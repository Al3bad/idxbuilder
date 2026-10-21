#!/usr/bin/env python3
# /// script
# requires-python = ">=3.9"
# dependencies = ["textual>=1.0", "python-docx>=1.0"]
# ///
"""
SANS Index Builder: a one-line, keyboard-only indexer for SANS courses.

    pip install textual python-docx
    python sans_index.py              # pick or create an index in ./.Index/
    python sans_index.py SEC504       # open/create ./.Index/SEC504.db
    python sans_index.py SEC504 --import-csv old.csv    # Term,Book,Page,Note CSV (old script format OK)
    python sans_index.py SEC504 --transparent           # use the terminal's own (transparent) background
                                                        # or set SANS_INDEX_TRANSPARENT=1 to make it permanent

Everything happens on one prompt line:

    Prefetch                                  add at the current book/page
    Prefetch ;; run count + last 8 run times  add with a note      (;; separates the note)
    Prefetch @41        Prefetch @2.41        one-off page / book.page (must end the line)
    $ tcpdump -nr x.pcap 'tcp port 80' | head ;; read pcap @3.22
                                              "$" = command, goes to the separate Command Index.
                                              | , ; quotes etc. are all safe. Write \\;; for a literal ;;
    ?word                                     search only (Enter edits the selected row)
    :p 47   :b 3   :b 3 47   :export   :undo   :redo   :q      (type ":" to see them all)

Keys:  ↑↓ select row · Tab reuse selected term · Ctrl+E edit · Ctrl+X delete
       Ctrl+Z undo · Ctrl+Y redo · PgDn/PgUp next/previous page · Esc clear/cancel · Ctrl+Q quit
"""
from __future__ import annotations

import argparse
import csv
import os
import re
import sqlite3
import sys
import time
from datetime import datetime
from pathlib import Path

from rich.text import Text
from textual import work
from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal
from textual.css.query import NoMatches
from textual.suggester import Suggester
from textual.widgets import DataTable, Input, Static

APP_NAME = "SANS Index"
INDEX_DIR = Path.cwd() / ".Index"
PANEL_LIMIT = 200
RECENT_LIMIT = 100
UNDO_LIMIT = 1000
BACKUPS_KEPT = 5

ACC, DIM, LOC, OK, CMD, ERR, WARN, PUR = (
    "#58a6ff", "#8b949e", "#e3b341", "#56d364", "#7ee787", "#ff7b72", "#d29922", "#bc8cff")

PLACEHOLDER = "term   ·   term ;; note   ·   … @page   ·   $ command   ·   ?search   ·   : for commands"
HINT = (" Enter add · ↑↓ select · Tab reuse · ^E edit · ^X delete · ^Z undo · ^Y redo · "
        "PgDn/PgUp page · Esc clear · ^Q quit")

HELP_ROWS = [
    ("Prefetch", "Add at the current book / page"),
    ("Prefetch ;; run count + times", "Add with a note"),
    ("Prefetch @41   ·   Prefetch @2.41", "One-off page or book.page (end of line). Current page stays"),
    ("$ tcpdump -nr x.pcap | head", "Add a command → separate Command Index (| , ; quotes are safe)"),
    ("?word", "Search only. ↑↓ pick, Enter or ^E edit, ^X delete"),
    (":p 47", "Go to page 47       (PgDn / PgUp = next / previous page)"),
    (":b 3", "Go to book 3, page 1     ( :b 3 47  sets both )"),
    (":undo   :redo", "Same as Ctrl+Z / Ctrl+Y (unlimited, this session)"),
    (":export", "CSV + Index DOCX + Command Index DOCX  →  .Index/exports/"),
    (":export csv", "CSV only"),
    (":export combined", "CSV + one DOCX with the Command Index at the end"),
    (":q", "Quit (everything is already saved)"),
]


# ════════════════════════════════════════════════════════════════════════════
#  Parsing & formatting
# ════════════════════════════════════════════════════════════════════════════
LOC_TAIL = re.compile(
    r"(?:^|\s|;;)@\s*(?:(?P<book>[A-Za-z]*\d+|[A-Za-z]+)\s*\.\s*)?(?P<page>\d+)\s*$")
ESC_SEP = "\x00"


def parse_line(line: str):
    """'[$ ]term [;; note] [@[book.]page]' -> (term, note, kind, (book|None, page)|None)."""
    s = line.strip().replace(r"\;;", ESC_SEP)
    loc = None
    m = LOC_TAIL.search(s)
    if m:
        book = m.group("book")
        loc = (book.upper() if book else None, int(m.group("page")))
        s = s[:m.start()].rstrip()
        if s.endswith(";;"):
            s = s[:-2].rstrip()
    kind = "term"
    if s.startswith("$"):
        kind, s = "cmd", s[1:].lstrip()
    term, _, note = s.partition(";;")
    return (term.strip().replace(ESC_SEP, ";;"), note.strip().replace(ESC_SEP, ";;"), kind, loc)


def format_entry(r) -> str:
    """Inverse of parse_line: used to load an entry into the prompt for editing."""
    s = ("$ " if r["kind"] == "cmd" else "") + r["term"].replace(";;", r"\;;")
    if r["note"]:
        s += " ;; " + r["note"].replace(";;", r"\;;")
    return f"{s} @{r['book']}.{r['page']}"


def loc_str(book, page) -> str:
    return f"{book}.{page}"


def book_key(book: str):
    return (0, int(book), "") if str(book).isdigit() else (1, 0, str(book).lower())


def sort_key(term: str) -> str:
    return re.sub(r"^[^0-9A-Za-z]+", "", term).lower() or term.lower()


def letter_of(term: str) -> str:
    c = sort_key(term)[:1].upper()
    return c if c.isalpha() else "#"


def ago(ts: float) -> str:
    s = int(time.time() - ts)
    if s < 10:
        return "just now"
    if s < 60:
        return f"{s}s ago"
    if s < 3600:
        return f"{s // 60}m ago"
    if s < 86400:
        return f"{s // 3600}h ago"
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d")


def like_escape(s: str) -> str:
    return s.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def short(s: str, n=40) -> str:
    return s if len(s) <= n else s[: n - 1] + "…"


# ════════════════════════════════════════════════════════════════════════════
#  Storage (SQLite + FTS5) - fast at hundreds of thousands of rows
# ════════════════════════════════════════════════════════════════════════════
SCHEMA = """
PRAGMA journal_mode=WAL;
PRAGMA synchronous=NORMAL;
CREATE TABLE IF NOT EXISTS entries (
    id      INTEGER PRIMARY KEY,
    term    TEXT NOT NULL COLLATE NOCASE,
    note    TEXT NOT NULL DEFAULT '',
    book    TEXT NOT NULL COLLATE NOCASE,
    page    INTEGER NOT NULL,
    kind    TEXT NOT NULL DEFAULT 'term',          -- 'term' or 'cmd'
    created REAL NOT NULL
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_loc ON entries(term, book, page);
CREATE TABLE IF NOT EXISTS meta (k TEXT PRIMARY KEY, v TEXT);
"""
FTS_SCHEMA = """
CREATE VIRTUAL TABLE IF NOT EXISTS entries_fts USING fts5(
    term, note, content='entries', content_rowid='id', tokenize="unicode61 tokenchars '-_'");
CREATE TRIGGER IF NOT EXISTS e_ai AFTER INSERT ON entries BEGIN
  INSERT INTO entries_fts(rowid, term, note) VALUES (new.id, new.term, new.note); END;
CREATE TRIGGER IF NOT EXISTS e_ad AFTER DELETE ON entries BEGIN
  INSERT INTO entries_fts(entries_fts, rowid, term, note) VALUES ('delete', old.id, old.term, old.note); END;
CREATE TRIGGER IF NOT EXISTS e_au AFTER UPDATE ON entries BEGIN
  INSERT INTO entries_fts(entries_fts, rowid, term, note) VALUES ('delete', old.id, old.term, old.note);
  INSERT INTO entries_fts(rowid, term, note) VALUES (new.id, new.term, new.note); END;
"""


class Duplicate(Exception):
    pass


class IndexDB:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.con = sqlite3.connect(str(self.path))
        self.con.row_factory = sqlite3.Row
        self._migrate_v1()
        self.con.executescript(SCHEMA)
        try:
            self.con.executescript(FTS_SCHEMA)
            self.fts = True
        except sqlite3.OperationalError:
            self.fts = False
        self.con.commit()
        self.total = self.con.execute("SELECT COUNT(*) FROM entries").fetchone()[0]
        self.cmds = self.con.execute("SELECT COUNT(*) FROM entries WHERE kind='cmd'").fetchone()[0]

    def _migrate_v1(self):
        """Upgrade databases made by the first (tag-based) version of this tool."""
        cols = {r[1] for r in self.con.execute("PRAGMA table_info(entries)")}
        if "description" not in cols:
            return
        self.con.executescript("""
            DROP TRIGGER IF EXISTS e_ai; DROP TRIGGER IF EXISTS e_ad; DROP TRIGGER IF EXISTS e_au;
            DROP TABLE IF EXISTS entries_fts; DROP INDEX IF EXISTS ux_loc; DROP INDEX IF EXISTS ix_book;
            ALTER TABLE entries RENAME TO entries_v1;""")
        self.con.executescript(SCHEMA)
        self.con.execute("""INSERT INTO entries(id, term, note, book, page, kind, created)
            SELECT id, term, description, book, page,
                   CASE WHEN (',' || replace(tags, ' ', '') || ',') LIKE '%,cmd,%' THEN 'cmd' ELSE 'term' END,
                   created FROM entries_v1""")
        self.con.execute("DROP TABLE entries_v1")
        self.con.executescript(FTS_SCHEMA)
        self.con.execute("INSERT INTO entries_fts(entries_fts) VALUES ('rebuild')")
        self.con.commit()

    # -- meta / backup -----------------------------------------------------------
    def get_meta(self, k, default=""):
        r = self.con.execute("SELECT v FROM meta WHERE k=?", (k,)).fetchone()
        return r[0] if r else default

    def set_meta(self, k, v):
        self.con.execute("INSERT OR REPLACE INTO meta VALUES (?,?)", (k, str(v)))
        self.con.commit()

    def backup(self):
        if not self.total:
            return
        bdir = self.path.parent / "backups"
        bdir.mkdir(exist_ok=True)
        with sqlite3.connect(str(bdir / f"{self.path.stem}_{datetime.now():%Y%m%d_%H%M%S}.db")) as b:
            self.con.backup(b)
        for old in sorted(bdir.glob(f"{self.path.stem}_*.db"))[:-BACKUPS_KEPT]:
            old.unlink(missing_ok=True)

    # -- CRUD (rows are plain dicts so they can live on the undo/redo stacks) -----
    def get(self, eid):
        r = self.con.execute("SELECT * FROM entries WHERE id=?", (eid,)).fetchone()
        return dict(r) if r else None

    def insert(self, row: dict) -> dict:
        try:
            cur = self.con.execute(
                "INSERT INTO entries(id, term, note, book, page, kind, created) VALUES (?,?,?,?,?,?,?)",
                (row.get("id"), row["term"], row.get("note", ""), str(row["book"]), int(row["page"]),
                 row.get("kind", "term"), row.get("created") or time.time()))
        except sqlite3.IntegrityError:
            raise Duplicate(f"“{short(row['term'])}” is already indexed at {loc_str(row['book'], row['page'])}")
        self.con.commit()
        self.total += 1
        self.cmds += row.get("kind") == "cmd"
        return self.get(cur.lastrowid)

    def update(self, row: dict) -> dict:
        old = self.get(row["id"])
        try:
            self.con.execute("UPDATE entries SET term=?, note=?, book=?, page=?, kind=? WHERE id=?",
                             (row["term"], row["note"], str(row["book"]), int(row["page"]), row["kind"], row["id"]))
        except sqlite3.IntegrityError:
            raise Duplicate(f"“{short(row['term'])}” is already indexed at {loc_str(row['book'], row['page'])}")
        self.con.commit()
        self.cmds += (row["kind"] == "cmd") - (old["kind"] == "cmd")
        return self.get(row["id"])

    def delete(self, eid) -> dict | None:
        row = self.get(eid)
        if row:
            self.con.execute("DELETE FROM entries WHERE id=?", (eid,))
            self.con.commit()
            self.total -= 1
            self.cmds -= row["kind"] == "cmd"
        return row

    # -- lookups -----------------------------------------------------------------
    def recent(self, limit=RECENT_LIMIT):
        return [dict(r) for r in self.con.execute(
            "SELECT * FROM entries ORDER BY id DESC LIMIT ?", (limit,))]

    def find_term(self, term: str):
        """Exact (case-insensitive) term -> {'term', 'locs'} or None."""
        rows = self.con.execute("SELECT term, book, page FROM entries WHERE term=?", (term,)).fetchall()
        if not rows:
            return None
        locs = sorted({(r["book"], r["page"]) for r in rows}, key=lambda x: (book_key(x[0]), x[1]))
        return {"term": rows[0]["term"], "locs": ", ".join(loc_str(*l) for l in locs)}

    def first_with_prefix(self, prefix: str, kind: str):
        r = self.con.execute(
            "SELECT term FROM entries WHERE term >= ? AND term < ? AND kind = ? ORDER BY term LIMIT 1",
            (prefix, prefix + "￿", kind)).fetchone()
        return r[0] if r else None

    def similar(self, q: str, limit=PANEL_LIMIT):
        """Prefix-matched full-text search; exact and prefix hits first."""
        q = q.strip()
        if not q:
            return []
        order = ("ORDER BY CASE WHEN e.term = ? THEN 0 WHEN e.term LIKE ? ESCAPE '\\' THEN 1 ELSE 2 END, "
                 "e.term, length(e.book), e.book, e.page LIMIT ?")
        oparams = (q, like_escape(q) + "%", limit)
        tokens = re.findall(r"[\w\-]+", q)
        if self.fts and tokens:
            match = " AND ".join(f'"{t}"*' for t in tokens)
            try:
                rows = self.con.execute(
                    "SELECT e.* FROM entries e WHERE e.id IN "
                    f"(SELECT rowid FROM entries_fts WHERE entries_fts MATCH ?) {order}",
                    (match, *oparams)).fetchall()
                if rows or "".join(tokens) == re.sub(r"\s+", "", q):
                    return [dict(r) for r in rows]
            except sqlite3.OperationalError:
                pass
        # Fallback for symbol-heavy input that FTS can't tokenise (e.g. "| grep", "::1")
        pat = f"%{like_escape(q)}%"
        return [dict(r) for r in self.con.execute(
            f"SELECT e.* FROM entries e WHERE e.term LIKE ? ESCAPE '\\' OR e.note LIKE ? ESCAPE '\\' {order}",
            (pat, pat, *oparams))]


# ════════════════════════════════════════════════════════════════════════════
#  CSV import / export
# ════════════════════════════════════════════════════════════════════════════
CSV_FIELDS = ["Term", "Note", "Book", "Page", "Type", "Created"]


def export_csv(con, path: Path):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:     # BOM so Excel shows UTF-8 properly
        w = csv.writer(f)
        w.writerow(CSV_FIELDS)
        for r in con.execute("SELECT term, note, book, page, kind, created FROM entries "
                             "ORDER BY term, length(book), book, page"):
            w.writerow([r[0], r[1], r[2], r[3], "command" if r[4] == "cmd" else "term",
                        datetime.fromtimestamp(r[5]).isoformat(timespec="seconds")])


def import_csv(db: IndexDB, path: Path):
    """Accepts this tool's CSV and the old Term,Book,Page,Note format (also Title/Description)."""
    added = skipped = 0
    with open(path, newline="", encoding="utf-8-sig") as f:
        for raw in csv.DictReader(f):
            row = {(k or "").strip().lower(): (v or "").strip() for k, v in raw.items()}
            term = row.get("term") or row.get("title") or ""
            kind = "cmd" if row.get("type", "").lower().startswith("c") or term.startswith("$") else "term"
            term = term.lstrip("$ ").strip() if kind == "cmd" else term
            try:
                created = datetime.fromisoformat(row["created"]).timestamp() if row.get("created") else None
                db.insert({"term": term, "note": row.get("note") or row.get("description") or "",
                           "book": (row.get("book") or "1").upper(), "page": int(row["page"]),
                           "kind": kind, "created": created})
                added += 1
            except (Duplicate, ValueError, KeyError):
                skipped += 1
    return added, skipped


# ════════════════════════════════════════════════════════════════════════════
#  Printable DOCX: A-Z Index + Command Index (grouped by tool)
# ════════════════════════════════════════════════════════════════════════════
SKIP_PREFIX = {"sudo", "doas", "time", "nohup", "env", "exec"}
INTERPRETERS = re.compile(r"^(python[\d.]*|py|perl|ruby|bash|sh|zsh|powershell(\.exe)?|pwsh|node)$", re.I)


def tool_of(cmd: str) -> str:
    """First real program in a command line: 'sudo python3 vol.py -f x' -> 'vol.py'."""
    toks = cmd.split()
    i = 0
    while i < len(toks) and (toks[i] in SKIP_PREFIX or re.match(r"^[A-Za-z_]\w*=", toks[i])):
        i += 1
    if i >= len(toks):
        return toks[0] if toks else "?"
    tool = toks[i]
    if INTERPRETERS.match(tool) and i + 1 < len(toks) and not toks[i + 1].startswith("-"):
        tool = toks[i + 1]
    return re.split(r"[\\/]", tool)[-1] or tool


def merged(con, kind):
    """All entries of one kind, merged by term (case-insensitive), sorted A-Z."""
    groups = {}
    for term, note, book, page in con.execute(
            "SELECT term, note, book, page FROM entries WHERE kind=?", (kind,)):
        g = groups.setdefault(term.lower(), {"term": term, "note": "", "locs": set(), "xref": ""})
        if len(note) > len(g["note"]):
            g["note"] = note
        g["locs"].add((book, page))
    out = list(groups.values())
    for g in out:
        g["locs"] = sorted(g["locs"], key=lambda x: (book_key(x[0]), x[1]))
    out.sort(key=lambda g: (sort_key(g["term"]), g["term"].lower()))
    return out


def command_groups(con):
    groups = {}
    for g in merged(con, "cmd"):
        groups.setdefault(tool_of(g["term"]), []).append(g)
    return sorted(groups.items(), key=lambda kv: (sort_key(kv[0]), kv[0].lower()))


def index_entries(con, groups):
    """A-Z terms plus cross-references from tool names to the Command Index."""
    terms = merged(con, "term")
    by_name = {g["term"].lower(): g for g in terms}
    for tool, _ in groups:
        if tool.lower() in by_name:
            by_name[tool.lower()]["xref"] = "see also Command Index"
        else:
            terms.append({"term": tool, "note": "", "locs": [], "xref": "see Command Index"})
    terms.sort(key=lambda g: (sort_key(g["term"]), g["term"].lower()))
    return terms


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_BAD_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _x(text: str) -> str:
    return _BAD_XML.sub("", text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _run(text, bold=False, italic=False, color=None, size=None, mono=False) -> str:
    """One <w:r> as raw XML (python-docx's per-property setters are ~100x slower)."""
    pr = ""
    if mono:
        pr += '<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:eastAsia="Consolas" w:cs="Consolas"/>'
    if bold:
        pr += "<w:b/>"
    if italic:
        pr += "<w:i/>"
    if color:
        pr += f'<w:color w:val="{color}"/>'
    if size:
        pr += f'<w:sz w:val="{size}"/>'
    return f'<w:r><w:rPr>{pr}</w:rPr><w:t xml:space="preserve">{_x(text)}</w:t></w:r>'


class DocxWriter:
    """Two-column A4 index. Paragraphs are buffered as XML and inserted in one go,
    so 100k+ entries export in seconds."""

    def __init__(self, header: str):
        try:
            from docx import Document
            from docx.enum.section import WD_SECTION
            from docx.oxml import OxmlElement, parse_xml
            from docx.oxml.ns import qn
            from docx.shared import Cm, Pt, RGBColor
        except ImportError:
            raise RuntimeError("DOCX export needs python-docx  (pip install python-docx)")
        self.Pt, self.RGB, self.Cm, self.OX, self.qn, self.WD = Pt, RGBColor, Cm, OxmlElement, qn, WD_SECTION
        self.parse_xml = parse_xml
        self.buf: list[str] = []
        self.doc = Document()
        st = self.doc.styles["Normal"]
        st.font.name, st.font.size = "Calibri", Pt(8.5)
        st.paragraph_format.space_after = st.paragraph_format.space_before = Pt(0)
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)                   # A4
        for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
            setattr(sec, side, Cm(1.3))
        self._setup(sec, header)

    def _setup(self, sec, header):
        cols = sec._sectPr.find(self.qn("w:cols"))
        if cols is None:
            cols = self.OX("w:cols")
            sec._sectPr.append(cols)
        cols.set(self.qn("w:num"), "2")
        cols.set(self.qn("w:space"), "425")
        sec.header.is_linked_to_previous = False
        p = sec.header.paragraphs[0]
        p.text = header
        p.runs[0].font.size = self.Pt(8)
        p.runs[0].font.color.rgb = self.RGB(0x66, 0x66, 0x66)

    def _flush(self):
        if not self.buf:
            return
        frag = self.parse_xml(f'<w:body xmlns:w="{W_NS}">{"".join(self.buf)}</w:body>')
        anchor = self.doc.element.body.sectPr
        for el in list(frag):
            anchor.addprevious(el)
        self.buf.clear()

    def new_section(self, header):
        self._flush()
        self._setup(self.doc.add_section(self.WD.NEW_PAGE), header)

    def title(self, text, sub):
        self.buf.append(f"<w:p>{_run(text, bold=True, size=32)}</w:p>"
                        f'<w:p><w:pPr><w:spacing w:after="80"/></w:pPr>{_run(sub, color="666666", size=16)}</w:p>')

    def heading(self, text, mono=False):
        self.buf.append(
            '<w:p><w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="1F4E79"/>'
            '</w:pBdr><w:spacing w:before="120" w:after="40"/></w:pPr>'
            f'{_run(text, bold=True, color="1F4E79", size=20 if mono else 24, mono=mono)}</w:p>')

    def entry(self, term, locs, note="", mono=False, xref=""):
        runs = _run(term, bold=not mono, mono=mono)
        if locs:
            runs += _run("  " + ", ".join(loc_str(b, pg) for b, pg in locs), bold=True, color="9C5500")
        if note:
            runs += _run(f"  {note}", italic=True, color="555555", size=15)
        if xref:
            runs += _run(f"  → {xref}", italic=True, color="1F4E79", size=15)
        self.buf.append(f'<w:p><w:pPr><w:ind w:left="198" w:hanging="198"/></w:pPr>{runs}</w:p>')

    def write_index(self, entries):
        cur = None
        for g in entries:
            if letter_of(g["term"]) != cur:
                cur = letter_of(g["term"])
                self.heading(cur)
            self.entry(g["term"], g["locs"], g["note"], xref=g["xref"])

    def write_commands(self, groups):
        for tool, ents in groups:
            self.heading(tool, mono=True)
            for g in ents:
                self.entry(g["term"], g["locs"], g["note"], mono=True)

    def save(self, path):
        self._flush()
        self.doc.save(str(path))


def export_all(db_path: Path, out_dir: Path, name: str, mode: str = "all"):
    """mode: all | csv | docx | combined.  Returns (written_files, warnings)."""
    out_dir.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(str(db_path))
    written, warnings = [], []
    try:
        if mode in ("all", "csv", "combined"):
            export_csv(con, out_dir / f"{name}.csv")
            written.append(f"{name}.csv")
        if mode in ("all", "docx", "combined"):
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
            groups = command_groups(con)
            entries = index_entries(con, groups)
            n_cmd = sum(len(e) for _, e in groups)
            try:
                w = DocxWriter(f"{name} · Index")
                w.title(f"{name} Index", f"{len(entries):,} terms · generated {stamp}")
                w.write_index(entries)
                if mode == "combined" and groups:
                    w.new_section(f"{name} · Command Index")
                    w.title("Command Index", f"{n_cmd:,} commands in {len(groups)} tools")
                    w.write_commands(groups)
                w.save(out_dir / f"{name}_index.docx")
                written.append(f"{name}_index.docx")
                if mode != "combined":
                    w = DocxWriter(f"{name} · Command Index")
                    w.title(f"{name} Command Index", f"{n_cmd:,} commands in {len(groups)} tools · {stamp}")
                    w.write_commands(groups)
                    w.save(out_dir / f"{name}_commands.docx")
                    written.append(f"{name}_commands.docx")
            except RuntimeError as ex:
                warnings.append(str(ex))
    finally:
        con.close()
    return written, warnings


# ════════════════════════════════════════════════════════════════════════════
#  TUI
# ════════════════════════════════════════════════════════════════════════════
class TermSuggester(Suggester):
    """Inline ghost-text completion from existing terms (accept with → or End)."""

    def __init__(self, db: IndexDB):
        super().__init__(use_cache=False, case_sensitive=True)
        self.db = db

    async def get_suggestion(self, value: str):
        if not value or ";;" in value or value[0] in ":?" or "@" in value[-8:]:
            return None
        kind, rest = ("cmd", value[1:].lstrip()) if value.startswith("$") else ("term", value)
        if not rest:
            return None
        hit = self.db.first_with_prefix(rest, kind)
        return value + hit[len(rest):] if hit and hit.lower() != rest.lower() else None


CSS = f"""
Screen {{ background: #0d1117; }}
#top {{ height: 1; background: #161b22; }}
#title {{ height: 1; padding: 0 1; margin-top: 1; }}
#panel {{ height: 1fr; background: #0d1117; border: round #30363d; scrollbar-size-horizontal: 0; }}
#panel > .datatable--header {{ background: #161b22; color: {ACC}; text-style: bold; }}
#panel > .datatable--even-row {{ background: #10151c; }}
#panel > .datatable--cursor, #panel:blur > .datatable--cursor {{ background: #1f3a5f; text-style: bold; }}
#msg {{ height: 1; padding: 0 1; }}
#prompt-row {{ height: 3; }}
#label {{ width: auto; padding: 1 0 0 1; color: {PUR}; text-style: bold; }}
#label.editing {{ color: {WARN}; }}
#prompt {{ width: 1fr; border: tall {ACC}; background: #161b22; }}
#prompt.editing {{ border: tall {WARN}; }}
#hint {{ height: 1; color: #6e7681; background: #161b22; }}

/* --transparent: let the terminal's own background (and its transparency) show through */
Screen.transparent, Screen.transparent #top, Screen.transparent #panel, Screen.transparent #hint,
Screen.transparent #prompt, Screen.transparent #prompt.editing {{ background: ansi_default; }}
Screen.transparent #panel > .datatable--header,
Screen.transparent #panel > .datatable--even-row,
Screen.transparent #panel > .datatable--odd-row {{ background: ansi_default; }}
"""


class IndexApp(App):
    TITLE = APP_NAME
    CSS = CSS
    ENABLE_COMMAND_PALETTE = False
    BINDINGS = [
        Binding("ctrl+q,ctrl+c", "quit", "Quit", priority=True),
        Binding("ctrl+z", "undo", "Undo", priority=True),
        Binding("ctrl+y", "redo", "Redo", priority=True),
        Binding("ctrl+e", "edit_selected", "Edit", priority=True),
        Binding("ctrl+x", "delete_selected", "Delete", priority=True),
        Binding("up", "move(-1)", show=False, priority=True),
        Binding("down", "move(1)", show=False, priority=True),
        Binding("tab", "reuse", show=False, priority=True),
        Binding("pagedown", "page(1)", show=False, priority=True),
        Binding("pageup", "page(-1)", show=False, priority=True),
        Binding("escape", "escape", show=False, priority=True),
    ]

    def __init__(self, db: IndexDB, transparent: bool = False):
        # ansi_color lets Textual emit the terminal's *default* background instead of a solid colour
        super().__init__(ansi_color=True if transparent else None)
        self.transparent = transparent
        self.db = db
        self.index_name = db.path.stem
        self.book = db.get_meta("book", "1")
        self.page = int(db.get_meta("page", "1") or 1)
        self.undo_stack: list[tuple] = []
        self.redo_stack: list[tuple] = []
        self.rows: list[dict] = []
        self.panel_kind = None          # 'rows' | 'help'
        self.editing: dict | None = None
        self.last_id = None
        self._timer = None

    # -- layout ------------------------------------------------------------------
    def compose(self) -> ComposeResult:
        yield Static(id="top")
        yield Static(id="title")
        yield DataTable(id="panel", cursor_type="row", zebra_stripes=True)
        yield Static(id="msg")
        with Horizontal(id="prompt-row"):
            yield Static("STI~>", id="label")
            yield Input(id="prompt", placeholder=PLACEHOLDER, suggester=TermSuggester(self.db))
        yield Static(HINT, id="hint")

    def on_mount(self):
        self.theme = "textual-dark"
        if self.transparent:
            self.screen.add_class("transparent")
        self.table = self.query_one("#panel", DataTable)
        self.table.can_focus = False                       # focus never leaves the prompt
        self.prompt = self.query_one("#prompt", Input)
        self.prompt.focus()
        self.refresh_top()
        self.refresh_panel()
        self.say("Type a term and press Enter.  Type “:” for all commands and syntax.", DIM)
        self.set_interval(30, lambda: self._safe_refresh(keep_cursor=True))   # keeps “Added” fresh

    def on_resize(self):
        if not hasattr(self, "prompt"):          # resize fires once before on_mount
            return
        self.panel_kind = None
        self.refresh_panel(keep_cursor=True)

    # -- rendering ----------------------------------------------------------------
    def refresh_top(self):
        t = Text()
        t.append(f" ◆ {APP_NAME} ", style=f"bold #0d1117 on {ACC}")
        t.append(f"  {self.index_name}  ", style="bold")
        t.append("│ Book ", style=DIM); t.append(str(self.book), style=f"bold {LOC}")
        t.append("  Page ", style=DIM); t.append(str(self.page), style=f"bold {LOC}")
        t.append(f"  │ {self.db.total:,} entries · {self.db.cmds:,} commands", style=DIM)
        t.append(f"  │ undo {len(self.undo_stack)} · redo {len(self.redo_stack)}", style=DIM)
        self.query_one("#top", Static).update(t)

    def say(self, msg: str, style: str = OK):
        self.query_one("#msg", Static).update(Text(msg, style=style))

    def set_title(self, *parts):
        self.query_one("#title", Static).update(Text.assemble(*parts))

    def refresh_panel(self, keep_cursor=False):
        v = self.prompt.value
        cursor = self.table.cursor_row if keep_cursor else 0
        if v == ":" or re.match(r":[A-Za-z]", v):
            return self.show_help(v[1:].split(" ")[0].lower(), cursor)
        if v.startswith("?"):
            q = v[1:].strip()
            rows = self.db.similar(q) if q else self.db.recent()
            self.set_title((" Search ", f"bold #0d1117 on {ACC}"),
                           (f"  “{q}”  {len(rows)}{'+' if len(rows) >= PANEL_LIMIT else ''} result(s)  ", "bold") if q
                           else ("  type to search terms and notes  ", DIM),
                           ("↑↓ pick · Enter/^E edit · ^X delete", DIM))
        else:
            term = parse_line(v)[0] if v.strip() else ""
            if term and self.editing:
                rows = self.db.similar(term)
                self.set_title((f" Editing #{self.editing['id']} ", f"bold #0d1117 on {WARN}"),
                               ("  Enter saves · Esc cancels · Ctrl+Z afterwards reverts it", DIM))
            elif term:
                rows = self.db.similar(term)
                hit = self.db.find_term(term)
                self.set_title((" Similar ", f"bold #0d1117 on {PUR}"), (f"  “{short(term)}”  ", "bold"),
                               (f"exists at {hit['locs']} → Enter adds this location" if hit
                                else "new term" if not rows else f"{len(rows)} similar", OK if hit else DIM))
            else:
                rows = self.db.recent()
                self.set_title((" Recently added ", f"bold #0d1117 on {OK}"), ("  newest first", DIM))
        self.show_rows(rows, cursor)

    def _columns(self, kind):
        if self.panel_kind == kind:
            self.table.clear()
            return
        self.table.clear(columns=True)
        w = max(60, self.size.width - 6)
        if kind == "rows":
            tw = int(w * 0.45)
            self.table.add_column("Term", width=tw)
            self.table.add_column("Loc", width=7)
            self.table.add_column("Note", width=max(10, w - tw - 7 - 9 - 8))
            self.table.add_column("Added", width=9)
        else:
            self.table.add_column("Type this", width=36)
            self.table.add_column("What it does", width=max(20, w - 36 - 4))
        self.panel_kind = kind

    def show_rows(self, rows, cursor=0):
        self._columns("rows")
        self.rows = rows
        for r in rows:
            cmd = r["kind"] == "cmd"
            style = f"bold {OK}" if r["id"] == self.last_id else (CMD if cmd else "")
            self.table.add_row(
                Text.assemble(("$ ", DIM) if cmd else "", (r["term"], style)),
                Text(loc_str(r["book"], r["page"]), style=LOC),
                Text(r["note"], style=DIM),
                Text(ago(r["created"]), style=DIM),
                key=str(r["id"]))
        if rows:
            self.table.move_cursor(row=min(cursor, len(rows) - 1), animate=False)

    def show_help(self, prefix, cursor=0):
        self._columns("help")
        self.rows = []
        self.set_title((" Commands & syntax ", f"bold #0d1117 on {WARN}"),
                       ("  Enter runs a : command · Tab copies it", DIM))
        for k, d in HELP_ROWS:
            if not prefix or (k.startswith(":") and k[1:].lower().startswith(prefix)):
                self.table.add_row(Text(k, style=f"bold {CMD if k[0] in ':?$' else 'white'}"), Text(d, style=DIM))
        if self.table.row_count:
            self.table.move_cursor(row=min(cursor, self.table.row_count - 1), animate=False)

    # -- input events --------------------------------------------------------------
    def on_input_changed(self, _):
        if self._timer:
            self._timer.stop()
        self._timer = self.set_timer(0.06, self._safe_refresh)

    def _safe_refresh(self, keep_cursor=False):
        try:                                  # a timer can fire while the app is shutting down
            self.refresh_panel(keep_cursor)
        except NoMatches:
            pass

    def on_input_submitted(self, event: Input.Submitted):
        v = event.value.strip()
        if not v:
            return
        if v == ":" or re.match(r":[A-Za-z]", v):
            self.run_command(v[1:])
        elif v.startswith("?"):
            self.action_edit_selected()
        elif self.editing:
            self.save_edit(v)
        else:
            self.add(v)

    # -- entry actions -------------------------------------------------------------
    def push(self, action):
        self.undo_stack.append(action)
        del self.undo_stack[:-UNDO_LIMIT]
        self.redo_stack.clear()

    def clear_prompt(self):
        self.prompt.value = ""
        self.refresh_top()
        self.refresh_panel()

    def add(self, line):
        term, note, kind, loc = parse_line(line)
        if not term:
            return self.say("Nothing to add: the term is empty", ERR)
        book, page = (loc[0] or self.book, loc[1]) if loc else (self.book, self.page)
        if page < 1:
            return self.say("Page must be 1 or more", ERR)
        existing = self.db.find_term(term)
        if existing:
            term = existing["term"]                    # keep the first spelling/casing
        try:
            row = self.db.insert({"term": term, "note": note, "book": book, "page": page, "kind": kind})
        except Duplicate as ex:
            return self.say(str(ex) + " (nothing added)", WARN)
        self.push(("add", row))
        self.last_id = row["id"]
        extra = f"   · also at {existing['locs']}" if existing else ""
        self.say(f"✔ Added {'command ' if kind == 'cmd' else ''}“{short(term)}” @ {loc_str(book, page)}{extra}")
        self.clear_prompt()

    def selected(self):
        if self.panel_kind != "rows" or not self.rows:
            return None
        return self.rows[min(self.table.cursor_row, len(self.rows) - 1)]

    def action_edit_selected(self):
        r = self.selected()
        if not r:
            return self.say("Select a row first (↑↓)", WARN)
        self.editing = r
        self.prompt.value = format_entry(r)
        self.prompt.cursor_position = len(self.prompt.value)
        self.query_one("#label", Static).update(f"EDIT #{r['id']}>")
        self.query_one("#label").add_class("editing")
        self.prompt.add_class("editing")
        self.say("Editing: change the line, Enter saves, Esc cancels", WARN)

    def stop_editing(self):
        self.editing = None
        self.query_one("#label", Static).update("STI~>")
        self.query_one("#label").remove_class("editing")
        self.prompt.remove_class("editing")

    def save_edit(self, line):
        before = self.editing
        term, note, kind, loc = parse_line(line)
        if not term:
            return self.say("The term can't be empty", ERR)
        book, page = (loc[0] or before["book"], loc[1]) if loc else (before["book"], before["page"])
        after = dict(before, term=term, note=note, kind=kind, book=book, page=page)
        if all(after[k] == before[k] for k in ("term", "note", "kind", "book", "page")):
            self.stop_editing()
            self.say("No changes", DIM)
            return self.clear_prompt()
        try:
            after = self.db.update(after)
        except Duplicate as ex:
            return self.say(str(ex), WARN)
        self.push(("edit", before, after))
        self.stop_editing()
        self.say(f"✔ Updated “{short(after['term'])}” @ {loc_str(after['book'], after['page'])}")
        self.clear_prompt()

    def action_delete_selected(self):
        r = self.selected()
        if not r:
            return self.say("Select a row first (↑↓)", WARN)
        if self.editing and self.editing["id"] == r["id"]:
            self.stop_editing()
            self.prompt.value = ""
        self.db.delete(r["id"])
        self.push(("del", r))
        self.say(f"✖ Deleted “{short(r['term'])}” @ {loc_str(r['book'], r['page'])}   (Ctrl+Z restores it)", WARN)
        self.refresh_top()
        self.refresh_panel(keep_cursor=True)

    # -- undo / redo -------------------------------------------------------------
    def _apply(self, action, undo: bool):
        kind = action[0]
        if kind == "add":
            self.db.delete(action[1]["id"]) if undo else self.db.insert(action[1])
        elif kind == "del":
            self.db.insert(action[1]) if undo else self.db.delete(action[1]["id"])
        else:
            self.db.update(action[1] if undo else action[2])

    @staticmethod
    def _describe(action, undo):
        r = action[2] if action[0] == "edit" else action[1]
        verb = {"add": ("removed", "re-added"), "del": ("restored", "deleted again"),
                "edit": ("reverted edit of", "re-applied edit of")}[action[0]][0 if undo else 1]
        return f"{verb} “{short(r['term'])}” @ {loc_str(r['book'], r['page'])}"

    def _step(self, src, dst, undo):
        if not src:
            return self.say(f"Nothing to {'undo' if undo else 'redo'}", DIM)
        if self.editing:
            self.stop_editing()
            self.prompt.value = ""
        action = src.pop()
        try:
            self._apply(action, undo)
        except Duplicate as ex:
            src.append(action)
            return self.say(f"Can't {'undo' if undo else 'redo'}: {ex}", ERR)
        dst.append(action)
        self.last_id = None
        self.say(("↶ Undo: " if undo else "↷ Redo: ") + self._describe(action, undo)
                 + f"    ({len(self.undo_stack)} undo · {len(self.redo_stack)} redo left)", ACC)
        self.refresh_top()
        self.refresh_panel(keep_cursor=True)

    def action_undo(self):
        self._step(self.undo_stack, self.redo_stack, True)

    def action_redo(self):
        self._step(self.redo_stack, self.undo_stack, False)

    # -- navigation ---------------------------------------------------------------
    def action_move(self, delta: int):
        if self.table.row_count:
            self.table.move_cursor(row=max(0, min(self.table.row_count - 1, self.table.cursor_row + delta)),
                                   animate=False)

    def action_reuse(self):
        """Tab: copy the selected term (or : command) into the prompt."""
        if self.panel_kind == "help" and self.table.row_count:
            key = str(self.table.get_row_at(self.table.cursor_row)[0])
            if key.startswith(":"):
                self.prompt.value = key.split()[0] + " "
        elif (r := self.selected()):
            self.prompt.value = ("$ " if r["kind"] == "cmd" else "") + r["term"]
        self.prompt.cursor_position = len(self.prompt.value)

    def action_escape(self):
        if self.editing:
            self.stop_editing()
            self.say("Edit cancelled", DIM)
        self.prompt.value = ""

    def set_page(self, page: int, book: str | None = None):
        if book is not None:
            self.book = book
            self.db.set_meta("book", book)
        self.page = max(1, page)
        self.db.set_meta("page", self.page)
        self.refresh_top()
        self.say(f"Book {self.book} · page {self.page}", LOC)

    def action_page(self, delta: int):
        self.set_page(self.page + delta)

    # -- : commands ---------------------------------------------------------------
    def run_command(self, text: str):
        parts = text.split()
        if not parts:
            return
        name, args = parts[0].lower(), parts[1:]
        m = re.fullmatch(r"([pb])(\w+)", name)          # also accept :p47 and :b3
        if m and name not in ("pp",):
            name, args = m.group(1), [m.group(2)] + args
        self.prompt.value = ""
        if name in ("p", "page"):
            if args and args[0].isdigit():
                return self.set_page(int(args[0]))
            return self.say("Usage: :p 47", ERR)
        if name in ("b", "book"):
            if not args:
                return self.say("Usage: :b 3   or   :b 3 47", ERR)
            page = int(args[1]) if len(args) > 1 and args[1].isdigit() else 1
            return self.set_page(page, args[0].upper())
        if name in ("u", "undo"):
            return self.action_undo()
        if name in ("r", "redo"):
            return self.action_redo()
        if name in ("export", "x", "e"):
            mode = (args[0].lower() if args else "all")
            if mode not in ("all", "csv", "docx", "combined"):
                return self.say("Usage: :export   :export csv   :export docx   :export combined", ERR)
            self.say("Exporting…", DIM)
            return self.do_export(mode)
        if name in ("q", "quit", "exit"):
            return self.exit()
        if name in ("h", "help"):
            self.prompt.value = ":"
            return
        self.say(f"Unknown command :{name}   (type “:” to list commands)", ERR)

    @work(thread=True, exclusive=True)
    def do_export(self, mode: str):
        out = self.db.path.parent / "exports"
        try:
            written, warnings = export_all(self.db.path, out, self.index_name, mode)
        except PermissionError as ex:
            self.call_from_thread(self.say, f"Export failed: is the file open in Word? ({ex})", ERR)
            return
        except Exception as ex:                                # report anything else in the message line
            self.call_from_thread(self.say, f"Export failed: {ex}", ERR)
            return
        msg = f"✔ Exported to {out}:  " + "  ·  ".join(written)
        if warnings:
            msg += "   ⚠ " + "; ".join(warnings)
        self.call_from_thread(self.say, msg, WARN if warnings else OK)


# ════════════════════════════════════════════════════════════════════════════
#  Startup
# ════════════════════════════════════════════════════════════════════════════
def resolve_path(arg: str) -> Path:
    p = Path(arg).expanduser()
    if p.suffix.lower() == ".db" or len(p.parts) > 1:
        return p if p.suffix.lower() == ".db" else p.with_suffix(".db")
    return INDEX_DIR / f"{arg}.db"


def choose_index() -> Path:
    INDEX_DIR.mkdir(parents=True, exist_ok=True)
    dbs = sorted(INDEX_DIR.glob("*.db"), key=lambda p: p.name.lower())
    print(f"\033[1;34m◆ {APP_NAME}\033[0m   ({INDEX_DIR})\n")
    for i, p in enumerate(dbs, 1):
        print(f"  \033[32m{i:>2}\033[0m  {p.stem}")
    prompt = "\nNumber to open, or a new name to create: " if dbs else "New index name (e.g. SEC504): "
    while True:
        try:
            choice = input(prompt).strip()
        except (EOFError, KeyboardInterrupt):
            raise SystemExit(0)
        if choice.isdigit() and 1 <= int(choice) <= len(dbs):
            return dbs[int(choice) - 1]
        if choice and not choice.isdigit():
            return INDEX_DIR / f"{re.sub(r'[^A-Za-z0-9_.-]+', '_', choice)}.db"


def main():
    ap = argparse.ArgumentParser(description=APP_NAME, formatter_class=argparse.RawDescriptionHelpFormatter,
                                 epilog=__doc__)
    ap.add_argument("index", nargs="?", help="index name (e.g. SEC504) or path to a .db file")
    ap.add_argument("--import-csv", metavar="FILE", help="import a CSV (Term,Book,Page,Note) and exit")
    ap.add_argument("--transparent", action="store_true",
                    help="use the terminal's background (for transparent terminals); "
                         "or set SANS_INDEX_TRANSPARENT=1")
    ap.add_argument("--export", metavar="MODE", nargs="?", const="all",
                    help="export without opening the TUI: all | csv | docx | combined")
    args = ap.parse_args()

    path = resolve_path(args.index) if args.index else choose_index()
    path.parent.mkdir(parents=True, exist_ok=True)
    db = IndexDB(path)

    if args.import_csv:
        added, skipped = import_csv(db, Path(args.import_csv))
        print(f"Imported {added:,} entries into {path.name} ({skipped:,} duplicates/invalid skipped)")
        return
    if args.export:
        written, warnings = export_all(path, path.parent / "exports", path.stem, args.export)
        print("Wrote:", ", ".join(written), *(f"\nWarning: {w}" for w in warnings))
        return

    db.backup()
    transparent = args.transparent or os.environ.get("SANS_INDEX_TRANSPARENT", "").lower() in ("1", "true", "yes", "on")
    IndexApp(db, transparent=transparent).run()


if __name__ == "__main__":
    main()
