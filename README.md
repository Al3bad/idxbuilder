# SANS Index Builder

A fast, keyboard-only terminal tool for building SANS course indexes. You type one line per entry, see similar entries live as you type, and export a print-ready index for the exam.

> This is a fork of [grahamhelton/SansTerminalIndexer](https://github.com/grahamhelton/SansTerminalIndexer). I've modified it with the help of AI to suit my needs.
>
> "The goal of this project is to create a very fast/efficient way to create an index. This program was greatly inspired by Matthew Toussain's [Voltaire](https://voltaire.publickey.io/) and Lesley Carhart's [Better GIAC testing With Pancakes](https://tisiphone.net/2015/08/18/giac-testing/)"

![image](./image.png)

## Features

* **One prompt line for everything.** No forms, no menus, and the cursor never leaves the prompt.
* **Live similar entries** while you type, so you can see whether a term is already indexed and where.
* **Recently added list** whenever the prompt is empty.
* **Safe for CLI commands.** Pipes, commas, semicolons and quotes are stored exactly as typed.
* **Separate Command Index.** Commands are grouped by tool, and the main index cross-references them.
* **Unlimited undo and redo** for adds, edits and deletes.
* **Instant search** across terms and notes, even with hundreds of thousands of entries (SQLite FTS5).
* **Print-ready export** to A4 two-column DOCX, plus CSV.
* **Automatic backups**: the last five copies are kept on every launch.

## Requirements

* Python 3.9 or newer
* A terminal with colour support (any modern terminal on Linux, macOS or Windows)
* Python packages: [`textual`](https://github.com/Textualize/textual) and [`python-docx`](https://github.com/python-openxml/python-docx). `python-docx` is only needed for DOCX export; without it, `:export` still writes the CSV.

## Installation and running

```bash
git clone https://github.com/Al3bad/idxbuilder
cd idxbuilder
```

### Option 1: uv (recommended)

The script declares its own dependencies ([PEP 723](https://peps.python.org/pep-0723/)), so [uv](https://docs.astral.sh/uv/) installs them into a cached environment the first time you run it. There's nothing else to set up.

```bash
uv run index.py              # pick or create an index
uv run index.py SEC504       # open or create the SEC504 index
```

### Option 2: plain Python

```bash
pip install textual python-docx
python3 index.py             # pick or create an index
python3 index.py SEC504      # open or create the SEC504 index
```

If pip refuses with an *"externally-managed-environment"* error (recent Ubuntu, Debian and Kali), use a virtual environment:

```bash
python3 -m venv .venv
source .venv/bin/activate         # Windows: .venv\Scripts\activate
pip install textual python-docx
python index.py SEC504
```

### Transparent background (optional)
 
By default the tool draws its own dark background. If your terminal is transparent and you want that to show through, start it with `--transparent`:
 
```bash
uv run sans_index.py SEC504 --transparent
python3 sans_index.py SEC504 --transparent
```
 
To make it permanent, set the environment variable in your shell profile (`~/.bashrc`, `~/.zshrc`):
 
```bash
export SANS_INDEX_TRANSPARENT=1
```

## Adding entries

Type an entry and press **Enter**. The book and page shown in the top bar are used unless you override them.

```
Prefetch                                      add at the current book/page
Prefetch ;; run count + last 8 run times      add with a note        (;; separates the note)
Prefetch @41                                  this entry only: page 41
Prefetch ;; evidence of execution @2.41       this entry only: book 2, page 41
$ tcpdump -nr x.pcap 'tcp port 80' | head     add a command (goes to the Command Index)
$ cut -d, -f1 a.csv | sort | uniq -c ;; count values @3.22
```

* `;;` separates the term from the optional note. Write `\;;` if a command really contains `;;`.
* `@page` or `@book.page` must be at the **end** of the line. It applies to that one entry, and the current page stays the same.
* Books can be numbers or letters, for example `@W.12` for a lab workbook.
* If the term already exists, the new page is added to it rather than creating a separate term. The exact same term on the same page is rejected as a duplicate.
* Press **→** or **End** to accept the grey inline suggestion of an existing term.

## Keys

```
Key          Action
---------------------------------------------------------
Enter        Add the entry (or save, when editing)
↑ / ↓        Select a row in the list
Tab          Copy the selected term into the prompt
Ctrl+E       Edit the selected entry in the prompt
Ctrl+X       Delete the selected entry (Ctrl+Z restores it)
Ctrl+Z       Undo
Ctrl+Y       Redo
PgDn / PgUp  Next / previous page
Esc          Clear the prompt / cancel editing
Ctrl+Q       Quit (Ctrl+C also works)
```

## Commands

Commands start with `:`. Type `:` on its own to show every command and the entry syntax in the app.

```
Command            Description
---------------------------------------------------------
:p 47              Go to page 47
:b 3               Go to book 3, page 1
:b 3 47            Go to book 3, page 47
:undo   :redo      Same as Ctrl+Z / Ctrl+Y
:export            CSV + Index DOCX + Command Index DOCX
:export csv        CSV only
:export docx       DOCX files only
:export combined   CSV + one DOCX with the Command Index at the end
:q                 Quit (everything is saved as you go)
?word              Search only; Enter or Ctrl+E edits the selected row
```

## Editing an entry

1. Find the entry: select it in the list with **↑ / ↓**, or search with `?word`.
2. Press **Ctrl+E**. The entry is loaded into the prompt in the same syntax you type it in, e.g. `Prefetch ;; note @2.41`.
3. Change anything (term, note, page or book) and press **Enter**. Press **Esc** to cancel.

Edits can be undone and redone like any other change.

## Exporting for print

`:export` writes into `.Index/exports/`:

| File | Contents |
|---|---|
| `SEC504_index.docx` | A–Z index of terms in two columns, with locations merged per term, e.g. `ShellBags 2.42, 4.9`. Tool names point to the Command Index (`tcpdump → see Command Index`). |
| `SEC504_commands.docx` | Commands grouped by tool (`tcpdump`, `vol.py`, `Get-WinEvent`…), printed in monospace. |
| `SEC504.csv` | Every entry with `Term, Note, Book, Page, Type, Created`. Opens in Excel and can be re-imported. |

Print the two DOCX files separately and keep them side by side in the exam. Use `:export combined` if you prefer a single document.

You can also export without opening the app:

```bash
uv run index.py SEC504 --export          # or: python3 index.py SEC504 --export
uv run index.py SEC504 --export csv
```

## Importing an existing CSV

CSV files from the original script (`Term,Book,Page,Note`) and this tool's own CSV export can be imported:

```bash
uv run index.py SEC504 --import-csv .Index/SEC504.csv
# or
python3 index.py SEC504 --import-csv .Index/SEC504.csv
```

Duplicates (the same term on the same book and page) are skipped. Terms starting with `$ ` are imported as commands.

## Index location

Indexes are SQLite databases in a hidden `.Index` directory inside the current working directory:

```
idxbuilder/
├── index.py
└── .Index/
    ├── SEC504.db          ← your index
    ├── SEC503.db
    ├── backups/           ← last 5 copies of each index, taken at every launch
    └── exports/           ← DOCX / CSV output
```

You can also pass a path directly: `python3 index.py ~/indexes/FOR508.db`.

## Troubleshooting

* **Search finds nothing.** Your Python's SQLite may be missing FTS5, which standard builds include. To check:
  ```bash
  python3 -c "import sqlite3; sqlite3.connect(':memory:').execute('create virtual table t using fts5(a)'); print('FTS5 OK')"
  ```
* **`DOCX export needs python-docx`.** Run `pip install python-docx`, or use `uv run`, which installs it for you.
* **Export fails with a permission error.** Close the DOCX in Word and run `:export` again.
