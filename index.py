#!/usr/bin/env python3

from __future__ import annotations

import os
import shutil
import subprocess
import csv
import signal
from collections import deque
from pathlib import Path

try:
    import readline
except ImportError:
    readline = None

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

INDEX_DIRECTORY = Path.cwd() / ".Index"
CSV_HEADERS = ["Term", "Book", "Page", "Note"]
NOTE_DELIMITER = "|"
HISTORY_LIMIT = 10
PREVIEW_ROWS = 5

RESET = "\033[0m"
BLACK = "\033[30m"
RED = "\033[31m"
GREEN = "\033[32m"
YELLOW = "\033[33m"
BLUE = "\033[34m"
MAGENTA = "\033[35m"
CYAN = "\033[36m"
WHITE = "\033[37m"
BOLD = "\033[1m"


SEPARATOR = f"{BLUE}{'-' * 60}{RESET}"
SUCCESS = f"{BLUE}[{GREEN}+{BLUE}]{RESET}"
ERROR = f"{YELLOW}[{RED}X{YELLOW}]{RESET}"
INFO = f"{BLUE}[{YELLOW}!{BLUE}]{RESET}"


# ---------------------------------------------------------------------------
# Readline prompt
# ---------------------------------------------------------------------------


def readline_colour(colour: str) -> str:
    """
    Mark an ANSI colour sequence as non-printing.

    This prevents readline from incorrectly calculating the prompt width
    when navigating command history with the arrow keys.
    """
    return f"\001{colour}\002"


ENTRY_PROMPT = f"{readline_colour(MAGENTA)}STI~> {readline_colour(YELLOW)}"


# ---------------------------------------------------------------------------
# Terminal helpers
# ---------------------------------------------------------------------------


def clear_screen() -> None:
    """Clear the terminal."""
    print("\033[2J\033[H", end="")


def pause(message: str = "Press Enter to continue.") -> None:
    """Pause until the user presses Enter."""
    input(message)


def handle_interrupt(signum: int, frame: object) -> None:
    """Exit cleanly when Ctrl+C is pressed."""
    del signum, frame

    print(f"\n{RED}Exiting.{RESET}")
    raise SystemExit(0)


# ---------------------------------------------------------------------------
# Index directory and file selection
# ---------------------------------------------------------------------------


def ensure_index_directory() -> Path:
    """Create and return the local .Index directory."""
    if not INDEX_DIRECTORY.exists():
        print(f"{INFO} Creating index directory at {CYAN}{INDEX_DIRECTORY}{RESET}")

    INDEX_DIRECTORY.mkdir(parents=True, exist_ok=True)
    return INDEX_DIRECTORY


def find_csv_files(directory: Path) -> list[Path]:
    """Return the CSV files in a directory, sorted alphabetically."""
    return sorted(
        (
            path
            for path in directory.iterdir()
            if path.is_file() and path.suffix.lower() == ".csv"
        ),
        key=lambda path: path.name.lower(),
    )


def create_index_file(directory: Path) -> Path:
    """Create a new CSV index."""
    while True:
        filename = input(f"{SUCCESS} {GREEN}New index name:{RESET}\n> ").strip()

        if not filename:
            print(f"{ERROR} The filename cannot be empty.")
            continue

        csv_path = directory / filename

        if csv_path.suffix.lower() != ".csv":
            csv_path = csv_path.with_suffix(".csv")

        if csv_path.exists():
            open_existing = (
                input(f"{INFO} {csv_path.name} already exists. Open it? [y/N]: ")
                .strip()
                .lower()
            )

            if open_existing == "y":
                return csv_path

            continue

        try:
            with csv_path.open(
                "w",
                newline="",
                encoding="utf-8",
            ) as csv_file:
                writer = csv.writer(csv_file)
                writer.writerow(CSV_HEADERS)
        except OSError as error:
            print(f"{ERROR} Could not create the index: {error}")
            continue

        print(f"{SUCCESS} Created {CYAN}{csv_path.name}{RESET}")
        return csv_path


def select_index_file(directory: Path) -> Path:
    """Select an existing index or create a new one."""
    while True:
        clear_screen()

        csv_files = find_csv_files(directory)

        print(f"{SUCCESS} Welcome to SANS Terminal Indexer")
        print(SEPARATOR)

        if csv_files:
            print(f"{INFO} Select an index:\n")

            for number, csv_path in enumerate(csv_files, start=1):
                print(
                    f"{GREEN}{number:>2}{RESET} "
                    f"{RED}-->{RESET} "
                    f"{CYAN}{csv_path.name}{RESET}"
                )
        else:
            print(f"{INFO} No indexes were found.\n")

        print(f"{GREEN} 0{RESET} {RED}-->{RESET} {CYAN}Create a new index{RESET}")
        print(SEPARATOR)

        selection = input("Selection: ").strip()

        if selection == "0":
            return create_index_file(directory)

        try:
            selected_number = int(selection)
            return csv_files[selected_number - 1]
        except (ValueError, IndexError):
            pause(f"{ERROR} Invalid selection. Press Enter to try again.")


# ---------------------------------------------------------------------------
# CSV operations
# ---------------------------------------------------------------------------


def normalise_entry(entry: dict[str, str]) -> dict[str, str]:
    """Return an entry using the current column names."""
    return {
        # Fall back to "Title" so indexes created before this column was
        # renamed still load correctly.
        "Term": entry.get("Term", entry.get("Title", "")).strip(),
        "Book": entry.get("Book", entry.get("book", "")).strip(),
        "Page": entry.get("Page", "").strip(),
        "Note": entry.get("Note", "").strip(),
    }


def migrate_index_file(csv_path: Path) -> None:
    """
    Bring an index file's header and column order up to date.

    Indexes created before the Note column (or the Title -> Term rename)
    still have their original header row on disk. Appending new rows to
    such a file with today's column order silently corrupts Book/Page and
    drops Note, because csv.DictReader maps columns by name using the
    header row, not by position. Rewriting the file here, once, before any
    further appends, prevents that.
    """
    try:
        with csv_path.open("r", newline="", encoding="utf-8") as csv_file:
            header = next(csv.reader(csv_file), [])
    except OSError:
        return

    if header == CSV_HEADERS:
        return

    entries = load_entries(csv_path)

    try:
        with csv_path.open("w", newline="", encoding="utf-8") as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(CSV_HEADERS)

            for entry in entries:
                writer.writerow([entry[column] for column in CSV_HEADERS])
    except OSError as error:
        print(
            f"{ERROR} Could not update {csv_path.name} to the current format: {error}"
        )
        return

    print(f"{INFO} Updated {CYAN}{csv_path.name}{RESET} to the current index format.")


def load_entries(csv_path: Path) -> list[dict[str, str]]:
    """Read all entries from an index."""
    try:
        with csv_path.open(
            "r",
            newline="",
            encoding="utf-8",
        ) as csv_file:
            reader = csv.DictReader(csv_file)
            return [normalise_entry(entry) for entry in reader]

    except (OSError, csv.Error) as error:
        print(f"{ERROR} Could not read {csv_path.name}: {error}")
        return []


def add_entry(
    csv_path: Path,
    term: str,
    book: int,
    page: int,
    note: str = "",
) -> None:
    """Append one entry to an index."""
    with csv_path.open(
        "a",
        newline="",
        encoding="utf-8",
    ) as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow([term, book, page, note])


# ---------------------------------------------------------------------------
# Index display
# ---------------------------------------------------------------------------


def print_index(csv_path: Path) -> None:
    entries = load_entries(csv_path)

    if not entries:
        print(f"{INFO} The index is empty.")
        return

    headers = ["#", *CSV_HEADERS]

    if len(entries) <= PREVIEW_ROWS * 2:
        visible = list(enumerate(entries, start=1))
    else:
        visible = (
            list(enumerate(entries[:PREVIEW_ROWS], start=1))
            + [(None, None)]
            + list(
                enumerate(
                    entries[-PREVIEW_ROWS:],
                    start=len(entries) - PREVIEW_ROWS + 1,
                )
            )
        )

    rows = []

    for number, entry in visible:
        if entry is None:
            rows.append(["...", "...", "...", "...", "..."])
        else:
            rows.append(
                [
                    str(number),
                    entry["Term"],
                    entry["Book"],
                    entry["Page"],
                    entry["Note"],
                ]
            )

    widths = [
        max(
            len(headers[col]),
            *(len(row[col]) for row in rows),
        )
        for col in range(len(headers))
    ]

    print(" | ".join(h.ljust(w) for h, w in zip(headers, widths)))
    print("-+-".join("-" * w for w in widths))

    for row in rows:
        print(" | ".join(v.ljust(w) for v, w in zip(row, widths)))

    print(f"\nTotal entries: {CYAN}{len(entries)}{RESET}")


# ---------------------------------------------------------------------------
# Term history
# ---------------------------------------------------------------------------


def load_term_history(
    csv_path: Path,
    limit: int = HISTORY_LIMIT,
) -> deque[str]:
    """Load the latest terms from an index."""
    terms = [entry["Term"] for entry in load_entries(csv_path) if entry["Term"]]

    return deque(terms[-limit:], maxlen=limit)


def load_readline_history(term_history: deque[str]) -> None:
    """Load recent terms into readline history."""
    if readline:
        readline.clear_history()
        for term in term_history:
            readline.add_history(term)
        readline.set_history_length(term_history.maxlen or HISTORY_LIMIT)


# ---------------------------------------------------------------------------
# Input handling
# ---------------------------------------------------------------------------


def prompt_for_book(current_book: int | None = None) -> int:
    """Prompt for a book number."""
    while True:
        if current_book is None:
            prompt = "Book number: "
        else:
            prompt = f"Book number [{current_book}]: "

        value = input(f"{YELLOW}{prompt}{RESET}").strip()

        if value and current_book is not None:
            return current_book

        if value.isdigit() and int(value) > 0:
            return int(value)

        print(f"{ERROR} The book number cannot be empty.")


def prompt_for_page(current_page: int | None = None) -> int:
    """Prompt for a positive page number."""
    while True:
        if current_page is None:
            prompt = "Starting page: "
        else:
            prompt = f"Page [{current_page}]: "

        value = input(f"{YELLOW}{prompt}{RESET}").strip()

        if not value and current_page is not None:
            return current_page

        if value.isdigit() and int(value) > 0:
            return int(value)

        print(f"{ERROR} Enter a positive page number.")


def parse_entry(
    user_input: str,
    current_page: int,
) -> tuple[str, int, str] | None:
    """
    Parse one of:

        Term
        Term, Page
        Term | Note
        Term, Page | Note

    A term without an explicit page uses the current page. A note is
    always optional and, when present, comes after a pipe so it can
    safely contain commas of its own.
    """
    value = user_input.strip()

    if not value:
        return None

    if NOTE_DELIMITER in value:
        value, note = value.rsplit(NOTE_DELIMITER, maxsplit=1)
        value = value.strip()
        note = note.strip()
    else:
        note = ""

    if not value:
        return None

    if "," not in value:
        return value, current_page, note

    term, page_text = value.rsplit(",", maxsplit=1)

    term = term.strip()
    page_text = page_text.strip()

    if not term:
        return None

    try:
        page = int(page_text)
    except ValueError:
        return None

    if page < 1:
        return None

    return term, page, note


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------


def print_header() -> None:
    """Display the current context and available commands."""
    print(f"{GREEN}SANS Terminal Indexer{RESET}")
    print(SEPARATOR)

    print(f"{GREEN}npage{RESET}     {MAGENTA}Next page{RESET}")
    print(f"{GREEN}ppage{RESET}     {MAGENTA}Previous page{RESET}")
    print(f"{GREEN}p<num>{RESET}    {MAGENTA}Jump to page <num>{RESET}")
    print(f"{GREEN}b<num>{RESET}    {MAGENTA}Change the current book{RESET}")
    print(f"{GREEN}e{RESET}         {MAGENTA}Edit entries{RESET}")
    print(f"{GREEN}index{RESET}     {MAGENTA}Switch index files{RESET}")
    print(f"{GREEN}help{RESET}      {MAGENTA}Show this command list{RESET}")
    print(f"{GREEN}exit{RESET}      {MAGENTA}Exit the program{RESET}")

    print(SEPARATOR)


def show_help() -> None:
    """Display detailed usage instructions."""
    clear_screen()

    print(f"{GREEN}SANS Terminal Indexer — Commands{RESET}")
    print(SEPARATOR)

    print("Add an entry using the current page:")
    print(f"  {CYAN}Windows Registry{RESET}")
    print()

    print("Add an entry with a different page:")
    print(f"  {CYAN}Windows Registry, 125{RESET}")
    print()

    print("Add an entry with a note (uses the current page):")
    print(f"  {CYAN}Windows Registry | HKLM hive locations{RESET}")
    print()

    print("Add an entry with a page and a note:")
    print(f"  {CYAN}Windows Registry, 125 | HKLM hive locations{RESET}")
    print()

    print("Page commands:")
    print(f"  {CYAN}npage{RESET}        Next page")
    print(f"  {CYAN}ppage{RESET}        Previous page")
    print(f"  {CYAN}p<num>{RESET}       Jump to page <num>")
    print()

    print("Editing commands:")
    print(f"  {CYAN}e{RESET}            Edit the latest item")
    print()

    print("Other commands:")
    print(f"  {CYAN}b<num>{RESET}       Change the current book")
    print(f"  {CYAN}index{RESET}        Switch index files")
    print(f"  {CYAN}exit{RESET}         Exit")

    print(SEPARATOR)
    pause()


# ---------------------------------------------------------------------------
# Command handling
# ---------------------------------------------------------------------------


def handle_page_command(
    command: str,
    current_page: int,
) -> int | None:
    """
    Handle page commands.

    Return:
        A new page number when the input is a page command.
        None when the input is not a page command.
    """
    if command == "npage":
        return current_page + 1

    if command == "ppage":
        return max(1, current_page - 1)

    if command.startswith("p") and command[1:].isdigit():
        requested_page = int(command[1:])

        if requested_page > 0:
            return requested_page

    return None


def handle_book_command(
    command: str,
    book: int,
) -> int | None:
    """
    Handle book commands.

    Return:
        A new book number when the input is a book command.
        None when the input is not a book command.
    """
    if command == "nbook":
        return book + 1

    if command == "pbook":
        return max(1, book - 1)

    if command.startswith("b") and command[1:].isdigit():
        requested_book = int(command[1:])

        if requested_book > 0:
            return requested_book

    return None


def edit_index(csv_path: Path) -> bool:
    """
    Open the current index in the user's editor.

    Returns True if the editor exited successfully.
    """
    editor = os.environ.get("EDITOR")

    if not editor:
        editor = shutil.which("nvim") or shutil.which("vim") or shutil.which("nano")

    if not editor:
        print(f"{ERROR} No editor found.")
        return False

    try:
        subprocess.run(
            [editor, "+normal G", str(csv_path)],
            check=True,
        )
    except subprocess.CalledProcessError:
        print(f"{ERROR} Editor exited with an error.")
        return False
    except OSError as error:
        print(f"{ERROR} Could not launch editor: {error}")
        return False

    return True


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------


def run() -> None:
    """Run the terminal indexer."""
    signal.signal(signal.SIGINT, handle_interrupt)

    directory = ensure_index_directory()
    csv_path = select_index_file(directory)
    migrate_index_file(csv_path)

    term_history = load_term_history(csv_path)

    clear_screen()
    book = prompt_for_book()
    current_page = prompt_for_page()

    while True:
        clear_screen()

        print_header()

        print_index(csv_path)
        print(SEPARATOR)

        load_readline_history(term_history)

        print(
            f"{SUCCESS} Enter a term, add "
            f"{GREEN}, Page{RESET} to override the current page, and/or add "
            f"{GREEN}| Note{RESET} for a short note.\n"
        )

        print(f"----------------------------------------------------------------")
        print(
            f"Index: {CYAN}{csv_path.name}{RESET}  "
            f"Book: {CYAN}{book}{RESET}  "
            f"Page: {CYAN}{current_page}{RESET}"
        )
        print(f"----------------------------------------------------------------")
        user_input = input(ENTRY_PROMPT).strip()
        print(RESET, end="")

        if not user_input:
            continue

        command = user_input.lower()

        if command == "exit":
            print(f"{SUCCESS} Index saved. Goodbye.")
            return

        if command == "help":
            show_help()
            continue

        if command == "index":
            csv_path = select_index_file(directory)
            migrate_index_file(csv_path)
            term_history = load_term_history(csv_path)

            clear_screen()
            book = prompt_for_book(book)
            current_page = prompt_for_page(current_page)
            continue

        new_book = handle_book_command(
            command,
            book,
        )

        new_page = handle_page_command(
            command,
            current_page,
        )

        if new_page is not None:
            if command == "ppage" and current_page == 1:
                pause(f"{ERROR} Page cannot be less than 1. Press Enter.")
            else:
                current_page = new_page

            continue

        if new_book is not None:
            book = new_book
            continue

        if command == "e":
            clear_screen()
            if edit_index(csv_path):
                term_history = load_term_history(csv_path)
            clear_screen()
            continue

        parsed_entry = parse_entry(
            user_input,
            current_page,
        )

        if parsed_entry is None:
            pause(
                f"{ERROR} Invalid entry.\n\n"
                "Use one of these formats:\n"
                "  Term\n"
                "  Term, Page\n"
                "  Term | Note\n"
                "  Term, Page | Note\n\n"
                "Press Enter to try again."
            )
            continue

        term, page, note = parsed_entry

        try:
            add_entry(
                csv_path=csv_path,
                term=term,
                book=book,
                page=page,
                note=note,
            )
        except OSError as error:
            pause(
                f"{ERROR} Could not save the entry: {error}\nPress Enter to continue."
            )
            continue

        term_history.append(term)


if __name__ == "__main__":
    run()
