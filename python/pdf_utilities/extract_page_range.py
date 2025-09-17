"""Interactive utility for extracting page ranges from PDF files.

This module provides both an interactive Tk-based workflow and programmatic
helpers for slicing the same page range from one or many PDFs.  The script can
be launched directly from the terminal::

    python python/pdf_utilities/extract_page_range.py

Running it presents file pickers and dialogs for selecting the source PDFs, an
output directory, and the page range to extract.  The lower-level helpers are
fully documented and may also be imported and reused from other Python code.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Sequence, Tuple

from PyPDF2 import PdfReader, PdfWriter

import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox


def parse_range(range_str: str) -> Tuple[int, int]:
    """Parse a range string like ``"260-278"`` into a ``(start, end)`` tuple.

    Parameters
    ----------
    range_str : str
        A string representation of the desired page range, using a hyphen to
        separate the starting and ending page numbers.

    Returns
    -------
    Tuple[int, int]
        A tuple containing the inclusive start and end page numbers as
        integers.

    Raises
    ------
    ValueError
        If the string does not contain exactly one hyphen, if either number
        cannot be converted to an integer, or if the numeric range is invalid
        (e.g. negative numbers or start greater than end).
    """

    try:
        # ``str.split("-")`` yields two substrings for inputs like "10-25".
        # ``str.strip()`` removes any surrounding whitespace the user may have
        # entered, allowing inputs such as " 10 - 25 ".
        start_str, end_str = range_str.split("-")
        start, end = int(start_str.strip()), int(end_str.strip())

        # Guard against inverted or out-of-bounds ranges by validating that
        # both numbers are positive and that the start page comes before the
        # end page.
        if start < 1 or end < 1 or start > end:
            raise ValueError
        return start, end
    except Exception:  # pragma: no cover - defensive user-facing validation
        # Provide a descriptive error message that can be surfaced in dialogs.
        raise ValueError(
            f"Invalid range format: '{range_str}'. Use e.g. '10-25'."
        )


def extract_page_range(pdf_path: Path, start: int, end: int, output_path: Path) -> None:
    """Write a new PDF containing pages ``start`` through ``end`` (inclusive).

    Parameters
    ----------
    pdf_path : Path
        The input PDF file that should be sliced.
    start : int
        The first page to extract, using 1-based indexing.
    end : int
        The last page to extract, using 1-based indexing.
    output_path : Path
        Destination path for the generated PDF file.  The parent directories
        are created automatically if they do not exist.

    Raises
    ------
    ValueError
        If ``start`` lies beyond the end of the document.
    """

    # ``PdfReader`` lazily loads the source document and exposes its pages via
    # ``reader.pages``.  ``PdfWriter`` accumulates the chosen pages and writes
    # them to disk once we are finished.
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()

    num_pages = len(reader.pages)
    if start > num_pages:
        raise ValueError(
            f"Range {start}-{end} exceeds PDF '{pdf_path.name}' ({num_pages} pages)."
        )

    # Clamp the end value so that asking for pages beyond the document simply
    # stops at the last available page.  This mirrors how many PDF tools behave
    # and prevents index errors.
    end = min(end, num_pages)

    # PyPDF2 pages are 0-indexed, hence the ``start - 1`` offset.  Slicing is
    # inclusive of ``start`` and exclusive of ``end``, so we rely on Python's
    # slice semantics with ``end`` already clamped above.
    for page in reader.pages[start - 1 : end]:
        writer.add_page(page)

    # Ensure the output directory exists before attempting to write the file.
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as fh:
        writer.write(fh)


def process_many(pdfs: Sequence[Path], start: int, end: int, output_dir: Path) -> None:
    """Extract the same page range from multiple PDFs and write them to disk.

    Parameters
    ----------
    pdfs : Sequence[Path]
        Iterable of input PDF file paths.
    start : int
        Inclusive starting page number (1-indexed).
    end : int
        Inclusive ending page number (1-indexed).
    output_dir : Path
        Folder where the generated PDFs should be placed.  Created if missing.
    """

    output_dir.mkdir(parents=True, exist_ok=True)

    for pdf_path in pdfs:
        # Construct a descriptive filename that indicates the source document
        # and the extracted page range, mirroring the behaviour of the GUI.
        out_name = f"{pdf_path.stem}_pages_{start}-{end}{pdf_path.suffix}"
        dest = output_dir / out_name
        print(f"→ Extracting {pdf_path.name} [{start}-{end}] -> {dest}")
        extract_page_range(pdf_path, start, end, dest)

    print("✅ Done.")


def run_interactive() -> None:
    """Launch the Tkinter-powered workflow for selecting PDFs and ranges."""

    # Initialising ``Tk`` creates the root window.  We hide it immediately to
    # avoid displaying a blank window before the dialogs appear.
    root = tk.Tk()
    root.withdraw()

    # --- Select source PDFs -------------------------------------------------
    file_paths = filedialog.askopenfilenames(
        title="Select one or more PDF files", filetypes=[("PDF files", "*.pdf")]
    )
    if not file_paths:
        print("No files selected. Exiting.")
        return
    pdfs = [Path(p) for p in file_paths]

    # --- Choose destination directory --------------------------------------
    out_dir_str = filedialog.askdirectory(title="Select output directory")
    if not out_dir_str:
        print("No output directory selected. Exiting.")
        return
    output_dir = Path(out_dir_str)

    # --- Ask for the desired page range ------------------------------------
    range_str = simpledialog.askstring(
        "Page range", "Enter page range (e.g. 10-25):"
    )
    if not range_str:
        print("No range provided. Exiting.")
        return

    try:
        start, end = parse_range(range_str)
    except ValueError as e:
        # When validation fails we show the error in a modal dialog so the user
        # understands what needs to be corrected.
        messagebox.showerror("Input error", str(e))
        return

    # --- Perform extraction for each selected PDF --------------------------
    errors = []
    for pdf in pdfs:
        try:
            out_name = f"{pdf.stem}_pages_{start}-{end}{pdf.suffix}"
            dest = output_dir / out_name
            print(f"→ Extracting {pdf.name} [{start}-{end}] -> {dest}")
            extract_page_range(pdf, start, end, dest)
        except Exception as e:  # pragma: no cover - error reporting path
            errors.append((pdf.name, str(e)))
            print(f"✖ Failed for {pdf.name}: {e}")

    # Summarise the outcome with a message box so the user receives feedback
    # without needing to watch the terminal.
    if errors:
        msg = "\n".join(f"- {name}: {err}" for name, err in errors)
        messagebox.showwarning("Finished with errors", f"Some files failed:\n\n{msg}")
    else:
        messagebox.showinfo("Success", f"Finished! Files saved to:\n{output_dir}")
        print("✅ Done.")


if __name__ == "__main__":
    try:
        run_interactive()
    except Exception as exc:  # pragma: no cover - protects CLI usage
        print(f"Unexpected error: {exc}", file=sys.stderr)
        raise
