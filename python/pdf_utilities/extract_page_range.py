"""Utility helpers to extract a range of pages from one or more PDF files.

Example usages
--------------
Extract a subset of pages from a single PDF and store it alongside the source::

    from pathlib import Path
    from python.pdf_utilities.extract_page_range import extract_page_ranges

    extract_page_ranges(
        pdfs=[Path("report.pdf")],
        start=1,
        end=3,
        output_dir=Path("out"),
    )

Process multiple PDFs in one call to build a collection of slices::

    extract_page_ranges(
        pdfs=[Path("chapter1.pdf"), Path("chapter2.pdf")],
        start=2,
        end=5,
        output_dir=Path("slices"),
    )

Combine the helpers with :meth:`pathlib.Path.glob` to handle folders::

    pdfs = sorted(Path("data").glob("*.pdf"))
    extract_page_ranges(
        pdfs=pdfs,
        start=10,
        end=12,
        output_dir=Path("subset"),
    )

These examples can be executed from interactive Python sessions or editor
run-configurations without relying on a command-line entry point.
"""

from __future__ import annotations

from pathlib import Path
from typing import Sequence

from PyPDF2 import PdfReader, PdfWriter


def extract_page_range(pdf_path: Path, start: int, end: int, output_path: Path) -> Path:
    """Extract pages from *start* to *end* from *pdf_path* and write to *output_path*.

    Parameters
    ----------
    pdf_path : Path
        The source PDF file.
    start : int
        The first page to extract (1-indexed).
    end : int
        The last page to extract (inclusive, 1-indexed).
    output_path : Path
        Where the resulting PDF will be written.

    Returns
    -------
    Path
        The path to the written PDF file.
    """
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()

    # Convert to 0-indexed for PyPDF2 and iterate over the requested pages
    for page in reader.pages[start - 1:end]:
        writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("wb") as fh:
        writer.write(fh)
    return output_path


def extract_page_ranges(
    pdfs: Sequence[Path | str],
    start: int,
    end: int,
    output_dir: Path | str,
) -> list[Path]:
    """Extract the same page range from multiple PDFs.

    The resulting files are written inside *output_dir* with the pattern
    ``<stem>_pages_<start>-<end><suffix>``. The output directory is created if it
    does not already exist.

    Parameters
    ----------
    pdfs : Sequence[Path | str]
        Collection of input PDF paths.
    start : int
        First page to extract (1-indexed).
    end : int
        Last page to extract (inclusive, 1-indexed).
    output_dir : Path | str
        Directory where the extracted PDFs will be saved.

    Returns
    -------
    list[Path]
        Paths to the generated PDF files, in the same order as ``pdfs``.
    """
    output_directory = Path(output_dir)
    output_directory.mkdir(parents=True, exist_ok=True)

    written_files: list[Path] = []
    for pdf in pdfs:
        pdf_path = Path(pdf)
        destination = output_directory / f"{pdf_path.stem}_pages_{start}-{end}{pdf_path.suffix}"
        written_files.append(
            extract_page_range(
                pdf_path=pdf_path,
                start=start,
                end=end,
                output_path=destination,
            )
        )
    return written_files
