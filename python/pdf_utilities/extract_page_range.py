import argparse
from pathlib import Path
from PyPDF2 import PdfReader, PdfWriter


def extract_page_range(pdf_path: Path, start: int, end: int, output_path: Path) -> None:
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
    """
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()

    # Convert to 0-indexed for PyPDF2 and iterate over the requested pages
    for page in reader.pages[start - 1:end]:
        writer.add_page(page)

    with output_path.open("wb") as fh:
        writer.write(fh)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract a range of pages from multiple PDF files.")
    parser.add_argument(
        "pdfs",
        nargs="+",
        help="One or more source PDF paths",
    )
    parser.add_argument(
        "--start",
        required=True,
        type=int,
        help="First page to extract (1-indexed)",
    )
    parser.add_argument(
        "--end",
        required=True,
        type=int,
        help="Last page to extract (inclusive, 1-indexed)",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        type=Path,
        help="Directory where extracted PDFs will be saved",
    )
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)

    for pdf in args.pdfs:
        pdf_path = Path(pdf)
        out_name = f"{pdf_path.stem}_pages_{args.start}-{args.end}{pdf_path.suffix}"
        extract_page_range(pdf_path, args.start, args.end, args.output_dir / out_name)


if __name__ == "__main__":
    main()
