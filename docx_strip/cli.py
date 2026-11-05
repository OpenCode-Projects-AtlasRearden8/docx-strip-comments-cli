"""Command line interface: python -m docx_strip input.docx [-o output.docx]"""

import argparse
import sys
from pathlib import Path

from . import DEFAULT_SUFFIX, DocxStripError, __version__, strip_docx


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docx-strip",
        description=(
            "Copy a .docx and strip tracked changes, comments, and personal "
            "metadata so the clean copy is safe to share. The original file "
            "is never modified. Runs fully offline."
        ),
        epilog=(
            "Tracked insertions are kept (as if all changes were accepted); "
            "tracked deletions, comments, author names, revision history, "
            "editing-time statistics, and rsid session data are removed."
        ),
    )
    parser.add_argument("input", help="the .docx file to clean (left untouched)")
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="where to write the clean copy (default: <input>_clean.docx)",
    )
    parser.add_argument(
        "--suffix",
        default=DEFAULT_SUFFIX,
        help="suffix for the default output name (default: _clean)",
    )
    parser.add_argument(
        "-q", "--quiet", action="store_true", help="suppress the summary report"
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}"
    )
    return parser


def main(argv=None) -> int:
    args = build_parser().parse_args(argv)
    src = Path(args.input)
    try:
        report = strip_docx(src, args.output, args.suffix)
    except DocxStripError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1
    if not args.quiet:
        print(report.summary())
    return 0


if __name__ == "__main__":
    sys.exit(main())
