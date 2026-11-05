"""Shared pytest fixtures built on top of tests/fixture_builder.py."""

import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from fixture_builder import _cellar, _plain_document_xml, write_docx  # noqa: E402


@pytest.fixture
def messy_docx(tmp_path) -> Path:
    """A document with tracked changes, comments, and personal metadata."""
    path = tmp_path / "messy.docx"
    write_docx(path, _cellar())
    return path


@pytest.fixture
def plain_docx(tmp_path) -> Path:
    """A document with no revisions and no metadata parts to carry them."""
    path = tmp_path / "plain.docx"
    cellar = _cellar()
    write_docx(
        path,
        {
            "[Content_Types].xml": cellar["[Content_Types].xml"],
            "_rels/.rels": cellar["_rels/.rels"],
            "word/document.xml": _plain_document_xml(),
        },
    )
    return path


@pytest.fixture
def original_bytes(messy_docx: Path) -> bytes:
    return messy_docx.read_bytes()
