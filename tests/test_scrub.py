"""Tests for the scrubbing API: docx_strip.scrub.strip_docx()."""

import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from docx_strip import DEFAULT_SUFFIX, DocxStripError, strip_docx

from fixture_builder import (
    DELETED_TEXT,
    FAKE_AUTHOR,
    FAKE_COMMENT,
    FAKE_INITIALS,
    INSERTED_TEXT,
    MOVE_DEST_TEXT,
    MOVE_SOURCE_TEXT,
    STATIC_TEXT,
    _cellar,
    write_docx,
)

RSID_VALUES = [b"00AA11BB", b"00CC22DD", b"00EE33FF", b"00EE44AA"]


def members(path) -> dict:
    """Every archive entry, decompressed."""
    out = {}
    with zipfile.ZipFile(path) as zf:
        for name in zf.namelist():
            out[name] = zf.read(name)
    return out


def document_xml(path) -> ET.Element:
    root = ET.fromstring(members(path)["word/document.xml"])
    return root


def body_text(root) -> str:
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    return "".join(
        (elem.text or "")
        for elem in root.iter(f"{{{w_ns}}}t")
    )


@pytest.fixture
def cleaned(messy_docx, tmp_path) -> Path:
    dst = tmp_path / "out.docx"
    strip_docx(messy_docx, dst)
    return dst


class TestCoreBehavior:
    def test_original_file_is_untouched(self, messy_docx, original_bytes, tmp_path):
        strip_docx(messy_docx, tmp_path / "out.docx")
        assert messy_docx.read_bytes() == original_bytes

    def test_output_exists_and_is_valid_zip(self, cleaned):
        with zipfile.ZipFile(cleaned) as zf:
            assert zf.testzip() is None
        assert "word/document.xml" in members(cleaned)

    def test_inserted_text_is_kept(self, cleaned):
        text = body_text(document_xml(cleaned))
        assert INSERTED_TEXT in text
        assert "nested revised text" in text

    def test_deleted_text_is_gone_everywhere(self, cleaned):
        for name, data in members(cleaned).items():
            assert DELETED_TEXT.encode() not in data, f"leak in {name}"

    def test_move_tracking_resolved(self, cleaned):
        text = body_text(document_xml(cleaned))
        assert MOVE_DEST_TEXT in text
        for name, data in members(cleaned).items():
            assert MOVE_SOURCE_TEXT.encode() not in data, f"leak in {name}"

    def test_comment_content_is_gone_everywhere(self, cleaned):
        for name, data in members(cleaned).items():
            assert FAKE_COMMENT.encode() not in data, f"leak in {name}"

    def test_comment_parts_are_removed(self, cleaned):
        names = set(members(cleaned))
        for name in ["word/comments.xml", "word/people.xml", "docProps/app.xml",
                     "docProps/custom.xml", "docProps/thumbnail.jpeg",
                     "customXml/item1.xml"]:
            assert name not in names

    def test_comment_marks_are_removed_from_document(self, cleaned):
        root = document_xml(cleaned)
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        leftovers = [
            elem.tag for elem in root.iter()
            if elem.tag.rsplit("}", 1)[-1]
            in ("commentRangeStart", "commentRangeEnd", "commentReference")
        ]
        assert leftovers == []

    def test_author_never_leaks(self, cleaned):
        for name, data in members(cleaned).items():
            assert FAKE_AUTHOR.encode() not in data, f"leak in {name}"

    def test_author_initials_never_leak(self, cleaned):
        for name, data in members(cleaned).items():
            assert FAKE_INITIALS.encode() not in data, f"leak in {name}"

    def test_rsids_are_gone(self, cleaned):
        root = ET.fromstring(members(cleaned)["word/document.xml"])
        rsid_tags = [
            attr for elem in root.iter() for attr in elem.attrib
            if attr.rsplit("}", 1)[-1].startswith("rsid")
        ]
        assert rsid_tags == []
        for name, data in members(cleaned).items():
            for value in RSID_VALUES:
                assert value not in data, f"rsid leak in {name}: {value}"

    def test_settings_has_no_trackchanges_or_rsids(self, cleaned):
        settings = members(cleaned)["word/settings.xml"].decode()
        assert "trackChanges" not in settings
        assert "rsid" not in settings.lower()

    def test_core_metadata_is_scrubbed(self, cleaned):
        core = members(cleaned)["docProps/core.xml"].decode()
        assert FAKE_AUTHOR not in core
        assert "<dc:creator" not in core
        assert "lastModifiedBy" not in core
        assert "<cp:revision" not in core
        assert "dcterms:created" not in core
        assert "dcterms:modified" not in core
        assert "Sample Title" in core  # chosen title is preserved

    def test_relationships_point_only_at_surviving_parts(self, cleaned):
        doc_rels = members(cleaned)["word/_rels/document.xml.rels"].decode()
        assert "comments.xml" not in doc_rels
        assert "people.xml" not in doc_rels
        assert "customXml" not in doc_rels
        root_rels = members(cleaned)["_rels/.rels"].decode()
        assert "custom.xml" not in root_rels

    def test_content_types_no_longer_list_dropped_parts(self, cleaned):
        ct = members(cleaned)["[Content_Types].xml"].decode()
        for leaked in ["comments.xml", "app.xml", "custom.xml", "thumbnail"]:
            assert leaked not in ct

    def test_zip_timestamps_are_normalized(self, cleaned):
        with zipfile.ZipFile(cleaned) as zf:
            for info in zf.infolist():
                assert info.date_time == (1980, 1, 1, 0, 0, 0)

    def test_report_counts_match(self, messy_docx, tmp_path):
        report = strip_docx(messy_docx, tmp_path / "out.docx")
        assert report.insertions_kept == 3      # ins + nested ins + moveTo
        assert report.deletions_removed == 3    # del + para-mark del + moveFrom
        assert report.change_records_removed == 1
        assert report.comment_marks_removed == 5
        assert report.range_markers_removed == 2
        assert len(report.parts_dropped) == 6   # comments, people, app, custom,
        # thumbnail, customXml/item1.xml (no orphan .rels sidecars existed)

    def test_document_survives_schema_shaped(self, cleaned):
        root = document_xml(cleaned)
        assert root.tag.endswith("document")
        # price of acceptance: markup remains well-formed
        ET.fromstring(members(cleaned)["docProps/core.xml"])
        ET.fromstring(members(cleaned)["word/settings.xml"])


class TestReportAndDefaults:
    def test_default_output_name_uses_suffix(self, messy_docx):
        report = strip_docx(messy_docx)
        assert Path(report.output).name == f"messy{DEFAULT_SUFFIX}.docx"
        assert Path(report.output).is_file()

    def test_custom_output_path(self, messy_docx, tmp_path):
        custom = tmp_path / "deep" / "clean.docx"
        custom.parent.mkdir()
        report = strip_docx(messy_docx, custom)
        assert report.output == str(custom)
        assert custom.is_file()

    def test_custom_suffix(self, messy_docx):
        report = strip_docx(messy_docx, suffix="_share-ready")
        assert Path(report.output).name == "messy_share-ready.docx"

    def test_summary_is_runnable_text(self, messy_docx, tmp_path):
        report = strip_docx(messy_docx, tmp_path / "out.docx")
        s = report.summary()
        assert str(messy_docx) in s
        assert "Parts dropped" in s


class TestPlainDocuments:
    def test_plain_document_round_trips(self, plain_docx, tmp_path):
        out = tmp_path / "clean.docx"
        report = strip_docx(plain_docx, out)
        assert report.insertions_kept == 0
        assert report.deletions_removed == 0
        assert report.parts_dropped == []
        assert body_text(document_xml(out)) == STATIC_TEXT

    def test_scrubbing_is_idempotent(self, cleaned, tmp_path):
        second = strip_docx(cleaned, tmp_path / "second.docx")
        assert second.parts_dropped == []
        assert body_text(document_xml(tmp_path / "second.docx")) == \
            body_text(document_xml(cleaned))


class TestErrors:
    def test_missing_file(self, tmp_path):
        with pytest.raises(DocxStripError, match="not found"):
            strip_docx(tmp_path / "nope.docx")

    def test_not_a_zip(self, tmp_path):
        target = tmp_path / "faked.docx"
        target.write_bytes(b"This is not a zip at all.")
        with pytest.raises(DocxStripError, match="not a valid .docx"):
            strip_docx(target)

    def test_zip_without_document_part(self, tmp_path):
        target = tmp_path / "similar.docx"
        write_docx(target, {"[Content_Types].xml": _cellar()["[Content_Types].xml"]})
        with pytest.raises(DocxStripError, match="not a Word"):
            strip_docx(target)

    def test_word_docx_never_written_when_invalid(self, tmp_path):
        target = tmp_path / "faked.docx"
        target.write_bytes(b"Not a zip.")
        out = tmp_path / "out.docx"
        with pytest.raises(DocxStripError):
            strip_docx(target, out)
        assert not out.exists()
