"""Core scrubbing logic.

A .docx file is a ZIP archive of XML parts. This module copies the archive
part by part into a *new* file and never opens the original for writing.

* Tracked changes in ``word/document.xml`` (plus headers, footers,
  footnotes, endnotes, and glossary) are resolved as if every change were
  accepted: inserted revisions are kept, deleted revisions are removed,
  and all revision markers, change records, and rsid attributes are dropped.
* Comment parts (``word/comments*.xml``, ``word/people.xml``) are removed
  entirely, along with their references, content-type overrides, and
  relationships.
* Personal metadata is removed: ``docProps/core.xml`` author /
  lastModifiedBy / revision / date fields are dropped, ``docProps/app.xml``
  (company, manager, editing-time statistics) and ``docProps/custom.xml``
  are removed wholesale, ``w:trackChanges`` and ``w:rsids`` session data
  are removed from settings, thumbnails and customXml data parts are
  dropped, and every ZIP entry timestamp is normalized.
"""

import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DCTERMS = "http://purl.org/dc/terms/"

# Fixed timestamp for every entry so the copy leaks no edit-session timing.
FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)

DEFAULT_SUFFIX = "_clean"

_XML_DECL = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

# Tracked revisions whose content should be kept (i.e. accepted).
_UNWRAP_TAGS = {"ins", "moveTo"}

# Revision plumbing and comment marks whose content must disappear.
_DROP_TAGS = {
    "del",
    "moveFrom",
    "rPrChange",
    "pPrChange",
    "sectPrChange",
    "tblPrChange",
    "trPrChange",
    "tcPrChange",
    "numberingChange",
    "cellIns",
    "cellDel",
    "cellMerge",
    "commentRangeStart",
    "commentRangeEnd",
    "commentReference",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
}

# Attributes that carry author/date/session fingerprints.
_ATTR_RSID = re.compile(r"^rsid")
_ATTR_AUTHOR = {"author", "authorInitials", "date"}

# Parts removed wholesale from the archive.
_DROP_PART_RE = re.compile(
    r"^(word/(comments[^/]*|commentInformation|people)\.xml|"
    r"customXml/.+|docProps/app\.xml|docProps/custom\.xml|"
    r"docProps/thumbnail\.[A-Za-z0-9]+)$"
)

# Parts that get the full tracked-changes transform.
_BODY_PART_RE = re.compile(
    r"^(word/document\.xml|word/footnotes\.xml|word/endnotes\.xml|"
    r"word/header[0-9]+\.xml|word/footer[0-9]+\.xml|word/glossary/document\.xml)$"
)

_SETTINGS_DROP_TAGS = {"trackChanges", "rsids", "writeProtection"}

_CORE_DROP_TAGS = {
    f"{{{NS_DC}}}creator",
    f"{{{NS_CP}}}lastModifiedBy",
    f"{{{NS_CP}}}revision",
    f"{{{NS_DCTERMS}}}created",
    f"{{{NS_DCTERMS}}}modified",
}


@dataclass
class StripReport:
    """Summary of what was removed from one document."""

    source: str = ""
    output: str = ""
    parts_dropped: list = field(default_factory=list)
    insertions_kept: int = 0
    deletions_removed: int = 0
    change_records_removed: int = 0
    comment_marks_removed: int = 0
    range_markers_removed: int = 0
    rsid_attrs_removed: int = 0
    author_attrs_removed: int = 0
    core_fields_removed: int = 0
    settings_items_removed: int = 0
    relationships_removed: int = 0
    content_types_removed: int = 0

    def summary(self) -> str:
        return "\n".join(
            [
                f"Source: {self.source}",
                f"Clean copy: {self.output}",
                "Parts dropped: " + (", ".join(self.parts_dropped) or "none"),
                f"Tracked insertions kept: {self.insertions_kept}",
                f"Tracked deletions removed: {self.deletions_removed}",
                f"Change records removed: {self.change_records_removed}",
                f"Comment marks removed: {self.comment_marks_removed}",
                f"Revision range markers removed: {self.range_markers_removed}",
                f"Author/date attributes removed: {self.author_attrs_removed}",
                f"rsid attributes removed: {self.rsid_attrs_removed}",
                f"Core metadata fields removed: {self.core_fields_removed}",
                f"Settings items removed: {self.settings_items_removed}",
                f"Relationships removed: {self.relationships_removed}",
                f"Content-type entries removed: {self.content_types_removed}",
            ]
        )


class DocxStripError(Exception):
    """Raised for unreadable or non-.docx inputs."""


def _localname(tag) -> str:
    if isinstance(tag, str) and tag.startswith("{"):
        return tag[tag.index("}") + 1 :]
    return str(tag)


def _resolve_tracked_changes(root) -> dict:
    """Accept all tracked changes in place. Returns per-type counters."""

    counts = {
        "unwrapped": 0,
        "dropped": 0,
        "change_records": 0,
        "comment_marks": 0,
        "range_markers": 0,
    }

    def walk(elem) -> None:
        idx = 0
        while idx < len(elem):
            child = elem[idx]
            name = _localname(child.tag)
            if name in _DROP_TAGS:
                if name.startswith("comment") or name.startswith("commentRange"):
                    counts["comment_marks"] += 1
                elif name.endswith("Change"):
                    counts["change_records"] += 1
                elif name.endswith("RangeStart") or name.endswith("RangeEnd"):
                    counts["range_markers"] += 1
                else:
                    counts["dropped"] += 1
                elem.remove(child)
            elif name in _UNWRAP_TAGS:
                # Accept the revision: recursively resolve its subtree,
                # then splice the surviving children into the parent.
                walk(child)
                grandchildren = list(child)
                elem.remove(child)
                for offset, grandchild in enumerate(grandchildren):
                    elem.insert(idx + offset, grandchild)
                counts["unwrapped"] += 1
                idx += len(grandchildren)
            else:
                walk(child)
                idx += 1

    walk(root)
    return counts


def _clean_revision_attributes(root) -> dict:
    """Drop author/date/rsid fingerprint attributes document-wide."""
    rsid = 0
    author = 0
    for elem in root.iter():
        for attr in list(elem.attrib):
            local = _localname(attr)
            if _ATTR_RSID.match(local):
                del elem.attrib[attr]
                rsid += 1
            elif local in _ATTR_AUTHOR and attr.startswith("{%s}" % W_NS):
                del elem.attrib[attr]
                author += 1
    return {"rsid": rsid, "author": author}


def _strip_document_part(data: bytes, report: StripReport) -> bytes:
    root = ET.fromstring(data)
    counts = _resolve_tracked_changes(root)
    attrs = _clean_revision_attributes(root)
    report.insertions_kept += counts["unwrapped"]
    report.deletions_removed += counts["dropped"]
    report.change_records_removed += counts["change_records"]
    report.comment_marks_removed += counts["comment_marks"]
    report.range_markers_removed += counts["range_markers"]
    report.rsid_attrs_removed += attrs["rsid"]
    report.author_attrs_removed += attrs["author"]
    ET.register_namespace("", W_NS)
    return _XML_DECL + ET.tostring(root, encoding="utf-8")


def _strip_core_part(data: bytes, report: StripReport) -> bytes:
    root = ET.fromstring(data)
    removed = 0
    for parent in root.iter():
        for child in list(parent):
            if child.tag in _CORE_DROP_TAGS:
                parent.remove(child)
                removed += 1
    report.core_fields_removed += removed
    ET.register_namespace("", NS_CP)
    return _XML_DECL + ET.tostring(root, encoding="utf-8")


def _strip_settings_part(data: bytes, report: StripReport) -> bytes:
    root = ET.fromstring(data)
    removed = 0
    for parent in root.iter():
        for child in list(parent):
            if _localname(child.tag) in _SETTINGS_DROP_TAGS:
                parent.remove(child)
                removed += 1
    report.settings_items_removed += removed
    ET.register_namespace("", W_NS)
    return _XML_DECL + ET.tostring(root, encoding="utf-8")


def _rels_base_dir(rels_name: str) -> str:
    """Directory of the part that owns a .rels file.

    ``_rels/.rels`` -> "" (package root)
    ``word/_rels/document.xml.rels`` -> "word/"
    """
    if rels_name.startswith("_rels/"):
        return ""
    return rels_name.split("_rels/")[0]


def _resolve_rel_target(base_dir: str, target: str) -> str:
    """Resolve an OPC relationship target to a part name."""
    if target.startswith("/"):
        return target[1:]
    base = base_dir
    while target.startswith("../"):
        target = target[3:]
        if base:
            base = base.rstrip("/")
            if "/" in base:
                base = base.rsplit("/", 1)[0] + "/"
            else:
                base = ""
        else:
            return target  # escapes the package: cannot resolve
    return base + target


def _strip_relationships(data: bytes, base_dir: str, dropped_parts: set):
    root = ET.fromstring(data)
    removed = 0
    for rel in list(root):
        if (rel.get("TargetMode") or "Internal") == "External":
            continue
        resolved = _resolve_rel_target(base_dir, rel.get("Target") or "")
        if resolved in dropped_parts:
            root.remove(rel)
            removed += 1
    return _XML_DECL + ET.tostring(root, encoding="utf-8"), removed


def _strip_content_types(data: bytes, dropped_parts: set):
    root = ET.fromstring(data)
    removed = 0
    for override in list(root):
        part_name = (override.get("PartName") or "").lstrip("/")
        if part_name in dropped_parts:
            root.remove(override)
            removed += 1
    return _XML_DECL + ET.tostring(root, encoding="utf-8"), removed


def strip_docx(source, output=None, suffix: str = DEFAULT_SUFFIX) -> StripReport:
    """Copy *source* .docx to *output* with tracked changes, comments, and
    personal metadata removed. The source file is never modified.

    Args:
        source: Path to the original .docx.
        output: Path for the clean copy. Defaults to ``<name>_clean.docx``
            next to the source.
        suffix: Suffix used when deriving the default output name.

    Returns:
        StripReport describing everything that was removed.

    Raises:
        DocxStripError: if the source is missing, not a ZIP, or not a
            Word document (no ``word/document.xml`` part).
    """
    src = Path(source)
    if not src.is_file():
        raise DocxStripError(f"input file not found: {src}")
    if output is None:
        dst = src.with_name(src.stem + suffix + ".docx")
    else:
        dst = Path(output)

    report = StripReport(source=str(src), output=str(dst))

    try:
        with zipfile.ZipFile(src, "r") as zin:
            bad = zin.testzip()
            if bad is not None:
                raise DocxStripError(f"corrupt archive entry: {bad}")
            names = zin.namelist()
            if "word/document.xml" not in names:
                raise DocxStripError(
                    f"not a Word (.docx) document: {src} "
                    "(missing word/document.xml part)"
                )

            dropped = {n for n in names if _DROP_PART_RE.match(n)}
            for part in list(dropped):
                parent = part.rsplit("/", 1)[0] if "/" in part else ""
                if parent:
                    dropped.add(f"{parent}/_rels/{part.rsplit('/', 1)[1]}.rels")
            dropped = {n for n in dropped if n in set(names)}
            report.parts_dropped = sorted(dropped)

            with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    name = info.filename
                    if name in dropped:
                        continue
                    data = zin.read(name)

                    if _BODY_PART_RE.match(name):
                        data = _strip_document_part(data, report)
                    elif name == "docProps/core.xml":
                        data = _strip_core_part(data, report)
                    elif name == "word/settings.xml":
                        data = _strip_settings_part(data, report)
                    elif name == "[Content_Types].xml":
                        data, ct_removed = _strip_content_types(data, dropped)
                        report.content_types_removed += ct_removed
                    elif name.endswith(".rels") and "_rels/" in name:
                        base = _rels_base_dir(name)
                        data, rel_removed = _strip_relationships(
                            data, base, dropped
                        )
                        report.relationships_removed += rel_removed

                    new_info = zipfile.ZipInfo(name, date_time=FIXED_ZIP_TIME)
                    new_info.compress_type = zipfile.ZIP_DEFLATED
                    new_info.external_attr = info.external_attr
                    zout.writestr(new_info, data)

    except zipfile.BadZipFile as exc:
        raise DocxStripError(f"{src} is not a valid .docx (zip) file: {exc}") from exc
    except ET.ParseError as exc:
        raise DocxStripError(f"malformed XML inside {src}: {exc}") from exc
    except OSError as exc:
        raise DocxStripError(f"file system error: {exc}") from exc

    return report
