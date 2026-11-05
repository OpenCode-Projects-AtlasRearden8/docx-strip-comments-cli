"""Fixture builders.

Every fixture .docx is generated here with plain stdlib so the test suite
carries no binary files and no secrets: the fake author name and the fake
comment text used by these tests are obviously fictional.
"""

import zipfile
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
NS_DC = "http://purl.org/dc/elements/1.1/"
NS_DCTERMS = "http://purl.org/dc/terms/"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_WRE = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

FAKE_AUTHOR = "Dr. Sample Author"
FAKE_INITIALS = "DSA"
FAKE_COMMENT = "Sample comment: please rewrite this whole paragraph"
DELETED_TEXT = "This deleted sentence must not survive scrubbing"
INSERTED_TEXT = "This inserted sentence should be kept"
MOVE_SOURCE_TEXT = "Old moved source sentence"
MOVE_DEST_TEXT = "New moved destination sentence"
STATIC_TEXT = "Ordinary body paragraph that is neither revised nor commented."


def _xml(tag, attrs=None, children=None, text=None):
    elem = ET.Element(tag, attrs or {})
    if text is not None:
        elem.text = text
    for child in children or []:
        elem.append(child)
    return elem


def _document_xml() -> bytes:
    """Body with: static text, tracked deletion, tracked insertion,
    move tracking, comment anchors, and rsid/author fingerprints."""
    w = "{%s}" % W_NS

    static_run = _xml(
        f"{w}r",
        attrs={f"{w}rsidR": "00AA11BB", f"{w}rsidRPr": "00CC22DD"},
        children=[_xml(f"{w}t", text=STATIC_TEXT)],
    )

    deleted_run = _xml(f"{w}r", children=[_xml(f"{w}delText", text=DELETED_TEXT)])
    tracked_deletion = _xml(
        f"{w}del",
        attrs={
            f"{w}id": "10",
            f"{w}author": FAKE_AUTHOR,
            f"{w}date": "2026-03-14T09:26:00Z",
        },
        children=[deleted_run],
    )

    inserted_run = _xml(
        f"{w}r",
        attrs={f"{w}rsidR": "00EE33FF"},
        children=[_xml(f"{w}t", text=INSERTED_TEXT)],
    )
    tracked_insertion = _xml(
        f"{w}ins",
        attrs={
            f"{w}id": "11",
            f"{w}author": FAKE_AUTHOR,
            f"{w}authorInitials": FAKE_INITIALS,
            f"{w}date": "2026-03-14T10:00:00Z",
        },
        children=[inserted_run],
    )

    move_from_run = _xml(f"{w}r", children=[_xml(f"{w}t", text=MOVE_SOURCE_TEXT)])
    tracked_move_from = _xml(
        f"{w}moveFrom",
        attrs={f"{w}id": "12", f"{w}author": FAKE_AUTHOR, f"{w}date": "2026-03-14T11:00:00Z"},
        children=[move_from_run],
    )
    move_to_run = _xml(f"{w}r", children=[_xml(f"{w}t", text=MOVE_DEST_TEXT)])
    tracked_move_to = _xml(
        f"{w}moveTo",
        attrs={f"{w}id": "13", f"{w}author": FAKE_AUTHOR, f"{w}date": "2026-03-14T11:05:00Z"},
        children=[move_to_run],
    )

    paragraph = _xml(
        f"{w}p",
        attrs={f"{w}rsidP": "00EE44AA"},
        children=[
            _xml(f"{w}commentRangeStart", attrs={f"{w}id": "3"}),
            static_run,
            tracked_deletion,
            tracked_insertion,
            _xml(f"{w}moveFromRangeStart", attrs={f"{w}id": "4"}),
            tracked_move_from,
            tracked_move_to,
            _xml(f"{w}moveFromRangeEnd", attrs={f"{w}id": "4"}),
            _xml(
                f"{w}r",
                children=[
                    _xml(f"{w}commentReference", attrs={f"{w}id": "3"}),
                ],
            ),
            _xml(f"{w}commentRangeEnd", attrs={f"{w}id": "3"}),
        ],
    )
    # paragraph-mark revision + style change record on a second paragraph
    para_mark_delete = _xml(
        f"{w}p",
        children=[
            _xml(
                f"{w}pPr",
                children=[
                    _xml(
                        f"{w}rPr",
                        children=[
                            _xml(
                                f"{w}del",
                                attrs={f"{w}id": "15", f"{w}author": FAKE_AUTHOR},
                            )
                        ],
                    ),
                    _xml(
                        f"{w}pPrChange",
                        attrs={f"{w}id": "16", f"{w}author": FAKE_AUTHOR},
                        children=[_xml(f"{w}pPr", children=[_xml(f"{w}spacing")])],
                    ),
                ],
            ),
            _xml(f"{w}r", children=[_xml(f"{w}t", text="still here after merge")]),
        ],
    )

    body = _xml(
        f"{w}body",
        children=[
            paragraph,
            para_mark_delete,
            _xml(f"{w}sectPr", children=[_xml(f"{w}pgSz", attrs={f"{w}w": "12240"})]),
        ],
    )
    document = _xml(
        f"{w}document",
        children=[body],
    )
    # nested comment anchor inside a tracked insertion (double-wrapped case)
    nested_par = _xml(
        f"{w}p",
        children=[
            _xml(
                f"{w}ins",
                attrs={f"{w}id": "17", f"{w}author": FAKE_AUTHOR},
                children=[
                    _xml(f"{w}commentRangeStart", attrs={f"{w}id": "3"}),
                    _xml(
                        f"{w}r",
                        children=[
                            _xml(f"{w}t", text="nested revised text"),
                            _xml(f"{w}commentReference", attrs={f"{w}id": "3"}),
                        ],
                    ),
                ],
            ),
        ],
    )
    body.insert(1, nested_par)

    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(document, encoding="utf-8")
    )


def _content_types(extra_overrides=()) -> bytes:
    w = "{%s}" % NS_CT
    overrides = [
        (f"{w}Override", {"PartName": "/word/document.xml", "ContentType":
            "application/vnd.openxmlformats-officedocument.wordprocessingml"
            ".document.main+xml"}),
        (f"{w}Override", {"PartName": "/docProps/core.xml", "ContentType":
            "application/vnd.openxmlformats-package.core-properties+xml"}),
    ]
    for part, ct in extra_overrides:
        overrides.append((f"{w}Override", {"PartName": part, "ContentType": ct}))
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(
            _xml(f"{w}Types", children=[_xml(t, a) for t, a in overrides]),
            encoding="utf-8",
        )
    )


def _root_rels() -> bytes:
    r = "{%s}" % NS_REL
    rels = _xml(
        f"{r}Relationships",
        children=[
            _xml(
                f"{r}Relationship",
                attrs={
                    "Id": "rId1",
                    "Type": "http://schemas.openxmlformats.org/officeDocument"
                    "/2006/relationships/officeDocument",
                    "Target": "word/document.xml",
                },
            ),
            _xml(
                f"{r}Relationship",
                attrs={
                    "Id": "rId2",
                    "Type": "http://schemas.openxmlformats.org/package"
                    "/2006/relationships/metadata/core-properties",
                    "Target": "docProps/core.xml",
                },
            ),
            _xml(
                f"{r}Relationship",
                attrs={
                    "Id": "rId3",
                    "Type": "http://schemas.openxmlformats.org/officeDocument"
                    "/2006/relationships/customXmlProps",
                    "Target": "docProps/custom.xml",
                },
            ),
        ],
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(rels, encoding="utf-8")
    )


def _document_rels() -> bytes:
    r = "{%s}" % NS_REL
    rels = _xml(
        f"{r}Relationships",
        children=[
            _xml(
                f"{r}Relationship",
                attrs={
                    "Id": "rIdComments",
                    "Type": "http://schemas.openxmlformats.org/officeDocument"
                    "/2006/relationships/comments",
                    "Target": "comments.xml",
                },
            ),
            _xml(
                f"{r}Relationship",
                attrs={
                    "Id": "rIdPeople",
                    "Type": "http://schemas.microsoft.com/office/word"
                    "/2010/wordpeople",
                    "Target": "people.xml",
                },
            ),
            _xml(
                f"{r}Relationship",
                attrs={
                    "Id": "rIdCustom",
                    "Type": "http://schemas.openxmlformats.org/officeDocument"
                    "/2006/relationships/customXml",
                    "Target": "../customXml/item1.xml",
                },
            ),
        ],
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(rels, encoding="utf-8")
    )


def _core_xml() -> bytes:
    root = _xml(
        "{%s}coreProperties" % NS_CP,
        children=[
            _xml("{%s}title" % NS_DC, text="Sample Title"),
            _xml("{%s}creator" % NS_DC, text=FAKE_AUTHOR),
            _xml("{%s}lastModifiedBy" % NS_CP, text=FAKE_AUTHOR),
            _xml("{%s}revision" % NS_CP, text="42"),
            _xml(
                "{%s}created" % NS_DCTERMS,
                attrs={"{http://www.w3.org/XML/1998/namespace}lang": "en-US"},
                text="2026-03-14T09:00:00Z",
            ),
            _xml("{%s}modified" % NS_DCTERMS, text="2026-03-14T12:00:00Z"),
        ],
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(root, encoding="utf-8")
    )


def _comments_xml() -> bytes:
    w = "{%s}" % W_NS
    comment = _xml(
        f"{w}comment",
        attrs={
            f"{w}id": "3",
            f"{w}author": FAKE_AUTHOR,
            f"{w}initials": FAKE_INITIALS,
            f"{w}date": "2026-03-14T09:30:00Z",
        },
        children=[
            _xml(f"{w}p", children=[_xml(f"{w}r", children=[
                _xml(f"{w}t", text=FAKE_COMMENT),
                _xml(f"{w}annotationRef"),
            ])]),
        ],
    )
    root = _xml(f"{w}comments", children=[comment])
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(root, encoding="utf-8")
    )


def _people_xml() -> bytes:
    w = "{%s}" % W_NS
    root = _xml(
        f"{w}people",
        children=[
            _xml(
                f"{w}person",
                attrs={f"{w}author": FAKE_AUTHOR},
                children=[
                    _xml(
                        f"{w}presenceInfo",
                        attrs={f"{w}userId": f"jabber:{FAKE_INITIALS}@example.invalid"},
                    )
                ],
            )
        ],
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(root, encoding="utf-8")
    )


def _settings_xml() -> bytes:
    w = "{%s}" % W_NS
    root = _xml(
        f"{w}settings",
        children=[
            _xml(f"{w}trackChanges"),
            _xml(f"{w}proofState", attrs={f"{w}spelling": "clean"}),
            _xml(
                f"{w}rsids",
                children=[
                    _xml(f"{w}rsidRoot", attrs={f"{w}val": "00AA11BB"}),
                    _xml(f"{w}rsid", attrs={f"{w}val": "00CC22DD"}),
                ],
            ),
        ],
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(root, encoding="utf-8")
    )


def _cellar() -> dict:
    """All parts for the full 'messy' document."""
    return {
        "[Content_Types].xml": _content_types(
            extra_overrides=[
                (
                    "/word/comments.xml",
                    "application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.comments+xml",
                ),
                (
                    "/docProps/app.xml",
                    "application/vnd.openxmlformats-officedocument"
                    ".extended-properties+xml",
                ),
                (
                    "/docProps/custom.xml",
                    "application/vnd.openxmlformats-officedocument"
                    ".customProperties+xml",
                ),
                (
                    "/docProps/thumbnail.jpeg",
                    "image/jpeg",
                ),
            ]
        ),
        "_rels/.rels": _root_rels(),
        "word/document.xml": _document_xml(),
        "word/_rels/document.xml.rels": _document_rels(),
        "word/comments.xml": _comments_xml(),
        "word/people.xml": _people_xml(),
        "word/settings.xml": _settings_xml(),
        "word/styles.xml": _styles_xml(),
        "word/media/image1.png": b"\x89PNG\r\n\x1a\n" + b"\x00" * 8,
        "docProps/core.xml": _core_xml(),
        "docProps/app.xml": (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            b'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument'
            b'/2006/extended-properties"><Application>Microsoft Office Word'
            b"</Application><TotalTime>123</TotalTime>"
            b"<Company>Sample Company Inc</Company>"
            b"<Manager>Sample Manager</Manager><Words>500</Words></Properties>"
        ),
        "docProps/custom.xml": (
            b'<?xml version="1.0" encoding="UTF-8"?>\n'
            b"<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument"
            b'/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/'
            b'officeDocument/2006/docPropsVTypes"><property name="Editor" vt:lpwstr="'
            + FAKE_AUTHOR.encode()
            + b'"/></Properties>'
        ),
        "docProps/thumbnail.jpeg": b"\xff\xd8\xff" + b"\x00" * 32,
        "customXml/item1.xml": (
            b'<?xml version="1.0" encoding="UTF-8"?>\n'
            b"<data xmlns='http://example.invalid/ns'>internal notes</data>"
        ),
    }


def _plain_document_xml() -> bytes:
    """A document with no tracked changes and no comment anchors at all."""
    w = "{%s}" % W_NS
    root = _xml(
        f"{w}document",
        children=[
            _xml(
                f"{w}body",
                children=[
                    _xml(
                        f"{w}p",
                        children=[
                            _xml(f"{w}r", children=[_xml(f"{w}t", text=STATIC_TEXT)])
                        ],
                    )
                ],
            )
        ],
    )
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(root, encoding="utf-8")
    )


def _styles_xml() -> bytes:
    w = "{%s}" % W_NS
    root = _xml(f"{w}styles", children=[_xml(f"{w}docDefaults")])
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        + ET.tostring(root, encoding="utf-8")
    )


def write_docx(path, parts: dict) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=(2026, 3, 14, 9, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(info, data)
