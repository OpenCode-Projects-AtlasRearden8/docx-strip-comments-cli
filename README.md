# docx-strip-comments-cli

Copy a Word document and strip everything you wouldn't want a stranger to read:
tracked changes, inline comments, and your personal metadata — so the *clean
copy* is safe to share. The original file is never touched.

This is the polite share version, not a forensic tool. It doesn't analyze or
recover anything; it just makes a tidy copy you can hand to an editor, a
classmate, or a client without leaking "who wrote what, when".

## Why this exists

Turning on Track Changes in Word records a revision history inside the file,
and every comment you or a reviewer adds carries the author's name, initials,
and a timestamp. Share the original file as-is and all of that ships with it.

This tool accepts every tracked change, removes every comment, and scrubs the
author/revision metadata — offline, with nothing leaving your computer. It is
deliberately different from tools like `email-headers-inspector-cli` (which
inspects and reports on email headers) or generic redaction tools (which aim
to make content unreadable). Here the goal is simple: keep the finished text,
drop the fingerprints.

## Highlights

- **Offline by design.** Pure Python standard library — `zipfile` and
  `xml.etree`. No network access, no API keys, no accounts.
- **Original always safe.** The source file is opened read-only; the cleaned
  document is written as a brand-new file.
- **Report you can read.** Every run prints exactly what was removed, in
  plain English.
- **Calm defaults.** Output lands next to the original as `<name>_clean.docx`
  unless you say otherwise.

## What gets removed

| Category | Detail |
| --- | --- |
| Tracked edits | Deletions removed; insertions kept (every change is accepted); move tracking resolved; format-change records (`rPrChange`, `pPrChange`, `sectPrChange`, table equivalents) dropped |
| Comments | `word/comments*.xml` and `word/people.xml` parts deleted, plus all comment anchors/ranges in the body |
| Author traces | `dc:creator`, `lastModifiedBy`, `cp:revision`, created/modified dates; `w:author`, `w:authorInitials`, `w:date` attributes |
| Revision history | `w:rsid*` session fingerprints (document + `w:rsids` in settings), `w:trackChanges` flag |
| Extras | `docProps/app.xml` (company, manager, total editing time), `docProps/custom.xml`, document thumbnail, `customXml` data parts — removed outright |
| ZIP hygiene | Every entry timestamp normalized so the copy leaks no edit-session timing |

Kept: your text and formatting, styles, images, headers/footers content,
footnotes/endnotes content, and any title/subject you set in core properties.

## Requirements

Python 3.9 or newer. Nothing to install.

## Usage

```bash
# as a module (works straight from this folder)
python -m docx_strip report.docx

# with an explicit destination
python -m docx_strip report.docx -o "for-review.docx"

# custom default suffix, silent mode, version
python -m docx_strip report.docx --suffix "_share-ready"
python -m docx_strip report.docx --quiet
python -m docx_strip --version
```

Example output:

```
Source: report.docx
Clean copy: report_clean.docx
Parts dropped: docProps/app.xml, docProps/custom.xml, docProps/thumbnail.jpeg, word/comments.xml, word/people.xml
Tracked insertions kept: 3
Tracked deletions removed: 3
Change records removed: 1
Comment marks removed: 5
Revision range markers removed: 2
Core metadata fields removed: 5
```

You can also install it as a command if you prefer:

```bash
pip install .
docx-strip report.docx
```

## How it works

A `.docx` file is a ZIP archive of XML parts. The tool walks the archive,
rewrites the body parts with every tracked change resolved (accepted) and every
comment anchor removed, drops the comment/metadata parts entirely from
`[Content_Types].xml` and the relationship files, and repacks everything else
byte-for-byte into a fresh copy with normalized timestamps.

## Verifying it

```bash
python -m pytest
```

The test suite builds fixture `.docx` files (with fake author names and fake
comments — no real data anywhere) and then opens every surviving archive
member to prove nothing readable leaks through: no deleted sentences, no
comment text, no author strings, no rsid values, no stale timestamps.

## Known limits

- Accepting a *paragraph-mark* deletion keeps the paragraph mark (the text of
  both paragraphs survives). Word's own "Accept All" would merge those
  paragraphs; the reading content is identical either way.
- Content controls (`w:customXml` bound controls) are stripped of revision
  markup; their underlying `customXml` data parts are removed entirely.
- Older `.doc` (pre-2007) files are intentionally out of scope — this is a
  plain-XML tool. Convert first.
- Not a renderer: if a document displays oddly in Word afterwards, send us a
  bug report with a description (never the document itself).
- The file-open password hash stored in `w:writeProtection` (protection to
  *edit* a file) is removed. Passwords to *open* the document, and an
  encrypted body, keep the file as-is: those are content, not metadata.

## Status

See [STATUS.md](STATUS.md) for the current build status and verification
results.

## License

MIT — see [LICENSE](LICENSE).

Copyright (c) 2026 Asdale Dynamics, LLC. Built by Dylan Asdale for the
Massive AI Manual project: plain-English AI tools that respect your privacy,
run locally, and never ask for your keys.
