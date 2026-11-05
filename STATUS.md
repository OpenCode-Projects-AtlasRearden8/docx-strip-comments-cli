# STATUS — docx-strip-comments-cli

Last updated: 2026-11-05

## What was built

An offline CLI that copies a `.docx` and produces a clean share copy:

- **Tracked changes resolved** as if every change were accepted — insertions
  kept, deletions removed, move tracking resolved, and every change record
  (`rPrChange`, `pPrChange`, `sectPrChange`, table equivalents) plus all
  comment anchors/ranges stripped from the body XML (document, headers,
  footers, footnotes, endnotes, glossary).
- **Comment parts deleted outright**: `word/comments*.xml`,
  `word/commentInformation.xml`, `word/people.xml`, with their relationships
  and `[Content_Types].xml` overrides cleaned up.
- **Personal metadata scrubbed**: `dc:creator`, `lastModifiedBy`,
  `cp:revision`, created/modified dates; `w:author` / `w:authorInitials` /
  `w:date` / `w:rsid*` attributes; `w:rsids` + `w:trackChanges` in settings;
  `docProps/app.xml` (company, manager, editing time), `docProps/custom.xml`,
  thumbnails, and `customXml/` data parts dropped; every ZIP entry timestamp
  normalized to 1980-01-01 so the copy leaks no edit-session timing.
- **Original never modified** — source is opened read-only; the clean copy is
  written as a new file (default `<name>_clean.docx` next to the source).

Distinct positioning (kept intentionally): this is the polite-share tool, not
a forensic analyzer and not a redaction tool; `email-headers-inspector-cli`
inspects/reports — this one quietly fixes a Word file.

## Repository layout

```
docx_strip/          package (scrub.py core, cli.py argparse front end)
tests/               pytest suite; fixtures are generated in fixture_builder.py
                     (fake author names/comment text only — no real data)
README.md, LICENSE, pyproject.toml, STATUS.md
```

## How to run

```bash
# straight from the repo, no install
python -m docx_strip report.docx            # -> report_clean.docx
python -m docx_strip report.docx -o out.docx
python -m docx_strip --help

# or install as a command
pip install .          # (offline: pip install --no-build-isolation .)
docx-strip report.docx

# tests
python -m pytest
```

## Verification results (this session)

- `python -m pytest tests` → **39 passed**, 0 failures.
- Fixture leak-proofs: every surviving archive member is decompressed and
  scanned — deleted sentences, comment text, author name, initials, editor
  custom property, rsid values, comment part names, app/custom/thumbnail/
  customXml parts: **all absent** from the output. Inserted and moved-into
  text **present**; original file bytes **byte-identical** after runs.
- Installed package check: `pip install .` into a target dir succeeded and
  `docx_strip.cli.main` plus the `--version` flag ran correctly from the
  installed copy.
- Manual CLI run on the messy fixture produced a complete plain-English
  report (6 parts dropped, 3 insertions kept, 3 deletions removed, 5 comment
  marks, 5 core fields, ZIP timestamps normalized) and exit code 0.
- Error paths verified: missing file, non-zip, zip-without-document all exit
  1 with a clear one-line message and never write output.
- Python 3.13.5, pytest 9.1.1, stdlib only — no third-party runtime deps.

## Known gaps

- Accepting a *paragraph-mark* deletion keeps the paragraph mark (both
  paragraphs' text survives, no merge). Word's Accept-All would merge them;
  displayed text is the same.
- `w:documentProtection` (editing restrictions) is left in place; only the
  `w:writeProtection` open-for-edit password hash is dropped.
- Very old constructs (deprecated `word/commentsEx.xml` alias, custom XML
  bound controls' markup in newer schemas) are handled by pattern rather
  than per-spec; real-world Word output is covered by the patterns used.
- Pre-2007 binary `.doc` is out of scope by design.
- No shell-completion / packaging beyond a console script entry point.
