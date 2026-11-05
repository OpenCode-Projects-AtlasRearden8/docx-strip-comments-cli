"""End-to-end tests for the command line interface."""

import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

from docx_strip import __version__
from docx_strip.cli import main

from fixture_builder import DELETED_TEXT, FAKE_AUTHOR, INSERTED_TEXT


REPO_ROOT = Path(__file__).resolve().parents[1]


def run_cli(*args, cwd=None):
    proc = subprocess.run(
        [sys.executable, "-m", "docx_strip", *args],
        cwd=cwd or REPO_ROOT,
        capture_output=True,
        text=True,
        timeout=60,
    )
    return proc


def zip_names(path) -> set:
    with zipfile.ZipFile(path) as zf:
        return set(zf.namelist())


class TestCliSuccess:
    def test_default_output_and_summary(self, messy_docx):
        proc = run_cli(str(messy_docx))
        assert proc.returncode == 0, proc.stderr
        clean = messy_docx.with_name("messy_clean.docx")
        assert clean.is_file()
        assert "Clean copy:" in proc.stdout
        assert "word/comments.xml" in proc.stdout  # listed as dropped

    def test_explicit_output_flag(self, messy_docx, tmp_path):
        out = tmp_path / "share-me.docx"
        proc = run_cli(str(messy_docx), "-o", str(out))
        assert proc.returncode == 0, proc.stderr
        assert out.is_file()

    def test_quiet_mode(self, messy_docx):
        proc = run_cli(str(messy_docx), "--quiet")
        assert proc.returncode == 0
        assert proc.stdout == ""

    def test_original_untouched_via_cli(self, messy_docx, original_bytes):
        run_cli(str(messy_docx))
        assert messy_docx.read_bytes() == original_bytes

    def test_nothing_readable_leaks_via_cli(self, messy_docx):
        run_cli(str(messy_docx))
        clean = messy_docx.with_name("messy_clean.docx")
        assert INSERTED_TEXT.encode() in zipfile.ZipFile(clean).read(
            "word/document.xml"
        )
        with zipfile.ZipFile(clean) as zf:
            for name in zf.namelist():
                data = zf.read(name)
                assert DELETED_TEXT.encode() not in data
                assert FAKE_AUTHOR.encode() not in data

    def test_version_flag(self):
        proc = run_cli("--version")
        assert proc.returncode == 0
        assert __version__ in proc.stdout

    def test_help_describes_goal(self):
        proc = run_cli("--help")
        assert proc.returncode == 0
        assert "tracked changes" in proc.stdout
        assert "never modified" in proc.stdout


class TestCliErrors:
    def test_missing_file_exit_code(self, tmp_path):
        proc = run_cli(str(tmp_path / "ghost.docx"))
        assert proc.returncode == 1
        assert "error:" in proc.stderr

    def test_non_zip_exit_code(self, tmp_path):
        fake = tmp_path / "fake.docx"
        fake.write_bytes(b"plain text, definitely not a zip")
        proc = run_cli(str(fake))
        assert proc.returncode == 1
        assert "not a valid .docx" in proc.stderr

    def test_zip_without_document_exit_code(self, tmp_path):
        import zipfile as zfmod

        fake = tmp_path / "somezip.docx"
        with zfmod.ZipFile(fake, "w") as zf:
            zf.writestr("readme.txt", "just a zip")
        proc = run_cli(str(fake))
        assert proc.returncode == 1
        assert "not a Word" in proc.stderr

    def test_bad_output_dir_does_not_crash_traceback(self, messy_docx, tmp_path):
        proc = run_cli(str(messy_docx), "-o", str(tmp_path / "no" / "such" / "out.docx"))
        assert proc.returncode == 1
        assert "Traceback" not in proc.stderr
