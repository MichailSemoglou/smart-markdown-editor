"""Tests for the headless CLI (src/__main__.py)."""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

_PYTHON = sys.executable
_CLI = [_PYTHON, "-m", "src"]

# Root of the repo (two levels up from this file)
_REPO_ROOT = Path(__file__).parent.parent
_SAMPLE = _REPO_ROOT / "test_sample.md"


def _run(*args: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [*_CLI, *args],
        capture_output=True,
        text=True,
        cwd=str(_REPO_ROOT),
    )


class TestCLIVersion:
    def test_version_exits_zero(self):
        result = _run("--version")
        assert result.returncode == 0

    def test_version_contains_name_and_number(self):
        result = _run("--version")
        assert "Smart Markdown Editor" in result.stdout
        assert "1.1.0" in result.stdout


class TestCLIStats:
    def test_stats_exits_zero(self):
        result = _run("--stats", str(_SAMPLE))
        assert result.returncode == 0

    def test_stats_produces_valid_json(self):
        result = _run("--stats", str(_SAMPLE))
        data = json.loads(result.stdout)
        assert isinstance(data, dict)

    def test_stats_contains_required_keys(self):
        result = _run("--stats", str(_SAMPLE))
        data = json.loads(result.stdout)
        required = {
            "word_count",
            "char_count",
            "line_count",
            "reading_time",
            "headings",
            "links",
            "images",
            "code_blocks",
            "lists",
            "blockquotes",
            "tables",
            "readability_score",
            "structure_quality",
        }
        assert required.issubset(data.keys())

    def test_stats_omits_broken_links_key(self):
        result = _run("--stats", str(_SAMPLE))
        data = json.loads(result.stdout)
        assert "broken_links" not in data

    def test_stats_word_count_positive(self):
        result = _run("--stats", str(_SAMPLE))
        data = json.loads(result.stdout)
        assert data["word_count"] > 0

    def test_stats_headings_is_dict(self):
        result = _run("--stats", str(_SAMPLE))
        data = json.loads(result.stdout)
        assert isinstance(data["headings"], dict)
        assert set(data["headings"].keys()) == {"h1", "h2", "h3", "h4", "h5", "h6"}


class TestCLILint:
    def test_lint_produces_valid_json(self):
        result = _run("--lint", str(_SAMPLE))
        data = json.loads(result.stdout)
        assert isinstance(data, dict)

    def test_lint_json_schema(self):
        result = _run("--lint", str(_SAMPLE))
        data = json.loads(result.stdout)
        assert "file" in data
        assert "clean" in data
        assert "issue_count" in data
        assert "issues" in data
        assert "readability_score" in data
        assert "readability_level" in data
        assert "structure_quality" in data

    def test_lint_readability_level_valid(self):
        result = _run("--lint", str(_SAMPLE))
        data = json.loads(result.stdout)
        assert data["readability_level"] in ("green", "orange", "red")

    def test_lint_clean_document_exits_zero(self, tmp_path):
        md = tmp_path / "clean.md"
        md.write_text(
            "# Document Title\n\n"
            "This is a well-formed document with simple prose.\n"
        )
        result = _run("--lint", str(md))
        data = json.loads(result.stdout)
        assert data["clean"] is True
        assert data["issue_count"] == 0
        assert result.returncode == 0

    def test_lint_document_with_issues_exits_one(self, tmp_path):
        md = tmp_path / "broken.md"
        md.write_text("# Title\n\nA link with no URL: [click here]()\n")
        result = _run("--lint", str(md))
        data = json.loads(result.stdout)
        assert data["clean"] is False
        assert data["issue_count"] > 0
        assert len(data["issues"]) > 0
        assert result.returncode == 1

    def test_lint_issues_is_list(self):
        result = _run("--lint", str(_SAMPLE))
        data = json.loads(result.stdout)
        assert isinstance(data["issues"], list)


class TestCLIErrors:
    def test_missing_file_exits_two(self):
        result = _run("--lint", "/nonexistent/path/does_not_exist.md")
        assert result.returncode == 2

    def test_missing_file_error_on_stderr(self):
        result = _run("--lint", "/nonexistent/path/does_not_exist.md")
        err = json.loads(result.stderr)
        assert "error" in err

    def test_no_args_exits_nonzero(self):
        result = _run()
        assert result.returncode != 0

    @pytest.mark.parametrize("flag", ["--stats", "--lint"])
    def test_only_flag_no_file_exits_nonzero(self, flag):
        result = _run(flag)
        assert result.returncode != 0
