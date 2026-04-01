"""Headless CLI entry-point for smart-markdown-editor.

Usage::

    python -m src --stats document.md   # JSON statistics
    python -m src --lint  document.md   # JSON lint report; exits 1 if issues found
    python -m src --version             # print version and exit

Or via installed scripts::

    smart-md --stats document.md
    smart-md --lint  document.md
    smart-md-lint document.md
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

_VERSION = "1.1.0"

# Sentinel returned by MarkdownAnalyzer when no real issues are found.
_CLEAN_SENTINEL = ["No issues detected"]


def _read_file(path: str) -> str:
    p = Path(path)
    if not p.exists():
        print(json.dumps({"error": f"File not found: {path}"}), file=sys.stderr)
        sys.exit(2)
    if not p.is_file():
        print(json.dumps({"error": f"Not a file: {path}"}), file=sys.stderr)
        sys.exit(2)
    return p.read_text(encoding="utf-8")


def cmd_stats(path: str) -> None:
    from src.core.analyzer import MarkdownAnalyzer

    text = _read_file(path)
    metrics = MarkdownAnalyzer(text).analyze()
    stats = {k: v for k, v in metrics.items() if k != "broken_links"}
    print(json.dumps(stats, indent=2))


def cmd_lint(path: str) -> None:
    from src.core.analyzer import MarkdownAnalyzer, get_readability_color

    text = _read_file(path)
    metrics = MarkdownAnalyzer(text).analyze()
    raw_issues = metrics["broken_links"]
    has_issues = raw_issues != _CLEAN_SENTINEL
    output = {
        "file": path,
        "clean": not has_issues,
        "issue_count": len(raw_issues) if has_issues else 0,
        "issues": raw_issues if has_issues else [],
        "readability_score": metrics["readability_score"],
        "readability_level": get_readability_color(metrics["readability_score"]),
        "structure_quality": metrics["structure_quality"],
    }
    print(json.dumps(output, indent=2))
    sys.exit(1 if has_issues else 0)


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="smart-md",
        description="Smart Markdown Editor — headless document analysis CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "examples:\n"
            "  smart-md --stats  document.md\n"
            "  smart-md --lint   document.md\n"
            "  smart-md --version\n"
        ),
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"Smart Markdown Editor {_VERSION}",
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument(
        "--stats",
        metavar="FILE",
        help="Print document statistics as JSON.",
    )
    group.add_argument(
        "--lint",
        metavar="FILE",
        help="Lint document and print issues as JSON. Exits 1 if issues are found.",
    )

    args = parser.parse_args()
    if args.stats:
        cmd_stats(args.stats)
    else:
        cmd_lint(args.lint)


def lint_entry() -> None:
    """Entry point for the ``smart-md-lint FILE`` convenience script."""
    parser = argparse.ArgumentParser(
        prog="smart-md-lint",
        description="Lint a Markdown file and report issues as JSON.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"Smart Markdown Editor {_VERSION}",
    )
    parser.add_argument("file", metavar="FILE", help="Path to the Markdown file.")
    args = parser.parse_args()
    cmd_lint(args.file)


if __name__ == "__main__":
    main()
