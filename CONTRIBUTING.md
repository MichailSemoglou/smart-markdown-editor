# Contributing to Smart Markdown Editor

Thank you for your interest in contributing! This document explains the process
and conventions for contributing code, documentation, or bug reports.

## Table of Contents

1. [Getting Started](#getting-started)
2. [Development Setup](#development-setup)
3. [Code Style](#code-style)
4. [Testing](#testing)
5. [Submitting Changes](#submitting-changes)
6. [Reporting Issues](#reporting-issues)

---

## Getting Started

1. **Fork** the repository on GitHub.
2. **Clone** your fork locally:
   ```bash
   git clone https://github.com/<your-username>/smart-markdown-editor.git
   cd smart-markdown-editor
   ```
3. Create a **feature branch** from `main`:
   ```bash
   git checkout -b feat/my-feature
   ```

---

## Development Setup

Python 3.9 or later is required.

```bash
# Create and activate a virtual environment
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# Install all dependencies (including development extras)
pip install -r requirements.txt
pip install pytest pytest-cov ruff mypy
```

On Linux you also need the system libraries used by WeasyPrint:

```bash
sudo apt-get install libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz0b \
                     libcairo2 libgdk-pixbuf2.0-0
```

Verify your setup:

```bash
python markdown_editor.py   # launches the editor (requires a display)
pytest tests/ test_exports.py -v
```

---

## Code Style

This project uses **Ruff** for linting and formatting.

```bash
ruff check .          # lint
ruff format .         # auto-format
```

Key conventions:

- Follow [PEP 8](https://peps.python.org/pep-0008/).
- Use type annotations for all new public functions and methods.
- Keep lines to 100 characters or fewer.
- Do not introduce new module-level side effects.
- Static/class methods that do not need `self` should be decorated as such.

---

## Testing

All changes must be covered by automated tests.

```bash
pytest tests/ test_exports.py -v --cov=src --cov-report=term-missing
```

- Unit tests live in `tests/`.
- Export integration tests live in `test_exports.py`.
- Use `pytest.mark.skipif` (not `unittest.skip`) to skip tests whose
  dependencies are unavailable.
- Do **not** use `print()` statements in test files; use `assert`.

---

## Submitting Changes

1. Ensure all tests pass and `ruff check .` reports no errors.
2. Update `CHANGELOG.md` under the `[Unreleased]` section.
3. Push your branch and open a **Pull Request** against `main`.
4. Fill in the PR template describing what changed and why.
5. Address any review comments before the PR is merged.

---

## Reporting Issues

Please open a GitHub issue and include:

- A clear title and description.
- Steps to reproduce the problem.
- Expected vs. actual behaviour.
- Python version and operating system.
- Any relevant error messages or stack traces (wrapped in a code block).
