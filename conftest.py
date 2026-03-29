"""Root conftest.py — makes the project root importable as a package root."""
import sys
from pathlib import Path

# Ensure the repository root is on sys.path so that `import src` works from
# any test directory when running pytest from the project root.
sys.path.insert(0, str(Path(__file__).parent))
