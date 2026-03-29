"""Smart Markdown Editor — UI sub-package.

Exposes the top-level MainWindow so that external callers can do::

    from src.ui import MainWindow
"""

from src.ui.main_window import MainWindow  # noqa: F401

__all__ = ["MainWindow"]
