#!/usr/bin/env python3
"""Main entry point for the Smart Markdown Editor application.

Example:
    $ python -m src.main
    $ python src/main.py
"""

import logging
import sys

from PySide6.QtWidgets import QApplication

from src.config import APP_NAME, APP_ORGANIZATION, APP_VERSION
from src.exporters.builtin import register_all as register_exporters
from src.ui.main_window import MainWindow


def setup_logging() -> None:
    """Configure application-level logging."""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def main() -> int:
    """Launch the Smart Markdown Editor.

    Returns:
        int: Exit code (0 on success).
    """
    setup_logging()
    logger = logging.getLogger(__name__)
    logger.info(f"Starting {APP_NAME} v{APP_VERSION}")

    register_exporters()

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName(APP_ORGANIZATION)

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
