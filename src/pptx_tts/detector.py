"""Detect PowerPoint application launches and newly-opened .pptx files."""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Callable

import psutil

logger = logging.getLogger(__name__)

# Process names that indicate PowerPoint (or LibreOffice Impress) is running
POWERPOINT_PROCESS_NAMES = {
    "powerpnt.exe",     # Microsoft PowerPoint (Windows)
    "powerpnt",
    "soffice.bin",      # LibreOffice (Linux/macOS)
    "soffice",
    "libreoffice",
    "impress",          # LibreOffice Impress directly
}


def _find_pptx_in_cmdline(proc: psutil.Process) -> str | None:
    """Try to extract a .pptx file path from a process's command line."""
    try:
        for arg in proc.cmdline():
            if arg.lower().endswith(".pptx") and Path(arg).is_file():
                return arg
    except (psutil.AccessDenied, psutil.NoSuchProcess):
        pass
    return None


class PowerPointDetector:
    """Polls running processes for PowerPoint with an open .pptx file."""

    def __init__(self, poll_interval: float = 3.0) -> None:
        self._poll_interval = poll_interval
        self._seen_files: set[str] = set()
        self._running = False

    def watch(self, on_file_opened: Callable[[str], None]) -> None:
        """Block and poll until stopped. Calls *on_file_opened* for each new .pptx.

        Args:
            on_file_opened: Callback receiving the path to a newly-detected .pptx file.
        """
        self._running = True
        logger.info(
            "Watching for PowerPoint processes (poll every %.1fs)...",
            self._poll_interval,
        )

        while self._running:
            for proc in psutil.process_iter(["name"]):
                try:
                    name = (proc.info["name"] or "").lower()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

                if name not in POWERPOINT_PROCESS_NAMES:
                    continue

                pptx_path = _find_pptx_in_cmdline(proc)
                if pptx_path and pptx_path not in self._seen_files:
                    self._seen_files.add(pptx_path)
                    logger.info("Detected new file: %s", pptx_path)
                    on_file_opened(pptx_path)

            time.sleep(self._poll_interval)

    def stop(self) -> None:
        """Signal the watch loop to exit."""
        self._running = False
