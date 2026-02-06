"""Detect PowerPoint slideshow mode and newly-opened .pptx files."""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Callable

import psutil
import win32gui

logger = logging.getLogger(__name__)

# Process names that indicate PowerPoint is running
POWERPOINT_PROCESS_NAMES = {
    "powerpnt.exe",
    "powerpnt",
}

# Window class name used by PowerPoint's slideshow window
SLIDESHOW_WINDOW_CLASS = "screenClass"


def _find_pptx_in_cmdline(proc: psutil.Process) -> str | None:
    """Try to extract a .pptx file path from a process's command line."""
    try:
        for arg in proc.cmdline():
            if arg.lower().endswith(".pptx") and Path(arg).is_file():
                return arg
    except (psutil.AccessDenied, psutil.NoSuchProcess):
        pass
    return None


def _is_slideshow_active() -> bool:
    """Check if a PowerPoint slideshow window is currently open."""
    hwnd = win32gui.FindWindow(SLIDESHOW_WINDOW_CLASS, None)
    return hwnd != 0


def _find_powerpoint_pptx() -> str | None:
    """Scan running processes for PowerPoint with an open .pptx file."""
    for proc in psutil.process_iter(["name"]):
        try:
            name = (proc.info["name"] or "").lower()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

        if name not in POWERPOINT_PROCESS_NAMES:
            continue

        pptx_path = _find_pptx_in_cmdline(proc)
        if pptx_path:
            return pptx_path
    return None


class PowerPointDetector:
    """Polls for PowerPoint slideshow mode and triggers TTS accordingly."""

    def __init__(self, poll_interval: float = 3.0) -> None:
        self._poll_interval = poll_interval
        self._running = False

    def watch(
        self,
        on_slideshow_started: Callable[[str], None],
        on_slideshow_ended: Callable[[], None],
    ) -> None:
        """Block and poll until stopped.

        Waits for PowerPoint to have a .pptx open AND be in slideshow mode.
        Calls *on_slideshow_started* with the file path when slideshow begins.
        Calls *on_slideshow_ended* when the slideshow window disappears.
        """
        self._running = True
        logger.info(
            "Watching for PowerPoint slideshow mode (poll every %.1fs)...",
            self._poll_interval,
        )

        slideshow_active = False
        current_file: str | None = None

        while self._running:
            pptx_path = _find_powerpoint_pptx()
            is_presenting = _is_slideshow_active()

            if not slideshow_active and pptx_path and is_presenting:
                # Slideshow just started
                slideshow_active = True
                current_file = pptx_path
                logger.info("Slideshow started: %s", pptx_path)
                on_slideshow_started(pptx_path)

            elif slideshow_active and not is_presenting:
                # Slideshow just ended
                slideshow_active = False
                logger.info("Slideshow ended: %s", current_file)
                current_file = None
                on_slideshow_ended()

            time.sleep(self._poll_interval)

    def stop(self) -> None:
        """Signal the watch loop to exit."""
        self._running = False
