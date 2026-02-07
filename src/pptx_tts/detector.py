"""Detect PowerPoint slideshow mode and newly-opened .pptx files."""

from __future__ import annotations

import logging
import platform
import subprocess
import time
from pathlib import Path
from typing import Callable

import psutil

logger = logging.getLogger(__name__)

_SYSTEM = platform.system()

# ---------------------------------------------------------------------------
# Windows helpers
# ---------------------------------------------------------------------------

_WIN_PROCESS_NAMES = {"powerpnt.exe", "powerpnt"}

# Window class name used by PowerPoint's slideshow window (Windows)
_WIN_SLIDESHOW_CLASS = "screenClass"


def _win_is_slideshow_active() -> bool:
    """Check if a PowerPoint slideshow window is currently open (Windows)."""
    import win32gui  # available only on Windows

    hwnd = win32gui.FindWindow(_WIN_SLIDESHOW_CLASS, None)
    return hwnd != 0


def _win_find_powerpoint_pptx() -> str | None:
    """Scan running processes for PowerPoint with an open .pptx file (Windows)."""
    for proc in psutil.process_iter(["name"]):
        try:
            name = (proc.info["name"] or "").lower()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        if name not in _WIN_PROCESS_NAMES:
            continue
        pptx_path = _find_pptx_in_cmdline(proc)
        if pptx_path:
            return pptx_path
    return None


# ---------------------------------------------------------------------------
# macOS helpers (use AppleScript to talk to PowerPoint)
# ---------------------------------------------------------------------------

_MAC_PROCESS_NAMES = {"microsoft powerpoint"}


def _applescript(script: str) -> str | None:
    """Run an AppleScript snippet and return stripped stdout, or None on error."""
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=5,
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except (FileNotFoundError, subprocess.TimeoutExpired):
        pass
    return None


def _mac_is_slideshow_active() -> bool:
    """Check if PowerPoint is currently in slideshow mode (macOS)."""
    out = _applescript(
        'tell application "System Events" to '
        'return (name of processes) contains "Microsoft PowerPoint"'
    )
    if out != "true":
        return False

    out = _applescript(
        'tell application "Microsoft PowerPoint" to return running of slide show window of active presentation'
    )
    return out == "true"


def _mac_find_powerpoint_pptx() -> str | None:
    """Ask PowerPoint for the path of the active presentation (macOS)."""
    # First check PowerPoint is actually running via psutil (cheap)
    for proc in psutil.process_iter(["name"]):
        try:
            name = (proc.info["name"] or "").lower()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        if name in _MAC_PROCESS_NAMES:
            break
    else:
        return None

    # Ask PowerPoint for the file path via AppleScript
    out = _applescript(
        'tell application "Microsoft PowerPoint" to return full name of active presentation'
    )
    if out and out.lower().endswith(".pptx") and Path(out).is_file():
        return out

    # Fallback: scan command-line args (works for some versions)
    for proc in psutil.process_iter(["name"]):
        try:
            name = (proc.info["name"] or "").lower()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        if name in _MAC_PROCESS_NAMES:
            pptx_path = _find_pptx_in_cmdline(proc)
            if pptx_path:
                return pptx_path
    return None


# ---------------------------------------------------------------------------
# Shared / public API
# ---------------------------------------------------------------------------


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
    """Check if a PowerPoint slideshow is currently active."""
    if _SYSTEM == "Windows":
        return _win_is_slideshow_active()
    elif _SYSTEM == "Darwin":
        return _mac_is_slideshow_active()
    else:
        logger.warning("Slideshow detection not supported on %s", _SYSTEM)
        return False


def _find_powerpoint_pptx() -> str | None:
    """Scan for PowerPoint with an open .pptx file."""
    if _SYSTEM == "Windows":
        return _win_find_powerpoint_pptx()
    elif _SYSTEM == "Darwin":
        return _mac_find_powerpoint_pptx()
    else:
        logger.warning("PowerPoint detection not supported on %s", _SYSTEM)
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
