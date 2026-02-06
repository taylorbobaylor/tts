"""Tests for the PowerPoint detector module."""

from unittest.mock import MagicMock, patch

from pptx_tts.detector import PowerPointDetector, _find_pptx_in_cmdline


class TestFindPptxInCmdline:
    def test_finds_pptx_arg(self, tmp_path):
        pptx = tmp_path / "deck.pptx"
        pptx.write_bytes(b"fake")
        proc = MagicMock()
        proc.cmdline.return_value = ["powerpnt.exe", str(pptx)]
        assert _find_pptx_in_cmdline(proc) == str(pptx)

    def test_returns_none_when_no_pptx(self):
        proc = MagicMock()
        proc.cmdline.return_value = ["powerpnt.exe"]
        assert _find_pptx_in_cmdline(proc) is None

    def test_returns_none_on_access_denied(self):
        import psutil

        proc = MagicMock()
        proc.cmdline.side_effect = psutil.AccessDenied(pid=1)
        assert _find_pptx_in_cmdline(proc) is None


class TestPowerPointDetector:
    def test_stop_ends_watch_loop(self):
        detector = PowerPointDetector(poll_interval=0.01)
        callback = MagicMock()

        # Stop immediately on first poll cycle
        original_sleep = None
        import time

        def fake_sleep(secs):
            detector.stop()

        with patch("pptx_tts.detector.time.sleep", side_effect=fake_sleep):
            detector.watch(callback)

        # Loop exited without error
        assert not detector._running
