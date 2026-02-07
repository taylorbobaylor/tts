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
        started_cb = MagicMock()
        ended_cb = MagicMock()

        # Stop immediately on first poll cycle
        def fake_sleep(secs):
            detector.stop()

        with patch("pptx_tts.detector.time.sleep", side_effect=fake_sleep), \
             patch("pptx_tts.detector._find_powerpoint_pptx", return_value=None), \
             patch("pptx_tts.detector._is_slideshow_active", return_value=False):
            detector.watch(started_cb, ended_cb)

        assert not detector._running

    def test_callback_fires_only_during_slideshow(self):
        """Callback should NOT fire when PowerPoint has a file but no slideshow."""
        detector = PowerPointDetector(poll_interval=0.01)
        started_cb = MagicMock()
        ended_cb = MagicMock()
        call_count = 0

        def fake_sleep(secs):
            nonlocal call_count
            call_count += 1
            if call_count >= 2:
                detector.stop()

        with patch("pptx_tts.detector.time.sleep", side_effect=fake_sleep), \
             patch("pptx_tts.detector._find_powerpoint_pptx", return_value="/tmp/test.pptx"), \
             patch("pptx_tts.detector._is_slideshow_active", return_value=False):
            detector.watch(started_cb, ended_cb)

        started_cb.assert_not_called()
        ended_cb.assert_not_called()

    def test_callback_fires_when_slideshow_starts(self):
        """Callback fires when both a .pptx file and slideshow window are detected."""
        detector = PowerPointDetector(poll_interval=0.01)
        started_cb = MagicMock()
        ended_cb = MagicMock()

        def fake_sleep(secs):
            detector.stop()

        with patch("pptx_tts.detector.time.sleep", side_effect=fake_sleep), \
             patch("pptx_tts.detector._find_powerpoint_pptx", return_value="/tmp/deck.pptx"), \
             patch("pptx_tts.detector._is_slideshow_active", return_value=True):
            detector.watch(started_cb, ended_cb)

        started_cb.assert_called_once_with("/tmp/deck.pptx")
        ended_cb.assert_not_called()

    def test_ended_callback_fires_when_slideshow_ends(self):
        """on_slideshow_ended fires when the slideshow window disappears."""
        detector = PowerPointDetector(poll_interval=0.01)
        started_cb = MagicMock()
        ended_cb = MagicMock()
        poll_count = 0

        slideshow_states = [True, False]  # active, then inactive

        def fake_sleep(secs):
            nonlocal poll_count
            poll_count += 1
            if poll_count >= 2:
                detector.stop()

        def slideshow_active():
            if poll_count < len(slideshow_states):
                return slideshow_states[poll_count]
            return False

        with patch("pptx_tts.detector.time.sleep", side_effect=fake_sleep), \
             patch("pptx_tts.detector._find_powerpoint_pptx", return_value="/tmp/deck.pptx"), \
             patch("pptx_tts.detector._is_slideshow_active", side_effect=slideshow_active):
            detector.watch(started_cb, ended_cb)

        started_cb.assert_called_once_with("/tmp/deck.pptx")
        ended_cb.assert_called_once()


class TestMacOSHelpers:
    """Test macOS-specific detection logic with mocked subprocess calls."""

    @patch("pptx_tts.detector._applescript")
    def test_mac_slideshow_active(self, mock_as):
        from pptx_tts.detector import _mac_is_slideshow_active

        # PowerPoint running + slideshow running
        mock_as.side_effect = ["true", "true"]
        assert _mac_is_slideshow_active() is True

    @patch("pptx_tts.detector._applescript")
    def test_mac_slideshow_not_active(self, mock_as):
        from pptx_tts.detector import _mac_is_slideshow_active

        # PowerPoint running, slideshow NOT running
        mock_as.side_effect = ["true", "false"]
        assert _mac_is_slideshow_active() is False

    @patch("pptx_tts.detector._applescript")
    def test_mac_powerpoint_not_running(self, mock_as):
        from pptx_tts.detector import _mac_is_slideshow_active

        mock_as.return_value = "false"
        assert _mac_is_slideshow_active() is False

    @patch("pptx_tts.detector._applescript")
    @patch("pptx_tts.detector.psutil.process_iter")
    def test_mac_find_pptx_via_applescript(self, mock_procs, mock_as, tmp_path):
        from pptx_tts.detector import _mac_find_powerpoint_pptx

        pptx = tmp_path / "slides.pptx"
        pptx.write_bytes(b"fake")

        # Simulate PowerPoint process in psutil
        proc = MagicMock()
        proc.info = {"name": "Microsoft PowerPoint"}
        mock_procs.return_value = [proc]

        mock_as.return_value = str(pptx)
        assert _mac_find_powerpoint_pptx() == str(pptx)

    @patch("pptx_tts.detector.psutil.process_iter")
    def test_mac_find_pptx_not_running(self, mock_procs):
        from pptx_tts.detector import _mac_find_powerpoint_pptx

        # No PowerPoint process
        proc = MagicMock()
        proc.info = {"name": "Safari"}
        mock_procs.return_value = [proc]

        assert _mac_find_powerpoint_pptx() is None
