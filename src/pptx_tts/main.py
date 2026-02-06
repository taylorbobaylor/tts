"""CLI entry point for pptx-tts-reader."""

from __future__ import annotations

import argparse
import logging
import signal
import sys

from pptx_tts.detector import PowerPointDetector
from pptx_tts.playback import PlaybackController
from pptx_tts.synthesizer import VoiceConfig


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx-reader",
        description="Read PowerPoint presentations aloud using text-to-speech.",
    )
    sub = parser.add_subparsers(dest="command")

    # --- read: speak a single .pptx file ---
    read_p = sub.add_parser("read", help="Read a .pptx file aloud")
    read_p.add_argument("file", help="Path to a .pptx file")
    read_p.add_argument(
        "--delay",
        type=float,
        default=1.5,
        help="Seconds to pause between slides (default: 1.5)",
    )
    read_p.add_argument(
        "--rate",
        type=int,
        default=175,
        help="Speech rate in words per minute (default: 175)",
    )
    read_p.add_argument(
        "--no-joke",
        action="store_true",
        help="Skip the ending joke",
    )

    # --- watch: background mode — detect PowerPoint and read automatically ---
    watch_p = sub.add_parser(
        "watch", help="Run in the background and auto-read when PowerPoint enters slideshow mode"
    )
    watch_p.add_argument(
        "--poll",
        type=float,
        default=3.0,
        help="Seconds between process polls (default: 3.0)",
    )
    watch_p.add_argument("--delay", type=float, default=1.5)
    watch_p.add_argument("--rate", type=int, default=175)
    watch_p.add_argument("--no-joke", action="store_true")

    # --- voices: list available TTS voices ---
    sub.add_parser("voices", help="List available TTS voices on this system")

    return parser


def _cmd_read(args: argparse.Namespace) -> None:
    voice = VoiceConfig(rate=args.rate)
    ctrl = PlaybackController(
        voice_config=voice, slide_delay=args.delay, tell_joke=not args.no_joke
    )
    ctrl.play_presentation(args.file)


def _cmd_watch(args: argparse.Namespace) -> None:
    voice = VoiceConfig(rate=args.rate)
    ctrl = PlaybackController(
        voice_config=voice, slide_delay=args.delay, tell_joke=not args.no_joke
    )
    detector = PowerPointDetector(poll_interval=args.poll)

    def _on_sigint(sig: int, frame: object) -> None:
        print("\nStopping...")
        ctrl.stop()
        detector.stop()

    signal.signal(signal.SIGINT, _on_sigint)
    signal.signal(signal.SIGTERM, _on_sigint)

    detector.watch(
        on_slideshow_started=ctrl.play_presentation,
        on_slideshow_ended=ctrl.finish_with_joke,
    )


def _cmd_voices(_args: argparse.Namespace) -> None:
    from pptx_tts.synthesizer import Synthesizer

    synth = Synthesizer()
    voices = synth.list_voices()
    if not voices:
        print("No voices found.")
        return
    for v in voices:
        print(f"  {v['name']}  ({v['id']})")


def main() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )

    parser = _build_parser()
    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    commands = {
        "read": _cmd_read,
        "watch": _cmd_watch,
        "voices": _cmd_voices,
    }
    commands[args.command](args)


if __name__ == "__main__":
    main()
