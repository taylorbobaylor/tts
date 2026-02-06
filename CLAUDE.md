# CLAUDE.md - PowerPoint Text-to-Speech Auto-Reader

## Project Overview

A Python tool that reads PowerPoint presentations aloud using offline AI text-to-speech. It can read a single `.pptx` file on demand, or run as a background watcher that auto-detects when PowerPoint opens a file.

## Repository Structure

```
tts/
├── src/pptx_tts/          # Main package
│   ├── __init__.py
│   ├── main.py             # CLI entry point (argparse)
│   ├── extractor.py        # Slide text extraction via python-pptx
│   ├── synthesizer.py      # TTS wrapper around pyttsx3
│   ├── playback.py         # Orchestrates reading a full presentation
│   └── detector.py         # Polls for PowerPoint/LibreOffice processes
├── tests/
│   ├── test_extractor.py   # Extraction + SlideContent tests
│   ├── test_playback.py    # Playback controller tests (mocked TTS)
│   └── test_detector.py    # Detector + process scanning tests
├── pyproject.toml           # Build config, deps, CLI entry point
├── .gitignore
├── README.MD
└── CLAUDE.md
```

## Quick Reference

```bash
# Install (editable, with dev deps)
pip install -e ".[dev]"

# Read a presentation aloud
pptx-reader read slides.pptx
pptx-reader read slides.pptx --rate 200 --delay 2.0 --no-joke

# Watch for PowerPoint and auto-read
pptx-reader watch

# List available TTS voices
pptx-reader voices

# Run tests
pytest
```

## Architecture

| Module | Responsibility |
|--------|---------------|
| `extractor.py` | Parses `.pptx` files with `python-pptx`, returns `SlideContent` dataclasses |
| `synthesizer.py` | Wraps `pyttsx3` engine — `speak()`, `stop()`, voice config |
| `playback.py` | Loops over slides, calls synthesizer, handles inter-slide delay and ending joke |
| `detector.py` | Polls `psutil.process_iter` for PowerPoint/LibreOffice, extracts `.pptx` from cmdline |
| `main.py` | CLI with subcommands: `read`, `watch`, `voices` |

## Dependencies

- **python-pptx** — Read `.pptx` slide content
- **pyttsx3** — Offline text-to-speech (no API keys needed)
- **psutil** — Cross-platform process monitoring
- **pytest / pytest-cov** — Testing (dev)

## Conventions

- Commit messages should be clear and descriptive
- The project has a lighthearted personality — the ending joke is part of the design; maintain that tone in user-facing strings
- Tests mock the TTS engine (`pyttsx3`) so they run silently and fast
- `README.MD` contains the user-facing project description
