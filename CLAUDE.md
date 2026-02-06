# CLAUDE.md - PowerPoint Text-to-Speech Auto-Reader

## Project Overview

This is a **PowerPoint Text-to-Speech Auto-Reader** — an automation tool that reads PowerPoint presentations aloud using AI text-to-speech voices. It runs as a background service, detects when PowerPoint is opened, extracts slide text, synthesizes speech, and auto-advances slides.

**Status**: Early planning/initialization phase. No implementation code exists yet — only a README describing the project vision.

## Repository Structure

```
tts/
├── README.MD          # Project description, features, and architecture overview
└── CLAUDE.md          # This file
```

## Planned Architecture (from README)

The system is designed around five modules:

1. **Detection** — Monitor for PowerPoint application launches
2. **Content Extraction** — Read text from slides via PowerPoint API
3. **Voice Synthesis** — Convert text to speech using AI voice models
4. **Playback Control** — Manage audio playback and slide timing
5. **Navigation** — Auto-advance slides (1.5s delay after each slide completes)

## Development Setup

No build system, dependency manager, or tooling has been configured yet. When implementation begins, these will need to be set up based on the chosen language/framework.

## Key Decisions Still Needed

- **Language**: Not yet chosen (Python is a likely candidate given TTS library availability)
- **TTS engine**: README mentions "high-quality, free AI text-to-speech voices" — specific engine TBD
- **PowerPoint integration**: COM automation (Windows), python-pptx, or similar
- **Service model**: How the background service runs (systemd, Windows service, tray app, etc.)
- **Dependency management**: pip/poetry/uv for Python, npm for Node, etc.

## Conventions for AI Assistants

- The README.MD contains the authoritative project vision and feature set
- No tests, linting, CI/CD, or formatting tools exist yet — set these up when adding code
- No `.gitignore` exists — create one appropriate for the chosen language when adding code
- Commit messages should be clear and descriptive
- The project has a lighthearted personality (ending joke feature) — maintain that tone in user-facing strings
