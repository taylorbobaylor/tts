"""Voice synthesis using pyttsx3 (offline, free TTS)."""

from __future__ import annotations

from dataclasses import dataclass

import pyttsx3


@dataclass
class VoiceConfig:
    """Configuration for the TTS voice."""

    rate: int = 175          # Words per minute
    volume: float = 1.0      # 0.0 – 1.0
    voice_id: str | None = None  # Platform-specific voice identifier


class Synthesizer:
    """Wraps pyttsx3 to speak text aloud."""

    def __init__(self, config: VoiceConfig | None = None) -> None:
        self._config = config or VoiceConfig()
        self._engine = pyttsx3.init()
        self._apply_config()

    def _apply_config(self) -> None:
        self._engine.setProperty("rate", self._config.rate)
        self._engine.setProperty("volume", self._config.volume)
        if self._config.voice_id:
            self._engine.setProperty("voice", self._config.voice_id)

    def speak(self, text: str) -> None:
        """Speak *text* synchronously (blocks until done)."""
        if not text.strip():
            return
        self._engine.say(text)
        self._engine.runAndWait()

    def list_voices(self) -> list[dict[str, str]]:
        """Return available voices on this system."""
        return [
            {"id": v.id, "name": v.name, "languages": str(v.languages)}
            for v in self._engine.getProperty("voices")
        ]

    def stop(self) -> None:
        """Stop any in-progress speech."""
        self._engine.stop()
