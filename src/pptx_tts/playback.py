"""Playback controller — orchestrates reading a full presentation."""

from __future__ import annotations

import logging
import time

from pptx_tts.extractor import SlideContent, extract_slides
from pptx_tts.synthesizer import Synthesizer, VoiceConfig

logger = logging.getLogger(__name__)

ENDING_JOKES = [
    "Any questions? ...Hello? Oh right, I forgot I'm just a computer program. Classic me!",
    "And that concludes the presentation! I'd take questions, but my listening skills are... non-existent.",
    "That's all folks! Feel free to ask questions — I'll just pretend I can hear you.",
]


class PlaybackController:
    """Reads an entire presentation aloud, slide by slide."""

    def __init__(
        self,
        voice_config: VoiceConfig | None = None,
        slide_delay: float = 1.5,
        tell_joke: bool = True,
    ) -> None:
        self._synth = Synthesizer(voice_config)
        self._slide_delay = slide_delay
        self._tell_joke = tell_joke
        self._stopped = False

    def play_presentation(self, filepath: str) -> None:
        """Extract slides from *filepath* and read them aloud sequentially."""
        slides = extract_slides(filepath)
        if not slides:
            logger.warning("No slides found in %s", filepath)
            return

        logger.info("Starting presentation: %s (%d slides)", filepath, len(slides))

        for slide in slides:
            if self._stopped:
                logger.info("Playback stopped by user")
                return
            self._read_slide(slide)

        if self._tell_joke:
            self._deliver_ending()

        logger.info("Presentation complete.")

    def _read_slide(self, slide: SlideContent) -> None:
        """Read a single slide, then pause before advancing."""
        text = slide.full_text
        if not text:
            logger.debug("Slide %d is empty, skipping", slide.number)
            return

        logger.info("Slide %d: %s", slide.number, text[:80])
        self._synth.speak(text)

        if not self._stopped:
            time.sleep(self._slide_delay)

    def _deliver_ending(self) -> None:
        """Tell a joke at the end of the presentation."""
        import random

        joke = random.choice(ENDING_JOKES)
        logger.info("Ending: %s", joke)
        self._synth.speak(joke)

    def stop(self) -> None:
        """Signal playback to stop after the current slide."""
        self._stopped = True
        self._synth.stop()
