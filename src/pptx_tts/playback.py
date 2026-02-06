"""Playback controller — orchestrates reading a full presentation."""

from __future__ import annotations

import itertools
import logging
import time

from pptx_tts.extractor import SlideContent, extract_slides
from pptx_tts.synthesizer import Synthesizer, VoiceConfig

logger = logging.getLogger(__name__)

ENDING_JOKES = [
    "So, does everything look good? Great! Because I can't hear you — I just read off the presentation.",
    "And that's the last slide! If you have questions, please write them on a sticky note and throw them at the screen.",
    "Presentation complete! I'd take a bow, but I don't have a body. Or legs. Or... anything, really.",
    "That's all folks! I'd ask for applause, but honestly, the silence is less awkward for both of us.",
    "Any questions? Just kidding — I literally cannot process your answers. Good luck out there!",
    "And we're done! If that didn't make sense, don't worry — I just read the words, I don't understand them either.",
    "Thank you for listening! Or sleeping. Either way, my job here is done.",
    "End of presentation! Fun fact: I rehearsed this zero times and still nailed it. Probably.",
    "That concludes today's slides. Remember, if you didn't learn anything, that's a content problem, not a me problem.",
    "And scene! I hope that was informative. If not, at least it was... audible?",
    "Presentation over! I'd stick around for the Q and A, but I have another deck to read in five minutes.",
    "We made it to the end! High five! Oh wait, I'm software. Air five? No air either. Never mind.",
    "That's a wrap! If you need me to read it again, just hit F5. I'll be here. I'm always here.",
    "All done! I hope I pronounced everything correctly. If not, blame the person who made the slides.",
    "And that's the presentation! Now if you'll excuse me, I need to go recharge. Just kidding — I run on pure determination.",
]

# Cycle through jokes so they rotate across presentations
_joke_cycle = itertools.cycle(ENDING_JOKES)


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
        self._stopped = False
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

        if self._tell_joke and not self._stopped:
            self._deliver_ending()

        logger.info("Presentation complete.")

    def finish_with_joke(self) -> None:
        """Stop current speech and deliver an ending joke after a short pause."""
        self._stopped = True
        self._synth.stop()
        if self._tell_joke:
            time.sleep(1.0)
            self._deliver_ending()

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
        """Tell the next joke in rotation."""
        joke = next(_joke_cycle)
        logger.info("Ending: %s", joke)
        self._synth.speak(joke)

    def stop(self) -> None:
        """Signal playback to stop after the current slide."""
        self._stopped = True
        self._synth.stop()
