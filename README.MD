# PowerPoint Text-to-Speech Auto-Reader

## Overview
This project automatically reads PowerPoint presentations aloud using natural-sounding AI voices. It runs silently in the background and springs into action whenever you open a PowerPoint file.

## Features

### 🎤 Natural Voice Synthesis
- Uses high-quality, free AI text-to-speech voices
- Natural intonation and pacing that doesn't sound robotic
- Optimized tone for presentation content

### 🔄 Automatic Slide Navigation
- Reads each slide's content completely
- Waits 1.5 seconds after finishing before advancing
- Automatically flips to the next slide
- Continues until the end of the presentation

### 🤖 Background Operation
- Runs as a background service
- Automatically detects when PowerPoint is opened
- No manual intervention needed
- Lightweight and unobtrusive

### 😄 Personality
At the end of the presentation, the system will ask if there are any questions and crack a joke about not being able to hear your response (because, well, it can't).

## How It Works

1. **Detection**: The service monitors for PowerPoint application launches
2. **Content Extraction**: Reads text content from each slide
3. **Voice Synthesis**: Converts text to speech using AI voice models
4. **Playback Control**: Manages audio playback and slide timing
5. **Navigation**: Automatically advances slides at appropriate intervals

## Technical Details

- Runs as a background process/service
- Integrates with PowerPoint's API for slide detection and navigation
- Uses modern TTS engines for natural voice generation
- Configurable timing between slides (default: 1.5 seconds)

## Getting Started

(Installation and setup instructions would go here)

## Configuration

- Voice selection
- Reading speed
- Pause duration between slides
- Enable/disable the end-of-presentation joke

## Notes

This is designed to make presentations more accessible and to provide a hands-free presentation experience. Perfect for rehearsing presentations, accessibility needs, or just letting your slides present themselves!

---

*"Any questions? ...Hello? Oh right, I forgot I'm just a computer program. Classic me!"* 🎭
```