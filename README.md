# ProSlide Viewer

ProSlide Viewer is a Windows presentation player for PDF and PPTX files with fullscreen playback, custom timing control, transitions, and live presenter controls.

## Why ProSlide Viewer

- Fast fullscreen playback for PDF pages and PPTX slides
- Flexible timing modes for fixed, custom-list, or random duration
- Presenter-first controls for jump, blackout, timing tweaks, direction, and snapshots
- Ready-to-run launcher and EXE build script for deployment

## Features

- Fullscreen viewer for PDF pages and PPTX slides
- User-defined timing modes:
  - One value for all slides (default)
  - Comma-separated timings per slide (example: `10,20,15`)
  - Random timing per slide with min/max range
- Slide controls in fullscreen:
  - `Right` / `Left`: next/previous slide
  - `Space`: pause/resume autoplay
  - `B`: blackout screen toggle
  - `F`: toggle fullscreen on/off
  - `G`: show/hide guide overlays
  - `J`: jump to a slide number
  - `S`: save current slide snapshot
  - `+` / `-`: increase/decrease current slide duration
  - `R`: toggle autoplay direction (forward/reverse)
  - `Esc`: stop and exit fullscreen
- Header/Footer guide overlays with controls help and status
- On-screen `Exit (Esc)` button in fullscreen
- Transition options: `None`, `Fade`, `Slide Left`, `Zoom In`
- Playback options:
  - Start from a specific slide
  - Loop playback on/off
  - Shuffle slide order
  - Countdown timer on/off
  - Pure fullscreen mode (guides stay hidden unless manually toggled)
- Presenter UI enhancements:
  - Progress bar
  - Current clock display
  - Slide timing/status line

## Keyboard Shortcuts (Presentation)

- `Left` / `Right`: Previous or next slide
- `Space`: Pause or resume autoplay
- `B`: Toggle blackout screen
- `F`: Toggle fullscreen state
- `G`: Show or hide guide overlays
- `J`: Jump to a specific slide
- `S`: Save snapshot of current slide
- `+` / `-`: Increase or decrease current slide duration
- `R`: Toggle autoplay direction (forward/reverse)
- `Esc`: Exit presentation

## Requirements

- Windows
- Python 3.10+
- For PPTX support: Microsoft PowerPoint installed

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Optional for PPTX support:

```powershell
pip install pywin32
```

## Run

```powershell
python viewer.py
```

Or use one-click launcher:

```powershell
run_viewer.bat
```

## Project Structure

- `viewer.py`: Main application code
- `requirements.txt`: Python dependencies
- `run_viewer.bat`: One-click local run script
- `build_product.bat`: One-click EXE build script

## Build for Deployment (EXE)

Use:

```powershell
build_product.bat
```

This creates a distributable app at:

- `dist/ProSlideViewer/ProSlideViewer.exe`

## Notes

- PPTX rendering is done by exporting slides through PowerPoint (COM automation, via `pywin32`).
- If custom timing list has fewer values than slide count, the last value is repeated.
- Slide snapshots are exported to the `exports/` folder.

## License

Use your preferred license for distribution (for example MIT) before public release.
