# PDF/PPTX Fullscreen Viewer

A Windows desktop app that lets you:
- Pick a `.pdf` or `.pptx`
- Run it in fullscreen
- Set slide/page timing (`10`, `20`, etc.)
- Use transitions and slide controls

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
