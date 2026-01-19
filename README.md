# MSFS Sim Rate Monitor

As we all know, Microsoft Flight Simulator 2024 still doesn't give us any proper way to know what the **current sim rate** is.  
A lot of people use sim rate to speed up cruise, but the game only gives you accelerate/decelerate — no actual display.  

This little tool simply **shows the current sim rate**, so you can have a smoother experience.

## Features

- **Real-time sim rate display** — Always know your current speed
- **Overlay controls** — Increase/decrease sim rate with `<` `>` buttons directly on overlay
- **Speed-based colors** — Cyan for slow, white for normal, orange for fast
- **Auto-hide** — Overlay hides when MSFS loses focus
- **Position memory** — Overlay position persists across app restarts
- **Start with Windows** — Optional auto-start via Startup folder
- **Modern UI** — Dark theme with JetBrains Mono font

## QuickStart

1. Download `MSFS-SimRate-Monitor.exe` from [Releases](../../releases)
2. Run it after or before MSFS, doesn't matter
3. Drag the overlay to your preferred position

## Screenshots

<img width="1120" height="928" alt="New Design" src="https://github.com/user-attachments/assets/d7d6bd6a-2668-44b9-873e-9bb51734b3aa" />


## Configuration

| Option | Description |
|--------|-------------|
| **Overlay Size** | S / M / L / XL / Hide |
| **Auto-hide** | Hide overlay when MSFS is not focused |
| **Start with Windows** | Launch automatically on system startup |

Config file location: `%APPDATA%\MSFS-SimRateMonitor\config.json`

## Development

```bash
# Install dependencies
uv sync

# Run
uv run mini_gui.py

# Build
.\build.bat
```

## Tech Stack

- Python 3.11+ with Tkinter
- SimConnect library for MSFS integration
- PyInstaller for packaging

## License

MIT License
