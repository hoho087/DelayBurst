# RAIN-DelayBurst

[中文](https://github.com/hoho087/DelayBurst/blob/main/README.md)

A Windows network testing utility (`PySide6 + WinDivert`) for simulating outbound/inbound packet behavior with **squeeze replay** or **drop mode** on a selected target process.

![image](https://github.com/hoho087/DelayBurst/blob/main/screenshot.png)

## Quick Start

1. Install dependencies
```powershell
pip install -r requirements.txt
```

2. Keep these files in the same folder
- `RAIN-DelayBurst.py`
- `WinDivert.dll`
- `WinDivert64.sys`

3. Run as Administrator
```powershell
python RAIN-DelayBurst.py
```

## Usage

1. For target selection, use `Select Target` (pick `.exe`) or `Select Process` (pick from running processes, then choose PID-only or same-path scope)
2. Click `Bind Hotkey`, then press the key you want
3. Configure outbound / inbound settings independently
4. Use hotkey or `Toggle Effect` to start, `End Effect` to stop
5. Settings are auto-saved to `RAIN-DelayBurst.config.json` in the same folder and auto-loaded on next launch
6. You can still use `Save Config` / `Load Config` manually

## Modes

- `Squeeze Replay`: packets are held first, then replayed after effect ends
- `Drop Mode`: packets are dropped by probability while effect is active

## How It Works

1. Uses `tasklist` / `netstat` to find the target PID and active local ports.
2. Builds WinDivert filters for outbound/inbound traffic.
3. In `Squeeze Replay`, packets are queued, then replayed using jitter/bandwidth/loss settings.
4. In `Drop Mode`, packets are dropped according to configured drop rate.

## Library Used

- WinDivert: <https://github.com/basil00/WinDivert>

## Common Issues

- `Cannot initialize Qt platform plugin`: rebuild with correct PySide6 plugin settings.
- `WinDivert DLL not found`: place `WinDivert.dll` beside the executable/script.
- `Run as Administrator required`: start terminal/app with admin privilege.
