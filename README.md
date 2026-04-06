# USB COM Monitor

Small always-on-top desktop widget that shows all connected COM ports at a glance — VID/PID, connection age, open/free status, and a flash animation when a new port appears.

Useful during embedded development when you need to watch ports enumerate and de-enumerate without keeping Device Manager open.

![dark floating widget with COM port rows]

## Requirements

```
pip install pyserial pywin32
```

`pywin32` is Windows-only and used for reliable "is this port open" detection via `CreateFile` without actually opening the port. Falls back to a pyserial probe if unavailable.

## Run

```
python tools/com_monitor/monitor.py
```

The window has no taskbar entry. Drag it by the title bar; close with **✕**, **Esc**, or **Ctrl+Q**.

## Columns

| Column | What it shows |
|---|---|
| **Port** | COM port name (e.g. `COM19`) |
| **VID:PID** | USB Vendor and Product ID in hex, sourced from Windows SetupAPI — same data as Device Manager, no port opening needed |
| **Age** | Time since the port first appeared (`5s`, `1m23s`, `2h04m`) |
| **●** | Colour dot: green = free, red = open by another process, yellow = appeared < 8 s ago |
| **Status** | `OPEN` / `free` — whether another process holds the port |
| **Serial / Loc** | USB serial number and hub location string from SetupAPI |
| **Description** | Human-readable device name from Windows |

## Flash on connect

When a port appears **after the monitor starts**, the row flashes amber and fades back to the normal background over roughly a minute.

The envelope shape is:

```
brightness(t) = exp(−K × √t)    where t = seconds since connect
```

This gives a fast initial drop (most of the brightness is gone in the first 15 s) followed by a very slow tail — like a struck note decaying. Ports present at startup do not flash.

Text is **bold** while the row is still above ~10 % brightness (first ~30 s with default settings).

## Controls

| Action | Effect |
|---|---|
| Drag title bar | Move window |
| **◑** button | Toggle between opaque and dimmed (35 % alpha) |
| **✕** / Esc / Ctrl+Q | Close |

## Tunables (top of `monitor.py`)

| Constant | Default | Effect |
|---|---|---|
| `REFRESH_MS` | `500` | Poll interval in milliseconds |
| `NEW_DOT_S` | `8` | Seconds the dot/age text stays yellow after connect |
| `C_ROW_FLASH` | `#8a6218` | Peak amber colour of the flash row |
| `FLASH_ATTACK_S` | `0.3` | Attack ramp duration (visual only at sub-second refresh) |
| `FLASH_DECAY_K` | `0.50` | Decay speed — higher = faster early drop, same ~0 at 60 s |
| `BOLD_THRESHOLD` | `0.1` | Brightness level below which text reverts to normal weight |
| `ALPHA_OPAQUE` | `0.96` | Window opacity in normal mode |
| `ALPHA_DIM` | `0.35` | Window opacity when dimmed with ◑ |
