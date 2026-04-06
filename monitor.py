#!/usr/bin/env python3
"""
USB COM Port Monitor
Small always-on-top desktop widget.

Usage:  python monitor.py
Needs:  pip install pyserial pywin32
"""

import tkinter as tk
import math
import time

try:
    import serial.tools.list_ports
except ImportError:
    print("We need pyserial library (and optionaly pywin32).");

try:
    import win32file, pywintypes
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# ── tunables ──────────────────────────────────────────────────────────────────
REFRESH_MS      = 500
NEW_DOT_S       = 8       # seconds the dot/age text stays yellow

C_ROW_FLASH     = "#8a6218"   # amber peak colour for new-port flash
FLASH_ATTACK_S  = 0.3         # seconds to reach peak brightness
FLASH_DECAY_K   = 0.50        # controls decay speed (higher = faster early drop)
#   brightness(t) = exp(-K * sqrt(t - attack))  ← ~0 by 60 s

BOLD_THRESHOLD  = 0.1        # keep bold text while brightness is above this

# ── palette ───────────────────────────────────────────────────────────────────
C_BG     = "#1a1a1a"
C_HDR    = "#252525"
C_PORT   = "#61dafb"
C_VIDPID = "#f0e68c"
C_FREE   = "#7ec87e"
C_OPEN   = "#e06c75"
C_NEW    = "#ffcc44"
C_AGE    = "#666666"
C_DESC   = "#999999"
C_HEAD   = "#444444"
C_SER    = "#b0a0d0"

FONT      = ("Consolas", 10)
FONT_BOLD = ("Consolas", 10, "bold")
FONT_HDR  = ("Consolas", 9, "bold")

ALPHA_OPAQUE = 0.96
ALPHA_DIM    = 0.35

COLS = [
    ("Port",          "w", False),
    ("VID:PID",       "w", False),
    ("Age",           "e", False),
    ("",              "c", False),
    ("Status",        "w", False),
    ("Serial / Loc",  "w", False),
    ("Description",   "w", True),
]


# ── flash curve  ─────────────────────────────────────────────────────────────
def _flash_brightness(t: float) -> float:
    """
    0..1 flash intensity at t seconds after port appeared.

    Shape: linear attack up to FLASH_ATTACK_S, then
           exp(-K * sqrt(t - attack))  — fast initial drop, very slow tail.
    Reaches ≈0 around 60 s for default K=0.65.
    """
    if t <= 0:
        return 0.0
    if t < FLASH_ATTACK_S:
        return t / FLASH_ATTACK_S          # quick linear ramp to 1.0
    return math.exp(-FLASH_DECAY_K * math.sqrt(t - FLASH_ATTACK_S))


def _blend(c1: str, c2: str, t: float) -> str:
    """Interpolate two #rrggbb colours: t=0 → c1, t=1 → c2."""
    t = max(0.0, min(1.0, t))
    r = int(int(c1[1:3], 16) + (int(c2[1:3], 16) - int(c1[1:3], 16)) * t)
    g = int(int(c1[3:5], 16) + (int(c2[3:5], 16) - int(c1[3:5], 16)) * t)
    b = int(int(c1[5:7], 16) + (int(c2[5:7], 16) - int(c1[5:7], 16)) * t)
    return f"#{r:02x}{g:02x}{b:02x}"


# ── port-open detection ───────────────────────────────────────────────────────
def _is_open_win32(device: str) -> bool:
    path = r"\\.\ "[:-1] + device
    try:
        h = win32file.CreateFile(
            path,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0, None, win32file.OPEN_EXISTING, 0, None,
        )
        win32file.CloseHandle(h)
        return False
    except pywintypes.error as e:
        if e.winerror in (5, 32):
            return True
        return False


def _is_open_fallback(device: str) -> bool:
    import serial
    try:
        s = serial.Serial(device, timeout=0)
        s.close()
        return False
    except Exception:
        return True


def is_open(device: str) -> bool:
    return _is_open_win32(device) if HAS_WIN32 else _is_open_fallback(device)


# ── main window ───────────────────────────────────────────────────────────────
class ComMonitor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.attributes("-alpha", ALPHA_OPAQUE)
        self.configure(bg=C_BG)

        self._first_seen:  dict[str, float] = {}   # device → first-seen time
        self._flash_start: dict[str, float] = {}   # device → flash start time (new arrivals only)
        self._initialized  = False                 # skip flash for ports present at startup
        self._row_widgets: list[list[tk.Widget]] = []
        self._drag_ox = self._drag_oy = 0
        self._dimmed  = False

        self._build_ui()
        self._refresh()

    # ── layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        bar = tk.Frame(self, bg=C_HDR, cursor="fleur")
        bar.pack(fill=tk.X)
        bar.bind("<ButtonPress-1>", self._drag_start)
        bar.bind("<B1-Motion>",     self._drag_move)

        tk.Label(bar, text="  USB COM Monitor",
                 bg=C_HDR, fg="#666666", font=("Segoe UI", 8),
                 pady=4).pack(side=tk.LEFT)
        tk.Button(bar, text=" ✕ ", bg=C_HDR, fg="#666666", bd=0, relief="flat",
                  font=("Segoe UI", 8),
                  activebackground="#c0392b", activeforeground="white",
                  command=self.destroy).pack(side=tk.RIGHT)

        self._dim_btn = tk.Button(
            bar, text=" ◑ ", bg=C_HDR, fg="#666666", bd=0, relief="flat",
            font=("Segoe UI", 8),
            activebackground=C_HDR, activeforeground="#aaaaaa",
            command=self._toggle_dim)
        self._dim_btn.pack(side=tk.RIGHT)

        self._grid = tk.Frame(self, bg=C_BG)
        self._grid.pack(fill=tk.BOTH, expand=True, padx=6, pady=(3, 6))

        for c, (name, anchor, stretch) in enumerate(COLS):
            tk.Label(self._grid, text=name, bg=C_BG, fg=C_HEAD,
                     font=FONT_HDR, anchor=anchor, padx=3
                     ).grid(row=0, column=c, sticky="ew")
            if stretch:
                self._grid.columnconfigure(c, weight=1)

        tk.Frame(self._grid, bg="#333333", height=1).grid(
            row=1, column=0, columnspan=len(COLS), sticky="ew", pady=(1, 3))

        self._empty_lbl = tk.Label(
            self._grid, text="(no COM ports detected)",
            bg=C_BG, fg="#3a3a3a", font=FONT, anchor="w", padx=3)

        self.bind("<Escape>",    lambda _: self.destroy())
        self.bind("<Control-q>", lambda _: self.destroy())

    # ── transparency toggle ───────────────────────────────────────────────────
    def _toggle_dim(self):
        self._dimmed = not self._dimmed
        self.attributes("-alpha", ALPHA_DIM if self._dimmed else ALPHA_OPAQUE)
        self._dim_btn.config(fg="#aaaaaa" if self._dimmed else "#666666")

    # ── drag ─────────────────────────────────────────────────────────────────
    def _drag_start(self, e):
        self._drag_ox = e.x_root - self.winfo_x()
        self._drag_oy = e.y_root - self.winfo_y()

    def _drag_move(self, e):
        self.geometry(f"+{e.x_root - self._drag_ox}+{e.y_root - self._drag_oy}")

    # ── helpers ───────────────────────────────────────────────────────────────
    @staticmethod
    def _age_str(s: float) -> str:
        s = int(s)
        if s < 60:   return f"{s}s"
        m, s = divmod(s, 60)
        if m < 60:   return f"{m}m{s:02d}s"
        h, m = divmod(m, 60)
        return f"{h}h{m:02d}m"

    # ── refresh ───────────────────────────────────────────────────────────────
    def _refresh(self):
        now   = time.time()
        ports = sorted(serial.tools.list_ports.comports(), key=lambda p: p.device)

        # update registries
        current = {p.device for p in ports}
        for dev in list(self._first_seen):
            if dev not in current:
                del self._first_seen[dev]
                self._flash_start.pop(dev, None)
        for p in ports:
            if p.device not in self._first_seen:
                self._first_seen[p.device] = now
                if self._initialized:           # don't flash ports seen at startup
                    # shift back so first render lands at peak brightness, not t=0
                    self._flash_start[p.device] = now - FLASH_ATTACK_S
        self._initialized = True

        # rebuild rows
        for row in self._row_widgets:
            for w in row:
                w.destroy()
        self._row_widgets.clear()
        self._empty_lbl.grid_forget()

        if not ports:
            self._empty_lbl.grid(row=2, column=0, columnspan=len(COLS),
                                 sticky="w", pady=6)
        else:
            for i, p in enumerate(ports):
                age_s    = now - self._first_seen[p.device]
                occupied = is_open(p.device)

                vid = f"{p.vid:04X}" if p.vid is not None else "----"
                pid = f"{p.pid:04X}" if p.pid is not None else "----"

                serial_no = p.serial_number or ""
                location  = p.location      or ""
                ser_loc   = " / ".join(filter(None, [serial_no, location])) or "—"

                desc = p.description or ""
                if desc == p.device:
                    desc = ""

                # flash envelope
                flash_t = now - self._flash_start[p.device] if p.device in self._flash_start else None
                if flash_t is not None:
                    brightness = _flash_brightness(flash_t)
                    row_bg     = _blend(C_ROW_FLASH, C_BG, 1.0 - brightness)
                    row_font   = FONT_BOLD if brightness > BOLD_THRESHOLD else FONT
                else:
                    row_bg   = C_BG
                    row_font = FONT

                fresh  = age_s < NEW_DOT_S
                dot_c  = C_NEW if fresh else (C_OPEN if occupied else C_FREE)
                age_c  = C_NEW if fresh else C_AGE
                stat_t = "OPEN" if occupied else "free"
                stat_c = C_OPEN if occupied else C_FREE

                cells = [
                    dict(text=p.device,              fg=C_PORT,   anchor="w"),
                    dict(text=f"{vid}:{pid}",         fg=C_VIDPID, anchor="w"),
                    dict(text=self._age_str(age_s),   fg=age_c,    anchor="e"),
                    dict(text="●",                    fg=dot_c,    anchor="center"),
                    dict(text=stat_t,                 fg=stat_c,   anchor="w"),
                    dict(text=ser_loc,                fg=C_SER,    anchor="w"),
                    dict(text=desc,                   fg=C_DESC,   anchor="w"),
                ]
                row_w = []
                for c, kw in enumerate(cells):
                    lbl = tk.Label(self._grid, bg=row_bg, font=row_font,
                                   padx=3, pady=1, **kw)
                    lbl.grid(row=i + 2, column=c, sticky="ew")
                    row_w.append(lbl)
                self._row_widgets.append(row_w)

        self.after(REFRESH_MS, self._refresh)


if __name__ == "__main__":
    ComMonitor().mainloop()
