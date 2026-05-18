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
import json
import os

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
C_ROW_GONE = "#7a1a1a"   # red peak for vanished-port row flash
C_GONE     = "#a04848"   # foreground for "gone" status
C_DIM      = "#555555"   # darkened text for vanished ports

FONT      = ("Consolas", 10)
FONT_BOLD = ("Consolas", 10, "bold")
FONT_HDR  = ("Consolas", 9, "bold")

ALPHA_OPAQUE   = 0.96
ALPHA_DIM      = 0.35
ALPHA_FADE_MIN = 0.05    # alpha when fade-to-transparent has fully run
FADE_RAMP_S    = 3.0     # seconds to ramp a fade once its threshold is past

COLS = [
    ("Port",          "w", False),
    ("VID:PID",       "w", False),
    ("Age",           "e", False),
    ("",              "c", False),
    ("Status",        "w", False),
    ("Serial / Loc",  "w", False),
    ("Description",   "w", True),
]

# ── settings ──────────────────────────────────────────────────────────────────
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "settings.json")

DEFAULT_SETTINGS = {
    "highlight_duration_s":          20.0,    # effective row-flash fade time
    "always_on_top":                 True,
    "always_on_top_timeout_s":       120,     # 0 = stay on top forever
    "show_removed_s":                60,      # 0 = don't show removed ports at all
    "move_to_back_enable":           False,   # push window to back of Z-order on idle
    "move_to_back_timeout_s":        300,
    "fade_to_transparent_enable":    False,   # reduce window alpha on idle
    "fade_to_transparent_timeout_s": 600,
    "normal_alpha":                  ALPHA_OPAQUE,    # window opacity at rest
    "faded_alpha":                   ALPHA_FADE_MIN,  # target opacity when fully faded
    "window_x":                      None,    # last-known position; None = let WM place
    "window_y":                      None,
}


def load_settings() -> dict:
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return dict(DEFAULT_SETTINGS)
    merged = dict(DEFAULT_SETTINGS)
    for k, v in data.items():
        if k in DEFAULT_SETTINGS:
            merged[k] = v
    return merged


def save_settings(s: dict) -> None:
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(s, f, indent=2)
    except Exception as e:
        print(f"Failed to save settings: {e}")


# ── flash curve  ─────────────────────────────────────────────────────────────
def _flash_brightness(t: float, k: float = FLASH_DECAY_K) -> float:
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
    return math.exp(-k * math.sqrt(t - FLASH_ATTACK_S))


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
        self.settings = load_settings()

        self.overrideredirect(True)
        self.attributes("-topmost", self.settings["always_on_top"])
        self.attributes("-alpha", self.settings["normal_alpha"])
        self.configure(bg=C_BG)

        self._first_seen:     dict[str, float] = {}   # device → first-seen time
        self._flash_start:    dict[str, float] = {}   # device → flash start time (new arrivals only)
        self._disappeared_at: dict[str, float] = {}   # device → time it vanished
        self._port_info:      dict[str, dict]  = {}   # device → last-known port info
        self._initialized  = False                    # skip flash for ports present at startup
        self._row_widgets: list[list[tk.Widget]] = []
        self._drag_ox = self._drag_oy = 0
        self._dimmed  = False

        # AOT / push-to-back state
        self._known_ports: set[str] = set()
        self._last_change = time.time()
        self._topmost_active = self.settings["always_on_top"]
        self._bg_active      = False
        self._flash_decay_k = self._compute_flash_decay_k()

        self._build_ui()
        self._refresh()
        self._restore_window_position()

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

        tk.Button(bar, text=" ⚙ ", bg=C_HDR, fg="#666666", bd=0, relief="flat",
                  font=("Segoe UI", 8),
                  activebackground=C_HDR, activeforeground="#aaaaaa",
                  command=self._open_settings).pack(side=tk.RIGHT)

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

        # interaction wakes the window from any idle fade / push-to-back
        self.bind("<Button>",  self._wake)
        self.bind("<Enter>",   self._wake)
        self.bind("<FocusIn>", self._wake)

    # ── transparency toggle ───────────────────────────────────────────────────
    def _toggle_dim(self):
        self._dimmed = not self._dimmed
        self._dim_btn.config(fg="#aaaaaa" if self._dimmed else "#666666")
        self._update_alpha()

    def _update_alpha(self):
        normal = self.settings.get("normal_alpha", ALPHA_OPAQUE)
        faded  = self.settings.get("faded_alpha",  ALPHA_FADE_MIN)
        base = ALPHA_DIM if self._dimmed else normal
        f = 0.0
        if self.settings.get("fade_to_transparent_enable"):
            threshold = self.settings["fade_to_transparent_timeout_s"]
            idle = time.time() - self._last_change
            if idle > threshold:
                f = min(1.0, (idle - threshold) / FADE_RAMP_S)
        target = base - (base - faded) * f
        self.attributes("-alpha", target)

    def _wake(self, _event=None):
        """Restart the idle timeout — pull the window out of any fade or push-to-back."""
        self._last_change = time.time()
        if self.settings["always_on_top"] and not self._topmost_active:
            self.attributes("-topmost", True)
            self._topmost_active = True
        if self._bg_active:
            self.lift()
            self._bg_active = False
        self._update_alpha()

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

    # ── settings ──────────────────────────────────────────────────────────────
    def _compute_flash_decay_k(self) -> float:
        # Solve exp(-K * sqrt(d - attack)) = 0.1 → K = 2.3 / sqrt(d - attack)
        d = max(self.settings["highlight_duration_s"] - FLASH_ATTACK_S, 0.1)
        return 2.3 / math.sqrt(d)

    def _restore_window_position(self):
        x = self.settings.get("window_x")
        y = self.settings.get("window_y")
        if x is None or y is None:
            return
        self.update_idletasks()
        # clamp to current virtual screen so a saved off-screen pos is recoverable
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = max(0, min(int(x), sw - 50))
        y = max(0, min(int(y), sh - 50))
        self.geometry(f"+{x}+{y}")

    def _save_window_position(self):
        try:
            self.settings["window_x"] = self.winfo_x()
            self.settings["window_y"] = self.winfo_y()
            save_settings(self.settings)
        except Exception:
            pass

    def destroy(self):
        self._save_window_position()
        super().destroy()

    def _apply_settings(self):
        """Re-apply settings to derived state after they've been edited."""
        self._flash_decay_k  = self._compute_flash_decay_k()
        self.attributes("-topmost", self.settings["always_on_top"])
        self._topmost_active = self.settings["always_on_top"]
        self._last_change    = time.time()
        self._update_alpha()

    def _open_settings(self):
        SettingsDialog(self)

    # ── refresh ───────────────────────────────────────────────────────────────
    def _refresh(self):
        now   = time.time()
        ports = sorted(serial.tools.list_ports.comports(), key=lambda p: p.device)

        # update registries — active ports
        current = {p.device for p in ports}
        for p in ports:
            self._port_info[p.device] = {
                "vid": p.vid, "pid": p.pid,
                "description":   p.description or "",
                "serial_number": p.serial_number or "",
                "location":      p.location or "",
            }
            if p.device not in self._first_seen:
                self._first_seen[p.device] = now
                if self._initialized:           # don't flash ports seen at startup
                    # shift back so first render lands at peak brightness, not t=0
                    self._flash_start[p.device] = now - FLASH_ATTACK_S
            if p.device in self._disappeared_at:
                # port came back — clear gone marker and re-flash as a new arrival
                del self._disappeared_at[p.device]
                self._flash_start[p.device] = now - FLASH_ATTACK_S

        # mark newly-vanished ports (skip if user has disabled the feature)
        if self.settings["show_removed_s"] > 0:
            for dev in list(self._first_seen):
                if dev not in current and dev not in self._disappeared_at:
                    self._disappeared_at[dev] = now
                    self._flash_start.pop(dev, None)

        # expire vanished ports past the configured display duration
        show_removed = self.settings["show_removed_s"]
        for dev, t in list(self._disappeared_at.items()):
            if show_removed == 0 or (now - t) > show_removed:
                del self._disappeared_at[dev]

        # purge registries for devices that are neither active nor displayed-as-gone
        for dev in list(self._first_seen):
            if dev not in current and dev not in self._disappeared_at:
                self._first_seen.pop(dev, None)
                self._flash_start.pop(dev, None)
                self._port_info.pop(dev, None)

        # AOT timeout: re-enable on port change, drop after idle period
        if self._initialized and current != self._known_ports:
            self._last_change = now
            if self.settings["always_on_top"] and not self._topmost_active:
                self.attributes("-topmost", True)
                self._topmost_active = True
        self._known_ports = current

        if self.settings["always_on_top"] and self._topmost_active:
            timeout = self.settings["always_on_top_timeout_s"]
            if timeout > 0 and (now - self._last_change) > timeout:
                self.attributes("-topmost", False)
                self._topmost_active = False

        # idle fades: window alpha
        self._update_alpha()

        # push-to-back on idle (opposite of always-on-top)
        desired_back = False
        if self.settings.get("move_to_back_enable"):
            threshold = self.settings["move_to_back_timeout_s"]
            if threshold > 0 and (now - self._last_change) > threshold:
                desired_back = True
        if desired_back:
            if self._topmost_active:
                self.attributes("-topmost", False)
                self._topmost_active = False
            self.lower()                   # keep re-lowering so newly opened windows pass us
            self._bg_active = True
        elif self._bg_active:
            self.lift()
            self._bg_active = False

        self._initialized = True

        # rebuild rows
        for row in self._row_widgets:
            for w in row:
                w.destroy()
        self._row_widgets.clear()
        self._empty_lbl.grid_forget()

        rendered = sorted(set(current) | set(self._disappeared_at))
        if not rendered:
            self._empty_lbl.grid(row=2, column=0, columnspan=len(COLS),
                                 sticky="w", pady=6)
        else:
            for i, dev in enumerate(rendered):
                gone = dev in self._disappeared_at
                info = self._port_info.get(dev, {})

                vid = f"{info.get('vid'):04X}" if info.get("vid") is not None else "----"
                pid = f"{info.get('pid'):04X}" if info.get("pid") is not None else "----"

                ser_loc = " / ".join(filter(None, [info.get("serial_number", ""),
                                                   info.get("location", "")])) or "—"
                desc = info.get("description", "")
                if desc == dev:
                    desc = ""

                if gone:
                    age_s      = now - self._disappeared_at[dev]
                    brightness = _flash_brightness(age_s, self._flash_decay_k)
                    row_bg     = _blend(C_ROW_GONE, C_BG, 1.0 - brightness)
                    row_font   = FONT
                    dot_c      = C_GONE
                    age_c      = C_DIM
                    stat_t     = "gone"
                    stat_c     = C_GONE
                    port_c     = C_DIM
                    vid_c      = C_DIM
                    ser_c      = C_DIM
                    desc_c     = C_DIM
                else:
                    age_s    = now - self._first_seen[dev]
                    occupied = is_open(dev)

                    flash_t = now - self._flash_start[dev] if dev in self._flash_start else None
                    if flash_t is not None:
                        brightness = _flash_brightness(flash_t, self._flash_decay_k)
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
                    port_c = C_PORT
                    vid_c  = C_VIDPID
                    ser_c  = C_SER
                    desc_c = C_DESC

                cells = [
                    dict(text=dev,                    fg=port_c, anchor="w"),
                    dict(text=f"{vid}:{pid}",         fg=vid_c,  anchor="w"),
                    dict(text=self._age_str(age_s),   fg=age_c,  anchor="e"),
                    dict(text="●",                    fg=dot_c,  anchor="center"),
                    dict(text=stat_t,                 fg=stat_c, anchor="w"),
                    dict(text=ser_loc,                fg=ser_c,  anchor="w"),
                    dict(text=desc,                   fg=desc_c, anchor="w"),
                ]
                row_w = []
                for c, kw in enumerate(cells):
                    lbl = tk.Label(self._grid, bg=row_bg, font=row_font,
                                   padx=3, pady=1, **kw)
                    lbl.grid(row=i + 2, column=c, sticky="ew")
                    row_w.append(lbl)
                self._row_widgets.append(row_w)

        self.after(REFRESH_MS, self._refresh)


# ── settings dialog ───────────────────────────────────────────────────────────
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent: ComMonitor):
        super().__init__(parent)
        self.parent = parent
        self.title("Settings")
        self.configure(bg=C_BG)
        self.transient(parent)
        self.resizable(False, False)
        self.attributes("-topmost", True)

        # snapshot for revert-on-cancel (live-preview sliders mutate parent.settings)
        self._snapshot = dict(parent.settings)
        self._saved    = False

        self._build()

        self.update_idletasks()
        dw, dh = self.winfo_width(), self.winfo_height()
        # bias halfway between parent-centered and screen-centered so the dialog
        # drifts toward the middle when the parent is parked in a corner
        parent_cx = parent.winfo_rootx() + parent.winfo_width()  // 2
        parent_cy = parent.winfo_rooty() + parent.winfo_height() // 2
        screen_cx = self.winfo_screenwidth()  // 2
        screen_cy = self.winfo_screenheight() // 2
        cx = (parent_cx + screen_cx) // 2
        cy = (parent_cy + screen_cy) // 2
        px = cx - dw // 2
        py = cy - dh // 2
        self.geometry(f"+{max(px, 0)}+{max(py, 0)}")
        self.grab_set()
        self.bind("<Escape>", lambda _: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

    def _build(self):
        s   = self.parent.settings
        pad = dict(padx=8, pady=4)

        frm = tk.Frame(self, bg=C_BG)
        frm.pack(fill=tk.BOTH, expand=True, padx=12, pady=(12, 6))

        # Row highlight duration
        tk.Label(frm, text="Row highlight duration (s):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=0, column=0, sticky="w", **pad)
        self.var_hl = tk.DoubleVar(value=s["highlight_duration_s"])
        tk.Spinbox(frm, from_=1, to=600, increment=1, textvariable=self.var_hl,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=0, column=1, sticky="w", **pad)

        # Show removed ports for
        tk.Label(frm, text="Show removed ports for (s, 0 = hide):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=1, column=0, sticky="w", **pad)
        self.var_sr = tk.IntVar(value=s["show_removed_s"])
        tk.Spinbox(frm, from_=0, to=86400, increment=10, textvariable=self.var_sr,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=1, column=1, sticky="w", **pad)

        # Always on top
        self.var_aot = tk.BooleanVar(value=s["always_on_top"])
        tk.Checkbutton(frm, text="Always on top",
                       variable=self.var_aot,
                       bg=C_BG, fg=C_DESC, selectcolor=C_HDR,
                       activebackground=C_BG, activeforeground=C_DESC,
                       font=FONT, anchor="w"
                       ).grid(row=2, column=0, columnspan=2, sticky="w", **pad)

        # AOT idle timeout
        tk.Label(frm, text="     ↳ drop after idle (s, 0 = never):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=3, column=0, sticky="w", **pad)
        self.var_to = tk.IntVar(value=s["always_on_top_timeout_s"])
        tk.Spinbox(frm, from_=0, to=86400, increment=10, textvariable=self.var_to,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=3, column=1, sticky="w", **pad)

        # Move window to back on idle
        self.var_mtb = tk.BooleanVar(value=s["move_to_back_enable"])
        tk.Checkbutton(frm, text="Move to background on idle",
                       variable=self.var_mtb,
                       bg=C_BG, fg=C_DESC, selectcolor=C_HDR,
                       activebackground=C_BG, activeforeground=C_DESC,
                       font=FONT, anchor="w"
                       ).grid(row=4, column=0, columnspan=2, sticky="w", **pad)
        tk.Label(frm, text="     ↳ after (s):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=5, column=0, sticky="w", **pad)
        self.var_mtb_to = tk.IntVar(value=s["move_to_back_timeout_s"])
        tk.Spinbox(frm, from_=1, to=86400, increment=10, textvariable=self.var_mtb_to,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=5, column=1, sticky="w", **pad)

        # Fade to transparent on idle
        self.var_ftr = tk.BooleanVar(value=s["fade_to_transparent_enable"])
        tk.Checkbutton(frm, text="Fade window to transparent on idle",
                       variable=self.var_ftr,
                       bg=C_BG, fg=C_DESC, selectcolor=C_HDR,
                       activebackground=C_BG, activeforeground=C_DESC,
                       font=FONT, anchor="w"
                       ).grid(row=6, column=0, columnspan=2, sticky="w", **pad)
        tk.Label(frm, text="     ↳ after (s):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=7, column=0, sticky="w", **pad)
        self.var_ftr_to = tk.IntVar(value=s["fade_to_transparent_timeout_s"])
        tk.Spinbox(frm, from_=1, to=86400, increment=10, textvariable=self.var_ftr_to,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=7, column=1, sticky="w", **pad)

        # Normal / faded transparency sliders (live preview via _preview_alpha)
        tk.Label(frm, text="Normal transparency:",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=8, column=0, sticky="w", **pad)
        self.var_an = tk.DoubleVar(value=s["normal_alpha"])
        tk.Scale(frm, from_=0.10, to=1.0, resolution=0.01,
                 orient=tk.HORIZONTAL, variable=self.var_an,
                 command=lambda v: self._preview_alpha("normal_alpha", v),
                 bg=C_BG, fg=C_DESC, troughcolor=C_HDR,
                 highlightthickness=0, bd=0, length=160,
                 font=FONT, activebackground=C_PORT,
                 ).grid(row=8, column=1, sticky="w", **pad)

        tk.Label(frm, text="Faded transparency:",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=9, column=0, sticky="w", **pad)
        self.var_af = tk.DoubleVar(value=s["faded_alpha"])
        tk.Scale(frm, from_=0.0, to=1.0, resolution=0.01,
                 orient=tk.HORIZONTAL, variable=self.var_af,
                 command=lambda v: self._preview_alpha("faded_alpha", v),
                 bg=C_BG, fg=C_DESC, troughcolor=C_HDR,
                 highlightthickness=0, bd=0, length=160,
                 font=FONT, activebackground=C_PORT,
                 ).grid(row=9, column=1, sticky="w", **pad)

        # Buttons
        btns = tk.Frame(self, bg=C_BG)
        btns.pack(fill=tk.X, padx=12, pady=(6, 12))
        tk.Button(btns, text="Cancel", command=self._cancel,
                  bg=C_HDR, fg=C_DESC, bd=0, relief="flat", font=FONT,
                  activebackground=C_HDR, activeforeground="white",
                  padx=14, pady=3).pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(btns, text="Save", command=self._save,
                  bg=C_HDR, fg=C_PORT, bd=0, relief="flat", font=FONT,
                  activebackground=C_HDR, activeforeground="white",
                  padx=14, pady=3).pack(side=tk.RIGHT)

    def _preview_alpha(self, key: str, val):
        try:
            self.parent.settings[key] = float(val)
        except (TypeError, ValueError):
            return
        self.parent._update_alpha()

    def _cancel(self):
        if not self._saved:
            self.parent.settings.clear()
            self.parent.settings.update(self._snapshot)
            self.parent._update_alpha()
        self.destroy()

    def _save(self):
        s = self.parent.settings
        try:
            v = float(self.var_hl.get())
            if v > 0:
                s["highlight_duration_s"] = v
        except (tk.TclError, ValueError):
            pass
        s["always_on_top"] = bool(self.var_aot.get())
        try:
            v = int(self.var_to.get())
            if v >= 0:
                s["always_on_top_timeout_s"] = v
        except (tk.TclError, ValueError):
            pass
        try:
            v = int(self.var_sr.get())
            if v >= 0:
                s["show_removed_s"] = v
        except (tk.TclError, ValueError):
            pass
        s["move_to_back_enable"] = bool(self.var_mtb.get())
        try:
            v = int(self.var_mtb_to.get())
            if v > 0:
                s["move_to_back_timeout_s"] = v
        except (tk.TclError, ValueError):
            pass
        s["fade_to_transparent_enable"] = bool(self.var_ftr.get())
        try:
            v = int(self.var_ftr_to.get())
            if v > 0:
                s["fade_to_transparent_timeout_s"] = v
        except (tk.TclError, ValueError):
            pass
        try:
            s["normal_alpha"] = float(self.var_an.get())
        except (tk.TclError, ValueError):
            pass
        try:
            s["faded_alpha"] = float(self.var_af.get())
        except (tk.TclError, ValueError):
            pass
        save_settings(s)
        self.parent._apply_settings()
        self._saved = True
        self.destroy()


if __name__ == "__main__":
    ComMonitor().mainloop()
