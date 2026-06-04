#!/usr/bin/env python3
"""
USB COM Port Monitor
Small always-on-top desktop widget.

Usage:  python monitor.py
Needs:  pip install pyserial pywin32
"""

import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog
import math
import time
import json
import os
import sys

try:
    import serial.tools.list_ports
except ImportError:
    print("We need pyserial library (and optionaly pywin32).");

try:
    import win32file, pywintypes
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# ── virtual-desktop confinement (Windows) ──────────────────────────────────
# Borderless (override-redirect) windows carry no taskbar button, so Windows'
# Virtual Desktop Manager never tracks them and shows them on *every* virtual
# desktop. To confine the widget to one desktop it has to become trackable,
# which means a taskbar button: we force WS_EX_APPWINDOW on the main strip and
# make the rows window + terminals owned by it, so they follow its desktop
# without buttons of their own. One taskbar button total.
if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

    _GWL_EXSTYLE      = -20
    _GWLP_HWNDPARENT  = -8
    _WS_EX_TOOLWINDOW = 0x00000080
    _WS_EX_APPWINDOW  = 0x00040000
    _GA_ROOT          = 2
    _SW_HIDE          = 0
    _SW_SHOWNA        = 8        # show, don't activate (keep focus where it is)
    _WM_SETICON       = 0x0080
    _ICON_SMALL       = 0
    _ICON_BIG         = 1
    _IMAGE_ICON       = 1
    _LR_LOADFROMFILE  = 0x0010
    _SM_CXICON        = 11
    _SM_CYICON        = 12
    _SM_CXSMICON      = 49
    _SM_CYSMICON      = 50
    _GCLP_HICON       = -14
    _GCLP_HICONSM     = -34

    _user32 = ctypes.windll.user32
    _user32.GetAncestor.restype  = wintypes.HWND
    _user32.GetAncestor.argtypes = (wintypes.HWND, ctypes.c_uint)
    _user32.GetSystemMetrics.restype  = ctypes.c_int
    _user32.GetSystemMetrics.argtypes = (ctypes.c_int,)
    _user32.LoadImageW.restype  = ctypes.c_void_p
    _user32.LoadImageW.argtypes = (ctypes.c_void_p, ctypes.c_wchar_p,
                                   ctypes.c_uint, ctypes.c_int, ctypes.c_int,
                                   ctypes.c_uint)
    _user32.SendMessageW.restype  = ctypes.c_void_p
    _user32.SendMessageW.argtypes = (wintypes.HWND, ctypes.c_uint,
                                     ctypes.c_void_p, ctypes.c_void_p)

    # 64-bit-safe long-ptr accessors (fall back to the 32-bit names on Win32)
    _get_long = getattr(_user32, "GetWindowLongPtrW", _user32.GetWindowLongW)
    _set_long = getattr(_user32, "SetWindowLongPtrW", _user32.SetWindowLongW)
    _set_class_long = getattr(_user32, "SetClassLongPtrW", _user32.SetClassLongW)
    for _f in (_get_long, _set_class_long):
        _f.restype = ctypes.c_void_p
    _get_long.argtypes = (wintypes.HWND, ctypes.c_int)
    _set_long.restype  = ctypes.c_void_p
    _set_long.argtypes = (wintypes.HWND, ctypes.c_int, ctypes.c_void_p)
    _set_class_long.argtypes = (wintypes.HWND, ctypes.c_int, ctypes.c_void_p)

    def _root_hwnd(hwnd):
        return _user32.GetAncestor(hwnd, _GA_ROOT)

    def make_taskbar_window(hwnd):
        """Give a top-level window a taskbar button (clear TOOLWINDOW, set
        APPWINDOW) so the Virtual Desktop Manager tracks it and keeps it on a
        single desktop. The window is briefly hidden/reshown so the shell picks
        up the change."""
        h  = _root_hwnd(hwnd)
        ex = (int(_get_long(h, _GWL_EXSTYLE) or 0)
              & ~_WS_EX_TOOLWINDOW) | _WS_EX_APPWINDOW
        _user32.ShowWindow(h, _SW_HIDE)
        _set_long(h, _GWL_EXSTYLE, ex)
        _user32.ShowWindow(h, _SW_SHOWNA)

    def set_window_owner(hwnd, owner_hwnd):
        """Make `hwnd` an owned window of `owner_hwnd`, so it has no taskbar
        button of its own and follows the owner across virtual desktops."""
        _set_long(_root_hwnd(hwnd), _GWLP_HWNDPARENT, _root_hwnd(owner_hwnd))

    def set_window_icon(hwnd, ico_path):
        """Load `ico_path` and attach it as the window's big + small icons (and
        class icons), so the taskbar button shows it. Tk's iconbitmap sets the
        window icon but the taskbar button created by WS_EX_APPWINDOW reads the
        class icon, so push both explicitly."""
        h = _root_hwnd(hwnd)
        big = _user32.LoadImageW(None, ico_path, _IMAGE_ICON,
                                 _user32.GetSystemMetrics(_SM_CXICON),
                                 _user32.GetSystemMetrics(_SM_CYICON),
                                 _LR_LOADFROMFILE)
        small = _user32.LoadImageW(None, ico_path, _IMAGE_ICON,
                                   _user32.GetSystemMetrics(_SM_CXSMICON),
                                   _user32.GetSystemMetrics(_SM_CYSMICON),
                                   _LR_LOADFROMFILE)
        if big:
            _user32.SendMessageW(h, _WM_SETICON, _ICON_BIG, big)
            _set_class_long(h, _GCLP_HICON, big)
        if small:
            _user32.SendMessageW(h, _WM_SETICON, _ICON_SMALL, small)
            _set_class_long(h, _GCLP_HICONSM, small)

    # Run via python.exe the taskbar button inherits the interpreter's
    # AppUserModelID (and its icon), ignoring our per-window icon — so give the
    # process its own identity. Skip this for the frozen .exe: there the exe is
    # already its own identity with an embedded icon, and an explicit AppID with
    # no relaunch/icon info would break the pinned-to-taskbar icon. Must happen
    # before the first window appears.
    if not getattr(sys, "frozen", False):
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "com_monitor.widget")
        except Exception:                                   # noqa: BLE001
            pass
else:
    def make_taskbar_window(hwnd):
        pass

    def set_window_owner(hwnd, owner_hwnd):
        pass

    def set_window_icon(hwnd, ico_path):
        pass

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
C_WAIT     = "#e6a500"   # amber 'waiting' status (we want to open, port is locked)

FONT         = ("Consolas", 10)
FONT_BOLD    = ("Consolas", 10, "bold")
FONT_IT      = ("Consolas", 10, "italic")
FONT_BOLD_IT = ("Consolas", 10, "bold italic")
FONT_HDR     = ("Consolas", 9, "bold")

ALPHA_OPAQUE = 0.96
FADE_RAMP_S  = 2.5       # seconds for chrome to fade to fully transparent
FADE_TICK_MS = 40        # animation tick while the fade ramp is running

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
# Anchor settings.json next to the program. In a PyInstaller --onefile build
# __file__ points at the temp _MEIxxxx extraction dir (wiped on exit), so use the
# directory of the actual .exe instead; otherwise use the script's own folder.
if getattr(sys, "frozen", False):
    _APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    _APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(_APP_DIR, "com_monitor.json")


def _resource_path(name: str) -> str:
    """Path to a bundled read-only resource (e.g. icon.ico). In a PyInstaller
    --onefile build data files live in the temp extraction dir (sys._MEIPASS);
    from source they sit next to this script."""
    base = getattr(sys, "_MEIPASS", _APP_DIR)
    return os.path.join(base, name)

DEFAULT_SETTINGS = {
    "highlight_duration_s":     20.0,         # row-flash fade duration
    "show_removed_enable":      True,         # show disconnected ports at all
    "show_removed_s":           60,           # gated: 0 = keep forever this session, else seconds before purge
    "always_on_top":            True,         # pin the window above all others
    "always_on_top_timeout_s":  0,            # 0 = never drop; else drop after idle
    "show_on_all_desktops":     True,         # Windows: show on every virtual desktop
                                              # (off = confine to one, adds a taskbar button)
    "interaction_timeout_s":    2,            # seconds with no mouse/focus before fading
    "move_to_back_enable":      False,        # push window to back of Z-order on idle
    "move_to_back_timeout_s":   300,
    "normal_alpha":             ALPHA_OPAQUE, # window opacity (rows keep this)
    "window_x":                 None,         # last-known position
    "window_y":                 None,
    "terminal_size":            "640x400",    # default terminal WxH (per-device geometry overrides)
    "terminal_always_on_top":   True,         # terminal pin — independent of main window
    "terminal_alpha":           ALPHA_OPAQUE, # terminal window opacity
    # per-device config keyed by "VID:PID:SERIAL"; each entry may carry a
    # "name" and/or serial settings — either alone is fine.
    "device_settings":          {},
}

# serial/terminal portion of a device entry (the "name" key lives alongside
# these but is not a serial default).
DEFAULT_PORT_SETTINGS = {
    "baud":        115200,
    "data_bits":   8,
    "parity":      "N",                       # N/E/O/M/S
    "stop_bits":   1,                         # 1, 1.5, 2
    "line_ending": "\r\n",                    # "\r\n" | "\n" | "\r" | ""
    "reconnect_on_unplug": True,              # keep terminal open & reconnect when device returns
    "signals":     "controls",                # title-bar modem lines: "none" | "controls" | "full"
    "watch_enable": False,                    # master toggle for the file watcher
    "watch_file":  "",                        # path to monitor (also disabled when empty)
    "watch_reconnect_s": 5.0,                 # seconds to release the port after a change
}

TERM_POLL_MS  = 50                            # serial-read poll cadence
TERM_WATCH_POLL_MS = 100                      # file-watch poll cadence
TERM_BORDER   = 4                             # px resize-zone around the terminal
TERM_MIN_W    = 280                           # minimum terminal width
TERM_MIN_H    = 160                           # minimum terminal height

# ESP/Arduino auto-reset timing (mirrors esptool). UART = external USB-serial
# bridge (CH340/CP2102/…); USB = native USB-Serial-JTAG (ESP32-S2/S3/C3/…).
ESP_RESET_DELAY_S     = 0.1                    # EN-low / inter-step hold
ESP_BOOT_DELAY_S      = 0.05                   # GPIO0 hold after release (UART boot)
ESP_USB_RESET_DELAY_S = 0.2                    # longer pulses for native-USB reset


def load_settings() -> dict:
    # shallow-copy each default value so callers can mutate nested dicts
    # (e.g. device_settings) without bleeding into DEFAULT_SETTINGS.
    merged = {k: (dict(v) if isinstance(v, dict) else v)
              for k, v in DEFAULT_SETTINGS.items()}
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return merged
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


# ── per-device record ─────────────────────────────────────────────────────────
class Device:
    """One COM port currently or recently seen. Bundles its identity/data
    (from list_ports), transient UI state (first-seen / flash / gone timers),
    and connection state. Persisted config (name + serial settings) lives in
    settings["device_settings"] keyed by VID:PID:Serial and is reached through
    the owning ComMonitor — this object is the single access point for it."""

    def __init__(self, app: "ComMonitor", dev_id: str, info: dict, now: float):
        self.app  = app
        self.id   = dev_id                  # e.g. "COM7"
        self.info = info                    # vid/pid/description/serial_number/location
        # transient UI state
        self.first_seen     = now
        self.flash_start: "float | None"    = None   # None = not flashing
        self.disappeared_at: "float | None" = None   # None = present
        # connection state, authoritative while a Terminal is bound to us
        self.conn_state = "disconnected"    # "connected" | "waiting" | "disconnected"
        self.terminal: "Terminal | None" = None   # this device's open terminal, if any
        self.wanted = False                 # queued to open a terminal once free

    # ── identity / data ──
    @property
    def present(self) -> bool:        return self.disappeared_at is None
    @property
    def vid(self):                    return self.info.get("vid")
    @property
    def pid(self):                    return self.info.get("pid")
    @property
    def serial_number(self) -> str:   return self.info.get("serial_number", "")
    @property
    def location(self) -> str:        return self.info.get("location", "")
    @property
    def description(self) -> str:     return self.info.get("description", "")
    @property
    def custom_key(self):
        """Settings key for this device's config. Serial is included verbatim —
        an empty serial yields "VID:PID:" which still uniquely identifies the
        "no-serial" variant of that VID:PID. Location is intentionally ignored."""
        if self.vid is None or self.pid is None:
            return None
        return f"{self.vid:04X}:{self.pid:04X}:{self.serial_number or ''}"

    # ── persisted config (this device's entry in settings["device_settings"]) ──
    def _config(self) -> dict:
        """The stored config dict for this device, or {} if none/unkeyable.
        Returns the live dict from settings so reads see edits immediately."""
        k = self.custom_key
        return self.app.settings.get("device_settings", {}).get(k, {}) if k else {}

    @property
    def name(self) -> str:
        return self._config().get("name", "")

    def port_settings(self) -> dict:
        """DEFAULT_PORT_SETTINGS merged with this device's serial overrides."""
        s = dict(DEFAULT_PORT_SETTINGS)
        s.update(self._config())
        for k in ("name", "geometry"):                  # non-serial config keys
            s.pop(k, None)
        return s

    @property
    def geometry(self) -> "str | None":
        """Remembered terminal geometry "WxH+X+Y" for this device, or None."""
        return self._config().get("geometry")

    def save_geometry(self, geo: str):
        self.update_config(geometry=geo)

    def update_config(self, **fields):
        """Merge fields into this device's stored config, pruning empties.
        Pass name="" (or any blank value) to drop that key."""
        k = self.custom_key
        if k is None:
            return
        store = self.app.settings.setdefault("device_settings", {})
        entry = store.setdefault(k, {})
        for kk, vv in fields.items():
            if vv == "" or vv is None:
                entry.pop(kk, None)
            else:
                entry[kk] = vv
        if not entry:                                   # nothing left → drop entry
            store.pop(k, None)

    def set_name(self, name: str):
        self.update_config(name=name)


# ── main window ───────────────────────────────────────────────────────────────
class ComMonitor(tk.Tk):
    def __init__(self):
        super().__init__()
        self.settings = load_settings()

        # app/window icon — drives the taskbar button when confined to a desktop.
        # default=… applies it to every toplevel (terminals, dialogs) too.
        try:
            self.iconbitmap(default=_resource_path("icon.ico"))
        except tk.TclError:
            pass

        self.overrideredirect(True)
        self.attributes("-topmost", self.settings["always_on_top"])
        self.attributes("-alpha", self.settings["normal_alpha"])
        self.configure(bg=C_BG)

        self.devices: dict[str, Device] = {}          # COM id → Device (present or recently gone)
        self._initialized  = False                    # skip flash for ports present at startup
        self._row_widgets: list[list[tk.Widget]] = []
        self._drag_ox = self._drag_oy = 0

        # interaction-driven fade + on-change top-most state
        self._known_ports: set[str] = set()
        self._last_interaction = time.time()
        self._top_until        = 0.0          # window stays topmost while now < this
        self._top_active       = self.settings["always_on_top"]
        self._bg_active        = False
        self._last_chrome_alpha = -1.0
        self._editing_key      = None        # custom_key currently being edited
        self._edit_entry       = None        # tk.Entry overlaying a serial cell
        self._edit_lbl         = None        # serial Label hidden during edit
        self._edit_desc_lbl    = None        # description Label hidden during edit
        # per-device terminals + open-when-free intent live on each Device
        self._flash_decay_k = self._compute_flash_decay_k()

        # font handles used to pre-compute matching column widths in both grids
        self._f_row = tkfont.Font(family="Consolas", size=10)
        self._f_hdr = tkfont.Font(family="Consolas", size=9, weight="bold")

        self._build_ui()
        self.bind("<Configure>", self._sync_rows_geometry)
        self._refresh()
        self._animate_fade()
        self._restore_window_position()
        self._sync_rows_geometry()
        self._rows_win.deiconify()                      # reveal once placed
        self._apply_desktop_confinement()

    # ── layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Chrome window (self) holds the title bar and the column-header strip.
        # Its -alpha animates from normal → 0 over the fade ramp.
        bar = tk.Frame(self, bg=C_HDR, cursor="fleur")
        bar.pack(fill=tk.X)
        bar.bind("<ButtonPress-1>", self._drag_start)
        bar.bind("<B1-Motion>",     self._drag_move)

        tk.Label(bar, text="  USB COM Monitor",
                 bg=C_HDR, fg="#666666", font=("Segoe UI", 8),
                 pady=4).pack(side=tk.LEFT)

        tk.Button(bar, text=" ✕ ", bg=C_HDR, fg="#666666", bd=0,
                  relief="flat", font=("Segoe UI", 8),
                  activebackground="#c0392b", activeforeground="white",
                  command=self.destroy).pack(side=tk.RIGHT)

        tk.Button(bar, text=" ⚙ ", bg=C_HDR, fg="#666666", bd=0,
                  relief="flat", font=("Segoe UI", 8),
                  activebackground=C_HDR, activeforeground="#aaaaaa",
                  command=self._open_settings).pack(side=tk.RIGHT)

        self._hdr_grid = tk.Frame(self, bg=C_BG)
        self._hdr_grid.pack(fill=tk.X, padx=2, pady=(3, 0))

        for c, (name, anchor, stretch) in enumerate(COLS):
            tk.Label(self._hdr_grid, text=name, bg=C_BG, fg=C_HEAD,
                     font=FONT_HDR, anchor=anchor, padx=3
                     ).grid(row=0, column=c, sticky="ew")
            if stretch:
                self._hdr_grid.columnconfigure(c, weight=1)

        tk.Frame(self._hdr_grid, bg="#333333", height=1).grid(
            row=1, column=0, columnspan=len(COLS), sticky="ew", pady=(1, 0))

        # Rows window — borderless Toplevel that always tracks the chrome
        # window's position. Holds the data grid only; its -alpha stays at
        # normal_alpha so rows never fade.
        self._rows_win = tk.Toplevel(self, bg=C_BG)
        self._rows_win.withdraw()                       # hidden until positioned
        self._rows_win.overrideredirect(True)
        self._rows_win.attributes("-topmost", self.settings["always_on_top"])
        self._rows_win.attributes("-alpha",   self.settings["normal_alpha"])

        self._grid = tk.Frame(self._rows_win, bg=C_BG)
        self._grid.pack(fill=tk.BOTH, expand=True, padx=2, pady=(0, 2))

        for c, (_, _, stretch) in enumerate(COLS):
            if stretch:
                self._grid.columnconfigure(c, weight=1)

        self._empty_lbl = tk.Label(
            self._grid, text="(no COM ports detected)",
            bg=C_BG, fg="#3a3a3a", font=FONT, anchor="w", padx=3)

        # dragging anywhere in the rows area also moves the pair
        self._grid.bind("<ButtonPress-1>", self._drag_start)
        self._grid.bind("<B1-Motion>",     self._drag_move)

        for w in (self, self._rows_win):
            w.bind("<Escape>",    lambda _: self.destroy())
            w.bind("<Control-q>", lambda _: self.destroy())
            w.bind("<Button>",    self._on_interact)
            w.bind("<Enter>",     self._on_interact)
            w.bind("<FocusIn>",   self._on_interact)
            w.bind("<Motion>",    self._on_interact)

    # ── transparency ──────────────────────────────────────────────────────────
    def _update_alpha(self):
        normal    = self.settings.get("normal_alpha", ALPHA_OPAQUE)
        threshold = self.settings["interaction_timeout_s"]
        idle      = time.time() - self._last_interaction
        f = min(1.0, (idle - threshold) / FADE_RAMP_S) if idle > threshold else 0.0
        # rows window holds its opacity; only the chrome window fades
        try:
            self._rows_win.attributes("-alpha", normal)
        except tk.TclError:
            pass
        chrome_alpha = normal * (1.0 - f)
        if chrome_alpha != self._last_chrome_alpha:
            self.attributes("-alpha", chrome_alpha)
            self._last_chrome_alpha = chrome_alpha

    def _animate_fade(self):
        # poll on the fast tick so a motionless cursor resting over the window
        # holds it solid (the 500 ms refresh poll alone lets the alpha dim
        # between ticks).
        if self._pointer_holds_visible():
            self._last_interaction = time.time()
        self._update_alpha()
        self.after(FADE_TICK_MS, self._animate_fade)

    def _pointer_holds_visible(self) -> bool:
        """True while the pointer should keep the window awake: over the rows
        window, or over the chrome window *while it is still visible*, or a
        widget is focused. Once the chrome has faded to alpha 0 the pointer over
        its (invisible) strip is ignored, so it stays hidden until the pointer
        reaches the rows."""
        try:
            px, py = self.winfo_pointerxy()
            # rows window always holds it awake
            r = self._rows_win
            rx, ry = r.winfo_rootx(), r.winfo_rooty()
            rw, rh = r.winfo_width(), r.winfo_height()
            if rx <= px < rx + rw and ry <= py < ry + rh:
                return True
            # chrome window holds it only while still visible (alpha > 0)
            if self._last_chrome_alpha > 0.0:
                cx, cy = self.winfo_rootx(), self.winfo_rooty()
                cw, ch = self.winfo_width(), self.winfo_height()
                if cx <= px < cx + cw and cy <= py < cy + ch:
                    return True
            return self._rows_win.focus_displayof() is not None
        except tk.TclError:
            return False

    # ── two-window sync ───────────────────────────────────────────────────────
    def _sync_rows_geometry(self, _event=None):
        """Keep the rows Toplevel glued directly under the chrome window."""
        try:
            self.update_idletasks()
            x = self.winfo_rootx()
            y = self.winfo_rooty() + self.winfo_height()
            self._rows_win.geometry(f"+{x}+{y}")        # position only — width auto
        except tk.TclError:
            pass

    def _sync_column_widths(self):
        """Pre-compute the per-column pixel width needed for the widest text in
        either grid (header text or any data cell) and apply it as `minsize` to
        both grids. With both grids on the same minsize, columns line up and
        both windows end up the same total width."""
        cell_pad = 6                                   # padx=3 on each side
        col_w = [0] * len(COLS)

        for c, (name, _, _) in enumerate(COLS):
            col_w[c] = self._f_hdr.measure(name) + cell_pad

        for row in self._row_widgets:
            for lbl in row:
                try:
                    info = lbl.grid_info()
                    if int(info.get("columnspan", 1)) > 1:
                        continue           # spanned cell doesn't constrain its column
                    col  = int(info["column"])
                    text = lbl.cget("text")
                except (tk.TclError, KeyError, ValueError):
                    continue
                w = self._f_row.measure(text) + cell_pad
                if w > col_w[col]:
                    col_w[col] = w

        for c, w in enumerate(col_w):
            self._hdr_grid.columnconfigure(c, minsize=w)
            self._grid.columnconfigure(c, minsize=w)

        # force both windows to the same overall width so they read as one
        self.update_idletasks()
        self._rows_win.update_idletasks()
        target_w = max(self.winfo_reqwidth(), self._rows_win.winfo_reqwidth())
        x = self.winfo_rootx()
        y = self.winfo_rooty()
        self.geometry(f"{target_w}x{self.winfo_reqheight()}+{x}+{y}")
        self._sync_rows_geometry()
        self._rows_win.geometry(
            f"{target_w}x{self._rows_win.winfo_reqheight()}"
            f"+{x}+{y + self.winfo_reqheight()}")

    def _on_interact(self, _event=None):
        """Immediate wake — restart the interaction timer and restore visibility."""
        self._last_interaction = time.time()
        if self._bg_active:
            self._lift_pair()
            self._bg_active = False
        self._update_alpha()

    # ── paired Z-order helpers ───────────────────────────────────────────────
    def _terminals(self):
        """Every currently-open Terminal (at most one per device)."""
        return [d.terminal for d in self.devices.values() if d.terminal is not None]

    def _update_terminal_alpha(self):
        """Apply terminal_alpha to every open terminal (live-preview + save)."""
        a = float(self.settings.get("terminal_alpha", ALPHA_OPAQUE))
        for t in self._terminals():
            try:
                t.attributes("-alpha", a)
            except tk.TclError:
                pass

    def _lift_pair(self):
        """Bring the windows to the top of the stack while preserving their
        relative order (chrome below, rows above, terminals on top)."""
        try:
            self.lift()
            self._rows_win.lift()
            for t in self._terminals():
                t.lift()
        except tk.TclError:
            pass

    def _lower_pair(self):
        try:
            for t in self._terminals():
                t.lower()
            self._rows_win.lower()
            self.lower()
        except tk.TclError:
            pass

    def _set_topmost_pair(self, on: bool):
        # Note: the terminal keeps its own topmost state (toggled from its
        # title bar) and is intentionally not pinned/unpinned here.
        try:
            self.attributes("-topmost", on)
            self._rows_win.attributes("-topmost", on)
        except tk.TclError:
            pass

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

    def _on_status_click(self, device: Device):
        """Click on a row's status cell. Each device has its own terminal:
        surface it if open, otherwise open now (if free) or queue the wait.
        If the terminal exists but is user-disconnected, the click also asks
        it to reconnect — the row labelled 'free' should match the action."""
        if device.terminal is not None:
            device.terminal.lift()
            device.terminal.focus_set()
            if device.conn_state == "disconnected":
                device.terminal.reconnect()
            return
        if device.wanted:
            device.wanted = False                       # cancel wait
            return
        # free and present right now → open immediately
        if device.present and not is_open(device.id):
            self._open_terminal(device)
            return
        # OPEN (held elsewhere) or gone → queue and grab it as soon as it's
        # present and free again
        device.wanted = True

    def _open_terminal(self, device: Device):
        device.wanted = False
        try:
            device.terminal = Terminal(self, device)
            device.terminal.lift()
            self._confine_window(device.terminal)
        except Exception as e:                          # noqa: BLE001
            print(f"Failed to open terminal for {device.id}: {e}")
            device.terminal = None


    # ── inline edit of custom serial name ─────────────────────────────────────
    def _begin_serial_edit(self, lbl, device: Device, desc_lbl=None):
        """Overlay an Entry on the serial cell so the user can type a custom
        name keyed by VID:PID:Serial. While editing, row rebuilds are paused
        (see `_editing_key` short-circuit in `_refresh`)."""
        if self._editing_key is not None:
            return
        self._editing_key   = device.custom_key
        self._edit_lbl      = lbl
        self._edit_desc_lbl = desc_lbl

        info = lbl.grid_info()
        row, col = int(info["row"]), int(info["column"])
        row_bg   = lbl.cget("bg")
        current  = device.name

        lbl.grid_remove()
        if desc_lbl is not None:
            desc_lbl.grid_remove()

        entry = tk.Entry(self._grid, bg=row_bg, fg=C_SER, font=FONT_IT,
                         bd=0, relief="flat",
                         insertbackground=C_PORT,
                         highlightthickness=1,
                         highlightbackground=C_PORT,
                         highlightcolor=C_PORT)
        entry.insert(0, current)
        entry.grid(row=row, column=col, columnspan=2, sticky="ew")
        entry.focus_set()
        entry.select_range(0, tk.END)
        entry.icursor(tk.END)
        self._edit_entry = entry

        def commit(_e=None):
            if self._edit_entry is None:
                return "break"
            v = entry.get().strip()
            device.set_name(v)
            save_settings(self.settings)
            self._end_serial_edit()
            return "break"

        def cancel(_e=None):
            self._end_serial_edit()
            return "break"

        entry.bind("<Return>",   commit)
        entry.bind("<Escape>",   cancel)
        entry.bind("<FocusOut>", commit)

    def _end_serial_edit(self):
        if self._edit_entry is None:
            return
        try:
            self._edit_entry.destroy()
        except tk.TclError:
            pass
        self._edit_entry = None
        # restore the hidden labels — the next _refresh tick will rebuild the
        # row properly (italic + columnspan if a name was set, plain otherwise)
        for w in (self._edit_lbl, self._edit_desc_lbl):
            if w is not None:
                try:
                    w.grid()
                except tk.TclError:
                    pass
        self._edit_lbl      = None
        self._edit_desc_lbl = None
        self._editing_key   = None

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
        for t in self._terminals():
            t._close_quiet()
        super().destroy()

    def _apply_settings(self):
        """Re-apply settings to derived state after they've been edited."""
        self._flash_decay_k    = self._compute_flash_decay_k()
        self._last_interaction = time.time()
        self._set_topmost_pair(self.settings["always_on_top"])
        self._top_active = self.settings["always_on_top"]
        self._update_alpha()
        self._update_terminal_alpha()
        self._apply_desktop_confinement()

    # ── virtual-desktop confinement ─────────────────────────────────────────
    def _confine_to_desktop_enabled(self) -> bool:
        return (sys.platform == "win32"
                and not self.settings.get("show_on_all_desktops", True))

    def _apply_desktop_confinement(self):
        """If enabled (Windows), give the main strip a taskbar button so the
        Virtual Desktop Manager keeps it on a single desktop, and make the rows
        window + any open terminals owned by it so they follow along without
        adding buttons of their own. Default off → borderless, no taskbar
        button, shown on all desktops (unchanged). Takes effect on Save;
        turning it back off needs a restart."""
        if not self._confine_to_desktop_enabled():
            return
        make_taskbar_window(self.winfo_id())
        set_window_icon(self.winfo_id(), _resource_path("icon.ico"))
        set_window_owner(self._rows_win.winfo_id(), self.winfo_id())
        for t in self._terminals():
            try:
                set_window_owner(t.winfo_id(), self.winfo_id())
            except tk.TclError:
                pass

    def _confine_window(self, win):
        """Make a newly-opened window (a terminal) owned by the main strip so it
        shares its virtual desktop and adds no taskbar button of its own."""
        if not self._confine_to_desktop_enabled():
            return
        try:
            set_window_owner(win.winfo_id(), self.winfo_id())
        except tk.TclError:
            pass

    def _open_settings(self):
        SettingsDialog(self)

    # ── refresh ───────────────────────────────────────────────────────────────
    def _refresh(self):
        now   = time.time()
        ports = sorted(serial.tools.list_ports.comports(), key=lambda p: p.device)

        # update device records — active ports
        current = {p.device for p in ports}
        for p in ports:
            info = {
                "vid": p.vid, "pid": p.pid,
                "description":   p.description or "",
                "serial_number": p.serial_number or "",
                "location":      p.location or "",
            }
            d = self.devices.get(p.device)
            if d is None:                       # first time we've seen this port
                d = self.devices[p.device] = Device(self, p.device, info, now)
                if self._initialized:           # don't flash ports seen at startup
                    # shift back so first render lands at peak brightness, not t=0
                    d.flash_start = now - FLASH_ATTACK_S
            else:
                d.info = info
                if not d.present:               # port came back → re-flash as new
                    d.disappeared_at = None
                    d.flash_start = now - FLASH_ATTACK_S

        # mark newly-vanished ports (always tracked; show_removed only affects display)
        for d in self.devices.values():
            if d.id not in current and d.present:
                d.disappeared_at = now
                d.flash_start = None

        # expire/purge: drop records gone too long. Keep any device that still
        # has a terminal or a pending open, even once its row expires.
        # Semantics: checkbox off → drop immediately; checkbox on + 0 → keep for
        # the rest of the session; checkbox on + N → drop after N seconds.
        show_removed_en = self.settings["show_removed_enable"]
        show_removed    = self.settings["show_removed_s"]
        for dev_id, d in list(self.devices.items()):
            if d.present or d.terminal is not None or d.wanted:
                continue
            if not show_removed_en:
                expired = True
            elif show_removed == 0:
                expired = False
            else:
                expired = (d.disappeared_at is None
                           or (now - d.disappeared_at) > show_removed)
            if expired:
                del self.devices[dev_id]

        # ports queued to open: grab each as soon as it's present AND free
        for d in self.devices.values():
            if d.wanted and d.present and not is_open(d.id):
                self._open_terminal(d)

        # reconcile each open terminal's connection with its device's presence
        for d in self.devices.values():
            t = d.terminal
            if t is None:
                continue
            if t.state == "connected" and not d.present:
                # device unplugged
                t._close_serial()
                if t.reconnect_on_unplug():
                    t._set_state("waiting")
                    t._info("[device removed — waiting for it to return]")
                else:
                    t._close_quiet()
            elif t.state == "waiting" and d.present and not is_open(d.id):
                # device available again → reconnect (manual disconnect won't reach here)
                t.reconnect()

        # port changes lift the window pair and pin them on top for one
        # interaction timeout so the user actually notices the connect/disconnect.
        port_changed = self._initialized and current != self._known_ports
        if port_changed:
            self._top_until = now + self.settings["interaction_timeout_s"]
            self._lift_pair()
        self._known_ports = current

        # interaction poll: pointer resting over the window (or a focused
        # widget) holds the timer at zero. The chrome strip counts only while
        # still visible, so hovering it once invisible does not bring it back.
        if self._pointer_holds_visible():
            self._last_interaction = now

        idle = now - self._last_interaction

        # always-on-top with optional drop-after-idle timeout
        aot_on  = self.settings["always_on_top"]
        aot_to  = self.settings["always_on_top_timeout_s"]
        aot_live = aot_on and (aot_to == 0 or idle < aot_to)
        want_top = aot_live or now < self._top_until

        # push-to-back on idle (overridden by the port-change pin)
        mtb_on  = self.settings.get("move_to_back_enable", False)
        mtb_to  = self.settings.get("move_to_back_timeout_s", 0)
        want_back = (mtb_on and mtb_to > 0 and idle > mtb_to
                     and now >= self._top_until)

        if want_top and not self._top_active:
            self._set_topmost_pair(True)
            self._top_active = True
        elif not want_top and self._top_active:
            self._set_topmost_pair(False)
            self._top_active = False

        if want_back and not self._bg_active:
            self._lower_pair()
            self._bg_active = True
        elif not want_back and self._bg_active:
            self._lift_pair()
            self._bg_active = False

        self._update_alpha()

        self._initialized = True

        if self._editing_key is not None:
            # an inline-edit is in progress — leave the rows alone so the
            # Entry overlay stays put; non-row state (fade, Z-order, port
            # tracking) above has already updated.
            self.after(REFRESH_MS, self._refresh)
            return

        # rebuild rows
        for row in self._row_widgets:
            for w in row:
                w.destroy()
        self._row_widgets.clear()
        self._empty_lbl.grid_forget()

        # Hide removed ports if the checkbox is off; else show while either the
        # session-forever (0) flag is set or the per-port window hasn't elapsed.
        visible = [d for d in self.devices.values()
                   if d.present or (show_removed_en
                                    and (show_removed == 0
                                         or now - d.disappeared_at <= show_removed))]
        visible.sort(key=lambda d: d.id)
        if not visible:
            self._empty_lbl.grid(row=0, column=0, columnspan=len(COLS),
                                 sticky="w", pady=6)
        else:
            for i, d in enumerate(visible):
                gone = not d.present

                vid = f"{d.vid:04X}" if d.vid is not None else "----"
                pid = f"{d.pid:04X}" if d.pid is not None else "----"

                ser_loc = " / ".join(filter(None, [d.serial_number,
                                                   d.location])) or "—"
                desc = "" if d.description == d.id else d.description

                if gone:
                    age_s      = now - d.disappeared_at
                    brightness = _flash_brightness(age_s, self._flash_decay_k)
                    row_bg     = _blend(C_ROW_GONE, C_BG, 1.0 - brightness)
                    row_font   = FONT
                    dot_c      = C_GONE
                    age_c      = C_DIM
                    if d.wanted or d.conn_state == "waiting":
                        stat_t, stat_c = "waiting", C_WAIT
                    else:
                        stat_t, stat_c = "gone", C_GONE
                    port_c     = C_DIM
                    vid_c      = C_DIM
                    ser_c      = C_DIM
                    desc_c     = C_DIM
                else:
                    age_s = now - d.first_seen
                    cs    = d.conn_state
                    # our terminal owns the port only while actually connected;
                    # waiting/disconnected releases it so the row reflects that
                    occupied = (cs == "connected"
                                or (cs == "disconnected" and is_open(d.id)))
                    waiting  = (cs == "waiting"
                                or (cs == "disconnected" and d.wanted))

                    flash_t = now - d.flash_start if d.flash_start is not None else None
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
                    if waiting:
                        stat_t, stat_c = "waiting", C_WAIT
                    elif occupied:
                        stat_t, stat_c = "OPEN", C_OPEN
                    else:
                        stat_t, stat_c = "free", C_FREE
                    port_c = C_PORT
                    vid_c  = C_VIDPID
                    ser_c  = C_SER
                    desc_c = C_DESC

                custom_name = d.name

                # (column, columnspan, label-kwargs)
                cells = [
                    (0, 1, dict(text=d.id,                 fg=port_c, anchor="w")),
                    (1, 1, dict(text=f"{vid}:{pid}",       fg=vid_c,  anchor="w")),
                    (2, 1, dict(text=self._age_str(age_s), fg=age_c,  anchor="e")),
                    (3, 1, dict(text="●",                  fg=dot_c,  anchor="center")),
                    (4, 1, dict(text=stat_t,               fg=stat_c, anchor="w")),
                ]
                if custom_name:
                    it_font = FONT_BOLD_IT if row_font is FONT_BOLD else FONT_IT
                    cells.append((5, 2,
                        dict(text=custom_name, fg=ser_c, anchor="w", font=it_font)))
                else:
                    cells.append((5, 1, dict(text=ser_loc, fg=ser_c,  anchor="w")))
                    cells.append((6, 1, dict(text=desc,    fg=desc_c, anchor="w")))

                row_w = []
                serial_lbl = None
                desc_lbl   = None
                status_lbl = None
                for col, span, kw in cells:
                    kw.setdefault("font", row_font)
                    lbl = tk.Label(self._grid, bg=row_bg, padx=3, pady=1, **kw)
                    lbl.grid(row=i, column=col, columnspan=span, sticky="ew")
                    row_w.append(lbl)
                    if col == 4:
                        status_lbl = lbl
                    elif col == 5:
                        serial_lbl = lbl
                    elif col == 6:
                        desc_lbl = lbl

                if serial_lbl is not None and d.custom_key is not None:
                    serial_lbl.bind("<Double-Button-1>",
                        lambda _e, dd=d, sw=serial_lbl, dw=desc_lbl:
                            self._begin_serial_edit(sw, dd, dw))

                # status cell drives terminal open / wait / cancel — clickable
                # in every state (free → open now; OPEN/gone → queue the wait)
                if status_lbl is not None:
                    status_lbl.configure(cursor="hand2")
                    status_lbl.bind("<Button-1>",
                        lambda _e, dd=d: self._on_status_click(dd))

                self._row_widgets.append(row_w)

        self._sync_column_widths()
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

        # Show disconnected ports (checkbox), with an indented timeout below.
        self.var_sr_en = tk.BooleanVar(value=s["show_removed_enable"])
        tk.Checkbutton(frm, text="Show disconnected ports",
                       variable=self.var_sr_en,
                       bg=C_BG, fg=C_DESC, selectcolor=C_HDR,
                       activebackground=C_BG, activeforeground=C_DESC,
                       font=FONT, anchor="w"
                       ).grid(row=1, column=0, columnspan=2, sticky="w", **pad)
        tk.Label(frm, text="     ↳ remove after (s, 0 = never):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=2, column=0, sticky="w", **pad)
        self.var_sr = tk.IntVar(value=s["show_removed_s"])
        tk.Spinbox(frm, from_=0, to=86400, increment=10, textvariable=self.var_sr,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=2, column=1, sticky="w", **pad)

        # Always on top
        self.var_aot = tk.BooleanVar(value=s["always_on_top"])
        tk.Checkbutton(frm, text="Always on top",
                       variable=self.var_aot,
                       bg=C_BG, fg=C_DESC, selectcolor=C_HDR,
                       activebackground=C_BG, activeforeground=C_DESC,
                       font=FONT, anchor="w"
                       ).grid(row=3, column=0, columnspan=2, sticky="w", **pad)
        tk.Label(frm, text="     ↳ drop after idle (s, 0 = never):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=4, column=0, sticky="w", **pad)
        self.var_aot_to = tk.IntVar(value=s["always_on_top_timeout_s"])
        tk.Spinbox(frm, from_=0, to=86400, increment=10, textvariable=self.var_aot_to,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=4, column=1, sticky="w", **pad)

        # Move to background on idle
        self.var_mtb = tk.BooleanVar(value=s["move_to_back_enable"])
        tk.Checkbutton(frm, text="Move to background on idle",
                       variable=self.var_mtb,
                       bg=C_BG, fg=C_DESC, selectcolor=C_HDR,
                       activebackground=C_BG, activeforeground=C_DESC,
                       font=FONT, anchor="w"
                       ).grid(row=5, column=0, columnspan=2, sticky="w", **pad)
        tk.Label(frm, text="     ↳ after (s):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=6, column=0, sticky="w", **pad)
        self.var_mtb_to = tk.IntVar(value=s["move_to_back_timeout_s"])
        tk.Spinbox(frm, from_=1, to=86400, increment=10, textvariable=self.var_mtb_to,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=6, column=1, sticky="w", **pad)

        # Interaction timeout (seconds before fade begins; 0 = fade at once)
        tk.Label(frm, text="Fade after no interaction (s, 0 = immediate):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=7, column=0, sticky="w", **pad)
        self.var_it = tk.IntVar(value=s["interaction_timeout_s"])
        tk.Spinbox(frm, from_=0, to=3600, increment=1, textvariable=self.var_it,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=7, column=1, sticky="w", **pad)

        # Window transparency (rows keep this opacity; live preview)
        tk.Label(frm, text="Window transparency:",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=8, column=0, sticky="w", **pad)
        self.var_an = tk.DoubleVar(value=s["normal_alpha"])
        tk.Scale(frm, from_=0.4, to=1.0, resolution=0.01,
                 orient=tk.HORIZONTAL, variable=self.var_an,
                 command=lambda v: self._preview_alpha("normal_alpha", v),
                 bg=C_BG, fg=C_DESC, troughcolor=C_HDR,
                 highlightthickness=0, bd=0, length=160,
                 font=FONT, activebackground=C_PORT,
                 ).grid(row=8, column=1, sticky="w", **pad)

        # Terminal transparency (applies to every open terminal; live preview)
        tk.Label(frm, text="Terminal transparency:",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=9, column=0, sticky="w", **pad)
        self.var_tan = tk.DoubleVar(value=s.get("terminal_alpha", ALPHA_OPAQUE))
        tk.Scale(frm, from_=0.4, to=1.0, resolution=0.01,
                 orient=tk.HORIZONTAL, variable=self.var_tan,
                 command=self._preview_terminal_alpha,
                 bg=C_BG, fg=C_DESC, troughcolor=C_HDR,
                 highlightthickness=0, bd=0, length=160,
                 font=FONT, activebackground=C_PORT,
                 ).grid(row=9, column=1, sticky="w", **pad)

        # Show on all virtual desktops (Windows). Unchecking confines the widget
        # to one desktop, which adds a single taskbar button. Takes effect on
        # Save; re-enabling "all desktops" needs a restart.
        self.var_ad = tk.BooleanVar(value=s.get("show_on_all_desktops", True))
        tk.Checkbutton(frm, text="Show on all desktops (needs restart)",
                       variable=self.var_ad,
                       bg=C_BG, fg=C_DESC, selectcolor=C_HDR,
                       activebackground=C_BG, activeforeground=C_DESC,
                       font=FONT, anchor="w"
                       ).grid(row=10, column=0, columnspan=2, sticky="w", **pad)

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

    def _preview_terminal_alpha(self, val):
        try:
            self.parent.settings["terminal_alpha"] = float(val)
        except (TypeError, ValueError):
            return
        self.parent._update_terminal_alpha()

    def _cancel(self):
        if not self._saved:
            self.parent.settings.clear()
            self.parent.settings.update(self._snapshot)
            self.parent._update_alpha()
            self.parent._update_terminal_alpha()
        self.destroy()

    def _save(self):
        s = self.parent.settings
        try:
            v = float(self.var_hl.get())
            if v > 0:
                s["highlight_duration_s"] = v
        except (tk.TclError, ValueError):
            pass
        s["show_removed_enable"] = bool(self.var_sr_en.get())
        try:
            v = int(self.var_sr.get())
            if v >= 0:
                s["show_removed_s"] = v
        except (tk.TclError, ValueError):
            pass
        s["always_on_top"] = bool(self.var_aot.get())
        try:
            v = int(self.var_aot_to.get())
            if v >= 0:
                s["always_on_top_timeout_s"] = v
        except (tk.TclError, ValueError):
            pass
        s["move_to_back_enable"] = bool(self.var_mtb.get())
        try:
            v = int(self.var_mtb_to.get())
            if v > 0:
                s["move_to_back_timeout_s"] = v
        except (tk.TclError, ValueError):
            pass
        s["show_on_all_desktops"] = bool(self.var_ad.get())
        try:
            v = int(self.var_it.get())
            if v >= 0:
                s["interaction_timeout_s"] = v
        except (tk.TclError, ValueError):
            pass
        try:
            s["normal_alpha"] = float(self.var_an.get())
        except (tk.TclError, ValueError):
            pass
        try:
            s["terminal_alpha"] = float(self.var_tan.get())
        except (tk.TclError, ValueError):
            pass
        save_settings(s)
        self.parent._apply_settings()
        self._saved = True
        self.destroy()


# ── terminal ──────────────────────────────────────────────────────────────────
class Terminal(tk.Toplevel):
    def __init__(self, parent: ComMonitor, device: "Device"):
        super().__init__(parent)
        self.parent   = parent
        self.dev      = device
        self.serial: "serial.Serial | None" = None
        self._poll_id = None
        self._closing = False
        self._conn_btn = None
        # modem-control state: RTS/DTR are outputs we drive, CTS/DSR/DCD are
        # inputs we read. Widget refs are filled by _build_signals. ESP auto-reset
        # modes idle the lines de-asserted so merely opening or closing the port
        # doesn't pulse the board into reset (a 1→0 drop on close is what the
        # USB-Serial-JTAG reads as a reset); other modes keep the conventional
        # asserted-on-open behaviour.
        esp = self._signals_mode() in ("esptool", "esptool_usb")
        self._rts = self._dtr = not esp
        self._sig_frame = None
        self._rts_btn = self._dtr_btn = None
        self._cts_lbl = self._dsr_lbl = self._dcd_lbl = None
        # file-watch state. baseline is (exists, mtime, size) or None when no
        # file is being watched / no observation yet; _watch_holding marks the
        # release window after a detected change so the ComMonitor auto-reconnect
        # machinery doesn't grab the port back during the flash.
        self._watch_baseline = None
        self._watch_poll_id = None
        self._watch_release_id = None
        self._watch_holding = False

        self.title(f"Terminal — {device.id}")
        self.configure(bg=C_HDR)                        # exposed margin = subtle border
        self.overrideredirect(True)
        self._topmost = bool(parent.settings.get("terminal_always_on_top", True))
        self.attributes("-topmost", self._topmost)
        self.attributes("-alpha",
                        float(parent.settings.get("terminal_alpha", ALPHA_OPAQUE)))
        saved_geo = device.geometry
        self.geometry(saved_geo or parent.settings.get("terminal_size", "640x400"))

        # resize state
        self._resize_edge = None
        self._rs = (0, 0, 0, 0, 0, 0)                    # sx, sy, x, y, w, h
        self._drag_ox = self._drag_oy = 0

        self._build_ui()
        self.update_idletasks()
        if not saved_geo:                                # first open → place above main
            self._position_above_main()
        self._set_state("connected" if self._open_serial() else "waiting")
        self._schedule_poll()
        self._schedule_watch_poll()

        self.bind("<Configure>", self._on_configure)
        self.focus_force()                              # override-redirect needs a nudge
        self.inp.focus_set()

    def _build_ui(self):
        B = TERM_BORDER
        # inner panel inset by the border on every side; the exposed C_HEAD
        # margin around it doubles as the resize zone.
        inner = tk.Frame(self, bg=C_BG)
        inner.place(x=B, y=B, relwidth=1.0, width=-2 * B,
                    relheight=1.0, height=-2 * B)

        # ── resize grips (placed directly on the toplevel, around `inner`) ──
        grips = (
            ("n",  dict(x=B, y=0,  relwidth=1.0, width=-2 * B, height=B),  "sb_v_double_arrow"),
            ("s",  dict(x=B, rely=1.0, y=-B, relwidth=1.0, width=-2 * B, height=B), "sb_v_double_arrow"),
            ("w",  dict(x=0, y=B, relheight=1.0, height=-2 * B, width=B),  "sb_h_double_arrow"),
            ("e",  dict(relx=1.0, x=-B, y=B, relheight=1.0, height=-2 * B, width=B), "sb_h_double_arrow"),
            ("nw", dict(x=0, y=0, width=B, height=B), "size_nw_se"),
            ("se", dict(relx=1.0, x=-B, rely=1.0, y=-B, width=B, height=B), "size_nw_se"),
            ("ne", dict(relx=1.0, x=-B, y=0, width=B, height=B), "size_ne_sw"),
            ("sw", dict(x=0, rely=1.0, y=-B, width=B, height=B), "size_ne_sw"),
        )
        for edge, place_kw, cursor in grips:
            g = tk.Frame(self, bg=C_HDR, cursor=cursor)
            g.place(**place_kw)
            g.bind("<ButtonPress-1>", lambda e, ed=edge: self._resize_start(e, ed))
            g.bind("<B1-Motion>",     self._resize_move)

        # ── title bar (drag to move) ──
        top = tk.Frame(inner, bg=C_HDR, cursor="fleur")
        top.pack(fill=tk.X)
        top.bind("<ButtonPress-1>", self._drag_start)
        top.bind("<B1-Motion>",     self._drag_move)
        self._title = tk.Label(top, text=self._title_text(), bg=C_HDR,
                               fg=C_PORT, font=FONT_BOLD, padx=8, pady=4)
        self._title.pack(side=tk.LEFT)
        self._title.bind("<ButtonPress-1>", self._drag_start)
        self._title.bind("<B1-Motion>",     self._drag_move)
        # modem-line controls/indicators — sit right after the baud/8N1 text
        self._sig_frame = tk.Frame(top, bg=C_HDR)
        self._sig_frame.pack(side=tk.LEFT, padx=(6, 0))
        self._build_signals()
        tk.Button(top, text=" ✕ ", bg=C_HDR, fg="#666666", bd=0,
                  relief="flat", font=("Segoe UI", 8),
                  activebackground="#c0392b", activeforeground="white",
                  command=self._on_close).pack(side=tk.RIGHT)
        self._aot_btn = tk.Button(top, text=" 📌 ", bg=C_HDR, bd=0,
                  relief="flat", font=("Segoe UI", 8),
                  activebackground=C_HDR, activeforeground=C_NEW,
                  command=self._toggle_topmost)
        self._aot_btn.pack(side=tk.RIGHT)
        self._update_aot_btn()
        tk.Button(top, text=" ⚙ ", bg=C_HDR, fg="#666666", bd=0,
                  relief="flat", font=("Segoe UI", 8),
                  activebackground=C_HDR, activeforeground="#aaaaaa",
                  command=self._open_port_settings).pack(side=tk.RIGHT)
        tk.Button(top, text=" Clear ", bg=C_HDR, fg="#666666", bd=0,
                  relief="flat", font=("Segoe UI", 8),
                  activebackground=C_HDR, activeforeground="#aaaaaa",
                  command=self._clear).pack(side=tk.RIGHT)
        self._conn_btn = tk.Button(top, bg=C_HDR, bd=0,
                  relief="flat", font=("Segoe UI", 8),
                  activebackground=C_HDR, activeforeground="white",
                  command=self._toggle_connection)
        self._conn_btn.pack(side=tk.RIGHT)
        self._update_conn_btn()

        out_frm = tk.Frame(inner, bg=C_BG)
        out_frm.pack(fill=tk.BOTH, expand=True)
        self.out = tk.Text(out_frm, bg=C_BG, fg=C_DESC, font=FONT,
                           bd=0, wrap=tk.NONE, padx=6, pady=4,
                           insertbackground=C_PORT)
        scr = tk.Scrollbar(out_frm, command=self.out.yview,
                           bg=C_HDR, troughcolor=C_BG, bd=0)
        self.out.configure(yscrollcommand=scr.set, state="disabled")
        scr.pack(side=tk.RIGHT, fill=tk.Y)
        self.out.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.out.tag_configure("tx",   foreground=C_DIM)
        self.out.tag_configure("rx",   foreground=C_DESC)
        self.out.tag_configure("info", foreground=C_NEW)

        self.inp = tk.Entry(inner, bg=C_HDR, fg=C_PORT, font=FONT, bd=0,
                            relief="flat", insertbackground=C_PORT)
        self.inp.pack(fill=tk.X, padx=6, pady=(0, 6))
        self.inp.bind("<Return>", self._on_send)
        self.bind("<Escape>", lambda _: self._on_close())

    # ── title bar text ──────────────────────────────────────────────────────
    def _title_text(self) -> str:
        """e.g. 'COM7 - MyBoard - 115200 8N1', or without the name if unset."""
        s = self.dev.port_settings()
        params = f"{s['baud']} {s['data_bits']}{s['parity']}{s['stop_bits']}"
        name = self.dev.name.strip()
        parts = [self.dev.id] + ([name] if name else []) + [params]
        return " - ".join(parts)

    def _refresh_title(self):
        if getattr(self, "_title", None) is not None:
            self._title.configure(text=self._title_text())

    # ── modem-control lines (title-bar buttons/indicators) ───────────────────
    def _signals_mode(self) -> str:
        m = self.dev.port_settings().get("signals", "none")
        return m if m in ("none", "controls", "full",
                          "esptool", "esptool_usb") else "none"

    def _build_signals(self):
        """(Re)create the RTS/DTR controls and CTS/DSR/DCD indicators to match
        the device's configured signals mode. Called on build and after the
        port-settings dialog may have changed the mode."""
        if self._sig_frame is None:
            return
        for w in self._sig_frame.winfo_children():
            w.destroy()
        self._rts_btn = self._dtr_btn = None
        self._cts_lbl = self._dsr_lbl = self._dcd_lbl = None

        mode = self._signals_mode()
        if mode == "none":
            return

        if mode in ("esptool", "esptool_usb"):
            # ESP/Arduino auto-reset: two buttons that pulse RTS/DTR to reboot
            # into the user app or the ROM bootloader (download mode). The reset
            # sequence differs per mode (UART bridge vs native USB-Serial-JTAG);
            # see _esp_reset. No live line-state indicators here.
            tk.Button(self._sig_frame, text=" app ", bg=C_HDR, fg=C_FREE, bd=0,
                      relief="flat", font=("Segoe UI", 8),
                      activebackground=C_HDR, activeforeground="white",
                      command=lambda: self._esp_reset(False)).pack(side=tk.LEFT)
            tk.Button(self._sig_frame, text=" boot ", bg=C_HDR, fg=C_VIDPID, bd=0,
                      relief="flat", font=("Segoe UI", 8),
                      activebackground=C_HDR, activeforeground="white",
                      command=lambda: self._esp_reset(True)).pack(side=tk.LEFT)
            return

        # RTS / DTR — clickable outputs we drive
        self._rts_btn = tk.Button(self._sig_frame, text=" RTS ", bg=C_HDR, bd=0,
                  relief="flat", font=("Segoe UI", 8),
                  activebackground=C_HDR, activeforeground="white",
                  command=lambda: self._toggle_output("rts"))
        self._rts_btn.pack(side=tk.LEFT)
        self._dtr_btn = tk.Button(self._sig_frame, text=" DTR ", bg=C_HDR, bd=0,
                  relief="flat", font=("Segoe UI", 8),
                  activebackground=C_HDR, activeforeground="white",
                  command=lambda: self._toggle_output("dtr"))
        self._dtr_btn.pack(side=tk.LEFT)

        if mode == "full":
            # CTS / DSR / DCD — read-only inputs (updated each poll)
            self._cts_lbl = tk.Label(self._sig_frame, text=" CTS ", bg=C_HDR,
                                     fg=C_DIM, font=("Segoe UI", 8))
            self._cts_lbl.pack(side=tk.LEFT)
            self._dsr_lbl = tk.Label(self._sig_frame, text=" DSR ", bg=C_HDR,
                                     fg=C_DIM, font=("Segoe UI", 8))
            self._dsr_lbl.pack(side=tk.LEFT)
            self._dcd_lbl = tk.Label(self._sig_frame, text=" DCD ", bg=C_HDR,
                                     fg=C_DIM, font=("Segoe UI", 8))
            self._dcd_lbl.pack(side=tk.LEFT)

        self._update_signal_widgets()

    def _toggle_output(self, line: str):
        """Flip RTS or DTR; drive the live port if open and remember the state
        so it's re-applied on the next (re)connect."""
        if line == "rts":
            self._rts = not self._rts
        else:
            self._dtr = not self._dtr
        self._apply_modem_outputs()
        self._update_signal_widgets()

    def _apply_modem_outputs(self):
        """Drive RTS/DTR on the live port to our tracked state. No-op if the
        port isn't open."""
        if self.serial is None:
            return
        try:
            self.serial.rts = self._rts
            self.serial.dtr = self._dtr
        except (serial.SerialException, OSError):
            pass

    def _esp_reset(self, to_bootloader: bool):
        """Reboot an ESP/Arduino board into the user app or the ROM bootloader
        (download mode) by toggling RTS/DTR. Two strategies, picked by the
        device's signals mode:

        "esptool" (UART bridge — CH340/CP2102/…): the classic discrete-transistor
            auto-reset where RTS drives EN/RESET and DTR drives GPIO0. Mirrors
            esptool's ClassicReset / HardReset.

        "esptool_usb" (native USB-Serial-JTAG — ESP32-S2/S3/C3/…): the chip's
            built-in USB peripheral interprets the lines itself; the reset
            transition is deliberately routed through (1,1) rather than (0,0),
            and the chip re-enumerates (the COM port drops and returns) as part
            of the reset, which is expected, not a failure. Mirrors esptool's
            USBJTAGSerialReset / HardReset(uses_usb=True).

        Every write goes through set_dtr/set_rts so the two-transfer ordering and
        the Windows DTR-re-assert work-around match esptool on each edge."""
        ser = self.serial
        if ser is None:
            self._info("[not connected — can't reset]")
            return

        usb = (self._signals_mode() == "esptool_usb")

        def set_dtr(state: bool):
            ser.dtr = state

        def set_rts(state: bool):
            # Some Windows USB-serial drivers (usbser.sys) drop DTR when RTS is
            # changed, so re-assert the current DTR right after (esptool's fix).
            ser.rts = state
            ser.dtr = ser.dtr

        try:
            if usb:
                if to_bootloader:
                    # esptool USBJTAGSerialReset
                    set_rts(False)
                    set_dtr(False)                  # idle
                    time.sleep(ESP_RESET_DELAY_S)
                    set_dtr(True)                   # assert IO0 (boot)
                    set_rts(False)
                    time.sleep(ESP_RESET_DELAY_S)
                    set_rts(True)                   # reset — routed through (1,1)
                    set_dtr(False)
                    set_rts(True)                   # re-set RTS so Windows propagates DTR
                    time.sleep(ESP_RESET_DELAY_S)
                    set_dtr(False)
                    set_rts(False)                  # out of reset → enters ROM loader
                    self._info("[USB reset → bootloader / download mode]")
                else:
                    # esptool HardReset(uses_usb=True) — pulse EN, IO0 left high
                    set_dtr(False)                  # IO0 = HIGH (run, not boot)
                    set_rts(True)                   # EN  = LOW
                    time.sleep(ESP_USB_RESET_DELAY_S)
                    set_rts(False)                  # EN  = HIGH → run the app
                    time.sleep(ESP_USB_RESET_DELAY_S)
                    self._info("[USB reset → app mode]")
            if not usb:
                if to_bootloader:
                    # esptool ClassicReset
                    set_dtr(False)                  # GPIO0 = HIGH
                    set_rts(True)                   # EN    = LOW  (chip in reset)
                    time.sleep(ESP_RESET_DELAY_S)
                    set_dtr(True)                   # GPIO0 = LOW
                    set_rts(False)                  # EN    = HIGH (released → ROM loader)
                    time.sleep(ESP_BOOT_DELAY_S)
                    set_dtr(False)                  # GPIO0 = HIGH (release)
                    self._info("[reset → bootloader / download mode]")
                else:
                    # esptool HardReset — pulse EN with GPIO0 high
                    set_dtr(False)                  # GPIO0 = HIGH (high before the pulse)
                    set_rts(True)                   # EN    = LOW  (chip in reset)
                    time.sleep(ESP_RESET_DELAY_S)
                    set_rts(False)                  # EN    = HIGH (released → run app)
                    self._info("[reset → app mode]")
        except (serial.SerialException, OSError) as e:
            # native USB drops & re-enumerates as part of the reset — expected;
            # the main window's presence poll will reconnect when it returns
            if usb:
                self._info("[USB reset signal sent — device re-enumerating…]")
            else:
                self._info(f"[reset failed: {e}]")
            return

        # remember the resting line state (board running) so a later reconnect
        # re-applies it rather than re-pulsing reset
        try:
            self._rts = bool(ser.rts)
            self._dtr = bool(ser.dtr)
        except (serial.SerialException, OSError):
            pass

    # ── file watch: release the port and re-grab it after a build ──────────
    def _watch_path(self) -> str:
        return (self.dev.port_settings().get("watch_file") or "").strip()

    def _watch_snapshot(self):
        """(exists, mtime, size) for the watch file. None means no path set
        OR the previous baseline should be kept (transient stat error)."""
        p = self._watch_path()
        if not p:
            return None
        try:
            st = os.stat(p)
            return (True, st.st_mtime, st.st_size)
        except FileNotFoundError:
            return (False, None, None)
        except OSError:
            return None                                  # transient: keep baseline

    def _schedule_watch_poll(self):
        self._watch_poll_id = self.after(TERM_WATCH_POLL_MS, self._watch_poll)

    def _watch_poll(self):
        self._watch_poll_id = None
        if self._closing:
            return
        ps = self.dev.port_settings()
        enabled = bool(ps.get("watch_enable", False))
        if not enabled or not self._watch_path():
            # disabled (or no path) → drop baseline so a future enable starts
            # a fresh observation rather than firing on the very first poll
            self._watch_baseline = None
            self._schedule_watch_poll()
            return
        snap = self._watch_snapshot()
        if snap is not None:
            if self._watch_baseline is None:
                self._watch_baseline = snap              # first observation
            elif snap != self._watch_baseline:
                self._watch_baseline = snap
                self._watch_trigger()
        self._schedule_watch_poll()

    def _watch_trigger(self):
        """File changed. If we hold the port (or are actively waiting for it),
        release and enter a hold window so the flasher can grab it; repeated
        changes within the window restart the timer so we wait for the file
        to settle before reconnecting. Ignored if the user has manually
        disconnected."""
        try:
            timeout = float(self.dev.port_settings().get("watch_reconnect_s", 5.0))
        except (TypeError, ValueError):
            timeout = 5.0
        timeout = max(0.0, min(20.0, timeout))
        manual  = (timeout == 0.0)
        suffix  = "reconnect manually" if manual else f"for {timeout:.1f} s"

        if self.state == "connected":
            self._close_serial()
            self._set_state("disconnected")
            self._watch_holding = True
            self._info(f"[watch: file changed — releasing port; {suffix}]")
        elif self.state == "waiting":
            self._set_state("disconnected")
            self._watch_holding = True
            self._info(f"[watch: file changed — holding off; {suffix}]")
        elif self.state == "disconnected" and self._watch_holding:
            self._info("[watch: file changed again]" if manual
                       else f"[watch: file changed again — extending hold to {timeout:.1f} s]")
        else:
            return                                       # user-disconnected: ignore

        # Cancel any pending timer (handles the case where timeout was just
        # lowered to 0, leaving an in-flight reconnect that we no longer want).
        if self._watch_release_id is not None:
            try:
                self.after_cancel(self._watch_release_id)
            except tk.TclError:
                pass
            self._watch_release_id = None
        if not manual:
            self._watch_release_id = self.after(int(timeout * 1000), self._watch_reconnect)

    def _watch_reconnect(self):
        """Hold window elapsed — resume the connection. If the port is still
        busy, reconnect() falls back to 'waiting' which the main window's
        port-free poll picks up."""
        self._watch_release_id = None
        if self._closing or not self._watch_holding:
            return                                       # cancelled / overridden
        self._watch_holding = False
        self.reconnect()

    def _watch_cancel_hold(self):
        """Drop any pending watch-reconnect timer (manual user action wins)."""
        if self._watch_release_id is not None:
            try:
                self.after_cancel(self._watch_release_id)
            except tk.TclError:
                pass
            self._watch_release_id = None
        self._watch_holding = False

    def _update_signal_widgets(self):
        """Repaint RTS/DTR from our tracked output state and CTS/DSR/DCD from the
        live port. Green = asserted/high, dim = low/deasserted."""
        def paint(w, on):
            if w is not None:
                try:
                    w.configure(fg=C_FREE if on else C_DIM)
                except tk.TclError:
                    pass
        paint(self._rts_btn, self._rts)
        paint(self._dtr_btn, self._dtr)
        if self._cts_lbl is None:                       # not in "full" mode
            return
        cts = dsr = dcd = False
        if self.serial is not None:
            try:
                cts, dsr, dcd = self.serial.cts, self.serial.dsr, self.serial.cd
            except (serial.SerialException, OSError, AttributeError):
                cts = dsr = dcd = False
        paint(self._cts_lbl, cts)
        paint(self._dsr_lbl, dsr)
        paint(self._dcd_lbl, dcd)

    # ── connect / disconnect state machine ──────────────────────────────────
    # connection state lives on the Device ("connected"/"waiting"/"disconnected")
    # so the main window can read it straight off the row's device.
    @property
    def state(self) -> str:
        return self.dev.conn_state

    @state.setter
    def state(self, st: str):
        self.dev.conn_state = st

    def _set_state(self, st: str):
        self.state = st
        self._update_conn_btn()

    def _toggle_connection(self):
        # Manual user action wins over any watch hold in progress.
        self._watch_cancel_hold()
        if self.state == "disconnected":
            # user wants a connection: take it now, else wait for the device
            self._set_state("connected" if self._open_serial() else "waiting")
        else:
            # connected or waiting → user opts out; no auto-reconnect
            self._close_serial()
            self._set_state("disconnected")
            self._info("[disconnected]")

    def _update_conn_btn(self):
        if getattr(self, "_conn_btn", None) is None:
            return
        # label = the action a click performs; colour = the current state
        text, fg = {
            "connected":    (" Disconnect ", C_FREE),
            "waiting":      (" Cancel ",     C_WAIT),
            "disconnected": (" Connect ",    C_OPEN),
        }.get(self.state, (" Connect ", C_OPEN))
        try:
            self._conn_btn.configure(text=text, fg=fg)
        except tk.TclError:                              # window tearing down
            pass

    # ── always-on-top (own toggle, independent of the main window) ──────────
    def _toggle_topmost(self):
        self._topmost = not self._topmost
        try:
            self.attributes("-topmost", self._topmost)
        except tk.TclError:
            pass
        self.parent.settings["terminal_always_on_top"] = self._topmost
        save_settings(self.parent.settings)
        self._update_aot_btn()

    def _update_aot_btn(self):
        # lit when pinned, dimmed when not — a simple on/off switch
        self._aot_btn.configure(fg=C_NEW if self._topmost else "#666666")

    # ── move / resize ──────────────────────────────────────────────────────
    def _drag_start(self, e):
        self._drag_ox = e.x_root - self.winfo_rootx()
        self._drag_oy = e.y_root - self.winfo_rooty()

    def _drag_move(self, e):
        self.geometry(f"+{e.x_root - self._drag_ox}+{e.y_root - self._drag_oy}")

    def _resize_start(self, e, edge):
        self._resize_edge = edge
        self._rs = (e.x_root, e.y_root,
                    self.winfo_rootx(), self.winfo_rooty(),
                    self.winfo_width(), self.winfo_height())

    def _resize_move(self, e):
        if self._resize_edge is None:
            return
        sx, sy, x0, y0, w0, h0 = self._rs
        dx, dy = e.x_root - sx, e.y_root - sy
        x, y, w, h = x0, y0, w0, h0
        edge = self._resize_edge
        if "e" in edge:
            w = max(TERM_MIN_W, w0 + dx)
        if "s" in edge:
            h = max(TERM_MIN_H, h0 + dy)
        if "w" in edge:
            w = max(TERM_MIN_W, w0 - dx)
            x = x0 + (w0 - w)
        if "n" in edge:
            h = max(TERM_MIN_H, h0 - dy)
            y = y0 + (h0 - h)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _open_serial(self) -> bool:
        s = self.dev.port_settings()
        try:
            self.serial = serial.Serial(
                self.dev.id,
                baudrate=s["baud"],
                bytesize=s["data_bits"],
                parity=s["parity"],
                stopbits=s["stop_bits"],
                timeout=0)
            # Drive RTS/DTR immediately, then again after 100 ms — some USB-serial
            # adapters drop line-state commands issued in the first few ms after
            # enumeration, so a second write ensures the buttons' values stick.
            self._apply_modem_outputs()
            self.after(100, self._apply_modem_outputs)
            self._info(f"[connected @ {s['baud']} "
                       f"{s['data_bits']}{s['parity']}{s['stop_bits']}]")
            self._update_signal_widgets()
            return True
        except (serial.SerialException, OSError, ValueError) as e:
            self.serial = None
            self._info(f"[open failed: {e}]")
            return False

    def _close_serial(self):
        if self.serial is not None:
            try:
                self.serial.close()
            except Exception:                           # noqa: BLE001
                pass
            self.serial = None

    def _schedule_poll(self):
        self._poll_id = self.after(TERM_POLL_MS, self._poll)

    def _poll(self):
        self._poll_id = None
        if self._closing:
            return
        if self.serial is not None:
            try:
                n = self.serial.in_waiting
                if n:
                    data = self.serial.read(n)
                    self._append(data.decode("utf-8", errors="replace"), "rx")
            except (serial.SerialException, OSError) as e:
                self._info(f"[read error: {e}]")
                self._close_serial()
        if self._cts_lbl is not None:                   # "full" mode: poll inputs
            self._update_signal_widgets()
        self._schedule_poll()

    def _on_send(self, _e=None):
        line = self.inp.get()
        self.inp.delete(0, tk.END)
        if self.serial is None:
            self._info("[not connected]")
            return "break"
        s = self.dev.port_settings()
        try:
            self.serial.write((line + s["line_ending"])
                              .encode("utf-8", errors="replace"))
            self._append(line + "\n", "tx")
        except (serial.SerialException, OSError) as e:
            self._info(f"[write error: {e}]")
            self._close_serial()
        return "break"

    def _append(self, text, tag):
        # auto-scroll only when already pinned to the bottom; if the user has
        # scrolled up, hold position so they can read/copy while data arrives
        at_bottom = self.out.yview()[1] >= 0.999
        self.out.configure(state="normal")
        self.out.insert(tk.END, text, tag)
        self.out.configure(state="disabled")
        if at_bottom:
            self.out.see(tk.END)

    def _info(self, text):
        self._append(text + "\n", "info")

    def _clear(self):
        self.out.configure(state="normal")
        self.out.delete("1.0", tk.END)
        self.out.configure(state="disabled")

    def _open_port_settings(self):
        PortSettingsDialog(self)

    def reconnect(self):
        """Reopen the serial port with whatever port_settings are now in effect."""
        self._close_serial()
        # The watch path/timeout may have just changed; re-seed on the next poll.
        self._watch_baseline = None
        self._set_state("connected" if self._open_serial() else "waiting")
        self._refresh_title()
        self._build_signals()

    def reconnect_on_unplug(self) -> bool:
        return bool(self.dev.port_settings().get("reconnect_on_unplug", True))

    def _position_above_main(self):
        try:
            x = self.parent.winfo_rootx()
            y = self.parent.winfo_rooty() - self.winfo_height() - 5
            self.geometry(f"+{x}+{max(0, y)}")
        except tk.TclError:
            pass

    def _on_configure(self, _e=None):
        if self._closing:
            return
        try:
            w, h = self.winfo_width(), self.winfo_height()
            x, y = self.winfo_rootx(), self.winfo_rooty()
        except tk.TclError:
            return
        if w > 1 and h > 1:
            self.dev.save_geometry(f"{w}x{h}+{x}+{y}")

    def _close_quiet(self):
        """Tear down without notifying parent — used by unplug / shutdown paths."""
        if self._closing:
            return
        self._closing = True
        self.dev.conn_state = "disconnected"            # release the device's state
        self.dev.terminal = None                        # unbind from the device
        save_settings(self.parent.settings)             # persist geometry
        if self._poll_id is not None:
            try:
                self.after_cancel(self._poll_id)
            except tk.TclError:
                pass
            self._poll_id = None
        if self._watch_poll_id is not None:
            try:
                self.after_cancel(self._watch_poll_id)
            except tk.TclError:
                pass
            self._watch_poll_id = None
        self._watch_cancel_hold()
        self._close_serial()
        try:
            self.destroy()
        except tk.TclError:
            pass

    def _on_close(self):
        """User clicked the X — clean up and unbind from the device."""
        self._close_quiet()


# ── port settings dialog (per VID:PID:Serial) ─────────────────────────────────
class PortSettingsDialog(tk.Toplevel):
    BAUDS = (4800, 9600, 19200, 38400, 57600, 115200, 230400, 460800, 921600, 1000000, 2000000)
    LE_TO_LABEL = {"\r\n": "CRLF", "\n": "LF", "\r": "CR", "": "none"}
    LABEL_TO_LE = {v: k for k, v in LE_TO_LABEL.items()}
    SIG_TO_LABEL = {"none": "None", "controls": "RTS/DTR", "full": "Full",
                    "esptool": "App/Boot UART", "esptool_usb": "App/Boot USB-JTAG"}
    LABEL_TO_SIG = {v: k for k, v in SIG_TO_LABEL.items()}

    def __init__(self, terminal: "Terminal"):
        super().__init__(terminal)
        self.terminal = terminal
        self.title("Port settings")
        self.configure(bg=C_BG)
        self.transient(terminal)
        self.resizable(False, False)
        self.attributes("-topmost", True)
        self._build()
        self.update_idletasks()
        px = terminal.winfo_rootx() + (terminal.winfo_width()
                                        - self.winfo_width()) // 2
        py = terminal.winfo_rooty() + 80
        self.geometry(f"+{max(px, 0)}+{max(py, 0)}")
        self.grab_set()
        self.bind("<Escape>", lambda _: self.destroy())

    def _build(self):
        s   = self.terminal.dev.port_settings()
        pad = dict(padx=8, pady=4)

        frm = tk.Frame(self, bg=C_BG)
        frm.pack(padx=12, pady=(12, 6))

        def label(r, text):
            tk.Label(frm, text=text, bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                     ).grid(row=r, column=0, sticky="w", **pad)

        def spin(r, var, values, value):
            tk.Spinbox(frm, values=values, textvariable=var, width=10,
                       bg=C_HDR, fg=C_PORT, font=FONT,
                       buttonbackground=C_HDR, relief="flat",
                       insertbackground=C_PORT
                       ).grid(row=r, column=1, sticky="w", **pad)
            # `values=` overwrites the textvariable to its first entry on
            # creation; restore the real current value afterwards.
            var.set(value)

        label(0, "Baud:")
        self.var_baud = tk.IntVar()
        spin(0, self.var_baud, self.BAUDS, s["baud"])

        label(1, "Data bits:")
        self.var_data = tk.IntVar()
        spin(1, self.var_data, (5, 6, 7, 8), s["data_bits"])

        label(2, "Parity:")
        self.var_par = tk.StringVar()
        spin(2, self.var_par, ("N", "E", "O", "M", "S"), s["parity"])

        label(3, "Stop bits:")
        self.var_stop = tk.DoubleVar()
        spin(3, self.var_stop, (1, 1.5, 2), s["stop_bits"])

        label(4, "Line ending:")
        self.var_le = tk.StringVar()
        spin(4, self.var_le, tuple(self.LE_TO_LABEL.values()),
             self.LE_TO_LABEL.get(s["line_ending"], "CRLF"))

        label(5, "Signals:")
        self.var_sig = tk.StringVar(
            value=self.SIG_TO_LABEL.get(s.get("signals", "none"), "None"))
        sig_menu = tk.OptionMenu(frm, self.var_sig, *self.SIG_TO_LABEL.values())
        sig_menu.configure(width=13, bg=C_HDR, fg=C_PORT, font=FONT,
                           activebackground=C_HDR, activeforeground="white",
                           highlightthickness=0, bd=0, relief="flat", anchor="w")
        sig_menu["menu"].configure(bg=C_HDR, fg=C_PORT,
                                   activebackground=C_PORT, activeforeground=C_BG)
        sig_menu.grid(row=5, column=1, sticky="w", padx=8, pady=4)

        self.var_reconn = tk.BooleanVar(value=s["reconnect_on_unplug"])
        tk.Checkbutton(frm, text="Reconnect on unplug",
                       variable=self.var_reconn,
                       bg=C_BG, fg=C_DESC, selectcolor=C_HDR,
                       activebackground=C_BG, activeforeground=C_DESC,
                       font=FONT, anchor="w"
                       ).grid(row=6, column=0, columnspan=2, sticky="w", **pad)

        # ── file watch: release the port when a chosen file changes ────────
        self.var_watch_en = tk.BooleanVar(value=bool(s.get("watch_enable", False)))
        tk.Checkbutton(frm, text="Watch file (change=disconnect)",
                       variable=self.var_watch_en,
                       bg=C_BG, fg=C_DESC, selectcolor=C_HDR,
                       activebackground=C_BG, activeforeground=C_DESC,
                       font=FONT, anchor="w"
                       ).grid(row=7, column=0, columnspan=2, sticky="w", **pad)

        tk.Label(frm, text="     ↳ path:",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=8, column=0, sticky="w", **pad)
        self.var_watch = tk.StringVar(value=s.get("watch_file", ""))
        watch_row = tk.Frame(frm, bg=C_BG)
        watch_row.grid(row=8, column=1, sticky="we", **pad)
        tk.Entry(watch_row, textvariable=self.var_watch, width=30,
                 bg=C_HDR, fg=C_PORT, font=FONT, bd=0, relief="flat",
                 insertbackground=C_PORT
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(watch_row, text=" … ", command=self._browse_watch,
                  bg=C_HDR, fg=C_DESC, bd=0, relief="flat", font=FONT,
                  activebackground=C_HDR, activeforeground="white"
                  ).pack(side=tk.LEFT, padx=(4, 0))
        tk.Button(watch_row, text=" ✕ ",
                  command=lambda: self.var_watch.set(""),
                  bg=C_HDR, fg=C_DESC, bd=0, relief="flat", font=FONT,
                  activebackground=C_HDR, activeforeground="white"
                  ).pack(side=tk.LEFT, padx=(2, 0))

        tk.Label(frm, text="     ↳ reconnect after (s, 0 = manual):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=9, column=0, sticky="w", **pad)
        self.var_watch_to = tk.DoubleVar(value=float(s.get("watch_reconnect_s", 5.0)))
        tk.Spinbox(frm, from_=0, to=20, increment=0.1, format="%.1f",
                   textvariable=self.var_watch_to,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=9, column=1, sticky="w", **pad)

        btns = tk.Frame(self, bg=C_BG)
        btns.pack(fill=tk.X, padx=12, pady=(0, 12))
        tk.Button(btns, text="Cancel", command=self.destroy,
                  bg=C_HDR, fg=C_DESC, bd=0, relief="flat", font=FONT,
                  activebackground=C_HDR, activeforeground="white",
                  padx=14, pady=3).pack(side=tk.RIGHT, padx=(6, 0))
        tk.Button(btns, text="Save", command=self._save,
                  bg=C_HDR, fg=C_PORT, bd=0, relief="flat", font=FONT,
                  activebackground=C_HDR, activeforeground="white",
                  padx=14, pady=3).pack(side=tk.RIGHT)

    def _save(self):
        if self.terminal.dev.custom_key is None:
            self.destroy()
            return
        new = dict(DEFAULT_PORT_SETTINGS)
        try: new["baud"]      = int(self.var_baud.get())
        except (tk.TclError, ValueError): pass
        try: new["data_bits"] = int(self.var_data.get())
        except (tk.TclError, ValueError): pass
        try:
            p = str(self.var_par.get()).strip().upper()
            if p and p[0] in "NEOMS":
                new["parity"] = p[0]
        except (tk.TclError, ValueError): pass
        try:
            sv = float(self.var_stop.get())
            new["stop_bits"] = int(sv) if sv.is_integer() else sv
        except (tk.TclError, ValueError): pass
        try: new["line_ending"] = self.LABEL_TO_LE.get(
                                    self.var_le.get(), "\r\n")
        except (tk.TclError, ValueError): pass
        try: new["reconnect_on_unplug"] = bool(self.var_reconn.get())
        except (tk.TclError, ValueError): pass
        try: new["signals"] = self.LABEL_TO_SIG.get(self.var_sig.get(), "none")
        except (tk.TclError, ValueError): pass
        try: new["watch_enable"] = bool(self.var_watch_en.get())
        except (tk.TclError, ValueError): pass
        try: new["watch_file"] = str(self.var_watch.get()).strip()
        except (tk.TclError, ValueError): pass
        try:
            v = float(self.var_watch_to.get())
            new["watch_reconnect_s"] = max(0.0, min(20.0, v))
        except (tk.TclError, ValueError): pass

        # merge serial keys into the device's entry (a "name" set elsewhere survives)
        self.terminal.dev.update_config(**new)
        save_settings(self.terminal.parent.settings)
        self.terminal.reconnect()
        self.destroy()

    def _browse_watch(self):
        path = filedialog.askopenfilename(parent=self,
                                          title="Watch file",
                                          initialfile=self.var_watch.get())
        if path:
            self.var_watch.set(path)


if __name__ == "__main__":
    ComMonitor().mainloop()
