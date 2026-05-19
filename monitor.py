#!/usr/bin/env python3
"""
USB COM Port Monitor
Small always-on-top desktop widget.

Usage:  python monitor.py
Needs:  pip install pyserial pywin32
"""

import tkinter as tk
import tkinter.font as tkfont
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
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                             "settings.json")

DEFAULT_SETTINGS = {
    "highlight_duration_s":     20.0,         # row-flash fade duration
    "show_removed_s":           60,           # 0 = don't show removed ports
    "always_on_top":            True,         # pin the window above all others
    "always_on_top_timeout_s":  0,            # 0 = never drop; else drop after idle
    "interaction_timeout_s":    2,            # seconds with no mouse/focus before fading
    "move_to_back_enable":      False,        # push window to back of Z-order on idle
    "move_to_back_timeout_s":   300,
    "normal_alpha":             ALPHA_OPAQUE, # window opacity (rows keep this)
    "window_x":                 None,         # last-known position
    "window_y":                 None,
    "custom_names":             {},           # {"VID:PID:SERIAL": user-given name}
}


def load_settings() -> dict:
    # shallow-copy each default value so callers can mutate nested dicts
    # (e.g. custom_names) without bleeding into DEFAULT_SETTINGS.
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


def _custom_key(vid, pid, serial):
    """Map key for user-given names. Serial is included verbatim — an empty
    serial yields a key like "VID:PID:" which still uniquely identifies the
    "no-serial" variant of that VID:PID. Location is intentionally ignored."""
    if vid is None or pid is None:
        return None
    return f"{vid:04X}:{pid:04X}:{serial or ''}"


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
        self._update_alpha()
        self.after(FADE_TICK_MS, self._animate_fade)

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
    def _lift_pair(self):
        """Bring both windows to the top of the stack while preserving their
        relative order (chrome below, rows on top)."""
        try:
            self.lift()
            self._rows_win.lift()
        except tk.TclError:
            pass

    def _lower_pair(self):
        try:
            self._rows_win.lower()
            self.lower()
        except tk.TclError:
            pass

    def _set_topmost_pair(self, on: bool):
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

    # ── inline edit of custom serial name ─────────────────────────────────────
    def _begin_serial_edit(self, lbl, key, desc_lbl=None):
        """Overlay an Entry on the serial cell so the user can type a custom
        name keyed by VID:PID:Serial. While editing, row rebuilds are paused
        (see `_editing_key` short-circuit in `_refresh`)."""
        if self._editing_key is not None:
            return
        self._editing_key   = key
        self._edit_lbl      = lbl
        self._edit_desc_lbl = desc_lbl

        info = lbl.grid_info()
        row, col = int(info["row"]), int(info["column"])
        row_bg   = lbl.cget("bg")
        current  = self.settings["custom_names"].get(key, "")

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
            v  = entry.get().strip()
            cn = self.settings["custom_names"]
            if v:
                cn[key] = v
            else:
                cn.pop(key, None)
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
        super().destroy()

    def _apply_settings(self):
        """Re-apply settings to derived state after they've been edited."""
        self._flash_decay_k    = self._compute_flash_decay_k()
        self._last_interaction = time.time()
        self._set_topmost_pair(self.settings["always_on_top"])
        self._top_active = self.settings["always_on_top"]
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

        # port changes lift the window pair and pin them on top for one
        # interaction timeout so the user actually notices the connect/disconnect.
        port_changed = self._initialized and current != self._known_ports
        if port_changed:
            self._top_until = now + self.settings["interaction_timeout_s"]
            self._lift_pair()
        self._known_ports = current

        # interaction poll: pointer over EITHER window (or focus on either)
        # holds the timer at zero. Once both are gone, the timer starts counting.
        try:
            px, py = self.winfo_pointerxy()
            in_any = False
            for w in (self, self._rows_win):
                wx, wy = w.winfo_rootx(), w.winfo_rooty()
                ww, wh = w.winfo_width(), w.winfo_height()
                if wx <= px < wx + ww and wy <= py < wy + wh:
                    in_any = True
                    break
            focused = (self.focus_displayof() is not None
                       or self._rows_win.focus_displayof() is not None)
            if in_any or focused:
                self._last_interaction = now
        except tk.TclError:
            pass

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

        rendered = sorted(set(current) | set(self._disappeared_at))
        if not rendered:
            self._empty_lbl.grid(row=0, column=0, columnspan=len(COLS),
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

                custom_key  = _custom_key(info.get("vid"), info.get("pid"),
                                          info.get("serial_number"))
                custom_name = (self.settings["custom_names"].get(custom_key)
                               if custom_key else None)

                # (column, columnspan, label-kwargs)
                cells = [
                    (0, 1, dict(text=dev,                  fg=port_c, anchor="w")),
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
                for col, span, kw in cells:
                    kw.setdefault("font", row_font)
                    lbl = tk.Label(self._grid, bg=row_bg, padx=3, pady=1, **kw)
                    lbl.grid(row=i, column=col, columnspan=span, sticky="ew")
                    row_w.append(lbl)
                    if col == 5:
                        serial_lbl = lbl
                    elif col == 6:
                        desc_lbl = lbl

                if serial_lbl is not None and custom_key is not None:
                    serial_lbl.bind("<Double-Button-1>",
                        lambda _e, k=custom_key, sw=serial_lbl, dw=desc_lbl:
                            self._begin_serial_edit(sw, k, dw))

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
        tk.Label(frm, text="     ↳ drop after idle (s, 0 = never):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=3, column=0, sticky="w", **pad)
        self.var_aot_to = tk.IntVar(value=s["always_on_top_timeout_s"])
        tk.Spinbox(frm, from_=0, to=86400, increment=10, textvariable=self.var_aot_to,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=3, column=1, sticky="w", **pad)

        # Move to background on idle
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

        # Interaction timeout (seconds before fade begins)
        tk.Label(frm, text="Fade after no interaction (s):",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=6, column=0, sticky="w", **pad)
        self.var_it = tk.IntVar(value=s["interaction_timeout_s"])
        tk.Spinbox(frm, from_=1, to=3600, increment=1, textvariable=self.var_it,
                   width=8, bg=C_HDR, fg=C_PORT, font=FONT,
                   buttonbackground=C_HDR, relief="flat",
                   insertbackground=C_PORT
                   ).grid(row=6, column=1, sticky="w", **pad)

        # Window transparency (rows keep this opacity; live preview)
        tk.Label(frm, text="Window transparency:",
                 bg=C_BG, fg=C_DESC, font=FONT, anchor="w"
                 ).grid(row=7, column=0, sticky="w", **pad)
        self.var_an = tk.DoubleVar(value=s["normal_alpha"])
        tk.Scale(frm, from_=0.4, to=1.0, resolution=0.01,
                 orient=tk.HORIZONTAL, variable=self.var_an,
                 command=lambda v: self._preview_alpha("normal_alpha", v),
                 bg=C_BG, fg=C_DESC, troughcolor=C_HDR,
                 highlightthickness=0, bd=0, length=160,
                 font=FONT, activebackground=C_PORT,
                 ).grid(row=7, column=1, sticky="w", **pad)

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
        try:
            v = int(self.var_it.get())
            if v > 0:
                s["interaction_timeout_s"] = v
        except (tk.TclError, ValueError):
            pass
        try:
            s["normal_alpha"] = float(self.var_an.get())
        except (tk.TclError, ValueError):
            pass
        save_settings(s)
        self.parent._apply_settings()
        self._saved = True
        self.destroy()


if __name__ == "__main__":
    ComMonitor().mainloop()
