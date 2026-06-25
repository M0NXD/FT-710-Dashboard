# FT-710 CAT Control Dashboard
# Copyright (C) 2026 M0NXD
#
# This program is free software: you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation, either version 3 of the License, or (at your option) any later
# version. This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
# FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details. You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

import tkinter as tk
from tkinter import messagebox
import win32com.client
import win32com.client.gencache
import pythoncom
import time
import threading

try:
    import winsound
    _HAS_WINSOUND = True
except ImportError:
    _HAS_WINSOUND = False

# ── OmniRig RigParamX constants ──────────────────────────────────────────────
# Verified against the OmniRig type library (win32com.client.constants).
PM_CW_U  = 0x00800000
PM_CW_L  = 0x01000000
PM_SSB_U = 0x02000000
PM_SSB_L = 0x04000000
PM_DIG_U = 0x08000000
PM_DIG_L = 0x10000000
PM_AM    = 0x20000000
PM_FM    = 0x40000000

# Split-state bitmask
PM_SPLIT_ON = 0x00008000

# Tx-property values (PM_RX is also non-zero, so a plain bool() is wrong)
PM_RX = 0x00200000
PM_TX = 0x00400000

MODE_NAMES = [
    (PM_SSB_U, "USB"),
    (PM_SSB_L, "LSB"),
    (PM_CW_U,  "CW"),
    (PM_CW_L,  "CW-R"),
    (PM_DIG_U, "DIG-U"),
    (PM_DIG_L, "DIG-L"),
    (PM_AM,    "AM"),
    (PM_FM,    "FM"),
]

BANDS = [
    (1_800_000,   2_000_000,   "160m"),
    (3_500_000,   4_000_000,   "80m"),
    (5_000_000,   5_500_000,   "60m"),
    (7_000_000,   7_300_000,   "40m"),
    (10_100_000,  10_150_000,  "30m"),
    (14_000_000,  14_350_000,  "20m"),
    (18_068_000,  18_168_000,  "17m"),
    (21_000_000,  21_450_000,  "15m"),
    (24_890_000,  24_990_000,  "12m"),
    (28_000_000,  29_700_000,  "10m"),
    (50_000_000,  54_000_000,  "6m"),
]

# (label, band-centre Hz, FT-710 mode command)
BAND_DEFAULTS = [
    ("160m", 1_900_000,   "MD01;"),  # LSB
    ("80m",  3_700_000,   "MD01;"),  # LSB
    ("60m",  5_357_000,   "MD02;"),  # USB (60m is USB only)
    ("40m",  7_100_000,   "MD01;"),  # LSB
    ("30m",  10_120_000,  "MD03;"),  # CW
    ("20m",  14_225_000,  "MD02;"),  # USB
    ("17m",  18_120_000,  "MD02;"),  # USB
    ("15m",  21_200_000,  "MD02;"),  # USB
    ("12m",  24_940_000,  "MD02;"),  # USB
    ("10m",  28_400_000,  "MD02;"),  # USB
    ("6m",   50_150_000,  "MD02;"),  # USB
]

# ── S-meter calibration ──────────────────────────────────────────────────────
# Maps the rig's raw SM value (0-255) to the gauge scale (0-9 = S0..S9,
# 9-30 = S9+0..S9+63 at 3 dB/unit). Breakpoints follow the standard Yaesu
# S-meter calibration (raw -> S-unit), interpolated linearly between points.
SMETER_CAL = [
    (0,   0.0), (12, 1.0), (27, 2.0), (40, 3.0), (55,  4.0), (65,  5.0),
    (80,  6.0), (95, 7.0), (112, 8.0), (130, 9.0),
    (150, 12.33), (172, 15.67), (190, 19.0), (220, 22.33),
    (240, 25.67), (255, 29.0),
]


def raw_to_smeter(raw):
    """Convert a raw 0-255 SM reading to the S-meter gauge scale (0-30)."""
    if raw <= SMETER_CAL[0][0]:
        return SMETER_CAL[0][1]
    if raw >= SMETER_CAL[-1][0]:
        return SMETER_CAL[-1][1]
    for (r0, g0), (r1, g1) in zip(SMETER_CAL, SMETER_CAL[1:]):
        if r0 <= raw <= r1:
            return g0 + (raw - r0) / (r1 - r0) * (g1 - g0)
    return SMETER_CAL[-1][1]


def swr_from_raw(raw):
    """Convert the RM6 raw value (0-255) to an SWR ratio.

    The raw value tracks the reflection-coefficient magnitude |Γ| = raw/255,
    so SWR = (1+|Γ|)/(1-|Γ|). Verified on an FT-710: raw 12 -> ~1.1,
    raw ~100 -> ~2.2.
    """
    g = min(raw / 255.0, 0.95)
    return (1.0 + g) / (1.0 - g)


# ── Colour palette ────────────────────────────────────────────────────────────
BG     = "#1a1a2e"
BG2    = "#16213e"
FG     = "#e0e0f0"
FG_DIM = "#6677aa"
GREEN  = "#00cc66"
YELLOW = "#f0a030"
RED    = "#cc3333"
BLUE   = "#4090e0"

_app = None


# ── Horizontal bar meter widget ───────────────────────────────────────────────
class BarMeter(tk.Canvas):
    _W  = 360
    _H  = 32
    _LX = 58
    _BW = 218
    _BY = 16
    _BH = 14

    def __init__(self, parent, label, min_val, max_val,
                 unit="", zones=None, fmt=None, **kw):
        super().__init__(parent, width=self._W, height=self._H,
                         bg=BG, highlightthickness=0, **kw)
        self.min_val = min_val
        self.max_val = max_val
        self.unit    = unit
        self.zones   = zones or [(max_val, GREEN)]
        self.fmt     = fmt or (lambda v: f"{v:.0f}")

        y1 = self._BY - self._BH // 2
        y2 = self._BY + self._BH // 2

        self.create_text(self._LX - 6, self._BY, anchor="e",
                         text=label, fill=FG_DIM,
                         font=("Consolas", 10, "bold"))
        self.create_rectangle(self._LX, y1, self._LX + self._BW, y2,
                               fill="#252540", outline="#3a3a5c")

        for threshold, _ in self.zones[:-1]:
            frac = (threshold - min_val) / (max_val - min_val)
            xd   = self._LX + int(frac * self._BW)
            self.create_line(xd, y1 + 1, xd, y2 - 1,
                             fill="#3a3a5c", dash=(2, 2))

        self._fill = self.create_rectangle(
            self._LX, y1 + 2, self._LX, y2 - 2,
            fill=GREEN, outline=""
        )
        self._peak_line = self.create_line(
            self._LX, y1 + 1, self._LX, y2 - 1,
            fill="#ffffff", width=2, state="hidden"
        )
        self._val = self.create_text(
            self._LX + self._BW + 10, self._BY,
            anchor="w", text="---",
            fill=FG, font=("Consolas", 11, "bold")
        )

    def set_value(self, value):
        y1 = self._BY - self._BH // 2
        y2 = self._BY + self._BH // 2
        if value is None:
            self.itemconfig(self._val, text="---")
            self.coords(self._fill, self._LX, y1 + 2, self._LX, y2 - 2)
            return
        value  = max(self.min_val, min(self.max_val, value))
        frac   = (value - self.min_val) / (self.max_val - self.min_val)
        fill_x = self._LX + int(frac * self._BW)
        color  = self.zones[-1][1]
        for threshold, c in self.zones:
            if value <= threshold:
                color = c
                break
        self.coords(self._fill, self._LX, y1 + 2, fill_x, y2 - 2)
        self.itemconfig(self._fill, fill=color)
        txt = self.fmt(value)
        if self.unit:
            txt += f" {self.unit}"
        self.itemconfig(self._val, text=txt)

    def set_peak(self, value):
        if value is None:
            self.itemconfig(self._peak_line, state="hidden")
            return
        value = max(self.min_val, min(self.max_val, value))
        frac  = (value - self.min_val) / (self.max_val - self.min_val)
        px    = self._LX + int(frac * self._BW)
        y1    = self._BY - self._BH // 2
        y2    = self._BY + self._BH // 2
        self.coords(self._peak_line, px, y1 + 1, px, y2 - 1)
        self.itemconfig(self._peak_line, state="normal")


# ── OmniRig event sink ────────────────────────────────────────────────────────
class OmniRigEvents:
    def OnCustomReply(self, RigNumber, Command, Reply):
        _handle_reply(Reply)

    def CustomReply(self, RigNumber, Command, Reply):
        _handle_reply(Reply)


def _handle_reply(Reply):
    if _app is None:
        return
    try:
        if isinstance(Reply, str):
            s = Reply
        elif isinstance(Reply, (bytes, bytearray)):
            s = bytes(Reply).decode("ascii", errors="ignore")
        else:
            # OmniRig usually hands back a SAFEARRAY -> tuple/list of ints
            try:
                s = bytes(Reply).decode("ascii", errors="ignore")
            except TypeError:
                s = bytes(bytearray(Reply)).decode("ascii", errors="ignore")
        s = s.strip().rstrip(";")
        if not s:
            return
        if s.startswith("PC") and len(s) >= 5:
            _app.current_power = int(s[2:5])
        elif s.startswith("SM0") and len(s) >= 4 and s[3:].isdigit():
            _app.current_smeter = int(s[3:])     # raw 0-255
        elif s.startswith("RM") and len(s) >= 6 and s[3:6].isdigit():
            # RM answer is "RMP1 P2P2P2 P3P3P3;": P2 (s[3:6]) is the 0-255
            # meter value, P3 is fixed 000. (FT-710 CAT manual.)
            meter_id = s[2]
            raw_val  = int(s[3:6])
            if meter_id == "6":          # RM6 = SWR
                _app.current_swr_raw = raw_val
            elif meter_id == "4":        # RM4 = ALC
                _app.current_alc_raw = raw_val
    except Exception:
        pass


# ── Main application ──────────────────────────────────────────────────────────
class RigControlApp:
    def __init__(self, root):
        global _app
        _app = self

        self.root = root
        self.root.title("FT-710 Dashboard")
        self.root.resizable(False, False)
        self.root.configure(bg=BG)

        self.omnirig          = None
        self.rig              = None
        self.events_connected = False
        self._cw_keyed        = False
        self._tx_active       = False
        self._poll_job        = None
        self._tx_start_time   = None
        self._on_top          = False
        self._freq_editing    = False
        self._swr_last_alert  = 0.0
        self._key_change_time = 0.0   # when we last commanded a key/unkey

        self.current_power    = None
        self.current_smeter   = None
        self.current_swr_raw  = None
        self.current_alc_raw  = None
        self.current_split    = False
        self.peak_power       = None
        self.peak_swr         = None

        self._last_freq_hz    = 0
        self._last_band_str   = ""
        self._last_mode_str   = ""

        self.tx_log           = []   # list of dicts, newest first

        self._connect_omnirig()
        self._build_gui()
        self._pump_com()
        self._poll_status()
        self._update_clock()

    # ── OmniRig ───────────────────────────────────────────────────────────────
    def _connect_omnirig(self):
        try:
            self.omnirig = win32com.client.DispatchWithEvents(
                "OmniRig.OmniRigX", OmniRigEvents
            )
            self.rig = self.omnirig.Rig1
            self.events_connected = True
            return
        except Exception as e:
            print(f"DispatchWithEvents failed: {e}")
        try:
            self.omnirig = win32com.client.gencache.EnsureDispatch(
                "OmniRig.OmniRigX"
            )
            self.rig = self.omnirig.Rig1
        except Exception:
            try:
                self.omnirig = win32com.client.Dispatch("OmniRig.OmniRigX")
                self.rig = self.omnirig.Rig1
            except Exception as e:
                messagebox.showerror("OmniRig Error",
                    f"Could not connect to OmniRig Rig 1:\n{e}")

    def _reconnect(self):
        self.omnirig = None
        self.rig = None
        self.events_connected = False
        self._connect_omnirig()
        if self.rig:
            status = "Events connected" if self.events_connected \
                     else "No events — meter readings unavailable"
            messagebox.showinfo("Reconnected",
                f"Connected to OmniRig Rig 1.\n{status}")
        else:
            messagebox.showerror("Failed", "Could not connect to OmniRig.")

    # ── CAT helpers ───────────────────────────────────────────────────────────
    def _send_cat(self, cmd_str):
        if not self.rig:
            return
        try:
            self.rig.SendCustomCommand(cmd_str.encode("ascii"), 0, b"")
        except Exception as e:
            messagebox.showerror("CAT Error", f"SendCustomCommand failed:\n{e}")

    def _query_tx_meters(self):
        if not self.rig or not self.events_connected:
            return
        try:
            self.rig.SendCustomCommand(b"PC;",  6, b";")   # power setting
            self.rig.SendCustomCommand(b"RM6;", 0, b";")   # SWR
            self.rig.SendCustomCommand(b"RM4;", 0, b";")   # ALC
        except Exception:
            pass

    def _query_rx_meters(self):
        if not self.rig or not self.events_connected:
            return
        try:
            # SM reply "SM0nnn;" varies 3-4 digits — read to terminator
            self.rig.SendCustomCommand(b"SM0;", 0, b";")
        except Exception:
            pass

    # ── Lookups ───────────────────────────────────────────────────────────────
    @staticmethod
    def _mode_str(mode_val):
        if not mode_val:
            return "---"
        for mask, name in MODE_NAMES:
            if mode_val & mask:
                return name
        return f"0x{mode_val:08X}"

    @staticmethod
    def _band_str(freq_hz):
        for lo, hi, name in BANDS:
            if lo <= freq_hz <= hi:
                return name
        return ""

    # ── Power / alert getters ─────────────────────────────────────────────────
    def _get_cw_power(self):
        try:
            return max(1, min(100, int(self.cw_power_var.get())))
        except (ValueError, tk.TclError):
            return 49

    def _get_ssb_power(self):
        try:
            return max(1, min(100, int(self.ssb_power_var.get())))
        except (ValueError, tk.TclError):
            return 5

    def _get_swr_alert(self):
        try:
            return max(1.0, min(3.0, float(self.swr_alert_var.get())))
        except (ValueError, tk.TclError):
            return 2.0

    def _update_power_labels(self, *_):
        if not self._cw_keyed:
            self.btn_cw.config(text=f"CW  /  {self._get_cw_power()} W")
        self.btn_ssb.config(text=f"SSB  /  {self._get_ssb_power()} W")

    # ── GUI ───────────────────────────────────────────────────────────────────
    def _build_gui(self):
        # ── Info strip ────────────────────────────────────────────────────────
        info = tk.Frame(self.root, bg=BG2, padx=10, pady=6)
        info.pack(fill="x", padx=12, pady=(10, 4))

        # Row 1: freq (click to edit) | band | mode
        row1 = tk.Frame(info, bg=BG2)
        row1.pack(fill="x")

        freq_cell = tk.Frame(row1, bg=BG2)
        freq_cell.pack(side="left")

        self.lbl_freq = tk.Label(freq_cell, text="---",
                                  font=("Consolas", 15, "bold"),
                                  bg=BG2, fg=FG, cursor="xterm")
        self.lbl_freq.grid(row=0, column=0, sticky="w")
        self.lbl_freq.bind("<Button-1>", self._start_freq_edit)

        self._freq_entry = tk.Entry(
            freq_cell, width=13,
            font=("Consolas", 15, "bold"),
            bg="#252540", fg=FG,
            insertbackground=FG,
            relief="flat",
            highlightthickness=1,
            highlightbackground=BLUE,
        )
        self._freq_entry.grid(row=0, column=0, sticky="w")
        self._freq_entry.grid_remove()
        self._freq_entry.bind("<Return>", self._commit_freq_edit)
        self._freq_entry.bind("<Escape>", lambda e: self._cancel_freq_edit())
        self._freq_entry.bind("<FocusOut>", lambda e: self._cancel_freq_edit())

        self.lbl_band = tk.Label(row1, text="",
                                  font=("Consolas", 13),
                                  bg=BG2, fg="#4499dd")
        self.lbl_band.pack(side="left", padx=(10, 0))

        self.lbl_mode = tk.Label(row1, text="",
                                  font=("Consolas", 13),
                                  bg=BG2, fg=FG_DIM)
        self.lbl_mode.pack(side="left", padx=(10, 0))

        # Row 2: TX dot | TX/RX | timer | UTC | always-on-top | reconnect
        row2 = tk.Frame(info, bg=BG2)
        row2.pack(fill="x", pady=(5, 0))

        self._tx_cv = tk.Canvas(row2, width=14, height=14,
                                 bg=BG2, highlightthickness=0)
        self._tx_cv.pack(side="left")
        self._tx_dot = self._tx_cv.create_oval(1, 1, 13, 13,
                                                fill="#003355",
                                                outline="#224466")

        self._tx_label = tk.Label(row2, text=" RX",
                                   font=("Consolas", 11, "bold"),
                                   bg=BG2, fg=BLUE)
        self._tx_label.pack(side="left")

        self._tx_timer_lbl = tk.Label(row2, text="",
                                       font=("Consolas", 11, "bold"),
                                       bg=BG2, fg=RED)
        self._tx_timer_lbl.pack(side="left", padx=(10, 0))

        self._utc_lbl = tk.Label(row2, text="",
                                  font=("Consolas", 10),
                                  bg=BG2, fg=FG_DIM)
        self._utc_lbl.pack(side="left", padx=(14, 0))

        self._top_btn = tk.Button(row2, text="Pin",
                                   font=("Consolas", 9),
                                   bg="#252545", fg=FG_DIM,
                                   activebackground="#353560",
                                   activeforeground=FG,
                                   relief="flat", bd=0, padx=5, pady=2,
                                   command=self._toggle_topmost)
        self._top_btn.pack(side="right", padx=(4, 0))

        tk.Button(row2, text="Reconnect",
                  font=("Consolas", 9),
                  bg="#252545", fg=FG_DIM,
                  activebackground="#353560",
                  activeforeground=FG,
                  relief="flat", bd=0, padx=5, pady=2,
                  command=self._reconnect,
                  ).pack(side="right")

        # Row 3: VFO B | Swap | Split
        row3 = tk.Frame(info, bg=BG2)
        row3.pack(fill="x", pady=(4, 0))

        tk.Label(row3, text="B:", font=("Consolas", 10),
                 bg=BG2, fg=FG_DIM).pack(side="left")

        self.lbl_vfob = tk.Label(row3, text="---",
                                  font=("Consolas", 12, "bold"),
                                  bg=BG2, fg="#8899bb")
        self.lbl_vfob.pack(side="left", padx=(4, 0))

        self.lbl_vfob_band = tk.Label(row3, text="",
                                       font=("Consolas", 11),
                                       bg=BG2, fg="#4466aa")
        self.lbl_vfob_band.pack(side="left", padx=(8, 0))

        self._split_btn = tk.Button(row3, text="Split: off",
                                     font=("Consolas", 9),
                                     bg="#252545", fg=FG_DIM,
                                     activebackground="#353560",
                                     activeforeground=FG,
                                     relief="flat", bd=0, padx=5, pady=2,
                                     command=self._toggle_split)
        self._split_btn.pack(side="right")

        tk.Button(row3, text="Swap VFO",
                  font=("Consolas", 9),
                  bg="#252545", fg=FG_DIM,
                  activebackground="#353560",
                  activeforeground=FG,
                  relief="flat", bd=0, padx=5, pady=2,
                  command=self._swap_vfo,
                  ).pack(side="right", padx=(0, 4))

        # ── Band buttons ──────────────────────────────────────────────────────
        band_frame = tk.Frame(self.root, bg=BG)
        band_frame.pack(fill="x", padx=12, pady=(2, 4))

        self._band_buttons = []
        btn_kw = dict(
            font=("Consolas", 9),
            bg="#252540", fg=FG_DIM,
            activebackground="#3a3a6c",
            activeforeground=FG,
            relief="flat", bd=0,
            width=4, padx=1, pady=3,
        )
        for i, (label, freq, mode) in enumerate(BAND_DEFAULTS):
            row = i // 6
            col = i % 6
            b = tk.Button(
                band_frame, text=label, **btn_kw,
                command=lambda f=freq, m=mode: self._qsy_band(f, m),
            )
            b.grid(row=row, column=col, padx=1, pady=1, sticky="ew")
            self._band_buttons.append((label, b))

        for col in range(6):
            band_frame.columnconfigure(col, weight=1)

        # ── Dashboard ─────────────────────────────────────────────────────────
        dash = tk.LabelFrame(self.root, text=" Dashboard ",
                              bg=BG, fg=FG_DIM,
                              font=("Consolas", 9),
                              bd=1, relief="solid")
        dash.pack(fill="x", padx=12, pady=4)

        def _smeter_fmt(v):
            v = int(round(v))
            if v <= 9:
                return f"S{v}"
            db = min(60, (v - 9) * 3)
            return f"S9+{db}"

        self.m_power = BarMeter(
            dash, "PWR", 0, 100, unit="W",
            zones=[(70, GREEN), (90, YELLOW), (100, RED)],
        )
        self.m_power.pack(pady=(8, 3))

        self.m_swr = BarMeter(
            dash, "SWR", 1.0, 3.0,
            zones=[(1.5, GREEN), (2.0, YELLOW), (3.0, RED)],
            fmt=lambda v: f"{v:.2f}",
        )
        self.m_swr.pack(pady=3)

        self._swr_alert_lbl = tk.Label(
            dash, text="",
            bg=BG, fg=RED,
            font=("Consolas", 10, "bold"),
        )
        self._swr_alert_lbl.pack()

        self.m_alc = BarMeter(
            dash, "ALC", 0, 100, unit="%",
            zones=[(60, GREEN), (80, YELLOW), (100, RED)],
        )
        self.m_alc.pack(pady=3)

        self.m_smeter = BarMeter(
            dash, "S", 0, 30,
            zones=[(9, GREEN), (18, "#40cc80"), (30, "#80ddaa")],
            fmt=_smeter_fmt,
        )
        self.m_smeter.pack(pady=(3, 8))

        if not self.events_connected:
            tk.Label(dash,
                     text="  Meter readings unavailable — COM events not connected",
                     bg=BG, fg=RED, font=("Consolas", 9),
                     anchor="w").pack(fill="x", padx=8, pady=(0, 6))

        # ── Action buttons ────────────────────────────────────────────────────
        btn_frame = tk.Frame(self.root, bg=BG)
        btn_frame.pack(fill="x", padx=12, pady=(4, 2))

        self.btn_cw = tk.Button(
            btn_frame, text="CW  /  49 W",
            font=("Segoe UI", 13, "bold"),
            bg="#e8b030", fg="black",
            activebackground="#d4a020",
            width=20, height=2,
            command=self._toggle_cw,
        )
        self.btn_cw.pack(pady=4)

        self.btn_ssb = tk.Button(
            btn_frame, text="SSB  /  5 W",
            font=("Segoe UI", 13, "bold"),
            bg="#3090e0", fg="white",
            activebackground="#2070c0",
            width=20, height=2,
            command=self._set_ssb_5w,
        )
        self.btn_ssb.pack(pady=4)

        # ── Settings ──────────────────────────────────────────────────────────
        settings = tk.Frame(self.root, bg=BG2, padx=10, pady=7)
        settings.pack(fill="x", padx=12, pady=(0, 4))

        self.cw_power_var  = tk.IntVar(value=49)
        self.ssb_power_var = tk.IntVar(value=5)
        self.swr_alert_var = tk.StringVar(value="2.0")

        spin_kw = dict(
            font=("Consolas", 10), width=4,
            bg="#252540", fg=FG,
            buttonbackground="#3a3a5c",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#3a3a5c",
            insertbackground=FG,
        )

        tk.Label(settings, text="CW:", font=("Consolas", 10),
                 bg=BG2, fg=FG_DIM).pack(side="left")
        tk.Spinbox(settings, from_=1, to=100,
                   textvariable=self.cw_power_var, **spin_kw
                   ).pack(side="left", padx=(3, 2))
        tk.Label(settings, text="W", font=("Consolas", 10),
                 bg=BG2, fg=FG_DIM).pack(side="left")

        tk.Label(settings, text="  SSB:", font=("Consolas", 10),
                 bg=BG2, fg=FG_DIM).pack(side="left")
        tk.Spinbox(settings, from_=1, to=100,
                   textvariable=self.ssb_power_var, **spin_kw
                   ).pack(side="left", padx=(3, 2))
        tk.Label(settings, text="W", font=("Consolas", 10),
                 bg=BG2, fg=FG_DIM).pack(side="left")

        tk.Label(settings, text="  SWR alert:", font=("Consolas", 10),
                 bg=BG2, fg=FG_DIM).pack(side="left")
        tk.Spinbox(settings, from_=1.0, to=3.0, increment=0.1,
                   format="%.1f",
                   textvariable=self.swr_alert_var, **spin_kw
                   ).pack(side="left", padx=(3, 2))


        self.cw_power_var.trace_add("write", self._update_power_labels)
        self.ssb_power_var.trace_add("write", self._update_power_labels)

        # ── TX Log ────────────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(self.root, text=" TX Log ",
                                   bg=BG, fg=FG_DIM,
                                   font=("Consolas", 9),
                                   bd=1, relief="solid")
        log_frame.pack(fill="x", padx=12, pady=(0, 12))

        log_top = tk.Frame(log_frame, bg=BG)
        log_top.pack(fill="x", padx=6, pady=(4, 2))
        tk.Button(log_top, text="Clear",
                  font=("Consolas", 9),
                  bg="#252540", fg=FG_DIM,
                  activebackground="#3a3a5c",
                  relief="flat", bd=0, padx=5, pady=1,
                  command=self._clear_tx_log,
                  ).pack(side="right")

        self.tx_log_text = tk.Text(
            log_frame, height=5,
            bg="#0f0f1e", fg=FG,
            font=("Consolas", 9),
            state="disabled",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#3a3a5c",
            insertbackground=FG,
            selectbackground="#3a3a5c",
        )
        self.tx_log_text.pack(fill="x", padx=6, pady=(0, 6))

    # ── Frequency click-to-edit ───────────────────────────────────────────────
    def _start_freq_edit(self, event=None):
        if self._tx_active:
            return  # don't allow freq edit while transmitting
        self._freq_editing = True
        self._freq_entry.delete(0, "end")
        self._freq_entry.insert(0, f"{self._last_freq_hz / 1_000_000:.6f}")
        self._freq_entry.select_range(0, "end")
        self.lbl_freq.grid_remove()
        self._freq_entry.grid()
        self._freq_entry.focus_set()

    def _cancel_freq_edit(self):
        if not self._freq_editing:
            return
        self._freq_editing = False
        self._freq_entry.grid_remove()
        self.lbl_freq.grid()

    def _commit_freq_edit(self, event=None):
        text = self._freq_entry.get().strip()
        freq_hz = self._parse_freq_mhz(text)
        self._cancel_freq_edit()
        if freq_hz is None or not (100_000 < freq_hz < 450_000_000):
            return
        self._send_cat(f"FA{freq_hz:09d};")   # FT-710 FA is 9 digits
        self.root.after(150, self._update_display)

    @staticmethod
    def _parse_freq_mhz(text):
        text = text.replace(",", "").strip()
        try:
            val = float(text)
        except ValueError:
            return None
        if "." in text or val < 1000:
            return int(val * 1_000_000)     # MHz
        elif val < 100_000:
            return int(val * 1_000)          # kHz
        return int(val)                       # Hz

    # ── Band QSY ──────────────────────────────────────────────────────────────
    def _qsy_band(self, freq_hz, mode_cmd):
        if not self.rig:
            messagebox.showwarning("No Rig", "OmniRig is not connected.")
            return
        if self._tx_active:
            return  # don't QSY mid-transmission
        # OmniRig's Freq property write is ignored on this rig — use CAT FA.
        self._send_cat(f"FA{freq_hz:09d};")   # FT-710 FA is 9 digits
        self._send_cat(mode_cmd)
        self.root.after(150, self._update_display)

    # ── VFO / Split ───────────────────────────────────────────────────────────
    def _swap_vfo(self):
        self._send_cat("SV;")
        self.root.after(300, self._update_display)

    def _toggle_split(self):
        if self.current_split:
            self._send_cat("FT0;")
        else:
            self._send_cat("FT1;")
        self.root.after(200, self._update_display)

    # ── Always on top ─────────────────────────────────────────────────────────
    def _toggle_topmost(self):
        self._on_top = not self._on_top
        self.root.attributes("-topmost", self._on_top)
        if self._on_top:
            self._top_btn.config(text="Pinned", fg=BLUE,
                                  bg="#1a2a4e", activeforeground=BLUE)
        else:
            self._top_btn.config(text="Pin", fg=FG_DIM,
                                  bg="#252545", activeforeground=FG)

    # ── SWR alert ─────────────────────────────────────────────────────────────
    @staticmethod
    def _beep():
        try:
            winsound.Beep(1200, 350)
        except Exception:
            pass

    def _check_swr_alert(self, swr):
        threshold = self._get_swr_alert()
        if swr > threshold:
            self._swr_alert_lbl.config(text=f"  HIGH SWR: {swr:.2f}")
            now = time.time()
            if _HAS_WINSOUND and now - self._swr_last_alert > 5.0:
                self._swr_last_alert = now
                threading.Thread(target=self._beep, daemon=True).start()
        else:
            self._swr_alert_lbl.config(text="")

    # ── TX log ────────────────────────────────────────────────────────────────
    def _add_tx_log_entry(self):
        if not self._tx_start_time:
            return
        duration = int(time.time() - self._tx_start_time)
        entry = {
            "freq": self._last_freq_hz,
            "band": self._last_band_str,
            "mode": self._last_mode_str,
            "power": self.peak_power,
            "swr":   self.peak_swr,
            "secs":  duration,
        }
        self.tx_log.insert(0, entry)
        if len(self.tx_log) > 10:
            self.tx_log.pop()
        self._refresh_tx_log()

    def _refresh_tx_log(self):
        self.tx_log_text.config(state="normal")
        self.tx_log_text.delete("1.0", "end")
        for e in self.tx_log:
            m, s     = divmod(e["secs"], 60)
            freq_str = f"{e['freq'] / 1_000_000:10.6f}"
            band_str = f"{e['band']:4s}"
            mode_str = f"{e['mode']:5s}"
            pwr_str  = f"{e['power']:3d}W" if e["power"] is not None else " --W"
            swr_str  = f"{e['swr']:.2f}" if e["swr"] is not None else " --"
            time_str = f"{m}:{s:02d}"
            self.tx_log_text.insert(
                "end",
                f"{freq_str}  {band_str}  {mode_str}  {pwr_str}  SWR:{swr_str}  {time_str}\n"
            )
        self.tx_log_text.config(state="disabled")

    def _clear_tx_log(self):
        self.tx_log.clear()
        self._refresh_tx_log()

    # ── UTC clock (independent 1-second loop) ─────────────────────────────────
    def _update_clock(self):
        t = time.gmtime()
        self._utc_lbl.config(
            text=f"{t.tm_hour:02d}:{t.tm_min:02d}:{t.tm_sec:02d} UTC"
        )
        self.root.after(1000, self._update_clock)

    # ── Actions ───────────────────────────────────────────────────────────────
    def _toggle_cw(self):
        if not self.rig:
            messagebox.showwarning("No Rig", "OmniRig is not connected.")
            return
        try:
            self._key_change_time = time.time()
            if not self._cw_keyed:
                cw_p = self._get_cw_power()
                self._send_cat("MD03;")
                self._send_cat(f"PC{cw_p:03d};")
                self.root.after(150, lambda: self._send_cat("TX1;"))
                self._cw_keyed = True
                self.btn_cw.config(text="TX  (click to unkey)",
                                   bg="#cc2222", fg="white",
                                   activebackground="#aa1111")
                self._set_tx(True)
            else:
                self._send_cat("TX0;")
                self._cw_keyed = False
                self.btn_cw.config(text=f"CW  /  {self._get_cw_power()} W",
                                   bg="#e8b030", fg="black",
                                   activebackground="#d4a020")
                self._set_tx(False)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _set_ssb_5w(self):
        if not self.rig:
            messagebox.showwarning("No Rig", "OmniRig is not connected.")
            return
        try:
            if self._cw_keyed:
                self._key_change_time = time.time()
                self._send_cat("TX0;")
                self._cw_keyed = False
                self.btn_cw.config(text=f"CW  /  {self._get_cw_power()} W",
                                   bg="#e8b030", fg="black",
                                   activebackground="#d4a020")
                self._set_tx(False)
            freq = self.rig.Freq
            if freq >= 10_000_000 or 5_000_000 <= freq <= 5_500_000:
                self._send_cat("MD02;")
            else:
                self._send_cat("MD01;")
            self._send_cat(f"PC{self._get_ssb_power():03d};")
            self.root.after(300, self._update_display)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ── TX state management ───────────────────────────────────────────────────
    def _set_tx(self, active):
        was_tx = self._tx_active
        self._tx_active = active
        if active and not was_tx:
            self._tx_start_time = time.time()
            self.current_smeter = None
            self._swr_alert_lbl.config(text="")
        elif not active and was_tx:
            self._add_tx_log_entry()
            self._tx_start_time  = None
            self.current_power   = None
            self.current_swr_raw = None
            self.current_alc_raw = None
            self.peak_power      = None
            self.peak_swr        = None
            self.m_power.set_peak(None)
            self._swr_alert_lbl.config(text="")

    # ── Status refresh ────────────────────────────────────────────────────────
    def _update_display(self):
        if not self.rig:
            return

        # Frequency + Band
        try:
            freq_hz = self.rig.Freq
            self._last_freq_hz = freq_hz
            band = self._band_str(freq_hz)
            self._last_band_str = band
            if not self._freq_editing:
                self.lbl_freq.config(text=f"{freq_hz / 1_000_000:,.6f} MHz")
            self.lbl_band.config(text=band)
        except Exception:
            if not self._freq_editing:
                self.lbl_freq.config(text="---")
            self.lbl_band.config(text="")

        # Mode
        try:
            mode = self._mode_str(self.rig.Mode)
            self._last_mode_str = mode
            self.lbl_mode.config(text=mode)
        except Exception:
            self.lbl_mode.config(text="---")

        # VFO B — read OmniRig's native FreqB property (clean integer Hz)
        try:
            vfob = int(self.rig.FreqB)
        except Exception:
            vfob = 0
        if vfob:
            self.lbl_vfob.config(text=f"{vfob / 1_000_000:,.6f} MHz")
            self.lbl_vfob_band.config(text=self._band_str(vfob))
        else:
            self.lbl_vfob.config(text="---")
            self.lbl_vfob_band.config(text="")

        # Split state — read OmniRig's native Split property
        try:
            self.current_split = bool(int(self.rig.Split) & PM_SPLIT_ON)
        except Exception:
            pass

        # Split button label
        if self.current_split:
            self._split_btn.config(text="Split: on",
                                   fg=YELLOW, bg="#2a2a10")
        else:
            self._split_btn.config(text="Split: off",
                                   fg=FG_DIM, bg="#252545")

        # Band button highlight
        for label, btn in self._band_buttons:
            if label == self._last_band_str:
                btn.config(bg="#2a2a6c", fg=FG)
            else:
                btn.config(bg="#252540", fg=FG_DIM)

        # TX/RX indicator. Just after a button-driven key/unkey, trust our own
        # intent (_cw_keyed) so the rig lagging by a poll doesn't bounce the
        # state. Once settled, the rig's Tx bit is authoritative — that's what
        # catches an external footswitch/VOX.
        tx = self._cw_keyed
        try:
            if time.time() - self._key_change_time < 0.7:
                tx = self._cw_keyed
            else:
                tx = bool(int(self.rig.Tx) & PM_TX)
        except Exception:
            tx = self._cw_keyed
        if tx != self._tx_active:
            self._set_tx(tx)

        if tx:
            self._tx_cv.itemconfig(self._tx_dot, fill=RED, outline="#ff5555")
            self._tx_label.config(text=" TX", fg=RED)
        else:
            self._tx_cv.itemconfig(self._tx_dot, fill="#003355", outline="#224466")
            self._tx_label.config(text=" RX", fg=BLUE)

        # TX timer
        if self._tx_active and self._tx_start_time is not None:
            elapsed = int(time.time() - self._tx_start_time)
            m, s = divmod(elapsed, 60)
            self._tx_timer_lbl.config(text=f"{m:02d}:{s:02d}")
        else:
            self._tx_timer_lbl.config(text="")

        # Peak power tracking
        if self.current_power is not None:
            if self.peak_power is None or self.current_power > self.peak_power:
                self.peak_power = self.current_power
            self.m_power.set_peak(self.peak_power)

        # Meters
        self.m_power.set_value(self.current_power)

        swr_val = (None if self.current_swr_raw is None
                   else swr_from_raw(self.current_swr_raw))
        self.m_swr.set_value(swr_val)

        if swr_val is not None and self._tx_active:
            if self.peak_swr is None or swr_val > self.peak_swr:
                self.peak_swr = swr_val
            self._check_swr_alert(swr_val)
        else:
            self._swr_alert_lbl.config(text="")

        self.m_alc.set_value(
            None if self.current_alc_raw is None
            else self.current_alc_raw * 100 / 255
        )
        # S-meter raw (0-255) mapped to gauge via the Yaesu calibration curve
        self.m_smeter.set_value(
            None if self.current_smeter is None
            else raw_to_smeter(self.current_smeter)
        )

    def _pump_com(self):
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        self.root.after(100, self._pump_com)

    def _poll_status(self):
        # Refresh the display every cycle (property reads are cheap, so freq/
        # mode/band stay responsive). Meter CAT queries run each cycle too —
        # fast while transmitting, a bit slower on receive.
        if self._tx_active:
            self._query_tx_meters()
            interval = 200
        else:
            self._query_rx_meters()
            interval = 300
        self._update_display()
        self._poll_job = self.root.after(interval, self._poll_status)


if __name__ == "__main__":
    root = tk.Tk()
    app = RigControlApp(root)
    root.mainloop()
