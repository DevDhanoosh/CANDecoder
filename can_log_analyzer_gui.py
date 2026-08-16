#!/usr/bin/env python3
"""
CAN Signal Bench — desktop GUI for can_log_analyzer.

A graphical front-end mirroring the HTML tool, plus a Bus Health workspace
(per-ID timing, estimated bus load %) and a Vehicle Dynamics workspace that
derives regen recovery, energy/power consumption, efficiency, distance and
more from the decoded signals.

Load up to 10 DBCs and a CAN trace (BUSMASTER .log/.asc or IXXAT MiniMon .csv),
tick the signals you want, browse stats and plots, trim by clock time or a drag
on the plot, map signals to physical roles for the dynamics report, choose which
sheets to export, and save a decoded .xlsx.

Reuses decode / parse / export logic from can_log_analyzer.py — keep both files
in the SAME folder.

Run:  python can_log_analyzer_gui.py
Deps: numpy, matplotlib, openpyxl (Excel).  tkinter ships with Python on
      Windows/macOS; on Linux install python3-tk.
"""

import inspect
import os
import re
import sys
import threading
import traceback

# A windowed build (or pythonw.exe / a frozen console=False .exe) has no
# console attached, so sys.stdout/stderr/stdin are None. This file and
# can_log_analyzer.py print/write throughout — including Tkinter's own
# default callback-exception reporter and the worker threads' traceback
# dumps below — so redirect to a null sink up front instead of guarding
# every call site.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import can_log_analyzer as core
except Exception as exc:                                   # pragma: no cover
    msg = ("Could not import can_log_analyzer.py.\n\nKeep can_log_analyzer_gui.py "
           "in the SAME folder as can_log_analyzer.py.\n\nDetails: " + str(exc))
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        _r = _tk.Tk(); _r.withdraw(); _mb.showerror("CAN Signal Bench", msg)
    except Exception:
        print(msg)
    sys.exit(1)

core.tqdm = lambda iterable=None, **_: (iterable if iterable is not None else [])

try:
    import numpy as np
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    from matplotlib.widgets import SpanSelector
except Exception as exc:                                   # pragma: no cover
    msg = ("Missing a dependency: " + str(exc) +
           "\n\nInstall with:  pip install numpy matplotlib openpyxl   (python3-tk on Linux)")
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        _r = _tk.Tk(); _r.withdraw(); _mb.showerror("CAN Signal Bench", msg)
    except Exception:
        print(msg)
    sys.exit(1)


PALETTE = ["#f0a92a", "#2bb8d8", "#e05cc8", "#46c66a", "#5aa9ff", "#ff6d5e",
           "#9b8cff", "#c9cf3a"]
THEMES = {
    "dark": dict(bg="#0d1219", panel="#171e2b", well="#10161f", ink="#eaf0f7",
                 muted="#8ea0b5", faint="#5f7086", accent="#4c9ffa", accent2="#2a6fd6",
                 line="#2b3748", grid="#223143", entry="#0c1119", sel="#1d4e86",
                 head="#111825", stripe="#141c28"),
    "light": dict(bg="#eaeef3", panel="#ffffff", well="#f4f7fb", ink="#16202e",
                  muted="#586a7e", faint="#8595a7", accent="#0b63d6", accent2="#0a4fac",
                  line="#d7dee7", grid="#e4eaf1", entry="#ffffff", sel="#cfe2ff",
                  head="#f0f3f8", stripe="#f6f9fc"),
}

# physical roles for the vehicle-dynamics calculations, with name hints.
# hints are matched separator-insensitively (see _norm), most specific first,
# so the battery signals win over DC-DC / OBC currents on multi-node DBCs.
ROLES = [
    ("speed",        "Vehicle speed",   ["vehiclespeed", "vehspeed", "hudspeed", "vspeed", "vspd", "speed"]),
    ("batt_volt",    "Battery voltage", ["batteryvoltage", "packvoltage", "battvolt", "hvvolt", "dcvolt", "vbat", "vdc", "voltage"]),
    ("batt_curr",    "Battery current", ["batterycurrent", "packcurrent", "battcurr", "hvcurr", "ibat", "idc", "current"]),
    ("motor_torque", "Motor torque",    ["motoractualtorque", "actualtorque", "motortorque", "torque", "trq"]),
    ("motor_speed",  "Motor speed",     ["motorrpm", "motorspeed", "rotorspeed", "motspeed", "rpm"]),
    ("soc",          "State of charge", ["batterysoc", "soc", "stateofcharge"]),
    ("odometer",     "Odometer",        ["totaldistancekm", "totaldistance", "distancetravelled", "odometer", "odo"]),
]
SECTIONS = [
    ("bus_health", "Bus Health (per-ID timing, load %)"),
    ("dynamics",   "Vehicle Dynamics"),
    ("components", "Components (per-subsystem)"),
    ("faults",     "Faults (episodes)"),
    ("summary",    "Summary (stats)"),
    ("frames",     "CAN Frames (raw trace)"),
    ("merged",     "Merged (time-aligned, all DBC signals)"),
    ("charts",     "Charts (embedded plots, selected signals)"),
]
FAULT_HINTS = ["fault", "error", "dtc", "trip", "warn", "fail", "alarm",
               "malfunction", "overtemp", "overcurrent", "overvolt", "undervolt"]

# component / subsystem classification (checked in order; first match wins).
# matched separator-insensitively against "<message name> <signal name>".
COMPONENTS = [
    ("EPAS",    ["epas", "eps", "steer", "sas", "handwheel", "column", "torquesensor", "steeringangle"]),
    ("BMS",     ["bms", "battery", "cell", "pack", "soc", "amphour", "contactor", "insulation", "isolation"]),
    ("DC-DC",   ["dcdc", "dclink", "dc_dc", "converter", "lvbatt", "lvdc", "auxbatt", "12v", "lowvoltage"]),
    ("OBC",     ["obc", "charger", "charging", "chargeport", "evse"]),
    ("MCU",     ["mcu", "mc_", "motor", "mtr", "inverter", "igbt", "stator", "rotor", "phase", "torque", "rpm", "statusmotor"]),
    ("Vehicle", ["vcu", "hud", "vehicle", "speed", "brake", "throttle", "accel", "pedal", "gear",
                 "drive", "odo", "distance", "door", "light", "horn", "wiper", "hvac", "cluster"]),
]
COMPONENT_ORDER = [c for c, _ in COMPONENTS] + ["Other"]


def classify_component(msg_name, sig_name):
    hay = _norm(f"{msg_name} {sig_name}")
    for comp, keys in COMPONENTS:
        if any(_norm(k) in hay for k in keys):
            return comp
    return "Other"


def decimate(t, v, max_n=6000):
    if len(t) <= max_n:
        return t, v
    step = int(np.ceil(len(t) / max_n))
    return t[::step], v[::step]


def _trapz(y, x):
    return float(np.trapezoid(y, x)) if len(x) > 1 else 0.0


def detect_faults(items, mode, threshold, max_per_signal=2000):
    """
    Accumulate fault episodes from one or more signals.

    items     : list of (name, t_ndarray, v_ndarray)  (already trimmed)
    mode       : 'code'  -> active when value != 0 (episode key = the code value)
                 'flag'  -> active when value >= threshold (episode key = 1)
    threshold  : float, used only in 'flag' mode
    max_per_signal : signals that produce more episodes than this are treated as
                 continuous/counter signals (not discrete faults) and skipped.
    returns    : (episodes, skipped) where episodes is a list of
                 (name, code, start_abs, end_abs) sorted by start time, and
                 skipped is the list of signal names that blew the cap.
    """
    out, skipped = [], []
    for name, t, v in items:
        if len(t) == 0:
            continue
        if mode == "flag":
            codes = np.where(v >= threshold, 1, 0).astype(int)
        else:
            codes = np.rint(v).astype(int)          # 0 == inactive, else fault code
        eps, prev, start = [], 0, None
        for k in range(len(t)):
            code = int(codes[k])
            if code != prev:
                if prev != 0 and start is not None:
                    eps.append((name, prev, start, float(t[k])))
                start = float(t[k]) if code != 0 else None
                prev = code
                if len(eps) > max_per_signal:
                    break
        if len(eps) > max_per_signal:
            skipped.append(name)
            continue
        if prev != 0 and start is not None:
            eps.append((name, prev, start, float(t[-1])))
        out.extend(eps)
    out.sort(key=lambda r: r[2])
    return out, skipped


def _norm(s):
    return re.sub(r"[^a-z0-9]", "", s.lower())


# tokens that mark a signal as a setpoint/limit/counter rather than a live
# feedback value — de-prioritised in auto-mapping (still used if nothing else)
_AVOID = ("count", "cou", "limit", "demand", "target", "request",
          "setpoint", "maximum", "minimum", "estimate")


def guess_roles(names):
    """Map physical roles to signal names, separator-insensitively.
    Most-specific hints win; battery signals beat DC-DC / OBC currents; and
    live feedback signals beat counters/limits/demands/targets. Returns
    dict role_id -> name (or '— none —')."""
    norm = [(n, _norm(n)) for n in names]
    used, out = set(), {}
    for rid, _, hints in ROLES:
        pick = "— none —"
        for skip_avoid in (True, False):          # first pass avoids setpoints
            for hint in hints:
                h = _norm(hint)
                for n, nn in norm:
                    if n in used or h not in nn:
                        continue
                    if skip_avoid and any(a in nn for a in _AVOID):
                        continue
                    pick = n; used.add(n); break
                if pick != "— none —":
                    break
            if pick != "— none —":
                break
        out[rid] = pick
    return out


def compute_dynamics(roles, opts):
    """
    roles : dict role -> (t_ndarray, v_ndarray, unit_str)  (already trimmed)
    opts  : dict(speed_unit, mass, payload)
    returns (rows, raw, timeseries):
       rows       = list of (metric, value_str, unit)
       raw        = dict of numeric results
       timeseries = (grid_abs, power_w, cum_energy_wh) or None
    """
    present = {k: v for k, v in roles.items() if v is not None}
    if not present:
        return [("Map at least one signal to a role.", "", "")], {}, None

    t_lo = max(v[0].min() for v in present.values())
    t_hi = min(v[0].max() for v in present.values())
    if t_hi <= t_lo:
        return [("Mapped signals have no overlapping time range.", "", "")], {}, None
    grid = np.unique(np.concatenate([v[0] for v in present.values()]))
    grid = grid[(grid >= t_lo) & (grid <= t_hi)]
    if len(grid) < 2:
        return [("Not enough overlapping samples.", "", "")], {}, None

    def g(role):
        if role not in present:
            return None
        t, v, u = present[role]
        return np.interp(grid, t, v), (u or "")

    rows, raw = [], {}
    rows.append(("Window duration", f"{float(grid[-1]-grid[0]):.2f}", "s"))

    # ── electrical power = battery voltage × current (sign is in the CAN data)
    power = None
    v = g("batt_volt"); i = g("batt_curr")
    if v is not None and i is not None:
        power = v[0] * i[0]                             # +ve = discharge (per BMS sign)
    if power is not None:
        edis = _trapz(np.clip(power, 0, None), grid) / 3600.0
        ereg = -_trapz(np.clip(power, None, 0), grid) / 3600.0
        enet = _trapz(power, grid) / 3600.0
        pos, neg = power[power > 0], power[power < 0]
        used_avg = float(pos.mean()) / 1000.0 if pos.size else 0.0     # kW, discharge only
        regen_avg = -float(neg.mean()) / 1000.0 if neg.size else 0.0   # kW, regen only
        peak_dis = max(0.0, float(power.max())) / 1000.0
        peak_reg = max(0.0, -float(power.min())) / 1000.0
        raw.update(discharge_wh=edis, regen_wh=ereg, net_wh=enet,
                   used_avg_kw=used_avg, regen_avg_kw=regen_avg)
        rows += [
            ("Energy consumed (discharge)", f"{edis:.1f}", "Wh"),
            ("Energy recovered (regen)", f"{ereg:.1f}", "Wh"),
            ("Net battery energy", f"{enet:.1f}", "Wh"),
            ("Regen recovery", f"{(ereg/edis*100):.1f}" if edis > 0 else "—", "%"),
            ("Used power (avg discharge)", f"{used_avg:.2f}", "kW"),
            ("Regen power (avg)", f"{regen_avg:.2f}", "kW"),
            ("Peak discharge power", f"{peak_dis:.2f}", "kW"),
            ("Peak regen power", f"{peak_reg:.2f}", "kW"),
            ("Average net power", f"{float(power.mean())/1000:.2f}", "kW"),
        ]

    # ── speed / distance / accel ────────────────────────────────────────────
    sp = g("speed"); dist_km = None
    if sp is not None:
        kmh = sp[0] * 3.6 if opts.get("speed_unit") == "m/s" else sp[0]
        ms = kmh / 3.6
        dist_km = _trapz(ms, grid) / 1000.0
        raw["distance_km"] = dist_km
        accel = np.gradient(ms, grid)
        rows += [
            ("Distance travelled", f"{dist_km:.3f}", "km"),
            ("Average speed", f"{float(kmh.mean()):.1f}", "km/h"),
            ("Max speed", f"{float(kmh.max()):.1f}", "km/h"),
            ("Max acceleration", f"{float(accel.max()):.2f}", "m/s²"),
            ("Max deceleration", f"{float(accel.min()):.2f}", "m/s²"),
        ]

    # ── efficiency ──────────────────────────────────────────────────────────
    if power is not None and dist_km and dist_km > 1e-6:
        eff = raw["net_wh"] / dist_km
        raw["efficiency_wh_km"] = eff
        rows.append(("Energy efficiency", f"{eff:.1f}", "Wh/km"))
        rows.append(("Range per kWh", f"{1000.0/eff:.1f}" if eff > 0 else "—", "km/kWh"))

    # ── motor mechanical power ──────────────────────────────────────────────
    trq = g("motor_torque"); rpm = g("motor_speed")
    if trq is not None and rpm is not None:
        pmech = trq[0] * rpm[0] * np.pi / 30.0            # W  (τ·ω, ω = rpm·2π/60)
        emech = _trapz(np.clip(pmech, 0, None), grid) / 3600.0
        rows += [
            ("Motor mechanical energy", f"{emech:.1f}", "Wh"),
            ("Peak mechanical power", f"{float(pmech.max())/1000:.2f}", "kW"),
        ]
        if power is not None and raw.get("discharge_wh", 0) > 0:
            rows.append(("Drivetrain efficiency (mech/elec)",
                         f"{emech/raw['discharge_wh']*100:.1f}", "%"))

    # ── SoC ─────────────────────────────────────────────────────────────────
    soc = g("soc")
    if soc is not None:
        rows.append(("SoC change", f"{float(soc[0][-1]-soc[0][0]):+.1f}", "%"))
        rows.append(("SoC range", f"{float(soc[0].min()):.1f} – {float(soc[0].max()):.1f}", "%"))

    # ── kinetic energy (needs mass + speed) ─────────────────────────────────
    gross = (opts.get("mass") or 0.0) + (opts.get("payload") or 0.0)
    if gross and sp is not None:
        ms_peak = float((kmh.max()) / 3.6)
        rows.append(("Peak kinetic energy", f"{0.5*gross*ms_peak**2/3600:.1f}", "Wh"))

    ts = None
    if power is not None:
        cum = np.concatenate([[0.0], np.cumsum((power[1:] + power[:-1]) / 2 * np.diff(grid))]) / 3600.0
        ts = (grid, power, cum)
    return rows, raw, ts


class App:
    def __init__(self, root):
        self.root = root
        self.theme = "dark"
        self.dbc_paths, self.log_path = [], None
        self.messages, self.frames, self.meta = [], [], None
        self.frame_name_by_id, self.series_all = {}, []
        self.log_t0 = self.log_dur = 0.0
        self.trim_a = self.trim_b = None
        self.span = None
        self.busy = False
        self.sig_vars = []
        self.fault_vars = []
        self.dyn_rows = []
        self.fault_rows = []
        self.comp_groups = {}
        self._dyn_ts = None
        self.bus_health = None

        root.title("CAN Signal Bench")
        root.geometry("1300x860")
        root.minsize(1080, 700)
        # Tkinter's default handler for exceptions raised inside a widget
        # callback just prints the traceback — which crashes outright in a
        # windowed build (no console, sys.stderr redirected to devnull isn't
        # enough on its own to give the USER any feedback). Route it to a
        # real error dialog instead.
        root.report_callback_exception = self._on_callback_exception

        self.style = ttk.Style()
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._build_ui()
        self.apply_theme(self.theme)
        self._set_loginfo({})
        self.log("Ready.  Add DBC file(s) and a log, then Analyze.")
        self._check_core()

    def _check_core(self):
        try:
            params = inspect.signature(core.export_excel).parameters
        except (ValueError, TypeError):
            return
        missing = [k for k in ("components", "progress", "faults", "dynamics", "bus_health") if k not in params]
        if "analyze_bus_health" not in dir(core):
            missing.append("analyze_bus_health")
        if missing:
            self.log("⚠ can_log_analyzer.py in this folder looks OUTDATED "
                     "(missing: " + ", ".join(missing) + ").")
            self.log("  Replace it with the latest version — make sure the file "
                     "is named exactly 'can_log_analyzer.py' (no ' (1)', ' (2)' suffix).")

    # ── layout ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        self.header = tk.Frame(self.root, height=54)
        self.header.pack(fill="x"); self.header.pack_propagate(False)
        self.h_title = tk.Label(self.header, text="  CAN Signal Bench",
                                font=("Segoe UI", 17, "bold"), anchor="w")
        self.h_title.pack(side="left", padx=6)
        self.h_sub = tk.Label(self.header, text="DBC-guided decode · dynamics · report",
                              font=("Segoe UI", 10))
        self.h_sub.pack(side="left", padx=8)
        self.theme_btn = ttk.Button(self.header, text="◐ Light", width=10,
                                    style="Ghost.TButton", command=self.toggle_theme)
        self.theme_btn.pack(side="right", padx=14)

        body = ttk.Frame(self.root, style="TFrame")
        body.pack(fill="both", expand=True, padx=14, pady=12)
        main = ttk.Panedwindow(body, orient="horizontal")
        main.pack(fill="both", expand=True)
        self._main = main

        left = ttk.Frame(main, style="Card.TFrame", padding=14)
        main.add(left, weight=0); self._build_left(left)
        right = ttk.Frame(main, style="TFrame")
        main.add(right, weight=1); self._build_right(right)
        self.root.after(80, lambda: self._main.sashpos(0, 360))

    def _card_head(self, parent, text):
        ttk.Label(parent, text=text, style="Head.TLabel").pack(anchor="w", pady=(0, 6))

    def _build_left(self, p):
        p.configure(width=350)
        self._card_head(p, "SOURCE FILES")
        ttk.Button(p, text="＋  Add DBC file(s)", command=self.add_dbc).pack(fill="x", pady=(0, 6))
        self.dbc_list = tk.Listbox(p, height=4, width=30, activestyle="none", exportselection=False,
                                   relief="flat", highlightthickness=1)
        self.dbc_list.pack(fill="x")
        row = ttk.Frame(p, style="Card.TFrame"); row.pack(fill="x", pady=(4, 12))
        ttk.Button(row, text="Remove", style="Ghost.TButton", command=self.remove_dbc).pack(side="left")
        self.dbc_count = ttk.Label(row, text="0 / 10", style="Dim.TLabel"); self.dbc_count.pack(side="right")

        ttk.Button(p, text="🗎  Choose log trace", command=self.choose_log).pack(fill="x")
        self.log_lbl = ttk.Label(p, text="no log selected", style="Dim.TLabel")
        self.log_lbl.pack(anchor="w", pady=(4, 12))

        self.analyze_btn = ttk.Button(p, text="Analyze  ▶", style="Accent.TButton", command=self.analyze)
        self.analyze_btn.pack(fill="x", pady=(0, 14))

        self._card_head(p, "LOG INFO")
        info = ttk.Frame(p, style="Card.TFrame"); info.pack(fill="x", pady=(0, 14))
        self.info_vals = {}
        for i, k in enumerate(["Format", "Date", "Start", "End", "Duration", "Frames", "Signals"]):
            ttk.Label(info, text=k, style="Dim.TLabel").grid(row=i, column=0, sticky="w", pady=1)
            v = ttk.Label(info, text="—", style="Mono.TLabel"); v.grid(row=i, column=1, sticky="w", padx=(12, 0))
            self.info_vals[k] = v

        self._card_head(p, "CONSOLE")
        self.console = tk.Text(p, height=8, width=34, wrap="word", state="disabled",
                               relief="flat", font=("Consolas", 9), highlightthickness=1)
        self.console.pack(fill="both", expand=True)

    def _build_right(self, p):
        trim = ttk.Frame(p, style="Card.TFrame", padding=10); trim.pack(fill="x")
        ttk.Label(trim, text="TRIM", style="Head.TLabel").grid(row=0, column=0, padx=(0, 8))
        self.trim_mode = ttk.Combobox(trim, values=["Clock time", "Seconds"], width=11, state="readonly")
        self.trim_mode.current(0); self.trim_mode.grid(row=0, column=1)
        self.trim_mode.bind("<<ComboboxSelected>>", lambda e: self._reformat_trim_fields())
        self.trim_start = ttk.Entry(trim, width=13); self.trim_start.grid(row=0, column=2, padx=4)
        ttk.Label(trim, text="–", style="Card.TLabel").grid(row=0, column=3)
        self.trim_end = ttk.Entry(trim, width=13); self.trim_end.grid(row=0, column=4, padx=4)
        ttk.Button(trim, text="Apply", style="Ghost.TButton", command=self.apply_trim_fields).grid(row=0, column=5, padx=(6, 2))
        ttk.Button(trim, text="Reset", style="Ghost.TButton", command=self.reset_trim).grid(row=0, column=6)
        self.trim_info = ttk.Label(trim, text="", style="MonoDim.TLabel"); self.trim_info.grid(row=0, column=7, padx=(12, 0))
        for e in (self.trim_start, self.trim_end):
            e.bind("<Return>", lambda ev: self.apply_trim_fields())

        self.nb = ttk.Notebook(p); self.nb.pack(fill="both", expand=True, pady=(10, 0))
        t1 = ttk.Frame(self.nb, style="TFrame"); self.nb.add(t1, text="Signals"); self._build_signals_tab(t1)
        t2 = ttk.Frame(self.nb, style="TFrame"); self.nb.add(t2, text="Traces"); self._build_traces_tab(t2)
        t_bus = ttk.Frame(self.nb, style="TFrame"); self.nb.add(t_bus, text="Bus Health"); self._build_bus_health_tab(t_bus)
        t3 = ttk.Frame(self.nb, style="TFrame"); self.nb.add(t3, text="Overlay"); self._build_overlay_tab(t3)
        t4 = ttk.Frame(self.nb, style="TFrame"); self.nb.add(t4, text="Dynamics"); self._build_dynamics_tab(t4)
        t5 = ttk.Frame(self.nb, style="TFrame"); self.nb.add(t5, text="Components"); self._build_components_tab(t5)
        t6 = ttk.Frame(self.nb, style="TFrame"); self.nb.add(t6, text="Faults"); self._build_faults_tab(t6)
        t7 = ttk.Frame(self.nb, style="TFrame"); self.nb.add(t7, text="Report"); self._build_report_tab(t7)

        bar = ttk.Frame(p, style="TFrame"); bar.pack(fill="x", pady=(8, 0))
        self.status = ttk.Label(bar, text="", style="Dim.TLabel"); self.status.pack(side="left")
        self.progress = ttk.Progressbar(bar, mode="determinate", maximum=100, length=240,
                                        style="Green.Horizontal.TProgressbar")
        self.progress.pack(side="right")
        self.progress.pack_forget()   # shown only while working

    def _build_signals_tab(self, p):
        paned = ttk.Panedwindow(p, orient="horizontal"); paned.pack(fill="both", expand=True)
        lf = ttk.Frame(paned, style="Card.TFrame", padding=8); paned.add(lf, weight=0)
        hdr = ttk.Frame(lf, style="Card.TFrame"); hdr.pack(fill="x")
        ttk.Label(hdr, text="SIGNALS", style="Head.TLabel").pack(side="left")
        ttk.Button(hdr, text="All", width=4, style="Ghost.TButton", command=lambda: self._select_all(True)).pack(side="right")
        ttk.Button(hdr, text="None", width=5, style="Ghost.TButton", command=lambda: self._select_all(False)).pack(side="right", padx=4)

        wrap = ttk.Frame(lf, style="Card.TFrame"); wrap.pack(fill="both", expand=True, pady=(6, 0))
        self.sig_canvas = tk.Canvas(wrap, highlightthickness=0, width=290)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=self.sig_canvas.yview)
        self.sig_inner = ttk.Frame(self.sig_canvas, style="Card.TFrame")
        self.sig_inner.bind("<Configure>", lambda e: self.sig_canvas.configure(scrollregion=self.sig_canvas.bbox("all")))
        self.sig_canvas.create_window((0, 0), window=self.sig_inner, anchor="nw")
        self.sig_canvas.configure(yscrollcommand=sb.set)
        self.sig_canvas.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        self.sig_canvas.bind_all("<MouseWheel>", self._on_wheel)

        sf = ttk.Frame(paned, style="TFrame"); paned.add(sf, weight=1)
        cols = ("unit", "n", "min", "max", "mean", "std")
        self.stats = ttk.Treeview(sf, columns=cols, show="tree headings", height=22)
        self.stats.heading("#0", text="Signal"); self.stats.column("#0", width=150, anchor="w")
        for cc, w in dict(unit=55, n=70, min=80, max=80, mean=80, std=80).items():
            self.stats.heading(cc, text=cc.capitalize()); self.stats.column(cc, width=w, anchor="e")
        self.stats.pack(fill="both", expand=True)

    def _build_traces_tab(self, p):
        self.fig_tr = Figure(figsize=(7, 4.6), dpi=100)
        self.canvas_tr = FigureCanvasTkAgg(self.fig_tr, master=p)
        self.canvas_tr.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.canvas_tr, p)
        ttk.Label(p, text="Tip: drag across the plot to select a time window.",
                  style="Dim.TLabel").pack(anchor="w", padx=6, pady=(0, 4))

    def _build_overlay_tab(self, p):
        ctrl = ttk.Frame(p, style="Card.TFrame", padding=8); ctrl.pack(fill="x")
        ttk.Label(ctrl, text="Overlay 2–4 signals:", style="Card.TLabel").grid(row=0, column=0, padx=(0, 8))
        self.ov_combos = []
        for k in range(4):
            cb = ttk.Combobox(ctrl, width=22, state="readonly"); cb.grid(row=0, column=1 + k, padx=3)
            self.ov_combos.append(cb)
        self.ov_norm = tk.BooleanVar(value=False)
        ttk.Checkbutton(ctrl, text="Normalize 0–1", variable=self.ov_norm, style="Card.TCheckbutton",
                        command=self.draw_overlay).grid(row=0, column=6, padx=8)
        ttk.Button(ctrl, text="Plot", style="Accent.TButton", command=self.draw_overlay).grid(row=0, column=7)
        self.fig_ov = Figure(figsize=(7, 4.4), dpi=100)
        self.canvas_ov = FigureCanvasTkAgg(self.fig_ov, master=p)
        self.canvas_ov.get_tk_widget().pack(fill="both", expand=True)
        NavigationToolbar2Tk(self.canvas_ov, p)

    def _build_bus_health_tab(self, p):
        top = ttk.Frame(p, style="TFrame"); top.pack(fill="x")

        optc = ttk.Frame(top, style="Card.TFrame", padding=10); optc.pack(side="left", fill="y", padx=(0, 10))
        ttk.Label(optc, text="OPTIONS", style="Head.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(optc, text="Bus bit rate (bit/s)", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=2)
        self.bus_baud = ttk.Entry(optc, width=12)
        self.bus_baud.grid(row=1, column=1, padx=(10, 0))
        ttk.Label(optc, text="Auto-filled from the log header when\nBUSMASTER recorded it; otherwise you'll\nbe asked for it.",
                 style="Dim.TLabel", justify="left").grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 6))
        ttk.Button(optc, text="Analyze bus health  ⚙", style="Accent.TButton",
                  command=self.compute_bus_health).grid(row=3, column=0, columnspan=2, sticky="ew")

        tiles = ttk.Frame(top, style="Card.TFrame", padding=10); tiles.pack(side="left", fill="y", padx=(0, 10))
        ttk.Label(tiles, text="ESTIMATE", style="Head.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(tiles, text="Average load", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=2)
        self.bus_avg_lbl = ttk.Label(tiles, text="—", style="Mono.TLabel")
        self.bus_avg_lbl.grid(row=1, column=1, sticky="w", padx=(10, 0))
        ttk.Label(tiles, text="Peak load", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=2)
        self.bus_peak_lbl = ttk.Label(tiles, text="—", style="Mono.TLabel")
        self.bus_peak_lbl.grid(row=2, column=1, sticky="w", padx=(10, 0))
        ttk.Label(tiles, text="Unique IDs", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=2)
        self.bus_ids_lbl = ttk.Label(tiles, text="—", style="Mono.TLabel")
        self.bus_ids_lbl.grid(row=3, column=1, sticky="w", padx=(10, 0))

        notec = ttk.Frame(top, style="Card.TFrame", padding=10); notec.pack(side="left", fill="both", expand=True)
        ttk.Label(notec, text="HOW LOAD % IS CALCULATED", style="Head.TLabel").pack(anchor="w", pady=(0, 6))
        ttk.Label(notec, text=core.BUS_LOAD_FORMULA, style="Dim.TLabel",
                 wraplength=520, justify="left").pack(anchor="w", fill="x")
        self.bus_warn_lbl = ttk.Label(notec, text="", style="Card.TLabel", wraplength=520, justify="left")
        self.bus_warn_lbl.pack(anchor="w", fill="x", pady=(6, 0))

        mid = ttk.Panedwindow(p, orient="horizontal"); mid.pack(fill="both", expand=True, pady=(10, 0))
        rf = ttk.Frame(mid, style="TFrame"); mid.add(rf, weight=1)
        cols = ("count", "hz", "avg", "max", "dropout")
        self.bus_tree = ttk.Treeview(rf, columns=cols, show="tree headings", height=16)
        self.bus_tree.heading("#0", text="ID  ·  Message"); self.bus_tree.column("#0", width=220, anchor="w")
        for cc, label, w in (("count", "Count", 70), ("hz", "Hz", 70), ("avg", "Avg gap (ms)", 100),
                             ("max", "Max gap (ms)", 100), ("dropout", "Dropout?", 80)):
            self.bus_tree.heading(cc, text=label); self.bus_tree.column(cc, width=w, anchor="e")
        self.bus_tree.pack(fill="both", expand=True)

        pf = ttk.Frame(mid, style="TFrame"); mid.add(pf, weight=1)
        self.fig_bus = Figure(figsize=(6, 3.8), dpi=100)
        self.canvas_bus = FigureCanvasTkAgg(self.fig_bus, master=pf)
        self.canvas_bus.get_tk_widget().pack(fill="both", expand=True)

    def _build_dynamics_tab(self, p):
        top = ttk.Frame(p, style="TFrame"); top.pack(fill="x")
        mapc = ttk.Frame(top, style="Card.TFrame", padding=10); mapc.pack(side="left", fill="y", padx=(0, 10))
        ttk.Label(mapc, text="MAP SIGNALS TO ROLES", style="Head.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        self.role_combos = {}
        for i, (rid, label, _) in enumerate(ROLES):
            ttk.Label(mapc, text=label, style="Card.TLabel").grid(row=i + 1, column=0, sticky="w", pady=2)
            cb = ttk.Combobox(mapc, width=24, state="readonly"); cb.grid(row=i + 1, column=1, padx=(10, 0), pady=2)
            self.role_combos[rid] = cb

        optc = ttk.Frame(top, style="Card.TFrame", padding=10); optc.pack(side="left", fill="y")
        ttk.Label(optc, text="OPTIONS", style="Head.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(optc, text="Speed unit", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=2)
        self.opt_speed = ttk.Combobox(optc, values=["km/h", "m/s"], width=10, state="readonly")
        self.opt_speed.current(0); self.opt_speed.grid(row=1, column=1, padx=(10, 0))
        ttk.Label(optc, text="Vehicle mass (kg)", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=2)
        self.opt_mass = ttk.Entry(optc, width=12); self.opt_mass.grid(row=2, column=1, padx=(10, 0))
        ttk.Label(optc, text="Payload (kg)", style="Card.TLabel").grid(row=3, column=0, sticky="w", pady=2)
        self.opt_payload = ttk.Entry(optc, width=12); self.opt_payload.grid(row=3, column=1, padx=(10, 0))
        ttk.Button(optc, text="Compute dynamics  ⚙", style="Accent.TButton",
                   command=self.compute_dyn).grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        mid = ttk.Panedwindow(p, orient="horizontal"); mid.pack(fill="both", expand=True, pady=(10, 0))
        rf = ttk.Frame(mid, style="TFrame"); mid.add(rf, weight=0)
        self.dyn_tree = ttk.Treeview(rf, columns=("value", "unit"), show="tree headings", height=16)
        self.dyn_tree.heading("#0", text="Metric"); self.dyn_tree.column("#0", width=240, anchor="w")
        self.dyn_tree.heading("value", text="Value"); self.dyn_tree.column("value", width=120, anchor="e")
        self.dyn_tree.heading("unit", text="Unit"); self.dyn_tree.column("unit", width=70, anchor="w")
        self.dyn_tree.pack(fill="both", expand=True)

        pf = ttk.Frame(mid, style="TFrame"); mid.add(pf, weight=1)
        self.fig_dyn = Figure(figsize=(6, 3.8), dpi=100)
        self.canvas_dyn = FigureCanvasTkAgg(self.fig_dyn, master=pf)
        self.canvas_dyn.get_tk_widget().pack(fill="both", expand=True)

    def _build_components_tab(self, p):
        paned = ttk.Panedwindow(p, orient="horizontal"); paned.pack(fill="both", expand=True)

        lf = ttk.Frame(paned, style="Card.TFrame", padding=10); paned.add(lf, weight=0)
        ttk.Label(lf, text="SUBSYSTEM", style="Head.TLabel").pack(anchor="w", pady=(0, 6))
        self.comp_list = ttk.Treeview(lf, columns=("n",), show="tree headings",
                                      height=12, selectmode="browse")
        self.comp_list.heading("#0", text="Component"); self.comp_list.column("#0", width=130, anchor="w")
        self.comp_list.heading("n", text="Signals"); self.comp_list.column("n", width=64, anchor="e")
        self.comp_list.pack(fill="both", expand=True)
        self.comp_list.bind("<<TreeviewSelect>>", lambda e: self._on_component_select())

        rf = ttk.Frame(paned, style="TFrame"); paned.add(rf, weight=1)
        cols = ("unit", "n", "min", "max", "mean", "std")
        self.comp_stats = ttk.Treeview(rf, columns=cols, show="tree headings", height=20)
        self.comp_stats.heading("#0", text="Signal"); self.comp_stats.column("#0", width=190, anchor="w")
        for cc, w in dict(unit=60, n=70, min=85, max=85, mean=85, std=85).items():
            self.comp_stats.heading(cc, text=cc.capitalize()); self.comp_stats.column(cc, width=w, anchor="e")
        self.comp_stats.pack(fill="both", expand=True)

    def _build_faults_tab(self, p):
        top = ttk.Frame(p, style="TFrame"); top.pack(fill="x")

        selc = ttk.Frame(top, style="Card.TFrame", padding=10); selc.pack(side="left", fill="y", padx=(0, 10))
        hd = ttk.Frame(selc, style="Card.TFrame"); hd.pack(fill="x")
        ttk.Label(hd, text="FAULT SIGNALS", style="Head.TLabel").pack(side="left")
        ttk.Button(hd, text="All", width=4, style="Ghost.TButton",
                   command=lambda: self._fault_all(True)).pack(side="right")
        ttk.Button(hd, text="None", width=5, style="Ghost.TButton",
                   command=lambda: self._fault_all(False)).pack(side="right", padx=4)
        wrap = ttk.Frame(selc, style="Card.TFrame"); wrap.pack(fill="both", expand=True, pady=(6, 0))
        self.flt_canvas = tk.Canvas(wrap, highlightthickness=0, width=260, height=220)
        fsb = ttk.Scrollbar(wrap, orient="vertical", command=self.flt_canvas.yview)
        self.flt_inner = ttk.Frame(self.flt_canvas, style="Card.TFrame")
        self.flt_inner.bind("<Configure>", lambda e: self.flt_canvas.configure(scrollregion=self.flt_canvas.bbox("all")))
        self.flt_canvas.create_window((0, 0), window=self.flt_inner, anchor="nw")
        self.flt_canvas.configure(yscrollcommand=fsb.set)
        self.flt_canvas.pack(side="left", fill="both", expand=True); fsb.pack(side="right", fill="y")

        optc = ttk.Frame(top, style="Card.TFrame", padding=10); optc.pack(side="left", fill="y")
        ttk.Label(optc, text="OPTIONS", style="Head.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(optc, text="Active when", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=2)
        self.flt_mode = ttk.Combobox(optc, values=["value ≠ 0 (code)", "value ≥ threshold (flag)"],
                                     width=22, state="readonly")
        self.flt_mode.current(0); self.flt_mode.grid(row=1, column=1, padx=(10, 0))
        ttk.Label(optc, text="Threshold", style="Card.TLabel").grid(row=2, column=0, sticky="w", pady=2)
        self.flt_thr = ttk.Entry(optc, width=10); self.flt_thr.insert(0, "0.5")
        self.flt_thr.grid(row=2, column=1, sticky="w", padx=(10, 0))
        ttk.Button(optc, text="Detect faults  ⚑", style="Accent.TButton",
                   command=self.detect_faults_ui).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        self.flt_summary = ttk.Label(optc, text="", style="Dim.TLabel", wraplength=200, justify="left")
        self.flt_summary.grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        cols = ("code", "start", "end", "dur")
        self.flt_tree = ttk.Treeview(p, columns=cols, show="tree headings", height=14)
        self.flt_tree.heading("#0", text="Signal"); self.flt_tree.column("#0", width=200, anchor="w")
        self.flt_tree.heading("code", text="Code/Value"); self.flt_tree.column("code", width=90, anchor="e")
        self.flt_tree.heading("start", text="Start"); self.flt_tree.column("start", width=130, anchor="center")
        self.flt_tree.heading("end", text="End"); self.flt_tree.column("end", width=130, anchor="center")
        self.flt_tree.heading("dur", text="Duration (s)"); self.flt_tree.column("dur", width=100, anchor="e")
        self.flt_tree.pack(fill="both", expand=True, pady=(10, 0))

    def _build_report_tab(self, p):
        top = ttk.Frame(p, style="TFrame"); top.pack(fill="x")

        # report details
        dc = ttk.Frame(top, style="Card.TFrame", padding=14); dc.pack(side="left", fill="both", expand=True, padx=(0, 10))
        ttk.Label(dc, text="REPORT DETAILS", style="Head.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Label(dc, text="VIN", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=3)
        self.rep_vin = ttk.Entry(dc, width=26)
        self.rep_vin.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=3)
        ttk.Label(dc, text="Description", style="Card.TLabel").grid(row=2, column=0, sticky="nw", pady=3)
        self.rep_desc = tk.Text(dc, width=36, height=7, wrap="word", relief="flat",
                                highlightthickness=1, font=("Segoe UI", 9))
        self.rep_desc.grid(row=2, column=1, sticky="w", padx=(10, 0), pady=3)

        # sheet selection
        sc = ttk.Frame(top, style="Card.TFrame", padding=14); sc.pack(side="left", fill="both", expand=True)
        ttk.Label(sc, text="SHEETS TO INCLUDE", style="Head.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.sect_vars = {}
        for i, (sid, label) in enumerate(SECTIONS):
            var = tk.BooleanVar(value=True)
            ttk.Checkbutton(sc, text=label, variable=var, style="Card.TCheckbutton").grid(
                row=1 + i, column=0, sticky="w", pady=2)
            self.sect_vars[sid] = var
        ttk.Label(sc, text="Log Info is always included.", style="Dim.TLabel").grid(
            row=len(SECTIONS) + 1, column=0, sticky="w", pady=(8, 2))
        self.report_note = ttk.Label(sc, text="", style="Dim.TLabel")
        self.report_note.grid(row=len(SECTIONS) + 2, column=0, sticky="w")

        # action row
        act = ttk.Frame(p, style="TFrame"); act.pack(fill="x", pady=(12, 0))
        self.export_btn = ttk.Button(act, text="Export .xlsx…  ⭳", style="Accent.TButton",
                                     command=self.export, state="disabled")
        self.export_btn.pack(side="left")

    # ── theme ───────────────────────────────────────────────────────────────
    def toggle_theme(self):
        self.apply_theme("light" if self.theme == "dark" else "dark")

    def apply_theme(self, name):
        self.theme = name
        c = THEMES[name]
        self.theme_btn.configure(text="◑ Dark" if name == "light" else "◐ Light")
        st = self.style
        self.root.configure(bg=c["bg"])
        self.header.configure(bg=c["head"])
        self.h_title.configure(bg=c["head"], fg=c["ink"])
        self.h_sub.configure(bg=c["head"], fg=c["faint"])

        st.configure("TFrame", background=c["bg"])
        st.configure("Card.TFrame", background=c["panel"])
        st.configure("TLabel", background=c["bg"], foreground=c["ink"])
        st.configure("Card.TLabel", background=c["panel"], foreground=c["ink"])
        st.configure("Head.TLabel", background=c["panel"], foreground=c["accent"], font=("Segoe UI", 8, "bold"))
        st.configure("Dim.TLabel", background=c["panel"], foreground=c["muted"])
        st.configure("Mono.TLabel", background=c["panel"], foreground=c["ink"], font=("Consolas", 9))
        st.configure("MonoDim.TLabel", background=c["panel"], foreground=c["muted"], font=("Consolas", 9))
        st.configure("TButton", background=c["well"], foreground=c["ink"], borderwidth=0, padding=6, font=("Segoe UI", 9))
        st.map("TButton", background=[("active", c["line"])])
        st.configure("Ghost.TButton", background=c["panel"], foreground=c["muted"], borderwidth=1, padding=5)
        st.map("Ghost.TButton", background=[("active", c["well"])], foreground=[("active", c["accent"])])
        st.configure("Accent.TButton", background=c["accent"], foreground="#ffffff", padding=7, font=("Segoe UI", 9, "bold"))
        st.map("Accent.TButton", background=[("active", c["accent2"])])
        st.configure("TCheckbutton", background=c["bg"], foreground=c["ink"])
        st.configure("Card.TCheckbutton", background=c["panel"], foreground=c["ink"])
        st.map("Card.TCheckbutton", background=[("active", c["panel"])])
        st.configure("TNotebook", background=c["bg"], borderwidth=0)
        st.configure("TNotebook.Tab", background=c["well"], foreground=c["muted"], padding=(10, 6), borderwidth=0)
        st.map("TNotebook.Tab", background=[("selected", c["panel"])], foreground=[("selected", c["accent"])])
        st.configure("Treeview", background=c["panel"], fieldbackground=c["panel"], foreground=c["ink"], borderwidth=0, rowheight=23)
        st.configure("Treeview.Heading", background=c["head"], foreground=c["muted"], borderwidth=0, font=("Segoe UI", 8, "bold"))
        st.map("Treeview", background=[("selected", c["sel"])])
        st.configure("TCombobox", fieldbackground=c["entry"], background=c["well"], foreground=c["ink"], arrowcolor=c["muted"])
        st.map("TCombobox",
               fieldbackground=[("readonly", c["entry"]), ("disabled", c["entry"])],
               foreground=[("readonly", c["ink"])],
               selectbackground=[("readonly", c["entry"])],
               selectforeground=[("readonly", c["ink"])],
               background=[("readonly", c["well"]), ("active", c["well"])])
        st.configure("TEntry", fieldbackground=c["entry"], foreground=c["ink"])
        st.configure("Vertical.TScrollbar", background=c["well"], troughcolor=c["panel"], borderwidth=0, arrowcolor=c["muted"])
        st.configure("Green.Horizontal.TProgressbar", troughcolor=c["well"], bordercolor=c["well"],
                     background="#2ecc71", lightcolor="#2ecc71", darkcolor="#2ecc71")

        self.dbc_list.configure(bg=c["panel"], fg=c["ink"], selectbackground=c["sel"], selectforeground=c["ink"],
                                highlightbackground=c["line"], highlightcolor=c["line"])
        self.console.configure(bg=c["well"], fg=c["muted"], insertbackground=c["ink"], highlightbackground=c["line"])
        self.rep_desc.configure(bg=c["well"], fg=c["ink"], insertbackground=c["ink"], highlightbackground=c["line"])
        self.sig_canvas.configure(bg=c["panel"])
        self.flt_canvas.configure(bg=c["panel"])
        for tv in (self.stats, self.dyn_tree, self.flt_tree, self.comp_list, self.comp_stats):
            tv.tag_configure("odd", background=c["panel"])
            tv.tag_configure("even", background=c["stripe"])
        for fig in (self.fig_tr, self.fig_ov, self.fig_dyn):
            fig.set_facecolor(c["panel"])
        self._redraw_all()

    def _style_ax(self, ax):
        c = THEMES[self.theme]
        ax.set_facecolor(c["well"])
        for sp in ax.spines.values():
            sp.set_color(c["line"])
        ax.tick_params(colors=c["muted"], labelsize=8)
        ax.grid(True, color=c["grid"], alpha=0.6, linewidth=0.6)
        ax.xaxis.label.set_color(c["muted"]); ax.yaxis.label.set_color(c["muted"])
        ax.title.set_color(c["ink"])

    # ── misc ────────────────────────────────────────────────────────────────
    def _on_wheel(self, e):
        try:
            self.sig_canvas.yview_scroll(int(-e.delta / 120), "units")
        except Exception:
            pass

    def log(self, msg):
        self.console.configure(state="normal")
        self.console.insert("end", msg + "\n"); self.console.see("end")
        self.console.configure(state="disabled")

    def _on_callback_exception(self, exc, val, tb):
        """Replaces Tkinter's default report_callback_exception (which just
        prints — fatal with no console attached). Surfaces it in the app
        instead of losing it silently."""
        msg = "".join(traceback.format_exception(exc, val, tb))
        try:
            self.log("ERROR (unhandled):\n" + msg)
        except Exception:
            pass
        try:
            messagebox.showerror("CAN Signal Bench — error", f"{val}")
        except Exception:
            pass

    def set_status(self, m):
        self.status.configure(text=m)

    def _progress_show(self, on):
        if on:
            self.progress.pack(side="right")
            self.progress["value"] = 0
        else:
            self.progress.pack_forget()

    def _progress_set(self, frac, label=""):
        self.progress["value"] = max(0, min(100, frac * 100))
        if label:
            self.set_status(label)

    def _mk_progress(self):
        # thread-safe: workers call this; UI update marshalled onto main thread
        return lambda frac, label="": self.root.after(0, lambda: self._progress_set(frac, label))

    # ── files ───────────────────────────────────────────────────────────────
    def add_dbc(self):
        for p in filedialog.askopenfilenames(title="Select DBC file(s)",
                                              filetypes=[("DBC files", "*.dbc"), ("All files", "*.*")]):
            if len(self.dbc_paths) >= 10:
                self.log("DBC limit (10) reached."); break
            if p not in self.dbc_paths:
                self.dbc_paths.append(p)
        self._refresh_dbc()

    def remove_dbc(self):
        for i in reversed(self.dbc_list.curselection()):
            del self.dbc_paths[i]
        self._refresh_dbc()

    def _refresh_dbc(self):
        self.dbc_list.delete(0, "end")
        for p in self.dbc_paths:
            self.dbc_list.insert("end", os.path.basename(p))
        self.dbc_count.configure(text=f"{len(self.dbc_paths)} / 10")

    def choose_log(self):
        p = filedialog.askopenfilename(title="Select CAN log",
                                       filetypes=[("CAN logs", "*.log *.asc *.csv *.txt"), ("All files", "*.*")])
        if p:
            self.log_path = p; self.log_lbl.configure(text=os.path.basename(p))

    def run_async(self, work, done):
        def worker():
            try:
                res = work()
            except Exception as e:
                res = e; traceback.print_exc()
            self.root.after(0, lambda: done(res))
        threading.Thread(target=worker, daemon=True).start()

    # ── analyze ─────────────────────────────────────────────────────────────
    def analyze(self):
        if self.busy:
            return
        if not self.dbc_paths:
            messagebox.showwarning("CAN Signal Bench", "Add at least one DBC file."); return
        if not self.log_path:
            messagebox.showwarning("CAN Signal Bench", "Choose a log file."); return
        self.busy = True; self.analyze_btn.configure(state="disabled")
        self.set_status("Analyzing…"); self.log("\nParsing and decoding…")
        self._progress_show(True)
        prog = self._mk_progress()

        def work():
            prog(0.08, "Parsing DBC(s)…")
            messages, seen, dups = [], set(), 0
            for path in self.dbc_paths:
                with open(path, "r", errors="replace") as fh:
                    for m in core.parse_dbc(fh.read()):
                        if m["id"] in seen:
                            dups += 1; continue
                        seen.add(m["id"]); messages.append(m)
            prog(0.35, "Loading log…")
            fmt, frames, counts, meta = core.load_log(self.log_path)
            if not frames:
                raise RuntimeError("No frames parsed from the log.")
            prog(0.6, "Decoding signals…")
            series, matched = core.decode_all(frames, messages, None)
            prog(1.0, "")
            return dict(messages=messages, dups=dups, fmt=fmt, frames=frames,
                        counts=counts, meta=meta, series=series, matched=matched)
        self.run_async(work, self._analyze_done)

    def _analyze_done(self, res):
        self.busy = False; self.analyze_btn.configure(state="normal")
        self._progress_show(False)
        if isinstance(res, Exception):
            self.set_status("Analyze failed."); self.log("ERROR: " + str(res))
            messagebox.showerror("CAN Signal Bench", str(res)); return

        self.messages, self.frames, self.meta = res["messages"], res["frames"], res["meta"]
        self.frame_name_by_id = {m["id"]: m["name"] for m in self.messages}
        self.series_all = res["series"]
        fmt_label = "IXXAT MiniMon CSV" if res["fmt"] == "minimon" else "BUSMASTER"
        c = res["counts"]
        self.log(f"Detected {fmt_label} — {len(self.frames)} frames "
                 f"({c['std']} std, {c['ext']} ext, {c['malformed']} skipped).")
        if res["dups"]:
            self.log(f"{res['dups']} duplicate ID(s) across DBCs — kept first.")
        self.log(f"Decoded {len(self.series_all)} signal(s) from {len(res['matched'])} message ID(s).")

        ts = np.fromiter((f.t for f in self.frames), float, len(self.frames))
        self.log_t0 = float(ts.min()); self.log_dur = float(ts.max()) - self.log_t0
        self.trim_a = self.trim_b = None
        self.dyn_rows = []; self._dyn_ts = None; self.fault_rows = []
        self.bus_health = None
        for r in self.bus_tree.get_children():
            self.bus_tree.delete(r)
        self.bus_avg_lbl.configure(text="—"); self.bus_peak_lbl.configure(text="—")
        self.bus_ids_lbl.configure(text="—"); self.bus_warn_lbl.configure(text="")
        self.bus_baud.delete(0, "end")
        detected_baud = self.meta.get("baudrate")
        if detected_baud:
            self.bus_baud.insert(0, str(int(detected_baud)))
            self.log(f"Bus bit rate detected from log: {detected_baud/1000:.0f} kbit/s.")
        else:
            self.log("Bus bit rate not found in the log header — "
                     "Analyze bus health will ask you for it.")

        self._set_loginfo({"Format": self.meta["format"], "Date": self.meta["date"],
                           "Start": self.meta["start"], "End": self.meta["end"],
                           "Duration": self.meta["duration"], "Frames": str(len(self.frames)),
                           "Signals": str(len(self.series_all))})
        self._build_checklist()
        self._build_fault_checklist()
        self._populate_components()
        names = [s["name"] for s in self.series_all]
        for i, cb in enumerate(self.ov_combos):
            cb.configure(values=["— none —"] + names)
            cb.set(names[i] if i < 2 and i < len(names) else "— none —")
        self._autofill_roles(names)
        self._reset_trim_fields()
        self.export_btn.configure(state="normal")
        self._update_report_note()
        self.set_status(f"Done — {len(self.series_all)} signal(s).")
        self._redraw_all()

    def _set_loginfo(self, d):
        for k, lbl in self.info_vals.items():
            lbl.configure(text=d.get(k, "—"))

    # ── checkbox signal list ────────────────────────────────────────────────
    def _build_checklist(self):
        for w in self.sig_inner.winfo_children():
            w.destroy()
        self.sig_vars = []
        for s in self.series_all:
            var = tk.BooleanVar(value=True)
            unit = f" [{s['unit']}]" if s["unit"] else ""
            txt = f"0x{s['msg_id']:X}  {s['name']}{unit}  ({len(s['t'])})"
            ttk.Checkbutton(self.sig_inner, text=txt, variable=var, style="Card.TCheckbutton",
                            command=self._on_sig_select).pack(anchor="w", fill="x")
            self.sig_vars.append(var)

    def _select_all(self, on):
        for v in self.sig_vars:
            v.set(on)
        self._on_sig_select()

    def selected_series(self):
        return [s for s, v in zip(self.series_all, self.sig_vars) if v.get()]

    def _on_sig_select(self):
        self.draw_stats(); self.draw_traces(); self._update_report_note()

    # ── trim ────────────────────────────────────────────────────────────────
    def _clamp(self, v):
        return max(0.0, min(self.log_dur, v))

    def _rel_to_field(self, rel):
        if rel is None:
            return ""
        return core._sec_to_clock(self.log_t0 + rel) if self.trim_mode.get().startswith("Clock") else f"{rel:.2f}"

    def _field_to_rel(self, txt):
        txt = txt.strip()
        if not txt:
            return None
        try:
            return core.parse_trim_value(txt, self.log_t0)
        except ValueError:
            return "bad"

    def _reset_trim_fields(self):
        self.trim_start.delete(0, "end"); self.trim_end.delete(0, "end")
        self._update_trim_info()

    def _reformat_trim_fields(self):
        if self.trim_a is not None or self.trim_b is not None:
            a = self.trim_a if self.trim_a is not None else 0.0
            b = self.trim_b if self.trim_b is not None else self.log_dur
            self.trim_start.delete(0, "end"); self.trim_start.insert(0, self._rel_to_field(a))
            self.trim_end.delete(0, "end"); self.trim_end.insert(0, self._rel_to_field(b))

    def _update_trim_info(self):
        if not self.log_dur:
            self.trim_info.configure(text=""); return
        if self.trim_a is None and self.trim_b is None:
            self.trim_info.configure(text=f"full log · {self.log_dur:.1f} s")
        else:
            a = self.trim_a or 0.0
            b = self.trim_b if self.trim_b is not None else self.log_dur
            self.trim_info.configure(text=f"{core._sec_to_clock(self.log_t0+a)} – "
                                          f"{core._sec_to_clock(self.log_t0+b)}  ·  {b-a:.2f} s")

    def apply_trim_fields(self):
        if not self.series_all:
            return
        a = self._field_to_rel(self.trim_start.get()); b = self._field_to_rel(self.trim_end.get())
        if a == "bad" or b == "bad":
            messagebox.showwarning("CAN Signal Bench", "Use seconds (5) or clock (15:27:00)."); return
        a = 0.0 if a is None else self._clamp(a)
        b = self.log_dur if b is None else self._clamp(b)
        if b <= a:
            messagebox.showwarning("CAN Signal Bench", "Trim end must be after start."); return
        self._set_trim(a, b)

    def _set_trim(self, a, b):
        self.trim_a = None if a <= 1e-4 else a
        self.trim_b = None if b >= self.log_dur - 1e-4 else b
        self.trim_start.delete(0, "end"); self.trim_start.insert(0, self._rel_to_field(a))
        self.trim_end.delete(0, "end"); self.trim_end.insert(0, self._rel_to_field(b))
        self._update_trim_info()
        self.draw_stats(); self.draw_traces(); self.draw_overlay(keep=True)
        self._update_report_note()

    def reset_trim(self):
        self.trim_a = self.trim_b = None
        self._reset_trim_fields()
        self.draw_stats(); self.draw_traces(); self.draw_overlay(keep=True)
        self._update_report_note()

    def _trim_points(self, s):
        t, v = s["t"], s["v"]
        if self.trim_a is None and self.trim_b is None:
            return t, v
        a = -np.inf if self.trim_a is None else self.log_t0 + self.trim_a
        b = np.inf if self.trim_b is None else self.log_t0 + self.trim_b
        m = (t >= a) & (t <= b)
        return t[m], v[m]

    # ── drawing ─────────────────────────────────────────────────────────────
    def _redraw_all(self):
        if self.series_all:
            self.draw_stats(); self.draw_traces(); self.draw_overlay(keep=True)
            self._fill_dyn_tree(self.dyn_rows) if self.dyn_rows else None
            self._draw_dyn_plot()
            self._draw_bus_plot()
        else:
            for fig, canvas in ((self.fig_tr, self.canvas_tr), (self.fig_ov, self.canvas_ov),
                                (self.fig_dyn, self.canvas_dyn), (self.fig_bus, self.canvas_bus)):
                fig.clear(); canvas.draw()

    def draw_stats(self):
        for r in self.stats.get_children():
            self.stats.delete(r)
        i = 0
        for s in self.selected_series():
            t, v = self._trim_points(s)
            if not len(v):
                continue
            st = core.stats(v)
            self.stats.insert("", "end", text=s["name"], tags=("even" if i % 2 else "odd",),
                              values=(s["unit"] or "—", st["n"], f"{st['min']:.4f}",
                                      f"{st['max']:.4f}", f"{st['mean']:.4f}", f"{st['std']:.4f}"))
            i += 1

    def draw_traces(self):
        fig = self.fig_tr; fig.clear()
        c = THEMES[self.theme]; fig.set_facecolor(c["panel"])
        sel = [s for s in self.selected_series() if len(self._trim_points(s)[1])]
        if not sel:
            ax = fig.add_subplot(111); self._style_ax(ax)
            ax.text(0.5, 0.5, "Tick one or more signals", ha="center", va="center",
                    color=c["muted"], transform=ax.transAxes)
            self.canvas_tr.draw(); return
        sel = sel[:8]
        axes = fig.subplots(len(sel), 1, sharex=True, squeeze=False)[:, 0]
        for ax, s, col in zip(axes, sel, PALETTE):
            t, v = self._trim_points(s); t, v = decimate(t - self.log_t0, v)
            ax.plot(t, v, color=col, lw=1.0)
            ax.set_ylabel((s["unit"] or s["name"])[:14], fontsize=8)
            ax.set_title(s["name"], fontsize=9, loc="left")
            self._style_ax(ax)
        axes[-1].set_xlabel("t (s)")
        fig.tight_layout()
        self.span = SpanSelector(axes[-1], self._on_span, "horizontal", useblit=True,
                                 props=dict(alpha=0.18, facecolor=c["accent"]), interactive=False)
        self.canvas_tr.draw()

    def _on_span(self, xmin, xmax):
        if xmax - xmin < 1e-6 or not self.log_dur:
            return
        self._set_trim(self._clamp(xmin), self._clamp(xmax))

    def draw_overlay(self, keep=False):
        fig = self.fig_ov; fig.clear()
        c = THEMES[self.theme]; fig.set_facecolor(c["panel"])
        chosen, seen = [], set()
        for cb in self.ov_combos:
            n = cb.get()
            if n and n != "— none —" and n not in seen:
                s = next((x for x in self.series_all if x["name"] == n), None)
                if s is not None:
                    seen.add(n); chosen.append(s)
        chosen = [s for s in chosen if len(self._trim_points(s)[1])]
        ax0 = fig.add_subplot(111); self._style_ax(ax0)
        if len(chosen) < 2:
            ax0.text(0.5, 0.5, "Pick 2–4 signals, then Plot", ha="center", va="center",
                     color=c["muted"], transform=ax0.transAxes)
            self.canvas_ov.draw(); return
        normalize = self.ov_norm.get(); lines = []
        for i, s in enumerate(chosen):
            t, v = self._trim_points(s); t, v = decimate(t - self.log_t0, v)
            col = PALETTE[i % len(PALETTE)]
            if normalize:
                rng = (v.max() - v.min()) or 1.0
                ln, = ax0.plot(t, (v - v.min()) / rng, color=col, lw=1.1, label=s["name"])
            else:
                ax = ax0 if i == 0 else ax0.twinx()
                if i >= 2:
                    ax.spines["right"].set_position(("outward", 46 * (i - 1)))
                ln, = ax.plot(t, v, color=col, lw=1.1,
                              label=s["name"] + (f" [{s['unit']}]" if s["unit"] else ""))
                ax.set_ylabel(s["unit"] or s["name"], color=col, fontsize=8)
                ax.tick_params(axis="y", labelcolor=col, labelsize=8)
            lines.append(ln)
        ax0.set_xlabel("t (s)")
        if normalize:
            ax0.set_ylabel("normalized 0–1")
        ax0.legend(lines, [ln.get_label() for ln in lines], fontsize=8, loc="upper right")
        fig.tight_layout(); self.canvas_ov.draw()

    # ── components (per-subsystem grouping) ─────────────────────────────────
    def _classify_all(self):
        groups = {c: [] for c in COMPONENT_ORDER}
        for s in self.series_all:
            comp = classify_component(self.frame_name_by_id.get(s["msg_id"], ""), s["name"])
            groups[comp].append(s)
        return groups

    def _populate_components(self):
        self.comp_groups = self._classify_all()
        for r in self.comp_list.get_children():
            self.comp_list.delete(r)
        first = None
        for i, comp in enumerate(COMPONENT_ORDER):
            n = len(self.comp_groups.get(comp, []))
            if n == 0:
                continue
            iid = self.comp_list.insert("", "end", text=comp, values=(n,),
                                        tags=("even" if i % 2 else "odd",))
            if first is None:
                first = iid
        for r in self.comp_stats.get_children():
            self.comp_stats.delete(r)
        if first is not None:
            self.comp_list.selection_set(first)
            self._on_component_select()

    def _on_component_select(self):
        sel = self.comp_list.selection()
        if not sel:
            return
        comp = self.comp_list.item(sel[0], "text")
        for r in self.comp_stats.get_children():
            self.comp_stats.delete(r)
        i = 0
        for s in self.comp_groups.get(comp, []):
            t, v = self._trim_points(s)
            if not len(v):
                continue
            st = core.stats(v)
            self.comp_stats.insert("", "end", text=s["name"], tags=("even" if i % 2 else "odd",),
                                   values=(s["unit"] or "—", st["n"], f"{st['min']:.4f}",
                                           f"{st['max']:.4f}", f"{st['mean']:.4f}", f"{st['std']:.4f}"))
            i += 1

    def _component_export_rows(self, series_v):
        """rows for the Components report sheet, restricted to exported signals."""
        keep = {s["name"] for s in series_v}
        rows = []
        for comp in COMPONENT_ORDER:
            for s in self.comp_groups.get(comp, []):
                if s["name"] not in keep:
                    continue
                sv = next((x for x in series_v if x["name"] == s["name"]), None)
                if sv is None or not len(sv["v"]):
                    continue
                st = core.stats(sv["v"])
                rows.append((comp, s["name"], s["unit"], st["n"],
                             round(st["min"], 4), round(st["max"], 4),
                             round(st["mean"], 4), round(st["std"], 4)))
        return rows

    # ── vehicle dynamics ────────────────────────────────────────────────────
    def _autofill_roles(self, names):
        picks = guess_roles(names)
        for rid, _, _ in ROLES:
            self.role_combos[rid].configure(values=["— none —"] + names)
            self.role_combos[rid].set(picks.get(rid, "— none —"))

    def _role_series(self, rid):
        name = self.role_combos[rid].get()
        if not name or name == "— none —":
            return None
        s = next((x for x in self.series_all if x["name"] == name), None)
        if s is None:
            return None
        t, v = self._trim_points(s)
        if len(v) < 2:
            return None
        return (t, v, s["unit"])

    def compute_dyn(self):
        if not self.series_all:
            return
        roles = {rid: self._role_series(rid) for rid, _, _ in ROLES}
        def _num(entry):
            try:
                return float(entry.get()) if entry.get().strip() else None
            except ValueError:
                return None
        opts = dict(speed_unit=self.opt_speed.get(),
                    mass=_num(self.opt_mass), payload=_num(self.opt_payload))
        rows, raw, tsdata = compute_dynamics(roles, opts)
        self.dyn_rows = rows; self._dyn_ts = tsdata
        self._fill_dyn_tree(rows); self._draw_dyn_plot()
        self.log(f"Computed {len([r for r in rows if r[1] != ''])} dynamics metric(s).")

    def _fill_dyn_tree(self, rows):
        for r in self.dyn_tree.get_children():
            self.dyn_tree.delete(r)
        for i, (metric, value, unit) in enumerate(rows):
            self.dyn_tree.insert("", "end", text=metric, values=(value, unit),
                                 tags=("even" if i % 2 else "odd",))

    def _draw_dyn_plot(self):
        fig = self.fig_dyn; fig.clear()
        c = THEMES[self.theme]; fig.set_facecolor(c["panel"])
        ts = self._dyn_ts
        ax = fig.add_subplot(111); self._style_ax(ax)
        if not ts:
            ax.text(0.5, 0.5, "Map battery V+I (or power), then Compute",
                    ha="center", va="center", color=c["muted"], transform=ax.transAxes)
            self.canvas_dyn.draw(); return
        grid, power, cum = ts
        x = grid - self.log_t0
        xp, pp = decimate(x, power / 1000.0)
        ax.plot(xp, pp, color=PALETTE[0], lw=1.0)
        ax.axhline(0, color=c["line"], lw=0.7)
        ax.set_xlabel("t (s)"); ax.set_ylabel("Power (kW)", color=PALETTE[0])
        ax.tick_params(axis="y", labelcolor=PALETTE[0])
        ax2 = ax.twinx()
        xc, cc = decimate(x, cum)
        ax2.plot(xc, cc, color=PALETTE[1], lw=1.2)
        ax2.set_ylabel("Cumulative energy (Wh)", color=PALETTE[1])
        ax2.tick_params(axis="y", labelcolor=PALETTE[1], labelsize=8)
        ax.set_title("Battery power & energy", fontsize=9, loc="left")
        fig.tight_layout(); self.canvas_dyn.draw()

    # ── bus health ──────────────────────────────────────────────────────────
    def _trimmed_frames(self):
        if self.trim_a is None and self.trim_b is None:
            return self.frames
        a = -np.inf if self.trim_a is None else self.log_t0 + self.trim_a
        b = np.inf if self.trim_b is None else self.log_t0 + self.trim_b
        return [f for f in self.frames if a <= f.t <= b]

    @staticmethod
    def _parse_baud(txt):
        try:
            v = float(txt)
            return v if v > 0 else None
        except (ValueError, TypeError):
            return None

    def _ask_baudrate(self):
        """The log didn't declare a bit rate (or it's an invalid one) —
        prompt for it rather than silently assuming a value."""
        while True:
            val = simpledialog.askstring(
                "CAN Signal Bench",
                "Bus bit rate wasn't found in the log header.\n"
                "Enter the bus bit rate in bit/s (e.g. 500000, 250000, 125000):",
                parent=self.root)
            if val is None:
                return None
            b = self._parse_baud(val)
            if b is not None:
                return b
            messagebox.showwarning("CAN Signal Bench", "Enter a positive number, e.g. 500000.")

    def compute_bus_health(self):
        if not self.frames:
            return
        baud = self._parse_baud(self.bus_baud.get())
        if baud is None:
            baud = self._ask_baudrate()
            if baud is None:
                return
            self.bus_baud.delete(0, "end"); self.bus_baud.insert(0, str(int(baud)))
        fv = self._trimmed_frames()
        dur = (self.trim_b if self.trim_b is not None else self.log_dur) - (self.trim_a or 0.0)
        if not fv or dur <= 0:
            messagebox.showwarning("CAN Signal Bench", "No frames in this window.")
            return
        health = core.analyze_bus_health(fv, dur, self.frame_name_by_id, baudrate=baud)
        health["baudrate"] = baud
        self.bus_health = health
        self._fill_bus_tree(health)
        self._draw_bus_plot(health)
        self.bus_avg_lbl.configure(text=f"{health['load_avg']:.2f} %")
        self.bus_peak_lbl.configure(text=f"{health['load_peak']:.2f} %")
        self.bus_ids_lbl.configure(text=str(len(health["per_id"])))
        overload = health["load_peak"] > 100.0
        if overload:
            self.bus_warn_lbl.configure(
                text=f"⚠ Peak load {health['load_peak']:.0f}% is impossible on a real bus — "
                     f"the assumed bit rate is very likely too low. Try the next standard "
                     f"speed up (e.g. 250000, 500000, 1000000).")
        else:
            self.bus_warn_lbl.configure(text="")
        self.log(f"Bus health: {len(health['per_id'])} ID(s) — avg {health['load_avg']:.2f}% / "
                f"peak {health['load_peak']:.2f}% (assumed {baud/1000:.0f} kbit/s).")

    def _fill_bus_tree(self, health):
        for r in self.bus_tree.get_children():
            self.bus_tree.delete(r)
        for i, r in enumerate(health["per_id"]):
            label = f"0x{r['id']:X}  {r['name']}" if r["name"] else f"0x{r['id']:X}"
            self.bus_tree.insert("", "end", text=label, tags=("even" if i % 2 else "odd",),
                                 values=(r["count"], f"{r['hz']:.2f}", f"{r['avg_gap_ms']:.1f}",
                                         f"{r['max_gap_ms']:.1f}", "YES" if r["dropout"] else "—"))

    def _draw_bus_plot(self, health=None):
        fig = self.fig_bus; fig.clear()
        c = THEMES[self.theme]; fig.set_facecolor(c["panel"])
        ax = fig.add_subplot(111); self._style_ax(ax)
        health = health or self.bus_health
        bins = health["bins"] if health else []
        if health is None or len(bins) == 0:
            ax.text(0.5, 0.5, "Set a bit rate, then Analyze bus health",
                    ha="center", va="center", color=c["muted"], transform=ax.transAxes)
            self.canvas_bus.draw(); return
        ax.plot(bins, health["load"], color=PALETTE[5], lw=1.1)
        ax.axhline(100, color=c["line"], lw=0.8, linestyle="--")
        ax.set_xlabel("t (s)"); ax.set_ylabel("load (%)")
        ax.set_title("Bus load over time (estimated)", fontsize=9, loc="left")
        fig.tight_layout(); self.canvas_bus.draw()

    def _render_bus_png(self, path, health):
        """Bus-load-over-time chart, embedded on the Bus Health sheet —
        same data as the Bus Health tab's live plot."""
        fig = Figure(figsize=(6.8, 3.0), dpi=110)
        ax = fig.add_subplot(111)
        ax.plot(health["bins"], health["load"], color=PALETTE[5], lw=1.2)
        ax.axhline(100, color="#888888", lw=0.8, linestyle="--")
        ax.set_xlabel("t (s)", fontsize=9); ax.set_ylabel("Load (%)", fontsize=9)
        ax.grid(True, alpha=0.3)
        ax.set_title(f"Bus load over time — estimated at {health['baudrate']/1000:.0f} kbit/s",
                    fontsize=11, fontweight="bold")
        fig.tight_layout(); FigureCanvasAgg(fig).print_png(path)

    # ── faults ──────────────────────────────────────────────────────────────
    def _build_fault_checklist(self):
        for w in self.flt_inner.winfo_children():
            w.destroy()
        self.fault_vars = []
        for s in self.series_all:
            nl = s["name"].lower()
            guess = any(h in nl for h in FAULT_HINTS)
            var = tk.BooleanVar(value=guess)
            unit = f" [{s['unit']}]" if s["unit"] else ""
            ttk.Checkbutton(self.flt_inner, text=f"{s['name']}{unit}", variable=var,
                            style="Card.TCheckbutton").pack(anchor="w", fill="x")
            self.fault_vars.append(var)

    def _fault_all(self, on):
        for v in self.fault_vars:
            v.set(on)

    def detect_faults_ui(self):
        if not self.series_all:
            return
        items = []
        for s, v in zip(self.series_all, self.fault_vars):
            if v.get():
                t, val = self._trim_points(s)
                items.append((s["name"], t, val))
        if not items:
            messagebox.showwarning("CAN Signal Bench", "Tick at least one fault signal.")
            return
        mode = "flag" if self.flt_mode.get().startswith("value ≥") else "code"
        try:
            thr = float(self.flt_thr.get())
        except ValueError:
            thr = 0.5
        episodes, skipped = detect_faults(items, mode, thr)

        # format for table + export; keep absolute times for clock rendering
        self.fault_rows = []
        for name, code, sa, ea in episodes:
            self.fault_rows.append((
                name, code,
                core._sec_to_clock(sa), core._sec_to_clock(ea),
                f"{ea - sa:.3f}", f"{sa - self.log_t0:.3f}", f"{ea - self.log_t0:.3f}"))
        self._fill_fault_tree()
        total = len(self.fault_rows)
        active = sum(float(r[4]) for r in self.fault_rows)
        note = (f"{total} episode(s) · {active:.2f} s active total"
                if total else "No fault episodes in this window.")
        if skipped:
            note += f"\nSkipped {len(skipped)} non-discrete signal(s): " + ", ".join(skipped[:4])
            if len(skipped) > 4:
                note += " …"
        self.flt_summary.configure(text=note)
        self.log(f"Detected {total} fault episode(s)."
                 + (f" Skipped {len(skipped)} continuous signal(s)." if skipped else ""))

    def _fill_fault_tree(self):
        for r in self.flt_tree.get_children():
            self.flt_tree.delete(r)
        for i, r in enumerate(self.fault_rows):
            name, code, sc, ec, dur, _, _ = r
            self.flt_tree.insert("", "end", text=name, values=(code, sc, ec, dur),
                                 tags=("even" if i % 2 else "odd",))

    # ── report / export ─────────────────────────────────────────────────────
    def _update_report_note(self):
        if not self.series_all:
            self.report_note.configure(text=""); return
        n = len(self.selected_series())
        trim = "full log" if (self.trim_a is None and self.trim_b is None) else "trimmed window"
        self.report_note.configure(text=f"{n} signal(s) selected · {trim}")

    def export(self):
        if self.busy or not self.series_all:
            return
        sel = self.selected_series()
        if not sel:
            messagebox.showwarning("CAN Signal Bench", "Tick at least one signal."); return
        sections = {sid for sid, var in self.sect_vars.items() if var.get()}
        default = os.path.splitext(os.path.basename(self.log_path))[0] + "_report.xlsx"
        path = filedialog.asksaveasfilename(title="Save decoded report", defaultextension=".xlsx",
                                            initialfile=default, filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        self.busy = True; self.export_btn.configure(state="disabled"); self.set_status("Exporting…")
        self._progress_show(True)
        prog = self._mk_progress()

        t0 = self.log_t0
        trim_a, trim_b, dur = self.trim_a, self.trim_b, self.log_dur
        frames, meta = self.frames, dict(self.meta)
        series_all = self.series_all
        dyn_ts = self._dyn_ts
        bus_health = dict(self.bus_health) if (self.bus_health and "bus_health" in sections) else None
        # report metadata: VIN / payload / description
        meta["vin"] = self.rep_vin.get().strip().upper()
        try:
            pl = self.opt_payload.get().strip()
            meta["payload"] = f"{float(pl):g} kg" if pl else ""
        except ValueError:
            meta["payload"] = self.opt_payload.get().strip()
        meta["description"] = self.rep_desc.get("1.0", "end").strip()
        dbc_names = [os.path.basename(p) for p in self.dbc_paths]
        fname_by_id = self.frame_name_by_id
        dyn = list(self.dyn_rows) if ("dynamics" in sections and self.dyn_rows) else None
        flt = list(self.fault_rows) if ("faults" in sections and self.fault_rows) else None
        want_components = "components" in sections

        def work():
            a = -np.inf if trim_a is None else t0 + trim_a
            b = np.inf if trim_b is None else t0 + trim_b
            trimmed = trim_a is not None or trim_b is not None
            fv = [f for f in frames if a <= f.t <= b] if trimmed else frames

            def _trim_series(items):
                out = []
                for s in items:
                    mask = (s["t"] >= a) & (s["t"] <= b)
                    if mask.any():
                        out.append(dict(name=s["name"], unit=s["unit"], msg_id=s["msg_id"],
                                        t=s["t"][mask], v=s["v"][mask]))
                return out

            series_v = _trim_series(sel)             # ticked signals — charts only
            series_full = _trim_series(series_all)    # every DBC signal — Merged/Summary
            meta["trim"] = (f"{(trim_a or 0):.3f}-{(trim_b if trim_b is not None else dur):.3f} s"
                            if trimmed else "full log")
            comp = self._component_export_rows(series_v) if want_components else None

            plotdir = os.path.join(os.path.dirname(path), "plots")
            need_dyn_chart = dyn is not None and dyn_ts is not None
            if ("charts" in sections) or need_dyn_chart:
                os.makedirs(plotdir, exist_ok=True)

            plot_paths = []
            if "charts" in sections:
                n = max(1, len(series_v))
                for i, s in enumerate(series_v):
                    pp = os.path.join(plotdir, re.sub(r"[^\w.-]", "_", s["name"]) + ".png")
                    self._render_png(s, t0, pp, i); plot_paths.append(pp)
                    prog(0.01 + 0.03 * i / n, "Rendering charts…")

            dyn_chart_path = None
            if need_dyn_chart:
                dyn_chart_path = os.path.join(plotdir, "_vehicle_dynamics.png")
                self._render_dyn_png(dyn_chart_path, dyn_ts)

            if bus_health is not None and len(bus_health.get("bins", [])):
                os.makedirs(plotdir, exist_ok=True)
                bus_chart_path = os.path.join(plotdir, "_bus_health.png")
                self._render_bus_png(bus_chart_path, bus_health)
                bus_health["chart"] = bus_chart_path

            # pass only kwargs the installed core supports, so an older
            # can_log_analyzer.py degrades gracefully instead of crashing
            kw = dict(sections=sections, dynamics=dyn, faults=flt)
            params = inspect.signature(core.export_excel).parameters
            if "components" in params:
                kw["components"] = comp
            if "progress" in params:
                kw["progress"] = prog
            if "dyn_chart" in params:
                kw["dyn_chart"] = dyn_chart_path
            if "bus_health" in params:
                kw["bus_health"] = bus_health
            core.export_excel(path, meta, dbc_names, series_full, fv, fname_by_id,
                              0.1, plot_paths, t0, **kw)
            return path
        self.run_async(work, self._export_done)

    def _render_png(self, s, t0, path, i):
        fig = Figure(figsize=(6.4, 2.3), dpi=110)
        ax = fig.add_subplot(111)
        t, v = decimate(s["t"] - t0, s["v"])
        ax.plot(t, v, color=PALETTE[i % len(PALETTE)], lw=1.1)
        ax.set_title(s["name"] + (f" [{s['unit']}]" if s["unit"] else ""), fontsize=11, fontweight="bold")
        ax.set_xlabel("t (s)", fontsize=9); ax.grid(True, alpha=0.3)
        fig.tight_layout(); FigureCanvasAgg(fig).print_png(path)

    def _render_dyn_png(self, path, ts):
        """Battery power + cumulative energy chart, embedded on the Vehicle
        Dynamics sheet — same data as the Dynamics tab's live plot."""
        grid, power, cum = ts
        fig = Figure(figsize=(6.8, 3.0), dpi=110)
        ax = fig.add_subplot(111)
        x = grid - self.log_t0
        xp, pp = decimate(x, power / 1000.0)
        ax.plot(xp, pp, color=PALETTE[0], lw=1.1, label="Power (kW)")
        ax.axhline(0, color="#888888", lw=0.7)
        ax.set_xlabel("t (s)", fontsize=9)
        ax.set_ylabel("Power (kW)", color=PALETTE[0], fontsize=9)
        ax.tick_params(axis="y", labelcolor=PALETTE[0])
        ax.grid(True, alpha=0.3)
        ax2 = ax.twinx()
        xc, cc = decimate(x, cum)
        ax2.plot(xc, cc, color=PALETTE[1], lw=1.3, label="Cumulative energy (Wh)")
        ax2.set_ylabel("Cumulative energy (Wh)", color=PALETTE[1], fontsize=9)
        ax2.tick_params(axis="y", labelcolor=PALETTE[1])
        ax.set_title("Battery power & cumulative energy", fontsize=11, fontweight="bold")
        lines = ax.get_lines() + ax2.get_lines()
        ax.legend(lines, [ln.get_label() for ln in lines], fontsize=8, loc="upper left")
        fig.tight_layout(); FigureCanvasAgg(fig).print_png(path)

    def _export_done(self, res):
        self.busy = False; self.export_btn.configure(state="normal")
        self._progress_show(False)
        if isinstance(res, Exception):
            self.set_status("Export failed."); self.log("Export error: " + str(res))
            messagebox.showerror("CAN Signal Bench", str(res)); return
        self.set_status("Saved."); self.log(f"Saved {res}")
        messagebox.showinfo("CAN Signal Bench", f"Saved:\n{res}")


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
