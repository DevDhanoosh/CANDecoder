# CANDecode

**Live tool:** [https://devdhanoosh.github.io/CANDecoder/](https://devdhanoosh.github.io/CANDecoder/)

Open-source, DBC-guided CAN bus log decoder, plotter, bus-health analyzer, and
Excel exporter.

Give it one or more Vector **DBC** databases and a **CAN trace**, and it
decodes every signal the two agree on, charts it, checks bus health, and
exports a tidy Excel workbook. Runs entirely on your own computer — no file
is ever uploaded, no account, no server in the loop.

## Three ways to run it

| | Trace formats | Highlights |
|---|---|---|
| **`candecode.html`** (the live tool above) | BUSMASTER, quoted CSV, **candump/SocketCAN** | Sidebar dashboard, CAN-ID checkbox tree, chart zoom/pan/pinch + expand, radial bus-load gauges, min/max/avg on every chart, BUS-OFF/Power-OFF detection, a Custom Graph tool to overlay 2–4 signals |
| **CAN Signal Bench** — `can_log_analyzer.py` + `can_log_analyzer_gui.py` (or the packaged `.exe`) | BUSMASTER, quoted CSV, **candump/SocketCAN** | Desktop GUI: CAN-ID checkbox tree, chart scroll-zoom/pan/expand, a Bus Health workspace with BUS-OFF/Power-OFF detection, trim by dragging on the plot |
| **`candecode.py`** | BUSMASTER, quoted CSV | The original standalone CLI script — lightest option, opens file-picker dialogs if run with no arguments |

All three decode independently of each other — pick whichever fits your
workflow. `candecode.py` predates the candump/SocketCAN and BUS-OFF/Power-OFF
work added to the other two, so it doesn't have those yet; reach for
`can_log_analyzer.py` if you want that on the command line.

## What it decodes

- **DBC**: standard Vector `BO_` / `SG_` syntax — Intel (little-endian) and
  Motorola (big-endian, "sawtooth") bit layouts, 2's-complement signed
  signals, scale/offset/unit.
- **Trace log**, auto-detected:
  - **BUSMASTER** ASCII log (`.log` / `.asc` / `.txt`)
  - **Quoted CSV trace** (`.csv`) — a header row starting `Time,...` followed
    by rows of `"HH:MM:SS.ffffff","<ID hex>","Std/Ext","<DLC>","<hex bytes>"`
    (MiniMon-compatible)
  - **candump / SocketCAN** — *(candecode.html and CAN Signal Bench)* — both
    the `candump -l` log format (`(1699887726.123456) can0 123#DEADBEEF`) and
    the live/`-t`-timestamped console format (`can0  123   [8]  11 22 ...`).
    Handles extended IDs, remote frames, and basic CAN FD.

## Signal selection and charts

- **CAN-ID checkbox tree** *(candecode.html, CAN Signal Bench)* — signals are
  grouped under their CAN ID; tick the ID to select/deselect every signal
  under it at once, with a proper indeterminate state when only some are
  ticked.
- **Zoom, pan, expand** *(candecode.html, CAN Signal Bench)* — scroll to
  zoom, drag to pan, pinch on touch (browser only), and an "expand" button
  pops any chart into a bigger view with its own controls.
- **Min / max / avg on every chart** *(candecode.html)* — drawn directly on
  the chart, the same way the PNG plots exported by the Python tools already
  caption theirs.
- **Custom Graph** *(candecode.html)* — overlay 2–4 signals on one chart,
  either on independent y-axes or normalized 0–1 to compare shape regardless
  of units. Mirrors CAN Signal Bench's Overlay tab.
- **Radial bus-load gauges and KPI widgets** *(candecode.html)* — average and
  peak bus load shown as speedometer-style gauges; frame count, error count,
  and BUS-OFF count shown as widget tiles instead of a plain table cell.

## Bus health, including BUS-OFF / Power-OFF

- **Per-ID timing** (frequency, gap mean/std/max, dropout flag) is computed
  directly from the frame timestamps in your trace, so it's exact. A
  "dropout" flag means that ID's longest gap was more than 3× its own
  median gap — i.e. something that normally arrives on a regular beat went
  quiet for a while.
- **BUS-OFF / Power-OFF detection** *(candecode.html, CAN Signal Bench, and
  `can_log_analyzer.py --bus-health`)* — flags any stretch where **no frame
  of any CAN ID** was seen for longer than a configurable threshold (default
  2.0s): the signature of the bus going off, or the logger/vehicle losing
  power mid-capture. This is different from the per-ID dropout flag above,
  which only fires when one ID goes quiet while the rest of the bus keeps
  talking. Silence alone is a suspected interruption, not a confirmed one —
  shown as a tile, an events table (start/end/duration), and shaded red
  directly on the bus-load chart.
- **Bus load %** is an *estimate*. A plain ASCII trace doesn't record the
  exact bit-stuffed length of each frame, so load is derived from the
  standard CAN frame overhead (arbitration/control/CRC/ACK/EOF/IFS) plus a
  configurable average stuffing inflation factor (`--stuff-factor`, default
  `1.1`). Good for spotting load trends and hotspots, not a certified
  measurement — use `--baudrate` to match your actual bus speed.
- **Baud-rate mismatch detection.** If your DBC declares its bus speed via
  a `BS_:` record, CANDecode reads it and uses it as the default (still
  overridable with `--baudrate`); if you pass an explicit `--baudrate` that
  disagrees with the DBC, you'll get a warning. Separately, if the
  estimated peak load comes out **over 100%** — physically impossible on a
  real bus — that's a strong sign the assumed baud rate is too low, and
  CANDecode flags it explicitly rather than silently reporting a bogus
  number.
- **Error frames** (Stuff / Form / CRC / ACK / Bit / Overload) are only as
  good as what your logging tool actually wrote to the trace. CANDecode
  recognizes lines that indicate an error frame and reads off a sub-type
  keyword when the tool included one; when it didn't, the frame is still
  counted, just under "Error Frame (type not logged)". Most simple ASCII
  traces don't capture hardware-level error interrupts at all unless the
  tool was explicitly configured to log them.

## What it exports

One `.xlsx` workbook with:

- **Log Info** — format, date/time window, duration, DBC files used, message
  IDs matched, signal count.
- **Summary** — min/max/mean/std for **every** decoded signal.
- **CAN Frames** — the full raw trace, one row per frame.
- **Merged** — every decoded signal time-aligned onto one shared timeline
  (nearest-value join within a configurable tolerance), one column per
  signal. This always covers the full decoded set, independent of what
  you've ticked for charting.
- **Charts** — embedded plots for only the signals you ticked / requested.
- **Bus Health** — per-message-ID frequency (Hz), inter-frame gap
  mean/std/max, dropout flags, an estimated bus-load-over-time chart,
  error-frame counts, and (candecode.html / CAN Signal Bench /
  `can_log_analyzer.py --bus-health`) BUS-OFF/Power-OFF events.

There are no per-signal sheets — with dozens or hundreds of signals in a
DBC, one sheet per signal stops being useful; the Merged sheet is the
column-per-signal alternative.

## Using the HTML tool

Open `candecode.html` (or the live site above) — a sidebar lists seven
numbered stages: Source Files, Console, Signals Decoded, Stats, Bus Health,
Charts, and Support. Drop in your DBC(s) and log, set the merge tolerance,
bus bit rate, and BUS-OFF gap threshold, hit Decode, tick signals in the
checkbox tree, then export.

Everything — parsing, decoding, charting, workbook generation — runs
locally in JavaScript (Chart.js + ExcelJS, loaded from a CDN); no file is
ever uploaded anywhere.

## Using CAN Signal Bench (desktop GUI)

```bash
pip install numpy matplotlib openpyxl        # tqdm optional
python can_log_analyzer_gui.py
```

Load up to 10 DBCs and a trace (BUSMASTER, quoted CSV, or candump), tick
signals in the CAN-ID checkbox tree, scroll/drag/expand charts, set a bit
rate and BUS-OFF gap threshold in the Bus Health tab, trim by dragging on
the plot, then export.

## Using can_log_analyzer.py (CLI, with candump + BUS-OFF support)

```bash
pip install numpy matplotlib openpyxl        # tqdm optional

python can_log_analyzer.py --dbc vehicle.dbc --log trace.log --out ./out
python can_log_analyzer.py --dbc vehicle.dbc --log candump.log --out ./out \
    --bus-health --baudrate 500000 --gap-threshold 2.0
```

Run it with no arguments and it opens file-picker dialogs instead (falls
back to typed paths in a headless / no-Tk environment).

Key flags:

| Flag | Meaning |
|---|---|
| `--dbc FILE [FILE ...]` | up to 10 `.dbc` files; first definition of a message ID wins |
| `--log FILE` | the trace to decode (BUSMASTER, quoted CSV, or candump — auto-detected) |
| `--signals NAME [NAME ...]` | restrict **charting** to these signals (Summary/Merged always cover everything) |
| `--start` / `--end` | trim window — seconds from log start, or clock time `HH:MM:SS` |
| `--tol` | merge tolerance (seconds) for the Merged sheet, default `0.1` |
| `--bus-health` | add a Bus Health sheet: per-ID timing, bus load %, BUS-OFF/Power-OFF gaps |
| `--baudrate` | CAN bus bit rate, bit/s — used for the bus-load estimate, default `500000` |
| `--gap-threshold` | seconds of silence across ALL IDs before flagging BUS-OFF/Power-OFF, default `2.0` |
| `--no-plots` / `--no-excel` | skip PNGs / the workbook |
| `--show` | also pop up an interactive overlay plot |

## Using candecode.py (original CLI script)

```bash
pip install numpy matplotlib openpyxl        # tqdm optional, for a nicer progress bar

python candecode.py --dbc vehicle.dbc --log trace.log --out ./out
python candecode.py --dbc a.dbc b.dbc --log capture.csv --out ./out \
    --signals MotorTorque PackCurrent --start 5 --end 120
```

Run it with no arguments and it opens file-picker dialogs instead. Same
BUSMASTER/quoted-CSV-only format support as always; see `python
candecode.py --help` for its flags (`--baudrate`, `--stuff-factor`,
`--no-bus-health`, and the same trim/signal/export flags as above).

## Building a standalone .exe

No Python install needed for end users — build a single Windows executable:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name CANDecode candecode.py
# → dist/CANDecode.exe

pyinstaller --onefile --windowed --name CANSignalBench can_log_analyzer_gui.py
# → dist/CANSignalBench.exe
```

`--windowed` builds it without a console window, so double-clicking opens
straight into the file-picker dialogs (or the GUI) instead of a black
terminal — the scripts detect this and redirect their own console output,
showing a message box for the final summary and any errors instead.

`CANDecode.spec` / `CANSignalBench.spec` (generated on first build) record
the exact build config; re-run `pyinstaller <name>.spec` on later builds
instead of retyping flags.

## Privacy

Every version of this tool runs on your own machine. The HTML tool decodes
entirely in your browser using JavaScript — your DBC and trace files are
read locally and never sent anywhere. The Python scripts and the desktop
GUI never open a network connection for decoding either. Nothing is
uploaded, no account is needed, and no data leaves your computer.

## License

MIT — see [LICENSE](LICENSE). This is an independent personal project with
no proprietary or vendor DBC content bundled or referenced; bring your own.

## Author

**Dhanoosh B**
GitHub: [github.com/DevDhanoosh](https://github.com/DevDhanoosh)
LinkedIn: [linkedin.com/in/dhanoosh-b-754945199](https://www.linkedin.com/in/dhanoosh-b-754945199/)
Email: [dhanoosh2001@gmail.com](mailto:dhanoosh2001@gmail.com)
