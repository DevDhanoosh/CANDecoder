# CANDecode

**Live tool:** [https://devdhanoosh.github.io/CANDecode/](https://devdhanoosh.github.io/CANDecoder/)

Open-source, DBC-guided CAN bus log decoder, plotter, and Excel exporter.

Give it one or more Vector **DBC** databases and a **CAN trace**, and it decodes
every signal the two agree on, plots it, and exports a single tidy Excel
workbook.

Two independent ways to run it — pick whichever suits you:

| | |
|---|---|
| **`candecode.py`** | A standalone Python script — command line, or launch it with no arguments for interactive file-picker dialogs. |
| **`candecode.html`** | A single self-contained HTML page — open it in any modern browser, no install, no server, nothing leaves your machine. |

Both implement the exact same decode logic independently, so pick whichever
fits your workflow — there is no dependency between them.

## What it decodes

- **DBC**: standard Vector `BO_` / `SG_` syntax — Intel (little-endian) and
  Motorola (big-endian, "sawtooth") bit layouts, 2's-complement signed
  signals, scale/offset/unit.
- **Trace log**, auto-detected:
  - **BUSMASTER** ASCII log (`.log` / `.asc` / `.txt`)
  - **Quoted CSV trace** (`.csv`) — a header row starting `Time,...` followed
    by rows of `"HH:MM:SS.ffffff","<ID hex>","Std/Ext","<DLC>","<hex bytes>"`

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
  mean/std/max, dropout flags, an estimated bus-load-over-time chart, and
  error-frame counts (see below).

There are no per-signal sheets — with dozens or hundreds of signals in a
DBC, one sheet per signal stops being useful; the Merged sheet is the
column-per-signal alternative.

## Bus load & error frames — what's real vs. estimated

- **Per-ID timing** (frequency, gap mean/std/max, dropout flag) is computed
  directly from the frame timestamps in your trace, so it's exact. A
  "dropout" flag means that ID's longest gap was more than 3× its own
  median gap — i.e. something that normally arrives on a regular beat went
  quiet for a while.
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

## Using the Python script

```bash
pip install numpy matplotlib openpyxl        # tqdm optional, for a nicer progress bar

python candecode.py --dbc vehicle.dbc --log trace.log --out ./out
python candecode.py --dbc a.dbc b.dbc --log capture.csv --out ./out \
    --signals MotorTorque PackCurrent --start 5 --end 120
```

Run it with no arguments and it opens file-picker dialogs instead (falls
back to typed paths in a headless / no-Tk environment).

Key flags:

| Flag | Meaning |
|---|---|
| `--dbc FILE [FILE ...]` | up to 10 `.dbc` files; first definition of a message ID wins |
| `--log FILE` | the trace to decode |
| `--signals NAME [NAME ...]` | restrict **charting** to these signals (Summary/Merged always cover everything) |
| `--start` / `--end` | trim window — seconds from log start, or clock time `HH:MM:SS` |
| `--tol` | merge tolerance (seconds) for the Merged sheet, default `0.1` |
| `--baudrate` | CAN bus bit rate, bit/s — used for the bus-load estimate, default `500000` |
| `--stuff-factor` | average bit-stuffing inflation for the bus-load estimate, default `1.1` |
| `--no-plots` / `--no-excel` / `--no-frames` / `--no-bus-health` | skip PNGs / the workbook / the CAN Frames sheet / the Bus Health sheet |
| `--show` | also pop up an interactive overlay plot |

## Using the HTML tool

Just open `candecode.html` in a browser — every section is laid out on one
page, numbered `01`–`06`: Source Files / Console, Signals Decoded, Stats,
Bus Health, and Charts. Drop in your DBC(s) and log, set the merge
tolerance and bus bit rate, hit Decode, then tick the signals you want
charted and export the `.xlsx`.

Everything — parsing, decoding, plotting, workbook generation — runs
locally in JavaScript; no file is ever uploaded anywhere.

## Building a standalone .exe

No Python install needed for end users — build a single Windows executable:

```bash
pip install pyinstaller
pyinstaller --onefile --name CANDecode candecode.py
# → dist/CANDecode.exe
```

`CANDecode.spec` (generated on first build) records the exact build config;
re-run `pyinstaller CANDecode.spec` on later builds instead of retyping flags.

## License

MIT — see [LICENSE](LICENSE). This is an independent personal project with
no proprietary or vendor DBC content bundled or referenced; bring your own.

## Author

**Dhanoosh B**
GitHub: [github.com/DevDhanoosh](https://github.com/DevDhanoosh)
Email: [dhanoosh2001@gmail.com](mailto:dhanoosh2001@gmail.com)
