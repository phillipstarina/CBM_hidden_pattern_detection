"""
CBM AI Analytics Platform — v5.2
─────────────────────────────────────────────
NEW in v5.2:
• FAST Access loader  — chunked read, background thread, no UI freeze
  even for 477 MB+ files; progress bar shows MB loaded
• WELL SELECTOR  — dropdown + listbox populated from well-ID column;
  pick one well to analyse vs all others
• MIXED-TYPE CLUSTERING  — numeric columns scaled normally;
  categorical/string columns label-encoded then scaled;
  ALL parameters contribute to clustering automatically
• HIDDEN PATTERN NAMING  — after per-well analysis a dedicated
  "Hidden Patterns" insight section names WHICH parameter differs
  between the selected well and its cluster peers, computes the
  z-score gap, and explains WHY in plain English
• Everything from v5.1 retained (per-param anomaly, RPM filter, events)
"""

import tkinter as tk
from tkinter import (filedialog, Listbox, Canvas, Scrollbar,
                     messagebox, ttk as tk_ttk)
import numpy as np
import matplotlib.patches as mpatches
import matplotlib.gridspec as gridspec
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.pyplot as plt
import threading
import traceback
import os
import datetime
import math
import matplotlib
matplotlib.use("TkAgg")

try:
    import ttkbootstrap as ttk
    from ttkbootstrap.constants import *
    HAVE_TTKBS = True
except ImportError:
    import tkinter.ttk as ttk
    HAVE_TTKBS = False


# ══════════════════════════════════════════════════════════════════════════════
# FAST MICROSOFT ACCESS LOADER
# Strategy: pyodbc with server-side cursor (fetchmany chunks) for large files
# Falls back to pandas_access on Linux/Mac
# ══════════════════════════════════════════════════════════════════════════════
def _get_access_driver():
    """Return first available Access ODBC driver name, or None."""
    try:
        import pyodbc
        for d in pyodbc.drivers():
            if "access" in d.lower() or "mdb" in d.lower():
                return d
    except ImportError:
        pass
    return None


def _list_access_tables_pyodbc(filepath, driver):
    import pyodbc
    cs  = f"Driver={{{driver}}};Dbq={filepath};Uid=Admin;Pwd=;Exclusive=No;"
    con = pyodbc.connect(cs, timeout=30, autocommit=True)
    cur = con.cursor()
    tables = [r.table_name for r in cur.tables(tableType="TABLE")
              if not r.table_name.startswith("MSys")]
    cur.close(); con.close()
    return tables


def _fast_read_access_table(filepath, table_name, driver,
                             chunk_size=50_000, progress_cb=None):
    """
    Read an Access table in chunks so the UI stays responsive.
    progress_cb(rows_loaded, total_estimate) called every chunk.
    Returns a pandas DataFrame.
    """
    import pyodbc, pandas as pd
    cs  = f"Driver={{{driver}}};Dbq={filepath};Uid=Admin;Pwd=;Exclusive=No;"
    con = pyodbc.connect(cs, timeout=120, autocommit=True)

    # Get row count estimate (may fail on some Access versions — ignore)
    total_est = 0
    try:
        cur2 = con.cursor()
        cur2.execute(f"SELECT COUNT(*) FROM [{table_name}]")
        total_est = cur2.fetchone()[0]
        cur2.close()
    except Exception:
        total_est = 0

    cur = con.cursor()
    cur.execute(f"SELECT * FROM [{table_name}]")
    columns = [col[0] for col in cur.description]

    chunks = []
    loaded = 0
    while True:
        rows = cur.fetchmany(chunk_size)
        if not rows:
            break
        chunks.append(pd.DataFrame.from_records(rows, columns=columns))
        loaded += len(rows)
        if progress_cb:
            progress_cb(loaded, total_est)

    cur.close(); con.close()
    if not chunks:
        return pd.DataFrame(columns=columns)
    return pd.concat(chunks, ignore_index=True)


def _pick_table_dialog(tables, filename):
    """Modal dialog — user picks one table from the Access DB."""
    result = [tables[0]]
    dlg = tk.Toplevel()
    dlg.title(f"Select Table — {filename}")
    dlg.configure(bg="#0d1117")
    dlg.resizable(False, False)
    dlg.grab_set()

    tk.Label(dlg, text=f"Select table to load from:\n{filename}",
             bg="#0d1117", fg="#e2e8f0",
             font=("Courier New", 10, "bold"), pady=8).pack(padx=16)

    lb_frame = tk.Frame(dlg, bg="#0d1117"); lb_frame.pack(padx=16, pady=6, fill="both")
    sb2 = tk.Scrollbar(lb_frame); sb2.pack(side="right", fill="y")
    lb  = Listbox(lb_frame, yscrollcommand=sb2.set, bg="#131b2a", fg="#f1f5f9",
                  selectbackground="#3b82f6", font=("Courier New", 9),
                  height=min(len(tables), 12), width=44, relief="flat", bd=0)
    lb.pack(side="left", fill="both"); sb2.config(command=lb.yview)
    for t in tables:
        lb.insert("end", f"  {t}")
    lb.selection_set(0)

    def _ok():
        sel = lb.curselection()
        if sel: result[0] = tables[sel[0]]
        dlg.destroy()

    tk.Button(dlg, text="  Load Selected Table  ", command=_ok,
              bg="#1e3a8a", fg="#93c5fd", font=("Courier New", 9, "bold"),
              relief="flat", bd=0, padx=14, pady=8, cursor="hand2").pack(pady=(4, 14))
    dlg.wait_window()
    return result[0]


def load_access_file_fast(filepath, progress_cb=None):
    """
    Fast Access loader.  Returns (df, table_name, method_used).
    progress_cb(rows_loaded, total_est) for UI updates.
    """
    import pandas as pd

    driver = _get_access_driver()

    # ── pyodbc path (Windows + ACE driver) ───────────────────────────────────
    if driver:
        tables = _list_access_tables_pyodbc(filepath, driver)
        if not tables:
            raise RuntimeError("No user tables found in this Access database.")
        table = tables[0] if len(tables) == 1 else \
                _pick_table_dialog(tables, os.path.basename(filepath))
        df = _fast_read_access_table(filepath, table, driver,
                                     chunk_size=50_000,
                                     progress_cb=progress_cb)
        return df, table, f"pyodbc / {driver}"

    # ── pandas_access path (Linux/Mac) ────────────────────────────────────────
    try:
        import pandas_access as mdb
        tables = [t for t in mdb.list_tables(filepath)
                  if not t.startswith("MSys")]
        if not tables:
            raise RuntimeError("No user tables found (pandas_access).")
        table = tables[0] if len(tables) == 1 else \
                _pick_table_dialog(tables, os.path.basename(filepath))
        if progress_cb:
            progress_cb(0, 0)
        df = mdb.read_table(filepath, table)
        return df, table, "pandas_access"
    except ImportError:
        pass

    raise RuntimeError(
        f"Cannot open '{os.path.basename(filepath)}'.\n\n"
        "On Windows: install pyodbc  +  Microsoft Access Database Engine:\n"
        "  https://www.microsoft.com/en-us/download/details.aspx?id=54920\n"
        "  (match 64-bit/32-bit to your Python)\n\n"
        "On Linux/Mac: sudo apt install mdbtools && pip install pandas-access"
    )


def load_any_file(filepath, progress_cb=None):
    """Universal loader. Returns (df, label)."""
    import pandas as pd
    ext = os.path.splitext(filepath)[1].lower().strip(".")

    if ext in ("mdb", "accdb"):
        df, table, method = load_access_file_fast(filepath, progress_cb)
        return df, f"Access [{table}]  via {method}"

    strategies = []
    if ext in ("csv", "tsv", "txt", ""):
        sep = "\t" if ext in ("tsv", "txt") else ","
        strategies += [
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="utf-8"),
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="latin-1"),
            lambda f: pd.read_csv(f, sep=None, engine="python"),
        ]
    if ext in ("xlsx", "xlsm", "xlsb", "xls", "ods", "odf"):
        for eng in (["openpyxl"] if ext in ("xlsx","xlsm") else []) + \
                   (["pyxlsb"]   if ext == "xlsb"            else []) + \
                   (["xlrd"]     if ext == "xls"              else []) + \
                   (["odf"]      if ext in ("ods","odf")      else []) + [None]:
            if eng:
                strategies.append(lambda f, e=eng: pd.read_excel(f, engine=e))
            else:
                strategies.append(lambda f: pd.read_excel(f))
    if ext == "json":
        strategies += [lambda f: pd.read_json(f, orient="records"),
                       lambda f: pd.read_json(f)]
    if ext == "parquet": strategies += [lambda f: pd.read_parquet(f)]
    if ext == "feather": strategies += [lambda f: pd.read_feather(f)]
    if ext in ("h5","hdf5"): strategies += [lambda f: pd.read_hdf(f)]
    if ext in ("pkl","pickle"): strategies += [lambda f: pd.read_pickle(f)]
    if not strategies:
        strategies = [lambda f: pd.read_csv(f), lambda f: pd.read_excel(f)]

    errors = []
    for s in strategies:
        try:
            df = s(filepath)
            if df is not None and len(df.columns) > 0:
                return df, os.path.basename(filepath)
        except Exception as e:
            errors.append(str(e))
    raise RuntimeError(
        f"Could not read '{os.path.basename(filepath)}'.\n"
        + "\n".join(f"  • {e}" for e in errors[:5])
    )


# ══════════════════════════════════════════════════════════════════════════════
# MATPLOTLIB THEME
# ══════════════════════════════════════════════════════════════════════════════
plt.rcParams.update({
    "figure.facecolor":"#0d1117","axes.facecolor":"#161d2e",
    "text.color":"#e2e8f0","axes.labelcolor":"#94a3b8",
    "xtick.color":"#94a3b8","ytick.color":"#94a3b8",
    "axes.edgecolor":"#2d3f5e","grid.color":"#1e3a5f",
    "axes.grid":True,"grid.linewidth":0.5,"grid.alpha":0.4,
    "legend.facecolor":"#161d2e","legend.edgecolor":"#2d3f5e",
    "legend.fontsize":9,"figure.autolayout":False,"font.family":"monospace",
})

# ── Colour palette ─────────────────────────────────────────────────────────
BG_DEEP="#060a10"; BG="#0d1117"; BG_MID="#0f1520"
SIDEBAR="#080c14"; CARD="#131b2a"; CARD2="#1a2438"
BORDER="#1e3a5f"; BORDER2="#0e2040"

COL_DATASET="#22d3ee"; COL_CLUSTER="#a78bfa"; COL_WEIGHTS="#f59e0b"
COL_ANOMALY="#f87171"; COL_EVENTS="#34d399"; COL_FEATURES="#60a5fa"
COL_EXPORT="#4ade80"; COL_PREVIEW="#94a3b8"; COL_PARAM_ANOM="#fb923c"
COL_WELL_SEL="#e879f9"   # magenta accent for well-selector panel

ACCENT="#3b82f6"; FG="#f1f5f9"; FG_MID="#cbd5e1"; FG_DIM="#4a6080"
SUCCESS="#10b981"; SUCCESS2="#34d399"; WARN="#f59e0b"
ERR="#ef4444"; ANOMALY_C="#ef4444"; NORMAL_C="#3b82f6"

CLUSTER_PALETTE=[
    "#f59e0b","#3b82f6","#10b981","#ec4899",
    "#8b5cf6","#06b6d4","#ef4444","#84cc16","#f97316","#a78bfa",
]
EVENT_COLOURS={
    "well on":"#10b981","well off":"#ef4444",
    "maintenance shutdown":"#f59e0b","pump failure":"#f97316",
    "water breakthrough":"#06b6d4",
}
SEV_COLORS={
    "NONE":"#2d3f5e","LOW":"#10b981","MODERATE":"#f59e0b",
    "ELEVATED":"#f97316","HIGH":"#ef4444",
}

FONT_H1=("Georgia",13,"bold"); FONT_H2=("Georgia",11,"bold")
FONT_H3=("Courier New",10,"bold"); FONT_SM=("Courier New",9)
FONT_XS=("Courier New",8)

# ── Global state ──────────────────────────────────────────────────────────────
raw_data=None; active_df=None; active_X=None; active_xcols=[]
active_anomaly_result=None; active_weight_result=None
active_event_result=None; active_rpm_result=None
active_param_anom=None; active_well_analysis=None
active_figures={}

weight_vars={}; weight_row_frames=[]
weight_sum_var=None; weight_sum_lbl=None

_well_id_col=None       # detected well-ID column name
_all_well_ids=[]        # sorted list of distinct well IDs


# ══════════════════════════════════════════════════════════════════════════════
# UTILITY
# ══════════════════════════════════════════════════════════════════════════════
def cluster_hex(n):
    return [CLUSTER_PALETTE[i%len(CLUSTER_PALETTE)] for i in range(n)]

def _lighten(hex_col, amount=30):
    try:
        h=hex_col.lstrip("#"); r,g,b=int(h[:2],16),int(h[2:4],16),int(h[4:],16)
        return "#{:02x}{:02x}{:02x}".format(
            min(r+amount,255),min(g+amount,255),min(b+amount,255))
    except Exception: return hex_col

def ts(): return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def set_status(msg, color=FG_DIM):
    status_var.set(msg); status_lbl.config(fg=color)

def _save_df(df, path):
    ext
