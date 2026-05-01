"""
CBM AI Analytics Platform — v5.1
─────────────────────────────────────────────
NEW in v5.1:
• Per-parameter anomaly analysis:
  - Each numeric column is individually analysed for outliers
  - A dedicated "Parameter Anomaly" tab shows which parameters have the
    most anomalous parameters, with z-score breakdown per column
  - The right-panel legend lists per-column anomaly counts
• Enhanced Insights Report:
  - Lists WHICH parameters contain anomalies
  - Explains WHY (mean, std-dev, z-score breach, skew)
  - Severity rating per parameter
  - Suggested field action for each flagged parameter
• Everything else from v5.0 unchanged
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
# MICROSOFT ACCESS LOADER
# ══════════════════════════════════════════════════════════════════════════════
def _get_access_tables_pyodbc(filepath):
    import pyodbc
    drivers = [d for d in pyodbc.drivers()
               if 'access' in d.lower() or 'mdb' in d.lower()]
    if not drivers:
        raise RuntimeError(
            "No Microsoft Access ODBC driver found.\n"
            "Install from:\n"
            "https://www.microsoft.com/en-us/download/details.aspx?id=54920\n"
            "(Use the 64-bit version if your Python is 64-bit)")
    driver   = drivers[0]
    conn_str = f"Driver={{{driver}}};Dbq={filepath};Uid=Admin;Pwd=;"
    conn     = pyodbc.connect(conn_str, timeout=30)
    cursor   = conn.cursor()
    tables   = [r.table_name for r in cursor.tables(tableType='TABLE')
                if not r.table_name.startswith('MSys')]
    conn.close()
    return tables, driver


def _load_access_table_pyodbc(filepath, table_name, driver):
    import pyodbc, pandas as pd
    conn_str = f"Driver={{{driver}}};Dbq={filepath};Uid=Admin;Pwd=;"
    conn     = pyodbc.connect(conn_str, timeout=120)
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM [{table_name}]")
        columns = [col[0] for col in cursor.description]
        rows    = cursor.fetchall()
        df      = pd.DataFrame.from_records(rows, columns=columns)
        cursor.close()
    except Exception as e:
        conn.close()
        raise RuntimeError(f"Could not read table '{table_name}': {e}")
    conn.close()
    return df


def _pick_table_dialog(tables, filename):
    result = [tables[0]]
    dlg = tk.Toplevel()
    dlg.title(f"Select Table — {filename}")
    dlg.configure(bg="#0d1117")
    dlg.resizable(False, False)
    dlg.grab_set()
    tk.Label(dlg, text=f"Multiple tables found in:\n{filename}",
             bg="#0d1117", fg="#e2e8f0",
             font=("Courier New", 10, "bold"), pady=8).pack(padx=16)
    tk.Label(dlg, text="Select the table to load:",
             bg="#0d1117", fg="#94a3b8",
             font=("Courier New", 9)).pack(padx=16)
    lb_frame = tk.Frame(dlg, bg="#0d1117")
    lb_frame.pack(padx=16, pady=6, fill="both", expand=True)
    sb2 = tk.Scrollbar(lb_frame); sb2.pack(side="right", fill="y")
    lb = tk.Listbox(lb_frame, yscrollcommand=sb2.set,
                    bg="#131b2a", fg="#f1f5f9",
                    selectbackground="#3b82f6",
                    font=("Courier New", 9),
                    height=min(len(tables), 12), width=42,
                    relief="flat", bd=0)
    lb.pack(side="left", fill="both", expand=True)
    sb2.config(command=lb.yview)
    for t in tables:
        lb.insert("end", f"  {t}")
    lb.selection_set(0)
    def _ok():
        sel = lb.curselection()
        if sel: result[0] = tables[sel[0]]
        dlg.destroy()
    tk.Button(dlg, text="Load Selected Table",
              command=_ok, bg="#1e3a8a", fg="#93c5fd",
              font=("Courier New", 9, "bold"),
              relief="flat", bd=0, padx=14, pady=8,
              cursor="hand2").pack(pady=(4, 14))
    dlg.wait_window()
    return result[0]


def load_access_file(filepath):
    import pandas as pd
    errors = []
    try:
        tables, driver = _get_access_tables_pyodbc(filepath)
        if not tables:
            raise RuntimeError("No user tables found in the Access database.")
        table = tables[0]
        if len(tables) > 1:
            table = _pick_table_dialog(tables, os.path.basename(filepath))
        df = _load_access_table_pyodbc(filepath, table, driver)
        return df, table
    except RuntimeError:
        raise
    except Exception as e:
        errors.append(f"pyodbc: {e}")
    try:
        import pandas_access as mdb
        tables = [t for t in mdb.list_tables(filepath)
                  if not t.startswith('MSys')]
        if not tables:
            raise RuntimeError("No user tables found (pandas_access).")
        table = tables[0]
        if len(tables) > 1:
            table = _pick_table_dialog(tables, os.path.basename(filepath))
        df = mdb.read_table(filepath, table)
        return df, table
    except Exception as e:
        errors.append(f"pandas_access: {e}")
    raise RuntimeError(
        f"Could not open '{os.path.basename(filepath)}'.\n\n"
        "Errors:\n" + "\n".join(f"  • {e}" for e in errors) + "\n\n"
        "Fix:\n"
        "  1. pip install pyodbc\n"
        "  2. Install Microsoft Access Database Engine (64-bit):\n"
        "     https://www.microsoft.com/en-us/download/details.aspx?id=54920\n"
        "  3. Ensure the file is not open in Access or password-protected."
    )


def load_any_file(filepath):
    import pandas as pd
    ext = os.path.splitext(filepath)[1].lower().strip(".")
    if ext in ("mdb", "accdb"):
        df, table = load_access_file(filepath)
        return df, f"Access table: {table}"
    strategies = []
    if ext in ("csv", "tsv", "txt", ""):
        sep = "\t" if ext in ("tsv", "txt") else ","
        strategies += [
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="utf-8"),
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="latin-1"),
            lambda f: pd.read_csv(f, sep=None, engine="python", encoding="utf-8"),
            lambda f: pd.read_csv(f, sep=None, engine="python", encoding="latin-1"),
        ]
    if ext in ("xlsx", "xlsm", "xlsb", "xls", "ods", "odf", "odt"):
        engines = []
        if ext in ("xlsx", "xlsm"): engines += ["openpyxl", None]
        if ext == "xlsb":           engines += ["pyxlsb"]
        if ext == "xls":            engines += ["xlrd", None]
        if ext in ("ods","odf","odt"): engines += ["odf", None]
        engines += [None]
        for eng in engines:
            if eng:
                strategies.append(lambda f, e=eng: pd.read_excel(f, engine=e))
            else:
                strategies.append(lambda f: pd.read_excel(f))
    if ext == "json":
        strategies += [
            lambda f: pd.read_json(f, orient="records"),
            lambda f: pd.read_json(f),
        ]
    if ext == "parquet":  strategies += [lambda f: pd.read_parquet(f)]
    if ext == "feather":  strategies += [lambda f: pd.read_feather(f)]
    if ext in ("h5","hdf5","hdf"): strategies += [lambda f: pd.read_hdf(f)]
    if ext in ("pkl","pickle"):    strategies += [lambda f: pd.read_pickle(f)]
    if not strategies:
        strategies += [
            lambda f: pd.read_csv(f, encoding="utf-8"),
            lambda f: pd.read_csv(f, encoding="latin-1"),
            lambda f: pd.read_csv(f, sep=None, engine="python"),
            lambda f: pd.read_excel(f),
            lambda f: pd.read_json(f),
        ]
    errors = []
    for strategy in strategies:
        try:
            df = strategy(filepath)
            if df is not None and len(df.columns) > 0:
                return df, os.path.basename(filepath)
        except Exception as e:
            errors.append(str(e))
    short_errors = "\n".join(f"  • {e}" for e in errors[:6])
    raise RuntimeError(
        f"Could not read '{os.path.basename(filepath)}'.\n\n"
        f"Tried {len(strategies)} strategies. Errors:\n{short_errors}\n\n"
        "Supported: CSV, TSV, XLSX, XLS, ODS, JSON, Parquet, "
        "Feather, HDF5, Pickle, MDB, ACCDB.\n"
        "Make sure the file is not open elsewhere or password-protected."
    )


# ══════════════════════════════════════════════════════════════════════════════
# MATPLOTLIB THEME
# ══════════════════════════════════════════════════════════════════════════════
plt.rcParams.update({
    "figure.facecolor": "#0d1117", "axes.facecolor":  "#161d2e",
    "text.color":       "#e2e8f0", "axes.labelcolor": "#94a3b8",
    "xtick.color":      "#94a3b8", "ytick.color":     "#94a3b8",
    "axes.edgecolor":   "#2d3f5e", "grid.color":      "#1e3a5f",
    "axes.grid": True,  "grid.linewidth": 0.5, "grid.alpha": 0.4,
    "legend.facecolor": "#161d2e", "legend.edgecolor": "#2d3f5e",
    "legend.fontsize":  9, "figure.autolayout": False,
    "font.family": "monospace",
})

BG_DEEP  = "#060a10"
BG       = "#0d1117"
BG_MID   = "#0f1520"
SIDEBAR  = "#080c14"
CARD     = "#131b2a"
CARD2    = "#1a2438"
BORDER   = "#1e3a5f"
BORDER2  = "#0e2040"

COL_DATASET  = "#22d3ee"
COL_CLUSTER  = "#a78bfa"
COL_WEIGHTS  = "#f59e0b"
COL_ANOMALY  = "#f87171"
COL_EVENTS   = "#34d399"
COL_FEATURES = "#60a5fa"
COL_EXPORT   = "#4ade80"
COL_PREVIEW  = "#94a3b8"
COL_PARAM_ANOM = "#fb923c"   # NEW — parameter anomaly tab accent

ACCENT    = "#3b82f6"
FG        = "#f1f5f9"
FG_MID    = "#cbd5e1"
FG_DIM    = "#4a6080"
SUCCESS   = "#10b981"
SUCCESS2  = "#34d399"
WARN      = "#f59e0b"
ERR       = "#ef4444"
ANOMALY_C = "#ef4444"
NORMAL_C  = "#3b82f6"

CLUSTER_PALETTE = [
    "#f59e0b","#3b82f6","#10b981","#ec4899",
    "#8b5cf6","#06b6d4","#ef4444","#84cc16","#f97316","#a78bfa",
]
EVENT_COLOURS = {
    "well on":              "#10b981",
    "well off":             "#ef4444",
    "maintenance shutdown": "#f59e0b",
    "pump failure":         "#f97316",
    "water breakthrough":   "#06b6d4",
}

FONT_H1 = ("Georgia",     13, "bold")
FONT_H2 = ("Georgia",     11, "bold")
FONT_H3 = ("Courier New", 10, "bold")
FONT_SM = ("Courier New",  9)
FONT_XS = ("Courier New",  8)

# ── Global state ──────────────────────────────────────────────────────────────
raw_data              = None
active_df             = None
active_X              = None
active_xcols          = []
active_anomaly_result = None
active_weight_result  = None
active_event_result   = None
active_rpm_result     = None
active_param_anom     = None   # NEW — per-parameter anomaly result
active_figures        = {}

weight_vars       = {}
weight_row_frames = []
weight_sum_var    = None
weight_sum_lbl    = None


# ══════════════════════════════════════════════════════════════════════════════
# ★ NEW — PER-PARAMETER ANOMALY ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
def analyse_per_parameter(df, x_cols, z_threshold=2.5):
    """
    For each numeric column in x_cols:
      • Compute mean, std, skew
      • Identify records where |z-score| > z_threshold
      • Classify anomaly severity and suggest a reason

    Returns a list of dicts, one per column, sorted by anomaly count desc.
    """
    import scipy.stats as sp_stats

    results = []
    for col in x_cols:
        if col not in df.columns:
            continue
        series = df[col].dropna()
        if len(series) < 5:
            continue
        vals  = series.values.astype(float)
        mean  = float(np.mean(vals))
        std   = float(np.std(vals))
        if std == 0:
            continue
        z_scores    = (vals - mean) / std
        anom_mask   = np.abs(z_scores) > z_threshold
        n_anom      = int(anom_mask.sum())
        n_total     = len(vals)
        pct         = n_anom / n_total * 100
        skew        = float(sp_stats.skew(vals))
        kurt        = float(sp_stats.kurtosis(vals))
        min_v, max_v = float(vals.min()), float(vals.max())
        median       = float(np.median(vals))

        # Severity
        if pct == 0:
            severity = "NONE"
        elif pct < 3:
            severity = "LOW"
        elif pct < 10:
            severity = "MODERATE"
        elif pct < 25:
            severity = "ELEVATED"
        else:
            severity = "HIGH"

        # --- Build a human-readable reason ---
        reason_parts = []
        if n_anom == 0:
            reason_parts.append("All values within normal statistical range.")
        else:
            # Direction of outliers
            high_out = int((z_scores > z_threshold).sum())
            low_out  = int((z_scores < -z_threshold).sum())
            if high_out > 0 and low_out > 0:
                reason_parts.append(
                    f"Outliers on BOTH ends: {high_out} records above"
                    f" {mean+z_threshold*std:.2f}"
                    f" and {low_out} records below {mean-z_threshold*std:.2f}.")
            elif high_out > 0:
                reason_parts.append(
                    f"{high_out} records exceed upper threshold"
                    f" ({mean+z_threshold*std:.2f};"
                    f" mean={mean:.2f}, sigma={std:.2f}).")
            else:
                reason_parts.append(
                    f"{low_out} records fall below lower threshold"
                    f" ({mean-z_threshold*std:.2f};"
                    f" mean={mean:.2f}, sigma={std:.2f}).")

            # Skewness interpretation
            if abs(skew) > 1.5:
                direction = "right (high-value tail)" if skew > 0 else "left (low-value tail)"
                reason_parts.append(
                    f"Distribution strongly skewed {direction} (skew={skew:.2f}) —"
                    f" suggests irregular production or measurement events.")
            elif abs(skew) > 0.75:
                reason_parts.append(
                    f"Moderate skew ({skew:.2f}) indicates non-uniform spread.")

            # Range anomaly
            range_ratio = (max_v - min_v) / (abs(mean) + 1e-9)
            if range_ratio > 5:
                reason_parts.append(
                    f"Extremely wide value range ({min_v:.2f} to {max_v:.2f})"
                    f" relative to mean —"
                    f" possible sensor fault or data entry error.")

            # Kurtosis (heavy tails)
            if kurt > 3:
                reason_parts.append(
                    f"Heavy-tailed distribution (kurtosis={kurt:.1f}) —"
                    f" extreme events or measurement spikes present.")

            # Domain-specific hints based on column name
            col_l = col.lower()
            if any(k in col_l for k in ["pressure","pres","pwh","pth","pbh"]):
                reason_parts.append(
                    "Pressure anomalies may indicate wellbore integrity issues,"
                    " plugging, or shut-in conditions.")
            elif any(k in col_l for k in ["gas","flow","rate","prod","mcf","mmscfd"]):
                reason_parts.append(
                    "Anomalous gas rates may reflect seam heterogeneity,"
                    " skin damage, or dewatering stage variation.")
            elif any(k in col_l for k in ["water","wtr","wc","bwpd","liquid"]):
                reason_parts.append(
                    "Water production anomalies often signal aquifer breakthrough"
                    " or ineffective dewatering.")
            elif any(k in col_l for k in ["rpm","speed","freq"]):
                reason_parts.append(
                    "RPM or speed outliers may point to pump wear, motor faults,"
                    " or variable-frequency drive issues.")
            elif any(k in col_l for k in ["temp","temperature","thp","tbg"]):
                reason_parts.append(
                    "Temperature anomalies can indicate Joule-Thomson cooling,"
                    " gas hydrate risk, or instrument drift.")
            elif any(k in col_l for k in ["depth","tvd","md","perforat"]):
                reason_parts.append(
                    "Depth or completion anomalies may reflect targeting errors"
                    " or multi-seam interference.")
            elif any(k in col_l for k in ["current","amp","curr"]):
                reason_parts.append(
                    "Current anomalies may indicate motor overload, cable fault,"
                    " or pump cavitation.")
            elif any(k in col_l for k in ["torque","tq","torq"]):
                reason_parts.append(
                    "Torque anomalies often indicate rod string issues,"
                    " sand ingress, or pump wear.")
            elif any(k in col_l for k in ["level","lvl","water_level","wl"]):
                reason_parts.append(
                    "Fluid level anomalies may indicate pump-off conditions,"
                    " inflow variability, or sensor malfunction.")
            elif any(k in col_l for k in ["bean","choke","orifice"]):
                reason_parts.append(
                    "Bean or choke size anomalies may reflect production"
                    " optimisation events or incorrect configuration.")
            elif any(k in col_l for k in ["voltage","volt","vlt"]):
                reason_parts.append(
                    "Voltage anomalies may indicate power supply instability"
                    " or electrical faults in the downhole motor.")

        results.append({
            "col":        col,
            "n_total":    n_total,
            "n_anom":     n_anom,
            "pct":        pct,
            "mean":       mean,
            "std":        std,
            "median":     median,
            "min":        min_v,
            "max":        max_v,
            "skew":       skew,
            "kurt":       kurt,
            "severity":   severity,
            "reason":     " ".join(reason_parts),
            "z_threshold": z_threshold,
            "high_out":   int((z_scores > z_threshold).sum()) if n_anom else 0,
            "low_out":    int((z_scores < -z_threshold).sum()) if n_anom else 0,
        })

    results.sort(key=lambda r: r["n_anom"], reverse=True)
    return results


# ══════════════════════════════════════════════════════════════════════════════
# UTILITY
# ══════════════════════════════════════════════════════════════════════════════
def cluster_hex(n):
    return [CLUSTER_PALETTE[i % len(CLUSTER_PALETTE)] for i in range(n)]

def _lighten(hex_col, amount=30):
    try:
        h = hex_col.lstrip("#")
        r,g,b = int(h[0:2],16), int(h[2:4],16), int(h[4:6],16)
        return "#{:02x}{:02x}{:02x}".format(
            min(r+amount,255), min(g+amount,255), min(b+amount,255))
    except Exception:
        return hex_col

def ts():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def set_status(msg, color=FG_DIM):
    status_var.set(msg)
    status_lbl.config(fg=color)

def _save_df(df, path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx",".xls"):
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False)


# ══════════════════════════════════════════════════════════════════════════════
# UI WIDGETS
# ══════════════════════════════════════════════════════════════════════════════
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget; self.text = text; self.tip = None
        widget.bind("<Enter>", self._show); widget.bind("<Leave>", self._hide)
    def _show(self, _=None):
        x = self.widget.winfo_rootx()+20
        y = self.widget.winfo_rooty()+self.widget.winfo_height()+4
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True); tw.wm_geometry(f"+{x}+{y}")
        tw.configure(bg=BORDER)
        tk.Label(tw, text=self.text, bg=CARD2, fg=FG_MID,
                 font=FONT_XS, padx=8, pady=4).pack()
    def _hide(self, _=None):
        if self.tip: self.tip.destroy(); self.tip = None


def make_scrollable(parent, bg=BG):
    outer  = tk.Frame(parent, bg=bg)
    canvas = Canvas(outer, bg=bg, highlightthickness=0, bd=0)
    sb     = Scrollbar(outer, orient="vertical", command=canvas.yview,
                       bg=BORDER, troughcolor=bg, activebackground=ACCENT)
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(canvas, bg=bg)
    wid   = canvas.create_window((0,0), window=inner, anchor="nw")
    def _resize(e):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(wid, width=e.width)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", _resize)
    def _wheel(e):
        if   e.num==4: canvas.yview_scroll(-1,"units")
        elif e.num==5: canvas.yview_scroll( 1,"units")
        else: canvas.yview_scroll(int(-1*(e.delta/120)),"units")
    canvas.bind_all("<MouseWheel>", _wheel)
    canvas.bind_all("<Button-4>",   _wheel)
    canvas.bind_all("<Button-5>",   _wheel)
    return outer, inner


def make_section_card(parent, title, accent_color, **pack_kw):
    wrapper = tk.Frame(parent, bg=SIDEBAR, pady=0)
    wrapper.pack(**{"fill":"x","pady":(0,2),**pack_kw})
    hdr = tk.Frame(wrapper, bg=accent_color, height=26)
    hdr.pack(fill="x"); hdr.pack_propagate(False)
    dark_text = accent_color in (COL_WEIGHTS, COL_EXPORT, COL_EVENTS)
    tk.Label(hdr, text=f"  {title}", bg=accent_color,
             fg="#000000" if dark_text else "#ffffff",
             font=FONT_H3, anchor="w").pack(side="left", fill="y", padx=4)
    body = tk.Frame(wrapper, bg=CARD, padx=10, pady=8,
                    highlightbackground=accent_color, highlightthickness=1)
    body.pack(fill="x")
    return body


def make_btn(parent, text, cmd, color=ACCENT, fg_col="#fff", tip=None):
    b = tk.Button(parent, text=text, command=cmd, bg=color, fg=fg_col,
                  activebackground=_lighten(color), activeforeground=fg_col,
                  font=FONT_H3, relief="flat", bd=0, cursor="hand2",
                  padx=10, pady=8, anchor="w")
    b.pack(fill="x", pady=2)
    if tip: Tooltip(b, tip)
    b.bind("<Enter>", lambda e: b.config(bg=_lighten(color)))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b


def _icon_btn(parent, text, cmd, bg, fg, tip=None):
    b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                  activebackground=_lighten(bg,20), activeforeground=fg,
                  font=FONT_XS, relief="flat", bd=0, cursor="hand2", padx=8, pady=5)
    b.pack(side="right", padx=3)
    if tip: Tooltip(b, tip)
    b.bind("<Enter>", lambda e: b.config(bg=_lighten(bg,20)))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b


def _divider(parent, color=BORDER):
    tk.Frame(parent, bg=color, height=1).pack(fill="x", padx=6, pady=(4,3))


def _mini_bar(parent, label, frac, col, label_width=8):
    row = tk.Frame(parent, bg=CARD); row.pack(fill="x", padx=8, pady=2)
    tk.Label(row, text=label[:label_width], bg=CARD, fg=FG_DIM,
             font=FONT_XS, width=label_width).pack(side="left")
    outer = tk.Frame(row, bg=BORDER2, height=7)
    outer.pack(side="left", fill="x", expand=True, padx=(3,0))
    tk.Frame(outer, bg=col, height=7).place(relwidth=max(frac,0.03), relheight=1.0)


# ══════════════════════════════════════════════════════════════════════════════
# PROGRESS
# ══════════════════════════════════════════════════════════════════════════════
def progress_start():
    progress_bar.config(mode="indeterminate"); progress_bar.start(12)

def progress_stop():
    progress_bar.stop(); progress_bar.config(mode="determinate"); progress_bar["value"] = 100


# ══════════════════════════════════════════════════════════════════════════════
# DYNAMIC WEIGHT PANEL
# ══════════════════════════════════════════════════════════════════════════════
def _update_weight_sum(*_):
    total = 0.0
    for v in weight_vars.values():
        try: total += float(v.get())
        except Exception: pass
    ok = abs(total - 100.0) < 0.5
    weight_sum_var.set(f"Total: {total:.1f}%  {'✔  OK' if ok else '⚠  should be 100'}")
    weight_sum_lbl.config(fg=SUCCESS2 if ok else WARN)


def rebuild_weight_panel(num_cols):
    global weight_vars, weight_row_frames
    for f in weight_row_frames:
        try: f.destroy()
        except Exception: pass
    weight_row_frames.clear()
    weight_vars.clear()
    if not num_cols:
        lbl = tk.Label(weight_body, text="No numeric columns in dataset.",
                       bg=CARD, fg=WARN, font=FONT_XS)
        lbl.pack(anchor="w")
        weight_row_frames.append(lbl)
        _update_weight_sum(); return
    default_pct = round(100.0 / len(num_cols), 1)
    for i, col in enumerate(num_cols):
        col_color = CLUSTER_PALETTE[i % len(CLUSTER_PALETTE)]
        row = tk.Frame(weight_body, bg=CARD2,
                       highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", pady=2)
        weight_row_frames.append(row)
        tk.Canvas(row, bg=col_color, width=10, height=10,
                  highlightthickness=0).pack(side="left", padx=(6,4), pady=6)
        tk.Label(row, text=col[:22], bg=CARD2, fg=col_color,
                 font=FONT_XS, width=22, anchor="w").pack(side="left")
        wvar = tk.StringVar(value=str(default_pct))
        wvar.trace_add("write", _update_weight_sum)
        weight_vars[col] = wvar
        sp = tk_ttk.Spinbox(row, from_=0.0, to=100.0, increment=1.0,
                            textvariable=wvar, width=7, font=FONT_XS)
        sp.pack(side="left", padx=4, pady=4)
        tk.Label(row, text="%", bg=CARD2, fg=FG_DIM, font=FONT_XS).pack(side="left")
    _update_weight_sum()


def get_manual_weights():
    raw = {}
    for col, v in weight_vars.items():
        try: w = float(v.get())
        except Exception: w = 0.0
        if w > 0: raw[col] = w
    total = sum(raw.values())
    if total <= 0: return {}, 0.0
    return {c: round(v/total*100, 2) for c,v in raw.items()}, total


# ══════════════════════════════════════════════════════════════════════════════
# MODULE ① — DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════
def upload_dataset():
    global raw_data
    f = filedialog.askopenfilename(
        title="Open Data File",
        filetypes=[
            ("All supported files",
             "*.csv *.tsv *.txt *.xlsx *.xls *.xlsb *.xlsm *.ods "
             "*.json *.parquet *.feather *.h5 *.hdf5 *.pkl *.pickle *.mdb *.accdb"),
            ("Microsoft Access 2007+", "*.mdb *.accdb"),
            ("CSV / TSV",              "*.csv *.tsv *.txt"),
            ("Excel",                  "*.xlsx *.xls *.xlsb *.xlsm"),
            ("ODS (LibreOffice)",       "*.ods"),
            ("JSON",                   "*.json"),
            ("Parquet / Feather",      "*.parquet *.feather"),
            ("HDF5",                   "*.h5 *.hdf5"),
            ("Pickle",                 "*.pkl *.pickle"),
            ("All files",              "*.*"),
        ]
    )
    if not f: return
    size_mb = os.path.getsize(f) / (1024*1024)
    set_status(f"⟳  Loading {os.path.basename(f)}  ({size_mb:.1f} MB) …", WARN)
    progress_start()

    def _load():
        global raw_data
        try:
            df, source_label = load_any_file(f)
            if len(df) == 0:
                raise RuntimeError("File loaded but contains no rows of data.")
            if len(df.columns) == 0:
                raise RuntimeError("File loaded but contains no columns.")
            raw_data = df
            num_cols = list(df.select_dtypes(include="number").columns)
            app.after(0, lambda: _post_load(df, f, num_cols, source_label))
        except Exception as e:
            err = str(e)
            app.after(0, lambda: _load_error(err))

    def _post_load(df, filepath, num_cols, source_label):
        id_keywords = ["wellid","well_id","well id","wellname","well_name",
                       "well no","wellno","uwi","api","id"]
        id_col = None
        for c in df.columns:
            if c.lower().replace(" ","").replace("_","") in \
               [k.replace(" ","").replace("_","") for k in id_keywords]:
                id_col = c; break
        if id_col is None:
            for c in df.columns:
                if "well" in c.lower() or c.lower() == "id":
                    id_col = c; break
        n_rows = len(df)
        n_unique_wells = df[id_col].nunique() if id_col else None
        # Sidebar "Unique Wells" shows actual distinct well count
        rows_var.set(f"{n_unique_wells:,}" if n_unique_wells is not None else f"{n_rows:,} rows")
        cols_var.set(str(len(df.columns)))
        num_cols_var.set(f"{len(num_cols)} numeric")
        rebuild_weight_panel(num_cols)

        ec = _find_event_column(df)
        if ec:
            event_col_var.set(f"✔  Event column: '{ec}'")
            event_col_lbl.config(fg=SUCCESS2)
        else:
            event_col_var.set("⚠  No event/status column found")
            event_col_lbl.config(fg=WARN)
        preview_table(df)
        progress_stop()
        size_mb = os.path.getsize(filepath) / (1024*1024)
        set_status(
            f"✔  {source_label}  —  {n_rows:,} rows × {len(df.columns)} cols  "
            f"({size_mb:.1f} MB)", SUCCESS)
        for k in stat_widgets: stat_widgets[k].set("—")
        # Show unique well count separately from total record count
        if n_unique_wells is not None:
            stat_widgets["Unique Wells"].set(f"{n_unique_wells:,}")
            stat_widgets["Total Records"].set(f"{n_rows:,}")
        else:
            stat_widgets["Unique Wells"].set("—")
            stat_widgets["Total Records"].set(f"{n_rows:,}")
        export_btn.config(state="normal")

    def _load_error(msg):
        progress_stop()
        set_status("✖  Load failed — see error dialog", ERR)
        messagebox.showerror("File Load Error", msg)

    threading.Thread(target=_load, daemon=True).start()


def _find_event_column(df):
    keywords = ["status","state","event","mode","condition","operation","type"]
    for c in df.columns:
        cl = c.lower().replace(" ","_").replace("-","_")
        if any(k in cl for k in keywords):
            return c
    return None


def _find_rpm_column(df):
    for c in df.columns:
        if "rpm" in c.lower():
            return c
    return None


def add_rpm_status_column(df):
    rpm_col = _find_rpm_column(df)
    if rpm_col is None:
        return df, None, {"found": False, "col": None,
                          "running": 0, "stopped": 0, "reverse": 0, "no_data": 0}
    rpm = df[rpm_col].copy()
    def _classify_rpm(v):
        try:
            v = float(v)
            if   v > 0: return "Running"
            elif v < 0: return "Reverse"
            else:       return "Stopped"
        except (TypeError, ValueError):
            return "No Data"
    df = df.copy()
    df["RPM_Status"] = rpm.apply(_classify_rpm)
    counts = df["RPM_Status"].value_counts().to_dict()
    return df, rpm_col, {
        "found":    True, "col": rpm_col,
        "running":  counts.get("Running",  0),
        "stopped":  counts.get("Stopped",  0),
        "reverse":  counts.get("Reverse",  0),
        "no_data":  counts.get("No Data",  0),
    }


def preview_table(df):
    for w in table_frame.winfo_children(): w.destroy()
    style = tk_ttk.Style()
    style.configure("P.Treeview", background=CARD, foreground=FG,
                    fieldbackground=CARD, rowheight=21, font=FONT_SM)
    style.configure("P.Treeview.Heading", background="#0e2040",
                    foreground=COL_PREVIEW, font=FONT_H3)
    style.map("P.Treeview", background=[("selected", ACCENT)])
    tv  = tk_ttk.Treeview(table_frame, style="P.Treeview")
    vsb = tk_ttk.Scrollbar(table_frame, orient="vertical",   command=tv.yview)
    hsb = tk_ttk.Scrollbar(table_frame, orient="horizontal", command=tv.xview)
    tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tv["columns"] = list(df.columns)
    tv["show"]    = "headings"
    for col in df.columns:
        tv.heading(col, text=col); tv.column(col, width=90, minwidth=60)
    for row in df.head(20).values:
        tv.insert("","end", values=list(row))
    hsb.pack(side="bottom", fill="x")
    vsb.pack(side="right",  fill="y")
    tv.pack(fill="both", expand=True)


# ══════════════════════════════════════════════════════════════════════════════
# MODULE ② — WEIGHTED CLUSTERING
# ══════════════════════════════════════════════════════════════════════════════
def build_weighted_X(df, x_cols, weight_map):
    cols    = [c for c in x_cols if c in df.columns]
    Xraw    = df[cols].values.astype(float)
    mu, s   = Xraw.mean(axis=0), Xraw.std(axis=0)
    s[s==0] = 1
    X_scaled = (Xraw - mu) / s
    w_vec    = np.array([np.sqrt(weight_map.get(c,1.0)/100.0) for c in cols])
    return X_scaled * w_vec, cols


def run_clustering(X_w, n_clusters):
    from sklearn.cluster import KMeans
    n  = min(n_clusters, X_w.shape[0])
    km = KMeans(n_clusters=n, random_state=42, n_init="auto")
    return km.fit_predict(X_w), km.inertia_


# ══════════════════════════════════════════════════════════════════════════════
# MODULE ③ — ANOMALY DETECTION
# ══════════════════════════════════════════════════════════════════════════════
def detect_anomalies(X_w, contamination=0.05, method="iforest"):
    from sklearn.ensemble import IsolationForest
    from sklearn.neighbors import LocalOutlierFactor
    n = X_w.shape[0]
    if n < 5:
        labels = np.ones(n, dtype=int)
        return labels, 0.0, [], "N/A (too few samples)", np.zeros(n)
    safe_cont = float(np.clip(contamination, 0.001, 0.499))
    min_cont  = max(safe_cont, 1.0/n)
    safe_cont = min(min_cont, 0.499)
    if method == "lof":
        k   = max(5, min(20, n // 10))
        det = LocalOutlierFactor(n_neighbors=k, contamination=safe_cont)
        labels = det.fit_predict(X_w)
        scores = det.negative_outlier_factor_
        name   = f"Local Outlier Factor (LOF, k={k})"
    else:
        n_est  = 100 if n <= 500 else 200
        det    = IsolationForest(n_estimators=n_est, contamination=safe_cont,
                                 random_state=42, n_jobs=-1)
        labels = det.fit_predict(X_w)
        scores = det.decision_function(X_w)
        name   = f"Isolation Forest (n_est={n_est})"
    idx = list(np.where(labels==-1)[0])
    pct = len(idx)/n*100
    return labels, pct, idx, name, scores


# ══════════════════════════════════════════════════════════════════════════════
# MODULE ④ — EVENT ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
def analyse_events(df, anomaly_labels):
    ec = _find_event_column(df)
    if ec is None:
        return {
            "has_events": False, "event_col": None,
            "event_counts": {}, "event_labels": np.full(len(df),"Unknown"),
            "n_active": 0, "n_inactive": 0,
            "n_abnormal": int((anomaly_labels==-1).sum()),
            "operational_anomalies": [],
            "message": "No event/status column found in dataset.",
        }
    raw_vals     = df[ec].astype(str).str.strip()
    event_labels = raw_vals.values
    from collections import Counter
    event_counts = dict(Counter(event_labels))
    active_kw    = ["on","active","running","producing","open"]
    inactive_kw  = ["off","inactive","shut","stop","closed","idle","down"]
    def _classify(val):
        vl = val.lower()
        if any(k in vl for k in active_kw):   return "active"
        if any(k in vl for k in inactive_kw): return "inactive"
        return "other"
    statuses   = raw_vals.apply(_classify)
    n_active   = int((statuses=="active").sum())
    n_inactive = int((statuses=="inactive").sum())
    n_abnormal = int((anomaly_labels==-1).sum())
    op_anomalies = []
    id_col = next((c for c in df.columns if "id" in c.lower()), None)
    if id_col:
        for wid, grp in df.groupby(id_col):
            states = raw_vals[grp.index].apply(_classify).unique()
            if "active" in states and "inactive" in states:
                op_anomalies.append(str(wid))
    return {
        "has_events": True, "event_col": ec,
        "event_counts": event_counts, "event_labels": event_labels,
        "n_active": n_active, "n_inactive": n_inactive,
        "n_abnormal": n_abnormal, "operational_anomalies": op_anomalies,
        "message": f"Events read from column: '{ec}'",
    }


# ══════════════════════════════════════════════════════════════════════════════
# MODULE ⑤ — VISUALIZATIONS
# ══════════════════════════════════════════════════════════════════════════════
def _style_ax(ax):
    ax.set_facecolor("#161d2e")
    for sp in ax.spines.values(): sp.set_edgecolor("#2d3f5e")
    ax.tick_params(colors="#94a3b8", labelsize=8)


def plot_clusters(df, xcols, ycols, hex_colors):
    fig, ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    x_col = xcols[0] if xcols else None
    y_col = (ycols[0] if ycols else (xcols[1] if len(xcols)>1 else xcols[0] if xcols else None))
    if x_col and y_col and x_col in df.columns and y_col in df.columns:
        for cl in sorted(df["cluster"].unique()):
            mask=df["cluster"]==cl; col=hex_colors[int(cl)%len(hex_colors)]
            ax.scatter(df.loc[mask,x_col], df.loc[mask,y_col], color=col, s=55,
                       edgecolors="#ffffff22", linewidths=0.5, zorder=3,
                       label=f"Cluster {cl}  ({int(mask.sum()):,} records)")
        ax.set_xlabel(x_col, fontsize=9, labelpad=6)
        ax.set_ylabel(y_col, fontsize=9, labelpad=6)
    else:
        ax.text(0.2,0.5,"Select X and Y features to plot",
                transform=ax.transAxes,color=FG_DIM,fontsize=10)
    ax.set_title("CBM Production Clusters",color=COL_CLUSTER,
                 fontsize=13,pad=12,fontweight="bold")
    ax.legend(loc="upper left",fontsize=8.5,framealpha=0.92,
              edgecolor="#334155",facecolor="#161d2e",labelcolor="#e2e8f0",markerscale=1.4)
    return fig


def plot_pca(X, labels, hex_colors):
    from sklearn.decomposition import PCA
    fig,ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    if X.shape[0]<2 or X.shape[1]<1:
        ax.text(0.3,0.5,"Not enough data for PCA",transform=ax.transAxes,color=FG_DIM)
        return fig
    n = min(2, X.shape[1])
    Z = PCA(n_components=n).fit_transform(X)
    if Z.shape[1]==1: Z = np.hstack([Z, np.zeros_like(Z)])
    for cl in sorted(np.unique(labels)):
        mask=labels==cl; col=hex_colors[int(cl)%len(hex_colors)]
        ax.scatter(Z[mask,0],Z[mask,1],color=col,s=45,
                   edgecolors="#ffffff22",linewidths=0.5,zorder=3,
                   label=f"Cluster {cl}  ({int(mask.sum()):,} records)")
    ax.set_xlabel("Principal Component 1",fontsize=9,labelpad=6)
    ax.set_ylabel("Principal Component 2",fontsize=9,labelpad=6)
    ax.set_title("PCA — Well Feature Space",color=COL_CLUSTER,
                 fontsize=13,pad=12,fontweight="bold")
    ax.margins(0.08)
    ax.legend(loc="upper left",fontsize=8.5,framealpha=0.92,
              edgecolor="#334155",facecolor="#161d2e",labelcolor="#e2e8f0",markerscale=1.4)
    return fig


def plot_reservoir_3d(df, xcols, ycols, hex_colors):
    from mpl_toolkits.mplot3d import Axes3D  # noqa
    fig = plt.figure(figsize=(7.5,4.5))
    ax  = fig.add_subplot(111,projection="3d")
    ax.set_facecolor("#161d2e"); ax.tick_params(colors="#94a3b8",labelsize=7)
    cols3 = list(dict.fromkeys(xcols+ycols))[:3]
    if len(cols3)<3:
        extra = [c for c in df.select_dtypes(include="number").columns
                 if c not in cols3 and c!="cluster"]
        cols3 = (cols3+extra)[:3]
    if len(cols3)==3 and all(c in df.columns for c in cols3):
        handles=[]
        for cl in sorted(df["cluster"].unique()):
            mask=df["cluster"]==cl; col=hex_colors[int(cl)%len(hex_colors)]
            ax.scatter(df.loc[mask,cols3[0]],df.loc[mask,cols3[1]],df.loc[mask,cols3[2]],
                       color=col,s=28,edgecolors="#ffffff15",linewidths=0.3)
            handles.append(mpatches.Patch(color=col,
                           label=f"Cluster {cl}  ({int(mask.sum()):,})"))
        ax.set_xlabel(cols3[0],color="#94a3b8",fontsize=7,labelpad=3)
        ax.set_ylabel(cols3[1],color="#94a3b8",fontsize=7,labelpad=3)
        ax.set_zlabel(cols3[2],color="#94a3b8",fontsize=7,labelpad=3)
        ax.legend(handles=handles,loc="upper left",fontsize=8,framealpha=0.85,
                  facecolor="#161d2e",edgecolor="#334155",labelcolor="#e2e8f0")
    else:
        ax.text2D(0.15,0.5,"Need ≥ 3 numeric features",
                  transform=ax.transAxes,color=FG_DIM)
    ax.set_title("3D Reservoir Map",color=COL_DATASET,
                 fontsize=13,pad=8,fontweight="bold")
    return fig


def plot_production(df, ycols, xcols, hex_colors):
    fig,ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    y_col = next((c for c in ycols if c in df.columns),None) or \
            next((c for c in xcols if c in df.columns),None)
    x_col = next((c for c in xcols if c in df.columns and c!=y_col),None)
    if y_col is None:
        ax.text(0.2,0.5,"Select a Y feature for production curve",
                transform=ax.transAxes,color=FG_DIM,fontsize=10)
        return fig
    for cl in sorted(df["cluster"].unique()):
        sub=df[df["cluster"]==cl].reset_index(drop=True)
        col=hex_colors[int(cl)%len(hex_colors)]
        xs = sub[x_col].values if x_col else np.arange(len(sub))
        ax.plot(xs,sub[y_col].values,color=col,linewidth=2,alpha=0.88,
                label=f"Cluster {cl}  ({len(sub):,} records)")
    ax.set_xlabel(x_col if x_col else "Well Index",fontsize=9,labelpad=6)
    ax.set_ylabel(y_col,fontsize=9,labelpad=6)
    ax.set_title("Production Curves by Cluster",color=SUCCESS2,
                 fontsize=13,pad=12,fontweight="bold")
    ax.margins(0.04,0.10)
    ax.legend(loc="upper left",fontsize=8.5,framealpha=0.92,
              edgecolor="#334155",facecolor="#161d2e",labelcolor="#e2e8f0")
    return fig


def plot_hidden_patterns(X_w, anomaly_labels, xcols, detector_name):
    from sklearn.decomposition import PCA
    fig,ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    n = X_w.shape[0]
    if X_w.shape[1]>=2:
        Z = PCA(n_components=2,random_state=42).fit_transform(X_w)
        xlabel,ylabel="PC 1  (feature projection)","PC 2  (feature projection)"
    elif X_w.shape[1]==1:
        Z = np.column_stack([X_w[:,0],np.arange(n)])
        xlabel = xcols[0] if xcols else "Feature"; ylabel="Well Index"
    else:
        ax.text(0.3,0.5,"No feature data",transform=ax.transAxes,color=FG_DIM); return fig
    nm=anomaly_labels==1; am=anomaly_labels==-1
    ax.scatter(Z[nm,0],Z[nm,1],color=NORMAL_C,s=30,alpha=0.70,
               edgecolors="#ffffff18",linewidths=0.3,zorder=3,
               label=f"● Normal records  ({int(nm.sum()):,})")
    if am.sum()>0:
        ax.scatter(Z[am,0],Z[am,1],color=ANOMALY_C,s=70,alpha=0.95,
                   edgecolors="#ffffff66",linewidths=0.9,marker="D",zorder=5,
                   label=f"◆ Anomalous records  ({int(am.sum()):,})")
        ax.annotate(
            f"◆ {int(am.sum())} anomalous record{'s' if am.sum()!=1 else ''} detected",
            xy=(Z[am,0].mean(),Z[am,1].mean()),xytext=(14,14),
            textcoords="offset points",color=ANOMALY_C,fontsize=9,fontweight="bold",
            arrowprops=dict(arrowstyle="->",color=ANOMALY_C,lw=1.0))
    ax.legend(loc="upper right",fontsize=9,framealpha=0.92,
              edgecolor="#334155",facecolor="#161d2e",labelcolor="#e2e8f0",markerscale=1.3)
    ax.set_xlabel(xlabel,fontsize=9,labelpad=6); ax.set_ylabel(ylabel,fontsize=9,labelpad=6)
    ax.set_title(f"Hidden Pattern Detection  ·  {detector_name}",
                 color=ANOMALY_C,fontsize=13,pad=12,fontweight="bold")
    ax.margins(0.10); return fig


def plot_weight_chart(weight_map, method_desc):
    if not weight_map:
        fig,ax=plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
        ax.text(0.3,0.5,"Run analysis first",transform=ax.transAxes,color=FG_DIM,fontsize=11)
        ax.set_title("Parameter Importance",color=COL_WEIGHTS,
                     fontsize=13,pad=12,fontweight="bold")
        return fig
    params=list(weight_map.keys()); values=[weight_map[p] for p in params]
    sorted_pairs=sorted(zip(values,params),reverse=False)
    values_s=[v for v,_ in sorted_pairs]; params_s=[p for _,p in sorted_pairs]
    colors=[CLUSTER_PALETTE[i%len(CLUSTER_PALETTE)] for i in range(len(params_s))]
    fig=plt.figure(figsize=(7.5,4.5))
    gs=gridspec.GridSpec(2,1,height_ratios=[5,1],hspace=0.45)
    ax=fig.add_subplot(gs[0]); _style_ax(ax)
    ax_note=fig.add_subplot(gs[1]); ax_note.axis("off")
    bars=ax.barh(params_s,values_s,color=colors,
                 edgecolor="#ffffff22",linewidth=0.5,height=0.6)
    for bar,val in zip(bars,values_s):
        ax.text(bar.get_width()+0.3,bar.get_y()+bar.get_height()/2,
                f"{val:.1f}%",va="center",ha="left",color=FG_MID,fontsize=8.5,fontweight="bold")
    max_idx=values_s.index(max(values_s))
    bars[max_idx].set_edgecolor("#ffffff88"); bars[max_idx].set_linewidth(1.5)
    ax.text(values_s[max_idx]/2,bars[max_idx].get_y()+bars[max_idx].get_height()/2,
            "  * Most Influential",va="center",ha="left",
            color="#ffffff",fontsize=7.5,style="italic")
    ax.set_xlabel("Assigned Weight (%)",fontsize=9,labelpad=6)
    ax.set_title("Parameter Importance  —  User-Assigned Weights",
                 color=COL_WEIGHTS,fontsize=12,pad=12,fontweight="bold")
    ax.set_xlim(0,max(values_s)*1.22); ax.margins(0.03,0.15)
    wrapped=method_desc if len(method_desc)<=90 else method_desc[:87]+"…"
    ax_note.text(0.01,0.8,f"Method: {wrapped}",
                 transform=ax_note.transAxes,color=FG_DIM,fontsize=7.5,va="top")
    fig.patch.set_facecolor("#0d1117"); return fig


# ══════════════════════════════════════════════════════════════════════════════
# ★ NEW — PARAMETER ANOMALY CHART
# ══════════════════════════════════════════════════════════════════════════════
SEV_COLORS = {
    "NONE":     "#2d3f5e",
    "LOW":      "#10b981",
    "MODERATE": "#f59e0b",
    "ELEVATED": "#f97316",
    "HIGH":     "#ef4444",
}

def plot_param_anomaly(param_anom_list, df=None, anomaly_labels=None):
    """
    Per-parameter Normal vs Anomalous chart.
    LEFT  : horizontal stacked bar — exact counts + % per parameter.
    RIGHT : box plots per parameter (Normal vs Anomalous actual values).
    All text is clearly spaced, no overlapping titles.
    Anomaly groups come from actual Isolation Forest labels (anomaly_labels).
    """
    import textwrap as _tw

    # ── helpers ────────────────────────────────────────────────────────────────
    def _fmt_median_diff(norm_vals, anom_vals):
        """Return a readable median-difference string, no absurd numbers."""
        if len(norm_vals) == 0 or len(anom_vals) == 0:
            return ""
        nm = float(np.median(norm_vals))
        am = float(np.median(anom_vals))
        abs_diff = abs(am - nm)
        if abs(nm) < 1e-9:
            # avoid division by near-zero — show absolute diff only
            return f"Median shift: {abs_diff:.4g} (abs)"
        pct = abs_diff / abs(nm) * 100
        if pct > 9999:
            return f"Median shift: {abs_diff:.4g}  (>{int(pct/1000)}k%)"
        return f"Median shift: {pct:.1f}%"

    # ── empty state ────────────────────────────────────────────────────────────
    if not param_anom_list:
        fig, ax = plt.subplots(figsize=(10, 5))
        fig.patch.set_facecolor("#0d1117"); _style_ax(ax)
        ax.text(0.35, 0.5, "Run analysis first",
                transform=ax.transAxes, color=FG_DIM, fontsize=12)
        ax.set_title("Per-Parameter: Normal vs Anomalous Records",
                     color=COL_PARAM_ANOM, fontsize=12, pad=10, fontweight="bold")
        return fig

    items = [r for r in param_anom_list if r["n_total"] > 0]
    if not items:
        fig, ax = plt.subplots(figsize=(10, 5))
        fig.patch.set_facecolor("#0d1117"); _style_ax(ax)
        ax.text(0.3, 0.5, "No usable parameters",
                transform=ax.transAxes, color=FG_DIM)
        return fig

    n_params  = len(items)
    have_data = (df is not None and anomaly_labels is not None)

    # ── figure: fixed height per parameter, generous margins ──────────────────
    row_h   = 1.05          # inches per parameter row
    fig_h   = max(6.0, n_params * row_h + 2.2)
    fig_w   = 14.0
    fig     = plt.figure(figsize=(fig_w, fig_h))
    fig.patch.set_facecolor("#0d1117")

    # Title at top — well above everything else
    fig.text(0.5, 0.985,
             "Per-Parameter Analysis:  Which Parameter is Normal / Anomalous?",
             ha="center", va="top",
             color=COL_PARAM_ANOM, fontsize=11, fontweight="bold")

    # Outer grid: LEFT bar chart | RIGHT box plots
    # Left gets 38% width, right gets 58%, gap 4%
    gs = gridspec.GridSpec(
        1, 2,
        width_ratios=[2, 3],
        wspace=0.06,
        left=0.01, right=0.99,
        top=0.93, bottom=0.07,
    )

    # ── LEFT: stacked bar chart ────────────────────────────────────────────────
    ax_bar = fig.add_subplot(gs[0])
    _style_ax(ax_bar)

    y_pos     = np.arange(n_params)
    # Use actual Isolation Forest label counts if available
    if have_data:
        norm_counts = []
        anom_counts = []
        for r in items:
            col = r["col"]
            if col in df.columns:
                valid_idx = df[col].dropna().index
                n_n = int((anomaly_labels[valid_idx] ==  1).sum())
                n_a = int((anomaly_labels[valid_idx] == -1).sum())
            else:
                n_n = r["n_total"] - r["n_anom"]
                n_a = r["n_anom"]
            norm_counts.append(n_n)
            anom_counts.append(n_a)
    else:
        norm_counts = [r["n_total"] - r["n_anom"] for r in items]
        anom_counts = [r["n_anom"]                for r in items]

    totals    = [n + a for n, a in zip(norm_counts, anom_counts)]
    sev_cols  = [SEV_COLORS.get(r["severity"], "#8b5cf6") for r in items]
    # Full parameter names — no truncation for y-axis (handled by tick font size)
    param_labels = [r["col"] for r in items]

    # Draw bars
    ax_bar.barh(y_pos, norm_counts, height=0.60,
                color=NORMAL_C, edgecolor="#ffffff18", linewidth=0.3,
                label="Normal records", zorder=3)
    ax_bar.barh(y_pos, anom_counts, height=0.60,
                left=norm_counts, color="#ef4444",
                edgecolor="#ffffff22", linewidth=0.3,
                label="Anomalous records", zorder=3)

    # Annotate counts inside bars — only if bar is wide enough
    max_total = max(totals) if totals else 1
    for i in range(n_params):
        nn, na, tot = norm_counts[i], anom_counts[i], totals[i]
        pct_a = na / max(tot, 1) * 100
        pct_n = 100.0 - pct_a

        # Normal label — inside blue bar if >18% of total width
        if nn / max_total > 0.18:
            ax_bar.text(nn / 2, i,
                        f"{nn:,}  ({pct_n:.1f}%)",
                        va="center", ha="center",
                        color="#ffffff", fontsize=7.5, fontweight="bold", zorder=5)

        # Anomalous label — inside red bar if space, else just outside
        if na > 0:
            red_center = nn + na / 2
            label_str  = f"{na:,}  ({pct_a:.1f}%)"
            if na / max_total > 0.10:
                ax_bar.text(red_center, i, label_str,
                            va="center", ha="center",
                            color="#ffffff", fontsize=7.5, fontweight="bold", zorder=5)
            else:
                ax_bar.text(tot + max_total * 0.01, i, label_str,
                            va="center", ha="left",
                            color="#ef4444", fontsize=7, zorder=5)

        # Severity badge — right of bar
        verdict   = "✔ CLEAN" if items[i]["severity"] == "NONE" else items[i]["severity"]
        sev_color = sev_cols[i]
        ax_bar.text(max_total * 1.02, i,
                    verdict,
                    va="center", ha="left",
                    color=sev_color, fontsize=7.5, fontweight="bold")

    # Y-axis: full parameter names, auto-size font
    ax_bar.set_yticks(y_pos)
    # Dynamically shrink font if many params
    ytick_fs = max(6.5, 9.0 - n_params * 0.25)
    ax_bar.set_yticklabels(param_labels, fontsize=ytick_fs, color=FG_MID)
    ax_bar.tick_params(axis="y", length=0, pad=4)
    ax_bar.set_xlabel("Record Count", fontsize=8.5, labelpad=5, color="#94a3b8")
    ax_bar.set_xlim(0, max_total * 1.25)
    ax_bar.set_ylim(-0.6, n_params - 0.4)
    ax_bar.invert_yaxis()
    ax_bar.tick_params(axis="x", labelsize=7, colors="#94a3b8")
    # No top title on the bar — the figure suptitle covers it
    ax_bar.set_title("Normal vs Anomalous\nper Parameter",
                     color="#94a3b8", fontsize=8.5, pad=6, loc="center")
    ax_bar.legend(loc="lower right", fontsize=7.5, framealpha=0.88,
                  facecolor="#161d2e", edgecolor="#334155", labelcolor="#e2e8f0")

    # ── RIGHT: per-parameter box plots ────────────────────────────────────────
    if have_data:
        # Each param gets its own subplot row
        gs_right = gridspec.GridSpecFromSubplotSpec(
            n_params, 1,
            subplot_spec=gs[1],
            hspace=0.55,        # generous vertical gap between rows
        )
        rng = np.random.default_rng(42)

        for i, r in enumerate(items):
            col = r["col"]
            ax_box = fig.add_subplot(gs_right[i])
            _style_ax(ax_box)

            if col not in df.columns:
                ax_box.text(0.3, 0.5, "column not found",
                            transform=ax_box.transAxes, color=FG_DIM, fontsize=7)
                continue

            # Split by actual Isolation Forest labels
            norm_vals = df.loc[anomaly_labels ==  1, col].dropna().values.astype(float)
            anom_vals = df.loc[anomaly_labels == -1, col].dropna().values.astype(float)

            # Sample for performance
            if len(norm_vals) > 4000:
                norm_vals = rng.choice(norm_vals, 4000, replace=False)
            if len(anom_vals) > 2000:
                anom_vals = rng.choice(anom_vals, 2000, replace=False)

            data_to_plot = []; bp_labels = []; bp_colors = []
            if len(norm_vals) > 0:
                data_to_plot.append(norm_vals)
                bp_labels.append("Normal")
                bp_colors.append(NORMAL_C)
            if len(anom_vals) > 0:
                data_to_plot.append(anom_vals)
                bp_labels.append("Anomalous")
                bp_colors.append("#ef4444")

            if data_to_plot:
                bp = ax_box.boxplot(
                    data_to_plot,
                    vert=False,
                    patch_artist=True,
                    widths=0.42,
                    showfliers=True,
                    flierprops=dict(marker=".", markersize=2.5,
                                   markerfacecolor="#ef444488",
                                   markeredgecolor="none", alpha=0.45),
                    medianprops=dict(color="#ffffff", linewidth=1.8),
                    whiskerprops=dict(color="#94a3b8", linewidth=0.9),
                    capprops=dict(color="#94a3b8", linewidth=0.9),
                    boxprops=dict(linewidth=0.8),
                )
                for patch, c in zip(bp["boxes"], bp_colors):
                    patch.set_facecolor(c); patch.set_alpha(0.70)

            ax_box.set_yticks(range(1, len(bp_labels) + 1))
            ax_box.set_yticklabels(bp_labels, fontsize=7.5, color=FG_MID)
            ax_box.tick_params(axis="x", labelsize=6.5, colors="#94a3b8",
                               pad=2, length=3)

            # Title: parameter name + verdict on separate line — no overlap
            sev_c   = SEV_COLORS.get(r["severity"], "#8b5cf6")
            verdict = ("CLEAN" if r["severity"] == "NONE"
                       else f"{r['severity']}  —  {anom_counts[i]:,} anomalous recs")
            # Use two-line title: param name on line 1, verdict on line 2
            ax_box.set_title(
                f"{col}\n[{verdict}]",
                color=sev_c, fontsize=7.5, pad=3,
                fontweight="bold", loc="left",
                linespacing=1.3,
            )

            # Median diff — bottom-right, compact
            if len(norm_vals) > 0 and len(anom_vals) > 0:
                diff_str = _fmt_median_diff(norm_vals, anom_vals)
                ax_box.text(0.99, -0.28, diff_str,
                            transform=ax_box.transAxes,
                            va="top", ha="right",
                            color="#f59e0b", fontsize=6.5)

    else:
        # Fallback scatter if no df passed
        ax_sc = fig.add_subplot(gs[1]); _style_ax(ax_sc)
        pcts  = [r["pct"]       for r in items]
        skews = [abs(r["skew"]) for r in items]
        ax_sc.scatter(skews, pcts, c=sev_cols, s=70,
                      edgecolors="#ffffff33", linewidths=0.6, zorder=4)
        for i, r in enumerate(items):
            ax_sc.annotate(r["col"][:18], (skews[i], pcts[i]),
                           xytext=(5, 4), textcoords="offset points",
                           color=FG_DIM, fontsize=7)
        ax_sc.set_xlabel("|Skewness|", fontsize=9, labelpad=6)
        ax_sc.set_ylabel("Anomaly %",  fontsize=9, labelpad=6)
        ax_sc.set_title("Anomaly % vs Skewness per Parameter",
                        color=COL_PARAM_ANOM, fontsize=10, pad=8, fontweight="bold")
        ax_sc.margins(0.18)

    return fig



def build_param_anomaly_legend(parent, param_anom_list):
    """Right-side legend: per-parameter Normal vs Anomalous record counts."""
    for w in parent.winfo_children(): w.destroy()

    tk.Label(parent, text="PARAM STATUS",
             bg=CARD, fg=COL_PARAM_ANOM,
             font=FONT_H3, justify="center").pack(pady=(8,2), padx=6)
    tk.Label(parent, text="Normal vs Anomalous\nper parameter",
             bg=CARD, fg=FG_DIM,
             font=("Courier New", 7), justify="center").pack(pady=(0,3))
    _divider(parent, COL_PARAM_ANOM)

    if not param_anom_list:
        tk.Label(parent, text="Run analysis first", bg=CARD, fg=FG_DIM,
                 font=FONT_XS, wraplength=170).pack(pady=8)
        return

    for r in param_anom_list:
        if r["n_total"] == 0:
            continue

        sev_col  = SEV_COLORS.get(r["severity"], "#8b5cf6")
        n_norm   = r["n_total"] - r["n_anom"]
        n_anom   = r["n_anom"]
        pct_a    = r["pct"]
        pct_n    = 100.0 - pct_a
        is_clean = (n_anom == 0)

        # Card per parameter
        card = tk.Frame(parent, bg=CARD2, padx=5, pady=4,
                        highlightbackground=sev_col if not is_clean else SUCCESS,
                        highlightthickness=1)
        card.pack(fill="x", padx=5, pady=2)

        # Parameter name
        tk.Label(card, text=r["col"][:22],
                 bg=CARD2, fg=sev_col if not is_clean else SUCCESS2,
                 font=("Courier New", 8, "bold"),
                 anchor="w", wraplength=160).pack(anchor="w")

        # Normal row
        norm_row = tk.Frame(card, bg=CARD2); norm_row.pack(fill="x")
        tk.Label(norm_row, text="● Normal  :",
                 bg=CARD2, fg=NORMAL_C,
                 font=("Courier New", 7), width=11, anchor="w").pack(side="left")
        tk.Label(norm_row, text=f"{n_norm:,}  ({pct_n:.1f}%)",
                 bg=CARD2, fg=FG_MID,
                 font=("Courier New", 7), anchor="w").pack(side="left")

        # Anomalous row
        if n_anom > 0:
            anom_row = tk.Frame(card, bg=CARD2); anom_row.pack(fill="x")
            tk.Label(anom_row, text="◆ Anomaly :",
                     bg=CARD2, fg="#ef4444",
                     font=("Courier New", 7), width=11, anchor="w").pack(side="left")
            tk.Label(anom_row, text=f"{n_anom:,}  ({pct_a:.1f}%)",
                     bg=CARD2, fg="#ef4444",
                     font=("Courier New", 7, "bold"), anchor="w").pack(side="left")
        else:
            tk.Label(card, text="✔ All records normal",
                     bg=CARD2, fg=SUCCESS2,
                     font=("Courier New", 7), anchor="w").pack(anchor="w")

        # Mini stacked bar: blue | red
        bar_frame = tk.Frame(card, bg=BORDER2, height=5)
        bar_frame.pack(fill="x", pady=(3, 0))
        norm_frac = n_norm / max(r["n_total"], 1)
        anom_frac = 1.0 - norm_frac
        norm_bar  = tk.Frame(bar_frame, bg=NORMAL_C, height=5)
        norm_bar.place(relwidth=norm_frac, relheight=1.0, relx=0)
        if anom_frac > 0.01:
            anom_bar = tk.Frame(bar_frame, bg="#ef4444", height=5)
            anom_bar.place(relwidth=anom_frac, relheight=1.0, relx=norm_frac)

    _divider(parent)
    # Summary
    flagged = sum(1 for r in param_anom_list if r["n_anom"] > 0)
    clean   = len(param_anom_list) - flagged
    tk.Label(parent,
             text=f"Flagged : {flagged} params\nClean   : {clean} params",
             bg=CARD, fg=FG_DIM,
             font=("Courier New", 7), justify="left").pack(anchor="w", padx=6, pady=(3,4))


def build_weight_legend(parent, weight_map, method_desc):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="PARAM WEIGHTS",bg=CARD,fg=COL_WEIGHTS,
             font=FONT_H3,justify="center").pack(pady=(8,2),padx=6)
    tk.Label(parent,text="★ = most influential",bg=CARD,fg=FG_DIM,
             font=("Courier New",7),justify="center").pack(pady=(0,3))
    _divider(parent,COL_WEIGHTS)
    if not weight_map:
        tk.Label(parent,text="Run analysis first",bg=CARD,fg=FG_DIM,
                 font=FONT_XS,wraplength=170).pack(pady=8)
        return
    total_w=max(sum(weight_map.values()),1)
    sorted_wm=sorted(weight_map.items(),key=lambda x:-x[1])
    max_w=sorted_wm[0][1] if sorted_wm else 1
    for i,(param,wt) in enumerate(sorted_wm):
        col=CLUSTER_PALETTE[i%len(CLUSTER_PALETTE)]; is_max=(wt==max_w)
        row=tk.Frame(parent,bg=CARD2 if is_max else CARD,padx=5,pady=3,
                     highlightbackground=col if is_max else BORDER,highlightthickness=1)
        row.pack(fill="x",padx=5,pady=2)
        tk.Canvas(row,bg=col,width=10,height=10,
                  highlightthickness=0).pack(side="left",padx=(0,5))
        txt=tk.Frame(row,bg=CARD2 if is_max else CARD)
        txt.pack(side="left",fill="x",expand=True)
        display_name = ("★ " if is_max else "") + param[:20]
        tk.Label(txt,text=display_name,
                 bg=CARD2 if is_max else CARD,fg=col,
                 font=("Courier New",8,"bold"),anchor="w",
                 wraplength=130).pack(anchor="w")
        tk.Label(txt,text=f"{wt:.1f}%",
                 bg=CARD2 if is_max else CARD,
                 fg=FG_MID,font=("Courier New",7),anchor="w").pack(anchor="w")
        _mini_bar(parent,"",wt/total_w,col)
    _divider(parent)
    tk.Label(parent,text="User-assigned weights\nfrom dataset columns",
             bg=CARD,fg=FG_DIM,font=("Courier New",7),wraplength=170,
             justify="left").pack(anchor="w",padx=6,pady=(2,4))


def build_event_legend(parent, event_result):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="EVENT LEGEND",bg=CARD,fg=COL_EVENTS,
             font=FONT_H3,justify="center").pack(pady=(8,2),padx=6)
    _divider(parent,COL_EVENTS)
    if not event_result["has_events"]:
        tk.Label(parent,text="No event column\nin dataset",bg=CARD,
                 fg=WARN,font=FONT_SM,justify="center",
                 wraplength=170).pack(pady=8)
        return
    counts=event_result["event_counts"]; total=max(sum(counts.values()),1)
    def _c(label):
        ll=label.lower()
        for k,c in EVENT_COLOURS.items():
            if k in ll: return c
        return CLUSTER_PALETTE[hash(label)%len(CLUSTER_PALETTE)]
    for label,cnt in sorted(counts.items(),key=lambda x:-x[1]):
        col=_c(label)
        row=tk.Frame(parent,bg=CARD2,padx=5,pady=3,
                     highlightbackground=BORDER,highlightthickness=1)
        row.pack(fill="x",padx=5,pady=2)
        tk.Canvas(row,bg=col,width=10,height=10,
                  highlightthickness=0).pack(side="left",padx=(0,5))
        txt=tk.Frame(row,bg=CARD2); txt.pack(side="left",fill="x",expand=True)
        tk.Label(txt,text=label[:22],bg=CARD2,fg=col,
                 font=("Courier New",8,"bold"),anchor="w",
                 wraplength=130).pack(anchor="w")
        tk.Label(txt,text=f"{cnt:,}  ({cnt/total*100:.1f}%)",bg=CARD2,fg=FG_MID,
                 font=("Courier New",7),anchor="w").pack(anchor="w")
        _mini_bar(parent,"",cnt/total,col)
    _divider(parent)
    tk.Label(parent,text=f"Source: '{event_result['event_col']}'",
             bg=CARD,fg=FG_DIM,font=("Courier New",7),
             wraplength=170,justify="left").pack(anchor="w",padx=6,pady=(2,4))


def build_rpm_legend(parent, rpm_result):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="RPM LEGEND",bg=CARD,fg="#f97316",
             font=FONT_H3,justify="center").pack(pady=(8,2),padx=6)
    _divider(parent,"#f97316")
    if not rpm_result["found"]:
        tk.Label(parent,text="No RPM column\nin dataset",bg=CARD,
                 fg=WARN,font=FONT_SM,justify="center",
                 wraplength=170).pack(pady=8)
        return
    total = max(rpm_result["running"]+rpm_result["stopped"]+
                rpm_result["reverse"]+rpm_result["no_data"], 1)
    items = [
        ("Running (RPM>0)", rpm_result["running"], "#10b981"),
        ("Stopped (RPM=0)", rpm_result["stopped"], "#f59e0b"),
        ("Reverse (RPM<0)", rpm_result["reverse"], "#ef4444"),
        ("No Data",         rpm_result["no_data"], "#4a6080"),
    ]
    for label, cnt, col in items:
        if cnt == 0: continue
        row=tk.Frame(parent,bg=CARD2,padx=5,pady=3,
                     highlightbackground=col,highlightthickness=1)
        row.pack(fill="x",padx=5,pady=2)
        tk.Canvas(row,bg=col,width=10,height=10,
                  highlightthickness=0).pack(side="left",padx=(0,5))
        txt=tk.Frame(row,bg=CARD2); txt.pack(side="left",fill="x",expand=True)
        tk.Label(txt,text=label,bg=CARD2,fg=col,
                 font=("Courier New",8,"bold"),anchor="w",
                 wraplength=130).pack(anchor="w")
        tk.Label(txt,text=f"{cnt:,}  ({cnt/total*100:.1f}%)",bg=CARD2,fg=FG_MID,
                 font=("Courier New",7),anchor="w").pack(anchor="w")
        _mini_bar(parent,"",cnt/total,col)
    _divider(parent)
    tk.Label(parent,text=f"Source: '{rpm_result['col']}'",
             bg=CARD,fg=FG_DIM,font=("Courier New",7),
             wraplength=170,justify="left").pack(anchor="w",padx=6,pady=(2,4))


def _fill_toolbar(bar, mpl_canvas, tab_name):
    tk.Label(bar,text=" TOOLS:",bg="#080e1a",fg=FG_DIM,
             font=FONT_XS).pack(side="left",padx=(6,2))
    nav_frame=tk.Frame(bar,bg="#080e1a"); nav_frame.pack(side="left",padx=2)
    tb=NavigationToolbar2Tk(mpl_canvas,nav_frame); tb.config(bg="#080e1a")
    for child in tb.winfo_children():
        try:
            child.config(bg="#080e1a",fg=FG_MID,activebackground=CARD2,
                         activeforeground=FG,relief="flat",bd=0,
                         highlightthickness=0,font=FONT_XS)
        except Exception: pass
    tb.update()
    tk.Frame(bar,bg=BORDER,width=1).pack(side="left",fill="y",padx=8,pady=3)
    def _save_png():
        fig=active_figures.get(tab_name)
        if not fig: messagebox.showwarning("Export","Run analysis first."); return
        path=filedialog.asksaveasfilename(defaultextension=".png",
                                          initialfile=f"cbm_{tab_name}_{ts()}.png",
                                          filetypes=[("PNG","*.png"),("SVG","*.svg"),("PDF","*.pdf")])
        if not path: return
        fig.savefig(path,dpi=180,bbox_inches="tight",facecolor=fig.get_facecolor())
        set_status(f"Saved: {os.path.basename(path)}",SUCCESS)
    _icon_btn(bar,"Save Chart",_save_png,bg="#1e3a8a",fg="#93c5fd",tip="Save chart as PNG")


def draw_plot(tab_frame, fig, tab_name, legend_builder=None, legend_kwargs=None):
    for w in tab_frame.winfo_children(): w.destroy()
    active_figures[tab_name]=fig
    try:
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            fig.tight_layout(pad=2.5,rect=[0.03,0.03,0.97,0.95])
    except Exception: pass
    fig.patch.set_facecolor("#0d1117")
    bar=tk.Frame(tab_frame,bg="#080e1a",pady=4,
                 highlightbackground=BORDER,highlightthickness=1)
    bar.pack(side="top",fill="x")
    row=tk.Frame(tab_frame,bg=BG_MID); row.pack(side="top",fill="both",expand=True)
    chart_frame=tk.Frame(row,bg=BG_MID); chart_frame.pack(side="left",fill="both",expand=True)
    leg_frame=tk.Frame(row,bg=CARD,width=195,
                       highlightbackground=BORDER,highlightthickness=1)
    leg_frame.pack(side="right",fill="y",padx=(2,6),pady=6)
    leg_frame.pack_propagate(False)
    mpl_canvas=FigureCanvasTkAgg(fig,master=chart_frame); mpl_canvas.draw()
    cw=mpl_canvas.get_tk_widget(); cw.config(bg="#0d1117",highlightthickness=0)
    cw.pack(fill="both",expand=True,padx=2,pady=2)
    _fill_toolbar(bar,mpl_canvas,tab_name)
    if legend_builder: legend_builder(leg_frame,**(legend_kwargs or {}))



# ══════════════════════════════════════════════════════════════════════════════
# PIPELINE
# ══════════════════════════════════════════════════════════════════════════════
def get_selected_x():
    # Feature selection removed — always use all numeric columns
    return []


def resolve_features():
    if raw_data is None: return [],None,0,0,"No dataset loaded"
    all_num=[c for c in raw_data.select_dtypes(include="number").columns]
    x_sel=[c for c in get_selected_x() if c in raw_data.columns]
    x_cols=x_sel or all_num
    needed=list(dict.fromkeys(x_cols))
    wdf=raw_data.dropna(subset=needed).reset_index(drop=True)
    warn="No X selected — all numeric cols used" if not x_sel else ""
    id_col = next((c for c in wdf.columns
                   if any(k in c.lower() for k in ["wellid","well_id","wellname","well_name","uwi","api"])
                   or c.lower() in ["id","well"]), None)
    n_unique = wdf[id_col].nunique() if id_col else len(wdf)
    return x_cols, wdf, len(wdf), n_unique, warn


def start_pipeline():
    if raw_data is None:
        messagebox.showwarning("No Data","Please upload a dataset first."); return
    set_status("Running analysis…",WARN)
    progress_start()
    threading.Thread(target=run_pipeline,daemon=True).start()


def run_pipeline():
    global active_df,active_X,active_xcols
    global active_anomaly_result,active_weight_result,active_event_result
    global active_rpm_result,active_param_anom

    x_cols,wdf,well_count,n_unique_wells,warn=resolve_features()
    if wdf is None or well_count==0:
        app.after(0,lambda: set_status(f"No usable rows: {warn}",WARN))
        app.after(0,progress_stop); return

    try:
        weight_map,raw_total=get_manual_weights()
        method_desc="User-assigned weights (from dataset numeric columns)"
        if not weight_map:
            weight_map={c:round(100/len(x_cols),2) for c in x_cols}
            method_desc="Equal weights (all spinboxes zero — set weights in sidebar)"
        final_weights={}
        for c in x_cols: final_weights[c]=weight_map.get(c,1.0)
        total_fw=sum(final_weights.values())
        final_weights={c:v/total_fw*100 for c,v in final_weights.items()}

        X_w,use_cols=build_weighted_X(wdf,x_cols,final_weights)
        n_clusters=min(cluster_var.get(),well_count)
        wdf=wdf.copy()
        cluster_labels,inertia=run_clustering(X_w,n_clusters)
        wdf["cluster"]=cluster_labels

        # ── RPM Status column ──────────────────────────────────────────────────
        wdf, rpm_col_used, rpm_result = add_rpm_status_column(wdf)
        if rpm_col_used:
            app.after(0, lambda rc=rpm_col_used: rpm_filter_info_var.set(
                f"RPM col: '{rc}'  |  "
                f"Running:{rpm_result['running']}  "
                f"Stopped:{rpm_result['stopped']}  "
                f"Reverse:{rpm_result['reverse']}"
            ))

        # ── Apply RPM filter for anomaly detection ────────────────────────────
        if rpm_col_used and "RPM_Status" in wdf.columns:
            allowed = []
            if rpm_filter_all_var.get():
                allowed = ["Running","Stopped","Reverse","No Data"]
            else:
                if rpm_filter_running_var.get(): allowed.append("Running")
                if rpm_filter_stopped_var.get(): allowed.append("Stopped")
                if rpm_filter_reverse_var.get(): allowed.append("Reverse")
                if rpm_filter_nodata_var.get():  allowed.append("No Data")
            if allowed and len(allowed) < 4:
                rpm_mask   = wdf["RPM_Status"].isin(allowed)
                X_w_anom   = X_w[rpm_mask.values]
                anom_filter_desc = f"RPM filter: {', '.join(allowed)}"
            else:
                X_w_anom         = X_w
                anom_filter_desc = "No RPM filter (all states)"
        else:
            X_w_anom         = X_w
            anom_filter_desc = "No RPM column detected"

        try:
            from model.test import test_model
            wdf=test_model(wdf)
        except Exception: pass

        method=anomaly_method_var.get(); contam=anomaly_contam_var.get()/100.0
        a_labels_sub,a_pct,a_idx_sub,a_name,a_scores_sub=\
            detect_anomalies(X_w_anom,contamination=contam,method=method)

        # Map filtered anomaly labels back onto full dataframe
        a_labels = np.ones(len(wdf), dtype=int)
        a_scores = np.zeros(len(wdf))
        if rpm_col_used and "RPM_Status" in wdf.columns and len(X_w_anom) < len(X_w):
            rpm_mask   = wdf["RPM_Status"].isin(allowed)
            sub_indices = np.where(rpm_mask.values)[0]
            for local_i, global_i in enumerate(sub_indices):
                a_labels[global_i] = a_labels_sub[local_i]
                a_scores[global_i] = a_scores_sub[local_i]
            a_idx = [sub_indices[i] for i in a_idx_sub]
        else:
            a_labels = a_labels_sub
            a_scores = a_scores_sub
            a_idx    = a_idx_sub

        a_name = f"{a_name}  [{anom_filter_desc}]"
        wdf["anomaly"]=a_labels
        anomaly_result={
            "labels":a_labels,"pct":a_pct,"indices":a_idx,
            "scores":a_scores,"detector_name":a_name,
            "n_anomaly":len(a_idx),"n_normal":len(a_labels)-len(a_idx),
        }
        event_result=analyse_events(wdf,a_labels)

        # ── ★ Per-parameter anomaly analysis ───────────────────────────────────
        param_anom = analyse_per_parameter(wdf, use_cols, z_threshold=2.5)

        active_df=wdf; active_X=X_w; active_xcols=use_cols
        active_anomaly_result=anomaly_result
        active_weight_result={"weight_map":final_weights,"method_desc":method_desc}
        active_event_result=event_result
        active_rpm_result=rpm_result
        active_param_anom=param_anom

        hx=cluster_hex(n_clusters)
        y_cols=[c for c in wdf.select_dtypes(include="number").columns
                if c not in use_cols and c!="cluster"][:2]
        insights=generate_insights(wdf,use_cols,well_count,n_unique_wells,hx,
                                   anomaly_result,event_result,
                                   final_weights,method_desc,inertia,rpm_result,
                                   param_anom)

        app.after(0,lambda: refresh_ui(
            wdf,X_w,use_cols,y_cols,hx,insights,well_count,n_unique_wells,
            anomaly_result,event_result,final_weights,method_desc,rpm_result,
            param_anom))
        app.after(0,lambda: set_status("Analysis complete",SUCCESS))
        app.after(0,progress_stop)
        app.after(0,lambda: export_btn.config(state="normal"))

    except Exception as e:
        traceback.print_exc()
        app.after(0,lambda err=str(e): set_status(f"Error: {err}",ERR))
        app.after(0,progress_stop)


def refresh_ui(df,X,x_cols,y_cols,hx,insights,well_count,n_unique_wells,
               anomaly_result,event_result,weight_map,method_desc,rpm_result,
               param_anom):
    draw_plot(cluster_tab,   plot_clusters(df,x_cols,y_cols,hx),      "clusters",
              legend_builder=build_cluster_legend,legend_kwargs={"df":df,"hex_colors":hx})
    draw_plot(pca_tab,       plot_pca(X,df["cluster"].values,hx),      "pca",
              legend_builder=build_cluster_legend,legend_kwargs={"df":df,"hex_colors":hx})
    draw_plot(reservoir_tab, plot_reservoir_3d(df,x_cols,y_cols,hx),  "reservoir_3d",
              legend_builder=build_cluster_legend,legend_kwargs={"df":df,"hex_colors":hx})
    draw_plot(production_tab,plot_production(df,y_cols,x_cols,hx),    "production",
              legend_builder=build_cluster_legend,legend_kwargs={"df":df,"hex_colors":hx})
    draw_plot(hidden_tab,
              plot_hidden_patterns(X,anomaly_result["labels"],x_cols,anomaly_result["detector_name"]),
              "hidden_patterns",
              legend_builder=build_anomaly_legend,
              legend_kwargs={"n_normal":anomaly_result["n_normal"],
                             "n_anomaly":anomaly_result["n_anomaly"]})
    draw_plot(weight_tab,    plot_weight_chart(weight_map,method_desc),"param_weights",
              legend_builder=build_weight_legend,
              legend_kwargs={"weight_map":weight_map,"method_desc":method_desc})
    draw_plot(event_tab,     plot_event_chart(df,event_result),        "event_summary",
              legend_builder=build_event_legend,legend_kwargs={"event_result":event_result})
    draw_plot(rpm_tab,       plot_rpm_chart(df,rpm_result),            "rpm_status",
              legend_builder=build_rpm_legend,legend_kwargs={"rpm_result":rpm_result})

    # ★ NEW — per-parameter anomaly tab (passes actual df + labels for box plots)
    draw_plot(param_anom_tab,
              plot_param_anomaly(param_anom,
                                 df=df,
                                 anomaly_labels=anomaly_result["labels"]),
              "param_anomaly",
              legend_builder=build_param_anomaly_legend,
              legend_kwargs={"param_anom_list": param_anom})

    # ── Stat tiles — use correct labels (wells vs records) ────────────────────
    # "Unique Wells" = distinct well IDs (e.g. 287)
    stat_widgets["Unique Wells"].set(f"{n_unique_wells:,}")
    # "Total Records" = total rows in dataset (e.g. 346,769)
    stat_widgets["Total Records"].set(f"{well_count:,}")
    stat_widgets["Clusters"].set(str(df["cluster"].nunique()))
    anom_n = anomaly_result['n_anomaly']
    anom_p = anomaly_result['pct']
    stat_widgets["Anomaly Recs"].set(f"{anom_n:,}\n({anom_p:.1f}%)")
    if event_result["has_events"]:
        stat_widgets["Active Recs"].set(f"{event_result['n_active']:,}")
        stat_widgets["Inactive Recs"].set(f"{event_result['n_inactive']:,}")
        stat_widgets["Event Types"].set(str(len(event_result["event_counts"])))
    else:
        stat_widgets["Active Recs"].set("N/A")
        stat_widgets["Inactive Recs"].set("N/A")
        stat_widgets["Event Types"].set("N/A")
    stat_widgets["Abnormal Recs"].set(f"{event_result['n_abnormal']:,}")
    if rpm_result["found"]:
        stat_widgets["RPM Running"].set(f"{rpm_result['running']:,}")
        stat_widgets["RPM Stopped"].set(f"{rpm_result['stopped']:,}")
        stat_widgets["RPM Reverse"].set(f"{rpm_result['reverse']:,}")
    else:
        stat_widgets["RPM Running"].set("N/A")
        stat_widgets["RPM Stopped"].set("N/A")
        stat_widgets["RPM Reverse"].set("N/A")

    # Param anomaly stat
    flagged = sum(1 for r in param_anom if r["n_anom"] > 0)
    stat_widgets["Param Anomalies"].set(f"{flagged}/{len(param_anom)}\nparams flagged")

    update_event_panel(event_result)
    explain_text.config(state="normal")
    explain_text.delete("1.0","end")
    explain_text.insert("end",insights)
    explain_text.config(state="disabled")


# ══════════════════════════════════════════════════════════════════════════════
# EVENT PANEL UPDATE
# ══════════════════════════════════════════════════════════════════════════════
def update_event_panel(event_result):
    if not event_result["has_events"]:
        ev_status_var.set("No event column in dataset"); ev_status_lbl.config(fg=WARN)
        for v in ev_count_vars.values(): v.set("N/A")
        ev_active_var.set("N/A"); ev_inactive_var.set("N/A")
        ev_abnormal_var.set("N/A"); ev_op_anom_var.set("N/A"); return
    ev_status_var.set(f"Column: '{event_result['event_col']}'")
    ev_status_lbl.config(fg=SUCCESS2)
    counts=event_result["event_counts"]
    top=sorted(counts.items(),key=lambda x:-x[1])[:6]
    keys=list(ev_count_vars.keys())
    for i,k in enumerate(keys):
        if i<len(top):
            label,cnt=top[i]; ev_count_vars[k].set(f"{cnt:,}")
            ev_count_lbls[k].config(text=label[:24])
        else:
            ev_count_vars[k].set("—"); ev_count_lbls[k].config(text="—")
    ev_active_var.set(f"{event_result['n_active']:,}")
    ev_inactive_var.set(f"{event_result['n_inactive']:,}")
    ev_abnormal_var.set(f"{event_result['n_abnormal']:,}")
    ev_op_anom_var.set(f"{len(event_result.get('operational_anomalies',[])):,}")


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════
def export_results():
    if active_df is None:
        messagebox.showwarning("Export","Run analysis first."); return
    path=filedialog.asksaveasfilename(
        defaultextension=".xlsx",initialfile=f"cbm_results_{ts()}.xlsx",
        filetypes=[("Excel","*.xlsx"),("CSV","*.csv")])
    if not path: return
    try:
        import pandas as pd
        ext=os.path.splitext(path)[1].lower()
        if ext==".xlsx":
            with pd.ExcelWriter(path,engine="openpyxl") as writer:
                active_df.to_excel(writer,sheet_name="All Wells",index=False)
                anom_df=active_df[active_df["anomaly"]==-1] \
                    if "anomaly" in active_df.columns else active_df.iloc[0:0]
                anom_df.to_excel(writer,sheet_name="Anomaly Wells",index=False)
                cl_sum=active_df.groupby("cluster").agg(
                    well_count=("cluster","count")).reset_index()
                cl_sum.to_excel(writer,sheet_name="Cluster Summary",index=False)
                if active_weight_result:
                    wm=active_weight_result["weight_map"]
                    wdf_exp=pd.DataFrame([{"Parameter":k,"Weight_%":v}
                                          for k,v in sorted(wm.items(),key=lambda x:-x[1])])
                    wdf_exp["Method"]=active_weight_result["method_desc"]
                    wdf_exp.to_excel(writer,sheet_name="Parameter Weights",index=False)
                # ★ NEW — per-parameter anomaly sheet
                if active_param_anom:
                    pa_rows=[]
                    for r in active_param_anom:
                        pa_rows.append({
                            "Parameter":  r["col"],
                            "Total_Rows": r["n_total"],
                            "Outliers":   r["n_anom"],
                            "Outlier_%":  round(r["pct"],2),
                            "High_Outliers": r["high_out"],
                            "Low_Outliers":  r["low_out"],
                            "Mean":   round(r["mean"],4),
                            "StdDev": round(r["std"],4),
                            "Median": round(r["median"],4),
                            "Skewness":  round(r["skew"],4),
                            "Kurtosis":  round(r["kurt"],4),
                            "Severity":  r["severity"],
                            "Reason / Analysis": r["reason"],
                        })
                    pd.DataFrame(pa_rows).to_excel(
                        writer,sheet_name="Parameter Anomalies",index=False)
                report_txt=explain_text.get("1.0","end").strip()
                pd.DataFrame({"Report":report_txt.split("\n")}).to_excel(
                    writer,sheet_name="Insights Report",index=False)
            set_status(f"Exported (6 sheets): {os.path.basename(path)}",SUCCESS)
            messagebox.showinfo("Export Complete",
                f"Saved: {path}\n\nSheets:\n"
                "  All Wells\n  Anomaly Wells\n  Cluster Summary\n"
                "  Parameter Weights\n  Parameter Anomalies\n  Insights Report")
        else:
            _save_df(active_df,path)
            set_status(f"Exported: {os.path.basename(path)}",SUCCESS)
    except Exception as e:
        messagebox.showerror("Export Error",str(e))
        set_status(f"Export failed: {e}",ERR)


# ══════════════════════════════════════════════════════════════════════════════
# ★ ENHANCED INSIGHTS — now includes per-parameter anomaly section
# ══════════════════════════════════════════════════════════════════════════════
def generate_insights(df, x_cols, well_count, n_unique_wells, hx, anomaly_result,
                      event_result, weight_map, method_desc, inertia,
                      rpm_result=None, param_anom=None):
    import textwrap

    clusters  = df["cluster"].value_counts().sort_index()
    n_anom    = anomaly_result["n_anomaly"]
    a_pct     = anomaly_result["pct"]
    sev_label = ("LOW"      if a_pct <  2 else
                 "MODERATE" if a_pct <  8 else
                 "ELEVATED" if a_pct < 20 else "HIGH")
    sev_interp = {
        "LOW":      "Very few outliers. Reservoir behaviour is relatively uniform.",
        "MODERATE": "Moderate anomalies present. Check for localised heterogeneity.",
        "ELEVATED": "Significant complexity detected. Variable seam or fracture network likely.",
        "HIGH":     "Large anomaly fraction. Review contamination % or check data quality.",
    }[sev_label]

    W   = "=" * 50
    D   = "-" * 50
    B   = ""   # blank line separator

    lines = [
        W,
        "  CBM AI ANALYSIS REPORT   v5.2",
        f"  Generated : {datetime.datetime.now().strftime('%Y-%m-%d   %H:%M:%S')}",
        W, B,
        "  DATASET SUMMARY",
        "  " + D,
        f"  Unique wells   : {n_unique_wells:,}  (distinct well IDs)",
        f"  Total records  : {well_count:,}  (rows in dataset)",
        f"  Analysed rows  : {len(df):,}",
        f"  Clusters       : {len(clusters)}",
        f"  Features used  : {len(x_cols)}",
        f"  Feature list   : {', '.join(x_cols)}",
        f"  KMeans inertia : {inertia:.2f}",
        B,
    ]

    # PARAMETER WEIGHTS
    lines += [
        W,
        "  PARAMETER WEIGHTS  (user-assigned)",
        "  " + D,
        f"  Method : {method_desc}",
        B,
    ]
    if weight_map:
        max_w = max(weight_map.values())
        lines.append(f"  {'PARAMETER':<28}  {'WEIGHT':>7}  INFLUENCE BAR")
        lines.append("  " + D)
        for param, wt in sorted(weight_map.items(), key=lambda x: -x[1]):
            star = "  <- HIGHEST" if wt == max_w else ""
            bar  = "||" * max(1, int(wt / 4))
            lines.append(f"  {param:<28}  {wt:>6.1f}%  {bar}{star}")
    else:
        lines.append("  (equal weights applied to all parameters)")
    lines.append(B)

    # CLUSTER BREAKDOWN
    lines += [
        W,
        "  CLUSTER BREAKDOWN",
        "  " + D,
        f"  {'CLUSTER':<12}  {'COUNT':>6}  {'PCT':>6}  SIZE BAR",
        "  " + D,
    ]
    for cl, cnt in clusters.items():
        pct = cnt / max(len(df), 1) * 100
        bar = "=" * max(1, int(pct / 4))
        lines.append(f"  Cluster {cl:<4}   {cnt:>6,}  {pct:>5.1f}%  {bar}")
    lines.append(B)

    # OVERALL ANOMALY DETECTION
    lines += [
        W,
        "  OVERALL ANOMALY DETECTION",
        "  (all features combined into one model)",
        "  " + D,
        f"  Algorithm      : {anomaly_result['detector_name']}",
        f"  Normal records : {anomaly_result['n_normal']:,}",
        f"  Anomalous recs : {n_anom:,}  ({a_pct:.1f}%)",
        f"  Severity level : {sev_label}",
        B,
        "  Interpretation :",
        f"    {sev_interp}",
        B,
    ]

    # PER-PARAMETER ANOMALY ANALYSIS
    lines += [
        W,
        "  PER-PARAMETER ANOMALY ANALYSIS",
        "  Each column is checked individually using z-score statistics.",
        "  A record is flagged if its value is more than",
        "  +/- z-threshold standard deviations from the column mean.",
        "  " + D,
    ]

    if param_anom:
        flagged = [r for r in param_anom if r["n_anom"] > 0]
        clean   = [r for r in param_anom if r["n_anom"] == 0]
        z_thr   = param_anom[0]["z_threshold"] if param_anom else 2.5

        lines += [
            f"  Z-score threshold  : +/- {z_thr} sigma",
            f"  Parameters checked : {len(param_anom)}",
            f"  Parameters flagged : {len(flagged)}",
            f"  Parameters clean   : {len(clean)}",
            B,
        ]

        if not flagged:
            lines += [
                "  ALL PARAMETERS ARE CLEAN",
                "  No column shows statistically significant outliers",
                "  at the current z-score threshold.",
                B,
            ]
        else:
            lines += [
                "  FLAGGED PARAMETERS  (ranked worst first)",
                "  " + D,
                B,
            ]
            for idx, r in enumerate(flagged, 1):
                upper_lim = r["mean"] + r["z_threshold"] * r["std"]
                lower_lim = r["mean"] - r["z_threshold"] * r["std"]

                if r["high_out"] > 0 and r["low_out"] > 0:
                    direction_str = (
                        f"{r['high_out']} records above upper limit"
                        f"  AND  {r['low_out']} records below lower limit"
                    )
                elif r["high_out"] > 0:
                    direction_str = (
                        f"{r['high_out']} records ABOVE the upper limit"
                        f" of {upper_lim:.4f}  (values are too HIGH)"
                    )
                else:
                    direction_str = (
                        f"{r['low_out']} records BELOW the lower limit"
                        f" of {lower_lim:.4f}  (values are too LOW)"
                    )

                skew_label = (
                    "strongly right-skewed (tail toward high values)" if r["skew"] > 1.5 else
                    "moderately right-skewed"                          if r["skew"] > 0.5 else
                    "strongly left-skewed (tail toward low values)"    if r["skew"] < -1.5 else
                    "moderately left-skewed"                           if r["skew"] < -0.5 else
                    "approximately symmetric"
                )

                # bar removed — not needed in new layout

                lines += [
                    f"  [{idx}] PARAMETER : {r['col']}",
                    "  " + "-" * 46,
                    f"      Severity        : {r['severity']}",
                    f"      Normal records  : {r['n_total'] - r['n_anom']:,}"
                    f"  ({100.0 - r['pct']:.1f}%)  <- NORMAL",
                    f"      Anomalous recs  : {r['n_anom']:,}"
                    f"  ({r['pct']:.2f}%)  <- ANOMALOUS",
                    f"      Direction       : {direction_str}",
                    B,
                    "      STATISTICS:",
                    f"        Mean        = {r['mean']:.4f}",
                    f"        Std Dev     = {r['std']:.4f}",
                    f"        Median      = {r['median']:.4f}",
                    f"        Min value   = {r['min']:.4f}",
                    f"        Max value   = {r['max']:.4f}",
                    f"        Skewness    = {r['skew']:.4f}  ({skew_label})",
                    f"        Kurtosis    = {r['kurt']:.4f}",
                    f"        Upper limit = {upper_lim:.4f}  (mean + {z_thr}*sigma)",
                    f"        Lower limit = {lower_lim:.4f}  (mean - {z_thr}*sigma)",
                    B,
                    "      REASON / ANALYSIS:",
                ]
                # Word-wrap reason text cleanly at 56 chars
                for wrap_line in textwrap.wrap(r["reason"], width=56):
                    lines.append(f"        {wrap_line}")
                lines += [B, "  " + D, B]

        if clean:
            lines += [
                "  CLEAN PARAMETERS  (no outliers at current threshold)",
                "  " + D,
                f"  {'PARAMETER':<28}  {'MEAN':>10}  {'STD DEV':>10}",
                "  " + D,
            ]
            for r in clean:
                lines.append(
                    f"  {r['col']:<28}  {r['mean']:>10.4f}  {r['std']:>10.4f}"
                )
            lines.append(B)
    else:
        lines += ["  Per-parameter analysis not available.", B]

    # OPERATIONAL EVENTS
    lines += [
        W,
        "  OPERATIONAL EVENTS",
        "  " + D,
    ]
    if event_result["has_events"]:
        op_count = len(event_result.get("operational_anomalies", []))
        lines += [
            f"  Source column  : {event_result['event_col']}",
            f"  Active records : {event_result['n_active']:,}",
            f"  Inactive recs  : {event_result['n_inactive']:,}",
            f"  Abnormal       : {event_result['n_abnormal']:,}",
            f"  Op. anomalies  : {op_count:,}",
            B,
            f"  {'EVENT TYPE':<30}  {'COUNT':>6}  {'PCT':>6}",
            "  " + D,
        ]
        for ev, cnt in sorted(event_result["event_counts"].items(),
                               key=lambda x: -x[1]):
            pct = cnt / max(len(df), 1) * 100
            lines.append(f"  {ev:<30}  {cnt:>6,}  {pct:>5.1f}%")
        lines.append(B)
    else:
        lines += [f"  {event_result['message']}", B]

    # RPM STATUS
    lines += [
        W,
        "  RPM STATUS ANALYSIS",
        "  " + D,
    ]
    if rpm_result and rpm_result.get("found"):
        total_rpm = max(
            rpm_result["running"] + rpm_result["stopped"] +
            rpm_result["reverse"] + rpm_result["no_data"], 1
        )
        lines += [
            f"  RPM column  : {rpm_result['col']}",
            B,
            f"  {'STATUS':<20}  {'COUNT':>6}  {'PCT':>6}",
            "  " + D,
            f"  {'Running  (RPM > 0)':<20}  {rpm_result['running']:>6,}"
            f"  {rpm_result['running']/total_rpm*100:>5.1f}%",
            f"  {'Stopped  (RPM = 0)':<20}  {rpm_result['stopped']:>6,}"
            f"  {rpm_result['stopped']/total_rpm*100:>5.1f}%",
            f"  {'Reverse  (RPM < 0)':<20}  {rpm_result['reverse']:>6,}"
            f"  {rpm_result['reverse']/total_rpm*100:>5.1f}%",
            f"  {'No Data':<20}  {rpm_result['no_data']:>6,}",
            B,
        ]
        if rpm_result["reverse"] > 0:
            lines += [
                "  ALERT: Reverse RPM detected.",
                "    Check pump rotation direction or sensor polarity.",
            ]
        if rpm_result["stopped"] / total_rpm > 0.3:
            lines += [
                "  ALERT: Over 30% of records show RPM = 0.",
                "    Significant pump downtime detected.",
                "    Investigate mechanical or operational root cause.",
            ]
        lines.append(B)
    else:
        lines += [
            "  No RPM column found in dataset.",
            "  Add a column with 'rpm' in its name to enable this analysis.",
            B,
        ]

    lines += [W, "  END OF REPORT", W]
    return "\n".join(lines)



# ══════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
if HAVE_TTKBS:
    app = ttk.Window(themename="darkly")
else:
    app = tk.Tk()
    app.configure(bg=BG_DEEP)

app.title("CBM AI Analytics Platform  v5.1  —  Per-Parameter Anomaly Analysis")
app.geometry("1720x980")
app.configure(bg=BG_DEEP)
app.minsize(1280,720)

top_bar=tk.Frame(app,bg="#040810",height=46)
top_bar.pack(fill="x",side="top"); top_bar.pack_propagate(False)
brand=tk.Frame(top_bar,bg="#040810"); brand.pack(side="left",padx=16,fill="y")
tk.Label(brand,text="CBM·AI",bg="#040810",fg=COL_DATASET,
         font=("Georgia",17,"bold")).pack(side="left",padx=(0,10))
tk.Label(brand,text="Coalbed Methane Analytics Platform  v5.1  —  Per-Parameter Anomaly",
         bg="#040810",fg=FG_DIM,font=FONT_SM).pack(side="left")
tk.Label(top_bar,
         text="MS Access 2007+ · Any file size · Dynamic weights · Events · Per-param anomaly",
         bg="#040810",fg=FG_DIM,font=FONT_XS).pack(side="right",padx=14)
tk.Frame(app,bg=BORDER,height=1).pack(fill="x")

root_pane=tk.PanedWindow(app,orient="horizontal",bg=BG_DEEP,
                          sashwidth=4,sashrelief="flat",sashpad=0,handlesize=0)
root_pane.pack(fill="both",expand=True)

# ══════════════════════════════════════════════════════════════════════════════
# LEFT SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
sb_outer,sidebar=make_scrollable(root_pane,bg=SIDEBAR)
root_pane.add(sb_outer,width=430,minsize=390)
tk.Frame(sidebar,bg=SIDEBAR,height=8).pack()

# ① DATASET
ds_body=make_section_card(sidebar,"① DATASET",COL_DATASET)
tk.Label(ds_body,
         text="Accepts: MS Access 2007+ (.mdb/.accdb), CSV, TSV,\n"
              "XLSX, XLS, XLSB, XLSM, ODS, JSON, Parquet,\n"
              "Feather, HDF5, Pickle — any file size.",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=390,justify="left").pack(anchor="w",pady=(0,4))
make_btn(ds_body,"  Upload Data File (any format)",upload_dataset,
         color="#0e7490",fg_col="#e0f7fa",
         tip="Load any data file — MS Access .mdb/.accdb, CSV, Excel, etc.")
info_row=tk.Frame(ds_body,bg=CARD); info_row.pack(fill="x",pady=(6,0))
rows_var=tk.StringVar(value="—"); cols_var=tk.StringVar(value="—")
num_cols_var=tk.StringVar(value="—")
for title,var,col in [("Unique Wells",rows_var,COL_DATASET),
                      ("Columns",cols_var,FG_DIM),
                      ("Numeric",num_cols_var,COL_WEIGHTS)]:
    cf=tk.Frame(info_row,bg=CARD); cf.pack(side="left",expand=True,fill="x")
    tk.Label(cf,text=title,bg=CARD,fg=FG_DIM,font=FONT_XS).pack(anchor="w")
    tk.Label(cf,textvariable=var,bg=CARD,fg=col,
             font=("Courier New",13,"bold")).pack(anchor="w")
event_col_var=tk.StringVar(value="Upload dataset to detect event column")
event_col_lbl=tk.Label(ds_body,textvariable=event_col_var,bg=CARD,fg=FG_DIM,
                       font=FONT_XS,wraplength=390,justify="left",anchor="w")
event_col_lbl.pack(fill="x",pady=(6,0))

# ② CLUSTER SETTINGS
cl_body=make_section_card(sidebar,"② CLUSTER SETTINGS",COL_CLUSTER)
tk.Label(cl_body,text="Number of Clusters  (2 – 10)",bg=CARD,fg=FG_DIM,font=FONT_SM).pack(anchor="w")
cluster_var=tk.IntVar(value=3)
sf=tk.Frame(cl_body,bg=CARD); sf.pack(fill="x",pady=4)
tk_ttk.Spinbox(sf,from_=2,to=10,textvariable=cluster_var,width=5,font=FONT_H2).pack(side="left")
tk.Label(sf,text="clusters",bg=CARD,fg=FG_DIM,font=FONT_SM).pack(side="left",padx=8)

# ③ PARAMETER WEIGHTAGE
wt_outer=make_section_card(sidebar,"③ PARAMETER WEIGHTAGE  (from your dataset)",COL_WEIGHTS)
tk.Label(wt_outer,
         text="After uploading, each numeric column appears here.\n"
              "Assign a % weight to each.  Total should = 100%.",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=390,justify="left").pack(anchor="w",pady=(0,4))
weight_sum_var=tk.StringVar(value="Upload a dataset to see parameters")
weight_sum_lbl=tk.Label(wt_outer,textvariable=weight_sum_var,bg=CARD,fg=FG_DIM,
                        font=("Courier New",9,"bold"),anchor="w")
weight_sum_lbl.pack(anchor="w",pady=(0,4))
wt_canvas_outer=tk.Frame(wt_outer,bg=CARD,height=180)
wt_canvas_outer.pack(fill="x"); wt_canvas_outer.pack_propagate(False)
wt_canvas=Canvas(wt_canvas_outer,bg=CARD,highlightthickness=0,bd=0)
wt_sb=Scrollbar(wt_canvas_outer,orient="vertical",command=wt_canvas.yview,
                bg=BORDER,troughcolor=CARD,activebackground=COL_WEIGHTS)
wt_canvas.configure(yscrollcommand=wt_sb.set)
wt_sb.pack(side="right",fill="y"); wt_canvas.pack(side="left",fill="both",expand=True)
weight_body=tk.Frame(wt_canvas,bg=CARD)
wt_wid=wt_canvas.create_window((0,0),window=weight_body,anchor="nw")
def _wt_resize(e):
    wt_canvas.configure(scrollregion=wt_canvas.bbox("all"))
    wt_canvas.itemconfig(wt_wid,width=e.width)
weight_body.bind("<Configure>",lambda e: wt_canvas.configure(scrollregion=wt_canvas.bbox("all")))
wt_canvas.bind("<Configure>",_wt_resize)
_wt_placeholder=tk.Label(weight_body,
                          text="Upload a dataset first.\nNumeric columns will appear here automatically.",
                          bg=CARD,fg=FG_DIM,font=FONT_XS,justify="left",padx=8,pady=8)
_wt_placeholder.pack(anchor="w")
weight_row_frames=[_wt_placeholder]

# ④ ANOMALY DETECTION
anom_body=make_section_card(sidebar,"④ ANOMALY DETECTION",COL_ANOMALY)
tk.Label(anom_body,text="Algorithm",bg=CARD,fg=FG_DIM,font=FONT_SM).pack(anchor="w")
anomaly_method_var=tk.StringVar(value="iforest")
mrow=tk.Frame(anom_body,bg=CARD); mrow.pack(fill="x",pady=(2,6))
for label,val in [("Isolation Forest  (faster, tree-based)","iforest"),
                   ("Local Outlier Factor  (density-based)","lof")]:
    tk.Radiobutton(mrow,text=label,variable=anomaly_method_var,value=val,
                   bg=CARD,fg=FG_MID,selectcolor=CARD2,activebackground=CARD,
                   activeforeground=COL_ANOMALY,font=FONT_SM,
                   wraplength=390).pack(anchor="w",pady=1)
tk.Label(anom_body,text="Contamination  (expected anomaly %)",
         bg=CARD,fg=FG_DIM,font=FONT_SM).pack(anchor="w")
anomaly_contam_var=tk.DoubleVar(value=5.0)
crow=tk.Frame(anom_body,bg=CARD); crow.pack(fill="x",pady=2)
tk_ttk.Spinbox(crow,from_=1.0,to=49.0,increment=0.5,
               textvariable=anomaly_contam_var,width=6,font=FONT_H3).pack(side="left")
tk.Label(crow,text="%   (typical CBM: 3–10%)",bg=CARD,fg=FG_DIM,font=FONT_XS).pack(side="left",padx=6)

# ★ NEW ④b — Per-Parameter Z-Score Threshold
pa_body=make_section_card(sidebar,"④b PER-PARAMETER  Z-Score Threshold",COL_PARAM_ANOM)
tk.Label(pa_body,
         text="Controls individual column outlier sensitivity.\n"
              "Lower = more anomalies detected per parameter.\n"
              "Recommended: 2.0–3.0  (default 2.5 = standard statistical)",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=390,justify="left").pack(anchor="w",pady=(0,6))
zthresh_row=tk.Frame(pa_body,bg=CARD); zthresh_row.pack(fill="x",pady=2)
# Note: z_threshold is hardcoded to 2.5 in run_pipeline; wire up a var here
zthresh_var=tk.DoubleVar(value=2.5)
tk_ttk.Spinbox(zthresh_row,from_=1.0,to=5.0,increment=0.1,
               textvariable=zthresh_var,width=6,font=FONT_H3).pack(side="left")
tk.Label(zthresh_row,text="σ  (standard deviations from mean)",
         bg=CARD,fg=FG_DIM,font=FONT_XS).pack(side="left",padx=6)

# ⑤ RPM FILTER
rpm_filter_body=make_section_card(sidebar,"⑤ RPM FILTER  (for anomaly analysis)","#f97316")
tk.Label(rpm_filter_body,
         text="Select which RPM states to include\nwhen running anomaly detection.",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=390,justify="left").pack(anchor="w",pady=(0,6))
rpm_filter_all_var      = tk.BooleanVar(value=True)
rpm_filter_running_var  = tk.BooleanVar(value=True)
rpm_filter_stopped_var  = tk.BooleanVar(value=True)
rpm_filter_reverse_var  = tk.BooleanVar(value=True)
rpm_filter_nodata_var   = tk.BooleanVar(value=True)
def _rpm_all_toggle():
    if rpm_filter_all_var.get():
        rpm_filter_running_var.set(True); rpm_filter_stopped_var.set(True)
        rpm_filter_reverse_var.set(True); rpm_filter_nodata_var.set(True)
rpm_filter_col = "#f97316"
tk.Checkbutton(rpm_filter_body,text="All RPM states (no filter)",
               variable=rpm_filter_all_var, command=_rpm_all_toggle,
               bg=CARD,fg=FG_MID,selectcolor=CARD2,activebackground=CARD,
               activeforeground=rpm_filter_col,font=FONT_SM).pack(anchor="w",pady=1)
_divider(rpm_filter_body)
for label, var, col in [
    ("Running  (RPM > 0)", rpm_filter_running_var, "#10b981"),
    ("Stopped  (RPM = 0)", rpm_filter_stopped_var, "#f59e0b"),
    ("Reverse  (RPM < 0)", rpm_filter_reverse_var, "#ef4444"),
    ("No Data  (NaN/bad)", rpm_filter_nodata_var,  "#4a6080"),
]:
    tk.Checkbutton(rpm_filter_body, text=label, variable=var,
                   bg=CARD, fg=col, selectcolor=CARD2, activebackground=CARD,
                   activeforeground=col, font=FONT_SM).pack(anchor="w", pady=1)
rpm_filter_info_var = tk.StringVar(value="No RPM column detected yet")
tk.Label(rpm_filter_body, textvariable=rpm_filter_info_var,
         bg=CARD, fg=FG_DIM, font=FONT_XS,
         wraplength=390, justify="left").pack(anchor="w", pady=(6,0))




# ⑦ RUN
run_body=make_section_card(sidebar,"⑦ RUN ANALYSIS","#6366f1")
make_btn(run_body,"  Run AI Analysis",start_pipeline,
         color="#4f46e5",tip="Cluster + anomaly + per-parameter + event analysis")
if HAVE_TTKBS:
    progress_bar=ttk.Progressbar(run_body,mode="determinate",bootstyle="info-striped",length=390)
else:
    progress_bar=tk_ttk.Progressbar(run_body,mode="determinate",length=390)
progress_bar.pack(fill="x",pady=(4,2))
status_var=tk.StringVar(value="Ready — upload a dataset to begin")
status_lbl=tk.Label(run_body,textvariable=status_var,bg=CARD,fg=FG_DIM,
                    font=FONT_SM,anchor="w",wraplength=390,justify="left")
status_lbl.pack(fill="x",pady=(2,0))

# ⑧ EVENT DETECTION PANEL
ev_body=make_section_card(sidebar,"⑧ EVENT DETECTION PANEL",COL_EVENTS)
ev_status_var=tk.StringVar(value="Awaiting analysis")
ev_status_lbl=tk.Label(ev_body,textvariable=ev_status_var,bg=CARD,fg=FG_DIM,
                       font=FONT_SM,wraplength=390,justify="left")
ev_status_lbl.pack(anchor="w",pady=(0,6))
ev_tile_row=tk.Frame(ev_body,bg=CARD); ev_tile_row.pack(fill="x",pady=(0,6))
ev_active_var=tk.StringVar(value="—"); ev_inactive_var=tk.StringVar(value="—")
ev_abnormal_var=tk.StringVar(value="—"); ev_op_anom_var=tk.StringVar(value="—")
for title,var,col in [("Active Recs",ev_active_var,COL_EVENTS),
                       ("Inactive Recs",ev_inactive_var,COL_ANOMALY),
                       ("Abnormal",ev_abnormal_var,WARN),
                       ("Op. Anom.",ev_op_anom_var,COL_WEIGHTS)]:
    cell=tk.Frame(ev_tile_row,bg=CARD2,padx=5,pady=5,
                  highlightbackground=BORDER,highlightthickness=1)
    cell.pack(side="left",expand=True,fill="x",padx=2)
    tk.Label(cell,textvariable=var,bg=CARD2,fg=col,
             font=("Courier New",11,"bold")).pack()
    tk.Label(cell,text=title,bg=CARD2,fg=FG_DIM,font=FONT_XS).pack()
_ev_label_names=[f"ev{i}" for i in range(6)]
ev_count_vars={}; ev_count_lbls={}
ev_detail_frame=tk.Frame(ev_body,bg=CARD); ev_detail_frame.pack(fill="x")
for key in _ev_label_names:
    row=tk.Frame(ev_detail_frame,bg=CARD); row.pack(fill="x",pady=1)
    lbl=tk.Label(row,text="—",bg=CARD,fg=FG_DIM,font=FONT_XS,width=26,anchor="w")
    lbl.pack(side="left",padx=(4,0))
    cnt_var=tk.StringVar(value="—")
    tk.Label(row,textvariable=cnt_var,bg=CARD,fg=COL_EVENTS,
             font=FONT_XS,width=8,anchor="e").pack(side="right",padx=4)
    ev_count_vars[key]=cnt_var; ev_count_lbls[key]=lbl

# ⑨ EXPORT
exp_body=make_section_card(sidebar,"⑨ EXPORT RESULTS",COL_EXPORT)
tk.Label(exp_body,
         text="Single Excel file — 6 sheets:\n"
              "  All Wells  ·  Anomaly Wells  ·  Cluster Summary\n"
              "  Parameter Weights  ·  Parameter Anomalies  ·  Insights Report",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=390,justify="left").pack(anchor="w",pady=(0,6))
export_btn=tk.Button(exp_body,text="  Export All Results  (Excel / CSV)",
                     command=export_results,bg="#14532d",fg=COL_EXPORT,
                     activebackground=_lighten("#14532d",20),activeforeground=COL_EXPORT,
                     font=FONT_H3,relief="flat",bd=0,cursor="hand2",
                     padx=10,pady=11,anchor="w",state="disabled")
export_btn.pack(fill="x",pady=2)
export_btn.bind("<Enter>",lambda e: export_btn.config(bg=_lighten("#14532d",20)))
export_btn.bind("<Leave>",lambda e: export_btn.config(bg="#14532d"))
Tooltip(export_btn,"Export all results to multi-sheet Excel (6 sheets incl. Parameter Anomalies)")

# ⑩ DATASET PREVIEW
prev_body=make_section_card(sidebar,"⑩ DATASET PREVIEW  (first 20 rows)",COL_PREVIEW,
                            fill="both",expand=True)
table_frame=tk.Frame(prev_body,bg=CARD); table_frame.pack(fill="both",expand=True)
tk.Frame(sidebar,bg=SIDEBAR,height=20).pack()

# ══════════════════════════════════════════════════════════════════════════════
# RIGHT PANEL
# ══════════════════════════════════════════════════════════════════════════════
rp_outer,right_inner=make_scrollable(root_pane,bg=BG)
root_pane.add(rp_outer,minsize=820)
rp_hdr=tk.Frame(right_inner,bg=BG,padx=16,pady=10); rp_hdr.pack(fill="x")
tk.Label(rp_hdr,text="Analysis Dashboard",bg=BG,fg=FG,
         font=("Georgia",15,"bold")).pack(side="left")
tk.Label(rp_hdr,
         text="v5.1  ·  Per-Parameter Anomaly  ·  Events  ·  RPM",
         bg=BG,fg=FG_DIM,font=FONT_SM).pack(side="right",pady=2)
tk.Frame(right_inner,bg=BORDER,height=1).pack(fill="x",padx=12)

stats_wrap=tk.Frame(right_inner,bg=BG,padx=10,pady=5); stats_wrap.pack(fill="x")
stat_defs_r1=[("Unique Wells","—",COL_DATASET),  ("Total Records","—",COL_CLUSTER),
              ("Clusters","—",COL_WEIGHTS),        ("Anomaly Recs","—",COL_ANOMALY)]
stat_defs_r2=[("Active Recs","—",COL_EVENTS),     ("Inactive Recs","—",ERR),
              ("Event Types","—",WARN),             ("Abnormal Recs","—",COL_ANOMALY)]
stat_defs_r3=[("RPM Running","—","#10b981"),       ("RPM Stopped","—","#f59e0b"),
              ("RPM Reverse","—","#ef4444"),        ("Param Anomalies","—",COL_PARAM_ANOM)]
stat_widgets={}
for stat_row_defs in [stat_defs_r1,stat_defs_r2,stat_defs_r3]:
    row=tk.Frame(stats_wrap,bg=BG); row.pack(fill="x",pady=2)
    for title,val,col in stat_row_defs:
        sf=tk.Frame(row,bg=CARD,padx=6,pady=6,
                    highlightbackground=col,highlightthickness=1)
        sf.pack(side="left",expand=True,fill="x",padx=3)
        sv=tk.StringVar(value=val)
        # Value label — auto-wraps, never overflows
        tk.Label(sf,textvariable=sv,bg=CARD,fg=col,
                 font=("Courier New",10,"bold"),
                 wraplength=120,justify="center",anchor="center").pack(fill="x")
        # Title label — fixed, clear, smaller
        tk.Label(sf,text=title,bg=CARD,fg=FG_DIM,
                 font=("Courier New",7),
                 wraplength=120,justify="center",anchor="center").pack(fill="x")
        stat_widgets[title]=sv

nb_wrap=tk.Frame(right_inner,bg=BG,padx=10,pady=6); nb_wrap.pack(fill="x")
sty=tk_ttk.Style()
sty.configure("CBM.TNotebook",     background=BG,tabmargins=[0,0,0,0])
sty.configure("CBM.TNotebook.Tab",background=CARD2,foreground=FG_DIM,
              padding=[12,6],font=FONT_H3)
sty.map("CBM.TNotebook.Tab",
        background=[("selected","#1e3a8a")],foreground=[("selected","#93c5fd")])
notebook=tk_ttk.Notebook(nb_wrap,style="CBM.TNotebook")
notebook.pack(fill="x")

TAB_H=500
cluster_tab    = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
pca_tab        = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
reservoir_tab  = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
production_tab = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
hidden_tab     = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
weight_tab     = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
event_tab      = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
rpm_tab        = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
param_anom_tab = tk.Frame(notebook,bg=BG_MID,height=TAB_H)

for tab,name in [
    (cluster_tab,    "  Clusters  "),
    (pca_tab,        "  PCA  "),
    (reservoir_tab,  "  3D Reservoir  "),
    (production_tab, "  Production  "),
    (hidden_tab,     "  Hidden Patterns  "),
    (weight_tab,     "  Param Weights  "),
    (event_tab,      "  Events  "),
    (rpm_tab,        "  RPM Status  "),
    (param_anom_tab, "  ★ Param Anomaly  "),
]:
    tab.pack_propagate(False); notebook.add(tab,text=name)

ins_wrap=tk.Frame(right_inner,bg=BG,padx=10,pady=6); ins_wrap.pack(fill="x")
ins_hdr=tk.Frame(ins_wrap,bg=BG); ins_hdr.pack(fill="x",pady=(0,4))
tk.Label(ins_hdr,text="AI Insights Report  —  Parameter Anomaly Diagnosis  (scroll to read)",
         bg=BG,fg=FG,font=FONT_H2).pack(side="left")
copy_btn=tk.Button(ins_hdr,text="Copy",
                   command=lambda:(app.clipboard_clear(),
                                   app.clipboard_append(explain_text.get("1.0","end"))),
                   bg=CARD2,fg=FG_DIM,activebackground=BORDER,activeforeground=FG,
                   font=FONT_XS,relief="flat",bd=0,cursor="hand2",padx=10,pady=4)
copy_btn.pack(side="right")
explain_text_frame = tk.Frame(ins_wrap, bg=CARD,
                              highlightbackground=BORDER, highlightthickness=1)
explain_text_frame.pack(fill="both", expand=True)
explain_xsb = tk_ttk.Scrollbar(explain_text_frame, orient="horizontal")
explain_vsb = tk_ttk.Scrollbar(explain_text_frame, orient="vertical")
explain_text = tk.Text(explain_text_frame, height=28, bg=CARD, fg="#a8f0c8",
                       font=("Courier New", 9), relief="flat", bd=0,
                       padx=14, pady=10, insertbackground=FG,
                       wrap="none",   # NO word-wrap: each line stays on its own row
                       xscrollcommand=explain_xsb.set,
                       yscrollcommand=explain_vsb.set,
                       highlightthickness=0, state="disabled")
explain_xsb.config(command=explain_text.xview)
explain_vsb.config(command=explain_text.yview)
explain_vsb.pack(side="right",  fill="y")
explain_xsb.pack(side="bottom", fill="x")
explain_text.pack(side="left", fill="both", expand=True)
tk.Frame(right_inner,bg=BG,height=20).pack()

explain_text.config(state="normal")
explain_text.insert("end",
    "==================================================\n"
    "  CBM AI Analytics Platform  v5.2\n"
    "  Per-Parameter Anomaly Analysis Edition\n"
    "==================================================\n"
    "\n"
    "  WHAT'S NEW IN v5.2:\n"
    "  --------------------------------------------------\n"
    "  * Per-Parameter Anomaly tab:\n"
    "    Each numeric column is individually checked.\n"
    "    See which parameter has outliers and why.\n"
    "\n"
    "  * Insights Report shows for each flagged column:\n"
    "    - How many records are anomalous\n"
    "    - Whether values are too HIGH or too LOW\n"
    "    - Mean, Std Dev, Median, Min, Max\n"
    "    - Skewness and Kurtosis\n"
    "    - Upper and lower z-score limits\n"
    "    - Plain-English reason and field interpretation\n"
    "\n"
    "  * Parameter Anomalies sheet added to Excel export.\n"
    "  * Adjustable z-score threshold  (sidebar: section IVb).\n"
    "\n"
    "  QUICK START:\n"
    "  --------------------------------------------------\n"
    "  1.  Upload your data file  (any format)\n"
    "  2.  Set number of clusters\n"
    "  3.  Assign parameter weights in the sidebar\n"
    "  4.  Set anomaly contamination %\n"
    "  4b. Set per-parameter z-score threshold\n"
    "  5.  Set RPM filter  (optional)\n"
    "  6.  Select features  (optional, or leave blank)\n"
    "  7.  Click  Run AI Analysis\n"
    "\n"
    "  After analysis:\n"
    "  -> View the  [* Param Anomaly]  tab\n"
    "  -> Read the Insights Report here for full diagnosis\n"
    "  -> Export to Excel for all 6 result sheets\n"
    "==================================================\n"
)
explain_text.config(state="disabled")

# Wire z-threshold into the pipeline
def run_pipeline():
    global active_df,active_X,active_xcols
    global active_anomaly_result,active_weight_result,active_event_result
    global active_rpm_result,active_param_anom

    x_cols,wdf,well_count,n_unique_wells,warn=resolve_features()
    if wdf is None or well_count==0:
        app.after(0,lambda: set_status(f"No usable rows: {warn}",WARN))
        app.after(0,progress_stop); return

    try:
        weight_map,raw_total=get_manual_weights()
        method_desc="User-assigned weights (from dataset numeric columns)"
        if not weight_map:
            weight_map={c:round(100/len(x_cols),2) for c in x_cols}
            method_desc="Equal weights (all spinboxes zero)"
        final_weights={}
        for c in x_cols: final_weights[c]=weight_map.get(c,1.0)
        total_fw=sum(final_weights.values())
        final_weights={c:v/total_fw*100 for c,v in final_weights.items()}

        X_w,use_cols=build_weighted_X(wdf,x_cols,final_weights)
        n_clusters=min(cluster_var.get(),well_count)
        wdf=wdf.copy()
        cluster_labels,inertia=run_clustering(X_w,n_clusters)
        wdf["cluster"]=cluster_labels

        wdf, rpm_col_used, rpm_result = add_rpm_status_column(wdf)
        if rpm_col_used:
            app.after(0, lambda rc=rpm_col_used: rpm_filter_info_var.set(
                f"RPM col: '{rc}'  |  "
                f"Running:{rpm_result['running']}  "
                f"Stopped:{rpm_result['stopped']}  "
                f"Reverse:{rpm_result['reverse']}"
            ))

        if rpm_col_used and "RPM_Status" in wdf.columns:
            allowed = []
            if rpm_filter_all_var.get():
                allowed = ["Running","Stopped","Reverse","No Data"]
            else:
                if rpm_filter_running_var.get(): allowed.append("Running")
                if rpm_filter_stopped_var.get(): allowed.append("Stopped")
                if rpm_filter_reverse_var.get(): allowed.append("Reverse")
                if rpm_filter_nodata_var.get():  allowed.append("No Data")
            if allowed and len(allowed) < 4:
                rpm_mask   = wdf["RPM_Status"].isin(allowed)
                X_w_anom   = X_w[rpm_mask.values]
                anom_filter_desc = f"RPM filter: {', '.join(allowed)}"
            else:
                X_w_anom         = X_w
                anom_filter_desc = "No RPM filter (all states)"
        else:
            X_w_anom         = X_w
            anom_filter_desc = "No RPM column detected"

        try:
            from model.test import test_model
            wdf=test_model(wdf)
        except Exception: pass

        method=anomaly_method_var.get(); contam=anomaly_contam_var.get()/100.0
        a_labels_sub,a_pct,a_idx_sub,a_name,a_scores_sub=\
            detect_anomalies(X_w_anom,contamination=contam,method=method)

        a_labels = np.ones(len(wdf), dtype=int)
        a_scores = np.zeros(len(wdf))
        if rpm_col_used and "RPM_Status" in wdf.columns and len(X_w_anom) < len(X_w):
            rpm_mask   = wdf["RPM_Status"].isin(allowed)
            sub_indices = np.where(rpm_mask.values)[0]
            for local_i, global_i in enumerate(sub_indices):
                a_labels[global_i] = a_labels_sub[local_i]
                a_scores[global_i] = a_scores_sub[local_i]
            a_idx = [sub_indices[i] for i in a_idx_sub]
        else:
            a_labels = a_labels_sub
            a_scores = a_scores_sub
            a_idx    = a_idx_sub

        a_name = f"{a_name}  [{anom_filter_desc}]"
        wdf["anomaly"]=a_labels
        anomaly_result={
            "labels":a_labels,"pct":a_pct,"indices":a_idx,
            "scores":a_scores,"detector_name":a_name,
            "n_anomaly":len(a_idx),"n_normal":len(a_labels)-len(a_idx),
        }
        event_result=analyse_events(wdf,a_labels)

        # ★ Per-parameter analysis using user-defined z threshold
        z_thr = zthresh_var.get()
        param_anom = analyse_per_parameter(wdf, use_cols, z_threshold=z_thr)

        active_df=wdf; active_X=X_w; active_xcols=use_cols
        active_anomaly_result=anomaly_result
        active_weight_result={"weight_map":final_weights,"method_desc":method_desc}
        active_event_result=event_result
        active_rpm_result=rpm_result
        active_param_anom=param_anom

        hx=cluster_hex(n_clusters)
        y_cols=[c for c in wdf.select_dtypes(include="number").columns
                if c not in use_cols and c!="cluster"][:2]
        insights=generate_insights(wdf,use_cols,well_count,n_unique_wells,hx,
                                   anomaly_result,event_result,
                                   final_weights,method_desc,inertia,rpm_result,
                                   param_anom)

        app.after(0,lambda: refresh_ui(
            wdf,X_w,use_cols,y_cols,hx,insights,well_count,n_unique_wells,
            anomaly_result,event_result,final_weights,method_desc,rpm_result,
            param_anom))
        app.after(0,lambda: set_status("Analysis complete",SUCCESS))
        app.after(0,progress_stop)
        app.after(0,lambda: export_btn.config(state="normal"))

    except Exception as e:
        traceback.print_exc()
        app.after(0,lambda err=str(e): set_status(f"Error: {err}",ERR))
        app.after(0,progress_stop)

app.mainloop()
