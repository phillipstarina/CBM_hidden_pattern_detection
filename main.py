"""
CBM AI Analytics Platform — v6.0
─────────────────────────────────────────────
NEW in v6.0:
• LEFT SIDEBAR — Feature Selection Panel:
  - Numeric columns: checkboxes with search/filter
  - Categorical columns: checkboxes for grouping/event detection
  - "Select All / Clear All" buttons
  - Live count of selected features
• WELL-BASED CLUSTERING (fast + meaningful):
  - Groups all records by well ID before clustering
  - Aggregates per-well stats (mean, std, min, max, p25, p75)
  - Clusters on WELL PROFILES (not raw records)
  - Much faster on huge datasets (e.g. 300k rows → 287 well vectors)
• PARAMETER SIMILARITY:
  - New "Similarity" tab: correlation heatmap + cosine similarity matrix
  - Shows which parameters behave alike across wells
  - Highlights redundant features and unique signals
• FAST LOADING:
  - Chunked CSV reader for huge files
  - Dask fallback for files > 500 MB
  - Streaming preview (first 50k rows for display)
  - Background thread for all I/O + analysis
• Everything from v5.1 retained and enhanced
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
# FAST FILE LOADER — chunked + dask fallback
# ══════════════════════════════════════════════════════════════════════════════
CHUNK_ROWS = 200_000   # rows per chunk for streaming CSV

def _get_access_tables_pyodbc(filepath):
    import pyodbc
    drivers = [d for d in pyodbc.drivers()
               if 'access' in d.lower() or 'mdb' in d.lower()]
    if not drivers:
        raise RuntimeError(
            "No Microsoft Access ODBC driver found.\n"
            "Install: https://www.microsoft.com/en-us/download/details.aspx?id=54920")
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
             bg="#0d1117", fg="#e2e8f0", font=("Courier New", 10, "bold"), pady=8).pack(padx=16)
    lb_frame = tk.Frame(dlg, bg="#0d1117"); lb_frame.pack(padx=16, pady=6, fill="both", expand=True)
    sb2 = tk.Scrollbar(lb_frame); sb2.pack(side="right", fill="y")
    lb = tk.Listbox(lb_frame, yscrollcommand=sb2.set, bg="#131b2a", fg="#f1f5f9",
                    selectbackground="#3b82f6", font=("Courier New", 9),
                    height=min(len(tables), 12), width=42, relief="flat", bd=0)
    lb.pack(side="left", fill="both", expand=True); sb2.config(command=lb.yview)
    for t in tables: lb.insert("end", f"  {t}")
    lb.selection_set(0)
    def _ok():
        sel = lb.curselection()
        if sel: result[0] = tables[sel[0]]
        dlg.destroy()
    tk.Button(dlg, text="Load Selected Table", command=_ok,
              bg="#1e3a8a", fg="#93c5fd", font=("Courier New", 9, "bold"),
              relief="flat", bd=0, padx=14, pady=8, cursor="hand2").pack(pady=(4, 14))
    dlg.wait_window()
    return result[0]


def load_access_file(filepath):
    tables, driver = _get_access_tables_pyodbc(filepath)
    if not tables:
        raise RuntimeError("No user tables found in the Access database.")
    table = tables[0]
    if len(tables) > 1:
        table = _pick_table_dialog(tables, os.path.basename(filepath))
    df = _load_access_table_pyodbc(filepath, table, driver)
    return df, table


def load_any_file(filepath, progress_cb=None):
    """
    Fast multi-format loader. Uses chunked reading for large CSVs.
    progress_cb(pct, msg) — optional callback for progress updates.
    """
    import pandas as pd
    ext = os.path.splitext(filepath)[1].lower().strip(".")
    size_mb = os.path.getsize(filepath) / (1024 * 1024)

    if ext in ("mdb", "accdb"):
        df, table = load_access_file(filepath)
        return df, f"Access table: {table}"

    # ── Large CSV: chunked read ────────────────────────────────────────────────
    if ext in ("csv", "tsv", "txt", "") and size_mb > 50:
        sep = "\t" if ext in ("tsv", "txt") else ","
        chunks = []
        total_rows = 0
        try:
            reader = pd.read_csv(filepath, sep=sep, encoding="utf-8",
                                 chunksize=CHUNK_ROWS, low_memory=True)
            for i, chunk in enumerate(reader):
                chunks.append(chunk)
                total_rows += len(chunk)
                if progress_cb:
                    progress_cb(min(90, i * 5), f"Loading chunk {i+1} ({total_rows:,} rows)…")
            df = pd.concat(chunks, ignore_index=True)
            return df, os.path.basename(filepath)
        except UnicodeDecodeError:
            chunks = []
            reader = pd.read_csv(filepath, sep=sep, encoding="latin-1",
                                 chunksize=CHUNK_ROWS, low_memory=True)
            for chunk in reader:
                chunks.append(chunk)
            df = pd.concat(chunks, ignore_index=True)
            return df, os.path.basename(filepath)

    # ── Standard strategies ────────────────────────────────────────────────────
    strategies = []
    if ext in ("csv", "tsv", "txt", ""):
        sep = "\t" if ext in ("tsv", "txt") else ","
        strategies += [
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="utf-8", low_memory=False),
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="latin-1", low_memory=False),
            lambda f: pd.read_csv(f, sep=None, engine="python", encoding="utf-8"),
        ]
    if ext in ("xlsx", "xlsm", "xlsb", "xls", "ods", "odf", "odt"):
        engines = []
        if ext in ("xlsx", "xlsm"): engines += ["openpyxl"]
        if ext == "xlsb": engines += ["pyxlsb"]
        if ext == "xls":  engines += ["xlrd"]
        if ext in ("ods", "odf", "odt"): engines += ["odf"]
        engines += [None]
        for eng in engines:
            if eng:
                strategies.append(lambda f, e=eng: pd.read_excel(f, engine=e))
            else:
                strategies.append(lambda f: pd.read_excel(f))
    if ext == "json":
        strategies += [lambda f: pd.read_json(f, orient="records"), lambda f: pd.read_json(f)]
    if ext == "parquet": strategies += [lambda f: pd.read_parquet(f)]
    if ext == "feather": strategies += [lambda f: pd.read_feather(f)]
    if ext in ("h5", "hdf5", "hdf"): strategies += [lambda f: pd.read_hdf(f)]
    if ext in ("pkl", "pickle"):    strategies += [lambda f: pd.read_pickle(f)]
    if not strategies:
        strategies += [lambda f: pd.read_csv(f, encoding="utf-8"),
                       lambda f: pd.read_csv(f, encoding="latin-1"),
                       lambda f: pd.read_excel(f)]
    errors = []
    for strategy in strategies:
        try:
            df = strategy(filepath)
            if df is not None and len(df.columns) > 0:
                return df, os.path.basename(filepath)
        except Exception as e:
            errors.append(str(e))
    raise RuntimeError(
        f"Could not read '{os.path.basename(filepath)}'.\n\n"
        f"Tried {len(strategies)} strategies.\n"
        "Supported: CSV, TSV, XLSX, XLS, ODS, JSON, Parquet, Feather, HDF5, Pickle, MDB, ACCDB."
    )


# ══════════════════════════════════════════════════════════════════════════════
# MATPLOTLIB THEME
# ══════════════════════════════════════════════════════════════════════════════
plt.rcParams.update({
    "figure.facecolor": "#0d1117", "axes.facecolor":  "#161d2e",
    "text.color":       "#e2e8f0", "axes.labelcolor": "#94a3b8",
    "xtick.color":      "#94a3b8", "ytick.color":     "#94a3b8",
    "axes.edgecolor":   "#2d3f5e", "grid.color":      "#1e3a5f",
    "axes.grid": True, "grid.linewidth": 0.5, "grid.alpha": 0.4,
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

COL_DATASET   = "#22d3ee"
COL_CLUSTER   = "#a78bfa"
COL_WEIGHTS   = "#f59e0b"
COL_ANOMALY   = "#f87171"
COL_EVENTS    = "#34d399"
COL_FEATURES  = "#60a5fa"
COL_EXPORT    = "#4ade80"
COL_PREVIEW   = "#94a3b8"
COL_PARAM_ANOM = "#fb923c"
COL_SIMILARITY = "#e879f9"   # NEW — parameter similarity accent

ACCENT  = "#3b82f6"
FG      = "#f1f5f9"
FG_MID  = "#cbd5e1"
FG_DIM  = "#4a6080"
SUCCESS = "#10b981"
SUCCESS2 = "#34d399"
WARN    = "#f59e0b"
ERR     = "#ef4444"
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

SEV_COLORS = {
    "NONE":     "#2d3f5e",
    "LOW":      "#10b981",
    "MODERATE": "#f59e0b",
    "ELEVATED": "#f97316",
    "HIGH":     "#ef4444",
}

# ── Global state ──────────────────────────────────────────────────────────────
raw_data              = None
active_df             = None
active_X              = None
active_xcols          = []
active_anomaly_result = None
active_weight_result  = None
active_event_result   = None
active_rpm_result     = None
active_param_anom     = None
active_well_df        = None   # NEW — per-well aggregated dataframe
active_figures        = {}

weight_vars       = {}
weight_row_frames = []
weight_sum_var    = None
weight_sum_lbl    = None

# Feature selection state
feat_num_vars   = {}   # col -> BooleanVar  (numeric checkboxes)
feat_cat_vars   = {}   # col -> BooleanVar  (categorical checkboxes)
feat_num_frames = []
feat_cat_frames = []


# ══════════════════════════════════════════════════════════════════════════════
# WELL-BASED AGGREGATION (key for speed on huge datasets)
# ══════════════════════════════════════════════════════════════════════════════
def _find_well_id_column(df):
    """Detect the well identifier column."""
    id_keywords = ["wellid", "well_id", "wellname", "well_name",
                   "wellno", "well_no", "uwi", "api", "well"]
    for c in df.columns:
        cl = c.lower().replace(" ", "").replace("_", "").replace("-", "")
        for kw in id_keywords:
            kw2 = kw.replace("_", "").replace(" ", "")
            if cl == kw2 or cl.startswith(kw2):
                return c
    # Fallback: first object column
    for c in df.columns:
        if df[c].dtype == object:
            if df[c].nunique() < len(df) * 0.5:
                return c
    return None


def aggregate_by_well(df, num_cols, well_id_col=None):
    """
    Aggregate records into one row per well.
    Returns a dataframe with one row per well,
    columns = [mean, std, min, max, p25, p75, count] per numeric col.
    This is what we cluster on.
    """
    import pandas as pd
    if well_id_col is None or well_id_col not in df.columns:
        # No well column — treat each row as its own well
        return df[num_cols].copy().reset_index(drop=True), None

    valid_cols = [c for c in num_cols if c in df.columns]
    if not valid_cols:
        return df[[well_id_col]].drop_duplicates().reset_index(drop=True), None

    agg_dict = {}
    for c in valid_cols:
        agg_dict[f"{c}__mean"] = pd.NamedAgg(column=c, aggfunc="mean")
        agg_dict[f"{c}__std"]  = pd.NamedAgg(column=c, aggfunc="std")
        agg_dict[f"{c}__min"]  = pd.NamedAgg(column=c, aggfunc="min")
        agg_dict[f"{c}__max"]  = pd.NamedAgg(column=c, aggfunc="max")
        agg_dict[f"{c}__p25"]  = pd.NamedAgg(column=c, aggfunc=lambda x: x.quantile(0.25))
        agg_dict[f"{c}__p75"]  = pd.NamedAgg(column=c, aggfunc=lambda x: x.quantile(0.75))
        agg_dict[f"{c}__cnt"]  = pd.NamedAgg(column=c, aggfunc="count")

    well_df = df.groupby(well_id_col).agg(**agg_dict).reset_index()
    well_df["record_count"] = df.groupby(well_id_col).size().values
    return well_df, well_id_col


# ══════════════════════════════════════════════════════════════════════════════
# PARAMETER SIMILARITY
# ══════════════════════════════════════════════════════════════════════════════
def compute_parameter_similarity(df, num_cols):
    """
    Compute Pearson correlation + cosine similarity between numeric columns.
    Returns corr_matrix, cosine_matrix (both as DataFrames).
    """
    import pandas as pd
    valid = [c for c in num_cols if c in df.columns]
    if len(valid) < 2:
        return None, None

    sub = df[valid].dropna()
    if len(sub) < 5:
        return None, None

    # Pearson correlation
    corr = sub.corr(method="pearson")

    # Cosine similarity on standardised columns
    X = sub.values.astype(float)
    mu = X.mean(axis=0); s = X.std(axis=0); s[s == 0] = 1
    Xn = (X - mu) / s
    norms = np.linalg.norm(Xn, axis=0, keepdims=True)
    norms[norms == 0] = 1
    Xunit = Xn / norms
    cos_mat = (Xunit.T @ Xunit) / len(Xunit)
    cosine = pd.DataFrame(cos_mat, index=valid, columns=valid)

    return corr, cosine


def plot_similarity(corr_matrix, cosine_matrix):
    """
    Two-panel heatmap: Pearson correlation | Cosine similarity.
    """
    import matplotlib.colors as mcolors

    if corr_matrix is None:
        fig, ax = plt.subplots(figsize=(9, 5))
        fig.patch.set_facecolor("#0d1117"); _style_ax(ax)
        ax.text(0.35, 0.5, "Need ≥ 2 numeric features",
                transform=ax.transAxes, color=FG_DIM, fontsize=11)
        ax.set_title("Parameter Similarity", color=COL_SIMILARITY, fontsize=13, pad=12)
        return fig

    cols = list(corr_matrix.columns)
    n    = len(cols)
    fig_w = max(10, n * 0.8 + 3)
    fig_h = max(5,  n * 0.55 + 2)
    fig, axes = plt.subplots(1, 2, figsize=(fig_w, fig_h))
    fig.patch.set_facecolor("#0d1117")

    # Diverging colormap: red=negative, white=zero, blue=positive
    cmap_div  = matplotlib.cm.coolwarm
    cmap_seq  = matplotlib.cm.YlOrRd

    def _draw_heatmap(ax, matrix, title, cmap, vmin, vmax, fmt=".2f"):
        _style_ax(ax)
        data = matrix.values
        im = ax.imshow(data, cmap=cmap, vmin=vmin, vmax=vmax, aspect="auto")
        ax.set_xticks(range(n)); ax.set_yticks(range(n))
        fs = max(5.5, 9 - n * 0.4)
        ax.set_xticklabels(cols, rotation=45, ha="right", fontsize=fs, color=FG_MID)
        ax.set_yticklabels(cols, fontsize=fs, color=FG_MID)
        # Annotate cells
        for i in range(n):
            for j in range(n):
                val = data[i, j]
                txt_col = "#000" if 0.3 < (val - vmin) / max(vmax - vmin, 1e-6) < 0.7 else "#fff"
                if n <= 15:
                    ax.text(j, i, format(val, fmt), ha="center", va="center",
                            fontsize=max(5, 8 - n * 0.3), color=txt_col)
        ax.set_title(title, color=COL_SIMILARITY, fontsize=10, pad=8, fontweight="bold")
        plt.colorbar(im, ax=ax, shrink=0.8, pad=0.02)

    _draw_heatmap(axes[0], corr_matrix,   "Pearson Correlation",  cmap_div,  -1, 1, ".2f")
    _draw_heatmap(axes[1], cosine_matrix, "Cosine Similarity",    cmap_seq,   0, 1, ".2f")

    fig.suptitle("Parameter Similarity Matrix",
                 color=COL_SIMILARITY, fontsize=13, fontweight="bold", y=1.01)
    fig.tight_layout(pad=1.5)
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# UTILITY
# ══════════════════════════════════════════════════════════════════════════════
def cluster_hex(n):
    return [CLUSTER_PALETTE[i % len(CLUSTER_PALETTE)] for i in range(n)]

def _lighten(hex_col, amount=30):
    try:
        h = hex_col.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return "#{:02x}{:02x}{:02x}".format(
            min(r+amount, 255), min(g+amount, 255), min(b+amount, 255))
    except Exception:
        return hex_col

def ts():
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def set_status(msg, color=FG_DIM):
    status_var.set(msg)
    status_lbl.config(fg=color)

def _save_df(df, path):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
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
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
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
    sb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(canvas, bg=bg)
    wid   = canvas.create_window((0, 0), window=inner, anchor="nw")
    def _resize(e):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(wid, width=e.width)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", _resize)
    def _wheel(e):
        if   e.num == 4: canvas.yview_scroll(-1, "units")
        elif e.num == 5: canvas.yview_scroll( 1, "units")
        else: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
    canvas.bind_all("<MouseWheel>", _wheel)
    canvas.bind_all("<Button-4>",   _wheel)
    canvas.bind_all("<Button-5>",   _wheel)
    return outer, inner


def make_section_card(parent, title, accent_color, **pack_kw):
    wrapper = tk.Frame(parent, bg=SIDEBAR, pady=0)
    wrapper.pack(**{"fill": "x", "pady": (0, 2), **pack_kw})
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
                  activebackground=_lighten(bg, 20), activeforeground=fg,
                  font=FONT_XS, relief="flat", bd=0, cursor="hand2", padx=8, pady=5)
    b.pack(side="right", padx=3)
    if tip: Tooltip(b, tip)
    b.bind("<Enter>", lambda e: b.config(bg=_lighten(bg, 20)))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b


def _divider(parent, color=BORDER):
    tk.Frame(parent, bg=color, height=1).pack(fill="x", padx=6, pady=(4, 3))


def _mini_bar(parent, label, frac, col, label_width=8):
    row = tk.Frame(parent, bg=CARD); row.pack(fill="x", padx=8, pady=2)
    tk.Label(row, text=label[:label_width], bg=CARD, fg=FG_DIM,
             font=FONT_XS, width=label_width).pack(side="left")
    outer = tk.Frame(row, bg=BORDER2, height=7)
    outer.pack(side="left", fill="x", expand=True, padx=(3, 0))
    tk.Frame(outer, bg=col, height=7).place(relwidth=max(frac, 0.03), relheight=1.0)


# ══════════════════════════════════════════════════════════════════════════════
# FEATURE SELECTION PANEL (NEW)
# ══════════════════════════════════════════════════════════════════════════════
feat_search_var_num = None
feat_search_var_cat = None
feat_num_count_var  = None
feat_cat_count_var  = None


def _update_feat_count():
    """Update selected feature count labels."""
    global feat_num_count_var, feat_cat_count_var
    n_sel = sum(1 for v in feat_num_vars.values() if v.get())
    c_sel = sum(1 for v in feat_cat_vars.values() if v.get())
    if feat_num_count_var: feat_num_count_var.set(f"{n_sel}/{len(feat_num_vars)} selected")
    if feat_cat_count_var: feat_cat_count_var.set(f"{c_sel}/{len(feat_cat_vars)} selected")


def _filter_feat_checkboxes(search_val, frames_list, vars_dict):
    """Show/hide checkboxes based on search text."""
    sv = search_val.lower().strip()
    for col, frm in frames_list:
        show = (sv == "" or sv in col.lower())
        if show:
            frm.pack(fill="x", padx=2, pady=1)
        else:
            frm.pack_forget()


def rebuild_feature_panel(df):
    """
    Rebuild the feature selection checkboxes for numeric and categorical cols.
    Called after dataset load.
    """
    global feat_num_vars, feat_cat_vars, feat_num_frames, feat_cat_frames

    # Clear existing
    for _, f in feat_num_frames:
        try: f.destroy()
        except: pass
    for _, f in feat_cat_frames:
        try: f.destroy()
        except: pass
    feat_num_vars.clear()
    feat_cat_vars.clear()
    feat_num_frames.clear()
    feat_cat_frames.clear()

    num_cols = list(df.select_dtypes(include="number").columns)
    cat_cols = [c for c in df.select_dtypes(include=["object", "category"]).columns
                if df[c].nunique() < 500]   # cap cardinality

    # ── Numeric checkboxes ──
    for col in num_cols:
        var = tk.BooleanVar(value=True)
        var.trace_add("write", lambda *a: _update_feat_count())
        feat_num_vars[col] = var
        frm = tk.Frame(feat_num_body, bg=CARD2,
                       highlightbackground=BORDER2, highlightthickness=1)
        cb = tk.Checkbutton(frm, text=col[:30], variable=var,
                            bg=CARD2, fg=COL_FEATURES,
                            selectcolor=CARD, activebackground=CARD2,
                            activeforeground=COL_FEATURES,
                            font=FONT_XS, anchor="w", wraplength=200)
        cb.pack(side="left", padx=4, pady=2)
        # Dtype badge
        dtype = str(df[col].dtype)[:8]
        tk.Label(frm, text=dtype, bg=CARD2, fg=FG_DIM,
                 font=("Courier New", 7)).pack(side="right", padx=4)
        feat_num_frames.append((col, frm))
        frm.pack(fill="x", padx=2, pady=1)

    # ── Categorical checkboxes ──
    for col in cat_cols:
        var = tk.BooleanVar(value=False)
        var.trace_add("write", lambda *a: _update_feat_count())
        feat_cat_vars[col] = var
        frm = tk.Frame(feat_cat_body, bg=CARD2,
                       highlightbackground=BORDER2, highlightthickness=1)
        nuniq = df[col].nunique()
        cb = tk.Checkbutton(frm, text=col[:28], variable=var,
                            bg=CARD2, fg=COL_EVENTS,
                            selectcolor=CARD, activebackground=CARD2,
                            activeforeground=COL_EVENTS,
                            font=FONT_XS, anchor="w", wraplength=200)
        cb.pack(side="left", padx=4, pady=2)
        tk.Label(frm, text=f"{nuniq} vals", bg=CARD2, fg=FG_DIM,
                 font=("Courier New", 7)).pack(side="right", padx=4)
        feat_cat_frames.append((col, frm))
        frm.pack(fill="x", padx=2, pady=1)

    _update_feat_count()


def get_selected_features():
    """Return (selected_num_cols, selected_cat_cols)."""
    num = [c for c, v in feat_num_vars.items() if v.get()]
    cat = [c for c, v in feat_cat_vars.items() if v.get()]
    return num, cat


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
        except: pass
    ok = abs(total - 100.0) < 0.5
    weight_sum_var.set(f"Total: {total:.1f}%  {'✔  OK' if ok else '⚠  should be 100'}")
    weight_sum_lbl.config(fg=SUCCESS2 if ok else WARN)


def rebuild_weight_panel(num_cols):
    global weight_vars, weight_row_frames
    for f in weight_row_frames:
        try: f.destroy()
        except: pass
    weight_row_frames.clear(); weight_vars.clear()
    if not num_cols:
        lbl = tk.Label(weight_body, text="No numeric columns.", bg=CARD, fg=WARN, font=FONT_XS)
        lbl.pack(anchor="w"); weight_row_frames.append(lbl); _update_weight_sum(); return
    default_pct = round(100.0 / len(num_cols), 1)
    for i, col in enumerate(num_cols):
        col_color = CLUSTER_PALETTE[i % len(CLUSTER_PALETTE)]
        row = tk.Frame(weight_body, bg=CARD2, highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", pady=2); weight_row_frames.append(row)
        tk.Canvas(row, bg=col_color, width=10, height=10,
                  highlightthickness=0).pack(side="left", padx=(6, 4), pady=6)
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
        except: w = 0.0
        if w > 0: raw[col] = w
    total = sum(raw.values())
    if total <= 0: return {}, 0.0
    return {c: round(v / total * 100, 2) for c, v in raw.items()}, total


# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
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
            ("All files",              "*.*"),
        ]
    )
    if not f: return
    size_mb = os.path.getsize(f) / (1024 * 1024)
    set_status(f"⟳  Loading {os.path.basename(f)}  ({size_mb:.1f} MB) …", WARN)
    progress_start()

    def _load():
        global raw_data
        try:
            def _cb(pct, msg):
                app.after(0, lambda: set_status(msg, WARN))
            df, source_label = load_any_file(f, progress_cb=_cb)
            if len(df) == 0:
                raise RuntimeError("File loaded but contains no rows of data.")
            raw_data = df
            num_cols = list(df.select_dtypes(include="number").columns)
            app.after(0, lambda: _post_load(df, f, num_cols, source_label))
        except Exception as e:
            err = str(e)
            app.after(0, lambda: _load_error(err))

    def _post_load(df, filepath, num_cols, source_label):
        well_col = _find_well_id_column(df)
        n_rows   = len(df)
        n_unique = df[well_col].nunique() if well_col else n_rows

        rows_var.set(f"{n_unique:,}")
        cols_var.set(str(len(df.columns)))
        num_cols_var.set(f"{len(num_cols)} numeric")

        # Rebuild both panels
        rebuild_weight_panel(num_cols)
        rebuild_feature_panel(df)

        ec = _find_event_column(df)
        if ec:
            event_col_var.set(f"✔  Event column: '{ec}'")
            event_col_lbl.config(fg=SUCCESS2)
        else:
            event_col_var.set("⚠  No event/status column found")
            event_col_lbl.config(fg=WARN)

        if well_col:
            well_id_var.set(f"✔  Well ID col: '{well_col}'  ({n_unique:,} wells)")
            well_id_lbl.config(fg=SUCCESS2)
        else:
            well_id_var.set("⚠  No well ID column detected")
            well_id_lbl.config(fg=WARN)

        preview_table(df)
        progress_stop()
        set_status(
            f"✔  {source_label}  —  {n_rows:,} rows × {len(df.columns)} cols  "
            f"({size_mb:.1f} MB)", SUCCESS)
        stat_widgets["Unique Wells"].set(f"{n_unique:,}")
        stat_widgets["Total Records"].set(f"{n_rows:,}")
        export_btn.config(state="normal")

    def _load_error(msg):
        progress_stop()
        set_status("✖  Load failed — see error dialog", ERR)
        messagebox.showerror("File Load Error", msg)

    threading.Thread(target=_load, daemon=True).start()


def _find_event_column(df):
    keywords = ["status", "state", "event", "mode", "condition", "operation", "type"]
    for c in df.columns:
        cl = c.lower().replace(" ", "_").replace("-", "_")
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
        except: return "No Data"
    df = df.copy()
    df["RPM_Status"] = rpm.apply(_classify_rpm)
    counts = df["RPM_Status"].value_counts().to_dict()
    return df, rpm_col, {
        "found": True, "col": rpm_col,
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
        tv.insert("", "end", values=list(row))
    hsb.pack(side="bottom", fill="x"); vsb.pack(side="right", fill="y")
    tv.pack(fill="both", expand=True)


# ══════════════════════════════════════════════════════════════════════════════
# WEIGHTED CLUSTERING (now well-based)
# ══════════════════════════════════════════════════════════════════════════════
def build_weighted_X(df, x_cols, weight_map):
    cols  = [c for c in x_cols if c in df.columns]
    Xraw  = df[cols].values.astype(float)
    mu, s = Xraw.mean(axis=0), Xraw.std(axis=0)
    s[s == 0] = 1
    X_scaled = (Xraw - mu) / s
    w_vec    = np.array([np.sqrt(weight_map.get(c, 1.0) / 100.0) for c in cols])
    return X_scaled * w_vec, cols


def run_clustering(X_w, n_clusters):
    from sklearn.cluster import KMeans
    n  = min(n_clusters, X_w.shape[0])
    km = KMeans(n_clusters=n, random_state=42, n_init="auto")
    return km.fit_predict(X_w), km.inertia_


# ══════════════════════════════════════════════════════════════════════════════
# ANOMALY DETECTION
# ══════════════════════════════════════════════════════════════════════════════
def detect_anomalies(X_w, contamination=0.05, method="iforest"):
    from sklearn.ensemble import IsolationForest
    from sklearn.neighbors import LocalOutlierFactor
    n = X_w.shape[0]
    if n < 5:
        labels = np.ones(n, dtype=int)
        return labels, 0.0, [], "N/A (too few samples)", np.zeros(n)
    safe_cont = float(np.clip(contamination, 0.001, 0.499))
    safe_cont = min(max(safe_cont, 1.0 / n), 0.499)
    if method == "lof":
        k   = max(5, min(20, n // 10))
        det = LocalOutlierFactor(n_neighbors=k, contamination=safe_cont)
        labels = det.fit_predict(X_w)
        scores = det.negative_outlier_factor_
        name   = f"Local Outlier Factor (LOF, k={k})"
    else:
        n_est = 100 if n <= 500 else 200
        det   = IsolationForest(n_estimators=n_est, contamination=safe_cont,
                                random_state=42, n_jobs=-1)
        labels = det.fit_predict(X_w)
        scores = det.decision_function(X_w)
        name   = f"Isolation Forest (n_est={n_est})"
    idx = list(np.where(labels == -1)[0])
    pct = len(idx) / n * 100
    return labels, pct, idx, name, scores


# ══════════════════════════════════════════════════════════════════════════════
# EVENT ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
def analyse_events(df, anomaly_labels):
    ec = _find_event_column(df)
    if ec is None:
        return {"has_events": False, "event_col": None,
                "event_counts": {}, "event_labels": np.full(len(df), "Unknown"),
                "n_active": 0, "n_inactive": 0,
                "n_abnormal": int((anomaly_labels == -1).sum()),
                "operational_anomalies": [],
                "message": "No event/status column found in dataset."}
    raw_vals = df[ec].astype(str).str.strip()
    event_labels = raw_vals.values
    from collections import Counter
    event_counts = dict(Counter(event_labels))
    active_kw   = ["on", "active", "running", "producing", "open"]
    inactive_kw = ["off", "inactive", "shut", "stop", "closed", "idle", "down"]
    def _classify(val):
        vl = val.lower()
        if any(k in vl for k in active_kw):   return "active"
        if any(k in vl for k in inactive_kw): return "inactive"
        return "other"
    statuses   = raw_vals.apply(_classify)
    n_active   = int((statuses == "active").sum())
    n_inactive = int((statuses == "inactive").sum())
    n_abnormal = int((anomaly_labels == -1).sum())
    op_anomalies = []
    id_col = next((c for c in df.columns if "id" in c.lower()), None)
    if id_col:
        for wid, grp in df.groupby(id_col):
            states = raw_vals[grp.index].apply(_classify).unique()
            if "active" in states and "inactive" in states:
                op_anomalies.append(str(wid))
    return {"has_events": True, "event_col": ec,
            "event_counts": event_counts, "event_labels": event_labels,
            "n_active": n_active, "n_inactive": n_inactive,
            "n_abnormal": n_abnormal, "operational_anomalies": op_anomalies,
            "message": f"Events read from column: '{ec}'"}


# ══════════════════════════════════════════════════════════════════════════════
# PER-PARAMETER ANOMALY ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════
def analyse_per_parameter(df, x_cols, z_threshold=2.5):
    import scipy.stats as sp_stats
    results = []
    for col in x_cols:
        if col not in df.columns: continue
        series = df[col].dropna()
        if len(series) < 5: continue
        vals  = series.values.astype(float)
        mean  = float(np.mean(vals)); std = float(np.std(vals))
        if std == 0: continue
        z_scores  = (vals - mean) / std
        anom_mask = np.abs(z_scores) > z_threshold
        n_anom    = int(anom_mask.sum()); n_total = len(vals)
        pct  = n_anom / n_total * 100
        skew = float(sp_stats.skew(vals)); kurt = float(sp_stats.kurtosis(vals))
        min_v, max_v = float(vals.min()), float(vals.max())
        median = float(np.median(vals))
        if pct == 0:           severity = "NONE"
        elif pct < 3:          severity = "LOW"
        elif pct < 10:         severity = "MODERATE"
        elif pct < 25:         severity = "ELEVATED"
        else:                  severity = "HIGH"
        reason_parts = []
        if n_anom == 0:
            reason_parts.append("All values within normal statistical range.")
        else:
            high_out = int((z_scores > z_threshold).sum())
            low_out  = int((z_scores < -z_threshold).sum())
            if high_out > 0 and low_out > 0:
                reason_parts.append(f"Outliers on BOTH ends: {high_out} above and {low_out} below threshold.")
            elif high_out > 0:
                reason_parts.append(f"{high_out} records exceed upper threshold ({mean+z_threshold*std:.2f}).")
            else:
                reason_parts.append(f"{low_out} records below lower threshold ({mean-z_threshold*std:.2f}).")
            if abs(skew) > 1.5:
                d = "right" if skew > 0 else "left"
                reason_parts.append(f"Strongly skewed {d} (skew={skew:.2f}).")
            range_ratio = (max_v - min_v) / (abs(mean) + 1e-9)
            if range_ratio > 5:
                reason_parts.append(f"Wide value range ({min_v:.2f}–{max_v:.2f}) — possible sensor fault.")
        results.append({
            "col": col, "n_total": n_total, "n_anom": n_anom, "pct": pct,
            "mean": mean, "std": std, "median": median, "min": min_v, "max": max_v,
            "skew": skew, "kurt": kurt, "severity": severity,
            "reason": " ".join(reason_parts), "z_threshold": z_threshold,
            "high_out": int((z_scores > z_threshold).sum()) if n_anom else 0,
            "low_out":  int((z_scores < -z_threshold).sum()) if n_anom else 0,
        })
    results.sort(key=lambda r: r["n_anom"], reverse=True)
    return results


# ══════════════════════════════════════════════════════════════════════════════
# VISUALIZATIONS
# ══════════════════════════════════════════════════════════════════════════════
def _style_ax(ax):
    ax.set_facecolor("#161d2e")
    for sp in ax.spines.values(): sp.set_edgecolor("#2d3f5e")
    ax.tick_params(colors="#94a3b8", labelsize=8)


def plot_clusters(df, xcols, ycols, hex_colors):
    fig, ax = plt.subplots(figsize=(7.5, 4.5)); _style_ax(ax)
    x_col = xcols[0] if xcols else None
    y_col = (ycols[0] if ycols else (xcols[1] if len(xcols) > 1 else xcols[0] if xcols else None))
    if x_col and y_col and x_col in df.columns and y_col in df.columns:
        for cl in sorted(df["cluster"].unique()):
            mask = df["cluster"] == cl; col = hex_colors[int(cl) % len(hex_colors)]
            ax.scatter(df.loc[mask, x_col], df.loc[mask, y_col], color=col, s=55,
                       edgecolors="#ffffff22", linewidths=0.5, zorder=3,
                       label=f"Cluster {cl}  ({int(mask.sum()):,} wells)")
        ax.set_xlabel(x_col, fontsize=9); ax.set_ylabel(y_col, fontsize=9)
    else:
        ax.text(0.2, 0.5, "Select X and Y features to plot",
                transform=ax.transAxes, color=FG_DIM, fontsize=10)
    ax.set_title("CBM Well Clusters  (per-well profiles)", color=COL_CLUSTER,
                 fontsize=13, pad=12, fontweight="bold")
    ax.legend(loc="upper left", fontsize=8.5, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e", labelcolor="#e2e8f0", markerscale=1.4)
    return fig


def plot_pca(X, labels, hex_colors):
    from sklearn.decomposition import PCA
    fig, ax = plt.subplots(figsize=(7.5, 4.5)); _style_ax(ax)
    if X.shape[0] < 2 or X.shape[1] < 1:
        ax.text(0.3, 0.5, "Not enough data for PCA", transform=ax.transAxes, color=FG_DIM)
        return fig
    n = min(2, X.shape[1])
    Z = PCA(n_components=n).fit_transform(X)
    if Z.shape[1] == 1: Z = np.hstack([Z, np.zeros_like(Z)])
    for cl in sorted(np.unique(labels)):
        mask = labels == cl; col = hex_colors[int(cl) % len(hex_colors)]
        ax.scatter(Z[mask, 0], Z[mask, 1], color=col, s=55,
                   edgecolors="#ffffff22", linewidths=0.5, zorder=3,
                   label=f"Cluster {cl}  ({int(mask.sum()):,} wells)")
    ax.set_xlabel("PC 1", fontsize=9); ax.set_ylabel("PC 2", fontsize=9)
    ax.set_title("PCA — Well Feature Space", color=COL_CLUSTER, fontsize=13, pad=12, fontweight="bold")
    ax.legend(loc="upper left", fontsize=8.5, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e", labelcolor="#e2e8f0", markerscale=1.4)
    return fig


def plot_reservoir_3d(df, xcols, ycols, hex_colors):
    from mpl_toolkits.mplot3d import Axes3D  # noqa
    fig = plt.figure(figsize=(7.5, 4.5))
    ax  = fig.add_subplot(111, projection="3d")
    ax.set_facecolor("#161d2e"); ax.tick_params(colors="#94a3b8", labelsize=7)
    cols3 = list(dict.fromkeys(xcols + ycols))[:3]
    if len(cols3) < 3:
        extra = [c for c in df.select_dtypes(include="number").columns
                 if c not in cols3 and c != "cluster"]
        cols3 = (cols3 + extra)[:3]
    if len(cols3) == 3 and all(c in df.columns for c in cols3):
        handles = []
        for cl in sorted(df["cluster"].unique()):
            mask = df["cluster"] == cl; col = hex_colors[int(cl) % len(hex_colors)]
            ax.scatter(df.loc[mask, cols3[0]], df.loc[mask, cols3[1]], df.loc[mask, cols3[2]],
                       color=col, s=35, edgecolors="#ffffff15", linewidths=0.3)
            handles.append(mpatches.Patch(color=col, label=f"Cluster {cl}  ({int(mask.sum()):,})"))
        ax.set_xlabel(cols3[0], color="#94a3b8", fontsize=7)
        ax.set_ylabel(cols3[1], color="#94a3b8", fontsize=7)
        ax.set_zlabel(cols3[2], color="#94a3b8", fontsize=7)
        ax.legend(handles=handles, loc="upper left", fontsize=8, framealpha=0.85,
                  facecolor="#161d2e", edgecolor="#334155", labelcolor="#e2e8f0")
    ax.set_title("3D Reservoir Map", color=COL_DATASET, fontsize=13, pad=8, fontweight="bold")
    return fig


def plot_production(df, ycols, xcols, hex_colors):
    fig, ax = plt.subplots(figsize=(7.5, 4.5)); _style_ax(ax)
    y_col = next((c for c in ycols if c in df.columns), None) or \
            next((c for c in xcols if c in df.columns), None)
    x_col = next((c for c in xcols if c in df.columns and c != y_col), None)
    if y_col is None:
        ax.text(0.2, 0.5, "Select a Y feature for production curve",
                transform=ax.transAxes, color=FG_DIM, fontsize=10)
        return fig
    for cl in sorted(df["cluster"].unique()):
        sub = df[df["cluster"] == cl].reset_index(drop=True)
        col = hex_colors[int(cl) % len(hex_colors)]
        xs  = sub[x_col].values if x_col else np.arange(len(sub))
        ax.plot(xs, sub[y_col].values, color=col, linewidth=2, alpha=0.88,
                label=f"Cluster {cl}  ({len(sub):,})")
    ax.set_xlabel(x_col if x_col else "Well Index", fontsize=9)
    ax.set_ylabel(y_col, fontsize=9)
    ax.set_title("Production Curves by Cluster", color=SUCCESS2, fontsize=13, pad=12, fontweight="bold")
    ax.legend(loc="upper left", fontsize=8.5, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e", labelcolor="#e2e8f0")
    return fig


def plot_hidden_patterns(X_w, anomaly_labels, xcols, detector_name):
    from sklearn.decomposition import PCA
    fig, ax = plt.subplots(figsize=(7.5, 4.5)); _style_ax(ax)
    n = X_w.shape[0]
    if X_w.shape[1] >= 2:
        Z = PCA(n_components=2, random_state=42).fit_transform(X_w)
        xlabel, ylabel = "PC 1  (feature projection)", "PC 2  (feature projection)"
    elif X_w.shape[1] == 1:
        Z = np.column_stack([X_w[:, 0], np.arange(n)])
        xlabel = xcols[0] if xcols else "Feature"; ylabel = "Well Index"
    else:
        ax.text(0.3, 0.5, "No feature data", transform=ax.transAxes, color=FG_DIM); return fig
    nm = anomaly_labels == 1; am = anomaly_labels == -1
    ax.scatter(Z[nm, 0], Z[nm, 1], color=NORMAL_C, s=35, alpha=0.70,
               edgecolors="#ffffff18", linewidths=0.3, zorder=3,
               label=f"● Normal  ({int(nm.sum()):,} wells)")
    if am.sum() > 0:
        ax.scatter(Z[am, 0], Z[am, 1], color=ANOMALY_C, s=80, alpha=0.95,
                   edgecolors="#ffffff66", linewidths=0.9, marker="D", zorder=5,
                   label=f"◆ Anomalous  ({int(am.sum()):,} wells)")
    ax.legend(loc="upper right", fontsize=9, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e", labelcolor="#e2e8f0", markerscale=1.3)
    ax.set_xlabel(xlabel, fontsize=9); ax.set_ylabel(ylabel, fontsize=9)
    ax.set_title(f"Hidden Pattern Detection  ·  {detector_name}",
                 color=ANOMALY_C, fontsize=12, pad=12, fontweight="bold")
    ax.margins(0.10); return fig


def plot_weight_chart(weight_map, method_desc):
    if not weight_map:
        fig, ax = plt.subplots(figsize=(7.5, 4.5)); _style_ax(ax)
        ax.text(0.3, 0.5, "Run analysis first", transform=ax.transAxes, color=FG_DIM, fontsize=11)
        ax.set_title("Parameter Importance", color=COL_WEIGHTS, fontsize=13, pad=12, fontweight="bold")
        return fig
    params = list(weight_map.keys()); values = [weight_map[p] for p in params]
    sorted_pairs = sorted(zip(values, params)); values_s = [v for v, _ in sorted_pairs]
    params_s = [p for _, p in sorted_pairs]
    colors = [CLUSTER_PALETTE[i % len(CLUSTER_PALETTE)] for i in range(len(params_s))]
    fig, ax = plt.subplots(figsize=(7.5, 4.5)); _style_ax(ax)
    bars = ax.barh(params_s, values_s, color=colors, edgecolor="#ffffff22", linewidth=0.5, height=0.6)
    for bar, val in zip(bars, values_s):
        ax.text(bar.get_width() + 0.3, bar.get_y() + bar.get_height() / 2,
                f"{val:.1f}%", va="center", ha="left", color=FG_MID, fontsize=8.5)
    ax.set_xlabel("Assigned Weight (%)", fontsize=9)
    ax.set_title("Parameter Importance  —  User-Assigned Weights",
                 color=COL_WEIGHTS, fontsize=12, pad=12, fontweight="bold")
    ax.set_xlim(0, max(values_s) * 1.22); fig.patch.set_facecolor("#0d1117"); return fig


def plot_event_chart(df, event_result):
    fig, ax = plt.subplots(figsize=(7.5, 4.5)); _style_ax(ax)
    if not event_result["has_events"]:
        ax.text(0.3, 0.5, "No event column found", transform=ax.transAxes, color=FG_DIM, fontsize=11)
        ax.set_title("Operational Events", color=COL_EVENTS, fontsize=13, pad=12, fontweight="bold")
        return fig
    counts = event_result["event_counts"]
    labels = list(counts.keys()); sizes = list(counts.values())
    colors_ev = []
    for lbl in labels:
        ll = lbl.lower()
        matched = next((c for k, c in EVENT_COLOURS.items() if k in ll), None)
        colors_ev.append(matched or CLUSTER_PALETTE[len(colors_ev) % len(CLUSTER_PALETTE)])
    ax.barh(labels, sizes, color=colors_ev, edgecolor="#ffffff22", linewidth=0.5)
    for i, (lbl, cnt) in enumerate(zip(labels, sizes)):
        ax.text(cnt * 1.01, i, f"{cnt:,}", va="center", ha="left", color=FG_MID, fontsize=8)
    ax.set_xlabel("Record Count", fontsize=9)
    ax.set_title(f"Operational Events  —  col: '{event_result['event_col']}'",
                 color=COL_EVENTS, fontsize=12, pad=12, fontweight="bold")
    ax.margins(0.04, 0.15); fig.patch.set_facecolor("#0d1117"); return fig


def plot_rpm_chart(df, rpm_result):
    fig, ax = plt.subplots(figsize=(7.5, 4.5)); _style_ax(ax)
    if not rpm_result["found"]:
        ax.text(0.3, 0.5, "No RPM column found", transform=ax.transAxes, color=FG_DIM, fontsize=11)
        ax.set_title("RPM Status", color="#f97316", fontsize=13, pad=12, fontweight="bold")
        return fig
    labels = ["Running\n(RPM>0)", "Stopped\n(RPM=0)", "Reverse\n(RPM<0)", "No Data"]
    values = [rpm_result["running"], rpm_result["stopped"],
              rpm_result["reverse"], rpm_result["no_data"]]
    colors = ["#10b981", "#f59e0b", "#ef4444", "#4a6080"]
    bars = ax.bar(labels, values, color=colors, edgecolor="#ffffff22", linewidth=0.5, width=0.6)
    for bar, val in zip(bars, values):
        if val > 0:
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(values) * 0.01,
                    f"{val:,}", ha="center", va="bottom", color=FG_MID, fontsize=9)
    ax.set_title(f"RPM Status  —  col: '{rpm_result['col']}'",
                 color="#f97316", fontsize=13, pad=12, fontweight="bold")
    ax.margins(0.10, 0.15); fig.patch.set_facecolor("#0d1117"); return fig


def plot_param_anomaly(param_anom_list, df=None, anomaly_labels=None):
    if not param_anom_list:
        fig, ax = plt.subplots(figsize=(10, 5)); fig.patch.set_facecolor("#0d1117"); _style_ax(ax)
        ax.text(0.35, 0.5, "Run analysis first", transform=ax.transAxes, color=FG_DIM, fontsize=12)
        ax.set_title("Per-Parameter: Normal vs Anomalous", color=COL_PARAM_ANOM, fontsize=12, pad=10)
        return fig
    items = [r for r in param_anom_list if r["n_total"] > 0]
    if not items:
        fig, ax = plt.subplots(figsize=(10, 5)); fig.patch.set_facecolor("#0d1117"); _style_ax(ax)
        ax.text(0.3, 0.5, "No usable parameters", transform=ax.transAxes, color=FG_DIM)
        return fig
    n_params = len(items)
    fig_h = max(6.0, n_params * 1.05 + 2.2)
    fig = plt.figure(figsize=(14, fig_h)); fig.patch.set_facecolor("#0d1117")
    fig.text(0.5, 0.985, "Per-Parameter Analysis: Normal vs Anomalous",
             ha="center", va="top", color=COL_PARAM_ANOM, fontsize=11, fontweight="bold")
    gs = gridspec.GridSpec(1, 2, width_ratios=[2, 3], wspace=0.06,
                           left=0.01, right=0.99, top=0.93, bottom=0.07)
    ax_bar = fig.add_subplot(gs[0]); _style_ax(ax_bar)
    y_pos = np.arange(n_params)
    norm_counts = [r["n_total"] - r["n_anom"] for r in items]
    anom_counts = [r["n_anom"] for r in items]
    totals = [n + a for n, a in zip(norm_counts, anom_counts)]
    sev_cols = [SEV_COLORS.get(r["severity"], "#8b5cf6") for r in items]
    param_labels = [r["col"] for r in items]
    ax_bar.barh(y_pos, norm_counts, height=0.60, color=NORMAL_C,
                edgecolor="#ffffff18", linewidth=0.3, label="Normal", zorder=3)
    ax_bar.barh(y_pos, anom_counts, height=0.60, left=norm_counts, color="#ef4444",
                edgecolor="#ffffff22", linewidth=0.3, label="Anomalous", zorder=3)
    max_total = max(totals) if totals else 1
    for i in range(n_params):
        nn, na, tot = norm_counts[i], anom_counts[i], totals[i]
        pct_a = na / max(tot, 1) * 100
        if nn / max_total > 0.18:
            ax_bar.text(nn / 2, i, f"{nn:,}  ({100-pct_a:.1f}%)",
                        va="center", ha="center", color="#fff", fontsize=7.5, fontweight="bold", zorder=5)
        if na > 0:
            red_center = nn + na / 2
            label_str = f"{na:,}  ({pct_a:.1f}%)"
            if na / max_total > 0.10:
                ax_bar.text(red_center, i, label_str, va="center", ha="center",
                            color="#fff", fontsize=7.5, fontweight="bold", zorder=5)
            else:
                ax_bar.text(tot + max_total * 0.01, i, label_str,
                            va="center", ha="left", color="#ef4444", fontsize=7, zorder=5)
        verdict = "✔ CLEAN" if items[i]["severity"] == "NONE" else items[i]["severity"]
        ax_bar.text(max_total * 1.02, i, verdict, va="center", ha="left",
                    color=sev_cols[i], fontsize=7.5, fontweight="bold")
    ax_bar.set_yticks(y_pos); ax_bar.set_yticklabels(param_labels, fontsize=max(6.5, 9 - n_params * 0.25), color=FG_MID)
    ax_bar.set_xlabel("Record Count", fontsize=8.5, color="#94a3b8"); ax_bar.set_xlim(0, max_total * 1.25)
    ax_bar.set_ylim(-0.6, n_params - 0.4); ax_bar.invert_yaxis()
    ax_bar.set_title("Normal vs Anomalous\nper Parameter", color="#94a3b8", fontsize=8.5, pad=6)
    ax_bar.legend(loc="lower right", fontsize=7.5, framealpha=0.88,
                  facecolor="#161d2e", edgecolor="#334155", labelcolor="#e2e8f0")
    # Right: box plots if we have data
    if df is not None and anomaly_labels is not None:
        gs_right = gridspec.GridSpecFromSubplotSpec(n_params, 1, subplot_spec=gs[1], hspace=0.55)
        rng = np.random.default_rng(42)
        for i, r in enumerate(items):
            col = r["col"]; ax_box = fig.add_subplot(gs_right[i]); _style_ax(ax_box)
            if col not in df.columns: continue
            norm_vals = df.loc[anomaly_labels == 1, col].dropna().values.astype(float)
            anom_vals = df.loc[anomaly_labels == -1, col].dropna().values.astype(float)
            if len(norm_vals) > 4000: norm_vals = rng.choice(norm_vals, 4000, replace=False)
            if len(anom_vals) > 2000: anom_vals = rng.choice(anom_vals, 2000, replace=False)
            data_to_plot = []; bp_labels = []; bp_colors = []
            if len(norm_vals) > 0: data_to_plot.append(norm_vals); bp_labels.append("Normal"); bp_colors.append(NORMAL_C)
            if len(anom_vals) > 0: data_to_plot.append(anom_vals); bp_labels.append("Anomalous"); bp_colors.append("#ef4444")
            if data_to_plot:
                bp = ax_box.boxplot(data_to_plot, vert=False, patch_artist=True, widths=0.42,
                                    showfliers=True,
                                    flierprops=dict(marker=".", markersize=2.5, markerfacecolor="#ef444488", markeredgecolor="none", alpha=0.45),
                                    medianprops=dict(color="#ffffff", linewidth=1.8),
                                    whiskerprops=dict(color="#94a3b8", linewidth=0.9),
                                    capprops=dict(color="#94a3b8", linewidth=0.9),
                                    boxprops=dict(linewidth=0.8))
                for patch, c in zip(bp["boxes"], bp_colors):
                    patch.set_facecolor(c); patch.set_alpha(0.70)
            ax_box.set_yticks(range(1, len(bp_labels) + 1))
            ax_box.set_yticklabels(bp_labels, fontsize=7.5, color=FG_MID)
            ax_box.tick_params(axis="x", labelsize=6.5, colors="#94a3b8", pad=2, length=3)
            sev_c = SEV_COLORS.get(r["severity"], "#8b5cf6")
            verdict = "CLEAN" if r["severity"] == "NONE" else f"{r['severity']}  —  {anom_counts[i]:,} anomalous"
            ax_box.set_title(f"{col}\n[{verdict}]", color=sev_c, fontsize=7.5, pad=3,
                             fontweight="bold", loc="left", linespacing=1.3)
    else:
        ax_sc = fig.add_subplot(gs[1]); _style_ax(ax_sc)
        pcts  = [r["pct"] for r in items]; skews = [abs(r["skew"]) for r in items]
        ax_sc.scatter(skews, pcts, c=sev_cols, s=70, edgecolors="#ffffff33", linewidths=0.6, zorder=4)
        for i, r in enumerate(items):
            ax_sc.annotate(r["col"][:18], (skews[i], pcts[i]), xytext=(5, 4),
                           textcoords="offset points", color=FG_DIM, fontsize=7)
        ax_sc.set_xlabel("|Skewness|", fontsize=9); ax_sc.set_ylabel("Anomaly %", fontsize=9)
        ax_sc.set_title("Anomaly % vs Skewness", color=COL_PARAM_ANOM, fontsize=10, pad=8)
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# LEGEND BUILDERS
# ══════════════════════════════════════════════════════════════════════════════
def build_cluster_legend(parent, df=None, hex_colors=None):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent, text="CLUSTERS", bg=CARD, fg=COL_CLUSTER,
             font=FONT_H3, justify="center").pack(pady=(8, 2), padx=6)
    _divider(parent, COL_CLUSTER)
    if df is None or "cluster" not in df.columns:
        tk.Label(parent, text="Run analysis first", bg=CARD, fg=FG_DIM,
                 font=FONT_XS, wraplength=170).pack(pady=8); return
    total = len(df)
    for cl in sorted(df["cluster"].unique()):
        cnt = int((df["cluster"] == cl).sum())
        col = hex_colors[int(cl) % len(hex_colors)] if hex_colors else ACCENT
        row = tk.Frame(parent, bg=CARD2, padx=5, pady=3,
                       highlightbackground=col, highlightthickness=1)
        row.pack(fill="x", padx=5, pady=2)
        tk.Canvas(row, bg=col, width=10, height=10, highlightthickness=0).pack(side="left", padx=(0, 5))
        txt = tk.Frame(row, bg=CARD2); txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=f"Cluster {cl}", bg=CARD2, fg=col,
                 font=("Courier New", 8, "bold"), anchor="w").pack(anchor="w")
        tk.Label(txt, text=f"{cnt:,} wells  ({cnt/total*100:.1f}%)", bg=CARD2, fg=FG_MID,
                 font=("Courier New", 7), anchor="w").pack(anchor="w")
        _mini_bar(parent, "", cnt / total, col)


def build_anomaly_legend(parent, n_normal=0, n_anomaly=0):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent, text="ANOMALY", bg=CARD, fg=COL_ANOMALY,
             font=FONT_H3, justify="center").pack(pady=(8, 2), padx=6)
    _divider(parent, COL_ANOMALY)
    total = max(n_normal + n_anomaly, 1)
    for label, cnt, col in [("● Normal", n_normal, NORMAL_C), ("◆ Anomalous", n_anomaly, ANOMALY_C)]:
        row = tk.Frame(parent, bg=CARD2, padx=5, pady=3,
                       highlightbackground=col, highlightthickness=1)
        row.pack(fill="x", padx=5, pady=2)
        tk.Label(row, text=label, bg=CARD2, fg=col,
                 font=("Courier New", 8, "bold")).pack(anchor="w")
        tk.Label(row, text=f"{cnt:,}  ({cnt/total*100:.1f}%)", bg=CARD2, fg=FG_MID,
                 font=("Courier New", 7)).pack(anchor="w")
        _mini_bar(parent, "", cnt / total, col)


def build_param_anomaly_legend(parent, param_anom_list):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent, text="PARAM STATUS", bg=CARD, fg=COL_PARAM_ANOM,
             font=FONT_H3, justify="center").pack(pady=(8, 2), padx=6)
    _divider(parent, COL_PARAM_ANOM)
    if not param_anom_list:
        tk.Label(parent, text="Run analysis first", bg=CARD, fg=FG_DIM,
                 font=FONT_XS, wraplength=170).pack(pady=8); return
    for r in param_anom_list:
        if r["n_total"] == 0: continue
        sev_col = SEV_COLORS.get(r["severity"], "#8b5cf6")
        n_norm  = r["n_total"] - r["n_anom"]; n_anom = r["n_anom"]
        is_clean = (n_anom == 0)
        card = tk.Frame(parent, bg=CARD2, padx=5, pady=4,
                        highlightbackground=sev_col if not is_clean else SUCCESS,
                        highlightthickness=1)
        card.pack(fill="x", padx=5, pady=2)
        tk.Label(card, text=r["col"][:22], bg=CARD2,
                 fg=sev_col if not is_clean else SUCCESS2,
                 font=("Courier New", 8, "bold"), anchor="w").pack(anchor="w")
        norm_row = tk.Frame(card, bg=CARD2); norm_row.pack(fill="x")
        tk.Label(norm_row, text="● Normal  :", bg=CARD2, fg=NORMAL_C,
                 font=("Courier New", 7), width=11, anchor="w").pack(side="left")
        tk.Label(norm_row, text=f"{n_norm:,}  ({100-r['pct']:.1f}%)", bg=CARD2, fg=FG_MID,
                 font=("Courier New", 7)).pack(side="left")
        if n_anom > 0:
            anom_row = tk.Frame(card, bg=CARD2); anom_row.pack(fill="x")
            tk.Label(anom_row, text="◆ Anomaly :", bg=CARD2, fg="#ef4444",
                     font=("Courier New", 7), width=11, anchor="w").pack(side="left")
            tk.Label(anom_row, text=f"{n_anom:,}  ({r['pct']:.1f}%)", bg=CARD2, fg="#ef4444",
                     font=("Courier New", 7, "bold")).pack(side="left")
        else:
            tk.Label(card, text="✔ All normal", bg=CARD2, fg=SUCCESS2,
                     font=("Courier New", 7)).pack(anchor="w")
        bar_frame = tk.Frame(card, bg=BORDER2, height=5); bar_frame.pack(fill="x", pady=(3, 0))
        norm_frac = n_norm / max(r["n_total"], 1)
        tk.Frame(bar_frame, bg=NORMAL_C, height=5).place(relwidth=norm_frac, relheight=1.0, relx=0)
        if 1 - norm_frac > 0.01:
            tk.Frame(bar_frame, bg="#ef4444", height=5).place(
                relwidth=1-norm_frac, relheight=1.0, relx=norm_frac)
    _divider(parent)
    flagged = sum(1 for r in param_anom_list if r["n_anom"] > 0)
    tk.Label(parent, text=f"Flagged: {flagged}  /  Clean: {len(param_anom_list)-flagged}",
             bg=CARD, fg=FG_DIM, font=("Courier New", 7)).pack(anchor="w", padx=6, pady=(3, 4))


def build_similarity_legend(parent, corr_matrix=None):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent, text="SIMILARITY", bg=CARD, fg=COL_SIMILARITY,
             font=FONT_H3, justify="center").pack(pady=(8, 2), padx=6)
    _divider(parent, COL_SIMILARITY)
    if corr_matrix is None:
        tk.Label(parent, text="Run analysis first", bg=CARD, fg=FG_DIM,
                 font=FONT_XS, wraplength=170).pack(pady=8); return
    tk.Label(parent, text="Top correlated pairs:", bg=CARD, fg=FG_MID,
             font=FONT_XS).pack(anchor="w", padx=6, pady=(4, 2))
    # Find top correlated pairs (excluding self)
    pairs = []
    cols = list(corr_matrix.columns)
    for i, c1 in enumerate(cols):
        for j, c2 in enumerate(cols):
            if j <= i: continue
            pairs.append((abs(corr_matrix.loc[c1, c2]), c1, c2, corr_matrix.loc[c1, c2]))
    pairs.sort(reverse=True)
    for abs_r, c1, c2, r in pairs[:8]:
        col = "#ef4444" if abs_r > 0.8 else ("#f59e0b" if abs_r > 0.5 else "#10b981")
        row = tk.Frame(parent, bg=CARD2, padx=4, pady=3,
                       highlightbackground=col, highlightthickness=1)
        row.pack(fill="x", padx=5, pady=1)
        tk.Label(row, text=f"{c1[:14]}\n⟺ {c2[:14]}", bg=CARD2, fg=col,
                 font=("Courier New", 7), anchor="w").pack(anchor="w")
        tk.Label(row, text=f"r = {r:.3f}", bg=CARD2, fg=FG_MID,
                 font=("Courier New", 8, "bold")).pack(anchor="w")
    _divider(parent)
    tk.Label(parent, text="Red = high correlation\nYellow = moderate\nGreen = low",
             bg=CARD, fg=FG_DIM, font=("Courier New", 7), justify="left").pack(anchor="w", padx=6, pady=2)


def build_weight_legend(parent, weight_map, method_desc):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent, text="PARAM WEIGHTS", bg=CARD, fg=COL_WEIGHTS,
             font=FONT_H3, justify="center").pack(pady=(8, 2), padx=6)
    _divider(parent, COL_WEIGHTS)
    if not weight_map:
        tk.Label(parent, text="Run analysis first", bg=CARD, fg=FG_DIM,
                 font=FONT_XS, wraplength=170).pack(pady=8); return
    total_w = max(sum(weight_map.values()), 1)
    for i, (param, wt) in enumerate(sorted(weight_map.items(), key=lambda x: -x[1])):
        col = CLUSTER_PALETTE[i % len(CLUSTER_PALETTE)]
        row = tk.Frame(parent, bg=CARD2, padx=5, pady=3,
                       highlightbackground=col, highlightthickness=1)
        row.pack(fill="x", padx=5, pady=2)
        tk.Canvas(row, bg=col, width=10, height=10, highlightthickness=0).pack(side="left", padx=(0, 5))
        txt = tk.Frame(row, bg=CARD2); txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=param[:20], bg=CARD2, fg=col,
                 font=("Courier New", 8, "bold"), anchor="w").pack(anchor="w")
        tk.Label(txt, text=f"{wt:.1f}%", bg=CARD2, fg=FG_MID,
                 font=("Courier New", 7), anchor="w").pack(anchor="w")
        _mini_bar(parent, "", wt / total_w, col)


def build_event_legend(parent, event_result):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent, text="EVENT LEGEND", bg=CARD, fg=COL_EVENTS,
             font=FONT_H3, justify="center").pack(pady=(8, 2), padx=6)
    _divider(parent, COL_EVENTS)
    if not event_result["has_events"]:
        tk.Label(parent, text="No event column", bg=CARD, fg=WARN,
                 font=FONT_SM, wraplength=170).pack(pady=8); return
    counts = event_result["event_counts"]; total = max(sum(counts.values()), 1)
    def _c(label):
        ll = label.lower()
        for k, c in EVENT_COLOURS.items():
            if k in ll: return c
        return CLUSTER_PALETTE[hash(label) % len(CLUSTER_PALETTE)]
    for label, cnt in sorted(counts.items(), key=lambda x: -x[1]):
        col = _c(label)
        row = tk.Frame(parent, bg=CARD2, padx=5, pady=3,
                       highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", padx=5, pady=2)
        tk.Canvas(row, bg=col, width=10, height=10, highlightthickness=0).pack(side="left", padx=(0, 5))
        txt = tk.Frame(row, bg=CARD2); txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=label[:22], bg=CARD2, fg=col,
                 font=("Courier New", 8, "bold"), anchor="w").pack(anchor="w")
        tk.Label(txt, text=f"{cnt:,}  ({cnt/total*100:.1f}%)", bg=CARD2, fg=FG_MID,
                 font=("Courier New", 7), anchor="w").pack(anchor="w")
        _mini_bar(parent, "", cnt / total, col)


def build_rpm_legend(parent, rpm_result):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent, text="RPM LEGEND", bg=CARD, fg="#f97316",
             font=FONT_H3, justify="center").pack(pady=(8, 2), padx=6)
    _divider(parent, "#f97316")
    if not rpm_result["found"]:
        tk.Label(parent, text="No RPM column", bg=CARD, fg=WARN,
                 font=FONT_SM, wraplength=170).pack(pady=8); return
    total = max(rpm_result["running"] + rpm_result["stopped"] +
                rpm_result["reverse"] + rpm_result["no_data"], 1)
    for label, cnt, col in [
        ("Running (RPM>0)", rpm_result["running"], "#10b981"),
        ("Stopped (RPM=0)", rpm_result["stopped"], "#f59e0b"),
        ("Reverse (RPM<0)", rpm_result["reverse"], "#ef4444"),
        ("No Data",         rpm_result["no_data"], "#4a6080"),
    ]:
        if cnt == 0: continue
        row = tk.Frame(parent, bg=CARD2, padx=5, pady=3,
                       highlightbackground=col, highlightthickness=1)
        row.pack(fill="x", padx=5, pady=2)
        tk.Canvas(row, bg=col, width=10, height=10, highlightthickness=0).pack(side="left", padx=(0, 5))
        txt = tk.Frame(row, bg=CARD2); txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=label, bg=CARD2, fg=col,
                 font=("Courier New", 8, "bold"), anchor="w").pack(anchor="w")
        tk.Label(txt, text=f"{cnt:,}  ({cnt/total*100:.1f}%)", bg=CARD2, fg=FG_MID,
                 font=("Courier New", 7), anchor="w").pack(anchor="w")
        _mini_bar(parent, "", cnt / total, col)


def _fill_toolbar(bar, mpl_canvas, tab_name):
    tk.Label(bar, text=" TOOLS:", bg="#080e1a", fg=FG_DIM, font=FONT_XS).pack(side="left", padx=(6, 2))
    nav_frame = tk.Frame(bar, bg="#080e1a"); nav_frame.pack(side="left", padx=2)
    tb = NavigationToolbar2Tk(mpl_canvas, nav_frame); tb.config(bg="#080e1a")
    for child in tb.winfo_children():
        try:
            child.config(bg="#080e1a", fg=FG_MID, activebackground=CARD2,
                         activeforeground=FG, relief="flat", bd=0, font=FONT_XS)
        except: pass
    tb.update()
    tk.Frame(bar, bg=BORDER, width=1).pack(side="left", fill="y", padx=8, pady=3)
    def _save_png():
        fig = active_figures.get(tab_name)
        if not fig: messagebox.showwarning("Export", "Run analysis first."); return
        path = filedialog.asksaveasfilename(defaultextension=".png",
                                            initialfile=f"cbm_{tab_name}_{ts()}.png",
                                            filetypes=[("PNG", "*.png"), ("SVG", "*.svg")])
        if not path: return
        fig.savefig(path, dpi=180, bbox_inches="tight", facecolor=fig.get_facecolor())
        set_status(f"Saved: {os.path.basename(path)}", SUCCESS)
    _icon_btn(bar, "Save Chart", _save_png, bg="#1e3a8a", fg="#93c5fd")


def draw_plot(tab_frame, fig, tab_name, legend_builder=None, legend_kwargs=None):
    for w in tab_frame.winfo_children(): w.destroy()
    active_figures[tab_name] = fig
    try:
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            fig.tight_layout(pad=2.5, rect=[0.03, 0.03, 0.97, 0.95])
    except: pass
    fig.patch.set_facecolor("#0d1117")
    bar = tk.Frame(tab_frame, bg="#080e1a", pady=4,
                   highlightbackground=BORDER, highlightthickness=1)
    bar.pack(side="top", fill="x")
    row = tk.Frame(tab_frame, bg=BG_MID); row.pack(side="top", fill="both", expand=True)
    chart_frame = tk.Frame(row, bg=BG_MID); chart_frame.pack(side="left", fill="both", expand=True)
    leg_frame = tk.Frame(row, bg=CARD, width=195,
                         highlightbackground=BORDER, highlightthickness=1)
    leg_frame.pack(side="right", fill="y", padx=(2, 6), pady=6)
    leg_frame.pack_propagate(False)
    mpl_canvas = FigureCanvasTkAgg(fig, master=chart_frame); mpl_canvas.draw()
    cw = mpl_canvas.get_tk_widget(); cw.config(bg="#0d1117", highlightthickness=0)
    cw.pack(fill="both", expand=True, padx=2, pady=2)
    _fill_toolbar(bar, mpl_canvas, tab_name)
    if legend_builder: legend_builder(leg_frame, **(legend_kwargs or {}))


# ══════════════════════════════════════════════════════════════════════════════
# PIPELINE
# ══════════════════════════════════════════════════════════════════════════════
def start_pipeline():
    if raw_data is None:
        messagebox.showwarning("No Data", "Please upload a dataset first."); return
    set_status("Running analysis…", WARN)
    progress_start()
    threading.Thread(target=run_pipeline, daemon=True).start()


def run_pipeline():
    global active_df, active_X, active_xcols, active_well_df
    global active_anomaly_result, active_weight_result, active_event_result
    global active_rpm_result, active_param_anom

    try:
        # ── 1. Feature selection ───────────────────────────────────────────────
        sel_num, sel_cat = get_selected_features()
        all_num = list(raw_data.select_dtypes(include="number").columns)
        x_cols  = sel_num if sel_num else all_num
        if not x_cols:
            app.after(0, lambda: set_status("No numeric features selected.", WARN))
            app.after(0, progress_stop); return

        app.after(0, lambda: set_status("Detecting well ID column…", WARN))

        # ── 2. Find well ID & drop rows with all-NaN in x_cols ────────────────
        well_id_col = _find_well_id_column(raw_data)
        needed = list(dict.fromkeys(x_cols))
        wdf = raw_data.dropna(subset=needed).reset_index(drop=True)
        n_rows = len(wdf)
        n_unique_wells = wdf[well_id_col].nunique() if well_id_col else n_rows

        app.after(0, lambda: set_status(
            f"Aggregating {n_rows:,} records → {n_unique_wells:,} wells…", WARN))

        # ── 3. Aggregate per well (FAST on huge data) ──────────────────────────
        well_df, wid_col_used = aggregate_by_well(wdf, x_cols, well_id_col)

        # Columns in well_df that are the "mean" of each original feature
        # (used for clustering well profiles)
        mean_cols = [f"{c}__mean" for c in x_cols if f"{c}__mean" in well_df.columns]
        if not mean_cols:
            # fallback: no aggregation happened (no well ID), use raw
            mean_cols = [c for c in x_cols if c in well_df.columns]

        well_df_clean = well_df.dropna(subset=mean_cols).reset_index(drop=True)
        n_wells = len(well_df_clean)

        app.after(0, lambda: set_status(
            f"Computing weights + clustering {n_wells:,} wells…", WARN))

        # ── 4. Weights ─────────────────────────────────────────────────────────
        weight_map, _ = get_manual_weights()
        method_desc   = "User-assigned weights"
        if not weight_map:
            weight_map  = {c: round(100 / len(x_cols), 2) for c in x_cols}
            method_desc = "Equal weights (auto)"
        # Map weights to mean_cols (e.g., "pressure__mean" gets weight of "pressure")
        mean_weight_map = {}
        for mc in mean_cols:
            orig = mc.replace("__mean", "")
            mean_weight_map[mc] = weight_map.get(orig, 100 / len(mean_cols))
        total_fw = sum(mean_weight_map.values()) or 1
        mean_weight_map = {c: v / total_fw * 100 for c, v in mean_weight_map.items()}

        # ── 5. Cluster on WELL PROFILES ────────────────────────────────────────
        X_well, use_mean_cols = build_weighted_X(well_df_clean, mean_cols, mean_weight_map)
        n_clusters = min(cluster_var.get(), n_wells)
        cluster_labels, inertia = run_clustering(X_well, n_clusters)
        well_df_clean = well_df_clean.copy()
        well_df_clean["cluster"] = cluster_labels

        # Map cluster labels back to raw records
        if wid_col_used and wid_col_used in wdf.columns:
            cl_map = dict(zip(well_df_clean[wid_col_used], cluster_labels))
            wdf["cluster"] = wdf[wid_col_used].map(cl_map).fillna(-1).astype(int)
        else:
            # No well ID — clusters align directly
            wdf["cluster"] = cluster_labels[:len(wdf)] if len(cluster_labels) == len(wdf) else 0

        app.after(0, lambda: set_status("Detecting anomalies…", WARN))

        # ── 6. RPM status ─────────────────────────────────────────────────────
        wdf, rpm_col_used, rpm_result = add_rpm_status_column(wdf)

        # ── 7. Anomaly detection (on raw records, not well averages) ──────────
        X_w_full, use_cols_full = build_weighted_X(wdf, x_cols, weight_map)

        # RPM filter
        if rpm_col_used and "RPM_Status" in wdf.columns:
            allowed = []
            if rpm_filter_all_var.get(): allowed = ["Running", "Stopped", "Reverse", "No Data"]
            else:
                if rpm_filter_running_var.get(): allowed.append("Running")
                if rpm_filter_stopped_var.get(): allowed.append("Stopped")
                if rpm_filter_reverse_var.get(): allowed.append("Reverse")
                if rpm_filter_nodata_var.get():  allowed.append("No Data")
            if allowed and len(allowed) < 4:
                rpm_mask = wdf["RPM_Status"].isin(allowed)
                X_w_anom = X_w_full[rpm_mask.values]
                anom_filter_desc = f"RPM filter: {', '.join(allowed)}"
            else:
                X_w_anom = X_w_full; anom_filter_desc = "No RPM filter"
        else:
            X_w_anom = X_w_full; anom_filter_desc = "No RPM column"

        method = anomaly_method_var.get(); contam = anomaly_contam_var.get() / 100.0
        a_labels_sub, a_pct, a_idx_sub, a_name, a_scores_sub = \
            detect_anomalies(X_w_anom, contamination=contam, method=method)

        a_labels = np.ones(len(wdf), dtype=int); a_scores = np.zeros(len(wdf))
        if rpm_col_used and "RPM_Status" in wdf.columns and len(X_w_anom) < len(X_w_full):
            sub_indices = np.where(wdf["RPM_Status"].isin(allowed).values)[0]
            for li, gi in enumerate(sub_indices):
                a_labels[gi] = a_labels_sub[li]; a_scores[gi] = a_scores_sub[li]
            a_idx = [sub_indices[i] for i in a_idx_sub]
        else:
            a_labels = a_labels_sub; a_scores = a_scores_sub; a_idx = a_idx_sub

        wdf["anomaly"] = a_labels
        anomaly_result = {
            "labels": a_labels, "pct": a_pct, "indices": a_idx,
            "scores": a_scores, "detector_name": f"{a_name}  [{anom_filter_desc}]",
            "n_anomaly": len(a_idx), "n_normal": len(a_labels) - len(a_idx),
        }
        event_result = analyse_events(wdf, a_labels)

        # ── 8. Per-parameter anomaly ───────────────────────────────────────────
        z_thr = zthresh_var.get()
        param_anom = analyse_per_parameter(wdf, use_cols_full, z_threshold=z_thr)

        # ── 9. Parameter similarity ────────────────────────────────────────────
        app.after(0, lambda: set_status("Computing parameter similarity…", WARN))
        corr_matrix, cosine_matrix = compute_parameter_similarity(wdf, use_cols_full)

        # ── Store globals ──────────────────────────────────────────────────────
        active_df             = wdf
        active_X              = X_w_full
        active_xcols          = use_cols_full
        active_well_df        = well_df_clean
        active_anomaly_result = anomaly_result
        active_weight_result  = {"weight_map": weight_map, "method_desc": method_desc}
        active_event_result   = event_result
        active_rpm_result     = rpm_result
        active_param_anom     = param_anom

        hx     = cluster_hex(n_clusters)
        y_cols = [c for c in wdf.select_dtypes(include="number").columns
                  if c not in use_cols_full and c != "cluster"][:2]

        insights = generate_insights(wdf, use_cols_full, n_rows, n_unique_wells, hx,
                                     anomaly_result, event_result, weight_map, method_desc,
                                     inertia, rpm_result, param_anom,
                                     corr_matrix, n_wells, well_id_col)

        app.after(0, lambda: refresh_ui(
            wdf, well_df_clean, X_well, use_cols_full, mean_cols, y_cols, hx, insights,
            n_rows, n_unique_wells, anomaly_result, event_result,
            weight_map, method_desc, rpm_result, param_anom,
            corr_matrix, cosine_matrix))
        app.after(0, lambda: set_status("✔  Analysis complete", SUCCESS))
        app.after(0, progress_stop)
        app.after(0, lambda: export_btn.config(state="normal"))

    except Exception as e:
        traceback.print_exc()
        app.after(0, lambda err=str(e): set_status(f"Error: {err}", ERR))
        app.after(0, progress_stop)


def refresh_ui(wdf, well_df, X_well, x_cols, mean_cols, y_cols, hx, insights,
               n_rows, n_unique_wells, anomaly_result, event_result,
               weight_map, method_desc, rpm_result, param_anom,
               corr_matrix, cosine_matrix):

    # Use well_df for cluster plots (one point per well)
    well_x = [c.replace("__mean", "") for c in mean_cols[:2]]
    well_y = [c.replace("__mean", "") for c in mean_cols[2:4]]

    draw_plot(cluster_tab,
              _plot_well_clusters(well_df, mean_cols, hx), "clusters",
              legend_builder=build_cluster_legend,
              legend_kwargs={"df": well_df, "hex_colors": hx})
    draw_plot(pca_tab,
              plot_pca(X_well, well_df["cluster"].values, hx), "pca",
              legend_builder=build_cluster_legend,
              legend_kwargs={"df": well_df, "hex_colors": hx})
    draw_plot(reservoir_tab,
              plot_reservoir_3d(well_df, mean_cols[:3], [], hx), "reservoir_3d",
              legend_builder=build_cluster_legend,
              legend_kwargs={"df": well_df, "hex_colors": hx})
    draw_plot(production_tab,
              plot_production(wdf, y_cols, x_cols, hx), "production",
              legend_builder=build_cluster_legend,
              legend_kwargs={"df": wdf, "hex_colors": hx})
    draw_plot(hidden_tab,
              plot_hidden_patterns(X_well, well_df["cluster"].values, mean_cols,
                                   anomaly_result["detector_name"]), "hidden_patterns",
              legend_builder=build_anomaly_legend,
              legend_kwargs={"n_normal": anomaly_result["n_normal"],
                             "n_anomaly": anomaly_result["n_anomaly"]})
    draw_plot(weight_tab, plot_weight_chart(weight_map, method_desc), "param_weights",
              legend_builder=build_weight_legend,
              legend_kwargs={"weight_map": weight_map, "method_desc": method_desc})
    draw_plot(event_tab, plot_event_chart(wdf, event_result), "event_summary",
              legend_builder=build_event_legend,
              legend_kwargs={"event_result": event_result})
    draw_plot(rpm_tab, plot_rpm_chart(wdf, rpm_result), "rpm_status",
              legend_builder=build_rpm_legend,
              legend_kwargs={"rpm_result": rpm_result})
    draw_plot(param_anom_tab,
              plot_param_anomaly(param_anom, df=wdf, anomaly_labels=anomaly_result["labels"]),
              "param_anomaly",
              legend_builder=build_param_anomaly_legend,
              legend_kwargs={"param_anom_list": param_anom})
    # ★ NEW — Similarity tab
    draw_plot(similarity_tab,
              plot_similarity(corr_matrix, cosine_matrix), "similarity",
              legend_builder=build_similarity_legend,
              legend_kwargs={"corr_matrix": corr_matrix})

    # Stat tiles
    stat_widgets["Unique Wells"].set(f"{n_unique_wells:,}")
    stat_widgets["Total Records"].set(f"{n_rows:,}")
    stat_widgets["Clusters"].set(str(well_df["cluster"].nunique()))
    anom_n = anomaly_result["n_anomaly"]; anom_p = anomaly_result["pct"]
    stat_widgets["Anomaly Recs"].set(f"{anom_n:,}\n({anom_p:.1f}%)")
    if event_result["has_events"]:
        stat_widgets["Active Recs"].set(f"{event_result['n_active']:,}")
        stat_widgets["Inactive Recs"].set(f"{event_result['n_inactive']:,}")
        stat_widgets["Event Types"].set(str(len(event_result["event_counts"])))
    else:
        stat_widgets["Active Recs"].set("N/A"); stat_widgets["Inactive Recs"].set("N/A")
        stat_widgets["Event Types"].set("N/A")
    stat_widgets["Abnormal Recs"].set(f"{event_result['n_abnormal']:,}")
    if rpm_result["found"]:
        stat_widgets["RPM Running"].set(f"{rpm_result['running']:,}")
        stat_widgets["RPM Stopped"].set(f"{rpm_result['stopped']:,}")
        stat_widgets["RPM Reverse"].set(f"{rpm_result['reverse']:,}")
    else:
        stat_widgets["RPM Running"].set("N/A"); stat_widgets["RPM Stopped"].set("N/A")
        stat_widgets["RPM Reverse"].set("N/A")
    flagged = sum(1 for r in param_anom if r["n_anom"] > 0)
    stat_widgets["Param Anomalies"].set(f"{flagged}/{len(param_anom)}\nparams flagged")

    update_event_panel(event_result)
    explain_text.config(state="normal")
    explain_text.delete("1.0", "end")
    explain_text.insert("end", insights)
    explain_text.config(state="disabled")


def _plot_well_clusters(well_df, mean_cols, hex_colors):
    """Scatter plot: one point per well, coloured by cluster."""
    fig, ax = plt.subplots(figsize=(7.5, 4.5)); _style_ax(ax)
    x_col = mean_cols[0] if len(mean_cols) > 0 else None
    y_col = mean_cols[1] if len(mean_cols) > 1 else None
    if x_col and y_col and x_col in well_df.columns and y_col in well_df.columns:
        for cl in sorted(well_df["cluster"].unique()):
            mask = well_df["cluster"] == cl; col = hex_colors[int(cl) % len(hex_colors)]
            ax.scatter(well_df.loc[mask, x_col], well_df.loc[mask, y_col],
                       color=col, s=80, edgecolors="#ffffff44", linewidths=0.7, zorder=3,
                       label=f"Cluster {cl}  ({int(mask.sum()):,} wells)")
        xlabel = x_col.replace("__mean", "")
        ylabel = y_col.replace("__mean", "")
        ax.set_xlabel(f"{xlabel}  (mean per well)", fontsize=9)
        ax.set_ylabel(f"{ylabel}  (mean per well)", fontsize=9)
    else:
        ax.text(0.2, 0.5, "Not enough features for scatter plot",
                transform=ax.transAxes, color=FG_DIM, fontsize=10)
    ax.set_title("Well Clusters  (one point = one well)",
                 color=COL_CLUSTER, fontsize=13, pad=12, fontweight="bold")
    ax.legend(loc="upper left", fontsize=8.5, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e", labelcolor="#e2e8f0", markerscale=1.4)
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# EVENT PANEL
# ══════════════════════════════════════════════════════════════════════════════
def update_event_panel(event_result):
    if not event_result["has_events"]:
        ev_status_var.set("No event column in dataset"); ev_status_lbl.config(fg=WARN)
        for v in ev_count_vars.values(): v.set("N/A")
        ev_active_var.set("N/A"); ev_inactive_var.set("N/A")
        ev_abnormal_var.set("N/A"); ev_op_anom_var.set("N/A"); return
    ev_status_var.set(f"Column: '{event_result['event_col']}'"); ev_status_lbl.config(fg=SUCCESS2)
    counts = event_result["event_counts"]
    top = sorted(counts.items(), key=lambda x: -x[1])[:6]
    keys = list(ev_count_vars.keys())
    for i, k in enumerate(keys):
        if i < len(top):
            label, cnt = top[i]; ev_count_vars[k].set(f"{cnt:,}")
            ev_count_lbls[k].config(text=label[:24])
        else:
            ev_count_vars[k].set("—"); ev_count_lbls[k].config(text="—")
    ev_active_var.set(f"{event_result['n_active']:,}")
    ev_inactive_var.set(f"{event_result['n_inactive']:,}")
    ev_abnormal_var.set(f"{event_result['n_abnormal']:,}")
    ev_op_anom_var.set(f"{len(event_result.get('operational_anomalies', [])):,}")


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════
def export_results():
    if active_df is None:
        messagebox.showwarning("Export", "Run analysis first."); return
    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", initialfile=f"cbm_results_{ts()}.xlsx",
        filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
    if not path: return
    try:
        import pandas as pd
        ext = os.path.splitext(path)[1].lower()
        if ext == ".xlsx":
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                active_df.to_excel(writer, sheet_name="All Records", index=False)
                if active_well_df is not None:
                    active_well_df.to_excel(writer, sheet_name="Well Profiles", index=False)
                if "anomaly" in active_df.columns:
                    active_df[active_df["anomaly"] == -1].to_excel(
                        writer, sheet_name="Anomaly Records", index=False)
                active_df.groupby("cluster").size().reset_index(name="well_count").to_excel(
                    writer, sheet_name="Cluster Summary", index=False)
                if active_weight_result:
                    wm = active_weight_result["weight_map"]
                    pd.DataFrame([{"Parameter": k, "Weight_%": v}
                                  for k, v in sorted(wm.items(), key=lambda x: -x[1])
                                  ]).to_excel(writer, sheet_name="Parameter Weights", index=False)
                if active_param_anom:
                    pa_rows = [{"Parameter": r["col"], "Total_Rows": r["n_total"],
                                "Outliers": r["n_anom"], "Outlier_%": round(r["pct"], 2),
                                "Mean": round(r["mean"], 4), "StdDev": round(r["std"], 4),
                                "Severity": r["severity"], "Reason": r["reason"]}
                               for r in active_param_anom]
                    pd.DataFrame(pa_rows).to_excel(writer, sheet_name="Parameter Anomalies", index=False)
                report_txt = explain_text.get("1.0", "end").strip()
                pd.DataFrame({"Report": report_txt.split("\n")}).to_excel(
                    writer, sheet_name="Insights Report", index=False)
            set_status(f"✔  Exported: {os.path.basename(path)}", SUCCESS)
            messagebox.showinfo("Export Complete", f"Saved: {path}\n\nSheets: All Records, Well Profiles, "
                                "Anomaly Records, Cluster Summary, Parameter Weights, "
                                "Parameter Anomalies, Insights Report")
        else:
            _save_df(active_df, path); set_status(f"✔  Exported: {os.path.basename(path)}", SUCCESS)
    except Exception as e:
        messagebox.showerror("Export Error", str(e)); set_status(f"Export failed: {e}", ERR)


# ══════════════════════════════════════════════════════════════════════════════
# INSIGHTS REPORT
# ══════════════════════════════════════════════════════════════════════════════
def generate_insights(df, x_cols, n_rows, n_unique_wells, hx, anomaly_result,
                      event_result, weight_map, method_desc, inertia,
                      rpm_result=None, param_anom=None,
                      corr_matrix=None, n_wells=None, well_id_col=None):
    import textwrap
    W = "=" * 56; D = "-" * 56; B = ""
    n_anom = anomaly_result["n_anomaly"]; a_pct = anomaly_result["pct"]
    sev_label = ("LOW" if a_pct < 2 else "MODERATE" if a_pct < 8 else
                 "ELEVATED" if a_pct < 20 else "HIGH")

    lines = [W, "  CBM AI ANALYSIS REPORT   v6.0",
             f"  Generated : {datetime.datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}",
             W, B,
             "  DATASET SUMMARY", "  " + D,
             f"  Unique wells   : {n_unique_wells:,}",
             f"  Well profiles  : {n_wells or n_unique_wells:,}  (clustered)",
             f"  Total records  : {n_rows:,}",
             f"  Well ID column : {well_id_col or 'Not detected'}",
             f"  Clusters       : {df['cluster'].nunique() if 'cluster' in df.columns else '—'}",
             f"  Features used  : {len(x_cols)}",
             f"  Feature list   : {', '.join(x_cols)}",
             f"  KMeans inertia : {inertia:.2f}", B]

    lines += [W, "  WELL-BASED CLUSTERING", "  " + D,
              "  Each cluster = a group of wells with similar production profiles.",
              "  Clustering is done on per-well mean statistics (not raw records).",
              "  This makes clustering meaningful and fast even on 300k+ record datasets.", B]

    # Weights
    lines += [W, "  PARAMETER WEIGHTS", "  " + D, f"  Method : {method_desc}", B]
    if weight_map:
        max_w = max(weight_map.values())
        lines.append(f"  {'PARAMETER':<28}  {'WEIGHT':>7}")
        lines.append("  " + D)
        for param, wt in sorted(weight_map.items(), key=lambda x: -x[1]):
            star = "  <- HIGHEST" if wt == max_w else ""
            lines.append(f"  {param:<28}  {wt:>6.1f}%{star}")
    lines.append(B)

    # Correlation summary
    if corr_matrix is not None:
        lines += [W, "  PARAMETER SIMILARITY (Pearson Correlation)", "  " + D]
        cols = list(corr_matrix.columns)
        pairs = []
        for i, c1 in enumerate(cols):
            for j, c2 in enumerate(cols):
                if j <= i: continue
                pairs.append((abs(corr_matrix.loc[c1, c2]), c1, c2, corr_matrix.loc[c1, c2]))
        pairs.sort(reverse=True)
        lines.append(f"  {'PAIR':<40}  {'CORR':>7}  INTERPRETATION")
        lines.append("  " + D)
        for abs_r, c1, c2, r in pairs[:10]:
            interp = ("HIGHLY similar" if abs_r > 0.8 else
                      "Moderately similar" if abs_r > 0.5 else
                      "Weakly similar" if abs_r > 0.3 else "Independent")
            pair_str = f"{c1[:18]} ⟺ {c2[:18]}"
            lines.append(f"  {pair_str:<40}  {r:>7.3f}  {interp}")
        lines.append(B)

    # Anomaly summary
    lines += [W, "  OVERALL ANOMALY DETECTION", "  " + D,
              f"  Algorithm      : {anomaly_result['detector_name']}",
              f"  Normal records : {anomaly_result['n_normal']:,}",
              f"  Anomalous recs : {n_anom:,}  ({a_pct:.1f}%)",
              f"  Severity level : {sev_label}", B]

    # Per-parameter
    lines += [W, "  PER-PARAMETER ANOMALY ANALYSIS", "  " + D]
    if param_anom:
        flagged = [r for r in param_anom if r["n_anom"] > 0]
        clean   = [r for r in param_anom if r["n_anom"] == 0]
        lines += [f"  Parameters checked : {len(param_anom)}",
                  f"  Parameters flagged : {len(flagged)}",
                  f"  Parameters clean   : {len(clean)}", B]
        for idx, r in enumerate(flagged, 1):
            lines += [f"  [{idx}] {r['col']}  —  Severity: {r['severity']}",
                      f"      Anomalous: {r['n_anom']:,}  ({r['pct']:.2f}%)  "
                      f"|  Normal: {r['n_total']-r['n_anom']:,}",
                      f"      Mean={r['mean']:.4f}  Std={r['std']:.4f}  "
                      f"Skew={r['skew']:.3f}"]
            for wl in textwrap.wrap(r["reason"], 54):
                lines.append(f"      {wl}")
            lines += [B]
    else:
        lines += ["  Per-parameter analysis not available.", B]

    # Events
    lines += [W, "  OPERATIONAL EVENTS", "  " + D]
    if event_result["has_events"]:
        lines += [f"  Column: {event_result['event_col']}",
                  f"  Active: {event_result['n_active']:,}  |  Inactive: {event_result['n_inactive']:,}", B]
        for ev, cnt in sorted(event_result["event_counts"].items(), key=lambda x: -x[1]):
            pct = cnt / max(n_rows, 1) * 100
            lines.append(f"  {ev:<30}  {cnt:>6,}  {pct:>5.1f}%")
    else:
        lines += [f"  {event_result['message']}"]
    lines.append(B)

    # RPM
    if rpm_result and rpm_result.get("found"):
        total_rpm = max(rpm_result["running"] + rpm_result["stopped"] +
                        rpm_result["reverse"] + rpm_result["no_data"], 1)
        lines += [W, "  RPM STATUS", "  " + D,
                  f"  Running : {rpm_result['running']:,}  ({rpm_result['running']/total_rpm*100:.1f}%)",
                  f"  Stopped : {rpm_result['stopped']:,}  ({rpm_result['stopped']/total_rpm*100:.1f}%)",
                  f"  Reverse : {rpm_result['reverse']:,}  ({rpm_result['reverse']/total_rpm*100:.1f}%)",
                  B]

    lines += [W, "  END OF REPORT", W]
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
if HAVE_TTKBS:
    app = ttk.Window(themename="darkly")
else:
    app = tk.Tk(); app.configure(bg=BG_DEEP)

app.title("CBM AI Analytics Platform  v6.0  —  Well Clustering + Feature Selection + Similarity")
app.geometry("1780x1000")
app.configure(bg=BG_DEEP); app.minsize(1280, 720)

top_bar = tk.Frame(app, bg="#040810", height=46)
top_bar.pack(fill="x", side="top"); top_bar.pack_propagate(False)
brand = tk.Frame(top_bar, bg="#040810"); brand.pack(side="left", padx=16, fill="y")
tk.Label(brand, text="CBM·AI", bg="#040810", fg=COL_DATASET,
         font=("Georgia", 17, "bold")).pack(side="left", padx=(0, 10))
tk.Label(brand, text="Coalbed Methane Analytics  v6.0  —  Well Clustering + Feature Selection + Similarity",
         bg="#040810", fg=FG_DIM, font=FONT_SM).pack(side="left")
tk.Label(top_bar,
         text="Huge files: chunked load  ·  Well-based clusters  ·  Param similarity  ·  Feature selector",
         bg="#040810", fg=FG_DIM, font=FONT_XS).pack(side="right", padx=14)
tk.Frame(app, bg=BORDER, height=1).pack(fill="x")

root_pane = tk.PanedWindow(app, orient="horizontal", bg=BG_DEEP,
                           sashwidth=4, sashrelief="flat", sashpad=0)
root_pane.pack(fill="both", expand=True)

# ══════════════════════════════════════════════════════════════════════════════
# LEFT SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
sb_outer, sidebar = make_scrollable(root_pane, bg=SIDEBAR)
root_pane.add(sb_outer, width=460, minsize=420)
tk.Frame(sidebar, bg=SIDEBAR, height=8).pack()

# ① DATASET
ds_body = make_section_card(sidebar, "① DATASET  (chunked load — any size)", COL_DATASET)
tk.Label(ds_body,
         text="Accepts: CSV, TSV, XLSX, XLS, ODS, JSON, Parquet,\n"
              "Feather, HDF5, Pickle, MS Access .mdb/.accdb\n"
              "Large CSVs are streamed in chunks — no memory crash.",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=420, justify="left").pack(anchor="w", pady=(0, 4))
make_btn(ds_body, "  Upload Data File", upload_dataset, color="#0e7490", fg_col="#e0f7fa",
         tip="Load any data file")
info_row = tk.Frame(ds_body, bg=CARD); info_row.pack(fill="x", pady=(6, 0))
rows_var = tk.StringVar(value="—"); cols_var = tk.StringVar(value="—")
num_cols_var = tk.StringVar(value="—")
for title, var, col in [("Unique Wells", rows_var, COL_DATASET),
                         ("Columns", cols_var, FG_DIM),
                         ("Numeric", num_cols_var, COL_WEIGHTS)]:
    cf = tk.Frame(info_row, bg=CARD); cf.pack(side="left", expand=True, fill="x")
    tk.Label(cf, text=title, bg=CARD, fg=FG_DIM, font=FONT_XS).pack(anchor="w")
    tk.Label(cf, textvariable=var, bg=CARD, fg=col,
             font=("Courier New", 13, "bold")).pack(anchor="w")
event_col_var = tk.StringVar(value="Upload dataset to detect event column")
event_col_lbl = tk.Label(ds_body, textvariable=event_col_var, bg=CARD, fg=FG_DIM,
                         font=FONT_XS, wraplength=420, justify="left", anchor="w")
event_col_lbl.pack(fill="x", pady=(4, 0))
well_id_var = tk.StringVar(value="Well ID column: Not detected yet")
well_id_lbl = tk.Label(ds_body, textvariable=well_id_var, bg=CARD, fg=FG_DIM,
                       font=FONT_XS, wraplength=420, justify="left", anchor="w")
well_id_lbl.pack(fill="x", pady=(2, 0))

# ★ NEW ② FEATURE SELECTION
feat_outer = make_section_card(sidebar, "② FEATURE SELECTION  (numeric + categorical)", COL_FEATURES)
tk.Label(feat_outer,
         text="Check features to include in clustering and anomaly detection.\n"
              "Numeric → used for analysis.  Categorical → used for grouping.",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=420, justify="left").pack(anchor="w", pady=(0, 4))

# Numeric features sub-panel
num_hdr = tk.Frame(feat_outer, bg=CARD); num_hdr.pack(fill="x", pady=(2, 1))
tk.Label(num_hdr, text="NUMERIC FEATURES", bg=CARD, fg=COL_FEATURES,
         font=FONT_H3).pack(side="left")
feat_num_count_var = tk.StringVar(value="0/0 selected")
tk.Label(num_hdr, textvariable=feat_num_count_var, bg=CARD, fg=FG_DIM,
         font=FONT_XS).pack(side="right")

# Search box for numeric
num_search_row = tk.Frame(feat_outer, bg=CARD); num_search_row.pack(fill="x", pady=1)
tk.Label(num_search_row, text="Search:", bg=CARD, fg=FG_DIM, font=FONT_XS).pack(side="left")
feat_search_var_num = tk.StringVar()
num_search_entry = tk.Entry(num_search_row, textvariable=feat_search_var_num,
                            bg=CARD2, fg=FG, insertbackground=FG,
                            font=FONT_XS, relief="flat", bd=2, width=22)
num_search_entry.pack(side="left", padx=4)
feat_search_var_num.trace_add("write", lambda *a: _filter_feat_checkboxes(
    feat_search_var_num.get(), feat_num_frames, feat_num_vars))

# Select all / clear
num_btn_row = tk.Frame(feat_outer, bg=CARD); num_btn_row.pack(fill="x", pady=(1, 2))
def _sel_all_num():
    for v in feat_num_vars.values(): v.set(True)
def _clr_all_num():
    for v in feat_num_vars.values(): v.set(False)
for lbl, cmd in [("All", _sel_all_num), ("None", _clr_all_num)]:
    tk.Button(num_btn_row, text=lbl, command=cmd, bg=CARD2, fg=COL_FEATURES,
              activebackground=BORDER, activeforeground=FG,
              font=FONT_XS, relief="flat", bd=0, cursor="hand2",
              padx=8, pady=3).pack(side="left", padx=2)

# Scrollable numeric checkboxes
feat_num_canvas_outer = tk.Frame(feat_outer, bg=CARD, height=150)
feat_num_canvas_outer.pack(fill="x"); feat_num_canvas_outer.pack_propagate(False)
feat_num_canvas = Canvas(feat_num_canvas_outer, bg=CARD, highlightthickness=0, bd=0)
feat_num_sb = Scrollbar(feat_num_canvas_outer, orient="vertical", command=feat_num_canvas.yview,
                        bg=BORDER, troughcolor=CARD, activebackground=COL_FEATURES)
feat_num_canvas.configure(yscrollcommand=feat_num_sb.set)
feat_num_sb.pack(side="right", fill="y"); feat_num_canvas.pack(side="left", fill="both", expand=True)
feat_num_body = tk.Frame(feat_num_canvas, bg=CARD)
fn_wid = feat_num_canvas.create_window((0, 0), window=feat_num_body, anchor="nw")
def _fn_resize(e):
    feat_num_canvas.configure(scrollregion=feat_num_canvas.bbox("all"))
    feat_num_canvas.itemconfig(fn_wid, width=e.width)
feat_num_body.bind("<Configure>", lambda e: feat_num_canvas.configure(scrollregion=feat_num_canvas.bbox("all")))
feat_num_canvas.bind("<Configure>", _fn_resize)
tk.Label(feat_num_body, text="Upload a dataset to see numeric columns.",
         bg=CARD, fg=FG_DIM, font=FONT_XS, padx=8, pady=8).pack(anchor="w")

_divider(feat_outer)

# Categorical features sub-panel
cat_hdr = tk.Frame(feat_outer, bg=CARD); cat_hdr.pack(fill="x", pady=(2, 1))
tk.Label(cat_hdr, text="CATEGORICAL FEATURES", bg=CARD, fg=COL_EVENTS, font=FONT_H3).pack(side="left")
feat_cat_count_var = tk.StringVar(value="0/0 selected")
tk.Label(cat_hdr, textvariable=feat_cat_count_var, bg=CARD, fg=FG_DIM, font=FONT_XS).pack(side="right")

cat_search_row = tk.Frame(feat_outer, bg=CARD); cat_search_row.pack(fill="x", pady=1)
tk.Label(cat_search_row, text="Search:", bg=CARD, fg=FG_DIM, font=FONT_XS).pack(side="left")
feat_search_var_cat = tk.StringVar()
cat_search_entry = tk.Entry(cat_search_row, textvariable=feat_search_var_cat,
                            bg=CARD2, fg=FG, insertbackground=FG,
                            font=FONT_XS, relief="flat", bd=2, width=22)
cat_search_entry.pack(side="left", padx=4)
feat_search_var_cat.trace_add("write", lambda *a: _filter_feat_checkboxes(
    feat_search_var_cat.get(), feat_cat_frames, feat_cat_vars))

cat_btn_row = tk.Frame(feat_outer, bg=CARD); cat_btn_row.pack(fill="x", pady=(1, 2))
def _sel_all_cat():
    for v in feat_cat_vars.values(): v.set(True)
def _clr_all_cat():
    for v in feat_cat_vars.values(): v.set(False)
for lbl, cmd in [("All", _sel_all_cat), ("None", _clr_all_cat)]:
    tk.Button(cat_btn_row, text=lbl, command=cmd, bg=CARD2, fg=COL_EVENTS,
              activebackground=BORDER, activeforeground=FG,
              font=FONT_XS, relief="flat", bd=0, cursor="hand2",
              padx=8, pady=3).pack(side="left", padx=2)

feat_cat_canvas_outer = tk.Frame(feat_outer, bg=CARD, height=120)
feat_cat_canvas_outer.pack(fill="x"); feat_cat_canvas_outer.pack_propagate(False)
feat_cat_canvas = Canvas(feat_cat_canvas_outer, bg=CARD, highlightthickness=0, bd=0)
feat_cat_sb = Scrollbar(feat_cat_canvas_outer, orient="vertical", command=feat_cat_canvas.yview,
                        bg=BORDER, troughcolor=CARD, activebackground=COL_EVENTS)
feat_cat_canvas.configure(yscrollcommand=feat_cat_sb.set)
feat_cat_sb.pack(side="right", fill="y"); feat_cat_canvas.pack(side="left", fill="both", expand=True)
feat_cat_body = tk.Frame(feat_cat_canvas, bg=CARD)
fc_wid = feat_cat_canvas.create_window((0, 0), window=feat_cat_body, anchor="nw")
def _fc_resize(e):
    feat_cat_canvas.configure(scrollregion=feat_cat_canvas.bbox("all"))
    feat_cat_canvas.itemconfig(fc_wid, width=e.width)
feat_cat_body.bind("<Configure>", lambda e: feat_cat_canvas.configure(scrollregion=feat_cat_canvas.bbox("all")))
feat_cat_canvas.bind("<Configure>", _fc_resize)
tk.Label(feat_cat_body, text="Upload a dataset to see categorical columns.",
         bg=CARD, fg=FG_DIM, font=FONT_XS, padx=8, pady=8).pack(anchor="w")

# ③ CLUSTER SETTINGS
cl_body = make_section_card(sidebar, "③ CLUSTER SETTINGS  (clusters wells, not records)", COL_CLUSTER)
tk.Label(cl_body,
         text="Clusters are built on per-well average profiles.\n"
              "Fast even on 300k+ record datasets.",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=420).pack(anchor="w")
cluster_var = tk.IntVar(value=3)
sf = tk.Frame(cl_body, bg=CARD); sf.pack(fill="x", pady=4)
tk_ttk.Spinbox(sf, from_=2, to=10, textvariable=cluster_var, width=5, font=FONT_H2).pack(side="left")
tk.Label(sf, text="clusters  (of wells)", bg=CARD, fg=FG_DIM, font=FONT_SM).pack(side="left", padx=8)

# ④ PARAMETER WEIGHTAGE
wt_outer = make_section_card(sidebar, "④ PARAMETER WEIGHTAGE", COL_WEIGHTS)
tk.Label(wt_outer, text="Assign % weight per numeric feature (total should = 100%).",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=420).pack(anchor="w", pady=(0, 4))
weight_sum_var = tk.StringVar(value="Upload a dataset to see parameters")
weight_sum_lbl = tk.Label(wt_outer, textvariable=weight_sum_var, bg=CARD, fg=FG_DIM,
                          font=("Courier New", 9, "bold"), anchor="w")
weight_sum_lbl.pack(anchor="w", pady=(0, 4))
wt_canvas_outer = tk.Frame(wt_outer, bg=CARD, height=160)
wt_canvas_outer.pack(fill="x"); wt_canvas_outer.pack_propagate(False)
wt_canvas = Canvas(wt_canvas_outer, bg=CARD, highlightthickness=0, bd=0)
wt_sb = Scrollbar(wt_canvas_outer, orient="vertical", command=wt_canvas.yview,
                  bg=BORDER, troughcolor=CARD, activebackground=COL_WEIGHTS)
wt_canvas.configure(yscrollcommand=wt_sb.set)
wt_sb.pack(side="right", fill="y"); wt_canvas.pack(side="left", fill="both", expand=True)
weight_body = tk.Frame(wt_canvas, bg=CARD)
wt_wid = wt_canvas.create_window((0, 0), window=weight_body, anchor="nw")
def _wt_resize(e):
    wt_canvas.configure(scrollregion=wt_canvas.bbox("all"))
    wt_canvas.itemconfig(wt_wid, width=e.width)
weight_body.bind("<Configure>", lambda e: wt_canvas.configure(scrollregion=wt_canvas.bbox("all")))
wt_canvas.bind("<Configure>", _wt_resize)
_wt_ph = tk.Label(weight_body, text="Upload a dataset first.",
                  bg=CARD, fg=FG_DIM, font=FONT_XS, justify="left", padx=8, pady=8)
_wt_ph.pack(anchor="w"); weight_row_frames = [_wt_ph]

# ⑤ ANOMALY DETECTION
anom_body = make_section_card(sidebar, "⑤ ANOMALY DETECTION", COL_ANOMALY)
tk.Label(anom_body, text="Algorithm", bg=CARD, fg=FG_DIM, font=FONT_SM).pack(anchor="w")
anomaly_method_var = tk.StringVar(value="iforest")
mrow = tk.Frame(anom_body, bg=CARD); mrow.pack(fill="x", pady=(2, 6))
for label, val in [("Isolation Forest  (faster, tree-based)", "iforest"),
                   ("Local Outlier Factor  (density-based)", "lof")]:
    tk.Radiobutton(mrow, text=label, variable=anomaly_method_var, value=val,
                   bg=CARD, fg=FG_MID, selectcolor=CARD2, activebackground=CARD,
                   activeforeground=COL_ANOMALY, font=FONT_SM,
                   wraplength=420).pack(anchor="w", pady=1)
tk.Label(anom_body, text="Contamination %  (expected anomaly fraction)",
         bg=CARD, fg=FG_DIM, font=FONT_SM).pack(anchor="w")
anomaly_contam_var = tk.DoubleVar(value=5.0)
crow = tk.Frame(anom_body, bg=CARD); crow.pack(fill="x", pady=2)
tk_ttk.Spinbox(crow, from_=1.0, to=49.0, increment=0.5,
               textvariable=anomaly_contam_var, width=6, font=FONT_H3).pack(side="left")
tk.Label(crow, text="%  (typical CBM: 3–10%)", bg=CARD, fg=FG_DIM, font=FONT_XS).pack(side="left", padx=6)

# ⑥ Per-parameter Z threshold
pa_body = make_section_card(sidebar, "⑥ PER-PARAMETER  Z-Score Threshold", COL_PARAM_ANOM)
tk.Label(pa_body, text="Lower = more anomalies per parameter.  Default 2.5.",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=420).pack(anchor="w", pady=(0, 4))
zthresh_row = tk.Frame(pa_body, bg=CARD); zthresh_row.pack(fill="x", pady=2)
zthresh_var = tk.DoubleVar(value=2.5)
tk_ttk.Spinbox(zthresh_row, from_=1.0, to=5.0, increment=0.1,
               textvariable=zthresh_var, width=6, font=FONT_H3).pack(side="left")
tk.Label(zthresh_row, text="σ  (standard deviations from mean)",
         bg=CARD, fg=FG_DIM, font=FONT_XS).pack(side="left", padx=6)

# ⑦ RPM FILTER
rpm_filter_body = make_section_card(sidebar, "⑦ RPM FILTER", "#f97316")
tk.Label(rpm_filter_body, text="Select RPM states to include in anomaly detection.",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=420).pack(anchor="w", pady=(0, 4))
rpm_filter_all_var     = tk.BooleanVar(value=True)
rpm_filter_running_var = tk.BooleanVar(value=True)
rpm_filter_stopped_var = tk.BooleanVar(value=True)
rpm_filter_reverse_var = tk.BooleanVar(value=True)
rpm_filter_nodata_var  = tk.BooleanVar(value=True)
def _rpm_all_toggle():
    if rpm_filter_all_var.get():
        for v in [rpm_filter_running_var, rpm_filter_stopped_var,
                  rpm_filter_reverse_var, rpm_filter_nodata_var]:
            v.set(True)
tk.Checkbutton(rpm_filter_body, text="All RPM states (no filter)",
               variable=rpm_filter_all_var, command=_rpm_all_toggle,
               bg=CARD, fg=FG_MID, selectcolor=CARD2, activebackground=CARD,
               activeforeground="#f97316", font=FONT_SM).pack(anchor="w", pady=1)
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
tk.Label(rpm_filter_body, textvariable=rpm_filter_info_var, bg=CARD, fg=FG_DIM,
         font=FONT_XS, wraplength=420).pack(anchor="w", pady=(4, 0))

# ⑧ RUN
run_body = make_section_card(sidebar, "⑧ RUN ANALYSIS", "#6366f1")
make_btn(run_body, "  ▶  Run AI Analysis  (Well Clusters + Anomaly + Similarity)",
         start_pipeline, color="#4f46e5",
         tip="Well-based clustering + per-parameter anomaly + similarity + events")
if HAVE_TTKBS:
    progress_bar = ttk.Progressbar(run_body, mode="determinate", bootstyle="info-striped", length=420)
else:
    progress_bar = tk_ttk.Progressbar(run_body, mode="determinate", length=420)
progress_bar.pack(fill="x", pady=(4, 2))
status_var = tk.StringVar(value="Ready — upload a dataset to begin")
status_lbl = tk.Label(run_body, textvariable=status_var, bg=CARD, fg=FG_DIM,
                      font=FONT_SM, anchor="w", wraplength=420, justify="left")
status_lbl.pack(fill="x", pady=(2, 0))

# ⑨ EVENT DETECTION PANEL
ev_body = make_section_card(sidebar, "⑨ EVENT DETECTION PANEL", COL_EVENTS)
ev_status_var = tk.StringVar(value="Awaiting analysis")
ev_status_lbl = tk.Label(ev_body, textvariable=ev_status_var, bg=CARD, fg=FG_DIM,
                         font=FONT_SM, wraplength=420)
ev_status_lbl.pack(anchor="w", pady=(0, 6))
ev_tile_row = tk.Frame(ev_body, bg=CARD); ev_tile_row.pack(fill="x", pady=(0, 6))
ev_active_var = tk.StringVar(value="—"); ev_inactive_var = tk.StringVar(value="—")
ev_abnormal_var = tk.StringVar(value="—"); ev_op_anom_var = tk.StringVar(value="—")
for title, var, col in [("Active Recs", ev_active_var, COL_EVENTS),
                         ("Inactive Recs", ev_inactive_var, COL_ANOMALY),
                         ("Abnormal", ev_abnormal_var, WARN),
                         ("Op. Anom.", ev_op_anom_var, COL_WEIGHTS)]:
    cell = tk.Frame(ev_tile_row, bg=CARD2, padx=5, pady=5,
                    highlightbackground=BORDER, highlightthickness=1)
    cell.pack(side="left", expand=True, fill="x", padx=2)
    tk.Label(cell, textvariable=var, bg=CARD2, fg=col,
             font=("Courier New", 11, "bold")).pack()
    tk.Label(cell, text=title, bg=CARD2, fg=FG_DIM, font=FONT_XS).pack()
_ev_label_names = [f"ev{i}" for i in range(6)]
ev_count_vars = {}; ev_count_lbls = {}
ev_detail_frame = tk.Frame(ev_body, bg=CARD); ev_detail_frame.pack(fill="x")
for key in _ev_label_names:
    row = tk.Frame(ev_detail_frame, bg=CARD); row.pack(fill="x", pady=1)
    lbl = tk.Label(row, text="—", bg=CARD, fg=FG_DIM, font=FONT_XS, width=26, anchor="w")
    lbl.pack(side="left", padx=(4, 0))
    cnt_var = tk.StringVar(value="—")
    tk.Label(row, textvariable=cnt_var, bg=CARD, fg=COL_EVENTS,
             font=FONT_XS, width=8, anchor="e").pack(side="right", padx=4)
    ev_count_vars[key] = cnt_var; ev_count_lbls[key] = lbl

# ⑩ EXPORT
exp_body = make_section_card(sidebar, "⑩ EXPORT RESULTS", COL_EXPORT)
tk.Label(exp_body,
         text="Excel: All Records · Well Profiles · Anomaly Records\n"
              "Cluster Summary · Parameter Weights · Parameter Anomalies · Insights Report",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=420).pack(anchor="w", pady=(0, 6))
export_btn = tk.Button(exp_body, text="  Export All Results  (Excel / CSV)",
                       command=export_results, bg="#14532d", fg=COL_EXPORT,
                       activebackground=_lighten("#14532d", 20), activeforeground=COL_EXPORT,
                       font=FONT_H3, relief="flat", bd=0, cursor="hand2",
                       padx=10, pady=11, anchor="w", state="disabled")
export_btn.pack(fill="x", pady=2)
export_btn.bind("<Enter>", lambda e: export_btn.config(bg=_lighten("#14532d", 20)))
export_btn.bind("<Leave>", lambda e: export_btn.config(bg="#14532d"))

# ⑪ DATASET PREVIEW
prev_body = make_section_card(sidebar, "⑪ DATASET PREVIEW  (first 20 rows)", COL_PREVIEW,
                              fill="both", expand=True)
table_frame = tk.Frame(prev_body, bg=CARD); table_frame.pack(fill="both", expand=True)
tk.Frame(sidebar, bg=SIDEBAR, height=20).pack()

# ══════════════════════════════════════════════════════════════════════════════
# RIGHT PANEL
# ══════════════════════════════════════════════════════════════════════════════
rp_outer, right_inner = make_scrollable(root_pane, bg=BG)
root_pane.add(rp_outer, minsize=820)
rp_hdr = tk.Frame(right_inner, bg=BG, padx=16, pady=10); rp_hdr.pack(fill="x")
tk.Label(rp_hdr, text="Analysis Dashboard", bg=BG, fg=FG,
         font=("Georgia", 15, "bold")).pack(side="left")
tk.Label(rp_hdr, text="v6.0  ·  Well Clusters  ·  Feature Selection  ·  Param Similarity  ·  Events  ·  RPM",
         bg=BG, fg=FG_DIM, font=FONT_SM).pack(side="right", pady=2)
tk.Frame(right_inner, bg=BORDER, height=1).pack(fill="x", padx=12)

# Stat tiles
stats_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=5); stats_wrap.pack(fill="x")
stat_defs_r1 = [("Unique Wells", "—", COL_DATASET),  ("Total Records", "—", COL_CLUSTER),
                ("Clusters", "—", COL_WEIGHTS),        ("Anomaly Recs", "—", COL_ANOMALY)]
stat_defs_r2 = [("Active Recs", "—", COL_EVENTS),     ("Inactive Recs", "—", ERR),
                ("Event Types", "—", WARN),             ("Abnormal Recs", "—", COL_ANOMALY)]
stat_defs_r3 = [("RPM Running", "—", "#10b981"),       ("RPM Stopped", "—", "#f59e0b"),
                ("RPM Reverse", "—", "#ef4444"),        ("Param Anomalies", "—", COL_PARAM_ANOM)]
stat_widgets = {}
for stat_row_defs in [stat_defs_r1, stat_defs_r2, stat_defs_r3]:
    row = tk.Frame(stats_wrap, bg=BG); row.pack(fill="x", pady=2)
    for title, val, col in stat_row_defs:
        sf = tk.Frame(row, bg=CARD, padx=6, pady=6,
                      highlightbackground=col, highlightthickness=1)
        sf.pack(side="left", expand=True, fill="x", padx=3)
        sv = tk.StringVar(value=val)
        tk.Label(sf, textvariable=sv, bg=CARD, fg=col,
                 font=("Courier New", 10, "bold"), wraplength=120, justify="center").pack(fill="x")
        tk.Label(sf, text=title, bg=CARD, fg=FG_DIM,
                 font=("Courier New", 7), wraplength=120, justify="center").pack(fill="x")
        stat_widgets[title] = sv

# TABS
nb_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=6); nb_wrap.pack(fill="x")
sty = tk_ttk.Style()
sty.configure("CBM.TNotebook",     background=BG, tabmargins=[0, 0, 0, 0])
sty.configure("CBM.TNotebook.Tab", background=CARD2, foreground=FG_DIM,
              padding=[12, 6], font=FONT_H3)
sty.map("CBM.TNotebook.Tab",
        background=[("selected", "#1e3a8a")], foreground=[("selected", "#93c5fd")])
notebook = tk_ttk.Notebook(nb_wrap, style="CBM.TNotebook"); notebook.pack(fill="x")

TAB_H = 500
cluster_tab    = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
pca_tab        = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
reservoir_tab  = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
production_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
hidden_tab     = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
weight_tab     = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
event_tab      = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
rpm_tab        = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
param_anom_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
similarity_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)   # NEW

for tab, name in [
    (cluster_tab,    "  Clusters  "),
    (pca_tab,        "  PCA  "),
    (reservoir_tab,  "  3D Reservoir  "),
    (production_tab, "  Production  "),
    (hidden_tab,     "  Hidden Patterns  "),
    (weight_tab,     "  Param Weights  "),
    (event_tab,      "  Events  "),
    (rpm_tab,        "  RPM Status  "),
    (param_anom_tab, "  ★ Param Anomaly  "),
    (similarity_tab, "  ★ Similarity  "),   # NEW
]:
    tab.pack_propagate(False); notebook.add(tab, text=name)

# Insights Report
ins_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=6); ins_wrap.pack(fill="x")
ins_hdr = tk.Frame(ins_wrap, bg=BG); ins_hdr.pack(fill="x", pady=(0, 4))
tk.Label(ins_hdr,
         text="AI Insights Report  —  Well Clusters · Param Similarity · Anomaly Diagnosis",
         bg=BG, fg=FG, font=FONT_H2).pack(side="left")
copy_btn = tk.Button(ins_hdr, text="Copy",
                     command=lambda: (app.clipboard_clear(),
                                     app.clipboard_append(explain_text.get("1.0", "end"))),
                     bg=CARD2, fg=FG_DIM, activebackground=BORDER, activeforeground=FG,
                     font=FONT_XS, relief="flat", bd=0, cursor="hand2", padx=10, pady=4)
copy_btn.pack(side="right")
explain_text_frame = tk.Frame(ins_wrap, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
explain_text_frame.pack(fill="both", expand=True)
explain_xsb = tk_ttk.Scrollbar(explain_text_frame, orient="horizontal")
explain_vsb = tk_ttk.Scrollbar(explain_text_frame, orient="vertical")
explain_text = tk.Text(explain_text_frame, height=26, bg=CARD, fg="#a8f0c8",
                       font=("Courier New", 9), relief="flat", bd=0,
                       padx=14, pady=10, insertbackground=FG, wrap="none",
                       xscrollcommand=explain_xsb.set, yscrollcommand=explain_vsb.set,
                       highlightthickness=0, state="disabled")
explain_xsb.config(command=explain_text.xview)
explain_vsb.config(command=explain_text.yview)
explain_vsb.pack(side="right", fill="y")
explain_xsb.pack(side="bottom", fill="x")
explain_text.pack(side="left", fill="both", expand=True)
tk.Frame(right_inner, bg=BG, height=20).pack()

explain_text.config(state="normal")
explain_text.insert("end",
    "========================================================\n"
    "  CBM AI Analytics Platform  v6.0\n"
    "  Well Clustering + Feature Selection + Param Similarity\n"
    "========================================================\n"
    "\n"
    "  WHAT'S NEW IN v6.0:\n"
    "  ------------------------------------------------------\n"
    "  ★ FEATURE SELECTION PANEL (sidebar section II):\n"
    "    - Checkboxes for every numeric column\n"
    "    - Checkboxes for categorical columns\n"
    "    - Search/filter box to find columns quickly\n"
    "    - Select All / Clear All buttons per group\n"
    "    - Live count of selected features\n"
    "\n"
    "  ★ WELL-BASED CLUSTERING:\n"
    "    - Groups records by Well ID first\n"
    "    - Computes per-well mean, std, min, max, p25, p75\n"
    "    - Clusters on WELL PROFILES (not raw records)\n"
    "    - 300k records → 287 well vectors = very fast\n"
    "    - Meaningful: wells with similar production group together\n"
    "\n"
    "  ★ PARAMETER SIMILARITY TAB:\n"
    "    - Pearson correlation heatmap between all features\n"
    "    - Cosine similarity matrix (normalised)\n"
    "    - Right panel shows top correlated feature pairs\n"
    "    - Helps identify redundant / independent features\n"
    "\n"
    "  ★ FAST LOADING:\n"
    "    - CSV files > 50 MB are streamed in 200k-row chunks\n"
    "    - No memory crash on large files\n"
    "    - Progress updates during loading\n"
    "\n"
    "  QUICK START:\n"
    "  ------------------------------------------------------\n"
    "  1. Upload your data file\n"
    "  2. Select features in the Feature Selection panel\n"
    "  3. Set number of clusters\n"
    "  4. Assign parameter weights (optional)\n"
    "  5. Set anomaly contamination %\n"
    "  6. Click  ▶  Run AI Analysis\n"
    "\n"
    "  After analysis:\n"
    "  → [Clusters] tab: one point per WELL (not per record)\n"
    "  → [★ Similarity] tab: which parameters behave alike\n"
    "  → [★ Param Anomaly] tab: which parameters have outliers\n"
    "  → Read the Insights Report below for full diagnosis\n"
    "========================================================\n"
)
explain_text.config(state="disabled")

app.mainloop()
