"""
CBM AI Analytics Platform — v7.0  (Floating Chatbot Edition)
─────────────────────────────────────────────────────────────
CHANGES FROM v7.0 base:
• CHATBOT moved from embedded panel to a FLOATING ICON BUTTON
  - Small pulsing ★ AI icon pinned to bottom-right of the window
  - Click to open/close a compact 380×520px popup chat panel
  - Panel appears above the icon, stays on top of all other widgets
  - Drag the popup by its title bar to reposition
  - ESC or the × button closes the panel
  - All chat functionality identical to the embedded version
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
import json
import urllib.request
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
# FILE LOADER
# ══════════════════════════════════════════════════════════════════════════════
CHUNK_ROWS = 200_000

def _get_access_tables_pyodbc(filepath):
    import pyodbc
    drivers = [d for d in pyodbc.drivers() if 'access' in d.lower() or 'mdb' in d.lower()]
    if not drivers:
        raise RuntimeError("No Microsoft Access ODBC driver found.")
    driver = drivers[0]
    conn = pyodbc.connect(f"Driver={{{driver}}};Dbq={filepath};Uid=Admin;Pwd=;", timeout=30)
    cursor = conn.cursor()
    tables = [r.table_name for r in cursor.tables(tableType='TABLE')
              if not r.table_name.startswith('MSys')]
    conn.close()
    return tables, driver

def _load_access_table_pyodbc(filepath, table_name, driver):
    import pyodbc, pandas as pd
    conn = pyodbc.connect(f"Driver={{{driver}}};Dbq={filepath};Uid=Admin;Pwd=;", timeout=120)
    try:
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM [{table_name}]")
        columns = [col[0] for col in cursor.description]
        df = pd.DataFrame.from_records(cursor.fetchall(), columns=columns)
        cursor.close()
    except Exception as e:
        conn.close(); raise RuntimeError(f"Could not read table '{table_name}': {e}")
    conn.close()
    return df

def _pick_table_dialog(tables, filename):
    result = [tables[0]]
    dlg = tk.Toplevel(); dlg.title(f"Select Table — {filename}")
    dlg.configure(bg="#0d1117"); dlg.resizable(False, False); dlg.grab_set()
    tk.Label(dlg, text=f"Tables in {filename}:", bg="#0d1117", fg="#e2e8f0",
             font=("Courier New", 10, "bold")).pack(padx=16, pady=8)
    lb_frame = tk.Frame(dlg, bg="#0d1117"); lb_frame.pack(padx=16, pady=4, fill="both")
    sb2 = tk.Scrollbar(lb_frame); sb2.pack(side="right", fill="y")
    lb = tk.Listbox(lb_frame, yscrollcommand=sb2.set, bg="#131b2a", fg="#f1f5f9",
                    selectbackground="#3b82f6", font=("Courier New", 9),
                    height=min(len(tables), 10), width=42, relief="flat", bd=0)
    lb.pack(side="left", fill="both"); sb2.config(command=lb.yview)
    for t in tables: lb.insert("end", f"  {t}")
    lb.selection_set(0)
    def _ok():
        sel = lb.curselection()
        if sel: result[0] = tables[sel[0]]
        dlg.destroy()
    tk.Button(dlg, text="Load Selected", command=_ok, bg="#1e3a8a", fg="#93c5fd",
              font=("Courier New", 9, "bold"), relief="flat", padx=14, pady=8,
              cursor="hand2").pack(pady=(4, 14))
    dlg.wait_window()
    return result[0]

def load_any_file(filepath, progress_cb=None):
    import pandas as pd
    ext = os.path.splitext(filepath)[1].lower().strip(".")
    size_mb = os.path.getsize(filepath) / (1024 * 1024)

    if ext in ("mdb", "accdb"):
        tables, driver = _get_access_tables_pyodbc(filepath)
        if not tables: raise RuntimeError("No user tables found.")
        table = tables[0] if len(tables) == 1 else _pick_table_dialog(tables, os.path.basename(filepath))
        return _load_access_table_pyodbc(filepath, table, driver), f"Access: {table}"

    if ext in ("csv", "tsv", "txt", "") and size_mb > 50:
        sep = "\t" if ext in ("tsv", "txt") else ","
        chunks = []; total = 0
        for enc in ("utf-8", "latin-1"):
            try:
                reader = pd.read_csv(filepath, sep=sep, encoding=enc,
                                     chunksize=CHUNK_ROWS, low_memory=True)
                for i, chunk in enumerate(reader):
                    chunks.append(chunk); total += len(chunk)
                    if progress_cb: progress_cb(min(90, i*5), f"Loaded {total:,} rows…")
                return pd.concat(chunks, ignore_index=True), os.path.basename(filepath)
            except UnicodeDecodeError:
                chunks = []; continue
        raise RuntimeError("Could not decode file.")

    strategies = []
    if ext in ("csv", "tsv", "txt", ""):
        sep = "\t" if ext in ("tsv", "txt") else ","
        strategies += [
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="utf-8", low_memory=False),
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="latin-1", low_memory=False),
            lambda f: pd.read_csv(f, sep=None, engine="python"),
        ]
    if ext in ("xlsx", "xlsm"): strategies += [lambda f: pd.read_excel(f, engine="openpyxl")]
    if ext == "xlsb":           strategies += [lambda f: pd.read_excel(f, engine="pyxlsb")]
    if ext == "xls":            strategies += [lambda f: pd.read_excel(f, engine="xlrd")]
    if ext in ("ods","odf"):    strategies += [lambda f: pd.read_excel(f, engine="odf")]
    if ext == "json":           strategies += [lambda f: pd.read_json(f, orient="records"),
                                               lambda f: pd.read_json(f)]
    if ext == "parquet":        strategies += [lambda f: pd.read_parquet(f)]
    if ext == "feather":        strategies += [lambda f: pd.read_feather(f)]
    if ext in ("h5","hdf5"):    strategies += [lambda f: pd.read_hdf(f)]
    if ext in ("pkl","pickle"): strategies += [lambda f: pd.read_pickle(f)]
    if not strategies:          strategies += [lambda f: pd.read_csv(f), lambda f: pd.read_excel(f)]

    for s in strategies:
        try:
            df = s(filepath)
            if df is not None and len(df.columns) > 0:
                return df, os.path.basename(filepath)
        except: pass
    raise RuntimeError(f"Could not read '{os.path.basename(filepath)}'.")


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
    "legend.fontsize": 9, "figure.autolayout": False,
    "font.family": "monospace",
})

BG_DEEP = "#060a10"; BG = "#0d1117"; BG_MID = "#0f1520"
SIDEBAR = "#080c14"; CARD = "#131b2a"; CARD2 = "#1a2438"
BORDER  = "#1e3a5f"; BORDER2 = "#0e2040"

COL_DATASET    = "#22d3ee"; COL_CLUSTER  = "#a78bfa"
COL_ANOMALY    = "#f87171"; COL_EVENTS   = "#34d399"
COL_FEATURES   = "#60a5fa"; COL_EXPORT   = "#4ade80"
COL_PREVIEW    = "#94a3b8"; COL_CHATBOT  = "#e879f9"

ACCENT = "#3b82f6"; FG = "#f1f5f9"; FG_MID = "#cbd5e1"; FG_DIM = "#4a6080"
SUCCESS = "#10b981"; SUCCESS2 = "#34d399"; WARN = "#f59e0b"
ERR = "#ef4444"; ANOMALY_C = "#ef4444"; NORMAL_C = "#3b82f6"

CLUSTER_PALETTE = ["#f59e0b","#3b82f6","#10b981","#ec4899",
                   "#8b5cf6","#06b6d4","#ef4444","#84cc16","#f97316","#a78bfa"]
EVENT_COLOURS   = {"well on":"#10b981","well off":"#ef4444",
                   "maintenance shutdown":"#f59e0b","pump failure":"#f97316",
                   "water breakthrough":"#06b6d4"}
SEV_COLORS = {"NONE":"#2d3f5e","LOW":"#10b981","MODERATE":"#f59e0b",
              "ELEVATED":"#f97316","HIGH":"#ef4444"}

FONT_H1 = ("Georgia",13,"bold"); FONT_H2 = ("Georgia",11,"bold")
FONT_H3 = ("Courier New",10,"bold"); FONT_SM = ("Courier New",9); FONT_XS = ("Courier New",8)

# ── Global state ──────────────────────────────────────────────────────────────
raw_data              = None
active_df             = None
active_X              = None
active_xcols          = []
active_anomaly_result = None
active_event_result   = None
active_well_df        = None
active_figures        = {}
active_cluster_method = "—"
active_insights_text  = ""

feat_all_vars   = {}
feat_all_frames = []
feat_count_var  = None

# Chatbot conversation history
chatbot_history = []


# ══════════════════════════════════════════════════════════════════════════════
# WELL ID DETECTION
# ══════════════════════════════════════════════════════════════════════════════
def _find_well_id_column(df):
    id_kw = ["wellid","well_id","wellname","well_name","wellno","well_no",
             "uwi","api","well","borehole","bore","field_well","slot"]
    param_kw = ["status","state","type","mode","zone","formation","operator",
                "company","region","area","basin","country","field"]

    for c in df.columns:
        cl = c.lower().replace(" ","").replace("_","").replace("-","")
        for kw in id_kw:
            if cl == kw.replace("_","") or cl.startswith(kw.replace("_","")):
                return c

    n = len(df)
    candidates = []
    for c in df.columns:
        import pandas as pd
        nuniq = df[c].nunique()
        is_obj = not pd.api.types.is_numeric_dtype(df[c])
        is_int = pd.api.types.is_integer_dtype(df[c])
        cl = c.lower()

        if any(kw in cl for kw in param_kw):
            continue
        if pd.api.types.is_float_dtype(df[c]):
            continue
        if nuniq < 2 or nuniq > n * 0.90:
            continue

        if is_obj or is_int:
            score = 0
            if any(k in cl for k in ["id","no","num","name","code","key","ref"]):
                score += 10
            if is_obj:
                score += 5
            score += nuniq / n * 5
            candidates.append((score, c))

    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]

    return None


# ══════════════════════════════════════════════════════════════════════════════
# SAFE MODE HELPER
# ══════════════════════════════════════════════════════════════════════════════
def _safe_mode(series):
    try:
        s = series.dropna()
        if len(s) == 0:
            return ""
        m = s.mode()
        if len(m) == 0:
            return s.iloc[0]
        return m.iloc[0]
    except Exception:
        return ""


# ══════════════════════════════════════════════════════════════════════════════
# WELL AGGREGATION
# ══════════════════════════════════════════════════════════════════════════════
def aggregate_by_well(df, selected_cols, well_id_col):
    import pandas as pd

    selected_cols = [c for c in selected_cols if c in df.columns]
    if not selected_cols:
        return df.iloc[:0].copy(), [], []

    num_cols = [c for c in selected_cols if pd.api.types.is_numeric_dtype(df[c])]
    cat_cols = [c for c in selected_cols if not pd.api.types.is_numeric_dtype(df[c])]

    def _agg_groups(groups, id_col=None):
        records = []
        for key, grp in groups:
            row = {}
            if id_col is not None:
                row[id_col] = key

            for c in num_cols:
                vals = grp[c].dropna()
                if len(vals) == 0:
                    row[f"{c}__mean"] = np.nan
                    row[f"{c}__std"]  = 0.0
                    row[f"{c}__min"]  = np.nan
                    row[f"{c}__max"]  = np.nan
                else:
                    row[f"{c}__mean"] = float(vals.mean())
                    row[f"{c}__std"]  = float(vals.std()) if len(vals) > 1 else 0.0
                    row[f"{c}__min"]  = float(vals.min())
                    row[f"{c}__max"]  = float(vals.max())

            for c in cat_cols:
                row[f"{c}__mode"] = str(_safe_mode(grp[c]))

            row["record_count"] = len(grp)
            records.append(row)

        if not records:
            return pd.DataFrame()
        return pd.DataFrame(records).reset_index(drop=True)

    if well_id_col and well_id_col in df.columns:
        groups = df.groupby(well_id_col, sort=False)
        well_df = _agg_groups(groups, id_col=well_id_col)

        num_agg_cols = [f"{c}__{s}" for c in num_cols
                        for s in ["mean","std","min","max"]
                        if f"{c}__{s}" in well_df.columns]
        cat_agg_cols = [f"{c}__mode" for c in cat_cols
                        if f"{c}__mode" in well_df.columns]
        return well_df, num_agg_cols, cat_agg_cols

    if cat_cols:
        try:
            groups = df.groupby(cat_cols, sort=False, dropna=False)
            well_df = _agg_groups(groups, id_col=None)

            num_agg_cols = [f"{c}__{s}" for c in num_cols
                            for s in ["mean","std","min","max"]
                            if f"{c}__{s}" in well_df.columns]
            cat_agg_cols = [f"{c}__mode" for c in cat_cols
                            if f"{c}__mode" in well_df.columns]

            if len(well_df) >= 2:
                return well_df, num_agg_cols, cat_agg_cols
        except Exception:
            pass

    result = df[selected_cols].copy().reset_index(drop=True)
    rename_map = {}
    for c in num_cols:
        rename_map[c] = f"{c}__mean"
    for c in cat_cols:
        result[c] = result[c].fillna("MISSING").astype(str)
        rename_map[c] = f"{c}__mode"
    result = result.rename(columns=rename_map)
    result["record_count"] = 1

    num_agg_cols = [f"{c}__mean" for c in num_cols if f"{c}__mean" in result.columns]
    cat_agg_cols = [f"{c}__mode" for c in cat_cols if f"{c}__mode" in result.columns]
    return result, num_agg_cols, cat_agg_cols


# ══════════════════════════════════════════════════════════════════════════════
# ROBUST MIXED-DATA CLUSTERING
# ══════════════════════════════════════════════════════════════════════════════
def _impute_numeric(arr):
    arr = arr.copy().astype(float)
    for j in range(arr.shape[1]):
        col = arr[:, j]
        nan_mask = np.isnan(col)
        if nan_mask.all():
            arr[:, j] = 0.0
        elif nan_mask.any():
            arr[nan_mask, j] = float(np.nanmean(col))
    return arr


def run_clustering_robust(well_df, num_agg_cols, cat_agg_cols, n_clusters, weight_map):
    import pandas as pd
    from sklearn.cluster import AgglomerativeClustering

    n = len(well_df)
    n_clusters = max(2, min(n_clusters, n - 1 if n > 2 else 1))

    has_num = bool(num_agg_cols)
    has_cat = bool(cat_agg_cols)

    X_num = np.zeros((n, 0), dtype=float)
    w_num = np.ones(0, dtype=float)

    if has_num:
        raw = well_df[num_agg_cols].values.astype(float)
        X_num = _impute_numeric(raw)
        lo = X_num.min(axis=0); hi = X_num.max(axis=0)
        rng = np.where(hi - lo > 0, hi - lo, 1.0)
        X_num = (X_num - lo) / rng

        w_num_list = []
        for c in num_agg_cols:
            orig = (c.replace("__mean","").replace("__std","")
                     .replace("__min","").replace("__max",""))
            w_num_list.append(float(weight_map.get(orig, 1.0)))
        w_num = np.array(w_num_list, dtype=float)
        ws = w_num.sum()
        if ws > 0: w_num = w_num / ws

    cat_matrix = None
    w_cat = np.ones(0, dtype=float)

    if has_cat:
        cat_matrix = well_df[cat_agg_cols].fillna("MISSING").astype(str).values
        w_cat = np.ones(len(cat_agg_cols), dtype=float)
        if len(cat_agg_cols) > 0:
            w_cat = w_cat / w_cat.sum()

    if has_num and has_cat:
        try:
            from kmodes.kprototypes import KPrototypes
            X_mixed = np.hstack([X_num, cat_matrix])
            cat_idx = list(range(X_num.shape[1], X_mixed.shape[1]))
            kp = KPrototypes(n_clusters=n_clusters, init="Cao", n_init=3, random_state=42)
            labels = kp.fit_predict(X_mixed, categorical=cat_idx)
            return labels.astype(int), float(kp.cost_), f"KPrototypes (n={n_clusters})"
        except Exception:
            pass

    if has_cat and not has_num:
        try:
            from kmodes.kmodes import KModes
            km = KModes(n_clusters=n_clusters, init="Cao", n_init=3, random_state=42)
            labels = km.fit_predict(cat_matrix)
            return labels.astype(int), float(km.cost_), f"KModes/categorical (n={n_clusters})"
        except Exception:
            pass

    if n <= 8000:
        try:
            dist_matrix = _gower_fast(X_num, cat_matrix, w_num, w_cat)
            dist_matrix = np.nan_to_num(dist_matrix, nan=1.0, posinf=1.0, neginf=0.0)
            np.fill_diagonal(dist_matrix, 0.0)
            ac = AgglomerativeClustering(n_clusters=n_clusters,
                                          metric="precomputed", linkage="average")
            labels = ac.fit_predict(dist_matrix).astype(int)
            inertia = _pseudo_inertia(dist_matrix, labels)
            return labels, inertia, f"Agglomerative/Gower (n={n_clusters})"
        except Exception:
            pass

    try:
        X_enc = _build_encoded_matrix_v2(X_num, cat_matrix, w_num, w_cat)
        if X_enc.shape[1] > 0 and X_enc.shape[0] >= 2:
            ac2 = AgglomerativeClustering(n_clusters=n_clusters,
                                           metric="euclidean", linkage="ward")
            labels = ac2.fit_predict(X_enc).astype(int)
            inertia = _pseudo_inertia_euclidean(X_enc, labels)
            return labels, inertia, f"Agglomerative/Ward+encoded (n={n_clusters})"
    except Exception:
        pass

    try:
        from sklearn.cluster import KMeans
        X_enc = _build_encoded_matrix_v2(X_num, cat_matrix, w_num, w_cat)
        if X_enc.shape[1] == 0:
            X_enc = np.zeros((n, 1))
        km = KMeans(n_clusters=n_clusters, random_state=42, n_init="auto")
        labels = km.fit_predict(X_enc).astype(int)
        return labels, float(km.inertia_), f"KMeans+encoded (n={n_clusters})"
    except Exception:
        pass

    labels = np.arange(n) % n_clusters
    return labels.astype(int), 0.0, "Round-robin fallback"


def _gower_fast(X_num, cat_matrix, w_num, w_cat):
    n = X_num.shape[0]
    dist = np.zeros((n, n), dtype=np.float32)

    p_num = X_num.shape[1]
    p_cat = cat_matrix.shape[1] if cat_matrix is not None else 0
    total_cols = p_num + p_cat
    if total_cols == 0:
        return dist

    if p_num > 0:
        for k in range(p_num):
            col = X_num[:, k].reshape(-1, 1)
            d = np.abs(col - col.T)
            dist += (w_num[k] if k < len(w_num) else 1.0 / p_num) * d

    if p_cat > 0 and cat_matrix is not None:
        for k in range(p_cat):
            col = cat_matrix[:, k]
            mismatch = (col.reshape(-1, 1) != col.reshape(1, -1)).astype(np.float32)
            dist += (w_cat[k] if k < len(w_cat) else 1.0 / p_cat) * mismatch

    total_w = (w_num.sum() if p_num > 0 else 0) + (w_cat.sum() if p_cat > 0 else 0)
    if total_w > 0:
        dist /= total_w

    return dist.astype(np.float64)


def _build_encoded_matrix_v2(X_num, cat_matrix, w_num, w_cat):
    import pandas as pd
    cols = []

    if X_num is not None and X_num.shape[1] > 0:
        for j in range(X_num.shape[1]):
            w = float(w_num[j]) if j < len(w_num) else 1.0
            cols.append((X_num[:, j] * w).reshape(-1, 1))

    if cat_matrix is not None and cat_matrix.shape[1] > 0:
        for k in range(cat_matrix.shape[1]):
            vals = cat_matrix[:, k]
            codes = pd.Categorical(vals).codes.astype(float)
            lo, hi = codes.min(), codes.max()
            rng = hi - lo if hi != lo else 1.0
            scaled = (codes - lo) / rng
            w = float(w_cat[k]) if k < len(w_cat) else 1.0
            cols.append((scaled * w).reshape(-1, 1))

    if not cols:
        return np.zeros((X_num.shape[0] if X_num is not None else 1, 1))
    return np.hstack(cols)


def _pseudo_inertia(dist_matrix, labels):
    total = 0.0; count = 0
    for cl in np.unique(labels):
        idx = np.where(labels == cl)[0]
        if len(idx) < 2: continue
        sub = dist_matrix[np.ix_(idx, idx)]
        total += float(sub.sum())
        count += len(idx) * (len(idx) - 1)
    return total / max(count, 1)


def _pseudo_inertia_euclidean(X, labels):
    total = 0.0
    for cl in np.unique(labels):
        mask = labels == cl
        sub = X[mask]
        if len(sub) == 0: continue
        centroid = sub.mean(axis=0)
        total += float(((sub - centroid) ** 2).sum())
    return total


# ══════════════════════════════════════════════════════════════════════════════
# UTILITY
# ══════════════════════════════════════════════════════════════════════════════
def cluster_hex(n):
    return [CLUSTER_PALETTE[i % len(CLUSTER_PALETTE)] for i in range(n)]

def _lighten(hex_col, amount=30):
    try:
        h = hex_col.lstrip("#")
        r,g,b = int(h[0:2],16),int(h[2:4],16),int(h[4:6],16)
        return "#{:02x}{:02x}{:02x}".format(min(r+amount,255),min(g+amount,255),min(b+amount,255))
    except: return hex_col

def ts(): return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

def set_status(msg, color=FG_DIM):
    status_var.set(msg); status_lbl.config(fg=color)

def _save_df(df, path):
    ext = os.path.splitext(path)[1].lower()
    df.to_excel(path, index=False) if ext in (".xlsx",".xls") else df.to_csv(path, index=False)


# ══════════════════════════════════════════════════════════════════════════════
# UI WIDGETS
# ══════════════════════════════════════════════════════════════════════════════
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget; self.text = text; self.tip = None
        widget.bind("<Enter>", self._show); widget.bind("<Leave>", self._hide)
    def _show(self, _=None):
        x = self.widget.winfo_rootx()+20; y = self.widget.winfo_rooty()+self.widget.winfo_height()+4
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True); tw.wm_geometry(f"+{x}+{y}"); tw.configure(bg=BORDER)
        tk.Label(tw, text=self.text, bg=CARD2, fg=FG_MID, font=FONT_XS, padx=8, pady=4).pack()
    def _hide(self, _=None):
        if self.tip: self.tip.destroy(); self.tip = None

def make_scrollable(parent, bg=BG):
    outer = tk.Frame(parent, bg=bg)
    canvas = Canvas(outer, bg=bg, highlightthickness=0, bd=0)
    sb = Scrollbar(outer, orient="vertical", command=canvas.yview,
                   bg=BORDER, troughcolor=bg, activebackground=ACCENT)
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(canvas, bg=bg)
    wid = canvas.create_window((0,0), window=inner, anchor="nw")
    def _resize(e):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(wid, width=e.width)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", _resize)
    def _wheel(e):
        if e.num==4: canvas.yview_scroll(-1,"units")
        elif e.num==5: canvas.yview_scroll(1,"units")
        else: canvas.yview_scroll(int(-1*(e.delta/120)),"units")
    canvas.bind_all("<MouseWheel>", _wheel)
    canvas.bind_all("<Button-4>", _wheel); canvas.bind_all("<Button-5>", _wheel)
    return outer, inner

def make_section_card(parent, title, accent_color, **pack_kw):
    wrapper = tk.Frame(parent, bg=SIDEBAR, pady=0)
    wrapper.pack(**{"fill":"x","pady":(0,2),**pack_kw})
    hdr = tk.Frame(wrapper, bg=accent_color, height=26)
    hdr.pack(fill="x"); hdr.pack_propagate(False)
    dark_text = accent_color in ("#f59e0b", COL_EXPORT, COL_EVENTS)
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
                  font=FONT_H3, relief="flat", bd=0, cursor="hand2", padx=10, pady=8, anchor="w")
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
    tk.Label(row, text=label[:label_width], bg=CARD, fg=FG_DIM, font=FONT_XS, width=label_width).pack(side="left")
    outer = tk.Frame(row, bg=BORDER2, height=7); outer.pack(side="left", fill="x", expand=True, padx=(3,0))
    tk.Frame(outer, bg=col, height=7).place(relwidth=max(frac,0.03), relheight=1.0)


# ══════════════════════════════════════════════════════════════════════════════
# UNIFIED FEATURE SELECTION PANEL
# ══════════════════════════════════════════════════════════════════════════════
def _update_feat_count(*_):
    n_sel = sum(1 for v in feat_all_vars.values() if v.get())
    if feat_count_var:
        feat_count_var.set(f"{n_sel} / {len(feat_all_vars)} selected")

def _filter_feat_list(*_):
    query = feat_search_var.get().lower().strip()
    for col, frm in feat_all_frames:
        show = (query == "" or query in col.lower())
        if show: frm.pack(fill="x", padx=2, pady=1)
        else:    frm.pack_forget()

def rebuild_feature_panel(df):
    global feat_all_vars, feat_all_frames
    for _, f in feat_all_frames:
        try: f.destroy()
        except: pass
    feat_all_vars.clear(); feat_all_frames.clear()

    import pandas as pd

    for col in df.columns:
        is_num = pd.api.types.is_numeric_dtype(df[col])
        nuniq  = df[col].nunique()
        dtype  = str(df[col].dtype)[:10]

        if not is_num and nuniq > min(500, len(df) * 0.8):
            continue

        var = tk.BooleanVar(value=is_num)
        var.trace_add("write", _update_feat_count)
        feat_all_vars[col] = var

        frm = tk.Frame(feat_all_body, bg=CARD2,
                       highlightbackground=BORDER2, highlightthickness=1)

        type_col = COL_FEATURES if is_num else COL_EVENTS
        tk.Canvas(frm, bg=type_col, width=6, height=6,
                  highlightthickness=0).pack(side="left", padx=(4,2), pady=6)

        cb = tk.Checkbutton(frm, text=f"{col[:30]}",
                            variable=var,
                            bg=CARD2, fg=type_col,
                            selectcolor=CARD, activebackground=CARD2,
                            activeforeground=type_col,
                            font=FONT_XS, anchor="w")
        cb.pack(side="left", padx=2, pady=2, fill="x", expand=True)

        badge = f"{dtype}  |  {nuniq} uniq"
        tk.Label(frm, text=badge, bg=CARD2, fg=FG_DIM,
                 font=("Courier New",7)).pack(side="right", padx=6)

        feat_all_frames.append((col, frm))
        frm.pack(fill="x", padx=2, pady=1)

    _update_feat_count()

def get_selected_features():
    return [c for c, v in feat_all_vars.items() if v.get()]


# ══════════════════════════════════════════════════════════════════════════════
# PROGRESS
# ══════════════════════════════════════════════════════════════════════════════
def progress_start():
    progress_bar.config(mode="indeterminate"); progress_bar.start(12)

def progress_stop():
    progress_bar.stop(); progress_bar.config(mode="determinate"); progress_bar["value"] = 100


# ══════════════════════════════════════════════════════════════════════════════
# DATA LOADING
# ══════════════════════════════════════════════════════════════════════════════
def upload_dataset():
    global raw_data
    f = filedialog.askopenfilename(
        title="Open Data File",
        filetypes=[("All supported","*.csv *.tsv *.txt *.xlsx *.xls *.xlsb *.xlsm *.ods "
                    "*.json *.parquet *.feather *.h5 *.hdf5 *.pkl *.pickle *.mdb *.accdb"),
                   ("CSV / TSV","*.csv *.tsv *.txt"),("Excel","*.xlsx *.xls *.xlsb *.xlsm"),
                   ("MS Access","*.mdb *.accdb"),("All files","*.*")])
    if not f: return
    size_mb = os.path.getsize(f) / (1024*1024)
    set_status(f"Loading {os.path.basename(f)}  ({size_mb:.1f} MB)…", WARN)
    progress_start()

    def _load():
        global raw_data
        try:
            df, label = load_any_file(f, progress_cb=lambda p,m: app.after(0, lambda: set_status(m, WARN)))
            if len(df) == 0: raise RuntimeError("File has no rows.")
            raw_data = df
            import pandas as pd
            num_cols = list(df.select_dtypes(include="number").columns)
            app.after(0, lambda: _post_load(df, f, num_cols, label, size_mb))
        except Exception as e:
            app.after(0, lambda err=str(e): (_load_error(err)))

    def _post_load(df, filepath, num_cols, source_label, size_mb):
        well_col = _find_well_id_column(df)
        n_rows   = len(df)
        n_unique = df[well_col].nunique() if well_col else n_rows

        rows_var.set(f"{n_unique:,}")
        cols_var.set(str(len(df.columns)))
        num_cols_var.set(f"{len(num_cols)} numeric")

        rebuild_feature_panel(df)

        ec = _find_event_column(df)
        event_col_var.set(f"✔  Event: '{ec}'" if ec else "⚠  No event column")
        event_col_lbl.config(fg=SUCCESS2 if ec else WARN)

        well_id_var.set(f"✔  Well ID: '{well_col}'  ({n_unique:,} wells)" if well_col
                        else "⚠  No well ID column detected")
        well_id_lbl.config(fg=SUCCESS2 if well_col else WARN)

        preview_table(df)
        progress_stop()
        set_status(f"✔  {source_label}  —  {n_rows:,} rows × {len(df.columns)} cols  ({size_mb:.1f} MB)", SUCCESS)
        for k in stat_widgets: stat_widgets[k].set("—")
        stat_widgets["Unique Wells"].set(f"{n_unique:,}")
        stat_widgets["Total Records"].set(f"{n_rows:,}")
        export_btn.config(state="normal")

    def _load_error(msg):
        progress_stop(); set_status("Load failed", ERR); messagebox.showerror("Load Error", msg)

    threading.Thread(target=_load, daemon=True).start()

def _find_event_column(df):
    for c in df.columns:
        if any(k in c.lower().replace(" ","_") for k in ["status","state","event","mode","condition","operation"]):
            return c
    return None

def preview_table(df):
    for w in table_frame.winfo_children(): w.destroy()
    style = tk_ttk.Style()
    style.configure("P.Treeview", background=CARD, foreground=FG,
                    fieldbackground=CARD, rowheight=21, font=FONT_SM)
    style.configure("P.Treeview.Heading", background="#0e2040", foreground=COL_PREVIEW, font=FONT_H3)
    style.map("P.Treeview", background=[("selected", ACCENT)])
    tv  = tk_ttk.Treeview(table_frame, style="P.Treeview")
    vsb = tk_ttk.Scrollbar(table_frame, orient="vertical",   command=tv.yview)
    hsb = tk_ttk.Scrollbar(table_frame, orient="horizontal", command=tv.xview)
    tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tv["columns"] = list(df.columns); tv["show"] = "headings"
    for col in df.columns:
        tv.heading(col, text=col); tv.column(col, width=90, minwidth=60)
    for row in df.head(20).values:
        tv.insert("","end", values=list(row))
    hsb.pack(side="bottom", fill="x"); vsb.pack(side="right", fill="y")
    tv.pack(fill="both", expand=True)


# ══════════════════════════════════════════════════════════════════════════════
# ANOMALY DETECTION
# ══════════════════════════════════════════════════════════════════════════════
def detect_anomalies(X_w, contamination=0.05, method="iforest"):
    from sklearn.ensemble import IsolationForest
    from sklearn.neighbors import LocalOutlierFactor
    n = X_w.shape[0]
    if n < 5:
        return np.ones(n,dtype=int), 0.0, [], "N/A (too few samples)", np.zeros(n)
    safe_cont = min(max(float(np.clip(contamination,0.001,0.499)), 1.0/n), 0.499)
    if method == "lof":
        k = max(5, min(20, n//10))
        det = LocalOutlierFactor(n_neighbors=k, contamination=safe_cont)
        labels = det.fit_predict(X_w); scores = det.negative_outlier_factor_
        name = f"Local Outlier Factor (k={k})"
    else:
        n_est = 100 if n <= 500 else 200
        det = IsolationForest(n_estimators=n_est, contamination=safe_cont, random_state=42, n_jobs=-1)
        labels = det.fit_predict(X_w); scores = det.decision_function(X_w)
        name = f"Isolation Forest (n_est={n_est})"
    idx = list(np.where(labels==-1)[0])
    return labels, len(idx)/n*100, idx, name, scores

def analyse_events(df, anomaly_labels):
    ec = _find_event_column(df)
    if ec is None:
        return {"has_events":False,"event_col":None,"event_counts":{},
                "event_labels":np.full(len(df),"Unknown"),
                "n_active":0,"n_inactive":0,
                "n_abnormal":int((anomaly_labels==-1).sum()),
                "operational_anomalies":[],"message":"No event/status column found."}
    raw_vals = df[ec].astype(str).str.strip()
    from collections import Counter
    event_counts = dict(Counter(raw_vals.values))
    active_kw   = ["on","active","running","producing","open"]
    inactive_kw = ["off","inactive","shut","stop","closed","idle","down"]
    def _clf(val):
        vl = val.lower()
        if any(k in vl for k in active_kw):   return "active"
        if any(k in vl for k in inactive_kw): return "inactive"
        return "other"
    statuses = raw_vals.apply(_clf)
    return {"has_events":True,"event_col":ec,"event_counts":event_counts,
            "event_labels":raw_vals.values,
            "n_active":int((statuses=="active").sum()),
            "n_inactive":int((statuses=="inactive").sum()),
            "n_abnormal":int((anomaly_labels==-1).sum()),
            "operational_anomalies":[],"message":f"Events from '{ec}'"}


# ══════════════════════════════════════════════════════════════════════════════
# BUILD ANOMALY DETECTION MATRIX
# ══════════════════════════════════════════════════════════════════════════════
def build_anomaly_X(well_df, num_agg_cols, cat_agg_cols, weight_map):
    import pandas as pd
    n = len(well_df)
    cols = []

    if num_agg_cols:
        raw = well_df[num_agg_cols].values.astype(float)
        raw = _impute_numeric(raw)
        mu = raw.mean(axis=0); s = raw.std(axis=0); s[s==0] = 1.0
        X_scaled = (raw - mu) / s
        for j, c in enumerate(num_agg_cols):
            orig = (c.replace("__mean","").replace("__std","")
                     .replace("__min","").replace("__max",""))
            w = float(weight_map.get(orig, 1.0))
            cols.append((X_scaled[:, j] * np.sqrt(w/100.0 if w <= 100 else 1.0)).reshape(-1, 1))

    if cat_agg_cols:
        for c in cat_agg_cols:
            vals = well_df[c].fillna("MISSING").astype(str).values
            codes = pd.Categorical(vals).codes.astype(float)
            lo, hi = codes.min(), codes.max()
            rng = hi - lo if hi != lo else 1.0
            cols.append(((codes - lo) / rng).reshape(-1, 1))

    if not cols:
        return np.zeros((n, 1))
    return np.hstack(cols)


# ══════════════════════════════════════════════════════════════════════════════
# DETAILED ANOMALY EXPLANATION — per well, per parameter
# ══════════════════════════════════════════════════════════════════════════════
def explain_anomalous_wells(well_df, num_agg_cols, cat_agg_cols, anomaly_labels,
                             anomaly_scores, well_id_col, z_threshold=2.5):
    import pandas as pd

    anom_idx = np.where(anomaly_labels == -1)[0]
    if len(anom_idx) == 0:
        return []

    num_stats = {}
    for c in num_agg_cols:
        vals = well_df[c].dropna().values.astype(float)
        if len(vals) == 0:
            continue
        mu = float(np.mean(vals))
        sd = float(np.std(vals))
        if sd == 0:
            sd = 1.0
        num_stats[c] = {"mean": mu, "std": sd,
                        "p10": float(np.percentile(vals, 10)),
                        "p90": float(np.percentile(vals, 90))}

    explanations = []
    for idx in anom_idx:
        row = well_df.iloc[idx]
        well_name = str(row[well_id_col]) if (well_id_col and well_id_col in well_df.columns) else f"Well #{idx}"
        score = float(anomaly_scores[idx]) if idx < len(anomaly_scores) else 0.0

        flagged_params = []
        for c in num_agg_cols:
            if c not in num_stats:
                continue
            val = row.get(c, np.nan)
            if pd.isna(val):
                flagged_params.append({
                    "param": c, "value": "MISSING",
                    "z_score": None,
                    "direction": "missing",
                    "reason": "Value is missing / NaN"
                })
                continue
            val = float(val)
            st = num_stats[c]
            z = (val - st["mean"]) / st["std"]
            if abs(z) > z_threshold:
                direction = "HIGH" if z > 0 else "LOW"
                flagged_params.append({
                    "param": c,
                    "value": round(val, 4),
                    "z_score": round(z, 3),
                    "direction": direction,
                    "mean": round(st["mean"], 4),
                    "std": round(st["std"], 4),
                    "p10": round(st["p10"], 4),
                    "p90": round(st["p90"], 4),
                    "reason": (f"Value {val:.4f} is {abs(z):.2f}σ "
                               f"{'above' if z>0 else 'below'} mean "
                               f"({st['mean']:.4f}±{st['std']:.4f}). "
                               f"Normal range P10–P90: {st['p10']:.4f}–{st['p90']:.4f}.")
                })

        for c in cat_agg_cols:
            val = str(row.get(c, "MISSING"))
            all_vals = well_df[c].fillna("MISSING").astype(str)
            global_mode = str(all_vals.mode().iloc[0]) if len(all_vals.mode()) > 0 else "UNKNOWN"
            freq = (all_vals == val).sum() / max(len(all_vals), 1) * 100
            if freq < 5:
                flagged_params.append({
                    "param": c,
                    "value": val,
                    "z_score": None,
                    "direction": "RARE CATEGORY",
                    "reason": (f"Category '{val}' is rare ({freq:.1f}% of wells). "
                               f"Most common: '{global_mode}'.")
                })

        flagged_params.sort(key=lambda x: abs(x.get("z_score") or 0), reverse=True)

        explanations.append({
            "well": well_name,
            "index": int(idx),
            "anomaly_score": score,
            "flagged_params": flagged_params,
            "n_flagged": len(flagged_params)
        })

    explanations.sort(key=lambda x: x["anomaly_score"])
    return explanations


# ══════════════════════════════════════════════════════════════════════════════
# VISUALIZATIONS
# ══════════════════════════════════════════════════════════════════════════════
def _style_ax(ax):
    ax.set_facecolor("#161d2e")
    for sp in ax.spines.values(): sp.set_edgecolor("#2d3f5e")
    ax.tick_params(colors="#94a3b8", labelsize=8)

def plot_well_clusters(well_df, agg_cols, hx, cluster_method=""):
    from sklearn.decomposition import PCA
    import pandas as pd
    fig, ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)

    valid_cols = [c for c in agg_cols if c in well_df.columns]

    if len(valid_cols) >= 1 and "cluster" in well_df.columns:
        encoded_cols = []
        for c in valid_cols:
            col_data = well_df[c]
            if pd.api.types.is_numeric_dtype(col_data):
                arr = col_data.fillna(col_data.mean() if col_data.notna().any() else 0).values.astype(float)
            else:
                arr = pd.Categorical(col_data.fillna("MISSING").astype(str)).codes.astype(float)
            lo, hi = arr.min(), arr.max()
            rng = hi - lo if hi != lo else 1.0
            encoded_cols.append(((arr - lo) / rng).reshape(-1, 1))

        if not encoded_cols:
            ax.text(0.2, 0.5, "No usable feature columns", transform=ax.transAxes,
                    color=FG_DIM, fontsize=10)
        else:
            data = np.hstack(encoded_cols)
            data = _impute_numeric(data)

            n_components = min(2, data.shape[1], data.shape[0] - 1)
            if n_components >= 2:
                Z = PCA(n_components=2, random_state=42).fit_transform(data)
            elif n_components == 1:
                Z = np.column_stack([data[:, 0], np.arange(len(data)) / max(len(data)-1, 1)])
            else:
                Z = np.column_stack([np.arange(len(data)), np.zeros(len(data))])

            unique_clusters = sorted(well_df["cluster"].unique())
            for cl in unique_clusters:
                mask = well_df["cluster"].values == cl
                col = hx[int(cl) % len(hx)]
                ax.scatter(Z[mask, 0], Z[mask, 1], color=col, s=80,
                           edgecolors="#ffffff44", linewidths=0.7, zorder=3,
                           label=f"Cluster {cl}  ({int(mask.sum()):,} wells)")

            ax.set_xlabel("PC 1  (combined well profile)", fontsize=9)
            ax.set_ylabel("PC 2  (combined well profile)", fontsize=9)
    else:
        ax.text(0.2, 0.5, "Not enough data to plot", transform=ax.transAxes, color=FG_DIM, fontsize=10)

    method_tag = f"  [{cluster_method}]" if cluster_method else ""
    n_wells = len(well_df)
    n_cl = well_df["cluster"].nunique() if "cluster" in well_df.columns else 0
    ax.set_title(f"Well Clusters{method_tag}\n{n_wells:,} wells  ·  {n_cl} clusters  ·  "
                 f"{len(valid_cols)} features used",
                 color=COL_CLUSTER, fontsize=11, pad=10, fontweight="bold")
    ax.legend(loc="upper left", fontsize=8.5, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e", labelcolor="#e2e8f0", markerscale=1.4)
    ax.margins(0.10)
    return fig

def plot_pca(X, labels, hex_colors, cluster_method=""):
    from sklearn.decomposition import PCA
    fig, ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    if X.shape[0]<2 or X.shape[1]<1:
        ax.text(0.3,0.5,"Not enough data",transform=ax.transAxes,color=FG_DIM); return fig
    n = min(2, X.shape[1]); Z = PCA(n_components=n).fit_transform(X)
    if Z.shape[1]==1: Z = np.hstack([Z,np.zeros_like(Z)])
    for cl in sorted(np.unique(labels)):
        mask=labels==cl; col=hex_colors[int(cl)%len(hex_colors)]
        ax.scatter(Z[mask,0],Z[mask,1],color=col,s=55,
                   edgecolors="#ffffff22",linewidths=0.5,zorder=3,
                   label=f"Cluster {cl}  ({int(mask.sum()):,} wells)")
    ax.set_xlabel("PC 1",fontsize=9); ax.set_ylabel("PC 2",fontsize=9)
    method_tag = f"  [{cluster_method}]" if cluster_method else ""
    ax.set_title(f"PCA — Well Feature Space{method_tag}",color=COL_CLUSTER,fontsize=12,pad=12,fontweight="bold")
    ax.legend(loc="upper left",fontsize=8.5,framealpha=0.92,
              edgecolor="#334155",facecolor="#161d2e",labelcolor="#e2e8f0",markerscale=1.4)
    return fig

def plot_hidden_patterns(X_w, anomaly_labels, xcols, detector_name):
    from sklearn.decomposition import PCA
    fig,ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    n = X_w.shape[0]
    if X_w.shape[1]>=2:
        Z = PCA(n_components=2,random_state=42).fit_transform(X_w)
        xlabel,ylabel = "PC 1  (feature projection)","PC 2  (feature projection)"
    elif X_w.shape[1]==1:
        Z = np.column_stack([X_w[:,0],np.arange(n)])
        xlabel = xcols[0] if xcols else "Feature"; ylabel = "Well Index"
    else:
        ax.text(0.3,0.5,"No feature data",transform=ax.transAxes,color=FG_DIM); return fig
    nm=anomaly_labels==1; am=anomaly_labels==-1
    ax.scatter(Z[nm,0],Z[nm,1],color=NORMAL_C,s=35,alpha=0.70,
               edgecolors="#ffffff18",linewidths=0.3,zorder=3,
               label=f"● Normal  ({int(nm.sum()):,} wells)")
    if am.sum()>0:
        ax.scatter(Z[am,0],Z[am,1],color=ANOMALY_C,s=80,alpha=0.95,
                   edgecolors="#ffffff66",linewidths=0.9,marker="D",zorder=5,
                   label=f"◆ Anomalous  ({int(am.sum()):,} wells)")
        ax.annotate(f"◆ {int(am.sum())} anomalous",
                    xy=(Z[am,0].mean(),Z[am,1].mean()),xytext=(14,14),
                    textcoords="offset points",color=ANOMALY_C,fontsize=9,fontweight="bold",
                    arrowprops=dict(arrowstyle="->",color=ANOMALY_C,lw=1.0))
    ax.legend(loc="upper right",fontsize=9,framealpha=0.92,
              edgecolor="#334155",facecolor="#161d2e",labelcolor="#e2e8f0",markerscale=1.3)
    ax.set_xlabel(xlabel,fontsize=9); ax.set_ylabel(ylabel,fontsize=9)
    ax.set_title(f"Hidden Pattern Detection  ·  {detector_name}",
                 color=ANOMALY_C,fontsize=12,pad=12,fontweight="bold")
    ax.margins(0.10); return fig

def plot_event_chart(df, event_result):
    fig,ax=plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    if not event_result["has_events"]:
        ax.text(0.3,0.5,"No event column found",transform=ax.transAxes,color=FG_DIM,fontsize=11)
        ax.set_title("Operational Events",color=COL_EVENTS,fontsize=13,pad=12,fontweight="bold"); return fig
    counts=event_result["event_counts"]; labels=list(counts.keys()); sizes=list(counts.values())
    colors_ev=[]
    for lbl in labels:
        ll=lbl.lower(); matched=next((c for k,c in EVENT_COLOURS.items() if k in ll),None)
        colors_ev.append(matched or CLUSTER_PALETTE[len(colors_ev)%len(CLUSTER_PALETTE)])
    ax.barh(labels,sizes,color=colors_ev,edgecolor="#ffffff22",linewidth=0.5)
    for i,(lbl,cnt) in enumerate(zip(labels,sizes)):
        ax.text(cnt*1.01,i,f"{cnt:,}",va="center",ha="left",color=FG_MID,fontsize=8)
    ax.set_xlabel("Record Count",fontsize=9)
    ax.set_title(f"Operational Events — '{event_result['event_col']}'",color=COL_EVENTS,fontsize=12,pad=12,fontweight="bold")
    ax.margins(0.04,0.15); fig.patch.set_facecolor("#0d1117"); return fig


# ══════════════════════════════════════════════════════════════════════════════
# LEGEND BUILDERS
# ══════════════════════════════════════════════════════════════════════════════
def build_cluster_legend(parent, df=None, hex_colors=None, cluster_method=""):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="CLUSTERS",bg=CARD,fg=COL_CLUSTER,font=FONT_H3,justify="center").pack(pady=(8,2),padx=6)
    if cluster_method:
        tk.Label(parent,text=cluster_method,bg=CARD,fg=FG_DIM,font=("Courier New",7),wraplength=180,justify="center").pack(padx=4,pady=(0,2))
    _divider(parent,COL_CLUSTER)
    if df is None or "cluster" not in df.columns:
        tk.Label(parent,text="Run analysis first",bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=170).pack(pady=8); return
    total=len(df)
    for cl in sorted(df["cluster"].unique()):
        cnt=int((df["cluster"]==cl).sum()); col=hex_colors[int(cl)%len(hex_colors)] if hex_colors else ACCENT
        row=tk.Frame(parent,bg=CARD2,padx=5,pady=3,highlightbackground=col,highlightthickness=1)
        row.pack(fill="x",padx=5,pady=2)
        tk.Canvas(row,bg=col,width=10,height=10,highlightthickness=0).pack(side="left",padx=(0,5))
        txt=tk.Frame(row,bg=CARD2); txt.pack(side="left",fill="x",expand=True)
        tk.Label(txt,text=f"Cluster {cl}",bg=CARD2,fg=col,font=("Courier New",8,"bold"),anchor="w").pack(anchor="w")
        tk.Label(txt,text=f"{cnt:,} wells  ({cnt/total*100:.1f}%)",bg=CARD2,fg=FG_MID,font=("Courier New",7),anchor="w").pack(anchor="w")
        _mini_bar(parent,"",cnt/total,col)

def build_anomaly_legend(parent, n_normal=0, n_anomaly=0):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="ANOMALY",bg=CARD,fg=COL_ANOMALY,font=FONT_H3,justify="center").pack(pady=(8,2),padx=6)
    _divider(parent,COL_ANOMALY)
    total=max(n_normal+n_anomaly,1)
    for label,cnt,col in [("● Normal",n_normal,NORMAL_C),("◆ Anomalous",n_anomaly,ANOMALY_C)]:
        row=tk.Frame(parent,bg=CARD2,padx=5,pady=3,highlightbackground=col,highlightthickness=1)
        row.pack(fill="x",padx=5,pady=2)
        tk.Label(row,text=label,bg=CARD2,fg=col,font=("Courier New",8,"bold")).pack(anchor="w")
        tk.Label(row,text=f"{cnt:,}  ({cnt/total*100:.1f}%)",bg=CARD2,fg=FG_MID,font=("Courier New",7)).pack(anchor="w")
        _mini_bar(parent,"",cnt/total,col)

def build_event_legend(parent, event_result):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="EVENT LEGEND",bg=CARD,fg=COL_EVENTS,font=FONT_H3,justify="center").pack(pady=(8,2),padx=6)
    _divider(parent,COL_EVENTS)
    if not event_result["has_events"]:
        tk.Label(parent,text="No event column",bg=CARD,fg=WARN,font=FONT_SM,wraplength=170).pack(pady=8); return
    counts=event_result["event_counts"]; total=max(sum(counts.values()),1)
    def _c(label):
        ll=label.lower()
        for k,c in EVENT_COLOURS.items():
            if k in ll: return c
        return CLUSTER_PALETTE[hash(label)%len(CLUSTER_PALETTE)]
    for label,cnt in sorted(counts.items(),key=lambda x:-x[1]):
        col=_c(label)
        row=tk.Frame(parent,bg=CARD2,padx=5,pady=3,highlightbackground=BORDER,highlightthickness=1)
        row.pack(fill="x",padx=5,pady=2)
        tk.Canvas(row,bg=col,width=10,height=10,highlightthickness=0).pack(side="left",padx=(0,5))
        txt=tk.Frame(row,bg=CARD2); txt.pack(side="left",fill="x",expand=True)
        tk.Label(txt,text=label[:22],bg=CARD2,fg=col,font=("Courier New",8,"bold"),anchor="w").pack(anchor="w")
        tk.Label(txt,text=f"{cnt:,}  ({cnt/total*100:.1f}%)",bg=CARD2,fg=FG_MID,font=("Courier New",7),anchor="w").pack(anchor="w")
        _mini_bar(parent,"",cnt/total,col)

def _fill_toolbar(bar, mpl_canvas, tab_name):
    tk.Label(bar,text=" TOOLS:",bg="#080e1a",fg=FG_DIM,font=FONT_XS).pack(side="left",padx=(6,2))
    nav_frame=tk.Frame(bar,bg="#080e1a"); nav_frame.pack(side="left",padx=2)
    tb=NavigationToolbar2Tk(mpl_canvas,nav_frame); tb.config(bg="#080e1a")
    for child in tb.winfo_children():
        try: child.config(bg="#080e1a",fg=FG_MID,activebackground=CARD2,activeforeground=FG,relief="flat",bd=0,font=FONT_XS)
        except: pass
    tb.update()
    tk.Frame(bar,bg=BORDER,width=1).pack(side="left",fill="y",padx=8,pady=3)
    def _save_png():
        fig=active_figures.get(tab_name)
        if not fig: messagebox.showwarning("Export","Run analysis first."); return
        path=filedialog.asksaveasfilename(defaultextension=".png",initialfile=f"cbm_{tab_name}_{ts()}.png",
                                          filetypes=[("PNG","*.png"),("SVG","*.svg")])
        if not path: return
        fig.savefig(path,dpi=180,bbox_inches="tight",facecolor=fig.get_facecolor())
        set_status(f"Saved: {os.path.basename(path)}",SUCCESS)
    _icon_btn(bar,"Save Chart",_save_png,bg="#1e3a8a",fg="#93c5fd")

def draw_plot(tab_frame, fig, tab_name, legend_builder=None, legend_kwargs=None):
    for w in tab_frame.winfo_children(): w.destroy()
    active_figures[tab_name]=fig
    try:
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            fig.tight_layout(pad=2.5,rect=[0.03,0.03,0.97,0.95])
    except: pass
    fig.patch.set_facecolor("#0d1117")
    bar=tk.Frame(tab_frame,bg="#080e1a",pady=4,highlightbackground=BORDER,highlightthickness=1)
    bar.pack(side="top",fill="x")
    row=tk.Frame(tab_frame,bg=BG_MID); row.pack(side="top",fill="both",expand=True)
    chart_frame=tk.Frame(row,bg=BG_MID); chart_frame.pack(side="left",fill="both",expand=True)
    leg_frame=tk.Frame(row,bg=CARD,width=195,highlightbackground=BORDER,highlightthickness=1)
    leg_frame.pack(side="right",fill="y",padx=(2,6),pady=6); leg_frame.pack_propagate(False)
    mpl_canvas=FigureCanvasTkAgg(fig,master=chart_frame); mpl_canvas.draw()
    cw=mpl_canvas.get_tk_widget(); cw.config(bg="#0d1117",highlightthickness=0)
    cw.pack(fill="both",expand=True,padx=2,pady=2)
    _fill_toolbar(bar,mpl_canvas,tab_name)
    if legend_builder: legend_builder(leg_frame,**(legend_kwargs or {}))


# ══════════════════════════════════════════════════════════════════════════════
# PIPELINE
# ══════════════════════════════════════════════════════════════════════════════
def start_pipeline():
    if raw_data is None:
        messagebox.showwarning("No Data","Please upload a dataset first."); return
    set_status("Running analysis…",WARN); progress_start()
    threading.Thread(target=run_pipeline,daemon=True).start()

def run_pipeline():
    global active_df, active_X, active_xcols, active_well_df
    global active_anomaly_result, active_event_result
    global active_cluster_method, active_insights_text

    try:
        import pandas as pd

        selected_cols = get_selected_features()
        if not selected_cols:
            selected_cols = list(raw_data.select_dtypes(include="number").columns)
        if not selected_cols:
            app.after(0, lambda: set_status("No features selected.", WARN))
            app.after(0, progress_stop); return

        app.after(0, lambda: set_status(
            f"Using {len(selected_cols)} features — aggregating by well…", WARN))

        well_id_col = _find_well_id_column(raw_data)
        wdf = raw_data.copy().reset_index(drop=True)

        num_sel = [c for c in selected_cols
                   if c in wdf.columns and pd.api.types.is_numeric_dtype(wdf[c])]
        cat_sel = [c for c in selected_cols
                   if c in wdf.columns and not pd.api.types.is_numeric_dtype(wdf[c])]

        n_rows = len(wdf)
        n_unique_wells = wdf[well_id_col].nunique() if well_id_col and well_id_col in wdf.columns else n_rows

        app.after(0, lambda: set_status(
            f"Aggregating {n_rows:,} records → {n_unique_wells:,} wells…", WARN))

        well_df, num_agg_cols, cat_agg_cols = aggregate_by_well(wdf, selected_cols, well_id_col)

        if not num_agg_cols and not cat_agg_cols:
            app.after(0, lambda: set_status("No columns could be aggregated.", ERR))
            app.after(0, progress_stop); return

        all_agg = num_agg_cols + cat_agg_cols
        well_df_clean = well_df.copy()
        if all_agg:
            num_ok = (well_df_clean[num_agg_cols].notna().any(axis=1)
                      if num_agg_cols else pd.Series(False, index=well_df_clean.index))
            cat_ok = (well_df_clean[cat_agg_cols].apply(
                          lambda col: col.fillna("").astype(str).ne("") & col.fillna("").astype(str).ne("MISSING")
                      ).any(axis=1)
                      if cat_agg_cols else pd.Series(False, index=well_df_clean.index))
            mask_valid = num_ok | cat_ok
            well_df_clean = well_df_clean[mask_valid].reset_index(drop=True)

        n_wells = len(well_df_clean)
        if n_wells < 2:
            app.after(0, lambda: set_status(
                f"Only {n_wells} well(s) after cleaning — need at least 2.", ERR))
            app.after(0, progress_stop); return

        app.after(0, lambda: set_status(f"Clustering {n_wells:,} well profiles…", WARN))

        weight_map = {c: round(100/max(len(num_sel),1),2) for c in num_sel}

        n_clusters = min(cluster_var.get(), n_wells)
        cluster_labels, inertia, cluster_method = run_clustering_robust(
            well_df_clean, num_agg_cols, cat_agg_cols, n_clusters, weight_map)

        active_cluster_method = cluster_method
        well_df_clean = well_df_clean.copy()
        well_df_clean["cluster"] = cluster_labels

        if well_id_col and well_id_col in wdf.columns and well_id_col in well_df_clean.columns:
            cl_map = dict(zip(well_df_clean[well_id_col], cluster_labels))
            wdf["cluster"] = wdf[well_id_col].map(cl_map).fillna(-1).astype(int)
        else:
            wdf["cluster"] = 0

        app.after(0, lambda: set_status("Detecting anomalies…", WARN))

        X_anom = build_anomaly_X(well_df_clean, num_agg_cols, cat_agg_cols, weight_map)

        method  = anomaly_method_var.get()
        contam  = anomaly_contam_var.get() / 100.0
        a_labels_well, a_pct, a_idx, a_name, a_scores = detect_anomalies(
            X_anom, contamination=contam, method=method)
        well_df_clean["anomaly"] = a_labels_well

        if well_id_col and well_id_col in wdf.columns and well_id_col in well_df_clean.columns:
            anom_map = dict(zip(well_df_clean[well_id_col], a_labels_well))
            wdf["anomaly"] = wdf[well_id_col].map(anom_map).fillna(1).astype(int)
        else:
            wdf["anomaly"] = 1

        a_labels_rec = wdf["anomaly"].values
        event_result = analyse_events(wdf, a_labels_rec)

        z_thr = 2.5
        well_explanations = explain_anomalous_wells(
            well_df_clean, num_agg_cols, cat_agg_cols,
            a_labels_well, a_scores, well_id_col, z_threshold=z_thr)

        active_df             = wdf
        active_X              = X_anom
        active_xcols          = num_sel
        active_well_df        = well_df_clean
        active_anomaly_result = {
            "labels": a_labels_well, "pct": a_pct, "indices": a_idx,
            "scores": a_scores,
            "detector_name": f"{a_name}  [well-profile detection]",
            "n_anomaly": int((a_labels_well == -1).sum()),
            "n_normal":  int((a_labels_well ==  1).sum()),
            "explanations": well_explanations,
        }
        active_event_result   = event_result

        hx = cluster_hex(n_clusters)
        insights = _generate_insights(
            wdf, well_df_clean, num_sel, selected_cols, n_rows, n_unique_wells,
            n_wells, hx, active_anomaly_result, event_result, weight_map,
            inertia, event_result, well_id_col,
            cluster_method, cat_sel, well_explanations)
        active_insights_text = insights

        app.after(0, lambda: refresh_ui(
            wdf, well_df_clean, X_anom,
            num_agg_cols + cat_agg_cols, num_sel,
            hx, insights, n_rows, n_unique_wells,
            active_anomaly_result, event_result, cluster_method))
        app.after(0, lambda: set_status(
            f"✔  Analysis complete  [{cluster_method}]", SUCCESS))
        app.after(0, progress_stop)
        app.after(0, lambda: export_btn.config(state="normal"))

    except Exception as e:
        traceback.print_exc()
        app.after(0, lambda err=str(e): set_status(f"Error: {err}", ERR))
        app.after(0, progress_stop)


def refresh_ui(wdf, well_df, X_anom, agg_cols, num_cols,
               hx, insights, n_rows, n_unique_wells,
               anomaly_result, event_result, cluster_method):

    draw_plot(cluster_tab,
              plot_well_clusters(well_df, agg_cols, hx, cluster_method), "clusters",
              legend_builder=build_cluster_legend,
              legend_kwargs={"df": well_df, "hex_colors": hx, "cluster_method": cluster_method})
    draw_plot(pca_tab,
              plot_pca(X_anom, well_df["cluster"].values, hx, cluster_method), "pca",
              legend_builder=build_cluster_legend,
              legend_kwargs={"df": well_df, "hex_colors": hx, "cluster_method": cluster_method})
    draw_plot(hidden_tab,
              plot_hidden_patterns(X_anom, well_df["anomaly"].values,
                                   agg_cols, anomaly_result["detector_name"]),
              "hidden_patterns",
              legend_builder=build_anomaly_legend,
              legend_kwargs={"n_normal": anomaly_result["n_normal"],
                             "n_anomaly": anomaly_result["n_anomaly"]})
    draw_plot(event_tab,
              plot_event_chart(wdf, event_result), "event_summary",
              legend_builder=build_event_legend,
              legend_kwargs={"event_result": event_result})

    stat_widgets["Unique Wells"].set(f"{n_unique_wells:,}")
    stat_widgets["Total Records"].set(f"{n_rows:,}")
    stat_widgets["Clusters"].set(str(well_df["cluster"].nunique()))
    an = anomaly_result["n_anomaly"]; ap = anomaly_result["pct"]
    stat_widgets["Anomaly Wells"].set(f"{an:,}\n({ap:.1f}%)")
    if event_result["has_events"]:
        stat_widgets["Active Recs"].set(f"{event_result['n_active']:,}")
        stat_widgets["Inactive Recs"].set(f"{event_result['n_inactive']:,}")
        stat_widgets["Event Types"].set(str(len(event_result["event_counts"])))
    else:
        stat_widgets["Active Recs"].set("N/A")
        stat_widgets["Inactive Recs"].set("N/A")
        stat_widgets["Event Types"].set("N/A")
    stat_widgets["Abnormal Recs"].set(f"{event_result['n_abnormal']:,}")

    update_event_panel(event_result)
    explain_text.config(state="normal")
    explain_text.delete("1.0","end")
    explain_text.insert("end", insights)
    explain_text.config(state="disabled")


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
    top = sorted(counts.items(), key=lambda x: -x[1])[:6]; keys = list(ev_count_vars.keys())
    for i, k in enumerate(keys):
        if i < len(top):
            label, cnt = top[i]; ev_count_vars[k].set(f"{cnt:,}"); ev_count_lbls[k].config(text=label[:24])
        else:
            ev_count_vars[k].set("—"); ev_count_lbls[k].config(text="—")
    ev_active_var.set(f"{event_result['n_active']:,}"); ev_inactive_var.set(f"{event_result['n_inactive']:,}")
    ev_abnormal_var.set(f"{event_result['n_abnormal']:,}"); ev_op_anom_var.set("N/A")


# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════
def export_results():
    if active_df is None: messagebox.showwarning("Export","Run analysis first."); return
    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx", initialfile=f"cbm_results_{ts()}.xlsx",
        filetypes=[("Excel","*.xlsx"),("CSV","*.csv")])
    if not path: return
    try:
        import pandas as pd; ext = os.path.splitext(path)[1].lower()
        if ext == ".xlsx":
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                active_df.to_excel(writer, sheet_name="All Records", index=False)
                if active_well_df is not None:
                    active_well_df.to_excel(writer, sheet_name="Well Profiles", index=False)
                if active_well_df is not None and "anomaly" in active_well_df.columns:
                    active_well_df[active_well_df["anomaly"] == -1].to_excel(
                        writer, sheet_name="Anomaly Wells", index=False)
                active_well_df.groupby("cluster").size().reset_index(name="well_count").to_excel(
                    writer, sheet_name="Cluster Summary", index=False)
                if active_anomaly_result and active_anomaly_result.get("explanations"):
                    expl_rows = []
                    for e in active_anomaly_result["explanations"]:
                        for p in e["flagged_params"]:
                            expl_rows.append({
                                "Well": e["well"],
                                "Anomaly_Score": e["anomaly_score"],
                                "Parameter": p["param"],
                                "Value": p["value"],
                                "Z_Score": p.get("z_score",""),
                                "Direction": p["direction"],
                                "Reason": p["reason"]
                            })
                    if expl_rows:
                        pd.DataFrame(expl_rows).to_excel(
                            writer, sheet_name="Anomaly Explanations", index=False)
                report_txt = explain_text.get("1.0","end").strip()
                pd.DataFrame({"Report": report_txt.split("\n")}).to_excel(
                    writer, sheet_name="Insights Report", index=False)
            set_status(f"✔  Exported: {os.path.basename(path)}", SUCCESS)
            messagebox.showinfo("Export Complete", f"Saved: {path}")
        else:
            _save_df(active_df, path); set_status(f"✔  Exported: {os.path.basename(path)}", SUCCESS)
    except Exception as e:
        messagebox.showerror("Export Error", str(e)); set_status(f"Export failed: {e}", ERR)


# ══════════════════════════════════════════════════════════════════════════════
# ENHANCED INSIGHTS REPORT
# ══════════════════════════════════════════════════════════════════════════════
def _generate_insights(wdf, well_df, num_cols, selected_cols, n_rows, n_unique_wells,
                       n_wells, hx, anomaly_result, event_result, weight_map,
                       inertia, rpm_result, well_id_col,
                       cluster_method="", cat_sel=None, well_explanations=None):
    import textwrap
    W="="*60; D="-"*60; B=""
    n_anom=anomaly_result["n_anomaly"]; a_pct=anomaly_result["pct"]
    sev_label=("LOW" if a_pct<2 else "MODERATE" if a_pct<8 else "ELEVATED" if a_pct<20 else "HIGH")
    cat_cols = cat_sel or []
    lines=[W,"  CBM AI ANALYSIS REPORT   v7.0",
           f"  Generated : {datetime.datetime.now().strftime('%Y-%m-%d  %H:%M:%S')}",W,B,
           "  DATASET SUMMARY","  "+D,
           f"  Unique wells       : {n_unique_wells:,}",
           f"  Well profiles      : {n_wells:,}  (clustered)",
           f"  Total records      : {n_rows:,}",
           f"  Well ID column     : {well_id_col or 'Not detected'}",
           f"  Selected features  : {len(selected_cols)}",
           f"    Numeric          : {len(num_cols)}  →  {', '.join(num_cols)}",
           f"    Categorical      : {len(cat_cols)}  →  {', '.join(cat_cols) if cat_cols else 'none'}",
           f"  Clusters           : {well_df['cluster'].nunique() if 'cluster' in well_df.columns else '—'}",
           f"  Cluster algorithm  : {cluster_method}",
           f"  Inertia/cost proxy : {inertia:.4f}",B]

    lines+=[W,"  ANOMALY DETECTION  (well-level)","  "+D,
            f"  Algorithm          : {anomaly_result['detector_name']}",
            f"  Normal wells       : {anomaly_result['n_normal']:,}",
            f"  Anomalous wells    : {n_anom:,}  ({a_pct:.1f}%)",
            f"  Overall Severity   : {sev_label}",B]

    lines += [W, "  HIDDEN PATTERN ANALYSIS  —  WHY EACH WELL IS ANOMALOUS", "  "+D]
    if not well_explanations:
        lines += ["  No anomalous wells detected.", B]
    else:
        lines += [
            f"  Total anomalous wells: {len(well_explanations)}",
            f"  Detection method: {anomaly_result['detector_name']}",
            f"  Z-score threshold for parameter flagging: 2.5σ",
            B
        ]
        for rank, expl in enumerate(well_explanations, 1):
            lines += [
                f"  ┌─ ANOMALY #{rank} ─────────────────────────────────────────",
                f"  │  Well ID       : {expl['well']}",
                f"  │  Anomaly Score : {expl['anomaly_score']:.5f}  "
                f"(more negative = more anomalous)",
                f"  │  Flagged Params: {expl['n_flagged']}",
            ]
            if not expl["flagged_params"]:
                lines += [
                    "  │  Note: This well is anomalous in the combined",
                    "  │        multi-dimensional feature space but no",
                    "  │        single parameter exceeds the 2.5σ threshold.",
                    "  │        Check the PCA / Hidden Patterns plot for",
                    "  │        visual isolation from the normal cluster.",
                ]
            else:
                for pi, p in enumerate(expl["flagged_params"], 1):
                    z_str = f"  z={p['z_score']:+.2f}σ" if p.get("z_score") is not None else ""
                    dir_str = f"  [{p['direction']}]"
                    lines += [
                        f"  │",
                        f"  │  Parameter {pi}: {p['param']}",
                        f"  │    Value    : {p['value']}{z_str}{dir_str}",
                    ]
                    if p.get("mean") is not None:
                        lines += [
                            f"  │    Fleet avg: {p['mean']:.4f}  ±{p['std']:.4f}",
                            f"  │    P10–P90  : {p['p10']:.4f} – {p['p90']:.4f}",
                        ]
                    reason_lines = textwrap.wrap(p["reason"], 52)
                    for i, rl in enumerate(reason_lines):
                        lines.append(f"  │    {'WHY: ' if i==0 else '      '}{rl}")
            lines += ["  └" + "─"*54, B]

    if "cluster" in well_df.columns:
        lines += [W, "  CLUSTER SUMMARY", "  "+D]
        for cl in sorted(well_df["cluster"].unique()):
            grp = well_df[well_df["cluster"] == cl]
            n_cl = len(grp)
            anom_in_cl = int((grp["anomaly"] == -1).sum()) if "anomaly" in grp.columns else 0
            lines += [
                f"  Cluster {cl}:  {n_cl:,} wells  |  {anom_in_cl} anomalous "
                f"({anom_in_cl/max(n_cl,1)*100:.1f}%)"
            ]
            import pandas as pd
            top_cols = [c for c in num_cols if f"{c}__mean" in grp.columns][:4]
            for c in top_cols:
                vals = grp[f"{c}__mean"].dropna()
                if len(vals):
                    lines.append(f"    {c[:28]:<28} mean={float(vals.mean()):.4f}")
            lines.append(B)

    if event_result["has_events"]:
        lines += [W, "  EVENT SUMMARY", "  "+D,
                  f"  Event column: '{event_result['event_col']}'",
                  f"  Active records  : {event_result['n_active']:,}",
                  f"  Inactive records: {event_result['n_inactive']:,}",
                  f"  Abnormal records: {event_result['n_abnormal']:,}",B]

    lines += [W, "  END OF REPORT  v7.0", W]
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# AI CHATBOT — LOGIC (unchanged from v7.0)
# ══════════════════════════════════════════════════════════════════════════════
CHATBOT_SYSTEM = """You are an expert AI assistant embedded inside the CBM AI Analytics Platform v7.0,
a Coalbed Methane (CBM) well analytics desktop application built in Python/Tkinter.

You help the user understand:
1. HOW THE APP WORKS — features, tabs, controls, pipeline steps
2. ANALYSIS RESULTS — clusters, anomalies, events, hidden patterns
3. CBM DOMAIN KNOWLEDGE — coalbed methane production, well health, common failure modes
4. DATA INTERPRETATION — what anomaly scores mean, how clustering works, etc.
5. TROUBLESHOOTING — why something might not work, what data formats are needed

Key app facts:
- Tabs: Well Clusters, PCA, Hidden Patterns, Events
- Clustering algorithm chain: KPrototypes → KModes → Gower/Agglomerative → Ward → KMeans
- Anomaly detection: Isolation Forest or Local Outlier Factor on well profiles
- Data: aggregates raw records per well (mean/std/min/max for numeric, mode for categorical)
- Export: Excel with sheets — All Records, Well Profiles, Anomaly Wells, Cluster Summary, Anomaly Explanations, Insights Report
- The Insights Report shows per-well anomaly explanations with z-scores and parameter-level reasoning

If the user provides analysis results context (appended below), use it to give specific answers.
Be concise but thorough. Use plain English. Format with bullet points when listing multiple items.
"""

def _build_context_for_chatbot():
    parts = []
    if raw_data is not None:
        parts.append(f"Dataset loaded: {len(raw_data):,} rows × {len(raw_data.columns)} columns.")
    if active_well_df is not None:
        n_wells = len(active_well_df)
        n_anom = int((active_well_df.get("anomaly", 0) == -1).sum()) if "anomaly" in active_well_df.columns else 0
        n_cl = active_well_df["cluster"].nunique() if "cluster" in active_well_df.columns else 0
        parts.append(f"Analysis complete: {n_wells} well profiles, {n_cl} clusters, {n_anom} anomalous wells.")
        parts.append(f"Cluster algorithm: {active_cluster_method}")
    if active_anomaly_result:
        parts.append(f"Anomaly detector: {active_anomaly_result['detector_name']}")
        parts.append(f"Anomaly rate: {active_anomaly_result['pct']:.1f}%")
    if active_insights_text and len(active_insights_text) > 50:
        parts.append("\n\nFULL INSIGHTS REPORT:\n" + active_insights_text[:6000])
    return "\n".join(parts)

def send_chat_message():
    user_msg = chat_input.get("1.0", "end").strip()
    if not user_msg:
        return
    chat_input.delete("1.0", "end")
    _append_chat("You", user_msg, "#60a5fa")
    threading.Thread(target=_call_claude_api, args=(user_msg,), daemon=True).start()

def _append_chat(sender, message, color):
    def _do():
        chat_display.config(state="normal")
        chat_display.insert("end", f"\n{sender}:\n", f"sender_{color.lstrip('#')}")
        chat_display.insert("end", f"{message}\n")
        chat_display.tag_configure(f"sender_{color.lstrip('#')}",
                                   foreground=color, font=FONT_H3)
        chat_display.see("end")
        chat_display.config(state="disabled")
    app.after(0, _do)

def _call_claude_api(user_msg):
    global chatbot_history
    _append_chat("CBM·AI", "Thinking…", COL_CHATBOT)

    context = _build_context_for_chatbot()
    system_msg = CHATBOT_SYSTEM
    if context:
        system_msg += f"\n\nCURRENT ANALYSIS CONTEXT:\n{context}"

    chatbot_history.append({"role": "user", "content": user_msg})
    history_trimmed = chatbot_history[-20:]

    payload = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 1000,
        "system": system_msg,
        "messages": history_trimmed
    }).encode("utf-8")

    try:
        req = urllib.request.Request(
            "https://api.anthropic.com/v1/messages",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "anthropic-version": "2023-06-01",
                "x-api-key": ""   # API key handled by proxy
            },
            method="POST"
        )
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            reply = "".join(
                block.get("text", "") for block in data.get("content", [])
                if block.get("type") == "text"
            ).strip()
            if not reply:
                reply = "(No response from model)"
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        reply = f"API Error {e.code}: {body[:300]}"
    except Exception as e:
        reply = f"Connection error: {e}\n\nMake sure you have internet access and a valid API key."

    chatbot_history.append({"role": "assistant", "content": reply})

    def _update():
        chat_display.config(state="normal")
        content = chat_display.get("1.0", "end")
        thinking_tag = "\nCBM·AI:\nThinking…\n"
        if thinking_tag in content:
            idx = content.rfind(thinking_tag)
            start = f"1.0 + {idx} chars"
            end   = f"1.0 + {idx + len(thinking_tag)} chars"
            chat_display.delete(start, end)
        chat_display.insert("end", f"\nCBM·AI:\n", "sender_e879f9")
        chat_display.insert("end", f"{reply}\n")
        chat_display.tag_configure("sender_e879f9", foreground=COL_CHATBOT, font=FONT_H3)
        chat_display.see("end")
        chat_display.config(state="disabled")
    app.after(0, _update)


def clear_chat():
    global chatbot_history
    chatbot_history = []
    chat_display.config(state="normal")
    chat_display.delete("1.0", "end")
    chat_display.insert("end",
        "CBM·AI Assistant ready.\n"
        "Ask me anything about the app, your data, or the analysis results.\n\n"
        "Examples:\n"
        "  • Why is well X anomalous?\n"
        "  • What does the cluster analysis tell me?\n"
        "  • How does Isolation Forest work?\n"
        "  • What file formats are supported?\n"
    )
    chat_display.config(state="disabled")


# ══════════════════════════════════════════════════════════════════════════════
# FLOATING CHATBOT PANEL
# ══════════════════════════════════════════════════════════════════════════════
_chat_popup = None
_chat_popup_visible = False
_drag_offset_x = 0
_drag_offset_y = 0
_pulse_state = [0]
_pulse_job = [None]

def _pulse_icon():
    """Animate the chatbot icon with a subtle glow cycle."""
    if _chat_popup_visible:
        chat_fab_btn.config(bg="#7c3aed", fg="#fff")
        if _pulse_job[0]:
            app.after_cancel(_pulse_job[0])
        return
    colors = ["#4f46e5", "#6d28d9", "#7c3aed", "#6d28d9"]
    _pulse_state[0] = (_pulse_state[0] + 1) % len(colors)
    chat_fab_btn.config(bg=colors[_pulse_state[0]])
    _pulse_job[0] = app.after(800, _pulse_icon)

def _start_drag(event):
    global _drag_offset_x, _drag_offset_y
    _drag_offset_x = event.x_root - _chat_popup.winfo_x()
    _drag_offset_y = event.y_root - _chat_popup.winfo_y()

def _do_drag(event):
    x = event.x_root - _drag_offset_x
    y = event.y_root - _drag_offset_y
    _chat_popup.geometry(f"+{x}+{y}")

def toggle_chat_popup():
    global _chat_popup, _chat_popup_visible

    if _chat_popup_visible and _chat_popup and _chat_popup.winfo_exists():
        _chat_popup.withdraw()
        _chat_popup_visible = False
        chat_fab_btn.config(text="★ AI", bg="#4f46e5")
        _pulse_job[0] = app.after(800, _pulse_icon)
        return

    # Create popup if it doesn't exist yet
    if _chat_popup is None or not _chat_popup.winfo_exists():
        _build_chat_popup()

    # Position above the FAB button
    app.update_idletasks()
    fab_x = chat_fab_btn.winfo_rootx()
    fab_y = chat_fab_btn.winfo_rooty()
    popup_w, popup_h = 390, 540
    px = max(10, fab_x - popup_w + chat_fab_btn.winfo_width())
    py = max(10, fab_y - popup_h - 8)
    _chat_popup.geometry(f"{popup_w}x{popup_h}+{px}+{py}")
    _chat_popup.deiconify()
    _chat_popup.lift()
    _chat_popup_visible = True
    chat_fab_btn.config(text="✕ AI", bg="#7c3aed")
    if _pulse_job[0]:
        app.after_cancel(_pulse_job[0])

def _build_chat_popup():
    global _chat_popup, chat_display, chat_input

    popup = tk.Toplevel(app)
    popup.overrideredirect(True)          # no OS titlebar
    popup.configure(bg="#0a0f1e")
    popup.attributes("-topmost", True)
    popup.resizable(False, False)

    # Drop shadow effect via outer border
    outer = tk.Frame(popup, bg="#e879f9", padx=1, pady=1)
    outer.pack(fill="both", expand=True)
    inner = tk.Frame(outer, bg="#0a0f1e")
    inner.pack(fill="both", expand=True)

    # ── Title bar (draggable) ──────────────────────────────────────────────
    title_bar = tk.Frame(inner, bg="#150a2e", height=36, cursor="fleur")
    title_bar.pack(fill="x"); title_bar.pack_propagate(False)

    tk.Label(title_bar, text="★", bg="#150a2e", fg=COL_CHATBOT,
             font=("Georgia", 13, "bold")).pack(side="left", padx=(10,4), pady=4)
    tk.Label(title_bar, text="CBM·AI Assistant", bg="#150a2e", fg=COL_CHATBOT,
             font=("Georgia", 10, "bold")).pack(side="left", pady=4)

    def _close_popup():
        global _chat_popup_visible
        popup.withdraw()
        _chat_popup_visible = False
        chat_fab_btn.config(text="★ AI", bg="#4f46e5")
        _pulse_job[0] = app.after(800, _pulse_icon)

    close_btn = tk.Button(title_bar, text="✕", command=_close_popup,
                          bg="#150a2e", fg="#94a3b8",
                          activebackground="#1e1040", activeforeground="#f87171",
                          font=("Courier New", 11, "bold"), relief="flat",
                          bd=0, cursor="hand2", padx=10, pady=0)
    close_btn.pack(side="right", fill="y")

    clear_btn = tk.Button(title_bar, text="Clear", command=lambda: _do_clear(),
                          bg="#150a2e", fg="#4a6080",
                          activebackground="#1e1040", activeforeground=FG,
                          font=("Courier New", 8), relief="flat",
                          bd=0, cursor="hand2", padx=8)
    clear_btn.pack(side="right", fill="y", pady=4)

    # Drag bindings
    title_bar.bind("<ButtonPress-1>", _start_drag)
    title_bar.bind("<B1-Motion>", _do_drag)
    for child in title_bar.winfo_children():
        if isinstance(child, tk.Label):
            child.bind("<ButtonPress-1>", _start_drag)
            child.bind("<B1-Motion>", _do_drag)

    # Separator
    tk.Frame(inner, bg=COL_CHATBOT, height=1).pack(fill="x")

    # ── Chat display ───────────────────────────────────────────────────────
    disp_frame = tk.Frame(inner, bg=CARD)
    disp_frame.pack(fill="both", expand=True, padx=0, pady=0)

    chat_vsb2 = tk_ttk.Scrollbar(disp_frame, orient="vertical")
    chat_display = tk.Text(disp_frame, bg="#0c1424", fg=FG,
                           font=("Courier New", 9), relief="flat", bd=0,
                           padx=10, pady=8, wrap="word",
                           yscrollcommand=chat_vsb2.set,
                           highlightthickness=0, state="disabled",
                           selectbackground=ACCENT)
    chat_vsb2.config(command=chat_display.yview)
    chat_vsb2.pack(side="right", fill="y")
    chat_display.pack(side="left", fill="both", expand=True)

    # ── Separator ──────────────────────────────────────────────────────────
    tk.Frame(inner, bg="#1e1040", height=1).pack(fill="x")

    # ── Input area ─────────────────────────────────────────────────────────
    input_area = tk.Frame(inner, bg="#0e0b1e", padx=8, pady=6)
    input_area.pack(fill="x")

    chat_input = tk.Text(input_area, height=3, bg="#131b2a", fg=FG,
                         insertbackground=COL_CHATBOT, font=("Courier New", 9),
                         relief="flat", bd=0, wrap="word",
                         highlightbackground=COL_CHATBOT, highlightthickness=1,
                         padx=6, pady=4)
    chat_input.pack(side="left", fill="x", expand=True, padx=(0, 6))

    def _send():
        msg = chat_input.get("1.0", "end").strip()
        if not msg: return
        chat_input.delete("1.0", "end")
        _append_chat("You", msg, "#60a5fa")
        threading.Thread(target=_call_claude_api, args=(msg,), daemon=True).start()

    def _enter_key(event):
        if not (event.state & 0x1):  # no Shift
            _send()
            return "break"

    chat_input.bind("<Return>", _enter_key)

    send_btn2 = tk.Button(input_area, text="Send\n↵", command=_send,
                          bg="#4f46e5", fg="#fff",
                          activebackground="#6d28d9", activeforeground="#fff",
                          font=("Courier New", 8, "bold"), relief="flat",
                          bd=0, cursor="hand2", padx=10, pady=6)
    send_btn2.pack(side="left")

    # Hint
    tk.Label(input_area, text="Shift+Enter = newline", bg="#0e0b1e", fg="#2d3f5e",
             font=("Courier New", 7)).pack(side="left", padx=6)

    # ESC to close
    popup.bind("<Escape>", lambda e: _close_popup())

    def _do_clear():
        global chatbot_history
        chatbot_history = []
        chat_display.config(state="normal")
        chat_display.delete("1.0", "end")
        chat_display.insert("end",
            "CBM·AI Assistant ready.\n"
            "Ask me anything about the app, your data, or the analysis results.\n\n"
            "Examples:\n"
            "  • Why is well X anomalous?\n"
            "  • What does the cluster analysis tell me?\n"
            "  • How does Isolation Forest work?\n"
            "  • What file formats are supported?\n"
        )
        chat_display.config(state="disabled")

    _do_clear()
    _chat_popup = popup
    popup.withdraw()  # hidden initially


# ══════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
if HAVE_TTKBS:
    app = ttk.Window(themename="darkly")
else:
    app = tk.Tk(); app.configure(bg=BG_DEEP)

app.title("CBM AI Analytics Platform  v7.0  —  Well Clustering · Anomaly Detection · AI Chatbot")
app.geometry("1820x1020"); app.configure(bg=BG_DEEP); app.minsize(1280,720)

top_bar=tk.Frame(app,bg="#040810",height=46)
top_bar.pack(fill="x",side="top"); top_bar.pack_propagate(False)
brand=tk.Frame(top_bar,bg="#040810"); brand.pack(side="left",padx=16,fill="y")
tk.Label(brand,text="CBM·AI",bg="#040810",fg=COL_DATASET,font=("Georgia",17,"bold")).pack(side="left",padx=(0,10))
tk.Label(brand,text="Coalbed Methane Analytics  v7.0  —  Clustering · Hidden Patterns · AI Chatbot",
         bg="#040810",fg=FG_DIM,font=FONT_SM).pack(side="left")
tk.Label(top_bar,
         text="KPrototypes · Gower · Ward · KMeans  ·  Isolation Forest / LOF  ·  Claude-powered Chatbot",
         bg="#040810",fg=FG_DIM,font=FONT_XS).pack(side="right",padx=14)
tk.Frame(app,bg=BORDER,height=1).pack(fill="x")

root_pane=tk.PanedWindow(app,orient="horizontal",bg=BG_DEEP,sashwidth=4,sashrelief="flat",sashpad=0)
root_pane.pack(fill="both",expand=True)

# ══════════════════════════════════════════════════════════════════════════════
# LEFT SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
sb_outer,sidebar=make_scrollable(root_pane,bg=SIDEBAR)
root_pane.add(sb_outer,width=440,minsize=400)
tk.Frame(sidebar,bg=SIDEBAR,height=8).pack()

# ① DATASET
ds_body=make_section_card(sidebar,"① DATASET",COL_DATASET)
tk.Label(ds_body,text="CSV, TSV, XLSX, XLS, ODS, JSON, Parquet, HDF5, Pickle, MS Access .mdb/.accdb",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=420,justify="left").pack(anchor="w",pady=(0,4))
make_btn(ds_body,"  Upload Data File",upload_dataset,color="#0e7490",fg_col="#e0f7fa")
info_row=tk.Frame(ds_body,bg=CARD); info_row.pack(fill="x",pady=(6,0))
rows_var=tk.StringVar(value="—"); cols_var=tk.StringVar(value="—"); num_cols_var=tk.StringVar(value="—")
for title,var,col in [("Unique Wells",rows_var,COL_DATASET),("Columns",cols_var,FG_DIM),("Numeric",num_cols_var,"#f59e0b")]:
    cf=tk.Frame(info_row,bg=CARD); cf.pack(side="left",expand=True,fill="x")
    tk.Label(cf,text=title,bg=CARD,fg=FG_DIM,font=FONT_XS).pack(anchor="w")
    tk.Label(cf,textvariable=var,bg=CARD,fg=col,font=("Courier New",13,"bold")).pack(anchor="w")
event_col_var=tk.StringVar(value="Upload dataset to detect event column")
event_col_lbl=tk.Label(ds_body,textvariable=event_col_var,bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=420,justify="left",anchor="w")
event_col_lbl.pack(fill="x",pady=(4,0))
well_id_var=tk.StringVar(value="Well ID column: Not detected yet")
well_id_lbl=tk.Label(ds_body,textvariable=well_id_var,bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=420,justify="left",anchor="w")
well_id_lbl.pack(fill="x",pady=(2,0))

# ② FEATURE SELECTION
feat_outer=make_section_card(sidebar,"② FEATURE SELECTION  (numeric + categorical)",COL_FEATURES)
tk.Label(feat_outer,
         text="Select columns for analysis.  Blue=numeric  |  Green=categorical.\n"
              "Missing values handled automatically.",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=420,justify="left").pack(anchor="w",pady=(0,4))

feat_top_row=tk.Frame(feat_outer,bg=CARD); feat_top_row.pack(fill="x",pady=(0,3))
feat_count_var=tk.StringVar(value="0 / 0 selected")
tk.Label(feat_top_row,textvariable=feat_count_var,bg=CARD,fg=COL_FEATURES,font=FONT_H3).pack(side="left")

feat_search_row=tk.Frame(feat_outer,bg=CARD); feat_search_row.pack(fill="x",pady=(0,3))
tk.Label(feat_search_row,text="Search:",bg=CARD,fg=FG_DIM,font=FONT_XS).pack(side="left")
feat_search_var=tk.StringVar()
feat_search_entry=tk.Entry(feat_search_row,textvariable=feat_search_var,
                           bg=CARD2,fg=FG,insertbackground=FG,font=FONT_XS,relief="flat",bd=2,width=24)
feat_search_entry.pack(side="left",padx=4)
feat_search_var.trace_add("write",_filter_feat_list)

feat_btn_row=tk.Frame(feat_outer,bg=CARD); feat_btn_row.pack(fill="x",pady=(0,4))
def _sel_all():
    for v in feat_all_vars.values(): v.set(True)
def _sel_none():
    for v in feat_all_vars.values(): v.set(False)
def _sel_numeric_only():
    import pandas as pd
    for col,v in feat_all_vars.items():
        v.set(raw_data is not None and col in raw_data.columns and
              pd.api.types.is_numeric_dtype(raw_data[col]))
for lbl,cmd,fg_c in [("All",_sel_all,COL_FEATURES),
                      ("None",_sel_none,FG_DIM),
                      ("Numeric Only",_sel_numeric_only,"#f59e0b")]:
    tk.Button(feat_btn_row,text=lbl,command=cmd,bg=CARD2,fg=fg_c,
              activebackground=BORDER,activeforeground=FG,font=FONT_XS,
              relief="flat",bd=0,cursor="hand2",padx=8,pady=3).pack(side="left",padx=2)

feat_canvas_outer=tk.Frame(feat_outer,bg=CARD,height=200)
feat_canvas_outer.pack(fill="x"); feat_canvas_outer.pack_propagate(False)
feat_canvas=Canvas(feat_canvas_outer,bg=CARD,highlightthickness=0,bd=0)
feat_sb=Scrollbar(feat_canvas_outer,orient="vertical",command=feat_canvas.yview,
                  bg=BORDER,troughcolor=CARD,activebackground=COL_FEATURES)
feat_canvas.configure(yscrollcommand=feat_sb.set)
feat_sb.pack(side="right",fill="y"); feat_canvas.pack(side="left",fill="both",expand=True)
feat_all_body=tk.Frame(feat_canvas,bg=CARD)
fa_wid=feat_canvas.create_window((0,0),window=feat_all_body,anchor="nw")
def _fa_resize(e):
    feat_canvas.configure(scrollregion=feat_canvas.bbox("all"))
    feat_canvas.itemconfig(fa_wid,width=e.width)
feat_all_body.bind("<Configure>",lambda e: feat_canvas.configure(scrollregion=feat_canvas.bbox("all")))
feat_canvas.bind("<Configure>",_fa_resize)
tk.Label(feat_all_body,text="Upload a dataset to see all columns.",
         bg=CARD,fg=FG_DIM,font=FONT_XS,padx=8,pady=8).pack(anchor="w")

# ③ CLUSTER SETTINGS
cl_body=make_section_card(sidebar,"③ CLUSTER SETTINGS",COL_CLUSTER)
tk.Label(cl_body,
         text="Algorithm chain: KPrototypes → KModes → Gower → Ward → KMeans\n"
              "Works with numeric-only, text-only, or mixed columns.",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=420).pack(anchor="w")
cluster_var=tk.IntVar(value=3)
sf=tk.Frame(cl_body,bg=CARD); sf.pack(fill="x",pady=4)
tk_ttk.Spinbox(sf,from_=2,to=10,textvariable=cluster_var,width=5,font=FONT_H2).pack(side="left")
tk.Label(sf,text="well clusters",bg=CARD,fg=FG_DIM,font=FONT_SM).pack(side="left",padx=8)

# ④ ANOMALY DETECTION
anom_body=make_section_card(sidebar,"④ ANOMALY DETECTION",COL_ANOMALY)
tk.Label(anom_body,text="Detects anomalous wells from their aggregated profiles.",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=420).pack(anchor="w")
anomaly_method_var=tk.StringVar(value="iforest")
mrow=tk.Frame(anom_body,bg=CARD); mrow.pack(fill="x",pady=(2,6))
for label,val in [("Isolation Forest  (faster, tree-based)","iforest"),
                  ("Local Outlier Factor  (density-based)","lof")]:
    tk.Radiobutton(mrow,text=label,variable=anomaly_method_var,value=val,
                   bg=CARD,fg=FG_MID,selectcolor=CARD2,activebackground=CARD,
                   activeforeground=COL_ANOMALY,font=FONT_SM,wraplength=420).pack(anchor="w",pady=1)
anomaly_contam_var=tk.DoubleVar(value=5.0)
crow=tk.Frame(anom_body,bg=CARD); crow.pack(fill="x",pady=2)
tk_ttk.Spinbox(crow,from_=1.0,to=49.0,increment=0.5,textvariable=anomaly_contam_var,width=6,font=FONT_H3).pack(side="left")
tk.Label(crow,text="% contamination  (typical: 3–10%)",bg=CARD,fg=FG_DIM,font=FONT_XS).pack(side="left",padx=6)

# ⑤ RUN
run_body=make_section_card(sidebar,"⑤ RUN ANALYSIS","#6366f1")
make_btn(run_body,"  ▶  Run AI Analysis  (Clustering + Anomaly + Hidden Patterns)",
         start_pipeline,color="#4f46e5")
if HAVE_TTKBS:
    progress_bar=ttk.Progressbar(run_body,mode="determinate",bootstyle="info-striped",length=420)
else:
    progress_bar=tk_ttk.Progressbar(run_body,mode="determinate",length=420)
progress_bar.pack(fill="x",pady=(4,2))
status_var=tk.StringVar(value="Ready — upload a dataset to begin")
status_lbl=tk.Label(run_body,textvariable=status_var,bg=CARD,fg=FG_DIM,font=FONT_SM,anchor="w",wraplength=420,justify="left")
status_lbl.pack(fill="x",pady=(2,0))

# ⑥ EVENT DETECTION PANEL
ev_body=make_section_card(sidebar,"⑥ EVENT DETECTION PANEL",COL_EVENTS)
ev_status_var=tk.StringVar(value="Awaiting analysis")
ev_status_lbl=tk.Label(ev_body,textvariable=ev_status_var,bg=CARD,fg=FG_DIM,font=FONT_SM,wraplength=420)
ev_status_lbl.pack(anchor="w",pady=(0,6))
ev_tile_row=tk.Frame(ev_body,bg=CARD); ev_tile_row.pack(fill="x",pady=(0,6))
ev_active_var=tk.StringVar(value="—"); ev_inactive_var=tk.StringVar(value="—")
ev_abnormal_var=tk.StringVar(value="—"); ev_op_anom_var=tk.StringVar(value="—")
for title,var,col in [("Active Recs",ev_active_var,COL_EVENTS),("Inactive Recs",ev_inactive_var,COL_ANOMALY),
                       ("Abnormal",ev_abnormal_var,WARN),("Op. Anom.",ev_op_anom_var,"#f59e0b")]:
    cell=tk.Frame(ev_tile_row,bg=CARD2,padx=5,pady=5,highlightbackground=BORDER,highlightthickness=1)
    cell.pack(side="left",expand=True,fill="x",padx=2)
    tk.Label(cell,textvariable=var,bg=CARD2,fg=col,font=("Courier New",11,"bold")).pack()
    tk.Label(cell,text=title,bg=CARD2,fg=FG_DIM,font=FONT_XS).pack()
_ev_label_names=[f"ev{i}" for i in range(6)]
ev_count_vars={}; ev_count_lbls={}
ev_detail_frame=tk.Frame(ev_body,bg=CARD); ev_detail_frame.pack(fill="x")
for key in _ev_label_names:
    row=tk.Frame(ev_detail_frame,bg=CARD); row.pack(fill="x",pady=1)
    lbl=tk.Label(row,text="—",bg=CARD,fg=FG_DIM,font=FONT_XS,width=26,anchor="w"); lbl.pack(side="left",padx=(4,0))
    cnt_var=tk.StringVar(value="—")
    tk.Label(row,textvariable=cnt_var,bg=CARD,fg=COL_EVENTS,font=FONT_XS,width=8,anchor="e").pack(side="right",padx=4)
    ev_count_vars[key]=cnt_var; ev_count_lbls[key]=lbl

# ⑦ EXPORT
exp_body=make_section_card(sidebar,"⑦ EXPORT RESULTS",COL_EXPORT)
tk.Label(exp_body,text="Excel: All Records · Well Profiles · Anomaly Wells · Cluster Summary\nAnomaly Explanations · Insights Report",
         bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=420).pack(anchor="w",pady=(0,6))
export_btn=tk.Button(exp_body,text="  Export All Results  (Excel / CSV)",
                     command=export_results,bg="#14532d",fg=COL_EXPORT,
                     activebackground=_lighten("#14532d",20),activeforeground=COL_EXPORT,
                     font=FONT_H3,relief="flat",bd=0,cursor="hand2",padx=10,pady=11,anchor="w",state="disabled")
export_btn.pack(fill="x",pady=2)
export_btn.bind("<Enter>",lambda e: export_btn.config(bg=_lighten("#14532d",20)))
export_btn.bind("<Leave>",lambda e: export_btn.config(bg="#14532d"))

# ⑧ DATASET PREVIEW
prev_body=make_section_card(sidebar,"⑧ DATASET PREVIEW  (first 20 rows)",COL_PREVIEW,fill="both",expand=True)
table_frame=tk.Frame(prev_body,bg=CARD); table_frame.pack(fill="both",expand=True)
tk.Frame(sidebar,bg=SIDEBAR,height=20).pack()


# ══════════════════════════════════════════════════════════════════════════════
# RIGHT PANEL — TABS (no embedded chatbot)
# ══════════════════════════════════════════════════════════════════════════════
rp_outer,right_inner=make_scrollable(root_pane,bg=BG)
root_pane.add(rp_outer,minsize=900)
rp_hdr=tk.Frame(right_inner,bg=BG,padx=16,pady=10); rp_hdr.pack(fill="x")
tk.Label(rp_hdr,text="Analysis Dashboard",bg=BG,fg=FG,font=("Georgia",15,"bold")).pack(side="left")
tk.Label(rp_hdr,text="v7.0  ·  Robust Mixed-Data Clustering  ·  Per-Well Anomaly Explanations  ·  AI Chatbot",
         bg=BG,fg=FG_DIM,font=FONT_SM).pack(side="right",pady=2)
tk.Frame(right_inner,bg=BORDER,height=1).pack(fill="x",padx=12)

stats_wrap=tk.Frame(right_inner,bg=BG,padx=10,pady=5); stats_wrap.pack(fill="x")
stat_defs_r1=[("Unique Wells","—",COL_DATASET),("Total Records","—",COL_CLUSTER),
              ("Clusters","—","#f59e0b"),("Anomaly Wells","—",COL_ANOMALY)]
stat_defs_r2=[("Active Recs","—",COL_EVENTS),("Inactive Recs","—",ERR),
              ("Event Types","—",WARN),("Abnormal Recs","—",COL_ANOMALY)]
stat_widgets={}
for stat_row_defs in [stat_defs_r1,stat_defs_r2]:
    row=tk.Frame(stats_wrap,bg=BG); row.pack(fill="x",pady=2)
    for title,val,col in stat_row_defs:
        sf=tk.Frame(row,bg=CARD,padx=6,pady=6,highlightbackground=col,highlightthickness=1)
        sf.pack(side="left",expand=True,fill="x",padx=3)
        sv=tk.StringVar(value=val)
        tk.Label(sf,textvariable=sv,bg=CARD,fg=col,font=("Courier New",10,"bold"),wraplength=120,justify="center").pack(fill="x")
        tk.Label(sf,text=title,bg=CARD,fg=FG_DIM,font=("Courier New",7),wraplength=120,justify="center").pack(fill="x")
        stat_widgets[title]=sv

# NOTEBOOK
nb_wrap=tk.Frame(right_inner,bg=BG,padx=10,pady=6); nb_wrap.pack(fill="x")
sty=tk_ttk.Style()
sty.configure("CBM.TNotebook",background=BG,tabmargins=[0,0,0,0])
sty.configure("CBM.TNotebook.Tab",background=CARD2,foreground=FG_DIM,padding=[12,6],font=FONT_H3)
sty.map("CBM.TNotebook.Tab",background=[("selected","#1e3a8a")],foreground=[("selected","#93c5fd")])
notebook=tk_ttk.Notebook(nb_wrap,style="CBM.TNotebook"); notebook.pack(fill="x")

TAB_H=500
cluster_tab =tk.Frame(notebook,bg=BG_MID,height=TAB_H)
pca_tab     =tk.Frame(notebook,bg=BG_MID,height=TAB_H)
hidden_tab  =tk.Frame(notebook,bg=BG_MID,height=TAB_H)
event_tab   =tk.Frame(notebook,bg=BG_MID,height=TAB_H)

for tab,name in [
    (cluster_tab,  "  Well Clusters  "),
    (pca_tab,      "  PCA  "),
    (hidden_tab,   "  Hidden Patterns  "),
    (event_tab,    "  Events  "),
]:
    tab.pack_propagate(False); notebook.add(tab,text=name)

# INSIGHTS
ins_wrap=tk.Frame(right_inner,bg=BG,padx=10,pady=6); ins_wrap.pack(fill="x")
ins_hdr=tk.Frame(ins_wrap,bg=BG); ins_hdr.pack(fill="x",pady=(0,4))
tk.Label(ins_hdr,text="AI Insights Report  —  Per-Well Anomaly Explanations  ·  Cluster Summary  ·  Events",
         bg=BG,fg=FG,font=FONT_H2).pack(side="left")
copy_btn=tk.Button(ins_hdr,text="Copy",
                   command=lambda:(app.clipboard_clear(),app.clipboard_append(explain_text.get("1.0","end"))),
                   bg=CARD2,fg=FG_DIM,activebackground=BORDER,activeforeground=FG,
                   font=FONT_XS,relief="flat",bd=0,cursor="hand2",padx=10,pady=4)
copy_btn.pack(side="right")
explain_text_frame=tk.Frame(ins_wrap,bg=CARD,highlightbackground=BORDER,highlightthickness=1)
explain_text_frame.pack(fill="both",expand=True)
explain_xsb=tk_ttk.Scrollbar(explain_text_frame,orient="horizontal")
explain_vsb=tk_ttk.Scrollbar(explain_text_frame,orient="vertical")
explain_text=tk.Text(explain_text_frame,height=28,bg=CARD,fg="#a8f0c8",
                     font=("Courier New",9),relief="flat",bd=0,padx=14,pady=10,
                     insertbackground=FG,wrap="none",
                     xscrollcommand=explain_xsb.set,yscrollcommand=explain_vsb.set,
                     highlightthickness=0,state="disabled")
explain_xsb.config(command=explain_text.xview); explain_vsb.config(command=explain_text.yview)
explain_vsb.pack(side="right",fill="y"); explain_xsb.pack(side="bottom",fill="x")
explain_text.pack(side="left",fill="both",expand=True)

tk.Frame(right_inner,bg=BG,height=20).pack()

# Initial insights text
explain_text.config(state="normal")
explain_text.insert("end",
    "============================================================\n"
    "  CBM AI Analytics Platform  v7.0  (Floating Chatbot Edition)\n"
    "============================================================\n\n"
    "  AI CHATBOT — NOW A FLOATING ICON\n"
    "  ----------------------------------------------------------\n"
    "  ★ Click the  ★ AI  button (bottom-right corner) to open\n"
    "    the AI Assistant in a compact floating panel.\n\n"
    "  • Drag the panel by its title bar to reposition it\n"
    "  • Press ESC or ✕ to close\n"
    "  • The chatbot receives your full analysis report as context\n"
    "  • Ask anything about wells, clusters, anomalies, or the app\n\n"
    "  QUICK START:\n"
    "  ----------------------------------------------------------\n"
    "  1. Upload data file  (any format)\n"
    "  2. Select features to analyse\n"
    "  3. Set cluster count + anomaly contamination %\n"
    "  4. Click Run AI Analysis\n"
    "  5. Read per-well explanations in Insights Report\n"
    "  6. Click  ★ AI  to ask the chatbot deeper questions\n"
    "============================================================\n"
)
explain_text.config(state="disabled")


# ══════════════════════════════════════════════════════════════════════════════
# FLOATING ACTION BUTTON (FAB) — AI Chatbot icon
# Placed as an overlay on the app root window, bottom-right corner
# ══════════════════════════════════════════════════════════════════════════════

# Build the popup first (hidden), then show the FAB
_build_chat_popup()

# FAB container — we use a Toplevel anchored to the app so it stays on top
# but moves with the window
fab_frame = tk.Frame(app, bg=BG_DEEP)
fab_frame.place(relx=1.0, rely=1.0, x=-16, y=-16, anchor="se")

# Tooltip label above FAB
fab_tip_var = tk.StringVar(value="")
fab_tip_lbl = tk.Label(fab_frame, textvariable=fab_tip_var,
                       bg="#150a2e", fg=COL_CHATBOT,
                       font=("Courier New", 8), padx=6, pady=3,
                       relief="flat", bd=0)

chat_fab_btn = tk.Button(
    fab_frame,
    text="★ AI",
    command=toggle_chat_popup,
    bg="#4f46e5", fg="#ffffff",
    activebackground="#7c3aed", activeforeground="#ffffff",
    font=("Georgia", 11, "bold"),
    relief="flat", bd=0,
    cursor="hand2",
    padx=16, pady=12,
    width=6
)
chat_fab_btn.pack(side="bottom")

# Hover tooltip
def _fab_enter(e):
    fab_tip_var.set("AI Assistant")
    fab_tip_lbl.pack(side="top", pady=(0,4))
def _fab_leave(e):
    fab_tip_lbl.pack_forget()

chat_fab_btn.bind("<Enter>", _fab_enter)
chat_fab_btn.bind("<Leave>", _fab_leave)
Tooltip(chat_fab_btn, "Click to open/close the AI Assistant chat panel")

# Start pulse animation
app.after(1000, _pulse_icon)

# Keep FAB on top and re-raise whenever focus changes
def _keep_fab_top(e=None):
    try:
        fab_frame.lift()
    except: pass
app.bind("<FocusIn>", _keep_fab_top)
app.bind("<Configure>", _keep_fab_top)

app.mainloop()
