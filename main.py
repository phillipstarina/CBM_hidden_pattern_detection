"""
CBM AI Analytics Platform — v6.0
─────────────────────────────────────────────
NEW in v6.0:
  • Multi-table Access loader — loads dbo_PRD, dbo_SS, dbo_XY simultaneously
  • Auto-detect join keys across all tables
  • Smart merge engine — produces one unified analytical DataFrame
  • Data Inspector panel — nulls, dtypes, distributions, inconsistencies
  • Table relationship viewer with visual summary
  • All v5.0 features retained
"""

import tkinter as tk
from tkinter import filedialog, Listbox, Canvas, Scrollbar, messagebox, ttk as tk_ttk
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
# COLOUR PALETTE
# ══════════════════════════════════════════════════════════════════════════════
BG_DEEP = "#060a10"
BG     = "#0d1117"
BG_MID = "#0f1520"
SIDEBAR= "#080c14"
CARD   = "#131b2a"
CARD2  = "#1a2438"
BORDER = "#1e3a5f"
BORDER2= "#0e2040"

COL_DATASET  = "#22d3ee"
COL_CLUSTER  = "#a78bfa"
COL_WEIGHTS  = "#f59e0b"
COL_ANOMALY  = "#f87171"
COL_EVENTS   = "#34d399"
COL_FEATURES = "#60a5fa"
COL_EXPORT   = "#4ade80"
COL_PREVIEW  = "#94a3b8"
COL_INSPECT  = "#e879f9"   # NEW — inspector accent
COL_MERGE    = "#38bdf8"   # NEW — merge accent

ACCENT    = "#3b82f6"
ACCENT_LT = "#60a5fa"
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

plt.rcParams.update({
    "figure.facecolor":"#0d1117","axes.facecolor":"#161d2e",
    "text.color":"#e2e8f0","axes.labelcolor":"#94a3b8",
    "xtick.color":"#94a3b8","ytick.color":"#94a3b8",
    "axes.edgecolor":"#2d3f5e","grid.color":"#1e3a5f",
    "axes.grid":True,"grid.linewidth":0.5,"grid.alpha":0.4,
    "legend.facecolor":"#161d2e","legend.edgecolor":"#2d3f5e",
    "legend.fontsize":9,"figure.autolayout":False,
    "font.family":"monospace",
})

# ── Global state ──────────────────────────────────────────────────────────────
raw_data             = None   # single-table or merged DataFrame
table_registry       = {}     # {table_name: DataFrame}  — all loaded tables
merge_report         = {}     # result of _auto_merge()
active_df            = None
active_X             = None
active_xcols         = []
active_anomaly_result= None
active_weight_result = None
active_event_result  = None
active_figures       = {}
weight_vars          = {}
weight_row_frames    = []
weight_sum_var       = None
weight_sum_lbl       = None


# ══════════════════════════════════════════════════════════════════════════════
# ACCESS CONNECTION HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def _get_access_connection(filepath):
    import sys
    filepath = os.path.abspath(filepath)
    errors = []
    if sys.platform.startswith("win"):
        try:
            import pyodbc
        except ImportError:
            errors.append("pyodbc not installed  →  pip install pyodbc")
            pyodbc = None
        if pyodbc is not None:
            ace_drivers = [
                "Microsoft Access Driver (*.mdb, *.accdb)",
                "Microsoft Access Driver (*.mdb)",
            ]
            for drv in ace_drivers:
                if drv not in pyodbc.drivers():
                    errors.append(f"Driver not installed: '{drv}'")
                    continue
                try:
                    conn_str = f"DRIVER={{{drv}}};DBQ={filepath};ExtendedAnsiSQL=1;"
                    conn = pyodbc.connect(conn_str, autocommit=True)
                    return conn, "pyodbc"
                except Exception as e:
                    errors.append(f"{drv}: {e}")
            short = "\n".join(f"  • {e}" for e in errors[:6])
            raise RuntimeError(
                "Cannot open Access file on Windows.\n\n"
                "Errors:\n" + short + "\n\n"
                "━━  FIX  ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                "Install the FREE Microsoft Access Database Engine:\n"
                "  https://www.microsoft.com/en-us/download/details.aspx?id=54920\n\n"
                "Match the bitness of your Python:\n"
                "  python -c \"import struct; print(struct.calcsize('P')*8, 'bit')\"\n"
                "  64-bit Python → accessdatabaseengine_X64.exe\n"
                "  32-bit Python → accessdatabaseengine.exe\n"
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
            )
    try:
        import pandas_access as mdb  # noqa
        return filepath, "pandas_access"
    except ImportError:
        raise RuntimeError(
            "Could not load Access file.\n\n"
            "Install:\n"
            "  sudo apt install mdbtools\n"
            "  brew install mdbtools\n"
            "  pip install pandas-access\n"
        )


def _list_access_tables(conn, backend):
    if backend == "pyodbc":
        cursor = conn.cursor()
        tables = [
            r.table_name for r in cursor.tables(tableType="TABLE")
            if not r.table_name.startswith("MSys")
        ]
        cursor.close()
        return tables
    else:
        import pandas_access as mdb
        return list(mdb.list_tables(conn))


def _read_access_table(conn, backend, table_name):
    import pandas as pd
    if backend == "pyodbc":
        return pd.read_sql(f"SELECT * FROM [{table_name}]", conn)
    else:
        import pandas_access as mdb
        return mdb.read_table(conn, table_name)


# ══════════════════════════════════════════════════════════════════════════════
# ★ NEW — MULTI-TABLE LOADER
# ══════════════════════════════════════════════════════════════════════════════
# Target table name fragments (case-insensitive partial match)
TARGET_TABLES = ["prd", "ss", "xy"]


def _is_target_table(name):
    nl = name.lower()
    return any(f"_{t}" in nl or nl.endswith(t) or nl.startswith(t)
               for t in TARGET_TABLES)


def load_all_access_tables(filepath):
    """
    Connect to the Access file, load ALL user tables, return dict {name: df}.
    Prioritises tables matching dbo_PRD / dbo_SS / dbo_XY naming.
    """
    conn, backend = _get_access_connection(filepath)
    all_tables    = _list_access_tables(conn, backend)
    if backend == "pyodbc":
        conn.close()

    results = {}
    errors  = {}
    for tname in all_tables:
        try:
            conn2, _ = _get_access_connection(filepath)
            df = _read_access_table(conn2, backend, tname)
            if backend == "pyodbc":
                conn2.close()
            results[tname] = df
        except Exception as e:
            errors[tname] = str(e)

    return results, errors, all_tables


# ══════════════════════════════════════════════════════════════════════════════
# ★ NEW — AUTO-JOIN ENGINE
# ══════════════════════════════════════════════════════════════════════════════
def _find_common_keys(df1, df2):
    """Return list of column names that exist in both tables (case-insensitive)."""
    c1 = {c.lower(): c for c in df1.columns}
    c2 = {c.lower(): c for c in df2.columns}
    common_lower = set(c1.keys()) & set(c2.keys())
    # Prefer columns that look like IDs
    id_hints = ["id","well","pad","location","name","no","num","key","code"]
    scored = []
    for cl in common_lower:
        score = sum(h in cl for h in id_hints)
        scored.append((score, cl, c1[cl], c2[cl]))
    scored.sort(reverse=True)
    return [(c1n, c2n) for _, _, c1n, c2n in scored]


def _auto_merge(tables: dict):
    """
    Given {name: DataFrame}, detect relationships and produce a merged DataFrame.
    Returns a report dict with full diagnostics.
    """
    import pandas as pd

    report = {
        "tables": {},
        "relationships": [],
        "merged_df": None,
        "merge_log": [],
        "warnings": [],
        "nulls_pre_merge": {},
        "nulls_post_merge": {},
    }

    if not tables:
        report["warnings"].append("No tables loaded.")
        return report

    # 1. Per-table inspection
    for name, df in tables.items():
        num_cols  = list(df.select_dtypes(include="number").columns)
        cat_cols  = list(df.select_dtypes(exclude="number").columns)
        null_info = {c: int(df[c].isna().sum()) for c in df.columns if df[c].isna().sum() > 0}
        dup_count = int(df.duplicated().sum())
        report["tables"][name] = {
            "rows":      len(df),
            "cols":      len(df.columns),
            "columns":   list(df.columns),
            "dtypes":    {c: str(df[c].dtype) for c in df.columns},
            "num_cols":  num_cols,
            "cat_cols":  cat_cols,
            "null_cols": null_info,
            "duplicates":dup_count,
            "sample":    df.head(3).to_dict(orient="records"),
        }
        report["nulls_pre_merge"][name] = null_info

    # 2. Detect pairwise relationships
    names = list(tables.keys())
    for i in range(len(names)):
        for j in range(i+1, len(names)):
            n1, n2 = names[i], names[j]
            pairs = _find_common_keys(tables[n1], tables[n2])
            if pairs:
                c1, c2 = pairs[0]  # best key
                overlap = set(tables[n1][c1].dropna().astype(str)) & \
                          set(tables[n2][c2].dropna().astype(str))
                report["relationships"].append({
                    "table_a": n1, "table_b": n2,
                    "key_a":   c1, "key_b":   c2,
                    "overlap": len(overlap),
                    "all_candidates": pairs,
                })

    # 3. Build merged DataFrame
    # Strategy: start with PRD (largest/time-series), left-join SS then XY
    def _score(name):
        nl = name.lower()
        if "prd" in nl: return 0
        if "ss"  in nl: return 1
        if "xy"  in nl: return 2
        return 3

    ordered = sorted(names, key=_score)
    merged  = tables[ordered[0]].copy()
    log     = [f"Base table: {ordered[0]}  ({len(merged):,} rows)"]

    for tname in ordered[1:]:
        right = tables[tname]
        pairs = _find_common_keys(merged, right)
        if not pairs:
            report["warnings"].append(
                f"No common key found between merged table and '{tname}' — skipped.")
            log.append(f"  SKIP {tname} — no common key")
            continue
        lkey, rkey = pairs[0]
        # Avoid duplicate columns: suffix strategy
        overlap_cols = [c for c in right.columns
                        if c in merged.columns and c != rkey]
        right_renamed = right.rename(
            columns={c: f"{c}__{tname}" for c in overlap_cols}
        )
        before = len(merged)
        merged = merged.merge(right_renamed, left_on=lkey, right_on=rkey,
                              how="left", suffixes=("", f"__{tname}"))
        after  = len(merged)
        log.append(
            f"  LEFT JOIN {tname} ON {lkey}={rkey}  "
            f"({before:,} → {after:,} rows, {len(right.columns)} new cols)"
        )

    report["merged_df"]  = merged
    report["merge_log"]  = log

    # 4. Post-merge null check
    null_post = {c: int(merged[c].isna().sum())
                 for c in merged.columns if merged[c].isna().sum() > 0}
    report["nulls_post_merge"] = null_post

    # 5. Inconsistency flags
    flags = []
    for rel in report["relationships"]:
        n1, n2 = rel["table_a"], rel["table_b"]
        c1, c2 = rel["key_a"], rel["key_b"]
        set1 = set(tables[n1][c1].dropna().astype(str))
        set2 = set(tables[n2][c2].dropna().astype(str))
        only_in_1 = set1 - set2
        only_in_2 = set2 - set1
        if only_in_1:
            flags.append(
                f"  ⚠ {len(only_in_1)} keys in [{n1}].{c1} "
                f"not found in [{n2}].{c2}  → orphan rows after join"
            )
        if only_in_2:
            flags.append(
                f"  ⚠ {len(only_in_2)} keys in [{n2}].{c2} "
                f"not found in [{n1}].{c1}  → unmatched reference rows"
            )
    report["warnings"].extend(flags)

    return report


# ══════════════════════════════════════════════════════════════════════════════
# ★ NEW — DATA INSPECTOR TEXT REPORT
# ══════════════════════════════════════════════════════════════════════════════
def _build_inspector_report(report: dict) -> str:
    lines = [
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        "  CBM DATA INSPECTOR REPORT  v6.0",
        f"  {datetime.datetime.now().strftime('%Y-%m-%d  %H:%M')}",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    ]

    # Per-table
    for tname, info in report["tables"].items():
        lines += [
            "",
            f"  TABLE: {tname}",
            f"  {'─'*40}",
            f"  Rows        : {info['rows']:,}",
            f"  Columns     : {info['cols']}",
            f"  Duplicates  : {info['duplicates']:,}",
            "",
            "  Columns & types:",
        ]
        for col, dtype in info["dtypes"].items():
            null_flag = f"  ← {info['null_cols'][col]} nulls" \
                        if col in info["null_cols"] else ""
            lines.append(f"    {col:<35} {dtype:<12}{null_flag}")

        if info["null_cols"]:
            lines.append(f"\n  NULL columns: {list(info['null_cols'].keys())}")
        else:
            lines.append("\n  NULL columns: none")

    # Relationships
    lines += [
        "",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        "  DETECTED RELATIONSHIPS",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    ]
    if report["relationships"]:
        for rel in report["relationships"]:
            lines += [
                f"  [{rel['table_a']}].{rel['key_a']}",
                f"     ↔  [{rel['table_b']}].{rel['key_b']}",
                f"     Overlapping keys: {rel['overlap']:,}",
                f"     All candidates  : "
                + ", ".join(f"{a}={b}" for a, b in rel["all_candidates"]),
                "",
            ]
    else:
        lines.append("  No common keys detected between any tables.")

    # Merge log
    lines += [
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        "  MERGE EXECUTION LOG",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    ]
    lines.extend(f"  {l}" for l in report["merge_log"])

    # Merged result
    mdf = report.get("merged_df")
    if mdf is not None:
        lines += [
            "",
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
            "  MERGED TABLE SUMMARY",
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
            f"  Total rows    : {len(mdf):,}",
            f"  Total columns : {len(mdf.columns)}",
            "",
            "  Post-merge nulls:",
        ]
        if report["nulls_post_merge"]:
            for col, cnt in sorted(report["nulls_post_merge"].items(),
                                   key=lambda x: -x[1]):
                pct = cnt / max(len(mdf), 1) * 100
                lines.append(f"    {col:<40} {cnt:>6,}  ({pct:.1f}%)")
        else:
            lines.append("    None — clean merge!")

    # Warnings
    if report["warnings"]:
        lines += [
            "",
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
            "  WARNINGS & INCONSISTENCIES",
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        ]
        lines.extend(f"  {w}" for w in report["warnings"])

    # Recommended SQL
    lines += [
        "",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        "  RECOMMENDED JOIN QUERY (Access SQL)",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    ]
    if report["relationships"]:
        rels  = report["relationships"]
        names = list(report["tables"].keys())
        base  = rels[0]["table_a"]
        lines.append(f"  SELECT p.*, s.*, x.*")
        lines.append(f"  FROM   [{base}] p")
        for rel in rels:
            other = rel["table_b"] if rel["table_a"] == base else rel["table_a"]
            ka    = rel["key_a"]   if rel["table_a"] == base else rel["key_b"]
            kb    = rel["key_b"]   if rel["table_a"] == base else rel["key_a"]
            alias = other[0].lower()
            lines.append(f"  LEFT JOIN [{other}] {alias} ON p.{ka} = {alias}.{kb}")
        lines.append(f"  ORDER BY p.{rels[0]['key_a']};")
    else:
        lines.append("  (No relationships detected — cannot suggest JOIN)")

    lines.append("\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# FILE LOADERS (non-Access)
# ══════════════════════════════════════════════════════════════════════════════
def load_any_file(filepath):
    import pandas as pd
    ext = os.path.splitext(filepath)[1].lower().strip(".")
    strategies = []
    if ext in ("csv","tsv","txt",""):
        sep = "\t" if ext in ("tsv","txt") else ","
        strategies += [
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="utf-8"),
            lambda f, s=sep: pd.read_csv(f, sep=s, encoding="latin-1"),
            lambda f:        pd.read_csv(f, sep=None, engine="python", encoding="utf-8"),
            lambda f:        pd.read_csv(f, sep=None, engine="python", encoding="latin-1"),
        ]
    if ext in ("xlsx","xlsm","xlsb","xls","ods","odf","odt"):
        for eng in (["openpyxl",None] if ext in ("xlsx","xlsm") else
                    ["pyxlsb"]        if ext == "xlsb" else
                    ["xlrd",None]     if ext == "xls"  else
                    ["odf",None]):
            strategies.append(lambda f, e=eng: pd.read_excel(f, engine=e) if e else pd.read_excel(f))
    if ext == "json":
        strategies += [lambda f: pd.read_json(f, orient="records"), lambda f: pd.read_json(f)]
    if ext == "parquet":
        strategies += [lambda f: pd.read_parquet(f)]
    if ext == "feather":
        strategies += [lambda f: pd.read_feather(f)]
    if ext in ("h5","hdf5","hdf"):
        strategies += [lambda f: pd.read_hdf(f)]
    if ext in ("pkl","pickle"):
        strategies += [lambda f: pd.read_pickle(f)]
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
                return df
        except Exception as e:
            errors.append(str(e))
    raise RuntimeError(
        f"Could not read '{os.path.basename(filepath)}'.\n\n"
        f"Attempted {len(strategies)} loading strategies.\n"
        + "\n".join(f"  • {e}" for e in errors[:5])
    )


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
# UI HELPERS
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
    sb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(canvas, bg=bg)
    wid   = canvas.create_window((0,0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: (
        canvas.configure(scrollregion=canvas.bbox("all")),
        canvas.itemconfig(wid, width=e.width)
    ))
    def _wheel(e):
        if   e.num==4: canvas.yview_scroll(-1,"units")
        elif e.num==5: canvas.yview_scroll(1,"units")
        else:          canvas.yview_scroll(int(-1*(e.delta/120)),"units")
    canvas.bind_all("<MouseWheel>",_wheel)
    canvas.bind_all("<Button-4>",  _wheel)
    canvas.bind_all("<Button-5>",  _wheel)
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
    progress_bar.stop(); progress_bar.config(mode="determinate"); progress_bar["value"]=100


# ══════════════════════════════════════════════════════════════════════════════
# WEIGHT PANEL
# ══════════════════════════════════════════════════════════════════════════════
def _update_weight_sum(*_):
    total = 0.0
    for v in weight_vars.values():
        try: total += float(v.get())
        except: pass
    ok = abs(total-100.0) < 0.5
    weight_sum_var.set(f"Total: {total:.1f}%  {'✔ OK' if ok else '⚠ should be 100'}")
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
    default_pct = round(100.0/len(num_cols), 1)
    for i, col in enumerate(num_cols):
        col_color = CLUSTER_PALETTE[i % len(CLUSTER_PALETTE)]
        row = tk.Frame(weight_body, bg=CARD2, highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", pady=2); weight_row_frames.append(row)
        tk.Canvas(row, bg=col_color, width=10, height=10,
                  highlightthickness=0).pack(side="left", padx=(6,4), pady=6)
        tk.Label(row, text=col[:22], bg=CARD2, fg=col_color,
                 font=FONT_XS, width=22, anchor="w").pack(side="left")
        wvar = tk.StringVar(value=str(default_pct))
        wvar.trace_add("write", _update_weight_sum)
        weight_vars[col] = wvar
        tk_ttk.Spinbox(row, from_=0.0, to=100.0, increment=1.0,
                       textvariable=wvar, width=7, font=FONT_XS).pack(side="left", padx=4, pady=4)
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
    return {c: round(v/total*100,2) for c,v in raw.items()}, total


# ══════════════════════════════════════════════════════════════════════════════
# ★ NEW — MULTI-TABLE UPLOAD ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════════
def upload_dataset():
    global raw_data, table_registry, merge_report
    f = filedialog.askopenfilename(
        title="Open Data File",
        filetypes=[
            ("All supported files",
             "*.mdb *.accdb *.csv *.tsv *.txt *.xlsx *.xls *.xlsb *.xlsm "
             "*.ods *.json *.parquet *.feather *.h5 *.hdf5 *.pkl *.pickle"),
            ("Microsoft Access 2007", "*.accdb *.mdb"),
            ("CSV / TSV",             "*.csv *.tsv *.txt"),
            ("Excel",                 "*.xlsx *.xls *.xlsb *.xlsm"),
            ("ODS (LibreOffice)",     "*.ods"),
            ("JSON",                  "*.json"),
            ("Parquet / Feather",     "*.parquet *.feather"),
            ("HDF5",                  "*.h5 *.hdf5"),
            ("Pickle",                "*.pkl *.pickle"),
            ("All files",             "*.*"),
        ]
    )
    if not f: return
    ext = os.path.splitext(f)[1].lower()
    set_status(f"⟳  Loading {os.path.basename(f)} …", WARN)
    progress_start()

    if ext in (".mdb", ".accdb"):
        # ── Access: load ALL tables, auto-merge ────────────────────────────
        def _access_load():
            try:
                tables, load_errors, all_names = load_all_access_tables(f)
                if not tables:
                    raise RuntimeError("No readable tables found in the Access database.")

                report = _auto_merge(tables)
                merged  = report.get("merged_df")
                if merged is None or len(merged) == 0:
                    # Fall back to largest table
                    merged = max(tables.values(), key=len)

                app.after(0, lambda: _post_access_load(
                    tables, report, merged, f, load_errors))
            except Exception as e:
                app.after(0, lambda err=str(e): _load_error(err))

        threading.Thread(target=_access_load, daemon=True).start()
        return

    # ── Non-Access ─────────────────────────────────────────────────────────
    def _load():
        try:
            df = load_any_file(f)
            if len(df) == 0:    raise RuntimeError("File loaded but contains no rows.")
            if len(df.columns) == 0: raise RuntimeError("File loaded but contains no columns.")
            _finish_single_load(df, f)
        except Exception as e:
            app.after(0, lambda err=str(e): _load_error(err))
    threading.Thread(target=_load, daemon=True).start()


def _post_access_load(tables, report, merged, filepath, load_errors):
    """Called on main thread after all Access tables are loaded and merged."""
    global raw_data, table_registry, merge_report
    table_registry = tables
    merge_report   = report
    raw_data       = merged

    num_cols = list(merged.select_dtypes(include="number").columns)
    rows_var.set(f"{len(merged):,}")
    cols_var.set(str(len(merged.columns)))
    num_cols_var.set(f"{len(num_cols)} numeric")

    # Table count badge
    n_tables = len(tables)
    table_count_var.set(f"{n_tables} tables merged")

    # Populate feature listbox
    x_listbox.delete(0, "end")
    for col in merged.columns:
        x_listbox.insert("end", f"  {col}")

    rebuild_weight_panel(num_cols)

    # Event column detection
    ec = _find_event_column(merged)
    if ec:
        event_col_var.set(f"✔  Event column: '{ec}'")
        event_col_lbl.config(fg=SUCCESS2)
    else:
        event_col_var.set("⚠  No event/status column found")
        event_col_lbl.config(fg=WARN)

    preview_table(merged)
    progress_stop()

    # Build inspector report and show it
    inspector_report = _build_inspector_report(report)
    _show_inspector_report(inspector_report, report)

    name = os.path.basename(filepath)
    set_status(
        f"✔  Loaded & merged: {name}  "
        f"({n_tables} tables → {len(merged):,} rows × {len(merged.columns)} cols)",
        SUCCESS
    )
    for k in stat_widgets:
        stat_widgets[k].set("—")
    stat_widgets["Total Wells"].set(f"{len(merged):,}")
    export_btn.config(state="normal")

    if load_errors:
        err_msg = "\n".join(f"  {t}: {e}" for t, e in load_errors.items())
        messagebox.showwarning(
            "Some Tables Failed to Load",
            f"The following tables could not be read:\n{err_msg}\n\n"
            "All other tables were loaded and merged successfully."
        )


def _finish_single_load(df, filepath):
    """Non-Access file post-load."""
    global raw_data
    raw_data = df
    num_cols = list(df.select_dtypes(include="number").columns)
    app.after(0, lambda: _post_single_load(df, filepath, num_cols))


def _post_single_load(df, filepath, num_cols):
    global raw_data
    raw_data = df
    rows_var.set(f"{len(df):,}")
    cols_var.set(str(len(df.columns)))
    num_cols_var.set(f"{len(num_cols)} numeric")
    table_count_var.set("single table")
    x_listbox.delete(0, "end")
    for col in df.columns:
        x_listbox.insert("end", f"  {col}")
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
    set_status(f"✔  Loaded: {os.path.basename(filepath)}  ({len(df):,} rows × {len(df.columns)} cols)", SUCCESS)
    for k in stat_widgets:
        stat_widgets[k].set("—")
    stat_widgets["Total Wells"].set(f"{len(df):,}")
    export_btn.config(state="normal")


def _load_error(msg):
    progress_stop()
    set_status("✖  Load failed — see error dialog", ERR)
    messagebox.showerror("File Load Error", msg)


# ══════════════════════════════════════════════════════════════════════════════
# ★ NEW — INSPECTOR REPORT DISPLAY
# ══════════════════════════════════════════════════════════════════════════════
def _show_inspector_report(text_report: str, report: dict):
    """Populate the Inspector tab with the text report and table cards."""
    # Text area
    inspector_text.config(state="normal")
    inspector_text.delete("1.0", "end")
    inspector_text.insert("end", text_report)
    inspector_text.config(state="disabled")

    # Relationship tiles
    for w in rel_tiles_frame.winfo_children():
        w.destroy()

    if report["relationships"]:
        for rel in report["relationships"]:
            col = CLUSTER_PALETTE[hash(rel["table_a"]) % len(CLUSTER_PALETTE)]
            tile = tk.Frame(rel_tiles_frame, bg=CARD2,
                            highlightbackground=col, highlightthickness=1)
            tile.pack(fill="x", pady=3, padx=6)
            tk.Label(tile,
                     text=f"  [{rel['table_a']}].{rel['key_a']}  ↔  [{rel['table_b']}].{rel['key_b']}",
                     bg=CARD2, fg=col, font=FONT_H3).pack(anchor="w", padx=6, pady=(4,0))
            tk.Label(tile,
                     text=f"  Overlapping keys: {rel['overlap']:,}   |   "
                          f"Candidates: {len(rel['all_candidates'])}",
                     bg=CARD2, fg=FG_DIM, font=FONT_XS).pack(anchor="w", padx=6, pady=(0,4))
    else:
        tk.Label(rel_tiles_frame, text="No relationships detected.",
                 bg=CARD, fg=WARN, font=FONT_SM).pack(anchor="w", padx=8, pady=6)

    # Switch to inspector tab
    notebook.select(inspector_tab)


# ══════════════════════════════════════════════════════════════════════════════
# DATA UTILITIES
# ══════════════════════════════════════════════════════════════════════════════
def _find_event_column(df):
    keywords = ["status","state","event","mode","condition","operation","type"]
    for c in df.columns:
        cl = c.lower().replace(" ","_").replace("-","_")
        if any(k in cl for k in keywords):
            return c
    return None


def preview_table(df):
    for w in table_frame.winfo_children():
        w.destroy()
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
    tv["columns"] = list(df.columns); tv["show"] = "headings"
    for col in df.columns:
        tv.heading(col, text=col); tv.column(col, width=90, minwidth=60)
    for row in df.head(20).values:
        tv.insert("","end", values=list(row))
    hsb.pack(side="bottom", fill="x"); vsb.pack(side="right", fill="y")
    tv.pack(fill="both", expand=True)


# ══════════════════════════════════════════════════════════════════════════════
# ANALYSIS MODULES (unchanged from v5.0)
# ══════════════════════════════════════════════════════════════════════════════
def build_weighted_X(df, x_cols, weight_map):
    cols  = [c for c in x_cols if c in df.columns]
    Xraw  = df[cols].values.astype(float)
    mu, s = Xraw.mean(axis=0), Xraw.std(axis=0)
    s[s==0] = 1
    X_scaled = (Xraw-mu)/s
    w_vec    = np.array([np.sqrt(weight_map.get(c,1.0)/100.0) for c in cols])
    return X_scaled*w_vec, cols


def run_clustering(X_w, n_clusters):
    from sklearn.cluster import KMeans
    n  = min(n_clusters, X_w.shape[0])
    km = KMeans(n_clusters=n, random_state=42, n_init="auto")
    return km.fit_predict(X_w), km.inertia_


def detect_anomalies(X_w, contamination=0.05, method="iforest"):
    from sklearn.ensemble import IsolationForest
    from sklearn.neighbors import LocalOutlierFactor
    n = X_w.shape[0]
    if n < 10:
        labels = np.ones(n, dtype=int)
        return labels, 0.0, [], "N/A (too few samples)", np.zeros(n)
    safe_cont = float(np.clip(contamination, 0.001, 0.499))
    if method == "lof":
        det    = LocalOutlierFactor(n_neighbors=min(20,n-1), contamination=safe_cont)
        labels = det.fit_predict(X_w)
        scores = det.negative_outlier_factor_
        name   = "Local Outlier Factor (LOF)"
    else:
        det    = IsolationForest(n_estimators=200, contamination=safe_cont,
                                 random_state=42, n_jobs=-1)
        labels = det.fit_predict(X_w)
        scores = det.decision_function(X_w)
        name   = "Isolation Forest"
    idx = list(np.where(labels==-1)[0])
    pct = len(idx)/n*100
    return labels, pct, idx, name, scores


def analyse_events(df, anomaly_labels):
    ec = _find_event_column(df)
    if ec is None:
        return {
            "has_events":False,"event_col":None,
            "event_counts":{},"event_labels":np.full(len(df),"Unknown"),
            "n_active":0,"n_inactive":0,
            "n_abnormal":int((anomaly_labels==-1).sum()),
            "operational_anomalies":[],
            "message":"No event/status column found in dataset.",
        }
    raw_vals    = df[ec].astype(str).str.strip()
    event_labels= raw_vals.values
    from collections import Counter
    event_counts= dict(Counter(event_labels))
    active_kw   = ["on","active","running","producing","open"]
    inactive_kw = ["off","inactive","shut","stop","closed","idle","down"]
    def _classify(val):
        vl = val.lower()
        if any(k in vl for k in active_kw):   return "active"
        if any(k in vl for k in inactive_kw): return "inactive"
        return "other"
    statuses  = raw_vals.apply(_classify)
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
        "has_events":True,"event_col":ec,
        "event_counts":event_counts,"event_labels":event_labels,
        "n_active":n_active,"n_inactive":n_inactive,
        "n_abnormal":n_abnormal,"operational_anomalies":op_anomalies,
        "message":f"Events read from column: '{ec}'",
    }


# ══════════════════════════════════════════════════════════════════════════════
# VISUALISATIONS
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
            mask = df["cluster"]==cl; col = hex_colors[int(cl)%len(hex_colors)]
            ax.scatter(df.loc[mask,x_col], df.loc[mask,y_col], color=col, s=55,
                       edgecolors="#ffffff22", linewidths=0.5, zorder=3,
                       label=f"Cluster {cl}  ({int(mask.sum()):,} wells)")
        ax.set_xlabel(x_col,fontsize=9,labelpad=6); ax.set_ylabel(y_col,fontsize=9,labelpad=6)
    else:
        ax.text(0.2,0.5,"Select X and Y features",transform=ax.transAxes,color=FG_DIM,fontsize=10)
    ax.set_title("CBM Production Clusters",color=COL_CLUSTER,fontsize=13,pad=12,fontweight="bold")
    ax.legend(loc="upper left",fontsize=8.5,framealpha=0.92,
              edgecolor="#334155",facecolor="#161d2e",labelcolor="#e2e8f0",markerscale=1.4)
    return fig

def plot_pca(X, labels, hex_colors):
    from sklearn.decomposition import PCA
    fig, ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    if X.shape[0]<2 or X.shape[1]<1:
        ax.text(0.3,0.5,"Not enough data",transform=ax.transAxes,color=FG_DIM); return fig
    n = min(2,X.shape[1]); Z = PCA(n_components=n).fit_transform(X)
    if Z.shape[1]==1: Z = np.hstack([Z,np.zeros_like(Z)])
    for cl in sorted(np.unique(labels)):
        mask = labels==cl; col = hex_colors[int(cl)%len(hex_colors)]
        ax.scatter(Z[mask,0],Z[mask,1],color=col,s=45,edgecolors="#ffffff22",
                   linewidths=0.5,zorder=3,label=f"Cluster {cl}  ({int(mask.sum()):,})")
    ax.set_xlabel("PC 1",fontsize=9,labelpad=6); ax.set_ylabel("PC 2",fontsize=9,labelpad=6)
    ax.set_title("PCA — Well Feature Space",color=COL_CLUSTER,fontsize=13,pad=12,fontweight="bold")
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
            handles.append(mpatches.Patch(color=col,label=f"Cluster {cl}  ({int(mask.sum()):,})"))
        ax.set_xlabel(cols3[0],color="#94a3b8",fontsize=7,labelpad=3)
        ax.set_ylabel(cols3[1],color="#94a3b8",fontsize=7,labelpad=3)
        ax.set_zlabel(cols3[2],color="#94a3b8",fontsize=7,labelpad=3)
        ax.legend(handles=handles,loc="upper left",fontsize=8,framealpha=0.85,
                  facecolor="#161d2e",edgecolor="#334155",labelcolor="#e2e8f0")
    else:
        ax.text2D(0.15,0.5,"Need ≥ 3 numeric features",transform=ax.transAxes,color=FG_DIM)
    ax.set_title("3D Reservoir Map",color=COL_DATASET,fontsize=13,pad=8,fontweight="bold")
    return fig

def plot_production(df, ycols, xcols, hex_colors):
    fig, ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    y_col = next((c for c in ycols if c in df.columns),None) or \
            next((c for c in xcols if c in df.columns),None)
    x_col = next((c for c in xcols if c in df.columns and c!=y_col),None)
    if y_col is None:
        ax.text(0.2,0.5,"Select a Y feature",transform=ax.transAxes,color=FG_DIM,fontsize=10); return fig
    for cl in sorted(df["cluster"].unique()):
        sub=df[df["cluster"]==cl].reset_index(drop=True); col=hex_colors[int(cl)%len(hex_colors)]
        xs = sub[x_col].values if x_col else np.arange(len(sub))
        ax.plot(xs,sub[y_col].values,color=col,linewidth=2,alpha=0.88,
                label=f"Cluster {cl}  ({len(sub):,} wells)")
    ax.set_xlabel(x_col if x_col else "Well Index",fontsize=9,labelpad=6)
    ax.set_ylabel(y_col,fontsize=9,labelpad=6)
    ax.set_title("Production Curves by Cluster",color=SUCCESS2,fontsize=13,pad=12,fontweight="bold")
    ax.margins(0.04,0.10)
    ax.legend(loc="upper left",fontsize=8.5,framealpha=0.92,
              edgecolor="#334155",facecolor="#161d2e",labelcolor="#e2e8f0")
    return fig

def plot_hidden_patterns(X_w, anomaly_labels, xcols, detector_name):
    from sklearn.decomposition import PCA
    fig, ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
    n = X_w.shape[0]
    if X_w.shape[1]>=2:
        Z=PCA(n_components=2,random_state=42).fit_transform(X_w)
        xlabel,ylabel="PC 1  (feature projection)","PC 2  (feature projection)"
    elif X_w.shape[1]==1:
        Z=np.column_stack([X_w[:,0],np.arange(n)])
        xlabel=xcols[0] if xcols else "Feature"; ylabel="Well Index"
    else:
        ax.text(0.3,0.5,"No feature data",transform=ax.transAxes,color=FG_DIM); return fig
    nm=anomaly_labels==1; am=anomaly_labels==-1
    ax.scatter(Z[nm,0],Z[nm,1],color=NORMAL_C,s=30,alpha=0.70,
               edgecolors="#ffffff18",linewidths=0.3,zorder=3,
               label=f"● Normal wells  ({int(nm.sum()):,})")
    if am.sum()>0:
        ax.scatter(Z[am,0],Z[am,1],color=ANOMALY_C,s=70,alpha=0.95,
                   edgecolors="#ffffff66",linewidths=0.9,marker="D",zorder=5,
                   label=f"◆ Anomalous wells  ({int(am.sum()):,})")
        ax.annotate(f"◆ {int(am.sum())} anomal{'y' if am.sum()==1 else 'ies'} detected",
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
        fig, ax = plt.subplots(figsize=(7.5,4.5)); _style_ax(ax)
        ax.text(0.3,0.5,"Run analysis first",transform=ax.transAxes,color=FG_DIM,fontsize=11)
        ax.set_title("Parameter Importance",color=COL_WEIGHTS,fontsize=13,pad=12,fontweight="bold")
        return fig
    params = list(weight_map.keys()); values = [weight_map[p] for p in params]
    sorted_pairs = sorted(zip(values,params),reverse=False)
    values_s=[v for v,_ in sorted_pairs]; params_s=[p for _,p in sorted_pairs]
    colors=[CLUSTER_PALETTE[i%len(CLUSTER_PALETTE)] for i in range(len(params_s))]
    fig = plt.figure(figsize=(7.5,4.5))
    gs  = gridspec.GridSpec(2,1,height_ratios=[5,1],hspace=0.45)
    ax  = fig.add_subplot(gs[0]); _style_ax(ax)
    ax_note = fig.add_subplot(gs[1]); ax_note.axis("off")
    bars = ax.barh(params_s,values_s,color=colors,edgecolor="#ffffff22",linewidth=0.5,height=0.6)
    for bar,val in zip(bars,values_s):
        ax.text(bar.get_width()+0.3,bar.get_y()+bar.get_height()/2,
                f"{val:.1f}%",va="center",ha="left",color=FG_MID,fontsize=8.5,fontweight="bold")
    max_idx=values_s.index(max(values_s)); bars[max_idx].set_edgecolor("#ffffff88"); bars[max_idx].set_linewidth(1.5)
    ax.text(values_s[max_idx]/2,bars[max_idx].get_y()+bars[max_idx].get_height()/2,
            "  ★ Most Influential",va="center",ha="left",color="#ffffff",fontsize=7.5,style="italic")
    ax.set_xlabel("Assigned Weight (%)",fontsize=9,labelpad=6)
    ax.set_title("Parameter Importance — User-Assigned Weights",
                 color=COL_WEIGHTS,fontsize=12,pad=12,fontweight="bold")
    ax.set_xlim(0,max(values_s)*1.22); ax.margins(0.03,0.15)
    wrapped=method_desc if len(method_desc)<=90 else method_desc[:87]+"…"
    ax_note.text(0.01,0.8,f"Method: {wrapped}",transform=ax_note.transAxes,
                 color=FG_DIM,fontsize=7.5,va="top")
    fig.patch.set_facecolor("#0d1117"); return fig

def plot_event_chart(df, event_result):
    fig = plt.figure(figsize=(7.5,4.5)); fig.patch.set_facecolor("#0d1117")
    if not event_result["has_events"]:
        ax=fig.add_subplot(111); _style_ax(ax)
        ax.text(0.5,0.58,"No Event Column Found",transform=ax.transAxes,
                ha="center",va="center",color=WARN,fontsize=14,fontweight="bold")
        ax.text(0.5,0.42,"Add a column named:\n'status','event','state','mode'",
                transform=ax.transAxes,ha="center",va="center",color=FG_DIM,fontsize=10)
        ax.set_title("Operational Event Analysis",color=COL_EVENTS,fontsize=13,pad=12,fontweight="bold")
        return fig
    ec=event_result["event_col"]; counts=event_result["event_counts"]
    labels=list(counts.keys()); values=list(counts.values()); total=max(sum(values),1)
    def _ec(label):
        ll=label.lower()
        for k,c in EVENT_COLOURS.items():
            if k in ll: return c
        return CLUSTER_PALETTE[hash(label)%len(CLUSTER_PALETTE)]
    bar_colors=[_ec(l) for l in labels]
    gs=gridspec.GridSpec(1,2,width_ratios=[3,2],wspace=0.38)
    ax1=fig.add_subplot(gs[0]); _style_ax(ax1)
    ax2=fig.add_subplot(gs[1]); _style_ax(ax2)
    sorted_pairs=sorted(zip(values,labels,bar_colors))
    vs=[v for v,_,_ in sorted_pairs]; ls=[l for _,l,_ in sorted_pairs]; cs=[c for _,_,c in sorted_pairs]
    bars=ax1.barh(ls,vs,color=cs,edgecolor="#ffffff22",linewidth=0.5,height=0.65)
    for bar,val in zip(bars,vs):
        ax1.text(bar.get_width()+0.3,bar.get_y()+bar.get_height()/2,
                 f"{val:,}  ({val/total*100:.1f}%)",va="center",ha="left",color=FG_MID,fontsize=8)
    ax1.set_xlabel("Well Count",fontsize=9,labelpad=6)
    ax1.set_title(f"Event Distribution\n(column: '{ec}')",color=COL_EVENTS,fontsize=10,pad=8,fontweight="bold")
    ax1.set_xlim(0,max(vs)*1.4 if vs else 1)
    wedge_props=dict(width=0.52,edgecolor="#0d1117",linewidth=2)
    pie_labels=[f"{l}\n{v:,}" for l,v in zip(labels,values)]
    ax2.pie(values,colors=bar_colors,labels=pie_labels,startangle=90,
            wedgeprops=wedge_props,textprops=dict(color=FG_MID,fontsize=7.5))
    ax2.set_title("Proportion",color=COL_EVENTS,fontsize=10,pad=8,fontweight="bold")
    op_count=len(event_result.get("operational_anomalies",[]))
    fig.text(0.5,0.01,
             f"Total: {total:,}  |  Anomalous: {event_result['n_abnormal']:,}  "
             f"|  Active: {event_result['n_active']:,}  |  Inactive: {event_result['n_inactive']:,}  "
             f"|  Op.Anomalies: {op_count:,}",
             ha="center",color=FG_DIM,fontsize=8)
    return fig


# ══════════════════════════════════════════════════════════════════════════════
# LEGEND BUILDERS
# ══════════════════════════════════════════════════════════════════════════════
def build_cluster_legend(parent, df, hex_colors):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="CLUSTER LEGEND",bg=CARD,fg=COL_CLUSTER,font=FONT_H3,
             justify="center").pack(pady=(12,4),padx=8)
    _divider(parent,COL_CLUSTER)
    clusters=sorted(df["cluster"].unique()); total=len(df)
    for cl in clusters:
        cnt=int((df["cluster"]==cl).sum()); pct=cnt/total*100
        col=hex_colors[int(cl)%len(hex_colors)]
        row=tk.Frame(parent,bg=CARD2,padx=6,pady=4,
                     highlightbackground=col,highlightthickness=1)
        row.pack(fill="x",padx=8,pady=3)
        tk.Canvas(row,bg=col,width=22,height=22,highlightthickness=0).pack(side="left",padx=(0,8))
        txt=tk.Frame(row,bg=CARD2); txt.pack(side="left",fill="x",expand=True)
        tk.Label(txt,text=f"Cluster {cl}",bg=CARD2,fg=col,font=FONT_H3,anchor="w").pack(anchor="w")
        tk.Label(txt,text=f"{cnt:,} wells  ({pct:.1f}%)",bg=CARD2,fg=FG_MID,
                 font=FONT_XS,anchor="w").pack(anchor="w")
    _divider(parent)
    for cl in clusters:
        cnt=int((df["cluster"]==cl).sum())
        _mini_bar(parent,f"Cl.{cl}",cnt/total,hex_colors[int(cl)%len(hex_colors)])

def build_anomaly_legend(parent, n_normal, n_anomaly):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="PATTERN LEGEND",bg=CARD,fg=ANOMALY_C,font=FONT_H3,
             justify="center").pack(pady=(12,4),padx=8)
    _divider(parent,ANOMALY_C)
    total=n_normal+n_anomaly
    for label,count,col,icon in [("Normal Wells",n_normal,NORMAL_C,"●"),
                                  ("Anomaly Wells",n_anomaly,ANOMALY_C,"◆")]:
        pct=count/total*100 if total>0 else 0
        row=tk.Frame(parent,bg=CARD2,padx=6,pady=4,
                     highlightbackground=col,highlightthickness=1)
        row.pack(fill="x",padx=8,pady=3)
        sw=tk.Canvas(row,bg=col,width=22,height=22,highlightthickness=0); sw.pack(side="left",padx=(0,8))
        sw.create_text(11,11,text=icon,fill="white",font=("Arial",10,"bold"))
        txt=tk.Frame(row,bg=CARD2); txt.pack(side="left",fill="x",expand=True)
        tk.Label(txt,text=label,bg=CARD2,fg=col,font=FONT_H3,anchor="w").pack(anchor="w")
        tk.Label(txt,text=f"{count:,} wells  ({pct:.1f}%)",bg=CARD2,fg=FG_MID,
                 font=FONT_XS,anchor="w").pack(anchor="w")
    if total:
        _divider(parent)
        _mini_bar(parent,"Norm",n_normal/total,NORMAL_C)
        _mini_bar(parent,"Anom",n_anomaly/total,ANOMALY_C)

def build_weight_legend(parent, weight_map, method_desc):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="PARAM WEIGHTS",bg=CARD,fg=COL_WEIGHTS,font=FONT_H3,
             justify="center").pack(pady=(10,2),padx=8)
    _divider(parent,COL_WEIGHTS)
    if not weight_map:
        tk.Label(parent,text="Run analysis first",bg=CARD,fg=FG_DIM,font=FONT_XS).pack(pady=8); return
    total_w=max(sum(weight_map.values()),1)
    sorted_wm=sorted(weight_map.items(),key=lambda x:-x[1])
    max_w=sorted_wm[0][1] if sorted_wm else 1
    for i,(param,wt) in enumerate(sorted_wm):
        col=CLUSTER_PALETTE[i%len(CLUSTER_PALETTE)]; is_max=(wt==max_w)
        row=tk.Frame(parent,bg=CARD2 if is_max else CARD,padx=6,pady=3,
                     highlightbackground=col if is_max else BORDER,highlightthickness=1)
        row.pack(fill="x",padx=8,pady=2)
        tk.Canvas(row,bg=col,width=18,height=18,highlightthickness=0).pack(side="left",padx=(0,6))
        txt=tk.Frame(row,bg=CARD2 if is_max else CARD); txt.pack(side="left",fill="x",expand=True)
        tk.Label(txt,text=f"{'★ ' if is_max else ''}{param[:22]}",
                 bg=CARD2 if is_max else CARD,fg=col,font=FONT_SM,anchor="w").pack(anchor="w")
        tk.Label(txt,text=f"{wt:.1f}%",bg=CARD2 if is_max else CARD,fg=FG_MID,
                 font=FONT_XS,anchor="w").pack(anchor="w")
        _mini_bar(parent,"",wt/total_w,col)
    _divider(parent)
    tk.Label(parent,text="User-assigned weights\nfrom dataset columns",
             bg=CARD,fg=FG_DIM,font=FONT_XS,wraplength=175,justify="left").pack(anchor="w",padx=8,pady=(2,4))

def build_event_legend(parent, event_result):
    for w in parent.winfo_children(): w.destroy()
    tk.Label(parent,text="EVENT LEGEND",bg=CARD,fg=COL_EVENTS,font=FONT_H3,
             justify="center").pack(pady=(10,2),padx=8)
    _divider(parent,COL_EVENTS)
    if not event_result["has_events"]:
        tk.Label(parent,text="No event column\nin dataset",bg=CARD,fg=WARN,
                 font=FONT_SM,justify="center").pack(pady=8); return
    counts=event_result["event_counts"]; total=max(sum(counts.values()),1)
    def _c(label):
        ll=label.lower()
        for k,c in EVENT_COLOURS.items():
            if k in ll: return c
        return CLUSTER_PALETTE[hash(label)%len(CLUSTER_PALETTE)]
    for label,cnt in sorted(counts.items(),key=lambda x:-x[1]):
        col=_c(label)
        row=tk.Frame(parent,bg=CARD2,padx=6,pady=3,
                     highlightbackground=BORDER,highlightthickness=1)
        row.pack(fill="x",padx=8,pady=2)
        tk.Canvas(row,bg=col,width=18,height=18,highlightthickness=0).pack(side="left",padx=(0,6))
        txt=tk.Frame(row,bg=CARD2); txt.pack(side="left",fill="x",expand=True)
        tk.Label(txt,text=label[:20],bg=CARD2,fg=col,font=FONT_SM,anchor="w").pack(anchor="w")
        tk.Label(txt,text=f"{cnt:,}  ({cnt/total*100:.1f}%)",bg=CARD2,fg=FG_MID,
                 font=FONT_XS,anchor="w").pack(anchor="w")
        _mini_bar(parent,"",cnt/total,col)
    _divider(parent)
    tk.Label(parent,text=f"Source:\n'{event_result['event_col']}'",
             bg=CARD,fg=FG_DIM,font=FONT_XS,justify="left").pack(anchor="w",padx=8,pady=(2,4))


# ══════════════════════════════════════════════════════════════════════════════
# CHART FRAME
# ══════════════════════════════════════════════════════════════════════════════
def _fill_toolbar(bar, mpl_canvas, tab_name):
    tk.Label(bar,text=" TOOLS:",bg="#080e1a",fg=FG_DIM,font=FONT_XS).pack(side="left",padx=(6,2))
    nav_frame=tk.Frame(bar,bg="#080e1a"); nav_frame.pack(side="left",padx=2)
    tb=NavigationToolbar2Tk(mpl_canvas,nav_frame); tb.config(bg="#080e1a")
    for child in tb.winfo_children():
        try: child.config(bg="#080e1a",fg=FG_MID,activebackground=CARD2,activeforeground=FG,
                          relief="flat",bd=0,highlightthickness=0,font=FONT_XS)
        except: pass
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
    try: fig.tight_layout(pad=2.5,rect=[0.03,0.03,0.97,0.95])
    except: pass
    fig.patch.set_facecolor("#0d1117")
    bar=tk.Frame(tab_frame,bg="#080e1a",pady=4,
                 highlightbackground=BORDER,highlightthickness=1)
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
def get_selected_x():
    return [x_listbox.get(i).strip() for i in x_listbox.curselection()]

def resolve_features():
    if raw_data is None: return [],"",0,"No dataset loaded"
    all_num=[c for c in raw_data.select_dtypes(include="number").columns]
    x_sel=[c for c in get_selected_x() if c in raw_data.columns]
    x_cols=x_sel or all_num
    needed=list(dict.fromkeys(x_cols))
    wdf=raw_data.dropna(subset=needed).reset_index(drop=True)
    warn="No X selected — all numeric cols used" if not x_sel else ""
    return x_cols, wdf, len(wdf), warn

def start_pipeline():
    if raw_data is None:
        messagebox.showwarning("No Data","Please upload a dataset first."); return
    set_status("Running analysis…",WARN); progress_start()
    threading.Thread(target=run_pipeline,daemon=True).start()

def run_pipeline():
    global active_df,active_X,active_xcols,active_anomaly_result,active_weight_result,active_event_result
    x_cols,wdf,well_count,warn = resolve_features()
    if wdf is None or well_count==0:
        app.after(0, lambda: set_status(f"No usable rows: {warn}",WARN))
        app.after(0, progress_stop); return
    try:
        weight_map,raw_total=get_manual_weights()
        method_desc="User-assigned weights (from dataset numeric columns)"
        if not weight_map:
            weight_map={c:round(100/len(x_cols),2) for c in x_cols}
            method_desc="Equal weights (set weights in sidebar)"
        final_weights={c:weight_map.get(c,1.0) for c in x_cols}
        total_fw=sum(final_weights.values())
        final_weights={c:v/total_fw*100 for c,v in final_weights.items()}
        X_w,use_cols=build_weighted_X(wdf,x_cols,final_weights)
        n_clusters=min(cluster_var.get(),well_count)
        wdf=wdf.copy(); cluster_labels,inertia=run_clustering(X_w,n_clusters)
        wdf["cluster"]=cluster_labels
        method=anomaly_method_var.get(); contam=anomaly_contam_var.get()/100.0
        a_labels,a_pct,a_idx,a_name,a_scores=detect_anomalies(X_w,contamination=contam,method=method)
        wdf["anomaly"]=a_labels
        anomaly_result={"labels":a_labels,"pct":a_pct,"indices":a_idx,
                        "scores":a_scores,"detector_name":a_name,
                        "n_anomaly":len(a_idx),"n_normal":len(a_labels)-len(a_idx)}
        event_result=analyse_events(wdf,a_labels)
        active_df=wdf; active_X=X_w; active_xcols=use_cols
        active_anomaly_result=anomaly_result
        active_weight_result={"weight_map":final_weights,"method_desc":method_desc}
        active_event_result=event_result
        hx=cluster_hex(n_clusters)
        y_cols=[c for c in wdf.select_dtypes(include="number").columns
                if c not in use_cols and c!="cluster"][:2]
        insights=generate_insights(wdf,use_cols,well_count,hx,anomaly_result,
                                   event_result,final_weights,method_desc,inertia)
        app.after(0,lambda: refresh_ui(wdf,X_w,use_cols,y_cols,hx,insights,well_count,
                                       anomaly_result,event_result,final_weights,method_desc))
        app.after(0,lambda: set_status("Analysis complete",SUCCESS))
        app.after(0,progress_stop)
        app.after(0,lambda: export_btn.config(state="normal"))
    except Exception as e:
        traceback.print_exc()
        app.after(0,lambda err=str(e): set_status(f"Error: {err}",ERR))
        app.after(0,progress_stop)

def refresh_ui(df,X,x_cols,y_cols,hx,insights,well_count,
               anomaly_result,event_result,weight_map,method_desc):
    draw_plot(cluster_tab,    plot_clusters(df,x_cols,y_cols,hx),"clusters",
              legend_builder=build_cluster_legend,legend_kwargs={"df":df,"hex_colors":hx})
    draw_plot(pca_tab,        plot_pca(X,df["cluster"].values,hx),"pca",
              legend_builder=build_cluster_legend,legend_kwargs={"df":df,"hex_colors":hx})
    draw_plot(reservoir_tab,  plot_reservoir_3d(df,x_cols,y_cols,hx),"reservoir_3d",
              legend_builder=build_cluster_legend,legend_kwargs={"df":df,"hex_colors":hx})
    draw_plot(production_tab, plot_production(df,y_cols,x_cols,hx),"production",
              legend_builder=build_cluster_legend,legend_kwargs={"df":df,"hex_colors":hx})
    draw_plot(hidden_tab,
              plot_hidden_patterns(X,anomaly_result["labels"],x_cols,anomaly_result["detector_name"]),
              "hidden_patterns",
              legend_builder=build_anomaly_legend,
              legend_kwargs={"n_normal":anomaly_result["n_normal"],"n_anomaly":anomaly_result["n_anomaly"]})
    draw_plot(weight_tab, plot_weight_chart(weight_map,method_desc),"param_weights",
              legend_builder=build_weight_legend,
              legend_kwargs={"weight_map":weight_map,"method_desc":method_desc})
    draw_plot(event_tab,  plot_event_chart(df,event_result),"event_summary",
              legend_builder=build_event_legend,legend_kwargs={"event_result":event_result})
    stat_widgets["Total Wells"].set(f"{well_count:,}")
    stat_widgets["Analysed"].set(f"{len(df):,}")
    stat_widgets["Clusters"].set(str(df["cluster"].nunique()))
    stat_widgets["Anomalies"].set(f"{anomaly_result['n_anomaly']}  ({anomaly_result['pct']:.1f}%)")
    if event_result["has_events"]:
        stat_widgets["Active Wells"].set(f"{event_result['n_active']:,}")
        stat_widgets["Inactive Wells"].set(f"{event_result['n_inactive']:,}")
        stat_widgets["Event Types"].set(str(len(event_result["event_counts"])))
    else:
        stat_widgets["Active Wells"].set("N/A")
        stat_widgets["Inactive Wells"].set("N/A")
        stat_widgets["Event Types"].set("N/A")
    stat_widgets["Abnormal"].set(f"{event_result['n_abnormal']:,}")
    update_event_panel(event_result)
    explain_text.config(state="normal"); explain_text.delete("1.0","end")
    explain_text.insert("end",insights); explain_text.config(state="disabled")

def update_event_panel(event_result):
    if not event_result["has_events"]:
        ev_status_var.set("No event column in dataset"); ev_status_lbl.config(fg=WARN)
        for v in ev_count_vars.values(): v.set("N/A")
        ev_active_var.set("N/A"); ev_inactive_var.set("N/A")
        ev_abnormal_var.set("N/A"); ev_op_anom_var.set("N/A"); return
    ev_status_var.set(f"Column: '{event_result['event_col']}'"); ev_status_lbl.config(fg=SUCCESS2)
    counts=event_result["event_counts"]
    top=sorted(counts.items(),key=lambda x:-x[1])[:6]
    keys=list(ev_count_vars.keys())
    for i,k in enumerate(keys):
        if i<len(top):
            label,cnt=top[i]; ev_count_vars[k].set(f"{cnt:,}"); ev_count_lbls[k].config(text=label[:24])
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
                cl_sum=active_df.groupby("cluster").agg(well_count=("cluster","count")).reset_index()
                cl_sum.to_excel(writer,sheet_name="Cluster Summary",index=False)
                if active_weight_result:
                    wm=active_weight_result["weight_map"]
                    wdf_exp=pd.DataFrame([{"Parameter":k,"Weight_%":v}
                                         for k,v in sorted(wm.items(),key=lambda x:-x[1])])
                    wdf_exp["Method"]=active_weight_result["method_desc"]
                    wdf_exp.to_excel(writer,sheet_name="Parameter Weights",index=False)
                # ★ NEW — Inspector report sheet
                if merge_report:
                    insp_txt=_build_inspector_report(merge_report)
                    pd.DataFrame({"Inspector Report":insp_txt.split("\n")}).to_excel(
                        writer,sheet_name="Data Inspector",index=False)
                report_txt=explain_text.get("1.0","end").strip()
                pd.DataFrame({"Report":report_txt.split("\n")}).to_excel(
                    writer,sheet_name="Insights Report",index=False)
            set_status(f"Exported (6 sheets): {os.path.basename(path)}",SUCCESS)
            messagebox.showinfo("Export Complete",
                                f"Saved: {path}\n\nSheets:\n"
                                "  All Wells\n  Anomaly Wells\n  Cluster Summary\n"
                                "  Parameter Weights\n  Data Inspector\n  Insights Report")
        else:
            _save_df(active_df,path)
            set_status(f"Exported: {os.path.basename(path)}",SUCCESS)
    except Exception as e:
        messagebox.showerror("Export Error",str(e))
        set_status(f"Export failed: {e}",ERR)


# ══════════════════════════════════════════════════════════════════════════════
# INSIGHTS
# ══════════════════════════════════════════════════════════════════════════════
def generate_insights(df,x_cols,well_count,hx,anomaly_result,
                      event_result,weight_map,method_desc,inertia):
    clusters=df["cluster"].value_counts().sort_index()
    n_anom=anomaly_result["n_anomaly"]; a_pct=anomaly_result["pct"]
    sev=("LOW" if a_pct<2 else "MODERATE" if a_pct<8 else "ELEVATED" if a_pct<20 else "HIGH")
    interp={"LOW":"Very few outliers. Reservoir is relatively uniform.",
             "MODERATE":"Moderate anomalies. Check localised heterogeneity.",
             "ELEVATED":"Significant complexity. Variable seam or fractures.",
             "HIGH":"Large fraction flagged. Review contamination % or data quality."}[sev]
    lines=["━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
           "  CBM AI ANALYSIS REPORT  v6.0",
           f"  {datetime.datetime.now().strftime('%Y-%m-%d  %H:%M')}",
           "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
           f"  Total wells : {well_count:,}",
           f"  Analysed    : {len(df):,}",
           f"  Clusters    : {len(clusters)}",
           f"  Features    : {', '.join(x_cols)}",
           f"  KMeans inertia: {inertia:.2f}",
           "","  PARAMETER IMPORTANCE","  ──────────────────────────────────────────",
           f"  Method: {method_desc}",""]
    if weight_map:
        max_w=max(weight_map.values())
        for param,wt in sorted(weight_map.items(),key=lambda x:-x[1]):
            star=" ★" if wt==max_w else ""; bar="█"*max(1,int(wt/2))
            lines.append(f"  {param:<30} {bar:<25} {wt:.1f}%{star}")
    lines+=["","  CLUSTER BREAKDOWN","  ──────────────────────────────────────────"]
    for cl,cnt in clusters.items():
        pct=cnt/len(df)*100; col=hx[int(cl)%len(hx)]; bar="█"*int(pct/3)
        lines.append(f"  Cluster {cl} [{col}]  {bar:<25} {cnt:>5,}  {pct:5.1f}%")
    lines+=["","━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
            "  ANOMALY DETECTION",
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
            f"  Algorithm    : {anomaly_result['detector_name']}",
            f"  Normal wells : {anomaly_result['n_normal']:,}",
            f"  Anomalies    : {n_anom:,}  ({a_pct:.1f}%)  [{sev}]",
            f"  Interpretation: {interp}",
            "","━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
            "  OPERATIONAL EVENTS",
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"]
    if event_result["has_events"]:
        op_count=len(event_result.get("operational_anomalies",[]))
        lines+=[f"  Source column  : '{event_result['event_col']}'",
                f"  Active wells   : {event_result['n_active']:,}",
                f"  Inactive wells : {event_result['n_inactive']:,}",
                f"  Abnormal       : {event_result['n_abnormal']:,}",
                f"  Op. anomalies  : {op_count:,}","","  Event breakdown:"]
        for ev,cnt in sorted(event_result["event_counts"].items(),key=lambda x:-x[1]):
            pct=cnt/len(df)*100; bar="█"*max(1,int(pct/3))
            lines.append(f"    {ev:<30} {bar:<20} {cnt:>5,}  {pct:.1f}%")
    else:
        lines.append("  "+event_result["message"])
    lines+=["","━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"]
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
if HAVE_TTKBS:
    app = ttk.Window(themename="darkly")
else:
    app = tk.Tk()
    app.configure(bg=BG_DEEP)

app.title("CBM AI Analytics Platform  v6.0  —  Multi-Table Inspector")
app.geometry("1760x1000")
app.configure(bg=BG_DEEP)
app.minsize(1280,720)

# Top bar
top_bar = tk.Frame(app, bg="#040810", height=46)
top_bar.pack(fill="x", side="top"); top_bar.pack_propagate(False)
brand = tk.Frame(top_bar, bg="#040810"); brand.pack(side="left", padx=16, fill="y")
tk.Label(brand, text="CBM·AI", bg="#040810", fg=COL_DATASET,
         font=("Georgia",17,"bold")).pack(side="left", padx=(0,10))
tk.Label(brand, text="Coalbed Methane Analytics Platform  v6.0",
         bg="#040810", fg=FG_DIM, font=FONT_SM).pack(side="left")
tk.Label(top_bar,
         text="Multi-table Access  ·  Auto-join  ·  Data Inspector  ·  Dynamic weights  ·  Event analysis",
         bg="#040810", fg=FG_DIM, font=FONT_XS).pack(side="right", padx=14)
tk.Frame(app, bg=BORDER, height=1).pack(fill="x")

root_pane = tk.PanedWindow(app, orient="horizontal", bg=BG_DEEP,
                           sashwidth=4, sashrelief="flat", sashpad=0, handlesize=0)
root_pane.pack(fill="both", expand=True)

# ══════════════════════════════════════════════════════════════════════════════
# LEFT SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
sb_outer, sidebar = make_scrollable(root_pane, bg=SIDEBAR)
root_pane.add(sb_outer, width=440, minsize=400)
tk.Frame(sidebar, bg=SIDEBAR, height=8).pack()

# ① DATASET
ds_body = make_section_card(sidebar, "① DATASET  (Access → auto-merges all tables)", COL_DATASET)
make_btn(ds_body, "  Upload Data File", upload_dataset,
         color="#0e7490", fg_col="#e0f7fa",
         tip="Access .accdb/.mdb → loads dbo_PRD + dbo_SS + dbo_XY and merges automatically")
info_row = tk.Frame(ds_body, bg=CARD); info_row.pack(fill="x", pady=(6,0))
rows_var       = tk.StringVar(value="—")
cols_var       = tk.StringVar(value="—")
num_cols_var   = tk.StringVar(value="—")
table_count_var= tk.StringVar(value="—")
for title, var, col in [("Total Rows", rows_var, COL_DATASET),
                        ("Columns",    cols_var, FG_DIM),
                        ("Numeric",    num_cols_var, COL_WEIGHTS),
                        ("Tables",     table_count_var, COL_INSPECT)]:
    cf = tk.Frame(info_row, bg=CARD); cf.pack(side="left", expand=True, fill="x")
    tk.Label(cf, text=title, bg=CARD, fg=FG_DIM, font=FONT_XS).pack(anchor="w")
    tk.Label(cf, textvariable=var, bg=CARD, fg=col,
             font=("Courier New",11,"bold")).pack(anchor="w")
event_col_var = tk.StringVar(value="Upload dataset to detect event column")
event_col_lbl = tk.Label(ds_body, textvariable=event_col_var, bg=CARD, fg=FG_DIM,
                         font=FONT_XS, wraplength=400, justify="left", anchor="w")
event_col_lbl.pack(fill="x", pady=(6,0))

# ② CLUSTER SETTINGS
cl_body = make_section_card(sidebar, "② CLUSTER SETTINGS", COL_CLUSTER)
tk.Label(cl_body, text="Number of Clusters  (2 – 10)",
         bg=CARD, fg=FG_DIM, font=FONT_SM).pack(anchor="w")
cluster_var = tk.IntVar(value=3)
sf = tk.Frame(cl_body, bg=CARD); sf.pack(fill="x", pady=4)
tk_ttk.Spinbox(sf, from_=2, to=10, textvariable=cluster_var,
               width=5, font=FONT_H2).pack(side="left")
tk.Label(sf, text="clusters", bg=CARD, fg=FG_DIM, font=FONT_SM).pack(side="left", padx=8)

# ③ PARAMETER WEIGHTAGE
wt_outer = make_section_card(sidebar, "③ PARAMETER WEIGHTAGE", COL_WEIGHTS)
weight_sum_var = tk.StringVar(value="Upload a dataset to see parameters")
weight_sum_lbl = tk.Label(wt_outer, textvariable=weight_sum_var, bg=CARD, fg=FG_DIM,
                          font=("Courier New",9,"bold"), anchor="w")
weight_sum_lbl.pack(anchor="w", pady=(0,4))
wt_canvas_outer = tk.Frame(wt_outer, bg=CARD, height=180)
wt_canvas_outer.pack(fill="x"); wt_canvas_outer.pack_propagate(False)
wt_canvas = Canvas(wt_canvas_outer, bg=CARD, highlightthickness=0, bd=0)
wt_sb = Scrollbar(wt_canvas_outer, orient="vertical", command=wt_canvas.yview,
                  bg=BORDER, troughcolor=CARD, activebackground=COL_WEIGHTS)
wt_canvas.configure(yscrollcommand=wt_sb.set)
wt_sb.pack(side="right", fill="y"); wt_canvas.pack(side="left", fill="both", expand=True)
weight_body = tk.Frame(wt_canvas, bg=CARD)
wt_wid = wt_canvas.create_window((0,0), window=weight_body, anchor="nw")
weight_body.bind("<Configure>", lambda e: wt_canvas.configure(scrollregion=wt_canvas.bbox("all")))
wt_canvas.bind("<Configure>", lambda e: (
    wt_canvas.configure(scrollregion=wt_canvas.bbox("all")),
    wt_canvas.itemconfig(wt_wid, width=e.width)
))
_wt_placeholder = tk.Label(weight_body, text="Upload a dataset first.",
                           bg=CARD, fg=FG_DIM, font=FONT_XS, justify="left", padx=8, pady=8)
_wt_placeholder.pack(anchor="w"); weight_row_frames = [_wt_placeholder]

# ④ ANOMALY DETECTION
anom_body = make_section_card(sidebar, "④ ANOMALY DETECTION", COL_ANOMALY)
tk.Label(anom_body, text="Algorithm", bg=CARD, fg=FG_DIM, font=FONT_SM).pack(anchor="w")
anomaly_method_var = tk.StringVar(value="iforest")
mrow = tk.Frame(anom_body, bg=CARD); mrow.pack(fill="x", pady=(2,6))
for label, val in [("Isolation Forest  (faster)","iforest"),
                   ("Local Outlier Factor  (density)","lof")]:
    tk.Radiobutton(mrow, text=label, variable=anomaly_method_var, value=val,
                   bg=CARD, fg=FG_MID, selectcolor=CARD2, activebackground=CARD,
                   activeforeground=COL_ANOMALY, font=FONT_SM).pack(anchor="w", pady=1)
tk.Label(anom_body, text="Contamination %", bg=CARD, fg=FG_DIM, font=FONT_SM).pack(anchor="w")
anomaly_contam_var = tk.DoubleVar(value=5.0)
crow = tk.Frame(anom_body, bg=CARD); crow.pack(fill="x", pady=2)
tk_ttk.Spinbox(crow, from_=1.0, to=49.0, increment=0.5,
               textvariable=anomaly_contam_var, width=6, font=FONT_H3).pack(side="left")
tk.Label(crow, text="%  (typical: 3–10%)", bg=CARD, fg=FG_DIM, font=FONT_XS).pack(side="left", padx=6)

# ⑤ FEATURE SELECTION
feat_body = make_section_card(sidebar, "⑤ FEATURE SELECTION", COL_FEATURES)
tk.Label(feat_body, text="Ctrl+click = multi-select  |  blank = all numeric",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=400, justify="left").pack(anchor="w")
x_listbox = Listbox(feat_body, height=5, selectmode="multiple",
                    bg="#080f1e", fg=FG, selectbackground=COL_FEATURES,
                    selectforeground="#fff", font=FONT_SM, relief="flat", bd=0,
                    highlightthickness=1, highlightcolor=COL_FEATURES, highlightbackground=BORDER)
x_listbox.pack(fill="x", pady=(4,2))

# ⑥ RUN
run_body = make_section_card(sidebar, "⑥ RUN ANALYSIS", "#6366f1")
make_btn(run_body, "  Run AI Analysis", start_pipeline,
         color="#4f46e5", tip="Cluster + anomaly + event analysis on merged table")
if HAVE_TTKBS:
    progress_bar = ttk.Progressbar(run_body, mode="determinate", bootstyle="info-striped", length=420)
else:
    progress_bar = tk_ttk.Progressbar(run_body, mode="determinate", length=420)
progress_bar.pack(fill="x", pady=(4,2))
status_var = tk.StringVar(value="Ready — upload an Access file (.accdb/.mdb) to begin")
status_lbl = tk.Label(run_body, textvariable=status_var, bg=CARD, fg=FG_DIM,
                      font=FONT_SM, anchor="w", wraplength=420, justify="left")
status_lbl.pack(fill="x", pady=(2,0))

# ⑦ EVENT DETECTION
ev_body = make_section_card(sidebar, "⑦ EVENT DETECTION", COL_EVENTS)
ev_status_var = tk.StringVar(value="Awaiting analysis")
ev_status_lbl = tk.Label(ev_body, textvariable=ev_status_var, bg=CARD, fg=FG_DIM,
                         font=FONT_SM, wraplength=420, justify="left")
ev_status_lbl.pack(anchor="w", pady=(0,6))
ev_tile_row = tk.Frame(ev_body, bg=CARD); ev_tile_row.pack(fill="x", pady=(0,6))
ev_active_var  = tk.StringVar(value="—")
ev_inactive_var= tk.StringVar(value="—")
ev_abnormal_var= tk.StringVar(value="—")
ev_op_anom_var = tk.StringVar(value="—")
for title, var, col in [("Active",ev_active_var,COL_EVENTS),("Inactive",ev_inactive_var,COL_ANOMALY),
                        ("Abnormal",ev_abnormal_var,WARN),("Op.Anom.",ev_op_anom_var,COL_WEIGHTS)]:
    cell=tk.Frame(ev_tile_row,bg=CARD2,padx=5,pady=5,
                  highlightbackground=BORDER,highlightthickness=1)
    cell.pack(side="left",expand=True,fill="x",padx=2)
    tk.Label(cell,textvariable=var,bg=CARD2,fg=col,font=("Courier New",11,"bold")).pack()
    tk.Label(cell,text=title,bg=CARD2,fg=FG_DIM,font=FONT_XS).pack()
_ev_label_names = [f"ev{i}" for i in range(6)]
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

# ⑧ EXPORT
exp_body = make_section_card(sidebar, "⑧ EXPORT RESULTS", COL_EXPORT)
export_btn = tk.Button(exp_body, text="  Export All Results  (Excel 6-sheet)",
                       command=export_results, bg="#14532d", fg=COL_EXPORT,
                       activebackground=_lighten("#14532d",20), activeforeground=COL_EXPORT,
                       font=FONT_H3, relief="flat", bd=0, cursor="hand2",
                       padx=10, pady=11, anchor="w", state="disabled")
export_btn.pack(fill="x", pady=2)
export_btn.bind("<Enter>", lambda e: export_btn.config(bg=_lighten("#14532d",20)))
export_btn.bind("<Leave>", lambda e: export_btn.config(bg="#14532d"))
Tooltip(export_btn, "Export 6-sheet Excel: Wells · Anomalies · Clusters · Weights · Inspector · Insights")

# ⑨ DATASET PREVIEW
prev_body = make_section_card(sidebar, "⑨ DATASET PREVIEW  (first 20 rows of merged table)",
                              COL_PREVIEW, fill="both", expand=True)
table_frame = tk.Frame(prev_body, bg=CARD); table_frame.pack(fill="both", expand=True)
tk.Frame(sidebar, bg=SIDEBAR, height=20).pack()

# ══════════════════════════════════════════════════════════════════════════════
# RIGHT PANEL
# ══════════════════════════════════════════════════════════════════════════════
rp_outer, right_inner = make_scrollable(root_pane, bg=BG)
root_pane.add(rp_outer, minsize=860)

rp_hdr = tk.Frame(right_inner, bg=BG, padx=16, pady=10); rp_hdr.pack(fill="x")
tk.Label(rp_hdr, text="Analysis Dashboard", bg=BG, fg=FG,
         font=("Georgia",15,"bold")).pack(side="left")
tk.Label(rp_hdr,
         text="Multi-table Access  ·  Auto-join  ·  Data Inspector  ·  Dynamic weights  ·  Events",
         bg=BG, fg=FG_DIM, font=FONT_SM).pack(side="right", pady=2)
tk.Frame(right_inner, bg=BORDER, height=1).pack(fill="x", padx=12)

# Stats
stats_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=5); stats_wrap.pack(fill="x")
stat_defs_r1=[("Total Wells","—",COL_DATASET),("Analysed","—",COL_CLUSTER),
              ("Clusters","—",COL_WEIGHTS),   ("Anomalies","—",COL_ANOMALY)]
stat_defs_r2=[("Active Wells","—",COL_EVENTS),("Inactive Wells","—",ERR),
              ("Event Types","—",WARN),        ("Abnormal","—",COL_ANOMALY)]
stat_widgets={}
for stat_row_defs in [stat_defs_r1, stat_defs_r2]:
    row=tk.Frame(stats_wrap,bg=BG); row.pack(fill="x",pady=2)
    for title,val,col in stat_row_defs:
        sf=tk.Frame(row,bg=CARD,padx=10,pady=7,
                    highlightbackground=col,highlightthickness=1)
        sf.pack(side="left",expand=True,fill="x",padx=3)
        sv=tk.StringVar(value=val)
        tk.Label(sf,textvariable=sv,bg=CARD,fg=col,font=("Courier New",12,"bold")).pack()
        tk.Label(sf,text=title,bg=CARD,fg=FG_DIM,font=FONT_XS).pack()
        stat_widgets[title]=sv

# Notebook — 8 tabs (7 original + 1 Inspector)
nb_wrap=tk.Frame(right_inner,bg=BG,padx=10,pady=6); nb_wrap.pack(fill="x")
sty=tk_ttk.Style()
sty.configure("CBM.TNotebook",     background=BG,tabmargins=[0,0,0,0])
sty.configure("CBM.TNotebook.Tab", background=CARD2,foreground=FG_DIM,
              padding=[12,6],font=FONT_H3)
sty.map("CBM.TNotebook.Tab",
        background=[("selected","#1e3a8a")],foreground=[("selected","#93c5fd")])
notebook=tk_ttk.Notebook(nb_wrap,style="CBM.TNotebook"); notebook.pack(fill="x")

TAB_H=500
cluster_tab    = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
pca_tab        = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
reservoir_tab  = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
production_tab = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
hidden_tab     = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
weight_tab     = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
event_tab      = tk.Frame(notebook,bg=BG_MID,height=TAB_H)
inspector_tab  = tk.Frame(notebook,bg=BG_MID,height=TAB_H)   # ★ NEW

for tab,name in [(cluster_tab,"  Clusters  "),(pca_tab,"  PCA  "),
                 (reservoir_tab,"  3D Reservoir  "),(production_tab,"  Production  "),
                 (hidden_tab,"  Hidden Patterns  "),(weight_tab,"  Param Weights  "),
                 (event_tab,"  Events  "),(inspector_tab,"  Data Inspector  ")]:
    tab.pack_propagate(False); notebook.add(tab,text=name)

# ── Inspector Tab Layout ──────────────────────────────────────────────────────
insp_pane = tk.PanedWindow(inspector_tab, orient="horizontal", bg=BG_MID,
                           sashwidth=4, sashrelief="flat")
insp_pane.pack(fill="both", expand=True)

# Left: relationship tiles
insp_left_outer, insp_left = make_scrollable(insp_pane, bg=CARD)
insp_pane.add(insp_left_outer, width=280, minsize=220)
tk.Label(insp_left, text="TABLE RELATIONSHIPS", bg=CARD, fg=COL_INSPECT,
         font=FONT_H3).pack(anchor="w", padx=8, pady=(10,4))
_divider(insp_left, COL_INSPECT)
rel_tiles_frame = tk.Frame(insp_left, bg=CARD)
rel_tiles_frame.pack(fill="x", padx=4)
tk.Label(rel_tiles_frame, text="Upload an Access file\nto detect relationships.",
         bg=CARD, fg=FG_DIM, font=FONT_SM, justify="left").pack(anchor="w", padx=8, pady=8)

# Right: inspector text
insp_right = tk.Frame(insp_pane, bg=BG_MID)
insp_pane.add(insp_right, minsize=500)
insp_hdr2 = tk.Frame(insp_right, bg=BG_MID, padx=6, pady=6); insp_hdr2.pack(fill="x")
tk.Label(insp_hdr2, text="Full Inspector Report", bg=BG_MID, fg=FG,
         font=FONT_H2).pack(side="left")
tk.Button(insp_hdr2, text="Copy",
          command=lambda: (app.clipboard_clear(),
                           app.clipboard_append(inspector_text.get("1.0","end"))),
          bg=CARD2, fg=FG_DIM, activebackground=BORDER, activeforeground=FG,
          font=FONT_XS, relief="flat", bd=0, cursor="hand2", padx=10, pady=4).pack(side="right")
inspector_text = tk.Text(insp_right, bg=CARD, fg=COL_INSPECT,
                         font=("Courier New",9), relief="flat", bd=0,
                         padx=12, pady=10, insertbackground=FG, wrap="word",
                         highlightbackground=BORDER, highlightthickness=1, state="disabled")
inspector_text.pack(fill="both", expand=True, padx=6, pady=(0,6))

# ── Insights ──────────────────────────────────────────────────────────────────
ins_wrap=tk.Frame(right_inner,bg=BG,padx=10,pady=6); ins_wrap.pack(fill="x")
ins_hdr=tk.Frame(ins_wrap,bg=BG); ins_hdr.pack(fill="x",pady=(0,4))
tk.Label(ins_hdr,text="AI Insights Report",bg=BG,fg=FG,font=FONT_H2).pack(side="left")
tk.Button(ins_hdr,text="Copy",
          command=lambda:(app.clipboard_clear(),
                          app.clipboard_append(explain_text.get("1.0","end"))),
          bg=CARD2,fg=FG_DIM,activebackground=BORDER,activeforeground=FG,
          font=FONT_XS,relief="flat",bd=0,cursor="hand2",padx=10,pady=4).pack(side="right")
explain_text=tk.Text(ins_wrap,height=22,bg=CARD,fg=SUCCESS2,
                     font=("Courier New",9),relief="flat",bd=0,
                     padx=12,pady=10,insertbackground=FG,wrap="word",
                     highlightbackground=BORDER,highlightthickness=1,state="disabled")
explain_text.pack(fill="x")
tk.Frame(right_inner,bg=BG,height=20).pack()

# Welcome text
explain_text.config(state="normal")
explain_text.insert("end",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
    "  CBM AI Analytics Platform  v6.0\n"
    "  Multi-Table Access Inspector\n"
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
    "  WHAT'S NEW IN v6.0:\n\n"
    "  ★  Multi-table Access loader\n"
    "     Loads dbo_PRD + dbo_SS + dbo_XY\n"
    "     simultaneously — no table picker needed\n\n"
    "  ★  Auto-join engine\n"
    "     Detects common keys, scores candidates,\n"
    "     executes LEFT JOIN chain automatically\n\n"
    "  ★  Data Inspector tab\n"
    "     Per-table structure, null counts,\n"
    "     duplicate detection, relationship tiles,\n"
    "     post-merge diagnostics, SQL suggestion\n\n"
    "  ★  Inspector exported as sheet 5 in Excel\n\n"
    "  QUICK START:\n\n"
    "  ①  Upload your .accdb or .mdb file\n"
    "      All tables load & merge automatically\n\n"
    "  ②  Review the Data Inspector tab\n"
    "      (opens automatically after load)\n\n"
    "  ③  Adjust cluster count and weights\n\n"
    "  ④  Click Run AI Analysis\n\n"
    "  ⑤  Export 6-sheet Excel report\n"
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
)
explain_text.config(state="disabled")

app.mainloop()
