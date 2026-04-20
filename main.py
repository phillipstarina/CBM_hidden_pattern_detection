"""
CBM AI Analytics Platform — v2.1 (Bug Fixed)
─────────────────────────────────────────────
FIX: Toolbar TclError resolved — toolbar frame is created FIRST (pack top),
     then the chart+legend row is packed below it. No re-parenting needed.
ALSO:
  • Sidebar 380px — no text clipping
  • Export buttons: green/red/amber/blue distinct colours
  • In-chart matplotlib legend per cluster with colour swatch
  • Right panel legend with large colour squares + counts
  • Distinct cluster colours (not just plasma)
"""

import tkinter as tk
from tkinter import filedialog, Listbox, Canvas, Scrollbar, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import numpy as np
import matplotlib.patches as mpatches
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.pyplot as plt
import threading
import traceback
import os
import datetime
import matplotlib
matplotlib.use("TkAgg")

try:
    from utils.loader import load_dataset
except ImportError:
    def load_dataset(f):
        import pandas as pd
        return pd.read_csv(f) if str(f).lower().endswith(".csv") \
            else __import__("pandas").read_excel(f)

try:
    from model.test import test_model
except ImportError:
    def test_model(df): return df

plt.rcParams.update({
    "figure.facecolor":  "#111827",
    "axes.facecolor":    "#1a2235",
    "text.color":        "#e2e8f0",
    "axes.labelcolor":   "#94a3b8",
    "xtick.color":       "#94a3b8",
    "ytick.color":       "#94a3b8",
    "axes.edgecolor":    "#334155",
    "grid.color":        "#1e3a5f",
    "axes.grid":         True,
    "grid.linewidth":    0.5,
    "grid.alpha":        0.5,
    "legend.facecolor":  "#161d2e",
    "legend.edgecolor":  "#334155",
    "legend.fontsize":   9,
    "figure.autolayout": False,
    "font.family":       "monospace",
})

BG_DEEP = "#080c14"
BG = "#0d1117"
BG_MID = "#111827"
SIDEBAR = "#0a0f1a"
CARD = "#161d2e"
CARD2 = "#1a2438"
BORDER = "#1e3a5f"
ACCENT = "#3b82f6"
ACCENT_LT = "#60a5fa"
ACCENT2 = "#8b5cf6"
FG = "#f1f5f9"
FG_MID = "#cbd5e1"
FG_DIM = "#64748b"
SUCCESS = "#10b981"
SUCCESS2 = "#34d399"
WARN = "#f59e0b"
ERR = "#ef4444"

EXP_ALL_BG = "#065f46"
EXP_ALL_FG = "#34d399"
EXP_ANOM_BG = "#7f1d1d"
EXP_ANOM_FG = "#f87171"
EXP_RPT_BG = "#78350f"
EXP_RPT_FG = "#fbbf24"
EXP_PNG_BG = "#1e3a8a"
EXP_PNG_FG = "#93c5fd"

ANOMALY_COLOR = "#ef4444"
NORMAL_COLOR = "#3b82f6"

FONT_H2 = ("Georgia",     11, "bold")
FONT_H3 = ("Courier New", 10, "bold")
FONT_SM = ("Courier New",  9)
FONT_XS = ("Courier New",  8)

CLUSTER_PALETTE = [
    "#f59e0b", "#3b82f6", "#10b981", "#ec4899",
    "#8b5cf6", "#06b6d4", "#ef4444", "#84cc16",
    "#f97316", "#a78bfa",
]

raw_data = None
active_df = None
active_X = None
active_xcols = []
active_ycols = []
active_anomaly_result = None
active_figures = {}


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


def _save_dataframe(df, path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx":
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False)


class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _=None):
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.configure(bg=BORDER)
        tk.Label(tw, text=self.text, bg=CARD2, fg=FG_MID,
                 font=FONT_XS, padx=8, pady=4, relief="flat", bd=0).pack()

    def _hide(self, _=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


def make_scrollable(parent, bg=BG):
    outer = tk.Frame(parent, bg=bg)
    canvas = Canvas(outer, bg=bg, highlightthickness=0, bd=0)
    sb = Scrollbar(outer, orient="vertical", command=canvas.yview,
                   bg=BORDER, troughcolor=bg, activebackground=ACCENT)
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    canvas.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(canvas, bg=bg)
    wid = canvas.create_window((0, 0), window=inner, anchor="nw")

    def _resize(e):
        canvas.configure(scrollregion=canvas.bbox("all"))
        canvas.itemconfig(wid, width=e.width)
    inner.bind("<Configure>", lambda e: canvas.configure(
        scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", _resize)

    def _wheel(e):
        if e.num == 4:
            canvas.yview_scroll(-1, "units")
        elif e.num == 5:
            canvas.yview_scroll(1, "units")
        else:
            canvas.yview_scroll(int(-1*(e.delta/120)), "units")
    canvas.bind_all("<MouseWheel>", _wheel)
    canvas.bind_all("<Button-4>",   _wheel)
    canvas.bind_all("<Button-5>",   _wheel)
    return outer, inner


def make_card(parent, title=None, accent_color=None, **pack_kw):
    f = tk.LabelFrame(parent,
                      text=f"  {title}  " if title else "",
                      fg=accent_color or FG_DIM,
                      bg=CARD, font=FONT_H3, relief="flat", bd=0,
                      highlightbackground=accent_color or BORDER,
                      highlightthickness=1, padx=10, pady=8)
    f.pack(**{"fill": "x", "pady": 5, "padx": 6, **pack_kw})
    return f


def make_btn(parent, text, cmd, color=ACCENT, fg_col="#fff", tip=None):
    b = tk.Button(parent, text=text, command=cmd,
                  bg=color, fg=fg_col,
                  activebackground=_lighten(color), activeforeground=fg_col,
                  font=FONT_H3, relief="flat", bd=0,
                  cursor="hand2", padx=10, pady=8, anchor="w")
    b.pack(fill="x", pady=2)
    if tip:
        Tooltip(b, tip)
    b.bind("<Enter>", lambda e: b.config(bg=_lighten(color)))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b


def make_export_btn(parent, text, cmd, bg, fg, tip=None, state="normal"):
    b = tk.Button(parent, text=text, command=cmd,
                  bg=bg, fg=fg,
                  activebackground=_lighten(bg, 20), activeforeground=fg,
                  font=FONT_H3, relief="flat", bd=0,
                  cursor="hand2", padx=10, pady=9,
                  anchor="w", wraplength=320, justify="left", state=state)
    b.pack(fill="x", pady=3)
    if tip:
        Tooltip(b, tip)
    b.bind("<Enter>", lambda e: b.config(bg=_lighten(bg, 20)))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b


def _icon_btn(parent, text, cmd, bg, fg, tip=None):
    b = tk.Button(parent, text=text, command=cmd,
                  bg=bg, fg=fg,
                  activebackground=_lighten(bg, 20), activeforeground=fg,
                  font=FONT_XS, relief="flat", bd=0,
                  cursor="hand2", padx=8, pady=5)
    b.pack(side="right", padx=3)
    if tip:
        Tooltip(b, tip)
    b.bind("<Enter>", lambda e: b.config(bg=_lighten(bg, 20)))
    b.bind("<Leave>", lambda e: b.config(bg=bg))
    return b


_progress_running = False


def progress_start():
    global _progress_running
    _progress_running = True
    progress_bar.config(mode="indeterminate")
    progress_bar.start(12)


def progress_stop():
    global _progress_running
    _progress_running = False
    progress_bar.stop()
    progress_bar.config(mode="determinate")
    progress_bar["value"] = 100


def detect_hidden_patterns(X_scaled, contamination=0.05, method="iforest"):
    from sklearn.ensemble import IsolationForest
    from sklearn.neighbors import LocalOutlierFactor
    n = X_scaled.shape[0]
    if n < 10:
        labels = np.ones(n, dtype=int)
        return labels, 0.0, [], "N/A (too few samples)", np.zeros(n)
    safe_cont = float(np.clip(contamination, 0.001, 0.499))
    if method == "lof":
        det = LocalOutlierFactor(n_neighbors=min(
            20, n-1), contamination=safe_cont)
        labels = det.fit_predict(X_scaled)
        scores = det.negative_outlier_factor_
        name = "Local Outlier Factor (LOF)"
    else:
        det = IsolationForest(n_estimators=200, contamination=safe_cont,
                              random_state=42, n_jobs=-1)
        labels = det.fit_predict(X_scaled)
        scores = det.decision_function(X_scaled)
        name = "Isolation Forest"
    idx = list(np.where(labels == -1)[0])
    pct = len(idx) / n * 100
    return labels, pct, idx, name, scores


def _divider(parent):
    tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", padx=6, pady=(4, 3))


def _mini_bar(parent, label, frac, col):
    row = tk.Frame(parent, bg=CARD)
    row.pack(fill="x", padx=8, pady=2)
    tk.Label(row, text=label[:6], bg=CARD, fg=FG_DIM,
             font=FONT_XS, width=6).pack(side="left")
    outer = tk.Frame(row, bg=BORDER, height=7)
    outer.pack(side="left", fill="x", expand=True, padx=(3, 0))
    tk.Frame(outer, bg=col, height=7).place(
        relwidth=max(frac, 0.04), relheight=1.0)


def build_legend(parent, df, hex_colors):
    for w in parent.winfo_children():
        w.destroy()
    tk.Label(parent, text="CLUSTER LEGEND", bg=CARD, fg=ACCENT_LT,
             font=FONT_H3, justify="center").pack(pady=(12, 4), padx=8)
    _divider(parent)
    clusters = sorted(df["cluster"].unique())
    total = len(df)
    for cl in clusters:
        cnt = int((df["cluster"] == cl).sum())
        pct = cnt / total * 100
        col = hex_colors[int(cl) % len(hex_colors)]
        row = tk.Frame(parent, bg=CARD2, padx=6, pady=4,
                       highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", padx=8, pady=3)
        tk.Canvas(row, bg=col, width=22, height=22,
                  highlightthickness=0).pack(side="left", padx=(0, 8))
        txt = tk.Frame(row, bg=CARD2)
        txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=f"Cluster {cl}", bg=CARD2, fg=col,
                 font=FONT_H3, anchor="w").pack(anchor="w")
        tk.Label(txt, text=f"{cnt:,} wells", bg=CARD2, fg=FG_MID,
                 font=FONT_XS, anchor="w").pack(anchor="w")
        tk.Label(txt, text=f"{pct:.1f}%", bg=CARD2, fg=FG_DIM,
                 font=FONT_XS, anchor="w").pack(anchor="w")
    _divider(parent)
    tk.Label(parent, text="PROPORTION", bg=CARD, fg=FG_DIM,
             font=FONT_XS).pack(anchor="w", padx=8, pady=(2, 0))
    for cl in clusters:
        cnt = int((df["cluster"] == cl).sum())
        _mini_bar(parent, f"Cl.{cl}", cnt/total,
                  hex_colors[int(cl) % len(hex_colors)])


def build_anomaly_legend(parent, n_normal, n_anomaly):
    for w in parent.winfo_children():
        w.destroy()
    tk.Label(parent, text="PATTERN LEGEND", bg=CARD, fg=ANOMALY_COLOR,
             font=FONT_H3, justify="center").pack(pady=(12, 4), padx=8)
    _divider(parent)
    total = n_normal + n_anomaly
    for label, count, col, icon in [
        ("Normal Wells",  n_normal,  NORMAL_COLOR,  "●"),
        ("Anomaly Wells", n_anomaly, ANOMALY_COLOR, "◆"),
    ]:
        pct = count / total * 100 if total > 0 else 0
        row = tk.Frame(parent, bg=CARD2, padx=6, pady=4,
                       highlightbackground=BORDER, highlightthickness=1)
        row.pack(fill="x", padx=8, pady=3)
        sw = tk.Canvas(row, bg=col, width=22, height=22, highlightthickness=0)
        sw.pack(side="left", padx=(0, 8))
        sw.create_text(11, 11, text=icon, fill="white",
                       font=("Arial", 10, "bold"))
        txt = tk.Frame(row, bg=CARD2)
        txt.pack(side="left", fill="x", expand=True)
        tk.Label(txt, text=label,   bg=CARD2, fg=col,
                 font=FONT_H3, anchor="w").pack(anchor="w")
        tk.Label(txt, text=f"{count:,} wells", bg=CARD2, fg=FG_MID,
                 font=FONT_XS, anchor="w").pack(anchor="w")
        tk.Label(txt, text=f"{pct:.1f}%", bg=CARD2, fg=FG_DIM,
                 font=FONT_XS, anchor="w").pack(anchor="w")
    _divider(parent)
    tk.Label(parent, text="PROPORTION", bg=CARD, fg=FG_DIM,
             font=FONT_XS).pack(anchor="w", padx=8, pady=(2, 0))
    if total:
        _mini_bar(parent, "Norm", n_normal/total, NORMAL_COLOR)
        _mini_bar(parent, "Anom", n_anomaly/total, ANOMALY_COLOR)


def _fill_toolbar_bar(bar, mpl_canvas, tab_name):
    """
    Populate a pre-existing toolbar frame with:
      left  — label + matplotlib NavigationToolbar2Tk
      right — Save PNG button + CSV button
    The bar frame must already be packed before this is called.
    """
    tk.Label(bar, text=" 🔧 TOOLS:",
             bg="#0d1525", fg=FG_DIM, font=FONT_XS).pack(side="left", padx=(6, 2))

    nav_frame = tk.Frame(bar, bg="#0d1525")
    nav_frame.pack(side="left", padx=2)

    toolbar = NavigationToolbar2Tk(mpl_canvas, nav_frame)
    toolbar.config(bg="#0d1525")
    for child in toolbar.winfo_children():
        try:
            child.config(bg="#0d1525", fg=FG_MID,
                         activebackground=CARD2, activeforeground=FG,
                         relief="flat", bd=0, highlightthickness=0,
                         font=FONT_XS)
        except Exception:
            pass
    toolbar.update()

    tk.Frame(bar, bg=BORDER, width=1).pack(
        side="left", fill="y", padx=8, pady=3)

    def _export_png():
        fig = active_figures.get(tab_name)
        if fig is None:
            messagebox.showwarning("Export", "Run analysis first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            initialfile=f"cbm_{tab_name}_{ts()}.png",
            filetypes=[("PNG", "*.png"), ("SVG", "*.svg"), ("PDF", "*.pdf")])
        if not path:
            return
        fig.savefig(path, dpi=180, bbox_inches="tight",
                    facecolor=fig.get_facecolor())
        set_status(f"✔  Chart saved → {os.path.basename(path)}", SUCCESS)

    def _export_csv():
        if active_df is None:
            messagebox.showwarning("Export", "Run analysis first.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=f"cbm_data_{ts()}.csv",
            filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")])
        if not path:
            return
        _save_dataframe(active_df, path)
        set_status(f"✔  Data saved → {os.path.basename(path)}", SUCCESS)

    _icon_btn(bar, "💾 Save PNG", _export_png,
              bg=EXP_PNG_BG, fg=EXP_PNG_FG, tip="Save this chart as PNG/SVG/PDF")
    _icon_btn(bar, "📊 CSV Data", _export_csv,
              bg=EXP_ALL_BG, fg=EXP_ALL_FG, tip="Export full dataset as CSV")


def draw_plot(tab_frame, fig, tab_name, df=None, hex_colors=None,
              legend_type="cluster", anomaly_counts=None):
    """
    Layout (top→bottom inside tab_frame):
      1. toolbar_bar  (side=top)   ← created FIRST so it packs above chart
      2. row          (side=top, expand)
           left:  chart_frame  (expand)
           right: leg_frame    (fixed width)
    The matplotlib canvas is built inside chart_frame.
    The toolbar is then filled using _fill_toolbar_bar AFTER the canvas exists
    (NavigationToolbar2Tk needs a canvas) but the bar frame itself is already
    in place — so there is no re-parenting or 'before=' hack needed.
    """
    for w in tab_frame.winfo_children():
        w.destroy()
    active_figures[tab_name] = fig

    try:
        fig.tight_layout(pad=2.5, rect=[0.03, 0.03, 0.97, 0.95])
    except Exception:
        pass
    fig.patch.set_facecolor("#111827")

    # ── STEP 1: toolbar bar frame — pack first so it sits on top ──────────────
    toolbar_bar = tk.Frame(tab_frame, bg="#0d1525", pady=4,
                           highlightbackground=BORDER, highlightthickness=1)
    toolbar_bar.pack(side="top", fill="x")

    # ── STEP 2: content row (chart + legend) packed below toolbar ─────────────
    row = tk.Frame(tab_frame, bg=BG_MID)
    row.pack(side="top", fill="both", expand=True)

    chart_frame = tk.Frame(row, bg=BG_MID)
    chart_frame.pack(side="left", fill="both", expand=True)

    leg_frame = tk.Frame(row, bg=CARD, width=195,
                         highlightbackground=BORDER, highlightthickness=1)
    leg_frame.pack(side="right", fill="y", padx=(2, 6), pady=6)
    leg_frame.pack_propagate(False)

    # ── STEP 3: matplotlib canvas ──────────────────────────────────────────────
    mpl_canvas = FigureCanvasTkAgg(fig, master=chart_frame)
    mpl_canvas.draw()
    cw = mpl_canvas.get_tk_widget()
    cw.config(bg="#111827", highlightthickness=0)
    cw.pack(fill="both", expand=True, padx=2, pady=2)

    # ── STEP 4: fill toolbar bar (canvas now exists for NavigationToolbar) ─────
    _fill_toolbar_bar(toolbar_bar, mpl_canvas, tab_name)

    # ── STEP 5: legend panel ───────────────────────────────────────────────────
    if legend_type == "anomaly" and anomaly_counts is not None:
        build_anomaly_legend(leg_frame, *anomaly_counts)
    elif legend_type == "cluster" and df is not None and hex_colors:
        build_legend(leg_frame, df, hex_colors)


def _style_ax(ax):
    ax.set_facecolor("#1a2235")
    for sp in ax.spines.values():
        sp.set_edgecolor("#334155")
    ax.tick_params(colors="#94a3b8", labelsize=8)


def plot_clusters(df, xcols, ycols, hex_colors):
    fig, ax = plt.subplots(figsize=(7.5, 4.5))
    _style_ax(ax)
    x_col = xcols[0] if xcols else None
    y_col = (ycols[0] if ycols else
             (xcols[1] if len(xcols) > 1 else xcols[0] if xcols else None))
    if x_col and y_col and x_col in df.columns and y_col in df.columns:
        for cl in sorted(df["cluster"].unique()):
            mask = df["cluster"] == cl
            col = hex_colors[int(cl) % len(hex_colors)]
            ax.scatter(df.loc[mask, x_col], df.loc[mask, y_col],
                       color=col, s=45, edgecolors="#ffffff22", linewidths=0.5,
                       zorder=3, label=f"Cluster {cl}  ({int(mask.sum()):,} wells)")
        ax.set_xlabel(x_col, fontsize=9, labelpad=6)
        ax.set_ylabel(y_col, fontsize=9, labelpad=6)
        ax.margins(0.08)
    else:
        ax.text(0.2, 0.5, "Select X and Y features to plot",
                transform=ax.transAxes, color=FG_DIM, fontsize=10)
    ax.set_title("CBM Production Clusters",
                 color="#f1f5f9", fontsize=12, pad=12)
    ax.legend(loc="upper left", fontsize=8.5, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e",
              labelcolor="#e2e8f0", markerscale=1.4)
    return fig


def plot_pca(X, labels, hex_colors):
    from sklearn.decomposition import PCA
    fig, ax = plt.subplots(figsize=(7.5, 4.5))
    _style_ax(ax)
    if X.shape[0] < 2 or X.shape[1] < 1:
        ax.text(0.3, 0.5, "Not enough data for PCA",
                transform=ax.transAxes, color=FG_DIM)
        return fig
    n = min(2, X.shape[1])
    Z = PCA(n_components=n).fit_transform(X)
    if Z.shape[1] == 1:
        Z = np.hstack([Z, np.zeros_like(Z)])
    for cl in sorted(np.unique(labels)):
        mask = labels == cl
        col = hex_colors[int(cl) % len(hex_colors)]
        ax.scatter(Z[mask, 0], Z[mask, 1], color=col, s=45,
                   edgecolors="#ffffff22", linewidths=0.5, zorder=3,
                   label=f"Cluster {cl}  ({int(mask.sum()):,} wells)")
    ax.set_xlabel("Principal Component 1", fontsize=9, labelpad=6)
    ax.set_ylabel("Principal Component 2", fontsize=9, labelpad=6)
    ax.set_title("PCA — Well Feature Space",
                 color="#f1f5f9", fontsize=12, pad=12)
    ax.margins(0.08)
    ax.legend(loc="upper left", fontsize=8.5, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e",
              labelcolor="#e2e8f0", markerscale=1.4)
    return fig


def plot_reservoir_3d(df, xcols, ycols, hex_colors):
    from mpl_toolkits.mplot3d import Axes3D  # noqa
    fig = plt.figure(figsize=(7.5, 4.5))
    ax = fig.add_subplot(111, projection="3d")
    ax.set_facecolor("#1a2235")
    ax.tick_params(colors="#94a3b8", labelsize=7)
    cols3 = list(dict.fromkeys(xcols + ycols))[:3]
    if len(cols3) < 3:
        extra = [c for c in df.select_dtypes(include="number").columns
                 if c not in cols3 and c != "cluster"]
        cols3 = (cols3 + extra)[:3]
    if len(cols3) == 3 and all(c in df.columns for c in cols3):
        handles = []
        for cl in sorted(df["cluster"].unique()):
            mask = df["cluster"] == cl
            col = hex_colors[int(cl) % len(hex_colors)]
            ax.scatter(df.loc[mask, cols3[0]], df.loc[mask, cols3[1]],
                       df.loc[mask, cols3[2]], color=col, s=28,
                       edgecolors="#ffffff15", linewidths=0.3)
            handles.append(mpatches.Patch(
                color=col, label=f"Cluster {cl}  ({int(mask.sum()):,})"))
        ax.set_xlabel(cols3[0], color="#94a3b8", fontsize=7, labelpad=3)
        ax.set_ylabel(cols3[1], color="#94a3b8", fontsize=7, labelpad=3)
        ax.set_zlabel(cols3[2], color="#94a3b8", fontsize=7, labelpad=3)
        ax.legend(handles=handles, loc="upper left", fontsize=8,
                  framealpha=0.85, facecolor="#161d2e",
                  edgecolor="#334155", labelcolor="#e2e8f0")
    else:
        ax.text2D(0.15, 0.5, "Need ≥ 3 numeric features",
                  transform=ax.transAxes, color=FG_DIM)
    ax.set_title("3D Reservoir Map", color="#f1f5f9", fontsize=12, pad=8)
    return fig


def plot_production(df, ycols, xcols, hex_colors):
    fig, ax = plt.subplots(figsize=(7.5, 4.5))
    _style_ax(ax)
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
        xs = sub[x_col].values if x_col else np.arange(len(sub))
        ax.plot(xs, sub[y_col].values, color=col, linewidth=2, alpha=0.88,
                label=f"Cluster {cl}  ({len(sub):,} wells)")
    ax.set_xlabel(x_col if x_col else "Well Index", fontsize=9, labelpad=6)
    ax.set_ylabel(y_col, fontsize=9, labelpad=6)
    ax.set_title("Production Curves by Cluster",
                 color="#f1f5f9", fontsize=12, pad=12)
    ax.margins(0.04, 0.10)
    ax.legend(loc="upper left", fontsize=8.5, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e", labelcolor="#e2e8f0")
    return fig


def plot_hidden_patterns(X_scaled, anomaly_labels, xcols, detector_name):
    from sklearn.decomposition import PCA
    fig, ax = plt.subplots(figsize=(7.5, 4.5))
    _style_ax(ax)
    n = X_scaled.shape[0]
    if X_scaled.shape[1] >= 2:
        Z = PCA(n_components=2, random_state=42).fit_transform(X_scaled)
        xlabel, ylabel = "PC 1  (feature projection)", "PC 2  (feature projection)"
    elif X_scaled.shape[1] == 1:
        Z = np.column_stack([X_scaled[:, 0], np.arange(n)])
        xlabel = xcols[0] if xcols else "Feature"
        ylabel = "Well Index"
    else:
        ax.text(0.3, 0.5, "No feature data available",
                transform=ax.transAxes, color=FG_DIM, fontsize=10)
        ax.set_title("Hidden Pattern Detection",
                     color="#f1f5f9", fontsize=12, pad=12)
        return fig
    nm = anomaly_labels == 1
    am = anomaly_labels == -1
    n_norm = int(nm.sum())
    n_anom = int(am.sum())
    ax.scatter(Z[nm, 0], Z[nm, 1], color=NORMAL_COLOR, s=30, alpha=0.70,
               edgecolors="#ffffff18", linewidths=0.3, zorder=3,
               label=f"● Normal wells  ({n_norm:,})")
    if n_anom > 0:
        ax.scatter(Z[am, 0], Z[am, 1], color=ANOMALY_COLOR, s=70, alpha=0.95,
                   edgecolors="#ffffff66", linewidths=0.9, marker="D", zorder=5,
                   label=f"◆ Anomalous wells  ({n_anom:,})")
        ax.annotate(
            f"◆ {n_anom} anomal{'y' if n_anom==1 else 'ies'} detected",
            xy=(Z[am, 0].mean(), Z[am, 1].mean()),
            xytext=(14, 14), textcoords="offset points",
            color=ANOMALY_COLOR, fontsize=9, fontweight="bold",
            arrowprops=dict(arrowstyle="->", color=ANOMALY_COLOR, lw=1.0))
    ax.legend(loc="upper right", fontsize=9, framealpha=0.92,
              edgecolor="#334155", facecolor="#161d2e",
              labelcolor="#e2e8f0", markerscale=1.3)
    ax.set_xlabel(xlabel, fontsize=9, labelpad=6)
    ax.set_ylabel(ylabel, fontsize=9, labelpad=6)
    ax.set_title(f"Hidden Pattern Detection  ·  {detector_name}",
                 color="#f1f5f9", fontsize=12, pad=12)
    ax.margins(0.10)
    return fig


def get_selected(lb):
    return [lb.get(i).strip() for i in lb.curselection()]


def resolve_features():
    if raw_data is None:
        return [], [], None, 0, "No dataset loaded"
    all_num = list(raw_data.select_dtypes(include="number").columns)
    x_sel = [c for c in get_selected(x_listbox) if c in raw_data.columns]
    y_sel = [c for c in get_selected(y_listbox) if c in raw_data.columns]
    warns = []
    x_cols = x_sel or all_num
    if not x_sel:
        warns.append("No X selected → all numeric cols used")
    y_cols = y_sel
    if not y_sel:
        warns.append("No Y selected → cluster colour used")
    needed = list(dict.fromkeys(x_cols + y_cols))
    wdf = raw_data.dropna(subset=needed).reset_index(drop=True)
    return x_cols, y_cols, wdf, len(wdf), " | ".join(warns)


def upload_dataset():
    global raw_data
    f = filedialog.askopenfilename(
        filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx *.xls"), ("All", "*.*")])
    if not f:
        return
    try:
        raw_data = load_dataset(f)
        rows_var.set(f"{len(raw_data):,}")
        cols_var.set(str(len(raw_data.columns)))
        x_listbox.delete(0, "end")
        y_listbox.delete(0, "end")
        for col in raw_data.columns:
            x_listbox.insert("end", f"  {col}")
            y_listbox.insert("end", f"  {col}")
        preview_table(raw_data)
        set_status(
            f"✔  Loaded: {os.path.basename(f)}  ({len(raw_data):,} rows)", SUCCESS)
        for k in stat_widgets:
            stat_widgets[k].set("—")
        stat_widgets["Total Wells"].set(f"{len(raw_data):,}")
        export_all_btn.config(state="normal")
    except Exception as e:
        messagebox.showerror("Load Error", str(e))
        set_status(f"✖  Load failed: {e}", ERR)


def preview_table(df):
    for w in table_frame.winfo_children():
        w.destroy()
    s = ttk.Style()
    s.configure("P.Treeview", background=CARD, foreground=FG,
                fieldbackground=CARD, rowheight=21, font=FONT_SM)
    s.configure("P.Treeview.Heading", background=BORDER,
                foreground=FG_MID, font=FONT_H3)
    s.map("P.Treeview", background=[("selected", ACCENT)])
    tv = ttk.Treeview(table_frame, style="P.Treeview")
    vsb = ttk.Scrollbar(table_frame, orient="vertical",   command=tv.yview)
    hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tv.xview)
    tv.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tv["columns"] = list(df.columns)
    tv["show"] = "headings"
    for col in df.columns:
        tv.heading(col, text=col)
        tv.column(col, width=90, minwidth=60)
    for row in df.head(20).values:
        tv.insert("", "end", values=list(row))
    hsb.pack(side="bottom", fill="x")
    vsb.pack(side="right",  fill="y")
    tv.pack(fill="both", expand=True)


def start_pipeline():
    set_status("⟳  Running analysis…", WARN)
    progress_start()
    threading.Thread(target=run_pipeline, daemon=True).start()


def run_pipeline():
    global active_df, active_X, active_xcols, active_ycols, active_anomaly_result
    x_cols, y_cols, wdf, well_count, warn = resolve_features()
    if wdf is None or well_count == 0:
        app.after(0, lambda: set_status(
            f"⚠  {warn or 'No usable rows'}", WARN))
        app.after(0, progress_stop)
        return
    if warn:
        app.after(0, lambda: set_status(f"ℹ  {warn}", WARN))
    try:
        from sklearn.cluster import KMeans
        n_clusters = min(cluster_var.get(), well_count)
        Xraw = wdf[x_cols].values.astype(float)
        mu, s = Xraw.mean(axis=0), Xraw.std(axis=0)
        s[s == 0] = 1
        X_scaled = (Xraw - mu) / s
        wdf = wdf.copy()
        km = KMeans(n_clusters=n_clusters, random_state=42, n_init="auto")
        wdf["cluster"] = km.fit_predict(X_scaled)
        wdf = test_model(wdf)
        method = anomaly_method_var.get()
        contam = anomaly_contam_var.get() / 100.0
        a_labels, a_pct, a_idx, a_name, a_scores = \
            detect_hidden_patterns(
                X_scaled, contamination=contam, method=method)
        wdf["anomaly"] = a_labels
        anomaly_result = {
            "labels":        a_labels,   "pct":      a_pct,
            "indices":       a_idx,      "scores":   a_scores,
            "detector_name": a_name,
            "n_anomaly":     len(a_idx),
            "n_normal":      len(a_labels) - len(a_idx),
        }
        active_df = wdf
        active_X = X_scaled
        active_xcols = x_cols
        active_ycols = y_cols
        active_anomaly_result = anomaly_result
        hx = cluster_hex(n_clusters)
        insights = generate_insights(wdf, x_cols, y_cols,
                                     well_count, hx, anomaly_result)
        app.after(0, lambda: refresh_ui(wdf, X_scaled, x_cols, y_cols,
                                        hx, insights, well_count, anomaly_result))
        app.after(0, lambda: set_status(
            "✔  Analysis complete — results ready", SUCCESS))
        app.after(0, progress_stop)
        app.after(0, lambda: export_anomaly_btn.config(state="normal"))
        app.after(0, lambda: export_report_btn.config(state="normal"))
    except Exception as e:
        traceback.print_exc()
        app.after(0, lambda: set_status(f"✖  {e}", ERR))
        app.after(0, progress_stop)


def refresh_ui(df, X, x_cols, y_cols, hx, insights, well_count, anomaly_result):
    draw_plot(cluster_tab,
              plot_clusters(df, x_cols, y_cols, hx), "clusters", df, hx)
    draw_plot(pca_tab,
              plot_pca(X, df["cluster"].values, hx), "pca", df, hx)
    draw_plot(reservoir_tab,
              plot_reservoir_3d(df, x_cols, y_cols, hx), "reservoir_3d", df, hx)
    draw_plot(production_tab,
              plot_production(df, y_cols, x_cols, hx), "production", df, hx)
    draw_plot(hidden_tab,
              plot_hidden_patterns(X, anomaly_result["labels"],
                                   x_cols, anomaly_result["detector_name"]),
              "hidden_patterns", legend_type="anomaly",
              anomaly_counts=(anomaly_result["n_normal"],
                              anomaly_result["n_anomaly"]))
    stat_widgets["Total Wells"].set(f"{well_count:,}")
    stat_widgets["Analysed"].set(f"{len(df):,}")
    stat_widgets["Clusters"].set(str(df["cluster"].nunique()))
    stat_widgets["Anomalies"].set(
        f"{anomaly_result['n_anomaly']}  ({anomaly_result['pct']:.1f}%)")
    explain_text.config(state="normal")
    explain_text.delete("1.0", "end")
    explain_text.insert("end", insights)
    explain_text.config(state="disabled")


def export_all_data():
    if active_df is None:
        messagebox.showwarning("Export", "Run analysis first.")
        return
    path = filedialog.asksaveasfilename(
        defaultextension=".csv", initialfile=f"cbm_all_wells_{ts()}.csv",
        filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")])
    if not path:
        return
    _save_dataframe(active_df, path)
    set_status(f"✔  Full dataset exported → {os.path.basename(path)}", SUCCESS)


def export_anomaly_wells():
    if active_df is None or "anomaly" not in active_df.columns:
        messagebox.showwarning("Export", "Run analysis first.")
        return
    anom_df = active_df[active_df["anomaly"] == -1].copy()
    if len(anom_df) == 0:
        messagebox.showinfo("Export", "No anomalous wells detected.")
        return
    path = filedialog.asksaveasfilename(
        defaultextension=".csv", initialfile=f"cbm_anomaly_wells_{ts()}.csv",
        filetypes=[("CSV", "*.csv"), ("Excel", "*.xlsx")])
    if not path:
        return
    _save_dataframe(anom_df, path)
    set_status(
        f"✔  {len(anom_df)} anomaly wells exported → {os.path.basename(path)}", SUCCESS)


def export_text_report():
    txt = explain_text.get("1.0", "end").strip()
    if not txt:
        messagebox.showwarning("Export", "Run analysis first.")
        return
    path = filedialog.asksaveasfilename(
        defaultextension=".txt", initialfile=f"cbm_report_{ts()}.txt",
        filetypes=[("Text", "*.txt"), ("All", "*.*")])
    if not path:
        return
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("CBM AI Analytics Platform — Report\n")
        fh.write(
            f"Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        fh.write("="*50+"\n\n")
        fh.write(txt)
    set_status(f"✔  Report saved → {os.path.basename(path)}", SUCCESS)


def export_all_charts():
    if not active_figures:
        messagebox.showwarning("Export", "Run analysis first.")
        return
    folder = filedialog.askdirectory(title="Choose folder for chart images")
    if not folder:
        return
    saved = 0
    for name, fig in active_figures.items():
        fpath = os.path.join(folder, f"cbm_{name}_{ts()}.png")
        try:
            fig.savefig(fpath, dpi=180, bbox_inches="tight",
                        facecolor=fig.get_facecolor())
            saved += 1
        except Exception as ex:
            print(f"Could not save {name}: {ex}")
    set_status(
        f"✔  {saved} charts saved to {os.path.basename(folder)}/", SUCCESS)


def generate_insights(df, x_cols, y_cols, well_count, hex_colors, anomaly_result):
    clusters = df["cluster"].value_counts().sort_index()
    n_anom = anomaly_result["n_anomaly"]
    a_pct = anomaly_result["pct"]
    det_name = anomaly_result["detector_name"]
    if a_pct < 2:
        sev = "LOW"
        interp = "Very few outlier wells. Reservoir is relatively uniform."
    elif a_pct < 8:
        sev = "MODERATE"
        interp = "Moderate anomalies. Check localised heterogeneity or faults."
    elif a_pct < 20:
        sev = "ELEVATED"
        interp = "Significant complexity. Variable seam thickness or fractures."
    else:
        sev = "HIGH"
        interp = "Large fraction flagged. Review contamination % or data quality."
    lines = [
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        "  CBM AI ANALYSIS REPORT",
        f"  {datetime.datetime.now().strftime('%Y-%m-%d  %H:%M')}",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        f"  Total wells : {well_count:,}",
        f"  Used wells  : {len(df):,}",
        f"  Clusters    : {len(clusters)}",
        f"  X features  : {', '.join(x_cols) or 'all numeric'}",
        f"  Y features  : {', '.join(y_cols) or 'none'}",
        "", "  CLUSTER BREAKDOWN",
        "  ──────────────────────────────────────",
    ]
    for cl, cnt in clusters.items():
        pct = cnt / len(df) * 100
        col = hex_colors[int(cl) % len(hex_colors)]
        bar = "█" * int(pct / 3)
        lines.append(
            f"  ■ Cluster {cl}  [{col}]  {bar:<25} {cnt:>5,}  {pct:5.1f}%")
    lines += [
        "", "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        "  HIDDEN PATTERN DETECTION",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
        f"  Algorithm    : {det_name}",
        f"  Normal wells : {anomaly_result['n_normal']:,}",
        f"  Anomalies    : {n_anom:,}",
        f"  Anomaly rate : {a_pct:.1f}%  [{sev} severity]",
        "", f"  INTERPRETATION: {interp}",
        "", "  Possible causes:",
        "  • Reservoir heterogeneity (variable permeability / coal rank)",
        "  • Abnormal formation pressure or aquifer encroachment",
        "  • Equipment malfunction (pump failures, casing leaks)",
        "  • Operational issues (shut-ins, choke changes, workover)",
        "  • Natural fracture network interference (faults, cleats)",
        "", "  CHART LEGEND GUIDE",
        "  ──────────────────────────────────────",
        "  In-chart legend (upper-left): colour swatch + cluster + count",
        "  Right panel: large colour squares + counts + proportion bars",
        "  Clusters/PCA/Production → coloured by cluster",
        "  Hidden Patterns → blue●=normal, red◆=anomaly",
        "  Toolbar above each chart: Home | Pan | Zoom | Back | Fwd | Save",
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    ]
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ══════════════════════════════════════════════════════════════════════════════
app = ttk.Window(themename="darkly")
app.title("CBM AI Analytics Platform  v2.1")
app.geometry("1600x900")
app.configure(bg=BG_DEEP)
app.minsize(1200, 700)

top_bar = tk.Frame(app, bg="#050810", height=48)
top_bar.pack(fill="x", side="top")
top_bar.pack_propagate(False)
brand = tk.Frame(top_bar, bg="#050810")
brand.pack(side="left", padx=16, fill="y")
tk.Label(brand, text="CBM·AI", bg="#050810", fg=ACCENT,
         font=("Georgia", 16, "bold")).pack(side="left", padx=(0, 10))
tk.Label(brand, text="Coalbed Methane Analytics Platform",
         bg="#050810", fg=FG_DIM, font=FONT_SM).pack(side="left")
tk.Label(top_bar, text="v2.1  ·  KMeans + Isolation Forest / LOF",
         bg="#050810", fg=FG_DIM, font=FONT_XS).pack(side="right", padx=14)
tk.Frame(app, bg=BORDER, height=1).pack(fill="x")

root_pane = tk.PanedWindow(app, orient="horizontal", bg=BG_DEEP,
                           sashwidth=4, sashrelief="flat",
                           sashpad=0, handlesize=0)
root_pane.pack(fill="both", expand=True)

# LEFT SIDEBAR
sb_outer, sidebar = make_scrollable(root_pane, bg=SIDEBAR)
root_pane.add(sb_outer, width=380, minsize=340)
tk.Frame(sidebar, bg=SIDEBAR, height=10).pack()

ds_card = make_card(sidebar, "DATASET", accent_color=SUCCESS2)
make_btn(ds_card, "⬆  Upload Dataset (CSV / Excel)", upload_dataset,
         color=SUCCESS, fg_col="#000",
         tip="Load a CSV or Excel file to begin analysis")
info_row = tk.Frame(ds_card, bg=CARD)
info_row.pack(fill="x", pady=(6, 0))
rows_var = tk.StringVar(value="—")
cols_var = tk.StringVar(value="—")
for title, var, col in [("Total Wells", rows_var, ACCENT_LT), ("Columns", cols_var, FG_DIM)]:
    cf = tk.Frame(info_row, bg=CARD)
    cf.pack(side="left", expand=True, fill="x")
    tk.Label(cf, text=title, bg=CARD, fg=FG_DIM, font=FONT_XS).pack(anchor="w")
    tk.Label(cf, textvariable=var, bg=CARD, fg=col,
             font=("Courier New", 14, "bold")).pack(anchor="w")

cl_card = make_card(sidebar, "CLUSTER SETTINGS", accent_color=ACCENT)
tk.Label(cl_card, text="Number of Clusters  (2 – 10)",
         bg=CARD, fg=FG_DIM, font=FONT_SM).pack(anchor="w")
cluster_var = tk.IntVar(value=3)
sf = tk.Frame(cl_card, bg=CARD)
sf.pack(fill="x", pady=4)
ttk.Spinbox(sf, from_=2, to=10, textvariable=cluster_var,
            width=5, font=FONT_H2).pack(side="left")
tk.Label(sf, text="clusters", bg=CARD, fg=FG_DIM,
         font=FONT_SM).pack(side="left", padx=8)

anom_card = make_card(sidebar, "ANOMALY DETECTION", accent_color=ERR)
tk.Label(anom_card, text="Algorithm", bg=CARD,
         fg=FG_DIM, font=FONT_SM).pack(anchor="w")
anomaly_method_var = tk.StringVar(value="iforest")
mrow = tk.Frame(anom_card, bg=CARD)
mrow.pack(fill="x", pady=(2, 6))
for label, val in [("Isolation Forest  (faster, tree-based)", "iforest"),
                   ("Local Outlier Factor  (density-based)", "lof")]:
    tk.Radiobutton(mrow, text=label, variable=anomaly_method_var, value=val,
                   bg=CARD, fg=FG_MID, selectcolor=CARD2,
                   activebackground=CARD, activeforeground=ACCENT,
                   font=FONT_SM, wraplength=320).pack(anchor="w", pady=1)
tk.Label(anom_card, text="Contamination  (expected anomaly %)",
         bg=CARD, fg=FG_DIM, font=FONT_SM).pack(anchor="w")
anomaly_contam_var = tk.DoubleVar(value=5.0)
crow = tk.Frame(anom_card, bg=CARD)
crow.pack(fill="x", pady=2)
ttk.Spinbox(crow, from_=1.0, to=49.0, increment=0.5,
            textvariable=anomaly_contam_var, width=6, font=FONT_H3).pack(side="left")
tk.Label(crow, text="%   (typical CBM: 3–10%)",
         bg=CARD, fg=FG_DIM, font=FONT_XS).pack(side="left", padx=6)

feat_card = make_card(sidebar, "FEATURE SELECTION", accent_color=ACCENT2)
tk.Label(feat_card, text="X — Input Features  (Ctrl+click for multi-select)",
         bg=CARD, fg=ACCENT_LT, font=FONT_SM, wraplength=340).pack(anchor="w")
x_listbox = Listbox(feat_card, height=4, selectmode="multiple",
                    bg="#0e1520", fg=FG, selectbackground=ACCENT,
                    selectforeground="#fff", font=FONT_SM,
                    relief="flat", bd=0, highlightthickness=1,
                    highlightcolor=ACCENT, highlightbackground=BORDER)
x_listbox.pack(fill="x", pady=(2, 7))
tk.Label(feat_card, text="Y — Target / Overlay Features",
         bg=CARD, fg=ACCENT2, font=FONT_SM).pack(anchor="w")
y_listbox = Listbox(feat_card, height=4, selectmode="multiple",
                    bg="#0e1520", fg=FG, selectbackground=ACCENT2,
                    selectforeground="#fff", font=FONT_SM,
                    relief="flat", bd=0, highlightthickness=1,
                    highlightcolor=ACCENT2, highlightbackground=BORDER)
y_listbox.pack(fill="x", pady=2)
tk.Label(feat_card, text="ℹ  Leave blank → all numeric columns used as X features",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=340).pack(anchor="w", pady=(3, 0))

run_card = make_card(sidebar)
make_btn(run_card, "▶   Run AI Analysis", start_pipeline,
         color=ACCENT, tip="Run KMeans clustering + anomaly detection")
progress_bar = ttk.Progressbar(run_card, mode="determinate",
                               bootstyle="info-striped", length=340)
progress_bar.pack(fill="x", pady=(4, 2))
status_var = tk.StringVar(value="Ready — upload a dataset to begin")
status_lbl = tk.Label(run_card, textvariable=status_var,
                      bg=CARD, fg=FG_DIM, font=FONT_SM,
                      anchor="w", wraplength=340, justify="left")
status_lbl.pack(fill="x", pady=(2, 0))

exp_card = make_card(sidebar, "EXPORT", accent_color=WARN)
tk.Label(exp_card, text="Each button exports different data:",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=340,
         justify="left").pack(anchor="w", pady=(0, 4))
export_all_btn = make_export_btn(
    exp_card, "📊  Export All Wells (CSV)\n     Full dataset with cluster & anomaly columns",
    export_all_data, bg=EXP_ALL_BG, fg=EXP_ALL_FG,
    tip="Green — all wells", state="disabled")
export_anomaly_btn = make_export_btn(
    exp_card, "⚠  Export Anomaly Wells (CSV)\n     Only flagged anomalous wells",
    export_anomaly_wells, bg=EXP_ANOM_BG, fg=EXP_ANOM_FG,
    tip="Red — anomaly wells only", state="disabled")
export_report_btn = make_export_btn(
    exp_card, "📄  Export Insights Report (TXT)\n     Full AI analysis text summary",
    export_text_report, bg=EXP_RPT_BG, fg=EXP_RPT_FG,
    tip="Amber — text report", state="disabled")
export_charts_btn = make_export_btn(
    exp_card, "🖼  Export All Charts (PNG)\n     Save all 5 chart images to a folder",
    export_all_charts, bg=EXP_PNG_BG, fg=EXP_PNG_FG,
    tip="Blue — all chart PNGs")

prev_card = make_card(sidebar, "DATASET PREVIEW  (first 20 rows)",
                      fill="both", expand=True)
prev_card.config(pady=6)
table_frame = tk.Frame(prev_card, bg=CARD)
table_frame.pack(fill="both", expand=True)
tk.Frame(sidebar, bg=SIDEBAR, height=16).pack()

# RIGHT PANEL
rp_outer, right_inner = make_scrollable(root_pane, bg=BG)
root_pane.add(rp_outer, minsize=760)

rp_hdr = tk.Frame(right_inner, bg=BG, padx=16, pady=10)
rp_hdr.pack(fill="x")
tk.Label(rp_hdr, text="Analysis Dashboard", bg=BG, fg=FG,
         font=("Georgia", 14, "bold")).pack(side="left")
tk.Label(rp_hdr, text="Coalbed Methane  ·  AI-Powered  ·  Distinct cluster colours",
         bg=BG, fg=FG_DIM, font=FONT_SM).pack(side="right", pady=2)
tk.Frame(right_inner, bg=BORDER, height=1).pack(fill="x", padx=12)

nb_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=6)
nb_wrap.pack(fill="x")
sty = ttk.Style()
sty.configure("CBM.TNotebook",     background=BG, tabmargins=[0, 0, 0, 0])
sty.configure("CBM.TNotebook.Tab", background=CARD, foreground=FG_DIM,
              padding=[14, 6], font=FONT_H3)
sty.map("CBM.TNotebook.Tab",
        background=[("selected", ACCENT)],
        foreground=[("selected", "#ffffff")])
notebook = ttk.Notebook(nb_wrap, style="CBM.TNotebook")
notebook.pack(fill="x")

TAB_H = 480
cluster_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
pca_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
reservoir_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
production_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
hidden_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
for tab, name in [
    (cluster_tab,    "  📍 Clusters  "),
    (pca_tab,        "  📈 PCA  "),
    (reservoir_tab,  "  🗺 3D Reservoir  "),
    (production_tab, "  ⚡ Production  "),
    (hidden_tab,     "  ◆ Hidden Patterns  "),
]:
    tab.pack_propagate(False)
    notebook.add(tab, text=name)

stats_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=5)
stats_wrap.pack(fill="x")
stat_defs = [("Total Wells", "—", ACCENT_LT), ("Analysed", "—", SUCCESS2),
             ("Clusters", "—", ACCENT2), ("Anomalies", "—", ERR)]
stat_widgets = {}
for title, val, col in stat_defs:
    sf = tk.Frame(stats_wrap, bg=CARD, padx=12, pady=9,
                  highlightbackground=BORDER, highlightthickness=1)
    sf.pack(side="left", expand=True, fill="x", padx=3)
    sv = tk.StringVar(value=val)
    tk.Label(sf, textvariable=sv, bg=CARD, fg=col,
             font=("Courier New", 14, "bold")).pack()
    tk.Label(sf, text=title, bg=CARD, fg=FG_DIM, font=FONT_XS).pack()
    stat_widgets[title] = sv

ins_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=6)
ins_wrap.pack(fill="x")
ins_hdr = tk.Frame(ins_wrap, bg=BG)
ins_hdr.pack(fill="x", pady=(0, 4))
tk.Label(ins_hdr, text="AI Insights Report", bg=BG, fg=FG,
         font=FONT_H2).pack(side="left")
copy_btn = tk.Button(ins_hdr, text="📋 Copy Report",
                     command=lambda: (app.clipboard_clear(),
                                      app.clipboard_append(explain_text.get("1.0", "end"))),
                     bg=CARD2, fg=FG_DIM, activebackground=BORDER,
                     activeforeground=FG, font=FONT_XS,
                     relief="flat", bd=0, cursor="hand2", padx=10, pady=4)
copy_btn.pack(side="right")
Tooltip(copy_btn, "Copy entire report to clipboard")

explain_text = tk.Text(ins_wrap, height=18, bg=CARD, fg=SUCCESS2,
                       font=("Courier New", 9), relief="flat", bd=0,
                       padx=12, pady=10, insertbackground=FG, wrap="word",
                       highlightbackground=BORDER, highlightthickness=1,
                       state="disabled")
explain_text.pack(fill="x")
tk.Frame(right_inner, bg=BG, height=20).pack()

explain_text.config(state="normal")
explain_text.insert("end",
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                    "  CBM AI Analytics Platform  v2.1\n"
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
                    "  FIXES IN v2.1:\n"
                    "  ✔  TclError fixed — toolbar frame created first,\n"
                    "     no re-parenting or 'before=' needed\n"
                    "  ✔  Toolbar above each chart (Home/Pan/Zoom/Save)\n"
                    "  ✔  In-chart legend: colour + cluster + well count\n"
                    "  ✔  Export: green/red/amber/blue distinct buttons\n"
                    "  ✔  Sidebar wider — no text cut off\n"
                    "  ✔  Distinct cluster colours (amber/blue/green/pink…)\n\n"
                    "  TO BEGIN:\n"
                    "  1. Click ⬆ Upload Dataset\n"
                    "  2. Optionally select X / Y features\n"
                    "  3. Set cluster count and anomaly %\n"
                    "  4. Click ▶ Run AI Analysis\n"
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                    )
explain_text.config(state="disabled")

app.mainloop()
