for key in _ev_label_names:
    row = tk.Frame(ev_detail_frame, bg=CARD)
    row.pack(fill="x", pady=1)
    lbl = tk.Label(row, text="—", bg=CARD, fg=FG_DIM,
                   font=FONT_XS, width=26, anchor="w")
    lbl.pack(side="left", padx=(4, 0))
    cnt_var = tk.StringVar(value="—")
    tk.Label(row, textvariable=cnt_var, bg=CARD, fg=COL_EVENTS,
             font=FONT_XS, width=8, anchor="e").pack(side="right", padx=4)
    ev_count_vars[key] = cnt_var
    ev_count_lbls[key] = lbl

# ── ⑧ EXPORT ──────────────────────────────────────────────────────────────────
exp_body = make_section_card(sidebar, "⑧ EXPORT RESULTS", COL_EXPORT)
tk.Label(exp_body,
         text="Single Excel file — 5 sheets:\n"
              "  All Wells  ·  Anomaly Wells  ·  Cluster Summary\n"
              "  Parameter Weights  ·  Insights Report",
         bg=CARD, fg=FG_DIM, font=FONT_XS, wraplength=390, justify="left").pack(anchor="w", pady=(0, 6))
export_btn = tk.Button(exp_body, text="  Export All Results  (Excel / CSV)",
                       command=export_results, bg="#14532d", fg=COL_EXPORT,
                       activebackground=_lighten("#14532d", 20), activeforeground=COL_EXPORT,
                       font=FONT_H3, relief="flat", bd=0, cursor="hand2",
                       padx=10, pady=11, anchor="w", state="disabled")
export_btn.pack(fill="x", pady=2)
export_btn.bind("<Enter>", lambda e: export_btn.config(
    bg=_lighten("#14532d", 20)))
export_btn.bind("<Leave>", lambda e: export_btn.config(bg="#14532d"))
Tooltip(export_btn, "Export all results to multi-sheet Excel")

# ── ⑨ DATASET PREVIEW ──────────────────────────────────────────────────────────
prev_body = make_section_card(sidebar, "⑨ DATASET PREVIEW  (first 20 rows)", COL_PREVIEW,
                              fill="both", expand=True)
table_frame = tk.Frame(prev_body, bg=CARD)
table_frame.pack(fill="both", expand=True)
tk.Frame(sidebar, bg=SIDEBAR, height=20).pack()

# ══════════════════════════════════════════════════════════════════════════════
# RIGHT PANEL
# ══════════════════════════════════════════════════════════════════════════════
rp_outer, right_inner = make_scrollable(root_pane, bg=BG)
root_pane.add(rp_outer, minsize=820)

rp_hdr = tk.Frame(right_inner, bg=BG, padx=16, pady=10)
rp_hdr.pack(fill="x")
tk.Label(rp_hdr, text="Analysis Dashboard", bg=BG, fg=FG,
         font=("Georgia", 15, "bold")).pack(side="left")
tk.Label(rp_hdr, text="Dynamic parameter weights  ·  Any file format  ·  Event analysis",
         bg=BG, fg=FG_DIM, font=FONT_SM).pack(side="right", pady=2)
tk.Frame(right_inner, bg=BORDER, height=1).pack(fill="x", padx=12)

# Stats
stats_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=5)
stats_wrap.pack(fill="x")
stat_defs_r1 = [("Total Wells", "—", COL_DATASET), ("Analysed", "—", COL_CLUSTER),
                ("Clusters", "—", COL_WEIGHTS),   ("Anomalies", "—", COL_ANOMALY)]
stat_defs_r2 = [("Active Wells", "—", COL_EVENTS), ("Inactive Wells", "—", ERR),
                ("Event Types", "—", WARN),         ("Abnormal", "—", COL_ANOMALY)]
stat_widgets = {}
for stat_row_defs in [stat_defs_r1, stat_defs_r2]:
    row = tk.Frame(stats_wrap, bg=BG)
    row.pack(fill="x", pady=2)
    for title, val, col in stat_row_defs:
        sf = tk.Frame(row, bg=CARD, padx=10, pady=7,
                      highlightbackground=col, highlightthickness=1)
        sf.pack(side="left", expand=True, fill="x", padx=3)
        sv = tk.StringVar(value=val)
        tk.Label(sf, textvariable=sv, bg=CARD, fg=col,
                 font=("Courier New", 12, "bold")).pack()
        tk.Label(sf, text=title, bg=CARD, fg=FG_DIM, font=FONT_XS).pack()
        stat_widgets[title] = sv

# Notebook
nb_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=6)
nb_wrap.pack(fill="x")
sty = tk_ttk.Style()
sty.configure("CBM.TNotebook",     background=BG, tabmargins=[0, 0, 0, 0])
sty.configure("CBM.TNotebook.Tab", background=CARD2, foreground=FG_DIM,
              padding=[12, 6], font=FONT_H3)
sty.map("CBM.TNotebook.Tab",
        background=[("selected", "#1e3a8a")], foreground=[("selected", "#93c5fd")])
notebook = tk_ttk.Notebook(nb_wrap, style="CBM.TNotebook")
notebook.pack(fill="x")

TAB_H = 500
cluster_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
pca_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
reservoir_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
production_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
hidden_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
weight_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)
event_tab = tk.Frame(notebook, bg=BG_MID, height=TAB_H)

for tab, name in [
    (cluster_tab,    "  Clusters  "),
    (pca_tab,        "  PCA  "),
    (reservoir_tab,  "  3D Reservoir  "),
    (production_tab, "  Production  "),
    (hidden_tab,     "  Hidden Patterns  "),
    (weight_tab,     "  Param Weights  "),
    (event_tab,      "  Events  "),
]:
    tab.pack_propagate(False)
    notebook.add(tab, text=name)

# Insights
ins_wrap = tk.Frame(right_inner, bg=BG, padx=10, pady=6)
ins_wrap.pack(fill="x")
ins_hdr = tk.Frame(ins_wrap, bg=BG)
ins_hdr.pack(fill="x", pady=(0, 4))
tk.Label(ins_hdr, text="AI Insights Report", bg=BG,
         fg=FG, font=FONT_H2).pack(side="left")
copy_btn = tk.Button(ins_hdr, text="Copy",
                     command=lambda: (app.clipboard_clear(),
                                      app.clipboard_append(explain_text.get("1.0", "end"))),
                     bg=CARD2, fg=FG_DIM, activebackground=BORDER, activeforeground=FG,
                     font=FONT_XS, relief="flat", bd=0, cursor="hand2", padx=10, pady=4)
copy_btn.pack(side="right")

explain_text = tk.Text(ins_wrap, height=22, bg=CARD, fg=SUCCESS2,
                       font=("Courier New", 9), relief="flat", bd=0,
                       padx=12, pady=10, insertbackground=FG, wrap="word",
                       highlightbackground=BORDER, highlightthickness=1, state="disabled")
explain_text.pack(fill="x")
tk.Frame(right_inner, bg=BG, height=20).pack()

# Welcome text
explain_text.config(state="normal")
explain_text.insert("end",
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                    "  CBM AI Analytics Platform  v5.0\n"
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
                    "  WHAT'S NEW IN v5.0:\n\n"
                    "  DYNAMIC PARAMETER WEIGHTS:\n"
                    "  Upload any dataset and the numeric columns\n"
                    "  appear automatically in Section ③.\n"
                    "  Set a % weight for each column you care about.\n"
                    "  Total should sum to 100% (live indicator shown).\n"
                    "  Weights shape how wells are grouped by KMeans.\n\n"
                    "  ANY FILE FORMAT:\n"
                    "  The uploader accepts CSV, Excel (all variants),\n"
                    "  ODS, JSON, Parquet, Feather, HDF5, Pickle.\n"
                    "  Unknown extensions are tried automatically.\n"
                    "  No crash — a clear error message is shown.\n\n"
                    "  QUICK START:\n"
                    "  ①  Upload your data file\n"
                    "  ②  Set number of clusters\n"
                    "  ③  Adjust parameter weights for your columns\n"
                    "  ④  Set anomaly detection settings\n"
                    "  ⑤  Optionally select X-axis features\n"
                    "  ⑥  Click Run AI Analysis\n"
                    "  ⑦  View event summary panel\n"
                    "  ⑧  Export results to Excel\n"
                    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
                    )
explain_text.config(state="disabled")

app.mainloop()
