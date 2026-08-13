#!/usr/bin/env python3
"""
Fire Suppression Quote Wizard – Python GUI Edition
Ported from the original Excel VBA macros.
Requires: xlwings (pip install xlwings) + a local Excel install for export.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import math
import os
import itertools
import traceback
from typing import List, Dict, Optional

# ── Optional dependency ───────────────────────
# Uses xlwings (drives real Excel via COM) instead of openpyxl, so exporting
# never rewrites the template file structure -- Excel Tables, conditional
# formatting extensions, and print areas in the template are left untouched.
try:
    import xlwings as xw
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# ── Engineering calculations ──────────────────
def ceil(x: float) -> int:
    return math.ceil(x)

def even(x: int) -> int:
    return x if x % 2 == 0 else x + 1

def lookup_table(fill_per_cyl: float, thresholds: List[tuple]):
    for limit, val in thresholds:
        if fill_per_cyl <= limit:
            return val
    return None

def calc_fm200(agent_qty: float, height: float) -> dict:
    T, U = 0.5621, max(height or 3, 0.5)
    K = ceil(agent_qty / 162)
    L = ceil(agent_qty / K)
    E = K * L
    M = lookup_table(L, [(8,8),(16,16),(32,32),(52,52),(95,106),(132,147),(162,180)])
    N = lookup_table(L, [(8,"303.205.015"),(16,"303.205.016"),(32,"303.205.017"),
                         (52,"303.205.018"),(95,"303.205.019"),(132,"303.205.020"),
                         (162,"303.205.021")])
    O = 314205322 if N in ("303.205.015","303.205.016","303.205.018") else 314205306
    P = 311700014 if N in ("303.205.015","303.205.016","303.205.017") else 311700006
    Q = ceil(K / 11)
    R = K - Q
    S = K - Q
    a_base = ceil(agent_qty / (T * U * 95))
    a = max(a_base * ceil(U / 4.87), 1)
    V = a if (agent_qty / a) <= 100 else a * 2
    W = max(even(ceil(agent_qty / (T * U * 25))), 2)
    X, Y = W // 2, W // 2
    return dict(K=K, L=L, E=E, M=M, N=N, O=O, P=P, Q=Q, R=R, S=S, V=V, W=W, X=X, Y=Y)

def calc_novec(agent_qty: float, height: float) -> dict:
    T, U = 0.5621, max(height or 3, 0.5)
    K = ceil(agent_qty / 187)
    L = ceil(agent_qty / K)
    E = K * L
    M = lookup_table(L, [(9.5,8),(19,16),(38,32),(62,52),(114,106),(158,147),(187,180)])
    N = lookup_table(L, [(9.5,"303.207.001"),(19,"303.207.002"),(38,"303.207.003"),
                         (62,"303.207.004"),(114,"303.207.005"),(158,"303.207.006"),
                         (187,"303.207.007")])
    O = 314207208 if N in ("303.207.001","303.207.002","303.207.004") else 314207321
    P = 311700014 if N in ("303.207.001","303.207.002","303.207.004") else 311700006
    Q = ceil(K / 11)
    R = K - Q
    S = K - Q
    a_base = ceil(agent_qty / (T * U * 96))
    a = max(a_base * ceil(U / 4.3), 1)
    V = a if (agent_qty / a) <= 136 else a * 2
    W = max(even(ceil(agent_qty / (T * U * 25))), 2)
    X, Y = W // 2, W // 2
    return dict(K=K, L=L, E=E, M=M, N=N, O=O, P=P, Q=Q, R=R, S=S, V=V, W=W, X=X, Y=Y)

# ── Package builders ──────────────────────────
def detection_full(D: str, C: int, X: int, Y: int) -> List[dict]:
    return [
        dict(D=D, F="4004-9302", H=1*C, K="", L="FA", M="Control Panel"),
        dict(D=D, F="batt",       H=2,     K="ch", L="FA", M="Local"),
        dict(D=D, F="Pro",        H=1,     K="ch", L="FA", M="Service"),
        dict(D=D, F="4098-5610",  H=X*C,  K="", L="FA", M="Field Device"),
        dict(D=D, F="4098-5261",  H=1,     K="ch", L="FA", M="Field Device"),
        dict(D=D, F="4098-5601",  H=Y*C,  K="", L="FA", M="Field Device"),
        dict(D=D, F="4098-5261",  H=1,     K="ch", L="FA", M="Field Device"),
        dict(D=D, F="2099-9149",  H=1*C, K="", L="FA", M="Field Device"),
        dict(D=D, F="2080-9067",  H=1*C, K="", L="FA", M="Field Device"),
        dict(D=D, F="4906-9127",  H=1*C, K="", L="FA", M="Notification"),
        dict(D=D, F="INT-GB6-24", H=1*C, K="", L="FA", M="Notification"),
    ]

def detection_co2(D: str, C: int=1) -> List[dict]:
    items = [
        dict(D=D, F="4004-9302", H=1, K="", L="FA", M="Control Panel"),
        dict(D=D, F="batt",       H=2, K="ch", L="FA", M="Local"),
        dict(D=D, F="pro",        H=1, K="ch", L="FA", M="Service"),
        dict(D=D, F="4098-5610",  H=1, K="", L="FA", M="Field Device"),
        dict(D=D, F="4098-5261",  H=1, K="ch", L="FA", M="Field Device"),
        dict(D=D, F="4098-5601",  H=1, K="", L="FA", M="Field Device"),
        dict(D=D, F="4098-5261",  H=1, K="ch", L="FA", M="Field Device"),
        dict(D=D, F="2099-9149",  H=1, K="", L="FA", M="Field Device"),
        dict(D=D, F="4906-9127",  H=1, K="", L="FA", M="Notification"),
        dict(D=D, F="INT-GB6-24", H=1, K="", L="FA", M="Notification"),
    ]
    if C > 1:
        for item in items:
            item["H"] *= C
    return items

def build_fm200_package(room: str, agent_qty: float, repeat: int, height: float) -> Optional[List[dict]]:
    c = calc_fm200(agent_qty, height)
    if c['M'] is None or c['N'] is None:
        return None
    D = "Tyco - Hygood (FM)"
    items = [
        dict(type="header2", text=room),
        dict(type="intro", text="FM-200 Systems Package includes below items :"),
    ]
    if c['K'] == 1:
        items += [
            dict(type="item", D=D, F=c['E'], H=1*repeat, K="", L="FF", M="Subtotal"),
            dict(type="item", D=D, F="300.205.001", H=c['E'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F=c['N'], H=1, K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="304205030", H=1, K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="304.205.006", H=1, K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="302.205.025", H=1, K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F=c['O'], H=1, K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F=c['P'], H=1, K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="310.205.224", H=c['V']*repeat, K="", L="FF", M="FM200"),
        ]
    else:
        items += [
            dict(type="item", D=D, F=c['E'], H=1*repeat, K="", L="FF", M="Subtotal"),
            dict(type="item", D=D, F="300.205.001", H=c['E'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F=c['N'], H=c['K'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="304205030", H=c['Q'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="304.205.006", H=c['K'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="302.205.025", H=c['K'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="304.209.004", H=c['R'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="306.205.003", H=c['S'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F=c['O'], H=c['K'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F=c['P'], H=c['K'], K="ch", L="FF", M="FM200"),
            dict(type="item", D=D, F="310.205.224", H=c['V']*repeat, K="", L="FF", M="FM200"),
        ]
    items.append(dict(type="intro", text="Detection Package includes below items:"))
    for it in detection_full("Simplex", repeat, c['X'], c['Y']):
        items.append(dict(type="item", **it))
    return items

def build_novec_package(room: str, agent_qty: float, repeat: int, height: float) -> Optional[List[dict]]:
    c = calc_novec(agent_qty, height)
    if c['M'] is None or c['N'] is None:
        return None
    D = "Tyco - Hygood (SP)"
    items = [
        dict(type="header2", text=room),
        dict(type="intro", text="Novec Systems Package includes below items :"),
    ]
    if c['K'] == 1:
        items += [
            dict(type="item", D=D, F=c['E'], H=1*repeat, K="", L="FF", M="Subtotal"),
            dict(type="item", D=D, F="452011", H=c['E'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F=c['N'], H=1, K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="304205030", H=1, K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="304.205.006", H=1, K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="302207021", H=1, K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F=c['O'], H=1, K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F=c['P'], H=1, K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="310.207.220", H=c['V']*repeat, K="", L="FF", M="Novec"),
        ]
    else:
        items += [
            dict(type="item", D=D, F=c['E'], H=1*repeat, K="", L="FF", M="Subtotal"),
            dict(type="item", D=D, F="452011", H=c['E'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F=c['N'], H=c['K'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="304205030", H=c['Q'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="304.205.006", H=c['K'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="302207021", H=c['K'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="304.209.004", H=c['R'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="306.205.003", H=c['S'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F=c['O'], H=c['K'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F=c['P'], H=c['K'], K="ch", L="FF", M="Novec"),
            dict(type="item", D=D, F="310.207.220", H=c['V']*repeat, K="", L="FF", M="Novec"),
        ]
    items.append(dict(type="intro", text="Detection Package includes below items:"))
    for it in detection_full("Simplex", repeat, c['X'], c['Y']):
        items.append(dict(type="item", **it))
    return items

def build_co2_ansul(room: str, cylinder_count: int=1) -> List[dict]:
    D = "Ansul"
    items = [
        dict(type="header2", text=room),
        dict(type="intro", text="CO2 Systems Package includes below items :"),
    ]
    for _ in range(cylinder_count):
        items += [
            dict(type="section", text="MASTER"),
            dict(type="item", D=D, F="451500", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="73327", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="428949", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="427082", H=1, K="ch", L="FF", M="CO2"),
            dict(type="section", text="SLAVE"),
            dict(type="item", D=D, F="451500", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="427082", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="426112", H=1, K="", L="FF", M="CO2"),
        ]
    items.append(dict(type="intro", text="Detection Package includes below items:"))
    for it in detection_co2("Simplex", cylinder_count):
        items.append(dict(type="item", **it))
    return items

def build_co2_hygood(room: str, cylinder_count: int=1) -> List[dict]:
    D = "Tyco - Hygood (LPG)"
    items = [
        dict(type="header2", text=room),
        dict(type="intro", text="CO2 Systems Package includes below items :"),
    ]
    for _ in range(cylinder_count):
        items += [
            dict(type="section", text="MASTER"),
            dict(type="item", D=D, F="71104006h", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="30521041", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="91104001", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="20006020", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="30027301", H=1, K="ch", L="FF", M="CO2"),
            dict(type="section", text="SLAVE"),
            dict(type="item", D=D, F="71104007h", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="30521041", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="30100102", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="30500041", H=1, K="ch", L="FF", M="CO2"),
            dict(type="item", D=D, F="30460002", H=1, K="", L="FF", M="CO2"),
        ]
    items.append(dict(type="intro", text="Detection Package includes below items:"))
    for it in detection_co2("Simplex", cylinder_count):
        items.append(dict(type="item", **it))
    return items

SYSTEM_HEADER1 = {
    "fm200": "Tyco / Hygood FM-200 Systems - UL Listed & FM Approved",
    "novec": "Hygood FK-5-1-12 Systems - UL Listed & FM Approved",
    "co2ansul": "Ansul CO2 - UL listed & FM Approved",
    "co2hygood": "Hygood (LPG) CO2 Systems - VdS Approved",
}

# First writable data row on the "Offer" sheet -- rows above this are the
# template's fixed header/title block. Matches the CCTV export tool's layout.
OFFER_DATA_START_ROW = 9

# Number of rows to check below a candidate "last part number" row before
# trusting it. Normal header gaps in this template are at most 2 empty
# cells, so this stays comfortably above that with margin.
F_COLUMN_GAP_CHECK = 5

# How far below OFFER_DATA_START_ROW to bulk-read column F when hunting for
# the last real part number. Generous on purpose -- it's a single COM call
# either way, so there's no real cost to scanning a wide range.
F_COLUMN_SCAN_ROWS = 5000

# ── Main Application ──────────────────────────
class QuoteWizardApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Fire Suppression Quote Wizard")
        self.root.geometry("1200x720")
        self.root.minsize(1000, 600)

        # Force a theme that Tk draws entirely itself (no native OS theming, no
        # extra files). The default 'vista' theme on Windows relies on native
        # Common-Controls-v6 theming, which needs an application manifest that
        # PyInstaller doesn't add by default -- in frozen builds this can leave
        # every ttk widget (Notebook, Treeview, Buttons, etc.) failing to draw
        # at all, even though plain tk widgets still render fine.
        style = ttk.Style(self.root)
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass

        # We'll build the UI step by step inside a try block
        self.quote: List[dict] = []
        self._uid_counter = itertools.count(1)
        self.template_path = None
        self.template_filename = None

        self.ws_toggle = tk.BooleanVar(value=True)
        self.ca_toggle = tk.BooleanVar(value=True)
        self.tab_widgets = {}
        self.calc_label = None
        self._status_label = None

        # Show a loading message immediately
        self.loading_label = ttk.Label(self.root, text="Loading interface, please wait...",
                                       font=("", 12))
        self.loading_label.pack(pady=20)
        self.root.update()

        # Now build everything
        try:
            self._create_layout()
            self._update_display()
            # Remove loading label
            self.loading_label.destroy()
            if not EXCEL_AVAILABLE:
                self._status_label.config(
                    text="⚠ xlwings not installed – Excel export disabled. Run: pip install xlwings",
                    foreground="red"
                )
            else:
                self._status_label.config(text="✔ xlwings available – Excel export enabled.", foreground="green")

            # Force an explicit repaint. Frozen (PyInstaller) builds on Windows sometimes
            # never receive an initial paint/expose event, leaving the window blank/white
            # until the user manually resizes it. Nudging the geometry forces Tk to redraw.
            self.root.update_idletasks()
            self.root.geometry(self.root.geometry())
        except Exception:
            # Show the full traceback in a messagebox
            messagebox.showerror("Startup Error",
                                 f"An error occurred while building the interface:\n\n{traceback.format_exc()}")
            raise

    def _create_layout(self):
        # Remove the loading label (it will be replaced later)
        self.loading_label.pack_forget()

        # Container that actually owns both panels, so grid geometry management
        # is unambiguous (each panel's real Tk parent is the widget laying it out).
        container = ttk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1, minsize=380)
        container.columnconfigure(1, weight=2)
        container.rowconfigure(0, weight=1)

        left_frame = ttk.Frame(container, padding=10)
        left_frame.grid(row=0, column=0, sticky="nsew")
        self.left_frame = left_frame

        right_frame = ttk.Frame(container, padding=10)
        right_frame.grid(row=0, column=1, sticky="nsew")
        self.right_frame = right_frame

        self.notebook = ttk.Notebook(left_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # FM‑200 tab
        fm200_tab = ttk.Frame(self.notebook)
        self._build_agent_tab(fm200_tab, "FM-200", "fm200")
        self.notebook.add(fm200_tab, text="FM-200")

        # Novec tab
        novec_tab = ttk.Frame(self.notebook)
        self._build_agent_tab(novec_tab, "Novec (FK-5-1-12)", "novec")
        self.notebook.add(novec_tab, text="Novec")

        # CO2 Ansul tab
        co2ansul_tab = ttk.Frame(self.notebook)
        self._build_co2_tab(co2ansul_tab, "Ansul CO2", "co2ansul")
        self.notebook.add(co2ansul_tab, text="CO2 – Ansul")

        # CO2 Hygood tab
        co2hygood_tab = ttk.Frame(self.notebook)
        self._build_co2_tab(co2hygood_tab, "Hygood (LPG) CO2", "co2hygood")
        self.notebook.add(co2hygood_tab, text="CO2 – Hygood")

        # Add button
        self.add_btn = ttk.Button(left_frame, text="+ Add Package to Quote", command=self._on_add)
        self.add_btn.pack(fill=tk.X, pady=(10, 0))

        # Calculation preview (removed wraplength to avoid layout issues)
        self.calc_label = tk.Label(left_frame, text="", anchor=tk.W, justify=tk.LEFT,
                                   bg="#f8f5ef", fg="#4a453d", font=("Courier", 10),
                                   padx=5, pady=5, relief=tk.GROOVE)
        self.calc_label.pack(fill=tk.BOTH, pady=(10, 0))

        # Toolbar
        toolbar = ttk.Frame(right_frame)
        toolbar.pack(fill=tk.X, pady=(0, 10))

        export_state = tk.NORMAL if EXCEL_AVAILABLE else tk.DISABLED
        self.export_btn = ttk.Button(toolbar, text="⬇ Export to Excel (Offer tab)",
                                     command=self._on_export, state=export_state)
        self.export_btn.pack(side=tk.LEFT, padx=2)

        self.continue_btn = ttk.Button(toolbar, text="⤵ Continue Offer",
                                       command=self._on_continue_offer, state=export_state)
        self.continue_btn.pack(side=tk.LEFT, padx=2)

        self.load_btn = ttk.Button(toolbar, text="Use my template file…",
                                   command=self._on_load_template, state=export_state)
        self.load_btn.pack(side=tk.LEFT, padx=2)

        self.clear_btn = ttk.Button(toolbar, text="Clear Quote", command=self._on_clear)
        self.clear_btn.pack(side=tk.LEFT, padx=2)

        toggle_frame = ttk.Frame(toolbar)
        toggle_frame.pack(side=tk.RIGHT, padx=10)
        ttk.Checkbutton(toggle_frame, text="Warning Sign", variable=self.ws_toggle,
                        command=self._update_display).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(toggle_frame, text="Class A Module", variable=self.ca_toggle,
                        command=self._update_display).pack(side=tk.LEFT, padx=5)

        self._status_label = ttk.Label(right_frame, text="", font=("", 9))
        self._status_label.pack(anchor=tk.W, pady=(0, 5))

        # Treeview
        tree_frame = ttk.Frame(right_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("mfr", "part", "qty", "unit", "type", "category")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                                 selectmode="browse")
        for col in columns:
            self.tree.heading(col, text=col.capitalize())
        self.tree.column("mfr", width=120)
        self.tree.column("part", width=120)
        self.tree.column("qty", width=60)
        self.tree.column("unit", width=60)
        self.tree.column("type", width=60)
        self.tree.column("category", width=150)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.bind("<Button-3>", self._on_tree_right_click)
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Delete Selected Row", command=self._delete_selected_row)

        ttk.Label(right_frame, text="Right‑click a row to delete it.",
                  foreground="gray", font=("", 9)).pack(anchor=tk.W, pady=(5, 0))

    def _build_agent_tab(self, parent: ttk.Frame, label: str, sys_id: str):
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Room / System Name").pack(anchor=tk.W)
        room_entry = ttk.Entry(frame)
        room_entry.pack(fill=tk.X, pady=(0, 5))

        grid = ttk.Frame(frame)
        grid.pack(fill=tk.X, pady=5)

        ttk.Label(grid, text="Agent Quantity (kg)").grid(row=0, column=0, sticky=tk.W, padx=(0,5))
        qty_entry = ttk.Entry(grid, width=15)
        qty_entry.grid(row=1, column=0, sticky=tk.W, padx=(0,5))

        ttk.Label(grid, text="Repeated Times").grid(row=0, column=1, sticky=tk.W)
        repeat_entry = ttk.Entry(grid, width=10)
        repeat_entry.insert(0, "1")
        repeat_entry.grid(row=1, column=1, sticky=tk.W)

        ttk.Label(frame, text="Room Height (m)").pack(anchor=tk.W, pady=(5,0))
        height_entry = ttk.Entry(frame)
        height_entry.insert(0, "3")
        height_entry.pack(fill=tk.X)

        ttk.Label(frame, text=f"{label} sizing is automatic.",
                  foreground="gray", font=("", 9)).pack(anchor=tk.W, pady=(2,0))

        self.tab_widgets[sys_id] = {
            'room': room_entry, 'qty': qty_entry,
            'repeat': repeat_entry, 'height': height_entry
        }
        for w in [qty_entry, height_entry, repeat_entry]:
            w.bind("<KeyRelease>", lambda e, sid=sys_id: self._update_preview(sid))

    def _build_co2_tab(self, parent: ttk.Frame, label: str, sys_id: str):
        frame = ttk.Frame(parent, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Room / Package Description").pack(anchor=tk.W)
        room_entry = ttk.Entry(frame)
        room_entry.pack(fill=tk.X, pady=(0,5))

        ttk.Label(frame, text="Number of Cylinders (Master+Slave pairs)").pack(anchor=tk.W)
        cyl_entry = ttk.Entry(frame)
        cyl_entry.insert(0, "1")
        cyl_entry.pack(fill=tk.X)

        ttk.Label(frame, text=f"{label} package repeats Master/Slave BOM per cylinder.",
                  foreground="gray", font=("", 9)).pack(anchor=tk.W, pady=(2,0))

        self.tab_widgets[sys_id] = {'room': room_entry, 'cyl': cyl_entry}

    # ── Preview update ─────────────────────────
    def _update_preview(self, sys_id: str):
        if sys_id not in ("fm200", "novec"):
            self.calc_label.config(text="")
            return
        w = self.tab_widgets[sys_id]
        try:
            qty = float(w['qty'].get())
            if qty <= 0:
                raise ValueError
        except (ValueError, KeyError):
            self.calc_label.config(text="Enter an agent quantity to preview the sizing.")
            return

        try:
            height = float(w['height'].get()) or 3.0
            if height <= 0:
                height = 3.0
        except:
            height = 3.0

        calc_fn = calc_fm200 if sys_id == "fm200" else calc_novec
        c = calc_fn(qty, height)

        if c['M'] is None or c['N'] is None:
            self.calc_label.config(text="Cylinder fill exceeds maximum capacity.\nReduce agent quantity.")
            return

        txt = (f"Cylinder Qty: {c['K']}   Fill/Cyl: {c['L']} kg   Vol: {c['M']} L\n"
               f"P/N: {c['N']}   Label: {c['O']}   Bracket: {c['P']}\n"
               f"Elec Act: {c['Q']}   Pneum Act: {c['R']}   Hoses: {c['S']}\n"
               f"Nozzles: {c['V']}   Detectors: {c['W']} ({c['X']}H / {c['Y']}S)")
        self.calc_label.config(text=txt)

    # ── Add package ────────────────────────────
    def _on_add(self):
        tab_idx = self.notebook.select()
        tab_text = self.notebook.tab(tab_idx, "text")
        sys_map = {
            "FM-200": "fm200", "Novec": "novec",
            "CO2 – Ansul": "co2ansul", "CO2 – Hygood": "co2hygood"
        }
        sys_id = sys_map.get(tab_text)
        if not sys_id:
            return

        w = self.tab_widgets[sys_id]
        room = w['room'].get().strip()
        if not room:
            messagebox.showwarning("Input Error", "Please enter a room / package name.")
            return

        try:
            if sys_id in ("fm200", "novec"):
                qty = float(w['qty'].get())
                if qty <= 0:
                    raise ValueError
                repeat = int(w['repeat'].get()) or 1
                height = float(w['height'].get()) or 3.0
                if height <= 0:
                    height = 3.0

                builder = build_fm200_package if sys_id == "fm200" else build_novec_package
                items = builder(room, qty, repeat, height)
                if items is None:
                    messagebox.showerror("Error", "Cylinder fill exceeds maximum capacity.\nReduce agent quantity.")
                    return

                w['room'].delete(0, tk.END)
                w['qty'].delete(0, tk.END)
                w['repeat'].delete(0, tk.END); w['repeat'].insert(0, "1")
                w['height'].delete(0, tk.END); w['height'].insert(0, "3")

            else:
                cyl = int(w['cyl'].get()) or 1
                if cyl < 1:
                    raise ValueError
                builder = build_co2_ansul if sys_id == "co2ansul" else build_co2_hygood
                items = builder(room, cyl)
                w['room'].delete(0, tk.END)
                w['cyl'].delete(0, tk.END); w['cyl'].insert(0, "1")

        except ValueError:
            messagebox.showwarning("Input Error", "Please enter valid numerical values.")
            return

        if not any(r.get("type") == "header1" and r.get("system") == sys_id for r in self.quote):
            self.quote.append(dict(type="header1", system=sys_id, text=SYSTEM_HEADER1[sys_id],
                                   uid=next(self._uid_counter)))

        for item in items:
            item["system"] = sys_id
            item["uid"] = next(self._uid_counter)
        self.quote.extend(items)
        self._update_display()

    def _on_clear(self):
        if self.quote and not messagebox.askyesno("Clear Quote", "Remove all items?"):
            return
        self.quote.clear()
        self._update_display()

    # ── Toggles & display ─────────────────────
    def _apply_toggles(self) -> List[dict]:
        rows = []
        ws = self.ws_toggle.get()
        ca = self.ca_toggle.get()
        for r in self.quote:
            rows.append(r)
            if r.get("type") == "item":
                if ws and r.get("F") == "INT-GB6-24":
                    rows.append(dict(type="item", system=r["system"], D="Simplex",
                                     F="4906-9101", H=1, K="", L="FF", M="Local",
                                     added=True))
                if ca and r.get("F") in ("Pro", "pro"):
                    rows.append(dict(type="item", system=r["system"], D="Simplex",
                                     F="4004-9864", H=1, K="", L="FA", M="Control Panel",
                                     added=True))
        return rows

    def _update_display(self, *args):
        self.tree.delete(*self.tree.get_children())
        rows = self._apply_toggles()

        self.tree.tag_configure("header1", background="#1b2430",
                                foreground="white", font=("", 10, "bold"))
        self.tree.tag_configure("header2", background="#e7e0d2",
                                font=("", 10, "bold"))
        self.tree.tag_configure("intro", font=("", 10, "bold"))
        self.tree.tag_configure("section", foreground="gray",
                                font=("", 9, "italic"))
        self.tree.tag_configure("item", font=("", 10))

        for idx, row in enumerate(rows):
            if row.get("type") in ("header1", "header2", "intro", "section"):
                values = (row["text"], "", "", "", "", "")
            else:
                values = (row.get("D", ""), row.get("F", ""),
                          row.get("H", ""), row.get("K", ""),
                          row.get("L", ""), row.get("M", ""))
            tags = [row.get("type", "item")]
            if not row.get("added") and "uid" in row:
                tags.append(f"orig_{row['uid']}")
            self.tree.insert("", tk.END, iid=str(idx), values=values, tags=tuple(tags))

    def _on_tree_right_click(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def _delete_selected_row(self):
        sel = self.tree.selection()
        if not sel:
            return
        item_iid = sel[0]
        for tag in self.tree.item(item_iid, "tags"):
            if tag.startswith("orig_"):
                try:
                    orig_uid = int(tag.split("_")[1])
                except ValueError:
                    continue
                for i, r in enumerate(self.quote):
                    if r.get("uid") == orig_uid:
                        del self.quote[i]
                        self._update_display()
                        return
        messagebox.showinfo("Cannot delete",
                            "This row was added by a toggle. Uncheck the toggle to remove it.")

    def _write_offer_rows(self, ws, rows, start_row: int) -> int:
        """Write the quote rows into the Offer sheet starting at start_row.
        Returns the row number one past the last row written."""
        r = start_row
        for row in rows:
            rtype = row.get("type")
            if rtype == "header1":
                ws.range(f"G{r}").value = row["text"]
                ws.range(f"D{r}:M{r}").color = (27, 36, 48)
                ws.range(f"G{r}").font.color = (255, 255, 255)
                ws.range(f"G{r}").font.bold = True
                ws.range(f"G{r}").font.size = 13
                ws.range(f"{r}:{r}").row_height = 22
            elif rtype == "header2":
                ws.range(f"G{r}").value = row["text"]
                ws.range(f"G{r}").color = (231, 224, 210)
                ws.range(f"G{r}").font.bold = True
                ws.range(f"G{r}").font.size = 11
                ws.range(f"{r}:{r}").row_height = 18
            elif rtype == "intro":
                ws.range(f"G{r}").value = row["text"]
                ws.range(f"G{r}").font.bold = True
            elif rtype == "section":
                ws.range(f"F{r}").value = row["text"]
                ws.range(f"F{r}").font.italic = True
            else:
                if row.get("D", "") != "":
                    ws.range(f"D{r}").value = row.get("D", "")
                if row.get("F", "") != "":
                    ws.range(f"F{r}").value = row.get("F", "")
                if row.get("H", "") != "":
                    ws.range(f"H{r}").value = row.get("H", "")
                if row.get("K", "") != "":
                    ws.range(f"K{r}").value = row.get("K", "")
                if row.get("L", "") != "":
                    ws.range(f"L{r}").value = row.get("L", "")
                if row.get("M", "") != "":
                    ws.range(f"M{r}").value = row.get("M", "")
            r += 1
        return r

    def _find_offer_sheet(self, wb):
        for name in wb.sheet_names:
            if name.lower() == "offer":
                return wb.sheets[name]
        raise Exception('Sheet "Offer" was not found in this file.')

    def _resolve_template_path(self, dialog_title: str):
        path = self.template_path
        if not path:
            path = filedialog.askopenfilename(
                title=dialog_title,
                filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        return path

    # ── Excel export ───────────────────────────
    def _on_load_template(self):
        if not EXCEL_AVAILABLE:
            messagebox.showinfo("Missing Module", "Excel export requires xlwings.\nInstall with: pip install xlwings")
            return
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        if not path:
            return
        self.template_path = path
        self.template_filename = os.path.basename(path)
        self._status_label.config(
            text=f"Template loaded: {self.template_filename} – export will append to its 'Offer' tab.",
            foreground="blue"
        )

    def _export_common(self, dialog_title: str, get_start_row):
        """Shared flow for Export and Continue Offer: resolve the file, open
        it via xlwings, locate the Offer sheet, compute the start row with
        the given strategy, write the current quote, and save."""
        if not self.quote:
            messagebox.showinfo("Empty Quote", "Add at least one package first.")
            return
        if not EXCEL_AVAILABLE:
            messagebox.showinfo("Missing Module", "Excel export requires xlwings.\nInstall with: pip install xlwings")
            return

        template_path = self._resolve_template_path(dialog_title)
        if not template_path:
            return

        save_as_new = messagebox.askyesno(
            "Save Option",
            "Yes = Save as a new file\nNo = Overwrite the file")
        output_path = template_path
        if save_as_new:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=self.template_filename or "Fire_Suppression_Offer.xlsx"
            )
            if not output_path:
                return

        rows = self._apply_toggles()

        app = None
        wb = None
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(template_path)
            ws = self._find_offer_sheet(wb)

            start_row = get_start_row(ws)
            self._write_offer_rows(ws, rows, start_row)

            wb.save(output_path)
            wb.close()
            app.quit()
            messagebox.showinfo("Done", f"Offer sheet updated starting at row {start_row}.\nSaved to {output_path}")

        except Exception as e:
            try:
                if wb is not None:
                    wb.close()
            except Exception:
                pass
            try:
                if app is not None:
                    app.quit()
            except Exception:
                pass
            messagebox.showerror("Export Error", str(e))

    def _on_export(self):
        self._export_common(
            "Select Excel Template",
            lambda ws: OFFER_DATA_START_ROW
        )

    def _on_continue_offer(self):
        self._export_common(
            "Select Existing Offer File",
            lambda ws: self._last_f_row(ws) + 1
        )

    def _last_f_row(self, ws) -> int:
        """Row of the last real part-number row in column F.

        Deliberately does NOT use Excel's End(xlUp)/Ctrl+Up: when a column
        belongs to an Excel Table (a ListObject -- e.g. this template's
        "TQuotation" table), End(xlUp) snaps to the Table's outer boundary
        rather than the true last populated cell, even across a long run of
        genuinely blank cells inside the Table body. That produced wrong
        results here (jumping to the bottom of a table that extends far
        past the real data).

        Instead, read column F as one bulk block of actual values and scan
        it directly. A candidate "last part number" row is only trusted
        once F_COLUMN_GAP_CHECK consecutive rows after it are confirmed
        empty (normal header gaps in this template are at most ~2 blank
        rows) -- otherwise treat whatever's found next as more real data
        and keep scanning past it.
        """
        scan_end = OFFER_DATA_START_ROW + F_COLUMN_SCAN_ROWS
        try:
            vals = ws.range(f"F{OFFER_DATA_START_ROW}:F{scan_end}").value
        except Exception:
            return OFFER_DATA_START_ROW - 1
        if not isinstance(vals, list):
            vals = [vals]

        last_row = OFFER_DATA_START_ROW - 1
        for i, val in enumerate(vals):
            if val in (None, ""):
                continue
            row = OFFER_DATA_START_ROW + i
            window = vals[i + 1: i + 1 + F_COLUMN_GAP_CHECK]
            if all(v in (None, "") for v in window):
                return row
            last_row = row  # more data follows within the gap window -- keep going

        return last_row


# ── Entry point with robust error handling ────
def main():
    root = tk.Tk()
    try:
        app = QuoteWizardApp(root)
    except Exception:
        # If anything goes wrong, show the error and close
        messagebox.showerror("Fatal Error",
                             f"An unexpected error occurred:\n\n{traceback.format_exc()}")
        root.destroy()
        return
    root.mainloop()

if __name__ == "__main__":
    main()
