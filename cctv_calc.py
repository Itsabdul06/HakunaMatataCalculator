#!/usr/bin/env python3
"""
CCTV Master Calculator
Rewrite of KantechCalc with improved GUI.
Maintains all original functionality: camera entry, NVR management,
HDD pricing, auto/manual calculation, report export to Excel and PDF.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math, itertools, json, os, sys
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed
import threading

# Try to import xlwings for Excel export
try:
    import xlwings as xw
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False

# Try to import reportlab for PDF export
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

# ─────────────────────────── Persistence ───────────────────────────────────
DATA_FILE = "system_data.json"

DEFAULT_HDD_PRICES = {
    1: 87.0, 2: 131.0, 3: 145.0, 4: 239.0,
    6: 375.0, 8: 427.0, 10: 500.0, 12: 614.0,
    14: 1114.0, 18: 1291.0, 22: 1145.85, 24: 1568.0, 26: 1385.0,
}

DEFAULT_NVR_DATA = [
    {"Name": "1U RAID",        "SKU": "ADVER00N0NP16G", "CH": 32,  "MB": 50,   "Slots": 4,  "Price": 3750.0,  "mode": "RAID",   "brand": "Tyco - American Dynamics"},
    {"Name": "2U 64 Ch",       "SKU": "ADVER12R0N2H",   "CH": 64,  "MB": 300,  "Slots": 6,  "Price": 10416.7, "mode": "RAID",   "brand": "Tyco - American Dynamics"},
    {"Name": "2U 100 Ch",      "SKU": "ADVER00RN2J",    "CH": 100, "MB": 600,  "Slots": 8,  "Price": 11666.7, "mode": "RAID",   "brand": "Tyco - American Dynamics"},
    {"Name": "2U 128 Ch",      "SKU": "ADVER72R5N2H",   "CH": 128, "MB": 600,  "Slots": 12, "Price": 25000.0, "mode": "RAID",   "brand": "Tyco - American Dynamics"},
    {"Name": "2U Rack 175 Ch", "SKU": "ADVER02RDK",     "CH": 175, "MB": 1000, "Slots": 12, "Price": 13854.2, "mode": "RAID",   "brand": "Tyco - American Dynamics"},
    {"Name": "2U Rack 200 Ch", "SKU": "ADVER02RDK",     "CH": 200, "MB": 1500, "Slots": 12, "Price": 12812.5, "mode": "RAID",   "brand": "Tyco - American Dynamics"},
    {"Name": "Micro NVR",      "SKU": "ADVEM00N0NP8AH", "CH": 8,   "MB": 80,   "Slots": 1,  "Price": 1500.0,  "mode": "JBOD",   "brand": "Tyco - American Dynamics"},
    {"Name": "Desktop JBOD",   "SKU": "ADVED00N0N5H",   "CH": 50,  "MB": 200,  "Slots": 2,  "Price": 2291.7,  "mode": "JBOD",   "brand": "Tyco - American Dynamics"},
    {"Name": "2U 75 Ch",       "SKU": "ADVER00N0N2J",   "CH": 75,  "MB": 400,  "Slots": 4,  "Price": 5312.5,  "mode": "JBOD",   "brand": "Tyco - American Dynamics"},
    {"Name": "Holis 8 Ch",     "SKU": "HRN-08013P",     "CH": 8,   "MB": 160,  "Slots": 1,  "Price": 520.85,  "mode": "JBOD",   "brand": "Tyco - Holis"},
    {"Name": "Holis 16 Ch",    "SKU": "HRN-16023P",     "CH": 16,  "MB": 320,  "Slots": 2,  "Price": 770.85,  "mode": "JBOD",   "brand": "Tyco - Holis"},
    {"Name": "Exacq 1U",       "SKU": "EXQ-1U-001",     "CH": 32,  "MB": 100,  "Slots": 4,  "Price": 3200.0,  "mode": "RAID",   "brand": "Exacq"},
    {"Name": "Exacq 2U",       "SKU": "EXQ-2U-001",     "CH": 64,  "MB": 200,  "Slots": 8,  "Price": 5500.0,  "mode": "RAID",   "brand": "Exacq"},
]

# ─────────────────────────── Colors & Fonts ────────────────────────────────
BG       = "#0f1520"
SURFACE  = "#151d2e"
SURFACE2 = "#1a2540"
SURFACE3 = "#1f2d4a"
BORDER   = "#253046"
ACCENT   = "#00d4ff"
ACCENT_D = "#0099bb"
GREEN    = "#22d3a5"
GOLD     = "#f59e0b"
RED      = "#f87171"
TEXT     = "#e2e8f0"
TEXT2    = "#7a90b0"
TEXT3    = "#3d5070"
WHITE    = "#ffffff"

FONT_H1   = ("Segoe UI", 16, "bold")
FONT_H2   = ("Segoe UI", 11, "bold")
FONT_H3   = ("Segoe UI", 10, "bold")
FONT_BODY = ("Segoe UI", 9)
FONT_MONO = ("Consolas", 9)
FONT_BTN  = ("Segoe UI", 9, "bold")
FONT_LRGE = ("Segoe UI", 11, "bold")

# ================= OPTIMIZED HDD CACHE =================
hdd_cache = {}

def get_best_hdd_cached(required_tb, slots, parity, price_dict):
    """Cached HDD calculation - critical for performance"""
    key = (round(required_tb, 2), slots, parity)
    if key in hdd_cache:
        return hdd_cache[key]

    best_cost, best_cfg = float('inf'), None
    for cap, price in price_dict.items():
        if cap <= 0:
            continue

        drives_needed = math.ceil(required_tb / cap)
        if parity > 0:
            drives_needed += parity
        
        if drives_needed > slots:
            continue

        drives_needed = max(drives_needed, parity + 1 if parity else 1)
        cost = drives_needed * price

        if cost < best_cost:
            best_cost = cost
            best_cfg = {
                "cap": cap,
                "qty": drives_needed,
                "data": drives_needed - parity,
                "cost": cost,
                "total_capacity": drives_needed * cap
            }

    hdd_cache[key] = best_cfg
    return best_cfg

# ================= OPTIMIZED SOLVER =================
def solve_combo(flat_cams, nvrs, raid_mode, hdd_prices):
    """Optimized solver with pruning and early termination"""
    
    # Precompute prefix sums for O(1) range queries
    total_cameras = len(flat_cams)
    bw_prefix = [0] * (total_cameras + 1)
    st_prefix = [0] * (total_cameras + 1)
    
    for i, (_, mbps, storage) in enumerate(flat_cams):
        bw_prefix[i + 1] = bw_prefix[i] + mbps
        st_prefix[i + 1] = st_prefix[i] + storage
    
    n_nvrs = len(nvrs)
    parity = 0 if raid_mode == "JBOD" else (1 if raid_mode == "RAID 5" else 2)
    
    # Pre-calculate minimum possible cost for pruning
    min_nvr_cost = min(n["Price"] for n in nvrs)
    min_hdd_cost = min(hdd_prices.values())
    
    best_result = None
    best_cost = float('inf')
    
    # Use iterative deepening with pruning
    def dfs(idx, start, remaining, current_cost, assignment):
        nonlocal best_result, best_cost
        
        # Prune if we can't beat best_cost
        lower_bound = current_cost + (n_nvrs - idx) * (min_nvr_cost + min_hdd_cost)
        if lower_bound >= best_cost:
            return
        
        if idx == n_nvrs - 1:
            # Last NVR takes all remaining
            assignment.append(remaining)
            
            # Quick check: last NVR must be able to handle bandwidth
            if start < total_cameras:
                end = total_cameras
                bw = bw_prefix[end] - bw_prefix[start]
                if bw > nvrs[idx]["MB"]:
                    assignment.pop()
                    return
            
            # Build result
            pos = 0
            result = []
            total = 0
            valid = True
            
            for i, nvr in enumerate(nvrs):
                take = assignment[i]
                if take > 0:
                    s, e = pos, pos + take
                    bw = bw_prefix[e] - bw_prefix[s]
                    st = st_prefix[e] - st_prefix[s]
                    pos = e
                    
                    if take > nvr["CH"] or bw > nvr["MB"]:
                        valid = False
                        break
                    
                    hdd = get_best_hdd_cached(st, nvr["Slots"], parity, hdd_prices)
                    if not hdd:
                        valid = False
                        break
                    
                    cost = nvr["Price"] + hdd["cost"]
                    total += cost
                    
                    # Count camera types for display
                    cam_counts = {}
                    for j in range(s, e):
                        name = flat_cams[j][0]
                        cam_counts[name] = cam_counts.get(name, 0) + 1
                    
                    result.append({
                        "nvr": nvr,
                        "camera_count": take,
                        "cam_breakdown": cam_counts,
                        "total_storage": st,
                        "total_bandwidth": bw,
                        "hdd_config": hdd,
                        "cost": cost
                    })
            
            if valid and pos == total_cameras and total < best_cost:
                best_cost = total
                best_result = result
            
            assignment.pop()
            return
        
        # Calculate realistic range for this NVR
        nvr = nvrs[idx]
        
        # Max by channel
        max_take = min(nvr["CH"], remaining - (n_nvrs - idx - 1))
        
        # Max by bandwidth (estimate using prefix sums)
        if start < total_cameras:
            # Binary search for max cameras this NVR can handle by bandwidth
            low, high = 1, min(max_take, remaining)
            max_by_bw = 0
            while low <= high:
                mid = (low + high) // 2
                if start + mid <= total_cameras:
                    bw = bw_prefix[start + mid] - bw_prefix[start]
                    if bw <= nvr["MB"]:
                        max_by_bw = mid
                        low = mid + 1
                    else:
                        high = mid - 1
                else:
                    high = mid - 1
            max_take = min(max_take, max_by_bw)
        
        # Min take: at least 1, and leave enough for remaining NVRs
        min_take = max(1, remaining - (n_nvrs - idx - 1) * max(nvr["CH"] for nvr in nvrs[idx+1:]))
        min_take = max(min_take, 1)
        
        if min_take > max_take:
            return
        
        # Try from largest to smallest (finds cheaper solutions faster)
        for take in range(max_take, min_take - 1, -1):
            # Estimate cost for this branch
            if start + take <= total_cameras:
                st = st_prefix[start + take] - st_prefix[start]
                hdd_est = get_best_hdd_cached(st, nvr["Slots"], parity, hdd_prices)
                if hdd_est:
                    est_cost = current_cost + nvr["Price"] + hdd_est["cost"]
                    if est_cost >= best_cost:
                        continue
            
            assignment.append(take)
            dfs(idx + 1, start + take, remaining - take, 
                current_cost + nvr["Price"], assignment)
            assignment.pop()
    
    dfs(0, 0, total_cameras, 0, [])
    return best_result, best_cost

# ================= CAMERA DATABASE LOADER =================
def get_resource_path():
    """Get the correct path for resources whether running as script or compiled EXE"""
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    else:
        return os.path.dirname(os.path.abspath(__file__))

def load_camera_database():
    """Load camera database from JSON file"""
    try:
        json_path = os.path.join(get_resource_path(), "Cameras_JSON.json")
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print("Warning: Cameras_JSON.json not found. Using empty database.")
        return {}
    except json.JSONDecodeError as e:
        print(f"Error parsing Cameras_JSON.json: {e}")
        return {}

# Calculate storage in TB per camera
def calculate_storage_tb(mbps, days):
    """Calculate storage in TB per camera for given Mbps and retention days"""
    return mbps * days * 0.01055

# ─────────────────────────── Widget Helpers ────────────────────────────────
def mk_frame(parent, bg=SURFACE, **kw):
    return tk.Frame(parent, bg=bg, **kw)

def mk_label(parent, text, font=FONT_BODY, fg=TEXT2, bg=SURFACE, anchor="w", **kw):
    return tk.Label(parent, text=text, font=font, fg=fg, bg=bg, anchor=anchor, **kw)

def mk_entry(parent, textvariable=None, width=12, font=FONT_MONO, **kw):
    defaults = dict(
        bg=SURFACE2, fg=TEXT, insertbackground=ACCENT,
        relief="flat", bd=0,
        highlightthickness=1, highlightbackground=BORDER,
        highlightcolor=ACCENT,
    )
    defaults.update(kw)
    e = tk.Entry(parent, textvariable=textvariable, width=width,
                 font=font, **defaults)
    return e

def mk_btn(parent, text, command, style="normal", **kw):
    colors = {
        "primary": (ACCENT,   "#000000", ACCENT_D),
        "danger":  (SURFACE2, RED,       SURFACE3),
        "ghost":   (SURFACE2, TEXT2,     SURFACE3),
        "success": (GREEN,    "#000000", "#18a87f"),
        "normal":  (SURFACE3, TEXT,      BORDER),
    }
    bg, fg, abg = colors.get(style, colors["normal"])
    return tk.Button(parent, text=text, command=command,
                     bg=bg, fg=fg, activebackground=abg, activeforeground=fg,
                     font=FONT_BTN, relief="flat", bd=0,
                     cursor="hand2", padx=10, pady=5, **kw)

def sep(parent, bg=BORDER, vertical=False):
    if vertical:
        return tk.Frame(parent, bg=bg, width=1)
    return tk.Frame(parent, bg=bg, height=1)

# ─────────────────────────── Application ───────────────────────────────────
class CCTVApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CCTV Master Calculator")
        self.root.configure(bg=BG)
        self.root.geometry("1200x820")
        self.root.minsize(1000, 700)

        self.last_report = None
        self.last_calculation_result = None
        self.hdd_ents = {}
        self.nvr_price_vars = {}
        self.progress_window = None
        self.brand_filter = tk.StringVar(value="All")
        
        # Camera database
        self.camera_db = load_camera_database()
        
        # UI Variables for camera entry
        self.selected_camera_type = tk.StringVar(value="All")
        self.selected_resolution = tk.StringVar(value="All")
        self.selected_camera = tk.StringVar()
        self.selected_codec = tk.StringVar()
        self.selected_fps = tk.StringVar()
        self.camera_quantity = tk.StringVar(value="1")
        self.retention_days = tk.StringVar(value="30")
        self.calculated_mbps = tk.StringVar(value="0.00")
        self.calculated_storage = tk.StringVar(value="0.00")
        
        # Get unique camera types and resolutions from database
        types = set()
        resolutions = set()
        for cam_data in self.camera_db.values():
            if cam_data.get("type"):
                types.add(cam_data.get("type"))
            if cam_data.get("resolution"):
                resolutions.add(cam_data.get("resolution"))
        
        self.camera_types = ["All"] + sorted(types)
        self.resolutions = ["All"] + sorted(resolutions, key=lambda x: [int(n) for n in x.split('x')] if 'x' in x else [0, 0])
        
        # Bind trace to update Mbps and Storage when selections change
        self.selected_camera_type.trace('w', self.update_camera_dropdown)
        self.selected_resolution.trace('w', self.update_camera_dropdown)
        self.selected_camera.trace('w', self.update_codec_dropdown)
        self.selected_codec.trace('w', self.update_fps_dropdown)
        self.selected_fps.trace('w', self.update_mbps_and_storage)
        self.retention_days.trace('w', self.update_storage_only)

        self.load_all_data()
        self.setup_ui()
        self._apply_ttk_styles()
        
        # Populate camera dropdown after UI is built
        self.update_camera_dropdown()

    def load_all_data(self):
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, "r") as f:
                    data = json.load(f)
                self.hdd_prices = {int(k): float(v) for k, v in data.get("hdd", {}).items()}
                self.nvr_list = [dict(x) for x in data.get("nvr", [])]
                return
            except Exception:
                pass
        self.hdd_prices = dict(DEFAULT_HDD_PRICES)
        self.nvr_list = [dict(n) for n in DEFAULT_NVR_DATA]

    def save_all_data(self):
        with open(DATA_FILE, "w") as f:
            json.dump({"hdd": self.hdd_prices, "nvr": self.nvr_list}, f, indent=2)

    def _apply_ttk_styles(self):
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("TNotebook", background=BG, borderwidth=0, tabmargins=0)
        s.configure("TNotebook.Tab",
                    background=SURFACE, foreground=TEXT2,
                    font=FONT_H3, padding=(16, 8),
                    borderwidth=0, focuscolor=BG)
        s.map("TNotebook.Tab",
              background=[("selected", SURFACE2), ("active", SURFACE3)],
              foreground=[("selected", ACCENT), ("active", TEXT)])
        s.configure("Treeview",
                    background=SURFACE, foreground=TEXT,
                    fieldbackground=SURFACE, rowheight=24,
                    font=FONT_MONO, borderwidth=0)
        s.configure("Treeview.Heading",
                    background=SURFACE2, foreground=ACCENT,
                    font=FONT_H3, relief="flat", borderwidth=0)
        s.map("Treeview",
              background=[("selected", ACCENT_D)],
              foreground=[("selected", WHITE)])
        s.map("Treeview.Heading", relief=[("active", "flat")])
        s.configure("Vertical.TScrollbar", background=BORDER, troughcolor=SURFACE, arrowcolor=TEXT3, borderwidth=0)
        s.configure("Horizontal.TScrollbar", background=BORDER, troughcolor=SURFACE, arrowcolor=TEXT3, borderwidth=0)
        s.configure("TCombobox",
                    fieldbackground=SURFACE2, background=SURFACE2,
                    foreground=TEXT, bordercolor=BORDER,
                    arrowcolor=ACCENT, selectbackground=SURFACE2,
                    selectforeground=TEXT)
        s.map("TCombobox",
              fieldbackground=[("readonly", SURFACE2)],
              foreground=[("readonly", TEXT)])

    def setup_ui(self):
        hdr = mk_frame(self.root, bg=BG)
        hdr.pack(fill="x", padx=24, pady=(18, 0))
        mk_label(hdr, "CCTV Master Calculator", font=FONT_H1, fg=WHITE, bg=BG).pack(side="left")
        mk_label(hdr, "  v37.0", font=FONT_BODY, fg=TEXT3, bg=BG).pack(side="left", pady=(6, 0))
        sep(self.root).pack(fill="x", padx=24, pady=10)

        self.nb = ttk.Notebook(self.root, style="TNotebook")
        self.nb.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        self.tabs = []
        for title in ["Cameras", "Calculate", "NVR Models", "HDD Prices"]:
            f = mk_frame(self.nb, bg=SURFACE2)
            self.nb.add(f, text=f"  {title}  ")
            self.tabs.append(f)

        self._build_cameras_tab(self.tabs[0])
        self._build_calc_tab(self.tabs[1])
        self._build_nvr_tab(self.tabs[2])
        self._build_hdd_tab(self.tabs[3])

    # ── Tab 1: Cameras (Database with Type and Resolution Filters) ────────
    def update_camera_dropdown(self, *args):
        """Update camera dropdown based on selected camera type and resolution"""
        camera_type = self.selected_camera_type.get()
        resolution = self.selected_resolution.get()
        
        filtered_cameras = []
        for name, data in self.camera_db.items():
            # Filter by type
            if camera_type != "All" and data.get("type", "") != camera_type:
                continue
            # Filter by resolution
            if resolution != "All" and data.get("resolution", "") != resolution:
                continue
            filtered_cameras.append(name)
        
        filtered_cameras.sort()
        
        if filtered_cameras:
            self.camera_dropdown['values'] = filtered_cameras
            self.selected_camera.set(filtered_cameras[0])
        else:
            self.camera_dropdown['values'] = ["No cameras found"]
            self.selected_camera.set("")
    
    def update_codec_dropdown(self, *args):
        """Update codec dropdown based on selected camera"""
        camera_name = self.selected_camera.get()
        if camera_name and camera_name in self.camera_db:
            codecs = list(self.camera_db[camera_name].get("throughputs", {}).keys())
            self.codec_dropdown['values'] = codecs
            if codecs:
                self.selected_codec.set(codecs[0])
            else:
                self.selected_codec.set("")
        else:
            self.codec_dropdown['values'] = []
            self.selected_codec.set("")
    
    def update_fps_dropdown(self, *args):
        """Update FPS dropdown based on selected camera and codec"""
        camera_name = self.selected_camera.get()
        codec = self.selected_codec.get()
        if camera_name and camera_name in self.camera_db and codec:
            throughputs = self.camera_db[camera_name].get("throughputs", {}).get(codec, {})
            fps_list = sorted(throughputs.keys(), key=lambda x: int(x.replace('fps', '')))
            self.fps_dropdown['values'] = fps_list
            if fps_list:
                self.selected_fps.set(fps_list[0])
            else:
                self.selected_fps.set("")
        else:
            self.fps_dropdown['values'] = []
            self.selected_fps.set("")
    
    def update_mbps_and_storage(self, *args):
        """Update Mbps and Storage based on current selections"""
        camera_name = self.selected_camera.get()
        codec = self.selected_codec.get()
        fps = self.selected_fps.get()
        
        if camera_name and camera_name in self.camera_db and codec and fps:
            mbps = self.camera_db[camera_name].get("throughputs", {}).get(codec, {}).get(fps, 0)
            self.calculated_mbps.set(f"{mbps:.2f}")
            self.update_storage_only()
        else:
            self.calculated_mbps.set("0.00")
            self.calculated_storage.set("0.00")
    
    def update_storage_only(self, *args):
        """Update only storage calculation (when retention days changes)"""
        try:
            mbps = float(self.calculated_mbps.get())
            days = float(self.retention_days.get()) if self.retention_days.get() else 0
            storage_tb = calculate_storage_tb(mbps, days)
            self.calculated_storage.set(f"{storage_tb:.2f}")
        except ValueError:
            self.calculated_storage.set("0.00")
    
    def add_camera_from_database(self):
        """Add camera to the tree using values from database selection"""
        try:
            camera_name = self.selected_camera.get()
            if not camera_name or camera_name == "No cameras found":
                messagebox.showerror("Error", "Please select a valid camera model.")
                return
            
            quantity = self.camera_quantity.get().strip()
            if not quantity:
                raise ValueError("Quantity cannot be empty")
            
            mbps = float(self.calculated_mbps.get())
            if mbps <= 0:
                raise ValueError("Invalid Mbps value")
            
            storage_tb = float(self.calculated_storage.get())
            if storage_tb <= 0:
                raise ValueError("Invalid storage value")
            
            qty = int(quantity)
            if qty <= 0:
                raise ValueError("Quantity must be positive")
            
            # Add to tree
            tag = "even" if len(self.tree.get_children()) % 2 == 0 else "odd"
            self.tree.insert("", "end", values=(camera_name, qty, f"{mbps:.2f}", f"{storage_tb:.2f}"), tags=(tag,))
            self.refresh_nvr_dropdowns()
            
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid input: {e}")
    
    def update_selected_camera(self):
        """Update the selected camera in the tree with current values"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select a camera to update.")
            return
        
        try:
            camera_name = self.selected_camera.get()
            if not camera_name or camera_name == "No cameras found":
                messagebox.showerror("Error", "Please select a valid camera model.")
                return
            
            quantity = self.camera_quantity.get().strip()
            if not quantity:
                raise ValueError("Quantity cannot be empty")
            
            mbps = float(self.calculated_mbps.get())
            if mbps <= 0:
                raise ValueError("Invalid Mbps value")
            
            storage_tb = float(self.calculated_storage.get())
            if storage_tb <= 0:
                raise ValueError("Invalid storage value")
            
            qty = int(quantity)
            if qty <= 0:
                raise ValueError("Quantity must be positive")
            
            # Update the selected item
            tag = "even" if self.tree.index(sel[0]) % 2 == 0 else "odd"
            self.tree.item(sel[0], values=(camera_name, qty, f"{mbps:.2f}", f"{storage_tb:.2f}"), tags=(tag,))
            self.refresh_nvr_dropdowns()
            
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid input: {e}")
    
    def _build_cameras_tab(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(4, weight=1)

        # ─────────────────────────────────────────────────────────────────────
        # SECTION: Add Camera from Database
        # ─────────────────────────────────────────────────────────────────────
        db_frame = mk_frame(tab, bg=SURFACE)
        db_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 5))

        mk_label(db_frame, "Add Camera from Database", font=FONT_H2, fg=ACCENT, bg=SURFACE).grid(
            row=0, column=0, columnspan=10, sticky="w", padx=14, pady=(10, 8))

        # Row 1: Camera Type Filter
        mk_label(db_frame, "Camera Type:", bg=SURFACE, fg=TEXT2).grid(row=1, column=0, sticky="w", padx=(14, 5), pady=5)
        type_dropdown = ttk.Combobox(db_frame, textvariable=self.selected_camera_type, width=20, state="readonly")
        type_dropdown.grid(row=1, column=1, sticky="w", padx=(0, 10), pady=5)
        type_dropdown['values'] = self.camera_types
        
        # Row 2: Resolution Filter
        mk_label(db_frame, "Resolution:", bg=SURFACE, fg=TEXT2).grid(row=2, column=0, sticky="w", padx=(14, 5), pady=5)
        resolution_dropdown = ttk.Combobox(db_frame, textvariable=self.selected_resolution, width=20, state="readonly")
        resolution_dropdown.grid(row=2, column=1, sticky="w", padx=(0, 10), pady=5)
        resolution_dropdown['values'] = self.resolutions
        
        # Row 3: Camera Model
        mk_label(db_frame, "Camera Model:", bg=SURFACE, fg=TEXT2).grid(row=3, column=0, sticky="w", padx=(14, 5), pady=5)
        self.camera_dropdown = ttk.Combobox(db_frame, textvariable=self.selected_camera, width=50, state="readonly")
        self.camera_dropdown.grid(row=3, column=1, columnspan=3, sticky="w", padx=(0, 10), pady=5)
        
        # Row 4: Codec
        mk_label(db_frame, "Codec:", bg=SURFACE, fg=TEXT2).grid(row=4, column=0, sticky="w", padx=(14, 5), pady=5)
        self.codec_dropdown = ttk.Combobox(db_frame, textvariable=self.selected_codec, width=10, state="readonly")
        self.codec_dropdown.grid(row=4, column=1, sticky="w", padx=(0, 10), pady=5)
        
        # Row 5: FPS
        mk_label(db_frame, "FPS:", bg=SURFACE, fg=TEXT2).grid(row=5, column=0, sticky="w", padx=(14, 5), pady=5)
        self.fps_dropdown = ttk.Combobox(db_frame, textvariable=self.selected_fps, width=10, state="readonly")
        self.fps_dropdown.grid(row=5, column=1, sticky="w", padx=(0, 10), pady=5)
        
        # Row 6: Quantity
        mk_label(db_frame, "Quantity:", bg=SURFACE, fg=TEXT2).grid(row=6, column=0, sticky="w", padx=(14, 5), pady=5)
        qty_entry = mk_entry(db_frame, textvariable=self.camera_quantity, width=10)
        qty_entry.grid(row=6, column=1, sticky="w", padx=(0, 10), pady=5)
        
        # Row 7: Retention Days
        mk_label(db_frame, "Retention Days:", bg=SURFACE, fg=TEXT2).grid(row=7, column=0, sticky="w", padx=(14, 5), pady=5)
        days_entry = mk_entry(db_frame, textvariable=self.retention_days, width=10)
        days_entry.grid(row=7, column=1, sticky="w", padx=(0, 10), pady=5)
        
        # Row 8: Calculated Mbps
        mk_label(db_frame, "Mbps (calculated):", bg=SURFACE, fg=TEXT2).grid(row=8, column=0, sticky="w", padx=(14, 5), pady=5)
        mbps_label = mk_label(db_frame, "", font=FONT_MONO, fg=ACCENT, bg=SURFACE, width=12)
        mbps_label.grid(row=8, column=1, sticky="w", padx=(0, 10), pady=5)
        self.calculated_mbps.trace('w', lambda *args: mbps_label.config(text=self.calculated_mbps.get()))
        
        # Row 9: Calculated Storage
        mk_label(db_frame, "Storage TB/cam (calculated):", bg=SURFACE, fg=TEXT2).grid(row=9, column=0, sticky="w", padx=(14, 5), pady=5)
        storage_label = mk_label(db_frame, "", font=FONT_MONO, fg=ACCENT, bg=SURFACE, width=12)
        storage_label.grid(row=9, column=1, sticky="w", padx=(0, 10), pady=5)
        self.calculated_storage.trace('w', lambda *args: storage_label.config(text=self.calculated_storage.get()))
        
        # Buttons for database section
        db_btn_f = mk_frame(db_frame, bg=SURFACE)
        db_btn_f.grid(row=10, column=0, columnspan=4, pady=(10, 0))
        mk_btn(db_btn_f, "Add Camera", self.add_camera_from_database, style="primary").pack(side="left", padx=(0, 6))
        mk_btn(db_btn_f, "Update Selected", self.update_selected_camera, style="ghost").pack(side="left", padx=(0, 6))
        mk_btn(db_btn_f, "Delete Selected", self.delete_camera, style="danger").pack(side="left")

        sep1 = mk_frame(tab, bg=BORDER, height=2)
        sep1.grid(row=1, column=0, sticky="ew", padx=16, pady=10)

        # ─────────────────────────────────────────────────────────────────────
        # Camera Tree
        # ─────────────────────────────────────────────────────────────────────
        tree_f = mk_frame(tab, bg=SURFACE2)
        tree_f.grid(row=2, column=0, sticky="nsew", padx=16, pady=14)
        tree_f.columnconfigure(0, weight=1)
        tree_f.rowconfigure(0, weight=1)

        cols = ("Name", "Count", "Mbps/cam", "Storage TB/cam")
        self.tree = ttk.Treeview(tree_f, columns=cols, show="headings")
        widths = [400, 80, 100, 130]
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="center" if c != "Name" else "w")
        self.tree.tag_configure("odd", background=SURFACE)
        self.tree.tag_configure("even", background=SURFACE2)

        vsb = ttk.Scrollbar(tree_f, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self.tree.bind("<<TreeviewSelect>>", self._on_cam_select)

    def _on_cam_select(self, event):
        sel = self.tree.selection()
        if not sel: 
            return
        vals = self.tree.item(sel[0])["values"]
        if vals:
            camera_name = vals[0]
            if camera_name in self.camera_db:
                cam_type = self.camera_db[camera_name].get("type", "All")
                cam_res = self.camera_db[camera_name].get("resolution", "All")
                if cam_type in self.camera_types:
                    self.selected_camera_type.set(cam_type)
                if cam_res in self.resolutions:
                    self.selected_resolution.set(cam_res)
                self.selected_camera.set(camera_name)
                self.camera_quantity.set(str(vals[1]))

    def delete_camera(self):
        for s in self.tree.selection():
            self.tree.delete(s)
        self.refresh_nvr_dropdowns()

    # ── Tab 2: Calculate ──────────────────────────────────────────────────
    def _build_calc_tab(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        ctrl = mk_frame(tab, bg=SURFACE)
        ctrl.grid(row=0, column=0, sticky="ew", padx=16, pady=14)

        mk_label(ctrl, "Calculation Settings", font=FONT_H2, fg=ACCENT, bg=SURFACE).pack(
            anchor="w", padx=14, pady=(10, 8))

        row1 = mk_frame(ctrl, bg=SURFACE)
        row1.pack(fill="x", padx=14, pady=(0, 10))

        mk_label(row1, "Mode:", bg=SURFACE, fg=TEXT2).pack(side="left", padx=(0, 6))
        self.auto_mode = tk.StringVar(value="AUTO")
        rb_auto = tk.Radiobutton(row1, text="Auto (find best NVR combo)", variable=self.auto_mode, value="AUTO",
                                 bg=SURFACE, fg=TEXT2, selectcolor=SURFACE2,
                                 activebackground=SURFACE, activeforeground=TEXT,
                                 font=FONT_BODY, command=self._on_mode_change)
        rb_auto.pack(side="left", padx=(0, 16))
        rb_manual = tk.Radiobutton(row1, text="Manual (choose NVR below)", variable=self.auto_mode, value="MANUAL",
                                   bg=SURFACE, fg=TEXT2, selectcolor=SURFACE2,
                                   activebackground=SURFACE, activeforeground=TEXT,
                                   font=FONT_BODY, command=self._on_mode_change)
        rb_manual.pack(side="left", padx=(0, 16))

        mk_label(row1, "RAID Level:", bg=SURFACE, fg=TEXT2).pack(side="left", padx=(16, 6))
        self.raid_var = tk.StringVar(value="JBOD")
        cb_raid = ttk.Combobox(row1, textvariable=self.raid_var, width=10,
                               state="readonly", values=["JBOD", "RAID 5", "RAID 6"])
        cb_raid.pack(side="left")

        row2 = mk_frame(ctrl, bg=SURFACE)
        row2.pack(fill="x", padx=14, pady=(0, 10))
        mk_label(row2, "NVR Brand:", bg=SURFACE, fg=TEXT2).pack(side="left", padx=(0, 6))
        self.brand_filter = tk.StringVar(value="All")
        brand_combo = ttk.Combobox(row2, textvariable=self.brand_filter, width=25,
                                   state="readonly", values=["All", "Tyco - American Dynamics", "Tyco - Holis", "Exacq"])
        brand_combo.bind("<<ComboboxSelected>>", lambda x: self.refresh_nvr_dropdowns())
        brand_combo.pack(side="left")
        mk_label(row2, "(Filters NVRs shown below)", bg=SURFACE, fg=TEXT3, font=FONT_BODY).pack(side="left", padx=(10, 0))

        self.manual_frame = mk_frame(ctrl, bg=SURFACE)
        self.manual_frame.pack(fill="x", padx=14, pady=(0, 10))
        manual_label = mk_label(self.manual_frame, "Manual NVR Selection:", font=FONT_H2, fg=ACCENT, bg=SURFACE)
        manual_label.pack(anchor="w", pady=(0, 10))
        self.manual_combos = []
        for i in range(6):
            row_frame = mk_frame(self.manual_frame, bg=SURFACE)
            row_frame.pack(fill="x", pady=2)
            mk_label(row_frame, f"NVR {i+1}:", bg=SURFACE, fg=TEXT2, width=8).pack(side="left", padx=(0, 5))
            var = tk.StringVar(value="None")
            cb = ttk.Combobox(row_frame, textvariable=var, width=30,
                             state="readonly", values=["None"])
            cb.pack(side="left", padx=(0, 10))
            self.manual_combos.append(cb)
        self.manual_frame.pack_forget()

        btn_row = mk_frame(ctrl, bg=SURFACE)
        btn_row.pack(fill="x", padx=14, pady=(0, 12))
        mk_btn(btn_row, "⚡  Run Calculation", self.run_logic, style="primary").pack(side="left", padx=(0, 10))
        mk_btn(btn_row, "Export to Excel", self.export_to_excel, style="success").pack(side="left", padx=(0, 10))
        mk_btn(btn_row, "Export to PDF", self.export_to_pdf, style="ghost").pack(side="left", padx=(0, 10))
        self.calc_status = mk_label(btn_row, "", fg=TEXT2, bg=SURFACE, font=FONT_BODY)
        self.calc_status.pack(side="left", padx=16)

        sep(tab).grid(row=0, column=0, sticky="ew", padx=16)

        res_f = mk_frame(tab, bg=SURFACE2)
        res_f.grid(row=1, column=0, sticky="nsew", padx=16, pady=14)
        res_f.columnconfigure(0, weight=1)
        res_f.rowconfigure(1, weight=1)

        mk_label(res_f, "Results", font=FONT_H2, fg=WHITE, bg=SURFACE2).grid(
            row=0, column=0, columnspan=2, sticky="w", padx=14, pady=(10, 6))

        self.res_txt = tk.Text(res_f, bg=SURFACE, fg=TEXT, font=FONT_MONO,
                               relief="flat", bd=0, state="disabled",
                               highlightthickness=0, wrap="none",
                               padx=14, pady=10)
        vsb2 = ttk.Scrollbar(res_f, orient="vertical", command=self.res_txt.yview)
        hsb2 = ttk.Scrollbar(res_f, orient="horizontal", command=self.res_txt.xview)
        self.res_txt.configure(yscrollcommand=vsb2.set, xscrollcommand=hsb2.set)
        self.res_txt.grid(row=1, column=0, sticky="nsew")
        vsb2.grid(row=1, column=1, sticky="ns")
        hsb2.grid(row=2, column=0, sticky="ew")

        self.res_txt.tag_configure("header", foreground=ACCENT, font=("Consolas", 9, "bold"))
        self.res_txt.tag_configure("best", foreground=GREEN, font=("Consolas", 9, "bold"))
        self.res_txt.tag_configure("label", foreground=TEXT2)
        self.res_txt.tag_configure("value", foreground=TEXT)
        self.res_txt.tag_configure("divider", foreground=TEXT3)
        self.res_txt.tag_configure("cost", foreground=GOLD, font=("Consolas", 9, "bold"))
        self.res_txt.tag_configure("error", foreground=RED)

        self.refresh_nvr_dropdowns()
        self._on_mode_change()

    def _on_mode_change(self):
        if self.auto_mode.get() == "MANUAL":
            self.refresh_nvr_dropdowns()
            self.manual_frame.pack(fill="x", padx=14, pady=(0, 10))
        else:
            self.manual_frame.pack_forget()

    def refresh_nvr_dropdowns(self):
        brand = self.brand_filter.get()
        if brand == "All":
            filtered_nvrs = self.nvr_list
        else:
            filtered_nvrs = [n for n in self.nvr_list if n.get("brand", "") == brand]
        names = ["None"] + [n["Name"] for n in filtered_nvrs]
        for combo in self.manual_combos:
            current = combo.get()
            combo['values'] = names
            if current not in names:
                combo.set("None")

    def show_progress(self):
        if self.progress_window and self.progress_window.winfo_exists():
            return
        self.progress_window = tk.Toplevel(self.root)
        self.progress_window.title("Calculating...")
        self.progress_window.configure(bg=SURFACE)
        self.progress_window.geometry("300x100")
        self.progress_window.transient(self.root)
        self.progress_window.grab_set()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 150
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 50
        self.progress_window.geometry(f"300x100+{x}+{y}")
        mk_label(self.progress_window, "Analyzing configurations...",
                font=FONT_H2, fg=ACCENT, bg=SURFACE).pack(pady=(20, 10))
        self.progress_bar = ttk.Progressbar(self.progress_window, mode='indeterminate')
        self.progress_bar.pack(padx=20, pady=10, fill='x')
        self.progress_bar.start(10)
        self.root.update()

    def hide_progress(self):
        if self.progress_window and self.progress_window.winfo_exists():
            self.progress_window.destroy()
        self.progress_window = None

    # ── Tab 3: NVR Models (Treeview with proper alignment) ──────────────────
    def _build_nvr_tab(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(1, weight=1)

        add_f = mk_frame(tab, bg=SURFACE)
        add_f.grid(row=0, column=0, sticky="ew", padx=16, pady=14)
        mk_label(add_f, "Add New NVR Model", font=FONT_H2, fg=ACCENT, bg=SURFACE).grid(
            row=0, column=0, columnspan=16, sticky="w", padx=14, pady=(10, 8))

        self.nf = {}
        fields = [("Name", 14), ("SKU", 14), ("CH", 6), ("MB", 6), ("Slots", 6), ("Price", 8)]
        for col, (f, w) in enumerate(fields):
            mk_label(add_f, f, bg=SURFACE, fg=TEXT2).grid(row=1, column=col*2, sticky="w", padx=(14 if col==0 else 6, 3))
            var = tk.StringVar()
            e = mk_entry(add_f, textvariable=var, width=w)
            e.grid(row=1, column=col*2+1, padx=(0, 2), pady=(0, 10))
            self.nf[f] = var

        self.na = tk.StringVar(value="RAID")
        mk_label(add_f, "RAID/JBOD", bg=SURFACE, fg=TEXT2).grid(row=1, column=12, sticky="w", padx=(6, 3))
        ttk.Combobox(add_f, textvariable=self.na, width=7,
                     state="readonly", values=["RAID", "JBOD"]).grid(row=1, column=13, padx=(0, 6), pady=(0, 10))
        
        self.nf_brand = tk.StringVar(value="Tyco - American Dynamics")
        mk_label(add_f, "Brand:", bg=SURFACE, fg=TEXT2).grid(row=1, column=14, sticky="w", padx=(6, 3))
        ttk.Combobox(add_f, textvariable=self.nf_brand, width=20,
                     state="readonly", values=["Tyco - American Dynamics", "Tyco - Holis", "Exacq"]).grid(row=1, column=15, padx=(0, 6), pady=(0, 10))
        
        mk_btn(add_f, "ADD TO DATABASE", self.add_new_nvr, style="primary").grid(
            row=1, column=16, padx=(6, 14), pady=(0, 10))
        
        mk_btn(add_f, "DELETE SELECTED", self._delete_nvr_from_tree, style="danger").grid(
            row=1, column=17, padx=(6, 14), pady=(0, 10))

        sep(tab).grid(row=0, column=0, sticky="ew", padx=16)

        list_frame = mk_frame(tab, bg=SURFACE2)
        list_frame.grid(row=1, column=0, sticky="nsew", padx=16, pady=14)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        columns = ("Name", "SKU", "Channels", "Bandwidth", "Slots", "Price", "Mode", "Brand")
        self.nvr_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)
        
        column_configs = [
            ("Name", 200, "w"),
            ("SKU", 150, "w"),
            ("Channels", 70, "center"),
            ("Bandwidth", 80, "center"),
            ("Slots", 60, "center"),
            ("Price", 100, "e"),
            ("Mode", 60, "center"),
            ("Brand", 180, "w"),
        ]
        
        for col, width, anchor in column_configs:
            self.nvr_tree.heading(col, text=col)
            self.nvr_tree.column(col, width=width, anchor=anchor)
        
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.nvr_tree.yview)
        hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=self.nvr_tree.xview)
        self.nvr_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.nvr_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        self.nvr_tree.bind("<Double-1>", self._on_nvr_double_click)
        
        self.refresh_nvr_list_tab()

    def refresh_nvr_list_tab(self):
        for item in self.nvr_tree.get_children():
            self.nvr_tree.delete(item)
        
        for i, n in enumerate(self.nvr_list):
            tag = "even" if i % 2 == 0 else "odd"
            price_str = f"${n['Price']:,.2f}"
            
            self.nvr_tree.insert("", "end", values=(
                n["Name"],
                n["SKU"],
                n["CH"],
                n["MB"],
                n["Slots"],
                price_str,
                n.get("mode", "RAID"),
                n.get("brand", "Tyco - American Dynamics"),
            ), tags=(tag,))
        
        self.nvr_tree.tag_configure("odd", background=SURFACE)
        self.nvr_tree.tag_configure("even", background=SURFACE2)

    def _on_nvr_double_click(self, event):
        item = self.nvr_tree.selection()[0] if self.nvr_tree.selection() else None
        if not item:
            return
        
        values = self.nvr_tree.item(item, "values")
        if not values:
            return
        
        current_price = float(values[5].replace("$", "").replace(",", ""))
        
        edit_window = tk.Toplevel(self.root)
        edit_window.title("Edit Price")
        edit_window.configure(bg=SURFACE)
        edit_window.geometry("300x120")
        edit_window.transient(self.root)
        edit_window.grab_set()
        
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 150
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 60
        edit_window.geometry(f"300x120+{x}+{y}")
        
        mk_label(edit_window, f"Edit price for {values[0]}", font=FONT_H2, fg=ACCENT, bg=SURFACE).pack(pady=(10, 5))
        
        price_var = tk.StringVar(value=f"{current_price:.2f}")
        price_entry = mk_entry(edit_window, textvariable=price_var, width=15)
        price_entry.pack(pady=5)
        
        def save_price():
            try:
                new_price = float(price_var.get())
                if new_price <= 0:
                    raise ValueError("Price must be positive")
                
                for idx, nvr in enumerate(self.nvr_list):
                    if nvr["Name"] == values[0] and nvr["SKU"] == values[1]:
                        self.nvr_list[idx]["Price"] = new_price
                        break
                
                self.refresh_nvr_list_tab()
                edit_window.destroy()
                
            except ValueError as e:
                messagebox.showerror("Error", f"Invalid price: {e}")
        
        btn_frame = mk_frame(edit_window, bg=SURFACE)
        btn_frame.pack(pady=10)
        mk_btn(btn_frame, "Save", save_price, style="primary").pack(side="left", padx=5)
        mk_btn(btn_frame, "Cancel", edit_window.destroy, style="ghost").pack(side="left", padx=5)

    def _delete_nvr_from_tree(self):
        selected = self.nvr_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an NVR to delete.")
            return
        
        if messagebox.askyesno("Confirm", "Delete the selected NVR model?"):
            for item in selected:
                values = self.nvr_tree.item(item, "values")
                if values:
                    for idx, nvr in enumerate(self.nvr_list):
                        if nvr["Name"] == values[0] and nvr["SKU"] == values[1]:
                            self.nvr_list.pop(idx)
                            break
                self.nvr_tree.delete(item)
            
            self.save_all_data()
            self.refresh_nvr_dropdowns()

    def add_new_nvr(self):
        try:
            name = self.nf["Name"].get().strip()
            if not name:
                raise ValueError("Name required")
            sku = self.nf["SKU"].get().strip()
            if not sku:
                raise ValueError("SKU required")
            ch_str = self.nf["CH"].get().strip()
            if not ch_str:
                raise ValueError("Channels required")
            mb_str = self.nf["MB"].get().strip()
            if not mb_str:
                raise ValueError("Max MB/s required")
            slots_str = self.nf["Slots"].get().strip()
            if not slots_str:
                raise ValueError("HDD Slots required")
            price_str = self.nf["Price"].get().strip()
            if not price_str:
                raise ValueError("Price required")

            row = {
                "Name": name, "SKU": sku, "CH": int(ch_str), "MB": int(mb_str),
                "Slots": int(slots_str), "Price": float(price_str),
                "mode": self.na.get(), "brand": self.nf_brand.get(),
            }

            if row["CH"] <= 0 or row["MB"] <= 0 or row["Slots"] <= 0 or row["Price"] <= 0:
                raise ValueError("All values must be positive")

            self.nvr_list.append(row)
            self.save_all_data()
            self.refresh_nvr_dropdowns()
            self.refresh_nvr_list_tab()
            
            for f in self.nf.values():
                f.set("")
            self.na.set("RAID")
            self.nf_brand.set("Tyco - American Dynamics")
            
            messagebox.showinfo("Success", "NVR Added.")
        except Exception as e:
            messagebox.showerror("Error", f"Invalid input: {e}")

    # ── Tab 4: HDD Prices ─────────────────────────────────────────────────
    def _build_hdd_tab(self, tab):
        tab.columnconfigure(0, weight=1)

        outer = mk_frame(tab, bg=SURFACE)
        outer.grid(row=0, column=0, sticky="nsew", padx=16, pady=14)

        mk_label(outer, "Hard Drive Prices  (EGP per drive)", font=FONT_H2, fg=ACCENT, bg=SURFACE).pack(
            anchor="w", padx=14, pady=(12, 10))

        grid = mk_frame(outer, bg=SURFACE)
        grid.pack(fill="x", padx=14, pady=(0, 10))

        self.hdd_ents = {}
        for i, cap in enumerate(sorted(self.hdd_prices.keys())):
            col, row = (i % 4) * 3, i // 4
            mk_label(grid, f"{cap} TB", fg=TEXT2, bg=SURFACE, width=6).grid(
                row=row, column=col, sticky="w", padx=(0, 4), pady=5)
            var = tk.StringVar(value=f"{self.hdd_prices[cap]:.2f}")
            e = mk_entry(grid, textvariable=var, width=10)
            e.grid(row=row, column=col+1, padx=(0, 24), pady=5)
            self.hdd_ents[cap] = var

        btn_row = mk_frame(outer, bg=SURFACE)
        btn_row.pack(anchor="w", padx=14, pady=(6, 14))
        mk_btn(btn_row, "Save HDD Prices", self.save_hdds, style="success").pack(side="left")

    def save_hdds(self):
        for cap, var in self.hdd_ents.items():
            try:
                price = float(var.get())
                if price <= 0:
                    raise ValueError("Price must be positive")
                self.hdd_prices[cap] = price
            except ValueError as e:
                messagebox.showerror("Error", f"Invalid price for {cap}TB: {e}")
                return
        self.save_all_data()
        messagebox.showinfo("Saved", "HDD Prices Updated.")

    # ── MAIN CALCULATION LOGIC ────────────────────────────────────────────────
    def run_logic(self):
        camera_rows = [self.tree.item(i)["values"] for i in self.tree.get_children()]
        if not camera_rows:
            messagebox.showwarning("Warning", "Add cameras first.")
            return
        
        cameras = []
        for row in camera_rows:
            cameras.append((row[0], int(row[1]), float(row[2]), float(row[3])))
        
        self.calc_status.config(text="Calculating...", fg=GOLD)
        self.show_progress()
        
        def worker():
            try:
                # Check which mode is selected
                if self.auto_mode.get() == "AUTO":
                    result = self.auto_calculate_optimized(cameras)
                else:
                    # MANUAL mode - use selected NVRs from dropdowns
                    result = self.manual_calculate(cameras)
                
                self.root.after(0, lambda: self._finish_calc(result))
            except Exception as e:
                import traceback
                traceback.print_exc()
                self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        
        threading.Thread(target=worker, daemon=True).start()
    
    def _finish_calc(self, result):
        self.hide_progress()
        
        if not result:
            self.calc_status.config(text="No solution found", fg=RED)
            self._show_result_error("ERROR: No valid configuration found.\n\nPossible reasons:\n• HDD sizes cannot meet storage requirements\n• NVR channel/slot limits exceeded\n• No compatible NVRs available for selected RAID mode")
            return
        
        self.last_calculation_result = {
            "cameras": [self.tree.item(i)["values"] for i in self.tree.get_children()],
            "nvr_config": result,
            "raid_mode": self.raid_var.get()
        }
        
        self.display_results(result)
        total_cost = sum(u["cost"] for u in result)
        self.calc_status.config(text=f"Done — Total: ${total_cost:,.2f}", fg=GREEN)
    
    def auto_calculate_optimized(self, cameras):
        brand = self.brand_filter.get()
        if brand == "All":
            available_nvrs = self.nvr_list.copy()
        else:
            available_nvrs = [n for n in self.nvr_list if n.get("brand", "") == brand]
        
        raid_mode = self.raid_var.get()
        compatible_nvrs = []
        for nvr in available_nvrs:
            nvr_mode = nvr.get("mode", "RAID")
            if raid_mode == "JBOD" and nvr_mode == "JBOD":
                compatible_nvrs.append(nvr)
            elif raid_mode != "JBOD" and nvr_mode == "RAID":
                compatible_nvrs.append(nvr)
        
        if not compatible_nvrs:
            compatible_nvrs = available_nvrs
        
        compatible_nvrs = self.filter_dominated_nvrs(compatible_nvrs)
        
        flat_cams = []
        for name, count, mbps, storage in cameras:
            for _ in range(count):
                flat_cams.append((name, mbps, storage))
        
        total_cam = len(flat_cams)
        total_bw = sum(x[1] for x in flat_cams)
        
        combos_to_test = []
        for k in range(1, min(5, len(compatible_nvrs) + 2)):
            for combo in itertools.combinations_with_replacement(compatible_nvrs, k):
                if sum(n["CH"] for n in combo) < total_cam:
                    continue
                if sum(n["MB"] for n in combo) < total_bw:
                    continue
                combos_to_test.append((flat_cams.copy(), list(combo), raid_mode, self.hdd_prices))
        
        if not combos_to_test:
            return None
        
        best_result = None
        best_cost = float('inf')
        
        with ProcessPoolExecutor(max_workers=min(4, len(combos_to_test))) as executor:
            futures = {executor.submit(solve_combo, *c): c for c in combos_to_test}
            
            for future in as_completed(futures):
                try:
                    res, cost = future.result()
                    if res and cost < best_cost:
                        best_cost = cost
                        best_result = res
                except Exception as e:
                    print(f"Error processing combo: {e}")
        
        return best_result
    def manual_calculate(self, cameras):
        """Manual calculation - uses user-selected NVRs"""
        selected_nvrs = []
        for combo in self.manual_combos:
            nvr_name = combo.get()
            if nvr_name != "None" and nvr_name:
                nvr = next((n for n in self.nvr_list if n["Name"] == nvr_name), None)
                if nvr:
                    selected_nvrs.append(nvr)
        
        if not selected_nvrs:
            messagebox.showwarning("Warning", "Select at least one NVR.")
            return None
        
        # Get RAID mode
        raid_mode = self.raid_var.get()
        
        # Calculate total requirements
        total_cameras = sum(c[1] for c in cameras)
        total_bandwidth = sum(c[1] * c[2] for c in cameras)
        
        # Quick feasibility check
        total_channels = sum(nvr["CH"] for nvr in selected_nvrs)
        if total_channels < total_cameras:
            messagebox.showwarning("Warning", f"Selected NVRs have only {total_channels} channels, but you need {total_cameras} cameras.")
            return None
        
        total_bandwidth_capacity = sum(nvr["MB"] for nvr in selected_nvrs)
        if total_bandwidth_capacity < total_bandwidth:
            messagebox.showwarning("Warning", f"Selected NVRs have only {total_bandwidth_capacity} Mbps bandwidth, but you need {total_bandwidth:.1f} Mbps.")
            return None
        
        # Flatten cameras
        flat_cams = []
        for name, count, mbps, storage in cameras:
            for _ in range(count):
                flat_cams.append((name, mbps, storage))
        
        # Use the distribution function
        return self.distribute_cameras_simple(cameras, selected_nvrs)
        
    def distribute_cameras_simple(self, cameras, nvrs):
        """Simple camera distribution for manual mode"""
        # Flatten cameras
        flat_cams = []
        for name, count, mbps, storage in cameras:
            for _ in range(count):
                flat_cams.append((name, mbps, storage))
        
        total_cams = len(flat_cams)
        n_nvrs = len(nvrs)
        
        # Sort NVRs by bandwidth (smallest first) to put limited NVRs first
        nvrs_sorted = sorted(enumerate(nvrs), key=lambda x: x[1]["MB"])
        
        # Calculate target cameras for each NVR
        target_cams = [0] * n_nvrs
        remaining = total_cams
        
        # First pass: allocate to smallest bandwidth NVRs first
        for idx, nvr in nvrs_sorted:
            if remaining <= 0:
                break
            
            # Calculate average bandwidth of remaining cameras
            avg_bandwidth = sum(c[1] for c in flat_cams[:remaining]) / remaining if remaining > 0 else 0
            max_by_bandwidth = int(nvr["MB"] / avg_bandwidth) if avg_bandwidth > 0 else nvr["CH"]
            max_for_nvr = min(nvr["CH"], max_by_bandwidth, remaining)
            
            if max_for_nvr > 0:
                target_cams[idx] = max_for_nvr
                remaining -= max_for_nvr
        
        # If we couldn't allocate all cameras, distribute remaining evenly
        if remaining > 0:
            for idx, nvr in enumerate(nvrs):
                if remaining <= 0:
                    break
                if target_cams[idx] < nvr["CH"]:
                    take = min(nvr["CH"] - target_cams[idx], remaining)
                    target_cams[idx] += take
                    remaining -= take
        
        if remaining > 0:
            return None
        
        # Distribute cameras according to target
        result = []
        idx = 0
        raid_mode = self.raid_var.get()
        parity = 0 if raid_mode == "JBOD" else (1 if raid_mode == "RAID 5" else 2)
        
        for i, nvr in enumerate(nvrs):
            take = target_cams[i]
            if take <= 0:
                continue
            
            cam_slice = flat_cams[idx:idx + take]
            idx += take
            
            total_storage = sum(c[2] for c in cam_slice)
            total_bandwidth = sum(c[1] for c in cam_slice)
            
            # Check bandwidth limit
            if total_bandwidth > nvr["MB"]:
                return None
            
            # Get HDD configuration
            hdd = get_best_hdd_cached(total_storage, nvr["Slots"], parity, self.hdd_prices)
            if hdd is None:
                return None
            
            # Count camera types
            cam_counts = {}
            for c in cam_slice:
                cam_counts[c[0]] = cam_counts.get(c[0], 0) + 1
            
            result.append({
                "nvr": nvr,
                "cameras": cam_slice,
                "camera_count": take,
                "cam_breakdown": cam_counts,
                "total_storage": total_storage,
                "total_bandwidth": total_bandwidth,
                "hdd_config": hdd,
                "cost": nvr["Price"] + hdd["cost"]
            })
        
        return result if idx == total_cams else None
    
    def filter_dominated_nvrs(self, nvrs):
        filtered = []
        for i, a in enumerate(nvrs):
            dominated = False
            for j, b in enumerate(nvrs):
                if i == j:
                    continue
                if (b["CH"] >= a["CH"] and 
                    b["MB"] >= a["MB"] and 
                    b["Slots"] >= a["Slots"] and
                    b["Price"] <= a["Price"]):
                    dominated = True
                    break
            if not dominated:
                filtered.append(a)
        return filtered

    def display_results(self, result):
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        total = sum(u["cost"] for u in result)
        lines = []

        def write(text, tag="value"):
            lines.append((text, tag))

        write("=" * 72 + "\n", "divider")
        write(f" CCTV DESIGN REPORT  —  {now}\n", "header")
        write(f" SYSTEM TOTAL: ${total:,.2f}\n", "cost")
        write("=" * 72 + "\n", "divider")

        for i, u in enumerate(result, 1):
            nvr = u["nvr"]
            hdd = u["hdd_config"]
            write(f"\nUNIT #{i}: {nvr['Name']}\n", "best")
            write("-" * 50 + "\n", "divider")
            write(f"  Mode:     ", "label")
            write(f"{self.raid_var.get()}\n", "value")
            write(f"  Load:     ", "label")
            
            mbps_total = u["total_bandwidth"]
            mbps_capacity = nvr["MB"]
            load_pct = (mbps_total / mbps_capacity * 100) if mbps_capacity > 0 else 0
            
            write(f"{mbps_total:.1f} Mbps  ({load_pct:.1f}% of {nvr['MB']} Mbps capacity)\n", "value")
            
            write(f"  Cameras:  ", "label")
            write(f"{u['camera_count']} total  ", "value")
            if u.get("cam_breakdown"):
                parts = ",  ".join(f"{n}: {c}" for n, c in u["cam_breakdown"].items())
                write(f"({parts})\n", "value")
            else:
                write("\n", "value")
            write(f"  Storage:  ", "label")
            drive_str = f"{hdd['qty']} × {hdd['cap']} TB"
            total_cap = hdd['qty'] * hdd['cap']
            usable_tb = hdd['data'] * hdd['cap']
            write(f"{drive_str}  = {total_cap:.1f} TB  ", "value")
            write(f"(usable: {usable_tb:.1f} TB)\n", "label")
            write(f"  Cost:     ", "label")
            write(f"NVR ${nvr['Price']:,.2f}  +  HDD ${hdd['cost']:,.2f}  =  ${u['cost']:,.2f}\n", "cost")

        write("\n" + "=" * 72 + "\n", "divider")
        write(f" GRAND TOTAL:  ${total:,.2f}\n", "cost")
        write("=" * 72 + "\n", "divider")

        self.res_txt.config(state="normal")
        self.res_txt.delete("1.0", "end")
        for text, tag in lines:
            self.res_txt.insert("end", text, tag)
        self.res_txt.config(state="disabled")

        self.last_report = "".join(t for t, _ in lines)
        self.nb.select(self.tabs[1])

    def _show_result_error(self, msg):
        self.res_txt.config(state="normal")
        self.res_txt.delete("1.0", "end")
        self.res_txt.insert("end", msg, "error")
        self.res_txt.config(state="disabled")

    # ── PDF Export ──────────────────────────────────────────────────────────
    def export_to_pdf(self):
        """Export calculation results to PDF (without prices)"""
        if not self.last_calculation_result:
            messagebox.showwarning("Warning", "Run a calculation first before exporting!")
            return
        
        if not PDF_AVAILABLE:
            messagebox.showerror("Error", 
                "PDF export requires reportlab library.\n\n"
                "Please install it using:\npip install reportlab")
            return
        
        # Ask for save location
        pdf_file = filedialog.asksaveasfilename(
            title="Save PDF Report",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialfile=f"CCTV_Design_{datetime.now().strftime('%Y%m%d_%H%M')}"
        )
        
        if not pdf_file:
            return
        
        # Show progress
        progress_msg = tk.Toplevel(self.root)
        progress_msg.title("Generating PDF...")
        progress_msg.configure(bg=SURFACE)
        progress_msg.geometry("300x80")
        progress_msg.transient(self.root)
        progress_msg.grab_set()
        mk_label(progress_msg, "Generating PDF...",
                font=FONT_H2, fg=ACCENT, bg=SURFACE).pack(pady=(20, 10))
        self.root.update()
        
        try:
            # Create PDF document
            doc = SimpleDocTemplate(pdf_file, pagesize=landscape(letter),
                                   rightMargin=36, leftMargin=36,
                                   topMargin=72, bottomMargin=72)
            
            styles = getSampleStyleSheet()
            title_style = styles['Title']
            heading_style = styles['Heading2']
            normal_style = styles['Normal']
            
            # Create a custom style for wrapped text
            wrapped_style = ParagraphStyle(
                'WrappedStyle',
                parent=styles['Normal'],
                fontSize=9,
                leading=11,
                wordWrap='CJK'
            )
            
            story = []
            
            # Title
            story.append(Paragraph("CCTV Design Report", title_style))
            story.append(Spacer(1, 0.25 * inch))
            story.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", normal_style))
            story.append(Spacer(1, 0.3 * inch))
            
            # Camera List
            story.append(Paragraph("Camera List", heading_style))
            story.append(Spacer(1, 0.1 * inch))
            
            camera_data = [["Camera Name", "Qty", "Mbps", "Storage (TB)"]]
            for cam in self.last_calculation_result["cameras"]:
                camera_name = Paragraph(str(cam[0]), wrapped_style)
                quantity = str(cam[1])
                
                try:
                    mbps_val = float(cam[2])
                    mbps_str = f"{mbps_val:.2f}"
                except (ValueError, TypeError):
                    mbps_str = str(cam[2])
                
                try:
                    storage_val = float(cam[3])
                    storage_str = f"{storage_val:.2f}"
                except (ValueError, TypeError):
                    storage_str = str(cam[3])
                
                camera_data.append([camera_name, quantity, mbps_str, storage_str])
            
            camera_table = Table(camera_data, colWidths=[3.2*inch, 0.5*inch, 0.7*inch, 0.9*inch])
            camera_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(camera_table)
            story.append(Spacer(1, 0.3 * inch))
            
            # NVR Configuration (without prices)
            story.append(Paragraph("NVR Configuration", heading_style))
            story.append(Spacer(1, 0.1 * inch))
            
            nvr_data = [["Unit", "NVR Model", "Cameras", "Bandwidth", "Storage"]]
            for i, unit in enumerate(self.last_calculation_result["nvr_config"], 1):
                nvr = unit["nvr"]
                hdd = unit["hdd_config"]
                
                nvr_name = Paragraph(str(nvr["Name"]), wrapped_style)
                unit_num = str(i)
                camera_count = str(unit["camera_count"])
                
                try:
                    bw_val = float(unit["total_bandwidth"])
                    bw_str = f"{bw_val:.1f} Mbps"
                except (ValueError, TypeError):
                    bw_str = str(unit["total_bandwidth"])
                
                try:
                    hdd_qty = int(hdd["qty"])
                    hdd_cap = float(hdd["cap"])
                    storage_str = f"{hdd_qty} x {hdd_cap:.0f} TB"
                except (ValueError, TypeError):
                    storage_str = str(hdd.get("qty", "?")) + " x " + str(hdd.get("cap", "?")) + " TB"
                
                nvr_data.append([unit_num, nvr_name, camera_count, bw_str, storage_str])
            
            nvr_table = Table(nvr_data, colWidths=[0.5*inch, 3*inch, 0.7*inch, 1.2*inch, 1.2*inch])
            nvr_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            story.append(nvr_table)
            
            # Build PDF
            doc.build(story)
            
            progress_msg.destroy()
            messagebox.showinfo("Success", f"PDF report saved to:\n{pdf_file}")
            
        except Exception as e:
            progress_msg.destroy()
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to create PDF:\n{str(e)}")

    # ── Excel Export ──────────────────────────────────────────────────────
    def export_to_excel(self):
        if not self.last_calculation_result:
            messagebox.showwarning("Warning", "Run a calculation first before exporting!")
            return
        if not EXCEL_AVAILABLE:
            messagebox.showerror("Error", "Excel export requires xlwings.\nInstall: pip install xlwings")
            return

        template_file = filedialog.askopenfilename(
            title="Select Excel Template",
            filetypes=[("Excel files", "*.xlsx")])
        if not template_file:
            return

        save_option = messagebox.askyesno("Save Option",
            "Yes = Save as new file\nNo = Overwrite template")
        output_file = template_file
        if save_option:
            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"CCTV_Quote_{datetime.now().strftime('%Y%m%d_%H%M')}")
            if not output_file:
                return

        progress_msg = tk.Toplevel(self.root)
        progress_msg.title("Exporting...")
        progress_msg.configure(bg=SURFACE)
        progress_msg.geometry("300x80")
        progress_msg.transient(self.root)
        progress_msg.grab_set()
        mk_label(progress_msg, "Exporting to Excel...",
                font=FONT_H2, fg=ACCENT, bg=SURFACE).pack(pady=(20, 10))
        self.root.update()

        app = None
        wb = None
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(template_file)

            sheet_name = None
            for name in wb.sheet_names:
                if name.lower() == "offer":
                    sheet_name = name
                    break
            if not sheet_name:
                raise Exception("Sheet 'offer' not found!")

            ws = wb.sheets[sheet_name]
            cameras = self.last_calculation_result["cameras"]
            nvr_config = self.last_calculation_result["nvr_config"]

            # Group identical NVRs and detect brands
            nvr_groups = {}
            has_exacq = False
            
            for unit in nvr_config:
                sku = unit["nvr"]["SKU"]
                hdd_cap = unit["hdd_config"]["cap"]
                hdd_qty = unit["hdd_config"]["qty"]
                brand = unit["nvr"].get("brand", "Tyco - American Dynamics")
                key = (sku, hdd_cap, hdd_qty, brand)
                if key not in nvr_groups:
                    nvr_groups[key] = {"sku": sku, "hdd_cap": hdd_cap, "hdd_qty": hdd_qty, "count": 1, "brand": brand}
                else:
                    nvr_groups[key]["count"] += 1
                
                if "Exacq" in brand:
                    has_exacq = True

            # Prepare data rows
            excel_rows = []
            
            # Cameras header
            excel_rows.append(("", "", "", "", "", "", "header", "Cameras", None))
            
            # Store first camera info for Override sheet
            first_camera = cameras[0] if cameras else None
            first_camera_fps = None
            first_camera_retention = None
            
            for i, cam in enumerate(cameras):
                cam_name = cam[0]
                cam_qty = int(cam[1])
                cam_sku = self.camera_db.get(cam_name, {}).get("sku", cam_name)
                cam_brand = self.camera_db.get(cam_name, {}).get("brand", "")
                excel_rows.append((cam_sku, cam_qty, "", cam_brand, "CCTV", "Camera", "data", "", None))
                
                if has_exacq:
                    excel_rows.append(("EVIP-01", 1, "ch", "Exacq", "CCTV", "Software", "data", "", None))
                else:
                    excel_rows.append(("ADVEC01", 1, "ch", "Tyco - American Dynamics", "CCTV", "Software", "data", "", None))
                
                if i == 0:
                    first_camera_fps = self.selected_fps.get() if self.selected_fps.get() else "30fps"
                    first_camera_retention = self.retention_days.get() if self.retention_days.get() else "30"
            
            # NVRs header
            excel_rows.append(("", "", "", "", "", "", "header", "NVRs", None))
            
            for key, group in nvr_groups.items():
                excel_rows.append((group["sku"], group["count"], "", group["brand"], "CCTV", "NVR", "data", "", None))
                excel_rows.append((f"{group['hdd_cap']}TB HDD", group["hdd_qty"], "ch", "", "CCTV", "HDD", "data", "", None))
            
            # VMS header
            excel_rows.append(("", "", "", "", "", "", "header", "VMS", None))
            
            if has_exacq:
                excel_rows.append(("EXACQVMS", 1, "", "Exacq", "CCTV", "Software", "data", "", None))
            else:
                excel_rows.append(("ADVASC01", 1, "", "Tyco - American Dynamics", "CCTV", "Software", "data", "", None))
            
            # Add Workstation and Monitor rows
            if has_exacq:
                excel_rows.append(("Monitor", 1, "ch", "", "CCTV", "Local", "data", "", None))
            else:
                excel_rows.append(("Workstation", 1, "ch", "", "CCTV", "Local", "data", "", None))
                excel_rows.append(("Monitor", 1, "ch", "", "CCTV", "Local", "data", "", None))

            current_row = 9
            
            header_rows = []
            row_counter = current_row
            for row_data in excel_rows:
                if row_data[6] == "header":
                    header_rows.append(row_counter)
                row_counter += 1
            
            for row in header_rows:
                ws.range(f"A{row}:M{row}").value = None

            for row_data in excel_rows:
                part_no, qty, sys, brand, solution, category, row_type, header_text, _ = row_data
                
                if row_type == "header":
                    if header_text:
                        ws.range(f"G{current_row}").value = header_text
                    try:
                        ws.range(f"A{current_row}:M{current_row}").api.Style = "CG - Header 1"
                    except:
                        pass
                else:
                    if part_no:
                        ws.range(f"F{current_row}").value = part_no
                    if qty:
                        ws.range(f"H{current_row}").value = qty
                    if sys:
                        ws.range(f"K{current_row}").value = sys
                    if brand:
                        ws.range(f"D{current_row}").value = brand
                    if solution:
                        ws.range(f"L{current_row}").value = solution
                    if category:
                        ws.range(f"M{current_row}").value = category
                
                current_row += 1

            # Write to Override sheet if it exists
            try:
                override_sheet = wb.sheets["Override"]
                if first_camera_fps:
                    override_sheet.range("R6").value = first_camera_fps
                if first_camera_retention:
                    override_sheet.range("Q6").value = int(first_camera_retention)
            except Exception:
                pass

            wb.save(output_file)
            wb.close()
            app.quit()
            progress_msg.destroy()
            messagebox.showinfo("Success", f"Exported to {os.path.basename(output_file)}")

        except Exception as e:
            progress_msg.destroy()
            if app:
                try:
                    app.quit()
                except:
                    pass
            messagebox.showerror("Error", f"Export failed: {str(e)}")

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    
    root = tk.Tk()
    app = CCTVApp(root)
    root.mainloop()
