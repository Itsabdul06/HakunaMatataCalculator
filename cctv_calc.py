#!/usr/bin/env python3
"""
CCTV Master Calculator
Rewrite of KantechCalc with improved GUI.
Maintains all original functionality: camera entry, NVR management,
HDD pricing, auto/manual calculation, report export to Excel.
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
        # Running as compiled EXE
        return sys._MEIPASS
    else:
        # Running as script
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

def save_camera_database(db):
    """Save camera database to JSON file"""
    try:
        json_path = os.path.join(get_resource_path(), "Cameras_JSON.json")
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(db, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving camera database: {e}")
        return False

# Calculate storage in TB per camera
def calculate_storage_tb(mbps, days):
    """Calculate storage in TB per camera for given Mbps and retention days"""
    # Formula: (Mbps × 324 × (days/30)) / 1024
    # Simplified: Mbps × days × 0.01055
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
        self.selected_camera = tk.StringVar()
        self.selected_codec = tk.StringVar()
        self.selected_fps = tk.StringVar()
        self.camera_quantity = tk.StringVar(value="1")
        self.retention_days = tk.StringVar(value="30")
        self.calculated_mbps = tk.StringVar(value="0.00")
        self.calculated_storage = tk.StringVar(value="0.00")
        
        # Custom camera variables
        self.custom_vars = {}
        
        # Bind trace to update Mbps and Storage when selections change
        self.selected_camera.trace('w', self.update_codec_dropdown)
        self.selected_codec.trace('w', self.update_fps_dropdown)
        self.selected_fps.trace('w', self.update_mbps_and_storage)
        self.retention_days.trace('w', self.update_storage_only)

        self.load_all_data()
        self.setup_ui()
        self._apply_ttk_styles()
        
        # Populate camera dropdown after UI is built
        self.populate_camera_dropdown()

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

    # ── Tab 1: Cameras (Database + Manual Entry) ────────────────────────────
    def populate_camera_dropdown(self):
        """Populate camera dropdown with names from database"""
        camera_names = sorted(self.camera_db.keys())
        if camera_names:
            self.camera_dropdown['values'] = camera_names
            self.selected_camera.set(camera_names[0])
        else:
            self.camera_dropdown['values'] = ["No cameras found"]
    
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
    
    def add_custom_camera(self):
        """Add a manually entered camera to the tree"""
        try:
            name = self.custom_vars["custom_name"].get().strip()
            if not name:
                raise ValueError("Camera name is required")
            
            sku = self.custom_vars["custom_sku"].get().strip()
            if not sku:
                sku = name
            
            brand = self.custom_vars["custom_brand"].get().strip()
            if not brand:
                brand = "Custom"
            
            mbps_str = self.custom_vars["custom_mbps"].get().strip()
            if not mbps_str:
                raise ValueError("Mbps is required")
            mbps = float(mbps_str)
            if mbps <= 0:
                raise ValueError("Mbps must be positive")
            
            storage_str = self.custom_vars["custom_storage"].get().strip()
            if not storage_str:
                raise ValueError("Storage TB/cam is required")
            storage = float(storage_str)
            if storage <= 0:
                raise ValueError("Storage must be positive")
            
            quantity_str = self.custom_vars["custom_quantity"].get().strip()
            if not quantity_str:
                raise ValueError("Quantity is required")
            quantity = int(quantity_str)
            if quantity <= 0:
                raise ValueError("Quantity must be positive")
            
            # Add to tree
            tag = "even" if len(self.tree.get_children()) % 2 == 0 else "odd"
            self.tree.insert("", "end", values=(name, quantity, f"{mbps:.2f}", f"{storage:.2f}"), tags=(tag,))
            
            # Optionally add to database
            if self.add_to_database.get():
                if name not in self.camera_db:
                    self.camera_db[name] = {
                        "sku": sku,
                        "brand": brand,
                        "resolution": "Custom",
                        "throughputs": {
                            "Custom": {
                                "1fps": mbps
                            }
                        }
                    }
                    save_camera_database(self.camera_db)
                    self.populate_camera_dropdown()
                    messagebox.showinfo("Success", f"Camera '{name}' added to database.")
                else:
                    messagebox.showwarning("Warning", f"Camera '{name}' already exists in database. Not added.")
            
            # Clear custom fields
            for key in self.custom_vars:
                self.custom_vars[key].set("")
            self.add_to_database.set(False)
            
            self.refresh_nvr_dropdowns()
            
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid input: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add camera: {e}")
    
    def _build_cameras_tab(self, tab):
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(4, weight=1)

        # ─────────────────────────────────────────────────────────────────────
        # SECTION 1: Add Camera from Database
        # ─────────────────────────────────────────────────────────────────────
        db_frame = mk_frame(tab, bg=SURFACE)
        db_frame.grid(row=0, column=0, sticky="ew", padx=16, pady=(14, 5))

        mk_label(db_frame, "Add Camera from Database", font=FONT_H2, fg=ACCENT, bg=SURFACE).grid(
            row=0, column=0, columnspan=10, sticky="w", padx=14, pady=(10, 8))

        # Row 1: Camera Model
        mk_label(db_frame, "Camera Model:", bg=SURFACE, fg=TEXT2).grid(row=1, column=0, sticky="w", padx=(14, 5), pady=5)
        self.camera_dropdown = ttk.Combobox(db_frame, textvariable=self.selected_camera, width=50, state="readonly")
        self.camera_dropdown.grid(row=1, column=1, columnspan=3, sticky="w", padx=(0, 10), pady=5)
        self.camera_dropdown.bind("<<ComboboxSelected>>", self.update_codec_dropdown)
        
        # Row 2: Codec
        mk_label(db_frame, "Codec:", bg=SURFACE, fg=TEXT2).grid(row=2, column=0, sticky="w", padx=(14, 5), pady=5)
        self.codec_dropdown = ttk.Combobox(db_frame, textvariable=self.selected_codec, width=10, state="readonly")
        self.codec_dropdown.grid(row=2, column=1, sticky="w", padx=(0, 10), pady=5)
        
        # Row 3: FPS
        mk_label(db_frame, "FPS:", bg=SURFACE, fg=TEXT2).grid(row=3, column=0, sticky="w", padx=(14, 5), pady=5)
        self.fps_dropdown = ttk.Combobox(db_frame, textvariable=self.selected_fps, width=10, state="readonly")
        self.fps_dropdown.grid(row=3, column=1, sticky="w", padx=(0, 10), pady=5)
        
        # Row 4: Quantity
        mk_label(db_frame, "Quantity:", bg=SURFACE, fg=TEXT2).grid(row=4, column=0, sticky="w", padx=(14, 5), pady=5)
        qty_entry = mk_entry(db_frame, textvariable=self.camera_quantity, width=10)
        qty_entry.grid(row=4, column=1, sticky="w", padx=(0, 10), pady=5)
        
        # Row 5: Retention Days
        mk_label(db_frame, "Retention Days:", bg=SURFACE, fg=TEXT2).grid(row=5, column=0, sticky="w", padx=(14, 5), pady=5)
        days_entry = mk_entry(db_frame, textvariable=self.retention_days, width=10)
        days_entry.grid(row=5, column=1, sticky="w", padx=(0, 10), pady=5)
        
        # Row 6: Calculated Mbps
        mk_label(db_frame, "Mbps (calculated):", bg=SURFACE, fg=TEXT2).grid(row=6, column=0, sticky="w", padx=(14, 5), pady=5)
        mbps_label = mk_label(db_frame, "", font=FONT_MONO, fg=ACCENT, bg=SURFACE, width=12)
        mbps_label.grid(row=6, column=1, sticky="w", padx=(0, 10), pady=5)
        self.calculated_mbps.trace('w', lambda *args: mbps_label.config(text=self.calculated_mbps.get()))
        
        # Row 7: Calculated Storage
        mk_label(db_frame, "Storage TB/cam (calculated):", bg=SURFACE, fg=TEXT2).grid(row=7, column=0, sticky="w", padx=(14, 5), pady=5)
        storage_label = mk_label(db_frame, "", font=FONT_MONO, fg=ACCENT, bg=SURFACE, width=12)
        storage_label.grid(row=7, column=1, sticky="w", padx=(0, 10), pady=5)
        self.calculated_storage.trace('w', lambda *args: storage_label.config(text=self.calculated_storage.get()))
        
        # Buttons for database section
        db_btn_f = mk_frame(db_frame, bg=SURFACE)
        db_btn_f.grid(row=8, column=0, columnspan=4, pady=(10, 0))
        mk_btn(db_btn_f, "Add Camera from DB", self.add_camera_from_database, style="primary").pack(side="left", padx=(0, 6))

        sep1 = mk_frame(tab, bg=BORDER, height=2)
        sep1.grid(row=1, column=0, sticky="ew", padx=16, pady=10)

        # ─────────────────────────────────────────────────────────────────────
        # SECTION 2: Add Custom Camera (Manual Entry)
        # ─────────────────────────────────────────────────────────────────────
        custom_frame = mk_frame(tab, bg=SURFACE)
        custom_frame.grid(row=2, column=0, sticky="ew", padx=16, pady=(5, 14))

        mk_label(custom_frame, "Add Custom Camera (Manual Entry)", font=FONT_H2, fg=ACCENT, bg=SURFACE).grid(
            row=0, column=0, columnspan=10, sticky="w", padx=14, pady=(10, 8))

        custom_fields = [
            ("Camera Name:", "custom_name", 25),
            ("SKU (Part No.):", "custom_sku", 20),
            ("Brand:", "custom_brand", 25),
            ("Mbps:", "custom_mbps", 10),
            ("Storage TB/cam:", "custom_storage", 10),
            ("Quantity:", "custom_quantity", 10),
        ]
        
        for i, (label, key, width) in enumerate(custom_fields):
            mk_label(custom_frame, label, bg=SURFACE, fg=TEXT2).grid(row=i+1, column=0, sticky="w", padx=(14, 5), pady=5)
            var = tk.StringVar()
            e = mk_entry(custom_frame, textvariable=var, width=width)
            e.grid(row=i+1, column=1, sticky="w", padx=(0, 10), pady=5)
            self.custom_vars[key] = var
        
        # Option to add to database checkbox
        self.add_to_database = tk.BooleanVar(value=False)
        add_to_db_check = tk.Checkbutton(custom_frame, text="Add this camera to database (for future use)", 
                                          variable=self.add_to_database,
                                          bg=SURFACE, fg=TEXT2, selectcolor=SURFACE2,
                                          activebackground=SURFACE, activeforeground=TEXT,
                                          font=FONT_BODY)
        add_to_db_check.grid(row=len(custom_fields)+1, column=0, columnspan=2, sticky="w", padx=(14, 5), pady=5)
        
        # Buttons for custom section
        custom_btn_f = mk_frame(custom_frame, bg=SURFACE)
        custom_btn_f.grid(row=len(custom_fields)+2, column=0, columnspan=2, pady=(10, 0))
        mk_btn(custom_btn_f, "Add Custom Camera", self.add_custom_camera, style="primary").pack(side="left", padx=(0, 6))

        sep2 = mk_frame(tab, bg=BORDER, height=2)
        sep2.grid(row=3, column=0, sticky="ew", padx=16, pady=10)

        # ─────────────────────────────────────────────────────────────────────
        # Camera Tree
        # ─────────────────────────────────────────────────────────────────────
        tree_f = mk_frame(tab, bg=SURFACE2)
        tree_f.grid(row=4, column=0, sticky="nsew", padx=16, pady=14)
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

        # Delete selected button for the tree
        tree_btn_f = mk_frame(tree_f, bg=SURFACE2)
        tree_btn_f.grid(row=1, column=0, pady=10)
        mk_btn(tree_btn_f, "Delete Selected Camera", self.delete_camera, style="danger").pack()

        self.tree.bind("<<TreeviewSelect>>", self._on_cam_select)

    def _on_cam_select(self, event):
        sel = self.tree.selection()
        if not sel: return
        # Just keep selection, don't populate inputs

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
                                   state="readonly", values=["All", "Tyco - American Dynamics", "Tyco - Holis"])
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
                     state="readonly", values=["Tyco - American Dynamics", "Tyco - Holis"]).grid(row=1, column=15, padx=(0, 6), pady=(0, 10))
        
        mk_btn(add_f, "ADD TO DATABASE", self.add_new_nvr, style="primary").grid(
            row=1, column=16, padx=(6, 14), pady=(0, 10))
        
        # Delete button for selected NVR
        mk_btn(add_f, "DELETE SELECTED", self._delete_nvr_from_tree, style="danger").grid(
            row=1, column=17, padx=(6, 14), pady=(0, 10))

        sep(tab).grid(row=0, column=0, sticky="ew", padx=16)

        # Create a frame for the NVR list with Treeview
        list_frame = mk_frame(tab, bg=SURFACE2)
        list_frame.grid(row=1, column=0, sticky="nsew", padx=16, pady=14)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # Create Treeview for NVR list
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
        
        # Add scrollbars
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.nvr_tree.yview)
        hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=self.nvr_tree.xview)
        self.nvr_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.nvr_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        # Bind double-click to edit price
        self.nvr_tree.bind("<Double-1>", self._on_nvr_double_click)
        
        self.refresh_nvr_list_tab()

    def refresh_nvr_list_tab(self):
        """Refresh the NVR list in the treeview"""
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
        """Handle double-click on NVR row to edit price"""
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
        """Delete the selected NVR from the treeview"""
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
            
            # Clear input fields
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
                result = self.auto_calculate_optimized(cameras)
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
        """Optimized auto-calculation with multiprocessing"""
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
        
        # Remove dominated NVRs
        compatible_nvrs = self.filter_dominated_nvrs(compatible_nvrs)
        
        # Flatten cameras
        flat_cams = []
        for name, count, mbps, storage in cameras:
            for _ in range(count):
                flat_cams.append((name, mbps, storage))
        
        total_cam = len(flat_cams)
        total_bw = sum(x[1] for x in flat_cams)
        
        # Generate combinations
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
    
    def filter_dominated_nvrs(self, nvrs):
        """Remove NVRs that are strictly worse than another"""
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

    # ── Excel Export with Brand Column (D) ────────────────────────────────
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

            # Group identical NVRs
            nvr_groups = {}
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

            # Prepare data rows
            excel_rows = []
            
            # Cameras header
            excel_rows.append(("", "", "", "", "", "", "header", "Cameras", None))
            
            for cam in cameras:
                cam_name = cam[0]
                cam_qty = int(cam[1])
                cam_sku = self.camera_db.get(cam_name, {}).get("sku", cam_name)
                cam_brand = self.camera_db.get(cam_name, {}).get("brand", "")
                excel_rows.append((cam_sku, cam_qty, "", cam_brand, "CCTV", "Camera", "data", "", None))
                excel_rows.append(("ADVEC01", 1, "ch", "", "Tyco - American Dynamics","CCTV", "Software", "data", "", None))
            
            # NVRs header
            excel_rows.append(("", "", "", "", "", "", "header", "NVRs", None))
            
            for key, group in nvr_groups.items():
                excel_rows.append((group["sku"], group["count"], "", group["brand"], "CCTV", "NVR", "data", "", None))
                excel_rows.append((f"{group['hdd_cap']}TB HDD", group["hdd_qty"], "ch", "", "CCTV", "HDD", "data", "", None))
            
            # VMS header
            excel_rows.append(("", "", "", "", "", "", "header", "VMS", None))
            excel_rows.append(("ADVASC01", 1, "", "", "CCTV", "Software", "data", "", None))

            current_row = 9
            
            # Track which rows are headers to only clear those
            header_rows = []
            row_counter = current_row
            for row_data in excel_rows:
                if row_data[6] == "header":
                    header_rows.append(row_counter)
                row_counter += 1
            
            # ONLY clear the header rows
            for row in header_rows:
                ws.range(f"A{row}:M{row}").value = None

            # Write new data
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
