"""
Lastrada Report Processor
Automatically renames and organizes Shelly & Sands Lastrada quality control PDF reports
"""

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import tkinter as tk
import winreg
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from pypdf import PdfReader, PdfWriter
from datetime import datetime

try:
    import requests as _requests
except ImportError:
    _requests = None  # update feature disabled if requests is not installed

try:
    import pystray as _pystray
    from pystray import MenuItem as _TrayItem
except ImportError:
    _pystray = None  # tray icon not available; overwatch still works without it

from PIL import Image as _PILImage, ImageDraw as _PILDraw


# App version — bump this string before publishing a new GitHub release
VERSION = "1.0.13"

# How often (seconds) the Overwatch mode scans the source folder for new files
OVERWATCH_INTERVAL = 30

# GitHub auto-update settings (private repo, read-only token)
GITHUB_OWNER     = "zhuffman93"
GITHUB_REPO_NAME = "lastrada-report-processor"
GITHUB_TOKEN     = "[REDACTED-PAT]"

# Config file location - persists settings between sessions
CONFIG_PATH = Path(os.environ.get('APPDATA', Path.home())) / 'LastradaReportProcessor' / 'config.json'
STATS_PATH  = Path(os.environ.get('APPDATA', Path.home())) / 'LastradaReportProcessor' / 'stats.json'
MERGES_PATH      = Path(os.environ.get('APPDATA', Path.home())) / 'LastradaReportProcessor' / 'merges.json'

# IPC paths used by the pdf_filler watcher (Excel workbooks drop a request here)
PDF_REQUEST_PATH = Path(os.environ.get('APPDATA', Path.home())) / 'LastradaReportProcessor' / 'pdf_request.json'
PDF_RESULT_PATH  = Path(os.environ.get('APPDATA', Path.home())) / 'LastradaReportProcessor' / 'pdf_result.json'

# Labs available in the dropdown
LABS = [
    "Lab 1", "Lab 2", "Lab 3", "Lab 4", "Lab 6", "Lab 7", "Lab 8", "Lab 9",
    "Lab 10", "Lab 12", "Lab 13", "Lab 14", "Lab 17", "Lab 21", "Lab 22",
    "Lab 23", "Lab 24", "Lab 26", "Lab 27", "Lab 28",
]

# OneDrive base path - same structure on every machine, only the username differs
_ONEDRIVE_BASE = (
    Path.home()
    / "OneDrive - Shelly & Sands, Inc"
    / "Mar-Zane Lab - MZ Lab Tech Info"
    / "Plant Labs"
)

# Destination "Plant Testing" folder for each lab, built dynamically from the user's home directory
LAB_DESTINATIONS = {lab: str(_ONEDRIVE_BASE / lab / "Plant Testing") for lab in LABS}

# Colour themes
LIGHT_THEME = {
    'bg':         '#f0f0f0',
    'fg':         '#333333',
    'entry_bg':   'white',
    'entry_fg':   '#000000',
    'tree_bg':    'white',
    'heading_bg': '#e0e0e0',
    'log_bg':     'white',
    'log_fg':     '#333333',
    'select_bg':  '#0078d4',
    # Treeview row tag colours
    'tag_pending': '#856404',
    'tag_ready':   '#004085',
    'tag_success': '#155724',
    'tag_skipped': '#856404',
    'tag_error':   '#721c24',
}
DARK_THEME = {
    'bg':         '#2b2b2b',
    'fg':         '#e0e0e0',
    'entry_bg':   '#3c3c3c',
    'entry_fg':   '#e0e0e0',
    'tree_bg':    '#3c3c3c',
    'heading_bg': '#3c3c3c',
    'log_bg':     '#1e1e1e',
    'log_fg':     '#d4d4d4',
    'select_bg':  '#005a9e',
    # Treeview row tag colours — lighter shades for dark backgrounds
    'tag_pending': '#ffc107',
    'tag_ready':   '#6ea8fe',
    'tag_success': '#75b798',
    'tag_skipped': '#ffc107',
    'tag_error':   '#f87171',
}


# ── pdf_filler IPC helpers ────────────────────────────────────────────────────
# These functions are called by the background watcher thread when an Excel
# workbook drops a pdf_request.json file.  Running the extraction inside the
# already-trusted Lastrada process avoids Sophos/AV blocking a child exe that
# was spawned from a VBA macro.

def _abbreviate_material(name: str) -> str:
    """Return a short label (≤4 chars) for an aggregate material name."""
    n = name.strip()
    if not n:
        return ""
    m = re.match(r'(?:Natural\s+)?Gravel\s+0*(\d+)', n, re.IGNORECASE)
    if m:
        return ("G" + m.group(1))[:4]
    if re.search(r'limestone\s+sand',    n, re.IGNORECASE): return "LSS"
    if re.search(r'natural\s+sand',      n, re.IGNORECASE): return "NS"
    if re.search(r'baghouse',            n, re.IGNORECASE): return "BHF"
    if re.search(r'crushed\s+limestone', n, re.IGNORECASE): return "CLS"
    if re.search(r'limestone',           n, re.IGNORECASE): return "LS"
    if re.search(r'stone\s+sand',        n, re.IGNORECASE): return "StS"
    if re.search(r'screening',           n, re.IGNORECASE): return "SCR"
    if re.search(r'crushed\s+stone',     n, re.IGNORECASE): return "CS"
    if re.search(r'manufactured\s+sand', n, re.IGNORECASE): return "MfS"
    if re.search(r'slag',                n, re.IGNORECASE): return "SLG"
    return "".join(w[0].upper() for w in n.split() if w)[:4]


def _pdf_filler_extract(pdf_path: str) -> dict:
    """Extract all Marshall Mix Design fields from *pdf_path*.

    Combines the logic from pdf_filler.py's three functions
    (extract_fields, parse_page3_materials, parse_odot_spec_band)
    into a single call.  Returns a flat dict ready to JSON-encode.
    """
    try:
        import pdfplumber
    except ImportError:
        return {"error": "pdfplumber is not available in this build"}

    result = {}

    # ── Page-1 / page-2 scalar fields ─────────────────────────────────────
    try:
        with pdfplumber.open(pdf_path) as pdf:
            p1 = pdf.pages[0].extract_text() or ""
            p2 = pdf.pages[1].extract_text() if len(pdf.pages) > 1 else ""
            all_text = "\n".join((pg.extract_text() or "") for pg in pdf.pages)

        producer = plant1 = mix_type = ""
        binder_supplier = binder_grade = ""
        jmf_number = calib_number = ""
        virgin_binder = rap_binder = ""

        m = re.search(r"Mix Producer Name:\s*(.+)", p1)
        if m: producer = m.group(1).strip()

        m = re.search(r"Plant 1 Name:\s*(.+)", p1)
        if m: plant1 = m.group(1).strip()

        m = re.search(r"Mix Type\s+(.+?)(?:\s{2,}|\n|$)", p1)
        if m: mix_type = m.group(1).strip()

        m = re.search(r"Selected Virgin Binder Grade\s+(.+?)\s+(?:Neat\b|Producer Name)", p2)
        if m:
            binder_grade = m.group(1).strip()
        else:
            m = re.search(r"Selected Virgin Binder Grade\s+(.+)", p2)
            if m: binder_grade = m.group(1).strip()

        m = re.search(r"Binder Supplier\s+(.+?)\s+Brand Name", p2)
        if m:
            binder_supplier = m.group(1).strip()
        else:
            m = re.search(r"Binder Supplier\s+(.+)", p2)
            if m: binder_supplier = m.group(1).strip()

        m = re.search(r'%\s*Virgin\s*Binder\s+([\d.]+)',      p2, re.IGNORECASE)
        if m: virgin_binder = m.group(1).strip()

        m = re.search(r'%\s*Binder\s+from\s+RAP\s+([\d.]+)', p2, re.IGNORECASE)
        if m: rap_binder = m.group(1).strip()

        m = re.search(r"(B\d+)\s*~",    all_text)
        if m: jmf_number = m.group(1).strip()

        m = re.search(r"Calib#\s*(\d+)", all_text)
        if m: calib_number = m.group(1).strip()

        result.update({
            "producer": producer, "plant1": plant1, "mix_type": mix_type,
            "binder_supplier": binder_supplier, "binder_grade": binder_grade,
            "jmf_number": jmf_number, "calib_number": calib_number,
            "virgin_binder": virgin_binder, "rap_binder": rap_binder,
        })

    except Exception as e:
        return {"error": str(e)}

    # ── Page-3 aggregate / RAP / AC fields ────────────────────────────────
    try:
        from openpyxl.utils import get_column_letter as _gcl

        with pdfplumber.open(pdf_path) as pdf:
            p3 = pdf.pages[2].extract_text() or ""
            p2 = pdf.pages[1].extract_text() or ""

        lines = p3.split('\n')

        def _sec(name):
            for i, ln in enumerate(lines):
                if name in ln: return i
            return None

        coarse_i = _sec('Coarse Aggregates')
        fine_i   = _sec('Fine Aggregates')
        bag_i    = _sec('Baghouse Fines')
        rap_i    = _sec('RAP')
        blend_i  = _sec('Blend Gsb')

        def _agg(line):
            t = line.strip().split()
            if len(t) < 8: return None
            try:
                pct = float(t[0]); gsb = float(t[-1]); float(t[-2])
            except ValueError:
                return None
            if pct == 0.0: return None
            return {"material": t[-5]+' '+t[-4]+' '+t[-3],
                    "producer": ' '.join(t[2:-5]), "pct": pct, "gsb": gsb}

        def _bag(line):
            t = line.strip().split()
            if len(t) < 5: return None
            try: pct = float(t[0]); gsb = float(t[-1])
            except ValueError: return None
            if pct == 0.0: return None
            return {"material": t[2]+' '+t[3],
                    "producer": ' '.join(t[4:-1]), "pct": pct, "gsb": gsb}

        def _rap(line):
            mx = re.search(r'Method \d+\s+(.+?)\s+[A-Z]+/[A-Z]+\s+(\d+\.\d+)', line)
            if mx:
                t = line.strip().split()
                try: pct = float(t[0])
                except: pct = 0.0
                return {"pile": mx.group(1).strip().replace('"', 'in.'),
                        "pct": pct, "gse": float(mx.group(2))}
            return None

        coarse = []
        if coarse_i is not None and fine_i is not None:
            for ln in lines[coarse_i+1:fine_i]:
                r = _agg(ln)
                if r: r["item"] = "703.50"; coarse.append(r)

        fine = []
        if fine_i is not None and bag_i is not None:
            for ln in lines[fine_i+1:bag_i]:
                r = _agg(ln)
                if r: r["item"] = "703.05"; fine.append(r)

        bags = []
        if bag_i is not None and rap_i is not None:
            for ln in lines[bag_i+1:rap_i]:
                r = _bag(ln)
                if r: r["item"] = "703.05"; bags.append(r)

        rap_data = {"pile": "", "pct": 0.0, "gse": 0.0}
        if rap_i is not None:
            end = blend_i if blend_i else len(lines)
            for ln in lines[rap_i+1:end]:
                r = _rap(ln)
                if r: rap_data = r; break

        BASE_COL   = 6
        fine_start = BASE_COL + len(coarse) + 1
        bag_start  = fine_start + len(fine)

        aggs = []
        for i, item in enumerate(coarse): aggs.append({**item, "col": _gcl(BASE_COL + i)})
        for i, item in enumerate(fine):   aggs.append({**item, "col": _gcl(fine_start + i)})
        for i, item in enumerate(bags):   aggs.append({**item, "col": _gcl(bag_start  + i)})

        empty = {"material": "", "producer": "", "item": "", "pct": "", "gsb": ""}
        slots = list(coarse[:4]) + [empty] + list(fine[:4]) + list(bags)
        while len(slots) < 6: slots.append(empty)
        slots = slots[:6]

        ac_pct = binder_gb = ""
        mx = re.search(r'% Binder Content.*?Opt.*?Air Voids\s+([\d.]+)', p2, re.IGNORECASE)
        if mx: ac_pct = mx.group(1)
        mx = re.search(r'Binder Gb\s+([\d.]+)', p2, re.IGNORECASE)
        if mx: binder_gb = mx.group(1)

        mats = {
            "rap_pile": rap_data["pile"],  "rap_pct":  str(rap_data["pct"]),
            "rap_gse":  str(rap_data["gse"]), "ac_pct":   ac_pct,
            "binder_gb": binder_gb,        "agg_count": str(len(aggs)),
        }
        for i, s in enumerate(slots, 1):
            mats[f"material_{i}"] = s["material"]
            mats[f"producer_{i}"] = s["producer"]
            mats[f"item_{i}"]     = s.get("item", "")
        for i, a in enumerate(aggs, 1):
            mats[f"agg_{i}_col"]    = a["col"]
            mats[f"agg_{i}_pct"]    = str(a["pct"])
            mats[f"agg_{i}_gsb"]    = str(a["gsb"])
            mats[f"agg_{i}_abbrev"] = _abbreviate_material(a["material"])

        result.update(mats)

    except Exception as e:
        result["materials_error"] = str(e)

    # ── ODOT sieve band ────────────────────────────────────────────────────
    try:
        SIEVES = [
            (0,  r'2"\s*\(50'),      (1,  r'1-1/2"\s*\(38'),
            (2,  r'1"\s*\(25'),      (3,  r'3/4"\s*\(19\)'),
            (4,  r'1/2"\s*\(12'),    (5,  r'3/8"\s*\(9'),
            (6,  r'#4\s*\(4'),       (7,  r'#8\s*\(2'),
            (8,  r'#16\s*\(1\.1'),   (9,  r'#30\s*\(0\.6\)'),
            (10, r'#50\s*\(0\.3\)'), (11, r'#100\s*\(0\.1'),
            (12, r'#200\s*\(0\.0'),
        ]
        sieve_data = {}
        for i in range(13):
            sieve_data[f"sieve_{i}_jmf"] = ""
            sieve_data[f"sieve_{i}_mr"]  = ""

        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) >= 2:
                p2_lines = (pdf.pages[1].extract_text() or "").split('\n')
                for (idx, pattern) in SIEVES:
                    for line in p2_lines:
                        if re.search(pattern, line):
                            mx = re.search(r'\)', line)
                            if mx:
                                nums = re.findall(r'\d+\.?\d*', line[mx.end():])
                                if nums:
                                    sieve_data[f"sieve_{idx}_jmf"] = nums[0]
                                    if len(nums) >= 3:
                                        sieve_data[f"sieve_{idx}_mr"] = f"{nums[1]} / {nums[2]}"
                            break

        result.update(sieve_data)

    except Exception as e:
        result["sieve_error"] = str(e)

    return result


class CollapsibleSection(ttk.Frame):
    """A section that shows only a header button when collapsed, full content when expanded."""

    def __init__(self, parent, title, expanded=False, **kwargs):
        super().__init__(parent, **kwargs)
        self._title = title
        self._expanded = expanded

        self._btn = ttk.Button(self, text=self._label(), command=self._toggle)
        self._btn.pack(fill=tk.X)

        self.interior = ttk.Frame(self, padding=(0, 6, 0, 4))
        if self._expanded:
            self.interior.pack(fill=tk.X)

    def _label(self):
        symbol = "-" if self._expanded else "+"
        return f"{symbol}  {self._title}"

    def _toggle(self):
        self._expanded = not self._expanded
        self._btn.configure(text=self._label())
        if self._expanded:
            self.interior.pack(fill=tk.X)
        else:
            self.interior.pack_forget()


class FPCProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lastrada Report Processor")
        self.root.geometry("900x700")
        self.root.minsize(700, 600)
        self.root.resizable(True, True)
        
        # Variables
        self.source_folder = tk.StringVar()
        self.dest_folder = tk.StringVar()
        self.selected_lab = tk.StringVar()
        self.files_to_process = []
        self.success_count = 0
        self.error_count = 0
        self.skipped_count = 0
        self.dark_mode     = False
        self.dark_mode_var = tk.BooleanVar(value=False)   # kept in sync; used by menu checkbutton

        # Startup preference flags
        self.start_in_overwatch = tk.BooleanVar(value=False)
        self.run_on_startup     = tk.BooleanVar(value=False)

        # Overwatch state
        self.overwatch_mode = False
        self._overwatch_stop  = threading.Event()
        self._overwatch_thread = None
        self._overwatch_done            = set()   # source filenames successfully handled
        self._overwatch_notified_errors = set()   # filenames already error-notified (avoid repeat spam)
        self._tray_icon                 = None

        # pdf_filler IPC watcher state
        self._pdf_filler_stop = threading.Event()

        # Configure style
        self.setup_styles()

        # Create UI
        self.create_widgets()

        # Create log window (hidden by default)
        self.create_log_window()

        # Load saved settings (must be after widgets are created)
        self.load_config()

        # Hide to tray (not quit) when X is clicked while Overwatch is running
        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)

        # Auto-start Overwatch if the user has that preference set
        if self.start_in_overwatch.get():
            # Delay slightly so the window is fully rendered before we hide it
            self.root.after(1200, self._auto_start_overwatch)

        # Extract bundled pdf_filler.exe to AppData (kept for backwards compatibility)
        self._extract_pdf_filler()

        # Clean up any stale update artifacts left by a previous update attempt
        self._cleanup_update_artifacts()

        # Start the pdf_filler IPC watcher so Excel workbooks can request extractions
        # without spawning a new process (avoids Sophos/AV blocking VBA → exe calls)
        self._start_pdf_filler_watcher()

        # Check GitHub for a newer release (runs in background, never blocks startup)
        self._check_for_updates()
        
    def setup_styles(self):
        """Configure ttk styles for better appearance"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure colors
        style.configure('Section.TLabel',
                       font=('Segoe UI', 10, 'bold'),
                       foreground='#333')
        style.configure('Primary.TButton',
                       font=('Segoe UI', 9, 'bold'),
                       padding=10)
        style.configure('Success.TButton',
                       font=('Segoe UI', 10, 'bold'),
                       padding=8)
        style.configure('Toolbar.TFrame', background='#e8e8e8')
        
    def create_widgets(self):
        """Create all UI widgets"""

        # ── Toolbar — a themed ttk frame so it respects dark mode ──────────
        # (tk.Menu attached to root ignores background on Windows; ttk widgets don't)
        self._toolbar = ttk.Frame(self.root, style='Toolbar.TFrame')
        self._toolbar.pack(fill=tk.X, padx=0, pady=0)

        # Settings dropdown — attached to a plain ttk.Button that posts the menu on click
        settings_menu = tk.Menu(self._toolbar, tearoff=0)
        self._settings_menu = settings_menu

        settings_menu.add_checkbutton(
            label="Dark Mode",
            variable=self.dark_mode_var,
            command=self.toggle_dark_mode,
        )
        settings_menu.add_separator()
        settings_menu.add_checkbutton(
            label="Start in Overwatch mode automatically",
            variable=self.start_in_overwatch,
            command=self._on_startup_prefs_changed,
        )
        settings_menu.add_checkbutton(
            label="Run on Windows startup",
            variable=self.run_on_startup,
            command=self._on_startup_prefs_changed,
        )
        settings_menu.add_separator()
        settings_menu.add_command(label="Check for Updates", command=self.check_for_updates_manual)
        settings_menu.add_separator()
        settings_menu.add_command(label="View Log", command=self.toggle_log_window)

        self._settings_btn = ttk.Button(
            self._toolbar,
            text="Settings",
            command=self._post_settings_menu,
        )
        self._settings_btn.pack(side=tk.LEFT, padx=4, pady=2)

        ttk.Button(
            self._toolbar,
            text="Open Source Folder",
            command=self.open_source_folder,
        ).pack(side=tk.LEFT, padx=(0, 4), pady=2)

        ttk.Button(
            self._toolbar,
            text="Open Destination Folder",
            command=self.open_destination_folder,
        ).pack(side=tk.LEFT, padx=(0, 4), pady=2)

        self.overwatch_btn = ttk.Button(
            self._toolbar,
            text="Start Overwatch",
            command=self.toggle_overwatch,
        )
        self.overwatch_btn.pack(side=tk.LEFT, padx=(0, 4), pady=2)

        # Version label — right-aligned in the toolbar
        self._version_lbl = ttk.Label(
            self._toolbar,
            text=f"v{VERSION}",
            style='Toolbar.TLabel',
        )
        self._version_lbl.pack(side=tk.RIGHT, padx=(0, 8), pady=2)

        # ── Status bar — packed BOTTOM before content so it reserves space ──
        self.status_bar = ttk.Label(
            self.root, text="Ready",
            relief=tk.SUNKEN, anchor=tk.W,
            padding=(6, 2), font=('Segoe UI', 8),
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # ── Main content ─────────────────────────────────────────────────
        content = ttk.Frame(self.root, padding="12")
        content.pack(fill=tk.BOTH, expand=True)

        # Lab selector
        lab_frame = ttk.LabelFrame(content, text="Select Lab", padding="10")
        lab_frame.pack(fill=tk.X, pady=(0, 8))

        self.lab_combo = ttk.Combobox(
            lab_frame, textvariable=self.selected_lab,
            values=LABS, state='readonly', width=20,
        )
        self.lab_combo.pack(side=tk.LEFT)
        self.lab_combo.bind('<<ComboboxSelected>>', self.on_lab_changed)

        self.lab_status_label = ttk.Label(lab_frame, text="", foreground='#888')
        self.lab_status_label.pack(side=tk.LEFT, padx=(12, 0))

        # Source folder - collapsed by default
        source_section = CollapsibleSection(content, "Source Folder")
        source_section.pack(fill=tk.X, pady=(0, 4))

        ttk.Entry(source_section.interior,
                  textvariable=self.source_folder,
                  width=70).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(source_section.interior,
                   text="Browse", command=self.select_source_folder,
                   style='Primary.TButton').pack(side=tk.LEFT)

        # Destination folder - collapsed by default
        dest_section = CollapsibleSection(content, "Destination Folder")
        dest_section.pack(fill=tk.X, pady=(0, 4))

        ttk.Entry(dest_section.interior,
                  textvariable=self.dest_folder,
                  width=70).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(dest_section.interior,
                   text="Browse", command=self.select_dest_folder,
                   style='Primary.TButton').pack(side=tk.LEFT)

        # Scan button — full width, aligns with entry fields above
        ttk.Button(content,
                   text="Scan & Preview Lastrada PDFs",
                   command=self.scan_folder).pack(fill=tk.X, pady=(0, 8))

        # File list
        list_frame = ttk.LabelFrame(content, text="Files to Process (0)", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        columns = ('selected', 'filename', 'new_name', 'status')
        self.file_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)

        self.file_tree.heading('selected', text='Process')
        self.file_tree.heading('filename', text='Original Filename')
        self.file_tree.heading('new_name', text='New Filename')
        self.file_tree.heading('status',   text='Status')

        self.file_tree.column('selected', width=60,  anchor=tk.CENTER, stretch=False)
        self.file_tree.column('filename', width=270)
        self.file_tree.column('new_name', width=320)
        self.file_tree.column('status',   width=100)

        self.file_tree.bind('<ButtonRelease-1>', self.on_tree_click)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_tree.tag_configure('pending', foreground='#856404')
        self.file_tree.tag_configure('ready',   foreground='#004085')
        self.file_tree.tag_configure('success', foreground='#155724')
        self.file_tree.tag_configure('skipped', foreground='#856404')
        self.file_tree.tag_configure('error',   foreground='#721c24')

        # Process buttons
        btn_frame = ttk.Frame(content)
        btn_frame.pack(fill=tk.X, pady=(0, 6))

        self.process_btn = ttk.Button(btn_frame,
                                      text="Process Files",
                                      command=self.process_files,
                                      state='disabled',
                                      style='Success.TButton')
        self.process_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))

        self.process_print_btn = ttk.Button(btn_frame,
                                            text="Process and Print",
                                            command=self.process_and_print,
                                            state='disabled',
                                            style='Success.TButton')
        self.process_print_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        # Print Selected
        self.print_selected_btn = ttk.Button(content,
                                             text="Print Selected",
                                             command=self.print_selected,
                                             state='disabled',
                                             style='Primary.TButton')
        self.print_selected_btn.pack(fill=tk.X, pady=(0, 6))

        # Stats (hidden until first processing run)
        self.stats_frame = ttk.Frame(content)
        stats_inner = ttk.Frame(self.stats_frame)
        stats_inner.pack(fill=tk.X)

        success_frame = ttk.Frame(stats_inner, relief=tk.RAISED, borderwidth=2)
        success_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        self.success_label = ttk.Label(success_frame, text="0",
                                       font=('Segoe UI', 24, 'bold'), foreground='#28a745')
        self.success_label.pack(pady=(10, 0))
        ttk.Label(success_frame, text="Processed", font=('Segoe UI', 9)).pack(pady=(0, 10))

        error_frame = ttk.Frame(stats_inner, relief=tk.RAISED, borderwidth=2)
        error_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        self.error_label = ttk.Label(error_frame, text="0",
                                     font=('Segoe UI', 24, 'bold'), foreground='#dc3545')
        self.error_label.pack(pady=(10, 0))
        ttk.Label(error_frame, text="Errors", font=('Segoe UI', 9)).pack(pady=(0, 10))

        skipped_frame = ttk.Frame(stats_inner, relief=tk.RAISED, borderwidth=2)
        skipped_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        self.skipped_label = ttk.Label(skipped_frame, text="0",
                                       font=('Segoe UI', 24, 'bold'), foreground='#856404')
        self.skipped_label.pack(pady=(10, 0))
        ttk.Label(skipped_frame, text="Skipped", font=('Segoe UI', 9)).pack(pady=(0, 10))

        total_frame = ttk.Frame(stats_inner, relief=tk.RAISED, borderwidth=2)
        total_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        self.total_label = ttk.Label(total_frame, text="0",
                                     font=('Segoe UI', 24, 'bold'), foreground='#007f3a')
        self.total_label.pack(pady=(10, 0))
        ttk.Label(total_frame, text="Total", font=('Segoe UI', 9)).pack(pady=(0, 10))

        self.list_frame_widget = list_frame
        
    def _post_settings_menu(self):
        """Drop the Settings menu directly below the Settings button."""
        btn = self._settings_btn
        x = btn.winfo_rootx()
        y = btn.winfo_rooty() + btn.winfo_height()
        self._settings_menu.post(x, y)

    def open_destination_folder(self):
        """Open the current destination folder in Windows Explorer."""
        dest = self.dest_folder.get()
        if not dest:
            messagebox.showwarning("No Destination", "No destination folder is selected.")
            return
        if not os.path.exists(dest):
            messagebox.showwarning("Folder Not Found",
                                   f"The destination folder does not exist yet:\n{dest}")
            return
        os.startfile(dest)

    def open_source_folder(self):
        """Open the current source folder in Windows Explorer."""
        source = self.source_folder.get()
        if not source:
            messagebox.showwarning("No Source", "No source folder is selected.")
            return
        if not os.path.exists(source):
            messagebox.showwarning("Folder Not Found",
                                   f"The source folder does not exist:\n{source}")
            return
        os.startfile(source)

    def select_source_folder(self):
        """Open folder dialog for source selection"""
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.source_folder.set(folder)
            self.add_log(f"Source folder selected: {folder}", "info")
            self.save_config()

    def select_dest_folder(self):
        """Open folder dialog for destination selection"""
        folder = filedialog.askdirectory(title="Select Destination Folder")
        if folder:
            self.dest_folder.set(folder)
            self.add_log(f"Destination folder selected: {folder}", "info")
            self.save_config()
    
    def on_lab_changed(self, event=None):
        """Auto-fill destination folder when a lab is selected"""
        lab = self.selected_lab.get()
        dest = LAB_DESTINATIONS.get(lab, "")
        if dest:
            self.dest_folder.set(dest)
            self.lab_status_label.configure(text="Destination folder set.", foreground='#007f3a')
        else:
            self.dest_folder.set("")
            self.lab_status_label.configure(text="No destination configured for this lab yet.", foreground='#cc6600')
        self.save_config()

    def toggle_dark_mode(self):
        """Switch between light and dark theme and save the preference.
        Called both by the menu checkbutton (dark_mode_var already updated)
        and by any other code path."""
        self.apply_theme(self.dark_mode_var.get())
        self.save_config()

    def apply_theme(self, dark):
        """Apply light or dark colour theme to all widgets"""
        self.dark_mode = dark
        t = DARK_THEME if dark else LIGHT_THEME
        s = ttk.Style()

        # Re-applying the base theme forces all existing widgets to re-render,
        # which is required for LabelFrames and Comboboxes to pick up colour changes
        s.theme_use('clam')

        # Derive button colours from the theme
        btn_bg     = '#4a4a4a' if dark else '#e0e0e0'
        btn_active = '#5a5a5a' if dark else '#d0d0d0'
        btn_disabled = '#3a3a3a' if dark else '#cccccc'

        # --- Frames and labels ---
        s.configure('TFrame',              background=t['bg'])
        s.configure('TLabelframe',         background=t['bg'])
        s.configure('TLabelframe.Label',   background=t['bg'], foreground=t['fg'])
        s.configure('TLabel',              background=t['bg'], foreground=t['fg'])
        s.configure('TCheckbutton',        background=t['bg'], foreground=t['fg'])

        # --- Entry ---
        s.configure('TEntry', fieldbackground=t['entry_bg'], foreground=t['entry_fg'],
                               insertcolor=t['fg'])

        # --- Combobox (clam has two readonly states: focused and unfocused) ---
        s.configure('TCombobox', fieldbackground=t['entry_bg'], foreground=t['entry_fg'],
                                 selectbackground=t['select_bg'], selectforeground='white',
                                 background=btn_bg)
        s.map('TCombobox',
              fieldbackground=[
                  ('readonly', 'focus', t['entry_bg']),
                  ('readonly',          t['entry_bg']),
                  ('disabled',          t['bg']),
                  ('',                  t['entry_bg']),
              ],
              foreground=[
                  ('readonly', 'focus', t['entry_fg']),
                  ('readonly',          t['entry_fg']),
                  ('disabled',          '#888888'),
                  ('',                  t['entry_fg']),
              ],
              selectbackground=[('readonly', t['select_bg'])])
        # Style the internal dropdown listbox (uses tk option database)
        self.root.option_add('*TCombobox*Listbox.background',       t['entry_bg'])
        self.root.option_add('*TCombobox*Listbox.foreground',       t['entry_fg'])
        self.root.option_add('*TCombobox*Listbox.selectBackground', t['select_bg'])
        self.root.option_add('*TCombobox*Listbox.selectForeground', 'white')

        # --- Buttons (configure + map needed on Windows for all states) ---
        for style_name, font, padding in [
            ('TButton',        ('Segoe UI', 9),        5),
            ('Primary.TButton', ('Segoe UI', 9, 'bold'), 10),
            ('Success.TButton', ('Segoe UI', 10, 'bold'), 8),
        ]:
            s.configure(style_name, background=btn_bg, foreground=t['fg'],
                        font=font, padding=padding)
            s.map(style_name,
                  background=[('active', btn_active), ('disabled', btn_disabled)],
                  foreground=[('disabled', '#888888')])

        # --- Treeview (configure + map overrides Windows native row colours) ---
        s.configure('Treeview',
                    background=t['tree_bg'], fieldbackground=t['tree_bg'],
                    foreground=t['fg'], rowheight=22)
        s.configure('Treeview.Heading',
                    background=t['heading_bg'], foreground=t['fg'])
        s.map('Treeview.Heading',
              background=[('active', t['heading_bg'])],
              foreground=[('active', t['fg'])])
        s.map('Treeview',
              background=[('selected', t['select_bg']), ('!selected', t['tree_bg'])],
              foreground=[('selected', 'white')])

        # --- Treeview row tag colours ---
        if hasattr(self, 'file_tree'):
            self.file_tree.tag_configure('pending', foreground=t['tag_pending'])
            self.file_tree.tag_configure('ready',   foreground=t['tag_ready'])
            self.file_tree.tag_configure('success', foreground=t['tag_success'])
            self.file_tree.tag_configure('skipped', foreground=t['tag_skipped'])
            self.file_tree.tag_configure('error',   foreground=t['tag_error'])

        # --- Root window, status bar, and log window ---
        self.root.configure(bg=t['bg'])
        if hasattr(self, 'status_bar'):
            self.status_bar.configure(background=t['bg'], foreground=t['fg'])
        if hasattr(self, 'log_text'):
            self.log_text.configure(
                bg=t['log_bg'], fg=t['log_fg'], insertbackground=t['log_fg'])
        if hasattr(self, 'log_window'):
            self.log_window.configure(bg=t['bg'])

        # --- Toolbar and its dropdown menu ---
        toolbar_bg = '#3a3a3a' if dark else '#e8e8e8'
        s.configure('Toolbar.TFrame', background=toolbar_bg)
        s.configure('Toolbar.TLabel', background=toolbar_bg, foreground=t['fg'],
                    font=('Segoe UI', 8))
        if hasattr(self, '_settings_menu'):
            self._settings_menu.configure(
                background=t['entry_bg'], foreground=t['fg'],
                activebackground=t['select_bg'], activeforeground='white',
            )

        # --- Keep dark_mode_var in sync so the menu checkmark reflects reality ---
        if hasattr(self, 'dark_mode_var'):
            self.dark_mode_var.set(dark)

    def load_config(self):
        """Load saved settings from config file"""
        defaults = {
            'selected_lab':       '',
            'source_folder':      str(Path.home() / 'Downloads'),
            'dest_folder':        '',
            'dark_mode':          False,
            'start_in_overwatch': False,
            'run_on_startup':     False,
        }
        config = defaults.copy()

        if CONFIG_PATH.exists():
            try:
                with open(CONFIG_PATH, 'r') as f:
                    saved = json.load(f)
                config.update(saved)
            except Exception:
                pass  # Corrupted config - fall back to defaults

        # Apply source folder (default to Downloads if blank)
        source = config['source_folder'] or str(Path.home() / 'Downloads')
        self.source_folder.set(source)

        # Apply lab selection and trigger destination auto-fill
        if config['selected_lab'] in LABS:
            self.selected_lab.set(config['selected_lab'])
            self.on_lab_changed()

        # Only restore saved dest if lab didn't set one
        if not self.dest_folder.get() and config['dest_folder']:
            self.dest_folder.set(config['dest_folder'])

        # Apply saved theme
        self.apply_theme(config['dark_mode'])

        # Restore startup preferences (checkboxes only — overwatch auto-start happens in __init__)
        self.start_in_overwatch.set(config['start_in_overwatch'])
        self.run_on_startup.set(config['run_on_startup'])

    def save_config(self):
        """Save current settings to config file"""
        try:
            CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
            config = {
                'selected_lab':       self.selected_lab.get(),
                'source_folder':      self.source_folder.get(),
                'dest_folder':        self.dest_folder.get(),
                'dark_mode':          self.dark_mode,
                'start_in_overwatch': self.start_in_overwatch.get(),
                'run_on_startup':     self.run_on_startup.get(),
            }
            with open(CONFIG_PATH, 'w') as f:
                json.dump(config, f, indent=2)
        except Exception:
            pass  # Non-critical - silently skip if save fails

    def on_tree_click(self, event):
        """Toggle the selected state of a row when the Process column is clicked"""
        if self.file_tree.identify_region(event.x, event.y) != 'cell':
            return
        if self.file_tree.identify_column(event.x) != '#1':
            return
        item = self.file_tree.identify_row(event.y)
        if not item:
            return
        idx = self.file_tree.index(item)
        file_info = self.files_to_process[idx]
        file_info['selected'] = not file_info['selected']
        values = list(self.file_tree.item(item, 'values'))
        values[0] = '[x]' if file_info['selected'] else '[ ]'
        self.file_tree.item(item, values=values)

        # Enable Print Selected whenever at least one file is checked
        any_selected = any(f['selected'] for f in self.files_to_process)
        self.print_selected_btn.configure(state='normal' if any_selected else 'disabled')

    def scan_folder(self):
        """Scan source folder for Shelly & Sands Lastrada PDFs and preview new filenames"""
        source = self.source_folder.get()

        if not source:
            messagebox.showwarning("No Source Folder", "Please select a source folder first.")
            return

        if not os.path.exists(source):
            messagebox.showerror("Invalid Folder", "Source folder does not exist.")
            return

        # Clear previous results
        self.files_to_process = []
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        # Collect matching files (case-insensitive for both name and extension)
        pdf_files = [
            f for f in os.listdir(source)
            if f.lower().endswith(".pdf")
        ]
        pdf_count = len(pdf_files)

        self.list_frame_widget.configure(text=f"Files to Process ({pdf_count})")

        if pdf_count == 0:
            self._set_process_buttons('disabled')
            self.add_log("No PDF files found in source folder", "error")
            return

        self.add_log(f"Found {pdf_count} PDF file(s) - extracting previews...", "info")
        self._set_process_buttons('disabled')

        ready_count = 0
        for filename in pdf_files:
            filepath = os.path.join(source, filename)

            info = self.extract_info_from_pdf(filepath)

            new_name = ''
            status_text = 'Pending'
            tag = 'pending'

            pdf_type = info.get('pdf_type') if info else None

            if pdf_type == 'jmf_adjustment':
                required_ok = info and all([info['project'], info['jmf'], info['date']])
            else:
                required_ok = info and all([info['project'], info['material'], info['jmf'],
                                            info['production_day'], info['date']])

            if required_ok:
                if pdf_type == 'jmf_adjustment':
                    new_name = (
                        f"{info['project']} {info['jmf']} {info['date']} "
                        f"JMF Adjustment Letter.pdf"
                    )
                else:
                    new_name = (
                        f"{info['project']} {info['material']} {info['jmf']} "
                        f"{info['production_day']} {info['date']}.pdf"
                    )
                    new_name = new_name.replace('Intermediate', 'Int')
                    new_name = new_name.replace('Surface', 'Sur')
                new_name = re.sub(r'[<>:"|?*]', '_', new_name)
                status_text = 'Ready'
                tag = 'ready'
                ready_count += 1
            else:
                missing = []
                if not info:
                    missing.append("read error")
                else:
                    if not info['project']: missing.append("project")
                    if pdf_type != 'jmf_adjustment':
                        if not info['material']: missing.append("material")
                        if not info['production_day']: missing.append("production day")
                    if not info['jmf']: missing.append("JMF")
                    if not info['date']: missing.append("date")
                status_text = f"Missing: {', '.join(missing)}"
                tag = 'error'
                self.add_log(f"Preview failed: {filename} (missing: {', '.join(missing)})", "error")

            self.files_to_process.append({
                'path': filepath,
                'name': filename,
                'new_name': new_name,
                'info': info,
                'status': tag,
                'selected': True,
            })

            self.file_tree.insert('', tk.END, values=('[x]', filename, new_name, status_text), tags=(tag,))
            self.root.update()

        self.add_log(
            f"Preview complete: {ready_count} ready, {pdf_count - ready_count} with issues",
            "info"
        )

        if ready_count > 0:
            self._set_process_buttons('normal')
    
    def extract_info_from_pdf(self, filepath):
        """Extract metadata from a PDF file.

        Detects the document type from the first page and routes to the
        appropriate extraction logic.  Always returns a dict with at least
        a 'pdf_type' key, or None on read failure.

        Supported types:
          'jmf_adjustment' — Mar-Zane JMF Adjustment Letter
          'lastrada'       — Shelly & Sands Lastrada QC report (default)
        """
        try:
            reader = PdfReader(filepath)
            text = reader.pages[0].extract_text() or ""

            if 'JMF Adjustment Letter' in text:
                return self._extract_jmf_adjustment_info(text)
            elif 'Mar-Zane Lab' in text:
                return self._extract_pills_rice_info(text)
            else:
                return self._extract_lastrada_info(text)

        except Exception as e:
            self.add_log(f"Error reading PDF: {e}", "error")
            return None

    def _extract_jmf_adjustment_info(self, text):
        """Extract fields from a Mar-Zane JMF Adjustment Letter."""
        info = {
            'pdf_type':      'jmf_adjustment',
            'project':       '',
            'jmf':           '',
            'date':          '',
            # Unused by this type but kept so shared code doesn't KeyError
            'material':      '',
            'production_day': '',
        }

        # Project: 254-25
        m = re.search(r'Project:\s*(\S+)', text, re.IGNORECASE)
        if m:
            info['project'] = m.group(1).strip()

        # JMF = B240868
        m = re.search(r'JMF\s*=\s*(B\d+)', text, re.IGNORECASE)
        if m:
            info['jmf'] = m.group(1).strip()

        # Date: 04/21/2026
        m = re.search(r'Date:\s*(\d{1,2})/(\d{1,2})/(\d{4})', text, re.IGNORECASE)
        if m:
            info['date'] = f"{m.group(1).zfill(2)}-{m.group(2).zfill(2)}-{m.group(3)}"

        return info

    def _extract_pills_rice_info(self, text):
        """Extract fields from a Mar-Zane Lab Pills & Rice report (TE-220 / TE-221)."""
        info = {
            'pdf_type':      'pills_rice',
            'project':       '',
            'jmf':           '',
            'material':      '',
            'production_day': '',
            'date':          '',
        }

        # Project Number: 287-25
        m = re.search(r'Project\s+Number:\s*(\S+)', text, re.IGNORECASE)
        if m:
            info['project'] = m.group(1).strip()

        # JMF Number: B260208
        m = re.search(r'JMF\s+Number:\s*(B\d+)', text, re.IGNORECASE)
        if m:
            info['jmf'] = m.group(1).strip()

        # Mix Type: Type 2 Intermediate  (stop before the next labelled field)
        # Use \s* (zero or more spaces) because pypdf sometimes concatenates adjacent fields
        # without whitespace (e.g. "IntermediateTest Number:1").
        m = re.search(
            r'Mix\s+Type:\s*(.+?)(?=\s*Test\s+Number|\s*Sample\s+Type|\s*Day\s+Number)',
            text, re.IGNORECASE,
        )
        if m:
            info['material'] = ' '.join(m.group(1).split())

        # Day Number: 1
        m = re.search(r'Day\s+Number:\s*(\d+)', text, re.IGNORECASE)
        if m:
            info['production_day'] = f"Day{m.group(1).strip()}"

        # Date: 04/20/2026
        m = re.search(r'Date:\s*(\d{1,2})/(\d{1,2})/(\d{4})', text, re.IGNORECASE)
        if m:
            info['date'] = f"{m.group(1).zfill(2)}-{m.group(2).zfill(2)}-{m.group(3)}"

        return info

    def _extract_lastrada_info(self, text):
        """Extract fields from a Shelly & Sands Lastrada QC report."""
        info = {
            'pdf_type':      'lastrada',
            'project':       '',
            'material':      '',
            'jmf':           '',
            'production_day': '',
            'date':          '',
        }

        # Project: X-XX
        m = re.search(r'Project:\s*(\d+-\d+)', text, re.IGNORECASE)
        if m:
            info['project'] = m.group(1).strip()

        # Material: XXmm Intermediate/Surface  (stop before "Quantity(")
        m = re.search(r'Material:\s*([^\n]+?)(?:\s+Quantity\s*\(|$)', text, re.IGNORECASE)
        if m:
            info['material'] = ' '.join(m.group(1).split())

        # JMF: BXXXXXX
        m = re.search(r'JMF:\s*(B\d+)', text, re.IGNORECASE)
        if m:
            info['jmf'] = m.group(1).strip()

        # Production Day: X
        m = re.search(r'Production\s+Day:\s*(\d+)', text, re.IGNORECASE)
        if m:
            info['production_day'] = f"Day{m.group(1).strip()}"

        # Date Tested / Date Sampled: MM/DD/YYYY
        m = re.search(r'Date\s+(?:Tested|Sampled):\s*(\d{1,2})/(\d{1,2})/(\d{4})', text, re.IGNORECASE)
        if m:
            info['date'] = f"{m.group(1).zfill(2)}-{m.group(2).zfill(2)}-{m.group(3)}"

        return info
    
    def _is_already_merged(self, source_path: str, dest_path: Path) -> bool:
        """Return True if this exact source file has already been merged into dest_path.
        Uses source file size + destination path as a lightweight fingerprint."""
        try:
            if not MERGES_PATH.exists():
                return False
            with open(MERGES_PATH, 'r') as f:
                merges = json.load(f)
            source_size = os.path.getsize(source_path)
            dest_str = str(dest_path)
            return any(
                m.get('source_size') == source_size and m.get('dest') == dest_str
                for m in merges
            )
        except Exception:
            return False  # if the log can't be read, allow the merge

    def _record_merge(self, source_path: str, dest_path: Path):
        """Record a successful Pills & Rice merge so future scans can skip it."""
        try:
            merges = []
            if MERGES_PATH.exists():
                with open(MERGES_PATH, 'r') as f:
                    merges = json.load(f)
            merges.append({
                'source':      source_path,
                'source_size': os.path.getsize(source_path),
                'dest':        str(dest_path),
                'merged_at':   datetime.now().isoformat(),
            })
            MERGES_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(MERGES_PATH, 'w') as f:
                json.dump(merges, f, indent=2)
        except Exception:
            pass  # merge log is non-critical

    def _merge_pdfs(self, base_path: Path, append_path: str):
        """Append all pages from append_path to the end of base_path, overwriting base_path.
        Reads both files fully before writing so there is no conflict when paths overlap."""
        writer = PdfWriter()
        for page in PdfReader(str(base_path)).pages:
            writer.add_page(page)
        for page in PdfReader(str(append_path)).pages:
            writer.add_page(page)
        tmp = base_path.with_suffix('.mergetmp.pdf')
        with open(tmp, 'wb') as fh:
            writer.write(fh)
        tmp.replace(base_path)  # atomic rename overwrites the original

    def _ensure_project_subfolders(self, project_path: Path):
        """Create the standard subfolders inside a project folder if they don't exist yet."""
        for folder_name in ("Antistrip Reports", "Moistures", "Random Numbers"):
            (project_path / folder_name).mkdir(parents=True, exist_ok=True)

    def _set_process_buttons(self, state):
        """Enable or disable both process buttons together"""
        self.process_btn.configure(state=state)
        self.process_print_btn.configure(state=state)

    def process_files(self):
        """Process selected files without printing"""
        self._run_processing(print_after=False)

    def process_and_print(self):
        """Process selected files and send each to the default printer"""
        self._run_processing(print_after=True)

    def _get_default_printer(self):
        """Return the Windows default printer name, or None if it cannot be determined"""
        try:
            result = subprocess.run(
                ['powershell', '-NoProfile', '-command',
                 '(Get-WmiObject -Query "SELECT * FROM Win32_Printer WHERE Default=$true").Name'],
                capture_output=True, text=True, timeout=5,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            name = result.stdout.strip()
            return name if name else None
        except Exception:
            return None

    def print_selected(self):
        """Print all currently checked files without processing"""
        dest = self.dest_folder.get()
        printed = 0

        for file_info in self.files_to_process:
            if not file_info['selected']:
                continue

            print_path = None
            info = file_info.get('info')

            # For successfully processed or skipped files, print the renamed destination copy
            if file_info['status'] in ('success', 'skipped') and info and file_info['new_name']:
                pdf_type = info.get('pdf_type', 'lastrada')
                if pdf_type == 'jmf_adjustment':
                    dest_path = (
                        Path(dest) / info['project']
                        / 'JMF Adjustments'
                        / file_info['new_name']
                    )
                else:
                    material_short = (
                        info['material']
                        .replace('Intermediate', 'Int')
                        .replace('Surface', 'Sur')
                    )
                    dest_path = (
                        Path(dest) / info['project']
                        / f"{info['jmf']} {material_short}"
                        / file_info['new_name']
                    )
                if dest_path.exists():
                    print_path = dest_path

            # For everything else (ready, error) print the source file
            if print_path is None:
                source = Path(file_info['path'])
                if source.exists():
                    print_path = source

            if print_path:
                self._print_pdf(print_path)
                printed += 1
            else:
                self.add_log(f"Could not find file to print: {file_info['name']}", "error")

        if printed:
            self.add_log(f"Sent {printed} file(s) to printer", "info")

    def _print_pdf(self, filepath):
        """Send a PDF to the default printer using the best available method"""
        # Method 1: SumatraPDF - silent, reliable, works regardless of default PDF app
        sumatra_locations = [
            r'C:\Program Files\SumatraPDF\SumatraPDF.exe',
            r'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe',
            os.path.join(os.environ.get('LOCALAPPDATA', ''), 'SumatraPDF', 'SumatraPDF.exe'),
        ]
        for sumatra in sumatra_locations:
            if os.path.exists(sumatra):
                subprocess.Popen(
                    [sumatra, '-print-to-default', '-silent', str(filepath)],
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                self.add_log(f"Sent to printer: {filepath.name}", "info")
                return

        # Method 2: Adobe Reader/Acrobat with /t flag - prints directly to a named printer,
        # bypassing Adobe's own last-used printer setting which can default to Print to PDF
        printer_name = self._get_default_printer()
        adobe_locations = [
            r'C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe',
            r'C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe',
            r'C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe',
            r'C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe',
        ]
        if printer_name:
            for adobe in adobe_locations:
                if os.path.exists(adobe):
                    subprocess.Popen(
                        [adobe, '/t', str(filepath), printer_name],
                        creationflags=subprocess.CREATE_NO_WINDOW,
                    )
                    self.add_log(f"Sent to printer ({printer_name}): {filepath.name}", "info")
                    return

        # Method 3: Generic shell print verb - works with Foxit and other dedicated PDF apps
        try:
            os.startfile(str(filepath), 'print')
            self.add_log(f"Sent to printer: {filepath.name}", "info")
            return
        except OSError:
            pass

        # Method 4: No print verb registered (e.g. Edge only) - open for manual printing
        try:
            os.startfile(str(filepath))
            self.add_log(
                f"Auto-print unavailable on this machine - {filepath.name} opened for manual printing. "
                f"Install SumatraPDF for automatic printing.",
                "error"
            )
        except Exception as e:
            self.add_log(f"Could not open {filepath.name} for printing: {e}", "error")

    def _record_stat(self, file_info, info):
        """Append one success record to stats.json — non-critical, fails silently"""
        try:
            existing = []
            if STATS_PATH.exists():
                with open(STATS_PATH, 'r') as f:
                    existing = json.load(f)

            existing.append({
                'lab':               self.selected_lab.get(),
                'processed_at':      datetime.now().isoformat(),
                'project':           info.get('project', ''),
                'material':          info.get('material', ''),
                'jmf':               info.get('jmf', ''),
                'production_day':    info.get('production_day', ''),
                'date':              info.get('date', ''),
                'new_filename':      file_info['new_name'],
            })

            STATS_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(STATS_PATH, 'w') as f:
                json.dump(existing, f, indent=2)
        except Exception:
            pass  # Stats are non-critical

    def _run_processing(self, print_after=False):
        """Core processing logic shared by both process buttons"""
        dest = self.dest_folder.get()
        
        if not dest:
            messagebox.showwarning("No Destination", "Please select a destination folder first.")
            return
        
        if not os.path.exists(dest):
            messagebox.showerror("Invalid Folder", "Destination folder does not exist.")
            return
        
        # Reset counters
        self.success_count = 0
        self.error_count = 0
        self.skipped_count = 0
        
        # Show stats
        self.stats_frame.pack(fill=tk.X, pady=(0, 10))

        # Clear log
        self.log_text.delete(1.0, tk.END)
        
        # Disable process buttons during run
        self._set_process_buttons('disabled')
        
        # Process non-pills_rice files first so Lastrada reports are already at the destination
        # before any Pills & Rice merge attempt (handles the case where both arrive together).
        def _sort_key(i):
            pdf_type = (self.files_to_process[i].get('info') or {}).get('pdf_type', 'lastrada')
            return 1 if pdf_type == 'pills_rice' else 0

        process_order = sorted(range(len(self.files_to_process)), key=_sort_key)

        # Process each file using info already extracted during scan/preview
        for idx in process_order:
            file_info = self.files_to_process[idx]
            try:
                if not file_info['selected']:
                    # User deselected this file - leave it as-is in the list
                    pass
                elif file_info['status'] == 'ready' and file_info['new_name']:
                    info = file_info['info']
                    pdf_type = info.get('pdf_type') if info else 'lastrada'

                    project_path = Path(dest) / info['project']
                    if pdf_type == 'jmf_adjustment':
                        # Plant Testing / Project / JMF Adjustments /
                        dest_folder_path = project_path / 'JMF Adjustments'
                        dest_folder_path.mkdir(parents=True, exist_ok=True)
                        dest_path = dest_folder_path / file_info['new_name']
                        relative = Path(info['project']) / 'JMF Adjustments' / file_info['new_name']
                    else:
                        # Apply same abbreviations used in the filename
                        material_short = (
                            info['material']
                            .replace('Intermediate', 'Int')
                            .replace('Surface', 'Sur')
                        )
                        # Plant Testing / Project / JMF Material /
                        dest_folder_path = project_path / f"{info['jmf']} {material_short}"
                        dest_folder_path.mkdir(parents=True, exist_ok=True)
                        dest_path = dest_folder_path / file_info['new_name']
                        relative = Path(info['project']) / f"{info['jmf']} {material_short}" / file_info['new_name']
                        self._ensure_project_subfolders(dest_folder_path)

                    if dest_path.exists() and pdf_type == 'pills_rice':
                        if self._is_already_merged(file_info['path'], dest_path):
                            file_info['status'] = 'skipped'
                            self.skipped_count += 1
                            self.add_log(f"Skipped (already merged): {file_info['new_name']}", "info")
                        else:
                            # Merge: append pills & rice pages to the end of the existing report
                            self._merge_pdfs(dest_path, file_info['path'])
                            self._record_merge(file_info['path'], dest_path)
                            file_info['status'] = 'success'
                            self.success_count += 1
                            self.add_log(f"Merged: {file_info['name']} -> {relative}", "success")
                            self._record_stat(file_info, info)
                            if print_after:
                                self._print_pdf(dest_path)
                    elif dest_path.exists():
                        file_info['status'] = 'skipped'
                        self.skipped_count += 1
                        self.add_log(f"Skipped (already exists): {file_info['new_name']}", "info")
                    else:
                        if pdf_type == 'pills_rice':
                            # Never save Pills & Rice standalone — it must merge into an
                            # existing TE199 report.  If the TE199 hasn't been processed
                            # yet, leave this file alone so a future rescan can merge it.
                            file_info['status'] = 'error'
                            self.error_count += 1
                            self.add_log(
                                f"Skipped: {file_info['name']} — no matching TE199 report "
                                f"found at destination yet. Process the TE199 first, then rescan.",
                                "error",
                            )
                        else:
                            shutil.copy2(file_info['path'], str(dest_path))
                            file_info['status'] = 'success'
                            self.success_count += 1
                            self.add_log(f"Processed: {file_info['name']} -> {relative}", "success")
                            self._record_stat(file_info, info)
                            if print_after:
                                self._print_pdf(dest_path)
                elif file_info['status'] in ('success', 'skipped'):
                    pass  # already handled in a previous run - leave silently
                else:
                    # Missing fields flagged during preview - cannot process
                    file_info['status'] = 'error'
                    self.error_count += 1
                    self.add_log(f"Skipped: {file_info['name']} (missing fields - see preview)", "error")

            except Exception as e:
                file_info['status'] = 'error'
                self.error_count += 1
                self.add_log(f"Error processing {file_info['name']}: {str(e)}", "error")

            # Auto-deselect all files after processing
            file_info['selected'] = False

            # Update tree view
            tree_item = self.file_tree.get_children()[idx]
            tag = file_info['status']
            checkbox = '[x]' if file_info['selected'] else '[ ]'

            self.file_tree.item(tree_item,
                               values=(checkbox, file_info['name'], file_info['new_name'], file_info['status'].upper()),
                               tags=(tag,))

            # Update stats
            self.update_stats()
            self.root.update()

        # Final summary
        self.add_log(
            f"\nProcessing complete! {self.success_count} processed, "
            f"{self.skipped_count} skipped, {self.error_count} errors",
            "info"
        )

        if self.success_count > 0:
            messagebox.showinfo("Complete",
                              f"Processing complete!\n\n"
                              f"{self.success_count} file(s) processed successfully\n"
                              f"{self.skipped_count} file(s) skipped (already exist)\n"
                              f"{self.error_count} file(s) had errors\n\n"
                              f"Check the log for details.")

        # Re-enable process buttons so the user can process again without re-scanning
        if self.files_to_process:
            self._set_process_buttons('normal')

        # Print Selected stays disabled until user manually re-checks a file
        self.print_selected_btn.configure(state='disabled')
    
    # ── Startup preferences ───────────────────────────────────────────────

    def _on_startup_prefs_changed(self):
        """Called whenever either startup checkbox is toggled."""
        self._apply_run_on_startup(self.run_on_startup.get())
        self.save_config()

    def _apply_run_on_startup(self, enable: bool):
        """Add or remove the app from the Windows HKCU Run registry key."""
        app_name = "LastradaReportProcessor"
        reg_path = r"Software\Microsoft\Windows\CurrentVersion\Run"

        # Build the launch command — .exe when compiled, pythonw.exe for the script
        # (pythonw.exe suppresses the console window on startup)
        if getattr(sys, 'frozen', False):
            launch_cmd = f'"{sys.executable}"'
        else:
            pythonw = Path(sys.executable).with_name("pythonw.exe")
            script   = Path(__file__).resolve()
            launch_cmd = f'"{pythonw}" "{script}"'

        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, reg_path, 0,
                winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE,
            )
            if enable:
                winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, launch_cmd)
            else:
                try:
                    winreg.DeleteValue(key, app_name)
                except FileNotFoundError:
                    pass  # wasn't registered — nothing to remove
            winreg.CloseKey(key)
        except Exception as e:
            messagebox.showerror(
                "Startup Registry Error",
                f"Could not update the Windows startup registry:\n{e}",
            )
            # Revert the checkbox so it reflects the actual state
            self.run_on_startup.set(not enable)

    def _auto_start_overwatch(self):
        """Triggered on launch when 'Start in Overwatch mode' is set.
        Validates that the required settings are in place before activating."""
        if not self.source_folder.get() or not self.dest_folder.get() or not self.selected_lab.get():
            # Settings not configured yet — show the window normally and let the user set up
            return
        self._start_overwatch()

    # ── Overwatch mode ────────────────────────────────────────────────────

    def toggle_overwatch(self):
        """Start or stop Overwatch mode."""
        if self.overwatch_mode:
            self._stop_overwatch()
        else:
            if not self.source_folder.get():
                messagebox.showwarning("Overwatch", "Please select a source folder first.")
                return
            if not self.dest_folder.get():
                messagebox.showwarning("Overwatch", "Please select a destination folder first.")
                return
            if not self.selected_lab.get():
                messagebox.showwarning("Overwatch", "Please select a lab first.")
                return
            self._start_overwatch()

    def _start_overwatch(self):
        """Activate Overwatch: start the background scan thread and minimise to tray."""
        self.overwatch_mode = True
        self._overwatch_stop.clear()
        self._overwatch_notified_errors.clear()

        self.overwatch_btn.configure(text="Stop Overwatch")
        self.status_bar.configure(
            text=f"Overwatch active  —  scanning every {OVERWATCH_INTERVAL}s  —  processing automatically",
            foreground='#007f3a',
        )
        self.add_log(
            f"Overwatch mode started — auto-scanning '{self.source_folder.get()}' "
            f"every {OVERWATCH_INTERVAL}s",
            "info",
        )

        # Try to create a system tray icon.
        # Only hide the window if the tray icon launched successfully.
        if _pystray is not None:
            self._create_tray_icon()
            # Give pystray a moment to register the icon with Windows before hiding
            self.root.after(500, self.root.withdraw)
        else:
            # pystray not installed — minimise to taskbar so the user can still get back
            self.root.iconify()
            messagebox.showinfo(
                "Overwatch Active (no tray icon)",
                "Overwatch is running but the system tray icon is unavailable.\n\n"
                "Install pystray for full tray support:\n"
                "    pip install pystray\n\n"
                "The app is minimised to the taskbar.\n"
                "Click it in the taskbar to check status or stop Overwatch.",
            )

        # Launch the background worker
        self._overwatch_thread = threading.Thread(
            target=self._overwatch_worker, daemon=True
        )
        self._overwatch_thread.start()

    def _stop_overwatch(self):
        """Deactivate Overwatch: stop background thread and restore the window."""
        self.overwatch_mode = False
        self._overwatch_stop.set()

        # Remove tray icon
        if self._tray_icon:
            try:
                self._tray_icon.stop()
            except Exception:
                pass
            self._tray_icon = None

        # Restore main window
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

        self.overwatch_btn.configure(text="Start Overwatch")
        self.status_bar.configure(text="Overwatch stopped.", foreground='#888888')
        self.add_log("Overwatch mode stopped.", "info")
        self.save_config()

    def _overwatch_worker(self):
        """Background daemon thread.  Fires a scan on the main thread immediately,
        then every OVERWATCH_INTERVAL seconds until the stop event is set."""
        # Scan once right away
        self.root.after(0, self._overwatch_scan_and_process)
        # Then repeat until stopped
        while not self._overwatch_stop.wait(timeout=OVERWATCH_INTERVAL):
            if not self.overwatch_mode:
                break
            self.root.after(0, self._overwatch_scan_and_process)

    def _overwatch_scan_and_process(self):
        """Main-thread callback: scan the source folder and auto-process any new files.
        Does NOT print (printing is always a deliberate user action)."""
        if not self.overwatch_mode:
            return

        source = self.source_folder.get()
        dest   = self.dest_folder.get()

        if not source or not os.path.exists(source):
            return
        if not dest or not os.path.exists(dest):
            return

        try:
            pdf_files = [
                f for f in os.listdir(source)
                if f.lower().endswith(".pdf")
            ]
        except Exception:
            return

        new_files = [f for f in pdf_files if f not in self._overwatch_done]
        if not new_files:
            return

        self.add_log(f"[Overwatch] {len(new_files)} new file(s) found — processing...", "info")

        processed_names = []          # collect successes for the summary notification
        new_error_msgs  = {}          # filename -> short reason, for first-time error notifications

        # Extract info for all new files upfront, then sort so pills_rice is processed last.
        # This ensures the matching Lastrada report is already at the destination before
        # any merge attempt when both arrive in the same scan.
        file_queue = []
        for filename in new_files:
            filepath = os.path.join(source, filename)
            info = self.extract_info_from_pdf(filepath)
            file_queue.append((filename, filepath, info))

        file_queue.sort(
            key=lambda x: 1 if (x[2] or {}).get('pdf_type') == 'pills_rice' else 0
        )

        for filename, filepath, info in file_queue:
            pdf_type = info.get('pdf_type') if info else None

            # Validate required fields depending on PDF type
            if pdf_type == 'jmf_adjustment':
                fields_ok = info and all([info['project'], info['jmf'], info['date']])
            else:
                fields_ok = info and all([
                    info['project'], info['material'], info['jmf'],
                    info['production_day'], info['date'],
                ])

            if not fields_ok:
                missing = []
                if not info:
                    missing.append("read error")
                else:
                    if not info['project']:        missing.append("project")
                    if pdf_type != 'jmf_adjustment':
                        if not info['material']:       missing.append("material")
                        if not info['production_day']: missing.append("production day")
                    if not info['jmf']:            missing.append("JMF")
                    if not info['date']:           missing.append("date")
                self.add_log(
                    f"[Overwatch] Skipping {filename} — missing: {', '.join(missing)} "
                    f"(will retry next scan)",
                    "error",
                )
                # Do NOT add to _overwatch_done — retry next interval
                self._update_or_add_tree_row(filename, '', info, 'error', f"Missing: {', '.join(missing)}")
                if filename not in self._overwatch_notified_errors:
                    new_error_msgs[filename] = f"Missing: {', '.join(missing)}"
                continue

            # Build the new filename and destination path (same logic as manual processing)
            project_path = Path(dest) / info['project']
            if pdf_type == 'jmf_adjustment':
                new_name = (
                    f"{info['project']} {info['jmf']} {info['date']} "
                    f"JMF Adjustment Letter.pdf"
                )
                new_name = re.sub(r'[<>:"|?*]', '_', new_name)
                dest_folder_path = project_path / 'JMF Adjustments'
                relative = Path(info['project']) / 'JMF Adjustments' / new_name
            else:
                new_name = (
                    f"{info['project']} {info['material']} {info['jmf']} "
                    f"{info['production_day']} {info['date']}.pdf"
                )
                new_name = new_name.replace('Intermediate', 'Int')
                new_name = new_name.replace('Surface', 'Sur')
                new_name = re.sub(r'[<>:"|?*]', '_', new_name)
                material_short = (
                    info['material']
                    .replace('Intermediate', 'Int')
                    .replace('Surface', 'Sur')
                )
                dest_folder_path = project_path / f"{info['jmf']} {material_short}"
                relative = Path(info['project']) / f"{info['jmf']} {material_short}" / new_name

            dest_path = dest_folder_path / new_name

            if dest_path.exists() and pdf_type == 'pills_rice':
                if self._is_already_merged(filepath, dest_path):
                    # Already merged in a previous session — treat like a normal skip
                    self._overwatch_done.add(filename)
                    self.add_log(f"[Overwatch] Skipped (already merged): {new_name}", "info")
                    self._update_or_add_tree_row(filename, new_name, info, 'skipped', 'SKIPPED')
                else:
                    # Merge: append pills & rice pages to the end of the existing report
                    try:
                        self._merge_pdfs(dest_path, filepath)
                        self._record_merge(filepath, dest_path)
                        self._overwatch_done.add(filename)
                        self.add_log(f"[Overwatch] Merged: {filename}  ->  {relative}", "success")
                        self._update_or_add_tree_row(filename, new_name, info, 'success', 'SUCCESS')
                        self._record_stat({'new_name': new_name}, info)
                        processed_names.append(new_name)
                    except Exception as e:
                        self.add_log(f"[Overwatch] Error merging {filename}: {e}", "error")
                        self._update_or_add_tree_row(filename, new_name, info, 'error', 'ERROR')
                        if filename not in self._overwatch_notified_errors:
                            new_error_msgs[filename] = str(e)
                continue

            if dest_path.exists():
                self._overwatch_done.add(filename)
                self.add_log(f"[Overwatch] Skipped (already exists): {new_name}", "info")
                self._update_or_add_tree_row(filename, new_name, info, 'skipped', 'SKIPPED')
                continue

            if pdf_type == 'pills_rice':
                # Never save Pills & Rice standalone — it must merge into an existing
                # TE199 report.  Leave it out of _overwatch_done so it is retried on
                # every scan until the matching TE199 has been processed.
                self.add_log(
                    f"[Overwatch] {filename} — waiting for matching TE199 report, will retry next scan",
                    "info",
                )
                self._update_or_add_tree_row(filename, new_name, info, 'pending', 'Waiting for TE199')
                continue

            try:
                dest_folder_path.mkdir(parents=True, exist_ok=True)
                if pdf_type != 'jmf_adjustment':
                    self._ensure_project_subfolders(dest_folder_path)
                shutil.copy2(filepath, str(dest_path))
                self._overwatch_done.add(filename)
                self.add_log(f"[Overwatch] Processed: {filename}  ->  {relative}", "success")
                self._update_or_add_tree_row(filename, new_name, info, 'success', 'SUCCESS')
                self._record_stat({'new_name': new_name}, info)
                processed_names.append(new_name)
            except Exception as e:
                self.add_log(f"[Overwatch] Error processing {filename}: {e}", "error")
                self._update_or_add_tree_row(filename, new_name, info, 'error', 'ERROR')
                if filename not in self._overwatch_notified_errors:
                    new_error_msgs[filename] = str(e)

        # Refresh the section header count
        self.list_frame_widget.configure(text=f"Files to Process ({len(self.files_to_process)})")

        # Send a single Windows notification summarising what was processed
        if processed_names:
            if len(processed_names) == 1:
                self._tray_notify("Report Processed", processed_names[0])
            else:
                summary = "\n".join(processed_names[:5])
                if len(processed_names) > 5:
                    summary += f"\n…and {len(processed_names) - 5} more"
                self._tray_notify(f"{len(processed_names)} Reports Processed", summary)

        # Send one error notification per new failing file (never repeat for the same file)
        if new_error_msgs:
            self._overwatch_notified_errors.update(new_error_msgs.keys())
            if len(new_error_msgs) == 1:
                fname, reason = next(iter(new_error_msgs.items()))
                self._tray_notify("Processing Error", f"{fname}\n{reason}")
            else:
                lines = [f"{fn}: {reason}" for fn, reason in list(new_error_msgs.items())[:5]]
                if len(new_error_msgs) > 5:
                    lines.append(f"…and {len(new_error_msgs) - 5} more")
                self._tray_notify(f"{len(new_error_msgs)} Files Could Not Be Processed",
                                  "\n".join(lines))

    def _update_or_add_tree_row(self, filename, new_name, info, status_tag, status_text):
        """Update an existing row in the file tree or add a new one.
        Keeps self.files_to_process in sync so printing works correctly."""
        children = self.file_tree.get_children()

        for i, file_info in enumerate(self.files_to_process):
            if file_info['name'] == filename:
                # Update the data model
                file_info['new_name'] = new_name
                file_info['info']     = info
                file_info['status']   = status_tag
                # Update the tree row (preserve the checkbox state)
                if i < len(children):
                    checkbox = '[x]' if file_info['selected'] else '[ ]'
                    self.file_tree.item(
                        children[i],
                        values=(checkbox, filename, new_name, status_text),
                        tags=(status_tag,),
                    )
                return

        # Not in the list yet — insert a new row (selected by default)
        self.files_to_process.append({
            'path':     os.path.join(self.source_folder.get(), filename),
            'name':     filename,
            'new_name': new_name,
            'info':     info,
            'status':   status_tag,
            'selected': True,
        })
        self.file_tree.insert(
            '', tk.END,
            values=('[x]', filename, new_name, status_text),
            tags=(status_tag,),
        )
        # Enable Print Selected whenever files exist in the list
        self.print_selected_btn.configure(state='normal')

    # ── System tray helpers ───────────────────────────────────────────────

    def _create_tray_image(self):
        """Draw a small green 'eye' icon for the system tray using PIL."""
        size = 64
        img  = _PILImage.new('RGBA', (size, size), (0, 0, 0, 0))
        d    = _PILDraw.Draw(img)
        # Green circle background
        d.ellipse([2, 2, 62, 62], fill='#007f3a')
        # White eye outline
        d.ellipse([12, 22, 52, 42], fill='white')
        # Green iris
        d.ellipse([24, 26, 40, 38], fill='#007f3a')
        # Black pupil
        d.ellipse([29, 29, 35, 35], fill='black')
        return img

    def _create_tray_icon(self):
        """Build and run the pystray icon in a daemon thread.
        If pystray is not installed the app just stays hidden (accessible via taskbar)."""
        if _pystray is None:
            return

        image = self._create_tray_image()
        menu  = _pystray.Menu(
            _TrayItem(
                'Open Lastrada Report Processor',
                self._tray_open,
                default=True,
            ),
            _TrayItem('Stop Overwatch', self._tray_stop),
            _pystray.Menu.SEPARATOR,
            _TrayItem('Exit', self._tray_exit),
        )
        self._tray_icon = _pystray.Icon(
            'LastradaReportProcessor',
            image,
            'Lastrada Report Processor\nOverwatch Active',
            menu,
        )
        threading.Thread(target=self._tray_icon.run, daemon=True).start()

    def _tray_notify(self, title: str, message: str):
        """Send a Windows balloon notification from the tray icon.
        Silently does nothing if the tray icon is not running."""
        if self._tray_icon:
            try:
                self._tray_icon.notify(message, title)
            except Exception:
                pass  # notifications are non-critical

    def _tray_open(self, icon=None, item=None):
        """Restore main window from tray (called from tray thread)."""
        self.root.after(0, self._do_restore_window)

    def _do_restore_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _tray_stop(self, icon=None, item=None):
        """Stop Overwatch from tray menu (called from tray thread)."""
        self.root.after(0, self._stop_overwatch)

    def _tray_exit(self, icon=None, item=None):
        """Quit the whole application from tray menu (called from tray thread)."""
        self.root.after(0, self._quit_app)

    def _on_window_close(self):
        """X button handler.  While Overwatch is active, hide to tray instead of quitting."""
        if self.overwatch_mode:
            if _pystray is not None:
                self.root.withdraw()   # hide — tray icon is the only way back
            else:
                self.root.iconify()    # minimise to taskbar (no tray available)
        else:
            self._quit_app()

    def _quit_app(self):
        """Clean shutdown: stop overwatch / tray icon / pdf_filler watcher, then destroy the window."""
        self._overwatch_stop.set()
        self._pdf_filler_stop.set()
        if self._tray_icon:
            try:
                self._tray_icon.stop()
            except Exception:
                pass
        self.save_config()
        self.root.destroy()

    # ── Auto-update methods ────────────────────────────────────────────────

    def _cleanup_update_artifacts(self):
        """Delete any stale files left behind by a previous update attempt.

        Because the exe can be renamed to a new version on each update, we
        use directory globs rather than fixed paths so we find artifacts
        regardless of which version they came from.
        """
        if not getattr(sys, 'frozen', False):
            return  # dev mode — nothing to clean up
        try:
            exe_dir = Path(sys.executable).parent
            # Glob patterns cover any version name:
            #   *.exe.bak  — old exe renamed aside during swap (current mechanism)
            #   *.exe.old  — legacy v1.0.10 rename artifact
            #   *.new.exe  — legacy v1.0.9 temp download name
            for pattern in ("*.exe.bak", "*.exe.old", "*.new.exe"):
                for stale in exe_dir.glob(pattern):
                    try:
                        stale.unlink()
                    except Exception:
                        pass  # locked or already gone — skip silently
            # Fixed-name temp download file (current mechanism)
            tmp = exe_dir / "Lastrada_download.tmp"
            if tmp.exists():
                try:
                    tmp.unlink()
                except Exception:
                    pass
            # Legacy: .update.ps1 next to the exe (v1.0.8 mechanism)
            ps1 = Path(sys.executable).with_suffix(".update.ps1")
            if ps1.exists():
                try:
                    ps1.unlink()
                except Exception:
                    pass
        except Exception:
            pass

    def _extract_pdf_filler(self):
        """When running as a compiled .exe, extract the bundled pdf_filler.exe to
        %APPDATA%\\LastradaReportProcessor\\ so the Excel workbook can call it from
        a fixed, known path on any machine without needing Python installed."""
        if not getattr(sys, 'frozen', False):
            return  # dev mode — nothing to extract
        try:
            bundled = Path(sys._MEIPASS) / 'pdf_filler.exe'
            if not bundled.exists():
                return
            target = CONFIG_PATH.parent / 'pdf_filler.exe'
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(str(bundled), str(target))
        except Exception:
            pass  # non-critical — silently ignore if extraction fails

    # ── pdf_filler IPC watcher ─────────────────────────────────────────────

    def _start_pdf_filler_watcher(self):
        """Launch the background thread that serves pdf_filler requests from Excel."""
        self._pdf_filler_stop.clear()
        threading.Thread(target=self._pdf_filler_watch_loop, daemon=True).start()

    def _pdf_filler_watch_loop(self):
        """Daemon thread: poll every second for pdf_request.json.

        Performs the extraction IN-PROCESS via _pdf_filler_extract().  No
        child process is spawned, which sidesteps Sophos / AV behavioural
        rules that flag Office macros causing the parent app to spawn
        children.  pdfplumber is bundled in the main exe via PyInstaller.
        All activity is logged to the Lastrada log window.
        """
        while not self._pdf_filler_stop.wait(timeout=1.0):
            try:
                if not PDF_REQUEST_PATH.exists():
                    continue

                # Read and immediately delete the request file.  If we can't
                # parse it, log the actual exception so we know WHY (this is
                # invaluable for debugging IPC issues from VBA).
                try:
                    with open(PDF_REQUEST_PATH, 'r', encoding='utf-8') as fh:
                        request = json.load(fh)
                    PDF_REQUEST_PATH.unlink(missing_ok=True)
                except Exception as parse_err:
                    parse_err_msg = f'{type(parse_err).__name__}: {parse_err}'
                    self.root.after(0, lambda m=parse_err_msg: self.add_log(
                        f'[PDF Filler] IPC parse failed: {m}', 'error'))
                    try: PDF_REQUEST_PATH.unlink(missing_ok=True)
                    except Exception: pass
                    continue

                pdf_path = request.get('pdf_path', '').strip()
                self.root.after(0, lambda p=pdf_path: self.add_log(
                    f'[PDF Filler] Request received: {os.path.basename(p)}', 'info'))

                if not pdf_path or not os.path.exists(pdf_path):
                    err = f'PDF not found: {pdf_path}'
                    self._write_pdf_result({'error': err})
                    self.root.after(0, lambda m=err: self.add_log(f'[PDF Filler] {m}', 'error'))
                    continue

                # Extract IN-PROCESS - no subprocess spawned
                error_msg = None
                result    = None
                try:
                    result = _pdf_filler_extract(pdf_path)
                    if isinstance(result, dict) and 'error' in result and len(result) == 1:
                        error_msg = str(result['error'])
                        result = None
                    else:
                        self.root.after(0, lambda: self.add_log(
                            '[PDF Filler] Extraction complete', 'success'))
                except Exception as e:
                    error_msg = f'extraction failed: {type(e).__name__}: {e}'

                if error_msg:
                    self._write_pdf_result({'error': error_msg})
                    self.root.after(0, lambda m=error_msg: self.add_log(
                        f'[PDF Filler] Error: {m}', 'error'))
                else:
                    self._write_pdf_result(result)

            except Exception:
                pass  # never let an exception kill the watcher thread

    def _write_pdf_result(self, result: dict):
        """Write the extraction result dict to pdf_result.json."""
        try:
            PDF_RESULT_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(PDF_RESULT_PATH, 'w', encoding='utf-8') as fh:
                json.dump(result, fh, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _check_for_updates(self):
        """Startup auto-check: runs silently in background, only prompts if update found.
        Skipped when running as a .py script (only the compiled .exe auto-updates)."""
        if _requests is None:
            return
        if not getattr(sys, 'frozen', False):
            return  # dev mode — skip auto-check, use Settings > Check for Updates instead
        threading.Thread(
            target=self._update_check_worker, kwargs={'silent': True}, daemon=True
        ).start()

    def check_for_updates_manual(self):
        """Settings menu: manual check — works in both .exe and .py, shows result dialog."""
        if _requests is None:
            messagebox.showwarning(
                "Update Check Unavailable",
                "The 'requests' library is not installed.\n"
                "Run:  pip install requests"
            )
            return
        threading.Thread(
            target=self._update_check_worker, kwargs={'silent': False}, daemon=True
        ).start()

    def _update_check_worker(self, silent: bool = True):
        """Background thread: call the GitHub releases API and act on the result.

        silent=True  — startup auto-check: swallow all errors, only show dialog if update found.
        silent=False — manual check: always show a result dialog (update / up-to-date / error).
        """
        error_msg   = None
        latest_tag  = None
        asset_url   = None

        try:
            api_url = (
                f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
            )
            headers = {
                "Authorization": f"Bearer {GITHUB_TOKEN}",
                "Accept": "application/vnd.github+json",
            }
            resp = _requests.get(api_url, headers=headers, timeout=10)

            if resp.status_code == 401:
                error_msg = "Authentication failed (HTTP 401).\nThe update token may have expired or been revoked."
            elif resp.status_code == 404:
                error_msg = "No releases found on GitHub yet."
            elif resp.status_code != 200:
                error_msg = f"GitHub API returned HTTP {resp.status_code}."
            else:
                data       = resp.json()
                latest_tag = data.get("tag_name", "").lstrip("v")
                for asset in data.get("assets", []):
                    if asset["name"].lower().endswith(".exe"):
                        asset_url  = asset["url"]
                        asset_size = asset.get("size", 0)
                        break

        except Exception as e:
            error_msg = f"Network error: {e}"

        # Schedule UI feedback on the main thread
        if error_msg:
            if not silent:
                self.root.after(0, lambda: messagebox.showerror(
                    "Update Check Failed", error_msg))
            return

        if not self._version_is_newer(latest_tag, VERSION):
            if not silent:
                self.root.after(0, lambda: messagebox.showinfo(
                    "Up to Date",
                    f"You are running the latest version ({VERSION})."))
            return

        if not asset_url:
            if not silent:
                self.root.after(0, lambda: messagebox.showwarning(
                    "Update Available",
                    f"Version {latest_tag} is available but the installer\n"
                    f"has not been attached to the release yet."))
            return

        self.root.after(0, lambda: self._prompt_update(latest_tag, asset_url, asset_size))

    @staticmethod
    def _version_is_newer(latest: str, current: str) -> bool:
        """Return True if *latest* is strictly newer than *current*.
        Both strings are expected to be in 'MAJOR.MINOR.PATCH' format."""
        try:
            def to_tuple(v):
                return tuple(int(x) for x in v.strip().split("."))
            return to_tuple(latest) > to_tuple(current)
        except ValueError:
            return False

    def _prompt_update(self, latest_version: str, asset_url: str, asset_size: int = 0):
        """Show a dialog asking the user if they want to update now."""
        answer = messagebox.askyesno(
            "Update Available",
            f"A new version of Lastrada Report Processor is available!\n\n"
            f"  Your version:   {VERSION}\n"
            f"  New version:    {latest_version}\n\n"
            f"Download and install now?\n"
            f"(The app will restart automatically after the update.)",
            icon="info",
        )
        if answer:
            self._download_and_apply_update(asset_url, asset_size, latest_version)

    def _download_and_apply_update(self, asset_url: str, expected_size: int = 0,
                                     new_version: str = ""):
        """Download the new .exe and swap it in place, keeping the same filename.

        The update replaces the exe at its current path so shortcuts and
        pinned taskbar items keep working.  Versioned names (e.g.
        'Lastrada Report Processor v1.0.12.exe') only appear on GitHub
        release downloads — once installed, the file stays whatever name
        the user chose.

        Temp files:
          <exe_dir>\\Lastrada_download.tmp  — download in progress
          <current_exe_path>.bak            — backup of old exe during swap
        Both are cleaned up on the next startup if they survive a crash.
        """
        current_exe = Path(sys.executable)
        exe_dir     = current_exe.parent
        tmp_exe     = exe_dir / "Lastrada_download.tmp"

        # The new exe replaces the current one at exactly the same path.
        target_exe  = current_exe

        # Backup: just append ".bak" to the full filename (avoids any
        # pathlib suffix-validation edge cases with multi-dot extensions).
        bak_exe     = Path(str(current_exe) + ".bak")

        try:
            self.add_log("Downloading update...", "info")

            headers = {
                "Authorization": f"Bearer {GITHUB_TOKEN}",
                "Accept": "application/octet-stream",
            }
            resp = _requests.get(asset_url, headers=headers, timeout=60, stream=True)
            if resp.status_code != 200:
                messagebox.showerror("Update Failed",
                                     f"Could not download the update (HTTP {resp.status_code}).\n"
                                     f"Please try again later.")
                return

            with open(tmp_exe, "wb") as fh:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        fh.write(chunk)

            # Verify the download is complete before swapping
            if expected_size > 0:
                actual_size = tmp_exe.stat().st_size
                if actual_size != expected_size:
                    tmp_exe.unlink(missing_ok=True)
                    messagebox.showerror(
                        "Update Failed",
                        f"The download was incomplete ({actual_size:,} of {expected_size:,} bytes).\n"
                        f"Please check your connection and try again."
                    )
                    return

            self.add_log("Download complete. Preparing update...", "info")

            # Save settings NOW so dark mode and all preferences survive the restart.
            self.save_config()

            # ── Inline PowerShell swap (no .ps1 file written) ─────────────────
            #
            #  Strategy (safe in-place replacement):
            #  1. Sleep 5 s — lets PyInstaller finish releasing file locks
            #  2. Rename current exe → <same_name>.bak
            #     Windows allows renaming a running exe on NTFS; overwriting is blocked.
            #  3. Move Lastrada_download.tmp → original exe path (no conflict now)
            #  4a. SUCCESS: start the new exe, delete the .bak
            #  4b. FAILURE: rename .bak back to original so the OLD version is still
            #               launchable — user loses the update but doesn't lose the app
            #
            cur = str(current_exe).replace("'", "''")
            tmp = str(tmp_exe).replace("'", "''")
            tgt = str(target_exe).replace("'", "''")   # same as cur
            bak = str(bak_exe).replace("'", "''")

            ps_cmd = (
                # 5 s head-start so PyInstaller temp folder and AV get a moment to settle
                f"Start-Sleep -Seconds 5; "
                f"$ok = $false; "

                # ── Step 1: rename running exe aside ──────────────────────────
                # Windows lets you rename a running exe; overwriting is blocked.
                # Retry up to 5 × with 2 s gaps (handles brief AV locks).
                f"$renamed = $false; "
                f"for ($i=0; $i -lt 5; $i++) {{ "
                f"  try {{ Rename-Item -Force '{cur}' '{bak}'; $renamed=$true; break }} "
                f"  catch {{ Start-Sleep -Seconds 2 }} "
                f"}}; "
                f"if (-not $renamed) {{ exit 1 }}; "

                # ── Step 2: move download into place ──────────────────────────
                # Retry up to 15 × with 2 s gaps = up to 30 s total.
                # This outlasts a typical Sophos/Defender on-access scan of the
                # newly written file, which is the most common cause of failure.
                f"for ($i=0; $i -lt 15; $i++) {{ "
                f"  try {{ Move-Item -Force '{tmp}' '{tgt}'; $ok=$true; break }} "
                f"  catch {{ Start-Sleep -Seconds 2 }} "
                f"}}; "

                # ── Step 3a: success ──────────────────────────────────────────
                f"if ($ok) {{ "
                f"  Start-Process '{tgt}'; "
                f"  Start-Sleep -Seconds 3; "
                f"  Remove-Item '{bak}' -ErrorAction SilentlyContinue "
                f"}} else {{ "

                # ── Step 3b: failed — restore old exe so nothing is lost ──────
                f"  for ($i=0; $i -lt 5; $i++) {{ "
                f"    try {{ Rename-Item -Force '{bak}' '{tgt}'; break }} "
                f"    catch {{ Start-Sleep -Seconds 2 }} "
                f"  }}; "
                f"  Start-Process '{tgt}' "
                f"}}"
            )

            subprocess.Popen(
                [
                    "powershell.exe",
                    "-NoProfile", "-NonInteractive",
                    "-WindowStyle", "Hidden",
                    "-ExecutionPolicy", "Bypass",
                    "-Command", ps_cmd,
                ],
                creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
                close_fds=True,
            )

            self.root.destroy()

        except Exception as e:
            # Clean up partial download so a retry starts fresh
            tmp_exe.unlink(missing_ok=True)
            messagebox.showerror("Update Failed",
                                 f"An error occurred during the update:\n{e}\n\n"
                                 f"Please try again or update manually.")

    # ── Statistics display ─────────────────────────────────────────────────

    def update_stats(self):
        """Update statistics display"""
        self.success_label.configure(text=str(self.success_count))
        self.error_label.configure(text=str(self.error_count))
        self.skipped_label.configure(text=str(self.skipped_count))
        self.total_label.configure(text=str(len(self.files_to_process)))
    
    def create_log_window(self):
        """Create the log window, hidden by default"""
        self.log_window = tk.Toplevel(self.root)
        self.log_window.title("Processing Log")
        self.log_window.geometry("750x300")
        self.log_window.resizable(True, True)
        self.log_window.withdraw()  # hidden until user requests it

        # Closing the window hides it rather than destroying it
        self.log_window.protocol("WM_DELETE_WINDOW", self.log_window.withdraw)

        log_frame = ttk.Frame(self.log_window, padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, wrap=tk.WORD, font=('Consolas', 9))
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.log_text.tag_configure('success', foreground='#28a745')
        self.log_text.tag_configure('error', foreground='#dc3545')
        self.log_text.tag_configure('info', foreground='#17a2b8')

    def toggle_log_window(self):
        """Show the log window if hidden, bring it to front if already open"""
        if self.log_window.winfo_viewable():
            self.log_window.withdraw()
        else:
            self.log_window.deiconify()
            self.log_window.lift()

    def add_log(self, message, level='info'):
        """Add entry to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"

        self.log_text.insert(tk.END, log_message, level)
        self.log_text.see(tk.END)


def main():
    # ── Single-instance guard ─────────────────────────────────────────────
    # CreateMutexW returns a handle; GetLastError() == 183 (ERROR_ALREADY_EXISTS)
    # means another instance already owns the mutex.  We keep a reference to
    # the handle so Python's GC doesn't release it before the process exits.
    import ctypes
    _mutex_handle = ctypes.windll.kernel32.CreateMutexW(
        None, True, "Global\\LastradaReportProcessor_v1"
    )
    if ctypes.windll.kernel32.GetLastError() == 183:   # ERROR_ALREADY_EXISTS
        # Build a minimal Tk root just long enough to show the message box,
        # then exit — avoids the error sound of a bare messagebox on some Windows versions.
        _tmp = tk.Tk()
        _tmp.withdraw()
        messagebox.showwarning(
            "Already Running",
            "Lastrada Report Processor is already running.\n\n"
            "Check your system tray if you don't see the window.",
            parent=_tmp,
        )
        _tmp.destroy()
        sys.exit(0)

    root = tk.Tk()
    app = FPCProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
