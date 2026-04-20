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
from pypdf import PdfReader
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
VERSION = "1.0.1"

# How often (seconds) the Overwatch mode scans the source folder for new files
OVERWATCH_INTERVAL = 30

# GitHub auto-update settings (private repo, read-only token)
GITHUB_OWNER     = "zhuffman93"
GITHUB_REPO_NAME = "lastrada-report-processor"
GITHUB_TOKEN     = "[REDACTED-PAT]"

# Config file location - persists settings between sessions
CONFIG_PATH = Path(os.environ.get('APPDATA', Path.home())) / 'LastradaReportProcessor' / 'config.json'
STATS_PATH  = Path(os.environ.get('APPDATA', Path.home())) / 'LastradaReportProcessor' / 'stats.json'

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

            if info and all([info['project'], info['material'], info['jmf'],
                             info['production_day'], info['date']]):
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
                    if not info['material']: missing.append("material")
                    if not info['jmf']: missing.append("JMF")
                    if not info['production_day']: missing.append("production day")
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
        """Extract metadata from PDF file"""
        try:
            reader = PdfReader(filepath)
            
            # Get text from first page
            first_page = reader.pages[0]
            text = first_page.extract_text()
            
            info = {
                'project': '',
                'material': '',
                'jmf': '',
                'production_day': '',
                'date': ''
            }
            
            # Extract Project (looking for "Project: X-XX" pattern)
            project_match = re.search(r'Project:\s*(\d+-\d+)', text, re.IGNORECASE)
            if project_match:
                info['project'] = project_match.group(1).strip()
            
            # Extract Material (looking for "Material: XXmm Intermediate/Surface" but exclude "Quantity(Tons)")
            material_match = re.search(r'Material:\s*([^\n]+?)(?:\s+Quantity\s*\(|$)', text, re.IGNORECASE)
            if material_match:
                # Clean up the material text - remove extra whitespace
                info['material'] = ' '.join(material_match.group(1).split())
            
            # Extract JMF (looking for "JMF: BXXXXXX")
            jmf_match = re.search(r'JMF:\s*([B]\d+)', text, re.IGNORECASE)
            if jmf_match:
                info['jmf'] = jmf_match.group(1).strip()
            
            # Extract Production Day (looking for "Production Day: X")
            day_match = re.search(r'Production\s+Day:\s*(\d+)', text, re.IGNORECASE)
            if day_match:
                info['production_day'] = f"Day{day_match.group(1).strip()}"
            
            # Extract Date (looking for "Date Tested: XX/XX/XXXX" or "Date Sampled: XX/XX/XXXX")
            date_match = re.search(r'Date\s+(?:Tested|Sampled):\s*(\d{1,2})/(\d{1,2})/(\d{4})', text, re.IGNORECASE)
            if date_match:
                month = date_match.group(1).zfill(2)
                day = date_match.group(2).zfill(2)
                year = date_match.group(3)
                info['date'] = f"{month}-{day}-{year}"
            
            return info
            
        except Exception as e:
            self.add_log(f"Error extracting from PDF: {str(e)}", "error")
            return None
    
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
        
        # Process each file using info already extracted during scan/preview
        for idx, file_info in enumerate(self.files_to_process):
            try:
                if not file_info['selected']:
                    # User deselected this file - leave it as-is in the list
                    pass
                elif file_info['status'] == 'ready' and file_info['new_name']:
                    info = file_info['info']

                    # Apply same abbreviations used in the filename
                    material_short = (
                        info['material']
                        .replace('Intermediate', 'Int')
                        .replace('Surface', 'Sur')
                    )

                    # Build subfolder structure: Plant Testing / Project / JMF Material /
                    project_folder = Path(dest) / info['project']
                    mix_folder = project_folder / f"{info['jmf']} {material_short}"
                    mix_folder.mkdir(parents=True, exist_ok=True)

                    dest_path = mix_folder / file_info['new_name']

                    if dest_path.exists():
                        file_info['status'] = 'skipped'
                        self.skipped_count += 1
                        self.add_log(f"Skipped (already exists): {file_info['new_name']}", "info")
                    else:
                        shutil.copy2(file_info['path'], str(dest_path))
                        file_info['status'] = 'success'
                        self.success_count += 1
                        relative = Path(info['project']) / f"{info['jmf']} {material_short}" / file_info['new_name']
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

        for filename in new_files:
            filepath = os.path.join(source, filename)
            info = self.extract_info_from_pdf(filepath)

            if not info or not all([
                info['project'], info['material'], info['jmf'],
                info['production_day'], info['date'],
            ]):
                missing = []
                if not info:
                    missing.append("read error")
                else:
                    if not info['project']:        missing.append("project")
                    if not info['material']:       missing.append("material")
                    if not info['jmf']:            missing.append("JMF")
                    if not info['production_day']: missing.append("production day")
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

            # Build the new filename (same logic as manual processing)
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
            mix_folder = Path(dest) / info['project'] / f"{info['jmf']} {material_short}"
            dest_path  = mix_folder / new_name

            if dest_path.exists():
                self._overwatch_done.add(filename)
                self.add_log(f"[Overwatch] Skipped (already exists): {new_name}", "info")
                self._update_or_add_tree_row(filename, new_name, info, 'skipped', 'SKIPPED')
                continue

            try:
                mix_folder.mkdir(parents=True, exist_ok=True)
                shutil.copy2(filepath, str(dest_path))
                self._overwatch_done.add(filename)
                relative = Path(info['project']) / f"{info['jmf']} {material_short}" / new_name
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
        """Clean shutdown: stop overwatch / tray icon, then destroy the window."""
        self._overwatch_stop.set()
        if self._tray_icon:
            try:
                self._tray_icon.stop()
            except Exception:
                pass
        self.root.destroy()

    # ── Auto-update methods ────────────────────────────────────────────────

    def _check_for_updates(self):
        """Kick off a background thread to check GitHub for a newer release.
        The check is skipped silently if 'requests' is not installed or if
        the app is running as a plain .py script (not a compiled .exe)."""
        if _requests is None:
            return  # requests not installed - feature disabled
        # Only auto-update the compiled .exe; during dev the .py script is fine as-is
        if not getattr(sys, 'frozen', False):
            return
        t = threading.Thread(target=self._update_check_worker, daemon=True)
        t.start()

    def _update_check_worker(self):
        """Run in a background thread: hit the GitHub releases API and schedule
        a UI callback on the main thread if a newer version is found."""
        try:
            api_url = (
                f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO_NAME}/releases/latest"
            )
            headers = {
                "Authorization": f"Bearer {GITHUB_TOKEN}",
                "Accept": "application/vnd.github+json",
            }
            resp = _requests.get(api_url, headers=headers, timeout=10)
            if resp.status_code != 200:
                return  # no release yet, or token/network issue - fail silently

            data = resp.json()
            latest_tag = data.get("tag_name", "").lstrip("v")

            if not self._version_is_newer(latest_tag, VERSION):
                return  # already up to date

            # Find the .exe asset in the release
            asset_url = None
            for asset in data.get("assets", []):
                if asset["name"].lower().endswith(".exe"):
                    # For private repos we must use the API URL (not browser_download_url)
                    # and include the auth header when downloading
                    asset_url = asset["url"]
                    break

            if not asset_url:
                return  # release exists but no .exe attached yet

            # Schedule the update prompt on the main thread (safe from background threads)
            self.root.after(0, lambda: self._prompt_update(latest_tag, asset_url))

        except Exception:
            pass  # network unavailable, timeout, etc. - fail silently

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

    def _prompt_update(self, latest_version: str, asset_url: str):
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
            self._download_and_apply_update(asset_url)

    def _download_and_apply_update(self, asset_url: str):
        """Download the new .exe and swap it in via a batch script that runs
        after this process exits, then restart the application."""
        try:
            self.add_log("Downloading update...", "info")

            headers = {
                "Authorization": f"Bearer {GITHUB_TOKEN}",
                # This Accept header tells GitHub to redirect to the actual binary
                "Accept": "application/octet-stream",
            }
            resp = _requests.get(asset_url, headers=headers, timeout=60, stream=True)
            if resp.status_code != 200:
                messagebox.showerror("Update Failed",
                                     f"Could not download the update (HTTP {resp.status_code}).\n"
                                     f"Please try again later.")
                return

            # Save the downloaded .exe next to the current .exe with a temp name
            current_exe = Path(sys.executable)
            tmp_exe     = current_exe.with_suffix(".new.exe")

            with open(tmp_exe, "wb") as fh:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        fh.write(chunk)

            self.add_log("Download complete. Preparing update...", "info")

            # Write a small batch script that:
            #   1. Waits 2 seconds for our process to fully exit
            #   2. Replaces the current .exe with the downloaded one
            #   3. Restarts the application
            #   4. Deletes itself
            bat_path = current_exe.with_suffix(".update.bat")
            bat_content = (
                "@echo off\n"
                "timeout /t 2 /nobreak >nul\n"
                f'move /y "{tmp_exe}" "{current_exe}"\n'
                f'start "" "{current_exe}"\n'
                'del "%~f0"\n'
            )
            bat_path.write_text(bat_content)

            # Launch the batch script as a fully detached process, then quit
            subprocess.Popen(
                ["cmd.exe", "/c", str(bat_path)],
                creationflags=(
                    subprocess.CREATE_NEW_PROCESS_GROUP |
                    subprocess.DETACHED_PROCESS
                ),
                close_fds=True,
            )

            self.root.destroy()

        except Exception as e:
            messagebox.showerror("Update Failed",
                                 f"An error occurred during the update:\n{e}\n\n"
                                 f"Please update manually.")

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
    root = tk.Tk()
    app = FPCProcessorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
