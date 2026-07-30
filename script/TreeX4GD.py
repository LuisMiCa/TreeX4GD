import sys
import os
import time
import json
import re
import threading
import queue
import webbrowser
import platform
import pathlib
import shutil
import subprocess  # Added for macOS Dark Mode detection
from html import escape
import random
from datetime import datetime
from collections import defaultdict
from dataclasses import dataclass
import tkinter as tk
from tkinter import filedialog, scrolledtext

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# ---------------- Config ----------------
SCRIPT_NAME = "TreeX4GD"
SCRIPT_SUBTITLE = "File Tree Exporter for Google Drive"
SCRIPT_FULL_NAME = f"{SCRIPT_NAME} - {SCRIPT_SUBTITLE}"
SCRIPT_VERSION = "1.0.0"
EXPECTED_TEMPLATE_VERSION = "1.0.x"

# Scope for Google Drive API (read-only metadata access)
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly']

# Icons used in log
ICON_FOLDER = "📂"
ICON_FILE = "📄"

# Symbols for Excel export
TREE_SYMBOL_FOLDER = "■"
TREE_SYMBOL_FILE = "●"

SEPARATOR = " ‖ "

# Animated spinner characters
SPINNER = ["⠋", "⠙", "⠚", "⠞", "⠖", "⠦", "⠴", "⠲", "⠳", "⠓"]

# Configurable sleep interval BEFORE API calls (0.1s ensures max ~10 queries/sec)
SLEEP_INTERVAL = 0.1
MAX_API_RETRIES = 5  # Exponential backoff limit for 429/403 errors

# Long MIME type simplification map
MIME_TYPE_MAP = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "Word Document",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "Excel Spreadsheet",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "PowerPoint Presentation",
}

# Month name mapping (name → number) for natural sorting
MONTH_MAP = {
    # English
    'january': 1, 'february': 2, 'march': 3, 'april': 4,
    'may': 5, 'june': 6, 'july': 7, 'august': 8,
    'september': 9, 'october': 10, 'november': 11, 'december': 12,
    # Portuguese
    'janeiro': 1, 'fevereiro': 2, 'março': 3, 'marco': 3, 'abril': 4,
    'maio': 5, 'junho': 6, 'julho': 7, 'agosto': 8, 'setembro': 9,
    'outubro': 10, 'novembro': 11, 'dezembro': 12,
    # Spanish
    'enero': 1, 'febrero': 2, 'marzo': 3, 'abril': 4,
    'mayo': 5, 'junio': 6, 'julio': 7, 'agosto': 8,
    'septiembre': 9, 'setiembre': 9,
    'octubre': 10, 'noviembre': 11, 'diciembre': 12,
    # Italian
    'gennaio': 1, 'febbraio': 2, 'marzo': 3, 'aprile': 4,
    'maggio': 5, 'giugno': 6, 'luglio': 7, 'agosto': 8,
    'settembre': 9, 'ottobre': 10, 'novembre': 11, 'dicembre': 12,
    # French
    'janvier': 1, 'février': 2, 'fevrier': 2, 'mars': 3, 'avril': 4, 'mai': 5,
    'juin': 6, 'juillet': 7, 'août': 8, 'aout': 8,
    'septembre': 9, 'octobre': 10, 'novembre': 11, 'décembre': 12, 'decembre': 12,
    # German
    'januar': 1, 'februar': 2, 'märz': 3, 'maerz': 3, 'april': 4, 'mai': 5,
    'juni': 6, 'juli': 7, 'august': 8, 'september': 9, 'oktober': 10,
    'november': 11, 'dezember': 12
}


# ---------------- Path Handling ----------------
def get_base_path():
    """Get the absolute path to the script's directory (works for .py and frozen .exe)."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return base_path


# ---------- Data Structure Definition ----------
@dataclass
class DataRow:
    """Holds all data for a single file/folder item."""
    tree: str
    name: str
    type: str
    size: str         # Formatted string
    raw_size: int     # Raw size in bytes (integer)
    date: str         # Modified date
    created_date: str # Created date
    mime: str
    level: int
    parent: str
    id: str
    owner: str
    modified_by: str  # Full display name
    md5: str          # MD5 Checksum (files only)
    shared: bool      # Shared status
    version: int      # Version number


# ---------- Custom Centered MessageBox ----------
class CustomMessageBox(tk.Toplevel):
    """A custom, centered, modal message box."""

    def __init__(self, parent, title, message, dialog_type="ok", neg_text=None, pos_text=None):
        super().__init__(parent)
        self.withdraw()  # Hide initially to prevent visual jump
        self.transient(parent)
        self.grab_set()

        # Register this modal on the parent so background workers can close it.
        self.parent_ref = parent
        try:
            setattr(parent, "_active_modal", self)
        except Exception:
            pass

        self.title(title)
        self.message = message
        self.neg_text = neg_text
        self.pos_text = pos_text
        self.dialog_type = dialog_type
        self.result = None

        # --- Layout ---
        self.resizable(False, False)
        main_frame = tk.Frame(self, padx=20, pady=15)
        main_frame.pack(expand=True, fill="both")

        message_label = tk.Label(main_frame, text=self.message, wraplength=400, justify="left")
        message_label.pack(pady=(0, 20))

        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x")

        if self.dialog_type == "ok":
            ok_button = tk.Button(button_frame, text="OK", width=10, command=self._on_ok)
            ok_button.pack()
            self.bind("<Return>", self._on_ok)
            self.bind("<Escape>", self._on_ok)

        elif self.dialog_type == "negpos":
            button_frame.columnconfigure((0, 1), weight=1)
            neg_button = tk.Button(button_frame, text=self.neg_text, width=20, command=self._on_no)
            neg_button.pack(side="left", expand=True, padx=5)
            pos_button = tk.Button(button_frame, text=self.pos_text, width=20, command=self._on_yes)
            pos_button.pack(side="left", expand=True, padx=5)
            self.bind("<Escape>", self._on_no)
            self.bind("<Return>", self._on_yes)

        elif self.dialog_type == "yesnocancel":
            yes_button = tk.Button(button_frame, text="Yes", width=10, command=self._on_yes)
            yes_button.pack(side="left", expand=True, padx=5)
            no_button = tk.Button(button_frame, text="No", width=10, command=self._on_no)
            no_button.pack(side="left", expand=True, padx=5)
            cancel_button = tk.Button(button_frame, text="Cancel", width=10, command=self._on_cancel)
            cancel_button.pack(side="left", expand=True, padx=5)
            self.bind("<Return>", self._on_yes)
            self.bind("<Escape>", self._on_cancel)

        # --- Centering ---
        self.update_idletasks()

        # Get parent absolute screen coordinates
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_w = parent.winfo_width()
        parent_h = parent.winfo_height()

        # Get dialog dimensions (use reqwidth/reqheight if actual is not yet calculated)
        dialog_w = self.winfo_width()
        dialog_h = self.winfo_height()
        if dialog_w < 2:
            dialog_w = self.winfo_reqwidth()
        if dialog_h < 2:
            dialog_h = self.winfo_reqheight()

        # Calculate position relative to screen
        pos_x = parent_x + (parent_w // 2) - (dialog_w // 2)
        pos_y = parent_y + (parent_h // 2) - (dialog_h // 2)
        
        # Safety check to prevent negative coordinates
        if pos_x < 0: pos_x = 0
        if pos_y < 0: pos_y = 0

        self.geometry(f"+{pos_x}+{pos_y}")
        self.deiconify()  # Show the window at the correct position

        if self.dialog_type == "yesnocancel":
            self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        elif self.dialog_type == "negpos":
            self.protocol("WM_DELETE_WINDOW", self._on_no)
        else:
            self.protocol("WM_DELETE_WINDOW", self._on_ok)

        self.wait_window()

    def _on_ok(self, event=None):
        self.result = True
        self.destroy()

    def _on_yes(self, event=None):
        self.result = True
        self.destroy()

    def _on_no(self, event=None):
        self.result = False
        self.destroy()

    def _on_cancel(self, event=None):
        self.result = None
        self.destroy()

    def destroy(self):
        """Clear the parent's _active_modal reference if it points to this dialog, then destroy."""
        try:
            parent = getattr(self, 'parent_ref', None)
            if parent is not None:
                if getattr(parent, '_active_modal', None) is self:
                    setattr(parent, '_active_modal', None)
        except Exception:
            pass
        finally:
            super().destroy()


# Helper functions to mimic the standard messagebox API
def custom_showerror(parent, title, message):
    CustomMessageBox(parent, title, message, dialog_type="ok")


def custom_showwarning(parent, title, message):
    CustomMessageBox(parent, title, message, dialog_type="ok")


def custom_asknegpos(parent, title, message, neg_text, pos_text):
    dialog = CustomMessageBox(parent, title, message, dialog_type="negpos", neg_text=neg_text, pos_text=pos_text)
    return dialog.result


def custom_askyesnocancel(parent, title, message):
    dialog = CustomMessageBox(parent, title, message, dialog_type="yesnocancel")
    return dialog.result


# -------------------------
# --- Application Class ---
# -------------------------

class TreeX4GDApp:
    def __init__(self, root):
        self.root = root
        self._data_lock = threading.Lock()

        # ---------------- Theme Detection ----------------
        self.is_dark_mode = self._detect_macos_dark_mode()
        self.colors = self._get_theme_colors(self.is_dark_mode)

        # ---------------- Instance Attributes ----------------
        self.current_task_id = 0
        self.total_files = 0
        self.total_folders = 0
        self.data_rows = []              # List of DataRow objects
        self.log_queue = queue.Queue()
        self.root.after(100, self._process_log_queue)
        self.is_paused = False
        self.is_stopped = False
        self.should_auto_scroll = True
        self.app_state = "standby"
        self.spinner_index = 0
        self.spinner_task = None
        self.has_unsaved_data = False
        self.is_importing = False
        self.selected_folder_name = ""
        self.auth_email = ""
        self.data_account_email = ""
        self.is_live_scan_data = False
        self.folder_size_totals = defaultdict(int)
        self.parent_ids_by_child = {}
        self.g_service = None
        self.acquisition_start = None
        self._folder_entry_mouse_focus_pending = False

        # Constants for HTML Template Injection
        self.HTML_GLOBAL_PREFIX = f"{SCRIPT_NAME}_"

        # ---------------- GUI Setup ----------------
        self.root.title(f"{SCRIPT_NAME} v{SCRIPT_VERSION} — {SCRIPT_SUBTITLE}")
        self.root.geometry("900x600")

        # --- Window Icon setup ---
        try:
            base_dir = get_base_path()
            if sys.platform.startswith('win'):
                icon_path = os.path.join(base_dir, 'TreeX4GD.ico')
                if os.path.exists(icon_path):
                    self.root.iconbitmap(icon_path)
            elif sys.platform.startswith('linux'):
                icon_path = os.path.join(base_dir, 'TreeX4GD.png')
                if os.path.exists(icon_path):
                    icon_img = tk.PhotoImage(file=icon_path)
                    self.root.iconphoto(True, icon_img)
             # macOS dock icon is handled by the BUNDLE in the .spec file.
        except Exception as e:
            print(f"Warning: Could not load window icon: {e}")

        # --- Grid configuration for responsiveness ---
        self.root.rowconfigure(4, weight=1)
        self.root.columnconfigure(1, weight=1)
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)

        # --- Row 0: Credentials ---
        tk.Label(self.root, text="Credentials (OAuth 2.0):").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self.credentials_entry = tk.Entry(self.root)
        self.credentials_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        self.browse_btn = tk.Button(self.root, text="Select Credentials (JSON file)", command=self.on_browse_credentials)
        self.browse_btn.grid(row=0, column=2, columnspan=2, sticky="ew", padx=(6, 6), pady=6)

        # --- Row 1: Folder ID ---
        tk.Label(self.root, text="Folder ID or drive URL:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        self.folder_entry = tk.Entry(self.root)
        self.folder_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=6)
        self.folder_entry.unbind("<FocusIn>")
        self.folder_entry.bind("<Button-1>", self.on_folder_entry_mouse_down, "+")
        self.folder_entry.bind("<FocusIn>", self.on_folder_entry_focus_in, "+")
        self.folder_entry.bind("<FocusOut>", self.on_folder_entry_focus_out, "+")

        self.root_btn = tk.Button(self.root, text="Set Root", command=self.on_set_root)
        self.root_btn.grid(row=1, column=2, sticky="ew", padx=(6, 2), pady=6)
        self.open_drive_btn = tk.Button(self.root, text="Browse Root", command=self.on_open_drive)
        self.open_drive_btn.grid(row=1, column=3, sticky="ew", padx=(2, 6), pady=6)

        # --- Controls row ---
        tk.Label(self.root, text="Listing Controls:").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        controls_frame = tk.Frame(self.root)
        controls_frame.grid(row=2, column=1, columnspan=3, sticky="ew", padx=6, pady=6)
        controls_frame.columnconfigure(0, weight=1)
        controls_frame.columnconfigure(1, weight=1)
        controls_frame.columnconfigure(2, weight=0)

        self.start_btn = tk.Button(controls_frame, text="Start ▶", command=self.on_start_pause)
        self.start_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))
        self.stop_btn = tk.Button(controls_frame, text="Stop ⏹", command=self.on_stop, state="disabled")
        self.stop_btn.grid(row=0, column=1, sticky="ew", padx=(4, 8))
        self.auto_scroll_var = tk.BooleanVar(value=True)
        self.auto_scroll_chk = tk.Checkbutton(
            controls_frame,
            text="Auto Scroll",
            variable=self.auto_scroll_var,
            command=self.on_toggle_autoscroll
        )
        self.auto_scroll_chk.grid(row=0, column=2, sticky="e")

        # --- Log Box (Theme Aware) ---
        self.log_box = scrolledtext.ScrolledText(
            self.root, 
            wrap=tk.NONE, 
            font=("Courier New", 10),
            bg=self.colors['log_bg'],
            fg=self.colors['log_fg'],
            insertbackground=self.colors['log_fg'] # Cursor color
        )
        self.log_box.grid(row=4, column=0, columnspan=4, sticky="nsew", padx=10, pady=(4, 8))
        h_scroll = tk.Scrollbar(self.root, orient='horizontal', command=self.log_box.xview)
        h_scroll.grid(row=5, column=0, columnspan=4, sticky='ew', padx=10)
        self.log_box.config(xscrollcommand=h_scroll.set, state="disabled")

        # Tags configuration (Theme Aware)
        for tag, color in self.colors['tags'].items():
            self.log_box.tag_config(tag, foreground=color)

        export_frame = tk.Frame(self.root)
        export_frame.grid(row=6, column=0, columnspan=4, sticky="ew", padx=8, pady=(4, 4))
        for i in range(1, 5):
            export_frame.columnconfigure(i, weight=1)
        tk.Label(export_frame, text="Export:").grid(row=0, column=0, padx=(0, 6))
        self.save_html_btn = tk.Button(export_frame, text="HTML", command=self.save_html_file, state="disabled")
        self.save_html_btn.grid(row=0, column=1, sticky="ew", padx=4)
        self.save_excel_btn = tk.Button(export_frame, text="EXCEL", command=self.save_excel_file, state="disabled")
        self.save_excel_btn.grid(row=0, column=2, sticky="ew", padx=4)
        self.save_json_btn = tk.Button(export_frame, text="JSON", command=self.save_json_file, state="disabled")
        self.save_json_btn.grid(row=0, column=3, sticky="ew", padx=4)
        self.save_log_btn = tk.Button(export_frame, text="Log (txt)", command=self.on_save_log, state="disabled")
        self.save_log_btn.grid(row=0, column=4, sticky="ew", padx=4)

        bottom_frame = tk.Frame(self.root)
        left_sub, center_sub, right_sub = tk.Frame(bottom_frame), tk.Frame(bottom_frame), tk.Frame(bottom_frame)
        left_sub.pack(side="left", fill="x", expand=True)
        center_sub.pack(side="left")
        right_sub.pack(side="right")
        bottom_frame.grid(row=7, column=0, columnspan=4, sticky="ew", padx=8, pady=8)

        # Removed 'foreground="black"' to allow dynamic system theme inheritance
        tk.Label(left_sub, text="Status:").pack(side="left", padx=(2, 4))
        self.status_dynamic_label = tk.Label(left_sub, text="Standby", foreground=self.colors['status_standby'])
        self.status_dynamic_label.pack(side="left", padx=(0, 10))
        self.folders_label = tk.Label(center_sub, text="Folders: 0")
        self.folders_label.pack(side="left", padx=8)
        self.files_label = tk.Label(center_sub, text="Files: 0")
        self.files_label.pack(side="left", padx=8)

        self.clear_log_btn = tk.Button(right_sub, text="Clear Log", command=self.on_clear_log, state="disabled")
        self.clear_log_btn.pack(side="left", padx=(4, 6))
        self.import_btn = tk.Button(right_sub, text="Import Data", command=self.on_import_json)
        self.import_btn.pack(side="left", padx=(4, 6))
        self.help_btn = tk.Button(right_sub, text="   Help   ", command=self.on_help)
        self.help_btn.pack(side="left", padx=(4, 6))
        self.exit_btn = tk.Button(right_sub, text="   Exit   ", command=self.on_exit)
        self.exit_btn.pack(side="left", padx=(6, 2))

        # ------------------- State Machine Config -------------------
        self.STATE_CONFIG = {
            "standby": {
                self.start_btn: {"state": "normal", "text": "Start ▶"},
                self.stop_btn: {"state": "disabled"},
                self.browse_btn: {"state": "normal"},
                self.root_btn: {"state": "normal"},
                self.open_drive_btn: {"state": "normal"},
                self.credentials_entry: {"state": "normal"},
                self.folder_entry: {"state": "normal"},
                self.import_btn: {"state": "normal"},
                self.save_excel_btn: {"state": "disabled"},
                self.save_json_btn: {"state": "disabled"},
                self.save_html_btn: {"state": "disabled"},
                self.save_log_btn: {"state": "disabled"},
                self.clear_log_btn: {"state": "disabled"},
            },
            "running": {
                self.start_btn: {"state": "normal", "text": "Pause ⏸"},
                self.stop_btn: {"state": "normal"},
                self.browse_btn: {"state": "disabled"},
                self.root_btn: {"state": "disabled"},
                self.open_drive_btn: {"state": "disabled"},
                self.credentials_entry: {"state": "disabled"},
                self.folder_entry: {"state": "disabled"},
                self.save_log_btn: {"state": "disabled"},
                self.clear_log_btn: {"state": "disabled"},
                self.save_excel_btn: {"state": "disabled"},
                self.save_json_btn: {"state": "disabled"},
                self.save_html_btn: {"state": "disabled"},
                self.import_btn: {"state": "disabled"},
            },
            "paused": {
                self.start_btn: {"state": "normal", "text": "Resume ▶"},
                self.stop_btn: {"state": "normal"},
                self.browse_btn: {"state": "disabled"},
                self.root_btn: {"state": "disabled"},
                self.open_drive_btn: {"state": "disabled"},
                self.credentials_entry: {"state": "disabled"},
                self.folder_entry: {"state": "disabled"},
                self.save_log_btn: {"state": "disabled"},
                self.clear_log_btn: {"state": "disabled"},
                self.save_excel_btn: {"state": "disabled"},
                self.save_json_btn: {"state": "disabled"},
                self.save_html_btn: {"state": "disabled"},
                self.import_btn: {"state": "disabled"},
            },
            "stopped": {
                self.start_btn: {"state": "normal", "text": "Start ▶"},
                self.stop_btn: {"state": "disabled"},
                self.browse_btn: {"state": "normal"},
                self.root_btn: {"state": "normal"},
                self.open_drive_btn: {"state": "normal"},
                self.credentials_entry: {"state": "normal"},
                self.folder_entry: {"state": "normal"},
                self.import_btn: {"state": "normal"},
                self.save_excel_btn: {"state": "disabled"},
                self.save_json_btn: {"state": "disabled"},
                self.save_html_btn: {"state": "disabled"},
                self.save_log_btn: {"state": "disabled"},
                self.clear_log_btn: {"state": "disabled"},
            },
            "auth_wait": {
                self.start_btn: {"state": "disabled", "text": "Start ▶"},
                self.stop_btn: {"state": "normal"},
                self.browse_btn: {"state": "disabled"},
                self.root_btn: {"state": "disabled"},
                self.open_drive_btn: {"state": "disabled"},
                self.credentials_entry: {"state": "disabled"},
                self.folder_entry: {"state": "disabled"},
                self.save_log_btn: {"state": "disabled"},
                self.clear_log_btn: {"state": "disabled"},
                self.save_excel_btn: {"state": "disabled"},
                self.save_json_btn: {"state": "disabled"},
                self.save_html_btn: {"state": "disabled"},
                self.import_btn: {"state": "disabled"},
            },
            "exporting": {
                self.start_btn: {"state": "disabled", "text": "Start ▶"},
                self.stop_btn: {"state": "disabled"},
                self.browse_btn: {"state": "disabled"},
                self.root_btn: {"state": "disabled"},
                self.open_drive_btn: {"state": "disabled"},
                self.credentials_entry: {"state": "disabled"},
                self.folder_entry: {"state": "disabled"},
                self.save_log_btn: {"state": "disabled"},
                self.clear_log_btn: {"state": "disabled"},
                self.save_excel_btn: {"state": "disabled"},
                self.save_json_btn: {"state": "disabled"},
                self.save_html_btn: {"state": "disabled"},
                self.import_btn: {"state": "disabled"},
            }
        }

        # 2. State-Change Actions
        self.STATE_ACTIONS = {
            "standby": {"spinner": self.spinner_stop, "status": ("Standby", self.colors['status_standby'])},
            "running": {"spinner": self.spinner_start, "status": ("Collecting data...", self.colors['status_running'])},
            "paused": {"spinner": self.spinner_stop, "status": ("Paused ⏸", self.colors['status_paused'])},
            "stopped": {"spinner": self.spinner_stop, "status": ("Stopped ⏹", self.colors['status_stopped'])},
            "auth_wait": {"spinner": None, "status": ("Waiting for authentication...", self.colors['status_standby'])},
            "exporting": {"spinner": self.spinner_stop, "status": ("Exporting...", self.colors['status_running'])},
        }

        # --- Set initial state ---
        self.set_state("standby")

    # ---------- Theme Methods ----------

    def _detect_macos_dark_mode(self):
        """Detects if macOS is running in Dark Mode."""
        if sys.platform != 'darwin':
            return False
        try:
            val = subprocess.check_output(['defaults', 'read', '-g', 'AppleInterfaceStyle'], stderr=subprocess.DEVNULL)
            return b'Dark' in val
        except subprocess.CalledProcessError:
            return False
        except Exception:
            return False

    def _get_theme_colors(self, is_dark):
        """Returns a dictionary of colors based on the theme."""
        if is_dark:
            # DARK MODE PALETTE
            return {
                'log_bg': '#1e1e1e',     # Dark gray (VS Code-ish)
                'log_fg': '#e0e0e0',     # Light gray
                
                # Status Variables
                'status_standby': '#64B5F6', # Light Blue
                'status_running': '#81C784', # Light Green
                'status_paused':  '#FFB74D', # Light Orange
                'status_stopped': '#E57373', # Light Red

                # Tag Colors (Pastels/Neon)
                'tags': {
                    "info": "#e0e0e0",
                    "tree": "#90EE90",       # LightGreen
                    "status": "#64B5F6",     # Light Blue
                    "folder": "#98FB98",     # PaleGreen
                    "warning": "#FF8A80",    # Light Red/Salmon
                    "parent": "#DEB887",     # Burlywood
                    "size": "#D8BFD8",       # Thistle
                    "date": "#FFCC80",       # Light Orange
                    "mime": "#87CEFA",       # LightSkyBlue
                    "level": "#F08080",      # Light Coral
                    "id": "#B0C4DE",         # LightSteelBlue
                    "success": "#66BB6A",    # Medium Green
                    "filename_highlight": "#69F0AE", # Neon Green
                    "owner": "#FFAB91",      # Light Salmon
                    "created": "#E1BEE7",    # Light Purple
                    "modifier": "#EF9A9A",   # Light Red
                    "md5": "#C5E1A5",        # Light Green
                    "shared": "#FF5252",     # Red Accent
                    "version": "#9FA8DA"     # Light Indigo
                }
            }
        else:
            # LIGHT MODE PALETTE (Original)
            return {
                'log_bg': 'white',
                'log_fg': 'black',

                # Status Variables
                'status_standby': '#1E90FF', # DodgerBlue
                'status_running': '#28A745', # Green
                'status_paused':  '#FFA500', # Orange
                'status_stopped': '#DC3545', # Red

                # Tag Colors (Darker for white background)
                'tags': {
                    "info": "black",
                    "tree": "#008000",
                    "status": "blue",
                    "folder": "green",
                    "warning": "red",
                    "parent": "#A0522D",
                    "size": "#800080",
                    "date": "#FF8C00",
                    "mime": "#1E90FF",
                    "level": "#B22222",
                    "id": "#2F4F4F",
                    "success": "#28A745",
                    "filename_highlight": "#008000",
                    "owner": "#D2691E",
                    "created": "#8B008B",
                    "modifier": "#8B0000",
                    "md5": "#556B2F",
                    "shared": "#DC143C",
                    "version": "#4B0082"
                }
            }


    # ---------- Helper Methods ----------

    def simplify_mime_type(self, mime_type):
        return MIME_TYPE_MAP.get(mime_type, mime_type)

    def ts_now(self):
        """Return current timestamp (HH:MM:SS)."""
        return datetime.now().strftime("%H:%M:%S")

    def format_size(self, size):
        """Convert bytes to a human-readable string (B, KB, MB, GB)."""
        if not size:
            return "0 B"
        try:
            s = int(size)
        except Exception:
            return str(size)
        if s < 1024:
            return f"{s} B"
        elif s < 1024 ** 2:
            return f"{s / 1024:.2f} KB"
        elif s < 1024 ** 3:
            return f"{s / (1024 ** 2):.2f} MB"
        else:
            return f"{s / (1024 ** 3):.2f} GB"

    def sanitize_name(self, name: str) -> str:
        """
        Replace '/' with '／' (fullwidth slash U+FF0F).
        This avoids visual confusion in Parent_Path in Log and Excel exports.
        """
        if not name:
            return ""
        return name.replace("/", "／")

    def apply_folder_size_totals(self, rows, totals_snapshot):
        """
        Replace folder size column values with aggregated totals.
        Updates both raw and formatted size columns.
        """
        for r in rows:
            try:
                if (r.type or "").lower() == "folder":
                    fid = r.id if r.id else ""
                    total_bytes = totals_snapshot.get(fid, 0)
                    r.raw_size = total_bytes
                    r.size = self.format_size(total_bytes)
            except Exception as e:
                self.gui_log_insert(f"⚠ Error applying totals: {e}", "warning")
                continue
        return rows

    def format_date(self, date_str):
        """Convert ISO date string from Google API to local time DD/MM/YYYY HH:MM format."""
        if not date_str:
            return "N/A"
        try:
            if date_str.endswith("Z"):
                date_str = date_str[:-1] + "+00:00"
            dt_utc = datetime.fromisoformat(date_str)
            dt_local = dt_utc.astimezone(None)
            return dt_local.strftime("%d/%m/%Y %H:%M")
        except (ValueError, TypeError):
            return date_str

    # ---------- GUI-safe operations ----------

    def gui_log_insert(self, line, tag="info", add_ts=True):
        """Insert a line into the log box from worker thread, safely via `root.after`."""
        if add_ts:
            line = f"[{self.ts_now()}] {line}"

        def _insert():
            self.log_box.config(state="normal")
            self.log_box.insert(tk.END, line + "\n", tag)
            self.log_box.config(state="disabled")
            if self.should_auto_scroll:
                self.log_box.see(tk.END)

        self.root.after(0, _insert)

    def gui_log_separator(self):
        """Inserts a visual separator line into the log box, safely."""
        line = "─" * 120

        def _insert():
            self.log_box.config(state="normal")
            self.log_box.insert(tk.END, line + "\n", "info")
            self.log_box.config(state="disabled")
            if self.should_auto_scroll:
                self.log_box.see(tk.END)

        self.root.after(0, _insert)

    def gui_update_counters(self):
        """Update file/folder counters in the GUI safely."""
        with self._data_lock:
            folders = self.total_folders
            files = self.total_files

        def _u():
            self.folders_label.config(text=f"Folders: {folders}")
            self.files_label.config(text=f"Files: {files}")

        self.root.after(0, _u)

    def gui_set_status(self, text, color=None):
        """Update status label safely with optional color."""
        def _set():
            self.status_dynamic_label.config(text=text)
            if color:
                self.status_dynamic_label.config(foreground=color)

        self.root.after(0, _set)

    def _close_active_modal_if_any(self, set_negative=True):
        """Close any active CustomMessageBox registered on the root."""
        modal = getattr(self.root, "_active_modal", None)

        if modal and isinstance(modal, tk.Toplevel):
            try:
                try:
                    self.gui_log_insert("ℹ️ Auto-closing active confirmation dialog...", "status")
                except Exception:
                    pass

                if set_negative:
                    try:
                        modal.result = False
                    except Exception:
                        pass

                try:
                    modal.destroy()
                except Exception:
                    try:
                        tk.Toplevel.destroy(modal)
                    except Exception:
                        pass
            finally:
                try:
                    if getattr(self.root, '_active_modal', None) is modal:
                        setattr(self.root, '_active_modal', None)
                except Exception:
                    pass

    # ------------------- State Machine Engine -------------------

    def set_state(self, new_state):
        """
        Change application state, update controls from the declarative state machine,
        and handle dynamic logic.
        """
        self.app_state = new_state

        def _update_gui():
            config = self.STATE_CONFIG.get(new_state, {})
            actions = self.STATE_ACTIONS.get(new_state, {})

            # Apply base widget configurations
            for widget, props in config.items():
                try:
                    widget.config(**props)
                except Exception as e:
                    print(f"Warning: Failed to configure widget {widget}: {e}")

            # Perform state-change actions
            if actions.get("spinner"):
                actions["spinner"]()

            if actions.get("status"):
                status_text, status_color = actions["status"]
                self.gui_set_status(status_text, status_color)

            # Handle DYNAMIC logic
            if new_state == "running":
                if self.is_importing:
                    self.start_btn.config(state="disabled", text="Start ▶")
                    self.gui_set_status("Importing data...", self.colors['status_running'])

            # --- DYNAMIC AUTH CHECK ---
            # Checks if we have an active session in EITHER 'standby' OR 'stopped' states.
            if new_state in ["standby", "stopped"]:
                with self._data_lock:
                    is_authenticated = (self.g_service is not None)
                
                if is_authenticated:
                    self.credentials_entry.config(state="disabled")

            # Enable/Disable Export buttons based on data presence
            if new_state in ["standby", "stopped"]:
                with self._data_lock:
                    has_data = bool(self.data_rows)

                self.save_excel_btn.config(state="normal" if has_data else "disabled")
                self.save_json_btn.config(state="normal" if has_data else "disabled")
                self.save_html_btn.config(state="normal" if has_data else "disabled")

                has_log_entries = bool(self.log_box.get(1.0, tk.END).strip())
                self.save_log_btn.config(state="normal" if has_log_entries else "disabled")
                self.clear_log_btn.config(state="normal" if has_log_entries else "disabled")

        self.root.after(0, _update_gui)

    # ---------- Spinner ----------

    def spinner_step(self):
        """Advance spinner animation in the status bar while listing is running."""
        if self.app_state == "running":
            ch = SPINNER[self.spinner_index % len(SPINNER)]
            base_status = "Importing data" if self.is_importing else "Collecting data"
            self.gui_set_status(f"{base_status} {ch}", self.colors['status_running'])
            self.spinner_index += 1
            self.spinner_task = self.root.after(140, self.spinner_step)
        else:
            self.spinner_task = None

    def spinner_start(self):
        """Start the spinner animation."""
        if self.spinner_task is None:
            self.spinner_step()

    def spinner_stop(self):
        """Stop the spinner animation if active."""
        if self.spinner_task:
            try:
                self.root.after_cancel(self.spinner_task)
            except Exception:
                pass
            self.spinner_task = None

    # ---------- Logging ----------

    def append_log(self, name_with_indent, item_type, size_str, date_str, created_date_str, mime, level, rel_path,
                   node_id, owner, modified_by, md5, shared, version):
        """Queue a log entry to be processed by the batch worker."""
        with self._data_lock:
            self.has_unsaved_data = True

        timestamp = self.ts_now()

        # Package the data for the queue
        log_data = {
            'timestamp': timestamp,
            'name_with_indent': name_with_indent,
            'rel_path': rel_path,
            'size_str': size_str,
            'date_str': date_str,
            'created_date_str': created_date_str,
            'mime': mime,
            'owner': owner,
            'modified_by': modified_by,
            'md5': md5,
            'shared': shared,
            'version': version,
            'level': level,
            'node_id': node_id
        }
        
        # Put it in the queue for the GUI thread to pick up
        self.log_queue.put(log_data)

    def _process_log_queue(self):
        """Batch process queued log entries to keep the GUI responsive."""
        if not self.log_queue.empty():
            items_to_process = []
            
            # Pull everything currently in the queue
            while not self.log_queue.empty():
                try:
                    items_to_process.append(self.log_queue.get_nowait())
                except queue.Empty:
                    break

            if items_to_process:
                self.log_box.config(state="normal")
                
                for item in items_to_process:
                    # Timestamp
                    self.log_box.insert(tk.END, f"[{item['timestamp']}] ", "info")
                    # Tree / Name
                    self.log_box.insert(tk.END, item['name_with_indent'], "tree")
                    # Parent Path
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, item['rel_path'], "parent")
                    # Size
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, item['size_str'], "size")
                    # Mod Date
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, item['date_str'], "date")
                    # Created Date
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, item['created_date_str'], "created")
                    # MIME
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, item['mime'], "mime")
                    # Owner
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, item['owner'], "owner")
                    # Modified By
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, item['modified_by'], "modifier")
                    # MD5
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    if item['md5']:
                        self.log_box.insert(tk.END, item['md5'], "md5")
                    else:
                        self.log_box.insert(tk.END, "--", "info")
                    # Shared
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    if item['shared']:
                        self.log_box.insert(tk.END, "SHARED", "shared")
                    else:
                        self.log_box.insert(tk.END, "--", "info")
                    # Version
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, f"v{item['version']}", "version")
                    # Level
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, str(item['level']), "level")
                    # File/Folder ID
                    self.log_box.insert(tk.END, SEPARATOR, "info")
                    self.log_box.insert(tk.END, item['node_id'], "id")
        
                    self.log_box.insert(tk.END, "\n")
                    
                self.log_box.config(state="disabled")
                
                # Scroll exactly ONCE per batch
                if self.should_auto_scroll:
                    self.log_box.see(tk.END)

        # Reschedule the next check
        self.root.after(100, self._process_log_queue)

    # ---------- Export Helpers (shared by exports) ----------

    def _tokenize_text_numbers(self, fragment: str):
        """Tokenize into numbers and text fragments for natural sort."""
        tokens = []
        for p in re.split(r'(\d+)', fragment):
            if not p:
                continue
            if p.isdigit():
                tokens.append((0, int(p)))
            else:
                tokens.append((2, p.lower()))
        return tokens

    def _tokenize_part(self, part: str):
        """Tokenize one filename/path component with special handling for months."""
        part = (part or "").strip()
        if not part:
            return [(2, "")]
        pl = part.lower()
        for mname, midx in MONTH_MAP.items():
            if pl == mname:
                return [(1, midx)]
            if pl.startswith(mname) and len(pl) > len(mname) and not pl[len(mname)].isalpha():
                rest = part[len(mname):]
                return [(1, midx)] + self._tokenize_text_numbers(rest)
        return self._tokenize_text_numbers(part)

    def _name_key(self, path: str):
        """Return sorting key for natural ordering by last non-empty path component."""
        comp = path.rstrip("/").split('/')[-1]
        return self._tokenize_part(comp)

    def _build_tree(self, rows, parent_map):
        """Build parent→children relationships by file/folder ID for tree rendering."""
        children = defaultdict(list)
        type_by_id = {}
        row_by_id = {}

        for r in rows:
            fid = r.id if r.id else ""
            if not fid:
                continue
            row_by_id[fid] = r
            type_by_id[fid] = (r.type or "").lower()

        for fid in list(row_by_id.keys()):
            parent = parent_map.get(fid, None)
            # Normalize parent key
            if parent is None or parent == "" or parent == "--":
                pkey = ""
            else:
                pkey = parent
            children[pkey].append(fid)

        return children, type_by_id, row_by_id

    def sort_rows_folders_first(self, rows, parent_map):
        """
        Sort rows in pre-order with folders first, then files, natural order by name.
        Mutates the 'tree' and 'level' attributes of the DataRow objects.
        """
        children, type_by_id, row_by_id = self._build_tree(rows, parent_map)
        out = []

        def name_key_for_id(fid):
            name = row_by_id[fid].name if fid in row_by_id else ""
            return self._name_key(name)

        def process_children(parent_id, prefix_str, depth):
            kids = children.get(parent_id, [])
            kids.sort(key=lambda fid: (0 if type_by_id.get(fid) == 'folder' else 1, name_key_for_id(fid)))
            for idx, fid in enumerate(kids):
                is_last = (idx == len(kids) - 1)
                typ = type_by_id.get(fid, '')
                indent_repr = prefix_str + ("└─" if is_last else "├─") + (
                    TREE_SYMBOL_FOLDER if typ == 'folder' else TREE_SYMBOL_FILE)

                if fid in row_by_id:
                    orig_row = row_by_id[fid]
                    orig_row.tree = indent_repr
                    orig_row.level = depth
                    out.append(orig_row)

                child_prefix = prefix_str + ("  " if is_last else "│ ")
                if typ == 'folder':
                    process_children(fid, child_prefix, depth + 1)

        # Determine explicit roots (rows that originally had parent == '--')
        roots = [fid for fid, r in row_by_id.items() if r.parent == "--"]

        if roots:
            roots.sort(key=lambda fid: name_key_for_id(fid))
            for root in roots:
                typ = type_by_id.get(root, '')
                if root in row_by_id:
                    orig_row = row_by_id[root]
                    orig_row.tree = (TREE_SYMBOL_FOLDER if typ == 'folder' else TREE_SYMBOL_FILE)
                    orig_row.level = 0
                    out.append(orig_row)
                process_children(root, "", 1)
            return out
        else:
            # fallback: print children of virtual root key ''
            process_children("", "", 0)
            return out

    # ---------- Export Methods ----------

    def save_excel_file(self):
        """
        Save the collected data into an Excel file (.xlsx), including the raw size.
        Applies folder size totals from `folder_size_totals` before writing.
        """
        with self._data_lock:
            if not self.data_rows:
                self.gui_log_insert("⚠ Nothing to save in Excel.", "warning")
                return

        self.set_state('exporting')

        try:
            # Take snapshots
            with self._data_lock:
                rows_snapshot = list(self.data_rows)
                totals_snapshot = dict(self.folder_size_totals)
                parents_snapshot = dict(self.parent_ids_by_child)
                folder_name_snapshot = self.selected_folder_name

            default_name = folder_name_snapshot if folder_name_snapshot else "listing"
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=f"{default_name}.xlsx"
            )
            if not path:
                return

            while True:
                try:
                    dummy_wb = openpyxl.Workbook()
                    dummy_wb.save(path)
                    break
                except PermissionError:
                    retry = custom_askyesnocancel(
                        self.root,
                        "File In Use",
                        f"The file '{os.path.basename(path)}'\n is open in another program and cannot be overwritten.\n\nChoose:\n"
                        " • 'Yes' — after closing the file — to try saving again\n"
                        " • 'No' to choose another name\n"
                        " • 'Cancel' to abort."
                    )
                    if retry is None:
                        return
                    elif retry is False:
                        new_path = filedialog.asksaveasfilename(
                            defaultextension=".xlsx",
                            filetypes=[("Excel files", "*.xlsx")],
                            initialfile=f"{default_name}.xlsx"
                        )
                        if not new_path:
                            return
                        path = new_path
                        continue
                    else:
                        continue
                except Exception as e:
                    custom_showerror(self.root, "Error", f"Cannot write to file: {e}")
                    return

            self.gui_log_insert("⚙️ Preparing data for Excel file...", "status")
            try:
                self.root.update()
            except Exception:
                pass

            try:
                rows = self.apply_folder_size_totals(rows_snapshot, totals_snapshot)
                rows = self.sort_rows_folders_first(rows, parents_snapshot)

                # Build Workbook
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "File Tree"

                headers = [
                    "#", "Tree", "Name", "Size", "Raw Size", "Modified Date", "Created Date",
                    "Type", "MIME Type", "Parent Path", "Level", "Owner", "Modified By",
                    "Shared", "Version", "MD5 Checksum", "File/Folder ID"
                ]
                ws.append(headers)
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                ws.freeze_panes = 'A2'

                row_num = 2
                for row_index, r_obj in enumerate(rows, 1):
                    ws.append([
                        row_index,
                        r_obj.tree,
                        r_obj.name,
                        r_obj.size,
                        r_obj.raw_size,
                        r_obj.date,
                        r_obj.created_date,
                        r_obj.type,
                        self.simplify_mime_type(r_obj.mime),
                        r_obj.parent,
                        r_obj.level,
                        r_obj.owner,
                        r_obj.modified_by,
                        r_obj.shared,
                        r_obj.version,
                        r_obj.md5,
                        r_obj.id
                    ])

                    if r_obj.id:
                        cell = ws.cell(row=row_num, column=3)
                        cell.hyperlink = f"https://drive.google.com/open?id={r_obj.id}"
                        cell.style = "Hyperlink"

                    row_num += 1

                mono = Font(name="Courier New")
                for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
                    for cell in row:
                        cell.font = mono

                # Center align ID and Levels
                col_idx_level = 11
                col_letter = get_column_letter(col_idx_level)
                center_align = Alignment(horizontal="center", vertical="center")
                ws[f"{col_letter}1"].alignment = center_align
                for row in ws.iter_rows(min_row=2, min_col=col_idx_level, max_col=col_idx_level, max_row=ws.max_row):
                    for cell in row:
                        cell.alignment = center_align

                col_idx_hash = 1
                col_letter_hash = get_column_letter(col_idx_hash)
                ws[f"{col_letter_hash}1"].alignment = center_align
                for row in ws.iter_rows(min_row=2, min_col=col_idx_hash, max_col=col_idx_hash, max_row=ws.max_row):
                    for cell in row:
                        cell.alignment = center_align

                for col in ws.columns:
                    max_len = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        if cell.value is not None:
                            l = len(str(cell.value))
                            if l > max_len:
                                max_len = l

                    if col_letter == 'A':
                        width = int(max(6, max_len + 2))
                    elif col_letter == 'B':
                        width = int(min(120, max(20, max_len * 1.2 + 4)))
                    else:
                        width = int(min(60, max(10, max_len + 2)))

                    ws.column_dimensions[col_letter].width = width

                self.gui_log_insert(f"📄 Data ready ({len(rows)} rows). Saving...", "status")
                try:
                    self.root.update()
                except Exception:
                    pass

            except Exception as e:
                custom_showerror(self.root, "Error", f"Failed to prepare Excel file: {e}")
                self.gui_log_insert(f"⚠ Excel export preparation error: {e}", "warning")
                self.gui_log_separator()
                return

            try:
                wb.save(path)
                self.gui_log_insert(f"💾 Data exported successfully to: {path}", "status")
                self.gui_log_separator()
                self.has_unsaved_data = False
                self.is_live_scan_data = False

            except Exception as e:
                custom_showerror(self.root, "Error",
                                 f"Failed to save Excel file: {e}\n\nThe file may have become locked after the initial check.")
                self.gui_log_insert(f"⚠ Excel save error (race condition?): {e}", "warning")
                self.gui_log_separator()

        finally:
            self.set_state('standby')


    def save_html_file(self):
        """Save the collected data into an HTML file using an external template."""
        with self._data_lock:
            if not self.data_rows:
                self.gui_log_insert("⚠ Nothing to save in HTML.", "warning")
                return

        self.set_state('exporting')
        export_canceled = True

        try:
            # Take snapshots
            with self._data_lock:
                rows_snapshot = list(self.data_rows)
                totals_snapshot = dict(self.folder_size_totals)
                parents_snapshot = dict(self.parent_ids_by_child)
                folder_name_snapshot = self.selected_folder_name
                account_snapshot = self.data_account_email
                folders_snapshot = self.total_folders
                files_snapshot = self.total_files
                start_snapshot = self.acquisition_start

            default_name = folder_name_snapshot if folder_name_snapshot else "listing"
            path = filedialog.asksaveasfilename(
                defaultextension=".html",
                filetypes=[("HTML files", "*.html")],
                initialfile=f"{default_name}.html"
            )

            export_canceled = not path
            if export_canceled:
                return

            # Styling & Calculation Constants
            STYLE_CONSTS = {
                'TREE_FONT_FAMILY': "'Courier New', monospace",
                'TREE_FONT_SIZE_PX': 24,
                'TREE_LETTER_SPACING_PX': -3,
                'NAME_FONT_SIZE_PX': 16,
                'NAME_LETTER_SPACING_PX': 0,
                'ICON_FOLDER_FONT_SIZE_PX': 20,
                'ICON_FILE_FONT_SIZE_PX': 15,
            }

            tree_char_width_px = (STYLE_CONSTS['TREE_FONT_SIZE_PX'] * 0.6) + STYLE_CONSTS['TREE_LETTER_SPACING_PX']
            name_char_width_px = (STYLE_CONSTS['NAME_FONT_SIZE_PX'] * 0.5) + STYLE_CONSTS['NAME_LETTER_SPACING_PX']
            icon_width_px = STYLE_CONSTS['ICON_FOLDER_FONT_SIZE_PX']
            visible_name_chars = 8
            general_padding_px = 40

            def safe_str(v):
                return "" if v is None else str(v)

            rows = self.apply_folder_size_totals(rows_snapshot, totals_snapshot)
            rows = self.sort_rows_folders_first(rows, parents_snapshot)
            self.gui_log_insert(f"📤 Exporting {len(rows)} rows to HTML...", "status")
            try:
                self.root.update()
            except Exception:
                pass

            # --- Calculate Max Content Widths for Cols 2–5 ---
            # NOTE: HTML export intentionally excludes technical fields (MD5, Version, Shared) to remain clean.
            headers = {
                'size': "Size",
                'date': "Modified Date",
                'mime': "MIME Type",
                'owner': "Owner"
            }
            max_lens = {k: len(v) for k, v in headers.items()}

            for r in rows:
                max_lens['size'] = max(max_lens['size'], len(safe_str(r.size)))
                max_lens['date'] = max(max_lens['date'], len(safe_str(r.date)))
                max_lens['mime'] = max(max_lens['mime'], len(self.simplify_mime_type(safe_str(r.mime))))
                max_lens['owner'] = max(max_lens['owner'], len(safe_str(r.owner)))

            CHAR_WIDTH_ESTIMATE = 8
            CELL_PADDING = 30

            col_widths_px = {
                'COL_SIZE_WIDTH_PX': (max_lens['size'] * CHAR_WIDTH_ESTIMATE) + CELL_PADDING,
                'COL_DATE_WIDTH_PX': (max_lens['date'] * CHAR_WIDTH_ESTIMATE) + CELL_PADDING,
                'COL_MIME_WIDTH_PX': (max_lens['mime'] * CHAR_WIDTH_ESTIMATE) + CELL_PADDING,
                'COL_OWNER_WIDTH_PX': (max_lens['owner'] * CHAR_WIDTH_ESTIMATE) + CELL_PADDING
            }

            for k in col_widths_px:
                col_widths_px[k] = int(max(90, min(400, col_widths_px[k])))

            fixed_buffer_for_other_cols = sum(col_widths_px.values())

            # --- Column 1 MIN width only (Auto handles max) ---
            max_tree_prefix_len = max(max(0, len(safe_str(r.tree)) - 1) for r in rows) if rows else 0

            tree_prefix_width_min = max_tree_prefix_len * tree_char_width_px
            visible_name_width_min = visible_name_chars * name_char_width_px

            tree_name_min_px = int(
                tree_prefix_width_min +
                icon_width_px +
                visible_name_width_min +
                general_padding_px
            )

            total_min_px = int(tree_name_min_px + fixed_buffer_for_other_cols)
            storage_hash = f"{random.getrandbits(48):012x}"

            # Read Template File
            try:
                base_dir = get_base_path()
            except NameError:
                base_dir = os.getcwd()
            template_path = os.path.join(base_dir, "TreeX4GD-template.html")

            # Check if the template exists before trying to open it
            if not os.path.exists(template_path):
                msg = (
                    "The required HTML template file could not be found.\n\n"
                    "Please ensure 'TreeX4GD-template.html' is located in the same folder as the application:\n"
                    f"{base_dir}"
                )
                custom_showerror(self.root, "Missing Template", msg)
                self.gui_log_insert("❌ Export canceled. 'TreeX4GD-template.html' is missing.", "warning")
                return

            # Attempt to read the file safely
            try:
                with open(template_path, "r", encoding="utf-8") as f:
                    template_content = f.read()
            except Exception as e:
                custom_showerror(self.root, "Error Reading Template", f"Failed to read the HTML template file:\n{e}")
                self.gui_log_insert(f"❌ Export canceled. Could not read template: {e}", "warning")
                return

            # Check Template Version
            try:
                content_head = template_content[:1024]
                match = re.search(
                    r'<meta\s+name="template-version"\s+content="([^"]+)"',
                    content_head,
                    re.IGNORECASE
                )
                template_version = match.group(1).strip() if match else None

                if not template_version:
                    raise ValueError("Template version metadata missing.")

                # Extract digits for major/minor comparison (safely drops the 'x')
                tmpl_parts = [int(p) for p in template_version.split('.') if p.isdigit()]
                exp_parts = [int(p) for p in EXPECTED_TEMPLATE_VERSION.split('.') if p.isdigit()]

                if len(tmpl_parts) < 2 or len(exp_parts) < 2:
                    raise ValueError(f"Version format invalid (Found {template_version}, Expected {EXPECTED_TEMPLATE_VERSION})")

                # Strict Major.Minor Check (Ignore patch version completely)
                if tmpl_parts[0] != exp_parts[0] or tmpl_parts[1] != exp_parts[1]:
                    found_ver_str = f"v{template_version}"
                    msg_popup = (
                        "The 'TreeX4GD-template.html' file is incompatible.\n\n"
                        f"Its version is {found_ver_str}, but the script requires "
                        f"v{exp_parts[0]}.{exp_parts[1]}.x.\n\n"
                        "Please ensure all script files are from the same release."
                    )
                    custom_showerror(self.root, "HTML Template Incompatible", msg_popup)
                    self.gui_log_insert(
                        f"❌ Export canceled. Template version incompatible. "
                        f"Expected v{EXPECTED_TEMPLATE_VERSION}, found {found_ver_str}.",
                        "warning"
                    )
                    return

                # If major and minor match, it's automatically valid.
                self.gui_log_insert(f"ℹ️ HTML template version OK (v{template_version})", "status")

            except ValueError as ve:
                custom_showerror(self.root, "HTML Template Error", f"Version check failed: {ve}")
                self.gui_log_insert(f"❌ Export canceled. {ve}", "warning")
                return
            except Exception as e:
                custom_showerror(
                    self.root,
                    "HTML Template Error",
                    f"An unexpected error occurred while parsing the HTML template.\n\n{e}"
                )
                self.gui_log_insert(f"❌ Export canceled. Failed to parse HTML template: {e}", "warning")
                return

                



            # --- Generate HTML Rows ---
            table_rows_html = []

            for r in rows:
                level_val = safe_str(r.level)
                raw_tree = safe_str(r.tree)
                name_val = safe_str(r.name)
                node_id = safe_str(r.id)
                node_key = escape(raw_tree + "|" + name_val)

                table_rows_html.append(
                    f"<tr data-level='{escape(level_val)}'" +
                    (f" data-nodeid='{escape(node_id)}'>" if node_id else
                    f" data-nodekey='{node_key}'>")
                )

                is_folder = (safe_str(r.type).lower() == "folder")
                emoji = "📂" if is_folder else "📄"
                tree_part = raw_tree.replace("■", emoji).replace("●", emoji)
                escaped_name = escape(name_val)

                if node_id:
                    url = f"https://drive.google.com/open?id={escape(node_id)}"
                    title_text = escape(f'Open "{name_val}" in Google Drive')
                    link_html = (
                        f"<a href='{url}' target='_blank' "
                        f"rel='noopener noreferrer' title='{title_text}' "
                        f"style='cursor: alias;'>{escaped_name}</a>"
                    )
                else:
                    link_html = escaped_name

                tree_lines = tree_part[:-1] if tree_part else ""
                tree_icon = tree_part[-1] if tree_part else emoji

                icon_class = "icon-folder" if is_folder else "icon-file"

                table_rows_html.append(
                    f"<td><span class='tree'>{escape(tree_lines)}"
                    f"<span class='{icon_class}'>{tree_icon}</span></span>&nbsp;{link_html}</td>"
                )

                size_val = escape(safe_str(r.size))
                table_rows_html.append(
                    f"<td class='fsize'>{size_val}</td>" if is_folder
                    else f"<td class='fsize-file'>{size_val}</td>"
                )

                table_rows_html.append(f"<td class='date meta'>{escape(safe_str(r.date))}</td>")
                table_rows_html.append(
                    f"<td class='mime meta'>{escape(self.simplify_mime_type(safe_str(r.mime)))}</td>"
                )
                table_rows_html.append(
                    f"<td class='owner meta'>{escape(safe_str(r.owner))}</td>"
                )

                table_rows_html.append("</tr>")

            table_rows_string = "\n".join(table_rows_html)

            # --- Inject into Template ---
            root_total = rows[0].size if rows else "--"
            max_depth = max((r.level for r in rows_snapshot), default=0)
            acq_date_str = (
                start_snapshot.strftime("%Y-%m-%d %H:%M:%S")
                if start_snapshot else "N/A"
            )

            final_html = template_content.replace("{{PAGE_TITLE}}", escape(default_name))
            final_html = final_html.replace("{{TOTAL_MIN_PX}}", str(total_min_px))
            final_html = final_html.replace("{{TREE_NAME_MIN_PX}}", str(tree_name_min_px))

            for key, value in col_widths_px.items():
                final_html = final_html.replace(f"{{{{{key}}}}}", str(value))

            for key, value in STYLE_CONSTS.items():
                final_html = final_html.replace(f"{{{{{key}}}}}", str(value))

            final_html = final_html.replace("{{STORAGE_HASH}}", storage_hash)
            final_html = final_html.replace("{{GLOBAL_PREFIX}}", self.HTML_GLOBAL_PREFIX)

            final_html = final_html.replace("{{SUMMARY_ACCOUNT}}", escape(account_snapshot or "(unavailable)"))
            final_html = final_html.replace("{{SUMMARY_ROOT_FOLDER}}", escape(default_name))
            final_html = final_html.replace("{{SUMMARY_TOTAL_FOLDERS}}", escape(str(folders_snapshot)))
            final_html = final_html.replace("{{SUMMARY_TOTAL_FILES}}", escape(str(files_snapshot)))
            final_html = final_html.replace("{{SUMMARY_ROWS_EXPORTED}}", escape(str(len(rows))))
            final_html = final_html.replace("{{SUMMARY_TOTAL_SIZE}}", escape(str(root_total)))
            final_html = final_html.replace("{{SUMMARY_MAX_DEPTH}}", escape(str(max_depth)))
            final_html = final_html.replace("{{SUMMARY_ACQ_DATE}}", escape(acq_date_str))
            final_html = final_html.replace("{{SUMMARY_CREATED_BY}}", escape(SCRIPT_FULL_NAME))
            final_html = final_html.replace("{{SUMMARY_VERSION}}", escape(SCRIPT_VERSION))

            final_html = final_html.replace("{{TABLE_ROWS}}", table_rows_string)

            with open(path, "w", encoding="utf-8") as f:
                f.write(final_html)

            self.gui_log_insert(f"💾 HTML saved to: {path}", "status")
            self.has_unsaved_data = False
            self.is_live_scan_data = False

        except Exception as e:
            custom_showerror(self.root, "Error", f"Failed to save HTML file: {e}")
            self.gui_log_insert(f"⚠ HTML save error: {e}", "warning")

        finally:
            if not export_canceled:
                self.gui_log_separator()
            self.set_state('standby')


    def save_json_file(self):
        """
        Save the collected data into a hierarchical JSON file.
        Includes metadata like script name, version, and account email.
        """
        with self._data_lock:
            if not self.data_rows:
                self.gui_log_insert("⚠ Nothing to save in JSON.", "warning")
                return

        self.set_state('exporting')

        try:
            # Take snapshots
            with self._data_lock:
                rows_snapshot = list(self.data_rows)
                totals_snapshot = dict(self.folder_size_totals)
                parents_snapshot = dict(self.parent_ids_by_child)
                folder_name_snapshot = self.selected_folder_name
                account_snapshot = self.data_account_email
                start_snapshot = self.acquisition_start

            rows = self.apply_folder_size_totals(rows_snapshot, totals_snapshot)
            sorted_rows = self.sort_rows_folders_first(rows, parents_snapshot)

            default_name = folder_name_snapshot if folder_name_snapshot else "listing"
            path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")],
                initialfile=f"{default_name}.json"
            )
            if not path:
                return

            self.gui_log_insert(f"📤 Exporting {len(sorted_rows)} rows to JSON...", "status")
            try:
                self.root.update()
            except Exception:
                pass

            try:
                tree_data = []
                parent_stack = []

                for r_obj in sorted_rows:
                    node = {
                        "tree_prefix": r_obj.tree,
                        "name": r_obj.name,
                        "type": r_obj.type,
                        "size": r_obj.size,
                        "raw_size": r_obj.raw_size,
                        "modified_date": r_obj.date,
                        "created_date": r_obj.created_date,
                        "mime_type": r_obj.mime,
                        "level": r_obj.level,
                        "parent_path": r_obj.parent,
                        "id": r_obj.id,
                        "owner": r_obj.owner,
                        "modified_by": r_obj.modified_by,
                        "md5": r_obj.md5,
                        "shared": r_obj.shared,
                        "version": r_obj.version
                    }
                    if node.get('type', '').lower() == 'folder':
                        node['children'] = []
                    else:
                        node['children'] = None
                    level = node.get('level', 0)
                    if level == 0:
                        tree_data.append(node)
                        parent_stack = [node['children']] if node['children'] is not None else []
                    else:
                        while len(parent_stack) > level:
                            parent_stack.pop()
                        if parent_stack:
                            parent_stack[-1].append(node)
                        if node['children'] is not None:
                            parent_stack.append(node['children'])

                json_output = {
                    "script_name": SCRIPT_FULL_NAME,
                    "script_version": SCRIPT_VERSION,
                    "account_email": account_snapshot,
                    "root_folder_name": folder_name_snapshot,
                    "acquisition_date": start_snapshot.isoformat() if start_snapshot else datetime.now().isoformat(),
                    "tree_data": tree_data
                }

                with open(path, 'w', encoding='utf-8') as f:
                    json.dump(json_output, f, indent=4, ensure_ascii=False)

                self.gui_log_insert(f"💾 JSON saved to: {path}", "status")
                self.gui_log_separator()
                self.has_unsaved_data = False
                self.is_live_scan_data = False

            except Exception as e:
                custom_showerror(self.root, "Error", f"Failed to save JSON file: {e}")
                self.gui_log_insert(f"⚠ JSON save error: {e}", "warning")
                self.gui_log_separator()

        finally:
            self.set_state('standby')

    # ---------- Core (worker methods) ----------

    def _calculate_folder_sizes(self):
        """
        Iterates through all files and calculates cumulative sizes for all parent folders.
        This is run as a separate pass AFTER all data is loaded.
        """
        with self._data_lock:
            rows_snapshot = list(self.data_rows)
            parents_snapshot = dict(self.parent_ids_by_child)

        new_totals = defaultdict(int)

        for row in rows_snapshot:
            if row.type.lower() != 'folder':
                try:
                    raw_size = int(row.raw_size)
                    file_id = row.id
                    if raw_size > 0 and file_id:
                        ancestor = parents_snapshot.get(file_id)
                        while ancestor:
                            new_totals[ancestor] += raw_size
                            ancestor = parents_snapshot.get(ancestor)
                except (ValueError, IndexError):
                    continue

        with self._data_lock:
            self.folder_size_totals = new_totals

    def list_drive_folder_worker(self, service, folder_id, level=0, parent_path="", prefix_state=None):
        """Recursive worker that lists the contents of a Google Drive folder."""
        if prefix_state is None:
            prefix_state = []

        page_token = None
        query = f"'{folder_id}' in parents and trashed=false"
        fields = (
            'nextPageToken, '
            'files(id, name, mimeType, size, modifiedTime, createdTime, owners(displayName), '
            'lastModifyingUser(displayName), md5Checksum, shared, version)'
        )

        try:
            while True:
                if self.is_stopped:
                    return
                while self.is_paused:
                    time.sleep(0.2)
                    if self.is_stopped:
                        return

                # 1. API Throttle: Sleep lightly right before we hit the API
                time.sleep(SLEEP_INTERVAL)

                # 2. Stoppable Retry Logic with Exponential Backoff
                retries = 0
                backoff = 2.0  # Start with 2 seconds on first rate-limit
                resp = None

                while retries <= MAX_API_RETRIES:
                    try:
                        resp = service.files().list(
                            q=query,
                            spaces='drive',
                            fields=fields,
                            pageToken=page_token,
                            pageSize=1000,  # <--- MAXIMIZED: 10x fewer API calls for large folders
                            includeItemsFromAllDrives=True,
                            supportsAllDrives=True
                        ).execute()
                        break  # Success, break out of retry loop

                    except HttpError as he:
                        status_code = he.resp.status
                        # 403 = Rate Limit / Quota Exceeded, 429 = Too Many Requests, 50x = Server Errors
                        if status_code in [403, 429, 500, 502, 503, 504]:
                            self.root.after(0, self.gui_log_insert, 
                                            f"⏳ Rate Limit hit ({status_code}). Retrying in {backoff}s... (Attempt {retries + 1}/{MAX_API_RETRIES})", 
                                            "warning")
                            
                            # Stoppable wait (so the user can still hit 'Stop' during a long backoff)
                            elapsed = 0
                            while elapsed < backoff:
                                if self.is_stopped:
                                    return
                                time.sleep(0.2)
                                elapsed += 0.2

                            retries += 1
                            backoff *= 2  # Double the wait time for the next attempt
                        else:
                            raise  # Unrecoverable error (e.g., 404 Not Found), bubble up

                if resp is None:
                    self.root.after(0, self.gui_log_insert, f"❌ Failed to list folder after {MAX_API_RETRIES} attempts. Skipping.", "warning")
                    return  # Give up on this specific folder branch

                items = resp.get('files', [])

                for idx, item in enumerate(items):
                    if self.is_stopped:
                        return
                    while self.is_paused:
                        time.sleep(0.2)
                        if self.is_stopped:
                            return

                    item_id = item.get('id')
                    mime = item.get('mimeType', '')
                    item_type = "Folder" if mime == 'application/vnd.google-apps.folder' else "File"
                    raw_name = item.get('name', '')
                    safe_name = self.sanitize_name(raw_name)

                    parent_path_only = parent_path if (parent_path and parent_path.endswith("/")) else (
                            parent_path + "/" if parent_path else "")
                    full_path = parent_path_only + safe_name
                    rel_path = parent_path_only

                    is_last = (idx == len(items) - 1)
                    parts = ["│ " if is_more else "  " for is_more in prefix_state]
                    parts.append("└─" if is_last else "├─")
                    tree_prefix = "".join(parts)

                    symbol = ICON_FOLDER if item_type == "Folder" else ICON_FILE
                    name_with_indent = f"{tree_prefix}{symbol} {raw_name}"

                    raw_size = 0
                    if item_type != "Folder":
                        try:
                            raw_size = int(item.get('size', 0) or 0)
                        except Exception:
                            raw_size = 0

                    size_str = self.format_size(raw_size) if item_type != "Folder" else "--"
                    date_str = self.format_date(item.get('modifiedTime'))
                    created_date_str = self.format_date(item.get('createdTime'))

                    item_owner = "N/A"
                    try:
                        owners = item.get('owners', [])
                        if owners:
                            item_owner = owners[0].get('displayName', 'N/A')
                    except IndexError:
                        pass

                    modified_by = item.get('lastModifyingUser', {}).get('displayName', 'N/A')
                    md5 = item.get('md5Checksum', '')
                    shared = item.get('shared', False)
                    version = int(item.get('version', 0))

                    self.root.after(0, self.append_log, name_with_indent, item_type, size_str, date_str, created_date_str, mime, level,
                                    rel_path, item_id, item_owner, modified_by, md5, shared, version)

                    excel_symbol = TREE_SYMBOL_FOLDER if item_type == 'Folder' else TREE_SYMBOL_FILE
                    indent_repr = tree_prefix + excel_symbol
                    size_out = '--' if item_type == 'Folder' else size_str

                    with self._data_lock:
                        self.data_rows.append(DataRow(
                            tree=indent_repr,
                            name=safe_name,
                            type=item_type,
                            size=size_out,
                            raw_size=raw_size,
                            date=date_str,
                            created_date=created_date_str,
                            mime=mime,
                            level=level,
                            parent=parent_path_only,
                            id=item_id,
                            owner=item_owner,
                            modified_by=modified_by,
                            md5=md5,
                            shared=shared,
                            version=version
                        ))

                        try:
                            self.parent_ids_by_child[item_id] = folder_id
                        except Exception:
                            pass

                        if item_type == "Folder":
                            self.total_folders += 1
                        else:
                            self.total_files += 1

                    self.root.after(0, self.gui_update_counters)

                    # 3. UI Throttle: Prevent Tkinter mainloop from freezing when processing huge folders.
                    # We only sleep briefly every 50 items rather than every single item.
                    if idx % 50 == 0:
                        time.sleep(0.01)

                    if item_type == "Folder":
                        new_state = prefix_state + [not is_last]
                        self.list_drive_folder_worker(service, item.get('id'), level + 1, full_path, new_state)

                page_token = resp.get('nextPageToken', None)
                if not page_token:
                    break

        except Exception as e:
            self.root.after(0, self.gui_log_insert, f"⚠ Error listing: {e}", "warning")


    def get_authenticated_service(self, creds_file, task_id=None):
        """
        Central function to get an authenticated Google Drive service.
        If already authenticated, returns existing session.
        """
        with self._data_lock:
            g_service = self.g_service
            auth_email = self.auth_email

        if g_service:
            self.root.after(0, self.gui_log_insert, "🔓 Authentication already established — using existing session.", "status")
            self.root.after(0, self.gui_log_insert, f"👤 Account: {auth_email}", "status")
            return g_service

        self.root.after(0, self.gui_log_insert, "⏳ Waiting for authentication in Browser...", "status")
        self.root.after(0, self.set_state, "auth_wait")

        service = None
        try:
            # 1. Attempt to load the credentials file specifically
            try:
                flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
            except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as file_err:
                self.root.after(0, self.gui_log_insert, "⚠ Authentication Failed: Invalid Credentials File.", "warning")
                self.root.after(0, self.gui_log_insert, "ℹ The file you selected is not a valid JSON file or is corrupted.", "info")
                self.root.after(0, self.gui_log_separator)
                self.root.after(0, self.set_state, "standby")
                return None

            # 2. Proceed with browser authentication
            creds = flow.run_local_server(port=0)

            # --- ZOMBIE CHECK ---
            if task_id is not None and task_id != self.current_task_id:
                return None  # Silent exit
            # --------------------

            if self.is_stopped:
                raise Exception("Authentication cancelled by user.")

            service = build('drive', 'v3', credentials=creds)
            self.root.after(0, self.gui_log_insert, "🔓 Authentication successful", "status")

            user_info = service.about().get(fields="user(emailAddress)").execute()

            with self._data_lock:
                self.auth_email = user_info["user"]["emailAddress"]
                self.g_service = service
                auth_email = self.auth_email

            self.root.after(0, self.gui_log_insert, f"👤 Account: {auth_email}", "status")
            return service

        except Exception as e:
            err_msg = str(e)
            
            if "access_denied" in err_msg:
                self.root.after(0, self.gui_log_insert, "⚠ Authentication cancelled by user.", "warning")
                
            elif "cancelled" in err_msg.lower():
                self.root.after(0, self.gui_log_insert, "⚠ Authentication cancelled.", "warning")
                
            elif "File validation failed" in err_msg:
                pass
                
            else:
                self.root.after(0, self.gui_log_insert, f"⚠ Authentication error: {e}", "warning")

            self.root.after(0, self.gui_log_separator)

            if self.is_stopped:
                self.root.after(0, self.set_state, "stopped")
            else:
                self.root.after(0, self.set_state, "standby")

            return None



    def open_drive_worker(self, creds_file, task_id):
        """Worker thread that only handles authenticating and opening the browser."""
        service = self.get_authenticated_service(creds_file, task_id)

        if service is None:
            # If service is None, it means either:
            # 1. An error occurred (already logged with separator in get_authenticated_service).
            # 2. It was a zombie task (task_id mismatch), so we die silently.
            return

        try:
            # Final Zombie Check before opening browser
            if task_id != self.current_task_id:
                return
            
            with self._data_lock:
                auth_email = self.auth_email

            url = f"https://drive.google.com/drive/my-drive?authuser={auth_email}"
            self.root.after(0, self.gui_log_insert, "🌐 Opening 'My Drive' using default web browser...", "status")
            self.root.after(0, lambda u=url: webbrowser.open(u, new=2))
            self.root.after(0, self.gui_log_separator)
        except Exception as e:
            self.root.after(0, self.gui_log_insert, f"⚠ Could not open browser: {e}", "warning")
            self.root.after(0, self.gui_log_separator)
        finally:
            # Only reset state if we are still effectively waiting and haven't been stopped/redirected
            if self.app_state == "auth_wait" and not self.is_stopped:
                self.root.after(0, self.set_state, "standby")
            # Reset the stopped flag if it was used
            if self.is_stopped:
                self.is_stopped = False
                self.root.after(0, self.set_state, "stopped")




    def acquisition_worker(self, creds_file, folder_id, task_id):
        """Main worker thread for folder listing."""
        with self._data_lock:
            self.total_files, self.total_folders = 0, 0
            self.data_rows = []
            self.is_paused, self.is_stopped = False, False
            self.acquisition_start = None
            self.folder_size_totals = defaultdict(int)
            self.parent_ids_by_child = {}
            self.is_live_scan_data = False

        service = self.get_authenticated_service(creds_file, task_id)

        if service is None:
            if self.is_stopped:
                self.is_stopped = False
            return

        # --- EXTRACT ID FROM URL (if present) ---
        is_url_extraction = False
        original_input = folder_id
        if "google.com" in folder_id and "/" in folder_id:
            match = re.search(r'(?:folders|d)\/([a-zA-Z0-9_-]+)', folder_id)
            if match:
                folder_id = match.group(1)
                is_url_extraction = True
            else:
                try:
                    last_part = folder_id.split('/')[-1]
                    new_id = last_part.split('?')[0].split('#')[0]
                    if new_id and new_id != folder_id:
                        folder_id = new_id
                        is_url_extraction = True
                except Exception:
                    pass

        # Convert 'my-drive' to 'root'
        if folder_id == "my-drive":
            folder_id = "root"

        if is_url_extraction and folder_id:
            self.root.after(0, self.gui_log_insert, f"ℹ️ Folder ID extracted from URL: {original_input}", "status")

            def _update_entry(fid=folder_id):
                self.folder_entry.delete(0, tk.END)
                self.folder_entry.insert(0, fid)

            self.root.after(0, _update_entry)

        # --- Proceed with folder validation ---
        try:
            fields = "id, name, mimeType, modifiedTime, createdTime, owners(displayName), lastModifyingUser(displayName), md5Checksum, shared, version"
            f = service.files().get(fileId=folder_id, fields=fields).execute()

            if f.get('mimeType') != 'application/vnd.google-apps.folder':
                self.root.after(0, self.gui_log_insert, f"⚠ The ID: {folder_id} is not associated with a folder.", "warning")
                self.root.after(0, self.gui_log_separator)
                self.root.after(0, self.set_state, "standby")
                return

            raw_name = f.get('name', 'folder')
            safe_name = self.sanitize_name(raw_name)
            self.selected_folder_name = safe_name
            root_mime = f.get('mimeType')
            root_modified = f.get('modifiedTime')
            root_created = f.get('createdTime')

            root_owner = "N/A"
            try:
                root_owner = f.get('owners', [{}])[0].get('displayName', 'N/A')
            except IndexError:
                pass

            root_mod_by = f.get('lastModifyingUser', {}).get('displayName', 'N/A')
            root_md5 = f.get('md5Checksum', '')
            root_shared = f.get('shared', False)
            root_version = int(f.get('version', 0))

        except HttpError as he:
            if he.resp.status in [401, 403]:
                self.root.after(0, self.gui_log_insert, f"⚠ Auth error. Please clear log and try again.", "warning")
                with self._data_lock:
                    self.g_service = None
            else:
                self.root.after(0, self.gui_log_insert, f"⚠ Invalid or inaccessible folder ID: {he}", "warning")
            self.root.after(0, self.gui_log_separator)
            self.root.after(0, self.set_state, "standby")
            return
        except Exception as e:
            self.root.after(0, self.gui_log_insert, f"⚠ Error validating folder: {e}", "warning")
            self.root.after(0, self.gui_log_separator)
            self.root.after(0, self.set_state, "standby")
            return

        self.root.after(0, self.gui_log_insert, f"✅ Folder ID validated ({folder_id})", "status")
        self.root.after(0, self.gui_log_insert, f"📁 Selected folder: {self.selected_folder_name}", "status")

        self.acquisition_start = datetime.now()
        self.root.after(0, self.gui_log_insert, "🚀 Data acquisition started", "status")

        def _insert_header():
            ts = self.ts_now()
            self.log_box.config(state="normal")
            self.log_box.insert(tk.END, f"[{ts}] ", "info")
            self.log_box.insert(tk.END, "📋 Showing: ", "status")
            self.log_box.insert(tk.END, "Tree with Name", "tree")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "Parent Path", "parent")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "Size", "size")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "Mod. Date", "date")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "Creat. Date", "created")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "MIME", "mime")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "Owner", "owner")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "Modified By", "modifier")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "MD5 Checksum", "md5")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "Shared", "shared")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "Version", "version")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "Level", "level")
            self.log_box.insert(tk.END, SEPARATOR, "info")
            self.log_box.insert(tk.END, "File/Folder ID", "id")
            self.log_box.insert(tk.END, "\n")
            self.log_box.config(state="disabled")
            if self.should_auto_scroll:
                self.log_box.see(tk.END)

        self.root.after(0, _insert_header)

        # --- Start the scan ---
        self.root.after(0, self.set_state, "running")

        indent_repr = TREE_SYMBOL_FOLDER
        mime_value = "root" if folder_id == "root" else (root_mime or "application/vnd.google-apps.folder")
        date_str_root = self.format_date(root_modified) if root_modified else "N/A"
        date_created_root = self.format_date(root_created) if root_created else "N/A"

        with self._data_lock:
            self.data_rows.append(DataRow(
                tree=indent_repr,
                name=safe_name,
                type="Folder",
                size="--",
                raw_size=0,
                date=date_str_root,
                created_date=date_created_root,
                mime=mime_value,
                level=0,
                parent="--",
                id=folder_id,
                owner=root_owner,
                modified_by=root_mod_by,
                md5=root_md5,
                shared=root_shared,
                version=root_version
            ))
            self.total_folders += 1
            try:
                self.parent_ids_by_child[folder_id] = None
            except Exception:
                pass

            self.is_live_scan_data = True
            self.data_account_email = self.auth_email

        self.root.after(0, self.append_log, ICON_FOLDER + " " + raw_name, "Folder", "--", date_str_root, date_created_root, mime_value, 0,
                        "--", folder_id, root_owner, root_mod_by, root_md5, root_shared, root_version)

        self.list_drive_folder_worker(service, folder_id, level=1, parent_path=self.selected_folder_name)

        # --- Finalize scan ---
        if self.is_stopped:
            self.root.after(0, self.gui_log_insert, "⛔ Listing interrupted by user.", "warning")
            final_state = "stopped"
            self.is_stopped = False
        else:
            self.root.after(0, self.gui_log_insert, "✅ Listing completed", "status")
            final_state = "standby"

        try:
            self.root.after(0, self._close_active_modal_if_any)
        except Exception:
            pass

        self.root.after(0, self.gui_log_insert, "🧮 Calculating raw folder sizes...", "status")
        self._calculate_folder_sizes()

        with self._data_lock:
            folders = self.total_folders
            files = self.total_files
            max_depth = max((r.level for r in self.data_rows), default=0)
            rows_count = len(self.data_rows)

        self.root.after(0, self.gui_log_insert, f"📂 Folders: {folders}   📄 Files: {files}", "status")
        self.root.after(0, self.gui_log_insert, f"📏 Maximum depth level: {max_depth}", "status")
        if self.acquisition_start:
            elapsed = str(datetime.now() - self.acquisition_start).split('.')[0]
            self.root.after(0, self.gui_log_insert, f"⏱ Acquisition time: {elapsed}", "status")

        self.root.after(0, self.gui_log_insert, f"💾 Ready to export {rows_count} rows of data", "status")
        self.root.after(0, self.gui_log_separator)
        self.root.after(0, self.set_state, final_state)



    def import_worker(self, file_path, file_name):
        """Worker thread to process imported JSON data and populate internal structures."""
        self.is_stopped = False
        self.root.after(0, self.set_state, "running")

        def _log_import_start():
            self.log_box.config(state="normal")
            self.log_box.insert(tk.END, f"[{self.ts_now()}] ", "status")
            self.log_box.insert(tk.END, "📥 Starting import process from \"", "status")
            self.log_box.insert(tk.END, file_name, "filename_highlight")
            self.log_box.insert(tk.END, "\"...\n", "status")
            self.log_box.config(state="disabled")
            if self.should_auto_scroll:
                self.log_box.see(tk.END)

        self.root.after(0, _log_import_start)
        self.root.after(0, self.gui_log_insert, "🔎 Checking file...", "status")

        def _clear_internal_data():
            with self._data_lock:
                self.data_rows = []
                self.folder_size_totals, self.parent_ids_by_child = defaultdict(int), {}
                self.total_files, self.total_folders = 0, 0
                self.has_unsaved_data = False
                self.acquisition_start = None
                self.data_account_email = ""

            self.folders_label.config(text="Folders: 0")
            self.files_label.config(text="Files: 0")

        items_processed = 0

        def _traverse_and_populate(nodes, parent_id=None):
            nonlocal items_processed
            for node in nodes:
                if self.is_stopped:
                    return
                item_type = node.get('type', 'File')
                item_id = node.get('id', '')
                owner = node.get('owner', 'N/A')

                with self._data_lock:
                    if item_id:
                        self.parent_ids_by_child[item_id] = parent_id

                    if item_type.lower() == 'folder':
                        self.total_folders += 1
                    else:
                        self.total_files += 1

                    self.data_rows.append(DataRow(
                        tree=node.get('tree_prefix', ''),
                        name=node.get('name', ''),
                        type=item_type,
                        size=node.get('size', '0 B'),
                        raw_size=int(node.get('raw_size', 0)),
                        date=node.get('modified_date', 'N/A'),
                        created_date=node.get('created_date', 'N/A'),
                        mime=node.get('mime_type', ''),
                        level=int(node.get('level', 0)),
                        parent=node.get('parent_path', ''),
                        id=item_id,
                        owner=owner,
                        modified_by=node.get('modified_by', 'N/A'),
                        md5=node.get('md5', ''),
                        shared=node.get('shared', False),
                        version=int(node.get('version', 0))
                    ))

                items_processed += 1
                if items_processed % 250 == 0:
                    self.root.after(0, self.gui_update_counters)

                if item_type.lower() == 'folder' and 'children' in node and node['children']:
                    _traverse_and_populate(node.get('children', []), parent_id=item_id)

        import_success = False

        try:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
            except json.JSONDecodeError:
                raise ValueError("Invalid JSON file. Could not parse.")
            except Exception as e:
                raise ValueError(f"Failed to read file: {e}")

            if not isinstance(json_data, dict):
                raise ValueError("Invalid JSON format. Expected a dictionary object.")
            if json_data.get('script_name') != SCRIPT_FULL_NAME:
                raise ValueError(f"Invalid '{SCRIPT_NAME}' JSON file.")
            tree_nodes = json_data.get('tree_data', [])
            if not tree_nodes:
                raise ValueError("JSON is valid, but 'tree_data' is empty.")

            imported_version = json_data.get('script_version', 'N/A')
            self.root.after(0, self.gui_log_insert, f"🆗 File OK (created with version: {imported_version}).", "status")

            imported_account = json_data.get('account_email', '(unavailable)')
            self.root.after(0, self.gui_log_insert, f"👤 Contains data from account: {imported_account}", "status")

            self.selected_folder_name = json_data.get('root_folder_name', 'imported_data')
            self.root.after(0, self.gui_log_insert, f"📂 Main folder: {self.selected_folder_name}", "status")

            with self._data_lock:
                has_data = bool(self.data_rows) or self.total_files > 0 or self.total_folders > 0

            if has_data:
                self.root.after(0, self.gui_log_insert, "🧹 Clearing previous data...", "status")
            _clear_internal_data()

            self.root.after(0, self.gui_log_insert, "🔧 Data structures prepared.", "status")
            self.root.after(0, self.gui_log_insert, "⏳ Importing data...", "status")

            with self._data_lock:
                self.data_account_email = imported_account
            try:
                acquisition_date_str = json_data.get('acquisition_date')
                self.acquisition_start = datetime.fromisoformat(
                    acquisition_date_str) if acquisition_date_str else None
            except (ValueError, TypeError):
                self.acquisition_start = None

            _traverse_and_populate(tree_nodes)

            self.root.after(0, self.gui_log_insert, "🧮 Recalculating raw folder sizes...", "status")
            self._calculate_folder_sizes()

            self.root.after(0, self.gui_update_counters)
            self.root.after(0, self.gui_log_insert, f"✅ {items_processed} rows imported successfully.", "status")
            self.root.after(0, self.gui_log_insert, "👍 Ready to export.", "success")
            self.root.after(0, self.gui_log_separator)

            with self._data_lock:
                self.has_unsaved_data = True
                self.is_live_scan_data = False

            import_success = True

        except ValueError as ve:
            self.root.after(0, self.gui_log_insert, f"❌ Error: {ve}", "warning")
            self.root.after(0, lambda ve=ve: custom_showerror(self.root, "Import Error", str(ve)))
            self.root.after(0, self.gui_log_separator)
        except Exception as e:
            self.root.after(0, self.gui_log_insert, f"💥 An unexpected error occurred during import: {e}", "warning")
            self.root.after(0, lambda e=e: custom_showerror(self.root, "Import Error", str(e)))
            self.root.after(0, self.gui_log_separator)
        finally:
            self.is_importing = False
            try:
                if import_success and getattr(self.root, "_active_modal", None) is not None:
                    self.root.after(0, self._close_active_modal_if_any)
            except Exception:
                pass

            self.root.after(0, self.set_state, "standby")

    # ---------- Help related Methods ----------
    def get_help_export_path(self):
        """Returns a path that all browsers (even Snaps) can definitely read."""
        # We use Downloads folders for html help file on macOS/Linux systems
        # because it's the most 'public' user folder for browsers
        return pathlib.Path.home() / "Downloads" / "TreeX4GD_Help_Files"



    def on_help(self):
        """Open help file. Direct for Windows, Downloads folder for macOS/Linux."""
        try:
            # 1. Resolve internal bundle path
            try:
                base_path = get_base_path()
            except NameError:
                base_path = os.getcwd()

            internal_help = pathlib.Path(base_path) / "TreeX4GD-help.html"
            internal_imgs = pathlib.Path(base_path) / "TreeX4GD-help-images"

            if not internal_help.exists():
                custom_showerror(self.root, "Error", "Help file not found in bundle.")
                return

            # 2. Major.Minor Version Check (using SCRIPT_VERSION and target-script-tier)
            try:
                with open(internal_help, "r", encoding="utf-8") as f:
                    head = f.read(1024)
                    
                match = re.search(r'<meta\s+name="target-script-tier"\s+content="([^"]+)"', head, re.IGNORECASE)
                if match:
                    found_tier = match.group(1).strip()
                    found_parts = [int(x) for x in found_tier.split('.') if x.isdigit()]
                    exp_parts = [int(x) for x in SCRIPT_VERSION.split('.') if x.isdigit()]
                    
                    if len(found_parts) >= 2 and len(exp_parts) >= 2:
                        # Compare Major and Minor ONLY
                        if found_parts[0] != exp_parts[0] or found_parts[1] != exp_parts[1]:
                            custom_showwarning(self.root, "Version Mismatch", 
                                f"The Help file is meant for script tier v{found_parts[0]}.{found_parts[1]}.x.\n\n"
                                f"The current app version is v{SCRIPT_VERSION}.\n"
                                "Some documentation might be inaccurate or missing.")
                else:
                    custom_showwarning(self.root, "Missing Metadata", 
                        "The Help file is missing its version metadata.\n\n" "That should not happen...")

            except Exception as e:
                print(f"Warning: Could not verify help file version: {e}")

            # 3. Windows: Open directly using os.startfile
            # This bypasses URI encoding issues with chars like », «, £, §, etc.
            if sys.platform.startswith('win'):
                try:
                    # Convert Path object to a standard string for Windows
                    os.startfile(str(internal_help.absolute()))
                except Exception:
                    # Fallback to webbrowser if startfile fails
                    webbrowser.open(internal_help.absolute().as_uri())
                return

            # 4. macOS/Linux: Copy to Downloads (Ensures Firefox/Chrome Snap access)
            dest_dir = self.get_help_export_path()
            dest_dir.mkdir(parents=True, exist_ok=True)
            
            dest_help = dest_dir / "TreeX4GD-help.html"
            dest_imgs = dest_dir / "TreeX4GD-help-images"

            # Copy files
            shutil.copy2(internal_help, dest_help)
            if internal_imgs.exists():
                if dest_imgs.exists():
                    shutil.rmtree(dest_imgs)
                shutil.copytree(internal_imgs, dest_imgs)
            
            # Open the html help file copy
            if sys.platform.startswith("linux"):
                # Open new window on Linux because of notification issues
                webbrowser.open(dest_help.absolute().as_uri(), new=1)
            else:  
                # Open new tab on macOS
                webbrowser.open(dest_help.absolute().as_uri(), new=2)

        except Exception as e:
            custom_showerror(self.root, "Error", f"Could not open help: {e}")



    # »»»»»»»»»»»»»»»»»» UI Callbacks ««««««««««««««««««««««««««

    # ---------- Folder ID Entry Mouse Focus Handlers ----------
    def on_folder_entry_mouse_down(self, event):
        """
        Mouse press on the entry. Mark that a mouse-driven focus is in progress.
        Do NOT return "break" here — allow Tk to perform normal focus/caret work.
        """
        try:
            if self.root.focus_get() is not self.folder_entry:
                self._folder_entry_mouse_focus_pending = True
                self.folder_entry.focus_set()
        except Exception:
            self._folder_entry_mouse_focus_pending = False
        return None

    def on_folder_entry_focus_in(self, event):
        """FocusIn handler."""
        try:
            if getattr(self, "_folder_entry_mouse_focus_pending", False):
                self._folder_entry_mouse_focus_pending = False
                self.folder_entry.after(5, self._select_all_folder_entry)
        except Exception:
            pass
        return None

    def _select_all_folder_entry(self):
        """Helper: ensure entry still has focus and then select-all."""
        try:
            if self.root.focus_get() is self.folder_entry:
                self.folder_entry.select_range(0, 'end')
                self.folder_entry.icursor('end')
        except Exception:
            pass

    def on_folder_entry_focus_out(self, event):
        """Handles widget losing focus."""
        try:
            self.folder_entry.select_clear()
        except Exception:
            pass

    # ---------- End of Folder ID Entry Mouse Focus Handlers ----------

    def on_browse_credentials(self):
        """Open file dialog to select JSON credentials file, handling active session conflicts."""
        # 1. Check if we have an active session
        with self._data_lock:
            has_session = (self.g_service is not None)

        # 2. If active, ask the user what to do
        if has_session:
            if not custom_asknegpos(
                self.root,
                "Active Authentication",
                "Current Authentication will be cleared before you can pick new credentials...",
                "Stay Authenticated",       # Negative answer (Left button)
                "Clear Auth. & Pick New"    # Positive answer (Right button)
            ):
                return
            
            # 3. User chose to clear: Reset service and UI
            with self._data_lock:
                self.g_service = None
                self.auth_email = ""
            
            self.gui_log_insert("🔓 Authentication session cleared by user.", "status")
            self.gui_log_separator()
            
            # Force a state refresh to re-enable the input box immediately
            self.set_state(self.app_state)

        # 4. Proceed with file selection
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if path:
            # Better enable the entry to write to it, just in case
            self.credentials_entry.config(state="normal") 
            self.credentials_entry.delete(0, tk.END)
            self.credentials_entry.insert(0, path)

    def on_open_drive(self):
        """Open Google Drive in the browser."""
        try:
            with self._data_lock:
                g_service = self.g_service
                auth_email = self.auth_email

            if g_service and auth_email:
                url = f"https://drive.google.com/drive/my-drive?authuser={auth_email}"
                webbrowser.open(url, new=2)
                return

            creds = self.credentials_entry.get().strip()
            if not creds:
                custom_showwarning(self.root, "Warning", "Please select a Credentials File first.")
                return
            if not os.path.exists(creds):
                custom_showerror(self.root, "Error", "Credentials file not found!")
                return
            
            self.is_stopped = False

            # Increment ID and pass to thread:
            self.current_task_id += 1    
            threading.Thread(target=self.open_drive_worker, args=(creds, self.current_task_id), daemon=True).start()

        except Exception as e:
            self.gui_log_insert(f"⚠ Could not open browser: {e}", "warning")

    def on_start_pause(self):
        """Unified handler for Start / Pause / Resume button."""
        if self.app_state == "running":
            self.is_paused = True
            self.set_state("paused")
            self.gui_log_insert("⏸ Listing paused.", "status")
            return

        if self.app_state == "paused":
            self.is_paused = False
            self.set_state("running")
            self.gui_log_insert("▶ Listing resumed.", "status")
            return

        creds = self.credentials_entry.get().strip()
        folder_id = self.folder_entry.get().strip()

        if not creds and not folder_id:
            custom_showwarning(self.root, "Warning",
                               "Credentials file and Google Drive folder ID not set.\n\nPlease fill in both fields!")
            return
        elif not creds:
            custom_showwarning(self.root, "Warning", "Credentials file not set.\n\nPlease select your credentials file!")
            return
        elif not folder_id:
            custom_showwarning(self.root, "Warning", "Google Drive folder ID not set.\n\nPlease enter the folder ID!")
            return

        if not os.path.exists(creds):
            custom_showerror(self.root, "Error", "Credentials file not found!")
            return

        with self._data_lock:
            has_data = bool(self.data_rows)
            is_live = self.is_live_scan_data
            has_unsaved = self.has_unsaved_data

        if has_data:
            if is_live and has_unsaved:
                if not custom_asknegpos(self.root, "Unexported Data",
                                        "Data from previous scan has not been exported.",
                                        "Go Back",
                                        "Discard and Continue"):
                    return

            with self._data_lock:
                self.data_rows.clear()
                self.folder_size_totals = defaultdict(int)
                self.has_unsaved_data = False
                self.is_live_scan_data = False

            self.gui_update_counters()
            self.gui_log_insert("🧹 Previous data discarded.", "status")

        # Increment ID and pass to thread:
        self.current_task_id += 1 
        threading.Thread(target=self.acquisition_worker, args=(creds, folder_id, self.current_task_id), daemon=True).start()

    def on_stop(self):
        """Stop the current listing after confirmation, without pausing the background task."""
        state_at_stop = self.app_state

        if state_at_stop not in ["running", "paused", "auth_wait"]:
            self.gui_log_insert("ℹ️ No active process to stop.", "status")
            self.gui_log_separator()
            return

        if state_at_stop == "paused":
            if not custom_asknegpos(
                    self.root,
                    "Stop Process",
                    "A task is currently in standby.",
                    "Remain Paused",
                    "Stop the task"
            ):
                return
        elif state_at_stop == "auth_wait":
            if not custom_asknegpos(
                    self.root,
                    "Stop Process",
                    "The app is waiting for authentication.",
                    "Keep Waiting",
                    "Stop Auth. Process"
            ):
                return
        else:  # running
            if not custom_asknegpos(
                    self.root,
                    "Stop Process",
                    "A task is currently running.",
                    "Keep it Running",
                    "Stop the task"
            ):
                return

        self.is_stopped = True
        self.is_paused = False
        self.current_task_id += 1  # <--- Invalidate any pending background auth tasks

        self.set_state("stopped")
        self.gui_log_insert("⏹ Stop request confirmed.", "status")

        if state_at_stop == "auth_wait":
            self.gui_log_separator()

    def on_set_root(self):
        """Confirm using Drive 'root' as the folder ID."""
        if custom_asknegpos(
                self.root,
                "Use Root Folder",
                "By setting the Folder ID to \"root\", this script will scan your entire\nGoogle Drive's \"My Drive\".",
                "Cancel",
                "Yes, Use Root"
        ):
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, "root")

    def on_save_log(self):
        """Save log output to a text file with a header."""
        self.root.after(0, self.set_state, 'exporting')

        try:
            default_name = self.selected_folder_name if self.selected_folder_name else "log"
            path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")],
                                                initialfile=f"{default_name}.txt")

            if not path:
                return

            # Header Generation
            now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            try:
                py_ver = sys.version.split()[0]
            except Exception:
                py_ver = "N/A"
            try:
                os_str = f"{platform.system()} {platform.release()}"
            except Exception:
                os_str = "N/A"

            WIDTH = 54
            TITLE = f"{SCRIPT_NAME} – Export Log"
            header_lines = [
                f"{'=' * 16} {TITLE} {'=' * 16}",
                f"{'Date:'.ljust(8)} {now_str}",
                f"{'Tool:'.ljust(8)} {SCRIPT_FULL_NAME}",
                f"{'Version:'.ljust(8)} {SCRIPT_VERSION}",
                f"{'OS:'.ljust(8)} {os_str}",
                f"{'Python:'.ljust(8)} {py_ver}",
                f"{'=' * WIDTH}",
                ""
            ]

            header = "\n".join(header_lines)

            log_content = self.log_box.get(1.0, tk.END).rstrip()
            if not log_content:
                num_lines = 0
            else:
                num_lines = len(log_content.split('\n'))

            self.gui_log_insert(f"📤 Exporting {num_lines} log lines to TXT file...", "status")

            with open(path, "w", encoding="utf-8") as f:
                f.write(header + "\n" + log_content)

            self.gui_log_insert(f"📝 Log saved to: {path}", "status")
            self.gui_log_separator()

        except Exception as e:
            custom_showerror(self.root, "Error", f"Failed to save log: {e}")
            self.gui_log_insert(f"⚠ Log save error: {e}", "warning")
            self.gui_log_separator()

        finally:
            self.root.after(0, self.set_state, 'standby')

    def on_clear_log(self):
        """Clear the log box and reset all session and data variables."""
        with self._data_lock:
            has_data = bool(self.data_rows)
            is_live = self.is_live_scan_data
            session_data_present = bool(self.g_service)

        message_parts = []

        if session_data_present:
            message_parts.append("• Current Authentication Session;")

        if has_data:
            if is_live:
                message_parts.append("• Recently scanned data;")
            else:
                message_parts.append("• Imported data;")

        final_message = ""
        if not message_parts:
            final_message = "The entire log will be cleared."
        else:
            message_parts.insert(0, "You are about to clear:")
            message_parts.append("• The entire log.")
            final_message = "\n".join(message_parts)

        if not custom_asknegpos(self.root, "Clear Log & Data", final_message, "Cancel", "Clear Data"):
            return

        # Clear GUI log box
        self.log_box.config(state="normal")
        self.log_box.delete(1.0, tk.END)
        self.log_box.config(state="disabled")

        with self._data_lock:
            self.data_rows = []
            self.total_files = 0
            self.total_folders = 0
            self.has_unsaved_data = False
            self.is_live_scan_data = False
            self.folder_size_totals = defaultdict(int)
            self.parent_ids_by_child = {}
            self.g_service = None
            self.auth_email = ""
            self.data_account_email = ""

        self.gui_update_counters()
        self.set_state("standby")

    def on_import_json(self):
        """Import data from a JSON file, replacing current data."""
        with self._data_lock:
            is_live = self.is_live_scan_data
            has_unsaved = self.has_unsaved_data

        if is_live and has_unsaved:
            if not custom_asknegpos(self.root, "Import Data",
                                    "You have unsaved data from previous scan.\nThe import process will discard it.",
                                    "Keep Current Data",
                                    "Import Anyway"):
                return
        path = filedialog.askopenfilename(title="Select JSON file to import", filetypes=[("JSON files", "*.json")])
        if not path:
            return

        self.is_stopped = False

        self.current_task_id += 1  # <--- Invalidate any pending background auth tasks
        self.is_importing = True
        threading.Thread(target=self.import_worker, args=(path, os.path.basename(path)), daemon=True).start()

    def on_exit(self):
        """Exit the program, with confirmation if listing is active."""
        if self.app_state in ["running", "auth_wait", "paused"]:
            if self.app_state == "paused":
                message = "A task is currently paused."
                neg_text = "Stay"
            elif self.app_state == "auth_wait":
                message = "The program is waiting for authentication."
                neg_text = "Keep Waiting"
            else:  # running
                message = "A task is currently running."
                neg_text = "Keep Running"

            response = custom_asknegpos(self.root, "Exit Confirmation", message, neg_text, "Exit Anyway")
            if not response:
                return
            self.is_stopped = True

        else:
            with self._data_lock:
                has_unsaved = self.has_unsaved_data
                is_live = self.is_live_scan_data

            if has_unsaved:
                if is_live:
                    message = "Data from the recent Google Drive scan has not been exported yet."
                    if not custom_asknegpos(self.root, "Unexported Data", message, "Go Back", "Exit Without Exporting"):
                        return
                else:
                    pass

        # --- CLEANUP HELP FILES on macOS/Linux BEFORE CLOSING ---
        try:
            if not sys.platform.startswith('win'):
                export_path = self.get_help_export_path()
                if export_path.exists():
                    shutil.rmtree(export_path)
        except Exception:
            pass # Don't let cleanup errors block the exit

        self.root.quit()
        self.root.destroy()

    def on_toggle_autoscroll(self):
        """Enable/disable automatic scrolling of the log window."""
        self.should_auto_scroll = self.auto_scroll_var.get()


# ---------- Main execution ----------
if __name__ == "__main__":
    
    root = tk.Tk(className="TreeX4GD") # Using className because of app icon in Linux systems
    app = TreeX4GDApp(root)

    # --- Fix for macOS window layering (move in front of terminal) ---
    root.lift()                          # Raise window to the top of the stack
    root.attributes('-topmost', True)    # Force it to be the "topmost" window
    root.after_idle(root.attributes, '-topmost', False) # Release the "topmost" lock once the app is idle
    root.focus_force()                   # Optional: Try to grab keyboard focus
    # -----------------------------------------------------------------

    root.mainloop()