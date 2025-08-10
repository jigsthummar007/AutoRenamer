# main.py - Auto File Renamer (Final Pro Version with Help & About)
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
from pathlib import Path
import csv
import json
import logging
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import re
import webbrowser
import requests  # For checking online version
import webbrowser  # To open GitHub download link

# ============ Version Info ============
APP_NAME = "Auto File Renamer Pro"
VERSION = "1.5.0"
AUTHOR = "Jignesh Thummar"
COPYRIGHT = "© 2025 Jignesh Thummar. All Rights Reserved."
LICENSE = "Proprietary Software. Do not distribute."

# ============ Create Directories ============
project_dir = Path(__file__).parent
logs_dir = project_dir / "logs"
codes_dir = project_dir / "codes"
config_dir = project_dir / "config"
backup_dir = project_dir / "backup"
logs_dir.mkdir(exist_ok=True)
codes_dir.mkdir(exist_ok=True)
config_dir.mkdir(exist_ok=True)
backup_dir.mkdir(exist_ok=True)

# ============ Keyword Config Path ============
keywords_file = config_dir / "keywords.json"
DEFAULT_KEYWORDS = ["copy", "copies", "pcs", "pieces", "x"]

# ============ Configure Logging ============
log_file = logs_dir / f"{datetime.now().strftime('%Y-%m-%d')}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# ============ Set Appearance ============
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# ============ Fonts ============
FONT = ("Segoe UI", 12)
TITLE_FONT = ("Segoe UI", 16, "bold")
CODE_FONT = ("Consolas", 13)
SMALL_FONT = ("Segoe UI", 10)


class RenameHistory:
    def __init__(self):
        self.history = []
        self.index = -1

    def add(self, old, new):
        self.history = self.history[:self.index + 1]
        self.history.append({
            "old": str(old),
            "new": str(new),
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        self.index += 1

    def undo(self):
        if self.index < 0:
            return None
        item = self.history[self.index]
        self.index -= 1
        return item

    def redo(self):
        if self.index >= len(self.history) - 1:
            return None
        self.index += 1
        return self.history[self.index]

    def clear(self):
        self.history = []
        self.index = -1


class FileRenamerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{VERSION}")
        self.geometry("1000x780")
        self.resizable(True, True)

        # --- Paths ---
        self.config_file = project_dir / "config.json"

        # --- Data ---
        self.selected_root = None
        self.selected_file = None
        self.allowed_extensions = {'.plt', '.jpg', '.jpeg', '.jpe', '.jfif'}
        self.party_map = {}
        self.history = RenameHistory()
        self.auto_observer = None
        self.file_path_list = []
        self.filtered_file_list = []  # For search
        self.machine_var = ctk.StringVar(value="(C.S)")
        self.show_done_var = ctk.BooleanVar(value=False)
        self.last_folder = ""
        self.quantity_keywords = []
        self.first_run = True

                # ============ Version & Update Config ============
        self.CURRENT_VERSION = "1.5.0"
        self.UPDATE_URL = "https://raw.githubusercontent.com/jigsthummar007/AutoRenamer/main/version.txt"

        # --- Load Config ---
        self.load_config()
        self.load_keywords()

        # --- Menu Bar ---
        self.create_menu_bar()

        # --- Startup Wizard ---
        self.after(100, self.run_startup_wizard)

        # ============ Grid Layout ============
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # ============ Sidebar ============
        self.sidebar_frame = ctk.CTkFrame(self, width=260, corner_radius=12, fg_color="#2a2d30")
        self.sidebar_frame.grid(row=0, column=0, rowspan=7, sticky="nswe", padx=12, pady=12)
        self.sidebar_frame.grid_propagate(False)

        ctk.CTkLabel(self.sidebar_frame, text="📁 File Manager", font=TITLE_FONT).pack(pady=(12, 8))

        self.folder_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="📁 Select 2025 Folder",
            command=self.select_folder,
            height=40,
            font=FONT
        )
        self.folder_btn.pack(pady=8, padx=16, fill="x")

        self.scan_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="⟳ Scan Files",
            command=self.scan_folder,
            height=35
        )
        self.scan_btn.pack(pady=6, padx=16, fill="x")

        # --- Auto-scan ---
        self.auto_scan_var = ctk.BooleanVar(value=True)
        self.auto_scan_switch = ctk.CTkSwitch(
            self.sidebar_frame,
            text="🔁 Auto-scan",
            variable=self.auto_scan_var,
            command=self.toggle_auto_scan,
            font=FONT
        )
        self.auto_scan_switch.pack(pady=10, padx=16, anchor="w")

        # --- Show Done Files ---
        self.show_done_switch = ctk.CTkSwitch(
            self.sidebar_frame,
            text="👁️ Show Finalize Files",
            variable=self.show_done_var,
            command=self.scan_folder,
            font=FONT
        )
        self.show_done_switch.pack(pady=8, padx=16, anchor="w")

        # --- Machine Type ---
        ctk.CTkLabel(self.sidebar_frame, text="🖨️ Machine:", font=FONT).pack(pady=(12, 4), anchor="w", padx=20)
        self.machine_dropdown = ctk.CTkComboBox(
            self.sidebar_frame,
            values=["(C.S)", "(C.E)"],
            variable=self.machine_var,
            font=FONT,
            dropdown_font=FONT
        )
        self.machine_dropdown.pack(pady=6, padx=20, fill="x")

        # --- Edit Keywords Button ---
        self.keywords_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="🔧 Edit Quantity Keywords",
            command=self.open_keywords_editor,
            fg_color="orange",
            hover_color="dark orange",
            height=32
        )
        self.keywords_btn.pack(pady=8, padx=16, fill="x")

        # --- Reload CSV ---
        self.reload_csv_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="🔁 Reload Parties",
            command=self.load_parties_csv,
            fg_color="#8a2be2",
            hover_color="#7a1dd1",
            height=32
        )
        self.reload_csv_btn.pack(pady=12, padx=16, fill="x")

        # --- Export Log Button ---
        self.export_log_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="📊 Export Rename Log",
            command=self.export_rename_log,
            fg_color="teal",
            hover_color="dark teal",
            height=32
        )
        self.export_log_btn.pack(pady=8, padx=16, fill="x")

        # ============ Main Area ============
        self.main_frame = ctk.CTkFrame(self, corner_radius=12)
        self.main_frame.grid(row=0, column=1, sticky="nswe", padx=12, pady=12)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(3, weight=1)

        # --- Search Box ---
        search_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        search_frame.grid(row=0, column=0, columnspan=2, sticky="we", pady=(0, 5))
        ctk.CTkLabel(search_frame, text="🔍 Search:", font=FONT).pack(side="left")
        self.search_var = ctk.StringVar()
        self.search_var.trace("w", self.on_search_change)
        self.search_entry = ctk.CTkEntry(search_frame, textvariable=self.search_var, placeholder_text="Filter files...")
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(5, 0))

        # --- Selected File ---
        self.file_label = ctk.CTkLabel(self.main_frame, text="No file selected", font=TITLE_FONT)
        self.file_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 5))

        # --- Preview ---
        self.preview_label = ctk.CTkLabel(
            self.main_frame,
            text="Preview: --",
            font=CODE_FONT,
            text_color="lightgray",
            justify="left",
            anchor="w"
        )
        self.preview_label.grid(row=2, column=0, columnspan=2, pady=10, sticky="w")

        # --- Buttons ---
        btn_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        btn_frame.grid(row=3, column=0, columnspan=2, sticky="w", pady=5)

        self.rename_btn = ctk.CTkButton(btn_frame, text="✅ Rename", command=self.rename_file, width=100)
        self.rename_btn.grid(row=0, column=0, padx=(0, 10))

        self.undo_btn = ctk.CTkButton(btn_frame, text="↩ Undo", command=self.undo_rename, width=90)
        self.undo_btn.grid(row=0, column=1, padx=(0, 5))

        self.redo_btn = ctk.CTkButton(btn_frame, text="⟳ Redo", command=self.redo_rename, width=90)
        self.redo_btn.grid(row=0, column=2)

        self.undo_all_btn = ctk.CTkButton(btn_frame, text="↩ Undo All", command=self.undo_all_batch, width=90, fg_color="red")
        self.undo_all_btn.grid(row=0, column=3, padx=(5, 0))

        # --- File List ---
        self.file_listbox = ctk.CTkTextbox(self.main_frame, font=CODE_FONT, wrap="none")
        self.file_listbox.grid(row=4, column=0, columnspan=2, sticky="nswe", pady=5)
        self.file_listbox.bind("<ButtonRelease-1>", self.on_file_click)
        self.file_listbox.configure(cursor="hand2")

        # --- Select All Button ---
        self.select_all_btn = ctk.CTkButton(
            self.main_frame,
            text="📁 Batch Rename",
            command=self.select_all_files,
            fg_color="green",
            hover_color="dark green",
            height=36
        )
        self.select_all_btn.grid(row=5, column=0, columnspan=2, sticky="we", padx=10, pady=(6, 0))

        # --- Status Bar ---
        self.status_label = ctk.CTkLabel(self, text="Ready", anchor="w", font=FONT)
        self.status_label.grid(row=6, column=0, columnspan=2, sticky="we", padx=16, pady=4)

        # --- Footer ---
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.grid(row=7, column=0, columnspan=2, pady=(0, 10), sticky="s")

        ctk.CTkLabel(footer_frame, text="📱", font=SMALL_FONT).pack(side="left", padx=(0, 2))
        whatsapp_btn = ctk.CTkButton(
            footer_frame, text="WhatsApp", width=80, height=20, font=SMALL_FONT,
            command=lambda: webbrowser.open("https://wa.me/919825531314")
        )
        whatsapp_btn.pack(side="left", padx=2)

        ctk.CTkLabel(footer_frame, text="📸", font=SMALL_FONT).pack(side="left", padx=(10, 2))
        insta_btn = ctk.CTkButton(
            footer_frame, text="@official.jignesh.1", width=120, height=20, font=SMALL_FONT,
            command=lambda: webbrowser.open("https://instagram.com/official.jignesh.1")
        )
        insta_btn.pack(side="left", padx=2)

        ctk.CTkLabel(footer_frame, text="✉️", font=SMALL_FONT).pack(side="left", padx=(10, 2))
        email_btn = ctk.CTkButton(
            footer_frame, text="Email Me", width=80, height=20, font=SMALL_FONT,
            command=lambda: webbrowser.open("mailto:Jigsthummar1990@gmail.com")
        )
        email_btn.pack(side="left", padx=2)

        # ============ Keyboard Shortcuts ============
        self.bind("<Control-o>", lambda e: self.select_folder())
        self.bind("<Control-s>", lambda e: self.scan_folder())
        self.bind("<Control-r>", lambda e: self.rename_file() if self.selected_file else None)
        self.bind("<Control-z>", lambda e: self.undo_rename())
        self.bind("<Control-y>", lambda e: self.redo_rename())
        self.bind("<Control-a>", lambda e: self.select_all_files())

        # --- Auto-scan init ---
        self.toggle_auto_scan()

        # --- Load last folder ---
        if self.last_folder and Path(self.last_folder).exists():
            self.selected_root = Path(self.last_folder)
            self.scan_folder()
            self.start_auto_scan()

        self.load_parties_csv()

        # --- Create Backup (must come after status_label is created)
        self.create_backup()
                # 🔔 Check for update 2 seconds after launch
        self.after(2000, self.check_for_update)

    def create_menu_bar(self):
        """Create menu bar with Help and About"""
        menubar = tk.Menu(self)
        
        # Help Menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Usage Guide", command=self.show_help_usage)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_help_shortcuts)
        menubar.add_cascade(label="📘 Help", menu=help_menu)

        # About Menu
        menubar.add_command(label="ℹ️ About", command=self.show_about)

        self.configure(menu=menubar)

    def show_help_shortcuts(self):
        """Show keyboard shortcuts"""
        msg = """
📌 Keyboard Shortcuts:

• Ctrl + O → Select Folder
• Ctrl + S → Scan Files
• Ctrl + R → Rename Selected File
• Ctrl + Z → Undo Last Rename
• Ctrl + Y → Redo Rename
• Ctrl + A → Batch Rename All Files

💡 Tip: Click any file to preview rename.
💡 Tip: Use 'Show Done Files' to finalize with %8%.
"""
        messagebox.showinfo("📘 Help: Keyboard Shortcuts", msg)

    def show_help_usage(self):
        """Show how to use the app"""
        msg = """
📘 How to Use Auto File Renamer:

1. Click '📁 Select 2025 Folder' to set root
2. Files will appear in the list
3. Click any file to see rename preview
4. Click '✅ Rename' to process
5. File moves to 'Done' folder automatically
6. Use '📁 Batch Rename' for multiple files
7. Use '🔧 Edit Keywords' to add 'layout', 'design', etc.
8. Use '👁️ Show Finalize Files' to manually add %8%

📁 Folder Structure:
2025 → Month → Date → Party → File.plt

📤 Output:
{Code}_{Name} (C.S)(FT.1x2)(Q.2)%%.plt → Moved to 'Done'

🔧 Keywords:
- Use 'x' → '2 x' → (Q.2)
- Add more via UI
"""
        messagebox.showinfo("📘 Help: Usage Guide", msg)

    def show_about(self):
        """Show about window"""
        popup = ctk.CTkToplevel(self)
        popup.title("ℹ️ About")
        popup.geometry("450x500")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        x = self.winfo_x() + (self.winfo_width() // 2) - 225
        y = self.winfo_y() + (self.winfo_height() // 2) - 250
        popup.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(popup, text=APP_NAME, font=("Segoe UI", 18, "bold")).pack(pady=10)
        ctk.CTkLabel(popup, text=f"Version {VERSION}", font=("Segoe UI", 14)).pack(pady=5)
        ctk.CTkLabel(popup, text=f"By: {AUTHOR}", font=("Segoe UI", 14)).pack(pady=5)

        ctk.CTkLabel(popup, text="📱 Contact:", font=("Segoe UI", 12, "bold")).pack(pady=(20, 5))
        ctk.CTkLabel(popup, text="WhatsApp: +91 98255 31314", font=("Segoe UI", 12)).pack()
        ctk.CTkLabel(popup, text="Instagram: @official.jignesh.1", font=("Segoe UI", 12)).pack()
        ctk.CTkLabel(popup, text="Email: Jigsthummar1990@gmail.com", font=("Segoe UI", 12)).pack()

        ctk.CTkLabel(popup, text="📜 License", font=("Segoe UI", 12, "bold")).pack(pady=(20, 5))
        ctk.CTkLabel(popup, text=COPYRIGHT, font=("Segoe UI", 10), wraplength=400).pack(pady=5)
        ctk.CTkLabel(popup, text="Proprietary Software. Do not distribute.", font=("Segoe UI", 10), wraplength=400).pack(pady=5)

        ctk.CTkButton(
            popup, text="📧 Send Email", width=120,
            command=lambda: webbrowser.open("mailto:Jigsthummar1990@gmail.com")
        ).pack(pady=10)

        ctk.CTkButton(popup, text="Close", width=100, command=popup.destroy).pack(pady=10)


    def check_for_update(self):
        """Check if a new version is available"""
        try:
            response = requests.get(self.UPDATE_URL, timeout=5)
            latest_version = response.text.strip()

            if latest_version > self.CURRENT_VERSION:
                self.show_update_prompt(latest_version)
            else:
                self.status_label.configure(text=f"✅ Up to date (v{self.CURRENT_VERSION})")

        except requests.RequestException as e:
            logging.warning(f"Update check failed: {e}")
            self.status_label.configure(text=f"✅ Ready (v{self.CURRENT_VERSION}) - Update check failed")
        except Exception as e:
            logging.error(f"Unexpected error in update check: {e}")
            self.status_label.configure(text=f"✅ Ready (v{self.CURRENT_VERSION})")

    def show_update_prompt(self, latest_version):
        """Show update popup"""
        popup = ctk.CTkToplevel(self)
        popup.title("📢 Update Available!")
        popup.geometry("400x180")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        # Center popup
        popup.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (popup.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (popup.winfo_height() // 2)
        popup.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(popup, text=f"New Version {latest_version} Available!", font=("Helvetica", 16, "bold")).pack(pady=10)
        ctk.CTkLabel(popup, text=f"You are on v{self.CURRENT_VERSION}", font=("Helvetica", 12)).pack(pady=5)

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=15)

        ctk.CTkButton(
            btn_frame,
            text="📥 Download Now",
            command=lambda: webbrowser.open("https://github.com/jigsthummar007/AutoRenamer/releases"),
            fg_color="green"
        ).pack(side="left", padx=5)

        ctk.CTkButton(
            btn_frame,
            text="Later",
            command=popup.destroy
        ).pack(side="left", padx=5)

        self.status_label.configure(text=f"🔔 Update v{latest_version} available!")    

    def run_startup_wizard(self):
        """Show setup guide on first run"""
        if not self.first_run:
            return

        wizard = ctk.CTkToplevel(self)
        wizard.title("🚀 Setup Wizard")
        wizard.geometry("500x400")
        wizard.resizable(False, False)
        wizard.transient(self)
        wizard.grab_set()

        x = self.winfo_x() + (self.winfo_width() // 2) - 250
        y = self.winfo_y() + (self.winfo_height() // 2) - 200
        wizard.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(wizard, text="Welcome to Auto File Renamer", font=TITLE_FONT).pack(pady=10)
        ctk.CTkLabel(wizard, text="First-time setup guide:", font=FONT).pack(pady=(0, 10))

        steps = [
            "1. Click '📁 Select 2025 Folder' to set root",
            "2. Your files will appear in the list",
            "3. Click any file to preview rename",
            "4. Use '🔧 Edit Keywords' to add 'layout', 'design', etc.",
            "5. Click '✅ Rename' to process",
            "6. Use '📁 Batch Rename' for multiple files"
        ]

        for step in steps:
            ctk.CTkLabel(wizard, text=step, font=("Segoe UI", 11), anchor="w").pack(pady=2, padx=20, anchor="w")

        def finish():
            self.first_run = False
            wizard.destroy()

        ctk.CTkButton(wizard, text="Let's Go!", command=finish, height=40, font=FONT).pack(pady=20)

    def create_backup(self):
        """Backup config, parties.csv, keywords.json"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_folder = backup_dir / timestamp
        backup_folder.mkdir(exist_ok=True)

        items = [
            (self.config_file, "config.json"),
            (codes_dir / "parties.csv", "parties.csv"),
            (keywords_file, "keywords.json")
        ]

        for src, name in items:
            try:
                if src.exists():
                    shutil.copy2(src, backup_folder / name)
            except Exception as e:
                logging.warning(f"Backup failed for {name}: {e}")

        self.status_label.configure(text=f"✅ Backup created: {timestamp}")

    def export_rename_log(self):
        """Export rename history to CSV"""
        if not self.history.history:
            messagebox.showinfo("Export Log", "No rename history to export.")
            return

        file = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            title="Save Rename Log"
        )
        if not file:
            return

        try:
            with open(file, mode='w', encoding='utf-8', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(["Timestamp", "Original", "New Name"])
                for item in self.history.history:
                    writer.writerow([item["timestamp"], item["old"], item["new"]])
            messagebox.showinfo("Success", f"Rename log exported to:\n{file}")
        except Exception as e:
            messagebox.showerror("Export Failed", str(e))

    def load_keywords(self):
        """Load quantity keywords from config/keywords.json"""
        if keywords_file.exists():
            try:
                with open(keywords_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.quantity_keywords = data.get("quantity_keywords", DEFAULT_KEYWORDS)
            except Exception as e:
                logging.warning(f"Failed to load keywords: {e}")
                self.quantity_keywords = DEFAULT_KEYWORDS
        else:
            self.quantity_keywords = DEFAULT_KEYWORDS
            self.save_keywords()

    def save_keywords(self):
        """Save keywords to file"""
        try:
            with open(keywords_file, "w", encoding="utf-8") as f:
                json.dump({"quantity_keywords": self.quantity_keywords}, f, indent=2)
        except Exception as e:
            logging.error(f"Failed to save keywords: {e}")

    def open_keywords_editor(self):
        """Open UI to edit quantity keywords"""
        popup = ctk.CTkToplevel(self)
        popup.title("🔧 Edit Quantity Keywords")
        popup.geometry("500x400")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()

        x = self.winfo_x() + (self.winfo_width() // 2) - 250
        y = self.winfo_y() + (self.winfo_height() // 2) - 200
        popup.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(popup, text="Manage Quantity Keywords", font=TITLE_FONT).pack(pady=10)

        list_frame = ctk.CTkScrollableFrame(popup, height=200)
        list_frame.pack(pady=10, padx=20, fill="both", expand=True)

        keyword_vars = []
        delete_btns = []

        def refresh_list():
            for var, btn in zip(keyword_vars, delete_btns):
                var.destroy()
                btn.destroy()
            keyword_vars.clear()
            delete_btns.clear()

            for kw in self.quantity_keywords:
                inner_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
                inner_frame.pack(fill="x", pady=2)

                var = ctk.CTkLabel(inner_frame, text=kw, font=CODE_FONT, width=100, anchor="w")
                var.pack(side="left", padx=(0, 10))
                keyword_vars.append(var)

                btn = ctk.CTkButton(
                    inner_frame,
                    text="🗑️",
                    width=40,
                    height=28,
                    fg_color="#d42222",
                    hover_color="#a00",
                    command=lambda k=kw: remove_keyword(k, popup)
                )
                btn.pack(side="right")
                delete_btns.append(btn)

        def remove_keyword(kw, win):
            self.quantity_keywords.remove(kw)
            refresh_list()

        def add_keyword():
            new_kw = add_entry.get().strip().lower()
            if new_kw and new_kw not in self.quantity_keywords:
                self.quantity_keywords.append(new_kw)
                add_entry.delete(0, "end")
                refresh_list()

        def save_and_close():
            self.save_keywords()
            popup.destroy()
            self.status_label.configure(text=f"✅ Keywords updated: {len(self.quantity_keywords)} active")

        def reset_default():
            self.quantity_keywords = DEFAULT_KEYWORDS.copy()
            refresh_list()

        refresh_list()

        add_frame = ctk.CTkFrame(popup)
        add_frame.pack(pady=10, padx=20, fill="x")
        ctk.CTkLabel(add_frame, text="Add New:").pack(side="left")
        add_entry = ctk.CTkEntry(add_frame, placeholder_text="e.g., layout, design")
        add_entry.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkButton(add_frame, text="Add", command=add_keyword).pack(side="left")

        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="💾 Save & Close", command=save_and_close).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="🔄 Reset", command=reset_default).pack(side="left", padx=5)

        popup.bind("<Return>", lambda e: add_keyword())
        popup.focus()

    def extract_dimensions(self, filename: str) -> str:
        """Extract dimensions and convert to feet using custom bucket rules"""
        clean = re.sub(r'\s+', ' ', filename.replace('X', 'x').lower())
        match = re.search(r'(\d+\.?\d*)\s*x\s*(\d+\.?\d*)', clean)
        if not match:
            return ""
        w_in, h_in = float(match.group(1)), float(match.group(2))

        def to_feet(inch):
            if inch <= 20: return 1
            elif inch <= 26: return 2
            elif inch <= 32: return 3
            elif inch <= 38: return 3
            elif inch <= 50: return 4
            elif inch <= 62: return 5
            elif inch <= 74: return 6
            elif inch <= 98: return 8
            else: return 10

        w_ft = to_feet(w_in)
        h_ft = to_feet(h_in)

        # Enforce minimum 2 sq ft → (1x2)
        if w_ft * h_ft < 2:
            w_ft, h_ft = 1, 2

        return f"{w_ft}x{h_ft}"

    def detect_quantity(self, filename: str) -> int:
        """
        Detect quantity only from numbers adjacent to keywords.
        Ignores dimensions like '60 x 36'.
        """
        clean = re.sub(r'\s+', ' ', filename.strip().lower())
        # Remove dimension part first
        clean = re.sub(r'\d+\s*x\s*\d+', '', clean)
        clean = re.sub(r'\s+', ' ', clean).strip()

        for kw in self.quantity_keywords:
            # Match: "2 x", "3 copy", "4   pcs", etc.
            match = re.search(rf'\b(\d+)\s*{re.escape(kw)}\b', clean)
            if match:
                return int(match.group(1))
            # Match: "x 2", "copy 3", etc.
            match = re.search(rf'\b{re.escape(kw)}\s*(\d+)\b', clean)
            if match:
                return int(match.group(1))
        return 1  # Default

    def generate_new_filename(self, original_stem: str, party_code: str, extension: str, dim_str: str = "") -> str:
        qty = self.detect_quantity(original_stem)
        machine = self.machine_var.get()
        dim_part = f"(FT.{dim_str})" if dim_str else ""
        return f"{party_code}_{original_stem} {machine}{dim_part}(Q.{qty})%%{extension}"

    def update_preview(self):
        if not self.selected_file or not self.selected_file.exists():
            self.preview_label.configure(text="Preview: --", text_color="gray")
            return
        stem = self.selected_file.stem
        ext = self.selected_file.suffix
        party_name = self.selected_file.parent.name
        code = self.party_map.get(party_name, "?")
        color = "red" if code == "?" else "lightgreen"
        dim = self.extract_dimensions(self.selected_file.name)
        new_name = self.generate_new_filename(stem, code, ext, dim)
        self.preview_label.configure(text=f"Preview: {new_name}", text_color=color)

    def select_folder(self):
        try:
            folder = filedialog.askdirectory(title="Select Year Folder (e.g. 2025)")
            if not folder: return
            path = Path(folder)
            if not path.is_dir(): return
            self.selected_root = path
            self.last_folder = str(folder)
            self.save_config()
            self.scan_folder()
            self.start_auto_scan()
            self.status_label.configure(text=f"📁 Active: {path.name}")
        except Exception as e:
            self.status_label.configure(text="❌ Folder selection failed")
            messagebox.showerror("Error", f"Could not open folder:\n{e}")

    def scan_folder(self):
        if not self.selected_root or not self.selected_root.exists():
            self.file_listbox.delete("0.0", "end")
            self.file_listbox.insert("0.0", "❌ Root folder not found.")
            return
        try:
            files = []
            for ext in self.allowed_extensions:
                if self.show_done_var.get():
                    files.extend(self.selected_root.rglob(f"Done/*{ext}"))
                    files.extend(self.selected_root.rglob(f"Done/*{ext.upper()}"))
                else:
                    files.extend(self.selected_root.rglob(f"*{ext}"))
                    files.extend(self.selected_root.rglob(f"*{ext.upper()}"))
            files = sorted(set(files), key=lambda x: x.name.lower())
            self.file_listbox.delete("0.0", "end")
            mode = "Finalize Mode" if self.show_done_var.get() else "New Files"
            self.file_listbox.insert("0.0", f"📁 {mode}\nFiles:\n")
            self.file_path_list = []
            for file_path in files:
                in_done = "done" in [p.lower() for p in file_path.parts]
                if in_done != self.show_done_var.get(): continue
                if "[ok]" in file_path.name and self.show_done_var.get(): continue
                party = file_path.parent.name if not in_done else "Done"
                code = self.party_map.get(party, "?") if not in_done else "?"
                display = f"{code} | {party} | {file_path.name}\n"
                self.file_listbox.insert("end", display)
                self.file_path_list.append(file_path)
            self.filtered_file_list = self.file_path_list[:]
            if self.show_done_var.get():
                self.select_all_btn.grid_remove()
            else:
                self.select_all_btn.grid()
            self.status_label.configure(text=f"✅ {mode}: {len(self.file_path_list)} files")
        except Exception as e:
            self.status_label.configure(text=f"❌ Scan error: {e}")

    def on_search_change(self, *args):
        query = self.search_var.get().strip().lower()
        self.file_listbox.delete("0.0", "end")
        self.file_listbox.insert("0.0", "Search Results:\n")
        self.filtered_file_list = []
        for file_path in self.file_path_list:
            if query in file_path.name.lower():
                party = file_path.parent.name
                code = self.party_map.get(party, "?")
                display = f"{code} | {party} | {file_path.name}\n"
                self.file_listbox.insert("end", display)
                self.filtered_file_list.append(file_path)
        self.status_label.configure(text=f"🔍 Found {len(self.filtered_file_list)} matching files")

    def on_file_click(self, event):
        try:
            index = self.file_listbox.index(f"@{event.x},{event.y}")
            line_num = int(index.split('.')[0])
            HEADER_LINES = 3
            file_index = line_num - HEADER_LINES
            if file_index < 0 or file_index >= len(self.filtered_file_list): return
            selected_path = self.filtered_file_list[file_index]
            if not selected_path.exists(): return
            self.selected_file = selected_path.resolve()
            if self.show_done_var.get():
                self.open_manual_finalize_popup()
                return
            self.file_label.configure(text=f"Selected: {selected_path.name}")
            self.update_preview()
            self.status_label.configure(text=f"📁 {selected_path.parent.name} | Ready")
        except Exception as e:
            self.status_label.configure(text=f"❌ Click error: {e}")

    def open_manual_finalize_popup(self):
        files_left = [p for p in self.filtered_file_list if "[ok]" not in p.name]
        try:
            current_idx = files_left.index(self.selected_file)
        except ValueError:
            current_idx = 0

        def finalize_and_next(qty, cat):
            try:
                old_path = self.selected_file
                if "[ok]" in old_path.name: return
                stem = old_path.stem
                ext = old_path.suffix
                if "(Q." in stem:
                    base = re.split(r'\(Q\.\d+\)%%?', stem)[0]
                else:
                    base = stem
                machine = self.machine_var.get()
                if machine in base and "[ok]" not in base:
                    base = base.replace(machine, f"{machine}[ok]")
                new_name = f"{base}(Q.{qty})"
                new_name += f"%{cat}%" if cat else "%%"
                new_name += ext
                new_path = old_path.parent / new_name
                if new_path.exists():
                    messagebox.showerror("Exists", "File already exists.")
                    return
                old_path.rename(new_path)
                self.history.add(old_path, new_path)
                self.status_label.configure(text=f"✅ Finalized: {new_path.name}")
                if current_idx < len(files_left) - 1:
                    next_file = files_left[current_idx + 1]
                    self.selected_file = next_file
                    self.after(100, self.open_manual_finalize_popup)
                else:
                    self.scan_folder()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        self.show_manual_input_popup(finalize_and_next)

    def show_manual_input_popup(self, callback):
        popup = ctk.CTkToplevel(self)
        popup.title("✏️ Finalize File")
        popup.geometry("380x180")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grab_set()
        x = self.winfo_x() + (self.winfo_width() // 2) - 190
        y = self.winfo_y() + (self.winfo_height() // 2) - 90
        popup.geometry(f"+{int(x)}+{int(y)}")
        ctk.CTkLabel(popup, text="Enter Final Details", font=TITLE_FONT).pack(pady=10)
        frame = ctk.CTkFrame(popup, fg_color="transparent")
        frame.pack(pady=8)
        ctk.CTkLabel(frame, text="Qty:").pack(side="left", padx=5)
        qty_var = ctk.StringVar(value="1")
        qty_entry = ctk.CTkEntry(frame, textvariable=qty_var, width=60)
        qty_entry.pack(side="left", padx=10)
        qty_entry.focus()
        ctk.CTkLabel(frame, text="Cat:").pack(side="left", padx=5)
        cat_var = ctk.StringVar()
        cat_entry = ctk.CTkEntry(frame, textvariable=cat_var, width=60)
        cat_entry.pack(side="left", padx=10)
        btn_frame = ctk.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(pady=10)
        def submit():
            qty = qty_var.get().strip()
            if not qty.isdigit() or int(qty) < 1:
                messagebox.showwarning("Invalid", "Quantity must be ≥ 1")
                return
            callback(qty, cat_var.get().strip())
            popup.destroy()
        def cancel(): popup.destroy()
        ctk.CTkButton(btn_frame, text="✅ OK", command=submit, width=80).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=cancel, width=80).pack(side="left", padx=5)
        popup.bind("<Return>", lambda e: submit())
        popup.bind("<Escape>", lambda e: cancel())
        qty_entry.bind("<Return>", lambda e: cat_entry.focus())
        cat_entry.bind("<Return>", lambda e: submit())

    def rename_file(self):
        if not self.selected_file: return
        machine = self.machine_var.get()
        if not machine: return
        file_path = self.selected_file
        if not file_path.exists(): return
        party_name = file_path.parent.name
        code = self.party_map.get(party_name)
        if not code: return
        stem = file_path.stem
        ext = file_path.suffix
        dim = self.extract_dimensions(file_path.name)
        new_name = self.generate_new_filename(stem, code, ext, dim)
        new_path = file_path.parent / new_name
        counter = 1
        original = new_path
        while new_path.exists():
            new_path = original.parent / f"{original.stem} ({counter}){ext}"
            counter += 1
        try:
            file_path.rename(new_path)
            done_folder = file_path.parent / "Done"
            done_folder.mkdir(exist_ok=True)
            final_path = done_folder / new_path.name
            shutil.move(str(new_path), str(final_path))
            self.history.add(file_path, final_path)
            self.status_label.configure(text=f"✅ Renamed: {final_path.name}")
            self.scan_folder()
            self.selected_file = None
            self.file_label.configure(text="No file selected")
            self.preview_label.configure(text="Preview: --", text_color="gray")
        except Exception as e:
            self.status_label.configure(text=f"❌ Error: {e}")

    def select_all_files(self):
        if not self.filtered_file_list or self.show_done_var.get(): return
        machine = self.machine_var.get()
        if not machine: return
        confirm = messagebox.askyesno("Confirm", "Rename all files?")
        if not confirm: return
        renamed_count = 0
        error_count = 0
        for file_path in self.filtered_file_list[:]:
            try:
                if not file_path.exists(): continue
                party_name = file_path.parent.name
                code = self.party_map.get(party_name)
                if not code: continue
                stem = file_path.stem
                ext = file_path.suffix
                dim = self.extract_dimensions(file_path.name)
                new_name = self.generate_new_filename(stem, code, ext, dim)
                new_path = file_path.parent / new_name
                counter = 1
                while new_path.exists():
                    new_path = file_path.parent / f"{new_name[:-len(ext)]} ({counter}){ext}"
                    counter += 1
                file_path.rename(new_path)
                done_folder = file_path.parent / "Done"
                done_folder.mkdir(exist_ok=True)
                final_path = done_folder / new_path.name
                shutil.move(str(new_path), str(final_path))
                self.history.add(file_path, final_path)
                renamed_count += 1
            except Exception as e:
                error_count += 1
        self.scan_folder()
        self.status_label.configure(text=f"✅ Batch: {renamed_count} renamed, {error_count} failed")

    def undo_all_batch(self):
        """Undo all renames from last batch"""
        if not self.history.history:
            messagebox.showinfo("Undo", "Nothing to undo.")
            return
        confirm = messagebox.askyesno("Undo All", "Undo ALL renames from this session?")
        if not confirm:
            return
        restored = 0
        while self.history.history:
            item = self.history.undo()
            if not item:
                break
            try:
                src = Path(item["new"])
                dst = Path(item["old"])
                dst.parent.mkdir(exist_ok=True)
                shutil.move(str(src), str(dst))
                restored += 1
            except Exception as e:
                logging.error(f"Undo failed: {e}")
        self.history.clear()
        self.scan_folder()
        self.status_label.configure(text=f"↩ Undo All: {restored} files restored")

    def undo_rename(self):
        item = self.history.undo()
        if not item: return
        try:
            src = Path(item["new"])
            dst = Path(item["old"])
            dst.parent.mkdir(exist_ok=True)
            shutil.move(str(src), str(dst))
            self.status_label.configure(text=f"↩ Undo: {dst.name}")
            self.scan_folder()
        except Exception as e:
            self.status_label.configure(text=f"❌ Undo failed: {e}")

    def redo_rename(self):
        item = self.history.redo()
        if not item: return
        try:
            src = Path(item["old"])
            dst = Path(item["new"])
            counter = 1
            orig = dst
            while dst.exists():
                dst = orig.parent / f"{orig.stem} ({counter}){orig.suffix}"
                counter += 1
            src.rename(dst)
            done = dst.parent / "Done"
            done.mkdir(exist_ok=True)
            shutil.move(str(dst), str(done / dst.name))
            self.status_label.configure(text=f"⟳ Redo: {dst.name}")
            self.scan_folder()
        except Exception as e:
            self.status_label.configure(text=f"❌ Redo failed: {e}")

    def load_parties_csv(self):
        csv_path = codes_dir / "parties.csv"
        if not csv_path.exists():
            try:
                with open(csv_path, "w", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow(["Party Name", "Code"])
                    writer.writerow(["Creative", "2"])
                    writer.writerow(["Pranam Maheta", "7"])
                    writer.writerow(["XYZ Designs", "5"])
                    writer.writerow(["Sunrise", "3"])
                    writer.writerow(["Vikas", "9"])
                self.status_label.configure(text=f"✅ Created default CSV: {csv_path.name}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create CSV: {e}")
                self.party_map = {}
                return
        self.party_map = {}
        try:
            with open(csv_path, mode='r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    name = row.get("Party Name", "").strip()
                    code = row.get("Code", "").strip()
                    if name and code:
                        self.party_map[name] = code
            if self.selected_root:
                self.scan_folder()
        except Exception as e:
            self.status_label.configure(text="❌ Failed to load CSV")
            messagebox.showerror("Error", f"Failed to load parties.csv:\n{e}")

    def load_config(self):
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.last_folder = data.get("last_folder", "")
        except Exception:
            self.last_folder = ""

    def save_config(self):
        try:
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump({"last_folder": self.last_folder}, f, indent=2)
        except Exception as e:
            logging.error(f"Config save error: {e}")

    def toggle_auto_scan(self):
        if self.auto_scan_var.get() and self.selected_root:
            self.start_auto_scan()
        else:
            self.stop_auto_scan()

    def start_auto_scan(self):
        self.stop_auto_scan()
        if not self.selected_root: return
        class Handler(FileSystemEventHandler):
            def __init__(self, app):
                self.app = app
            def on_created(self, event):
                if not event.is_directory:
                    ext = Path(event.src_path).suffix.lower()
                    if ext in self.app.allowed_extensions:
                        self.app.after(0, self.app.scan_folder)
        self.auto_observer = Observer()
        self.auto_observer.schedule(Handler(self), str(self.selected_root), recursive=True)
        self.auto_observer.start()

    def stop_auto_scan(self):
        if self.auto_observer:
            self.auto_observer.stop()
            self.auto_observer.join()
            self.auto_observer = None

    def destroy(self):
        self.stop_auto_scan()
        self.save_config()
        super().destroy()


if __name__ == "__main__":
    app = FileRenamerApp()
    app.mainloop()