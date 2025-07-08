import subprocess
import os
import tkinter as tk
from tkinter import messagebox, END
from tkinter.scrolledtext import ScrolledText
import ttkbootstrap as ttk
from ttkbootstrap import Style
from ttkbootstrap.widgets import Button, Frame, Label
from ttkbootstrap.tableview import Tableview
from ttkbootstrap.constants import *
import threading
from datetime import datetime

MPCMDRUN = r"C:\Program Files\Windows Defender\MpCmdRun.exe"

# Prepare Recovered folder on Desktop
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
recovered_path = os.path.join(desktop_path, "Recovered")
os.makedirs(recovered_path, exist_ok=True)

class DefenderRestoreGUI:
    def __init__(self):
        self.app = ttk.Window(themename="darkly")
        self.app.title("🛡️ Windows Defender Quarantine Manager")
        self.app.geometry("1200x800")
        self.app.resizable(True, True)
        
        # Configure style
        self.style = Style("darkly")
        self.style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        self.style.configure("Subtitle.TLabel", font=("Segoe UI", 10))
        
        self.quarantine_items = []
        self.setup_gui()
        self.refresh_quarantine_list()
        
    def setup_gui(self):
        # Main container with padding
        main_container = ttk.Frame(self.app, padding=20)
        main_container.pack(fill=BOTH, expand=True)
        
        # Header section
        header_frame = ttk.Frame(main_container)
        header_frame.pack(fill=X, pady=(0, 20))
        
        title_label = ttk.Label(
            header_frame, 
            text="🛡️ Windows Defender Quarantine Manager",
            style="Title.TLabel",
            foreground="#4a9eff"
        )
        title_label.pack(anchor=W)
        
        subtitle_label = ttk.Label(
            header_frame,
            text="Safely restore quarantined files with automatic exclusions",
            style="Subtitle.TLabel",
            foreground="#888888"
        )
        subtitle_label.pack(anchor=W, pady=(5, 0))
        
        # Create main content with paned window for resizable sections
        paned_window = ttk.PanedWindow(main_container, orient=VERTICAL)
        paned_window.pack(fill=BOTH, expand=True)
        
        # Top section - Quarantine items
        top_frame = ttk.Frame(paned_window, padding=10)
        paned_window.add(top_frame, weight=2)
        
        # Quarantine items header with controls
        items_header = ttk.Frame(top_frame)
        items_header.pack(fill=X, pady=(0, 10))
        
        ttk.Label(
            items_header, 
            text="📋 Quarantined Items", 
            font=("Segoe UI", 12, "bold")
        ).pack(side=LEFT)
        
        # Control buttons
        button_frame = ttk.Frame(items_header)
        button_frame.pack(side=RIGHT)
        
        self.refresh_btn = ttk.Button(
            button_frame,
            text="🔄 Refresh",
            bootstyle="info-outline",
            command=self.refresh_quarantine_list_threaded
        )
        self.refresh_btn.pack(side=LEFT, padx=(0, 5))
        
        self.select_all_btn = ttk.Button(
            button_frame,
            text="☑️ Select All",
            bootstyle="secondary-outline",
            command=self.select_all_items
        )
        self.select_all_btn.pack(side=LEFT, padx=(0, 5))
        
        self.clear_selection_btn = ttk.Button(
            button_frame,
            text="☐ Clear Selection",
            bootstyle="secondary-outline",
            command=self.clear_selection
        )
        self.clear_selection_btn.pack(side=LEFT)
        
        # Table for quarantine items
        columns = [
            {"text": "Type", "width": 80},
            {"text": "File Path / Threat Name", "width": 400},
            {"text": "Status", "width": 100},
            {"text": "Details", "width": 300}
        ]
        
        self.tree = Tableview(
            master=top_frame,
            coldata=columns,
            searchable=True,
            autofit=True,
            height=15,
            paginated=True,
            pagesize=20
        )
        self.tree.pack(fill=BOTH, expand=True, pady=(0, 10))
        
        # Action buttons
        action_frame = ttk.Frame(top_frame)
        action_frame.pack(fill=X, pady=(10, 0))
        
        self.restore_btn = ttk.Button(
            action_frame,
            text="🔓 Restore Selected Items",
            bootstyle="success",
            command=self.restore_selected_threaded
        )
        self.restore_btn.pack(side=LEFT, padx=(0, 10))
        
        self.status_label = ttk.Label(
            action_frame,
            text="Ready",
            foreground="#10b981"
        )
        self.status_label.pack(side=LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            action_frame,
            mode='indeterminate',
            length=200
        )
        self.progress.pack(side=RIGHT)
        
        # Bottom section - Debug output
        bottom_frame = ttk.Frame(paned_window, padding=10)
        paned_window.add(bottom_frame, weight=1)
        
        debug_header = ttk.Frame(bottom_frame)
        debug_header.pack(fill=X, pady=(0, 10))
        
        ttk.Label(
            debug_header,
            text="🔍 Debug Output",
            font=("Segoe UI", 12, "bold")
        ).pack(side=LEFT)
        
        clear_debug_btn = ttk.Button(
            debug_header,
            text="🗑️ Clear Log",
            bootstyle="warning-outline",
            command=self.clear_debug_log
        )
        clear_debug_btn.pack(side=RIGHT)
        
        # Debug output with improved styling
        debug_frame = ttk.Frame(bottom_frame)
        debug_frame.pack(fill=BOTH, expand=True)
        
        self.debug_box = ScrolledText(
            debug_frame,
            height=10,
            font=("Consolas", 9),
            bg="#1a202c",
            fg="#e2e8f0",
            insertbackground="#4a9eff",
            selectbackground="#4a9eff",
            selectforeground="#ffffff"
        )
        self.debug_box.pack(fill=BOTH, expand=True)
        
        # Configure debug text tags for colored output
        self.debug_box.tag_configure("INFO", foreground="#10b981")
        self.debug_box.tag_configure("SUCCESS", foreground="#34d399")
        self.debug_box.tag_configure("WARNING", foreground="#f59e0b")
        self.debug_box.tag_configure("ERROR", foreground="#ef4444")
        self.debug_box.tag_configure("CMD", foreground="#8b5cf6")
        
        # Footer
        footer_frame = ttk.Frame(main_container)
        footer_frame.pack(fill=X, pady=(20, 0))
        
        ttk.Label(
            footer_frame,
            text=f"📁 Recovery Path: {recovered_path}",
            font=("Segoe UI", 9),
            foreground="#6b7280"
        ).pack(anchor=W)
        
    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"
        
        # Insert with appropriate color tag
        self.debug_box.insert(END, full_message, level)
        self.debug_box.see(END)
        self.app.update_idletasks()
        
    def clear_debug_log(self):
        self.debug_box.delete(1.0, END)
        self.log("Debug log cleared", "INFO")
        
    def update_status(self, message, color="#10b981"):
        self.status_label.config(text=message, foreground=color)
        self.app.update_idletasks()
        
    def list_quarantine_items(self):
        try:
            self.log("Retrieving quarantined items...", "INFO")
            output = subprocess.check_output([MPCMDRUN, "-Restore", "-ListAll"], text=True)
            
            entries = []
            for line in output.strip().splitlines():
                if line.strip() and not line.strip().startswith("Successfully"):
                    entries.append(line.strip())
            
            self.log(f"Found {len(entries)} quarantined items", "SUCCESS")
            return entries
            
        except subprocess.CalledProcessError as e:
            self.log(f"Could not retrieve quarantined items: {e}", "ERROR")
            return []
        except FileNotFoundError:
            self.log("Windows Defender MpCmdRun.exe not found. Please run as administrator.", "ERROR")
            return []
            
    def parse_quarantine_item(self, item):
        """Parse quarantine item and return structured data"""
        if item.lower().startswith("file:"):
            filepath = item[5:].split(" quarantined at")[0].strip()
            return {
                "type": "File",
                "path": filepath,
                "status": "Quarantined",
                "details": f"File: {os.path.basename(filepath)}"
            }
        elif "Path:" in item:
            filepath = item.split("Path:")[-1].strip()
            return {
                "type": "File",
                "path": filepath,
                "status": "Quarantined",
                "details": f"File: {os.path.basename(filepath)}"
            }
        elif item.lower().startswith("threatname ="):
            threat_name = item.split("=", 1)[-1].strip()
            return {
                "type": "Threat",
                "path": threat_name,
                "status": "Quarantined",
                "details": f"Threat: {threat_name}"
            }
        elif "Name:" in item:
            threat_name = item.split("Name:")[-1].strip()
            return {
                "type": "Threat",
                "path": threat_name,
                "status": "Quarantined",
                "details": f"Threat: {threat_name}"
            }
        else:
            return {
                "type": "Unknown",
                "path": item,
                "status": "Unknown",
                "details": "Unrecognized format"
            }
            
    def refresh_quarantine_list(self):
        self.tree.delete_rows()
        self.quarantine_items = self.list_quarantine_items()
        
        if not self.quarantine_items:
            self.tree.insert_row(END, ["No Items", "No quarantined items found", "N/A", ""])
            self.log("No items found in quarantine", "INFO")
        else:
            for i, item in enumerate(self.quarantine_items):
                parsed = self.parse_quarantine_item(item)
                self.tree.insert_row(END, [
                    parsed["type"],
                    parsed["path"],
                    parsed["status"],
                    parsed["details"]
                ])
            self.log(f"Loaded {len(self.quarantine_items)} quarantined items", "SUCCESS")
            
    def refresh_quarantine_list_threaded(self):
        self.progress.start()
        self.refresh_btn.config(state="disabled")
        self.update_status("Refreshing...", "#f59e0b")
        
        def refresh_thread():
            self.refresh_quarantine_list()
            self.app.after(0, self.refresh_complete)
            
        threading.Thread(target=refresh_thread, daemon=True).start()
        
    def refresh_complete(self):
        self.progress.stop()
        self.refresh_btn.config(state="normal")
        self.update_status("Ready", "#10b981")
        
    def select_all_items(self):
        if self.quarantine_items:
            # Select all rows in the tableview
            for i in range(len(self.quarantine_items)):
                self.tree.view.selection_add(self.tree.view.get_children()[i])
            self.log("Selected all items", "INFO")
            
    def clear_selection(self):
        self.tree.view.selection_remove(self.tree.view.selection())
        self.log("Cleared selection", "INFO")
        
    def get_selected_items(self):
        # Get selected items from the tableview
        selected_rows = self.tree.view.selection()
        if not selected_rows:
            return []
        
        selected_items = []
        for item_id in selected_rows:
            # Get the row index
            row_index = self.tree.view.index(item_id)
            if row_index < len(self.quarantine_items):
                selected_items.append((row_index, self.quarantine_items[row_index]))
        
        return selected_items
        
    def restore_selected_threaded(self):
        selected_items = self.get_selected_items()
        if not selected_items:
            messagebox.showwarning("No Selection", "Please select at least one item to restore.")
            return
            
        self.progress.start()
        self.restore_btn.config(state="disabled")
        self.update_status("Restoring items...", "#f59e0b")
        
        def restore_thread():
            self.restore_selected_items(selected_items)
            self.app.after(0, self.restore_complete)
            
        threading.Thread(target=restore_thread, daemon=True).start()
        
    def restore_selected_items(self, selected_items):
        restored_count = 0
        total_count = len(selected_items)
        
        self.log(f"Starting restoration of {total_count} items", "INFO")
        
        for i, (row_index, item) in enumerate(selected_items):
            self.log(f"Processing item {i+1}/{total_count}: {item}", "INFO")
            
            filepath = ""
            threat_name = ""
            
            if item.lower().startswith("file:"):
                filepath = item[5:].split(" quarantined at")[0].strip()
                cmd = [MPCMDRUN, "-Restore", "-FilePath", filepath, "-Path", recovered_path]
            elif "Path:" in item:
                filepath = item.split("Path:")[-1].strip()
                cmd = [MPCMDRUN, "-Restore", "-FilePath", filepath, "-Path", recovered_path]
            elif item.lower().startswith("threatname ="):
                threat_name = item.split("=", 1)[-1].strip()
                cmd = [MPCMDRUN, "-Restore", "-Name", threat_name]
            elif "Name:" in item:
                threat_name = item.split("Name:")[-1].strip()
                cmd = [MPCMDRUN, "-Restore", "-Name", threat_name]
            else:
                self.log(f"Unrecognized item format: {item}", "WARNING")
                continue
                
            self.log(f"Executing: {' '.join(cmd)}", "CMD")
            
            try:
                subprocess.run(cmd, shell=True, check=True)
                if filepath:
                    restored_file_path = os.path.join(recovered_path, os.path.basename(filepath))
                    self.log(f"Restored file to: {restored_file_path}", "SUCCESS")
                    self.exclude_file(restored_file_path)
                else:
                    self.log(f"Restored threat: {threat_name}", "SUCCESS")
                restored_count += 1
            except subprocess.CalledProcessError as e:
                self.log(f"Restore failed: {e}", "ERROR")
                
        self.log(f"Restoration complete: {restored_count}/{total_count} items restored", "SUCCESS")
        
    def restore_complete(self):
        self.progress.stop()
        self.restore_btn.config(state="normal")
        self.update_status("Ready", "#10b981")
        self.refresh_quarantine_list()
        messagebox.showinfo("Restoration Complete", 
                          f"Items have been restored to:\n{recovered_path}\n\nCheck the debug log for details.")
        
    def exclude_file(self, filepath):
        ps_command = f'Add-MpPreference -ExclusionPath "{filepath}"'
        self.log(f"Adding exclusion for: {filepath}", "INFO")
        
        try:
            subprocess.run(["powershell", "-Command", ps_command], check=True)
            self.log(f"Exclusion added successfully", "SUCCESS")
        except subprocess.CalledProcessError as e:
            self.log(f"Failed to add exclusion: {e}", "ERROR")
            
    def run(self):
        self.log("Windows Defender Quarantine Manager started", "INFO")
        self.log(f"Recovery path: {recovered_path}", "INFO")
        self.app.mainloop()

if __name__ == "__main__":
    app = DefenderRestoreGUI()
    app.run()