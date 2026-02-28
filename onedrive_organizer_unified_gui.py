#!/usr/bin/env python3
"""
OneDrive File Organizer - Unified GUI
Microsoft 365 / Business with Device Code Auth and Archive Feature
Devoped by Jeff Lauria, iCorps Technologies Release 2/5/2025 
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import json
from pathlib import Path
from queue import Queue


class OneDriveOrganizerGUI:
    """Unified GUI with all features"""
    
    CLR_BG = "#f3f3f3"
    CLR_PRIMARY = "#0078d4"
    
    _CFG_PATH = Path.home() / ".onedrive_organizer_unified.json"
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OneDrive File Organizer - Microsoft 365 / Business")
        self.root.geometry("850x750")
        self.root.configure(bg=self.CLR_BG)
        
        # State
        self.anthropic_key = tk.StringVar()
        self.chk_large_files = tk.BooleanVar(value=False)
        self.chk_archive = tk.BooleanVar(value=False)
        self.chk_cleanup = tk.BooleanVar(value=False)
        self.chk_downloads = tk.BooleanVar(value=False)
        self.archive_months = tk.StringVar(value="24")
        self.status_var = tk.StringVar(value="Ready")
        self.device_code_var = tk.StringVar(value="")
        self.device_url_var = tk.StringVar(value="")
        
        # Backend
        self._organizer = None
        self.message_queue = Queue()
        
        # Load settings
        self._load_settings()
        
        # Build UI
        self._build_ui()
        
        # Start queue checker
        self._check_queue()
    
    def _build_ui(self):
        main = ttk.Frame(self.root, padding=20)
        main.pack(fill=tk.BOTH, expand=True)
        
        # Header
        ttk.Label(main, text="OneDrive File Organizer", 
                 font=("Segoe UI", 16, "bold")).pack(pady=(0, 5))
        ttk.Label(main, text="Microsoft 365 / Business - Full OneDrive Scan + Archive",
                 font=("Segoe UI", 9), foreground="#666").pack(pady=(0, 15))
        
        # API Key
        frame1 = ttk.LabelFrame(main, text="Step 1 – AI Configuration (Optional)", padding=10)
        frame1.pack(fill=tk.X, pady=(0, 10))
        
        row = ttk.Frame(frame1)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Anthropic API Key:", width=20).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self.anthropic_key, width=50, show="*").pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="(Leave blank for keyword-based sorting)", 
                 foreground="gray", font=("Segoe UI", 8)).pack(side=tk.LEFT)
        
        # Authentication
        frame2 = ttk.LabelFrame(main, text="Step 2 – Connect to OneDrive", padding=10)
        frame2.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(frame2, text="Sign in with your Microsoft 365 / Business account").pack(pady=5)
        
        # Device code display frame (initially hidden)
        self.device_frame = tk.Frame(frame2, bg=self.CLR_BG)
        
        code_display = tk.Frame(self.device_frame, bg="#fff", relief=tk.SOLID, borderwidth=1)
        code_display.pack(pady=10, padx=20, fill=tk.X)
        
        tk.Label(code_display, text="1. Visit:", font=("Segoe UI", 9), bg="#fff").pack(anchor=tk.W, padx=10, pady=(10, 0))
        url_entry = tk.Entry(code_display, textvariable=self.device_url_var, font=("Segoe UI", 10), 
                            state="readonly", relief=tk.FLAT, bg="#f9f9f9")
        url_entry.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(code_display, text="2. Enter this code:", font=("Segoe UI", 9), bg="#fff").pack(anchor=tk.W, padx=10, pady=(10, 0))
        code_entry = tk.Entry(code_display, textvariable=self.device_code_var, font=("Courier New", 14, "bold"),
                             state="readonly", justify=tk.CENTER, relief=tk.FLAT, bg="#fffacd")
        code_entry.pack(fill=tk.X, padx=10, pady=5)
        
        copy_btn = ttk.Button(code_display, text="📋 Copy Code", command=self._copy_code)
        copy_btn.pack(pady=(0, 10))
        
        self._auth_button = ttk.Button(frame2, text="🔐 Connect to OneDrive", 
                                       command=self._start_auth)
        self._auth_button.pack(pady=5)
        
        # Options
        frame3 = ttk.LabelFrame(main, text="Step 3 – Options", padding=10)
        frame3.pack(fill=tk.X, pady=(0, 10))
        
        tk.Checkbutton(frame3, text="Include large files (ISO, executables, disk images)",
                      variable=self.chk_large_files, bg=self.CLR_BG).pack(anchor=tk.W)
        
        # Archive option
        archive_frame = tk.Frame(frame3, bg=self.CLR_BG)
        archive_frame.pack(fill=tk.X, pady=(10, 0))
        
        tk.Checkbutton(archive_frame, text="Archive old files (not accessed in",
                      variable=self.chk_archive, bg=self.CLR_BG).pack(side=tk.LEFT)
        
        ttk.Entry(archive_frame, textvariable=self.archive_months, width=5).pack(side=tk.LEFT, padx=5)
        tk.Label(archive_frame, text="months)", bg=self.CLR_BG).pack(side=tk.LEFT)
        
        archive_info = tk.Label(frame3, 
                               text="Archives files/folders to 'Archive' folder, organized by year (e.g., Archive/2022/)",
                               bg=self.CLR_BG, foreground="#666", font=("Segoe UI", 8), wraplength=750, justify=tk.LEFT)
        archive_info.pack(anchor=tk.W, pady=(5, 0), padx=20)
        
        # Cleanup option
        tk.Checkbutton(frame3, text="Clean up empty folders after organizing",
                      variable=self.chk_cleanup, bg=self.CLR_BG).pack(anchor=tk.W, pady=(10, 0))
        
        cleanup_info = tk.Label(frame3,
                               text="Removes empty folders left behind after moving files (organized folders preserved)",
                               bg=self.CLR_BG, foreground="#666", font=("Segoe UI", 8), wraplength=750, justify=tk.LEFT)
        cleanup_info.pack(anchor=tk.W, pady=(5, 0), padx=20)
        
        # Downloads processing option
        tk.Checkbutton(frame3, text="Process Downloads folder (move files to root before organizing)",
                      variable=self.chk_downloads, bg=self.CLR_BG).pack(anchor=tk.W, pady=(10, 0))
        
        downloads_info = tk.Label(frame3,
                                 text="Moves all files from Downloads folder to root, then organizes them normally",
                                 bg=self.CLR_BG, foreground="#666", font=("Segoe UI", 8), wraplength=750, justify=tk.LEFT)
        downloads_info.pack(anchor=tk.W, pady=(5, 0), padx=20)
        
        # Organize
        frame4 = ttk.LabelFrame(main, text="Step 4 – Organize Files", padding=10)
        frame4.pack(fill=tk.X, pady=(0, 10))
        
        self._organize_button = ttk.Button(frame4, text="📁 Start Organizing", 
                                          command=self._start_organize, state=tk.DISABLED)
        self._organize_button.pack(pady=5)
        
        # Log
        frame5 = ttk.LabelFrame(main, text="Activity Log", padding=5)
        frame5.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self._log_text = scrolledtext.ScrolledText(frame5, height=12, 
                                                    font=("Consolas", 9), wrap=tk.WORD)
        self._log_text.pack(fill=tk.BOTH, expand=True)
        
        # Status
        status_frame = ttk.Frame(main)
        status_frame.pack(fill=tk.X)
        ttk.Label(status_frame, textvariable=self.status_var, 
                 font=("Segoe UI", 9), foreground="#666").pack(side=tk.LEFT)
        
        # Credits
        credits_frame = ttk.Frame(main)
        credits_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Separator(credits_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(0, 10))
        
        credits_text = tk.Label(credits_frame,
                               text="Application written and designed by Jeffery Lauria\nFor support: jeff@Brokencomputer.net",
                               bg=self.CLR_BG, foreground="#666", font=("Segoe UI", 8), justify=tk.CENTER)
        credits_text.pack()
    
    def _copy_code(self):
        """Copy device code to clipboard"""
        code = self.device_code_var.get()
        if code:
            self.root.clipboard_clear()
            self.root.clipboard_append(code)
            self._log("✓ Code copied to clipboard!")
            messagebox.showinfo("Copied", "Device code copied to clipboard!")
    
    # ------------------------------------------------------------------
    # Authentication
    # ------------------------------------------------------------------
    def _start_auth(self):
        self._auth_button.config(state=tk.DISABLED)
        self._set_status("Starting authentication...")
        self._log("Initiating device code authentication...")
        threading.Thread(target=self._auth_thread, daemon=True).start()
    
    def _auth_thread(self):
        try:
            from onedrive_organizer_unified import OneDriveOrganizerBusiness
            
            self._organizer = OneDriveOrganizerBusiness(
                anthropic_api_key=self.anthropic_key.get() or None,
                include_large_files=self.chk_large_files.get(),
                archive_old_files=self.chk_archive.get(),
                archive_months=int(self.archive_months.get()) if self.archive_months.get().isdigit() else 24,
                cleanup_empty_folders=self.chk_cleanup.get(),
                process_downloads=self.chk_downloads.get(),
            )
            self._organizer.log = lambda msg: self.message_queue.put(("log", msg))
            
            # Authenticate
            ok = self._organizer.authenticate()
            
            if ok:
                self.message_queue.put(("auth_success", None))
            else:
                self.message_queue.put(("auth_failed", None))
        
        except Exception as exc:
            self.message_queue.put(("error", f"Authentication error: {exc}"))
    
    def _show_device_code(self, url: str, code: str):
        """Display device code for user"""
        self.device_url_var.set(url)
        self.device_code_var.set(code)
        self.device_frame.pack(pady=10)
        self._log(f"Device code: {code}")
        self._log(f"Visit: {url}")
    
    def _on_auth_success(self):
        self._set_status("✓ Connected to OneDrive")
        self._log("✓ Authentication successful!")
        self.device_frame.pack_forget()  # Hide device code
        self._organize_button.config(state=tk.NORMAL)
        self._auth_button.config(text="✓ Connected", state=tk.DISABLED)
        messagebox.showinfo("Success", "Connected to your OneDrive!\n\nYou can now organize your files.")
    
    def _on_auth_failed(self):
        self._set_status("✗ Authentication failed")
        self._log("✗ Authentication failed")
        self.device_frame.pack_forget()  # Hide device code
        self._auth_button.config(state=tk.NORMAL)
        messagebox.showerror("Authentication Failed", 
                           "Could not connect to OneDrive.\n\nPlease try again.")
    
    # ------------------------------------------------------------------
    # Organization
    # ------------------------------------------------------------------
    def _start_organize(self):
        if not self._organizer:
            messagebox.showwarning("Not Connected", "Please connect to OneDrive first.")
            return
        
        # Build confirmation message
        msg_parts = [
            "This will scan and organize your ENTIRE OneDrive:\n",
            "• Scans ALL folders and subfolders recursively",
            "• Files will be analyzed and categorized",
            "• Organized folders created at root level",
            "• Items will be moved (not deleted)"
        ]
        
        if self.chk_downloads.get():
            msg_parts.append(f"\n📥 DOWNLOADS PROCESSING:")
            msg_parts.append(f"  All files in Downloads folder will be")
            msg_parts.append(f"  moved to root before organizing")
        
        if self.chk_archive.get():
            months = self.archive_months.get()
            msg_parts.append(f"\n⚠ ARCHIVE MODE ENABLED:")
            msg_parts.append(f"  Files/folders not accessed in {months}+ months")
            msg_parts.append(f"  will be moved to 'Archive' folder")
        
        if self.chk_cleanup.get():
            msg_parts.append(f"\n🧹 CLEANUP ENABLED:")
            msg_parts.append(f"  Empty folders will be deleted after organizing")
        
        msg_parts.append("\nProceed?")
        msg = "\n".join(msg_parts)
        
        if not messagebox.askyesno("Confirm", msg):
            return
        
        # Save settings
        self._save_settings()
        
        # Start organization
        self._organize_button.config(state=tk.DISABLED)
        self._set_status("Organizing files...")
        threading.Thread(target=self._organize_thread, daemon=True).start()
    
    def _organize_thread(self):
        try:
            self._organizer.organize_files()
            self.message_queue.put(("organize_complete", None))
        except Exception as exc:
            self.message_queue.put(("error", f"Organization error: {exc}"))
    
    def _on_organize_complete(self):
        self._set_status("✓ Organization complete!")
        self._organize_button.config(state=tk.NORMAL)
        messagebox.showinfo("Complete", 
                          "File organization complete!\n\nCheck your OneDrive to see the organized folders.")
    
    # ------------------------------------------------------------------
    # Message Queue Handler
    # ------------------------------------------------------------------
    def _check_queue(self):
        try:
            while True:
                msg_type, data = self.message_queue.get_nowait()
                
                if msg_type == "log":
                    self._log(data)
                    # Parse device code from log
                    if "Enter code:" in data:
                        code = data.split("Enter code:")[-1].strip()
                        self.device_code_var.set(code)
                    elif "Go to:" in data:
                        url = data.split("Go to:")[-1].strip()
                        self.device_url_var.set(url)
                        if self.device_code_var.get():
                            self.device_frame.pack(pady=10)
                elif msg_type == "auth_success":
                    self._on_auth_success()
                elif msg_type == "auth_failed":
                    self._on_auth_failed()
                elif msg_type == "organize_complete":
                    self._on_organize_complete()
                elif msg_type == "error":
                    self._log(f"ERROR: {data}")
                    self._set_status(f"Error: {data}")
                    messagebox.showerror("Error", str(data))
        except:
            pass
        
        self.root.after(100, self._check_queue)
    
    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _log(self, message: str):
        self._log_text.insert(tk.END, message + "\n")
        self._log_text.see(tk.END)
    
    def _set_status(self, message: str):
        self.status_var.set(message)
    
    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------
    def _load_settings(self):
        if self._CFG_PATH.exists():
            try:
                cfg = json.loads(self._CFG_PATH.read_text())
                self.anthropic_key.set(cfg.get("api_key", ""))
                self.chk_large_files.set(cfg.get("include_large_files", False))
                self.chk_archive.set(cfg.get("archive_old_files", False))
                self.archive_months.set(str(cfg.get("archive_months", 24)))
                self.chk_cleanup.set(cfg.get("cleanup_empty_folders", False))
                self.chk_downloads.set(cfg.get("process_downloads", False))
            except:
                pass
    
    def _save_settings(self):
        try:
            cfg = {
                "api_key": self.anthropic_key.get(),
                "include_large_files": self.chk_large_files.get(),
                "archive_old_files": self.chk_archive.get(),
                "archive_months": int(self.archive_months.get()) if self.archive_months.get().isdigit() else 24,
                "cleanup_empty_folders": self.chk_cleanup.get(),
                "process_downloads": self.chk_downloads.get(),
            }
            self._CFG_PATH.write_text(json.dumps(cfg))
        except:
            pass


def main():
    root = tk.Tk()
    app = OneDriveOrganizerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
