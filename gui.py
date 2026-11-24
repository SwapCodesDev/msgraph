"""Single-file graphical user interface (GUI) for Microsoft Graph Excel Integration.

Provides an intuitive desktop UI to:
- Connect and authenticate via Device Code or Client Credentials.
- Inspect, create, and open the remote Excel file in Excel Online.
- View worksheets and browse cells in an interactive spreadsheet grid.
- Append rows and update cell ranges without writing commands.
- Edit and save .env configuration directly in the app.
"""

import json
import os
import sys
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, simpledialog, ttk
from typing import Any, Dict, List, Optional
import webbrowser
from pathlib import Path
from auth import GraphAuthenticator
from config import BASE_DIR, Config
from graph_client import GraphAPIError, GraphExcelClient


# Enable High-DPI support on Windows if available
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass


class DeviceCodeDialog(tk.Toplevel):
    """Modal dialog displaying Microsoft Device Code instructions with quick copy and open."""

    def __init__(self, parent, flow_data):
        super().__init__(parent)
        self.flow_data = flow_data
        self.title("Microsoft Authentication Required")
        self.geometry("520x340")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self._build_ui()
        self.center_window(parent)

    def center_window(self, parent):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w = self.winfo_width()
        h = self.winfo_height()
        x = px + max(0, (pw - w) // 2)
        y = py + max(0, (ph - h) // 2)
        self.geometry(f"+{x}+{y}")

    def _build_ui(self):
        container = ttk.Frame(self, padding=20)
        container.pack(fill=tk.BOTH, expand=True)

        lbl_title = ttk.Label(
            container,
            text="Sign in to Microsoft 365",
            font=("Segoe UI", 13, "bold"),
        )
        lbl_title.pack(anchor="w", pady=(0, 10))

        lbl_desc = ttk.Label(
            container,
            text="To grant access, copy your one-time code below and paste it in the login window:",
            wraplength=480,
            justify="left",
        )
        lbl_desc.pack(anchor="w", pady=(0, 15))

        # Code display box
        code_frame = ttk.Frame(container)
        code_frame.pack(fill=tk.X, pady=(0, 15))

        user_code = self.flow_data.get("user_code", "UNKNOWN")
        self.txt_code = ttk.Entry(
            code_frame,
            font=("Consolas", 18, "bold"),
            justify="center",
        )
        self.txt_code.insert(0, user_code)
        self.txt_code.configure(state="readonly")
        self.txt_code.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        btn_copy = ttk.Button(code_frame, text="📋 Copy Code", command=self.copy_code)
        btn_copy.pack(side=tk.RIGHT)

        # Action buttons
        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill=tk.X, pady=(10, 15))

        btn_open = ttk.Button(
            btn_frame,
            text="🌐 Open Login Page in Browser",
            command=self.open_browser,
        )
        btn_open.pack(fill=tk.X, ipady=4)

        self.lbl_status = ttk.Label(
            container,
            text="⏳ Waiting for you to complete sign-in in browser...",
            font=("Segoe UI", 9, "italic"),
            foreground="#666666",
        )
        self.lbl_status.pack(pady=(10, 0))

    def copy_code(self):
        code = self.flow_data.get("user_code", "")
        self.clipboard_clear()
        self.clipboard_append(code)
        messagebox.showinfo("Copied", f"Code '{code}' copied to clipboard!", parent=self)

    def open_browser(self):
        url = self.flow_data.get("verification_uri", "https://microsoft.com/devicelogin")
        webbrowser.open(url)


class ExcelGraphGUI(tk.Tk):
    """Main application window for Microsoft Graph Excel Manager."""

    def __init__(self):
        super().__init__()
        self.title("Microsoft Graph Excel Manager")
        self.geometry("960x720")
        self.minsize(860, 600)

        self.client: Optional[GraphExcelClient] = None
        self.active_device_dialog: Optional[DeviceCodeDialog] = None
        self.current_web_url: Optional[str] = None

        self._configure_styles()
        self._build_ui()
        self._load_config_to_fields()

        # Log startup banner
        self.log("Ready. Click 'Connect / Check Status' to initialize connection to Microsoft Graph.")

    def _configure_styles(self):
        self.style = ttk.Style(self)
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.style.configure(".", font=("Segoe UI", 9))
        self.style.configure("Header.TLabel", font=("Segoe UI", 10, "bold"))
        self.style.configure("Status.TLabel", font=("Segoe UI", 9))
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _build_ui(self):
        # ---------------------------------------------------------------------
        # Top Header Bar (Connection Status & Quick Actions)
        # ---------------------------------------------------------------------
        top_bar = ttk.Frame(self, padding=(15, 10))
        top_bar.pack(fill=tk.X)

        lbl_app = ttk.Label(top_bar, text="📊 Excel Graph Manager", font=("Segoe UI", 13, "bold"))
        lbl_app.pack(side=tk.LEFT)

        self.lbl_auth_status = ttk.Label(
            top_bar,
            text="⚪ Status: Disconnected",
            foreground="#666666",
            font=("Segoe UI", 9, "bold"),
        )
        self.lbl_auth_status.pack(side=tk.LEFT, padx=(20, 0))

        self.btn_open_web = ttk.Button(
            top_bar,
            text="🌐 Open in Excel Online",
            command=self.open_excel_online,
            state="disabled",
        )
        self.btn_open_web.pack(side=tk.RIGHT, padx=(5, 0))

        self.btn_connect = ttk.Button(
            top_bar,
            text="🔌 Connect / Check Status",
            command=self.on_connect_clicked,
        )
        self.btn_connect.pack(side=tk.RIGHT, padx=(5, 0))

        # ---------------------------------------------------------------------
        # Main Tabbed Notebook
        # ---------------------------------------------------------------------
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 10))

        # Tab 1: Spreadsheet Viewer
        self.tab_view = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_view, text="  📈 Spreadsheet Grid  ")
        self._build_tab_view()

        # Tab 2: Write & Append
        self.tab_write = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_write, text="  ✏️ Update & Append  ")
        self._build_tab_write()

        # Tab 3: Configuration & Auth
        self.tab_config = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.tab_config, text="  ⚙️ Settings & Azure Auth  ")
        self._build_tab_config()

        # ---------------------------------------------------------------------
        # Bottom Activity Console Log
        # ---------------------------------------------------------------------
        log_frame = ttk.LabelFrame(self, text=" Activity Log ", padding=5)
        log_frame.pack(fill=tk.X, padx=15, pady=(0, 10))

        self.txt_log = scrolledtext.ScrolledText(
            log_frame,
            height=6,
            font=("Consolas", 9),
            wrap=tk.WORD,
        )
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        self.txt_log.tag_config("ERROR", foreground="#d32f2f")
        self.txt_log.tag_config("SUCCESS", foreground="#2e7d32")
        self.txt_log.tag_config("INFO", foreground="#1976d2")

    # -------------------------------------------------------------------------
    # Tab 1: Spreadsheet Viewer UI
    # -------------------------------------------------------------------------
    def _build_tab_view(self):
        ctrl_frame = ttk.Frame(self.tab_view)
        ctrl_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(ctrl_frame, text="Sheet:").pack(side=tk.LEFT, padx=(0, 5))
        self.combo_sheets = ttk.Combobox(ctrl_frame, state="readonly", width=18)
        self.combo_sheets.pack(side=tk.LEFT, padx=(0, 5))

        btn_refresh_sheets = ttk.Button(ctrl_frame, text="↻ Refresh Sheets", command=self.on_refresh_sheets)
        btn_refresh_sheets.pack(side=tk.LEFT, padx=(0, 5))

        btn_add_sheet = ttk.Button(ctrl_frame, text="+ New Sheet", command=self.on_add_sheet)
        btn_add_sheet.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(ctrl_frame, text="Range:").pack(side=tk.LEFT, padx=(0, 5))
        self.ent_range = ttk.Entry(ctrl_frame, width=12)
        self.ent_range.insert(0, "")
        self.ent_range.pack(side=tk.LEFT, padx=(0, 5))

        self.var_used_only = tk.BooleanVar(value=True)
        chk_used = ttk.Checkbutton(ctrl_frame, text="Used Range (Auto)", variable=self.var_used_only)
        chk_used.pack(side=tk.LEFT, padx=(0, 10))

        btn_read = ttk.Button(ctrl_frame, text="📥 Load Data", command=self.on_read_data)
        btn_read.pack(side=tk.LEFT)

        # Treeview grid for Excel cells
        grid_frame = ttk.Frame(self.tab_view)
        grid_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(grid_frame, show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(grid_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(grid_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        grid_frame.grid_rowconfigure(0, weight=1)
        grid_frame.grid_columnconfigure(0, weight=1)

    # -------------------------------------------------------------------------
    # Tab 2: Update & Append UI
    # -------------------------------------------------------------------------
    def _build_tab_write(self):
        # Card 1: Append Row
        grp_append = ttk.LabelFrame(self.tab_write, text=" ➕ Append Row ", padding=10)
        grp_append.pack(fill=tk.X, pady=(0, 10))

        f_app1 = ttk.Frame(grp_append)
        f_app1.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(f_app1, text="Row Values (comma-separated or JSON array):").pack(side=tk.LEFT)
        self.ent_append_values = ttk.Entry(f_app1)
        self.ent_append_values.insert(0, "Timestamp, Active, 100")
        self.ent_append_values.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(10, 0))

        f_app2 = ttk.Frame(grp_append)
        f_app2.pack(fill=tk.X, pady=(5, 0))

        ttk.Label(f_app2, text="Destination:").pack(side=tk.LEFT, padx=(0, 5))
        self.var_append_dest = tk.StringVar(value="sheet")
        r1 = ttk.Radiobutton(f_app2, text="Sheet End", variable=self.var_append_dest, value="sheet")
        r1.pack(side=tk.LEFT, padx=(0, 10))
        r2 = ttk.Radiobutton(f_app2, text="Excel Table", variable=self.var_append_dest, value="table")
        r2.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Label(f_app2, text="Table Name:").pack(side=tk.LEFT, padx=(10, 5))
        self.ent_append_table = ttk.Entry(f_app2, width=15)
        self.ent_append_table.pack(side=tk.LEFT, padx=(0, 15))

        btn_do_append = ttk.Button(f_app2, text="Append Row Now", command=self.on_do_append)
        btn_do_append.pack(side=tk.RIGHT)

        # Card 2: Update Specific Range
        grp_update = ttk.LabelFrame(self.tab_write, text=" 💾 Write to Cell Range ", padding=10)
        grp_update.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        f_up1 = ttk.Frame(grp_update)
        f_up1.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(f_up1, text="Target Range Address (e.g. A1:B2):").pack(side=tk.LEFT, padx=(0, 5))
        self.ent_update_addr = ttk.Entry(f_up1, width=15)
        self.ent_update_addr.insert(0, "A1:B2")
        self.ent_update_addr.pack(side=tk.LEFT, padx=(0, 15))

        btn_do_update = ttk.Button(f_up1, text="Write Range Now", command=self.on_do_update_range)
        btn_do_update.pack(side=tk.RIGHT)

        ttk.Label(
            grp_update,
            text="Values (JSON 2D array, e.g. [[\"Product\", \"Price\"], [\"Laptop\", 1200]]):",
        ).pack(anchor="w", pady=(5, 2))

        self.txt_update_values = scrolledtext.ScrolledText(grp_update, height=6, font=("Consolas", 9))
        self.txt_update_values.insert(tk.END, '[[\"Product\", \"Price\"],\n [\"Laptop\", 1200]]')
        self.txt_update_values.pack(fill=tk.BOTH, expand=True)

        # Card 3: Quick Utilities
        grp_tools = ttk.LabelFrame(self.tab_write, text=" 🛠️ Table & Range Utilities ", padding=10)
        grp_tools.pack(fill=tk.X)

        f_t1 = ttk.Frame(grp_tools)
        f_t1.pack(fill=tk.X)

        ttk.Label(f_t1, text="Convert Range to Table:").pack(side=tk.LEFT, padx=(0, 5))
        self.ent_tbl_range = ttk.Entry(f_t1, width=10)
        self.ent_tbl_range.insert(0, "A1:C10")
        self.ent_tbl_range.pack(side=tk.LEFT, padx=(0, 5))

        ttk.Label(f_t1, text="Table Name:").pack(side=tk.LEFT, padx=(5, 5))
        self.ent_tbl_name = ttk.Entry(f_t1, width=14)
        self.ent_tbl_name.insert(0, "SalesTable")
        self.ent_tbl_name.pack(side=tk.LEFT, padx=(0, 10))

        btn_create_tbl = ttk.Button(f_t1, text="Create Table", command=self.on_create_table)
        btn_create_tbl.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(f_t1, text="Clear Range:").pack(side=tk.LEFT, padx=(10, 5))
        self.ent_clear_addr = ttk.Entry(f_t1, width=10)
        self.ent_clear_addr.insert(0, "A2:C20")
        self.ent_clear_addr.pack(side=tk.LEFT, padx=(0, 5))

        btn_clear = ttk.Button(f_t1, text="Clear Range", command=self.on_clear_range)
        btn_clear.pack(side=tk.LEFT)

    # -------------------------------------------------------------------------
    # Tab 3: Configuration & Auth UI
    # -------------------------------------------------------------------------
    def _build_tab_config(self):
        container = ttk.Frame(self.tab_config)
        container.pack(fill=tk.BOTH, expand=True)

        row = 0

        # Flow
        ttk.Label(container, text="Auth Flow:").grid(row=row, column=0, sticky="w", pady=5)
        self.combo_flow = ttk.Combobox(
            container,
            values=["device_code", "client_credentials"],
            state="readonly",
            width=28,
        )
        self.combo_flow.grid(row=row, column=1, sticky="w", pady=5)
        ttk.Label(
            container,
            text="(device_code = Interactive/Personal; client_credentials = Headless Daemon)",
            foreground="#666",
        ).grid(row=row, column=2, sticky="w", padx=10)
        row += 1

        # Client ID
        ttk.Label(container, text="Azure Client ID:").grid(row=row, column=0, sticky="w", pady=5)
        self.ent_client_id = ttk.Entry(container, width=32)
        self.ent_client_id.grid(row=row, column=1, sticky="w", pady=5)
        ttk.Label(container, text="Application (client) ID GUID from Azure Portal", foreground="#666").grid(
            row=row, column=2, sticky="w", padx=10
        )
        row += 1

        # Tenant ID
        ttk.Label(container, text="Tenant ID:").grid(row=row, column=0, sticky="w", pady=5)
        self.ent_tenant_id = ttk.Entry(container, width=32)
        self.ent_tenant_id.grid(row=row, column=1, sticky="w", pady=5)
        ttk.Label(container, text="'common' for personal accounts or Tenant GUID for corporate", foreground="#666").grid(
            row=row, column=2, sticky="w", padx=10
        )
        row += 1

        # Client Secret
        ttk.Label(container, text="Client Secret:").grid(row=row, column=0, sticky="w", pady=5)
        self.ent_secret = ttk.Entry(container, width=32, show="•")
        self.ent_secret.grid(row=row, column=1, sticky="w", pady=5)
        ttk.Label(container, text="Required ONLY for client_credentials flow", foreground="#666").grid(
            row=row, column=2, sticky="w", padx=10
        )
        row += 1

        # Excel File Path
        ttk.Label(container, text="Excel File Path:").grid(row=row, column=0, sticky="w", pady=5)
        self.ent_filepath = ttk.Entry(container, width=32)
        self.ent_filepath.grid(row=row, column=1, sticky="w", pady=5)
        ttk.Label(container, text="OneDrive path (e.g. 'my-app/data.xlsx')", foreground="#666").grid(
            row=row, column=2, sticky="w", padx=10
        )
        row += 1

        # Drive Target
        ttk.Label(container, text="Drive Target:").grid(row=row, column=0, sticky="w", pady=5)
        self.combo_target = ttk.Combobox(
            container,
            values=["me", "user", "site", "drive"],
            state="readonly",
            width=28,
        )
        self.combo_target.grid(row=row, column=1, sticky="w", pady=5)
        ttk.Label(container, text="'me' for signed-in user OneDrive; 'user' for specific user", foreground="#666").grid(
            row=row, column=2, sticky="w", padx=10
        )
        row += 1

        # Target ID
        ttk.Label(container, text="Target User/Site ID:").grid(row=row, column=0, sticky="w", pady=5)
        self.ent_target_id = ttk.Entry(container, width=32)
        self.ent_target_id.grid(row=row, column=1, sticky="w", pady=5)
        ttk.Label(container, text="Used only if Drive Target is 'user', 'site', or 'drive'", foreground="#666").grid(
            row=row, column=2, sticky="w", padx=10
        )
        row += 1

        # Action Buttons
        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=row, column=0, columnspan=3, pady=(20, 0), sticky="w")

        btn_save = ttk.Button(btn_frame, text="💾 Save Settings to .env", command=self.on_save_config)
        btn_save.pack(side=tk.LEFT, padx=(0, 10))

        btn_logout = ttk.Button(btn_frame, text="🚪 Sign Out & Clear Token Cache", command=self.on_logout)
        btn_logout.pack(side=tk.LEFT, padx=(0, 10))

    # -------------------------------------------------------------------------
    # Helper & Logging Methods
    # -------------------------------------------------------------------------
    def log(self, text: str, tag: str = "INFO"):
        """Logs a message to the bottom console in a thread-safe manner."""
        def _write():
            self.txt_log.insert(tk.END, f"[{tag}] {text}\n", tag)
            self.txt_log.see(tk.END)
        self.after(0, _write)

    def set_status(self, text: str, color: str = "#333"):
        def _update():
            self.lbl_auth_status.configure(text=text, foreground=color)
        self.after(0, _update)

    def run_async(self, worker_func, *args, **kwargs):
        """Runs a task in a background daemon thread to keep UI completely responsive."""
        thread = threading.Thread(target=worker_func, args=args, kwargs=kwargs, daemon=True)
        thread.start()

    # -------------------------------------------------------------------------
    # Configuration Loading & Saving
    # -------------------------------------------------------------------------
    def _load_config_to_fields(self):
        self.combo_flow.set(Config.AUTH_FLOW or "device_code")
        self.ent_client_id.delete(0, tk.END)
        self.ent_client_id.insert(0, Config.AZURE_CLIENT_ID)
        self.ent_tenant_id.delete(0, tk.END)
        self.ent_tenant_id.insert(0, Config.AZURE_TENANT_ID)
        self.ent_secret.delete(0, tk.END)
        self.ent_secret.insert(0, Config.AZURE_CLIENT_SECRET)
        self.ent_filepath.delete(0, tk.END)
        self.ent_filepath.insert(0, Config.EXCEL_FILE_PATH)
        self.combo_target.set(Config.DRIVE_TARGET or "me")
        self.ent_target_id.delete(0, tk.END)
        self.ent_target_id.insert(0, Config.TARGET_USER_ID or Config.TARGET_SITE_ID or "")

    def on_save_config(self):
        flow = self.combo_flow.get().strip()
        client_id = self.ent_client_id.get().strip()
        tenant_id = self.ent_tenant_id.get().strip()
        secret = self.ent_secret.get().strip()
        filepath = self.ent_filepath.get().strip()
        drive_target = self.combo_target.get().strip()
        target_id = self.ent_target_id.get().strip()

        env_path = BASE_DIR / ".env"
        lines = [
            "# Microsoft Graph API Configuration (Saved via GUI)\n",
            f"AUTH_FLOW={flow}\n",
            f"AZURE_CLIENT_ID={client_id}\n",
            f"AZURE_CLIENT_SECRET={secret}\n",
            f"AZURE_TENANT_ID={tenant_id}\n",
            f"EXCEL_FILE_PATH={filepath}\n",
            f"DRIVE_TARGET={drive_target}\n",
            f"TARGET_USER_ID={target_id if drive_target == 'user' else ''}\n",
            f"TARGET_SITE_ID={target_id if drive_target == 'site' else ''}\n",
            f"TARGET_DRIVE_ID={target_id if drive_target == 'drive' else ''}\n",
        ]

        try:
            with open(env_path, "w", encoding="utf-8") as f:
                f.writelines(lines)

            # Update running Config
            Config.AUTH_FLOW = flow
            Config.AZURE_CLIENT_ID = client_id
            Config.AZURE_CLIENT_SECRET = secret
            Config.AZURE_TENANT_ID = tenant_id
            Config.EXCEL_FILE_PATH = filepath
            Config.DRIVE_TARGET = drive_target
            if drive_target == "user":
                Config.TARGET_USER_ID = target_id
            elif drive_target == "site":
                Config.TARGET_SITE_ID = target_id

            self.client = None
            self.log(f"Settings saved to '{env_path.name}'.", "SUCCESS")
            messagebox.showinfo("Saved", "Settings successfully saved to .env file!")
        except Exception as e:
            self.log(f"Failed to save settings: {e}", "ERROR")
            messagebox.showerror("Error", f"Could not save .env: {e}")

    # -------------------------------------------------------------------------
    # Authentication & Connection
    # -------------------------------------------------------------------------
    def on_device_code_callback(self, flow):
        """Called when MSAL generates a device code."""
        def _open_dialog():
            self.active_device_dialog = DeviceCodeDialog(self, flow)
        self.after(0, _open_dialog)

    def _worker_connect(self):
        self.set_status("🟡 Status: Connecting...", "#e65100")
        self.log("Initializing Microsoft Graph client...")

        try:
            auth = GraphAuthenticator()
            self.client = GraphExcelClient(
                authenticator=auth,
                device_code_callback=self.on_device_code_callback,
            )

            # Close popup if active
            if self.active_device_dialog:
                self.after(0, self.active_device_dialog.destroy)
                self.active_device_dialog = None

            self.set_status("🟢 Status: Connected", "#2e7d32")
            self.log("Authenticated successfully with Microsoft Graph.", "SUCCESS")

            # Check / Ensure file exists
            self.log(f"Checking remote file '{Config.EXCEL_FILE_PATH}'...")
            meta = self.client.ensure_file_exists()
            self.current_web_url = meta.get("webUrl")
            self.log(f"Remote file ready: {meta.get('name')} (ID: {meta.get('id')})", "SUCCESS")

            if self.current_web_url:
                self.after(0, lambda: self.btn_open_web.configure(state="normal"))

            # Refresh sheets automatically
            self._worker_refresh_sheets()

        except Exception as e:
            if self.active_device_dialog:
                self.after(0, self.active_device_dialog.destroy)
                self.active_device_dialog = None

            self.set_status("🔴 Status: Failed", "#c62828")
            self.log(f"Connection failed: {e}", "ERROR")
            self.after(0, lambda: messagebox.showerror("Connection Error", str(e)))

    def on_connect_clicked(self):
        errors = Config.validate()
        if errors:
            self.log("Configuration validation failed. Check settings tab.", "ERROR")
            messagebox.showwarning("Config Required", "\n".join(errors))
            self.notebook.select(self.tab_config)
            return

        self.run_async(self._worker_connect)

    def on_logout(self):
        try:
            auth = GraphAuthenticator()
            auth.clear_cache()
            self.client = None
            self.set_status("⚪ Status: Disconnected", "#666")
            self.btn_open_web.configure(state="disabled")
            self.combo_sheets["values"] = []
            self.combo_sheets.set("")
            self.log("Token cache cleared. Signed out.", "INFO")
            messagebox.showinfo("Signed Out", "Local credentials cache deleted.")
        except Exception as e:
            self.log(f"Error signing out: {e}", "ERROR")

    def open_excel_online(self):
        if self.current_web_url:
            webbrowser.open(self.current_web_url)
        else:
            messagebox.showinfo("Notice", "No web link available. Connect to file first.")

    # -------------------------------------------------------------------------
    # Sheets Management
    # -------------------------------------------------------------------------
    def _worker_refresh_sheets(self):
        if not self.client:
            return
        try:
            self.log("Fetching worksheets...")
            sheets = self.client.list_worksheets()
            sheet_names = [s.get("name") for s in sheets if s.get("name")]

            def _update():
                self.combo_sheets["values"] = sheet_names
                if sheet_names:
                    if not self.combo_sheets.get() or self.combo_sheets.get() not in sheet_names:
                        self.combo_sheets.set(sheet_names[0])
            self.after(0, _update)
            self.log(f"Found {len(sheet_names)} worksheet(s): {', '.join(sheet_names)}", "INFO")
        except Exception as e:
            self.log(f"Failed to fetch worksheets: {e}", "ERROR")

    def on_refresh_sheets(self):
        if not self.client:
            messagebox.showwarning("Not Connected", "Please connect first.")
            return
        self.run_async(self._worker_refresh_sheets)

    def on_add_sheet(self):
        if not self.client:
            messagebox.showwarning("Not Connected", "Please connect first.")
            return
        name = simpledialog.askstring("New Worksheet", "Enter name for new worksheet:", parent=self)
        if not name or not name.strip():
            return

        def _worker():
            try:
                self.client.add_worksheet(name.strip())
                self.log(f"Created worksheet '{name.strip()}'.", "SUCCESS")
                self._worker_refresh_sheets()
            except Exception as e:
                self.log(f"Failed to add worksheet: {e}", "ERROR")
                self.after(0, lambda: messagebox.showerror("Error", str(e)))

        self.run_async(_worker)

    # -------------------------------------------------------------------------
    # Spreadsheet Grid Display
    # -------------------------------------------------------------------------
    def _worker_read_data(self, sheet: str, address: Optional[str], used_only: bool):
        try:
            self.log(f"Reading data from '{sheet}'...")
            if address:
                res = self.client.read_range(sheet_name=sheet, address=address)
            else:
                res = self.client.read_used_range(sheet_name=sheet, values_only=True)
                if not res:
                    self.log(f"Worksheet '{sheet}' is completely empty.", "INFO")
                    self._render_grid([])
                    return

            values = res.get("values", [])
            self.log(f"Loaded {len(values)} row(s) from '{res.get('address', '')}'.", "SUCCESS")
            self._render_grid(values)
        except Exception as e:
            self.log(f"Failed to read data: {e}", "ERROR")
            self.after(0, lambda: messagebox.showerror("Read Error", str(e)))

    def _render_grid(self, values: list):
        def _update():
            # Clear existing items
            self.tree.delete(*self.tree.get_children())

            if not values:
                self.tree["columns"] = ("Empty",)
                self.tree.heading("Empty", text="No data in worksheet")
                self.tree.column("Empty", width=300, anchor="center")
                return

            max_cols = max(len(row) for row in values)
            def col_letter(i):
                res = ""
                while i >= 0:
                    res = chr(i % 26 + ord("A")) + res
                    i = i // 26 - 1
                return res

            col_ids = ["#Row"] + [col_letter(i) for i in range(max_cols)]
            self.tree["columns"] = col_ids

            self.tree.heading("#Row", text="Row")
            self.tree.column("#Row", width=55, anchor="center")

            for col in col_ids[1:]:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=120, anchor="w")

            for r_idx, row in enumerate(values):
                row_vals = [str(r_idx + 1)] + [str(v) if v is not None else "" for v in row]
                # Pad to max_cols
                while len(row_vals) < len(col_ids):
                    row_vals.append("")
                self.tree.insert("", tk.END, values=row_vals)

        self.after(0, _update)

    def on_read_data(self):
        if not self.client:
            messagebox.showwarning("Not Connected", "Please click 'Connect / Check Status' first.")
            return

        sheet = self.combo_sheets.get().strip() or "Sheet1"
        addr = self.ent_range.get().strip()
        used_only = self.var_used_only.get()

        target_addr = None if (used_only or not addr) else addr
        self.run_async(self._worker_read_data, sheet, target_addr, used_only)

    # -------------------------------------------------------------------------
    # Data Write / Append Handlers
    # -------------------------------------------------------------------------
    def on_do_append(self):
        if not self.client:
            messagebox.showwarning("Not Connected", "Please connect first.")
            return

        val_str = self.ent_append_values.get().strip()
        if not val_str:
            messagebox.showwarning("Input Required", "Enter values to append.")
            return

        # Parse values
        try:
            parsed = json.loads(val_str)
            if not isinstance(parsed, list):
                parsed = [parsed]
        except Exception:
            parts = [p.strip().strip("'\"") for p in val_str.split(",")]
            parsed = []
            for p in parts:
                try:
                    parsed.append(int(p))
                except ValueError:
                    try:
                        parsed.append(float(p))
                    except ValueError:
                        parsed.append(p)

        sheet = self.combo_sheets.get().strip() or "Sheet1"
        dest = self.var_append_dest.get()
        table_name = self.ent_append_table.get().strip()

        def _worker():
            try:
                if dest == "table":
                    if not table_name:
                        self.after(0, lambda: messagebox.showwarning("Input Required", "Please specify Table Name."))
                        return
                    res = self.client.append_table_row(table_name_or_id=table_name, row_values=parsed)
                    self.log(f"Row appended to table '{table_name}': {parsed}", "SUCCESS")
                else:
                    res = self.client.append_row_to_range(row_values=parsed, sheet_name=sheet)
                    self.log(f"Row appended to '{sheet}' at {res.get('address')}: {parsed}", "SUCCESS")

                # Refresh grid view automatically
                self._worker_read_data(sheet, None, True)
            except Exception as e:
                self.log(f"Append failed: {e}", "ERROR")
                self.after(0, lambda: messagebox.showerror("Append Error", str(e)))

        self.run_async(_worker)

    def on_do_update_range(self):
        if not self.client:
            messagebox.showwarning("Not Connected", "Please connect first.")
            return

        sheet = self.combo_sheets.get().strip() or "Sheet1"
        addr = self.ent_update_addr.get().strip()
        raw_vals = self.txt_update_values.get("1.0", tk.END).strip()

        if not addr or not raw_vals:
            messagebox.showwarning("Input Required", "Specify range address and values.")
            return

        try:
            values = json.loads(raw_vals)
            if not isinstance(values, list) or (values and not isinstance(values[0], list)):
                raise ValueError("Values must be a 2D JSON array, e.g. [['A', 'B'], ['C', 'D']]")
        except Exception as pe:
            messagebox.showerror("JSON Error", f"Invalid 2D JSON array format: {pe}")
            return

        def _worker():
            try:
                res = self.client.update_range(sheet_name=sheet, address=addr, values=values)
                self.log(f"Updated range '{res.get('address')}' in sheet '{sheet}'.", "SUCCESS")
                self._worker_read_data(sheet, None, True)
            except Exception as e:
                self.log(f"Update failed: {e}", "ERROR")
                self.after(0, lambda: messagebox.showerror("Write Error", str(e)))

        self.run_async(_worker)

    def on_create_table(self):
        if not self.client:
            messagebox.showwarning("Not Connected", "Please connect first.")
            return
        sheet = self.combo_sheets.get().strip() or "Sheet1"
        addr = self.ent_tbl_range.get().strip()
        tname = self.ent_tbl_name.get().strip() or None
        if not addr:
            messagebox.showwarning("Input Required", "Specify table range address (e.g. A1:C10).")
            return

        def _worker():
            try:
                res = self.client.create_table(address=addr, table_name=tname, sheet_name=sheet)
                self.log(f"Created Excel Table '{res.get('name')}' (ID: {res.get('id')}).", "SUCCESS")
                messagebox.showinfo("Success", f"Table '{res.get('name')}' created successfully!")
            except Exception as e:
                self.log(f"Failed to create table: {e}", "ERROR")
                self.after(0, lambda: messagebox.showerror("Error", str(e)))

        self.run_async(_worker)

    def on_clear_range(self):
        if not self.client:
            messagebox.showwarning("Not Connected", "Please connect first.")
            return
        sheet = self.combo_sheets.get().strip() or "Sheet1"
        addr = self.ent_clear_addr.get().strip()
        if not addr:
            messagebox.showwarning("Input Required", "Specify range to clear.")
            return

        if not messagebox.askyesno("Confirm Clear", f"Are you sure you want to clear '{sheet}!{addr}'?"):
            return

        def _worker():
            try:
                self.client.clear_range(sheet_name=sheet, address=addr)
                self.log(f"Cleared range '{sheet}!{addr}'.", "SUCCESS")
                self._worker_read_data(sheet, None, True)
            except Exception as e:
                self.log(f"Clear failed: {e}", "ERROR")
                self.after(0, lambda: messagebox.showerror("Error", str(e)))

        self.run_async(_worker)


def main():
    app = ExcelGraphGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
