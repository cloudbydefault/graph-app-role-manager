#!/usr/bin/env python3
"""Graph App Role Manager v1.1.

Cross-platform Tkinter GUI using a persistent PowerShell 7 backend.
Authentication is handled by Connect-MgGraph, so no custom App Registration,
client ID, secret, certificate, MSAL package, or requests package is required.
"""
from __future__ import annotations

import json
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import uuid
import webbrowser
from dataclasses import dataclass
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Any, Callable

PROTOCOL_PREFIX = "__GARM__"
APP_VERSION = "1.1.0"


@dataclass(frozen=True)
class AppRole:
    id: str
    value: str
    display_name: str
    description: str


class PowerShellBridge:
    def __init__(
        self,
        backend_path: Path,
        event_callback: Callable[[str, Any], None],
    ) -> None:
        self.backend_path = backend_path
        self.event_callback = event_callback
        self.process: subprocess.Popen[str] | None = None
        self.pending: dict[str, queue.Queue[dict[str, Any]]] = {}
        self.lock = threading.Lock()
        self.ready = threading.Event()
        self.startup_errors: list[str] = []

    @staticmethod
    def find_pwsh() -> str | None:
        return shutil.which("pwsh")

    def start(self) -> None:
        pwsh = self.find_pwsh()
        if not pwsh:
            raise RuntimeError(
                "PowerShell 7 (pwsh) was not found. Install it first, then restart the tool."
            )
        if not self.backend_path.exists():
            raise RuntimeError(f"PowerShell backend not found: {self.backend_path}")

        self.process = subprocess.Popen(
            [pwsh, "-NoLogo", "-NoProfile", "-File", str(self.backend_path)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )
        threading.Thread(target=self._read_stdout, daemon=True).start()
        threading.Thread(target=self._read_stderr, daemon=True).start()

        if not self.ready.wait(timeout=15):
            details = "\n".join(self.startup_errors[-10:]).strip()
            exit_code = self.process.poll() if self.process else None
            self.stop()
            message = "The PowerShell backend did not start within 15 seconds."
            if exit_code is not None:
                message += f" Process exited with code {exit_code}."
            if details:
                message += f"\n\nPowerShell details:\n{details}"
            raise RuntimeError(message)

        if not self.process or self.process.poll() is not None:
            details = "\n".join(self.startup_errors[-10:]).strip()
            exit_code = self.process.poll() if self.process else None
            self.stop()
            message = f"The PowerShell backend exited during startup (code {exit_code})."
            if details:
                message += f"\n\nPowerShell details:\n{details}"
            raise RuntimeError(message)

    def _read_stdout(self) -> None:
        assert self.process and self.process.stdout
        for raw_line in self.process.stdout:
            line = raw_line.rstrip("\r\n")
            if not line.startswith(PROTOCOL_PREFIX):
                if line:
                    self.event_callback("log", f"PowerShell: {line}")
                continue
            try:
                payload = json.loads(line[len(PROTOCOL_PREFIX) :])
            except json.JSONDecodeError:
                self.event_callback("log", f"Invalid backend response: {line}")
                continue

            event = payload.get("event")
            if event:
                if event == "ready":
                    self.ready.set()
                else:
                    self.event_callback(event, payload.get("data"))
                continue

            request_id = payload.get("requestId")
            with self.lock:
                response_queue = self.pending.get(request_id)
            if response_queue:
                response_queue.put(payload)

        # Reaching EOF means the backend exited. Do not mark it as ready.

    def _read_stderr(self) -> None:
        assert self.process and self.process.stderr
        for raw_line in self.process.stderr:
            line = raw_line.rstrip("\r\n")
            if line:
                self.startup_errors.append(line)
                self.event_callback("log", f"PowerShell error stream: {line}")

    def call(self, command: str, timeout: int = 180, **kwargs: Any) -> Any:
        if not self.process or self.process.poll() is not None or not self.process.stdin:
            raise RuntimeError("The PowerShell backend is not running.")

        request_id = str(uuid.uuid4())
        response_queue: queue.Queue[dict[str, Any]] = queue.Queue(maxsize=1)
        with self.lock:
            self.pending[request_id] = response_queue

        request = {"requestId": request_id, "command": command, **kwargs}
        try:
            self.process.stdin.write(json.dumps(request, separators=(",", ":")) + "\n")
            self.process.stdin.flush()
            response = response_queue.get(timeout=timeout)
        except queue.Empty as exc:
            raise RuntimeError(f"PowerShell command '{command}' timed out.") from exc
        finally:
            with self.lock:
                self.pending.pop(request_id, None)

        if not response.get("ok"):
            raise RuntimeError(response.get("error") or "Unknown PowerShell backend error.")
        return response.get("data")

    def stop(self) -> None:
        process = self.process
        self.process = None
        if not process:
            return
        try:
            if process.stdin and process.poll() is None:
                process.stdin.close()
            process.terminate()
            process.wait(timeout=3)
        except Exception:
            try:
                process.kill()
            except Exception:
                pass


class GraphAppRoleManager(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"Graph App Role Manager v{APP_VERSION}")
        self.geometry("1180x800")

        # Use the bundled PNG as the window/application icon when supported.
        self._app_icon = None
        icon_path = Path(__file__).resolve().with_name("graph-app-role-manager-icon.png")
        if icon_path.exists():
            try:
                self._app_icon = tk.PhotoImage(file=str(icon_path))
                self.iconphoto(True, self._app_icon)
            except tk.TclError:
                pass
        self.minsize(1000, 700)
        self.configure(background="#f4f6f8")

        backend = Path(__file__).with_name("graph_backend.ps1")
        self.ui_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.bridge = PowerShellBridge(backend, self._queue_backend_event)
        self.roles: list[AppRole] = []
        self.filtered_roles: list[AppRole] = []
        self.selected_role_ids: set[str] = set()
        self.rendering_permissions = False
        self.identity_results: list[dict[str, Any]] = []
        self.selected_sp: dict[str, Any] | None = None
        self.current_assignments: dict[str, dict[str, Any]] = {}
        self.connected = False
        self.device_code_dialog: tk.Toplevel | None = None

        self._configure_style()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self._process_ui_queue)
        self.after(200, self.start_backend)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Header.TFrame", background="#1f2937")
        style.configure(
            "HeaderTitle.TLabel",
            background="#1f2937",
            foreground="white",
            font=("Segoe UI", 20, "bold"),
        )
        style.configure(
            "HeaderSubtitle.TLabel",
            background="#1f2937",
            foreground="#d1d5db",
            font=("Segoe UI", 10),
        )
        base_font = (
            "SF Pro Text"
            if self.tk.call("tk", "windowingsystem") == "aqua"
            else "Segoe UI",
            10,
        )
        style.configure("TFrame", background="#f4f6f8")
        style.configure(
            "TLabel",
            background="#f4f6f8",
            foreground="#172033",
            font=base_font,
        )
        style.configure(
            "TLabelframe",
            background="#f4f6f8",
            bordercolor="#cbd5e1",
            relief="solid",
        )
        style.configure(
            "TLabelframe.Label",
            background="#f4f6f8",
            foreground="#172033",
            font=(base_font[0], 10, "bold"),
        )
        style.configure("TNotebook", background="#f4f6f8", borderwidth=0)
        style.configure("TNotebook.Tab", padding=(14, 7), font=base_font)
        style.configure("TEntry", padding=6)
        style.configure("TCombobox", padding=5)
        style.configure("TButton", padding=(12, 7), font=base_font)
        style.configure(
            "Primary.TButton",
            font=(base_font[0], 10, "bold"),
            padding=(14, 8),
        )
        style.configure(
            "Success.TButton",
            font=(base_font[0], 10, "bold"),
            padding=(14, 8),
        )
        style.configure(
            "Danger.TButton",
            font=(base_font[0], 10, "bold"),
            padding=(12, 7),
        )
        style.configure(
            "Treeview",
            background="#ffffff",
            fieldbackground="#ffffff",
            foreground="#172033",
            rowheight=28,
            borderwidth=0,
        )
        style.configure(
            "Treeview.Heading",
            background="#e9eef5",
            foreground="#172033",
            font=(base_font[0], 10, "bold"),
            relief="flat",
        )
        style.map(
            "Treeview",
            background=[("selected", "#2563eb")],
            foreground=[("selected", "#ffffff")],
        )

    def _build_ui(self) -> None:
        header = ttk.Frame(self, style="Header.TFrame", padding=(20, 14))
        header.pack(fill="x")
        left = ttk.Frame(header, style="Header.TFrame")
        left.pack(side="left", fill="x", expand=True)
        ttk.Label(left, text="Graph App Role Manager", style="HeaderTitle.TLabel").pack(anchor="w")
        ttk.Label(
            left,
            text="Manage Microsoft Graph application permissions for Managed Identities and Service Principals",
            style="HeaderSubtitle.TLabel",
        ).pack(anchor="w")
        self.connect_button = ttk.Button(
            header,
            text="Connect to Microsoft Graph",
            style="Primary.TButton",
            command=self.connect,
        )
        self.connect_button.pack(side="right", padx=(15, 0))

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=14, pady=14)
        main_tab = ttk.Frame(notebook, padding=12)
        log_tab = ttk.Frame(notebook, padding=12)
        notebook.add(main_tab, text="Assign permissions")
        notebook.add(log_tab, text="Activity log")

        prereq = ttk.LabelFrame(main_tab, text="Connection", padding=12)
        prereq.pack(fill="x", pady=(0, 10))
        self.prereq_var = tk.StringVar(value="Checking PowerShell prerequisites...")
        ttk.Label(prereq, textvariable=self.prereq_var).pack(side="left")
        ttk.Label(prereq, text="Tenant ID/domain (optional):").pack(side="left", padx=(30, 8))
        self.tenant_var = tk.StringVar()
        ttk.Entry(prereq, textvariable=self.tenant_var, width=34).pack(side="left")

        identity = ttk.LabelFrame(main_tab, text="1. Select the target Managed Identity or Service Principal", padding=12)
        identity.pack(fill="x", pady=(0, 10))
        row = ttk.Frame(identity)
        row.pack(fill="x")
        ttk.Label(row, text="Display name").pack(side="left")
        self.identity_name_var = tk.StringVar()
        identity_entry = ttk.Entry(row, textvariable=self.identity_name_var)
        identity_entry.pack(side="left", fill="x", expand=True, padx=10)
        ttk.Button(row, text="Search", command=self.search_identity).pack(side="left")
        self.identity_combo = ttk.Combobox(identity, state="readonly")
        self.identity_combo.pack(fill="x", pady=(10, 8))
        self.identity_combo.bind("<<ComboboxSelected>>", self.identity_selected)
        self.identity_summary_var = tk.StringVar(value="No service principal selected.")
        ttk.Label(identity, textvariable=self.identity_summary_var).pack(anchor="w")

        body = ttk.Panedwindow(main_tab, orient=tk.HORIZONTAL)
        body.pack(fill="both", expand=True)
        permissions = ttk.LabelFrame(
            body, text="2. Select Microsoft Graph application permissions", padding=12
        )
        current = ttk.LabelFrame(body, text="3. Current Microsoft Graph assignments", padding=12)
        body.add(permissions, weight=3)
        body.add(current, weight=2)

        filter_row = ttk.Frame(permissions)
        filter_row.pack(fill="x")
        self.permission_filter_var = tk.StringVar()
        permission_entry = ttk.Entry(filter_row, textvariable=self.permission_filter_var)
        permission_entry.pack(side="left", fill="x", expand=True)
        permission_entry.bind("<KeyRelease>", lambda _event: self.apply_permission_filter())
        ttk.Button(filter_row, text="Load permissions", command=self.load_permissions).pack(
            side="left", padx=(8, 0)
        )

        list_frame = ttk.Frame(permissions)
        list_frame.pack(fill="both", expand=True, pady=10)
        self.permission_list = tk.Listbox(
            list_frame,
            selectmode=tk.MULTIPLE,
            activestyle="none",
            exportselection=False,
            background="#ffffff",
            foreground="#172033",
            selectbackground="#2563eb",
            selectforeground="#ffffff",
            highlightbackground="#cbd5e1",
            highlightcolor="#2563eb",
            highlightthickness=1,
            borderwidth=0,
            relief="flat",
            font=(
                "SF Pro Text"
                if self.tk.call("tk", "windowingsystem") == "aqua"
                else "Segoe UI",
                10,
            ),
        )
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.permission_list.yview)
        self.permission_list.configure(yscrollcommand=scrollbar.set)
        self.permission_list.bind("<<ListboxSelect>>", self.permission_selection_changed)
        self.permission_list.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        bottom = ttk.Frame(permissions)
        bottom.pack(fill="x")
        self.permission_count_var = tk.StringVar(value="0 permission(s) displayed")
        ttk.Label(bottom, textvariable=self.permission_count_var).pack(side="left")
        self.selected_permission_count_var = tk.StringVar(value="0 permission(s) selected")
        ttk.Label(bottom, textvariable=self.selected_permission_count_var).pack(side="left", padx=(14, 0))
        ttk.Button(bottom, text="Clear selection", command=self.clear_permission_selection).pack(side="right", padx=(8, 0))
        ttk.Button(
            bottom,
            text="Assign selected permissions",
            style="Success.TButton",
            command=self.assign_permissions,
        ).pack(side="right")

        refresh_row = ttk.Frame(current)
        refresh_row.pack(fill="x")
        ttk.Button(refresh_row, text="Refresh", command=self.refresh_assignments).pack(side="left")
        self.assigned_count_var = tk.StringVar(value="0 permission(s) assigned")
        ttk.Label(refresh_row, textvariable=self.assigned_count_var).pack(side="left", padx=10)
        ttk.Button(
            refresh_row,
            text="Remove selected",
            style="Danger.TButton",
            command=self.remove_permissions,
        ).pack(side="right")
        self.assignment_tree = ttk.Treeview(
            current, columns=("permission", "display"), show="headings"
        )
        self.assignment_tree.heading("permission", text="Permission")
        self.assignment_tree.heading("display", text="Display name")
        self.assignment_tree.column("permission", width=260)
        self.assignment_tree.column("display", width=220)
        self.assignment_tree.pack(fill="both", expand=True, pady=(10, 0))

        self.log_text = tk.Text(
            log_tab,
            wrap="none",
            background="#111827",
            foreground="#e5e7eb",
            insertbackground="white",
            font=("Menlo" if self.tk.call("tk", "windowingsystem") == "aqua" else "Consolas", 10),
        )
        self.log_text.pack(fill="both", expand=True)
        self.log(f"Graph App Role Manager v{APP_VERSION} started.")

    def _queue_backend_event(self, event: str, payload: Any) -> None:
        self.ui_queue.put((event, payload))

    def start_backend(self) -> None:
        self.set_busy(True)

        def worker() -> None:
            try:
                self.bridge.start()
                prereq = self.bridge.call("prerequisites", timeout=30)
                self.ui_queue.put(("backend_ready", prereq))
            except Exception as exc:  # noqa: BLE001
                self.ui_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def run_async(self, event: str, operation: Callable[[], Any]) -> None:
        self.set_busy(True)

        def worker() -> None:
            try:
                self.ui_queue.put((event, operation()))
            except Exception as exc:  # noqa: BLE001
                self.ui_queue.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _process_ui_queue(self) -> None:
        try:
            while True:
                event, payload = self.ui_queue.get_nowait()
                if event == "log":
                    self.log(payload)
                elif event == "deviceCode":
                    self.show_device_code(payload)
                elif event == "authMessage":
                    self.log(payload.get("message", "Microsoft Graph authentication update."))
                elif event == "backend_ready":
                    self.on_backend_ready(payload)
                elif event == "connected":
                    self.on_connected(payload)
                elif event == "identities":
                    self.on_identities(payload)
                elif event == "roles":
                    self.on_roles(payload)
                elif event == "assignments":
                    self.on_assignments(payload)
                elif event == "assignment_complete":
                    self.on_assignment_complete(payload)
                elif event == "removal_complete":
                    self.on_removal_complete(payload)
                elif event == "error":
                    self.set_busy(False)
                    self.dismiss_device_code_dialog()
                    self.log(payload, "ERROR")
                    messagebox.showerror("Graph App Role Manager", payload)
        except queue.Empty:
            pass
        self.after(100, self._process_ui_queue)

    def set_busy(self, busy: bool) -> None:
        self.configure(cursor="watch" if busy else "")
        self.connect_button.configure(state="disabled" if busy else "normal")
        self.update_idletasks()

    def log(self, message: str, level: str = "INFO") -> None:
        from datetime import datetime

        self.log_text.insert("end", f"[{datetime.now():%H:%M:%S}][{level}] {message}\n")
        self.log_text.see("end")

    def on_backend_ready(self, prereq: dict[str, Any]) -> None:
        self.set_busy(False)
        version = prereq.get("powershellVersion", "unknown")
        graph_ok = prereq.get("graphModuleInstalled", False)
        graph_version = prereq.get("graphModuleVersion") or "not installed"
        self.prereq_var.set(f"PowerShell {version} | Microsoft.Graph.Authentication {graph_version}")
        self.log(f"PowerShell backend ready. Version: {version}", "SUCCESS")
        if not graph_ok:
            messagebox.showwarning(
                "Missing Microsoft Graph module",
                "Microsoft.Graph.Authentication is not installed.\n\nRun in PowerShell 7:\n"
                "Install-Module Microsoft.Graph.Authentication -Scope CurrentUser",
            )

    def connect(self) -> None:
        tenant = self.tenant_var.get().strip()
        self.log("Starting Microsoft Graph authentication...")
        self.run_async(
            "connected",
            lambda: self.bridge.call("connect", tenantId=tenant, timeout=300),
        )

    def show_device_code(self, payload: dict[str, Any]) -> None:
        self.dismiss_device_code_dialog()

        verification_uri = (
            payload.get("verificationUri") or "https://microsoft.com/devicelogin"
        )
        user_code = payload.get("userCode") or ""
        message = payload.get("message") or (
            f"Open {verification_uri} and enter code {user_code}."
        )
        self.log(message)

        dialog = tk.Toplevel(self)
        self.device_code_dialog = dialog
        dialog.title("Sign in to Microsoft Graph")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.configure(background="#eef2f7")

        card = tk.Frame(
            dialog,
            background="white",
            highlightthickness=1,
            highlightbackground="#cbd5e1",
        )
        card.pack(fill="both", expand=True, padx=20, pady=20)

        header = tk.Frame(card, background="#1f2937")
        header.pack(fill="x")
        tk.Label(
            header,
            text="Complete Microsoft Graph sign-in",
            background="#1f2937",
            foreground="white",
            font=("Segoe UI", 15, "bold"),
            padx=18,
            pady=14,
        ).pack(anchor="w")

        content = tk.Frame(card, background="white", padx=18, pady=18)
        content.pack(fill="both", expand=True)
        tk.Label(
            content,
            text="A browser has been opened. Enter this one-time code:",
            background="white",
            foreground="#172033",
            font=("Segoe UI", 11),
        ).pack(anchor="w")

        code_entry = ttk.Entry(
            content,
            justify="center",
            font=("Consolas", 20, "bold"),
            width=18,
        )
        code_entry.insert(0, user_code)
        code_entry.configure(state="readonly")
        code_entry.pack(fill="x", pady=(14, 8))

        ttk.Label(
            content,
            text=verification_uri,
            foreground="#2563eb",
        ).pack(anchor="center", pady=(0, 12))

        button_row = ttk.Frame(content)
        button_row.pack(fill="x")

        def copy_code() -> None:
            self.clipboard_clear()
            self.clipboard_append(user_code)
            self.update()
            self.log("Device sign-in code copied to the clipboard.")

        ttk.Button(button_row, text="Copy code", command=copy_code).pack(side="left")
        ttk.Button(
            button_row,
            text="Open browser",
            style="Primary.TButton",
            command=lambda: webbrowser.open(verification_uri),
        ).pack(side="right")

        dialog.protocol("WM_DELETE_WINDOW", self.dismiss_device_code_dialog)
        dialog.update_idletasks()
        width, height = 520, 290
        x = self.winfo_rootx() + max(0, (self.winfo_width() - width) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - height) // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.lift()
        webbrowser.open(verification_uri)

    def dismiss_device_code_dialog(self) -> None:
        dialog = self.device_code_dialog
        self.device_code_dialog = None
        if dialog and dialog.winfo_exists():
            dialog.destroy()

    def on_connected(self, result: dict[str, Any]) -> None:
        self.set_busy(False)
        self.dismiss_device_code_dialog()
        self.connected = True
        account = result.get("account") or "Connected"
        self.connect_button.configure(text=f"Connected: {account}")
        self.log(f"Connected as {account} in tenant {result.get('tenantId')}", "SUCCESS")
        self.load_permissions()

    def populate_identity_combo(self, values: list[dict[str, Any]]) -> None:
        self.identity_results = values
        labels = [
            f"{item.get('displayName')} | {item.get('servicePrincipalType')} | {item.get('id')}"
            for item in values
        ]
        self.identity_combo["values"] = labels
        if values:
            self.identity_combo.current(0)
            self.identity_selected()
        else:
            self.identity_combo.set("")
            self.selected_sp = None
            self.identity_summary_var.set("No matching Managed Identity or Service Principal found.")

    def search_identity(self) -> None:
        if not self.connected:
            messagebox.showwarning("Not connected", "Connect to Microsoft Graph first.")
            return
        name = self.identity_name_var.get().strip()
        if not name:
            messagebox.showwarning("Missing name", "Enter a display name.")
            return
        self.log(f"Searching Managed Identities and Service Principals containing '{name}'...")
        self.run_async(
            "identities", lambda: self.bridge.call("searchServicePrincipals", prefix=name)
        )

    def on_identities(self, values: list[dict[str, Any]]) -> None:
        self.set_busy(False)
        self.populate_identity_combo(values)
        if values:
            self.log(f"Found {len(values)} matching service principal(s).", "SUCCESS")

    def identity_selected(self, _event=None) -> None:
        index = self.identity_combo.current()
        if index < 0:
            return
        self.selected_sp = self.identity_results[index]
        sp_type = self.selected_sp.get("servicePrincipalType") or "ServicePrincipal"
        enabled = self.selected_sp.get("accountEnabled")
        status = "Enabled" if enabled is True else "Disabled" if enabled is False else "Unknown"
        self.identity_summary_var.set(
            f"Name: {self.selected_sp.get('displayName')}   |   "
            f"Type: {sp_type}   |   Status: {status}   |   "
            f"Object ID: {self.selected_sp.get('id')}   |   "
            f"App ID: {self.selected_sp.get('appId')}"
        )
        self.refresh_assignments()

    def load_permissions(self) -> None:
        if not self.connected:
            return
        self.log("Loading Microsoft Graph application permissions...")
        self.run_async("roles", lambda: self.bridge.call("loadGraphRoles"))

    def on_roles(self, values: list[dict[str, Any]]) -> None:
        self.set_busy(False)
        self.roles = [
            AppRole(
                id=item["id"],
                value=item["value"],
                display_name=item.get("displayName", ""),
                description=item.get("description", ""),
            )
            for item in values
        ]
        self.apply_permission_filter()
        self.log(f"Loaded {len(self.roles)} application permissions.", "SUCCESS")

    def apply_permission_filter(self) -> None:
        # Preserve selections made under previous filters while rebuilding the visible list.
        self.permission_selection_changed()
        term = self.permission_filter_var.get().strip().lower()
        self.filtered_roles = [
            role
            for role in self.roles
            if not term
            or term in role.value.lower()
            or term in role.display_name.lower()
            or term in role.description.lower()
        ]
        self.rendering_permissions = True
        try:
            self.permission_list.delete(0, "end")
            for index, role in enumerate(self.filtered_roles):
                self.permission_list.insert("end", f"{role.value} — {role.display_name}")
                if role.id in self.selected_role_ids:
                    self.permission_list.selection_set(index)
        finally:
            self.rendering_permissions = False
        self.permission_count_var.set(f"{len(self.filtered_roles)} permission(s) displayed")
        self.selected_permission_count_var.set(f"{len(self.selected_role_ids)} permission(s) selected")

    def permission_selection_changed(self, _event=None) -> None:
        if self.rendering_permissions:
            return
        visible_ids = {role.id for role in self.filtered_roles}
        self.selected_role_ids.difference_update(visible_ids)
        for index in self.permission_list.curselection():
            self.selected_role_ids.add(self.filtered_roles[index].id)
        self.selected_permission_count_var.set(f"{len(self.selected_role_ids)} permission(s) selected")

    def clear_permission_selection(self) -> None:
        self.selected_role_ids.clear()
        self.permission_filter_var.set("")
        self.apply_permission_filter()
        self.permission_list.selection_clear(0, "end")
        self.selected_permission_count_var.set("0 permission(s) selected")

    def refresh_assignments(self) -> None:
        if not self.selected_sp or not self.connected:
            return
        target_id = self.selected_sp["id"]
        self.run_async(
            "assignments", lambda: self.bridge.call("getAssignments", targetId=target_id)
        )

    def on_assignments(self, assignments: list[dict[str, Any]]) -> None:
        self.set_busy(False)
        self.current_assignments = {
            item["id"]: item for item in assignments if item.get("id")
        }
        for item in self.assignment_tree.get_children():
            self.assignment_tree.delete(item)
        role_by_id = {role.id: role for role in self.roles}
        for assignment in assignments:
            role = role_by_id.get(str(assignment.get("appRoleId")))
            self.assignment_tree.insert(
                "",
                "end",
                iid=assignment.get("id"),
                values=(
                    role.value if role else assignment.get("appRoleId", "Unknown"),
                    role.display_name if role else "",
                ),
            )
        self.assigned_count_var.set(f"{len(assignments)} permission(s) assigned")

    def show_assignment_confirmation(self, selected_roles: list[AppRole]) -> bool:
        dialog = tk.Toplevel(self)
        dialog.title("Confirm permission assignment")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(background="#eef2f7")

        result = {"confirmed": False}
        card = tk.Frame(dialog, background="white", highlightthickness=1, highlightbackground="#cbd5e1")
        card.pack(fill="both", expand=True, padx=20, pady=20)

        header = tk.Frame(card, background="#1f2937")
        header.pack(fill="x")
        tk.Label(header, text="Confirm permission assignment", background="#1f2937", foreground="white",
                 font=("Segoe UI", 16, "bold"), padx=18, pady=14).pack(anchor="w")

        content = tk.Frame(card, background="white", padx=18, pady=16)
        content.pack(fill="both", expand=True)
        target = self.selected_sp or {}
        details = (
            f"Target: {target.get('displayName')}\n"
            f"Type: {target.get('servicePrincipalType') or 'ServicePrincipal'}\n"
            f"Object ID: {target.get('id')}\n"
            f"App ID: {target.get('appId')}"
        )
        tk.Label(content, text=details, justify="left", anchor="w", background="white",
                 font=("Segoe UI", 11)).pack(fill="x", anchor="w")

        tk.Label(content, text=f"Permissions to assign ({len(selected_roles)}):", background="white",
                 font=("Segoe UI", 11, "bold"), pady=10).pack(anchor="w")
        list_frame = tk.Frame(content, background="white")
        list_frame.pack(fill="both", expand=True)
        permission_box = tk.Listbox(list_frame, height=min(10, max(3, len(selected_roles))),
                                    background="#f9fafb", foreground="#111827",
                                    selectbackground="#dbeafe", relief="flat",
                                    font=("Menlo" if self.tk.call("tk", "windowingsystem") == "aqua" else "Consolas", 10))
        permission_box.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=permission_box.yview)
        scroll.pack(side="right", fill="y")
        permission_box.configure(yscrollcommand=scroll.set)
        for role in selected_roles:
            permission_box.insert("end", role.value)

        buttons = tk.Frame(card, background="#f9fafb", padx=18, pady=14)
        buttons.pack(fill="x")

        def cancel() -> None:
            dialog.destroy()

        def confirm() -> None:
            result["confirmed"] = True
            dialog.destroy()

        ttk.Button(buttons, text="Cancel", command=cancel).pack(side="right")
        ttk.Button(
            buttons,
            text="Assign permissions",
            command=confirm,
            style="Success.TButton",
        ).pack(side="right", padx=(0, 10))

        dialog.update_idletasks()
        width = 680
        height = min(650, 360 + len(selected_roles) * 22)
        x = self.winfo_rootx() + max(0, (self.winfo_width() - width) // 2)
        y = self.winfo_rooty() + max(0, (self.winfo_height() - height) // 2)
        dialog.geometry(f"{width}x{height}+{x}+{y}")
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.wait_window()
        return result["confirmed"]

    def assign_permissions(self) -> None:
        if not self.selected_sp:
            messagebox.showwarning("No target", "Select a Managed Identity or Service Principal first.")
            return
        self.permission_selection_changed()
        role_by_id = {role.id: role for role in self.roles}
        selected_roles = [role_by_id[role_id] for role_id in self.selected_role_ids if role_id in role_by_id]
        selected_roles.sort(key=lambda role: role.value.lower())
        if not selected_roles:
            messagebox.showwarning("No permissions selected", "Select at least one permission.")
            return
        if not self.show_assignment_confirmation(selected_roles):
            return
        target_id = self.selected_sp["id"]

        def operation() -> dict[str, list[str]]:
            existing = self.bridge.call("getAssignments", targetId=target_id)
            existing_ids = {str(item.get("appRoleId")) for item in existing}
            summary: dict[str, list[str]] = {"assigned": [], "skipped": [], "failed": []}
            for role in selected_roles:
                if role.id in existing_ids:
                    summary["skipped"].append(role.value)
                    continue
                try:
                    self.bridge.call("assignRole", targetId=target_id, appRoleId=role.id)
                    summary["assigned"].append(role.value)
                except Exception as exc:  # noqa: BLE001
                    summary["failed"].append(f"{role.value}: {exc}")
            return summary

        self.run_async("assignment_complete", operation)

    def on_assignment_complete(self, summary: dict[str, list[str]]) -> None:
        self.set_busy(False)
        for value in summary["assigned"]:
            self.log(f"Assigned successfully: {value}", "SUCCESS")
        for value in summary["skipped"]:
            self.log(f"Already assigned: {value}", "WARNING")
        for value in summary["failed"]:
            self.log(f"Failed: {value}", "ERROR")
        messagebox.showinfo(
            "Graph App Role Manager",
            "Completed.\n\n"
            f"Assigned: {len(summary['assigned'])}\n"
            f"Already present: {len(summary['skipped'])}\n"
            f"Failed: {len(summary['failed'])}",
        )
        if not summary["failed"]:
            self.clear_permission_selection()
        self.refresh_assignments()

    def remove_permissions(self) -> None:
        if not self.selected_sp:
            messagebox.showwarning("No target", "Select a Managed Identity or Service Principal first.")
            return
        selected_ids = list(self.assignment_tree.selection())
        if not selected_ids:
            messagebox.showwarning("No permissions selected", "Select at least one assignment.")
            return
        if not messagebox.askyesno(
            "Confirm removal",
            f"Remove {len(selected_ids)} Microsoft Graph permission assignment(s) from\n"
            f"{self.selected_sp.get('displayName')}?",
        ):
            return
        target_id = self.selected_sp["id"]

        def operation() -> dict[str, list[str]]:
            summary: dict[str, list[str]] = {"removed": [], "failed": []}
            for assignment_id in selected_ids:
                try:
                    self.bridge.call(
                        "removeAssignment", targetId=target_id, assignmentId=assignment_id
                    )
                    summary["removed"].append(assignment_id)
                except Exception as exc:  # noqa: BLE001
                    summary["failed"].append(f"{assignment_id}: {exc}")
            return summary

        self.run_async("removal_complete", operation)

    def on_removal_complete(self, summary: dict[str, list[str]]) -> None:
        self.set_busy(False)
        messagebox.showinfo(
            "Graph App Role Manager",
            f"Removed: {len(summary['removed'])}\nFailed: {len(summary['failed'])}",
        )
        self.refresh_assignments()

    def on_close(self) -> None:
        try:
            if self.connected:
                self.bridge.call("disconnect", timeout=10)
        except Exception:
            pass
        self.bridge.stop()
        self.destroy()


if __name__ == "__main__":
    try:
        app = GraphAppRoleManager()
        app.mainloop()
    except tk.TclError as exc:
        print(f"Tkinter could not start: {exc}", file=sys.stderr)
        sys.exit(1)
