from __future__ import annotations

import json
import os
import site
import shutil
import csv
import subprocess
import sys
import threading
import tempfile
import urllib.error
import urllib.request
import uuid
import base64
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

APP_DIR = Path(__file__).resolve().parent
USER_SITE = site.getusersitepackages()
if USER_SITE and USER_SITE not in sys.path:
    sys.path.append(USER_SITE)
BRIDGE_VENDOR = APP_DIR / "bridge_vendor"
if BRIDGE_VENDOR.exists() and str(BRIDGE_VENDOR) not in sys.path:
    sys.path.insert(0, str(BRIDGE_VENDOR))

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except Exception:
    pystray = None
    Image = None
    ImageDraw = None
    TRAY_AVAILABLE = False

CONFIG_FILE = APP_DIR / "codex_bridge_config.json"
DEFAULT_BRIDGE_ROOT = Path.home() / "Syncthing" / "codex-bridge"
SECRETS_FILE = Path.home() / ".codex_bridge_secrets.json"
DEFAULT_JOBS_CSV = Path.home() / "Mitchell EMS" / "Mitchell Data" / "jobs.csv"
POLL_INTERVAL_MS = 3_000
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)
SSH_NONINTERACTIVE_ARGS = ["-o", "BatchMode=yes", "-o", "StrictHostKeyChecking=accept-new"]


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def make_id(prefix: str) -> str:
    return f"{prefix}-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:8]}"


@dataclass
class BridgeConfig:
    local_node: str = "home"
    display_name: str = "Home Codex"
    bridge_root: str = str(DEFAULT_BRIDGE_ROOT)
    project_roots: list[str] = field(default_factory=list)
    focus_project: str = ""
    focus_note: str = ""
    auto_reply_enabled: bool = False
    auto_reply_model: str = "gpt-4.1-mini"
    home_host: str = "100.102.108.49"
    home_user: str = "ferna"
    home_key: str = str(Path.home() / ".ssh" / "home_pc_id_ed25519")
    office_host: str = "100.69.179.72"
    office_user: str = "ferna"
    office_key: str = str(Path.home() / ".ssh" / "office_pc_access_v2")
    auto_execute_trusted_commands: bool = True
    auto_convert_trusted_chat_commands: bool = True


class CodexBridgeApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Codex Bridge")
        self.root.geometry("1380x860")

        self.config = self.load_config()
        self.active_commands: set[str] = set()
        self.active_auto_replies: set[str] = set()
        self.latest_messages: list[dict] = []
        self.latest_commands: list[dict] = []
        self.latest_projects: dict[str, dict] = {}
        self.latest_nodes: dict[str, dict] = {}
        self.local_entries: list[dict] = []
        self.remote_entries: list[dict] = []
        self.remote_refresh_inflight = False
        self.remote_refresh_generation = 0
        self.tray_icon: pystray.Icon | None = None
        self.tray_thread: threading.Thread | None = None
        self.tray_started = False
        self.is_hidden_to_tray = False

        self.local_node_var = tk.StringVar(value=self.config.local_node)
        self.display_name_var = tk.StringVar(value=self.config.display_name)
        self.bridge_root_var = tk.StringVar(value=self.config.bridge_root)
        self.focus_project_var = tk.StringVar(value=self.config.focus_project)
        self.chat_target_var = tk.StringVar(value="all")
        self.command_target_var = tk.StringVar(value=self.opposite_node(self.config.local_node))
        self.command_project_var = tk.StringVar()
        self.command_workdir_var = tk.StringVar()
        self.command_requires_approval_var = tk.BooleanVar(value=True)
        self.command_summary_var = tk.StringVar()
        self.explorer_target_var = tk.StringVar(value=self.opposite_node(self.config.local_node))
        self.local_path_var = tk.StringVar(value=str(Path.home()))
        self.remote_path_var = tk.StringVar(value=r"C:\Users\ferna")
        self.auto_reply_enabled_var = tk.BooleanVar(value=self.config.auto_reply_enabled)
        self.auto_reply_model_var = tk.StringVar(value=self.config.auto_reply_model)
        self.home_host_var = tk.StringVar(value=self.config.home_host)
        self.home_user_var = tk.StringVar(value=self.config.home_user)
        self.home_key_var = tk.StringVar(value=self.config.home_key)
        self.office_host_var = tk.StringVar(value=self.config.office_host)
        self.office_user_var = tk.StringVar(value=self.config.office_user)
        self.office_key_var = tk.StringVar(value=self.config.office_key)
        self.jobs_csv_var = tk.StringVar(value=str(DEFAULT_JOBS_CSV))
        self.ro_lookup_var = tk.StringVar()
        tray_note = "" if TRAY_AVAILABLE else " Tray support is unavailable on this Python build, so closing the window will exit the bridge."
        self.status_var = tk.StringVar(value="Configure the shared bridge folder, then keep the app open on both PCs." + tray_note)

        self._build_ui()
        self.ensure_bridge_dirs()
        self.setup_tray()
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.refresh_all()
        self.root.after(POLL_INTERVAL_MS, self.poll_bridge)

    def load_config(self) -> BridgeConfig:
        if not CONFIG_FILE.exists():
            return BridgeConfig()
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8-sig"))
            return BridgeConfig(**data)
        except Exception:
            return BridgeConfig()

    def save_config(self) -> None:
        self.config.local_node = self.local_node_var.get().strip() or "home"
        self.config.display_name = self.display_name_var.get().strip() or self.default_display_name(self.config.local_node)
        self.config.bridge_root = self.bridge_root_var.get().strip() or str(DEFAULT_BRIDGE_ROOT)
        self.config.focus_project = self.focus_project_var.get().strip()
        self.config.focus_note = self.focus_note_text.get("1.0", "end").strip()
        self.config.project_roots = list(self.project_roots_list.get(0, "end"))
        self.config.auto_reply_enabled = bool(self.auto_reply_enabled_var.get())
        self.config.auto_reply_model = self.auto_reply_model_var.get().strip() or "gpt-4.1-mini"
        self.config.home_host = self.home_host_var.get().strip()
        self.config.home_user = self.home_user_var.get().strip() or "ferna"
        self.config.home_key = self.home_key_var.get().strip()
        self.config.office_host = self.office_host_var.get().strip()
        self.config.office_user = self.office_user_var.get().strip() or "ferna"
        self.config.office_key = self.office_key_var.get().strip()
        CONFIG_FILE.write_text(json.dumps(asdict(self.config), indent=2), encoding="utf-8")

    def default_display_name(self, node: str) -> str:
        return "Home Codex" if node == "home" else "Office Codex"

    def bridge_root(self) -> Path:
        text = self.bridge_root_var.get().strip()
        return Path(text) if text else DEFAULT_BRIDGE_ROOT

    def messages_dir(self) -> Path:
        return self.bridge_root() / "messages"

    def commands_dir(self) -> Path:
        return self.bridge_root() / "commands"

    def tasks_dir(self) -> Path:
        return self.bridge_root() / "tasks"

    def state_dir(self) -> Path:
        return self.bridge_root() / "state"

    def ensure_bridge_dirs(self) -> None:
        for path in (self.bridge_root(), self.messages_dir(), self.commands_dir(), self.tasks_dir(), self.state_dir()):
            path.mkdir(parents=True, exist_ok=True)

    def _build_ui(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        header = ttk.Frame(self.root, padding=12)
        header.grid(row=0, column=0, sticky="ew")
        for column in range(7):
            header.columnconfigure(column, weight=1 if column in (1, 3, 5) else 0)

        ttk.Label(header, text="Node").grid(row=0, column=0, sticky="w")
        self.node_box = ttk.Combobox(header, textvariable=self.local_node_var, values=("home", "office"), state="readonly", width=12)
        self.node_box.grid(row=0, column=1, sticky="ew", padx=(6, 12))
        self.node_box.bind("<<ComboboxSelected>>", lambda _event: self.on_node_changed())

        ttk.Label(header, text="Display Name").grid(row=0, column=2, sticky="w")
        ttk.Entry(header, textvariable=self.display_name_var).grid(row=0, column=3, sticky="ew", padx=(6, 12))

        ttk.Label(header, text="Bridge Folder").grid(row=0, column=4, sticky="w")
        ttk.Entry(header, textvariable=self.bridge_root_var).grid(row=0, column=5, sticky="ew", padx=(6, 6))
        ttk.Button(header, text="Choose", command=self.choose_bridge_folder).grid(row=0, column=6, sticky="ew")
        ttk.Button(header, text="Hide To Tray", command=self.hide_to_tray).grid(row=1, column=5, sticky="ew", pady=(8, 0), padx=(6, 6))
        ttk.Button(header, text="Exit Bridge", command=self.exit_app).grid(row=1, column=6, sticky="ew", pady=(8, 0))

        controls = ttk.Frame(self.root, padding=(12, 0, 12, 12))
        controls.grid(row=1, column=0, sticky="nsew")
        controls.columnconfigure(0, weight=3)
        controls.columnconfigure(1, weight=2)
        controls.rowconfigure(0, weight=1)

        left = ttk.Notebook(controls)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right = ttk.Notebook(controls)
        right.grid(row=0, column=1, sticky="nsew")

        chat_tab = ttk.Frame(left, padding=10)
        commands_tab = ttk.Frame(left, padding=10)
        approvals_tab = ttk.Frame(left, padding=10)
        explorer_tab = ttk.Frame(left, padding=10)
        left.add(chat_tab, text="Chat")
        left.add(commands_tab, text="Commands")
        left.add(approvals_tab, text="Command Inbox")
        left.add(explorer_tab, text="Explorer")

        projects_tab = ttk.Frame(right, padding=10)
        setup_tab = ttk.Frame(right, padding=10)
        log_tab = ttk.Frame(right, padding=10)
        data_tab = ttk.Frame(right, padding=10)
        right.add(projects_tab, text="Projects")
        right.add(setup_tab, text="Setup")
        right.add(log_tab, text="Log")
        right.add(data_tab, text="Repair Center")

        self.build_chat_tab(chat_tab)
        self.build_commands_tab(commands_tab)
        self.build_approvals_tab(approvals_tab)
        self.build_explorer_tab(explorer_tab)
        self.build_projects_tab(projects_tab)
        self.build_setup_tab(setup_tab)
        self.build_log_tab(log_tab)
        self.build_data_tab(data_tab)

        status_bar = ttk.Label(self.root, textvariable=self.status_var, anchor="w", padding=(12, 6))
        status_bar.grid(row=2, column=0, sticky="ew")

    def build_chat_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        self.chat_text = tk.Text(parent, wrap="word", state="disabled")
        self.chat_text.grid(row=0, column=0, sticky="nsew")

        compose = ttk.Frame(parent)
        compose.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        compose.columnconfigure(1, weight=1)

        ttk.Label(compose, text="Target").grid(row=0, column=0, sticky="w")
        ttk.Combobox(compose, textvariable=self.chat_target_var, values=("all", "home", "office"), state="readonly", width=12).grid(row=0, column=1, sticky="w", padx=(6, 12))

        self.chat_entry = tk.Text(compose, height=5, wrap="word")
        self.chat_entry.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        ttk.Button(compose, text="Send Chat Message", command=self.send_chat_message).grid(row=2, column=2, sticky="e", pady=(8, 0))

    def build_commands_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        parent.rowconfigure(5, weight=1)

        ttk.Label(parent, text="Target Node").grid(row=0, column=0, sticky="w")
        ttk.Combobox(parent, textvariable=self.command_target_var, values=("home", "office"), state="readonly").grid(row=0, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(parent, text="Project").grid(row=1, column=0, sticky="w")
        ttk.Entry(parent, textvariable=self.command_project_var).grid(row=1, column=1, sticky="ew", pady=(0, 8))

        ttk.Label(parent, text="Working Directory").grid(row=2, column=0, sticky="w")
        workdir_row = ttk.Frame(parent)
        workdir_row.grid(row=2, column=1, sticky="ew", pady=(0, 8))
        workdir_row.columnconfigure(0, weight=1)
        ttk.Entry(workdir_row, textvariable=self.command_workdir_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(workdir_row, text="Choose", command=self.choose_command_workdir).grid(row=0, column=1, padx=(6, 0))

        ttk.Label(parent, text="Summary").grid(row=3, column=0, sticky="w")
        ttk.Entry(parent, textvariable=self.command_summary_var).grid(row=3, column=1, sticky="ew", pady=(0, 8))

        ttk.Checkbutton(parent, text="Require Approval Before Running", variable=self.command_requires_approval_var).grid(row=4, column=1, sticky="w", pady=(0, 8))

        ttk.Label(parent, text="PowerShell Command").grid(row=5, column=0, sticky="nw")
        self.command_text = tk.Text(parent, height=14, wrap="word")
        self.command_text.grid(row=5, column=1, sticky="nsew")

        ttk.Button(parent, text="Send Command Request", command=self.send_command_request).grid(row=6, column=1, sticky="e", pady=(10, 0))
        direct_row = ttk.Frame(parent)
        direct_row.grid(row=7, column=1, sticky="e", pady=(8, 0))
        ttk.Button(direct_row, text="Run Direct Via SSH", command=self.run_direct_ssh_command).grid(row=0, column=0, padx=(0, 8))
        ttk.Button(direct_row, text="Open Direct Shell", command=self.open_direct_ssh_shell).grid(row=0, column=1)

    def build_approvals_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        columns = ("created", "direction", "sender", "summary", "status")
        self.approvals_tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        for column, heading, width in (
            ("created", "Created", 160),
            ("direction", "Direction", 90),
            ("sender", "Sender", 120),
            ("summary", "Summary", 380),
            ("status", "Status", 140),
        ):
            self.approvals_tree.heading(column, text=heading)
            self.approvals_tree.column(column, width=width, anchor="w")
        self.approvals_tree.grid(row=0, column=0, sticky="nsew")
        self.approvals_tree.bind("<<TreeviewSelect>>", lambda _event: self.show_selected_command_details())

        detail_frame = ttk.LabelFrame(parent, text="Selected Request", padding=8)
        detail_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        detail_frame.columnconfigure(0, weight=1)
        self.approval_detail_text = tk.Text(detail_frame, height=12, wrap="word", state="disabled")
        self.approval_detail_text.grid(row=0, column=0, columnspan=3, sticky="ew")
        ttk.Button(detail_frame, text="Approve", command=self.approve_selected_command).grid(row=1, column=1, sticky="e", pady=(8, 0), padx=(0, 8))
        ttk.Button(detail_frame, text="Deny", command=self.deny_selected_command).grid(row=1, column=2, sticky="e", pady=(8, 0))

    def build_explorer_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        header = ttk.Frame(parent)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(5, weight=1)
        ttk.Label(header, text="Remote Node").grid(row=0, column=0, sticky="w")
        ttk.Combobox(header, textvariable=self.explorer_target_var, values=("home", "office"), state="readonly", width=12).grid(row=0, column=1, sticky="w", padx=(6, 12))
        ttk.Button(header, text="Refresh Both", command=self.refresh_explorer).grid(row=0, column=2, sticky="w")
        ttk.Button(header, text="Test SSH", command=self.test_explorer_remote).grid(row=0, column=3, sticky="w", padx=(8, 0))
        ttk.Button(header, text="Open Shell", command=self.open_explorer_shell).grid(row=0, column=4, sticky="w", padx=(8, 0))
        ttk.Label(header, text="Browse both PCs and transfer directly over SSH.").grid(row=0, column=5, sticky="e")

        panes = ttk.Panedwindow(parent, orient="horizontal")
        panes.grid(row=1, column=0, sticky="nsew")

        local_frame = ttk.LabelFrame(panes, text="This PC", padding=8)
        remote_frame = ttk.LabelFrame(panes, text="Remote PC", padding=8)
        panes.add(local_frame, weight=1)
        panes.add(remote_frame, weight=1)

        self.build_local_browser(local_frame)
        self.build_remote_browser(remote_frame)

        actions = ttk.Frame(parent)
        actions.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        for column in range(4):
            actions.columnconfigure(column, weight=1)
        ttk.Button(actions, text="Send Selected -> Remote Folder", command=self.transfer_local_to_remote).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(actions, text="<- Pull Selected To Local Folder", command=self.transfer_remote_to_local).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(actions, text="New Folder", command=self.create_folder_in_active_pane).grid(row=0, column=2, sticky="ew", padx=6)
        ttk.Button(actions, text="Rename / Delete", command=self.rename_or_delete_selected).grid(row=0, column=3, sticky="ew", padx=(6, 0))

    def build_local_browser(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        path_row = ttk.Frame(parent)
        path_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        path_row.columnconfigure(1, weight=1)
        ttk.Button(path_row, text="Up", command=self.local_up).grid(row=0, column=0, padx=(0, 6))
        ttk.Entry(path_row, textvariable=self.local_path_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(path_row, text="Go", command=self.refresh_local_browser).grid(row=0, column=2, padx=(6, 0))

        columns = ("name", "type", "size", "modified")
        self.local_tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        for column, heading, width in (
            ("name", "Name", 220),
            ("type", "Type", 80),
            ("size", "Size", 90),
            ("modified", "Modified", 150),
        ):
            self.local_tree.heading(column, text=heading)
            self.local_tree.column(column, width=width, anchor="w")
        self.local_tree.grid(row=1, column=0, sticky="nsew")
        self.local_tree.bind("<Double-1>", lambda _event: self.open_local_selected())

    def build_remote_browser(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        path_row = ttk.Frame(parent)
        path_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        path_row.columnconfigure(1, weight=1)
        ttk.Button(path_row, text="Up", command=self.remote_up).grid(row=0, column=0, padx=(0, 6))
        ttk.Entry(path_row, textvariable=self.remote_path_var).grid(row=0, column=1, sticky="ew")
        ttk.Button(path_row, text="Go", command=self.refresh_remote_browser).grid(row=0, column=2, padx=(6, 0))

        columns = ("name", "type", "size", "modified")
        self.remote_tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        for column, heading, width in (
            ("name", "Name", 220),
            ("type", "Type", 80),
            ("size", "Size", 90),
            ("modified", "Modified", 150),
        ):
            self.remote_tree.heading(column, text=heading)
            self.remote_tree.column(column, width=width, anchor="w")
        self.remote_tree.grid(row=1, column=0, sticky="nsew")
        self.remote_tree.bind("<Double-1>", lambda _event: self.open_remote_selected())

    def build_projects_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        focus_frame = ttk.LabelFrame(parent, text="Local Focus", padding=8)
        focus_frame.grid(row=0, column=0, sticky="ew")
        focus_frame.columnconfigure(1, weight=1)

        ttk.Label(focus_frame, text="Project").grid(row=0, column=0, sticky="w")
        ttk.Entry(focus_frame, textvariable=self.focus_project_var).grid(row=0, column=1, sticky="ew", pady=(0, 8))
        ttk.Label(focus_frame, text="What You're Doing").grid(row=1, column=0, sticky="nw")
        self.focus_note_text = tk.Text(focus_frame, height=4, wrap="word")
        self.focus_note_text.grid(row=1, column=1, sticky="ew")
        self.focus_note_text.insert("1.0", self.config.focus_note)
        ttk.Button(focus_frame, text="Publish Focus Update", command=self.publish_focus_update).grid(row=2, column=1, sticky="e", pady=(8, 0))

        projects_frame = ttk.LabelFrame(parent, text="Known Projects Across Both PCs", padding=8)
        projects_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        projects_frame.columnconfigure(0, weight=1)
        projects_frame.rowconfigure(0, weight=1)

        columns = ("project", "nodes", "last_update", "last_note")
        self.projects_tree = ttk.Treeview(projects_frame, columns=columns, show="headings")
        for column, heading, width in (
            ("project", "Project", 220),
            ("nodes", "Known On", 140),
            ("last_update", "Last Update", 160),
            ("last_note", "Latest Note", 360),
        ):
            self.projects_tree.heading(column, text=heading)
            self.projects_tree.column(column, width=width, anchor="w")
        self.projects_tree.grid(row=0, column=0, sticky="nsew")

    def build_setup_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        top = ttk.LabelFrame(parent, text="Bridge Setup", padding=8)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Bridge Folder").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.bridge_root_var).grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Button(top, text="Choose", command=self.choose_bridge_folder).grid(row=0, column=2, sticky="ew")

        ttk.Label(top, text="Local Node").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Combobox(top, textvariable=self.local_node_var, values=("home", "office"), state="readonly").grid(row=1, column=1, sticky="w", padx=(6, 0), pady=(8, 0))
        ttk.Button(top, text="Save Setup", command=self.save_setup).grid(row=1, column=2, sticky="e", pady=(8, 0))

        auto_frame = ttk.LabelFrame(parent, text="Auto Reply", padding=8)
        auto_frame.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        auto_frame.columnconfigure(1, weight=1)
        ttk.Checkbutton(auto_frame, text="Reply automatically to new chat messages targeted to this node", variable=self.auto_reply_enabled_var).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(auto_frame, text="Model").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(auto_frame, textvariable=self.auto_reply_model_var).grid(row=1, column=1, sticky="ew", padx=(6, 0), pady=(8, 0))
        ttk.Label(auto_frame, text="Tip: use chat target 'home' or 'office' for auto replies. API key is read from OPENAI_API_KEY or ~/.codex_bridge_secrets.json.").grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        access_frame = ttk.LabelFrame(parent, text="Direct SSH Access", padding=8)
        access_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        access_frame.columnconfigure(1, weight=1)
        access_frame.columnconfigure(3, weight=1)
        ttk.Label(access_frame, text="Home Host").grid(row=0, column=0, sticky="w")
        ttk.Entry(access_frame, textvariable=self.home_host_var).grid(row=0, column=1, sticky="ew", padx=(6, 12))
        ttk.Label(access_frame, text="Home User").grid(row=0, column=2, sticky="w")
        ttk.Entry(access_frame, textvariable=self.home_user_var).grid(row=0, column=3, sticky="ew", padx=(6, 0))
        ttk.Label(access_frame, text="Home Key").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(access_frame, textvariable=self.home_key_var).grid(row=1, column=1, columnspan=3, sticky="ew", padx=(6, 0), pady=(8, 0))
        ttk.Label(access_frame, text="Office Host").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(access_frame, textvariable=self.office_host_var).grid(row=2, column=1, sticky="ew", padx=(6, 12), pady=(8, 0))
        ttk.Label(access_frame, text="Office User").grid(row=2, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(access_frame, textvariable=self.office_user_var).grid(row=2, column=3, sticky="ew", padx=(6, 0), pady=(8, 0))
        ttk.Label(access_frame, text="Office Key").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(access_frame, textvariable=self.office_key_var).grid(row=3, column=1, columnspan=3, sticky="ew", padx=(6, 0), pady=(8, 0))
        ttk.Label(access_frame, text="These profiles power direct remote shell access without going through bridge chat.").grid(row=4, column=0, columnspan=4, sticky="w", pady=(8, 0))

        roots_frame = ttk.LabelFrame(parent, text="Local Project Roots To Publish", padding=8)
        roots_frame.grid(row=4, column=0, sticky="nsew", pady=(10, 0))
        roots_frame.columnconfigure(0, weight=1)
        roots_frame.rowconfigure(0, weight=1)

        self.project_roots_list = tk.Listbox(roots_frame, height=12)
        self.project_roots_list.grid(row=0, column=0, sticky="nsew")
        for path in self.config.project_roots:
            self.project_roots_list.insert("end", path)

        root_buttons = ttk.Frame(roots_frame)
        root_buttons.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(root_buttons, text="Add Root", command=self.add_project_root).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(root_buttons, text="Remove Selected", command=self.remove_project_root).grid(row=0, column=1)

    def build_log_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)
        self.log_text = tk.Text(parent, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")

    def build_data_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)

        source = ttk.LabelFrame(parent, text="Jobs Source", padding=8)
        source.grid(row=0, column=0, sticky="ew")
        source.columnconfigure(1, weight=1)
        ttk.Label(source, text="jobs.csv").grid(row=0, column=0, sticky="w")
        ttk.Entry(source, textvariable=self.jobs_csv_var).grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Button(source, text="Browse", command=self.choose_jobs_csv).grid(row=0, column=2)

        lookup = ttk.LabelFrame(parent, text="RO Lookup", padding=8)
        lookup.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        lookup.columnconfigure(1, weight=1)
        ttk.Label(lookup, text="RO Number").grid(row=0, column=0, sticky="w")
        ttk.Entry(lookup, textvariable=self.ro_lookup_var).grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Button(lookup, text="Lookup RO", command=self.lookup_ro).grid(row=0, column=2)

        results = ttk.LabelFrame(parent, text="Result", padding=8)
        results.grid(row=2, column=0, sticky="nsew", pady=(10, 0))
        results.columnconfigure(0, weight=1)
        results.rowconfigure(0, weight=1)
        self.ro_result_text = tk.Text(results, wrap="word", state="disabled")
        self.ro_result_text.grid(row=0, column=0, sticky="nsew")

    def setup_tray(self) -> None:
        if not TRAY_AVAILABLE:
            return
        if self.tray_started:
            return
        image = self.build_tray_image()
        menu = pystray.Menu(
            pystray.MenuItem("Show Bridge", self.on_tray_show),
            pystray.MenuItem("Hide Bridge", self.on_tray_hide),
            pystray.MenuItem("Exit Bridge", self.on_tray_exit),
        )
        self.tray_icon = pystray.Icon("codex-bridge", image, "Codex Bridge", menu)
        self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        self.tray_thread.start()
        self.tray_started = True

    def build_tray_image(self) -> Image.Image:
        if not TRAY_AVAILABLE or Image is None or ImageDraw is None:
            raise RuntimeError("Tray icons are not available on this system.")
        image = Image.new("RGB", (64, 64), "#14213d")
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((6, 6, 58, 58), radius=12, fill="#fca311")
        draw.text((18, 18), "CB", fill="#14213d")
        return image

    def on_tray_show(self, icon: pystray.Icon | None = None, item: object | None = None) -> None:
        self.root.after(0, self.show_from_tray)

    def on_tray_hide(self, icon: pystray.Icon | None = None, item: object | None = None) -> None:
        self.root.after(0, self.hide_to_tray)

    def on_tray_exit(self, icon: pystray.Icon | None = None, item: object | None = None) -> None:
        self.root.after(0, self.exit_app)

    def show_from_tray(self) -> None:
        self.is_hidden_to_tray = False
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.status_var.set("Codex Bridge restored from tray.")

    def hide_to_tray(self) -> None:
        if not TRAY_AVAILABLE:
            self.status_var.set("Tray support is unavailable on this Python build. Leaving the bridge window open will keep it running.")
            messagebox.showinfo("Tray Unavailable", "Tray support is unavailable on this Python build. Leave the bridge window open to keep it running.")
            return
        self.is_hidden_to_tray = True
        self.root.withdraw()
        self.status_var.set("Codex Bridge is still running in the system tray.")

    def exit_app(self) -> None:
        if self.tray_icon is not None:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
        self.root.destroy()

    def choose_bridge_folder(self) -> None:
        path = filedialog.askdirectory(title="Choose the synced bridge folder", initialdir=str(self.bridge_root()))
        if path:
            self.bridge_root_var.set(path)
            self.ensure_bridge_dirs()
            self.status_var.set(f"Bridge folder set to {path}")

    def choose_command_workdir(self) -> None:
        path = filedialog.askdirectory(title="Choose command working directory")
        if path:
            self.command_workdir_var.set(path)

    def direct_access_profile(self, target: str) -> tuple[str, str, str]:
        if target == "home":
            return (
                self.home_host_var.get().strip(),
                self.home_user_var.get().strip() or "ferna",
                self.home_key_var.get().strip(),
            )
        return (
            self.office_host_var.get().strip(),
            self.office_user_var.get().strip() or "ferna",
            self.office_key_var.get().strip(),
        )

    def validate_direct_access_profile(self, target: str) -> tuple[str, str, str] | None:
        host, user, key = self.direct_access_profile(target)
        if not host or not user or not key:
            messagebox.showinfo("Missing Access Profile", f"Set the {target} SSH host, user, and key in Setup first.")
            return None
        if not Path(key).exists():
            messagebox.showinfo("Missing Key File", f"The configured {target} SSH key does not exist:\n{key}")
            return None
        return host, user, key

    def encoded_powershell(self, command_text: str) -> str:
        return base64.b64encode(command_text.encode("utf-16le")).decode("ascii")

    def run_direct_ssh_command(self) -> None:
        command_text = self.command_text.get("1.0", "end").strip()
        target = self.command_target_var.get().strip()
        if not command_text:
            messagebox.showinfo("No Command", "Type a PowerShell command first.")
            return
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        summary = self.command_summary_var.get().strip() or command_text.splitlines()[0][:120]
        thread = threading.Thread(
            target=self.execute_direct_ssh_command,
            args=(target, host, user, key, command_text, summary),
            daemon=True,
        )
        thread.start()
        self.status_var.set(f"Running direct SSH command on {target}: {summary}")

    def execute_direct_ssh_command(
        self,
        target: str,
        host: str,
        user: str,
        key: str,
        command_text: str,
        summary: str,
        announce_result: bool = False,
    ) -> None:
        try:
            encoded = self.encoded_powershell(command_text)
            completed = subprocess.run(
                [
                    r"C:\Windows\System32\OpenSSH\ssh.exe",
                    *SSH_NONINTERACTIVE_ARGS,
                    "-i",
                    key,
                    f"{user}@{host}",
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-EncodedCommand",
                    encoded,
                ],
                capture_output=True,
                text=True,
                timeout=600,
                creationflags=CREATE_NO_WINDOW,
            )
            status_text = f"Direct SSH command on {target} finished with exit code {completed.returncode}: {summary}"
            self.log(status_text)
            if completed.stdout.strip():
                self.log("Direct SSH stdout:\n" + completed.stdout[-4000:])
            if completed.stderr.strip():
                self.log("Direct SSH stderr:\n" + completed.stderr[-4000:])
            self.root.after(0, lambda: self.status_var.set(status_text))
            if announce_result:
                self.send_system_message(f"Direct command on {target} finished with exit code {completed.returncode}: {summary}")
        except subprocess.TimeoutExpired:
            self.log(f"Direct SSH command timed out on {target}: {summary}")
            self.root.after(0, lambda: self.status_var.set(f"Direct SSH command timed out on {target}."))
            if announce_result:
                self.send_system_message(f"Direct command on {target} timed out: {summary}")
        except Exception as exc:
            self.log(f"Direct SSH command failed on {target}: {exc}")
            self.root.after(0, lambda: self.status_var.set(f"Direct SSH command failed on {target}."))
            if announce_result:
                self.send_system_message(f"Direct command on {target} failed: {summary}")

    def open_direct_ssh_shell(self) -> None:
        target = self.command_target_var.get().strip()
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        try:
            subprocess.Popen(
                [r"C:\Windows\System32\OpenSSH\ssh.exe", *SSH_NONINTERACTIVE_ARGS, "-i", key, f"{user}@{host}"],
                creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0),
            )
            self.status_var.set(f"Opened direct SSH shell to {target}.")
        except Exception as exc:
            self.log(f"Failed to open direct SSH shell to {target}: {exc}")
            messagebox.showerror("SSH Launch Failed", f"Could not open a direct SSH shell to {target}.\n\n{exc}")

    def powershell_quote(self, text: str) -> str:
        return "'" + text.replace("'", "''") + "'"

    def format_size(self, value: int | None) -> str:
        if not value:
            return ""
        amount = float(value)
        units = ["B", "KB", "MB", "GB", "TB"]
        index = 0
        while amount >= 1024 and index < len(units) - 1:
            amount /= 1024
            index += 1
        return f"{amount:.1f} {units[index]}" if index else f"{int(amount)} {units[index]}"

    def run_remote_json(self, target: str, command_text: str, timeout: int = 120) -> list[dict]:
        profile = self.validate_direct_access_profile(target)
        if not profile:
            raise RuntimeError(f"Missing {target} SSH profile.")
        host, user, key = profile
        encoded = self.encoded_powershell(command_text)
        completed = subprocess.run(
            [
                r"C:\Windows\System32\OpenSSH\ssh.exe",
                *SSH_NONINTERACTIVE_ARGS,
                "-i",
                key,
                f"{user}@{host}",
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-EncodedCommand",
                encoded,
            ],
            capture_output=True,
            text=True,
            timeout=timeout,
            creationflags=CREATE_NO_WINDOW,
        )
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or f"SSH failed with exit code {completed.returncode}")
        text = completed.stdout.strip()
        if not text:
            return []
        payload = json.loads(text)
        if isinstance(payload, list):
            return payload
        if isinstance(payload, dict):
            return [payload]
        return []

    def remote_scp_path(self, path_text: str) -> str:
        return path_text.replace("\\", "/")

    def local_dir_entries(self, path: Path) -> list[dict]:
        entries: list[dict] = []
        for item in sorted(path.iterdir(), key=lambda child: (not child.is_dir(), child.name.lower())):
            try:
                is_dir = item.is_dir()
                stat = item.stat()
            except OSError:
                continue
            entries.append(
                {
                    "Name": item.name,
                    "FullPath": str(item),
                    "IsDirectory": is_dir,
                    "Size": None if is_dir else stat.st_size,
                    "Modified": datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
                }
            )
        return entries

    def populate_tree(self, tree: ttk.Treeview, entries: list[dict]) -> None:
        for item in tree.get_children():
            tree.delete(item)
        for index, entry in enumerate(entries):
            tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    entry.get("Name", ""),
                    "Folder" if entry.get("IsDirectory") else "File",
                    self.format_size(entry.get("Size")),
                    entry.get("Modified", ""),
                ),
            )

    def selected_entry(self, tree: ttk.Treeview, entries: list[dict]) -> dict | None:
        selection = tree.selection()
        if not selection:
            return None
        try:
            return entries[int(selection[0])]
        except Exception:
            return None

    def refresh_explorer(self) -> None:
        self.refresh_local_browser()
        self.refresh_remote_browser()

    def refresh_local_browser(self) -> None:
        path = Path(self.local_path_var.get().strip() or str(Path.home())).resolve()
        self.local_path_var.set(str(path))
        self.local_entries = self.local_dir_entries(path)
        self.populate_tree(self.local_tree, self.local_entries)

    def refresh_remote_browser(self) -> None:
        if self.remote_refresh_inflight:
            return
        target = self.explorer_target_var.get().strip() or self.opposite_node(self.local_node_var.get().strip() or "home")
        path_text = self.remote_path_var.get().strip() or r"C:\Users\ferna"
        self.remote_refresh_inflight = True
        self.remote_refresh_generation += 1
        generation = self.remote_refresh_generation

        def worker() -> None:
            try:
                script = (
                    f"$path = {self.powershell_quote(path_text)}; "
                    "if (-not (Test-Path -LiteralPath $path)) { throw \"Path not found: $path\" }; "
                    "Get-ChildItem -LiteralPath $path | "
                    "Sort-Object @{Expression='PSIsContainer';Descending=$true}, Name | "
                    "Select-Object "
                    "@{Name='Name';Expression={$_.Name}},"
                    "@{Name='FullPath';Expression={$_.FullName}},"
                    "@{Name='IsDirectory';Expression={[bool]$_.PSIsContainer}},"
                    "@{Name='Size';Expression={if ($_.PSIsContainer) { $null } else { [int64]$_.Length }}},"
                    "@{Name='Modified';Expression={$_.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')}} | "
                    "ConvertTo-Json -Compress"
                )
                entries = self.run_remote_json(target, script)

                def apply_entries() -> None:
                    if generation != self.remote_refresh_generation:
                        return
                    self.remote_entries = entries
                    self.populate_tree(self.remote_tree, self.remote_entries)
                    self.status_var.set(f"Remote browser refreshed for {target}.")

                self.root.after(0, apply_entries)
            except Exception as exc:
                self.root.after(0, lambda: self.log(f"Remote browser refresh failed for {target}: {exc}"))
                self.root.after(0, lambda: self.status_var.set(f"Remote browser refresh failed for {target}."))
            finally:
                self.root.after(0, lambda: setattr(self, "remote_refresh_inflight", False))

        threading.Thread(target=worker, daemon=True).start()

    def local_up(self) -> None:
        current = Path(self.local_path_var.get().strip() or str(Path.home()))
        parent = current.parent if current.parent != current else current
        self.local_path_var.set(str(parent))
        self.refresh_local_browser()

    def remote_up(self) -> None:
        current = Path(self.remote_path_var.get().strip() or r"C:\Users\ferna")
        parent = current.parent if current.parent != current else current
        self.remote_path_var.set(str(parent))
        self.refresh_remote_browser()

    def open_local_selected(self) -> None:
        entry = self.selected_entry(self.local_tree, self.local_entries)
        if entry and entry.get("IsDirectory"):
            self.local_path_var.set(entry["FullPath"])
            self.refresh_local_browser()

    def open_remote_selected(self) -> None:
        entry = self.selected_entry(self.remote_tree, self.remote_entries)
        if entry and entry.get("IsDirectory"):
            self.remote_path_var.set(entry["FullPath"])
            self.refresh_remote_browser()

    def transfer_local_to_remote(self) -> None:
        entry = self.selected_entry(self.local_tree, self.local_entries)
        if not entry:
            messagebox.showinfo("No Selection", "Pick a local file or folder first.")
            return
        target = self.explorer_target_var.get().strip() or self.opposite_node(self.local_node_var.get().strip() or "home")
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        self.status_var.set(f"Sending {entry['Name']} to {target}...")
        threading.Thread(
            target=self.execute_scp_transfer,
            args=("upload", target, host, user, key, entry["FullPath"], self.remote_path_var.get().strip(), entry["Name"]),
            daemon=True,
        ).start()

    def transfer_remote_to_local(self) -> None:
        entry = self.selected_entry(self.remote_tree, self.remote_entries)
        if not entry:
            messagebox.showinfo("No Selection", "Pick a remote file or folder first.")
            return
        target = self.explorer_target_var.get().strip() or self.opposite_node(self.local_node_var.get().strip() or "home")
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        self.status_var.set(f"Pulling {entry['Name']} from {target}...")
        threading.Thread(
            target=self.execute_scp_transfer,
            args=("download", target, host, user, key, entry["FullPath"], self.local_path_var.get().strip(), entry["Name"]),
            daemon=True,
        ).start()

    def test_explorer_remote(self) -> None:
        target = self.explorer_target_var.get().strip() or self.opposite_node(self.local_node_var.get().strip() or "home")
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        self.status_var.set(f"Testing SSH to {target}...")
        threading.Thread(target=self.execute_ssh_test, args=(target, host, user, key), daemon=True).start()

    def execute_ssh_test(self, target: str, host: str, user: str, key: str) -> None:
        try:
            completed = subprocess.run(
                [r"C:\Windows\System32\OpenSSH\ssh.exe", *SSH_NONINTERACTIVE_ARGS, "-i", key, f"{user}@{host}", "hostname"],
                capture_output=True,
                text=True,
                timeout=120,
                creationflags=CREATE_NO_WINDOW,
            )
            if completed.returncode != 0:
                raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or f"SSH failed with exit code {completed.returncode}")
            host_name = completed.stdout.strip() or host
            self.root.after(0, lambda: messagebox.showinfo("SSH Connected", f"{target.title()} is reachable.\n\nHost replied: {host_name}"))
            self.root.after(0, lambda: self.status_var.set(f"SSH test succeeded for {target}: {host_name}"))
        except Exception as exc:
            self.log(f"SSH test failed for {target}: {exc}")
            self.root.after(0, lambda: messagebox.showerror("SSH Test Failed", str(exc)))
            self.root.after(0, lambda: self.status_var.set(f"SSH test failed for {target}."))

    def open_explorer_shell(self) -> None:
        self.command_target_var.set(self.explorer_target_var.get().strip() or self.command_target_var.get())
        self.open_direct_ssh_shell()

    def create_folder_in_active_pane(self) -> None:
        folder_name = simpledialog.askstring("New Folder", "Folder name:")
        if not folder_name:
            return
        target = self.explorer_target_var.get().strip() or self.opposite_node(self.local_node_var.get().strip() or "home")
        active_remote = bool(self.remote_tree.selection())
        active_local = bool(self.local_tree.selection()) or not active_remote
        try:
            if active_remote and not active_local:
                parent_path = Path(self.remote_path_var.get().strip() or r"C:\Users\ferna") / folder_name
                script = f"New-Item -ItemType Directory -Path {self.powershell_quote(str(parent_path))} -Force | Out-Null"
                self.run_remote_json(target, script)
                self.refresh_remote_browser()
                self.status_var.set(f"Created remote folder: {parent_path}")
            else:
                new_path = Path(self.local_path_var.get().strip() or str(Path.home())) / folder_name
                new_path.mkdir(parents=True, exist_ok=True)
                self.refresh_local_browser()
                self.status_var.set(f"Created local folder: {new_path}")
        except Exception as exc:
            messagebox.showerror("Create Folder Failed", str(exc))

    def rename_or_delete_selected(self) -> None:
        remote_entry = self.selected_entry(self.remote_tree, self.remote_entries)
        local_entry = self.selected_entry(self.local_tree, self.local_entries)
        entry = remote_entry or local_entry
        if not entry:
            messagebox.showinfo("No Selection", "Pick a file or folder first.")
            return
        action = messagebox.askyesnocancel("Rename or Delete", "Yes = Rename\nNo = Delete\nCancel = do nothing")
        if action is None:
            return
        try:
            if action:
                new_name = simpledialog.askstring("Rename", f"New name for {entry['Name']}:")
                if not new_name:
                    return
                self.rename_entry(entry, bool(remote_entry), new_name)
            else:
                confirm = messagebox.askyesno("Confirm Delete", f"Delete {entry['Name']}?")
                if confirm:
                    self.delete_entry(entry, bool(remote_entry))
        except Exception as exc:
            messagebox.showerror("Explorer Action Failed", str(exc))

    def rename_entry(self, entry: dict, is_remote: bool, new_name: str) -> None:
        old_path = Path(entry["FullPath"])
        new_path = old_path.parent / new_name
        if is_remote:
            target = self.explorer_target_var.get().strip() or self.opposite_node(self.local_node_var.get().strip() or "home")
            script = (
                f"Rename-Item -LiteralPath {self.powershell_quote(str(old_path))} "
                f"-NewName {self.powershell_quote(new_name)}"
            )
            self.run_remote_json(target, script)
            self.refresh_remote_browser()
            self.status_var.set(f"Renamed remote item to {new_name}")
        else:
            old_path.rename(new_path)
            self.refresh_local_browser()
            self.status_var.set(f"Renamed local item to {new_name}")

    def delete_entry(self, entry: dict, is_remote: bool) -> None:
        path = Path(entry["FullPath"])
        if is_remote:
            target = self.explorer_target_var.get().strip() or self.opposite_node(self.local_node_var.get().strip() or "home")
            script = f"Remove-Item -LiteralPath {self.powershell_quote(str(path))} -Recurse -Force"
            self.run_remote_json(target, script)
            self.refresh_remote_browser()
            self.status_var.set(f"Deleted remote item: {entry['Name']}")
        else:
            if path.is_dir():
                shutil.rmtree(path)
            else:
                path.unlink()
            self.refresh_local_browser()
            self.status_var.set(f"Deleted local item: {entry['Name']}")

    def choose_jobs_csv(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose Mitchell jobs.csv",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            initialdir=str(DEFAULT_JOBS_CSV.parent),
        )
        if path:
            self.jobs_csv_var.set(path)

    def load_jobs_csv(self) -> list[dict]:
        path = Path(self.jobs_csv_var.get().strip())
        if not path.exists():
            raise FileNotFoundError(f"jobs.csv not found: {path}")
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            return list(csv.DictReader(handle))

    def lookup_ro(self) -> None:
        ro = self.ro_lookup_var.get().strip()
        if not ro:
            messagebox.showinfo("No RO Number", "Type an RO number first.")
            return
        try:
            rows = self.load_jobs_csv()
            match = next((row for row in rows if str(row.get("RONumber", "")).strip() == ro), None)
            if not match:
                self.set_text(self.ro_result_text, f"No RO found for {ro}.")
                self.status_var.set(f"RO {ro} not found.")
                return
            lines = [
                f"RO Number: {match.get('RONumber', '')}",
                f"Customer: {match.get('CustomerLastOrCompanyName', '')}, {match.get('CustomerFirstName', '')}".strip(", "),
                f"Vehicle: {match.get('VehicleYear', '')} {match.get('VehicleMake', '')} {match.get('VehicleModel', '')}".strip(),
                f"Total Amount: {match.get('TotalAmount', '')}",
                f"Balance Due: {match.get('BalanceDue', '')}",
                f"Due In: {match.get('DueInDate', '')}",
                f"Due Out: {match.get('DueOutDate', '')}",
                f"Insurance: {match.get('InsuranceCompanyName', '')}",
                f"Claim Number: {match.get('ClaimNumber', '')}",
                f"Estimator: {match.get('EstimatorFullName', '')}",
                f"VIN: {match.get('VehicleVin', '')}",
                f"Plate: {match.get('VehicleLicense', '')}",
                f"EstimateId: {match.get('EstimateId', '')}",
            ]
            self.set_text(self.ro_result_text, "\n".join(lines))
            self.status_var.set(f"Loaded RO {ro}.")
        except Exception as exc:
            messagebox.showerror("RO Lookup Failed", str(exc))

    def execute_scp_transfer(self, mode: str, target: str, host: str, user: str, key: str, source_path: str, dest_path: str, display_name: str) -> None:
        try:
            scp_exe = r"C:\Windows\System32\OpenSSH\scp.exe"
            if mode == "upload":
                command = [scp_exe, *SSH_NONINTERACTIVE_ARGS, "-r", "-i", key, source_path, f"{user}@{host}:{self.remote_scp_path(dest_path)}"]
            else:
                Path(dest_path).mkdir(parents=True, exist_ok=True)
                command = [scp_exe, *SSH_NONINTERACTIVE_ARGS, "-r", "-i", key, f"{user}@{host}:{self.remote_scp_path(source_path)}", dest_path]
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=3600,
                creationflags=CREATE_NO_WINDOW,
            )
            if completed.returncode != 0:
                raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or f"SCP failed with exit code {completed.returncode}")
            self.log(f"{mode.title()} complete for {display_name} on {target}.")
            if mode == "upload":
                success_text = f"Sent {display_name} to {target} folder:\n{dest_path}"
            else:
                success_text = f"Pulled {display_name} from {target} into:\n{dest_path}"
            self.root.after(0, lambda: messagebox.showinfo("Transfer Complete", success_text))
            self.root.after(0, lambda: self.status_var.set(f"{mode.title()} complete: {display_name}"))
            self.root.after(0, self.refresh_explorer)
        except Exception as exc:
            self.log(f"{mode.title()} failed for {display_name} on {target}: {exc}")
            self.root.after(0, lambda: messagebox.showerror("Transfer Failed", str(exc)))
            self.root.after(0, lambda: self.status_var.set(f"{mode.title()} failed for {display_name}."))

    def on_node_changed(self) -> None:
        node = self.local_node_var.get().strip() or "home"
        if not self.display_name_var.get().strip():
            self.display_name_var.set(self.default_display_name(node))
        self.command_target_var.set(self.opposite_node(node))

    def opposite_node(self, node: str) -> str:
        return "office" if node == "home" else "home"

    def save_setup(self) -> None:
        self.save_config()
        self.ensure_bridge_dirs()
        self.publish_local_state()
        self.status_var.set("Bridge setup saved.")

    def add_project_root(self) -> None:
        path = filedialog.askdirectory(title="Choose project root to publish")
        if not path:
            return
        existing = set(self.project_roots_list.get(0, "end"))
        if path in existing:
            messagebox.showinfo("Already Added", "That project root is already in the list.")
            return
        self.project_roots_list.insert("end", path)

    def remove_project_root(self) -> None:
        selection = self.project_roots_list.curselection()
        if not selection:
            return
        self.project_roots_list.delete(selection[0])

    def load_secrets(self) -> dict:
        if not SECRETS_FILE.exists():
            return {}
        try:
            return json.loads(SECRETS_FILE.read_text(encoding="utf-8-sig"))
        except Exception:
            return {}

    def current_api_key(self) -> str:
        return os.environ.get("OPENAI_API_KEY", "").strip() or str(self.load_secrets().get("openai_api_key", "")).strip()

    def auto_reply_state_path(self) -> Path:
        local_node = self.local_node_var.get().strip() or "home"
        return self.state_dir() / f"{local_node}_auto_reply_state.json"

    def load_auto_reply_state(self) -> dict:
        path = self.auto_reply_state_path()
        if not path.exists():
            return {"initialized": False, "handled_ids": [], "updated_at": now_iso()}
        try:
            return json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            return {"initialized": False, "handled_ids": [], "updated_at": now_iso()}

    def save_auto_reply_state(self, state: dict) -> None:
        self.write_json(self.auto_reply_state_path(), state)

    def trusted_chat_state_path(self) -> Path:
        local_node = self.local_node_var.get().strip() or "home"
        return self.state_dir() / f"{local_node}_trusted_chat_state.json"

    def load_trusted_chat_state(self) -> dict:
        path = self.trusted_chat_state_path()
        if not path.exists():
            return {"initialized": False, "handled_ids": [], "updated_at": now_iso()}
        try:
            return json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            return {"initialized": False, "handled_ids": [], "updated_at": now_iso()}

    def save_trusted_chat_state(self, state: dict) -> None:
        self.write_json(self.trusted_chat_state_path(), state)

    def is_trusted_peer_sender(self, sender: str) -> bool:
        local_node = (self.local_node_var.get().strip() or "home").lower()
        sender = (sender or "").strip().lower()
        if sender not in {"home", "office"} or sender == local_node:
            return False
        return True

    def should_auto_execute_trusted_commands(self) -> bool:
        return bool(getattr(self.config, "auto_execute_trusted_commands", True))

    def should_auto_convert_trusted_chat_commands(self) -> bool:
        return bool(getattr(self.config, "auto_convert_trusted_chat_commands", True))

    def recent_chat_context(self, limit: int = 12) -> str:
        lines: list[str] = []
        for item in self.latest_messages[-limit:]:
            if item.get("kind") != "chat":
                continue
            sender = item.get("sender_name") or item.get("sender") or "unknown"
            lines.append(f"{sender}: {item.get('text', '')}")
        return "\n".join(lines).strip()

    def request_auto_reply(self, message: dict) -> str:
        api_key = self.current_api_key()
        if not api_key:
            raise RuntimeError("No OpenAI API key configured for bridge auto reply.")

        local_node = self.local_node_var.get().strip() or "home"
        model = self.auto_reply_model_var.get().strip() or "gpt-4.1-mini"
        sender_name = message.get("sender_name") or message.get("sender") or "Remote User"
        sender_node = message.get("sender") or self.opposite_node(local_node)
        system_prompt = (
            f"You are {self.display_name_var.get().strip() or self.default_display_name(local_node)} running on the {local_node} PC "
            "through Codex Bridge. Reply helpfully and concisely to the other PC. "
            "Do not claim you executed commands or changed files unless that actually happened. "
            "If you need the user to target a different node or approve a command, say so plainly."
        )
        user_prompt = (
            f"Recent bridge chat:\n{self.recent_chat_context() or '(no recent chat)'}\n\n"
            f"New message for you from {sender_name} on the {sender_node} PC:\n{message.get('text', '')}\n\n"
            "Reply as a practical teammate in 2-6 sentences unless more detail is needed."
        )
        payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            "temperature": 0.2,
        }
        request = urllib.request.Request(
            "https://api.openai.com/v1/chat/completions",
            data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        with urllib.request.urlopen(request, timeout=90) as response:
            body = json.loads(response.read().decode("utf-8"))
        content = body["choices"][0]["message"]["content"]
        if isinstance(content, list):
            return "\n".join(part.get("text", "") for part in content if isinstance(part, dict)).strip()
        return str(content).strip()

    def auto_reply_to_messages(self) -> None:
        if not self.auto_reply_enabled_var.get():
            return
        local_node = self.local_node_var.get().strip() or "home"
        state = self.load_auto_reply_state()
        handled_ids = set(state.get("handled_ids", []))
        incoming = [
            item for item in self.latest_messages
            if item.get("kind") == "chat"
            and item.get("sender") not in ("", None, local_node)
            and item.get("target") == local_node
            and not bool(item.get("routed_action"))
            and self.parse_action_from_chat(str(item.get("text") or "")) is None
        ]
        if not state.get("initialized"):
            state["initialized"] = True
            state["handled_ids"] = sorted({item.get("id") for item in incoming if item.get("id")})
            state["updated_at"] = now_iso()
            self.save_auto_reply_state(state)
            return
        for item in incoming:
            message_id = item.get("id")
            if not message_id or message_id in handled_ids or message_id in self.active_auto_replies:
                continue
            self.active_auto_replies.add(message_id)
            thread = threading.Thread(target=self.generate_auto_reply, args=(item,), daemon=True)
            thread.start()
            break

    def generate_auto_reply(self, message: dict) -> None:
        message_id = message.get("id")
        try:
            reply_text = self.request_auto_reply(message)
            if not reply_text:
                raise RuntimeError("Model returned an empty reply.")
            response = {
                "id": make_id("msg"),
                "kind": "chat",
                "sender": self.local_node_var.get().strip() or "home",
                "sender_name": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
                "target": message.get("sender") or self.opposite_node(self.local_node_var.get().strip() or "home"),
                "text": reply_text,
                "reply_to": message_id,
                "created_at": now_iso(),
            }
            filename = f"{response['created_at'].replace(':', '').replace('-', '')}_{response['id']}.json"
            self.write_bridge_payload("messages", filename, response)
            self.root.after(0, lambda: self.status_var.set("Auto reply sent."))
        except Exception as exc:
            self.log(f"Auto reply failed: {exc}")
            self.root.after(0, lambda: self.status_var.set("Auto reply failed. Check the Log tab."))
        finally:
            state = self.load_auto_reply_state()
            handled_ids = set(state.get("handled_ids", []))
            if message_id:
                handled_ids.add(message_id)
            state["initialized"] = True
            state["handled_ids"] = sorted(handled_ids)
            state["updated_at"] = now_iso()
            self.save_auto_reply_state(state)
            if message_id in self.active_auto_replies:
                self.active_auto_replies.remove(message_id)
            self.root.after(0, self.refresh_all)

    def write_json(self, path: Path, payload: dict) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path = path.with_suffix(path.suffix + ".tmp")
        temp_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        temp_path.replace(path)

    def bridge_subdir(self, name: str) -> Path:
        return self.bridge_root() / name

    def direct_bridge_target(self, target: str) -> str | None:
        local = (self.local_node_var.get().strip() or "home").lower()
        target = (target or "all").lower()
        if target in {"home", "office"} and target != local:
            return target
        if target == "all":
            return self.opposite_node(local)
        return None

    def deliver_bridge_payload_direct(self, target: str, subdir: str, filename: str, payload: dict) -> None:
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        remote_dir = str(self.bridge_subdir(subdir)).replace("/", "\\")
        payload_bytes = json.dumps(payload, indent=2).encode("utf-8")
        payload_b64 = base64.b64encode(payload_bytes).decode("ascii")
        script = (
            f"$dir = {self.powershell_quote(remote_dir)}; "
            "New-Item -ItemType Directory -Force -Path $dir | Out-Null; "
            f"$dest = Join-Path $dir {self.powershell_quote(filename)}; "
            f"$bytes = [Convert]::FromBase64String('{payload_b64}'); "
            "[System.IO.File]::WriteAllBytes($dest, $bytes)"
        )
        encoded = self.encoded_powershell(script)
        completed = subprocess.run(
            [
                r"C:\Windows\System32\OpenSSH\ssh.exe",
                *SSH_NONINTERACTIVE_ARGS,
                "-i",
                key,
                f"{user}@{host}",
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-EncodedCommand",
                encoded,
            ],
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=CREATE_NO_WINDOW,
        )
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or f"SSH failed with exit code {completed.returncode}")

    def remote_bridge_path_exists(self, target: str, subdir: str, filename: str) -> bool:
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return False
        host, user, key = profile
        remote_path = str(self.bridge_subdir(subdir) / filename).replace("/", "\\")
        script = f"if (Test-Path {self.powershell_quote(remote_path)}) {{ 'present' }} else {{ 'missing' }}"
        encoded = self.encoded_powershell(script)
        completed = subprocess.run(
            [
                r"C:\Windows\System32\OpenSSH\ssh.exe",
                *SSH_NONINTERACTIVE_ARGS,
                "-i",
                key,
                f"{user}@{host}",
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-EncodedCommand",
                encoded,
            ],
            capture_output=True,
            text=True,
            timeout=60,
            creationflags=CREATE_NO_WINDOW,
        )
        return completed.returncode == 0 and "present" in (completed.stdout or "").lower()

    def trigger_remote_responder_once(self, target: str) -> None:
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        runtime_script = r"C:\Users\ferna\bridge_runtime\codex_bridge_responder.py"
        script = f"& 'C:\\Program Files\\Python311\\python.exe' {self.powershell_quote(runtime_script)} --once"
        encoded = self.encoded_powershell(script)
        completed = subprocess.run(
            [
                r"C:\Windows\System32\OpenSSH\ssh.exe",
                *SSH_NONINTERACTIVE_ARGS,
                "-i",
                key,
                f"{user}@{host}",
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-EncodedCommand",
                encoded,
            ],
            capture_output=True,
            text=True,
            timeout=90,
            creationflags=CREATE_NO_WINDOW,
        )
        if completed.returncode != 0:
            raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or f"SSH failed with exit code {completed.returncode}")

    def write_bridge_payload(self, subdir: str, filename: str, payload: dict, *, mirror: bool = True) -> None:
        path = self.bridge_subdir(subdir) / filename
        self.write_json(path, payload)
        if not mirror:
            return
        target = self.direct_bridge_target(str(payload.get("target", "all")))
        if not target:
            return
        def worker() -> None:
            try:
                self.deliver_bridge_payload_direct(target, subdir, filename, payload)
                if not self.remote_bridge_path_exists(target, subdir, filename):
                    raise RuntimeError(f"Remote bridge did not confirm receipt of {filename}.")
                if subdir == "messages" and str(payload.get("kind")) == "chat":
                    self.trigger_remote_responder_once(target)
            except Exception as exc:
                self.root.after(0, lambda: self.log(f"Direct bridge delivery to {target} failed for {filename}: {exc}"))

        if threading.current_thread() is threading.main_thread():
            threading.Thread(target=worker, daemon=True).start()
        else:
            worker()

    def read_json_files(self, directory: Path) -> list[dict]:
        items: list[dict] = []
        if not directory.exists():
            return items
        for path in sorted(directory.glob("*.json")):
            try:
                payload = json.loads(path.read_text(encoding="utf-8-sig"))
                payload["_path"] = str(path)
                items.append(payload)
            except Exception:
                continue
        return items

    def normalize_chat_request(self, text: str) -> str:
        return " ".join(text.strip().split())

    def parse_direct_command_from_chat(self, text: str) -> tuple[str, str] | None:
        stripped = text.strip()
        lower = stripped.lower()
        for prefix in ("run:", "ps:", "powershell:", "cmd:"):
            if lower.startswith(prefix):
                command_text = stripped[len(prefix):].strip()
                if command_text:
                    summary = command_text.splitlines()[0][:120]
                    return command_text, summary
        wish_body = self.extract_i_wish_body(text)
        if wish_body and self.looks_like_shell_command(wish_body):
            summary = wish_body.splitlines()[0][:120]
            return wish_body, summary
        return None

    def extract_i_wish_body(self, text: str) -> str:
        stripped = text.strip()
        lower = stripped.lower()
        prefixes = ("i wish ", "i wish:", "iwish ", "iwish:")
        for prefix in prefixes:
            if lower.startswith(prefix):
                return stripped[len(prefix):].strip()
        return ""

    def parse_action_from_chat(self, text: str) -> tuple[str, str, str] | None:
        stripped = text.strip()
        lower = stripped.lower()
        if any(phrase in lower for phrase in ("print a test page", "print test page", "print the test page")):
            return ("print_test_page", "", "Print a test page")
        if any(phrase in lower for phrase in ("what apps are running", "what programs are running", "list running apps", "list running processes", "show running processes")):
            return ("shell_command", "Get-Process | Sort-Object ProcessName | Select-Object -First 120 ProcessName,Id,MainWindowTitle | Format-Table -AutoSize", "List running processes")
        if lower.startswith("open folder ") or lower.startswith("open the folder "):
            folder_text = stripped.split("folder", 1)[1].strip()
            if folder_text:
                quoted = self.powershell_quote(folder_text)
                return ("shell_command", f"if (Test-Path {quoted}) {{ Start-Process explorer.exe {quoted}; Write-Output 'Opened folder.' }} else {{ Write-Error 'Folder not found.' }}", f"Open folder {folder_text}")
        direct = self.parse_direct_command_from_chat(text)
        if direct:
            command_text, summary = direct
            return ("shell_command", command_text, summary)
        return None

    def looks_like_shell_command(self, text: str) -> bool:
        stripped = text.strip()
        if not stripped:
            return False
        lower = stripped.lower()
        power_shell_verbs = (
            "get-",
            "set-",
            "new-",
            "remove-",
            "copy-",
            "move-",
            "rename-",
            "restart-",
            "start-",
            "stop-",
            "test-",
            "select-",
            "where-",
            "invoke-",
            "add-",
            "clear-",
            "enable-",
            "disable-",
        )
        shell_starters = (
            "dir ",
            "cd ",
            "ls ",
            "pwd",
            "echo ",
            "type ",
            "cat ",
            "python ",
            "py ",
            "git ",
            "ssh ",
            "scp ",
            "robocopy ",
            ".\\",
            "& ",
        )
        if lower.startswith(power_shell_verbs) or lower.startswith(shell_starters):
            return True
        if any(token in stripped for token in ("|", ";", ">", "<", "$env:", ".ps1", ".bat", ".cmd")):
            return True
        if "\\" in stripped and ":" in stripped:
            return True
        return False

    def is_task_like_chat(self, text: str, target: str) -> bool:
        if target not in {"home", "office"}:
            return False
        normalized = self.normalize_chat_request(text).lower()
        if not normalized:
            return False
        if self.parse_direct_command_from_chat(text):
            return False
        if self.extract_i_wish_body(text):
            return True
        conversational_prefixes = (
            "hi",
            "hello",
            "hey",
            "thanks",
            "thank you",
            "reply",
            "can you reply",
            "are you there",
            "hello?",
        )
        if normalized.startswith(conversational_prefixes):
            return False
        task_prefixes = (
            "please ",
            "send ",
            "transfer ",
            "copy ",
            "move ",
            "pull ",
            "push ",
            "download ",
            "upload ",
            "print ",
            "check ",
            "inspect ",
            "verify ",
            "look ",
            "find ",
            "open ",
            "run ",
            "install ",
            "restart ",
            "scan ",
            "update ",
            "tell ",
            "ask ",
            "get ",
            "bring ",
        )
        if normalized.startswith(task_prefixes):
            return True
        if " please " in f" {normalized} " and len(normalized.split()) >= 4:
            return True
        if normalized.endswith("?") and any(word in normalized for word in ("check", "find", "look", "send", "transfer", "print", "verify", "status", "update")):
            return True
        return False

    def create_tracked_task(self, assigned_to: str, details: str, source_message_id: str = "") -> dict:
        normalized = self.normalize_chat_request(details)
        title = normalized[:96] if len(normalized) <= 96 else normalized[:93] + "..."
        payload = {
            "id": make_id("task"),
            "title": title or "Untitled task",
            "details": details.strip(),
            "assigned_to": assigned_to,
            "requested_by": self.local_node_var.get().strip() or "home",
            "requested_by_name": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            "destination_path": "",
            "status": "pending",
            "created_at": now_iso(),
            "updated_at": now_iso(),
            "target": assigned_to,
            "source_message_id": source_message_id,
        }
        filename = f"{payload['created_at'].replace(':', '').replace('-', '')}_{payload['id']}.json"
        self.write_bridge_payload("tasks", filename, payload)
        return payload

    def create_command_record_from_message(self, message: dict, action_kind: str, command_text: str, summary: str) -> dict:
        local_node = self.local_node_var.get().strip() or "home"
        sender = str(message.get("sender") or "").strip()
        sender_name = str(message.get("sender_name") or sender or "Remote Sender").strip()
        command_id = make_id("cmd")
        approval_required = not (self.should_auto_execute_trusted_commands() and self.is_trusted_peer_sender(sender))
        payload = {
            "id": command_id,
            "kind": action_kind,
            "action": action_kind,
            "sender": sender,
            "sender_name": sender_name,
            "target": local_node,
            "project": "",
            "summary": summary,
            "command": command_text,
            "workdir": "",
            "approval_required": approval_required,
            "status": "pending_approval" if approval_required else "approved",
            "created_at": now_iso(),
            "updated_at": now_iso(),
            "approved_by": sender_name if not approval_required else "",
            "source_message_id": str(message.get("id") or ""),
            "result": {},
        }
        filename = f"{payload['created_at'].replace(':', '').replace('-', '')}_{command_id}.json"
        self.write_bridge_payload("commands", filename, payload)
        return payload

    def create_outgoing_command_request(
        self,
        target: str,
        action_kind: str,
        command_text: str,
        summary: str,
        *,
        project: str = "",
        workdir: str = "",
        source_message_id: str = "",
        approval_required: bool | None = None,
    ) -> dict:
        local_node = self.local_node_var.get().strip() or "home"
        sender_name = self.display_name_var.get().strip() or self.default_display_name(local_node)
        if approval_required is None:
            approval_required = False if target in {"home", "office"} else True
        command_id = make_id("cmd")
        payload = {
            "id": command_id,
            "kind": action_kind,
            "action": action_kind,
            "sender": local_node,
            "sender_name": sender_name,
            "target": target,
            "project": project,
            "summary": summary,
            "command": command_text,
            "workdir": workdir,
            "approval_required": bool(approval_required),
            "status": "pending_approval" if approval_required else "approved",
            "created_at": now_iso(),
            "updated_at": now_iso(),
            "approved_by": "" if approval_required else sender_name,
            "source_message_id": source_message_id,
            "result": {},
        }
        filename = f"{payload['created_at'].replace(':', '').replace('-', '')}_{command_id}.json"
        self.write_bridge_payload("commands", filename, payload)
        return payload

    def summarize_result_for_chat(self, record: dict) -> str:
        result = record.get("result") or {}
        action = str(result.get("action") or record.get("action") or record.get("kind") or "command")
        summary = str(record.get("summary") or action).strip() or action
        exit_code = result.get("exit_code")
        if action == "print_test_page":
            printed_file = str(result.get("printed_file") or "").strip()
            if exit_code == 0:
                if printed_file:
                    return f"I sent a test page to the printer path. Windows accepted the request and used {printed_file}. I can't physically confirm paper output from here."
                return "I sent a test page to the printer path. Windows accepted the request, but I can't physically confirm paper output from here."
            return f"I tried to print a test page, but it failed. {str(result.get('stderr') or 'No printer error text was returned.').strip()}"
        stdout = str(result.get("stdout") or "").strip()
        stderr = str(result.get("stderr") or "").strip()
        if exit_code == 0:
            if stdout:
                first = stdout.splitlines()[0][:260]
                return f"I ran '{summary}' successfully. First result line: {first}"
            return f"I ran '{summary}' successfully."
        if stderr:
            return f"I tried to run '{summary}', but it failed. Error: {stderr.splitlines()[0][:260]}"
        return f"I tried to run '{summary}', but it failed with exit code {exit_code}."

    def auto_create_commands_from_incoming_chat(self) -> None:
        if not self.should_auto_convert_trusted_chat_commands():
            return
        local_node = self.local_node_var.get().strip() or "home"
        state = self.load_trusted_chat_state()
        handled_ids = set(state.get("handled_ids", []))
        incoming = [
            item for item in self.latest_messages
            if item.get("kind") == "chat"
            and item.get("sender") not in ("", None, local_node)
            and item.get("target") == local_node
            and self.is_trusted_peer_sender(str(item.get("sender") or ""))
            and not bool(item.get("routed_action"))
        ]
        if not state.get("initialized"):
            state["initialized"] = True
            state["handled_ids"] = sorted({item.get("id") for item in incoming if item.get("id")})
            state["updated_at"] = now_iso()
            self.save_trusted_chat_state(state)
            return
        for item in incoming:
            message_id = str(item.get("id") or "")
            if not message_id or message_id in handled_ids:
                continue
            parsed = self.parse_action_from_chat(str(item.get("text") or ""))
            if parsed is not None:
                action_kind, command_text, summary = parsed
                record = self.create_command_record_from_message(item, action_kind, command_text, summary)
                if record["status"] == "approved":
                    self.send_system_message(
                        f"Trusted bridge command accepted from {item.get('sender_name') or item.get('sender')}: {summary}",
                        target=str(item.get("sender") or "all"),
                    )
                else:
                    self.send_system_message(
                        f"Bridge command from {item.get('sender_name') or item.get('sender')} is waiting for approval: {summary}",
                        target=str(item.get("sender") or "all"),
                    )
            handled_ids.add(message_id)
        state["initialized"] = True
        state["handled_ids"] = sorted(handled_ids)
        state["updated_at"] = now_iso()
        self.save_trusted_chat_state(state)

    def route_chat_message(self, message: dict) -> str:
        target = str(message.get("target") or "all").strip().lower()
        text = str(message.get("text") or "")
        if target not in {"home", "office"}:
            return "chat"
        parsed_action = self.parse_action_from_chat(text)
        if parsed_action:
            action_kind, command_text, summary = parsed_action
            message["routed_action"] = True
            self.create_outgoing_command_request(
                target,
                action_kind,
                command_text,
                summary,
                source_message_id=str(message.get("id") or ""),
                approval_required=False,
            )
            self.send_system_message(f"Routed actionable chat to {target}: {summary}", target=target)
            return "command"
        wish_body = self.extract_i_wish_body(text)
        if self.is_task_like_chat(text, target):
            details = wish_body or text
            task = self.create_tracked_task(target, details, str(message.get("id") or ""))
            if wish_body:
                self.send_system_message(f"Immediate action request created for {target}: {task['title']}")
                return "wish_task"
            self.send_system_message(f"Tracked task created for {target}: {task['title']}")
            return "task"
        return "chat"

    def send_chat_message(self) -> None:
        text = self.chat_entry.get("1.0", "end").strip()
        if not text:
            messagebox.showinfo("No Message", "Type a message first.")
            return
        message = {
            "id": make_id("msg"),
            "kind": "chat",
            "sender": self.local_node_var.get().strip() or "home",
            "sender_name": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            "target": self.chat_target_var.get().strip() or "all",
            "text": text,
            "created_at": now_iso(),
        }
        route_result = self.route_chat_message(message)
        filename = f"{message['created_at'].replace(':', '').replace('-', '')}_{message['id']}.json"
        self.write_bridge_payload("messages", filename, message)
        self.chat_entry.delete("1.0", "end")
        if route_result == "command":
            self.status_var.set("Chat message sent and routed into a direct command.")
        elif route_result == "wish_task":
            self.status_var.set("Immediate action request sent and tracked.")
        elif route_result == "task":
            self.status_var.set("Chat message sent and tracked task created.")
        else:
            self.status_var.set("Chat message sent.")
        self.refresh_all()

    def send_system_message(self, text: str, target: str = "all") -> None:
        message = {
            "id": make_id("msg"),
            "kind": "system",
            "sender": self.local_node_var.get().strip() or "home",
            "sender_name": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            "target": target,
            "text": text,
            "created_at": now_iso(),
        }
        filename = f"{message['created_at'].replace(':', '').replace('-', '')}_{message['id']}.json"
        self.write_bridge_payload("messages", filename, message)

    def send_command_request(self) -> None:
        command_text = self.command_text.get("1.0", "end").strip()
        target = self.command_target_var.get().strip()
        if not command_text:
            messagebox.showinfo("No Command", "Type a PowerShell command first.")
            return
        if not target:
            messagebox.showinfo("No Target", "Choose which node should receive the command.")
            return
        command_id = make_id("cmd")
        payload = {
            "id": command_id,
            "kind": "shell_command",
            "sender": self.local_node_var.get().strip() or "home",
            "sender_name": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            "target": target,
            "project": self.command_project_var.get().strip(),
            "summary": self.command_summary_var.get().strip() or command_text.splitlines()[0][:120],
            "command": command_text,
            "workdir": self.command_workdir_var.get().strip(),
            "approval_required": bool(self.command_requires_approval_var.get()),
            "status": "pending_approval" if self.command_requires_approval_var.get() else "approved",
            "created_at": now_iso(),
            "updated_at": now_iso(),
            "approved_by": "",
            "result": {},
        }
        filename = f"{payload['created_at'].replace(':', '').replace('-', '')}_{command_id}.json"
        self.write_bridge_payload("commands", filename, payload)
        self.command_text.delete("1.0", "end")
        self.command_summary_var.set("")
        self.send_system_message(f"{payload['sender_name']} requested a command on {target}: {payload['summary']}")
        self.status_var.set("Command request sent.")
        self.refresh_all()

    def command_records(self) -> list[dict]:
        return sorted(self.read_json_files(self.commands_dir()), key=lambda item: item.get("created_at", ""))

    def message_records(self) -> list[dict]:
        return sorted(self.read_json_files(self.messages_dir()), key=lambda item: item.get("created_at", ""))

    def publish_focus_update(self) -> None:
        self.save_setup()
        self.send_system_message(
            f"{self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get())} "
            f"updated local focus: {self.focus_project_var.get().strip() or 'No project named'}"
        )
        self.refresh_all()

    def publish_local_state(self) -> None:
        self.save_config()
        projects = self.scan_projects()
        payload = {
            "node": self.local_node_var.get().strip() or "home",
            "display_name": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            "focus_project": self.focus_project_var.get().strip(),
            "focus_note": self.focus_note_text.get("1.0", "end").strip(),
            "project_roots": self.config.project_roots,
            "projects": projects,
            "updated_at": now_iso(),
        }
        filename = f"{payload['node']}.json"
        self.write_bridge_payload("state", filename, payload)

    def scan_projects(self) -> list[dict]:
        records: list[dict] = []
        for root_text in self.config.project_roots:
            root_path = Path(root_text)
            if not root_path.exists() or not root_path.is_dir():
                continue
            for child in sorted(root_path.iterdir(), key=lambda item: item.name.lower()):
                if child.name.startswith(".") or not child.is_dir():
                    continue
                records.append(
                    {
                        "name": child.name,
                        "path": str(child),
                        "root": str(root_path),
                        "modified_at": datetime.fromtimestamp(child.stat().st_mtime).isoformat(timespec="seconds"),
                    }
                )
        return records

    def load_node_states(self) -> dict[str, dict]:
        states: dict[str, dict] = {}
        for record in self.read_json_files(self.state_dir()):
            node = record.get("node")
            if node:
                states[node] = record
        return states

    def refresh_messages_view(self) -> None:
        self.latest_messages = self.message_records()
        lines: list[str] = []
        local_node = self.local_node_var.get().strip() or "home"
        opposite = self.opposite_node(local_node)
        for item in self.latest_messages[-200:]:
            target = item.get("target", "all")
            if target not in ("all", local_node, opposite):
                continue
            stamp = item.get("created_at", "")
            sender = item.get("sender_name") or item.get("sender") or "unknown"
            kind = item.get("kind", "chat")
            text = item.get("text", "")
            prefix = f"[{stamp}] {sender}"
            if kind != "chat":
                prefix += f" [{kind}]"
            lines.append(prefix)
            lines.append(text)
            lines.append("")
        self.set_text(self.chat_text, "\n".join(lines).strip())
        self.chat_text.see("end")

    def refresh_commands_view(self) -> None:
        self.latest_commands = self.command_records()
        for item in self.approvals_tree.get_children():
            self.approvals_tree.delete(item)
        local_node = self.local_node_var.get().strip() or "home"
        for record in self.latest_commands:
            is_incoming = record.get("target") == local_node
            is_outgoing = record.get("sender") == local_node
            if not (is_incoming or is_outgoing):
                continue
            direction = "Incoming" if is_incoming else "Outgoing"
            self.approvals_tree.insert(
                "",
                "end",
                iid=record["id"],
                values=(
                    record.get("created_at", ""),
                    direction,
                    record.get("sender_name") or record.get("sender", ""),
                    record.get("summary", ""),
                    record.get("status", ""),
                ),
            )
        self.show_selected_command_details()

    def show_selected_command_details(self) -> None:
        selection = self.approvals_tree.selection()
        if not selection:
            self.set_text(self.approval_detail_text, "Select a command request to review it here.")
            return
        command_id = selection[0]
        record = next((item for item in self.latest_commands if item.get("id") == command_id), None)
        if not record:
            self.set_text(self.approval_detail_text, "Command details unavailable.")
            return
        detail = [
            f"Summary: {record.get('summary', '')}",
            f"Action: {record.get('action') or record.get('kind', '')}",
            f"Sender: {record.get('sender_name') or record.get('sender', '')}",
            f"Target: {record.get('target', '')}",
            f"Project: {record.get('project', '')}",
            f"Status: {record.get('status', '')}",
            f"Approval required: {record.get('approval_required', False)}",
            f"Working dir: {record.get('workdir', '') or '(none)'}",
            f"Source message id: {record.get('source_message_id', '')}",
            "",
            "Command:",
            record.get("command", ""),
        ]
        result = record.get("result") or {}
        if result:
            detail.extend(
                [
                    "",
                    f"Exit code: {result.get('exit_code', '')}",
                    "Stdout:",
                    result.get("stdout", ""),
                    "",
                    "Stderr:",
                    result.get("stderr", ""),
                ]
            )
        self.set_text(self.approval_detail_text, "\n".join(detail).strip())

    def approve_selected_command(self) -> None:
        self.update_selected_command_status("approved")

    def deny_selected_command(self) -> None:
        self.update_selected_command_status("denied")

    def update_selected_command_status(self, new_status: str) -> None:
        selection = self.approvals_tree.selection()
        if not selection:
            messagebox.showinfo("No Selection", "Select a command request first.")
            return
        command_id = selection[0]
        record = next((item for item in self.latest_commands if item.get("id") == command_id), None)
        if not record:
            return
        record["status"] = new_status
        record["updated_at"] = now_iso()
        record["approved_by"] = self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get())
        self.write_json(Path(record["_path"]), {key: value for key, value in record.items() if not key.startswith("_")})
        action = "approved" if new_status == "approved" else "denied"
        message = f"{record['approved_by']} {action} command: {record.get('summary', '')}"
        self.send_system_message(message)
        self.status_var.set(message)
        self.refresh_all()

    def auto_approve_trusted_commands(self) -> None:
        if not self.should_auto_execute_trusted_commands():
            return
        local_node = self.local_node_var.get().strip() or "home"
        approved_any = False
        for record in self.latest_commands:
            if record.get("target") != local_node:
                continue
            if record.get("status") != "pending_approval":
                continue
            if not self.is_trusted_peer_sender(str(record.get("sender") or "")):
                continue
            record["status"] = "approved"
            record["updated_at"] = now_iso()
            record["approved_by"] = self.display_name_var.get().strip() or self.default_display_name(local_node)
            self.write_json(Path(record["_path"]), {key: value for key, value in record.items() if not key.startswith("_")})
            approved_any = True
            self.log(f"Auto-approved trusted command: {record.get('summary', '')}")
        if approved_any:
            self.latest_commands = self.command_records()

    def auto_run_approved_commands(self) -> None:
        local_node = self.local_node_var.get().strip() or "home"
        for record in self.latest_commands:
            if record.get("target") != local_node:
                continue
            if record.get("status") != "approved":
                continue
            command_id = record.get("id")
            if not command_id or command_id in self.active_commands:
                continue
            self.active_commands.add(command_id)
            thread = threading.Thread(target=self.execute_command, args=(record,), daemon=True)
            thread.start()

    def print_test_page(self) -> dict:
        temp_dir = Path(tempfile.gettempdir()) / "codex_bridge"
        temp_dir.mkdir(parents=True, exist_ok=True)
        file_path = temp_dir / f"codex_bridge_test_page_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        body = (
            "Codex Bridge Test Page\r\n"
            f"Printed at: {now_iso()}\r\n"
            f"Node: {self.local_node_var.get().strip() or 'home'}\r\n"
            f"Display: {self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get())}\r\n"
        )
        file_path.write_text(body, encoding="utf-8")
        command_text = f'Start-Process -FilePath notepad.exe -ArgumentList "/p","{file_path}" -Wait'
        completed = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-Command",
                command_text,
            ],
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=CREATE_NO_WINDOW,
        )
        return {
            "exit_code": completed.returncode,
            "stdout": completed.stdout[-10000:],
            "stderr": completed.stderr[-10000:],
            "printed_file": str(file_path),
        }

    def execute_command(self, record: dict) -> None:
        path = Path(record["_path"])
        try:
            record["status"] = "running"
            record["updated_at"] = now_iso()
            self.write_json(path, {key: value for key, value in record.items() if not key.startswith("_")})
            self.root.after(0, lambda: self.status_var.set(f"Running command: {record.get('summary', '')}"))
            action = str(record.get("action") or record.get("kind") or "shell_command").strip()
            if action == "print_test_page":
                result = self.print_test_page()
            else:
                completed = subprocess.run(
                    [
                        "powershell",
                        "-NoProfile",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-Command",
                        record.get("command", ""),
                    ],
                    capture_output=True,
                    text=True,
                    cwd=record.get("workdir") or None,
                    timeout=600,
                    creationflags=CREATE_NO_WINDOW,
                )
                result = {
                    "exit_code": completed.returncode,
                    "stdout": completed.stdout[-10000:],
                    "stderr": completed.stderr[-10000:],
                }
            exit_code = result.get("exit_code", 1)
            record["status"] = "completed" if exit_code == 0 else "failed"
            record["updated_at"] = now_iso()
            record["result"] = {
                **result,
                "ran_at": now_iso(),
                "action": action,
                "executor": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            }
            self.write_json(path, {key: value for key, value in record.items() if not key.startswith("_")})
            summary = (
                f"{record['result']['executor']} completed command '{record.get('summary', '')}' "
                f"with exit code {exit_code}."
            )
            self.send_system_message(summary)
            reply_target = str(record.get("sender") or "").strip()
            if reply_target:
                reply_id = make_id("msg")
                reply_created_at = now_iso()
                self.write_bridge_payload(
                    "messages",
                    f"{reply_created_at.replace(':', '').replace('-', '')}_{reply_id}.json",
                    {
                        "id": reply_id,
                        "kind": "chat",
                        "sender": self.local_node_var.get().strip() or "home",
                        "sender_name": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
                        "target": reply_target,
                        "text": self.summarize_result_for_chat(record),
                        "reply_to": record.get("source_message_id", ""),
                        "created_at": reply_created_at,
                    },
                )
            self.log(summary)
            if str(record["result"].get("stdout") or "").strip():
                self.log("Stdout:\n" + str(record["result"]["stdout"])[-4000:])
            if str(record["result"].get("stderr") or "").strip():
                self.log("Stderr:\n" + str(record["result"]["stderr"])[-4000:])
        except subprocess.TimeoutExpired:
            record["status"] = "failed"
            record["updated_at"] = now_iso()
            record["result"] = {
                "exit_code": "timeout",
                "stdout": "",
                "stderr": "Command timed out after 600 seconds.",
                "ran_at": now_iso(),
                "executor": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            }
            self.write_json(path, {key: value for key, value in record.items() if not key.startswith("_")})
            self.send_system_message(f"Command timed out: {record.get('summary', '')}")
        except Exception as exc:
            record["status"] = "failed"
            record["updated_at"] = now_iso()
            record["result"] = {
                "exit_code": "error",
                "stdout": "",
                "stderr": str(exc),
                "ran_at": now_iso(),
                "executor": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            }
            self.write_json(path, {key: value for key, value in record.items() if not key.startswith("_")})
            self.send_system_message(f"Command failed to start: {record.get('summary', '')}")
        finally:
            command_id = record.get("id")
            if command_id in self.active_commands:
                self.active_commands.remove(command_id)
            self.root.after(0, self.refresh_all)

    def refresh_projects_view(self) -> None:
        self.latest_nodes = self.load_node_states()
        aggregate: dict[str, dict] = {}
        for node, state in self.latest_nodes.items():
            node_name = state.get("display_name") or node
            if state.get("focus_project"):
                name = state["focus_project"]
                info = aggregate.setdefault(name, {"nodes": set(), "last_update": state.get("updated_at", ""), "last_note": ""})
                info["nodes"].add(node_name)
                info["last_update"] = max(info["last_update"], state.get("updated_at", ""))
                if state.get("focus_note"):
                    info["last_note"] = state["focus_note"]
            for project in state.get("projects", []):
                name = project.get("name")
                if not name:
                    continue
                info = aggregate.setdefault(name, {"nodes": set(), "last_update": "", "last_note": ""})
                info["nodes"].add(node_name)
                info["last_update"] = max(info["last_update"], project.get("modified_at", ""))
        self.latest_projects = aggregate

        for item in self.projects_tree.get_children():
            self.projects_tree.delete(item)
        for name in sorted(aggregate):
            info = aggregate[name]
            self.projects_tree.insert(
                "",
                "end",
                values=(
                    name,
                    ", ".join(sorted(info["nodes"])),
                    info.get("last_update", ""),
                    info.get("last_note", ""),
                ),
            )

    def set_text(self, widget: tk.Text, content: str) -> None:
        widget.config(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", content)
        widget.config(state="disabled")

    def log(self, text: str) -> None:
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"[{now_iso()}] {text}\n\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def refresh_all(self) -> None:
        self.ensure_bridge_dirs()
        self.publish_local_state()
        self.refresh_messages_view()
        self.refresh_commands_view()
        self.refresh_projects_view()
        try:
            self.refresh_explorer()
        except Exception as exc:
            self.log(f"Explorer refresh failed: {exc}")
        self.auto_create_commands_from_incoming_chat()
        self.refresh_commands_view()
        self.auto_approve_trusted_commands()
        self.auto_run_approved_commands()
        self.auto_reply_to_messages()

    def poll_bridge(self) -> None:
        try:
            self.refresh_all()
        except Exception as exc:
            self.log(f"Refresh failed: {exc}")
        finally:
            self.root.after(POLL_INTERVAL_MS, self.poll_bridge)


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    CodexBridgeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
