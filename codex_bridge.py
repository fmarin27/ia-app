from __future__ import annotations

import json
import os
import site
import subprocess
import sys
import threading
import urllib.error
import urllib.request
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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
POLL_INTERVAL_MS = 3_000


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
        self.auto_reply_enabled_var = tk.BooleanVar(value=self.config.auto_reply_enabled)
        self.auto_reply_model_var = tk.StringVar(value=self.config.auto_reply_model)
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
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
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

    def state_dir(self) -> Path:
        return self.bridge_root() / "state"

    def ensure_bridge_dirs(self) -> None:
        for path in (self.bridge_root(), self.messages_dir(), self.commands_dir(), self.state_dir()):
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
        left.add(chat_tab, text="Chat")
        left.add(commands_tab, text="Commands")
        left.add(approvals_tab, text="Approvals")

        projects_tab = ttk.Frame(right, padding=10)
        setup_tab = ttk.Frame(right, padding=10)
        log_tab = ttk.Frame(right, padding=10)
        right.add(projects_tab, text="Projects")
        right.add(setup_tab, text="Setup")
        right.add(log_tab, text="Log")

        self.build_chat_tab(chat_tab)
        self.build_commands_tab(commands_tab)
        self.build_approvals_tab(approvals_tab)
        self.build_projects_tab(projects_tab)
        self.build_setup_tab(setup_tab)
        self.build_log_tab(log_tab)

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

    def build_approvals_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        columns = ("created", "sender", "summary", "status")
        self.approvals_tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="browse")
        for column, heading, width in (
            ("created", "Created", 160),
            ("sender", "Sender", 120),
            ("summary", "Summary", 440),
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

        roots_frame = ttk.LabelFrame(parent, text="Local Project Roots To Publish", padding=8)
        roots_frame.grid(row=3, column=0, sticky="nsew", pady=(10, 0))
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
            path = self.messages_dir() / f"{response['created_at'].replace(':', '').replace('-', '')}_{response['id']}.json"
            self.write_json(path, response)
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
        path = self.messages_dir() / f"{message['created_at'].replace(':', '').replace('-', '')}_{message['id']}.json"
        self.write_json(path, message)
        self.chat_entry.delete("1.0", "end")
        self.status_var.set("Chat message sent.")
        self.refresh_all()

    def send_system_message(self, text: str) -> None:
        message = {
            "id": make_id("msg"),
            "kind": "system",
            "sender": self.local_node_var.get().strip() or "home",
            "sender_name": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            "target": "all",
            "text": text,
            "created_at": now_iso(),
        }
        path = self.messages_dir() / f"{message['created_at'].replace(':', '').replace('-', '')}_{message['id']}.json"
        self.write_json(path, message)

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
        path = self.commands_dir() / f"{payload['created_at'].replace(':', '').replace('-', '')}_{command_id}.json"
        self.write_json(path, payload)
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
        path = self.state_dir() / f"{payload['node']}.json"
        self.write_json(path, payload)

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

    def refresh_commands_view(self) -> None:
        self.latest_commands = self.command_records()
        for item in self.approvals_tree.get_children():
            self.approvals_tree.delete(item)
        local_node = self.local_node_var.get().strip() or "home"
        for record in self.latest_commands:
            if record.get("target") != local_node:
                continue
            self.approvals_tree.insert(
                "",
                "end",
                iid=record["id"],
                values=(
                    record.get("created_at", ""),
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
            f"Sender: {record.get('sender_name') or record.get('sender', '')}",
            f"Project: {record.get('project', '')}",
            f"Status: {record.get('status', '')}",
            f"Approval required: {record.get('approval_required', False)}",
            f"Working dir: {record.get('workdir', '') or '(none)'}",
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

    def execute_command(self, record: dict) -> None:
        path = Path(record["_path"])
        try:
            record["status"] = "running"
            record["updated_at"] = now_iso()
            self.write_json(path, {key: value for key, value in record.items() if not key.startswith("_")})
            self.root.after(0, lambda: self.status_var.set(f"Running command: {record.get('summary', '')}"))

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
            )
            record["status"] = "completed" if completed.returncode == 0 else "failed"
            record["updated_at"] = now_iso()
            record["result"] = {
                "exit_code": completed.returncode,
                "stdout": completed.stdout[-10000:],
                "stderr": completed.stderr[-10000:],
                "ran_at": now_iso(),
                "executor": self.display_name_var.get().strip() or self.default_display_name(self.local_node_var.get()),
            }
            self.write_json(path, {key: value for key, value in record.items() if not key.startswith("_")})
            summary = (
                f"{record['result']['executor']} completed command '{record.get('summary', '')}' "
                f"with exit code {completed.returncode}."
            )
            self.send_system_message(summary)
            self.log(summary)
            if completed.stdout.strip():
                self.log("Stdout:\n" + completed.stdout[-4000:])
            if completed.stderr.strip():
                self.log("Stderr:\n" + completed.stderr[-4000:])
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
