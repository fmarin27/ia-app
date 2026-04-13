from __future__ import annotations

import json
import site
import shutil
import subprocess
import sys
import tempfile
import uuid
from dataclasses import asdict, dataclass, field, fields
from datetime import datetime
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
USER_SITE = site.getusersitepackages()
if USER_SITE and USER_SITE not in sys.path:
    sys.path.append(USER_SITE)
QT_VENDOR = APP_DIR / "qt_vendor"
if QT_VENDOR.exists() and str(QT_VENDOR) not in sys.path:
    sys.path.insert(0, str(QT_VENDOR))

from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QAction, QColor, QFont, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QSplitter,
    QStatusBar,
    QSystemTrayIcon,
    QInputDialog,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

CONFIG_FILE = APP_DIR / "codex_bridge_config.json"
DEFAULT_BRIDGE_ROOT = Path.home() / "Syncthing" / "codex-bridge"
POLL_INTERVAL_MS = 3_000
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


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


def format_timestamp(value: str) -> str:
    if not value:
        return "Never"
    try:
        stamp = datetime.fromisoformat(value)
    except Exception:
        return value
    return stamp.strftime("%b %d, %I:%M %p")


def format_age(value: str) -> str:
    if not value:
        return "No heartbeat yet"
    try:
        stamp = datetime.fromisoformat(value)
    except Exception:
        return "Unknown"
    delta = datetime.now() - stamp
    seconds = int(delta.total_seconds())
    if seconds < 60:
        return "Just now"
    if seconds < 3600:
        return f"{seconds // 60}m ago"
    if seconds < 86400:
        return f"{seconds // 3600}h ago"
    return f"{seconds // 86400}d ago"


class StatCard(QFrame):
    def __init__(self, title: str) -> None:
        super().__init__()
        self.setObjectName("statCard")
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(2)

        self.title_label = QLabel(title)
        self.title_label.setObjectName("cardTitle")
        layout.addWidget(self.title_label)

        self.value_label = QLabel("--")
        self.value_label.setObjectName("cardValue")
        layout.addWidget(self.value_label)

        self.meta_label = QLabel("")
        self.meta_label.setObjectName("cardMeta")
        self.meta_label.setWordWrap(True)
        layout.addWidget(self.meta_label)

    def set_value(self, value: str, meta: str = "") -> None:
        self.value_label.setText(value)
        self.meta_label.setText(meta)


class BridgeStore:
    def __init__(self, config: BridgeConfig):
        self.config = config
        self.active_commands: set[str] = set()

    @property
    def bridge_root(self) -> Path:
        return Path(self.config.bridge_root)

    @property
    def messages_dir(self) -> Path:
        return self.bridge_root / "messages"

    @property
    def commands_dir(self) -> Path:
        return self.bridge_root / "commands"

    @property
    def state_dir(self) -> Path:
        return self.bridge_root / "state"

    def ensure_dirs(self) -> None:
        for path in (self.bridge_root, self.messages_dir, self.commands_dir, self.state_dir):
            path.mkdir(parents=True, exist_ok=True)

    def write_json(self, path: Path, payload: dict) -> None:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp = path.with_suffix(path.suffix + ".tmp")
        temp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        temp.replace(path)

    def read_json_files(self, directory: Path) -> list[dict]:
        items: list[dict] = []
        if not directory.exists():
            return items
        for path in sorted(directory.glob("*.json")):
            try:
                payload = json.loads(path.read_text(encoding="utf-8"))
                payload["_path"] = str(path)
                items.append(payload)
            except Exception:
                continue
        return items

    def publish_state(self) -> None:
        payload = {
            "node": self.config.local_node,
            "display_name": self.config.display_name,
            "focus_project": self.config.focus_project,
            "focus_note": self.config.focus_note,
            "project_roots": self.config.project_roots,
            "projects": self.scan_projects(),
            "updated_at": now_iso(),
        }
        self.write_json(self.state_dir / f"{self.config.local_node}.json", payload)

    def scan_projects(self) -> list[dict]:
        rows: list[dict] = []
        for root_text in self.config.project_roots:
            root = Path(root_text)
            if not root.exists() or not root.is_dir():
                continue
            for child in sorted(root.iterdir(), key=lambda item: item.name.lower()):
                if child.name.startswith(".") or not child.is_dir():
                    continue
                rows.append(
                    {
                        "name": child.name,
                        "path": str(child),
                        "root": str(root),
                        "modified_at": datetime.fromtimestamp(child.stat().st_mtime).isoformat(timespec="seconds"),
                    }
                )
        return rows

    def build_filename(self, created_at: str, payload_id: str) -> str:
        return f"{created_at.replace(':', '').replace('-', '')}_{payload_id}.json"

    def build_message_payload(self, kind: str, text: str, target: str = "all", extra: dict | None = None) -> dict:
        payload = {
            "id": make_id("msg"),
            "kind": kind,
            "sender": self.config.local_node,
            "sender_name": self.config.display_name,
            "target": target,
            "text": text,
            "created_at": now_iso(),
        }
        if extra:
            payload.update(extra)
        return payload

    def send_message(self, kind: str, text: str, target: str = "all", extra: dict | None = None) -> dict:
        payload = self.build_message_payload(kind, text, target, extra)
        filename = self.build_filename(payload["created_at"], payload["id"])
        self.write_json(self.messages_dir / filename, payload)
        return payload

    def build_command_payload(self, target: str, summary: str, command: str, workdir: str, project: str, approval_required: bool, action: str = "shell_command", source_message_id: str = "") -> dict:
        payload = {
            "id": make_id("cmd"),
            "kind": action,
            "action": action,
            "sender": self.config.local_node,
            "sender_name": self.config.display_name,
            "target": target,
            "project": project,
            "summary": summary or command.splitlines()[0][:120],
            "command": command,
            "workdir": workdir,
            "approval_required": approval_required,
            "status": "pending_approval" if approval_required else "approved",
            "created_at": now_iso(),
            "updated_at": now_iso(),
            "approved_by": "",
            "source_message_id": source_message_id,
            "result": {},
        }
        return payload

    def send_command(self, target: str, summary: str, command: str, workdir: str, project: str, approval_required: bool, action: str = "shell_command", source_message_id: str = "") -> dict:
        payload = self.build_command_payload(target, summary, command, workdir, project, approval_required, action, source_message_id)
        filename = self.build_filename(payload["created_at"], payload["id"])
        self.write_json(self.commands_dir / filename, payload)
        return payload

    def messages(self) -> list[dict]:
        return sorted(self.read_json_files(self.messages_dir), key=lambda item: item.get("created_at", ""))

    def commands(self) -> list[dict]:
        return sorted(self.read_json_files(self.commands_dir), key=lambda item: item.get("created_at", ""))

    def node_states(self) -> dict[str, dict]:
        result: dict[str, dict] = {}
        for item in self.read_json_files(self.state_dir):
            node = item.get("node")
            if node:
                result[node] = item
        return result

    def update_command(self, record: dict) -> None:
        path = Path(record["_path"])
        self.write_json(path, {k: v for k, v in record.items() if not k.startswith("_")})


class BridgeWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.config = self.load_config()
        self.store = BridgeStore(self.config)
        self.store.ensure_dirs()
        self.latest_messages: list[dict] = []
        self.latest_commands: list[dict] = []
        self.latest_states: dict[str, dict] = {}
        self.visible_commands: list[dict] = []
        self.local_entries: list[dict] = []
        self.remote_entries: list[dict] = []
        self.remote_refresh_inflight = False
        self.remote_refresh_generation = 0
        self.setWindowTitle("Codex Bridge Control Center")
        self.resize(1540, 980)
        self.setStyleSheet(self.stylesheet())
        self.build_ui()
        self.build_tray()
        self.refresh_all()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh_all)
        self.timer.start(POLL_INTERVAL_MS)

    def load_config(self) -> BridgeConfig:
        if not CONFIG_FILE.exists():
            return BridgeConfig()
        try:
            raw = json.loads(CONFIG_FILE.read_text(encoding="utf-8-sig"))
            allowed = {item.name for item in fields(BridgeConfig)}
            filtered = {key: value for key, value in raw.items() if key in allowed}
            return BridgeConfig(**filtered)
        except Exception:
            return BridgeConfig()

    def save_config(self) -> None:
        self.config.local_node = self.node_combo.currentText() or "home"
        self.config.display_name = self.display_name_edit.text().strip() or ("Home Codex" if self.config.local_node == "home" else "Office Codex")
        self.config.bridge_root = self.bridge_root_edit.text().strip() or str(DEFAULT_BRIDGE_ROOT)
        self.config.focus_project = self.focus_project_edit.text().strip()
        self.config.focus_note = self.focus_note_edit.toPlainText().strip()
        self.config.project_roots = [self.project_roots_list.item(i).text() for i in range(self.project_roots_list.count())]
        self.config.auto_reply_enabled = self.auto_reply_checkbox.isChecked()
        self.config.auto_reply_model = self.auto_reply_model_edit.text().strip() or "gpt-4.1-mini"
        self.config.home_host = self.home_host_edit.text().strip()
        self.config.home_user = self.home_user_edit.text().strip() or "ferna"
        self.config.home_key = self.home_key_edit.text().strip()
        self.config.office_host = self.office_host_edit.text().strip()
        self.config.office_user = self.office_user_edit.text().strip() or "ferna"
        self.config.office_key = self.office_key_edit.text().strip()
        self.config.auto_execute_trusted_commands = self.auto_execute_checkbox.isChecked()
        self.config.auto_convert_trusted_chat_commands = self.auto_convert_checkbox.isChecked()
        CONFIG_FILE.write_text(json.dumps(asdict(self.config), indent=2), encoding="utf-8")
        self.store.config = self.config
        self.store.ensure_dirs()

    def stylesheet(self) -> str:
        return """
        QMainWindow { background: #07111f; }
        QWidget { color: #ecf3ff; font-size: 12px; font-family: "Segoe UI Variable Text", "Segoe UI", sans-serif; }
        QFrame#card, QFrame#heroCard, QFrame#statCard {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #102038, stop:1 #0d1729);
            border: 1px solid #223754;
            border-radius: 18px;
        }
        QLabel#eyebrow { color: #8fb6ff; font-size: 10px; font-weight: 700; }
        QLabel#heroTitle { color: #f6fbff; font-size: 18px; font-weight: 700; }
        QLabel#heroSub { color: #9db4d0; font-size: 12px; }
        QLabel#cardTitle { color: #89a7d0; font-size: 11px; font-weight: 700; }
        QLabel#cardValue { color: #f5fbff; font-size: 15px; font-weight: 700; }
        QLabel#cardMeta { color: #98acc5; font-size: 12px; }
        QLineEdit, QPlainTextEdit, QComboBox, QListWidget, QTableWidget {
            background: #0b1628; border: 1px solid #26415f; border-radius: 12px; padding: 8px;
        }
        QTableWidget {
            gridline-color: #17314d;
            selection-background-color: #89c2ff;
            selection-color: #07111f;
        }
        QTableWidget::item:selected {
            background: #89c2ff;
            color: #07111f;
        }
        QTableCornerButton::section {
            background: #12263f;
            border: none;
        }
        QPushButton {
            background: #8dc2ff; color: #08111e; border: none; border-radius: 12px; padding: 8px 12px; font-weight: 700;
        }
        QPushButton:hover { background: #a9d0ff; }
        QTabWidget::pane { border: none; }
        QTabBar::tab { background: #0f1b2d; color: #91aac8; padding: 9px 14px; margin-right: 6px; border-top-left-radius: 12px; border-top-right-radius: 12px; }
        QTabBar::tab:selected { background: #19314e; color: #f2f8ff; }
        QHeaderView::section { background: #12263f; color: #eef4ff; padding: 8px; border: none; font-weight: 700; }
        QGroupBox {
            border: 1px solid #223754;
            border-radius: 16px;
            margin-top: 10px;
            padding-top: 14px;
            background: #0d1729;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 14px;
            padding: 0 6px;
            color: #b7cff0;
            font-weight: 700;
        }
        """

    def build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(8)

        header = QFrame(objectName="heroCard")
        header_layout = QGridLayout(header)
        header_layout.setContentsMargins(16, 12, 16, 12)
        header_layout.setHorizontalSpacing(12)
        header_layout.setVerticalSpacing(6)
        eyebrow = QLabel("Bridge Control Center")
        eyebrow.setObjectName("eyebrow")
        header_layout.addWidget(eyebrow, 0, 0, 1, 4)
        title = QLabel("Home + Office command bridge")
        title.setObjectName("heroTitle")
        header_layout.addWidget(title, 1, 0, 1, 4)
        subtitle = QLabel("Track project work, route commands, and keep both PCs coordinated from one cleaner control surface.")
        subtitle.setWordWrap(True)
        subtitle.setObjectName("heroSub")
        header_layout.addWidget(subtitle, 2, 0, 1, 4)

        header_layout.addWidget(QLabel("Node"), 0, 4)
        self.node_combo = QComboBox()
        self.node_combo.addItems(["home", "office"])
        self.node_combo.setCurrentText(self.config.local_node)
        header_layout.addWidget(self.node_combo, 0, 5)
        header_layout.addWidget(QLabel("Display Name"), 1, 4)
        self.display_name_edit = QLineEdit(self.config.display_name)
        header_layout.addWidget(self.display_name_edit, 1, 5)
        header_layout.addWidget(QLabel("Bridge Folder"), 2, 4)
        self.bridge_root_edit = QLineEdit(self.config.bridge_root)
        header_layout.addWidget(self.bridge_root_edit, 2, 5)
        choose_btn = QPushButton("Choose")
        choose_btn.setProperty("class", "secondary")
        choose_btn.clicked.connect(self.choose_bridge_root)
        header_layout.addWidget(choose_btn, 2, 6)
        hide_btn = QPushButton("Hide To Tray")
        hide_btn.setProperty("class", "secondary")
        hide_btn.clicked.connect(self.hide)
        header_layout.addWidget(hide_btn, 0, 6)
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_and_refresh)
        header_layout.addWidget(save_btn, 1, 6)
        layout.addWidget(header)

        stats_row = QHBoxLayout()
        stats_row.setSpacing(8)
        self.local_card = StatCard("This PC")
        self.remote_card = StatCard("Other PC")
        self.queue_card = StatCard("Command Queue")
        self.projects_card = StatCard("Project Radar")
        stats_row.addWidget(self.local_card)
        stats_row.addWidget(self.remote_card)
        stats_row.addWidget(self.queue_card)
        stats_row.addWidget(self.projects_card)
        layout.addLayout(stats_row)

        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter, 1)

        left_tabs = QTabWidget()
        right_tabs = QTabWidget()
        splitter.addWidget(left_tabs)
        splitter.addWidget(right_tabs)
        splitter.setSizes([1080, 420])

        self.explorer_target_combo = QComboBox()
        self.explorer_target_combo.addItems(["home", "office"])
        self.explorer_target_combo.setCurrentText("office" if self.config.local_node == "home" else "home")
        self.local_path_edit = QLineEdit(str(Path.home()))
        self.remote_path_edit = QLineEdit(r"C:\Users\ferna")

        left_tabs.addTab(self.build_chat_tab(), "Conversations")
        left_tabs.addTab(self.build_commands_tab(), "Run on Other PC")
        left_tabs.addTab(self.build_approvals_tab(), "Command Queue")
        left_tabs.addTab(self.build_explorer_tab(), "Explorer")
        right_tabs.addTab(self.build_projects_tab(), "Project Radar")
        right_tabs.addTab(self.build_setup_tab(), "Control Center")
        right_tabs.addTab(self.build_log_tab(), "Activity")

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Bridge ready.")

    def build_chat_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        helper = QLabel("Talk normally here. Actionable chat should evolve into verified bridge commands instead of staying passive chatter.")
        helper.setWordWrap(True)
        helper.setObjectName("cardMeta")
        layout.addWidget(helper)
        self.chat_log = QPlainTextEdit()
        self.chat_log.setReadOnly(True)
        layout.addWidget(self.chat_log)
        row = QHBoxLayout()
        row.addWidget(QLabel("Send To"))
        self.chat_target_combo = QComboBox()
        self.chat_target_combo.addItems(["all", "home", "office"])
        row.addWidget(self.chat_target_combo)
        row.addStretch(1)
        send_btn = QPushButton("Send")
        send_btn.clicked.connect(self.send_chat)
        row.addWidget(send_btn)
        layout.addLayout(row)
        self.chat_input = QPlainTextEdit()
        self.chat_input.setPlaceholderText("Examples: 'office print a test page', 'run: Get-Process', 'what apps are running over there?'")
        layout.addWidget(self.chat_input)
        return tab

    def build_commands_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        helper = QLabel("Use this when you want a structured action on the other PC. This is the cleaner path for work that should actually execute.")
        helper.setWordWrap(True)
        helper.setObjectName("cardMeta")
        layout.addWidget(helper)
        form = QFormLayout()
        self.command_target_combo = QComboBox()
        self.command_target_combo.addItems(["home", "office"])
        self.command_target_combo.setCurrentText("office" if self.config.local_node == "home" else "home")
        form.addRow("Target PC", self.command_target_combo)
        self.command_project_edit = QLineEdit()
        self.command_project_edit.setPlaceholderText("Optional project label")
        form.addRow("Project", self.command_project_edit)
        workdir_row = QHBoxLayout()
        self.command_workdir_edit = QLineEdit()
        self.command_workdir_edit.setPlaceholderText("Optional working directory")
        workdir_row.addWidget(self.command_workdir_edit)
        workdir_btn = QPushButton("Choose")
        workdir_btn.clicked.connect(self.choose_command_workdir)
        workdir_row.addWidget(workdir_btn)
        wrap = QWidget()
        wrap.setLayout(workdir_row)
        form.addRow("Working Dir", wrap)
        self.command_summary_edit = QLineEdit()
        self.command_summary_edit.setPlaceholderText("Human-friendly summary")
        form.addRow("Summary", self.command_summary_edit)
        self.command_requires_approval = QCheckBox("Require approval before running")
        self.command_requires_approval.setChecked(True)
        form.addRow("", self.command_requires_approval)
        self.command_input = QPlainTextEdit()
        self.command_input.setPlaceholderText("PowerShell command to run on the other PC.")
        form.addRow("Command", self.command_input)
        layout.addLayout(form)
        send_btn = QPushButton("Send Command Request")
        send_btn.clicked.connect(self.send_command)
        layout.addWidget(send_btn, alignment=Qt.AlignRight)
        return tab

    def build_approvals_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        helper = QLabel("One place to see what is waiting, running, finished, or failed across both sides.")
        helper.setWordWrap(True)
        helper.setObjectName("cardMeta")
        layout.addWidget(helper)
        self.approvals_table = QTableWidget(0, 5)
        self.approvals_table.setHorizontalHeaderLabels(["Created", "Direction", "Other Side", "Summary", "Status"])
        self.approvals_table.itemSelectionChanged.connect(self.populate_command_detail)
        self.approvals_table.verticalHeader().setVisible(False)
        self.approvals_table.setAlternatingRowColors(True)
        self.approvals_table.setShowGrid(False)
        self.approvals_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.approvals_table.setSelectionMode(QTableWidget.SingleSelection)
        self.approvals_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.approvals_table)
        self.command_detail = QPlainTextEdit()
        self.command_detail.setReadOnly(True)
        layout.addWidget(self.command_detail)
        row = QHBoxLayout()
        approve_btn = QPushButton("Approve")
        approve_btn.clicked.connect(lambda: self.change_selected_command("approved"))
        deny_btn = QPushButton("Deny")
        deny_btn.clicked.connect(lambda: self.change_selected_command("denied"))
        row.addStretch(1)
        row.addWidget(approve_btn)
        row.addWidget(deny_btn)
        layout.addLayout(row)
        return tab

    def build_explorer_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)

        helper = QLabel("Browse this PC and the other PC side by side, then move files or folders directly over SSH.")
        helper.setWordWrap(True)
        helper.setObjectName("cardMeta")
        layout.addWidget(helper)

        header = QHBoxLayout()
        header.addWidget(QLabel("Remote PC"))
        header.addWidget(self.explorer_target_combo)
        refresh_btn = QPushButton("Refresh Both")
        refresh_btn.clicked.connect(self.refresh_explorer)
        header.addWidget(refresh_btn)
        test_btn = QPushButton("Test SSH")
        test_btn.setObjectName("secondary")
        test_btn.clicked.connect(self.test_explorer_remote)
        header.addWidget(test_btn)
        header.addStretch(1)
        layout.addLayout(header)

        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter, 1)

        local_panel = QWidget()
        local_layout = QVBoxLayout(local_panel)
        local_layout.addWidget(QLabel("This PC"))
        local_path_row = QHBoxLayout()
        local_up_btn = QPushButton("Up")
        local_up_btn.setObjectName("secondary")
        local_up_btn.clicked.connect(self.local_up)
        local_path_row.addWidget(local_up_btn)
        local_path_row.addWidget(self.local_path_edit, 1)
        local_go_btn = QPushButton("Go")
        local_go_btn.setObjectName("secondary")
        local_go_btn.clicked.connect(self.refresh_local_browser)
        local_path_row.addWidget(local_go_btn)
        local_layout.addLayout(local_path_row)
        self.local_table = QTableWidget(0, 4)
        self.local_table.setHorizontalHeaderLabels(["Name", "Type", "Size", "Modified"])
        self.local_table.verticalHeader().setVisible(False)
        self.local_table.setAlternatingRowColors(True)
        self.local_table.setShowGrid(False)
        self.local_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.local_table.setSelectionMode(QTableWidget.SingleSelection)
        self.local_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.local_table.doubleClicked.connect(self.open_local_selected)
        local_layout.addWidget(self.local_table, 1)
        splitter.addWidget(local_panel)

        remote_panel = QWidget()
        remote_layout = QVBoxLayout(remote_panel)
        remote_layout.addWidget(QLabel("Remote PC"))
        remote_path_row = QHBoxLayout()
        remote_up_btn = QPushButton("Up")
        remote_up_btn.setObjectName("secondary")
        remote_up_btn.clicked.connect(self.remote_up)
        remote_path_row.addWidget(remote_up_btn)
        remote_path_row.addWidget(self.remote_path_edit, 1)
        remote_go_btn = QPushButton("Go")
        remote_go_btn.setObjectName("secondary")
        remote_go_btn.clicked.connect(self.refresh_remote_browser)
        remote_path_row.addWidget(remote_go_btn)
        remote_layout.addLayout(remote_path_row)
        self.remote_table = QTableWidget(0, 4)
        self.remote_table.setHorizontalHeaderLabels(["Name", "Type", "Size", "Modified"])
        self.remote_table.verticalHeader().setVisible(False)
        self.remote_table.setAlternatingRowColors(True)
        self.remote_table.setShowGrid(False)
        self.remote_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.remote_table.setSelectionMode(QTableWidget.SingleSelection)
        self.remote_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.remote_table.doubleClicked.connect(self.open_remote_selected)
        remote_layout.addWidget(self.remote_table, 1)
        splitter.addWidget(remote_panel)

        actions = QHBoxLayout()
        push_btn = QPushButton("Send Selected -> Remote Folder")
        push_btn.clicked.connect(self.transfer_local_to_remote)
        pull_btn = QPushButton("<- Pull Selected To Local Folder")
        pull_btn.clicked.connect(self.transfer_remote_to_local)
        actions.addWidget(push_btn)
        actions.addWidget(pull_btn)
        layout.addLayout(actions)
        return tab

    def build_projects_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        helper = QLabel("This should feel like a project radar, not a mystery box. It shows what each PC says it is working on and what folders it has published.")
        helper.setWordWrap(True)
        helper.setObjectName("cardMeta")
        layout.addWidget(helper)
        top = QFormLayout()
        self.focus_project_edit = QLineEdit(self.config.focus_project)
        self.focus_note_edit = QPlainTextEdit(self.config.focus_note)
        top.addRow("Current Project", self.focus_project_edit)
        top.addRow("Current Note", self.focus_note_edit)
        layout.addLayout(top)
        publish_btn = QPushButton("Publish Work Update")
        publish_btn.clicked.connect(self.publish_focus)
        layout.addWidget(publish_btn)
        self.projects_table = QTableWidget(0, 4)
        self.projects_table.setHorizontalHeaderLabels(["Project", "Known On", "Last Update", "Latest Note"])
        self.projects_table.verticalHeader().setVisible(False)
        self.projects_table.setAlternatingRowColors(True)
        self.projects_table.setShowGrid(False)
        self.projects_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.projects_table.setSelectionMode(QTableWidget.SingleSelection)
        self.projects_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.projects_table)
        return tab

    def build_setup_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        behavior_group = QGroupBox("Bridge Behavior")
        behavior_form = QFormLayout(behavior_group)
        self.auto_reply_checkbox = QCheckBox("Allow automatic chat replies to this node")
        self.auto_reply_checkbox.setChecked(self.config.auto_reply_enabled)
        behavior_form.addRow("", self.auto_reply_checkbox)
        self.auto_reply_model_edit = QLineEdit(self.config.auto_reply_model)
        behavior_form.addRow("Reply Model", self.auto_reply_model_edit)
        self.auto_convert_checkbox = QCheckBox("Auto-convert trusted actionable chat into commands")
        self.auto_convert_checkbox.setChecked(self.config.auto_convert_trusted_chat_commands)
        behavior_form.addRow("", self.auto_convert_checkbox)
        self.auto_execute_checkbox = QCheckBox("Auto-approve trusted incoming commands")
        self.auto_execute_checkbox.setChecked(self.config.auto_execute_trusted_commands)
        behavior_form.addRow("", self.auto_execute_checkbox)
        layout.addWidget(behavior_group)

        access_group = QGroupBox("Remote Access Profiles")
        access_layout = QGridLayout(access_group)
        access_layout.addWidget(QLabel("Home Host"), 0, 0)
        self.home_host_edit = QLineEdit(self.config.home_host)
        access_layout.addWidget(self.home_host_edit, 0, 1)
        access_layout.addWidget(QLabel("Home User"), 0, 2)
        self.home_user_edit = QLineEdit(self.config.home_user)
        access_layout.addWidget(self.home_user_edit, 0, 3)
        access_layout.addWidget(QLabel("Home Key"), 1, 0)
        self.home_key_edit = QLineEdit(self.config.home_key)
        access_layout.addWidget(self.home_key_edit, 1, 1, 1, 3)
        access_layout.addWidget(QLabel("Office Host"), 2, 0)
        self.office_host_edit = QLineEdit(self.config.office_host)
        access_layout.addWidget(self.office_host_edit, 2, 1)
        access_layout.addWidget(QLabel("Office User"), 2, 2)
        self.office_user_edit = QLineEdit(self.config.office_user)
        access_layout.addWidget(self.office_user_edit, 2, 3)
        access_layout.addWidget(QLabel("Office Key"), 3, 0)
        self.office_key_edit = QLineEdit(self.config.office_key)
        access_layout.addWidget(self.office_key_edit, 3, 1, 1, 3)
        layout.addWidget(access_group)

        self.project_roots_list = QListWidget()
        for path in self.config.project_roots:
            self.project_roots_list.addItem(path)
        roots_note = QLabel("Only publish roots you want the other side to know about. This keeps the radar useful instead of cluttered.")
        roots_note.setWordWrap(True)
        roots_note.setObjectName("cardMeta")
        layout.addWidget(roots_note)
        layout.addWidget(QLabel("Project roots to publish"))
        layout.addWidget(self.project_roots_list)
        row = QHBoxLayout()
        add_btn = QPushButton("Add Root")
        add_btn.clicked.connect(self.add_project_root)
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self.remove_project_root)
        row.addWidget(add_btn)
        row.addWidget(remove_btn)
        row.addStretch(1)
        layout.addLayout(row)
        return tab

    def build_log_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        helper = QLabel("Operational notes, command output highlights, and the things worth remembering while we harden the bridge.")
        helper.setWordWrap(True)
        helper.setObjectName("cardMeta")
        layout.addWidget(helper)
        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)
        return tab

    def build_tray(self) -> None:
        icon = QPixmap(32, 32)
        icon.fill(QColor("#14213d"))
        painter = QPainter(icon)
        painter.setPen(QColor("#fca311"))
        painter.setFont(QFont("Segoe UI", 12, QFont.Bold))
        painter.drawText(icon.rect(), Qt.AlignCenter, "CB")
        painter.end()
        self.tray = QSystemTrayIcon(QIcon(icon), self)
        menu = self.menuBar().addMenu("Bridge")
        show_action = QAction("Show", self)
        show_action.triggered.connect(self.showNormal)
        hide_action = QAction("Hide To Tray", self)
        hide_action.triggered.connect(self.hide)
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        tray_menu = QMenu(self)
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addAction(exit_action)
        self.tray.setContextMenu(tray_menu)
        self.tray.activated.connect(lambda reason: self.showNormal() if reason == QSystemTrayIcon.Trigger else None)
        self.tray.show()
        menu.addAction(show_action)
        menu.addAction(hide_action)
        menu.addAction(exit_action)

    def closeEvent(self, event) -> None:  # type: ignore[override]
        if self.tray.isVisible():
            self.hide()
            self.status.showMessage("Bridge hidden to tray and still running.")
            event.ignore()
            return
        super().closeEvent(event)

    def choose_bridge_root(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Choose synced bridge folder", self.bridge_root_edit.text())
        if path:
            self.bridge_root_edit.setText(path)

    def choose_command_workdir(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Choose command working directory")
        if path:
            self.command_workdir_edit.setText(path)

    def add_project_root(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Choose project root")
        if not path:
            return
        for index in range(self.project_roots_list.count()):
            if self.project_roots_list.item(index).text() == path:
                return
        self.project_roots_list.addItem(path)

    def remove_project_root(self) -> None:
        for item in self.project_roots_list.selectedItems():
            self.project_roots_list.takeItem(self.project_roots_list.row(item))

    def save_and_refresh(self) -> None:
        self.save_config()
        self.refresh_all()
        self.status.showMessage("Bridge settings saved.")

    def parse_direct_command_from_chat(self, text: str) -> tuple[str, str] | None:
        stripped = text.strip()
        lower = stripped.lower()
        for prefix in ("run:", "ps:", "powershell:", "cmd:"):
            if lower.startswith(prefix):
                command_text = stripped[len(prefix):].strip()
                if command_text:
                    return command_text, command_text.splitlines()[0][:120]
        return None

    def parse_action_from_chat(self, text: str) -> tuple[str, str, str] | None:
        stripped = text.strip()
        lower = stripped.lower()
        if any(phrase in lower for phrase in ("print a test page", "print test page", "print the test page")):
            return ("print_test_page", "", "Print a test page")
        if any(phrase in lower for phrase in ("what apps are running", "what programs are running", "list running apps", "list running processes", "show running processes")):
            return (
                "shell_command",
                "Get-Process | Sort-Object ProcessName | Select-Object -First 120 ProcessName,Id,MainWindowTitle | Format-Table -AutoSize",
                "List running processes",
            )
        direct = self.parse_direct_command_from_chat(text)
        if direct:
            command_text, summary = direct
            return ("shell_command", command_text, summary)
        return None

    def send_chat(self) -> None:
        text = self.chat_input.toPlainText().strip()
        if not text:
            QMessageBox.information(self, "No Message", "Type a message first.")
            return
        self.save_config()
        target = self.chat_target_combo.currentText()
        parsed = self.parse_action_from_chat(text) if target in {"home", "office"} else None
        extra: dict | None = None
        status_note = "Chat message sent."
        direct_target = self.direct_bridge_target(target)
        if parsed:
            action, command_text, summary = parsed
            command_payload = self.store.build_command_payload(
                target,
                summary,
                command_text,
                "",
                "",
                approval_required=False,
                action=action,
                source_message_id="",
            )
            command_filename = self.store.build_filename(command_payload["created_at"], command_payload["id"])
            self.write_bridge_payload("commands", command_filename, command_payload, direct_target=direct_target, direct_first=True)
            extra = {"routed_action": True}
            system_payload = self.store.build_message_payload("system", f"Routed actionable chat to {target}: {summary}", target)
            system_filename = self.store.build_filename(system_payload["created_at"], system_payload["id"])
            self.write_bridge_payload("messages", system_filename, system_payload, direct_target=direct_target, direct_first=True)
            status_note = "Actionable chat routed into a direct command."
        chat_payload = self.store.build_message_payload("chat", text, target, extra)
        chat_filename = self.store.build_filename(chat_payload["created_at"], chat_payload["id"])
        self.write_bridge_payload("messages", chat_filename, chat_payload, direct_target=direct_target if parsed else None, direct_first=bool(parsed))
        self.chat_input.clear()
        self.refresh_all()
        self.status.showMessage(status_note)

    def publish_focus(self) -> None:
        self.save_config()
        self.store.send_message("system", f"{self.config.display_name} updated current work: {self.config.focus_project or 'No project named'}")
        self.refresh_all()
        self.status.showMessage("Current work update published.")

    def send_command(self) -> None:
        command = self.command_input.toPlainText().strip()
        if not command:
            QMessageBox.information(self, "No Command", "Type a PowerShell command first.")
            return
        self.save_config()
        target = self.command_target_combo.currentText()
        payload = self.store.build_command_payload(
            target,
            self.command_summary_edit.text().strip(),
            command,
            self.command_workdir_edit.text().strip(),
            self.command_project_edit.text().strip(),
            self.command_requires_approval.isChecked(),
        )
        filename = self.store.build_filename(payload["created_at"], payload["id"])
        self.write_bridge_payload("commands", filename, payload, direct_target=self.direct_bridge_target(target), direct_first=True)
        self.command_input.clear()
        self.command_summary_edit.clear()
        self.refresh_all()
        self.status.showMessage("Command request sent.")

    def selected_command(self) -> dict | None:
        row = self.approvals_table.currentRow()
        if row < 0 or row >= len(getattr(self, "visible_commands", [])):
            return None
        command_id = self.approvals_table.item(row, 0).data(Qt.UserRole)
        for record in self.latest_commands:
            if record.get("id") == command_id:
                return record
        return None

    def populate_command_detail(self) -> None:
        record = self.selected_command()
        if not record:
            self.command_detail.setPlainText("Select a command request to inspect it.")
            return
        lines = [
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
            lines.extend(
                [
                    "",
                    f"Exit code: {result.get('exit_code', '')}",
                    f"Executed by: {result.get('executor', '')}",
                    f"Ran at: {result.get('ran_at', '')}",
                    "",
                    "Stdout:",
                    result.get("stdout", ""),
                    "",
                    "Stderr:",
                    result.get("stderr", ""),
                ]
            )
        self.command_detail.setPlainText("\n".join(lines))

    def change_selected_command(self, status: str) -> None:
        record = self.selected_command()
        if not record:
            return
        record["status"] = status
        record["updated_at"] = now_iso()
        record["approved_by"] = self.config.display_name
        self.store.update_command(record)
        self.store.send_message("system", f"{self.config.display_name} {status} command: {record.get('summary', '')}")
        self.refresh_all()

    def append_log(self, text: str) -> None:
        self.log_output.appendPlainText(f"[{now_iso()}] {text}\n")

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

    def direct_access_profile(self, target: str) -> tuple[str, str, str]:
        if target == "home":
            return (
                self.config.home_host.strip(),
                self.config.home_user.strip() or "ferna",
                self.config.home_key.strip(),
            )
        return (
            self.config.office_host.strip(),
            self.config.office_user.strip() or "ferna",
            self.config.office_key.strip(),
        )

    def validate_direct_access_profile(self, target: str) -> tuple[str, str, str] | None:
        self.save_config()
        host, user, key = self.direct_access_profile(target)
        if not host or not user or not key:
            QMessageBox.information(self, "Missing Access Profile", f"Set the {target} SSH host, user, and key in Control Center first.")
            return None
        if not Path(key).exists():
            QMessageBox.information(self, "Missing Key File", f"The configured {target} SSH key does not exist:\n{key}")
            return None
        return host, user, key

    def direct_bridge_target(self, target: str) -> str | None:
        local = (self.config.local_node or "home").lower()
        target = (target or "all").lower()
        if target in {"home", "office"} and target != local:
            return target
        return None

    def remote_bridge_path_exists(self, target: str, subdir: str, filename: str) -> bool:
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return False
        host, user, key = profile
        remote_path = str((self.store.bridge_root / subdir / filename)).replace("/", "\\")
        script = f"if (Test-Path {self.powershell_quote(remote_path)}) {{ 'present' }} else {{ 'missing' }}"
        encoded = subprocess.run(
            ["powershell", "-NoProfile", "-Command", "[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($input))"],
            input=script,
            capture_output=True,
            text=True,
            timeout=30,
            creationflags=CREATE_NO_WINDOW,
        ).stdout.strip()
        completed = subprocess.run(
            [
                r"C:\Windows\System32\OpenSSH\ssh.exe",
                "-o",
                "BatchMode=yes",
                "-o",
                "StrictHostKeyChecking=accept-new",
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

    def deliver_bridge_payload_direct(self, target: str, subdir: str, filename: str, payload: dict) -> None:
        profile = self.validate_direct_access_profile(target)
        if not profile:
            raise RuntimeError(f"Missing {target} SSH profile.")
        host, user, key = profile
        remote_dir = str((self.store.bridge_root / subdir)).replace("/", "\\")
        prep_script = f"New-Item -ItemType Directory -Force -Path {self.powershell_quote(remote_dir)} | Out-Null"
        prep_encoded = subprocess.run(
            ["powershell", "-NoProfile", "-Command", "[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($input))"],
            input=prep_script,
            capture_output=True,
            text=True,
            timeout=30,
            creationflags=CREATE_NO_WINDOW,
        ).stdout.strip()
        prep = subprocess.run(
            [
                r"C:\Windows\System32\OpenSSH\ssh.exe",
                "-o",
                "BatchMode=yes",
                "-o",
                "StrictHostKeyChecking=accept-new",
                "-i",
                key,
                f"{user}@{host}",
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-EncodedCommand",
                prep_encoded,
            ],
            capture_output=True,
            text=True,
            timeout=60,
            creationflags=CREATE_NO_WINDOW,
        )
        if prep.returncode != 0:
            raise RuntimeError(prep.stderr.strip() or prep.stdout.strip() or "Failed to prepare remote bridge folder.")

        temp_file = Path(tempfile.gettempdir()) / f"codex_bridge_{payload.get('id', 'payload')}.json"
        temp_file.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        try:
            completed = subprocess.run(
                [
                    r"C:\Windows\System32\OpenSSH\scp.exe",
                    "-o",
                    "BatchMode=yes",
                    "-o",
                    "StrictHostKeyChecking=accept-new",
                    "-i",
                    key,
                    str(temp_file),
                    f"{user}@{host}:{self.remote_scp_path(str((self.store.bridge_root / subdir / filename)))}",
                ],
                capture_output=True,
                text=True,
                timeout=120,
                creationflags=CREATE_NO_WINDOW,
            )
            if completed.returncode != 0:
                raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or "Direct bridge delivery failed.")
        finally:
            try:
                temp_file.unlink()
            except OSError:
                pass

    def write_bridge_payload(self, subdir: str, filename: str, payload: dict, direct_target: str | None = None, direct_first: bool = False) -> None:
        if direct_target and direct_first:
            self.deliver_bridge_payload_direct(direct_target, subdir, filename, payload)
            if not self.remote_bridge_path_exists(direct_target, subdir, filename):
                raise RuntimeError(f"Direct receipt verification failed for {direct_target}:{subdir}/{filename}")
        self.store.write_json(self.store.bridge_root / subdir / filename, payload)
        if direct_target and not direct_first:
            self.deliver_bridge_payload_direct(direct_target, subdir, filename, payload)
            if not self.remote_bridge_path_exists(direct_target, subdir, filename):
                raise RuntimeError(f"Direct receipt verification failed for {direct_target}:{subdir}/{filename}")

    def encoded_powershell(self, command_text: str) -> str:
        return command_text.encode("utf-16le").hex()

    def run_remote_json(self, target: str, command_text: str, timeout: int = 120) -> list[dict]:
        profile = self.validate_direct_access_profile(target)
        if not profile:
            raise RuntimeError(f"Missing {target} SSH profile.")
        host, user, key = profile
        command_text = "$ProgressPreference='SilentlyContinue'; $ErrorActionPreference='Stop'; " + command_text
        encoded = subprocess.run(
            ["powershell", "-NoProfile", "-Command", "[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($input))"],
            input=command_text,
            capture_output=True,
            text=True,
            timeout=30,
            creationflags=CREATE_NO_WINDOW,
        ).stdout.strip()
        completed = subprocess.run(
            [
                r"C:\Windows\System32\OpenSSH\ssh.exe",
                "-o",
                "BatchMode=yes",
                "-o",
                "StrictHostKeyChecking=accept-new",
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
        text = (completed.stdout or "").strip()
        if not text:
            return []
        lines = [line.strip() for line in text.splitlines() if line.strip() and not line.lstrip().startswith("#<")]
        cleaned = "\n".join(lines)
        start = min((idx for idx in [cleaned.find("["), cleaned.find("{")] if idx != -1), default=-1)
        if start > 0:
            cleaned = cleaned[start:]
        payload = json.loads(cleaned)
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

    def populate_table(self, table: QTableWidget, entries: list[dict]) -> None:
        table.setRowCount(len(entries))
        for row, entry in enumerate(entries):
            name_item = QTableWidgetItem(entry.get("Name", ""))
            name_item.setData(Qt.UserRole, row)
            table.setItem(row, 0, name_item)
            table.setItem(row, 1, QTableWidgetItem("Folder" if entry.get("IsDirectory") else "File"))
            table.setItem(row, 2, QTableWidgetItem(self.format_size(entry.get("Size"))))
            table.setItem(row, 3, QTableWidgetItem(entry.get("Modified", "")))
        table.resizeColumnsToContents()

    def selected_entry(self, table: QTableWidget, entries: list[dict]) -> dict | None:
        row = table.currentRow()
        if row < 0 or row >= len(entries):
            return None
        return entries[row]

    def refresh_explorer(self) -> None:
        self.refresh_local_browser()
        self.refresh_remote_browser()

    def refresh_local_browser(self) -> None:
        path = Path(self.local_path_edit.text().strip() or str(Path.home())).resolve()
        self.local_path_edit.setText(str(path))
        self.local_entries = self.local_dir_entries(path)
        self.populate_table(self.local_table, self.local_entries)

    def refresh_remote_browser(self) -> None:
        if self.remote_refresh_inflight:
            return
        target = self.explorer_target_combo.currentText().strip() or ("office" if self.config.local_node == "home" else "home")
        path_text = self.remote_path_edit.text().strip() or r"C:\Users\ferna"
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
                    self.populate_table(self.remote_table, self.remote_entries)
                    self.status.showMessage(f"Remote browser refreshed for {target}.")

                QTimer.singleShot(0, apply_entries)
            except Exception as exc:
                QTimer.singleShot(0, lambda: self.append_log(f"Remote browser refresh failed for {target}: {exc}"))
                QTimer.singleShot(0, lambda: self.status.showMessage(f"Remote browser refresh failed for {target}: {exc}"))
            finally:
                QTimer.singleShot(0, lambda: setattr(self, "remote_refresh_inflight", False))

        import threading
        threading.Thread(target=worker, daemon=True).start()

    def local_up(self) -> None:
        current = Path(self.local_path_edit.text().strip() or str(Path.home()))
        parent = current.parent if current.parent != current else current
        self.local_path_edit.setText(str(parent))
        self.refresh_local_browser()

    def remote_up(self) -> None:
        current = Path(self.remote_path_edit.text().strip() or r"C:\Users\ferna")
        parent = current.parent if current.parent != current else current
        self.remote_path_edit.setText(str(parent))
        self.refresh_remote_browser()

    def open_local_selected(self) -> None:
        entry = self.selected_entry(self.local_table, self.local_entries)
        if entry and entry.get("IsDirectory"):
            self.local_path_edit.setText(entry["FullPath"])
            self.refresh_local_browser()

    def open_remote_selected(self) -> None:
        entry = self.selected_entry(self.remote_table, self.remote_entries)
        if entry and entry.get("IsDirectory"):
            self.remote_path_edit.setText(entry["FullPath"])
            self.refresh_remote_browser()

    def execute_scp_transfer(self, mode: str, target: str, host: str, user: str, key: str, source_path: str, dest_path: str, display_name: str) -> None:
        try:
            scp_exe = r"C:\Windows\System32\OpenSSH\scp.exe"
            if mode == "upload":
                command = [scp_exe, "-o", "BatchMode=yes", "-o", "StrictHostKeyChecking=accept-new", "-r", "-i", key, source_path, f"{user}@{host}:{self.remote_scp_path(dest_path)}"]
            else:
                Path(dest_path).mkdir(parents=True, exist_ok=True)
                command = [scp_exe, "-o", "BatchMode=yes", "-o", "StrictHostKeyChecking=accept-new", "-r", "-i", key, f"{user}@{host}:{self.remote_scp_path(source_path)}", dest_path]
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=3600,
                creationflags=CREATE_NO_WINDOW,
            )
            if completed.returncode != 0:
                raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or f"SCP failed with exit code {completed.returncode}")
            self.append_log(f"{mode.title()} complete for {display_name} on {target}.")
            QTimer.singleShot(0, self.refresh_explorer)
            QTimer.singleShot(0, lambda: self.status.showMessage(f"{mode.title()} complete: {display_name}"))
        except Exception as exc:
            self.append_log(f"{mode.title()} failed for {display_name} on {target}: {exc}")
            QTimer.singleShot(0, lambda: QMessageBox.critical(self, "Transfer Failed", str(exc)))
            QTimer.singleShot(0, lambda: self.status.showMessage(f"{mode.title()} failed for {display_name}."))

    def transfer_local_to_remote(self) -> None:
        entry = self.selected_entry(self.local_table, self.local_entries)
        if not entry:
            QMessageBox.information(self, "No Selection", "Pick a local file or folder first.")
            return
        target = self.explorer_target_combo.currentText().strip() or ("office" if self.config.local_node == "home" else "home")
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        import threading
        threading.Thread(
            target=self.execute_scp_transfer,
            args=("upload", target, host, user, key, entry["FullPath"], self.remote_path_edit.text().strip(), entry["Name"]),
            daemon=True,
        ).start()
        self.status.showMessage(f"Sending {entry['Name']} to {target}...")

    def transfer_remote_to_local(self) -> None:
        entry = self.selected_entry(self.remote_table, self.remote_entries)
        if not entry:
            QMessageBox.information(self, "No Selection", "Pick a remote file or folder first.")
            return
        target = self.explorer_target_combo.currentText().strip() or ("office" if self.config.local_node == "home" else "home")
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        import threading
        threading.Thread(
            target=self.execute_scp_transfer,
            args=("download", target, host, user, key, entry["FullPath"], self.local_path_edit.text().strip(), entry["Name"]),
            daemon=True,
        ).start()
        self.status.showMessage(f"Pulling {entry['Name']} from {target}...")

    def test_explorer_remote(self) -> None:
        target = self.explorer_target_combo.currentText().strip() or ("office" if self.config.local_node == "home" else "home")
        profile = self.validate_direct_access_profile(target)
        if not profile:
            return
        host, user, key = profile
        try:
            completed = subprocess.run(
                [r"C:\Windows\System32\OpenSSH\ssh.exe", "-o", "BatchMode=yes", "-o", "StrictHostKeyChecking=accept-new", "-i", key, f"{user}@{host}", "hostname"],
                capture_output=True,
                text=True,
                timeout=120,
                creationflags=CREATE_NO_WINDOW,
            )
            if completed.returncode != 0:
                raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or f"SSH failed with exit code {completed.returncode}")
            QMessageBox.information(self, "SSH Connected", f"{target.title()} is reachable.\n\nHost replied: {completed.stdout.strip() or host}")
            self.status.showMessage(f"SSH test succeeded for {target}.")
        except Exception as exc:
            self.append_log(f"SSH test failed for {target}: {exc}")
            QMessageBox.critical(self, "SSH Test Failed", str(exc))

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
                return f"I ran '{summary}' successfully. First result line: {stdout.splitlines()[0][:260]}"
            return f"I ran '{summary}' successfully."
        if stderr:
            return f"I tried to run '{summary}', but it failed. Error: {stderr.splitlines()[0][:260]}"
        return f"I tried to run '{summary}', but it failed with exit code {exit_code}."

    def print_test_page(self) -> dict:
        temp_dir = Path(tempfile.gettempdir()) / "codex_bridge"
        temp_dir.mkdir(parents=True, exist_ok=True)
        file_path = temp_dir / f"codex_bridge_test_page_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        body = (
            "Codex Bridge Test Page\r\n"
            f"Printed at: {now_iso()}\r\n"
            f"Node: {self.config.local_node}\r\n"
            f"Display: {self.config.display_name}\r\n"
        )
        file_path.write_text(body, encoding="utf-8")
        command_text = f'Start-Process -FilePath notepad.exe -ArgumentList "/p","{file_path}" -Wait'
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command_text],
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

    def refresh_all(self) -> None:
        self.save_config()
        self.store.ensure_dirs()
        self.store.publish_state()
        self.latest_messages = self.store.messages()
        self.latest_commands = self.store.commands()
        self.latest_states = self.store.node_states()
        self.refresh_chat()
        self.refresh_approvals()
        self.refresh_projects()
        self.refresh_summary_cards()
        self.run_approved_commands()

    def refresh_chat(self) -> None:
        local = self.config.local_node
        opposite = "office" if local == "home" else "home"
        lines: list[str] = []
        for item in self.latest_messages[-200:]:
            target = item.get("target", "all")
            if target not in ("all", local, opposite):
                continue
            sender = item.get("sender_name") or item.get("sender", "")
            stamp = item.get("created_at", "")
            kind = item.get("kind", "chat")
            lines.append(f"[{format_timestamp(stamp)}] {sender} [{kind}]")
            lines.append(item.get("text", ""))
            lines.append("")
        self.chat_log.setPlainText("\n".join(lines).strip())
        self.chat_log.verticalScrollBar().setValue(self.chat_log.verticalScrollBar().maximum())

    def refresh_approvals(self) -> None:
        local = self.config.local_node
        visible = [record for record in self.latest_commands if record.get("target") == local or record.get("sender") == local]
        self.visible_commands = visible
        self.approvals_table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            direction = "Incoming" if record.get("target") == local else "Outgoing"
            other_side = record.get("sender_name") or record.get("sender", "") if direction == "Incoming" else record.get("target", "")
            created_item = QTableWidgetItem(format_timestamp(record.get("created_at", "")))
            created_item.setData(Qt.UserRole, record.get("id"))
            self.approvals_table.setItem(row, 0, created_item)
            self.approvals_table.setItem(row, 1, QTableWidgetItem(direction))
            self.approvals_table.setItem(row, 2, QTableWidgetItem(other_side))
            self.approvals_table.setItem(row, 3, QTableWidgetItem(record.get("summary", "")))
            self.approvals_table.setItem(row, 4, QTableWidgetItem(record.get("status", "")))
        self.approvals_table.resizeColumnsToContents()
        self.populate_command_detail()

    def refresh_projects(self) -> None:
        aggregate: dict[str, dict] = {}
        for node, state in self.latest_states.items():
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
        self.projects_table.setRowCount(len(aggregate))
        for row, name in enumerate(sorted(aggregate)):
            info = aggregate[name]
            self.projects_table.setItem(row, 0, QTableWidgetItem(name))
            self.projects_table.setItem(row, 1, QTableWidgetItem(", ".join(sorted(info["nodes"]))))
            self.projects_table.setItem(row, 2, QTableWidgetItem(format_timestamp(info.get("last_update", ""))))
            self.projects_table.setItem(row, 3, QTableWidgetItem(info.get("last_note", "")))
        self.projects_table.resizeColumnsToContents()

    def refresh_summary_cards(self) -> None:
        local_state = self.latest_states.get(self.config.local_node, {})
        other_node = "office" if self.config.local_node == "home" else "home"
        remote_state = self.latest_states.get(other_node, {})
        local_projects = len(local_state.get("projects", []))
        remote_projects = len(remote_state.get("projects", []))
        self.local_card.set_value(
            self.config.display_name,
            f"{local_projects} published project folders\nHeartbeat {format_age(local_state.get('updated_at', ''))}",
        )
        remote_name = remote_state.get("display_name") or ("Office Codex" if other_node == "office" else "Home Codex")
        remote_focus = remote_state.get("focus_project") or "No work note published"
        self.remote_card.set_value(
            remote_name,
            f"{remote_projects} published project folders\nCurrent work: {remote_focus}\nHeartbeat {format_age(remote_state.get('updated_at', ''))}",
        )
        incoming = sum(1 for record in self.latest_commands if record.get("target") == self.config.local_node and record.get("status") in {"pending_approval", "approved", "running"})
        outgoing = sum(1 for record in self.latest_commands if record.get("sender") == self.config.local_node and record.get("status") in {"pending_approval", "approved", "running"})
        self.queue_card.set_value(
            f"{incoming} in / {outgoing} out",
            "Live bridge work waiting or executing across both PCs.",
        )
        radar_count = len({item.get("name") for state in self.latest_states.values() for item in state.get("projects", []) if item.get("name")})
        self.projects_card.set_value(
            str(radar_count),
            "Distinct known project folders published across both PCs.",
        )

    def run_approved_commands(self) -> None:
        local = self.config.local_node
        for record in self.latest_commands:
            if record.get("target") != local or record.get("status") != "approved":
                continue
            command_id = record.get("id")
            if not command_id or command_id in self.store.active_commands:
                continue
            self.store.active_commands.add(command_id)
            self.execute_command(record)

    def execute_command(self, record: dict) -> None:
        record["status"] = "running"
        record["updated_at"] = now_iso()
        self.store.update_command(record)
        try:
            action = str(record.get("action") or record.get("kind") or "shell_command").strip()
            if action == "print_test_page":
                result = self.print_test_page()
                exit_code = result.get("exit_code", 1)
            else:
                completed = subprocess.run(
                    ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", record.get("command", "")],
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
                exit_code = completed.returncode
            record["status"] = "completed" if exit_code == 0 else "failed"
            record["result"] = {
                **result,
                "ran_at": now_iso(),
                "action": action,
                "executor": self.config.display_name,
            }
            self.store.update_command(record)
            self.store.send_message("system", f"{self.config.display_name} completed command '{record.get('summary', '')}' with exit code {exit_code}.")
            reply_target = str(record.get("sender") or "").strip()
            if reply_target:
                reply_payload = self.store.build_message_payload(
                    "chat",
                    self.summarize_result_for_chat(record),
                    reply_target,
                    {"reply_to": record.get("source_message_id", "")},
                )
                reply_filename = self.store.build_filename(reply_payload["created_at"], reply_payload["id"])
                self.write_bridge_payload("messages", reply_filename, reply_payload, direct_target=self.direct_bridge_target(reply_target), direct_first=True)
            if str(result.get("stdout") or "").strip():
                self.append_log("Stdout:\n" + str(result["stdout"])[-4000:])
            if str(result.get("stderr") or "").strip():
                self.append_log("Stderr:\n" + str(result["stderr"])[-4000:])
        except Exception as exc:
            record["status"] = "failed"
            record["result"] = {"exit_code": "error", "stdout": "", "stderr": str(exc), "ran_at": now_iso(), "executor": self.config.display_name}
            self.store.update_command(record)
            self.store.send_message("system", f"{self.config.display_name} failed command '{record.get('summary', '')}'.")
            self.append_log(f"Command failed: {exc}")
        finally:
            if record.get("id") in self.store.active_commands:
                self.store.active_commands.remove(record.get("id"))


def main() -> None:
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    window = BridgeWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
