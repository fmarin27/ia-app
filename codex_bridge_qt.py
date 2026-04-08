from __future__ import annotations

import json
import site
import subprocess
import sys
import uuid
from dataclasses import asdict, dataclass, field
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
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

CONFIG_FILE = APP_DIR / "codex_bridge_config.json"
DEFAULT_BRIDGE_ROOT = Path.home() / "Desktop" / "codex-bridge"
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

    def send_message(self, kind: str, text: str, target: str = "all") -> None:
        payload = {
            "id": make_id("msg"),
            "kind": kind,
            "sender": self.config.local_node,
            "sender_name": self.config.display_name,
            "target": target,
            "text": text,
            "created_at": now_iso(),
        }
        filename = f"{payload['created_at'].replace(':', '').replace('-', '')}_{payload['id']}.json"
        self.write_json(self.messages_dir / filename, payload)

    def send_command(self, target: str, summary: str, command: str, workdir: str, project: str, approval_required: bool) -> None:
        payload = {
            "id": make_id("cmd"),
            "kind": "shell_command",
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
            "result": {},
        }
        filename = f"{payload['created_at'].replace(':', '').replace('-', '')}_{payload['id']}.json"
        self.write_json(self.commands_dir / filename, payload)

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
        self.setWindowTitle("Codex Bridge Qt")
        self.resize(1460, 920)
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
            return BridgeConfig(**json.loads(CONFIG_FILE.read_text(encoding="utf-8")))
        except Exception:
            return BridgeConfig()

    def save_config(self) -> None:
        self.config.local_node = self.node_combo.currentText() or "home"
        self.config.display_name = self.display_name_edit.text().strip() or ("Home Codex" if self.config.local_node == "home" else "Office Codex")
        self.config.bridge_root = self.bridge_root_edit.text().strip() or str(DEFAULT_BRIDGE_ROOT)
        self.config.focus_project = self.focus_project_edit.text().strip()
        self.config.focus_note = self.focus_note_edit.toPlainText().strip()
        self.config.project_roots = [self.project_roots_list.item(i).text() for i in range(self.project_roots_list.count())]
        CONFIG_FILE.write_text(json.dumps(asdict(self.config), indent=2), encoding="utf-8")
        self.store.config = self.config
        self.store.ensure_dirs()

    def stylesheet(self) -> str:
        return """
        QMainWindow { background: #0d1321; }
        QWidget { color: #eef4ed; font-size: 13px; }
        QFrame#card { background: #1d2d44; border-radius: 16px; }
        QLineEdit, QPlainTextEdit, QComboBox, QListWidget, QTableWidget {
            background: #132033; border: 1px solid #3e5c76; border-radius: 10px; padding: 8px;
        }
        QPushButton {
            background: #748cab; color: #0d1321; border: none; border-radius: 10px; padding: 10px 14px; font-weight: 700;
        }
        QPushButton:hover { background: #9db4c0; }
        QTabWidget::pane { border: none; }
        QTabBar::tab { background: #1d2d44; padding: 10px 14px; margin-right: 6px; border-top-left-radius: 10px; border-top-right-radius: 10px; }
        QTabBar::tab:selected { background: #3e5c76; }
        QHeaderView::section { background: #243b53; color: #eef4ed; padding: 8px; border: none; }
        """

    def build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        header = QFrame(objectName="card")
        header_layout = QGridLayout(header)
        header_layout.addWidget(QLabel("Node"), 0, 0)
        self.node_combo = QComboBox()
        self.node_combo.addItems(["home", "office"])
        self.node_combo.setCurrentText(self.config.local_node)
        header_layout.addWidget(self.node_combo, 0, 1)
        header_layout.addWidget(QLabel("Display Name"), 0, 2)
        self.display_name_edit = QLineEdit(self.config.display_name)
        header_layout.addWidget(self.display_name_edit, 0, 3)
        header_layout.addWidget(QLabel("Bridge Folder"), 0, 4)
        self.bridge_root_edit = QLineEdit(self.config.bridge_root)
        header_layout.addWidget(self.bridge_root_edit, 0, 5)
        choose_btn = QPushButton("Choose")
        choose_btn.clicked.connect(self.choose_bridge_root)
        header_layout.addWidget(choose_btn, 0, 6)
        hide_btn = QPushButton("Hide To Tray")
        hide_btn.clicked.connect(self.hide)
        header_layout.addWidget(hide_btn, 1, 5)
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_and_refresh)
        header_layout.addWidget(save_btn, 1, 6)
        layout.addWidget(header)

        splitter = QSplitter(Qt.Horizontal)
        layout.addWidget(splitter, 1)

        left_tabs = QTabWidget()
        right_tabs = QTabWidget()
        splitter.addWidget(left_tabs)
        splitter.addWidget(right_tabs)
        splitter.setSizes([850, 560])

        left_tabs.addTab(self.build_chat_tab(), "Chat")
        left_tabs.addTab(self.build_commands_tab(), "Commands")
        left_tabs.addTab(self.build_approvals_tab(), "Approvals")
        right_tabs.addTab(self.build_projects_tab(), "Projects")
        right_tabs.addTab(self.build_setup_tab(), "Setup")
        right_tabs.addTab(self.build_log_tab(), "Log")

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Bridge ready.")

    def build_chat_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.chat_log = QPlainTextEdit()
        self.chat_log.setReadOnly(True)
        layout.addWidget(self.chat_log)
        row = QHBoxLayout()
        self.chat_target_combo = QComboBox()
        self.chat_target_combo.addItems(["all", "home", "office"])
        row.addWidget(self.chat_target_combo)
        send_btn = QPushButton("Send")
        send_btn.clicked.connect(self.send_chat)
        row.addWidget(send_btn)
        layout.addLayout(row)
        self.chat_input = QPlainTextEdit()
        self.chat_input.setPlaceholderText("Type a message to both Codexes or target one side.")
        layout.addWidget(self.chat_input)
        return tab

    def build_commands_tab(self) -> QWidget:
        tab = QWidget()
        layout = QFormLayout(tab)
        self.command_target_combo = QComboBox()
        self.command_target_combo.addItems(["home", "office"])
        self.command_target_combo.setCurrentText("office" if self.config.local_node == "home" else "home")
        layout.addRow("Target", self.command_target_combo)
        self.command_project_edit = QLineEdit()
        layout.addRow("Project", self.command_project_edit)
        workdir_row = QHBoxLayout()
        self.command_workdir_edit = QLineEdit()
        workdir_row.addWidget(self.command_workdir_edit)
        workdir_btn = QPushButton("Choose")
        workdir_btn.clicked.connect(self.choose_command_workdir)
        workdir_row.addWidget(workdir_btn)
        wrap = QWidget()
        wrap.setLayout(workdir_row)
        layout.addRow("Working Dir", wrap)
        self.command_summary_edit = QLineEdit()
        layout.addRow("Summary", self.command_summary_edit)
        self.command_requires_approval = QCheckBox("Require approval before running")
        self.command_requires_approval.setChecked(True)
        layout.addRow("", self.command_requires_approval)
        self.command_input = QPlainTextEdit()
        self.command_input.setPlaceholderText("PowerShell command to run on the other PC.")
        layout.addRow("Command", self.command_input)
        send_btn = QPushButton("Send Command Request")
        send_btn.clicked.connect(self.send_command)
        layout.addRow("", send_btn)
        return tab

    def build_approvals_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.approvals_table = QTableWidget(0, 4)
        self.approvals_table.setHorizontalHeaderLabels(["Created", "Sender", "Summary", "Status"])
        self.approvals_table.itemSelectionChanged.connect(self.populate_command_detail)
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

    def build_projects_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        top = QFormLayout()
        self.focus_project_edit = QLineEdit(self.config.focus_project)
        self.focus_note_edit = QPlainTextEdit(self.config.focus_note)
        top.addRow("Focus Project", self.focus_project_edit)
        top.addRow("Focus Note", self.focus_note_edit)
        layout.addLayout(top)
        publish_btn = QPushButton("Publish Focus Update")
        publish_btn.clicked.connect(self.publish_focus)
        layout.addWidget(publish_btn)
        self.projects_table = QTableWidget(0, 4)
        self.projects_table.setHorizontalHeaderLabels(["Project", "Known On", "Last Update", "Latest Note"])
        layout.addWidget(self.projects_table)
        return tab

    def build_setup_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
        self.project_roots_list = QListWidget()
        for path in self.config.project_roots:
            self.project_roots_list.addItem(path)
        layout.addWidget(QLabel("Project roots to publish"))
        layout.addWidget(self.project_roots_list)
        row = QHBoxLayout()
        add_btn = QPushButton("Add Root")
        add_btn.clicked.connect(self.add_project_root)
        remove_btn = QPushButton("Remove Selected")
        remove_btn.clicked.connect(self.remove_project_root)
        row.addWidget(add_btn)
        row.addWidget(remove_btn)
        layout.addLayout(row)
        return tab

    def build_log_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout(tab)
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

    def send_chat(self) -> None:
        text = self.chat_input.toPlainText().strip()
        if not text:
            QMessageBox.information(self, "No Message", "Type a message first.")
            return
        self.save_config()
        self.store.send_message("chat", text, self.chat_target_combo.currentText())
        self.chat_input.clear()
        self.refresh_all()

    def publish_focus(self) -> None:
        self.save_config()
        self.store.send_message("system", f"{self.config.display_name} updated focus: {self.config.focus_project or 'No project named'}")
        self.refresh_all()

    def send_command(self) -> None:
        command = self.command_input.toPlainText().strip()
        if not command:
            QMessageBox.information(self, "No Command", "Type a PowerShell command first.")
            return
        self.save_config()
        self.store.send_command(
            self.command_target_combo.currentText(),
            self.command_summary_edit.text().strip(),
            command,
            self.command_workdir_edit.text().strip(),
            self.command_project_edit.text().strip(),
            self.command_requires_approval.isChecked(),
        )
        self.command_input.clear()
        self.command_summary_edit.clear()
        self.refresh_all()

    def selected_command(self) -> dict | None:
        row = self.approvals_table.currentRow()
        if row < 0 or row >= len(self.latest_commands):
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
            lines.extend(["", f"Exit code: {result.get('exit_code', '')}", "Stdout:", result.get("stdout", ""), "", "Stderr:", result.get("stderr", "")])
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
            lines.append(f"[{stamp}] {sender} [{kind}]")
            lines.append(item.get("text", ""))
            lines.append("")
        self.chat_log.setPlainText("\n".join(lines).strip())
        self.chat_log.verticalScrollBar().setValue(self.chat_log.verticalScrollBar().maximum())

    def refresh_approvals(self) -> None:
        local = self.config.local_node
        visible = [record for record in self.latest_commands if record.get("target") == local]
        self.approvals_table.setRowCount(len(visible))
        for row, record in enumerate(visible):
            created_item = QTableWidgetItem(record.get("created_at", ""))
            created_item.setData(Qt.UserRole, record.get("id"))
            self.approvals_table.setItem(row, 0, created_item)
            self.approvals_table.setItem(row, 1, QTableWidgetItem(record.get("sender_name") or record.get("sender", "")))
            self.approvals_table.setItem(row, 2, QTableWidgetItem(record.get("summary", "")))
            self.approvals_table.setItem(row, 3, QTableWidgetItem(record.get("status", "")))
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
            self.projects_table.setItem(row, 2, QTableWidgetItem(info.get("last_update", "")))
            self.projects_table.setItem(row, 3, QTableWidgetItem(info.get("last_note", "")))
        self.projects_table.resizeColumnsToContents()

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
            completed = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", record.get("command", "")],
                capture_output=True,
                text=True,
                cwd=record.get("workdir") or None,
                timeout=600,
            )
            record["status"] = "completed" if completed.returncode == 0 else "failed"
            record["result"] = {
                "exit_code": completed.returncode,
                "stdout": completed.stdout[-10000:],
                "stderr": completed.stderr[-10000:],
                "ran_at": now_iso(),
                "executor": self.config.display_name,
            }
            self.store.update_command(record)
            self.store.send_message("system", f"{self.config.display_name} completed command '{record.get('summary', '')}' with exit code {completed.returncode}.")
            if completed.stdout.strip():
                self.append_log("Stdout:\n" + completed.stdout[-4000:])
            if completed.stderr.strip():
                self.append_log("Stderr:\n" + completed.stderr[-4000:])
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
