from __future__ import annotations

import json
import subprocess
import tempfile
import time
import uuid
from datetime import datetime
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "codex_bridge_config.json"
DEFAULT_BRIDGE_ROOT = Path.home() / "Syncthing" / "codex-bridge"
POLL_SECONDS = 15
STATUS_NUDGE_SECONDS = 120
CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def make_id(prefix: str) -> str:
    return f"{prefix}-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:8]}"


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    temp_path.replace(path)


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        return {
            "local_node": "office",
            "display_name": "Office Codex",
            "bridge_root": str(DEFAULT_BRIDGE_ROOT),
        }
    return read_json(CONFIG_FILE)


def load_json_files(directory: Path) -> list[dict]:
    items: list[dict] = []
    if not directory.exists():
        return items
    for path in sorted(directory.glob("*.json")):
        try:
            payload = read_json(path)
            payload["_path"] = str(path)
            items.append(payload)
        except Exception:
            continue
    return items


def summarize_result_for_chat(record: dict) -> str:
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


def parse_direct_command_from_chat(text: str) -> tuple[str, str] | None:
    stripped = text.strip()
    lower = stripped.lower()
    for prefix in ("run:", "ps:", "powershell:", "cmd:"):
        if lower.startswith(prefix):
            command_text = stripped[len(prefix):].strip()
            if command_text:
                summary = command_text.splitlines()[0][:120]
                return command_text, summary
    return None


def parse_action_from_chat(text: str) -> tuple[str, str, str] | None:
    stripped = text.strip()
    lower = stripped.lower()
    if any(phrase in lower for phrase in ("print a test page", "print test page", "print the test page")):
        return ("print_test_page", "", "Print a test page")
    if any(phrase in lower for phrase in ("what apps are running", "what programs are running", "list running apps", "list running processes", "show running processes")):
        return ("shell_command", "Get-Process | Sort-Object ProcessName | Select-Object -First 120 ProcessName,Id,MainWindowTitle | Format-Table -AutoSize", "List running processes")
    direct = parse_direct_command_from_chat(text)
    if direct:
        command_text, summary = direct
        return ("shell_command", command_text, summary)
    return None


def trusted_chat_state_path(state_dir: Path, local_node: str) -> Path:
    return state_dir / f"{local_node}_trusted_chat_state.json"


def load_state_file(path: Path, default: dict) -> dict:
    if not path.exists():
        return dict(default)
    try:
        return read_json(path)
    except Exception:
        return dict(default)


def create_command_record(commands_dir: Path, message: dict, action_kind: str, command_text: str, summary: str, local_node: str) -> dict:
    sender = str(message.get("sender") or "").strip()
    sender_name = str(message.get("sender_name") or sender or "Remote Sender").strip()
    command_id = make_id("cmd")
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
        "approval_required": False,
        "status": "approved",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "approved_by": sender_name,
        "source_message_id": str(message.get("id") or ""),
        "result": {},
    }
    filename = f"{payload['created_at'].replace(':', '').replace('-', '')}_{command_id}.json"
    write_json(commands_dir / filename, payload)
    return payload


def print_test_page() -> dict:
    temp_dir = Path(tempfile.gettempdir()) / "codex_bridge"
    temp_dir.mkdir(parents=True, exist_ok=True)
    file_path = temp_dir / f"codex_bridge_test_page_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    body = (
        "Codex Bridge Test Page\r\n"
        f"Printed at: {now_iso()}\r\n"
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


def execute_command_record(record: dict, path: Path, local_node: str, display_name: str, messages_dir: Path, other_node: str) -> None:
    record["status"] = "running"
    record["updated_at"] = now_iso()
    write_json(path, {k: v for k, v in record.items() if not k.startswith("_")})
    action = str(record.get("action") or record.get("kind") or "shell_command").strip()
    try:
        if action == "print_test_page":
            result = print_test_page()
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
        exit_code = result.get("exit_code", 1)
        record["status"] = "completed" if exit_code == 0 else "failed"
        record["updated_at"] = now_iso()
        record["result"] = {
            **result,
            "ran_at": now_iso(),
            "action": action,
            "executor": display_name,
        }
        write_json(path, {k: v for k, v in record.items() if not k.startswith("_")})
        append_message(messages_dir, local_node, display_name, other_node, "system", f"{display_name} completed command '{record.get('summary', '')}' with exit code {exit_code}.")
        append_message(messages_dir, local_node, display_name, record.get("sender") or other_node, "chat", summarize_result_for_chat(record), str(record.get("source_message_id") or ""))
    except subprocess.TimeoutExpired:
        record["status"] = "failed"
        record["updated_at"] = now_iso()
        record["result"] = {
            "exit_code": "timeout",
            "stdout": "",
            "stderr": "Command timed out after 600 seconds.",
            "ran_at": now_iso(),
            "action": action,
            "executor": display_name,
        }
        write_json(path, {k: v for k, v in record.items() if not k.startswith("_")})
        append_message(messages_dir, local_node, display_name, other_node, "system", f"Command timed out: {record.get('summary', '')}")
    except Exception as exc:
        record["status"] = "failed"
        record["updated_at"] = now_iso()
        record["result"] = {
            "exit_code": "error",
            "stdout": "",
            "stderr": str(exc),
            "ran_at": now_iso(),
            "action": action,
            "executor": display_name,
        }
        write_json(path, {k: v for k, v in record.items() if not k.startswith("_")})
        append_message(messages_dir, local_node, display_name, other_node, "system", f"Command failed to start: {record.get('summary', '')}")


def append_message(messages_dir: Path, sender: str, sender_name: str, target: str, kind: str, text: str, reply_to: str = "") -> None:
    payload = {
        "id": make_id("msg"),
        "kind": kind,
        "sender": sender,
        "sender_name": sender_name,
        "target": target,
        "text": text,
        "created_at": now_iso(),
    }
    if reply_to:
        payload["reply_to"] = reply_to
    filename = f"{payload['created_at'].replace(':', '').replace('-', '')}_{payload['id']}.json"
    write_json(messages_dir / filename, payload)


def parse_iso(value: str) -> datetime | None:
    try:
        return datetime.fromisoformat(value)
    except Exception:
        return None


def seconds_since(value: str) -> float | None:
    stamp = parse_iso(value)
    if not stamp:
        return None
    return max(0.0, (datetime.now() - stamp).total_seconds())


def main() -> None:
    config = load_config()
    bridge_root = Path(str(config.get("bridge_root") or DEFAULT_BRIDGE_ROOT))
    messages_dir = bridge_root / "messages"
    commands_dir = bridge_root / "commands"
    tasks_dir = bridge_root / "tasks"
    state_dir = bridge_root / "state"
    messages_dir.mkdir(parents=True, exist_ok=True)
    commands_dir.mkdir(parents=True, exist_ok=True)
    tasks_dir.mkdir(parents=True, exist_ok=True)
    state_dir.mkdir(parents=True, exist_ok=True)

    local_node = str(config.get("local_node") or "office").strip() or "office"
    display_name = str(config.get("display_name") or ("Office Codex" if local_node == "office" else "Home Codex")).strip()
    other_node = "home" if local_node == "office" else "office"
    runtime_path = state_dir / f"{local_node}_coordinator.json"

    if runtime_path.exists():
        try:
            runtime = read_json(runtime_path)
        except Exception:
            runtime = {}
    else:
        runtime = {}

    runtime.setdefault("announced_task_ids", [])
    runtime.setdefault("nudged_task_ids", {})
    trusted_state_default = {"initialized": False, "handled_ids": [], "updated_at": now_iso()}

    while True:
        config = load_config()
        local_node = str(config.get("local_node") or local_node).strip() or local_node
        display_name = str(config.get("display_name") or display_name).strip() or display_name
        other_node = "home" if local_node == "office" else "office"
        trusted_state_path = trusted_chat_state_path(state_dir, local_node)
        trusted_state = load_state_file(trusted_state_path, trusted_state_default)

        messages = load_json_files(messages_dir)
        incoming_messages = [
            item for item in messages
            if item.get("kind") == "chat"
            and item.get("sender") not in ("", None, local_node)
            and item.get("target") == local_node
            and not bool(item.get("routed_action"))
        ]
        handled_ids = set(trusted_state.get("handled_ids", []))
        if not trusted_state.get("initialized"):
            trusted_state["initialized"] = True
            trusted_state["handled_ids"] = sorted({item.get("id") for item in incoming_messages if item.get("id")})
            trusted_state["updated_at"] = now_iso()
            write_json(trusted_state_path, trusted_state)
        else:
            for message in incoming_messages:
                message_id = str(message.get("id") or "")
                if not message_id or message_id in handled_ids:
                    continue
                parsed = parse_action_from_chat(str(message.get("text") or ""))
                if parsed is not None:
                    action_kind, command_text, summary = parsed
                    create_command_record(commands_dir, message, action_kind, command_text, summary, local_node)
                    append_message(messages_dir, local_node, display_name, message.get("sender") or other_node, "system", f"Trusted bridge command accepted from {message.get('sender_name') or message.get('sender')}: {summary}")
                handled_ids.add(message_id)
            trusted_state["initialized"] = True
            trusted_state["handled_ids"] = sorted(handled_ids)
            trusted_state["updated_at"] = now_iso()
            write_json(trusted_state_path, trusted_state)

        commands = load_json_files(commands_dir)
        for command in commands:
            command_id = str(command.get("id") or "")
            if command.get("target") != local_node:
                continue
            if command.get("status") != "approved":
                continue
            if not command_id:
                continue
            execute_command_record(command, Path(command["_path"]), local_node, display_name, messages_dir, other_node)

        tasks = load_json_files(tasks_dir)
        announced = set(runtime.get("announced_task_ids", []))
        nudged: dict[str, str] = dict(runtime.get("nudged_task_ids", {}))
        active_titles: list[str] = []

        for task in tasks:
            task_id = str(task.get("id") or "")
            assigned_to = str(task.get("assigned_to") or "").strip()
            status = str(task.get("status") or "pending").strip()
            title = str(task.get("title") or task_id or "Untitled task").strip()
            if assigned_to != local_node or not task_id:
                continue

            active_titles.append(title)

            if task_id not in announced:
                if status == "pending":
                    task["status"] = "active"
                    task["claimed_by"] = local_node
                    task["claimed_by_name"] = display_name
                    task["updated_at"] = now_iso()
                    write_json(Path(task["_path"]), {k: v for k, v in task.items() if not k.startswith("_")})
                    status = "active"
                append_message(
                    messages_dir,
                    sender=local_node,
                    sender_name=display_name,
                    target=other_node,
                    kind="system",
                    text=f"{display_name} acknowledged task '{title}' and is tracking it in the background.",
                )
                announced.add(task_id)
                runtime["announced_task_ids"] = sorted(announced)

            if status == "active":
                age = seconds_since(str(task.get("updated_at") or task.get("created_at") or ""))
                last_nudge_age = seconds_since(str(nudged.get(task_id) or ""))
                if age is not None and age >= STATUS_NUDGE_SECONDS and (last_nudge_age is None or last_nudge_age >= STATUS_NUDGE_SECONDS):
                    append_message(
                        messages_dir,
                        sender=local_node,
                        sender_name=display_name,
                        target=other_node,
                        kind="system",
                        text=f"{display_name} is still working on '{title}'.",
                    )
                    nudged[task_id] = now_iso()

        runtime["nudged_task_ids"] = nudged
        runtime["updated_at"] = now_iso()
        runtime["active_tasks"] = active_titles
        write_json(runtime_path, runtime)
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
