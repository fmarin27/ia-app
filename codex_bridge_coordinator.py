from __future__ import annotations

import json
import time
import uuid
from datetime import datetime
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "codex_bridge_config.json"
DEFAULT_BRIDGE_ROOT = Path.home() / "Syncthing" / "codex-bridge"
POLL_SECONDS = 15
STATUS_NUDGE_SECONDS = 120


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
    tasks_dir = bridge_root / "tasks"
    state_dir = bridge_root / "state"
    messages_dir.mkdir(parents=True, exist_ok=True)
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

    while True:
        config = load_config()
        local_node = str(config.get("local_node") or local_node).strip() or local_node
        display_name = str(config.get("display_name") or display_name).strip() or display_name
        other_node = "home" if local_node == "office" else "office"

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
