from __future__ import annotations

import argparse
import json
import uuid
from datetime import datetime
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
CONFIG_FILE = APP_DIR / "codex_bridge_config.json"
DEFAULT_BRIDGE_ROOT = Path.home() / "Syncthing" / "codex-bridge"


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def make_id(prefix: str) -> str:
    return f"{prefix}-{datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:8]}"


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        return {"local_node": "office", "display_name": "Office Codex", "bridge_root": str(DEFAULT_BRIDGE_ROOT)}
    return json.loads(CONFIG_FILE.read_text(encoding="utf-8-sig"))


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    temp_path.replace(path)


def create_task(args: argparse.Namespace) -> None:
    config = load_config()
    bridge_root = Path(str(config.get("bridge_root") or DEFAULT_BRIDGE_ROOT))
    tasks_dir = bridge_root / "tasks"
    task = {
        "id": make_id("task"),
        "title": args.title.strip(),
        "details": args.details.strip(),
        "assigned_to": args.assigned_to,
        "requested_by": str(config.get("local_node") or "office"),
        "requested_by_name": str(config.get("display_name") or "Office Codex"),
        "destination_path": args.destination.strip(),
        "status": "pending",
        "created_at": now_iso(),
        "updated_at": now_iso(),
    }
    filename = f"{task['created_at'].replace(':', '').replace('-', '')}_{task['id']}.json"
    write_json(tasks_dir / filename, task)
    print(f"Created task {task['id']} for {args.assigned_to}: {args.title}")


def main() -> None:
    parser = argparse.ArgumentParser(description="Create background coordination tasks for Codex Bridge.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    create = subparsers.add_parser("create", help="Create a new tracked bridge task.")
    create.add_argument("--title", required=True, help="Short task title.")
    create.add_argument("--details", default="", help="Longer description.")
    create.add_argument("--assigned-to", choices=("home", "office"), required=True, help="Which node should handle the task.")
    create.add_argument("--destination", default="", help="Optional destination path for the work.")
    create.set_defaults(func=create_task)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
