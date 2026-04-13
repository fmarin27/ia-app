from __future__ import annotations

import json
import os
import time
import base64
import subprocess
import urllib.request
import uuid
from datetime import datetime
from pathlib import Path

CONFIG_PATH = Path(__file__).resolve().parent / "codex_bridge_config.json"
DEFAULT_BRIDGE_ROOT = Path.home() / "Syncthing" / "codex-bridge"
SECRETS_FILE = Path.home() / ".codex_bridge_secrets.json"
POLL_SECONDS = 3
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


def powershell_quote(text: str) -> str:
    return "'" + text.replace("'", "''") + "'"


def encoded_powershell(command_text: str) -> str:
    return base64.b64encode(command_text.encode("utf-16le")).decode("ascii")


def direct_bridge_target(local_node: str, target: str) -> str | None:
    local_node = (local_node or "home").strip().lower()
    target = (target or "all").strip().lower()
    if target in {"home", "office"} and target != local_node:
        return target
    if target == "all":
        return "office" if local_node == "home" else "home"
    return None


def validate_direct_access_profile(config: dict, target: str) -> tuple[str, str, str] | None:
    prefix = "home" if target == "home" else "office"
    host = str(config.get(f"{prefix}_host") or "").strip()
    user = str(config.get(f"{prefix}_user") or "").strip()
    key = str(config.get(f"{prefix}_key") or "").strip()
    if not host or not user or not key:
        return None
    return host, user, key


def deliver_bridge_payload_direct(config: dict, bridge_root: Path, target: str, subdir: str, filename: str, payload: dict) -> None:
    profile = validate_direct_access_profile(config, target)
    if not profile:
        return
    host, user, key = profile
    remote_dir = str((bridge_root / subdir)).replace("/", "\\")
    payload_bytes = json.dumps(payload, indent=2).encode("utf-8")
    payload_b64 = base64.b64encode(payload_bytes).decode("ascii")
    script = (
        f"$dir = {powershell_quote(remote_dir)}; "
        "New-Item -ItemType Directory -Force -Path $dir | Out-Null; "
        f"$dest = Join-Path $dir {powershell_quote(filename)}; "
        f"$bytes = [Convert]::FromBase64String('{payload_b64}'); "
        "[System.IO.File]::WriteAllBytes($dest, $bytes)"
    )
    encoded = encoded_powershell(script)
    completed = subprocess.run(
        [
            r"C:\Windows\System32\OpenSSH\ssh.exe",
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
    if completed.returncode != 0:
        raise RuntimeError(completed.stderr.strip() or completed.stdout.strip() or f"SSH failed with exit code {completed.returncode}")


def load_config() -> dict:
    if not CONFIG_PATH.exists():
        return {
            "local_node": "home",
            "display_name": "Home Codex",
            "bridge_root": str(DEFAULT_BRIDGE_ROOT),
            "auto_reply_enabled": False,
            "auto_reply_model": "gpt-4.1-mini",
        }
    return read_json(CONFIG_PATH)


def load_secrets() -> dict:
    if not SECRETS_FILE.exists():
        return {}
    return read_json(SECRETS_FILE)


def current_api_key() -> str:
    return os.environ.get("OPENAI_API_KEY", "").strip() or str(load_secrets().get("openai_api_key", "")).strip()


def read_json_files(directory: Path) -> list[dict]:
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


def recent_chat_context(messages: list[dict], limit: int = 12) -> str:
    lines: list[str] = []
    for item in messages[-limit:]:
        if item.get("kind") != "chat":
            continue
        sender = item.get("sender_name") or item.get("sender") or "unknown"
        lines.append(f"{sender}: {item.get('text', '')}")
    return "\n".join(lines).strip()


def request_auto_reply(config: dict, messages: list[dict], message: dict) -> str:
    api_key = current_api_key()
    if not api_key:
        raise RuntimeError("No OpenAI API key configured.")

    local_node = str(config.get("local_node") or "home").strip() or "home"
    display_name = str(config.get("display_name") or ("Home Codex" if local_node == "home" else "Office Codex")).strip()
    model = str(config.get("auto_reply_model") or "gpt-4.1-mini").strip() or "gpt-4.1-mini"
    sender_name = message.get("sender_name") or message.get("sender") or "Remote User"
    sender_node = message.get("sender") or ("office" if local_node == "home" else "home")
    system_prompt = (
        f"You are {display_name} running on the {local_node} PC through Codex Bridge. "
        "Reply helpfully and concisely to the other PC. "
        "Do not claim you executed commands or changed files unless that actually happened. "
        "If you need the user to target a different node or approve a command, say so plainly."
    )
    user_prompt = (
        f"Recent bridge chat:\n{recent_chat_context(messages) or '(no recent chat)'}\n\n"
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


def main() -> None:
    config = load_config()
    bridge_root = Path(str(config.get("bridge_root") or DEFAULT_BRIDGE_ROOT))
    messages_dir = bridge_root / "messages"
    state_dir = bridge_root / "state"
    state_dir.mkdir(parents=True, exist_ok=True)
    responder_state_path = state_dir / f"{str(config.get('local_node') or 'home')}_auto_reply.json"
    responder_runtime_path = state_dir / f"{str(config.get('local_node') or 'home')}_auto_reply_state.json"

    state = {"initialized": False, "handled_ids": [], "updated_at": now_iso()}
    if responder_runtime_path.exists():
        try:
            state = read_json(responder_runtime_path)
        except Exception:
            pass

    while True:
        config = load_config()
        enabled = bool(config.get("auto_reply_enabled"))
        messages = sorted(read_json_files(messages_dir), key=lambda item: item.get("created_at", ""))
        local_node = str(config.get("local_node") or "home").strip() or "home"

        heartbeat = {
            "node": local_node,
            "display_name": str(config.get("display_name") or local_node),
            "auto_reply_enabled": enabled,
            "auto_reply_model": str(config.get("auto_reply_model") or "gpt-4.1-mini"),
            "last_checked_at": now_iso(),
            "handled_count": len(state.get("handled_ids", [])),
            "status": "idle",
        }

        try:
            incoming = [
                item for item in messages
                if item.get("kind") == "chat"
                and item.get("sender") not in ("", None, local_node)
                and item.get("target") == local_node
            ]

            if not state.get("initialized"):
                state["initialized"] = True
                state["handled_ids"] = sorted({item.get("id") for item in incoming if item.get("id")})
                state["updated_at"] = now_iso()
                write_json(responder_runtime_path, state)
                heartbeat["status"] = "initialized"
                write_json(responder_state_path, heartbeat)
                time.sleep(POLL_SECONDS)
                continue

            handled_ids = set(state.get("handled_ids", []))
            pending = [item for item in incoming if item.get("id") and item.get("id") not in handled_ids]

            if enabled and pending:
                message = pending[0]
                heartbeat["status"] = "replying"
                heartbeat["replying_to"] = message.get("id")
                write_json(responder_state_path, heartbeat)

                reply_text = request_auto_reply(config, messages, message)
                if not reply_text:
                    raise RuntimeError("Model returned an empty reply.")

                response = {
                    "id": make_id("msg"),
                    "kind": "chat",
                    "sender": local_node,
                    "sender_name": str(config.get("display_name") or local_node),
                    "target": message.get("sender") or ("office" if local_node == "home" else "home"),
                    "text": reply_text,
                    "reply_to": message.get("id"),
                    "created_at": now_iso(),
                }
                filename = f"{response['created_at'].replace(':', '').replace('-', '')}_{response['id']}.json"
                write_json(messages_dir / filename, response)
                remote_target = direct_bridge_target(local_node, str(response.get("target") or "all"))
                if remote_target:
                    deliver_bridge_payload_direct(config, bridge_root, remote_target, "messages", filename, response)

                handled_ids.add(message["id"])
                state["handled_ids"] = sorted(handled_ids)
                state["updated_at"] = now_iso()
                write_json(responder_runtime_path, state)
                heartbeat["status"] = "replied"
                heartbeat["last_replied_at"] = now_iso()
                heartbeat["last_reply_to"] = message.get("id")
            else:
                heartbeat["status"] = "idle"

        except Exception as exc:
            heartbeat["status"] = "error"
            heartbeat["last_error"] = str(exc)

        write_json(responder_state_path, heartbeat)
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
