from __future__ import annotations

import json
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


WEB_DIR = Path(__file__).resolve().parent
APP_DIR = WEB_DIR.parent
DATA_FILE = APP_DIR / "claims_data.json"
HOST = "127.0.0.1"
PORT = 8010


def load_claims() -> list[dict]:
    if not DATA_FILE.exists():
        return []
    try:
        payload = json.loads(DATA_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []

    claims = payload.get("claims", {})
    records = []
    for item in claims.values():
        records.append(
            {
                "key": item.get("key", ""),
                "claim_id": item.get("claim_id", ""),
                "title": item.get("title", ""),
                "status": item.get("status", ""),
                "claim_type": item.get("claim_type", ""),
                "customer_name": item.get("customer_name", ""),
                "insurance_company": item.get("insurance_company", ""),
                "claim_number": item.get("claim_number", ""),
                "date_of_loss": item.get("date_of_loss", ""),
                "town": item.get("town", ""),
                "owner_address": item.get("owner_address", ""),
                "vehicle": item.get("vehicle", ""),
                "vin": item.get("vin", ""),
                "shop_name": item.get("shop_name", ""),
                "contact_phone": item.get("contact_phone", ""),
                "contact_email": item.get("contact_email", ""),
                "assignment_claim_notes": item.get("assignment_claim_notes", ""),
                "notes": item.get("notes", ""),
                "office_progress_status": item.get("office_progress_status", ""),
                "office_appt_when": item.get("office_appt_when", ""),
                "office_waiting_for_paperwork": item.get("office_waiting_for_paperwork", ""),
                "updated_at": item.get("updated_at", ""),
                "source_path": item.get("source_path", ""),
            }
        )

    records.sort(key=lambda record: ((record.get("status") or "") != "Open", record.get("claim_id") or record.get("title") or ""))
    return records


def summarize(records: list[dict]) -> dict:
    open_count = sum(1 for record in records if record.get("status") == "Open")
    closed_count = sum(1 for record in records if record.get("status") == "Closed")
    return {
        "all": len(records),
        "open": open_count,
        "closed": closed_count,
    }


class ClaimsWebHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path

        if path in ("/", "/index.html"):
            self._serve_file("index.html", "text/html; charset=utf-8")
            return
        if path == "/app.js":
            self._serve_file("app.js", "application/javascript; charset=utf-8")
            return
        if path == "/styles.css":
            self._serve_file("styles.css", "text/css; charset=utf-8")
            return
        if path == "/api/claims":
            self._serve_claims_api(parsed.query)
            return

        self.send_error(404, "Not Found")

    def log_message(self, format: str, *args) -> None:
        return

    def _serve_claims_api(self, query: str) -> None:
        params = parse_qs(query)
        status_filter = (params.get("status", ["all"])[0] or "all").lower()
        search = (params.get("search", [""])[0] or "").strip().lower()

        records = load_claims()
        if status_filter in {"open", "closed"}:
            records = [record for record in records if (record.get("status") or "").lower() == status_filter]

        if search:
            def matches(record: dict) -> bool:
                haystack = " ".join(
                    [
                        record.get("claim_id", ""),
                        record.get("customer_name", ""),
                        record.get("insurance_company", ""),
                        record.get("claim_number", ""),
                        record.get("town", ""),
                        record.get("vehicle", ""),
                        record.get("shop_name", ""),
                    ]
                ).lower()
                return search in haystack

            records = [record for record in records if matches(record)]

        payload = {
            "summary": summarize(load_claims()),
            "records": records,
        }
        body = json.dumps(payload).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _serve_file(self, name: str, content_type: str) -> None:
        path = WEB_DIR / name
        if not path.exists():
            self.send_error(404, "Not Found")
            return
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), ClaimsWebHandler)
    print(f"Claims web app running at http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
