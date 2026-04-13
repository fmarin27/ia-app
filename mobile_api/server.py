from __future__ import annotations

import json
import mimetypes
import os
import sys
import threading
from dataclasses import asdict
from datetime import datetime
from email.parser import BytesParser
from email.policy import default as email_policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

API_DIR = Path(__file__).resolve().parent
APP_DIR = API_DIR.parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from claim_manager_3.data_access import ClaimsRepository

try:
    from pypdf import PdfReader
except Exception:
    PdfReader = None
HOST = "0.0.0.0"
PORT = int(os.environ.get("CLAIM_MANAGER_MOBILE_API_PORT", "8011"))
IGNORED_FILE_NAMES = {"desktop.ini", "thumbs.db"}
PHOTO_EXTENSIONS = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".heic": "image/heic",
    ".webp": "image/webp",
}
ASSIGNMENT_MEASUREMENT_CACHE: dict[str, tuple[float, int, bool, str]] = {}
MEASUREMENT_SCAN_STATUS = {
    "running": False,
    "started_at": "",
    "completed_at": "",
    "claims_scanned": 0,
    "claims_with_assignment": 0,
    "claims_flagged": 0,
    "last_error": "",
}
MEASUREMENT_SCAN_LOCK = threading.Lock()


def iso_timestamp(path: Path) -> str:
    try:
        return datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds")
    except OSError:
        return ""


def clean_text(value: object) -> str:
    return str(value or "").strip()


def parse_multipart_file(body: bytes, content_type: str, field_name: str) -> tuple[str, str, bytes] | None:
    if "multipart/form-data" not in content_type.lower():
        return None

    header_block = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
    message = BytesParser(policy=email_policy).parsebytes(header_block + body)
    if not message.is_multipart():
        return None

    for part in message.iter_parts():
        name = clean_text(part.get_param("name", header="content-disposition"))
        if name != field_name:
            continue
        filename = clean_text(part.get_filename()) or "photo.jpg"
        payload = part.get_payload(decode=True) or b""
        if not payload:
            return None
        return filename, part.get_content_type() or "", payload
    return None


def placeholder_safe(value: object) -> str:
    text = clean_text(value)
    return "" if text.lower() in {"", "-", "n/a", "unknown", "none", "no shop chosen"} else text


def claim_folder_for(claim) -> Path | None:
    source_path = clean_text(getattr(claim, "source_path", ""))
    if not source_path or source_path == "Manual Entry":
        return None
    folder = Path(source_path)
    if folder.exists() and folder.is_dir():
        return folder
    return None


def find_assignment_pdf(claim) -> Path | None:
    candidates: list[Path] = []
    assign_pdf_path = clean_text(getattr(claim, "assign_pdf_path", ""))
    if assign_pdf_path:
        candidates.append(Path(assign_pdf_path))
    folder = claim_folder_for(claim)
    if folder:
        candidates.append(folder / "assign.pdf")
        candidates.extend(sorted(folder.glob("assign*.pdf"), key=lambda path: path.name.lower()))
        candidates.extend(sorted(folder.glob("*.pdf"), key=lambda path: path.name.lower()))

    seen: set[str] = set()
    for candidate in candidates:
        candidate_key = str(candidate).lower()
        if candidate_key in seen:
            continue
        seen.add(candidate_key)
        if not candidate.exists() or not candidate.is_file() or candidate.suffix.lower() != ".pdf":
            continue
        lower_name = candidate.name.lower()
        if lower_name == "assign.pdf" or "assign" in lower_name or lower_name[:-4].isdigit():
            return candidate
    return None


def _normalize_pdf_text(text: str) -> str:
    return " ".join((text or "").lower().split())


def _measurement_requirement_from_text(text: str) -> tuple[bool, str]:
    normalized = _normalize_pdf_text(text)
    if not normalized:
        return False, ""

    phrase_patterns = [
        "measured photos",
        "measurement photos",
        "measure photos",
        "measured photo",
        "measurement photo",
        "photo rod",
        "measuring rod",
    ]
    for phrase in phrase_patterns:
        if phrase in normalized:
            return True, phrase

    has_measurement_term = any(
        token in normalized
        for token in ["measure", "measured", "measurement", "measuring", "rod", "keesiq", "keysiq"]
    )
    has_photo_term = any(token in normalized for token in ["photo", "photos", "photo requirements", "instructions"])
    if has_measurement_term and has_photo_term:
        return True, "measurement requirement"
    return False, ""


def claim_needs_measurement_photos(claim) -> tuple[bool, str]:
    pdf_path = find_assignment_pdf(claim)
    if not pdf_path or not pdf_path.exists() or PdfReader is None:
        return False, ""

    try:
        stat = pdf_path.stat()
    except OSError:
        return False, ""

    cache_key = str(pdf_path)
    cached = ASSIGNMENT_MEASUREMENT_CACHE.get(cache_key)
    if cached and cached[0] == stat.st_mtime and cached[1] == stat.st_size:
        return cached[2], cached[3]

    needs_measurement = False
    matched_phrase = ""
    try:
        reader = PdfReader(str(pdf_path))
        extracted = "\n".join((page.extract_text() or "") for page in reader.pages[:3])
        needs_measurement, matched_phrase = _measurement_requirement_from_text(extracted)
    except Exception:
        needs_measurement, matched_phrase = False, ""

    ASSIGNMENT_MEASUREMENT_CACHE[cache_key] = (stat.st_mtime, stat.st_size, needs_measurement, matched_phrase)
    return needs_measurement, matched_phrase


def measurement_scan_snapshot() -> dict:
    with MEASUREMENT_SCAN_LOCK:
        return dict(MEASUREMENT_SCAN_STATUS)


def _run_measurement_scan(repository: ClaimsRepository, force: bool = False) -> None:
    started_at = datetime.now().isoformat(timespec="seconds")
    with MEASUREMENT_SCAN_LOCK:
        MEASUREMENT_SCAN_STATUS.update(
            {
                "running": True,
                "started_at": started_at,
                "completed_at": "",
                "claims_scanned": 0,
                "claims_with_assignment": 0,
                "claims_flagged": 0,
                "last_error": "",
            }
        )
    try:
        claims = repository.load_claims()
        scanned = 0
        with_assignment = 0
        flagged = 0
        if force:
            ASSIGNMENT_MEASUREMENT_CACHE.clear()
        for claim in claims:
            scanned += 1
            pdf_path = find_assignment_pdf(claim)
            if pdf_path:
                with_assignment += 1
            needs_measurement, _ = claim_needs_measurement_photos(claim)
            if needs_measurement:
                flagged += 1
        with MEASUREMENT_SCAN_LOCK:
            MEASUREMENT_SCAN_STATUS.update(
                {
                    "running": False,
                    "completed_at": datetime.now().isoformat(timespec="seconds"),
                    "claims_scanned": scanned,
                    "claims_with_assignment": with_assignment,
                    "claims_flagged": flagged,
                }
            )
    except Exception as exc:
        with MEASUREMENT_SCAN_LOCK:
            MEASUREMENT_SCAN_STATUS.update(
                {
                    "running": False,
                    "completed_at": datetime.now().isoformat(timespec="seconds"),
                    "last_error": str(exc),
                }
            )


def start_measurement_scan(repository: ClaimsRepository, force: bool = False) -> bool:
    with MEASUREMENT_SCAN_LOCK:
        if MEASUREMENT_SCAN_STATUS.get("running"):
            return False
    worker = threading.Thread(
        target=_run_measurement_scan,
        args=(repository, force),
        daemon=True,
        name="measurement-scan",
    )
    worker.start()
    return True


def route_display_address(claim) -> str:
    for candidate in (
        getattr(claim, "route_address_override", ""),
        getattr(claim, "location_of_vehicle", ""),
        getattr(claim, "shop_name", ""),
        getattr(claim, "owner_address", ""),
    ):
        text = placeholder_safe(candidate)
        if text:
            return text
    return "Address unavailable"


def summarize(records: list) -> dict:
    return {
        "all": len(records),
        "open": sum(1 for record in records if clean_text(record.status).lower() == "open"),
        "closed": sum(1 for record in records if clean_text(record.status).lower() == "closed"),
    }


class MobileApiHandler(BaseHTTPRequestHandler):
    server_version = "ClaimsMobileAPI/2.0"
    repository = ClaimsRepository()

    def do_OPTIONS(self) -> None:
        self.send_response(204)
        self._send_cors_headers()
        self.end_headers()

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/api/mobile/health":
            self._send_json({"ok": True})
            return
        if path == "/api/mobile/dashboard":
            self._serve_dashboard()
            return
        if path == "/api/mobile/claims":
            self._serve_claims(parsed.query)
            return
        if path == "/api/mobile/claim":
            self._serve_claim_detail(parsed.query)
            return
        if path == "/api/mobile/assignment":
            self._serve_assignment(parsed.query)
            return
        if path == "/api/mobile/files":
            self._serve_claim_files(parsed.query)
            return
        if path == "/api/mobile/file":
            self._serve_file(parsed.query)
            return
        if path == "/api/mobile/office-update":
            self._serve_office_update(parsed.query)
            return
        if path == "/api/mobile/route-planner":
            self._serve_route_planner(parsed.query)
            return
        if path == "/api/mobile/measurement-scan":
            self._serve_measurement_scan(parsed.query)
            return
        if path == "/api/mobile/route-plan":
            self._save_route_plan_from_query(parsed.query)
            return
        if path == "/api/mobile/claim-tools":
            self._serve_claim_tools()
            return
        if path == "/api/mobile/body-shops":
            self._serve_body_shops(parsed.query)
            return
        if path == "/api/mobile/insurance-companies":
            self._serve_insurance_companies(parsed.query)
            return

        self.send_error(404, "Not Found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/api/mobile/claim-note":
            self._append_claim_note()
            return
        if path == "/api/mobile/route-plan":
            self._save_route_plan()
            return
        if path == "/api/mobile/upload-photo":
            self._upload_photo(parsed.query)
            return

        self.send_error(404, "Not Found")

    def log_message(self, format: str, *args) -> None:
        return

    def _load_claims(self) -> list:
        return self.repository.load_claims()

    def _find_claim(self, key: str):
        if not key:
            return None
        return next((claim for claim in self._load_claims() if claim.key == key), None)

    def _serve_dashboard(self) -> None:
        claims = self._load_claims()
        settings = self.repository.load_settings()
        self._send_json(
            {
                "summary": summarize(claims),
                "settings": {
                    "route_home_address": clean_text(settings.get("route_home_address")),
                    "claim_tools_folder": clean_text(settings.get("claim_tools_folder")),
                    "watched_folders": [clean_text(path) for path in settings.get("watched_folders", [])],
                },
                "measurement_scan": measurement_scan_snapshot(),
                "tabs": [
                    "Claims",
                    "Route Planner",
                    "Claim Tools",
                    "Body Shops",
                    "Insurance Companies",
                ],
            }
        )

    def _serve_claims(self, query: str) -> None:
        params = parse_qs(query)
        status_filter = clean_text(params.get("status", ["open"])[0]).lower() or "open"
        search = clean_text(params.get("search", [""])[0]).lower()
        claims = self._load_claims()

        if status_filter in {"open", "closed"}:
            claims = [claim for claim in claims if clean_text(claim.status).lower() == status_filter]

        if search:
            claims = [
                claim
                for claim in claims
                if search
                in " ".join(
                    [
                        claim.claim_id,
                        claim.customer_name,
                        claim.title,
                        claim.insurance_company,
                        claim.claim_number,
                        claim.town,
                        claim.vehicle,
                        claim.shop_name,
                    ]
                ).lower()
            ]

        base_url = self._request_base_url()
        payload = {
            "summary": summarize(self._load_claims()),
            "records": [self._serialize_claim_summary(claim, base_url) for claim in claims],
        }
        self._send_json(payload)

    def _serve_claim_detail(self, query: str) -> None:
        params = parse_qs(query)
        key = unquote(clean_text(params.get("key", [""])[0]))
        claim = self._find_claim(key)
        if not claim:
            self._send_json({"error": "Claim not found."}, status=404)
            return

        self._send_json(self._serialize_claim_detail(claim, self._request_base_url()))

    def _serve_assignment(self, query: str) -> None:
        params = parse_qs(query)
        key = unquote(clean_text(params.get("key", [""])[0]))
        claim = self._find_claim(key)
        if not claim:
            self._send_json({"error": "Claim not found."}, status=404)
            return

        pdf_path = find_assignment_pdf(claim)
        if not pdf_path:
            self._send_json({"error": "Assignment sheet not found."}, status=404)
            return
        self._serve_path(pdf_path, filename=pdf_path.name, content_type="application/pdf")

    def _serve_claim_files(self, query: str) -> None:
        params = parse_qs(query)
        key = unquote(clean_text(params.get("key", [""])[0]))
        claim = self._find_claim(key)
        if not claim:
            self._send_json({"error": "Claim not found."}, status=404)
            return

        self._send_json({"files": self._list_claim_files(claim, self._request_base_url())})

    def _serve_file(self, query: str) -> None:
        params = parse_qs(query)
        key = unquote(clean_text(params.get("key", [""])[0]))
        relative_path = unquote(clean_text(params.get("path", [""])[0]))
        claim = self._find_claim(key)
        if not claim:
            self._send_json({"error": "Claim not found."}, status=404)
            return

        file_path = self._resolve_claim_file(claim, relative_path)
        if not file_path:
            self._send_json({"error": "File not found."}, status=404)
            return

        content_type, _ = mimetypes.guess_type(file_path.name)
        self._serve_path(file_path, filename=file_path.name, content_type=content_type or "application/octet-stream")

    def _serve_office_update(self, query: str) -> None:
        params = parse_qs(query)
        search = clean_text(params.get("search", [""])[0]).lower()
        claims = [claim for claim in self._load_claims() if clean_text(claim.status).lower() == "open"]
        if search:
            claims = [
                claim
                for claim in claims
                if search
                in " ".join(
                    [
                        claim.claim_id,
                        claim.customer_name,
                        claim.insurance_company,
                        claim.shop_name,
                        claim.office_progress_status,
                        claim.office_appt_when,
                        claim.office_additional_notes,
                    ]
                ).lower()
            ]

        records = [
            {
                "key": claim.key,
                "claim_id": claim.claim_id,
                "customer_name": claim.customer_name or claim.title,
                "insurance_company": claim.insurance_company,
                "shop_name": claim.shop_name,
                "status": claim.office_progress_status or claim.status,
                "inspection_when": claim.office_appt_when,
                "waiting_for_paperwork": claim.office_waiting_for_paperwork,
                "additional_notes": claim.office_additional_notes,
                "claim_type": claim.claim_type,
                "assignment_url": f"{self._request_base_url()}/api/mobile/assignment?key={quote(claim.key)}",
            }
            for claim in claims
        ]
        self._send_json({"records": records})

    def _serve_route_planner(self, query: str) -> None:
        params = parse_qs(query)
        search = clean_text(params.get("search", [""])[0]).lower()
        status_filter = clean_text(params.get("status", ["all"])[0]).lower() or "all"
        appointment_only = clean_text(params.get("appointment_only", ["0"])[0]).lower() in {"1", "true", "yes"}
        claims = self._load_claims()
        if status_filter in {"open", "closed"}:
            claims = [claim for claim in claims if clean_text(claim.status).lower() == status_filter]
        if appointment_only:
            claims = [
                claim
                for claim in claims
                if clean_text(claim.office_progress_status).lower() == "appointment scheduled"
            ]
        if search:
            claims = [
                claim
                for claim in claims
                if search
                in " ".join(
                    [
                        claim.claim_id,
                        claim.customer_name,
                        claim.insurance_company,
                        claim.vehicle,
                        claim.shop_name,
                        claim.town,
                        route_display_address(claim),
                    ]
                ).lower()
            ]

        settings = self.repository.load_settings()

        self._send_json(
            {
                "start_from": clean_text(settings.get("route_home_address")),
                "records": [self._route_record(claim) for claim in claims],
                "planned_records": [],
                "measurement_scan": measurement_scan_snapshot(),
            }
        )

    def _serve_measurement_scan(self, query: str) -> None:
        params = parse_qs(query)
        force = clean_text(params.get("force", ["0"])[0]).lower() in {"1", "true", "yes"}
        started = False
        if force:
            started = start_measurement_scan(self.repository, force=True)
        payload = measurement_scan_snapshot()
        payload["started"] = started
        self._send_json(payload)

    def _save_route_plan(self) -> None:
        payload = self._read_json_body()
        if payload is None:
            return

        raw_keys = payload.get("keys", [])
        if not isinstance(raw_keys, list):
            self._send_json({"error": "Route plan keys must be a list."}, status=400)
            return

        self._save_route_plan_keys(raw_keys)

    def _save_route_plan_from_query(self, query: str) -> None:
        params = parse_qs(query)
        raw_keys = params.get("key", [])
        if not raw_keys:
            raw_csv = clean_text(params.get("keys", [""])[0])
            if raw_csv:
                raw_keys = [item for item in raw_csv.split(",") if clean_text(item)]
        self._save_route_plan_keys(raw_keys)

    def _save_route_plan_keys(self, raw_keys: list[object]) -> None:
        available_claims = self._load_claims()
        available_keys = {claim.key for claim in available_claims}
        cleaned_keys: list[str] = []
        seen: set[str] = set()
        for raw_key in raw_keys:
            key = clean_text(raw_key)
            if not key or key in seen or key not in available_keys:
                continue
            seen.add(key)
            cleaned_keys.append(key)

        self.repository.save_route_plan_keys(cleaned_keys)
        self._serve_route_planner("status=all")

    def _serve_claim_tools(self) -> None:
        settings = self.repository.load_settings()
        contacts = [asdict(entry) for entry in self.repository.load_claim_tools_contacts()]
        files = [
            {
                "section": entry.section,
                "label": entry.label,
                "file_name": entry.file_name,
                "notes": entry.notes,
                "file_path": entry.file_path,
            }
            for entry in self.repository.load_claim_tools_files()
        ]
        self._send_json(
            {
                "claim_tools_folder": clean_text(settings.get("claim_tools_folder")),
                "contacts": contacts,
                "files": files,
            }
        )

    def _serve_body_shops(self, query: str) -> None:
        params = parse_qs(query)
        search = clean_text(params.get("search", [""])[0]).lower()
        entries = self.repository.load_body_shop_database()
        if search:
            entries = [
                entry
                for entry in entries
                if search
                in " ".join(
                    [
                        entry.shop_name,
                        entry.contact_name,
                        entry.phone,
                        entry.email,
                        entry.address,
                        entry.notes,
                    ]
                ).lower()
            ]
        self._send_json({"records": [asdict(entry) for entry in entries]})

    def _serve_insurance_companies(self, query: str) -> None:
        params = parse_qs(query)
        search = clean_text(params.get("search", [""])[0]).lower()
        entries = self.repository.load_insurance_company_database()
        if search:
            entries = [
                entry
                for entry in entries
                if search
                in " ".join(
                    [
                        entry.company_name,
                        entry.quick_summary,
                        entry.photo_rules,
                        entry.total_loss_rules,
                        entry.documentation_requirements,
                        entry.notes,
                    ]
                ).lower()
            ]
        self._send_json({"records": [asdict(entry) for entry in entries]})

    def _append_claim_note(self) -> None:
        payload = self._read_json_body()
        if payload is None:
            return

        claim_key = clean_text(payload.get("key"))
        note_text = clean_text(payload.get("text"))
        if not claim_key or not note_text:
            self._send_json({"error": "Missing claim key or note text."}, status=400)
            return

        self.repository.append_claim_note(claim_key, note_text)
        claim = self._find_claim(claim_key)
        self._send_json(
            {
                "ok": True,
                "claim": self._serialize_claim_detail(claim, self._request_base_url()) if claim else None,
            }
        )

    def _upload_photo(self, query: str) -> None:
        params = parse_qs(query)
        claim_key = unquote(clean_text(params.get("key", [""])[0]))
        claim = self._find_claim(claim_key)
        if not claim:
            self._send_json({"error": "Claim not found."}, status=404)
            return

        folder = claim_folder_for(claim)
        if not folder:
            self._send_json({"error": "Claim folder is unavailable."}, status=400)
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            content_length = 0
        body = self.rfile.read(content_length) if content_length else b""
        if not body:
            self._send_json({"error": "No upload payload received."}, status=400)
            return

        parsed_file = parse_multipart_file(body, self.headers.get("Content-Type", ""), "photo")
        if parsed_file is None:
            self._send_json({"error": "Photo file is missing."}, status=400)
            return

        original_name, parsed_content_type, file_bytes = parsed_file
        suffix = Path(original_name).suffix.lower() or ".jpg"
        content_type = clean_text(parsed_content_type) or PHOTO_EXTENSIONS.get(suffix, "image/jpeg")
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        destination = folder / f"mobile-photo-{timestamp}{suffix}"
        counter = 2
        while destination.exists():
            destination = folder / f"mobile-photo-{timestamp}-{counter}{suffix}"
            counter += 1

        with destination.open("wb") as handle:
            handle.write(file_bytes)

        self._send_json(
            {
                "ok": True,
                "file": self._serialize_file_entry(claim, destination, folder, self._request_base_url(), override_type=content_type),
            }
        )

    def _read_json_body(self) -> dict | None:
        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            content_length = 0
        raw_body = self.rfile.read(content_length) if content_length else b""
        try:
            payload = json.loads(raw_body.decode("utf-8") or "{}")
        except json.JSONDecodeError:
            self._send_json({"error": "Invalid JSON."}, status=400)
            return None
        if not isinstance(payload, dict):
            self._send_json({"error": "Invalid JSON payload."}, status=400)
            return None
        return payload

    def _route_record(self, claim) -> dict:
        needs_measurement_photos, measurement_photo_hint = claim_needs_measurement_photos(claim)
        return {
            "key": claim.key,
            "claim_id": claim.claim_id,
            "customer_name": claim.customer_name or claim.title,
            "address": route_display_address(claim),
            "shop_name": claim.shop_name,
            "town": claim.town,
            "vehicle": claim.vehicle,
            "insurance_company": claim.insurance_company,
            "status": claim.status,
            "status_label": claim.office_progress_status or claim.status,
            "needs_measurement_photos": needs_measurement_photos,
            "measurement_photo_hint": measurement_photo_hint,
        }

    def _serialize_claim_summary(self, claim, base_url: str) -> dict:
        needs_measurement_photos, measurement_photo_hint = claim_needs_measurement_photos(claim)
        return {
            "key": claim.key,
            "claim_id": claim.claim_id,
            "title": claim.title,
            "status": claim.status,
            "claim_type": claim.claim_type,
            "customer_name": claim.customer_name or claim.title,
            "insurance_company": claim.insurance_company,
            "claim_number": claim.claim_number,
            "date_of_loss": claim.date_of_loss,
            "town": claim.town,
            "shop_name": claim.shop_name,
            "office_progress_status": claim.office_progress_status,
            "office_appt_when": claim.office_appt_when,
            "vehicle": claim.vehicle,
            "assignment_url": f"{base_url}/api/mobile/assignment?key={quote(claim.key)}",
            "needs_measurement_photos": needs_measurement_photos,
            "measurement_photo_hint": measurement_photo_hint,
        }

    def _serialize_claim_detail(self, claim, base_url: str) -> dict:
        payload = asdict(claim)
        payload["assignment_url"] = f"{base_url}/api/mobile/assignment?key={quote(claim.key)}"
        payload["files"] = self._list_claim_files(claim, base_url)
        payload["route_display_address"] = route_display_address(claim)
        payload["folder_available"] = claim_folder_for(claim) is not None
        return payload

    def _list_claim_files(self, claim, base_url: str) -> list[dict]:
        folder = claim_folder_for(claim)
        if not folder:
            return []
        files: list[dict] = []
        for path in sorted(folder.rglob("*"), key=lambda item: (item.is_file() is False, item.name.lower())):
            if not path.is_file():
                continue
            if path.name.lower() in IGNORED_FILE_NAMES:
                continue
            files.append(self._serialize_file_entry(claim, path, folder, base_url))
        return files

    def _serialize_file_entry(
        self,
        claim,
        file_path: Path,
        root_folder: Path,
        base_url: str,
        override_type: str | None = None,
    ) -> dict:
        relative_path = file_path.relative_to(root_folder).as_posix()
        guessed_type, _ = mimetypes.guess_type(file_path.name)
        return {
            "name": file_path.name,
            "relative_path": relative_path,
            "size": file_path.stat().st_size,
            "modified_at": iso_timestamp(file_path),
            "content_type": override_type or guessed_type or "application/octet-stream",
            "file_url": f"{base_url}/api/mobile/file?key={quote(claim.key)}&path={quote(relative_path)}",
        }

    def _resolve_claim_file(self, claim, relative_path: str) -> Path | None:
        folder = claim_folder_for(claim)
        if not folder or not relative_path:
            return None
        candidate = (folder / relative_path).resolve()
        try:
            candidate.relative_to(folder.resolve())
        except ValueError:
            return None
        if candidate.exists() and candidate.is_file():
            return candidate
        return None

    def _serve_path(self, path: Path, filename: str, content_type: str) -> None:
        body = path.read_bytes()
        self.send_response(200)
        self._send_cors_headers()
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Content-Disposition", f'inline; filename="{filename}"')
        self.end_headers()
        self.wfile.write(body)

    def _request_base_url(self) -> str:
        host = self.headers.get("Host") or f"127.0.0.1:{PORT}"
        return f"http://{host}"

    def _send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self._send_cors_headers()
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_cors_headers(self) -> None:
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")


def main() -> None:
    start_measurement_scan(MobileApiHandler.repository, force=False)
    server = ThreadingHTTPServer((HOST, PORT), MobileApiHandler)
    print(f"Claims mobile API running at http://127.0.0.1:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
