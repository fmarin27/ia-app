from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    from pypdf import PdfReader
except Exception:
    try:
        from PyPDF2 import PdfReader  # type: ignore[assignment]
    except Exception:
        PdfReader = None  # type: ignore[assignment]


def _runtime_app_dir() -> Path:
    portable_home = (os.environ.get("CLAIM_MANAGER_HOME") or "").strip()
    if portable_home:
        return Path(portable_home).expanduser().resolve()
    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        for candidate in (exe_dir, exe_dir.parent):
            if (candidate / "claims_data.json").exists() or (candidate / "PENDING CLAIMS").exists():
                return candidate
        return exe_dir
    return Path(__file__).resolve().parents[1]


APP_DIR = _runtime_app_dir()
CLAIM_MANAGER_DIR = APP_DIR / "claim_manager_3" if (APP_DIR / "claim_manager_3").exists() else Path(__file__).resolve().parent
_desktop_root = Path.home() / "Desktop"


def _runtime_data_dir() -> Path:
    portable_home = (os.environ.get("CLAIM_MANAGER_HOME") or "").strip()
    if portable_home:
        return Path(portable_home).expanduser().resolve()
    if not getattr(sys, "frozen", False):
        return APP_DIR

    app_dir_text = str(APP_DIR).lower()
    program_files_markers = ("\\program files", "\\program files (x86)")
    if not any(marker in app_dir_text for marker in program_files_markers):
        return APP_DIR

    local_appdata = Path(os.environ.get("LOCALAPPDATA") or (Path.home() / "AppData" / "Local"))
    return local_appdata / "Claim Manager 3"


DATA_DIR = _runtime_data_dir()
DATA_DIR.mkdir(parents=True, exist_ok=True)
DATA_FILE = DATA_DIR / "claims_data.json"
LEGACY_DATA_FILE = APP_DIR / "claims_data.json"
DESKTOP_MODE = (os.environ.get("CLAIM_MANAGER_DESKTOP_MODE") or "home").strip().lower()
OFFICE_SYNC_ENABLED = DESKTOP_MODE == "office"
HOME_APP_ROOT = Path(os.environ.get("CLAIM_MANAGER_HOME_ROOT") or r"C:\Users\ferna\Desktop\IA APP")

def _preferred_existing_path(*candidates: Path) -> Path:
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


DEFAULT_CLOSED_CLAIMS_FOLDER = _preferred_existing_path(
    _desktop_root / "Closed claims",
    _desktop_root / "Closed Claims",
    _desktop_root / "CLOSED CLAIMS",
    DATA_DIR / "Closed Claims",
    APP_DIR / "Closed Claims",
)
DEFAULT_WATCHED_FOLDERS = [
    str(
        _preferred_existing_path(
            _desktop_root / "Pending claims",
            _desktop_root / "Pending Claims",
            _desktop_root / "PENDING CLAIMS",
            DATA_DIR / "PENDING CLAIMS",
            APP_DIR / "PENDING CLAIMS",
        )
    ),
    str(
        _preferred_existing_path(
            _desktop_root / "Closed claims",
            _desktop_root / "Closed Claims",
            _desktop_root / "CLOSED CLAIMS",
            DATA_DIR / "Closed Claims",
            APP_DIR / "Closed Claims",
        )
    ),
]
DEFAULT_CLAIM_TOOLS_FOLDER = str(
    _preferred_existing_path(
        _desktop_root / "Claim Tools",
        DATA_DIR / "Claim Tools",
        APP_DIR / "Claim Tools",
    )
)
DEFAULT_ROUTE_HOME_ADDRESS = "31 Cummings Ave, Fairfield, CT 06825"
TESSERACT_EXE = Path(r"C:\Program Files\PDF24\tesseract\tesseract.exe")
LOCAL_TESSDATA_DIR = CLAIM_MANAGER_DIR / "ocr" / "tessdata"
TESSDATA_DIR = LOCAL_TESSDATA_DIR if LOCAL_TESSDATA_DIR.exists() else Path(r"C:\Program Files\PDF24\tesseract\tessdata")
GHOSTSCRIPT_EXE = Path(r"C:\Program Files\PDF24\gs\bin\gswin64c.exe")
BRIDGE_CONFIG_FILE = Path(
    os.environ.get("CLAIM_MANAGER_BRIDGE_CONFIG")
    or (Path.home() / "Desktop" / "Apps" / "Business" / "Codex Bridge" / "codex_bridge_config.json")
)
HOME_PC_REMOTE_DATA_PATH = os.environ.get("CLAIM_MANAGER_HOME_REMOTE_DATA_PATH") or f"/{str(HOME_APP_ROOT / 'claims_data.json').replace(chr(92), '/')}"
LOCAL_ONLY_KEYS = {
    "watched_folders",
    "claim_tools_folder",
    "route_home_address",
    "route_plan_keys",
    "processed_apptrak_pdfs",
    "office_update_email_to",
    "office_update_email_cc",
    "office_review_email_to",
    "office_rmc_email_to",
    "office_supplement_email_to",
}
CLAIM_SAVED_STATE_FIELDS = (
    "manual_overrides",
    "notes",
    "note_history",
    "route_address_override",
    "office_progress_status",
    "office_appt_when",
    "office_waiting_for_paperwork",
    "office_additional_notes",
)
HOME_TO_LOCAL_PATH_MAP = (
    (str(HOME_APP_ROOT / "PENDING CLAIMS"), str(APP_DIR / "PENDING CLAIMS")),
    (str(HOME_APP_ROOT / "Closed Claims"), str(APP_DIR / "Closed Claims")),
    (str(HOME_APP_ROOT / "Claim Tools"), str(APP_DIR / "Claim Tools")),
)
LOCAL_TO_HOME_PATH_MAP = tuple((local_root, home_root) for home_root, local_root in HOME_TO_LOCAL_PATH_MAP)


@dataclass
class ClaimView:
    key: str
    claim_id: str
    title: str
    status: str
    source_path: str
    claim_type: str
    total_loss: bool
    updated_at: str
    closed_date: str
    customer_name: str
    insurance_company: str
    claim_number: str
    date_of_loss: str
    town: str
    owner_address: str
    location_of_vehicle: str
    vehicle: str
    vin: str
    shop_name: str
    shop_phone: str
    shop_email: str
    contact_phone: str
    contact_email: str
    notes: str
    note_history: list[dict[str, str]]
    assignment_claim_notes: str
    office_progress_status: str
    office_appt_when: str
    office_waiting_for_paperwork: str
    office_additional_notes: str
    route_address_override: str
    assign_pdf_path: str
    payroll_excluded: bool
    payroll_base_pay: float | None
    payroll_total_loss_pay: float | None
    payroll_note: str

    @property
    def display_customer(self) -> str:
        return self.customer_name or self.title or self.claim_id

    @property
    def assignment_path(self) -> str:
        candidates: list[Path] = []
        if self.assign_pdf_path:
            candidates.append(Path(self.assign_pdf_path))
        folder = Path(self.source_path) if self.source_path else None
        if folder and folder.exists() and folder.is_dir():
            candidates.append(folder / "assign.pdf")
            candidates.extend(sorted(folder.glob("assign*.pdf"), key=lambda item: item.name.lower()))
            candidates.extend(sorted(folder.glob("*.pdf"), key=lambda item: item.name.lower()))
        seen: set[str] = set()
        for candidate in candidates:
            candidate_key = str(candidate).lower()
            if candidate_key in seen:
                continue
            seen.add(candidate_key)
            if not candidate.exists() or not candidate.is_file() or candidate.suffix.lower() != ".pdf":
                continue
            lower_name = candidate.name.lower()
            if lower_name == "assign.pdf" or "assign" in lower_name or re.fullmatch(r"\d+\.pdf", lower_name):
                return str(candidate)
        return ""


@dataclass
class ClaimToolContact:
    name: str
    number: str
    prompt_guide: str
    notes: str


@dataclass
class ClaimToolFile:
    section: str
    label: str
    file_path: str
    notes: str

    @property
    def file_name(self) -> str:
        if not self.file_path:
            return ""
        return Path(self.file_path).name


@dataclass
class BodyShopEntry:
    shop_name: str
    contact_name: str
    phone: str
    email: str
    tax_id: str
    address: str
    body_rate: str
    paint_rate: str
    frame_rate: str
    mechanical_rate: str
    certifications: str
    negotiation_notes: str
    notes: str


@dataclass
class InsuranceCompanyEntry:
    company_name: str
    quick_summary: str
    fatal_errors: str
    photo_rules: str
    estimate_supp_rules: str
    parts_rules: str
    total_loss_rules: str
    tow_rules: str
    supplement_rules: str
    betterment_depreciation_rules: str
    documentation_requirements: str
    rates_and_sales_tax_rules: str
    miscellaneous_rules: str
    contact_information: str
    labor_rates: str
    total_loss_threshold: str
    notes: str


class ClaimsRepository:
    def __init__(self, data_file: Path = DATA_FILE) -> None:
        self.data_file = data_file
        self._bridge_config: dict[str, str] | None = None

    def _coerce_optional_float(self, value: Any) -> float | None:
        if value in (None, ""):
            return None
        try:
            return float(value)
        except (TypeError, ValueError):
            return None

    def _claim_state_aliases(self, claim_key: str, record: dict[str, Any]) -> list[str]:
        aliases: list[str] = []
        if claim_key:
            aliases.append(str(claim_key))

        source_path = str(record.get("source_path") or "").strip()
        if source_path:
            aliases.append(f"path::{source_path.lower()}")

        claim_id = str(record.get("claim_id") or "").strip()
        if claim_id:
            aliases.append(f"id::{claim_id}")

        seen: set[str] = set()
        unique_aliases: list[str] = []
        for alias in aliases:
            if alias not in seen:
                seen.add(alias)
                unique_aliases.append(alias)
        return unique_aliases

    def _saved_claim_state(
        self,
        payload: dict,
        claim_key: str,
        record: dict[str, Any],
    ) -> dict[str, Any]:
        saved_state = payload.get("claim_saved_state", {})
        if not isinstance(saved_state, dict):
            return {}

        saved_values: dict[str, Any] = {}
        for alias in reversed(self._claim_state_aliases(claim_key, record)):
            state_entry = saved_state.get(alias)
            if not isinstance(state_entry, dict):
                continue
            for field_name in CLAIM_SAVED_STATE_FIELDS:
                if field_name not in state_entry:
                    continue
                if field_name == "manual_overrides":
                    state_overrides = state_entry.get("manual_overrides", {})
                    if isinstance(state_overrides, dict):
                        merged_overrides = dict(saved_values.get("manual_overrides") or {})
                        merged_overrides.update(state_overrides)
                        saved_values["manual_overrides"] = merged_overrides
                else:
                    saved_values[field_name] = state_entry[field_name]
        return saved_values

    def _saved_manual_overrides(
        self,
        payload: dict,
        claim_key: str,
        record: dict[str, Any],
    ) -> dict[str, Any]:
        state_values = self._saved_claim_state(payload, claim_key, record)
        manual_overrides = state_values.get("manual_overrides", {})
        return dict(manual_overrides) if isinstance(manual_overrides, dict) else {}

    def _remember_claim_state(
        self,
        payload: dict,
        claim_key: str,
        record: dict[str, Any],
        state_values: dict[str, Any] | None,
    ) -> None:
        if not state_values:
            return
        saved_state = payload.setdefault("claim_saved_state", {})
        if not isinstance(saved_state, dict):
            return

        primary_entry = dict(saved_state.get(claim_key) or {})
        for alias in self._claim_state_aliases(claim_key, record):
            state_entry = dict(saved_state.get(alias) or primary_entry)
            for field_name, value in state_values.items():
                if field_name == "manual_overrides":
                    if not isinstance(value, dict):
                        continue
                    state_overrides = dict(state_entry.get("manual_overrides") or {})
                    state_overrides.update(value)
                    state_entry["manual_overrides"] = state_overrides
                elif field_name in CLAIM_SAVED_STATE_FIELDS:
                    state_entry[field_name] = value
            saved_state[alias] = state_entry

    def _remember_manual_overrides(
        self,
        payload: dict,
        claim_key: str,
        record: dict[str, Any],
        manual_overrides: dict[str, Any] | None,
    ) -> None:
        if not manual_overrides:
            return
        self._remember_claim_state(payload, claim_key, record, {"manual_overrides": manual_overrides})

    def _forget_manual_override_fields(
        self,
        payload: dict,
        claim_key: str,
        record: dict[str, Any],
        field_names: set[str],
    ) -> None:
        if not field_names:
            return
        saved_state = payload.get("claim_saved_state", {})
        if not isinstance(saved_state, dict):
            return
        for alias in self._claim_state_aliases(claim_key, record):
            state_entry = saved_state.get(alias)
            if not isinstance(state_entry, dict):
                continue
            state_overrides = state_entry.get("manual_overrides", {})
            if not isinstance(state_overrides, dict):
                continue
            changed = False
            for field_name in field_names:
                if field_name in state_overrides:
                    state_overrides.pop(field_name, None)
                    changed = True
            if not changed:
                continue
            if state_overrides:
                state_entry["manual_overrides"] = state_overrides
            else:
                state_entry.pop("manual_overrides", None)
            saved_state[alias] = state_entry

    def _load_payload(self) -> dict:
        local_payload = self._load_local_payload()
        if not OFFICE_SYNC_ENABLED:
            return local_payload

        remote_payload = self._fetch_home_payload()
        if not remote_payload:
            return local_payload

        payload = self._translate_payload_paths(remote_payload, HOME_TO_LOCAL_PATH_MAP)
        for key in LOCAL_ONLY_KEYS:
            if key in local_payload:
                payload[key] = local_payload[key]
        return payload

    def _save_payload(self, payload: dict) -> None:
        temp_path = self.data_file.with_suffix(f"{self.data_file.suffix}.tmp")
        temp_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        temp_path.replace(self.data_file)
        if OFFICE_SYNC_ENABLED:
            self._save_home_payload(payload)

    def _load_local_payload(self) -> dict:
        source_path = self.data_file
        if not source_path.exists() and LEGACY_DATA_FILE.exists():
            source_path = LEGACY_DATA_FILE
        if not source_path.exists():
            return {}
        try:
            payload = json.loads(source_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}
        if source_path == LEGACY_DATA_FILE and self.data_file != LEGACY_DATA_FILE:
            try:
                self.data_file.write_text(json.dumps(payload, indent=2), encoding="utf-8")
            except OSError:
                pass
        return payload

    def _load_bridge_config(self) -> dict[str, str]:
        if self._bridge_config is not None:
            return self._bridge_config
        if not OFFICE_SYNC_ENABLED or not BRIDGE_CONFIG_FILE.exists():
            self._bridge_config = {}
            return self._bridge_config
        try:
            raw = json.loads(BRIDGE_CONFIG_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            raw = {}
        self._bridge_config = {
            "home_host": str(raw.get("home_host") or "").strip(),
            "home_user": str(raw.get("home_user") or "").strip(),
            "home_key": str(raw.get("home_key") or "").strip(),
        }
        return self._bridge_config

    def _home_connection_args(self) -> list[str] | None:
        config = self._load_bridge_config()
        host = config.get("home_host", "")
        user = config.get("home_user", "")
        key_path = config.get("home_key", "")
        if not host or not user or not key_path:
            return None
        if not Path(key_path).exists():
            return None
        return ["-i", key_path, f"{user}@{host}"]

    def _fetch_home_payload(self) -> dict:
        connection_args = self._home_connection_args()
        if not connection_args or not HOME_PC_REMOTE_DATA_PATH:
            return {}
        with tempfile.TemporaryDirectory() as temp_dir:
            local_copy = Path(temp_dir) / "claims_data_home.json"
            command = ["scp", *connection_args[:2], f"{connection_args[2]}:{HOME_PC_REMOTE_DATA_PATH}", str(local_copy)]
            completed = subprocess.run(command, capture_output=True, text=True, check=False)
            if completed.returncode != 0 or not local_copy.exists():
                return {}
            try:
                return json.loads(local_copy.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                return {}

    def _save_home_payload(self, payload: dict) -> None:
        connection_args = self._home_connection_args()
        if not connection_args or not HOME_PC_REMOTE_DATA_PATH:
            return

        current_home_payload = self._fetch_home_payload()
        if not current_home_payload:
            return

        translated_payload = self._translate_payload_paths(payload, LOCAL_TO_HOME_PATH_MAP)
        for key, value in translated_payload.items():
            if key in LOCAL_ONLY_KEYS:
                continue
            current_home_payload[key] = value

        with tempfile.TemporaryDirectory() as temp_dir:
            local_copy = Path(temp_dir) / "claims_data_home.json"
            local_copy.write_text(json.dumps(current_home_payload, indent=2), encoding="utf-8")
            command = ["scp", *connection_args[:2], str(local_copy), f"{connection_args[2]}:{HOME_PC_REMOTE_DATA_PATH}"]
            subprocess.run(command, capture_output=True, text=True, check=False)

    def _translate_payload_paths(self, payload: Any, mappings: tuple[tuple[str, str], ...]) -> Any:
        if isinstance(payload, dict):
            translated: dict[Any, Any] = {}
            for key, value in payload.items():
                translated_key = self._translate_claim_key(key, mappings) if isinstance(key, str) else key
                translated[translated_key] = self._translate_payload_paths(value, mappings)
            return translated
        if isinstance(payload, list):
            return [self._translate_payload_paths(item, mappings) for item in payload]
        if isinstance(payload, str):
            return self._translate_string_paths(payload, mappings)
        return payload

    def _translate_claim_key(self, value: str, mappings: tuple[tuple[str, str], ...]) -> str:
        if not value.startswith("fs::"):
            return value
        translated = self._translate_windows_path(value[4:], mappings)
        return f"fs::{translated}"

    def _translate_string_paths(self, value: str, mappings: tuple[tuple[str, str], ...]) -> str:
        translated = self._translate_windows_path(value, mappings)
        if translated != value:
            return translated
        return value

    def _translate_windows_path(self, value: str, mappings: tuple[tuple[str, str], ...]) -> str:
        normalized_value = value.replace("/", "\\")
        lowered_value = normalized_value.lower()
        for source_root, destination_root in mappings:
            normalized_source = source_root.replace("/", "\\")
            normalized_destination = destination_root.replace("/", "\\")
            lowered_source = normalized_source.lower()
            if lowered_value == lowered_source:
                return normalized_destination
            prefix = f"{lowered_source}\\"
            if lowered_value.startswith(prefix):
                suffix = normalized_value[len(normalized_source):]
                return f"{normalized_destination}{suffix}"
        return value

    def load_claims(self) -> list[ClaimView]:
        payload = self._load_payload()
        if not payload:
            return []

        claims_block = payload.get("claims", {})
        results: list[ClaimView] = []
        changed = False
        for key, raw in claims_block.items():
            if not isinstance(raw, dict):
                continue
            shop_name = str(raw.get("shop_name") or "").strip()
            if not shop_name or shop_name == "-":
                shop_name = "No Shop Chosen"
                raw["shop_name"] = shop_name
                changed = True
            results.append(
                ClaimView(
                    key=str(raw.get("key") or key),
                    claim_id=str(raw.get("claim_id") or ""),
                    title=str(raw.get("title") or ""),
                    status=str(raw.get("status") or ""),
                    source_path=str(raw.get("source_path") or ""),
                    claim_type=str(raw.get("claim_type") or ""),
                    total_loss=bool(raw.get("total_loss")),
                    updated_at=str(raw.get("updated_at") or ""),
                    closed_date=str(raw.get("closed_date") or ""),
                    customer_name=str(raw.get("customer_name") or ""),
                    insurance_company=str(raw.get("insurance_company") or ""),
                    claim_number=str(raw.get("claim_number") or ""),
                    date_of_loss=str(raw.get("date_of_loss") or ""),
                    town=str(raw.get("town") or ""),
                    owner_address=str(raw.get("owner_address") or ""),
                    location_of_vehicle=str(raw.get("location_of_vehicle") or ""),
                    vehicle=str(raw.get("vehicle") or ""),
                    vin=str(raw.get("vin") or ""),
                    shop_name=shop_name,
                    shop_phone=str(raw.get("shop_phone") or ""),
                    shop_email=str(raw.get("shop_email") or ""),
                    contact_phone=str(raw.get("contact_phone") or ""),
                    contact_email=str(raw.get("contact_email") or ""),
                    notes=str(raw.get("notes") or ""),
                    note_history=[
                        {"timestamp": str(entry.get("timestamp") or ""), "text": str(entry.get("text") or "")}
                        for entry in (raw.get("note_history") or [])
                        if isinstance(entry, dict)
                    ],
                    assignment_claim_notes=str(raw.get("assignment_claim_notes") or ""),
                    office_progress_status=str(raw.get("office_progress_status") or ""),
                    office_appt_when=str(raw.get("office_appt_when") or ""),
                    office_waiting_for_paperwork=str(raw.get("office_waiting_for_paperwork") or ""),
                    office_additional_notes=str(raw.get("office_additional_notes") or ""),
                    route_address_override=str(raw.get("route_address_override") or ""),
                    assign_pdf_path=str(raw.get("assign_pdf_path") or ""),
                    payroll_excluded=bool(raw.get("payroll_excluded")),
                    payroll_base_pay=self._coerce_optional_float(raw.get("payroll_base_pay")),
                    payroll_total_loss_pay=self._coerce_optional_float(raw.get("payroll_total_loss_pay")),
                    payroll_note=str(raw.get("payroll_note") or ""),
                )
            )
        if changed:
            self._save_payload(payload)
        results.sort(key=lambda claim: (claim.status != "Open", claim.claim_id or claim.title))
        return results

    def refresh_scan(self) -> list[str]:
        payload = self._load_payload()
        watched_folders = [
            str(folder).strip()
            for folder in payload.get("watched_folders", [])
            if str(folder).strip()
        ] or DEFAULT_WATCHED_FOLDERS
        payload["watched_folders"] = watched_folders
        existing_claims = payload.get("claims", {})
        if not isinstance(existing_claims, dict):
            existing_claims = {}

        discovered: dict[str, dict[str, Any]] = {}
        scanned_keys: set[str] = set()
        scanned_roots: list[Path] = []
        access_issues: list[str] = []
        closed_job_numbers = self._closed_folder_job_numbers(watched_folders, access_issues)

        for folder in watched_folders:
            folder_path = Path(folder)
            if not folder_path.exists():
                continue
            default_status = self._default_status_for_folder(folder_path)
            if default_status != "Open":
                continue
            scanned_roots.append(folder_path.resolve())
            self._organize_loose_assignment_pdfs(folder_path, access_issues)
            for item in self._claim_entries_for_folder(folder_path, access_issues):
                key = f"fs::{item.resolve()}"
                scanned_keys.add(key)
                original_key = key
                previous = existing_claims.get(key, {})
                if not isinstance(previous, dict):
                    previous = {}
                record = self._build_scanned_claim_record(item, key, default_status, previous)
                # Reuse cached parsed details when the claim folder/file timestamp has not changed.
                if not self._scan_record_is_unchanged(record, previous):
                    details = self._extract_claim_details(item)
                    if details:
                        record.update(details)
                saved_state_values = self._saved_claim_state(payload, key, record)
                for field_name in CLAIM_SAVED_STATE_FIELDS:
                    if field_name == "manual_overrides":
                        continue
                    if field_name in saved_state_values:
                        record[field_name] = saved_state_values[field_name]
                overrides = dict(saved_state_values.get("manual_overrides") or {})
                previous_overrides = previous.get("manual_overrides", {})
                if isinstance(previous_overrides, dict):
                    overrides.update(previous_overrides)
                if overrides:
                    record["manual_overrides"] = {
                        **dict(record.get("manual_overrides") or {}),
                        **overrides,
                    }
                    for field_name, value in overrides.items():
                        record[field_name] = value
                if default_status == "Open":
                    record["status"] = "Open"
                    record["closed_date"] = ""
                    manual_overrides = dict(record.get("manual_overrides") or {})
                    manual_overrides.pop("status", None)
                    manual_overrides.pop("closed_date", None)
                    record["manual_overrides"] = manual_overrides
                    self._forget_manual_override_fields(payload, key, record, {"status", "closed_date"})
                for name_field in ("customer_name", "insured_name", "claimant_name"):
                    normalized_name = self._normalize_customer_name(str(record.get(name_field) or ""))
                    if normalized_name:
                        record[name_field] = normalized_name
                if default_status == "Open" and item.is_dir():
                    upgraded_item = self._maybe_upgrade_claim_folder_name(
                        item,
                        str(record.get("claim_id") or ""),
                        str(record.get("customer_name") or ""),
                    )
                    if upgraded_item != item:
                        item = upgraded_item
                        key = f"fs::{item.resolve()}"
                        scanned_keys.discard(original_key)
                        scanned_keys.add(key)
                        record["key"] = key
                        record["source_path"] = str(item.resolve())
                        record["title"] = item.name
                        self._rekey_claim_state_entries(payload, original_key, key)
                state_values_to_remember = {
                    field_name: record.get(field_name)
                    for field_name in CLAIM_SAVED_STATE_FIELDS
                    if field_name != "manual_overrides" and field_name in saved_state_values
                }
                manual_overrides = dict(record.get("manual_overrides") or {})
                if manual_overrides:
                    state_values_to_remember["manual_overrides"] = manual_overrides
                self._remember_claim_state(payload, key, record, state_values_to_remember)
                discovered[key] = record

        for key, raw in existing_claims.items():
            if not isinstance(raw, dict):
                continue
            if str(key).startswith("manual::"):
                discovered[str(key)] = dict(raw)
                continue
            if str(key) in discovered:
                continue
            if not self._claim_record_is_under_roots(raw, scanned_roots):
                discovered[str(key)] = dict(raw)

        self._apply_closed_folder_supplement_types(
            discovered,
            eligible_keys=scanned_keys,
            closed_job_numbers=closed_job_numbers,
        )
        payload["claims"] = discovered
        self._save_payload(payload)
        self.sync_reference_data_from_claims(include_claim_folders=False)
        return access_issues

    def _closed_folder_job_numbers(self, watched_folders: list[str], access_issues: list[str]) -> set[str]:
        job_numbers: set[str] = set()
        for folder in watched_folders:
            folder_path = Path(folder)
            if self._default_status_for_folder(folder_path) != "Closed" or not folder_path.exists():
                continue
            try:
                folders = [item for item in folder_path.rglob("*") if item.is_dir()]
            except (PermissionError, OSError):
                access_issues.append(str(folder_path))
                continue
            for item in folders:
                match = re.search(r"\b(\d{8})\b", item.name)
                if match:
                    job_numbers.add(match.group(1))
        return job_numbers

    def _claim_record_is_under_roots(self, record: dict[str, Any], roots: list[Path]) -> bool:
        source_path = str(record.get("source_path") or "").strip()
        if not source_path:
            return False
        try:
            source = Path(source_path).resolve()
        except OSError:
            source = Path(source_path)

        for root in roots:
            try:
                if source == root or source.is_relative_to(root):
                    return True
            except (OSError, ValueError):
                root_text = str(root).rstrip("\\/").lower()
                source_text = str(source).lower()
                if source_text == root_text or source_text.startswith(root_text + "\\"):
                    return True
        return False

    def _apply_closed_folder_supplement_types(
        self,
        claims: dict[str, dict[str, Any]],
        *,
        eligible_keys: set[str],
        closed_job_numbers: set[str],
    ) -> None:
        if not closed_job_numbers:
            return

        for key in eligible_keys:
            record = claims.get(key)
            if not isinstance(record, dict):
                continue
            job_number = str(record.get("claim_id") or "").strip()
            if not job_number or job_number not in closed_job_numbers:
                continue
            manual_overrides = record.get("manual_overrides", {})
            if isinstance(manual_overrides, dict) and "claim_type" in manual_overrides:
                continue
            record["claim_type"] = "Supplement"

    def _scan_record_is_unchanged(self, record: dict[str, Any], previous: dict[str, Any]) -> bool:
        if self._cached_record_needs_reparse(record, previous):
            return False
        return (
            str(previous.get("source_path") or "") == str(record.get("source_path") or "")
            and str(previous.get("updated_at") or "") == str(record.get("updated_at") or "")
        )

    def _cached_record_needs_reparse(self, record: dict[str, Any], previous: dict[str, Any]) -> bool:
        previous_customer = self._normalize_extracted_value(str(previous.get("customer_name") or ""))
        if not previous_customer:
            return True

        previous_insurance = self._normalize_extracted_value(str(previous.get("insurance_company") or ""))
        if previous_insurance and self._looks_suspicious_insurance_company(previous_insurance):
            return True

        if self._looks_suspicious_customer_name(previous_customer):
            return True

        claim_id = self._normalize_extracted_value(str(record.get("claim_id") or ""))
        title = self._normalize_extracted_value(str(record.get("title") or ""))
        source_name = Path(str(record.get("source_path") or "")).name
        source_stem = Path(source_name).stem
        suspicious_values = {
            claim_id,
            title,
            self._normalize_extracted_value(source_name),
            self._normalize_extracted_value(source_stem),
        }
        suspicious_values = {value for value in suspicious_values if value}
        if previous_customer in suspicious_values:
            return True

        if re.fullmatch(r"\d{5,8}", previous_customer):
            return True

        return False

    def _organize_loose_assignment_pdfs(self, folder_path: Path, access_issues: list[str]) -> None:
        try:
            loose_pdfs = [
                item for item in folder_path.iterdir()
                if item.is_file()
                and item.suffix.lower() == ".pdf"
                and item.name.lower() not in {"desktop.ini", "thumbs.db"}
            ]
        except (PermissionError, OSError):
            access_issues.append(str(folder_path))
            return

        for pdf_path in sorted(loose_pdfs, key=lambda item: item.name.lower()):
            lower_name = pdf_path.name.lower()
            if lower_name != "assign.pdf" and "assign" not in lower_name:
                continue
            try:
                details = self._normalize_parsed_details(self._extract_details_from_pdf(pdf_path))
                claim_id = str(details.get("claim_id") or "").strip()
                if not claim_id:
                    text = self._extract_pdf_text(pdf_path, max_pages=3)
                    claim_id = self._extract_claim_id_from_text(text)
                if not claim_id:
                    claim_id = self._claim_id_for_item(pdf_path)
                customer_name = str(details.get("customer_name") or "").strip() or "New Claim"
                target_folder = self._resolve_claim_folder_target(folder_path, claim_id, customer_name)
                target_folder.mkdir(parents=True, exist_ok=True)
                target_file = target_folder / "assign.pdf"
                if target_file.exists():
                    target_folder = self._unique_destination(target_folder)
                    target_folder.mkdir(parents=True, exist_ok=True)
                    target_file = target_folder / "assign.pdf"
                shutil.move(str(pdf_path), str(target_file))
            except (OSError, PermissionError):
                access_issues.append(str(pdf_path))

    def _extract_claim_id_from_text(self, text: str) -> str:
        if not text:
            return ""
        patterns = [
            r"\b(?:file|claim|job|inspection)\s*(?:number|#|id)?\s*[:#]?\s*(2\d{7})\b",
            r"\b(2\d{7})\b",
        ]
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        return ""

    def _resolve_claim_folder_target(self, root: Path, claim_id: str, customer_name: str) -> Path:
        preferred_name = self._preferred_claim_folder_name(claim_id, customer_name)
        existing = []
        try:
            existing = [item for item in root.iterdir() if item.is_dir()]
        except (PermissionError, OSError):
            return root / preferred_name

        if claim_id:
            prefix = f"{claim_id} -".lower()
            for item in existing:
                if item.name.lower() == claim_id.lower() or item.name.lower().startswith(prefix):
                    return self._maybe_upgrade_claim_folder_name(item, claim_id, customer_name)
        return root / preferred_name

    def _preferred_claim_folder_name(self, claim_id: str, customer_name: str) -> str:
        normalized_customer = re.sub(r"[^A-Za-z0-9&'.() ]+", " ", customer_name).strip()
        normalized_customer = (" ".join(normalized_customer.split()) or "New Claim").upper()
        return f"{claim_id} - {normalized_customer}" if claim_id else normalized_customer

    def _maybe_upgrade_claim_folder_name(self, folder: Path, claim_id: str, customer_name: str) -> Path:
        if not claim_id or not folder.exists() or not folder.is_dir():
            return folder

        preferred_name = self._preferred_claim_folder_name(claim_id, customer_name)
        current_name = folder.name.strip()
        if not preferred_name or current_name == preferred_name:
            return folder

        rename_allowed = current_name.lower() in {
            "new folder",
            "new claim",
            claim_id.lower(),
            f"{claim_id} - new claim".lower(),
            f"{claim_id} - {claim_id}".lower(),
        } or current_name.lower() == preferred_name.lower() or bool(re.fullmatch(r"new folder \(\d+\)", current_name.lower()))
        if not rename_allowed:
            return folder

        target = folder.with_name(preferred_name)
        if current_name.lower() == preferred_name.lower():
            temp_target = folder.with_name(f"{preferred_name}.__casefix__")
            if temp_target.exists():
                return folder
            try:
                folder.rename(temp_target)
                temp_target.rename(target)
            except (PermissionError, OSError):
                try:
                    if temp_target.exists():
                        temp_target.rename(folder)
                except (PermissionError, OSError):
                    pass
                return folder
            return target

        if target.exists():
            return folder

        try:
            folder.rename(target)
        except (PermissionError, OSError):
            return folder
        return target

    def _build_scanned_claim_record(
        self,
        item: Path,
        key: str,
        default_status: str,
        previous: dict[str, Any],
    ) -> dict[str, Any]:
        return {
            "key": key,
            "claim_id": self._claim_id_for_item(item),
            "title": item.name,
            "status": str(previous.get("status") or default_status),
            "source_path": str(item.resolve()),
            "claim_type": str(previous.get("claim_type") or self._guess_claim_type_from_name(item.name) or "Original"),
            "total_loss": bool(previous.get("total_loss")),
            "updated_at": datetime.fromtimestamp(item.stat().st_mtime).isoformat(timespec="seconds"),
            "closed_date": str(previous.get("closed_date") or ""),
            "customer_name": str(previous.get("customer_name") or ""),
            "insurance_company": str(previous.get("insurance_company") or ""),
            "claim_number": str(previous.get("claim_number") or ""),
            "date_of_loss": str(previous.get("date_of_loss") or ""),
            "town": str(previous.get("town") or ""),
            "owner_address": str(previous.get("owner_address") or ""),
            "location_of_vehicle": str(previous.get("location_of_vehicle") or ""),
            "vehicle": str(previous.get("vehicle") or ""),
            "vin": str(previous.get("vin") or ""),
            "shop_name": str(previous.get("shop_name") or ""),
            "shop_phone": str(previous.get("shop_phone") or ""),
            "shop_email": str(previous.get("shop_email") or ""),
            "contact_phone": str(previous.get("contact_phone") or ""),
            "contact_email": str(previous.get("contact_email") or ""),
            "notes": str(previous.get("notes") or ""),
            "note_history": [
                {"timestamp": str(entry.get("timestamp") or ""), "text": str(entry.get("text") or "")}
                for entry in (previous.get("note_history") or [])
                if isinstance(entry, dict)
            ],
            "assignment_claim_notes": str(previous.get("assignment_claim_notes") or ""),
            "office_progress_status": str(previous.get("office_progress_status") or ""),
            "office_appt_when": str(previous.get("office_appt_when") or ""),
            "office_waiting_for_paperwork": str(previous.get("office_waiting_for_paperwork") or "No"),
            "office_additional_notes": str(previous.get("office_additional_notes") or ""),
            "route_address_override": str(previous.get("route_address_override") or ""),
            "assign_pdf_path": str(previous.get("assign_pdf_path") or ""),
            "policy_number": str(previous.get("policy_number") or ""),
            "insured_name": str(previous.get("insured_name") or ""),
            "claimant_name": str(previous.get("claimant_name") or ""),
            "damage_description": str(previous.get("damage_description") or ""),
            "facts_of_loss": str(previous.get("facts_of_loss") or ""),
            "manual_overrides": dict(previous.get("manual_overrides") or {}),
        }

    def _claim_entries_for_folder(self, folder_path: Path, access_issues: list[str]) -> list[Path]:
        name = folder_path.name.strip().lower()
        if "closed" in name:
            return self._find_closed_claim_folders(folder_path, access_issues)
        try:
            entries = [
                entry
                for entry in folder_path.iterdir()
                if not (entry.is_file() and entry.name.lower() in {"desktop.ini", "thumbs.db"})
            ]
        except (PermissionError, OSError):
            access_issues.append(str(folder_path))
            return []
        return sorted(entries, key=lambda entry: entry.name.lower())

    def _find_closed_claim_folders(self, root: Path, access_issues: list[str]) -> list[Path]:
        claim_folders: list[Path] = []
        try:
            month_folders = [item for item in root.iterdir() if item.is_dir()]
        except (PermissionError, OSError):
            access_issues.append(str(root))
            return claim_folders

        for month_folder in sorted(month_folders, key=lambda entry: entry.name.lower()):
            try:
                day_folders = [item for item in month_folder.iterdir() if item.is_dir() and item.name.lower().startswith("claims ")]
            except (PermissionError, OSError):
                access_issues.append(str(month_folder))
                continue
            for day_folder in sorted(day_folders, key=lambda entry: entry.name.lower()):
                try:
                    entries = [item for item in day_folder.iterdir() if item.is_dir()]
                except (PermissionError, OSError):
                    access_issues.append(str(day_folder))
                    continue
                claim_folders.extend(sorted(entries, key=lambda entry: entry.name.lower()))
        return claim_folders

    def _default_status_for_folder(self, folder_path: Path) -> str:
        name = folder_path.name.strip().lower()
        if "closed" in name:
            return "Closed"
        if "pending" in name or "open" in name:
            return "Open"
        return "Open"

    def _claim_id_for_item(self, item: Path) -> str:
        source = item.stem if item.is_file() else item.name
        match = re.search(r"\b(\d{8})\b", source)
        if not match:
            match = re.search(r"\b(\d{5,7})\b", source)
        return match.group(1) if match else source

    def _candidate_pdf_paths(self, claim_item: Path) -> list[Path]:
        pdfs = sorted(claim_item.glob("*.pdf"), key=lambda path: path.name.lower())
        if not pdfs:
            return []

        def rank(path: Path) -> tuple[int, str]:
            name = path.name.lower()
            if name == "assign.pdf":
                return (0, name)
            if "assign" in name:
                return (1, name)
            if "claim summary" in name:
                return (2, name)
            if "estimate" in name:
                return (3, name)
            return (4, name)

        return sorted(pdfs, key=rank)

    def _claim_detail_parse_rank(self, path: Path) -> tuple[int, str]:
        name = path.name.lower()
        if "claim summary" in name or "summary" in name:
            return (0, name)
        if "supp prelim" in name or "preliminary supplement" in name or name.startswith("supp") or "supplement" in name:
            return (1, name)
        if "estimate" in name:
            return (2, name)
        if name == "assign.pdf" or "assign" in name:
            return (3, name)
        if "pre-scan" in name or "scan" in name:
            return (4, name)
        if "dop" in name or "direction to pay" in name:
            return (5, name)
        return (6, name)

    def _guess_claim_type_from_name(self, value: str) -> str:
        return self._normalize_claim_type(value, allow_blank=True)

    def _guess_claim_type_from_text(self, text: str, pdf_path: Path) -> str:
        name_guess = self._guess_claim_type_from_name(f"{pdf_path.name} {pdf_path.parent.name}")
        if "supplement" in name_guess.lower():
            return "Supplement"

        early_lines = "\n".join((text or "").splitlines()[:80]).lower()
        condensed = " ".join((text or "").split()).lower()
        supplement_patterns = (
            r"\bpreliminary supplement\b",
            r"(?m)^\s*supplement\s*$",
            r"\bsupplement photos?\b",
            r"\bsupplement request\b",
            r"\bsupplement estimate\b",
            r"\btear[\s-]?down supplement\b",
            r"\bre[\s-]?inspection\b",
            r"\badditional damage\b",
        )
        if any(re.search(pattern, early_lines) for pattern in supplement_patterns):
            return "Supplement"

        original_patterns = (
            r"\boriginal estimate\b",
            r"\binitial inspection\b",
            r"\bnew assignment\b",
            r"\bnew claim\b",
        )
        if any(re.search(pattern, condensed) for pattern in original_patterns):
            return "Original"

        return "Original" if name_guess == "Original" else ""

    def _normalize_claim_type(self, value: str, *, allow_blank: bool = False) -> str:
        lowered = (value or "").strip().lower()
        if not lowered:
            return "" if allow_blank else "Original"
        is_supplement_original = "supplement original" in lowered
        is_supplement = is_supplement_original or any(
            token in lowered
            for token in (
                "supp prelim",
                "preliminary supplement",
                "supplement",
                "reinspect",
                "re-inspection",
                "reinspection",
            )
        )
        is_supplement = is_supplement or bool(re.search(r"\bsupp\b", lowered))
        is_total_loss = any(
            token in lowered
            for token in (
                "total loss",
                "total-loss",
                "possible total",
            )
        )
        if is_supplement and is_total_loss:
            return "Supplement Total Loss"
        if is_total_loss:
            return "Total Loss"
        if is_supplement_original:
            return "Supplement Original"
        if is_supplement:
            return "Supplement"
        return "" if allow_blank else "Original"

    def _normalize_extracted_value(self, value: str) -> str:
        cleaned = " ".join(str(value or "").replace("\x00", " ").split()).strip(" -:\t")
        if not cleaned:
            return ""
        upper = cleaned.upper()
        blocked = {
            "CLAIM #:",
            "CLAIM #",
            "CLAIM:",
            "POLICY #:",
            "POLICY #",
            "POLICY",
            "ATTENTION:",
            "ATTENTION",
            "COMPANY:",
            "COMPANY",
            "PHONE:",
            "FAX:",
            "UNKNOWN",
            "NOT AVAILABLE",
        }
        if upper in blocked:
            return ""
        return cleaned

    def _normalize_customer_name(self, value: str) -> str:
        cleaned = self._normalize_extracted_value(value)
        if not cleaned:
            return ""
        if "," in cleaned:
            parts = [part.strip() for part in cleaned.split(",", 1)]
            if len(parts) == 2 and parts[0] and parts[1]:
                cleaned = f"{parts[1]} {parts[0]}"
        return cleaned.upper()

    def _fallback_customer_name_from_folder(self, claim_item: Path) -> str:
        raw_name = claim_item.name if claim_item.is_dir() else claim_item.stem
        raw_name = re.sub(r"^\s*\d{5,8}\s*-\s*", "", raw_name).strip()
        if not raw_name:
            return ""
        tokens: list[str] = []
        for token in raw_name.split():
            normalized = token.strip(",")
            if not normalized:
                continue
            if re.fullmatch(r"[A-Z0-9&'.]+", normalized):
                tokens.append(normalized)
                continue
            break
        if tokens:
            return " ".join(tokens).upper()
        return raw_name.split("-")[0].strip().upper()

    def _looks_suspicious_customer_name(self, value: str) -> bool:
        cleaned = self._normalize_extracted_value(value)
        if not cleaned:
            return True
        upper = cleaned.upper()
        bad_fragments = (
            "AT TIME OF INSPECTION",
            "CLAIMS PROCESS FULLY EXPLAINED",
            "SIGNED AUTHORIZATION",
            "DIRECTION OF PAY",
            "INVOICES AND PHOTOS ATTACHED",
            "SUPPLEMENTAL COMMENTS",
            "APPEARANCE ALLOWANCE",
            "JOB NUMBER",
            "INSPECTION LOCATION",
            "INSURANCE COMPANY",
        )
        if any(fragment in upper for fragment in bad_fragments):
            return True
        if any(char in cleaned for char in "{}[]"):
            return True
        if cleaned[:1] in {")", "]", "}", ">", "|"}:
            return True
        if any(token in upper for token in ("AUTO BODY", "BODY SHOP", "REPAIR FACILITY", "COLLISION CENTER")):
            return True
        return False

    def _looks_suspicious_insurance_company(self, value: str) -> bool:
        cleaned = self._normalize_extracted_value(value)
        if not cleaned:
            return True
        upper = cleaned.upper()
        if any(token in upper for token in ("AUTO BODY", "BODY SHOP", "REPAIR FACILITY", "INSPECTION LOCATION", "OWNER:")):
            return True
        if "OWNER" in upper and "INSURANCE" in upper:
            return True
        if cleaned.count(",") >= 1 and len(cleaned.split()) > 5:
            return True
        return False

    def _first_pattern_match(self, text: str, patterns: list[str], *, normalizer: Callable[[str], str] | None = None) -> str:
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
            if not match:
                continue
            value = match.group(1)
            cleaned = normalizer(value) if normalizer else self._normalize_extracted_value(value)
            if cleaned:
                return cleaned
        return ""

    def _detail_quality(self, key: str, value: Any) -> int:
        text = self._normalize_extracted_value(str(value or ""))
        if not text:
            return -1000
        lower = text.lower()
        score = 10
        if key == "customer_name":
            if self._looks_suspicious_customer_name(text):
                score -= 180
            if any(token in lower for token in ("auto body", "autobody", "repair", "collision", "motors", "shop")):
                score -= 40
            if re.search(r"\b[a-z]{2,}\b", text) and text == text.title():
                score += 5
            score += min(len(text), 40)
        elif key == "shop_name":
            if any(token in lower for token in ("auto body", "autobody", "collision", "motors", "shop", "bmw")):
                score += 40
            if any(token in lower for token in ("claims service", "factory service", "pre-repair", "pre-collision", "enform service", "days to repair")):
                score -= 120
            if "," in text:
                score -= 15
            if len(text.split()) > 8:
                score -= 20
            score += min(len(text), 35)
        elif key == "vehicle":
            if re.match(r"^(19|20)?\d{2}\s+", text):
                score += 50
            if "$" in text or "frame" in lower or "mech" in lower or "labor rate" in lower:
                score -= 120
            if any(token in lower for token in ("ccc intelligent solutions", "claims service", "factory service", "pre-repair", "pre-collision")):
                score -= 140
            score += min(len(text), 50)
        elif key == "insurance_company":
            if self._looks_suspicious_insurance_company(text):
                score -= 220
            if any(token in lower for token in ("insurance", "assurance", "mutual", "group", "casualty", "rock")):
                score += 40
            if any(token in lower for token in ("if applicable", "insurance company (", "|f applicable")):
                score -= 140
            if "|" in text or "(" in text or ")" in text:
                score -= 60
            score += min(len(text), 40)
        elif key == "town":
            if re.fullmatch(r"[A-Za-z .'-]+", text):
                score += 30
        elif key == "vin":
            if re.fullmatch(r"[A-HJ-NPR-Z0-9]{17}", text):
                score += 80
        elif key == "date_of_loss":
            if re.search(r"\d{1,2}/\d{1,2}/\d{2,4}", text):
                score += 50
        elif key == "claim_number":
            if re.search(r"\d", text):
                score += 40
            if re.search(r"[a-z]{3,}", text):
                score -= 60
        else:
            score += min(len(text), 30)
        return score

    def _extract_claim_details_from_text(self, text: str) -> dict[str, Any]:
        details: dict[str, Any] = {}

        customer_name = self._first_pattern_match(
            text,
            [
                r"Customer:\s*([^\n\r]+?)\s+Job Number:",
                r"Owner:\s*([^\n\r]+?)\s+Job Number:",
                r"Insured:\s*([^\n\r]+?)\s+Policy #:",
                r"(?s)Vehicle Owner.*?\n([A-Z][A-Z '&.,-]+)\n\d{1,6}\s+[A-Z0-9 .'-]+",
            ],
            normalizer=self._normalize_customer_name,
        )
        insurance_company = self._first_pattern_match(
            text,
            [
                r"For:\s*([^\n\r]+)",
                r"Insurance Company:\s*([^\n\r]+)",
                r"(?m)^Company:\s*([^\n\r]+)$",
            ],
        )
        claim_number = self._first_pattern_match(
            text,
            [
                r"Claim #:\s*([A-Z0-9][A-Z0-9 \-_/]{4,})",
                r"Claim\s*:\s*([A-Z0-9][A-Z0-9 \-_/]{4,})",
            ],
        )
        date_of_loss = self._first_pattern_match(
            text,
            [
                r"Date(?: Of)? Loss:\s*([0-9/:\- ]+(?:[AP]M)?)",
                r"Loss Date:\s*([0-9/:\- ]+(?:[AP]M)?)",
            ],
        )
        shop_name = self._first_pattern_match(
            text,
            [
                r"Shop Name:\s*([^\n\r]+)",
                r"Repair Facility:\s*([^\n\r]+)",
                r"Shop Name:\s*([^\n\r]+)",
                r"Inspection Location:\s*([^\n\r]+)",
                r"(?m)^([A-Z][A-Za-z0-9.&' -]{4,})\n(?:\d{1,6}\s+[A-Za-z0-9 .,'-]+)?\n(?:[A-Z][A-Za-z .'-]+,\s*[A-Z]{2}\s*\d{5}|EMAIL:|Phone:)",
                r"(?s)Estimate by:\s*\n\([^)]+\)\s*\n([^\n\r]+)",
            ],
        )
        vehicle = self._first_pattern_match(
            text,
            [
                r"(?m)^VEHICLE\s*\n([12]\d{3}[^\n\r]+)",
                r"(?s)Vehicle Information.*?\n([12]?\d{2,3}\s+[A-Z0-9][^\n\r]+)\n[A-HJ-NPR-Z0-9]{17}",
                r"(?m)^([12]\d{3}\s+[A-Z0-9][^\n\r]{6,})$",
                r"Vehicle\s*:\s*([^\n\r]+)",
            ],
        )
        vin = self._first_pattern_match(
            text,
            [
                r"VIN:\s*([A-HJ-NPR-Z0-9]{17})",
                r"\b([A-HJ-NPR-Z0-9]{17})\b",
            ],
        )
        contact_phone = self._first_pattern_match(
            text,
            [
                r"Owner:.*?\(\s*(\d{3}\)?[-.\s]*\d{3}[-.\s]*\d{4})\s*(?:Cell|Business|Day|Home|Work)",
                r"Customer:.*?\(\s*(\d{3}\)?[-.\s]*\d{3}[-.\s]*\d{4})",
                r"Insured:.*?\(\s*(\d{3}\)?[-.\s]*\d{3}[-.\s]*\d{4})",
                r"Contact Phone:\s*([0-9().\-\s]{10,})",
            ],
        )
        town = self._first_pattern_match(
            text,
            [
                r"(?s)Owner:.*?\n[^\n\r]+\n([A-Z][A-Za-z .'-]+),\s*[A-Z]{2}\s+\d{5}",
                r"(?s)Vehicle Owner.*?\n[A-Z][A-Z '&.,-]+\n[^\n\r]+\n([A-Z][A-Z .'-]+)\s+CT\s+\d{5}",
                r"Address:\s*[^\n\r,]+,\s*([A-Z][A-Za-z .'-]+)\s+[A-Z]{2}",
            ],
            normalizer=lambda value: self._normalize_extracted_value(value).title(),
        )

        if customer_name:
            details["customer_name"] = customer_name
        if insurance_company:
            details["insurance_company"] = insurance_company
        if claim_number:
            details["claim_number"] = claim_number
        if date_of_loss:
            details["date_of_loss"] = date_of_loss
        if shop_name:
            details["shop_name"] = shop_name
        if vehicle:
            details["vehicle"] = vehicle
        if vin:
            details["vin"] = vin
        if contact_phone:
            phone_digits = re.sub(r"\D", "", contact_phone)
            if len(phone_digits) >= 10:
                details["contact_phone"] = f"{phone_digits[-10:-7]}-{phone_digits[-7:-4]}-{phone_digits[-4:]}"
        if town:
            details["town"] = town
        if re.search(r"\bpossible total\b|\btotal[-\s]?loss\b(?!\s*:)", text, re.IGNORECASE):
            details["total_loss"] = True
        return details

    def _extract_claim_details(self, claim_item: Path) -> dict[str, Any]:
        if not claim_item.is_dir():
            return {}
        pdf_paths = self._candidate_pdf_paths(claim_item)
        if not pdf_paths:
            fallback_customer = self._fallback_customer_name_from_folder(claim_item)
            return {"customer_name": fallback_customer} if fallback_customer else {}

        assign_pdf = next(
            (path for path in pdf_paths if path.name.lower() == "assign.pdf" or "assign" in path.name.lower()),
            pdf_paths[0],
        )
        details: dict[str, Any] = {"assign_pdf_path": str(assign_pdf)}
        merged_details: dict[str, Any] = {}

        for pdf_path in sorted(pdf_paths, key=self._claim_detail_parse_rank):
            extracted = self._normalize_parsed_details(self._extract_details_from_pdf(pdf_path))
            for key, value in extracted.items():
                if key == "assign_pdf_path" or value in ("", None):
                    continue
                if isinstance(value, bool):
                    if value:
                        merged_details[key] = value
                    continue
                if self._detail_quality(key, value) <= 0:
                    continue
                if key not in merged_details or self._detail_quality(key, value) > self._detail_quality(key, merged_details.get(key)):
                    merged_details[key] = value

        folder_customer = self._fallback_customer_name_from_folder(claim_item)
        if folder_customer and not self._normalize_extracted_value(str(merged_details.get("customer_name") or "")):
            merged_details["customer_name"] = folder_customer

        for key, value in merged_details.items():
            if value:
                details[key] = value
        return details

    def _extract_details_from_pdf(self, pdf_path: Path) -> dict[str, Any]:
        text = self._extract_pdf_text(pdf_path, max_pages=5)
        if not text.strip():
            return {"assign_pdf_path": str(pdf_path)}

        if "Claim Summary" in text and "Owner:" in text:
            details = self._parse_claim_summary_pdf_text(text, pdf_path)
        elif "Assignment Sheet - Duhamel & Duhamel, LLC" in text or (
            "Insurance Information" in text and "Vehicle Owner" in text and "Appraisal Notes:" in text
        ):
            details = self._parse_assignment_pdf_text(text, pdf_path)
        elif all(fragment in text.lower() for fragment in ("instructions to estimator", "facts of loss", "vehicle information")):
            details = self._parse_email_thread_assignment_pdf_text(text, pdf_path)
        else:
            details = self._parse_generic_pdf_text(text, pdf_path)

        if not details.get("claim_id"):
            claim_id = self._extract_claim_id_from_text(text)
            if claim_id:
                details["claim_id"] = claim_id
        guessed_claim_type = self._guess_claim_type_from_text(text, pdf_path)
        if guessed_claim_type and not details.get("claim_type"):
            details["claim_type"] = guessed_claim_type
        details["assign_pdf_path"] = str(pdf_path)
        return details

    def _extract_labeled_section(self, text: str, headings: list[str], next_headings: list[str]) -> str:
        start_pattern = "|".join(re.escape(item) for item in headings)
        if "$" in next_headings:
            end_pattern = "$"
        else:
            end_pattern = "|".join(re.escape(item) for item in next_headings)
        pattern = rf"(?:{start_pattern})\s*:?\s*(.*?)(?=(?:{end_pattern})\s*:|$)"
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        return match.group(1).strip() if match else ""

    def _parse_contact_section(self, section_text: str) -> dict[str, str]:
        info = {"name": "", "address": "", "town": "", "phone": "", "email": ""}
        if not section_text:
            return info

        lines = [line.strip() for line in section_text.splitlines() if line.strip()]
        if not lines:
            return info

        name_match = re.search(r"\b(?:Owner|Claimant|Insured)\s*:\s*([^\n]+)", section_text, re.IGNORECASE)
        if name_match:
            name_candidate = name_match.group(1).strip()
        else:
            name_candidate = re.sub(r"^\((Claimant|Insured|Owner)\)\s*", "", lines[0], flags=re.IGNORECASE).strip()
        if not re.search(r"[@:]", name_candidate) and not re.search(r"\b(phone|cell|email|address)\b", name_candidate, re.IGNORECASE):
            info["name"] = name_candidate.title() if name_candidate.isupper() else name_candidate

        email_match = re.search(r"\bEmail\s*:\s*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})\b", section_text, re.IGNORECASE)
        if not email_match:
            email_match = re.search(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", section_text, re.IGNORECASE)
        if email_match:
            info["email"] = email_match.group(1) if email_match.lastindex else email_match.group(0)

        phone_match = re.search(r"\bPhone\s*:\s*\(?(\d{3})\)?[-.\s]*(\d{3})[-.\s]*(\d{4})", section_text, re.IGNORECASE)
        if not phone_match:
            phone_match = re.search(r"\(?(\d{3})\)?[-.\s]*(\d{3})[-.\s]*(\d{4})", section_text)
        if phone_match:
            info["phone"] = f"{phone_match.group(1)}-{phone_match.group(2)}-{phone_match.group(3)}"

        address_match = re.search(r"\bAddress\s*:\s*([^\n]+)", section_text, re.IGNORECASE)
        if address_match:
            info["address"] = " ".join(address_match.group(1).split())
        else:
            address_lines: list[str] = []
            for line in lines[1:] if info["name"] else lines:
                if info["email"] and info["email"] in line:
                    continue
                if re.search(r"\b(phone|cell|work|home|email|owner:|claimant:|insured:|address:)\b", line, re.IGNORECASE):
                    continue
                address_lines.append(line.title() if line.isupper() else line)
            if address_lines:
                info["address"] = " ".join(address_lines).strip()

        if info["address"]:
            town_match = re.search(r"\b([A-Z][A-Za-z .'-]+),\s*([A-Z]{2}),?\s*(\d{5}(?:-\d{4})?)\b", info["address"])
            if not town_match:
                town_match = re.search(r"\b([A-Z][A-Za-z .'-]+),\s*([A-Z]{2})\s+(\d{5}(?:-\d{4})?)\b", info["address"])
            if town_match:
                info["town"] = town_match.group(1).strip().title()

        return info

    def _parse_email_thread_assignment_pdf_text(self, text: str, pdf_path: Path) -> dict[str, Any]:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        normalized_text = "\n".join(lines)
        details: dict[str, Any] = {"assign_pdf_path": str(pdf_path)}

        file_number_match = re.search(r"\b(\d{5,8})\.pdf\b", pdf_path.name, re.IGNORECASE)
        if file_number_match:
            details["claim_id"] = file_number_match.group(1)

        claim_number_match = re.search(r"\bClaim\s*:\s*([A-Z0-9 \-]+)", normalized_text, re.IGNORECASE)
        if claim_number_match:
            details["claim_number"] = " ".join(claim_number_match.group(1).split())

        date_of_loss_match = re.search(r"Date Of Loss\s*:\s*([0-9/:-]+\s+[AP]M|[0-9/:-]+)", normalized_text, re.IGNORECASE)
        if date_of_loss_match:
            raw_dol = " ".join(date_of_loss_match.group(1).split())
            date_only_match = re.match(r"(\d{2}/\d{2}/\d{4})", raw_dol)
            details["date_of_loss"] = date_only_match.group(1) if date_only_match else raw_dol

        from_matches = re.findall(r"\bFrom\s*:\s*([^\n]+)", normalized_text, re.IGNORECASE)
        for insurer in from_matches:
            cleaned_insurer = insurer.strip(" ,")
            upper_insurer = cleaned_insurer.upper()
            if any(blocked in upper_insurer for blocked in ("DO_NOT_REPLY", "ASSIGNMENT")):
                continue
            details["insurance_company"] = cleaned_insurer
            break

        instructions_match = re.search(
            r"Instructions\s+to\s+Estimator\s*:\s*(.*?)Facts\s+of\s+Loss\s*:",
            text,
            re.IGNORECASE | re.DOTALL,
        )
        instructions_block = " ".join(instructions_match.group(1).split()) if instructions_match else ""
        if instructions_block:
            details["assignment_claim_notes"] = instructions_block
            damage_inline_match = re.search(r"Damage Description\s*:\s*(.+)", instructions_block, re.IGNORECASE)
            if damage_inline_match:
                details["damage_description"] = " ".join(damage_inline_match.group(1).split())

        facts_match = re.search(
            r"Facts\s+of\s+Loss\s*:\s*(.*?)Vehicle\s+Location",
            text,
            re.IGNORECASE | re.DOTALL,
        )
        facts_block = " ".join(facts_match.group(1).split()) if facts_match else ""
        if facts_block:
            details["facts_of_loss"] = facts_block

        vehicle_location_name_match = re.search(r"Vehicle Location Name\s*:\s*([^\n]+)", normalized_text, re.IGNORECASE)
        vehicle_location_match = re.search(r"Vehicle Location\s*:\s*([^\n]+)", normalized_text, re.IGNORECASE)
        vehicle_location_phone_match = re.search(
            r"Vehicle Location[\s\S]{0,200}?Phone\s*:\s*\(?(\d{3})\)?[-.\s]*(\d{3})[-.\s]*(\d{4})",
            normalized_text,
            re.IGNORECASE,
        )
        vehicle_location_email_match = re.search(
            r"Vehicle Location[\s\S]{0,250}?Email\s*:\s*([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})",
            normalized_text,
            re.IGNORECASE,
        )
        if vehicle_location_name_match:
            details["shop_name"] = vehicle_location_name_match.group(1).strip()
        if vehicle_location_match:
            details["location_of_vehicle"] = " ".join(vehicle_location_match.group(1).split())
            if not details.get("shop_name"):
                details["shop_name"] = details["location_of_vehicle"]
        if vehicle_location_phone_match:
            details["shop_phone"] = (
                f"{vehicle_location_phone_match.group(1)}-{vehicle_location_phone_match.group(2)}-{vehicle_location_phone_match.group(3)}"
            )
        if vehicle_location_email_match:
            details["shop_email"] = vehicle_location_email_match.group(1).strip()

        owner_block = self._extract_labeled_section(normalized_text, ["Owner", "Claimant"], ["Insured", "Vehicle Information"])
        owner_info = self._parse_contact_section(owner_block)
        if owner_info["name"]:
            details["customer_name"] = owner_info["name"]
            details["claimant_name"] = owner_info["name"]
        if owner_info["address"]:
            details["owner_address"] = owner_info["address"]
        if owner_info["town"]:
            details["town"] = owner_info["town"]
        if owner_info["phone"]:
            details["contact_phone"] = owner_info["phone"]
        if owner_info["email"]:
            details["contact_email"] = owner_info["email"]

        insured_block = self._extract_labeled_section(normalized_text, ["Insured"], ["Vehicle Information"])
        insured_info = self._parse_contact_section(insured_block)
        if insured_info["name"]:
            details["insured_name"] = insured_info["name"]
            if not details.get("customer_name"):
                details["customer_name"] = insured_info["name"]
        if not details.get("owner_address") and insured_info["address"]:
            details["owner_address"] = insured_info["address"]
        if not details.get("town") and insured_info["town"]:
            details["town"] = insured_info["town"]
        if not details.get("contact_phone") and insured_info["phone"]:
            details["contact_phone"] = insured_info["phone"]
        if not details.get("contact_email") and insured_info["email"]:
            details["contact_email"] = insured_info["email"]

        vehicle_block = self._extract_labeled_section(normalized_text, ["Vehicle Information"], ["Notice", "$"])
        if vehicle_block:
            vin_match = re.search(r"VIN\s*:\s*([A-HJ-NPR-Z0-9]{17})", vehicle_block, re.IGNORECASE)
            if vin_match:
                details["vin"] = vin_match.group(1).strip()
            vehicle_match = re.search(r"Vehicle\s*:\s*([^\n]+)", vehicle_block, re.IGNORECASE)
            if vehicle_match:
                details["vehicle"] = " ".join(vehicle_match.group(1).split())
            if not details.get("damage_description"):
                poi_match = re.search(r"Primary Point of Impact\s*:\s*([^\n]+)", vehicle_block, re.IGNORECASE)
                if poi_match:
                    details["damage_description"] = poi_match.group(1).strip()

        if re.search(r"\bpossible total\b|\btotal[-\s]?loss\b(?!\s*:)", normalized_text, re.IGNORECASE):
            details["total_loss"] = True

        return details

    def _parse_claim_summary_pdf_text(self, text: str, pdf_path: Path) -> dict[str, Any]:
        details: dict[str, Any] = {"assign_pdf_path": str(pdf_path)}

        owner_match = re.search(r"Owner:\s*([^\n\r]+)", text, re.IGNORECASE)
        if owner_match:
            details["customer_name"] = owner_match.group(1).strip()
        insured_match = re.search(r"Insured:\s*([^\n\r]+)", text, re.IGNORECASE)
        if insured_match:
            details["insured_name"] = insured_match.group(1).strip()
        insurer_match = re.search(r"Company:\s*([^\n\r]+?)\s+Claim #:", text, re.IGNORECASE)
        if insurer_match:
            details["insurance_company"] = insurer_match.group(1).strip()
        claim_match = re.search(r"Claim #:\s*([A-Z0-9 \-_]+)", text, re.IGNORECASE)
        if claim_match:
            details["claim_number"] = claim_match.group(1).strip()
        vin_match = re.search(r"VIN:\s*([A-HJ-NPR-Z0-9]{17})", text, re.IGNORECASE)
        if vin_match:
            details["vin"] = vin_match.group(1)
        vehicle_match = re.search(r"Owner:\s*[^\n\r]+\s*\n([^\n\r]+)", text, re.IGNORECASE)
        if vehicle_match:
            details["vehicle"] = vehicle_match.group(1).strip()
        town_match = re.search(r"\n([A-Z][A-Za-z .'-]+),\s*([A-Z]{2})\s+(\d{5})", text)
        if town_match:
            details["town"] = town_match.group(1).strip()
        phone_match = re.search(r"Owner:\s*[^\n\r]+\s+Cell:\s*\((\d{3})\)\s*(\d{3})-(\d{4})", text, re.IGNORECASE)
        if phone_match:
            details["contact_phone"] = f"{phone_match.group(1)}-{phone_match.group(2)}-{phone_match.group(3)}"
        shop_match = re.search(r"Place of Inspection:\s*([^\n\r]+)", text, re.IGNORECASE)
        if shop_match:
            details["shop_name"] = shop_match.group(1).strip()

        return details

    def _parse_assignment_pdf_text(self, text: str, pdf_path: Path) -> dict[str, Any]:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        insurance_company = ""
        claim_number = ""
        policy_number = ""
        customer_name = ""
        insured_name = ""
        claimant_name = ""
        date_of_loss = ""
        town = ""
        owner_address = ""
        location_of_vehicle = ""
        vehicle = ""
        vin = ""
        damage_description = ""
        facts_of_loss = ""
        shop_name = ""
        contact_phone = ""
        contact_email = ""
        shop_phone = ""
        shop_email = ""
        assignment_claim_notes = ""
        claim_type = ""

        def clean_line(value: str) -> str:
            value = value.strip()
            if not value or re.fullmatch(r"[-\s/@]+", value):
                return ""
            if value.upper() in {"UNKNOWN", "NOT AVAILABLE", "NOT READABLE"}:
                return ""
            return value

        def extract_name_from_owner_line(value: str) -> str:
            cleaned = clean_line(re.sub(r"^\((Claimant|Insured)\)\s*", "", value, flags=re.IGNORECASE))
            if not cleaned:
                return ""
            if re.search(r"\b(vehicle owner|vehicle information|home ph|work ph|pager|cell|ext)\b", cleaned, re.IGNORECASE):
                return ""
            if re.search(r"\b(not readable|manual|unknown|type|color|engine|trans|plate|style|mileage|damage|make|model|vin)\b", cleaned, re.IGNORECASE):
                return ""
            if re.search(r"\d", cleaned):
                return ""
            if not re.fullmatch(r"[A-Za-z][A-Za-z0-9&.,'() /-]*", cleaned):
                return ""
            return cleaned

        def clean_vehicle_line(value: str) -> str:
            cleaned = clean_line(value)
            if not cleaned:
                return ""
            year_match = re.match(r"^(\d{2})\s+(.+)$", cleaned)
            if year_match:
                year = int(year_match.group(1))
                full_year = 2000 + year if year <= 30 else 1900 + year
                cleaned = f"{full_year} {year_match.group(2)}"
            return cleaned.title() if cleaned.isupper() else cleaned

        def clean_date(value: str) -> str:
            match = re.match(r"\s*(\d{1,2})/(\d{1,2})/(\d{2,4})\b", value)
            if not match:
                return clean_line(value)
            month = int(match.group(1))
            day = int(match.group(2))
            year = match.group(3)[-2:]
            return f"{month:02d}/{day:02d}/{year}"

        def phone_from_line(value: str) -> str:
            match = re.search(r"\(?(\d{3})\)?[-.\s]*(\d{3})[-.\s]*(\d{4})", value)
            if not match:
                return ""
            return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"

        def looks_like_city_state_zip(value: str) -> bool:
            return bool(re.search(r"\b[A-Z][A-Z .'\-]+\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?\b", value))

        def looks_like_shop_name(value: str) -> bool:
            cleaned = clean_line(value)
            if not cleaned:
                return False
            if re.search(r"\d|@", cleaned):
                return False
            if re.fullmatch(r"[A-Z][A-Z .'\-&/]+", cleaned):
                return True
            return any(token in cleaned.lower() for token in ("auto", "body", "collision", "garage", "service", "repair"))

        try:
            insurance_idx = lines.index("Insurance Information")
            if insurance_idx >= 4:
                insurance_company = clean_line(lines[insurance_idx - 7]) if insurance_idx >= 7 else ""
                insured_name = clean_line(lines[insurance_idx - 6]) if insurance_idx >= 6 else ""
                claim_number = clean_line(lines[insurance_idx - 5]) if insurance_idx >= 5 else ""
                policy_number = clean_line(lines[insurance_idx - 4]) if insurance_idx >= 4 and lines[insurance_idx - 4] != "Not Available" else ""
        except ValueError:
            pass

        for date_pattern in (
            r"(\d{1,2}/\d{1,2}/\d{2,4})\s+Date of Loss:",
            r"Date of Loss:\s*(\d{1,2}/\d{1,2}/\d{2,4})",
        ):
            date_match = re.search(date_pattern, text, re.IGNORECASE)
            if date_match:
                date_of_loss = clean_date(date_match.group(1))
                break

        location_match = re.search(r"Location:\s*([^\n\r]+)", text, re.IGNORECASE)
        if location_match:
            location_of_vehicle = clean_line(location_match.group(1))

        if re.search(r"(?im)^\s*SUPPLEMENT\s*$", text):
            claim_type = "Supplement"

        try:
            owner_idx = lines.index("Vehicle Owner")
            owner_block = lines[max(0, owner_idx - 8):owner_idx]
            filtered = [
                line
                for line in owner_block
                if clean_line(line)
                and not re.search(r"\bvehicle location\b|\bvehicle owner\b|\bvehicle information\b", line, re.IGNORECASE)
            ]
            preferred_owner_names = [extract_name_from_owner_line(line) for line in filtered]
            preferred_owner_names = [line for line in preferred_owner_names if line]
            if preferred_owner_names:
                customer_name = preferred_owner_names[0]
            address_lines: list[str] = []
            town_line_index = -1
            for index, line in enumerate(filtered):
                town_match = re.search(r"\b([A-Z][A-Z ]+)\s+([A-Z]{2})\s+(\d{5})\b", line)
                if town_match:
                    town_line_index = index
                    town = town_match.group(1).title()
                    owner_address = " ".join(
                        part for part in [*address_lines, f"{town_match.group(1).title()}, {town_match.group(2)} {town_match.group(3)}"] if part
                    ).strip()
                    break
                if not preferred_owner_names or line != preferred_owner_names[0]:
                    cleaned_line = clean_line(line)
                    if not cleaned_line:
                        continue
                    if "not readable" in cleaned_line.lower():
                        continue
                    if re.search(r"\b(home ph|work ph|pager|cell|ext|vehicle owner|vehicle information)\b", cleaned_line, re.IGNORECASE):
                        continue
                    address_lines.append(cleaned_line.title() if cleaned_line.isupper() else cleaned_line)
            if not vehicle:
                vehicle_lines = filtered[town_line_index + 1:] if town_line_index >= 0 else filtered
                for index, line in enumerate(vehicle_lines):
                    vehicle_candidate = clean_line(line)
                    if re.match(r"^\d{2}\s+[A-Z]", vehicle_candidate, re.IGNORECASE):
                        vehicle = clean_vehicle_line(vehicle_candidate)
                        if index + 1 < len(vehicle_lines):
                            style_match = re.match(r"^[A-Z0-9-]+\s+([A-Z0-9][A-Z0-9 /-]{0,20})\s+Not Readable\b", vehicle_lines[index + 1], re.IGNORECASE)
                            if style_match:
                                style = style_match.group(1).strip()
                                if style and style.lower() not in vehicle.lower():
                                    vehicle = f"{vehicle} {style.upper() if style.isupper() else style}"
                        break
        except ValueError:
            pass

        owner_role_match = re.search(r"\((Claimant|Insured)\)[ \t]+([A-Z][A-Z .'\-]+)$", text, re.IGNORECASE | re.MULTILINE)
        if owner_role_match:
            owner_role = owner_role_match.group(1).lower()
            owner_name = clean_line(owner_role_match.group(2).title() if owner_role_match.group(2).isupper() else owner_role_match.group(2))
            customer_name = customer_name or owner_name
            if owner_role == "claimant":
                claimant_name = owner_name
            else:
                insured_name = owner_name
        elif "(Claimant)" in text:
            claimant_name = customer_name
        elif "(Insured)" in text:
            insured_name = insured_name or customer_name

        vin_match = re.search(r"\b([A-HJ-NPR-Z0-9]{17})\b", text)
        if vin_match:
            vin = vin_match.group(1)

        if vin:
            for index, line in enumerate(lines):
                if line == vin and index > 0:
                    vehicle_line = lines[index - 1]
                    if not vehicle_line.endswith(":"):
                        vehicle = clean_vehicle_line(vehicle_line)
                        if index + 1 < len(lines):
                            style_line = clean_line(lines[index + 1])
                            if (
                                style_line
                                and not style_line.endswith(":")
                                and not re.search(r"\b(vehicle owner|vehicle information|home ph|work ph|model|vin|plate|style|mileage|damage|make)\b", style_line, re.IGNORECASE)
                                and re.fullmatch(r"[A-Z0-9][A-Z0-9 /.'\-]{1,30}", style_line)
                            ):
                                style_text = style_line.title() if style_line.isupper() else style_line
                                if style_text.lower() not in vehicle.lower():
                                    vehicle = f"{vehicle} {style_text}"
                    break

        damage_match = re.search(r"Damage:\s*([^\n\r]+)", text, re.IGNORECASE)
        if damage_match:
            damage_description = clean_line(damage_match.group(1))
        if not damage_description or damage_description.lower() in {"make", "make:"}:
            try:
                damage_idx = lines.index("Damage:")
            except ValueError:
                damage_idx = -1
            if damage_idx >= 0:
                for follow_line in lines[damage_idx + 1:damage_idx + 8]:
                    candidate = clean_line(follow_line)
                    if not candidate or candidate.endswith(":"):
                        continue
                    damage_description = candidate
                    break

        facts_match = re.search(r"Facts of Loss:\s*([^\n\r]+)", text, re.IGNORECASE)
        if facts_match:
            facts_of_loss = clean_line(facts_match.group(1))
        if facts_match and len(facts_of_loss.split()) <= 2:
            facts_tail = text[facts_match.end():]
            stop_match = re.search(
                r"(?:^|\W|\d)(Impact\s+Notes|Appraisal\s+Notes|Registration\s+Expiration|Inspection\s+Number)\s*:",
                facts_tail,
                re.IGNORECASE,
            )
            facts_block = facts_tail[: stop_match.start()] if stop_match else facts_tail[:300]
            facts_of_loss = clean_line(f"{facts_of_loss} {' '.join(facts_block.split())}")

        shop_match = re.search(r"\(\s*[A-Z0-9]+\s*\)\s*\n([^\n]+)\n([^\n]+)\n([^\n]+)\n([\d-]+)", text)
        if shop_match:
            shop_name = clean_line(shop_match.group(1))
            shop_phone = clean_line(shop_match.group(4))
        if not shop_name:
            for index, line in enumerate(lines[:-4]):
                if not re.fullmatch(r"\(\s*[A-Z0-9]*\s*\)", line):
                    continue
                candidate_name = clean_line(lines[index + 1])
                candidate_city = clean_line(lines[index + 3])
                candidate_phone = phone_from_line(lines[index + 4])
                if looks_like_shop_name(candidate_name) and looks_like_city_state_zip(candidate_city) and candidate_phone:
                    shop_name = candidate_name
                    shop_phone = candidate_phone
                    break
        if location_of_vehicle and not shop_name:
            shop_name = location_of_vehicle
        if shop_name and not location_of_vehicle:
            location_of_vehicle = shop_name

        appraisal_idx = next((idx for idx, value in enumerate(lines) if value.lower().startswith("appraisal notes:")), -1)
        claimant_context = lines[max(0, appraisal_idx - 14):min(len(lines), appraisal_idx + 10)] if appraisal_idx >= 0 else []
        if claimant_context:
            claimant_phones: list[str] = []
            for line in claimant_context:
                for match in re.findall(r"\b\d{3}-\d{3}-\d{4}\b", line):
                    if match not in {"203-792-2150", "203-791-8066", "888-345-7797", "888-884-0885"}:
                        claimant_phones.append(match)
                for groups in re.findall(r"\(?(\d{3})\)?\s*(\d{3})-(\d{4})", line):
                    phone = f"{groups[0]}-{groups[1]}-{groups[2]}"
                    if phone not in {"203-792-2150", "203-791-8066", "888-345-7797", "888-884-0885"}:
                        claimant_phones.append(phone)
            if claimant_phones:
                contact_phone = claimant_phones[-1]
            for line in claimant_context:
                email_match = re.search(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", line, re.IGNORECASE)
                if email_match:
                    email_value = email_match.group(0)
                    if all(blocked not in email_value.lower() for blocked in ("claimsolution", "duhamel", "apptrak", "autobid", "claims@innovativeclaims")):
                        contact_email = email_value
                        break

        note_lines: list[str] = []
        note_prefixes = ("instructions to estimator:", "appraisal notes:")
        stop_prefixes = ("ext:", "location:", "fax:", "pager:", "cell:", "registration expiration:", "inspection number:", "estimate by:", "amount:")
        for idx, raw_line in enumerate(lines):
            lowered = raw_line.lower()
            if not lowered.startswith(note_prefixes):
                continue
            for follow_line in lines[idx:]:
                lowered_follow = follow_line.lower()
                if note_lines and lowered_follow.startswith(stop_prefixes):
                    break
                if "app trak" in lowered_follow or re.search(r"page \d+ of \d+", lowered_follow):
                    break
                cleaned_follow = follow_line.strip() if lowered_follow.startswith(note_prefixes) else clean_line(follow_line)
                if not cleaned_follow:
                    continue
                note_lines.append(cleaned_follow)
            if note_lines:
                assignment_claim_notes = " ".join(note_lines).strip()
                break
        if assignment_claim_notes.lower().rstrip(":") in {"appraisal notes", "instructions to estimator"}:
            assignment_claim_notes = ""

        return {
            "assign_pdf_path": str(pdf_path),
            "customer_name": customer_name,
            "insurance_company": insurance_company,
            "claim_number": claim_number,
            "policy_number": policy_number,
            "insured_name": insured_name,
            "claimant_name": claimant_name,
            "date_of_loss": date_of_loss,
            "town": town,
            "owner_address": owner_address,
            "location_of_vehicle": location_of_vehicle,
            "vehicle": vehicle,
            "vin": vin,
            "damage_description": damage_description,
            "facts_of_loss": facts_of_loss,
            "shop_name": shop_name,
            "shop_phone": shop_phone,
            "shop_email": shop_email,
            "contact_phone": contact_phone,
            "contact_email": contact_email,
            "assignment_claim_notes": assignment_claim_notes,
            "claim_type": claim_type,
        }

    def _parse_generic_pdf_text(self, text: str, pdf_path: Path) -> dict[str, Any]:
        details: dict[str, Any] = {"assign_pdf_path": str(pdf_path)}

        claim_match = re.search(r"Claim Number\s+([A-Z0-9-]+)", text, re.IGNORECASE)
        if not claim_match:
            claim_match = re.search(r"Claim #:\s*([A-Z0-9 \-_]+)", text, re.IGNORECASE)
        if claim_match:
            details["claim_number"] = claim_match.group(1).strip()

        owner_match = re.search(r"(?m)^Owner:\s*([^\n\r]+)", text, re.IGNORECASE)
        if owner_match:
            details["customer_name"] = owner_match.group(1).strip()

        town_match = re.search(r"\n[A-Z0-9 .'-]+\n([A-Z][A-Z ]+)\s+[A-Z]{2}\s+\d{5}\b", text)
        if town_match:
            details["town"] = town_match.group(1).title()

        insurer_match = re.search(r"\n([A-Z][A-Za-z &]+Insurance[^\n]*)\n", text)
        if not insurer_match:
            insurer_match = re.search(r"\n([A-Z][A-Za-z &]+Mutual[^\n]*)\n", text)
        if not insurer_match:
            insurer_match = re.search(r"(?m)^Company:\s*([^\n\r]+)", text, re.IGNORECASE)
        if insurer_match:
            details["insurance_company"] = insurer_match.group(1).strip()

        vin_match = re.search(r"\b([A-HJ-NPR-Z0-9]{17})\b", text)
        if vin_match:
            details["vin"] = vin_match.group(1)

        vehicle_match = re.search(r"\b(20\d{2}\s+[A-Z][A-Za-z0-9]+\s+[A-Z0-9][A-Za-z0-9 ]+)\b", text)
        if not vehicle_match:
            vehicle_match = re.search(r"Vehicle\s*:\s*([^\n\r]+)", text, re.IGNORECASE)
        if vehicle_match:
            details["vehicle"] = vehicle_match.group(1).strip()

        phone_match = re.search(r"Owner\s+[^\n]+\s+\((\d{3})\)\s*(\d{3})-(\d{4})", text, re.IGNORECASE)
        if not phone_match:
            phone_match = re.search(r"\(?(\d{3})\)?[-.\s]*(\d{3})[-.\s]*(\d{4})", text)
        if phone_match:
            details["contact_phone"] = f"{phone_match.group(1)}-{phone_match.group(2)}-{phone_match.group(3)}"

        email_match = re.search(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", text, re.IGNORECASE)
        if email_match:
            details["shop_email"] = email_match.group(0)

        estimate_shop_match = re.search(
            r"(?:MY WAY AUTO BODY|[A-Z][A-Za-z0-9&'.,\- ]+(?:AUTO BODY|BODY SHOP|COLLISION|COLLISION CENTER|MOTORS|SERVICE|REPAIR SERVICE|REPAIR))",
            text,
            re.IGNORECASE,
        )
        if estimate_shop_match:
            details["shop_name"] = estimate_shop_match.group(0).strip()

        if re.search(r"\bpossible total\b|\btotal[-\s]?loss\b(?!\s*:)", text, re.IGNORECASE):
            details["total_loss"] = True

        return details

    def _normalize_parsed_details(self, details: dict[str, Any]) -> dict[str, Any]:
        normalized: dict[str, Any] = {}
        for key, value in details.items():
            if value is None or value == "":
                continue
            if isinstance(value, bool):
                normalized[key] = value
                continue
            cleaned = self._normalize_extracted_value(str(value))
            if not cleaned:
                continue
            if key in {"customer_name", "insured_name", "claimant_name"}:
                cleaned = self._normalize_customer_name(cleaned)
            elif key == "claim_type":
                cleaned = self._normalize_claim_type(cleaned)
            elif key == "town":
                cleaned = cleaned.title()
            elif key in {"contact_phone", "shop_phone"}:
                digits = re.sub(r"\D", "", cleaned)
                if len(digits) >= 10:
                    cleaned = f"{digits[-10:-7]}-{digits[-7:-4]}-{digits[-4:]}"
            normalized[key] = cleaned
        return normalized

    def load_settings(self) -> dict[str, object]:
        payload = self._load_payload()
        watched_folders = [
            str(folder).strip()
            for folder in payload.get("watched_folders", [])
            if str(folder).strip()
        ]
        claim_tools_folder = str(payload.get("claim_tools_folder", "") or "").strip()
        return {
            "watched_folders": watched_folders or DEFAULT_WATCHED_FOLDERS,
            "claim_tools_folder": claim_tools_folder or DEFAULT_CLAIM_TOOLS_FOLDER,
            "route_home_address": str(payload.get("route_home_address", "") or DEFAULT_ROUTE_HOME_ADDRESS),
            "route_plan_keys": list(payload.get("route_plan_keys", [])),
            "processed_apptrak_pdfs": list(payload.get("processed_apptrak_pdfs", [])),
            "office_update_email_to": str(payload.get("office_update_email_to", "")),
            "office_update_email_cc": str(payload.get("office_update_email_cc", "")),
            "office_review_email_to": str(payload.get("office_review_email_to", "")),
            "office_rmc_email_to": str(payload.get("office_rmc_email_to", "")),
            "office_supplement_email_to": str(payload.get("office_supplement_email_to", "")),
        }

    def load_processed_apptrak_pdfs(self) -> list[str]:
        payload = self._load_payload()
        return [str(value) for value in payload.get("processed_apptrak_pdfs", []) if str(value).strip()]

    def load_claim_tools_contacts(self) -> list[ClaimToolContact]:
        payload = self._load_payload()
        results: list[ClaimToolContact] = []
        for raw in payload.get("claim_tools_contacts", []):
            if not isinstance(raw, dict):
                continue
            results.append(
                ClaimToolContact(
                    name=str(raw.get("name") or ""),
                    number=str(raw.get("number") or ""),
                    prompt_guide=str(raw.get("prompt_guide") or ""),
                    notes=str(raw.get("notes") or ""),
                )
            )
        return results

    def load_claim_tools_files(self) -> list[ClaimToolFile]:
        payload = self._load_payload()
        results: list[ClaimToolFile] = []
        for raw in payload.get("claim_tools_files", []):
            if not isinstance(raw, dict):
                continue
            results.append(
                ClaimToolFile(
                    section=str(raw.get("section") or ""),
                    label=str(raw.get("label") or ""),
                    file_path=str(raw.get("file_path") or ""),
                    notes=str(raw.get("notes") or ""),
                )
            )
        return results

    def load_body_shop_database(self) -> list[BodyShopEntry]:
        payload = self._load_payload()
        results: list[BodyShopEntry] = []
        for raw in payload.get("body_shop_database", []):
            if not isinstance(raw, dict):
                continue
            results.append(
                self._sanitize_body_shop_entry(
                    BodyShopEntry(
                        shop_name=str(raw.get("shop_name") or ""),
                        contact_name=str(raw.get("contact_name") or ""),
                        phone=str(raw.get("phone") or ""),
                        email=str(raw.get("email") or ""),
                        tax_id=str(raw.get("tax_id") or ""),
                        address=str(raw.get("address") or ""),
                        body_rate=str(raw.get("body_rate") or ""),
                        paint_rate=str(raw.get("paint_rate") or ""),
                        frame_rate=str(raw.get("frame_rate") or ""),
                        mechanical_rate=str(raw.get("mechanical_rate") or ""),
                        certifications=str(raw.get("certifications") or ""),
                        negotiation_notes=str(raw.get("negotiation_notes") or ""),
                        notes=str(raw.get("notes") or ""),
                    )
                )
            )
        return results

    def load_insurance_company_database(self) -> list[InsuranceCompanyEntry]:
        payload = self._load_payload()
        results: list[InsuranceCompanyEntry] = []
        for raw in payload.get("insurance_company_database", []):
            if not isinstance(raw, dict):
                continue
            results.append(
                self._sanitize_insurance_company_entry(
                    InsuranceCompanyEntry(
                        company_name=str(raw.get("company_name") or ""),
                        quick_summary=str(raw.get("quick_summary") or ""),
                        fatal_errors=str(raw.get("fatal_errors") or ""),
                        photo_rules=str(raw.get("photo_rules") or ""),
                        estimate_supp_rules=str(raw.get("estimate_supp_rules") or ""),
                        parts_rules=str(raw.get("parts_rules") or ""),
                        total_loss_rules=str(raw.get("total_loss_rules") or ""),
                        tow_rules=str(raw.get("tow_rules") or ""),
                        supplement_rules=str(raw.get("supplement_rules") or ""),
                        betterment_depreciation_rules=str(raw.get("betterment_depreciation_rules") or ""),
                        documentation_requirements=str(raw.get("documentation_requirements") or ""),
                        rates_and_sales_tax_rules=str(raw.get("rates_and_sales_tax_rules") or ""),
                        miscellaneous_rules=str(raw.get("miscellaneous_rules") or ""),
                        contact_information=str(raw.get("contact_information") or ""),
                        labor_rates=str(raw.get("labor_rates") or ""),
                        total_loss_threshold=str(raw.get("total_loss_threshold") or ""),
                        notes=str(raw.get("notes") or ""),
                    )
                )
            )
        return results

    def update_watched_folders(self, watched_folders: list[str]) -> None:
        payload = self._load_payload()
        payload["watched_folders"] = watched_folders
        self._save_payload(payload)

    def update_settings(self, watched_folders: list[str], claim_tools_folder: str) -> None:
        payload = self._load_payload()
        payload["watched_folders"] = watched_folders
        payload["claim_tools_folder"] = claim_tools_folder
        self._save_payload(payload)

    def save_route_plan_keys(self, route_plan_keys: list[str]) -> None:
        payload = self._load_payload()
        payload["route_plan_keys"] = route_plan_keys
        self._save_payload(payload)

    def save_route_home_address(self, route_home_address: str) -> None:
        payload = self._load_payload()
        payload["route_home_address"] = route_home_address
        self._save_payload(payload)

    def save_email_settings(self, settings: dict[str, str]) -> None:
        payload = self._load_payload()
        for key in (
            "office_update_email_to",
            "office_update_email_cc",
            "office_review_email_to",
            "office_rmc_email_to",
            "office_supplement_email_to",
        ):
            payload[key] = str(settings.get(key, "") or "").strip()
        self._save_payload(payload)

    def save_claim_tools_contacts(self, contacts: list[ClaimToolContact]) -> None:
        payload = self._load_payload()
        payload["claim_tools_contacts"] = [
            {
                "name": entry.name,
                "number": entry.number,
                "prompt_guide": entry.prompt_guide,
                "notes": entry.notes,
            }
            for entry in contacts
        ]
        self._save_payload(payload)

    def save_claim_tools_files(self, files: list[ClaimToolFile]) -> None:
        payload = self._load_payload()
        payload["claim_tools_files"] = [
            {
                "section": entry.section,
                "label": entry.label,
                "file_path": entry.file_path,
                "notes": entry.notes,
            }
            for entry in files
        ]
        self._save_payload(payload)

    def save_body_shop_database(self, entries: list[BodyShopEntry]) -> None:
        payload = self._load_payload()
        payload["body_shop_database"] = [
            {
                "shop_name": self._sanitize_body_shop_entry(entry).shop_name,
                "contact_name": self._sanitize_body_shop_entry(entry).contact_name,
                "phone": self._sanitize_body_shop_entry(entry).phone,
                "email": self._sanitize_body_shop_entry(entry).email,
                "tax_id": self._sanitize_body_shop_entry(entry).tax_id,
                "address": self._sanitize_body_shop_entry(entry).address,
                "body_rate": self._sanitize_body_shop_entry(entry).body_rate,
                "paint_rate": self._sanitize_body_shop_entry(entry).paint_rate,
                "frame_rate": self._sanitize_body_shop_entry(entry).frame_rate,
                "mechanical_rate": self._sanitize_body_shop_entry(entry).mechanical_rate,
                "certifications": self._sanitize_body_shop_entry(entry).certifications,
                "negotiation_notes": self._sanitize_body_shop_entry(entry).negotiation_notes,
                "notes": self._sanitize_body_shop_entry(entry).notes,
            }
            for entry in entries
        ]
        self._save_payload(payload)

    def save_insurance_company_database(self, entries: list[InsuranceCompanyEntry]) -> None:
        payload = self._load_payload()
        payload["insurance_company_database"] = [
            {
                "company_name": self._sanitize_insurance_company_entry(entry).company_name,
                "quick_summary": self._sanitize_insurance_company_entry(entry).quick_summary,
                "fatal_errors": self._sanitize_insurance_company_entry(entry).fatal_errors,
                "photo_rules": self._sanitize_insurance_company_entry(entry).photo_rules,
                "estimate_supp_rules": self._sanitize_insurance_company_entry(entry).estimate_supp_rules,
                "parts_rules": self._sanitize_insurance_company_entry(entry).parts_rules,
                "total_loss_rules": self._sanitize_insurance_company_entry(entry).total_loss_rules,
                "tow_rules": self._sanitize_insurance_company_entry(entry).tow_rules,
                "supplement_rules": self._sanitize_insurance_company_entry(entry).supplement_rules,
                "betterment_depreciation_rules": self._sanitize_insurance_company_entry(entry).betterment_depreciation_rules,
                "documentation_requirements": self._sanitize_insurance_company_entry(entry).documentation_requirements,
                "rates_and_sales_tax_rules": self._sanitize_insurance_company_entry(entry).rates_and_sales_tax_rules,
                "miscellaneous_rules": self._sanitize_insurance_company_entry(entry).miscellaneous_rules,
                "contact_information": self._sanitize_insurance_company_entry(entry).contact_information,
                "labor_rates": self._sanitize_insurance_company_entry(entry).labor_rates,
                "total_loss_threshold": self._sanitize_insurance_company_entry(entry).total_loss_threshold,
                "notes": self._sanitize_insurance_company_entry(entry).notes,
            }
            for entry in entries
        ]
        self._save_payload(payload)

    def sync_reference_data_from_claims(
        self,
        claim_keys: list[str] | None = None,
        include_claim_folders: bool = True,
    ) -> bool:
        payload = self._load_payload()
        claims_block = payload.get("claims", {})
        if not isinstance(claims_block, dict):
            return False

        body_shop_entries = self.load_body_shop_database()
        insurance_entries = self.load_insurance_company_database()
        changed = False

        changed = self._sync_insurance_company_files(payload, insurance_entries) or changed

        if include_claim_folders:
            for key, raw in claims_block.items():
                if claim_keys is not None and key not in claim_keys:
                    continue
                if not isinstance(raw, dict):
                    continue
                changed = self._sync_reference_data_for_record(raw, body_shop_entries, insurance_entries) or changed

        if changed:
            payload["body_shop_database"] = [self._body_shop_entry_payload(entry) for entry in body_shop_entries]
            payload["insurance_company_database"] = [self._insurance_company_entry_payload(entry) for entry in insurance_entries]
            self._save_payload(payload)
        return changed

    def save_processed_apptrak_pdfs(self, processed: list[str]) -> None:
        payload = self._load_payload()
        payload["processed_apptrak_pdfs"] = [str(value) for value in processed if str(value).strip()]
        self._save_payload(payload)

    def upsert_claim_record(self, claim_key: str, record: dict[str, Any]) -> None:
        payload = self._load_payload()
        claims_block = payload.setdefault("claims", {})
        claims_block[claim_key] = record
        self._save_payload(payload)

    def update_claim_fields(self, claim_key: str, updates: dict[str, Any]) -> None:
        payload = self._load_payload()
        claims_block = payload.setdefault("claims", {})
        record = claims_block.get(claim_key)
        if not isinstance(record, dict):
            return

        manual_overrides = dict(record.get("manual_overrides") or {})
        saved_state = payload.setdefault("claim_saved_state", {})
        state_entry = dict(saved_state.get(claim_key) or {})
        state_overrides = dict(state_entry.get("manual_overrides") or {})

        for field, value in updates.items():
            record[field] = value
            manual_overrides[field] = value
            state_overrides[field] = value

        record["manual_overrides"] = manual_overrides
        state_entry["manual_overrides"] = state_overrides
        saved_state[claim_key] = state_entry
        self._remember_manual_overrides(payload, claim_key, record, manual_overrides)
        self._save_payload(payload)

    def update_claim_payroll_settings(
        self,
        claim_key: str,
        *,
        excluded: bool | None = None,
        base_pay: float | None = None,
        total_loss_pay: float | None = None,
        note: str | None = None,
    ) -> None:
        payload = self._load_payload()
        claims_block = payload.setdefault("claims", {})
        record = claims_block.get(claim_key)
        if not isinstance(record, dict):
            return

        manual_overrides = dict(record.get("manual_overrides") or {})
        saved_state = payload.setdefault("claim_saved_state", {})
        state_entry = dict(saved_state.get(claim_key) or {})
        state_overrides = dict(state_entry.get("manual_overrides") or {})

        updates: dict[str, Any] = {
            "payroll_base_pay": None if base_pay is None else float(base_pay),
            "payroll_total_loss_pay": None if total_loss_pay is None else float(total_loss_pay),
        }
        if excluded is not None:
            updates["payroll_excluded"] = bool(excluded)
        if note is not None:
            updates["payroll_note"] = str(note).strip()

        for field, value in updates.items():
            record[field] = value
            manual_overrides[field] = value
            state_overrides[field] = value

        record["manual_overrides"] = manual_overrides
        state_entry["manual_overrides"] = state_overrides
        saved_state[claim_key] = state_entry
        self._remember_manual_overrides(payload, claim_key, record, manual_overrides)
        self._save_payload(payload)

    def append_claim_note(self, claim_key: str, note_text: str) -> None:
        payload = self._load_payload()
        claims_block = payload.setdefault("claims", {})
        record = claims_block.get(claim_key)
        if not isinstance(record, dict):
            return

        normalized_text = (note_text or "").strip()
        if not normalized_text:
            return

        timestamp = datetime.now().strftime("%m/%d/%Y %I:%M %p")
        note_history = list(record.get("note_history") or [])
        note_history.append({"timestamp": timestamp, "text": normalized_text})
        rendered_notes = render_note_history(note_history)

        record["note_history"] = note_history
        record["notes"] = rendered_notes
        record["updated_at"] = datetime.now().isoformat(timespec="seconds")

        saved_state = payload.setdefault("claim_saved_state", {})
        state_entry = dict(saved_state.get(claim_key) or {})
        state_entry["note_history"] = note_history
        state_entry["notes"] = rendered_notes
        saved_state[claim_key] = state_entry
        self._remember_claim_state(
            payload,
            claim_key,
            record,
            {"note_history": note_history, "notes": rendered_notes},
        )
        self._save_payload(payload)

    def update_office_fields(self, claim_key: str, updates: dict[str, str]) -> None:
        payload = self._load_payload()
        claims_block = payload.setdefault("claims", {})
        record = claims_block.get(claim_key)
        if not isinstance(record, dict):
            return

        office_updates = payload.setdefault("office_updates", {})
        office_entry = dict(office_updates.get(claim_key) or {})
        manual_overrides = dict(record.get("manual_overrides") or {})
        saved_state = payload.setdefault("claim_saved_state", {})
        state_entry = dict(saved_state.get(claim_key) or {})
        state_overrides = dict(state_entry.get("manual_overrides") or {})

        for field, value in updates.items():
            record[field] = value
            office_entry[field] = value
            manual_overrides[field] = value
            state_overrides[field] = value

        record["manual_overrides"] = manual_overrides
        office_updates[claim_key] = office_entry
        state_entry["manual_overrides"] = state_overrides
        saved_state[claim_key] = state_entry
        self._remember_manual_overrides(payload, claim_key, record, manual_overrides)
        self._save_payload(payload)

    def update_claim_status(self, claim_key: str, status: str) -> None:
        payload = self._load_payload()
        claims_block = payload.setdefault("claims", {})
        record = claims_block.get(claim_key)
        if not isinstance(record, dict):
            return

        timestamp = datetime.now().isoformat(timespec="seconds")
        record["status"] = status
        record["updated_at"] = timestamp
        if status.strip().lower() == "closed":
            record["closed_date"] = timestamp
        else:
            record["closed_date"] = ""

        manual_overrides = dict(record.get("manual_overrides") or {})
        manual_overrides["status"] = status
        record["manual_overrides"] = manual_overrides

        saved_state = payload.setdefault("claim_saved_state", {})
        state_entry = dict(saved_state.get(claim_key) or {})
        state_overrides = dict(state_entry.get("manual_overrides") or {})
        state_overrides["status"] = status
        state_entry["manual_overrides"] = state_overrides
        saved_state[claim_key] = state_entry
        self._remember_manual_overrides(payload, claim_key, record, manual_overrides)
        self._save_payload(payload)

    def move_claim_for_status(self, claim_key: str, status: str) -> str:
        payload = self._load_payload()
        claims_block = payload.setdefault("claims", {})
        record = claims_block.get(claim_key)
        if not isinstance(record, dict):
            return claim_key

        old_key = claim_key
        new_key = claim_key
        source_path = str(record.get("source_path") or "").strip()
        body_shop_entries = self.load_body_shop_database()
        insurance_entries = self.load_insurance_company_database()

        if source_path and source_path != "Manual Entry":
            source = Path(source_path)
            if source.exists() and source.is_dir():
                normalized_status = status.strip().lower()
                if normalized_status == "closed":
                    destination = self._build_closed_destination(payload, source)
                elif normalized_status == "open":
                    destination = self._build_pending_destination(payload, source)
                else:
                    destination = source

                if destination != source:
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    if source.resolve() != destination.resolve():
                        shutil.move(str(source), str(destination))
                    new_source = str(destination)
                    new_key = f"fs::{new_source}"
                    record["source_path"] = new_source
                    record["key"] = new_key
                    record["title"] = destination.name

        timestamp = datetime.now().isoformat(timespec="seconds")
        record["status"] = status
        record["updated_at"] = timestamp
        record["closed_date"] = timestamp if status.strip().lower() == "closed" else ""

        manual_overrides = dict(record.get("manual_overrides") or {})
        manual_overrides["status"] = status
        record["manual_overrides"] = manual_overrides

        if new_key != old_key:
            claims_block.pop(old_key, None)
        claims_block[new_key] = record
        self._rekey_claim_state_entries(payload, old_key, new_key)

        saved_state = payload.setdefault("claim_saved_state", {})
        state_entry = dict(saved_state.get(new_key) or {})
        state_overrides = dict(state_entry.get("manual_overrides") or {})
        state_overrides["status"] = status
        state_entry["manual_overrides"] = state_overrides
        saved_state[new_key] = state_entry
        self._remember_manual_overrides(payload, new_key, record, manual_overrides)
        if self._sync_reference_data_for_record(record, body_shop_entries, insurance_entries):
            payload["body_shop_database"] = [self._body_shop_entry_payload(entry) for entry in body_shop_entries]
            payload["insurance_company_database"] = [self._insurance_company_entry_payload(entry) for entry in insurance_entries]
        self._save_payload(payload)
        return new_key

    def _build_closed_destination(self, payload: dict, source: Path) -> Path:
        closed_folder = self._first_watched_folder(payload, "closed")
        if not closed_folder:
            raise RuntimeError("Closed Claims folder is not configured.")
        today = datetime.now()
        closed_root = Path(closed_folder)
        month_folder = closed_root / f"{today.month}-{today.year}"
        day_folder = month_folder / f"Claims {today.month}-{today.day}-{today.strftime('%y')}"
        return self._unique_destination(day_folder / source.name)

    def _build_pending_destination(self, payload: dict, source: Path) -> Path:
        pending_folder = self._first_watched_folder(payload, "pending")
        if not pending_folder:
            raise RuntimeError("Pending Claims folder is not configured.")
        return self._unique_destination(Path(pending_folder) / source.name)

    def _first_watched_folder(self, payload: dict, name_fragment: str) -> str:
        watched_folders = [str(folder) for folder in payload.get("watched_folders", [])]
        return next((folder for folder in watched_folders if name_fragment.lower() in folder.lower()), "")

    def _unique_destination(self, destination: Path) -> Path:
        if not destination.exists():
            return destination
        counter = 2
        while True:
            candidate = destination.with_name(f"{destination.name} ({counter})")
            if not candidate.exists():
                return candidate
            counter += 1

    def _rekey_claim_state_entries(self, payload: dict, old_key: str, new_key: str) -> None:
        if old_key == new_key:
            return

        for block_name in ("claim_saved_state", "office_updates"):
            block = payload.setdefault(block_name, {})
            if old_key in block and new_key not in block:
                block[new_key] = block.pop(old_key)
            elif old_key in block:
                block.pop(old_key, None)

        route_plan_keys = list(payload.get("route_plan_keys", []))
        payload["route_plan_keys"] = [new_key if key == old_key else key for key in route_plan_keys]

    def update_route_address(self, claim_key: str, route_address: str) -> None:
        payload = self._load_payload()
        claims_block = payload.setdefault("claims", {})
        record = claims_block.get(claim_key)
        if not isinstance(record, dict):
            return

        record["route_address_override"] = route_address

        manual_overrides = dict(record.get("manual_overrides") or {})
        manual_overrides["route_address_override"] = route_address
        record["manual_overrides"] = manual_overrides

        saved_state = payload.setdefault("claim_saved_state", {})
        state_entry = dict(saved_state.get(claim_key) or {})
        state_overrides = dict(state_entry.get("manual_overrides") or {})
        state_overrides["route_address_override"] = route_address
        state_entry["manual_overrides"] = state_overrides
        saved_state[claim_key] = state_entry
        self._remember_manual_overrides(payload, claim_key, record, manual_overrides)
        self._save_payload(payload)

    def _body_shop_entry_payload(self, entry: BodyShopEntry) -> dict[str, str]:
        entry = self._sanitize_body_shop_entry(entry)
        return {
            "shop_name": entry.shop_name,
            "contact_name": entry.contact_name,
            "phone": entry.phone,
            "email": entry.email,
            "tax_id": entry.tax_id,
            "address": entry.address,
            "body_rate": entry.body_rate,
            "paint_rate": entry.paint_rate,
            "frame_rate": entry.frame_rate,
            "mechanical_rate": entry.mechanical_rate,
            "certifications": entry.certifications,
            "negotiation_notes": entry.negotiation_notes,
            "notes": entry.notes,
        }

    def _insurance_company_entry_payload(self, entry: InsuranceCompanyEntry) -> dict[str, str]:
        entry = self._sanitize_insurance_company_entry(entry)
        return {
            "company_name": entry.company_name,
            "quick_summary": entry.quick_summary,
            "fatal_errors": entry.fatal_errors,
            "photo_rules": entry.photo_rules,
            "estimate_supp_rules": entry.estimate_supp_rules,
            "parts_rules": entry.parts_rules,
            "total_loss_rules": entry.total_loss_rules,
            "tow_rules": entry.tow_rules,
            "supplement_rules": entry.supplement_rules,
            "betterment_depreciation_rules": entry.betterment_depreciation_rules,
            "documentation_requirements": entry.documentation_requirements,
            "rates_and_sales_tax_rules": entry.rates_and_sales_tax_rules,
            "miscellaneous_rules": entry.miscellaneous_rules,
            "contact_information": entry.contact_information,
            "labor_rates": entry.labor_rates,
            "total_loss_threshold": entry.total_loss_threshold,
            "notes": entry.notes,
        }

    def _sync_reference_data_for_record(
        self,
        record: dict[str, Any],
        body_shop_entries: list[BodyShopEntry],
        insurance_entries: list[InsuranceCompanyEntry],
    ) -> bool:
        source_path = str(record.get("source_path") or "").strip()
        if not source_path or source_path == "Manual Entry":
            return False

        claim_folder = Path(source_path)
        if not claim_folder.exists() or not claim_folder.is_dir():
            return False

        changed = False

        shop_name = str(record.get("shop_name") or "").strip()
        if shop_name and shop_name != "No Shop Chosen":
            supporting_texts = self._collect_supporting_pdf_texts(claim_folder)
            changed = self._upsert_body_shop_from_claim(record, supporting_texts, body_shop_entries) or changed

        insurer_name = str(record.get("insurance_company") or "").strip()
        if insurer_name:
            assignment_text = self._extract_assignment_sheet_text(claim_folder)
            if assignment_text:
                changed = self._upsert_insurance_company_from_assignment(
                    insurer_name,
                    assignment_text,
                    insurance_entries,
                ) or changed

        return changed

    def _sync_insurance_company_files(
        self,
        payload: dict[str, Any],
        entries: list[InsuranceCompanyEntry],
    ) -> bool:
        changed = False
        for raw in payload.get("claim_tools_files", []):
            if not isinstance(raw, dict):
                continue
            section = str(raw.get("section") or "").strip().lower()
            if "insurance" not in section:
                continue
            company_name = str(raw.get("label") or "").strip()
            file_path = str(raw.get("file_path") or "").strip()
            if not company_name or not file_path:
                continue
            text = self._read_guideline_reference_text(Path(file_path))
            if not text:
                continue
            changed = self._upsert_insurance_company_from_assignment(
                company_name,
                text,
                entries,
                prefer_incoming=True,
            ) or changed
        return changed

    def _read_guideline_reference_text(self, file_path: Path) -> str:
        if not file_path.exists() or not file_path.is_file():
            return ""
        suffix = file_path.suffix.lower()
        if suffix in {".txt", ".md"}:
            try:
                return file_path.read_text(encoding="utf-8-sig")
            except UnicodeDecodeError:
                try:
                    return file_path.read_text(encoding="utf-8")
                except UnicodeDecodeError:
                    try:
                        return file_path.read_text(encoding="cp1252")
                    except OSError:
                        return ""
                except OSError:
                    return ""
            except OSError:
                return ""
        if suffix == ".pdf":
            return self._extract_pdf_text(file_path, max_pages=12)
        return ""

    def _normalize_lookup(self, value: str) -> str:
        return re.sub(r"[^a-z0-9]+", "", (value or "").strip().lower())

    def _find_body_shop_index(self, entries: list[BodyShopEntry], shop_name: str) -> int | None:
        normalized_target = self._normalize_lookup(shop_name)
        if not normalized_target:
            return None
        for index, entry in enumerate(entries):
            normalized_candidate = self._normalize_lookup(entry.shop_name)
            if normalized_candidate == normalized_target:
                return index
        for index, entry in enumerate(entries):
            normalized_candidate = self._normalize_lookup(entry.shop_name)
            if normalized_target in normalized_candidate or normalized_candidate in normalized_target:
                return index
        return None

    def _find_insurance_company_index(
        self,
        entries: list[InsuranceCompanyEntry],
        company_name: str,
    ) -> int | None:
        normalized_target = self._normalize_lookup(company_name)
        if not normalized_target:
            return None
        for index, entry in enumerate(entries):
            normalized_candidate = self._normalize_lookup(entry.company_name)
            if normalized_candidate == normalized_target:
                return index
        for index, entry in enumerate(entries):
            normalized_candidate = self._normalize_lookup(entry.company_name)
            if normalized_target in normalized_candidate or normalized_candidate in normalized_target:
                return index
        return None

    def _merge_text_field(self, existing: str, incoming: str) -> str:
        existing = (existing or "").strip()
        incoming = (incoming or "").strip()
        if not incoming:
            return existing
        if not existing:
            return incoming
        existing_lines = [line.strip() for line in existing.splitlines() if line.strip()]
        for line in [line.strip() for line in incoming.splitlines() if line.strip()]:
            if line not in existing_lines:
                existing_lines.append(line)
        return "\n".join(existing_lines)

    def _first_clean_line(self, value: str) -> str:
        for line in [line.strip() for line in str(value or "").splitlines() if line.strip()]:
            return line
        return ""

    def _sanitize_phone_value(self, value: str) -> str:
        matches = re.findall(r"(?:\+?1[-.\s]*)?(?:\(?\d{3}\)?[-.\s]*)\d{3}[-.\s]*\d{4}", str(value or ""))
        for match in matches:
            normalized = re.sub(r"\s+", " ", match).strip()
            if normalized:
                return normalized
        return ""

    def _sanitize_email_value(self, value: str) -> str:
        matches = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", str(value or ""), flags=re.IGNORECASE)
        for match in matches:
            lowered = match.lower()
            if any(domain in lowered for domain in ("duhamels.com", "ambroadjust.com")):
                continue
            return lowered
        return ""

    def _sanitize_address_value(self, value: str) -> str:
        text = str(value or "")
        for line in [line.strip(" ,") for line in text.splitlines() if line.strip()]:
            if re.search(
                r"\b\d{1,5}\s+[A-Za-z0-9.\- ]+\s(?:St|Street|Ave|Avenue|Rd|Road|Dr|Drive|Ln|Lane|Blvd|Boulevard|Ct|Court|Way|Pl|Place)\b",
                line,
                flags=re.IGNORECASE,
            ):
                return line
        match = re.search(
            r"\b\d{1,5}\s+[A-Za-z0-9.\- ]+\s(?:St|Street|Ave|Avenue|Rd|Road|Dr|Drive|Ln|Lane|Blvd|Boulevard|Ct|Court|Way|Pl|Place)\b[^\n,]*(?:,\s*[A-Za-z .'-]+,\s*[A-Z]{2}\s*\d{5})?",
            text,
            flags=re.IGNORECASE,
        )
        return match.group(0).strip(" ,") if match else ""

    def _sanitize_rate_value(self, value: str) -> str:
        match = re.search(r"\b(\d{2,3}(?:\.\d{1,2})?)\b", str(value or ""))
        if not match:
            return ""
        raw = match.group(1)
        try:
            parsed = float(raw)
        except ValueError:
            return ""
        if parsed < 10 or parsed > 250:
            return ""
        return f"{parsed:.2f}" if "." in raw else str(int(parsed))

    def _sanitize_certifications_value(self, value: str) -> str:
        found: list[str] = []
        for cert in ("I-CAR", "ASE", "Tesla", "Aluminum", "Certified"):
            if re.search(rf"\b{re.escape(cert)}\b", str(value or ""), flags=re.IGNORECASE) and cert not in found:
                found.append(cert)
        return ", ".join(found)

    def _sanitize_free_text(self, value: str, max_lines: int = 12) -> str:
        lines: list[str] = []
        for line in [re.sub(r"\s+", " ", line).strip() for line in str(value or "").splitlines()]:
            if not line:
                continue
            if line not in lines:
                lines.append(line)
        return "\n".join(lines[:max_lines])

    def _sanitize_guideline_text(self, value: str, max_lines: int = 10) -> str:
        lines: list[str] = []
        for raw_line in str(value or "").splitlines():
            cleaned = self._clean_guideline_line(raw_line)
            if not cleaned or self._looks_like_claim_metadata_line(cleaned):
                continue
            if cleaned not in lines:
                lines.append(cleaned)
        return "\n".join(lines[:max_lines])

    def _sanitize_body_shop_entry(self, entry: BodyShopEntry) -> BodyShopEntry:
        return BodyShopEntry(
            shop_name=(entry.shop_name or "").strip(),
            contact_name=self._first_clean_line(entry.contact_name),
            phone=self._sanitize_phone_value(entry.phone),
            email=self._sanitize_email_value(entry.email),
            tax_id=self._first_clean_line(entry.tax_id),
            address=self._sanitize_address_value(entry.address),
            body_rate=self._sanitize_rate_value(entry.body_rate),
            paint_rate=self._sanitize_rate_value(entry.paint_rate),
            frame_rate=self._sanitize_rate_value(entry.frame_rate),
            mechanical_rate=self._sanitize_rate_value(entry.mechanical_rate),
            certifications=self._sanitize_certifications_value(entry.certifications),
            negotiation_notes=self._sanitize_free_text(entry.negotiation_notes, max_lines=12),
            notes=self._sanitize_free_text(entry.notes, max_lines=18),
        )

    def _sanitize_insurance_company_entry(self, entry: InsuranceCompanyEntry) -> InsuranceCompanyEntry:
        return InsuranceCompanyEntry(
            company_name=(entry.company_name or "").strip(),
            quick_summary=self._sanitize_guideline_text(entry.quick_summary, max_lines=8),
            fatal_errors=self._sanitize_guideline_text(entry.fatal_errors, max_lines=12),
            photo_rules=self._sanitize_guideline_text(entry.photo_rules, max_lines=12),
            estimate_supp_rules=self._sanitize_guideline_text(entry.estimate_supp_rules, max_lines=12),
            parts_rules=self._sanitize_guideline_text(entry.parts_rules, max_lines=12),
            total_loss_rules=self._sanitize_guideline_text(entry.total_loss_rules, max_lines=12),
            tow_rules=self._sanitize_guideline_text(entry.tow_rules, max_lines=8),
            supplement_rules=self._sanitize_guideline_text(entry.supplement_rules, max_lines=10),
            betterment_depreciation_rules=self._sanitize_guideline_text(entry.betterment_depreciation_rules, max_lines=8),
            documentation_requirements=self._sanitize_guideline_text(entry.documentation_requirements, max_lines=10),
            rates_and_sales_tax_rules=self._sanitize_guideline_text(entry.rates_and_sales_tax_rules, max_lines=10),
            miscellaneous_rules=self._sanitize_guideline_text(entry.miscellaneous_rules, max_lines=12),
            contact_information=self._sanitize_guideline_text(entry.contact_information, max_lines=8),
            labor_rates=self._sanitize_guideline_text(entry.labor_rates, max_lines=8),
            total_loss_threshold=self._sanitize_guideline_text(entry.total_loss_threshold, max_lines=6),
            notes=self._sanitize_guideline_text(entry.notes, max_lines=12),
        )

    def _looks_like_claim_metadata_line(self, line: str) -> bool:
        upper = line.upper()
        metadata_tokens = (
            "FILE #",
            "CLAIM #",
            "DATE OF LOSS",
            "DATE INSPECTED",
            "LOSS TYPE",
            "OWNER:",
            "INSURED:",
            "ODOMETER",
            "VIN:",
            "LICENSE:",
            "APPRAISER:",
            "WRITTEN BY:",
            "ADJUSTER:",
            "JOB NUMBER",
            "POLICY #:",
            "POINT OF IMPACT",
            "PRIMARY IMPACT",
            "SECONDARY IMPACT",
            "VEHICLE INFORMATION",
            "VEHICLE OWNER",
            "CLAIM SUMMARY",
            "ESTIMATE OF RECORD",
            "APPRAISAL COMPANY",
            "INSURANCE COMPANY",
        )
        return any(token in upper for token in metadata_tokens)

    def _extract_instruction_candidate_lines(self, text: str) -> list[str]:
        keywords = (
            "photo",
            "photos",
            "total loss",
            "tow",
            "storage",
            "supplement",
            "paperwork",
            "release",
            "call",
            "email",
            "lkq",
            "aftermarket",
            "reconditioned",
            "labor rate",
            "sales tax",
            "invoice",
            "tear down",
            "salvage",
            "comparable",
            "guide",
            "instruction",
            "repairable",
        )
        lines: list[str] = []
        for raw_line in text.splitlines():
            cleaned = self._clean_guideline_line(raw_line)
            if not cleaned or self._looks_like_claim_metadata_line(cleaned):
                continue
            lower = cleaned.lower()
            if any(keyword in lower for keyword in keywords) and cleaned not in lines:
                lines.append(cleaned)
        return lines

    def _extract_shop_focus_text(self, shop_name: str, texts: dict[str, str]) -> str:
        normalized_name = self._normalize_lookup(shop_name)
        tokens = [
            token
            for token in re.split(r"[^A-Za-z0-9]+", normalized_name)
            if len(token) >= 4 and token not in {"auto", "body", "shop", "repair", "collision", "service"}
        ]
        snippets: list[str] = []
        for content in texts.values():
            lines = [line.strip() for line in content.splitlines() if line.strip()]
            for index, line in enumerate(lines):
                normalized_line = self._normalize_lookup(line)
                score = sum(1 for token in tokens if token in normalized_line)
                if normalized_name and normalized_name in normalized_line:
                    score += 2
                if score >= 2:
                    window = lines[max(0, index - 2) : min(len(lines), index + 4)]
                    snippets.extend(window)
        deduped: list[str] = []
        for line in snippets:
            if line not in deduped:
                deduped.append(line)
        return "\n".join(deduped)

    def _pick_shop_email(self, text: str) -> str:
        matches = re.findall(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, flags=re.IGNORECASE)
        for match in matches:
            lowered = match.lower()
            if any(domain in lowered for domain in ("duhamels.com", "ambroadjust.com")):
                continue
            return lowered
        return ""

    def _pick_best_structured_value(self, field: str, current: str, incoming: str) -> str:
        current = (current or "").strip()
        incoming = (incoming or "").strip()
        if not incoming:
            return current
        if not current:
            return incoming
        structured_fields = {"phone", "email", "address", "body_rate", "paint_rate", "frame_rate", "mechanical_rate"}
        if field not in structured_fields:
            return self._merge_text_field(current, incoming)
        if len(current.splitlines()) > 1 and len(incoming.splitlines()) <= 1:
            return incoming
        if field in {"phone", "email"} and len(incoming.splitlines()) < len(current.splitlines()):
            return incoming
        if field in {"body_rate", "paint_rate", "frame_rate", "mechanical_rate"} and self._sanitize_rate_value(incoming):
            return incoming
        if field == "address" and current != incoming and len(incoming) < len(current) and "," in incoming:
            return incoming
        return current

    def _upsert_body_shop_from_claim(
        self,
        record: dict[str, Any],
        texts: dict[str, str],
        entries: list[BodyShopEntry],
    ) -> bool:
        shop_name = str(record.get("shop_name") or "").strip()
        if not shop_name or shop_name == "No Shop Chosen":
            return False

        parsed = self._parse_shop_reference_details(shop_name, texts)
        candidate = BodyShopEntry(
            shop_name=shop_name,
            contact_name="",
            phone=str(record.get("shop_phone") or parsed.get("phone") or "").strip(),
            email=str(record.get("shop_email") or parsed.get("email") or "").strip(),
            tax_id="",
            address=str(parsed.get("address") or "").strip(),
            body_rate=str(parsed.get("body_rate") or "").strip(),
            paint_rate=str(parsed.get("paint_rate") or "").strip(),
            frame_rate=str(parsed.get("frame_rate") or "").strip(),
            mechanical_rate=str(parsed.get("mechanical_rate") or "").strip(),
            certifications=str(parsed.get("certifications") or "").strip(),
            negotiation_notes="",
            notes="",
        )

        index = self._find_body_shop_index(entries, shop_name)
        if index is None:
            if not any(
                [
                    candidate.phone,
                    candidate.email,
                    candidate.address,
                    candidate.body_rate,
                    candidate.paint_rate,
                    candidate.frame_rate,
                    candidate.mechanical_rate,
                    candidate.certifications,
                ]
            ):
                return False
            entries.append(candidate)
            return True

        existing = entries[index]
        updated = False
        for field in (
            "phone",
            "email",
            "address",
            "body_rate",
            "paint_rate",
            "frame_rate",
            "mechanical_rate",
            "certifications",
        ):
            incoming = getattr(candidate, field)
            current = getattr(existing, field)
            replacement = self._pick_best_structured_value(field, current, incoming)
            if replacement != current:
                setattr(existing, field, replacement)
                updated = True
        return updated

    def _upsert_insurance_company_from_assignment(
        self,
        company_name: str,
        assignment_text: str,
        entries: list[InsuranceCompanyEntry],
        prefer_incoming: bool = False,
    ) -> bool:
        parsed = self._parse_assignment_guidelines(assignment_text)
        if not any(parsed.values()):
            return False

        candidate = InsuranceCompanyEntry(
            company_name=company_name,
            quick_summary=parsed.get("quick_summary", ""),
            fatal_errors=parsed.get("fatal_errors", ""),
            photo_rules=parsed.get("photo_rules", ""),
            estimate_supp_rules=parsed.get("estimate_supp_rules", ""),
            parts_rules=parsed.get("parts_rules", ""),
            total_loss_rules=parsed.get("total_loss_rules", ""),
            tow_rules=parsed.get("tow_rules", ""),
            supplement_rules=parsed.get("supplement_rules", ""),
            betterment_depreciation_rules=parsed.get("betterment_depreciation_rules", ""),
            documentation_requirements=parsed.get("documentation_requirements", ""),
            rates_and_sales_tax_rules=parsed.get("rates_and_sales_tax_rules", ""),
            miscellaneous_rules=parsed.get("miscellaneous_rules", ""),
            contact_information=parsed.get("contact_information", ""),
            labor_rates=parsed.get("labor_rates", ""),
            total_loss_threshold=parsed.get("total_loss_threshold", ""),
            notes=parsed.get("notes", ""),
        )

        index = self._find_insurance_company_index(entries, company_name)
        if index is None:
            entries.append(candidate)
            return True

        existing = entries[index]
        updated = False
        for field in (
            "quick_summary",
            "fatal_errors",
            "photo_rules",
            "estimate_supp_rules",
            "parts_rules",
            "total_loss_rules",
            "tow_rules",
            "supplement_rules",
            "betterment_depreciation_rules",
            "documentation_requirements",
            "rates_and_sales_tax_rules",
            "miscellaneous_rules",
            "contact_information",
            "labor_rates",
            "total_loss_threshold",
            "notes",
        ):
            incoming = getattr(candidate, field)
            current = getattr(existing, field)
            replacement = incoming if prefer_incoming and incoming else self._merge_text_field(current, incoming)
            if replacement != current:
                setattr(existing, field, replacement)
                updated = True
        return updated

    def _collect_supporting_pdf_texts(self, claim_folder: Path) -> dict[str, str]:
        texts: dict[str, str] = {}
        for pdf_path in sorted(claim_folder.glob("*.pdf"), key=lambda item: item.name.lower()):
            name = pdf_path.name.lower()
            if not any(token in name for token in ("estimate", "supp", "claimsummary", "claim summary", "assign")):
                continue
            text = self._extract_pdf_text(pdf_path)
            if text:
                texts[name] = text
        return texts

    def _extract_assignment_sheet_text(self, claim_folder: Path) -> str:
        for pdf_path in sorted(claim_folder.glob("*.pdf"), key=lambda item: item.name.lower()):
            lower_name = pdf_path.name.lower()
            if lower_name == "assign.pdf" or "assign" in lower_name:
                text = self._extract_pdf_text(pdf_path)
                if text:
                    return text
        return ""

    def _extract_pdf_text(self, pdf_path: Path, max_pages: int = 8) -> str:
        if PdfReader is None:
            return self._ocr_pdf_text(pdf_path, max_pages=max_pages)
        try:
            reader = PdfReader(str(pdf_path))
        except Exception:
            return self._ocr_pdf_text(pdf_path, max_pages=max_pages)
        chunks: list[str] = []
        for page in reader.pages[:max_pages]:
            try:
                chunks.append(page.extract_text() or "")
            except Exception:
                continue
        text = "\n".join(chunks).strip()
        if text and len(re.sub(r"\W+", "", text)) >= 40:
            return text
        return self._ocr_pdf_text(pdf_path, max_pages=max_pages)

    def _ocr_pdf_text(self, pdf_path: Path, max_pages: int = 8) -> str:
        if not TESSERACT_EXE.exists() or not GHOSTSCRIPT_EXE.exists():
            return ""
        try:
            with tempfile.TemporaryDirectory(prefix="claim_ocr_", dir=str(DATA_DIR)) as temp_dir:
                temp_path = Path(temp_dir)
                pattern = str(temp_path / "page-%03d.png")
                gs_command = [
                    str(GHOSTSCRIPT_EXE),
                    "-dSAFER",
                    "-dBATCH",
                    "-dNOPAUSE",
                    "-sDEVICE=pnggray",
                    "-r220",
                    "-dTextAlphaBits=4",
                    "-dGraphicsAlphaBits=4",
                    "-dFirstPage=1",
                    f"-dLastPage={max_pages}",
                    f"-sOutputFile={pattern}",
                    str(pdf_path),
                ]
                subprocess.run(gs_command, check=True, capture_output=True, text=True, encoding="utf-8", errors="ignore")

                texts: list[str] = []
                for image_path in sorted(temp_path.glob("page-*.png")):
                    try:
                        env = os.environ.copy()
                        if TESSDATA_DIR.exists():
                            env["TESSDATA_PREFIX"] = str(TESSDATA_DIR) + os.sep
                        result = subprocess.run(
                            [str(TESSERACT_EXE), str(image_path), "stdout", "--psm", "6"],
                            check=True,
                            capture_output=True,
                            text=True,
                            encoding="utf-8",
                            errors="ignore",
                            env=env,
                        )
                    except Exception:
                        continue
                    content = (result.stdout or "").strip()
                    if content:
                        texts.append(content)
                return "\n".join(texts).strip()
        except Exception:
            return ""

    def _parse_shop_reference_details(self, shop_name: str, texts: dict[str, str]) -> dict[str, str]:
        combined = "\n".join(texts.values())
        if not combined.strip():
            return {}

        details: dict[str, str] = {}
        focus_text = self._extract_shop_focus_text(shop_name, texts)
        email = self._pick_shop_email(focus_text or combined)
        if email:
            details["email"] = email

        phone = self._sanitize_phone_value(focus_text)
        if phone:
            details["phone"] = phone

        address = self._sanitize_address_value(focus_text)
        if address:
            details["address"] = address

        rate_patterns = {
            "body_rate": r"(?:Body|Sheet Metal|Body Labor)\s*(?:Rate)?\s*[:$ ]+\s*(\d+(?:\.\d{1,2})?)",
            "paint_rate": r"(?:Paint|Refinish)\s*(?:Rate)?\s*[:$ ]+\s*(\d+(?:\.\d{1,2})?)",
            "frame_rate": r"(?:Frame)\s*(?:Rate)?\s*[:$ ]+\s*(\d+(?:\.\d{1,2})?)",
            "mechanical_rate": r"(?:Mechanical)\s*(?:Rate)?\s*[:$ ]+\s*(\d+(?:\.\d{1,2})?)",
        }
        rate_source = "\n".join(
            text for name, text in texts.items() if any(token in name for token in ("estimate", "supp"))
        )
        for field, pattern in rate_patterns.items():
            match = re.search(pattern, rate_source, flags=re.IGNORECASE)
            if match:
                details[field] = self._sanitize_rate_value(match.group(1).strip())

        certifications: list[str] = []
        for cert in ("I-CAR", "ASE", "Tesla", "Aluminum", "Certified"):
            if re.search(rf"\b{re.escape(cert)}\b", focus_text or combined, flags=re.IGNORECASE):
                certifications.append(cert)
        if certifications:
            details["certifications"] = ", ".join(dict.fromkeys(certifications))
        return details

    def _parse_assignment_guidelines(self, text: str) -> dict[str, str]:
        extracted_sections = self._extract_guideline_sections(text)
        buckets = {
            "quick_summary": [],
            "fatal_errors": [],
            "photo_rules": [],
            "estimate_supp_rules": [],
            "parts_rules": [],
            "total_loss_rules": [],
            "tow_rules": [],
            "supplement_rules": [],
            "betterment_depreciation_rules": [],
            "documentation_requirements": [],
            "rates_and_sales_tax_rules": [],
            "miscellaneous_rules": [],
            "contact_information": [],
            "labor_rates": [],
            "total_loss_threshold": [],
            "notes": [],
        }

        for section_name, lines in extracted_sections.items():
            for raw_line in lines:
                cleaned = self._clean_guideline_line(raw_line)
                if not cleaned:
                    continue
                target = self._bucket_guideline_line(cleaned, section_name)
                if target and cleaned not in buckets[target]:
                    buckets[target].append(cleaned)

        if not any(buckets.values()):
            for cleaned in self._extract_instruction_candidate_lines(text):
                target = self._bucket_guideline_line(cleaned, "miscellaneous_rules")
                if target and cleaned not in buckets[target]:
                    buckets[target].append(cleaned)

        if not buckets["quick_summary"] and buckets["miscellaneous_rules"]:
            buckets["quick_summary"] = buckets["miscellaneous_rules"][:3]

        return {key: "\n".join(values) for key, values in buckets.items() if values}

    def _extract_guideline_sections(self, text: str) -> dict[str, list[str]]:
        lines = [line.strip() for line in text.splitlines()]
        preamble: list[str] = []
        key_rules: list[str] = []
        labor_rates: list[str] = []
        explicit_sections: dict[str, list[str]] = {
            "quick_summary": [],
            "fatal_errors": [],
            "photo_rules": [],
            "estimate_supp_rules": [],
            "parts_rules": [],
            "total_loss_rules": [],
            "tow_rules": [],
            "supplement_rules": [],
            "betterment_depreciation_rules": [],
            "documentation_requirements": [],
            "rates_and_sales_tax_rules": [],
            "miscellaneous_rules": [],
            "contact_information": [],
            "labor_rates": [],
            "total_loss_threshold": [],
            "notes": [],
        }
        heading_map = {
            "CLIENT QUICK SUMMARY": "quick_summary",
            "QUICK SUMMARY": "quick_summary",
            "CLIENT FATAL ERROR LIST": "fatal_errors",
            "FATAL ERRORS": "fatal_errors",
            "CLIENT PHOTO RULES": "photo_rules",
            "PHOTO RULES": "photo_rules",
            "CLIENT ESTIMATE/SUPPLEMENT RELEASE RULES": "estimate_supp_rules",
            "ESTIMATE / SUPPLEMENT RELEASE RULES": "estimate_supp_rules",
            "ESTIMATE/SUPPLEMENT RELEASE RULES": "estimate_supp_rules",
            "CLIENT PARTS APPLICATION RULES": "parts_rules",
            "PARTS APPLICATION RULES": "parts_rules",
            "CLIENT TOTAL LOSS RULES": "total_loss_rules",
            "TOTAL LOSS RULES": "total_loss_rules",
            "CLIENT TOW CHARGE RULES": "tow_rules",
            "TOW CHARGE RULES": "tow_rules",
            "CLIENT SUPPLEMENT HANDLING RULES": "supplement_rules",
            "SUPPLEMENT RULES": "supplement_rules",
            "SUPPLEMENT HANDLING RULES": "supplement_rules",
            "CLIENT BETTERMENT/DEPRECIATION RULES": "betterment_depreciation_rules",
            "BETTERMENT / DEPRECIATION RULES": "betterment_depreciation_rules",
            "BETTERMENT/DEPRECIATION RULES": "betterment_depreciation_rules",
            "CLIENT DOCUMENTATION REQUIREMENTS": "documentation_requirements",
            "DOCUMENTATION REQUIREMENTS": "documentation_requirements",
            "CLIENT RATES AND SALES TAX RULES": "rates_and_sales_tax_rules",
            "RATES AND SALES TAX RULES": "rates_and_sales_tax_rules",
            "CLIENT MISCELLANEOUS RULES": "miscellaneous_rules",
            "MISCELLANEOUS RULES": "miscellaneous_rules",
            "CLIENT CONTACT INFORMATION": "contact_information",
            "CONTACT INFORMATION": "contact_information",
            "LABOR RATES": "labor_rates",
            "TOTAL LOSS THRESHOLD": "total_loss_threshold",
            "NOTES": "notes",
        }
        capture_preamble = True
        capture_key_rules = False
        capture_labor = False
        current_explicit: str | None = None

        for line in lines:
            upper = line.upper()
            if not line:
                continue
            if upper in heading_map:
                current_explicit = heading_map[upper]
                capture_preamble = False
                capture_key_rules = False
                capture_labor = False
                continue
            if upper.startswith("REQUIRED AND OPTIONAL DOCUMENTS/IMAGES FROM APPRAISER"):
                current_explicit = None
                capture_preamble = False
                capture_key_rules = False
                capture_labor = False
                continue
            if current_explicit:
                explicit_sections[current_explicit].append(line)
                continue
            if upper.startswith("LINK TO EXISTING IMAGES") or upper.startswith("STATUS:") or upper.startswith("SCA FILE #"):
                capture_preamble = False
                capture_key_rules = False
                capture_labor = False
                continue
            if upper.startswith("CLIENT'S KEY MD RULES"):
                capture_key_rules = True
                capture_labor = False
                continue
            if upper.startswith("AVERAGE LABOR RATES"):
                capture_labor = True
                capture_key_rules = False
                continue
            if upper.startswith("AREA OF DAMAGE") or upper.startswith("REMARKS") or upper.startswith("INSURED INFORMATION"):
                capture_key_rules = False
                capture_labor = False
                continue
            if capture_preamble:
                preamble.append(line)
            elif capture_key_rules:
                key_rules.append(line)
            elif capture_labor:
                labor_rates.append(line)

        if any(explicit_sections.values()):
            preamble = []

        sections = {
            "preamble": preamble,
            "key_rules": key_rules,
            "labor_rates": labor_rates,
        }
        sections.update(explicit_sections)
        return sections

    def _clean_guideline_line(self, line: str) -> str:
        cleaned = re.sub(r"\s+", " ", line).strip(" -;\t")
        cleaned = re.sub(r"^\d+[.)]?\s*", "", cleaned)
        if not cleaned or len(cleaned) < 3:
            return ""
        upper = cleaned.upper()
        bad_fragments = (
            "ASSIGNMENT SHEET",
            "SCA FILE #",
            "FILE #",
            "FEE:",
            "CLAIM #",
            "LOSS TYPE",
            "DATE OF LOSS",
            "DATE INSPECTED",
            "APPOINTMENT DATE",
            "LOCATE:",
            "DATE RECEIVED",
            "DATE ASSIGNED",
            "ZIP LOOKUP",
            "LOSS DATE",
            "OWNER INFORMATION",
            "INSURED INFORMATION",
            "VEHICLE LOCATION INFORMATION",
            "AREA OF DAMAGE",
            "REMARKS",
            "ADJUSTER",
            "CUSTOMER PHONE",
            "LOCATION PHONE",
            "HOME #",
            "WORK #",
            "CELL #",
            "EMAIL",
            "COMPANY",
            "NAME",
            "ADDRESS",
            "OWNER",
            "VIN",
            "LICENSE PLATE",
            "PHOTO:",
        )
        if any(fragment in upper for fragment in bad_fragments):
            return ""
        return cleaned

    def _bucket_guideline_line(self, line: str, section_name: str) -> str:
        upper = line.upper()
        explicit_targets = {
            "quick_summary": "quick_summary",
            "fatal_errors": "fatal_errors",
            "photo_rules": "photo_rules",
            "estimate_supp_rules": "estimate_supp_rules",
            "parts_rules": "parts_rules",
            "total_loss_rules": "total_loss_rules",
            "tow_rules": "tow_rules",
            "supplement_rules": "supplement_rules",
            "betterment_depreciation_rules": "betterment_depreciation_rules",
            "documentation_requirements": "documentation_requirements",
            "rates_and_sales_tax_rules": "rates_and_sales_tax_rules",
            "miscellaneous_rules": "miscellaneous_rules",
            "contact_information": "contact_information",
            "labor_rates": "labor_rates",
            "total_loss_threshold": "total_loss_threshold",
            "notes": "notes",
        }
        if section_name in explicit_targets:
            return explicit_targets[section_name]
        if section_name == "labor_rates":
            if "65%" in upper or "NADA" in upper or "KBB" in upper:
                return "total_loss_threshold"
            return "labor_rates"
        if "FATAL" in upper or "DO NOT" in upper and ("RELEASE" in upper or "COMPLETE" in upper):
            return "fatal_errors"
        if any(token in upper for token in ("PHOTO", "VIN", "ODOMETER", "LICENSE PLATE", "4 CORNERS", "INTERIOR", "SHARP EDGE", "ARROWS")):
            return "photo_rules"
        if "SUPPLEMENT" in upper or "PRIOR AUTHORIZATION" in upper:
            return "supplement_rules"
        if "ESTIMATE" in upper or "APPRAISAL COMMENTS" in upper or "RELEASE ANY PAPERWORK" in upper:
            return "estimate_supp_rules"
        if any(token in upper for token in ("PART", "LKQ", "A/M")):
            return "parts_rules"
        if "TOTAL LOSS" in upper:
            return "total_loss_rules"
        if any(token in upper for token in ("KBB", "NADA", "65%")):
            return "total_loss_threshold"
        if any(token in upper for token in ("TOW", "STORAGE")):
            return "tow_rules"
        if any(token in upper for token in ("BETTERMENT", "DEPRECIATION")):
            return "betterment_depreciation_rules"
        if any(token in upper for token in ("REGISTRATION", "DOCUMENTATION", "INVOICE")):
            return "documentation_requirements"
        if "LABOR RATE" in upper or "SHEET METAL/BODY" in upper or "REFINISH" in upper or "MECHANICAL" in upper or "FRAME" in upper or "PAINT AND MATERIALS" in upper:
            return "rates_and_sales_tax_rules"
        if any(token in upper for token in ("PHONE", "FAX", "CONTACT", "E-MAIL", "EMAIL")):
            return "contact_information"
        return "miscellaneous_rules"


def render_note_history(entries: list[dict[str, str]]) -> str:
    lines: list[str] = []
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        text = str(entry.get("text") or "").strip()
        if not text:
            continue
        timestamp = str(entry.get("timestamp") or "").strip()
        if timestamp:
            lines.append(f"[{timestamp}]")
        lines.append(text)
        lines.append("")
    return "\n".join(lines).strip()
