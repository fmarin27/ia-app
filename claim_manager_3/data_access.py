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
DATA_FILE = APP_DIR / "claims_data.json"
DESKTOP_MODE = (os.environ.get("CLAIM_MANAGER_DESKTOP_MODE") or "home").strip().lower()
OFFICE_SYNC_ENABLED = DESKTOP_MODE == "office"
HOME_APP_ROOT = Path(os.environ.get("CLAIM_MANAGER_HOME_ROOT") or r"C:\Users\ferna\Desktop\IA APP")
DEFAULT_WATCHED_FOLDERS = [
    str(APP_DIR / "PENDING CLAIMS"),
    str(APP_DIR / "Closed Claims"),
]
DEFAULT_CLAIM_TOOLS_FOLDER = str(APP_DIR / "Claim Tools")
DEFAULT_ROUTE_HOME_ADDRESS = "5 Richlee Rd, Norwalk, CT 06851"
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
        if not self.data_file.exists():
            return {}
        try:
            return json.loads(self.data_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return {}

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
                )
            )
        if changed:
            self._save_payload(payload)
        results.sort(key=lambda claim: (claim.status != "Open", claim.claim_id or claim.title))
        return results

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

    def update_claim_fields(self, claim_key: str, updates: dict[str, str]) -> None:
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
            with tempfile.TemporaryDirectory(prefix="claim_ocr_", dir=str(APP_DIR)) as temp_dir:
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
