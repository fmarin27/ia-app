from __future__ import annotations

import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

from claims_manager import ClaimsManagerApp, ClaimsStore, DATA_FILE  # noqa: E402


def main() -> None:
    app = ClaimsManagerApp.__new__(ClaimsManagerApp)
    app.store = ClaimsStore(DATA_FILE)
    app.selected_key = None
    app.scan_access_issues = []
    app.refresh_claim_tools_contacts_view = lambda: None
    app.refresh_claim_tools_files_view = lambda: None
    app.refresh_body_shops_view = lambda: None
    app.refresh_views = lambda select_key=None: None
    ClaimsManagerApp.scan_folders(app, show_access_warnings=False)


if __name__ == "__main__":
    main()
