from __future__ import annotations

import os

os.environ.setdefault("CLAIM_MANAGER_DESKTOP_MODE", "office")

import app as office_app


office_app.APP_NAME = "Claim Manager 3.0 Office"


if __name__ == "__main__":
    raise SystemExit(office_app.main())
