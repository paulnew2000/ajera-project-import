"""
Small Deltek Ajera API client - just what the project importer needs.

Settings come from environment variables (or a .env file next to this script):

    AJERA_API_URL          the full API URL from Ajera (Company > API Access)
    AJERA_USERNAME         the API user name
    AJERA_PASSWORD         the API user's password
    AJERA_EXPECTED_DB_ID   the DatabaseID you intend to write to (safety check, see README)

Ajera facts this client handles for you:
  * CreateProjects only exists in API version 2; most reads work on version 1.
  * A badly shaped CreateProjects request comes back "successful" but creates nothing,
    so a create only counts when Ajera returns a ProjectKey.
"""
from __future__ import annotations

import base64
import json
import os
import threading
import urllib.parse

import httpx


class AjeraError(Exception):
    def __init__(self, message: str, errors: list[str] | None = None):
        self.errors = errors or []
        super().__init__(message + (f" | {'; '.join(self.errors)}" if self.errors else ""))


def decode_url_token(url: str) -> dict:
    """The query string of an Ajera API URL is base64-encoded JSON:
    {"ClientID": ..., "DatabaseID": ..., "IsSampleData": ...}."""
    if "?" not in url:
        raise ValueError("AJERA_API_URL does not look like an Ajera API URL (no '?' part)")
    raw = urllib.parse.unquote(url.split("?", 1)[1])
    return json.loads(base64.b64decode(raw + "=" * (-len(raw) % 4)))


class Ajera:
    def __init__(self, url: str | None = None, username: str | None = None, password: str | None = None):
        self.url = url or os.environ.get("AJERA_API_URL", "")
        self._user = username or os.environ.get("AJERA_USERNAME", "")
        self._pwd = password or os.environ.get("AJERA_PASSWORD", "")
        missing = [n for n, v in (("AJERA_API_URL", self.url), ("AJERA_USERNAME", self._user),
                                  ("AJERA_PASSWORD", self._pwd)) if not v]
        if missing:
            raise RuntimeError(f"Missing setting(s): {', '.join(missing)}. Copy .env.example to .env and fill it in.")
        token = decode_url_token(self.url)
        self.database_id = token.get("DatabaseID")
        self.is_sample_data = bool(token.get("IsSampleData"))
        self._sessions: dict[int, str] = {}
        self._lock = threading.Lock()

    # -- safety ------------------------------------------------------------------------
    def check_write_target(self) -> None:
        """Writes are only allowed when AJERA_EXPECTED_DB_ID matches the database in the URL.
        This stops a copied/pasted production URL from receiving test data by accident."""
        expected = os.environ.get("AJERA_EXPECTED_DB_ID", "").strip()
        if not expected:
            raise RuntimeError(
                f"AJERA_EXPECTED_DB_ID is not set. Your URL points at DatabaseID {self.database_id}; "
                f"if that is the database you mean to write to, put AJERA_EXPECTED_DB_ID={self.database_id} in .env.")
        if str(self.database_id) != expected:
            raise RuntimeError(
                f"Your URL points at DatabaseID {self.database_id} but AJERA_EXPECTED_DB_ID is {expected}. "
                "Nothing was created.")

    # -- transport -----------------------------------------------------------------------
    def _post(self, payload: dict) -> dict:
        r = httpx.post(self.url, json=payload, timeout=180)
        r.raise_for_status()
        return r.json()

    def _session(self, version: int) -> str:
        with self._lock:
            if version not in self._sessions:
                d = self._post({"Method": "CreateAPISession", "APIVersion": version,
                                "Username": self._user, "Password": self._pwd})
                if d.get("ResponseCode") != 200:
                    raise AjeraError(f"Could not sign in to Ajera (API v{version}): {d.get('Message')}")
                self._sessions[version] = d["Content"]["SessionToken"]
            return self._sessions[version]

    def call(self, method: str, args: dict | None = None, version: int = 1) -> dict:
        for attempt in (1, 2):
            d = self._post({"Method": method, "SessionToken": self._session(version), "MethodArguments": args or {}})
            msg = str(d.get("Message", ""))
            if attempt == 1 and "session" in msg.lower() and d.get("ResponseCode") != 200:
                self._sessions.pop(version, None)  # expired session: sign in again once
                continue
            break
        if d.get("ResponseCode") != 200:
            errs = [e.get("ErrorMessage", "") for e in d.get("Errors", []) if e.get("ErrorMessage") != "An error has occurred."]
            raise AjeraError(f"{method}: {msg}", errs)
        return d.get("Content") or {}

    # -- reads ---------------------------------------------------------------------------
    def list_all_projects(self) -> list[dict]:
        """No status filter = every project. (A status list silently narrows to Active only.)"""
        return self.call("ListProjects").get("Projects", [])

    def get_projects(self, keys: list[int]) -> list[dict]:
        out: list[dict] = []
        for i in range(0, len(keys), 50):
            out += self.call("GetProjects", {"RequestedProjects": keys[i:i + 50]}).get("Projects", [])
        return out

    def get_employees(self, keys: list[int]) -> list[dict]:
        out: list[dict] = []
        for i in range(0, len(keys), 50):
            out += self.call("GetEmployees", {"RequestedEmployees": keys[i:i + 50]}).get("Employees", [])
        return out

    # -- write ---------------------------------------------------------------------------
    def create_project(self, args: dict) -> dict:
        """API v2 CreateProjects. Returns Ajera's created-project record (with ProjectKey)."""
        content = self.call("CreateProjects", args, version=2)
        projects = content.get("Projects") or []
        if not projects or not projects[0].get("ProjectKey"):
            raise AjeraError("Ajera did not create the project (it did not accept the request's shape)")
        return projects[0]
