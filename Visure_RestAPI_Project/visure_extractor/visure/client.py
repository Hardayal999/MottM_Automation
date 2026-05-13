"""
client.py
=========
The Visure API client. This is the ONLY file in the project that touches
the Visure API. Everything else works with Python objects returned from here.

Design principles
-----------------
1. ONLY GET requests for reading data. The one exception is POST /authenticate
   and POST /project/current — both required by Visure even to *read* anything.
   We use NO write endpoints (no POST/PUT/DELETE on requirements data).

2. Session-based: Visure pins API calls to a "current project". You authenticate,
   then call POST /project/current with {project, group}, and then subsequent
   /specifications calls return specs for that project. We hide this behind
   set_current_project() so the caller doesn't have to remember.

3. Every method raises a clear exception on failure. The caller decides what
   to do (retry, log, abort). No silent failures.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import requests

from visure.config import settings


# ---------------------------------------------------------------------------
# Custom exceptions — so callers can tell auth errors from other errors
# ---------------------------------------------------------------------------

class VisureAPIError(Exception):
    """Base class for any error from the Visure API."""


class VisureAuthError(VisureAPIError):
    """Raised when authentication fails."""


# ---------------------------------------------------------------------------
# Lightweight data holders
# ---------------------------------------------------------------------------

@dataclass
class Project:
    """A Visure project as returned by /authenticate or /projects.

    We pull out just the fields we actually use elsewhere. The raw dict
    is kept in `raw` in case we need anything else later.
    """
    id: int
    name: str
    code: str | None
    group_id: int | None    # the first group we have access to in this project
    raw: dict

    @classmethod
    def from_dict(cls, d: dict) -> "Project":
        # A project from /authenticate carries a `groups` array. We need at
        # least one group to call /project/current. We pick the first.
        # If a project has multiple groups and the user needs to choose,
        # we'll add a prompt later — out of scope for v1.
        groups = d.get("groups") or d.get("projectGroups") or []
        group_id = groups[0].get("id") if groups else None
        return cls(
            id=d["id"],
            name=d.get("name") or f"<unnamed project {d['id']}>",
            code=d.get("code"),
            group_id=group_id,
            raw=d,
        )


@dataclass
class Specification:
    """A specification (document) inside a project."""
    id: int
    name: str
    code: str | None
    raw: dict

    @classmethod
    def from_dict(cls, d: dict) -> "Specification":
        return cls(
            id=d["id"],
            name=d.get("name") or f"<unnamed spec {d['id']}>",
            code=d.get("code"),
            raw=d,
        )


# ---------------------------------------------------------------------------
# The client itself
# ---------------------------------------------------------------------------

class VisureClient:
    """Thin wrapper around requests.Session for the Visure API.

    Use as a context manager so the session is logged out cleanly:

        with VisureClient() as client:
            client.authenticate()
            projects = client.list_projects_from_auth()
    """

    def __init__(self, base_url: str | None = None, timeout: int = 30):
        # base_url defaults to whatever's in .env. Allowing override makes
        # the class testable against a mock server.
        self.base_url = (base_url or settings.base_url).rstrip("/")
        self.timeout = timeout

        # requests.Session reuses the underlying TCP connection across calls.
        # For ~25 API calls in a row, this is a real speedup.
        self.session = requests.Session()

        # Filled in by authenticate()
        self._token: str | None = None
        self._user_payload: dict | None = None

    # ----- Context manager plumbing ----------------------------------------

    def __enter__(self) -> "VisureClient":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        # Try to log out cleanly. Don't blow up if it fails — we're exiting anyway.
        try:
            self.logout()
        except Exception:
            pass
        self.session.close()

    # ----- The actual HTTP plumbing ----------------------------------------

    def _url(self, path: str) -> str:
        """Join the base URL with an API path."""
        return f"{self.base_url}{path}"

    def _headers(self) -> dict[str, str]:
        """Build request headers, including the bearer token if we have one."""
        h = {"Accept": "application/json", "Content-Type": "application/json"}
        if self._token:
            h["Authorization"] = f"Bearer {self._token}"
        return h

    def _get(self, path: str, params: dict | None = None) -> Any:
        """All read calls go through here. Centralised error handling."""
        url = self._url(path)
        response = self.session.get(
            url,
            headers=self._headers(),
            params=params,
            timeout=self.timeout,
        )
        self._raise_for_status(response, method="GET", path=path)
        # Some endpoints return 200 with empty body. Guard against that.
        if not response.content:
            return None
        return response.json()

    def _post(self, path: str, json_body: dict | None = None) -> Any:
        """Used ONLY for /authenticate, /project/current, and /logout.
        No requirement-modifying POSTs exist anywhere in this client."""
        url = self._url(path)
        response = self.session.post(
            url,
            headers=self._headers(),
            json=json_body or {},
            timeout=self.timeout,
        )
        self._raise_for_status(response, method="POST", path=path)
        if not response.content:
            return None
        # Some Visure endpoints reply with no body but return 200 — handle that.
        try:
            return response.json()
        except ValueError:
            return None

    @staticmethod
    def _raise_for_status(response: requests.Response, method: str, path: str) -> None:
        """Turn HTTP errors into our exception types with useful messages."""
        if response.ok:
            return

        # Trim the body so the error stays readable
        body_preview = (response.text or "")[:400]

        if response.status_code in (401, 403):
            raise VisureAuthError(
                f"{method} {path} -> {response.status_code} {response.reason}\n"
                f"Body: {body_preview}\n"
                f"Likely cause: bad credentials, expired token, or insufficient permissions."
            )

        raise VisureAPIError(
            f"{method} {path} -> {response.status_code} {response.reason}\n"
            f"Body: {body_preview}"
        )

    # ----- Authentication --------------------------------------------------

    def authenticate(self) -> dict:
        """POST /api/v1/authenticate.

        Returns the full User14 payload. The payload contains:
          - accessToken.token   <- the bearer token for subsequent calls
          - accessToken.refreshToken
          - projects[]          <- the projects this account can see (with groups)

        We store the token on `self` for transparent use by later calls.
        """
        body = {
            "username": settings.username,
            "password": settings.password,
            "licenseType": settings.license_type,
        }
        payload = self._post("/api/v1/authenticate", json_body=body)

        if not payload:
            raise VisureAuthError("Authenticate returned empty response.")

        # The token can live in payload["accessToken"]["token"] OR at top level
        # depending on Visure version. We handle both.
        token = None
        access_token_obj = payload.get("accessToken")
        if isinstance(access_token_obj, dict):
            token = access_token_obj.get("token")
        token = token or payload.get("token")

        if not token:
            raise VisureAuthError(
                "Authenticated, but no token in response. Payload keys: "
                f"{list(payload.keys())}"
            )

        self._token = token
        self._user_payload = payload
        return payload

    def logout(self) -> None:
        """POST /api/v1/logout. Cleanly closes the session on the Visure side."""
        if not self._token:
            return
        try:
            self._post("/api/v1/logout")
        finally:
            self._token = None

    # ----- Projects --------------------------------------------------------

    def list_projects_from_auth(self) -> list[Project]:
        """Return the projects array from the /authenticate response.

        This is a freebie — the auth response already contains the project
        list with groups. No extra API call needed. If for some reason auth
        didn't include projects, fall back to GET /api/v1/projects.
        """
        if not self._user_payload:
            raise VisureAPIError("Must call authenticate() before listing projects.")

        raw_projects = self._user_payload.get("projects")
        if not raw_projects:
            # Fallback: explicit list endpoint
            raw_projects = self._get("/api/v1/projects") or []

        return [Project.from_dict(p) for p in raw_projects]

    def set_current_project(self, project: Project) -> None:
        """POST /api/v1/project/current with {project, group}.

        Visure is session-stateful: many read endpoints (like /specifications)
        return data for whichever project is currently selected. You have to
        explicitly pick one before listing its specs.
        """
        if project.group_id is None:
            raise VisureAPIError(
                f"Project {project.name!r} (id={project.id}) has no group on the "
                f"user record. Cannot set as current project."
            )
        body = {"project": project.id, "group": project.group_id}
        self._post("/api/v1/project/current", json_body=body)

    # ----- Specifications --------------------------------------------------

    def list_specifications(self) -> list[Specification]:
        """GET /api/v1/specifications.

        Returns specs for the CURRENTLY SELECTED project. Call
        set_current_project() first. This is the session quirk we documented
        in the README — Visure remembers the current project on the server side.
        """
        raw = self._get("/api/v1/specifications") or []
        return [Specification.from_dict(s) for s in raw]

    # ----- Requirements (elements) inside a spec ---------------------------

    def get_specification_items(
        self,
        spec_id: int,
        *,
        include_all_attributes: bool = True,
    ) -> list[dict]:
        """GET /api/v1/specification/{id}/items?includeAllAttributes=true.

        This is the workhorse endpoint. ONE call returns the whole hierarchy
        of elements (requirements + section headings) for a spec, with every
        user-defined attribute already baked into each element.

        The raw shape is a list of Element14 dicts. We hand the raw dicts to
        the excel_writer to flatten — keeping API parsing concerns separate
        from output-formatting concerns.
        """
        params = {
            "includeAllAttributes": "true" if include_all_attributes else "false",
        }
        return self._get(f"/api/v1/specification/{spec_id}/items", params=params) or []
