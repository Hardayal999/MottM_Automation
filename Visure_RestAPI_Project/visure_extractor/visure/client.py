"""
client.py
=========
The Visure API client. This is the ONLY file in the project that touches
the Visure API. Everything else works with Python objects returned from here.

Design principles
-----------------
1. Read-only against requirements data. The POSTs we DO make are:
   - /authenticate            (sends credentials, gets a token)
   - /project/current         (session pin — no data changes)
   - /elements/linkeditems    (bulk READ — POST because it sends a list of IDs
                               in the body; doesn't write anything)
   - /logout                  (clean exit)
   No POST, PUT, or DELETE against requirements, attributes, or links.

2. Session-based. After authenticate() and set_current_project() the rest of
   the endpoints answer in the context of that project.

3. Every method raises a clear exception on failure. The caller decides what
   to do (retry, log, abort). No silent failures.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import requests

from visure.config import settings


# ---------------------------------------------------------------------------
# Custom exceptions
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
    """A Visure project as returned by /authenticate or /projects."""
    id: int
    name: str
    code: str | None
    group_id: int | None
    raw: dict

    @classmethod
    def from_dict(cls, d: dict) -> "Project":
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
# The client
# ---------------------------------------------------------------------------

class VisureClient:
    """Thin wrapper around requests.Session for the Visure API."""

    # Batch size for the bulk linked-items POST. The endpoint accepts a list
    # of element IDs in the body, but very large bodies can hit URL length
    # or timeout limits. 100 is a safe, fast batch size — adjust if needed.
    LINK_BATCH_SIZE = 100

    def __init__(self, base_url: str | None = None, timeout: int = 30):
        self.base_url = (base_url or settings.base_url).rstrip("/")
        self.timeout = timeout
        self.session = requests.Session()
        self._token: str | None = None
        self._user_payload: dict | None = None

    # ----- Context manager plumbing ----------------------------------------

    def __enter__(self) -> "VisureClient":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        try:
            self.logout()
        except Exception:
            pass
        self.session.close()

    # ----- HTTP plumbing ---------------------------------------------------

    def _url(self, path: str) -> str:
        return f"{self.base_url}{path}"

    def _headers(self) -> dict[str, str]:
        h = {"Accept": "application/json", "Content-Type": "application/json"}
        if self._token:
            h["Authorization"] = f"Bearer {self._token}"
        return h

    def _get(self, path: str, params: dict | None = None) -> Any:
        url = self._url(path)
        response = self.session.get(
            url, headers=self._headers(), params=params, timeout=self.timeout,
        )
        self._raise_for_status(response, method="GET", path=path)
        if not response.content:
            return None
        return response.json()

    def _post(self, path: str, json_body: Any = None) -> Any:
        """Used for /authenticate, /project/current, /logout, and the bulk
        linked-items READ endpoint. The bulk endpoint takes a JSON array, not
        a dict, so json_body is typed as Any."""
        url = self._url(path)
        response = self.session.post(
            url,
            headers=self._headers(),
            json=json_body if json_body is not None else {},
            timeout=self.timeout,
        )
        self._raise_for_status(response, method="POST", path=path)
        if not response.content:
            return None
        try:
            return response.json()
        except ValueError:
            return None

    @staticmethod
    def _raise_for_status(response: requests.Response, method: str, path: str) -> None:
        if response.ok:
            return
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
        """POST /api/v1/authenticate. Stores the bearer token on self."""
        body = {
            "username": settings.username,
            "password": settings.password,
            "licenseType": settings.license_type,
        }
        payload = self._post("/api/v1/authenticate", json_body=body)

        if not payload:
            raise VisureAuthError("Authenticate returned empty response.")

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
        if not self._user_payload:
            raise VisureAPIError("Must call authenticate() before listing projects.")

        raw_projects = self._user_payload.get("projects")
        if not raw_projects:
            raw_projects = self._get("/api/v1/projects") or []

        return [Project.from_dict(p) for p in raw_projects]

    def set_current_project(self, project: Project) -> None:
        """POST /api/v1/project/current with {project, group}.

        Visure is session-stateful: read endpoints (like /specifications)
        return data for whichever project is currently selected.
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
        """GET /api/v1/specifications for the currently selected project."""
        raw = self._get("/api/v1/specifications") or []
        return [Specification.from_dict(s) for s in raw]

    # ----- Elements --------------------------------------------------------

    def get_specification_items(
        self,
        spec_id: int,
        *,
        include_all_attributes: bool = True,
    ) -> list[dict]:
        """GET /api/v1/specification/{id}/items?includeAllAttributes=true.

        One call returns the whole hierarchy with every attribute baked in.
        """
        params = {
            "includeAllAttributes": "true" if include_all_attributes else "false",
        }
        return self._get(f"/api/v1/specification/{spec_id}/items", params=params) or []

    # ----- Traceability links ----------------------------------------------

    def get_link_types(self) -> list[dict]:
        """GET /api/v1/linktypes for the currently selected project.

        Returns the list of link type definitions (Derives, Verified By,
        Satisfies, etc). We don't strictly need this for the export — the
        link type name is already on each LinkedItem — but it's useful if
        you ever want to validate or filter by type in Power BI.
        """
        return self._get("/api/v1/linktypes") or []

    def get_linked_items_bulk(self, element_ids: list[int]) -> list[dict]:
        """POST /api/v1/elements/linkeditems with a JSON array of element IDs.

        Returns one LinkedItem14 dict per link found. Each link has:
            sourceItemID, targetItemID, linkType, isSuspect, direction,
            code (target's code), name (target's name), project, etc.

        We batch the IDs in groups of LINK_BATCH_SIZE to keep request bodies
        modest. The endpoint itself uses POST only because it needs to send a
        list in the body — it doesn't modify any data on the server.

        Empty input → empty output, no API call wasted.
        """
        if not element_ids:
            return []

        all_links: list[dict] = []
        batch_size = self.LINK_BATCH_SIZE
        for start in range(0, len(element_ids), batch_size):
            batch = element_ids[start:start + batch_size]
            result = self._post("/api/v1/elements/linkeditems", json_body=batch) or []
            all_links.extend(result)

        return all_links
