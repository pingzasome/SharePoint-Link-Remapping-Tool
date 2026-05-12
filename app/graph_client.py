from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from typing import Any
from urllib.parse import quote

import msal
import requests


GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]

logger = logging.getLogger(__name__)


@dataclass
class GraphConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    sharepoint_host: str
    sharepoint_site_path: str

    def validate(self) -> None:
        missing = [
            name
            for name, value in {
                "Tenant ID": self.tenant_id,
                "Client ID": self.client_id,
                "Client Secret": self.client_secret,
                "SharePoint Host": self.sharepoint_host,
                "SharePoint Site Path": self.sharepoint_site_path,
            }.items()
            if not value
        ]
        if missing:
            raise ValueError(f"Missing required configuration: {', '.join(missing)}")


class GraphClient:
    def __init__(self, config: GraphConfig):
        config.validate()
        self.config = config
        self.session = requests.Session()
        self._site_id: str | None = None
        self._drives: list[dict[str, str]] | None = None
        self._access_token_value: str | None = None
        self._access_token_expires_at = 0.0

    def search_site_files(self, filename: str) -> list[dict[str, str]]:
        if not filename:
            return []

        matches_by_key: dict[str, dict[str, str]] = {}
        encoded_query = quote(filename.replace("'", "''"), safe="")

        for drive in self.get_site_drives():
            url = f"{GRAPH_ROOT}/drives/{drive['id']}/root/search(q='{encoded_query}')"
            while url:
                payload = self._request("GET", url)
                for item in payload.get("value", []):
                    if "file" not in item:
                        continue
                    web_url = item.get("webUrl", "")
                    name = item.get("name", "")
                    match = {
                        "name": name,
                        "webUrl": web_url,
                        "id": item.get("id", ""),
                        "driveName": drive.get("name", ""),
                    }
                    matches_by_key[item.get("id") or web_url or f"{drive['id']}:{name}"] = match
                url = payload.get("@odata.nextLink")

        return list(matches_by_key.values())

    def get_site_id(self) -> str:
        if self._site_id:
            return self._site_id

        host = self.config.sharepoint_host.strip().strip("/")
        site_path = "/" + self.config.sharepoint_site_path.strip("/")
        encoded_path = quote(site_path, safe="/")
        url = f"{GRAPH_ROOT}/sites/{host}:{encoded_path}"
        payload = self._request("GET", url)
        self._site_id = payload["id"]
        return self._site_id

    def get_site_drives(self) -> list[dict[str, str]]:
        if self._drives is not None:
            return self._drives

        site_id = self.get_site_id()
        payload = self._request("GET", f"{GRAPH_ROOT}/sites/{site_id}/drives")
        self._drives = [
            {"id": item["id"], "name": item.get("name", "")}
            for item in payload.get("value", [])
            if item.get("id")
        ]
        if not self._drives:
            raise RuntimeError("No SharePoint document libraries/drives were found for this site.")
        return self._drives

    def _access_token(self) -> str:
        if self._access_token_value and time.time() < self._access_token_expires_at:
            return self._access_token_value

        authority = f"https://login.microsoftonline.com/{self.config.tenant_id}"
        app = msal.ConfidentialClientApplication(
            self.config.client_id,
            authority=authority,
            client_credential=self.config.client_secret,
        )
        result = app.acquire_token_for_client(scopes=SCOPES)
        if "access_token" not in result:
            error = result.get("error_description") or result.get("error") or "Unknown authentication error"
            raise RuntimeError(f"Microsoft Graph authentication failed: {error}")

        self._access_token_value = result["access_token"]
        self._access_token_expires_at = time.time() + max(int(result.get("expires_in", 3600)) - 300, 60)
        return self._access_token_value

    def _request(self, method: str, url: str, **kwargs: Any) -> dict[str, Any]:
        token = self._access_token()
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {token}"
        headers["Accept"] = "application/json"

        response = self.session.request(method, url, headers=headers, timeout=60, **kwargs)
        if response.status_code >= 400:
            logger.warning("Microsoft Graph request failed: %s", response.status_code)
            raise RuntimeError(f"Microsoft Graph error {response.status_code}: {_graph_error_message(response)}")
        return response.json()


def _graph_error_message(response: requests.Response) -> str:
    try:
        payload = response.json()
    except ValueError:
        return response.text[:500]

    error = payload.get("error", {})
    message = error.get("message") if isinstance(error, dict) else None
    return message or response.text[:500]
