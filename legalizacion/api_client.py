"""Client for the Silogtran legalization API hosted on Render."""
from __future__ import annotations

from datetime import date
from io import StringIO
import time

import pandas as pd
import requests

from .config import settings

CROSSED_REPORT = "legalizacion-viaje-sap"
RETRY_STATUS = {502, 503, 504}


class ApiError(RuntimeError):
    pass


def _error_message(r: requests.Response) -> str:
    try:
        return r.json().get("error", r.text)
    except ValueError:
        return r.text[:500] or r.reason


class LegalizacionApi:
    def __init__(self, base_url: str | None = None, api_key: str | None = None, timeout: int | None = None):
        self.base_url = (base_url or settings.api_base_url).rstrip("/")
        self.api_key = api_key or settings.api_key
        self.timeout = timeout or settings.api_timeout
        if not self.api_key or self.api_key == "change-me":
            raise ApiError("API_KEY is missing. Copy .env.example to .env and set it.")
        self.session = requests.Session()
        self.session.headers["x-api-key"] = self.api_key

    def wake_up(self) -> dict:
        """Render free tier sleeps; /health wakes the service (public endpoint)."""
        r = self.session.get(f"{self.base_url}/health", timeout=self.timeout)
        r.raise_for_status()
        return r.json()

    def _get(self, path: str, params: dict, retries: int = 3) -> requests.Response:
        url = f"{self.base_url}{path}"
        for attempt in range(1, retries + 1):
            try:
                r = self.session.get(url, params=params, timeout=self.timeout)
            except (requests.ConnectionError, requests.Timeout) as e:
                if attempt == retries:
                    raise ApiError(f"No connection to API: {e}") from e
            else:
                if r.ok:
                    return r
                msg = _error_message(r)
                # 502 with a JSON body is a Silogtran/business error from our API: do not retry.
                app_error = r.headers.get("content-type", "").startswith("application/json")
                if r.status_code not in RETRY_STATUS or app_error or attempt == retries:
                    raise ApiError(f"API {r.status_code}: {msg}")
            time.sleep(5 * attempt)
        raise ApiError("Unreachable")

    def legalizacion_viaje_sap(self, start: date, end: date) -> pd.DataFrame:
        """Trip legalizations with SAP plate (no PEAJE ELECTRÓNICO, no voided).

        All columns are returned as text; numeric parsing happens in the engine.
        """
        if end < start:
            raise ApiError("End date is before start date.")
        params = {"fechaInicial": start.isoformat(), "fechaFinal": end.isoformat(), "format": "csv"}
        r = self._get(f"/api/reports/{CROSSED_REPORT}", params)
        text = r.content.decode("utf-8-sig")
        if not text.strip():
            return pd.DataFrame()
        df = pd.read_csv(StringIO(text), dtype=str, keep_default_na=False)
        df.columns = [c.strip() for c in df.columns]
        return df
