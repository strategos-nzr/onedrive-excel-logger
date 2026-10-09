"""
Bluesky callback that logs runs to an Excel workbook on OneDrive/SharePoint
via Microsoft Graph.

Usage:
    from onedrive_excel_logger import OneDriveExcelLogger
    logger = OneDriveExcelLogger(
        share_url="https://1drv.ms/x/s!...",      # or a SharePoint link
        client_id="<azure-app-client-id>",
        tenant_id="<work-or-school-tenant-id>",
        worksheet="Sheet1",
        table_name="BeamtimeRuns",
        beam_spot_size=lambda: f"{slits.hgap.get():.1f} x {slits.vgap.get():.1f} um",
    )
    RE.subscribe(logger)
"""
import base64
import logging
import os
import threading
from concurrent.futures import ThreadPoolExecutor
from urllib.parse import quote

import msal
import requests
from bluesky.callbacks import CallbackBase

log = logging.getLogger(__name__)
GRAPH = "https://graph.microsoft.com/v1.0"

# Logical field -> table header text (case-insensitive, whitespace-normalized)
DEFAULT_COLUMNS = {
    "uid": "UID",
    "scan_id": "ScanID",
    "num_frames": "Number of frames",
    "integration_time": "Integration time per frame",
    "detector": "Detector",
    "beam_spot_size": "Beam spot size",
}


def _encode_share_url(url):
    b64 = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    return "u!" + b64


class OneDriveExcelLogger(CallbackBase):
    def __init__(
        self,
        share_url,
        client_id,
        tenant_id="common",
        client_secret=None,
        username=None,
        worksheet="Sheet1",
        columns=None,
        integration_time_keys=("acquire_time", "exposure_time", "count_time"),
        beam_spot_size=None,
        stream_name="primary",
        token_cache_path="~/.bluesky_onedrive_token.json",
        blocking=False,
        *,
        table_name="BeamtimeRuns",
    ):
        super().__init__()
        if client_secret is not None:
            raise ValueError(
                "Excel table rows/add does not support application permissions "
                "(client_secret). Use delegated work/school Files.ReadWrite access."
            )
        if tenant_id.lower() == "consumers":
            raise ValueError(
                "Excel table rows/add does not support personal Microsoft accounts. "
                "Use a work/school account and tenant."
            )
        if not isinstance(table_name, str) or not table_name.strip():
            raise ValueError("table_name must name an existing Excel table")
        self.share_url = share_url
        self.worksheet = worksheet
        self.table_name = table_name
        self.columns = columns or DEFAULT_COLUMNS
        self.integration_time_keys = integration_time_keys
        self.beam_spot_size = beam_spot_size
        self.stream_name = stream_name
        self.blocking = blocking
        self._pool = ThreadPoolExecutor(max_workers=1)
        self._lock = threading.Lock()
        self._drive_item = None
        self._init_auth(client_id, tenant_id, client_secret, username, token_cache_path)
        self._reset()

    def _init_auth(self, client_id, tenant_id, client_secret, username, cache_path):
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._username = username
        self._cache_path = os.path.expanduser(cache_path)
        self._cache = msal.SerializableTokenCache()
        if os.path.exists(self._cache_path):
            with open(self._cache_path) as f:
                self._cache.deserialize(f.read())
        self._app = msal.PublicClientApplication(
            client_id, authority=authority, token_cache=self._cache)
        self._scopes = ["Files.ReadWrite"]

    def _save_cache(self):
        if self._cache.has_state_changed:
            with open(self._cache_path, "w") as f:
                f.write(self._cache.serialize())

    def _token(self):
        accounts = self._app.get_accounts(username=self._username)
        result = (self._app.acquire_token_silent(self._scopes, account=accounts[0])
                  if accounts else None)
        if not result:
            flow = self._app.initiate_device_flow(scopes=self._scopes)
            print(flow["message"])
            result = self._app.acquire_token_by_device_flow(flow)
        self._save_cache()
        if "access_token" not in result:
            raise RuntimeError(f"Graph auth failed: {result.get('error_description')}")
        return result["access_token"]

    def login(self):
        """Sign in and validate the existing table before starting the beamtime."""
        self._token()
        with self._lock:
            self._resolve_table()

    def _req(self, method, url, **kw):
        h = {"Authorization": f"Bearer {self._token()}"}
        r = requests.request(method, url, headers=h, timeout=30, **kw)
        r.raise_for_status()
        return r.json() if r.content else {}

    def _resolve_item(self):
        if self._drive_item is None:
            item = self._req("GET", f"{GRAPH}/shares/{_encode_share_url(self.share_url)}/driveItem")
            self._drive_item = (item["parentReference"]["driveId"], item["id"])
        d, i = self._drive_item
        return (
            f"{GRAPH}/drives/{quote(d, safe='')}/items/{quote(i, safe='')}"
            f"/workbook/worksheets/{quote(self.worksheet, safe='')}"
        )

    def _resolve_table(self):
        ws = self._resolve_item()
        table = self._req("GET", f"{ws}/tables/{quote(self.table_name, safe='')}")
        url = f"{ws}/tables/{quote(table['id'], safe='')}"
        columns = []
        page_url = f"{url}/columns"
        while page_url:
            page = self._req("GET", page_url)
            columns.extend(page["value"])
            page_url = page.get("@odata.nextLink")
        columns.sort(key=lambda column: column["index"])
        headers = [" ".join(column["name"].split()).casefold() for column in columns]
        mapping = {}
        for field, header in self.columns.items():
            normalized = " ".join(header.split()).casefold()
            matches = [index for index, name in enumerate(headers) if name == normalized]
            if not matches:
                raise ValueError(f"Required column {header!r} missing from table {self.table_name!r}")
            if len(matches) != 1 or matches[0] in mapping:
                raise ValueError(f"Ambiguous column {header!r} in table {self.table_name!r}")
            mapping[matches[0]] = field
        return url, headers, mapping

    def _append_row(self, record):
        with self._lock:
            table, headers, mapping = self._resolve_table()
            row = [record.get(mapping[index], "") if index in mapping else ""
                   for index in range(len(headers))]
            # Do not retry this POST: a timeout may occur after Graph has appended it.
            result = self._req("POST", f"{table}/rows/add",
                               json={"index": None, "values": [row]})
            index = result.get("index")
            if index is None and result.get("value"):
                index = result["value"][0].get("index")
            log.info("Logged run %s to Excel table %s (returned row index: %s)",
                     record.get("uid"), self.table_name, index)

    def _reset(self):
        self._start = None
        self._num_frames = 0
        self._primary_desc = set()
        self._int_time = None

    def start(self, doc):
        self._reset()
        self._start = doc

    def descriptor(self, doc):
        if doc.get("name") == self.stream_name:
            self._primary_desc.add(doc["uid"])
        if self._int_time is None:
            self._int_time = self._find_integration_time(doc)

    def event(self, doc):
        if doc["descriptor"] in self._primary_desc:
            self._num_frames += 1

    def event_page(self, doc):
        if doc["descriptor"] in self._primary_desc:
            self._num_frames += len(doc["seq_num"])

    def stop(self, doc):
        if self._start is None:
            return
        md = self._start
        n = (doc.get("num_events") or {}).get(self.stream_name, self._num_frames)
        record = {
            "uid": md["uid"],
            "scan_id": md.get("scan_id", ""),
            "num_frames": n,
            "integration_time": md.get("integration_time", self._int_time) or "",
            "detector": ", ".join(md.get("detectors", [])),
            "beam_spot_size": self._get_spot_size(md),
        }
        fut = self._pool.submit(self._safe_append, record)
        if self.blocking:
            fut.result()
        self._reset()

    def _find_integration_time(self, desc):
        for obj_cfg in (desc.get("configuration") or {}).values():
            for key, val in (obj_cfg.get("data") or {}).items():
                if any(key.endswith(k) for k in self.integration_time_keys):
                    return val
        return None

    def _get_spot_size(self, md):
        src = self.beam_spot_size
        if callable(src):
            try:
                return src()
            except Exception as e:
                log.warning("beam_spot_size callable failed: %s", e)
                return ""
        if src is not None:
            return src
        return md.get("beam_spot_size", "")

    def _safe_append(self, record):
        try:
            self._append_row(record)
        except Exception:
            log.exception("Failed to write run %s to OneDrive Excel", record.get("uid"))
