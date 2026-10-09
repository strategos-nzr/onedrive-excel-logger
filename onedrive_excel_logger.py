"""
Bluesky callback that logs runs to an Excel workbook on OneDrive/SharePoint
via Microsoft Graph.

Usage:
    from onedrive_excel_logger import OneDriveExcelLogger
    logger = OneDriveExcelLogger(
        share_url="https://1drv.ms/x/s!...",      # or a SharePoint link
        client_id="<azure-app-client-id>",
        tenant_id="<tenant-id or 'consumers'>",
        worksheet="Sheet1",
        beam_spot_size=lambda: f"{slits.hgap.get():.1f} x {slits.vgap.get():.1f} um",
    )
    RE.subscribe(logger)
"""
import base64
import logging
import os
import threading
from concurrent.futures import ThreadPoolExecutor

import msal
import requests
from bluesky.callbacks import CallbackBase

log = logging.getLogger(__name__)
GRAPH = "https://graph.microsoft.com/v1.0"

# Logical field -> header text expected in row 1 of the sheet (case-insensitive)
DEFAULT_COLUMNS = {
    "uid": "UID",
    "scan_id": "ScanID",
    "num_frames": "Number of frames",
    "integration_time": "Integration time per frame",
    "detector": "Detector",
    "beam_spot_size": "Beam spot size",
}


def _col_letter(idx):
    """0-based column index -> Excel letter (0 -> A, 27 -> AB)."""
    s, idx = "", idx + 1
    while idx:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


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
    ):
        super().__init__()
        self.share_url = share_url
        self.worksheet = worksheet
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
        if client_secret:
            self._app = msal.ConfidentialClientApplication(
                client_id, authority=authority, client_credential=client_secret)
            self._scopes = ["https://graph.microsoft.com/.default"]
            self._confidential = True
        else:
            self._cache_path = os.path.expanduser(cache_path)
            self._cache = msal.SerializableTokenCache()
            if os.path.exists(self._cache_path):
                self._cache.deserialize(open(self._cache_path).read())
            self._app = msal.PublicClientApplication(
                client_id, authority=authority, token_cache=self._cache)
            self._scopes = ["Files.ReadWrite"]
            self._confidential = False

    def _save_cache(self):
        if not self._confidential and self._cache.has_state_changed:
            with open(self._cache_path, "w") as f:
                f.write(self._cache.serialize())

    def _token(self):
        if self._confidential:
            result = self._app.acquire_token_for_client(scopes=self._scopes)
        else:
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
        """Call once at the start of the beamtime so device-code login happens up front."""
        self._token()
        self._resolve_item()

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
        return f"{GRAPH}/drives/{d}/items/{i}/workbook/worksheets/{self.worksheet}"

    def _append_row(self, record):
        with self._lock:
            ws = self._resolve_item()
            used = self._req("GET", f"{ws}/usedRange(valuesOnly=true)")
            values = used.get("values") or [[]]
            headers = [str(h).strip().lower() for h in values[0]]
            addr = used["address"].split("!")[-1].split(":")[0]
            start_row = int("".join(c for c in addr if c.isdigit()) or 1)
            next_row = start_row + used.get("rowCount", len(values))

            for field, header in self.columns.items():
                try:
                    col = headers.index(header.lower())
                except ValueError:
                    log.warning("Column %r not found in sheet header; skipping", header)
                    continue
                cell = f"{_col_letter(col)}{next_row}"
                self._req("PATCH", f"{ws}/range(address='{cell}')",
                          json={"values": [[record.get(field, "")]]})
            log.info("Logged run %s to Excel row %d", record.get("uid"), next_row)

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
