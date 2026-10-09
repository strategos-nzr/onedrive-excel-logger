# onedrive-excel-logger

Bluesky callback for autofilling fields into a Microsoft OneDrive Excel spreadsheet.

## What it does

This module provides a Bluesky `RunEngine` callback that appends one row per completed run to an existing named Excel table in a workbook stored on OneDrive for Business or SharePoint via the Microsoft Graph API.

It can autofill the following fields:

- UID
- ScanID
- Number of frames
- Integration time per frame
- Detector
- Beam spot size

## How it works

- `start` document: records the run UID, ScanID, detector metadata, and plan metadata.
- `descriptor` documents: looks for integration time in detector configuration (not event readings).
- `event` documents: counts frames in the `primary` stream.
- `stop` document: resolves the table and its ordered headers, then submits one `POST .../tables/{id}/rows/add` with `{"index": null, "values": [row]}`. Graph appends the row at the end of the table; the callback never calculates worksheet cell addresses or PATCHes existing cells.

## Requirements

Install dependencies:

```bash
pip install msal requests bluesky
```

You also need an Azure / Entra ID app registration.

### Supported authentication

Use **delegated work/school account access** with device-code sign-in. Configure your app registration for work/school accounts, enable public client flows, and grant the delegated Microsoft Graph permission `Files.ReadWrite` (with consent as required by your organization). The signed-in account must have write access to the workbook.

The official [table rows/add documentation](https://learn.microsoft.com/en-us/graph/api/tablerowcollection-add?view=graph-rest-1.0) lists:

- Delegated work/school accounts: `Files.ReadWrite` supported.
- Delegated personal Microsoft accounts: **not supported**.
- Application permissions (app-only/client-secret authentication): **not supported**.

Do not use personal OneDrive or client-secret authentication for these workbook table writes. The constructor retains the existing `client_secret` argument for compatibility but raises a clear error if it is supplied; `tenant_id="consumers"` is also rejected. Use your work/school tenant ID (or `"organizations"`), and sign in with a work/school account.

## Notes

### Excel table setup

1. Open the workbook in Excel, on the worksheet you want to log to (for example `Sheet1`).
2. Enter the following headers in separate columns:

   - `UID`
   - `ScanID`
   - `Number of frames`
   - `Integration time per frame`
   - `Detector`
   - `Beam spot size`

3. Select the header cells and any existing run rows you want included in the table. Choose **Insert > Table** and check **My table has headers**.
4. Select a cell in the table. In **Table Design > Table Name** (or the table naming control in Excel for the web), name it `BeamtimeRuns`.
5. Save the workbook to OneDrive for Business or SharePoint. Supply its sharing link, worksheet name, and table name to the constructor.
6. Call `xl.login()` before subscribing or scanning. It signs in and validates the table and required headers without adding a row.

The table must already exist on the selected worksheet: the logger does not create one or fall back to raw worksheet writes. `table_name` defaults to `BeamtimeRuns` and is a new keyword-only argument; existing positional arguments retain their meaning.

### Header mapping

The actual table column order determines the submitted row order. Header matching ignores case, strips leading/trailing whitespace, and collapses repeated whitespace. Columns may be reordered. Every header in `columns` is required and must match exactly one table column; ambiguous required headers or multiple configured fields targeting one column are errors, not partial writes. Validation runs at login and again before each append.

To use different header text, pass a logical-field-to-header mapping, for example `columns={"uid": "Run UID", "scan_id": "Scan number"}`. Omitted logical fields are not logged. Extra table columns receive empty strings in each new row; they are not populated by the logger. Existing table rows are never targeted by its write requests.

### Integration time

The callback first checks `md['integration_time']`. If that is missing, it looks in detector configuration for keys ending in:

- `acquire_time`
- `exposure_time`
- `count_time`

If your detector uses a different key, pass `integration_time_keys=("your_key_suffix",)`. Event readings are not searched for integration time.

### Number of frames

This is counted from the `primary` stream in the stop document. If the stop document does not include `num_events`, the callback falls back to counting event documents directly.

### Beam spot size

You can provide this as:

- a static string
- a callable that returns the current beam spot size
- metadata in the run document

### Credentials

Device-code sign-in caches tokens in `~/.bluesky_onedrive_token.json` (customizable with `token_cache_path`). Protect this file as a credential and do not commit it or share it.

### Consecutive appends and limitations

Writes happen after each run finishes, in a background single-worker queue protected by a lock. Each successful append adds a new row at the table's end, so successive runs occupy consecutive new rows without overwriting previous run rows. There is no client-side incrementing row counter. The returned table row index is logged when Graph provides it. Set `blocking=True` to wait for each write attempt.

Other logger instances, external writers, or manual edits may interleave rows; ordering is guaranteed only within this logger's queue, not across writers. Avoid changing the table schema while logging. Append-only is **not an exactly-once guarantee**: after a timeout the server may already have appended the row, so the logger does not automatically retry append POSTs. Check the table by run UID before any manual retry to avoid duplicates. Failed writes are logged and do not stop scanning; they are not durably queued or automatically recovered.

## Example usage

```python
from onedrive_excel_logger import OneDriveExcelLogger

xl = OneDriveExcelLogger(
    share_url="https://<tenant>-my.sharepoint.com/:x:/g/personal/.../Exxxx",
    client_id="00000000-0000-0000-0000-000000000000",
    tenant_id="<your-tenant-id>",
    worksheet="Sheet1",
    table_name="BeamtimeRuns",
    beam_spot_size=lambda: f"{ss.hgap.get():.0f} x {ss.vgap.get():.0f} µm",
)

xl.login()  # device-code sign-in and table/header validation before scanning
RE.subscribe(xl)
```

## Tests

After installing the runtime dependencies above, run the standard-library mocked unit tests:

```bash
python -m unittest discover -v
```

Tests mock authentication and Graph HTTP requests; no credentials or live workbook are required.

## License

Add a license file if you plan to distribute this beyond your beamtime setup.
