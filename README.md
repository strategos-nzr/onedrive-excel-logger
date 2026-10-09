# onedrive-excel-logger

Bluesky callback for autofilling fields into a Microsoft OneDrive Excel spreadsheet.

## What it does

This package provides a Bluesky `RunEngine` callback that writes one row per completed run into an Excel workbook stored on OneDrive or SharePoint via the Microsoft Graph API.

It can autofill the following fields:

- UID
- ScanID
- Number of frames
- Integration time per frame
- Detector
- Beam spot size

## How it works

- `start` document: records the run UID, ScanID, detector metadata, and plan metadata.
- `descriptor` documents: looks for integration time in detector configuration or readings.
- `event` documents: counts frames in the `primary` stream.
- `stop` document: finds the next empty row in the Excel sheet and writes the collected values.

## Requirements

Install dependencies:

```bash
pip install msal requests bluesky
```

You also need an Azure / Entra ID app registration.

### Authentication options

#### Interactive / delegated access
Use this for personal or work OneDrive with device-code sign-in.

Required permissions:

- `Files.ReadWrite`
- `Sites.ReadWrite.All` if the workbook is stored in SharePoint

#### App-only access
Use this for unattended automation with a client secret.

- Requires `https://graph.microsoft.com/.default`
- Your tenant admin must grant consent

## Notes

### Sheet setup

The first row of the sheet should contain the headers:

- `UID`
- `ScanID`
- `Number of frames`
- `Integration time per frame`
- `Detector`
- `Beam spot size`

The columns can be in any order, and header matching is case-insensitive.

### Integration time

The callback first checks `md['integration_time']`. If that is missing, it looks in detector configuration for keys ending in:

- `acquire_time`
- `exposure_time`
- `count_time`

If your detector uses a different key, you can customize the matching logic.

### Number of frames

This is counted from the `primary` stream in the stop document. If the stop document does not include `num_events`, the callback falls back to counting event documents directly.

### Beam spot size

You can provide this as:

- a static string
- a callable that returns the current beam spot size
- metadata in the run document

### Credentials

Do not hard-code secrets into source files. Load them from environment variables, a secrets manager, or a keyring.

### Concurrency

Writes happen after the run finishes. If the workbook is being edited while scans are running, consider using an Excel table instead of raw sheet cells for safer concurrent writes.

## Example usage

```python
from onedrive_excel_logger import OneDriveExcelLogger

xl = OneDriveExcelLogger(
    share_url="https://<tenant>-my.sharepoint.com/:x:/g/personal/.../Exxxx",
    client_id="00000000-0000-0000-0000-000000000000",
    tenant_id="<your-tenant-id>",
    worksheet="Sheet1",
    beam_spot_size=lambda: f"{ss.hgap.get():.0f} x {ss.vgap.get():.0f} µm",
)

xl.login()
RE.subscribe(xl)
```

## License

Add a license file if you plan to distribute this beyond your beamtime setup.
