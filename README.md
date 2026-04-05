# Excel Form App

A desktop app built with CustomTkinter to manage Heat Number records in Excel.

It has:
- an Insert tab for adding rows
- a Search tab for finding and updating rows
- a Customize Fields panel for column setup

## Security

- Customize Fields is password-protected.
- Local credentials are stored in `secrets.py`.
- Keep `secrets.py` out of version control.

## Requirements

- Python 3.10+ (tested on 3.13)
- `customtkinter`
- `openpyxl`

Install dependencies:

```bash
pip install customtkinter openpyxl
```

Run locally:

```bash
python main.py
```

## Client Delivery (Windows)

This project is packaged in PyInstaller one-folder mode.

Do not send only the raw app executable from dist. Send the installer executable generated in installer-output. If only the dist EXE is sent, Windows will show errors like:

- `failed to load python dll`
- `LoadLibrary: impossible to find the specified module`

For client delivery, send only:

- `installer-output\ExcelFormSetup.exe`

## Maintainer Release Steps

### 1) Build app binaries

```powershell
pyinstaller .\ExcelForm.spec --clean
```

Expected output includes:

- `dist\ExcelForm\ExcelForm.exe`
- `dist\ExcelForm\_internal\python313.dll`

### 2) Build installer

Compile `ExcelForm.iss` with Inno Setup (`ISCC`):

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" .\ExcelForm.iss
```

Output:

- `installer-output\ExcelFormSetup.exe`

### 3) Optional signing (recommended for production)

```powershell
$thumb = "YOUR_CERT_SHA1_THUMBPRINT"
$ts = "http://timestamp.digicert.com"

signtool sign /sha1 $thumb /fd SHA256 /tr $ts /td SHA256 .\dist\ExcelForm\ExcelForm.exe
signtool sign /sha1 $thumb /fd SHA256 /tr $ts /td SHA256 .\installer-output\ExcelFormSetup.exe

signtool verify /pa /v .\dist\ExcelForm\ExcelForm.exe
signtool verify /pa /v .\installer-output\ExcelFormSetup.exe
```

### 4) One-click release script

Signed build:

```powershell
.\release.ps1 -CertThumbprint "YOUR_CERT_SHA1_THUMBPRINT"
```

Unsigned test build:

```powershell
.\release.ps1 -SkipSign
```

Skip signature verification:

```powershell
.\release.ps1 -CertThumbprint "YOUR_CERT_SHA1_THUMBPRINT" -SkipVerify
```

## macOS Release

Windows `.exe` files do not run on macOS.
For macOS clients, build and send a macOS app bundle zip.

### Local macOS build

Run this on a Mac:

```bash
chmod +x ./release-macos.sh
./release-macos.sh
```

Output file to send:

- `insta ller-output/ExcelForm-macOS.zip`

### GitHub Actions build

Workflow file:

- `.github/workflows/release-macos.yml`

It runs on `macos-latest` and uploads this artifact:

- `ExcelForm-macOS` (contains `installer-output/ExcelForm-macOS.zip`)

## Workbook Behavior

- Workbook path and sheet settings are managed in Workbook Settings from the app.
- App-managed records are written starting at row 3 in the output sheet.
- Search reads from the app-managed region starting at row 3.

## Troubleshooting

Cannot save workbook:
Close the workbook in Excel, then save again.

Description not showing immediately:
Use Refresh in Search, or reopen workbook if external recalculation is pending.

Rows appear in Search but not where expected:
App writes to row 3 onward by design.

Field customization issues:
If `Description` is enabled, ensure `ItemCode` exists.
