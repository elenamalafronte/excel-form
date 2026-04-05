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

### What to send to client IT (managed laptops)

If endpoint protection still blocks installation, send IT:

1. signed installer file name: `ExcelFormSetup.exe`
2. publisher name from your code-signing certificate
3. SHA256 hash:

```powershell
Get-FileHash .\installer-output\ExcelFormSetup.exe -Algorithm SHA256
```

Ask IT to allow by publisher certificate rule (preferred) or by hash rule.

## Workbook Notes

The app expects an Excel workbook with:
- source sheet: `CREXPD01`
- form/output sheet: `Heat Number`

The app uses `EXCEL_FILE` in `config.py` to decide which file to read/write.

On a first run, no workbook is loaded by default. After you choose a workbook in Workbook Settings, that selection is saved in the user config and reused the next time the app starts.

Data placement behavior:
- App-managed records are written starting from row 3 in the form/output sheet.
- Search reads from the app-managed table region that starts at row 3.
- This avoids mixing app rows with deep template/history ranges in large client workbooks.

## Troubleshooting

Cannot save workbook:
Close the workbook in Excel, then save again.

Description not showing immediately:
Use Refresh in Search, or reopen workbook if external recalculation is pending.

Rows appear in Search but not where expected:
App writes to row 3 onward by design.

Field customization issues:
If `Description` is enabled, ensure `ItemCode` exists.
