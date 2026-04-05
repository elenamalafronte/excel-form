[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$CertThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$TimeStampUrl = "http://timestamp.digicert.com",

    [Parameter(Mandatory = $false)]
    [switch]$SkipSign,

    [Parameter(Mandatory = $false)]
    [switch]$SkipVerify
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Get-CommandPath {
    param([string]$Name)

    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if ($null -eq $cmd) {
        return $null
    }

    return $cmd.Source
}

function Get-IsccPath {
    $fromPath = Get-CommandPath -Name "ISCC.exe"
    if ($fromPath) {
        return $fromPath
    }

    $candidates = @(
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
        "C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        "C:\Program Files\Inno Setup 6\ISCC.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Sign-File {
    param(
        [string]$SignTool,
        [string]$Thumbprint,
        [string]$Timestamp,
        [string]$FilePath
    )

    & $SignTool sign /sha1 $Thumbprint /fd SHA256 /tr $Timestamp /td SHA256 $FilePath
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Push-Location $scriptRoot

try {
    Write-Step "Validating required tools"

    $pyInstaller = Get-CommandPath -Name "pyinstaller.exe"
    if (-not $pyInstaller) {
        throw "PyInstaller not found. Install it and ensure pyinstaller.exe is on PATH."
    }

    $iscc = Get-IsccPath
    if (-not $iscc) {
        throw "Inno Setup Compiler (ISCC.exe) not found. Install Inno Setup 6 or add ISCC.exe to PATH."
    }

    $signTool = Get-CommandPath -Name "signtool.exe"
    if (-not $SkipSign -and -not $signTool) {
        throw "signtool.exe not found. Install Windows SDK and ensure signtool is on PATH, or run with -SkipSign."
    }

    if (-not $SkipSign -and [string]::IsNullOrWhiteSpace($CertThumbprint)) {
        throw "Signing enabled but no certificate thumbprint provided. Use -CertThumbprint or run with -SkipSign."
    }

    Write-Step "Building app with PyInstaller"
    & $pyInstaller .\ExcelForm.spec --clean

    $appExe = Join-Path $scriptRoot "dist\ExcelForm\ExcelForm.exe"
    if (-not (Test-Path $appExe)) {
        throw "Build did not produce expected app executable at $appExe"
    }

    Write-Step "Building installer with Inno Setup"
    & $iscc .\ExcelForm.iss

    $installerExe = Join-Path $scriptRoot "installer-output\ExcelFormSetup.exe"
    if (-not (Test-Path $installerExe)) {
        throw "Build did not produce expected installer at $installerExe"
    }

    if (-not $SkipSign) {
        Write-Step "Signing app executable"
        Sign-File -SignTool $signTool -Thumbprint $CertThumbprint -Timestamp $TimeStampUrl -FilePath $appExe

        Write-Step "Signing installer executable"
        Sign-File -SignTool $signTool -Thumbprint $CertThumbprint -Timestamp $TimeStampUrl -FilePath $installerExe
    }
    else {
        Write-Step "Skipping signing (-SkipSign provided)"
    }

    if (-not $SkipVerify -and -not $SkipSign) {
        Write-Step "Verifying signatures"
        & $signTool verify /pa /v $appExe
        & $signTool verify /pa /v $installerExe
    }
    elseif (-not $SkipVerify -and $SkipSign) {
        Write-Step "Skipping verify because signing was skipped"
    }
    else {
        Write-Step "Skipping signature verification (-SkipVerify provided)"
    }

    Write-Step "Computing installer SHA256"
    $hash = Get-FileHash $installerExe -Algorithm SHA256
    Write-Host "Installer: $installerExe" -ForegroundColor Green
    Write-Host "SHA256:   $($hash.Hash)" -ForegroundColor Green

    Write-Host "`nRelease completed successfully." -ForegroundColor Green
}
catch {
    Write-Error $_
    exit 1
}
finally {
    Pop-Location
}
