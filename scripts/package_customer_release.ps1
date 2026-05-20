param(
    [string]$ExePath = "dist\GmailConfirmationAutomation.exe",
    [string]$OutputRoot = "release",
    [switch]$BuildIfMissing,
    [string]$ReleaseName = "",
    [switch]$ForceRebuild
)

$ErrorActionPreference = "Stop"

function Invoke-Executable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string[]]$Args
    )
    & $Path @Args
    if ($LASTEXITCODE -ne 0) {
        throw "Executable command failed with exit code $LASTEXITCODE. Args: $($Args -join ' ')"
    }
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $repoRoot

$resolvedExePath = Join-Path $repoRoot $ExePath
$resolvedConsoleExePath = Join-Path $repoRoot "dist\GmailConfirmationAutomationConsole.exe"
if (-not (Test-Path $resolvedExePath)) {
    if ($BuildIfMissing) {
        Write-Host "Executable not found. Building first..."
        & (Join-Path $repoRoot "scripts\build_exe.ps1")
        if ($LASTEXITCODE -ne 0) {
            throw "Build failed while preparing customer release."
        }
    }
}

if ($ForceRebuild) {
    Write-Host "Force rebuild requested. Building executable from current branch..."
    & (Join-Path $repoRoot "scripts\build_exe.ps1") -SkipInstall
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed while preparing customer release."
    }
}

if (-not (Test-Path $resolvedExePath)) {
    throw "Executable not found: $resolvedExePath"
}
if (-not (Test-Path $resolvedConsoleExePath)) {
    throw "Console executable not found: $resolvedConsoleExePath"
}

$helpOutput = (& $resolvedConsoleExePath --help 2>&1 | Out-String)
if ($LASTEXITCODE -ne 0) {
    throw "Failed to run executable --help before packaging."
}
if ($helpOutput -notmatch "--modern-ui") {
    throw "Packaging aborted: executable does not include --modern-ui. Rebuild from the correct branch."
}

$releaseRoot = Join-Path $repoRoot $OutputRoot
New-Item -ItemType Directory -Path $releaseRoot -Force | Out-Null

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$releaseBaseName = ""
if ([string]::IsNullOrWhiteSpace($ReleaseName)) {
    $releaseBaseName = "GmailConfirmationAutomation_" + $stamp
}
else {
    $trimmedName = $ReleaseName.Trim()
    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $releaseBaseName = -join ($trimmedName.ToCharArray() | ForEach-Object {
            if ($invalidChars -contains $_) { "_" } else { $_ }
        })
    if ([string]::IsNullOrWhiteSpace($releaseBaseName)) {
        throw "ReleaseName is invalid after sanitization. Please provide at least one valid filename character."
    }
}

$releaseDir = Join-Path $releaseRoot $releaseBaseName
if (Test-Path $releaseDir) {
    Remove-Item -Path $releaseDir -Recurse -Force
}
New-Item -ItemType Directory -Path $releaseDir | Out-Null

$exeTargetPath = Join-Path $releaseDir "GmailConfirmationAutomation.exe"
Copy-Item -Path $resolvedExePath -Destination $exeTargetPath -Force
$consoleExeTargetPath = Join-Path $releaseDir "GmailConfirmationAutomationConsole.exe"
Copy-Item -Path $resolvedConsoleExePath -Destination $consoleExeTargetPath -Force

$tempConfigDir = Join-Path $env:TEMP ("gmail_auto_release_" + [guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempConfigDir | Out-Null
try {
    Push-Location $tempConfigDir
    Invoke-Executable -Path $consoleExeTargetPath -Args @("--init-config")
    Pop-Location

    $generatedConfigPath = Join-Path $tempConfigDir "config.json"
    if (-not (Test-Path $generatedConfigPath)) {
        throw "Could not generate sample config.json from executable."
    }
    Copy-Item -Path $generatedConfigPath -Destination (Join-Path $releaseDir "config.sample.json") -Force
}
finally {
    if ((Get-Location).Path -eq $tempConfigDir) {
        Pop-Location
    }
    Remove-Item -Path $tempConfigDir -Recurse -Force -ErrorAction SilentlyContinue
}

$quickstartPath = Join-Path $releaseDir "README_QUICKSTART.txt"
$quickstart = @"
Gmail Confirmation Automation - Customer Quick Start

1) Place the executable and config.sample.json in a working folder.
2) Rename config.sample.json to config.json and update sender_email and app_password.
3) Keep your master workbook in the same folder (or pass full path with --master).

Useful commands:
- Create a fresh config template:
  .\GmailConfirmationAutomation.exe --init-config

- Validate workbook/templates only:
  .\GmailConfirmationAutomation.exe --master "Information for External Balance Confirmations (1).xlsx" --mode validate

- Generate documents without sending:
  .\GmailConfirmationAutomation.exe --master "Information for External Balance Confirmations (1).xlsx" --mode preview

- Open modern compliance UI (default):
  .\GmailConfirmationAutomation.exe

- Open console mode for CLI output:
  .\GmailConfirmationAutomation.exe --console --mode validate

- Open legacy desktop UI:
  .\GmailConfirmationAutomation.exe --ui

- Open modern compliance UI:
  .\GmailConfirmationAutomation.exe --modern-ui

Notes:
- Python is NOT required on customer machines.
- Microsoft Word is required only for PDF conversion when convert_to_pdf=true.
- Database is created automatically at:
  %APPDATA%\GmailAutomation\compliance.db
"@
$quickstart | Set-Content -Path $quickstartPath -Encoding UTF8

$zipPath = Join-Path $releaseRoot ($releaseBaseName + ".zip")
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}
Compress-Archive -Path (Join-Path $releaseDir "*") -DestinationPath $zipPath -CompressionLevel Optimal -Force

Write-Host "Customer release folder: $releaseDir"
Write-Host "Customer release zip:    $zipPath"
