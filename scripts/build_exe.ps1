param(
    [switch]$SkipInstall,
    [switch]$SkipSmokeTest,
    [switch]$NoClean,
    [switch]$ThoroughValidate,
    [string]$ValidationMasterPath = "Information for External Balance Confirmations (1).xlsx"
)

$ErrorActionPreference = "Stop"

function Invoke-Python {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Args
    )

    if (Get-Command python -ErrorAction SilentlyContinue) {
        Write-Host "Using Python interpreter: python"
        & python @Args
        if ($LASTEXITCODE -ne 0) {
            throw "Python command failed with exit code $LASTEXITCODE"
        }
        return
    }

    if (Get-Command py -ErrorAction SilentlyContinue) {
        foreach ($version in @("3.13", "3.12", "3.11", "3.10", "3")) {
            & py "-$version" -c "import sys; print(sys.executable)" 1>$null 2>$null
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Using Python interpreter: py -$version"
                & py "-$version" @Args
                if ($LASTEXITCODE -ne 0) {
                    throw "Python command failed with exit code $LASTEXITCODE"
                }
                return
            }
        }
    }

    throw "No supported Python interpreter found. Install Python 3.10+ to build the executable."
}

function Invoke-Executable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath,
        [Parameter(Mandatory = $true)]
        [string[]]$Args
    )
    & $ExePath @Args
    if ($LASTEXITCODE -ne 0) {
        throw "Executable command failed with exit code $LASTEXITCODE. Args: $($Args -join ' ')"
    }
}

function Test-UiStartup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath,
        [Parameter(Mandatory = $true)]
        [string]$Argument,
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    $process = Start-Process -FilePath $ExePath -ArgumentList $Argument -PassThru
    Start-Sleep -Seconds 8
    if ($process.HasExited) {
        throw "$Name exited early with code $($process.ExitCode)"
    }
    Stop-Process -Id $process.Id -Force
    Write-Host "$Name startup check passed."
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Set-Location $repoRoot

$assetFiles = @(
    "Balance confirmation letter.docx",
    "On Vendor letter.docx",
    "Authorisation for Direct Balance Confirmation.pdf"
)

foreach ($asset in $assetFiles) {
    $assetPath = Join-Path $repoRoot $asset
    if (-not (Test-Path $assetPath)) {
        throw "Required build asset not found: $assetPath"
    }
}

if (-not $SkipInstall) {
    Write-Host "Installing/updating dependencies from requirements.txt..."
    Invoke-Python -Args @("-m", "pip", "install", "-r", "requirements.txt")
}

$pyInstallerArgs = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--onefile",
    "--specpath", "build",
    "--name", "GmailConfirmationAutomation",
    "--console",
    "--collect-data", "flet",
    "--collect-data", "gmail_automation",
    "--hidden-import", "win32com",
    "--hidden-import", "win32com.client",
    "--add-data", ((Join-Path $repoRoot "Balance confirmation letter.docx") + ";gmail_automation/resources"),
    "--add-data", ((Join-Path $repoRoot "On Vendor letter.docx") + ";gmail_automation/resources"),
    "--add-data", ((Join-Path $repoRoot "Authorisation for Direct Balance Confirmation.pdf") + ";gmail_automation/resources"),
    "run_gmail_automation.py"
)

if (-not $NoClean) {
    $pyInstallerArgs += "--clean"
}

Write-Host "Building standalone executable with PyInstaller..."
$exePath = Join-Path $repoRoot "dist\GmailConfirmationAutomation.exe"
if (Test-Path $exePath) {
    Remove-Item $exePath -Force
}
Invoke-Python -Args $pyInstallerArgs

if (-not (Test-Path $exePath)) {
    throw "Build finished but executable not found: $exePath"
}

if (-not $SkipSmokeTest) {
    Write-Host "Running smoke test: GmailConfirmationAutomation.exe --help"
    Invoke-Executable -ExePath $exePath -Args @("--help")
}

if ($ThoroughValidate) {
    Write-Host "Running thorough runtime validation..."
    $tempRoot = Join-Path $env:TEMP ("gmail_auto_build_validate_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempRoot | Out-Null
    Push-Location $tempRoot
    try {
        Invoke-Executable -ExePath $exePath -Args @("--init-config")
        $configPath = Join-Path $tempRoot "config.json"
        if (-not (Test-Path $configPath)) {
            throw "Thorough validation failed: config.json was not created."
        }

        Invoke-Executable -ExePath $exePath -Args @("--init-db")
        $dbPath = Join-Path $env:APPDATA "GmailAutomation\compliance.db"
        if (-not (Test-Path $dbPath)) {
            throw "Thorough validation failed: database file not found at $dbPath"
        }

        $validationMaster = Join-Path $repoRoot $ValidationMasterPath
        if (Test-Path $validationMaster) {
            Invoke-Executable -ExePath $exePath -Args @("--master", $validationMaster, "--mode", "validate", "--config", $configPath)
        }
        else {
            Write-Host "Validation workbook not found, skipping --mode validate: $validationMaster"
        }

        Test-UiStartup -ExePath $exePath -Argument "--ui" -Name "Legacy UI"
        Test-UiStartup -ExePath $exePath -Argument "--modern-ui" -Name "Modern UI"
    }
    finally {
        Pop-Location
        Remove-Item -Recurse -Force $tempRoot -ErrorAction SilentlyContinue
    }
    Write-Host "Thorough runtime validation passed."
}

Write-Host "Build complete: $exePath"
