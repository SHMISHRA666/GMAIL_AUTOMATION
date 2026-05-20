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
    $output = & $ExePath @Args 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Executable command failed with exit code $LASTEXITCODE. Args: $($Args -join ' ')`n$output"
    }
    return ($output | Out-String)
}

function Test-UiStartup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExePath,
        [Parameter(Mandatory = $false)]
        [string]$Argument = "",
        [Parameter(Mandatory = $true)]
        [string]$Name
    )
    if ([string]::IsNullOrWhiteSpace($Argument)) {
        $process = Start-Process -FilePath $ExePath -PassThru
    }
    else {
        $process = Start-Process -FilePath $ExePath -ArgumentList $Argument -PassThru
    }
    Start-Sleep -Seconds 8
    if ($process.HasExited) {
        throw "$Name exited early with code $($process.ExitCode)"
    }
    Stop-Process -Id $process.Id -Force
    Write-Host "$Name startup check passed."
}

function New-PyInstallerArgs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$WindowMode
    )
    return @(
        "-m", "PyInstaller",
        "--noconfirm",
        "--onefile",
        "--specpath", "build",
        "--name", $Name,
        $WindowMode,
        "--collect-data", "flet",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "flet_desktop",
        "--exclude-module", "IPython",
        "--exclude-module", "matplotlib",
        "--exclude-module", "jedi",
        "--exclude-module", "parso",
        "--exclude-module", "pytest",
        "--exclude-module", "tkinter",
        "--add-data", ((Join-Path $repoRoot "Balance confirmation letter.docx") + ";gmail_automation/resources"),
        "--add-data", ((Join-Path $repoRoot "On Vendor letter.docx") + ";gmail_automation/resources"),
        "--add-data", ((Join-Path $repoRoot "Authorisation for Direct Balance Confirmation.pdf") + ";gmail_automation/resources"),
        "--add-data", ($fletArtifactPath + ";flet_desktop/app"),
        "run_gmail_automation.py"
    )
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

# Bundle Flet desktop runtime so modern UI does not download on first launch.
$fletCacheDir = (Invoke-Python -Args @(
        "-c",
        "import flet_desktop; from pathlib import Path; cache=flet_desktop.ensure_client_cached(); print(Path(cache).resolve())"
) | Select-Object -Last 1).Trim()
if (-not $fletCacheDir) {
    throw "Could not resolve Flet desktop cache directory."
}
if (-not (Test-Path $fletCacheDir)) {
    throw "Flet desktop cache directory not found: $fletCacheDir"
}
$fletClientExe = Join-Path $fletCacheDir "flet\flet.exe"
if (-not (Test-Path $fletClientExe)) {
    throw "Flet desktop runtime is incomplete (missing flet.exe): $fletClientExe"
}
$fletArtifactPath = Join-Path $repoRoot "build\flet-windows.zip"
if (Test-Path $fletArtifactPath) {
    Remove-Item $fletArtifactPath -Force
}
Write-Host "Bundling cached Flet desktop runtime from: $fletCacheDir"
Compress-Archive -Path (Join-Path $fletCacheDir "*") -DestinationPath $fletArtifactPath -CompressionLevel Optimal -Force
if (-not (Test-Path $fletArtifactPath)) {
    throw "Failed to create bundled Flet artifact: $fletArtifactPath"
}

$windowedExePath = Join-Path $repoRoot "dist\GmailConfirmationAutomation.exe"
$consoleExePath = Join-Path $repoRoot "dist\GmailConfirmationAutomationConsole.exe"
foreach ($path in @($windowedExePath, $consoleExePath)) {
    if (Test-Path $path) {
        Remove-Item $path -Force
    }
}

$windowedArgs = New-PyInstallerArgs -Name "GmailConfirmationAutomation" -WindowMode "--windowed"
$consoleArgs = New-PyInstallerArgs -Name "GmailConfirmationAutomationConsole" -WindowMode "--console"
if (-not $NoClean) {
    $windowedArgs += "--clean"
}

Write-Host "Building standalone windowed launcher (default)..."
Invoke-Python -Args $windowedArgs
Write-Host "Building standalone console launcher..."
Invoke-Python -Args $consoleArgs

if (-not (Test-Path $windowedExePath)) {
    throw "Build finished but windowed executable not found: $windowedExePath"
}
if (-not (Test-Path $consoleExePath)) {
    throw "Build finished but console executable not found: $consoleExePath"
}

if (-not $SkipSmokeTest) {
    Write-Host "Running smoke test: GmailConfirmationAutomationConsole.exe --help"
    $helpOutput = Invoke-Executable -ExePath $consoleExePath -Args @("--help")
    if ($helpOutput -notmatch "--modern-ui") {
        throw "Build validation failed: --modern-ui is missing from executable help."
    }
    if ($helpOutput -notmatch "--console") {
        throw "Build validation failed: --console is missing from executable help."
    }
}

if ($ThoroughValidate) {
    Write-Host "Running thorough runtime validation..."
    $tempRoot = Join-Path $env:TEMP ("gmail_auto_build_validate_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempRoot | Out-Null
    Push-Location $tempRoot
    try {
        Invoke-Executable -ExePath $consoleExePath -Args @("--init-config") | Out-Null
        $configPath = Join-Path $tempRoot "config.json"
        if (-not (Test-Path $configPath)) {
            throw "Thorough validation failed: config.json was not created."
        }

        Invoke-Executable -ExePath $consoleExePath -Args @("--init-db") | Out-Null
        $dbPath = Join-Path $env:APPDATA "GmailAutomation\compliance.db"
        if (-not (Test-Path $dbPath)) {
            throw "Thorough validation failed: database file not found at $dbPath"
        }

        $validationMaster = Join-Path $repoRoot $ValidationMasterPath
        if (Test-Path $validationMaster) {
            Invoke-Executable -ExePath $consoleExePath -Args @("--master", $validationMaster, "--mode", "validate", "--config", $configPath) | Out-Null
        }
        else {
            Write-Host "Validation workbook not found, skipping --mode validate: $validationMaster"
        }

        Test-UiStartup -ExePath $windowedExePath -Name "Default launch (modern UI)"
        Test-UiStartup -ExePath $windowedExePath -Argument "--ui" -Name "Legacy UI"
        Test-UiStartup -ExePath $windowedExePath -Argument "--modern-ui" -Name "Modern UI"
    }
    finally {
        Pop-Location
        Remove-Item -Recurse -Force $tempRoot -ErrorAction SilentlyContinue
    }
    Write-Host "Thorough runtime validation passed."
}

Write-Host "Build complete (windowed): $windowedExePath"
Write-Host "Build complete (console):  $consoleExePath"
