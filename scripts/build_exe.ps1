$ErrorActionPreference = "Stop"

python -m pip install -r requirements.txt
python -m PyInstaller GmailAutomation.spec --clean --noconfirm

Write-Host "Build complete. Executable should be in dist\\GmailConfirmationAutomation.exe"
