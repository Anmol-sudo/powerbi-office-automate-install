# Check for Administrative privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this PowerShell window as Administrator!" -ForegroundColor Red
    exit
}

# 1. Define Paths and Raw GitHub URLs
$desktopPath = [System.IO.Path]::Combine($env:USERPROFILE, "Desktop")
$workDir = Join-Path $desktopPath "office"
$xmlUrl = "https://raw.githubusercontent.com/Anmol-sudo/powerbi-office-automate-install/main/configuration.xml"
$setupUrl = "https://raw.githubusercontent.com/Anmol-sudo/powerbi-office-automate-install/main/setup.exe"

# 2. Create the 'office' folder on Desktop
if (!(Test-Path $workDir)) {
    New-Item -ItemType Directory -Path $workDir
    Write-Host "Created folder at $workDir" -ForegroundColor Green
}
Set-Location $workDir

# 3. Download files from GitHub
Write-Host "Downloading configuration and setup files..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $xmlUrl -OutFile "$workDir\configuration.xml"
Invoke-WebRequest -Uri $setupUrl -OutFile "$workDir\setup.exe"

# 4. Run Office Installation
Write-Host "Starting Office 2024 Installation..." -ForegroundColor Gold
Start-Process -FilePath "$workDir\setup.exe" -ArgumentList "/configure configuration.xml" -Wait

# 5. Install Power BI via Winget
Write-Host "Installing Power BI Desktop..." -ForegroundColor Cyan
winget install --id Microsoft.PowerBIDesktop --silent --accept-source-agreements --accept-package-agreements

Write-Host "All tasks completed!" -ForegroundColor Green