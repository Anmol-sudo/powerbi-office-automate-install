# Check for Administrative privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "Please run this script as an Administrator."
    exit
}

# 1. Define variables
$workDir = "C:\Office2024"
$xmlUrl = "https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/configuration.xml"
$odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17328-20162.exe"

# 2. Setup Environment
if (!(Test-Path $workDir)) { New-Item -ItemType Directory -Path $workDir }
Set-Location $workDir

# 3. Install Power BI via Winget
Write-Host "--- Installing Power BI Desktop ---" -ForegroundColor Cyan
winget install --id Microsoft.PowerBIDesktop --silent --accept-source-agreements --accept-package-agreements

# 4. Download Office Deployment Files
Write-Host "--- Downloading Office Deployment Files ---" -ForegroundColor Cyan
Invoke-WebRequest -Uri $odtUrl -OutFile "$workDir\odt.exe"
Invoke-WebRequest -Uri $xmlUrl -OutFile "$workDir\configuration.xml"

# 5. Extract and Install Office 2024
Write-Host "--- Extracting ODT ---" -ForegroundColor Cyan
Start-Process -FilePath "$workDir\odt.exe" -ArgumentList "/extract:$workDir /quiet" -Wait

Write-Host "--- Starting Office 2024 Installation ---" -ForegroundColor Gold
Write-Host "This will run in the background. Please wait..." -ForegroundColor Gray
Start-Process -FilePath "$workDir\setup.exe" -ArgumentList "/configure configuration.xml" -Wait

Write-Host "--- All tasks completed successfully! ---" -ForegroundColor Green