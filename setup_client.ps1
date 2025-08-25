# Run as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

Write-Host "=== Distributed Worker Setup ===" -ForegroundColor Cyan

# 1. Set Execution Policy
Write-Host "Configuring execution policy..."
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# 2. Enable PSRemoting
Write-Host "Enabling PSRemoting..."
Enable-PSRemoting -Force

# 3. Set Network Profile to Private
Write-Host "Configuring network profile..."
Get-NetConnectionProfile | ForEach-Object { 
    Set-NetConnectionProfile -InterfaceIndex $_.InterfaceIndex -NetworkCategory Private -ErrorAction SilentlyContinue
}

# 4. Configure WinRM
Write-Host "Configuring WinRM..."
winrm quickconfig -force
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force

# 5. Configure Firewall
Write-Host "Configuring firewall..."
netsh advfirewall firewall add rule name="WinRM-HTTP" dir=in localport=5985 protocol=TCP action=allow

# 6. Restart WinRM Service
Write-Host "Restarting services..."
Restart-Service WinRM

Write-Host "`n=== Setup Instructions ===" -ForegroundColor Green
Write-Host "1. Create a directory for your client"
Write-Host "2. Create a file named 'worker_start.bat' in that directory"
Write-Host "3. Copy these files to the same directory:"
Write-Host "   - main.exe"
Write-Host "   - ASAMD.exe"
Write-Host "   - input.txt"
Write-Host "   - input_file.xlsx"
Write-Host "`nPress any key to view the worker_start.bat content..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# Display worker_start.bat content
Write-Host "`n=== worker_start.bat Content ===" -ForegroundColor Yellow
@"
@echo off
:: Change to the directory where this batch file is located
cd /d "%~dp0"

:: Check if required files exist
if not exist "main.exe" (
    echo Error: main.exe not found in %CD%
    pause
    exit /b 1
)

if not exist "ASAMD.exe" (
    echo Error: ASAMD.exe not found in %CD%
    pause
    exit /b 1
)

if not exist "input.txt" (
    echo Error: input.txt not found in %CD%
    pause
    exit /b 1
)

if not exist "input_file.xlsx" (
    echo Error: input_file.xlsx not found in %CD%
    pause
    exit /b 1
)

echo Starting worker process at %TIME%
start "" "main.exe" "input.txt" "input_file.xlsx"
echo Worker started at %TIME%
"@ | Write-Host -ForegroundColor White

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")