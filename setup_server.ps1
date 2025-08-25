# Run as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {  
    $arguments = "& '" + $myinvocation.mycommand.definition + "'"
    Start-Process powershell -Verb runAs -ArgumentList $arguments
    Break
}

Write-Host "=== Distributed Server Setup ===" -ForegroundColor Cyan

# 1. Set Execution Policy
Write-Host "Configuring execution policy..."
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# 2. Enable PSRemoting
Write-Host "Enabling PSRemoting..."
Enable-PSRemoting -Force

# 3. Configure WinRM
Write-Host "Configuring WinRM..."
winrm quickconfig -force
Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force

Write-Host "`n=== Setup Instructions ===" -ForegroundColor Green
Write-Host "1. Create a file named 'credentials.json' in this directory"
Write-Host "2. Add your worker configuration in this format:"
Write-Host @"
{
    "workers": [
        {
            "worker_id": "worker1",
            "ip": "WORKER_IP_ADDRESS",
            "username": "COMPUTERNAME\\USERNAME",
            "password": "PASSWORD"
        }
    ]
}
"@
Write-Host "`n3. Replace:"
Write-Host "   - WORKER_IP_ADDRESS: The IP of the worker machine"
Write-Host "   - COMPUTERNAME: The worker's computer name"
Write-Host "   - USERNAME: The username on the worker machine"
Write-Host "   - PASSWORD: The password for that user"
Write-Host "`n4. Run 'python distributed_runner.py' to start the server"

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")