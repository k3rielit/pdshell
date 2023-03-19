# PDShell - github.com/k3rielit/pdshell

# Get the current script path and the ID of the current user
$scriptPath = $MyInvocation.MyCommand.Path
$currentUserId = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
Write-Host "[UID:$currentUserId] PDShell 0.2"

# Check if the script is running with administrator privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    # If not, restart it with elevated permissions
    Start-Process -FilePath "powershell.exe" -ArgumentList "-File",$scriptPath -Verb RunAs
    Exit
}

# If the script is running with administrator privileges, continue with the main logic
$filePath = Join-Path $PSScriptRoot "config.txt"
$desktopPath = "C:\Users\Public\Desktop"

Get-Content $filePath | ForEach-Object {
    # Skip commented out or empty lines
    if ($_.StartsWith("#")) {
        return
    }

    $splitLine = $_.Split(";")
    $file = $splitLine[0].Trim()
    $installer = $splitLine[1].Trim()
    $iconName = $splitLine[2].Trim()

    if ($file.StartsWith(".\")) {
        $file = Join-Path $PSScriptRoot ($file.Substring(2))
    }

    if ($installer.StartsWith(".\")) {
        $installer = Join-Path $PSScriptRoot ($installer.Substring(2))
    }

    Write-Host " > File: $file > Installer: $installer"

    if (!(Test-Path $file)) {
        Start-Process $installer -Wait
    }

    $targetPath = Join-Path $desktopPath $iconName
    $wshShell = New-Object -ComObject WScript.Shell
    $shortcut = $wshShell.CreateShortcut($targetPath + ".lnk")
    $shortcut.TargetPath = $file
    $shortcut.WorkingDirectory = (Split-Path $file -Parent)
    $shortcut.Save()
}

# Wait for input to exit
Write-Host -NoNewLine 'Done';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');