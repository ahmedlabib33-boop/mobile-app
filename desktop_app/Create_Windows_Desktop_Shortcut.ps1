$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$launcher = Join-Path $scriptDir "Launch_Project_Intelligence_Hub.ps1"
if (-not (Test-Path -LiteralPath $launcher)) {
    throw "Launcher not found: $launcher"
}

$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "Project Intelligence Hub.lnk"
$powershell = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $powershell
$shortcut.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$launcher`""
$shortcut.WorkingDirectory = $scriptDir
$shortcut.Description = "Project Intelligence Hub desktop app"

$edgeIcon = "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
if (Test-Path -LiteralPath $edgeIcon) {
    $shortcut.IconLocation = $edgeIcon
}

$shortcut.Save()
Write-Host "Desktop shortcut created:"
Write-Host $shortcutPath
