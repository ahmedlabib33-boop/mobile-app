param(
    [int]$DebounceSeconds = 20,
    [string]$MessagePrefix = "Sync mobile app Excel and template data updates"
)

$scriptPath = Join-Path $PSScriptRoot "watch_excel_and_push_to_samco.ps1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -DebounceSeconds $DebounceSeconds -MessagePrefix $MessagePrefix
exit $LASTEXITCODE
