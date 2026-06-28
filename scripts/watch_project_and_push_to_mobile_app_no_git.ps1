param(
    [int]$DebounceSeconds = 25,
    [string]$MessagePrefix = "Auto sync Project Intelligence Hub mobile app",
    [string]$Owner = "ahmedlabib33-boop",
    [string]$Repo = "mobile-app",
    [string]$Branch = "main"
)

$scriptPath = Join-Path $PSScriptRoot "watch_project_and_push_to_samco_no_git.ps1"
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -DebounceSeconds $DebounceSeconds -MessagePrefix $MessagePrefix -Owner $Owner -Repo $Repo -Branch $Branch
exit $LASTEXITCODE
