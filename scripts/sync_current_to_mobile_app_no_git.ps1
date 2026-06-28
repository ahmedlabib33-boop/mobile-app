param(
    [string]$Owner = "ahmedlabib33-boop",
    [string]$Repo = "mobile-app",
    [string]$Branch = "main",
    [string]$Message = "No-Git sync mobile Streamlit app",
    [switch]$DryRun,
    [int]$ChunkSize = 500
)

$scriptPath = Join-Path $PSScriptRoot "sync_current_to_samco_no_git.ps1"
$arguments = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $scriptPath,
    "-Owner", $Owner,
    "-Repo", $Repo,
    "-Branch", $Branch,
    "-Message", $Message,
    "-ChunkSize", $ChunkSize
)
if ($DryRun) {
    $arguments += "-DryRun"
}
& powershell.exe @arguments
exit $LASTEXITCODE
