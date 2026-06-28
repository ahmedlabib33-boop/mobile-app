param(
    [int]$DebounceSeconds = 25,
    [string]$MessagePrefix = "Auto sync Project Intelligence Hub mobile app",
    [string]$Owner = "ahmedlabib33-boop",
    [string]$Repo = "mobile-app",
    [string]$Branch = "main"
)

$ErrorActionPreference = "Stop"

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$syncScript = Join-Path $PSScriptRoot "sync_current_to_samco_no_git.ps1"
$logPath = Join-Path $root "11-outputs\logs\pih_mobile_app_no_git_auto_sync_watcher.log"
New-Item -ItemType Directory -Force -Path (Split-Path $logPath -Parent) | Out-Null

$ignoredPathFragments = @(
    "\.git\",
    "\.venv\",
    "\venv\",
    "\.pih_mobile_app_sync_state\",
    "\__pycache__\",
    "\node_modules\",
    "\.pytest_cache\",
    "\.mypy_cache\",
    "\.ruff_cache\",
    "\.vite\",
    "\generated_outputs\logs\",
    "\11-outputs\logs\",
    "\.streamlit\delay_tia_upload_cache\"
)

$ignoredNames = @(
    "~$*",
    "*.pyc",
    "*.pyo",
    "*.log",
    "*.err",
    "*.out",
    "*.tmp",
    ".env",
    ".env.*"
)

$pending = $false
$lastChange = Get-Date
$pendingPaths = New-Object System.Collections.Generic.HashSet[string]

function Write-SyncLog([string]$Message) {
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$stamp] $Message"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line
}

function Test-SyncCandidate([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }
    $normalized = $Path.Replace("/", "\")
    foreach ($fragment in $ignoredPathFragments) {
        if ($normalized.IndexOf($fragment, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $false
        }
    }
    $name = [System.IO.Path]::GetFileName($Path)
    foreach ($pattern in $ignoredNames) {
        if ($name -like $pattern) {
            return $false
        }
    }
    return $true
}

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $root
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true
$watcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, DirectoryName, LastWrite, Size'

foreach ($eventName in @("Changed", "Created", "Deleted", "Renamed")) {
    Register-ObjectEvent -InputObject $watcher -EventName $eventName -Action {
        $eventPath = $Event.SourceEventArgs.FullPath
        if (Test-SyncCandidate $eventPath) {
            $script:pending = $true
            $script:lastChange = Get-Date
            [void]$script:pendingPaths.Add($eventPath)
            Write-SyncLog "Detected $($Event.SourceEventArgs.ChangeType): $eventPath"
        }
    } | Out-Null
}

Write-SyncLog "Started no-Git full-project watcher for $root"
Write-SyncLog "Target repository: $Owner/$Repo branch $Branch"
Write-SyncLog "Credential rule: set PIH_MOBILE_APP_GITHUB_TOKEN with repo write access before the first sync."

while ($true) {
    Start-Sleep -Seconds 2
    if ($pending -and ((Get-Date) - $lastChange).TotalSeconds -ge $DebounceSeconds) {
        $pending = $false
        $changed = @($pendingPaths)
        $pendingPaths.Clear()
        try {
            $message = "$MessagePrefix - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            Write-SyncLog "Running no-Git sync for $($changed.Count) detected path(s): $message"
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -Owner $Owner -Repo $Repo -Branch $Branch -Message $message *>> $logPath
            Write-SyncLog "No-Git sync completed."
        } catch {
            Write-SyncLog "No-Git sync failed: $($_.Exception.Message)"
        }
    }
}
