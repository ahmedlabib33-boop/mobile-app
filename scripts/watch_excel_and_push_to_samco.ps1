param(
    [int]$DebounceSeconds = 20,
    [string]$MessagePrefix = "Sync mobile app Excel and template data updates"
)

$ErrorActionPreference = "Stop"

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$syncScript = Join-Path $PSScriptRoot "sync_current_to_samco_no_git.ps1"
$logPath = Join-Path $root "11-outputs\logs\pih_mobile_app_excel_sync_to_github.log"
New-Item -ItemType Directory -Force -Path (Split-Path $logPath -Parent) | Out-Null

$watchRoots = @(
    "projects",
    "templates"
) | ForEach-Object { Join-Path $root $_ } | Where-Object { Test-Path -LiteralPath $_ }

$extensions = @(".xlsx", ".xls", ".xlsm", ".csv")
$pending = $false
$lastChange = Get-Date

function Write-SyncLog([string]$Message) {
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -LiteralPath $logPath -Value "[$stamp] $Message"
}

function Test-SyncCandidate([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }
    $name = [System.IO.Path]::GetFileName($Path)
    if ($name.StartsWith("~$")) {
        return $false
    }
    return $extensions -contains ([System.IO.Path]::GetExtension($Path).ToLowerInvariant())
}

$watcher = foreach ($watchRoot in $watchRoots) {
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $watchRoot
    $watcher.IncludeSubdirectories = $true
    $watcher.EnableRaisingEvents = $true
    $watcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite, Size'

    foreach ($eventName in @("Changed", "Created", "Deleted", "Renamed")) {
        Register-ObjectEvent -InputObject $watcher -EventName $eventName -Action {
            $eventPath = $Event.SourceEventArgs.FullPath
            if (Test-SyncCandidate $eventPath) {
                $script:pending = $true
                $script:lastChange = Get-Date
                Write-SyncLog "Detected $($Event.SourceEventArgs.ChangeType): $eventPath"
            }
        } | Out-Null
    }
    $watcher
}

Write-SyncLog "Started Excel/CSV watcher for: $($watchRoots -join '; ')"

while ($true) {
    Start-Sleep -Seconds 2
    if ($pending -and ((Get-Date) - $lastChange).TotalSeconds -ge $DebounceSeconds) {
        $pending = $false
        try {
            $message = "$MessagePrefix - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            Write-SyncLog "Running sync: $message"
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $syncScript -Owner "ahmedlabib33-boop" -Repo "mobile-app" -Branch "main" -Message $message *>> $logPath
            Write-SyncLog "Sync completed."
        } catch {
            Write-SyncLog "Sync failed: $($_.Exception.Message)"
        }
    }
}
