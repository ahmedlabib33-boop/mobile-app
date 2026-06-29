param(
    [ValidateSet("Watch", "Once", "DryRun")]
    [string]$Mode = "Watch",
    [int]$IntervalMinutes = 0,
    [int]$IntervalSeconds = 0,
    [string]$Message = "Synchronize Project Intelligence Hub workspace"
)

$ErrorActionPreference = "Stop"
$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$configPath = Join-Path $PSScriptRoot "github_sync_config.json"
$logPath = Join-Path $root "11-outputs\logs\pih_mobile_app_github_sync.log"
$statePath = Join-Path $root ".pih_mobile_app_sync_state\local_manifest.json"
New-Item -ItemType Directory -Force -Path (Split-Path $logPath -Parent), (Split-Path $statePath -Parent) | Out-Null

if (-not (Test-Path -LiteralPath $configPath)) { throw "Missing synchronization configuration: $configPath" }
$config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
if ($IntervalSeconds -le 0 -and $null -ne $config.interval_seconds) { $IntervalSeconds = [int]$config.interval_seconds }
if ($IntervalSeconds -le 0) {
    if ($IntervalMinutes -le 0) { $IntervalMinutes = [int]$config.interval_minutes }
    if ($IntervalMinutes -le 0) { $IntervalMinutes = 30 }
    $IntervalSeconds = $IntervalMinutes * 60
}
if ($IntervalSeconds -lt 10) { $IntervalSeconds = 10 }

$script:WatchMutex = $null
if ($Mode -eq "Watch") {
    $script:WatchMutex = New-Object System.Threading.Mutex($false, "Local\ProjectIntelligenceHubMobileAppGitHubSync")
    if (-not $script:WatchMutex.WaitOne(0, $false)) {
        Write-Host "A Project Intelligence Hub synchronization watcher is already running."
        exit 0
    }
}

function Write-SyncLog([string]$Text) {
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Text"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line
}

function Convert-ToRelativePath([string]$FullName) {
    $rootPath = $root.TrimEnd("\") + "\"
    $rootUri = New-Object System.Uri($rootPath)
    $fileUri = New-Object System.Uri([System.IO.Path]::GetFullPath($FullName))
    return ([System.Uri]::UnescapeDataString($rootUri.MakeRelativeUri($fileUri).ToString()) -replace "\\", "/")
}

function Test-Excluded([string]$FullName) {
    $relative = Convert-ToRelativePath $FullName
    $segments = $relative -split "/"
    foreach ($segment in $segments) {
        if ($segment -like "backup_*") { return $true }
        if ($segment -like "android_app_backup*") { return $true }
        if ($segment -eq "_RETURN_POINTS") { return $true }
    }
    foreach ($directory in @($config.excluded_directories)) {
        $normalizedDirectory = ([string]$directory).Trim("/")
        if ($segments -contains $normalizedDirectory -or $relative.Equals($normalizedDirectory, [System.StringComparison]::OrdinalIgnoreCase) -or $relative.StartsWith($normalizedDirectory + "/", [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    foreach ($file in @($config.excluded_files)) {
        if ($relative.Equals([string]$file, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    $name = [System.IO.Path]::GetFileName($FullName)
    foreach ($pattern in @($config.excluded_patterns)) {
        if ($name -like [string]$pattern) { return $true }
    }
    return $false
}

function Test-LegacyProjectRepositoryPath([string]$RelativePath) {
    return $RelativePath -match '^projects/[^/]+/(branding|contracts|evidence|notes)(/|$)'
}

function Ensure-EmptyDirectoryPlaceholders {
    $directories = Get-ChildItem -LiteralPath $root -Recurse -Directory -Force -ErrorAction SilentlyContinue |
        Where-Object { -not (Test-Excluded $_.FullName) } |
        Sort-Object FullName
    foreach ($directory in $directories) {
        $children = Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction SilentlyContinue |
            Where-Object { -not (Test-Excluded $_.FullName) }
        if (@($children).Count -eq 0) {
            $placeholder = Join-Path $directory.FullName ".gitkeep"
            if (-not (Test-Path -LiteralPath $placeholder)) {
                New-Item -ItemType File -Path $placeholder -Force | Out-Null
                Write-SyncLog "Created placeholder for empty folder: $(Convert-ToRelativePath $placeholder)"
            }
        }
    }
}

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-GitBlobSha([byte[]]$Bytes) {
    $prefix = [System.Text.Encoding]::ASCII.GetBytes("blob $($Bytes.Length)`0")
    $combined = New-Object byte[] ($prefix.Length + $Bytes.Length)
    [System.Buffer]::BlockCopy($prefix, 0, $combined, 0, $prefix.Length)
    [System.Buffer]::BlockCopy($Bytes, 0, $combined, $prefix.Length, $Bytes.Length)
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try { return -join ($sha1.ComputeHash($combined) | ForEach-Object { $_.ToString("x2") }) }
    finally { $sha1.Dispose() }
}

function Get-WorkspaceManifest {
    $files = Get-ChildItem -LiteralPath $root -Recurse -File -Force -ErrorAction SilentlyContinue |
        Where-Object { -not (Test-Excluded $_.FullName) }
    $entries = [ordered]@{}
    foreach ($file in $files) {
        $relative = Convert-ToRelativePath $file.FullName
        $entries[$relative] = [ordered]@{
            sha256 = Get-Sha256 $file.FullName
            size = $file.Length
            modified_utc = $file.LastWriteTimeUtc.ToString("o")
        }
    }
    return $entries
}

function Read-PreviousManifest {
    if (-not (Test-Path -LiteralPath $statePath)) { return @{} }
    try {
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        $files = @{}
        foreach ($property in $state.files.PSObject.Properties) {
            $files[$property.Name] = @{
                sha256 = [string]$property.Value.sha256
                size = [int64]$property.Value.size
                modified_utc = [string]$property.Value.modified_utc
            }
        }
        return $files
    } catch { return @{} }
}

function Write-LocalManifest([System.Collections.IDictionary]$Files, [string[]]$Changed, [string[]]$Deleted, [string]$Result) {
    $payload = [ordered]@{
        generated_at = (Get-Date).ToUniversalTime().ToString("o")
        workspace_root = $root
        mode = $Mode
        result = $Result
        changed_or_new = $Changed
        deleted_locally = $Deleted
        files = $Files
    }
    $payload | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $statePath -Encoding utf8
}

function Get-Token {
    foreach ($name in @("PIH_MOBILE_APP_GITHUB_TOKEN", "PIH_MOBILE_APP_GH_TOKEN")) {
        foreach ($scope in @("Process", "User", "Machine")) {
            $candidate = [Environment]::GetEnvironmentVariable($name, $scope)
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            $headers = @{
                Authorization = "Bearer $candidate"
                Accept = "application/vnd.github+json"
                "X-GitHub-Api-Version" = "2022-11-28"
                "User-Agent" = "Project-Intelligence-Hub-NoGit-Sync"
            }
            try {
                $repository = Invoke-RestMethod -Method Get -Uri "https://api.github.com/repos/$($config.owner)/$($config.repository)" -Headers $headers
                if ($null -ne $repository.permissions -and -not [bool]$repository.permissions.push) {
                    Write-SyncLog "Credential in $name ($scope scope) can read the repository but does not have push permission."
                    continue
                }
                Write-SyncLog "Credential accepted from $name ($scope scope) for $($config.owner)/$($config.repository)."
                return $candidate
            } catch {
                $statusCode = $null
                if ($null -ne $_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }
                if ($statusCode -eq 401) {
                    Write-SyncLog "Credential in $name ($scope scope) was rejected by GitHub (HTTP 401)."
                    continue
                }
                if ($statusCode -eq 403) {
                    Write-SyncLog "Credential in $name ($scope scope) cannot access the configured repository (HTTP 403)."
                    continue
                }
                throw
            }
        }
    }
    throw "No valid GitHub credential was found. Create a new repository-scoped token with Contents: Read and write, set it as PIH_MOBILE_APP_GITHUB_TOKEN, then open a new terminal. Codespaces user secrets are not available to this local Windows process."
}

function Invoke-GitHubApi([string]$Method, [string]$Uri, [object]$Body = $null) {
    $headers = @{
        Authorization = "Bearer $script:Token"
        Accept = "application/vnd.github+json"
        "X-GitHub-Api-Version" = "2022-11-28"
        "User-Agent" = "Project-Intelligence-Hub-NoGit-Sync"
    }
    $params = @{ Method = $Method; Uri = $Uri; Headers = $headers }
    if ($null -ne $Body) {
        $jsonBody = $Body | ConvertTo-Json -Depth 20 -Compress
        $params.Body = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
        $params.ContentType = "application/json"
    }
    try {
        return Invoke-RestMethod @params
    } catch {
        $statusCode = $null
        $responseBody = ""
        if ($null -ne $_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $responseBody = $reader.ReadToEnd()
                $reader.Dispose()
            } catch { $responseBody = "" }
        }
        $detail = if ($responseBody) { " Response: $responseBody" } else { "" }
        throw "GitHub API $Method $Uri failed$(if ($statusCode) { " (HTTP $statusCode)" }): $($_.Exception.Message)$detail"
    }
}

function Invoke-SyncCycle {
    Ensure-EmptyDirectoryPlaceholders
    $current = Get-WorkspaceManifest
    $previous = Read-PreviousManifest
    $changed = @($current.Keys | Where-Object { -not $previous.ContainsKey($_) -or $previous[$_].sha256 -ne $current[$_].sha256 } | Sort-Object)
    $deleted = @($previous.Keys | Where-Object { -not $current.Contains($_) } | Sort-Object)
    Write-SyncLog "Workspace scan: $($current.Count) files; changed/new: $($changed.Count); deleted locally: $($deleted.Count)."

    if ($Mode -eq "DryRun") {
        Write-LocalManifest $current $changed $deleted "dry-run"
        foreach ($path in ($changed | Select-Object -First 100)) { Write-SyncLog "DRY RUN changed/new: $path" }
        foreach ($path in ($deleted | Select-Object -First 100)) { Write-SyncLog "DRY RUN deleted locally: $path" }
        if ($changed.Count -gt 100) { Write-SyncLog "DRY RUN additional changed/new paths omitted from log: $($changed.Count - 100)" }
        if ($deleted.Count -gt 100) { Write-SyncLog "DRY RUN additional deleted paths omitted from log: $($deleted.Count - 100)" }
        return
    }

    $script:Token = Get-Token
    $apiBase = "https://api.github.com/repos/$($config.owner)/$($config.repository)"
    $emptyRepository = $false
    $baseCommitSha = ""
    $baseTreeSha = ""
    try {
        $ref = Invoke-GitHubApi "Get" "$apiBase/git/ref/heads/$($config.branch)"
        $baseCommitSha = $ref.object.sha
        $baseCommit = Invoke-GitHubApi "Get" "$apiBase/git/commits/$baseCommitSha"
        $baseTreeSha = $baseCommit.tree.sha
        $remoteTree = Invoke-GitHubApi "Get" "$apiBase/git/trees/$baseTreeSha`?recursive=1"
    } catch {
        if ($_.Exception.Message -match "\(HTTP 404\)" -or $_.Exception.Message -match "\(HTTP 409\)") {
            $emptyRepository = $true
            $remoteTree = @{ tree = @() }
            Write-SyncLog "Remote branch $($config.branch) does not exist. The repository appears empty; an initial commit will be created."
        } else {
            throw
        }
    }
    $remote = @{}
    foreach ($item in $remoteTree.tree) { if ($item.type -eq "blob") { $remote[$item.path] = $item.sha } }

    if ($emptyRepository) {
        $bootstrapText = "# Project Intelligence Hub`n`nRepository initialized for the mobile Streamlit deployment package.`n"
        $bootstrapBytes = [System.Text.Encoding]::UTF8.GetBytes($bootstrapText)
        Invoke-GitHubApi "Put" "$apiBase/contents/README.md" @{
            message = "Initialize mobile-app repository"
            content = [Convert]::ToBase64String($bootstrapBytes)
            branch = $config.branch
        } | Out-Null
        Write-SyncLog "Created initial repository commit on branch $($config.branch); continuing with full synchronization."

        $ref = Invoke-GitHubApi "Get" "$apiBase/git/ref/heads/$($config.branch)"
        $baseCommitSha = $ref.object.sha
        $baseCommit = Invoke-GitHubApi "Get" "$apiBase/git/commits/$baseCommitSha"
        $baseTreeSha = $baseCommit.tree.sha
        $remoteTree = Invoke-GitHubApi "Get" "$apiBase/git/trees/$baseTreeSha`?recursive=1"
        $remote = @{}
        foreach ($item in $remoteTree.tree) { if ($item.type -eq "blob") { $remote[$item.path] = $item.sha } }
        $emptyRepository = $false
    }

    $uploadEntries = @()
    $deleteEntries = @()
    $maxBytes = [int64]$config.max_file_size_mb * 1MB
    foreach ($path in $current.Keys) {
        $fullName = Join-Path $root ($path -replace "/", "\")
        $info = Get-Item -LiteralPath $fullName
        if ($info.Length -gt $maxBytes) { Write-SyncLog "Skipped over-size file: $path"; continue }
        $bytes = [System.IO.File]::ReadAllBytes($fullName)
        $localBlobSha = Get-GitBlobSha $bytes
        if ($remote.ContainsKey($path) -and $remote[$path] -eq $localBlobSha) { continue }
        $blob = Invoke-GitHubApi "Post" "$apiBase/git/blobs" @{ content = [Convert]::ToBase64String($bytes); encoding = "base64" }
        $uploadEntries += @{ path = $path; mode = "100644"; type = "blob"; sha = $blob.sha }
    }
    if ([bool]$config.sync_deletions -or [bool]$config.prune_legacy_project_folders) {
        foreach ($path in $remote.Keys) {
            $globalDeletionAllowed = [bool]$config.sync_deletions
            $legacyDeletionAllowed = [bool]$config.prune_legacy_project_folders -and (Test-LegacyProjectRepositoryPath $path)
            if (($globalDeletionAllowed -or $legacyDeletionAllowed) -and -not $current.Contains($path) -and -not (Test-Excluded (Join-Path $root ($path -replace "/", "\")))) {
                $deleteEntries += @{ path = $path; mode = "100644"; type = "blob"; sha = $null }
                if ($legacyDeletionAllowed -and -not $globalDeletionAllowed) { Write-SyncLog "Pruning legacy project path: $path" }
            }
        }
    } elseif ($deleted.Count -gt 0) {
        Write-SyncLog "General deletion synchronization is disabled; remote files were not deleted."
    }

    if (($uploadEntries.Count + $deleteEntries.Count) -eq 0) {
        Write-LocalManifest $current $changed $deleted "no-remote-change"
        Write-SyncLog "No remote changes required."
        return
    }

    function Invoke-TreeEntryBatches([array]$Entries, [string]$PhaseName, [string]$StartingCommitSha) {
        if ($Entries.Count -eq 0) { return $StartingCommitSha }
        $batchSize = 200
        $latestCommitSha = $StartingCommitSha
        for ($offset = 0; $offset -lt $Entries.Count; $offset += $batchSize) {
            $end = [math]::Min($offset + $batchSize - 1, $Entries.Count - 1)
            $batch = @($Entries[$offset..$end])
            $batchNumber = [int]([math]::Floor($offset / $batchSize) + 1)
            $batchTotal = [int]([math]::Ceiling($Entries.Count / $batchSize))

            $updated = $false
            for ($attempt = 1; $attempt -le 3 -and -not $updated; $attempt++) {
                $latestRef = Invoke-GitHubApi "Get" "$apiBase/git/ref/heads/$($config.branch)"
                $latestCommitSha = $latestRef.object.sha
                $latestCommit = Invoke-GitHubApi "Get" "$apiBase/git/commits/$latestCommitSha"
                $latestTreeSha = $latestCommit.tree.sha
                $newTree = Invoke-GitHubApi "Post" "$apiBase/git/trees" @{ base_tree = $latestTreeSha; tree = $batch }
                $commit = Invoke-GitHubApi "Post" "$apiBase/git/commits" @{
                    message = "$Message $PhaseName batch $batchNumber/$batchTotal - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
                    tree = $newTree.sha
                    parents = @($latestCommitSha)
                }
                try {
                    Invoke-GitHubApi "Patch" "$apiBase/git/refs/heads/$($config.branch)" @{ sha = $commit.sha; force = $true } | Out-Null
                    $latestCommitSha = $commit.sha
                    $updated = $true
                    Write-SyncLog "Synchronization $PhaseName batch $batchNumber/$batchTotal complete: $($commit.sha)"
                } catch {
                    if ($_.Exception.Message -match "\(HTTP 409\)" -and $attempt -lt 3) {
                        Write-SyncLog "GitHub branch update conflict on $PhaseName batch $batchNumber/$batchTotal; retrying with latest remote head (attempt $($attempt + 1)/3)."
                        Start-Sleep -Seconds 2
                    } else {
                        throw
                    }
                }
            }
        }
        return $latestCommitSha
    }

    $latestCommitSha = $baseCommitSha
    $latestCommitSha = Invoke-TreeEntryBatches -Entries ([array]$uploadEntries) -PhaseName "upload" -StartingCommitSha $latestCommitSha
    $latestCommitSha = Invoke-TreeEntryBatches -Entries ([array]$deleteEntries) -PhaseName "delete" -StartingCommitSha $latestCommitSha
    Write-LocalManifest $current $changed $deleted "synced:$latestCommitSha"
    Write-SyncLog "Synchronization complete: $latestCommitSha"
}

Write-SyncLog "Started mode=$Mode intervalSeconds=$IntervalSeconds root=$root target=$($config.owner)/$($config.repository):$($config.branch) deletionSync=$($config.sync_deletions)"
if ($Mode -in @("Once", "DryRun")) {
    try {
        Invoke-SyncCycle
        exit 0
    } catch {
        Write-SyncLog "Synchronization failed at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
        exit 1
    }
}
while ($true) {
    try { Invoke-SyncCycle }
    catch { Write-SyncLog "Synchronization failed at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)" }
    Start-Sleep -Seconds $IntervalSeconds
}
