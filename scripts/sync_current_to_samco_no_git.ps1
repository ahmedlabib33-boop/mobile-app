param(
    [string]$Owner = "ahmedlabib33-boop",
    [string]$Repo = "mobile-app",
    [string]$Branch = "main",
    [string]$Message = "No-Git sync mobile Streamlit app",
    [switch]$DryRun,
    [int]$ChunkSize = 500
)

$ErrorActionPreference = "Stop"

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$apiBase = "https://api.github.com/repos/$Owner/$Repo"
$logPath = Join-Path $root "11-outputs\logs\pih_mobile_app_no_git_sync.log"
New-Item -ItemType Directory -Force -Path (Split-Path $logPath -Parent) | Out-Null

$excludedDirs = @(
    ".git",
    ".venv",
    "venv",
    ".pih_mobile_app_sync_state",
    "logs",
    ".pytest_cache",
    ".mypy_cache",
    ".ruff_cache",
    ".vite",
    ".streamlit\delay_tia_upload_cache",
    "android_app\app\build",
    "android_app\.gradle",
    "dist\android_tooling",
    "__pycache__",
    "node_modules",
    "project_data"
)

$excludedFilePatterns = @(
    "*.pyc",
    "*.pyo",
    "*.log",
    "*.err",
    "*.out",
    "*.tmp",
    "~$*",
    ".env",
    ".env.*",
    "*.token",
    "*.secret",
    "NO_GIT_PROJECT_DIAGNOSTIC.txt",
    "PROJECT_DIAGNOSTIC_OUTPUT.txt"
)

function Write-SyncLog([string]$MessageText) {
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$stamp] $MessageText"
    Write-Host $line
    Add-Content -LiteralPath $logPath -Value $line
}

function Get-GitHubToken {
    if (-not [string]::IsNullOrWhiteSpace($env:PIH_MOBILE_APP_GITHUB_TOKEN)) {
        return $env:PIH_MOBILE_APP_GITHUB_TOKEN
    }
    if (-not [string]::IsNullOrWhiteSpace($env:PIH_MOBILE_APP_GH_TOKEN)) {
        return $env:PIH_MOBILE_APP_GH_TOKEN
    }
    $gh = Get-Command gh -ErrorAction SilentlyContinue
    if ($gh) {
        $token = & $gh.Source auth token 2>$null
        if (-not [string]::IsNullOrWhiteSpace($token)) {
            return $token.Trim()
        }
    }
    throw "No GitHub token found. Set PIH_MOBILE_APP_GITHUB_TOKEN with repo write access before using no-Git sync."
}

function Invoke-GitHubApi {
    param(
        [string]$Method,
        [string]$Uri,
        [object]$Body = $null
    )

    $headers = @{
        Authorization          = "Bearer $script:GitHubToken"
        Accept                 = "application/vnd.github+json"
        "X-GitHub-Api-Version" = "2022-11-28"
        "User-Agent"           = "Project-Intelligence-Hub-NoGit-Sync"
    }
    $params = @{
        Method  = $Method
        Uri     = $Uri
        Headers = $headers
    }
    if ($null -ne $Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 20 -Compress)
        $params.ContentType = "application/json"
    }
    return Invoke-RestMethod @params
}

function Convert-ToRepoPath([string]$FullName) {
    $relative = [System.IO.Path]::GetRelativePath($root, $FullName)
    foreach ($part in ($relative -split '[\\/]')) {
        if ($part -like "backup_*") {
            return $true
        }
    }
    return ($relative -replace "\\", "/")
}

function Test-ExcludedPath([string]$FullName) {
    $relative = [System.IO.Path]::GetRelativePath($root, $FullName)
    foreach ($dir in $excludedDirs) {
        if ($relative.Equals($dir, [System.StringComparison]::OrdinalIgnoreCase) -or
            $relative.StartsWith($dir + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase) -or
            $relative.StartsWith($dir + [System.IO.Path]::AltDirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    $name = [System.IO.Path]::GetFileName($FullName)
    foreach ($pattern in $excludedFilePatterns) {
        if ($name -like $pattern) {
            return $true
        }
    }
    return $false
}

function Get-GitBlobSha([byte[]]$Bytes) {
    $prefix = [System.Text.Encoding]::ASCII.GetBytes("blob $($Bytes.Length)`0")
    $combined = New-Object byte[] ($prefix.Length + $Bytes.Length)
    [System.Buffer]::BlockCopy($prefix, 0, $combined, 0, $prefix.Length)
    [System.Buffer]::BlockCopy($Bytes, 0, $combined, $prefix.Length, $Bytes.Length)
    $sha1 = [System.Security.Cryptography.SHA1]::Create()
    try {
        $hash = $sha1.ComputeHash($combined)
        return -join ($hash | ForEach-Object { $_.ToString("x2") })
    } finally {
        $sha1.Dispose()
    }
}

function New-GitHubBlob([byte[]]$Bytes) {
    $payload = @{
        content  = [System.Convert]::ToBase64String($Bytes)
        encoding = "base64"
    }
    $blob = Invoke-GitHubApi -Method "Post" -Uri "$apiBase/git/blobs" -Body $payload
    return $blob.sha
}

function Get-LocalFileMap {
    $files = Get-ChildItem -LiteralPath $root -Recurse -File -Force |
        Where-Object { -not (Test-ExcludedPath $_.FullName) }
    $map = @{}
    foreach ($file in $files) {
        $repoPath = Convert-ToRepoPath $file.FullName
        $map[$repoPath] = $file.FullName
    }
    return $map
}

$script:GitHubToken = Get-GitHubToken
Write-SyncLog "Starting no-Git sync from $root to ${Owner}/${Repo}:${Branch}"

$ref = Invoke-GitHubApi -Method "Get" -Uri "$apiBase/git/ref/heads/$Branch"
$baseCommitSha = $ref.object.sha
$baseCommit = Invoke-GitHubApi -Method "Get" -Uri "$apiBase/git/commits/$baseCommitSha"
$baseTreeSha = $baseCommit.tree.sha
$remoteTree = Invoke-GitHubApi -Method "Get" -Uri "$apiBase/git/trees/$baseTreeSha`?recursive=1"

$remoteBlobMap = @{}
foreach ($item in $remoteTree.tree) {
    if ($item.type -eq "blob") {
        $remoteBlobMap[$item.path] = $item.sha
    }
}

$localFileMap = Get-LocalFileMap
$localPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($path in $localFileMap.Keys) {
    [void]$localPaths.Add($path)
}

$treeEntries = New-Object System.Collections.Generic.List[object]
$changedCount = 0
$deletedCount = 0
$skippedLargeCount = 0
$maxBytes = 95MB

foreach ($path in ($localFileMap.Keys | Sort-Object)) {
    $fullName = $localFileMap[$path]
    $bytes = [System.IO.File]::ReadAllBytes($fullName)
    if ($bytes.Length -gt $maxBytes) {
        $skippedLargeCount++
        Write-SyncLog "Skipped over-size file: $path ($($bytes.Length) bytes)"
        continue
    }
    $localSha = Get-GitBlobSha $bytes
    if ($remoteBlobMap.ContainsKey($path) -and $remoteBlobMap[$path] -eq $localSha) {
        continue
    }
    $changedCount++
    if (-not $DryRun) {
        $blobSha = New-GitHubBlob $bytes
    } else {
        $blobSha = $localSha
    }
    $treeEntries.Add(@{
            path = $path
            mode = "100644"
            type = "blob"
            sha  = $blobSha
        })
}

foreach ($remotePath in ($remoteBlobMap.Keys | Sort-Object)) {
    if (-not $localPaths.Contains($remotePath)) {
        $deletedCount++
        $treeEntries.Add(@{
                path = $remotePath
                mode = "100644"
                type = "blob"
                sha  = $null
            })
    }
}

Write-SyncLog "Detected changed/new files: $changedCount; deleted files: $deletedCount; skipped large files: $skippedLargeCount."

if ($DryRun) {
    Write-SyncLog "Dry run complete. No GitHub changes were written."
    return
}

if ($treeEntries.Count -eq 0) {
    Write-SyncLog "No changes to sync."
    return
}

$currentTreeSha = $baseTreeSha
for ($i = 0; $i -lt $treeEntries.Count; $i += $ChunkSize) {
    $count = [Math]::Min($ChunkSize, $treeEntries.Count - $i)
    $chunk = $treeEntries.GetRange($i, $count)
    $treePayload = @{
        base_tree = $currentTreeSha
        tree      = $chunk
    }
    $newTree = Invoke-GitHubApi -Method "Post" -Uri "$apiBase/git/trees" -Body $treePayload
    $currentTreeSha = $newTree.sha
    Write-SyncLog "Created tree chunk $([Math]::Floor($i / $ChunkSize) + 1) with $count entries."
}

$commitPayload = @{
    message = $Message
    tree    = $currentTreeSha
    parents = @($baseCommitSha)
}
$newCommit = Invoke-GitHubApi -Method "Post" -Uri "$apiBase/git/commits" -Body $commitPayload
$updatePayload = @{
    sha   = $newCommit.sha
    force = $false
}
Invoke-GitHubApi -Method "Patch" -Uri "$apiBase/git/refs/heads/$Branch" -Body $updatePayload | Out-Null
Write-SyncLog "No-Git sync complete. Commit: $($newCommit.sha)"
