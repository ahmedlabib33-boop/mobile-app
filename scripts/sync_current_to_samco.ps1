param(
    [string]$Message = "Sync current Project Intelligence Hub updates",
    [switch]$NoPush
)

$ErrorActionPreference = "Stop"

function Get-GitExecutable {
    $git = Get-Command git -ErrorAction SilentlyContinue
    if ($git) {
        return $git.Source
    }

    $githubDesktopGit = Get-ChildItem -Path "$env:LOCALAPPDATA\GitHubDesktop" -Recurse -Filter git.exe -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -like "*\resources\app\git\cmd\git.exe" } |
        Sort-Object FullName -Descending |
        Select-Object -First 1

    if ($githubDesktopGit) {
        return $githubDesktopGit.FullName
    }

    throw "Git was not found. Install Git or GitHub Desktop before syncing."
}

function Invoke-RoboCopyMirror {
    param(
        [string]$Source,
        [string]$Target,
        [string[]]$ExtraExcludedDirs = @()
    )

    if (-not (Test-Path $Source)) {
        if (Test-Path $Target) {
            Remove-Item -LiteralPath $Target -Recurse -Force
        }
        return
    }

    New-Item -ItemType Directory -Path $Target -Force | Out-Null
    $excludedDirs = @("__pycache__", ".pytest_cache", ".mypy_cache", ".ruff_cache", ".vite", ".streamlit\delay_tia_upload_cache") + $ExtraExcludedDirs
    $excludedFiles = @("*.pyc", "*.pyo", "*.log", "*.err", "*.out", "*.tmp", "~$*")

    & robocopy $Source $Target /MIR /XD $excludedDirs /XF $excludedFiles /R:2 /W:1 /NFL /NDL /NJH /NJS /NP | Out-Null
    if ($LASTEXITCODE -gt 7) {
        throw "Robocopy failed for $Source -> $Target with exit code $LASTEXITCODE"
    }
}

function Copy-IfPresent {
    param(
        [string]$Source,
        [string]$Target
    )

    if (Test-Path $Source) {
        $parent = Split-Path -Parent $Target
        if ($parent) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        Copy-Item -LiteralPath $Source -Destination $Target -Force
    } elseif (Test-Path $Target) {
        Remove-Item -LiteralPath $Target -Force
    }
}

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$git = Get-GitExecutable
$repoUrl = "https://github.com/ahmedlabib33-boop/mobile-app.git"
$syncRoot = Join-Path $env:TEMP ("project-intelligence-hub-mobile-app-sync-worktree-" + [guid]::NewGuid().ToString("N"))

Push-Location $root
try {
    $remoteNames = & $git remote
    if ($remoteNames -notcontains "mobileapp") {
        & $git remote add mobileapp $repoUrl
    }
    & $git fetch mobileapp main

    if (Test-Path $syncRoot) {
        $resolvedSyncRoot = (Resolve-Path $syncRoot).Path
        $resolvedTemp = (Resolve-Path $env:TEMP).Path
        if (-not $resolvedSyncRoot.StartsWith($resolvedTemp, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to remove sync worktree outside TEMP: $resolvedSyncRoot"
        }
        & $git worktree remove --force $syncRoot 2>$null
        if (Test-Path $syncRoot) {
            Remove-Item -LiteralPath $syncRoot -Recurse -Force
        }
    }

    & $git worktree prune
    & $git worktree add --detach $syncRoot mobileapp/main
} finally {
    Pop-Location
}

$dirsToMirror = @(
    ".devcontainer",
    ".vscode",
    "assets",
    "data",
    "docs",
    "exports",
    "reports",
    "projects",
    "scripts",
    "src",
    "templates",
    "tests",
    "ui"
)

foreach ($dir in $dirsToMirror) {
    Invoke-RoboCopyMirror -Source (Join-Path $root $dir) -Target (Join-Path $syncRoot $dir)
}

Invoke-RoboCopyMirror -Source (Join-Path $root "analytics") -Target (Join-Path $syncRoot "analytics") -ExtraExcludedDirs @("node_modules")

$topLevelFiles = @(
    ".gitattributes",
    ".gitignore",
    "app.py",
    "beba.md",
    "contract_claims_center.py",
    "dashboard.py",
    "Delay_TIA_Conclusion_Methodology_Summary.md",
    "README.md",
    "requirements.txt",
    "runtime.txt",
    "RUN_APP.bat",
    "RUN_FULL_PROJECT_NO_GIT_SYNC.bat",
    "RUN_TUNNEL.bat",
    "RUN_LIVE_EXCEL_SYNC.bat"
)

foreach ($file in $topLevelFiles) {
    Copy-IfPresent -Source (Join-Path $root $file) -Target (Join-Path $syncRoot $file)
}

$legacyToolName = ("Cod" + "ex")
$legacyReportTier = ("Super" + "_" + "Prem" + "ium")
$obsoleteTopLevelFiles = @(
    ($legacyToolName + "_" + ("Pro" + "mpt") + "_" + $legacyReportTier + "_Interactive_Progress_Report.md"),
    ("README_TIA_" + $legacyToolName + ".md")
)

foreach ($file in $obsoleteTopLevelFiles) {
    $target = Join-Path $syncRoot $file
    if (Test-Path $target) {
        Remove-Item -LiteralPath $target -Force
    }
}

New-Item -ItemType Directory -Path (Join-Path $syncRoot ".streamlit") -Force | Out-Null
Copy-IfPresent -Source (Join-Path $root ".streamlit\config.toml") -Target (Join-Path $syncRoot ".streamlit\config.toml")
$cachePath = Join-Path $syncRoot ".streamlit\delay_tia_upload_cache"
if (Test-Path $cachePath) {
    Remove-Item -LiteralPath $cachePath -Recurse -Force
}

Push-Location $syncRoot
try {
    & $git add -A
    $status = & $git status --porcelain
    if (-not $status) {
        Write-Host "No source/template/data changes to sync."
        return
    }

    & $git commit -m $Message
    if (-not $NoPush) {
        & $git push $repoUrl HEAD:main
    }
} finally {
    Pop-Location
}
