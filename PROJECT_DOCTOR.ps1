$ErrorActionPreference = "Continue"

$ProjectRoot = "C:\Users\pc\OneDrive\Documents\Project Intelligence Hub"
$Stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Out = "$ProjectRoot\PROJECT_DOCTOR_REPORT_$Stamp.txt"

Set-Location $ProjectRoot

function Add-Title {
    param([string]$Title)
    "`n===== $Title =====" | Out-File $Out -Append -Encoding utf8
}

function Add-CommandOutput {
    param(
        [string]$Title,
        [scriptblock]$Command
    )
    Add-Title $Title
    try {
        & $Command 2>&1 | Out-File $Out -Append -Encoding utf8
    } catch {
        "ERROR: $($_.Exception.Message)" | Out-File $Out -Append -Encoding utf8
    }
}

"===== PROJECT DOCTOR REPORT =====" | Out-File $Out -Encoding utf8
"Date: $(Get-Date)" | Out-File $Out -Append -Encoding utf8
"ProjectRoot: $ProjectRoot" | Out-File $Out -Append -Encoding utf8

Add-CommandOutput "PYTHON" { python --version }
Add-CommandOutput "PIP" { python -m pip --version }
Add-CommandOutput "STREAMLIT" { python -m streamlit --version }

Add-Title "IMPORTANT FILES"
$Files = @(
    "dashboard.py",
    "app.py",
    "requirements.txt",
    "PROJECT_CONTEXT.md",
    ".streamlit\config.toml",
    "README.md",
    "contract_claims_center.py",
    "reports\tia_director_pack_generator.py",
    "exports\word_template_exporter.py",
    "src\construction_system\database.py",
    "src\construction_system\importers.py",
    "src\construction_system\steel_delay_tia.py"
)

foreach ($File in $Files) {
    if (Test-Path $File) {
        "FOUND: $File" | Out-File $Out -Append -Encoding utf8
    } else {
        "MISSING: $File" | Out-File $Out -Append -Encoding utf8
    }
}

Add-Title "REQUIREMENTS"
if (Test-Path "requirements.txt") {
    Get-Content "requirements.txt" | Out-File $Out -Append -Encoding utf8
} else {
    "requirements.txt missing" | Out-File $Out -Append -Encoding utf8
}

Add-Title "STREAMLIT CONFIG"
if (Test-Path ".streamlit\config.toml") {
    Get-Content ".streamlit\config.toml" | Out-File $Out -Append -Encoding utf8
} else {
    ".streamlit\config.toml missing" | Out-File $Out -Append -Encoding utf8
}

Add-Title "PYTHON SYNTAX CHECK"
$CompileTargets = @("dashboard.py", "app.py", "src", "reports", "exports", "scripts") | Where-Object { Test-Path $_ }

if ($CompileTargets.Count -gt 0) {
    python -m compileall $CompileTargets 2>&1 | Out-File $Out -Append -Encoding utf8
} else {
    "No compile targets found." | Out-File $Out -Append -Encoding utf8
}

Add-Title "IMPORT CHECK"
$ImportScript = Join-Path $env:TEMP "pih_import_check_$Stamp.py"

$ImportLines = @(
    "import importlib",
    "modules = ['streamlit','pandas','plotly','docx','openpyxl','pptx','reportlab','openai','numpy','pypdf']",
    "for m in modules:",
    "    try:",
    "        importlib.import_module(m)",
    "        print(f'OK: {m}')",
    "    except Exception as e:",
    "        print(f'MISSING/ERROR: {m} -> {e}')"
)

$ImportLines | Set-Content $ImportScript -Encoding UTF8
python $ImportScript 2>&1 | Out-File $Out -Append -Encoding utf8
Remove-Item $ImportScript -Force -ErrorAction SilentlyContinue

Add-CommandOutput "GIT STATUS - READ ONLY" { git status --short --branch }
Add-CommandOutput "REMOTES - READ ONLY" { git remote -v }

Add-Title "ROOT FILES"
Get-ChildItem -File |
ForEach-Object {
    "$($_.Name)`t$($_.Length)`t$($_.LastWriteTime)"
} | Out-File $Out -Append -Encoding utf8

Add-Title "LIGHT PROJECT TREE"
Get-ChildItem -Recurse -File |
Where-Object {
    $_.FullName -notmatch "\\\.git\\|\\\.venv\\|\\venv\\|\\__pycache__\\|\\node_modules\\|\\outputs\\|\\generated_outputs\\|\\logs\\|\\\.streamlit\\delay_tia_upload_cache\\"
} |
ForEach-Object {
    "$($_.FullName)`t$($_.Length)`t$($_.LastWriteTime)"
} | Out-File $Out -Append -Encoding utf8

Write-Host "Project Doctor Report created:" $Out -ForegroundColor Green
notepad $Out
