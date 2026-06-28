$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$SourceRoot = Join-Path $Root "ios_app_source"
$DistRoot = Join-Path $Root "dist"
$ZipPath = Join-Path $DistRoot "Project_Intelligence_Hub_iOS_Source.zip"

if (-not (Test-Path $SourceRoot)) {
    Write-Host "iOS source package was not generated." -ForegroundColor Yellow
    Write-Host "Blocker: ios_app_source folder is missing." -ForegroundColor Yellow
    exit 1
}

New-Item -ItemType Directory -Path $DistRoot -Force | Out-Null
if (Test-Path $ZipPath) {
    Remove-Item -LiteralPath $ZipPath -Force
}

Compress-Archive -Path (Join-Path $SourceRoot "*") -DestinationPath $ZipPath -Force

if (-not (Test-Path $ZipPath)) {
    Write-Host "iOS source package was not generated." -ForegroundColor Yellow
    Write-Host "Blocker: Compress-Archive completed without creating $ZipPath" -ForegroundColor Yellow
    exit 1
}

Write-Host "iOS source package generated successfully:" -ForegroundColor Green
Write-Host $ZipPath -ForegroundColor Green
Write-Host "This is source package only, not a signed IPA." -ForegroundColor Yellow
