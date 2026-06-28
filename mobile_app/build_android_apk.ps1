param(
    [string]$Configuration = "release"
)

$ErrorActionPreference = "Stop"

$MobileRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $MobileRoot
$DistRoot = Join-Path $ProjectRoot "dist"
$RootConfig = Join-Path $ProjectRoot "mobile_config.json"
$AssetConfig = Join-Path $MobileRoot "assets\mobile_config.json"
$ApkSource = Join-Path $MobileRoot "build\app\outputs\flutter-apk\app-release.apk"
$ApkTarget = Join-Path $DistRoot "Project_Intelligence_Hub.apk"

Write-Host "Project Intelligence Hub Android build" -ForegroundColor Cyan

if (-not (Get-Command flutter -ErrorAction SilentlyContinue)) {
    Write-Host ""
    Write-Host "Flutter is not installed or not available in PATH." -ForegroundColor Yellow
    Write-Host "Install Flutter for Windows for free from: https://docs.flutter.dev/get-started/install/windows"
    Write-Host "After installation, open a new PowerShell window and run:"
    Write-Host "  flutter doctor"
    Write-Host "  .\mobile_app\build_android_apk.ps1"
    exit 1
}

New-Item -ItemType Directory -Path $DistRoot -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $AssetConfig) -Force | Out-Null

if (Test-Path $RootConfig) {
    Copy-Item -LiteralPath $RootConfig -Destination $AssetConfig -Force
    Write-Host "Synced mobile_config.json into Flutter assets."
}

Push-Location $MobileRoot
try {
    if (-not (Test-Path "android\gradlew.bat") -or -not (Test-Path "ios\Runner.xcodeproj")) {
        Write-Host "Generating missing Flutter Android/iOS platform folders..."
        flutter create --platforms=android,ios --project-name project_intelligence_hub --org com.samco.projectintelligence .
    }

    $AndroidManifest = Join-Path $MobileRoot "android\app\src\main\AndroidManifest.xml"
    if (Test-Path $AndroidManifest) {
        $manifestText = Get-Content -LiteralPath $AndroidManifest -Raw
        $manifestText = $manifestText -replace 'android:label="project_intelligence_hub"', 'android:label="Project Intelligence Hub"'
        if ($manifestText -notmatch 'android.permission.INTERNET') {
            $manifestText = $manifestText -replace '<manifest([^>]*)>', "<manifest`$1>`n    <uses-permission android:name=`"android.permission.INTERNET`" />"
        }
        Set-Content -LiteralPath $AndroidManifest -Value $manifestText -Encoding UTF8
    }

    $InfoPlist = Join-Path $MobileRoot "ios\Runner\Info.plist"
    if (Test-Path $InfoPlist) {
        $plistText = Get-Content -LiteralPath $InfoPlist -Raw
        if ($plistText -notmatch 'CFBundleDisplayName') {
            $plistText = $plistText -replace '<key>CFBundleName</key>', "<key>CFBundleDisplayName</key>`n`t<string>Project Intelligence Hub</string>`n`t<key>CFBundleName</key>"
        }
        Set-Content -LiteralPath $InfoPlist -Value $plistText -Encoding UTF8
    }

    Write-Host "Running flutter pub get..."
    flutter pub get

    Write-Host "Running static analysis..."
    flutter analyze

    Write-Host "Building Android APK..."
    flutter build apk --release

    if (-not (Test-Path $ApkSource)) {
        throw "Expected APK was not generated at $ApkSource"
    }

    Copy-Item -LiteralPath $ApkSource -Destination $ApkTarget -Force
    Write-Host "APK copied to: $ApkTarget" -ForegroundColor Green
}
finally {
    Pop-Location
}
