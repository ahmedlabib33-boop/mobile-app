$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$AndroidRoot = Join-Path $Root "android_app"
$DistRoot = Join-Path $Root "dist"
$AndroidConfig = Join-Path $AndroidRoot "app\src\main\assets\mobile_config.json"
$KeyProperties = Join-Path $AndroidRoot "key.properties"
$ApkTarget = Join-Path $DistRoot "Project_Intelligence_Hub.apk"

function Stop-Blocker($Message) {
    Write-Host "APK was not generated." -ForegroundColor Yellow
    Write-Host "Blocker: $Message" -ForegroundColor Yellow
    exit 1
}

function Find-JdkBin {
    $javaCmd = Get-Command java -ErrorAction SilentlyContinue
    if ($javaCmd) { return Split-Path -Parent $javaCmd.Source }
    $candidates = @()
    foreach ($base in @("$env:ProgramFiles\Eclipse Adoptium", "$env:ProgramFiles\Java")) {
        if (Test-Path $base) {
            $candidates += Get-ChildItem -LiteralPath $base -Recurse -Filter java.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
        }
    }
    if ($candidates.Count -gt 0) { return Split-Path -Parent $candidates[0] }
    return ""
}

function Find-GradleCommand {
    $gradleCmd = Get-Command gradle -ErrorAction SilentlyContinue
    if ($gradleCmd) { return $gradleCmd.Source }
    $localGradle = Join-Path $Root "dist\android_tooling\gradle-8.10.2\bin\gradle.bat"
    if (Test-Path $localGradle) { return $localGradle }
    return ""
}

Write-Host "Building Project Intelligence Hub Android APK without Git" -ForegroundColor Cyan

if (-not (Test-Path $AndroidRoot)) {
    Stop-Blocker "android_app folder is missing."
}

$jdkBin = Find-JdkBin
if ($jdkBin) { $env:Path = "$jdkBin;$env:Path" }
if (-not (Get-Command java -ErrorAction SilentlyContinue)) {
    Stop-Blocker "Java/JDK is missing from PATH. Install JDK 17 or Android Studio and expose the bundled JDK."
}

if (-not (Test-Path $KeyProperties)) {
    Stop-Blocker "Android release signing file is missing: $KeyProperties. Run FINAL_MOBILE_BUILD_NO_GIT.ps1 -GenerateKeystore or create it from android_app\key.properties.example."
}

$defaultSdkPath = Join-Path $env:LOCALAPPDATA "Android\Sdk"
$sdkPath = @($defaultSdkPath, $env:ANDROID_HOME, $env:ANDROID_SDK_ROOT) | Where-Object { $_ -and (Test-Path $_) -and (Test-Path (Join-Path $_ "cmdline-tools")) } | Select-Object -First 1
if (-not $sdkPath) {
    Stop-Blocker "Android SDK is missing. Install Android Studio and SDK Build-Tools, or set ANDROID_HOME."
}
$env:ANDROID_HOME = $sdkPath
$env:ANDROID_SDK_ROOT = $sdkPath
$env:Path = "$sdkPath\platform-tools;$env:Path"

if (-not (Test-Path $AndroidConfig)) {
    Stop-Blocker "Android config file is missing: $AndroidConfig"
}

$config = Get-Content -LiteralPath $AndroidConfig -Raw | ConvertFrom-Json
$invalidUrlTokens = @("PUT_DEPLOYED_STREAMLIT_URL_HERE", "https://your-streamlit-url", "https://your-deployed-streamlit-app", "https://your-deployed-streamlit-url")
if (-not $config.streamlit_url -or $invalidUrlTokens -contains $config.streamlit_url -or -not $config.streamlit_url.StartsWith("https://") -or $config.streamlit_url.Contains("localhost") -or $config.streamlit_url.Contains("127.0.0.1")) {
    Stop-Blocker "Streamlit URL is missing, fake, localhost, or not HTTPS in $AndroidConfig"
}

New-Item -ItemType Directory -Path $DistRoot -Force | Out-Null

Push-Location $AndroidRoot
try {
    if (Test-Path ".\gradlew.bat") {
        .\gradlew.bat assembleRelease
    } elseif (Find-GradleCommand) {
        & (Find-GradleCommand) assembleRelease
    } else {
        Stop-Blocker "Gradle is missing and no Gradle wrapper exists. Install Gradle or build from Android Studio."
    }
} finally {
    Pop-Location
}

$apkCandidates = @(@(
    (Join-Path $AndroidRoot "app\build\outputs\apk\release\app-release.apk"),
    (Join-Path $AndroidRoot "app\build\outputs\apk\release\app-release-unsigned.apk")
) | Where-Object { Test-Path $_ })

if (-not $apkCandidates -or $apkCandidates.Count -eq 0) {
    Stop-Blocker "Gradle completed without producing a release APK in android_app\app\build\outputs\apk\release."
}

Copy-Item -LiteralPath $apkCandidates[0] -Destination $ApkTarget -Force
Write-Host "APK generated successfully:" -ForegroundColor Green
Write-Host $ApkTarget -ForegroundColor Green
