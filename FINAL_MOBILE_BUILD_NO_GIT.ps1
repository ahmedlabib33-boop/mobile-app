param(
    [string]$DeployedStreamlitUrl = "",
    [switch]$GenerateKeystore,
    [string]$KeyAlias = "project_intelligence_hub",
    [string]$StorePassword = "",
    [string]$KeyPassword = "",
    [switch]$InstallTools,
    [switch]$SkipAndroidBuild
)

$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$DistRoot = Join-Path $Root "dist"
$AndroidRoot = Join-Path $Root "android_app"
$IosRoot = Join-Path $Root "ios_app_source"
$RootConfig = Join-Path $Root "mobile_config.json"
$AndroidConfig = Join-Path $AndroidRoot "app\src\main\assets\mobile_config.json"
$IosConfig = Join-Path $IosRoot "Config\mobile_config.json"
$KeyProperties = Join-Path $AndroidRoot "key.properties"
$KeystorePath = Join-Path $AndroidRoot "release-key.jks"
$DirectorZip = Join-Path $DistRoot "Project_Intelligence_Hub_Director_Handover.zip"
$Status = [ordered]@{}

function Write-Step($Message) {
    Write-Host "`n== $Message ==" -ForegroundColor Cyan
}

function Update-JsonUrl($Path, $Url) {
    if (-not (Test-Path $Path)) { return }
    $json = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    $json.streamlit_url = $Url
    $json | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Update-PwaManifestUrl($Path, $Url) {
    if (-not (Test-Path $Path)) { return }
    $json = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
    $json.start_url = $Url
    $json.scope = $Url
    $json | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function New-RandomPassword {
    -join ((48..57 + 65..90 + 97..122) | Get-Random -Count 24 | ForEach-Object {[char]$_})
}

function Test-RealHttpsUrl($Url) {
    $invalidUrlTokens = @("PUT_DEPLOYED_STREAMLIT_URL_HERE", "https://your-streamlit-url", "https://your-deployed-streamlit-app", "https://your-deployed-streamlit-url")
    return [string]$Url -and ($invalidUrlTokens -notcontains $Url) -and $Url.StartsWith("https://") -and -not $Url.Contains("localhost") -and -not $Url.Contains("127.0.0.1")
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

New-Item -ItemType Directory -Path $DistRoot -Force | Out-Null

Write-Step "No-Git mobile build orchestrator"
Write-Host "Root: $Root"
Write-Host "Git is not required and this script does not call Git." -ForegroundColor Green

if ($DeployedStreamlitUrl) {
    if (-not (Test-RealHttpsUrl $DeployedStreamlitUrl)) {
        throw "DeployedStreamlitUrl must be a real HTTPS URL. Localhost, 127.0.0.1, and placeholder URLs are rejected."
    }
    Write-Step "Updating Streamlit URL in config files"
    Update-JsonUrl $RootConfig $DeployedStreamlitUrl
    Update-JsonUrl $AndroidConfig $DeployedStreamlitUrl
    Update-JsonUrl $IosConfig $DeployedStreamlitUrl
    Update-PwaManifestUrl (Join-Path $Root "pwa\manifest.json") $DeployedStreamlitUrl
    foreach ($contentPath in @((Join-Path $Root "DIRECTOR_HANDOVER_MESSAGE.md"), (Join-Path $DistRoot "DIRECTOR_INSTALL_GUIDE.html"), (Join-Path $Root "pwa\README_PWA.md"), (Join-Path $Root "pwa\index.html"))) {
        if (Test-Path $contentPath) {
            (Get-Content -LiteralPath $contentPath -Raw).Replace("PUT_DEPLOYED_STREAMLIT_URL_HERE", $DeployedStreamlitUrl).Replace("https://your-streamlit-url", $DeployedStreamlitUrl) | Set-Content -LiteralPath $contentPath -Encoding UTF8
        }
    }
}

Write-Step "Checking free Android tooling"
$jdkBin = Find-JdkBin
if ($jdkBin) { $env:Path = "$jdkBin;$env:Path" }
$java = Get-Command java -ErrorAction SilentlyContinue
$gradle = Find-GradleCommand
$defaultSdkPath = Join-Path $env:LOCALAPPDATA "Android\Sdk"
$sdkPath = @($defaultSdkPath, $env:ANDROID_HOME, $env:ANDROID_SDK_ROOT) | Where-Object { $_ -and (Test-Path $_) -and (Test-Path (Join-Path $_ "cmdline-tools")) } | Select-Object -First 1
$studio = @("$env:ProgramFiles\Android\Android Studio\bin\studio64.exe", "${env:ProgramFiles(x86)}\Android\Android Studio\bin\studio64.exe", "$env:LOCALAPPDATA\Programs\Android Studio\bin\studio64.exe") | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

if ($InstallTools) {
    Write-Step "Installing free Android tooling with winget"
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "winget is not available. Install Android Studio and JDK 17 manually. Git is not required."
    }
    if (-not $java) {
        winget install --id EclipseAdoptium.Temurin.17.JDK --source winget --accept-package-agreements --accept-source-agreements
    }
    if (-not $studio) {
        winget install --id Google.AndroidStudio --source winget --accept-package-agreements --accept-source-agreements
    }
    if (-not $gradle) {
        winget install --id Gradle.Gradle --source winget --accept-package-agreements --accept-source-agreements
    }
    Write-Host "Tool installation requested. Open a new PowerShell window and rerun this script after Android Studio SDK setup completes." -ForegroundColor Yellow
}

$Status.Java = if ($java) { $java.Source } else { "Missing JDK. Install JDK 17 or Android Studio bundled JDK." }
$Status.AndroidStudio = if ($studio) { $studio } else { "Missing Android Studio." }
$Status.AndroidSdk = if ($sdkPath) { $sdkPath } else { "Missing Android SDK." }
$Status.Gradle = if ($gradle) { $gradle } else { "Missing Gradle or Gradle wrapper." }

$urlConfigured = $false
if (Test-Path $AndroidConfig) {
    $androidJson = Get-Content -LiteralPath $AndroidConfig -Raw | ConvertFrom-Json
    $urlConfigured = Test-RealHttpsUrl $androidJson.streamlit_url
}
$Status.StreamlitUrl = if ($urlConfigured) { "Configured" } else { "Missing deployed HTTPS Streamlit URL." }

if ($GenerateKeystore) {
    Write-Step "Generating Android release keystore"
    if (-not $java) { throw "keytool requires Java/JDK. Install JDK 17 first." }
    $keytool = Join-Path (Split-Path -Parent $java.Source) "keytool.exe"
    if (-not (Test-Path $keytool)) { $keytool = "keytool" }
    if (-not $StorePassword) { $StorePassword = New-RandomPassword }
    if (-not $KeyPassword) { $KeyPassword = $StorePassword }
    if (-not (Test-Path $KeystorePath)) {
        & $keytool -genkeypair -v -keystore $KeystorePath -storepass $StorePassword -keypass $KeyPassword -alias $KeyAlias -keyalg RSA -keysize 2048 -validity 10000 -dname "CN=Project Intelligence Hub, OU=Planning Department, O=SAMCO, L=Cairo, S=Cairo, C=EG"
    }
    @"
storeFile=release-key.jks
storePassword=$StorePassword
keyAlias=$KeyAlias
keyPassword=$KeyPassword
"@ | Set-Content -LiteralPath $KeyProperties -Encoding ASCII
}

$Status.Signing = if (Test-Path $KeyProperties) { "key.properties exists" } else { "Missing key.properties. Run with -GenerateKeystore after JDK is installed." }

Write-Step "Packaging iOS source"
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "package_ios_source_no_git.ps1")
$IosZip = Join-Path $DistRoot "Project_Intelligence_Hub_iOS_Source.zip"
$Status.IosZip = if (Test-Path $IosZip) { $IosZip } else { "Not generated" }

if (-not $SkipAndroidBuild) {
    Write-Step "Attempting Android APK build"
    try {
        & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $Root "build_android_no_git.ps1")
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Yellow
    }
} else {
    Write-Host "Android build skipped by parameter." -ForegroundColor Yellow
}

$ApkPath = Join-Path $DistRoot "Project_Intelligence_Hub.apk"
$Status.AndroidApk = if (Test-Path $ApkPath) { $ApkPath } else { "Not generated" }

Write-Step "Creating director handover ZIP"
if (Test-Path $DirectorZip) { Remove-Item -LiteralPath $DirectorZip -Force }
$distHandoverMessage = Join-Path $DistRoot "DIRECTOR_HANDOVER_MESSAGE.md"
Copy-Item -LiteralPath (Join-Path $Root "DIRECTOR_HANDOVER_MESSAGE.md") -Destination $distHandoverMessage -Force
$handoverItems = @(
    (Join-Path $DistRoot "DIRECTOR_INSTALL_GUIDE.html"),
    $distHandoverMessage,
    (Join-Path $Root "MOBILE_BUILD_STATUS.md")
)
if (Test-Path $ApkPath) { $handoverItems += $ApkPath }
if (Test-Path $IosZip) { $handoverItems += $IosZip }
foreach ($pwaItem in @("pwa\README_PWA.md", "pwa\manifest.json", "pwa\service-worker.js", "pwa\index.html", "pwa\ICON_PLACEHOLDER_NOTES.md")) {
    $candidate = Join-Path $Root $pwaItem
    if (Test-Path $candidate) { $handoverItems += $candidate }
}
Compress-Archive -LiteralPath $handoverItems -DestinationPath $DirectorZip -Force
$Status.DirectorZip = $DirectorZip

Write-Step "Final status"
$Status.GetEnumerator() | ForEach-Object { Write-Host ("{0}: {1}" -f $_.Key, $_.Value) }
Write-Host "No Git used. Flutter not required. No fake APK or IPA claimed." -ForegroundColor Green
