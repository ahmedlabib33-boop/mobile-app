$ErrorActionPreference = "Continue"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$AndroidRoot = Join-Path $Root "android_app"
$DistRoot = Join-Path $Root "dist"
$AndroidConfig = Join-Path $AndroidRoot "app\src\main\assets\mobile_config.json"
$KeyProperties = Join-Path $AndroidRoot "key.properties"

function Write-Check($Name, $Ok, $Detail) {
    $status = if ($Ok) { "OK" } else { "MISSING" }
    $color = if ($Ok) { "Green" } else { "Yellow" }
    Write-Host ("[{0}] {1} - {2}" -f $status, $Name, $Detail) -ForegroundColor $color
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

Write-Host "Project Intelligence Hub Android no-Git toolchain check" -ForegroundColor Cyan
Write-Host "Root: $Root"
Write-Host "Git is not required and this script does not call Git." -ForegroundColor Green

$jdkBin = Find-JdkBin
if ($jdkBin) { $env:Path = "$jdkBin;$env:Path" }
$java = Get-Command java -ErrorAction SilentlyContinue
Write-Check "Java/JDK" ([bool]$java) ($(if ($java) { $java.Source } else { "Install JDK 17 or use Android Studio bundled JDK." }))

$studioCandidates = @(
    "$env:ProgramFiles\Android\Android Studio\bin\studio64.exe",
    "${env:ProgramFiles(x86)}\Android\Android Studio\bin\studio64.exe",
    "$env:LOCALAPPDATA\Programs\Android Studio\bin\studio64.exe"
) | Where-Object { $_ -and (Test-Path $_) }
Write-Check "Android Studio" ($studioCandidates.Count -gt 0) ($(if ($studioCandidates.Count) { $studioCandidates[0] } else { "Install Android Studio for free." }))

$defaultSdkPath = Join-Path $env:LOCALAPPDATA "Android\Sdk"
$sdkCandidates = @(@($defaultSdkPath, $env:ANDROID_HOME, $env:ANDROID_SDK_ROOT) | Where-Object { $_ -and (Test-Path $_) -and (Test-Path (Join-Path $_ "cmdline-tools")) })
$sdkPath = if ($sdkCandidates.Count -gt 0) { [string]$sdkCandidates[0] } else { "" }
Write-Check "Android SDK" ($sdkPath -ne "") ($(if ($sdkPath) { $sdkPath } else { "Set ANDROID_HOME or install SDK through Android Studio." }))

$gradle = Find-GradleCommand
$wrapper = Test-Path (Join-Path $AndroidRoot "gradlew.bat")
Write-Check "Gradle or wrapper" ([bool]$gradle -or $wrapper) ($(if ($gradle) { $gradle } elseif ($wrapper) { "android_app\gradlew.bat" } else { "Install Gradle or build from Android Studio." }))

$buildTools = $false
if ($sdkPath) {
    $buildToolsPath = Join-Path $sdkPath "build-tools"
    $buildTools = (Test-Path $buildToolsPath) -and ((Get-ChildItem -LiteralPath $buildToolsPath -Directory -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0)
}
Write-Check "Android build tools" $buildTools ($(if ($buildTools) { Join-Path $sdkPath "build-tools" } else { "Install Android SDK Build-Tools from Android Studio SDK Manager." }))

$configExists = Test-Path $AndroidConfig
Write-Check "Android mobile_config.json" $configExists $AndroidConfig
Write-Check "Android release signing key.properties" (Test-Path $KeyProperties) ($(if (Test-Path $KeyProperties) { $KeyProperties } else { "Missing. Run FINAL_MOBILE_BUILD_NO_GIT.ps1 -GenerateKeystore." }))

$urlConfigured = $false
if ($configExists) {
    try {
        $config = Get-Content -LiteralPath $AndroidConfig -Raw | ConvertFrom-Json
        $invalidUrlTokens = @("PUT_DEPLOYED_STREAMLIT_URL_HERE", "https://your-streamlit-url", "https://your-deployed-streamlit-app", "https://your-deployed-streamlit-url")
        $urlConfigured = [string]$config.streamlit_url -and ($invalidUrlTokens -notcontains $config.streamlit_url) -and $config.streamlit_url.StartsWith("https://") -and -not $config.streamlit_url.Contains("localhost") -and -not $config.streamlit_url.Contains("127.0.0.1")
    } catch {
        $urlConfigured = $false
    }
}
Write-Check "Streamlit HTTPS URL configured" $urlConfigured ($(if ($urlConfigured) { "Configured" } else { "Still placeholder or not HTTPS. Edit $AndroidConfig" }))

if (-not (Test-Path $DistRoot)) {
    New-Item -ItemType Directory -Path $DistRoot -Force | Out-Null
}
Write-Check "dist folder" (Test-Path $DistRoot) $DistRoot

if ($java -and $sdkPath -and ($gradle -or $wrapper) -and $buildTools -and $urlConfigured) {
    Write-Host "Toolchain status: ready to build APK without Git." -ForegroundColor Green
    exit 0
}

Write-Host "Toolchain status: not ready. See MISSING items above." -ForegroundColor Yellow
exit 1
