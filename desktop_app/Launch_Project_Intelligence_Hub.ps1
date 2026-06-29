param(
    [string]$Url = "https://samco-mob-intelligence-dashboard.streamlit.app/"
)

$ErrorActionPreference = "Stop"

function Find-Browser {
    $candidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    $command = Get-Command msedge.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    $command = Get-Command chrome.exe -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }

    return ""
}

$browser = Find-Browser
if ($browser) {
    Start-Process -FilePath $browser -ArgumentList @("--app=$Url", "--new-window")
} else {
    Start-Process $Url
}
