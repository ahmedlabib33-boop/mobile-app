param(
    [string]$Owner = "ahmedlabib33-boop",
    [string]$Repository = "mobile-app"
)

$ErrorActionPreference = "Stop"

function ConvertFrom-SecureValue([Security.SecureString]$SecureValue) {
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureValue)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer) }
}

$secureToken = Read-Host "New repository-scoped GitHub token" -AsSecureString
$token = ConvertFrom-SecureValue $secureToken
if ([string]::IsNullOrWhiteSpace($token)) { throw "Token cannot be empty." }

$headers = @{
    Authorization = "Bearer $token"
    Accept = "application/vnd.github+json"
    "X-GitHub-Api-Version" = "2022-11-28"
    "User-Agent" = "Project-Intelligence-Hub-Mobile-App-Credential-Setup"
}
$repo = Invoke-RestMethod -Method Get -Uri "https://api.github.com/repos/$Owner/$Repository" -Headers $headers
if ($null -ne $repo.permissions -and -not [bool]$repo.permissions.push) {
    throw "The token can read the repository but does not have Contents: Read and write permission."
}

$securePin = Read-Host "New private Streamlit synchronization admin PIN" -AsSecureString
$pin = ConvertFrom-SecureValue $securePin
if ([string]::IsNullOrWhiteSpace($pin)) { throw "Administrator PIN cannot be empty." }
if ($pin.Length -lt 8) { throw "Administrator PIN must contain at least 8 characters." }

[Environment]::SetEnvironmentVariable("PIH_MOBILE_APP_GITHUB_TOKEN", $token, "User")
[Environment]::SetEnvironmentVariable("PIH_MOBILE_APP_SYNC_ADMIN_PIN", $pin, "User")
$env:PIH_MOBILE_APP_GITHUB_TOKEN = $token
$env:PIH_MOBILE_APP_SYNC_ADMIN_PIN = $pin

Write-Host "Credentials validated and saved to the Windows user environment."
Write-Host "Restart Streamlit and open a new VS Code terminal before using the Output Studio controls."

$token = $null
$pin = $null
