# Project Intelligence Hub Mobile Package - No Git

## Final Architecture

```text
Existing Streamlit / Python App
-> Hosted HTTPS Streamlit URL
-> Native Android WebView App
-> Native iOS WKWebView Source
-> Android APK + iOS source package
```

The Streamlit app remains the dashboard and backend. The mobile apps are professional native wrappers around the deployed Streamlit URL.

## Why No Flutter

The requested no-Git strategy avoids Flutter as the primary path because Flutter installation commonly depends on Git. The primary deliverable is now native Android Java WebView plus native iOS Swift WKWebView source.

## Why No Git

The Android and iOS sources are local files only. The build scripts do not call Git, clone repositories, or use GitHub workflows.

## Android App

Folder:

```text
android_app
```

Package:

```text
com.samco.projectintelligencehub
```

Features:

- Native Android WebView shell.
- JavaScript enabled.
- DOM storage enabled.
- Safe browsing enabled where available.
- HTTPS-only production network policy.
- Internet and network state permissions.
- Loading progress bar.
- Offline/configuration error screen.
- Back button handling.
- Lightweight pull-to-refresh gesture.
- File download/export handling through Android DownloadManager or external app fallback.
- External links opened outside the app.
- Portrait orientation.
- Corporate status/navigation bar colors.

Streamlit URL config:

```text
android_app\app\src\main\assets\mobile_config.json
```

## Required Free Android Tools

- Android Studio.
- Android SDK and Build-Tools.
- JDK 17.
- Gradle or Android Studio build runner.

Git is not required.

## Android APK Build

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub"
.\check_android_toolchain_no_git.ps1
.\build_android_no_git.ps1
```

For the complete release flow:

```powershell
.\FINAL_MOBILE_BUILD_NO_GIT.ps1 -DeployedStreamlitUrl "https://mobile-app.streamlit.app" -GenerateKeystore
```

To request free tool installation through Windows Package Manager without Git:

```powershell
.\FINAL_MOBILE_BUILD_NO_GIT.ps1 -InstallTools -SkipAndroidBuild
```

After Android Studio installs, open it once and install Android SDK Build-Tools from SDK Manager.

APK output:

```text
D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\dist\Project_Intelligence_Hub.apk
```

The script refuses to build if Java, Android SDK, Gradle, Build-Tools, or a real HTTPS Streamlit URL are missing.

## Authentication

The Streamlit app now includes local SQLite authentication in `auth.py`.

Features:

- First-run admin setup.
- Username/email and password login.
- Sign-up flow for viewer accounts.
- PBKDF2-HMAC password hashing with salt.
- No plaintext passwords.
- Roles: `admin`, `director`, `viewer`.
- Admin user management.
- Director-focused dashboard navigation.
- Viewer read-only mode.
- Logout button and session protection.

## iOS App

Folder:

```text
ios_app_source
```

Source:

- SwiftUI app entry.
- WKWebView wrapper.
- Config JSON.
- Info.plist template.
- App icon placeholder notes.

Suggested bundle identifier:

```text
com.samco.projectintelligencehub
```

## iOS Build On Mac

Windows cannot build a signed iOS IPA.

On macOS:

1. Install Xcode.
2. Create a SwiftUI iOS app project.
3. Copy `ios_app_source\Sources` into the Xcode project.
4. Add `ios_app_source\Config\mobile_config.json` to the app target.
5. Apply the values from `ios_app_source\Config\Info.plist`.
6. Configure signing.
7. Build/run on Simulator or device.

App Store, TestFlight, and professional IPA distribution require Apple Developer Program signing.

## iOS Source Package

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub"
.\package_ios_source_no_git.ps1
```

Output:

```text
dist\Project_Intelligence_Hub_iOS_Source.zip
```

This ZIP is source only, not a signed IPA.

## Streamlit Mobile Improvements

The existing Streamlit dashboard was improved with:

- Responsive CSS.
- Collapsed sidebar default.
- Touch-friendly controls.
- Better KPI/metric cards.
- Mobile-safe table overflow.
- Responsive Plotly containers.
- Cleaner Plotly titles, legends, hover labels, and margins.
- Mobile wrapping for columns and tabs.
- Dark/light system compatibility.
- Safer deployed URL handling for shared dashboard links.

## PWA Backup

Folder:

```text
pwa
```

Includes:

- `manifest.json`
- `service-worker.js`
- icon placeholders
- `README_PWA.md`

The PWA option allows an HTTPS-hosted Streamlit app to be added to Android/iOS home screens when served through a static/proxy layer.

## Troubleshooting

- **APK not generated**: Run `check_android_toolchain_no_git.ps1` and fix missing items.
- **URL placeholder blocker**: Edit `android_app\app\src\main\assets\mobile_config.json`.
- **HTTP blocked**: Use HTTPS.
- **iOS IPA requested on Windows**: Not possible. Use macOS + Xcode.
- **Git missing**: Expected and acceptable. Git is not required.
- **No users exist**: The app opens the first-run admin setup screen.
- **Director password handover**: Create director accounts in the sidebar admin panel and send credentials through a secure channel.

## Files Modified

- `dashboard.py`
- `mobile_config.json`

## Files Created

- `android_app/`
- `ios_app_source/`
- `check_android_toolchain_no_git.ps1`
- `build_android_no_git.ps1`
- `package_ios_source_no_git.ps1`
- `ANDROID_STUDIO_BUILD_GUIDE.md`
- `MOBILE_COMPLETE_NO_GIT_README.md`
- `MOBILE_BUILD_STATUS.md`
- `DIRECTOR_HANDOVER_MESSAGE.md`
- `FINAL_MOBILE_BUILD_NO_GIT.ps1`
- `pwa/`

## Current Build Status

Java/JDK, Android SDK, and Gradle are not available in this shell, so the APK was not generated here. The iOS source ZIP can be generated locally because it only packages source files.

