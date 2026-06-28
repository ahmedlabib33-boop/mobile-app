# Android Studio Build Guide - No Git

## Open The Project

1. Install Android Studio from the official Android website.
2. Open Android Studio.
3. Choose **Open**.
4. Select:

```text
D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\android_app
```

Git is not required.

## Sync Gradle

Android Studio should detect the Gradle project and sync automatically. If it does not:

1. Open **File > Sync Project with Gradle Files**.
2. Install any missing Android SDK or Build-Tools versions requested by Android Studio.

## Set The Streamlit URL

Edit:

```text
android_app\app\src\main\assets\mobile_config.json
```

Replace:

```json
"streamlit_url": "PUT_DEPLOYED_STREAMLIT_URL_HERE"
```

with the deployed HTTPS URL:

```json
"streamlit_url": "https://your-deployed-streamlit-app"
```

Do not use localhost for the production APK.

## Build APK From Android Studio

1. Confirm the Streamlit URL is configured.
2. Create release signing files if they do not exist:

```powershell
.\FINAL_MOBILE_BUILD_NO_GIT.ps1 -GenerateKeystore -SkipAndroidBuild
```

This creates `android_app\release-key.jks` and `android_app\key.properties`. Keep both private.

3. Select **Build > Build Bundle(s) / APK(s) > Build APK(s)**.
4. Android Studio will create the APK under:

```text
android_app\app\build\outputs\apk\debug
```

For release APK, use **Build > Generate Signed Bundle / APK** or the PowerShell script below.

## Build APK From PowerShell

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub"
.\check_android_toolchain_no_git.ps1
.\build_android_no_git.ps1
```

Final copied APK path:

```text
D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\dist\Project_Intelligence_Hub.apk
```

## Common Errors

- **Java/JDK missing**: Install JDK 17 or use Android Studio bundled JDK and expose it to PATH.
- **Android SDK missing**: Install SDK through Android Studio SDK Manager and set `ANDROID_HOME`.
- **Build Tools missing**: Install Android SDK Build-Tools from SDK Manager.
- **Gradle missing**: Build inside Android Studio or install Gradle. No Git is required.
- **Signing missing**: Create `android_app\key.properties` from `key.properties.example` or run `FINAL_MOBILE_BUILD_NO_GIT.ps1 -GenerateKeystore`.
- **URL placeholder**: Edit `android_app\app\src\main\assets\mobile_config.json`.
- **HTTP URL blocked**: Use HTTPS for production.

## Confirmation

This Android app is native Java + WebView. It does not require Git, Flutter, React Native, Node, or paid services.

Optional tool installer:

```powershell
.\FINAL_MOBILE_BUILD_NO_GIT.ps1 -InstallTools -SkipAndroidBuild
```

This uses `winget` for free JDK, Android Studio, and Gradle installs. It does not install or require Git.
