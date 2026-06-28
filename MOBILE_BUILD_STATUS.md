# Mobile Build Status

Generated: 2026-06-29

## Live Mobile URL

```text
https://mobile-app.streamlit.app
```

This is the intended stable Streamlit Community Cloud URL for the mobile/PWA release. After deploying the GitHub repo on Streamlit Community Cloud, confirm Streamlit assigns this exact URL; if Streamlit assigns a different stable URL, rerun the final mobile build script with that URL and rebuild the APK.

## Android Native APK

Android APK generated: **Yes**

APK path:

```text
D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\dist\Project_Intelligence_Hub.apk
```

Current Android readiness:

- Native Android WebView source: **Ready**
- Package name: `com.samco.projectintelligencehub`
- App name: `Project Intelligence Hub`
- Streamlit URL configured: **Yes**
- JDK: **Ready**
- Android SDK / platform / build-tools: **Ready**
- Local Gradle: **Ready**
- Release keystore: **Ready**
- Signed release APK: **Generated**

## iPhone PWA App

iPhone native IPA generated: **No**

Reason: native iOS IPA requires macOS, Xcode, and Apple signing. This free package does not claim or include a native IPA.

iPhone free app ready: **Yes, via PWA/Add to Home Screen**

iPhone PWA URL:

```text
https://mobile-app.streamlit.app
```

Director iPhone steps:

1. Open the Streamlit URL in Safari.
2. Tap Share.
3. Tap Add to Home Screen.
4. Name it Project Intelligence Hub.
5. Open it like an app.
6. Login with director username/password.

## Authentication

Login/sign-up added: **Yes**

Roles added: **Yes**

Implemented:

- SQLite `runtime/project_intelligence_hub_mobile_auth.sqlite`
- username/email + password login
- sign-up page
- PBKDF2-HMAC salted password hashing
- first-run admin setup
- roles: `admin`, `director`, `viewer`
- admin user management
- director executive navigation
- viewer read-only restrictions
- logout
- session protection
- professional login screen and footer

## Director Handover

Director handover ZIP generated: **Yes**

Path:

```text
D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\dist\Project_Intelligence_Hub_Director_Handover.zip
```

The ZIP includes:

- `Project_Intelligence_Hub.apk`
- `DIRECTOR_INSTALL_GUIDE.html`
- `DIRECTOR_HANDOVER_MESSAGE.md`
- `MOBILE_BUILD_STATUS.md`
- PWA instructions/files
- `Project_Intelligence_Hub_iOS_Source.zip` as source only

## Confirmations

- GitHub repo update requested for deployment handoff; no Git is required to run the app.
- NO Flutter required.
- Android native APK exists physically.
- No fake APK claimed.
- iPhone PWA app is the free iOS route.
- No fake IPA claimed.
- Future native iOS route: macOS/Xcode/TestFlight/App Store with Apple signing.

