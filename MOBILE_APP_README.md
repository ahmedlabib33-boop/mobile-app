# Project Intelligence Hub Mobile App

## What Was Created

This project now has a free-first mobile architecture:

```text
Existing Streamlit app
-> hosted HTTPS Streamlit URL
-> Flutter WebView mobile shell
-> Android APK and iOS-ready Flutter source
```

Created folders and files:

- `mobile_app/` Flutter source for the Android/iOS WebView shell.
- `mobile_app/lib/main.dart` production mobile shell with WebView, loading state, error state, back handling, pull-to-refresh, external link/export handling, and no debug banner.
- `mobile_app/build_android_apk.ps1` Android APK build script.
- `mobile_app/IOS_BUILD_NOTES.md` iOS build and signing notes.
- `mobile_config.json` root configuration for app name, Streamlit URL, and output path.
- `pwa/` optional PWA manifest/service worker package.
- `dist/` output folder for the release APK.

## Architecture

The Python/Streamlit application remains the source of truth and should be deployed online. The Flutter app is a native mobile shell that opens the deployed Streamlit URL using `webview_flutter`.

This keeps the backend and dashboards in Python while giving users an installable Android app and an iOS-ready codebase from one free Flutter codebase.

## Deploy Streamlit Online

Use any free or existing hosting path that can run Python and Streamlit, for example:

- Streamlit Community Cloud if the repo can be published there.
- A free VM/container host.
- A local company server exposed through HTTPS.
- Streamlit Community Cloud for the stable mobile/PWA URL.

Run locally for testing:

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub"
python -m streamlit run app.py
```

For production mobile use, configure a public HTTPS URL. Do not use a local-only URL for the final APK.

## Change The Streamlit URL

Edit:

```powershell
D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\mobile_config.json
```

Set:

```json
"streamlit_url": "https://your-deployed-streamlit-url"
```

The Android build script automatically copies this file into:

```powershell
mobile_app\assets\mobile_config.json
```

## Build Android APK

Install Flutter for Windows first:

```powershell
flutter doctor
```

Then build:

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub"
.\mobile_app\build_android_apk.ps1
```

APK output:

```powershell
D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\dist\Project_Intelligence_Hub.apk
```

The script runs:

- Flutter availability check.
- Flutter platform bootstrap if `android/` or `ios/` are missing.
- `flutter pub get`
- `flutter analyze`
- `flutter build apk --release`
- APK copy into `dist/`.

## iOS Status

The Flutter source is iOS-ready, but a signed iOS IPA cannot be produced on Windows only.

Requirements:

- macOS
- Xcode
- Flutter installed on macOS
- Apple signing configured in Xcode

Free personal Apple ID may allow limited device testing. App Store publishing, TestFlight, and permanent signed IPA distribution require Apple Developer Program membership.

See:

```powershell
mobile_app\IOS_BUILD_NOTES.md
```

## Streamlit Mobile Improvements

The Streamlit dashboard now includes a global mobile-first CSS layer:

- Collapsed sidebar default preserved.
- Touch-friendly buttons and controls.
- Better metric/KPI card spacing.
- Mobile wrapping for columns.
- Readable tables with horizontal scrolling instead of page overflow.
- Responsive Plotly containers.
- Mobile tab overflow handling.
- Dark/light system compatibility.
- More polished alert/card styling.

Charts now share cleaner executive defaults through the central `style_plotly()` helper:

- Responsive autosize.
- Cleaner title alignment.
- Better hover labels.
- Mobile-safe margins.
- Horizontal legends.
- Axis automargins.

## PWA Option

The `pwa/` folder contains:

- `manifest.json`
- `service-worker.js`
- `README_PWA.md`

This can be used if the hosted Streamlit deployment is served with a small proxy/static layer that exposes the manifest, service worker, and icons over HTTPS.

## Free vs Paid Limitations

Free:

- Flutter framework.
- Android APK build.
- Local Android install.
- iOS Flutter source preparation.
- Limited iOS testing with a free personal Apple ID, subject to Apple restrictions.
- PWA home-screen option.

Not fully free:

- App Store distribution.
- TestFlight.
- Permanent signed iOS IPA distribution.
- Apple Developer Program signing for professional iOS release.

## Troubleshooting

If Flutter is missing:

```powershell
flutter doctor
```

If Android build fails, install Android Studio and Android SDK, then rerun:

```powershell
flutter doctor --android-licenses
flutter doctor
```

If the app shows the offline/config screen, confirm `mobile_config.json` has a public HTTPS Streamlit URL and rebuild the APK.

If downloads do not open inside the WebView, the shell routes export-like links to the external browser/app where Android can handle the file.

## Recommended Next Improvements

- Replace icon and splash placeholders with final branded artwork.
- Deploy Streamlit to a stable HTTPS host.
- Add Firebase-free crash logging through local server logs or Streamlit-side telemetry if needed.
- Add CI build instructions once Flutter is installed on a build machine.
- Test on at least one Android phone and one iPhone through Xcode before distribution.
