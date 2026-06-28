# Project Intelligence Hub iOS Source

This folder contains native Swift WKWebView source for the iOS version of Project Intelligence Hub.

Status:

- iOS source is prepared.
- Windows cannot build a signed iOS IPA.
- macOS + Xcode are required for final build.
- Apple Developer Program is required for App Store, TestFlight, and professional signed distribution.
- A free Apple ID may allow limited local device testing only, subject to Apple's current restrictions.
- This is an Apple ecosystem limitation, not a code failure.

Suggested bundle identifier:

```text
com.samco.projectintelligencehub
```

Set the deployed Streamlit URL in:

```text
Config/mobile_config.json
```

or update the fallback constant in:

```text
Sources/AppConfig.swift
```

Typical macOS/Xcode steps:

1. Create a new iOS App project in Xcode named `Project Intelligence Hub`.
2. Use Swift and SwiftUI.
3. Set bundle identifier to `com.samco.projectintelligencehub`.
4. Copy the files from `Sources/` into the Xcode project.
5. Add `Config/mobile_config.json` to the app target bundle.
6. Add the `Info.plist` values from `Config/Info.plist`.
7. Configure signing in Xcode.
8. Build and run on Simulator or a connected device.
9. Archive only after Apple signing is configured.

## TestFlight / App Store

TestFlight and App Store distribution require:

- macOS.
- Xcode.
- Apple Developer Program membership.
- A configured signing team and bundle identifier.
- App archive/upload through Xcode Organizer or Apple Transporter.

## Free iPhone Alternative

The immediate free iPhone route is the PWA/home-screen workflow:

1. Open the deployed Streamlit URL in Safari.
2. Tap Share.
3. Tap Add to Home Screen.
4. Open the Project Intelligence Hub icon.
5. Sign in with the director username and password.

No Git is required.
