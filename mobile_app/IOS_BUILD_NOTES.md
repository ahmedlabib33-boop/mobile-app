# iOS Build Notes

The Flutter iOS source is prepared in this `mobile_app` project. The app opens the deployed Streamlit URL in a native Flutter WebView shell.

Important limitations:

- Building iOS requires macOS with Xcode installed.
- Windows cannot produce a final signed iOS IPA.
- A free personal Apple ID may allow limited local device testing through Xcode, subject to Apple's current restrictions.
- Professional distribution, App Store publishing, TestFlight, and permanent signed IPA distribution require the paid Apple Developer Program.

Typical iOS preparation on macOS:

```powershell
cd mobile_app
flutter pub get
flutter create --platforms=ios --project-name project_intelligence_hub --org com.samco.projectintelligence .
flutter build ios --release
```

Then open `ios/Runner.xcworkspace` in Xcode, configure signing, select the target device, and build/archive.
