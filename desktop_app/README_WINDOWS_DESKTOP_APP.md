# Project Intelligence Hub - Windows Desktop App

This folder provides a free Windows desktop launcher for:

```text
https://samco-mob-intelligence-dashboard.streamlit.app/
```

It opens the deployed Streamlit app in a clean desktop app-style browser window using Microsoft Edge or Google Chrome if available.

## Create Desktop Shortcut

Run:

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\desktop_app"
.\Create_Windows_Desktop_Shortcut.ps1
```

Then open **Project Intelligence Hub** from the Windows desktop.

## Open Directly

Run:

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\desktop_app"
.\Launch_Project_Intelligence_Hub.bat
```

## Android APK On Windows

Windows does not run APK files natively. To open the Android APK from a desktop PC, use one of these routes:

- Android Studio Emulator with an Android Virtual Device.
- A physical Android phone connected by USB with APK install enabled.
- A third-party Android emulator.

The Windows desktop launcher in this folder is the cleaner desktop route because it opens the same stable Streamlit app directly.
