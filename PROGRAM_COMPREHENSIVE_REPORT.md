# Projects Intelligence Hub - Comprehensive Program Report

Generated: 2026-06-30  
Workspace: `D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub`  
Repository: `https://github.com/ahmedlabib33-boop/mobile-app`  
Live Streamlit URL: `https://samco-mob-intelligence-dashboard.streamlit.app/`

## 1. Executive Summary

Projects Intelligence Hub is a Streamlit-based integrated project controls system with:

- Main Streamlit dashboard: `dashboard.py`
- Separate owner admin console: `admin_app.py`
- Authentication and role control: `auth.py`
- Project-scoped data folders under `projects/`
- Contract and claims intelligence: `contract_claims_center.py`
- Data quality and export center: `src/construction_system/premium_platform.py`
- Android native WebView app: `android_app/`
- iPhone free app route: PWA/Add to Home Screen under `pwa/`
- No-Git GitHub sync system: `tools/github_no_git_sync.ps1`
- Direct and 30-second watcher BAT launchers

The system is designed to keep the Python/Streamlit web application as the source of truth, while Android and iPhone users access it through a mobile shell or PWA.

## 2. High-Level Architecture

```text
Local Project Folder
  |
  |-- Streamlit Dashboard: dashboard.py
  |-- Admin Console: admin_app.py
  |-- Auth DB: runtime/project_intelligence_hub_mobile_auth.sqlite
  |-- Project Data: projects/<Sector>/<Project>/
  |-- Source Modules: src/construction_system/
  |-- Reports/Exports: reports/, 11-outputs/, dist/
  |
  +--> GitHub No-Git Sync
        |
        v
      GitHub Repo: ahmedlabib33-boop/mobile-app
        |
        v
      Streamlit Cloud
        |
        v
      Web URL / Android WebView / iPhone PWA
```

## 3. Main Application Entry Points

| Purpose | File | Notes |
|---|---|---|
| Main client dashboard | `dashboard.py` | Main Streamlit app deployed to Streamlit Cloud. |
| Streamlit compatibility entry | `app.py` | Imports `dashboard.py`. |
| Admin console | `admin_app.py` | Separate owner-only Streamlit app/host. |
| Authentication | `auth.py` | Login, signup, roles, remember-me, admin approval. |
| Contract claims module | `contract_claims_center.py` | Contract/claims intelligence center. |
| Data quality/export module | `src/construction_system/premium_platform.py` | Editable data, validation, chart pack, reports, ZIP exports. |

## 4. Branding Rules

Current visible branding:

- Login title: `Projects Intelligence Hub`
- Login subtitle: `Integrated Project Controls System`
- All-projects dashboard title: `Projects Intelligence Hub`
- All-projects subtitle: `Decision Making Dashboard`
- Project page title: `Projects - Projects Intelligence Hub`
- Project page subtitle: `Integrated Project Controls System`
- Contractor display name hard-coded as: `SAMCO - NATIONAL`

The contractor hard-code is applied in:

- `dashboard.py`
- `src/construction_system/project_catalog.py`
- `reports/tia_director_pack_generator.py`

Generic technical/legal labels such as `Contractor Delay`, `Contractor Rights`, or `Contractor Mitigation` remain as role/category terms.

## 5. Authentication and Access Control

Authentication file:

`auth.py`

Runtime database:

`runtime/project_intelligence_hub_mobile_auth.sqlite`

Roles:

- `admin`: owner-only.
- `director`: executive dashboard access, based on allowed sections.
- `viewer`: read-only access, based on allowed sections.

Admin owner identity:

- Username: `Ahmed_Labib`
- Email: `ahmedlabib33@gmail.com`

Admin rule:

Only the owner identity can hold admin rights. Any non-owner admin attempt is blocked or downgraded.

Signup workflow:

1. User opens the main dashboard.
2. User selects `Sign Up`.
3. User enters full name, username, email, and password.
4. Account is created inactive/pending.
5. User cannot access the dashboard yet.
6. Owner opens Admin Console.
7. Owner approves the account.
8. Owner chooses role and allowed dashboard sections.
9. User logs in normally.

Remember-me:

- Implemented with hashed remember token stored in the auth SQLite DB.
- Token is stored in query parameters for the current device/session flow.
- Logout revokes the remember token.

## 6. Admin Console

Admin app file:

`admin_app.py`

Local launcher:

```bat
RUN_ADMIN_CONSOLE.bat
```

Local URL:

```text
http://localhost:18756
```

Desktop launcher:

```text
C:\Users\pc\OneDrive\Desktop\Project Intelligence Hub Admin Console.bat
```

Admin console functions:

- View all users.
- Create approved users.
- Approve pending signups.
- Assign role: `director` or `viewer`.
- Assign allowed dashboard sections.
- Reset passwords.
- Remove users except owner.
- Download access matrix CSV.
- View user status chart.

Streamlit Cloud admin deployment:

Create a second Streamlit Cloud app from the same GitHub repo:

- Repo: `ahmedlabib33-boop/mobile-app`
- Branch: `main`
- Main file path: `admin_app.py`
- Suggested app name: `samco-project-intelligence-admin`

Important limitation:

Two separate Streamlit Cloud apps do not automatically share the same local SQLite DB. Local main/admin hosts share the same DB because they run from the same folder. Cloud-shared user persistence requires a shared database backend if both apps must manage the exact same users in production.

## 7. Project Pipeline Folder Structure

Projects are discovered under:

```text
projects/<Sector>/<Project>/
```

Current top-level sectors:

- `projects/Bridges`
- `projects/Buildings`
- `projects/Tunnels`
- `projects/_PROJECT_TEMPLATE`

Detected project manifests:

- `projects/Bridges/LMD-Bridgs/project_manifest.json`
- `projects/Buildings/Sophia-Mall/project_manifest.json`
- `projects/Buildings/The BIG - PH01/project_manifest.json`
- `projects/Tunnels/Suez-Tunnel/project_manifest.json`

Standard project folder structure:

```text
01-data/import_templates/
02-delay_analysis/steel_delay_tia_templates/
02-delay_analysis/methodology/
03-schedule/
04-source_excel/
05-contracts/source/
05-contracts/clauses/
06-evidence/
07-letters_intelligence/inbox/From Contractor/
07-letters_intelligence/inbox/From Consultant/
08-branding/
09-notes/
10-deliverables/
11-outputs/
12-logs/
```

Project discovery logic:

- Implemented in `src/construction_system/project_catalog.py`.
- Reads `project_manifest.json`, `project.json`, and `01-data/import_templates/projects.csv`.
- Builds project catalog using `discover_projects()` and `projects_frame()`.
- Supports both direct project folders and sector/project nesting.

## 8. Core Data Wiring

Main dashboard data paths are defined in `dashboard.py`:

```text
PROJECTS_CSV_PATH
ACTIVITIES_CSV_PATH
EVM_CSV_PATH
CONTRACTS_CSV_PATH
PAYMENTS_CSV_PATH
DELAYS_CSV_PATH
RISKS_CSV_PATH
MILESTONES_CSV_PATH
CHANGE_ORDERS_CSV_PATH
S_CURVE_CSV_PATH
WBS_CSV_PATH
```

Data loading function:

`load_core_csv(path, project_id=None)`

Behavior:

- If a project is selected, data is loaded from that project folder.
- If no project is selected, data is aggregated across discovered projects.
- If the loaded CSV does not include `project_id`, the loader injects the discovered project id.
- If a CSV has mismatched `project_id`, the original is preserved as `source_project_id` and corrected to the discovered project id.
- Contractor/company display fields are forced to `SAMCO - NATIONAL`.

Data normalization:

- CSV frames are normalized through `normalize_import_template_frame()`.
- Missing/blank data is filled safely.
- Project-scoped data avoids global/shared pollution.

## 9. Data Linkage and Relationships

Primary data linkage keys:

- `project_id`
- `activity_id`
- `wbs_id`
- `contract_id`
- delay/event identifiers
- letter/reference identifiers

Important data relationships:

- Projects link to activities, WBS, EVM, contracts, payments, risks, milestones, delays, and change orders through `project_id`.
- Activities and WBS drive progress, cost, schedule, and EVM views.
- Delay/TIA workflows connect delay events to activities, schedule data, evidence, clauses, and report outputs.
- Letters Intelligence connects inbound/outbound correspondence and threads using SAMCO/ACE reference patterns.
- Contract Claims Center links contract clauses, evidence, rebuttals, and claim draft outputs inside each project folder.

Decision Making Dashboard:

- Builds portfolio registry from discovered projects.
- Aggregates contract value, paid amount, remaining value, progress, SPI, CPI, risk score, delay days, milestone count, and claims/EOT exposure.
- Shows all-projects management view when no project is selected.

Project pages:

- Use the active project context.
- Show overview, WBS, activities, milestones, S-Curve, EVM, contracts, letters, risks, TIA, claims, output studio, and data quality/export center.

## 10. Data Quality and Export Center

Module:

`src/construction_system/premium_platform.py`

Purpose:

- Discover source files.
- Clean safe formatting issues.
- Log automatic corrections.
- Run validation checks.
- Allow editable tables with `st.data_editor`.
- Generate readiness score.
- Generate chart pack.
- Export reports and package.

Supported source file types:

- `.csv`
- `.xlsx`
- `.xls`
- `.json`

Validation checks include:

- Empty tables.
- Duplicate identifiers.
- Missing/placeholder values.
- Missing status/owner fields.
- Invalid date parsing.
- Start date after finish date.
- Broken parent/predecessor/successor relationships.
- Orphan references.

Exports:

- Full PDF report.
- Executive summary PDF.
- Technical appendix PDF.
- Chart pack PDF.
- Excel workbook.
- JSON data model.
- CSV exports.
- Chart HTML files.
- Chart PNG files where Plotly/Kaleido supports image export.
- Complete ZIP package.

## 11. Streamlit Main Dashboard

Run locally:

```bat
RUN_APP.bat
```

Or:

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub"
python -m streamlit run dashboard.py --server.port=18755
```

Local URL:

```text
http://localhost:18755
```

Streamlit Cloud:

```text
https://samco-mob-intelligence-dashboard.streamlit.app/
```

Main file path for Streamlit Cloud:

```text
dashboard.py
```

Requirements:

```text
streamlit
pandas
numpy
plotly
python-docx
openpyxl
python-pptx
reportlab
openai
streamlit-cookies-controller
```

## 12. Android Mobile App

Android project folder:

```text
android_app/
```

Package:

```text
com.samco.projectintelligencehub
```

App label:

```text
Project Intelligence Hub
```

Android source files:

- `android_app/app/src/main/AndroidManifest.xml`
- `android_app/app/src/main/java/com/samco/projectintelligencehub/MainActivity.java`
- `android_app/app/src/main/assets/mobile_config.json`
- `android_app/app/src/main/res/values/styles.xml`

Android app type:

Native Android WebView shell.

Web URL loaded:

```text
https://samco-mob-intelligence-dashboard.streamlit.app
```

Features:

- Internet permission.
- Network state permission.
- JavaScript enabled.
- DOM storage enabled.
- Cookie support.
- Third-party cookies enabled.
- External link handling.
- DownloadManager support.
- Offline/error screen.
- Loading progress indicator.
- Back-button handling.
- Portrait orientation.

Build commands:

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\android_app"
$env:JAVA_HOME="C:\Program Files\Eclipse Adoptium\jdk-17.0.19.10-hotspot"
$env:ANDROID_HOME="C:\Users\pc\AppData\Local\Android\Sdk"
$env:Path="C:\Gradle\gradle-8.10.2\bin;$env:JAVA_HOME\bin;$env:ANDROID_HOME\platform-tools;$env:Path"
gradle clean assembleRelease
```

Copy output:

```powershell
Copy-Item "app\build\outputs\apk\release\app-release.apk" "..\dist\Project_Intelligence_Hub.apk" -Force
```

Current director files:

- `dist/Project_Intelligence_Hub.apk`
- `dist/Project_Intelligence_Hub_Android_WhatsApp_Package.zip`
- `dist/ANDROID_DIRECTOR_INSTALL_NOTE.txt`

Desktop copies:

- `C:\Users\pc\OneDrive\Desktop\Project_Intelligence_Hub.apk`
- `C:\Users\pc\OneDrive\Desktop\Project_Intelligence_Hub_Android_WhatsApp_Package.zip`

APK verification command:

```powershell
$env:JAVA_HOME="C:\Program Files\Eclipse Adoptium\jdk-17.0.19.10-hotspot"
$env:Path="$env:JAVA_HOME\bin;$env:Path"
& "C:\Users\pc\AppData\Local\Android\Sdk\build-tools\35.0.0\apksigner.bat" verify --verbose --print-certs "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\dist\Project_Intelligence_Hub.apk"
```

WhatsApp handover message:

```text
Project Intelligence Hub Android App

Please download the attached ZIP, extract it, then tap Project_Intelligence_Hub.apk to install.

If Android asks for permission, allow Install unknown apps for WhatsApp or Files, then continue installation.

After installation, open Project Intelligence Hub and login with the username/password shared separately.
```

## 13. iPhone / iOS Route

Free iPhone route:

PWA/Add to Home Screen.

PWA folder:

```text
pwa/
```

PWA files:

- `pwa/manifest.json`
- `pwa/service-worker.js`
- `pwa/index.html`
- `pwa/icons/icon.svg`
- `pwa/README_PWA.md`

iPhone install steps:

1. Open Safari on iPhone.
2. Open `https://samco-mob-intelligence-dashboard.streamlit.app/`.
3. Tap Share.
4. Tap Add to Home Screen.
5. Name it `Project Intelligence Hub`.
6. Open it like an app.
7. Login with provided credentials.

Native iOS IPA status:

No native IPA is generated in this free Windows package. Native iOS IPA signing/distribution requires macOS/Xcode and Apple signing. Professional App Store distribution requires Apple Developer Program.

iOS source/package notes:

- `ios_app_source/`
- `package_ios_source_no_git.ps1`
- `IPHONE_PWA_STATUS.md`

## 14. Desktop App / Launchers

Desktop launcher folder:

```text
desktop_app/
```

Desktop files:

- `desktop_app/Launch_Project_Intelligence_Hub.bat`
- `desktop_app/Launch_Project_Intelligence_Hub.ps1`
- `desktop_app/Create_Windows_Desktop_Shortcut.ps1`
- `desktop_app/README_WINDOWS_DESKTOP_APP.md`

Desktop shortcut/bat:

```text
C:\Users\pc\OneDrive\Desktop\Project Intelligence Hub.lnk
C:\Users\pc\OneDrive\Desktop\Project Intelligence Hub Desktop App.bat
```

## 15. GitHub Repository and No-Git Sync

Repository:

```text
https://github.com/ahmedlabib33-boop/mobile-app
```

Sync engine:

```text
tools/github_no_git_sync.ps1
```

Sync config:

```text
tools/github_sync_config.json
```

Target:

- Owner: `ahmedlabib33-boop`
- Repository: `mobile-app`
- Branch: `main`

Credentials:

The sync reads only:

- `PIH_MOBILE_APP_GITHUB_TOKEN`
- `PIH_MOBILE_APP_GH_TOKEN`

The token must have repository contents read/write access.

Direct one-time sync:

```bat
RUN_FULL_PROJECT_NO_GIT_SYNC.bat
```

30-second watcher:

```bat
RUN_FULL_PROJECT_WATCH_30SEC_NO_GIT.bat
```

Dry-run test:

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\github_no_git_sync.ps1 -Mode DryRun -IntervalSeconds 30
```

Sync logic:

- Scans included workspace files.
- Calculates SHA-256 local manifest.
- Compares with previous local manifest.
- Compares local Git blob SHA with remote GitHub tree.
- Uploads changed/new files.
- Deletes remote files removed locally when deletion sync is enabled.
- Writes sync log to `11-outputs/logs/pih_mobile_app_github_sync.log`.

CSV/data edit behavior:

CSV files are included. If a number inside a CSV changes, the SHA-256 changes and the sync detects/pushes it.

Excluded from sync:

- `_RETURN_POINTS`
- `.venv`
- `runtime`
- `dist`
- `android_signing`
- `android_app/app/build`
- `android_app/.gradle`
- `.env`
- `.env.local`
- `*.jks`
- `*.keystore`
- token/secret files

Reason:

This prevents leaking secrets, local auth DBs, signing keys, generated APKs, and build artifacts into GitHub.

## 16. Return Points and Recovery

Return point folder:

```text
_RETURN_POINTS/
```

Current return marker:

- `RETURN_POINT_CURRENT.json`
- `RETURN_POINT_CURRENT.md`

Known return points:

- `_RETURN_POINTS/return_point_20260629_200640`
- `_RETURN_POINTS/return_point_before_sync_rules_20260630_133647`

User instruction:

If the user types `RETURN`, restore to the current marked return point.

## 17. Report and Export Systems

Main report/export areas:

- Output Studio in `dashboard.py`
- Data Quality & Export Center in `premium_platform.py`
- TIA Director Pack generator in `reports/tia_director_pack_generator.py`
- Word template replacement utilities in `reports/word_template_exporter.py`
- UI report generator helpers in `ui/report_generator_page.py`

TIA director report output:

- Uses Word template: `reports/templates/time_impact_analysis_report_director_pack.docx`
- Tracks generated outputs in SQLite table `generated_outputs`
- Supports contractor display value `SAMCO - NATIONAL`

## 18. Contract and Claims Intelligence

Main file:

```text
contract_claims_center.py
```

Project-scoped folders:

```text
05-contracts/source/
05-contracts/clauses/
06-evidence/
11-outputs/
```

Features:

- Contract clause library.
- Clause search.
- Claim draft support.
- Evidence matrix.
- Rebuttal generation.
- Contractual entitlement mapping.
- Export to Excel/JSON/CSV/DOCX/PDF/HTML where available.

OpenAI note:

OpenAI-backed functions require valid OpenAI credentials where used. The rest of the app remains usable without paid third-party services.

## 19. Letters Intelligence

Main module:

```text
src/construction_system/letters_auto_ingest.py
```

Folders:

```text
07-letters_intelligence/inbox/From Contractor/
07-letters_intelligence/inbox/From Consultant/
```

Purpose:

- Reads correspondence references.
- Builds thread relationships.
- Classifies SAMCO/ACE directionality.
- Produces links and risk signals.

## 20. Delay / TIA / Steel Delay System

Main module:

```text
src/construction_system/steel_delay_tia.py
```

Delay/TIA data folders:

```text
02-delay_analysis/steel_delay_tia_templates/
03-schedule/
```

Core workflow:

- Load employer/client supply and project schedule evidence.
- Map fields.
- Run steel delay TIA analysis.
- Detect affected activities.
- Produce delay/TIA outputs and director pack support.

Important rule:

Contractor-supplied steel is treated as mitigation/visibility only and excluded from employer-steel entitlement calculations unless independently supported.

## 21. Running Prompts and Commands

### Main Streamlit Dashboard

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub"
python -m streamlit run dashboard.py --server.port=18755
```

Or:

```bat
RUN_APP.bat
```

### Admin Console

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub"
python -m streamlit run admin_app.py --server.port=18756
```

Or:

```bat
RUN_ADMIN_CONSOLE.bat
```

### Direct GitHub Sync

```bat
RUN_FULL_PROJECT_NO_GIT_SYNC.bat
```

### 30-Second GitHub Watcher

```bat
RUN_FULL_PROJECT_WATCH_30SEC_NO_GIT.bat
```

### Mobile App Sync

```bat
RUN_MOBILE_APP_SYNC.bat Once 30
```

### Full Project Dry Run

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\github_no_git_sync.ps1 -Mode DryRun -IntervalSeconds 30
```

### Android Toolchain Check

```powershell
.\check_android_toolchain_no_git.ps1
```

### Android Build

```powershell
.\build_android_no_git.ps1
```

Or manual Gradle:

```powershell
cd "D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub\android_app"
$env:JAVA_HOME="C:\Program Files\Eclipse Adoptium\jdk-17.0.19.10-hotspot"
$env:ANDROID_HOME="C:\Users\pc\AppData\Local\Android\Sdk"
$env:Path="C:\Gradle\gradle-8.10.2\bin;$env:JAVA_HOME\bin;$env:ANDROID_HOME\platform-tools;$env:Path"
gradle clean assembleRelease
```

### iOS/PWA Package Notes

```powershell
.\package_ios_source_no_git.ps1
```

### Project Doctor

```powershell
.\PROJECT_DOCTOR.ps1
```

### Compile Check

```powershell
python -m compileall -q auth.py admin_app.py dashboard.py src reports ui
```

## 22. Deployment Checklist

Main Streamlit Cloud app:

1. Repo: `ahmedlabib33-boop/mobile-app`
2. Branch: `main`
3. Main file: `dashboard.py`
4. URL: `https://samco-mob-intelligence-dashboard.streamlit.app/`

Admin Streamlit Cloud app:

1. Repo: `ahmedlabib33-boop/mobile-app`
2. Branch: `main`
3. Main file: `admin_app.py`
4. Use separate app URL.
5. Use shared persistence if cloud admin must manage same cloud auth data as main app.

Android:

1. Confirm `android_app/app/src/main/assets/mobile_config.json`.
2. Build release APK.
3. Verify signing.
4. Copy to `dist/Project_Intelligence_Hub.apk`.
5. Send ZIP via WhatsApp if APK attachment is blocked.

iPhone:

1. Open live Streamlit URL in Safari.
2. Add to Home Screen.
3. Use as iPhone PWA app.

GitHub sync:

1. Set `PIH_MOBILE_APP_GITHUB_TOKEN`.
2. Use direct sync or 30-second watcher.
3. Confirm Streamlit Cloud redeploys.

## 23. Security and Privacy Notes

Do not sync:

- `.env`
- `.env.local`
- `.streamlit/secrets.toml`
- signing keystores
- Android signing properties
- runtime auth DB
- generated APKs in `dist`

Why:

These contain local credentials, generated artifacts, or runtime state that should not be published to GitHub.

The sync config currently excludes these files/folders.

## 24. Known Limitations

- Streamlit Cloud local SQLite persistence is not a durable shared production database.
- Separate Streamlit Cloud apps do not share the same local SQLite file automatically.
- Native iOS IPA is not generated on Windows.
- Android APK is a lightweight WebView shell and requires internet access.
- Plotly PNG export depends on image export support such as Kaleido.
- OpenAI-related features require valid OpenAI API setup.

## 25. Current Operational Status

Main dashboard:

- Local run supported.
- Streamlit Cloud URL configured.
- GitHub repo sync supported.

Admin:

- Separate local admin host supported.
- Owner-only admin identity enforced.

Android:

- Native WebView source exists.
- Signed release APK generated locally.
- WhatsApp ZIP handover package exists.

iPhone:

- Free PWA route documented and configured.
- Native IPA not generated.

Sync:

- Direct sync supported.
- 30-second watcher supported.
- CSV numeric changes detected through content hash.

Data:

- Project folder pipeline is active.
- Multiple discovered project manifests exist.
- Contractor display forced to `SAMCO - NATIONAL`.

