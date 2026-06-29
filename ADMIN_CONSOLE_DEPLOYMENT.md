# Project Intelligence Hub Admin Console

The client-facing dashboard and the owner administration screen are now separated.

## Local hosts

- Main dashboard: `RUN_APP.bat` or `python -m streamlit run dashboard.py --server.port=18755`
- Admin console: `RUN_ADMIN_CONSOLE.bat` or `python -m streamlit run admin_app.py --server.port=18756`

The local hosts share the same project folder and therefore use the same SQLite auth database at `runtime/project_intelligence_hub_mobile_auth.sqlite`.

## Streamlit Cloud

Create a second Streamlit Cloud app from the same GitHub repository:

- Repository: `ahmedlabib33-boop/mobile-app`
- Branch: `main`
- Main file path: `admin_app.py`
- Suggested app name: `samco-project-intelligence-admin`

The existing public/mobile app remains:

- Main file path: `dashboard.py`
- URL: `https://samco-mob-intelligence-dashboard.streamlit.app/`

## Important free-hosting limitation

Two separate Streamlit Cloud apps do not automatically share the same local SQLite database. The admin console is fully separated as requested, but cloud-wide shared user persistence requires a shared database/storage backend. Locally, both hosts share the same SQLite database because they run from the same folder.

## Watcher behavior

The existing `RUN_MOBILE_APP_SYNC.bat` 30-second watcher pushes code changes for both apps to GitHub. Streamlit Cloud will reflect the new admin entry point after redeploying the second app from `admin_app.py`.
