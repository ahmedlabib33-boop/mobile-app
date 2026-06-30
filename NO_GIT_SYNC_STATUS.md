# No-Git Sync Status

Return point created before these sync-rule changes:

`_RETURN_POINTS/return_point_before_sync_rules_20260630_133647`

## Direct Sync

Run:

`RUN_FULL_PROJECT_NO_GIT_SYNC.bat`

Default behavior is now one immediate sync to:

`ahmedlabib33-boop/mobile-app`

## 30-Second Watch

Run:

`RUN_FULL_PROJECT_WATCH_30SEC_NO_GIT.bat`

This starts a watcher that scans every 30 seconds and pushes changed/new files.

## CSV/Data Changes

CSV files are included by default. If any included `.csv` file changes, including a changed number inside the file, the sync engine detects it by SHA-256 content hash and pushes it to the repository.

## Protected Local Files

The sync intentionally excludes runtime/build/secrets folders and files such as:

- `_RETURN_POINTS`
- `.venv`
- `.tmp`
- `runtime`
- `dist`
- `android_signing`
- `android_app/app/build`
- `android_app/.gradle`
- `.env`
- `.env.local`
- signing keys and keystores

This keeps deployment source and data synced without publishing local secrets or generated build artifacts.
