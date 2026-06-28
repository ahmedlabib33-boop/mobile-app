# Project Intelligence Hub Operating Prompt

Work only in `C:\Users\pc\OneDrive\Documents\Project Intelligence Hub`.

1. Preserve the existing `Dashboard project` selector and folder discovery.
2. Use `project_id` from `project_manifest.json` as the stable identity everywhere.
3. Use `ProjectContext` to resolve the selected ID to the folder's current actual path.
4. Support folder renames without changing project identity or mixing data.
5. Auto-discover unlimited future folders under `projects`, including `projects/<sector>/<project>` layouts, without code registration.
6. Scope every dashboard loader, Claims Intelligence operation, Delay TIA calculation, report, slide, export, cache, and database to the selected project.
7. In `Decision Making dashboard`, allow only explicit portfolio aggregations retaining `project_id`; keep it as standalone Phase 1 and show project-specific slides only after opening a Phase 2 project workspace.
8. Never fall back to a default, legacy, random, first, or previously selected project.
9. Missing selected-project data must show a clean empty/setup state.
10. Preserve all existing business logic, slides, reports, exports, and formats.
11. Do not overwrite user files. Standard project samples are created only when missing.
12. Maintain `data_to_program.md`, `data_lineage.json`, validation tests, and the active backup change log.

Run app:

```powershell
.\RUN_APP.bat
```

Validate:

```powershell
python -m pytest -q tests
python tools/validate_project_isolation.py --sync-probe
```

Full workspace no-Git repository sync only:

```powershell
.\RUN_MOBILE_APP_SYNC.bat Watch 30
```

Manual repository refresh is also available in Output Studio after `PIH_MOBILE_APP_GITHUB_TOKEN` and `PIH_MOBILE_APP_SYNC_ADMIN_PIN` are configured as environment variables. Never place either secret in code, prompts, configuration JSON, or documentation.

Target: `ahmedlabib33-boop/mobile-app`, branch `main`.

