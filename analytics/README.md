# analytics

`analytics` is a working OpenAI Agents SDK app for engineering release planning, data analytics, and delay-analysis launch workflows.

The frontend collects:

- product brief
- audience
- launch date
- constraints
- available assets
- delay-analysis context

The backend streams an Agents SDK run that uses tools to extract tasks, check launch readiness, generate owner checklists, and draft channel-specific launch copy. The response includes a prioritized plan, risk register, owner checklist, launch copy suggestions, and follow-up questions when details are missing.

## Structure

```text
analytics/
  src/                         React/Vite frontend
  backend/
    app/
      main.py                  FastAPI API routes
      agent.py                 Agents SDK setup and stream adapter
      tools.py                 Function tools
      schemas.py               Request and health schemas
    tests/                     Backend unit tests
    run_server.py              Local API server entrypoint
  scripts/
    verify_stream.py           End-to-end streamed POST verification
  VALIDATION_CHECKLIST.md
```

## Environment

The backend reads `OPENAI_API_KEY` from the workspace `.env.local` or `.env`.

Optional model override:

```powershell
$env:ANALYTICS_AGENT_MODEL="gpt-5.5"
```

The default model is `gpt-5.5`, aligned with current OpenAI model guidance at implementation time. If your account does not have access to that model, set `ANALYTICS_AGENT_MODEL` to a model available to your project.

## Install

From the workspace root:

```powershell
python -m pip install -r .\analytics\backend\requirements.txt
cd .\analytics
npm install
```

## Run

Terminal 1:

```powershell
cd "C:\Users\pc\OneDrive\Documents\Project Intelligence Hub\analytics\backend"
python .\run_server.py
```

Terminal 2:

```powershell
cd "C:\Users\pc\OneDrive\Documents\Project Intelligence Hub\analytics"
npm run dev
```

Open:

```text
http://127.0.0.1:5177
```

## Verify The Real Agent Stream

Do not stop at `/api/health`. Run the streamed endpoint check:

```powershell
cd "C:\Users\pc\OneDrive\Documents\Project Intelligence Hub\analytics"
python .\scripts\verify_stream.py
```

The script posts a real request to `http://127.0.0.1:8788/api/agent/stream` and fails unless it receives:

- at least one `tool_progress` event
- at least one `text_delta` event

## Tests

```powershell
cd "C:\Users\pc\OneDrive\Documents\Project Intelligence Hub\analytics\backend"
python -m pytest .\tests
```

Frontend build:

```powershell
cd "C:\Users\pc\OneDrive\Documents\Project Intelligence Hub\analytics"
npm run build
```

## Extending

Add new deterministic tools in `backend/app/tools.py`, import them in `backend/app/agent.py`, then update the agent instructions so the model knows when to call them.

Good extension candidates:

- Primavera/XER evidence preflight
- claim notice compliance checker
- schedule-risk scoring
- release owner escalation handoff
- export-to-Word or export-to-Excel report writer
