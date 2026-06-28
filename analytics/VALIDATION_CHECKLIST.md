# Validation Checklist

## Agent Behavior

- [ ] The agent calls at least one tool before final output.
- [ ] The response includes a prioritized plan.
- [ ] The response includes a risk register with delay-analysis risks.
- [ ] The response includes an owner checklist.
- [ ] The response includes launch copy suggestions for more than one channel.
- [ ] The response asks follow-up questions when details are missing.
- [ ] The response does not invent unavailable dates, costs, activity IDs, entitlement, or contract facts.

## Frontend Flow

- [ ] Product brief, audience, launch date, constraints, assets, and delay context can be edited.
- [ ] Submit button starts a streamed run.
- [ ] Model text appears progressively in the response panel.
- [ ] Tool activity appears in the right activity rail.
- [ ] API errors appear as visible user-facing errors.
- [ ] Mobile layout stacks without horizontal overflow.

## Tool Outputs

- [ ] `extract_tasks_from_brief` detects analytics and delay-analysis tasks.
- [ ] `check_launch_readiness` returns a score, status, rubric, launch timing, and asset count.
- [ ] `generate_owner_checklist` returns engineering, analytics, planning, contracts, and communications owners.
- [ ] `draft_channel_launch_copy` returns internal, executive, client, and standup copy suggestions.

## End-To-End Stream

- [ ] Backend server runs on `http://127.0.0.1:8788`.
- [ ] Frontend dev server runs on `http://127.0.0.1:5177`.
- [ ] `python .\scripts\verify_stream.py` receives at least one `tool_progress` event.
- [ ] `python .\scripts\verify_stream.py` receives at least one `text_delta` event.
- [ ] Stream verification uses the server process with `OPENAI_API_KEY` loaded from `.env.local` or environment.
