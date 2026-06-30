/godmode
//ooda

Project path:
D:\New SAMCO Dashboard\Mobile App\Project Intelligence Hub

Task:
Integrate the prepared Keep me signed in feature into the existing Streamlit authentication flow.

Prepared file:
remember_me.py

Requirements:
1. Import remember_me.py in auth.py or dashboard.py.
2. At app startup, before showing login page, call:
   remember_me.restore_session_from_remember_cookie()
3. On login screen, add checkbox:
   Keep me signed in
   Note: Recommended for trusted devices only.
4. If login succeeds and checkbox is checked, call:
   remember_me.issue_remember_cookie_for_user(user_id, user_agent="", device_label="Trusted device")
5. On Sign out button:
   call remember_me.sign_out_and_clear_session()
   then rerun app.
6. Add admin session management panel:
   - list remembered sessions using remember_me.list_remember_sessions()
   - revoke selected session using remember_me.revoke_session_by_id(session_id)
   - revoke all sessions for a user using remember_me.revoke_all_user_tokens(user_id)
   - revoke all sessions globally using remember_me.revoke_all_tokens()
7. Do not store passwords in cookies/localStorage.
8. Do not store raw tokens in SQLite.
9. Ensure Streamlit Cloud works.
10. Ensure iPhone PWA keeps login after close/reopen.
11. Ensure Android WebView keeps login after close/reopen.
12. User must be signed out only when clicking Sign out, admin revokes token, account disabled, or token expired.

Run:
python -m py_compile remember_me.py
python -m py_compile auth.py
python -m py_compile dashboard.py
python -c "import remember_me; print('remember_me import ok')"

Update:
KEEP_ME_SIGNED_IN_STATUS.md
MOBILE_BUILD_STATUS.md
pwa\README_PWA.md
DIRECTOR_HANDOVER_MESSAGE.md

Final report:
- Keep me signed in added: Yes/No
- Works on web: Yes/No
- Works on iPhone PWA: Yes/No
- Works on Android WebView: Yes/No
- Sign out clears persistent login: Yes/No
- Admin revoke sessions added: Yes/No
- requirements.txt updated: Yes/No
- files modified
- files created
- blockers