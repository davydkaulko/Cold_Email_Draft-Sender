

**Set up your spreadsheet**

Columns should be in this order:

| Column | Field         | Notes                              |
|--------|---------------|------------------------------------|
| A      | email         | Required — rows without it skipped |
| B      | first_name    | Optional                           |
| C      | last_name     | Optional                           |
| D      | company       | Defaults to "your company" if empty |
| E      | personal_note | The opening line of the email      |
| F      | status        | Leave empty — script writes here   |

---

## Running a campaign

```bash
# Step 1 — create drafts
python step1_create_drafts.py
```

It will ask for your spreadsheet ID (the long string in the URL between `/d/` and `/edit`). Then pick option 2 from the menu to start from row 2.

After it finishes, go to Gmail → Drafts and read through the emails. Make sure they look right.

```bash
# Step 2 — send on schedule
python step2_send_scheduler.py
```

Keep the terminal open. It will wait until each scheduled time and send. If you close the terminal, sending stops — just run it again and it picks up where it left off, skipping anything already sent.

To stop gracefully, press Ctrl+C once. It will finish the current wait and stop. Press Ctrl+C twice to exit immediately.

---

## Files in this folder

```
credentials.json    — Google API credentials (never commit this)
token.pickle        — saved auth token, generated on first run (never commit this)
schedule.json       — tracks draft IDs, send times, and statuses
```

Add `credentials.json` and `token.pickle` to your `.gitignore`.

---

## Configuring send intervals

Default is 11 to 22 minutes between emails. You can change this in the step 1 menu (option 4) during a session, or edit the defaults directly in the `ColdEmailSender` constructor:

```python
# step1_create_drafts.py, line ~30
def __init__(self, spreadsheet_id, min_interval=11, max_interval=22):
```

The randomization is intentional — a fixed interval looks like a bot to spam filters.

---

## Editing the email template

Everything is in the `create_email_body()` method in `step1_create_drafts.py`. It's a plain f-string with three variables available:

- `{full_name}` — first + last name, falls back to "there" if both are empty
- `{personal_note}` — the content from column E
- company name is used in the subject line only

---

## Resuming after a pause or crash

Just run step 2 again. It reads `schedule.json` and skips anything with `"status": "sent"`. Overdue emails (past their scheduled time) get sent immediately.

If a draft got deleted from Gmail manually, step 2 will log it as `error_draft_not_found` and move on.

---

## Starting a new campaign

Either use a new spreadsheet, or make sure column F is empty (or says "Not sent") for the rows you want to process. Rows with any other status in column F will be skipped.

To completely reset, delete or rename `schedule.json`.

---

## Things to know

- The script only processes rows where column F is empty or says "Not sent"
- Running step 1 twice is safe — already-scheduled rows won't be touched
- step 2 can be stopped and restarted at any time without losing progress
- If you edit a draft in Gmail after step 1 runs, the draft ID changes. step 2 handles this automatically — it falls back to finding the draft by recipient email address
- The spreadsheet ID prompt in step 2 is optional. If you skip it, emails still send but the sheet won't be updated with "Sent" status

---

## Dependencies

- Python 3.8+
- google-auth
- google-auth-oauthlib
- google-auth-httplib2
- google-api-python-client
