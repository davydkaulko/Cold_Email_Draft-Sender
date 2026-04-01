#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
STEP 2: Send Scheduler
Reads schedule.json created by step1, waits for each send time,
then sends the corresponding Gmail draft.

HOW TO USE:
  Start:  python step2_send_scheduler.py
  Stop:   Ctrl+C  → progress is saved, already-sent emails are skipped on resume
  Resume: python step2_send_scheduler.py  (runs again, skips already sent)

IMPORTANT: Keep this terminal window open while sending.
           If you close it or the PC sleeps, sending stops.
           Just run again and it continues from where it stopped.
"""

import os
import json
import time
import pickle
import signal
import sys
from datetime import datetime

# ── Dependency check ──────────────────────────────────────────────────────────
try:
    from google.auth.transport.requests import Request
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    import base64
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
except ImportError:
    print("\n❌  Missing Python libraries. Run this command first:")
    print("    pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client\n")
    sys.exit(1)

SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/spreadsheets',
]

TOKEN_FILE  = 'token.pickle'
CREDS_FILE  = 'credentials.json'
SCHEDULE_DB = 'schedule.json'

# ── Graceful Ctrl+C handling ──────────────────────────────────────────────────
_stop_requested = False

def _handle_sigint(sig, frame):
    global _stop_requested
    if _stop_requested:
        print("\n\n🛑 Force quit.")
        sys.exit(0)
    print("\n\n⚠️  Stop requested — finishing current wait, then stopping.")
    print("    Press Ctrl+C again to force quit immediately.")
    _stop_requested = True

signal.signal(signal.SIGINT, _handle_sigint)


# ── Authentication ────────────────────────────────────────────────────────────

def get_credentials():
    """Load saved credentials or ask user to log in via browser."""
    creds = None

    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, 'rb') as f:
            creds = pickle.load(f)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Refreshing expired token...")
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDS_FILE):
                print(f"\n❌  {CREDS_FILE} not found!")
                print("    Download it from Google Cloud Console and put it in this folder.")
                print("    See README.md for step-by-step instructions.\n")
                sys.exit(1)
            print("🔐 Opening browser for Google login...")
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_FILE, 'wb') as f:
            pickle.dump(creds, f)
        print("✅ Login saved — won't need to log in again.\n")

    return creds


# ── Schedule DB helpers ───────────────────────────────────────────────────────

def load_schedule():
    """Load schedule.json. Exits with error if file not found."""
    if not os.path.exists(SCHEDULE_DB):
        print(f"\n❌  {SCHEDULE_DB} not found!")
        print("    You need to run step1_create_drafts.py first to create drafts and the schedule.\n")
        sys.exit(1)

    with open(SCHEDULE_DB, 'r') as f:
        return json.load(f)


def save_schedule(db):
    """Save updated schedule back to schedule.json."""
    with open(SCHEDULE_DB, 'w') as f:
        json.dump(db, f, indent=2)


# ── Gmail helpers ─────────────────────────────────────────────────────────────

def find_draft_by_id(gmail, draft_id):
    """
    Find a Gmail draft by its ID.
    Returns the full draft object, or None if not found.
    """
    try:
        draft = gmail.users().drafts().get(
            userId='me',
            id=draft_id,
            format='full'
        ).execute()
        return draft
    except HttpError:
        return None


def find_draft_by_email(gmail, to_email):
    """
    Fallback: scan all drafts and find one addressed to this email.
    Used when draft_id lookup fails (e.g. draft was reopened/edited in Gmail).
    Returns draft object or None.
    """
    try:
        result = gmail.users().drafts().list(userId='me').execute()
        drafts = result.get('drafts', [])

        for d in drafts:
            detail = gmail.users().drafts().get(
                userId='me',
                id=d['id'],
                format='full'
            ).execute()
            headers = detail['message']['payload'].get('headers', [])
            for h in headers:
                if h['name'] == 'To' and to_email.lower() in h['value'].lower():
                    return detail

    except HttpError as e:
        print(f"   ⚠️  Error scanning drafts: {e}")

    return None


def extract_draft_content(draft):
    """
    Extract To, Subject, and body text from a Gmail draft object.
    Returns (to, subject, body) or (None, None, None) on failure.
    """
    payload = draft['message']['payload']
    headers = {h['name']: h['value'] for h in payload.get('headers', [])}

    to      = headers.get('To', '')
    subject = headers.get('Subject', '')
    body    = ''

    # Try multipart first (most common)
    if 'parts' in payload:
        for part in payload['parts']:
            if part['mimeType'] == 'text/plain':
                data = part.get('body', {}).get('data', '')
                if data:
                    body = base64.urlsafe_b64decode(data).decode('utf-8')
                    break
    # Fallback: single-part body
    elif 'data' in payload.get('body', {}):
        body = base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8')

    if not to or not subject or not body:
        return None, None, None

    return to, subject, body


def send_email_now(gmail, to, subject, body):
    """
    Build and send an email immediately using Gmail API.
    Returns message_id on success, None on failure.
    """
    try:
        msg = MIMEMultipart()
        msg['To']      = to
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')

        result = gmail.users().messages().send(
            userId='me',
            body={'raw': raw}
        ).execute()

        return result['id']

    except HttpError as e:
        print(f"   ❌  Gmail send error: {e}")
        return None


def delete_draft(gmail, draft_id):
    """Delete a Gmail draft after it's been sent (cleanup)."""
    try:
        gmail.users().drafts().delete(userId='me', id=draft_id).execute()
    except HttpError:
        pass  # Not critical if deletion fails


def update_sheet_status(sheets, spreadsheet_id, row, status):
    """Update column F in the Google Sheet with current status."""
    if not sheets or not spreadsheet_id:
        return
    try:
        sheets.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f'Sheet1!F{row}',
            valueInputOption='RAW',
            body={'values': [[status]]}
        ).execute()
    except HttpError as e:
        print(f"   ⚠️  Could not update sheet row {row}: {e}")


# ── Main scheduler loop ───────────────────────────────────────────────────────

def run_scheduler(spreadsheet_id=None):
    """
    Main loop: go through all pending emails in schedule.json,
    wait until each scheduled time, then send.
    """
    global _stop_requested

    db = load_schedule()

    # Collect pending items and sort by send time
    pending = []
    for row_key, item in db.items():
        if item.get('status') == 'pending':
            try:
                send_at = datetime.strptime(item['send_at'], '%Y-%m-%d %H:%M:%S')
                pending.append((row_key, send_at, item))
            except ValueError:
                print(f"⚠️  Skipping row {row_key}: bad timestamp '{item['send_at']}'")

    pending.sort(key=lambda x: x[1])   # sort by scheduled time

    if not pending:
        print("\n📭  Nothing to send — queue is empty.")
        print("    All emails may already be sent, or you need to run step1 first.\n")
        return

    total         = len(pending)
    already_past  = sum(1 for _, t, _ in pending if t <= datetime.now())
    future_items  = [(k, t, i) for k, t, i in pending if t > datetime.now()]

    print(f"\n📋  Queue: {total} email(s) to send")
    if already_past:
        print(f"⚡  {already_past} overdue (will send immediately)")
    if future_items:
        next_t = future_items[0][1]
        next_e = future_items[0][2]['email']
        wait_m = int((next_t - datetime.now()).total_seconds() / 60)
        print(f"⏰  Next scheduled: {next_t.strftime('%Y-%m-%d %H:%M:%S')} → {next_e}  (in ~{wait_m} min)")

    print(f"\n💡  Keep this window open. Ctrl+C to pause — progress is auto-saved.\n")

    # Authenticate
    creds  = get_credentials()
    gmail  = build('gmail',  'v1', credentials=creds)
    sheets = build('sheets', 'v4', credentials=creds) if spreadsheet_id else None

    sent_count  = 0
    error_count = 0

    for i, (row_key, send_at, item) in enumerate(pending, 1):

        if _stop_requested:
            remaining = total - i + 1
            print(f"\n⏸️   Paused. {remaining} email(s) still in queue.")
            print(f"    Run the script again any time to continue.\n")
            break

        email   = item['email']
        subject = item.get('subject', '(no subject)')

        print(f"{'─'*56}")
        print(f"  [{i}/{total}]  Row {row_key}  →  {email}")
        print(f"  Subject:   {subject}")
        print(f"  Scheduled: {send_at.strftime('%Y-%m-%d %H:%M:%S')}")

        # Wait until send time
        now      = datetime.now()
        wait_sec = (send_at - now).total_seconds()

        if wait_sec > 0:
            wait_min = int(wait_sec // 60)
            wait_s   = int(wait_sec % 60)
            print(f"  ⏳  Waiting {wait_min}m {wait_s}s ...")

            deadline = send_at
            while not _stop_requested:
                remaining_sec = (deadline - datetime.now()).total_seconds()
                if remaining_sec <= 0:
                    break
                # Sleep in 15-second chunks so Ctrl+C responds fast
                sleep_chunk = min(15, remaining_sec)
                time.sleep(sleep_chunk)

                # Show countdown every minute
                left = int((deadline - datetime.now()).total_seconds())
                if left > 60 and left % 60 == 0:
                    print(f"      … {left // 60}m left", end='\r')

            print()  # newline after countdown

        if _stop_requested:
            print(f"\n⏸️   Paused before sending row {row_key}.")
            break

        # Find the draft
        print(f"  🔍  Looking up draft...")
        draft = find_draft_by_id(gmail, item.get('draft_id', ''))

        if not draft:
            print(f"  ⚠️   Draft not found by ID — scanning by email address...")
            draft = find_draft_by_email(gmail, email)

        if not draft:
            print(f"  ❌  Draft not found at all. Was it deleted from Gmail?")
            db[row_key]['status'] = 'error_draft_not_found'
            save_schedule(db)
            update_sheet_status(sheets, spreadsheet_id, int(row_key), "Error — draft not found")
            error_count += 1
            print()
            continue

        # Extract content and send
        to, subj, body = extract_draft_content(draft)

        if not body:
            print(f"  ❌  Could not read email body from draft.")
            db[row_key]['status'] = 'error_empty_body'
            save_schedule(db)
            update_sheet_status(sheets, spreadsheet_id, int(row_key), "Error — could not read draft body")
            error_count += 1
            print()
            continue

        print(f"  📤  Sending now...")
        msg_id = send_email_now(gmail, to, subj, body)

        if msg_id:
            # Success!
            sent_at = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            db[row_key]['status']  = 'sent'
            db[row_key]['sent_at'] = sent_at
            db[row_key]['msg_id']  = msg_id
            save_schedule(db)
            update_sheet_status(sheets, spreadsheet_id, int(row_key), f"Sent | {sent_at}")

            # Clean up the draft
            delete_draft(gmail, draft['id'])

            print(f"  ✅  SENT at {sent_at}")
            sent_count += 1
        else:
            db[row_key]['status'] = 'error_send_failed'
            save_schedule(db)
            update_sheet_status(sheets, spreadsheet_id, int(row_key), "Error — send failed")
            print(f"  ❌  Failed to send")
            error_count += 1

        print()

    if not _stop_requested:
        print("=" * 56)
        print(f"🏁  All done!")
        print(f"   ✅ Sent:   {sent_count}")
        print(f"   ❌ Errors: {error_count}")
        print(f"   📊 Total:  {total}")
        print("=" * 56 + "\n")


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    print("\n" + "="*56)
    print("  📧  STEP 2 — SEND SCHEDULER")
    print("  Sends your Gmail drafts at their scheduled times.")
    print("  Keep this window open while sending is running.")
    print("="*56 + "\n")

    if not os.path.exists(SCHEDULE_DB):
        print(f"❌  {SCHEDULE_DB} not found.")
        print("    Run step1_create_drafts.py first!\n")
        return

    print("📊  Enter your Google Sheets ID to update statuses in the sheet.")
    print("    (press Enter to skip — emails will still send, sheet just won't update)")
    sid = input("Spreadsheet ID (or Enter to skip): ").strip() or None

    if sid:
        print(f"✅  Will update sheet: {sid[:20]}...\n")
    else:
        print("⚠️   No sheet ID — statuses won't be updated in Google Sheets.\n")

    run_scheduler(spreadsheet_id=sid)


if __name__ == '__main__':
    main()