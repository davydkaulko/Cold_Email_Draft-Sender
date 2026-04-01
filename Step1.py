#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cold Email Automation with Gmail API
STEP 1: Creates Gmail drafts and saves schedule to schedule.json
"""

import os
import pickle
import random
import time
import json
from datetime import datetime, timedelta
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import base64

# Gmail API scopes
SCOPES = [
    'https://www.googleapis.com/auth/gmail.send',
    'https://www.googleapis.com/auth/gmail.compose',
    'https://www.googleapis.com/auth/spreadsheets'
]

SCHEDULE_DB = 'schedule.json'


class ColdEmailSender:
    def __init__(self, spreadsheet_id, min_interval=11, max_interval=22):
        """
        Initialize cold email sender

        Args:
            spreadsheet_id: Google Sheets spreadsheet ID
            min_interval: minimum interval between sends (minutes)
            max_interval: maximum interval between sends (minutes)
        """
        self.spreadsheet_id = spreadsheet_id
        self.min_interval = min_interval
        self.max_interval = max_interval
        self.gmail_service = None
        self.sheets_service = None
        self.running = True

    def authenticate_gmail(self):
        """Authenticate Gmail API"""
        print("🔐 Connecting to Gmail API...")
        creds = None

        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)

            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.gmail_service = build('gmail', 'v1', credentials=creds)
        print("✅ Gmail API connected successfully\n")

    def authenticate_sheets(self):
        """Authenticate Google Sheets API"""
        print("📊 Connecting to Google Sheets API...")
        creds = None

        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)

            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.sheets_service = build('sheets', 'v4', credentials=creds)
        print("✅ Google Sheets API connected successfully\n")

    def get_full_name(self, first_name, last_name):
        """
        Combine first and last name intelligently

        Args:
            first_name: first name (can be empty)
            last_name: last name (can be empty)

        Returns:
            str: combined name or 'there' if both empty
        """
        first = (first_name or '').strip()
        last = (last_name or '').strip()

        if first and last:
            return f"{first} {last}"
        elif first:
            return first
        elif last:
            return last
        else:
            return "there"  # fallback if no name

    def get_spreadsheet_data(self, start_row=2):
        """
        Get data from spreadsheet

        Args:
            start_row: starting row (default 2, row 1 is headers)

        Returns:
            list: list of contact dictionaries
        """
        try:
            range_name = f'Sheet1!A{start_row}:F1000'
            result = self.sheets_service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheet_id,
                range=range_name
            ).execute()

            values = result.get('values', [])

            if not values:
                print("⚠️  No data in spreadsheet")
                return []

            contacts = []
            for idx, row in enumerate(values, start=start_row):
                # Skip if no email (minimum requirement)
                if len(row) < 1 or not row[0]:
                    continue

                # Extract data with safe defaults
                email = row[0].strip() if len(row) > 0 else ''
                first_name = row[1].strip() if len(row) > 1 else ''
                last_name = row[2].strip() if len(row) > 2 else ''
                company = row[3].strip() if len(row) > 3 else 'your company'
                personal_note = row[4].strip() if len(row) > 4 else ''
                status = row[5].strip() if len(row) > 5 else 'Not sent'

                # Skip if no email
                if not email:
                    continue

                # Get full name
                full_name = self.get_full_name(first_name, last_name)

                contact = {
                    'row_number': idx,
                    'email': email,
                    'first_name': first_name,
                    'last_name': last_name,
                    'full_name': full_name,
                    'company': company,
                    'personal_note': personal_note,
                    'status': status
                }
                contacts.append(contact)

            return contacts

        except HttpError as error:
            print(f"❌ Error reading spreadsheet: {error}")
            return []

    def update_status(self, row_number, status, scheduled_time=None):
        """
        Update send status in spreadsheet

        Args:
            row_number: row number
            status: new status
            scheduled_time: scheduled send time
        """
        try:
            range_name = f'Sheet1!F{row_number}'
            value = status
            if scheduled_time:
                value = f"{status} | {scheduled_time}"

            body = {
                'values': [[value]]
            }

            self.sheets_service.spreadsheets().values().update(
                spreadsheetId=self.spreadsheet_id,
                range=range_name,
                valueInputOption='RAW',
                body=body
            ).execute()

        except HttpError as error:
            print(f"❌ Error updating status: {error}")

    def create_email_body(self, full_name, company, personal_note):
        """
        Create email body from template

        Args:
            full_name: recipient full name
            company: company name
            personal_note: personalized message

        Returns:
            str: formatted email body
        """
        # If no personal note, use generic greeting
        if not personal_note:
            personal_note = f""

        template = f"""Hi {full_name},

{personal_note}

I'll be direct: Right now, we are looking to build a few high-impact case studies. We want to find ambitious companies where we can automate core processes or solve problems quickly and effectively.

I lead AVI Collective. We aren't a typical agency with account managers and layers of red tape. We are a decentralized group of specialists obsessed with high-velocity solutions and killing manual friction.

The Win-Win:
For you: You get a custom solution that solves a real bottleneck, delivered fast and at a significantly lower cost than a traditional firm would charge.
For us: We get a powerful, real-world case study to add to our portfolio.

We don't do "corporate fluff." We focus on pure logic and shipping fast. You can get a feel for our team and how we work here: https://avic.netlify.app/

What's currently eating up your team's time? If you have a task or workflow that feels like it should be fixed, I'd love to hop on a 15-minute chat. No formal pitch just tell me what's bothering you, and I'll tell you if we can fix it.

Davyd Kaulko, founder of AVI Collective"""

        return template

    def create_draft_email(self, to_email, subject, body):
        """
        Create draft email in Gmail

        Args:
            to_email: recipient email
            subject: email subject
            body: email body

        Returns:
            str: draft ID or None
        """
        try:
            message = MIMEMultipart()
            message['To'] = to_email
            message['Subject'] = subject

            msg = MIMEText(body, 'plain')
            message.attach(msg)

            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')

            draft = self.gmail_service.users().drafts().create(
                userId='me',
                body={'message': {'raw': raw_message}}
            ).execute()

            return draft['id']

        except HttpError as error:
            print(f"❌ Error creating draft: {error}")
            return None

    def save_to_schedule_db(self, row_number, email, draft_id, subject, send_time_str):
        """
        Save scheduled email info to local schedule.json file.
        This file is used by step2_send_scheduler.py to know what to send and when.

        Args:
            row_number: spreadsheet row number (used as key)
            email: recipient email
            draft_id: Gmail draft ID
            subject: email subject
            send_time_str: scheduled send time as string 'YYYY-MM-DD HH:MM:SS'
        """
        db = {}
        if os.path.exists(SCHEDULE_DB):
            try:
                with open(SCHEDULE_DB, 'r') as f:
                    db = json.load(f)
            except Exception:
                db = {}

        db[str(row_number)] = {
            'email': email,
            'draft_id': draft_id,
            'subject': subject,
            'send_at': send_time_str,
            'status': 'pending'
        }

        with open(SCHEDULE_DB, 'w') as f:
            json.dump(db, f, indent=2)

    def schedule_email(self, to_email, subject, body, send_time):
        """
        Create draft email and save to schedule.
        Does NOT send anything — only creates draft + saves time to schedule file.

        Args:
            to_email: recipient email
            subject: email subject
            body: email body
            send_time: scheduled send time

        Returns:
            tuple: (success bool, draft_id or None)
        """
        draft_id = self.create_draft_email(to_email, subject, body)

        if draft_id:
            print(f"   📝 Draft created (ID: {draft_id[:10]}...)")
            return True, draft_id
        else:
            print(f"   ❌ Failed to create draft")
            return False, None

    def send_test_email(self, test_email="your_email@gmail.com"):
        """
        Send test email to yourself

        Args:
            test_email: test email address
        """
        print("\n📧 TEST EMAIL")
        print("=" * 50)

        subject = "Partnership Proposal: Test Company x AVI Collective"
        body = self.create_email_body(
            "John Smith",
            "Test Company",
            "This is a test personal message to verify email formatting and content."
        )

        print(f"\n📬 Subject: {subject}")
        print(f"📬 Recipient: {test_email}")
        print(f"\n📝 Email body:\n{'-' * 50}\n{body}\n{'-' * 50}\n")

        confirm = input("Send test email? (y/n): ").lower()
        if confirm == 'y':
            draft_id = self.create_draft_email(test_email, subject, body)
            if draft_id:
                print("✅ Test draft created! Check Gmail Drafts folder.")
            else:
                print("❌ Error creating test draft")

    def calculate_next_send_time(self, current_time):
        """
        Calculate next send time with random interval

        Args:
            current_time: current time

        Returns:
            tuple: (next_datetime, interval_minutes)
        """
        interval_minutes = random.randint(self.min_interval, self.max_interval)
        next_time = current_time + timedelta(minutes=interval_minutes)
        return next_time, interval_minutes

    def process_emails(self, start_row=2, stop_row=None):
        """
        Process contacts and create draft emails.
        Also saves schedule to schedule.json for step2 to use.

        Args:
            start_row: starting row
            stop_row: ending row (None = until end)
        """
        print("\n🚀 CREATING DRAFTS")
        print("=" * 50)
        print(f"⏰ Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"📍 Start row: {start_row}")
        print(f"⏸️  End row: {stop_row if stop_row else 'Until end of spreadsheet'}")
        print(f"⏱️  Scheduled interval: {self.min_interval}-{self.max_interval} minutes")
        print("=" * 50)
        print("\n💡 Press Ctrl+C to stop")
        print("💡 This creates drafts ONLY — not sending yet")
        print(f"💡 Schedule will be saved to {SCHEDULE_DB}\n")

        contacts = self.get_spreadsheet_data(start_row)

        if not contacts:
            print("❌ No contacts to process")
            return

        pending_contacts = [
            c for c in contacts
            if c['status'].startswith('Not sent') or c['status'] == ''
        ]

        if stop_row:
            pending_contacts = [
                c for c in pending_contacts
                if c['row_number'] <= stop_row
            ]

        print(f"📊 Contacts to process: {len(pending_contacts)}\n")

        if not pending_contacts:
            print("⚠️  No pending contacts found.")
            print("    Check that column F says 'Not sent' or is empty for rows you want to process.")
            return

        current_send_time = datetime.now() + timedelta(minutes=2)
        created_count = 0
        error_count = 0

        for idx, contact in enumerate(pending_contacts, 1):
            try:
                row_num = contact['row_number']

                print(f"{'='*50}")
                print(f"📧 [{idx}/{len(pending_contacts)}] Processing row {row_num}")
                print(f"{'='*50}")
                print(f"   👤 Name: {contact['full_name']}")
                print(f"   🏢 Company: {contact['company']}")
                print(f"   📧 Email: {contact['email']}")

                note_preview = contact['personal_note'][:50] if contact['personal_note'] else '[No personal note]'
                print(f"   📝 Note: {note_preview}...")

                subject = f"Partnership Proposal: {contact['company']} x AVI Collective"
                body = self.create_email_body(
                    contact['full_name'],
                    contact['company'],
                    contact['personal_note']
                )

                scheduled_time_str = current_send_time.strftime('%Y-%m-%d %H:%M:%S')
                print(f"   ⏰ Scheduled for: {scheduled_time_str}")

                success, draft_id = self.schedule_email(
                    contact['email'],
                    subject,
                    body,
                    current_send_time
                )

                if success:
                    # Update spreadsheet status
                    self.update_status(
                        row_num,
                        "Scheduled",
                        scheduled_time_str
                    )

                    # Save to local schedule.json so step2 can find it
                    self.save_to_schedule_db(
                        row_num,
                        contact['email'],
                        draft_id,
                        subject,
                        scheduled_time_str
                    )

                    print(f"   ✅ Draft created successfully")
                    created_count += 1

                    next_time, interval = self.calculate_next_send_time(current_send_time)
                    print(f"   ⏭️  Next draft in {interval} min (scheduled send time)")
                    current_send_time = next_time
                else:
                    self.update_status(row_num, "Error")
                    error_count += 1
                    print(f"   ❌ Error creating draft")

                print()
                time.sleep(1)

            except KeyboardInterrupt:
                print("\n\n⚠️  Stopped by user...")
                print(f"✅ Processed: {idx-1} of {len(pending_contacts)}")
                print(f"📍 Last processed row: {contact['row_number']}")
                break
            except Exception as e:
                print(f"   ❌ Unexpected error: {e}")
                self.update_status(contact['row_number'], "Error")
                error_count += 1
                continue

        print("\n" + "="*50)
        print("🏁 DRAFTS CREATED")
        print("="*50)
        print(f"⏰ End time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"✅ Successfully created: {created_count}")
        print(f"❌ Errors: {error_count}")
        print(f"\n💡 Next steps:")
        print(f"   1. Open Gmail → Drafts to verify your emails look correct")
        print(f"   2. Run  python step2_send_scheduler.py  to start sending")
        print(f"   3. Keep terminal open while step2 is running\n")


def main():
    """Main function"""
    print("\n" + "="*50)
    print("📧 STEP 1: CREATE DRAFTS")
    print("="*50 + "\n")

    print("💡 This script ONLY creates drafts in Gmail")
    print("💡 Run step2_send_scheduler.py after this to actually send them\n")

    print("📊 Enter Google Sheets spreadsheet ID:")
    print("   (Copy it from your spreadsheet URL — the long string between /d/ and /edit)")
    print("   Example: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms")
    spreadsheet_id = input("\nSpreadsheet ID: ").strip()

    if not spreadsheet_id:
        print("❌ Spreadsheet ID not provided!")
        return

    sender = ColdEmailSender(spreadsheet_id)

    sender.authenticate_gmail()
    sender.authenticate_sheets()

    while True:
        print("\n" + "="*50)
        print("MENU")
        print("="*50)
        print("1. 📧 Send test email (check that template looks right)")
        print("2. 🚀 Start campaign (from row 2)")
        print("3. ▶️  Continue campaign (from a specific row)")
        print("4. ⚙️  Send intervals (default: 11-22 min)")
        print("5. 🚪 Exit")
        print("="*50)

        choice = input("\nSelect action (1-5): ").strip()

        if choice == '1':
            test_email = input("Enter your email to send test to: ").strip()
            if not test_email:
                test_email = "me"
            sender.send_test_email(test_email)

        elif choice == '2':
            confirm = input("\n⚠️  Start campaign from beginning (row 2)? (y/n): ").lower()
            if confirm == 'y':
                stop = input("Stop at row? (press Enter to process all): ").strip()
                stop_row = int(stop) if stop.isdigit() else None
                sender.process_emails(start_row=2, stop_row=stop_row)

        elif choice == '3':
            start = input("Start from which row number? ").strip()
            if start.isdigit():
                stop = input("Stop at row? (press Enter to process all): ").strip()
                stop_row = int(stop) if stop.isdigit() else None
                sender.process_emails(start_row=int(start), stop_row=stop_row)
            else:
                print("❌ Invalid row number — must be a number like 5 or 10")

        elif choice == '4':
            print(f"\nCurrent interval: {sender.min_interval}–{sender.max_interval} minutes between emails")
            min_int = input(f"New minimum interval (minutes): ").strip()
            max_int = input(f"New maximum interval (minutes): ").strip()
            if min_int.isdigit() and max_int.isdigit():
                sender.min_interval = int(min_int)
                sender.max_interval = int(max_int)
                print(f"✅ Updated: {sender.min_interval}–{sender.max_interval} minutes")
            else:
                print("❌ Invalid values — enter numbers only")

        elif choice == '5':
            print("\n👋 Goodbye!")
            break

        else:
            print("❌ Invalid choice, enter a number 1-5")


if __name__ == "__main__":
    main()