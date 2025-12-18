"""
Quick automated test - sends HTML email for Harriet to your test email
"""

from lab_report_sender import LabReportSender
from email_config import EmailConfig
import pandas as pd

# File paths
grading_file = "C:/Users/Franz-JamesWefagiKab/OneDrive - AmaliTech gGmbH/Lab Materials/BE Lab Grading Sheet.xlsx"
email_list_file = "C:/Users/Franz-JamesWefagiKab/OneDrive - AmaliTech gGmbH/Lab Materials/Quality Assurance - Copy.xlsx"

# Your test email
TEST_EMAIL = "franzjameswefagikaba@gmail.com"

# Initialize
sender = LabReportSender(grading_file, email_list_file)
config = EmailConfig()

# Load Module-2 data
print("Loading Module-2 data...")
sender.load_grading_data('Module-2')

# Find Harriet (row 31, but pandas 0-indexed after headers, so it depends)
print("Looking for Harriet Effah...")

harriet_row = None
for idx, row in sender.grading_data.iterrows():
    nsp_name = row.get('Name of NSP', '')
    if 'Harriet' in str(nsp_name):
        harriet_row = row
        print(f"Found: {nsp_name}")
        break

if harriet_row is None:
    print("Harriet not found in Module-2!")
    exit(1)

# Generate email
print("\nGenerating HTML email...")
subject, body = sender.generate_email_content(harriet_row)

# Add TEST tag
subject = f"{subject} [TEST]"

# Save HTML to file for inspection
print("Saving HTML to preview_email.html...")
with open('preview_email.html', 'w', encoding='utf-8') as f:
    f.write(body)
print("[OK] Saved! You can open preview_email.html in a browser to see how it looks.")

# Preview
print(f"\n{'='*80}")
print(f"PREVIEW")
print(f"{'='*80}")
print(f"To: {TEST_EMAIL}")
print(f"Subject: {subject}")
print(f"Body length: {len(body)} characters")
print(f"{'='*80}\n")

# Check for saved credentials
if not config.has_config():
    print("ERROR: No saved credentials found!")
    print("Please run: python lab_report_sender.py")
    print("And select option 2 to configure your email credentials first.")
    exit(1)

# Get credentials
smtp_server = config.get_smtp_server()
smtp_port = config.get_smtp_port()
sender_email = config.get_email()
sender_password = config.get_password()

print(f"Using saved credentials: {sender_email}")
print(f"\nSending test email to {TEST_EMAIL}...")
print("This will show you how the HTML email looks in your actual inbox.\n")

# Prepare email data
email_data = {
    'to': TEST_EMAIL,
    'to_name': 'Harriet Effah (TEST)',
    'subject': subject,
    'body': body
}

# Send
try:
    sender.send_emails([email_data], smtp_server, smtp_port, sender_email, sender_password)
    print("\n[OK] Test email sent successfully!")
    print(f"[OK] Check your inbox at {TEST_EMAIL}")
    print(f"[OK] Also check preview_email.html to see the HTML structure")
except Exception as e:
    print(f"\n[ERROR] Failed to send: {e}")
    print(f"But you can still view the HTML by opening preview_email.html")
