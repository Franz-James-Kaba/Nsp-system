"""
Lab Report Sender - Main Script
Sends graded lab reports to NSPs via email
"""

import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
from typing import Dict, List, Tuple
import getpass
from datetime import datetime
from email_config import EmailConfig


class LabReportSender:
    def __init__(self, grading_file: str, email_list_file: str):
        self.grading_file = grading_file
        self.email_list_file = email_list_file
        self.email_list = None
        self.grading_data = None

    def load_email_list(self):
        """Load NSP email list from Excel file"""
        print("Loading NSP email list...")
        df = pd.read_excel(self.email_list_file, sheet_name='QA Class List')
        self.email_list = df[['Full Name', 'AmaliTech Email']].copy()
        print(f"Loaded {len(self.email_list)} NSPs")
        return self.email_list

    def load_grading_data(self, module_name: str):
        """Load grading data from specific module sheet"""
        print(f"Loading grading data from {module_name}...")

        # Read Excel with proper structure:
        # Row 1 (index 0): Allocated Points
        # Row 2 (index 1): Passing Score, weights for 1st attempt
        # Row 3 (index 2): weights for 2nd attempt
        # Row 4 (index 3): Column headers (Review Date, Name of NSP, Reviewer, Lab Title, Attempt, rubrics, Total Score, etc.)
        # Row 5+ (index 4+): Actual data

        # Use row 3 (0-indexed) as header, which is Excel row 4
        df_data = pd.read_excel(self.grading_file, sheet_name=module_name, header=3)

        self.grading_data = df_data
        print(f"Loaded {len(df_data)} grading records")
        return df_data

    def match_nsp_email(self, nsp_name: str) -> str:
        """Match NSP name to email address"""
        if self.email_list is None:
            self.load_email_list()

        # Clean up name for matching
        nsp_name_clean = str(nsp_name).strip().lower()

        # Try exact match first
        match = self.email_list[self.email_list['Full Name'].str.strip().str.lower() == nsp_name_clean]

        if not match.empty:
            return match.iloc[0]['AmaliTech Email']

        # Try partial match - check if grading name is IN email list name
        match = self.email_list[self.email_list['Full Name'].str.lower().str.contains(nsp_name_clean, na=False)]

        if not match.empty:
            return match.iloc[0]['AmaliTech Email']

        # Try reverse partial match - check if any part of email list name is IN grading name
        # This catches cases like "Bernice Mawuena" in grading matching "Bernice Adime Mawuena" in email list
        for idx, row in self.email_list.iterrows():
            email_name = str(row['Full Name']).lower().strip()
            # Split both names into parts and check for significant overlap
            email_parts = set(email_name.split())
            grading_parts = set(nsp_name_clean.split())

            # If at least 2 name parts match (e.g., "bernice" and "mawuena"), consider it a match
            if len(email_parts.intersection(grading_parts)) >= 2:
                return row['AmaliTech Email']

        return None

    def is_grade_complete(self, row) -> Tuple[bool, str]:
        """
        Check if a student's grade is complete (has scores AND remarks)
        Returns: (is_complete, reason_if_incomplete)
        """
        # Check if total score exists and is > 0
        total_score = row.get('Total Score', 0)
        try:
            total_score_float = float(total_score) if pd.notna(total_score) else 0.0
        except:
            total_score_float = 0.0

        if total_score_float == 0:
            return False, "No total score"

        # Check if at least one rubric score exists
        excluded_cols = ['Review Date', 'Name of NSP', 'Reviewer', 'Lab Title', 'Attempt',
                        'Total Score', 'Re-do Lab', 'Plagiarism Check', 'Remarks: Strengths',
                        'Remarks: Gaps', 'Other Remarks']

        has_rubric_score = False
        for col in row.index:
            if col not in excluded_cols:
                score = row.get(col)
                try:
                    if pd.notna(score):
                        float(score)  # Try to convert to float
                        has_rubric_score = True
                        break
                except (ValueError, TypeError):
                    pass

        if not has_rubric_score:
            return False, "No rubric scores"

        # Check if remarks exist (at least Strengths or Gaps)
        strengths = row.get('Remarks: Strengths', '')
        gaps = row.get('Remarks: Gaps', '')

        has_remarks = pd.notna(strengths) or pd.notna(gaps)

        if not has_remarks:
            return False, "No remarks/feedback"

        return True, ""

    def generate_email_content(self, row) -> Tuple[str, str]:
        """Generate email subject and body from grading row"""

        # Extract basic data using column names
        nsp_name = row.get('Name of NSP', 'NSP')
        reviewer = row.get('Reviewer', 'N/A')
        lab_title = row.get('Lab Title', 'Lab')
        attempt = row.get('Attempt', 'N/A')
        total_score = row.get('Total Score', 0)
        redo_lab = row.get('Re-do Lab', 'No')
        plagiarism = row.get('Plagiarism Check', 'N/A')
        strengths = row.get('Remarks: Strengths', '')
        gaps = row.get('Remarks: Gaps', '')
        other_remarks = row.get('Other Remarks', '')

        # Determine pass/fail status
        passing_score = 0.8
        try:
            total_score_float = float(total_score) if pd.notna(total_score) else 0.0
        except:
            total_score_float = 0.0

        status = "PASSED" if total_score_float >= passing_score else "NEEDS RE-DO"

        # Get rubric scores
        rubric_data = []
        for col in row.index:
            if col not in ['Review Date', 'Name of NSP', 'Reviewer', 'Lab Title', 'Attempt',
                          'Total Score', 'Re-do Lab', 'Plagiarism Check', 'Remarks: Strengths',
                          'Remarks: Gaps', 'Other Remarks']:
                score = row.get(col)
                try:
                    if pd.notna(score):
                        score_val = float(score)
                        rubric_data.append((col, int(score_val) if score_val == int(score_val) else score_val))
                except (ValueError, TypeError):
                    pass

        # Create subject
        subject = f"Lab Grade: {lab_title} - {status}"

        # Generate HTML email body
        body = self._generate_html_email(
            nsp_name, lab_title, status, total_score_float, passing_score,
            attempt, redo_lab, plagiarism, rubric_data, strengths, gaps, other_remarks
        )

        return subject, body

    def _generate_html_email(self, nsp_name, lab_title, status, total_score, passing_score,
                            attempt, redo_lab, plagiarism, rubric_data, strengths, gaps, other_remarks):
        """Generate beautifully formatted HTML email"""

        # Status colors
        status_color = "#28a745" if status == "PASSED" else "#dc3545"
        status_bg = "#d4edda" if status == "PASSED" else "#f8d7da"
        status_icon = "✓" if status == "PASSED" else "✗"

        # Score percentage
        score_percent = int(total_score * 100)
        passing_percent = int(passing_score * 100)

        # Build rubric rows
        rubric_rows = ""
        if rubric_data:
            for rubric_name, score in rubric_data:
                # Calculate score bar width (assume max score is 5)
                bar_width = int((score / 5) * 100)
                bar_color = "#28a745" if score >= 4 else "#ffc107" if score >= 3 else "#dc3545"

                rubric_rows += f"""
                <tr>
                    <td style="padding: 12px; border-bottom: 1px solid #e9ecef; font-weight: 500;">{rubric_name}</td>
                    <td style="padding: 12px; border-bottom: 1px solid #e9ecef; text-align: center; font-weight: bold; font-size: 18px;">{score}</td>
                    <td style="padding: 12px; border-bottom: 1px solid #e9ecef;">
                        <div style="background-color: #e9ecef; border-radius: 10px; height: 10px; overflow: hidden;">
                            <div style="background-color: {bar_color}; width: {bar_width}%; height: 100%;"></div>
                        </div>
                    </td>
                </tr>
                """
        else:
            rubric_rows = '<tr><td colspan="3" style="padding: 12px; text-align: center; color: #6c757d;">No rubric scores available</td></tr>'

        # Format remarks
        strengths_text = strengths if pd.notna(strengths) else 'No feedback provided'
        gaps_text = gaps if pd.notna(gaps) else 'No feedback provided'
        other_text = other_remarks if pd.notna(other_remarks) else 'No additional remarks'

        html = f"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Lab Grade Report</title>
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8f9fa;">
    <table role="presentation" style="width: 100%; border-collapse: collapse;">
        <tr>
            <td style="padding: 20px 0;">
                <table role="presentation" style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.1);">

                    <!-- Header -->
                    <tr>
                        <td style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 30px; text-align: center;">
                            <h1 style="margin: 0; color: #ffffff; font-size: 28px; font-weight: 600;">Lab Grade Report</h1>
                            <p style="margin: 10px 0 0 0; color: #e9ecef; font-size: 16px;">{lab_title}</p>
                        </td>
                    </tr>

                    <!-- Greeting -->
                    <tr>
                        <td style="padding: 30px 30px 20px 30px;">
                            <p style="margin: 0; font-size: 16px; color: #495057;">Dear <strong>{nsp_name}</strong>,</p>
                            <p style="margin: 10px 0 0 0; font-size: 14px; color: #6c757d;">Your lab has been reviewed and graded. Here are your results:</p>
                        </td>
                    </tr>

                    <!-- Status Badge -->
                    <tr>
                        <td style="padding: 0 30px;">
                            <div style="background-color: {status_bg}; border-left: 4px solid {status_color}; padding: 20px; border-radius: 8px; text-align: center;">
                                <div style="font-size: 48px; margin-bottom: 10px;">{status_icon}</div>
                                <div style="font-size: 24px; font-weight: bold; color: {status_color}; margin-bottom: 5px;">{status}</div>
                                <div style="font-size: 32px; font-weight: bold; color: #495057;">{score_percent}%</div>
                                <div style="font-size: 14px; color: #6c757d; margin-top: 5px;">Passing Score: {passing_percent}%</div>
                            </div>
                        </td>
                    </tr>

                    <!-- Grade Details -->
                    <tr>
                        <td style="padding: 20px 30px;">
                            <table style="width: 100%; border-collapse: collapse; background-color: #f8f9fa; border-radius: 8px; overflow: hidden;">
                                <tr>
                                    <td style="padding: 12px 20px; border-bottom: 1px solid #dee2e6;">
                                        <strong style="color: #495057;">Attempt:</strong>
                                    </td>
                                    <td style="padding: 12px 20px; border-bottom: 1px solid #dee2e6; text-align: right; color: #6c757d;">
                                        {attempt}
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 12px 20px; border-bottom: 1px solid #dee2e6;">
                                        <strong style="color: #495057;">Re-do Required:</strong>
                                    </td>
                                    <td style="padding: 12px 20px; border-bottom: 1px solid #dee2e6; text-align: right; color: #6c757d;">
                                        {redo_lab}
                                    </td>
                                </tr>
                                <tr>
                                    <td style="padding: 12px 20px;">
                                        <strong style="color: #495057;">Plagiarism Check:</strong>
                                    </td>
                                    <td style="padding: 12px 20px; text-align: right; color: #6c757d;">
                                        {plagiarism}
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- Rubric Scores -->
                    <tr>
                        <td style="padding: 20px 30px;">
                            <h2 style="margin: 0 0 15px 0; color: #495057; font-size: 20px; border-bottom: 2px solid #667eea; padding-bottom: 10px;">📊 Rubric Scores</h2>
                            <table style="width: 100%; border-collapse: collapse; background-color: #ffffff; border: 1px solid #e9ecef; border-radius: 8px; overflow: hidden;">
                                <thead>
                                    <tr style="background-color: #f8f9fa;">
                                        <th style="padding: 12px; text-align: left; color: #495057; font-weight: 600; border-bottom: 2px solid #dee2e6;">Criteria</th>
                                        <th style="padding: 12px; text-align: center; color: #495057; font-weight: 600; border-bottom: 2px solid #dee2e6;">Score</th>
                                        <th style="padding: 12px; text-align: left; color: #495057; font-weight: 600; border-bottom: 2px solid #dee2e6; width: 40%;">Progress</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {rubric_rows}
                                </tbody>
                            </table>
                        </td>
                    </tr>

                    <!-- Strengths -->
                    <tr>
                        <td style="padding: 20px 30px;">
                            <h2 style="margin: 0 0 15px 0; color: #495057; font-size: 20px; border-bottom: 2px solid #28a745; padding-bottom: 10px;">✨ Strengths</h2>
                            <div style="background-color: #d4edda; border-left: 4px solid #28a745; padding: 15px; border-radius: 8px; color: #155724; line-height: 1.6;">
                                {strengths_text}
                            </div>
                        </td>
                    </tr>

                    <!-- Areas for Improvement -->
                    <tr>
                        <td style="padding: 20px 30px;">
                            <h2 style="margin: 0 0 15px 0; color: #495057; font-size: 20px; border-bottom: 2px solid #ffc107; padding-bottom: 10px;">📈 Areas for Improvement</h2>
                            <div style="background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; border-radius: 8px; color: #856404; line-height: 1.6;">
                                {gaps_text}
                            </div>
                        </td>
                    </tr>

                    <!-- Additional Remarks -->
                    <tr>
                        <td style="padding: 20px 30px;">
                            <h2 style="margin: 0 0 15px 0; color: #495057; font-size: 20px; border-bottom: 2px solid #17a2b8; padding-bottom: 10px;">💬 Additional Remarks</h2>
                            <div style="background-color: #d1ecf1; border-left: 4px solid #17a2b8; padding: 15px; border-radius: 8px; color: #0c5460; line-height: 1.6;">
                                {other_text}
                            </div>
                        </td>
                    </tr>

                    <!-- Footer -->
                    <tr>
                        <td style="padding: 30px; background-color: #f8f9fa; text-align: center; border-top: 1px solid #dee2e6;">
                            <p style="margin: 0 0 10px 0; color: #6c757d; font-size: 14px;">If you have any questions about your grade, please reach out during office hours.</p>
                            <p style="margin: 0; color: #495057; font-weight: 600;">Best regards,<br>Franz-James Kaba</p>
                        </td>
                    </tr>

                    <!-- Signature -->
                    <tr>
                        <td style="padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); text-align: center;">
                            <p style="margin: 0; color: #ffffff; font-size: 12px;">© Franzy Lab Grading System</p>
                        </td>
                    </tr>

                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"""
        return html

    def preview_emails(self, module_name: str) -> List[Dict]:
        """Preview all emails that will be sent"""

        if self.grading_data is None:
            self.load_grading_data(module_name)

        if self.email_list is None:
            self.load_email_list()

        emails_to_send = []
        skipped_no_email = []
        skipped_incomplete = []

        print(f"\n{'='*80}")
        print(f"EMAIL PREVIEW - {module_name}")
        print(f"{'='*80}\n")

        # Iterate through grading data
        for idx, row in self.grading_data.iterrows():
            # Get NSP name using column name
            nsp_name = row.get('Name of NSP', None)

            # Skip empty rows
            if pd.isna(nsp_name) or str(nsp_name).strip() == '':
                continue

            # Check if grade is complete
            is_complete, incomplete_reason = self.is_grade_complete(row)

            if not is_complete:
                skipped_incomplete.append((str(nsp_name), incomplete_reason))
                continue

            # Check if email exists
            email = self.match_nsp_email(str(nsp_name))

            if email:
                subject, body = self.generate_email_content(row)

                emails_to_send.append({
                    'to': email,
                    'to_name': nsp_name,
                    'subject': subject,
                    'body': body
                })

                print(f"[{len(emails_to_send)}] To: {nsp_name} <{email}>")
                print(f"    Subject: {subject}")
                print(f"    Preview: {body[:150]}...")
                print()
            else:
                skipped_no_email.append(str(nsp_name))

        print(f"{'='*80}")
        print(f"SUMMARY")
        print(f"{'='*80}")
        print(f"Emails to send: {len(emails_to_send)}")

        if skipped_incomplete:
            print(f"\nSkipped (incomplete grades): {len(skipped_incomplete)}")
            for name, reason in skipped_incomplete[:10]:
                print(f"  - {name}: {reason}")
            if len(skipped_incomplete) > 10:
                print(f"  ... and {len(skipped_incomplete) - 10} more")

        if skipped_no_email:
            print(f"\nSkipped (no email address): {len(skipped_no_email)}")
            print(f"  Students: {', '.join(skipped_no_email[:5])}")
            if len(skipped_no_email) > 5:
                print(f"  ... and {len(skipped_no_email) - 5} more")

        print(f"{'='*80}\n")

        return emails_to_send

    def send_emails(self, emails_to_send: List[Dict], smtp_server: str, smtp_port: int,
                   sender_email: str, sender_password: str):
        """Send emails via SMTP"""

        print(f"\nConnecting to {smtp_server}:{smtp_port}...")

        try:
            # Connect to SMTP server
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(sender_email, sender_password)

            print("Connected successfully!\n")

            sent_count = 0
            failed = []

            for email_data in emails_to_send:
                try:
                    # Create message
                    msg = MIMEMultipart('alternative')
                    msg['From'] = sender_email
                    msg['To'] = email_data['to']
                    msg['Subject'] = email_data['subject']

                    # Attach HTML content
                    msg.attach(MIMEText(email_data['body'], 'html', 'utf-8'))

                    # Send email
                    server.send_message(msg)
                    sent_count += 1
                    print(f"[OK] Sent to {email_data['to_name']} ({email_data['to']})")

                except Exception as e:
                    failed.append((email_data['to_name'], str(e)))
                    print(f"[FAILED] Failed to send to {email_data['to_name']}: {e}")

            server.quit()

            print(f"\n{'='*80}")
            print(f"Sending complete!")
            print(f"Sent: {sent_count}/{len(emails_to_send)}")
            if failed:
                print(f"Failed: {len(failed)}")
                for name, error in failed:
                    print(f"  - {name}: {error}")
            print(f"{'='*80}\n")

        except Exception as e:
            print(f"[ERROR] Failed to connect to SMTP server: {e}")


def main():
    """Main CLI interface"""

    print("""
╔═══════════════════════════════════════════════════════════════╗
║            LAB REPORT SENDER - NSP Grading System             ║
╚═══════════════════════════════════════════════════════════════╝
""")

    # File paths
    grading_file = "C:/Users/Franz-JamesWefagiKab/OneDrive - AmaliTech gGmbH/Lab Materials/BE Lab Grading Sheet.xlsx"
    email_list_file = "C:/Users/Franz-JamesWefagiKab/OneDrive - AmaliTech gGmbH/Lab Materials/Quality Assurance - Copy.xlsx"

    # Initialize sender and config
    sender = LabReportSender(grading_file, email_list_file)
    config = EmailConfig()

    # Show options menu
    print("\nOptions:")
    print("1. Send lab reports")
    print("2. Configure email credentials")
    print("3. Clear saved credentials")

    option = input("\nSelect option (1-3, or press Enter for option 1): ").strip() or "1"

    if option == "2":
        # Reconfigure credentials
        smtp_server, smtp_port, sender_email, sender_password = get_smtp_config(config, force_save=True)
        if smtp_server:
            print("\nCredentials updated successfully!")
        return
    elif option == "3":
        # Clear credentials
        config.clear_config()
        return
    elif option != "1":
        print("Invalid option!")
        return

    # Get module name
    print("\nAvailable modules:")
    print("1. Module-1")
    print("2. Module-2")
    print("3. Module-3")

    module_choice = input("\nSelect module number (1-3): ").strip()
    module_map = {'1': 'Module-1', '2': 'Module-2', '3': 'Module-3'}

    module_name = module_map.get(module_choice)

    if not module_name:
        print("Invalid module selection!")
        return

    # Preview emails
    emails_to_send = sender.preview_emails(module_name)

    if not emails_to_send:
        print("No emails to send!")
        return

    # Confirm sending
    confirm = input("\nDo you want to send these emails? (yes/no): ").strip().lower()

    if confirm != 'yes':
        print("Sending cancelled.")
        return

    # Email configuration - automatically use saved credentials if they exist
    if config.has_config():
        smtp_server = config.get_smtp_server()
        smtp_port = config.get_smtp_port()
        sender_email = config.get_email()
        sender_password = config.get_password()
        print(f"\nUsing saved credentials: {sender_email}")
    else:
        print("\n[No saved credentials found - First time setup]")
        smtp_server, smtp_port, sender_email, sender_password = get_smtp_config(config, force_save=True)

        if not smtp_server:
            print("Configuration cancelled.")
            return

    # Send emails
    sender.send_emails(emails_to_send, smtp_server, smtp_port, sender_email, sender_password)

    print("\nDone!")


def get_smtp_config(config, force_save=False):
    """Get SMTP configuration from user"""
    print("\nEmail Configuration:")
    print("1. Gmail (smtp.gmail.com)")
    print("2. Outlook (smtp-mail.outlook.com)")

    smtp_choice = input("Select email provider (1-2): ").strip()

    if smtp_choice == '1':
        smtp_server = 'smtp.gmail.com'
        smtp_port = 587
    elif smtp_choice == '2':
        smtp_server = 'smtp-mail.outlook.com'
        smtp_port = 587
    else:
        print("Invalid choice!")
        return None, None, None, None

    sender_email = input("Your email address: ").strip()
    sender_password = getpass.getpass("Your email password (or app password): ")

    # Save credentials automatically if force_save, otherwise ask
    if force_save:
        config.save_config(smtp_server, smtp_port, sender_email, sender_password)
        print("Credentials saved!")
    else:
        save = input("\nSave these credentials for future use? (yes/no): ").strip().lower()
        if save == 'yes':
            config.save_config(smtp_server, smtp_port, sender_email, sender_password)

    return smtp_server, smtp_port, sender_email, sender_password


if __name__ == "__main__":
    main()
