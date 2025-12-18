"""
Test Single Email Script
Allows testing the email sending functionality with a single student record
and a custom destination email address.
"""

from lab_report_sender import LabReportSender
from email_config import EmailConfig
import getpass
import pandas as pd
import sys

class TestLabReportSender(LabReportSender):
    def run_test(self):
        print("\n=== TEST SINGLE EMAIL MODE ===\n")

        # 1. Select Module
        print("Available modules:")
        print("1. Module-1")
        print("2. Module-2")
        print("3. Module-3")
        
        module_choice = input("\nSelect module number (1-3): ").strip()
        module_map = {'1': 'Module-1', '2': 'Module-2', '3': 'Module-3'}
        module_name = module_map.get(module_choice)
        
        if not module_name:
            print("Invalid module selection!")
            return

        # Load data
        self.load_grading_data(module_name)
        if self.email_list is None:
            self.load_email_list()

        # 2. Select Student
        print(f"\nStudents in {module_name}:")
        student_records = []
        
        # Filter for valid students (those with names)
        for idx, row in self.grading_data.iterrows():
            nsp_name = row.get('Name of NSP')
            if pd.notna(nsp_name) and str(nsp_name).strip() != '':
                student_records.append((idx, row))

        print(f"\nFound {len(student_records)} students.")
        search_query = input("Enter student name to search (or press Enter to list all): ").strip().lower()

        filtered_records = []
        if search_query:
            for i, (idx, row) in enumerate(student_records):
                nsp_name = str(row.get('Name of NSP')).strip()
                if search_query in nsp_name.lower():
                    filtered_records.append((i, idx, row)) # Keep original 'i' (index in full list) for consistency if needed, or just iterate
        else:
            # All records
            filtered_records = [(i, idx, row) for i, (idx, row) in enumerate(student_records)]

        if not filtered_records:
            print("No students found matching that name.")
            return

        print("\nSelect student:")
        for i, (original_list_idx, idx, row) in enumerate(filtered_records):
            nsp_name = row.get('Name of NSP')
            status = "PASSED" if row.get('Total Score', 0) >= 0.8 else "NEEDS RE-DO"
            print(f"{i+1}. {nsp_name} ({status})")

        try:
            selection = int(input("\nSelect number from above list: ").strip())
            if selection < 1 or selection > len(filtered_records):
                print("Invalid selection!")
                return
            
            _, selected_idx, selected_row = filtered_records[selection - 1]
            nsp_name = selected_row.get('Name of NSP')

        except ValueError:
            print("Invalid input!")
            return

        # 3. Enter Test Email
        test_email = input(f"\nEnter test email address to send report for '{nsp_name}': ").strip()
        if not test_email:
            print("Email is required!")
            return

        # 4. Generate and Preview
        subject, body = self.generate_email_content(selected_row)
        
        # Modify subject for testing
        subject = f"{subject} [TEST]"
        
        print(f"\n{'='*80}")
        print(f"PREVIEW FOR: {nsp_name}")
        print(f"SENDING TO: {test_email}")
        print(f"{'='*80}")
        print(f"Subject: {subject}\n")
        print(body)
        print(f"{'='*80}\n")

        # 5. Confirm
        confirm = input("Do you want to send this TEST email? (yes/no): ").strip().lower()
        if confirm != 'yes':
            print("Test cancelled.")
            return

        # 6. Email Configuration - automatically use saved credentials if they exist
        config = EmailConfig()

        if config.has_config():
            smtp_server = config.get_smtp_server()
            smtp_port = config.get_smtp_port()
            sender_email = config.get_email()
            sender_password = config.get_password()
            print(f"\nUsing saved credentials: {sender_email}")
        else:
            print("\n[No saved credentials found - First time setup]")
            smtp_server, smtp_port, sender_email, sender_password = self._get_smtp_config(config, force_save=True)

            if not smtp_server:
                print("Configuration cancelled.")
                return

        # Prepare single email payload
        email_data = {
            'to': test_email,
            'to_name': nsp_name,
            'subject': subject,
            'body': body
        }

        # Send
        self.send_emails([email_data], smtp_server, smtp_port, sender_email, sender_password)

    def _get_smtp_config(self, config, force_save=False):
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
    # File paths (same as main script)
    grading_file = "C:/Users/Franz-JamesWefagiKab/OneDrive - AmaliTech gGmbH/Lab Materials/BE Lab Grading Sheet.xlsx"
    email_list_file = "C:/Users/Franz-JamesWefagiKab/OneDrive - AmaliTech gGmbH/Lab Materials/Quality Assurance - Copy.xlsx"

    tester = TestLabReportSender(grading_file, email_list_file)
    tester.run_test()
