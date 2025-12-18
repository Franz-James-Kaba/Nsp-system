# Lab Report Sender System

Automated system for sending graded lab reports to NSPs via email.

## Features

- Reads grading data from Excel sheets (BE Lab Grading Sheet.xlsx)
- Matches student names with email addresses
- Generates personalized emails with:
  - Complete rubric scores for all grading categories
  - Total score and pass/fail status
  - Attempt number
  - Re-do requirement status
  - Plagiarism check result
  - Strengths, areas for improvement, and additional remarks
- Preview emails before sending
- Supports Gmail and Outlook SMTP
- Automatically skips students without email addresses

## Requirements

- Python 3.7+
- pandas
- openpyxl

## Installation

Dependencies are already installed:
```bash
pip install pandas openpyxl
```

## File Structure

```
lab_report_sender.py     # Main script
README_LAB_REPORT_SENDER.md  # This file
```

## Usage

### Step 1: Prepare Your Files

Ensure these files are in place:
- **Grading Sheet**: `C:/Users/Franz-JamesWefagiKab/OneDrive - AmaliTech gGmbH/Lab Materials/BE Lab Grading Sheet.xlsx`
- **Email List**: `C:/Users/Franz-JamesWefagiKab/OneDrive - AmaliTech gGmbH/Lab Materials/Quality Assurance - Copy.xlsx`

### Step 2: Run the Script

```bash
cd "C:\Users\Franz-JamesWefagiKab\IdeaProjects\Nsp system"
python lab_report_sender.py
```

### Step 3: Follow the Prompts

1. **Select Module**: Choose which module sheet to send reports for (1-3)
   - Module-1
   - Module-2
   - Module-3

2. **Preview Emails**: The system will show you:
   - How many emails will be sent
   - Preview of each email (first 150 characters)
   - List of students who will be skipped (no email address)

3. **Confirm Sending**: Type `yes` to proceed or `no` to cancel

4. **Choose Email Provider**:
   - Option 1: Gmail (smtp.gmail.com)
   - Option 2: Outlook (smtp-mail.outlook.com)

5. **Enter Credentials**:
   - Your email address
   - Your password or app password

### Step 4: Monitor Sending

The system will:
- Show progress as each email is sent
- Mark successful sends with [✓]
- Mark failed sends with [✗]
- Provide a final summary

## Email Configuration

### For Gmail

1. Use your Gmail address
2. **Important**: Use an [App Password](https://support.google.com/accounts/answer/185833), not your regular password
   - Go to Google Account → Security → 2-Step Verification → App Passwords
   - Generate a new app password for "Mail"
   - Use that 16-character password in the script

### For Outlook

1. Use your Outlook/AmaliTech email address
2. Use your regular email password
3. If you have 2FA enabled, you may need an app password

## Current Status

**Matched Students**: 9 students have email addresses and will receive reports
**Skipped Students**: 25 students in the grading sheet don't have emails in the current list

### Students Who Will Receive Emails:
1. Francis Roland Bissah
2. Lennox Owusu Afriyie
3. Nana Kwaku Sarpong
4. Harriet Effah
5. Tob Adoba
6. Samuel Oduro Duah Boakye (matched as "Samuel Oduro")
7. Fred Pekyi
8. Zakaria Osman
9. Patrick Koduah Assiamah

## Sample Email Output

```
Subject: Lab Grade: Grade Management System - PASSED

Dear Francis Roland Bissah,

Your lab grade for "Grade Management System" has been reviewed and graded.

--- GRADE SUMMARY ---
Status: PASSED
Total Score: 0.86 (Passing Score: 0.8)
Attempt: 1st
Re-do Required: No
Plagiarism Check: Passed

--- RUBRIC SCORES ---
OOP Principles: 4
Functionality: 5
Class Design: 5
DSA: 4
Code Quality: 3
Documentation: 4

--- STRENGTHS ---
[Your detailed feedback on strengths...]

--- AREAS FOR IMPROVEMENT ---
[Your detailed feedback on gaps...]

--- ADDITIONAL REMARKS ---
[Your additional comments...]

---

If you have any questions about your grade, please reach out during office hours.

Best regards,
Your Instructor
```

## Troubleshooting

### "No module named 'pandas'"
Run: `pip install pandas openpyxl`

### "Permission denied" on Excel files
Close the Excel files before running the script

### "No emails found"
- Check that student names in the grading sheet match names in the email list
- The system only sends to students who exist in BOTH files

### SMTP Authentication Failed
- Gmail: Use an app password, not your regular password
- Outlook: Check your email and password are correct
- Check if your email provider requires 2FA or app-specific passwords

### Some students not receiving emails
Only students who appear in BOTH the grading sheet AND the email list will receive emails. This is by design to prevent sending to incorrect addresses.

## Workflow Options

### Option A: Send All Graded Students in a Module
1. Finish grading all students in Module-1
2. Run the script and select Module-1
3. Review preview
4. Confirm and send

### Option C (Built-in): Preview Before Sending
The script automatically shows you a complete preview of all emails before asking for confirmation.

## Customization

### Change File Paths
Edit lines 260-261 in `lab_report_sender.py`:
```python
grading_file = "your/path/to/grading/sheet.xlsx"
email_list_file = "your/path/to/email/list.xlsx"
```

### Change Passing Score
Edit line 86 in `lab_report_sender.py`:
```python
passing_score = 0.8  # Change to your threshold
```

### Customize Email Template
Edit the `generate_email_content()` method starting at line 70

## Support

For issues or questions, check:
1. README troubleshooting section
2. Verify file paths are correct
3. Ensure Excel files are closed
4. Check email credentials

## Version

Current Version: 1.0
Last Updated: 2025-12-17
