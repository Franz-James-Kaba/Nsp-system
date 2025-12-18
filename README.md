# NSP Lab Report Sender System

The **NSP Lab Report Sender** is an automated grading and feedback distribution system designed to streamline the process of sending detailed lab reports to National Service Personnel (NSPs). It bridges the gap between Excel-based grading sheets and personalized email feedback, ensuring that every student receives a comprehensive breakdown of their performance.

## 📋 Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [System Architecture](#system-architecture)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage Guide](#usage-guide)
- [Technical Details](#technical-details)
    - [Student Matching Logic](#student-matching-logic)
    - [Grading Logic](#grading-logic)
    - [Security](#security)
- [Troubleshooting](#troubleshooting)
- [Customization](#customization)

---

## 🚀 Overview

Grading and providing feedback for large cohorts of students can be tedious and error-prone when done manually. This system automates the last mile of the grading process: **Communication**.

It reads grading data from a standardized Excel sheet, matches student names against an official email list, and generates individualized HTML emails containing:
- **Pass/Fail Status**
- **Rubric-level scores** (visualized with progress bars)
- **Specific feedback** (Strengths and Areas for Improvement)
- **Administrative details** (Attempt number, Plagiarism checks)

---

## ✨ Key Features

- **Automated Data Extraction**: Parses complex Excel grading sheets to extract scores, remarks, and metadata.
- **Smart Student Matching**: Uses a 3-tier matching algorithm (Exact, Partial, Reverse-Partial) to link grading names to email addresses, handling typos and name variations.
- **Rich HTML Emails**: Sends professionally formatted emails with color-coded status indicators (Green for Pass, Red for Re-do).
- **Interactive Preview**: Allows the instructor to preview every single email (subject and body snippet) before sending.
- **Secure Credential Management**: Stores SMTP credentials locally using base64 encoding to prevent repeated logic.
- **Fail-Safe Mechanisms**:
    - Skips students with incomplete grades (missing rubric scores or remarks).
    - Skips students without found email addresses.
    - Provides a summary report of successful sends and failures.

---

## 🏗️ System Architecture

The project consists of the following core components:

| File | Description |
|------|-------------|
| `lab_report_sender.py` | The **Main Application**. Handles UI, logic, parsing, and email dispatch. |
| `email_config.py` | **Configuration Manager**. Handles loading/saving of SMTP credentials. |
| `.email_config.json` | **Updates Automatically**. Stores user credentials locally (hidden file). |
| `test_single_email.py` | **Testing Utility**. Allows sending a test email to a specific student for debugging. |
| `send_test_email.py` | **Basic Test**. Sends a generic test email to verify SMTP connectivity. |

---

## 📝 Prerequisites

- **OS**: Windows (tested), macOS, or Linux.
- **Python**: Version 3.7 or higher.
- **Excel Files**:
    - **Grading Sheet**: `BE Lab Grading Sheet.xlsx` (Must follow the specific row structure).
    - **Email List**: `Quality Assurance - Copy.xlsx` (Must contain 'Full Name' and 'AmaliTech Email').

---

## 📦 Installation

1. **Clone the repository** (if applicable) or navigate to the project folder:
   ```bash
   cd "C:\Users\Franz-JamesWefagiKab\IdeaProjects\Nsp system"
   ```

2. **Install dependencies**:
   The system relies on `pandas` for data manipulation and `openpyxl` for Excel reading.
   ```bash
   pip install pandas openpyxl
   ```

---

## ⚙️ Configuration

### 1. File Paths
By default, the script looks for Excel files in specific absolute paths. To change these, modify the `main()` function in `lab_report_sender.py` (lines 503-504):

```python
grading_file = "C:/path/to/your/BE Lab Grading Sheet.xlsx"
email_list_file = "C:/path/to/your/Quality Assurance - Copy.xlsx"
```

### 2. Email Provider
The system supports:
- **Gmail** (`smtp.gmail.com`) - Requires an **App Password**.
- **Outlook** (`smtp-mail.outlook.com`) - Uses standard password (or App Password if 2FA is on).

---

## 📖 Usage Guide

### Step 1: Run the Application
Open your terminal and run:
```bash
python lab_report_sender.py
```

### Step 2: Main Menu
You will see the main interface:
```
1. Send lab reports
2. Configure email credentials
3. Clear saved credentials
```
Press **Enter** to select option 1.

### Step 3: Select Module
Choose the module you are grading (corresponding to the sheet name in Excel):
```
Available modules:
1. Module-1
2. Module-2
3. Module-3
```

### Step 4: Review Preview
The system will analyze the Excel sheet and print a preview:
- **[1] To: John Doe <john@example.com>**
- **Subject**: Lab Grade: Intro to Python - PASSED
- **Preview**: Dear John Doe...

It will also list **Skipped Students** (those with incomplete grades or no email found).

### Step 5: Confirm & Send
If the preview looks correct, type `yes` to proceed.
- If it's your first time, you will be prompted to enter your email and password.
- Credentials will be saved for future runs.

---

## 🔍 Technical Details

### Student Matching Logic
Finding the correct email for a student name is critical. The `match_nsp_email` method uses a robust strategy:
1.  **Exact Match**: Looks for an identical string (case-insensitive).
    -   *Example*: "John Doe" matches "John Doe".
2.  **Partial Match**: Checks if the grading name is contained within the email list name.
    -   *Example*: "John" in grading sheet will match "John Doe" in email list.
3.  **Reverse Partial / Component Match**: Splits names into parts and checks for significant overlap.
    -   *Example*: "Bernice Mawuena" (Grading) matches "Bernice Adime Mawuena" (Email List) because at least 2 parts match.

### Grading Logic
- **Passing Threshold**: 80% (0.8).
- **Incomplete Grade Detection**: A record is considered "Incomplete" and skipped if:
    -   Total Score is 0 or missing.
    -   No rubric scores are present.
    -   No textual feedback (remarks) is present.

### Security
- **Credentials**: Stored in `.email_config.json`.
- **Encryption**: Passwords are **base64 encoded** to prevent casual shoulder-surfing, but this is NOT strong encryption.
- **Recommendation**: Never commit `.email_config.json` to version control.

---

## 🔧 Troubleshooting

### "Permission Denied" Error
**Cause**: The Excel file is currently open in Microsoft Excel.
**Fix**: Close all Excel windows properly and try again.

### "Authentication Failed"
**Cause**: Incorrect password or security blocking the login.
**Fix**:
- **Gmail**: You MUST use an **App Password** (Google Account -> Security -> 2-Step Verification -> App Passwords). Your normal Gmail password will not work.
- **Outlook**: Ensure strict spam filters aren't blocking the script.

### "No emails to send"
**Cause**: The script found no rows that met the "Complete Grade" criteria, or no matched emails.
**Fix**:
- Ensure the "Total Score" column is filled.
- Ensure at least one rubric column has a value.
- Check "Remarks: Strengths" or "Remarks: Gaps" columns.
- Verify the student name spelling matches the email list.

---

## 🎨 Customization

To modify the email template, edit the `_generate_html_email` method in `lab_report_sender.py`.
- **Colors**: Defined in `status_color`, `status_bg`.
- **Header**: Change the `<h1 ...>Lab Grade Report</h1>` section.
- **Signature**: Update the footer text in the HTML string.

---

<p align="center">
  © 2025 Franzy Lab Grading System
</p>
