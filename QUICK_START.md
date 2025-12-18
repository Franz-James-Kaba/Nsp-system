# Lab Report Sender - Quick Start Guide

## One-Time Setup (First Run Only)

```bash
cd "C:\Users\Franz-JamesWefagiKab\IdeaProjects\Nsp system"
python lab_report_sender.py
```

**You will be prompted:**
1. Select option: `1` (Send lab reports)
2. Select module: `1`, `2`, or `3`
3. Review preview
4. Confirm send: `yes`
5. **[FIRST TIME ONLY]** Enter email credentials:
   - Email provider: `1` (Gmail) or `2` (Outlook)
   - Your email address
   - Your app password
   - ✅ **Credentials are automatically saved!**

## Every Subsequent Run

```bash
python lab_report_sender.py
```

**Simple workflow:**
1. Press Enter (or type `1` for Send lab reports)
2. Select module: `1`, `2`, or `3`
3. Review preview
4. Confirm send: `yes`
5. ✅ **Automatically uses saved credentials - NO PASSWORD ENTRY!**

## Managing Credentials

### Update Email Password
```bash
python lab_report_sender.py
```
- Select option: `2` (Configure email credentials)
- Enter new credentials
- Done!

### Clear Saved Credentials
```bash
python lab_report_sender.py
```
- Select option: `3` (Clear saved credentials)

Or manually delete:
```bash
rm .email_config.json
```

## Testing with Single Email

```bash
python test_single_email.py
```

**First time:**
- Prompts for credentials (automatically saved)

**Every time after:**
- Automatically uses saved credentials!

---

## Gmail App Password Setup

**Required for Gmail users:**

1. Go to: https://myaccount.google.com/apppasswords
2. Sign in if needed
3. Enable 2-Step Verification if not already enabled
4. Select app: "Mail"
5. Select device: "Windows Computer"
6. Click "Generate"
7. Copy the 16-character password
8. Use this password (remove spaces) in the script

---

## How It Works

**First Run:**
```
╔═══════════════════════════════════════════════════════════════╗
║            LAB REPORT SENDER - NSP Grading System             ║
╚═══════════════════════════════════════════════════════════════╝

Options:
1. Send lab reports
2. Configure email credentials
3. Clear saved credentials

Select option (1-3, or press Enter for option 1): [ENTER]

Available modules:
1. Module-1
2. Module-2
3. Module-3

Select module number (1-3): 2

[Preview shows students with complete grades]

Do you want to send these emails? (yes/no): yes

[No saved credentials found - First time setup]

Email Configuration:
1. Gmail (smtp.gmail.com)
2. Outlook (smtp-mail.outlook.com)

Select email provider (1-2): 1
Your email address: your@gmail.com
Your email password (or app password): ****************
Credentials saved!

Using saved credentials: your@gmail.com

[Sending emails...]
Done!
```

**Every Subsequent Run:**
```
╔═══════════════════════════════════════════════════════════════╗
║            LAB REPORT SENDER - NSP Grading System             ║
╚═══════════════════════════════════════════════════════════════╝

Options:
1. Send lab reports
2. Configure email credentials
3. Clear saved credentials

Select option (1-3, or press Enter for option 1): [ENTER]

Available modules:
1. Module-1
2. Module-2
3. Module-3

Select module number (1-3): 2

[Preview shows students with complete grades]

Do you want to send these emails? (yes/no): yes

Using saved credentials: your@gmail.com

[Sending emails...]
Done!
```

**That's it! No more password entry!** 🎉

---

## What Gets Sent?

**The system automatically:**
- ✅ Only sends to students with complete grades (scores + remarks)
- ✅ Skips students with incomplete grades
- ✅ Skips students without email addresses
- ✅ Shows clear summary of who will receive emails

**Each email includes:**
- Status (PASSED/NEEDS RE-DO)
- Total score and percentage
- All rubric scores (adapts per module)
- Strengths feedback
- Areas for improvement
- Additional remarks

---

## Troubleshooting

**"Authentication failed"**
- Gmail: Use App Password, not regular password
- Outlook: Try App Password if regular password fails

**"No emails to send"**
- Students need complete grades (scores AND remarks)
- Students need email addresses in the list

**"Permission denied" on Excel file**
- Close the Excel file before running script

**Want to use different email?**
```bash
python lab_report_sender.py
```
- Select option: `2`
- Enter new credentials

---

## Security Note

Your credentials are saved in `.email_config.json` in the project directory.
- Password is base64-encoded (not plaintext)
- File is local only
- Don't share this file or commit to git
- Use app-specific passwords, not main email password

---

**Questions? The system is now fully automated - just run and go!** 🚀
