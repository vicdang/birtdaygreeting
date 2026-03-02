# Birthday Card System 🎂

**Automated birthday card generation and sending system for DC34 PUB team**

## Overview

This system automatically generates personalized birthday cards with employee photos and sends them via Outlook email. It's designed to run daily at 7:00 AM and sends birthday wishes to team members.

**Status**: ✅ Production Ready

## Features

✅ **Automated Daily Execution** - Runs at 7:00 AM every day via Windows Task Scheduler  
✅ **Personalized Cards** - Generates cards with employee photos from HRM system  
✅ **Email Integration** - Sends via Windows Outlook (no SMTP credentials needed)  
✅ **Dual Sender Support** - Toggle between primary and alternative email addresses  
✅ **CC Recipients** - Automatically CC team leads and managers  
✅ **Rich HTML Email** - Embeds card images directly in email  
✅ **Comprehensive Logging** - All operations logged for audit trail  
✅ **Easy Management Menu** - Simple click-to-run batch file interface  

## Quick Start

### Daily Use

1. **Double-click** `birthday_menu.bat` in File Explorer
2. **Press 1** to send today's cards
3. System will automatically generate and send birthday emails

### Automatic Daily Execution

The system runs automatically at **7:00 AM** every day via Windows Task Scheduler. No manual action needed.

### Test Tomorrow's Cards

1. **Double-click** `birthday_menu.bat`
2. **Press 2** to test tomorrow's workflow
3. Generates and sends cards for tomorrow's birthdays

## Installation

### Prerequisites

- Windows 7 or later
- Python 3.10+
- Outlook configured with email account
- Network connectivity for HRM photo download

### Setup

1. Extract the project to: `d:\PlayGround\birthdaycard\birtdaygreeting`

2. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Configure the scheduler (as Administrator):
   ```powershell
   schtasks /create /tn "Birthday Card Sender" /tr "D:\Program Files\Python\Python314\python.exe d:\PlayGround\birthdaycard\birtdaygreeting\send_via_outlook.py" /sc daily /st 07:00 /rl highest /f
   ```

4. (Optional) Verify installation:
   ```bash
   python send_via_outlook.py
   ```

## Configuration

### Email Settings (`.env`)

```ini
SMTP_HOST=smtp.trn.com.vn
SMTP_PORT=465
SMTP_USER=abc@trn.com.vn
SENDER_NAME=DC34 PUB Birthday Greetings
SENDER_EMAIL=abc@trn.com.vn (primary sender)
ALT_SENDER_EMAIL=dc34_pub@trn.com.vn (alternative sender)
USE_ALT_SENDER=false (set to true to use alternative email)
```

### Email Configuration (`email_config.json`)

- **recipient**: Email addresses to send to
- **subject**: Email subject line
- **sender_name**: Display name for sender
- **sender_email**: Primary email address
- **alt_sender_email**: Alternative email address
- **use_alt_sender**: Toggle between emails (true/false)
- **additional_cc**: CC recipients (team leads, managers)
- **template**: Birthday message templates
- **styling**: Email styling colors and fonts

### Employee Data (`data/members.xlsx`)

Contains:
- 179 employees
- Member ID, Name, Birthday, Email, HRM photo URL
- Updated quarterly

## Menu Options

**Double-click `birthday_menu.bat` to access:**

| Option | Function | Use Case |
|--------|----------|----------|
| **1** | Send Today's Cards | Manual execution for today |
| **2** | Test Tomorrow's Workflow | Test cards for tomorrow's date |
| **3** | Generate Cards Only | Create cards without sending email |
| **4** | Check Scheduler Status | Verify task is running correctly |
| **5** | Change Sender Email | Switch between primary/alternative sender |
| **6** | PowerShell Commands | Reference for advanced management |
| **7** | View Configuration | Display all current settings |
| **8** | Exit | Close the menu |

## Email Sender Configuration

### Using Primary Email

Primary sender is `abc@trn.com.vn` (default)

- **Edit `.env`**: `USE_ALT_SENDER=false`
- **Edit `email_config.json`**: `"use_alt_sender": false`

### Using Alternative Email

Alternative sender is `dc34_pub@trn.com.vn`

- **Edit `.env`**: `USE_ALT_SENDER=true`
- **Edit `email_config.json`**: `"use_alt_sender": true`

Both `.env` and `email_config.json` must be synchronized.

## File Structure

```
birthday-bot/
├── birthday_bot/                 # Core Python package
│   ├── __init__.py
│   ├── config.py               # Config loading
│   ├── roster.py               # Excel roster management
│   ├── renderer.py             # Card generation
│   ├── mailer.py               # Email sending
│   └── utils.py                # Utilities
├── data/
│   └── members.xlsx            # Employee roster (179 members)
├── out/
│   └── birthday_cards_today/   # Generated card images
├── .env                        # Environment variables
├── .gitignore                  # Git ignore rules
├── birthday_menu.bat           # Main menu (double-click to run)
├── birthday_menu_gui.ps1       # Alternative PowerShell GUI menu
├── card_config.json            # Card design configuration
├── email_config.json           # Email settings and templates
├── greeting_messages.json      # 30+ birthday greeting templates
├── generate_today_birthdays.py # Generate-only script
├── send_via_outlook.py         # Daily sender script
├── test_tomorrow_workflow.py   # Tomorrow test script
├── requirements.txt            # Python dependencies
└── README.md                   # This file
```

## Daily Workflow

**Automatic (7:00 AM)**:
1. Check for employees with today's birthday
2. Download photos from HRM system
3. Generate personalized birthday cards
4. Compose HTML email with embedded cards
5. Send via Outlook to recipients with CC
6. Log all operations

**Manual**:
1. Double-click `birthday_menu.bat`
2. Press 1 (Send Today's Cards)
3. Same workflow as automatic

## Logs

All operations logged to: `birthday_bot.log`

View recent logs:
```powershell
Get-Content birthday_bot.log -Tail 20
```

## Task Scheduler Management

### Check Status
```powershell
Get-ScheduledTask -TaskName "Birthday Card Sender"
```

### Run Manually
```powershell
schtasks /run /tn "Birthday Card Sender"
```

### Disable Task
```powershell
schtasks /change /tn "Birthday Card Sender" /disable
```

### Enable Task
```powershell
schtasks /change /tn "Birthday Card Sender" /enable
```

### Delete Task
```powershell
schtasks /delete /tn "Birthday Card Sender" /f
```

## Troubleshooting

### No emails sent on a birthday

**Possible causes**:
1. No employees have birthday on that date → Check `data/members.xlsx`
2. Outlook not configured → Configure in Windows Settings
3. Email disabled in config → Check `email_config.json`: `"enabled": true`
4. Task not running → Check Task Scheduler status

**Solution**:
```bash
python send_via_outlook.py  # Test manually
```

### Cards not generating

**Possible causes**:
1. Python not installed → Install Python 3.10+
2. Missing dependencies → `pip install -r requirements.txt`
3. HRM photo URL unavailable → Check network connectivity

### Alternative sender not working

**Possible causes**:
1. Configuration mismatch → Sync `.env` and `email_config.json`
2. Outlook account not configured → Add account in Windows Outlook
3. Typo in email address → Verify `dc34_pub@trn.com.vn`

## Technical Details

**Technologies**:
- Python 3.10+
- pandas - Excel roster management
- Pillow - Image processing
- pywin32 - Windows Outlook COM interface
- JSON - Configuration management

**Python Execution**:
- Interpreter: `D:\Program Files\Python\Python314\python.exe`
- Working Directory: `d:\PlayGround\birthdaycard\birtdaygreeting`

**Email System**:
- Uses Windows Outlook COM interface (no SMTP login needed)
- Automatically uses configured Outlook account
- Supports HTML body with embedded images
- CC recipients from configuration

**Image Processing**:
- Downloads photos from HRM system
- Circular mask with soft shadow
- Embedded directly in email (no attachments)

## Maintenance

### Daily
- System automatically runs at 7:00 AM
- Check logs: `birthday_bot.log`

### Monthly
- Review sent emails and logs
- Update employee roster if needed

### Quarterly  
- Update `data/members.xlsx` with new employees
- Refresh employee photos in HRM

## Support

For issues or questions:

1. Check `birthday_bot.log` for detailed error messages
2. Run manual test: `python send_via_outlook.py`
3. Review `SCHEDULER_SETUP.md` for advanced tasks
4. Verify configuration: Menu Option 7

## Version

- **Version**: 1.0.0
- **Status**: Production Ready
- **Last Updated**: March 2, 2026
- **Deployed**: March 2, 2026

## References

- 📊 Employee Roster: `data/members.xlsx` (179 members)
- 🎨 Card Design: `card_config.json`
- 📧 Email Config: `email_config.json`
- 💬 Messages: `greeting_messages.json`
- 📅 Scheduler: Windows Task Scheduler - "Birthday Card Sender"
- ⏰ Run Time: Daily at 7:00 AM
- 👤 Primary Sender: abc@trn.com.vn
- 📤 Alternative Sender: dc34_pub@trn.com.vn

---

**Ready for daily use!** 🎉
