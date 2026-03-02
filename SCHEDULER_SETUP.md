# Birthday Card Scheduler Setup Guide

## Overview
This guide shows you how to set up automatic daily birthday card generation and sending at 7:00 AM using Windows Task Scheduler.

## Prerequisites
- Windows operating system (Windows 7 or later)
- Python 3.10+ installed
- Outlook configured with your email account
- Administrator access to your computer

## Quick Setup

### Step 1: Run the Setup Script

1. Open **PowerShell as Administrator**
   - Press `Win + X` and select "Windows PowerShell (Admin)"
   - Or search "PowerShell" and run as administrator

2. Navigate to the birthday-bot directory:
   ```powershell
   cd d:\PlayGround\birthdaycard\birthday-bot
   ```

3. Run the setup script:
   ```powershell
   .\setup_scheduler.ps1
   ```

4. The script will:
   - Create a scheduled task named "Birthday Card Sender"
   - Set it to run daily at 7:00 AM
   - Configure it with highest privileges
   - Display confirmation with task details

### Step 2: Verify the Task

After setup, verify the task was created:

```powershell
Get-ScheduledTask -TaskName "Birthday Card Sender"
```

You should see output showing the task is enabled.

## Manual Setup (Alternative)

If you prefer to create the task manually using Task Scheduler GUI:

1. Open Task Scheduler
   - Press `Win + R` and type: `taskschd.msc`
   - Press Enter

2. Click "Create Task" on the right panel

3. **General Tab:**
   - Name: `Birthday Card Sender`
   - Description: `Daily automatic birthday card generation and sending`
   - Run with highest privileges: ✓ Check

4. **Triggers Tab:**
   - Click "New..."
   - Begin the task: `On a schedule`
   - Daily
   - Start: Today (or any date)
   - Recur every: 1 day
   - Time: 07:00:00 AM
   - Click "OK"

5. **Actions Tab:**
   - Click "New..."
   - Action: `Start a program`
   - Program/script: `D:\Program Files\Python\Python314\python.exe`
   - Add arguments: `d:\PlayGround\birthdaycard\birthday-bot\send_via_outlook.py`
   - Start in: `d:\PlayGround\birthdaycard\birthday-bot`
   - Click "OK"

6. **Settings Tab:**
   - Allow task to be run on demand: ✓ Check
   - Run task as soon as possible after a scheduled start is missed: ✓ Check
   - If the running task does not end when requested: `Stop the task`
   - Click "OK"

7. Click "OK" to save the task

## Manage the Scheduled Task

### View Task Status
```powershell
Get-ScheduledTask -TaskName "Birthday Card Sender" | Select-Object State, TaskName, Description
```

### Enable/Disable the Task
```powershell
# Enable
Enable-ScheduledTask -TaskName "Birthday Card Sender"

# Disable
Disable-ScheduledTask -TaskName "Birthday Card Sender"
```

### View Last Run Info
```powershell
Get-ScheduledTaskInfo -TaskName "Birthday Card Sender"
```

### Delete the Task
```powershell
Unregister-ScheduledTask -TaskName "Birthday Card Sender" -Confirm:$false
```

### Run Task Manually
```powershell
Start-ScheduledTask -TaskName "Birthday Card Sender"
```

## Change Schedule Time

To change the run time (e.g., from 7 AM to 8 AM):

### Using PowerShell:
```powershell
$task = Get-ScheduledTask -TaskName "Birthday Card Sender"
$trigger = New-ScheduledTaskTrigger -Daily -At 08:00AM
Set-ScheduledTask -TaskName "Birthday Card Sender" -Trigger $trigger
```

### Using Task Scheduler GUI:
1. Open Task Scheduler
2. Find "Birthday Card Sender"
3. Right-click and select "Properties"
4. Go to "Triggers" tab
5. Select the trigger and click "Edit..."
6. Change the time
7. Click "OK"

## Logging

The task will log execution details to:
- **Application logs**: `d:\PlayGround\birthdaycard\birthday-bot\birthday_bot.log`
- **Windows Task Scheduler logs**: View in Event Viewer → Windows Logs → System

## Troubleshooting

### Task Not Running
1. Check if task is enabled: `Get-ScheduledTask -TaskName "Birthday Card Sender" | Select-Object State`
2. Verify time setting: `Get-ScheduledTaskInfo -TaskName "Birthday Card Sender"`
3. Check Windows Event Viewer for errors (Event Viewer → Windows Logs → System)

### Python/Script Not Found
- Verify Python path: `D:\Program Files\Python\Python314\python.exe` exists
- Verify script path: `d:\PlayGround\birthdaycard\birthday-bot\send_via_outlook.py` exists
- Check working directory is set correctly

### Outlook Not Sending
- Ensure Outlook is configured with the email account
- Check `.env` and `email_config.json` settings
- Verify sender email is configured in Windows Outlook

### Permission Errors
- Run PowerShell as Administrator
- Ensure the scheduler has permission to run the script
- Check Windows UAC (User Account Control) settings

## Features

✅ **Daily Execution**: Runs automatically every day at 7:00 AM  
✅ **Highest Privileges**: Runs with elevated permissions  
✅ **Network Aware**: Only runs if network is available  
✅ **Recovery**: Automatically retries if missed  
✅ **Error Handling**: Full logging of all operations  
✅ **Easy Management**: Enable/disable without deleting task  

## Testing the Setup

Before trusting the automation, test it manually:

```powershell
cd d:\PlayGround\birthdaycard\birthday-bot
python send_via_outlook.py
```

Verify it runs without errors. If there are any issues, fix them before relying on the scheduler.

## Backup/Export Task

Export the scheduled task for backup:

```powershell
Export-ScheduledTask -TaskName "Birthday Card Sender" -Path "birthday_card_task_backup.xml"
```

Import it on another machine:
```powershell
Import-ScheduledTask -Xml (Get-Content "birthday_card_task_backup.xml" | Out-String) -TaskName "Birthday Card Sender"
```

## FAQ

**Q: Will the script run if I shut down my computer?**  
A: No. The script only runs if your computer is on and the scheduled time is reached. The Task Scheduler can retry if the computer was off at the scheduled time.

**Q: Can I change the run time?**  
A: Yes, see "Change Schedule Time" section above.

**Q: What if there are no birthdays on a given day?**  
A: The script will run but won't send any emails (correct behavior). Check the logs to confirm it ran.

**Q: Can I run it multiple times a day?**  
A: Yes, create another task with a different name and time trigger.

**Q: How do I check if it ran today?**  
A: Check the log file: `d:\PlayGround\birthdaycard\birthday-bot\birthday_bot.log`

---

**Support**: For issues, check the log files and Windows Event Viewer for detailed error messages.
