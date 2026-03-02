@echo off
REM Birthday Card System - Interactive Menu
REM Double-click this file to run

setlocal enabledelayedexpansion
cd /d "%~dp0"

:menu
cls
echo.
echo ========================================
echo    Birthday Card System - Main Menu
echo ========================================
echo.
echo What would you like to do?
echo.
echo   1. Send Birthday Cards for TODAY
echo   2. Test Tomorrow's Workflow
echo   3. Generate Cards Only (No Email)
echo   4. Check Scheduler Status
echo   5. Change Sender Email
echo   6. Run Script in PowerShell
echo   7. View Configuration
echo   8. Exit
echo.
set /p choice="Enter your choice (1-8): "

if "%choice%"=="1" goto today
if "%choice%"=="2" goto tomorrow
if "%choice%"=="3" goto generate
if "%choice%"=="4" goto status
if "%choice%"=="5" goto sender
if "%choice%"=="6" goto pwsh
if "%choice%"=="7" goto config
if "%choice%"=="8" goto end
goto invalid

:today
cls
echo.
echo ========================================
echo    Running: Send Today's Birthday Cards
echo ========================================
echo.
echo Step 1: Clearing old cards...
del /q out\birthday_cards_today\* >nul 2>&1
echo Step 2: Generating today's cards...
python generate_today_birthdays.py
echo.
echo Step 3: Sending emails...
python send_via_outlook.py
echo.
echo Press any key to return to menu...
pause >nul
goto menu

:tomorrow
cls
echo.
echo ========================================
echo    Running: Tomorrow's Workflow Test
echo ========================================
echo.
python test_tomorrow_workflow.py
echo.
echo Press any key to return to menu...
pause >nul
goto menu

:generate
cls
echo.
echo ========================================
echo    Running: Generate Cards Only
echo ========================================
echo.
python generate_today_birthdays.py
echo.
echo Press any key to return to menu...
pause >nul
goto menu

:status
cls
echo.
echo ========================================
echo    Scheduler Status
echo ========================================
echo.
schtasks /query /tn "Birthday Card Sender" /v
echo.
echo Press any key to return to menu...
pause >nul
goto menu

:sender
cls
echo.
echo ========================================
echo    Change Sender Email
echo ========================================
echo.
echo Current Configuration:
echo.
python -c "import json; cfg=json.load(open('email_config.json')); e=cfg['email']; print(f'  Primary: {e.get(\"sender_email\")}'); print(f'  Alternative: {e.get(\"alt_sender_email\")}'); print(f'  Use Alternative: {e.get(\"use_alt_sender\")}')"
echo.
echo Options:
echo   1. Use Primary Email (abc@trna.com.vn)
echo   2. Use Alternative Email (dc34_pub@trna.com.vn)
echo   3. Back to Menu
echo.
set /p email_choice="Enter your choice (1-3): "

if "%email_choice%"=="1" (
    echo Switching to PRIMARY sender...
    python -c "import json; cfg=json.load(open('email_config.json')); cfg['email']['use_alt_sender']=False; json.dump(cfg, open('email_config.json','w'), indent=2)"
    echo Done! Primary sender activated.
) else if "%email_choice%"=="2" (
    echo Switching to ALTERNATIVE sender...
    python -c "import json; cfg=json.load(open('email_config.json')); cfg['email']['use_alt_sender']=True; json.dump(cfg, open('email_config.json','w'), indent=2)"
    echo Done! Alternative sender activated.
) else (
    goto menu
)
echo.
echo Press any key to return to menu...
pause >nul
goto menu

:config
cls
echo.
echo ========================================
echo    Current Configuration
echo ========================================
echo.
echo === .ENV FILE ===
echo.
findstr /v "^#" .env | findstr /v "^$"
echo.
echo === EMAIL CONFIGURATION FILE ===
echo.
type email_config.json
echo.
echo Press any key to return to menu...
pause >nul
goto menu

:pwsh
cls
echo.
echo ========================================
echo    PowerShell Commands Available
echo ========================================
echo.
echo You can run these commands in PowerShell:
echo.
echo   1. Check scheduler: Get-ScheduledTask -TaskName "Birthday Card Sender"
echo   2. Run manually: schtasks /run /tn "Birthday Card Sender"
echo   3. Disable: schtasks /change /tn "Birthday Card Sender" /disable
echo   4. Enable: schtasks /change /tn "Birthday Card Sender" /enable
echo   5. View logs: Get-Content birthday_bot.log -Tail 20
echo.
echo Press any key to return to menu...
pause >nul
goto menu

:invalid
cls
echo.
echo ERROR: Invalid choice. Please enter a number between 1 and 8.
echo.
echo Press any key to try again...
pause >nul
goto menu

:end
cls
echo.
echo Goodbye!
echo.
exit /b 0
