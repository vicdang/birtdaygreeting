#!/usr/bin/env powershell
# Birthday Card System - Interactive Menu GUI Version

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Show-MainMenu {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Birthday Card System - Control Center"
    $form.Size = New-Object System.Drawing.Size(500, 600)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    $form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 245)
    
    # Title Label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Birthday Card System"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(50, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(400, 40)
    $form.Controls.Add($titleLabel)
    
    # Today Button
    $btnToday = New-Object System.Windows.Forms.Button
    $btnToday.Text = "📧 Send Today's Cards"
    $btnToday.Location = New-Object System.Drawing.Point(30, 80)
    $btnToday.Size = New-Object System.Drawing.Size(440, 50)
    $btnToday.Font = New-Object System.Drawing.Font("Arial", 11)
    $btnToday.BackColor = [System.Drawing.Color]::FromArgb(52, 168, 224)
    $btnToday.ForeColor = [System.Drawing.Color]::White
    $btnToday.Add_Click({
        $form.Hide()
        Set-Location $scriptDir
        Remove-Item "out\birthday_cards_today\*" -Force -ErrorAction SilentlyContinue
        python generate_today_birthdays.py
        python send_via_outlook.py
        [System.Windows.Forms.MessageBox]::Show("Today's workflow completed!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Show()
    })
    $form.Controls.Add($btnToday)
    
    # Tomorrow Button
    $btnTomorrow = New-Object System.Windows.Forms.Button
    $btnTomorrow.Text = "🗓️ Test Tomorrow's Workflow"
    $btnTomorrow.Location = New-Object System.Drawing.Point(30, 145)
    $btnTomorrow.Size = New-Object System.Drawing.Size(440, 50)
    $btnTomorrow.Font = New-Object System.Drawing.Font("Arial", 11)
    $btnTomorrow.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $btnTomorrow.ForeColor = [System.Drawing.Color]::White
    $btnTomorrow.Add_Click({
        $form.Hide()
        Set-Location $scriptDir
        python test_tomorrow_workflow.py
        [System.Windows.Forms.MessageBox]::Show("Tomorrow's workflow completed!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Show()
    })
    $form.Controls.Add($btnTomorrow)
    
    # Generate Only Button
    $btnGenerate = New-Object System.Windows.Forms.Button
    $btnGenerate.Text = "🎨 Generate Cards Only"
    $btnGenerate.Location = New-Object System.Drawing.Point(30, 210)
    $btnGenerate.Size = New-Object System.Drawing.Size(440, 50)
    $btnGenerate.Font = New-Object System.Drawing.Font("Arial", 11)
    $btnGenerate.BackColor = [System.Drawing.Color]::FromArgb(255, 152, 0)
    $btnGenerate.ForeColor = [System.Drawing.Color]::White
    $btnGenerate.Add_Click({
        $form.Hide()
        Set-Location $scriptDir
        python generate_today_birthdays.py
        [System.Windows.Forms.MessageBox]::Show("Card generation completed!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Show()
    })
    $form.Controls.Add($btnGenerate)
    
    # Sender Email Button
    $btnSender = New-Object System.Windows.Forms.Button
    $btnSender.Text = "⚙️ Change Sender Email"
    $btnSender.Location = New-Object System.Drawing.Point(30, 275)
    $btnSender.Size = New-Object System.Drawing.Size(440, 50)
    $btnSender.Font = New-Object System.Drawing.Font("Arial", 11)
    $btnSender.BackColor = [System.Drawing.Color]::FromArgb(156, 39, 176)
    $btnSender.ForeColor = [System.Drawing.Color]::White
    $btnSender.Add_Click({ Show-SenderMenu })
    $form.Controls.Add($btnSender)
    
    # Status Button
    $btnStatus = New-Object System.Windows.Forms.Button
    $btnStatus.Text = "📊 Check Scheduler Status"
    $btnStatus.Location = New-Object System.Drawing.Point(30, 340)
    $btnStatus.Size = New-Object System.Drawing.Size(440, 50)
    $btnStatus.Font = New-Object System.Drawing.Font("Arial", 11)
    $btnStatus.BackColor = [System.Drawing.Color]::FromArgb(63, 81, 181)
    $btnStatus.ForeColor = [System.Drawing.Color]::White
    $btnStatus.Add_Click({
        $task = Get-ScheduledTask -TaskName "Birthday Card Sender" -ErrorAction SilentlyContinue
        if ($task) {
            $info = Get-ScheduledTaskInfo -TaskName "Birthday Card Sender"
            $msg = "Task: $($task.TaskName)`nStatus: $($task.State)`nNext Run: $($info.NextRunTime)`nLast Run: $($info.LastRunTime)"
            [System.Windows.Forms.MessageBox]::Show($msg, "Scheduler Status", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Scheduler task not found!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($btnStatus)
    
    # Exit Button
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text = "Exit"
    $btnExit.Location = New-Object System.Drawing.Point(30, 520)
    $btnExit.Size = New-Object System.Drawing.Size(440, 40)
    $btnExit.Font = New-Object System.Drawing.Font("Arial", 11)
    $btnExit.BackColor = [System.Drawing.Color]::FromArgb(128, 128, 128)
    $btnExit.ForeColor = [System.Drawing.Color]::White
    $btnExit.Add_Click({ $form.Close() })
    $form.Controls.Add($btnExit)
    
    $form.ShowDialog()
}

function Show-SenderMenu {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Sender Email"
    $form.Size = New-Object System.Drawing.Size(400, 250)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Which email should be used as the sender?"
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(360, 40)
    $label.Font = New-Object System.Drawing.Font("Arial", 11)
    $form.Controls.Add($label)
    
    $btnPrimary = New-Object System.Windows.Forms.Button
    $btnPrimary.Text = "Primary: abc@trna.com.vn"
    $btnPrimary.Location = New-Object System.Drawing.Point(20, 80)
    $btnPrimary.Size = New-Object System.Drawing.Size(360, 50)
    $btnPrimary.Font = New-Object System.Drawing.Font("Arial", 10)
    $btnPrimary.BackColor = [System.Drawing.Color]::FromArgb(52, 168, 224)
    $btnPrimary.ForeColor = [System.Drawing.Color]::White
    $btnPrimary.Add_Click({
        Set-Location $scriptDir
        python -c "import json; cfg=json.load(open('email_config.json')); cfg['email']['use_alt_sender']=False; json.dump(cfg, open('email_config.json','w'), indent=2)"
        [System.Windows.Forms.MessageBox]::Show("Primary sender activated!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Close()
    })
    $form.Controls.Add($btnPrimary)
    
    $btnAlt = New-Object System.Windows.Forms.Button
    $btnAlt.Text = "Alternative: dc34_pub@trna.com.vn"
    $btnAlt.Location = New-Object System.Drawing.Point(20, 145)
    $btnAlt.Size = New-Object System.Drawing.Size(360, 50)
    $btnAlt.Font = New-Object System.Drawing.Font("Arial", 10)
    $btnAlt.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $btnAlt.ForeColor = [System.Drawing.Color]::White
    $btnAlt.Add_Click({
        Set-Location $scriptDir
        python -c "import json; cfg=json.load(open('email_config.json')); cfg['email']['use_alt_sender']=True; json.dump(cfg, open('email_config.json','w'), indent=2)"
        [System.Windows.Forms.MessageBox]::Show("Alternative sender activated!", "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        $form.Close()
    })
    $form.Controls.Add($btnAlt)
    
    $form.ShowDialog()
}

# Run the menu
Show-MainMenu
