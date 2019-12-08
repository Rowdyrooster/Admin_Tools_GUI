#

#
<#
  ===================================================
   Written By: Joshua Montgomery
  ===================================================
#>


[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")  
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Windows.Forms.Application]::EnableVisualStyles() 

Import-Module -Name ActiveDirectory

$Main_Form = New-Object System.Windows.Forms.Form
$Main_Form.Text = 'Admin Tools'
$Main_Form.Width = 700
$Main_Form.Height = 600
$Main_Form.StartPosition = "CenterScreen"
$Main_Form.FormBorderStyle = 'Fixed3d'
$Main_Form.BackColor = "black"
$Main_Form.AutoSize = $true
$Main_Form.AutoScale = $true

#$icon = New-Object System.Drawing.Icon ("")
#$icon.Icon = $icon

$label = New-Object System.Windows.Forms.Label
$label.text = 'Last PSW Change'
$label.ForeColor = "White"
$label.location = New-Object System.Drawing.Point(0,10)
$label.AutoSize = $true
$Main_Form.Controls.Add($label)

$combobox = New-Object System.Windows.Forms.ComboBox
$combobox.Text = ""
$combobox.Width = 150
$combobox.Height = 10
$combobox.AutoCompleteMode = 'Suggest'
$combobox.AutoCompleteSource = 'ListItems'
$combobox.FormattingEnabled = $True
$combobox.AutoSize = $true
$combobox.Location = New-Object System.Drawing.Point(100,0)
$Main_Form.Controls.Add($combobox)


$Users = Get-ADUser -Filter * | Select-Object SamAccountName

Foreach ($Users in $Users)
{$combobox.Items.Add($Users.SamAccountName);
}

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(500,0)
$textBox.Size = New-Object System.Drawing.Size(450,500)
$textBox.AutoSize = $false
$textBox.Multiline = $true
$textBox.ReadOnly = $true
$textBox.WordWrap = $true
#$textbox.UseWaitCursor = $true
$textBox.TextAlign = 'Left'
$textBox.ScrollBars = 'Both'
$textBox.BackColor = "Black"
$textBox.ForeColor = "Green"
$textBox.Text = ""
$Main_Form.Controls.Add($textBox)

$label14 = New-Object System.Windows.Forms.Label
$label14.Text = "List of Enabled AD Users"
$label14.ForeColor = "White"
$label14.Location = New-Object System.Drawing.Point(0,40)
$main_form.Controls.Add($label14)

$label15 = New-Object System.Windows.Forms.Label
$label15.Text = "List of Disabled AD Users "
$label15.ForeColor = "White"
$label15.Location = New-Object System.Drawing.Point(0,80)
$main_form.Controls.Add($label15)

$combobox2 = New-Object System.Windows.Forms.ComboBox
$combobox2.Text = ""
$combobox2.Width = 150
$combobox2.Height = 10
$combobox2.AutoCompleteMode = 'Suggest'
$combobox2.AutoCompleteSource = 'ListItems'
$combobox2.FormattingEnabled = $True
$combobox2.AutoSize = $true
$combobox2.Location = New-Object System.Drawing.Point(100,80)
$Main_Form.Controls.Add($combobox2)

$disabledusers = Get-ADUser -Filter {Enabled -eq $false} | Select-Object SamAccountName

Foreach ($disabledusers in $disabledusers){
$combobox2.Items.Add($disabledusers.SamAccountName);
}

$combobox1 = New-Object System.Windows.Forms.ComboBox
$combobox1.Text = ""
$combobox1.Width = 150
$combobox1.Height = 10
$combobox1.AutoCompleteMode = 'Suggest'
$combobox1.AutoCompleteSource = 'ListItems'
$combobox1.FormattingEnabled = $True
$combobox1.AutoSize = $true
$combobox1.Location = New-Object System.Drawing.Point(100,40)
$Main_Form.Controls.Add($combobox1)

$enabledusers = Get-ADuser -Filter {Enabled -eq $true} | Select-Object SamAccountName

Foreach ($enabledusers in $enabledusers){
$combobox1.Items.Add($enabledusers.SamAccountName);
}

$label17 = New-Object System.Windows.Forms.Label
$label17.Text = "Tools"
$label17.ForeColor = "White"
$label17.Location = New-Object System.Drawing.Point(40,275)
$main_form.Controls.Add($label17)

$label18 = New-Object System.Windows.Forms.Label
$label18.Text = "AD Tools"
$label18.ForeColor = "White"
$label18.Location = New-Object System.Drawing.Point(150,275)
$main_form.Controls.Add($label18)

$label19 = New-Object System.Windows.Forms.Label
$label19.Text = "AD Manager Search"
$label19.ForeColor = "White"
$label19.Location = New-Object System.Drawing.Point(0,120)
$main_form.Controls.Add($label19)

#$label21 = New-Object System.Windows.Forms.Label
#$label21.Text = ""
#$label21.ForeColor = "White"
#$label21.Location = New-Object System.Drawing.Point(100,200)
#$main_form.Controls.Add($label21)

$label22 = New-Object System.Windows.Forms.Label
$label22.Text = "GpUpdate List Selection"
$label22.ForeColor = "White"
$label22.Location = New-Object System.Drawing.Point(0,160)
$main_form.Controls.Add($label22)

$label23 = New-Object System.Windows.Forms.Label
$label23.Text = "Ping Host"
$label23.ForeColor = "White"
$label23.Location = New-Object System.Drawing.Point(0,200)
$main_form.Controls.Add($label23)

$combobox5 = New-Object System.Windows.Forms.ComboBox
$combobox5.Text = ""
$combobox5.Width = 150
$combobox5.Height = 10
$combobox5.AutoCompleteMode = 'Suggest'
$combobox5.AutoCompleteSource = 'ListItems'
$combobox5.FormattingEnabled = $True
$combobox5.AutoSize = $false
$combobox5.Location = New-Object System.Drawing.Point(100,200)
$Main_Form.Controls.Add($combobox5)

$ping = Get-ADComputer -Filter * | Select-Object DNSHostName

Foreach($ping in $ping){
$combobox5.items.Add($ping.DNSHostName)
}

$combobox3 = New-Object System.Windows.Forms.ComboBox
$combobox3.Text = ""
$combobox3.Width = 150
$combobox3.Height = 10
$combobox3.AutoCompleteMode = 'Suggest'
$combobox3.AutoCompleteSource = 'ListItems'
$combobox3.FormattingEnabled = $True
$combobox3.AutoSize = $true
$combobox3.Location = New-Object System.Drawing.Point(100,120)
$Main_Form.Controls.Add($combobox3)

$Users2 = Get-ADUser -Filter * | Select-Object SamAccountName

Foreach ($Users2 in $Users2)
{$combobox3.Items.Add($Users2.SamAccountName);
}

$Side_Button = New-Object System.Windows.Forms.Button
$Side_Button.location = New-Object System.Drawing.Point(300,120)
$Side_Button.Size = New-Object System.Drawing.Size(120,23)
$Side_Button.Text = "Check"
$Side_Button.BackColor = "Green"
$Side_Button.Add_Click({$textBox.Text = ((Get-ADUser -Identity $combobox3.SelectedItem -Properties Manager).Manager).ToString()})
$main_form.controls.Add($Side_Button)

$Side_Button1 = New-Object System.Windows.Forms.Button
$Side_Button1.location = New-Object System.Drawing.Point(300,160)
$Side_Button1.Size = New-Object System.Drawing.Size(120,23)
$Side_Button1.AutoSize = $true
$Side_Button1.Text = "GPUpdate Execute"
$Side_Button1.BackColor = "Green"
$Side_Button1.Add_Click({$textBox.Text = Invoke-GPUpdate -Computer $combobox4.SelectedItem -Force | Out-String })
$main_form.controls.Add($Side_Button1)

$Side_Button2 = New-Object System.Windows.Forms.Button
$Side_Button2.location = New-Object System.Drawing.Point(300,200)
$Side_Button2.Size = New-Object System.Drawing.Size(120,23)
$Side_Button2.AutoSize = $true
$Side_Button2.Text = "Ping Execute"
$Side_Button2.BackColor = "Green"
$Side_Button2.Add_Click({$textBox.Text = ping $combobox5.SelectedItem | Out-String })
$main_form.controls.Add($Side_Button2)

$Side_Button3 = New-Object System.Windows.Forms.Button
$Side_Button3.location = New-Object System.Drawing.Point(300,0)
$Side_Button3.Size = New-Object System.Drawing.Size(120,23)
$Side_Button3.Text = "Check"
$Side_Button3.BackColor = "Green"
$Side_Button3.Add_Click({$textBox.Text = [datetime]::FromFileTime((Get-ADUser -Identity $combobox.SelectedItem -Properties pwdlastset).pwdlastset).ToString('MM dd yy')})
$main_form.controls.Add($Side_Button3)

$Button1 = New-Object System.Windows.Forms.Button
$Button1.location = New-Object System.Drawing.Point(300,40)
$Button1.Size = New-Object System.Drawing.Size(120,23)
$Button1.Text = "Open Grid"
$Button1.BackColor = "Green"
$Button1.Add_Click({Get-ADUser -Filter {Enabled -eq $true} | Select-Object SamAccountName, Enabled | Out-GridView })
$main_form.controls.Add($Button1)

$Button2 = New-Object System.Windows.Forms.Button
$Button2.location = New-Object System.Drawing.Point(300,80)
$Button2.Size = New-Object System.Drawing.Size(120,23)
$Button2.AutoSize = $true
$Button2.Text = "Open Grid"
$Button2.BackColor = "Green"
$Button2.Add_Click({Search-ADAccount -AccountDisabled -UsersOnly | Select-Object SamAccountName, Name, LastLogonDate, LockedOut, Enabled| Out-GridView})
$main_form.controls.Add($Button2)

$Button3 = New-Object System.Windows.Forms.Button
$Button3.location = New-Object System.Drawing.Point(0,300)
$Button3.Size = New-Object System.Drawing.Size(120,23)
$Button3.AutoSize = $true
$Button3.Text = "RDP Session"
$Button3.BackColor = "Green"
$Button3.Add_Click({mstsc.exe})
$main_form.controls.Add($Button3)

$Button4 = New-Object System.Windows.Forms.Button
$Button4.location = New-Object System.Drawing.Point(0,325)
$Button4.Size = New-Object System.Drawing.Size(120,23)
$Button4.AutoSize = $true
$Button4.Text = "CMD Window"
$Button4.BackColor = "Green"
$Button4.Add_Click({start cmd.exe})
$main_form.controls.Add($Button4)

$Button5 = New-Object System.Windows.Forms.Button
$Button5.location = New-Object System.Drawing.Point(0,350)
$Button5.Size = New-Object System.Drawing.Size(120,23)
$Button5.AutoSize = $true
$Button5.Text = "Powershell Window"
$Button5.BackColor = "Green"
$Button5.Add_Click({start powershell.exe})
$main_form.controls.Add($Button5)

$Button6 = New-Object System.Windows.Forms.Button
$Button6.location = New-Object System.Drawing.Point(0,375)
$Button6.Size = New-Object System.Drawing.Size(120,23)
$Button6.AutoSize = $true
$Button6.Text = "Send an Email"
$Button6.BackColor = "Green"
$Button6.Add_Click({start powershell.exe Send-MailMessage})
$main_form.controls.Add($Button6)

$Button7 = New-Object System.Windows.Forms.Button
$Button7.location = New-Object System.Drawing.Point(0,400)
$Button7.Size = New-Object System.Drawing.Size(120,23)
$Button7.AutoSize = $true
$Button7.Text = "Telnet Session"
$Button7.BackColor = "Green"
$Button7.Add_Click({start powershell.exe telnet})
$main_form.controls.Add($Button7)

$Button8 = New-Object System.Windows.Forms.Button
$Button8.location = New-Object System.Drawing.Point(0,425)
$Button8.Size = New-Object System.Drawing.Size(120,23)
$Button8.AutoSize = $true
$Button8.Text = "Putty Session"
$Button8.BackColor = "Green"
$Button8.Add_Click({putty.exe})
$main_form.controls.Add($Button8)

$Button9 = New-Object System.Windows.Forms.Button
$Button9.location = New-Object System.Drawing.Point(0,450)
$Button9.Size = New-Object System.Drawing.Size(120,23)
$Button9.AutoSize = $true
$Button9.Text = "Local IP Address"
$Button9.BackColor = "Green"
$Button9.Add_Click({$textBox.Text = Get-NetIPAddress -AddressFamily IPv4})
$main_form.controls.Add($Button9)

Get-NetIPAddress -AddressFamily IPv4 | Select-Object IPAddress
$Button10 = New-Object System.Windows.Forms.Button
$Button10.location = New-Object System.Drawing.Point(0,475)
$Button10.Size = New-Object System.Drawing.Size(120,23)
$Button10.AutoSize = $true
$Button10.Text = "Netstat"
$Button10.BackColor = "Green"
$Button10.Add_Click({$textBox.text = Netstat | out-string })
$main_form.controls.Add($Button10)

$combobox4 = New-Object System.Windows.Forms.ComboBox
$combobox4.Text = ""
$combobox4.Width = 150
$combobox4.Height = 10
$combobox4.AutoCompleteMode = 'Suggest'
$combobox4.AutoCompleteSource = 'ListItems'
$combobox4.FormattingEnabled = $True
$combobox4.AutoSize = $true
$combobox4.Location = New-Object System.Drawing.Point(100,160)
$Main_Form.Controls.Add($combobox4)

$Computer = Get-ADComputer -Filter * | Select-Object Name

Foreach ($Computer in $Computer)
{$combobox4.Items.Add($Computer.Name);
}

$Button11 = New-Object System.Windows.Forms.Button
$Button11.location = New-Object System.Drawing.Point(0,500)
$Button11.Size = New-Object System.Drawing.Size(120,23)
$Button11.AutoSize = $true
$Button11.Text = ""
$Button11.BackColor = "Green"
#$Button11.Add_Click({start powershell.exe Send-MailMessage})
$main_form.controls.Add($Button11)

$Button12 = New-Object System.Windows.Forms.Button
$Button12.location = New-Object System.Drawing.Point(0,525)
$Button12.Size = New-Object System.Drawing.Size(120,23)
$Button12.AutoSize = $true
$Button12.Text = ""
$Button12.BackColor = "Green"
#$Button12.Add_Click({start powershell.exe Send-MailMessage})
$main_form.controls.Add($Button12)

$Button13 = New-Object System.Windows.Forms.Button
$Button13.location = New-Object System.Drawing.Point(125,300)
$Button13.Size = New-Object System.Drawing.Size(120,23)
$Button13.AutoSize = $true
$Button13.Text = "Unlock AD Account"
$Button13.BackColor = "Green"
$Button13.Add_Click({start powershell.exe Unlock-ADAccount})
$main_form.controls.Add($Button13)

$Button14 = New-Object System.Windows.Forms.Button
$Button14.location = New-Object System.Drawing.Point(125,325)
$Button14.Size = New-Object System.Drawing.Size(120,23)
$Button14.AutoSize = $true
$Button14.Text = "Disable AD Account"
$Button14.BackColor = "Green"
$Button14.Add_Click({start powershell.exe Disable-ADAccount})
$main_form.controls.Add($Button14)

$Button15 = New-Object System.Windows.Forms.Button
$Button15.location = New-Object System.Drawing.Point(125,350)
$Button15.Size = New-Object System.Drawing.Size(120,23)
$Button15.AutoSize = $true
$Button15.Text = "Enable AD Account"
$Button15.BackColor = "Green"
$Button15.Add_Click({start powershell.exe Enable-ADAccount})
$main_form.controls.Add($Button15)

$Button16 = New-Object System.Windows.Forms.Button
$Button16.location = New-Object System.Drawing.Point(125,375)
$Button16.Size = New-Object System.Drawing.Size(120,23)
$Button16.AutoSize = $true
$Button16.Text = "Delete AD Account"
$Button16.BackColor = "Green"
$Button16.Add_Click({start powershell.exe Remove-ADUser})
$main_form.controls.Add($Button16)

$Button17 = New-Object System.Windows.Forms.Button
$Button17.location = New-Object System.Drawing.Point(125,400)
$Button17.Size = New-Object System.Drawing.Size(120,23)
$Button17.AutoSize = $true
$Button17.Text = "Locked Out Users"
$Button17.BackColor = "Green"
$Button17.Add_Click({Search-ADAccount -LockedOut | Select-Object SamAccountName, LockedOut, LastLogonDate  | Out-GridView})
$main_form.controls.Add($Button17)

$Button17 = New-Object System.Windows.Forms.Button
$Button17.location = New-Object System.Drawing.Point(125,425)
$Button17.Size = New-Object System.Drawing.Size(120,23)
$Button17.AutoSize = $true
$Button17.Text = "Expired AD Accounts"
$Button17.BackColor = "Green"
$Button17.Add_Click({Search-ADAccount -AccountExpired | Out-GridView})
$main_form.controls.Add($Button17)

$Button18 = New-Object System.Windows.Forms.Button
$Button18.location = New-Object System.Drawing.Point(125,450)
$Button18.Size = New-Object System.Drawing.Size(120,23)
$Button18.AutoSize = $true
$Button18.Text = "AD Accts Expiring"
$Button18.BackColor = "Green"
$Button18.Add_Click({Search-ADAccount -AccountExpiring})
$main_form.controls.Add($Button18)

$Button19 = New-Object System.Windows.Forms.Button
$Button19.location = New-Object System.Drawing.Point(125,475)
$Button19.Size = New-Object System.Drawing.Size(120,23)
$Button19.AutoSize = $true
$Button19.Text = "AD PC's Disabled"
$Button19.BackColor = "Green"
$Button19.Add_Click({Get-ADComputer -Filter {Enabled -eq $false} | Select-Object Name, Enabled | Out-GridView})
$main_form.controls.Add($Button19)

$Button20 = New-Object System.Windows.Forms.Button
$Button20.location = New-Object System.Drawing.Point(125,500)
$Button20.Size = New-Object System.Drawing.Size(120,23)
$Button20.AutoSize = $true
$Button20.Text = "AD PC's Enabled"
$Button20.BackColor = "Green"
$Button20.Add_Click({Get-ADComputer -Filter {Enabled -eq $true} | Select-Object Name, Enabled | Out-GridView})
$main_form.controls.Add($Button20)

$Button21 = New-Object System.Windows.Forms.Button
$Button21.location = New-Object System.Drawing.Point(125,525)
$Button21.Size = New-Object System.Drawing.Size(120,23)
$Button21.AutoSize = $true
$Button21.Text = "Aged AD Users"
$Button21.BackColor = "Green"
$Button21.Add_Click({$COMPAREDATE = GET-DATE
$NumberDays=120
$CSVFileLocation='C:\TEMP\OldComps.CSV'
$output = Search-ADAccount -AccountInactive -TimeSpan 90 -UsersOnly | Where-Object { $_.Enabled -eq $true } | Select-object Name, UserPrincipalName, LastLogonDate | Out-GridView})
$main_form.controls.Add($Button21)

$Button22 = New-Object System.Windows.Forms.Button
$Button22.location = New-Object System.Drawing.Point(250,300)
$Button22.Size = New-Object System.Drawing.Size(120,23)
$Button22.AutoSize = $true
$Button22.Text = "Aged AD PC's"
$Button22.BackColor = "Green"
$Button22.Add_Click({$COMPAREDATE = GET-DATE
$NoContact = Search-ADAccount -AccountInactive -TimeSpan 90 -ComputersOnly | Where-Object { $_.Enabled -eq $True } | Select-Object Name, DNSHostName, LastLogonDate | Out-GridView})
$main_form.controls.Add($Button22)

$Button23 = New-Object System.Windows.Forms.Button
$Button23.location = New-Object System.Drawing.Point(250,325)
$Button23.Size = New-Object System.Drawing.Size(120,23)
$Button23.AutoSize = $true
$Button23.Text = "Active AD Users"
$Button23.BackColor = "Green"
$Button23.Add_Click({Get-ADUser -Filter {Enabled -eq $true} |Select-object SamAccountName, Enabled | Out-GridView})
$main_form.controls.Add($Button23)

$Main_Form.ShowDialog()
