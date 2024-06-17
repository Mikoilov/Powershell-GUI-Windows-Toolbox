Add-Type -TypeDefinition @"
    using System;
    using System.Runtime.InteropServices;

    public class NativeMethods {
        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        public const int VK_LWIN = 0x5B;
        public const int VK_D = 0x44;
        public const int KEYEVENTF_KEYDOWN = 0;
        public const int KEYEVENTF_KEYUP = 2;

        public static void ShowDesktop() {
            keybd_event((byte)VK_LWIN, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero);
            keybd_event((byte)VK_D, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero);
            keybd_event((byte)VK_D, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            keybd_event((byte)VK_LWIN, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
        }
    }
"@
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#Form ja control objectid
$FormObject = [System.Windows.Forms.Form]
$TabControlObject = [System.Windows.Forms.TabControl]
$TabPageObject = [System.Windows.Forms.TabPage]
$LabelObject = [System.Windows.Forms.Label]
$ButtonObject = [System.Windows.Forms.Button]
$ComboBoxObject = [System.Windows.Forms.Combobox]
$PanelObject = [System.Windows.Forms.Panel]
$CheckBoxObject = [System.Windows.Forms.CheckBox]

#Peamine form
$MareForm = New-Object $FormObject
$MareForm.ClientSize = '1280,1024'
$MareForm.Text = 'Miko Windows Toolbox'
$MareForm.BackColor = "#ffffff"

#Tabcontrol ja Tabid
$TabControl = New-Object $TabControlObject
$TabControl.Size = '1260,1000'
$TabControl.Location = New-Object System.Drawing.Point(10,10)

$TabPage1 = New-Object $TabPageObject
$TabPage1.Text = "Peamised riistad"
$TabPage1.BackColor="#E718E7"
$TabPage2 = New-Object $TabPageObject
$TabPage2.Text = "Teenused"
$TabPage2.BackColor = "#FFC0CB"
$TabPage3 = New-Object $TabPageObject
$TabPage3.Text = "Rakendused"
$TabPage3.BackColor= "#9000CC"
$TabPage4 = New-Object $TabPageObject
$TabPage4.Text = "CLI"
$TabPage4.BackColor= "#93E9BE"


#Controls esimese tabi jaoks (Main Tools)
$lbltitle = New-Object $LabelObject
$lbltitle.Text = 'Riistad'
$lbltitle.Autosize = $true
$lbltitle.Font = 'times new roman,24,style=italic'
$lbltitle.Location = New-Object System.Drawing.Point(10,10)

$btnclrclip = New-Object $ButtonObject
$btnclrclip.Text = 'Clipboard tyhjaks'
$btnclrclip.Autosize = $true
$btnclrclip.Font = 'times new roman,12'
$btnclrclip.Location = New-Object System.Drawing.Point(10,70)

$btnspat = New-Object $ButtonObject
$btnspat.Text = 'KINNI'
$btnspat.Autosize = $true
$btnspat.Font = 'times new roman,12'
$btnspat.Location = New-Object System.Drawing.Point(10,130)

$btnIP = New-Object $ButtonObject
$btnIP.Text = 'Vaata IPd'
$btnIP.Autosize = $true
$btnIP.Font = 'times new roman,12'
$btnIP.Location = New-Object System.Drawing.Point(10,200)

$btnSysInfo = New-Object $ButtonObject
$btnSysInfo.Text = 'Systeemi info'
$btnSysInfo.Autosize = $true
$btnSysInfo.Font = 'times new roman,12'
$btnSysInfo.Location = New-Object System.Drawing.Point(10,270)

$btndisksize = New-Object $ButtonObject
$btndisksize.Text = 'Ketta ruum'
$btndisksize.Autosize = $true
$btndisksize.Font = 'times new roman,12'
$btndisksize.Location = New-Object System.Drawing.Point(10,340)

$btnminwin = New-Object $ButtonObject
$btnminwin.Text = 'Naita Desktopi'
$btnminwin.Autosize = $true
$btnminwin.Font = 'times new roman,12'
$btnminwin.Location = New-Object System.Drawing.Point(10,410)

$btnlockscreen = New-Object $ButtonObject
$btnlockscreen.Text = 'Lock Screen'
$btnlockscreen.Autosize = $true
$btnlockscreen.Font = 'times new roman,12'
$btnlockscreen.Location = New-Object System.Drawing.Point(10,480)

$btntime = New-Object $ButtonObject
$btntime.Text = 'AEG'
$btntime.Autosize = $true
$btntime.Font = 'times new roman,12'
$btntime.Location = New-Object System.Drawing.Point(10,550)

$listBoxUSBDevices = New-Object System.Windows.Forms.ListBox
$listBoxUSBDevices.Location = New-Object System.Drawing.Point(200, 80)
$listBoxUSBDevices.Size = New-Object System.Drawing.Size(250, 40)
$listBoxUSBDevices.SelectionMode = "Pole"
$Tabpage1.Controls.Add($listBoxUSBDevices)

$btnrefresh = New-Object System.Windows.Forms.Button
$btnrefresh.Text = "Varskenda USB Valikuid"
$btnrefresh.Location = New-Object System.Drawing.Point(200, 120)
$btnrefresh.Size = New-Object System.Drawing.Size(150, 40)
$TabPage1.Controls.Add($buttonRefresh)

$textboxPassword = New-Object System.Windows.Forms.TextBox
$textboxPassword.Location = New-Object System.Drawing.Point(10, 660)
$textboxPassword.Size = New-Object System.Drawing.Size(300, 30)
$TabPage1.Controls.Add($textboxPassword)

$btnGenPass = New-Object $ButtonObject
$btnGenPass.Text = 'Genereeri password'
$btnGenPass.Autosize = $true
$btnGenPass.Font = 'times new roman,12'
$btnGenPass.Location = New-Object System.Drawing.Point(10,620)

$btncopytoclipboard = New-Object System.Windows.Forms.Button
$btncopytoclipboard.Text = "Copy to Clipboard"
$btncopytoclipboard.Location = New-Object System.Drawing.Point(150,700)
$btncopytoclipboard.Size = New-Object System.Drawing.Size(150, 40)
$TabPage1.Controls.Add($buttonCopyToClipboard)

$TabPage1.Controls.AddRange(@($lbltitle, $btnclrclip, $btnspat, $btnIP, $btnSysInfo, $btndisksize, $btnminwin, $btnlockscreen, $btntime, $listBoxUSBDevices, $btnrefresh, $textboxPassword, $btnGenPass, $btncopytoclipboard))

#Controls teise tabi jaoks (Services)
$DefaultFont = 'Times new roman,12'

$lblService = New-Object $LabelObject
$lblService.Text = 'Teenused :'
$lblService.Autosize = $true
$lblService.Location = New-Object System.Drawing.Point(20,20)

$ddlService = New-Object $ComboBoxObject
$ddlService.Width = '300'
$ddlService.Location = New-Object System.Drawing.Point(125,20)

#Comboboxi paneb serviceid
Get-Service | ForEach-Object {$ddlService.Items.Add($_.Name)}

$ddlService.Text = 'Vali Service'
$lblForName = New-Object $LabelObject
$lblForName.Text = 'Teenuse sobralik nimi :'
$lblForName.Autosize = $true
$lblForName.Location = New-Object System.Drawing.Point(20,80)

$lblName = New-Object $LabelObject
$lblName.Text = ''
$lblName.Autosize = $true
$lblName.Location = New-Object System.Drawing.Point(240,80)

$lblForStatus = New-Object $LabelObject
$lblForStatus.Text = 'Olek'
$lblForStatus.Autosize = $true
$lblForStatus.Location = New-Object System.Drawing.Point(20,120)

$lblStatus = New-Object $LabelObject
$lblStatus.Text = ''
$lblStatus.Autosize = $true
$lblStatus.Location = New-Object System.Drawing.Point(240,120)

$TabPage2.Controls.AddRange(@($lblService, $ddlService, $lblForName, $lblName, $lblForStatus, $lblStatus))

#Controls kolmanda tabi jaoks (Rakendused)
$panel = New-Object $PanelObject
$panel.Size = New-Object System.Drawing.Size(350, 350)
$panel.Location = New-Object System.Drawing.Point(20, 20)

#Kõik rakendused
$applications = @("Chrome", "Opera", "Firefox", "Zoom", "Skype", "Teamviewer 15", "CCleaner", "WinDirStat", "Audacity", "VLC", "Spotify", "PuTTY", "KeePass 2", "7-Zip", "WinRAR")

#CheckBoxid iga rakenduse jaoks
$checkboxes = @{}
$yPos = 10
foreach ($app in $applications) {
    $checkbox = New-Object $CheckBoxObject
    $checkbox.Text = $app
    $checkbox.AutoSize = $true
    $checkbox.Location = New-Object System.Drawing.Point(10, $yPos)
    $panel.Controls.Add($checkbox)
    $checkboxes[$app] = $checkbox
    $yPos += 25
}

#Install nupp
$installButton = New-Object $ButtonObject
$installButton.Text = "Lae alla valitud rakendused"
$installButton.Size = New-Object System.Drawing.Size(350, 30)
$installButton.Location = New-Object System.Drawing.Point(20, 380)

$TabPage3.Controls.AddRange(@($panel, $installButton))

#Install commandid
$installCommands = @{
    "Chrome" = 'Invoke-WebRequest -Uri "https://dl.google.com/chrome/install/latest/chrome_installer.exe" -OutFile "chrome_installer.exe"; Start-Process -FilePath "chrome_installer.exe" -ArgumentList "/silent /install" -Wait'
    "Opera" = 'Invoke-WebRequest -Uri "https://net.geo.opera.com/opera/stable/windows" -OutFile "opera_installer.exe"; Start-Process -FilePath "opera_installer.exe" -ArgumentList "/silent /install" -Wait'
    "Firefox" = 'Invoke-WebRequest -Uri "https://download.mozilla.org/?product=firefox-latest-ssl&os=win&lang=en-US" -OutFile "firefox_installer.exe"; Start-Process -FilePath "firefox_installer.exe" -ArgumentList "-ms" -Wait'
    "Zoom" = 'Invoke-WebRequest -Uri "https://zoom.us/client/latest/ZoomInstaller.exe" -OutFile "ZoomInstaller.exe"; Start-Process -FilePath "ZoomInstaller.exe" -ArgumentList "/silent" -Wait'
    "Skype" = 'Invoke-WebRequest -Uri "https://go.skype.com/windows.desktop.download" -OutFile "SkypeInstaller.exe"; Start-Process -FilePath "SkypeInstaller.exe" -ArgumentList "/silent" -Wait'
    "Teamviewer 15" = 'Invoke-WebRequest -Uri "https://download.teamviewer.com/download/TeamViewer_Setup.exe" -OutFile "TeamViewer_Setup.exe"; Start-Process -FilePath "TeamViewer_Setup.exe" -ArgumentList "/S" -Wait'
    "CCleaner" = 'Invoke-WebRequest -Uri "https://download.ccleaner.com/ccsetup.exe" -OutFile "ccsetup.exe"; Start-Process -FilePath "ccsetup.exe" -ArgumentList "/S" -Wait'
    "WinDirStat" = 'Invoke-WebRequest -Uri "https://downloads.sourceforge.net/project/windirstat/windirstat/1.1.2/WinDirStat1_1_2_setup.exe" -OutFile "WinDirStat1_1_2_setup.exe"; Start-Process -FilePath "WinDirStat1_1_2_setup.exe" -ArgumentList "/S" -Wait'
    "Audacity" = 'Invoke-WebRequest -Uri "https://www.fosshub.com/Audacity.html/audacity-win-3.0.2.exe" -OutFile "audacity.exe"; Start-Process -FilePath "audacity.exe" -ArgumentList "/verysilent" -Wait'
    "VLC" = 'Invoke-WebRequest -Uri "https://get.videolan.org/vlc/3.0.11.1/win32/vlc-3.0.11.1-win32.exe" -OutFile "vlc.exe"; Start-Process -FilePath "vlc.exe" -ArgumentList "/S" -Wait'
    "Spotify" = 'Invoke-WebRequest -Uri "https://download.scdn.co/SpotifySetup.exe" -OutFile "SpotifySetup.exe"; Start-Process -FilePath "SpotifySetup.exe" -ArgumentList "/silent" -Wait'
    "PuTTY" = 'Invoke-WebRequest -Uri "https://the.earth.li/~sgtatham/putty/latest/w64/putty-64bit-0.74-installer.msi" -OutFile "putty_installer.msi"; Start-Process -FilePath "msiexec.exe" -ArgumentList "/i putty_installer.msi /quiet" -Wait'
    "KeePass 2" = 'Invoke-WebRequest -Uri "https://downloads.sourceforge.net/project/keepass/KeePass%202.x/2.47/KeePass-2.47-Setup.exe" -OutFile "KeePass-Setup.exe"; Start-Process -FilePath "KeePass-Setup.exe" -ArgumentList "/silent" -Wait'
    "7-Zip" = 'Invoke-WebRequest -Uri "https://www.7-zip.org/a/7z1900-x64.exe" -OutFile "7z1900-x64.exe"; Start-Process -FilePath "7z1900-x64.exe" -ArgumentList "/S" -Wait'
    "WinRAR" = 'Invoke-WebRequest -Uri "https://www.win-rar.com/fileadmin/winrar-versions/winrar/winrar-x64-600.exe" -OutFile "winrar-x64-600.exe"; Start-Process -FilePath "winrar-x64-600.exe" -ArgumentList "/S" -Wait'
}

#Event handler install nupu jaoks
$installButton.Add_Click({
    foreach ($app in $checkboxes.Keys) {
        if ($checkboxes[$app].Checked) {
            Write-Output "Installing $app..."
            Invoke-Expression $installCommands[$app]
        }
    }
    [System.Windows.Forms.MessageBox]::Show("Selected applications installed successfully.", "Installation Complete")
})

#Controls Neljanda Tabi jaoks (CLI: Input/Output)

# Textbox outputi jaoks
$textBoxoutput = New-Object System.Windows.Forms.TextBox
$textBoxoutput.Multiline = $true
$textBoxoutput.ScrollBars = "Vertical"
$textBoxoutput.Width = 560
$textBoxoutput.Height = 100
$textBoxoutput.Location = New-Object System.Drawing.Point(10,10)
$textBoxoutput.Font = New-Object System.Drawing.Font("Consolas", 10)
$TabPage4.Controls.Add($textBoxoutput)

#Textbox Inputi jaoks
$textBoxinput = New-Object System.Windows.Forms.TextBox
$textBoxinput.Multiline = $false #yhe realine input
$textBoxinput.Width = 400
$textBoxinput.Height = 30
$textBoxinput.Location = New-Object System.Drawing.Point(10, 120)
$textBoxinput.Font = New-Object System.Drawing.Font("Consolas", 10)
$TabPage4.Controls.Add($textBoxinput)

#Nupp et executida command
$btnexec = New-Object System.Windows.Forms.Button
$btnexec.Text = "Kaivita Skript"
$btnexec.Width = 100
$btnexec.Height = 30
$btnexec.Location = New-Object System.Drawing.Point(420,115)
$TabPage4.Controls.Add($btnexec)

#Function mis uuendab textboxi cli outputiga
function Execute-Script {
    param (
        [string]$script
    )

    try {
        $output = Invoke-Expression -Command $script 2>&1 | Out-String
        $textBoxOutput.Text = $output
    } catch {
        $textBoxOutput.Text = $_.Exception.Message
    }
}


#Eventi handler nupu vajutuse jaoks
$btnexec.Add_Click({
    Execute-Script -script $textBoxinput.Text
    })

$TabPage4.Controls.AddRange(@($textBoxoutput,$btnexec))

#Tabid Tabcontrolile
$TabControl.TabPages.AddRange(@($TabPage1, $TabPage2, $TabPage3, $TabPage4))

#Formi tabcontroli lisamine
$MareForm.Controls.Add($TabControl)

#Loogilised functionid esimese tabi jaoks
#Function IP addressi jaoks
function Check-IP {
    $ip = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -ne 'Loopback Pseudo-Interface 1' }).IPAddress
    [System.Windows.Forms.MessageBox]::Show("Praegune IP aadress: $ip", "IP Aadress")
}

#Function Systeemi info jaoks
function show-Systeminfo {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem
    $processor = Get-CimInstance -ClassName Win32_Processor
    $totalmemory = Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum | Foreach {"{0:N2}" -f ([math]::round(($_.Sum / 1GB),2))}
    $systemInfo = @"
Operatsiooni systeem: $($os.Caption) $($os.Version)
Arvuti nimi: $($os.CSName)
Protsessor: $($processor.Name)
RAM: $totalMemory GB
"@
    [System.Windows.Forms.MessageBox]::Show($systemInfo, "Systeemi info")
}

#Function Ketta Ruumi jaoks
function Check-DiskSpace {
    $diskInfo = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 } | Select-Object DeviceID, VolumeName, @{Name="FreeSpaceGB"; Expression={[math]::Round($_.FreeSpace / 1GB, 2)}}, @{Name="SizeGB"; Expression={[math]::Round($_.Size / 1GB, 2)}}
    
    $diskReport = "Ketta Ruumi Informatsioon`n`n"
    foreach ($disk in $diskInfo) {
        $diskReport += "Drive: $($disk.DeviceID)`n"
        $diskReport += "Volume Name: $($disk.VolumeName)`n"
        $diskReport += "Free Space: $($disk.FreeSpaceGB) GB`n"
        $diskReport += "Total Size: $($disk.SizeGB) GB`n`n"
    }
    
    [System.Windows.Forms.MessageBox]::Show($diskReport, "Ketta Ruumi informatsioon")
}

#Function et teha Clipboard Tyhjaks
function Clear-Clipboard {
    [System.Windows.Forms.Clipboard]::Clear()
    [System.Windows.Forms.MessageBox]::Show("Clipboard has been cleared.", "Clear Clipboard", "OK", "Information")
}

#Function panna arvuti sunniviisiliselt kinni
function Forcefully-ShutDownComputer {
    Stop-Computer -ComputerName localhost
}

#Paneb koik desktopi windowid minimized
function Show-Desktop {
try {
    [NativeMethods]::ShowDesktop()
    [System.Windows.Forms.MessageBox]::Show("Desktop is shown.", "Show Desktop", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error showing desktop:`n$($_.Exception.Message)", "Show Desktop", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
}

#Function Refreshida USB Seadmeid
function Refresh-USBDevices {
    $listBoxUSBDevices.Items.Clear()
    $usbDevices = Get-CimInstance -Query "SELECT * FROM Win32_USBControllerDevice" | ForEach-Object {
        $usbDevice = Get-CimInstance -ClassName Win32_PnPEntity -Filter "DeviceID = '$($_.Dependent.DeviceID -replace '\\', '\\')'"
        [PSCustomObject]@{
            'Device Name' = $usbDevice.Name
            'Description' = $usbDevice.Description
            'Manufacturer' = $usbDevice.Manufacturer
            'Service' = $usbDevice.Service
        }
    }
    
    foreach ($device in $usbDevices) {
        $listBoxUSBDevices.Items.Add("Device Name: $($device.'Device Name')")
        $listBoxUSBDevices.Items.Add("Description: $($device.Description)")
        $listBoxUSBDevices.Items.Add("Manufacturer: $($device.Manufacturer)")
        $listBoxUSBDevices.Items.Add("Service: $($device.Service)")
        $listBoxUSBDevices.Items.Add("")
    }
}

#Function genereerida suvaline passworrdddd
function Generate-RandomPassword {
    #Defineeri mis karaktereid kasutada
    $chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()-_=+[]{}|;:',<.>/?"
    
    #Pane passwordi pikkus
    $passwordLength = 12
    
    #Genereeri suvaline password
    $password = -join ((0..($passwordLength-1)) | ForEach-Object { $chars[(Get-Random -Minimum 0 -Maximum $chars.Length)] })
    
    #Naita password ette
    $textboxPassword.Text = $password
}



#Event handlerid Esimese tabi jaoks
$btnIP.Add_Click({Check-IP})

$btnSysInfo.Add_Click({show-Systeminfo})

$btndisksize.Add_Click({Check-DiskSpace})

$btnrestore.Add_Click({Create-RestorePoint})

$btnspat.Add_Click({Forcefully-ShutDownComputer})

$btnminwin.Add_Click({Show-Desktop})

$btnclrclip.Add_Click({Clear-Clipboard})

$btnlockscreen.Add_Click({
    #Kutsub Lock Screen Commandi esile
    try {
        rundll32.exe user32.dll,LockWorkStation
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to lock the screen:`n$($_.Exception.Message)", "Lock Screen", "OK", "Error")
    }
})

$btntime.Add_Click({
    $currentDateTime = Get-Date
    [System.Windows.Forms.MessageBox]::Show("Current Date and Time:`n$currentDateTime", "Date and Time")
})

$btnrefresh.Add_Click({
    Refresh-USBDevices
})

#Laeb USB Seadmed
Refresh-USBDevices

$btnGenPass.Add_Click({Generate-RandomPassword})

$btncopytoclipboard.Add_Click({
    if ($textboxPassword.Text -ne "") {
        $textboxPassword.SelectAll()
        $textboxPassword.Copy()
        [System.Windows.Forms.MessageBox]::Show("Password copied to clipboard.", "Copy Password")
    } else {
        [System.Windows.Forms.MessageBox]::Show("No password generated.", "Copy Password", "Error")
    }
})

#Functionid Teise tabi jaoks
function GetServiceDetails{
    $ServiceName = $ddlService.SelectedItem
    $Details = Get-Service -Name $ServiceName | select *
    $lblName.Text = $details.name
    $lblstatus.Text = $details.status

    if ($lblstatus.Text -eq 'Running') {
        $lblstatus.ForeColor = 'green'
    } else {
        $lblstatus.ForeColor = 'red'
    }
}

$btnrefresh.Add_Click({Refresh-USBDevices})


#Event handlerid teise tabi jaoks
$ddlService.Add_SelectedIndexChanged({ GetServiceDetails })

#Näita formi
$MareForm.ShowDialog()

#Cleanup
$MareForm.Dispose()