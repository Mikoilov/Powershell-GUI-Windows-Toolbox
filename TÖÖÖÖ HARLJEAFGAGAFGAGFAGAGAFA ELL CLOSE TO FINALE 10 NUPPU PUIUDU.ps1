Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Form ja control objectid
$FormObject = [System.Windows.Forms.Form]
$TabControlObject = [System.Windows.Forms.TabControl]
$TabPageObject = [System.Windows.Forms.TabPage]
$LabelObject = [System.Windows.Forms.Label]
$ButtonObject = [System.Windows.Forms.Button]
$ComboBoxObject = [System.Windows.Forms.Combobox]
$PanelObject = [System.Windows.Forms.Panel]
$CheckBoxObject = [System.Windows.Forms.CheckBox]

# Peamine form
$MareForm = New-Object $FormObject
$MareForm.ClientSize = '1280,1024'
$MareForm.Text = 'Miko Windows Toolbox'
$MareForm.BackColor = "#ffffff"

# Tabcontrol ja Tabid
$TabControl = New-Object $TabControlObject
$TabControl.Size = '1260,1000'
$TabControl.Location = New-Object System.Drawing.Point(10,10)

$TabPage1 = New-Object $TabPageObject
$TabPage1.Text = "Main Tools"
$TabPage2 = New-Object $TabPageObject
$TabPage2.Text = "Services"
$TabPage3 = New-Object $TabPageObject
$TabPage3.Text = "Rakendused"

# Controls esimese tabi jaoks (Main Tools)
$lbltitle = New-Object $LabelObject
$lbltitle.Text = 'Tools'
$lbltitle.Autosize = $true
$lbltitle.Font = 'times new roman,24,style=italic'
$lbltitle.Location = New-Object System.Drawing.Point(10,10)

$btnhello = New-Object $ButtonObject
$btnhello.Text = 'Ava aken'
$btnhello.Autosize = $true
$btnhello.Font = 'times new roman,12'
$btnhello.Location = New-Object System.Drawing.Point(10,70)

$TabPage1.Controls.AddRange(@($lbltitle, $btnhello))

# Controls teise tabi jaoks (Services)
$DefaultFont = 'Times new roman,12'

$lblService = New-Object $LabelObject
$lblService.Text = 'Services :'
$lblService.Autosize = $true
$lblService.Location = New-Object System.Drawing.Point(20,20)

$ddlService = New-Object $ComboBoxObject
$ddlService.Width = '300'
$ddlService.Location = New-Object System.Drawing.Point(125,20)

# Comboboxi paneb serviceid
Get-Service | ForEach-Object {$ddlService.Items.Add($_.Name)}

$ddlService.Text = 'Vali Service'

$lblForName = New-Object $LabelObject
$lblForName.Text = 'Service Friendly nimi :'
$lblForName.Autosize = $true
$lblForName.Location = New-Object System.Drawing.Point(20,80)

$lblName = New-Object $LabelObject
$lblName.Text = ''
$lblName.Autosize = $true
$lblName.Location = New-Object System.Drawing.Point(240,80)

$lblForStatus = New-Object $LabelObject
$lblForStatus.Text = 'Status'
$lblForStatus.Autosize = $true
$lblForStatus.Location = New-Object System.Drawing.Point(20,120)

$lblStatus = New-Object $LabelObject
$lblStatus.Text = ''
$lblStatus.Autosize = $true
$lblStatus.Location = New-Object System.Drawing.Point(240,120)

$TabPage2.Controls.AddRange(@($lblService, $ddlService, $lblForName, $lblName, $lblForStatus, $lblStatus))

# Controls kolmanda tabi jaoks (Rakendused)
$panel = New-Object $PanelObject
$panel.Size = New-Object System.Drawing.Size(350, 350)
$panel.Location = New-Object System.Drawing.Point(20, 20)

# Kõik rakendused
$applications = @("Chrome", "Opera", "Firefox", "Zoom", "Skype", "Teamviewer 15", "CCleaner", "WinDirStat", "Audacity", "VLC", "Spotify", "PuTTY", "KeePass 2", "7-Zip", "WinRAR")

# CheckBoxid iga rakenduse jaoks
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

# Install nupp
$installButton = New-Object $ButtonObject
$installButton.Text = "Install Selected Applications"
$installButton.Size = New-Object System.Drawing.Size(350, 30)
$installButton.Location = New-Object System.Drawing.Point(20, 380)

$TabPage3.Controls.AddRange(@($panel, $installButton))

# Install commandid
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

# Event handler install nupu jaoks
$installButton.Add_Click({
    foreach ($app in $checkboxes.Keys) {
        if ($checkboxes[$app].Checked) {
            Write-Output "Installing $app..."
            Invoke-Expression $installCommands[$app]
        }
    }
    [System.Windows.Forms.MessageBox]::Show("Selected applications installed successfully.", "Installation Complete")
})

# Tabid Tabcontrolile
$TabControl.TabPages.AddRange(@($TabPage1, $TabPage2, $TabPage3))

# Formi tabcontroli lisamine
$MareForm.Controls.Add($TabControl)

# Loogilised functionid esimese tabi jaoks
function SayHello {
    if ($lbltitle.Text -eq '') {
        $lbltitle.Text = 'Sinine marps'
    } else {
        $lbltitle.Text = ''
    }
}

# Event handlerid esimese tabi jaoks
$btnhello.add_Click({ SayHello })

# Functionid teise tabi jaoks
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

# Event handlerid teise tabi jaoks
$ddlService.Add_SelectedIndexChanged({ GetServiceDetails })

# Näita formi
$MareForm.ShowDialog()

# Cleanup
$MareForm.Dispose()