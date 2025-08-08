Add-Type -AssemblyName System.Windows.Forms

function Show-MessageBox($text, $title) {
    [System.Windows.Forms.MessageBox]::Show($text, $title)
}

function Show-PackageInfo($packageName) {
    $info = choco info $packageName | Out-String
    Show-MessageBox $info "Package Info: $packageName"
}

function Create-DesktopShortcut($name, $targetPath) {
    $desktopPath = [Environment]::GetFolderPath("Desktop")
    $shortcutPath = Join-Path $desktopPath "$name.lnk"
    if (Test-Path $shortcutPath) { Remove-Item $shortcutPath -Force }
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $targetPath
    $shortcut.Save()
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "IT Support Toolkit"
$form.Size = New-Object System.Drawing.Size(400, 450)
$form.StartPosition = "CenterScreen"

# Dropdown for actions
$actionLabel = New-Object System.Windows.Forms.Label
$actionLabel.Text = "Select Action:"
$actionLabel.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($actionLabel)

$actionDropdown = New-Object System.Windows.Forms.ComboBox
$actionDropdown.Location = New-Object System.Drawing.Point(120, 20)
$actionDropdown.Width = 200
$actionDropdown.Items.AddRange(@(
    "Reinstall Apps (Ninite)",
    "Reinstall Apps (Chocolatey)",
    "Install Apps (Choose)",
    "Install IT Tools",
    "Driver Management"
))
$form.Controls.Add($actionDropdown)

$actionDropdown.Add_SelectedIndexChanged({
    $runButton.Enabled = $true
    $itToolsGroup.Visible = ($actionDropdown.SelectedItem -eq "Install IT Tools")
})

# Device dropdown for drivers
$deviceLabel = New-Object System.Windows.Forms.Label
$deviceLabel.Text = "Select Device (Drivers):"
$deviceLabel.Location = New-Object System.Drawing.Point(10, 70)
$form.Controls.Add($deviceLabel)

$deviceDropdown = New-Object System.Windows.Forms.ComboBox
$deviceDropdown.Location = New-Object System.Drawing.Point(120, 70)
$deviceDropdown.Width = 250
$form.Controls.Add($deviceDropdown)

# Populate devices
$devices = Get-PnpDevice | Where-Object { $_.Status -eq 'OK' -and $_.Class -ne 'SoftwareDevice' }
foreach ($dev in $devices) {
    $item = New-Object PSObject -Property @{
        DisplayName = $dev.FriendlyName
        InstanceId = $dev.InstanceId
    }
    $deviceDropdown.Items.Add($item)
}
$deviceDropdown.DisplayMember = 'DisplayName'

# Button to choose driver folder
$driverButton = New-Object System.Windows.Forms.Button
$driverButton.Text = "Select Driver Folder"
$driverButton.Location = New-Object System.Drawing.Point(10, 120)
$form.Controls.Add($driverButton)

$driverPathLabel = New-Object System.Windows.Forms.Label
$driverPathLabel.Text = "No folder selected"
$driverPathLabel.Location = New-Object System.Drawing.Point(150, 125)
$driverPathLabel.Width = 200
$form.Controls.Add($driverPathLabel)

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$driverButton.Add_Click({
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $driverPathLabel.Text = $folderBrowser.SelectedPath
    }
})

# Group box for IT tool selection
$itToolsGroup = New-Object System.Windows.Forms.GroupBox
$itToolsGroup.Text = "Select IT Tools to Install"
$itToolsGroup.Location = New-Object System.Drawing.Point(10, 160)
$itToolsGroup.Size = New-Object System.Drawing.Size(360, 120)
$itToolsGroup.Visible = $false
$form.Controls.Add($itToolsGroup)

$toolCheckboxes = New-Object System.Collections.Generic.List[System.Windows.Forms.CheckBox]
$toolNames = @("cpu-z", "hwmonitor", "crystaldiskinfo", "sysinternals", "nirlauncher")

for ($i = 0; $i -lt $toolNames.Count; $i++) {
    $x = 10 + (([int]($i % 2)) * 180)
    $y = 20 + ([int]([math]::Floor($i / 2)) * 30)

    $toolCheckbox = New-Object System.Windows.Forms.CheckBox
    $toolCheckbox.Text = $toolNames[$i]
    $toolCheckbox.Location = New-Object System.Drawing.Point($x, $y)
    $toolCheckbox.Add_MouseDoubleClick({ Show-PackageInfo $_.Text })
    $itToolsGroup.Controls.Add($toolCheckbox)
    $toolCheckboxes.Add($toolCheckbox)
}

# Run button
$runButton = New-Object System.Windows.Forms.Button
$runButton.Text = "Run"
$runButton.Location = New-Object System.Drawing.Point(150, 300)
$runButton.Enabled = $false
$form.Controls.Add($runButton)

$runButton.Add_Click({
    switch ($actionDropdown.SelectedItem) {
        "Reinstall Apps (Ninite)" {
            $script = @"
Get-Package -Name 'Google Chrome','Mozilla Firefox','VLC media player','7-Zip','Notepad++' | ForEach-Object { Uninstall-Package -Name `$_.Name -Force }
Invoke-WebRequest -OutFile NiniteInstaller.exe https://ninite.com/chrome-firefox-vlc-7zip-np++/ninite.exe
Start-Process NiniteInstaller.exe -Wait
"@
            Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -Command `$script = @'`n$script`n'@; Invoke-Expression `$script"
            Show-MessageBox "Ninite reinstall completed." "Done"
        }
        "Reinstall Apps (Chocolatey)" {
            $apps = @(
                @{ Name = "googlechrome"; Path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" },
                @{ Name = "firefox"; Path = "C:\\Program Files\\Mozilla Firefox\\firefox.exe" },
                @{ Name = "vlc"; Path = "C:\\Program Files\\VideoLAN\\VLC\\vlc.exe" },
                @{ Name = "7zip"; Path = "C:\\Program Files\\7-Zip\\7zFM.exe" },
                @{ Name = "notepadplusplus"; Path = "C:\\Program Files\\Notepad++\\notepad++.exe" }
            )

            $installScript = @"
Set-ExecutionPolicy Bypass -Scope Process -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
"@

            foreach ($app in $apps) {
                $installScript += "choco uninstall $($app.Name) -y --ignore-checksums --remove-dependencies`n"
                $installScript += "choco install $($app.Name) -y --ignore-checksums --force`n"
            }

            $installScriptPath = "$env:TEMP\\install_apps.ps1"
            Set-Content -Path $installScriptPath -Value $installScript
            Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$installScriptPath`"" -Wait

            foreach ($app in $apps) {
                if (Test-Path $app.Path) {
                    Create-DesktopShortcut -name $app.Name -targetPath $app.Path
                }
            }
            Show-MessageBox "Chocolatey reinstall completed." "Done"
        }
        "Install Apps (Choose)" {
            Show-MessageBox "Use a preconfigured Chocolatey script for custom apps." "Info"
        }
        "Install IT Tools" {
            $selectedTools = $toolCheckboxes | Where-Object { $_.Checked } | ForEach-Object { $_.Text }
            if ($selectedTools.Count -eq 0) {
                Show-MessageBox "Please select at least one tool." "Error"
                return
            }
            foreach ($tool in $selectedTools) {
                $script = @"
Set-ExecutionPolicy Bypass -Scope Process -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
choco install $tool -y --ignore-checksums
"@
                Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -Command `$script = @'`n$script`n'@; Invoke-Expression `$script" -Wait
            }
            Show-MessageBox "Selected IT Tools installed." "Done"
        }
        "Driver Management" {
            if (-not $deviceDropdown.SelectedItem) {
                Show-MessageBox "Please select a device." "Error"
                return
            }
            if ($driverPathLabel.Text -eq "No folder selected") {
                Show-MessageBox "Please select a driver folder." "Error"
                return
            }

            $instanceId = $deviceDropdown.SelectedItem.InstanceId
            $driverPath = $driverPathLabel.Text

            $script = @"
Disable-PnpDevice -InstanceId '$instanceId' -Confirm:$false
Uninstall-PnpDevice -InstanceId '$instanceId' -Confirm:$false
Start-Sleep -Seconds 2
pnputil /add-driver '$driverPath\*.inf' /install
"@
            Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -Command `$script = @'`n$script`n'@; Invoke-Expression `$script"
            Show-MessageBox "Driver reinstalled for selected device." "Done"
        }
    }
})

# Show form
[void]$form.ShowDialog()
