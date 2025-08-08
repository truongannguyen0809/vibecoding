# IT Support Toolkit - Command Line Version
# This PowerShell script provides a text-based interface
# for common IT support tasks such as reinstalling applications
# via Ninite or Chocolatey, installing optional IT tools, and
# reinstalling device drivers. The script mirrors the functionality
# of the GUI version but uses console prompts instead.

function Show-Message($message) {
    Write-Host "[INFO] $message"
}

function Show-Error($message) {
    Write-Host "[ERROR] $message" -ForegroundColor Red
}

function Reinstall-NiniteApps {
    $script = @"
Get-Package -Name 'Google Chrome','Mozilla Firefox','VLC media player','7-Zip','Notepad++' | ForEach-Object { Uninstall-Package -Name `$_.Name -Force }
Invoke-WebRequest -OutFile NiniteInstaller.exe https://ninite.com/chrome-firefox-vlc-7zip-np++/ninite.exe
Start-Process NiniteInstaller.exe -Wait
"@
    Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -Command `$script = @'`n$script`n'@; Invoke-Expression `$script" -Wait
    Show-Message "Ninite reinstall completed."
}

function Reinstall-ChocoApps {
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
            $shell = New-Object -ComObject WScript.Shell
            $desktopPath = [Environment]::GetFolderPath("Desktop")
            $shortcut = $shell.CreateShortcut((Join-Path $desktopPath "$($app.Name).lnk"))
            $shortcut.TargetPath = $app.Path
            $shortcut.Save()
        }
    }
    Show-Message "Chocolatey reinstall completed."
}

function Install-ITTools {
    $toolNames = @("cpu-z", "hwmonitor", "crystaldiskinfo", "sysinternals", "nirlauncher")
    Write-Host "Select tools to install (comma separated numbers):"
    for ($i=0; $i -lt $toolNames.Count; $i++) {
        Write-Host "$($i+1)) $($toolNames[$i])"
    }
    $selection = Read-Host "Enter your choices"
    $indexes = $selection -split ',' | ForEach-Object { ($_ -as [int]) - 1 }
    $selected = @()
    foreach ($idx in $indexes) {
        if ($idx -ge 0 -and $idx -lt $toolNames.Count) {
            $selected += $toolNames[$idx]
        }
    }
    if ($selected.Count -eq 0) {
        Show-Error "No valid tools selected."
        return
    }
    foreach ($tool in $selected) {
        $script = @"
Set-ExecutionPolicy Bypass -Scope Process -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
choco install $tool -y --ignore-checksums
"@
        Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -Command `$script = @'`n$script`n'@; Invoke-Expression `$script" -Wait
    }
    Show-Message "Selected IT tools installed."
}

function Manage-Drivers {
    $devices = Get-PnpDevice | Where-Object { $_.Status -eq 'OK' -and $_.Class -ne 'SoftwareDevice' }
    if ($devices.Count -eq 0) {
        Show-Error "No devices found."
        return
    }
    Write-Host "Select a device to reinstall driver:";
    for ($i=0; $i -lt $devices.Count; $i++) {
        Write-Host "$($i+1)) $($devices[$i].FriendlyName)"
    }
    $choice = (Read-Host "Enter number") -as [int]
    if ($choice -lt 1 -or $choice -gt $devices.Count) {
        Show-Error "Invalid device selection."
        return
    }
    $instanceId = $devices[$choice-1].InstanceId
    $driverPath = Read-Host "Enter path to driver folder"
    if (-not (Test-Path $driverPath)) {
        Show-Error "Driver folder not found."
        return
    }
    $script = @"
Disable-PnpDevice -InstanceId '$instanceId' -Confirm:$false
Uninstall-PnpDevice -InstanceId '$instanceId' -Confirm:$false
Start-Sleep -Seconds 2
pnputil /add-driver '$driverPath\*.inf' /install
"@
    Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -Command `$script = @'`n$script`n'@; Invoke-Expression `$script" -Wait
    Show-Message "Driver reinstalled for selected device."
}

function Show-Menu {
    Write-Host "IT Support Toolkit"
    Write-Host "1) Reinstall Apps (Ninite)"
    Write-Host "2) Reinstall Apps (Chocolatey)"
    Write-Host "3) Install Apps (Choose)"
    Write-Host "4) Install IT Tools"
    Write-Host "5) Driver Management"
    Write-Host "Q) Quit"
}

while ($true) {
    Show-Menu
    $input = Read-Host "Select option"
    switch ($input) {
        '1' { Reinstall-NiniteApps }
        '2' { Reinstall-ChocoApps }
        '3' { Show-Message "Use a preconfigured Chocolatey script for custom apps." }
        '4' { Install-ITTools }
        '5' { Manage-Drivers }
        'q' { break }
        'Q' { break }
        default { Show-Error "Unknown selection." }
    }
}

