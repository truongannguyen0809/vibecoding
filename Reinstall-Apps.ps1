<# 
.SYNOPSIS
    Detect, select, uninstall, clean, and reinstall applications on Windows using winget.

.DESCRIPTION
    - Enumerates installed apps via winget (and augments with registry info for InstallLocation).
    - Lets you select apps interactively (Out-GridView if available, else console prompt) or pass -AppsById/-AppsByName.
    - Uninstalls via winget, optionally includes Microsoft Store apps with -IncludeStore.
    - Cleans leftover InstallLocation directories (safe-ish cleanup).
    - Reinstalls via winget (latest), optionally pin to current version with -PinToCurrentVersion (best-effort).
    - Logs to a timestamped file.

.PARAMETER AppsById
    One or more winget PackageIdentifiers to process (skips interactive selection).

.PARAMETER AppsByName
    One or more display names to match (best-effort fuzzy match; you’ll be shown what matched).

.PARAMETER IncludeStore
    Include Microsoft Store (msstore) packages.

.PARAMETER PinToCurrentVersion
    Try to reinstall the same version that was previously installed (if winget supports that version).

.PARAMETER DryRun
    Show what would be done without changing the system.

.PARAMETER CreateRestorePoint
    Attempt to create a system restore point before making changes.

.EXAMPLE
    # Interactive: pick apps, uninstall, cleanup, reinstall latest
    .\Reinstall-Apps.ps1

.EXAMPLE
    # Non-interactive by Id
    .\Reinstall-Apps.ps1 -AppsById Microsoft.VisualStudioCode, 7zip.7zip

.EXAMPLE
    # Include Microsoft Store apps and try to pin to the current version
    .\Reinstall-Apps.ps1 -IncludeStore -PinToCurrentVersion
#>

[CmdletBinding()]
param(
    [string[]] $AppsById,
    [string[]] $AppsByName,
    [switch]   $IncludeStore,
    [switch]   $PinToCurrentVersion,
    [switch]   $DryRun,
    [switch]   $CreateRestorePoint
)

#---------------------------- Helpers ----------------------------#

function Assert-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent()
       ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) { throw "Please run PowerShell as Administrator." }
}

function Assert-Winget {
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "winget (Windows Package Manager) is required. Update Windows 10/11 or install winget."
    }
}

$Script:LogPath = Join-Path -Path $PSScriptRoot -ChildPath ("ReinstallApps-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "u"), $Level.ToUpper(), $Message
    Write-Host $line
    Add-Content -Path $Script:LogPath -Value $line
}

function Try-Checkpoint {
    param([string]$Description)
    try {
        # May fail if System Protection is disabled
        Checkpoint-Computer -Description $Description -RestorePointType "MODIFY_SETTINGS" | Out-Null
        Write-Log "Created restore point: $Description"
    } catch {
        Write-Log "Could not create restore point: $($_.Exception.Message)" "WARN"
    }
}

# Parse 'winget list' output text into objects (Name, Id, Version, Source)
function Get-WingetInstalled {
    Write-Log "Querying installed packages via winget..."
    $out = & winget list --accept-source-agreements 2>$null
    if (-not $out) { return @() }

    # Skip header lines and handle variable columns
    $lines = $out | Where-Object { $_ -and $_.Trim() -ne "" } | Select-Object -Skip 2
    $apps = foreach ($line in $lines) {
        if ($line -match '^-{3,}') { continue }
        $cols = ($line -replace '\s{2,}', '|').Split('|')
        if ($cols.Count -ge 4) {
            if ($cols.Count -eq 4) {
                $name,$id,$version,$source = $cols
                [pscustomobject]@{
                    Name    = $name.Trim()
                    Id      = $id.Trim()
                    Version = $version.Trim()
                    Source  = $source.Trim()
                }
            } elseif ($cols.Count -ge 5) {
                $name,$id,$version,$available,$source = $cols[0..4]
                [pscustomobject]@{
                    Name    = $name.Trim()
                    Id      = $id.Trim()
                    Version = $version.Trim()
                    Source  = $source.Trim()
                }
            }
        }
    }

    # De-duplicate by Id if necessary
    $apps | Group-Object Id | ForEach-Object { $_.Group | Select-Object -First 1 }
}

# Registry uninstall entries -> used to fetch InstallLocation for safer cleanup
function Get-RegistryUninstallEntries {
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $all = foreach ($p in $paths) {
        Get-ItemProperty -Path $p -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -and ($_.DisplayName.Trim() -ne "")
        } | Select-Object DisplayName, DisplayVersion, InstallLocation, UninstallString, PSPath, Publisher
    }

    # collapse duplicates by DisplayName
    $all | Group-Object DisplayName | ForEach-Object { $_.Group | Select-Object -First 1 }
}

function Join-AppData {
    param([object[]]$WingetApps, [object[]]$RegEntries)

    foreach ($app in $WingetApps) {
        $match = $RegEntries | Where-Object { $_.DisplayName -eq $app.Name } | Select-Object -First 1
        if (-not $match) {
            $pattern = ('^{0}($|\s|\-|\()' -f [regex]::Escape($app.Name))
            $match = $RegEntries | Where-Object { $_.DisplayName -match $pattern } | Select-Object -First 1
        }

        $installLocation = $null
        $displayVersion  = $null
        $publisher       = $null
        $regKey          = $null

        if ($match) {
            $installLocation = $match.InstallLocation
            $displayVersion  = $match.DisplayVersion
            $publisher       = $match.Publisher
            $regKey          = $match.PSPath
        }

        [pscustomobject]@{
            Name            = $app.Name
            Id              = $app.Id
            Version         = if ($displayVersion) { $displayVersion } else { $app.Version }
            Source          = $app.Source
            InstallLocation = $installLocation
            RegistryKey     = $regKey
            Publisher       = $publisher
        }
    }
}

function Select-Apps {
    param([object[]]$Catalog, [string[]]$ById, [string[]]$ByName, [switch]$IncludeStore)

    $filtered = if ($IncludeStore) {
        $Catalog
    } else {
        $Catalog | Where-Object { $_.Source -ne 'msstore' }
    }

    if ($ById -and $ById.Count) {
        $chosen = foreach ($id in $ById) {
            $hit = $filtered | Where-Object { $_.Id -eq $id } | Select-Object -First 1
            if (-not $hit) { Write-Log "No match for Id '$id'." "WARN" } else { $hit }
        }
        return ($chosen | Where-Object { $_ })
    }

    if ($ByName -and $ByName.Count) {
        $chosen = foreach ($name in $ByName) {
            $hit = $filtered | Where-Object { $_.Name -like $name } |
                   Sort-Object Name | Select-Object -First 1
            if (-not $hit) { Write-Log "No match for Name '$name'." "WARN" } else { $hit }
        }
        return ($chosen | Where-Object { $_ })
    }

    # Interactive selection
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        $filtered | Sort-Object Name |
            Out-GridView -Title "Select apps to uninstall & reinstall (multi-select, then press OK)" -PassThru
    } else {
        Write-Host ""
        Write-Host "Interactive selection (no Out-GridView)."
        Write-Host "Enter numbers separated by spaces or commas."
        Write-Host ""
        $list = $filtered | Sort-Object Name | Select-Object Name,Id,Version,Source
        $i=1
        $map = @{}
        foreach ($item in $list) {
            Write-Host ("[{0}] {1}    ({2})    v{3}    [{4}]" -f $i, $item.Name, $item.Id, $item.Version, $item.Source)
            $map[$i] = $item
            $i++
        }
        $input = Read-Host "Pick one or more"
        $indices = $input -split '[, ]+' | Where-Object { $_ } | ForEach-Object { $_ -as [int] } | Where-Object { $_ -ge 1 -and $_ -lt $i }
        foreach ($n in $indices) { $map[$n] }
    }
}

function Invoke-Winget {
    param([string[]]$Args, [switch]$DryRun)

    Write-Log ("winget {0}" -f ($Args -join ' '))
    if ($DryRun) { return 0 }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = "winget"
    if ($psi.PSObject.Properties.Name -contains 'ArgumentList' -and $null -ne $psi.ArgumentList) {
        $psi.ArgumentList.AddRange($Args)
    } else {
        $psi.Arguments = ($Args -join ' ')
    }
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $stdout = $proc.StandardOutput.ReadToEnd()
    $stderr = $proc.StandardError.ReadToEnd()
    $proc.WaitForExit()

    if ($stdout) { Write-Log $stdout.Trim() }
    if ($stderr) { Write-Log $stderr.Trim() "WARN" }

    return $proc.ExitCode
}

function Uninstall-App {
    param($App, [switch]$DryRun)

    $args = @('uninstall','--id', $App.Id, '--accept-source-agreements','-e','--silent')
    if ($App.Source) { $args += @('--source', $App.Source) }

    $code = Invoke-Winget -Args $args -DryRun:$DryRun
    if ($code -ne 0) { Write-Log "Uninstall exit code $code for $($App.Name)." "WARN" }
    return $code
}

function Cleanup-App {
    param($App, [switch]$DryRun)

    if ($App.InstallLocation -and (Test-Path -LiteralPath $App.InstallLocation)) {
        Write-Log "Removing InstallLocation: $($App.InstallLocation)"
        if (-not $DryRun) {
            try {
                # Make files writable, then remove
                Get-ChildItem -LiteralPath $App.InstallLocation -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
                    try { $_.IsReadOnly = $false } catch {}
                }
                Remove-Item -LiteralPath $App.InstallLocation -Recurse -Force -ErrorAction Stop
            } catch {
                Write-Log "Failed to remove $($App.InstallLocation): $($_.Exception.Message)" "WARN"
            }
        }
    } else {
        Write-Log "No InstallLocation found for $($App.Name); skipping cleanup."
    }
}

function Reinstall-App {
    param($App, [switch]$PinToCurrentVersion, [switch]$DryRun)

    $args = @('install','--id', $App.Id, '--accept-source-agreements','--accept-package-agreements','-e','--silent')
    if ($App.Source) { $args += @('--source', $App.Source) }
    if ($PinToCurrentVersion -and $App.Version) { $args += @('--version', $App.Version) }

    $code = Invoke-Winget -Args $args -DryRun:$DryRun
    if ($code -ne 0) { Write-Log "Install exit code $code for $($App.Name)." "WARN" }
    return $code
}

#---------------------------- Main ----------------------------#

try {
    Assert-Admin
    Assert-Winget

    if ($CreateRestorePoint) {
        Try-Checkpoint -Description "Before Reinstall-Apps"
    }

    $wg = Get-WingetInstalled
    if (-not $wg -or $wg.Count -eq 0) {
        throw "No installed packages were returned by winget."
    }

    $reg = Get-RegistryUninstallEntries
    $catalog = Join-AppData -WingetApps $wg -RegEntries $reg

    $selected = Select-Apps -Catalog $catalog -ById $AppsById -ByName $AppsByName -IncludeStore:$IncludeStore
    if (-not $selected -or $selected.Count -eq 0) {
        Write-Log "Nothing selected. Exiting."
        return
    }

    Write-Log ("Selected {0} app(s): {1}" -f $selected.Count, (($selected | ForEach-Object { $_.Name }) -join ", "))

    foreach ($app in $selected) {
        if (-not $IncludeStore -and $app.Source -eq 'msstore') {
            Write-Log "Skipping Microsoft Store app (use -IncludeStore to process): $($app.Name)" "WARN"
            continue
        }

        Write-Log "----- Processing: $($app.Name) [$($app.Id)] v$($app.Version) -----"
        $u = Uninstall-App -App $app -DryRun:$DryRun
        if ($u -eq 0) {
            Cleanup-App -App $app -DryRun:$DryRun
            $i = Reinstall-App -App $app -PinToCurrentVersion:$PinToCurrentVersion -DryRun:$DryRun
            if ($i -eq 0) {
                Write-Log "Reinstalled: $($app.Name)"
            } else {
                Write-Log "Failed to reinstall: $($app.Name)" "WARN"
            }
        } else {
            Write-Log "Uninstall reported failure; skipping reinstall for: $($app.Name)" "WARN"
        }
    }

    Write-Log "Done. Log: $Script:LogPath"
} catch {
    Write-Log ("ERROR: {0}" -f $_.Exception.Message) "ERROR"
    throw
}
