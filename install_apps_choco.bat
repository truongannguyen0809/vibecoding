@echo off
setlocal EnableDelayedExpansion

REM Install Chocolatey if not installed
where choco >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Installing Chocolatey...
    powershell -NoProfile -InputFormat None -ExecutionPolicy Bypass -Command "Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))"
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to install Chocolatey.
        exit /b %ERRORLEVEL%
    )
)

REM Define list of popular applications
set total_apps=20
set app1=googlechrome
set app2=firefox
set app3=microsoft-edge
set app4=brave
set app5=opera
set app6=7zip
set app7=notepadplusplus
set app8=git
set app9=vscode
set app10=vlc
set app11=telegram
set app12=obs-studio
set app13=microsoft-365-apps
set app14=microsoft-teams
set app15=zoom
set app16=sharex
set app17=messenger
set app18=foxitreader
set app19=bulk-crap-uninstaller
set app20=advanced-renamer

echo Available applications:
for /l %%i in (1,1,%total_apps%) do (
    echo %%i. !app%%i!
)

set /p choices=Enter numbers of apps to install separated by spaces ^(or 'all' for all^): 

if /i "%choices%"=="all" (
    set selected=
    for /l %%i in (1,1,%total_apps%) do (
        set selected=!selected! !app%%i!
    )
) else (
    set selected=
    for %%i in (%choices%) do (
        set selected=!selected! !app%%i!
    )
)

if "!selected!"=="" (
    echo No applications selected. Exiting.
    exit /b 0
)

REM Validate that selected packages exist on Chocolatey
set valid=
for %%p in (!selected!) do (
    choco info %%p >nul 2>&1
    if !errorlevel! EQU 0 (
        set valid=!valid! %%p
    ) else (
        echo Package %%p not found on Chocolatey, skipping.
    )
)

if "!valid!"=="" (
    echo No valid packages selected. Exiting.
    exit /b 0
)

echo Installing valid packages: !valid!
choco install -y !valid!

