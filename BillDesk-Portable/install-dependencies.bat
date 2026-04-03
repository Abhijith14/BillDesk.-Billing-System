@echo off
title BillDesk - Dependency Installer
color 0A

echo.
echo  ============================================
echo    BillDesk. - Dependency Installer
echo    For Windows 10 / Windows 11
echo  ============================================
echo.
echo  This script will check and install the
echo  required components to run BillDesk:
echo.
echo    1. .NET Framework 4.8+  (usually pre-installed)
echo    2. Microsoft Access Database Engine 2016
echo       (needed to read Excel data files)
echo.
echo  ============================================
echo.

:: Check for admin rights
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo  [!] Administrator privileges required.
    echo      Right-click this file and select
    echo      "Run as administrator"
    echo.
    pause
    exit /b 1
)

echo  [*] Running as Administrator... OK
echo.

:: -----------------------------------------------
:: CHECK 1: .NET Framework 4.8+
:: -----------------------------------------------
echo  [1/2] Checking .NET Framework...

reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release >nul 2>&1
if %errorlevel% neq 0 (
    echo  [X] .NET Framework 4.8 is NOT installed.
    echo      Downloading installer...
    echo.
    powershell -Command "Start-BitsTransfer -Source 'https://go.microsoft.com/fwlink/?linkid=2088631' -Destination '%TEMP%\ndp48-web.exe'"
    if exist "%TEMP%\ndp48-web.exe" (
        echo  [*] Installing .NET Framework 4.8...
        "%TEMP%\ndp48-web.exe" /passive /norestart
        echo  [OK] .NET Framework 4.8 installed.
    ) else (
        echo  [!] Download failed. Please install manually:
        echo      https://dotnet.microsoft.com/download/dotnet-framework/net48
    )
) else (
    for /f "tokens=3" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" /v Release 2^>nul ^| find "Release"') do set DOTNET_REL=%%a
    if %DOTNET_REL% GEQ 528040 (
        echo  [OK] .NET Framework 4.8 or later is installed.
    ) else (
        echo  [~] .NET Framework found but older than 4.8.
        echo      Downloading update...
        powershell -Command "Start-BitsTransfer -Source 'https://go.microsoft.com/fwlink/?linkid=2088631' -Destination '%TEMP%\ndp48-web.exe'"
        if exist "%TEMP%\ndp48-web.exe" (
            "%TEMP%\ndp48-web.exe" /passive /norestart
            echo  [OK] .NET Framework 4.8 installed.
        )
    )
)
echo.

:: -----------------------------------------------
:: CHECK 2: Microsoft Access Database Engine
:: -----------------------------------------------
echo  [2/2] Checking Microsoft Access Database Engine...

set ACE_FOUND=0

:: Check 64-bit registry
reg query "HKLM\SOFTWARE\Classes\Microsoft.ACE.OLEDB.16.0" >nul 2>&1
if %errorlevel% equ 0 (
    set ACE_FOUND=1
    echo  [OK] Microsoft.ACE.OLEDB.16.0 found (64-bit).
)

if %ACE_FOUND% equ 0 (
    reg query "HKLM\SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0" >nul 2>&1
    if %errorlevel% equ 0 (
        set ACE_FOUND=1
        echo  [OK] Microsoft.ACE.OLEDB.12.0 found (64-bit).
    )
)

:: Check 32-bit registry
if %ACE_FOUND% equ 0 (
    reg query "HKLM\SOFTWARE\WOW6432Node\Classes\Microsoft.ACE.OLEDB.16.0" >nul 2>&1
    if %errorlevel% equ 0 (
        set ACE_FOUND=1
        echo  [OK] Microsoft.ACE.OLEDB.16.0 found (32-bit).
    )
)

if %ACE_FOUND% equ 0 (
    reg query "HKLM\SOFTWARE\WOW6432Node\Classes\Microsoft.ACE.OLEDB.12.0" >nul 2>&1
    if %errorlevel% equ 0 (
        set ACE_FOUND=1
        echo  [OK] Microsoft.ACE.OLEDB.12.0 found (32-bit).
    )
)

if %ACE_FOUND% equ 0 (
    echo  [X] Access Database Engine is NOT installed.
    echo      Downloading installer...
    echo.
    powershell -Command "Start-BitsTransfer -Source 'https://download.microsoft.com/download/3/5/C/35C84C36-661A-44E6-9324-8786B8DBE231/AccessDatabaseEngine_X64.exe' -Destination '%TEMP%\AccessDatabaseEngine_X64.exe'"
    if exist "%TEMP%\AccessDatabaseEngine_X64.exe" (
        echo  [*] Installing Access Database Engine 2016 (64-bit)...
        "%TEMP%\AccessDatabaseEngine_X64.exe" /passive /norestart
        echo  [OK] Access Database Engine installed.
    ) else (
        echo  [!] Download failed. Please install manually:
        echo      https://www.microsoft.com/en-us/download/details.aspx?id=54920
    )
)
echo.

:: -----------------------------------------------
:: DONE
:: -----------------------------------------------
echo  ============================================
echo    All checks complete!
echo  ============================================
echo.
echo  You can now run Fees_Management.exe
echo.
pause
