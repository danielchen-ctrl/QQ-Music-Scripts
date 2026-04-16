@echo off
setlocal
cd /d "%~dp0"

:: 解除从网络下载（GitHub pull）时 Windows 附加的 Internet 来源限制
powershell.exe -NoProfile -Command "Get-ChildItem '%~dp0*.ps1' | Unblock-File" 2>nul

where adb >nul 2>&1
if errorlevel 1 (
    powershell.exe -NoProfile -Command "$msg = 'adb not found. Please make sure adb is installed and available in PATH.'; $title = 'ADB Missing'; Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show($msg, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)" 1>nul
    exit /b 1
)

set "device_connected="
for /f "skip=1 tokens=1,2" %%A in ('adb devices') do (
    if "%%B"=="device" set "device_connected=1"
)

if not defined device_connected (
    powershell.exe -NoProfile -Command "$msg = -join ([char[]](0x6CA1,0x6709,0x68C0,0x6D4B,0x5230,0x5DF2,0x8FDE,0x63A5,0x7684,0x20,0x41,0x6E64,0x72,0x6F,0x69,0x64,0x20,0x624B,0x673A,0xFF0C,0x8BF7,0x5148,0x8FDE,0x63A5,0x624B,0x673A,0x5E76,0x786E,0x8BA4,0x20,0x61,0x64,0x62,0x20,0x64,0x65,0x76,0x69,0x63,0x65,0x73,0x20,0x80FD,0x770B,0x5230,0x20,0x64,0x65,0x76,0x69,0x63,0x65,0x20,0x72B6,0x6001,0x3002)); $title = -join ([char[]](0x8BBE,0x5907,0x672A,0x8FDE,0x63A5)); Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.MessageBox]::Show($msg, $title, [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)" 1>nul
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0qq_music_exchange_ui.ps1" %*
exit /b %errorlevel%
