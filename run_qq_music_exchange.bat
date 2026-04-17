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

:: 设备检测交给 UI 自己处理，不在 BAT 层强制退出
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0qq_music_exchange_ui.ps1" %*
exit /b %errorlevel%
