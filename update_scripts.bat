@echo off
chcp 65001 >nul
echo ============================================
echo   QQ音乐脚本 - 一键更新（从GitHub拉取最新版）
echo ============================================
echo.

:: 切换到脚本所在目录
cd /d "%~dp0"

:: 检查 git 是否可用
git --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Git，请先安装 Git for Windows
    echo 下载地址: https://git-scm.com/download/win
    pause
    exit /b 1
)

echo [1/3] 检查当前版本...
git log --oneline -1 2>&1
echo.

echo [2/3] 从 GitHub 拉取最新脚本...
git pull origin main
if errorlevel 1 (
    echo.
    echo [错误] 拉取失败，请检查网络连接或仓库权限
    pause
    exit /b 1
)
echo.

echo [3/3] 更新完成！当前版本：
git log --oneline -1 2>&1
echo.
echo ============================================
echo   所有脚本已更新到最新版本
echo ============================================
echo.
pause
