@echo off
chcp 65001 > nul
echo =====================================
echo  起動診断ツール
echo =====================================
echo.

REM PowerShell 7 (pwsh) が利用可能かチェック
where pwsh >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [PowerShell 7を使用]
    pwsh -ExecutionPolicy Bypass -NoProfile -File "%~dp0test-startup.ps1"
    goto :end
)

REM Windows PowerShell を使用
where powershell >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [Windows PowerShellを使用]
    powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0test-startup.ps1"
    goto :end
)

echo.
echo [エラー] PowerShellが見つかりませんでした。
pause
exit /b 1

:end

