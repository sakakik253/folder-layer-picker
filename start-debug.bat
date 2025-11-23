@echo off
chcp 65001 > nul
echo =====================================
echo  Folder Layer Picker デバッグ起動
echo =====================================
echo.
echo このウィンドウはアプリケーション終了後も残ります。
echo エラーメッセージを確認できます。
echo.
pause

REM PowerShell 7 (pwsh) が利用可能かチェック
where pwsh >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [PowerShell 7を使用（デバッグモード）]
    pwsh -ExecutionPolicy Bypass -NoProfile -NoExit -File "%~dp0folder-layer-picker.ps1"
    goto :end
)

REM Windows PowerShell を使用
where powershell >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [Windows PowerShellを使用（デバッグモード）]
    powershell -ExecutionPolicy Bypass -NoProfile -NoExit -File "%~dp0folder-layer-picker.ps1"
    goto :end
)

echo.
echo [エラー] PowerShellが見つかりませんでした。
pause
exit /b 1

:end
echo.
echo =====================================
echo デバッグセッション終了
echo =====================================
pause

