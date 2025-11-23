@echo off
chcp 65001 > nul
echo =====================================
echo  Folder Layer Picker 起動中...
echo =====================================
echo.

REM PowerShell 7 (pwsh) が利用可能かチェック
where pwsh >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [PowerShell 7を使用して起動します]
    pwsh -ExecutionPolicy Bypass -NoProfile -File "%~dp0folder-layer-picker.ps1"
    goto :end
)

REM Windows PowerShell (powershell) を使用
where powershell >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [Windows PowerShellを使用して起動します]
    powershell -ExecutionPolicy Bypass -NoProfile -File "%~dp0folder-layer-picker.ps1"
    goto :end
)

REM PowerShellが見つからない場合
echo.
echo [エラー] PowerShellが見つかりませんでした。
echo Windows 10/11には標準でインストールされています。
echo.
pause
exit /b 1

:end
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo =====================================
    echo  エラーが発生しました
    echo =====================================
    echo.
    echo 終了コード: %ERRORLEVEL%
    echo.
    pause
)

