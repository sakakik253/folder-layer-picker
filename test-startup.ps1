# 起動テストスクリプト
Write-Host "=== Folder Layer Picker 起動診断 ===" -ForegroundColor Cyan

# 1. PowerShellバージョンチェック
Write-Host "`n[1] PowerShellバージョン:" -ForegroundColor Yellow
$PSVersionTable.PSVersion | Format-Table

# 2. XAMLファイル存在チェック
Write-Host "[2] XAMLファイルチェック:" -ForegroundColor Yellow
$xamlPath = Join-Path $PSScriptRoot "MainWindow.xaml"
if (Test-Path $xamlPath) {
    Write-Host "  ✓ MainWindow.xaml が見つかりました" -ForegroundColor Green
    Write-Host "  パス: $xamlPath" -ForegroundColor Gray
} else {
    Write-Host "  ✗ MainWindow.xaml が見つかりません！" -ForegroundColor Red
    Write-Host "  パス: $xamlPath" -ForegroundColor Gray
    Read-Host "`nEnterキーを押して終了"
    exit 1
}

# 3. WPFアセンブリ読み込みチェック
Write-Host "`n[3] WPFアセンブリ読み込みチェック:" -ForegroundColor Yellow
try {
    Add-Type -AssemblyName PresentationFramework
    Write-Host "  ✓ PresentationFramework 読み込み成功" -ForegroundColor Green
    
    Add-Type -AssemblyName PresentationCore
    Write-Host "  ✓ PresentationCore 読み込み成功" -ForegroundColor Green
    
    Add-Type -AssemblyName WindowsBase
    Write-Host "  ✓ WindowsBase 読み込み成功" -ForegroundColor Green
}
catch {
    Write-Host "  ✗ WPFアセンブリの読み込みエラー: $_" -ForegroundColor Red
    Read-Host "`nEnterキーを押して終了"
    exit 1
}

# 4. XAML読み込みチェック
Write-Host "`n[4] XAML読み込みチェック:" -ForegroundColor Yellow
try {
    [xml]$xaml = Get-Content $xamlPath -Encoding UTF8
    Write-Host "  ✓ XAML解析成功" -ForegroundColor Green
    
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    Write-Host "  ✓ WPFウィンドウ作成成功" -ForegroundColor Green
}
catch {
    Write-Host "  ✗ XAML読み込みエラー:" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`n詳細:" -ForegroundColor Yellow
    Write-Host $_.Exception.ToString() -ForegroundColor Gray
    Read-Host "`nEnterキーを押して終了"
    exit 1
}

# 5. UI要素アクセステスト
Write-Host "`n[5] UI要素アクセステスト:" -ForegroundColor Yellow
try {
    $testElements = @(
        "txtFolderPath",
        "btnBrowse",
        "btnAnalyze",
        "pnlLayerSelection",
        "btnUpdatePreview",
        "btnExecute",
        "btnUndo",
        "btnClose"
    )
    
    $foundCount = 0
    foreach ($elementName in $testElements) {
        $element = $window.FindName($elementName)
        if ($element) {
            $foundCount++
        } else {
            Write-Host "  ⚠ 要素 '$elementName' が見つかりません" -ForegroundColor Yellow
        }
    }
    
    Write-Host "  ✓ $foundCount/$($testElements.Count) 個の要素を確認" -ForegroundColor Green
}
catch {
    Write-Host "  ✗ UI要素アクセスエラー: $_" -ForegroundColor Red
}

# 6. テスト起動
Write-Host "`n[6] テスト起動:" -ForegroundColor Yellow
Write-Host "  ウィンドウを表示します（閉じるボタンでテスト完了）..." -ForegroundColor Cyan

try {
    # 閉じるボタンだけ機能させる
    $btnClose = $window.FindName("btnClose")
    if ($btnClose) {
        $btnClose.Add_Click({
            $window.Close()
        })
    }
    
    $result = $window.ShowDialog()
    Write-Host "`n  ✓ ウィンドウ表示成功！" -ForegroundColor Green
}
catch {
    Write-Host "`n  ✗ ウィンドウ表示エラー:" -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "`nEnterキーを押して終了"
    exit 1
}

Write-Host "`n=== 診断完了 ===" -ForegroundColor Cyan
Write-Host "すべてのチェックが成功しました。" -ForegroundColor Green
Write-Host "メインスクリプト (folder-layer-picker.ps1) が起動するはずです。" -ForegroundColor Green

Read-Host "`nEnterキーを押して終了"

