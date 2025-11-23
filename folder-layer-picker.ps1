#Requires -Version 5.1

<#
.SYNOPSIS
    Folder Layer Picker - 深い階層構造のフォルダを整理するツール

.DESCRIPTION
    指定した階層のフォルダをルート直下に移動し、空フォルダを自動削除します。
    WPF UIを使用した対話型ツールです。

.NOTES
    Author: Folder Layer Picker
    Version: 1.0
#>

# ================================================================================
# グローバル変数
# ================================================================================

$script:RootPath = ""                    # 対象フォルダのパス
$script:HierarchyData = @{}              # 階層構造データ
$script:SelectedLayers = @()             # 選択された階層（複数）
$script:OperationMode = "MoveAndDeleteAll"  # 操作モード
$script:DeleteRange = "AllEmpty"         # 削除範囲
$script:MoveDestination = "Root"         # 移動先（Root or Parent）
$script:PreviewData = $null              # プレビューデータ
$script:LastBackupPath = ""              # 最後に作成したバックアップのパス
$script:LogFilePath = ""                 # ログファイルのパス
$script:AutoPreviewEnabled = $false      # 自動プレビュー更新が有効か

# ================================================================================
# ログ機能
# ================================================================================

function Write-Log {
    <#
    .SYNOPSIS
        ログファイルにメッセージを書き込む
    #>
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # コンソールにも出力
    switch ($Level) {
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logMessage -ForegroundColor Green }
        default   { Write-Host $logMessage }
    }
    
    # ログファイルに書き込み
    if ($script:LogFilePath) {
        try {
            Add-Content -Path $script:LogFilePath -Value $logMessage -Encoding UTF8
        }
        catch {
            Write-Host "ログ書き込みエラー: $_" -ForegroundColor Red
        }
    }
}

function Initialize-LogFile {
    <#
    .SYNOPSIS
        ログファイルを初期化する
    #>
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogFilePath = Join-Path $PSScriptRoot "folder-layer-picker_$timestamp.log"
    
    try {
        "=" * 80 | Out-File -FilePath $script:LogFilePath -Encoding UTF8
        "Folder Layer Picker ログファイル" | Out-File -FilePath $script:LogFilePath -Append -Encoding UTF8
        "作成日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" | Out-File -FilePath $script:LogFilePath -Append -Encoding UTF8
        "=" * 80 | Out-File -FilePath $script:LogFilePath -Append -Encoding UTF8
        Write-Log "ログファイルを初期化しました: $script:LogFilePath"
    }
    catch {
        Write-Host "ログファイル初期化エラー: $_" -ForegroundColor Red
    }
}

# ================================================================================
# エクスポート機能
# ================================================================================

function Get-FolderDetails {
    <#
    .SYNOPSIS
        フォルダの詳細情報を取得する
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )
    
    try {
        $files = Get-ChildItem -LiteralPath $FolderPath -File -ErrorAction SilentlyContinue
        $subfolders = Get-ChildItem -LiteralPath $FolderPath -Directory -ErrorAction SilentlyContinue
        
        $fileCount = $files.Count
        $subfolderCount = $subfolders.Count
        
        # ファイルサイズの合計
        $totalSize = ($files | Measure-Object -Property Length -Sum).Sum
        if (-not $totalSize) { $totalSize = 0 }
        
        # 最終更新日時
        $lastModified = $null
        if ($files.Count -gt 0) {
            $lastModified = ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
        }
        if ($subfolders.Count -gt 0) {
            $folderModified = ($subfolders | Sort-Object LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
            if (-not $lastModified -or ($folderModified -gt $lastModified)) {
                $lastModified = $folderModified
            }
        }
        
        # 拡張子別ファイル数
        $extensionStats = @{}
        foreach ($file in $files) {
            $ext = $file.Extension.ToLower()
            if (-not $ext) { $ext = "(なし)" }
            if (-not $extensionStats.ContainsKey($ext)) {
                $extensionStats[$ext] = 0
            }
            $extensionStats[$ext]++
        }
        
        return @{
            FileCount = $fileCount
            SubfolderCount = $subfolderCount
            TotalSize = $totalSize
            LastModified = $lastModified
            ExtensionStats = $extensionStats
            IsEmpty = ($fileCount -eq 0 -and $subfolderCount -eq 0)
        }
    }
    catch {
        return @{
            FileCount = 0
            SubfolderCount = 0
            TotalSize = 0
            LastModified = $null
            ExtensionStats = @{}
            IsEmpty = $true
        }
    }
}

function Export-AnalysisToText {
    <#
    .SYNOPSIS
        分析結果をテキスト形式でエクスポート
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$RootPath,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$HierarchyData,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    
    Write-Log "テキスト形式でエクスポート開始: $OutputPath"
    
    try {
        $output = @()
        
        # ヘッダー
        $output += "=" * 80
        $output += "フォルダ階層構造 分析結果"
        $output += "対象フォルダ: $RootPath"
        $output += "分析日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        
        # 統計情報の計算
        $totalFolders = ($HierarchyData.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
        $maxLevel = ($HierarchyData.Keys | Measure-Object -Maximum).Maximum
        
        # ルート配下の全ファイル数とサイズ
        $allFiles = Get-ChildItem -LiteralPath $RootPath -File -Recurse -ErrorAction SilentlyContinue
        $totalFiles = $allFiles.Count
        $totalSize = ($allFiles | Measure-Object -Property Length -Sum).Sum
        if (-not $totalSize) { $totalSize = 0 }
        $totalSizeGB = [math]::Round($totalSize / 1GB, 2)
        
        # 空フォルダ数
        $emptyFolders = 0
        Get-ChildItem -LiteralPath $RootPath -Directory -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            $items = Get-ChildItem -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            if ($items.Count -eq 0) { $emptyFolders++ }
        }
        
        $output += "最大階層: $maxLevel 階層"
        $output += "総フォルダ数: $totalFolders 個"
        $output += "総ファイル数: $totalFiles 個"
        $output += "総サイズ: $totalSizeGB GB"
        $output += "空フォルダ数: $emptyFolders 個"
        $output += ""
        $output += "階層別統計:"
        
        $sortedLevels = $HierarchyData.Keys | Sort-Object
        foreach ($level in $sortedLevels) {
            $count = $HierarchyData[$level].Count
            $levelFiles = 0
            foreach ($folder in $HierarchyData[$level].Folders) {
                $details = Get-FolderDetails -FolderPath $folder.FullPath
                $levelFiles += $details.FileCount
            }
            $output += "  $level 階層目: $count フォルダ ($levelFiles ファイル)"
        }
        
        $output += "=" * 80
        $output += "詳細構造"
        $output += ""
        
        # ツリー構造を生成
        $rootFolders = Get-ChildItem -LiteralPath $RootPath -Directory -ErrorAction SilentlyContinue
        foreach ($folder in $rootFolders) {
            $output += Build-FolderTree -FolderPath $folder.FullName -Prefix "" -IsLast $false -Level 1
        }
        
        # ファイルに出力（BOM付きUTF-8）
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllLines($OutputPath, $output, $utf8WithBom)
        
        Write-Log "テキストエクスポート完了: $OutputPath" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "テキストエクスポートエラー: $_" -Level ERROR
        return $false
    }
}

function Build-FolderTree {
    <#
    .SYNOPSIS
        フォルダツリー構造を再帰的に構築
    #>
    param(
        [string]$FolderPath,
        [string]$Prefix,
        [bool]$IsLast,
        [int]$Level
    )
    
    $output = @()
    
    try {
        $folderName = Split-Path $FolderPath -Leaf
        $details = Get-FolderDetails -FolderPath $FolderPath
        
        # サイズをMB単位で表示
        $sizeMB = [math]::Round($details.TotalSize / 1MB, 2)
        
        # 拡張子統計
        $extStr = ""
        if ($details.ExtensionStats.Count -gt 0) {
            $extArray = $details.ExtensionStats.GetEnumerator() | ForEach-Object { "$($_.Key)($($_.Value))" }
            $extStr = "拡張子: " + ($extArray -join ", ")
        }
        
        # 最終更新日時
        $modifiedStr = if ($details.LastModified) { $details.LastModified.ToString("yyyy-MM-dd") } else { "不明" }
        
        # ツリー記号
        $connector = if ($IsLast) { "└──" } else { "├──" }
        
        # フォルダ情報行
        $info = "$folderName\ (ファイル: $($details.FileCount), サイズ: $sizeMB MB, 更新: $modifiedStr)"
        $output += "$Prefix$connector $info"
        
        # 拡張子情報（ファイルがある場合）
        if ($extStr) {
            $extPrefix = if ($IsLast) { "    " } else { "│   " }
            $output += "$Prefix$extPrefix    $extStr"
        }
        
        # サブフォルダを処理
        $subfolders = Get-ChildItem -LiteralPath $FolderPath -Directory -ErrorAction SilentlyContinue
        if ($subfolders.Count -gt 0) {
            $newPrefix = $Prefix + (if ($IsLast) { "    " } else { "│   " })
            
            for ($i = 0; $i -lt $subfolders.Count; $i++) {
                $isLastSub = ($i -eq ($subfolders.Count - 1))
                $output += Build-FolderTree -FolderPath $subfolders[$i].FullName -Prefix $newPrefix -IsLast $isLastSub -Level ($Level + 1)
            }
        }
    }
    catch {
        $output += "$Prefix$connector [エラー: アクセスできません]"
    }
    
    return $output
}

function Export-AnalysisToCSV {
    <#
    .SYNOPSIS
        分析結果をCSV形式でエクスポート
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$RootPath,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$HierarchyData,
        
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )
    
    Write-Log "CSV形式でエクスポート開始: $OutputPath"
    
    try {
        $csvData = @()
        
        # ヘッダー
        $csvData += [PSCustomObject]@{
            "階層レベル" = "階層レベル"
            "フォルダ名" = "フォルダ名"
            "フルパス" = "フルパス"
            "ファイル数" = "ファイル数"
            "サブフォルダ数" = "サブフォルダ数"
            "フォルダサイズ(MB)" = "フォルダサイズ(MB)"
            "最終更新日時" = "最終更新日時"
            "空フォルダ" = "空フォルダ"
            "拡張子別ファイル数" = "拡張子別ファイル数"
        }
        
        # 階層順にデータを構築
        $sortedLevels = $HierarchyData.Keys | Sort-Object
        foreach ($level in $sortedLevels) {
            foreach ($folder in $HierarchyData[$level].Folders) {
                $details = Get-FolderDetails -FolderPath $folder.FullPath
                
                # インデント（階層-1 × 2文字）
                $indent = "  " * ($level - 1)
                $folderName = $indent + $folder.Name
                
                # サイズをMB単位
                $sizeMB = [math]::Round($details.TotalSize / 1MB, 2)
                
                # 拡張子統計
                $extStr = ""
                if ($details.ExtensionStats.Count -gt 0) {
                    $extArray = $details.ExtensionStats.GetEnumerator() | ForEach-Object { "$($_.Key)($($_.Value))" }
                    $extStr = $extArray -join ", "
                }
                
                # 最終更新日時
                $modifiedStr = if ($details.LastModified) { 
                    $details.LastModified.ToString("yyyy-MM-dd HH:mm:ss") 
                } else { 
                    "" 
                }
                
                # 空フォルダ判定
                $isEmpty = if ($details.IsEmpty) { "はい" } else { "いいえ" }
                
                $csvData += [PSCustomObject]@{
                    "階層レベル" = $level
                    "フォルダ名" = $folderName
                    "フルパス" = $folder.FullPath
                    "ファイル数" = $details.FileCount
                    "サブフォルダ数" = $details.SubfolderCount
                    "フォルダサイズ(MB)" = $sizeMB
                    "最終更新日時" = $modifiedStr
                    "空フォルダ" = $isEmpty
                    "拡張子別ファイル数" = $extStr
                }
            }
        }
        
        # CSVに出力（BOM付きUTF-8でExcel対応）
        $csvContent = $csvData | ConvertTo-Csv -NoTypeInformation
        $utf8WithBom = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllLines($OutputPath, $csvContent, $utf8WithBom)
        
        Write-Log "CSVエクスポート完了: $OutputPath" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "CSVエクスポートエラー: $_" -Level ERROR
        return $false
    }
}

# ================================================================================
# 階層分析機能
# ================================================================================

function Get-FolderHierarchy {
    <#
    .SYNOPSIS
        フォルダ構造をスキャンして階層別に分類する
    
    .PARAMETER RootPath
        分析するルートフォルダのパス
    
    .OUTPUTS
        階層別のフォルダ情報を含むハッシュテーブル
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$RootPath
    )
    
    Write-Log "フォルダ階層分析を開始: $RootPath"
    
    $hierarchyData = @{}
    
    try {
        # ルートフォルダ自体は階層0とする
        $rootItem = Get-Item -LiteralPath $RootPath -ErrorAction Stop
        
        # すべてのサブフォルダを取得
        $allFolders = Get-ChildItem -LiteralPath $RootPath -Directory -Recurse -ErrorAction SilentlyContinue
        
        foreach ($folder in $allFolders) {
            # 相対パスを取得
            $relativePath = $folder.FullName.Substring($RootPath.Length).TrimStart('\')
            
            # 階層レベルを計算（バックスラッシュの数 + 1）
            $level = ($relativePath.Split('\').Count)
            
            # 階層データに追加
            if (-not $hierarchyData.ContainsKey($level)) {
                $hierarchyData[$level] = @{
                    Count = 0
                    Folders = @()
                }
            }
            
            $hierarchyData[$level].Count++
            $hierarchyData[$level].Folders += @{
                FullPath = $folder.FullName
                RelativePath = $relativePath
                Name = $folder.Name
                Parent = Split-Path $relativePath -Parent
            }
        }
        
        Write-Log "階層分析完了: $($hierarchyData.Keys.Count) 階層、合計 $($allFolders.Count) フォルダ" -Level SUCCESS
        
        return $hierarchyData
    }
    catch {
        Write-Log "階層分析エラー: $_" -Level ERROR
        return @{}
    }
}

# ================================================================================
# プレビュー機能
# ================================================================================

function Get-PreviewData {
    <#
    .SYNOPSIS
        移動予測データを生成する（複数階層対応）
    
    .PARAMETER HierarchyData
        階層構造データ
    
    .PARAMETER TargetLevels
        移動対象の階層レベル（配列）
    
    .PARAMETER RootPath
        ルートフォルダのパス
    
    .PARAMETER OperationMode
        操作モード
    
    .PARAMETER DeleteRange
        削除範囲
    
    .PARAMETER MoveDestination
        移動先（Root or Parent）
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$HierarchyData,
        
        [Parameter(Mandatory=$true)]
        [array]$TargetLevels,
        
        [Parameter(Mandatory=$true)]
        [string]$RootPath,
        
        [string]$OperationMode = "MoveAndDeleteAll",
        
        [string]$DeleteRange = "AllEmpty",
        
        [string]$MoveDestination = "Root"
    )
    
    $levelList = $TargetLevels -join ", "
    Write-Log "プレビューデータを生成: 階層$levelList (モード: $OperationMode)"
    
    $previewData = @{
        TargetFolders = @()
        MoveOperations = @()
        EmptyFoldersToDelete = @()
        Warnings = @()
        OperationMode = $OperationMode
    }
    
    try {
        # 削除のみモードの場合、移動操作は実行しない
        if ($OperationMode -eq "DeleteOnly") {
            Write-Log "削除のみモード: 選択した階層のフォルダを削除"
            
            # 選択した階層のフォルダを削除対象として記録
            $foldersToDelete = @()
            foreach ($targetLevel in $TargetLevels) {
                if ($HierarchyData.ContainsKey($targetLevel)) {
                    foreach ($folder in $HierarchyData[$targetLevel].Folders) {
                        $foldersToDelete += @{
                            Path = $folder.FullPath
                            Level = $targetLevel
                            Name = $folder.Name
                        }
                    }
                }
            }
            
            $previewData.EmptyFoldersToDelete = $foldersToDelete
            $previewData.Warnings += "削除のみモード: $($foldersToDelete.Count) 個のフォルダが削除されます"
            
            Write-Log "削除のみモード: $($foldersToDelete.Count) 個のフォルダを削除予定" -Level SUCCESS
            return $previewData
        }
        
        # 対象階層のフォルダを取得（複数階層対応）
        $allTargetFolders = @()
        foreach ($targetLevel in $TargetLevels) {
            if ($HierarchyData.ContainsKey($targetLevel)) {
                $allTargetFolders += $HierarchyData[$targetLevel].Folders
                Write-Log "階層$targetLevel から $($HierarchyData[$targetLevel].Count) 個のフォルダを取得"
            }
        }
        
        if ($allTargetFolders.Count -eq 0) {
            Write-Log "対象フォルダが見つかりませんでした" -Level WARNING
            return $previewData
        }
        
        $targetFolders = $allTargetFolders
        
        # 移動先に応じた処理
        if ($MoveDestination -eq "Parent") {
            # 1階層上に移動
            Write-Log "移動先: 1階層上"
            
            foreach ($folder in $targetFolders) {
                $originalName = $folder.Name
                $currentParent = Split-Path $folder.FullPath -Parent
                
                # 親の親（1階層上）を取得
                $targetParent = Split-Path $currentParent -Parent
                
                # ルート直下の場合はスキップ
                if (-not $targetParent -or $targetParent -eq $RootPath) {
                    $previewData.Warnings += "スキップ: '$originalName' は既に最上位階層です"
                    continue
                }
                
                # 移動先の既存フォルダ名を取得
                $existingInTarget = @{}
                Get-ChildItem -LiteralPath $targetParent -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                    $existingInTarget[$_.Name.ToLower()] = $true
                }
                
                # 同名チェック
                $finalName = $originalName
                $counter = 1
                while ($existingInTarget.ContainsKey($finalName.ToLower())) {
                    $finalName = "${originalName}_$counter"
                    $counter++
                }
                
                # 移動操作を記録
                $moveOp = @{
                    SourcePath = $folder.FullPath
                    DestinationName = $finalName
                    DestinationPath = Join-Path $targetParent $finalName
                    OriginalName = $originalName
                    WasRenamed = ($finalName -ne $originalName)
                    TargetParent = $targetParent
                }
                
                $previewData.MoveOperations += $moveOp
                
                if ($moveOp.WasRenamed) {
                    $previewData.Warnings += "名前変更: '$originalName' → '$finalName'"
                }
            }
        }
        else {
            # ルート直下に移動（デフォルト）
            Write-Log "移動先: ルート直下"
            
            # ルート直下の既存フォルダ名を取得
            $existingNames = @{}
            Get-ChildItem -LiteralPath $RootPath -Directory | ForEach-Object {
                $existingNames[$_.Name.ToLower()] = $true
            }
            
            foreach ($folder in $targetFolders) {
                $originalName = $folder.Name
                $newName = $originalName
                
                # 親パスがある場合（ルート直下でない場合）
                if ($folder.Parent) {
                    # 親パスをプレフィックスとして追加
                    $parentPrefix = $folder.Parent -replace '\\', '_'
                    $newName = "${parentPrefix}_${originalName}"
                }
                
                # 同名チェック
                $finalName = $newName
                $counter = 1
                while ($existingNames.ContainsKey($finalName.ToLower())) {
                    $finalName = "${newName}_$counter"
                    $counter++
                }
                
                # 移動操作を記録
                $moveOp = @{
                    SourcePath = $folder.FullPath
                    DestinationName = $finalName
                    DestinationPath = Join-Path $RootPath $finalName
                    OriginalName = $originalName
                    WasRenamed = ($finalName -ne $originalName)
                }
                
                $previewData.MoveOperations += $moveOp
                $existingNames[$finalName.ToLower()] = $true
                
                if ($moveOp.WasRenamed) {
                    $previewData.Warnings += "名前変更: '$originalName' → '$finalName'"
                }
            }
        }
        
        # 移動後に空になるフォルダを予測
        $allAffectedPaths = @{}
        foreach ($moveOp in $previewData.MoveOperations) {
            $parentPath = Split-Path $moveOp.SourcePath -Parent
            while ($parentPath -and $parentPath -ne $RootPath) {
                $allAffectedPaths[$parentPath] = $true
                $parentPath = Split-Path $parentPath -Parent
            }
        }
            
        # 各フォルダについて、移動後も内容が残るかチェック
        foreach ($path in $allAffectedPaths.Keys) {
            $willBeEmpty = Test-WillBeEmptyAfterMove -FolderPath $path -MoveOperations $previewData.MoveOperations
            if ($willBeEmpty) {
                $previewData.EmptyFoldersToDelete += $path
            }
        }
        
        Write-Log "プレビュー生成完了: $($previewData.MoveOperations.Count) 個のフォルダを移動予定" -Level SUCCESS
        
        return $previewData
    }
    catch {
        Write-Log "プレビュー生成エラー: $_" -Level ERROR
        return $null
    }
}

function Test-WillBeEmptyAfterMove {
    <#
    .SYNOPSIS
        フォルダが移動後に空になるかをチェックする
    #>
    param(
        [string]$FolderPath,
        [array]$MoveOperations
    )
    
    try {
        $items = Get-ChildItem -LiteralPath $FolderPath -Force -ErrorAction Stop
        
        foreach ($item in $items) {
            # このアイテムが移動対象でない場合、フォルダは空にならない
            $isMoving = $false
            foreach ($moveOp in $MoveOperations) {
                if ($item.FullName -eq $moveOp.SourcePath) {
                    $isMoving = $true
                    break
                }
            }
            
            if (-not $isMoving) {
                # ファイルか、移動しないフォルダがある
                if (-not $item.PSIsContainer) {
                    return $false
                }
                
                # サブフォルダの場合、再帰的にチェック
                $subWillBeEmpty = Test-WillBeEmptyAfterMove -FolderPath $item.FullName -MoveOperations $MoveOperations
                if (-not $subWillBeEmpty) {
                    return $false
                }
            }
        }
        
        return $true
    }
    catch {
        return $false
    }
}

# ================================================================================
# 実行機能
# ================================================================================

function Move-MultipleLayers {
    <#
    .SYNOPSIS
        複数階層のフォルダを移動する（モード対応）
    #>
    param(
        [Parameter(Mandatory=$true)]
        [object]$PreviewData,
        
        [Parameter(Mandatory=$true)]
        [string]$RootPath,
        
        [string]$OperationMode = "MoveAndDeleteAll"
    )
    
    Write-Log "複数階層の移動を開始: モード=$OperationMode"
    
    if ($OperationMode -eq "DeleteOnly") {
        Write-Log "削除のみモード: 移動をスキップ"
        return @{ Success = 0; Error = 0 }
    }
    
    if ($PreviewData.MoveOperations.Count -eq 0) {
        Write-Log "移動対象のフォルダがありません"
        return @{ Success = 0; Error = 0 }
    }
    
    # Move-FoldersToRootを呼び出し
    return Move-FoldersToRoot -MoveOperations $PreviewData.MoveOperations -RootPath $RootPath
}

function Remove-SelectedLayers {
    <#
    .SYNOPSIS
        選択した階層のフォルダを削除する（削除のみモード用）
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$FoldersToDelete
    )
    
    Write-Log "選択階層のフォルダ削除を開始: $($FoldersToDelete.Count) 個"
    
    $deletedCount = 0
    $errorCount = 0
    
    foreach ($folder in $FoldersToDelete) {
        try {
            if (Test-Path -LiteralPath $folder.Path) {
                Remove-Item -LiteralPath $folder.Path -Recurse -Force -ErrorAction Stop
                Write-Log "削除: $($folder.Path)" -Level SUCCESS
                $deletedCount++
            }
            else {
                Write-Log "スキップ（存在しない）: $($folder.Path)" -Level WARNING
            }
        }
        catch {
            Write-Log "削除失敗: $($folder.Path) - $_" -Level ERROR
            $errorCount++
        }
    }
    
    Write-Log "選択階層削除完了: 成功 $deletedCount 個、失敗 $errorCount 個" -Level SUCCESS
    return $deletedCount
}

function Move-FoldersToRoot {
    <#
    .SYNOPSIS
        指定された階層のフォルダをルート直下に移動する
    #>
    param(
        [Parameter(Mandatory=$true)]
        [array]$MoveOperations,
        
        [Parameter(Mandatory=$true)]
        [string]$RootPath
    )
    
    Write-Log "フォルダ移動を開始: $($MoveOperations.Count) 個のフォルダ"
    
    $successCount = 0
    $errorCount = 0
    
    foreach ($moveOp in $MoveOperations) {
        try {
            Write-Log "移動: $($moveOp.SourcePath) → $($moveOp.DestinationPath)"
            
            # 移動実行
            Move-Item -LiteralPath $moveOp.SourcePath -Destination $moveOp.DestinationPath -ErrorAction Stop
            
            $successCount++
            Write-Log "移動成功: $($moveOp.DestinationName)" -Level SUCCESS
        }
        catch {
            $errorCount++
            Write-Log "移動失敗: $($moveOp.SourcePath) - $_" -Level ERROR
        }
    }
    
    Write-Log "フォルダ移動完了: 成功 $successCount 個、失敗 $errorCount 個" -Level SUCCESS
    
    return @{
        Success = $successCount
        Error = $errorCount
    }
}

function Remove-EmptyFolders {
    <#
    .SYNOPSIS
        空のフォルダを再帰的に削除する（モード対応）
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$RootPath,
        
        [string]$DeleteRange = "AllEmpty",
        
        [array]$SelectedLevels = @()
    )
    
    Write-Log "空フォルダの削除を開始: 削除範囲=$DeleteRange"
    
    # NoDeleteモードの場合はスキップ
    if ($DeleteRange -eq "NoDelete") {
        Write-Log "削除しない設定: スキップ"
        return 0
    }
    
    $deletedCount = 0
    
    try {
        # ボトムアップで削除するため、深い階層から処理
        do {
            $deleted = $false
            $allFolders = Get-ChildItem -LiteralPath $RootPath -Directory -Recurse -ErrorAction SilentlyContinue
            
            # 深さでソート（深い順）
            $sortedFolders = $allFolders | Sort-Object { $_.FullName.Split('\').Count } -Descending
            
            foreach ($folder in $sortedFolders) {
                try {
                    $items = Get-ChildItem -LiteralPath $folder.FullName -Force -ErrorAction Stop
                    
                    if ($items.Count -eq 0) {
                        Remove-Item -LiteralPath $folder.FullName -Force -ErrorAction Stop
                        Write-Log "削除: $($folder.FullName)"
                        $deletedCount++
                        $deleted = $true
                    }
                }
                catch {
                    Write-Log "削除失敗: $($folder.FullName) - $_" -Level WARNING
                }
            }
        } while ($deleted)
        
        Write-Log "空フォルダ削除完了: $deletedCount 個のフォルダを削除" -Level SUCCESS
    }
    catch {
        Write-Log "空フォルダ削除エラー: $_" -Level ERROR
    }
    
    return $deletedCount
}

# ================================================================================
# バックアップ機能
# ================================================================================

function New-Backup {
    <#
    .SYNOPSIS
        対象フォルダのバックアップを作成する
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SourcePath
    )
    
    Write-Log "バックアップを作成中..."
    
    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $sourceName = Split-Path $SourcePath -Leaf
        $parentPath = Split-Path $SourcePath -Parent
        $backupName = "${sourceName}_backup_${timestamp}"
        $backupPath = Join-Path $parentPath $backupName
        
        Write-Log "バックアップ先: $backupPath"
        
        # フォルダをコピー
        Copy-Item -LiteralPath $SourcePath -Destination $backupPath -Recurse -ErrorAction Stop
        
        $script:LastBackupPath = $backupPath
        Write-Log "バックアップ作成完了: $backupPath" -Level SUCCESS
        
        return $backupPath
    }
    catch {
        Write-Log "バックアップ作成エラー: $_" -Level ERROR
        return $null
    }
}

function Restore-FromBackup {
    <#
    .SYNOPSIS
        バックアップから復元する
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$BackupPath,
        
        [Parameter(Mandatory=$true)]
        [string]$DestinationPath
    )
    
    Write-Log "バックアップから復元中: $BackupPath"
    
    try {
        # 現在のフォルダを削除
        if (Test-Path $DestinationPath) {
            Write-Log "現在のフォルダを削除: $DestinationPath"
            Remove-Item -LiteralPath $DestinationPath -Recurse -Force -ErrorAction Stop
        }
        
        # バックアップから復元
        Write-Log "バックアップをコピー: $BackupPath → $DestinationPath"
        Copy-Item -LiteralPath $BackupPath -Destination $DestinationPath -Recurse -ErrorAction Stop
        
        Write-Log "復元完了" -Level SUCCESS
        return $true
    }
    catch {
        Write-Log "復元エラー: $_" -Level ERROR
        return $false
    }
}

function Get-LatestBackup {
    <#
    .SYNOPSIS
        最新のバックアップフォルダを検索する
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$OriginalPath
    )
    
    try {
        $sourceName = Split-Path $OriginalPath -Leaf
        $parentPath = Split-Path $OriginalPath -Parent
        
        # バックアップフォルダを検索
        $backupFolders = Get-ChildItem -Path $parentPath -Directory -Filter "${sourceName}_backup_*" -ErrorAction Stop
        
        if ($backupFolders.Count -eq 0) {
            return $null
        }
        
        # 最新のバックアップを返す
        $latest = $backupFolders | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        return $latest.FullName
    }
    catch {
        Write-Log "バックアップ検索エラー: $_" -Level ERROR
        return $null
    }
}

# ================================================================================
# UI イベントハンドラー
# ================================================================================

function Show-FolderBrowserDialog {
    <#
    .SYNOPSIS
        フォルダ選択ダイアログを表示する
    #>
    Add-Type -AssemblyName System.Windows.Forms
    
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "対象フォルダを選択してください"
    $folderBrowser.ShowNewFolderButton = $false
    
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    }
    
    return $null
}

function Show-LayerAnalysis {
    <#
    .SYNOPSIS
        階層分析結果をUIに表示する（チェックボックスで複数選択対応）
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$HierarchyData,
        
        [Parameter(Mandatory=$true)]
        [System.Windows.Controls.StackPanel]$Panel
    )
    
    $Panel.Children.Clear()
    
    $sortedLevels = $HierarchyData.Keys | Sort-Object
    
    foreach ($level in $sortedLevels) {
        $count = $HierarchyData[$level].Count
        
        # ラジオボタンからチェックボックスに変更
        $checkBox = New-Object System.Windows.Controls.CheckBox
        $checkBox.Content = "階層 $level ($count 個のフォルダ)"
        $checkBox.Tag = $level
        $checkBox.Margin = "0,0,0,5"
        
        # イベントハンドラー：チェック状態変更時
        $checkBox.Add_Checked({
            param($sender, $e)
            Update-SelectedLayers
        })
        
        $checkBox.Add_Unchecked({
            param($sender, $e)
            Update-SelectedLayers
        })
        
        $Panel.Children.Add($checkBox) | Out-Null
    }
}

function Update-SelectedLayers {
    <#
    .SYNOPSIS
        選択された階層リストを更新する
    #>
    $script:SelectedLayers = @()
    
    $panel = $window.FindName("pnlMoveLayerSelection")
    foreach ($child in $panel.Children) {
        if ($child -is [System.Windows.Controls.CheckBox] -and $child.IsChecked) {
            $script:SelectedLayers += $child.Tag
        }
    }
    
    # プレビューボタンの有効/無効を切り替え
    $btnUpdatePreview = $window.FindName("btnUpdatePreview")
    $txtStatusBar = $window.FindName("txtStatusBar")
    
    if ($script:SelectedLayers.Count -gt 0) {
        $btnUpdatePreview.IsEnabled = $true
        $layerList = $script:SelectedLayers -join ", "
        $txtStatusBar.Text = "階層 $layerList を選択しました（$($script:SelectedLayers.Count) 個）"
        
        # 自動プレビュー更新
        if ($script:AutoPreviewEnabled) {
            Invoke-AutoPreviewUpdate
        }
    }
    else {
        $btnUpdatePreview.IsEnabled = $false
        $txtStatusBar.Text = "階層を選択してください"
    }
}

function Invoke-AutoPreviewUpdate {
    <#
    .SYNOPSIS
        自動的にプレビューを更新する
    #>
    if ($script:SelectedLayers.Count -eq 0 -or -not $script:HierarchyData -or $script:HierarchyData.Count -eq 0) {
        return
    }
    
    try {
        # プレビューデータ生成
        $script:PreviewData = Get-PreviewData `
            -HierarchyData $script:HierarchyData `
            -TargetLevels $script:SelectedLayers `
            -RootPath $script:RootPath `
            -OperationMode $script:OperationMode `
            -DeleteRange $script:DeleteRange `
            -MoveDestination $script:MoveDestination
        
        if ($script:PreviewData) {
            Update-PreviewDisplay -PreviewData $script:PreviewData
        }
    }
    catch {
        Write-Log "自動プレビュー更新エラー: $_" -Level ERROR
    }
}

function Update-PreviewDisplay {
    <#
    .SYNOPSIS
        プレビュー結果をUIに表示する（複数階層・モード対応）
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$PreviewData
    )
    
    $txtPreviewMoveCount = $window.FindName("txtPreviewMoveCount")
    $txtPreviewDeleteCount = $window.FindName("txtPreviewDeleteCount")
    $txtPreviewStatus = $window.FindName("txtPreviewStatus")
    $txtPreviewDetails = $window.FindName("txtPreviewDetails")
    $btnExecute = $window.FindName("btnExecute")
    
    $moveCount = $PreviewData.MoveOperations.Count
    $deleteCount = $PreviewData.EmptyFoldersToDelete.Count
    $operationMode = $PreviewData.OperationMode
    
    # モードに応じた表示
    switch ($operationMode) {
        "DeleteOnly" {
            $txtPreviewMoveCount.Text = "削除されるフォルダ数: $deleteCount"
            $txtPreviewDeleteCount.Text = "移動: なし（削除のみモード）"
        }
        "MoveOnly" {
            $txtPreviewMoveCount.Text = "移動されるフォルダ数: $moveCount"
            $txtPreviewDeleteCount.Text = "削除: なし（移動のみモード）"
        }
        default {
            $txtPreviewMoveCount.Text = "移動されるフォルダ数: $moveCount"
            $txtPreviewDeleteCount.Text = "削除される空フォルダ数: $deleteCount"
        }
    }
    
    $txtPreviewStatus.Text = "プレビュー更新完了 (モード: $operationMode)"
    
    # 詳細表示
    $details = @()
    $details += "=" * 80
    $details += "操作モード: $operationMode"
    $details += "選択階層: " + ($script:SelectedLayers -join ", ")
    $details += "=" * 80
    $details += ""
    
    # 移動予定の表示
    if ($moveCount -gt 0) {
        $details += "=" * 80
        $details += "移動予定のフォルダ ($moveCount 個)"
        $details += "=" * 80
        
        # 階層ごとにグループ化
        $byLevel = @{}
        foreach ($moveOp in $PreviewData.MoveOperations) {
            # 元のパスから階層を判定
            $depth = ($moveOp.SourcePath.Replace($script:RootPath, "").TrimStart('\').Split('\').Count) - 1
            if (-not $byLevel.ContainsKey($depth)) {
                $byLevel[$depth] = @()
            }
            $byLevel[$depth] += $moveOp
        }
        
        foreach ($level in ($byLevel.Keys | Sort-Object)) {
            $details += ""
            $details += "[階層 $level] $($byLevel[$level].Count) 個"
            foreach ($moveOp in $byLevel[$level]) {
                if ($moveOp.WasRenamed) {
                    $details += "  [名前変更] $($moveOp.OriginalName) → $($moveOp.DestinationName)"
                } else {
                    $details += "  $($moveOp.DestinationName)"
                }
            }
        }
    }
    
    # 削除予定の表示
    if ($deleteCount -gt 0) {
        $details += ""
        $details += "=" * 80
        
        if ($operationMode -eq "DeleteOnly") {
            $details += "削除予定のフォルダ ($deleteCount 個)"
        } else {
            $details += "削除予定の空フォルダ ($deleteCount 個)"
        }
        
        $details += "=" * 80
        
        foreach ($folder in $PreviewData.EmptyFoldersToDelete) {
            if ($folder -is [hashtable]) {
                $details += "  [階層 $($folder.Level)] $($folder.Name)"
            }
            else {
                $relativePath = $folder.Replace($script:RootPath, "").TrimStart('\')
                $details += "  $relativePath"
            }
        }
    }
    
    if ($PreviewData.Warnings.Count -gt 0) {
        $details += ""
        $details += "=" * 80
        $details += "警告"
        $details += "=" * 80
        $details += $PreviewData.Warnings
    }
    
    $txtPreviewDetails.Text = $details -join "`r`n"
    $btnExecute.IsEnabled = $true
}

# ================================================================================
# メイン処理
# ================================================================================

function Initialize-Application {
    <#
    .SYNOPSIS
        アプリケーションを初期化する
    #>
    
    # ログファイル初期化
    Initialize-LogFile
    
    Write-Log "Folder Layer Picker を起動しました"
}

# WPF ウィンドウの起動
try {
    # WPFアセンブリの読み込み（最初に実行）
    Write-Host "WPFアセンブリを読み込んでいます..." -ForegroundColor Cyan
    Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
    Add-Type -AssemblyName PresentationCore -ErrorAction Stop
    Add-Type -AssemblyName WindowsBase -ErrorAction Stop
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Write-Host "✓ WPFアセンブリ読み込み完了" -ForegroundColor Green
    
    # アプリケーション初期化
    Initialize-Application
    
    # XAMLファイルの読み込み
    $xamlPath = Join-Path $PSScriptRoot "MainWindow.xaml"
    Write-Host "XAMLファイルを読み込んでいます: $xamlPath" -ForegroundColor Cyan
    
    if (-not (Test-Path $xamlPath)) {
        throw "XAMLファイルが見つかりません: $xamlPath"
    }
    
    [xml]$xaml = Get-Content $xamlPath -Encoding UTF8 -ErrorAction Stop
    Write-Host "✓ XAML読み込み完了" -ForegroundColor Green
    
    # XAMLリーダーの作成
    Write-Host "WPFウィンドウを作成しています..." -ForegroundColor Cyan
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    if (-not $window) {
        throw "WPFウィンドウの作成に失敗しました"
    }
    Write-Host "✓ WPFウィンドウ作成完了" -ForegroundColor Green
    
    # UI要素の取得
    $txtFolderPath = $window.FindName("txtFolderPath")
    $btnBrowse = $window.FindName("btnBrowse")
    $btnAnalyze = $window.FindName("btnAnalyze")
    $btnExportAnalysis = $window.FindName("btnExportAnalysis")
    $txtAnalysisStatus = $window.FindName("txtAnalysisStatus")
    
    # 操作モード関連
    $rdoMoveAndDeleteAll = $window.FindName("rdoMoveAndDeleteAll")
    $rdoMoveOnly = $window.FindName("rdoMoveOnly")
    $rdoDeleteOnly = $window.FindName("rdoDeleteOnly")
    $rdoCustom = $window.FindName("rdoCustom")
    $txtModeDescription = $window.FindName("txtModeDescription")
    
    # 階層選択関連
    $pnlMoveLayerSelection = $window.FindName("pnlMoveLayerSelection")
    $pnlDeleteOptions = $window.FindName("pnlDeleteOptions")
    $lblMoveLayersTitle = $window.FindName("lblMoveLayersTitle")
    
    # 削除オプション
    $rdoDeleteAllEmpty = $window.FindName("rdoDeleteAllEmpty")
    $rdoDeleteSelectedOnly = $window.FindName("rdoDeleteSelectedOnly")
    $rdoNoDelete = $window.FindName("rdoNoDelete")
    
    # 移動先選択
    $pnlMoveDestination = $window.FindName("pnlMoveDestination")
    $rdoMoveToRoot = $window.FindName("rdoMoveToRoot")
    $rdoMoveToParent = $window.FindName("rdoMoveToParent")
    
    # その他
    $btnUpdatePreview = $window.FindName("btnUpdatePreview")
    $btnExecute = $window.FindName("btnExecute")
    $btnUndo = $window.FindName("btnUndo")
    $btnClose = $window.FindName("btnClose")
    $chkCreateBackup = $window.FindName("chkCreateBackup")
    $txtExecutionStatus = $window.FindName("txtExecutionStatus")
    $txtStatusBar = $window.FindName("txtStatusBar")
    
    # ================================================================================
    # イベントハンドラーの設定
    # ================================================================================
    
    # 操作モード変更イベント
    $rdoMoveAndDeleteAll.Add_Checked({
        $script:OperationMode = "MoveAndDeleteAll"
        $txtModeDescription.Text = "選択した階層のフォルダを移動し、すべての空フォルダを削除します。"
        $lblMoveLayersTitle.Text = "移動する階層を選択（複数可）:"
        $pnlDeleteOptions.Visibility = "Collapsed"
        $pnlMoveDestination.Visibility = "Visible"
        if ($script:AutoPreviewEnabled) { Invoke-AutoPreviewUpdate }
    })
    
    $rdoMoveOnly.Add_Checked({
        $script:OperationMode = "MoveOnly"
        $txtModeDescription.Text = "選択した階層のフォルダを移動します。空フォルダは削除しません。"
        $lblMoveLayersTitle.Text = "移動する階層を選択（複数可）:"
        $pnlDeleteOptions.Visibility = "Collapsed"
        $pnlMoveDestination.Visibility = "Visible"
        if ($script:AutoPreviewEnabled) { Invoke-AutoPreviewUpdate }
    })
    
    $rdoDeleteOnly.Add_Checked({
        $script:OperationMode = "DeleteOnly"
        $txtModeDescription.Text = "選択した階層のフォルダを削除します。移動は行いません。"
        $lblMoveLayersTitle.Text = "削除する階層を選択（複数可）:"
        $pnlDeleteOptions.Visibility = "Collapsed"
        $pnlMoveDestination.Visibility = "Collapsed"
        if ($script:AutoPreviewEnabled) { Invoke-AutoPreviewUpdate }
    })
    
    $rdoCustom.Add_Checked({
        $script:OperationMode = "Custom"
        $txtModeDescription.Text = "移動と削除を個別に設定できます。"
        $lblMoveLayersTitle.Text = "移動する階層を選択（複数可）:"
        $pnlDeleteOptions.Visibility = "Visible"
        $pnlMoveDestination.Visibility = "Visible"
        if ($script:AutoPreviewEnabled) { Invoke-AutoPreviewUpdate }
    })
    
    # 削除範囲オプション
    $rdoDeleteAllEmpty.Add_Checked({
        $script:DeleteRange = "AllEmpty"
        if ($script:AutoPreviewEnabled) { Invoke-AutoPreviewUpdate }
    })
    
    $rdoDeleteSelectedOnly.Add_Checked({
        $script:DeleteRange = "SelectedOnly"
        if ($script:AutoPreviewEnabled) { Invoke-AutoPreviewUpdate }
    })
    
    $rdoNoDelete.Add_Checked({
        $script:DeleteRange = "NoDelete"
        if ($script:AutoPreviewEnabled) { Invoke-AutoPreviewUpdate }
    })
    
    # 移動先選択イベント
    $rdoMoveToRoot.Add_Checked({
        $script:MoveDestination = "Root"
        if ($script:AutoPreviewEnabled) { Invoke-AutoPreviewUpdate }
    })
    
    $rdoMoveToParent.Add_Checked({
        $script:MoveDestination = "Parent"
        if ($script:AutoPreviewEnabled) { Invoke-AutoPreviewUpdate }
    })
    
    # 参照ボタン
    $btnBrowse.Add_Click({
        $selectedPath = Show-FolderBrowserDialog
        if ($selectedPath) {
            $txtFolderPath.Text = $selectedPath
            $script:RootPath = $selectedPath
            $txtStatusBar.Text = "フォルダを選択しました: $selectedPath"
            $btnAnalyze.IsEnabled = $true
            Write-Log "対象フォルダを選択: $selectedPath"
        }
    })
    
    # エクスポートボタン
    $btnExportAnalysis.Add_Click({
        try {
            if (-not $script:HierarchyData -or $script:HierarchyData.Count -eq 0) {
                [System.Windows.MessageBox]::Show("先に分析を実行してください。", "確認", "OK", "Warning")
                return
            }
            
            $txtStatusBar.Text = "分析結果をエクスポートしています..."
            
            # 出力ファイル名の生成
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $parentPath = Split-Path $script:RootPath -Parent
            $txtPath = Join-Path $parentPath "分析結果_$timestamp.txt"
            $csvPath = Join-Path $parentPath "分析結果_$timestamp.csv"
            
            # テキスト形式でエクスポート
            $txtSuccess = Export-AnalysisToText -RootPath $script:RootPath -HierarchyData $script:HierarchyData -OutputPath $txtPath
            
            # CSV形式でエクスポート
            $csvSuccess = Export-AnalysisToCSV -RootPath $script:RootPath -HierarchyData $script:HierarchyData -OutputPath $csvPath
            
            if ($txtSuccess -and $csvSuccess) {
                $txtStatusBar.Text = "エクスポート完了"
                $message = "分析結果をエクスポートしました。`n`n" +
                           "テキスト: $txtPath`n" +
                           "CSV: $csvPath`n`n" +
                           "エクスプローラーで開きますか？"
                
                $result = [System.Windows.MessageBox]::Show($message, "エクスポート完了", "YesNo", "Information")
                
                if ($result -eq "Yes") {
                    explorer.exe "/select,$txtPath"
                }
            }
            else {
                $txtStatusBar.Text = "エクスポートエラー"
                [System.Windows.MessageBox]::Show("エクスポート中にエラーが発生しました。`nログファイルを確認してください。", "エラー", "OK", "Error")
            }
        }
        catch {
            Write-Log "エクスポートエラー: $_" -Level ERROR
            $txtStatusBar.Text = "エクスポートエラー"
            [System.Windows.MessageBox]::Show("エクスポート中にエラーが発生しました:`n$_", "エラー", "OK", "Error")
        }
    })
    
    # 分析ボタン
    $btnAnalyze.Add_Click({
        try {
            $txtAnalysisStatus.Text = "分析中..."
            $txtAnalysisStatus.Foreground = "Blue"
            $txtStatusBar.Text = "階層構造を分析しています..."
            
            # 階層分析実行
            $script:HierarchyData = Get-FolderHierarchy -RootPath $script:RootPath
            
            if ($script:HierarchyData.Count -eq 0) {
                $txtAnalysisStatus.Text = "フォルダが見つかりませんでした。"
                $txtAnalysisStatus.Foreground = "Red"
                $txtStatusBar.Text = "分析完了: フォルダなし"
                return
            }
            
            # 結果表示
            $totalFolders = ($script:HierarchyData.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
            $maxLevel = ($script:HierarchyData.Keys | Measure-Object -Maximum).Maximum
            
            $txtAnalysisStatus.Text = "分析完了: $maxLevel 階層、合計 $totalFolders 個のフォルダ"
            $txtAnalysisStatus.Foreground = "Green"
            $txtStatusBar.Text = "分析完了"
            
            # 階層選択UIを更新
            $pnlMoveLayerSelection = $window.FindName("pnlMoveLayerSelection")
            Show-LayerAnalysis -HierarchyData $script:HierarchyData -Panel $pnlMoveLayerSelection
            
            # 自動プレビュー更新を有効化
            $script:AutoPreviewEnabled = $true
            
            # エクスポートボタンを有効化
            $btnExportAnalysis.IsEnabled = $true
            
            Write-Log "階層分析完了: $maxLevel 階層" -Level SUCCESS
        }
        catch {
            $txtAnalysisStatus.Text = "エラー: $_"
            $txtAnalysisStatus.Foreground = "Red"
            $txtStatusBar.Text = "分析エラー"
            Write-Log "分析エラー: $_" -Level ERROR
            [System.Windows.MessageBox]::Show("階層分析中にエラーが発生しました:`n$_", "エラー", "OK", "Error")
        }
    })
    
    # プレビュー更新ボタン
    $btnUpdatePreview.Add_Click({
        try {
            if ($script:SelectedLayers.Count -eq 0) {
                [System.Windows.MessageBox]::Show("階層を選択してください。", "確認", "OK", "Warning")
                return
            }
            
            $txtStatusBar.Text = "プレビューを生成しています..."
            
            # プレビューデータ生成（複数階層対応）
            $script:PreviewData = Get-PreviewData `
                -HierarchyData $script:HierarchyData `
                -TargetLevels $script:SelectedLayers `
                -RootPath $script:RootPath `
                -OperationMode $script:OperationMode `
                -DeleteRange $script:DeleteRange `
                -MoveDestination $script:MoveDestination
            
            if ($script:PreviewData) {
                Update-PreviewDisplay -PreviewData $script:PreviewData
                $txtStatusBar.Text = "プレビュー更新完了"
                Write-Log "プレビュー更新完了" -Level SUCCESS
            }
            else {
                [System.Windows.MessageBox]::Show("プレビュー生成に失敗しました。", "エラー", "OK", "Error")
                $txtStatusBar.Text = "プレビュー生成エラー"
            }
        }
        catch {
            Write-Log "プレビュー更新エラー: $_" -Level ERROR
            [System.Windows.MessageBox]::Show("プレビュー更新中にエラーが発生しました:`n$_", "エラー", "OK", "Error")
        }
    })
    
    # 実行ボタン
    $btnExecute.Add_Click({
        try {
            if (-not $script:PreviewData) {
                [System.Windows.MessageBox]::Show("プレビューを更新してください。", "確認", "OK", "Warning")
                return
            }
            
            # モードに応じた確認メッセージ
            $moveCount = $script:PreviewData.MoveOperations.Count
            $deleteCount = $script:PreviewData.EmptyFoldersToDelete.Count
            
            $message = "以下の操作を実行します:`n`n"
            
            switch ($script:OperationMode) {
                "MoveAndDeleteAll" {
                    $message += "・$moveCount 個のフォルダをルート直下に移動`n"
                    $message += "・すべての空フォルダを削除`n"
                }
                "MoveOnly" {
                    $message += "・$moveCount 個のフォルダをルート直下に移動`n"
                    $message += "・空フォルダは削除しません`n"
                }
                "DeleteOnly" {
                    $message += "・$deleteCount 個のフォルダを削除`n"
                    $message += "・移動は行いません`n"
                }
                "Custom" {
                    $message += "・$moveCount 個のフォルダをルート直下に移動`n"
                    $message += "・削除範囲: $script:DeleteRange`n"
                }
            }
            
            $message += "`n実行してもよろしいですか？"
            
            $result = [System.Windows.MessageBox]::Show($message, "実行確認", "YesNo", "Question")
            
            if ($result -ne "Yes") {
                $txtStatusBar.Text = "実行をキャンセルしました"
                return
            }
            
            $txtExecutionStatus.Text = "実行中..."
            $txtExecutionStatus.Foreground = "Blue"
            $txtStatusBar.Text = "処理を実行しています..."
            $btnExecute.IsEnabled = $false
            
            # バックアップ作成
            if ($chkCreateBackup.IsChecked) {
                $txtExecutionStatus.Text = "バックアップを作成中..."
                $backupPath = New-Backup -SourcePath $script:RootPath
                
                if (-not $backupPath) {
                    $txtExecutionStatus.Text = "バックアップ作成失敗"
                    $txtExecutionStatus.Foreground = "Red"
                    [System.Windows.MessageBox]::Show("バックアップの作成に失敗しました。`n処理を中止します。", "エラー", "OK", "Error")
                    $btnExecute.IsEnabled = $true
                    return
                }
                
                $btnUndo.IsEnabled = $true
            }
            
            # モードに応じた処理実行
            $moveResult = @{ Success = 0; Error = 0 }
            $deleteCount = 0
            
            if ($script:OperationMode -eq "DeleteOnly") {
                # 削除のみモード
                $txtExecutionStatus.Text = "フォルダを削除中..."
                $deleteCount = Remove-SelectedLayers -FoldersToDelete $script:PreviewData.EmptyFoldersToDelete
                $txtExecutionStatus.Text = "完了: $deleteCount 個削除"
            }
            else {
                # フォルダ移動実行
                $txtExecutionStatus.Text = "フォルダを移動中..."
                $moveResult = Move-MultipleLayers -PreviewData $script:PreviewData -RootPath $script:RootPath -OperationMode $script:OperationMode
                
                # 空フォルダ削除（モードに応じて）
                if ($script:OperationMode -ne "MoveOnly") {
                    $txtExecutionStatus.Text = "空フォルダを削除中..."
                    $deleteCount = Remove-EmptyFolders -RootPath $script:RootPath -DeleteRange $script:DeleteRange -SelectedLevels $script:SelectedLayers
                }
                
                $txtExecutionStatus.Text = "完了: $($moveResult.Success) 個移動、$deleteCount 個削除"
            }
            
            $txtExecutionStatus.Foreground = "Green"
            $txtStatusBar.Text = "処理完了"
            
            Write-Log "処理完了" -Level SUCCESS
            
            # 結果メッセージ
            $message = "処理が完了しました。`n`n"
            
            if ($script:OperationMode -eq "DeleteOnly") {
                $message += "削除: $deleteCount 個のフォルダ"
            }
            else {
                $message += "移動: $($moveResult.Success) 個 (失敗: $($moveResult.Error) 個)`n"
                $message += "削除: $deleteCount 個の空フォルダ"
            }
            
            [System.Windows.MessageBox]::Show($message, "完了", "OK", "Information")
        }
        catch {
            $txtExecutionStatus.Text = "エラー: $_"
            $txtExecutionStatus.Foreground = "Red"
            $txtStatusBar.Text = "実行エラー"
            Write-Log "実行エラー: $_" -Level ERROR
            [System.Windows.MessageBox]::Show("実行中にエラーが発生しました:`n$_", "エラー", "OK", "Error")
        }
        finally {
            $btnExecute.IsEnabled = $true
        }
    })
    
    # 元に戻すボタン
    $btnUndo.Add_Click({
        try {
            # 最新のバックアップを検索
            $backupPath = Get-LatestBackup -OriginalPath $script:RootPath
            
            if (-not $backupPath) {
                if ($script:LastBackupPath) {
                    $backupPath = $script:LastBackupPath
                }
                else {
                    [System.Windows.MessageBox]::Show("バックアップが見つかりません。", "エラー", "OK", "Error")
                    return
                }
            }
            
            $message = "バックアップから復元します。`n現在の内容は失われます。`n`nバックアップ: $backupPath`n`n実行してもよろしいですか？"
            $result = [System.Windows.MessageBox]::Show($message, "復元確認", "YesNo", "Warning")
            
            if ($result -ne "Yes") {
                return
            }
            
            $txtExecutionStatus.Text = "復元中..."
            $txtExecutionStatus.Foreground = "Blue"
            $txtStatusBar.Text = "バックアップから復元しています..."
            
            $success = Restore-FromBackup -BackupPath $backupPath -DestinationPath $script:RootPath
            
            if ($success) {
                $txtExecutionStatus.Text = "復元完了"
                $txtExecutionStatus.Foreground = "Green"
                $txtStatusBar.Text = "復元完了"
                [System.Windows.MessageBox]::Show("バックアップから復元しました。", "完了", "OK", "Information")
            }
            else {
                $txtExecutionStatus.Text = "復元失敗"
                $txtExecutionStatus.Foreground = "Red"
                $txtStatusBar.Text = "復元エラー"
                [System.Windows.MessageBox]::Show("復元に失敗しました。", "エラー", "OK", "Error")
            }
        }
        catch {
            Write-Log "復元エラー: $_" -Level ERROR
            [System.Windows.MessageBox]::Show("復元中にエラーが発生しました:`n$_", "エラー", "OK", "Error")
        }
    })
    
    # 閉じるボタン
    $btnClose.Add_Click({
        Write-Log "アプリケーションを終了します"
        $window.Close()
    })
    
    # ウィンドウを表示
    Write-Host "UIを表示します..." -ForegroundColor Cyan
    Write-Log "UIを表示します"
    $window.ShowDialog() | Out-Null
    
    Write-Host "アプリケーションが正常に終了しました" -ForegroundColor Green
}
catch {
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "致命的なエラーが発生しました" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "`nエラー内容:" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "`nエラー発生箇所:" -ForegroundColor Yellow
    Write-Host $_.InvocationInfo.PositionMessage -ForegroundColor Gray
    
    if ($_.Exception.InnerException) {
        Write-Host "`n内部エラー:" -ForegroundColor Yellow
        Write-Host $_.Exception.InnerException.Message -ForegroundColor Red
    }
    
    Write-Host "`nスタックトレース:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "対処方法:" -ForegroundColor Yellow
    Write-Host "1. MainWindow.xamlが同じフォルダにあることを確認" -ForegroundColor White
    Write-Host "2. PowerShell 5.1以降を使用していることを確認" -ForegroundColor White
    Write-Host "3. test-startup.ps1で診断を実行" -ForegroundColor White
    Write-Host "========================================" -ForegroundColor Red
    
    Read-Host "`nEnterキーを押して終了してください"
    exit 1
}

