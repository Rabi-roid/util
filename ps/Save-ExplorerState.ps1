# エクスプローラーの状態を記録するスクリプト
# Save-ExplorerState.ps1

function Save-ExplorerState {
    $explorerPaths = @()
    
    # 現在開いているエクスプローラーウィンドウを取得
    $shell = New-Object -ComObject Shell.Application
    $windows = $shell.Windows()
    
    foreach ($window in $windows) {
        if ($window.Name -eq "エクスプローラー" -or $window.Name -eq "File Explorer") {
            $path = $window.LocationURL -replace "file:///", "" -replace "/", "\"
            $path = [System.Uri]::UnescapeDataString($path)
            if ($path -and $path -ne "") {
                $explorerPaths += $path
                Write-Host "記録: $path"
            }
        }
    }
    
    # パスをファイルに保存
    $explorerPaths | Out-File -FilePath "$env:TEMP\explorer_paths.txt" -Encoding UTF8
    Write-Host "エクスプローラーの状態を保存しました: $($explorerPaths.Count)個のタブ/ウィンドウ"
    
    return $explorerPaths
}

function Restart-Explorer {
    Write-Host "エクスプローラーを再起動しています..."
    
    # エクスプローラープロセスを終了
    Stop-Process -Name "explorer" -Force
    
    # 少し待機
    Start-Sleep -Seconds 2
    
    # エクスプローラーを再開
    Start-Process "explorer.exe"
    
    # エクスプローラーが完全に起動するまで待機
    Start-Sleep -Seconds 3
    
    Write-Host "エクスプローラーが再起動されました"
}

function Restore-ExplorerState {
    $pathsFile = "$env:TEMP\explorer_paths.txt"
    
    if (-not (Test-Path $pathsFile)) {
        Write-Host "保存されたパス情報が見つかりません"
        return
    }
    
    $savedPaths = Get-Content $pathsFile -Encoding UTF8
    
    Write-Host "保存されたパスを復元しています..."
    
    foreach ($path in $savedPaths) {
        if ($path -and (Test-Path $path)) {
            Start-Process "explorer.exe" -ArgumentList $path
            Write-Host "復元: $path"
            Start-Sleep -Milliseconds 500  # 各ウィンドウの起動間隔を空ける
        }
        else {
            Write-Host "パスが存在しません: $path"
        }
    }
    
    Write-Host "復元が完了しました"
}

# メイン実行部分
function Backup-And-Restart-Explorer {
    Write-Host "=== エクスプローラー状態の記録・再起動・復元 ==="
    
    # 1. 現在の状態を記録
    Save-ExplorerState
    
    # 2. エクスプローラーを再起動
    Restart-Explorer
    
    # 3. 状態を復元
    Restore-ExplorerState
}

# 使用例:
# 全工程を一度に実行する場合
# Backup-And-Restart-Explorer

# 個別に実行する場合
# Save-ExplorerState        # 状態を記録
# Restart-Explorer          # 再起動
# Restore-ExplorerState     # 復元
