# FileRenamer.psm1 - ファイル名をフォルダ名でリネームするPowerShellモジュール

function Rename-FilesToFolderName {
    <#
    .SYNOPSIS
    指定フォルダ内のファイル名をそれを包含するフォルダ名でリネームします。

    .DESCRIPTION
    ルートフォルダ内の各サブフォルダにある指定ファイル名のファイルを、
    そのサブフォルダ名と同じ名前にリネームします。

    .PARAMETER RootFolder
    処理対象のルートフォルダパス

    .PARAMETER TargetFileName
    リネーム対象のファイル名

    .PARAMETER Force
    確認プロンプトをスキップして強制実行

    .EXAMPLE
    Rename-FilesToFolderName -RootFolder "C:\Projects" -TargetFileName "report.xlsx"
    
    .EXAMPLE
    Rename-FilesToFolderName -RootFolder "C:\Projects" -TargetFileName "report.xlsx" -Force
    
    .OUTPUTS
    処理結果オブジェクト (成功数、エラー数、詳細情報)
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory=$true, 
                   Position=0,
                   HelpMessage="処理対象のルートフォルダパスを指定してください")]
        [ValidateScript({
            if (-not (Test-Path $_ -PathType Container)) {
                throw "指定されたフォルダが存在しません: $_"
            }
            $true
        })]
        [string]$RootFolder,
        
        [Parameter(Mandatory=$true, 
                   Position=1,
                   HelpMessage="リネーム対象のファイル名を指定してください")]
        [ValidateNotNullOrEmpty()]
        [string]$TargetFileName,
        
        [Parameter()]
        [switch]$Force
        
    )
    
    begin {
        # 結果オブジェクトの初期化
        $result = [PSCustomObject]@{
            Success = 0
            Errors = 0
            Total = 0
            Details = @()
        }
        
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host "ファイルリネーム処理" -ForegroundColor Cyan
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host "対象ルートフォルダ: $RootFolder" -ForegroundColor Green
        Write-Host "対象ファイル名: $TargetFileName" -ForegroundColor Green
        Write-Host ""
    }
    
    process {
        # サブフォルダを取得
        $subFolders = Get-ChildItem -Path $RootFolder -Directory
        
        if ($subFolders.Count -eq 0) {
            Write-Warning "サブフォルダが見つかりませんでした。"
            return $result
        }
        
        Write-Host "見つかったサブフォルダ数: $($subFolders.Count)" -ForegroundColor Green
        Write-Host ""
        
        # 処理対象ファイルの確認
        $processTargets = @()
        foreach ($folder in $subFolders) {
            $targetFilePath = Join-Path $folder.FullName $TargetFileName
            if (Test-Path $targetFilePath) {
                $processTargets += @{
                    Folder = $folder
                    FilePath = $targetFilePath
                }
            }
        }
        
        if ($processTargets.Count -eq 0) {
            Write-Warning "対象ファイル '$TargetFileName' が見つかりませんでした。"
            return $result
        }
        
        $result.Total = $processTargets.Count
        Write-Host "処理対象ファイル数: $($processTargets.Count)" -ForegroundColor Green
        Write-Host ""
        
        # 処理対象の一覧表示
        Write-Host "処理対象一覧:" -ForegroundColor Cyan
        foreach ($target in $processTargets) {
            $folderName = $target.Folder.Name
            $fileExtension = [System.IO.Path]::GetExtension($TargetFileName)
            $newFileName = $folderName + $fileExtension
            
            Write-Host "  フォルダ: $folderName" -ForegroundColor White
            Write-Host "    変更前: $TargetFileName" -ForegroundColor Gray
            Write-Host "    変更後: $newFileName" -ForegroundColor Gray
            Write-Host ""
        }
        
        # WhatIf モードの場合はここで終了
        if ($WhatIf) {
            Write-Host "WhatIfモード: 実際の処理は実行されませんでした。" -ForegroundColor Yellow
            return $result
        }
        
        # 実行確認（Force スイッチがない場合）
        if (-not $Force) {
            $confirmation = Read-Host "上記の処理を実行しますか？ (y/N)"
            if ($confirmation -ne 'y' -and $confirmation -ne 'Y') {
                Write-Host "処理をキャンセルしました。" -ForegroundColor Yellow
                return $result
            }
        }
        
        Write-Host ""
        Write-Host "処理を開始します..." -ForegroundColor Green
        Write-Host ""
        
        # リネーム処理実行
        foreach ($target in $processTargets) {
            $folderName = $target.Folder.Name
            $currentFilePath = $target.FilePath
            $fileExtension = [System.IO.Path]::GetExtension($TargetFileName)
            $newFileName = $folderName + $fileExtension
            $newFilePath = Join-Path $target.Folder.FullName $newFileName
            
            $detail = [PSCustomObject]@{
                FolderName = $folderName
                OldFileName = $TargetFileName
                NewFileName = $newFileName
                Status = ""
                Error = $null
            }
            
            try {
                # 同名ファイルが既に存在する場合のチェック
                if (Test-Path $newFilePath) {
                    $detail.Status = "Skipped"
                    $detail.Error = "変更後のファイル名が既に存在します"
                    Write-Host "  警告: '$folderName' - 変更後のファイル名が既に存在します。スキップします。" -ForegroundColor Yellow
                }
                elseif ($PSCmdlet.ShouldProcess($currentFilePath, "Rename to $newFileName")) {
                    # リネーム実行
                    Rename-Item -Path $currentFilePath -NewName $newFileName -ErrorAction Stop
                    $detail.Status = "Success"
                    Write-Host "  成功: '$folderName' - リネーム完了" -ForegroundColor Green
                    $result.Success++
                }
            }
            catch {
                $detail.Status = "Error"
                $detail.Error = $_.Exception.Message
                Write-Host "  エラー: '$folderName' - $($_.Exception.Message)" -ForegroundColor Red
                $result.Errors++
            }
            
            $result.Details += $detail
        }
    }
    
    end {
        # 処理結果の表示
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host "処理完了" -ForegroundColor Cyan
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host "成功: $($result.Success) 件" -ForegroundColor Green
        Write-Host "エラー: $($result.Errors) 件" -ForegroundColor Red
        Write-Host "合計: $($result.Total) 件" -ForegroundColor White
        
        if ($result.Errors -gt 0) {
            Write-Host ""
            Write-Host "エラーが発生したファイルがあります。戻り値の Details プロパティで詳細を確認してください。" -ForegroundColor Yellow
        }
        
        return $result
    }
}

# タブ補完機能の登録
Register-ArgumentCompleter -CommandName 'Rename-FilesToFolderName' -ParameterName 'RootFolder' -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    
    Get-ChildItem -Path "$wordToComplete*" -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        if ($_.FullName -like "*$wordToComplete*") {
            "'$($_.FullName)'"
        }
    }
}

Register-ArgumentCompleter -CommandName 'Rename-FilesToFolderName' -ParameterName 'TargetFileName' -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)
    
    $commonExtensions = @('.xlsx', '.xls', '.csv', '.txt', '.pdf', '.docx', '.doc', '.pptx', '.ppt')
    
    if ([string]::IsNullOrEmpty($wordToComplete)) {
        @('自動生成ファイル.xlsx', 'report.xlsx', 'data.csv', 'output.txt')
    } else {
        $commonExtensions | ForEach-Object {
            if ($wordToComplete.Contains('.')) {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($wordToComplete)
                "'$baseName$_'"
            } else {
                "'$wordToComplete$_'"
            }
        }
    }
}

# エクスポートする関数
Export-ModuleMember -Function Rename-FilesToFolderName
