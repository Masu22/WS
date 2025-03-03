#Chromeでアドレスを開いて、自動ダウンロード

# URLリストのテキストファイルのパス
$urlsFile = "urls.txt"

# Chromeのパス（デフォルトのインストール先）
$chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

# エラーログファイルのパス
$errorLog = "error_log.txt"

# エラーログをクリア（前回の実行結果を消す）
Clear-Content -Path $errorLog -ErrorAction SilentlyContinue

# URLリストを1行ずつ読み込んで処理
Get-Content $urlsFile | ForEach-Object {
    $url = $_.Trim()
    if ($url -ne "") {  # 空行をスキップ
        Write-Host "Opening: $url"
        try {
            Start-Process -FilePath $chromePath -ArgumentList $url -ErrorAction Stop
            Start-Sleep -Seconds 15  # ダウンロード待機時間（必要なら調整）
        } catch {
            $errorMessage = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Error opening $url : $_"
            Write-Host $errorMessage
            Add-Content -Path $errorLog -Value $errorMessage
        }
    }
}
Write-Host "All URLs processed!"

