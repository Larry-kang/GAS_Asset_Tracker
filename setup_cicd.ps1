$ErrorActionPreference = 'Stop'

# 目標：讀取 ~/.clasprc.json 內容並轉換為 Base64
# 將內容更新至 GitHub Secrets (CLASPRC_JSON_BASE64) 以供 CI/CD 使用

$claspPath = "$HOME/.clasprc.json"

if (!(Test-Path $claspPath)) {
    Write-Host "錯誤：找不到 .clasprc.json 路徑：$claspPath" -ForegroundColor Red
    Write-Host "請先執行 'clasp login' 進行登入。"
    exit 1
}

$claspContent = Get-Content $claspPath -Raw -Encoding utf8
$bytes = [System.Text.Encoding]::UTF8.GetBytes($claspContent)
$base64Content = [System.Convert]::ToBase64String($bytes)

Write-Host "正在讀取 .clasprc.json... 正在設定 GitHub Secret..."

try {
    # Pipe the base64 content to gh secret set
    $base64Content | gh secret set CLASPRC_JSON_BASE64 --app actions
    Write-Host "成功：GitHub Secret 已成功更新。" -ForegroundColor Green
}
catch {
    Write-Host "自動設定失敗。請手動更新以下 Base64 內容 (複製)：" -ForegroundColor Yellow
    Write-Host "---"
    Write-Host $base64Content
    Write-Host "---"
}

