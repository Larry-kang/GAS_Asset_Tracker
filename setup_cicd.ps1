$ErrorActionPreference = 'Stop'

# �ؼСGŪ�� ~/.clasprc.json ���e���ഫ�� Base64
# �M��N��ƻs�� GitHub Secrets (CLASPRC_JSON_BASE64) �H�� CI/CD �ϥ�

$claspPath = "$HOME/.clasprc.json"

if (!(Test-Path $claspPath)) {
    Write-Host "���~�G�䤣�� .clasprc.json ���|�G$claspPath" -ForegroundColor Red
    Write-Host "�Х����� 'clasp login' �i��n�J�C"
    exit 1
}

$claspContent = Get-Content $claspPath -Raw -Encoding utf8
$bytes = [System.Text.Encoding]::UTF8.GetBytes($claspContent)
$base64Content = [System.Convert]::ToBase64String($bytes)

Write-Host "������ .clasprc.json�C���b�]�w GitHub Secret..."

try {
    # Pipe the base64 content to gh secret set
    $base64Content | gh secret set CLASPRC_JSON_BASE64 --app actions
    Write-Host "�����IGitHub Secret �w���\��s�C" -ForegroundColor Green
}
catch {
    Write-Host "�۰ʳ]�w���ѡC�Ф�ʽƻs�H�U Base64 ���e (���)�G" -ForegroundColor Yellow
    Write-Host "---"
    Write-Host $base64Content
    Write-Host "---"
}
