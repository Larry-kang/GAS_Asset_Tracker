$ErrorActionPreference = 'Stop'
Write-Host "? Pushing code to Google Apps Script..."
clasp push -f
Write-Host "? Code Pushed! Please go to GAS Editor run 'getBinanceBalance' to verify."
