# Деплой НПП Регистр заявок на pm2 (Windows, порт 3043).
# Запуск: powershell -ExecutionPolicy Bypass -File deploy\deploy.ps1
$ErrorActionPreference = 'Stop'
Set-Location (Join-Path $PSScriptRoot '..')
$app = 'npp-register'

if (-not (Test-Path .env)) { throw 'Нет файла .env — скопируйте .env.example в .env и заполните' }
if (-not (Get-Command pm2 -ErrorAction SilentlyContinue)) { throw 'pm2 не найден: npm i -g pm2' }

Write-Host '==> Установка зависимостей'
npm --prefix server install --no-audit --no-fund
npm --prefix client install --no-audit --no-fund

Write-Host '==> Сборка'
npm run build
if ($LASTEXITCODE -ne 0) { throw 'Сборка не удалась' }

New-Item -ItemType Directory -Force logs | Out-Null

Write-Host '==> Запуск через pm2'
pm2 describe $app *> $null
if ($LASTEXITCODE -eq 0) { pm2 reload ecosystem.config.js --update-env } else { pm2 start ecosystem.config.js }
pm2 save

$ok = $false
for ($i = 0; $i -lt 20 -and -not $ok; $i++) {
  try { Invoke-RestMethod http://localhost:3043/api/health | Out-Null; $ok = $true } catch { Start-Sleep -Seconds 2 }
}
if (-not $ok) { pm2 logs $app --lines 30 --nostream; throw "Сервис не отвечает, логи: pm2 logs $app" }
Write-Host 'OK: http://localhost:3043'
