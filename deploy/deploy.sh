#!/usr/bin/env bash
# Деплой НПП Регистр заявок на pm2 (порт 3043).
# Запуск на сервере из любой директории:  ./deploy/deploy.sh
# Опции: --pull  — перед сборкой сделать git pull (если проект в git)
set -euo pipefail

APP_NAME="npp-register-v2"
cd "$(dirname "$0")/.."

command -v node >/dev/null || { echo "Node.js не найден"; exit 1; }
command -v pm2  >/dev/null || { echo "pm2 не найден: npm i -g pm2"; exit 1; }

[ -f .env ] || { echo "Нет файла .env — скопируйте .env.example в .env и заполните"; exit 1; }

if [ "${1:-}" = "--pull" ]; then
  git pull --ff-only
fi

echo "==> Установка зависимостей"
npm --prefix server ci --no-audit --no-fund || npm --prefix server install --no-audit --no-fund
npm --prefix client ci --no-audit --no-fund || npm --prefix client install --no-audit --no-fund

echo "==> Сборка (client + server)"
npm run build

mkdir -p logs

echo "==> Запуск через pm2"
if pm2 describe "$APP_NAME" >/dev/null 2>&1; then
  pm2 reload ecosystem.config.js --update-env
else
  pm2 start ecosystem.config.js
fi
pm2 save

echo "==> Проверка (ждём до 40 с)"
for i in $(seq 1 20); do
  if curl -fsS "http://localhost:3043/api/health" >/dev/null 2>&1; then
    echo "OK: http://localhost:3043"
    ok=1; break
  fi
  sleep 2
done
if [ "${ok:-0}" != "1" ]; then
  echo "Сервис не отвечает, логи: pm2 logs $APP_NAME"
  pm2 logs "$APP_NAME" --lines 30 --nostream || true
  exit 1
fi
echo "Автозапуск после перезагрузки сервера (один раз): pm2 startup"
