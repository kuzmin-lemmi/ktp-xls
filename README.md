# Конвертер КТП -> Excel

Веб-приложение для подготовки файлов импорта в Дневник.ру из КТП (`.docx`/`.odt`).

## Возможности

- Импорт таблиц из DOCX/ODT
- Выбор столбцов темы/ДЗ
- Режим без столбца ДЗ (`ДЗ нет`)
- Разделение на 4 четверти или 2 полугодия
- Редактор с отменой/возвратом уроков
- Drag&drop тем между строками
- Экспорт `XLS` и `XLSX`

## Локальный запуск

```bash
npm install
npm start
```

Открыть: `http://localhost:3000`

## Продакшен на VDS

Ниже готовый сценарий для Ubuntu 22.04+

### 1) Установка Node.js и Nginx

```bash
sudo apt update
sudo apt install -y curl nginx
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt install -y nodejs
node -v
npm -v
```

### 2) Копирование проекта

```bash
sudo mkdir -p /var/www/ktp-converter
sudo chown -R $USER:$USER /var/www/ktp-converter
cd /var/www/ktp-converter

# вариант 1: git clone
# git clone <repo_url> .

# вариант 2: загрузить архив и распаковать
```

### 3) Настройка окружения

```bash
cp .env.example .env
nano .env
```

Рекомендуемые значения:

```env
NODE_ENV=production
PORT=3000
HOST=127.0.0.1
```

### 4) Установка зависимостей

```bash
npm ci --omit=dev
```

### 5) Запуск через PM2

```bash
sudo npm i -g pm2
pm2 start ecosystem.config.cjs --env production
pm2 save
pm2 startup
```

Проверка:

```bash
pm2 status
curl http://127.0.0.1:3000/health
```

### 6) Настройка Nginx

Скопируйте готовый конфиг:

```bash
sudo cp deploy/nginx-ktp.conf /etc/nginx/sites-available/ktp-converter
```

Откройте и замените `YOUR_DOMAIN` на ваш домен:

```bash
sudo nano /etc/nginx/sites-available/ktp-converter
```

Активируйте сайт:

```bash
sudo ln -s /etc/nginx/sites-available/ktp-converter /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

### 7) SSL (Let's Encrypt)

```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d YOUR_DOMAIN -d www.YOUR_DOMAIN
```

## Полезные команды

```bash
pm2 logs ktp-converter
pm2 restart ktp-converter
pm2 stop ktp-converter
pm2 delete ktp-converter
```

## Структура деплой-файлов

- `ecosystem.config.cjs` - PM2 конфигурация
- `deploy/nginx-ktp.conf` - Nginx reverse proxy
- `.env.example` - пример переменных окружения
