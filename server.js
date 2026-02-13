const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || '0.0.0.0';

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.set('trust proxy', 1);

app.use((req, res, next) => {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'SAMEORIGIN');
  res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
  next();
});

// Логирование всех запросов
app.use((req, res, next) => {
  console.log(`[${new Date().toLocaleTimeString()}] ${req.method} ${req.url}`);
  next();
});

// Создаём директорию для загрузок
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// Раздача статических файлов фронтенда
app.use(express.static(path.join(__dirname, 'public')));

// API routes — только один mount на /api
const apiRoutes = require('./src/routes/api');
app.use('/api', apiRoutes);

app.get('/health', (req, res) => {
  res.json({ ok: true, uptime: process.uptime(), ts: Date.now() });
});

// Главная страница
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Обработка ошибок multer
app.use((err, req, res, next) => {
  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(413).json({ error: 'Файл слишком большой (макс. 50MB)' });
  }
  console.error('Ошибка сервера:', err);
  res.status(500).json({ error: err.message || 'Внутренняя ошибка сервера' });
});

// Запуск сервера
const server = app.listen(PORT, HOST, () => {
  console.log(`\n${'='.repeat(60)}`);
  console.log(`  Конвертер КТП -> Excel`);
  console.log(`${'='.repeat(60)}`);
  console.log(`  Сервер запущен: http://${HOST}:${PORT}`);
  console.log(`  Откройте браузер и перейдите по адресу выше`);
  console.log(`${'='.repeat(60)}\n`);
});

function shutdown(signal) {
  console.log(`Получен ${signal}, завершаю сервер...`);
  server.close(() => {
    console.log('HTTP сервер остановлен');
    process.exit(0);
  });
  setTimeout(() => process.exit(1), 10000);
}

process.on('SIGINT', () => shutdown('SIGINT'));
process.on('SIGTERM', () => shutdown('SIGTERM'));
