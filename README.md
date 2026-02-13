# ktp-xls (Vercel A2)

Клиентское приложение для конвертации КТП (`.docx`/`.odt`) в Excel (`.xls`/`.xlsx`) для импорта в Дневник.ру.

## Что важно

- Приложение работает **полностью в браузере** (без backend).
- Файлы не отправляются на сервер.
- Подходит для деплоя на **Vercel** как обычный Vite frontend.

## Локальный запуск

```bash
npm install
npm run dev
```

Откройте `http://localhost:5173`

## Production build

```bash
npm run build
npm run preview
```

## Деплой на Vercel

1. Загрузите репозиторий в GitHub.
2. В Vercel: **Add New Project** -> выберите репозиторий.
3. Настройки:
   - Framework preset: `Vite`
   - Build command: `npm run build`
   - Output directory: `dist`
4. Нажмите **Deploy**.

После этого каждый push в ветку, подключенную к проекту Vercel, будет деплоиться автоматически.

## Структура

- `index.html` — входная страница
- `src/main.js` — вся логика приложения
- `public/app.css` — стили
