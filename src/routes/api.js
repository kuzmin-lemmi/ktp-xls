const express = require('express');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const tableParser = require('../services/tableParser');
const partService = require('../services/partService');
const exportService = require('../services/exportService');
const path = require('path');
const fs = require('fs');

const router = express.Router();

// Хранилище сессий
const sessions = new Map();

// Настройка multer
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (['.docx', '.odt'].includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Поддерживаются только файлы .docx и .odt'));
    }
  }
});

// ===================================================================
// POST /api/upload — Загружает файл и парсит таблицы
// ===================================================================
router.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Файл не загружен' });
    }

    const fileExt = path.extname(req.file.originalname).toLowerCase();
    let tables;

    console.log(`[upload] Файл: ${req.file.originalname}, формат: ${fileExt}`);

    if (fileExt === '.docx') {
      tables = await tableParser.parseDocx(req.file.buffer);
    } else if (fileExt === '.odt') {
      tables = await tableParser.parseOdt(req.file.buffer);
    } else {
      return res.status(400).json({ error: 'Неподдерживаемый формат файла' });
    }

    if (!tables || tables.length === 0) {
      return res.status(400).json({ error: 'Таблицы не найдены в документе' });
    }

    console.log(`[upload] Найдено таблиц: ${tables.length}`);

    const sessionId = uuidv4();
    sessions.set(sessionId, {
      fileName: req.file.originalname,
      tables,
      selectedTableIdx: tables.length === 1 ? 0 : null,
      temaCol: null,
      dzCol: null,
      noDz: false,
      hasHeader: true,
      separator: '\n',
      periodMode: 'quarters',
      partCounts: [],
      parts: [],
      totalDataRows: 0
    });

    // Если одна таблица — автоопределяем столбцы сразу
    if (tables.length === 1) {
      autoDetectColumns(sessions.get(sessionId), 0);
    }

    res.json({
      sessionId,
      fileName: req.file.originalname,
      tables: tables.map((t, i) => ({
        id: i,
        rowCount: t.rows.length,
        colCount: Math.max(...t.rows.map(r => r.length))
      }))
    });
  } catch (error) {
    console.error('[upload] Ошибка:', error);
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// POST /api/session/:id/select-table — Выбирает таблицу
// ===================================================================
router.post('/session/:id/select-table', (req, res) => {
  try {
    const session = sessions.get(req.params.id);
    if (!session) {
      return res.status(404).json({ error: 'Сессия не найдена' });
    }

    const { tableIdx } = req.body;
    if (tableIdx < 0 || tableIdx >= session.tables.length) {
      return res.status(400).json({ error: 'Неверный индекс таблицы' });
    }

    session.selectedTableIdx = tableIdx;
    autoDetectColumns(session, tableIdx);

    console.log(`[select-table] Таблица ${tableIdx}, temaCol=${session.temaCol}, dzCol=${session.dzCol}`);

    res.json({
      success: true,
      temaCol: session.temaCol,
      dzCol: session.dzCol,
      hasHeader: session.hasHeader
    });
  } catch (error) {
    console.error('[select-table] Ошибка:', error);
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// GET /api/session/:id/preview — Предпросмотр таблицы
// ===================================================================
router.get('/session/:id/preview', (req, res) => {
  try {
    const session = sessions.get(req.params.id);
    if (!session) {
      return res.status(404).json({ error: 'Сессия не найдена' });
    }
    if (session.selectedTableIdx === null) {
      return res.status(400).json({ error: 'Таблица не выбрана' });
    }

    const table = session.tables[session.selectedTableIdx];
    const maxRows = Math.min(table.rows.length, 15);
    const previewRows = table.rows.slice(0, maxRows);
    const totalRows = table.rows.length;
    const dataRows = session.hasHeader ? totalRows - 1 : totalRows;

    res.json({
      headers: table.rows[0] || [],
      rows: previewRows,
      temaCol: session.temaCol,
      dzCol: session.dzCol,
      totalRows: totalRows,
      dataRows: dataRows
    });
  } catch (error) {
    console.error('[preview] Ошибка:', error);
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// POST /api/session/:id/columns — Устанавливает столбцы Тема и ДЗ
// ===================================================================
router.post('/session/:id/columns', (req, res) => {
  try {
    const session = sessions.get(req.params.id);
    if (!session) {
      return res.status(404).json({ error: 'Сессия не найдена' });
    }

    const { temaCol, dzCol, hasHeader, separator, noDz } = req.body;

    if (temaCol === undefined || temaCol === null) {
      return res.status(400).json({ error: 'Столбец темы должен быть выбран' });
    }
    if (!noDz && (dzCol === undefined || dzCol === null)) {
      return res.status(400).json({ error: 'Выберите столбец ДЗ или отметьте, что ДЗ нет' });
    }

    session.temaCol = temaCol;
    session.dzCol = noDz ? null : dzCol;
    session.noDz = !!noDz;
    session.hasHeader = hasHeader !== false;
    session.separator = separator || '\n';

    const table = session.tables[session.selectedTableIdx];
    const startRow = session.hasHeader ? 1 : 0;
    session.totalDataRows = table.rows.length - startRow;

    console.log(`[columns] temaCol=${temaCol}, dzCol=${dzCol}, dataRows=${session.totalDataRows}`);

    res.json({
      success: true,
      totalDataRows: session.totalDataRows
    });
  } catch (error) {
    console.error('[columns] Ошибка:', error);
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// GET /api/session/:id/info — Получить инфо о сессии
// ===================================================================
router.get('/session/:id/info', (req, res) => {
  try {
    const session = sessions.get(req.params.id);
    if (!session) {
      return res.status(404).json({ error: 'Сессия не найдена' });
    }

    const table = session.tables[session.selectedTableIdx];
    const startRow = session.hasHeader ? 1 : 0;
    const totalDataRows = table ? (table.rows.length - startRow) : 0;

    res.json({
      totalDataRows,
      periodMode: session.periodMode,
      partCounts: session.partCounts
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// POST /api/session/:id/parts — Разделяет на части
// ===================================================================
router.post('/session/:id/parts', (req, res) => {
  try {
    const session = sessions.get(req.params.id);
    if (!session) {
      return res.status(404).json({ error: 'Сессия не найдена' });
    }

    const { periodMode, partCounts, defaultDzForAll } = req.body;
    if (!periodMode || !partCounts || partCounts.length === 0) {
      return res.status(400).json({ error: 'Неверные параметры' });
    }

    // Валидация суммы
    const table = session.tables[session.selectedTableIdx];
    const startRow = session.hasHeader ? 1 : 0;
    const totalDataRows = table.rows.length - startRow;
    const sum = partCounts.reduce((a, b) => a + b, 0);

    if (sum !== totalDataRows) {
      return res.status(400).json({
        error: `Сумма частей (${sum}) не совпадает с количеством строк (${totalDataRows})`
      });
    }

    session.periodMode = periodMode;
    session.partCounts = partCounts;

    // Извлекаем и разделяем данные
    const dataRows = table.rows.slice(startRow).map(row => ({
      tema: (row[session.temaCol] || '').trim(),
      dz: session.noDz ? '' : (row[session.dzCol] || '').trim(),
      status: 'normal'
    }));

    if (defaultDzForAll === 'Конспект' || defaultDzForAll === 'Карточка') {
      dataRows.forEach((r) => {
        r.dz = defaultDzForAll;
      });
    }

    session.parts = partService.divideParts(dataRows, partCounts);

    console.log(`[parts] Разделено на ${session.parts.length} частей: ${partCounts.join(', ')}`);

    res.json({
      success: true,
      parts: session.parts.map((p, i) => ({
        id: i,
        rowCount: p.rows.length,
        label: (periodMode === 'quarters' ? 'Ч' : 'П') + (i + 1)
      }))
    });
  } catch (error) {
    console.error('[parts] Ошибка:', error);
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// GET /api/session/:id/part/:partIdx — Содержимое части
// ===================================================================
router.get('/session/:id/part/:partIdx', (req, res) => {
  try {
    const session = sessions.get(req.params.id);
    if (!session) {
      return res.status(404).json({ error: 'Сессия не найдена' });
    }

    const partIdx = parseInt(req.params.partIdx, 10);
    const part = session.parts[partIdx];
    if (!part) {
      return res.status(404).json({ error: 'Часть не найдена' });
    }

    res.json(part);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// POST /api/session/:id/part/:partIdx/update — Обновляет часть
// ===================================================================
router.post('/session/:id/part/:partIdx/update', (req, res) => {
  try {
    const session = sessions.get(req.params.id);
    if (!session) {
      return res.status(404).json({ error: 'Сессия не найдена' });
    }

    const partIdx = parseInt(req.params.partIdx, 10);
    const { rows } = req.body;

    if (!session.parts[partIdx]) {
      return res.status(404).json({ error: 'Часть не найдена' });
    }

    session.parts[partIdx].rows = rows;
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// POST /api/session/:id/part/:partIdx/compress — Сжатие программы
// ===================================================================
router.post('/session/:id/part/:partIdx/compress', (req, res) => {
  try {
    const session = sessions.get(req.params.id);
    if (!session) {
      return res.status(404).json({ error: 'Сессия не найдена' });
    }

    const partIdx = parseInt(req.params.partIdx, 10);
    const part = session.parts[partIdx];
    if (!part) {
      return res.status(404).json({ error: 'Часть не найдена' });
    }

    const { rows, warnings } = partService.compressPart(part.rows, session.separator);
    session.parts[partIdx].rows = rows;

    res.json({ success: true, warnings });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// POST /api/session/:id/export — Экспорт в Excel
// ===================================================================
router.post('/session/:id/export', async (req, res) => {
  try {
    const session = sessions.get(req.params.id);
    if (!session || !session.parts || session.parts.length === 0) {
      return res.status(404).json({ error: 'Данные не найдены' });
    }

    const uploadsDir = path.join(__dirname, '../../uploads');
    const files = await exportService.exportToExcel(
      session.parts,
      uploadsDir,
      session.periodMode
    );

    const downloads = files.map(filePath => ({
      fileName: path.basename(filePath),
      downloadUrl: `/api/download/${path.basename(filePath)}`
    }));

    console.log(`[export] Экспортировано ${files.length} файлов`);

    res.json({ success: true, files: downloads });
  } catch (error) {
    console.error('[export] Ошибка:', error);
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// GET /api/download/:fileName — Скачивание файла
// ===================================================================
router.get('/download/:fileName', (req, res) => {
  try {
    const fileName = req.params.fileName;
    if (fileName.includes('..') || fileName.includes('/') || fileName.includes('\\')) {
      return res.status(403).json({ error: 'Неверное имя файла' });
    }

    const uploadsDir = path.resolve(__dirname, '../../uploads');
    const filePath = path.join(uploadsDir, fileName);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'Файл не найден' });
    }

    res.download(filePath, fileName, (err) => {
      if (err && !res.headersSent) {
        console.error('[download] Ошибка:', err);
        res.status(500).json({ error: 'Ошибка при скачивании файла' });
      }
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// ===================================================================
// Вспомогательные функции
// ===================================================================
function autoDetectColumns(session, tableIdx) {
  session.temaCol = null;
  session.dzCol = null;

  const headerRow = session.tables[tableIdx].rows[0] || [];
  headerRow.forEach((cell, i) => {
    const low = (cell || '').toLowerCase();
    if (session.temaCol === null && (low.includes('тема') || low.includes('тем'))) {
      session.temaCol = i;
    }
    if (session.dzCol === null && (low.includes('домашн') || low.includes('задан') || low.includes('дз') || low.includes('д/з'))) {
      session.dzCol = i;
    }
  });
}

module.exports = router;
