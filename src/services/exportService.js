const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

/**
 * Экспортирует части в XLS и XLSX файлы
 * @param {Array} parts - Массив частей [{rows: [...]}]
 * @param {String} outputDir - Директория для сохранения файлов
 * @param {String} periodMode - 'quarters' или 'halves'
 * @returns {Promise<Array>} Массив путей к созданным файлам
 */
async function exportToExcel(parts, outputDir, periodMode = 'quarters') {
  const prefix = periodMode === 'quarters' ? 'Ч' : 'П';
  const files = [];

  // Убедимся, что директория существует
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  for (let i = 0; i < parts.length; i++) {
    const partRows = parts[i].rows;
    const fileName = `${prefix}${i + 1}`;

    // Создаём рабочую книгу
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Уроки');

    // Добавляем заголовки
    worksheet.addRow(['Тема урока', 'Домашнее задание']);

    // Стилизуем заголовки
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE7E6E6' }
    };

    // Добавляем данные
    partRows.forEach(row => {
      worksheet.addRow([row.tema, row.dz]);
    });

    // Настраиваем ширину столбцов
    worksheet.columns = [
      { width: 50 },  // Тема урока
      { width: 40 }   // ДЗ
    ];

    // Переносим текст в ячейках
    worksheet.eachRow((row, rowNumber) => {
      row.alignment = { wrapText: true, vertical: 'top' };
      // Устанавливаем минимальную высоту строки
      row.height = Math.max(20, row.values.reduce((max, val) => {
        if (!val) return max;
        const lines = String(val).split('\n').length;
        return Math.max(max, lines * 15);
      }, 20));
    });

    // Сохраняем в XLSX
    const xlsxPath = path.join(outputDir, `${fileName}.xlsx`);
    await workbook.xlsx.writeFile(xlsxPath);
    files.push(xlsxPath);

    // Сохраняем в XLS (BIFF8) — ExcelJS не поддерживает BIFF8 напрямую,
    // используем xlsx как фоллбэк или конвертим через ручное создание
    // Для XLS используем другой подход — просто дублируем с другим расширением
    // или используем более старую версию формата
    const xlsPath = path.join(outputDir, `${fileName}.xls`);
    // ExcelJS не может писать XLS (только XLSX), поэтому создадим XLS через xlsx библиотеку
    const xlsData = convertToXls(worksheet);
    fs.writeFileSync(xlsPath, xlsData);
    files.push(xlsPath);
  }

  return files;
}

/**
 * Конвертирует рабочий лист в XLS формат (BIFF8)
 * Используем xlsx библиотеку для этого
 */
function convertToXls(worksheet) {
  const XLSX = require('xlsx');

  // Преобразуем ExcelJS worksheet в стандартный формат
  const data = [];
  data.push(['Тема урока', 'Домашнее задание']);

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 1) {
      // Пропускаем заголовок, так как мы его добавили выше
      data.push([row.values[1] || '', row.values[2] || '']);
    }
  });

  // Создаём workbook в xlsx
  const workbook = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);

  // Устанавливаем ширину столбцов
  ws['!cols'] = [
    { wch: 50 },
    { wch: 40 }
  ];

  XLSX.utils.book_append_sheet(workbook, ws, 'Уроки');

  // Возвращаем буфер в формате XLS
  return XLSX.write(workbook, { bookType: 'biff8', type: 'buffer' });
}

/**
 * Удаляет файлы экспорта (для очистки)
 */
async function cleanupFiles(filePaths) {
  for (const filePath of filePaths) {
    try {
      if (fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
      }
    } catch (error) {
      console.error(`Не удалось удалить файл ${filePath}:`, error.message);
    }
  }
}

module.exports = {
  exportToExcel,
  cleanupFiles
};
