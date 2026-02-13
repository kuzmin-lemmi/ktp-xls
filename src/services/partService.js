/**
 * Сервис для работы с частями КТП (четверти/полугодия)
 */

/**
 * Разделяет строки таблицы на части
 * @param {Array} rows - Строки таблицы [{tema, dz, status}]
 * @param {Array} partCounts - Количество уроков в каждой части [12, 10, 15, 14]
 * @returns {Array} Массив частей
 */
function divideParts(rows, partCounts) {
  const parts = [];
  let offset = 0;

  for (let i = 0; i < partCounts.length; i++) {
    const count = partCounts[i];
    const partRows = rows.slice(offset, offset + count).map(row => ({
      ...row,
      status: 'normal'
    }));
    parts.push({ rows: partRows });
    offset += count;
  }

  return parts;
}

/**
 * Сжимает программу: переносит темы из отменённых уроков в следующие
 * Работает в пределах одной части
 * @param {Array} partRows - Строки одной части
 * @param {String} separator - Разделитель между объединениями
 * @returns {Object} {rows: compressed, warnings: []}
 */
function compressPart(partRows, separator = '\n') {
  const compressed = partRows.map(r => ({ ...r })); // Deep copy
  let debtTema = [];
  let debtDz = [];
  const warnings = [];

  for (let i = 0; i < compressed.length; i++) {
    const row = compressed[i];

    if (row.status === 'cancelled') {
      // Отменённая строка остаётся пустой, но её содержимое
      // считается долгом (если что-то было до отмены)
      // На самом деле, при отмене уже очищаем, так что долга нет
      continue;
    }

    // Обычный урок: добавляем к нему долг
    if (debtTema.length > 0 || debtDz.length > 0) {
      const newTema = [row.tema, ...debtTema].filter(t => t).join(separator);
      const newDz = [row.dz, ...debtDz].filter(d => d).join(separator);
      row.tema = newTema;
      row.dz = newDz;
      debtTema = [];
      debtDz = [];
    }
  }

  // Если остался долг — это ошибка
  if (debtTema.length > 0 || debtDz.length > 0) {
    warnings.push(
      `Осталось ${debtTema.length} тем, которые не уместились в этой части`
    );
  }

  return { rows: compressed, warnings };
}

/**
 * Перемещает строку с данными на позицию ниже и объединяет содержимое
 * @param {Array} partRows - Строки части
 * @param {Number} fromIdx - Индекс исходной строки
 * @param {Number} toIdx - Индекс целевой строки
 * @param {String} separator - Разделитель
 */
function dragRowDown(partRows, fromIdx, toIdx, separator = '\n') {
  if (toIdx <= fromIdx) {
    throw new Error('Можно перетаскивать только вниз');
  }

  const srcRow = partRows[fromIdx];
  const dstRow = partRows[toIdx];

  // Объединяем содержимое в целевой строке
  if (dstRow.tema || dstRow.dz) {
    dstRow.tema = [dstRow.tema, srcRow.tema].filter(t => t).join(separator);
    dstRow.dz = [dstRow.dz, srcRow.dz].filter(d => d).join(separator);
  } else {
    dstRow.tema = srcRow.tema;
    dstRow.dz = srcRow.dz;
  }

  // Исходная строка становится пустой
  srcRow.tema = '';
  srcRow.dz = '';
  srcRow.status = 'normal';

  // Перемещаем в массиве
  partRows.splice(fromIdx, 1);
  partRows.splice(toIdx, 0, srcRow);
}

/**
 * Отменяет строку (помечает как cancelled и очищает содержимое)
 */
function cancelRow(row) {
  row.status = 'cancelled';
  row.tema = '';
  row.dz = '';
}

/**
 * Восстанавливает строку
 */
function restoreRow(row) {
  row.status = 'normal';
  // Содержимое остаётся пустым — пользователь может заполнить вручную
}

/**
 * Вставляет шаблон в поле ДЗ для выделенных строк
 */
function insertTemplate(rows, selectedIndices, template) {
  selectedIndices.forEach(idx => {
    if (rows[idx]) {
      rows[idx].dz = template;
    }
  });
}

module.exports = {
  divideParts,
  compressPart,
  dragRowDown,
  cancelRow,
  restoreRow,
  insertTemplate
};
