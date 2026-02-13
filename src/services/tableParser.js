const mammoth = require('mammoth');
const JSZip = require('jszip');

/**
 * Парсит DOCX файл и извлекает таблицы
 */
async function parseDocx(buffer) {
  try {
    const result = await mammoth.convertToHtml({ buffer });
    const html = result.value;

    if (!html || !html.includes('<table')) {
      console.log('[parseDocx] HTML не содержит таблиц');
      return [];
    }

    const tables = [];
    // Используем простой regex-парсер
    const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
    let tableMatch;

    while ((tableMatch = tableRegex.exec(html)) !== null) {
      const tableHtml = tableMatch[0];
      const tbl = parseHtmlTable(tableHtml);
      if (tbl.rows.length > 0) {
        tables.push(tbl);
      }
    }

    console.log(`[parseDocx] Найдено таблиц: ${tables.length}`);
    return tables;
  } catch (error) {
    throw new Error(`Ошибка при парсинге DOCX: ${error.message}`);
  }
}

/**
 * Парсит ODT файл
 */
async function parseOdt(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const contentXmlFile = zip.file('content.xml');

    if (!contentXmlFile) {
      throw new Error('content.xml не найден в ODT файле');
    }

    const contentXml = await contentXmlFile.async('string');
    return parseOdtXml(contentXml);
  } catch (error) {
    throw new Error(`Ошибка при парсинге ODT: ${error.message}`);
  }
}

/**
 * Парсит HTML таблицу через regex (работает на сервере без DOM)
 */
function parseHtmlTable(htmlStr) {
  const rows = [];
  const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let rowMatch;

  while ((rowMatch = rowRegex.exec(htmlStr)) !== null) {
    const rowHtml = rowMatch[1];
    const cells = [];
    const cellRegex = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    let cellMatch;

    while ((cellMatch = cellRegex.exec(rowHtml)) !== null) {
      let content = cellMatch[1];
      // Заменяем <p>, <br> на переносы строк
      content = content.replace(/<br\s*\/?>/gi, '\n');
      content = content.replace(/<\/p>\s*<p[^>]*>/gi, '\n');
      // Удаляем все HTML теги
      content = content.replace(/<[^>]*>/g, '');
      // Декодируем entities
      content = decodeEntities(content);
      content = content.trim();
      cells.push(content);
    }

    if (cells.length > 0) {
      rows.push(cells);
    }
  }

  return { rows };
}

/**
 * Парсит ODT XML
 */
function parseOdtXml(xmlString) {
  const { DOMParser } = require('@xmldom/xmldom');
  const doc = new DOMParser().parseFromString(xmlString, 'application/xml');

  const NS_TABLE = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';
  const NS_TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';

  const tableEls = doc.getElementsByTagNameNS(NS_TABLE, 'table');
  const tables = [];

  for (let ti = 0; ti < tableEls.length; ti++) {
    const tableEl = tableEls[ti];
    const rows = [];
    const rowEls = tableEl.getElementsByTagNameNS(NS_TABLE, 'table-row');

    for (let ri = 0; ri < rowEls.length; ri++) {
      const rowEl = rowEls[ri];
      const cells = [];
      const cellEls = rowEl.getElementsByTagNameNS(NS_TABLE, 'table-cell');

      for (let ci = 0; ci < cellEls.length; ci++) {
        const cellEl = cellEls[ci];
        const repeated = parseInt(
          cellEl.getAttributeNS(NS_TABLE, 'number-columns-repeated') || '1', 10
        );

        const pEls = cellEl.getElementsByTagNameNS(NS_TEXT, 'p');
        const parts = [];
        for (let pi = 0; pi < pEls.length; pi++) {
          const txt = (pEls[pi].textContent || '').trim();
          if (txt) parts.push(txt);
        }
        const cellText = parts.join('\n');

        const reps = Math.min(repeated, 30);
        for (let r = 0; r < reps; r++) {
          cells.push(cellText);
        }
      }

      if (cells.some(c => c.trim() !== '')) {
        rows.push(cells);
      }
    }

    if (rows.length > 0) {
      tables.push({ rows });
    }
  }

  console.log(`[parseOdt] Найдено таблиц: ${tables.length}`);
  return tables;
}

/**
 * Декодирует HTML entities
 */
function decodeEntities(text) {
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, (_, n) => String.fromCharCode(parseInt(n, 10)));
}

module.exports = { parseDocx, parseOdt };
