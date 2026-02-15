import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import mammoth from 'mammoth/mammoth.browser';

const DRAFT_KEY = 'ktp-converter-draft-v1';

const state = {
  fileName: null,
  rawTables: [],
  selectedTableIdx: null,
  temaCol: null,
  dzCol: null,
  noDz: false,
  hasHeader: true,
  separator: '\n',
  periodMode: 'quarters',
  partCounts: [],
  parts: [],
  totalDataRows: 0,
  activeTab: 0,
  selectedRows: new Set(),
  currentStep: 1,
  previewHeaders: [],
  previewRows: [],
  defaultDzForAll: '',
  currentTemplate: 'Конспект',
  lastClickedRow: null
};

function cloneParts(parts) {
  return parts.map(p => ({ rows: p.rows.map(r => ({ ...r })) }));
}

function scheduleDraftSave() {
  if (state.currentStep !== 4 || !state.parts.length) return;
  clearTimeout(scheduleDraftSave._t);
  scheduleDraftSave._t = setTimeout(() => {
    try {
      localStorage.setItem(DRAFT_KEY, JSON.stringify({
        v: 1, ts: Date.now(),
        fileName: state.fileName,
        periodMode: state.periodMode,
        partCounts: state.partCounts,
        separator: state.separator,
        noDz: state.noDz,
        currentTemplate: state.currentTemplate,
        activeTab: state.activeTab,
        parts: cloneParts(state.parts)
      }));
    } catch {}
  }, 200);
}

function restoreDraftIfExists() {
  try {
    const raw = localStorage.getItem(DRAFT_KEY);
    if (!raw) return false;
    const d = JSON.parse(raw);
    if (!d || !Array.isArray(d.parts) || !d.parts.length) return false;
    if (!window.confirm(`Найден черновик${d.fileName ? ` (${d.fileName})` : ''}. Восстановить последний прогресс?`)) {
      localStorage.removeItem(DRAFT_KEY);
      return false;
    }
    Object.assign(state, {
      fileName: d.fileName || null,
      periodMode: d.periodMode === 'halves' ? 'halves' : 'quarters',
      partCounts: d.partCounts || [],
      separator: d.separator || '\n',
      noDz: !!d.noDz,
      currentTemplate: d.currentTemplate || 'Конспект',
      parts: cloneParts(d.parts),
      activeTab: Number.isInteger(d.activeTab) ? d.activeTab : 0
    });
    return true;
  } catch { return false; }
}

function getPartStats(partIdx) {
  const rows = state.parts[partIdx].rows;
  const total = rows.length;
  const cancelled = rows.filter(r => r.status === 'cancelled').length;
  const emptyDz = rows.filter(r => r.status !== 'cancelled' && !r.dz.trim()).length;
  return { total, cancelled, emptyDz, active: total - cancelled };
}

function syncTabStats() {
  const p = state.periodMode === 'quarters' ? 'Ч' : 'П';
  document.querySelectorAll('.tab-btn').forEach((btn, i) => {
    if (i >= state.parts.length) return;
    const s = getPartStats(i);
    const warn = s.cancelled > 0 || s.emptyDz > 0;
    btn.innerHTML = `${p}${i + 1} <span class="tab-stat${warn ? ' tab-stat-warn' : ''}">${s.active}/${s.total}</span>`;
    btn.classList.toggle('active', i === state.activeTab);
  });
}

function clearDraft() {
  if (!window.confirm('Очистить черновик? Текущие правки останутся, но автосохранение будет сброшено.')) return;
  try { localStorage.removeItem(DRAFT_KEY); } catch {}
  showInfo('Черновик удалён');
}

/**
 * INIT APP
 */
function initApp() {
  document.getElementById('root').innerHTML = `
    <header>
      <div class="logo">📚</div>
      <div>
        <h1>Конвертер КТП &rarr; Excel</h1>
        <span>Версия для Дневник.ру</span>
      </div>
    </header>
    <div class="wizard">
      <div class="stepper" id="stepper"></div>
      <div id="steps-container"></div>
    </div>`;
  renderStepper();
  if (restoreDraftIfExists()) renderStep(4);
  else renderStep(1);
}

/**
 * STEPPER
 */
function renderStepper() {
  const el = document.getElementById('stepper');
  const steps = [
    { n: 1, l: 'Файл' },
    { n: 2, l: 'Столбцы' },
    { n: 3, l: 'Части' },
    { n: 4, l: 'Редактор' }
  ];
  let h = '';
  steps.forEach((s, i) => {
    const active = state.currentStep === s.n;
    const done = state.currentStep > s.n;
    h += `
      <div class="step${active ? ' active' : ''}${done ? ' done' : ''}">
        <div class="step-num">${done ? '✓' : s.n}</div>
        <div class="step-label">${s.l}</div>
      </div>`;
    if (i < steps.length - 1) {
      h += `<div class="step-line${done ? ' done' : ''}"></div>`;
    }
  });
  el.innerHTML = h;
}

function renderStep(n) {
  state.currentStep = n;
  renderStepper();
  const c = document.getElementById('steps-container');
  c.innerHTML = '';
  if (n === 1) renderStep1(c);
  else if (n === 2) renderStep2(c);
  else if (n === 3) renderStep3(c);
  else if (n === 4) renderStep4(c);
}

/**
 * STEP 1 — FILE UPLOAD
 */
function renderStep1(c) {
  c.innerHTML = `
    <div class="card">
      <h2>Шаг 1. Загрузите файл КТП</h2>
      <p class="sub">Выберите ваш файл КТП (.docx или .odt). Файл обрабатывается прямо в браузере.</p>

      <div class="info-box">
        <h4><span style="font-size:1.2rem">💡</span> Как подготовить файл?</h4>
        <p>Нужен файл с <b>таблицей</b>, где есть темы уроков и (желательно) домашнее задание.</p>
        <details style="margin-top:8px">
          <summary style="cursor:pointer;font-weight:700;color:var(--primary)">Посмотреть пример правильной таблицы</summary>
          <table class="example-table">
            <thead><tr><th>№</th><th>Тема урока</th><th>Домашнее задание</th></tr></thead>
            <tbody>
              <tr><td>1</td><td>Техника безопасности</td><td>Параграф 1</td></tr>
              <tr><td>2</td><td>Устройство компьютера</td><td>Конспект</td></tr>
            </tbody>
          </table>
        </details>
      </div>

      <div class="dropzone" id="dropzone">
        <div class="dropzone-icon">📄</div>
        <h3>Перетащите файл сюда</h3>
        <p>или нажмите для выбора (.docx, .odt)</p>
      </div>
      <input type="file" id="file-input" accept=".docx,.odt">

      <div id="file-result" class="hidden">
        <div class="file-badge" id="file-badge"></div>
        <div style="margin-top:16px">
          <div style="font-size:0.9rem;font-weight:700;color:var(--slate-600);margin-bottom:8px">Выберите таблицу из списка:</div>
          <div class="table-list" id="table-list"></div>
        </div>
      </div>
      <div id="step1-error"></div>
      <div class="actions">
        <div></div>
        <button class="btn btn-primary btn-lg" id="btn-next1" disabled>Далее &rarr;</button>
      </div>
    </div>`;

  const dz = document.getElementById('dropzone');
  const fi = document.getElementById('file-input');

  dz.onclick = () => fi.click();
  dz.ondragover = e => { e.preventDefault(); dz.classList.add('drag-over'); };
  dz.ondragleave = () => dz.classList.remove('drag-over');
  dz.ondrop = e => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  };
  fi.onchange = () => { if (fi.files[0]) handleFile(fi.files[0]); };

  document.getElementById('btn-next1').onclick = () => {
    if (state.selectedTableIdx !== null) renderStep(2);
  };
}

async function handleFile(file) {
  const dz = document.getElementById('dropzone');
  dz.innerHTML = `
    <div style="padding:24px;display:flex;align-items:center;justify-content:center;gap:12px">
      <div class="loading"></div>
      <span style="font-weight:600">Читаем документ...</span>
    </div>`;

  try {
    const ext = file.name.split('.').pop().toLowerCase();
    let tables = [];
    if (ext === 'docx') tables = await parseDocx(file);
    else if (ext === 'odt') tables = await parseOdt(file);
    else throw new Error('Неподдерживаемый формат. Используйте Word (.docx) или LibreOffice (.odt)');

    if (!tables.length) throw new Error('В файле не найдено ни одной таблицы.');

    state.fileName = file.name;
    state.rawTables = tables;
    state.selectedTableIdx = tables.length === 1 ? 0 : null;
    autoDetectColumns();

    dz.innerHTML = '<div style="padding:24px;color:var(--success);font-weight:700">✅ Файл успешно прочитан</div>';
    document.getElementById('file-result').classList.remove('hidden');
    document.getElementById('file-badge').textContent = `✅ ${file.name}`;
    renderTableList();
  } catch (e) {
    dz.innerHTML = `
      <div class="error-msg">
        <b>Ошибка:</b> ${esc(e.message)}
        <br><small>Попробуйте открыть файл в Word и пересохранить его.</small>
      </div>`;
  }
}

function renderTableList() {
  const list = document.getElementById('table-list');
  list.innerHTML = '';

  state.rawTables.forEach((t, i) => {
    const isSelected = state.selectedTableIdx === i;
    const totalRows = t.rows.length;
    const cols = Math.max(...t.rows.map(r => r.length));
    const firstRow = t.rows[0] || [];
    const lastRows = t.rows.slice(-3);

    // Определяем похоже ли первая строка на заголовок
    const looksLikeHeader = firstRow.some(cell => {
      const low = String(cell || '').toLowerCase();
      return low.includes('тема') || low.includes('урок') || low.includes('дз') || low.includes('домашн') || low.includes('№');
    });

    const dataRows = totalRows - 1; // если есть заголовок
    const lessonWord = (n) => n === 1 ? 'урок' : (n >= 2 && n <= 4 ? 'урока' : 'уроков');

    // Превью первой строки
    const firstRowPreview = firstRow.slice(0, 4).map(c => `<span class="row-preview-cell${looksLikeHeader ? ' header-cell' : ''}">${esc(String(c || '').substring(0, 30))}</span>`).join('');

    // Превью последних строк
    const lastRowsPreview = lastRows.map((r, ri) => {
      const isEmpty = r.every(c => !String(c || '').trim());
      const isLast = ri === lastRows.length - 1;
      return `<div class="row-preview-line${isEmpty ? ' empty-row' : ''}${isLast ? ' last-row' : ''}">${r.slice(0, 4).map(c => `<span class="row-preview-cell">${esc(String(c || '').substring(0, 30))}</span>`).join('')}${isEmpty ? '<span class="empty-badge">пустая строка</span>' : ''}</div>`;
    }).join('');

    const d = document.createElement('div');
    d.className = `table-item${isSelected ? ' selected' : ''}`;
    const dataRowsActual = state.hasHeader ? totalRows - 1 : totalRows;
    d.innerHTML = `
      <div class="table-item-header">
        <div>
          <b>Таблица ${i + 1}</b>
          <span class="tbl-stat">всего строк: <b>${totalRows}</b></span>
          <span class="tbl-stat tbl-stat-lessons">уроков: <b>${dataRowsActual}</b> ${lessonWord(dataRowsActual)}</span>
          <span class="tbl-stat tbl-stat-cols">${cols} столбцов</span>
        </div>
      </div>
      <div class="table-item-preview">
        <div class="preview-section-label">Первая строка ${looksLikeHeader ? '<span class="badge-header">похожа на заголовок ✓</span>' : '<span class="badge-noheader">может быть урок</span>'}:</div>
        <div class="row-preview-line header-line">${firstRowPreview}</div>
        <div class="preview-section-label" style="margin-top:8px">Последние строки (проверьте на «мусор»):</div>
        ${lastRowsPreview}
      </div>
      <div class="table-item-footer" onclick="event.stopPropagation()">
        <label class="header-checkbox-label">
          <input type="checkbox" class="header-checkbox" data-idx="${i}" ${state.hasHeader ? 'checked' : ''}>
          <span>Первая строка — заголовки (не считать уроком)</span>
        </label>
      </div>`;

    d.onclick = () => {
      state.selectedTableIdx = i;
      autoDetectColumns();
      renderTableList();
      document.getElementById('btn-next1').disabled = false;
    };
    list.appendChild(d);
  });

  document.getElementById('btn-next1').disabled = state.selectedTableIdx === null;

  // Обработчики галочек "первая строка — заголовок"
  document.querySelectorAll('.header-checkbox').forEach(cb => {
    cb.onchange = (e) => {
      state.hasHeader = e.target.checked;
      // синхронизируем все остальные галочки
      document.querySelectorAll('.header-checkbox').forEach(x => x.checked = state.hasHeader);
      // пересчитываем счётчики уроков без полного ре-рендера
      document.querySelectorAll('.tbl-stat-lessons').forEach((el, idx) => {
        if (idx >= state.rawTables.length) return;
        const total = state.rawTables[idx].rows.length;
        const cnt = state.hasHeader ? total - 1 : total;
        const lw = cnt === 1 ? 'урок' : (cnt >= 2 && cnt <= 4 ? 'урока' : 'уроков');
        el.innerHTML = `уроков: <b>${cnt}</b> ${lw}`;
      });
    };
  });
}

function autoDetectColumns() {
  state.temaCol = null;
  state.dzCol = null;
  if (state.selectedTableIdx === null) return;
  const header = state.rawTables[state.selectedTableIdx].rows[0] || [];
  header.forEach((c, i) => {
    const low = String(c || '').toLowerCase();
    if (state.temaCol === null && (low.includes('тема') || low.includes('тем'))) state.temaCol = i;
    if (state.dzCol === null && (low.includes('дом') || low.includes('д/з') || low.includes('дз') || low.includes('задан'))) state.dzCol = i;
  });
}

/**
 * STEP 2 — COLUMN SELECTION
 */
function renderStep2(c) {
  c.innerHTML = `
    <div class="card">
      <h2>Шаг 2. Выберите столбцы</h2>
      <p class="sub">Нажмите на заголовки столбцов в таблице ниже, чтобы отметить «Тему» и «ДЗ».</p>

      <div id="col-status" class="col-status waiting">Инициализация...</div>

      <div class="preview-wrap">
        <table class="preview-table" id="preview-table"></table>
      </div>

      <div style="display:flex;flex-direction:column;gap:4px">
        <div class="option-row">
          <input type="checkbox" id="no-dz" ${state.noDz ? 'checked' : ''}>
          <label for="no-dz"><b>В этом КТП нет столбца ДЗ</b> (сделать его пустым)</label>
        </div>
        <div class="option-row">
          <label for="separator">Разделитель тем:</label>
          <select id="separator" class="styled">
            <option value="\n">Перенос строки</option>
            <option value="; ">Точка с запятой</option>
            <option value=" / ">Косая черта /</option>
          </select>
        </div>
      </div>

      <div class="actions">
        <button class="btn btn-secondary btn-lg" id="btn-back2">← Назад</button>
        <button class="btn btn-primary btn-lg" id="btn-next2" disabled>Далее →</button>
      </div>
    </div>`;

  document.getElementById('btn-back2').onclick = () => renderStep(1);
  document.getElementById('no-dz').onchange = (e) => { 
    state.noDz = e.target.checked; 
    if (state.noDz) state.dzCol = null; 
    drawPreview(); 
    updateColStatus(); 
  };
  document.getElementById('separator').onchange = (e) => { state.separator = e.target.value; };
  document.getElementById('btn-next2').onclick = () => renderStep(3);

  loadPreview();
}

function loadPreview() {
  const table = state.rawTables[state.selectedTableIdx];
  state.previewHeaders = table.rows[0] || [];
  state.previewRows = table.rows.slice(0, 12);
  state.totalDataRows = Math.max(0, table.rows.length - (state.hasHeader ? 1 : 0));
  drawPreview();
  updateColStatus();
}

function drawPreview() {
  const tbl = document.getElementById('preview-table');
  let h = '<thead><tr>';
  state.previewHeaders.forEach((hdr, i) => {
    let cls = '', badge = '';
    if (i === state.temaCol) { cls = ' col-tema'; badge = '<span class="col-badge badge-tema">ТЕМА</span>'; }
    if (!state.noDz && i === state.dzCol) { cls = ' col-dz'; badge = '<span class="col-badge badge-dz">ДЗ</span>'; }
    h += `<th class="${cls}" onclick="window.pickCol(${i})">${esc(hdr)}${badge}</th>`;
  });
  h += '</tr></thead><tbody>';
  state.previewRows.forEach((r, ri) => {
    if (ri === 0 && state.hasHeader) return;
    h += '<tr>';
    r.forEach((cell, ci) => {
      let cls = '';
      if (ci === state.temaCol) cls = ' col-tema';
      if (!state.noDz && ci === state.dzCol) cls = ' col-dz';
      h += `<td class="${cls}">${esc(cell || '')}</td>`;
    });
    h += '</tr>';
  });
  h += '</tbody>';
  tbl.innerHTML = h;
}

window.pickCol = function(i) {
  if (state.temaCol === null) state.temaCol = i;
  else if (state.temaCol === i) state.temaCol = null;
  else if (!state.noDz && state.dzCol === null && i !== state.temaCol) state.dzCol = i;
  else if (!state.noDz && state.dzCol === i) state.dzCol = null;
  else state.temaCol = i;
  drawPreview();
  updateColStatus();
};

function updateColStatus() {
  const el = document.getElementById('col-status');
  const btn = document.getElementById('btn-next2');
  if (state.temaCol === null) {
    el.className = 'col-status waiting';
    el.innerHTML = '👉 Нажмите на столбец <b>«Тема урока»</b>';
    btn.disabled = true;
  } else if (!state.noDz && state.dzCol === null) {
    el.className = 'col-status waiting';
    el.innerHTML = '✅ Тема выбрана. Теперь выберите <b>«Домашнее задание»</b>';
    btn.disabled = true;
  } else {
    el.className = 'col-status ready';
    el.innerHTML = state.noDz ? '✅ Всё готово для экспорта без ДЗ' : '✅ Столбцы выбраны. Нажмите «Далее»';
    btn.disabled = false;
  }
}

/**
 * STEP 3 — PARTS
 */
function renderStep3(c) {
  c.innerHTML = `
    <div class="card">
      <h2>Шаг 3. Разделите на части</h2>
      <p class="sub">Укажите количество уроков в каждой четверти или полугодии. <br>Всего в КТП: <b>${state.totalDataRows}</b> уроков.</p>
      
      <div class="radio-group">
        <div class="radio-btn ${state.periodMode === 'quarters' ? 'active' : ''}" id="rb-q">📅 4 четверти</div>
        <div class="radio-btn ${state.periodMode === 'halves' ? 'active' : ''}" id="rb-h">🗓 2 полугодия</div>
      </div>
      
      <div class="parts-grid" id="parts-grid"></div>
      
      <div id="count-info" class="count-info bad"></div>
      
      <div class="actions">
        <button class="btn btn-secondary btn-lg" id="btn-back3">← Назад</button>
        <button class="btn btn-primary btn-lg" id="btn-next3" disabled>Далее →</button>
      </div>
    </div>`;

  document.getElementById('btn-back3').onclick = () => renderStep(2);
  document.getElementById('rb-q').onclick = () => { state.periodMode = 'quarters'; state.partCounts = [0,0,0,0]; renderStep(3); };
  document.getElementById('rb-h').onclick = () => { state.periodMode = 'halves'; state.partCounts = [0,0]; renderStep(3); };
  document.getElementById('btn-next3').onclick = buildPartsNext;

  if (!state.partCounts.length) state.partCounts = state.periodMode === 'quarters' ? [0,0,0,0] : [0,0];
  buildPartsGrid();
}

function buildPartsGrid() {
  const n = state.periodMode === 'quarters' ? 4 : 2;
  const p = state.periodMode === 'quarters' ? 'Ч' : 'П';
  const grid = document.getElementById('parts-grid');
  grid.innerHTML = '';
  for (let i = 0; i < n; i++) {
    const d = document.createElement('div');
    d.className = 'part-input';
    d.innerHTML = `<label>${p}${i + 1} — количество уроков:</label><input id="pc-${i}" type="number" min="0" value="${state.partCounts[i] || ''}" placeholder="0">`;
    grid.appendChild(d);
    d.querySelector('input').addEventListener('input', recalcParts);
  }
  recalcParts();
}

function recalcParts() {
  const n = state.periodMode === 'quarters' ? 4 : 2;
  let sum = 0;
  for (let i = 0; i < n; i++) {
    const v = parseInt(document.getElementById(`pc-${i}`)?.value || '0', 10);
    state.partCounts[i] = Number.isNaN(v) ? 0 : v;
    sum += state.partCounts[i];
  }
  const total = state.totalDataRows;
  const remaining = total - sum;
  const ok = sum === total && sum > 0;
  const info = document.getElementById('count-info');
  info.className = `count-info ${ok ? 'ok' : 'bad'}`;

  let statusHtml = '';
  if (ok) {
    statusHtml = `<span class="ci-item ci-ok">✅ Совпадает</span>`;
  } else if (remaining > 0) {
    statusHtml = `<span class="ci-item ci-warn">Осталось распределить: <b>${remaining}</b></span>`;
  } else {
    statusHtml = `<span class="ci-item ci-err">Превышение на: <b>${Math.abs(remaining)}</b></span>`;
  }

  info.innerHTML = `
    <span class="ci-item">В КТП: <b>${total}</b></span>
    <span class="ci-item">Распределено: <b>${sum}</b></span>
    ${statusHtml}`;
  document.getElementById('btn-next3').disabled = !ok;
}

function buildPartsNext() {
  const counts = state.partCounts;
  const table = state.rawTables[state.selectedTableIdx];
  const start = state.hasHeader ? 1 : 0;
  const rows = table.rows.slice(start).map(r => ({
    tema: String(r[state.temaCol] || '').trim(),
    dz: state.noDz ? '' : String(r[state.dzCol] || '').trim(),
    status: 'normal'
  }));

  state.parts = [];
  let offset = 0;
  counts.forEach(cnt => {
    state.parts.push({ rows: rows.slice(offset, offset + cnt) });
    offset += cnt;
  });
  renderStep(4);
}

/**
 * STEP 4 — EDITOR
 */
function renderStep4(c) {
  const n = state.parts.length;
  const p = state.periodMode === 'quarters' ? 'Ч' : 'П';
  let tabs = '';
  for (let i = 0; i < n; i++) {
    const s = getPartStats(i);
    const warn = s.cancelled > 0 || s.emptyDz > 0;
    tabs += `<button class="tab-btn${i === 0 ? ' active' : ''}" data-ti="${i}">${p}${i + 1} <span class="tab-stat${warn ? ' tab-stat-warn' : ''}">${s.active}/${s.total}</span></button>`;
  }

  c.innerHTML = `
    <div class="card">
      <h2>Шаг 4. Редактор и экспорт</h2>
      <p class="sub">Проверьте данные. Темы можно перетаскивать мышкой для объединения уроков.</p>
      
      <div class="tabs" id="tabs">${tabs}</div>
      
      <div class="editor-toolbar">
        <button class="btn btn-danger btn-sm" id="btn-cancel">✖ Отменить урок</button>
        <button class="btn btn-secondary btn-sm" id="btn-restore">↩ Вернуть урок</button>
        <button class="btn btn-secondary btn-sm" id="btn-clear-draft" title="Очистить сохранённый черновик">🗑 Черновик</button>
        <div class="sep"></div>
        <span style="font-size:0.8rem;color:var(--slate-500);font-weight:700">ШАБЛОНЫ:</span>
        <button class="tpl-btn" data-tpl="Конспект">Конспект</button>
        <button class="tpl-btn" data-tpl="Карточка">Карточка</button>
        <button class="tpl-btn" data-tpl="§ ">§...</button>
        <button class="tpl-btn" data-tpl="Упр. ">Упр...</button>
        <div class="sep"></div>
        <button class="btn btn-secondary btn-sm" id="btn-all-part">Установить всем в части</button>
        <button class="btn btn-secondary btn-sm" id="btn-all-all">Установить всем во всех частях</button>
        <div id="tpl-state" style="font-size:0.75rem;color:var(--primary);font-weight:700;margin-left:auto"></div>
      </div>

      <div id="warn-area"></div>
      <div id="editor-content"></div>
      
      <div class="actions">
        <button class="btn btn-secondary btn-lg" id="btn-back4">← Назад</button>
        <button class="btn btn-success btn-lg" id="btn-export">💾 Экспорт в Excel (.xls)</button>
      </div>
    </div>`;

  state.activeTab = 0;
  state.selectedRows = new Set();
  
  document.getElementById('tabs').onclick = e => {
    const b = e.target.closest('.tab-btn');
    if (!b) return;
    state.activeTab = parseInt(b.dataset.ti, 10);
    state.selectedRows = new Set();
    document.querySelectorAll('.tab-btn').forEach((x, i) => x.classList.toggle('active', i === state.activeTab));
    drawPartTable();
  };

  document.getElementById('btn-back4').onclick = () => renderStep(3);
  document.getElementById('btn-cancel').onclick = cancelSelected;
  document.getElementById('btn-restore').onclick = restoreSelected;
  document.getElementById('btn-clear-draft').onclick = clearDraft;
  document.getElementById('btn-all-part').onclick = applyTemplateToPart;
  document.getElementById('btn-all-all').onclick = applyTemplateToAllParts;
  document.getElementById('btn-export').onclick = doExport;
  
  document.querySelectorAll('.tpl-btn').forEach(b => {
    if (b.dataset.tpl === state.currentTemplate) b.classList.add('active-tpl');
    b.onclick = () => setTemplate(b.dataset.tpl);
  });

  updateTemplateState();
  drawPartTable();
}

function captureScroll() {
  const w = document.querySelector('.editor-wrap');
  return { top: w ? w.scrollTop : 0, left: w ? w.scrollLeft : 0 };
}

function restoreScroll(snap) {
  if (!snap) return;
  const w = document.querySelector('.editor-wrap');
  if (w) { w.scrollTop = snap.top; w.scrollLeft = snap.left; }
}

function drawPartTable(opts = {}) {
  const snap = opts.keepScroll ? captureScroll() : null;
  const part = state.parts[state.activeTab];
  const ec = document.getElementById('editor-content');
  let h = `
    <div class="editor-wrap">
      <table class="editor-table">
        <thead>
          <tr>
            <th style="width:50px">№</th>
            <th>Тема урока (тяните для объединения)</th>
            <th>Домашнее задание</th>
            <th style="width:100px">Статус</th>
            <th style="width:120px">Действия</th>
          </tr>
        </thead>
        <tbody id="editor-tbody">
          <tr><td colspan="5" class="drag-hint">💡 Перетащите тему урока на другой урок чтобы объединить. Shift+клик — выделить диапазон.</td></tr>`;
  
  part.rows.forEach((r, i) => {
    const cancelled = r.status === 'cancelled';
    const actBtns = cancelled
      ? `<button class="row-act-btn" data-act="restore-row" data-row="${i}" title="Вернуть">↩</button>`
      : `<button class="row-act-btn" data-act="cancel-row" data-row="${i}" title="Отменить">✖</button><button class="row-act-btn" data-act="copy-dz-down" data-row="${i}" title="Копировать ДЗ вниз">↓</button>`;
    h += `
      <tr data-row="${i}" class="${cancelled ? 'cancelled' : ''}${state.selectedRows.has(i) ? ' selected-row' : ''}" draggable="${!cancelled}">
        <td class="td-num">${i + 1}</td>
        <td class="td-tema" data-f="tema" data-r="${i}">${cancelled ? '&nbsp;' : esc(r.tema)}</td>
        <td class="td-dz" data-f="dz" data-r="${i}">${cancelled ? '&nbsp;' : esc(r.dz)}</td>
        <td style="text-align:center">${cancelled ? '<span class="cancelled-badge">ОТМЕНЁН</span>' : ''}</td>
        <td class="row-actions">${actBtns}</td>
      </tr>`;
  });
  h += '</tbody></table></div>';
  ec.innerHTML = h;
  setupEditorEvents(part);
  restoreScroll(snap);
  syncTabStats();
  scheduleDraftSave();
}

function setupEditorEvents(part) {
  const tbody = document.getElementById('editor-tbody');

  tbody.onclick = (e) => {
    if (e.target.closest('[contenteditable]')) return;
    // Кнопки строки
    const actBtn = e.target.closest('.row-act-btn');
    if (actBtn) {
      const idx = parseInt(actBtn.dataset.row, 10);
      if (Number.isNaN(idx) || !part.rows[idx]) return;
      if (actBtn.dataset.act === 'cancel-row') {
        preserveForRestore(part.rows[idx]);
        part.rows[idx].status = 'cancelled';
        part.rows[idx].tema = ''; part.rows[idx].dz = '';
        drawPartTable({ keepScroll: true });
      } else if (actBtn.dataset.act === 'restore-row') {
        part.rows[idx].status = 'normal';
        if (!part.rows[idx].tema && part.rows[idx]._backupTema) part.rows[idx].tema = part.rows[idx]._backupTema;
        if (!part.rows[idx].dz && part.rows[idx]._backupDz) part.rows[idx].dz = part.rows[idx]._backupDz;
        drawPartTable({ keepScroll: true });
      } else if (actBtn.dataset.act === 'copy-dz-down') {
        if (idx + 1 >= part.rows.length) { showInfo('Ниже нет строки'); return; }
        if (part.rows[idx + 1].status === 'cancelled') { showInfo('Нижняя строка отменена'); return; }
        part.rows[idx + 1].dz = part.rows[idx].dz || '';
        drawPartTable({ keepScroll: true });
      }
      return;
    }
    // Выделение строк
    const tr = e.target.closest('tr[data-row]');
    if (!tr) return;
    const r = parseInt(tr.dataset.row, 10);
    if (e.shiftKey && state.lastClickedRow !== null) {
      const from = Math.min(state.lastClickedRow, r);
      const to = Math.max(state.lastClickedRow, r);
      for (let i = from; i <= to; i++) state.selectedRows.add(i);
    } else if (e.ctrlKey || e.metaKey) {
      state.selectedRows.has(r) ? state.selectedRows.delete(r) : state.selectedRows.add(r);
      state.lastClickedRow = r;
    } else {
      state.selectedRows = new Set([r]);
      state.lastClickedRow = r;
    }
    highlightRows();
  };

  tbody.ondblclick = (e) => {
    const td = e.target.closest('.td-tema, .td-dz');
    if (!td) return;
    const r = parseInt(td.dataset.r, 10);
    const f = td.dataset.f;
    if (part.rows[r].status === 'cancelled') return;
    td.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    td.contentEditable = 'true';
    td.focus();
    const save = () => {
      td.contentEditable = 'false';
      part.rows[r][f] = td.innerText.trim();
      if (f === 'tema' && !part.rows[r][f]) {
        preserveForRestore(part.rows[r]);
        part.rows[r].status = 'cancelled';
        part.rows[r].dz = '';
      }
      drawPartTable({ keepScroll: true });
    };
    td.onblur = save;
    td.onkeydown = ev => { if (ev.key === 'Enter') { ev.preventDefault(); td.blur(); } };
  };

  let src = null;
  let dragOverTr = null;
  const clearDragOver = () => { if (dragOverTr) { dragOverTr.classList.remove('drag-over-row', 'drag-insert-top', 'drag-insert-bottom'); dragOverTr = null; } };

  tbody.ondragstart = (e) => {
    const tr = e.target.closest('tr[data-row]');
    if (!tr) return;
    src = parseInt(tr.dataset.row, 10);
    tr.style.opacity = '0.4';
  };
  tbody.ondragover = (e) => {
    e.preventDefault();
    // автоскролл
    const wrap = document.querySelector('.editor-wrap');
    if (wrap) {
      const rect = wrap.getBoundingClientRect();
      if (e.clientY < rect.top + 56) wrap.scrollTop -= 18;
      else if (e.clientY > rect.bottom - 56) wrap.scrollTop += 18;
    }
    const tr = e.target.closest('tr[data-row]');
    if (!tr) return;
    if (dragOverTr !== tr) {
      clearDragOver();
      dragOverTr = tr;
      const rect = tr.getBoundingClientRect();
      const pos = e.clientY < rect.top + rect.height / 2 ? 'top' : 'bottom';
      tr.classList.add('drag-over-row', pos === 'top' ? 'drag-insert-top' : 'drag-insert-bottom');
    }
  };
  tbody.ondragleave = (e) => { if (!e.relatedTarget || !tbody.contains(e.relatedTarget)) clearDragOver(); };
  tbody.ondragend = (e) => {
    const tr = e.target.closest('tr[data-row]');
    if (tr) tr.style.opacity = '1';
    clearDragOver(); src = null;
  };
  tbody.ondrop = (e) => {
    e.preventDefault();
    clearDragOver();
    const tr = e.target.closest('tr[data-row]');
    if (!tr) return;
    const dst = parseInt(tr.dataset.row, 10);
    if (src === null || dst === src) return;
    const srcRow = part.rows[src], dstRow = part.rows[dst];
    if (dstRow.status === 'cancelled') {
      dstRow.tema = srcRow.tema; dstRow.dz = srcRow.dz; dstRow.status = 'normal';
    } else {
      dstRow.tema = [dstRow.tema, srcRow.tema].filter(Boolean).join(state.separator);
      dstRow.dz = [dstRow.dz, srcRow.dz].filter(Boolean).join(state.separator);
    }
    preserveForRestore(srcRow);
    srcRow.tema = ''; srcRow.dz = ''; srcRow.status = 'cancelled';
    drawPartTable({ keepScroll: true });
  };
}

function preserveForRestore(row) {
  if (!row._backupTema && row.tema) row._backupTema = row.tema;
  if (!row._backupDz && row.dz) row._backupDz = row.dz;
}

function highlightRows() {
  document.querySelectorAll('#editor-tbody tr[data-row]').forEach(tr => {
    tr.classList.toggle('selected-row', state.selectedRows.has(parseInt(tr.dataset.row, 10)));
  });
}

function cancelSelected() {
  if (!state.selectedRows.size) return alert('Выделите строки для отмены');
  const rows = state.parts[state.activeTab].rows;
  state.selectedRows.forEach(i => {
    preserveForRestore(rows[i]);
    rows[i].status = 'cancelled';
    rows[i].tema = ''; rows[i].dz = '';
  });
  drawPartTable();
}

function restoreSelected() {
  if (!state.selectedRows.size) return alert('Выделите строки для возврата');
  const rows = state.parts[state.activeTab].rows;
  state.selectedRows.forEach(i => {
    rows[i].status = 'normal';
    if (!rows[i].tema && rows[i]._backupTema) rows[i].tema = rows[i]._backupTema;
    if (!rows[i].dz && rows[i]._backupDz) rows[i].dz = rows[i]._backupDz;
  });
  drawPartTable();
}

function setTemplate(t) {
  state.currentTemplate = t;
  document.querySelectorAll('.tpl-btn').forEach(b => b.classList.toggle('active-tpl', b.dataset.tpl === t));
  updateTemplateState();
  if (state.selectedRows.size > 0) {
    const rows = state.parts[state.activeTab].rows;
    state.selectedRows.forEach(i => { if (rows[i].status !== 'cancelled') rows[i].dz = t; });
    drawPartTable();
  }
}

function applyTemplateToPart() {
  const tpl = state.currentTemplate;
  state.parts[state.activeTab].rows.forEach(r => { if (r.status !== 'cancelled') r.dz = tpl; });
  showInfo(`Шаблон «${tpl}» применен ко всей части`);
  drawPartTable();
}

function applyTemplateToAllParts() {
  const tpl = state.currentTemplate;
  state.parts.forEach(p => p.rows.forEach(r => { if (r.status !== 'cancelled') r.dz = tpl; }));
  showInfo(`Шаблон «${tpl}» применен ко всем урокам`);
  drawPartTable();
}

function updateTemplateState() {
  const el = document.getElementById('tpl-state');
  if (el) el.textContent = `ВЫБРАН: ${state.currentTemplate.toUpperCase()}`;
}

function showInfo(msg) {
  const wa = document.getElementById('warn-area');
  if (!wa) return;
  wa.innerHTML = `<div class="success-msg">${esc(msg)}</div>`;
  setTimeout(() => { if (wa) wa.innerHTML = ''; }, 2500);
}

function doExport() {
  // Проверка перед экспортом
  const issues = [];
  state.parts.forEach((part, pi) => {
    const lbl = (state.periodMode === 'quarters' ? 'Ч' : 'П') + (pi + 1);
    const emptyTema = part.rows.filter(r => r.status !== 'cancelled' && !r.tema.trim()).length;
    const emptyDz = part.rows.filter(r => r.status !== 'cancelled' && !r.dz.trim()).length;
    if (emptyTema) issues.push(`${lbl}: ${emptyTema} уроков с пустой темой`);
    if (emptyDz) issues.push(`${lbl}: ${emptyDz} уроков с пустым ДЗ`);
  });
  if (issues.length && !window.confirm(`Найдены незаполненные данные:\n\n${issues.join('\n')}\n\nЭкспортировать всё равно?`)) return;

  // Экран выбора названия класса
  const p = state.periodMode === 'quarters' ? 'Ч' : 'П';
  const partLabels = state.parts.map((_, i) => `${p}${i + 1}`).join(', ');

  document.getElementById('editor-content').innerHTML = `
    <div class="export-name-screen">
      <div class="big-icon">📝</div>
      <h3>Введите название класса</h3>
      <p>Файлы будут названы: <span id="name-preview-label" class="name-preview-example">10А ${p}1.xlsx, 10А ${p}2.xlsx...</span></p>
      <div class="export-name-row">
        <input id="class-name-input" class="class-name-input" type="text" placeholder="Например: 10А" maxlength="20" autofocus>
        <button class="btn btn-success btn-lg" id="btn-generate">Сформировать файлы →</button>
      </div>
      <div id="export-links-area"></div>
    </div>`;

  const input = document.getElementById('class-name-input');
  const preview = document.getElementById('name-preview-label');

  input.addEventListener('input', () => {
    const cls = input.value.trim();
    const prefix = cls ? `${cls} ` : '';
    preview.textContent = state.parts.map((_, i) => `${prefix}${p}${i+1}.xlsx`).join(', ');
  });

  input.addEventListener('keydown', e => { if (e.key === 'Enter') generateLinks(); });
  document.getElementById('btn-generate').onclick = generateLinks;

  function generateLinks() {
    const cls = input.value.trim();
    const prefix = cls ? `${cls} ` : '';
    const links = [];

    for (let i = 0; i < state.parts.length; i++) {
      const data = [['№ урока', 'Тема урока', 'Домашнее задание']];
      state.parts[i].rows.forEach((r, idx) => data.push([idx + 1, r.tema || '', r.dz || '']));
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(data);
      ws['!cols'] = [{wch: 12}, {wch: 55}, {wch: 40}];
      XLSX.utils.book_append_sheet(wb, ws, 'Уроки');
      const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      links.push({ name: `${prefix}${p}${i+1}.xlsx`, url });
    }

    const htmlLinks = links.map(f => `
      <a href="${f.url}" download="${f.name}" class="btn btn-primary btn-lg" style="text-decoration:none;width:100%">
        📥 Скачать ${esc(f.name)}
      </a>`).join('');

    document.getElementById('export-links-area').innerHTML = `
      <div class="export-links-ready">
        <p style="font-weight:700;color:var(--success);margin-bottom:12px">✅ Файлы готовы — нажмите для скачивания:</p>
        <div style="display:flex;flex-direction:column;gap:10px;width:100%;max-width:440px">${htmlLinks}</div>
      </div>`;
  }
}

async function parseDocx(file) {
  const buffer = await file.arrayBuffer();
  const res = await mammoth.convertToHtml({ arrayBuffer: buffer });
  return htmlToTables(res.value);
}

async function parseOdt(file) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const xml = await zip.file('content.xml')?.async('string');
  if (!xml) return [];
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const NS = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';
  const NS_T = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
  return Array.from(doc.getElementsByTagNameNS(NS, 'table')).map(t => {
    const rows = Array.from(t.getElementsByTagNameNS(NS, 'table-row')).map(tr => {
      const cells = [];
      Array.from(tr.getElementsByTagNameNS(NS, 'table-cell')).forEach(c => {
        const rep = parseInt(c.getAttributeNS(NS, 'number-columns-repeated') || '1', 10);
        const txt = Array.from(c.getElementsByTagNameNS(NS_T, 'p')).map(p => p.textContent || '').join('\n').trim();
        for (let i = 0; i < Math.min(rep, 20); i++) cells.push(txt);
      });
      return cells;
    }).filter(r => r.some(c => c));
    return { rows };
  }).filter(t => t.rows.length);
}

function htmlToTables(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return Array.from(doc.querySelectorAll('table')).map(table => ({
    rows: Array.from(table.querySelectorAll('tr')).map(tr =>
      Array.from(tr.querySelectorAll('th,td')).map(td => (td.textContent || '').trim())
    ).filter(r => r.length)
  })).filter(t => t.rows.length);
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = String(s || '');
  return d.innerHTML;
}

window.addEventListener('DOMContentLoaded', initApp);
