import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import * as mammoth from 'mammoth/mammoth.browser';

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
  defaultDzForAll: ''
};

function initApp() {
  document.getElementById('root').innerHTML = `
    <header>
      <div class="logo">📚</div>
      <div><h1>Конвертер КТП -> Excel</h1><span>Vercel A2 (без сервера)</span></div>
    </header>
    <div class="wizard"><div class="stepper" id="stepper"></div><div id="steps-container"></div></div>`;
  renderStepper();
  renderStep(1);
}

function renderStepper() {
  const el = document.getElementById('stepper');
  const steps = [{n:1,l:'Файл'},{n:2,l:'Столбцы'},{n:3,l:'Части'},{n:4,l:'Редактор'}];
  let h = '';
  steps.forEach((s, i) => {
    const active = state.currentStep === s.n;
    const done = state.currentStep > s.n;
    h += `<div class="step${active ? ' active' : ''}${done ? ' done' : ''}"><div class="step-num">${s.n}</div><div class="step-label">${s.l}</div></div>`;
    if (i < steps.length - 1) h += `<div class="step-line${done ? ' done' : ''}"></div>`;
  });
  el.innerHTML = h;
}

function renderStep(n) {
  state.currentStep = n;
  renderStepper();
  const c = document.getElementById('steps-container');
  c.innerHTML = '';
  [null, renderStep1, renderStep2, renderStep3, renderStep4][n](c);
}

function renderStep1(c) {
  c.innerHTML = `<div class="card">
    <h2>Шаг 1. Загрузите файл КТП</h2>
    <p class="sub">Файл обрабатывается на вашем компьютере, никуда не отправляется.</p>
    <div class="info-box">
      <h4>Какой файл нужен?</h4>
      <p><b>.docx</b> или <b>.odt</b> с таблицей, где есть колонка темы урока. Колонка ДЗ может отсутствовать.</p>
    </div>
    <div class="dropzone" id="dropzone"><div class="dropzone-icon">📄</div><h3>Перетащите файл сюда</h3><p>или нажмите для выбора (.docx, .odt)</p></div>
    <input type="file" id="file-input" accept=".docx,.odt">
    <div id="file-result" class="hidden"><div class="file-badge" id="file-badge"></div><div style="margin-top:20px"><div class="table-list" id="table-list"></div></div></div>
    <div id="step1-error"></div>
    <div class="actions"><div></div><button class="btn btn-primary btn-lg" id="btn-next1" disabled>Далее →</button></div>
  </div>`;

  const dz = document.getElementById('dropzone');
  const fi = document.getElementById('file-input');
  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', e => {
    e.preventDefault();
    dz.classList.remove('drag-over');
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  });
  fi.addEventListener('change', () => { if (fi.files[0]) handleFile(fi.files[0]); });
  document.getElementById('btn-next1').addEventListener('click', () => {
    if (state.selectedTableIdx !== null) renderStep(2);
  });
}

async function handleFile(file) {
  const dz = document.getElementById('dropzone');
  dz.innerHTML = '<div style="padding:24px;display:flex;align-items:center;justify-content:center;gap:10px"><div class="loading"></div><span>Читаю файл...</span></div>';
  try {
    const ext = file.name.split('.').pop().toLowerCase();
    let tables = [];
    if (ext === 'docx') tables = await parseDocx(file);
    else if (ext === 'odt') tables = await parseOdt(file);
    else throw new Error('Поддерживаются только .docx и .odt');

    if (!tables.length) throw new Error('Таблицы не найдены в файле');

    state.fileName = file.name;
    state.rawTables = tables;
    state.selectedTableIdx = tables.length === 1 ? 0 : null;
    autoDetectColumns();

    dz.innerHTML = '<div style="padding:24px;color:var(--success);font-weight:600">✅ Файл загружен</div>';
    document.getElementById('file-result').classList.remove('hidden');
    document.getElementById('file-badge').textContent = `✅ ${file.name}`;
    renderTableList();
  } catch (e) {
    dz.innerHTML = `<div class="error-msg">${esc(e.message)}</div>`;
  }
}

function renderTableList() {
  const list = document.getElementById('table-list');
  list.innerHTML = '';
  state.rawTables.forEach((t, i) => {
    const d = document.createElement('div');
    d.className = `table-item${state.selectedTableIdx === i ? ' selected' : ''}`;
    const cols = Math.max(...t.rows.map(r => r.length));
    d.innerHTML = `<b>Таблица ${i + 1}</b> — ${t.rows.length} строк, ${cols} столбцов`;
    d.onclick = () => {
      state.selectedTableIdx = i;
      autoDetectColumns();
      renderTableList();
      document.getElementById('btn-next1').disabled = false;
    };
    list.appendChild(d);
  });
  document.getElementById('btn-next1').disabled = state.selectedTableIdx === null;
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

function renderStep2(c) {
  c.innerHTML = `<div class="card">
    <h2>Шаг 2. Выберите столбцы</h2>
    <div id="col-status" class="col-status waiting">Выберите столбцы</div>
    <div class="preview-wrap"><table class="preview-table" id="preview-table"></table></div>
    <div class="option-row"><input type="checkbox" id="no-dz" ${state.noDz ? 'checked' : ''}><label for="no-dz"><b>В этом КТП нет столбца ДЗ</b></label></div>
    <div class="option-row"><input type="checkbox" id="has-header" ${state.hasHeader ? 'checked' : ''}><label for="has-header">Первая строка — заголовки</label></div>
    <div class="option-row"><label for="separator">Разделитель:</label><select id="separator" class="styled"><option value="\n">Перенос строки</option><option value="; ">Точка с запятой</option></select></div>
    <div class="actions"><button class="btn btn-secondary btn-lg" id="btn-back2">← Назад</button><button class="btn btn-primary btn-lg" id="btn-next2" disabled>Далее →</button></div>
  </div>`;

  document.getElementById('btn-back2').onclick = () => renderStep(1);
  document.getElementById('no-dz').onchange = (e) => { state.noDz = e.target.checked; if (state.noDz) state.dzCol = null; drawPreview(); updateColStatus(); };
  document.getElementById('has-header').onchange = (e) => { state.hasHeader = e.target.checked; };
  document.getElementById('separator').onchange = (e) => { state.separator = e.target.value; };
  document.getElementById('btn-next2').onclick = saveColumnsNext;

  loadPreview();
}

function loadPreview() {
  const table = state.rawTables[state.selectedTableIdx];
  state.previewHeaders = table.rows[0] || [];
  state.previewRows = table.rows.slice(0, 15);
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
    if (ri === 0) return;
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

window.pickCol = function pickCol(i) {
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
    el.innerHTML = 'Нажмите на столбец <b>Тема урока</b>';
    btn.disabled = true;
  } else if (!state.noDz && state.dzCol === null) {
    el.className = 'col-status waiting';
    el.innerHTML = 'Тема выбрана, теперь выберите столбец <b>ДЗ</b>';
    btn.disabled = true;
  } else {
    el.className = 'col-status ready';
    el.innerHTML = state.noDz ? 'Тема выбрана, ДЗ будет пустым' : 'Тема и ДЗ выбраны';
    btn.disabled = false;
  }
}

function saveColumnsNext() {
  const table = state.rawTables[state.selectedTableIdx];
  state.totalDataRows = Math.max(0, table.rows.length - (state.hasHeader ? 1 : 0));
  renderStep(3);
}

function renderStep3(c) {
  c.innerHTML = `<div class="card">
    <h2>Шаг 3. Разделите на части</h2>
    <p class="sub">Всего строк уроков: <b>${state.totalDataRows}</b></p>
    <div class="radio-group"><div class="radio-btn ${state.periodMode === 'quarters' ? 'active' : ''}" id="rb-q">4 четверти</div><div class="radio-btn ${state.periodMode === 'halves' ? 'active' : ''}" id="rb-h">2 полугодия</div></div>
    <div class="parts-grid" id="parts-grid"></div>
    <div class="option-row"><span style="font-weight:600">Заполнить ДЗ во всех уроках:</span><button class="btn btn-secondary btn-sm" id="dz-k">Конспект</button><button class="btn btn-secondary btn-sm" id="dz-c">Карточка</button><button class="btn btn-secondary btn-sm" id="dz-cl">Очистить</button><span id="dz-state"></span></div>
    <div id="count-info" class="count-info bad"></div>
    <div class="actions"><button class="btn btn-secondary btn-lg" id="btn-back3">← Назад</button><button class="btn btn-primary btn-lg" id="btn-next3" disabled>Далее →</button></div>
  </div>`;

  document.getElementById('btn-back3').onclick = () => renderStep(2);
  document.getElementById('rb-q').onclick = () => setMode('quarters');
  document.getElementById('rb-h').onclick = () => setMode('halves');
  document.getElementById('dz-k').onclick = () => setDzAll('Конспект');
  document.getElementById('dz-c').onclick = () => setDzAll('Карточка');
  document.getElementById('dz-cl').onclick = () => setDzAll('');
  document.getElementById('btn-next3').onclick = buildPartsNext;

  if (!state.partCounts.length) state.partCounts = state.periodMode === 'quarters' ? [0,0,0,0] : [0,0];
  buildPartsGrid();
  refreshDzState();
}

function setMode(m) {
  state.periodMode = m;
  state.partCounts = m === 'quarters' ? [0,0,0,0] : [0,0];
  renderStep(3);
}

function setDzAll(v) {
  state.defaultDzForAll = v;
  refreshDzState();
}

function refreshDzState() {
  const el = document.getElementById('dz-state');
  if (!el) return;
  el.textContent = state.defaultDzForAll ? `Сейчас: ${state.defaultDzForAll}` : 'Сейчас: без автозаполнения';
}

function buildPartsGrid() {
  const n = state.periodMode === 'quarters' ? 4 : 2;
  const p = state.periodMode === 'quarters' ? 'Ч' : 'П';
  const grid = document.getElementById('parts-grid');
  grid.innerHTML = '';
  for (let i = 0; i < n; i++) {
    const d = document.createElement('div');
    d.className = 'part-input';
    d.innerHTML = `<label>${p}${i + 1} — уроков:</label><input id="pc-${i}" type="number" min="0" value="${state.partCounts[i] || ''}">`;
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
  const ok = sum === state.totalDataRows && sum > 0;
  const info = document.getElementById('count-info');
  info.className = `count-info ${ok ? 'ok' : 'bad'}`;
  info.innerHTML = `<span>В КТП: <b>${state.totalDataRows}</b></span><span>Сумма: <b>${sum}</b></span>${ok ? '<span>✅ Совпадает</span>' : `<span>❌ Разница: ${Math.abs(sum - state.totalDataRows)}</span>`}`;
  document.getElementById('btn-next3').disabled = !ok;
}

function buildPartsNext() {
  const n = state.periodMode === 'quarters' ? 4 : 2;
  const counts = state.partCounts.slice(0, n);
  const table = state.rawTables[state.selectedTableIdx];
  const start = state.hasHeader ? 1 : 0;
  const rows = table.rows.slice(start).map(r => ({
    tema: String(r[state.temaCol] || '').trim(),
    dz: state.noDz ? '' : String(r[state.dzCol] || '').trim(),
    status: 'normal'
  }));
  if (state.defaultDzForAll) rows.forEach(r => { r.dz = state.defaultDzForAll; });

  state.parts = [];
  let offset = 0;
  counts.forEach(cnt => {
    state.parts.push({ rows: rows.slice(offset, offset + cnt) });
    offset += cnt;
  });
  renderStep(4);
}

function renderStep4(c) {
  const n = state.parts.length;
  const p = state.periodMode === 'quarters' ? 'Ч' : 'П';
  let tabs = '';
  for (let i = 0; i < n; i++) tabs += `<button class="tab-btn${i===0?' active':''}" data-ti="${i}">${p}${i + 1}</button>`;

  c.innerHTML = `<div class="card">
    <h2>Шаг 4. Редактор частей</h2>
    <div class="tabs" id="tabs">${tabs}</div>
    <div class="editor-toolbar"><button class="btn btn-danger btn-sm" id="btn-cancel">✖ Отменить урок</button><button class="btn btn-secondary btn-sm" id="btn-restore">↩ Вернуть урок</button><div class="sep"></div><span style="font-size:.85rem;color:var(--gray-500)">Шаблоны ДЗ:</span><button class="tpl-btn" data-tpl="Конспект">Конспект</button><button class="tpl-btn" data-tpl="Карточка">Карточка</button></div>
    <div id="editor-content"></div>
    <div class="actions"><button class="btn btn-secondary btn-lg" id="btn-back4">← Назад</button><button class="btn btn-success btn-lg" id="btn-export">💾 Экспорт</button></div>
  </div>`;

  state.activeTab = 0;
  state.selectedRows = new Set();
  document.getElementById('tabs').addEventListener('click', e => {
    const b = e.target.closest('.tab-btn');
    if (!b) return;
    state.activeTab = parseInt(b.dataset.ti, 10);
    document.querySelectorAll('.tab-btn').forEach((x, i) => x.classList.toggle('active', i === state.activeTab));
    drawPartTable();
  });
  document.getElementById('btn-back4').onclick = () => renderStep(3);
  document.getElementById('btn-cancel').onclick = cancelSelected;
  document.getElementById('btn-restore').onclick = restoreSelected;
  document.getElementById('btn-export').onclick = doExport;
  document.querySelectorAll('.tpl-btn').forEach(b => b.onclick = () => setTemplate(b.dataset.tpl));

  drawPartTable();
}

function drawPartTable() {
  const part = state.parts[state.activeTab];
  const ec = document.getElementById('editor-content');
  let h = `<div class="editor-wrap"><table class="editor-table"><thead><tr><th style="width:50px">№</th><th>Тема урока</th><th>Домашнее задание</th><th style="width:120px">Статус</th></tr></thead><tbody id="editor-tbody">`;
  part.rows.forEach((r, i) => {
    const cancelled = r.status === 'cancelled';
    h += `<tr data-row="${i}" class="${cancelled?'cancelled':''}${state.selectedRows.has(i)?' selected-row':''}" draggable="${!cancelled}"><td class="td-num">${i+1}</td><td class="td-tema" data-f="tema" data-r="${i}">${cancelled?'':esc(r.tema)}</td><td class="td-dz" data-f="dz" data-r="${i}">${cancelled?'':esc(r.dz)}</td><td style="text-align:center">${cancelled?'<span class="cancelled-badge">ОТМЕНЁН</span>':''}</td></tr>`;
  });
  h += '</tbody></table></div>';
  ec.innerHTML = h;
  setupEditorEvents(part);
}

function setupEditorEvents(part) {
  const tbody = document.getElementById('editor-tbody');
  tbody.onclick = (e) => {
    if (e.target.closest('[contenteditable]')) return;
    const tr = e.target.closest('tr[data-row]');
    if (!tr) return;
    const r = parseInt(tr.dataset.row, 10);
    if (e.ctrlKey || e.metaKey) state.selectedRows.has(r) ? state.selectedRows.delete(r) : state.selectedRows.add(r);
    else state.selectedRows = new Set([r]);
    highlightRows();
  };

  tbody.ondblclick = (e) => {
    const td = e.target.closest('.td-tema, .td-dz');
    if (!td) return;
    const r = parseInt(td.dataset.r, 10);
    const f = td.dataset.f;
    if (part.rows[r].status === 'cancelled') return;
    td.contentEditable = 'true';
    td.focus();
    const save = () => {
      td.contentEditable = 'false';
      const val = td.innerText.trim();
      part.rows[r][f] = val;
      if (f === 'tema' && !val) {
        preserveForRestore(part.rows[r]);
        part.rows[r].status = 'cancelled';
        part.rows[r].dz = '';
      }
      drawPartTable();
      td.removeEventListener('blur', save);
    };
    td.addEventListener('blur', save);
  };

  let src = null;
  tbody.addEventListener('dragstart', (e) => {
    const tr = e.target.closest('tr[data-row]');
    if (!tr) return;
    src = parseInt(tr.dataset.row, 10);
    e.dataTransfer.setData('text/plain', String(src));
  });
  tbody.addEventListener('dragover', (e) => e.preventDefault());
  tbody.addEventListener('drop', (e) => {
    e.preventDefault();
    const tr = e.target.closest('tr[data-row]');
    if (!tr) return;
    const dst = parseInt(tr.dataset.row, 10);
    if (src === null || dst === src) return;
    const srcRow = part.rows[src], dstRow = part.rows[dst];
    dstRow.tema = [dstRow.tema, srcRow.tema].filter(Boolean).join(state.separator);
    dstRow.dz = [dstRow.dz, srcRow.dz].filter(Boolean).join(state.separator);
    preserveForRestore(srcRow);
    srcRow.tema = '';
    srcRow.dz = '';
    srcRow.status = 'cancelled';
    drawPartTable();
  });
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
  if (!state.selectedRows.size) return alert('Выделите строки');
  const rows = state.parts[state.activeTab].rows;
  state.selectedRows.forEach(i => {
    preserveForRestore(rows[i]);
    rows[i].status = 'cancelled';
    rows[i].tema = '';
    rows[i].dz = '';
  });
  drawPartTable();
}

function restoreSelected() {
  if (!state.selectedRows.size) return alert('Выделите строки');
  const rows = state.parts[state.activeTab].rows;
  state.selectedRows.forEach(i => {
    rows[i].status = 'normal';
    if (!rows[i].tema && rows[i]._backupTema) rows[i].tema = rows[i]._backupTema;
    if (!rows[i].dz && rows[i]._backupDz) rows[i].dz = rows[i]._backupDz;
  });
  drawPartTable();
}

function setTemplate(t) {
  if (!state.selectedRows.size) return alert('Выделите строки');
  const rows = state.parts[state.activeTab].rows;
  state.selectedRows.forEach(i => { if (rows[i].status !== 'cancelled') rows[i].dz = t; });
  drawPartTable();
}

function doExport() {
  const p = state.periodMode === 'quarters' ? 'Ч' : 'П';
  for (let i = 0; i < state.parts.length; i++) {
    const data = [['Тема урока', 'Домашнее задание']];
    state.parts[i].rows.forEach(r => data.push([r.tema || '', r.dz || '']));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = [{wch: 60}, {wch: 40}];
    XLSX.utils.book_append_sheet(wb, ws, 'Уроки');
    XLSX.writeFile(wb, `${p}${i + 1}.xlsx`, { bookType: 'xlsx' });
    XLSX.writeFile(wb, `${p}${i + 1}.xls`, { bookType: 'biff8' });
  }
  alert('Готово! Файлы сохранены.');
}

async function parseDocx(file) {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });
  return htmlTablesToRows(result.value);
}

async function parseOdt(file) {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);
  const contentXml = await zip.file('content.xml')?.async('string');
  if (!contentXml) return [];
  const doc = new DOMParser().parseFromString(contentXml, 'application/xml');
  const NS_TABLE = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0';
  const NS_TEXT = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0';
  const tables = [];
  const tableEls = Array.from(doc.getElementsByTagNameNS(NS_TABLE, 'table'));
  tableEls.forEach(t => {
    const rows = [];
    const rowEls = Array.from(t.getElementsByTagNameNS(NS_TABLE, 'table-row'));
    rowEls.forEach(re => {
      const cells = [];
      const cellEls = Array.from(re.getElementsByTagNameNS(NS_TABLE, 'table-cell'));
      cellEls.forEach(c => {
        const rep = parseInt(c.getAttributeNS(NS_TABLE, 'number-columns-repeated') || '1', 10);
        const pEls = Array.from(c.getElementsByTagNameNS(NS_TEXT, 'p'));
        const txt = pEls.map(x => x.textContent || '').join('\n').trim();
        for (let i = 0; i < Math.min(rep, 20); i++) cells.push(txt);
      });
      if (cells.some(x => x.trim() !== '')) rows.push(cells);
    });
    if (rows.length) tables.push({ rows });
  });
  return tables;
}

function htmlTablesToRows(html) {
  const doc = new DOMParser().parseFromString(html, 'text/html');
  return Array.from(doc.querySelectorAll('table')).map(table => {
    const rows = Array.from(table.querySelectorAll('tr')).map(tr =>
      Array.from(tr.querySelectorAll('th,td')).map(td => (td.textContent || '').trim())
    ).filter(r => r.length);
    return { rows };
  }).filter(t => t.rows.length);
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = String(s || '');
  return d.innerHTML;
}

window.addEventListener('DOMContentLoaded', initApp);
