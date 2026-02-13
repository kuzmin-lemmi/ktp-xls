/**
 * Конвертер КТП -> Excel | Фронтенд
 */
const API = '/api';

const state = {
  sessionId: null, fileName: null, tables: [], selectedTableIdx: null,
  temaCol: null, dzCol: null, noDz: false, hasHeader: true, separator: '\n',
  defaultDzForAll: '',
  periodMode: 'quarters', partCounts: [], parts: [], totalDataRows: 0,
  activeTab: 0, selectedRows: new Set(), currentStep: 1,
  previewHeaders: [], previewRows: [],
  dragSrcRow: null // для drag&drop тем
};

// ============================================================
// INIT
// ============================================================
function initApp() {
  document.getElementById('root').innerHTML = `
    <header>
      <div class="logo">&#128218;</div>
      <div>
        <h1>Конвертер КТП &rarr; Excel</h1>
        <span>Импорт в Дневник.ру</span>
      </div>
    </header>
    <div class="wizard">
      <div class="stepper" id="stepper"></div>
      <div id="steps-container"></div>
    </div>`;
  renderStepper();
  renderStep(1);
}

// ============================================================
// STEPPER
// ============================================================
function renderStepper() {
  const el = document.getElementById('stepper');
  const steps = [{n:1,l:'Файл'},{n:2,l:'Столбцы'},{n:3,l:'Части'},{n:4,l:'Редактор'}];
  let h = '';
  steps.forEach((s,i) => {
    const act = state.currentStep === s.n, done = state.currentStep > s.n;
    h += `<div class="step${act?' active':''}${done?' done':''}"><div class="step-num">${s.n}</div><div class="step-label">${s.l}</div></div>`;
    if (i < steps.length-1) h += `<div class="step-line${done?' done':''}"></div>`;
  });
  el.innerHTML = h;
}
function renderStep(n) {
  state.currentStep = n; renderStepper();
  const c = document.getElementById('steps-container'); c.innerHTML = '';
  [null,renderStep1,renderStep2,renderStep3,renderStep4][n](c);
}

// ============================================================
// STEP 1 — FILE UPLOAD
// ============================================================
function renderStep1(c) {
  c.innerHTML = `<div class="card">
    <h2>Шаг 1. Загрузите файл КТП</h2>
    <p class="sub">Загрузите ваш файл с календарно-тематическим планированием</p>

    <div class="info-box">
      <h4>&#128196; Какой файл нужен?</h4>
      <p>Нужен ваш файл <b>КТП (календарно-тематическое планирование)</b> в формате <b>.docx</b> (Word) или <b>.odt</b> (LibreOffice).</p>
      <p style="margin-top:8px">Файл должен содержать <b>таблицу</b> с колонками &laquo;Тема урока&raquo; и &laquo;Домашнее задание&raquo;.</p>
      <details style="margin-top:6px">
        <summary style="cursor:pointer;font-weight:600">Показать пример таблицы</summary>
        <table class="example-table">
          <thead><tr><th>&#8470;</th><th>Тема урока</th><th>Кол-во часов</th><th>Домашнее задание</th></tr></thead>
          <tbody>
            <tr><td>1</td><td>Введение в предмет</td><td>1</td><td>&#167; 1, стр. 5</td></tr>
            <tr><td>2</td><td>Основные понятия</td><td>1</td><td>&#167; 2, упр. 3</td></tr>
          </tbody>
        </table>
      </details>
      <p style="margin-top:8px"><b>Не подходят:</b> PDF-файлы, сканы, фотографии документов.</p>
    </div>

    <div class="dropzone" id="dropzone">
      <div class="dropzone-icon">&#128451;</div>
      <h3>Перетащите файл сюда</h3>
      <p>или нажмите, чтобы выбрать файл (.docx или .odt)</p>
    </div>
    <input type="file" id="file-input" accept=".docx,.odt">

    <div id="file-result" class="hidden">
      <div class="file-badge" id="file-badge"></div>
      <div style="margin-top:20px" id="table-chooser">
        <div style="font-size:1rem;font-weight:600;color:var(--gray-700);margin-bottom:10px">Найденные таблицы:</div>
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
  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag-over'); if(e.dataTransfer.files[0]) uploadFile(e.dataTransfer.files[0]); });
  fi.addEventListener('change', () => { if(fi.files[0]) uploadFile(fi.files[0]); });
  document.getElementById('btn-next1').addEventListener('click', () => { if(state.sessionId && state.selectedTableIdx!==null) renderStep(2); });
}

async function uploadFile(file) {
  const dz = document.getElementById('dropzone');
  dz.innerHTML = '<div style="padding:32px;display:flex;align-items:center;justify-content:center;gap:12px"><div class="loading"></div><span style="font-size:1.1rem">Читаю файл...</span></div>';
  const fd = new FormData(); fd.append('file', file);
  try {
    const res = await fetch(`${API}/upload`, {method:'POST', body:fd});
    const data = await res.json();
    if(!res.ok) throw new Error(data.error);
    state.sessionId=data.sessionId; state.fileName=data.fileName; state.tables=data.tables;
    document.getElementById('file-badge').textContent = '\u2705 ' + data.fileName;
    document.getElementById('file-result').classList.remove('hidden');
    dz.innerHTML = '<div style="padding:32px;color:var(--success);font-size:1.1rem;font-weight:600">\u2705 Файл успешно загружен!</div>';
    if(data.tables.length===1){state.selectedTableIdx=0; await doSelectTable(0);}
    renderTableList();
  } catch(err) { dz.innerHTML = `<div class="error-msg">${esc(err.message)}</div>`; }
}

function renderTableList() {
  const list = document.getElementById('table-list'); list.innerHTML='';
  state.tables.forEach((t,i) => {
    const d = document.createElement('div');
    d.className = 'table-item'+(state.selectedTableIdx===i?' selected':'');
    d.innerHTML = `<b>Таблица ${i+1}</b> &mdash; ${t.rowCount} строк, ${t.colCount} столбцов`;
    d.onclick = () => doSelectTable(i);
    list.appendChild(d);
  });
}

async function doSelectTable(idx) {
  try {
    const res = await fetch(`${API}/session/${state.sessionId}/select-table`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({tableIdx:idx})});
    const data = await res.json(); if(!res.ok) throw new Error(data.error);
    state.selectedTableIdx=idx; state.temaCol=data.temaCol; state.dzCol=data.dzCol; state.hasHeader=data.hasHeader;
    renderTableList();
    document.getElementById('btn-next1').disabled = false;
  } catch(err) { document.getElementById('step1-error').innerHTML = `<div class="error-msg">${esc(err.message)}</div>`; }
}

// ============================================================
// STEP 2 — COLUMN SELECTION
// ============================================================
function renderStep2(c) {
  c.innerHTML = `<div class="card">
    <h2>Шаг 2. Выберите столбцы</h2>
    <p class="sub">Нажмите на заголовок столбца, чтобы отметить его.<br>Первый клик &mdash; &laquo;Тема урока&raquo;, второй &mdash; &laquo;Домашнее задание&raquo;.</p>
    <div class="flex-gap" style="margin-bottom:12px">
      <span style="background:#bfdbfe;padding:6px 14px;border-radius:8px;font-weight:600;font-size:0.9rem">&#128309; Синий = Тема</span>
      <span style="background:#bbf7d0;padding:6px 14px;border-radius:8px;font-weight:600;font-size:0.9rem">&#128994; Зелёный = ДЗ</span>
    </div>
    <div id="col-status" class="col-status waiting">Загружаю...</div>
    <div class="preview-wrap"><table class="preview-table" id="preview-table"></table></div>
    <div class="option-row">
      <input type="checkbox" id="no-dz" ${state.noDz?'checked':''}>
      <label for="no-dz"><b>В этом КТП нет столбца ДЗ</b> (создать пустой столбец автоматически)</label>
    </div>
    <div class="option-row">
      <input type="checkbox" id="has-header" ${state.hasHeader?'checked':''}>
      <label for="has-header">Первая строка &mdash; заголовки (не включать в уроки)</label>
    </div>
    <div class="option-row">
      <label for="separator-sel">Разделитель при объединении тем:</label>
      <select id="separator-sel" class="styled">
        <option value="\\n">Перенос строки</option>
        <option value="; ">Точка с запятой</option>
      </select>
    </div>
    <div class="actions">
      <button class="btn btn-secondary btn-lg" id="btn-back2">&larr; Назад</button>
      <button class="btn btn-primary btn-lg" id="btn-next2" disabled>Далее &rarr;</button>
    </div>
  </div>`;
  document.getElementById('btn-back2').addEventListener('click', () => renderStep(1));
  document.getElementById('no-dz').addEventListener('change', e => {
    state.noDz = e.target.checked;
    if (state.noDz) state.dzCol = null;
    drawPreviewTable();
    updateColStatus();
  });
  document.getElementById('has-header').addEventListener('change', e => { state.hasHeader=e.target.checked; });
  document.getElementById('separator-sel').addEventListener('change', e => { state.separator=e.target.value.replace('\\n','\n'); });
  document.getElementById('btn-next2').addEventListener('click', saveColumnsNext);
  loadPreview();
}

async function loadPreview() {
  try {
    const res = await fetch(`${API}/session/${state.sessionId}/preview`);
    const data = await res.json(); if(!res.ok) throw new Error(data.error);
    state.temaCol=data.temaCol; state.dzCol=data.dzCol;
    state.previewHeaders=data.headers; state.previewRows=data.rows; state.totalDataRows=data.dataRows;
    drawPreviewTable(); updateColStatus();
  } catch(err) { document.getElementById('col-status').innerHTML = `<span style="color:var(--danger)">${esc(err.message)}</span>`; }
}

function drawPreviewTable() {
  const tbl = document.getElementById('preview-table');
  let h = '<thead><tr>';
  state.previewHeaders.forEach((hdr,i) => {
    let cls='',badge='';
    if(i===state.temaCol){cls=' col-tema';badge='<span class="col-badge badge-tema">ТЕМА</span>';}
    if(!state.noDz && i===state.dzCol){cls=' col-dz';badge='<span class="col-badge badge-dz">ДЗ</span>';}
    h += `<th class="${cls}" onclick="clickCol(${i})">${esc(hdr)}${badge}</th>`;
  });
  h += '</tr></thead><tbody>';
  state.previewRows.forEach((row,r) => {
    if(r===0) return;
    h += '<tr>';
    row.forEach((cell,ci) => {
      let cls='';
      if(ci===state.temaCol) cls=' col-tema';
      if(!state.noDz && ci===state.dzCol) cls=' col-dz';
      h += `<td class="${cls}">${esc(cell||'')}</td>`;
    });
    h += '</tr>';
  });
  h += '</tbody>'; tbl.innerHTML=h;
}

window.clickCol = function(ci) {
  if(state.temaCol===null){state.temaCol=ci;}
  else if(state.temaCol===ci){state.temaCol=null;}
  else if(!state.noDz && state.dzCol===null && ci!==state.temaCol){state.dzCol=ci;}
  else if(state.dzCol===ci){state.dzCol=null;}
  else{state.temaCol=ci;if(state.dzCol===ci)state.dzCol=null;}
  drawPreviewTable(); updateColStatus();
};

function updateColStatus() {
  const el = document.getElementById('col-status'), btn = document.getElementById('btn-next2');
  if(state.temaCol===null){
    el.className='col-status waiting'; el.innerHTML='&#9755; Нажмите на столбец с <b>темами уроков</b>'; btn.disabled=true;
  } else if(!state.noDz && state.dzCol===null){
    el.className='col-status waiting'; el.innerHTML='\u2705 Тема выбрана! Теперь нажмите на столбец с <b>домашним заданием</b>'; btn.disabled=true;
  } else {
    el.className='col-status ready'; el.innerHTML=state.noDz
      ? '\u2705 Столбец темы выбран. ДЗ будет пустым. Нажмите &laquo;Далее&raquo;'
      : '\u2705 Оба столбца выбраны! Нажмите &laquo;Далее&raquo;';
    btn.disabled=false;
  }
}

async function saveColumnsNext() {
  try {
    const res = await fetch(`${API}/session/${state.sessionId}/columns`, {method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({temaCol:state.temaCol, dzCol:state.dzCol, noDz:state.noDz, hasHeader:state.hasHeader, separator:state.separator})});
    const data = await res.json(); if(!res.ok) throw new Error(data.error);
    state.totalDataRows=data.totalDataRows; renderStep(3);
  } catch(err){alert('Ошибка: '+err.message);}
}

// ============================================================
// STEP 3 — PARTS
// ============================================================
function renderStep3(c) {
  c.innerHTML = `<div class="card">
    <h2>Шаг 3. Разделите на части</h2>
    <p class="sub">Укажите, сколько уроков в каждой четверти (или полугодии).<br>
    Всего строк уроков в вашем КТП: <b style="font-size:1.2rem">${state.totalDataRows}</b></p>
    <div class="radio-group">
      <div class="radio-btn ${state.periodMode==='quarters'?'active':''}" id="rb-q">4 четверти</div>
      <div class="radio-btn ${state.periodMode==='halves'?'active':''}" id="rb-h">2 полугодия</div>
    </div>
    <div class="parts-grid" id="parts-grid"></div>
    <div class="option-row" style="margin-top:0">
      <span style="font-weight:600">Быстро заполнить ДЗ во всех уроках:</span>
      <button class="btn btn-secondary btn-sm" id="btn-dz-all-k">Конспект</button>
      <button class="btn btn-secondary btn-sm" id="btn-dz-all-c">Карточка</button>
      <button class="btn btn-secondary btn-sm" id="btn-dz-all-clear">Очистить</button>
      <span id="dz-all-state" style="font-size:0.85rem;color:var(--gray-500)"></span>
    </div>
    <div id="count-info" class="count-info bad"></div>
    <div class="actions">
      <button class="btn btn-secondary btn-lg" id="btn-back3">&larr; Назад</button>
      <button class="btn btn-primary btn-lg" id="btn-next3" disabled>Далее &rarr;</button>
    </div>
  </div>`;
  document.getElementById('btn-back3').addEventListener('click', () => renderStep(2));
  document.getElementById('rb-q').addEventListener('click', () => setMode('quarters'));
  document.getElementById('rb-h').addEventListener('click', () => setMode('halves'));
  document.getElementById('btn-dz-all-k').addEventListener('click', () => setDzForAll('Конспект'));
  document.getElementById('btn-dz-all-c').addEventListener('click', () => setDzForAll('Карточка'));
  document.getElementById('btn-dz-all-clear').addEventListener('click', () => setDzForAll(''));
  document.getElementById('btn-next3').addEventListener('click', savePartsNext);
  if(state.partCounts.length===0) state.partCounts = state.periodMode==='quarters'?[0,0,0,0]:[0,0];
  buildPartsGrid();
  refreshDzAllState();
}

function setDzForAll(value) {
  state.defaultDzForAll = value;
  refreshDzAllState();
}

function refreshDzAllState() {
  const el = document.getElementById('dz-all-state');
  if (!el) return;
  if (!state.defaultDzForAll) {
    el.textContent = 'Сейчас: без автозаполнения';
  } else {
    el.textContent = `Сейчас: "${state.defaultDzForAll}" во всех уроках`;
  }
}
function setMode(m) {
  state.periodMode=m; state.partCounts = m==='quarters'?[0,0,0,0]:[0,0];
  document.getElementById('rb-q').classList.toggle('active',m==='quarters');
  document.getElementById('rb-h').classList.toggle('active',m==='halves');
  buildPartsGrid();
}
function buildPartsGrid() {
  const n = state.periodMode==='quarters'?4:2, pf=state.periodMode==='quarters'?'Ч':'П';
  const grid = document.getElementById('parts-grid'); grid.innerHTML='';
  for(let i=0;i<n;i++){
    const d=document.createElement('div'); d.className='part-input';
    d.innerHTML=`<label>${pf}${i+1} &mdash; уроков:</label><input type="number" min="0" value="${state.partCounts[i]||''}" id="pc-${i}">`;
    grid.appendChild(d);
    d.querySelector('input').addEventListener('input', recalcParts);
  }
  recalcParts();
}
window.recalcParts = function() {
  const n = state.periodMode==='quarters'?4:2; let sum=0;
  for(let i=0;i<n;i++){const v=parseInt(document.getElementById(`pc-${i}`)?.value||'0',10); state.partCounts[i]=isNaN(v)?0:v; sum+=state.partCounts[i];}
  const total=state.totalDataRows, ok=sum===total&&sum>0;
  const info = document.getElementById('count-info');
  info.className='count-info '+(ok?'ok':'bad');
  info.innerHTML=`<span>В КТП: <b>${total}</b> строк</span><span>Сумма частей: <b>${sum}</b></span>${ok?'<span>\u2705 Совпадает!</span>':`<span>\u274C Разница: ${Math.abs(sum-total)}</span>`}`;
  document.getElementById('btn-next3').disabled = !ok;
};
async function savePartsNext() {
  try {
    const n = state.periodMode==='quarters'?4:2;
    const res = await fetch(`${API}/session/${state.sessionId}/parts`, {method:'POST', headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        periodMode:state.periodMode,
        partCounts:state.partCounts.slice(0,n),
        defaultDzForAll: state.defaultDzForAll
      })});
      
    const data = await res.json(); if(!res.ok) throw new Error(data.error);
    state.parts=data.parts; renderStep(4);
  } catch(err){alert('Ошибка: '+err.message);}
}

// ============================================================
// STEP 4 — EDITOR with Drag & Drop
// ============================================================
function renderStep4(c) {
  const pf = state.periodMode==='quarters'?'Ч':'П', n=state.parts.length;
  let tabsH=''; for(let i=0;i<n;i++) tabsH+=`<button class="tab-btn${i===0?' active':''}" data-ti="${i}">${pf}${i+1}</button>`;

  c.innerHTML = `<div class="card">
    <h2>Шаг 4. Редактор частей</h2>
    <p class="sub">Здесь вы можете отредактировать темы и домашние задания перед экспортом.</p>

    <div class="info-box" style="margin-bottom:20px">
      <h4>&#128161; Как пользоваться редактором</h4>
      <ul>
        <li><b>Двойной клик</b> по ячейке &mdash; редактирование текста</li>
        <li><b>Перетащите тему</b> мышкой в другую строку &mdash; темы объединятся, а исходный урок отменится</li>
        <li><b>Клик по строке</b> &mdash; выделение (Ctrl+клик для нескольких)</li>
        <li><b>Отменить урок</b> скрывает тему, но её можно вернуть кнопкой <b>&laquo;Вернуть урок&raquo;</b></li>
        <li>Если тема урока пустая &mdash; урок автоматически отменяется</li>
      </ul>
    </div>

    <div class="tabs" id="tabs">${tabsH}</div>
    <div class="editor-toolbar">
      <button class="btn btn-danger btn-sm" id="btn-cancel">&#10006; Отменить урок</button>
      <button class="btn btn-secondary btn-sm" id="btn-restore">&#8617; Вернуть урок</button>
      <div class="sep"></div>
      <span style="font-size:0.85rem;color:var(--gray-500)">Шаблоны ДЗ:</span>
      <button class="tpl-btn" data-tpl="Конспект">Конспект</button>
      <button class="tpl-btn" data-tpl="Карточка">Карточка</button>
      <button class="tpl-btn" data-tpl="§ ">&#167;...</button>
      <button class="tpl-btn" data-tpl="Упр. ">Упр...</button>
    </div>
    <div id="warn-area"></div>
    <div id="editor-content"><div style="padding:32px;display:flex;align-items:center;gap:10px"><div class="loading"></div> Загружаю...</div></div>
    <div class="actions">
      <button class="btn btn-secondary btn-lg" id="btn-back4">&larr; Назад</button>
      <button class="btn btn-success btn-lg" id="btn-export">&#128190; Экспорт в Excel</button>
    </div>
  </div>`;

  state.activeTab=0; state.selectedRows=new Set();
  document.getElementById('tabs').addEventListener('click', e => {
    const btn=e.target.closest('.tab-btn'); if(!btn) return;
    state.activeTab=parseInt(btn.dataset.ti); state.selectedRows=new Set();
    document.querySelectorAll('.tab-btn').forEach((b,i) => b.classList.toggle('active',i===state.activeTab));
    loadPart(state.activeTab);
  });
  document.getElementById('btn-cancel').addEventListener('click', doCancelSelected);
  document.getElementById('btn-restore').addEventListener('click', doRestoreSelected);
  document.getElementById('btn-back4').addEventListener('click', () => renderStep(3));
  document.getElementById('btn-export').addEventListener('click', doExport);
  document.querySelectorAll('.tpl-btn').forEach(btn => btn.addEventListener('click', () => doInsertTemplate(btn.dataset.tpl)));
  loadPart(0);
}

async function loadPart(idx) {
  try {
    const res = await fetch(`${API}/session/${state.sessionId}/part/${idx}`);
    const data = await res.json(); if(!res.ok) throw new Error(data.error);
    state.parts[idx]=data; drawPartTable(data,idx);
  } catch(err) { document.getElementById('editor-content').innerHTML = `<div class="error-msg">${esc(err.message)}</div>`; }
}

function drawPartTable(part, idx) {
  const ec = document.getElementById('editor-content');
  let h = `<div class="editor-wrap"><table class="editor-table">
    <thead><tr>
      <th style="width:50px">&#8470;</th>
      <th style="width:45%">&#128196; Тема урока <span style="font-weight:400;color:var(--gray-400)">(перетаскивайте)</span></th>
      <th style="width:40%">&#128221; Домашнее задание</th>
      <th style="width:70px">Статус</th>
    </tr></thead>
    <tbody id="editor-tbody">
    <tr><td colspan="4" class="drag-hint">&#128161; Перетащите тему урока мышкой на другую строку, чтобы объединить</td></tr>`;

  part.rows.forEach((row,r) => {
    const cancelled = row.status==='cancelled';
    h += `<tr data-row="${r}" class="${cancelled?'cancelled':''}${state.selectedRows.has(r)?' selected-row':''}" draggable="${!cancelled}">
      <td class="td-num">${r+1}</td>
      <td class="td-tema" data-f="tema" data-r="${r}">${cancelled?'':esc(row.tema)}</td>
      <td class="td-dz" data-f="dz" data-r="${r}">${cancelled?'':esc(row.dz)}</td>
      <td style="text-align:center">${cancelled?'<span class="cancelled-badge">ОТМЕНЁН</span>':''}</td>
    </tr>`;
  });

  h += '</tbody></table></div>';
  ec.innerHTML = h;
  setupEditorEvents(part);
}

function setupEditorEvents(part) {
  const tbody = document.getElementById('editor-tbody');

  // === Row selection ===
  tbody.addEventListener('click', e => {
    if(e.target.closest('[contenteditable]')) return;
    const tr = e.target.closest('tr[data-row]'); if(!tr) return;
    const r = parseInt(tr.dataset.row);
    if(e.ctrlKey||e.metaKey) state.selectedRows.has(r)?state.selectedRows.delete(r):state.selectedRows.add(r);
    else state.selectedRows = new Set([r]);
    highlightRows();
  });

  // === Double-click edit ===
  tbody.addEventListener('dblclick', e => {
    const td = e.target.closest('.td-tema, .td-dz'); if(!td) return;
    const r = parseInt(td.dataset.r), f = td.dataset.f, row = part.rows[r];
    if(row.status==='cancelled') return;
    td.contentEditable = 'true'; td.focus();
    const range = document.createRange(); range.selectNodeContents(td); range.collapse(false);
    window.getSelection().removeAllRanges(); window.getSelection().addRange(range);
    function save() {
      td.contentEditable='false';
      const val = td.innerText.trim(); row[f]=val;
      // Автоотмена: если тема пустая, отменяем урок
      if(f==='tema' && !val) {
        preserveRowForRestore(row);
        row.status='cancelled'; row.tema=''; row.dz='';
      }
      savePart(state.activeTab);
      drawPartTable(part, state.activeTab);
      td.removeEventListener('blur',save);
    }
    function onKey(e2) {
      if(e2.key==='Escape'){td.contentEditable='false'; td.textContent=row[f]; td.removeEventListener('blur',save); td.removeEventListener('keydown',onKey);}
      if(e2.key==='Enter'&&!e2.shiftKey){e2.preventDefault(); td.blur();}
    }
    td.addEventListener('blur',save);
    td.addEventListener('keydown',onKey);
  });

  // === DRAG & DROP тем ===
  let dragSrcRow = null;

  tbody.addEventListener('dragstart', e => {
    const tr = e.target.closest('tr[data-row]');
    if(!tr) return;
    const r = parseInt(tr.dataset.row);
    if(part.rows[r].status==='cancelled'){e.preventDefault(); return;}
    dragSrcRow = r;
    tr.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', r.toString());
    // Make drag image
    const tema = part.rows[r].tema;
    const ghost = document.createElement('div');
    ghost.textContent = tema.length > 40 ? tema.slice(0,40)+'...' : tema;
    ghost.style.cssText = 'position:absolute;top:-1000px;padding:8px 16px;background:var(--primary-light);border-radius:8px;font-size:0.9rem;font-weight:600;border:2px solid var(--primary);max-width:300px;';
    document.body.appendChild(ghost);
    e.dataTransfer.setDragImage(ghost, 10, 10);
    setTimeout(() => document.body.removeChild(ghost), 0);
  });

  tbody.addEventListener('dragover', e => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    const tr = e.target.closest('tr[data-row]');
    // Clear all highlights
    tbody.querySelectorAll('.drag-over-row').forEach(el => el.classList.remove('drag-over-row'));
    if(tr) {
      const targetRow = parseInt(tr.dataset.row);
      if(targetRow !== dragSrcRow) tr.classList.add('drag-over-row');
    }
  });

  tbody.addEventListener('dragleave', e => {
    const tr = e.target.closest('tr[data-row]');
    if(tr) tr.classList.remove('drag-over-row');
  });

  tbody.addEventListener('dragend', e => {
    dragSrcRow = null;
    tbody.querySelectorAll('.dragging').forEach(el => el.classList.remove('dragging'));
    tbody.querySelectorAll('.drag-over-row').forEach(el => el.classList.remove('drag-over-row'));
  });

  tbody.addEventListener('drop', e => {
    e.preventDefault();
    tbody.querySelectorAll('.drag-over-row').forEach(el => el.classList.remove('drag-over-row'));
    const tr = e.target.closest('tr[data-row]'); if(!tr) return;
    const srcR = parseInt(e.dataTransfer.getData('text/plain'));
    const dstR = parseInt(tr.dataset.row);
    if(isNaN(srcR) || isNaN(dstR) || srcR===dstR) return;

    const src = part.rows[srcR], dst = part.rows[dstR];
    const sep = state.separator;

    // Объединяем тему и ДЗ в целевую строку
    if(dst.status==='cancelled') {
      // Если целевая отменена — просто переносим
      dst.tema = src.tema; dst.dz = src.dz; dst.status = 'normal';
    } else {
      dst.tema = [dst.tema, src.tema].filter(Boolean).join(sep);
      dst.dz = [dst.dz, src.dz].filter(Boolean).join(sep);
    }

    // Исходный урок — отменяется, но с возможностью восстановления
    preserveRowForRestore(src);
    src.tema = ''; src.dz = ''; src.status = 'cancelled';

    savePart(state.activeTab);
    drawPartTable(part, state.activeTab);
  });
}

function highlightRows() {
  document.querySelectorAll('#editor-tbody tr[data-row]').forEach(tr => {
    tr.classList.toggle('selected-row', state.selectedRows.has(parseInt(tr.dataset.row)));
  });
}

async function savePart(idx) {
  const part = state.parts[idx]; if(!part) return;
  try {
    await fetch(`${API}/session/${state.sessionId}/part/${idx}/update`, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({rows:part.rows})});
  } catch(err) { console.error('Ошибка сохранения:', err); }
}

async function doCancelSelected() {
  if(!state.selectedRows.size) return alert('Сначала выделите строки (кликните по ним)');
  const part = state.parts[state.activeTab];
  state.selectedRows.forEach(r => {
    const row = part.rows[r];
    preserveRowForRestore(row);
    row.status='cancelled'; row.tema=''; row.dz='';
  });
  await savePart(state.activeTab); drawPartTable(part, state.activeTab);
}

async function doRestoreSelected() {
  if(!state.selectedRows.size) return alert('Сначала выделите строки (кликните по ним)');
  const part = state.parts[state.activeTab];
  state.selectedRows.forEach(r => {
    const row = part.rows[r];
    row.status='normal';
    if(!row.tema && row._backupTema) row.tema = row._backupTema;
    if(!row.dz && row._backupDz) row.dz = row._backupDz;
  });
  await savePart(state.activeTab); drawPartTable(part, state.activeTab);
}

async function doInsertTemplate(tpl) {
  if(!state.selectedRows.size) return alert('Сначала выделите строки (кликните по ним)');
  const part = state.parts[state.activeTab];
  state.selectedRows.forEach(r => { if(part.rows[r].status!=='cancelled') part.rows[r].dz=tpl; });
  await savePart(state.activeTab); drawPartTable(part, state.activeTab);
}

async function doExport() {
  const btn = document.getElementById('btn-export');
  btn.disabled=true; btn.innerHTML='<span class="loading"></span> Экспортирую...';
  try {
    for(let i=0;i<state.parts.length;i++) if(state.parts[i]&&state.parts[i].rows) await savePart(i);
    const res = await fetch(`${API}/session/${state.sessionId}/export`, {method:'POST'});
    const data = await res.json(); if(!res.ok) throw new Error(data.error);

    document.getElementById('editor-content').innerHTML = `
      <div class="export-done">
        <div class="big-icon">\u2705</div>
        <h3>Файлы готовы!</h3>
        <p>Скачайте нужные файлы вручную по кнопкам ниже</p>
        <div style="margin-top:20px;display:flex;flex-direction:column;gap:14px;width:100%;max-width:400px">
          ${data.files.map(f => `<a href="${f.downloadUrl}" download="${f.fileName}" class="btn btn-primary btn-lg" style="text-decoration:none">&#128229; Скачать ${f.fileName}</a>`).join('')}
        </div>
      </div>`;
  } catch(err) { alert('Ошибка экспорта: '+err.message); }
  finally { btn.disabled=false; btn.innerHTML='&#128190; Экспорт в Excel'; }
}

// ============================================================
function esc(s) { const d=document.createElement('div'); d.textContent=String(s||''); return d.innerHTML; }
function preserveRowForRestore(row) {
  if(!row._backupTema && row.tema) row._backupTema = row.tema;
  if(!row._backupDz && row.dz) row._backupDz = row.dz;
}
window.addEventListener('DOMContentLoaded', initApp);
