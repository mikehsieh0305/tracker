// === 全域狀態 ===
let importedSheets = {};   // { sheetName: rows[] }
let sheetNames = [];
let currentSheetName = '';
let areaCapacityChart = null;
let materialStatusChart = null;

let rawRows = [];          // 選定 sheet 的原始列
let materialRows = [];     // 表1：材料進度
let progressRows = [];     // 表2：工程進度
let issueRows = [];        // 表3：阻礙

// 阻礙類型選單
const ISSUE_TYPES = [
  '',
  '魚塭、堤防問題',
  '路權問題',
  '水電供應不足',
  '雨季影響',
  '材料到貨 delay',
  '工班不足',
  '陳情抗議',
  '其他'
];

// === DOM 取得 ===
const btnImport = document.getElementById('btnImport');
const btnExport = document.getElementById('btnExport');
const btnAddMaterial = document.getElementById('btnAddMaterial');
const btnAddProgress = document.getElementById('btnAddProgress');
const btnAddIssue = document.getElementById('btnAddIssue');

const currentFileLabel = document.getElementById('currentFile');
const sheetSelect = document.getElementById('sheetSelect');

const materialTbody = document.querySelector('#materialTable tbody');
const progressTbody = document.querySelector('#progressTable tbody');
const issueTbody = document.querySelector('#issueTable tbody');

const tabButtons = document.querySelectorAll('.tab-btn');
const tabPanels = document.querySelectorAll('.tab-panel');

// === 綁定事件 ===
btnImport.addEventListener('click', onImport);
btnExport.addEventListener('click', onExport);
btnAddMaterial.addEventListener('click', onAddMaterialRow);
btnAddProgress.addEventListener('click', onAddProgressRow);
btnAddIssue.addEventListener('click', onAddIssueRow);

sheetSelect.addEventListener('change', onSheetChange);

tabButtons.forEach(btn => btn.addEventListener('click', onTabClick));

// ---- Tab 切換 ----
function onTabClick(e) {
  const tab = e.currentTarget.dataset.tab;
  tabButtons.forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  tabPanels.forEach(p => p.classList.toggle('active', p.dataset.tabPanel === tab));
}

// ---- 匯入 Excel ----
async function onImport() {
  const result = await window.excelAPI.importAllSheets();
  if (!result) return;

  importedSheets = result.sheets || {};
  sheetNames = result.sheetNames || [];
  currentSheetName = '';
  rawRows = [];
  materialRows = [];
  progressRows = [];
  issueRows = [];

  currentFileLabel.textContent = `目前檔案：${result.filePath}`;

  setupSheetSelect();
}

// 填 Sheet 下拉
function setupSheetSelect() {
  sheetSelect.innerHTML = '';

  if (!sheetNames.length) return;

  sheetNames.forEach((name, idx) => {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name;
    sheetSelect.appendChild(opt);
    if (idx === 0) currentSheetName = name;
  });

  sheetSelect.value = currentSheetName;
  applySheet(currentSheetName);
}

function onSheetChange() {
  const name = sheetSelect.value;
  currentSheetName = name;
  applySheet(name);
}

// 套用某個 sheet → 重建三種資料表
function applySheet(name) {
  rawRows = importedSheets[name] || [];
  materialRows = buildMaterialRowsFromRaw(rawRows);
  progressRows = buildProgressRowsFromRaw(rawRows);
  issueRows = buildIssueRowsFromRaw(rawRows);

  renderMaterialTable();
  renderProgressTable();
  renderIssueTable();
  updateCharts();
}

// ===== 共用小工具 =====
function toPercentCell(rate) {
  if (rate == null || rate === '' || isNaN(rate)) return '';
  const r = Number(rate);
  if (r >= 0.999) return '✔';
  if (r <= 0) return '❌';
  // return Math.round(r * 100) + '%';
  let p = (r * 100).toFixed(2);
  return p.endsWith('.00') ? p.replace('.00', '') + '%' : p + '%';
}

function formatDateCell(v) {
  if (!v) return '';
  // JS Date 物件 → yyyy-mm-dd
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  // Excel 日期序號（如 45322）
  if (typeof v === 'number' && v > 40000 && v < 60000) {
    // Excel 起始 1899-12-30
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(epoch.getTime() + v * 24 * 3600 * 1000);
    const y = date.getUTCFullYear();
    const m = String(date.getUTCMonth() + 1).padStart(2, '0');
    const d = String(date.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  // 其餘保持原樣但轉字串
  return String(v);
}

// ===== 1. 從 rawRows 生成 Material =====
// function buildMaterialRowsFromRaw(rows) {
//   return rows.map((row, idx) => {
//     const area = row['區域'] || '';
//     const kw = Number(row['容量(Kw)'] || 0);

//     if (!area && !kw) return null; // 完全空的列不要

//     const pileRate = Number(row['基樁發料完成率'] || 0);
//     const steelMainRate = Number(row['鋼構-大料發料完成率'] || 0);
//     const steelSubRate = Number(row['鋼構-小料發料完成率'] || 0);
//     const moduleRate = Number(row['模組發料完成率'] || 0);

//     const pileCell = toPercentCell(pileRate);
//     const steelMainCell = toPercentCell(steelMainRate);
//     const steelSubCell = toPercentCell(steelSubRate);
//     const moduleCell = toPercentCell(moduleRate);

//     // 狀態文字
//     const isMainZero = steelMainRate === 0;
//     const isSubZero = steelSubRate === 0;
//     const isModuleZero = moduleRate === 0;

//     let statusText = '正常';
//     if (isMainZero && isSubZero && isModuleZero) {
//       statusText = '嚴重缺料';
//     } else {
//       const lacks = [];
//       if (isMainZero) lacks.push('缺大料');
//       if (isSubZero) lacks.push('缺小料');
//       if (isModuleZero) lacks.push('缺模組');
//       if (lacks.length > 0) statusText = lacks.join('、');
//     }

//     return {
//       __index: idx,         // 如未來要對應回原始列可用
//       '區域': area,
//       'Kw': kw,
//       '基樁': pileCell,
//       '支架大料': steelMainCell,
//       '支架小料': steelSubCell,
//       '模組架': '',
//       '模組': moduleCell,
//       '狀態': statusText
//     };
//   }).filter(Boolean);
// }
function buildMaterialRowsFromRaw(rows) {
  return rows.map((row, idx) => {
    const area = row['區域'] || '';
    const kw = row['容量(Kw)'] || '';

    if (!area && !kw) return null; // 完全空列就丟掉

    const pileRate = Number(row['基樁發料完成率'] || 0);
    const steelMainRate = Number(row['鋼構-大料發料完成率'] || 0);
    const steelSubRate = Number(row['鋼構-小料發料完成率'] || 0);
    const moduleRate = Number(row['模組發料完成率'] || 0);

    const pileCell = toPercentCell(pileRate);
    const steelMainCell = toPercentCell(steelMainRate);
    const steelSubCell = toPercentCell(steelSubRate);
    const moduleCell = toPercentCell(moduleRate);

    const steelArrive = row['鋼構到料狀態'] || '';  // 例如「已到」、「未到」或日期
    const steelRemark = row['鋼構缺料說明'] || '';  // 例如「缺上構」

    // ===== 材料狀態 =====
    const isMainZero = steelMainRate === 0;
    const isSubZero = steelSubRate === 0;
    const isModuleZero = moduleRate === 0;

    let statusText = '正常';
    if (isMainZero && isSubZero && isModuleZero) {
      statusText = '嚴重缺料';
    } else {
      const lacks = [];
      if (isMainZero) lacks.push('缺大料');
      if (isSubZero) lacks.push('缺小料');
      if (isModuleZero) lacks.push('缺模組');
      if (lacks.length > 0) statusText = lacks.join('、');
    }

    return {
      __index: idx,
      '區域': area,
      '容量(kW)': kw,
      '基樁完成率': pileCell,
      '鋼構大料完成率': steelMainCell,
      '鋼構小料完成率': steelSubCell,
      '模組完成率': moduleCell,
      '鋼構到料狀態': steelArrive,
      '鋼構缺料說明': steelRemark,
      '材料狀態': statusText,
      '材料備註': row['備註/注意事項'] || ''
    };
  }).filter(Boolean);
}


// ===== 2. 從 rawRows 生成 Progress =====
// function buildProgressRowsFromRaw(rows) {
//   const today = new Date();

//   return rows.map((row, idx) => {
//     const area = row['區域'] || '';
//     const kw = Number(row['容量(Kw)'] || 0);
//     if (!area && !kw) return null;

//     const taskName = row['施工進度'] || '整體工程';
//     // const startDate = row['施工起始日'] || '';
//     // const planDate  = row['預計完工日期'] || '';
//     const startDate = formatDateCell(row['施工起始日']);
//     const planDate = formatDateCell(row['預計完工日期']);


//     let currentStatus = '進行中';
//     const memo = row['備註/注意事項'] || '';

//     if (typeof memo === 'string' && memo.includes('缺料')) {
//       currentStatus = '缺料無法動工';
//     }

//     // 燈號（簡易版）
//     let light = '🟢 正常';
//     if (!planDate) {
//       light = '⚪ 未排程';
//     } else {
//       const plan = new Date(planDate);
//       const diffDays = (plan - today) / (1000 * 3600 * 24);
//       if (today > plan) {
//         light = '🔴 延誤';
//       } else if (diffDays <= 7 && diffDays >= 0) {
//         light = '🟡 即將到期';
//       }
//     }

//     return {
//       __index: idx,
//       '區域': area,
//       '工項': taskName,
//       '起始': startDate,
//       '預計完工': planDate,
//       '現況': currentStatus,
//       '狀態': light
//     };
//   }).filter(Boolean);
// }
function buildProgressRowsFromRaw(rows) {
  const today = new Date();

  return rows.map((row, idx) => {
    const area = row['區域'] || '';
    const kw = row['容量(Kw)'] || '';

    if (!area && !kw) return null;

    const taskName = row['施工進度'] || '整體工程';
    const startDate = formatDateCell(row['施工起始日']);
    const planDate = formatDateCell(row['預計完工日期']);
    const protestDur = row['陳抗影響工進起訖日'] || '';
    const protestDays = row['陳抗影響工進實際天數'] || '';

    const memo = row['備註/注意事項'] || '';

    // ===== 現況說明：先用備註，如果空就用施工進度 =====
    let currentStatus = memo || taskName;

    // ===== 工期狀態 & 燈號 =====
    let scheduleStatus = '正常';
    let light = '🟢 正常';

    if (!planDate) {
      scheduleStatus = '未排程';
      light = '⚪ 未排程';
    } else {
      const plan = new Date(planDate);
      const diffDays = (plan - today) / (1000 * 3600 * 24);

      if (today > plan) {
        scheduleStatus = '延誤>0天';
        light = '🔴 延誤';
      } else if (diffDays <= 7 && diffDays >= 0) {
        scheduleStatus = '即將到期(7天內)';
        light = '🟡 即將到期';
      }
    }

    // 若有「陳抗實際影響天數」，補在 scheduleStatus 裡
    if (protestDays) {
      scheduleStatus += `，陳抗影響 ${protestDays} 天`;
    }

    return {
      __index: idx,
      '區域': area,
      '容量(kW)': kw,
      '工項': taskName,
      '施工起始日': startDate,
      '預計完工日期': planDate,
      '陳抗影響期間': protestDur,
      '陳抗實際影響天數': protestDays,
      '現況說明': currentStatus,
      '工期狀態': scheduleStatus,
      '狀態燈號': light
    };
  }).filter(Boolean);
}


// ===== 3. 從 rawRows 生成 IssueLog =====
// function buildIssueRowsFromRaw(rows) {
//   return rows.map((row, idx) => {
//     const area = row['區域'] || '';
//     const kw = Number(row['容量(Kw)'] || 0);
//     if (!area && !kw) return null;

//     const note = row['備註/注意事項'] || '';
//     const issueDate = row['陳抗問題發現之日期及對應日期'] || '';
//     const impactDays = row['陳抗影響工進實際天數'];
//     const improve = row['陳抗時困難點之對策及執行做法'] || '';

//     let impactText = '';
//     if (impactDays != null && impactDays !== '' && !isNaN(impactDays)) {
//       impactText = `延誤 ${impactDays} 天`;
//     }

//     let status = '';
//     if (note && String(note).trim() !== '') {
//       status = '進行中';
//     }

//     return {
//       __index: idx,
//       '區域': area,
//       '問題': note || '(尚未填寫)',
//       '發生日': issueDate,
//       '影響': impactText,
//       '設計變更': improve,
//       '狀態': status
//     };
//   }).filter(Boolean);
// }
function buildIssueRowsFromRaw(rows) {
  return rows.map((row, idx) => {
    const area = row['區域'] || '';
    const kw = row['容量(Kw)'] || '';

    if (!area && !kw) return null;

    const protestDate = formatDateCell(row['陳抗問題發現之日期及對應日期']);
    const protestDur = row['陳抗影響工進起訖日'] || '';
    const protestDays = row['陳抗影響工進實際天數'] || '';

    const poolImpact = row['交池地主影響之地號、容量、基裝數量'] || '';
    const memo = row['備註/注意事項'] || '';
    const improve = row['陳抗時困難點之對策及執行做法'] || '';
    const contact = row['陳抗人（地主/臨池養殖户）聯絡方式'] || '';
    const grWindow = row['GR對應窗口'] || '';

    // 問題內容：合併「交池地主影響...」+「備註/注意事項」
    const issues = [];
    if (poolImpact) issues.push(poolImpact);
    if (memo) issues.push(memo);
    const issueText = issues.join('\n');

    let impactText = '';
    if (protestDays !== '' && !isNaN(protestDays)) {
      impactText = `延誤 ${protestDays} 天`;
    }

    // 預設狀態：如果有問題內容就當作「進行中」，否則空白
    let status = '';
    if (issueText.trim()) status = '進行中';

    // 🔍 自動判斷「問題類型」
    const fullText = (memo + ' ' + improve).toLowerCase();
    let issueType = '';

    if (fullText.match(/魚塭|養殖|堤防|護岸/)) {
      issueType = '魚塭、堤防問題';
    } else if (fullText.match(/路權|道路用地|出入口|通行/)) {
      issueType = '路權問題';
    } else if (fullText.match(/水電|電力不足|用電不足|抽水電|變壓器/)) {
      issueType = '水電供應不足';
    } else if (fullText.match(/雨季|豪雨|降雨|天候|氣候|颱風/)) {
      issueType = '雨季影響';
    } else if (fullText.match(/材料|到貨|交期|delay|延遲出貨/)) {
      issueType = '材料到貨 delay';
    } else if (fullText.match(/工班|人力不足|人手不足|缺工/)) {
      issueType = '工班不足';
    } else if (fullText.match(/陳情|抗議|請願/)) {
      issueType = '陳情抗議';
    } else if (memo || improve) {
      issueType = '其他';
    }

    return {
      __index: idx,
      '區域': area,
      '容量(kW)': kw,
      '問題類型': issueType,
      '問題發現日期': protestDate,
      '問題內容': issueText || '(尚未填寫)',
      '影響期間': protestDur,
      '實際影響天數': protestDays,
      '影響說明': impactText,
      '改善措施': improve,
      '地主/養殖戶聯絡方式': contact,
      'GR窗口': grWindow,
      '狀態': status
    };
  }).filter(Boolean);
}


// ===== 三張表的 render + 編輯回寫 =====
// function renderMaterialTable() {
//   materialTbody.innerHTML = '';
//   materialRows.forEach((row, idx) => {
//     const tr = document.createElement('tr');

//     function cell(field, type = 'text') {
//       const td = document.createElement('td');
//       if (field === '#') {
//         td.textContent = idx + 1;
//         return td;
//       }
//       const input = document.createElement('input');
//       input.type = type;
//       input.value = row[field] ?? '';
//       input.dataset.kind = 'material';
//       input.dataset.index = idx;
//       input.dataset.field = field;
//       input.addEventListener('change', onCellChange);
//       td.appendChild(input);
//       return td;
//     }

//     // tr.appendChild(cell('#'));
//     // tr.appendChild(cell('區域'));
//     // tr.appendChild(cell('Kw', 'number'));
//     // tr.appendChild(cell('基樁'));
//     // tr.appendChild(cell('支架大料'));
//     // tr.appendChild(cell('支架小料'));
//     // tr.appendChild(cell('模組架'));
//     // tr.appendChild(cell('模組'));
//     // tr.appendChild(cell('狀態'));
//      tr.appendChild(cell('#'));
//     tr.appendChild(cell('區域'));
//     tr.appendChild(cell('容量(kW)', 'number'));
//     tr.appendChild(cell('基樁完成率'));
//     tr.appendChild(cell('鋼構大料完成率'));
//     tr.appendChild(cell('鋼構小料完成率'));
//     tr.appendChild(cell('模組完成率'));
//     tr.appendChild(cell('鋼構到料狀態'));
//     tr.appendChild(cell('鋼構缺料說明'));
//     tr.appendChild(cell('材料狀態'));
//     tr.appendChild(cell('材料備註'));
//     materialTbody.appendChild(tr);
//   });
// }

// function renderProgressTable() {
//   progressTbody.innerHTML = '';
//   progressRows.forEach((row, idx) => {
//     const tr = document.createElement('tr');

//     function cell(field, type = 'text') {
//       const td = document.createElement('td');
//       if (field === '#') {
//         td.textContent = idx + 1;
//         return td;
//       }
//       const input = document.createElement('input');
//       input.type = type;
//       input.value = row[field] ?? '';
//       input.dataset.kind = 'progress';
//       input.dataset.index = idx;
//       input.dataset.field = field;
//       input.addEventListener('change', onCellChange);
//       td.appendChild(input);
//       return td;
//     }

//     // tr.appendChild(cell('#'));
//     // tr.appendChild(cell('區域'));
//     // tr.appendChild(cell('工項'));
//     // tr.appendChild(cell('起始'));
//     // tr.appendChild(cell('預計完工'));
//     // tr.appendChild(cell('現況'));
//     // tr.appendChild(cell('狀態'));
//     tr.appendChild(cell('#'));
//     tr.appendChild(cell('區域'));
//     tr.appendChild(cell('容量(kW)'));
//     tr.appendChild(cell('工項'));
//     tr.appendChild(cell('施工起始日'));
//     tr.appendChild(cell('預計完工日期'));
//      tr.appendChild(cell('陳抗影響期間'));
//     tr.appendChild(cell('陳抗實際影響天數'));
//     tr.appendChild(cell('現況說明'));
//     tr.appendChild(cell('工期狀態'));
//     tr.appendChild(cell('狀態燈號'));
//     progressTbody.appendChild(tr);
//   });
// }

// function renderIssueTable() {
//   issueTbody.innerHTML = '';
//   issueRows.forEach((row, idx) => {
//     const tr = document.createElement('tr');

//     function cell(field, type = 'text') {
//       const td = document.createElement('td');
//       if (field === '#') {
//         td.textContent = idx + 1;
//         return td;
//       }
//       const input = document.createElement('input');
//       input.type = type;
//       input.value = row[field] ?? '';
//       input.dataset.kind = 'issue';
//       input.dataset.index = idx;
//       input.dataset.field = field;
//       input.addEventListener('change', onCellChange);
//       td.appendChild(input);
//       return td;
//     }

//     tr.appendChild(cell('#'));
//     tr.appendChild(cell('區域'));
//     tr.appendChild(cell('容量(kW)'));
//     tr.appendChild(cell('問題發現日期'));
//     tr.appendChild(cell('問題內容'));
//     tr.appendChild(cell('影響期間'));
//     tr.appendChild(cell('實際影響天數'));
//     tr.appendChild(cell('影響說明'));
//     tr.appendChild(cell('改善措施'));
//     tr.appendChild(cell('地主/養殖戶聯絡方式'));
//     tr.appendChild(cell('GR窗口'));
//     tr.appendChild(cell('狀態'));
//     issueTbody.appendChild(tr);
//   });
// }

// 🔽 問題類型 select
function cellIssueType() {
  const td = document.createElement('td');
  const select = document.createElement('select');
  ISSUE_TYPES.forEach(optVal => {
    const opt = document.createElement('option');
    opt.value = optVal;
    opt.textContent = optVal || '（未分類）';
    if ((row['問題類型'] || '') === optVal) opt.selected = true;
    select.appendChild(opt);
  });
  select.dataset.kind = 'issue';
  select.dataset.index = idx;
  select.dataset.field = '問題類型';
  select.addEventListener('change', onCellChange);
  td.appendChild(select);
  return td;
}

function renderTable(tbody, rows, kind) {
  tbody.innerHTML = '';

  rows.forEach((row, idx) => {
    const tr = document.createElement('tr');

    // ==== 前導第0欄：序號 ====
    const tdIndex = document.createElement('td');
    tdIndex.textContent = idx + 1;
    tr.appendChild(tdIndex);

    // ==== 動態生成其它欄 ====
    Object.keys(row).forEach(field => {
      if (field === '__index') return; // 忽略內部欄位

      const td = document.createElement('td');
      const input = document.createElement('input');

      input.type = 'text';
      input.value = row[field] ?? '';
      input.dataset.kind = kind;
      input.dataset.index = idx;
      input.dataset.field = field;
      input.addEventListener('change', onCellChange);

      td.appendChild(input);
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });
}

function renderMaterialTable() {
  renderTable(materialTbody, materialRows, 'material');
  updateCharts();   // 材料更新後要更新圖
}

function renderProgressTable() {
  renderTable(progressTbody, progressRows, 'progress');
}

function renderIssueTable() {
  renderTable(issueTbody, issueRows, 'issue');
}

function computeMaterialStatus(row) {
  function toNum(v) {
    if (v === '✔') return 1;
    if (v === '❌') return 0;
    if (typeof v === 'string' && v.endsWith('%')) return Number(v.replace('%', '')) / 100;
    return Number(v) || 0;
  }

  const pile = toNum(row['基樁完成率']);
  const main = toNum(row['鋼構大料完成率']);
  const sub = toNum(row['鋼構小料完成率']);
  const module = toNum(row['模組完成率']);

  const rates = [pile, main, sub, module];
  const zeroCount = rates.filter(v => v === 0).length;

  if (zeroCount === 4) return '嚴重缺料';
  if (zeroCount >= 1) return '缺料';
  if (rates.some(v => v < 1)) return '未完成';
  return '正常';
}

function updateCharts() {
  const ctxCapacity = document.getElementById('areaCapacityChart');
  const ctxStatus = document.getElementById('materialStatusChart');
  if (!ctxCapacity || !ctxStatus) return;

  // 1) 各區域容量 (kW) 長條圖
  const labels = materialRows.map(r => r['區域']).filter(x => x);
  const dataKw = materialRows.map(r => Number(r['容量(kW)'] || 0));

  // 如果之前有 chart 先銷毀
  if (areaCapacityChart) areaCapacityChart.destroy();
  areaCapacityChart = new Chart(ctxCapacity, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '容量 (kW)',
        data: dataKw
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: true },
        tooltip: { enabled: true }
      },
      scales: {
        x: { title: { display: true, text: '區域' } },
        y: { title: { display: true, text: '容量(kW)' }, beginAtZero: true }
      }
    }
  });

  // 2) 材料狀態分佈 圓餅圖
  const statusCountMap = {};  // { '正常':3, '缺大料':2, '嚴重缺料':1 ... }
  materialRows.forEach(r => {
    const s = (r['材料狀態'] || '').trim() || '未標註';
    // const status = computeMaterialStatus(r);
    statusCountMap[s] = (statusCountMap[s] || 0) + 1;
    // statusCountMap[status] = (statusCountMap[status] || 0) + 1;
  });

  const statusLabels = Object.keys(statusCountMap);
  const statusData = statusLabels.map(k => statusCountMap[k]);

  if (materialStatusChart) materialStatusChart.destroy();
  materialStatusChart = new Chart(ctxStatus, {
    type: 'pie',
    data: {
      labels: statusLabels,
      datasets: [{
        data: statusData
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: 'bottom' }
      }
    }
  });
}

// 編輯回寫
function onCellChange(e) {
  const input = e.target;
  const kind = input.dataset.kind;
  const idx = Number(input.dataset.index);
  const field = input.dataset.field;
  const value = input.value;

  if (kind === 'material') {
    materialRows[idx][field] = value;
    updateCharts();
  } else if (kind === 'progress') {
    progressRows[idx][field] = value;
  } else if (kind === 'issue') {
    issueRows[idx][field] = value;
  }
}

// ===== 新增列功能 =====
function addEmptyRowFromTemplate(arr) {
  const template = {};
  Object.keys(arr[0] || {}).forEach(k => template[k] = '');
  delete template.__index;
  return template;
}

function onAddMaterialRow() {
  const newRow = addEmptyRowFromTemplate(materialRows);
  newRow.__index = -1;
  materialRows.push(newRow);
  renderMaterialTable();
  updateCharts();
}

function onAddProgressRow() {
  const newRow = addEmptyRowFromTemplate(progressRows);
  newRow.__index = -1;
  progressRows.push(newRow);
  renderProgressTable();
}

function onAddIssueRow() {
  const newRow = addEmptyRowFromTemplate(issueRows);
  newRow.__index = -1;
  issueRows.push(newRow);
  renderIssueTable();
}

// function onAddMaterialRow() {
//   materialRows.push({
//     __index: -1,
//     '區域': '',
//     '容量(kW)': '',
//     '基樁完成率': '',
//     '鋼構大料完成率': '',
//     '鋼構小料完成率': '',
//     '模組完成率': '',
//     '鋼構到料狀態': '',
//     '鋼構缺料說明': '',
//     '材料狀態': '',
//     '材料備註': ''
//   });
//   renderMaterialTable();
//   updateCharts();
// }

// function onAddProgressRow() {
//   progressRows.push({
//     __index: -1,
//     '區域': '',
//     '容量(kW)': '',
//     '工項': '',
//     '施工起始日': '',
//     '預計完工日期': '',
//     '陳抗影響期間': '',
//     '陳抗實際影響天數': '',
//     '現況說明': '',
//     '工期狀態': '',
//     '狀態燈號': '⚪ 未排程'
//   });
//   renderProgressTable();
// }

// function onAddIssueRow() {
//   issueRows.push({
//     __index: -1,
//     '區域': '',
//     '容量(kW)': '',
//     '問題發現日期': '',
//     '問題內容': '',
//     '影響期間': '',
//     '實際影響天數': '',
//     '影響說明': '',
//     '改善措施': '',
//     '地主/養殖戶聯絡方式': '',
//     'GR窗口': '',
//     '狀態': ''
//   });
//   renderIssueTable();
// }

// ===== 匯出前的重算 =====
// function recomputeMaterialStatus(rows) {
//   rows.forEach(r => {
//     const main = r['支架大料'] || '';
//     const sub = r['支架小料'] || '';
//     const mod = r['模組'] || '';

//     const isZero = (v) =>
//       v === '❌' ||
//       v === '' ||
//       (typeof v === 'string' && v.endsWith('%') && Number(v.replace('%', '')) === 0);

//     const isMainZero = isZero(main);
//     const isSubZero = isZero(sub);
//     const isModZero = isZero(mod);

//     if (isMainZero && isSubZero && isModZero) {
//       r['狀態'] = '嚴重缺料';
//     } else {
//       const lacks = [];
//       if (isMainZero) lacks.push('缺大料');
//       if (isSubZero) lacks.push('缺小料');
//       if (isModZero) lacks.push('缺模組');
//       r['狀態'] = lacks.length ? lacks.join('、') : '正常';
//     }
//   });
//   return rows;
// }
function recomputeMaterialStatus(rows) {
  rows.forEach(r => {
    const main = r['鋼構大料完成率'] || '';
    const sub = r['鋼構小料完成率'] || '';
    const mod = r['模組完成率'] || '';
    const arrive = r['鋼構到料狀態'] || '';

    const isZero = (v) =>
      v === '❌' ||
      v === '' ||
      (typeof v === 'string' && v.endsWith('%') && Number(v.replace('%', '')) === 0);

    const isMainZero = isZero(main);
    const isSubZero = isZero(sub);
    const isModZero = isZero(mod);
    const isSteelNotArrived =
      !arrive || arrive.includes('未') || arrive.includes('待') || arrive.includes('無');

    // ===== 狀態推論 =====
    if ((isMainZero && isSubZero && isModZero) || isSteelNotArrived) {
      r['材料狀態'] = '嚴重缺料';
    } else {
      const lacks = [];
      if (isMainZero) lacks.push('缺大料');
      if (isSubZero) lacks.push('缺小料');
      if (isModZero) lacks.push('缺模組');
      if (lacks.length) {
        r['材料狀態'] = lacks.join('、');
      } else {
        r['材料狀態'] = '正常';
      }
    }
  });
  return rows;
}


// function recomputeProgressLights(rows) {
//   const today = new Date();
//   rows.forEach(r => {
//     const txt = r['現況'] || '';
//     const plan = r['預計完工'];
//     let light = '🟢 正常';

//     if (txt.includes('完成') || txt.includes('完工')) {
//       light = '🟢 完成';
//     } else if (!plan) {
//       light = '⚪ 未排程';
//     } else {
//       const planDate = new Date(plan);
//       const diffDays = (planDate - today) / (1000 * 3600 * 24);
//       if (today > planDate) {
//         light = '🔴 延誤';
//       } else if (diffDays <= 7 && diffDays >= 0) {
//         light = '🟡 即將到期';
//       }
//     }

//     r['狀態'] = light;
//   });
//   return rows;
// }
function recomputeProgressLights(rows) {
  const today = new Date();
  rows.forEach(r => {
    const memo = r['現況說明'] || '';
    const plan = r['預計完工日期'];
    const delay = Number(r['陳抗實際影響天數'] || 0);

    // ===== 狀態燈號 =====
    let light = '🟢 正常';

    if (!plan) {
      light = '⚪ 未排程';
    } else {
      const planDate = new Date(plan);
      const diffDays = (planDate - today) / (1000 * 3600 * 24);

      if (today > planDate) {
        light = '🔴 延誤';
      } else if (diffDays <= 7 && diffDays >= 0) {
        light = '🟡 即將到期';
      }
    }

    // 如果備註包含「缺料」
    if (memo.includes('缺料') || memo.includes('未到') || memo.includes('無料')) {
      light = '🔴 缺料停工';
    }

    // ===== 工期狀態文字 =====
    let scheduleStatus = '正常';

    if (!plan) {
      scheduleStatus = '未排程';
    } else {
      if (today > new Date(plan)) {
        scheduleStatus = '延誤中';
      } else {
        const planDate = new Date(plan);
        const diffDays = (planDate - today) / (1000 * 3600 * 24);
        if (diffDays <= 7 && diffDays >= 0) {
          scheduleStatus = '即將到期(7天內)';
        }
      }
    }

    // 加上陳抗天數說明
    if (delay > 0) {
      scheduleStatus += `、陳抗影響 ${delay} 天`;
    }

    // ===== 寫回 =====
    r['狀態燈號'] = light;
    r['工期狀態'] = scheduleStatus;
  });
  return rows;
}

// function buildGanttRowsFromProgress(rows) {
//   return rows.map(r => {
//     const start = r['起始'] ? new Date(r['起始']) : null;
//     const plan = r['預計完工'] ? new Date(r['預計完工']) : null;

//     const daysPlan = start && plan
//       ? (plan - start) / (1000 * 3600 * 24)
//       : '';

//     return {
//       '區域': r['區域'],
//       '工項': r['工項'],
//       '施工起始': r['起始'],
//       '預計完工': r['預計完工'],
//       '實際完工': '',
//       '工期天數(預計)': daysPlan,
//       '工期天數(實際)': ''
//     };
//   });
// }

function splitDateRange(rangeStr) {
  if (!rangeStr) return { start: '', end: '' };

  const normalized = rangeStr
    .replace(/至|—|～|-/g, '~') // 將各種可能的符號轉成 ~
    .replace(/\s+/g, '');        // 去空白

  const parts = normalized.split('~');
  return {
    start: parts[0] || '',
    end: parts[1] || ''
  };
}

function buildGanttRowsFromProgress(rows) {
  return rows.map(r => {
    const startStr = r['施工起始日'] || '';
    const planStr = r['預計完工日期'] || '';

    const start = startStr ? new Date(startStr) : null;
    const plan = planStr ? new Date(planStr) : null;

    const daysPlan = (start && plan)
      ? (plan - start) / (1000 * 3600 * 24)
      : '';

    // 解析陳抗期間
    const protestRange = splitDateRange(r['陳抗影響期間']);
    const protestStart = protestRange.start;
    const protestEnd = protestRange.end;
    const autoProtestDays = (protestStart && protestEnd)
      ? (new Date(protestEnd) - new Date(protestStart)) / (1000 * 3600 * 24)
      : '';

    // return {
    //   '區域': r['區域'],
    //   '工項': r['工項'],
    //   '施工起始': startStr,
    //   '預計完工': planStr,
    //   '實際完工': '',
    //   '工期天數(預計)': daysPlan,
    //   '工期天數(實際)': '',
    //   '陳抗實際影響天數': r['陳抗實際影響天數'] || ''
    // };
    return {
      '案件編號': r['區域'],
      '開始日': protestStart,
      '持續天數': autoProtestDays,
      '陳抗截止日': protestEnd,
      // '預計完工': planStr,
      // '實際完工': '',
      // '工期天數(預計)': daysPlan,
      // '工期天數(實際)': '',
      // '陳抗實際影響天數': r['陳抗實際影響天數'] || ''
    };
  });
}

// 移除每列中的 __index 等內部欄位
function stripInternalFields(rows) {
  return rows.map(r => {
    const copy = { ...r };
    delete copy.__index;
    return copy;
  });
}


// ===== 匯出 =====
async function onExport() {
  if (!materialRows.length && !progressRows.length && !issueRows.length) {
    alert('尚未有任何資料可以匯出');
    return;
  }

  // const matForExport = recomputeMaterialStatus(
  //   JSON.parse(JSON.stringify(materialRows))
  // );
  // const progForExport = recomputeProgressLights(
  //   JSON.parse(JSON.stringify(progressRows))
  // );
  // const issueForExport = JSON.parse(JSON.stringify(issueRows));
  // 先深拷貝
  const matCopy = JSON.parse(JSON.stringify(materialRows));
  const progCopy = JSON.parse(JSON.stringify(progressRows));
  const issueCopy = JSON.parse(JSON.stringify(issueRows));

  // 重算狀態
  const matForExport = stripInternalFields(recomputeMaterialStatus(matCopy));
  const progForExport = stripInternalFields(recomputeProgressLights(progCopy));
  const issueForExport = stripInternalFields(issueCopy);
  const ganttRows = buildGanttRowsFromProgress(progForExport);

  const res = await window.excelAPI.exportSummary({
    materialRows: matForExport,
    progressRows: progForExport,
    issueRows: issueForExport,
    ganttRows
  });

  if (res && res.ok) {
    alert('已匯出：\n' + res.filePath);
  }
}
