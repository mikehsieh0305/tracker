const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const XLSX = require('xlsx');

function createWindow() {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  });

  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

/**
 * 匯入 Excel：回傳所有 sheet 名稱 + 每個 sheet 的 rows
 */
ipcMain.handle('excel:importAllSheets', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    title: '選擇漁電共生 Excel',
    filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
    properties: ['openFile']
  });

  if (canceled || !filePaths.length) return null;

  const filePath = filePaths[0];
  const wb = XLSX.readFile(filePath);

  const sheets = {};
  wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    // sheets[name] = XLSX.utils.sheet_to_json(ws, { defval: '' });
    sheets[name] = XLSX.utils.sheet_to_json(ws, {
      defval: '',
      raw: true   // 保留有效數值，不要自動格式化
    });
    // sheets[name] = XLSX.utils.sheet_to_json(ws, {
    //   defval: '',
    //   raw: false,              // 不要回傳原始 Date/數字，改用格式化文字
    //   dateNF: 'yyyy-mm-dd'     // 日期統一成 2025-08-19 這種格式
    // });
  });

  return {
    filePath,
    sheetNames: wb.SheetNames,
    sheets
  };
});

/**
 * 匯出 4 個 Sheet：
 * - Material（材料進度）
 * - Progress（工程進度）
 * - IssueLog（阻礙）
 * - Gantt（甘特圖用）
 */
ipcMain.handle('excel:exportSummary', async (event, payload) => {
  const {
    materialRows,
    progressRows,
    issueRows,
    ganttRows
  } = payload || {};

  const { canceled, filePath } = await dialog.showSaveDialog({
    title: '匯出漁電共生 工程追蹤 Excel',
    defaultPath: '漁電共生_工程追蹤.xlsx',
    filters: [{ name: 'Excel', extensions: ['xlsx'] }]
  });

  if (canceled || !filePath) return { ok: false };

  const wb = XLSX.utils.book_new();

  function addSheet(rows, headers, name) {
    const ws = XLSX.utils.json_to_sheet(rows || [], {
      header: headers,
      skipHeader: false
    });
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  // Sheet1: 材料進度
  // addSheet(
  //   materialRows,
  //   ['區域', 'Kw', '基樁', '支架大料', '支架小料', '模組架', '模組', '狀態'],
  //   '材料進度'
  // );
  addSheet(
    materialRows,
    ['區域', '容量(kW)', '基樁完成率', '鋼構大料完成率', '鋼構小料完成率', '模組完成率', '鋼構到料狀態', '鋼構缺料說明', '材料狀態', '材料備註'],
    '材料進度'
  );

  // Sheet2: 工程進度
  addSheet(
    progressRows,
    ['區域', '容量(kW)', '工項', '施工起始日', '預計完工日期', '陳抗影響期間', '陳抗實際影響天數', '現況說明', '工期狀態', '狀態燈號'],
    '工程進度'
  );

  // Sheet3: 阻礙 IssueLog
  addSheet(
    issueRows,
    ['區域', '容量(kW)', '問題類型', '問題發現日期', '問題內容', '影響期間', '實際影響天數', '影響說明', '改善措施', '地主/養殖戶聯絡方式', 'GR窗口', '狀態'],
    '阻礙與風險表'
  );

  // Sheet4: Gantt
  // addSheet(
  //   ganttRows,
  //   ['區域', '工項', '施工起始', '預計完工', '實際完工', '工期天數(預計)', '工期天數(實際)', '陳抗實際影響天數'],
  //   '甘特圖'
  // );
  addSheet(
    ganttRows,
    ['案件編號', '開始日', '持續天數', '陳抗截止日'],
    '陳抗甘特圖'
  );

  XLSX.writeFile(wb, filePath);
  return { ok: true, filePath };
});
