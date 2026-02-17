/**
 * AI建築見積システム v10.0
 * Code.gs - 完全版 (Phase 4完了 + Performance Tuning)
 */

const scriptProps = PropertiesService.getScriptProperties();
const PROPS = scriptProps ? scriptProps.getProperties() || {} : {};

const CONFIG = {
  API_KEY:             PROPS.GEMINI_API_KEY || '', 
  inputFolder:         PROPS.FOLDER_INPUT || '',        
  invoiceInputFolder: PROPS.FOLDER_INVOICE_INPUT || '', 
  saveFolder:          PROPS.FOLDER_SAVE || '',
  logoFileId:          PROPS.IMAGE_FILE_ID || '',
  sheetNames: {
    list: '見積リスト',
    order: '発注リスト',
    invoice: '受取請求書リスト',
    deposits: '入金リスト',
    payments: '出金リスト',
    masterBasic: '基本単価マスタ',
    masterClient: '元請別単価マスタ',
    masterSet: '見積セットマスタ',
    masterVendor: '発注先マスタ',
    journalConfig: '仕訳設定マスタ'
  }
};

const CACHE_TTL = 1500;
const CACHE_TTL_SHORT = 120;  // 2分（案件・発注等の更新頻度考慮）
const CACHE_TTL_ORDERS = 60;  // 1分（発注一覧の更新頻度を高める）

function invalidateDataCache_() {
  try {
    const c = CacheService.getScriptCache();
    c.remove("projects_data");
    c.remove("orders_data");
    c.remove("active_projects_data");
    c.remove("deposits_data");
    c.remove("payments_data");
    c.remove("masters_data");
    c.remove("products_data");
    const y = new Date().getFullYear();
    for (let i = y - 2; i <= y + 1; i++) c.remove("analysis_" + i);
  } catch (e) { /* ignore */ }
}

let _saveFolderCache = null;
function getSaveFolder() {
  if (_saveFolderCache) return _saveFolderCache;
  if (!CONFIG.saveFolder) return DriveApp.getRootFolder();
  try {
    _saveFolderCache = DriveApp.getFolderById(CONFIG.saveFolder);
    return _saveFolderCache;
  } catch (e) {
    return DriveApp.getRootFolder();
  }
}

// -----------------------------------------------------------
// ヘルパー関数 & 安全なInclude
// -----------------------------------------------------------

function include(filename) {
  try {
    var name = (filename != null && String(filename).trim() !== '') ? String(filename).trim() : 'logo';
    return HtmlService.createHtmlOutputFromFile(name).getContent();
  } catch (e) {
    console.warn("Template include failed: " + filename + " (" + (e && e.message) + ")");
    if (CONFIG.logoFileId) {
      try {
        return 'https://drive.google.com/uc?export=view&id=' + CONFIG.logoFileId;
      } catch (e2) { /* ignore */ }
    }
    return "";
  }
}

function parseCurrency(val) {
  if (!val) return 0;
  let str = String(val);
  str = str.replace(/[０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  const num = Number(str.replace(/[^0-9.-]+/g, ""));
  return isNaN(num) ? 0 : num;
}

function toHalfWidth(str) {
  if (!str) return "";
  return String(str).replace(/[！-～]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
}

function formatDate(d) { 
  try { 
    if (!d) return ""; 
    return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), "yyyy/MM/dd"); 
  } catch(e) { return d; } 
}

function getJapaneseDateStr(date) {
  try {
    const d = new Date(date);
    const year = d.getFullYear();
    const month = d.getMonth() + 1;
    const day = d.getDate();
    if (year > 2019 || (year === 2019 && month >= 5)) {
      const reiwaYear = year - 2018;
      return `令和${reiwaYear === 1 ? '元' : reiwaYear}年${month}月${day}日`;
    }
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy年MM月dd日");
  } catch (e) {
    return "";
  }
}

function getNextSequenceId(type) {
  const props = PropertiesService.getScriptProperties();
  const key = type === 'estimate' ? 'SEQ_ESTIMATE' : 'SEQ_ORDER';
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = lock.tryLock(5000);
    if (lockAcquired) {
      let current = Number(props.getProperty(key)) || 0;
      current++;
      props.setProperty(key, String(current));
      const seq = String(current).padStart(7, '0');
      return `${seq}-00`;
    } else {
      throw new Error("ID採番タイムアウト");
    }
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

/**
 * 高速削除ヘルパー
 * 連続する行をまとめて削除することでAPIコール数を削減
 */
function deleteRowsOptimized_(sheet, rows) {
  if (!sheet || !rows || rows.length === 0) return;
  
  // 行番号でソート (昇順)
  rows.sort((a, b) => a - b);
  
  const groups = [];
  let start = rows[0];
  let count = 1;
  
  for (let i = 1; i < rows.length; i++) {
    if (rows[i] === start + count) {
      count++;
    } else {
      groups.push({ start, count });
      start = rows[i];
      count = 1;
    }
  }
  groups.push({ start, count });
  
  // 下の行から削除しないとインデックスがずれるため逆順で実行
  for (let i = groups.length - 1; i >= 0; i--) {
    sheet.deleteRows(groups[i].start, groups[i].count);
  }
}

function deleteRowsById(sheet, targetId) {
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const rowsToDelete = [];
  let currentId = "";
  
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0]).trim(); 
    if (rowId !== "") {
      currentId = rowId;
    } else {
      // ID列が空行の場合、次にIDが見つかる行まで先読みして判断
      // 直前のIDが対象外なら、この行はスキップ
      if (currentId !== targetId) continue;
    }
    
    if (currentId === targetId) {
      // i=0 is header(row1), so data[i] is row i+1
      rowsToDelete.push(i + 1);
    }
  }
  
  if (rowsToDelete.length === 0) return false;
  
  // 高速削除実行
  deleteRowsOptimized_(sheet, rowsToDelete);
  return true;
}

function paginateItems(items, rowsPerPage, rowsPerPageSubsequent) {
  const firstPageRows = rowsPerPage;
  const nextPageRows = rowsPerPageSubsequent || rowsPerPage;
  const pages = [];
  const targetItems = (items && Array.isArray(items) && items.length > 0) ? items : [];
  const queue = targetItems.map(item => ({ ...item }));
  let isFirst = true;
  while (queue.length > 0) {
    const limit = isFirst ? firstPageRows : nextPageRows;
    const chunk = queue.splice(0, limit);
    while (chunk.length < limit) {
      chunk.push({ isPadding: true });
    }
    pages.push(chunk);
    isFirst = false;
  }
  if (pages.length === 0) {
      const chunk = [];
      for(let i=0; i<firstPageRows; i++) chunk.push({ isPadding: true });
      pages.push(chunk);
  }
  return pages;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('AI建築見積システム v10.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// -----------------------------------------------------------
// シート初期化
// -----------------------------------------------------------

function checkAndFixOrderHeader(sheet) {
  if (!sheet) return;
  const headers = ["ID", "日付", "発注先", "関連見積ID", "工種", "品名", "仕様", "数量", "単位", "単価", "金額", "納品場所", "状態", "備考", "作成者", "公開範囲"];
  if (sheet.getLastRow() === 0) { sheet.appendRow(headers); }
}

function checkAndFixInvoiceHeader(sheet) {
  if (!sheet) return;
  const headers = ["ID", "ステータス", "登録日時", "ファイルID", "工事ID", "工事名", "請求元", "請求日", "請求金額", "相殺額", "支払予定額", "内容", "備考", "登録番号"];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
}

function checkAndFixJournalConfig(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["出力項目名(CSVヘッダー)", "データソース", "固定値/フォーマット/デフォルト", "順序", "タイプ(仕入/売上/共通)", "", "【集計対象 取引先名 (売上)】", "【集計対象 取引先名 (仕入)】"]);
    const defaults = [
      ["取引先名", "client", "", 1, "売上"], ["前月繰越", "fixed", "0", 2, "売上"], ["当月発生高", "amount", "", 3, "売上"],
      ["当月値引割引高", "fixed", "0", 4, "売上"], ["現金・小切手(入金・支払)高", "cash_check", "", 5, "売上"], ["手　形", "bill", "", 6, "売上"], 
      ["相　殺", "fixed", "0", 7, "売上"], ["振込料", "fixed", "0", 8, "売上"], ["その他", "other", "", 9, "売上"], ["翌月繰越高", "fixed", "0", 10, "売上"], 
      ["取引先名", "supplier", "", 1, "仕入"], ["前月繰越", "fixed", "0", 2, "仕入"], ["当月発生高", "amount", "", 3, "仕入"],
      ["当月値引割引高", "fixed", "0", 4, "仕入"], ["現金・小切手(入金・支払)高", "cash_check", "", 5, "仕入"], ["手　形", "bill", "", 6, "仕入"],
      ["相　殺", "offset", "", 7, "仕入"], ["振込料", "fixed", "0", 8, "仕入"], ["その他", "other", "", 9, "仕入"], ["翌月繰越高", "fixed", "0", 10, "仕入"]
    ];
    sheet.getRange(2, 1, defaults.length, 5).setValues(defaults);
  }
}

function checkAndFixDepositsHeader(sheet) {
  if (!sheet) return;
  const headers = ["ID", "登録日時", "入金日", "関連見積ID", "取引先名", "工事名", "入金種別", "入金金額", "振込手数料", "相殺金額", "備考", "ステータス", "登録者", "公開範囲"];
  if (sheet.getLastRow() === 0) { sheet.appendRow(headers); }
}

function checkAndFixPaymentsHeader(sheet) {
  if (!sheet) return;
  const headers = ["ID", "登録日時", "出金日", "関連発注ID", "関連請求書ID", "支払先名", "工事名", "出金種別", "出金金額", "振込手数料", "相殺金額", "備考", "ステータス", "登録者", "公開範囲"];
  if (sheet.getLastRow() === 0) { sheet.appendRow(headers); }
}

// -----------------------------------------------------------
// 共通・マスタ系 API
// -----------------------------------------------------------

function apiGetAuthStatus() {
  try {
    const email = Session.getActiveUser().getEmail().toLowerCase();
    const props = PropertiesService.getScriptProperties();
    const adminStr = props.getProperty('ADMIN_USERS') || "";
    const admins = adminStr.split(',').map(function(e) { return e.trim().toLowerCase(); });
    const isAdmin = admins.includes(email);
    console.log("Login: " + email + ", Admin: " + isAdmin);
    return JSON.stringify({ isAdmin: isAdmin, email: email });
  } catch (e) {
    return JSON.stringify({ isAdmin: false, email: "unknown", error: e.toString() });
  }
}

function apiGetMasters() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("masters_data");
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mSheet = ss.getSheetByName(CONFIG.sheetNames.masterClient);
  const sSheet = ss.getSheetByName(CONFIG.sheetNames.masterSet);
  const vSheet = ss.getSheetByName(CONFIG.sheetNames.masterVendor);
  let clients = [], sets = [], vendors = [];
  
  if (mSheet && mSheet.getLastRow() > 1) { 
    clients = [...new Set(mSheet.getRange("A2:A" + mSheet.getLastRow()).getValues().flat().filter(String))]; 
  }
  if (sSheet && sSheet.getLastRow() > 1) { 
    sets = [...new Set(sSheet.getRange("A2:A" + sSheet.getLastRow()).getValues().flat().filter(String))]; 
  }
  if (vSheet && vSheet.getLastRow() > 1) {
    const vData = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 4).getValues(); 
    const map = new Map();
    vData.forEach(r => {
      const name = String(r[1]).trim();
      if (!name) return;
      const display = r[2] ? `${name} ${r[2]}` : name;
      map.set(display, { name, honorific: r[2]||'', displayName: display, account: r[3]||'' });
    });
    vendors = Array.from(map.values());
  }
  
  const result = JSON.stringify({ clients, sets, vendors });
  
  try {
    cache.put("masters_data", result, CACHE_TTL);
  } catch (e) {
    console.warn("Cache put failed (masters_data): " + e.message);
  }
  return result;
}

function apiGetUnifiedProducts() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("products_data");
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = new Map();
  const add = (item, source) => {
    const key = (source + "_" + item.product + "_" + (item.spec||"")).trim();
    if (!item.product || products.has(key)) return;
    item.source = source;
    products.set(key, item);
  };
  const bSheet = ss.getSheetByName(CONFIG.sheetNames.masterBasic);
  if (bSheet && bSheet.getLastRow() > 1) {
    bSheet.getRange(2, 1, bSheet.getLastRow()-1, 4).getValues().forEach(r => {
      if(r[0]) add({ category: "-", product: r[0], spec: r[1], unit: r[2], price: parseCurrency(r[3]) }, "基本");
    });
  }
  const cSheet = ss.getSheetByName(CONFIG.sheetNames.masterClient);
  if (cSheet && cSheet.getLastRow() > 1) {
    cSheet.getRange(2, 1, cSheet.getLastRow()-1, 6).getValues().forEach(r => {
      if(r[2]) add({ category: r[1], product: r[2], spec: r[3], unit: r[4], price: parseCurrency(r[5]) }, "元請:" + r[0]);
    });
  }
  const sSheet = ss.getSheetByName(CONFIG.sheetNames.masterSet);
  if (sSheet && sSheet.getLastRow() > 1) {
    sSheet.getRange(2, 1, sSheet.getLastRow()-1, 8).getValues().forEach(r => {
      if(r[2]) {
        const rawPrice = parseCurrency(r[6]);
        const rawAmount = parseCurrency(r[7]);
        const qty = Number(r[4]) || 0;
        // 単価が空欄で金額だけ記載されている場合、金額÷数量を単価とする
        const price = rawPrice > 0 ? rawPrice : (qty > 0 && rawAmount > 0 ? Math.round(rawAmount / qty) : 0);
        add({ category: r[1], product: r[2], spec: r[3], unit: r[5], price: price }, "セット");
      }
    });
  }
  const lSheet = ss.getSheetByName(CONFIG.sheetNames.list);
  if (lSheet && lSheet.getLastRow() > 1) {
    const lastRow = lSheet.getLastRow();
    const HISTORY_LIMIT = 500;  // 履歴データは直近500件に制限（起動時間短縮）
    const startRow = Math.max(2, lastRow - HISTORY_LIMIT + 1);
    const data = lSheet.getRange(startRow, 1, lastRow, 10).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const r = data[i];
      if (r[4] && r[9]) {
        add({ category: r[3], product: r[4], spec: r[5], unit: r[7], price: parseCurrency(r[9]) }, "履歴");
      }
    }
  }

  const result = JSON.stringify(Array.from(products.values()));
  
  try {
    cache.put("products_data", result, CACHE_TTL);
  } catch (e) {
    console.warn("Cache put failed (products_data): " + e.message);
  }
  return result;
}

function apiSearchSets(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sSheet = ss.getSheetByName(CONFIG.sheetNames.masterSet);
  if (!sSheet) return JSON.stringify([]);
  const normalizedKeyword = toHalfWidth(keyword || "").toLowerCase();
  const keywords = normalizedKeyword.split(/\s+/).filter(k => k);
  const data = sSheet.getDataRange().getDisplayValues().slice(1);
  const setMap = new Map();
  data.forEach(r => {
      const setName = r[0];
      if (!setName) return;
      if (keywords.length === 0 || keywords.every(k => setName.toLowerCase().includes(k))) {
          if (!setMap.has(setName)) setMap.set(setName, { name: setName, firstItem: r[2], totalPrice: 0, count: 0 });
          const current = setMap.get(setName);
          current.count++;
          // 単価が空でも金額があればそれを使って合計を計算
          const rawPrice = parseCurrency(r[6]);
          const rawAmount = parseCurrency(r[7]);
          const qty = parseCurrency(r[4]);
          // 金額が記載されていればそのまま使用、なければ単価×数量
          const lineAmount = rawAmount > 0 ? rawAmount : (rawPrice * qty);
          current.totalPrice += lineAmount;
      }
  });
  const result = Array.from(setMap.values()).filter(s => (s.totalPrice || 0) > 0);
  return JSON.stringify(result);
}

function apiGetSetDetails(setName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.masterSet);
  if (!sheet) return JSON.stringify([]);
  const data = sheet.getDataRange().getValues().slice(1);
  const items = data.filter(r => r[0] === setName).map(r => {
      const rawPrice = parseCurrency(r[6]);
      const rawAmount = parseCurrency(r[7]);
      const qty = Number(r[4]) || 0;
      // 単価が空欄で金額だけ記載されている場合、金額÷数量を単価とする（金額が正のときのみ）
      const price = rawPrice !== 0 ? rawPrice : (qty > 0 && rawAmount > 0 ? Math.round(rawAmount / qty) : 0);
      // 金額が記載されていればそのまま使用（負数含む：値引き等）、なければ単価×数量
      const amount = rawAmount !== 0 ? rawAmount : Math.round(qty * price);
      return {
        category: r[1], product: r[2], spec: r[3], qty: qty, unit: r[5], price: price, amount: amount, remarks: r[8] || ""
      };
    });
  return JSON.stringify(items);
}

function apiListDriveFiles() {
  if (!CONFIG.inputFolder) return JSON.stringify({ error: "未設定" });
  const cache = CacheService.getScriptCache();
  const cacheKey = "drive_files_" + CONFIG.inputFolder.slice(-8);
  const cached = cache.get(cacheKey);
  if (cached) return cached;
  try {
    const folder = DriveApp.getFolderById(CONFIG.inputFolder);
    const files = folder.getFiles(); 
    const result = [];
    while (files.hasNext()) { 
        const f = files.next();
        const m = f.getMimeType();
        if (m.includes("image") || m.includes("pdf") || m.includes("text")) {
            result.push({ id: f.getId(), name: f.getName(), mime: m, updated: formatDate(f.getLastUpdated()) }); 
        }
    }
    const json = JSON.stringify(result.sort((a,b)=>new Date(b.updated)-new Date(a.updated)).slice(0, 30));
    try { cache.put(cacheKey, json, 60); } catch (e) { /* ignore */ }
    return json;
  } catch(e) { return JSON.stringify({ error: e.toString() }); }
}

function apiGetClientHistory(clientName) {
  if (!clientName) return JSON.stringify([]);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.list);
  if (!sheet) return JSON.stringify([]);
  const data = sheet.getDataRange().getDisplayValues();
  const history = [];
  let currentId = "", currentHeader = null, items = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i]; const id = row[0];
    if (id) {
        if (currentId && currentHeader && currentHeader.client === clientName) { history.push({ header: currentHeader, items: items }); }
        currentId = id;
        currentHeader = { id: id, date: row[1], client: row[2], project: row[13], location: row[12], payment: row[15], status: row[17] };
        items = [];
    }
    if (currentId && row[4]) { 
        items.push({ category: row[3], product: row[4], spec: row[5], qty: row[6], unit: row[7], cost: row[8], price: row[9], amount: row[10] });
    }
  }
  if (currentId && currentHeader && currentHeader.client === clientName) { history.push({ header: currentHeader, items: items }); }
  return JSON.stringify(history.reverse());
}

// -----------------------------------------------------------
// プロジェクト・見積関連 API
// -----------------------------------------------------------

function apiGetProjects() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("projects_data");
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // パフォーマンス最適化: 全シート一括取得
  const sheets = {
    list: ss.getSheetByName(CONFIG.sheetNames.list),
    order: ss.getSheetByName(CONFIG.sheetNames.order),
    invoice: ss.getSheetByName(CONFIG.sheetNames.invoice),
    deposits: ss.getSheetByName(CONFIG.sheetNames.deposits)
  };
  if (!sheets.list) { ss.insertSheet(CONFIG.sheetNames.list); return JSON.stringify([]); }
  if (sheets.list.getLastRow() < 2) return JSON.stringify([]);

  const orderSummary = {}; 
  if (sheets.order && sheets.order.getLastRow() > 1) {
    const oData = sheets.order.getDataRange().getDisplayValues();
    if (oData.length > 1) {
      let hIdx = 0;
      for(let i=0; i<Math.min(10, oData.length); i++) { if(oData[i][0] === 'ID') { hIdx = i; break; } }
      const h = oData[hIdx];
      const col = {}; h.forEach((v, i) => col[String(v).trim()] = i);
      const idxEstId = col['関連見積ID']; const idxAmount = col['金額'];
      if (idxEstId !== undefined && idxAmount !== undefined) {
        for (let i = hIdx + 1; i < oData.length; i++) {
          const row = oData[i]; const estId = row[idxEstId];
          if (!estId) continue;
          const amount = parseCurrency(row[idxAmount]);
          if (!orderSummary[estId]) orderSummary[estId] = { totalCost: 0, orderCount: 0 };
          orderSummary[estId].totalCost += amount;
          orderSummary[estId].orderCount += 1;
        }
      }
    }
  }

  const invoiceSummary = {};
  if (sheets.invoice && sheets.invoice.getLastRow() > 1) {
    const iData = sheets.invoice.getDataRange().getDisplayValues();
    for (let i = 1; i < iData.length; i++) {
      const row = iData[i]; const constId = row[4]; const payAmount = parseCurrency(row[10]); 
      if (constId) {
        if (!invoiceSummary[constId]) invoiceSummary[constId] = { totalInvoiced: 0, invoiceCount: 0 };
        invoiceSummary[constId].totalInvoiced += payAmount;
        invoiceSummary[constId].invoiceCount += 1;
      }
    }
  }

  // 入金データ集計 (見積ID単位)
  const depositSummary = {};
  if (sheets.deposits && sheets.deposits.getLastRow() > 1) {
    const dData = sheets.deposits.getDataRange().getDisplayValues();
    for (let i = 1; i < dData.length; i++) {
      const row = dData[i];
      if (!row[0]) continue;
      const estId = String(row[3]).trim(); // 関連見積ID
      if (!estId) continue;
      const status = String(row[11]).trim();
      if (status === '取消') continue;
      const amount = parseCurrency(row[7]);
      if (!depositSummary[estId]) depositSummary[estId] = { totalDeposit: 0, depositCount: 0 };
      depositSummary[estId].totalDeposit += amount;
      depositSummary[estId].depositCount += 1;
    }
  }

  const data = sheets.list.getDataRange().getValues().slice(1);
  const projectMap = {};
  let currentId = "";
  data.forEach(row => {
    const id = String(row[0]); if (id) currentId = id; 
    if (currentId) {
      if (!projectMap[currentId]) { 
        const summary = orderSummary[currentId] || { totalCost: 0, orderCount: 0 };
        const invSummary = invoiceSummary[currentId] || { totalInvoiced: 0, invoiceCount: 0 };
        const depSummary = depositSummary[currentId] || { totalDeposit: 0, depositCount: 0 };
        projectMap[currentId] = { 
          id: currentId, date: formatDate(row[1]), client: row[2], project: row[13], location: row[12], 
          status: row[17] || "未作成", visibility: row[19] || 'public', 
          totalAmount: 0, totalOrderAmount: summary.totalCost, orderCount: summary.orderCount,
          totalInvoicedAmount: invSummary.totalInvoiced, invoiceCount: invSummary.invoiceCount,
          totalDeposit: depSummary.totalDeposit, depositCount: depSummary.depositCount
        }; 
      }
      projectMap[currentId].totalAmount += Number(row[10]) || 0;
    }
  });
  const result = JSON.stringify(Object.values(projectMap).sort((a, b) => new Date(b.date) - new Date(a.date)));
  try { cache.put("projects_data", result, CACHE_TTL_SHORT); } catch (e) { console.warn("Cache put failed (projects_data): " + e.message); }
  return result;
}

function apiGetActiveProjectsList() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("active_projects_data");
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.list);
  if (!sheet) return JSON.stringify([]);
  const data = sheet.getDataRange().getValues().slice(1);
  const projects = [];
  data.forEach(row => {
    if (row[0] && row[17] !== '完了' && row[17] !== '失注') {
      projects.push({ id: row[0], name: `${row[2]} ${row[13]}`, client: row[2], project: row[13] });
    }
  });
  const result = JSON.stringify(projects.reverse());
  try { cache.put("active_projects_data", result, CACHE_TTL_SHORT); } catch (e) { console.warn("Cache put failed (active_projects_data): " + e.message); }
  return result;
}

function apiGetDrafts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.list);
  if (!sheet) return JSON.stringify([]);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return JSON.stringify([]);
  const draftsMap = new Map();
  let currentId = null;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[0] || "").trim();
    if (id) {
      currentId = id;
      if (!draftsMap.has(id)) {
        draftsMap.set(id, { id: id, date: formatDate(row[1]), timestamp: new Date(row[1] || 0).getTime(), client: row[2], project: row[13], status: row[17], totalAmount: 0 });
      }
    }
    if (currentId && draftsMap.has(currentId) && String(row[4])) {
      const amount = Number(row[10]) || 0;
      draftsMap.get(currentId).totalAmount += amount;
    }
  }
  const list = Array.from(draftsMap.values());
  list.sort((a, b) => b.timestamp - a.timestamp);
  return JSON.stringify(list);
}

function _getEstimateData(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.list);
  const orderSheet = ss.getSheetByName(CONFIG.sheetNames.order);
  if (!sheet) return null;

  const orderAgg = {}; 
  if (orderSheet && orderSheet.getLastRow() > 1) {
    const oData = orderSheet.getDataRange().getDisplayValues();
    let hIdx = 0;
    for(let i=0; i<Math.min(10, oData.length); i++) { if(oData[i][0] === 'ID') { hIdx = i; break; } }
    const h = oData[hIdx]; const col = {}; h.forEach((v, i) => col[String(v).trim()] = i);
    const idxEstId = col['関連見積ID']; const idxProd = col['品名']; const idxSpec = col['仕様'];
    const idxQty = col['数量']; const idxAmt = col['金額']; const idxVendor = col['発注先'];

    if (idxEstId !== undefined) {
      for(let i = hIdx+1; i < oData.length; i++) {
        const r = oData[i];
        if (r[idxEstId] === id) {
          const key = `${r[idxProd]}_${r[idxSpec]}`;
          if (!orderAgg[key]) orderAgg[key] = { qty: 0, vendors: [], amount: 0 };
          orderAgg[key].qty += parseCurrency(r[idxQty]);
          orderAgg[key].amount += parseCurrency(r[idxAmt]);
          let vName = String(r[idxVendor]).replace(/(株式会社|有限会社|合同会社)/g, '').trim();
          if (vName && !orderAgg[key].vendors.includes(vName)) orderAgg[key].vendors.push(vName);
        }
      }
    }
  }

  const data = sheet.getDataRange().getValues().slice(1);
  let header = null;
  const items = [];
  let isTarget = false;
  data.forEach(row => {
    const rowId = String(row[0]);
    if (rowId !== "") {
      if (rowId === id) {
        isTarget = true;
        header = { id: rowId, date: formatDate(row[1]), client: row[2], location: row[12], project: row[13], period: row[14], payment: row[15], expiry: row[16], status: row[17], remarks: row[11], visibility: row[19] || 'public' };
      } else { isTarget = false; }
    }
    if (isTarget && String(row[4])) {
      const key = `${row[4]}_${row[5]}`;
      const ordered = orderAgg[key] || { qty: 0, vendors: [], amount: 0 };
      items.push({
        category: row[3], product: row[4], spec: row[5], qty: Number(row[6]), unit: row[7],
        cost: Number(row[8]) || 0, price: Number(row[9]), amount: Number(row[10]),
        remarks: row[11], vendor: row[18] || "",
        orderedQty: ordered.qty, orderedAmount: ordered.amount, orderedVendors: ordered.vendors.join(", ")
      });
    }
  });
  if (header) {
    const totalAmount = items.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    return { header, items, totalAmount };
  }
  return null;
}

function apiGetEstimateDetails(id) { return JSON.stringify(_getEstimateData(id)); }

function apiSaveUnifiedData(jsonData) {
  const data = JSON.parse(jsonData); 
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "Busy" });

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const estimateData = data.estimate || data;
    if (!estimateData || !estimateData.header) {
        return JSON.stringify({ success: false, message: "Invalid Data Structure" });
    }

    let estSheet = ss.getSheetByName(CONFIG.sheetNames.list);
    if (!estSheet) estSheet = ss.insertSheet(CONFIG.sheetNames.list);
    
    let saveId = estimateData.header.id;
    if (!saveId) saveId = getNextSequenceId('estimate');
    
    // Performance Tuning: Use Optimized Delete
    deleteRowsById(estSheet, saveId);
    
    const now = new Date();
    const saveTimestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
    
    const estItems = (estimateData.items && estimateData.items.length > 0) ? estimateData.items : [{category:'', product:'', spec:'', qty:0, unit:'', cost:0, price:0, amount:0, remarks:'', vendor:''}];
    
    const estValues = estItems.map((item, i) => {
      const isFirst = (i === 0);
      return [
        isFirst ? saveId : "", 
        isFirst ? saveTimestamp : "", 
        isFirst ? estimateData.header.client : "", 
        item.category, item.product, item.spec, item.qty, item.unit, item.cost, item.price, item.amount, item.remarks, 
        isFirst ? estimateData.header.location : "", 
        isFirst ? estimateData.header.project : "", 
        isFirst ? estimateData.header.period : "", 
        isFirst ? estimateData.header.payment : "", 
        isFirst ? estimateData.header.expiry : "", 
        isFirst ? (estimateData.header.status || "見積提出") : "", 
        item.vendor, 
        isFirst ? (estimateData.header.visibility || 'public') : "" 
      ];
    });
    
    estSheet.getRange(estSheet.getLastRow() + 1, 1, estValues.length, 20).setValues(estValues.map(r => { while(r.length < 20) r.push(""); return r; }));

    let orderSheet = ss.getSheetByName(CONFIG.sheetNames.order);
    if (!orderSheet) { orderSheet = ss.insertSheet(CONFIG.sheetNames.order); checkAndFixOrderHeader(orderSheet); }
    
    const oData = orderSheet.getDataRange().getValues();
    const oHeaders = oData[0] || [];
    const relEstIdColIdx = oHeaders.indexOf('関連見積ID');
    const orderRelEstCol = relEstIdColIdx !== -1 ? relEstIdColIdx : 3;
    const orderRowsToDelete = [];
    if (oData.length > 1) {
        for (let i = 1; i < oData.length; i++) {
            if (String(oData[i][orderRelEstCol]) === saveId) { orderRowsToDelete.push(i + 1); }
        }
        // Performance Tuning: Use Optimized Delete for Orders too
        if (orderRowsToDelete.length > 0) {
            deleteRowsOptimized_(orderSheet, orderRowsToDelete);
        }
    }

    const email = Session.getActiveUser().getEmail();
    const orderValues = [];
    const vendorGroups = {};
    if (estimateData.items) {
      estimateData.items.forEach((item) => {
          if (item.vendor && (Number(item.cost) > 0 || Number(item.qty) > 0)) {
              const v = String(item.vendor).trim();
              if (!vendorGroups[v]) vendorGroups[v] = [];
              vendorGroups[v].push(item);
          }
      });
    }
    Object.keys(vendorGroups).forEach((vendor) => {
      const items = vendorGroups[vendor];
      const orderId = getNextSequenceId('order');
      items.forEach((item, idx) => {
        orderValues.push([
            idx === 0 ? orderId : "",
            saveTimestamp, item.vendor, saveId,
            item.category, item.product, item.spec, item.qty, item.unit, item.cost, Math.round((Number(item.qty) || 0) * (Number(item.cost) || 0)),
            estimateData.header.location, "発注書作成", "", email, "public"
        ]);
      });
    });

    if (orderValues.length > 0) {
        const startRow = orderSheet.getLastRow() + 1;
        orderSheet.getRange(startRow, 1, orderValues.length, 16).setValues(orderValues.map(r => { while(r.length < 16) r.push(""); return r; }));
    }

    invalidateDataCache_();
    return JSON.stringify({ success: true, id: saveId });

  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

// 発注データ単体保存用API (本実装)
function apiSaveOrderOnly(jsonData) {
  const data = JSON.parse(jsonData); // { header:..., items:... }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "Busy" });

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let orderSheet = ss.getSheetByName(CONFIG.sheetNames.order);
    if (!orderSheet) { orderSheet = ss.insertSheet(CONFIG.sheetNames.order); checkAndFixOrderHeader(orderSheet); }
    
    const saveTimestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
    const email = Session.getActiveUser().getEmail();
    const relEstId = data.header.relEstId || "";
    const orderId = data.header.id || getNextSequenceId('order');

    if (data.header.id) {
      deleteRowsById(orderSheet, data.header.id);
    }

    const orderValues = [];

    data.items.forEach((item, idx) => {
        orderValues.push([
            idx === 0 ? orderId : "",
            saveTimestamp,
            data.header.vendor || item.vendor,
            relEstId,
            item.category || "",
            item.product,
            item.spec || "",
            item.qty,
            item.unit || "",
            item.cost,
            Math.round((Number(item.qty) || 0) * (Number(item.cost) || 0)),
            data.header.location || "",
            "発注書作成",
            "",
            email,
            "public"
        ]);
    });

    if (orderValues.length > 0) {
        const startRow = orderSheet.getLastRow() + 1;
        const padded = orderValues.map(r => { while (r.length < 16) r.push(""); return r; });
        orderSheet.getRange(startRow, 1, orderValues.length, 16).setValues(padded);
    }
    
    invalidateDataCache_();
    return JSON.stringify({ success: true, id: orderId, count: orderValues.length });

  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

function apiSaveAndCreateEstimatePdf(jsonData) {
  const data = JSON.parse(jsonData); 
  const savePayload = { estimate: data };
  const saveResJson = apiSaveUnifiedData(JSON.stringify(savePayload));
  const saveRes = JSON.parse(saveResJson);
  
  if (!saveRes.success) return saveResJson;
  const saveId = saveRes.id;
  data.header.id = saveId;

  try {
    const now = new Date();
    data.totalAmount = data.items.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
    data.header.date = getJapaneseDateStr(now);
    data.pages = paginateItems(data.items, 20, 35);

    let template;
    try { template = HtmlService.createTemplateFromFile('quote_template'); } 
    catch(e) { return JSON.stringify({ success: false, message: "見積書テンプレート(quote_template.html)が見つかりません。作成してください。" }); }
    
    template.data = data; 
    const html = template.evaluate().getContent();
    const cleanClient = (data.header.client || "").replace(/[\r\n\t\\/:*?"<>|]/g, '').trim();
    const fileName = `御見積書_${cleanClient}_${data.header.project || saveId}.pdf`;
    const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF).setName(fileName);
    
    const folder = getSaveFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return JSON.stringify({ success: true, id: saveId, url: file.getUrl() });

  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  }
}

function apiIssueBillFromId(id) {
  const estimateData = _getEstimateData(id);
  if (!estimateData) return JSON.stringify({ success: false, message: "データが見つかりません" });
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.list);
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === id) {
      sheet.getRange(i+1, 18).setValue("請求済"); 
      invalidateDataCache_();
      break;
    }
  }
  
  const now = new Date();
  estimateData.totalAmount = estimateData.items.reduce((s, i) => s + (Number(i.amount) || 0), 0);
  estimateData.header.date = getJapaneseDateStr(now);
  estimateData.pages = paginateItems(estimateData.items, 20, 35);
  
  let template;
  try { template = HtmlService.createTemplateFromFile('bill_template'); } 
  catch(e) { return JSON.stringify({ success: false, message: "請求書テンプレート(bill_template.html)が見つかりません。" }); }
  template.data = estimateData; 
  const blob = Utilities.newBlob(template.evaluate().getContent(), MimeType.HTML).getAs(MimeType.PDF).setName(`御請求書_${estimateData.header.client}_${estimateData.header.project}.pdf`);
  
  const folder = getSaveFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return JSON.stringify({ success: true, url: file.getUrl() });
}

// -----------------------------------------------------------
// 発注関連 API
// -----------------------------------------------------------

function apiGetOrders() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("orders_data");
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetNames.order);
  if (!sheet) { sheet = ss.insertSheet(CONFIG.sheetNames.order); checkAndFixOrderHeader(sheet); return JSON.stringify([]); }
  const allData = sheet.getDataRange().getDisplayValues();
  if (allData.length < 2) return JSON.stringify([]);

  let headerRowIndex = 0;
  for (let i = 0; i < Math.min(10, allData.length); i++) {
    if (allData[i][0] === "ID") { headerRowIndex = i; break; }
  }
  const headers = allData[headerRowIndex];
  const col = {}; headers.forEach((h, i) => { col[String(h).trim()] = i; });

  const IDX = {
    id: col["ID"] !== undefined ? col["ID"] : 0, date: col["日付"] !== undefined ? col["日付"] : 1, vendor: col["発注先"] !== undefined ? col["発注先"] : 2, relEstId: col["関連見積ID"] !== undefined ? col["関連見積ID"] : 3,
    product: col["品名"] !== undefined ? col["品名"] : 5,
    amount: col["金額"] !== undefined ? col["金額"] : 10, location: col["納品場所"] !== undefined ? col["納品場所"] : 11, status: col["状態"] !== undefined ? col["状態"] : 12, remarks: col["備考"] !== undefined ? col["備考"] : 13, creator: col["作成者"] !== undefined ? col["作成者"] : 14, visibility: col["公開範囲"] !== undefined ? col["公開範囲"] : 15
  };

  // 出金データ集計 (発注ID単位)
  const paymentSummary = {};
  const paySheet = ss.getSheetByName(CONFIG.sheetNames.payments);
  if (paySheet && paySheet.getLastRow() > 1) {
    const pData = paySheet.getDataRange().getDisplayValues();
    for (let i = 1; i < pData.length; i++) {
      const row = pData[i];
      if (!row[0]) continue;
      const orderId = String(row[3]).trim(); // 関連発注ID
      if (!orderId) continue;
      const status = String(row[12]).trim();
      if (status === '取消') continue;
      const amount = parseCurrency(row[8]);
      if (!paymentSummary[orderId]) paymentSummary[orderId] = { totalPaid: 0, paymentCount: 0 };
      paymentSummary[orderId].totalPaid += amount;
      paymentSummary[orderId].paymentCount += 1;
    }
  }

  // PDF存在チェック用: ドライブフォルダ内のPDFファイル名を収集
  let pdfFileNames = [];
  try {
    const folder = getSaveFolder();
    const pdfFiles = folder.getFilesByType(MimeType.PDF);
    while (pdfFiles.hasNext()) {
      pdfFileNames.push(pdfFiles.next().getName());
    }
  } catch(e) { /* ignore */ }

  const orderMap = new Map();
  let currentId = ""; 
  for (let i = headerRowIndex + 1; i < allData.length; i++) {
    const row = allData[i]; const idCell = row[IDX.id]; 
    if (idCell && idCell !== "ID") { currentId = idCell; }
    if (!currentId) continue;

    if (!orderMap.has(currentId)) {
      const paySummary = paymentSummary[currentId] || { totalPaid: 0, paymentCount: 0 };
      orderMap.set(currentId, {
        id: currentId, date: row[IDX.date], vendor: row[IDX.vendor], relEstId: row[IDX.relEstId], location: row[IDX.location],
        status: row[IDX.status], remarks: row[IDX.remarks], creator: row[IDX.creator] || '', visibility: row[IDX.visibility] || 'public', totalAmount: 0,
        totalPaid: paySummary.totalPaid, paymentCount: paySummary.paymentCount,
        hasPdf: false, project: ''
      });
    }
    const amount = parseCurrency(row[IDX.amount]);
    const currentData = orderMap.get(currentId);
    if (currentData) { currentData.totalAmount += amount; }
  }

  // PDF存在チェック & プロジェクト名取得
  const list = Array.from(orderMap.values());
  list.forEach(order => {
    // PDF存在チェック: 発注書_業者名_ でファイル名マッチ
    const cleanVendor = (order.vendor || '').replace(/[\r\n\t\\/:*?"<>|]/g, '').trim();
    order.hasPdf = pdfFileNames.some(fn => fn.includes('発注書_' + cleanVendor));
    // 関連見積IDからプロジェクト名を推定 (見積リストの工事名)
    if (order.relEstId) {
      const est = _getEstimateHeaderOnly(order.relEstId);
      if (est) order.project = est.project || '';
    }
  });

  list.sort((a, b) => new Date(b.date) - new Date(a.date));
  const result = JSON.stringify(list);
  try { cache.put("orders_data", result, CACHE_TTL_ORDERS); } catch (e) { console.warn("Cache put failed (orders_data): " + e.message); }
  return result;
}

// 軽量ヘッダー取得 (apiGetOrdersから利用、明細不要)
function _getEstimateHeaderOnly(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.list);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      return { project: data[i][13], client: data[i][2], location: data[i][12] };
    }
  }
  return null;
}

function apiCreateOrderPdf(jsonData, targetVendor) {
  const data = JSON.parse(jsonData);
  
  if (targetVendor) {
    data.header.vendor = targetVendor;
    const filtered = data.items.filter(item => item.vendor === targetVendor);
    const hasAnyVendor = data.items.some(item => (item.vendor || '').trim() !== '');
    
    if (filtered.length > 0) {
      data.items = filtered;
    } else if (!hasAnyVendor) {
      // 明細にvendorが無い場合（単独発注画面等）は全件を対象にvendorを付与
      data.items = data.items.map(item => Object.assign({}, item, { vendor: targetVendor }));
    } else {
      data.items = filtered; // 該当発注先の明細が無い
    }
  }

  if (!data.items || data.items.length === 0) {
    return JSON.stringify({ success: false, message: "指定された発注先の明細がありません。" });
  }

  const now = new Date();
  data.header.honorific = " 御中"; 
  data.header.date = getJapaneseDateStr(now);
  data.totalAmount = data.items.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
  data.pages = paginateItems(data.items, 22, 35);

  // 関連見積IDがある場合、見積データから工事名・工期・決済条件・有効期限を取得
  if (data.header.relEstId) {
    const estimateData = _getEstimateData(data.header.relEstId);
    if (estimateData && estimateData.header) {
      data.header.project = estimateData.header.project || data.header.project || "";
      data.header.location = data.header.location || estimateData.header.location || "";
      data.header.period = estimateData.header.period || data.header.period || "";
      data.header.payment = estimateData.header.payment || data.header.payment || "";
      data.header.expiry = estimateData.header.expiry || data.header.expiry || "";
    }
  }
  if (!data.header.project) data.header.project = "";
  if (!data.header.period) data.header.period = "";
  if (!data.header.payment) data.header.payment = "";
  if (!data.header.expiry) data.header.expiry = "";

  let template; 
  try { template = HtmlService.createTemplateFromFile('order_template'); } 
  catch(e) { return JSON.stringify({ success: false, message: "発注書テンプレート(order_template.html)が見つかりません。作成してください。" }); }

  template.data = data;
  const cleanVendor = (targetVendor || data.header.vendor || "発注先不明").replace(/[\r\n\t\\/:*?"<>|]/g, '').trim();
  const fileName = `発注書_${cleanVendor}_${data.header.project || '案件'}.pdf`;
  const blob = Utilities.newBlob(template.evaluate().getContent(), MimeType.HTML).getAs(MimeType.PDF).setName(fileName);
  
  const folder = getSaveFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return JSON.stringify({ success: true, url: file.getUrl() });
}

// --- 同じ見積+発注先の既存発注IDを取得 ---
function apiFindOrderByEstimateAndVendor(relEstId, vendor) {
  if (!relEstId || !vendor) return JSON.stringify(null);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.order);
  if (!sheet) return JSON.stringify(null);
  const data = sheet.getDataRange().getDisplayValues();
  let hIdx = 0;
  for (let i = 0; i < Math.min(10, data.length); i++) { if (data[i][0] === 'ID') { hIdx = i; break; } }
  const headers = data[hIdx];
  const col = {}; headers.forEach((h, i) => { col[String(h).trim()] = i; });
  const idxRelEstId = col['関連見積ID'] !== undefined ? col['関連見積ID'] : 3;
  const idxVendor = col['発注先'] !== undefined ? col['発注先'] : 2;
  let currentId = '';
  for (let i = hIdx + 1; i < data.length; i++) {
    const row = data[i];
    const idCell = row[0];
    if (idCell && idCell !== 'ID') currentId = idCell;
    if (!currentId) continue;
    const rRel = String(row[idxRelEstId] || '').trim();
    const rVendor = String(row[idxVendor] || '').trim();
    if (rRel === String(relEstId).trim() && rVendor === String(vendor).trim()) {
      return JSON.stringify(currentId);
    }
  }
  return JSON.stringify(null);
}

// --- 発注データ詳細取得（履歴からの編集用） ---
function apiGetOrderDetails(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.order);
  if (!sheet) return JSON.stringify({ error: "発注データが見つかりません" });
  const data = sheet.getDataRange().getDisplayValues();
  let hIdx = 0;
  for (let i = 0; i < Math.min(10, data.length); i++) { if (data[i][0] === 'ID') { hIdx = i; break; } }
  const headers = data[hIdx];
  const col = {}; headers.forEach((h, i) => { col[String(h).trim()] = i; });
  const items = [];
  let header = null;
  let currentId = '';
  for (let i = hIdx + 1; i < data.length; i++) {
    const row = data[i];
    const idCell = row[col['ID']];
    if (idCell && idCell !== 'ID') { currentId = idCell; }
    if (!currentId || currentId !== orderId) continue;
    if (!header) {
      header = {
        id: orderId,
        vendor: row[col['発注先']],
        date: row[col['日付']],
        relEstId: row[col['関連見積ID']] || '',
        location: row[col['納品場所']] || '',
        remarks: row[col['備考']] || ''
      };
    }
    items.push({
      category: row[col['工種']] || '',
      product: row[col['品名']] || '',
      spec: row[col['仕様']] || '',
      qty: parseCurrency(row[col['数量']]) || 0,
      unit: row[col['単位']] || '',
      cost: parseCurrency(row[col['単価']]) || 0,
      amount: parseCurrency(row[col['金額']]) || 0
    });
  }
  if (!header) return JSON.stringify({ error: "指定された発注が見つかりません" });
  // 関連見積IDがある場合、見積から工事名・工期・決済条件・有効期限を取得
  if (header.relEstId) {
    const est = _getEstimateData(header.relEstId);
    if (est && est.header) {
      header.project = est.header.project || "";
      header.period = est.header.period || "";
      header.payment = est.header.payment || "";
      header.expiry = est.header.expiry || "";
      if (!header.location && est.header.location) header.location = est.header.location;
    }
  }
  return JSON.stringify({ header, items });
}

// --- Phase 4 追加機能: 保存済み発注書のPDF再発行 ---
function apiReprintOrderPdf(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.order);
  if (!sheet) return JSON.stringify({ success: false, message: "発注データが見つかりません" });
  
  const data = sheet.getDataRange().getDisplayValues();
  // ヘッダー検索
  let hIdx = 0;
  for(let i=0; i<Math.min(10, data.length); i++) { if(data[i][0] === 'ID') { hIdx = i; break; } }
  const headers = data[hIdx];
  const col = {}; headers.forEach((h, i) => { col[String(h).trim()] = i; });
  
  // データ抽出（apiGetOrderDetailsと同様: 同一orderIdの行はIDが空でも続く）
  const items = [];
  let header = null;
  let currentId = '';
  
  for(let i=hIdx+1; i<data.length; i++) {
    const row = data[i];
    const idCell = row[col['ID']];
    if (idCell && idCell !== 'ID') { currentId = idCell; }
    if (!currentId || currentId !== orderId) continue;
    
    if(!header) {
      header = {
        id: orderId,
        date: row[col['日付']],
        vendor: row[col['発注先']],
        relEstId: row[col['関連見積ID']],
        location: row[col['納品場所']],
        remarks: row[col['備考']],
        honorific: " 御中"
      };
    }
    items.push({
      product: row[col['品名']] || '',
      spec: row[col['仕様']] || '',
      qty: parseCurrency(row[col['数量']]) || 0,
      unit: row[col['単位']] || '',
      cost: parseCurrency(row[col['単価']]) || 0,
      amount: parseCurrency(row[col['金額']]) || 0
    });
  }
  
  if(!header) return JSON.stringify({ success: false, message: "指定された発注書が見つかりません" });
  
  // 関連見積IDがある場合、見積から工事名・工期・決済条件・有効期限を取得
  if (header.relEstId) {
    const est = _getEstimateData(header.relEstId);
    if (est && est.header) {
      header.project = est.header.project || header.project || "";
      header.period = est.header.period || "";
      header.payment = est.header.payment || "";
      header.expiry = est.header.expiry || "";
      if (!header.location && est.header.location) header.location = est.header.location;
    }
  }
  if (!header.project) header.project = "";
  if (!header.period) header.period = "";
  if (!header.payment) header.payment = "";
  if (!header.expiry) header.expiry = "";
  
  const totalAmount = items.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
  // PDF生成（totalAmountを追加してテンプレートで合計表示）
  const pdfData = { header: header, items: items, totalAmount: totalAmount, pages: paginateItems(items, 22, 35) };
  
  try {
    let template;
    try { template = HtmlService.createTemplateFromFile('order_template'); }
    catch(e) { return JSON.stringify({ success: false, message: "order_template.html が見つかりません" }); }
    
    template.data = pdfData;
    const blob = Utilities.newBlob(template.evaluate().getContent(), MimeType.HTML).getAs(MimeType.PDF).setName(`発注書_${header.vendor}_${header.id}.pdf`);
    
    const folder = getSaveFolder();
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return JSON.stringify({ success: true, url: file.getUrl() });
  } catch(e) {
    return JSON.stringify({ success: false, message: e.toString() });
  }
}

// -----------------------------------------------------------
// 請求書受取・AI解析 API
// -----------------------------------------------------------

function apiListInvoiceDriveFiles() {
  if (!CONFIG.invoiceInputFolder) return JSON.stringify({ error: "請求書受取フォルダIDが未設定です" });
  const cache = CacheService.getScriptCache();
  const cacheKey = "invoice_files_" + String(CONFIG.invoiceInputFolder).slice(-8);
  const cached = cache.get(cacheKey);
  if (cached) return cached;
  try {
    const folder = DriveApp.getFolderById(CONFIG.invoiceInputFolder);
    const files = folder.getFiles(); 
    const result = [];
    while (files.hasNext()) { 
        const f = files.next();
        const m = f.getMimeType();
        if (m.includes("image") || m.includes("pdf") || m.includes("text")) {
            result.push({ id: f.getId(), name: f.getName(), mime: m, updated: formatDate(f.getLastUpdated()) }); 
        }
    }
    const json = JSON.stringify(result.sort((a,b)=>new Date(b.updated)-new Date(a.updated)).slice(0, 30));
    try { cache.put(cacheKey, json, 60); } catch (e) { /* ignore */ }
    return json;
  } catch(e) { return JSON.stringify({ error: e.toString() }); }
}

function apiParseInvoiceFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mime = file.getMimeType();
    const name = file.getName();
    if (mime.includes("text") || name.endsWith(".txt")) { return JSON.stringify(_parseTextInvoice(file)); } 
    else if (mime.includes("image") || mime.includes("pdf")) { return JSON.stringify(_parseInvoiceImageWithGemini(file)); }
    return JSON.stringify({ error: "Unsupported file type" });
  } catch (e) { return JSON.stringify({ error: e.toString() }); }
}

function _parseTextInvoice(file) {
  let content = "";
  try {
    content = file.getBlob().getDataAsString();
    if (!content.match(/工事|現場|請求|金額|日付|業者/)) { content = file.getBlob().getDataAsString('Shift_JIS'); }
  } catch(e) {}
  const lines = content.split(/\r\n|\n/);
  const result = { constructionId: "", project: "", supplier: "", amount: 0, content: "", date: "" };
  const keyMap = [
    { key: "constructionId", keywords: ["工事番号", "工事ID", "No"] },
    { key: "project", keywords: ["現場名", "工事名", "案件名", "件名"] },
    { key: "supplier", keywords: ["請求業者", "業者名", "請求元", "会社名"] },
    { key: "amount", keywords: ["金額", "請求金額", "合計", "税込金額"] },
    { key: "content", keywords: ["内容", "但し書き", "品名", "詳細"] },
    { key: "date", keywords: ["日付", "請求日", "発行日"] }
  ];
  lines.forEach(line => {
    const l = line.trim(); if (!l) return;
    keyMap.forEach(map => {
      map.keywords.forEach(keyword => {
        let value = "";
        const regexBracket = new RegExp(`^【\\s*${keyword}\\s*】\\s*(.*)$`);
        const matchBracket = l.match(regexBracket);
        if (matchBracket) value = matchBracket[1].trim();
        if (!value) {
           const regexColon = new RegExp(`^${keyword}\\s*[:：]\\s*(.*)$`);
           const matchColon = l.match(regexColon);
           if (matchColon) value = matchColon[1].trim();
        }
        if (value) {
          if (map.key === "amount") { result[map.key] = parseCurrency(value); } else { result[map.key] = value; }
        }
      });
    });
  });
  return result;
}

function _parseInvoiceImageWithGemini(file) {
  if (!CONFIG.API_KEY) return { error: "APIキーなし" };
  const projectsJson = apiGetActiveProjectsList();
  const projects = JSON.parse(projectsJson).map(p => `${p.id}: ${p.name}`).join("\n");
  const mime = file.getMimeType();
  const base64 = Utilities.base64Encode(file.getBlob().getBytes());
  const prompt = `あなたは建築積算のプロです。画像から情報を抽出してください。\n【重要】以下のリストを参照し、最も関連性が高い「工事番号(constructionId)」を推測してください。\nリスト: ${projects}\n抽出項目: constructionId, supplier, date(yyyy/MM/dd), amount(税込), content, registrationNumber(Tから始まる13桁の番号)`;
  const parts = [{ text: prompt }, { inline_data: { mime_type: mime, data: base64 } }];
  const responseSchema = {
    "type": "OBJECT",
    "properties": {
      "constructionId": { "type": "STRING" }, "supplier": { "type": "STRING" }, "date": { "type": "STRING" },
      "amount": { "type": "NUMBER" }, "content": { "type": "STRING" }, "registrationNumber": { "type": "STRING", "description": "T+13 digits" }
    }
  };
  const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${CONFIG.API_KEY}`, {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({ contents: [{ parts }], generationConfig: { response_mime_type: "application/json", response_schema: responseSchema } }),
    muteHttpExceptions: true
  });
  return JSON.parse(JSON.parse(res.getContentText()).candidates[0].content.parts[0].text);
}

function apiSaveInvoice(jsonData) {
  const data = JSON.parse(jsonData);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
  if (!sheet) { sheet = ss.insertSheet(CONFIG.sheetNames.invoice); checkAndFixInvoiceHeader(sheet); }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "Busy" });
  try {
    let id = data.id; 
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
    const payment = (Number(data.amount) || 0) - (Number(data.offset) || 0);
    let rowIndex = -1;
    if (id) {
        const sheetData = sheet.getDataRange().getValues();
        for (let i = 1; i < sheetData.length; i++) {
            if (String(sheetData[i][0]) === String(id)) { rowIndex = i + 1; break; }
        }
    } else {
        id = "INV-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMddHHmmss");
    }
    const rowValues = [ 
        id, data.status || "未確認", now, data.fileId || "", data.constructionId || "", 
        data.project || "", data.supplier || "", data.date || "", data.amount || 0, 
        data.offset || 0, payment, data.content || "", data.remarks || "", data.registrationNumber || ""
    ];
    if (rowIndex > 0) {
        const currentStatus = sheet.getRange(rowIndex, 2).getValue();
        rowValues[1] = currentStatus; 
        sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
        sheet.appendRow(rowValues);
    }
    invalidateDataCache_();
    return JSON.stringify({ success: true });
  } catch(e) { return JSON.stringify({ success: false, message: e.toString() }); } finally { lock.releaseLock(); }
}

function apiGetInvoices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
  if (!sheet) { sheet = ss.insertSheet(CONFIG.sheetNames.invoice); checkAndFixInvoiceHeader(sheet); return JSON.stringify([]); }
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return JSON.stringify([]);
  const invoices = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      invoices.push({
        id: row[0], status: row[1], registeredAt: row[2], fileId: row[3], constructionId: row[4], 
        project: row[5], supplier: row[6], date: row[7],
        amount: parseCurrency(row[8]), offset: parseCurrency(row[9]), payment: parseCurrency(row[10]), 
        content: row[11], remarks: row[12], registrationNumber: row[13] || ""
      });
    }
  }
  return JSON.stringify(invoices.reverse());
}

function apiUpdateInvoiceStatus(id, newStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
  if (!sheet) return JSON.stringify({ success: false });
  const data = sheet.getDataRange().getValues();
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) { sheet.getRange(i + 1, 2).setValue(newStatus); found = true; invalidateDataCache_(); break; }
  }
  return JSON.stringify({ success: found });
}

function apiGetOrderBalance(constructionId, supplierName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oSheet = ss.getSheetByName(CONFIG.sheetNames.order);
  const iSheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
  if (!oSheet) return JSON.stringify({ error: "発注シートがありません" });
  if (!supplierName) return JSON.stringify({ totalOrder: 0, totalPaid: 0, balance: 0 });
  const normSupplier = supplierName.replace(/[\s\u3000]/g, "");
  const oData = oSheet.getDataRange().getDisplayValues();
  let totalOrder = 0;
  for (let i = 1; i < oData.length; i++) {
    const row = oData[i];
    const estId = row[3];
    const vendor = row[2].replace(/[\s\u3000]/g, "");
    if (estId && (estId === constructionId || estId.startsWith(constructionId))) {
      if (vendor.includes(normSupplier) || normSupplier.includes(vendor)) totalOrder += parseCurrency(row[10]);
    }
  }
  let totalPaid = 0;
  if (iSheet && iSheet.getLastRow() > 1) {
    const iData = iSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < iData.length; i++) {
      const row = iData[i];
      const invEstId = row[4];
      const invVendor = row[6].replace(/[\s\u3000]/g, "");
      if (invEstId && (invEstId === constructionId || invEstId.startsWith(constructionId))) {
        if (invVendor.includes(normSupplier) || normSupplier.includes(invVendor)) totalPaid += parseCurrency(row[10]); 
      }
    }
  }
  return JSON.stringify({ totalOrder: totalOrder, totalPaid: totalPaid, balance: totalOrder - totalPaid });
}

// -----------------------------------------------------------
// 入出金管理 API
// -----------------------------------------------------------

function getNextDepositId_() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
  const props = PropertiesService.getScriptProperties();
  const key = "SEQ_DEPOSIT_" + dateStr;
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = lock.tryLock(5000);
    if (lockAcquired) {
      let current = Number(props.getProperty(key)) || 0;
      current++;
      props.setProperty(key, String(current));
      return "DEP-" + dateStr + "-" + String(current).padStart(5, "0");
    }
    throw new Error("ID採番タイムアウト");
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function getNextPaymentId_() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd");
  const props = PropertiesService.getScriptProperties();
  const key = "SEQ_PAYMENT_" + dateStr;
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = lock.tryLock(5000);
    if (lockAcquired) {
      let current = Number(props.getProperty(key)) || 0;
      current++;
      props.setProperty(key, String(current));
      return "PAY-" + dateStr + "-" + String(current).padStart(5, "0");
    }
    throw new Error("ID採番タイムアウト");
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

function apiGetDeposits() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("deposits_data");
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetNames.deposits);
  if (!sheet) { sheet = ss.insertSheet(CONFIG.sheetNames.deposits); checkAndFixDepositsHeader(sheet); return JSON.stringify([]); }
  if (sheet.getLastRow() < 2) return JSON.stringify([]);

  const data = sheet.getDataRange().getDisplayValues();
  const deposits = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    deposits.push({
      id: row[0], registeredAt: row[1], date: row[2], estimateId: row[3], client: row[4], project: row[5],
      type: row[6], amount: parseCurrency(row[7]), fee: parseCurrency(row[8]), offset: parseCurrency(row[9]),
      remarks: row[10], status: row[11], registrant: row[12], visibility: row[13] || "public"
    });
  }
  deposits.sort((a, b) => new Date(b.date) - new Date(a.date));
  const result = JSON.stringify(deposits);
  try { cache.put("deposits_data", result, CACHE_TTL_ORDERS); } catch (e) { /* ignore */ }
  return result;
}

function apiGetPayments() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("payments_data");
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetNames.payments);
  if (!sheet) { sheet = ss.insertSheet(CONFIG.sheetNames.payments); checkAndFixPaymentsHeader(sheet); return JSON.stringify([]); }
  if (sheet.getLastRow() < 2) return JSON.stringify([]);

  const data = sheet.getDataRange().getDisplayValues();
  const payments = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    payments.push({
      id: row[0], registeredAt: row[1], date: row[2], orderId: row[3], invoiceId: row[4], supplier: row[5], project: row[6],
      type: row[7], amount: parseCurrency(row[8]), fee: parseCurrency(row[9]), offset: parseCurrency(row[10]),
      remarks: row[11], status: row[12], registrant: row[13], visibility: row[14] || "public"
    });
  }
  payments.sort((a, b) => new Date(b.date) - new Date(a.date));
  const result = JSON.stringify(payments);
  try { cache.put("payments_data", result, CACHE_TTL_ORDERS); } catch (e) { /* ignore */ }
  return result;
}

function apiSaveDeposit(jsonData) {
  const data = JSON.parse(jsonData);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "Busy" });

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.sheetNames.deposits);
    if (!sheet) { sheet = ss.insertSheet(CONFIG.sheetNames.deposits); checkAndFixDepositsHeader(sheet); }

    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
    const email = Session.getActiveUser().getEmail();
    let id = data.id;

    if (id) {
      const sheetData = sheet.getDataRange().getValues();
      for (let i = 1; i < sheetData.length; i++) {
        if (String(sheetData[i][0]) === String(id)) {
          const rowValues = [
            id, now, data.date || now, data.estimateId || "", data.client || "", data.project || "",
            data.type || "振込", Number(data.amount) || 0, Number(data.fee) || 0, Number(data.offset) || 0,
            data.remarks || "", data.status || "確認済", email, data.visibility || "public"
          ];
          sheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
          invalidateDataCache_();
          return JSON.stringify({ success: true, id: id });
        }
      }
    }

    id = id || getNextDepositId_();
    const rowValues = [
      id, now, data.date || now, data.estimateId || "", data.client || "", data.project || "",
      data.type || "振込", Number(data.amount) || 0, Number(data.fee) || 0, Number(data.offset) || 0,
      data.remarks || "", data.status || "確認済", email, data.visibility || "public"
    ];
    sheet.appendRow(rowValues);
    invalidateDataCache_();
    return JSON.stringify({ success: true, id: id });
  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

function apiSavePayment(jsonData) {
  const data = JSON.parse(jsonData);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "Busy" });

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.sheetNames.payments);
    if (!sheet) { sheet = ss.insertSheet(CONFIG.sheetNames.payments); checkAndFixPaymentsHeader(sheet); }

    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
    const email = Session.getActiveUser().getEmail();
    let id = data.id;

    if (id) {
      const sheetData = sheet.getDataRange().getValues();
      for (let i = 1; i < sheetData.length; i++) {
        if (String(sheetData[i][0]) === String(id)) {
          const rowValues = [
            id, now, data.date || now, data.orderId || "", data.invoiceId || "", data.supplier || "", data.project || "",
            data.type || "振込", Number(data.amount) || 0, Number(data.fee) || 0, Number(data.offset) || 0,
            data.remarks || "", data.status || "確認済", email, data.visibility || "public"
          ];
          sheet.getRange(i + 1, 1, 1, rowValues.length).setValues([rowValues]);
          invalidateDataCache_();
          return JSON.stringify({ success: true, id: id });
        }
      }
    }

    id = id || getNextPaymentId_();
    const rowValues = [
      id, now, data.date || now, data.orderId || "", data.invoiceId || "", data.supplier || "", data.project || "",
      data.type || "振込", Number(data.amount) || 0, Number(data.fee) || 0, Number(data.offset) || 0,
      data.remarks || "", data.status || "確認済", email, data.visibility || "public"
    ];
    sheet.appendRow(rowValues);
    invalidateDataCache_();
    return JSON.stringify({ success: true, id: id });
  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  } finally {
    lock.releaseLock();
  }
}

function apiGetDepositsByEstimate(estimateId) {
  if (!estimateId) return JSON.stringify([]);
  const allJson = apiGetDeposits();
  const all = JSON.parse(allJson);
  const filtered = all.filter(d => String(d.estimateId || "").trim() === String(estimateId).trim());
  return JSON.stringify(filtered);
}

function apiGetPaymentsByOrder(orderId) {
  if (!orderId) return JSON.stringify([]);
  const allJson = apiGetPayments();
  const all = JSON.parse(allJson);
  const filtered = all.filter(p => String(p.orderId || "").trim() === String(orderId).trim());
  return JSON.stringify(filtered);
}

// -----------------------------------------------------------
// 会計・台帳・分析
// -----------------------------------------------------------

function apiGetJournalYears() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const years = new Set();
  const addYearsFromSheet = (sheet, dateColIndex) => {
    if (!sheet || sheet.getLastRow() < 2) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const dStr = data[i][dateColIndex] || data[i][2];
      try { const d = new Date(dStr); if (!isNaN(d.getTime())) years.add(d.getFullYear()); } catch(e) {}
    }
  };
  addYearsFromSheet(ss.getSheetByName(CONFIG.sheetNames.invoice), 7);
  addYearsFromSheet(ss.getSheetByName(CONFIG.sheetNames.deposits), 2);
  addYearsFromSheet(ss.getSheetByName(CONFIG.sheetNames.payments), 2);
  const list = Array.from(years).sort((a, b) => b - a);
  if (list.length === 0) list.push(new Date().getFullYear());
  return JSON.stringify(list);
}

function apiGenerateJournalData(year, month, includeSales, includePurchases) {
  const previewJson = apiPreviewJournalData(year, month, includeSales, includePurchases);
  const data = JSON.parse(previewJson);
  if (data.rows.length === 0) return JSON.stringify({ error: "対象データがありません" });
  const csvRows = [];
  csvRows.push(data.headers.map(v => `"${String(v).replace(/"/g, '""')}"`).join(","));
  data.rows.forEach(row => { csvRows.push(row.map(v => `"${String(v).replace(/"/g, '""')}"`).join(",")); });
  const csvString = csvRows.join("\r\n");
  const blob = Utilities.newBlob('\uFEFF' + csvString, 'text/csv', `集計表_${year}年${month}月.csv`);
  return JSON.stringify({ success: true, data: Utilities.base64Encode(blob.getBytes()), filename: `集計表_${year}年${month}月.csv`, count: data.rows.length });
}

function apiPreviewJournalData(year, month, includeSales, includePurchases) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(CONFIG.sheetNames.journalConfig);
  if (!configSheet) { configSheet = ss.insertSheet(CONFIG.sheetNames.journalConfig); checkAndFixJournalConfig(configSheet); }
  const configRaw = configSheet.getDataRange().getValues();
  const configHeaders = configRaw.slice(1).filter(r => r[0]); 
  const configPurchase = configHeaders.filter(r => r[4] === '仕入' || r[4] === '共通').sort((a, b) => (Number(a[3])||999) - (Number(b[3])||999));
  const configSales = configHeaders.filter(r => r[4] === '売上' || r[4] === '共通').sort((a, b) => (Number(a[3])||999) - (Number(b[3])||999));
  const targetConfig = (includeSales && includePurchases) ? configSales.concat(configPurchase.filter(c => !configSales.some(s => s[0] === c[0] && s[3] === c[3]))) : (includePurchases ? configPurchase : configSales);
  const headers = targetConfig.map(c => c[0]);
  const rows = [];
  let salesClients = [];
  let purchaseSuppliers = [];
  if (configRaw.length > 1) {
      salesClients = configRaw.slice(1).map(r => String(r[6]||"").trim()).filter(String);
      purchaseSuppliers = configRaw.slice(1).map(r => String(r[7]||"").trim()).filter(String);
  }
  if (includePurchases) {
    const agg = {};
    const initPurchaseAgg = () => ({ amount: 0, offset: 0, cash: 0, check: 0, bill: 0, transfer: 0, other: 0 });
    purchaseSuppliers.forEach(s => agg[s] = initPurchaseAgg());
    const iSheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
    if (iSheet && iSheet.getLastRow() > 1) {
        const iData = iSheet.getDataRange().getValues();
        for (let i = 1; i < iData.length; i++) {
            const row = iData[i];
            if (row[1] !== '確認済' && row[1] !== '支払済') continue;
            let dStr = row[7] || row[2];
            const date = new Date(dStr);
            if (isNaN(date.getTime()) || date.getFullYear() != year || (date.getMonth() + 1) != month) continue;
            const supplier = String(row[6]).trim();
            if (!agg[supplier]) {
                if (purchaseSuppliers.length === 0) agg[supplier] = initPurchaseAgg(); 
                else continue; 
            }
            agg[supplier].amount += (Number(row[8]) || 0);
            agg[supplier].offset += (Number(row[9]) || 0);
        }
    }
    const pSheet = ss.getSheetByName(CONFIG.sheetNames.payments);
    if (pSheet && pSheet.getLastRow() > 1) {
        const pData = pSheet.getDataRange().getValues();
        for (let i = 1; i < pData.length; i++) {
            const row = pData[i];
            if (row[12] === '取消') continue;
            const date = new Date(row[2]);
            if (isNaN(date.getTime()) || date.getFullYear() != year || (date.getMonth() + 1) != month) continue;
            const supplier = String(row[5]).trim();
            if (!agg[supplier]) {
                if (purchaseSuppliers.length === 0) agg[supplier] = initPurchaseAgg();
                else continue;
            }
            const type = String(row[7]).trim();
            const amount = Number(row[8]) || 0;
            if (type === '現金') agg[supplier].cash += amount;
            else if (type === '小切手') agg[supplier].check += amount;
            else if (type === '手形') agg[supplier].bill += amount;
            else if (type === '振込') agg[supplier].transfer += amount;
            else if (type === '相殺') agg[supplier].offset += (Number(row[10]) || 0);
            else agg[supplier].other += amount;
        }
    }
    const targetSuppliers = purchaseSuppliers.length > 0 ? purchaseSuppliers : Object.keys(agg);
    targetSuppliers.forEach(supplier => {
        const data = agg[supplier] || initPurchaseAgg();
        const rowData = configPurchase.map(c => {
            const source = c[1]; const fixed = c[2];
            if (source === "fixed") return fixed;
            if (source === "date") return `${year}/${String(month).padStart(2,'0')}`;
            if (source === "amount") return data.amount;
            if (source === "offset") return data.offset;
            if (source === "cash") return data.cash || 0;
            if (source === "check") return data.check || 0;
            if (source === "bill") return data.bill || 0;
            if (source === "transfer") return data.transfer || 0;
            if (source === "other") return data.other || 0;
            if (source === "cash_check") return (data.cash || 0) + (data.check || 0);
            if (source === "supplier" || source === "client") return supplier;
            return "";
        });
        rows.push(rowData);
    });
  }
  if (includeSales) {
    const agg = {};
    const initSalesAgg = () => ({ amount: 0, cash: 0, check: 0, bill: 0, transfer: 0, other: 0 });
    salesClients.forEach(c => agg[c] = initSalesAgg());
    const lSheet = ss.getSheetByName(CONFIG.sheetNames.list);
    if (lSheet && lSheet.getLastRow() > 1) {
        const lData = lSheet.getDataRange().getValues();
        let currentId = "", headerRow = null, tempAmount = 0;
        const processAgg = () => {
            if (!currentId || !headerRow) return;
            const status = headerRow[17];
            if (status !== '請求済' && status !== '完了') return;
            const date = new Date(headerRow[1]);
            if (isNaN(date.getTime()) || date.getFullYear() != year || (date.getMonth() + 1) != month) return;
            const client = String(headerRow[2]).trim();
            if (!agg[client]) {
                if (salesClients.length === 0) agg[client] = initSalesAgg();
                else return;
            }
            agg[client].amount += tempAmount;
        };
        for (let i = 1; i < lData.length; i++) {
            const row = lData[i]; const id = String(row[0]);
            if (id) { processAgg(); currentId = id; headerRow = row; tempAmount = 0; }
            if (currentId) tempAmount += Number(row[10]) || 0;
        }
        processAgg();
    }
    const dSheet = ss.getSheetByName(CONFIG.sheetNames.deposits);
    if (dSheet && dSheet.getLastRow() > 1) {
        const dData = dSheet.getDataRange().getValues();
        for (let i = 1; i < dData.length; i++) {
            const row = dData[i];
            if (row[11] === '取消') continue;
            const date = new Date(row[2]);
            if (isNaN(date.getTime()) || date.getFullYear() != year || (date.getMonth() + 1) != month) continue;
            const client = String(row[4]).trim();
            if (!agg[client]) {
                if (salesClients.length === 0) agg[client] = initSalesAgg();
                else continue;
            }
            const type = String(row[6]).trim();
            const amount = Number(row[7]) || 0;
            if (type === '現金') agg[client].cash += amount;
            else if (type === '小切手') agg[client].check += amount;
            else if (type === '手形') agg[client].bill += amount;
            else if (type === '振込') agg[client].transfer += amount;
            else agg[client].other += amount;
        }
    }
    const targetClients = salesClients.length > 0 ? salesClients : Object.keys(agg);
    targetClients.forEach(client => {
        const data = agg[client] || initSalesAgg();
        const rowData = configSales.map(c => {
            const source = c[1]; const fixed = c[2];
            if (source === "fixed") return fixed;
            if (source === "date") return `${year}/${String(month).padStart(2,'0')}`;
            if (source === "amount") return data.amount;
            if (source === "cash") return data.cash || 0;
            if (source === "check") return data.check || 0;
            if (source === "bill") return data.bill || 0;
            if (source === "transfer") return data.transfer || 0;
            if (source === "other") return data.other || 0;
            if (source === "cash_check") return (data.cash || 0) + (data.check || 0);
            if (source === "supplier" || source === "client") return client;
            return "";
        });
        rows.push(rowData);
    });
  }
  return JSON.stringify({ headers: headers, rows: rows });
}

function apiPredictUnitPrice(product, spec) {
  if (!CONFIG.API_KEY) return JSON.stringify({ error: "APIキーなし" });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.list);
  if (!sheet) return JSON.stringify({ error: "データなし" });
  const data = sheet.getDataRange().getValues();
  const historyLines = [];
  for (let i = data.length - 1; i > 0 && historyLines.length < 50; i--) {
    const row = data[i];
    if (row[4] && row[9]) { historyLines.push(`品名:${row[4]} | 仕様:${row[5]} | 単価:${row[9]} | 単位:${row[7]}`); }
  }
  const context = historyLines.join("\n");
  const prompt = `あなたは建築積算のプロです。以下の過去実績を参考に、新しい項目の適正単価(数値のみ)を予測してください。\n【過去実績】\n${context}\n【対象】\n品名: ${product}\n仕様: ${spec}\n回答は数値(円)のみ。予測不能なら0。`;
  try {
    const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${CONFIG.API_KEY}`, {
      method: "post", contentType: "application/json", payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }), muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    const text = json.candidates[0].content.parts[0].text;
    return JSON.stringify({ price: parseCurrency(text) });
  } catch (e) { return JSON.stringify({ error: e.toString() }); }
}

function apiGetAnalysisData(year) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "analysis_" + year;
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listSheet = ss.getSheetByName(CONFIG.sheetNames.list);
  const orderSheet = ss.getSheetByName(CONFIG.sheetNames.order);
  const projectMap = {}; // id -> { date, client, totalAmount, totalOrderAmount }

  if (listSheet && listSheet.getLastRow() > 1) {
    const listData = listSheet.getRange(2, 1, listSheet.getLastRow() - 1, 20).getValues();
    let currentId = "";
    for (let i = 0; i < listData.length; i++) {
      const row = listData[i];
      const id = String(row[0]);
      if (id) {
        currentId = id;
        if (!projectMap[currentId]) {
          projectMap[currentId] = { date: row[1], client: row[2] || "(不明)", totalAmount: 0, totalOrderAmount: 0 };
        }
      }
      if (currentId) projectMap[currentId].totalAmount += Number(row[10]) || 0;
    }
  }

  if (orderSheet && orderSheet.getLastRow() > 1) {
    const oData = orderSheet.getDataRange().getDisplayValues();
    let hIdx = 0;
    for (let i = 0; i < Math.min(10, oData.length); i++) { if (oData[i][0] === 'ID') { hIdx = i; break; } }
    const h = oData[hIdx];
    const col = {}; h.forEach((v, i) => col[String(v).trim()] = i);
    const idxEstId = col['関連見積ID']; const idxAmount = col['金額'];
    if (idxEstId !== undefined && idxAmount !== undefined) {
      for (let i = hIdx + 1; i < oData.length; i++) {
        const row = oData[i];
        const estId = row[idxEstId];
        if (!estId) continue;
        if (!projectMap[estId]) projectMap[estId] = { date: "", client: "(不明)", totalAmount: 0, totalOrderAmount: 0 };
        projectMap[estId].totalOrderAmount += parseCurrency(row[idxAmount]);
      }
    }
  }

  const monthlyStats = Array(12).fill(0).map(() => ({ sales: 0, cost: 0, profit: 0 }));
  const clientStats = {};
  Object.values(projectMap).forEach(p => {
    const d = new Date(p.date);
    if (isNaN(d.getTime())) return;
    if (d.getFullYear() != year) return;
    const monthIdx = d.getMonth();
    const sales = Number(p.totalAmount) || 0;
    const cost = Number(p.totalOrderAmount) || 0;
    const profit = sales - cost;
    monthlyStats[monthIdx].sales += sales;
    monthlyStats[monthIdx].cost += cost;
    monthlyStats[monthIdx].profit += profit;
    const client = p.client || "(不明)";
    if (!clientStats[client]) clientStats[client] = { name: client, sales: 0, profit: 0, count: 0 };
    clientStats[client].sales += sales;
    clientStats[client].profit += profit;
    clientStats[client].count += 1;
  });
  const clientRanking = Object.values(clientStats).sort((a, b) => b.sales - a.sales).slice(0, 10);
  const result = JSON.stringify({ monthly: monthlyStats, ranking: clientRanking });
  try { cache.put(cacheKey, result, CACHE_TTL_SHORT); } catch (e) { console.warn("Cache put failed (analysis): " + e.message); }
  return result;
}

function apiGetProjectLedger(projectId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const estSheet = ss.getSheetByName(CONFIG.sheetNames.list);
  let estimate = { header: {}, items: [] };
  if (estSheet) {
    const rawEst = _getEstimateData(projectId); 
    if (rawEst) estimate = rawEst;
  }
  const ordSheet = ss.getSheetByName(CONFIG.sheetNames.order);
  const orders = [];
  if (ordSheet) {
    const oData = ordSheet.getDataRange().getDisplayValues();
    for(let i=1; i<oData.length; i++) {
      const r = oData[i];
      if(r[3] === projectId || String(r[3]).startsWith(projectId + '-')) {
        orders.push({ date: r[1], vendor: r[2], item: `${r[5]} ${r[6]}`, amount: parseCurrency(r[10]) });
      }
    }
  }
  const invSheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
  const invoices = [];
  if (invSheet) {
    const iData = invSheet.getDataRange().getDisplayValues();
    for(let i=1; i<iData.length; i++) {
      const r = iData[i];
      if(r[4] === projectId || String(r[4]).startsWith(projectId + '-')) {
        invoices.push({ date: r[7] || r[2], vendor: r[6], item: r[11], amount: parseCurrency(r[10]) });
      }
    }
  }

  // 入金データ取得
  const depSheet = ss.getSheetByName(CONFIG.sheetNames.deposits);
  const depositEntries = [];
  let totalDepositAmount = 0;
  if (depSheet && depSheet.getLastRow() > 1) {
    const dData = depSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < dData.length; i++) {
      const r = dData[i];
      if (!r[0]) continue;
      const estId = String(r[3]).trim();
      if (estId === projectId || estId.startsWith(projectId + '-')) {
        const status = String(r[11]).trim();
        if (status === '取消') continue;
        const amount = parseCurrency(r[7]);
        const fee = parseCurrency(r[8]);
        depositEntries.push({ date: r[2], client: r[4], type: r[6], amount: amount, fee: fee, remarks: r[10] || '' });
        totalDepositAmount += amount;
      }
    }
  }

  // 出金データ取得
  const paySheet = ss.getSheetByName(CONFIG.sheetNames.payments);
  const paymentEntries = [];
  let totalPaymentAmount = 0;
  if (paySheet && paySheet.getLastRow() > 1) {
    const pData = paySheet.getDataRange().getDisplayValues();
    for (let i = 1; i < pData.length; i++) {
      const r = pData[i];
      if (!r[0]) continue;
      // 出金は発注IDで紐付けるため、発注IDから関連見積IDを逆引き
      const orderId = String(r[3]).trim();
      const status = String(r[12]).trim();
      if (status === '取消') continue;
      const amount = parseCurrency(r[8]);
      const fee = parseCurrency(r[9]);
      // 工事名でも照合
      const payProject = String(r[6]).trim();
      const payEstHeader = estimate.header || {};
      const matchByOrder = orderId && orders.some(o => true); // 発注IDがある場合は発注明細と照合
      const matchByProject = payProject && (payProject === (payEstHeader.project || '') || payProject === (payEstHeader.client || ''));
      
      // 関連発注IDで照合: 発注IDから関連見積IDを取得
      let matchByOrderEstId = false;
      if (orderId && ordSheet) {
        const oData2 = ordSheet.getDataRange().getDisplayValues();
        for (let j = 1; j < oData2.length; j++) {
          if (oData2[j][0] === orderId && (oData2[j][3] === projectId || String(oData2[j][3]).startsWith(projectId + '-'))) {
            matchByOrderEstId = true;
            break;
          }
        }
      }
      
      if (matchByOrderEstId || matchByProject) {
        paymentEntries.push({ date: r[2], supplier: r[5], type: r[7], amount: amount, fee: fee, remarks: r[11] || '' });
        totalPaymentAmount += amount;
      }
    }
  }

  const totalSales = estimate.totalAmount || 0;
  const totalOrder = orders.reduce((s,o) => s + o.amount, 0);
  const totalInvoicePayment = invoices.reduce((s,i) => s + i.amount, 0);
  const profit = totalSales - totalOrder; 
  const profitRate = totalSales ? ((profit / totalSales) * 100).toFixed(1) : 0;
  return JSON.stringify({
    project: estimate.header, sales: totalSales, totalOrder: totalOrder, totalPayment: totalInvoicePayment,
    profit: profit, profitRate: profitRate, orders: orders, invoices: invoices,
    deposits: depositEntries, totalDeposit: totalDepositAmount,
    payments: paymentEntries, totalWithdrawal: totalPaymentAmount
  });
}

function apiCreateLedgerPdf(jsonData) {
  const data = JSON.parse(jsonData);
  const now = new Date();
  data.printDate = getJapaneseDateStr(now);
  let template;
  try { template = HtmlService.createTemplateFromFile('ledger_template'); } 
  catch(e) { return JSON.stringify({ success: false, message: "ledger_template.html missing" }); }
  template.data = data;
  const blob = Utilities.newBlob(template.evaluate().getContent(), MimeType.HTML).getAs(MimeType.PDF).setName(`工事台帳_${data.project.project}.pdf`);
  const folder = getSaveFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return JSON.stringify({ success: true, url: file.getUrl() });
}

function apiDeleteData(id) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return JSON.stringify({ success: false, message: "Busy" });
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let deleted = false;
    if (deleteRowsById(ss.getSheetByName(CONFIG.sheetNames.list), id)) deleted = true;
    if (deleteRowsById(ss.getSheetByName(CONFIG.sheetNames.order), id)) deleted = true;
    if (deleteRowsById(ss.getSheetByName(CONFIG.sheetNames.invoice), id)) deleted = true;
    if (deleteRowsById(ss.getSheetByName(CONFIG.sheetNames.deposits), id)) deleted = true;
    if (deleteRowsById(ss.getSheetByName(CONFIG.sheetNames.payments), id)) deleted = true;
    invalidateDataCache_();
    return JSON.stringify({ success: deleted, message: deleted ? "" : "Not found" });
  } catch (e) { return JSON.stringify({ success: false, message: e.toString() }); } finally { lock.releaseLock(); }
}

/** @deprecated フロントエンドから未使用。将来の検索パネル用に残置 */
function apiSearchItems(keyword, type) {
  return JSON.stringify([]);
}

/**
 * 一括初期化API（軽量版）: 起動時は認証とマスタのみ取得
 * projects, orders, products, invoices は画面遷移時に遅延取得で起動時間を短縮
 */
function apiBatchInit() {
  const results = {};
  results.auth = apiGetAuthStatus();
  results.masters = apiGetMasters();
  return JSON.stringify(results);
}

// ==========================================
// AIチャットボット機能 (System Expert Bot)
// ==========================================

/**
 * AIチャットボットAPI
 * 知識ファイルを読み込み、ユーザーの質問に対する回答を生成します。
 * @param {string} userMessage - ユーザーからの質問
 * @return {string} JSON形式の回答 { reply: "...", error: "..." }
 */
function apiChatWithSystemBot(userMessage) {
  // APIキーの確認
  if (!CONFIG.API_KEY) {
    return JSON.stringify({ error: "APIキーが設定されていません。管理者に連絡してください。" });
  }

  try {
    // 1. 知識ファイル取得
    const props = PropertiesService.getScriptProperties();
    const knowledgeFileId = props.getProperty('system_knowledge');
    
    if (!knowledgeFileId) {
      return JSON.stringify({ error: "知識ファイルIDが設定されていません。管理者に連絡してください。" });
    }

    const knowledgeFile = DriveApp.getFileById(knowledgeFileId);
    const systemContext = knowledgeFile.getBlob().getDataAsString();

    // 2. プロンプトの構築
    const promptText = `
あなたは「AI建築見積システム」の操作サポート専門アシスタントです。
以下の【システム情報】を基に、ユーザーからの質問に答えてください。

■回答ルール
1. 操作方法の質問には、具体的なボタン名や画面上の場所、手順を簡潔に案内してください
2. 技術的な仕組みや実装の詳細は、明示的に聞かれた場合のみ説明してください
3. 回答は以下の形式で統一してください：
   - マークダウン記号（**、##、###など）は使わない
   - 箇条書きは「・」を使用
   - 手順は「1. 」「2. 」のように番号付き
   - 適度な改行で読みやすく整形
   - シンプルで分かりやすい日本語

■制約事項
・【システム情報】に記載されていないことは「わかりません」と答えてください
・回答は日本語のみで行ってください

【システム情報】
${systemContext}

【ユーザーの質問】
${userMessage}
`;

    // 3. Gemini APIへのリクエスト
    const payload = {
      contents: [{ parts: [{ text: promptText }] }],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 2000
      }
    };

    const res = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${CONFIG.API_KEY}`,
      {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    const json = JSON.parse(res.getContentText());
    
    // エラーハンドリング
    if (json.error) {
      console.error("Gemini API Error: " + JSON.stringify(json.error));
      return JSON.stringify({ error: "AIの応答エラー: " + json.error.message });
    }
    
    // 回答の抽出
    const reply = json.candidates && json.candidates[0].content.parts[0].text;
    if (!reply) {
      return JSON.stringify({ error: "回答を生成できませんでした。" });
    }

    return JSON.stringify({ reply: reply });

  } catch (e) {
    console.error("apiChatWithSystemBot Exception: " + e.toString());
    return JSON.stringify({ error: "システムエラーが発生しました: " + e.toString() });
  }
}