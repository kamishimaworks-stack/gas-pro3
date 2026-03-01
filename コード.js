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
  LINE_TOKEN:          PROPS.LINE_CHANNEL_TOKEN || '',
  LINE_USER_ID:        PROPS.LINE_USER_ID || '',
  LINE_CHANNEL_SECRET: PROPS.LINE_CHANNEL_SECRET || '',
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
    journalConfig: '仕訳設定マスタ',
    masterMaterial: '材料費単価マスタ',
    masterMaterialIncl: '材料（施工込）単価マスタ',
    progressDb: '出来高DB'
  }
};

// 出来高DB 列インデックス（1-based: GAS Range用）
const PROG_COL = {
  no: 1, name: 2, spec: 3, totalQty: 4, unit: 5, price: 6,
  estimateAmt: 7, prevCumQty: 8, currCumQty: 9,
  progressAmt: 10, progressRate: 11, periodQty: 12, periodPayment: 13,
  estId: 14, orderId: 15, reportMonth: 16
};

const PROGRESS_HEADERS = [
  'No.', '品名', '仕様', '全体数量', '単位', '単価', '見積金額',
  '前月末累積数量', '現在の累積数量', '出来高金額', '出来高比率',
  '今回数量', '今回支払金額', '関連見積ID', '関連発注ID', '報告月'
];

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
    c.remove("material_prices");
    c.remove("progress_data_all");
    c.remove("progress_report_list");
    c.remove("progress_report_list_all");
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
 * 請求書ファイル連番を採番（イニシャル別に4桁連番）
 */
function getNextInvoiceFileNo_(initial) {
  const props = PropertiesService.getScriptProperties();
  const key = 'SEQ_INV_FILE_' + (initial || 'X');
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  try {
    lockAcquired = lock.tryLock(5000);
    if (lockAcquired) {
      let current = Number(props.getProperty(key)) || 0;
      current++;
      props.setProperty(key, String(current));
      return String(current).padStart(4, '0');
    }
    return '0000';
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

// 漢字→ローマ字イニシャル変換テーブル（共通）
const KANJI_ROMAJI_ = {
  '阿':'A','安':'A','青':'A','赤':'A','秋':'A','朝':'A','浅':'A','荒':'A','有':'A','新':'A','飯':'I','池':'I','石':'I','泉':'I','井':'I','伊':'I','磯':'I','一':'I','稲':'I','今':'I','岩':'I','上':'U','内':'U','宇':'U','梅':'U','浦':'U','遠':'E','江':'E','榎':'E','大':'O','岡':'O','奥':'O','小':'O','尾':'O','荻':'O',
  '加':'K','柿':'K','角':'K','笠':'K','片':'K','金':'K','鎌':'K','上':'K','亀':'K','川':'K','河':'K','神':'K','菊':'K','岸':'K','北':'K','木':'K','吉':'K','久':'K','国':'K','熊':'K','栗':'K','黒':'K','桑':'K','小':'K','古':'K','後':'G','五':'G',
  '佐':'S','斉':'S','斎':'S','坂':'S','桜':'S','笹':'S','沢':'S','澤':'S','塩':'S','柴':'S','島':'S','嶋':'S','清':'S','白':'S','新':'S','進':'S','杉':'S','鈴':'S','須':'S','関':'S','瀬':'S',
  '高':'T','竹':'T','田':'T','谷':'T','丹':'T','千':'T','塚':'T','土':'T','鶴':'T','寺':'T','天':'T','東':'T','徳':'T','富':'T','豊':'T',
  '中':'N','永':'N','長':'N','西':'N','二':'N','野':'N','能':'N',
  '橋':'H','畑':'H','浜':'H','濱':'H','林':'H','原':'H','春':'H','樋':'H','久':'H','平':'H','広':'H','廣':'H','蜂':'H','長谷':'H','羽':'H','花':'H','福':'F','藤':'F','船':'F','古':'F',
  '前':'M','牧':'M','松':'M','丸':'M','三':'M','水':'M','溝':'M','南':'M','宮':'M','村':'M','森':'M','諸':'M',
  '八':'Y','山':'Y','矢':'Y','柳':'Y','横':'Y','吉':'Y','米':'Y',
  '若':'W','渡':'W','和':'W',
  '栄':'S','相':'A','足':'A','天':'A','綾':'A','粟':'A',
  '利':'R','陸':'R','龍':'R','竜':'R',
  'A':'A','B':'B','C':'C','D':'D','E':'E','F':'F','G':'G','H':'H','I':'I','J':'J','K':'K','L':'L','M':'M','N':'N','O':'O','P':'P','Q':'Q','R':'R','S':'S','T':'T','U':'U','V':'V','W':'W','X':'X','Y':'Y','Z':'Z',
  'Ａ':'A','Ｂ':'B','Ｃ':'C','Ｄ':'D','Ｅ':'E','Ｆ':'F','Ｇ':'G','Ｈ':'H','Ｉ':'I','Ｊ':'J','Ｋ':'K','Ｌ':'L','Ｍ':'M','Ｎ':'N','Ｏ':'O','Ｐ':'P','Ｑ':'Q','Ｒ':'R','Ｓ':'S','Ｔ':'T','Ｕ':'U','Ｖ':'V','Ｗ':'W','Ｘ':'X','Ｙ':'Y','Ｚ':'Z',
  'あ':'A','い':'I','う':'U','え':'E','お':'O','か':'K','き':'K','く':'K','け':'K','こ':'K','さ':'S','し':'S','す':'S','せ':'S','そ':'S','た':'T','ち':'T','つ':'T','て':'T','と':'T','な':'N','に':'N','ぬ':'N','ね':'N','の':'N','は':'H','ひ':'H','ふ':'F','へ':'H','ほ':'H','ま':'M','み':'M','む':'M','め':'M','も':'M','や':'Y','ゆ':'Y','よ':'Y','ら':'R','り':'R','る':'R','れ':'R','ろ':'R','わ':'W',
  'ア':'A','イ':'I','ウ':'U','エ':'E','オ':'O','カ':'K','キ':'K','ク':'K','ケ':'K','コ':'K','サ':'S','シ':'S','ス':'S','セ':'S','ソ':'S','タ':'T','チ':'T','ツ':'T','テ':'T','ト':'T','ナ':'N','ニ':'N','ヌ':'N','ネ':'N','ノ':'N','ハ':'H','ヒ':'H','フ':'F','ヘ':'H','ホ':'H','マ':'M','ミ':'M','ム':'M','メ':'M','モ':'M','ヤ':'Y','ユ':'Y','ヨ':'Y','ラ':'R','リ':'R','ル':'R','レ':'R','ロ':'R','ワ':'W'
};

/**
 * 企業名からイニシャルを推定（漢字→ローマ字変換）
 */
function guessInitialFromName_(name) {
  if (!name) return '';
  const cleaned = name.trim().replace(/^(株式会社|有限会社|合同会社|（株）|\(株\))/, '');
  const first = cleaned.charAt(0);
  return KANJI_ROMAJI_[first] || '';
}

/**
 * 会計年度を取得（4月始まり：4〜12月→その年、1〜3月→前年）
 */
function getFiscalYear_(date) {
  const d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return new Date().getFullYear();
  const month = d.getMonth() + 1;
  return month >= 4 ? d.getFullYear() : d.getFullYear() - 1;
}

/**
 * 同一年度内で同じ工事名の既存工事番号を検索
 */
function lookupConstructionNumber_(projectName, date) {
  if (!projectName) return '';
  const fiscalYear = getFiscalYear_(date || new Date());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
  if (!sheet || sheet.getLastRow() < 2) return '';
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const existingProject = String(row[5]).trim();
    const existingDate = row[7];
    const existingConstructionId = String(row[4]).trim();
    if (existingProject === projectName.trim() &&
        existingConstructionId &&
        getFiscalYear_(existingDate) === fiscalYear) {
      return existingConstructionId;
    }
  }
  return '';
}

/**
 * 工事番号を決定（既存再利用 or 新規採番）
 */
function resolveConstructionId_(data) {
  if (data.constructionId) return data.constructionId;
  const personName = data.person || '';
  // 同一年度内の同じ工事名から再利用
  const existing = lookupConstructionNumber_(data.project, data.date || new Date());
  if (existing) return existing;
  // イニシャル決定: 明示指定 > 担当者名から推定
  let initial = '';
  if (data.initial) {
    initial = data.initial.toUpperCase().replace(/[^A-Z]/g, '');
  }
  if (!initial) {
    initial = guessInitialFromName_(personName) || getVendorInitial_(personName);
  }
  if (!initial) initial = 'X';
  // 新規採番
  const seqNo = getNextInvoiceFileNo_(initial);
  return initial + seqNo;
}

/**
 * 完了自動処理: 同一工事名・担当者・施工者のレコードを全て「完了」に更新
 */
function apiAutoCompleteOnKanryo_(projectName, personName, contractorName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
  if (!sheet || sheet.getLastRow() < 2) return { updated: 0, ids: [] };
  const data = sheet.getDataRange().getValues();
  const updatedIds = [];
  const pj = (projectName || '').trim();
  const ps = (personName || '').trim();
  const ct = (contractorName || '').trim();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowProject = String(row[5]).trim();
    const rowPerson = String(row[6]).trim();
    const rowContractor = String(row[14] || '').trim();
    const status = String(row[1]);
    if (status === '完了' || status === '支払済') continue;
    // 工事名＋担当者は必須一致、施工者は指定時のみチェック
    if (rowProject === pj && rowPerson === ps &&
        (!ct || !rowContractor || rowContractor === ct)) {
      sheet.getRange(i + 1, 2).setValue('完了');
      updatedIds.push(String(row[0]));
    }
  }
  if (updatedIds.length > 0) invalidateDataCache_();
  return { updated: updatedIds.length, ids: updatedIds };
}

/**
 * 発注先マスタからイニシャルを取得（未設定なら企業名から推定）
 */
function getVendorInitial_(supplierName) {
  if (!supplierName) return '';
  // 1. マスタから検索
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vSheet = ss.getSheetByName(CONFIG.sheetNames.masterVendor);
    if (vSheet && vSheet.getLastRow() >= 2) {
      const cols = Math.min(vSheet.getLastColumn(), 5);
      const data = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, cols).getValues();
      const target = supplierName.trim();
      for (const r of data) {
        const name = String(r[1]).trim();
        if (name && (target === name || target.startsWith(name) || name.startsWith(target))) {
          const initial = cols >= 5 && r[4] ? String(r[4]).trim() : '';
          if (initial) return initial.toUpperCase();
        }
      }
    }
  } catch (e) { /* ignore */ }
  // 2. フォールバック: 企業名から推定
  return guessInitialFromName_(supplierName);
}

/**
 * 発注先マスタのE列にローマ字イニシャルを一括入力（未入力のみ）
 * GASエディタから手動実行: fillVendorInitials()
 */
function fillVendorInitials() {
  const KANJI_ROMAJI = KANJI_ROMAJI_;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.masterVendor);
  if (!sheet || sheet.getLastRow() < 2) return;
  // E1にヘッダーがなければ追加
  if (!sheet.getRange(1, 5).getValue()) sheet.getRange(1, 5).setValue('イニシャル');
  const lastRow = sheet.getLastRow();
  const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const existing = sheet.getRange(2, 5, lastRow - 1, 1).getValues();
  const updates = [];
  for (let i = 0; i < names.length; i++) {
    if (existing[i][0]) { updates.push([existing[i][0]]); continue; }
    const name = String(names[i][0]).trim().replace(/^(株式会社|有限会社|合同会社|（株）|\(株\))/, '');
    const first = name.charAt(0);
    updates.push([KANJI_ROMAJI[first] || '']);
  }
  sheet.getRange(2, 5, updates.length, 1).setValues(updates);
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

// 関連見積IDで発注行を一括削除（自動生成発注の再作成用）
function deleteOrdersByEstimateId_(sheet, estId) {
  if (!sheet || !estId) return;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  // ヘッダー行を探す
  let hIdx = 0;
  for (let i = 0; i < Math.min(10, data.length); i++) { if (String(data[i][0]).trim() === 'ID') { hIdx = i; break; } }
  const headers = data[hIdx];
  let relEstCol = 3; // デフォルト
  for (let i = 0; i < headers.length; i++) { if (String(headers[i]).trim() === '関連見積ID') { relEstCol = i; break; } }

  const rowsToDelete = [];
  let currentId = "";
  let currentRelEstId = "";
  for (let i = hIdx + 1; i < data.length; i++) {
    const rowId = String(data[i][0]).trim();
    if (rowId) {
      currentId = rowId;
      currentRelEstId = String(data[i][relEstCol]).trim();
    }
    if (currentRelEstId === estId) {
      rowsToDelete.push(i + 1);
    }
  }
  if (rowsToDelete.length > 0) {
    deleteRowsOptimized_(sheet, rowsToDelete);
  }
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
// LINE Webhook (doPost)
// -----------------------------------------------------------

function doPost(e) {
  try {
    // URLクエリパラメータでトークン検証
    const params = e.parameter || {};
    const secret = CONFIG.LINE_CHANNEL_SECRET || PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_SECRET') || '';
    if (!secret) {
      console.warn('doPost: LINE_CHANNEL_SECRET is not configured, skipping event processing');
      return ContentService.createTextOutput('OK');
    }
    if (params.token !== secret) {
      console.warn('doPost: invalid token');
      return ContentService.createTextOutput('OK');
    }

    const body = JSON.parse(e.postData.contents);
    if (!body || !Array.isArray(body.events)) {
      return ContentService.createTextOutput('OK');
    }

    body.events.forEach(event => {
      try {
        processLineEvent_(event);
      } catch (err) {
        console.error('processLineEvent_ error: ' + err.toString());
      }
    });
  } catch (ex) {
    console.error('doPost error: ' + ex.toString());
  }
  return ContentService.createTextOutput('OK');
}

function processLineEvent_(event) {
  const token = CONFIG.LINE_TOKEN || PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_TOKEN') || '';
  if (!token) { console.warn('LINE_TOKEN not set'); return; }

  if (event.type === 'follow') {
    // 友だち追加時: ユーザーID自動登録 + ウェルカムメッセージ
    const userId = event.source && event.source.userId;
    if (userId) {
      const existing = getLineUserIds_();
      if (!existing.includes(userId)) {
        existing.push(userId);
        PropertiesService.getScriptProperties().setProperty('LINE_USER_IDS', existing.join(','));
      }
      sendLineReply_(event.replyToken, 'AI建築見積システムと連携しました。\nメッセージを送信すると工事データとして受け付けます。\n\n【入力フォーマット例】\n工事名、○○邸新築工事\n担当者、○○\nイニシャル、T\n着工日、2026/02/20\n施工者、○○建設\n\n※「着工日」の代わりに「完了日」も使えます');
    }
    return;
  }

  if (event.type === 'message' && event.message && event.message.type === 'text') {
    const text = event.message.text;
    if (!text || text.length > 5000) {
      sendLineReply_(event.replyToken, 'メッセージが長すぎます（5000文字以内にしてください）');
      return;
    }

    const parsed = parseLineMessage_(text);
    if (!parsed || !parsed.person) {
      sendLineReply_(event.replyToken, '内容を解析できませんでした。\n以下のフォーマットで送信してください:\n\n工事名、○○邸新築工事\n担当者、○○\nイニシャル、T\n着工日、2026/02/20\n施工者、○○建設\n\n※「着工日」の代わりに「完了日」も使えます');
      return;
    }

    // === 完了通知の場合: ステータス更新のみ（ファイル作成・工事番号採番しない） ===
    if (parsed.detectedKanryo && parsed.project && parsed.person) {
      const result = apiAutoCompleteOnKanryo_(parsed.project, parsed.person, parsed.contractor || '');
      let reply = '';
      if (result.updated > 0) {
        reply += '【完了処理】' + result.updated + '件のレコードを「完了」に更新しました。\n\n';
      } else {
        reply += '完了対象の着工中レコードが見つかりませんでした。\n工事名・担当者が登録済みデータと一致しているか確認してください。\n\n';
      }
      if (parsed.project) reply += '工事名: ' + parsed.project + '\n';
      if (parsed.person) reply += '担当者: ' + parsed.person + '\n';
      if (parsed.contractor) reply += '施工者: ' + parsed.contractor + '\n';
      if (parsed.date) reply += '完了日: ' + parsed.date + '\n';
      sendLineReply_(event.replyToken, reply);
      return;
    }

    // === 通常の工事データ登録 ===
    const folderId = CONFIG.invoiceInputFolder;
    if (!folderId) {
      sendLineReply_(event.replyToken, '請求書受取フォルダが未設定です。管理者に連絡してください。');
      return;
    }

    const now = new Date();
    const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const projectName = (parsed.project || '').replace(/[\\/:*?"<>|]/g, '');
    const personName = (parsed.person || '').replace(/[\\/:*?"<>|]/g, '');
    const contractorName = (parsed.contractor || '').replace(/[\\/:*?"<>|]/g, '');
    // イニシャル未指定時は担当者名の漢字から推測（1文字）
    const hadExplicitInitial = !!parsed.initial;
    if (!parsed.initial && parsed.person) {
      const inferred = guessInitialFromName_(parsed.person) || getVendorInitial_(parsed.person);
      if (inferred) parsed.initial = inferred.charAt(0);
    }
    if (parsed.initial) parsed.initial = parsed.initial.charAt(0).toUpperCase();
    if (!hadExplicitInitial && parsed.initial) parsed._initialInferred = true;
    // 工事番号を必ず生成（年度内同一工事名チェック含む）
    if (!parsed.constructionId) {
      parsed.constructionId = resolveConstructionId_(parsed);
    }
    let fileName;
    if (parsed.constructionId && projectName && personName) {
      fileName = parsed.constructionId + '_' + projectName + '\u3000' + personName + '.txt';
    } else if (projectName && personName) {
      fileName = projectName + '\u3000' + personName + '.txt';
    } else {
      fileName = 'LINE_請求書_' + ts + '.txt';
    }

    // _parseTextInvoice() が読み取れる key: value 形式でファイル内容を生成
    let content = '';
    content += '工事番号: ' + parsed.constructionId + '\n';
    if (parsed.project) content += '現場名: ' + parsed.project + '\n';
    if (parsed.person) content += '担当者: ' + parsed.person + '\n';
    content += 'イニシャル: ' + (parsed.initial || 'X') + '\n';
    if (parsed.date) content += '着工日: ' + parsed.date + '\n';
    if (parsed.contractor) content += '施工者: ' + parsed.contractor + '\n';
    if (parsed.content) content += '内容: ' + parsed.content + '\n';
    if (parsed.amount) content += '請求金額: ' + parsed.amount + '\n';
    if (parsed.location) content += '工事場所: ' + parsed.location + '\n';

    const folder = DriveApp.getFolderById(folderId);
    folder.createFile(fileName, content, MimeType.PLAIN_TEXT);

    // 請求書ファイルキャッシュを無効化
    try {
      const cache = CacheService.getScriptCache();
      cache.remove("invoice_files_" + String(folderId).slice(-8));
    } catch (e) { /* ignore */ }

    let reply = '工事データを受け付けました。\n';
    reply += 'Webアプリの請求書受取画面で確認・登録してください。\n\n';
    reply += '工事番号: ' + parsed.constructionId + '\n';
    if (parsed.project) reply += '工事名: ' + parsed.project + '\n';
    if (parsed.person) reply += '担当者: ' + parsed.person + '\n';
    reply += 'イニシャル: ' + (parsed.initial || 'X') + (parsed._initialInferred ? '（自動推測）' : '') + '\n';
    if (parsed.date) reply += '着工日: ' + parsed.date + '\n';
    if (parsed.contractor) reply += '施工者: ' + parsed.contractor + '\n';
    if (parsed.content) reply += '内容: ' + parsed.content + '\n';
    if (parsed.amount) reply += '金額: ' + Number(parsed.amount).toLocaleString() + '円\n';
    if (parsed.location) reply += '工事場所: ' + parsed.location + '\n';

    sendLineReply_(event.replyToken, reply);
    return;
  }

  // 画像・ファイルメッセージ対応
  if (event.type === 'message' && event.message &&
      (event.message.type === 'image' || event.message.type === 'file')) {
    const folderId = CONFIG.invoiceInputFolder;
    if (!folderId) {
      sendLineReply_(event.replyToken, '請求書受取フォルダが未設定です。管理者に連絡してください。');
      return;
    }
    const messageId = event.message.id;
    const res = UrlFetchApp.fetch(
      'https://api-data.line.me/v2/bot/message/' + messageId + '/content',
      { method: 'get', headers: { 'Authorization': 'Bearer ' + token } }
    );
    const blob = res.getBlob();
    const now = new Date();
    const ts = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    let fileName;
    if (event.message.type === 'file' && event.message.fileName) {
      fileName = 'LINE_請求書_' + ts + '_' + event.message.fileName;
    } else {
      const ext = (blob.getContentType() || 'image/jpeg').split('/').pop().replace('jpeg', 'jpg');
      fileName = 'LINE_請求書_' + ts + '.' + ext;
    }
    const folder = DriveApp.getFolderById(folderId);
    folder.createFile(blob.setName(fileName));
    // キャッシュ無効化
    try {
      CacheService.getScriptCache().remove("invoice_files_" + String(folderId).slice(-8));
    } catch (e) { /* ignore */ }
    sendLineReply_(event.replyToken, '請求書画像を受け付けました。\nWebアプリの請求書受取画面で確認・登録してください。');
    return;
  }
}

function parseLineMessage_(text) {
  // 1. 構造化フォーマット（正規表現）
  const structured = parseStructuredMessage_(text);
  if (structured && structured.person) return structured;

  // 2. Gemini APIフォールバック
  try {
    return parseLineMessageWithGemini_(text);
  } catch (e) {
    console.error('parseLineMessageWithGemini_ failed: ' + e.toString());
    return null;
  }
}

function parseStructuredMessage_(text) {
  const result = {};
  const patterns = {
    person:         /(?:担当者|担当|請求元|業者名|会社名|企業名)[：:、,\s]*(.+)/,
    contractor:     /(?:施工者|施工業者|施工担当|作業者)[：:、,\s]*(.+)/,
    amount:         /(?:金額|請求金額|合計)[：:、,\s]*([0-9０-９,，]+)/,
    date:           /(?:日付|着工日|施工日|完了日|請求日|発行日|開始日)[：:、,\s]*(.+)/,
    content:        /(?:内容|品名|但し書き)[：:、,\s]*(.+)/,
    project:        /(?:工事名|現場名|案件名)[：:、,\s]*(.+)/,
    constructionId: /(?:工事番号|工事ID)[：:、,\s]*(.+)/,
    location:       /(?:工事場所|現場住所|住所|場所)[：:、,\s]*(.+)/,
    initial:        /(?:イニシャル|頭文字)[：:、,\s]*([A-Za-zＡ-Ｚ])/
  };
  const lines = text.split(/\r?\n/);
  lines.forEach(line => {
    const trimmed = line.trim();
    Object.keys(patterns).forEach(key => {
      const m = trimmed.match(patterns[key]);
      if (m && !result[key]) {
        let val = m[1].trim();
        if (key === 'amount') {
          val = toHalfWidth(val).replace(/[,，]/g, '');
        }
        result[key] = val;
      }
    });
  });
  // 完了キーワード検出（「完了日」指定 or テキストに「完了」を含む場合）
  if (/完了日[：:、,\s]/.test(text) || text.includes('完了')) {
    result.detectedKanryo = true;
  }
  return result.person ? result : null;
}

function parseLineMessageWithGemini_(text) {
  if (!CONFIG.API_KEY) return null;

  const prompt = '以下のテキストから工事に関する情報を抽出してJSON形式で返してください。' +
    '主要キー: project(工事名・現場名), person(担当者・業者名・企業名), initial(イニシャル・頭文字・アルファベット1文字), date(着工日・完了日・日付・yyyy/MM/dd形式), contractor(施工者・施工業者)。' +
    '補助キー: amount(金額・数値), content(内容・品名), constructionId(工事番号), location(工事場所・現場住所), detectedKanryo(テキストに「完了」や「完了日」が含まれるかboolean)。' +
    '該当しない項目はnullにしてください。\n\nテキスト:\n' + text;

  const responseSchema = {
    "type": "OBJECT",
    "properties": {
      "person":         { "type": "STRING" },
      "contractor":     { "type": "STRING" },
      "amount":         { "type": "NUMBER" },
      "date":           { "type": "STRING" },
      "content":        { "type": "STRING" },
      "project":        { "type": "STRING" },
      "constructionId": { "type": "STRING" },
      "location":       { "type": "STRING" },
      "initial":        { "type": "STRING" },
      "detectedKanryo": { "type": "BOOLEAN" }
    }
  };

  const res = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${CONFIG.API_KEY}`,
    {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { response_mime_type: 'application/json', response_schema: responseSchema }
      }),
      muteHttpExceptions: true
    }
  );

  const json = JSON.parse(res.getContentText());
  if (json.error || !json.candidates || !json.candidates[0]) return null;
  const parsed = JSON.parse(json.candidates[0].content.parts[0].text);
  return parsed && parsed.person ? parsed : null;
}


function sendLineReply_(replyToken, message) {
  const token = CONFIG.LINE_TOKEN || PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_TOKEN') || '';
  if (!token || !replyToken) return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post', contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ replyToken: replyToken, messages: [{ type: 'text', text: message }] }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('sendLineReply_ failed: ' + e.toString());
  }
}

function sendLinePush_(userId, message) {
  const token = CONFIG.LINE_TOKEN || PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_TOKEN') || '';
  if (!token || !userId) return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post', contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ to: userId, messages: [{ type: 'text', text: message }] }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('sendLinePush_ failed: ' + e.toString());
  }
}

function apiSaveLineChannelSecret(secret) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('LINE_CHANNEL_SECRET', secret || '');
    CONFIG.LINE_CHANNEL_SECRET = secret || '';
    return JSON.stringify({ success: true, message: 'チャネルシークレットを保存しました' });
  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  }
}

function apiGetWebhookUrl() {
  try {
    const url = ScriptApp.getService().getUrl();
    const secret = CONFIG.LINE_CHANNEL_SECRET || PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_SECRET') || '';
    const webhookUrl = secret ? url + '?token=' + encodeURIComponent(secret) : url;
    return JSON.stringify({ success: true, url: webhookUrl });
  } catch (e) {
    return JSON.stringify({ success: false, message: 'Webhook URLを取得できません。ウェブアプリとしてデプロイしてください。' });
  }
}

// -----------------------------------------------------------
// シート初期化
// -----------------------------------------------------------

function checkAndFixEstimateHeader(sheet) {
  if (!sheet) return;
  const headers = ["ID", "日付", "顧客名", "工種", "品名", "仕様", "数量", "単位", "原価", "単価", "金額", "備考", "工事場所", "工事名", "工期", "支払条件", "有効期限", "状態", "発注先", "公開範囲", "税区分"];
  if (sheet.getLastRow() === 0) { sheet.appendRow(headers); }
}

function checkAndFixOrderHeader(sheet) {
  if (!sheet) return;
  const headers = ["ID", "日付", "発注先", "関連見積ID", "工種", "品名", "仕様", "数量", "単位", "単価", "金額", "納品場所", "状態", "備考", "作成者", "公開範囲"];
  if (sheet.getLastRow() === 0) { sheet.appendRow(headers); }
}

function checkAndFixInvoiceHeader(sheet) {
  if (!sheet) return;
  const headers = ["ID", "ステータス", "登録日時", "ファイルID", "工事ID", "工事名", "担当者", "着工日", "請求金額", "相殺額", "支払予定額", "内容", "備考", "登録番号", "施工者", "工事場所"];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    // 既存シートのヘッダー行を自動修正
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (currentHeaders.length < headers.length || String(currentHeaders[6]) !== '担当者') {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
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
    const vData = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, Math.min(vSheet.getLastColumn(), 5)).getValues();
    const map = new Map();
    vData.forEach(r => {
      const name = String(r[1]).trim();
      if (!name) return;
      const display = r[2] ? `${name} ${r[2]}` : name;
      map.set(display, { name, honorific: r[2]||'', displayName: display, account: r[3]||'', initial: r[4] ? String(r[4]).trim() : '' });
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
    const numRows = lastRow - startRow + 1;
    const data = lSheet.getRange(startRow, 1, numRows, 10).getValues();
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

// -----------------------------------------------------------
// 材料費単価マスタ
// -----------------------------------------------------------

function checkAndFixMaterialHeader(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["品名", "仕様", "単位", "単価", "更新日"]);
  }
}

function getMaterialSheetName_(sheetType) {
  return sheetType === 'incl' ? CONFIG.sheetNames.masterMaterialIncl : CONFIG.sheetNames.masterMaterial;
}

function getMaterialCacheKey_(sheetType) {
  return sheetType === 'incl' ? 'material_incl_prices' : 'material_prices';
}

function getOrCreateMaterialSheet_(sheetType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getMaterialSheetName_(sheetType);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    checkAndFixMaterialHeader(sheet);
  }
  if (sheet.getLastRow() === 0) checkAndFixMaterialHeader(sheet);
  return sheet;
}

function apiGetMaterialPrices(sheetType) {
  const cacheKey = getMaterialCacheKey_(sheetType);
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const sheet = getOrCreateMaterialSheet_(sheetType);
  if (sheet.getLastRow() <= 1) return JSON.stringify([]);

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getDisplayValues();
  const items = data.filter(r => r[0]).map(r => ({
    product: r[0],
    spec: r[1] || '',
    unit: r[2] || '',
    price: parseCurrency(r[3]),
    updatedAt: r[4] || ''
  }));

  const result = JSON.stringify(items);
  try {
    cache.put(cacheKey, result, CACHE_TTL);
  } catch (e) {
    console.warn("Cache put failed (" + cacheKey + "): " + e.message);
  }
  return result;
}

function apiUpsertMaterialPrices(jsonItems, sheetType) {
  const items = JSON.parse(jsonItems);
  if (!items || !items.length) return JSON.stringify({ success: true, updated: 0, added: 0 });

  const sheet = getOrCreateMaterialSheet_(sheetType);

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
  let updated = 0, added = 0;

  // 既存データ読込
  const lastRow = sheet.getLastRow();
  let existing = [];
  if (lastRow > 1) {
    existing = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  }

  // キーマップ作成 (品名+仕様 → 行インデックス)
  const keyMap = new Map();
  existing.forEach((r, i) => {
    const key = (String(r[0]).trim() + "\t" + String(r[1]).trim()).toLowerCase();
    keyMap.set(key, i);
  });

  const newRows = [];
  items.forEach(item => {
    const product = String(item.product || '').trim();
    if (!product) return;
    const spec = String(item.spec || '').trim();
    const unit = String(item.unit || '').trim();
    const price = Number(item.price) || 0;
    const key = (product + "\t" + spec).toLowerCase();

    if (keyMap.has(key)) {
      // 上書き更新
      const rowIdx = keyMap.get(key);
      existing[rowIdx] = [product, spec, unit, price, now];
      updated++;
    } else {
      // 新規追加
      newRows.push([product, spec, unit, price, now]);
      keyMap.set(key, existing.length + newRows.length - 1);
      added++;
    }
  });

  // 既存行を一括書き戻し
  if (existing.length > 0) {
    sheet.getRange(2, 1, existing.length, 5).setValues(existing);
  }
  // 新規行を追加
  if (newRows.length > 0) {
    sheet.getRange(lastRow + 1, 1, newRows.length, 5).setValues(newRows);
  }

  invalidateDataCache_();
  return JSON.stringify({ success: true, updated: updated, added: added });
}

function apiDeleteMaterialPrice(product, spec, sheetType) {
  const sheet = getOrCreateMaterialSheet_(sheetType);
  if (sheet.getLastRow() <= 1) return JSON.stringify({ success: false, message: "データなし" });

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const targetProduct = String(product || '').trim().toLowerCase();
  const targetSpec = String(spec || '').trim().toLowerCase();

  for (let i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim().toLowerCase() === targetProduct &&
        String(data[i][1]).trim().toLowerCase() === targetSpec) {
      sheet.deleteRow(i + 2);
      invalidateDataCache_();
      return JSON.stringify({ success: true });
    }
  }
  return JSON.stringify({ success: false, message: "該当データが見つかりません" });
}

// ==================== マスタデータ編集API ====================

const MASTER_TYPE_MAP = {
  basic: CONFIG.sheetNames.masterBasic,
  client: CONFIG.sheetNames.masterClient,
  set: CONFIG.sheetNames.masterSet,
  vendor: CONFIG.sheetNames.masterVendor,
  material: CONFIG.sheetNames.masterMaterial,
  material_incl: CONFIG.sheetNames.masterMaterialIncl,
  journal: CONFIG.sheetNames.journalConfig
};

function getMasterSheet_(masterType) {
  const sheetName = MASTER_TYPE_MAP[masterType];
  if (!sheetName) throw new Error('不明なマスタ種別: ' + masterType);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('シートが見つかりません: ' + sheetName);
  return sheet;
}

function apiGetMasterData(masterType) {
  try {
    const sheet = getMasterSheet_(masterType);
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length === 0) return JSON.stringify({ headers: [], data: [] });
    const headers = data[0];
    const rows = data.slice(1);
    Logger.log('apiGetMasterData: ' + masterType + ' のデータ取得完了 (' + rows.length + '件)');
    return JSON.stringify({ headers: headers, data: rows });
  } catch (e) {
    Logger.log('apiGetMasterData: エラー - ' + e.message);
    return JSON.stringify({ error: e.message });
  }
}

function apiUpdateMasterRow(masterType, rowIndex, valuesJson) {
  try {
    const sheet = getMasterSheet_(masterType);
    const values = JSON.parse(valuesJson);
    const sheetRow = rowIndex + 2; // ヘッダー行分+1、0始まり→1始まり+1
    if (sheetRow < 2 || sheetRow > sheet.getLastRow()) throw new Error('行番号が範囲外です');
    sheet.getRange(sheetRow, 1, 1, values.length).setValues([values]);
    invalidateDataCache_();
    Logger.log('apiUpdateMasterRow: ' + masterType + ' の行' + rowIndex + 'を更新しました');
    return JSON.stringify({ success: true });
  } catch (e) {
    Logger.log('apiUpdateMasterRow: エラー - ' + e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
}

function apiAddMasterRow(masterType, valuesJson) {
  try {
    const sheet = getMasterSheet_(masterType);
    const values = JSON.parse(valuesJson);
    sheet.appendRow(values);
    invalidateDataCache_();
    Logger.log('apiAddMasterRow: ' + masterType + ' に行を追加しました');
    return JSON.stringify({ success: true });
  } catch (e) {
    Logger.log('apiAddMasterRow: エラー - ' + e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
}

function apiDeleteMasterRow(masterType, rowIndex) {
  try {
    const sheet = getMasterSheet_(masterType);
    const sheetRow = rowIndex + 2; // ヘッダー行分+1、0始まり→1始まり+1
    if (sheetRow < 2 || sheetRow > sheet.getLastRow()) throw new Error('行番号が範囲外です');
    sheet.deleteRow(sheetRow);
    invalidateDataCache_();
    Logger.log('apiDeleteMasterRow: ' + masterType + ' の行' + rowIndex + 'を削除しました');
    return JSON.stringify({ success: true });
  } catch (e) {
    Logger.log('apiDeleteMasterRow: エラー - ' + e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
}

function apiClearMaterialPrices(sheetType) {
  try {
    const sheetName = getMaterialSheetName_(sheetType);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return JSON.stringify({ success: true, deleted: 0 });
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return JSON.stringify({ success: true, deleted: 0 });
    const count = lastRow - 1;
    sheet.deleteRows(2, count);
    invalidateDataCache_();
    Logger.log('apiClearMaterialPrices: ' + sheetName + ' から' + count + '件のデータを削除しました');
    return JSON.stringify({ success: true, deleted: count });
  } catch (e) {
    Logger.log('apiClearMaterialPrices: エラー - ' + e.message);
    return JSON.stringify({ success: false, message: e.message });
  }
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
          id: currentId, date: formatDate(row[1]), updatedAt: row[1] ? new Date(row[1]).getTime() : 0, client: row[2], project: row[13], location: row[12], period: row[14] || '', payment: row[15] || '', expiry: row[16] || '',
          status: row[17] || "未作成", visibility: row[19] || 'public', taxMode: row[20] || '税別',
          totalAmount: 0, totalOrderAmount: summary.totalCost, orderCount: summary.orderCount,
          totalInvoicedAmount: invSummary.totalInvoiced, invoiceCount: invSummary.invoiceCount,
          totalDeposit: depSummary.totalDeposit, depositCount: depSummary.depositCount
        }; 
      }
      projectMap[currentId].totalAmount += Number(row[10]) || 0;
    }
  });
  const result = JSON.stringify(Object.values(projectMap).sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0)));
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
        draftsMap.set(id, { id: id, date: formatDate(row[1]), timestamp: new Date(row[1] || 0).getTime(), client: row[2], project: row[13], location: row[12] || '', period: row[14] || '', payment: row[15] || '', expiry: row[16] || '', status: row[17], taxMode: row[20] || '税別', totalAmount: 0 });
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
        header = { id: rowId, date: formatDate(row[1]), client: row[2], location: row[12], project: row[13], period: row[14], payment: row[15], expiry: row[16], status: row[17], remarks: row[11], visibility: row[19] || 'public', taxMode: row[20] || '税別' };
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
    if (!estSheet) { estSheet = ss.insertSheet(CONFIG.sheetNames.list); checkAndFixEstimateHeader(estSheet); }
    
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
        isFirst ? (estimateData.header.visibility || 'public') : "",
        isFirst ? (estimateData.header.taxMode || '税別') : ""
      ];
    });

    estSheet.getRange(estSheet.getLastRow() + 1, 1, estValues.length, 21).setValues(estValues.map(r => { while(r.length < 21) r.push(""); return r; }));

    // 見積保存時の自動発注生成: 発注先が記載された明細を発注先ごとにグループ化して保存
    const vendorGroups = {};
    estItems.forEach(item => {
      const v = String(item.vendor || '').trim();
      if (!v || !String(item.product || '').trim()) return;
      if (!vendorGroups[v]) vendorGroups[v] = [];
      vendorGroups[v].push(item);
    });

    if (Object.keys(vendorGroups).length > 0) {
      let orderSheet = ss.getSheetByName(CONFIG.sheetNames.order);
      if (!orderSheet) { orderSheet = ss.insertSheet(CONFIG.sheetNames.order); checkAndFixOrderHeader(orderSheet); }

      // この見積に紐づく既存の自動生成発注を削除
      deleteOrdersByEstimateId_(orderSheet, saveId);

      const email = Session.getActiveUser().getEmail();
      Object.keys(vendorGroups).forEach(vendor => {
        const items = vendorGroups[vendor];
        const orderId = getNextSequenceId('order');
        const orderValues = items.map((item, idx) => {
          const qty = Number(item.qty) || 0;
          const cost = Number(item.cost) || 0;
          return [
            idx === 0 ? orderId : "",
            saveTimestamp,
            vendor,
            saveId,
            item.category || "",
            item.product || "",
            item.spec || "",
            qty,
            item.unit || "",
            cost,
            Math.round(qty * cost),
            estimateData.header.location || "",
            "発注書作成",
            "",
            email,
            "public"
          ];
        });
        const padded = orderValues.map(r => { while (r.length < 16) r.push(""); return r; });
        orderSheet.getRange(orderSheet.getLastRow() + 1, 1, orderValues.length, 16).setValues(padded);
      });
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
  // ※キャッシュなし。リクエスト毎にDrive APIを呼び出すため、大量PDFの場合は要最適化
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
        hasPdf: false, project: '', client: '', period: '', payment: '', expiry: '', taxMode: ''
      });
    }
    const amount = parseCurrency(row[IDX.amount]);
    const currentData = orderMap.get(currentId);
    if (currentData) { currentData.totalAmount += amount; }
  }

  // PDF存在チェック & プロジェクト名取得
  // 見積ヘッダーMapを事前構築 (N+1回避)
  const estHeaderMap = new Map();
  const estListSheet = ss.getSheetByName(CONFIG.sheetNames.list);
  if (estListSheet && estListSheet.getLastRow() > 1) {
    const estData = estListSheet.getDataRange().getValues();
    for (let i = 1; i < estData.length; i++) {
      const eid = String(estData[i][0]);
      if (eid && !estHeaderMap.has(eid)) {
        estHeaderMap.set(eid, { project: estData[i][13], client: estData[i][2], location: estData[i][12], period: estData[i][14] || '', payment: estData[i][15] || '', expiry: estData[i][16] || '', taxMode: estData[i][20] || '税別' });
      }
    }
  }
  const list = Array.from(orderMap.values());
  list.forEach(order => {
    // PDF存在チェック: 発注書_業者名_ でファイル名マッチ
    const cleanVendor = (order.vendor || '').replace(/[\r\n\t\\/:*?"<>|]/g, '').trim();
    order.hasPdf = pdfFileNames.some(fn => fn.includes('発注書_' + cleanVendor));
    // 関連見積IDからプロジェクト名を推定 (Map参照)
    if (order.relEstId) {
      const est = estHeaderMap.get(order.relEstId);
      if (est) {
        order.project = est.project || '';
        order.client = est.client || '';
        order.period = est.period || '';
        order.payment = est.payment || '';
        order.expiry = est.expiry || '';
        order.taxMode = est.taxMode || '';
        if (!order.location) order.location = est.location || '';
      }
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
  // 発注先マスタから敬称を取得
  data.header.honorific = " 御中";
  if (targetVendor) {
    try {
      const vSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetNames.masterVendor);
      if (vSheet && vSheet.getLastRow() > 1) {
        const vData = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 3).getValues();
        for (const r of vData) {
          if (String(r[1]).trim() === targetVendor.replace(/\s*(御中|様|殿)\s*$/, "").trim()) {
            data.header.honorific = r[2] ? ` ${r[2]}` : " 御中";
            break;
          }
        }
      }
    } catch(e) { /* fallback to 御中 */ }
  }
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
      amount: parseCurrency(row[col['金額']]) || 0,
      vendor: row[col['発注先']] || ''
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
  const result = { constructionId: "", project: "", person: "", contractor: "", amount: 0, content: "", date: "", location: "", initial: "", detectedKanryo: false };
  const keyMap = [
    { key: "constructionId", keywords: ["工事番号", "工事ID", "No"] },
    { key: "project", keywords: ["現場名", "工事名", "案件名", "件名"] },
    { key: "person", keywords: ["担当者", "担当", "請求業者", "業者名", "請求元", "会社名", "企業名"] },
    { key: "contractor", keywords: ["施工者", "施工業者", "施工担当", "作業者"] },
    { key: "amount", keywords: ["金額", "請求金額", "合計", "税込金額"] },
    { key: "content", keywords: ["内容", "但し書き", "品名", "詳細"] },
    { key: "date", keywords: ["日付", "着工日", "請求日", "発行日", "開始日"] },
    { key: "location", keywords: ["工事場所", "現場住所", "住所", "場所"] },
    { key: "initial", keywords: ["イニシャル", "頭文字"] }
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
           const regexColon = new RegExp(`^${keyword}\\s*[:：、,]\\s*(.*)$`);
           const matchColon = l.match(regexColon);
           if (matchColon) value = matchColon[1].trim();
        }
        if (value) {
          if (map.key === "amount") { result[map.key] = parseCurrency(value); } else { result[map.key] = value; }
        }
      });
    });
  });
  // 完了キーワード検出
  if (content.includes('完了')) {
    result.detectedKanryo = true;
  }
  return result;
}

function _parseInvoiceImageWithGemini(file) {
  if (!CONFIG.API_KEY) return { error: "APIキーなし" };
  const projectsJson = apiGetActiveProjectsList();
  const projects = JSON.parse(projectsJson).map(p => `${p.id}: ${p.name}`).join("\n");
  const mime = file.getMimeType();
  const base64 = Utilities.base64Encode(file.getBlob().getBytes());
  const prompt = `あなたは建築積算のプロです。画像から情報を抽出してください。\n【重要】以下のリストを参照し、最も関連性が高い「工事番号(constructionId)」を推測してください。\nリスト: ${projects}\n抽出項目: constructionId, person(担当者/請求元), contractor(施工者/施工業者), date(着工日/開始日, yyyy/MM/dd), amount(税込), content, registrationNumber(Tから始まる13桁の番号), location(工事場所), initial(担当者のイニシャル/頭文字, A-Zの1文字), detectedKanryo(文書に「完了」という文言があればtrue)`;
  const parts = [{ text: prompt }, { inline_data: { mime_type: mime, data: base64 } }];
  const responseSchema = {
    "type": "OBJECT",
    "properties": {
      "constructionId": { "type": "STRING" }, "person": { "type": "STRING" }, "contractor": { "type": "STRING" },
      "date": { "type": "STRING" }, "amount": { "type": "NUMBER" }, "content": { "type": "STRING" },
      "registrationNumber": { "type": "STRING", "description": "T+13 digits" },
      "location": { "type": "STRING" }, "initial": { "type": "STRING" },
      "detectedKanryo": { "type": "BOOLEAN" }
    }
  };
  const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${CONFIG.API_KEY}`, {
    method: "post", contentType: "application/json",
    payload: JSON.stringify({ contents: [{ parts }], generationConfig: { response_mime_type: "application/json", response_schema: responseSchema } }),
    muteHttpExceptions: true
  });
  const json = JSON.parse(res.getContentText());
  if (json.error) throw new Error("Gemini API Error: " + json.error.message);
  if (!json.candidates || !json.candidates[0]) throw new Error("AIから回答を取得できませんでした");
  return JSON.parse(json.candidates[0].content.parts[0].text);
}

// ── 請求書OCR (ブラウザアップロード) ──────────────────
function apiParseInvoiceFromBase64(base64Data, mimeType) {
  try {
    if (!base64Data || !mimeType) return JSON.stringify({ error: "データが不足しています" });
    if (!CONFIG.API_KEY) return JSON.stringify({ error: "APIキーが未設定です" });
    const projectsJson = apiGetActiveProjectsList();
    const projects = JSON.parse(projectsJson).map(p => `${p.id}: ${p.name}`).join("\n");
    const prompt = `あなたは建築積算のプロです。画像から情報を抽出してください。\n【重要】以下のリストを参照し、最も関連性が高い「工事番号(constructionId)」を推測してください。\nリスト: ${projects}\n抽出項目: constructionId, person(担当者/請求元), contractor(施工者/施工業者), date(着工日/開始日, yyyy/MM/dd), amount(税込), content, registrationNumber(Tから始まる13桁の番号), location(工事場所), initial(担当者のイニシャル/頭文字, A-Zの1文字), detectedKanryo(文書に「完了」という文言があればtrue)`;
    const parts = [{ text: prompt }, { inline_data: { mime_type: mimeType, data: base64Data } }];
    const responseSchema = {
      "type": "OBJECT",
      "properties": {
        "constructionId": { "type": "STRING" }, "person": { "type": "STRING" }, "contractor": { "type": "STRING" },
        "date": { "type": "STRING" }, "amount": { "type": "NUMBER" }, "content": { "type": "STRING" },
        "registrationNumber": { "type": "STRING", "description": "T+13 digits" },
        "location": { "type": "STRING" }, "initial": { "type": "STRING" },
        "detectedKanryo": { "type": "BOOLEAN" }
      }
    };
    const res = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${CONFIG.API_KEY}`, {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({ contents: [{ parts }], generationConfig: { response_mime_type: "application/json", response_schema: responseSchema } }),
      muteHttpExceptions: true
    });
    const json = JSON.parse(res.getContentText());
    if (json.error) throw new Error("Gemini API Error: " + json.error.message);
    if (!json.candidates || !json.candidates[0]) throw new Error("AIから回答を取得できませんでした");
    return JSON.stringify(JSON.parse(json.candidates[0].content.parts[0].text));
  } catch (e) { return JSON.stringify({ error: e.toString() }); }
}

// ── 見積OCR機能 ──────────────────────────────────
function apiOcrEstimateFromDrive(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const mime = file.getMimeType();
    if (mime.includes("text") || file.getName().endsWith(".txt")) {
      return JSON.stringify(_parseTextEstimateItems(file));
    }
    return JSON.stringify(_parseEstimateImageWithGemini(
      Utilities.base64Encode(file.getBlob().getBytes()), mime
    ));
  } catch (e) { return JSON.stringify({ error: e.toString() }); }
}

function apiOcrEstimateFromBase64(base64Data, mimeType) {
  try {
    if (!base64Data || !mimeType) return JSON.stringify({ error: "データが不足しています" });
    return JSON.stringify(_parseEstimateImageWithGemini(base64Data, mimeType));
  } catch (e) { return JSON.stringify({ error: e.toString() }); }
}

function _parseEstimateImageWithGemini(base64, mime) {
  if (!CONFIG.API_KEY) return { error: "APIキーが未設定です" };
  const prompt = `あなたは建築積算の専門家です。この画像/PDFから見積書・請求書・納品書の明細行を読み取り、すべての品目を抽出してください。
【抽出ルール】
- 各行の「工種(カテゴリ)」「品名」「仕様・規格」「数量」「単位」「単価」「金額」を読み取る
- 数量が空欄の場合は 0 とする（後で金額から逆算する）
- 単位が空欄の場合は "式" とする
- 単価が不明な場合は 0 とする
- 金額が記載されていれば必ず読み取る。金額が空欄の場合は 0 とする
- ヘッダ行・合計行・消費税行は除外する
- 工種/カテゴリが明記されていない場合は空文字列にする`;
  const parts = [{ text: prompt }, { inline_data: { mime_type: mime, data: base64 } }];
  const responseSchema = {
    "type": "OBJECT",
    "properties": {
      "items": {
        "type": "ARRAY",
        "items": {
          "type": "OBJECT",
          "properties": {
            "category": { "type": "STRING", "description": "工種・カテゴリ" },
            "product":  { "type": "STRING", "description": "品名・品目名" },
            "spec":     { "type": "STRING", "description": "仕様・規格・サイズ" },
            "qty":      { "type": "NUMBER", "description": "数量" },
            "unit":     { "type": "STRING", "description": "単位(式, m, m2, 枚, 個 等)" },
            "price":    { "type": "NUMBER", "description": "単価(円)" },
            "amount":   { "type": "NUMBER", "description": "金額(円)" }
          },
          "required": ["product"]
        }
      }
    },
    "required": ["items"]
  };
  const res = UrlFetchApp.fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${CONFIG.API_KEY}`,
    { method: "post", contentType: "application/json",
      payload: JSON.stringify({ contents: [{ parts }], generationConfig: { response_mime_type: "application/json", response_schema: responseSchema } }),
      muteHttpExceptions: true }
  );
  const json = JSON.parse(res.getContentText());
  if (json.error) return { error: "Gemini API Error: " + json.error.message };
  if (!json.candidates || !json.candidates[0]) return { error: "AIから回答を取得できませんでした" };
  const parsed = JSON.parse(json.candidates[0].content.parts[0].text);
  return { items: (parsed.items || []).map(item => {
    let qty = Number(item.qty) || 0;
    let price = Number(item.price) || 0;
    let amount = Number(item.amount) || 0;
    // 金額があるが数量・単価が不完全な場合の逆算
    if (amount > 0) {
      if (price > 0 && qty === 0) {
        // 金額と単価から数量を逆算
        qty = Math.round((amount / price) * 100) / 100;
      } else if (qty > 0 && price === 0) {
        // 金額と数量から単価を逆算
        price = Math.round(amount / qty);
      } else if (qty === 0 && price === 0) {
        // 金額のみ: 数量=1, 単価=金額
        qty = 1; price = amount;
      }
    }
    // 数量が0のままなら1にフォールバック
    if (qty === 0) qty = 1;
    // 金額の整合性チェック: OCR金額が無い場合は計算で補完
    if (amount === 0 && price > 0) amount = Math.round(qty * price);
    return {
      category: item.category || "", product: item.product || "", spec: item.spec || "",
      qty: qty, unit: item.unit || "式", price: price, amount: amount
    };
  })};
}

function _parseTextEstimateItems(file) {
  let content = "";
  try {
    content = file.getBlob().getDataAsString();
    if (!content.match(/品名|工種|数量|単価|仕様/)) content = file.getBlob().getDataAsString('Shift_JIS');
  } catch(e) { return { error: "テキストの読み取りに失敗しました" }; }
  const lines = content.split(/\r\n|\n/).filter(l => l.trim());
  const items = [];
  lines.forEach(line => {
    const cols = line.split(/\t/);
    if (cols.length >= 2) {
      let qty = parseCurrency(cols[3]) || 0;
      let price = parseCurrency(cols[5]) || 0;
      let amount = parseCurrency(cols[6]) || 0;
      if (amount > 0) {
        if (price > 0 && qty === 0) qty = Math.round((amount / price) * 100) / 100;
        else if (qty > 0 && price === 0) price = Math.round(amount / qty);
        else if (qty === 0 && price === 0) { qty = 1; price = amount; }
      }
      if (qty === 0) qty = 1;
      if (amount === 0 && price > 0) amount = Math.round(qty * price);
      items.push({ category: cols[0] || "", product: cols[1] || "", spec: cols[2] || "",
        qty: qty, unit: cols[4] || "式", price: price, amount: amount });
    }
  });
  return { items };
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
    // 工事番号の自動解決
    const constructionId = resolveConstructionId_(data);
    const rowValues = [
        id, data.status || "着工", now, data.fileId || "", constructionId,
        data.project || "", data.person || "", data.date || "", data.amount || 0,
        data.offset || 0, payment, data.content || "", data.remarks || "", data.registrationNumber || "",
        data.contractor || "", data.location || ""
    ];
    if (rowIndex > 0) {
        // 既存レコードの更新時はステータスを維持する (ユーザーが手動変更した値を上書きしないため)
        const currentStatus = sheet.getRange(rowIndex, 2).getValue();
        rowValues[1] = currentStatus;
        sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
        sheet.appendRow(rowValues);
    }
    // 完了自動処理
    let autoCompleteResult = null;
    if (data.detectedKanryo && data.project && data.person) {
      autoCompleteResult = apiAutoCompleteOnKanryo_(data.project, data.person, data.contractor || '');
    }
    invalidateDataCache_();
    return JSON.stringify({ success: true, constructionId: constructionId, autoComplete: autoCompleteResult });
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
        project: row[5], person: row[6], date: row[7],
        amount: parseCurrency(row[8]), offset: parseCurrency(row[9]), payment: parseCurrency(row[10]),
        content: row[11], remarks: row[12], registrationNumber: row[13] || "",
        contractor: row[14] || "", location: row[15] || ""
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

function apiGetOrderBalance(constructionId, personName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oSheet = ss.getSheetByName(CONFIG.sheetNames.order);
  const iSheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
  if (!oSheet) return JSON.stringify({ error: "発注シートがありません" });
  if (!personName) return JSON.stringify({ totalOrder: 0, totalPaid: 0, balance: 0 });
  const normSupplier = personName.replace(/[\s\u3000]/g, "");
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

    id = getNextDepositId_();
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

    id = getNextPaymentId_();
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
            if (row[1] !== '着工' && row[1] !== '完了') continue;
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
    if (json.error) return JSON.stringify({ error: "Gemini API Error: " + json.error.message });
    if (!json.candidates || !json.candidates[0]) return JSON.stringify({ error: "AIから回答を取得できませんでした" });
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
  // 発注ID→関連見積IDのMapを事前構築 (N+1回避)
  const orderIdToEstIdMap = new Map();
  if (ordSheet && ordSheet.getLastRow() > 1) {
    const oData2 = ordSheet.getDataRange().getDisplayValues();
    for (let j = 1; j < oData2.length; j++) {
      if (oData2[j][0]) orderIdToEstIdMap.set(oData2[j][0], oData2[j][3] || '');
    }
  }
  const paySheet = ss.getSheetByName(CONFIG.sheetNames.payments);
  const paymentEntries = [];
  let totalPaymentAmount = 0;
  if (paySheet && paySheet.getLastRow() > 1) {
    const pData = paySheet.getDataRange().getDisplayValues();
    for (let i = 1; i < pData.length; i++) {
      const r = pData[i];
      if (!r[0]) continue;
      const orderId = String(r[3]).trim();
      const status = String(r[12]).trim();
      if (status === '取消') continue;
      const amount = parseCurrency(r[8]);
      const fee = parseCurrency(r[9]);
      // 工事名でも照合
      const payProject = String(r[6]).trim();
      const payEstHeader = estimate.header || {};
      const matchByProject = payProject && (payProject === (payEstHeader.project || '') || payProject === (payEstHeader.client || ''));

      // 関連発注IDで照合: Map参照で関連見積IDを取得
      let matchByOrderEstId = false;
      if (orderId) {
        const relEstId = orderIdToEstIdMap.get(orderId) || '';
        matchByOrderEstId = (relEstId === projectId || relEstId.startsWith(projectId + '-'));
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

/**
 * 指定IDのデータを全シート(見積・発注・請求書・入金・出金)から削除する。
 * フロントエンドからの汎用削除APIとして使用。
 * @param {string} id - 削除対象のレコードID
 * @returns {string} JSON { success: boolean, message: string }
 */
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
// ==========================================
// リマインド機能
// ==========================================

/**
 * 決済条件テキストから支払期日を算出
 * @param {string} paymentText - 決済条件テキスト（例: "月末締め翌月末払い"）
 * @param {Date} baseDate - 起算日
 * @return {Date|null} 支払期日。パース不可の場合はnull
 */
function parsePaymentTerms_(paymentText, baseDate) {
  if (!paymentText || !baseDate) return null;
  const text = String(paymentText).replace(/\s+/g, '').trim();
  if (!text) return null;

  const d = new Date(baseDate);
  if (isNaN(d.getTime())) return null;

  // 即日 / 現金
  if (/^(即日|現金|即金|cash)$/i.test(text)) return new Date(d);

  // 受領後○日 / NET○
  const netMatch = text.match(/(?:受領後|NET|net)(\d+)日?/);
  if (netMatch) {
    const result = new Date(d);
    result.setDate(result.getDate() + parseInt(netMatch[1], 10));
    return result;
  }

  // ○日締め翌月○日払い / ○日締め翌々月○日払い
  const customMatch = text.match(/(\d+)日締め(翌々?月)(\d+)日払/);
  if (customMatch) {
    const closeDay = parseInt(customMatch[1], 10);
    const monthsAhead = customMatch[2] === '翌々月' ? 2 : 1;
    const payDay = parseInt(customMatch[3], 10);
    // 締め日を基準に月を進める
    let closeDate = new Date(d.getFullYear(), d.getMonth(), closeDay);
    if (d.getDate() > closeDay) closeDate.setMonth(closeDate.getMonth() + 1);
    const result = new Date(closeDate.getFullYear(), closeDate.getMonth() + monthsAhead, payDay);
    return result;
  }

  // 月末締め翌月○日払い
  const endMonthDayMatch = text.match(/月末締め(翌々?月)(\d+)日払/);
  if (endMonthDayMatch) {
    const monthsAhead = endMonthDayMatch[1] === '翌々月' ? 2 : 1;
    const payDay = parseInt(endMonthDayMatch[2], 10);
    const endOfMonth = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    const result = new Date(endOfMonth.getFullYear(), endOfMonth.getMonth() + monthsAhead, payDay);
    return result;
  }

  // 月末締め翌月末払い / 月末締め翌々月末払い
  const endMonthMatch = text.match(/月末締め(翌々?月)末払/);
  if (endMonthMatch) {
    const monthsAhead = endMonthMatch[1] === '翌々月' ? 2 : 1;
    const endOfMonth = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    const result = new Date(endOfMonth.getFullYear(), endOfMonth.getMonth() + monthsAhead + 1, 0);
    return result;
  }

  return null;
}

/**
 * リマインド一覧を取得
 * @return {string} JSON配列
 */
function apiGetReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reminders = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // --- PDF存在チェック用: ドライブフォルダ内のPDFファイル名を収集 ---
  let pdfFileNames = [];
  try {
    const folder = getSaveFolder();
    const pdfFiles = folder.getFilesByType(MimeType.PDF);
    while (pdfFiles.hasNext()) pdfFileNames.push(pdfFiles.next().getName());
  } catch (e) { /* ignore */ }

  // --- 見積リスト読み込み ---
  const listSheet = ss.getSheetByName(CONFIG.sheetNames.list);
  const estimateMap = {}; // id -> { client, project, status, date, payment, totalAmount }
  if (listSheet && listSheet.getLastRow() > 1) {
    const listData = listSheet.getDataRange().getValues().slice(1);
    let currentId = '', currentHeader = null;
    listData.forEach(row => {
      const id = String(row[0]);
      if (id) {
        currentId = id;
        currentHeader = {
          client: String(row[2] || ''),
          project: String(row[13] || ''),
          status: String(row[17] || ''),
          date: row[1],
          payment: String(row[15] || ''),
          totalAmount: 0
        };
        estimateMap[currentId] = currentHeader;
      }
      if (currentId && estimateMap[currentId]) {
        estimateMap[currentId].totalAmount += Number(row[10]) || 0;
      }
    });
  }

  // --- A. PDF未作成リマインド (見積書) ---
  Object.keys(estimateMap).forEach(id => {
    const est = estimateMap[id];
    const cleanClient = est.client.replace(/[\r\n\t\\/:*?"<>|]/g, '').trim();
    if (est.status === '見積提出' && cleanClient) {
      const hasQuotePdf = pdfFileNames.some(fn => fn.includes('御見積書_' + cleanClient));
      if (!hasQuotePdf) {
        reminders.push({
          type: 'pdf_missing', docType: '見積書', id: id,
          label: est.client + ' / ' + est.project,
          severity: 'info'
        });
      }
    }
    // PDF未作成リマインド (請求書)
    if (est.status === '請求済' && cleanClient) {
      const hasBillPdf = pdfFileNames.some(fn => fn.includes('御請求書_' + cleanClient));
      if (!hasBillPdf) {
        reminders.push({
          type: 'pdf_missing', docType: '請求書', id: id,
          label: est.client + ' / ' + est.project,
          severity: 'info'
        });
      }
    }
  });

  // --- A. PDF未作成リマインド (発注書) ---
  const orderSheet = ss.getSheetByName(CONFIG.sheetNames.order);
  if (orderSheet && orderSheet.getLastRow() > 1) {
    const oData = orderSheet.getDataRange().getDisplayValues();
    let hIdx = 0;
    for (let i = 0; i < Math.min(10, oData.length); i++) { if (oData[i][0] === 'ID') { hIdx = i; break; } }
    const h = oData[hIdx];
    const col = {}; h.forEach((v, i) => col[String(v).trim()] = i);
    const idxVendor = col['発注先'] !== undefined ? col['発注先'] : 2;
    const idxStatus = col['状態'] !== undefined ? col['状態'] : 12;
    const idxRelEstId = col['関連見積ID'] !== undefined ? col['関連見積ID'] : 3;

    const seenOrders = new Set();
    for (let i = hIdx + 1; i < oData.length; i++) {
      const row = oData[i];
      const orderId = row[0];
      if (orderId && !seenOrders.has(orderId)) {
        seenOrders.add(orderId);
        const vendor = String(row[idxVendor] || '');
        const status = String(row[idxStatus] || '');
        const cleanVendor = vendor.replace(/[\r\n\t\\/:*?"<>|]/g, '').trim();
        if (status === '発注書作成' && cleanVendor) {
          const hasPdf = pdfFileNames.some(fn => fn.includes('発注書_' + cleanVendor));
          if (!hasPdf) {
            const relEstId = row[idxRelEstId];
            const est = relEstId ? estimateMap[relEstId] : null;
            reminders.push({
              type: 'pdf_missing', docType: '発注書', id: orderId,
              label: vendor + (est ? ' / ' + est.project : ''),
              severity: 'info'
            });
          }
        }
      }
    }
  }

  // --- 入金データ集計 (見積ID単位) ---
  const depositByEstimate = {};
  const depSheet = ss.getSheetByName(CONFIG.sheetNames.deposits);
  if (depSheet && depSheet.getLastRow() > 1) {
    const dData = depSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < dData.length; i++) {
      const row = dData[i];
      if (!row[0]) continue;
      const estId = String(row[3]).trim();
      if (!estId) continue;
      const status = String(row[11]).trim();
      if (status === '取消') continue;
      const amount = parseCurrency(row[7]);
      if (!depositByEstimate[estId]) depositByEstimate[estId] = 0;
      depositByEstimate[estId] += amount;
    }
  }

  // --- B. 入金期日リマインド ---
  Object.keys(estimateMap).forEach(id => {
    const est = estimateMap[id];
    if (est.status !== '請求済') return;
    if (!est.payment) return;
    const baseDate = est.date ? new Date(est.date) : null;
    const dueDate = parsePaymentTerms_(est.payment, baseDate);
    if (!dueDate) return;
    dueDate.setHours(0, 0, 0, 0);

    const depositTotal = depositByEstimate[id] || 0;
    const requiredAmount = Math.floor(est.totalAmount * 1.1);
    if (depositTotal >= requiredAmount) return; // 入金済み

    const diffDays = Math.round((dueDate - today) / (1000 * 60 * 60 * 24));
    if (diffDays > 3) return; // まだ余裕あり

    reminders.push({
      type: diffDays < 0 ? 'deposit_overdue' : 'deposit_due',
      id: id,
      label: est.client + ' / ' + est.project,
      dueDate: formatDate(dueDate),
      daysLeft: diffDays,
      severity: diffDays < 0 ? 'danger' : 'warning'
    });
  });

  // --- 出金データ集計 (請求書ID単位) ---
  const paymentByInvoice = {};
  const paySheet = ss.getSheetByName(CONFIG.sheetNames.payments);
  if (paySheet && paySheet.getLastRow() > 1) {
    const pData = paySheet.getDataRange().getDisplayValues();
    for (let i = 1; i < pData.length; i++) {
      const row = pData[i];
      if (!row[0]) continue;
      const invoiceId = String(row[4]).trim(); // 関連請求書ID
      if (!invoiceId) continue;
      const status = String(row[12]).trim();
      if (status === '取消') continue;
      const amount = parseCurrency(row[8]);
      if (!paymentByInvoice[invoiceId]) paymentByInvoice[invoiceId] = 0;
      paymentByInvoice[invoiceId] += amount;
    }
  }

  // --- C. 出金期日リマインド ---
  const invSheet = ss.getSheetByName(CONFIG.sheetNames.invoice);
  if (invSheet && invSheet.getLastRow() > 1) {
    const invData = invSheet.getDataRange().getValues();
    for (let i = 1; i < invData.length; i++) {
      const row = invData[i];
      const invoiceId = String(row[0]);
      if (!invoiceId) continue;
      const invStatus = String(row[1] || '');
      if (invStatus === '取消' || invStatus === '完了') continue;

      const plannedAmount = parseCurrency(row[10]); // 支払予定額
      if (plannedAmount <= 0) continue;
      const paidTotal = paymentByInvoice[invoiceId] || 0;
      if (paidTotal >= plannedAmount) continue; // 支払済み

      // 決済条件: 関連見積IDから取得
      const constId = String(row[4] || '');
      const est = constId ? estimateMap[constId] : null;
      const paymentTerms = est ? est.payment : '';
      const invoiceDate = row[7] ? new Date(row[7]) : null;

      let dueDate = null;
      if (paymentTerms && invoiceDate) {
        dueDate = parsePaymentTerms_(paymentTerms, invoiceDate);
      }
      // フォールバック: 請求日+30日
      if (!dueDate && invoiceDate && !isNaN(invoiceDate.getTime())) {
        dueDate = new Date(invoiceDate);
        dueDate.setDate(dueDate.getDate() + 30);
      }
      if (!dueDate) continue;
      dueDate.setHours(0, 0, 0, 0);

      const diffDays = Math.round((dueDate - today) / (1000 * 60 * 60 * 24));
      if (diffDays > 3) continue;

      const supplier = String(row[6] || '');
      const projectName = String(row[5] || '');
      reminders.push({
        type: diffDays < 0 ? 'payment_overdue' : 'payment_due',
        id: invoiceId,
        label: supplier + ' / ' + projectName,
        dueDate: formatDate(dueDate),
        daysLeft: diffDays,
        severity: diffDays < 0 ? 'danger' : 'warning'
      });
    }
  }

  // severity順にソート: danger > warning > info
  const severityOrder = { danger: 0, warning: 1, info: 2 };
  reminders.sort((a, b) => (severityOrder[a.severity] || 9) - (severityOrder[b.severity] || 9));

  return JSON.stringify(reminders);
}

/**
 * リマインドメール送信（トリガーから呼び出し）
 */
function sendReminderEmail_() {
  try {
    const remindersJson = apiGetReminders();
    const reminders = JSON.parse(remindersJson);
    const urgent = reminders.filter(r => r.severity === 'danger' || r.severity === 'warning');
    if (urgent.length === 0) return;

    const typeLabel = {
      deposit_overdue: '入金期日超過',
      deposit_due: '入金期日接近',
      payment_overdue: '出金期日超過',
      payment_due: '出金期日接近',
      pdf_missing: 'PDF未作成'
    };

    let body = 'AI建築見積システム - リマインド通知\n';
    body += '==========================================\n\n';
    urgent.forEach(r => {
      const icon = r.severity === 'danger' ? '[!!]' : '[!]';
      const type = typeLabel[r.type] || r.type;
      body += icon + ' ' + type + ': ' + r.label;
      if (r.dueDate) {
        body += ' (期日: ' + r.dueDate;
        if (r.daysLeft < 0) body += ', ' + Math.abs(r.daysLeft) + '日超過';
        else body += ', あと' + r.daysLeft + '日';
        body += ')';
      }
      body += '\n';
    });
    body += '\n--\nAI建築見積システム v10.0';

    // メール送信
    const email = Session.getActiveUser().getEmail();
    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: '[リマインド] 未処理項目が ' + urgent.length + ' 件あります',
        body: body
      });
    }

    // LINE送信
    sendReminderLine_(urgent, typeLabel);
  } catch (e) {
    console.error('sendReminderEmail_ failed: ' + e.toString());
  }
}

/**
 * LINE送信先ユーザーIDリストを取得（カンマ区切りで保存）
 */
function getLineUserIds_() {
  const raw = PropertiesService.getScriptProperties().getProperty('LINE_USER_IDS') || CONFIG.LINE_USER_ID || '';
  return raw.split(',').map(s => s.trim()).filter(s => s.length > 0);
}

/**
 * LINE Messaging APIでリマインド通知を送信（multicast対応）
 */
function sendReminderLine_(urgentItems, typeLabel) {
  const token = CONFIG.LINE_TOKEN;
  const userIds = getLineUserIds_();
  if (!token || userIds.length === 0) return;

  try {
    let text = '\u{1F514} リマインド通知\n\n';
    urgentItems.forEach(r => {
      const icon = r.severity === 'danger' ? '\u{1F534}' : '\u{1F7E1}';
      const type = typeLabel[r.type] || r.type;
      text += icon + ' ' + type + '\n  ' + r.label;
      if (r.dueDate) {
        text += '\n  期日: ' + r.dueDate;
        if (r.daysLeft < 0) text += '(' + Math.abs(r.daysLeft) + '日超過)';
        else text += '(あと' + r.daysLeft + '日)';
      }
      text += '\n\n';
    });

    if (userIds.length === 1) {
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post', contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: JSON.stringify({ to: userIds[0], messages: [{ type: 'text', text: text.trim() }] }),
        muteHttpExceptions: true
      });
    } else {
      UrlFetchApp.fetch('https://api.line.me/v2/bot/message/multicast', {
        method: 'post', contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: JSON.stringify({ to: userIds, messages: [{ type: 'text', text: text.trim() }] }),
        muteHttpExceptions: true
      });
    }
  } catch (e) {
    console.error('sendReminderLine_ failed: ' + e.toString());
  }
}

/**
 * LINE送信先を追加
 */
function apiAddLineUser(userId) {
  if (!userId || !userId.trim()) return JSON.stringify({ success: false, message: 'ユーザーIDが空です' });
  const props = PropertiesService.getScriptProperties();
  const existing = getLineUserIds_();
  const id = userId.trim();
  if (existing.includes(id)) return JSON.stringify({ success: false, message: 'このユーザーIDは既に登録されています' });
  existing.push(id);
  props.setProperty('LINE_USER_IDS', existing.join(','));
  return JSON.stringify({ success: true, message: '送信先を追加しました' });
}

/**
 * LINE送信先を削除
 */
function apiRemoveLineUser(userId) {
  if (!userId) return JSON.stringify({ success: false, message: 'ユーザーIDが空です' });
  const props = PropertiesService.getScriptProperties();
  const existing = getLineUserIds_();
  const filtered = existing.filter(id => id !== userId.trim());
  props.setProperty('LINE_USER_IDS', filtered.join(','));
  return JSON.stringify({ success: true, message: '送信先を削除しました' });
}

/**
 * LINEトークンの保存
 */
function apiSaveLineToken(token) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('LINE_CHANNEL_TOKEN', token || '');
    CONFIG.LINE_TOKEN = token || '';
    return JSON.stringify({ success: true, message: 'トークンを保存しました' });
  } catch (e) {
    return JSON.stringify({ success: false, message: e.toString() });
  }
}

/**
 * LINE設定の取得
 */
function apiGetLineSettings() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('LINE_CHANNEL_TOKEN') || '';
  const userIds = getLineUserIds_();
  return JSON.stringify({
    hasToken: token.length > 0,
    tokenPreview: token ? token.substring(0, 8) + '...' : '',
    userIds: userIds
  });
}

/**
 * LINEテスト送信（全登録ユーザーに送信）
 */
function apiTestLineSend() {
  const token = CONFIG.LINE_TOKEN || PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_TOKEN') || '';
  const userIds = getLineUserIds_();
  if (!token) return JSON.stringify({ success: false, message: 'チャネルアクセストークンが未設定です' });
  if (userIds.length === 0) return JSON.stringify({ success: false, message: '送信先ユーザーIDが未登録です' });

  try {
    const msg = [{ type: 'text', text: '\u{2705} AI建築見積システム\nLINE通知のテスト送信です。この通知が届いていれば設定は正常です。' }];
    let res;
    if (userIds.length === 1) {
      res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
        method: 'post', contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: JSON.stringify({ to: userIds[0], messages: msg }),
        muteHttpExceptions: true
      });
    } else {
      res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/multicast', {
        method: 'post', contentType: 'application/json',
        headers: { 'Authorization': 'Bearer ' + token },
        payload: JSON.stringify({ to: userIds, messages: msg }),
        muteHttpExceptions: true
      });
    }
    const code = res.getResponseCode();
    if (code === 200) return JSON.stringify({ success: true, message: 'テスト送信成功（' + userIds.length + '人）！LINEを確認してください' });
    const errBody = JSON.parse(res.getContentText());
    return JSON.stringify({ success: false, message: 'LINE APIエラー (HTTP ' + code + '): ' + (errBody.message || res.getContentText()) });
  } catch (e) {
    return JSON.stringify({ success: false, message: 'テスト送信失敗: ' + e.toString() });
  }
}

/**
 * リマインドメール用トリガーをセットアップ（毎朝9時）
 */
function setupReminderTrigger() {
  removeReminderTrigger();
  ScriptApp.newTrigger('sendReminderEmail_')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  return JSON.stringify({ success: true, message: 'リマインドトリガーを設定しました（毎朝9時）' });
}

/**
 * リマインドメール用トリガーを解除
 */
function removeReminderTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'sendReminderEmail_') ScriptApp.deleteTrigger(t);
  });
  return JSON.stringify({ success: true, message: 'リマインドトリガーを解除しました' });
}

/**
 * リマインドトリガーの状態を取得
 */
function apiGetReminderTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const active = triggers.some(t => t.getHandlerFunction() === 'sendReminderEmail_');
  return JSON.stringify({ active: active });
}

function apiChatWithSystemBot(userMessage, screenContext) {
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

    // 2. 画面コンテキスト情報の構築
    let screenSection = '';
    if (screenContext && screenContext.screenName) {
      screenSection = `\n【現在の画面】\nユーザーは現在「${screenContext.screenName}」を表示しています。`;
      if (screenContext.subScreen) {
        screenSection += `\nサブ画面: ${screenContext.subScreen}`;
      }
      if (screenContext.sidePanel) {
        screenSection += `\nサイドパネル: ${screenContext.sidePanel}を開いています`;
      }
      if (screenContext.leftPanel) {
        screenSection += `\n左パネル: ${screenContext.leftPanel}を開いています`;
      }
    }

    // 3. 組み込みシステム機能ドキュメント
    const builtInDocs = `
【システム機能一覧（組み込みドキュメント）】

■ メインメニュー
アプリ起動時に表示される画面。以下のメニューカードから各機能にアクセスできる。
・見積書作成: 新規見積書の作成。OCR取込、過去データ参照、新規作成が可能
・発注・原価管理: 発注書の単独作成
・案件一覧・売上統計: 全案件の進捗確認、売上集計
・請求書受取: 受取請求書の確認・登録
・管理者画面: 承認フロー、マスタ管理（管理者のみ表示）

■ リマインド機能
メニュー画面の上部に「リマインドパネル」が表示される。アプリ起動時に自動で確認が行われる。
リマインドの種類:
・PDF未作成（青）: 見積書・発注書・請求書がシステムに登録されているが、PDFがまだ作成されていない場合に通知
・入金期日接近（黄）: 請求済みの案件で、入金期日まで3日以内の場合に警告
・入金期日超過（赤）: 請求済みの案件で、入金期日を過ぎても入金が確認されていない場合に警告
・出金期日接近（黄）: 受取請求書の支払期日まで3日以内の場合に警告
・出金期日超過（赤）: 受取請求書の支払期日を過ぎても支払が行われていない場合に警告
入金・出金の期日は、見積書の「決済条件」テキスト（例: 月末締め翌月末払い）から自動計算される。
決済条件が未入力またはパースできない場合はリマインド対象外になる。
受取請求書の出金期日は、決済条件が取得できない場合は請求日から30日後がデフォルト。
メール通知: 管理者がトリガーを設定すると、毎朝9時に期日接近・超過のリマインドメールが自動送信される。
LINE通知: LINE Messaging APIと連携可能。メールと同時にLINEにもリマインド通知が届く。
LINE通知の設定手順:
  1. LINE Developers（https://developers.line.biz/）にログイン
  2. 「プロバイダー」を作成（会社名など任意の名前でOK）
  3. プロバイダー内で「Messaging APIチャネル」を新規作成
  4. チャネル設定画面の「Messaging API設定」タブを開く
  5. 一番下の「チャネルアクセストークン（長期）」の「発行」ボタンを押してトークンをコピー
  6. 作成されたLINE公式アカウントを、通知を受け取りたいLINEアカウントで友だち追加する（QRコードがMessaging API設定タブにある）
  7. ユーザーIDの取得: チャネル基本設定タブの「あなたのユーザーID」を確認。またはWebhookでフォローイベントから取得
  8. 本システムの管理者画面→「設定」タブ→「LINE通知設定」で、コピーしたトークンとユーザーIDを入力して「保存」
  9. 「テスト送信」ボタンを押してLINEにメッセージが届くか確認
注意: リマインドメール通知のトリガーが「無効」の場合、LINE通知も送信されない。必ずトリガーを「有効」にすること。

■ 見積書作成・編集画面
・ヘッダー情報: 顧客名、工事名、工事場所、工期、決済条件、有効期限を入力
・明細テーブル: 工種、品名、仕様、数量、単位、原価、単価、金額、備考、発注先を入力
・左パネル「単価マスタ」: 基本単価マスタや元請別単価マスタから明細に反映
・左パネル「セットテンプレート」: セットマスタから複数明細を一括追加
・右パネル「見積履歴」: 過去の見積データを参照して読み込み
・右パネル「発注先変更」: 明細の発注先を一括変更、発注書作成
・PDF作成: 保存と同時にPDFを生成。御見積書として顧客名・工事名入りのファイル名で保存
・請求書発行: 見積データから御請求書PDFを生成。ステータスが「請求済」に変更される

■ 発注・原価管理画面（単独発注）
・見積書を経由せず、直接発注書を作成できる画面
・発注先、明細（品名・仕様・数量・単位・単価・金額）を入力
・PDF作成: 発注書PDFを生成
・履歴: 過去の発注書を一覧表示、検索、再印刷が可能

■ 案件一覧・売上統計画面
・見積案件一覧: 全見積案件をリスト表示。売上、発注額、入金状況を確認
・発注案件一覧: 全発注をリスト表示。支払状況を確認
・統計カード: 総売上、総発注額、粗利、入金済額などの集計値を表示
・入金ステータスバッジ: 入金済（緑）、一部入金（黄）で表示
・工事台帳: 案件ごとの売上・発注・入出金の詳細台帳をPDF出力

■ 請求書受取画面
・Google Driveの指定フォルダからファイルを取得
・AIが請求書の内容を自動解析（OCR + Gemini）
・工事ID紐付け、請求元、請求金額、相殺額、支払予定額を登録
・登録した請求書は受取請求書リストで管理

■ 管理者画面
・承認管理: 管理者確認中の見積・発注を承認/却下
・経営分析: 年度別の月次売上・粗利グラフ、案件ランキング
・会計連携: 仕訳データのプレビューとCSVダウンロード
・入金管理: 入金記録の登録・編集・一覧表示
・出金管理: 出金記録の登録・編集・一覧表示
・設定: リマインドメール通知のON/OFF切替、LINE通知設定（トークン・ユーザーID入力、テスト送信）

■ AIチャットボット（このシステム解説AI）
・画面右下の紫色のボタンをクリックして開く
・現在表示中の画面を自動認識し、その画面に関する質問に的確に回答
・ドラッグで移動可能
・操作方法や機能の説明を質問できる

■ データ管理
・スプレッドシート: 見積リスト、発注リスト、受取請求書リスト、入金リスト、出金リスト、各種マスタ
・Google Drive: PDFファイルの保存先
・キャッシュ: データの高速読み込みのため、CacheServiceでキャッシュ管理
・権限: 管理者と一般ユーザーで表示・操作権限が異なる
`;

    // 4. プロンプトの構築
    const promptText = `
あなたは「AI建築見積システム v10.0」の操作サポート専門アシスタントです。
以下の【組み込みドキュメント】と【システム情報】の両方を基に、ユーザーからの質問に答えてください。

■回答ルール
1. 操作方法の質問には、具体的なボタン名や画面上の場所、手順を簡潔に案内してください
2. 技術的な仕組みや実装の詳細は、明示的に聞かれた場合のみ説明してください
3. 回答は以下の形式で統一してください：
   - マークダウン記号（**、##、###など）は使わない
   - 箇条書きは「・」を使用
   - 手順は「1. 」「2. 」のように番号付き
   - 適度な改行で読みやすく整形
   - シンプルで分かりやすい日本語
4. ユーザーが「この画面」「ここ」「今の画面」などの指示語を使った場合、【現在の画面】の情報を基に回答してください

■制約事項
・【組み込みドキュメント】と【システム情報】のどちらにも記載されていないことは「わかりません」と答えてください
・回答は日本語のみで行ってください
${screenSection}

${builtInDocs}

【システム情報（外部ナレッジ）】
${systemContext}

【ユーザーの質問】
${userMessage}
`;

    // 4. Gemini APIへのリクエスト
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
    const reply = json.candidates && json.candidates.length > 0 && json.candidates[0].content.parts[0].text;
    if (!reply) {
      return JSON.stringify({ error: "回答を生成できませんでした。" });
    }

    return JSON.stringify({ reply: reply });

  } catch (e) {
    console.error("apiChatWithSystemBot Exception: " + e.toString());
    return JSON.stringify({ error: "システムエラーが発生しました: " + e.toString() });
  }
}

// ── AI操作アシスタント ──────────────────

function buildAssistantPrompt_(userInstruction, ctx) {
  const header = ctx.header || {};
  const items = (ctx.items || []).map((it, i) =>
    `${i}: cat=${it.category||''} prod=${it.product||''} spec=${it.spec||''} qty=${it.qty||0} unit=${it.unit||''} price=${it.price||0} cost=${it.cost||0} vendor=${it.vendor||''}`
  ).join('\n');

  return `あなたは建築見積システムの操作アシスタントです。
ユーザーの自然言語の指示を解析し、見積データを操作するためのJSONアクションを生成してください。

【現在の見積データ】
顧客: ${header.client || ''} / 工事名: ${header.project || ''} / 現場: ${header.location || ''}
明細(${(ctx.items||[]).length}件):
${items}

【アクションタイプ】
- update_item: 明細行のフィールドを更新。filterで対象を絞り、changesで値を設定、またはmodifierで計算。
- update_header: ヘッダー(client,project,location,period,payment,expiry)を更新。
- add_item: 新しい明細行を追加。newItemにcategory,product,spec,qty,unit,price,cost,vendorを指定。
- remove_item: 明細行を削除。filterで対象を絞り込む。

【filterの仕様】
- operator: "all"(全件), "equals"(完全一致), "contains"(部分一致)
- field: "category","product","vendor","spec" のいずれか
- value: 比較する値

【changesの仕様】
変更可能フィールド: vendor, category, product, spec, qty, unit, price, cost, remarks

【modifierの仕様】
- field: "price","qty","cost"
- operation: "multiply"(乗算), "add"(加算), "set"(直接設定)
- value: 数値

【ユーザーの指示】
${userInstruction}

注意:
- confidenceは0〜1で、指示が曖昧な場合は低い値を返してください
- summaryは日本語で変更内容を簡潔に説明してください
- 指示が見積操作と無関係な場合はactions空配列でsummaryに説明を入れてください`;
}

function apiAiAssistantAction(userInstruction, estimateContextJson) {
  try {
    if (!CONFIG.API_KEY) return JSON.stringify({ error: "APIキーが未設定です" });
    if (!userInstruction) return JSON.stringify({ error: "指示を入力してください" });

    const ctx = JSON.parse(estimateContextJson || '{}');
    const prompt = buildAssistantPrompt_(userInstruction, ctx);

    const responseSchema = {
      "type": "OBJECT",
      "properties": {
        "actions": {
          "type": "ARRAY",
          "items": {
            "type": "OBJECT",
            "properties": {
              "type": { "type": "STRING", "description": "update_item|update_header|add_item|remove_item" },
              "filter": {
                "type": "OBJECT",
                "properties": {
                  "field": { "type": "STRING" },
                  "operator": { "type": "STRING", "description": "all|equals|contains" },
                  "value": { "type": "STRING" }
                }
              },
              "changes": {
                "type": "OBJECT",
                "properties": {
                  "vendor": { "type": "STRING" }, "category": { "type": "STRING" },
                  "product": { "type": "STRING" }, "spec": { "type": "STRING" },
                  "qty": { "type": "NUMBER" }, "unit": { "type": "STRING" },
                  "price": { "type": "NUMBER" }, "cost": { "type": "NUMBER" },
                  "remarks": { "type": "STRING" }
                }
              },
              "modifier": {
                "type": "OBJECT",
                "properties": {
                  "field": { "type": "STRING" }, "operation": { "type": "STRING" }, "value": { "type": "NUMBER" }
                }
              },
              "newItem": {
                "type": "OBJECT",
                "properties": {
                  "category": { "type": "STRING" }, "product": { "type": "STRING" },
                  "spec": { "type": "STRING" }, "qty": { "type": "NUMBER" },
                  "unit": { "type": "STRING" }, "price": { "type": "NUMBER" },
                  "cost": { "type": "NUMBER" }, "vendor": { "type": "STRING" }
                }
              }
            }
          }
        },
        "summary": { "type": "STRING" },
        "confidence": { "type": "NUMBER" }
      }
    };

    const res = UrlFetchApp.fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=${CONFIG.API_KEY}`,
      {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: {
            response_mime_type: "application/json",
            response_schema: responseSchema,
            temperature: 0.1
          }
        }),
        muteHttpExceptions: true
      }
    );

    const json = JSON.parse(res.getContentText());
    if (json.error) {
      console.error("AI Assistant Gemini Error: " + JSON.stringify(json.error));
      return JSON.stringify({ error: "AIの応答エラー: " + json.error.message });
    }
    if (!json.candidates || !json.candidates[0]) {
      return JSON.stringify({ error: "AIから回答を取得できませんでした" });
    }

    return json.candidates[0].content.parts[0].text;

  } catch (e) {
    console.error("apiAiAssistantAction Exception: " + e.toString());
    return JSON.stringify({ error: "システムエラー: " + e.toString() });
  }
}

// ── 統合AIアシスタント（全画面対応） ──────────────────

function apiAiAssistantUnified(userInstruction, contextJson) {
  try {
    if (!CONFIG.API_KEY) return JSON.stringify({ error: "APIキーが未設定です" });
    if (!userInstruction) return JSON.stringify({ error: "指示を入力してください" });

    var ctx = JSON.parse(contextJson || '{}');
    var prompt = buildUnifiedPrompt_(userInstruction, ctx);

    // まずFlashで試行、失敗時(タイムアウト/RECITATION等)はProにフォールバック
    var models = ['gemini-3-flash-preview', 'gemini-3.1-pro-preview'];
    var aiResult = null;
    var lastError = '';

    for (var mi = 0; mi < models.length; mi++) {
      var model = models[mi];
      console.log('apiAiAssistantUnified: モデル ' + model + ' で実行');

      try {
        var res = UrlFetchApp.fetch(
          "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + CONFIG.API_KEY,
          {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify({
              contents: [{ parts: [{ text: prompt }] }],
              generationConfig: {
                response_mime_type: "application/json",
                temperature: 0.1
              }
            }),
            muteHttpExceptions: true
          }
        );

        var json = JSON.parse(res.getContentText());
        if (json.error) {
          lastError = json.error.message;
          console.error("AI Unified (" + model + "): APIエラー: " + lastError);
          continue;
        }
        var candidate = json.candidates && json.candidates[0];
        if (!candidate) {
          lastError = 'candidates無し';
          console.error("AI Unified (" + model + "): " + lastError);
          continue;
        }
        if (!candidate.content || !candidate.content.parts || !candidate.content.parts[0]) {
          var reason = candidate.finishReason || 'UNKNOWN';
          lastError = '応答不完全 (' + reason + ')';
          console.log("AI Unified (" + model + "): " + lastError + " → フォールバック");
          continue;
        }

        // 正常レスポンス取得
        aiResult = JSON.parse(candidate.content.parts[0].text);
        console.log('apiAiAssistantUnified: ' + model + ' で正常応答取得: ' + JSON.stringify(aiResult).substring(0, 500));
        break;

      } catch (fetchErr) {
        // タイムアウト等のfetch例外 → 次のモデルへフォールバック
        lastError = fetchErr.toString();
        console.error("AI Unified (" + model + "): fetch例外: " + lastError + " → フォールバック");
        continue;
      }
    }

    if (!aiResult) {
      return JSON.stringify({ error: "AIから応答を取得できませんでした（" + lastError + "）。しばらく待ってから再度お試しください。" });
    }

    // queryタイプの場合はデータ取得して要約
    if (aiResult.responseType === 'query') {
      try {
        var queryData = executeAiQuery_(aiResult.queryType, aiResult.queryParams || {});
        var summary = summarizeQueryResult_(userInstruction, queryData, aiResult.queryType);
        aiResult.queryResult = summary;
      } catch (qe) {
        console.error("executeAiQuery_: エラー発生: " + qe.toString());
        aiResult.queryResult = "データ取得中にエラーが発生しました: " + qe.toString();
      }
    }

    return JSON.stringify(aiResult);

  } catch (e) {
    console.error("apiAiAssistantUnified: 例外発生: " + e.toString());
    return JSON.stringify({ error: "システムエラー: " + e.toString() });
  }
}

function buildUnifiedPrompt_(userInstruction, ctx) {
  var screen = ctx.screen || 'menu';
  var formData = ctx.formData || {};
  var summaryData = ctx.summaryData || {};
  var history = ctx.history || [];

  var screenLabel = { menu: 'ホーム', list: '一覧/統計', edit: '見積編集', admin: '管理', invoice: '請求書受取', dedicated_order: '発注書作成' }[screen] || screen;

  // 画面固有のデータコンテキスト（明細は上限20件に制限しAI処理を高速化）
  var MAX_CONTEXT_ITEMS = 20;
  var dataContext = '';
  if (screen === 'edit' && formData.header) {
    var header = formData.header;
    var allItems = formData.items || [];
    var truncated = allItems.length > MAX_CONTEXT_ITEMS;
    var items = allItems.slice(0, MAX_CONTEXT_ITEMS).map(function(it, i) {
      return i + ': cat=' + (it.category||'') + ' prod=' + (it.product||'') + ' spec=' + (it.spec||'') + ' qty=' + (it.qty||0) + ' unit=' + (it.unit||'') + ' price=' + (it.price||0) + ' cost=' + (it.cost||0) + ' vendor=' + (it.vendor||'');
    }).join('\n');
    if (truncated) items += '\n... (他 ' + (allItems.length - MAX_CONTEXT_ITEMS) + '件省略)';
    dataContext = '\n【現在の見積データ】\n顧客: ' + (header.client||'') + ' / 工事名: ' + (header.project||'') + ' / 現場: ' + (header.location||'') + '\n明細(全' + allItems.length + '件):\n' + items;
  } else if (screen === 'dedicated_order' && formData.header) {
    var oh = formData.header;
    var allOItems = formData.items || [];
    var oTruncated = allOItems.length > MAX_CONTEXT_ITEMS;
    var oItems = allOItems.slice(0, MAX_CONTEXT_ITEMS).map(function(it, i) {
      return i + ': prod=' + (it.product||'') + ' spec=' + (it.spec||'') + ' vendor=' + (it.vendor||'') + ' estQty=' + (it.estQty||0) + ' estPrice=' + (it.estPrice||0) + ' exeQty=' + (it.exeQty||0) + ' exePrice=' + (it.exePrice||0);
    }).join('\n');
    if (oTruncated) oItems += '\n... (他 ' + (allOItems.length - MAX_CONTEXT_ITEMS) + '件省略)';
    dataContext = '\n【現在の発注データ】\n発注先: ' + (oh.vendor||'') + ' / 工事名: ' + (oh.project||'') + ' / 現場: ' + (oh.location||'') + '\n明細(全' + allOItems.length + '件):\n' + oItems;
  } else if (screen === 'invoice' && formData.constructionId !== undefined) {
    dataContext = '\n【現在の請求書データ】\n工事ID: ' + (formData.constructionId||'') + ' / 工事名: ' + (formData.project||'') + ' / 日付: ' + (formData.date||'') + ' / 担当: ' + (formData.person||'') + ' / 業者: ' + (formData.contractor||'') + ' / 金額: ' + (formData.amount||0) + ' / 相殺: ' + (formData.offset||0) + ' / 内容: ' + (formData.content||'') + ' / 現場: ' + (formData.location||'') + ' / ステータス: ' + (formData.status||'');
  } else if (screen === 'list') {
    dataContext = '\n【現在の一覧画面】\nサブタブ: ' + (summaryData.subTab||'projects') + ' / 表示件数: ' + (summaryData.count||0) + '件';
  } else if (screen === 'admin') {
    dataContext = '\n【現在の管理画面】\nサブタブ: ' + (summaryData.subTab||'');
  } else if (screen === 'menu') {
    dataContext = '\n【ホーム画面】\nリマインダー: ' + (summaryData.reminderCount||0) + '件';
  }

  // 画面固有のアクション定義
  var estimateActionsDef = '\n- update_item: 明細行のフィールドを更新。filterで対象を絞り、changesで値を設定、modifierで計算。\n- update_header: ヘッダー(client,project,location,period,payment,expiry)を更新。\n- add_item: 新しい明細行を追加。newItemにcategory,product,spec,qty,unit,price,cost,vendorを指定。\n- remove_item: 明細行を削除。filterで対象を絞り込む。';

  var actionDefs = '\n【共通アクション（全画面）】\n- navigate: 画面遷移。target に遷移先 (menu/list/edit/admin/invoice/dedicated_order) を指定。targetId で特定のデータを開く。targetSubTab でサブタブ指定。\n- query: データ検索。queryType (projects/orders/invoices/analysis/deposits/payments) と queryParams (year, vendor, status, project 等) を指定。\n- load_estimate: 見積履歴から工事名で検索して読み込む。projectNameに工事名（部分一致）を指定。読み込み後は自動的に見積編集画面に遷移する。\n- create_pdf: 現在の見積を保存してPDFを作成する。見積編集画面でのみ使用可能。' +
  '\n\n【複合操作】\n複数の操作を順番に実行できます。actionsの配列に実行順で並べてください。\n順序の原則: 読込(load_estimate) → 変更(update_header/update_item等) → 出力(create_pdf)\n' +
  '\n【複合操作パターン（{A},{B}等はユーザー指定の値で置換）】\n' +
  'パターンA: 「{A}の見積をベースに、工事名を{B}に変えてPDF作成して」\n' +
  '  → actions: [\n' +
  '    { "type": "load_estimate", "projectName": "{A}" },\n' +
  '    { "type": "update_header", "changes": { "project": "{B}" } },\n' +
  '    { "type": "create_pdf" }\n' +
  '  ]\n' +
  'パターンB: 「{A}の見積を開いて、単価を全部{N}倍にしてPDF作って」\n' +
  '  → actions: [\n' +
  '    { "type": "load_estimate", "projectName": "{A}" },\n' +
  '    { "type": "update_item", "filter": { "operator": "all" }, "modifier": { "field": "price", "operation": "multiply", "value": {N} } },\n' +
  '    { "type": "create_pdf" }\n' +
  '  ]\n' +
  'パターンC: 「{A}の見積をベースに、顧客を{B}に、現場を{C}にして」\n' +
  '  → actions: [\n' +
  '    { "type": "load_estimate", "projectName": "{A}" },\n' +
  '    { "type": "update_header", "changes": { "client": "{B}", "location": "{C}" } }\n' +
  '  ]\n' +
  '\n【重要】指示に「〜に変えて」「〜に変更して」等の変更指示がある場合、必ずupdate_headerまたはupdate_itemアクションを生成してください。変更指示を省略しないでください。\n' +
  '\n【曖昧表現の解釈ルール】\n' +
  '- 「〜の内容を探して反映」「〜をベースに」「〜を元に」「〜をコピーして」 → load_estimate\n' +
  '- 「工事名を〜に変えて」 → update_header changes: { "project": "値" }\n' +
  '- 「顧客を〜に変えて」 → update_header changes: { "client": "値" }\n' +
  '- 「現場を〜に変えて」 → update_header changes: { "location": "値" }\n' +
  '- 「単価を〜倍にして」 → update_item modifier: { "field": "price", "operation": "multiply", "value": 数値 }\n' +
  '- 「発注先を〜にして」 → update_item changes: { "vendor": "値" }\n' +
  '- 「見積書を作成して」「PDFにして」「PDF作って」 → create_pdf\n' +
  '- 「全部」「全て」「すべて」 → filter: { "operator": "all" }\n' +
  '- 工事名の検索は部分一致。ユーザー指定のキーワードをそのままprojectNameに設定\n' +
  '- 複数の変更指示は1つのupdate_headerのchangesにまとめる\n' +
  '\nload_estimate後に使える見積操作アクション:' + estimateActionsDef;

  if (screen === 'edit') {
    actionDefs += '\n【見積操作アクション（現在の画面）】' + estimateActionsDef;
  } else if (screen === 'dedicated_order') {
    actionDefs += '\n【発注操作アクション】\n- update_order_item: 発注明細行のフィールドを更新。filterで対象を絞り、changesでproduct,spec,vendor,estQty,estUnit,estPrice,exeQty,exeUnit,exePriceを設定、modifierで計算。\n- update_order_header: 発注ヘッダー(vendor,project,location,period,payment,expiry,date)を更新。\n- add_order_item: 新しい発注明細行を追加。\n- remove_order_item: 発注明細行を削除。';
  } else if (screen === 'invoice') {
    actionDefs += '\n【請求書操作アクション】\n- update_invoice: 請求書フォームのフィールドを更新。changesにconstructionId,project,date,person,contractor,amount,offset,content,location,statusを指定。';
  }

  // 会話履歴コンテキスト
  var historyContext = '';
  if (history.length > 0) {
    historyContext = '\n【直近の会話】\n' + history.map(function(h) {
      return (h.role === 'user' ? 'ユーザー: ' : 'AI: ') + h.text;
    }).join('\n');
  }

  return 'あなたは建築見積システムの操作アシスタントです。\nユーザーの自然言語の指示を解析し、適切なレスポンスタイプ(navigate/query/action)を判別してJSONを生成してください。\n\n【最重要ルール】\n操作(action)は必ず「現在表示中の画面のデータのみ」を対象にしてください。\n「全て」「全部」等の指示は「現在の画面に表示されている明細全て」を意味します。\nデータベースや他の画面のデータを変更する操作は絶対に生成しないでください。\n例: 見積編集画面で「発注先を全て○○にして」→ 今開いている見積の明細だけ変更する\n\n【現在の画面】' + screenLabel + ' (' + screen + ')' + dataContext + actionDefs + '\n\n【filterの仕様】\n- operator: "all"(現在の画面の全件), "equals"(完全一致), "contains"(部分一致)\n- field: 検索対象のフィールド名\n- value: 比較する値\n\n【modifierの仕様（数値変更用）】\n- field: 対象フィールド\n- operation: "multiply"(乗算), "add"(加算), "set"(直接設定)\n- value: 数値\n\n【queryTypeの仕様】\n- projects: 見積一覧。queryParamsにyear,status,clientでフィルタ可。\n- orders: 発注一覧。queryParamsにvendor,statusでフィルタ可。\n- invoices: 請求書一覧。queryParamsにstatus,contractorでフィルタ可。\n- analysis: 分析データ。queryParamsにyearを指定。\n- deposits: 入金一覧。\n- payments: 出金一覧。' + historyContext + '\n\n【ユーザーの指示】\n' + userInstruction + '\n\n【出力JSON形式】\n必ず以下の形式のJSONを返してください。各フィールドにはユーザーが指定した具体的な値を入れてください。空文字""は禁止です。\n{\n  "responseType": "navigate" | "query" | "action",\n  "queryType": "projects|orders|invoices|analysis|deposits|payments",\n  "queryParams": { ... },\n  "actions": [\n    {\n      "type": "アクション種別",\n      "projectName": "load_estimate時の工事名（ユーザー指定の値をそのまま入れる）",\n      "target": "navigate先画面",\n      "targetId": "対象ID",\n      "changes": { "フィールド名": "ユーザー指定の値" },\n      "filter": { "field": "フィールド名", "operator": "all|equals|contains", "value": "値" },\n      "modifier": { "field": "フィールド名", "operation": "multiply|add|set", "value": 数値 },\n      "newItem": { "category": "", "product": "", "spec": "", "qty": 0, "unit": "", "price": 0, "cost": 0, "vendor": "" }\n    }\n  ],\n  "summary": "日本語で操作内容を簡潔に説明",\n  "confidence": 0.0～1.0\n}\n\n注意:\n- responseTypeは navigate / query / action のいずれかを必ず指定\n- confidenceは0〜1で、指示が曖昧な場合は低い値を返してください\n- summaryは日本語で内容を簡潔に説明してください\n- 画面遷移の指示にはnavigate、データに関する質問にはquery、フォーム操作にはactionを使用\n- actionは現在表示中の画面のフォームデータのみを対象とし、他の画面やDBのデータは変更しない\n- 現在の画面で使えないアクションは生成しないでください\n- 指示が操作と無関係な場合はactions空配列でsummaryに説明を入れてください\n- actions内の各オブジェクトには、そのアクションに必要なフィールドだけ含めてください\n- projectName、changes等にはユーザーの指示から抽出した具体的な値を必ず入れてください';
}

function buildUnifiedResponseSchema_(screen) {
  // 見積ヘッダー変更用フィールド
  var headerChangeProps = {
    "client": { "type": "STRING", "description": "顧客名" },
    "project": { "type": "STRING", "description": "工事名" },
    "location": { "type": "STRING", "description": "現場" },
    "period": { "type": "STRING", "description": "工期" },
    "payment": { "type": "STRING", "description": "支払条件" },
    "expiry": { "type": "STRING", "description": "有効期限" }
  };
  // 明細変更用フィールド
  var itemChangeProps = {
    "vendor": { "type": "STRING", "description": "発注先" },
    "category": { "type": "STRING", "description": "工種" },
    "product": { "type": "STRING", "description": "品名" },
    "spec": { "type": "STRING", "description": "仕様" },
    "qty": { "type": "NUMBER", "description": "数量" },
    "unit": { "type": "STRING", "description": "単位" },
    "price": { "type": "NUMBER", "description": "単価" },
    "cost": { "type": "NUMBER", "description": "原価" },
    "remarks": { "type": "STRING", "description": "備考" }
  };
  // 画面ごとに適切なchangesプロパティを選択
  var changesProps = headerChangeProps;
  if (screen === 'dedicated_order') {
    changesProps = {
      "vendor": { "type": "STRING" }, "project": { "type": "STRING" }, "location": { "type": "STRING" },
      "product": { "type": "STRING" }, "spec": { "type": "STRING" },
      "estQty": { "type": "NUMBER" }, "estUnit": { "type": "STRING" }, "estPrice": { "type": "NUMBER" },
      "exeQty": { "type": "NUMBER" }, "exeUnit": { "type": "STRING" }, "exePrice": { "type": "NUMBER" }
    };
  } else if (screen === 'invoice') {
    changesProps = {
      "constructionId": { "type": "STRING" }, "project": { "type": "STRING" }, "date": { "type": "STRING" },
      "person": { "type": "STRING" }, "contractor": { "type": "STRING" }, "amount": { "type": "NUMBER" },
      "offset": { "type": "NUMBER" }, "content": { "type": "STRING" }, "location": { "type": "STRING" }, "status": { "type": "STRING" }
    };
  } else {
    // edit画面: ヘッダーと明細の両方のフィールドを含める
    changesProps = {};
    for (var k in headerChangeProps) changesProps[k] = headerChangeProps[k];
    for (var k2 in itemChangeProps) changesProps[k2] = itemChangeProps[k2];
  }

  var actionProperties = {
    "type": { "type": "STRING", "description": "navigate|update_item|update_header|add_item|remove_item|update_order_item|update_order_header|add_order_item|remove_order_item|update_invoice|load_estimate|create_pdf" },
    "target": { "type": "STRING", "description": "navigate先: menu/list/edit/admin/invoice/dedicated_order" },
    "targetId": { "type": "STRING", "description": "navigate先で開くデータID" },
    "targetSubTab": { "type": "STRING", "description": "navigate先のサブタブ" },
    "filter": {
      "type": "OBJECT",
      "properties": {
        "field": { "type": "STRING" },
        "operator": { "type": "STRING", "description": "all|equals|contains" },
        "value": { "type": "STRING" }
      }
    },
    "changes": { "type": "OBJECT", "description": "変更するフィールドと値のマップ。該当フィールドにユーザー指定の値を設定", "properties": changesProps },
    "modifier": {
      "type": "OBJECT",
      "properties": {
        "field": { "type": "STRING" },
        "operation": { "type": "STRING" },
        "value": { "type": "NUMBER" }
      }
    },
    "newItem": { "type": "OBJECT", "description": "追加する新規行データ", "properties": itemChangeProps },
    "projectName": { "type": "STRING", "description": "load_estimate時の工事名（部分一致検索）。ユーザーが指定した工事名をそのまま設定" }
  };

  return {
    "type": "OBJECT",
    "properties": {
      "responseType": { "type": "STRING", "description": "navigate|query|action" },
      "queryType": { "type": "STRING", "description": "projects|orders|invoices|analysis|deposits|payments" },
      "queryParams": { "type": "OBJECT", "description": "クエリのフィルタパラメータ" },
      "actions": {
        "type": "ARRAY",
        "items": { "type": "OBJECT", "properties": actionProperties }
      },
      "summary": { "type": "STRING" },
      "confidence": { "type": "NUMBER" }
    }
  };
}

function executeAiQuery_(queryType, params) {
  var QUERY_WHITELIST = ['projects', 'orders', 'invoices', 'analysis', 'deposits', 'payments'];
  if (QUERY_WHITELIST.indexOf(queryType) === -1) {
    throw new Error('不正なqueryType: ' + queryType);
  }

  var data;
  if (queryType === 'projects') {
    data = JSON.parse(apiGetProjects());
    if (params.year) data = data.filter(function(p) { return String(p.date || '').indexOf(String(params.year)) >= 0; });
    if (params.status) data = data.filter(function(p) { return p.status === params.status; });
    if (params.client) data = data.filter(function(p) { return String(p.client || '').indexOf(params.client) >= 0; });
    if (params.project) data = data.filter(function(p) { return String(p.project || '').indexOf(params.project) >= 0; });
  } else if (queryType === 'orders') {
    data = JSON.parse(apiGetOrders());
    if (params.vendor) data = data.filter(function(o) { return String(o.vendor || '').indexOf(params.vendor) >= 0; });
    if (params.status) data = data.filter(function(o) { return o.status === params.status; });
  } else if (queryType === 'invoices') {
    data = JSON.parse(apiGetInvoices());
    if (params.status) data = data.filter(function(inv) { return inv.status === params.status; });
    if (params.contractor) data = data.filter(function(inv) { return String(inv.contractor || '').indexOf(params.contractor) >= 0; });
  } else if (queryType === 'analysis') {
    data = JSON.parse(apiGetAnalysisData(params.year || new Date().getFullYear()));
  } else if (queryType === 'deposits') {
    data = JSON.parse(apiGetDeposits());
  } else if (queryType === 'payments') {
    data = JSON.parse(apiGetPayments());
  }

  // 大量データのトランケート（上位50件）
  if (Array.isArray(data) && data.length > 50) {
    data = { items: data.slice(0, 50), total: data.length, truncated: true };
  }

  return data;
}

function summarizeQueryResult_(question, data, queryType) {
  if (!CONFIG.API_KEY) return JSON.stringify(data);

  var dataStr = JSON.stringify(data);
  if (dataStr.length > 8000) dataStr = dataStr.substring(0, 8000) + '... (truncated)';

  var prompt = 'あなたは建築見積システムのデータアナリストです。\nユーザーの質問に対して、以下のデータを元に日本語で簡潔に回答してください。\n\n【ユーザーの質問】\n' + question + '\n\n【データタイプ】' + queryType + '\n\n【取得データ】\n' + dataStr + '\n\n金額は3桁カンマ区切りで表示してください。回答は簡潔に、必要なデータのみ含めてください。';

  try {
    var res = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=" + CONFIG.API_KEY,
      {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.3 }
        }),
        muteHttpExceptions: true
      }
    );
    var json = JSON.parse(res.getContentText());
    if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0]) {
      return json.candidates[0].content.parts[0].text;
    }
  } catch (e) {
    console.error("summarizeQueryResult_ error: " + e.toString());
  }
  return 'データを取得しました（' + (Array.isArray(data) ? data.length : '1') + '件）。詳細は一覧画面で確認してください。';
}

// ===========================================================
// 出来高報告書 機能
// ===========================================================

/** 出来高キャッシュ一括破棄（orderId対応、現在月±2ヶ月分のキーを削除） */
function invalidateProgressCache_(orderId) {
  try {
    var cc = CacheService.getScriptCache();
    cc.remove("progress_data_all");
    cc.remove("progress_report_list");
    cc.remove("progress_report_list_all");
    if (orderId) {
      cc.remove("progress_data_" + orderId);
      cc.remove("progress_data_" + orderId + "_");
    }
    // 現在月±2ヶ月分のキャッシュキーを削除
    var now = new Date();
    for (var delta = -2; delta <= 2; delta++) {
      var d = new Date(now.getFullYear(), now.getMonth() + delta, 1);
      var ym = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
      cc.remove("progress_report_list_" + ym);
      if (orderId) {
        cc.remove("progress_data_" + orderId + "_" + ym);
      }
    }
  } catch (_) {}
}

/** 出来高DBシート取得（なければ作成してヘッダー・ARRAYFORMULA・書式設定） */
function getOrCreateProgressSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.sheetNames.progressDb);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.sheetNames.progressDb);
    Logger.log('出来高DBシート作成');
  }
  // ヘッダー確認・設定
  const firstRow = sheet.getRange(1, 1, 1, PROGRESS_HEADERS.length).getValues()[0];
  const needsHeader = PROGRESS_HEADERS.some((h, i) => firstRow[i] !== h);
  if (needsHeader) {
    sheet.getRange(1, 1, 1, PROGRESS_HEADERS.length).setValues([PROGRESS_HEADERS]);
    sheet.getRange(1, 1, 1, PROGRESS_HEADERS.length)
      .setFontWeight('bold').setBackground('#0d47a1').setFontColor('#ffffff').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 出来高DBにARRAYFORMULA設置 */
function setProgressFormulas_(sheet) {
  sheet.getRange('A2').setFormula('=ARRAYFORMULA(IF(B2:B="","",ROW(B2:B)-1))');
  sheet.getRange('G2').setFormula('=ARRAYFORMULA(IF(B2:B="","",D2:D*F2:F))');
  sheet.getRange('J2').setFormula('=ARRAYFORMULA(IF(B2:B="","",I2:I*F2:F))');
  sheet.getRange('K2').setFormula('=ARRAYFORMULA(IF(B2:B="","",IFERROR(J2:J/G2:G,"")))');
  sheet.getRange('L2').setFormula('=ARRAYFORMULA(IF(B2:B="","",I2:I-H2:H))');
  sheet.getRange('M2').setFormula('=ARRAYFORMULA(IF(B2:B="","",L2:L*F2:F))');
  Logger.log('出来高DB: ARRAYFORMULA設置完了');
}

/** 出来高DBの書式設定 */
function formatProgressSheet_(sheet) {
  sheet.setColumnWidth(PROG_COL.no, 40);
  sheet.setColumnWidth(PROG_COL.name, 180);
  sheet.setColumnWidth(PROG_COL.spec, 140);
  sheet.setColumnWidth(PROG_COL.totalQty, 90);
  sheet.setColumnWidth(PROG_COL.unit, 50);
  sheet.setColumnWidth(PROG_COL.price, 100);
  sheet.setColumnWidth(PROG_COL.estimateAmt, 120);
  sheet.setColumnWidth(PROG_COL.prevCumQty, 110);
  sheet.setColumnWidth(PROG_COL.currCumQty, 110);
  sheet.setColumnWidth(PROG_COL.progressAmt, 120);
  sheet.setColumnWidth(PROG_COL.progressRate, 80);
  sheet.setColumnWidth(PROG_COL.periodQty, 90);
  sheet.setColumnWidth(PROG_COL.periodPayment, 120);
  sheet.setColumnWidth(PROG_COL.estId, 100);
  sheet.setColumnWidth(PROG_COL.orderId, 100);
  sheet.setColumnWidth(PROG_COL.reportMonth, 90);
  var maxRows = 500;
  sheet.getRange(2, PROG_COL.price, maxRows, 1).setNumberFormat('#,##0');
  sheet.getRange(2, PROG_COL.estimateAmt, maxRows, 1).setNumberFormat('#,##0');
  sheet.getRange(2, PROG_COL.progressAmt, maxRows, 1).setNumberFormat('#,##0');
  sheet.getRange(2, PROG_COL.progressRate, maxRows, 1).setNumberFormat('0.0%');
  sheet.getRange(2, PROG_COL.periodPayment, maxRows, 1).setNumberFormat('#,##0');
}

/** 出来高DB全行読み取り */
function readProgressItems_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetNames.progressDb);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const lastRow = sheet.getLastRow();
  const disp = sheet.getRange(2, 1, lastRow - 1, PROGRESS_HEADERS.length).getDisplayValues();
  const raw = sheet.getRange(2, 1, lastRow - 1, PROGRESS_HEADERS.length).getValues();
  return disp.map((row, i) => ({
    rowIndex: i + 2,
    no: row[0], name: row[1], spec: row[2],
    totalQty: raw[i][3], unit: row[4], price: raw[i][5],
    estimateAmt: raw[i][6], prevCumQty: raw[i][7], currCumQty: raw[i][8],
    progressAmt: raw[i][9], progressRate: raw[i][10],
    periodQty: raw[i][11], periodPayment: raw[i][12],
    estId: row[13] || '', orderId: row[14] || '',
    reportMonth: row[15] || '',
    estimateAmtDisp: row[6], progressAmtDisp: row[9],
    progressRateDisp: row[10], periodPaymentDisp: row[12]
  }));
}

// --- 出来高 API関数 ---

/** 報告書一覧取得（orderId単位でグルーピング、reportMonth対応） */
function apiProgressGetReportList(reportMonth) {
  try {
    const c = CacheService.getScriptCache();
    const cacheKey = "progress_report_list_" + (reportMonth || "all");
    const cached = c.get(cacheKey);
    if (cached) return cached;
    var items = readProgressItems_();
    // reportMonth指定時: その月のデータまたは月未設定データのみに絞る
    if (reportMonth) {
      items = items.filter(function(item) { return !item.reportMonth || item.reportMonth === reportMonth; });
    }
    // orderId でグルーピング
    const groups = {};
    items.forEach(item => {
      const oid = item.orderId || '__none__';
      if (!groups[oid]) groups[oid] = [];
      groups[oid].push(item);
    });
    // ヘッダー情報マップ読み取り
    var headersMap = {};
    try {
      var hJson = PropertiesService.getScriptProperties().getProperty('PROGRESS_REPORT_HEADERS');
      if (hJson) headersMap = JSON.parse(hJson);
    } catch (_) {}
    // 旧キーからのマイグレーション
    if (Object.keys(headersMap).length === 0) {
      try {
        var oldJson = PropertiesService.getScriptProperties().getProperty('PROGRESS_REPORT_HEADER');
        if (oldJson) {
          var oldH = JSON.parse(oldJson);
          if (oldH && Object.keys(oldH).length > 0) {
            // 旧データがあれば __none__ に格納
            headersMap['__none__'] = oldH;
          }
        }
      } catch (_) {}
    }
    var result = [];
    Object.keys(groups).forEach(oid => {
      var grp = groups[oid];
      var estimateTotal = 0, progressTotal = 0, periodPaymentTotal = 0;
      grp.forEach(item => {
        estimateTotal += Number(item.estimateAmt) || 0;
        progressTotal += Number(item.progressAmt) || 0;
        periodPaymentTotal += Number(item.periodPayment) || 0;
      });
      var overallRate = estimateTotal > 0 ? progressTotal / estimateTotal : 0;
      result.push({
        orderId: oid === '__none__' ? '' : oid,
        itemCount: grp.length,
        estimateTotal: estimateTotal,
        progressTotal: progressTotal,
        overallRate: overallRate,
        periodPaymentTotal: periodPaymentTotal,
        headerInfo: headersMap[oid] || {}
      });
    });
    var json = JSON.stringify(result);
    try { c.put(cacheKey, json, 60); } catch (_) {}
    return json;
  } catch (e) {
    Logger.log('apiProgressGetReportList: エラー - ' + e.message);
    return '[]';
  }
}

/** 初期設定（シート作成+ARRAYFORMULA+書式） */
function apiProgressSetupSheet() {
  try {
    const sheet = getOrCreateProgressSheet_();
    setProgressFormulas_(sheet);
    formatProgressSheet_(sheet);
    return JSON.stringify({ success: true, message: '出来高DBシートを作成・設定しました' });
  } catch (e) {
    Logger.log('apiProgressSetupSheet: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}

/** 品目取得（orderId指定時: フィルタ返却、reportMonth対応） */
function apiProgressGetItems(orderId, reportMonth) {
  try {
    const c = CacheService.getScriptCache();
    const cacheKey = "progress_data_" + (orderId || "all") + "_" + (reportMonth || "");
    const cached = c.get(cacheKey);
    if (cached) return cached;
    var items = readProgressItems_();
    if (orderId) {
      items = items.filter(function(item) { return item.orderId === orderId; });
    }
    if (reportMonth) {
      items = items.filter(function(item) { return !item.reportMonth || item.reportMonth === reportMonth; });
    }
    const json = JSON.stringify(items);
    try { c.put(cacheKey, json, 60); } catch (_) {}
    return json;
  } catch (e) {
    Logger.log('apiProgressGetItems: エラー - ' + e.message);
    return '[]';
  }
}

/** 累積数量更新（LockService + flush + 行再読込） */
function apiProgressUpdateCumQty(rowIndex, qty) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return JSON.stringify({ success: false, error: '他の処理が実行中です' });
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.sheetNames.progressDb);
    if (!sheet) return JSON.stringify({ success: false, error: '出来高DBが見つかりません' });
    sheet.getRange(rowIndex, PROG_COL.currCumQty).setValue(Number(qty));
    SpreadsheetApp.flush();
    // キャッシュ破棄
    invalidateProgressCache_('');
    // 更新後の行データ返却
    const row = sheet.getRange(rowIndex, 1, 1, PROGRESS_HEADERS.length).getDisplayValues()[0];
    const raw = sheet.getRange(rowIndex, 1, 1, PROGRESS_HEADERS.length).getValues()[0];
    return JSON.stringify({
      success: true,
      row: {
        rowIndex: rowIndex, no: row[0], name: row[1], spec: row[2],
        totalQty: raw[3], unit: row[4], price: raw[5],
        estimateAmt: raw[6], prevCumQty: raw[7], currCumQty: raw[8],
        progressAmt: raw[9], progressRate: raw[10],
        periodQty: raw[11], periodPayment: raw[12],
        estId: row[13] || '', orderId: row[14] || '',
        reportMonth: row[15] || '',
        estimateAmtDisp: row[6], progressAmtDisp: row[9],
        progressRateDisp: row[10], periodPaymentDisp: row[12]
      }
    });
  } catch (e) {
    Logger.log('apiProgressUpdateCumQty: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  } finally {
    lock.releaseLock();
  }
}

/** 見積から品目取込 */
function apiProgressImportFromEstimate(estimateId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return JSON.stringify({ success: false, error: '他の処理が実行中です' });
  }
  try {
    const estData = _getEstimateData(estimateId);
    if (!estData || !estData.items || estData.items.length === 0) {
      return JSON.stringify({ success: false, error: '見積データが見つかりません: ' + estimateId });
    }
    const sheet = getOrCreateProgressSheet_();
    setProgressFormulas_(sheet);
    // 既存キー取得（重複チェック）
    const existingKeys = {};
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existing = sheet.getRange(2, PROG_COL.name, lastRow - 1, 2).getValues();
      existing.forEach(r => { existingKeys[String(r[0]).trim() + '|' + String(r[1]).trim()] = true; });
    }
    const newRows = [];
    let skipped = 0;
    estData.items.forEach(item => {
      const name = String(item.product || '').trim();
      const spec = String(item.spec || '').trim();
      if (!name) return;
      const key = name + '|' + spec;
      if (existingKeys[key]) { skipped++; return; }
      existingKeys[key] = true;
      newRows.push(['', name, spec, Number(item.qty) || 0, item.unit || '', Number(item.price) || 0, '', 0, 0, '', '', '', '', estimateId, '', '']);
    });
    if (newRows.length === 0) {
      return JSON.stringify({ success: true, added: 0, skipped: skipped, message: '新規データはありませんでした' });
    }
    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, newRows.length, PROGRESS_HEADERS.length).setValues(newRows);
    invalidateProgressCache_('');
    Logger.log('apiProgressImportFromEstimate: ' + newRows.length + '件取込, ' + skipped + '件スキップ');
    return JSON.stringify({ success: true, added: newRows.length, skipped: skipped });
  } catch (e) {
    Logger.log('apiProgressImportFromEstimate: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  } finally {
    lock.releaseLock();
  }
}

/** 発注から品目取込（reportMonth対応） */
function apiProgressImportFromOrder(orderId, reportMonth) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return JSON.stringify({ success: false, error: '他の処理が実行中です' });
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orderSheet = ss.getSheetByName(CONFIG.sheetNames.order);
    if (!orderSheet || orderSheet.getLastRow() <= 1) {
      return JSON.stringify({ success: false, error: '発注データが見つかりません' });
    }
    // 発注シートのヘッダー解析
    const oData = orderSheet.getDataRange().getDisplayValues();
    let hIdx = 0;
    for (let i = 0; i < Math.min(10, oData.length); i++) { if (oData[i][0] === 'ID') { hIdx = i; break; } }
    const h = oData[hIdx]; const col = {}; h.forEach((v, i) => col[String(v).trim()] = i);
    const idxId = col['ID'], idxProd = col['品名'], idxSpec = col['仕様'];
    const idxQty = col['数量'], idxUnit = col['単位'], idxPrice = col['単価'];
    const idxEstId = col['関連見積ID'];

    // 対象発注のデータ収集（ID列は先頭行のみ、後続行は空 → currentId追跡）
    const orderItems = [];
    let currentId = '';
    for (let i = hIdx + 1; i < oData.length; i++) {
      const r = oData[i];
      const idCell = String(r[idxId] || '').trim();
      if (idCell && idCell !== 'ID') { currentId = idCell; }
      if (!currentId || currentId !== orderId) continue;
      orderItems.push({
        name: String(r[idxProd] || '').trim(),
        spec: String(r[idxSpec] || '').trim(),
        qty: parseCurrency(r[idxQty]),
        unit: String(r[idxUnit] || '').trim(),
        price: parseCurrency(r[idxPrice]),
        estId: idxEstId !== undefined ? String(r[idxEstId] || '').trim() : ''
      });
    }
    if (orderItems.length === 0) {
      return JSON.stringify({ success: false, error: '発注ID ' + orderId + ' のデータが見つかりません' });
    }
    // 出来高DBに追加（重複チェックは同一orderId+reportMonth内のみ）
    const sheet = getOrCreateProgressSheet_();
    setProgressFormulas_(sheet);
    const existingKeys = {};
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existingData = sheet.getRange(2, PROG_COL.name, lastRow - 1, PROGRESS_HEADERS.length).getValues();
      existingData.forEach(r => {
        var rowOrderId = String(r[PROG_COL.orderId - PROG_COL.name] || '').trim();
        var rowRM = String(r[PROG_COL.reportMonth - PROG_COL.name] || '').trim();
        if (rowOrderId === orderId && (!reportMonth || !rowRM || rowRM === reportMonth)) {
          existingKeys[String(r[0]).trim() + '|' + String(r[1]).trim()] = true;
        }
      });
    }
    const newRows = [];
    let skipped = 0;
    orderItems.forEach(item => {
      if (!item.name) return;
      const key = item.name + '|' + item.spec;
      if (existingKeys[key]) { skipped++; return; }
      existingKeys[key] = true;
      newRows.push(['', item.name, item.spec, item.qty, item.unit, item.price, '', 0, 0, '', '', '', '', item.estId, orderId, reportMonth || '']);
    });
    if (newRows.length === 0) {
      return JSON.stringify({ success: true, added: 0, skipped: skipped, message: '新規データはありませんでした' });
    }
    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, newRows.length, PROGRESS_HEADERS.length).setValues(newRows);
    invalidateProgressCache_(orderId);
    Logger.log('apiProgressImportFromOrder: ' + newRows.length + '件取込, ' + skipped + '件スキップ');
    return JSON.stringify({ success: true, added: newRows.length, skipped: skipped });
  } catch (e) {
    Logger.log('apiProgressImportFromOrder: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  } finally {
    lock.releaseLock();
  }
}

/** 手動品目追加 */
function apiProgressImportManual(jsonArray) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return JSON.stringify({ success: false, error: '他の処理が実行中です' });
  }
  try {
    const items = JSON.parse(jsonArray);
    if (!items || items.length === 0) {
      return JSON.stringify({ success: false, error: 'データがありません' });
    }
    const sheet = getOrCreateProgressSheet_();
    setProgressFormulas_(sheet);
    const existingKeys = {};
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const existing = sheet.getRange(2, PROG_COL.name, lastRow - 1, 2).getValues();
      existing.forEach(r => { existingKeys[String(r[0]).trim() + '|' + String(r[1]).trim()] = true; });
    }
    const newRows = [];
    let skipped = 0;
    items.forEach(item => {
      const name = String(item.name || '').trim();
      const spec = String(item.spec || '').trim();
      if (!name) return;
      const key = name + '|' + spec;
      if (existingKeys[key]) { skipped++; return; }
      existingKeys[key] = true;
      newRows.push(['', name, spec, parseCurrency(item.qty), String(item.unit || '').trim(), parseCurrency(item.price), '', 0, 0, '', '', '', '', '', '', '']);
    });
    if (newRows.length === 0) {
      return JSON.stringify({ success: true, added: 0, skipped: skipped, message: '新規データはありませんでした' });
    }
    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, newRows.length, PROGRESS_HEADERS.length).setValues(newRows);
    invalidateProgressCache_('');
    Logger.log('apiProgressImportManual: ' + newRows.length + '件追加, ' + skipped + '件スキップ');
    return JSON.stringify({ success: true, added: newRows.length, skipped: skipped });
  } catch (e) {
    Logger.log('apiProgressImportManual: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  } finally {
    lock.releaseLock();
  }
}

/** 工事単位月締め（orderId + currentYM → 翌月行を新規作成） */
function apiProgressMonthlyCloseOrder(orderId, currentYM) {
  if (!orderId || !currentYM) {
    return JSON.stringify({ success: false, error: '発注IDと対象年月が必要です' });
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return JSON.stringify({ success: false, error: '他の処理が実行中です' });
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.sheetNames.progressDb);
    if (!sheet || sheet.getLastRow() <= 1) {
      return JSON.stringify({ success: false, error: '出来高DBにデータがありません' });
    }
    // nextYM算出
    var parts = currentYM.split('-');
    var ny = Number(parts[0]), nm = Number(parts[1]) + 1;
    if (nm > 12) { nm = 1; ny++; }
    var nextYM = ny + '-' + String(nm).padStart(2, '0');

    // 対象行を取得（orderId一致 かつ reportMonth=currentYM または空）
    var allItems = readProgressItems_();
    var targetItems = allItems.filter(function(item) {
      return item.orderId === orderId && (!item.reportMonth || item.reportMonth === currentYM);
    });
    if (targetItems.length === 0) {
      return JSON.stringify({ success: false, error: '対象データが見つかりません（orderId=' + orderId + ', 月=' + currentYM + '）' });
    }

    // 二重締め防止: nextYMのデータが既にあればエラー
    var nextExists = allItems.some(function(item) {
      return item.orderId === orderId && item.reportMonth === nextYM;
    });
    if (nextExists) {
      return JSON.stringify({ success: false, error: nextYM + ' のデータが既に存在します。二重締めはできません。' });
    }

    // 1) reportMonth空の行にcurrentYMをタグ付け
    targetItems.forEach(function(item) {
      if (!item.reportMonth) {
        sheet.getRange(item.rowIndex, PROG_COL.reportMonth).setValue(currentYM);
      }
    });

    // 2) 翌月行を新規作成
    var newRows = [];
    targetItems.forEach(function(item) {
      var currCum = Number(item.currCumQty) || 0;
      newRows.push([
        '', item.name, item.spec, Number(item.totalQty) || 0, item.unit, Number(item.price) || 0,
        '', currCum, currCum, '', '', '', '',
        item.estId || '', orderId, nextYM
      ]);
    });
    var startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, newRows.length, PROGRESS_HEADERS.length).setValues(newRows);
    setProgressFormulas_(sheet);
    SpreadsheetApp.flush();
    invalidateProgressCache_(orderId);

    Logger.log('apiProgressMonthlyCloseOrder: orderId=' + orderId + ' ' + currentYM + '→' + nextYM + ' ' + newRows.length + '件作成');
    return JSON.stringify({ success: true, nextMonth: nextYM, count: newRows.length });
  } catch (e) {
    Logger.log('apiProgressMonthlyCloseOrder: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  } finally {
    lock.releaseLock();
  }
}

/** 報告書ヘッダー取得（ScriptProperties、orderId対応マップ形式） */
function apiProgressGetReportHeader(orderId) {
  try {
    const props = PropertiesService.getScriptProperties();
    var headersMap = {};
    var hJson = props.getProperty('PROGRESS_REPORT_HEADERS');
    if (hJson) {
      headersMap = JSON.parse(hJson);
    } else {
      // 旧キーからのマイグレーション
      var oldJson = props.getProperty('PROGRESS_REPORT_HEADER');
      if (oldJson) {
        var oldH = JSON.parse(oldJson);
        if (oldH && Object.keys(oldH).length > 0) {
          headersMap['__none__'] = oldH;
          props.setProperty('PROGRESS_REPORT_HEADERS', JSON.stringify(headersMap));
        }
      }
    }
    var key = orderId || '__none__';
    return JSON.stringify(headersMap[key] || {});
  } catch (e) {
    Logger.log('apiProgressGetReportHeader: エラー - ' + e.message);
    return JSON.stringify({});
  }
}

/** 報告書ヘッダー保存（orderId対応マップ形式） */
function apiProgressSaveReportHeader(headerJson, orderId) {
  try {
    const props = PropertiesService.getScriptProperties();
    var headersMap = {};
    var hJson = props.getProperty('PROGRESS_REPORT_HEADERS');
    if (hJson) headersMap = JSON.parse(hJson);
    var key = orderId || '__none__';
    headersMap[key] = JSON.parse(headerJson);
    props.setProperty('PROGRESS_REPORT_HEADERS', JSON.stringify(headersMap));
    Logger.log('apiProgressSaveReportHeader: ヘッダー保存完了 (orderId=' + key + ')');
    return JSON.stringify({ success: true });
  } catch (e) {
    Logger.log('apiProgressSaveReportHeader: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}

/** PDF生成（HTMLテンプレート → Drive保存、orderId+reportMonth対応） */
function apiProgressGeneratePdf(headerJson, orderId, reportMonth) {
  try {
    const header = JSON.parse(headerJson);
    var items = readProgressItems_();
    if (orderId) {
      items = items.filter(function(item) { return item.orderId === orderId; });
    }
    if (reportMonth) {
      items = items.filter(function(item) { return !item.reportMonth || item.reportMonth === reportMonth; });
    }
    if (items.length === 0) {
      return JSON.stringify({ success: false, error: '出来高DBにデータがありません' });
    }
    // サマリー計算
    let estimateTotal = 0, progressTotal = 0, periodPaymentTotal = 0;
    const pdfItems = items.map(item => {
      const ea = Number(item.estimateAmt) || 0;
      const pa = Number(item.progressAmt) || 0;
      const pp = Number(item.periodPayment) || 0;
      estimateTotal += ea;
      progressTotal += pa;
      periodPaymentTotal += pp;
      return {
        name: item.name, spec: item.spec,
        totalQty: item.totalQty, unit: item.unit, price: item.price,
        estimateAmt: ea,
        currCumQty: item.currCumQty,
        progressAmt: pa,
        progressRate: item.progressRate,
        periodQty: item.periodQty,
        periodPayment: pp
      };
    });
    const overallRate = estimateTotal > 0 ? progressTotal / estimateTotal : 0;
    const summary = { estimateTotal, progressTotal, overallRate, periodPaymentTotal };

    // ページネーション（1ページ目20行、2ページ目以降25行）
    const pages = paginateItems(pdfItems, 20, 25);

    // テンプレート展開
    const tmpl = HtmlService.createTemplateFromFile('progress_report_template');
    tmpl.data = {
      header: header,
      items: pdfItems,
      pages: pages,
      summary: summary,
      createDate: getJapaneseDateStr(new Date())
    };
    const htmlOutput = tmpl.evaluate().getContent();

    // PDF変換
    const blob = Utilities.newBlob(htmlOutput, 'text/html', 'report.html');
    const pdfBlob = blob.getAs('application/pdf');
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
    const projectName = header.projectName ? '_' + header.projectName : '';
    const fileName = '工事出来高報告書' + projectName + '_' + dateStr + '.pdf';
    pdfBlob.setName(fileName);

    // Drive保存
    const folder = getSaveFolder();
    const file = folder.createFile(pdfBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    Logger.log('apiProgressGeneratePdf: ' + fileName + ' を保存 → ' + file.getUrl());
    return JSON.stringify({ success: true, url: file.getUrl(), fileName: fileName });
  } catch (e) {
    Logger.log('apiProgressGeneratePdf: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  }
}

/** 出来高 累積数量一括更新 */
function apiProgressBatchUpdate(jsonArray) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return JSON.stringify({ success: false, error: '他の処理が実行中です' });
  }
  try {
    const rows = JSON.parse(jsonArray);
    if (!rows || rows.length === 0) {
      return JSON.stringify({ success: false, error: '更新データがありません' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.sheetNames.progressDb);
    if (!sheet) return JSON.stringify({ success: false, error: '出来高DBが見つかりません' });
    rows.forEach(function(r) {
      const ri = Number(r.rowIndex);
      if (ri >= 2) {
        sheet.getRange(ri, PROG_COL.currCumQty).setValue(Number(r.currCumQty) || 0);
      }
    });
    SpreadsheetApp.flush();
    invalidateProgressCache_('');
    const items = readProgressItems_();
    Logger.log('apiProgressBatchUpdate: ' + rows.length + '行を一括更新');
    return JSON.stringify({ success: true, items: items });
  } catch (e) {
    Logger.log('apiProgressBatchUpdate: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  } finally {
    lock.releaseLock();
  }
}

/** 出来高 行削除 */
function apiProgressDeleteRow(rowIndex) {
  const ri = Number(rowIndex);
  if (ri < 2) {
    return JSON.stringify({ success: false, error: 'ヘッダー行は削除できません' });
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return JSON.stringify({ success: false, error: '他の処理が実行中です' });
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.sheetNames.progressDb);
    if (!sheet) return JSON.stringify({ success: false, error: '出来高DBが見つかりません' });
    if (ri > sheet.getLastRow()) {
      return JSON.stringify({ success: false, error: '指定行が存在しません' });
    }
    sheet.deleteRow(ri);
    setProgressFormulas_(sheet);
    SpreadsheetApp.flush();
    invalidateProgressCache_('');
    const items = readProgressItems_();
    Logger.log('apiProgressDeleteRow: 行 ' + ri + ' を削除');
    return JSON.stringify({ success: true, items: items });
  } catch (e) {
    Logger.log('apiProgressDeleteRow: エラー - ' + e.message);
    return JSON.stringify({ success: false, error: e.message });
  } finally {
    lock.releaseLock();
  }
}