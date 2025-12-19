/**
 * ==========================================================================
 * Google Apps Script 後端 API 腳本
 * 用途：作為網頁前端與 Google Sheets 資料庫之間的中介層 (CRUD 操作)
 * 功能：
 * 1. 讀取 (Read): 支援分頁 (Pagination) 與 多欄位搜尋 (Search)
 * 2. 新增 (Create): 單筆或批次新增，自動產生 UUID 與 createdAt
 * 3. 更新 (Update): 依據 id 更新資料，自動補齊舊資料缺少的系統欄位
 * 4. 刪除 (Delete): 依據 id 刪除資料
 * ==========================================================================
 */

// 定義 Google 試算表的 ID (請確認此 ID 對應到正確的試算表檔案)
const SPREADSHEET_ID = '1Ev9CL1iuSblh27YaazjoKe2B_O50193Z4-jf25-8a2U';

// Firebase Web API Key (用於驗證 ID Token)
const FIREBASE_API_KEY = 'AIzaSyDxtxWF-jQMMxvAeMjP5-HJ6y6_QBxLdsY';

// 允許存取的 Email 白名單 (與前端 PREDEFINED 一致)
const ALLOWED_EMAILS = [
  'alenchen@stust.edu.tw',
  'v1277.chen@gmail.com.tw'
];

/**
 * 驗證 Firebase ID Token
 * @param {string} idToken - Firebase ID Token
 * @return {Object} { valid: boolean, email: string, error?: string }
 */
function verifyFirebaseToken(idToken) {
  if (!idToken) {
    return { valid: false, error: 'No token provided' };
  }
  
  try {
    // 使用 Firebase Auth REST API 驗證 Token
    const url = 'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY;
    const response = UrlFetchApp.fetch(url, {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify({ idToken: idToken }),
      muteHttpExceptions: true
    });
    
    const statusCode = response.getResponseCode();
    const data = JSON.parse(response.getContentText());
    
    if (statusCode === 200 && data.users && data.users.length > 0) {
      const user = data.users[0];
      return { 
        valid: true, 
        email: user.email,
        uid: user.localId
      };
    } else {
      return { valid: false, error: data.error ? data.error.message : 'Token validation failed' };
    }
  } catch (e) {
    Logger.log('Token verification error: ' + e);
    return { valid: false, error: e.toString() };
  }
}

/**
 * 檢查 Email 是否在白名單或已授權使用者名單中
 * @param {string} email - 使用者 Email
 * @return {boolean} 是否允許存取
 */
function isEmailAllowed(email) {
  if (!email) return false;
  // 白名單永遠允許
  if (ALLOWED_EMAILS.includes(email)) return true;
  // 其他使用者：你可以加入額外邏輯，例如檢查 Firestore 中的 sys_users
  // 目前簡化處理：只要 Token 有效且 email 存在就允許
  return true;
}

/**
 * 記錄 API 請求日誌至 log 頁籤
 * @param {Spreadsheet} ss - 試算表物件
 * @param {Object} logData - 日誌資料
 */
function logRequest(ss, logData) {
  try {
    let logSheet = ss.getSheetByName('log');
    
    // 若 log 頁籤不存在則建立
    if (!logSheet) {
      logSheet = ss.insertSheet('log');
      // 設定標題列
      logSheet.appendRow(['timestamp', 'email', 'action', 'sheet', 'targetId', 'status', 'message', 'duration_ms']);
      // 凍結標題列
      logSheet.setFrozenRows(1);
    }
    
    // 寫入日誌
    logSheet.appendRow([
      logData.timestamp || new Date().toISOString(),
      logData.email || '',
      logData.action || '',
      logData.sheet || '',
      logData.targetId || '',
      logData.status || '',
      logData.message || '',
      logData.duration || ''
    ]);
    
    // 可選：保持日誌在 10000 筆以內，超過則刪除最舊的
    const maxRows = 10000;
    const currentRows = logSheet.getLastRow();
    if (currentRows > maxRows) {
      logSheet.deleteRows(2, currentRows - maxRows);
    }
  } catch (e) {
    // 日誌寫入失敗不應影響主要功能
    Logger.log('Log write error: ' + e);
  }
}

/**
 * 記錄搜尋日誌至 search 頁籤
 * @param {Spreadsheet} ss - 試算表物件
 * @param {Object} logData - 日誌資料
 */
function logSearch(ss, logData) {
  try {
    let searchSheet = ss.getSheetByName('search');
    
    // 若 search 頁籤不存在則建立
    if (!searchSheet) {
      searchSheet = ss.insertSheet('search');
      // 設定標題列
      searchSheet.appendRow(['timestamp', 'email', 'sheet', 'searchCriteria', 'resultCount', 'duration_ms']);
      // 凍結標題列
      searchSheet.setFrozenRows(1);
    }
    
    // 寫入日誌
    searchSheet.appendRow([
      logData.timestamp || new Date().toISOString(),
      logData.email || '',
      logData.sheet || '',
      logData.searchCriteria || '',
      logData.resultCount || 0,
      logData.duration || ''
    ]);
    
    // 可選：保持日誌在 10000 筆以內
    const maxRows = 10000;
    const currentRows = searchSheet.getLastRow();
    if (currentRows > maxRows) {
      searchSheet.deleteRows(2, currentRows - maxRows);
    }
  } catch (e) {
    Logger.log('Search log write error: ' + e);
  }
}

/**
 * 處理 HTTP GET 請求的進入點
 * GET 請求通常用於讀取資料 (Read)
 * @param {Object} e - 事件參數，包含查詢參數 (e.parameter)
 * @return {GoogleAppsScript.Content.TextOutput} JSON 格式的回傳結果
 */
function doGet(e) { return handleRequest(e); }

/**
 * 處理 HTTP POST 請求的進入點
 * POST 請求通常用於寫入或修改資料 (Create, Update, Delete)
 * @param {Object} e - 事件參數，包含 POST Body (e.postData)
 * @return {GoogleAppsScript.Content.TextOutput} JSON 格式的回傳結果
 */
function doPost(e) { return handleRequest(e); }

/**
 * 核心請求處理函式
 * 負責解析參數、分派動作、並處理併發鎖定 (LockService) 以確保資料一致性
 * [SECURITY] 新增 Firebase ID Token 驗證
 * @param {Object} e - 事件參數
 */
function handleRequest(e) {
  // 取得腳本鎖定，防止多位使用者同時寫入導致資料錯亂 (Race Condition)
  // 嘗試鎖定 30 秒 (30000 ms)，若逾時則會拋出錯誤
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    const params = e.parameter || {};
    
    // 解析 POST Body (如果是 POST 請求，資料通常在 body 中)
    let postData = null;
    if (e.postData && e.postData.contents) {
      try { postData = JSON.parse(e.postData.contents); } catch (err) {
        // 若解析 JSON 失敗，postData 維持 null，後續邏輯會 fallback 使用 e.parameter
      }
    }
    
    // === [SECURITY] Firebase ID Token 驗證 ===
    const idToken = postData ? postData.idToken : params.idToken;
    const authResult = verifyFirebaseToken(idToken);
    const startTime = Date.now(); // 記錄開始時間
    
    // 預先取得動作參數 (用於日誌記錄)
    const reqAction = postData ? postData.action : params.action;
    const reqSheet = postData ? postData.sheet : params.sheet;
    const reqId = postData ? postData.id : params.id;
    
    // 開啟試算表 (提前開啟以便記錄日誌)
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    if (!authResult.valid) {
      // 記錄驗證失敗日誌
      logRequest(ss, {
        timestamp: new Date().toISOString(),
        email: 'UNKNOWN',
        action: reqAction,
        sheet: reqSheet,
        targetId: reqId,
        status: 'AUTH_FAILED',
        message: authResult.error || 'Invalid token',
        duration: Date.now() - startTime
      });
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: 'Unauthorized: ' + (authResult.error || 'Invalid token'),
        code: 'AUTH_FAILED'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (!isEmailAllowed(authResult.email)) {
      // 記錄權限不足日誌
      logRequest(ss, {
        timestamp: new Date().toISOString(),
        email: authResult.email,
        action: reqAction,
        sheet: reqSheet,
        targetId: reqId,
        status: 'ACCESS_DENIED',
        message: 'Email not authorized',
        duration: Date.now() - startTime
      });
      return ContentService.createTextOutput(JSON.stringify({ 
        status: 'error', 
        message: 'Forbidden: Email not authorized',
        code: 'ACCESS_DENIED'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    // === 驗證通過 ===
    
    // 資料內容 (物件或陣列)
    const reqData = postData ? postData.data : null;

    // 分頁與搜尋參數 (通常用於 read 動作)
    const page = parseInt(postData ? postData.page : params.page) || 1;           // 頁碼，預設第 1 頁
    const pageSize = parseInt(postData ? postData.pageSize : params.pageSize) || 0; // 每頁筆數，0 代表全部
    const searchParams = postData ? postData.search : (params.search ? JSON.parse(params.search) : {}); // 搜尋條件物件

    let result = {};

    // 根據動作分派給對應函式
    if (reqAction === 'read') {
      result = readSheet(ss, reqSheet, page, pageSize, searchParams);
      // 記錄搜尋日誌至 search 頁籤
      logSearch(ss, {
        timestamp: new Date().toISOString(),
        email: authResult.email,
        sheet: reqSheet,
        searchCriteria: JSON.stringify(searchParams),
        resultCount: result.total || 0,
        duration: Date.now() - startTime
      });
    }
    else if (reqAction === 'create') result = createRow(ss, reqSheet, reqData);
    else if (reqAction === 'createBatch') result = createBatch(ss, reqSheet, reqData);
    else if (reqAction === 'update') result = updateRow(ss, reqSheet, reqId, reqData);
    else if (reqAction === 'delete') result = deleteRow(ss, reqSheet, reqId);
    else if (reqAction === 'getPrxCount') result = getPrxCount(ss, reqId); // reqId = CID1
    else if (reqAction === 'updateAllPrxCount') result = updateAllPrxCount(ss); // 批次更新所有 prxCount
    else result = { status: 'error', message: 'Unknown action: ' + reqAction };

    // 只記錄寫入操作「失敗」的日誌 (成功不記錄)
    if (reqAction !== 'read' && result.status === 'error') {
      logRequest(ss, {
        timestamp: new Date().toISOString(),
        email: authResult.email,
        action: reqAction,
        sheet: reqSheet,
        targetId: reqId,
        status: 'error',
        message: result.message || '',
        duration: Date.now() - startTime
      });
    }

    // 設定回傳格式為 JSON
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    // 全域錯誤處理：回傳錯誤訊息
    return ContentService.createTextOutput(JSON.stringify({ 
      status: 'error', 
      message: err.toString(), 
      stack: err.stack 
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    // 無論成功與否，最後必須釋放鎖定
    lock.releaseLock();
  }
}

/**
 * 格式化寫入的值 (解決 Google Sheets 掉零問題與格式轉換)
 * 針對 CID1 (客戶編號) 與 TEL (電話)，強制加上單引號，讓 Sheet 視為純文字，避免開頭的 0 被去除。
 * @param {string} key - 欄位名稱
 * @param {any} value - 欄位值
 * @return {string} 格式化後的值
 */
function formatValue(key, value) {
  if (value === undefined || value === null) return '';
  const strVal = String(value);
  if (key === 'CID1' || key === 'TEL') {
    // 如果已經有單引號則不加，否則加上單引號
    return strVal.startsWith("'") ? strVal : "'" + strVal;
  }
  return strVal;
}

/**
 * 讀取資料 (Read) - 支援分頁與搜尋
 * @param {Spreadsheet} ss - 試算表物件
 * @param {string} sheetName - 工作表名稱
 * @param {number} page - 目前頁碼 (從 1 開始)
 * @param {number} pageSize - 每頁筆數 (0 代表不分頁，回傳全部)
 * @param {Object} searchParams - 搜尋過濾條件 (Key-Value 對應)
 * @return {Object} 包含資料列表、總筆數、頁次資訊的物件
 */
/**
 * 讀取資料 (Read) - 支援分頁與搜尋
 * [Optimized] 使用 TextFinder 在 Sheet 層級做搜尋，大幅提升效能
 * @param {Spreadsheet} ss - 試算表物件
 * @param {string} sheetName - 工作表名稱
 * @param {number} page - 目前頁碼 (從 1 開始)
 * @param {number} pageSize - 每頁筆數 (0 代表不分頁，回傳全部)
 * @param {Object} searchParams - 搜尋過濾條件 (Key-Value 對應)
 * @return {Object} 包含資料列表、總筆數、頁次資訊的物件
 */
function readSheet(ss, sheetName, page, pageSize, searchParams) {
  const sheet = getOrCreateSheet(ss, sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // 若試算表完全空白，回傳空結果
  if (lastRow === 0) return { status: 'success', data: [], total: 0, page: 1, totalPages: 0 };
  
  // 取得標題列
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // 決定要讀取的列號集合
  let targetRowNumbers = null; // null 代表全部讀取
  
  // --- 使用 TextFinder 優化搜尋 ---
  if (searchParams && Object.keys(searchParams).length > 0) {
    const searchKeys = Object.keys(searchParams).filter(k => searchParams[k] && String(searchParams[k]).trim());
    
    if (searchKeys.length > 0) {
      // 對每個搜尋條件用 TextFinder 找出符合的列
      let matchedRowSets = [];
      
      for (const key of searchKeys) {
        const searchVal = String(searchParams[key]).trim();
        const colIndex = headers.indexOf(key);
        
        if (colIndex === -1 || !searchVal) continue;
        
        // 用 TextFinder 搜尋該欄位 (只搜資料區，不含標題列)
        const searchRange = sheet.getRange(2, colIndex + 1, lastRow - 1, 1);
        
        // CID1 欄位使用精確匹配，其他欄位使用模糊搜尋
        const useExactMatch = (key === 'CID1');
        const finder = searchRange.createTextFinder(searchVal)
          .matchCase(false)                   // 不區分大小寫
          .matchEntireCell(useExactMatch);    // CID1 精確匹配，其他模糊搜尋
        
        const matches = finder.findAll();
        const rowSet = new Set(matches.map(cell => cell.getRow()));
        matchedRowSets.push(rowSet);
      }
      
      // 取所有條件的交集 (AND 邏輯)
      if (matchedRowSets.length > 0) {
        targetRowNumbers = matchedRowSets.reduce((acc, set) => {
          return new Set([...acc].filter(x => set.has(x)));
        });
      } else {
        // 搜尋條件都無效，回傳空
        targetRowNumbers = new Set();
      }
    }
  }
  
  // --- 讀取資料 ---
  let allRows = [];
  
  if (targetRowNumbers === null) {
    // 無搜尋條件，讀取全部資料
    const rawData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    allRows = convertToObjects(rawData, headers);
  } else if (targetRowNumbers.size > 0) {
    // 有搜尋結果，只讀取符合條件的列
    const sortedRows = Array.from(targetRowNumbers).sort((a, b) => a - b);
    
    // 批次讀取優化：如果符合的列很多，分批讀取避免太多 API 呼叫
    if (sortedRows.length > 50) {
      // 符合列數較多時，一次讀全部再篩選（避免多次 getRange 呼叫）
      const rawData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      const rowSet = new Set(sortedRows);
      for (let i = 0; i < rawData.length; i++) {
        const actualRow = i + 2; // rawData index 0 = sheet row 2
        if (rowSet.has(actualRow)) {
          const rowObj = convertRowToObject(rawData[i], headers);
          if (rowObj) allRows.push(rowObj);
        }
      }
    } else {
      // 符合列數較少，逐列讀取更有效率
      for (const rowNum of sortedRows) {
        const rowData = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
        const rowObj = convertRowToObject(rowData, headers);
        if (rowObj) allRows.push(rowObj);
      }
    }
  }
  // else: targetRowNumbers.size === 0，沒有符合的資料，allRows 維持空陣列

  // --- 分頁邏輯 ---
  const totalCount = allRows.length;
  let pagedData = allRows;
  let totalPages = 1;

  if (pageSize > 0) {
    const startIndex = (page - 1) * pageSize;
    pagedData = allRows.slice(startIndex, startIndex + pageSize);
    totalPages = Math.ceil(totalCount / pageSize);
  }

  return { 
    status: 'success', 
    data: pagedData, 
    total: totalCount, 
    page: page, 
    totalPages: totalPages,
    pageSize: pageSize 
  };
}

/**
 * 將二維陣列轉換為物件陣列
 * @param {Array} rawData - 二維陣列資料
 * @param {Array} headers - 標題列
 * @return {Array} 物件陣列
 */
function convertToObjects(rawData, headers) {
  const result = [];
  for (let i = 0; i < rawData.length; i++) {
    const rowObj = convertRowToObject(rawData[i], headers);
    if (rowObj) result.push(rowObj);
  }
  return result;
}

/**
 * 將單列資料轉換為物件
 * @param {Array} row - 單列資料
 * @param {Array} headers - 標題列
 * @return {Object|null} 物件或 null (空列)
 */
function convertRowToObject(row, headers) {
  const rowObj = {};
  let hasData = false;
  
  for (let j = 0; j < headers.length; j++) {
    const header = headers[j];
    const val = row[j];
    
    if (val instanceof Date) {
      rowObj[header] = val.toISOString();
    } else {
      rowObj[header] = String(val);
    }
    
    if (val !== "") hasData = true;
  }
  
  return hasData ? rowObj : null;
}

/**
 * 新增單筆資料 (Create)
 * @param {Spreadsheet} ss - 試算表物件
 * @param {string} sheetName - 工作表名稱
 * @param {Object} data - 要新增的資料物件
 * @return {Object} 執行結果與寫入的資料
 */
function createRow(ss, sheetName, data) {
  const sheet = getOrCreateSheet(ss, sheetName);
  const headers = getHeaders(sheet);
  
  // 自動生成 ID (UUID) - 用於系統內部唯一識別
  if (!data.id) data.id = Utilities.getUuid();
  // 自動生成 createdAt (ISO String) - 用於記錄建立時間
  if (!data.createdAt) data.createdAt = new Date().toISOString();
  
  // 若傳入的資料包含新欄位，自動更新 Sheet 的標頭
  const newHeaders = updateHeaders(sheet, headers, data);
  const row = [];
  
  // 依據最新標頭的順序填入資料，並套用格式化 (防掉零)
  for (const header of newHeaders) {
    row.push(formatValue(header, data[header]));
  }
  
  sheet.appendRow(row);
  return { status: 'success', data: data };
}

/**
 * 批次新增資料 (Batch Create) - 用於 CSV 匯入效能優化
 * @param {Spreadsheet} ss - 試算表物件
 * @param {string} sheetName - 工作表名稱
 * @param {Array<Object>} dataArray - 資料物件陣列
 * @return {Object} 成功筆數
 */
function createBatch(ss, sheetName, dataArray) {
  if (!Array.isArray(dataArray) || dataArray.length === 0) return { status: 'success', count: 0 };
  const sheet = getOrCreateSheet(ss, sheetName);
  let headers = getHeaders(sheet);
  
  // 收集所有資料的所有欄位名稱 (聯集)
  const allKeys = new Set(headers);
  dataArray.forEach(d => Object.keys(d).forEach(k => allKeys.add(k)));
  const newHeaderList = Array.from(allKeys);
  
  // 如果有新欄位，更新 Sheet 標頭
  if (newHeaderList.length > headers.length) {
    sheet.getRange(1, 1, 1, newHeaderList.length).setValues([newHeaderList]);
    headers = newHeaderList;
  }

  // 準備二維陣列進行一次性寫入，大幅提升效能
  const rows = dataArray.map(data => {
    if (!data.id) data.id = Utilities.getUuid();
    if (!data.createdAt) data.createdAt = new Date().toISOString();
    return headers.map(header => formatValue(header, data[header]));
  });

  if (rows.length > 0) {
    // 取得最後一列的位置，接續寫入
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
  return { status: 'success', count: rows.length };
}

/**
 * 更新資料 (Update)
 * 優先使用 ID 搜尋，若無 ID 則嘗試使用 CID1 搜尋 (相容舊資料)
 * 特點：若找到舊資料且該列沒有 id 或 createdAt，會在此時自動補上
 * @param {Spreadsheet} ss - 試算表物件
 * @param {string} sheetName - 工作表名稱
 * @param {string} id - 目標資料 ID
 * @param {Object} data - 更新內容
 * @return {Object} 更新後的資料或錯誤訊息
 */
function updateRow(ss, sheetName, id, data) {
  const sheet = getOrCreateSheet(ss, sheetName);
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  
  let rowIndex = -1;
  const idIndex = headers.indexOf('id');
  const createdAtIndex = headers.indexOf('createdAt');
  
  // 1. 優先嘗試用 ID 找
  if (id && idIndex !== -1) {
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][idIndex]) === String(id)) { rowIndex = i + 1; break; }
    }
  }
  // 2. 如果沒 ID，嘗試用 CID1 找 (處理舊資料遷移)
  if (rowIndex === -1 && data.CID1) {
     const cidIndex = headers.indexOf('CID1');
     if (cidIndex !== -1) {
       for (let i = 1; i < allData.length; i++) {
         if (String(allData[i][cidIndex]) === String(data.CID1)) { rowIndex = i + 1; break; }
       }
     }
  }

  if (rowIndex === -1) return { status: 'error', message: 'Row not found' };

  // 若有新欄位，更新標頭
  const newHeaders = updateHeaders(sheet, headers, data);
  
  // 【自動補齊機制】
  // 檢查該列 id 欄位是否為空 (舊資料)
  if (idIndex !== -1) {
    const currentId = sheet.getRange(rowIndex, idIndex + 1).getValue();
    if (!currentId) {
      // 補上新的 ID
      const newId = data.id || Utilities.getUuid();
      sheet.getRange(rowIndex, idIndex + 1).setValue(newId);
      data.id = newId; // 更新 data 物件以便回傳給前端
    }
  }

  // 檢查 createdAt 欄位是否為空
  if (createdAtIndex !== -1) {
     const currentCreated = sheet.getRange(rowIndex, createdAtIndex + 1).getValue();
     if (!currentCreated) {
        sheet.getRange(rowIndex, createdAtIndex + 1).setValue(new Date().toISOString());
     }
  }

  // 逐一更新資料欄位
  for (const key in data) {
    const colIndex = newHeaders.indexOf(key);
    // 更新資料，套用格式化，且不隨意覆寫 id (除非上方補齊邏輯)
    if (colIndex !== -1 && key !== 'id') {
      sheet.getRange(rowIndex, colIndex + 1).setValue(formatValue(key, data[key]));
    }
  }
  
  return { status: 'success', message: 'Updated', data: data };
}

/**
 * 刪除資料 (Delete)
 * @param {Spreadsheet} ss - 試算表物件
 * @param {string} sheetName - 工作表名稱
 * @param {string} id - 要刪除的資料 ID
 * @return {Object} 執行結果
 */
function deleteRow(ss, sheetName, id) {
  const sheet = getOrCreateSheet(ss, sheetName);
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const idIndex = headers.indexOf('id');
  
  if (idIndex === -1) return { status: 'error', message: 'ID column not found' };
  
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][idIndex]) === String(id)) {
      sheet.deleteRow(i + 1); // 刪除整列
      return { status: 'success', message: 'Deleted' };
    }
  }
  return { status: 'error', message: 'ID not found' };
}

/**
 * 取得指定客戶的驗光紀錄筆數，並回寫至 CMF 表的 prxCount 欄位
 * @param {Spreadsheet} ss - 試算表物件
 * @param {string} cid1 - 客戶編號 (CID1)
 * @return {Object} 包含 prxCount 的結果
 */
function getPrxCount(ss, cid1) {
  if (!cid1) return { status: 'error', message: 'CID1 is required' };
  
  const targetCid = String(cid1).trim();
  let prxCount = 0;

  try {
    // 1. 計算 PRXMF 表中該客戶的筆數
    const prxSheet = ss.getSheetByName('PRXMF');
    if (prxSheet && prxSheet.getLastRow() > 1) {
      const prxHeaders = prxSheet.getRange(1, 1, 1, prxSheet.getLastColumn()).getValues()[0];
      const cidIndex = prxHeaders.indexOf('CID1');
      
      if (cidIndex !== -1) {
        const cidValues = prxSheet.getRange(2, cidIndex + 1, prxSheet.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < cidValues.length; i++) {
          if (String(cidValues[i][0]).trim() === targetCid) {
            prxCount++;
          }
        }
      }
    }

    // 2. 將 prxCount 回寫至 CMF 表
    const cmfSheet = ss.getSheetByName('CMF');
    if (cmfSheet && cmfSheet.getLastRow() > 1) {
      const cmfHeaders = cmfSheet.getRange(1, 1, 1, cmfSheet.getLastColumn()).getValues()[0];
      const cmfCidIndex = cmfHeaders.indexOf('CID1');
      let prxCountIndex = cmfHeaders.indexOf('prxCount');
      
      // 若 prxCount 欄位不存在，則新增
      if (prxCountIndex === -1) {
        prxCountIndex = cmfHeaders.length;
        cmfSheet.getRange(1, prxCountIndex + 1).setValue('prxCount');
      }
      
      // 找到該客戶的列並更新 prxCount
      if (cmfCidIndex !== -1) {
        const cmfData = cmfSheet.getRange(2, 1, cmfSheet.getLastRow() - 1, cmfSheet.getLastColumn()).getValues();
        for (let i = 0; i < cmfData.length; i++) {
          if (String(cmfData[i][cmfCidIndex]).trim() === targetCid) {
            cmfSheet.getRange(i + 2, prxCountIndex + 1).setValue(prxCount);
            break;
          }
        }
      }
    }

    return { status: 'success', cid1: targetCid, prxCount: prxCount };
  } catch (e) {
    Logger.log('getPrxCount error: ' + e);
    return { status: 'error', message: e.toString(), prxCount: prxCount };
  }
}

/**
 * 批次更新所有客戶的驗光紀錄筆數
 * 一次性計算 PRXMF 中每個 CID1 的數量，並批次回寫到 CMF 表的 prxCount 欄位
 * @param {Spreadsheet} ss - 試算表物件
 * @return {Object} 更新結果
 */
function updateAllPrxCount(ss) {
  try {
    const cmfSheet = ss.getSheetByName('CMF');
    const prxSheet = ss.getSheetByName('PRXMF');
    
    if (!cmfSheet || cmfSheet.getLastRow() <= 1) {
      return { status: 'success', message: 'No CMF data', updated: 0 };
    }
    
    // 1. 計算 PRXMF 中每個 CID1 的出現次數
    const prxCountMap = {};
    if (prxSheet && prxSheet.getLastRow() > 1) {
      const prxHeaders = prxSheet.getRange(1, 1, 1, prxSheet.getLastColumn()).getValues()[0];
      const prxCidIndex = prxHeaders.indexOf('CID1');
      
      if (prxCidIndex !== -1) {
        const prxCidValues = prxSheet.getRange(2, prxCidIndex + 1, prxSheet.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < prxCidValues.length; i++) {
          const cid = String(prxCidValues[i][0]).trim();
          if (cid) {
            prxCountMap[cid] = (prxCountMap[cid] || 0) + 1;
          }
        }
      }
    }
    
    // 2. 讀取 CMF 表頭
    const cmfHeaders = cmfSheet.getRange(1, 1, 1, cmfSheet.getLastColumn()).getValues()[0];
    const cmfCidIndex = cmfHeaders.indexOf('CID1');
    let prxCountIndex = cmfHeaders.indexOf('prxCount');
    
    // 若 prxCount 欄位不存在，則新增
    if (prxCountIndex === -1) {
      prxCountIndex = cmfHeaders.length;
      cmfSheet.getRange(1, prxCountIndex + 1).setValue('prxCount');
    }
    
    // 3. 讀取所有 CMF 資料
    const cmfData = cmfSheet.getRange(2, 1, cmfSheet.getLastRow() - 1, cmfSheet.getLastColumn()).getValues();
    
    // 4. 準備批次更新資料 (只需要 prxCount 欄位的值)
    const updateValues = cmfData.map(row => {
      const cid = String(row[cmfCidIndex]).trim();
      return [prxCountMap[cid] || 0];
    });
    
    // 5. 批次寫入 prxCount 欄位
    cmfSheet.getRange(2, prxCountIndex + 1, updateValues.length, 1).setValues(updateValues);
    
    return { 
      status: 'success', 
      message: 'All prxCount updated', 
      updated: updateValues.length,
      prxCountMap: prxCountMap // 回傳統計資料供前端使用
    };
  } catch (e) {
    Logger.log('updateAllPrxCount error: ' + e);
    return { status: 'error', message: e.toString() };
  }
}

// --- 輔助函式 (Helper Functions) ---

/**
 * 取得或建立工作表
 * 若工作表不存在則建立，並預設加入 id 與 createdAt 欄位
 */
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) { 
    sheet = ss.insertSheet(name); 
    sheet.appendRow(['id', 'createdAt']); // 初始化標題列
  }
  return sheet;
}

/**
 * 取得工作表目前的標題列
 */
function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

/**
 * 檢查並更新工作表標題列
 * 若 data 中有新欄位，則將其附加到標題列後方
 */
function updateHeaders(sheet, currentHeaders, newData) {
  const newKeys = Object.keys(newData);
  let updatedHeaders = [...currentHeaders];
  let isUpdated = false;
  
  for (const key of newKeys) {
    if (!updatedHeaders.includes(key)) { 
      updatedHeaders.push(key); 
      isUpdated = true; 
    }
  }
  
  if (isUpdated) {
    sheet.getRange(1, 1, 1, updatedHeaders.length).setValues([updatedHeaders]);
  }
  return updatedHeaders;
}