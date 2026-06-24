/**
 * ============================================================
 * VanTran Mobile — Utils.gs
 * Các hàm tiện ích dùng chung cho toàn hệ thống
 * ============================================================
 */

// Cache lưu trữ dữ liệu của các sheet trong một lần chạy để tối ưu hiệu năng đọc
let _sheetDataCache = {};
let _sheetIndexCache = {};
let _holdsLock = false;

function clearSheetCache(sheetName) {
  if (sheetName) {
    delete _sheetDataCache[sheetName];
    delete _sheetIndexCache[sheetName];
    if (
      sheetName === SHEET_NAMES.PHU_KIEN &&
      typeof clearPhuKienCompositeIndex === "function"
    ) {
      clearPhuKienCompositeIndex();
    }
  } else {
    _sheetDataCache = {};
    _sheetIndexCache = {};
    if (typeof clearPhuKienCompositeIndex === "function") {
      clearPhuKienCompositeIndex();
    }
  }
}

// NOTE: Global column enums (COL_*), proxy (createEnumProxy), and init flags are now declared globally in SchemaDef.gs

function initializeColumnEnums() {
  if (_columnEnumsInitialized || _loadingColumnEnums) return;
  _loadingColumnEnums = true;

  const cache = CacheService.getScriptCache();
  const cached = cache.get("system_column_enums_cache");
  if (cached) {
    try {
      const data = JSON.parse(cached);
      Object.assign(COL_DT, data.COL_DT);
      Object.assign(COL_PK, data.COL_PK);
      Object.assign(COL_DH, data.COL_DH);
      Object.assign(COL_TM, data.COL_TM);
      Object.assign(COL_NK, data.COL_NK);
      Object.assign(COL_TG, data.COL_TG);
      Object.assign(COL_LSTG, data.COL_LSTG);
      Object.assign(COL_DV, data.COL_DV);
      Object.assign(COL_DT_TRA, data.COL_DT_TRA);
      Object.assign(COL_KH, data.COL_KH);
      Object.assign(COL_BH, data.COL_BH);
      Object.assign(COL_NV, data.COL_NV || {});
      _columnEnumsInitialized = true;
      _loadingColumnEnums = false;
      return;
    } catch (e) {
      Logger.log("Error parsing cached column enums: " + e.message);
    }
  }

  try {
    reconcileAllSchemas();
    _columnEnumsInitialized = true;
  } finally {
    _loadingColumnEnums = false;
  }
}

function reconcileAllSchemas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Mapping of sheet name to its corresponding global COL_XXX object
  const enumMapping = {};
  enumMapping[SHEET_NAMES.DIEN_THOAI] = COL_DT;
  enumMapping[SHEET_NAMES.PHU_KIEN] = COL_PK;
  enumMapping[SHEET_NAMES.DON_HANG] = COL_DH;
  enumMapping[SHEET_NAMES.THU_MUA] = COL_TM;
  enumMapping[SHEET_NAMES.NHAP_KHO] = COL_NK;
  enumMapping[SHEET_NAMES.TRA_GOP] = COL_TG;
  enumMapping[SHEET_NAMES.LICH_SU_TRA_GOP] = COL_LSTG;
  enumMapping[SHEET_NAMES.DICH_VU] = COL_DV;
  enumMapping[SHEET_NAMES.DOI_TRA] = COL_DT_TRA;
  enumMapping[SHEET_NAMES.KHACH_HANG] = COL_KH;
  enumMapping[SHEET_NAMES.BAO_HANH] = COL_BH;
  enumMapping[SHEET_NAMES.NHAN_VIEN] = COL_NV;

  for (const sheetName in SCHEMA) {
    const colEnum = enumMapping[sheetName];
    if (!colEnum) continue;

    // Load actual column map from the sheet
    const actualMap = getColMapFromSheet(sheetName);
    const schemaList = SCHEMA[sheetName];

    if (actualMap && Object.keys(actualMap).length > 0) {
      // Reconcile index based on current sheet header label
      schemaList.forEach(function (colDef) {
        const code = colDef[0];
        const label = colDef[1].toLowerCase();
        
        if (actualMap[label] !== undefined) {
          colEnum[code] = actualMap[label];
        }
      });
    }
  }
  
  // Reconstruct exact property structure to preserve compatibility with existing cache consumer logic
  const cachePayload = {
    COL_DT: COL_DT,
    COL_PK: COL_PK,
    COL_DH: COL_DH,
    COL_TM: COL_TM,
    COL_NK: COL_NK,
    COL_TG: COL_TG,
    COL_LSTG: COL_LSTG,
    COL_DV: COL_DV,
    COL_DT_TRA: COL_DT_TRA,
    COL_KH: COL_KH,
    COL_BH: COL_BH,
    COL_NV: COL_NV
  };

  const cache = CacheService.getScriptCache();
  try {
    cache.put("system_column_enums_cache", JSON.stringify(cachePayload), 21600);
  } catch (e) {
    Logger.log("Error writing column enums cache: " + e.message);
  }
}

/**
 * Xóa cache column enums của hệ thống
 */
function clearColumnEnumsCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove("system_column_enums_cache");
    _columnEnumsInitialized = false;
  } catch (e) {
    Logger.log("Error clearing column enums cache: " + e.message);
  }
}

/**
 * Lấy bản đồ cột (column mapping) của một sheet.
 * Chỉ đọc tiêu đề hiện có từ sheet, không ghi/sửa đổi cấu trúc cột.
 * 
 * @param {string} sheetName Tên sheet cần lấy bản đồ cột.
 * @return {Object|null} Bản đồ cột (chuyển chữ thường) hoặc null nếu lỗi/không tìm thấy sheet.
 */
function getColMapFromSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return null;

    const lastCol = sheet.getLastColumn();
    if (lastCol <= 0) return {};

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const map = {};
    for (let i = 0; i < headers.length; i++) {
      map[String(headers[i]).trim().toLowerCase()] = i + 1;
    }
    return map;
  } catch (e) {
    Logger.log("Error getColMapFromSheet for " + sheetName + ": " + e.message);
    return null;
  }
}

/**
 * Đảm bảo các cột định nghĩa trong SHEET_HEADERS tồn tại trong sheet.
 * Chỉ chạy khi cần migration hoặc cài đặt ban đầu.
 * 
 * @param {string} sheetName Tên sheet cần kiểm tra và cập nhật cấu trúc.
 */
function ensureSheetColumnsExist(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    const expected = SHEET_HEADERS[sheetName];
    if (!sheet || !expected) return;

    const current = sheet.getLastColumn();
    if (expected.length > current) {
      const missing = expected.slice(current);
      sheet.getRange(1, current + 1, 1, missing.length).setValues([missing]);
    }
  } catch (e) {
    Logger.log("Error ensureSheetColumnsExist for " + sheetName + ": " + e.message);
  }
}

function getSheetIndex(sheetName, searchCol) {
  if (!_sheetIndexCache[sheetName]) {
    _sheetIndexCache[sheetName] = {};
  }
  if (!_sheetIndexCache[sheetName][searchCol]) {
    const indexMap = {};
    const data = getSheetDataCached(sheetName);
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.length < searchCol) continue;
      const rawVal = row[searchCol - 1];
      let cellVal = "";
      if (sheetName === SHEET_NAMES.DOANH_SO && searchCol === 1) {
        cellVal = formatMonthYear(rawVal);
      } else {
        cellVal = String(rawVal).trim();
      }
      cellVal = cellVal.toLowerCase();
      if (!(cellVal in indexMap)) {
        indexMap[cellVal] = i; // Lưu index 0-indexed trong mảng data
      }
    }
    _sheetIndexCache[sheetName][searchCol] = indexMap;
  }
  return _sheetIndexCache[sheetName][searchCol];
}

// ==================== CACHE SERVICE (PERSISTENT ACROSS REQUESTS) ====================

const DROPDOWN_CACHE_TTL = 3600; // 1 giờ

/**
 * Lưu dữ liệu vào CacheService với chunking (tránh giới hạn 100KB/key)
 *
 * @param {string} key - Cache key
 * @param {*} data - Dữ liệu cần cache (sẽ được JSON.stringify)
 */
function setChunkedCache(key, data) {
  try {
    const cache = CacheService.getScriptCache();
    const json = JSON.stringify(data);
    const CHUNK_SIZE = 90000; // ~90KB per chunk (buffer từ giới hạn 100KB)
    const numChunks = Math.ceil(json.length / CHUNK_SIZE);

    const cacheObj = {};
    cacheObj[key + "_m"] = JSON.stringify({ c: numChunks });

    for (let i = 0; i < numChunks; i++) {
      cacheObj[key + "_" + i] = json.substring(
        i * CHUNK_SIZE,
        (i + 1) * CHUNK_SIZE,
      );
    }

    cache.putAll(cacheObj, DROPDOWN_CACHE_TTL);
  } catch (e) {
    Logger.log("CacheService setChunkedCache error: " + e.message);
  }
}

/**
 * Đọc dữ liệu đã chunk từ CacheService
 *
 * @param {string} key - Cache key
 * @return {*|null} Dữ liệu đã parse hoặc null nếu cache miss
 */
function getChunkedCache(key) {
  try {
    const cache = CacheService.getScriptCache();
    const meta = cache.get(key + "_m");
    if (!meta) return null;

    const metaObj = JSON.parse(meta);
    const numChunks = metaObj.c;

    const keys = [];
    for (let i = 0; i < numChunks; i++) {
      keys.push(key + "_" + i);
    }

    const all = cache.getAll(keys);
    let json = "";
    for (let i = 0; i < numChunks; i++) {
      const chunk = all[key + "_" + i];
      if (!chunk) return null; // Partial cache miss
      json += chunk;
    }

    return JSON.parse(json);
  } catch (e) {
    Logger.log("CacheService getChunkedCache error: " + e.message);
    return null;
  }
}

/**
 * Xóa cache dropdown liên quan đến sheet đã thay đổi
 *
 * @param {string} sheetName - Tên sheet đã thay đổi dữ liệu
 */
function invalidateDropdownCache(sheetName) {
  let cacheKeys = [];
  if (sheetName === SHEET_NAMES.KHACH_HANG) {
    cacheKeys = ["dd_kh"];
  } else if (sheetName === SHEET_NAMES.DIEN_THOAI) {
    cacheKeys = ["dd_dt"];
  } else if (sheetName === SHEET_NAMES.PHU_KIEN) {
    cacheKeys = ["dd_pk", "dd_pku"];
  } else if (sheetName === SHEET_NAMES.TRA_GOP) {
    cacheKeys = ["dd_tg"];
  }

  if (cacheKeys.length === 0) return;

  try {
    const cache = CacheService.getScriptCache();
    for (let k = 0; k < cacheKeys.length; k++) {
      const key = cacheKeys[k];
      const meta = cache.get(key + "_m");
      if (meta) {
        const metaObj = JSON.parse(meta);
        const removeKeys = [key + "_m"];
        for (let i = 0; i < metaObj.c; i++) {
          removeKeys.push(key + "_" + i);
        }
        cache.removeAll(removeKeys);
      }
    }
  } catch (e) {
    Logger.log("CacheService invalidateDropdownCache error: " + e.message);
  }
}

function getSheetDataCached(sheetName) {
  if (_sheetDataCache[sheetName]) {
    return _sheetDataCache[sheetName];
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  _sheetDataCache[sheetName] = data;
  return data;
}

/**
 * Tạo mã tự động theo prefix + số thứ tự
 * VD: generateId('DH', 'DonHang') → 'DH001', 'DH002', ...
 *
 * @param {string} prefix - Tiền tố mã (VD: 'DH', 'NV', 'KH')
 * @param {string} sheetName - Tên sheet để đếm số dòng
 * @return {string} Mã mới
 */
function getNewIdCounter(prefix, sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0;

  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let maxNum = 0;

  data.forEach(function (row) {
    const val = String(row[0]).trim();
    if (val.indexOf(prefix) === 0) {
      const numPart = parseInt(val.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  });

  return maxNum;
}

function generateId(prefix, sheetName) {
  const lock = LockService.getScriptLock();
  // Chờ tối đa 5 giây nếu có người khác đang tạo ID
  lock.waitLock(5000);
  try {
    const props = PropertiesService.getScriptProperties();
    const key = "id_counter_" + prefix + "_" + sheetName;
    // Đọc counter cũ, nếu chưa có thì gán là 0, sau đó cộng 1
    const count = Number(props.getProperty(key) || "0") + 1;
    // Lưu counter mới lại ngay lập tức
    props.setProperty(key, String(count));
    // Định dạng số thành 5 chữ số (00001, 00002...)
    const padded = ("00000" + count).slice(-5);
    // Tạo ID kết hợp Counter và Random (Cần đảm bảo hàm getRandomLetters() đã được định nghĩa)
    return prefix + padded; 
  } finally {
    // Luôn luôn giải phóng khóa dù có lỗi xảy ra hay không
    lock.releaseLock(); 
  }
}
/**
 * Lookup giá trị từ sheet
 *
 * @param {string} sheetName - Tên sheet
 * @param {number} searchCol - Cột tìm kiếm (1-indexed)
 * @param {*} searchVal - Giá trị tìm
 * @param {number} returnCol - Cột trả về (1-indexed)
 * @return {*} Giá trị tìm được hoặc null
 */
function lookupValue(sheetName, searchCol, searchVal, returnCol) {
  const indexMap = getSheetIndex(sheetName, searchCol);
  const searchStr = String(searchVal).trim().toLowerCase();
  if (searchStr in indexMap) {
    const data = getSheetDataCached(sheetName);
    const rowIndex = indexMap[searchStr];
    const row = data[rowIndex];
    if (row.length >= returnCol) {
      return row[returnCol - 1];
    }
  }
  return null;
}

/**
 * Lookup nhiều cột cùng lúc từ 1 dòng
 *
 * @param {string} sheetName - Tên sheet
 * @param {number} searchCol - Cột tìm kiếm (1-indexed)
 * @param {*} searchVal - Giá trị tìm
 * @param {number[]} returnCols - Mảng các cột cần trả về (1-indexed)
 * @return {Object|null} Object với key là index cột, value là giá trị
 */
function lookupMultipleValues(sheetName, searchCol, searchVal, returnCols) {
  const indexMap = getSheetIndex(sheetName, searchCol);
  const searchStr = String(searchVal).trim().toLowerCase();
  if (searchStr in indexMap) {
    const data = getSheetDataCached(sheetName);
    const rowIndex = indexMap[searchStr];
    const row = data[rowIndex];
    const result = {};
    returnCols.forEach(function (col) {
      result[col] = row.length >= col ? row[col - 1] : "";
    });
    result._rowIndex = rowIndex + 2; // Actual row in sheet (1-indexed, skip header)
    return result;
  }
  return null;
}

/**
 * Định dạng số tiền VND
 *
 * @param {number} number - Số tiền
 * @return {string} Chuỗi đã format: "1,234,567"
 */
function formatCurrency(number) {
  if (isNaN(number) || number === null || number === undefined) return "0";
  return Number(number).toLocaleString("vi-VN");
}

/**
 * Định dạng ngày dd/MM/yyyy
 *
 * @param {Date} date - Đối tượng Date
 * @return {string} Chuỗi ngày "dd/MM/yyyy"
 */
function formatDate(date) {
  if (!date || !(date instanceof Date)) return "";
  const dd = ("0" + date.getDate()).slice(-2);
  const mm = ("0" + (date.getMonth() + 1)).slice(-2);
  const yyyy = date.getFullYear();
  return dd + "/" + mm + "/" + yyyy;
}

/**
 * Chuyển đổi cell value (Date hoặc String) thành định dạng MM/yyyy
 *
 * @param {*} val - Giá trị cell
 * @return {string} Chuỗi định dạng "MM/yyyy"
 */
function formatMonthYear(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const mm = ("0" + (val.getMonth() + 1)).slice(-2);
    return mm + "/" + val.getFullYear();
  }
  const str = String(val).trim();
  if (str.length > 7 && !isNaN(Date.parse(str))) {
    const d = new Date(str);
    const mm = ("0" + (d.getMonth() + 1)).slice(-2);
    return mm + "/" + d.getFullYear();
  }
  return str;
}

/**
 * Đọc giá trị cấu hình từ sheet CauHinh
 *
 * @param {string} key - Tên cấu hình (cột A)
 * @return {string} Giá trị cấu hình (cột B) hoặc ''
 */
function getConfig(key) {
  const val = lookupValue(SHEET_NAMES.CAU_HINH, 1, key, 2);
  return val !== null ? String(val) : "";
}

/**
 * Đọc giá trị cấu hình dạng số
 *
 * @param {string} key - Tên cấu hình
 * @return {number} Giá trị số hoặc 0
 */
function getConfigNumber(key) {
  const val = getConfig(key);
  const num = parseFloat(val);
  return isNaN(num) ? 0 : num;
}

/**
 * Lấy toàn bộ dữ liệu từ sheet (trừ header)
 *
 * @param {string} sheetName - Tên sheet
 * @return {Array[]} Mảng 2 chiều dữ liệu
 */
function getAllData(sheetName) {
  return getSheetDataCached(sheetName);
}

/**
 * Thêm 1 dòng dữ liệu vào cuối sheet
 *
 * @param {string} sheetName - Tên sheet
 * @param {Array} rowData - Mảng dữ liệu 1 dòng
 * @return {number} Số dòng mới (row number)
 */
/**
 * Hàm cập nhật hoặc chèn mới dòng dữ liệu an toàn (DRY Compliance)
 *
 * @param {string} sheetName - Tên sheet
 * @param {number} row - Dòng cần sửa (1-indexed), nếu là -1 sẽ chèn mới vào cuối sheet
 * @param {Object} columnValueMap - Bản đồ cột và giá trị { [Mã cột enum]: giá trị }
 * @return {number} Số dòng đã ghi nhận (row number)
 */
function saveRowData(sheetName, row, columnValueMap) {
  clearSheetCache(sheetName);
  invalidateDropdownCache(sheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" không tồn tại!');

  // Xác định số cột lớn nhất cần thiết
  let maxColNeeded = 0;
  Object.keys(columnValueMap).forEach(function(colStr) {
    const col = parseInt(colStr, 10);
    if (!isNaN(col) && col > maxColNeeded) {
      maxColNeeded = col;
    }
  });

  // Tự động mở rộng số cột nếu thiếu
  const maxCols = sheet.getMaxColumns();
  if (maxCols < maxColNeeded) {
    sheet.insertColumnsAfter(maxCols, maxColNeeded - maxCols);
  }

  // Nếu là dòng mới (-1), append dòng trống để lấy số dòng cuối
  if (row === -1) {
    sheet.appendRow([]);
    row = sheet.getLastRow();
  }

  const range = sheet.getRange(row, 1, 1, maxColNeeded);
  const rowValues = range.getValues()[0];

  // Áp dụng dữ liệu cập nhật
  Object.keys(columnValueMap).forEach(function(colStr) {
    const col = parseInt(colStr, 10);
    if (isNaN(col)) return;
    let value = columnValueMap[colStr];

    // Tránh mất số 0 ở đầu đối với chuỗi số điện thoại hoặc mã bắt đầu bằng 0
    if (typeof value === "string" && value.indexOf("0") === 0 && value.length > 1 && /^\d+$/.test(value)) {
      value = "'" + value;
    }
    rowValues[col - 1] = (value !== undefined && value !== null) ? value : "";
  });

  range.setValues([rowValues]);
  range.setFontFamily("Times New Roman");
  range.setFontSize(12);

  return row;
}

function appendRow(sheetName, rowData) {
  clearSheetCache(sheetName);
  invalidateDropdownCache(sheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" không tồn tại!');

  // Tự động mở rộng số cột nếu dòng dữ liệu mới dài hơn số cột hiện tại
  const maxCols = sheet.getMaxColumns();
  if (maxCols < rowData.length) {
    sheet.insertColumnsAfter(maxCols, rowData.length - maxCols);
  }

  // Tránh mất số 0 ở đầu đối với chuỗi số điện thoại hoặc mã bắt đầu bằng 0
  for (let i = 0; i < rowData.length; i++) {
    const val = rowData[i];
    if (typeof val === "string" && val.indexOf("0") === 0 && val.length > 1 && /^\d+$/.test(val)) {
      rowData[i] = "'" + val;
    }
  }

  sheet.appendRow(rowData);
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(lastRow, 1, 1, rowData.length);
  range.setFontFamily("Times New Roman");
  range.setFontSize(12);

  return lastRow;
}

function appendRows(sheetName, rowsData) {
  if (!rowsData || rowsData.length === 0) return 0;
  clearSheetCache(sheetName);
  invalidateDropdownCache(sheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" không tồn tại!');

  // Tìm số cột dữ liệu lớn nhất trong các dòng
  const maxDataCols = Math.max.apply(null, rowsData.map(function(r) { return r.length; }));
  const maxCols = sheet.getMaxColumns();
  if (maxCols < maxDataCols) {
    sheet.insertColumnsAfter(maxCols, maxDataCols - maxCols);
  }

  // Tránh mất số 0 ở đầu đối với chuỗi số điện thoại hoặc mã bắt đầu bằng 0
  rowsData.forEach(function (rowData) {
    for (let i = 0; i < rowData.length; i++) {
      const val = rowData[i];
      if (typeof val === "string" && val.indexOf("0") === 0 && val.length > 1 && /^\d+$/.test(val)) {
        rowData[i] = "'" + val;
      }
    }
  });

  const lastRow = sheet.getLastRow();
  // Pad các hàng ngắn để khớp với kích thước mảng 2 chiều
  const normalizedRows = rowsData.map(function(r) {
    const arr = r.slice();
    while (arr.length < maxDataCols) {
      arr.push("");
    }
    return arr;
  });

  const range = sheet.getRange(lastRow + 1, 1, normalizedRows.length, maxDataCols);
  range.setValues(normalizedRows);
  range.setFontFamily("Times New Roman");
  range.setFontSize(12);

  return lastRow + normalizedRows.length;
}

/**
 * Cập nhật giá trị 1 cell
 *
 * @param {string} sheetName - Tên sheet
 * @param {number} row - Dòng (1-indexed)
 * @param {number} col - Cột (1-indexed)
 * @param {*} value - Giá trị mới
 */
function updateCell(sheetName, row, col, value) {
  clearSheetCache(sheetName);
  invalidateDropdownCache(sheetName);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" không tồn tại!');

  // Tự động mở rộng số cột nếu cột cần cập nhật vượt quá số cột hiện tại
  const maxCols = sheet.getMaxColumns();
  if (maxCols < col) {
    sheet.insertColumnsAfter(maxCols, col - maxCols);
  }

  const cell = sheet.getRange(row, col);
  // Tránh mất số 0 ở đầu đối với chuỗi số điện thoại hoặc mã bắt đầu bằng 0
  if (typeof value === "string" && value.indexOf("0") === 0 && value.length > 1 && /^\d+$/.test(value)) {
    value = "'" + value;
  }
  cell.setValue(value);
  cell.setFontFamily("Times New Roman");
  cell.setFontSize(12);
}

/**
 * Tìm dòng theo giá trị cột
 *
 * @param {string} sheetName - Tên sheet
 * @param {number} searchCol - Cột tìm (1-indexed)
 * @param {*} searchVal - Giá trị tìm
 * @return {number} Row number (1-indexed) hoặc -1 nếu không tìm thấy
 */
function findRow(sheetName, searchCol, searchVal) {
  const indexMap = getSheetIndex(sheetName, searchCol);
  const searchStr = String(searchVal).trim().toLowerCase();
  if (searchStr in indexMap) {
    return indexMap[searchStr] + 2; // +2 vì skip header và 0-indexed
  }
  return -1;
}

/**
 * Hiển thị thông báo toast
 *
 * @param {string} message - Nội dung thông báo
 * @param {string} title - Tiêu đề (optional)
 * @param {number} duration - Thời gian hiển thị (giây, default 5)
 */
function showToast(message, title, duration) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      message,
      title || "VanTran Mobile",
      duration || 5,
    );
  } catch (e) {
    Logger.log("Toast: [" + (title || "VanTran Mobile") + "] " + message);
  }
}

/**
 * Hiển thị alert
 *
 * @param {string} title - Tiêu đề
 * @param {string} message - Nội dung
 */
function showAlert(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui) {
      ui.alert(title, message, ui.ButtonSet.OK);
      return;
    }
  } catch (e) {
    Logger.log("Alert: [" + title + "] " + message);
  }
}

/**
 * Kiểm tra thương hiệu có phải Apple không
 *
 * @param {string} thuongHieu - Tên thương hiệu
 * @return {boolean}
 */
function isApple(thuongHieu) {
  return String(thuongHieu).trim().toLowerCase() === "apple";
}

/**
 * Lấy tháng/năm hiện tại dạng MM/YYYY
 *
 * @return {string} VD: "06/2026"
 */
function getCurrentMonthYear() {
  const now = new Date();
  const mm = ("0" + (now.getMonth() + 1)).slice(-2);
  return mm + "/" + now.getFullYear();
}

/**
 * Lấy danh sách các chi nhánh dưới dạng mảng chuỗi
 *
 * @return {string[]}
 */
function getBranchesList() {
  const branchesStr = getConfig("Danh sách chi nhánh");
  if (!branchesStr) return ["Chi nhánh 1", "Chi nhánh 2", "Chi nhánh 3"];
  return branchesStr.split(",").map(function (item) {
    return item.trim();
  });
}

/**
 * Lấy danh sách các chi nhánh dưới dạng cấu trúc dropdown cho UI
 *
 * @return {Object[]}
 */
function getBranchesDropdown() {
  const list = getBranchesList();
  return list.map(function (item) {
    return { value: item, text: item };
  });
}

/**
 * Lấy danh sách các công ty tài chính dưới dạng mảng chuỗi
 *
 * @return {string[]}
 */
function getFinanceCompaniesList() {
  const listStr = getConfig("Danh sách công ty tài chính");
  if (!listStr)
    return ["FE Credit", "Home Credit", "HD Saison", "MIRAES Asset"];
  return listStr.split(",").map(function (item) {
    return item.trim();
  });
}

/**
 * Lấy danh sách các công ty tài chính dưới dạng cấu trúc dropdown cho UI
 *
 * @return {Object[]}
 */
function getFinanceCompaniesDropdown() {
  const list = getFinanceCompaniesList();
  return list.map(function (item) {
    return { value: item, text: item };
  });
}

/**
 * Đảm bảo khoá cấu hình tồn tại trong sheet CauHinh
 * Nếu chưa tồn tại, tự động chèn thêm dòng mới với giá trị mặc định và mô tả
 *
 * @param {string} key - Tên cấu hình
 * @param {string} defaultValue - Giá trị mặc định
 * @param {string} description - Mô tả cấu hình
 */
function ensureConfigKey(key, defaultValue, description) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CAU_HINH);
  if (!sheet) return;
  const row = findRow(SHEET_NAMES.CAU_HINH, 1, key);
  if (row === -1) {
    const lastRow = sheet.getLastRow();
    const range = sheet.getRange(lastRow + 1, 1, 1, 3);
    range.setValues([[key, defaultValue, description]]);
    range.setFontFamily("Times New Roman").setFontSize(12);
  }
}

/**
 * Lấy danh sách các thương hiệu điện thoại dưới dạng mảng chuỗi
 *
 * @return {string[]}
 */
function getBrandsList() {
  ensureConfigKey(
    "Danh sách thương hiệu",
    "Apple, Samsung, Xiaomi, OPPO, Vivo, Realme, Khác",
    "Danh sách các thương hiệu điện thoại, phân cách bằng dấu phẩy",
  );
  const brandsStr = getConfig("Danh sách thương hiệu");
  if (!brandsStr)
    return ["Apple", "Samsung", "Xiaomi", "OPPO", "Vivo", "Realme", "Khác"];
  return brandsStr.split(",").map(function (item) {
    return item.trim();
  });
}

/**
 * Lấy danh sách các thương hiệu điện thoại dưới dạng cấu trúc dropdown cho UI
 *
 * @return {Object[]}
 */
function getBrandsDropdown() {
  const list = getBrandsList();
  return list.map(function (item) {
    return { value: item, text: item };
  });
}

/**
 * Phân tích hình thức thanh toán hỗn hợp (ví dụ: "Tiền mặt: 500,000 + Chuyển khoản: 1,000,000")
 *
 * @param {string} hinhThucTT - Chuỗi hình thức thanh toán
 * @param {number} totalAmount - Tổng số tiền thu/chi thực tế của bản ghi (amountCollected, hoanTien, v.v.)
 * @return {Object} { tm: number, ck: number } số tiền mặt và chuyển khoản (bao gồm POS)
 */
function parseMixedPayment(hinhThucTT, totalAmount) {
  const str = String(hinhThucTT || "").trim();
  if (str === "Tiền mặt") {
    return { tm: totalAmount, ck: 0 };
  }
  if (str === "Chuyển khoản" || str === "Quẹt thẻ (POS)") {
    return { tm: 0, ck: totalAmount };
  }

  if (str.indexOf("+") !== -1 || str.indexOf(":") !== -1) {
    const tmMatch = str.match(/Tiền mặt:\s*([\d,]+)/i);
    const ckMatch = str.match(/Chuyển khoản:\s*([\d,]+)/i);
    const posMatch = str.match(/Quẹt thẻ\s*\(POS\):\s*([\d,]+)/i);

    const tmVal = tmMatch ? Number(tmMatch[1].replace(/,/g, "")) : 0;
    const ckVal = ckMatch ? Number(ckMatch[1].replace(/,/g, "")) : 0;
    const posVal = posMatch ? Number(posMatch[1].replace(/,/g, "")) : 0;

    const sum = tmVal + ckVal + posVal;
    if (sum > 0) {
      // Tính tỷ lệ để phân bổ chính xác theo totalAmount thực tế thu/chi
      const ratio = totalAmount / sum;
      return {
        tm: Math.round(tmVal * ratio),
        ck: Math.round((ckVal + posVal) * ratio),
      };
    }
  }

  // Mặc định chuyển khoản nếu không khớp
  return { tm: 0, ck: totalAmount };
}

/**
 * Phân tích chuỗi số tiền hỗn hợp dạng "chuyenKhoan,tienMat"
 * @param {any} val - Giá trị ô hoặc chuỗi
 * @return {Object|null} { ck: number, tm: number } hoặc null nếu không phải định dạng hỗn hợp
 */
function parseHybridAmount(val) {
  const str = String(val || "").trim();
  if (str.indexOf(",") !== -1) {
    const parts = str.split(",");
    if (parts.length === 2) {
      const ck = Number(parts[0]);
      const tm = Number(parts[1]);
      if (!isNaN(ck) && !isNaN(tm)) {
        return { ck: ck, tm: tm };
      }
    }
  }
  return null;
}

/**
 * Lấy tổng số tiền (cho cả trường hợp số đơn hoặc chuỗi hỗn hợp)
 * @param {any} val
 * @return {number}
 */
function parseAmountVal(val) {
  if (val === undefined || val === null || val === "") return 0;
  const parsed = parseHybridAmount(val);
  if (parsed) {
    return parsed.ck + parsed.tm;
  }
  return Number(val) || 0;
}

/**
 * Xoá toàn bộ bộ nhớ đệm (cache) hệ thống ở máy chủ
 */
function clearAllCaches() {
  clearColumnEnumsCache();
  clearSheetCache();

  // Xoá tất cả cache của dropdown
  const cache = CacheService.getScriptCache();
  const keys = ["dd_kh", "dd_dt", "dd_pk", "dd_pku", "dd_tg"];
  keys.forEach(function (key) {
    const meta = cache.get(key + "_m");
    if (meta) {
      try {
        const metaObj = JSON.parse(meta);
        const removeKeys = [key + "_m"];
        for (let i = 0; i < metaObj.c; i++) {
          removeKeys.push(key + "_" + i);
        }
        cache.removeAll(removeKeys);
      } catch (ex) {
        cache.remove(key + "_m");
      }
    }
  });
  return true;
}

/**
 * Lấy phần trăm lãi suất trả góp cửa hàng từ cấu hình
 * @return {number} Phần trăm lãi suất (VD: 2 nếu là 2%)
 */
function getInterestRateConfig() {
  const val = getConfig("Lãi suất trả góp cửa hàng (%)");
  if (!val) return 0;
  const cleanVal = val.replace("%", "").trim();
  const num = parseFloat(cleanVal);
  return isNaN(num) ? 0 : num;
}

/**
 * Thực thi một hàm an sau trong môi trường khóa tài liệu (Document Lock)
 * 
 * @param {Function} callback Hàm chứa logic nghiệp vụ cần thực hiện trong lock
 * @return {*} Kết quả trả về của callback
 */
function withDocumentLock(callback) {
  if (_holdsLock) {
    return callback();
  }
  const lock = LockService.getDocumentLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    throw new Error("Hệ thống hiện đang bận xử lý giao dịch khác. Vui lòng thử lại sau vài giây!");
  }
  
  _holdsLock = true;
  try {
    return callback();
  } finally {
    _holdsLock = false;
    try {
      lock.releaseLock();
    } catch (err) {
      Logger.log("Không thể giải phóng khóa: " + err.message);
    }
  }
}

/**
 * Tính toán và xác thực phân bổ thanh toán (Tiền mặt / Chuyển khoản / Hỗn hợp)
 * 
 * @param {Object} data Đối tượng chứa hinhThucThanhToan, splitTienMat, splitChuyenKhoan
 * @param {number} totalAmount Tổng số tiền cần thanh toán
 * @return {Object} { tienMat, chuyenKhoan, hinhThucTTDisplay }
 */
function calculatePaymentSplit(data, totalAmount) {
  let tienMat = 0;
  let chuyenKhoan = 0;
  const hinhThucTTDisplay = data.hinhThucThanhToan || "Tiền mặt";
  
  if (data.hinhThucThanhToan === "Hỗn hợp") {
    const splitTienMat = Number(data.splitTienMat) || 0;
    const splitChuyenKhoan = Number(data.splitChuyenKhoan) || 0;
    if (Math.abs(splitTienMat + splitChuyenKhoan - totalAmount) > 1) {
      throw new Error(
        "Lỗi dữ liệu: Tổng tiền mặt (" +
          splitTienMat +
          ") và chuyển khoản (" +
          splitChuyenKhoan +
          ") không khớp với số tiền cần thanh toán (" +
          totalAmount +
          ")!"
      );
    }
    tienMat = splitTienMat;
    chuyenKhoan = splitChuyenKhoan;
  } else if (data.hinhThucThanhToan === "Tiền mặt") {
    tienMat = totalAmount;
  } else {
    chuyenKhoan = totalAmount;
  }
  
  return {
    tienMat: tienMat,
    chuyenKhoan: chuyenKhoan,
    hinhThucTTDisplay: hinhThucTTDisplay
  };
}

/**
 * Kiểm tra xem ngày tháng có khớp với tháng và năm chỉ định không
 * 
 * @param {*} date Giá trị ngày cần kiểm tra
 * @param {number} month Tháng (1-12)
 * @param {number} year Năm
 * @return {boolean}
 */
function isSameMonthYear(date, month, year) {
  return (
    date instanceof Date &&
    date.getMonth() + 1 === month &&
    date.getFullYear() === year
  );
}
