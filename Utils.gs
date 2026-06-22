/**
 * ============================================================
 * VanTran Mobile — Utils.gs
 * Các hàm tiện ích dùng chung cho toàn hệ thống
 * ============================================================
 */

// Cache lưu trữ dữ liệu của các sheet trong một lần chạy để tối ưu hiệu năng đọc
var _sheetDataCache = {};

function clearSheetCache(sheetName) {
  if (sheetName) {
    delete _sheetDataCache[sheetName];
  } else {
    _sheetDataCache = {};
  }
}

// ==================== CACHE SERVICE (PERSISTENT ACROSS REQUESTS) ====================

var DROPDOWN_CACHE_TTL = 3600; // 1 giờ

/**
 * Lưu dữ liệu vào CacheService với chunking (tránh giới hạn 100KB/key)
 *
 * @param {string} key - Cache key
 * @param {*} data - Dữ liệu cần cache (sẽ được JSON.stringify)
 */
function setChunkedCache(key, data) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(data);
    var CHUNK_SIZE = 90000; // ~90KB per chunk (buffer từ giới hạn 100KB)
    var numChunks = Math.ceil(json.length / CHUNK_SIZE);

    var cacheObj = {};
    cacheObj[key + '_m'] = JSON.stringify({ c: numChunks });

    for (var i = 0; i < numChunks; i++) {
      cacheObj[key + '_' + i] = json.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    }

    cache.putAll(cacheObj, DROPDOWN_CACHE_TTL);
  } catch (e) {
    Logger.log('CacheService setChunkedCache error: ' + e.message);
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
    var cache = CacheService.getScriptCache();
    var meta = cache.get(key + '_m');
    if (!meta) return null;

    var metaObj = JSON.parse(meta);
    var numChunks = metaObj.c;

    var keys = [];
    for (var i = 0; i < numChunks; i++) {
      keys.push(key + '_' + i);
    }

    var all = cache.getAll(keys);
    var json = '';
    for (var i = 0; i < numChunks; i++) {
      var chunk = all[key + '_' + i];
      if (!chunk) return null; // Partial cache miss
      json += chunk;
    }

    return JSON.parse(json);
  } catch (e) {
    Logger.log('CacheService getChunkedCache error: ' + e.message);
    return null;
  }
}

/**
 * Xóa cache dropdown liên quan đến sheet đã thay đổi
 *
 * @param {string} sheetName - Tên sheet đã thay đổi dữ liệu
 */
function invalidateDropdownCache(sheetName) {
  var cacheKeys = [];
  if (sheetName === SHEET_NAMES.KHACH_HANG) {
    cacheKeys = ['dd_kh'];
  } else if (sheetName === SHEET_NAMES.DIEN_THOAI) {
    cacheKeys = ['dd_dt'];
  } else if (sheetName === SHEET_NAMES.PHU_KIEN) {
    cacheKeys = ['dd_pk', 'dd_pku'];
  } else if (sheetName === SHEET_NAMES.TRA_GOP) {
    cacheKeys = ['dd_tg'];
  }

  if (cacheKeys.length === 0) return;

  try {
    var cache = CacheService.getScriptCache();
    for (var k = 0; k < cacheKeys.length; k++) {
      var key = cacheKeys[k];
      var meta = cache.get(key + '_m');
      if (meta) {
        var metaObj = JSON.parse(meta);
        var removeKeys = [key + '_m'];
        for (var i = 0; i < metaObj.c; i++) {
          removeKeys.push(key + '_' + i);
        }
        cache.removeAll(removeKeys);
      }
    }
  } catch (e) {
    Logger.log('CacheService invalidateDropdownCache error: ' + e.message);
  }
}


function getSheetDataCached(sheetName) {
  if (_sheetDataCache[sheetName]) {
    return _sheetDataCache[sheetName];
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
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
function generateId(prefix, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return prefix + "001";

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return prefix + "001";

  // Lấy cột mã (cột 1) và tìm số lớn nhất
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxNum = 0;

  data.forEach(function (row) {
    var val = String(row[0]);
    if (val.indexOf(prefix) === 0) {
      var numPart = parseInt(val.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  });

  var nextNum = maxNum + 1;
  var padded = ("000" + nextNum).slice(-3);
  return prefix + padded;
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
  var data = getSheetDataCached(sheetName);
  var searchStr = String(searchVal).trim().toLowerCase();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row.length < searchCol || row.length < returnCol) continue;
    var cellVal = String(row[searchCol - 1]).trim().toLowerCase();
    if (cellVal === searchStr) {
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
  var data = getSheetDataCached(sheetName);
  var searchStr = String(searchVal).trim().toLowerCase();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row.length < searchCol) continue;
    if (String(row[searchCol - 1]).trim().toLowerCase() === searchStr) {
      var result = {};
      returnCols.forEach(function (col) {
        result[col] = row.length >= col ? row[col - 1] : "";
      });
      result._rowIndex = i + 2; // Actual row in sheet (1-indexed, skip header)
      return result;
    }
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
  var dd = ("0" + date.getDate()).slice(-2);
  var mm = ("0" + (date.getMonth() + 1)).slice(-2);
  var yyyy = date.getFullYear();
  return dd + "/" + mm + "/" + yyyy;
}

/**
 * Đọc giá trị cấu hình từ sheet CauHinh
 *
 * @param {string} key - Tên cấu hình (cột A)
 * @return {string} Giá trị cấu hình (cột B) hoặc ''
 */
function getConfig(key) {
  var val = lookupValue(SHEET_NAMES.CAU_HINH, 1, key, 2);
  return val !== null ? String(val) : "";
}

/**
 * Đọc giá trị cấu hình dạng số
 *
 * @param {string} key - Tên cấu hình
 * @return {number} Giá trị số hoặc 0
 */
function getConfigNumber(key) {
  var val = getConfig(key);
  var num = parseFloat(val);
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
function appendRow(sheetName, rowData) {
  clearSheetCache(sheetName);
  invalidateDropdownCache(sheetName);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" không tồn tại!');

  // Tự động mở rộng số cột nếu dòng dữ liệu mới dài hơn số cột hiện tại
  var maxCols = sheet.getMaxColumns();
  if (maxCols < rowData.length) {
    sheet.insertColumnsAfter(maxCols, rowData.length - maxCols);
  }

  sheet.appendRow(rowData);
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(lastRow, 1, 1, rowData.length);
  range.setFontFamily("Times New Roman");
  range.setFontSize(12);

  // Auto-resize columns
  for (var i = 1; i <= rowData.length; i++) {
    sheet.autoResizeColumn(i);
  }

  return lastRow;
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" không tồn tại!');

  // Tự động mở rộng số cột nếu cột cần cập nhật vượt quá số cột hiện tại
  var maxCols = sheet.getMaxColumns();
  if (maxCols < col) {
    sheet.insertColumnsAfter(maxCols, col - maxCols);
  }

  var cell = sheet.getRange(row, col);
  cell.setValue(value);
  cell.setFontFamily("Times New Roman");
  cell.setFontSize(12);
  sheet.autoResizeColumn(col);
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
  var data = getSheetDataCached(sheetName);
  var searchStr = String(searchVal).trim().toLowerCase();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (row.length < searchCol) continue;
    if (String(row[searchCol - 1]).trim().toLowerCase() === searchStr) {
      return i + 2; // +2 vì skip header và 0-indexed
    }
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
    var ui = SpreadsheetApp.getUi();
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
  var now = new Date();
  var mm = ("0" + (now.getMonth() + 1)).slice(-2);
  return mm + "/" + now.getFullYear();
}

/**
 * Lấy danh sách các chi nhánh dưới dạng mảng chuỗi
 *
 * @return {string[]}
 */
function getBranchesList() {
  var branchesStr = getConfig("Danh sách chi nhánh");
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
  var list = getBranchesList();
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
  var listStr = getConfig("Danh sách công ty tài chính");
  if (!listStr) return ["FE Credit", "Home Credit", "HD Saison", "MIRAES Asset"];
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
  var list = getFinanceCompaniesList();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CAU_HINH);
  if (!sheet) return;
  var row = findRow(SHEET_NAMES.CAU_HINH, 1, key);
  if (row === -1) {
    var lastRow = sheet.getLastRow();
    var range = sheet.getRange(lastRow + 1, 1, 1, 3);
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
  ensureConfigKey("Danh sách thương hiệu", "Apple, Samsung, Xiaomi, OPPO, Vivo, Realme, Khác", "Danh sách các thương hiệu điện thoại, phân cách bằng dấu phẩy");
  var brandsStr = getConfig("Danh sách thương hiệu");
  if (!brandsStr) return ["Apple", "Samsung", "Xiaomi", "OPPO", "Vivo", "Realme", "Khác"];
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
  var list = getBrandsList();
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
  var str = String(hinhThucTT || "").trim();
  if (str === "Tiền mặt") {
    return { tm: totalAmount, ck: 0 };
  }
  if (str === "Chuyển khoản" || str === "Quẹt thẻ (POS)") {
    return { tm: 0, ck: totalAmount };
  }
  
  if (str.indexOf("+") !== -1 || str.indexOf(":") !== -1) {
    var tmMatch = str.match(/Tiền mặt:\s*([\d,]+)/i);
    var ckMatch = str.match(/Chuyển khoản:\s*([\d,]+)/i);
    var posMatch = str.match(/Quẹt thẻ\s*\(POS\):\s*([\d,]+)/i);
    
    var tmVal = tmMatch ? Number(tmMatch[1].replace(/,/g, '')) : 0;
    var ckVal = ckMatch ? Number(ckMatch[1].replace(/,/g, '')) : 0;
    var posVal = posMatch ? Number(posMatch[1].replace(/,/g, '')) : 0;
    
    var sum = tmVal + ckVal + posVal;
    if (sum > 0) {
      // Tính tỷ lệ để phân bổ chính xác theo totalAmount thực tế thu/chi
      var ratio = totalAmount / sum;
      return {
        tm: tmVal * ratio,
        ck: (ckVal + posVal) * ratio
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
  var str = String(val || "").trim();
  if (str.indexOf(",") !== -1) {
    var parts = str.split(",");
    if (parts.length === 2) {
      var ck = Number(parts[0]);
      var tm = Number(parts[1]);
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
  var parsed = parseHybridAmount(val);
  if (parsed) {
    return parsed.ck + parsed.tm;
  }
  return Number(val) || 0;
}
