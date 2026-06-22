/**
 * ============================================================
 * VanTran Mobile — KhachHang.gs
 * Quản lý thông tin khách hàng
 * ============================================================
 */

/**
 * Lấy danh sách tất cả khách hàng
 *
 * @return {Object[]}
 */
function getAllKhachHang() {
  initializeColumnEnums();
  var data = getAllData(SHEET_NAMES.KHACH_HANG);
  var headers = SHEET_HEADERS[SHEET_NAMES.KHACH_HANG];
  var result = [];

  data.forEach(function (row) {
    var obj = {};
    headers.forEach(function (h, i) {
      obj[h] = row[i];
    });
    result.push(obj);
  });

  return result;
}

/**
 * Lấy danh sách KH cho dropdown
 *
 * @return {Object[]} [{value: 'KH001', text: 'KH001 - Nguyễn Văn A (0901234567)'}, ...]
 */
function getKhachHangDropdown() {
  initializeColumnEnums();
  var data = getAllData(SHEET_NAMES.KHACH_HANG);
  var result = [];

  data.forEach(function (row) {
    var ma = String(row[COL_KH.MA_KH - 1] || "");
    if (ma.trim() !== "") {
      var ten = String(row[COL_KH.HO_TEN - 1] || "");
      var sdt = String(row[COL_KH.SO_DIEN_THOAI - 1] || "");
      var displayText =
        ma === sdt || !sdt
          ? ma + " - " + ten
          : ma + " - " + ten + " (" + sdt + ")";
      result.push({
        value: ma,
        text: displayText,
      });
    }
  });

  return result;
}

/**
 * Thêm khách hàng mới
 *
 * @param {Object} data - {hoTen, soDienThoai, cccd, diaChi, ghiChu}
 * @return {string} Mã KH mới
 */
function addKhachHang(data) {
  initializeColumnEnums();
  var phone = String(data.soDienThoai || "").trim();
  var maKH = phone ? phone : generateId("KH", SHEET_NAMES.KHACH_HANG);

  // Kiểm tra SĐT/Mã KH trùng
  var existing = lookupValue(
    SHEET_NAMES.KHACH_HANG,
    COL_KH.MA_KH,
    maKH,
    COL_KH.MA_KH,
  );
  if (phone && !existing && COL_KH.SO_DIEN_THOAI !== COL_KH.MA_KH) {
    existing = lookupValue(
      SHEET_NAMES.KHACH_HANG,
      COL_KH.SO_DIEN_THOAI,
      phone,
      COL_KH.MA_KH,
    );
  }
  if (existing) {
    throw new Error(
      'Khách hàng với SĐT/Mã KH "' +
        (phone || maKH) +
        '" đã tồn tại (Mã: ' +
        existing +
        ")",
    );
  }

  var rowData = [];
  rowData[COL_KH.MA_KH - 1] = maKH;
  rowData[COL_KH.HO_TEN - 1] = data.hoTen || "";
  if (COL_KH.SO_DIEN_THOAI !== COL_KH.MA_KH) {
    rowData[COL_KH.SO_DIEN_THOAI - 1] = phone;
  }
  rowData[COL_KH.CCCD - 1] = data.cccd || "";
  rowData[COL_KH.DIA_CHI - 1] = data.diaChi || "";
  rowData[COL_KH.NGAY_TAO - 1] = new Date();
  rowData[COL_KH.GHI_CHU - 1] = data.ghiChu || "";

  appendRow(SHEET_NAMES.KHACH_HANG, rowData);
  showToast("Đã thêm khách hàng: " + data.hoTen + " (" + maKH + ")");
  return maKH;
}

/**
 * Cập nhật thông tin khách hàng
 *
 * @param {string} maKH - Mã KH
 * @param {Object} data - Dữ liệu cập nhật
 * @return {boolean}
 */
function updateKhachHang(maKH, data) {
  initializeColumnEnums();
  var row = findRow(SHEET_NAMES.KHACH_HANG, COL_KH.MA_KH, maKH);
  if (row === -1) {
    showAlert("Lỗi", "Không tìm thấy khách hàng: " + maKH);
    return false;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.KHACH_HANG);

  if (data.soDienThoai !== undefined) {
    sheet.getRange(row, COL_KH.MA_KH).setValue(data.soDienThoai);
    if (COL_KH.SO_DIEN_THOAI !== COL_KH.MA_KH) {
      sheet.getRange(row, COL_KH.SO_DIEN_THOAI).setValue(data.soDienThoai);
    }
  }
  if (data.hoTen !== undefined)
    sheet.getRange(row, COL_KH.HO_TEN).setValue(data.hoTen);
  if (data.cccd !== undefined)
    sheet.getRange(row, COL_KH.CCCD).setValue(data.cccd);
  if (data.diaChi !== undefined)
    sheet.getRange(row, COL_KH.DIA_CHI).setValue(data.diaChi);
  if (data.ghiChu !== undefined)
    sheet.getRange(row, COL_KH.GHI_CHU).setValue(data.ghiChu);

  showToast("Đã cập nhật KH: " + maKH);
  return true;
}

/**
 * Tìm kiếm khách hàng theo từ khóa
 *
 * @param {string} keyword - Tên, SĐT, hoặc CCCD
 * @return {Object[]}
 */
function searchKhachHang(keyword) {
  initializeColumnEnums();
  var data = getAllData(SHEET_NAMES.KHACH_HANG);
  var headers = SHEET_HEADERS[SHEET_NAMES.KHACH_HANG];
  var result = [];
  var kw = String(keyword).trim().toLowerCase();

  data.forEach(function (row) {
    var hoTen = String(row[COL_KH.HO_TEN - 1] || "").toLowerCase();
    var sdt = String(row[COL_KH.SO_DIEN_THOAI - 1] || "").toLowerCase();
    var cccd = String(row[COL_KH.CCCD - 1] || "").toLowerCase();

    if (
      hoTen.indexOf(kw) !== -1 ||
      sdt.indexOf(kw) !== -1 ||
      cccd.indexOf(kw) !== -1
    ) {
      var obj = {};
      headers.forEach(function (h, i) {
        obj[h] = row[i];
      });
      result.push(obj);
    }
  });

  return result;
}

/**
 * Build cache dữ liệu KH cho tìm kiếm nhanh
 * @private
 */
function _buildKhachHangDropdownCache() {
  initializeColumnEnums();
  var data = getAllData(SHEET_NAMES.KHACH_HANG);
  var result = [];

  data.forEach(function (row) {
    var ma = String(row[COL_KH.MA_KH - 1] || "");
    if (ma.trim() !== "") {
      var ten = String(row[COL_KH.HO_TEN - 1] || "");
      var sdt = String(row[COL_KH.SO_DIEN_THOAI - 1] || "");
      var cccd = String(row[COL_KH.CCCD - 1] || "");
      var displayText =
        ma === sdt || !sdt
          ? ma + " - " + ten
          : ma + " - " + ten + " (" + sdt + ")";
      result.push({
        v: ma,
        t: displayText,
        _s: (ma + " " + ten + " " + sdt + " " + cccd).toLowerCase(),
      });
    }
  });

  return result;
}

/**
 * Lấy danh sách KH cho dropdown theo từ khóa tìm kiếm (Tên, SĐT, CCCD, Mã KH)
 * Sử dụng CacheService để tránh đọc sheet mỗi lần tìm kiếm
 *
 * @param {string} keyword - Từ khóa tìm
 * @return {Object[]} [{value: 'KH001', text: 'KH001 - Nguyễn Văn A (0901234567)'}, ...]
 */
function getKhachHangDropdownSearch(keyword) {
  var kw = String(keyword).trim().toLowerCase();

  // Đọc từ cache, nếu miss thì build lại
  var allItems = getChunkedCache("dd_kh");
  if (!allItems) {
    allItems = _buildKhachHangDropdownCache();
    setChunkedCache("dd_kh", allItems);
  }

  var result = [];
  for (var i = 0; i < allItems.length; i++) {
    if (allItems[i]._s.indexOf(kw) !== -1) {
      result.push({
        value: allItems[i].v,
        text: allItems[i].t,
      });
      if (result.length >= 100) break;
    }
  }

  return result;
}

/**
 * Đảm bảo khách hàng tồn tại trong danh sách (tự động thêm mới nếu chưa có)
 * 
 * @param {string} maKH Mã khách hàng (thường là SĐT)
 * @param {string} hoTen Tên khách hàng
 * @param {string} [soDienThoai] Số điện thoại (tuỳ chọn, mặc định lấy maKH nếu là số)
 * @return {string} Tên khách hàng (đã được làm sạch)
 */
function ensureKhachHangExists(maKH, hoTen, soDienThoai) {
  if (!maKH) return "";
  
  initializeColumnEnums();
  
  // Kiểm tra xem mã khách hàng đã có trong DB chưa
  var existingRow = findRow(SHEET_NAMES.KHACH_HANG, COL_KH.MA_KH, maKH);
  
  if (existingRow === -1 && hoTen && hoTen !== maKH) {
    // Khách hàng chưa tồn tại, tự động thêm mới
    var phone = soDienThoai || (/\d{9,11}/.test(maKH) ? maKH : "");
    var rowData = [];
    rowData[COL_KH.MA_KH - 1] = maKH;
    rowData[COL_KH.HO_TEN - 1] = hoTen;
    if (COL_KH.SO_DIEN_THOAI !== COL_KH.MA_KH) {
      rowData[COL_KH.SO_DIEN_THOAI - 1] = phone;
    }
    rowData[COL_KH.NGAY_TAO - 1] = new Date();
    
    appendRow(SHEET_NAMES.KHACH_HANG, rowData);
    
    // Clear cache để dropdown cập nhật
    clearSheetCache(SHEET_NAMES.KHACH_HANG);
    setChunkedCache("dd_kh", null); 
  }
  
  return hoTen;
}
