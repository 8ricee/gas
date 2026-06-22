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
  var data = getAllData(SHEET_NAMES.KHACH_HANG);
  var result = [];

  data.forEach(function (row) {
    if (row[0] && String(row[0]).trim() !== "") {
      result.push({
        value: String(row[0]),
        text: row[0] + " - " + row[1] + " (" + row[2] + ")",
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
  // Kiểm tra SĐT trùng
  if (data.soDienThoai) {
    var existing = lookupValue(SHEET_NAMES.KHACH_HANG, 3, data.soDienThoai, 1);
    if (existing) {
      throw new Error(
        'SĐT "' + data.soDienThoai + '" đã tồn tại (Mã: ' + existing + ")",
      );
    }
  }

  var maKH = generateId("KH", SHEET_NAMES.KHACH_HANG);

  var rowData = [
    maKH,
    data.hoTen || "",
    data.soDienThoai || "",
    data.cccd || "",
    data.diaChi || "",
    new Date(),
    data.ghiChu || "",
  ];

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
  var row = findRow(SHEET_NAMES.KHACH_HANG, 1, maKH);
  if (row === -1) {
    showAlert("Lỗi", "Không tìm thấy khách hàng: " + maKH);
    return false;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.KHACH_HANG);

  if (data.hoTen !== undefined) sheet.getRange(row, 2).setValue(data.hoTen);
  if (data.soDienThoai !== undefined)
    sheet.getRange(row, 3).setValue(data.soDienThoai);
  if (data.cccd !== undefined) sheet.getRange(row, 4).setValue(data.cccd);
  if (data.diaChi !== undefined) sheet.getRange(row, 5).setValue(data.diaChi);
  if (data.ghiChu !== undefined) sheet.getRange(row, 7).setValue(data.ghiChu);

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
  var data = getAllData(SHEET_NAMES.KHACH_HANG);
  var headers = SHEET_HEADERS[SHEET_NAMES.KHACH_HANG];
  var result = [];
  var kw = String(keyword).trim().toLowerCase();

  data.forEach(function (row) {
    var hoTen = String(row[1]).toLowerCase();
    var sdt = String(row[2]).toLowerCase();
    var cccd = String(row[3]).toLowerCase();

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
  var data = getAllData(SHEET_NAMES.KHACH_HANG);
  var result = [];

  data.forEach(function (row) {
    if (row[0] && String(row[0]).trim() !== "") {
      result.push({
        v: String(row[0]),
        t: row[0] + " - " + row[1] + " (" + row[2] + ")",
        _s: (String(row[0]) + " " + String(row[1]) + " " + String(row[2]) + " " + String(row[3])).toLowerCase(),
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
  var allItems = getChunkedCache('dd_kh');
  if (!allItems) {
    allItems = _buildKhachHangDropdownCache();
    setChunkedCache('dd_kh', allItems);
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


