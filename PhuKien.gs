/**
 * ============================================================
 * VanTran Mobile — PhuKien.gs
 * Quản lý danh mục phụ kiện + tồn kho chi tiết
 * ============================================================
 */

/**
 * Lấy index cột (0-indexed) cho Phụ kiện
 * @private
 */
function _getPhuKienIndices() {
  initializeColumnEnums();
  return {
    maPK: COL_PK.MA_PK - 1,
    tenSP: COL_PK.TEN_SP - 1,
    loaiPK: COL_PK.LOAI_PK - 1,
    thuongHieu: COL_PK.THUONG_HIEU - 1,
    giaNhap: COL_PK.GIA_NHAP - 1,
    giaBan: COL_PK.GIA_BAN - 1,
    soLuongTon: COL_PK.SO_LUONG_TON - 1,
    moTa: COL_PK.MO_TA - 1,
    trangThai: COL_PK.TRANG_THAI - 1,
    chiNhanh: COL_PK.CHI_NHANH - 1,
  };
}

/**
 * Lấy danh sách tất cả phụ kiện
 *
 * @param {boolean} [chiConHang] - Nếu true, chỉ lấy SP còn hàng (tồn > 0 & đang bán)
 * @param {string} [chiNhanh] - Chỉ lấy SP ở chi nhánh
 * @return {Object[]}
 */
function getAllPhuKien(chiConHang, chiNhanh) {
  var data = getAllData(SHEET_NAMES.PHU_KIEN);
  var result = [];
  var c = _getPhuKienIndices();

  data.forEach(function (row) {
    if (row.length <= c.chiNhanh) return;
    if (chiConHang) {
      if (
        String(row[c.trangThai]) === "Ngừng bán" ||
        Number(row[c.soLuongTon]) <= 0
      )
        return;
    }
    if (chiNhanh) {
      if (String(row[c.chiNhanh] || "") !== chiNhanh) return;
    }
    result.push({
      MaPK: String(row[c.maPK]),
      TenSP: String(row[c.tenSP]),
      LoaiPK: String(row[c.loaiPK]),
      ThuongHieu: String(row[c.thuongHieu]),
      GiaNhap: Number(row[c.giaNhap]) || 0,
      GiaBan: Number(row[c.giaBan]) || 0,
      SoLuongTon: Number(row[c.soLuongTon]) || 0,
      MoTa: String(row[c.moTa]),
      TrangThai: String(row[c.trangThai]),
      ChiNhanh: String(row[c.chiNhanh] || ""),
    });
  });

  return result;
}

/**
 * Lấy danh sách phụ kiện cho dropdown theo chi nhánh
 *
 * @param {string} [chiNhanh] - Chi nhánh lọc
 * @return {Object[]} [{value: 'PK001', text: 'PK001 - Ốp lưng iPhone 15 (Tồn: 10)'}, ...]
 */
function getPhuKienDropdown(chiNhanh) {
  var data = getAllData(SHEET_NAMES.PHU_KIEN);
  var result = [];
  var c = _getPhuKienIndices();

  data.forEach(function (row) {
    if (row.length <= c.chiNhanh) return;
    if (
      row[c.maPK] &&
      String(row[c.maPK]).trim() !== "" &&
      String(row[c.trangThai]) !== "Ngừng bán" &&
      Number(row[c.soLuongTon]) > 0
    ) {
      var rowChiNhanh = String(row[c.chiNhanh] || "");
      if (!chiNhanh || rowChiNhanh === chiNhanh) {
        result.push({
          value: String(row[c.maPK]),
          text:
            row[c.maPK] +
            " - " +
            row[c.tenSP] +
            " (Tồn: " +
            row[c.soLuongTon] +
            " | " +
            rowChiNhanh +
            ")",
          giaBan: Number(row[c.giaBan]),
          soLuongTon: Number(row[c.soLuongTon]),
          chiNhanh: rowChiNhanh,
        });
      }
    }
  });

  return result;
}

/**
 * Thêm phụ kiện mới
 *
 * @param {Object} data - {tenSP, loaiPK, thuongHieu, giaNhap, giaBan, soLuongTon, moTa, chiNhanh}
 * @return {string} Mã PK mới
 */
function addPhuKien(data) {
  initializeColumnEnums();
  var maPK = generateId("PK", SHEET_NAMES.PHU_KIEN);

  var rowData = [];
  rowData[COL_PK.MA_PK - 1] = maPK;
  rowData[COL_PK.TEN_SP - 1] = data.tenSP || "";
  rowData[COL_PK.LOAI_PK - 1] = data.loaiPK || "Khác";
  rowData[COL_PK.THUONG_HIEU - 1] = data.thuongHieu || "";
  rowData[COL_PK.GIA_NHAP - 1] = Number(data.giaNhap) || 0;
  rowData[COL_PK.GIA_BAN - 1] = Number(data.giaBan) || 0;
  rowData[COL_PK.SO_LUONG_TON - 1] = Number(data.soLuongTon) || 0;
  rowData[COL_PK.MO_TA - 1] = data.moTa || "";
  rowData[COL_PK.TRANG_THAI - 1] = "Đang bán";
  rowData[COL_PK.CHI_NHANH - 1] = data.chiNhanh || "";

  appendRow(SHEET_NAMES.PHU_KIEN, rowData);
  showToast("Đã thêm phụ kiện: " + data.tenSP + " (" + maPK + ")");
  return maPK;
}

/**
 * Cập nhật thông tin phụ kiện
 *
 * @param {string} maPK - Mã phụ kiện
 * @param {Object} data - Dữ liệu cập nhật (phải có chiNhanh nếu muốn cập nhật đúng dòng!)
 * @return {boolean}
 */
function updatePhuKien(maPK, data) {
  initializeColumnEnums();
  var chiNhanh = data.chiNhanh;
  if (!chiNhanh) {
    showAlert("❌ Lỗi", "Cần cung cấp chi nhánh để cập nhật phụ kiện: " + maPK);
    return false;
  }

  var row = findPhuKienRow(maPK, chiNhanh);
  if (row === -1) {
    showAlert(
      "❌ Lỗi",
      "Không tìm thấy phụ kiện: " + maPK + " tại " + chiNhanh,
    );
    return false;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  invalidateDropdownCache(SHEET_NAMES.PHU_KIEN);

  // Đảm bảo đủ cột
  var maxColNeeded = Math.max(
    COL_PK.MA_PK,
    COL_PK.TEN_SP,
    COL_PK.LOAI_PK,
    COL_PK.THUONG_HIEU,
    COL_PK.GIA_NHAP,
    COL_PK.GIA_BAN,
    COL_PK.SO_LUONG_TON,
    COL_PK.MO_TA,
    COL_PK.TRANG_THAI,
    COL_PK.CHI_NHANH,
  );
  var maxCols = sheet.getMaxColumns();
  if (maxCols < maxColNeeded) {
    sheet.insertColumnsAfter(maxCols, maxColNeeded - maxCols);
  }

  var range = sheet.getRange(row, 1, 1, maxColNeeded);
  var rowValues = range.getValues()[0];

  if (data.tenSP !== undefined)
    rowValues[COL_PK.TEN_SP - 1] = data.tenSP;
  if (data.loaiPK !== undefined)
    rowValues[COL_PK.LOAI_PK - 1] = data.loaiPK;
  if (data.thuongHieu !== undefined)
    rowValues[COL_PK.THUONG_HIEU - 1] = data.thuongHieu;
  if (data.giaNhap !== undefined)
    rowValues[COL_PK.GIA_NHAP - 1] = Number(data.giaNhap);
  if (data.giaBan !== undefined)
    rowValues[COL_PK.GIA_BAN - 1] = Number(data.giaBan);
  if (data.soLuongTon !== undefined)
    rowValues[COL_PK.SO_LUONG_TON - 1] = Number(data.soLuongTon);
  if (data.moTa !== undefined)
    rowValues[COL_PK.MO_TA - 1] = data.moTa;
  if (data.trangThai !== undefined)
    rowValues[COL_PK.TRANG_THAI - 1] = data.trangThai;

  range.setValues([rowValues]);

  showToast("Đã cập nhật PK: " + maPK + " (" + chiNhanh + ")");
  return true;
}

var _phuKienCompositeIndex = null;

function clearPhuKienCompositeIndex() {
  _phuKienCompositeIndex = null;
}

/**
 * Tìm dòng phụ kiện khớp cả mã và chi nhánh
 *
 * @param {string} maPK
 * @param {string} chiNhanh
 * @return {number} Row index (1-indexed) hoặc -1
 */
function findPhuKienRow(maPK, chiNhanh) {
  initializeColumnEnums();
  if (!_phuKienCompositeIndex) {
    _phuKienCompositeIndex = {};
    var data = getAllData(SHEET_NAMES.PHU_KIEN);
    var maPKIdx = COL_PK.MA_PK - 1;
    var chiNhanhIdx = COL_PK.CHI_NHANH - 1;

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row.length <= chiNhanhIdx) continue;
      var key =
        String(row[maPKIdx]).trim().toLowerCase() +
        "_" +
        String(row[chiNhanhIdx] || "")
          .trim()
          .toLowerCase();
      if (!(key in _phuKienCompositeIndex)) {
        _phuKienCompositeIndex[key] = i + 2; // +2 vì bỏ qua header
      }
    }
  }

  var searchKey =
    String(maPK).trim().toLowerCase() +
    "_" +
    String(chiNhanh).trim().toLowerCase();
  return _phuKienCompositeIndex[searchKey] || -1;
}

/**
 * Cập nhật tồn kho phụ kiện theo chi nhánh
 *
 * @param {string} maPK - Mã phụ kiện
 * @param {number} soLuong - Số lượng thay đổi
 * @param {string} type - 'nhap' (cộng) hoặc 'xuat' (trừ)
 * @param {string} chiNhanh - Chi nhánh lưu kho
 * @return {boolean}
 */
function updateTonKhoPhuKien(maPK, soLuong, type, chiNhanh) {
  initializeColumnEnums();
  if (!chiNhanh) {
    throw new Error("Thiếu chi nhánh khi cập nhật tồn kho phụ kiện!");
  }

  var row = findPhuKienRow(maPK, chiNhanh);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);

  if (row === -1) {
    if (type === "nhap") {
      // Nếu chưa có ở chi nhánh này, sao chép thông tin từ chi nhánh bất kỳ
      var details = lookupMultipleValues(
        SHEET_NAMES.PHU_KIEN,
        COL_PK.MA_PK,
        maPK,
        [
          COL_PK.TEN_SP,
          COL_PK.LOAI_PK,
          COL_PK.THUONG_HIEU,
          COL_PK.GIA_NHAP,
          COL_PK.GIA_BAN,
          COL_PK.MO_TA,
        ],
      );
      if (details) {
        var rowData = [];
        rowData[COL_PK.MA_PK - 1] = maPK;
        rowData[COL_PK.TEN_SP - 1] = details[COL_PK.TEN_SP] || "";
        rowData[COL_PK.LOAI_PK - 1] = details[COL_PK.LOAI_PK] || "Khác";
        rowData[COL_PK.THUONG_HIEU - 1] = details[COL_PK.THUONG_HIEU] || "";
        rowData[COL_PK.GIA_NHAP - 1] = Number(details[COL_PK.GIA_NHAP]) || 0;
        rowData[COL_PK.GIA_BAN - 1] = Number(details[COL_PK.GIA_BAN]) || 0;
        rowData[COL_PK.SO_LUONG_TON - 1] = Number(soLuong);
        rowData[COL_PK.MO_TA - 1] = details[COL_PK.MO_TA] || "";
        rowData[COL_PK.TRANG_THAI - 1] = "Đang bán";
        rowData[COL_PK.CHI_NHANH - 1] = chiNhanh;

        appendRow(SHEET_NAMES.PHU_KIEN, rowData);
        return true;
      } else {
        throw new Error(
          "Không tìm thấy thông tin sản phẩm phụ kiện: " +
            maPK +
            " ở các chi nhánh khác!",
        );
      }
    } else {
      throw new Error(
        "Sản phẩm " + maPK + " không tồn tại tại " + chiNhanh + " để xuất kho!",
      );
    }
  }

  // Cập nhật dòng có sẵn
  invalidateDropdownCache(SHEET_NAMES.PHU_KIEN);
  var currentTon =
    Number(sheet.getRange(row, COL_PK.SO_LUONG_TON).getValue()) || 0;
  if (type === "nhap") {
    sheet
      .getRange(row, COL_PK.SO_LUONG_TON)
      .setValue(currentTon + Number(soLuong));
  } else if (type === "xuat") {
    var newTon = currentTon - Number(soLuong);
    if (newTon < 0) {
      throw new Error(
        "Không đủ tồn kho tại " +
          chiNhanh +
          "! Hiện tại: " +
          currentTon +
          ", cần xuất: " +
          soLuong,
      );
    }
    sheet.getRange(row, COL_PK.SO_LUONG_TON).setValue(newTon);
  }

  return true;
}

/**
 * Lấy danh sách phụ kiện sắp hết hàng (tồn <= 5)
 *
 * @return {Object[]}
 */
function getPhuKienSapHet() {
  var data = getAllData(SHEET_NAMES.PHU_KIEN);
  var result = [];
  var c = _getPhuKienIndices();

  data.forEach(function (row) {
    if (row.length <= c.chiNhanh) return;
    if (
      row[c.maPK] &&
      String(row[c.maPK]).trim() !== "" &&
      String(row[c.trangThai]) === "Đang bán" &&
      Number(row[c.soLuongTon]) <= 5 &&
      Number(row[c.soLuongTon]) > 0
    ) {
      result.push({
        MaPK: String(row[c.maPK]),
        TenSP: String(row[c.tenSP]),
        LoaiPK: String(row[c.loaiPK]),
        ThuongHieu: String(row[c.thuongHieu]),
        GiaNhap: Number(row[c.giaNhap]) || 0,
        GiaBan: Number(row[c.giaBan]) || 0,
        SoLuongTon: Number(row[c.soLuongTon]) || 0,
        MoTa: String(row[c.moTa]),
        TrangThai: String(row[c.trangThai]),
        ChiNhanh: String(row[c.chiNhanh] || ""),
      });
    }
  });

  return result;
}

/**
 * Lấy danh sách phụ kiện duy nhất (không trùng mã) để phục vụ nhập kho cộng tồn
 * Bao gồm cả các sản phẩm đã hết hàng (tồn = 0)
 *
 * @return {Object[]}
 */
function getPhuKienUniqueList() {
  var data = getAllData(SHEET_NAMES.PHU_KIEN);
  var result = [];
  var seen = {};
  var c = _getPhuKienIndices();

  data.forEach(function (row) {
    if (row.length <= c.chiNhanh) return;
    var maPK = String(row[c.maPK]).trim();
    if (maPK && maPK !== "" && String(row[c.trangThai]) !== "Ngừng bán") {
      if (!seen[maPK]) {
        seen[maPK] = true;
        result.push({
          value: maPK,
          text:
            maPK +
            " - " +
            row[c.tenSP] +
            " (" +
            row[c.loaiPK] +
            " | " +
            row[c.thuongHieu] +
            ")",
          giaBan: Number(row[c.giaBan]),
          chiNhanh: String(row[c.chiNhanh] || ""),
        });
      }
    }
  });

  return result;
}

/**
 * Lấy số lượng tồn kho của phụ kiện tại chi nhánh cụ thể
 *
 * @param {string} maPK
 * @param {string} chiNhanh
 * @return {number} Số lượng tồn
 */
function getPhuKienStockAtBranch(maPK, chiNhanh) {
  initializeColumnEnums();
  if (!maPK || !chiNhanh) return 0;
  var row = findPhuKienRow(maPK, chiNhanh);
  if (row === -1) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  if (!sheet) return 0;
  return Number(sheet.getRange(row, COL_PK.SO_LUONG_TON).getValue()) || 0;
}

/**
 * Build cache phụ kiện cho tìm kiếm nhanh (theo chi nhánh, còn hàng)
 * @private
 */
function _buildPhuKienDropdownCache() {
  var data = getAllData(SHEET_NAMES.PHU_KIEN);
  var result = [];
  var c = _getPhuKienIndices();

  data.forEach(function (row) {
    if (row.length <= c.chiNhanh) return;
    if (
      row[c.maPK] &&
      String(row[c.maPK]).trim() !== "" &&
      String(row[c.trangThai]) !== "Ngừng bán" &&
      Number(row[c.soLuongTon]) > 0
    ) {
      var cn = String(row[c.chiNhanh] || "");
      result.push({
        v: String(row[c.maPK]),
        t:
          row[c.maPK] +
          " - " +
          row[c.tenSP] +
          " (Tồn: " +
          row[c.soLuongTon] +
          " | " +
          cn +
          ")",
        gb: Number(row[c.giaBan]),
        sl: Number(row[c.soLuongTon]),
        cn: cn,
        _s: (String(row[c.maPK]) + " " + String(row[c.tenSP])).toLowerCase(),
      });
    }
  });

  return result;
}

/**
 * Build cache phụ kiện duy nhất (không trùng mã) cho tìm kiếm nhanh
 * @private
 */
function _buildPhuKienUniqueCache() {
  var data = getAllData(SHEET_NAMES.PHU_KIEN);
  var result = [];
  var seen = {};
  var c = _getPhuKienIndices();

  data.forEach(function (row) {
    if (row.length <= c.chiNhanh) return;
    var maPK = String(row[c.maPK]).trim();
    if (maPK && maPK !== "" && String(row[c.trangThai]) !== "Ngừng bán") {
      if (!seen[maPK]) {
        seen[maPK] = true;
        result.push({
          v: maPK,
          t:
            maPK +
            " - " +
            row[c.tenSP] +
            " (" +
            row[c.loaiPK] +
            " | " +
            row[c.thuongHieu] +
            ")",
          gb: Number(row[c.giaBan]),
          cn: String(row[c.chiNhanh] || ""),
          _s: (maPK + " " + String(row[c.tenSP])).toLowerCase(),
        });
      }
    }
  });

  return result;
}

/**
 * Lấy danh sách phụ kiện cho dropdown theo chi nhánh và từ khóa tìm kiếm
 * Sử dụng CacheService để tránh đọc sheet mỗi lần tìm kiếm
 *
 * @param {string} chiNhanh - Chi nhánh lọc
 * @param {string} keyword - Từ khóa tìm
 * @return {Object[]}
 */
function getPhuKienDropdownSearch(chiNhanh, keyword) {
  var kw = String(keyword).trim().toLowerCase();

  var allItems = getChunkedCache("dd_pk");
  if (!allItems) {
    allItems = _buildPhuKienDropdownCache();
    setChunkedCache("dd_pk", allItems);
  }

  var result = [];
  for (var i = 0; i < allItems.length; i++) {
    var item = allItems[i];
    if (chiNhanh && item.cn !== chiNhanh) continue;
    if (item._s.indexOf(kw) !== -1) {
      result.push({
        value: item.v,
        text: item.t,
        giaBan: item.gb,
        soLuongTon: item.sl,
        chiNhanh: item.cn,
      });
      // Nếu tìm kiếm cụ thể (có keyword) thì giới hạn 100; nếu là preload (không keyword) thì giới hạn 5000
      if (kw && result.length >= 100) break;
      if (!kw && result.length >= 5000) break;
    }
  }

  return result;
}

/**
 * Lấy danh sách phụ kiện duy nhất theo từ khóa (Mã PK, Tên PK)
 * Sử dụng CacheService để tránh đọc sheet mỗi lần tìm kiếm
 *
 * @param {string} keyword - Từ khóa tìm
 * @return {Object[]}
 */
function getPhuKienUniqueListSearch(keyword) {
  var kw = String(keyword).trim().toLowerCase();

  var allItems = getChunkedCache("dd_pku");
  if (!allItems) {
    allItems = _buildPhuKienUniqueCache();
    setChunkedCache("dd_pku", allItems);
  }

  var result = [];
  for (var i = 0; i < allItems.length; i++) {
    if (allItems[i]._s.indexOf(kw) !== -1) {
      result.push({
        value: allItems[i].v,
        text: allItems[i].t,
        giaBan: allItems[i].gb,
        chiNhanh: allItems[i].cn,
      });
      // Nếu tìm kiếm cụ thể (có keyword) thì giới hạn 100; nếu là preload (không keyword) thì giới hạn 5000
      if (kw && result.length >= 100) break;
      if (!kw && result.length >= 5000) break;
    }
  }

  return result;
}
