/**
 * ============================================================
 * VanTran Mobile — PhuKien.gs
 * Quản lý danh mục phụ kiện + tồn kho chi tiết
 * ============================================================
 */

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

  data.forEach(function (row) {
    if (chiConHang) {
      if (String(row[8]) === "Ngừng bán" || Number(row[6]) <= 0) return;
    }
    if (chiNhanh) {
      if (String(row[9] || "") !== chiNhanh) return;
    }
    result.push({
      MaPK: String(row[0]),
      TenSP: String(row[1]),
      LoaiPK: String(row[2]),
      ThuongHieu: String(row[3]),
      GiaNhap: Number(row[4]) || 0,
      GiaBan: Number(row[5]) || 0,
      SoLuongTon: Number(row[6]) || 0,
      MoTa: String(row[7]),
      TrangThai: String(row[8]),
      ChiNhanh: String(row[9] || ""),
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

  data.forEach(function (row) {
    if (
      row[0] &&
      String(row[0]).trim() !== "" &&
      String(row[8]) !== "Ngừng bán" &&
      Number(row[6]) > 0
    ) {
      var rowChiNhanh = String(row[9] || "");
      if (!chiNhanh || rowChiNhanh === chiNhanh) {
        result.push({
          value: String(row[0]),
          text:
            row[0] +
            " - " +
            row[1] +
            " (Tồn: " +
            row[6] +
            " | " +
            rowChiNhanh +
            ")",
          giaBan: Number(row[5]),
          soLuongTon: Number(row[6]),
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
  var maPK = generateId("PK", SHEET_NAMES.PHU_KIEN);

  var rowData = [
    maPK,
    data.tenSP || "",
    data.loaiPK || "Khác",
    data.thuongHieu || "",
    Number(data.giaNhap) || 0,
    Number(data.giaBan) || 0,
    Number(data.soLuongTon) || 0,
    data.moTa || "",
    "Đang bán",
    data.chiNhanh || "", // 10th column
  ];

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
  if (data.tenSP !== undefined) sheet.getRange(row, 2).setValue(data.tenSP);
  if (data.loaiPK !== undefined) sheet.getRange(row, 3).setValue(data.loaiPK);
  if (data.thuongHieu !== undefined)
    sheet.getRange(row, 4).setValue(data.thuongHieu);
  if (data.giaNhap !== undefined)
    sheet.getRange(row, 5).setValue(Number(data.giaNhap));
  if (data.giaBan !== undefined)
    sheet.getRange(row, 6).setValue(Number(data.giaBan));
  if (data.soLuongTon !== undefined)
    sheet.getRange(row, 7).setValue(Number(data.soLuongTon));
  if (data.moTa !== undefined) sheet.getRange(row, 8).setValue(data.moTa);
  if (data.trangThai !== undefined)
    sheet.getRange(row, 9).setValue(data.trangThai);

  showToast("Đã cập nhật PK: " + maPK + " (" + chiNhanh + ")");
  return true;
}

/**
 * Tìm dòng phụ kiện khớp cả mã và chi nhánh
 *
 * @param {string} maPK
 * @param {string} chiNhanh
 * @return {number} Row index (1-indexed) hoặc -1
 */
function findPhuKienRow(maPK, chiNhanh) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  if (!sheet) return -1;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;

  var pkData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var cnData = sheet.getRange(2, 10, lastRow - 1, 1).getValues(); // 10th column is ChiNhanh

  var searchPK = String(maPK).trim().toLowerCase();
  var searchCN = String(chiNhanh).trim().toLowerCase();

  for (var i = 0; i < pkData.length; i++) {
    var rowPK = String(pkData[i][0]).trim().toLowerCase();
    var rowCN = String(cnData[i][0]).trim().toLowerCase();

    if (rowPK === searchPK && rowCN === searchCN) {
      return i + 2;
    }
  }

  return -1;
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
        1,
        maPK,
        [2, 3, 4, 5, 6, 8],
      );
      if (details) {
        var rowData = [
          maPK,
          details[2] || "", // tenSP
          details[3] || "Khác", // loaiPK
          details[4] || "", // thuongHieu
          Number(details[5]) || 0, // giaNhap
          Number(details[6]) || 0, // giaBan
          Number(soLuong), // soLuongTon
          details[8] || "", // moTa
          "Đang bán",
          chiNhanh, // 10th column: ChiNhanh
        ];
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
  var currentTon = Number(sheet.getRange(row, 7).getValue()) || 0;
  if (type === "nhap") {
    sheet.getRange(row, 7).setValue(currentTon + Number(soLuong));
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
    sheet.getRange(row, 7).setValue(newTon);
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

  data.forEach(function (row) {
    if (
      row[0] &&
      String(row[0]).trim() !== "" &&
      String(row[8]) === "Đang bán" &&
      Number(row[6]) <= 5 &&
      Number(row[6]) > 0
    ) {
      result.push({
        MaPK: String(row[0]),
        TenSP: String(row[1]),
        LoaiPK: String(row[2]),
        ThuongHieu: String(row[3]),
        GiaNhap: Number(row[4]) || 0,
        GiaBan: Number(row[5]) || 0,
        SoLuongTon: Number(row[6]) || 0,
        MoTa: String(row[7]),
        TrangThai: String(row[8]),
        ChiNhanh: String(row[9] || ""),
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

  data.forEach(function (row) {
    var maPK = String(row[0]).trim();
    if (maPK && maPK !== "" && String(row[8]) !== "Ngừng bán") {
      if (!seen[maPK]) {
        seen[maPK] = true;
        result.push({
          value: maPK,
          text: maPK + " - " + row[1] + " (" + row[2] + " | " + row[3] + ")",
          giaBan: Number(row[5]),
          chiNhanh: String(row[9] || "")
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
  if (!maPK || !chiNhanh) return 0;
  var row = findPhuKienRow(maPK, chiNhanh);
  if (row === -1) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  if (!sheet) return 0;
  return Number(sheet.getRange(row, 7).getValue()) || 0;
}

/**
 * Build cache phụ kiện cho tìm kiếm nhanh (theo chi nhánh, còn hàng)
 * @private
 */
function _buildPhuKienDropdownCache() {
  var data = getAllData(SHEET_NAMES.PHU_KIEN);
  var result = [];

  data.forEach(function (row) {
    if (
      row[0] &&
      String(row[0]).trim() !== "" &&
      String(row[8]) !== "Ngừng bán" &&
      Number(row[6]) > 0
    ) {
      var cn = String(row[9] || "");
      result.push({
        v: String(row[0]),
        t: row[0] + " - " + row[1] + " (Tồn: " + row[6] + " | " + cn + ")",
        gb: Number(row[5]),
        sl: Number(row[6]),
        cn: cn,
        _s: (String(row[0]) + " " + String(row[1])).toLowerCase(),
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

  data.forEach(function (row) {
    var maPK = String(row[0]).trim();
    if (maPK && maPK !== "" && String(row[8]) !== "Ngừng bán") {
      if (!seen[maPK]) {
        seen[maPK] = true;
        result.push({
          v: maPK,
          t: maPK + " - " + row[1] + " (" + row[2] + " | " + row[3] + ")",
          gb: Number(row[5]),
          cn: String(row[9] || ""),
          _s: (maPK + " " + String(row[1])).toLowerCase(),
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

  var allItems = getChunkedCache('dd_pk');
  if (!allItems) {
    allItems = _buildPhuKienDropdownCache();
    setChunkedCache('dd_pk', allItems);
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
      if (result.length >= 100) break;
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

  var allItems = getChunkedCache('dd_pku');
  if (!allItems) {
    allItems = _buildPhuKienUniqueCache();
    setChunkedCache('dd_pku', allItems);
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
      if (result.length >= 100) break;
    }
  }

  return result;
}



