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
  const data = getAllData(SHEET_NAMES.PHU_KIEN);
  const result = [];

  data.forEach(function (row) {
    const obj = mapRowToObject(row, SHEET_NAMES.PHU_KIEN);
    if (!obj.MA_PK) return;
    if (chiConHang) {
      if (
        String(obj.TRANG_THAI) === PK_STATUS.INACTIVE ||
        Number(obj.SO_LUONG_TON) <= 0
      )
        return;
    }
    if (chiNhanh) {
      if (String(obj.CHI_NHANH || "") !== chiNhanh) return;
    }
    result.push({
      MaPK: String(obj.MA_PK),
      TenSP: String(obj.TEN_SP),
      LoaiPK: String(obj.LOAI_PK),
      ThuongHieu: String(obj.THUONG_HIEU),
      GiaNhap: Number(obj.GIA_NHAP) || 0,
      GiaBan: Number(obj.GIA_BAN) || 0,
      SoLuongTon: Number(obj.SO_LUONG_TON) || 0,
      MoTa: String(obj.MO_TA),
      TrangThai: String(obj.TRANG_THAI),
      ChiNhanh: String(obj.CHI_NHANH || ""),
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
  const data = getAllData(SHEET_NAMES.PHU_KIEN);
  const result = [];

  data.forEach(function (row) {
    const obj = mapRowToObject(row, SHEET_NAMES.PHU_KIEN);
    if (
      obj.MA_PK &&
      String(obj.MA_PK).trim() !== "" &&
      String(obj.TRANG_THAI) !== PK_STATUS.INACTIVE &&
      Number(obj.SO_LUONG_TON) > 0
    ) {
      const rowChiNhanh = String(obj.CHI_NHANH || "");
      if (!chiNhanh || rowChiNhanh === chiNhanh) {
        result.push({
          value: String(obj.MA_PK),
          text:
            obj.MA_PK +
            " - " +
            obj.TEN_SP +
            " (Tồn: " +
            obj.SO_LUONG_TON +
            " | " +
            rowChiNhanh +
            ")",
          giaBan: Number(obj.GIA_BAN),
          soLuongTon: Number(obj.SO_LUONG_TON),
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
  const maPK = generateId("PK", SHEET_NAMES.PHU_KIEN);

  const rowData = [];
  rowData[COL_PK.MA_PK - 1] = maPK;
  rowData[COL_PK.TEN_SP - 1] = data.tenSP || "";
  rowData[COL_PK.LOAI_PK - 1] = data.loaiPK || "Khác";
  rowData[COL_PK.THUONG_HIEU - 1] = data.thuongHieu || "";
  rowData[COL_PK.GIA_NHAP - 1] = Number(data.giaNhap) || 0;
  rowData[COL_PK.GIA_BAN - 1] = Number(data.giaBan) || 0;
  rowData[COL_PK.SO_LUONG_TON - 1] = Number(data.soLuongTon) || 0;
  rowData[COL_PK.MO_TA - 1] = data.moTa || "";
  rowData[COL_PK.TRANG_THAI - 1] = PK_STATUS.ACTIVE;
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
  const chiNhanh = data.chiNhanh;
  if (!chiNhanh) {
    showAlert("❌ Lỗi", "Cần cung cấp chi nhánh để cập nhật phụ kiện: " + maPK);
    return false;
  }

  const row = findPhuKienRow(maPK, chiNhanh);
  if (row === -1) {
    showAlert(
      "❌ Lỗi",
      "Không tìm thấy phụ kiện: " + maPK + " tại " + chiNhanh,
    );
    return false;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  invalidateDropdownCache(SHEET_NAMES.PHU_KIEN);

  // Đảm bảo đủ cột
  const maxColNeeded = Math.max(
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
  const maxCols = sheet.getMaxColumns();
  if (maxCols < maxColNeeded) {
    sheet.insertColumnsAfter(maxCols, maxColNeeded - maxCols);
  }

  const range = sheet.getRange(row, 1, 1, maxColNeeded);
  const rowValues = range.getValues()[0];

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

let _phuKienCompositeIndex = null;

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
  if (!_phuKienCompositeIndex) {
    _phuKienCompositeIndex = {};
    const data = getAllData(SHEET_NAMES.PHU_KIEN);
    const maPKIdx = COL_PK.MA_PK - 1;
    const chiNhanhIdx = COL_PK.CHI_NHANH - 1;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row.length <= chiNhanhIdx) continue;
      const key =
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

  const searchKey =
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
  return withDocumentLock(function () {
    if (!chiNhanh) {
      throw new Error("Thiếu chi nhánh khi cập nhật tồn kho phụ kiện!");
    }

    const row = findPhuKienRow(maPK, chiNhanh);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);

    if (row === -1) {
      if (type === "nhap") {
        // Nếu chưa có ở chi nhánh này, sao chép thông tin từ chi nhánh bất kỳ
        const details = lookupMultipleValues(
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
          const rowData = [];
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
    const currentTon =
      Number(sheet.getRange(row, COL_PK.SO_LUONG_TON).getValue()) || 0;
    if (type === "nhap") {
      sheet
        .getRange(row, COL_PK.SO_LUONG_TON)
        .setValue(currentTon + Number(soLuong));
    } else if (type === "xuat") {
      const newTon = currentTon - Number(soLuong);
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
  });
}

/**
 * Lấy danh sách phụ kiện sắp hết hàng (tồn <= 5)
 *
 * @return {Object[]}
 */
function getPhuKienSapHet() {
  const data = getAllData(SHEET_NAMES.PHU_KIEN);
  const result = [];

  data.forEach(function (row) {
    const obj = mapRowToObject(row, SHEET_NAMES.PHU_KIEN);
    if (
      obj.MA_PK &&
      String(obj.MA_PK).trim() !== "" &&
      String(obj.TRANG_THAI) === PK_STATUS.ACTIVE &&
      Number(obj.SO_LUONG_TON) <= 5 &&
      Number(obj.SO_LUONG_TON) > 0
    ) {
      result.push({
        MaPK: String(obj.MA_PK),
        TenSP: String(obj.TEN_SP),
        LoaiPK: String(obj.LOAI_PK),
        ThuongHieu: String(obj.THUONG_HIEU),
        GiaNhap: Number(obj.GIA_NHAP) || 0,
        GiaBan: Number(obj.GIA_BAN) || 0,
        SoLuongTon: Number(obj.SO_LUONG_TON) || 0,
        MoTa: String(obj.MO_TA),
        TrangThai: String(obj.TRANG_THAI),
        ChiNhanh: String(obj.CHI_NHANH || ""),
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
  const data = getAllData(SHEET_NAMES.PHU_KIEN);
  const result = [];
  const seen = {};

  data.forEach(function (row) {
    const obj = mapRowToObject(row, SHEET_NAMES.PHU_KIEN);
    const maPK = String(obj.MA_PK || "").trim();
    if (maPK && maPK !== "" && String(obj.TRANG_THAI) !== PK_STATUS.INACTIVE) {
      if (!seen[maPK]) {
        seen[maPK] = true;
        result.push({
          value: maPK,
          text:
            maPK +
            " - " +
            obj.TEN_SP +
            " (" +
            obj.LOAI_PK +
            " | " +
            obj.THUONG_HIEU +
            ")",
          giaBan: Number(obj.GIA_BAN),
          chiNhanh: String(obj.CHI_NHANH || ""),
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
  const row = findPhuKienRow(maPK, chiNhanh);
  if (row === -1) return 0;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  if (!sheet) return 0;
  return Number(sheet.getRange(row, COL_PK.SO_LUONG_TON).getValue()) || 0;
}

/**
 * Build cache phụ kiện cho tìm kiếm nhanh (theo chi nhánh, còn hàng)
 * @private
 */
function _buildPhuKienDropdownCache() {
  const data = getAllData(SHEET_NAMES.PHU_KIEN);
  const result = [];

  data.forEach(function (row) {
    const obj = mapRowToObject(row, SHEET_NAMES.PHU_KIEN);
    if (
      obj.MA_PK &&
      String(obj.MA_PK).trim() !== "" &&
      String(obj.TRANG_THAI) !== PK_STATUS.INACTIVE &&
      Number(obj.SO_LUONG_TON) > 0
    ) {
      const cn = String(obj.CHI_NHANH || "");
      result.push({
        v: String(obj.MA_PK),
        t:
          obj.MA_PK +
          " - " +
          obj.TEN_SP +
          " (Tồn: " +
          obj.SO_LUONG_TON +
          " | " +
          cn +
          ")",
        gb: Number(obj.GIA_BAN),
        sl: Number(obj.SO_LUONG_TON),
        cn: cn,
        _s: (String(obj.MA_PK) + " " + String(obj.TEN_SP)).toLowerCase(),
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
  const data = getAllData(SHEET_NAMES.PHU_KIEN);
  const result = [];
  const seen = {};

  data.forEach(function (row) {
    const obj = mapRowToObject(row, SHEET_NAMES.PHU_KIEN);
    const maPK = String(obj.MA_PK || "").trim();
    if (maPK && maPK !== "" && String(obj.TRANG_THAI) !== PK_STATUS.INACTIVE) {
      if (!seen[maPK]) {
        seen[maPK] = true;
        result.push({
          v: maPK,
          t:
            maPK +
            " - " +
            obj.TEN_SP +
            " (" +
            obj.LOAI_PK +
            " | " +
            obj.THUONG_HIEU +
            ")",
          gb: Number(obj.GIA_BAN),
          cn: String(obj.CHI_NHANH || ""),
          _s: (maPK + " " + String(obj.TEN_SP)).toLowerCase(),
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
  const kw = String(keyword).trim().toLowerCase();

  let allItems = getChunkedCache("dd_pk");
  if (!allItems) {
    allItems = _buildPhuKienDropdownCache();
    setChunkedCache("dd_pk", allItems);
  }

  const result = [];
  for (let i = 0; i < allItems.length; i++) {
    const item = allItems[i];
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
  const kw = String(keyword).trim().toLowerCase();

  let allItems = getChunkedCache("dd_pku");
  if (!allItems) {
    allItems = _buildPhuKienUniqueCache();
    setChunkedCache("dd_pku", allItems);
  }

  const result = [];
  for (let i = 0; i < allItems.length; i++) {
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
