/**
 * ============================================================
 * VanTran Mobile — Formatting.gs
 * Cấu hình định dạng bảng tính & đổ màu có điều kiện
 * ============================================================
 */

/**
 * Định dạng Times New Roman cỡ 12 cho toàn hệ thống
 */
function formatAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Thiết lập locale tiếng Việt cho spreadsheet để các định dạng ngày giờ chuẩn hóa theo vi_VN
  try {
    ss.setSpreadsheetLocale("vi_VN");
  } catch (e) {
    Logger.log("Không thể thiết lập locale vi_VN: " + e.message);
  }

  var sheets = ss.getSheets();

  // Danh sách các sheet dữ liệu chuẩn cần tự động format định dạng tiền tệ và ngày giờ
  var standardDataSheets = [
    SHEET_NAMES.NHAN_VIEN,
    SHEET_NAMES.DIEN_THOAI,
    SHEET_NAMES.PHU_KIEN,
    SHEET_NAMES.KHACH_HANG,
    SHEET_NAMES.DON_HANG,
    SHEET_NAMES.DICH_VU,
    SHEET_NAMES.TRA_GOP,
    SHEET_NAMES.LICH_SU_TRA_GOP,
    SHEET_NAMES.DOANH_SO,
    SHEET_NAMES.NHAP_KHO,
    SHEET_NAMES.DOI_TRA,
    SHEET_NAMES.THU_MUA,
    SHEET_NAMES.BAO_HANH
  ];

  sheets.forEach(function (sheet) {
    var sheetName = sheet.getName();
    var maxRows = sheet.getMaxRows();
    var maxCols = sheet.getMaxColumns();
    if (maxRows > 0 && maxCols > 0) {
      var range = sheet.getRange(1, 1, maxRows, maxCols);
      // Apply basic font settings
      range.setFontFamily("Times New Roman");
      range.setFontSize(12);
    }

    // Reformat headers to remain bold and lefted (skip Tồn kho to preserve custom layout and column widths)
    var lastCol = sheet.getLastColumn();
    var lastRow = sheet.getLastRow();

    if (lastCol > 0 && sheetName !== SHEET_NAMES.TON_KHO) {
      var headerRange = sheet.getRange(1, 1, 1, lastCol);
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("left");

      // Auto resize columns
      for (var col = 1; col <= lastCol; col++) {
        sheet.autoResizeColumn(col);
      }

      // Tự động định dạng tiền tệ #,##0 và ngày giờ cho các cột tương ứng nếu thuộc danh sách sheet dữ liệu chuẩn
      if (standardDataSheets.indexOf(sheetName) !== -1 && maxRows > 1) {
        var headers = headerRange.getValues()[0];
        for (var col = 1; col <= lastCol; col++) {
          var headerName = headers[col - 1];
          var formatRange = sheet.getRange(2, col, maxRows - 1, 1);
          if (isFinancialHeader(headerName)) {
            formatRange.setNumberFormat("#,##0");
          } else if (isDateTimeHeader(headerName)) {
            formatRange.setNumberFormat("dd/MM/yyyy HH:mm:ss");
          } else if (isDateHeader(headerName)) {
            formatRange.setNumberFormat("dd/MM/yyyy");
          }
        }
      }
    }
  });
}

function isFinancialHeader(name) {
  if (!name) return false;
  var n = String(name).trim().toLowerCase();
  
  // Các từ khóa chỉ tiền tệ, giá cả, chi phí, thu nhập
  var keywords = [
    "giá",
    "tiền",
    "phí",
    "thành tiền",
    "đơn giá",
    "hoa hồng",
    "thu nhập",
    "thưởng",
    "mặt", // Tiền mặt
    "khoản", // Chuyển khoản
    "trước", // Trả trước
    "lại", // Còn lại, Tiền hoàn trả, Đã trả lại
    "doanh thu"
  ];
  
  // Loại trừ các trường hợp không phải tiền tệ
  if (n.indexOf("ngày") !== -1) return false;
  if (n.indexOf("nguồn") !== -1) return false;
  if (n.indexOf("tài khoản") !== -1) return false;
  if (n.indexOf("cccd") !== -1) return false;
  if (n.indexOf("số điện thoại") !== -1) return false;
  if (n.indexOf("sđt") !== -1) return false;

  for (var i = 0; i < keywords.length; i++) {
    if (n.indexOf(keywords[i]) !== -1) {
      return true;
    }
  }
  return false;
}

function isDateTimeHeader(name) {
  if (!name) return false;
  var n = String(name).trim().toLowerCase();
  return n === "thời gian" || n.indexOf("thời gian") !== -1;
}

function isDateHeader(name) {
  if (!name) return false;
  var n = String(name).trim().toLowerCase();
  return n.indexOf("ngày") !== -1;
}

/**
 * Đổ màu có điều kiện cho tất cả các dropdown trên bảng tính
 */
function applyConditionalFormatting(ss) {
  initializeColumnEnums();
  var branches = getBranchesList();
  var brands = getBrandsList();
  var colorMap = _buildColorMap(branches, brands);

  _applyNhanVienFormatting(ss, colorMap);
  _applyDienThoaiFormatting(ss, colorMap, brands, branches);
  _applyPhuKienFormatting(ss, colorMap, brands, branches);
  _applyDonHangFormatting(ss, colorMap, branches);
  _applyDichVuFormatting(ss, colorMap, branches);
  _applyTraGopFormatting(ss, colorMap, branches);
  _applyLichSuTraGopFormatting(ss, colorMap);
  _applyNhapKhoFormatting(ss, colorMap, branches);
  _applyDoiTraFormatting(ss, colorMap, branches);
  _applyThuMuaFormatting(ss, colorMap, brands, branches);
  _applyBaoHanhFormatting(ss, colorMap, branches);
}

function _buildColorMap(branches, brands) {
  var colorMap = {};
  for (var key in CONDITIONAL_COLOR_MAP) {
    colorMap[key] = CONDITIONAL_COLOR_MAP[key];
  }

  var BRANCH_COLORS = [
    { bg: "#e8f0fe", fg: "#1a73e8" }, // Blue
    { bg: "#fce8e6", fg: "#c5221f" }, // Red
    { bg: "#fef7e0", fg: "#b06000" }, // Yellow
    { bg: "#e6f4ea", fg: "#137333" }, // Green
    { bg: "#f3e8fd", fg: "#a142f4" }, // Purple
  ];
  branches.forEach(function (branch, index) {
    var colorIdx = index % BRANCH_COLORS.length;
    colorMap[branch] = BRANCH_COLORS[colorIdx];
  });

  var BRAND_COLORS = [
    { bg: "#e8f0fe", fg: "#1a73e8" }, // Blue
    { bg: "#e6f4ea", fg: "#137333" }, // Green
    { bg: "#fef7e0", fg: "#b06000" }, // Orange/Yellow
    { bg: "#f3e8fd", fg: "#a142f4" }, // Purple
    { bg: "#e0f7fa", fg: "#006064" }, // Cyan
    { bg: "#fce4ec", fg: "#880e4f" }, // Pink
    { bg: "#e8f8f5", fg: "#0e6251" }, // Teal
    { bg: "#fbf9ff", fg: "#6a1b9a" }, // Dark Violet
  ];
  brands.forEach(function (brand, index) {
    var lowerBrand = brand.toLowerCase();
    if (lowerBrand === "apple") {
      colorMap[brand] = { bg: "#cbd5e1", fg: "#1e293b" };
    } else if (lowerBrand === "khác") {
      colorMap[brand] = { bg: "#e2e8f0", fg: "#475569" };
    } else {
      var colorIdx = index % BRAND_COLORS.length;
      colorMap[brand] = BRAND_COLORS[colorIdx];
    }
  });

  // Bổ sung các giá trị đặc biệt cần tô màu
  colorMap["__NOT_EMPTY__"] = { bg: "#ffffff", fg: "#334155" };

  return colorMap;
}

function _applyRulesForSheet(sheet, mappings, colorMap) {
  if (!sheet) return;
  var rules = [];
  mappings.forEach(function (map) {
    var range = sheet.getRange(map.range);
    map.values.forEach(function (val) {
      var color = colorMap[val];
      if (color) {
        var builder = SpreadsheetApp.newConditionalFormatRule().setRanges([
          range,
        ]);
        if (val === "__NOT_EMPTY__") {
          builder.whenCellNotEmpty();
        } else {
          builder.whenTextEqualTo(val);
        }
        var rule = builder
          .setBackground(color.bg)
          .setFontColor(color.fg)
          .build();
        rules.push(rule);
      }
    });
  });
  sheet.setConditionalFormatRules(rules);
}

function _applyNhanVienFormatting(ss, colorMap) {
  var sheet = ss.getSheetByName(SHEET_NAMES.NHAN_VIEN);
  var colVaiTro = columnToLetter(COL_NV.VAI_TRO);
  var colQuyen = columnToLetter(COL_NV.QUYEN_XUAT);
  var colTrangThai = columnToLetter(COL_NV.TRANG_THAI);

  _applyRulesForSheet(sheet, [
    { range: colVaiTro + "2:" + colVaiTro, values: ["Bán hàng", "Kế toán", "Kỹ thuật"] },
    { range: colQuyen + "2:" + colQuyen, values: ["✓", "✗"] },
    { range: colTrangThai + "2:" + colTrangThai, values: ["Đang làm", "Nghỉ việc"] },
  ], colorMap);
}

function _applyDienThoaiFormatting(ss, colorMap, brands, branches) {
  var sheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  var colThuongHieuLetter = columnToLetter(COL_DT.THUONG_HIEU);
  var colTinhTrangLetter = columnToLetter(COL_DT.TINH_TRANG);
  var colTrangThaiLetter = columnToLetter(COL_DT.TRANG_THAI_KHO);
  var colChiNhanhLetter = columnToLetter(COL_DT.CHI_NHANH);

  _applyRulesForSheet(sheet, [
    { range: colThuongHieuLetter + "2:" + colThuongHieuLetter, values: brands },
    { range: colTinhTrangLetter + "2:" + colTinhTrangLetter, values: ["Mới 100%", "Like New", "__NOT_EMPTY__"] },
    {
      range: colTrangThaiLetter + "2:" + colTrangThaiLetter,
      values: ["Còn hàng", "Đã bán", "Đang trả góp", "Đã trả lại"],
    },
    { range: colChiNhanhLetter + "2:" + colChiNhanhLetter, values: branches },
  ], colorMap);
}

function _applyPhuKienFormatting(ss, colorMap, brands, branches) {
  var sheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  var colPKLoaiPK = columnToLetter(COL_PK.LOAI_PK);
  var colPKThuongHieu = columnToLetter(COL_PK.THUONG_HIEU);
  var colPKTrangThai = columnToLetter(COL_PK.TRANG_THAI);
  var colPKChiNhanh = columnToLetter(COL_PK.CHI_NHANH);

  _applyRulesForSheet(sheet, [
    {
      range: colPKLoaiPK + "2:" + colPKLoaiPK,
      values: ["Sạc", "Ốp lưng", "Tai nghe", "Cường lực", "Cáp", "Khác"],
    },
    { range: colPKThuongHieu + "2:" + colPKThuongHieu, values: brands },
    { range: colPKTrangThai + "2:" + colPKTrangThai, values: ["Đang bán", "Ngừng bán"] },
    { range: colPKChiNhanh + "2:" + colPKChiNhanh, values: branches },
  ], colorMap);
}

function _applyDonHangFormatting(ss, colorMap, branches) {
  var sheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  var colDHNguonSP = columnToLetter(COL_DH.NGUON_SP);
  var colDHHinhThucBan = columnToLetter(COL_DH.HINH_THUC_BAN);
  var colDHHTTT = columnToLetter(COL_DH.HINH_THUC_TT);
  var colDHTrangThai = columnToLetter(COL_DH.TRANG_THAI);
  var colDHChiNhanh = columnToLetter(COL_DH.CHI_NHANH);
  var colDHCoNhanQua = columnToLetter(COL_DH.CO_NHAN_QUA);

  _applyRulesForSheet(sheet, [
    { range: colDHNguonSP + "2:" + colDHNguonSP, values: ["Điện thoại", "Phụ kiện"] },
    { range: colDHHinhThucBan + "2:" + colDHHinhThucBan, values: ["Bán thẳng", "Trả góp"] },
    { range: colDHHTTT + "2:" + colDHHTTT, values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"] },
    { range: colDHTrangThai + "2:" + colDHTrangThai, values: ["Hoàn thành", "Đang xử lý", "Huỷ", "Đổi trả"] },
    { range: colDHChiNhanh + "2:" + colDHChiNhanh, values: branches },
    { range: colDHCoNhanQua + "2:" + colDHCoNhanQua, values: ["✓", "✗"] },
  ], colorMap);
}

function _applyDichVuFormatting(ss, colorMap, branches) {
  var sheet = ss.getSheetByName(SHEET_NAMES.DICH_VU);
  var colDVLoaiDV = columnToLetter(COL_DV.LOAI_DV);
  var colDVHTTT = columnToLetter(COL_DV.HINH_THUC_TT);
  var colDVTrangThai = columnToLetter(COL_DV.TRANG_THAI);
  var colDVChiNhanh = columnToLetter(COL_DV.CHI_NHANH);

  _applyRulesForSheet(sheet, [
    {
      range: colDVLoaiDV + "2:" + colDVLoaiDV,
      values: ["Chuyển khoản hộ", "Rút tiền mặt", "Nạp thẻ điện thoại"],
    },
    { range: colDVHTTT + "2:" + colDVHTTT, values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"] },
    { range: colDVTrangThai + "2:" + colDVTrangThai, values: ["Hoàn thành", "Huỷ"] },
    { range: colDVChiNhanh + "2:" + colDVChiNhanh, values: branches },
  ], colorMap);
}

function _applyTraGopFormatting(ss, colorMap, branches) {
  var sheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
  var colTGLoaiTG = columnToLetter(COL_TG.LOAI_TRA_GOP);
  var colTGTrangThai = columnToLetter(COL_TG.TRANG_THAI);
  var colTGChiNhanh = columnToLetter(COL_TG.CHI_NHANH);

  _applyRulesForSheet(sheet, [
    { range: colTGLoaiTG + "2:" + colTGLoaiTG, values: ["Cửa hàng", "Công ty tài chính"] },
    { range: colTGTrangThai + "2:" + colTGTrangThai, values: ["Đang trả", "Hoàn tất", "Quá hạn", "Đã huỷ"] },
    { range: colTGChiNhanh + "2:" + colTGChiNhanh, values: branches },
  ], colorMap);
}

function _applyLichSuTraGopFormatting(ss, colorMap) {
  var sheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
  var colLSTGHTTT = columnToLetter(COL_LSTG.HINH_THUC_TT);
  var colLSTGTrangThai = columnToLetter(COL_LSTG.TRANG_THAI);

  _applyRulesForSheet(sheet, [
    { range: colLSTGHTTT + "2:" + colLSTGHTTT, values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"] },
    { range: colLSTGTrangThai + "2:" + colLSTGTrangThai, values: ["Đã trả", "Chưa trả", "Quá hạn", "Đã huỷ"] },
  ], colorMap);
}

function _applyNhapKhoFormatting(ss, colorMap, branches) {
  var sheet = ss.getSheetByName(SHEET_NAMES.NHAP_KHO);
  _applyRulesForSheet(sheet, [
    { range: "C2:C", values: ["Điện thoại", "Phụ kiện"] },
    { range: "K2:K", values: branches },
  ], colorMap);
}

function _applyDoiTraFormatting(ss, colorMap, branches) {
  var sheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
  var colDTRLoaiGD = columnToLetter(COL_DT_TRA.LOAI_GD);
  var colDTRHTTT = columnToLetter(COL_DT_TRA.HINH_THUC_TT);
  var colDTRChiNhanh = columnToLetter(COL_DT_TRA.CHI_NHANH);
  var colDTRTrangThai = columnToLetter(COL_DT_TRA.TRANG_THAI);

  _applyRulesForSheet(sheet, [
    { range: colDTRLoaiGD + "2:" + colDTRLoaiGD, values: ["Trả máy", "Đổi máy"] },
    { range: colDTRHTTT + "2:" + colDTRHTTT, values: ["Tiền mặt", "Chuyển khoản", "Hỗn hợp"] },
    { range: colDTRChiNhanh + "2:" + colDTRChiNhanh, values: branches },
    { range: colDTRTrangThai + "2:" + colDTRTrangThai, values: ["Hoàn thành", "Huỷ"] },
  ], colorMap);
}

function _applyThuMuaFormatting(ss, colorMap, brands, branches) {
  var sheet = ss.getSheetByName(SHEET_NAMES.THU_MUA);
  var colTMBrand = columnToLetter(COL_TM.THUONG_HIEU_THU);
  var colTMTinhTrang = columnToLetter(COL_TM.TINH_TRANG_THU);
  var colTMLoaiGD = columnToLetter(COL_TM.LOAI_GD);
  var colTMHTTT = columnToLetter(COL_TM.HINH_THUC_TT);
  var colTMChiNhanh = columnToLetter(COL_TM.CHI_NHANH);

  _applyRulesForSheet(sheet, [
    { range: colTMBrand + "2:" + colTMBrand, values: brands },
    { range: colTMTinhTrang + "2:" + colTMTinhTrang, values: ["Mới 100%", "Like New", "__NOT_EMPTY__"] },
    { range: colTMLoaiGD + "2:" + colTMLoaiGD, values: ["Bán thẳng", "Thu cũ đổi mới"] },
    { range: colTMHTTT + "2:" + colTMHTTT, values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Trả góp", "Hỗn hợp"] },
    { range: colTMChiNhanh + "2:" + colTMChiNhanh, values: branches },
  ], colorMap);
}

function _applyBaoHanhFormatting(ss, colorMap, branches) {
  var sheet = ss.getSheetByName(SHEET_NAMES.BAO_HANH);
  var colBHLoaiDV = columnToLetter(COL_BH.LOAI_DICH_VU);
  var colBHHTTT = columnToLetter(COL_BH.HINH_THUC_TT);
  var colBHTrangThai = columnToLetter(COL_BH.TRANG_THAI);
  var colBHChiNhanh = columnToLetter(COL_BH.CHI_NHANH);

  _applyRulesForSheet(sheet, [
    { range: colBHLoaiDV + "2:" + colBHLoaiDV, values: ["Sửa chữa", "Bảo hành"] },
    {
      range: colBHHTTT + "2:" + colBHHTTT,
      values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"],
    },
    { range: colBHTrangThai + "2:" + colBHTrangThai, values: ["Đang xử lý", "Hoàn thành", "Huỷ"] },
    { range: colBHChiNhanh + "2:" + colBHChiNhanh, values: branches },
  ], colorMap);
}
