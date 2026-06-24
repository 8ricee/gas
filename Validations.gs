/**
 * ============================================================
 * VanTran Mobile — Validations.gs
 * Cấu hình xác thực dữ liệu (Data Validation) cho các Sheet
 * ============================================================
 */

/**
 * Thiết lập Data Validation cho các sheet
 */
function _setupDataValidations(ss) {
  initializeColumnEnums();

  var branches = getBranchesList();
  var brands = getBrandsList();

  var chiNhanhRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(branches, true)
    .setAllowInvalid(false)
    .build();

  var thuongHieuRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(brands, true)
    .setAllowInvalid(false)
    .build();

  _setupNhanVienValidations(ss, chiNhanhRule);
  _setupDienThoaiValidations(ss, chiNhanhRule, thuongHieuRule);
  _setupPhuKienValidations(ss, chiNhanhRule, thuongHieuRule);
  _setupDonHangValidations(ss, chiNhanhRule);
  _setupDichVuValidations(ss, chiNhanhRule);
  _setupTraGopValidations(ss, chiNhanhRule);
  _setupLichSuTraGopValidations(ss);
  _setupNhapKhoValidations(ss, chiNhanhRule);
  _setupDoiTraValidations(ss, chiNhanhRule);
  _setupThuMuaValidations(ss, chiNhanhRule, thuongHieuRule);
  _setupBaoCaoDoanhSoValidations(ss);
  _setupDoanhSoValidations(ss);
  _setupBaoHanhValidations(ss, chiNhanhRule);
}

function _clearSheetDataValidations(sheet) {
  if (!sheet) return;
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  if (maxRows > 1 && maxCols > 0) {
    sheet.getRange(2, 1, maxRows - 1, maxCols).clearDataValidations();
  }
}

/**
 * Áp dụng quy tắc dữ liệu danh sách cho một cột cụ thể của sheet (DRY Compliance)
 * 
 * @param {Sheet} sheet Đối tượng sheet của Google
 * @param {number} colIndex Vị trí cột (1-indexed)
 * @param {string[]} listValues Danh sách các giá trị hợp lệ
 */
function setColumnListValidation(sheet, colIndex, listValues) {
  var colLetter = columnToLetter(colIndex);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(listValues, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(colLetter + "2:" + colLetter).setDataValidation(rule);
}

function _setupNhanVienValidations(ss, chiNhanhRule) {
  var nvSheet = ss.getSheetByName(SHEET_NAMES.NHAN_VIEN);
  if (!nvSheet) return;
  _clearSheetDataValidations(nvSheet);

  setColumnListValidation(nvSheet, COL_NV.VAI_TRO, ["Bán hàng", "Kế toán", "Kỹ thuật"]);
  setColumnListValidation(nvSheet, COL_NV.QUYEN_XUAT, ["✓", "✗"]);
  setColumnListValidation(nvSheet, COL_NV.TRANG_THAI, ["Đang làm", "Nghỉ việc"]);
}

function _setupDienThoaiValidations(ss, chiNhanhRule, thuongHieuRule) {
  var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  if (!dtSheet) return;
  _clearSheetDataValidations(dtSheet);

  var colTinhTrangLetter = columnToLetter(COL_DT.TINH_TRANG);
  dtSheet.getRange(colTinhTrangLetter + "2:" + colTinhTrangLetter).clearDataValidations();

  setColumnListValidation(dtSheet, COL_DT.TRANG_THAI_KHO, ["Còn hàng", "Đã bán", "Đang trả góp", "Đã trả lại"]);
  dtSheet.getRange(columnToLetter(COL_DT.CHI_NHANH) + "2:" + columnToLetter(COL_DT.CHI_NHANH)).setDataValidation(chiNhanhRule);
  dtSheet.getRange(columnToLetter(COL_DT.THUONG_HIEU) + "2:" + columnToLetter(COL_DT.THUONG_HIEU)).setDataValidation(thuongHieuRule);
}

function _setupPhuKienValidations(ss, chiNhanhRule, thuongHieuRule) {
  var pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  if (!pkSheet) return;
  _clearSheetDataValidations(pkSheet);

  // Cột Loại PK (cột C - cột 3)
  setColumnListValidation(pkSheet, 3, ["Sạc", "Ốp lưng", "Tai nghe", "Cường lực", "Cáp", "Khác"]);
  setColumnListValidation(pkSheet, COL_PK.TRANG_THAI, ["Đang bán", "Ngừng bán"]);
  pkSheet.getRange(columnToLetter(COL_PK.THUONG_HIEU) + "2:" + columnToLetter(COL_PK.THUONG_HIEU)).setDataValidation(thuongHieuRule);
  pkSheet.getRange(columnToLetter(COL_PK.CHI_NHANH) + "2:" + columnToLetter(COL_PK.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupDonHangValidations(ss, chiNhanhRule) {
  var dhSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  if (!dhSheet) return;
  _clearSheetDataValidations(dhSheet);

  setColumnListValidation(dhSheet, COL_DH.HINH_THUC_BAN, ["Bán thẳng", "Trả góp"]);
  setColumnListValidation(dhSheet, COL_DH.HINH_THUC_TT, ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"]);
  setColumnListValidation(dhSheet, COL_DH.NGUON_SP, ["Điện thoại", "Phụ kiện"]);
  setColumnListValidation(dhSheet, COL_DH.TRANG_THAI, ["Hoàn thành", "Đang xử lý", "Huỷ", "Đổi trả"]);
  setColumnListValidation(dhSheet, COL_DH.CO_NHAN_QUA, ["✓", "✗"]);
  dhSheet.getRange(columnToLetter(COL_DH.CHI_NHANH) + "2:" + columnToLetter(COL_DH.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupDichVuValidations(ss, chiNhanhRule) {
  var dvSheet = ss.getSheetByName(SHEET_NAMES.DICH_VU);
  if (!dvSheet) return;
  _clearSheetDataValidations(dvSheet);

  setColumnListValidation(dvSheet, COL_DV.LOAI_DV, ["Chuyển khoản hộ", "Rút tiền mặt", "Nạp thẻ điện thoại"]);
  setColumnListValidation(dvSheet, COL_DV.HINH_THUC_TT, ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"]);
  setColumnListValidation(dvSheet, COL_DV.TRANG_THAI, ["Hoàn thành", "Huỷ"]);
  dvSheet.getRange(columnToLetter(COL_DV.CHI_NHANH) + "2:" + columnToLetter(COL_DV.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupTraGopValidations(ss, chiNhanhRule) {
  var tgSheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
  if (!tgSheet) return;
  _clearSheetDataValidations(tgSheet);

  setColumnListValidation(tgSheet, COL_TG.LOAI_TRA_GOP, ["Cửa hàng", "Công ty tài chính"]);
  setColumnListValidation(tgSheet, COL_TG.TRANG_THAI, ["Đang trả", "Hoàn tất", "Quá hạn", "Đã huỷ"]);
  tgSheet.getRange(columnToLetter(COL_TG.CHI_NHANH) + "2:" + columnToLetter(COL_TG.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupLichSuTraGopValidations(ss) {
  var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
  if (!lstgSheet) return;
  _clearSheetDataValidations(lstgSheet);

  setColumnListValidation(lstgSheet, COL_LSTG.HINH_THUC_TT, ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"]);
  setColumnListValidation(lstgSheet, COL_LSTG.TRANG_THAI, ["Đã trả", "Chưa trả", "Quá hạn", "Đã huỷ"]);
}

function _setupNhapKhoValidations(ss, chiNhanhRule) {
  var nkSheet = ss.getSheetByName(SHEET_NAMES.NHAP_KHO);
  if (!nkSheet) return;
  _clearSheetDataValidations(nkSheet);

  setColumnListValidation(nkSheet, 3, ["Điện thoại", "Phụ kiện"]); // Cột C là loại sản phẩm
  nkSheet.getRange("K2:K").setDataValidation(chiNhanhRule);
}

function _setupDoiTraValidations(ss, chiNhanhRule) {
  var doiTraSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
  if (!doiTraSheet) return;
  _clearSheetDataValidations(doiTraSheet);

  setColumnListValidation(doiTraSheet, COL_DT_TRA.LOAI_GD, ["Trả máy", "Đổi máy"]);
  setColumnListValidation(doiTraSheet, COL_DT_TRA.HINH_THUC_TT, ["Tiền mặt", "Chuyển khoản", "Hỗn hợp"]);
  setColumnListValidation(doiTraSheet, COL_DT_TRA.TRANG_THAI, ["Hoàn thành", "Huỷ"]);
  doiTraSheet.getRange(columnToLetter(COL_DT_TRA.CHI_NHANH) + "2:" + columnToLetter(COL_DT_TRA.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupThuMuaValidations(ss, chiNhanhRule, thuongHieuRule) {
  var thuMuaSheet = ss.getSheetByName(SHEET_NAMES.THU_MUA);
  if (!thuMuaSheet) return;
  _clearSheetDataValidations(thuMuaSheet);

  var colTMTinhTrang = columnToLetter(COL_TM.TINH_TRANG_THU);
  thuMuaSheet.getRange(colTMTinhTrang + "2:" + colTMTinhTrang).clearDataValidations();

  setColumnListValidation(thuMuaSheet, COL_TM.LOAI_GD, ["Bán thẳng", "Thu cũ đổi mới"]);
  setColumnListValidation(thuMuaSheet, COL_TM.HINH_THUC_TT, ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Trả góp", "Hỗn hợp"]);
  thuMuaSheet.getRange(columnToLetter(COL_TM.CHI_NHANH) + "2:" + columnToLetter(COL_TM.CHI_NHANH)).setDataValidation(chiNhanhRule);
  thuMuaSheet.getRange(columnToLetter(COL_TM.THUONG_HIEU_THU) + "2:" + columnToLetter(COL_TM.THUONG_HIEU_THU)).setDataValidation(thuongHieuRule);
}

function _setupBaoCaoDoanhSoValidations(ss) {
  var reportDS = ss.getSheetByName(SHEET_NAMES.BAO_CAO_DOANH_SO);
  if (!reportDS) return;

  var staffList = ["Tất cả"];
  var nvSheet = ss.getSheetByName(SHEET_NAMES.NHAN_VIEN);
  if (nvSheet) {
    initializeColumnEnums();
    var lastRow = nvSheet.getLastRow();
    if (lastRow > 1) {
      var nvData = nvSheet.getRange(2, 1, lastRow - 1, 8).getValues();
      var c_maNV = COL_NV.MA_NV - 1;
      var c_tenNV = COL_NV.HO_TEN - 1;
      var c_trangThai = COL_NV.TRANG_THAI - 1;
      nvData.forEach(function (row) {
        if (row.length <= c_trangThai) return;
        var maNV = String(row[c_maNV]).trim();
        var tenNV = String(row[c_tenNV]).trim();
        var trangThai = String(row[c_trangThai]).trim();
        if (maNV && trangThai !== "Nghỉ việc") {
          staffList.push(maNV + " - " + tenNV);
        }
      });
    }
  }
  var staffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(staffList, true)
    .setAllowInvalid(false)
    .build();
  reportDS.getRange(3, 6).setDataValidation(staffRule);

  var gdList = [
    "Tất cả",
    "Đơn hàng (Bán máy)",
    "Đơn hàng (Hỗ trợ)",
    "Dịch vụ: Chuyển khoản hộ",
    "Dịch vụ: Rút tiền mặt",
    "Dịch vụ: Nạp thẻ điện thoại",
    "Dịch vụ: Sửa chữa",
    "Dịch vụ: Bảo hành",
  ];
  var gdRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(gdList, true)
    .setAllowInvalid(false)
    .build();
  reportDS.getRange(3, 8).setDataValidation(gdRule);
}

function _setupDoanhSoValidations(ss) {
  var dsSheet = ss.getSheetByName(SHEET_NAMES.DOANH_SO);
  if (dsSheet) {
    _clearSheetDataValidations(dsSheet);
  }
}

function _setupBaoHanhValidations(ss, chiNhanhRule) {
  var bhSheet = ss.getSheetByName(SHEET_NAMES.BAO_HANH);
  if (!bhSheet) return;
  _clearSheetDataValidations(bhSheet);

  setColumnListValidation(bhSheet, COL_BH.LOAI_DICH_VU, ["Sửa chữa", "Bảo hành"]);
  setColumnListValidation(bhSheet, COL_BH.HINH_THUC_TT, ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"]);
  setColumnListValidation(bhSheet, COL_BH.TRANG_THAI, ["Đang xử lý", "Hoàn thành", "Huỷ"]);
  bhSheet.getRange(columnToLetter(COL_BH.CHI_NHANH) + "2:" + columnToLetter(COL_BH.CHI_NHANH)).setDataValidation(chiNhanhRule);
}
