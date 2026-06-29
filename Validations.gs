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

  const branches = getBranchesList();
  const brands = getBrandsList();

  const chiNhanhRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(branches, true)
    .setAllowInvalid(false)
    .build();

  const thuongHieuRule = SpreadsheetApp.newDataValidation()
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

/**
 * Áp dụng quy tắc dữ liệu danh sách cho một cột cụ thể của sheet (DRY Compliance)
 * 
 * @param {Sheet} sheet Đối tượng sheet của Google
 * @param {number} colIndex Vị trí cột (1-indexed)
 * @param {string[]} listValues Danh sách các giá trị hợp lệ
 */
function setColumnListValidation(sheet, colIndex, listValues) {
  const colLetter = columnToLetter(colIndex);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(listValues, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(colLetter + "2:" + colLetter).setDataValidation(rule);
}

function _setupNhanVienValidations(ss, chiNhanhRule) {
  const nvSheet = ss.getSheetByName(SHEET_NAMES.NHAN_VIEN);
  if (!nvSheet) return;

  setColumnListValidation(nvSheet, COL_NV.VAI_TRO, ["Bán hàng", "Kế toán", "Kỹ thuật"]);
  setColumnListValidation(nvSheet, COL_NV.QUYEN_XUAT, ["✓", "✗"]);
  setColumnListValidation(nvSheet, COL_NV.TRANG_THAI, ["Đang làm", "Nghỉ việc"]);
}

function _setupDienThoaiValidations(ss, chiNhanhRule, thuongHieuRule) {
  const dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  if (!dtSheet) return;

  const colTinhTrangLetter = columnToLetter(COL_DT.TINH_TRANG);
  dtSheet.getRange(colTinhTrangLetter + "2:" + colTinhTrangLetter).clearDataValidations();

  setColumnListValidation(dtSheet, COL_DT.TRANG_THAI_KHO, [STOCK_STATUS.IN_STOCK, STOCK_STATUS.SOLD, STOCK_STATUS.INSTALLMENT, STOCK_STATUS.RETURNED, STOCK_STATUS.FAULTY]);
  dtSheet.getRange(columnToLetter(COL_DT.CHI_NHANH) + "2:" + columnToLetter(COL_DT.CHI_NHANH)).setDataValidation(chiNhanhRule);
  dtSheet.getRange(columnToLetter(COL_DT.THUONG_HIEU) + "2:" + columnToLetter(COL_DT.THUONG_HIEU)).setDataValidation(thuongHieuRule);
}

function _setupPhuKienValidations(ss, chiNhanhRule, thuongHieuRule) {
  const pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  if (!pkSheet) return;

  // Cột Loại PK
  setColumnListValidation(pkSheet, COL_PK.LOAI_PK, ["Sạc", "Ốp lưng", "Tai nghe", "Cường lực", "Cáp", "Khác"]);
  setColumnListValidation(pkSheet, COL_PK.TRANG_THAI, [PK_STATUS.ACTIVE, PK_STATUS.INACTIVE]);
  pkSheet.getRange(columnToLetter(COL_PK.THUONG_HIEU) + "2:" + columnToLetter(COL_PK.THUONG_HIEU)).setDataValidation(thuongHieuRule);
  pkSheet.getRange(columnToLetter(COL_PK.CHI_NHANH) + "2:" + columnToLetter(COL_PK.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupDonHangValidations(ss, chiNhanhRule) {
  const dhSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  if (!dhSheet) return;

  setColumnListValidation(dhSheet, COL_DH.HINH_THUC_BAN, ["Bán thẳng", "Trả góp"]);
  setColumnListValidation(dhSheet, COL_DH.HINH_THUC_TT, [PAYMENT_METHOD.CASH, PAYMENT_METHOD.TRANSFER, PAYMENT_METHOD.POS, PAYMENT_METHOD.MIXED]);
  setColumnListValidation(dhSheet, COL_DH.NGUON_SP, ["Điện thoại", "Phụ kiện"]);
  setColumnListValidation(dhSheet, COL_DH.TRANG_THAI, [ORDER_STATUS.DONE, ORDER_STATUS.PROCESSING, ORDER_STATUS.CANCELLED, ORDER_STATUS.EXCHANGED]);
  setColumnListValidation(dhSheet, COL_DH.CO_NHAN_QUA, ["✓", "✗"]);
  dhSheet.getRange(columnToLetter(COL_DH.CHI_NHANH) + "2:" + columnToLetter(COL_DH.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupDichVuValidations(ss, chiNhanhRule) {
  const dvSheet = ss.getSheetByName(SHEET_NAMES.DICH_VU);
  if (!dvSheet) return;

  setColumnListValidation(dvSheet, COL_DV.LOAI_DV, ["Chuyển khoản hộ", "Rút tiền mặt", "Nạp thẻ điện thoại"]);
  setColumnListValidation(dvSheet, COL_DV.HINH_THUC_TT, [PAYMENT_METHOD.CASH, PAYMENT_METHOD.TRANSFER, PAYMENT_METHOD.POS, PAYMENT_METHOD.MIXED]);
  setColumnListValidation(dvSheet, COL_DV.TRANG_THAI, [ORDER_STATUS.DONE, ORDER_STATUS.CANCELLED]);
  dvSheet.getRange(columnToLetter(COL_DV.CHI_NHANH) + "2:" + columnToLetter(COL_DV.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupTraGopValidations(ss, chiNhanhRule) {
  const tgSheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
  if (!tgSheet) return;

  setColumnListValidation(tgSheet, COL_TG.LOAI_TRA_GOP, ["Cửa hàng", "Công ty tài chính"]);
  setColumnListValidation(tgSheet, COL_TG.TRANG_THAI, [INSTALLMENT_STATUS.RUNNING, INSTALLMENT_STATUS.DONE, INSTALLMENT_STATUS.LATE, INSTALLMENT_STATUS.CANCELLED]);
  tgSheet.getRange(columnToLetter(COL_TG.CHI_NHANH) + "2:" + columnToLetter(COL_TG.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupLichSuTraGopValidations(ss) {
  const lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
  if (!lstgSheet) return;

  setColumnListValidation(lstgSheet, COL_LSTG.HINH_THUC_TT, [PAYMENT_METHOD.CASH, PAYMENT_METHOD.TRANSFER, PAYMENT_METHOD.POS, PAYMENT_METHOD.MIXED]);
  setColumnListValidation(lstgSheet, COL_LSTG.TRANG_THAI, [LSTG_STATUS.PAID, LSTG_STATUS.UNPAID, LSTG_STATUS.LATE, LSTG_STATUS.CANCELLED]);
}

function _setupNhapKhoValidations(ss, chiNhanhRule) {
  const nkSheet = ss.getSheetByName(SHEET_NAMES.NHAP_KHO);
  if (!nkSheet) return;

  setColumnListValidation(nkSheet, COL_NK.NGUON_NHAP, ["Điện thoại", "Phụ kiện"]); // Cột nguồn nhập là loại sản phẩm
  const colBranchLetter = columnToLetter(COL_NK.CHI_NHANH);
  nkSheet.getRange(colBranchLetter + "2:" + colBranchLetter).setDataValidation(chiNhanhRule);
}

function _setupDoiTraValidations(ss, chiNhanhRule) {
  const doiTraSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
  if (!doiTraSheet) return;

  setColumnListValidation(doiTraSheet, COL_DT_TRA.LOAI_GD, ["Trả máy", "Đổi máy"]);
  setColumnListValidation(doiTraSheet, COL_DT_TRA.HINH_THUC_TT, [PAYMENT_METHOD.CASH, PAYMENT_METHOD.TRANSFER, PAYMENT_METHOD.MIXED]);
  setColumnListValidation(doiTraSheet, COL_DT_TRA.TRANG_THAI, [ORDER_STATUS.DONE, ORDER_STATUS.CANCELLED]);
  doiTraSheet.getRange(columnToLetter(COL_DT_TRA.CHI_NHANH) + "2:" + columnToLetter(COL_DT_TRA.CHI_NHANH)).setDataValidation(chiNhanhRule);
}

function _setupThuMuaValidations(ss, chiNhanhRule, thuongHieuRule) {
  const thuMuaSheet = ss.getSheetByName(SHEET_NAMES.THU_MUA);
  if (!thuMuaSheet) return;

  const colTMTinhTrang = columnToLetter(COL_TM.TINH_TRANG_THU);
  thuMuaSheet.getRange(colTMTinhTrang + "2:" + colTMTinhTrang).clearDataValidations();

  setColumnListValidation(thuMuaSheet, COL_TM.LOAI_GD, ["Bán thẳng", "Thu cũ đổi mới"]);
  setColumnListValidation(thuMuaSheet, COL_TM.HINH_THUC_TT, ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Trả góp", "Hỗn hợp", "Trừ vào đơn mới"]);
  setColumnListValidation(thuMuaSheet, COL_TM.TRANG_THAI, [ORDER_STATUS.PROCESSING, ORDER_STATUS.DONE, ORDER_STATUS.CANCELLED]);
  thuMuaSheet.getRange(columnToLetter(COL_TM.CHI_NHANH) + "2:" + columnToLetter(COL_TM.CHI_NHANH)).setDataValidation(chiNhanhRule);
  thuMuaSheet.getRange(columnToLetter(COL_TM.THUONG_HIEU_THU) + "2:" + columnToLetter(COL_TM.THUONG_HIEU_THU)).setDataValidation(thuongHieuRule);
}

function _setupBaoCaoDoanhSoValidations(ss) {
  const reportDS = ss.getSheetByName(SHEET_NAMES.BAO_CAO_DOANH_SO);
  if (!reportDS) return;

  const staffList = ["Tất cả"];
  const nvSheet = ss.getSheetByName(SHEET_NAMES.NHAN_VIEN);
  if (nvSheet) {
    const lastRow = nvSheet.getLastRow();
    if (lastRow > 1) {
      const nvData = nvSheet.getRange(2, 1, lastRow - 1, 8).getValues();
      const c_maNV = COL_NV.MA_NV - 1;
      const c_tenNV = COL_NV.HO_TEN - 1;
      const c_trangThai = COL_NV.TRANG_THAI - 1;
      nvData.forEach(function (row) {
        if (row.length <= c_trangThai) return;
        const maNV = String(row[c_maNV]).trim();
        const tenNV = String(row[c_tenNV]).trim();
        const trangThai = String(row[c_trangThai]).trim();
        if (maNV && trangThai !== "Nghỉ việc") {
          staffList.push(maNV + " - " + tenNV);
        }
      });
    }
  }
  const staffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(staffList, true)
    .setAllowInvalid(false)
    .build();
  reportDS.getRange(3, 6).setDataValidation(staffRule);

  const gdList = [
    "Tất cả",
    "Đơn hàng (Bán máy)",
    "Đơn hàng (Hỗ trợ)",
    "Dịch vụ: Chuyển khoản hộ",
    "Dịch vụ: Rút tiền mặt",
    "Dịch vụ: Nạp thẻ điện thoại",
    "Dịch vụ: Sửa chữa",
    "Dịch vụ: Bảo hành",
  ];
  const gdRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(gdList, true)
    .setAllowInvalid(false)
    .build();
  reportDS.getRange(3, 8).setDataValidation(gdRule);
}

function _setupDoanhSoValidations(ss) {
  // Doanh số sheet does not require data validation setup
}

function _setupBaoHanhValidations(ss, chiNhanhRule) {
  const bhSheet = ss.getSheetByName(SHEET_NAMES.BAO_HANH);
  if (!bhSheet) return;

  setColumnListValidation(bhSheet, COL_BH.LOAI_DICH_VU, ["Sửa chữa", "Bảo hành"]);
  setColumnListValidation(bhSheet, COL_BH.HINH_THUC_TT, [PAYMENT_METHOD.CASH, PAYMENT_METHOD.TRANSFER, PAYMENT_METHOD.POS, PAYMENT_METHOD.MIXED]);
  setColumnListValidation(bhSheet, COL_BH.TRANG_THAI, [ORDER_STATUS.PROCESSING, ORDER_STATUS.DONE, ORDER_STATUS.CANCELLED]);
  bhSheet.getRange(columnToLetter(COL_BH.CHI_NHANH) + "2:" + columnToLetter(COL_BH.CHI_NHANH)).setDataValidation(chiNhanhRule);
}
