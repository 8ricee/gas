/**
 * ============================================================
 * VanTran Mobile — ChuyenKho.gs
 * Nghiệp vụ điều chuyển sản phẩm giữa các chi nhánh
 * ============================================================
 */

/**
 * Chuyển kho sản phẩm giữa các chi nhánh
 *
 * @param {Object} data - {nguonSP ('Điện thoại'/'Phụ kiện'), maSP, chiNhanhNguon, chiNhanhDich, soLuong}
 * @return {boolean}
 */
function chuyenKho(data) {
  var nguonSP = data.nguonSP || "Điện thoại";
  var maSP = data.maSP;
  var cnNguon = data.chiNhanhNguon;
  var cnDich = data.chiNhanhDich;
  var soLuong = Number(data.soLuong) || 1;

  if (!maSP || !cnNguon || !cnDich) {
    throw new Error(
      "Vui lòng chọn đầy đủ sản phẩm, chi nhánh nguồn và chi nhánh đích!",
    );
  }

  if (cnNguon === cnDich) {
    throw new Error("Chi nhánh nguồn và chi nhánh đích phải khác nhau!");
  }

  if (nguonSP === "Điện thoại") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
    var row = -1;

    if (data.imei) {
      row = findRow(SHEET_NAMES.DIEN_THOAI, 4, data.imei);
    } else {
      // Tự động tìm máy có mã SP tại chi nhánh nguồn và đang "Còn hàng"
      var dtData = sheet.getDataRange().getValues();
      for (var i = 1; i < dtData.length; i++) {
        if (String(dtData[i][0]) === maSP && 
            String(dtData[i][12]) === cnNguon && 
            String(dtData[i][10]) === "Còn hàng") {
          row = i + 1;
          data.imei = String(dtData[i][3]); // Lưu lại IMEI tìm được
          break;
        }
      }
      if (row === -1) {
        row = findRow(SHEET_NAMES.DIEN_THOAI, 1, maSP);
      }
    }

    if (row === -1) throw new Error("Không tìm thấy điện thoại: " + (data.imei ? "IMEI: " + data.imei : maSP));

    var currentBranch = sheet.getRange(row, 13).getValue();
    var currentStatus = sheet.getRange(row, 11).getValue();

    if (currentBranch !== cnNguon) {
      throw new Error(
        "Điện thoại không nằm ở chi nhánh nguồn " +
          cnNguon +
          " (Hiện tại ở: " +
          currentBranch +
          ")",
      );
    }
    if (currentStatus !== "Còn hàng") {
      throw new Error(
        'Điện thoại không ở trạng thái "Còn hàng" (Hiện tại: ' +
          currentStatus +
          ")",
      );
    }

    // 2. Chuyển chi nhánh
    updateCell(SHEET_NAMES.DIEN_THOAI, row, 13, cnDich);
  } else {
    // Phụ kiện
    // 1. Trừ kho ở chi nhánh nguồn
    updateTonKhoPhuKien(maSP, soLuong, "xuat", cnNguon);

    // 2. Cộng kho ở chi nhánh đích
    updateTonKhoPhuKien(maSP, soLuong, "nhap", cnDich);
  }

  showToast(
    "✅ Chuyển kho thành công: " + maSP + " (" + cnNguon + " ➔ " + cnDich + ")",
  );
  return true;
}
