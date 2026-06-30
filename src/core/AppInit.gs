/**
 * ============================================================
 * VanTran Mobile — AppInit.gs
 * Logic khởi tạo hệ thống, chuẩn hóa headers & cài đặt cấu hình
 * ============================================================
 */

/**
 * Khởi tạo toàn bộ hệ thống: tạo 13 sheet với header + dữ liệu cấu hình mặc định
 */
function setupSheets() {
  clearColumnEnumsCache();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let runSetup = false;

  try {
    const ui = SpreadsheetApp.getUi();
    if (ui) {
      const result = ui.alert(
        "Khởi tạo hệ thống",
        "Hệ thống sẽ tạo các sheet mới.\n\nCác sheet đã tồn tại sẽ KHÔNG bị ghi đè.\n\nBạn có muốn tiếp tục?",
        ui.ButtonSet.YES_NO,
      );
      if (result === ui.Button.YES) {
        runSetup = true;
      } else {
        ui.alert("Đã hủy khởi tạo.");
        return;
      }
    } else {
      runSetup = true;
    }
  } catch (e) {
    runSetup = true;
  }

  if (!runSetup) return;

  const sheetOrder = SHEET_ORDER;

  let createdCount = 0;

  sheetOrder.forEach(function (name) {
    const existing = ss.getSheetByName(name);
    if (!existing) {
      ss.insertSheet(name);
      createdCount++;
    }
  });

  // Chèn cột "Trừ tiền thu máy" ở đúng vị trí mới (nếu có dữ liệu cũ cần dịch chuyển)
  migrateTienThuMuaColumnPosition(ss);

  // Chèn cột "Trạng thái" ở sheet Thu mua ở đúng vị trí mới (nếu cần)
  migrateThuMuaStatusColumn(ss);

  // Đồng bộ tiêu đề cột cho tất cả các sheet (cũ & mới)
  syncSheetHeaders(ss);

  // Insert default config data if CauHinh is freshly created
  const cauHinhSheet = ss.getSheetByName(SHEET_NAMES.CAU_HINH);
  if (cauHinhSheet && cauHinhSheet.getLastRow() <= 1) {
    cauHinhSheet
      .getRange(2, 1, DEFAULT_CONFIG.length, DEFAULT_CONFIG[0].length)
      .setValues(DEFAULT_CONFIG);
  }

  // Setup data validations
  _setupDataValidations(ss);

  // Format font Times New Roman size 12 globally & Auto-resize columns
  formatAllSheets();

  // Setup conditional formatting colors
  applyConditionalFormatting(ss);

  // Khởi tạo sheet Tồn kho nếu vừa được tạo mới hoặc chưa có dữ liệu
  const tonKhoSheet = ss.getSheetByName(SHEET_NAMES.TON_KHO);
  if (tonKhoSheet && tonKhoSheet.getLastRow() <= 1) {
    rebuildTonKhoSheet(ss);
  }

  // Khởi tạo Báo cáo ngày nếu vừa được tạo mới hoặc chưa có dữ liệu
  const reportSheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_NGAY);
  if (reportSheet && reportSheet.getLastRow() <= 1) {
    reportSheet.getRange(3, 1).setValue("Ngày báo cáo:").setFontWeight("bold");
    reportSheet
      .getRange(3, 2)
      .setValue(new Date())
      .setNumberFormat("dd/MM/yyyy");
    reportSheet
      .getRange(4, 1)
      .setValue("Trạng thái cập nhật:")
      .setFontWeight("bold");
    try {
      updateDailyReportFromSheet();
    } catch (e) {
      Logger.log("Không thể chạy báo cáo ban đầu: " + e.message);
    }
  }

  // Khởi tạo Báo cáo doanh số nếu vừa được tạo mới hoặc chưa có dữ liệu
  const salesReportSheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_DOANH_SO);
  if (salesReportSheet && salesReportSheet.getLastRow() <= 1) {
    salesReportSheet.getRange(3, 1).setValue("Từ ngày:").setFontWeight("bold");
    const now = new Date();
    const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
    salesReportSheet
      .getRange(3, 2)
      .setValue(firstDay)
      .setNumberFormat("dd/MM/yyyy");
    salesReportSheet.getRange(3, 3).setValue("Đến ngày:").setFontWeight("bold");
    salesReportSheet.getRange(3, 4).setValue(now).setNumberFormat("dd/MM/yyyy");
    salesReportSheet
      .getRange(3, 5)
      .setValue("Nhân viên:")
      .setFontWeight("bold");
    salesReportSheet.getRange(3, 6).setValue("Tất cả");
    salesReportSheet
      .getRange(3, 7)
      .setValue("Loại GD/Dịch vụ:")
      .setFontWeight("bold");
    salesReportSheet.getRange(3, 8).setValue("Tất cả");
    salesReportSheet
      .getRange(4, 1)
      .setValue("Trạng thái cập nhật:")
      .setFontWeight("bold");
    try {
      updateSalesReportFromSheet();
    } catch (e) {
      Logger.log("Không thể chạy báo cáo doanh số ban đầu: " + e.message);
    }
  }

  // Remove default "Sheet1" if it exists and is empty
  const defaultSheet =
    ss.getSheetByName("Sheet1") ||
    ss.getSheetByName("Trang tính1") ||
    ss.getSheetByName("Sheet 1");
  if (defaultSheet && defaultSheet.getLastRow() === 0) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {
      Logger.log("[WARN] Không thể xóa sheet mặc định: " + e.message);
    }
  }

  showAlert(
    "Hoàn tất!",
    "Đã tạo " +
      createdCount +
      " sheet mới.\n\nHệ thống sẵn sàng hoạt động với font Times New Roman cỡ 12!",
  );
}

/**
 * Đồng bộ cấu trúc cột và tiêu đề cho tất cả các sheet dựa trên SHEET_HEADERS.
 * Nếu thiếu cột hoặc tiêu đề mới ở cuối, sẽ tự động chèn thêm và định dạng lại mà không làm mất dữ liệu cũ.
 */
function syncSheetHeaders(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrder = SHEET_ORDER;

  const updatedSheets = [];

  sheetOrder.forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    const targetHeaders = SHEET_HEADERS[name];
    if (!sheet || !targetHeaders || targetHeaders.length === 0) return;

    const maxCols = sheet.getMaxColumns();
    const lastCol = sheet.getLastColumn();
    let currentHeaders = [];

    if (lastCol > 0) {
      currentHeaders = sheet
        .getRange(1, 1, 1, lastCol)
        .getValues()[0]
        .map(function (h) {
          return String(h).trim();
        });
    }

    let needsUpdate = false;

    // Nếu số cột hiện tại ít hơn số tiêu đề định nghĩa
    if (maxCols < targetHeaders.length) {
      sheet.insertColumnsAfter(maxCols, targetHeaders.length - maxCols);
      needsUpdate = true;
    }

    // Kiểm tra từng tiêu đề cột xem có khớp hoặc thiếu không
    for (let colIdx = 0; colIdx < targetHeaders.length; colIdx++) {
      const expectedHeader = targetHeaders[colIdx];
      const actualHeader = currentHeaders[colIdx];

      if (String(actualHeader).trim() !== expectedHeader) {
        sheet.getRange(1, colIdx + 1).setValue(expectedHeader);
        needsUpdate = true;
      }
    }

    if (needsUpdate) {
      // Định dạng lại dòng tiêu đề cho đồng bộ
      const headerRange = sheet.getRange(1, 1, 1, targetHeaders.length);
      headerRange.setBackground("#1a73e8");
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("left");
      sheet.setFrozenRows(1);
      updatedSheets.push(name);
    }
  });

  return updatedSheets;
}

/**
 * Migration function to remove the redundant customer phone number columns in
 * existing Dịch vụ, Thu mua, and Bảo hành sheets, shifting remaining columns left.
 */
function migrateDeletePhoneColumns(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const targetSheets = [
    { name: SHEET_NAMES.DICH_VU, header: "Số điện thoại khách" },
    { name: SHEET_NAMES.THU_MUA, header: "Số điện thoại khách" },
    { name: SHEET_NAMES.BAO_HANH, header: "Số điện thoại" }
  ];
  
  let deletedCount = 0;
  
  targetSheets.forEach(function(target) {
    const sheet = ss.getSheetByName(target.name);
    if (!sheet) return;
    
    const lastCol = sheet.getLastColumn();
    if (lastCol <= 0) return;
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let colIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i]).trim().toLowerCase();
      if (h === target.header.toLowerCase()) {
        colIndex = i + 1;
        break;
      }
    }
    
    if (colIndex !== -1) {
      sheet.deleteColumn(colIndex);
      deletedCount++;
      Logger.log("Deleted column " + target.header + " in sheet " + target.name + " at index " + colIndex);
    }
  });
  
  if (deletedCount > 0) {
    // Clear dynamic column cache
    clearSheetCache();
    // Synchronize headers to reflect new column structure
    syncSheetHeaders(ss);
  }
}

/**
 * Migration function to move or insert the "Trừ tiền thu máy" column to its correct position (between "Đơn giá" and "Thành tiền")
 */
function migrateTienThuMuaColumnPosition(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  if (lastCol <= 0) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });

  const currentIdx = headers.indexOf("Trừ tiền thu máy");
  const donGiaIdx = headers.indexOf("Đơn giá");

  if (donGiaIdx === -1) return;

  const targetColNum = donGiaIdx + 2; // Column L (Index 12)

  if (currentIdx !== -1) {
    const currentColNum = currentIdx + 1;
    if (currentColNum === targetColNum) {
      // Already in the correct position
      return;
    }
    // Delete the column from its incorrect position
    sheet.deleteColumn(currentColNum);
  }

  // Insert the column at the correct position
  sheet.insertColumnBefore(targetColNum);
  sheet.getRange(1, targetColNum).setValue("Trừ tiền thu máy");
  
  clearSheetCache(SHEET_NAMES.DON_HANG);
}

/**
 * Đảm bảo các cấu hình mặc định (bao gồm cả các cấu hình mới thêm sau này)
 * tồn tại trong sheet Cấu hình.
 */
function ensureDefaultConfigsExist(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CAU_HINH);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  let keys = [];
  if (lastRow > 1) {
    keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(row) {
      return String(row[0]).trim().toLowerCase();
    });
  }

    DEFAULT_CONFIG.forEach(function(row) {
    const key = String(row[0]).trim().toLowerCase();
    if (keys.indexOf(key) === -1) {
      sheet.appendRow(row);
      const newRow = sheet.getLastRow();
      sheet.getRange(newRow, 1, 1, row.length)
        .setFontFamily("Times New Roman")
        .setFontSize(12)
        .setHorizontalAlignment("left");
    }
  });
}

/**
 * Migration function to extract IMEI from Ghi chú and save it to the new IMEI column.
 * It first synchronizes headers to add the IMEI column if missing.
 */
function migrateImeiToColumn(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  if (!sheet) return;
  
  // 1. Kiểm tra cấu trúc cột thực tế và chèn cột IMEI cạnh Tên sản phẩm
  let lastCol = sheet.getLastColumn();
  let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); }) : [];
  let tenSpIdx = -1;
  let imeiIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === "Tên sản phẩm") {
      tenSpIdx = i + 1;
    }
    if (headers[i] === "IMEI") {
      imeiIdx = i + 1;
    }
  }
  
  // Nếu có cột IMEI nhưng ở vị trí sai (không phải ngay sau Tên sản phẩm)
  if (imeiIdx !== -1 && imeiIdx !== tenSpIdx + 1) {
    sheet.deleteColumn(imeiIdx);
    imeiIdx = -1;
    // Cập nhật lại headers và lastCol sau khi xóa cột
    lastCol = sheet.getLastColumn();
    headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] === "Tên sản phẩm") {
        tenSpIdx = i + 1;
      }
    }
  }
  
  // Nếu chưa có cột IMEI, chèn ngay sau Tên sản phẩm
  if (imeiIdx === -1 && tenSpIdx !== -1) {
    sheet.insertColumnAfter(tenSpIdx);
    sheet.getRange(1, tenSpIdx + 1).setValue("IMEI");
  }
  
  // Reset cache và đồng bộ lại tiêu đề & cấu hình cột hệ thống
  clearColumnEnumsCache();
  initializeColumnEnums();
  syncSheetHeaders(ss);
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    showAlert("Thông báo", "Không có dữ liệu đơn hàng để di chuyển.");
    return;
  }
  
  // Định dạng cột IMEI là text thuần không dấu phân chia
  const colImeiLetter = columnToLetter(COL_DH.IMEI);
  sheet.getRange(colImeiLetter + "2:" + colImeiLetter).setNumberFormat("@");
  
  const currentLastCol = sheet.getLastColumn();
  const range = sheet.getRange(2, 1, lastRow - 1, currentLastCol);
  const values = range.getValues();
  
  let migratedCount = 0;
  
  for (let i = 0; i < values.length; i++) {
    const rowValues = values[i];
    const ghiChu = String(rowValues[COL_DH.GHI_CHU - 1] || "");
    const imeiMatch = ghiChu.match(/\[IMEI:\s*([^\s\]]+)\]/);
    
    if (imeiMatch) {
      const imei = imeiMatch[1].trim();
      // Điền IMEI dưới dạng text thuần
      rowValues[COL_DH.IMEI - 1] = "'" + imei;
      // Dọn dẹp ghi chú (loại bỏ tag IMEI)
      rowValues[COL_DH.GHI_CHU - 1] = ghiChu.replace(/\[IMEI:\s*([^\s\]]+)\]/, "").trim();
      migratedCount++;
    } else {
      // Đảm bảo các dòng cũ có IMEI nhưng không có tag (nếu có) cũng được định dạng text thuần
      const existingImei = String(rowValues[COL_DH.IMEI - 1] || "").trim();
      if (existingImei && /^\d+$/.test(existingImei)) {
        rowValues[COL_DH.IMEI - 1] = "'" + existingImei;
      }
    }
  }
  
  range.setValues(values);
  clearSheetCache(SHEET_NAMES.DON_HANG);
  Logger.log("Migrated " + migratedCount + " orders to IMEI column next to Product Name.");
  
  showAlert(
    "Thành công",
    "Đã di chuyển cột IMEI cạnh Tên sản phẩm, thiết lập định dạng text thuần và tách IMEI thành công cho " + migratedCount + " đơn hàng!"
  );
}

/**
 * Migration function to move or insert the "Trạng thái máy" column in Thu mua sheet
 */
function migrateThuMuaStatusColumn(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.THU_MUA);
  if (!sheet) return;

  const lastCol = sheet.getLastColumn();
  if (lastCol <= 0) return;

  let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });

  // Tìm cột "Trạng thái máy" trước
  let currentIdx = headers.indexOf("Trạng thái máy");

  // Nếu không thấy "Trạng thái máy", thử tìm "Trạng thái" cũ để đổi tên thành "Trạng thái máy"
  if (currentIdx === -1) {
    const oldIdx = headers.indexOf("Trạng thái");
    if (oldIdx !== -1) {
      sheet.getRange(1, oldIdx + 1).setValue("Trạng thái máy");
      headers[oldIdx] = "Trạng thái máy";
      currentIdx = oldIdx;
      clearSheetCache(SHEET_NAMES.THU_MUA);
    }
  }

  const nguoiThucHienIdx = headers.indexOf("Người thực hiện");
  if (nguoiThucHienIdx === -1) return;

  const targetColNum = nguoiThucHienIdx + 2; // Right after Người thực hiện

  if (currentIdx !== -1) {
    const currentColNum = currentIdx + 1;
    if (currentColNum === targetColNum) {
      // Already in the correct position
      return;
    }
    // Delete the column from its incorrect position
    sheet.deleteColumn(currentColNum);
  }

  // Insert the column at the correct position
  sheet.insertColumnBefore(targetColNum);
  sheet.getRange(1, targetColNum).setValue("Trạng thái máy");
  
  clearSheetCache(SHEET_NAMES.THU_MUA);

  // Điền giá trị mặc định "Đang xử lý" cho các dòng cũ
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    clearColumnEnumsCache();
    initializeColumnEnums();

    const range = sheet.getRange(2, COL_TM.TRANG_THAI, lastRow - 1, 1);
    const values = range.getValues();
    let updated = false;
    for (let i = 0; i < values.length; i++) {
      if (!values[i][0]) {
        values[i][0] = "Đang xử lý";
        updated = true;
      }
    }
    if (updated) {
      range.setValues(values);
      clearSheetCache(SHEET_NAMES.THU_MUA);
    }
  }
}

