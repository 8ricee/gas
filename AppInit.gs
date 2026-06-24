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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var runSetup = false;

  try {
    var ui = SpreadsheetApp.getUi();
    if (ui) {
      var result = ui.alert(
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

  var sheetOrder = SHEET_ORDER;

  var createdCount = 0;

  sheetOrder.forEach(function (name) {
    var existing = ss.getSheetByName(name);
    if (!existing) {
      ss.insertSheet(name);
      createdCount++;
    }
  });

  // Đồng bộ tiêu đề cột cho tất cả các sheet (cũ & mới)
  syncSheetHeaders(ss);

  // Insert default config data if CauHinh is freshly created
  var cauHinhSheet = ss.getSheetByName(SHEET_NAMES.CAU_HINH);
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
  var tonKhoSheet = ss.getSheetByName(SHEET_NAMES.TON_KHO);
  if (tonKhoSheet && tonKhoSheet.getLastRow() <= 1) {
    rebuildTonKhoSheet(ss);
  }

  // Khởi tạo Báo cáo ngày nếu vừa được tạo mới hoặc chưa có dữ liệu
  var reportSheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_NGAY);
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
  var salesReportSheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_DOANH_SO);
  if (salesReportSheet && salesReportSheet.getLastRow() <= 1) {
    salesReportSheet.getRange(3, 1).setValue("Từ ngày:").setFontWeight("bold");
    var now = new Date();
    var firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
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
  var defaultSheet =
    ss.getSheetByName("Sheet1") ||
    ss.getSheetByName("Trang tính1") ||
    ss.getSheetByName("Sheet 1");
  if (defaultSheet && defaultSheet.getLastRow() === 0) {
    try {
      ss.deleteSheet(defaultSheet);
    } catch (e) {}
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
  var sheetOrder = SHEET_ORDER;

  var updatedSheets = [];

  sheetOrder.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    var targetHeaders = SHEET_HEADERS[name];
    if (!sheet || !targetHeaders || targetHeaders.length === 0) return;

    var maxCols = sheet.getMaxColumns();
    var lastCol = sheet.getLastColumn();
    var currentHeaders = [];

    if (lastCol > 0) {
      currentHeaders = sheet
        .getRange(1, 1, 1, lastCol)
        .getValues()[0]
        .map(function (h) {
          return String(h).trim();
        });
    }

    var needsUpdate = false;

    // Nếu số cột hiện tại ít hơn số tiêu đề định nghĩa
    if (maxCols < targetHeaders.length) {
      sheet.insertColumnsAfter(maxCols, targetHeaders.length - maxCols);
      needsUpdate = true;
    }

    // Kiểm tra từng tiêu đề cột xem có khớp hoặc thiếu không
    for (var colIdx = 0; colIdx < targetHeaders.length; colIdx++) {
      var expectedHeader = targetHeaders[colIdx];
      var actualHeader = currentHeaders[colIdx];

      if (String(actualHeader).trim() !== expectedHeader) {
        sheet.getRange(1, colIdx + 1).setValue(expectedHeader);
        needsUpdate = true;
      }
    }

    if (needsUpdate) {
      // Định dạng lại dòng tiêu đề cho đồng bộ
      var headerRange = sheet.getRange(1, 1, 1, targetHeaders.length);
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
  
  var targetSheets = [
    { name: SHEET_NAMES.DICH_VU, header: "Số điện thoại khách" },
    { name: SHEET_NAMES.THU_MUA, header: "Số điện thoại khách" },
    { name: SHEET_NAMES.BAO_HANH, header: "Số điện thoại" }
  ];
  
  var deletedCount = 0;
  
  targetSheets.forEach(function(target) {
    var sheet = ss.getSheetByName(target.name);
    if (!sheet) return;
    
    var lastCol = sheet.getLastColumn();
    if (lastCol <= 0) return;
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colIndex = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim().toLowerCase();
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
 * Đảm bảo các cấu hình mặc định (bao gồm cả các cấu hình mới thêm sau này)
 * tồn tại trong sheet Cấu hình.
 */
function ensureDefaultConfigsExist(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CAU_HINH);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  var keys = [];
  if (lastRow > 1) {
    keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(row) {
      return String(row[0]).trim().toLowerCase();
    });
  }

  DEFAULT_CONFIG.forEach(function(row) {
    var key = String(row[0]).trim().toLowerCase();
    if (keys.indexOf(key) === -1) {
      sheet.appendRow(row);
      var newRow = sheet.getLastRow();
      sheet.getRange(newRow, 1, 1, row.length)
        .setFontFamily("Times New Roman")
        .setFontSize(12)
        .setHorizontalAlignment("left");
    }
  });
}
