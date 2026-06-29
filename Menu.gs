/**
 * ============================================================
 * VanTran Mobile — Menu.gs
 * Custom menu + sidebar launcher functions
 * ============================================================
 */

/**
 * Tạo custom menu khi mở Spreadsheet
 */
function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Bà xã hơm")
    .addItem("Dịch vụ", "showSidebar")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Kiểm tra nhanh")
        .addItem("Phụ kiện sắp hết", "menuPhuKienSapHet")
        .addItem("DS trả góp quá hạn", "menuTraGopQuaHan"),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Báo cáo & Doanh số")
        .addItem("Cập nhật báo cáo ngày", "updateDailyReportFromSheet")
        .addItem("Cập nhật báo cáo doanh số", "updateSalesReportFromSheet")
        .addItem("Chốt doanh số tháng", "menuChotDoanhSo")
        .addItem("Xem báo cáo doanh số", "menuXemBaoCao"),
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Hệ thống")
        .addItem("Khởi tạo hệ thống", "setupSheets")
        .addItem("Đồng bộ & Chuẩn hóa toàn bộ hệ thống", "menuNormalizeSystem")
        .addItem("Di chuyển IMEI sang cột riêng", "migrateImeiToColumn")
        .addItem("Làm mới bộ nhớ đệm", "menuClearAllCaches"),
    )
    .addToUi();
}

// ======================== SIDEBAR LAUNCHERS ========================

/**
 * Hiển thị sidebar Quản lý cửa hàng (Bảng điều khiển chung)
 */
function showSidebar() {
  const html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = "donHang"; // Mặc định mở phần Đơn hàng
  const output = html
    .evaluate()
    .setWidth(900)
    .setHeight(950)
    .setTitle("BichLoan");
  SpreadsheetApp.getUi().showSidebar(output);
}

// ======================== MENU ACTION FUNCTIONS ========================

/**
 * Hiển thị danh sách trả góp quá hạn
 */
function menuTraGopQuaHan() {
  markOverdueInstallments();
  const quaHan = getTraGopQuaHan();

  if (quaHan.length === 0) {
    showAlert("Trả góp", "Không có khoản trả góp nào quá hạn!");
    return;
  }

  let msg = "Có " + quaHan.length + " kỳ trả góp quá hạn:\n\n";

  quaHan.forEach(function (item, index) {
    msg += index + 1 + ". HĐ " + item.MaTG + " - Kỳ " + item.KySo + "\n";
    msg += "   KH: " + item.TenKH + "\n";
    msg += "   Cần trả: " + formatCurrency(item.SoTienCanTra) + "đ\n";
    msg += "   Quá hạn: " + item.SoNgayQuaHan + " ngày\n\n";
  });

  showAlert("Trả góp quá hạn", msg);
}

/**
 * Hiển thị danh sách phụ kiện sắp hết
 */
function menuPhuKienSapHet() {
  const sapHet = getPhuKienSapHet();

  if (sapHet.length === 0) {
    showAlert("Tồn kho", "Không có phụ kiện nào sắp hết hàng!");
    return;
  }

  let msg = "Có " + sapHet.length + " phụ kiện sắp hết:\n\n";

  sapHet.forEach(function (item, index) {
    msg +=
      index +
      1 +
      ". " +
      item.MaPK +
      " - " +
      item.TenSP +
      " (" +
      item.ChiNhanh +
      ")\n";
    msg += "   Loại: " + item.LoaiPK + " | Tồn: " + item.SoLuongTon + "\n\n";
  });

  showAlert("Phụ kiện sắp hết", msg);
}

/**
 * Menu action to format all sheets with Times New Roman 12
 */
function menuFormatAllSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Tái xây dựng báo cáo Tồn kho và bảng màu trước
    rebuildTonKhoSheet(ss);
    applyConditionalFormatting(ss);

    // Đồng bộ font toàn bộ
    formatAllSheets();

    showAlert(
      "Định dạng",
      "Đã đồng bộ thành công Tồn kho chi nhánh, màu thương hiệu và Font chữ Times New Roman 12 cho toàn bộ hệ thống!",
    );
  } catch (e) {
    showAlert("Lỗi", "Không thể định dạng: " + e.message);
  }
}

/**
 * Menu action to synchronize, normalize, and format the entire system
 */
function menuNormalizeSystem() {
  try {
    // Xoá cache cấu trúc cột cũ để nạp lại cấu trúc mới nhất
    clearColumnEnumsCache();

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Tự động bổ sung các cấu hình mặc định còn thiếu
    ensureDefaultConfigsExist(ss);

    // 1. Chuẩn hóa cấu trúc cột và điền Ngày nhập/xuất cho Điện thoại (chạy ẩn, không show alert riêng)
    backfillDienThoaiDates(true);

    // 1.5. Xóa cột số điện thoại khách hàng thừa (nếu có) và đồng bộ lại headers
    migrateDeletePhoneColumns(ss);

    // 1.5.5. Di chuyển hoặc chèn cột "Trừ tiền thu máy" vào vị trí mới (giữa Đơn giá và Thành tiền)
    migrateTienThuMuaColumnPosition(ss);

    // 1.5.6. Chèn hoặc di chuyển cột "Trạng thái" vào vị trí mới ở sheet Thu mua
    migrateThuMuaStatusColumn(ss);

    // 1.6. Làm sạch và thiết lập lại toàn bộ Data Validation (dropdown) chính xác cho tất cả các sheet
    _setupDataValidations(ss);

    // 2. Tái xây dựng báo cáo Tồn kho chi nhánh
    rebuildTonKhoSheet(ss);

    // 3. Thiết lập định dạng màu sắc có điều kiện cho các dropdown
    applyConditionalFormatting(ss);

    // 4. Đồng bộ font chữ Times New Roman 12 cho toàn hệ thống
    formatAllSheets();

    showAlert(
      "Hoàn tất",
      "Đã đồng bộ và chuẩn hóa toàn bộ hệ thống thành công!\n\n" +
        "- Bổ sung và điền Ngày nhập/xuất Điện thoại cũ.\n" +
        "- Tái dựng bảng Tồn kho chi nhánh.\n" +
        "- Định dạng lại màu sắc dropdown.\n" +
        "- Áp dụng Font Times New Roman 12 cho tất cả các trang tính.",
    );
  } catch (e) {
    showAlert("Lỗi", "Không thể chuẩn hóa hệ thống: " + e.message);
  }
}

/**
 * Menu action to clear all caches manually
 */
function menuClearAllCaches() {
  try {
    clearAllCaches();
    showAlert(
      "Thành công",
      "Đã xóa toàn bộ bộ nhớ đệm (cache) và làm mới hệ thống!",
    );
  } catch (e) {
    showAlert("Lỗi", "Không thể xóa bộ nhớ đệm: " + e.message);
  }
}
