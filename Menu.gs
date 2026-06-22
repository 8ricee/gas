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
  var ui = SpreadsheetApp.getUi();

  ui.createMenu("VanTran Store")
    // Bán hàng
    .addItem("Tạo đơn hàng mới", "showSidebarDonHang")
    .addItem("Thêm khách hàng", "showSidebarKhachHang")
    .addSeparator()
    // Sản phẩm & Kho
    .addSubMenu(
      ui
        .createMenu("Sản phẩm & Kho")
        .addItem("Nhập kho sản phẩm", "showSidebarNhapKho")
        .addItem("Chuyển kho chi nhánh", "showSidebarChuyenKho")
        .addItem("Phụ kiện sắp hết", "menuPhuKienSapHet"),
    )
    .addSeparator()
    // Dịch vụ & Giao dịch
    .addSubMenu(
      ui
        .createMenu("Dịch vụ & Giao dịch")
        .addItem("Giao dịch dịch vụ", "showSidebarDichVu")
        .addItem("Đổi trả & Đổi quà", "showSidebarDoiTra")
        .addItem("Thu mua & Thu cũ đổi mới", "showSidebarThuMua"),
    )
    .addSeparator()
    // Trả góp
    .addSubMenu(
      ui
        .createMenu("Trả góp")
        .addItem("Ghi nhận thanh toán", "showSidebarThanhToan")
        .addItem("DS trả góp quá hạn", "menuTraGopQuaHan"),
    )
    .addSeparator()
    // Báo cáo & Doanh số
    .addItem("Cập nhật báo cáo ngày", "updateDailyReportFromSheet")
    .addItem("Cập nhật báo cáo doanh số", "updateSalesReportFromSheet")
    .addItem("Chốt doanh số tháng", "menuChotDoanhSo")
    .addItem("Xem báo cáo doanh số", "menuXemBaoCao")
    .addSeparator()
    // Hệ thống
    .addSubMenu(
      ui
        .createMenu("Hệ thống")
        .addItem("Khởi tạo hệ thống", "setupSheets")
        .addItem("Đồng bộ & Chuẩn hóa toàn bộ hệ thống", "menuNormalizeSystem"),
    )
    .addToUi();
}

// ======================== SIDEBAR LAUNCHERS ========================

/**
 * Hiển thị sidebar tạo đơn hàng
 */
function showSidebarDonHang() {
  var html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = "donHang";
  var output = html.evaluate().setTitle("Tạo đơn hàng mới");
  SpreadsheetApp.getUi().showSidebar(output);
}

/**
 * Hiển thị sidebar thêm khách hàng
 */
function showSidebarKhachHang() {
  var html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = "khachHang";
  var output = html.evaluate().setTitle("Thêm khách hàng");
  SpreadsheetApp.getUi().showSidebar(output);
}

/**
 * Hiển thị sidebar nhập kho
 */
function showSidebarNhapKho() {
  var html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = "nhapKho";
  var output = html.evaluate().setTitle("Nhập kho sản phẩm");
  SpreadsheetApp.getUi().showSidebar(output);
}

/**
 * Hiển thị sidebar dịch vụ
 */
function showSidebarDichVu() {
  var html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = "dichVu";
  var output = html.evaluate().setTitle("Giao dịch dịch vụ");
  SpreadsheetApp.getUi().showSidebar(output);
}

/**
 * Hiển thị sidebar ghi nhận thanh toán trả góp
 */
function showSidebarThanhToan() {
  var html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = "thanhToan";
  var output = html.evaluate().setTitle("Thanh toán trả góp");
  SpreadsheetApp.getUi().showSidebar(output);
}

/**
 * Hiển thị sidebar chuyển kho
 */
function showSidebarChuyenKho() {
  var html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = "chuyenKho";
  var output = html.evaluate().setTitle("Chuyển kho chi nhánh");
  SpreadsheetApp.getUi().showSidebar(output);
}

/**
 * Hiển thị sidebar đổi trả
 */
function showSidebarDoiTra() {
  var html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = "doiTra";
  var output = html.evaluate().setTitle("Đổi trả điện thoại");
  SpreadsheetApp.getUi().showSidebar(output);
}

/**
 * Hiển thị sidebar thu mua
 */
function showSidebarThuMua() {
  var html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = "thuMua";
  var output = html.evaluate().setTitle("Thu mua & Thu cũ đổi mới");
  SpreadsheetApp.getUi().showSidebar(output);
}

// ======================== MENU ACTION FUNCTIONS ========================

/**
 * Hiển thị danh sách trả góp quá hạn
 */
function menuTraGopQuaHan() {
  var quaHan = getTraGopQuaHan();

  if (quaHan.length === 0) {
    showAlert("Trả góp", "Không có khoản trả góp nào quá hạn!");
    return;
  }

  var msg = "Có " + quaHan.length + " kỳ trả góp quá hạn:\n\n";

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
  var sapHet = getPhuKienSapHet();

  if (sapHet.length === 0) {
    showAlert("Tồn kho", "Không có phụ kiện nào sắp hết hàng!");
    return;
  }

  var msg = "Có " + sapHet.length + " phụ kiện sắp hết:\n\n";

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
    var ss = SpreadsheetApp.getActiveSpreadsheet();

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
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Chuẩn hóa cấu trúc cột và điền Ngày nhập/xuất cho Điện thoại (chạy ẩn, không show alert riêng)
    backfillDienThoaiDates(true);

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
