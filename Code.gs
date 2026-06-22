/**
 * ============================================================
 * VanTran Mobile — Hệ thống Dịch vụ & Buôn bán Trả góp
 * Entry Point: Code.gs
 * ============================================================
 */

// ======================== CONSTANTS ========================

var SHEET_NAMES = {
  CAU_HINH: "Cấu hình",
  NHAN_VIEN: "Nhân viên",
  DIEN_THOAI: "Điện thoại",
  PHU_KIEN: "Phụ kiện",
  KHACH_HANG: "Khách hàng",
  DON_HANG: "Đơn hàng",
  DICH_VU: "Dịch vụ",
  TRA_GOP: "Trả góp",
  LICH_SU_TRA_GOP: "Lịch sử trả góp",
  DOANH_SO: "Doanh số",
  NHAP_KHO: "Nhập kho",
  DOI_TRA: "Đổi trả",
  THU_MUA: "Thu mua",
  BAO_CAO_NGAY: "Báo cáo ngày",
  BAO_CAO_DOANH_SO: "Báo cáo doanh số",
  TON_KHO: "Tồn kho",
  BAO_HANH: "Bảo hành",
};

// ======================== SHEET HEADERS (VIETNAMESE) ========================

var SHEET_HEADERS = {};

SHEET_HEADERS[SHEET_NAMES.CAU_HINH] = ["Tên cấu hình", "Giá trị", "Ghi chú"];

SHEET_HEADERS[SHEET_NAMES.NHAN_VIEN] = [
  "Mã nhân viên",
  "Họ tên",
  "Số điện thoại",
  "Email",
  "Vai trò",
  "Quyền xuất máy",
  "Ngày vào làm",
  "Trạng thái",
];

SHEET_HEADERS[SHEET_NAMES.DIEN_THOAI] = [
  "Mã điện thoại",
  "Tên sản phẩm",
  "Thương hiệu",
  "IMEI",
  "Màu sắc",
  "Dung lượng",
  "Tình trạng",
  "Giá nhập",
  "Giá bán",
  "Giá trả góp",
  "Trạng thái kho",
  "Ghi chú",
  "Chi nhánh",
  "Ngày nhập",
  "Ngày xuất",
];

SHEET_HEADERS[SHEET_NAMES.PHU_KIEN] = [
  "Mã phụ kiện",
  "Tên sản phẩm",
  "Loại phụ kiện",
  "Thương hiệu",
  "Giá nhập",
  "Giá bán",
  "Số lượng tồn",
  "Mô tả",
  "Trạng thái",
  "Chi nhánh",
];

SHEET_HEADERS[SHEET_NAMES.KHACH_HANG] = [
  "Mã khách hàng",
  "Họ tên",
  "CCCD",
  "Địa chỉ",
  "Ngày tạo",
  "Ghi chú",
];

SHEET_HEADERS[SHEET_NAMES.DON_HANG] = [
  "Mã đơn hàng",
  "Ngày bán",
  "Mã khách hàng",
  "Tên khách hàng",
  "Mã sản phẩm",
  "Tên sản phẩm",
  "Nguồn sản phẩm",
  "Thương hiệu",
  "Số lượng",
  "Đơn giá",
  "Thành tiền",
  "Hình thức bán",
  "Hình thức thanh toán",
  "Người bán",
  "Tên người bán",
  "Có quyền xuất máy",
  "Người hỗ trợ",
  "Tên người hỗ trợ",
  "Trạng thái",
  "Ghi chú",
  "Chi nhánh",
  "Mã quà tặng",
  "Tên quà tặng",
  "Có nhận quà",
  "Tiền giảm giá",
  "Tiền mặt",
  "Chuyển khoản",
];

SHEET_HEADERS[SHEET_NAMES.DICH_VU] = [
  "Mã dịch vụ",
  "Ngày giao dịch",
  "Loại dịch vụ",
  "Mã khách hàng",
  "Tên khách hàng",
  "Số điện thoại khách",
  "Số tiền giao dịch",
  "Phí dịch vụ",
  "Hình thức thanh toán",
  "Người thực hiện",
  "Tên người thực hiện",
  "Trạng thái",
  "Ghi chú",
  "Chi nhánh",
  "Tiền mặt",
  "Chuyển khoản",
];

SHEET_HEADERS[SHEET_NAMES.TRA_GOP] = [
  "Mã trả góp",
  "Mã đơn hàng",
  "Mã khách hàng",
  "Tên khách hàng",
  "Tổng tiền",
  "Trả trước",
  "Còn lại",
  "Số kỳ",
  "Tiền mỗi kỳ",
  "Ngày bắt đầu",
  "Ngày kết thúc",
  "Đã trả số kỳ",
  "Đã trả số tiền",
  "Loại trả góp",
  "Công ty tài chính",
  "Trạng thái",
  "Chi nhánh",
  "Tiền mặt",
  "Chuyển khoản",
];

SHEET_HEADERS[SHEET_NAMES.LICH_SU_TRA_GOP] = [
  "Mã lịch sử",
  "Mã trả góp",
  "Kỳ số",
  "Số tiền cần trả",
  "Số tiền đã trả",
  "Ngày cần trả",
  "Ngày thực trả",
  "Hình thức thanh toán",
  "Trạng thái",
  "Ghi chú",
  "Tiền mặt",
  "Chuyển khoản",
];

SHEET_HEADERS[SHEET_NAMES.DOANH_SO] = [
  "Tháng năm",
  "Mã nhân viên",
  "Tên nhân viên",
  "Vai trò",
  "Có quyền xuất máy",
  "Số máy bán Apple",
  "Số máy bán khác",
  "Số máy hỗ trợ Apple",
  "Số máy hỗ trợ khác",
  "Hoa hồng bán",
  "Hoa hồng hỗ trợ",
  "Tổng hoa hồng",
  "Doanh thu dịch vụ",
  "Thưởng",
  "Tổng thu nhập",
  "Trạng thái",
];

SHEET_HEADERS[SHEET_NAMES.NHAP_KHO] = [
  "Mã nhập kho",
  "Ngày nhập",
  "Nguồn nhập",
  "Mã sản phẩm",
  "Tên sản phẩm",
  "Số lượng",
  "Giá nhập",
  "Thành tiền",
  "Nhà cung cấp",
  "Ghi chú",
  "Chi nhánh",
];

SHEET_HEADERS[SHEET_NAMES.DOI_TRA] = [
  "Mã đổi trả",
  "Ngày đổi trả",
  "Mã đơn hàng",
  "Mã khách hàng",
  "Tên khách hàng",
  "Loại giao dịch",
  "Mã sản phẩm trả",
  "Tên sản phẩm trả",
  "IMEI trả",
  "Mã sản phẩm nhận",
  "Tên sản phẩm nhận",
  "IMEI nhận",
  "Tiền hoàn trả",
  "Phí đổi trả",
  "Hình thức thanh toán",
  "Chi nhánh",
  "Người thực hiện",
  "Trạng thái",
  "Ghi chú",
  "Tiền mặt",
  "Chuyển khoản",
];

SHEET_HEADERS[SHEET_NAMES.THU_MUA] = [
  "Mã thu mua",
  "Ngày thu mua",
  "Mã khách hàng",
  "Tên khách hàng",
  "Số điện thoại khách",
  "Tên sản phẩm thu",
  "Thương hiệu thu",
  "IMEI thu",
  "Màu sắc thu",
  "Dung lượng thu",
  "Tình trạng thu",
  "Giá thu mua",
  "Loại giao dịch",
  "Mã đơn hàng mới",
  "Tiền hỗ trợ",
  "Tổng tiền trả khách",
  "Hình thức thanh toán",
  "Chi nhánh",
  "Người thực hiện",
  "Ghi chú",
  "Tiền mặt",
  "Chuyển khoản",
];

SHEET_HEADERS[SHEET_NAMES.BAO_HANH] = [
  "Mã bảo hành",
  "Ngày nhận",
  "Mã khách hàng",
  "Tên khách hàng",
  "Số điện thoại",
  "Tên sản phẩm",
  "Tình trạng lỗi",
  "Loại dịch vụ",
  "Phí sửa chữa",
  "Hình thức thanh toán",
  "Người tiếp nhận",
  "Người sửa",
  "Trạng thái",
  "Ghi chú",
  "Chi nhánh",
  "Tiền mặt",
  "Chuyển khoản",
];

// ======================== DEFAULT CONFIG ========================

var DEFAULT_CONFIG = [
  ["Tên cửa hàng", "VanTran Mobile", "Tên hiển thị cửa hàng"],
  ["Địa chỉ", "", "Địa chỉ cửa hàng"],
  ["Số điện thoại", "", "SĐT liên hệ cửa hàng"],
  ["Tháng chốt DS", "", "Tháng đang chốt doanh số (MM/YYYY)"],
  ["HH Bán máy - Apple", "100000", "Hoa hồng người bán / sản phẩm Apple (VNĐ)"],
  [
    "HH Hỗ trợ - Apple",
    "50000",
    "Hoa hồng người hỗ trợ / sản phẩm Apple (VNĐ)",
  ],
  ["HH Bán máy - Khác", "50000", "Hoa hồng người bán / SP khác (VNĐ)"],
  ["HH Hỗ trợ - Khác", "25000", "Hoa hồng người hỗ trợ / SP khác (VNĐ)"],
  ["Phí CK hộ", "0", "Phí dịch vụ chuyển khoản hộ (VNĐ, cố định)"],
  ["Phí Rút tiền", "0", "Phí dịch vụ rút tiền mặt (VNĐ, cố định)"],
  ["Phí Nạp thẻ", "0", "Phí dịch vụ nạp thẻ (VNĐ, MIỄN PHÍ)"],
  [
    "Danh sách chi nhánh",
    "Chi nhánh 1, Chi nhánh 2, Chi nhánh 3",
    "Danh sách các chi nhánh, phân cách bằng dấu phẩy",
  ],
  [
    "Danh sách công ty tài chính",
    "FE Credit, Home Credit, HD Saison, MIRAES Asset",
    "Danh sách các công ty tài chính đối tác, phân cách bằng dấu phẩy",
  ],
  [
    "Danh sách thương hiệu",
    "Apple, Samsung, Xiaomi, OPPO, Vivo, Realme, Khác",
    "Danh sách các thương hiệu điện thoại, phân cách bằng dấu phẩy",
  ],
];

// ======================== SETUP ========================

/**
 * Khởi tạo toàn bộ hệ thống: tạo 13 sheet với header + dữ liệu cấu hình mặc định
 */
function setupSheets() {
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

  var sheetOrder = [
    SHEET_NAMES.CAU_HINH,
    SHEET_NAMES.NHAN_VIEN,
    SHEET_NAMES.DIEN_THOAI,
    SHEET_NAMES.PHU_KIEN,
    SHEET_NAMES.TON_KHO,
    SHEET_NAMES.KHACH_HANG,
    SHEET_NAMES.DON_HANG,
    SHEET_NAMES.DICH_VU,
    SHEET_NAMES.TRA_GOP,
    SHEET_NAMES.LICH_SU_TRA_GOP,
    SHEET_NAMES.DOANH_SO,
    SHEET_NAMES.NHAP_KHO,
    SHEET_NAMES.DOI_TRA,
    SHEET_NAMES.THU_MUA,
    SHEET_NAMES.BAO_CAO_NGAY,
    SHEET_NAMES.BAO_CAO_DOANH_SO,
    SHEET_NAMES.BAO_HANH,
  ];

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
  var sheetOrder = [
    SHEET_NAMES.CAU_HINH,
    SHEET_NAMES.NHAN_VIEN,
    SHEET_NAMES.DIEN_THOAI,
    SHEET_NAMES.PHU_KIEN,
    SHEET_NAMES.TON_KHO,
    SHEET_NAMES.KHACH_HANG,
    SHEET_NAMES.DON_HANG,
    SHEET_NAMES.DICH_VU,
    SHEET_NAMES.TRA_GOP,
    SHEET_NAMES.LICH_SU_TRA_GOP,
    SHEET_NAMES.DOANH_SO,
    SHEET_NAMES.NHAP_KHO,
    SHEET_NAMES.DOI_TRA,
    SHEET_NAMES.THU_MUA,
    SHEET_NAMES.BAO_CAO_NGAY,
    SHEET_NAMES.BAO_CAO_DOANH_SO,
    SHEET_NAMES.BAO_HANH,
  ];

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
 * Thiết lập Data Validation cho các sheet
 */
function _setupDataValidations(ss) {
  // Đọc danh sách chi nhánh động
  var chSheet = ss.getSheetByName(SHEET_NAMES.CAU_HINH);
  var branches = ["Chi nhánh 1", "Chi nhánh 2", "Chi nhánh 3"];
  if (chSheet) {
    var data = chSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === "danh sách chi nhánh") {
        var val = String(data[i][1]).trim();
        if (val) {
          branches = val.split(",").map(function (item) {
            return item.trim();
          });
        }
        break;
      }
    }
  }

  var chiNhanhRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(branches, true)
    .setAllowInvalid(false)
    .build();

  // Đọc danh sách thương hiệu động
  var brands = ["Apple", "Samsung", "Xiaomi", "OPPO", "Vivo", "Realme", "Khác"];
  if (chSheet) {
    var data = chSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === "danh sách thương hiệu") {
        var val = String(data[i][1]).trim();
        if (val) {
          brands = val.split(",").map(function (item) {
            return item.trim();
          });
        }
        break;
      }
    }
  }

  var thuongHieuRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(brands, true)
    .setAllowInvalid(false)
    .build();

  // NhanVien - VaiTro, TrangThai, ChiNhanh
  var nvSheet = ss.getSheetByName(SHEET_NAMES.NHAN_VIEN);
  if (nvSheet) {
    var vaiTroRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Bán hàng", "Kế toán", "Kỹ thuật"], true)
      .setAllowInvalid(false)
      .build();
    nvSheet.getRange("E2:E").setDataValidation(vaiTroRule);

    var quyenRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["✓", "✗"], true)
      .setAllowInvalid(false)
      .build();
    nvSheet.getRange("F2:F").setDataValidation(quyenRule);

    var ttNVRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Đang làm", "Nghỉ việc"], true)
      .setAllowInvalid(false)
      .build();
    nvSheet.getRange("H2:H").setDataValidation(ttNVRule);
  }

  // DienThoai - TinhTrang, TrangThaiKho, ChiNhanh
  var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  if (dtSheet) {
    // Xoá dropdown (Data Validation) của Tình trạng máy để người dùng nhập liệu tự do
    dtSheet.getRange("G2:G").clearDataValidations();

    var ttKhoRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["Còn hàng", "Đã bán", "Đang trả góp", "Đã trả lại"],
        true,
      )
      .setAllowInvalid(false)
      .build();
    dtSheet.getRange("K2:K").setDataValidation(ttKhoRule);
    dtSheet.getRange("M2:M").setDataValidation(chiNhanhRule);
    dtSheet.getRange("C2:C").setDataValidation(thuongHieuRule);
  }

  // PhuKien - TrangThai, ChiNhanh
  var pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
  if (pkSheet) {
    var loaiPKRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["Sạc", "Ốp lưng", "Tai nghe", "Cường lực", "Cáp", "Khác"],
        true,
      )
      .setAllowInvalid(false)
      .build();
    pkSheet.getRange("C2:C").setDataValidation(loaiPKRule);

    var ttPKRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Đang bán", "Ngừng bán"], true)
      .setAllowInvalid(false)
      .build();
    pkSheet.getRange("I2:I").setDataValidation(ttPKRule);
    pkSheet.getRange("J2:J").setDataValidation(chiNhanhRule);
  }

  // DonHang - HinhThucBan, HinhThucThanhToan, TrangThai, ChiNhanh, CoNhanQua
  var dhSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  if (dhSheet) {
    var htBanRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Bán thẳng", "Trả góp"], true)
      .setAllowInvalid(false)
      .build();
    dhSheet.getRange("L2:L").setDataValidation(htBanRule);

    var htTTRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"],
        true,
      )
      .setAllowInvalid(false)
      .build();
    dhSheet.getRange("M2:M").setDataValidation(htTTRule);

    var nguonSPRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Điện thoại", "Phụ kiện"], true)
      .setAllowInvalid(false)
      .build();
    dhSheet.getRange("G2:G").setDataValidation(nguonSPRule);

    var ttDHRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Hoàn thành", "Đang xử lý", "Huỷ", "Đổi trả"], true)
      .setAllowInvalid(false)
      .build();
    dhSheet.getRange("S2:S").setDataValidation(ttDHRule);
    dhSheet.getRange("U2:U").setDataValidation(chiNhanhRule);

    var coNhanQuaRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["✓", "✗"], true)
      .setAllowInvalid(false)
      .build();
    dhSheet.getRange("X2:X").setDataValidation(coNhanQuaRule);
  }

  // DichVu - LoaiDichVu, HinhThucThanhToan, TrangThai, ChiNhanh
  var dvSheet = ss.getSheetByName(SHEET_NAMES.DICH_VU);
  if (dvSheet) {
    var loaiDVRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["Chuyển khoản hộ", "Rút tiền mặt", "Nạp thẻ điện thoại"],
        true,
      )
      .setAllowInvalid(false)
      .build();
    dvSheet.getRange("C2:C").setDataValidation(loaiDVRule);

    var htTTDVRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"],
        true,
      )
      .setAllowInvalid(false)
      .build();
    dvSheet.getRange("I2:I").setDataValidation(htTTDVRule);

    var ttDVRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Hoàn thành", "Huỷ"], true)
      .setAllowInvalid(false)
      .build();
    dvSheet.getRange("L2:L").setDataValidation(ttDVRule);
    dvSheet.getRange("N2:N").setDataValidation(chiNhanhRule);
  }

  // TraGop - LoaiTraGop, TrangThai, ChiNhanh
  var tgSheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
  if (tgSheet) {
    var loaiTGRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Cửa hàng", "Công ty tài chính"], true)
      .setAllowInvalid(false)
      .build();
    tgSheet.getRange("N2:N").setDataValidation(loaiTGRule);

    var ttTGRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Đang trả", "Hoàn tất", "Quá hạn", "Đã huỷ"], true)
      .setAllowInvalid(false)
      .build();
    tgSheet.getRange("P2:P").setDataValidation(ttTGRule);
    tgSheet.getRange("Q2:Q").setDataValidation(chiNhanhRule);
  }

  // LichSuTraGop - HinhThucThanhToan, TrangThai
  var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
  if (lstgSheet) {
    var htTTLSTGRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"],
        true,
      )
      .setAllowInvalid(false)
      .build();
    lstgSheet.getRange("H2:H").setDataValidation(htTTLSTGRule);

    var ttLSTGRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Đã trả", "Chưa trả", "Quá hạn", "Đã huỷ"], true)
      .setAllowInvalid(false)
      .build();
    lstgSheet.getRange("I2:I").setDataValidation(ttLSTGRule);
  }

  // NhapKho - NguonNhap, ChiNhanh
  var nkSheet = ss.getSheetByName(SHEET_NAMES.NHAP_KHO);
  if (nkSheet) {
    var nnRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Điện thoại", "Phụ kiện"], true)
      .setAllowInvalid(false)
      .build();
    nkSheet.getRange("C2:C").setDataValidation(nnRule);
    nkSheet.getRange("K2:K").setDataValidation(chiNhanhRule);
  }

  // DoiTra - LoaiGiaoDich, HinhThucThanhToan, ChiNhanh, TrangThai
  var doiTraSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
  if (doiTraSheet) {
    var lgdRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Trả máy", "Đổi máy"], true)
      .setAllowInvalid(false)
      .build();
    doiTraSheet.getRange("F2:F").setDataValidation(lgdRule);

    var htttRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Tiền mặt", "Chuyển khoản", "Hỗn hợp"], true)
      .setAllowInvalid(false)
      .build();
    doiTraSheet.getRange("O2:O").setDataValidation(htttRule);
    doiTraSheet.getRange("P2:P").setDataValidation(chiNhanhRule);

    var ttDTRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Hoàn thành", "Huỷ"], true)
      .setAllowInvalid(false)
      .build();
    doiTraSheet.getRange("R2:R").setDataValidation(ttDTRule);
  }

  // ThuMua - TinhTrang, LoaiGiaoDich, HinhThucThanhToan, ChiNhanh
  var thuMuaSheet = ss.getSheetByName(SHEET_NAMES.THU_MUA);
  if (thuMuaSheet) {
    // Xoá dropdown (Data Validation) của Tình trạng máy để người dùng nhập liệu tự do
    thuMuaSheet.getRange("K2:K").clearDataValidations();

    var lgdRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Bán thẳng", "Thu cũ đổi mới"], true)
      .setAllowInvalid(false)
      .build();
    thuMuaSheet.getRange("M2:M").setDataValidation(lgdRule);

    var htttRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Trả góp", "Hỗn hợp"],
        true,
      )
      .setAllowInvalid(false)
      .build();
    thuMuaSheet.getRange("Q2:Q").setDataValidation(htttRule);
    thuMuaSheet.getRange("R2:R").setDataValidation(chiNhanhRule);
    thuMuaSheet.getRange("G2:G").setDataValidation(thuongHieuRule);
  }

  // DoanhSo - ChiNhanh
  var dsSheet = ss.getSheetByName(SHEET_NAMES.DOANH_SO);
  if (dsSheet) {
    dsSheet.getRange("Q2:Q").setDataValidation(chiNhanhRule);
  }

  // BaoCaoDoanhSo - Danh sách nhân viên và thiết lập bộ lọc
  var reportDS = ss.getSheetByName(SHEET_NAMES.BAO_CAO_DOANH_SO);
  if (reportDS) {
    var staffList = ["Tất cả"];
    if (nvSheet) {
      var lastRow = nvSheet.getLastRow();
      if (lastRow > 1) {
        var nvData = nvSheet.getRange(2, 1, lastRow - 1, 8).getValues();
        nvData.forEach(function (row) {
          var maNV = String(row[0]).trim();
          var tenNV = String(row[1]).trim();
          var trangThai = String(row[7]).trim();
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

  // BaoHanh - LoaiDichVu, HinhThucThanhToan, TrangThai, ChiNhanh
  var bhSheet = ss.getSheetByName(SHEET_NAMES.BAO_HANH);
  if (bhSheet) {
    var loaiBHRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Sửa chữa", "Bảo hành"], true)
      .setAllowInvalid(false)
      .build();
    bhSheet.getRange("H2:H").setDataValidation(loaiBHRule);

    var htTTBHRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(
        ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"],
        true,
      )
      .setAllowInvalid(false)
      .build();
    bhSheet.getRange("J2:J").setDataValidation(htTTBHRule);

    var ttBHRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Đang xử lý", "Hoàn thành", "Huỷ"], true)
      .setAllowInvalid(false)
      .build();
    bhSheet.getRange("M2:M").setDataValidation(ttBHRule);
    bhSheet.getRange("O2:O").setDataValidation(chiNhanhRule);
  }
}

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
      if (standardDataSheets.indexOf(sheetName) !== -1 && lastRow > 1) {
        var headers = headerRange.getValues()[0];
        for (var col = 1; col <= lastCol; col++) {
          var headerName = headers[col - 1];
          if (isFinancialHeader(headerName)) {
            sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat("#,##0");
          } else if (isDateTimeHeader(headerName)) {
            sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
          } else if (isDateHeader(headerName)) {
            sheet.getRange(2, col, lastRow - 1, 1).setNumberFormat("dd/MM/yyyy");
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
  var chSheet = ss.getSheetByName(SHEET_NAMES.CAU_HINH);
  var branches = ["Chi nhánh 1", "Chi nhánh 2", "Chi nhánh 3"];
  var brands = ["Apple", "Samsung", "Xiaomi", "OPPO", "Vivo", "Realme", "Khác"];

  if (chSheet) {
    var data = chSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][0]).trim().toLowerCase();
      var val = String(data[i][1]).trim();
      if (val) {
        if (key === "danh sách chi nhánh") {
          branches = val.split(",").map(function (item) {
            return item.trim();
          });
        } else if (key === "danh sách thương hiệu") {
          brands = val.split(",").map(function (item) {
            return item.trim();
          });
        }
      }
    }
  }

  var COLOR_MAP = {
    // 1. Trạng thái tốt, hoàn thành, đang làm, có, mới
    "Đang làm": { bg: "#e6f4ea", fg: "#137333" },
    "Còn hàng": { bg: "#e6f4ea", fg: "#137333" },
    "Hoàn thành": { bg: "#e6f4ea", fg: "#137333" },
    "Hoàn tất": { bg: "#e6f4ea", fg: "#137333" },
    "Đã trả": { bg: "#e6f4ea", fg: "#137333" },
    "Đang bán": { bg: "#e6f4ea", fg: "#137333" },
    "✓": { bg: "#e6f4ea", fg: "#137333" },
    "Mới 100%": { bg: "#e6f4ea", fg: "#137333" },
    "Tiền mặt": { bg: "#e6f4ea", fg: "#137333" },
    "Rút tiền mặt": { bg: "#e6f4ea", fg: "#137333" },

    // 2. Trạng thái nghỉ việc, huỷ, quá hạn, không, trả máy, đổi trả
    "Nghỉ việc": { bg: "#fce8e6", fg: "#c5221f" },
    "Huỷ": { bg: "#fce8e6", fg: "#c5221f" },
    "Quá hạn": { bg: "#fce8e6", fg: "#c5221f" },
    "Đã huỷ": { bg: "#fce8e6", fg: "#c5221f" },
    "Ngừng bán": { bg: "#fce8e6", fg: "#c5221f" },
    "✗": { bg: "#fce8e6", fg: "#c5221f" },
    "Trả máy": { bg: "#fce8e6", fg: "#c5221f" },
    "Đã trả lại": { bg: "#fce8e6", fg: "#c5221f" },
    "Đổi trả": { bg: "#fce8e6", fg: "#c5221f" },

    // 3. Trạng thái chờ, đang xử lý, đang trả, chưa trả, đổi máy, phụ kiện
    "Đang trả góp": { bg: "#fef7e0", fg: "#b06000" },
    "Đang xử lý": { bg: "#fef7e0", fg: "#b06000" },
    "Đang trả": { bg: "#fef7e0", fg: "#b06000" },
    "Chưa trả": { bg: "#fef7e0", fg: "#b06000" },
    "Thu cũ đổi mới": { bg: "#fef7e0", fg: "#b06000" },
    "Đổi máy": { bg: "#fef7e0", fg: "#b06000" },
    "Nạp thẻ điện thoại": { bg: "#fef7e0", fg: "#b06000" },
    "Cường lực": { bg: "#fef7e0", fg: "#b06000" },
    "Quẹt thẻ (POS)": { bg: "#fef7e0", fg: "#b06000" },
    "Kỹ thuật": { bg: "#fef7e0", fg: "#b06000" },

    // 4. Trạng thái khác, đã qua sử dụng, đã bán
    "Đã bán": { bg: "#cbd5e1", fg: "#334155" },       // Slate 300 / 700 (Thay cho màu mặc định #f1f3f4 / #5f6368)
    "Đã qua sử dụng": { bg: "#e2e8f0", fg: "#334155" }, // Slate 200 / 700
    "Khác": { bg: "#e2e8f0", fg: "#475569" },           // Slate 200 / 600

    // 5. Chi nhánh, vai trò, hình thức, nguồn
    "Bán hàng": { bg: "#e8f0fe", fg: "#1a73e8" },
    "Chuyển khoản": { bg: "#e8f0fe", fg: "#1a73e8" },
    "Chuyển khoản hộ": { bg: "#e8f0fe", fg: "#1a73e8" },
    "Sạc": { bg: "#e8f0fe", fg: "#1a73e8" },
    "Like New": { bg: "#e8f0fe", fg: "#1a73e8" },
    "Điện thoại": { bg: "#e8f0fe", fg: "#1a73e8" },

    "Kế toán": { bg: "#f3e8fd", fg: "#a142f4" },
    "Tai nghe": { bg: "#f3e8fd", fg: "#a142f4" },
    "Phụ kiện": { bg: "#f3e8fd", fg: "#a142f4" },
    "Công ty tài chính": { bg: "#f3e8fd", fg: "#a142f4" },

    "Ốp lưng": { bg: "#e0f7fa", fg: "#006064" },
    "Cáp": { bg: "#fce4ec", fg: "#880e4f" },
    "Cửa hàng": { bg: "#fef7e0", fg: "#b06000" },

    // Bổ sung các hình thức thanh toán & loại dịch vụ
    "Bán thẳng": { bg: "#e6f4ea", fg: "#137333" },
    "Trả góp": { bg: "#f3e8fd", fg: "#a142f4" },
    "Hỗn hợp": { bg: "#e0f7fa", fg: "#006064" },
    "Sửa chữa": { bg: "#e0f7fa", fg: "#006064" },
    "Bảo hành": { bg: "#fce4ec", fg: "#880e4f" }
  };

  // Dynamic colors for branches
  var BRANCH_COLORS = [
    { bg: "#e8f0fe", fg: "#1a73e8" }, // Blue
    { bg: "#fce8e6", fg: "#c5221f" }, // Red
    { bg: "#fef7e0", fg: "#b06000" }, // Yellow
    { bg: "#e6f4ea", fg: "#137333" }, // Green
    { bg: "#f3e8fd", fg: "#a142f4" }, // Purple
  ];
  branches.forEach(function (branch, index) {
    var colorIdx = index % BRANCH_COLORS.length;
    COLOR_MAP[branch] = BRANCH_COLORS[colorIdx];
  });

  // Dynamic colors for brands
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
      COLOR_MAP[brand] = { bg: "#cbd5e1", fg: "#1e293b" }; // Slate 300 / 900 cho Apple (Không dùng màu mặc định f1f3f4)
    } else if (lowerBrand === "khác") {
      COLOR_MAP[brand] = { bg: "#e2e8f0", fg: "#475569" }; // Slate 200 / 600 cho Khác
    } else {
      var colorIdx = index % BRAND_COLORS.length;
      COLOR_MAP[brand] = BRAND_COLORS[colorIdx];
    }
  });

  function _applyRulesForSheet(sheet, mappings) {
    if (!sheet) return;
    var rules = [];
    mappings.forEach(function (map) {
      var range = sheet.getRange(map.range);
      map.values.forEach(function (val) {
        var color = COLOR_MAP[val];
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

  // 1. Nhân viên
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.NHAN_VIEN), [
    { range: "E2:E", values: ["Bán hàng", "Kế toán", "Kỹ thuật"] },
    { range: "F2:F", values: ["✓", "✗"] },
    { range: "H2:H", values: ["Đang làm", "Nghỉ việc"] },
    { range: "I2:I", values: branches },
  ]);

  // 2. Điện thoại
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.DIEN_THOAI), [
    { range: "C2:C", values: brands },
    { range: "G2:G", values: ["Mới 100%", "Like New", "__NOT_EMPTY__"] },
    {
      range: "K2:K",
      values: ["Còn hàng", "Đã bán", "Đang trả góp", "Đã trả lại"],
    },
    { range: "M2:M", values: branches },
  ]);

  // 3. Phụ kiện
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.PHU_KIEN), [
    {
      range: "C2:C",
      values: ["Sạc", "Ốp lưng", "Tai nghe", "Cường lực", "Cáp", "Khác"],
    },
    { range: "D2:D", values: brands },
    { range: "I2:I", values: ["Đang bán", "Ngừng bán"] },
    { range: "J2:J", values: branches },
  ]);

  // 4. Đơn hàng
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.DON_HANG), [
    { range: "G2:G", values: ["Điện thoại", "Phụ kiện"] },
    { range: "L2:L", values: ["Bán thẳng", "Trả góp"] },
    { range: "M2:M", values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"] },
    { range: "S2:S", values: ["Hoàn thành", "Đang xử lý", "Huỷ", "Đổi trả"] },
    { range: "U2:U", values: branches },
    { range: "X2:X", values: ["✓", "✗"] },
  ]);

  // 5. Dịch vụ
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.DICH_VU), [
    {
      range: "C2:C",
      values: ["Chuyển khoản hộ", "Rút tiền mặt", "Nạp thẻ điện thoại"],
    },
    { range: "I2:I", values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"] },
    { range: "L2:L", values: ["Hoàn thành", "Huỷ"] },
    { range: "N2:N", values: branches },
  ]);

  // 6. Trả góp
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.TRA_GOP), [
    { range: "N2:N", values: ["Cửa hàng", "Công ty tài chính"] },
    { range: "P2:P", values: ["Đang trả", "Hoàn tất", "Quá hạn", "Đã huỷ"] },
    { range: "Q2:Q", values: branches },
  ]);

  // 7. Lịch sử trả góp
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP), [
    { range: "H2:H", values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"] },
    { range: "I2:I", values: ["Đã trả", "Chưa trả", "Quá hạn", "Đã huỷ"] },
  ]);

  // 8. Nhập kho
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.NHAP_KHO), [
    { range: "C2:C", values: ["Điện thoại", "Phụ kiện"] },
    { range: "K2:K", values: branches },
  ]);

  // 9. Đổi trả
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.DOI_TRA), [
    { range: "F2:F", values: ["Trả máy", "Đổi máy"] },
    { range: "O2:O", values: ["Tiền mặt", "Chuyển khoản", "Hỗn hợp"] },
    { range: "P2:P", values: branches },
    { range: "R2:R", values: ["Hoàn thành", "Huỷ"] },
  ]);

  // 10. Thu mua
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.THU_MUA), [
    { range: "G2:G", values: brands },
    { range: "K2:K", values: ["Mới 100%", "Like New", "__NOT_EMPTY__"] },
    { range: "M2:M", values: ["Bán thẳng", "Thu cũ đổi mới"] },
    { range: "Q2:Q", values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Trả góp", "Hỗn hợp"] },
    { range: "R2:R", values: branches },
  ]);

  // 11. Doanh số
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.DOANH_SO), [
    { range: "Q2:Q", values: branches },
  ]);

  // 12. Bảo hành
  _applyRulesForSheet(ss.getSheetByName(SHEET_NAMES.BAO_HANH), [
    { range: "H2:H", values: ["Sửa chữa", "Bảo hành"] },
    {
      range: "J2:J",
      values: ["Tiền mặt", "Chuyển khoản", "Quẹt thẻ (POS)", "Hỗn hợp"],
    },
    { range: "M2:M", values: ["Đang xử lý", "Hoàn thành", "Huỷ"] },
    { range: "O2:O", values: branches },
  ]);
}

/**
 * onOpen trigger — tạo custom menu khi mở Spreadsheet
 */
function onOpen() {
  createCustomMenu();
}

/**
 * Web App entry point: doGet
 * Serves the HTML interface (Sidebar.html) when accessed as a Web App
 */
function doGet(e) {
  var html = HtmlService.createTemplateFromFile("Sidebar");
  return html
    .evaluate()
    .setTitle("VanTran Mobile — Hệ thống Dịch vụ & Buôn bán Trả góp")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * onEdit trigger — auto-calculate khi chỉnh sửa
 */
function onEdit(e) {
  if (!e) return;

  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();

  // Skip header row
  if (row <= 1) return;

  try {
    range.setFontFamily("Times New Roman");
    range.setFontSize(12);
  } catch (ex) {}

  try {
    clearSheetCache(sheetName);
    invalidateDropdownCache(sheetName);
  } catch (ex) {}

  try {
    switch (sheetName) {
      case SHEET_NAMES.CAU_HINH:
        _setupDataValidations(e.source);
        break;
      case SHEET_NAMES.DON_HANG:
        _onEditDonHang(sheet, row, col, e);
        break;
      case SHEET_NAMES.NHAP_KHO:
        _onEditNhapKho(sheet, row, col, e);
        break;
      case SHEET_NAMES.BAO_CAO_NGAY:
        _onEditBaoCaoNgay(sheet, row, col, e);
        break;
      case SHEET_NAMES.BAO_CAO_DOANH_SO:
        _onEditBaoCaoDoanhSo(sheet, row, col, e);
        break;
      case SHEET_NAMES.LICH_SU_TRA_GOP:
        _onEditLichSuTraGop(sheet, row, col, e);
        break;
      case SHEET_NAMES.TRA_GOP:
        _onEditTraGop(sheet, row, col, e);
        break;
    }
  } catch (err) {
    Logger.log("onEdit error: " + err.message);
  }
}

/**
 * Xử lý sự kiện chỉnh sửa trên sheet Báo cáo ngày
 */
function _onEditBaoCaoNgay(sheet, row, col, e) {
  if (row === 3 && col === 2) {
    // Ô B3
    updateDailyReportFromSheet();
  }
}

/**
 * Xử lý sự kiện chỉnh sửa trên sheet Báo cáo doanh số
 */
function _onEditBaoCaoDoanhSo(sheet, row, col, e) {
  if (row === 3 && (col === 2 || col === 4 || col === 6 || col === 8)) {
    updateSalesReportFromSheet();
  }
}

/**
 * Tự động cập nhật khi sửa trực tiếp trên sheet Lịch sử trả góp
 */
function _onEditLichSuTraGop(sheet, row, col, e) {
  // Lắng nghe thay đổi Số tiền đã trả (col 5) hoặc Trạng thái (col 9)
  if (col === 5 || col === 9) {
    var maTG = sheet.getRange(row, 2).getValue();
    if (!maTG) return;

    var status = String(sheet.getRange(row, 9).getValue()).trim();

    // Nếu sửa trạng thái thành Đã trả, tự động điền các trường mặc định nếu trống
    if (col === 9 && status === "Đã trả") {
      var ngayThucTraRange = sheet.getRange(row, 7);
      if (!ngayThucTraRange.getValue()) {
        ngayThucTraRange.setValue(new Date());
      }
      var htttRange = sheet.getRange(row, 8);
      if (!htttRange.getValue()) {
        htttRange.setValue("Tiền mặt");
      }
      var soTienCanTra = Number(sheet.getRange(row, 4).getValue()) || 0;
      var soTienDaTraRange = sheet.getRange(row, 5);
      if (!soTienDaTraRange.getValue()) {
        soTienDaTraRange.setValue(soTienCanTra);
      }
    }

    // Nếu sửa trạng thái thành Chưa trả hoặc Đã huỷ, xoá thông tin thanh toán
    if (col === 9 && (status === "Chưa trả" || status === "Đã huỷ")) {
      sheet.getRange(row, 5).setValue(""); // Số tiền đã trả
      sheet.getRange(row, 7).setValue(""); // Ngày thực trả
      sheet.getRange(row, 8).setValue(""); // Hình thức thanh toán
    }

    // Cập nhật lại tổng trong TraGop và trạng thái máy
    _capNhatTongTraGop(maTG);
  }
}

/**
 * Tự động cập nhật khi sửa trực tiếp trên sheet Trả góp
 */
function _onEditTraGop(sheet, row, col, e) {
  // Lắng nghe thay đổi Trạng thái hợp đồng (col 16)
  if (col === 16) {
    var statusVal = String(e.value).trim();
    if (
      statusVal === "Đã huỷ" ||
      statusVal.toLowerCase() === "huy" ||
      statusVal.toLowerCase() === "huỷ"
    ) {
      var maTG = sheet.getRange(row, 1).getValue();
      var maDH = sheet.getRange(row, 2).getValue();

      if (maTG && maDH) {
        // 1. Hoàn trả kho máy sang "Còn hàng" nếu là điện thoại
        var maSP = lookupValue(SHEET_NAMES.DON_HANG, 1, maDH, 5);
        var nguonSP = lookupValue(SHEET_NAMES.DON_HANG, 1, maDH, 7);
        if (nguonSP === "Điện thoại" && maSP) {
          var ghiChuDH = lookupValue(SHEET_NAMES.DON_HANG, 1, maDH, 20) || "";
          var imeiMatch = ghiChuDH.match(/\[IMEI:\s*([^\s\]]+)\]/);
          var imei = imeiMatch ? imeiMatch[1] : "";
          updateTrangThaiKhoDT(imei || maSP, "Còn hàng");
        }

        // 2. Chuyển các kỳ chưa trả của HĐ này sang "Đã huỷ"
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
        if (lstgSheet) {
          var lastRow = lstgSheet.getLastRow();
          if (lastRow > 1) {
            var allLichSu = lstgSheet
              .getRange(2, 2, lastRow - 1, 8)
              .getValues();
            for (var i = 0; i < allLichSu.length; i++) {
              if (String(allLichSu[i][0]) === maTG) {
                var status = String(allLichSu[i][7]);
                if (status !== "Đã trả") {
                  lstgSheet.getRange(i + 2, 9).setValue("Đã huỷ");
                }
              }
            }
          }
        }

        showToast(
          "Hợp đồng " +
            maTG +
            " đã được huỷ. Đã hoàn kho sản phẩm và huỷ lịch sử trả góp!",
        );
      }
    }
  }
}

/**
 * Auto-calculate cho sheet DonHang (Đơn hàng)
 */
function _onEditDonHang(sheet, row, col, e) {
  // Auto tính ThanhTien khi SoLuong(9) hoặc DonGia(10) hoặc TienGiamGia(25) thay đổi
  if (col === 9 || col === 10 || col === 25) {
    var soLuong = sheet.getRange(row, 9).getValue() || 0;
    var donGia = sheet.getRange(row, 10).getValue() || 0;
    var giamGia = sheet.getRange(row, 25).getValue() || 0;
    sheet.getRange(row, 11).setValue(soLuong * donGia - giamGia);
  }

  // Auto lookup TenKH khi nhập MaKH (col 3)
  if (col === 3) {
    var maKH = e.value;
    if (maKH) {
      var tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, 1, maKH, 2);
      sheet.getRange(row, 4).setValue(tenKH || "");
    }
  }

  // Auto lookup TenSP, NguonSP, ThuongHieu khi nhập MaSP (col 5)
  if (col === 5) {
    var maSP = e.value;
    if (maSP) {
      // Thử tìm trong Điện thoại trước
      var tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP, 2);
      if (tenDT) {
        sheet.getRange(row, 6).setValue(tenDT);
        sheet.getRange(row, 7).setValue("Điện thoại");
        var thuongHieu = lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP, 3);
        sheet.getRange(row, 8).setValue(thuongHieu || "");
        var giaBan = lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP, 9);
        sheet.getRange(row, 10).setValue(giaBan || 0);
        sheet.getRange(row, 9).setValue(1); // Điện thoại luôn SL = 1
      } else {
        // Thử tìm trong Phụ kiện
        var tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, 1, maSP, 2);
        if (tenPK) {
          sheet.getRange(row, 6).setValue(tenPK);
          sheet.getRange(row, 7).setValue("Phụ kiện");
          var thuongHieuPK = lookupValue(SHEET_NAMES.PHU_KIEN, 1, maSP, 4);
          sheet.getRange(row, 8).setValue(thuongHieuPK || "");
          var giaBanPK = lookupValue(SHEET_NAMES.PHU_KIEN, 1, maSP, 6);
          sheet.getRange(row, 10).setValue(giaBanPK || 0);
        }
      }
    }
  }

  // Auto lookup TenNguoiBan + QuyenXuatMay khi nhập NguoiBan (col 14)
  if (col === 14) {
    var maNVBan = e.value;
    if (maNVBan) {
      var tenNVBan = lookupValue(SHEET_NAMES.NHAN_VIEN, 1, maNVBan, 2);
      sheet.getRange(row, 15).setValue(tenNVBan || "");
      var quyenXM = lookupValue(SHEET_NAMES.NHAN_VIEN, 1, maNVBan, 6);
      sheet.getRange(row, 16).setValue(quyenXM || "✗");
    }
  }

  // Auto lookup TenNguoiHoTro khi nhập NguoiHoTro (col 17)
  if (col === 17) {
    var maNVHT = e.value;
    if (maNVHT) {
      var tenNVHT = lookupValue(SHEET_NAMES.NHAN_VIEN, 1, maNVHT, 2);
      sheet.getRange(row, 18).setValue(tenNVHT || "");
    }
  }

  // Lắng nghe thay đổi trạng thái đơn hàng (col 19) để tự động xử lý kho và trả góp khi đơn bị Huỷ trực tiếp trên sheet
  if (col === 19) {
    var statusVal = String(e.value).trim();
    if (
      statusVal === "Huỷ" ||
      statusVal.toLowerCase() === "huy" ||
      statusVal.toLowerCase() === "huỷ"
    ) {
      var maDH = sheet.getRange(row, 1).getValue();
      var maSP = sheet.getRange(row, 5).getValue();
      var nguonSP = sheet.getRange(row, 7).getValue();
      var hinhThucBan = sheet.getRange(row, 12).getValue();
      var soLuong = Number(sheet.getRange(row, 9).getValue()) || 1;
      var branch = sheet.getRange(row, 21).getValue();
      var coNhanQua = sheet.getRange(row, 24).getValue();
      var maQua = sheet.getRange(row, 22).getValue();

      // 1. Hoàn trả kho sản phẩm chính
      if (nguonSP === "Điện thoại") {
        var ghiChuDH = sheet.getRange(row, 20).getValue() || "";
        var imeiMatch = ghiChuDH.match(/\[IMEI:\s*([^\s\]]+)\]/);
        var imei = imeiMatch ? imeiMatch[1] : "";
        updateTrangThaiKhoDT(imei || maSP, "Còn hàng");
        // Huỷ hợp đồng trả góp nếu bán trả góp
        if (hinhThucBan === "Trả góp") {
          huyHopDongTraGop(maDH);
        }
      } else {
        updateTonKhoPhuKien(maSP, soLuong, "nhap", branch);
      }

      // 2. Hoàn trả kho quà tặng nếu có
      if (coNhanQua === "✓" && maQua) {
        var maQuaList = String(maQua).split(",");
        for (var i = 0; i < maQuaList.length; i++) {
          var code = maQuaList[i].trim();
          if (code) {
            updateTonKhoPhuKien(code, 1, "nhap", branch);
          }
        }
      }

      showToast(
        "Đơn hàng " +
          maDH +
          " đã được huỷ trực tiếp trên sheet. Đã đồng bộ kho hàng!",
      );
    }
  }
}

/**
 * Auto-calculate cho sheet NhapKho (Nhập kho)
 */
function _onEditNhapKho(sheet, row, col, e) {
  // Auto tính ThanhTien khi SoLuong(6) hoặc GiaNhap(7) thay đổi
  if (col === 6 || col === 7) {
    var soLuong = sheet.getRange(row, 6).getValue() || 0;
    var giaNhap = sheet.getRange(row, 7).getValue() || 0;
    sheet.getRange(row, 8).setValue(soLuong * giaNhap);
  }

  // Auto lookup TenSP khi nhập MaSP (col 4)
  if (col === 4) {
    var maSP = e.value;
    var nguonNhap = sheet.getRange(row, 3).getValue();
    if (maSP) {
      if (nguonNhap === "Điện thoại") {
        var tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP, 2);
        sheet.getRange(row, 5).setValue(tenDT || "");
      } else {
        var tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, 1, maSP, 2);
        sheet.getRange(row, 5).setValue(tenPK || "");
      }
    }
  }
}

/**
 * Tái cấu trúc báo cáo sheet Tồn kho dựa trên danh sách chi nhánh động
 * Xoay chiều (pivot) số lượng tồn theo chi nhánh và bổ sung cột Tổng tồn
 *
 * @param {Spreadsheet} [ss] - Active spreadsheet
 */
function rebuildTonKhoSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var tonKhoSheet = ss.getSheetByName(SHEET_NAMES.TON_KHO);
  if (!tonKhoSheet) return;

  tonKhoSheet.clear();
  tonKhoSheet.clearFormats();

  // Đọc danh sách chi nhánh động
  var branches = getBranchesList();
  var N = branches.length || 1;

  // Đọc danh sách thương hiệu động
  var brands = getBrandsList();
  var B = brands.length || 1;

  // --- 1. TỔNG SỐ MÁY ĐIỆN THOẠI CỦA MỖI CHI NHÁNH (Bảng 1) ---
  var totalSummaryColIdx = 2 + N; // Cột 'Thương hiệu' (1), N Chi nhánh, Cột 'Tổng cộng' (1)

  // Dòng 1: Tiêu đề trang tính
  var sheetTitleRange = tonKhoSheet.getRange(1, 1, 1, totalSummaryColIdx);
  sheetTitleRange.merge();
  sheetTitleRange.setValue("BÁO CÁO TỒN KHO HỆ THỐNG & TÌM KIẾM");
  sheetTitleRange
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  // Dòng 2: Ô tìm kiếm nhanh
  var searchLabelCell = tonKhoSheet.getRange(2, 1);
  searchLabelCell
    .setValue("Tìm kiếm:")
    .setFontWeight("bold")
    .setHorizontalAlignment("right")
    .setBackground("#f1f3f4");

  var searchInputCell = tonKhoSheet.getRange(2, 2);
  searchInputCell
    .setBackground("#ffffff")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#cccccc",
      SpreadsheetApp.BorderStyle.SOLID,
    )
    .setHorizontalAlignment("left");

  var searchTipRange = tonKhoSheet.getRange(
    2,
    3,
    1,
    Math.max(1, totalSummaryColIdx - 2),
  );
  searchTipRange.merge();
  searchTipRange
    .setValue("Nhập B2 để tìm kiếm")
    .setFontStyle("italic")
    .setFontColor("#5f6368")
    .setBackground("#f1f3f4")
    .setHorizontalAlignment("left");

  // Dòng 3: Dòng spacer trống
  tonKhoSheet.setRowHeight(3, 15);

  // Dòng 4: Tiêu đề Bảng 1
  var summaryTitleRange = tonKhoSheet.getRange(4, 1, 1, totalSummaryColIdx);
  summaryTitleRange.merge();
  summaryTitleRange.setValue("TỔNG SỐ MÁY ĐIỆN THOẠI CỦA MỖI CHI NHÁNH");
  summaryTitleRange
    .setFontWeight("bold")
    .setFontSize(12)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  // Dòng 5: Headers Bảng 1
  var summaryHeaders = ["Thương hiệu"];
  branches.forEach(function (branch) {
    summaryHeaders.push(branch);
  });
  summaryHeaders.push("Tổng cộng");
  tonKhoSheet.getRange(5, 1, 1, totalSummaryColIdx).setValues([summaryHeaders]);

  // Dòng 6 đến 5 + B: Ghi danh sách thương hiệu & công thức
  var brandValues = [];
  brands.forEach(function (brand) {
    brandValues.push([brand]);
  });
  tonKhoSheet.getRange(6, 1, B, 1).setValues(brandValues);

  // Chèn công thức COUNTIFS cho các chi nhánh & SUM cho cột Tổng cộng
  for (var r = 6; r <= 5 + B; r++) {
    for (var c = 2; c <= 1 + N; c++) {
      var colLetter = columnToLetter(c);
      tonKhoSheet
        .getRange(r, c)
        .setFormula(
          "=COUNTIFS('Điện thoại'!$C:$C, $A" +
            r +
            ", 'Điện thoại'!$M:$M, " +
            colLetter +
            "$5, 'Điện thoại'!$K:$K, \"Còn hàng\")",
        );
    }
    var lastBranchColLetter = columnToLetter(1 + N);
    tonKhoSheet
      .getRange(r, 2 + N)
      .setFormula("=SUM(B" + r + ":" + lastBranchColLetter + r + ")");
  }

  // Dòng 6 + B: Dòng Tổng cộng của bảng 1
  var totalRowIdx = 6 + B;
  tonKhoSheet.getRange(totalRowIdx, 1).setValue("Tổng cộng");
  for (var c = 2; c <= 1 + N; c++) {
    var colLetter = columnToLetter(c);
    tonKhoSheet
      .getRange(totalRowIdx, c)
      .setFormula("=SUM(" + colLetter + "6:" + colLetter + (5 + B) + ")");
  }
  tonKhoSheet
    .getRange(totalRowIdx, 2 + N)
    .setFormula(
      "=SUM(B" + totalRowIdx + ":" + lastBranchColLetter + totalRowIdx + ")",
    );

  // Format Table 1
  var summaryHeaderRange = tonKhoSheet.getRange(5, 1, 1, totalSummaryColIdx);
  summaryHeaderRange
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  var summaryTotalRowRange = tonKhoSheet.getRange(
    totalRowIdx,
    1,
    1,
    totalSummaryColIdx,
  );
  summaryTotalRowRange
    .setBackground("#f1f3f4")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  // Dòng spacer trống là dòng 7 + B
  tonKhoSheet.setRowHeight(7 + B, 15);
  var startRowB2B3 = 8 + B;

  // --- 2. TỒN KHO ĐIỆN THOẠI THEO CHI NHÁNH (Bảng 2) ---
  var totalPhoneColIdx = 3 + N + 1; // 3 cột thông tin + N chi nhánh + 1 cột tổng
  var phoneTitleRange = tonKhoSheet.getRange(
    startRowB2B3,
    1,
    1,
    totalPhoneColIdx,
  );
  phoneTitleRange.merge();
  phoneTitleRange.setValue("TỒN KHO ĐIỆN THOẠI THEO CHI NHÁNH");
  phoneTitleRange
    .setFontWeight("bold")
    .setFontSize(13)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  // Công thức QUERY xoay chiều theo chi nhánh M ở dòng startRowB2B3 + 1 (có lọc theo ô Tìm kiếm B2)
  tonKhoSheet
    .getRange(startRowB2B3 + 1, 1)
    .setFormula(
      '=QUERY(\'Điện thoại\'!A:M, "select B, E, F, count(A) where A is not null and K = \'Còn hàng\' " & IF(ISBLANK($B$2), "", "and (lower(B) contains \'"&LOWER($B$2)&"\' or lower(E) contains \'"&LOWER($B$2)&"\' or lower(F) contains \'"&LOWER($B$2)&"\' or lower(C) contains \'"&LOWER($B$2)&"\')") & " group by B, E, F pivot M", 1)',
    );

  // Đầu cột Tổng tồn
  var phoneTotalHeaderCell = tonKhoSheet.getRange(
    startRowB2B3 + 1,
    totalPhoneColIdx,
  );
  phoneTotalHeaderCell
    .setValue("Tổng tồn")
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8");

  // Công thức tính Tổng tồn dùng MAP và OFFSET động từ dòng startRowB2B3 + 2
  var phoneTotalFormulaCell = tonKhoSheet.getRange(
    startRowB2B3 + 2,
    totalPhoneColIdx,
  );
  phoneTotalFormulaCell.setFormula(
    "=MAP(A" +
      (startRowB2B3 + 2) +
      ':A, LAMBDA(val, IF(val="", "", SUM(OFFSET(val, 0, 3, 1, ' +
      N +
      ")))))",
  );

  // --- 3. TỒN KHO PHỤ KIỆN THEO CHI NHÁNH (Bảng 3) ---
  var spacerColIdx = totalPhoneColIdx + 1; // Cột khoảng cách trống
  var startAccColIdx = spacerColIdx + 1; // Cột bắt đầu bảng Phụ kiện
  var totalAccColIdx = startAccColIdx + 4 + N; // 4 cột thông tin + N chi nhánh + 1 cột tổng

  var accTitleRange = tonKhoSheet.getRange(
    startRowB2B3,
    startAccColIdx,
    1,
    4 + N + 1,
  );
  accTitleRange.merge();
  accTitleRange.setValue("TỒN KHO PHỤ KIỆN THEO CHI NHÁNH");
  accTitleRange
    .setFontWeight("bold")
    .setFontSize(13)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  // Công thức QUERY xoay chiều theo chi nhánh J (có lọc theo ô Tìm kiếm B2)
  var startAccColLetter = columnToLetter(startAccColIdx);
  tonKhoSheet
    .getRange(startRowB2B3 + 1, startAccColIdx)
    .setFormula(
      '=QUERY(\'Phụ kiện\'!A:J, "select A, B, C, D, sum(G) where A is not null and I = \'Đang bán\' and G > 0 " & IF(ISBLANK($B$2), "", "and (lower(A) contains \'"&LOWER($B$2)&"\' or lower(B) contains \'"&LOWER($B$2)&"\' or lower(C) contains \'"&LOWER($B$2)&"\' or lower(D) contains \'"&LOWER($B$2)&"\')") & " group by A, B, C, D pivot J", 1)',
    );

  // Đầu cột Tổng tồn của Phụ kiện
  var accTotalHeaderCell = tonKhoSheet.getRange(
    startRowB2B3 + 1,
    totalAccColIdx,
  );
  accTotalHeaderCell
    .setValue("Tổng tồn")
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8");

  // Công thức tính Tổng tồn Phụ kiện
  var accTotalFormulaCell = tonKhoSheet.getRange(
    startRowB2B3 + 2,
    totalAccColIdx,
  );
  accTotalFormulaCell.setFormula(
    "=MAP(" +
      startAccColLetter +
      (startRowB2B3 + 2) +
      ":" +
      startAccColLetter +
      ', LAMBDA(val, IF(val="", "", SUM(OFFSET(val, 0, 4, 1, ' +
      N +
      ")))))",
  );

  // Đảm bảo số cột trong sheet đủ lớn
  var lastRow = tonKhoSheet.getMaxRows();
  var lastCol = tonKhoSheet.getMaxColumns();
  var neededCols = Math.max(totalAccColIdx, totalSummaryColIdx);
  if (lastCol < neededCols) {
    tonKhoSheet.insertColumnsAfter(lastCol, neededCols - lastCol);
    lastCol = neededCols;
  }

  // Định dạng toàn bộ bảng
  var fullRange = tonKhoSheet.getRange(1, 1, lastRow, lastCol);
  fullRange.setFontFamily("Times New Roman").setFontSize(12);

  // Tô màu nền cho header dòng startRowB2B3 + 1 của cả 2 bảng
  var phoneHeaderRange = tonKhoSheet.getRange(
    startRowB2B3 + 1,
    1,
    1,
    totalPhoneColIdx,
  );
  phoneHeaderRange
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  var accHeaderRange = tonKhoSheet.getRange(
    startRowB2B3 + 1,
    startAccColIdx,
    1,
    4 + N + 1,
  );
  accHeaderRange
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  // Điều chỉnh độ rộng cột
  for (var c = 1; c <= lastCol; c++) {
    if (c === spacerColIdx) {
      tonKhoSheet.setColumnWidth(c, 40);
    } else {
      tonKhoSheet.autoResizeColumn(c);
    }
  }
}

/**
 * Chuyển số cột thành chữ cái (ví dụ: 1 -> A, 27 -> AA)
 *
 * @param {number} column - Chỉ số cột (1-indexed)
 * @return {string}
 */
function columnToLetter(column) {
  var temp,
    letter = "";
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Lấy dữ liệu tồn kho để preload xuống Client-side cache
 *
 * @return {Object} Chứa mảng dienThoai, phuKien và phuKienUnique
 */
function getPreloadData() {
  return {
    dienThoai: getDienThoaiDropdown(""),
    phuKien: getPhuKienDropdown(""),
    phuKienUnique: getPhuKienUniqueList(),
  };
}
