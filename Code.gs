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

var SHEET_ORDER = [
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
  "IMEI 2",
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
  ["HH Bán máy - Khác", "50000", "Hoa hồng người bán / sản phẩm khác (VNĐ)"],
  ["HH Hỗ trợ - Khác", "25000", "Hoa hồng người hỗ trợ / sản phẩm khác (VNĐ)"],
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
  [
    "Lãi suất trả góp cửa hàng (%)",
    "2",
    "Lãi suất trả góp cửa hàng cố định hàng tháng (%)",
  ],
];

// ======================== CONDITIONAL FORMATTING COLORS ========================
var CONDITIONAL_COLOR_MAP = {
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
  "Đã bán": { bg: "#cbd5e1", fg: "#334155" },
  "Đã qua sử dụng": { bg: "#e2e8f0", fg: "#334155" },
  "Khác": { bg: "#e2e8f0", fg: "#475569" },

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

  "Bán thẳng": { bg: "#e6f4ea", fg: "#137333" },
  "Trả góp": { bg: "#f3e8fd", fg: "#a142f4" },
  "Hỗn hợp": { bg: "#e0f7fa", fg: "#006064" },
  "Sửa chữa": { bg: "#e0f7fa", fg: "#006064" },
  "Bảo hành": { bg: "#fce4ec", fg: "#880e4f" }
};

// ======================== TRIGGERS & WebApp Serving ========================

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
  initializeColumnEnums();
  // Lắng nghe thay đổi Số tiền đã trả hoặc Trạng thái
  if (col === COL_LSTG.SO_TIEN_DA_TRA || col === COL_LSTG.TRANG_THAI) {
    var maTG = sheet.getRange(row, COL_LSTG.MA_TG).getValue();
    if (!maTG) return;

    var status = String(sheet.getRange(row, COL_LSTG.TRANG_THAI).getValue()).trim();

    // Nếu sửa trạng thái thành Đã trả, tự động điền các trường mặc định nếu trống
    if (col === COL_LSTG.TRANG_THAI && status === "Đã trả") {
      var ngayThucTraRange = sheet.getRange(row, COL_LSTG.NGAY_THUC_TRA);
      if (!ngayThucTraRange.getValue()) {
        ngayThucTraRange.setValue(new Date());
      }
      var htttRange = sheet.getRange(row, COL_LSTG.HINH_THUC_TT);
      if (!htttRange.getValue()) {
        htttRange.setValue("Tiền mặt");
      }
      var soTienCanTra = Number(sheet.getRange(row, COL_LSTG.SO_TIEN_CAN_TRA).getValue()) || 0;
      var soTienDaTraRange = sheet.getRange(row, COL_LSTG.SO_TIEN_DA_TRA);
      if (!soTienDaTraRange.getValue()) {
        soTienDaTraRange.setValue(soTienCanTra);
      }
    }

    // Nếu sửa trạng thái thành Chưa trả hoặc Đã huỷ, xoá thông tin thanh toán
    if (col === COL_LSTG.TRANG_THAI && (status === "Chưa trả" || status === "Đã huỷ")) {
      sheet.getRange(row, COL_LSTG.SO_TIEN_DA_TRA).setValue(""); // Số tiền đã trả
      sheet.getRange(row, COL_LSTG.NGAY_THUC_TRA).setValue(""); // Ngày thực trả
      sheet.getRange(row, COL_LSTG.HINH_THUC_TT).setValue(""); // Hình thức thanh toán
    }

    // Cập nhật lại tổng trong TraGop và trạng thái máy
    _capNhatTongTraGop(maTG);
  }
}

/**
 * Tự động cập nhật khi sửa trực tiếp trên sheet Trả góp
 */
function _onEditTraGop(sheet, row, col, e) {
  initializeColumnEnums();
  // Lắng nghe thay đổi Trạng thái hợp đồng
  if (col === COL_TG.TRANG_THAI) {
    var statusVal = String(e.value).trim();
    if (
      statusVal === "Đã huỷ" ||
      statusVal.toLowerCase() === "huy" ||
      statusVal.toLowerCase() === "huỷ"
    ) {
      var maTG = sheet.getRange(row, COL_TG.MA_TG).getValue();
      var maDH = sheet.getRange(row, COL_TG.MA_DH).getValue();

      if (maTG && maDH) {
        // 1. Hoàn trả kho máy sang "Còn hàng" nếu là điện thoại
        var maSP = lookupValue(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH, COL_DH.MA_SP);
        var nguonSP = lookupValue(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH, COL_DH.NGUON_SP);
        if (nguonSP === "Điện thoại" && maSP) {
          var ghiChuDH = lookupValue(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH, COL_DH.GHI_CHU) || "";
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
              .getRange(2, COL_LSTG.MA_TG, lastRow - 1, COL_LSTG.TRANG_THAI - COL_LSTG.MA_TG + 1)
              .getValues();
            for (var i = 0; i < allLichSu.length; i++) {
              if (String(allLichSu[i][0]) === maTG) {
                var status = String(allLichSu[i][COL_LSTG.TRANG_THAI - COL_LSTG.MA_TG]);
                if (status !== "Đã trả") {
                  lstgSheet.getRange(i + 2, COL_LSTG.TRANG_THAI).setValue("Đã huỷ");
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
  initializeColumnEnums();

  // Auto tính ThanhTien khi SoLuong hoặc DonGia hoặc TienGiamGia thay đổi
  if (col === COL_DH.SO_LUONG || col === COL_DH.DON_GIA || col === COL_DH.TIEN_GIAM_GIA) {
    var soLuong = sheet.getRange(row, COL_DH.SO_LUONG).getValue() || 0;
    var donGia = sheet.getRange(row, COL_DH.DON_GIA).getValue() || 0;
    var giamGia = sheet.getRange(row, COL_DH.TIEN_GIAM_GIA).getValue() || 0;
    sheet.getRange(row, COL_DH.THANH_TIEN).setValue(soLuong * donGia - giamGia);
  }

  // Auto lookup TenKH khi nhập MaKH
  if (col === COL_DH.MA_KH) {
    var maKH = e.value;
    if (maKH) {
      var tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, COL_KH.MA_KH, maKH, COL_KH.HO_TEN);
      sheet.getRange(row, COL_DH.TEN_KH).setValue(tenKH || "");
    }
  }

  // Auto lookup TenSP, NguonSP, ThuongHieu khi nhập MaSP
  if (col === COL_DH.MA_SP) {
    var maSP = e.value;
    if (maSP) {
      // Thử tìm trong Điện thoại trước
      var tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.TEN_SP);
      if (tenDT) {
        sheet.getRange(row, COL_DH.TEN_SP).setValue(tenDT);
        sheet.getRange(row, COL_DH.NGUON_SP).setValue("Điện thoại");
        var thuongHieu = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.THUONG_HIEU);
        sheet.getRange(row, COL_DH.THUONG_HIEU).setValue(thuongHieu || "");
        var giaBan = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.GIA_BAN);
        sheet.getRange(row, COL_DH.DON_GIA).setValue(giaBan || 0);
        sheet.getRange(row, COL_DH.SO_LUONG).setValue(1); // Điện thoại luôn SL = 1
      } else {
        // Thử tìm trong Phụ kiện
        var tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.TEN_SP);
        if (tenPK) {
          sheet.getRange(row, COL_DH.TEN_SP).setValue(tenPK);
          sheet.getRange(row, COL_DH.NGUON_SP).setValue("Phụ kiện");
          var thuongHieuPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.THUONG_HIEU);
          sheet.getRange(row, COL_DH.THUONG_HIEU).setValue(thuongHieuPK || "");
          var giaBanPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.GIA_BAN);
          sheet.getRange(row, COL_DH.DON_GIA).setValue(giaBanPK || 0);
        }
      }
    }
  }

  // Auto lookup TenNguoiBan + QuyenXuatMay khi nhập NguoiBan (Sử dụng dynamic enum COL_NV)
  if (col === COL_DH.NGUOI_BAN) {
    var maNVBan = e.value;
    if (maNVBan) {
      var tenNVBan = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNVBan, COL_NV.HO_TEN);
      sheet.getRange(row, COL_DH.TEN_NGUOI_BAN).setValue(tenNVBan || "");
      var quyenXM = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNVBan, COL_NV.QUYEN_XUAT);
      sheet.getRange(row, COL_DH.QUYEN_XUAT).setValue(quyenXM || "✗");
    }
  }

  // Auto lookup TenNguoiHoTro khi nhập NguoiHoTro (Sử dụng dynamic enum COL_NV)
  if (col === COL_DH.NGUOI_HO_TRO) {
    var maNVHT = e.value;
    if (maNVHT) {
      var tenNVHT = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNVHT, COL_NV.HO_TEN);
      sheet.getRange(row, COL_DH.TEN_NGUOI_HO_TRO).setValue(tenNVHT || "");
    }
  }

  // Lắng nghe thay đổi trạng thái đơn hàng để tự động xử lý kho và trả góp khi đơn bị Huỷ trực tiếp trên sheet
  if (col === COL_DH.TRANG_THAI) {
    var statusVal = String(e.value).trim();
    if (
      statusVal === "Huỷ" ||
      statusVal.toLowerCase() === "huy" ||
      statusVal.toLowerCase() === "huỷ"
    ) {
      var maDH = sheet.getRange(row, COL_DH.MA_DH).getValue();
      var maSP = sheet.getRange(row, COL_DH.MA_SP).getValue();
      var nguonSP = sheet.getRange(row, COL_DH.NGUON_SP).getValue();
      var hinhThucBan = sheet.getRange(row, COL_DH.HINH_THUC_BAN).getValue();
      var soLuong = Number(sheet.getRange(row, COL_DH.SO_LUONG).getValue()) || 1;
      var branch = sheet.getRange(row, COL_DH.CHI_NHANH).getValue();
      var coNhanQua = sheet.getRange(row, COL_DH.CO_NHAN_QUA).getValue();
      var maQua = sheet.getRange(row, COL_DH.MA_QUA_TANG).getValue();

      // 1. Hoàn trả kho sản phẩm chính
      if (nguonSP === "Điện thoại") {
        var ghiChuDH = sheet.getRange(row, COL_DH.GHI_CHU).getValue() || "";
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
  initializeColumnEnums();

  // Auto tính ThanhTien khi SoLuong hoặc GiaNhap thay đổi
  if (col === COL_NK.SO_LUONG || col === COL_NK.GIA_NHAP) {
    var soLuong = sheet.getRange(row, COL_NK.SO_LUONG).getValue() || 0;
    var giaNhap = sheet.getRange(row, COL_NK.GIA_NHAP).getValue() || 0;
    sheet.getRange(row, COL_NK.THANH_TIEN).setValue(soLuong * giaNhap);
  }

  // Auto lookup TenSP khi nhập MaSP
  if (col === COL_NK.MA_SP) {
    var maSP = e.value;
    var nguonNhap = sheet.getRange(row, COL_NK.NGUON_NHAP).getValue();
    if (maSP) {
      if (nguonNhap === "Điện thoại") {
        var tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.TEN_SP);
        sheet.getRange(row, COL_NK.TEN_SP).setValue(tenDT || "");
      } else {
        var tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.TEN_SP);
        sheet.getRange(row, COL_NK.TEN_SP).setValue(tenPK || "");
      }
    }
  }
}

// ======================== TỒN KHO & TEMPLATE UTILITIES ========================

/**
 * Tái cấu trúc báo cáo sheet Tồn kho dựa trên danh sách chi nhánh động
 * Xoay chiều (pivot) số lượng tồn theo chi nhánh và bổ sung cột Tổng tồn
 *
 * @param {Spreadsheet} [ss] - Active spreadsheet
 */
function rebuildTonKhoSheet(ss) {
  initializeColumnEnums();
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  var colThuongHieu = columnToLetter(COL_DT.THUONG_HIEU);
  var colChiNhanh = columnToLetter(COL_DT.CHI_NHANH);
  var colTrangThai = columnToLetter(COL_DT.TRANG_THAI_KHO);
  var colTenSP = columnToLetter(COL_DT.TEN_SP);
  var colMauSac = columnToLetter(COL_DT.MAU_SAC);
  var colDungLuong = columnToLetter(COL_DT.DUNG_LUONG);
  var colMaDT = columnToLetter(COL_DT.MA_DT);
  var colImei = columnToLetter(COL_DT.IMEI);
  var colImei2 = columnToLetter(COL_DT.IMEI_2);

  var maxColLetter = columnToLetter(Math.max(
    COL_DT.MA_DT, COL_DT.TEN_SP, COL_DT.THUONG_HIEU, COL_DT.IMEI, COL_DT.IMEI_2,
    COL_DT.MAU_SAC, COL_DT.DUNG_LUONG, COL_DT.TRANG_THAI_KHO, COL_DT.CHI_NHANH
  ));

  var colMaPK = columnToLetter(COL_PK.MA_PK);
  var colTenPK = columnToLetter(COL_PK.TEN_SP);
  var colLoaiPK = columnToLetter(COL_PK.LOAI_PK);
  var colThuongHieuPK = columnToLetter(COL_PK.THUONG_HIEU);
  var colSoLuongTon = columnToLetter(COL_PK.SO_LUONG_TON);
  var colTrangThaiPK = columnToLetter(COL_PK.TRANG_THAI);
  var colChiNhanhPK = columnToLetter(COL_PK.CHI_NHANH);

  var maxColPKLetter = columnToLetter(Math.max(
    COL_PK.MA_PK, COL_PK.TEN_SP, COL_PK.LOAI_PK, COL_PK.THUONG_HIEU,
    COL_PK.SO_LUONG_TON, COL_PK.TRANG_THAI, COL_PK.CHI_NHANH
  ));

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
          "=COUNTIFS('Điện thoại'!$" + colThuongHieu + ":$" + colThuongHieu + "; $A" +
            r +
            "; 'Điện thoại'!$" + colChiNhanh + ":$" + colChiNhanh + "; " +
            colLetter +
            "$5; 'Điện thoại'!$" + colTrangThai + ":$" + colTrangThai + "; \"Còn hàng\")",
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

  var colTenSPCol = "Col" + COL_DT.TEN_SP;
  var colMauSacCol = "Col" + COL_DT.MAU_SAC;
  var colDungLuongCol = "Col" + COL_DT.DUNG_LUONG;
  var colMaDTCol = "Col" + COL_DT.MA_DT;
  var colTrangThaiCol = "Col" + COL_DT.TRANG_THAI_KHO;
  var colThuongHieuCol = "Col" + COL_DT.THUONG_HIEU;
  var colImeiCol = "Col" + COL_DT.IMEI;
  var colImei2Col = "Col" + COL_DT.IMEI_2;
  var colChiNhanhCol = "Col" + COL_DT.CHI_NHANH;

  // Công thức QUERY xoay chiều theo chi nhánh (có lọc theo ô Tìm kiếm B2)
  var queryFormula = 
    "=QUERY(INDEX(TO_TEXT('Điện thoại'!A:" + maxColLetter + ")); \"select " + colTenSPCol + ", " + colMauSacCol + ", " + colDungLuongCol + ", count(" + colMaDTCol + ") " +
    "where " + colMaDTCol + " is not null and " + colTrangThaiCol + " = 'Còn hàng' \" & " +
    "IF(ISBLANK($B$2); \"\"; \"and (lower(" + colTenSPCol + ") contains '\"&LOWER($B$2)&\"' or lower(" + colMauSacCol + ") contains '\"&LOWER($B$2)&\"' or lower(" + colDungLuongCol + ") contains '\"&LOWER($B$2)&\"' or lower(" + colThuongHieuCol + ") contains '\"&LOWER($B$2)&\"' or lower(" + colImeiCol + ") contains '\"&LOWER($B$2)&\"' or lower(" + colImei2Col + ") contains '\"&LOWER($B$2)&\"')\") & " +
    "\" group by " + colTenSPCol + ", " + colMauSacCol + ", " + colDungLuongCol + " pivot " + colChiNhanhCol + "\"; 1)";

  tonKhoSheet
    .getRange(startRowB2B3 + 1, 1)
    .setFormula(queryFormula);

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
      ':A; LAMBDA(val; IF(val=""; ""; SUM(OFFSET(val; 0; 3; 1; ' +
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

  // Công thức QUERY xoay chiều theo chi nhánh (có lọc theo ô Tìm kiếm B2)
  var startAccColLetter = columnToLetter(startAccColIdx);
  var queryAccFormula =
    "=QUERY('Phụ kiện'!A:" + maxColPKLetter + "; \"select " + colMaPK + ", " + colTenPK + ", " + colLoaiPK + ", " + colThuongHieuPK + ", sum(" + colSoLuongTon + ") " +
    "where " + colMaPK + " is not null and " + colTrangThaiPK + " = 'Đang bán' and " + colSoLuongTon + " > 0 \" & " +
    "IF(ISBLANK($B$2); \"\"; \"and (lower(" + colMaPK + ") contains '\"&LOWER($B$2)&\"' or lower(" + colTenPK + ") contains '\"&LOWER($B$2)&\"' or lower(" + colLoaiPK + ") contains '\"&LOWER($B$2)&\"' or lower(" + colThuongHieuPK + ") contains '\"&LOWER($B$2)&\"')\") & " +
    "\" group by " + colMaPK + ", " + colTenPK + ", " + colLoaiPK + ", " + colThuongHieuPK + " pivot " + colChiNhanhPK + "\"; 1)";

  tonKhoSheet
    .getRange(startRowB2B3 + 1, startAccColIdx)
    .setFormula(queryAccFormula);

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
      '; LAMBDA(val; IF(val=""; ""; SUM(OFFSET(val; 0; 4; 1; ' +
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
    dienThoai: getDienThoaiDropdownSearch("", ""),
    phuKien: getPhuKienDropdownSearch("", ""),
    phuKienUnique: getPhuKienUniqueListSearch(""),
    storeInterestRate: getInterestRateConfig(),
  };
}

/**
 * Lấy toàn bộ dữ liệu khởi tạo cho Sidebar (gộp 5 yêu cầu thành 1 để tối ưu tốc độ tải)
 *
 * @return {Object} Chứa preloadData, branches, staff, financeCompanies, brands
 */
function getInitialSidebarData() {
  initializeColumnEnums();
  return {
    preloadData: getPreloadData(),
    branches: getBranchesDropdown(),
    staff: getNhanVienDropdown(),
    financeCompanies: getFinanceCompaniesDropdown(),
    brands: getBrandsDropdown(),
    storeInterestRate: getInterestRateConfig(),
  };
}
