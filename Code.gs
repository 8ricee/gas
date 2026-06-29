/**
 * ============================================================
 * VanTran Mobile — Hệ thống Dịch vụ & Buôn bán Trả góp
 * Entry Point: Code.gs
 * ============================================================
 */

// ======================== CONSTANTS ========================

const SHEET_NAMES = {
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

const SHEET_ORDER = [
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

const SHEET_HEADERS = {};

// ======================== DEFAULT CONFIG ========================

const DEFAULT_CONFIG = [
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
const CONDITIONAL_COLOR_MAP = {
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
  "Máy lỗi": { bg: "#fce8e6", fg: "#c5221f" },

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
  "Trừ vào đơn mới": { bg: "#fef7e0", fg: "#b06000" },
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
  const html = HtmlService.createTemplateFromFile("Sidebar");
  html.mode = (e && e.parameter && e.parameter.mode) || "donHang";
  return html
    .evaluate()
    .setTitle("VanTran Mobile — Hệ thống Dịch vụ & Buôn bán Trả góp")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

/**
 * Chuyển số cột thành chữ cái (ví dụ: 1 -> A, 27 -> AA)
 *
 * @param {number} column - Chỉ số cột (1-indexed)
 * @return {string}
 */
function columnToLetter(column) {
  let temp;
  let letter = "";
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
  return {
    preloadData: getPreloadData(),
    branches: getBranchesDropdown(),
    staff: getNhanVienDropdown(),
    financeCompanies: getFinanceCompaniesDropdown(),
    brands: getBrandsDropdown(),
  };
}
