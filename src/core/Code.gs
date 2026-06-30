/**
 * ============================================================
 * VanTran Mobile — Hệ thống Dịch vụ & Buôn bán Trả góp
 * Entry Point: Code.gs
 * ============================================================
 */

// ======================== CONSTANTS ========================

// SHEET_NAMES and SHEET_ORDER have been moved to constants/Constants.gs to resolve load-order issues.


// ======================== SHEET HEADERS (VIETNAMESE) ========================

// SHEET_HEADERS has been moved to constants/Constants.gs to resolve load-order issues.

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
let _conditionalColorMapCache = null;

function getConditionalColorMap() {
  if (_conditionalColorMapCache) return _conditionalColorMapCache;

  _conditionalColorMapCache = {
    // 1. Trạng thái tốt, hoàn thành, đang làm, có, mới
    [EMPLOYEE_STATUS.ACTIVE]: { bg: "#e6f4ea", fg: "#137333" },
    [STOCK_STATUS.IN_STOCK]: { bg: "#e6f4ea", fg: "#137333" },
    [ORDER_STATUS.DONE]: { bg: "#e6f4ea", fg: "#137333" },
    [INSTALLMENT_STATUS.DONE]: { bg: "#e6f4ea", fg: "#137333" },
    [LSTG_STATUS.PAID]: { bg: "#e6f4ea", fg: "#137333" },
    [PK_STATUS.ACTIVE]: { bg: "#e6f4ea", fg: "#137333" },
    [YES_NO_VALUES.YES]: { bg: "#e6f4ea", fg: "#137333" },
    "Mới 100%": { bg: "#e6f4ea", fg: "#137333" },
    [PAYMENT_METHOD.CASH]: { bg: "#e6f4ea", fg: "#137333" },
    [SERVICE_TYPES.CASH_WITHDRAWAL]: { bg: "#e6f4ea", fg: "#137333" },

    // 2. Trạng thái nghỉ việc, huỷ, quá hạn, không, trả máy, đổi trả
    [EMPLOYEE_STATUS.INACTIVE]: { bg: "#fce8e6", fg: "#c5221f" },
    [ORDER_STATUS.CANCELLED]: { bg: "#fce8e6", fg: "#c5221f" },
    [INSTALLMENT_STATUS.LATE]: { bg: "#fce8e6", fg: "#c5221f" },
    [INSTALLMENT_STATUS.CANCELLED]: { bg: "#fce8e6", fg: "#c5221f" },
    [LSTG_STATUS.LATE]: { bg: "#fce8e6", fg: "#c5221f" },
    [LSTG_STATUS.CANCELLED]: { bg: "#fce8e6", fg: "#c5221f" },
    [PK_STATUS.INACTIVE]: { bg: "#fce8e6", fg: "#c5221f" },
    [YES_NO_VALUES.NO]: { bg: "#fce8e6", fg: "#c5221f" },
    [EXCHANGE_TYPES.RETURN_DEVICE]: { bg: "#fce8e6", fg: "#c5221f" },
    [EXCHANGE_TYPES.RETURN_GOODS]: { bg: "#fce8e6", fg: "#c5221f" },
    [STOCK_STATUS.RETURNED]: { bg: "#fce8e6", fg: "#c5221f" },
    [ORDER_STATUS.EXCHANGED]: { bg: "#fce8e6", fg: "#c5221f" },
    [STOCK_STATUS.FAULTY]: { bg: "#fce8e6", fg: "#c5221f" },

    // 3. Trạng thái chờ, đang xử lý, đang trả, chưa trả, đổi máy, phụ kiện
    [STOCK_STATUS.INSTALLMENT]: { bg: "#fef7e0", fg: "#b06000" },
    [ORDER_STATUS.PROCESSING]: { bg: "#fef7e0", fg: "#b06000" },
    [INSTALLMENT_STATUS.RUNNING]: { bg: "#fef7e0", fg: "#b06000" },
    [LSTG_STATUS.UNPAID]: { bg: "#fef7e0", fg: "#b06000" },
    [SALES_METHOD.INSTALLMENT]: { bg: "#f3e8fd", fg: "#a142f4" }, // Tím nhẹ cho trả góp
    [EXCHANGE_TYPES.EXCHANGE_DEVICE]: { bg: "#fef7e0", fg: "#b06000" },
    [EXCHANGE_TYPES.EXCHANGE_GOODS]: { bg: "#fef7e0", fg: "#b06000" },
    [SERVICE_TYPES.PHONE_TOPUP]: { bg: "#fef7e0", fg: "#b06000" },
    [INSTALLMENT_STATUS.PENDING]: { bg: "#fef7e0", fg: "#b06000" },
    "Cường lực": { bg: "#fef7e0", fg: "#b06000" },
    [PAYMENT_METHOD.POS]: { bg: "#fef7e0", fg: "#b06000" },
    "Kỹ thuật": { bg: "#fef7e0", fg: "#b06000" },

    // 4. Trạng thái khác, đã qua sử dụng, đã bán
    [STOCK_STATUS.SOLD]: { bg: "#cbd5e1", fg: "#334155" },
    "Đã qua sử dụng": { bg: "#e2e8f0", fg: "#334155" },
    "Khác": { bg: "#e2e8f0", fg: "#475569" },

    // 5. Chi nhánh, vai trò, hình thức, nguồn
    "Bán hàng": { bg: "#e8f0fe", fg: "#1a73e8" },
    [PAYMENT_METHOD.TRANSFER]: { bg: "#e8f0fe", fg: "#1a73e8" },
    [SERVICE_TYPES.TRANSFER_HO]: { bg: "#e8f0fe", fg: "#1a73e8" },
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

    [SALES_METHOD.DIRECT]: { bg: "#e6f4ea", fg: "#137333" },
    [PAYMENT_METHOD.MIXED]: { bg: "#e0f7fa", fg: "#006064" },
    "Trừ vào đơn mới": { bg: "#fef7e0", fg: "#b06000" },
    "Sửa chữa": { bg: "#e0f7fa", fg: "#006064" },
    "Bảo hành": { bg: "#fce4ec", fg: "#880e4f" }
  };

  return _conditionalColorMapCache;
}

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
  const html = HtmlService.createTemplateFromFile("ui/Sidebar");
  const validModes = UI_MODES;
  let mode = (e && e.parameter && e.parameter.mode) || "donHang";
  if (validModes.indexOf(mode) === -1) {
    mode = "donHang";
  }
  html.mode = mode;
  return html
    .evaluate()
    .setTitle("BichLoan")
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
