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
  const html = HtmlService.createTemplateFromFile("Sidebar");
  return html
    .evaluate()
    .setTitle("VanTran Mobile — Hệ thống Dịch vụ & Buôn bán Trả góp")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ======================== EDIT TRIGGER REGISTER ========================

EDIT_HANDLERS[SHEET_NAMES.CAU_HINH] = function (sheet, row, col, e) {
  _setupDataValidations(e.source);
};
EDIT_HANDLERS[SHEET_NAMES.DON_HANG] = _onEditDonHang;
EDIT_HANDLERS[SHEET_NAMES.NHAP_KHO] = _onEditNhapKho;
EDIT_HANDLERS[SHEET_NAMES.BAO_CAO_NGAY] = _onEditBaoCaoNgay;
EDIT_HANDLERS[SHEET_NAMES.BAO_CAO_DOANH_SO] = _onEditBaoCaoDoanhSo;
EDIT_HANDLERS[SHEET_NAMES.LICH_SU_TRA_GOP] = _onEditLichSuTraGop;

/**
 * onEdit trigger — auto-calculate khi chỉnh sửa
 */
function onEdit(e) {
  if (!e || e.range.getRow() <= 1) return;
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Kiểm tra xem cột được sửa đổi có liên quan đến các bộ xử lý (EDIT_HANDLERS) không.
  // Tránh xoá cache và chạy xử lý vô ích đối với các cột không cần auto-lookup/auto-calculate.
  let isRelevant = false;
  if (sheetName === SHEET_NAMES.DON_HANG) {
    isRelevant = (col === COL_DH.SO_LUONG || col === COL_DH.DON_GIA || col === COL_DH.TIEN_GIAM_GIA ||
                  col === COL_DH.MA_KH || col === COL_DH.MA_SP || col === COL_DH.NGUOI_BAN || col === COL_DH.NGUOI_HO_TRO);
  } else if (sheetName === SHEET_NAMES.NHAP_KHO) {
    isRelevant = (col === COL_NK.SO_LUONG || col === COL_NK.GIA_NHAP || col === COL_NK.MA_SP);
  } else if (sheetName === SHEET_NAMES.LICH_SU_TRA_GOP) {
    isRelevant = (col === COL_LSTG.SO_TIEN_DA_TRA || col === COL_LSTG.TRANG_THAI);
  } else if (sheetName === SHEET_NAMES.BAO_CAO_NGAY) {
    isRelevant = (row === 3 && col === 2);
  } else if (sheetName === SHEET_NAMES.BAO_CAO_DOANH_SO) {
    isRelevant = (row === 3 && (col === 2 || col === 4 || col === 6 || col === 8));
  } else if (sheetName === SHEET_NAMES.CAU_HINH) {
    isRelevant = true; // Luôn chạy khi sửa đổi cấu hình để cập nhật validation
  }

  if (!isRelevant) return;

  try {
    clearSheetCache(sheetName);
    invalidateDropdownCache(sheetName);
    if (EDIT_HANDLERS[sheetName]) {
      EDIT_HANDLERS[sheetName](sheet, row, col, e);
    }
  } catch (err) {
    Logger.log("[onEdit][" + sheetName + "] Error: " + err.message);
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
  
  // Lắng nghe thay đổi Số tiền đã trả hoặc Trạng thái
  if (col !== COL_LSTG.SO_TIEN_DA_TRA && col !== COL_LSTG.TRANG_THAI) return;

  const maTG = sheet.getRange(row, COL_LSTG.MA_TG).getValue();
  if (!maTG) return;

  // Nếu sửa cột trạng thái
  if (col === COL_LSTG.TRANG_THAI) {
    // Tận dụng e.value nếu có (nhanh hơn), nếu không có mới lấy từ sheet (trường hợp xoá/paste nhiều ô)
    const status = e.value !== undefined ? String(e.value).trim() : String(sheet.getRange(row, COL_LSTG.TRANG_THAI).getValue()).trim();

    if (status === "Đã trả") {
      const ngayThucTraRange = sheet.getRange(row, COL_LSTG.NGAY_THUC_TRA);
      if (!ngayThucTraRange.getValue()) ngayThucTraRange.setValue(new Date());

      const htttRange = sheet.getRange(row, COL_LSTG.HINH_THUC_TT);
      if (!htttRange.getValue()) htttRange.setValue("Tiền mặt");

      const soTienCanTra = Number(sheet.getRange(row, COL_LSTG.SO_TIEN_CAN_TRA).getValue()) || 0;
      const soTienDaTraRange = sheet.getRange(row, COL_LSTG.SO_TIEN_DA_TRA);
      if (!soTienDaTraRange.getValue()) soTienDaTraRange.setValue(soTienCanTra);

    } else if (status === "Chưa trả" || status === "Đã huỷ") {
      // Gọi clearContent() sẽ đúng bản chất hơn là setValue("") khi muốn làm trống ô
      sheet.getRange(row, COL_LSTG.SO_TIEN_DA_TRA).clearContent();
      sheet.getRange(row, COL_LSTG.NGAY_THUC_TRA).clearContent();
      sheet.getRange(row, COL_LSTG.HINH_THUC_TT).clearContent();
    }
  }

  // Cập nhật lại tổng trong TraGop và trạng thái máy (chạy cho cả 2 trường hợp thay đổi cột trạng thái hoặc số tiền)
  _capNhatTongTraGop(maTG);
}



/**
 * Auto-calculate cho sheet DonHang (Đơn hàng)
 */
function _onEditDonHang(sheet, row, col, e) {

  // Auto tính ThanhTien khi SoLuong hoặc DonGia hoặc TienGiamGia thay đổi
  if (col === COL_DH.SO_LUONG || col === COL_DH.DON_GIA || col === COL_DH.TIEN_GIAM_GIA) {
    const soLuong = sheet.getRange(row, COL_DH.SO_LUONG).getValue() || 0;
    const donGia = sheet.getRange(row, COL_DH.DON_GIA).getValue() || 0;
    const giamGia = sheet.getRange(row, COL_DH.TIEN_GIAM_GIA).getValue() || 0;
    sheet.getRange(row, COL_DH.THANH_TIEN).setValue(soLuong * donGia - giamGia);
  }

  // Auto lookup TenKH khi nhập MaKH
  if (col === COL_DH.MA_KH) {
    const maKH = e.value;
    if (maKH) {
      const tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, COL_KH.MA_KH, maKH, COL_KH.HO_TEN);
      sheet.getRange(row, COL_DH.TEN_KH).setValue(tenKH || "");
    }
  }

  // Auto lookup TenSP, NguonSP, ThuongHieu khi nhập MaSP
  if (col === COL_DH.MA_SP) {
    const maSP = e.value;
    if (maSP) {
      // Thử tìm trong Điện thoại trước
      const tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.TEN_SP);
      if (tenDT) {
        sheet.getRange(row, COL_DH.TEN_SP).setValue(tenDT);
        sheet.getRange(row, COL_DH.NGUON_SP).setValue("Điện thoại");
        const thuongHieu = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.THUONG_HIEU);
        sheet.getRange(row, COL_DH.THUONG_HIEU).setValue(thuongHieu || "");
        const giaBan = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.GIA_BAN);
        sheet.getRange(row, COL_DH.DON_GIA).setValue(giaBan || 0);
        sheet.getRange(row, COL_DH.SO_LUONG).setValue(1); // Điện thoại luôn SL = 1
      } else {
        // Thử tìm trong Phụ kiện
        const tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.TEN_SP);
        if (tenPK) {
          sheet.getRange(row, COL_DH.TEN_SP).setValue(tenPK);
          sheet.getRange(row, COL_DH.NGUON_SP).setValue("Phụ kiện");
          const thuongHieuPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.THUONG_HIEU);
          sheet.getRange(row, COL_DH.THUONG_HIEU).setValue(thuongHieuPK || "");
          const giaBanPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.GIA_BAN);
          sheet.getRange(row, COL_DH.DON_GIA).setValue(giaBanPK || 0);
        }
      }
    }
  }

  // Auto lookup TenNguoiBan + QuyenXuatMay khi nhập NguoiBan (Sử dụng dynamic enum COL_NV)
  if (col === COL_DH.NGUOI_BAN) {
    const maNVBan = e.value;
    if (maNVBan) {
      const tenNVBan = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNVBan, COL_NV.HO_TEN);
      sheet.getRange(row, COL_DH.TEN_NGUOI_BAN).setValue(tenNVBan || "");
      const quyenXM = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNVBan, COL_NV.QUYEN_XUAT);
      sheet.getRange(row, COL_DH.QUYEN_XUAT).setValue(quyenXM || "✗");
    }
  }

  // Auto lookup TenNguoiHoTro khi nhập NguoiHoTro (Sử dụng dynamic enum COL_NV)
  if (col === COL_DH.NGUOI_HO_TRO) {
    const maNVHT = e.value;
    if (maNVHT) {
      const tenNVHT = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNVHT, COL_NV.HO_TEN);
      sheet.getRange(row, COL_DH.TEN_NGUOI_HO_TRO).setValue(tenNVHT || "");
    }
  }

}

/**
 * Auto-calculate cho sheet NhapKho (Nhập kho)
 */
function _onEditNhapKho(sheet, row, col, e) {

  // Auto tính ThanhTien khi SoLuong hoặc GiaNhap thay đổi
  if (col === COL_NK.SO_LUONG || col === COL_NK.GIA_NHAP) {
    const soLuong = sheet.getRange(row, COL_NK.SO_LUONG).getValue() || 0;
    const giaNhap = sheet.getRange(row, COL_NK.GIA_NHAP).getValue() || 0;
    sheet.getRange(row, COL_NK.THANH_TIEN).setValue(soLuong * giaNhap);
  }

  // Auto lookup TenSP khi nhập MaSP
  if (col === COL_NK.MA_SP) {
    const maSP = e.value;
    const nguonNhap = sheet.getRange(row, COL_NK.NGUON_NHAP).getValue();
    if (maSP) {
      if (nguonNhap === "Điện thoại") {
        const tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.TEN_SP);
        sheet.getRange(row, COL_NK.TEN_SP).setValue(tenDT || "");
      } else {
        const tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.TEN_SP);
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
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  const tonKhoSheet = ss.getSheetByName(SHEET_NAMES.TON_KHO);
  if (!tonKhoSheet) return;

  const branches = getBranchesList();
  const brands = getBrandsList();

  const N = branches.length || 1;
  const totalSummaryColIdx = 2 + N;

  // 1. Chuẩn bị sheet và ô tìm kiếm
  _prepareTonKhoSheet(tonKhoSheet, totalSummaryColIdx);

  // 2. Xây dựng Bảng 1: Tổng số máy điện thoại của mỗi chi nhánh
  const startRowB2B3 = _buildSummaryTable(tonKhoSheet, branches, brands);

  // 3. Xây dựng Bảng 2: Tồn kho điện thoại theo chi nhánh
  const totalPhoneColIdx = _buildPhoneInventoryTable(tonKhoSheet, startRowB2B3, branches);

  // 4. Xây dựng Bảng 3: Tồn kho phụ kiện theo chi nhánh
  const accInfo = _buildAccessoryInventoryTable(tonKhoSheet, startRowB2B3, branches, totalPhoneColIdx);

  // 5. Đảm bảo số cột trong sheet đủ lớn và định dạng font toàn bảng
  const lastRow = tonKhoSheet.getMaxRows();
  let lastCol = tonKhoSheet.getMaxColumns();
  const neededCols = Math.max(accInfo.totalAccColIdx, totalSummaryColIdx);
  if (lastCol < neededCols) {
    tonKhoSheet.insertColumnsAfter(lastCol, neededCols - lastCol);
    lastCol = neededCols;
  }

  const fullRange = tonKhoSheet.getRange(1, 1, lastRow, lastCol);
  fullRange.setFontFamily("Times New Roman").setFontSize(12);

  // Điều chỉnh độ rộng cột
  for (let c = 1; c <= lastCol; c++) {
    if (c === accInfo.spacerColIdx) {
      tonKhoSheet.setColumnWidth(c, 40);
    } else {
      tonKhoSheet.autoResizeColumn(c);
    }
  }
}

/**
 * Chuẩn bị sheet Tồn kho: clear dữ liệu cũ, thiết lập tiêu đề và ô Tìm kiếm
 * @private
 */
function _prepareTonKhoSheet(sheet, totalSummaryColIdx) {
  sheet.clear();
  sheet.clearFormats();

  // Dòng 1: Tiêu đề trang tính
  const sheetTitleRange = sheet.getRange(1, 1, 1, totalSummaryColIdx);
  sheetTitleRange.merge();
  sheetTitleRange.setValue("BÁO CÁO TỒN KHO HỆ THỐNG & TÌM KIẾM");
  sheetTitleRange
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  // Dòng 2: Ô tìm kiếm nhanh
  const searchLabelCell = sheet.getRange(2, 1);
  searchLabelCell
    .setValue("Tìm kiếm:")
    .setFontWeight("bold")
    .setHorizontalAlignment("right")
    .setBackground("#f1f3f4");

  const searchInputCell = sheet.getRange(2, 2);
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

  const searchTipRange = sheet.getRange(
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
  sheet.setRowHeight(3, 15);
}

/**
 * Xây dựng Bảng 1: Tổng số máy điện thoại của mỗi chi nhánh
 * @private
 * @return {number} Chỉ số dòng tiếp theo sau bảng 1
 */
function _buildSummaryTable(sheet, branches, brands) {
  const N = branches.length || 1;
  const B = brands.length || 1;
  const totalSummaryColIdx = 2 + N;

  const colThuongHieu = columnToLetter(COL_DT.THUONG_HIEU);
  const colChiNhanh = columnToLetter(COL_DT.CHI_NHANH);
  const colTrangThai = columnToLetter(COL_DT.TRANG_THAI_KHO);

  // Dòng 4: Tiêu đề Bảng 1
  const summaryTitleRange = sheet.getRange(4, 1, 1, totalSummaryColIdx);
  summaryTitleRange.merge();
  summaryTitleRange.setValue("TỔNG SỐ MÁY ĐIỆN THOẠI CỦA MỖI CHI NHÁNH");
  summaryTitleRange
    .setFontWeight("bold")
    .setFontSize(12)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  // Dòng 5: Headers Bảng 1
  const summaryHeaders = ["Thương hiệu"];
  branches.forEach(function (branch) {
    summaryHeaders.push(branch);
  });
  summaryHeaders.push("Tổng cộng");
  sheet.getRange(5, 1, 1, totalSummaryColIdx).setValues([summaryHeaders]);

  // Dòng 6 đến 5 + B: Ghi danh sách thương hiệu & công thức
  const brandValues = [];
  brands.forEach(function (brand) {
    brandValues.push([brand]);
  });
  sheet.getRange(6, 1, B, 1).setValues(brandValues);

  // Chèn công thức COUNTIFS cho các chi nhánh & SUM cho cột Tổng cộng
  for (let r = 6; r <= 5 + B; r++) {
    for (let c = 2; c <= 1 + N; c++) {
      const colLetter = columnToLetter(c);
      sheet
        .getRange(r, c)
        .setFormula(
          "=COUNTIFS('Điện thoại'!$" + colThuongHieu + ":$" + colThuongHieu + "; $A" +
            r +
            "; 'Điện thoại'!$" + colChiNhanh + ":$" + colChiNhanh + "; " +
            colLetter +
            "$5; 'Điện thoại'!$" + colTrangThai + ":$" + colTrangThai + "; \"Còn hàng\")",
        );
    }
    const lastBranchColLetter = columnToLetter(1 + N);
    sheet
      .getRange(r, 2 + N)
      .setFormula("=SUM(B" + r + ":" + lastBranchColLetter + r + ")");
  }

  // Dòng 6 + B: Dòng Tổng cộng của bảng 1
  const totalRowIdx = 6 + B;
  sheet.getRange(totalRowIdx, 1).setValue("Tổng cộng");
  for (let c = 2; c <= 1 + N; c++) {
    const colLetter = columnToLetter(c);
    sheet
      .getRange(totalRowIdx, c)
      .setFormula("=SUM(" + colLetter + "6:" + colLetter + (5 + B) + ")");
  }

  const lastBranchColLetter = columnToLetter(1 + N);
  sheet
    .getRange(totalRowIdx, 2 + N)
    .setFormula(
      "=SUM(B" + totalRowIdx + ":" + lastBranchColLetter + totalRowIdx + ")",
    );

  // Format Table 1
  const summaryHeaderRange = sheet.getRange(5, 1, 1, totalSummaryColIdx);
  summaryHeaderRange
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  const summaryTotalRowRange = sheet.getRange(
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
  sheet.setRowHeight(7 + B, 15);

  return 8 + B; // Dòng bắt đầu của Bảng 2 và 3
}

/**
 * Xây dựng Bảng 2: Tồn kho điện thoại theo chi nhánh
 * @private
 * @return {number} Chỉ số cột Tổng tồn của điện thoại
 */
function _buildPhoneInventoryTable(sheet, startRow, branches) {
  const N = branches.length || 1;
  const totalPhoneColIdx = 3 + N + 1; // 3 cột thông tin + N chi nhánh + 1 cột tổng

  // Tiêu đề Bảng 2
  const phoneTitleRange = sheet.getRange(
    startRow,
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

  const maxColLetter = columnToLetter(Math.max(
    COL_DT.MA_DT, COL_DT.TEN_SP, COL_DT.THUONG_HIEU, COL_DT.IMEI, COL_DT.IMEI_2,
    COL_DT.MAU_SAC, COL_DT.DUNG_LUONG, COL_DT.TRANG_THAI_KHO, COL_DT.CHI_NHANH
  ));

  // Công thức QUERY xoay chiều theo chi nhánh (có lọc theo ô Tìm kiếm B2)
  const queryFormula = _buildPhoneQueryFormula(maxColLetter, {
    colTenSP: "Col" + COL_DT.TEN_SP,
    colMauSac: "Col" + COL_DT.MAU_SAC,
    colDungLuong: "Col" + COL_DT.DUNG_LUONG,
    colMaDT: "Col" + COL_DT.MA_DT,
    colTrangThai: "Col" + COL_DT.TRANG_THAI_KHO,
    colThuongHieu: "Col" + COL_DT.THUONG_HIEU,
    colImei: "Col" + COL_DT.IMEI,
    colImei2: "Col" + COL_DT.IMEI_2,
    colChiNhanh: "Col" + COL_DT.CHI_NHANH
  });

  sheet
    .getRange(startRow + 1, 1)
    .setFormula(queryFormula);

  // Đầu cột Tổng tồn
  const phoneTotalHeaderCell = sheet.getRange(
    startRow + 1,
    totalPhoneColIdx,
  );
  phoneTotalHeaderCell
    .setValue("Tổng tồn")
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8");

  // Công thức tính Tổng tồn dùng MAP và OFFSET động từ dòng startRow + 2
  const phoneTotalFormulaCell = sheet.getRange(
    startRow + 2,
    totalPhoneColIdx,
  );
  phoneTotalFormulaCell.setFormula(
    "=MAP(A" +
      (startRow + 2) +
      ':A; LAMBDA(val; IF(val=""; ""; SUM(OFFSET(val; 0; 3; 1; ' +
      N +
      ")))))",
  );

  // Định dạng Header Dòng startRow + 1
  const phoneHeaderRange = sheet.getRange(
    startRow + 1,
    1,
    1,
    totalPhoneColIdx,
  );
  phoneHeaderRange
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  return totalPhoneColIdx;
}

/**
 * Xây dựng Bảng 3: Tồn kho phụ kiện theo chi nhánh
 * @private
 * @return {Object} { spacerColIdx, totalAccColIdx }
 */
function _buildAccessoryInventoryTable(sheet, startRow, branches, totalPhoneColIdx) {
  const N = branches.length || 1;
  const spacerColIdx = totalPhoneColIdx + 1; // Cột khoảng cách trống
  const startAccColIdx = spacerColIdx + 1; // Cột bắt đầu bảng Phụ kiện
  const totalAccColIdx = startAccColIdx + 4 + N; // 4 cột thông tin + N chi nhánh + 1 cột tổng

  const accTitleRange = sheet.getRange(
    startRow,
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

  const maxColPKLetter = columnToLetter(Math.max(
    COL_PK.MA_PK, COL_PK.TEN_SP, COL_PK.LOAI_PK, COL_PK.THUONG_HIEU,
    COL_PK.SO_LUONG_TON, COL_PK.TRANG_THAI, COL_PK.CHI_NHANH
  ));

  const startAccColLetter = columnToLetter(startAccColIdx);

  // Công thức QUERY xoay chiều theo chi nhánh (có lọc theo ô Tìm kiếm B2)
  const queryAccFormula = _buildAccessoryQueryFormula(maxColPKLetter, {
    colMaPK: columnToLetter(COL_PK.MA_PK),
    colTenPK: columnToLetter(COL_PK.TEN_SP),
    colLoaiPK: columnToLetter(COL_PK.LOAI_PK),
    colThuongHieuPK: columnToLetter(COL_PK.THUONG_HIEU),
    colSoLuongTon: columnToLetter(COL_PK.SO_LUONG_TON),
    colTrangThai: columnToLetter(COL_PK.TRANG_THAI),
    colChiNhanh: columnToLetter(COL_PK.CHI_NHANH)
  });

  sheet
    .getRange(startRow + 1, startAccColIdx)
    .setFormula(queryAccFormula);

  // Đầu cột Tổng tồn của Phụ kiện
  const accTotalHeaderCell = sheet.getRange(
    startRow + 1,
    totalAccColIdx,
  );
  accTotalHeaderCell
    .setValue("Tổng tồn")
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8");

  // Công thức tính Tổng tồn Phụ kiện
  const accTotalFormulaCell = sheet.getRange(
    startRow + 2,
    totalAccColIdx,
  );
  accTotalFormulaCell.setFormula(
    "=MAP(" +
      startAccColLetter +
      (startRow + 2) +
      ":" +
      startAccColLetter +
      '; LAMBDA(val; IF(val=""; ""; SUM(OFFSET(val; 0; 4; 1; ' +
      N +
      ")))))",
  );

  // Định dạng Header Dòng startRow + 1
  const accHeaderRange = sheet.getRange(
    startRow + 1,
    startAccColIdx,
    1,
    4 + N + 1,
  );
  accHeaderRange
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  return { spacerColIdx: spacerColIdx, totalAccColIdx: totalAccColIdx };
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
    storeInterestRate: getInterestRateConfig(),
  };
}

/**
 * Tạo công thức QUERY cho bảng tồn kho điện thoại
 * @private
 */
function _buildPhoneQueryFormula(maxColLetter, params) {
  const {
    colTenSP,
    colMauSac,
    colDungLuong,
    colMaDT,
    colTrangThai,
    colThuongHieu,
    colImei,
    colImei2,
    colChiNhanh,
    searchCell = "$B$2"
  } = params;

  const selectClause = `select ${colTenSP}, ${colMauSac}, ${colDungLuong}, count(${colMaDT})`;
  const whereBase = `where ${colMaDT} is not null and ${colTrangThai} = 'Còn hàng'`;
  
  const searchFilter = `and (lower(${colTenSP}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colMauSac}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colDungLuong}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colThuongHieu}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colImei}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colImei2}) contains '\"&LOWER(${searchCell})&\"')`;

  const groupPivotClause = `group by ${colTenSP}, ${colMauSac}, ${colDungLuong} pivot ${colChiNhanh}`;

  return `=QUERY(INDEX(TO_TEXT('Điện thoại'!A:${maxColLetter})); "${selectClause} ${whereBase} " & IF(ISBLANK(${searchCell}); ""; "${searchFilter}") & " ${groupPivotClause}"; 1)`;
}

/**
 * Tạo công thức QUERY cho bảng tồn kho phụ kiện
 * @private
 */
function _buildAccessoryQueryFormula(maxColPKLetter, params) {
  const {
    colMaPK,
    colTenPK,
    colLoaiPK,
    colThuongHieuPK,
    colSoLuongTon,
    colTrangThai,
    colChiNhanh,
    searchCell = "$B$2"
  } = params;

  const selectClause = `select ${colMaPK}, ${colTenPK}, ${colLoaiPK}, ${colThuongHieuPK}, sum(${colSoLuongTon})`;
  const whereBase = `where ${colMaPK} is not null and ${colTrangThai} = 'Đang bán' and ${colSoLuongTon} > 0`;

  const searchFilter = `and (lower(${colMaPK}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colTenPK}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colLoaiPK}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colThuongHieuPK}) contains '\"&LOWER(${searchCell})&\"')`;

  const groupPivotClause = `group by ${colMaPK}, ${colTenPK}, ${colLoaiPK}, ${colThuongHieuPK} pivot ${colChiNhanh}`;

  return `=QUERY('Phụ kiện'!A:${maxColPKLetter}; "${selectClause} ${whereBase} " & IF(ISBLANK(${searchCell}); ""; "${searchFilter}") & " ${groupPivotClause}"; 1)`;
}
