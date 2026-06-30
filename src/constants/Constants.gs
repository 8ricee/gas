/**
 * ============================================================
 * VanTran Mobile — StatusConstants.gs
 * Định nghĩa hằng số trạng thái dùng chung toàn hệ thống
 * ============================================================
 */

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

const SHEET_HEADERS = {};

const UI_MODES = Object.freeze([
  "donHang",
  "khachHang",
  "nhapKho",
  "dichVu",
  "thanhToan",
  "chuyenKho",
  "doiTra",
  "thuMua",
  "doanhSo",
  "baoHanh",
]);

const YES_NO_VALUES = Object.freeze({
  YES: "✓",
  NO: "✗",
});

const EMPLOYEE_ROLES = Object.freeze(["Bán hàng", "Kế toán", "Kỹ thuật"]);

const EMPLOYEE_STATUS = Object.freeze({
  ACTIVE: "Đang làm",
  INACTIVE: "Nghỉ việc",
});

const SALES_METHOD = Object.freeze({
  DIRECT: "Bán thẳng",
  INSTALLMENT: "Trả góp",
});

const SERVICE_TYPES = Object.freeze({
  TRANSFER_HO: "Chuyển khoản hộ",
  CASH_WITHDRAWAL: "Rút tiền mặt",
  PHONE_TOPUP: "Nạp thẻ điện thoại",
});

const EXCHANGE_TYPES = Object.freeze({
  RETURN_DEVICE: "Trả máy",
  EXCHANGE_DEVICE: "Đổi máy",
  RETURN_GOODS: "Trả hàng",
  EXCHANGE_GOODS: "Đổi hàng",
});

const STOCK_STATUS = Object.freeze({
  IN_STOCK: "Còn hàng",
  SOLD: "Đã bán",
  INSTALLMENT: "Đang trả góp",
  RETURNED: "Đã trả lại",
  FAULTY: "Máy lỗi",
});

const ORDER_STATUS = Object.freeze({
  DONE: "Hoàn thành",
  PROCESSING: "Đang xử lý",
  CANCELLED: "Huỷ",
  EXCHANGED: "Đổi trả",
});

const PK_STATUS = Object.freeze({
  ACTIVE: "Đang bán",
  INACTIVE: "Ngừng bán",
});

const PAYMENT_METHOD = Object.freeze({
  CASH: "Tiền mặt",
  TRANSFER: "Chuyển khoản",
  MIXED: "Hỗn hợp",
  POS: "Quẹt thẻ (POS)",
});

const INSTALLMENT_STATUS = Object.freeze({
  RUNNING: "Đang trả",
  DONE: "Hoàn tất",
  LATE: "Quá hạn",
  CANCELLED: "Đã huỷ",
  PENDING: "Chờ thanh khoản",
});

const LSTG_STATUS = Object.freeze({
  UNPAID: "Chưa trả",
  PAID: "Đã trả",
  LATE: "Quá hạn",
  CANCELLED: "Đã huỷ",
});

const INSTALLMENT_TYPE = Object.freeze({
  STORE: "Cửa hàng",
  FINANCE: "Công ty tài chính",
});

const PRODUCT_SOURCE = Object.freeze({
  PHONE: "Điện thoại",
  ACCESSORY: "Phụ kiện",
});

/**
 * Kiểm tra trạng thái có phải là "Huỷ" hay không (chấp nhận cả chữ thường, không dấu)
 * @param {any} value - Trạng thái cần kiểm tra
 * @return {boolean}
 */
function isCancelStatus(value) {
  const s = String(value || "").trim().toLowerCase();
  return s === "huỷ" || s === "huy" || s === "đã huỷ";
}
