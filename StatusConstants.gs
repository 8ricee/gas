/**
 * ============================================================
 * VanTran Mobile — StatusConstants.gs
 * Định nghĩa hằng số trạng thái dùng chung toàn hệ thống
 * ============================================================
 */

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
