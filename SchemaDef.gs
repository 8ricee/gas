/**
 * ============================================================
 * VanTran Mobile — SchemaDef.gs
 * Single Source of Truth for spreadsheet schema, enums & headers
 * ============================================================
 */

// Define the Single Source of Truth schema mapping
// Each array holds [property_code_name, column_header_label] in exact column order (0-indexed)
const SCHEMA = {};

SCHEMA[SHEET_NAMES.CAU_HINH] = [
  ["TEN_CFG", "Tên cấu hình"],
  ["GIA_TRI", "Giá trị"],
  ["GHI_CHU", "Ghi chú"]
];

SCHEMA[SHEET_NAMES.NHAN_VIEN] = [
  ["MA_NV", "Mã nhân viên"],
  ["HO_TEN", "Họ tên"],
  ["SO_DIEN_THOAI", "Số điện thoại"],
  ["EMAIL", "Email"],
  ["VAI_TRO", "Vai trò"],
  ["QUYEN_XUAT", "Quyền xuất máy"],
  ["NGAY_VAO", "Ngày vào làm"],
  ["TRANG_THAI", "Trạng thái"]
];

SCHEMA[SHEET_NAMES.DIEN_THOAI] = [
  ["MA_DT", "Mã điện thoại"],
  ["TEN_SP", "Tên sản phẩm"],
  ["THUONG_HIEU", "Thương hiệu"],
  ["IMEI", "IMEI"],
  ["IMEI_2", "IMEI 2"],
  ["MAU_SAC", "Màu sắc"],
  ["DUNG_LUONG", "Dung lượng"],
  ["TINH_TRANG", "Tình trạng"],
  ["GIA_NHAP", "Giá nhập"],
  ["GIA_BAN", "Giá bán"],
  ["GIA_TRA_GOP", "Giá trả góp"],
  ["TRANG_THAI_KHO", "Trạng thái kho"],
  ["GHI_CHU", "Ghi chú"],
  ["CHI_NHANH", "Chi nhánh"],
  ["NGAY_NHAP", "Ngày nhập"],
  ["NGAY_XUAT", "Ngày xuất"]
];

SCHEMA[SHEET_NAMES.PHU_KIEN] = [
  ["MA_PK", "Mã phụ kiện"],
  ["TEN_SP", "Tên sản phẩm"],
  ["LOAI_PK", "Loại phụ kiện"],
  ["THUONG_HIEU", "Thương hiệu"],
  ["GIA_NHAP", "Giá nhập"],
  ["GIA_BAN", "Giá bán"],
  ["SO_LUONG_TON", "Số lượng tồn"],
  ["MO_TA", "Mô tả"],
  ["TRANG_THAI", "Trạng thái"],
  ["CHI_NHANH", "Chi nhánh"]
];

SCHEMA[SHEET_NAMES.KHACH_HANG] = [
  ["MA_KH", "Mã khách hàng"],
  ["HO_TEN", "Họ tên"],
  ["CCCD", "CCCD"],
  ["DIA_CHI", "Địa chỉ"],
  ["NGAY_TAO", "Ngày tạo"],
  ["GHI_CHU", "Ghi chú"]
];

SCHEMA[SHEET_NAMES.DON_HANG] = [
  ["MA_DH", "Mã đơn hàng"],
  ["NGAY_BAN", "Ngày bán"],
  ["MA_KH", "Mã khách hàng"],
  ["TEN_KH", "Tên khách hàng"],
  ["MA_SP", "Mã sản phẩm"],
  ["TEN_SP", "Tên sản phẩm"],
  ["IMEI", "IMEI"],
  ["NGUON_SP", "Nguồn sản phẩm"],
  ["THUONG_HIEU", "Thương hiệu"],
  ["SO_LUONG", "Số lượng"],
  ["DON_GIA", "Đơn giá"],
  ["TIEN_THU_MUA", "Trừ tiền thu máy"],
  ["THANH_TIEN", "Thành tiền"],
  ["HINH_THUC_BAN", "Hình thức bán"],
  ["HINH_THUC_TT", "Hình thức thanh toán"],
  ["NGUOI_BAN", "Người bán"],
  ["TEN_NGUOI_BAN", "Tên người bán"],
  ["QUYEN_XUAT", "Có quyền xuất máy"],
  ["NGUOI_HO_TRO", "Người hỗ trợ"],
  ["TEN_NGUOI_HO_TRO", "Tên người hỗ trợ"],
  ["TRANG_THAI", "Trạng thái"],
  ["GHI_CHU", "Ghi chú"],
  ["CHI_NHANH", "Chi nhánh"],
  ["MA_QUA_TANG", "Mã quà tặng"],
  ["TEN_QUA_TANG", "Tên quà tặng"],
  ["CO_NHAN_QUA", "Có nhận quà"],
  ["TIEN_GIAM_GIA", "Tiền giảm giá"],
  ["TIEN_MAT", "Tiền mặt"],
  ["CHUYEN_KHOAN", "Chuyển khoản"]
];

SCHEMA[SHEET_NAMES.DICH_VU] = [
  ["MA_DV", "Mã dịch vụ"],
  ["NGAY_GD", "Ngày giao dịch"],
  ["LOAI_DV", "Loại dịch vụ"],
  ["MA_KH", "Mã khách hàng"],
  ["TEN_KH", "Tên khách hàng"],
  ["SO_TIEN_GD", "Số tiền giao dịch"],
  ["PHI_DV", "Phí dịch vụ"],
  ["HINH_THUC_TT", "Hình thức thanh toán"],
  ["NGUOI_THUC_HIEN", "Người thực hiện"],
  ["TEN_NGUOI_THUC_HIEN", "Tên người thực hiện"],
  ["TRANG_THAI", "Trạng thái"],
  ["GHI_CHU", "Ghi chú"],
  ["CHI_NHANH", "Chi nhánh"],
  ["TIEN_MAT", "Tiền mặt"],
  ["CHUYEN_KHOAN", "Chuyển khoản"]
];

SCHEMA[SHEET_NAMES.TRA_GOP] = [
  ["MA_TG", "Mã trả góp"],
  ["MA_DH", "Mã đơn hàng"],
  ["MA_KH", "Mã khách hàng"],
  ["TEN_KH", "Tên khách hàng"],
  ["TONG_TIEN", "Tổng tiền"],
  ["TRA_TRUOC", "Trả trước"],
  ["CON_LAI", "Còn lại"],
  ["SO_KY", "Số kỳ"],
  ["TIEN_MOI_KY", "Tiền mỗi kỳ"],
  ["NGAY_BAT_DAU", "Ngày bắt đầu"],
  ["NGAY_KET_THUC", "Ngày kết thúc"],
  ["DA_TRA_SO_KY", "Đã trả số kỳ"],
  ["DA_TRA_SO_TIEN", "Đã trả số tiền"],
  ["LOAI_TRA_GOP", "Loại trả góp"],
  ["CONG_TY_TC", "Công ty tài chính"],
  ["TRANG_THAI", "Trạng thái"],
  ["CHI_NHANH", "Chi nhánh"],
  ["TIEN_MAT", "Tiền mặt"],
  ["CHUYEN_KHOAN", "Chuyển khoản"]
];

SCHEMA[SHEET_NAMES.LICH_SU_TRA_GOP] = [
  ["MA_LS", "Mã lịch sử"],
  ["MA_TG", "Mã trả góp"],
  ["KY_SO", "Kỳ số"],
  ["SO_TIEN_CAN_TRA", "Số tiền cần trả"],
  ["SO_TIEN_DA_TRA", "Số tiền đã trả"],
  ["NGAY_CAN_TRA", "Ngày cần trả"],
  ["NGAY_THUC_TRA", "Ngày thực trả"],
  ["HINH_THUC_TT", "Hình thức thanh toán"],
  ["TRANG_THAI", "Trạng thái"],
  ["GHI_CHU", "Ghi chú"],
  ["TIEN_MAT", "Tiền mặt"],
  ["CHUYEN_KHOAN", "Chuyển khoản"]
];

SCHEMA[SHEET_NAMES.DOANH_SO] = [
  ["THANG_NAM", "Tháng năm"],
  ["MA_NV", "Mã nhân viên"],
  ["TEN_NV", "Tên nhân viên"],
  ["VAI_TRO", "Vai trò"],
  ["QUYEN_XUAT", "Có quyền xuất máy"],
  ["SO_MAY_BAN_AP", "Số máy bán Apple"],
  ["SO_MAY_BAN_KHAC", "Số máy bán khác"],
  ["SO_MAY_HT_AP", "Số máy hỗ trợ Apple"],
  ["SO_MAY_HT_KHAC", "Số máy hỗ trợ khác"],
  ["HOA_HONG_BAN", "Hoa hồng bán"],
  ["HOA_HONG_HT", "Hoa hồng hỗ trợ"],
  ["TONG_HOA_HONG", "Tổng hoa hồng"],
  ["DOANH_THU_DV", "Doanh thu dịch vụ"],
  ["THUONG", "Thưởng"],
  ["TONG_THU_NHAP", "Tổng thu nhập"],
  ["TRANG_THAI", "Trạng thái"]
];

SCHEMA[SHEET_NAMES.NHAP_KHO] = [
  ["MA_NK", "Mã nhập kho"],
  ["NGAY_NHAP", "Ngày nhập"],
  ["NGUON_NHAP", "Nguồn nhập"],
  ["MA_SP", "Mã sản phẩm"],
  ["TEN_SP", "Tên sản phẩm"],
  ["SO_LUONG", "Số lượng"],
  ["GIA_NHAP", "Giá nhập"],
  ["THANH_TIEN", "Thành tiền"],
  ["NHA_CUNG_CAP", "Nhà cung cấp"],
  ["GHI_CHU", "Ghi chú"],
  ["CHI_NHANH", "Chi nhánh"]
];

SCHEMA[SHEET_NAMES.DOI_TRA] = [
  ["MA_DT", "Mã đổi trả"],
  ["NGAY_DT", "Ngày đổi trả"],
  ["MA_DH", "Mã đơn hàng"],
  ["MA_KH", "Mã khách hàng"],
  ["TEN_KH", "Tên khách hàng"],
  ["LOAI_GD", "Loại giao dịch"],
  ["MA_SP_TRA", "Mã sản phẩm trả"],
  ["TEN_SP_TRA", "Tên sản phẩm trả"],
  ["IMEI_TRA", "IMEI trả"],
  ["MA_SP_NHAN", "Mã sản phẩm nhận"],
  ["TEN_SP_NHAN", "Tên sản phẩm nhận"],
  ["IMEI_NHAN", "IMEI nhận"],
  ["TIEN_HOAN_TRA", "Tiền hoàn trả"],
  ["PHI_DOI_TRA", "Phí đổi trả"],
  ["HINH_THUC_TT", "Hình thức thanh toán"],
  ["CHI_NHANH", "Chi nhánh"],
  ["NGUOI_THUC_HIEN", "Người thực hiện"],
  ["TRANG_THAI", "Trạng thái"],
  ["GHI_CHU", "Ghi chú"],
  ["TIEN_MAT", "Tiền mặt"],
  ["CHUYEN_KHOAN", "Chuyển khoản"]
];

SCHEMA[SHEET_NAMES.THU_MUA] = [
  ["MA_TM", "Mã thu mua"],
  ["NGAY_TM", "Ngày thu mua"],
  ["MA_KH", "Mã khách hàng"],
  ["TEN_KH", "Tên khách hàng"],
  ["TEN_SP_THU", "Tên sản phẩm thu"],
  ["THUONG_HIEU_THU", "Thương hiệu thu"],
  ["IMEI_THU", "IMEI thu"],
  ["MAU_SAC_THU", "Màu sắc thu"],
  ["DUNG_LUONG_THU", "Dung lượng thu"],
  ["TINH_TRANG_THU", "Tình trạng thu"],
  ["GIA_THU_MUA", "Giá thu mua"],
  ["LOAI_GD", "Loại giao dịch"],
  ["MA_DH_MOI", "Mã đơn hàng mới"],
  ["TIEN_HO_TRO", "Tiền hỗ trợ"],
  ["TONG_TIEN_TRA", "Tổng tiền trả khách"],
  ["HINH_THUC_TT", "Hình thức thanh toán"],
  ["CHI_NHANH", "Chi nhánh"],
  ["NGUOI_THUC_HIEN", "Người thực hiện"],
  ["TRANG_THAI", "Trạng thái"],
  ["GHI_CHU", "Ghi chú"],
  ["TIEN_MAT", "Tiền mặt"],
  ["CHUYEN_KHOAN", "Chuyển khoản"]
];

SCHEMA[SHEET_NAMES.BAO_HANH] = [
  ["MA_BH", "Mã bảo hành"],
  ["NGAY_NHAN", "Ngày nhận"],
  ["MA_KH", "Mã khách hàng"],
  ["TEN_KH", "Tên khách hàng"],
  ["TEN_SP", "Tên sản phẩm"],
  ["TINH_TRANG_LOI", "Tình trạng lỗi"],
  ["LOAI_DICH_VU", "Loại dịch vụ"],
  ["PHI_SUA_CHUA", "Phí sửa chữa"],
  ["HINH_THUC_TT", "Hình thức thanh toán"],
  ["NGUOI_TIEP_NHAN", "Người tiếp nhận"],
  ["NGUOI_SUA", "Người sửa"],
  ["TRANG_THAI", "Trạng thái"],
  ["GHI_CHU", "Ghi chú"],
  ["CHI_NHANH", "Chi nhánh"],
  ["TIEN_MAT", "Tiền mặt"],
  ["CHUYEN_KHOAN", "Chuyển khoản"]
];

// Helper to construct default enum structure from schema definition (1-indexed)
function getDefaultSchemaIndices(sheetName) {
  const indices = {};
  const schemaList = SCHEMA[sheetName];
  if (schemaList) {
    schemaList.forEach(function (colDef, i) {
      indices[colDef[0]] = i + 1;
    });
  }
  return indices;
}

// Populate the SHEET_HEADERS object on script load
function populateSheetHeaders() {
  for (const sheetName in SCHEMA) {
    SHEET_HEADERS[sheetName] = SCHEMA[sheetName].map(function (colDef) {
      return colDef[1];
    });
  }
}

var _columnEnumsInitialized = false;
var _loadingColumnEnums = false;

// Global proxy generator for column enums
function createEnumProxy(defaultValues) {
  return new Proxy(defaultValues, {
    get(target, prop) {
      if (typeof prop === "string" && !prop.startsWith("_") && prop !== "toString" && prop !== "toJSON") {
        initializeColumnEnums();
      }
      return target[prop];
    },
    set(target, prop, value) {
      target[prop] = value;
      return true;
    },
    ownKeys(target) {
      initializeColumnEnums();
      return Reflect.ownKeys(target);
    },
    getOwnPropertyDescriptor(target, prop) {
      initializeColumnEnums();
      return Reflect.getOwnPropertyDescriptor(target, prop);
    }
  });
}

// Declare enums dynamically as Proxies wrapping the default mapping indices
const COL_DT = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.DIEN_THOAI));
const COL_PK = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.PHU_KIEN));
const COL_DH = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.DON_HANG));
const COL_TM = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.THU_MUA));
const COL_NK = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.NHAP_KHO));
const COL_TG = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.TRA_GOP));
const COL_LSTG = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.LICH_SU_TRA_GOP));
const COL_DV = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.DICH_VU));
const COL_DT_TRA = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.DOI_TRA));
const COL_KH = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.KHACH_HANG));
const COL_BH = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.BAO_HANH));
const COL_NV = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.NHAN_VIEN));
const COL_DS = createEnumProxy(getDefaultSchemaIndices(SHEET_NAMES.DOANH_SO));

populateSheetHeaders();
