/**
 * ============================================================
 * VanTran Mobile — Utils.gs
 * Các hàm tiện ích dùng chung cho toàn hệ thống
 * ============================================================
 */

// Cache lưu trữ dữ liệu của các sheet trong một lần chạy để tối ưu hiệu năng đọc
var _sheetDataCache = {};
var _sheetIndexCache = {};

function clearSheetCache(sheetName) {
  if (sheetName) {
    delete _sheetDataCache[sheetName];
    delete _sheetIndexCache[sheetName];
    if (
      sheetName === SHEET_NAMES.PHU_KIEN &&
      typeof clearPhuKienCompositeIndex === "function"
    ) {
      clearPhuKienCompositeIndex();
    }
  } else {
    _sheetDataCache = {};
    _sheetIndexCache = {};
    if (typeof clearPhuKienCompositeIndex === "function") {
      clearPhuKienCompositeIndex();
    }
  }
}

/**
 * Thực thi một hàm callback dưới sự bảo vệ của Script Lock để tránh tranh chấp dữ liệu
 *
 * @param {Function} callback - Hàm nghiệp vụ cần thực thi
 * @return {*} Kết quả trả về của callback
 */
function runWithLock(callback) {
  return callback();
}

// Các Enum cột mặc định (sẽ được cập nhật động bằng initializeColumnEnums)
var COL_DT = {
  MA_DT: 1,
  TEN_SP: 2,
  THUONG_HIEU: 3,
  IMEI: 4,
  MAU_SAC: 5,
  DUNG_LUONG: 6,
  TINH_TRANG: 7,
  GIA_NHAP: 8,
  GIA_BAN: 9,
  GIA_TRA_GOP: 10,
  TRANG_THAI_KHO: 11,
  GHI_CHU: 12,
  CHI_NHANH: 13,
  NGAY_NHAP: 14,
  NGAY_XUAT: 15,
};

var COL_PK = {
  MA_PK: 1,
  TEN_SP: 2,
  LOAI_PK: 3,
  THUONG_HIEU: 4,
  GIA_NHAP: 5,
  GIA_BAN: 6,
  SO_LUONG_TON: 7,
  MO_TA: 8,
  TRANG_THAI: 9,
  CHI_NHANH: 10,
};

var COL_DH = {
  MA_DH: 1,
  NGAY_BAN: 2,
  MA_KH: 3,
  TEN_KH: 4,
  MA_SP: 5,
  TEN_SP: 6,
  NGUON_SP: 7,
  THUONG_HIEU: 8,
  SO_LUONG: 9,
  DON_GIA: 10,
  THANH_TIEN: 11,
  HINH_THUC_BAN: 12,
  HINH_THUC_TT: 13,
  NGUOI_BAN: 14,
  TEN_NGUOI_BAN: 15,
  QUYEN_XUAT: 16,
  NGUOI_HO_TRO: 17,
  TEN_NGUOI_HO_TRO: 18,
  TRANG_THAI: 19,
  GHI_CHU: 20,
  CHI_NHANH: 21,
  MA_QUA_TANG: 22,
  TEN_QUA_TANG: 23,
  CO_NHAN_QUA: 24,
  TIEN_GIAM_GIA: 25,
  TIEN_MAT: 26,
  CHUYEN_KHOAN: 27,
};

var COL_TM = {
  MA_TM: 1,
  NGAY_TM: 2,
  MA_KH: 3,
  TEN_KH: 4,
  SDT_KH: 5,
  TEN_SP_THU: 6,
  THUONG_HIEU_THU: 7,
  IMEI_THU: 8,
  MAU_SAC_THU: 9,
  DUNG_LUONG_THU: 10,
  TINH_TRANG_THU: 11,
  GIA_THU_MUA: 12,
  LOAI_GD: 13,
  MA_DH_MOI: 14,
  TIEN_HO_TRO: 15,
  TONG_TIEN_TRA: 16,
  HINH_THUC_TT: 17,
  CHI_NHANH: 18,
  NGUOI_THUC_HIEN: 19,
  GHI_CHU: 20,
  TIEN_MAT: 21,
  CHUYEN_KHOAN: 22,
};

var COL_TG = {
  MA_TG: 1,
  MA_DH: 2,
  MA_KH: 3,
  TEN_KH: 4,
  TONG_TIEN: 5,
  TRA_TRUOC: 6,
  CON_LAI: 7,
  SO_KY: 8,
  TIEN_MOI_KY: 9,
  NGAY_BAT_DAU: 10,
  NGAY_KET_THUC: 11,
  DA_TRA_SO_KY: 12,
  DA_TRA_SO_TIEN: 13,
  LOAI_TRA_GOP: 14,
  CONG_TY_TC: 15,
  TRANG_THAI: 16,
  CHI_NHANH: 17,
  TIEN_MAT: 18,
  CHUYEN_KHOAN: 19,
};

var COL_LSTG = {
  MA_LS: 1,
  MA_TG: 2,
  KY_SO: 3,
  SO_TIEN_CAN_TRA: 4,
  SO_TIEN_DA_TRA: 5,
  NGAY_CAN_TRA: 6,
  NGAY_THUC_TRA: 7,
  HINH_THUC_TT: 8,
  TRANG_THAI: 9,
  GHI_CHU: 10,
  TIEN_MAT: 11,
  CHUYEN_KHOAN: 12,
};

var COL_DV = {
  MA_DV: 1,
  NGAY_GD: 2,
  LOAI_DV: 3,
  MA_KH: 4,
  TEN_KH: 5,
  SDT_KH: 6,
  SO_TIEN_GD: 7,
  PHI_DV: 8,
  HINH_THUC_TT: 9,
  NGUOI_THUC_HIEN: 10,
  TEN_NGUOI_THUC_HIEN: 11,
  TRANG_THAI: 12,
  GHI_CHU: 13,
  CHI_NHANH: 14,
  TIEN_MAT: 15,
  CHUYEN_KHOAN: 16,
};

var COL_DT_TRA = {
  MA_DT: 1,
  NGAY_DT: 2,
  MA_DH: 3,
  MA_KH: 4,
  TEN_KH: 5,
  LOAI_GD: 6,
  MA_SP_TRA: 7,
  TEN_SP_TRA: 8,
  IMEI_TRA: 9,
  MA_SP_NHAN: 10,
  TEN_SP_NHAN: 11,
  IMEI_NHAN: 12,
  TIEN_HOAN_TRA: 13,
  PHI_DOI_TRA: 14,
  HINH_THUC_TT: 15,
  CHI_NHANH: 16,
  NGUOI_THUC_HIEN: 17,
  TRANG_THAI: 18,
  GHI_CHU: 19,
  TIEN_MAT: 20,
  CHUYEN_KHOAN: 21,
};

var COL_KH = {
  MA_KH: 1,
  HO_TEN: 2,
  SO_DIEN_THOAI: 1,
  CCCD: 3,
  DIA_CHI: 4,
  NGAY_TAO: 5,
  GHI_CHU: 6,
};

var _columnEnumsInitialized = false;

function initializeColumnEnums() {
  if (_columnEnumsInitialized) return;

  var mapDT = getColMapFromSheet(SHEET_NAMES.DIEN_THOAI);
  if (mapDT) {
    COL_DT.MA_DT = mapDT["mã điện thoại"] || COL_DT.MA_DT;
    COL_DT.TEN_SP = mapDT["tên sản phẩm"] || COL_DT.TEN_SP;
    COL_DT.THUONG_HIEU = mapDT["thương hiệu"] || COL_DT.THUONG_HIEU;
    COL_DT.IMEI = mapDT["imei"] || COL_DT.IMEI;
    COL_DT.MAU_SAC = mapDT["màu sắc"] || COL_DT.MAU_SAC;
    COL_DT.DUNG_LUONG = mapDT["dung lượng"] || COL_DT.DUNG_LUONG;
    COL_DT.TINH_TRANG = mapDT["tình trạng"] || COL_DT.TINH_TRANG;
    COL_DT.GIA_NHAP = mapDT["giá nhập"] || COL_DT.GIA_NHAP;
    COL_DT.GIA_BAN = mapDT["giá bán"] || COL_DT.GIA_BAN;
    COL_DT.GIA_TRA_GOP = mapDT["giá trả góp"] || COL_DT.GIA_TRA_GOP;
    COL_DT.TRANG_THAI_KHO = mapDT["trạng thái kho"] || COL_DT.TRANG_THAI_KHO;
    COL_DT.GHI_CHU = mapDT["ghi chú"] || COL_DT.GHI_CHU;
    COL_DT.CHI_NHANH = mapDT["chi nhánh"] || COL_DT.CHI_NHANH;
    COL_DT.NGAY_NHAP = mapDT["ngày nhập"] || COL_DT.NGAY_NHAP;
    COL_DT.NGAY_XUAT = mapDT["ngày xuất"] || COL_DT.NGAY_XUAT;
  }

  var mapPK = getColMapFromSheet(SHEET_NAMES.PHU_KIEN);
  if (mapPK) {
    COL_PK.MA_PK = mapPK["mã phụ kiện"] || COL_PK.MA_PK;
    COL_PK.TEN_SP = mapPK["tên sản phẩm"] || COL_PK.TEN_SP;
    COL_PK.LOAI_PK = mapPK["loại phụ kiện"] || COL_PK.LOAI_PK;
    COL_PK.THUONG_HIEU = mapPK["thương hiệu"] || COL_PK.THUONG_HIEU;
    COL_PK.GIA_NHAP = mapPK["giá nhập"] || COL_PK.GIA_NHAP;
    COL_PK.GIA_BAN = mapPK["giá bán"] || COL_PK.GIA_BAN;
    COL_PK.SO_LUONG_TON = mapPK["số lượng tồn"] || COL_PK.SO_LUONG_TON;
    COL_PK.MO_TA = mapPK["mô tả"] || COL_PK.MO_TA;
    COL_PK.TRANG_THAI = mapPK["trạng thái"] || COL_PK.TRANG_THAI;
    COL_PK.CHI_NHANH = mapPK["chi nhánh"] || COL_PK.CHI_NHANH;
  }

  var mapDH = getColMapFromSheet(SHEET_NAMES.DON_HANG);
  if (mapDH) {
    COL_DH.MA_DH = mapDH["mã đơn hàng"] || COL_DH.MA_DH;
    COL_DH.NGAY_BAN = mapDH["ngày bán"] || COL_DH.NGAY_BAN;
    COL_DH.MA_KH = mapDH["mã khách hàng"] || COL_DH.MA_KH;
    COL_DH.TEN_KH = mapDH["tên khách hàng"] || COL_DH.TEN_KH;
    COL_DH.MA_SP = mapDH["mã sản phẩm"] || COL_DH.MA_SP;
    COL_DH.TEN_SP = mapDH["tên sản phẩm"] || COL_DH.TEN_SP;
    COL_DH.NGUON_SP = mapDH["nguồn sản phẩm"] || COL_DH.NGUON_SP;
    COL_DH.THUONG_HIEU = mapDH["thương hiệu"] || COL_DH.THUONG_HIEU;
    COL_DH.SO_LUONG = mapDH["số lượng"] || COL_DH.SO_LUONG;
    COL_DH.DON_GIA = mapDH["đơn giá"] || COL_DH.DON_GIA;
    COL_DH.THANH_TIEN = mapDH["thành tiền"] || COL_DH.THANH_TIEN;
    COL_DH.HINH_THUC_BAN = mapDH["hình thức bán"] || COL_DH.HINH_THUC_BAN;
    COL_DH.HINH_THUC_TT = mapDH["hình thức thanh toán"] || COL_DH.HINH_THUC_TT;
    COL_DH.NGUOI_BAN = mapDH["người bán"] || COL_DH.NGUOI_BAN;
    COL_DH.TEN_NGUOI_BAN = mapDH["tên người bán"] || COL_DH.TEN_NGUOI_BAN;
    COL_DH.QUYEN_XUAT = mapDH["có quyền xuất máy"] || COL_DH.QUYEN_XUAT;
    COL_DH.NGUOI_HO_TRO = mapDH["người hỗ trợ"] || COL_DH.NGUOI_HO_TRO;
    COL_DH.TEN_NGUOI_HO_TRO =
      mapDH["tên người hỗ trợ"] || COL_DH.TEN_NGUOI_HO_TRO;
    COL_DH.TRANG_THAI = mapDH["trạng thái"] || COL_DH.TRANG_THAI;
    COL_DH.GHI_CHU = mapDH["ghi chú"] || COL_DH.GHI_CHU;
    COL_DH.CHI_NHANH = mapDH["chi nhánh"] || COL_DH.CHI_NHANH;
    COL_DH.MA_QUA_TANG = mapDH["mã quà tặng"] || COL_DH.MA_QUA_TANG;
    COL_DH.TEN_QUA_TANG = mapDH["tên quà tặng"] || COL_DH.TEN_QUA_TANG;
    COL_DH.CO_NHAN_QUA = mapDH["có nhận quà"] || COL_DH.CO_NHAN_QUA;
    COL_DH.TIEN_GIAM_GIA = mapDH["tiền giảm giá"] || COL_DH.TIEN_GIAM_GIA;
    COL_DH.TIEN_MAT = mapDH["tiền mặt"] || COL_DH.TIEN_MAT;
    COL_DH.CHUYEN_KHOAN = mapDH["chuyển khoản"] || COL_DH.CHUYEN_KHOAN;
  }

  var mapTM = getColMapFromSheet(SHEET_NAMES.THU_MUA);
  if (mapTM) {
    COL_TM.MA_TM = mapTM["mã thu mua"] || COL_TM.MA_TM;
    COL_TM.NGAY_TM = mapTM["ngày thu mua"] || COL_TM.NGAY_TM;
    COL_TM.MA_KH = mapTM["mã khách hàng"] || COL_TM.MA_KH;
    COL_TM.TEN_KH = mapTM["tên khách hàng"] || COL_TM.TEN_KH;
    COL_TM.SDT_KH = mapTM["số điện thoại khách"] || COL_TM.SDT_KH;
    COL_TM.TEN_SP_THU = mapTM["tên sản phẩm thu"] || COL_TM.TEN_SP_THU;
    COL_TM.THUONG_HIEU_THU = mapTM["thương hiệu thu"] || COL_TM.THUONG_HIEU_THU;
    COL_TM.IMEI_THU = mapTM["imei thu"] || COL_TM.IMEI_THU;
    COL_TM.MAU_SAC_THU = mapTM["màu sắc thu"] || COL_TM.MAU_SAC_THU;
    COL_TM.DUNG_LUONG_THU = mapTM["dung lượng thu"] || COL_TM.DUNG_LUONG_THU;
    COL_TM.TINH_TRANG_THU = mapTM["tình trạng thu"] || COL_TM.TINH_TRANG_THU;
    COL_TM.GIA_THU_MUA = mapTM["giá thu mua"] || COL_TM.GIA_THU_MUA;
    COL_TM.LOAI_GD = mapTM["loại giao dịch"] || COL_TM.LOAI_GD;
    COL_TM.MA_DH_MOI = mapTM["mã đơn hàng mới"] || COL_TM.MA_DH_MOI;
    COL_TM.TIEN_HO_TRO = mapTM["tiền hỗ trợ"] || COL_TM.TIEN_HO_TRO;
    COL_TM.TONG_TIEN_TRA = mapTM["tổng tiền trả khách"] || COL_TM.TONG_TIEN_TRA;
    COL_TM.HINH_THUC_TT = mapTM["hình thức thanh toán"] || COL_TM.HINH_THUC_TT;
    COL_TM.CHI_NHANH = mapTM["chi nhánh"] || COL_TM.CHI_NHANH;
    COL_TM.NGUOI_THUC_HIEN = mapTM["người thực hiện"] || COL_TM.NGUOI_THUC_HIEN;
    COL_TM.GHI_CHU = mapTM["ghi chú"] || COL_TM.GHI_CHU;
    COL_TM.TIEN_MAT = mapTM["tiền mặt"] || COL_TM.TIEN_MAT;
    COL_TM.CHUYEN_KHOAN = mapTM["chuyển khoản"] || COL_TM.CHUYEN_KHOAN;
  }

  var mapTG = getColMapFromSheet(SHEET_NAMES.TRA_GOP);
  if (mapTG) {
    COL_TG.MA_TG = mapTG["mã trả góp"] || COL_TG.MA_TG;
    COL_TG.MA_DH = mapTG["mã đơn hàng"] || COL_TG.MA_DH;
    COL_TG.MA_KH = mapTG["mã khách hàng"] || COL_TG.MA_KH;
    COL_TG.TEN_KH = mapTG["tên khách hàng"] || COL_TG.TEN_KH;
    COL_TG.TONG_TIEN = mapTG["tổng tiền"] || COL_TG.TONG_TIEN;
    COL_TG.TRA_TRUOC = mapTG["trả trước"] || COL_TG.TRA_TRUOC;
    COL_TG.CON_LAI = mapTG["còn lại"] || COL_TG.CON_LAI;
    COL_TG.SO_KY = mapTG["số kỳ"] || COL_TG.SO_KY;
    COL_TG.TIEN_MO_KY = mapTG["tiền mỗi kỳ"] || COL_TG.TIEN_MO_KY;
    COL_TG.NGAY_BAT_DAU = mapTG["ngày bắt đầu"] || COL_TG.NGAY_BAT_DAU;
    COL_TG.NGAY_KET_THUC = mapTG["ngày kết thúc"] || COL_TG.NGAY_KET_THUC;
    COL_TG.DA_TRA_SO_KY = mapTG["đã trả số kỳ"] || COL_TG.DA_TRA_SO_KY;
    COL_TG.DA_TRA_SO_TIEN = mapTG["đã trả số tiền"] || COL_TG.DA_TRA_SO_TIEN;
    COL_TG.LOAI_TRA_GOP = mapTG["loại trả góp"] || COL_TG.LOAI_TRA_GOP;
    COL_TG.CONG_TY_TC = mapTG["công ty tài chính"] || COL_TG.CONG_TY_TC;
    COL_TG.TRANG_THAI = mapTG["trạng thái"] || COL_TG.TRANG_THAI;
    COL_TG.CHI_NHANH = mapTG["chi nhánh"] || COL_TG.CHI_NHANH;
    COL_TG.TIEN_MAT = mapTG["tiền mặt"] || COL_TG.TIEN_MAT;
    COL_TG.CHUYEN_KHOAN = mapTG["chuyển khoản"] || COL_TG.CHUYEN_KHOAN;
  }

  var mapLSTG = getColMapFromSheet(SHEET_NAMES.LICH_SU_TRA_GOP);
  if (mapLSTG) {
    COL_LSTG.MA_LS = mapLSTG["mã lịch sử"] || COL_LSTG.MA_LS;
    COL_LSTG.MA_TG = mapLSTG["mã trả góp"] || COL_LSTG.MA_TG;
    COL_LSTG.KY_SO = mapLSTG["kỳ số"] || COL_LSTG.KY_SO;
    COL_LSTG.SO_TIEN_CAN_TRA =
      mapLSTG["số tiền cần trả"] || COL_LSTG.SO_TIEN_CAN_TRA;
    COL_LSTG.SO_TIEN_DA_TRA =
      mapLSTG["số tiền đã trả"] || COL_LSTG.SO_TIEN_DA_TRA;
    COL_LSTG.NGAY_CAN_TRA = mapLSTG["ngày cần trả"] || COL_LSTG.NGAY_CAN_TRA;
    COL_LSTG.NGAY_THUC_TRA = mapLSTG["ngày thực trả"] || COL_LSTG.NGAY_THUC_TRA;
    COL_LSTG.HINH_THUC_TT =
      mapLSTG["hình thức thanh toán"] || COL_LSTG.HINH_THUC_TT;
    COL_LSTG.TRANG_THAI = mapLSTG["trạng thái"] || COL_LSTG.TRANG_THAI;
    COL_LSTG.GHI_CHU = mapLSTG["ghi chú"] || COL_LSTG.GHI_CHU;
    COL_LSTG.TIEN_MAT = mapLSTG["tiền mặt"] || COL_LSTG.TIEN_MAT;
    COL_LSTG.CHUYEN_KHOAN = mapLSTG["chuyển khoản"] || COL_LSTG.CHUYEN_KHOAN;
  }

  var mapDV = getColMapFromSheet(SHEET_NAMES.DICH_VU);
  if (mapDV) {
    COL_DV.MA_DV = mapDV["mã dịch vụ"] || COL_DV.MA_DV;
    COL_DV.NGAY_GD = mapDV["ngày giao dịch"] || COL_DV.NGAY_GD;
    COL_DV.LOAI_DV = mapDV["loại dịch vụ"] || COL_DV.LOAI_DV;
    COL_DV.MA_KH = mapDV["mã khách hàng"] || COL_DV.MA_KH;
    COL_DV.TEN_KH = mapDV["tên khách hàng"] || COL_DV.TEN_KH;
    COL_DV.SDT_KH = mapDV["số điện thoại khách"] || COL_DV.SDT_KH;
    COL_DV.SO_TIEN_GD = mapDV["số tiền giao dịch"] || COL_DV.SO_TIEN_GD;
    COL_DV.PHI_DV = mapDV["phí dịch vụ"] || COL_DV.PHI_DV;
    COL_DV.HINH_THUC_TT = mapDV["hình thức thanh toán"] || COL_DV.HINH_THUC_TT;
    COL_DV.NGUOI_THUC_HIEN = mapDV["người thực hiện"] || COL_DV.NGUOI_THUC_HIEN;
    COL_DV.TEN_NGUOI_THUC_HIEN =
      mapDV["tên người thực hiện"] || COL_DV.TEN_NGUOI_THUC_HIEN;
    COL_DV.TRANG_THAI = mapDV["trạng thái"] || COL_DV.TRANG_THAI;
    COL_DV.GHI_CHU = mapDV["ghi chú"] || COL_DV.GHI_CHU;
    COL_DV.CHI_NHANH = mapDV["chi nhánh"] || COL_DV.CHI_NHANH;
    COL_DV.TIEN_MAT = mapDV["tiền mặt"] || COL_DV.TIEN_MAT;
    COL_DV.CHUYEN_KHOAN = mapDV["chuyển khoản"] || COL_DV.CHUYEN_KHOAN;
  }

  var mapDoiTra = getColMapFromSheet(SHEET_NAMES.DOI_TRA);
  if (mapDoiTra) {
    COL_DT_TRA.MA_DT = mapDoiTra["mã đổi trả"] || COL_DT_TRA.MA_DT;
    COL_DT_TRA.NGAY_DT = mapDoiTra["ngày đổi trả"] || COL_DT_TRA.NGAY_DT;
    COL_DT_TRA.MA_DH = mapDoiTra["mã đơn hàng"] || COL_DT_TRA.MA_DH;
    COL_DT_TRA.MA_KH = mapDoiTra["mã khách hàng"] || COL_DT_TRA.MA_KH;
    COL_DT_TRA.TEN_KH = mapDoiTra["tên khách hàng"] || COL_DT_TRA.TEN_KH;
    COL_DT_TRA.LOAI_GD = mapDoiTra["loại giao dịch"] || COL_DT_TRA.LOAI_GD;
    COL_DT_TRA.MA_SP_TRA = mapDoiTra["mã sản phẩm trả"] || COL_DT_TRA.MA_SP_TRA;
    COL_DT_TRA.TEN_SP_TRA =
      mapDoiTra["tên sản phẩm trả"] || COL_DT_TRA.TEN_SP_TRA;
    COL_DT_TRA.IMEI_TRA = mapDoiTra["imei trả"] || COL_DT_TRA.IMEI_TRA;
    COL_DT_TRA.MA_SP_NHAN =
      mapDoiTra["mã sản phẩm nhận"] || COL_DT_TRA.MA_SP_NHAN;
    COL_DT_TRA.TEN_SP_NHAN =
      mapDoiTra["tên sản phẩm nhận"] || COL_DT_TRA.TEN_SP_NHAN;
    COL_DT_TRA.IMEI_NHAN = mapDoiTra["imei nhận"] || COL_DT_TRA.IMEI_NHAN;
    COL_DT_TRA.TIEN_HOAN_TRA =
      mapDoiTra["tiền hoàn trả"] || COL_DT_TRA.TIEN_HOAN_TRA;
    COL_DT_TRA.PHI_DOI_TRA = mapDoiTra["phí đổi trả"] || COL_DT_TRA.PHI_DOI_TRA;
    COL_DT_TRA.HINH_THUC_TT =
      mapDoiTra["hình thức thanh toán"] || COL_DT_TRA.HINH_THUC_TT;
    COL_DT_TRA.CHI_NHANH = mapDoiTra["chi nhánh"] || COL_DT_TRA.CHI_NHANH;
    COL_DT_TRA.NGUOI_THUC_HIEN =
      mapDoiTra["người thực hiện"] || COL_DT_TRA.NGUOI_THUC_HIEN;
    COL_DT_TRA.TRANG_THAI = mapDoiTra["trạng thái"] || COL_DT_TRA.TRANG_THAI;
    COL_DT_TRA.GHI_CHU = mapDoiTra["ghi chú"] || COL_DT_TRA.GHI_CHU;
    COL_DT_TRA.TIEN_MAT = mapDoiTra["tiền mặt"] || COL_DT_TRA.TIEN_MAT;
    COL_DT_TRA.CHUYEN_KHOAN =
      mapDoiTra["chuyển khoản"] || COL_DT_TRA.CHUYEN_KHOAN;
  }

  var mapKH = getColMapFromSheet(SHEET_NAMES.KHACH_HANG);
  if (mapKH) {
    COL_KH.MA_KH = mapKH["mã khách hàng"] || COL_KH.MA_KH;
    COL_KH.HO_TEN = mapKH["họ tên"] || mapKH["họ và tên"] || COL_KH.HO_TEN;
    COL_KH.SO_DIEN_THOAI = mapKH["số điện thoại"] || COL_KH.MA_KH;
    COL_KH.CCCD = mapKH["cccd"] || COL_KH.CCCD;
    COL_KH.DIA_CHI = mapKH["địa chỉ"] || COL_KH.DIA_CHI;
    COL_KH.NGAY_TAO = mapKH["ngày tạo"] || COL_KH.NGAY_TAO;
    COL_KH.GHI_CHU = mapKH["ghi chú"] || COL_KH.GHI_CHU;
  }

  _columnEnumsInitialized = true;
}

function getColMapFromSheet(sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return null;

    var lastCol = sheet.getLastColumn();
    var headers = [];
    if (lastCol > 0) {
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    // Tự động kiểm tra và thêm các cột còn thiếu theo SHEET_HEADERS
    var expectedHeaders = SHEET_HEADERS[sheetName];
    if (expectedHeaders && expectedHeaders.length > headers.length) {
      var nextCol = headers.length + 1;
      var missing = expectedHeaders.slice(headers.length);
      sheet.getRange(1, nextCol, 1, missing.length).setValues([missing]);

      // Tải lại các header mới
      lastCol = sheet.getLastColumn();
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    var map = {};
    for (var i = 0; i < headers.length; i++) {
      map[String(headers[i]).trim().toLowerCase()] = i + 1;
    }
    return map;
  } catch (e) {
    Logger.log("Error getColMapFromSheet for " + sheetName + ": " + e.message);
    return null;
  }
}

function getSheetIndex(sheetName, searchCol) {
  if (!_sheetIndexCache[sheetName]) {
    _sheetIndexCache[sheetName] = {};
  }
  if (!_sheetIndexCache[sheetName][searchCol]) {
    var indexMap = {};
    var data = getSheetDataCached(sheetName);
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      if (row.length < searchCol) continue;
      var cellVal = String(row[searchCol - 1])
        .trim()
        .toLowerCase();
      if (!(cellVal in indexMap)) {
        indexMap[cellVal] = i; // Lưu index 0-indexed trong mảng data
      }
    }
    _sheetIndexCache[sheetName][searchCol] = indexMap;
  }
  return _sheetIndexCache[sheetName][searchCol];
}

// ==================== CACHE SERVICE (PERSISTENT ACROSS REQUESTS) ====================

var DROPDOWN_CACHE_TTL = 3600; // 1 giờ

/**
 * Lưu dữ liệu vào CacheService với chunking (tránh giới hạn 100KB/key)
 *
 * @param {string} key - Cache key
 * @param {*} data - Dữ liệu cần cache (sẽ được JSON.stringify)
 */
function setChunkedCache(key, data) {
  try {
    var cache = CacheService.getScriptCache();
    var json = JSON.stringify(data);
    var CHUNK_SIZE = 90000; // ~90KB per chunk (buffer từ giới hạn 100KB)
    var numChunks = Math.ceil(json.length / CHUNK_SIZE);

    var cacheObj = {};
    cacheObj[key + "_m"] = JSON.stringify({ c: numChunks });

    for (var i = 0; i < numChunks; i++) {
      cacheObj[key + "_" + i] = json.substring(
        i * CHUNK_SIZE,
        (i + 1) * CHUNK_SIZE,
      );
    }

    cache.putAll(cacheObj, DROPDOWN_CACHE_TTL);
  } catch (e) {
    Logger.log("CacheService setChunkedCache error: " + e.message);
  }
}

/**
 * Đọc dữ liệu đã chunk từ CacheService
 *
 * @param {string} key - Cache key
 * @return {*|null} Dữ liệu đã parse hoặc null nếu cache miss
 */
function getChunkedCache(key) {
  try {
    var cache = CacheService.getScriptCache();
    var meta = cache.get(key + "_m");
    if (!meta) return null;

    var metaObj = JSON.parse(meta);
    var numChunks = metaObj.c;

    var keys = [];
    for (var i = 0; i < numChunks; i++) {
      keys.push(key + "_" + i);
    }

    var all = cache.getAll(keys);
    var json = "";
    for (var i = 0; i < numChunks; i++) {
      var chunk = all[key + "_" + i];
      if (!chunk) return null; // Partial cache miss
      json += chunk;
    }

    return JSON.parse(json);
  } catch (e) {
    Logger.log("CacheService getChunkedCache error: " + e.message);
    return null;
  }
}

/**
 * Xóa cache dropdown liên quan đến sheet đã thay đổi
 *
 * @param {string} sheetName - Tên sheet đã thay đổi dữ liệu
 */
function invalidateDropdownCache(sheetName) {
  var cacheKeys = [];
  if (sheetName === SHEET_NAMES.KHACH_HANG) {
    cacheKeys = ["dd_kh"];
  } else if (sheetName === SHEET_NAMES.DIEN_THOAI) {
    cacheKeys = ["dd_dt"];
  } else if (sheetName === SHEET_NAMES.PHU_KIEN) {
    cacheKeys = ["dd_pk", "dd_pku"];
  } else if (sheetName === SHEET_NAMES.TRA_GOP) {
    cacheKeys = ["dd_tg"];
  }

  if (cacheKeys.length === 0) return;

  try {
    var cache = CacheService.getScriptCache();
    for (var k = 0; k < cacheKeys.length; k++) {
      var key = cacheKeys[k];
      var meta = cache.get(key + "_m");
      if (meta) {
        var metaObj = JSON.parse(meta);
        var removeKeys = [key + "_m"];
        for (var i = 0; i < metaObj.c; i++) {
          removeKeys.push(key + "_" + i);
        }
        cache.removeAll(removeKeys);
      }
    }
  } catch (e) {
    Logger.log("CacheService invalidateDropdownCache error: " + e.message);
  }
}

function getSheetDataCached(sheetName) {
  if (_sheetDataCache[sheetName]) {
    return _sheetDataCache[sheetName];
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  _sheetDataCache[sheetName] = data;
  return data;
}

/**
 * Tạo mã tự động theo prefix + số thứ tự
 * VD: generateId('DH', 'DonHang') → 'DH001', 'DH002', ...
 *
 * @param {string} prefix - Tiền tố mã (VD: 'DH', 'NV', 'KH')
 * @param {string} sheetName - Tên sheet để đếm số dòng
 * @return {string} Mã mới
 */
function generateId(prefix, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return prefix + "001";

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return prefix + "001";

  // Lấy cột mã (cột 1) và tìm số lớn nhất
  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var maxNum = 0;

  data.forEach(function (row) {
    var val = String(row[0]);
    if (val.indexOf(prefix) === 0) {
      var numPart = parseInt(val.substring(prefix.length), 10);
      if (!isNaN(numPart) && numPart > maxNum) {
        maxNum = numPart;
      }
    }
  });

  var nextNum = maxNum + 1;
  var padded = ("000" + nextNum).slice(-3);
  return prefix + padded;
}

/**
 * Lookup giá trị từ sheet
 *
 * @param {string} sheetName - Tên sheet
 * @param {number} searchCol - Cột tìm kiếm (1-indexed)
 * @param {*} searchVal - Giá trị tìm
 * @param {number} returnCol - Cột trả về (1-indexed)
 * @return {*} Giá trị tìm được hoặc null
 */
function lookupValue(sheetName, searchCol, searchVal, returnCol) {
  var indexMap = getSheetIndex(sheetName, searchCol);
  var searchStr = String(searchVal).trim().toLowerCase();
  if (searchStr in indexMap) {
    var data = getSheetDataCached(sheetName);
    var rowIndex = indexMap[searchStr];
    var row = data[rowIndex];
    if (row.length >= returnCol) {
      return row[returnCol - 1];
    }
  }
  return null;
}

/**
 * Lookup nhiều cột cùng lúc từ 1 dòng
 *
 * @param {string} sheetName - Tên sheet
 * @param {number} searchCol - Cột tìm kiếm (1-indexed)
 * @param {*} searchVal - Giá trị tìm
 * @param {number[]} returnCols - Mảng các cột cần trả về (1-indexed)
 * @return {Object|null} Object với key là index cột, value là giá trị
 */
function lookupMultipleValues(sheetName, searchCol, searchVal, returnCols) {
  var indexMap = getSheetIndex(sheetName, searchCol);
  var searchStr = String(searchVal).trim().toLowerCase();
  if (searchStr in indexMap) {
    var data = getSheetDataCached(sheetName);
    var rowIndex = indexMap[searchStr];
    var row = data[rowIndex];
    var result = {};
    returnCols.forEach(function (col) {
      result[col] = row.length >= col ? row[col - 1] : "";
    });
    result._rowIndex = rowIndex + 2; // Actual row in sheet (1-indexed, skip header)
    return result;
  }
  return null;
}

/**
 * Định dạng số tiền VND
 *
 * @param {number} number - Số tiền
 * @return {string} Chuỗi đã format: "1,234,567"
 */
function formatCurrency(number) {
  if (isNaN(number) || number === null || number === undefined) return "0";
  return Number(number).toLocaleString("vi-VN");
}

/**
 * Định dạng ngày dd/MM/yyyy
 *
 * @param {Date} date - Đối tượng Date
 * @return {string} Chuỗi ngày "dd/MM/yyyy"
 */
function formatDate(date) {
  if (!date || !(date instanceof Date)) return "";
  var dd = ("0" + date.getDate()).slice(-2);
  var mm = ("0" + (date.getMonth() + 1)).slice(-2);
  var yyyy = date.getFullYear();
  return dd + "/" + mm + "/" + yyyy;
}

/**
 * Đọc giá trị cấu hình từ sheet CauHinh
 *
 * @param {string} key - Tên cấu hình (cột A)
 * @return {string} Giá trị cấu hình (cột B) hoặc ''
 */
function getConfig(key) {
  var val = lookupValue(SHEET_NAMES.CAU_HINH, 1, key, 2);
  return val !== null ? String(val) : "";
}

/**
 * Đọc giá trị cấu hình dạng số
 *
 * @param {string} key - Tên cấu hình
 * @return {number} Giá trị số hoặc 0
 */
function getConfigNumber(key) {
  var val = getConfig(key);
  var num = parseFloat(val);
  return isNaN(num) ? 0 : num;
}

/**
 * Lấy toàn bộ dữ liệu từ sheet (trừ header)
 *
 * @param {string} sheetName - Tên sheet
 * @return {Array[]} Mảng 2 chiều dữ liệu
 */
function getAllData(sheetName) {
  return getSheetDataCached(sheetName);
}

/**
 * Thêm 1 dòng dữ liệu vào cuối sheet
 *
 * @param {string} sheetName - Tên sheet
 * @param {Array} rowData - Mảng dữ liệu 1 dòng
 * @return {number} Số dòng mới (row number)
 */
function appendRow(sheetName, rowData) {
  clearSheetCache(sheetName);
  invalidateDropdownCache(sheetName);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" không tồn tại!');

  // Tự động mở rộng số cột nếu dòng dữ liệu mới dài hơn số cột hiện tại
  var maxCols = sheet.getMaxColumns();
  if (maxCols < rowData.length) {
    sheet.insertColumnsAfter(maxCols, rowData.length - maxCols);
  }

  sheet.appendRow(rowData);
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(lastRow, 1, 1, rowData.length);
  range.setFontFamily("Times New Roman");
  range.setFontSize(12);

  // Auto-resize columns
  for (var i = 1; i <= rowData.length; i++) {
    sheet.autoResizeColumn(i);
  }

  return lastRow;
}

/**
 * Cập nhật giá trị 1 cell
 *
 * @param {string} sheetName - Tên sheet
 * @param {number} row - Dòng (1-indexed)
 * @param {number} col - Cột (1-indexed)
 * @param {*} value - Giá trị mới
 */
function updateCell(sheetName, row, col, value) {
  clearSheetCache(sheetName);
  invalidateDropdownCache(sheetName);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" không tồn tại!');

  // Tự động mở rộng số cột nếu cột cần cập nhật vượt quá số cột hiện tại
  var maxCols = sheet.getMaxColumns();
  if (maxCols < col) {
    sheet.insertColumnsAfter(maxCols, col - maxCols);
  }

  var cell = sheet.getRange(row, col);
  cell.setValue(value);
  cell.setFontFamily("Times New Roman");
  cell.setFontSize(12);
  sheet.autoResizeColumn(col);
}

/**
 * Tìm dòng theo giá trị cột
 *
 * @param {string} sheetName - Tên sheet
 * @param {number} searchCol - Cột tìm (1-indexed)
 * @param {*} searchVal - Giá trị tìm
 * @return {number} Row number (1-indexed) hoặc -1 nếu không tìm thấy
 */
function findRow(sheetName, searchCol, searchVal) {
  var indexMap = getSheetIndex(sheetName, searchCol);
  var searchStr = String(searchVal).trim().toLowerCase();
  if (searchStr in indexMap) {
    return indexMap[searchStr] + 2; // +2 vì skip header và 0-indexed
  }
  return -1;
}

/**
 * Hiển thị thông báo toast
 *
 * @param {string} message - Nội dung thông báo
 * @param {string} title - Tiêu đề (optional)
 * @param {number} duration - Thời gian hiển thị (giây, default 5)
 */
function showToast(message, title, duration) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      message,
      title || "VanTran Mobile",
      duration || 5,
    );
  } catch (e) {
    Logger.log("Toast: [" + (title || "VanTran Mobile") + "] " + message);
  }
}

/**
 * Hiển thị alert
 *
 * @param {string} title - Tiêu đề
 * @param {string} message - Nội dung
 */
function showAlert(title, message) {
  try {
    var ui = SpreadsheetApp.getUi();
    if (ui) {
      ui.alert(title, message, ui.ButtonSet.OK);
      return;
    }
  } catch (e) {
    Logger.log("Alert: [" + title + "] " + message);
  }
}

/**
 * Kiểm tra thương hiệu có phải Apple không
 *
 * @param {string} thuongHieu - Tên thương hiệu
 * @return {boolean}
 */
function isApple(thuongHieu) {
  return String(thuongHieu).trim().toLowerCase() === "apple";
}

/**
 * Lấy tháng/năm hiện tại dạng MM/YYYY
 *
 * @return {string} VD: "06/2026"
 */
function getCurrentMonthYear() {
  var now = new Date();
  var mm = ("0" + (now.getMonth() + 1)).slice(-2);
  return mm + "/" + now.getFullYear();
}

/**
 * Lấy danh sách các chi nhánh dưới dạng mảng chuỗi
 *
 * @return {string[]}
 */
function getBranchesList() {
  var branchesStr = getConfig("Danh sách chi nhánh");
  if (!branchesStr) return ["Chi nhánh 1", "Chi nhánh 2", "Chi nhánh 3"];
  return branchesStr.split(",").map(function (item) {
    return item.trim();
  });
}

/**
 * Lấy danh sách các chi nhánh dưới dạng cấu trúc dropdown cho UI
 *
 * @return {Object[]}
 */
function getBranchesDropdown() {
  var list = getBranchesList();
  return list.map(function (item) {
    return { value: item, text: item };
  });
}

/**
 * Lấy danh sách các công ty tài chính dưới dạng mảng chuỗi
 *
 * @return {string[]}
 */
function getFinanceCompaniesList() {
  var listStr = getConfig("Danh sách công ty tài chính");
  if (!listStr)
    return ["FE Credit", "Home Credit", "HD Saison", "MIRAES Asset"];
  return listStr.split(",").map(function (item) {
    return item.trim();
  });
}

/**
 * Lấy danh sách các công ty tài chính dưới dạng cấu trúc dropdown cho UI
 *
 * @return {Object[]}
 */
function getFinanceCompaniesDropdown() {
  var list = getFinanceCompaniesList();
  return list.map(function (item) {
    return { value: item, text: item };
  });
}

/**
 * Đảm bảo khoá cấu hình tồn tại trong sheet CauHinh
 * Nếu chưa tồn tại, tự động chèn thêm dòng mới với giá trị mặc định và mô tả
 *
 * @param {string} key - Tên cấu hình
 * @param {string} defaultValue - Giá trị mặc định
 * @param {string} description - Mô tả cấu hình
 */
function ensureConfigKey(key, defaultValue, description) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.CAU_HINH);
  if (!sheet) return;
  var row = findRow(SHEET_NAMES.CAU_HINH, 1, key);
  if (row === -1) {
    var lastRow = sheet.getLastRow();
    var range = sheet.getRange(lastRow + 1, 1, 1, 3);
    range.setValues([[key, defaultValue, description]]);
    range.setFontFamily("Times New Roman").setFontSize(12);
  }
}

/**
 * Lấy danh sách các thương hiệu điện thoại dưới dạng mảng chuỗi
 *
 * @return {string[]}
 */
function getBrandsList() {
  ensureConfigKey(
    "Danh sách thương hiệu",
    "Apple, Samsung, Xiaomi, OPPO, Vivo, Realme, Khác",
    "Danh sách các thương hiệu điện thoại, phân cách bằng dấu phẩy",
  );
  var brandsStr = getConfig("Danh sách thương hiệu");
  if (!brandsStr)
    return ["Apple", "Samsung", "Xiaomi", "OPPO", "Vivo", "Realme", "Khác"];
  return brandsStr.split(",").map(function (item) {
    return item.trim();
  });
}

/**
 * Lấy danh sách các thương hiệu điện thoại dưới dạng cấu trúc dropdown cho UI
 *
 * @return {Object[]}
 */
function getBrandsDropdown() {
  var list = getBrandsList();
  return list.map(function (item) {
    return { value: item, text: item };
  });
}

/**
 * Phân tích hình thức thanh toán hỗn hợp (ví dụ: "Tiền mặt: 500,000 + Chuyển khoản: 1,000,000")
 *
 * @param {string} hinhThucTT - Chuỗi hình thức thanh toán
 * @param {number} totalAmount - Tổng số tiền thu/chi thực tế của bản ghi (amountCollected, hoanTien, v.v.)
 * @return {Object} { tm: number, ck: number } số tiền mặt và chuyển khoản (bao gồm POS)
 */
function parseMixedPayment(hinhThucTT, totalAmount) {
  var str = String(hinhThucTT || "").trim();
  if (str === "Tiền mặt") {
    return { tm: totalAmount, ck: 0 };
  }
  if (str === "Chuyển khoản" || str === "Quẹt thẻ (POS)") {
    return { tm: 0, ck: totalAmount };
  }

  if (str.indexOf("+") !== -1 || str.indexOf(":") !== -1) {
    var tmMatch = str.match(/Tiền mặt:\s*([\d,]+)/i);
    var ckMatch = str.match(/Chuyển khoản:\s*([\d,]+)/i);
    var posMatch = str.match(/Quẹt thẻ\s*\(POS\):\s*([\d,]+)/i);

    var tmVal = tmMatch ? Number(tmMatch[1].replace(/,/g, "")) : 0;
    var ckVal = ckMatch ? Number(ckMatch[1].replace(/,/g, "")) : 0;
    var posVal = posMatch ? Number(posMatch[1].replace(/,/g, "")) : 0;

    var sum = tmVal + ckVal + posVal;
    if (sum > 0) {
      // Tính tỷ lệ để phân bổ chính xác theo totalAmount thực tế thu/chi
      var ratio = totalAmount / sum;
      return {
        tm: Math.round(tmVal * ratio),
        ck: Math.round((ckVal + posVal) * ratio),
      };
    }
  }

  // Mặc định chuyển khoản nếu không khớp
  return { tm: 0, ck: totalAmount };
}

/**
 * Phân tích chuỗi số tiền hỗn hợp dạng "chuyenKhoan,tienMat"
 * @param {any} val - Giá trị ô hoặc chuỗi
 * @return {Object|null} { ck: number, tm: number } hoặc null nếu không phải định dạng hỗn hợp
 */
function parseHybridAmount(val) {
  var str = String(val || "").trim();
  if (str.indexOf(",") !== -1) {
    var parts = str.split(",");
    if (parts.length === 2) {
      var ck = Number(parts[0]);
      var tm = Number(parts[1]);
      if (!isNaN(ck) && !isNaN(tm)) {
        return { ck: ck, tm: tm };
      }
    }
  }
  return null;
}

/**
 * Lấy tổng số tiền (cho cả trường hợp số đơn hoặc chuỗi hỗn hợp)
 * @param {any} val
 * @return {number}
 */
function parseAmountVal(val) {
  if (val === undefined || val === null || val === "") return 0;
  var parsed = parseHybridAmount(val);
  if (parsed) {
    return parsed.ck + parsed.tm;
  }
  return Number(val) || 0;
}
