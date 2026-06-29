/**
 * ============================================================
 * VanTran Mobile — Triggers.gs
 * Logic xử lý các triggers và sự kiện chỉnh sửa bảng tính (onEdit)
 * ============================================================
 */

// ======================== EDIT TRIGGER REGISTER ========================

const EDIT_HANDLERS = {};

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
    isRelevant = (col === COL_DH.SO_LUONG || col === COL_DH.DON_GIA || col === COL_DH.TIEN_GIAM_GIA || col === COL_DH.TIEN_THU_MUA ||
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
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  const range = sheet.getRange(row, 1, 1, lastCol);
  const v = range.getValues()[0];

  // Auto tính ThanhTien khi SoLuong hoặc DonGia hoặc TienGiamGia hoặc TienThuMua thay đổi
  if (col === COL_DH.SO_LUONG || col === COL_DH.DON_GIA || col === COL_DH.TIEN_GIAM_GIA || col === COL_DH.TIEN_THU_MUA) {
    const soLuong = Number(v[COL_DH.SO_LUONG - 1]) || 0;
    const donGia = Number(v[COL_DH.DON_GIA - 1]) || 0;
    const giamGia = Number(v[COL_DH.TIEN_GIAM_GIA - 1]) || 0;
    const tienThuMua = Number(v[COL_DH.TIEN_THU_MUA - 1]) || 0;
    v[COL_DH.THANH_TIEN - 1] = soLuong * donGia - giamGia - tienThuMua;
  }

  // Auto lookup TenKH khi nhập MaKH
  if (col === COL_DH.MA_KH) {
    const maKH = e.value;
    if (maKH) {
      const tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, COL_KH.MA_KH, maKH, COL_KH.HO_TEN);
      v[COL_DH.TEN_KH - 1] = tenKH || "";
    }
  }

  // Auto lookup TenSP, NguonSP, ThuongHieu khi nhập MaSP
  if (col === COL_DH.MA_SP) {
    const maSP = e.value;
    if (maSP) {
      // Thử tìm trong Điện thoại trước
      const tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.TEN_SP);
      if (tenDT) {
        v[COL_DH.TEN_SP - 1] = tenDT;
        v[COL_DH.NGUON_SP - 1] = "Điện thoại";
        const thuongHieu = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.THUONG_HIEU);
        v[COL_DH.THUONG_HIEU - 1] = thuongHieu || "";
        // Chỉ điền giá trị mặc định nếu ô đó hiện đang trống (tránh ghi đè giá/SL tay đã nhập trước đó)
        if (!v[COL_DH.DON_GIA - 1]) {
          const giaBan = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.GIA_BAN);
          v[COL_DH.DON_GIA - 1] = giaBan || 0;
        }
        if (!v[COL_DH.SO_LUONG - 1]) {
          v[COL_DH.SO_LUONG - 1] = 1; // Điện thoại luôn mặc định SL = 1
        }
        // Tính lại thành tiền
        const soLuong = Number(v[COL_DH.SO_LUONG - 1]) || 0;
        const donGia = Number(v[COL_DH.DON_GIA - 1]) || 0;
        const giamGia = Number(v[COL_DH.TIEN_GIAM_GIA - 1]) || 0;
        v[COL_DH.THANH_TIEN - 1] = soLuong * donGia - giamGia;
      } else {
        // Thử tìm trong Phụ kiện
        const tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.TEN_SP);
        if (tenPK) {
          v[COL_DH.TEN_SP - 1] = tenPK;
          v[COL_DH.NGUON_SP - 1] = "Phụ kiện";
          const thuongHieuPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.THUONG_HIEU);
          v[COL_DH.THUONG_HIEU - 1] = thuongHieuPK || "";
          // Chỉ điền giá trị mặc định nếu ô đó hiện đang trống
          if (!v[COL_DH.DON_GIA - 1]) {
            const giaBanPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.GIA_BAN);
            v[COL_DH.DON_GIA - 1] = giaBanPK || 0;
          }
          if (!v[COL_DH.SO_LUONG - 1]) {
            v[COL_DH.SO_LUONG - 1] = 1;
          }
          // Tính lại thành tiền
          const soLuong = Number(v[COL_DH.SO_LUONG - 1]) || 0;
          const donGia = Number(v[COL_DH.DON_GIA - 1]) || 0;
          const giamGia = Number(v[COL_DH.TIEN_GIAM_GIA - 1]) || 0;
          v[COL_DH.THANH_TIEN - 1] = soLuong * donGia - giamGia;
        }
      }
    }
  }

  // Auto lookup TenNguoiBan + QuyenXuatMay khi nhập NguoiBan (Sử dụng dynamic enum COL_NV)
  if (col === COL_DH.NGUOI_BAN) {
    const maNVBan = e.value;
    if (maNVBan) {
      const tenNVBan = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNVBan, COL_NV.HO_TEN);
      v[COL_DH.TEN_NGUOI_BAN - 1] = tenNVBan || "";
      const quyenXM = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNVBan, COL_NV.QUYEN_XUAT);
      v[COL_DH.QUYEN_XUAT - 1] = quyenXM || "✗";
    }
  }

  // Auto lookup TenNguoiHoTro khi nhập NguoiHoTro (Sử dụng dynamic enum COL_NV)
  if (col === COL_DH.NGUOI_HO_TRO) {
    const maNVHT = e.value;
    if (maNVHT) {
      const tenNVHT = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNVHT, COL_NV.HO_TEN);
      v[COL_DH.TEN_NGUOI_HO_TRO - 1] = tenNVHT || "";
    }
  }

  range.setValues([v]);
}

/**
 * Auto-calculate cho sheet NhapKho (Nhập kho)
 */
function _onEditNhapKho(sheet, row, col, e) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  const range = sheet.getRange(row, 1, 1, lastCol);
  const v = range.getValues()[0];

  // Auto tính ThanhTien khi SoLuong hoặc GiaNhap thay đổi
  if (col === COL_NK.SO_LUONG || col === COL_NK.GIA_NHAP) {
    const soLuong = Number(v[COL_NK.SO_LUONG - 1]) || 0;
    const giaNhap = Number(v[COL_NK.GIA_NHAP - 1]) || 0;
    v[COL_NK.THANH_TIEN - 1] = soLuong * giaNhap;
  }

  // Auto lookup TenSP khi nhập MaSP
  if (col === COL_NK.MA_SP) {
    const maSP = e.value;
    const nguonNhap = v[COL_NK.NGUON_NHAP - 1];
    if (maSP) {
      if (nguonNhap === "Điện thoại") {
        const tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.TEN_SP);
        v[COL_NK.TEN_SP - 1] = tenDT || "";
      } else {
        const tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.TEN_SP);
        v[COL_NK.TEN_SP - 1] = tenPK || "";
      }
    }
  }

  range.setValues([v]);
}
