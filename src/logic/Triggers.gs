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

  // 1. Kiểm tra xem có cần xóa cache dropdown/sheet không
  let shouldInvalidateCache = false;
  let shouldInvalidateDropdown = false;

  if (sheetName === SHEET_NAMES.DIEN_THOAI) {
    shouldInvalidateCache = true;
    shouldInvalidateDropdown = (col === COL_DT.IMEI || col === COL_DT.TEN_SP || col === COL_DT.TRANG_THAI_KHO || col === COL_DT.GIA_BAN || col === COL_DT.CHI_NHANH);
  } else if (sheetName === SHEET_NAMES.PHU_KIEN) {
    shouldInvalidateCache = true;
    shouldInvalidateDropdown = (col === COL_PK.MA_PK || col === COL_PK.TEN_SP || col === COL_PK.SO_LUONG_TON || col === COL_PK.GIA_BAN || col === COL_PK.CHI_NHANH);
  } else if (sheetName === SHEET_NAMES.KHACH_HANG) {
    shouldInvalidateCache = true;
    shouldInvalidateDropdown = (col === COL_KH.MA_KH || col === COL_KH.HO_TEN || col === COL_KH.SO_DIEN_THOAI);
  } else if (sheetName === SHEET_NAMES.NHAN_VIEN) {
    shouldInvalidateCache = true;
    shouldInvalidateDropdown = (col === COL_NV.MA_NV || col === COL_NV.HO_TEN || col === COL_NV.TRANG_THAI);
  }

  // 2. Kiểm tra xem cột được sửa đổi có liên quan đến các bộ xử lý (EDIT_HANDLERS) không.
  // Tránh xoá cache và chạy xử lý vô ích đối với các cột không cần auto-lookup/auto-calculate.
  let isRelevant = false;
  if (sheetName === SHEET_NAMES.DON_HANG) {
    isRelevant = (col === COL_DH.SO_LUONG || col === COL_DH.DON_GIA || col === COL_DH.TIEN_GIAM_GIA || col === COL_DH.TIEN_THU_MUA ||
                  col === COL_DH.MA_KH || col === COL_DH.MA_SP || col === COL_DH.NGUOI_BAN || col === COL_DH.NGUOI_HO_TRO);
    shouldInvalidateCache = isRelevant;
  } else if (sheetName === SHEET_NAMES.NHAP_KHO) {
    isRelevant = (col === COL_NK.SO_LUONG || col === COL_NK.GIA_NHAP || col === COL_NK.MA_SP);
    shouldInvalidateCache = isRelevant;
  } else if (sheetName === SHEET_NAMES.LICH_SU_TRA_GOP) {
    isRelevant = (col === COL_LSTG.SO_TIEN_DA_TRA || col === COL_LSTG.TRANG_THAI);
    shouldInvalidateCache = isRelevant;
  } else if (sheetName === SHEET_NAMES.BAO_CAO_NGAY) {
    isRelevant = (row === 3 && col === 2);
  } else if (sheetName === SHEET_NAMES.BAO_CAO_DOANH_SO) {
    isRelevant = (row === 3 && (col === 2 || col === 4 || col === 6 || col === 8));
  } else if (sheetName === SHEET_NAMES.CAU_HINH) {
    isRelevant = true; // Luôn chạy khi sửa đổi cấu hình để cập nhật validation
    shouldInvalidateCache = true;
  }

  if (!isRelevant && !shouldInvalidateCache) return;

  try {
    if (shouldInvalidateCache) {
      clearSheetCache(sheetName);
    }
    if (shouldInvalidateDropdown) {
      invalidateDropdownCache(sheetName);
    }
    if (isRelevant && EDIT_HANDLERS[sheetName]) {
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
  // Trích xuất an toàn giá trị chỉnh sửa (tương thích dán đè dữ liệu hàng loạt)
  const val = e.value !== undefined ? e.value : sheet.getRange(row, col).getValue();

  // Auto tính ThanhTien khi SoLuong hoặc DonGia hoặc TienGiamGia hoặc TienThuMua thay đổi
  if (col === COL_DH.SO_LUONG || col === COL_DH.DON_GIA || col === COL_DH.TIEN_GIAM_GIA || col === COL_DH.TIEN_THU_MUA) {
    const soLuong = Number(sheet.getRange(row, COL_DH.SO_LUONG).getValue()) || 0;
    const donGia = Number(sheet.getRange(row, COL_DH.DON_GIA).getValue()) || 0;
    const giamGia = Number(sheet.getRange(row, COL_DH.TIEN_GIAM_GIA).getValue()) || 0;
    const tienThuMua = Number(sheet.getRange(row, COL_DH.TIEN_THU_MUA).getValue()) || 0;
    sheet.getRange(row, COL_DH.THANH_TIEN).setValue(calcThanhTien(soLuong, donGia, giamGia, tienThuMua));
  }

  // Auto lookup TenKH khi nhập MaKH
  if (col === COL_DH.MA_KH) {
    if (val) {
      const tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, COL_KH.MA_KH, val, COL_KH.HO_TEN);
      sheet.getRange(row, COL_DH.TEN_KH).setValue(tenKH || "");
    }
  }

  // Auto lookup TenSP, NguonSP, ThuongHieu khi nhập MaSP
  if (col === COL_DH.MA_SP) {
    if (val) {
      // Thử tìm trong Điện thoại trước
      const tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, val, COL_DT.TEN_SP);
      if (tenDT) {
        sheet.getRange(row, COL_DH.TEN_SP).setValue(tenDT);
        sheet.getRange(row, COL_DH.NGUON_SP).setValue(PRODUCT_SOURCE.PHONE);
        const thuongHieu = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, val, COL_DT.THUONG_HIEU);
        sheet.getRange(row, COL_DH.THUONG_HIEU).setValue(thuongHieu || "");
        
        const donGiaCell = sheet.getRange(row, COL_DH.DON_GIA);
        if (!donGiaCell.getValue()) {
          const giaBan = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, val, COL_DT.GIA_BAN);
          donGiaCell.setValue(giaBan || 0);
        }
        
        const soLuongCell = sheet.getRange(row, COL_DH.SO_LUONG);
        if (!soLuongCell.getValue()) {
          soLuongCell.setValue(1); // Điện thoại luôn mặc định SL = 1
        }
        
        // Tính lại thành tiền
        const soLuong = Number(soLuongCell.getValue()) || 0;
        const donGia = Number(donGiaCell.getValue()) || 0;
        const giamGia = Number(sheet.getRange(row, COL_DH.TIEN_GIAM_GIA).getValue()) || 0;
        sheet.getRange(row, COL_DH.THANH_TIEN).setValue(calcThanhTien(soLuong, donGia, giamGia, 0));
      } else {
        // Thử tìm trong Phụ kiện
        const tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, val, COL_PK.TEN_SP);
        if (tenPK) {
          sheet.getRange(row, COL_DH.TEN_SP).setValue(tenPK);
          sheet.getRange(row, COL_DH.NGUON_SP).setValue(PRODUCT_SOURCE.ACCESSORY);
          const thuongHieuPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, val, COL_PK.THUONG_HIEU);
          sheet.getRange(row, COL_DH.THUONG_HIEU).setValue(thuongHieuPK || "");
          
          const donGiaCell = sheet.getRange(row, COL_DH.DON_GIA);
          if (!donGiaCell.getValue()) {
            const giaBanPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, val, COL_PK.GIA_BAN);
            donGiaCell.setValue(giaBanPK || 0);
          }
          
          const soLuongCell = sheet.getRange(row, COL_DH.SO_LUONG);
          if (!soLuongCell.getValue()) {
            soLuongCell.setValue(1);
          }
          
          // Tính lại thành tiền
          const soLuong = Number(soLuongCell.getValue()) || 0;
          const donGia = Number(donGiaCell.getValue()) || 0;
          const giamGia = Number(sheet.getRange(row, COL_DH.TIEN_GIAM_GIA).getValue()) || 0;
          sheet.getRange(row, COL_DH.THANH_TIEN).setValue(calcThanhTien(soLuong, donGia, giamGia, 0));
        }
      }
    }
  }

  // Auto lookup TenNguoiBan + QuyenXuatMay khi nhập NguoiBan (Sử dụng dynamic enum COL_NV)
  if (col === COL_DH.NGUOI_BAN) {
    if (val) {
      const tenNVBan = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, val, COL_NV.HO_TEN);
      sheet.getRange(row, COL_DH.TEN_NGUOI_BAN).setValue(tenNVBan || "");
      const quyenXM = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, val, COL_NV.QUYEN_XUAT);
      sheet.getRange(row, COL_DH.QUYEN_XUAT).setValue(quyenXM || "✗");
    }
  }

  // Auto lookup TenNguoiHoTro khi nhập NguoiHoTro (Sử dụng dynamic enum COL_NV)
  if (col === COL_DH.NGUOI_HO_TRO) {
    if (val) {
      const tenNVHT = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, val, COL_NV.HO_TEN);
      sheet.getRange(row, COL_DH.TEN_NGUOI_HO_TRO).setValue(tenNVHT || "");
    }
  }
}

/**
 * Auto-calculate cho sheet NhapKho (Nhập kho)
 */
function _onEditNhapKho(sheet, row, col, e) {
  // Trích xuất an toàn giá trị chỉnh sửa (tương thích dán đè dữ liệu hàng loạt)
  const val = e.value !== undefined ? e.value : sheet.getRange(row, col).getValue();

  // Auto tính ThanhTien khi SoLuong hoặc GiaNhap thay đổi
  if (col === COL_NK.SO_LUONG || col === COL_NK.GIA_NHAP) {
    const soLuong = Number(sheet.getRange(row, COL_NK.SO_LUONG).getValue()) || 0;
    const giaNhap = Number(sheet.getRange(row, COL_NK.GIA_NHAP).getValue()) || 0;
    sheet.getRange(row, COL_NK.THANH_TIEN).setValue(calcThanhTien(soLuong, giaNhap, 0, 0));
  }

  // Auto lookup TenSP khi nhập MaSP
  if (col === COL_NK.MA_SP) {
    const nguonNhap = sheet.getRange(row, COL_NK.NGUON_NHAP).getValue();
    if (val) {
      if (nguonNhap === PRODUCT_SOURCE.PHONE) {
        const tenDT = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, val, COL_DT.TEN_SP);
        sheet.getRange(row, COL_NK.TEN_SP).setValue(tenDT || "");
      } else {
        const tenPK = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, val, COL_PK.TEN_SP);
        sheet.getRange(row, COL_NK.TEN_SP).setValue(tenPK || "");
      }
    }
  }
}
