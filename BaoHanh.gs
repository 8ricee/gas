/**
 * ============================================================
 * VanTran Mobile — BaoHanh.gs
 * Quản lý Dịch vụ Sửa chữa & Bảo hành
 * ============================================================
 */

/**
 * Tạo giao dịch bảo hành / sửa chữa mới
 *
 * @param {Object} data
 * @return {string} Mã BH
 */
function taoBaoHanh(data) {
  return withDocumentLock(function () {
    initializeColumnEnums();
    var maBH = generateId("BH", SHEET_NAMES.BAO_HANH);
    var chiNhanh = data.chiNhanh;

    if (!chiNhanh) {
      throw new Error("Vui lòng chọn chi nhánh!");
    }

    var tenNguoiTiepNhan = getNhanVienName(data.nguoiTiepNhan);
    var tenKH = ensureKhachHangExists(data.maKH, data.tenKH);
    var phiSuaChua = Number(data.phiSuaChua) || 0;

    var splitResult = calculatePaymentSplit(data, phiSuaChua);
    var tienMat = splitResult.tienMat;
    var chuyenKhoan = splitResult.chuyenKhoan;
    var hinhThucTTDisplay = splitResult.hinhThucTTDisplay;

    var rowData = [];
    rowData[COL_BH.MA_BH - 1] = maBH;
    rowData[COL_BH.NGAY_NHAN - 1] = new Date();
    rowData[COL_BH.MA_KH - 1] = data.maKH || "";
    rowData[COL_BH.TEN_KH - 1] = tenKH;
    rowData[COL_BH.TEN_SP - 1] = data.tenSP || "";
    rowData[COL_BH.TINH_TRANG_LOI - 1] = data.tinhTrangLoi || "";
    rowData[COL_BH.LOAI_DICH_VU - 1] = data.loaiDichVu || "Sửa chữa"; // Sửa chữa hoặc Bảo hành
    rowData[COL_BH.PHI_SUA_CHUA - 1] = phiSuaChua;
    rowData[COL_BH.HINH_THUC_TT - 1] = hinhThucTTDisplay;
    rowData[COL_BH.NGUOI_TIEP_NHAN - 1] = data.nguoiTiepNhan || "";
    rowData[COL_BH.NGUOI_SUA - 1] = data.nguoiSua || "";
    rowData[COL_BH.TRANG_THAI - 1] = data.trangThai || "Đang xử lý";
    rowData[COL_BH.GHI_CHU - 1] = data.ghiChu || "";
    rowData[COL_BH.CHI_NHANH - 1] = chiNhanh;
    rowData[COL_BH.TIEN_MAT - 1] = tienMat;
    rowData[COL_BH.CHUYEN_KHOAN - 1] = chuyenKhoan;

    appendRow(SHEET_NAMES.BAO_HANH, rowData);
    showToast("Đã tiếp nhận " + data.loaiDichVu + " (" + maBH + ")");
    return maBH;
  });
}

/**
 * Cập nhật trạng thái và ghi chú của phiếu bảo hành
 */
function capNhatTrangThaiBaoHanh(maBH, trangThaiMoi, ghiChuMoi) {
  return withDocumentLock(function () {
    var row = findRow(SHEET_NAMES.BAO_HANH, COL_BH.MA_BH, maBH);
    if (row === -1) {
      throw new Error("Không tìm thấy phiếu " + maBH);
    }
    updateCell(SHEET_NAMES.BAO_HANH, row, COL_BH.TRANG_THAI, trangThaiMoi);
    if (ghiChuMoi) {
      updateCell(SHEET_NAMES.BAO_HANH, row, COL_BH.GHI_CHU, ghiChuMoi);
    }
    showToast("Cập nhật trạng thái thành công: " + maBH);
    return true;
  });
}

/**
 * Lấy danh sách bảo hành theo tháng năm
 */
function getBaoHanhTheoThang(thang, nam) {
  initializeColumnEnums();
  var data = getAllData(SHEET_NAMES.BAO_HANH);
  var result = [];

  data.forEach(function (row) {
    var ngay = row[COL_BH.NGAY_NHAN - 1];
    if (isSameMonthYear(ngay, thang, nam)) {
        result.push({
          MaBH: String(row[COL_BH.MA_BH - 1] || ""),
          NgayNhan: row[COL_BH.NGAY_NHAN - 1],
          MaKH: String(row[COL_BH.MA_KH - 1] || ""),
          TenKH: String(row[COL_BH.TEN_KH - 1] || ""),
          SoDienThoaiKH: String(row[COL_BH.MA_KH - 1] || ""), // Mã KH là SĐT
          TenSP: String(row[COL_BH.TEN_SP - 1] || ""),
          TinhTrangLoi: String(row[COL_BH.TINH_TRANG_LOI - 1] || ""),
          LoaiDichVu: String(row[COL_BH.LOAI_DICH_VU - 1] || ""),
          PhiSuaChua: Number(row[COL_BH.PHI_SUA_CHUA - 1]) || 0,
          HinhThucThanhToan: String(row[COL_BH.HINH_THUC_TT - 1] || ""),
          NguoiTiepNhan: String(row[COL_BH.NGUOI_TIEP_NHAN - 1] || ""),
          NguoiSua: String(row[COL_BH.NGUOI_SUA - 1] || ""),
          TrangThai: String(row[COL_BH.TRANG_THAI - 1] || ""),
          GhiChu: String(row[COL_BH.GHI_CHU - 1] || ""),
          ChiNhanh: String(row[COL_BH.CHI_NHANH - 1] || ""),
        });
      }
    }
  );

  return result;
}
