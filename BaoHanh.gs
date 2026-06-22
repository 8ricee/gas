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
  initializeColumnEnums();
  var maBH = generateId("BH", SHEET_NAMES.BAO_HANH);
  var chiNhanh = data.chiNhanh;

  if (!chiNhanh) {
    throw new Error("Vui lòng chọn chi nhánh!");
  }

  var tenNguoiTiepNhan = "";
  if (data.nguoiTiepNhan) {
    tenNguoiTiepNhan = lookupValue(SHEET_NAMES.NHAN_VIEN, 1, data.nguoiTiepNhan, 2) || "";
  }

  var tenKH = ensureKhachHangExists(data.maKH, data.tenKH, data.soDienThoaiKH);
  if (!tenKH && data.maKH) {
    tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, 1, data.maKH, 2) || "";
  }

  var phiSuaChua = Number(data.phiSuaChua) || 0;

  var tienMat = 0;
  var chuyenKhoan = 0;
  var hinhThucTTDisplay = data.hinhThucThanhToan || "Tiền mặt";

  if (data.hinhThucThanhToan === "Hỗn hợp") {
    var splitTienMat = Number(data.splitTienMat) || 0;
    var splitChuyenKhoan = Number(data.splitChuyenKhoan) || 0;
    if (Math.abs(splitTienMat + splitChuyenKhoan - phiSuaChua) > 1) {
      throw new Error("Lỗi dữ liệu: Tổng thanh toán không khớp với Phí sửa chữa!");
    }
    tienMat = splitTienMat;
    chuyenKhoan = splitChuyenKhoan;
  } else if (data.hinhThucThanhToan === "Tiền mặt") {
    tienMat = phiSuaChua;
  } else {
    chuyenKhoan = phiSuaChua;
  }

  var rowData = [];
  rowData[COL_BH.MA_BH - 1] = maBH;
  rowData[COL_BH.NGAY_NHAN - 1] = new Date();
  rowData[COL_BH.MA_KH - 1] = data.maKH || "";
  rowData[COL_BH.TEN_KH - 1] = tenKH;
  rowData[COL_BH.SDT_KH - 1] = data.soDienThoaiKH || "";
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
}

/**
 * Cập nhật trạng thái và ghi chú của phiếu bảo hành
 */
function capNhatTrangThaiBaoHanh(maBH, trangThaiMoi, ghiChuMoi) {
  var row = findRow(SHEET_NAMES.BAO_HANH, COL_BH.MA_BH, maBH);
  if (row === -1) {
    throw new Error("Không tìm thấy phiếu " + maBH);
  }
  updateCell(SHEET_NAMES.BAO_HANH, row, COL_BH.TRANG_THAI, trangThaiMoi);
  if (ghiChuMoi) {
      updateCell(SHEET_NAMES.BAO_HANH, row, COL_BH.GHI_CHU, ghiChuMoi);
  }
  showToast("✅ Cập nhật trạng thái thành công: " + maBH);
  return true;
}

/**
 * Lấy danh sách bảo hành theo tháng năm
 */
function getBaoHanhTheoThang(thang, nam) {
  var data = getAllData(SHEET_NAMES.BAO_HANH);
  var result = [];

  data.forEach(function (row) {
    var ngay = row[1];
    if (ngay instanceof Date) {
      if (ngay.getMonth() + 1 === thang && ngay.getFullYear() === nam) {
        result.push({
          MaBH: String(row[0]),
          NgayNhan: row[1],
          MaKH: String(row[2]),
          TenKH: String(row[3]),
          SoDienThoaiKH: String(row[4]),
          TenSP: String(row[5]),
          TinhTrangLoi: String(row[6]),
          LoaiDichVu: String(row[7]),
          PhiSuaChua: Number(row[8]) || 0,
          HinhThucThanhToan: String(row[9]),
          NguoiTiepNhan: String(row[10]),
          NguoiSua: String(row[11]),
          TrangThai: String(row[12]),
          GhiChu: String(row[13]),
          ChiNhanh: String(row[14] || ""),
        });
      }
    }
  });

  return result;
}
