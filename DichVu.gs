/**
 * ============================================================
 * VanTran Mobile — DichVu.gs
 * Quản lý dịch vụ: Chuyển khoản hộ, Rút tiền mặt, Nạp thẻ
 * ============================================================
 */

/**
 * Tạo giao dịch dịch vụ mới
 *
 * @param {Object} data - {
 *   loaiDichVu: 'Chuyển khoản hộ' / 'Rút tiền mặt' / 'Nạp thẻ điện thoại',
 *   maKH, tenKH, soDienThoaiKH,
 *   soTienGiaoDich, phiDichVu,
 *   hinhThucThanhToan: 'Tiền mặt' / 'Chuyển khoản' / 'Quẹt thẻ (POS)',
 *   nguoiThucHien (MaNV), ghiChu, chiNhanh
 * }
 * @return {string} Mã dịch vụ mới
 */
function taoDichVu(data) {
    initializeColumnEnums();
    var maDV = generateId("DV", SHEET_NAMES.DICH_VU);
    var chiNhanh = data.chiNhanh;

    if (!chiNhanh) {
      throw new Error("Vui lòng chọn chi nhánh thực hiện giao dịch!");
    }

    // Auto lookup tên NV
    var tenNguoiThucHien = "";
    if (data.nguoiThucHien) {
      tenNguoiThucHien =
        lookupValue(SHEET_NAMES.NHAN_VIEN, 1, data.nguoiThucHien, 2) || "";
    }

    // Auto lookup tên KH nếu có mã KH
    var tenKH = ensureKhachHangExists(data.maKH, data.tenKH);
    if (!tenKH && data.maKH) {
      tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, 1, data.maKH, 2) || "";
    }

    // Xác định phí dịch vụ
    var phiDichVu = Number(data.phiDichVu) || 0;
    if (data.loaiDichVu === "Nạp thẻ điện thoại") {
      phiDichVu = 0; // Nạp thẻ miễn phí
    } else if (phiDichVu === 0) {
      // Lấy phí từ cấu hình nếu chưa nhập
      if (data.loaiDichVu === "Chuyển khoản hộ") {
        phiDichVu = getConfigNumber("Phí CK hộ");
      } else if (data.loaiDichVu === "Rút tiền mặt") {
        phiDichVu = getConfigNumber("Phí Rút tiền");
      }
    }

    var soTienGiaoDich = Number(data.soTienGiaoDich) || 0;

    // Backend validation and calculation for Hỗn hợp / Tiền mặt / Chuyển khoản payments
    var tienMat = 0;
    var chuyenKhoan = 0;
    var hinhThucTTDisplay = data.hinhThucThanhToan;

    if (data.hinhThucThanhToan === "Hỗn hợp") {
      var splitTienMat = Number(data.splitTienMat) || 0;
      var splitChuyenKhoan = Number(data.splitChuyenKhoan) || 0;
      var totalNeeded = soTienGiaoDich + phiDichVu;
      if (Math.abs(splitTienMat + splitChuyenKhoan - totalNeeded) > 1) {
        throw new Error(
          "Lỗi dữ liệu: Tổng tiền mặt (" +
            splitTienMat +
            ") và chuyển khoản (" +
            splitChuyenKhoan +
            ") không khớp với số tiền cần thanh toán (" +
            totalNeeded +
            ")!",
        );
      }
      tienMat = splitTienMat;
      chuyenKhoan = splitChuyenKhoan;
      hinhThucTTDisplay = "Hỗn hợp";
    } else if (data.hinhThucThanhToan === "Tiền mặt") {
      tienMat = soTienGiaoDich + phiDichVu;
    } else {
      // Chuyển khoản hoặc Quẹt thẻ (POS)
      chuyenKhoan = soTienGiaoDich + phiDichVu;
    }

    var rowData = [];
    rowData[COL_DV.MA_DV - 1] = maDV;
    rowData[COL_DV.NGAY_GD - 1] = new Date();
    rowData[COL_DV.LOAI_DV - 1] = data.loaiDichVu || "";
    rowData[COL_DV.MA_KH - 1] = data.maKH || "";
    rowData[COL_DV.TEN_KH - 1] = tenKH;
    rowData[COL_DV.SO_TIEN_GD - 1] = soTienGiaoDich;
    rowData[COL_DV.PHI_DV - 1] = phiDichVu;
    rowData[COL_DV.HINH_THUC_TT - 1] = hinhThucTTDisplay;
    rowData[COL_DV.NGUOI_THUC_HIEN - 1] = data.nguoiThucHien || "";
    rowData[COL_DV.TEN_NGUOI_THUC_HIEN - 1] = tenNguoiThucHien;
    rowData[COL_DV.TRANG_THAI - 1] = "Hoàn thành";
    rowData[COL_DV.GHI_CHU - 1] = data.ghiChu || "";
    rowData[COL_DV.CHI_NHANH - 1] = chiNhanh;
    rowData[COL_DV.TIEN_MAT - 1] = tienMat;
    rowData[COL_DV.CHUYEN_KHOAN - 1] = chuyenKhoan;

    appendRow(SHEET_NAMES.DICH_VU, rowData);
    showToast(
      "✅ Đã tạo dịch vụ: " +
        data.loaiDichVu +
        " (" +
        maDV +
        ")\nSố tiền: " +
        formatCurrency(data.soTienGiaoDich) +
        "đ | Phí: " +
        formatCurrency(phiDichVu) +
        "đ",
    );
    return maDV;
}

/**
 * Lấy giao dịch dịch vụ theo tháng/năm
 *
 * @param {number} thang
 * @param {number} nam
 * @return {Object[]}
 */
function getDichVuTheoThang(thang, nam) {
  initializeColumnEnums();
  var data = getAllData(SHEET_NAMES.DICH_VU);
  var result = [];

  data.forEach(function (row) {
    var ngay = row[COL_DV.NGAY_GD - 1];
    if (ngay instanceof Date) {
      if (ngay.getMonth() + 1 === thang && ngay.getFullYear() === nam) {
        result.push({
          MaDV: String(row[COL_DV.MA_DV - 1] || ""),
          NgayGiaoDich: row[COL_DV.NGAY_GD - 1],
          LoaiDichVu: String(row[COL_DV.LOAI_DV - 1] || ""),
          MaKH: String(row[COL_DV.MA_KH - 1] || ""),
          TenKH: String(row[COL_DV.TEN_KH - 1] || ""),
          SoDienThoaiKH: String(row[COL_DV.MA_KH - 1] || ""), // Mã KH là SĐT
          SoTienGiaoDich: Number(row[COL_DV.SO_TIEN_GD - 1]) || 0,
          PhiDichVu: Number(row[COL_DV.PHI_DV - 1]) || 0,
          HinhThucThanhToan: String(row[COL_DV.HINH_THUC_TT - 1] || ""),
          NguoiThucHien: String(row[COL_DV.NGUOI_THUC_HIEN - 1] || ""),
          TenNguoiThucHien: String(row[COL_DV.TEN_NGUOI_THUC_HIEN - 1] || ""),
          TrangThai: String(row[COL_DV.TRANG_THAI - 1] || ""),
          GhiChu: String(row[COL_DV.GHI_CHU - 1] || ""),
          ChiNhanh: String(row[COL_DV.CHI_NHANH - 1] || ""),
        });
      }
    }
  });

  return result;
}

/**
 * Tổng phí dịch vụ theo NV trong tháng
 *
 * @param {string} maNV - Mã nhân viên
 * @param {number} thang
 * @param {number} nam
 * @return {number} Tổng phí
 */
function getTongPhiDichVu(maNV, thang, nam) {
  var dichVus = getDichVuTheoThang(thang, nam);
  var tong = 0;

  dichVus.forEach(function (dv) {
    if (String(dv.NguoiThucHien) === maNV && dv.TrangThai !== "Huỷ") {
      tong += Number(dv.PhiDichVu) || 0;
    }
  });

  return tong;
}

/**
 * Tổng kết dịch vụ trong tháng
 *
 * @param {number} thang
 * @param {number} nam
 * @return {Object}
 */
function getTongKetDichVu(thang, nam) {
  var dichVus = getDichVuTheoThang(thang, nam);
  var result = {
    tongGiaoDich: 0,
    tongTienGD: 0,
    tongPhi: 0,
    chiTiet: {
      chuyenKhoan: { soGD: 0, tongTien: 0, tongPhi: 0 },
      rutTien: { soGD: 0, tongTien: 0, tongPhi: 0 },
      napThe: { soGD: 0, tongTien: 0, tongPhi: 0 },
    },
  };

  dichVus.forEach(function (dv) {
    if (dv.TrangThai !== "Huỷ") {
      result.tongGiaoDich++;
      result.tongTienGD += Number(dv.SoTienGiaoDich) || 0;
      result.tongPhi += Number(dv.PhiDichVu) || 0;

      var loai = dv.LoaiDichVu;
      if (loai === "Chuyển khoản hộ") {
        result.chiTiet.chuyenKhoan.soGD++;
        result.chiTiet.chuyenKhoan.tongTien += Number(dv.SoTienGiaoDich) || 0;
        result.chiTiet.chuyenKhoan.tongPhi += Number(dv.PhiDichVu) || 0;
      } else if (loai === "Rút tiền mặt") {
        result.chiTiet.rutTien.soGD++;
        result.chiTiet.rutTien.tongTien += Number(dv.SoTienGiaoDich) || 0;
        result.chiTiet.rutTien.tongPhi += Number(dv.PhiDichVu) || 0;
      } else if (loai === "Nạp thẻ điện thoại") {
        result.chiTiet.napThe.soGD++;
        result.chiTiet.napThe.tongTien += Number(dv.SoTienGiaoDich) || 0;
        result.chiTiet.napThe.tongPhi += Number(dv.PhiDichVu) || 0;
      }
    }
  });

  return result;
}
