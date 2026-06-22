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
  return runWithLock(function() {
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
    var tenKH = data.tenKH || "";
    if (data.maKH && !tenKH) {
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

    // Backend validation for Hỗn hợp payments
    if (data.hinhThucThanhToan === "Hỗn hợp") {
      var tongThanhToan = (Number(data.splitChuyenKhoan) || 0) + (Number(data.splitTienMat) || 0);
      var soTienYeuCau = soTienGiaoDich + phiDichVu;
      if (Math.abs(tongThanhToan - soTienYeuCau) > 1) {
        throw new Error("Lỗi dữ liệu: Tổng tiền mặt (" + data.splitTienMat + ") và chuyển khoản (" + data.splitChuyenKhoan + ") không khớp với số tiền cần thanh toán (" + soTienYeuCau + ")!");
      }
    }

    var tienMat = 0;
    var chuyenKhoan = 0;
    if (data.hinhThucThanhToan === "Hỗn hợp") {
      tienMat = Number(data.splitTienMat) || 0;
      chuyenKhoan = Number(data.splitChuyenKhoan) || 0;
    } else if (data.hinhThucThanhToan === "Tiền mặt") {
      tienMat = soTienGiaoDich + phiDichVu;
    } else {
      chuyenKhoan = soTienGiaoDich + phiDichVu;
    }

    var rowData = [];
    rowData[COL_DV.MA_DV - 1] = maDV;
    rowData[COL_DV.NGAY_GD - 1] = new Date();
    rowData[COL_DV.LOAI_DV - 1] = data.loaiDichVu || "";
    rowData[COL_DV.MA_KH - 1] = data.maKH || "";
    rowData[COL_DV.TEN_KH - 1] = tenKH;
    rowData[COL_DV.SDT_KH - 1] = data.soDienThoaiKH || "";
    rowData[COL_DV.SO_TIEN_GD - 1] = soTienGiaoDich;
    rowData[COL_DV.PHI_DV - 1] = phiDichVu;
    rowData[COL_DV.HINH_THUC_TT - 1] = data.hinhThucThanhToan || "Tiền mặt";
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
  });
}

/**
 * Lấy giao dịch dịch vụ theo tháng/năm
 *
 * @param {number} thang
 * @param {number} nam
 * @return {Object[]}
 */
function getDichVuTheoThang(thang, nam) {
  var data = getAllData(SHEET_NAMES.DICH_VU);
  var result = [];

  data.forEach(function (row) {
    var ngay = row[1];
    if (ngay instanceof Date) {
      if (ngay.getMonth() + 1 === thang && ngay.getFullYear() === nam) {
        result.push({
          MaDV: String(row[0]),
          NgayGiaoDich: row[1],
          LoaiDichVu: String(row[2]),
          MaKH: String(row[3]),
          TenKH: String(row[4]),
          SoDienThoaiKH: String(row[5]),
          SoTienGiaoDich: Number(row[6]) || 0,
          PhiDichVu: Number(row[7]) || 0,
          HinhThucThanhToan: String(row[8]),
          NguoiThucHien: String(row[9]),
          TenNguoiThucHien: String(row[10]),
          TrangThai: String(row[11]),
          GhiChu: String(row[12]),
          ChiNhanh: String(row[13] || ""),
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
