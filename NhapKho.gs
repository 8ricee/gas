/**
 * ============================================================
 * VanTran Mobile — NhapKho.gs
 * Quản lý nhập kho — Điện thoại (tạo mới) & Phụ kiện (cộng tồn)
 * ============================================================
 */

/**
 * Nhập hàng vào kho
 *
 * @param {Object} data - {
 *   nguonNhap: 'Điện thoại' / 'Phụ kiện',
 *   // Nếu Điện thoại:
 *   tenSP, thuongHieu, imei, mauSac, dungLuong, tinhTrang, giaNhap, giaBan, giaTraGop,
 *   // Nếu Phụ kiện (đã tồn tại):
 *   maSP, soLuong, giaNhap,
 *   // Nếu Phụ kiện (mới):
 *   tenSP, loaiPK, thuongHieu, giaNhap, giaBan, soLuong, moTa,
 *   // Chung:
 *   nhaCungCap, ghiChu, chiNhanh
 * }
 * @return {string} Mã nhập kho
 */
function nhapKho(data) {
  return runWithLock(function() {
    initializeColumnEnums();
    var maNK = generateId("NK", SHEET_NAMES.NHAP_KHO);
    var nguonNhap = data.nguonNhap || "Điện thoại";
    var maSP = "";
    var tenSP = "";
    var soLuong = Number(data.soLuong) || 1;
    var giaNhap = Number(data.giaNhap) || 0;
    var chiNhanh = data.chiNhanh;

    if (!chiNhanh) {
      throw new Error("Vui lòng chọn chi nhánh nhập kho!");
    }

    if (nguonNhap === "Điện thoại") {
      // Tạo sản phẩm mới trong danh mục điện thoại
      maSP = addDienThoai({
        tenSP: data.tenSP,
        thuongHieu: data.thuongHieu,
        imei: data.imei,
        mauSac: data.mauSac,
        dungLuong: data.dungLuong,
        tinhTrang: data.tinhTrang || "Mới 100%",
        giaNhap: giaNhap,
        giaBan: Number(data.giaBan) || 0,
        giaTraGop: Number(data.giaTraGop) || 0,
        ghiChu: data.ghiChu || "",
        chiNhanh: chiNhanh,
      });
      tenSP = data.tenSP;
      soLuong = 1; // Mỗi lần nhập 1 máy
    } else {
      // Phụ kiện
      if (data.maSP) {
        // Phụ kiện đã tồn tại → cộng tồn kho
        maSP = data.maSP;
        tenSP = lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.TEN_SP) || "";
        updateTonKhoPhuKien(maSP, soLuong, "nhap", chiNhanh);

        // Cập nhật giá nhập mới nếu có (cho đúng dòng chi nhánh!)
        if (giaNhap > 0) {
          var row = findPhuKienRow(maSP, chiNhanh);
          if (row !== -1) {
            updateCell(SHEET_NAMES.PHU_KIEN, row, COL_PK.GIA_NHAP, giaNhap);
          }
        }
      } else {
        // Phụ kiện mới → tạo mới trong danh mục
        maSP = addPhuKien({
          tenSP: data.tenSP,
          loaiPK: data.loaiPK || "Khác",
          thuongHieu: data.thuongHieu || "",
          giaNhap: giaNhap,
          giaBan: Number(data.giaBan) || 0,
          soLuongTon: soLuong,
          moTa: data.moTa || "",
          chiNhanh: chiNhanh,
        });
        tenSP = data.tenSP;
      }
    }

    // Ghi vào sheet NhapKho
    var rowData = [
      maNK,
      new Date(), // NgayNhap
      nguonNhap,
      maSP,
      tenSP,
      soLuong,
      giaNhap,
      soLuong * giaNhap, // ThanhTien
      data.nhaCungCap || "",
      data.ghiChu || "",
      chiNhanh, // 11th column
    ];

    appendRow(SHEET_NAMES.NHAP_KHO, rowData);
    showToast(
      "✅ Đã nhập kho: " +
        tenSP +
        " (" +
        maNK +
        ")\nSố lượng: " +
        soLuong +
        " | Tổng: " +
        formatCurrency(soLuong * giaNhap) +
        "đ",
    );
    return maNK;
  });
}

/**
 * Lịch sử nhập kho theo mã SP
 *
 * @param {string} maSP
 * @return {Object[]}
 */
function getLichSuNhapKho(maSP) {
  var data = getAllData(SHEET_NAMES.NHAP_KHO);
  var result = [];

  data.forEach(function (row) {
    if (String(row[3]) === maSP) {
      result.push({
        MaNK: String(row[0]),
        NgayNhap: row[1],
        NguonNhap: String(row[2]),
        MaSP: String(row[3]),
        TenSP: String(row[4]),
        SoLuong: Number(row[5]) || 0,
        GiaNhap: Number(row[6]) || 0,
        ThanhTien: Number(row[7]) || 0,
        NhaCungCap: String(row[8]),
        GhiChu: String(row[9]),
        ChiNhanh: String(row[10] || ""),
      });
    }
  });

  return result;
}

/**
 * Tổng kết nhập kho trong tháng
 *
 * @param {number} thang
 * @param {number} nam
 * @return {Object}
 */
function getTongKetNhapKho(thang, nam) {
  var data = getAllData(SHEET_NAMES.NHAP_KHO);
  var result = {
    tongNhap: 0,
    tongTien: 0,
    dienThoai: { soLuong: 0, tongTien: 0 },
    phuKien: { soLuong: 0, tongTien: 0 },
  };

  data.forEach(function (row) {
    var ngay = row[1];
    if (
      ngay instanceof Date &&
      ngay.getMonth() + 1 === thang &&
      ngay.getFullYear() === nam
    ) {
      result.tongNhap++;
      result.tongTien += Number(row[7]) || 0;

      if (String(row[2]) === "Điện thoại") {
        result.dienThoai.soLuong += Number(row[5]) || 0;
        result.dienThoai.tongTien += Number(row[7]) || 0;
      } else {
        result.phuKien.soLuong += Number(row[5]) || 0;
        result.phuKien.tongTien += Number(row[7]) || 0;
      }
    }
  });

  return result;
}
