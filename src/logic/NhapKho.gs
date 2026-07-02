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
  return withDocumentLock(function () {
    const maNK = generateId("NK", SHEET_NAMES.NHAP_KHO);
    const nguonNhap = data.nguonNhap || PRODUCT_SOURCE.PHONE;
    let maSP = "";
    let tenSP = "";
    let soLuong = Number(data.soLuong) || 1;
    const giaNhap = Number(data.giaNhap) || 0;
    
    validateRequiredFields(data, [
      { key: "chiNhanh", label: "Chi nhánh" }
    ]);
    const chiNhanh = data.chiNhanh;

    if (nguonNhap === PRODUCT_SOURCE.PHONE) {
      // Tạo sản phẩm mới trong danh mục điện thoại
      maSP = addDienThoai({
        maDT: data.maDT,
        tenSP: data.tenSP,
        thuongHieu: data.thuongHieu,
        imei: data.imei,
        imei2: data.imei2 || "",
        mauSac: data.mauSac,
        dungLuong: data.dungLuong,
        tinhTrang: data.tinhTrang || "New",
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
        tenSP =
          lookupValue(
            SHEET_NAMES.PHU_KIEN,
            COL_PK.MA_PK,
            maSP,
            COL_PK.TEN_SP,
          ) || "";
        updateTonKhoPhuKien(maSP, soLuong, "nhap", chiNhanh);

        // Cập nhật giá nhập mới nếu có (cho đúng dòng chi nhánh!)
        if (giaNhap > 0) {
          const row = findPhuKienRow(maSP, chiNhanh);
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

    // Ghi vào sheet NhapKho tuân thủ Schema
    const rowData = buildRowData(COL_NK, {
      MA_NK: maNK,
      NGAY_NHAP: new Date(),
      NGUON_NHAP: nguonNhap,
      MA_SP: maSP,
      TEN_SP: tenSP,
      SO_LUONG: soLuong,
      GIA_NHAP: giaNhap,
      THANH_TIEN: soLuong * giaNhap,
      NHA_CUNG_CAP: data.nhaCungCap || "",
      GHI_CHU: data.ghiChu || "",
      CHI_NHANH: chiNhanh
    });

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
  const data = getAllData(SHEET_NAMES.NHAP_KHO);
  const result = [];

  data.forEach(function (row) {
    const obj = mapRowToObject(row, SHEET_NAMES.NHAP_KHO);
    if (String(obj.MA_SP) === maSP) {
      result.push({
        MaNK: String(obj.MA_NK),
        NgayNhap: obj.NGAY_NHAP,
        NguonNhap: String(obj.NGUON_NHAP),
        MaSP: String(obj.MA_SP),
        TenSP: String(obj.TEN_SP),
        SoLuong: Number(obj.SO_LUONG) || 0,
        GiaNhap: Number(obj.GIA_NHAP) || 0,
        ThanhTien: Number(obj.THANH_TIEN) || 0,
        NhaCungCap: String(obj.NHA_CUNG_CAP),
        GhiChu: String(obj.GHI_CHU),
        ChiNhanh: String(obj.CHI_NHANH || ""),
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
  const data = getAllData(SHEET_NAMES.NHAP_KHO);
  const result = {
    tongNhap: 0,
    tongTien: 0,
    dienThoai: { soLuong: 0, tongTien: 0 },
    phuKien: { soLuong: 0, tongTien: 0 },
  };

  data.forEach(function (row) {
    const obj = mapRowToObject(row, SHEET_NAMES.NHAP_KHO);
    const ngay = obj.NGAY_NHAP;
    if (isSameMonthYear(ngay, thang, nam)) {
      result.tongNhap++;
      result.tongTien += Number(obj.THANH_TIEN) || 0;

      if (String(obj.NGUON_NHAP) === PRODUCT_SOURCE.PHONE) {
        result.dienThoai.soLuong += Number(obj.SO_LUONG) || 0;
        result.dienThoai.tongTien += Number(obj.THANH_TIEN) || 0;
      } else {
        result.phuKien.soLuong += Number(obj.SO_LUONG) || 0;
        result.phuKien.tongTien += Number(obj.THANH_TIEN) || 0;
      }
    }
  });

  return result;
}
