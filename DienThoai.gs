/**
 * ============================================================
 * VanTran Mobile — DienThoai.gs
 * Quản lý danh mục điện thoại (theo IMEI)
 * ============================================================
 */

/**
 * Lấy index cột (0-indexed) cho Điện thoại
 * @private
 */
function _getDienThoaiIndices() {
  return {
    maDT: COL_DT.MA_DT - 1,
    tenSP: COL_DT.TEN_SP - 1,
    thuongHieu: COL_DT.THUONG_HIEU - 1,
    imei: COL_DT.IMEI - 1,
    imei2: COL_DT.IMEI_2 - 1,
    mauSac: COL_DT.MAU_SAC - 1,
    dungLuong: COL_DT.DUNG_LUONG - 1,
    tinhTrang: COL_DT.TINH_TRANG - 1,
    giaNhap: COL_DT.GIA_NHAP - 1,
    giaBan: COL_DT.GIA_BAN - 1,
    giaTraGop: COL_DT.GIA_TRA_GOP - 1,
    trangThaiKho: COL_DT.TRANG_THAI_KHO - 1,
    ghiChu: COL_DT.GHI_CHU - 1,
    chiNhanh: COL_DT.CHI_NHANH - 1,
    ngayNhap: COL_DT.NGAY_NHAP - 1,
    ngayXuat: COL_DT.NGAY_XUAT - 1,
  };
}

/**
 * Lấy danh sách tất cả điện thoại
 *
 * @param {string} [filter] - Lọc theo TrangThaiKho: 'Còn hàng', 'Đã bán', etc.
 * @return {Object[]} Mảng objects điện thoại
 */
function getAllDienThoai(filter) {
  const data = getAllData(SHEET_NAMES.DIEN_THOAI);
  const result = [];
  const c = _getDienThoaiIndices();

  data.forEach(function (row) {
    if (row.length <= c.trangThaiKho) return;
    if (!filter || String(row[c.trangThaiKho]) === filter) {
      result.push({
        MaDT: String(row[c.maDT]),
        TenSP: String(row[c.tenSP]),
        ThuongHieu: String(row[c.thuongHieu]),
        IMEI: String(row[c.imei]),
        IMEI2:
          c.imei2 !== undefined && row.length > c.imei2
            ? String(row[c.imei2] || "")
            : "",
        MauSac: String(row[c.mauSac]),
        DungLuong: String(row[c.dungLuong]),
        TinhTrang: String(row[c.tinhTrang]),
        GiaNhap: Number(row[c.giaNhap]) || 0,
        GiaBan: Number(row[c.giaBan]) || 0,
        GiaTraGop: Number(row[c.giaTraGop]) || 0,
        TrangThaiKho: String(row[c.trangThaiKho]),
        GhiChu: String(row[c.ghiChu]),
        ChiNhanh: String(row[c.chiNhanh] || ""),
        NgayNhap: row[c.ngayNhap] ? new Date(row[c.ngayNhap]) : null,
        NgayXuat: row[c.ngayXuat] ? new Date(row[c.ngayXuat]) : null,
      });
    }
  });

  return result;
}

/**
 * Lấy danh sách điện thoại còn hàng tại một chi nhánh để hiển thị dropdown
 *
 * @param {string} [chiNhanh] - Chi nhánh lọc
 * @return {Object[]} [{value: 'DT001', text: 'DT001 - iPhone 15 (IMEI: xxx)'}, ...]
 */
function getDienThoaiDropdown(chiNhanh) {
  const data = getAllData(SHEET_NAMES.DIEN_THOAI);
  const result = [];
  const c = _getDienThoaiIndices();

  data.forEach(function (row) {
    if (row.length <= c.trangThaiKho) return;
    if (
      row[c.maDT] &&
      String(row[c.maDT]).trim() !== "" &&
      String(row[c.trangThaiKho]) === "Còn hàng"
    ) {
      const rowChiNhanh = String(row[c.chiNhanh] || "");
      if (!chiNhanh || rowChiNhanh === chiNhanh) {
        const imeiVal = String(row[c.imei] || "");
        const imei2Val =
          c.imei2 !== undefined && row.length > c.imei2
            ? String(row[c.imei2] || "")
            : "";
        let imeiText = imeiVal;
        if (imei2Val) {
          imeiText += " / " + imei2Val;
        }

        result.push({
          value: String(row[c.maDT]),
          text:
            row[c.maDT] +
            " - " +
            row[c.tenSP] +
            " (" +
            row[c.mauSac] +
            ", " +
            row[c.dungLuong] +
            " | IMEI: " +
            imeiText +
            " | " +
            rowChiNhanh +
            ")",
          thuongHieu: String(row[c.thuongHieu]),
          giaBan: Number(row[c.giaBan]),
          giaTraGop: Number(row[c.giaTraGop]),
          chiNhanh: rowChiNhanh,
        });
      }
    }
  });

  return result;
}

/**
 * Thêm điện thoại mới vào danh mục
 *
 * @param {Object} data - {tenSP, thuongHieu, imei, mauSac, dungLuong, tinhTrang, giaNhap, giaBan, giaTraGop, ghiChu, chiNhanh}
 * @return {string} Mã DT mới
 */
function addDienThoai(data) {
  // Kiểm tra IMEI trùng
  if (data.imei) {
    const existing = lookupValue(
      SHEET_NAMES.DIEN_THOAI,
      COL_DT.IMEI,
      data.imei,
      COL_DT.MA_DT,
    );
    if (existing) {
      throw new Error(
        'IMEI "' +
          data.imei +
          '" đã tồn tại trong hệ thống (Mã: ' +
          existing +
          ")",
      );
    }
  }

  const maDT = generateId("DT", SHEET_NAMES.DIEN_THOAI);

  // Tạo mảng dữ liệu với độ dài lớn nhất là 15
  const rowData = [];
  rowData[COL_DT.MA_DT - 1] = maDT;
  rowData[COL_DT.TEN_SP - 1] = data.tenSP || "";
  rowData[COL_DT.THUONG_HIEU - 1] = data.thuongHieu || "";
  rowData[COL_DT.IMEI - 1] = data.imei || "";
  rowData[COL_DT.MAU_SAC - 1] = data.mauSac || "";
  rowData[COL_DT.DUNG_LUONG - 1] = data.dungLuong || "";
  rowData[COL_DT.TINH_TRANG - 1] = data.tinhTrang || "Mới 100%";
  rowData[COL_DT.GIA_NHAP - 1] = Number(data.giaNhap) || 0;
  rowData[COL_DT.GIA_BAN - 1] = Number(data.giaBan) || 0;
  rowData[COL_DT.GIA_TRA_GOP - 1] = Number(data.giaTraGop) || 0;
  rowData[COL_DT.TRANG_THAI_KHO - 1] = "Còn hàng";
  rowData[COL_DT.GHI_CHU - 1] = data.ghiChu || "";
  rowData[COL_DT.CHI_NHANH - 1] = data.chiNhanh || "";
  rowData[COL_DT.NGAY_NHAP - 1] = new Date();
  rowData[COL_DT.NGAY_XUAT - 1] = "";

  appendRow(SHEET_NAMES.DIEN_THOAI, rowData);
  showToast("Đã thêm điện thoại: " + data.tenSP + " (" + maDT + ")");
  return maDT;
}

/**
 * Cập nhật thông tin điện thoại
 *
 * @param {string} maDT - Mã điện thoại
 * @param {Object} data - Dữ liệu cập nhật
 * @return {boolean}
 */
function updateDienThoai(maDT, data) {
  const row = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maDT);
  if (row === -1) {
    showAlert("Lỗi", "Không tìm thấy điện thoại: " + maDT);
    return false;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  invalidateDropdownCache(SHEET_NAMES.DIEN_THOAI);

  // Đảm bảo đủ số cột
  const maxColNeeded = Math.max(
    COL_DT.MA_DT,
    COL_DT.TEN_SP,
    COL_DT.THUONG_HIEU,
    COL_DT.IMEI,
    COL_DT.MAU_SAC,
    COL_DT.DUNG_LUONG,
    COL_DT.TINH_TRANG,
    COL_DT.GIA_NHAP,
    COL_DT.GIA_BAN,
    COL_DT.GIA_TRA_GOP,
    COL_DT.TRANG_THAI_KHO,
    COL_DT.GHI_CHU,
    COL_DT.CHI_NHANH,
    COL_DT.NGAY_NHAP,
    COL_DT.NGAY_XUAT,
  );
  const maxCols = sheet.getMaxColumns();
  if (maxCols < maxColNeeded) {
    sheet.insertColumnsAfter(maxCols, maxColNeeded - maxCols);
  }

  const range = sheet.getRange(row, 1, 1, maxColNeeded);
  const rowValues = range.getValues()[0];

  if (data.tenSP !== undefined)
    rowValues[COL_DT.TEN_SP - 1] = data.tenSP;
  if (data.thuongHieu !== undefined)
    rowValues[COL_DT.THUONG_HIEU - 1] = data.thuongHieu;
  if (data.imei !== undefined)
    rowValues[COL_DT.IMEI - 1] = data.imei;
  if (data.mauSac !== undefined)
    rowValues[COL_DT.MAU_SAC - 1] = data.mauSac;
  if (data.dungLuong !== undefined)
    rowValues[COL_DT.DUNG_LUONG - 1] = data.dungLuong;
  if (data.tinhTrang !== undefined)
    rowValues[COL_DT.TINH_TRANG - 1] = data.tinhTrang;
  if (data.giaNhap !== undefined)
    rowValues[COL_DT.GIA_NHAP - 1] = Number(data.giaNhap);
  if (data.giaBan !== undefined)
    rowValues[COL_DT.GIA_BAN - 1] = Number(data.giaBan);
  if (data.giaTraGop !== undefined)
    rowValues[COL_DT.GIA_TRA_GOP - 1] = Number(data.giaTraGop);
  if (data.trangThaiKho !== undefined) {
    rowValues[COL_DT.TRANG_THAI_KHO - 1] = data.trangThaiKho;
    if (
      data.trangThaiKho === "Đã bán" ||
      data.trangThaiKho === "Đang trả góp"
    ) {
      rowValues[COL_DT.NGAY_XUAT - 1] = new Date();
    } else {
      rowValues[COL_DT.NGAY_XUAT - 1] = "";
    }
  }
  if (data.ghiChu !== undefined)
    rowValues[COL_DT.GHI_CHU - 1] = data.ghiChu;
  if (data.chiNhanh !== undefined)
    rowValues[COL_DT.CHI_NHANH - 1] = data.chiNhanh;
  if (data.ngayNhap !== undefined)
    rowValues[COL_DT.NGAY_NHAP - 1] = data.ngayNhap ? new Date(data.ngayNhap) : "";
  if (data.ngayXuat !== undefined)
    rowValues[COL_DT.NGAY_XUAT - 1] = data.ngayXuat ? new Date(data.ngayXuat) : "";

  range.setValues([rowValues]);

  showToast("Đã cập nhật ĐT: " + maDT);
  return true;
}

/**
 * Cập nhật trạng thái kho cho điện thoại
 *
 * @param {string} maDT - Mã điện thoại
 * @param {string} trangThai - 'Còn hàng' / 'Đã bán' / 'Đang trả góp' / 'Đã trả lại'
 * @return {boolean}
 */
function updateTrangThaiKhoDT(maDT_or_imei, trangThai) {
  let row = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.IMEI, maDT_or_imei); // Tìm theo IMEI trước
  if (row === -1 && COL_DT.IMEI_2) {
    row = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.IMEI_2, maDT_or_imei); // Fallback tìm theo IMEI 2
  }
  if (row === -1) {
    row = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maDT_or_imei); // Fallback tìm theo Mã điện thoại
  }
  if (row === -1) return false;

  updateCell(SHEET_NAMES.DIEN_THOAI, row, COL_DT.TRANG_THAI_KHO, trangThai);

  // Cập nhật Ngày xuất tương ứng
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  const maxColNeeded = Math.max(COL_DT.TRANG_THAI_KHO, COL_DT.NGAY_XUAT);
  const maxCols = sheet.getMaxColumns();
  if (maxCols < maxColNeeded) {
    sheet.insertColumnsAfter(maxCols, maxColNeeded - maxCols);
  }
  if (trangThai === "Đã bán" || trangThai === "Đang trả góp") {
    sheet.getRange(row, COL_DT.NGAY_XUAT).setValue(new Date());
  } else {
    sheet.getRange(row, COL_DT.NGAY_XUAT).setValue("");
  }
  return true;
}

/**
 * Tìm kiếm điện thoại theo từ khóa
 *
 * @param {string} keyword - Từ khóa tìm (tên, IMEI, thương hiệu)
 * @return {Object[]} Kết quả tìm kiếm
 */
function searchDienThoai(keyword) {
  const data = getAllData(SHEET_NAMES.DIEN_THOAI);
  const result = [];
  const kw = String(keyword).trim().toLowerCase();
  const c = _getDienThoaiIndices();

  data.forEach(function (row) {
    if (row.length <= c.imei) return;
    const tenSP = String(row[c.tenSP]).toLowerCase();
    const thuongHieu = String(row[c.thuongHieu]).toLowerCase();
    const imei = String(row[c.imei]).toLowerCase();
    const imei2 =
      c.imei2 !== undefined && row.length > c.imei2
        ? String(row[c.imei2]).toLowerCase()
        : "";

    if (
      tenSP.indexOf(kw) !== -1 ||
      thuongHieu.indexOf(kw) !== -1 ||
      imei.indexOf(kw) !== -1 ||
      imei2.indexOf(kw) !== -1
    ) {
      result.push({
        MaDT: String(row[c.maDT]),
        TenSP: String(row[c.tenSP]),
        ThuongHieu: String(row[c.thuongHieu]),
        IMEI: String(row[c.imei]),
        IMEI2:
          c.imei2 !== undefined && row.length > c.imei2
            ? String(row[c.imei2] || "")
            : "",
        MauSac: String(row[c.mauSac]),
        DungLuong: String(row[c.dungLuong]),
        TinhTrang: String(row[c.tinhTrang]),
        GiaNhap: Number(row[c.giaNhap]) || 0,
        GiaBan: Number(row[c.giaBan]) || 0,
        GiaTraGop: Number(row[c.giaTraGop]) || 0,
        TrangThaiKho: String(row[c.trangThaiKho]),
        GhiChu: String(row[c.ghiChu]),
        ChiNhanh: String(row[c.chiNhanh] || ""),
        NgayNhap: row[c.ngayNhap] ? new Date(row[c.ngayNhap]) : null,
        NgayXuat: row[c.ngayXuat] ? new Date(row[c.ngayXuat]) : null,
      });
    }
  });

  return result;
}

/**
 * Chuẩn hóa và tự động điền Ngày nhập & Ngày xuất cho các dữ liệu điện thoại đã có sẵn
 */
function backfillDienThoaiDates(silent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  if (!dtSheet) return;

  // 1. Đảm bảo sheet Điện thoại có đủ các cột Ngày nhập và Ngày xuất
  const maxColNeeded = Math.max(COL_DT.NGAY_NHAP, COL_DT.NGAY_XUAT);
  let maxCols = dtSheet.getMaxColumns();
  if (maxCols < maxColNeeded) {
    dtSheet.insertColumnsAfter(maxCols, maxColNeeded - maxCols);
  }

  // 2. Kiểm tra và ghi tiêu đề cột Ngày nhập và Ngày xuất nếu chưa có hoặc sai tiêu đề
  let needsHeaderStyle = false;
  if (
    String(dtSheet.getRange(1, COL_DT.NGAY_NHAP).getValue()).trim() !==
    "Ngày nhập"
  ) {
    dtSheet.getRange(1, COL_DT.NGAY_NHAP).setValue("Ngày nhập");
    needsHeaderStyle = true;
  }
  if (
    String(dtSheet.getRange(1, COL_DT.NGAY_XUAT).getValue()).trim() !==
    "Ngày xuất"
  ) {
    dtSheet.getRange(1, COL_DT.NGAY_XUAT).setValue("Ngày xuất");
    needsHeaderStyle = true;
  }

  if (needsHeaderStyle) {
    const newHeadersRange = dtSheet.getRange(
      1,
      Math.min(COL_DT.NGAY_NHAP, COL_DT.NGAY_XUAT),
      1,
      2,
    );
    newHeadersRange.setBackground("#1a73e8");
    newHeadersRange.setFontColor("#ffffff");
    newHeadersRange.setFontWeight("bold");
    newHeadersRange.setHorizontalAlignment("left");
  }

  const lastRow = dtSheet.getLastRow();
  const lastCol = dtSheet.getLastColumn();
  if (lastRow <= 1) {
    if (!silent)
      showAlert(
        "Thông báo",
        "Sheet Điện thoại chưa có dòng dữ liệu nào để chuẩn hóa!",
      );
    return;
  }

  // 3. Đọc dữ liệu động theo số cột thực tế
  const dtData = dtSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const orders = getAllData(SHEET_NAMES.DON_HANG);
  const imports = getAllData(SHEET_NAMES.NHAP_KHO);
  const buybacks = getAllData(SHEET_NAMES.THU_MUA);

  let updatedCount = 0;

  for (let i = 0; i < dtData.length; i++) {
    const maDT = String(dtData[i][COL_DT.MA_DT - 1]);
    const imei = String(dtData[i][COL_DT.IMEI - 1]);
    const status = String(dtData[i][COL_DT.TRANG_THAI_KHO - 1]);

    const currentNgayNhap = dtData[i][COL_DT.NGAY_NHAP - 1];
    const currentNgayXuat = dtData[i][COL_DT.NGAY_XUAT - 1];
    let changed = false;

    // A. Tìm Ngày nhập (nếu trống)
    if (!currentNgayNhap) {
      let ngayNhap = null;
      // Thử tìm trong Nhập kho theo Mã điện thoại (cột 4, index 3)
      for (let j = 0; j < imports.length; j++) {
        if (String(imports[j][3]) === maDT) {
          ngayNhap = imports[j][1]; // Ngày nhập (cột 2, index 1)
          break;
        }
      }
      // Nếu không thấy, tìm trong Thu mua theo IMEI (cột 8, index 7)
      if (!ngayNhap && imei) {
        for (let k = 0; k < buybacks.length; k++) {
          if (String(buybacks[k][7]) === imei) {
            ngayNhap = buybacks[k][1]; // Ngày thu mua (cột 2, index 1)
            break;
          }
        }
      }

      if (ngayNhap) {
        dtData[i][COL_DT.NGAY_NHAP - 1] = new Date(ngayNhap);
      } else {
        dtData[i][COL_DT.NGAY_NHAP - 1] = new Date(); // Mặc định là ngày hiện tại nếu không tìm thấy lịch sử
      }
      changed = true;
    }

    // B. Tìm/Xóa Ngày xuất
    if (status === "Đã bán" || status === "Đang trả góp") {
      if (!currentNgayXuat) {
        let ngayXuat = null;
        // Tìm trong Đơn hàng theo Mã điện thoại (cột 5, index 4)
        for (let j = 0; j < orders.length; j++) {
          if (String(orders[j][4]) === maDT) {
            const orderStatus = String(orders[j][18]);
            if (orderStatus !== "Huỷ" && orderStatus !== "Đổi trả") {
              ngayXuat = orders[j][1]; // Ngày bán (cột 2, index 1)
              break;
            }
          }
        }
        if (ngayXuat) {
          dtData[i][COL_DT.NGAY_XUAT - 1] = new Date(ngayXuat);
          changed = true;
        }
      }
    } else {
      // Nếu trạng thái khác (Còn hàng, Đã trả lại) mà có ngày xuất, thực hiện xoá ngày xuất
      if (currentNgayXuat) {
        dtData[i][COL_DT.NGAY_XUAT - 1] = "";
        changed = true;
      }
    }

    if (changed) {
      updatedCount++;
    }
  }

  // 4. Ghi ngược kết quả bằng Batch Write để tối ưu hóa hiệu suất
  const ngayNhapVals = dtData.map(function (row) {
    return [
      row[COL_DT.NGAY_NHAP - 1] ? new Date(row[COL_DT.NGAY_NHAP - 1]) : "",
    ];
  });
  const ngayXuatVals = dtData.map(function (row) {
    return [
      row[COL_DT.NGAY_XUAT - 1] ? new Date(row[COL_DT.NGAY_XUAT - 1]) : "",
    ];
  });

  dtSheet.getRange(2, COL_DT.NGAY_NHAP, lastRow - 1, 1).setValues(ngayNhapVals);
  dtSheet.getRange(2, COL_DT.NGAY_XUAT, lastRow - 1, 1).setValues(ngayXuatVals);

  // 5. Định dạng lại cột hiển thị
  const rangeNgayNhap = dtSheet.getRange(2, COL_DT.NGAY_NHAP, lastRow - 1, 1);
  const rangeNgayXuat = dtSheet.getRange(2, COL_DT.NGAY_XUAT, lastRow - 1, 1);
  rangeNgayNhap
    .setNumberFormat("dd/MM/yyyy")
    .setFontFamily("Times New Roman")
    .setFontSize(12)
    .setHorizontalAlignment("left");
  rangeNgayXuat
    .setNumberFormat("dd/MM/yyyy")
    .setFontFamily("Times New Roman")
    .setFontSize(12)
    .setHorizontalAlignment("left");

  // Tự động căn chỉnh lại độ rộng cột N & O
  dtSheet.autoResizeColumn(COL_DT.NGAY_NHAP);
  dtSheet.autoResizeColumn(COL_DT.NGAY_XUAT);

  if (!silent) {
    showAlert(
      "Thành công",
      "Đã chuẩn hóa cấu trúc cột và cập nhật " +
        updatedCount +
        " dòng dữ liệu Điện thoại cũ thành công!",
    );
  }
}

/**
 * Build cache dữ liệu điện thoại còn hàng cho tìm kiếm nhanh
 * @private
 */
function _buildDienThoaiDropdownCache() {
  const data = getAllData(SHEET_NAMES.DIEN_THOAI);
  const result = [];
  const c = _getDienThoaiIndices();

  data.forEach(function (row) {
    if (row.length <= c.trangThaiKho) return;
    if (
      row[c.maDT] &&
      String(row[c.maDT]).trim() !== "" &&
      String(row[c.trangThaiKho]) === "Còn hàng"
    ) {
      const cn = String(row[c.chiNhanh] || "");
      const imeiVal = String(row[c.imei] || "");
      const imei2Val =
        c.imei2 !== undefined && row.length > c.imei2
          ? String(row[c.imei2] || "")
          : "";
      let imeiText = imeiVal;
      if (imei2Val) {
        imeiText += " / " + imei2Val;
      }

      result.push({
        v: String(row[c.maDT]),
        t:
          row[c.maDT] +
          " - " +
          row[c.tenSP] +
          " (" +
          row[c.mauSac] +
          ", " +
          row[c.dungLuong] +
          " | IMEI: " +
          imeiText +
          " | " +
          cn +
          ")",
        th: String(row[c.thuongHieu]),
        gb: Number(row[c.giaBan]),
        gtg: Number(row[c.giaTraGop]),
        cn: cn,
        _s: (
          String(row[c.maDT]) +
          " " +
          String(row[c.tenSP]) +
          " " +
          imeiVal +
          " " +
          imei2Val
        ).toLowerCase(),
      });
    }
  });

  return result;
}

/**
 * Lấy danh sách điện thoại cho dropdown theo chi nhánh và từ khóa (Tên, Mã, IMEI)
 * Sử dụng CacheService để tránh đọc sheet mỗi lần tìm kiếm
 *
 * @param {string} chiNhanh - Chi nhánh lọc
 * @param {string} keyword - Từ khóa tìm
 * @return {Object[]}
 */
function getDienThoaiDropdownSearch(chiNhanh, keyword) {
  const kw = String(keyword).trim().toLowerCase();

  let allItems = getChunkedCache("dd_dt");
  if (!allItems) {
    allItems = _buildDienThoaiDropdownCache();
    setChunkedCache("dd_dt", allItems);
  }

  const result = [];
  for (let i = 0; i < allItems.length; i++) {
    const item = allItems[i];
    if (chiNhanh && item.cn !== chiNhanh) continue;
    if (item._s.indexOf(kw) !== -1) {
      result.push({
        value: item.v,
        text: item.t,
        thuongHieu: item.th,
        giaBan: item.gb,
        giaTraGop: item.gtg,
        chiNhanh: item.cn,
      });
      // Nếu tìm kiếm cụ thể (có keyword) thì giới hạn 100; nếu là preload (không keyword) thì giới hạn 5000
      if (kw && result.length >= 100) break;
      if (!kw && result.length >= 5000) break;
    }
  }

  return result;
}
