/**
 * ============================================================
 * VanTran Mobile — DienThoai.gs
 * Quản lý danh mục điện thoại (theo IMEI)
 * ============================================================
 */

/**
 * Lấy danh sách tất cả điện thoại
 *
 * @param {string} [filter] - Lọc theo TrangThaiKho: 'Còn hàng', 'Đã bán', etc.
 * @return {Object[]} Mảng objects điện thoại
 */
function getAllDienThoai(filter) {
  var data = getAllData(SHEET_NAMES.DIEN_THOAI);
  var result = [];

  data.forEach(function (row) {
    if (!filter || String(row[10]) === filter) {
      result.push({
        MaDT: String(row[0]),
        TenSP: String(row[1]),
        ThuongHieu: String(row[2]),
        IMEI: String(row[3]),
        MauSac: String(row[4]),
        DungLuong: String(row[5]),
        TinhTrang: String(row[6]),
        GiaNhap: Number(row[7]) || 0,
        GiaBan: Number(row[8]) || 0,
        GiaTraGop: Number(row[9]) || 0,
        TrangThaiKho: String(row[10]),
        GhiChu: String(row[11]),
        ChiNhanh: String(row[12] || ""),
        NgayNhap: row[13] ? new Date(row[13]) : null,
        NgayXuat: row[14] ? new Date(row[14]) : null,
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
  var data = getAllData(SHEET_NAMES.DIEN_THOAI);
  var result = [];

  data.forEach(function (row) {
    if (
      row[0] &&
      String(row[0]).trim() !== "" &&
      String(row[10]) === "Còn hàng"
    ) {
      var rowChiNhanh = String(row[12] || "");
      if (!chiNhanh || rowChiNhanh === chiNhanh) {
        result.push({
          value: String(row[0]),
          text:
            row[0] +
            " - " +
            row[1] +
            " (" +
            row[4] +
            ", " +
            row[5] +
            " | IMEI: " +
            row[3] +
            " | " +
            rowChiNhanh +
            ")",
          thuongHieu: String(row[2]),
          giaBan: Number(row[8]),
          giaTraGop: Number(row[9]),
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
    var existing = lookupValue(SHEET_NAMES.DIEN_THOAI, 4, data.imei, 1);
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

  var maDT = generateId("DT", SHEET_NAMES.DIEN_THOAI);

  var rowData = [
    maDT,
    data.tenSP || "",
    data.thuongHieu || "",
    data.imei || "",
    data.mauSac || "",
    data.dungLuong || "",
    data.tinhTrang || "Mới 100%",
    Number(data.giaNhap) || 0,
    Number(data.giaBan) || 0,
    Number(data.giaTraGop) || 0,
    "Còn hàng",
    data.ghiChu || "",
    data.chiNhanh || "",
    new Date(), // Ngày nhập (Cột N)
    "",         // Ngày xuất (Cột O)
  ];

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
  var row = findRow(SHEET_NAMES.DIEN_THOAI, 1, maDT);
  if (row === -1) {
    showAlert("Lỗi", "Không tìm thấy điện thoại: " + maDT);
    return false;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  invalidateDropdownCache(SHEET_NAMES.DIEN_THOAI);
  var maxCols = sheet.getMaxColumns();
  if (maxCols < 15) {
    sheet.insertColumnsAfter(maxCols, 15 - maxCols);
  }

  if (data.tenSP !== undefined) sheet.getRange(row, 2).setValue(data.tenSP);
  if (data.thuongHieu !== undefined)
    sheet.getRange(row, 3).setValue(data.thuongHieu);
  if (data.imei !== undefined) sheet.getRange(row, 4).setValue(data.imei);
  if (data.mauSac !== undefined) sheet.getRange(row, 5).setValue(data.mauSac);
  if (data.dungLuong !== undefined)
    sheet.getRange(row, 6).setValue(data.dungLuong);
  if (data.tinhTrang !== undefined)
    sheet.getRange(row, 7).setValue(data.tinhTrang);
  if (data.giaNhap !== undefined)
    sheet.getRange(row, 8).setValue(Number(data.giaNhap));
  if (data.giaBan !== undefined)
    sheet.getRange(row, 9).setValue(Number(data.giaBan));
  if (data.giaTraGop !== undefined)
    sheet.getRange(row, 10).setValue(Number(data.giaTraGop));
  if (data.trangThaiKho !== undefined) {
    sheet.getRange(row, 11).setValue(data.trangThaiKho);
    if (data.trangThaiKho === "Đã bán" || data.trangThaiKho === "Đang trả góp") {
      sheet.getRange(row, 15).setValue(new Date());
    } else {
      sheet.getRange(row, 15).setValue("");
    }
  }
  if (data.ghiChu !== undefined) sheet.getRange(row, 12).setValue(data.ghiChu);
  if (data.chiNhanh !== undefined)
    sheet.getRange(row, 13).setValue(data.chiNhanh);
  if (data.ngayNhap !== undefined) sheet.getRange(row, 14).setValue(data.ngayNhap ? new Date(data.ngayNhap) : "");
  if (data.ngayXuat !== undefined) sheet.getRange(row, 15).setValue(data.ngayXuat ? new Date(data.ngayXuat) : "");

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
  var row = findRow(SHEET_NAMES.DIEN_THOAI, 4, maDT_or_imei); // Tìm theo IMEI trước
  if (row === -1) {
    row = findRow(SHEET_NAMES.DIEN_THOAI, 1, maDT_or_imei); // Fallback tìm theo Mã điện thoại
  }
  if (row === -1) return false;

  updateCell(SHEET_NAMES.DIEN_THOAI, row, 11, trangThai);

  // Cập nhật Ngày xuất tương ứng
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  var maxCols = sheet.getMaxColumns();
  if (maxCols < 15) {
    sheet.insertColumnsAfter(maxCols, 15 - maxCols);
  }
  if (trangThai === "Đã bán" || trangThai === "Đang trả góp") {
    sheet.getRange(row, 15).setValue(new Date());
  } else {
    sheet.getRange(row, 15).setValue("");
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
  var data = getAllData(SHEET_NAMES.DIEN_THOAI);
  var result = [];
  var kw = String(keyword).trim().toLowerCase();

  data.forEach(function (row) {
    var tenSP = String(row[1]).toLowerCase();
    var thuongHieu = String(row[2]).toLowerCase();
    var imei = String(row[3]).toLowerCase();

    if (
      tenSP.indexOf(kw) !== -1 ||
      thuongHieu.indexOf(kw) !== -1 ||
      imei.indexOf(kw) !== -1
    ) {
      result.push({
        MaDT: String(row[0]),
        TenSP: String(row[1]),
        ThuongHieu: String(row[2]),
        IMEI: String(row[3]),
        MauSac: String(row[4]),
        DungLuong: String(row[5]),
        TinhTrang: String(row[6]),
        GiaNhap: Number(row[7]) || 0,
        GiaBan: Number(row[8]) || 0,
        GiaTraGop: Number(row[9]) || 0,
        TrangThaiKho: String(row[10]),
        GhiChu: String(row[11]),
        ChiNhanh: String(row[12] || ""),
        NgayNhap: row[13] ? new Date(row[13]) : null,
        NgayXuat: row[14] ? new Date(row[14]) : null,
      });
    }
  });

  return result;
}

/**
 * Chuẩn hóa và tự động điền Ngày nhập & Ngày xuất cho các dữ liệu điện thoại đã có sẵn
 */
function backfillDienThoaiDates(silent) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
  if (!dtSheet) return;

  // 1. Đảm bảo sheet Điện thoại có đủ ít nhất 15 cột (đến cột O)
  var maxCols = dtSheet.getMaxColumns();
  if (maxCols < 15) {
    dtSheet.insertColumnsAfter(maxCols, 15 - maxCols);
  }

  // 2. Kiểm tra và ghi tiêu đề cột N và O nếu chưa có hoặc sai tiêu đề
  var headerRange = dtSheet.getRange(1, 14, 1, 2);
  var headerValues = headerRange.getValues()[0];
  var needsHeaderStyle = false;

  if (String(headerValues[0]).trim() !== "Ngày nhập") {
    dtSheet.getRange(1, 14).setValue("Ngày nhập");
    needsHeaderStyle = true;
  }
  if (String(headerValues[1]).trim() !== "Ngày xuất") {
    dtSheet.getRange(1, 15).setValue("Ngày xuất");
    needsHeaderStyle = true;
  }

  if (needsHeaderStyle) {
    var newHeadersRange = dtSheet.getRange(1, 14, 1, 2);
    newHeadersRange.setBackground("#1a73e8");
    newHeadersRange.setFontColor("#ffffff");
    newHeadersRange.setFontWeight("bold");
    newHeadersRange.setHorizontalAlignment("center");
  }

  var lastRow = dtSheet.getLastRow();
  if (lastRow <= 1) {
    if (!silent) showAlert("Thông báo", "Sheet Điện thoại chưa có dòng dữ liệu nào để chuẩn hóa!");
    return;
  }

  // 3. Đọc dữ liệu 15 cột
  var dtData = dtSheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var orders = getAllData(SHEET_NAMES.DON_HANG);
  var imports = getAllData(SHEET_NAMES.NHAP_KHO);
  var buybacks = getAllData(SHEET_NAMES.THU_MUA);

  var updatedCount = 0;

  for (var i = 0; i < dtData.length; i++) {
    var maDT = String(dtData[i][0]);
    var imei = String(dtData[i][3]);
    var status = String(dtData[i][10]);
    
    var currentNgayNhap = dtData[i][13];
    var currentNgayXuat = dtData[i][14];
    var changed = false;

    // A. Tìm Ngày nhập (nếu trống)
    if (!currentNgayNhap) {
      var ngayNhap = null;
      // Thử tìm trong Nhập kho theo Mã điện thoại (cột 4, index 3)
      for (var j = 0; j < imports.length; j++) {
        if (String(imports[j][3]) === maDT) {
          ngayNhap = imports[j][1]; // Ngày nhập (cột 2, index 1)
          break;
        }
      }
      // Nếu không thấy, tìm trong Thu mua theo IMEI (cột 8, index 7)
      if (!ngayNhap && imei) {
        for (var k = 0; k < buybacks.length; k++) {
          if (String(buybacks[k][7]) === imei) {
            ngayNhap = buybacks[k][1]; // Ngày thu mua (cột 2, index 1)
            break;
          }
        }
      }
      
      if (ngayNhap) {
        dtData[i][13] = new Date(ngayNhap);
      } else {
        dtData[i][13] = new Date(); // Mặc định là ngày hiện tại nếu không tìm thấy lịch sử
      }
      changed = true;
    }

    // B. Tìm/Xóa Ngày xuất
    if (status === "Đã bán" || status === "Đang trả góp") {
      if (!currentNgayXuat) {
        var ngayXuat = null;
        // Tìm trong Đơn hàng theo Mã điện thoại (cột 5, index 4)
        for (var j = 0; j < orders.length; j++) {
          if (String(orders[j][4]) === maDT) {
            var orderStatus = String(orders[j][18]);
            if (orderStatus !== "Huỷ" && orderStatus !== "Đổi trả") {
              ngayXuat = orders[j][1]; // Ngày bán (cột 2, index 1)
              break;
            }
          }
        }
        if (ngayXuat) {
          dtData[i][14] = new Date(ngayXuat);
          changed = true;
        }
      }
    } else {
      // Nếu trạng thái khác (Còn hàng, Đã trả lại) mà có ngày xuất, thực hiện xoá ngày xuất
      if (currentNgayXuat) {
        dtData[i][14] = "";
        changed = true;
      }
    }

    if (changed) {
      updatedCount++;
    }
  }

  // 4. Ghi ngược kết quả bằng Batch Write để tối ưu hóa hiệu suất
  var datesToUpdate = dtData.map(function (row) {
    return [
      row[13] ? new Date(row[13]) : "",
      row[14] ? new Date(row[14]) : "",
    ];
  });

  var updateRange = dtSheet.getRange(2, 14, lastRow - 1, 2);
  updateRange.setValues(datesToUpdate);

  // 5. Định dạng lại cột hiển thị
  updateRange.setNumberFormat("dd/MM/yyyy");
  updateRange.setFontFamily("Times New Roman");
  updateRange.setFontSize(12);
  updateRange.setHorizontalAlignment("center");

  // Tự động căn chỉnh lại độ rộng cột N & O
  dtSheet.autoResizeColumn(14);
  dtSheet.autoResizeColumn(15);

  if (!silent) {
    showAlert(
      "Thành công",
      "Đã chuẩn hóa cấu trúc cột và cập nhật " + updatedCount + " dòng dữ liệu Điện thoại cũ thành công!",
    );
  }
}

/**
 * Build cache dữ liệu điện thoại còn hàng cho tìm kiếm nhanh
 * @private
 */
function _buildDienThoaiDropdownCache() {
  var data = getAllData(SHEET_NAMES.DIEN_THOAI);
  var result = [];

  data.forEach(function (row) {
    if (
      row[0] &&
      String(row[0]).trim() !== "" &&
      String(row[10]) === "Còn hàng"
    ) {
      var cn = String(row[12] || "");
      result.push({
        v: String(row[0]),
        t: row[0] + " - " + row[1] + " (" + row[4] + ", " + row[5] + " | IMEI: " + row[3] + " | " + cn + ")",
        th: String(row[2]),
        gb: Number(row[8]),
        gtg: Number(row[9]),
        cn: cn,
        _s: (String(row[0]) + " " + String(row[1]) + " " + String(row[3])).toLowerCase(),
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
  var kw = String(keyword).trim().toLowerCase();

  var allItems = getChunkedCache('dd_dt');
  if (!allItems) {
    allItems = _buildDienThoaiDropdownCache();
    setChunkedCache('dd_dt', allItems);
  }

  var result = [];
  for (var i = 0; i < allItems.length; i++) {
    var item = allItems[i];
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
      if (result.length >= 100) break;
    }
  }

  return result;
}


