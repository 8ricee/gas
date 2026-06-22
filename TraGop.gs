/**
 * ============================================================
 * VanTran Mobile — TraGop.gs
 * Quản lý hợp đồng trả góp + lịch sử thanh toán từng kỳ
 * Hỗ trợ: Cửa hàng tự cho trả góp + Công ty tài chính
 * ============================================================
 */

/**
 * Tạo hợp đồng trả góp + tự động sinh lịch thanh toán từng kỳ
 *
 * @param {Object} data - {maDH, maKH, tongTien, traTruoc, soKy, loaiTraGop, congTyTC, chiNhanh}
 * @return {string} Mã trả góp mới
 */
function taoHopDongTraGop(data) {
  var maTG = generateId("TG", SHEET_NAMES.TRA_GOP);

  var tongTien = Number(data.tongTien) || 0;
  var traTruoc = data.traTruoc;
  var conLai = tongTien - parseAmountVal(traTruoc);
  var soKy = Number(data.soKy) || 1;
  
  var loaiTraGop = data.loaiTraGop || "Cửa hàng";
  var isCTTC = (loaiTraGop === "Công ty tài chính");

  // Nếu là công ty tài chính, sử dụng số tiền đóng mỗi kỳ truyền lên (đã bao gồm lãi)
  var tienMoiKy = isCTTC ? (Number(data.tienMoiKy) || 0) : Math.ceil(conLai / soKy);

  var tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, 1, data.maKH, 2) || "";

  var ngayBatDau = new Date();
  var ngayKetThuc = new Date();
  ngayKetThuc.setMonth(ngayKetThuc.getMonth() + soKy);

  var rowData = [
    maTG,
    data.maDH || "",
    data.maKH || "",
    tenKH,
    tongTien,
    traTruoc,
    conLai,
    soKy,
    tienMoiKy,
    ngayBatDau,
    ngayKetThuc,
    isCTTC ? soKy : 0, // DaTraSoKy: Nếu CTTC, ghi nhận hoàn tất ngay
    isCTTC ? tongTien : traTruoc, // DaTraSoTien: Nếu CTTC, ghi nhận nhận đủ 100% (cả phần CTTC giải ngân)
    loaiTraGop,
    data.congTyTC || "",
    isCTTC ? "Hoàn tất" : "Đang trả", // TrangThai
    data.chiNhanh || "", // 17th column
  ];

  appendRow(SHEET_NAMES.TRA_GOP, rowData);

  // Chỉ tạo lịch thanh toán từng kỳ cho Cửa hàng tự trả góp
  if (!isCTTC) {
    for (var i = 1; i <= soKy; i++) {
      var maLS = generateId("LS", SHEET_NAMES.LICH_SU_TRA_GOP);
      var ngayCanTra = new Date(ngayBatDau);
      ngayCanTra.setMonth(ngayCanTra.getMonth() + i);

      var lichData = [
        maLS,
        maTG,
        i, // KySo
        tienMoiKy, // SoTienCanTra
        0, // SoTienDaTra
        ngayCanTra, // NgayCanTra
        "", // NgayThucTra
        "", // HinhThucThanhToan
        "Chưa trả", // TrangThai
        "", // GhiChu
      ];

      appendRow(SHEET_NAMES.LICH_SU_TRA_GOP, lichData);
    }
  }

  var msgToast = isCTTC 
    ? "✅ Đã tạo HĐ trả góp qua CTTC: " + maTG
    : "✅ Đã tạo HĐ trả góp: " + maTG + "\n" + soKy + " kỳ × " + formatCurrency(tienMoiKy) + "đ/kỳ";
  showToast(msgToast);
  return maTG;
}

/**
 * Ghi nhận thanh toán 1 kỳ trả góp
 *
 * @param {Object} data - {maTG, kySo, soTien, hinhThucThanhToan, ghiChu}
 * @return {boolean}
 */
function ghiNhanThanhToan(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Tìm dòng lịch sử trả góp
  var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
  if (!lstgSheet) {
    showAlert("❌ Lỗi", "Không tìm thấy bảng Lịch sử trả góp!");
    return false;
  }

  var lastRow = lstgSheet.getLastRow();
  if (lastRow <= 1) {
    showAlert("❌ Lỗi", "Không có dữ liệu lịch sử trả góp!");
    return false;
  }

  var allData = lstgSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var targetRow = -1;

  for (var i = 0; i < allData.length; i++) {
    if (
      String(allData[i][1]) === data.maTG &&
      Number(allData[i][2]) === Number(data.kySo)
    ) {
      targetRow = i + 2;
      break;
    }
  }

  if (targetRow === -1) {
    showAlert(
      "❌ Lỗi",
      "Không tìm thấy kỳ " + data.kySo + " của HĐ " + data.maTG,
    );
    return false;
  }

  var soTien = (data.hinhThucThanhToan === "Hỗn hợp" && data.splitChuyenKhoan !== undefined)
    ? (data.splitChuyenKhoan + "," + data.splitTienMat)
    : (Number(data.soTien) || 0);

  var displayAmount = parseAmountVal(soTien);

  // Cập nhật lịch sử
  lstgSheet.getRange(targetRow, 5).setValue(soTien); // SoTienDaTra
  lstgSheet.getRange(targetRow, 7).setValue(new Date()); // NgayThucTra
  lstgSheet
    .getRange(targetRow, 8)
    .setValue(data.hinhThucThanhToan || "Tiền mặt");
  lstgSheet.getRange(targetRow, 9).setValue("Đã trả");
  if (data.ghiChu) {
    lstgSheet.getRange(targetRow, 10).setValue(data.ghiChu);
  }
  clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);

  // Cập nhật thông tin tổng trong TraGop
  _capNhatTongTraGop(data.maTG);

  showToast(
    "✅ Đã ghi nhận thanh toán kỳ " +
      data.kySo +
      " - HĐ " +
      data.maTG +
      "\nSố tiền: " +
      formatCurrency(displayAmount) +
      "đ",
  );
  return true;
}

/**
 * Cập nhật tổng số kỳ đã trả và tổng tiền đã trả trong sheet TraGop
 *
 * @param {string} maTG - Mã trả góp
 * @private
 */
function _capNhatTongTraGop(maTG) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Đếm kỳ đã trả và tổng tiền
  var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
  var lastRow = lstgSheet.getLastRow();
  if (lastRow <= 1) return;

  var allData = lstgSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var daTraSoKy = 0;
  var daTraSoTien = 0;

  allData.forEach(function (row) {
    if (String(row[1]) === maTG && String(row[8]) === "Đã trả") {
      daTraSoKy++;
      daTraSoTien += parseAmountVal(row[4]);
    }
  });

  // Tìm dòng trả góp
  var tgRow = findRow(SHEET_NAMES.TRA_GOP, 1, maTG);
  if (tgRow === -1) return;

  var tgSheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
  var traTruoc = parseAmountVal(tgSheet.getRange(tgRow, 6).getValue());
  var soKy = Number(tgSheet.getRange(tgRow, 8).getValue()) || 0;

  tgSheet.getRange(tgRow, 12).setValue(daTraSoKy); // DaTraSoKy
  tgSheet.getRange(tgRow, 13).setValue(traTruoc + daTraSoTien); // DaTraSoTien (cộng tiền trả trước)

  // Cập nhật trạng thái
  if (daTraSoKy >= soKy) {
    tgSheet.getRange(tgRow, 16).setValue("Hoàn tất");

    // Cập nhật trạng thái kho điện thoại → Đã bán
    var maDH = tgSheet.getRange(tgRow, 2).getValue();
    var maSP = lookupValue(SHEET_NAMES.DON_HANG, 1, maDH, 5);
    var nguonSP = lookupValue(SHEET_NAMES.DON_HANG, 1, maDH, 7);
    if (nguonSP === "Điện thoại" && maSP) {
      var ghiChuDH = lookupValue(SHEET_NAMES.DON_HANG, 1, maDH, 20) || "";
      var imeiMatch = ghiChuDH.match(/\[IMEI:\s*([^\s\]]+)\]/);
      var imei = imeiMatch ? imeiMatch[1] : "";
      updateTrangThaiKhoDT(imei || maSP, "Đã bán");
    }
  }
  clearSheetCache(SHEET_NAMES.TRA_GOP);
  invalidateDropdownCache(SHEET_NAMES.TRA_GOP);
}

/**
 * Huỷ hợp đồng trả góp
 *
 * @param {string} maDH - Mã đơn hàng
 */
function huyHopDongTraGop(maDH) {
  var maTG = lookupValue(SHEET_NAMES.TRA_GOP, 2, maDH, 1);
  if (maTG) {
    var tgRow = findRow(SHEET_NAMES.TRA_GOP, 1, maTG);
    if (tgRow !== -1) {
      updateCell(SHEET_NAMES.TRA_GOP, tgRow, 16, "Đã huỷ");
      
      // Cập nhật trạng thái các kỳ chưa thanh toán trong Lịch sử trả góp thành "Đã huỷ"
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
      if (lstgSheet) {
        var lastRow = lstgSheet.getLastRow();
        if (lastRow > 1) {
          var allLichSu = lstgSheet.getRange(2, 2, lastRow - 1, 8).getValues(); // Cột B (MaTG) đến I (TrangThai)
          for (var i = 0; i < allLichSu.length; i++) {
            if (String(allLichSu[i][0]) === maTG) {
              var status = String(allLichSu[i][7]);
              if (status !== "Đã trả") {
                lstgSheet.getRange(i + 2, 9).setValue("Đã huỷ"); // Cột I (Cột 9)
              }
            }
          }
          clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);
        }
      }
    }
  }
}

/**
 * Lấy danh sách trả góp quá hạn
 *
 * @return {Object[]} Danh sách kỳ trả góp đã quá hạn
 */
function getTraGopQuaHan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
  if (!lstgSheet) return [];

  var lastRow = lstgSheet.getLastRow();
  if (lastRow <= 1) return [];

  var allData = lstgSheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var result = [];

  allData.forEach(function (row, index) {
    var trangThai = String(row[8]);
    var ngayCanTra = row[5];

    if (trangThai === "Chưa trả" && ngayCanTra instanceof Date) {
      ngayCanTra.setHours(0, 0, 0, 0);
      if (ngayCanTra < today) {
        // Đánh dấu quá hạn
        lstgSheet.getRange(index + 2, 9).setValue("Quá hạn");

        var obj = {
          MaLS: String(row[0]),
          MaTG: String(row[1]),
          KySo: Number(row[2]),
          SoTienCanTra: Number(row[3]),
          SoTienDaTra: Number(row[4]),
          NgayCanTra: row[5],
          NgayThucTra: row[6],
          HinhThucThanhToan: String(row[7]),
          TrangThai: "Quá hạn",
          GhiChu: String(row[9]),
        };

        // Thêm thông tin KH
        var maTG = String(row[1]);
        obj.TenKH = lookupValue(SHEET_NAMES.TRA_GOP, 1, maTG, 4) || "";
        obj.SoNgayQuaHan = Math.floor(
          (today - ngayCanTra) / (1000 * 60 * 60 * 24),
        );

        result.push(obj);
      }
    }
  });

  return result;
}

/**
 * Xem tình trạng trả góp
 *
 * @param {string} maTG - Mã trả góp
 * @return {Object|null}
 */
function getTrangThaiTraGop(maTG) {
  var row = findRow(SHEET_NAMES.TRA_GOP, 1, maTG);
  if (row === -1) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
  var r = sheet.getRange(row, 1, 1, 17).getValues()[0];

  var result = {
    MaTG: String(r[0]),
    MaDH: String(r[1]),
    MaKH: String(r[2]),
    TenKH: String(r[3]),
    TongTien: Number(r[4]) || 0,
    TraTruoc: Number(r[5]) || 0,
    ConLai: Number(r[6]) || 0,
    SoKy: Number(r[7]) || 0,
    TienMoiKy: Number(r[8]) || 0,
    NgayBatDau: r[9],
    NgayKetThuc: r[10],
    DaTraSoKy: Number(r[11]) || 0,
    DaTraSoTien: Number(r[12]) || 0,
    LoaiTraGop: String(r[13]),
    CongTyTC: String(r[14]),
    TrangThai: String(r[15]),
    ChiNhanh: String(r[16] || ""),
  };

  // Lấy lịch sử thanh toán
  var allLSTG = getAllData(SHEET_NAMES.LICH_SU_TRA_GOP);
  result.lichSu = [];

  allLSTG.forEach(function (rowVals) {
    if (String(rowVals[1]) === maTG) {
      result.lichSu.push({
        MaLS: String(rowVals[0]),
        MaTG: String(rowVals[1]),
        KySo: Number(rowVals[2]),
        SoTienCanTra: Number(rowVals[3]),
        SoTienDaTra: Number(rowVals[4]),
        NgayCanTra: rowVals[5],
        NgayThucTra: rowVals[6],
        HinhThucThanhToan: String(rowVals[7]),
        TrangThai: String(rowVals[8]),
        GhiChu: String(rowVals[9]),
      });
    }
  });

  return result;
}

/**
 * Lấy danh sách hợp đồng trả góp cho dropdown
 *
 * @return {Object[]} [{value: 'TG001', text: 'TG001 - Nguyễn Văn A - Đang trả (5/12 kỳ)'}, ...]
 */
function getTraGopDropdown() {
  var data = getAllData(SHEET_NAMES.TRA_GOP);
  var result = [];

  data.forEach(function (row) {
    if (
      row[0] &&
      String(row[0]).trim() !== "" &&
      String(row[15]) === "Đang trả" &&
      String(row[13]) === "Cửa hàng"
    ) {
      result.push({
        value: String(row[0]),
        text:
          row[0] +
          " - " +
          row[3] +
          " - Đang trả (" +
          row[11] +
          "/" +
          row[7] +
          " kỳ)",
        soKy: Number(row[7]),
        daTraSoKy: Number(row[11]),
        tienMoiKy: Number(row[8]),
      });
    }
  });

  return result;
}

/**
 * Lấy danh sách kỳ chưa trả của 1 hợp đồng
 *
 * @param {string} maTG - Mã trả góp
 * @return {Object[]} Các kỳ chưa trả
 */
function getKyChuaTra(maTG) {
  var allLSTG = getAllData(SHEET_NAMES.LICH_SU_TRA_GOP);
  var result = [];

  allLSTG.forEach(function (row) {
    if (String(row[1]) === maTG && String(row[8]) !== "Đã trả") {
      result.push({
        value: Number(row[2]),
        text:
          "Kỳ " +
          row[2] +
          " - " +
          formatCurrency(row[3]) +
          "đ - Hạn: " +
          formatDate(row[5]),
        soTienCanTra: Number(row[3]),
        trangThai: String(row[8]),
      });
    }
  });

  return result;
}

/**
 * Build cache hợp đồng trả góp đang hoạt động cho tìm kiếm nhanh
 * @private
 */
function _buildTraGopDropdownCache() {
  var data = getAllData(SHEET_NAMES.TRA_GOP);
  var result = [];

  data.forEach(function (row) {
    if (
      row[0] &&
      String(row[0]).trim() !== "" &&
      String(row[15]) === "Đang trả" &&
      String(row[13]) === "Cửa hàng"
    ) {
      result.push({
        v: String(row[0]),
        t: row[0] + " - " + row[3] + " - Đang trả (" + row[11] + "/" + row[7] + " kỳ)",
        sk: Number(row[7]),
        dtsk: Number(row[11]),
        tmk: Number(row[8]),
        _s: (String(row[0]) + " " + String(row[3])).toLowerCase(),
      });
    }
  });

  return result;
}

/**
 * Lấy danh sách hợp đồng trả góp hoạt động theo từ khóa tìm kiếm (Mã TG, Tên khách hàng)
 * Sử dụng CacheService để tránh đọc sheet mỗi lần tìm kiếm
 *
 * @param {string} keyword - Từ khóa tìm
 * @return {Object[]}
 */
function getTraGopDropdownSearch(keyword) {
  var kw = String(keyword).trim().toLowerCase();

  var allItems = getChunkedCache('dd_tg');
  if (!allItems) {
    allItems = _buildTraGopDropdownCache();
    setChunkedCache('dd_tg', allItems);
  }

  var result = [];
  for (var i = 0; i < allItems.length; i++) {
    if (allItems[i]._s.indexOf(kw) !== -1) {
      result.push({
        value: allItems[i].v,
        text: allItems[i].t,
        soKy: allItems[i].sk,
        daTraSoKy: allItems[i].dtsk,
        tienMoiKy: allItems[i].tmk,
      });
      if (result.length >= 100) break;
    }
  }

  return result;
}


