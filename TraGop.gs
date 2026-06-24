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
  initializeColumnEnums();
  var maTG = generateId("TG", SHEET_NAMES.TRA_GOP);

  var tongTien = Number(data.tongTien) || 0;
  var traTruocNum = parseAmountVal(data.traTruoc);
  var conLai = tongTien - traTruocNum;
  var soKy = Number(data.soKy) || 1;

  var loaiTraGop = data.loaiTraGop || "Cửa hàng";
  var isCTTC = loaiTraGop === "Công ty tài chính";

  // Nếu là công ty tài chính, sử dụng số tiền đóng mỗi kỳ truyền lên (đã bao gồm lãi)
  // Nếu là cửa hàng, tính tiền mỗi kỳ = tiền gốc mỗi tháng + tiền lãi mỗi tháng
  var interestRate = getInterestRateConfig();
  var tienMoiKy = isCTTC
    ? Number(data.tienMoiKy) || 0
    : Math.ceil(conLai / soKy) + Math.round(conLai * (interestRate / 100));

  var tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, 1, data.maKH, 2) || "";

  var ngayBatDau = new Date();
  var ngayKetThuc = new Date();
  ngayKetThuc.setMonth(ngayKetThuc.getMonth() + soKy);

  var tienMat = 0;
  var chuyenKhoan = 0;
  if (typeof data.traTruoc === "string" && data.traTruoc.indexOf(",") !== -1) {
    var parts = data.traTruoc.split(",");
    chuyenKhoan = Number(parts[0]) || 0;
    tienMat = Number(parts[1]) || 0;
  } else {
    if (
      data.splitTienMat !== undefined ||
      data.splitChuyenKhoan !== undefined
    ) {
      tienMat = Number(data.splitTienMat) || 0;
      chuyenKhoan = Number(data.splitChuyenKhoan) || 0;
    } else {
      if (
        data.hinhThucThanhToan === "Chuyển khoản" ||
        data.hinhThucThanhToan === "Quẹt thẻ (POS)"
      ) {
        chuyenKhoan = traTruocNum;
      } else {
        tienMat = traTruocNum;
      }
    }
  }

  var rowData = [];
  rowData[COL_TG.MA_TG - 1] = maTG;
  rowData[COL_TG.MA_DH - 1] = data.maDH || "";
  rowData[COL_TG.MA_KH - 1] = data.maKH || "";
  rowData[COL_TG.TEN_KH - 1] = tenKH;
  rowData[COL_TG.TONG_TIEN - 1] = tongTien;
  rowData[COL_TG.TRA_TRUOC - 1] = traTruocNum;
  rowData[COL_TG.CON_LAI - 1] = conLai;
  rowData[COL_TG.SO_KY - 1] = soKy;
  rowData[COL_TG.TIEN_MO_KY - 1] = tienMoiKy;
  rowData[COL_TG.NGAY_BAT_DAU - 1] = ngayBatDau;
  rowData[COL_TG.NGAY_KET_THUC - 1] = ngayKetThuc;
  rowData[COL_TG.DA_TRA_SO_KY - 1] = isCTTC ? soKy : 0;
  rowData[COL_TG.DA_TRA_SO_TIEN - 1] = isCTTC ? tongTien : traTruocNum;
  rowData[COL_TG.LOAI_TRA_GOP - 1] = loaiTraGop;
  rowData[COL_TG.CONG_TY_TC - 1] = data.congTyTC || "";
  rowData[COL_TG.TRANG_THAI - 1] = isCTTC ? "Hoàn tất" : "Đang trả";
  rowData[COL_TG.CHI_NHANH - 1] = data.chiNhanh || "";
  rowData[COL_TG.TIEN_MAT - 1] = tienMat;
  rowData[COL_TG.CHUYEN_KHOAN - 1] = chuyenKhoan;

  appendRow(SHEET_NAMES.TRA_GOP, rowData);

  // Chỉ tạo lịch thanh toán từng kỳ cho Cửa hàng tự trả góp
  if (!isCTTC) {
    var historyRows = [];
    var baseNum = getNewIdCounter("LS", SHEET_NAMES.LICH_SU_TRA_GOP);
    for (var i = 1; i <= soKy; i++) {
      var nextNum = baseNum + i;
      var padded = ("00000" + nextNum).slice(-5);
      var maLS = "LS" + padded + getRandomLetters();
      var ngayCanTra = new Date(ngayBatDau);
      ngayCanTra.setMonth(ngayCanTra.getMonth() + i);

      var lichData = [];
      lichData[COL_LSTG.MA_LS - 1] = maLS;
      lichData[COL_LSTG.MA_TG - 1] = maTG;
      lichData[COL_LSTG.KY_SO - 1] = i;
      lichData[COL_LSTG.SO_TIEN_CAN_TRA - 1] = tienMoiKy;
      lichData[COL_LSTG.SO_TIEN_DA_TRA - 1] = 0;
      lichData[COL_LSTG.NGAY_CAN_TRA - 1] = ngayCanTra;
      lichData[COL_LSTG.NGAY_THUC_TRA - 1] = "";
      lichData[COL_LSTG.HINH_THUC_TT - 1] = "";
      lichData[COL_LSTG.TRANG_THAI - 1] = "Chưa trả";
      lichData[COL_LSTG.GHI_CHU - 1] = "";
      lichData[COL_LSTG.TIEN_MAT - 1] = 0;
      lichData[COL_LSTG.CHUYEN_KHOAN - 1] = 0;

      historyRows.push(lichData);
    }
    appendRows(SHEET_NAMES.LICH_SU_TRA_GOP, historyRows);
  }

  var msgToast = isCTTC
    ? "✅ Đã tạo HĐ trả góp qua CTTC: " + maTG
    : "✅ Đã tạo HĐ trả góp: " +
      maTG +
      "\n" +
      soKy +
      " kỳ × " +
      formatCurrency(tienMoiKy) +
      "đ/kỳ";
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
  return withDocumentLock(function () {
    clearSheetCache();
    initializeColumnEnums();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var soTien = Number(data.soTien) || 0;
    var splitResult = calculatePaymentSplit(data, soTien);
    var tienMat = splitResult.tienMat;
    var chuyenKhoan = splitResult.chuyenKhoan;

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

    // Cập nhật lịch sử hàng loạt trong 1 range
    var lastCol = lstgSheet.getLastColumn();
    var range = lstgSheet.getRange(targetRow, 1, 1, lastCol);
    var rowValues = range.getValues()[0];

    rowValues[COL_LSTG.SO_TIEN_DA_TRA - 1] = soTien;
    rowValues[COL_LSTG.NGAY_THUC_TRA - 1] = new Date();
    rowValues[COL_LSTG.HINH_THUC_TT - 1] = splitResult.hinhThucTTDisplay;
    rowValues[COL_LSTG.TRANG_THAI - 1] = "Đã trả";
    if (data.ghiChu) {
      rowValues[COL_LSTG.GHI_CHU - 1] = data.ghiChu;
    }
    rowValues[COL_LSTG.TIEN_MAT - 1] = tienMat;
    rowValues[COL_LSTG.CHUYEN_KHOAN - 1] = chuyenKhoan;

    range.setValues([rowValues]);

    clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);

    // Cập nhật thông tin tổng trong TraGop
    _capNhatTongTraGop(data.maTG);

    showToast(
      "✅ Đã ghi nhận thanh toán kỳ " +
        data.kySo +
        " - HĐ " +
        data.maTG +
        "\nSố tiền: " +
        formatCurrency(soTien) +
        "đ",
    );
    return true;
  });
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
  initializeColumnEnums();
  var tgRow = findRow(SHEET_NAMES.TRA_GOP, COL_TG.MA_TG, maTG);
  if (tgRow === -1) return;

  var tgSheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
  var lastCol = tgSheet.getLastColumn();
  var range = tgSheet.getRange(tgRow, 1, 1, lastCol);
  var rowValues = range.getValues()[0];

  var traTruoc = parseAmountVal(rowValues[COL_TG.TRA_TRUOC - 1]);
  var soKy = Number(rowValues[COL_TG.SO_KY - 1]) || 0;

  rowValues[COL_TG.DA_TRA_SO_KY - 1] = daTraSoKy;
  rowValues[COL_TG.DA_TRA_SO_TIEN - 1] = traTruoc + daTraSoTien;

  // Cập nhật trạng thái
  if (daTraSoKy >= soKy) {
    rowValues[COL_TG.TRANG_THAI - 1] = "Hoàn tất";
    
    // Cập nhật trạng thái kho điện thoại → Đã bán
    var maDH = rowValues[COL_TG.MA_DH - 1];
    var maSP = lookupValue(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH, COL_DH.MA_SP);
    var nguonSP = lookupValue(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH, COL_DH.NGUON_SP);
    if (nguonSP === "Điện thoại" && maSP) {
      var ghiChuDH = lookupValue(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH, COL_DH.GHI_CHU) || "";
      var imeiMatch = ghiChuDH.match(/\[IMEI:\s*([^\s\]]+)\]/);
      var imei = imeiMatch ? imeiMatch[1] : "";
      updateTrangThaiKhoDT(imei || maSP, "Đã bán");
    }
  }

  range.setValues([rowValues]);

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
        t:
          row[0] +
          " - " +
          row[3] +
          " - Đang trả (" +
          row[11] +
          "/" +
          row[7] +
          " kỳ)",
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

  var allItems = getChunkedCache("dd_tg");
  if (!allItems) {
    allItems = _buildTraGopDropdownCache();
    setChunkedCache("dd_tg", allItems);
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
