/**
 * ============================================================
 * VanTran Mobile — DoanhSo.gs
 * Chốt doanh số cuối tháng + Báo cáo
 * ============================================================
 */

/**
 * Chốt doanh số cuối tháng
 *
 * @param {number} thang - Tháng (1-12)
 * @param {number} nam - Năm
 * @return {boolean}
 */
function chotDoanhSoThang(thang, nam) {
  var thangNam = ("0" + thang).slice(-2) + "/" + nam;

  // Kiểm tra đã chốt chưa
  var daChot = lookupValue(SHEET_NAMES.DOANH_SO, 1, thangNam, 16);
  if (daChot === "Đã chốt") {
    try {
      var ui = SpreadsheetApp.getUi();
      if (ui) {
        var result = ui.alert(
          "Cảnh báo",
          "Doanh số tháng " +
            thangNam +
            " đã được chốt trước đó.\n\nBạn có muốn chốt lại (xóa dữ liệu cũ)?",
          ui.ButtonSet.YES_NO,
        );
        if (result !== ui.Button.YES) return false;
      }
    } catch (e) {
      Logger.log(
        "Doanh số tháng " +
          thangNam +
          " đã được chốt trước đó. Tự động ghi đè trong môi trường không có UI.",
      );
    }

    // Xóa dữ liệu cũ
    _xoaDoanhSoThang(thangNam);
  }

  // Lấy danh sách NV đang làm
  var allNV = getAllData(SHEET_NAMES.NHAN_VIEN);

  // Lấy đơn hàng trong tháng (chỉ điện thoại, hoàn thành)
  var donHangs = getDonHangTheoThang(thang, nam);
  var dtDonHangs = donHangs.filter(function (dh) {
    return (
      dh.NguonSP === "Điện thoại" &&
      dh.TrangThai !== "Huỷ" &&
      dh.TrangThai !== "Đổi trả"
    );
  });

  // Lấy bảo hành trong tháng
  var baoHanhs = typeof getBaoHanhTheoThang === 'function' ? getBaoHanhTheoThang(thang, nam) : [];

  // Lấy cấu hình hoa hồng
  var hhBanApple = getConfigNumber("HH Bán máy - Apple");
  var hhHTApple = getConfigNumber("HH Hỗ trợ - Apple");
  var hhBanKhac = getConfigNumber("HH Bán máy - Khác");
  var hhHTKhac = getConfigNumber("HH Hỗ trợ - Khác");

  // Tổng hợp doanh số theo từng NV
  var nvDoanhSo = {};

  allNV.forEach(function (nv) {
    if (String(nv[7]) === "Nghỉ việc") return;

    var maNV = String(nv[0]);
    nvDoanhSo[maNV] = {
      tenNV: nv[1],
      vaiTro: String(nv[4]),
      coQuyenXuatMay: String(nv[5]) === "✓",
      soMayBan_Apple: 0,
      soMayBan_Khac: 0,
      soMayHoTro_Apple: 0,
      soMayHoTro_Khac: 0,
    };
  });

  // Quét đơn hàng và phân bổ
  dtDonHangs.forEach(function (dh) {
    var nguoiBan = String(dh.NguoiBan);
    var nguoiHoTro = String(dh.NguoiHoTro);
    var laApple = isApple(dh.ThuongHieu);

    // Tính cho người bán
    if (nguoiBan && nvDoanhSo[nguoiBan]) {
      if (laApple) {
        nvDoanhSo[nguoiBan].soMayBan_Apple++;
      } else {
        nvDoanhSo[nguoiBan].soMayBan_Khac++;
      }
    }

    // Tính cho người hỗ trợ
    if (nguoiHoTro && nvDoanhSo[nguoiHoTro]) {
      if (laApple) {
        nvDoanhSo[nguoiHoTro].soMayHoTro_Apple++;
      } else {
        nvDoanhSo[nguoiHoTro].soMayHoTro_Khac++;
      }
    }
  });

  // Ghi vào sheet DoanhSo
  var rows = [];

  Object.keys(nvDoanhSo).forEach(function (maNV) {
    var nv = nvDoanhSo[maNV];

    // Tính hoa hồng bán (chỉ khi có quyền xuất máy)
    var hhBan = 0;
    if (nv.coQuyenXuatMay) {
      hhBan = nv.soMayBan_Apple * hhBanApple + nv.soMayBan_Khac * hhBanKhac;
    }

    // Hoa hồng hỗ trợ (luôn được tính)
    var hhHoTro =
      nv.soMayHoTro_Apple * hhHTApple + nv.soMayHoTro_Khac * hhHTKhac;

    var tongHoaHong = hhBan + hhHoTro;

    // Doanh thu dịch vụ
    var doanhThuDV = getTongPhiDichVu(maNV, thang, nam);
    baoHanhs.forEach(function(bh) {
      if (String(bh.NguoiTiepNhan) === maNV && bh.TrangThai !== "Huỷ") {
        doanhThuDV += Number(bh.PhiSuaChua) || 0;
      }
    });

    // Chỉ ghi nếu có doanh số hoặc dịch vụ
    if (
      nv.soMayBan_Apple +
        nv.soMayBan_Khac +
        nv.soMayHoTro_Apple +
        nv.soMayHoTro_Khac >
        0 ||
      doanhThuDV > 0
    ) {
      rows.push([
        thangNam,
        maNV,
        nv.tenNV,
        nv.vaiTro,
        nv.coQuyenXuatMay ? "✓" : "✗",
        nv.soMayBan_Apple,
        nv.soMayBan_Khac,
        nv.soMayHoTro_Apple,
        nv.soMayHoTro_Khac,
        hhBan,
        hhHoTro,
        tongHoaHong,
        doanhThuDV,
        0, // Thuong (nhập thủ công)
        tongHoaHong, // TongThuNhap (chưa tính thưởng)
        "Đã chốt",
      ]);
    }
  });

  // Ghi dữ liệu
  if (rows.length > 0) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.DOANH_SO);
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

    // Format tiền
    var colsTienStart = 10; // Hoa hồng bán
    for (var c = colsTienStart; c <= 15; c++) {
      sheet.getRange(startRow, c, rows.length, 1).setNumberFormat("#,##0");
    }
  }

  showAlert(
    "✅ Chốt doanh số thành công!",
    "Tháng: " +
      thangNam +
      "\nSố NV có doanh số: " +
      rows.length +
      "\nTổng đơn hàng ĐT: " +
      dtDonHangs.length +
      "\n\nVui lòng kiểm tra sheet Doanh số để xem chi tiết.",
  );

  return true;
}

/**
 * Xóa doanh số cũ của 1 tháng
 * @private
 */
function _xoaDoanhSoThang(thangNam) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.DOANH_SO);
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  // Xóa từ dưới lên để tránh lệch index
  for (var i = data.length - 1; i >= 0; i--) {
    if (formatMonthYear(data[i][0]) === thangNam) {
      sheet.deleteRow(i + 2);
    }
  }
}

/**
 * Lấy báo cáo doanh số tháng
 *
 * @param {number} thang
 * @param {number} nam
 * @return {Object}
 */
function getBaoCaoDoanhSo(thang, nam) {
  var thangNam = ("0" + thang).slice(-2) + "/" + nam;
  var data = getAllData(SHEET_NAMES.DOANH_SO);
  var result = {
    thangNam: thangNam,
    nhanVien: [],
    tongSoMayBan: 0,
    tongHoaHong: 0,
    tongDoanhThuDV: 0,
    tongThuNhap: 0,
  };

  data.forEach(function (row) {
    if (formatMonthYear(row[0]) === thangNam) {
      var nv = {
        maNV: String(row[1]),
        tenNV: String(row[2]),
        vaiTro: String(row[3]),
        coQuyenXuatMay: String(row[4]) === "✓",
        soMayBan_Apple: Number(row[5]) || 0,
        soMayBan_Khac: Number(row[6]) || 0,
        soMayHoTro_Apple: Number(row[7]) || 0,
        soMayHoTro_Khac: Number(row[8]) || 0,
        hhBan: Number(row[9]) || 0,
        hhHoTro: Number(row[10]) || 0,
        tongHoaHong: Number(row[11]) || 0,
        doanhThuDV: Number(row[12]) || 0,
        thuong: Number(row[13]) || 0,
        tongThuNhap: Number(row[14]) || 0,
      };

      result.nhanVien.push(nv);
      result.tongSoMayBan += nv.soMayBan_Apple + nv.soMayBan_Khac;
      result.tongHoaHong += nv.tongHoaHong;
      result.tongDoanhThuDV += nv.doanhThuDV;
      result.tongThuNhap += nv.tongThuNhap;
    }
  });

  return result;
}

/**
 * Hàm gọi từ menu — chốt doanh số với dialog chọn tháng
 */
function menuChotDoanhSo() {
  try {
    var ui = SpreadsheetApp.getUi();
    if (!ui) return;
    var now = new Date();
    var defaultThang = now.getMonth() + 1;
    var defaultNam = now.getFullYear();

    var response = ui.prompt(
      "Chốt doanh số tháng",
      "Nhập tháng/năm cần chốt (VD: " + defaultThang + "/" + defaultNam + "):",
      ui.ButtonSet.OK_CANCEL,
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;

    var input = response.getResponseText().trim();
    var parts = input.split("/");

    if (parts.length !== 2) {
      showAlert(
        "❌ Lỗi",
        "Định dạng không đúng! Vui lòng nhập: tháng/năm (VD: 6/2026)",
      );
      return;
    }

    var thang = parseInt(parts[0], 10);
    var nam = parseInt(parts[1], 10);

    if (isNaN(thang) || thang < 1 || thang > 12 || isNaN(nam)) {
      showAlert("❌ Lỗi", "Tháng hoặc năm không hợp lệ!");
      return;
    }

    chotDoanhSoThang(thang, nam);
  } catch (e) {
    Logger.log("menuChotDoanhSo error: " + e.message);
  }
}

/**
 * Hàm gọi từ menu — xem báo cáo doanh số
 */
function menuXemBaoCao() {
  try {
    var ui = SpreadsheetApp.getUi();
    if (!ui) return;
    var now = new Date();
    var defaultThang = now.getMonth() + 1;
    var defaultNam = now.getFullYear();

    var response = ui.prompt(
      "Xem báo cáo doanh số",
      "Nhập tháng/năm cần xem (VD: " + defaultThang + "/" + defaultNam + "):",
      ui.ButtonSet.OK_CANCEL,
    );

    if (response.getSelectedButton() !== ui.Button.OK) return;

    var input = response.getResponseText().trim();
    var parts = input.split("/");

    if (parts.length !== 2) {
      showAlert("❌ Lỗi", "Định dạng không đúng!");
      return;
    }

    var thang = parseInt(parts[0], 10);
    var nam = parseInt(parts[1], 10);
    var baoCao = getBaoCaoDoanhSo(thang, nam);

    if (baoCao.nhanVien.length === 0) {
      showAlert(
        "Báo cáo tháng " + baoCao.thangNam,
        "Chưa có dữ liệu doanh số cho tháng này.\nVui lòng chốt doanh số trước.",
      );
      return;
    }

    var msg = "═══ BÁO CÁO DOANH SỐ THÁNG " + baoCao.thangNam + " ═══\n\n";
    msg += "Tổng máy bán: " + baoCao.tongSoMayBan + "\n";
    msg += "Tổng hoa hồng: " + formatCurrency(baoCao.tongHoaHong) + "đ\n";
    msg +=
      "Tổng DT dịch vụ: " + formatCurrency(baoCao.tongDoanhThuDV) + "đ\n\n";

    msg += "─── Chi tiết ───\n";
    baoCao.nhanVien.forEach(function (nv) {
      msg += "\n " + nv.tenNV + " (" + nv.maNV + ")";
      msg +=
        "\n   Bán: Apple " + nv.soMayBan_Apple + " + Khác " + nv.soMayBan_Khac;
      msg +=
        "\n   HT:  Apple " +
        nv.soMayHoTro_Apple +
        " + Khác " +
        nv.soMayHoTro_Khac;
      msg += "\n   HH:  " + formatCurrency(nv.tongHoaHong) + "đ";
      if (nv.doanhThuDV > 0) {
        msg += "\n   DV:  " + formatCurrency(nv.doanhThuDV) + "đ";
      }
      msg += "\n";
    });

    showAlert("Báo cáo doanh số", msg);
  } catch (e) {
    Logger.log("menuXemBaoCao error: " + e.message);
  }
}

/**
 * Cập nhật báo cáo doanh số từ các bộ lọc trên sheet Báo cáo doanh số
 */
function updateSalesReportFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_DOANH_SO);
  if (!sheet) return;

  var startDateVal = sheet.getRange(3, 2).getValue(); // B3
  var endDateVal = sheet.getRange(3, 4).getValue(); // D3
  var staffVal = sheet.getRange(3, 6).getValue(); // F3
  var gdFilterVal = sheet.getRange(3, 8).getValue(); // H3

  var startDate = _parseDate(startDateVal);
  var endDate = _parseDate(endDateVal);

  if (!startDate || !endDate) {
    sheet
      .getRange(4, 2)
      .setValue("Ngày không hợp lệ! Vui lòng nhập định dạng dd/MM/yyyy");
    return;
  }

  if (startDate > endDate) {
    sheet
      .getRange(4, 2)
      .setValue("Lỗi: Ngày bắt đầu không thể lớn hơn ngày kết thúc!");
    return;
  }

  sheet.getRange(4, 2).setValue("Đang tải báo cáo...");
  SpreadsheetApp.flush();

  try {
    generateSalesReportOnSheet(startDate, endDate, staffVal, gdFilterVal);
    var nowStr = Utilities.formatDate(
      new Date(),
      "Asia/Ho_Chi_Minh",
      "HH:mm:ss dd/MM/yyyy",
    );
    sheet.getRange(4, 2).setValue("Hoàn thành lúc: " + nowStr);
  } catch (e) {
    sheet.getRange(4, 2).setValue("Lỗi: " + e.message);
  }
}

/**
 * Tạo báo cáo doanh số trên sheet Báo cáo doanh số
 *
 * @param {Date} startDate - Ngày bắt đầu
 * @param {Date} endDate - Ngày kết thúc
 * @param {string} staffVal - Lựa chọn nhân viên (VD: "Tất cả" hoặc "NV001 - Nguyễn Văn A")
 * @param {string} gdFilterVal - Lọc loại giao dịch (VD: "Tất cả", "Đơn hàng (Bán máy)", v.v.)
 */
function generateSalesReportOnSheet(startDate, endDate, staffVal, gdFilterVal) {
  initializeColumnEnums();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_DOANH_SO);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.BAO_CAO_DOANH_SO);
  }

  // 1. Dọn dẹp dữ liệu cũ (từ dòng 6 trở đi)
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow >= 6) {
    sheet
      .getRange(6, 1, lastRow - 5, Math.max(lastCol, 12))
      .clearContent()
      .clearFormat();
  }

  // 2. Tính toán doanh số live
  var listNV = _tinhDoanhSoTuNgayDenNgay(startDate, endDate);

  // Lọc theo nhân viên nếu cần
  var targetMaNV = "tat_ca";
  if (staffVal && staffVal !== "Tất cả") {
    var parts = staffVal.split(" - ");
    if (parts.length > 0) {
      targetMaNV = parts[0].trim();
    }
  }

  var filteredListNV = listNV;
  if (targetMaNV !== "tat_ca") {
    filteredListNV = listNV.filter(function (nv) {
      return nv.maNV === targetMaNV;
    });
  }

  // 3. Tính toán tổng hợp
  var tongSoMayBan = 0;
  var tongHoaHong = 0;
  var tongDoanhThuDV = 0;
  var tongThuNhap = 0;

  filteredListNV.forEach(function (nv) {
    tongSoMayBan += nv.soMayBan_Apple + nv.soMayBan_Khac;
    tongHoaHong += nv.tongHoaHong;
    tongDoanhThuDV += nv.doanhThuDV;
    tongThuNhap += nv.tongThuNhap;
  });

  // 4. Ghi bảng tổng hợp chỉ tiêu tài chính (dòng 6 -> 10)
  var summaryTitleRange = sheet.getRange(6, 1, 1, 2);
  summaryTitleRange.merge();
  summaryTitleRange
    .setValue("TỔNG HỢP CHỈ TIÊU DOANH SỐ & COMMISSION")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var summaryData = [
    ["Chỉ tiêu", "Giá trị"],
    ["1. Tổng máy bán", tongSoMayBan],
    ["2. Tổng hoa hồng nhân viên", tongHoaHong],
    ["3. Tổng doanh thu dịch vụ", tongDoanhThuDV],
    ["4. Tổng thu nhập nhân viên", tongThuNhap],
  ];
  sheet.getRange(7, 1, summaryData.length, 2).setValues(summaryData);
  sheet
    .getRange(7, 1, 1, 2)
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("center");
  sheet.getRange(8, 2, summaryData.length - 1, 1).setNumberFormat("#,##0");
  sheet.getRange(8, 1, summaryData.length - 1, 2).setFontWeight("bold");
  sheet.getRange(11, 1, 1, 2).setFontColor("#1a73e8").setFontWeight("bold");
  sheet
    .getRange(6, 1, summaryData.length + 1, 2)
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#dadce0",
      SpreadsheetApp.BorderStyle.SOLID,
    );

  // 5. Ghi bảng chi tiết nhân viên (bắt đầu từ dòng 13)
  var startRow = 13;
  var detailTitleRange = sheet.getRange(startRow, 1, 1, 12);
  detailTitleRange.merge();
  var periodStr = formatDate(startDate) + " - " + formatDate(endDate);
  detailTitleRange
    .setValue("CHI TIẾT DOANH SỐ NHÂN VIÊN (" + periodStr + ")")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var detailHeaders = [
    "Mã nhân viên",
    "Tên nhân viên",
    "Vai trò",
    "Quyền xuất",
    "Bán Apple",
    "Bán Khác",
    "HT Apple",
    "HT Khác",
    "HH Bán máy",
    "HH Hỗ trợ",
    "Doanh thu DV",
    "Tổng thu nhập",
  ];
  sheet.getRange(startRow + 1, 1, 1, 12).setValues([detailHeaders]);
  sheet
    .getRange(startRow + 1, 1, 1, 12)
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("center");

  var totalRowIdx = startRow + 2;

  if (filteredListNV.length > 0) {
    var detailRows = [];
    var sumBanAP = 0,
      sumBanKhac = 0,
      sumHTAP = 0,
      sumHTKhac = 0;
    var sumHHBan = 0,
      sumHHHT = 0,
      sumDV = 0,
      sumTongTN = 0;

    filteredListNV.forEach(function (nv) {
      detailRows.push([
        nv.maNV,
        nv.tenNV,
        nv.vaiTro,
        nv.coQuyenXuatMay ? "✓" : "✗",
        nv.soMayBan_Apple,
        nv.soMayBan_Khac,
        nv.soMayHoTro_Apple,
        nv.soMayHoTro_Khac,
        nv.hhBan,
        nv.hhHoTro,
        nv.doanhThuDV,
        nv.tongThuNhap,
      ]);
      sumBanAP += nv.soMayBan_Apple;
      sumBanKhac += nv.soMayBan_Khac;
      sumHTAP += nv.soMayHoTro_Apple;
      sumHTKhac += nv.soMayHoTro_Khac;
      sumHHBan += nv.hhBan;
      sumHHHT += nv.hhHoTro;
      sumDV += nv.doanhThuDV;
      sumTongTN += nv.tongThuNhap;
    });

    sheet
      .getRange(startRow + 2, 1, detailRows.length, 12)
      .setValues(detailRows);
    sheet
      .getRange(startRow + 2, 5, detailRows.length, 8)
      .setNumberFormat("#,##0");

    // Dòng tổng cộng nhân viên
    totalRowIdx = startRow + 2 + detailRows.length;
    sheet.getRange(totalRowIdx, 1).setValue("Tổng cộng").setFontWeight("bold");
    sheet
      .getRange(totalRowIdx, 5, 1, 8)
      .setValues([
        [
          sumBanAP,
          sumBanKhac,
          sumHTAP,
          sumHTKhac,
          sumHHBan,
          sumHHHT,
          sumDV,
          sumTongTN,
        ],
      ]);
    sheet
      .getRange(totalRowIdx, 5, 1, 8)
      .setNumberFormat("#,##0")
      .setFontWeight("bold");
    sheet
      .getRange(totalRowIdx, 1, 1, 12)
      .setBackground("#f1f3f4")
      .setFontWeight("bold");

    sheet
      .getRange(startRow, 1, detailRows.length + 3, 12)
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#dadce0",
        SpreadsheetApp.BorderStyle.SOLID,
      );
  } else {
    var emptyRange = sheet.getRange(startRow + 2, 1, 1, 12);
    emptyRange.merge();
    emptyRange
      .setValue(
        "Không có dữ liệu doanh số nào phát sinh trong khoảng thời gian này.",
      )
      .setFontStyle("italic")
      .setHorizontalAlignment("center");

    sheet
      .getRange(startRow, 1, 3, 12)
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#dadce0",
        SpreadsheetApp.BorderStyle.SOLID,
      );
    totalRowIdx = startRow + 2;
  }

  // 6. Giao dịch chi tiết (bắt đầu từ dòng totalRowIdx + 3)
  var txStartRow = totalRowIdx + 3;
  var txTitleRange = sheet.getRange(txStartRow, 1, 1, 12);
  txTitleRange.merge();
  txTitleRange
    .setValue("DANH SÁCH GIAO DỊCH CHI TIẾT")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var txHeaders = [
    "Thời gian",
    "Mã giao dịch",
    "Nhân viên",
    "Vai trò",
    "Loại giao dịch",
    "Sản phẩm / Chi tiết",
    "Doanh thu / Số tiền",
    "Hoa hồng",
    "Phí dịch vụ",
    "Chi nhánh",
    "Khách hàng",
    "Ghi chú",
  ];
  sheet.getRange(txStartRow + 1, 1, 1, 12).setValues([txHeaders]);
  sheet
    .getRange(txStartRow + 1, 1, 1, 12)
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("center");

  // Thu thập giao dịch chi tiết
  var allNV = getAllData(SHEET_NAMES.NHAN_VIEN);
  var orders = getAllData(SHEET_NAMES.DON_HANG);
  var services = getAllData(SHEET_NAMES.DICH_VU);

  var nvMap = {};
  allNV.forEach(function (nv) {
    nvMap[String(nv[0])] = {
      tenNV: String(nv[1]),
      vaiTro: String(nv[4]),
      coQuyenXuatMay: String(nv[5]) === "✓",
    };
  });

  var hhBanApple = getConfigNumber("HH Bán máy - Apple");
  var hhHTApple = getConfigNumber("HH Hỗ trợ - Apple");
  var hhBanKhac = getConfigNumber("HH Bán máy - Khác");
  var hhHTKhac = getConfigNumber("HH Hỗ trợ - Khác");

  var txList = [];

  // A. Đơn hàng
  orders.forEach(function (row) {
    var ngayBan = row[COL_DH.NGAY_BAN - 1];
    if (!(ngayBan instanceof Date)) return;
    var status = String(row[COL_DH.TRANG_THAI - 1]);
    if (
      ngayBan >= startDate &&
      ngayBan <= endDate &&
      String(row[COL_DH.NGUON_SP - 1]) === "Điện thoại" &&
      status !== "Huỷ" &&
      status !== "Đổi trả"
    ) {
      var maDH = String(row[COL_DH.MA_DH - 1]);
      var tenKH = String(row[COL_DH.TEN_KH - 1]);
      var tenSP = String(row[COL_DH.TEN_SP - 1]);
      var sl = Number(row[COL_DH.SO_LUONG - 1]) || 0;
      var donGia = Number(row[COL_DH.DON_GIA - 1]) || 0;
      var thanhTien = Number(row[COL_DH.THANH_TIEN - 1]) || 0;
      var nguoiBan = String(row[COL_DH.NGUOI_BAN - 1]);
      var nguoiHoTro = String(row[COL_DH.NGUOI_HO_TRO - 1]);
      var laApple = isApple(String(row[COL_DH.THUONG_HIEU - 1]));
      var branch = String(row[COL_DH.CHI_NHANH - 1] || "");
      var ghiChu = String(row[COL_DH.GHI_CHU - 1] || "");

      // Bán chính
      if (nguoiBan && (targetMaNV === "tat_ca" || nguoiBan === targetMaNV)) {
        var nvInfo = nvMap[nguoiBan] || {
          tenNV: nguoiBan,
          vaiTro: "Bán hàng",
          coQuyenXuatMay: false,
        };
        var hh = 0;
        if (nvInfo.coQuyenXuatMay) {
          hh = laApple ? hhBanApple : hhBanKhac;
        }
        txList.push({
          time: ngayBan,
          maGD: maDH,
          maNV: nguoiBan,
          tenNV: nvInfo.tenNV,
          vaiTro: "Bán chính",
          loaiGD: "Đơn hàng (Bán máy)",
          detail: tenSP + " (SL: " + sl + " x " + formatCurrency(donGia) + "đ)",
          revenue: thanhTien,
          commission: hh * sl,
          serviceFee: 0,
          branch: branch,
          customer: tenKH,
          ghiChu: ghiChu,
        });
      }

      // Hỗ trợ
      if (
        nguoiHoTro &&
        (targetMaNV === "tat_ca" || nguoiHoTro === targetMaNV)
      ) {
        var nvInfo = nvMap[nguoiHoTro] || {
          tenNV: nguoiHoTro,
          vaiTro: "Bán hàng",
          coQuyenXuatMay: false,
        };
        var hh = laApple ? hhHTApple : hhHTKhac;
        txList.push({
          time: ngayBan,
          maGD: maDH,
          maNV: nguoiHoTro,
          tenNV: nvInfo.tenNV,
          vaiTro: "Hỗ trợ",
          loaiGD: "Đơn hàng (Hỗ trợ)",
          detail: "Hỗ trợ bán: " + tenSP,
          revenue: 0,
          commission: hh * sl,
          serviceFee: 0,
          branch: branch,
          customer: tenKH,
          ghiChu: ghiChu,
        });
      }
    }
  });

  // B. Dịch vụ
  services.forEach(function (row) {
    var ngayGD = row[COL_DV.NGAY_GD - 1];
    if (!(ngayGD instanceof Date)) return;
    var status = String(row[COL_DV.TRANG_THAI - 1]);
    if (ngayGD >= startDate && ngayGD <= endDate && status !== "Huỷ") {
      var maDV = String(row[COL_DV.MA_DV - 1]);
      var loaiDV = String(row[COL_DV.LOAI_DV - 1]);
      var tenKH = String(row[COL_DV.TEN_KH - 1]);
      var soTien = Number(row[COL_DV.SO_TIEN_GD - 1]) || 0;
      var phi = Number(row[COL_DV.PHI_DV - 1]) || 0;
      var nguoiThucHien = String(row[COL_DV.NGUOI_THUC_HIEN - 1]);
      var branch = String(row[COL_DV.CHI_NHANH - 1] || "");
      var ghiChu = String(row[COL_DV.GHI_CHU - 1] || "");

      if (
        nguoiThucHien &&
        (targetMaNV === "tat_ca" || nguoiThucHien === targetMaNV)
      ) {
        var nvInfo = nvMap[nguoiThucHien] || {
          tenNV: nguoiThucHien,
          vaiTro: "Dịch vụ",
          coQuyenXuatMay: false,
        };
        txList.push({
          time: ngayGD,
          maGD: maDV,
          maNV: nguoiThucHien,
          tenNV: nvInfo.tenNV,
          vaiTro: "Thực hiện",
          loaiGD: "Dịch vụ: " + loaiDV,
          detail: loaiDV + " (Số tiền: " + formatCurrency(soTien) + "đ)",
          revenue: soTien,
          commission: 0,
          serviceFee: phi,
          branch: branch,
          customer: tenKH,
          ghiChu: ghiChu,
        });
      }
    }
  });

  // Sắp xếp tăng dần theo thời gian
  txList.sort(function (a, b) {
    return a.time.getTime() - b.time.getTime();
  });

  // Lọc theo Loại giao dịch
  if (gdFilterVal && gdFilterVal !== "Tất cả") {
    txList = txList.filter(function (tx) {
      return tx.loaiGD === gdFilterVal;
    });
  }

  if (txList.length > 0) {
    var txRows = [];
    var sumRev = 0,
      sumComm = 0,
      sumFee = 0;

    txList.forEach(function (tx) {
      txRows.push([
        Utilities.formatDate(
          tx.time,
          "Asia/Ho_Chi_Minh",
          "dd/MM/yyyy HH:mm:ss",
        ),
        tx.maGD,
        tx.tenNV + " (" + tx.maNV + ")",
        tx.vaiTro,
        tx.loaiGD,
        tx.detail,
        tx.revenue,
        tx.commission,
        tx.serviceFee,
        tx.branch,
        tx.customer,
        tx.ghiChu,
      ]);
      sumRev += tx.revenue;
      sumComm += tx.commission;
      sumFee += tx.serviceFee;
    });

    sheet.getRange(txStartRow + 2, 1, txRows.length, 12).setValues(txRows);
    sheet
      .getRange(txStartRow + 2, 7, txRows.length, 3)
      .setNumberFormat("#,##0");

    // Dòng tổng cộng giao dịch
    var txTotalRowIdx = txStartRow + 2 + txRows.length;
    sheet
      .getRange(txTotalRowIdx, 1)
      .setValue("Tổng cộng")
      .setFontWeight("bold");
    sheet
      .getRange(txTotalRowIdx, 7, 1, 3)
      .setValues([[sumRev, sumComm, sumFee]]);
    sheet
      .getRange(txTotalRowIdx, 7, 1, 3)
      .setNumberFormat("#,##0")
      .setFontWeight("bold");
    sheet
      .getRange(txTotalRowIdx, 1, 1, 12)
      .setBackground("#f1f3f4")
      .setFontWeight("bold");

    sheet
      .getRange(txStartRow, 1, txRows.length + 3, 12)
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#dadce0",
        SpreadsheetApp.BorderStyle.SOLID,
      );
  } else {
    var emptyRange = sheet.getRange(txStartRow + 2, 1, 1, 12);
    emptyRange.merge();
    emptyRange
      .setValue("Không có giao dịch chi tiết nào phù hợp với bộ lọc.")
      .setFontStyle("italic")
      .setHorizontalAlignment("center");

    sheet
      .getRange(txStartRow, 1, 3, 12)
      .setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#dadce0",
        SpreadsheetApp.BorderStyle.SOLID,
      );
  }

  // Format font Times New Roman size 12
  var maxRows = sheet.getMaxRows();
  var maxCols = sheet.getMaxColumns();
  sheet
    .getRange(1, 1, maxRows, maxCols)
    .setFontFamily("Times New Roman")
    .setFontSize(12);

  // Auto resize columns
  for (var col = 1; col <= 12; col++) {
    sheet.autoResizeColumn(col);
  }
}

/**
 * Tính toán doanh số live cho từng nhân viên trong khoảng ngày
 * @private
 */
function _tinhDoanhSoTuNgayDenNgay(startDate, endDate) {
  var allNV = getAllData(SHEET_NAMES.NHAN_VIEN);
  var orders = getAllData(SHEET_NAMES.DON_HANG);
  var services = getAllData(SHEET_NAMES.DICH_VU);

  // Lọc đơn hàng thỏa mãn
  var dtDonHangs = orders.filter(function (row) {
    var ngayBan = row[COL_DH.NGAY_BAN - 1];
    if (!(ngayBan instanceof Date)) return false;
    var status = String(row[COL_DH.TRANG_THAI - 1]);
    return (
      ngayBan >= startDate &&
      ngayBan <= endDate &&
      String(row[COL_DH.NGUON_SP - 1]) === "Điện thoại" &&
      status !== "Huỷ" &&
      status !== "Đổi trả"
    );
  });

  // Lấy cấu hình hoa hồng
  var hhBanApple = getConfigNumber("HH Bán máy - Apple");
  var hhHTApple = getConfigNumber("HH Hỗ trợ - Apple");
  var hhBanKhac = getConfigNumber("HH Bán máy - Khác");
  var hhHTKhac = getConfigNumber("HH Hỗ trợ - Khác");

  var nvDoanhSo = {};

  allNV.forEach(function (nv) {
    if (String(nv[7]) === "Nghỉ việc") return;

    var maNV = String(nv[0]);
    nvDoanhSo[maNV] = {
      maNV: maNV,
      tenNV: nv[1],
      vaiTro: String(nv[4]),
      coQuyenXuatMay: String(nv[5]) === "✓",
      soMayBan_Apple: 0,
      soMayBan_Khac: 0,
      soMayHoTro_Apple: 0,
      soMayHoTro_Khac: 0,
      hhBan: 0,
      hhHoTro: 0,
      tongHoaHong: 0,
      doanhThuDV: 0,
      thuong: 0,
      tongThuNhap: 0,
    };
  });

  // Quét đơn hàng phân bổ
  dtDonHangs.forEach(function (row) {
    var nguoiBan = String(row[COL_DH.NGUOI_BAN - 1]);
    var nguoiHoTro = String(row[COL_DH.NGUOI_HO_TRO - 1]);
    var laApple = isApple(String(row[COL_DH.THUONG_HIEU - 1]));

    if (nguoiBan && nvDoanhSo[nguoiBan]) {
      if (laApple) nvDoanhSo[nguoiBan].soMayBan_Apple++;
      else nvDoanhSo[nguoiBan].soMayBan_Khac++;
    }

    if (nguoiHoTro && nvDoanhSo[nguoiHoTro]) {
      if (laApple) nvDoanhSo[nguoiHoTro].soMayHoTro_Apple++;
      else nvDoanhSo[nguoiHoTro].soMayHoTro_Khac++;
    }
  });

  // Quét dịch vụ phân bổ
  services.forEach(function (row) {
    var ngayGD = row[COL_DV.NGAY_GD - 1];
    if (!(ngayGD instanceof Date)) return;
    var status = String(row[COL_DV.TRANG_THAI - 1]);
    if (ngayGD >= startDate && ngayGD <= endDate && status !== "Huỷ") {
      var maNV = String(row[COL_DV.NGUOI_THUC_HIEN - 1]);
      var phi = Number(row[COL_DV.PHI_DV - 1]) || 0;
      if (maNV && nvDoanhSo[maNV]) {
        nvDoanhSo[maNV].doanhThuDV += phi;
      }
    }
  });

  var listNV = [];
  Object.keys(nvDoanhSo).forEach(function (maNV) {
    var nv = nvDoanhSo[maNV];
    if (nv.coQuyenXuatMay) {
      nv.hhBan = nv.soMayBan_Apple * hhBanApple + nv.soMayBan_Khac * hhBanKhac;
    }
    nv.hhHoTro =
      nv.soMayHoTro_Apple * hhHTApple + nv.soMayHoTro_Khac * hhHTKhac;
    nv.tongHoaHong = nv.hhBan + nv.hhHoTro;
    nv.tongThuNhap = nv.tongHoaHong + nv.doanhThuDV;

    // Chỉ lấy nhân viên có hoạt động
    if (
      nv.soMayBan_Apple +
        nv.soMayBan_Khac +
        nv.soMayHoTro_Apple +
        nv.soMayHoTro_Khac >
        0 ||
      nv.doanhThuDV > 0
    ) {
      listNV.push(nv);
    }
  });

  return listNV;
}

/**
 * Lấy báo cáo doanh số động theo cấu hình bộ lọc từ client (Sidebar hỗ trợ nếu cần)
 *
 * @param {Object} params - { loaiThoiGian, thang, nam, tuNgay, denNgay, maNV }
 * @return {Object} Báo cáo doanh số
 */
function getBaoCaoDoanhSoDong(params) {
  initializeColumnEnums();
  var loaiTG = params.loaiThoiGian || "thang";
  var targetMaNV = params.maNV || "tat_ca";
  var result = {
    thangNam: "",
    nhanVien: [],
    tongSoMayBan: 0,
    tongHoaHong: 0,
    tongDoanhThuDV: 0,
    tongThuNhap: 0,
  };

  var listNV = [];

  if (loaiTG === "thang") {
    var thangNam = ("0" + params.thang).slice(-2) + "/" + params.nam;
    result.thangNam = thangNam;

    // Xem báo cáo từ dữ liệu đã chốt
    var dataDS = getAllData(SHEET_NAMES.DOANH_SO);
    var records = [];
    dataDS.forEach(function (row) {
      if (formatMonthYear(row[0]) === thangNam) {
        records.push({
          maNV: String(row[1]),
          tenNV: String(row[2]),
          vaiTro: String(row[3]),
          coQuyenXuatMay: String(row[4]) === "✓",
          soMayBan_Apple: Number(row[5]) || 0,
          soMayBan_Khac: Number(row[6]) || 0,
          soMayHoTro_Apple: Number(row[7]) || 0,
          soMayHoTro_Khac: Number(row[8]) || 0,
          hhBan: Number(row[9]) || 0,
          hhHoTro: Number(row[10]) || 0,
          tongHoaHong: Number(row[11]) || 0,
          doanhThuDV: Number(row[12]) || 0,
          thuong: Number(row[13]) || 0,
          tongThuNhap: Number(row[14]) || 0,
        });
      }
    });

    if (records.length > 0) {
      listNV = records;
    } else {
      var ngayDau = new Date(params.nam, params.thang - 1, 1, 0, 0, 0);
      var ngayCuoi = new Date(params.nam, params.thang, 0, 23, 59, 59);
      listNV = _tinhDoanhSoTuNgayDenNgay(ngayDau, ngayCuoi);
    }
  } else {
    var tuNgay = new Date(params.tuNgay + "T00:00:00");
    var denNgay = new Date(params.denNgay + "T23:59:59");
    result.thangNam = formatDate(tuNgay) + " - " + formatDate(denNgay);
    listNV = _tinhDoanhSoTuNgayDenNgay(tuNgay, denNgay);
  }

  if (targetMaNV !== "tat_ca") {
    listNV = listNV.filter(function (nv) {
      return nv.maNV === targetMaNV;
    });
  }

  listNV.forEach(function (nv) {
    result.nhanVien.push(nv);
    result.tongSoMayBan += nv.soMayBan_Apple + nv.soMayBan_Khac;
    result.tongHoaHong += nv.tongHoaHong;
    result.tongDoanhThuDV += nv.doanhThuDV;
    result.tongThuNhap += nv.tongThuNhap;
  });

  return result;
}
