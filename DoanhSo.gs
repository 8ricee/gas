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
    if (String(data[i][0]) === thangNam) {
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
    if (String(row[0]) === thangNam) {
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
      msg += "\n👤 " + nv.tenNV + " (" + nv.maNV + ")";
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
