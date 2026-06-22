/**
 * ============================================================
 * VanTran Mobile — BaoCaoNgay.gs
 * Nghiệp vụ tổng hợp báo cáo tài chính và giao dịch theo ngày
 * ============================================================
 */

/**
 * Tạo/Cập nhật báo cáo ngày dựa trên ngày được chọn ở ô B3
 */
function updateDailyReportFromSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_NGAY);
  if (!sheet) return;

  var dateVal = sheet.getRange(3, 2).getValue(); // Cell B3
  var date = _parseDate(dateVal);
  if (!date) {
    sheet
      .getRange(4, 2)
      .setValue("Ngày không hợp lệ! Vui lòng nhập định dạng dd/MM/yyyy");
    return;
  }

  sheet.getRange(4, 2).setValue("Đang tải báo cáo...");
  SpreadsheetApp.flush();

  try {
    generateDailyReport(date);
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
 * Tạo báo cáo ngày chi tiết
 *
 * @param {Date} targetDate - Ngày cần báo cáo
 */
function generateDailyReport(targetDate) {
  initializeColumnEnums();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var reportSheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_NGAY);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(SHEET_NAMES.BAO_CAO_NGAY);
  }

  // Đọc danh sách chi nhánh động
  var branches = getBranchesList();
  var fallbackBranch = "Khác/Hệ thống";

  // Khởi tạo accumulator tài chính theo từng chi nhánh
  var branchMetrics = {};
  branches.forEach(function (b) {
    branchMetrics[b] = { doanhSo: 0, loiNhuan: 0, thuTM: 0, chiTM: 0, thuCK: 0, chiCK: 0 };
  });
  branchMetrics[fallbackBranch] = { doanhSo: 0, loiNhuan: 0, thuTM: 0, chiTM: 0, thuCK: 0, chiCK: 0 };

  // 1. Thu thập giao dịch của ngày targetDate
  var transactions = [];

  // --- A. ĐƠN HÀNG (DON_HANG) ---
  var orders = getAllData(SHEET_NAMES.DON_HANG);
  orders.forEach(function (row) {
    var ngayBan = row[COL_DH.NGAY_BAN - 1];
    if (ngayBan instanceof Date && _isSameDay(ngayBan, targetDate)) {
      var status = String(row[COL_DH.TRANG_THAI - 1]);
      if (status === "Huỷ" || status === "Đổi trả") return; // Skip canceled/returned orders

      var maDH = String(row[COL_DH.MA_DH - 1]);
      var tenKH = String(row[COL_DH.TEN_KH - 1]);
      var maSP = String(row[COL_DH.MA_SP - 1]);
      var tenSP = String(row[COL_DH.TEN_SP - 1]);
      var nguonSP = String(row[COL_DH.NGUON_SP - 1]);
      var sl = Number(row[COL_DH.SO_LUONG - 1]) || 0;
      var donGia = Number(row[COL_DH.DON_GIA - 1]) || 0;
      var thanhTien = parseAmountVal(row[COL_DH.THANH_TIEN - 1]);
      var hinhThucBan = String(row[COL_DH.HINH_THUC_BAN - 1]);
      var hinhThucTT = String(row[COL_DH.HINH_THUC_TT - 1]);
      var branch = String(row[COL_DH.CHI_NHANH - 1] || "").trim();
      var maQua = String(row[COL_DH.MA_QUA_TANG - 1] || "");
      var coNhanQua = String(row[COL_DH.CO_NHAN_QUA - 1] || "✗");
      var giamGia = Number(row[COL_DH.TIEN_GIAM_GIA - 1] || 0);

      // Tra cứu giá nhập sản phẩm chính
      var giaNhapSP = 0;
      if (nguonSP === "Điện thoại") {
        giaNhapSP = Number(lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.IMEI, maSP, COL_DT.GIA_NHAP)) || 
                    Number(lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.GIA_NHAP)) || 0;
      } else {
        var pkRow = findPhuKienRow(maSP, branch);
        if (pkRow !== -1) {
          giaNhapSP = Number(ss.getSheetByName(SHEET_NAMES.PHU_KIEN).getRange(pkRow, COL_PK.GIA_NHAP).getValue()) || 0;
        }
      }
      var costSP = giaNhapSP * sl;

      // Tra cứu giá nhập quà tặng (nếu nhận quà)
      var costQua = 0;
      if (coNhanQua === "✓" && maQua) {
        var maQuaList = String(maQua).split(",");
        for (var i = 0; i < maQuaList.length; i++) {
          var code = maQuaList[i].trim();
          if (!code) continue;
          var quaRow = findPhuKienRow(code, branch);
          if (quaRow !== -1) {
            costQua += Number(ss.getSheetByName(SHEET_NAMES.PHU_KIEN).getRange(quaRow, COL_PK.GIA_NHAP).getValue()) || 0;
          }
        }
      }

      var loiNhuan = thanhTien - costSP - costQua;

      // Tính dòng tiền
      var thuTM = 0;
      var thuCK = 0;

      var cellTM = row[COL_DH.TIEN_MAT - 1];
      var cellCK = row[COL_DH.CHUYEN_KHOAN - 1];

      // Check if we have new columns populated with non-zero values (new flow)
      var hasNewColumns = (cellTM !== undefined && cellTM !== "" && Number(cellTM) !== 0) || 
                          (cellCK !== undefined && cellCK !== "" && Number(cellCK) !== 0);

      if (hasNewColumns) {
        thuTM = Number(cellTM) || 0;
        thuCK = Number(cellCK) || 0;

        // Với trả góp CTTC, công ty tài chính giải ngân phần còn lại qua CK
        if (hinhThucBan === "Trả góp") {
          var traTruocRaw = lookupValue(SHEET_NAMES.TRA_GOP, COL_TG.MA_DH, maDH, COL_TG.TRA_TRUOC);
          var traTruoc = Number(traTruocRaw) || 0;
          var loaiTG = String(lookupValue(SHEET_NAMES.TRA_GOP, COL_TG.MA_DH, maDH, COL_TG.LOAI_TRA_GOP));
          var conLai = thanhTien - traTruoc;
          if (loaiTG === "Công ty tài chính") {
            thuCK += conLai;
          }
        }
      } else {
        // Fallback cho dữ liệu cũ (chưa chia cột)
        var amountCollected = thanhTien;
        var rawAmount = row[COL_DH.THANH_TIEN - 1];

        if (hinhThucBan === "Trả góp") {
          var traTruocRaw = lookupValue(SHEET_NAMES.TRA_GOP, COL_TG.MA_DH, maDH, COL_TG.TRA_TRUOC);
          var traTruoc = parseAmountVal(traTruocRaw);
          var loaiTG = String(lookupValue(SHEET_NAMES.TRA_GOP, COL_TG.MA_DH, maDH, COL_TG.LOAI_TRA_GOP));
          var conLai = thanhTien - traTruoc;

          if (loaiTG === "Công ty tài chính") {
            // Công ty tài chính giải ngân qua chuyển khoản
            thuCK += conLai;
          }
          amountCollected = traTruoc;
          rawAmount = traTruocRaw;
        }

        var hybrid = parseHybridAmount(hinhThucTT === "Hỗn hợp" ? rawAmount : "");
        if (hybrid) {
          thuTM += hybrid.tm;
          thuCK += hybrid.ck;
        } else {
          var payment = parseMixedPayment(hinhThucTT, amountCollected);
          thuTM += payment.tm;
          thuCK += payment.ck;
        }
      }

      var brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = { doanhSo: 0, loiNhuan: 0, thuTM: 0, chiTM: 0, thuCK: 0, chiCK: 0 };
      }
      branchMetrics[brKey].doanhSo += thanhTien;
      branchMetrics[brKey].loiNhuan += loiNhuan;
      branchMetrics[brKey].thuTM += thuTM;
      branchMetrics[brKey].thuCK += thuCK;

      transactions.push({
        maGD: maDH,
        time: ngayBan,
        loaiGD: "Đơn hàng",
        khachHang: tenKH,
        branch: branch,
        thuTM: thuTM,
        chiTM: 0,
        thuCK: thuCK,
        chiCK: 0,
        loiNhuan: loiNhuan,
        detail: tenSP + " (SL: " + sl + ", " + hinhThucBan + " - " + hinhThucTT + ")",
      });
    }
  });

  // --- B. DỊCH VỤ (DICH_VU) ---
  var services = getAllData(SHEET_NAMES.DICH_VU);
  services.forEach(function (row) {
    var ngayGD = row[COL_DV.NGAY_GD - 1];
    if (ngayGD instanceof Date && _isSameDay(ngayGD, targetDate)) {
      var status = String(row[COL_DV.TRANG_THAI - 1]);
      if (status === "Huỷ") return;

      var maDV = String(row[COL_DV.MA_DV - 1]);
      var loaiDV = String(row[COL_DV.LOAI_DV - 1]);
      var tenKH = String(row[COL_DV.TEN_KH - 1]);
      var rawSoTien = row[COL_DV.SO_TIEN_GD - 1];
      var soTien = parseAmountVal(rawSoTien);
      var phi = Number(row[COL_DV.PHI_DV - 1]) || 0;
      var hinhThucTT = String(row[COL_DV.HINH_THUC_TT - 1]);
      var branch = String(row[COL_DV.CHI_NHANH - 1] || "").trim();

      var thuTM = 0;
      var chiTM = 0;
      var thuCK = 0;
      var chiCK = 0;
      var loiNhuan = phi;

      var cellTM = row[COL_DV.TIEN_MAT - 1];
      var cellCK = row[COL_DV.CHUYEN_KHOAN - 1];

      var hasNewColumns = (cellTM !== undefined && cellTM !== "" && Number(cellTM) !== 0) || 
                          (cellCK !== undefined && cellCK !== "" && Number(cellCK) !== 0);

      if (hasNewColumns) {
        thuTM = Number(cellTM) || 0;
        thuCK = Number(cellCK) || 0;

        if (loaiDV === "Chuyển khoản hộ" || loaiDV === "Nạp thẻ điện thoại") {
          chiCK = soTien;
        } else if (loaiDV === "Rút tiền mặt") {
          chiTM = soTien;
        }
      } else {
        // Fallback cho dữ liệu cũ (chưa chia cột)
        var hybrid = parseHybridAmount(hinhThucTT === "Hỗn hợp" ? rawSoTien : "");
        if (loaiDV === "Chuyển khoản hộ") {
          if (hybrid) {
            thuTM = hybrid.tm;
            thuCK = hybrid.ck;
            chiCK = soTien - phi;
          } else {
            var payment = parseMixedPayment(hinhThucTT, soTien + phi);
            thuTM = payment.tm;
            thuCK = payment.ck;
            chiCK = soTien;
          }
        } else if (loaiDV === "Rút tiền mặt") {
          thuCK = soTien + phi;
          chiTM = soTien;
        } else if (loaiDV === "Nạp thẻ điện thoại") {
          if (hybrid) {
            thuTM = hybrid.tm;
            thuCK = hybrid.ck;
            chiCK = soTien - phi;
          } else {
            var payment = parseMixedPayment(hinhThucTT, soTien + phi);
            thuTM = payment.tm;
            thuCK = payment.ck;
            chiCK = soTien;
          }
        }
      }

      var brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = { doanhSo: 0, loiNhuan: 0, thuTM: 0, chiTM: 0, thuCK: 0, chiCK: 0 };
      }
      branchMetrics[brKey].loiNhuan += loiNhuan;
      branchMetrics[brKey].thuTM += thuTM;
      branchMetrics[brKey].chiTM += chiTM;
      branchMetrics[brKey].thuCK += thuCK;
      branchMetrics[brKey].chiCK += chiCK;

      transactions.push({
        maGD: maDV,
        time: ngayGD,
        loaiGD: "Dịch vụ",
        khachHang: tenKH,
        branch: branch,
        thuTM: thuTM,
        chiTM: chiTM,
        thuCK: thuCK,
        chiCK: chiCK,
        loiNhuan: loiNhuan,
        detail: loaiDV + " (" + formatCurrency(soTien) + "đ, Phí: " + formatCurrency(phi) + "đ)",
      });
    }
  });

  // --- C. TRẢ GÓP (LICH_SU_TRA_GOP - repayments) ---
  var repayments = getAllData(SHEET_NAMES.LICH_SU_TRA_GOP);
  repayments.forEach(function (row) {
    var ngayTra = row[COL_LSTG.NGAY_THUC_TRA - 1];
    var status = String(row[COL_LSTG.TRANG_THAI - 1]);
    if (ngayTra instanceof Date && _isSameDay(ngayTra, targetDate) && status === "Đã trả") {
      var maLS = String(row[COL_LSTG.MA_LS - 1]);
      var maTG = String(row[COL_LSTG.MA_TG - 1]);
      var kySo = String(row[COL_LSTG.KY_SO - 1]);
      var rawSoTienDaTra = row[COL_LSTG.SO_TIEN_DA_TRA - 1];
      var soTienDaTra = parseAmountVal(rawSoTienDaTra);
      var hinhThucTT = String(row[COL_LSTG.HINH_THUC_TT - 1]);

      var maKH = String(lookupValue(SHEET_NAMES.TRA_GOP, COL_TG.MA_TG, maTG, COL_TG.MA_KH));
      var tenKH = String(lookupValue(SHEET_NAMES.KHACH_HANG, 1, maKH, 2)) || "Khách trả góp";
      var branch = String(lookupValue(SHEET_NAMES.TRA_GOP, COL_TG.MA_TG, maTG, COL_TG.CHI_NHANH) || "").trim();

      var thuTM = 0;
      var thuCK = 0;

      var cellTM = row[COL_LSTG.TIEN_MAT - 1];
      var cellCK = row[COL_LSTG.CHUYEN_KHOAN - 1];

      var hasNewColumns = (cellTM !== undefined && cellTM !== "" && Number(cellTM) !== 0) || 
                          (cellCK !== undefined && cellCK !== "" && Number(cellCK) !== 0);

      if (hasNewColumns) {
        thuTM = Number(cellTM) || 0;
        thuCK = Number(cellCK) || 0;
      } else {
        // Fallback cho dữ liệu cũ (chưa chia cột)
        var hybrid = parseHybridAmount(hinhThucTT === "Hỗn hợp" ? rawSoTienDaTra : "");
        if (hybrid) {
          thuTM = hybrid.tm;
          thuCK = hybrid.ck;
        } else {
          var payment = parseMixedPayment(hinhThucTT, soTienDaTra);
          thuTM = payment.tm;
          thuCK = payment.ck;
        }
      }

      var brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = { doanhSo: 0, loiNhuan: 0, thuTM: 0, chiTM: 0, thuCK: 0, chiCK: 0 };
      }
      branchMetrics[brKey].thuTM += thuTM;
      branchMetrics[brKey].thuCK += thuCK;

      transactions.push({
        maGD: maLS,
        time: ngayTra,
        loaiGD: "Thu trả góp",
        khachHang: tenKH,
        branch: branch,
        thuTM: thuTM,
        chiTM: 0,
        thuCK: thuCK,
        chiCK: 0,
        loiNhuan: 0,
        detail: "Thanh toán HĐ " + maTG + " - Kỳ " + kySo,
      });
    }
  });

  // --- D. NHẬP KHO (NHAP_KHO) ---
  var imports = getAllData(SHEET_NAMES.NHAP_KHO);
  imports.forEach(function (row) {
    var ngayNhap = row[1];
    if (ngayNhap instanceof Date && _isSameDay(ngayNhap, targetDate)) {
      var maNK = String(row[0]);
      var nguonNhap = String(row[2]);
      var tenSP = String(row[4]);
      var sl = Number(row[5]) || 0;
      var thanhTien = Number(row[7]) || 0;
      var nhaCC = String(row[8] || "");
      var branch = String(row[10] || "").trim();

      var chiCK = thanhTien;

      var brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = { doanhSo: 0, loiNhuan: 0, thuTM: 0, chiTM: 0, thuCK: 0, chiCK: 0 };
      }
      branchMetrics[brKey].chiCK += chiCK;

      transactions.push({
        maGD: maNK,
        time: ngayNhap,
        loaiGD: "Nhập kho",
        khachHang: nhaCC || "Nhà cung cấp",
        branch: branch,
        thuTM: 0,
        chiTM: 0,
        thuCK: 0,
        chiCK: chiCK,
        loiNhuan: 0,
        detail: "Nhập " + nguonNhap + ": " + tenSP + " (SL: " + sl + ")",
      });
    }
  });

  // --- E. ĐỔI TRẢ (DOI_TRA) ---
  var returns = getAllData(SHEET_NAMES.DOI_TRA);
  returns.forEach(function (row) {
    var ngayDT = row[COL_DT_TRA.NGAY_DT - 1];
    if (ngayDT instanceof Date && _isSameDay(ngayDT, targetDate)) {
      var status = String(row[COL_DT_TRA.TRANG_THAI - 1]);
      if (status === "Huỷ") return;

      var maDT = String(row[COL_DT_TRA.MA_DT - 1]);
      var maDH = String(row[COL_DT_TRA.MA_DH - 1]);
      var tenKH = String(row[COL_DT_TRA.TEN_KH - 1]);
      var loaiGD = String(row[COL_DT_TRA.LOAI_GD - 1]);
      var tenSPTra = String(row[COL_DT_TRA.TEN_SP_TRA - 1]);
      var tenSPNhan = String(row[COL_DT_TRA.TEN_SP_NHAN - 1]);
      var rawHoanTien = row[COL_DT_TRA.TIEN_HOAN_TRA - 1];
      var hoanTien = parseAmountVal(rawHoanTien);
      var phi = Number(row[COL_DT_TRA.PHI_DOI_TRA - 1]) || 0;
      var hinhThucTT = String(row[COL_DT_TRA.HINH_THUC_TT - 1]);
      var branch = String(row[COL_DT_TRA.CHI_NHANH - 1] || "").trim();

      var thuTM = 0;
      var chiTM = 0;
      var thuCK = 0;
      var chiCK = 0;

      var originalOrder = getDonHangDetails(maDH);
      var costOriginal = 0;
      var priceOriginal = 0;
      if (originalOrder) {
        priceOriginal = parseAmountVal(originalOrder.thanhTien);
        var costSP = 0;
        if (originalOrder.nguonSP === "Điện thoại") {
          costSP = Number(lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.IMEI, originalOrder.maSP, COL_DT.GIA_NHAP)) || 
                   Number(lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, originalOrder.maSP, COL_DT.GIA_NHAP)) || 0;
        } else {
          var pkRow = findPhuKienRow(originalOrder.maSP, originalOrder.chiNhanh);
          if (pkRow !== -1) {
            costSP = Number(ss.getSheetByName(SHEET_NAMES.PHU_KIEN).getRange(pkRow, COL_PK.GIA_NHAP).getValue()) || 0;
          }
        }
        costOriginal = costSP * originalOrder.soLuong;

        if (originalOrder.coNhanQua === "✓" && originalOrder.maQuaTang) {
          var maQuaOriginalList = String(originalOrder.maQuaTang).split(",");
          for (var i = 0; i < maQuaOriginalList.length; i++) {
            var code = maQuaOriginalList[i].trim();
            if (!code) continue;
            var quaRow = findPhuKienRow(code, originalOrder.chiNhanh);
            if (quaRow !== -1) {
              costOriginal += Number(ss.getSheetByName(SHEET_NAMES.PHU_KIEN).getRange(quaRow, COL_PK.GIA_NHAP).getValue()) || 0;
            }
          }
        }
      }

      var lossDoanhSo = 0;
      var lossLoiNhuan = 0;

      var cellTM = row[COL_DT_TRA.TIEN_MAT - 1];
      var cellCK = row[COL_DT_TRA.CHUYEN_KHOAN - 1];

      var hasNewColumns = (cellTM !== undefined && cellTM !== "" && Number(cellTM) !== 0) || 
                          (cellCK !== undefined && cellCK !== "" && Number(cellCK) !== 0);

      if (hasNewColumns) {
        var valTM = Number(cellTM) || 0;
        var valCK = Number(cellCK) || 0;

        if (hoanTien >= 0) {
          chiTM = valTM;
          chiCK = valCK;
        } else {
          thuTM = valTM;
          thuCK = valCK;
        }
      } else {
        // Fallback cho dữ liệu cũ (chưa chia cột)
        var hybrid = parseHybridAmount(hinhThucTT === "Hỗn hợp" ? rawHoanTien : "");
        if (loaiGD === "Trả máy") {
          if (hybrid) {
            chiTM = hybrid.tm;
            chiCK = hybrid.ck;
          } else {
            var payment = parseMixedPayment(hinhThucTT, hoanTien);
            chiTM = payment.tm;
            chiCK = payment.ck;
          }
        } else {
          if (hoanTien >= 0) {
            if (hybrid) {
              chiTM = hybrid.tm;
              chiCK = hybrid.ck;
            } else {
              var payment = parseMixedPayment(hinhThucTT, hoanTien);
              chiTM = payment.tm;
              chiCK = payment.ck;
            }
          } else {
            if (hybrid) {
              thuTM = hybrid.tm;
              thuCK = hybrid.ck;
            } else {
              var payment = parseMixedPayment(hinhThucTT, Math.abs(hoanTien));
              thuTM = payment.tm;
              thuCK = payment.ck;
            }
          }
        }
      }

      lossDoanhSo = -priceOriginal;
      lossLoiNhuan = -(priceOriginal - costOriginal);
      lossLoiNhuan += phi;

      var brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = { doanhSo: 0, loiNhuan: 0, thuTM: 0, chiTM: 0, thuCK: 0, chiCK: 0 };
      }
      branchMetrics[brKey].doanhSo += lossDoanhSo;
      branchMetrics[brKey].loiNhuan += lossLoiNhuan;
      branchMetrics[brKey].thuTM += thuTM;
      branchMetrics[brKey].chiTM += chiTM;
      branchMetrics[brKey].thuCK += thuCK;
      branchMetrics[brKey].chiCK += chiCK;

      transactions.push({
        maGD: maDT,
        time: ngayDT,
        loaiGD: "Đổi trả",
        khachHang: tenKH,
        branch: branch,
        thuTM: thuTM,
        chiTM: chiTM,
        thuCK: thuCK,
        chiCK: chiCK,
        loiNhuan: lossLoiNhuan,
        detail: loaiGD + ": " + tenSPTra + (loaiGD === "Đổi máy" ? " ➔ " + tenSPNhan : "") + " (Phí: " + formatCurrency(phi) + "đ)",
      });
    }
  });

  // --- F. THU MUA (THU_MUA) ---
  var buybacks = getAllData(SHEET_NAMES.THU_MUA);
  buybacks.forEach(function (row) {
    var ngayTM = row[COL_TM.NGAY_TM - 1];
    if (ngayTM instanceof Date && _isSameDay(ngayTM, targetDate)) {
      var maTM = String(row[COL_TM.MA_TM - 1]);
      var tenKH = String(row[COL_TM.TEN_KH - 1]);
      var tenSPThu = String(row[COL_TM.TEN_SP_THU - 1]);
      var rawTongTienTra = row[COL_TM.TONG_TIEN_TRA - 1];
      var tongTienTra = parseAmountVal(rawTongTienTra);
      var hinhThucTT = String(row[COL_TM.HINH_THUC_TT - 1]);
      var branch = String(row[COL_TM.CHI_NHANH - 1] || "").trim();

      var chiTM = 0;
      var chiCK = 0;

      var cellTM = row[COL_TM.TIEN_MAT - 1];
      var cellCK = row[COL_TM.CHUYEN_KHOAN - 1];

      var hasNewColumns = (cellTM !== undefined && cellTM !== "" && Number(cellTM) !== 0) || 
                          (cellCK !== undefined && cellCK !== "" && Number(cellCK) !== 0);

      if (hasNewColumns) {
        chiTM = Number(cellTM) || 0;
        chiCK = Number(cellCK) || 0;
      } else {
        // Fallback cho dữ liệu cũ (chưa chia cột)
        var hybrid = parseHybridAmount(hinhThucTT === "Hỗn hợp" ? rawTongTienTra : "");
        if (hybrid) {
          chiTM = hybrid.tm;
          chiCK = hybrid.ck;
        } else {
          var payment = parseMixedPayment(hinhThucTT, tongTienTra);
          chiTM = payment.tm;
          chiCK = payment.ck;
        }
      }

      var brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = { doanhSo: 0, loiNhuan: 0, thuTM: 0, chiTM: 0, thuCK: 0, chiCK: 0 };
      }
      branchMetrics[brKey].chiTM += chiTM;
      branchMetrics[brKey].chiCK += chiCK;

      transactions.push({
        maGD: maTM,
        time: ngayTM,
        loaiGD: "Thu mua",
        khachHang: tenKH,
        branch: branch,
        thuTM: 0,
        chiTM: chiTM,
        thuCK: 0,
        chiCK: chiCK,
        loiNhuan: 0,
        detail: "Thu mua máy cũ: " + tenSPThu + " (Giá: " + formatCurrency(tongTienTra) + "đ)",
      });
    }
  });

  // Sắp xếp tăng dần theo thời gian
  transactions.sort(function (a, b) {
    return a.time.getTime() - b.time.getTime();
  });

  // --- 2. GHI BÁO CÁO LÊN SHEET ---
  reportSheet.clearContents();
  reportSheet.getRange("A1:M1000").clearFormat();

  // Tạo tiêu đề
  reportSheet.getRange(1, 1)
             .setValue("BÁO CÁO GIAO DỊCH THEO NGÀY")
             .setFontSize(14)
             .setFontWeight("bold");

  reportSheet.getRange(3, 1)
             .setValue("Ngày báo cáo:")
             .setFontWeight("bold");

  var dateCell = reportSheet.getRange(3, 2);
  dateCell.setValue(targetDate)
          .setNumberFormat("dd/MM/yyyy")
          .setBackground("#ffffff")
          .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID)
          .setHorizontalAlignment("left");

  reportSheet.getRange(4, 1)
             .setValue("Trạng thái cập nhật:")
             .setFontWeight("bold");

  // Chuẩn bị danh sách các chi nhánh hiển thị ở bảng tổng hợp
  var activeBranches = branches.slice();
  var fb = branchMetrics[fallbackBranch];
  if (fb.doanhSo !== 0 || fb.loiNhuan !== 0 || fb.thuTM !== 0 || fb.chiTM !== 0 || fb.thuCK !== 0 || fb.chiCK !== 0) {
    activeBranches.push(fallbackBranch);
  }
  var numActiveBranches = activeBranches.length;
  var totalSummaryColIdx = numActiveBranches + 2;

  // Tiêu đề bảng tổng hợp chỉ tiêu tài chính
  var summaryTitleRange = reportSheet.getRange(6, 1, 1, totalSummaryColIdx);
  summaryTitleRange.merge();
  summaryTitleRange.setValue("TỔNG HỢP CHỈ TIÊU TÀI CHÍNH")
                    .setBackground("#1a73e8")
                    .setFontColor("#ffffff")
                    .setFontWeight("bold")
                    .setHorizontalAlignment("center");

  // Tiêu đề cột bảng tổng hợp
  var summaryHeaders = ["Chỉ tiêu"];
  activeBranches.forEach(function (b) {
    summaryHeaders.push(b);
  });
  summaryHeaders.push("Tổng cộng");
  reportSheet.getRange(7, 1, 1, totalSummaryColIdx).setValues([summaryHeaders]);
  reportSheet.getRange(7, 1, 1, totalSummaryColIdx)
             .setFontWeight("bold")
             .setBackground("#e8f0fe")
             .setHorizontalAlignment("center");

  // Ghi dữ liệu bảng tổng hợp
  var metricsList = [
    { name: "1. Doanh số bán lẻ", key: "doanhSo" },
    { name: "2. Lợi nhuận gộp", key: "loiNhuan" },
    { name: "3. Tiền mặt Thu", key: "thuTM" },
    { name: "4. Tiền mặt Chi", key: "chiTM" },
    { name: "5. Chuyển khoản Thu", key: "thuCK" },
    { name: "6. Chuyển khoản Chi", key: "chiCK" },
    { name: "7. Dòng tiền ròng trong ngày", key: "netCash" }
  ];

  var summaryData = [];
  metricsList.forEach(function (m) {
    var row = [m.name];
    var totalVal = 0;
    activeBranches.forEach(function (b) {
      var val = 0;
      if (m.key === "netCash") {
        val = branchMetrics[b].thuTM + branchMetrics[b].thuCK - (branchMetrics[b].chiTM + branchMetrics[b].chiCK);
      } else {
        val = branchMetrics[b][m.key];
      }
      row.push(val);
      totalVal += val;
    });
    row.push(totalVal);
    summaryData.push(row);
  });

  reportSheet.getRange(8, 1, summaryData.length, totalSummaryColIdx).setValues(summaryData);

  // Định dạng số và kiểu chữ bảng tổng hợp
  reportSheet.getRange(8, 2, summaryData.length, totalSummaryColIdx - 1)
             .setNumberFormat("#,##0");
  reportSheet.getRange(14, 1, 1, totalSummaryColIdx)
             .setFontColor("#1a73e8")
             .setFontWeight("bold"); // Dòng tiền ròng nổi bật
  reportSheet.getRange(8, totalSummaryColIdx, summaryData.length, 1)
             .setFontWeight("bold"); // Cột tổng cộng in đậm

  // Vẽ viền cho bảng tổng hợp
  reportSheet.getRange(6, 1, summaryData.length + 2, totalSummaryColIdx)
             .setBorder(true, true, true, true, true, true, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);

  // --- 3. GHI CÁC BẢNG LỊCH SỬ CHI TIẾT THEO CHI NHÁNH ---
  var startRow = 17;

  activeBranches.forEach(function (branchName) {
    var branchTxs = transactions.filter(function (tx) {
      return (tx.branch === branchName) || (branchName === fallbackBranch && !tx.branch);
    });

    // 1. Tiêu đề bảng chi tiết chi nhánh
    var titleRange = reportSheet.getRange(startRow, 1, 1, 10);
    titleRange.merge();
    titleRange.setValue("CHI TIẾT GIAO DỊCH - " + branchName.toUpperCase())
              .setBackground("#1a73e8")
              .setFontColor("#ffffff")
              .setFontWeight("bold")
              .setHorizontalAlignment("center");

    // 2. Headers cột giao dịch (lược bỏ cột Chi nhánh)
    var detailHeaders = [
      "Mã giao dịch",
      "Thời gian",
      "Loại giao dịch",
      "Khách hàng",
      "Thu Tiền mặt",
      "Chi Tiền mặt",
      "Thu Chuyển khoản",
      "Chi Chuyển khoản",
      "Lợi nhuận",
      "Chi tiết"
    ];
    reportSheet.getRange(startRow + 1, 1, 1, 10).setValues([detailHeaders]);
    reportSheet.getRange(startRow + 1, 1, 1, 10)
               .setFontWeight("bold")
               .setBackground("#e8f0fe")
               .setHorizontalAlignment("center");

    // 3. Ghi dữ liệu giao dịch chi tiết
    if (branchTxs.length > 0) {
      var detailRows = [];
      var sumThuTM = 0, sumChiTM = 0, sumThuCK = 0, sumChiCK = 0, sumLoiNhuan = 0;

      branchTxs.forEach(function (tx) {
        detailRows.push([
          tx.maGD,
          Utilities.formatDate(tx.time, "Asia/Ho_Chi_Minh", "HH:mm:ss"),
          tx.loaiGD,
          tx.khachHang,
          tx.thuTM,
          tx.chiTM,
          tx.thuCK,
          tx.chiCK,
          tx.loiNhuan,
          tx.detail
        ]);
        sumThuTM += tx.thuTM;
        sumChiTM += tx.chiTM;
        sumThuCK += tx.thuCK;
        sumChiCK += tx.chiCK;
        sumLoiNhuan += tx.loiNhuan;
      });

      reportSheet.getRange(startRow + 2, 1, detailRows.length, 10).setValues(detailRows);
      
      // Format tiền tệ
      reportSheet.getRange(startRow + 2, 5, detailRows.length, 5).setNumberFormat("#,##0");

      // Dòng tổng cộng của riêng chi nhánh
      var totalRowIdx = startRow + 2 + detailRows.length;
      reportSheet.getRange(totalRowIdx, 1).setValue("Tổng cộng chi nhánh").setFontWeight("bold");
      reportSheet.getRange(totalRowIdx, 5, 1, 5).setValues([[sumThuTM, sumChiTM, sumThuCK, sumChiCK, sumLoiNhuan]]);
      reportSheet.getRange(totalRowIdx, 5, 1, 5).setNumberFormat("#,##0").setFontWeight("bold");
      reportSheet.getRange(totalRowIdx, 1, 1, 10).setBackground("#f1f3f4").setFontWeight("bold");

      // Vẽ viền cho bảng (bao gồm cả dòng tổng cộng chi nhánh, tổng cộng 3 dòng tiêu đề/tổng + data)
      reportSheet.getRange(startRow, 1, detailRows.length + 3, 10)
                 .setBorder(true, true, true, true, true, true, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);

      // Điểm bắt đầu bảng tiếp theo
      startRow = totalRowIdx + 3;
    } else {
      // Báo không có giao dịch
      var emptyRange = reportSheet.getRange(startRow + 2, 1, 1, 10);
      emptyRange.merge();
      emptyRange.setValue("Không có giao dịch nào phát sinh trong ngày tại chi nhánh này.")
                 .setFontStyle("italic")
                 .setHorizontalAlignment("center");

      // Vẽ viền cho bảng trống
      reportSheet.getRange(startRow, 1, 3, 10)
                 .setBorder(true, true, true, true, true, true, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);

      startRow = startRow + 5;
    }
  });

  // Áp dụng font Times New Roman 12 cho toàn sheet
  var maxRows = reportSheet.getMaxRows();
  var maxCols = reportSheet.getMaxColumns();
  reportSheet.getRange(1, 1, maxRows, maxCols)
             .setFontFamily("Times New Roman")
             .setFontSize(12);

  // Auto resize các cột hiển thị
  var colsToResize = Math.max(10, totalSummaryColIdx);
  for (var col = 1; col <= colsToResize; col++) {
    reportSheet.autoResizeColumn(col);
  }
}

/**
 * Phân tích cú pháp giá trị ngày từ ô nhập liệu
 *
 * @param {*} val - Giá trị cần phân tích
 * @return {Date|null} Đối tượng Date hoặc null nếu không hợp lệ
 */
function _parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  if (typeof val === "string") {
    var match = val.trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (match) {
      var d = parseInt(match[1], 10);
      var m = parseInt(match[2], 10) - 1;
      var y = parseInt(match[3], 10);
      var date = new Date(y, m, d);
      if (
        date.getDate() === d &&
        date.getMonth() === m &&
        date.getFullYear() === y
      ) {
        return date;
      }
    }
  }
  if (typeof val === "number") {
    var baseDate = new Date(1899, 11, 30);
    return new Date(baseDate.getTime() + val * 24 * 60 * 60 * 1000);
  }
  return null;
}

/**
 * Kiểm tra xem hai ngày có cùng ngày, tháng, năm hay không
 *
 * @param {Date} d1 - Ngày thứ nhất
 * @param {Date} d2 - Ngày thứ hai
 * @return {boolean} True nếu trùng ngày, ngược lại False
 */
function _isSameDay(d1, d2) {
  if (!(d1 instanceof Date) || !(d2 instanceof Date)) return false;
  return (
    d1.getDate() === d2.getDate() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getFullYear() === d2.getFullYear()
  );
}
