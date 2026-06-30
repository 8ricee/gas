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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_NGAY);
  if (!sheet) return;

  const dateVal = sheet.getRange(3, 2).getValue(); // Cell B3
  const date = _parseDate(dateVal);
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
    const nowStr = Utilities.formatDate(
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let reportSheet = ss.getSheetByName(SHEET_NAMES.BAO_CAO_NGAY);
  if (!reportSheet) {
    reportSheet = ss.insertSheet(SHEET_NAMES.BAO_CAO_NGAY);
  }

  // Đọc danh sách chi nhánh động
  const branches = getBranchesList();
  const fallbackBranch = "Khác/Hệ thống";

  // Khởi tạo accumulator tài chính theo từng chi nhánh
  const branchMetrics = _initBranchMetrics(branches, fallbackBranch);
  const transactions = [];

  // Thu thập giao dịch theo ngày
  _collectDonHang(targetDate, branchMetrics, fallbackBranch, transactions, ss);
  _collectDichVu(targetDate, branchMetrics, fallbackBranch, transactions);
  _collectRepayments(targetDate, branchMetrics, fallbackBranch, transactions);
  _collectNhapKho(targetDate, branchMetrics, fallbackBranch, transactions);
  _collectDoiTra(targetDate, branchMetrics, fallbackBranch, transactions, ss);
  _collectThuMua(targetDate, branchMetrics, fallbackBranch, transactions);
  _collectBaoHanh(targetDate, branchMetrics, fallbackBranch, transactions);

  // Sắp xếp tăng dần theo thời gian
  transactions.sort(function (a, b) {
    return a.time.getTime() - b.time.getTime();
  });

  // --- 2. GHI BÁO CÁO LÊN SHEET ---
  const totalRows = reportSheet.getMaxRows();
  const totalCols = reportSheet.getMaxColumns();
  if (totalRows >= 6) {
    reportSheet.getRange(6, 1, totalRows - 5, totalCols).clear();
  }

  // Tạo tiêu đề
  reportSheet
    .getRange(1, 1)
    .setValue("BÁO CÁO GIAO DỊCH THEO NGÀY")
    .setFontSize(14)
    .setFontWeight("bold");

  reportSheet.getRange(3, 1).setValue("Ngày báo cáo:").setFontWeight("bold");

  const dateCell = reportSheet.getRange(3, 2);
  dateCell
    .setValue(targetDate)
    .setNumberFormat("dd/MM/yyyy")
    .setBackground("#ffffff")
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#cccccc",
      SpreadsheetApp.BorderStyle.SOLID,
    )
    .setHorizontalAlignment("left");

  reportSheet
    .getRange(4, 1)
    .setValue("Trạng thái cập nhật:")
    .setFontWeight("bold");

  // Chuẩn bị danh sách các chi nhánh hiển thị ở bảng tổng hợp
  const activeBranches = branches.slice();
  const fb = branchMetrics[fallbackBranch];
  if (
    fb.doanhSo !== 0 ||
    fb.loiNhuan !== 0 ||
    fb.thuTM !== 0 ||
    fb.chiTM !== 0 ||
    fb.thuCK !== 0 ||
    fb.chiCK !== 0
  ) {
    activeBranches.push(fallbackBranch);
  }
  const numActiveBranches = activeBranches.length;
  const totalSummaryColIdx = numActiveBranches + 2;

  // Ghi bảng tổng hợp chỉ tiêu tài chính
  _writeSummaryTableToSheet(reportSheet, branchMetrics, activeBranches, totalSummaryColIdx);

  // Ghi chi tiết giao dịch các chi nhánh
  _writeDetailedTablesToSheet(reportSheet, transactions, activeBranches, fallbackBranch);

  // Áp dụng font Times New Roman 12 cho toàn sheet
  const maxRows = reportSheet.getMaxRows();
  const maxCols = reportSheet.getMaxColumns();
  reportSheet
    .getRange(1, 1, maxRows, maxCols)
    .setFontFamily("Times New Roman")
    .setFontSize(12);

  // Auto resize các cột hiển thị
  const colsToResize = Math.max(10, totalSummaryColIdx);
  for (let col = 1; col <= colsToResize; col++) {
    reportSheet.autoResizeColumn(col);
  }
}

/**
 * Khởi tạo accumulator tài chính theo từng chi nhánh
 * @private
 */
function _initBranchMetrics(branches, fallbackBranch) {
  const branchMetrics = {};
  branches.forEach(function (b) {
    branchMetrics[b] = {
      doanhSo: 0,
      loiNhuan: 0,
      thuTM: 0,
      chiTM: 0,
      thuCK: 0,
      chiCK: 0,
      congNoCTTC : 0,
    };
  });
  branchMetrics[fallbackBranch] = {
    doanhSo: 0,
    loiNhuan: 0,
    thuTM: 0,
    chiTM: 0,
    thuCK: 0,
    chiCK: 0,
    congNoCTTC : 0,
  };
  return branchMetrics;
}

/**
 * Ghi bảng tổng hợp chỉ tiêu tài chính lên reportSheet
 * @private
 */
function _writeSummaryTableToSheet(reportSheet, branchMetrics, activeBranches, totalSummaryColIdx) {
  const summaryTitleRange = reportSheet.getRange(6, 1, 1, totalSummaryColIdx);
  summaryTitleRange.merge();
  summaryTitleRange
    .setValue("TỔNG HỢP CHỈ TIÊU TÀI CHÍNH")
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  // Tiêu đề cột bảng tổng hợp
  const summaryHeaders = ["Chỉ tiêu"];
  activeBranches.forEach(function (b) {
    summaryHeaders.push(b);
  });
  summaryHeaders.push("Tổng cộng");
  reportSheet.getRange(7, 1, 1, totalSummaryColIdx).setValues([summaryHeaders]);
  reportSheet
    .getRange(7, 1, 1, totalSummaryColIdx)
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setHorizontalAlignment("left");

  // Ghi dữ liệu bảng tổng hợp
  const metricsList = [
    { name: "1. Doanh số bán lẻ", key: "doanhSo" },
    { name: "2. Lợi nhuận gộp", key: "loiNhuan" },
    { name: "3. Tiền mặt Thu", key: "thuTM" },
    { name: "4. Tiền mặt Chi", key: "chiTM" },
    { name: "5. Chuyển khoản Thu", key: "thuCK" },
    { name: "6. Chuyển khoản Chi", key: "chiCK" },
    { name: "7. Phải thu Trả góp (Công nợ)", key: "congNoCTTC" },
    { name: "8. Dòng tiền ròng trong ngày", key: "netCash" },
  ];

  const summaryData = [];
  metricsList.forEach(function (m) {
    const row = [m.name];
    let totalVal = 0;
    activeBranches.forEach(function (b) {
      let val = 0;
      if (m.key === "netCash") {
        val =
          branchMetrics[b].thuTM +
          branchMetrics[b].thuCK -
          (branchMetrics[b].chiTM + branchMetrics[b].chiCK);
      } else {
        val = branchMetrics[b][m.key];
      }
      row.push(val);
      totalVal += val;
    });
    row.push(totalVal);
    summaryData.push(row);
  });

  reportSheet
    .getRange(8, 1, summaryData.length, totalSummaryColIdx)
    .setValues(summaryData);

  // Định dạng số và kiểu chữ bảng tổng hợp
  reportSheet
    .getRange(8, 2, summaryData.length, totalSummaryColIdx - 1)
    .setNumberFormat("#,##0");
  reportSheet
    .getRange(14, 1, 1, totalSummaryColIdx)
    .setFontColor("#1a73e8")
    .setFontWeight("bold"); // Dòng tiền ròng nổi bật
  reportSheet
    .getRange(8, totalSummaryColIdx, summaryData.length, 1)
    .setFontWeight("bold"); // Cột tổng cộng in đậm

  // Vẽ viền cho bảng tổng hợp
  reportSheet
    .getRange(6, 1, summaryData.length + 2, totalSummaryColIdx)
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

/**
 * Ghi các bảng chi tiết giao dịch theo chi nhánh lên reportSheet
 * @private
 */
function _writeDetailedTablesToSheet(reportSheet, transactions, activeBranches, fallbackBranch) {
  let startRow = 17;

  activeBranches.forEach(function (branchName) {
    const branchTxs = transactions.filter(function (tx) {
      return (
        tx.branch === branchName ||
        (branchName === fallbackBranch && !tx.branch)
      );
    });

    // 1. Tiêu đề bảng chi tiết chi nhánh
    const titleRange = reportSheet.getRange(startRow, 1, 1, 11);
    titleRange.merge();
    titleRange
      .setValue("CHI TIẾT GIAO DỊCH - " + branchName.toUpperCase())
      .setBackground("#1a73e8")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("left");

    // 2. Headers cột giao dịch (lược bỏ cột Chi nhánh)
    const detailHeaders = [
      "Mã giao dịch",
      "Thời gian",
      "Loại giao dịch",
      "Khách hàng",
      "Thu Tiền mặt",
      "Chi Tiền mặt",
      "Thu Chuyển khoản",
      "Chi Chuyển khoản",
      "Công nợ",
      "Lợi nhuận",
      "Chi tiết",
    ];
    reportSheet.getRange(startRow + 1, 1, 1, 11).setValues([detailHeaders]);
    reportSheet
      .getRange(startRow + 1, 1, 1, 11)
      .setFontWeight("bold")
      .setBackground("#e8f0fe")
      .setHorizontalAlignment("left");

    // 3. Ghi dữ liệu giao dịch chi tiết
    if (branchTxs.length > 0) {
      const detailRows = [];
      let sumThuTM = 0,
        sumChiTM = 0,
        sumThuCK = 0,
        sumChiCK = 0,
        sumCongNo = 0,
        sumLoiNhuan = 0;

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
          tx.congNo || 0,
          tx.loiNhuan,
          tx.detail,
        ]);
        sumThuTM += tx.thuTM;
        sumChiTM += tx.chiTM;
        sumThuCK += tx.thuCK;
        sumChiCK += tx.chiCK;
        sumCongNo += tx.congNo || 0;
        sumLoiNhuan += tx.loiNhuan;
      });

      reportSheet
        .getRange(startRow + 2, 1, detailRows.length, 11)
        .setValues(detailRows);

      // Format tiền tệ
      reportSheet
        .getRange(startRow + 2, 5, detailRows.length, 6)
        .setNumberFormat("#,##0");

      // Dòng tổng cộng của riêng chi nhánh
      const totalRowIdx = startRow + 2 + detailRows.length;
      reportSheet
        .getRange(totalRowIdx, 1)
        .setValue("Tổng cộng chi nhánh")
        .setFontWeight("bold");
      reportSheet
        .getRange(totalRowIdx, 5, 1, 6)
        .setValues([[sumThuTM, sumChiTM, sumThuCK, sumChiCK, sumCongNo, sumLoiNhuan]]);
      reportSheet
        .getRange(totalRowIdx, 5, 1, 6)
        .setNumberFormat("#,##0")
        .setFontWeight("bold");
      reportSheet
        .getRange(totalRowIdx, 1, 1, 11)
        .setBackground("#f1f3f4")
        .setFontWeight("bold");

      // Vẽ viền cho bảng
      reportSheet
        .getRange(startRow, 1, detailRows.length + 3, 11)
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

      // Điểm bắt đầu bảng tiếp theo
      startRow = totalRowIdx + 3;
    } else {
      // Báo không có giao dịch
      const emptyRange = reportSheet.getRange(startRow + 2, 1, 1, 11);
      emptyRange.merge();
      emptyRange
        .setValue(
          "Không có giao dịch nào phát sinh trong ngày tại chi nhánh này.",
        )
        .setFontStyle("italic")
        .setHorizontalAlignment("left");

      // Vẽ viền cho bảng trống
      reportSheet
        .getRange(startRow, 1, 3, 11)
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

      startRow = startRow + 5;
    }
  });
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
    const match = val.trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (match) {
      const d = parseInt(match[1], 10);
      const m = parseInt(match[2], 10) - 1;
      const y = parseInt(match[3], 10);
      const date = new Date(y, m, d);
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
    const baseDate = new Date(1899, 11, 30);
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
