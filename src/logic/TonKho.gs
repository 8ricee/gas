/**
 * ============================================================
 * VanTran Mobile — TonKho.gs
 * Nghiệp vụ tái cấu trúc và quản lý sheet Tồn kho hệ thống
 * ============================================================
 */

/**
 * Tái cấu trúc báo cáo sheet Tồn kho dựa trên danh sách chi nhánh động
 * Xoay chiều (pivot) số lượng tồn theo chi nhánh và bổ sung cột Tổng tồn
 *
 * @param {Spreadsheet} [ss] - Active spreadsheet
 */
function rebuildTonKhoSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  const tonKhoSheet = ss.getSheetByName(SHEET_NAMES.TON_KHO);
  if (!tonKhoSheet) return;

  const branches = getBranchesList();
  const brands = getBrandsList();

  const N = branches.length || 1;
  const totalSummaryColIdx = 2 + N;

  // Đảm bảo số cột tối thiểu của sheet Tồn kho để tránh lỗi vượt quá phạm vi khi vẽ bảng
  const expectedNeededCols = 3 * N + 15;
  const maxCols = tonKhoSheet.getMaxColumns();
  if (maxCols < expectedNeededCols) {
    tonKhoSheet.insertColumnsAfter(maxCols, expectedNeededCols - maxCols);
  }

  // 1. Chuẩn bị sheet và ô tìm kiếm
  _prepareTonKhoSheet(tonKhoSheet, totalSummaryColIdx);

  // 2. Xây dựng Bảng 1: Tổng số máy điện thoại của mỗi chi nhánh
  const startRowB2B3 = _buildSummaryTable(tonKhoSheet, branches, brands);

  // 3. Xây dựng Bảng 2: Tồn kho điện thoại theo chi nhánh
  const totalPhoneColIdx = _buildPhoneInventoryTable(tonKhoSheet, startRowB2B3, branches);

  // 4. Xây dựng Bảng 3: Tồn kho phụ kiện theo chi nhánh
  const accInfo = _buildAccessoryInventoryTable(tonKhoSheet, startRowB2B3, branches, totalPhoneColIdx);

  // 5. Đảm bảo số cột trong sheet đủ lớn và định dạng font toàn bảng
  const lastRow = tonKhoSheet.getMaxRows();
  let lastCol = tonKhoSheet.getMaxColumns();
  const neededCols = Math.max(accInfo.totalAccColIdx, totalSummaryColIdx);
  if (lastCol < neededCols) {
    tonKhoSheet.insertColumnsAfter(lastCol, neededCols - lastCol);
    lastCol = neededCols;
  }

  const fullRange = tonKhoSheet.getRange(1, 1, lastRow, lastCol);
  fullRange.setFontFamily("Times New Roman").setFontSize(12);

  // Điều chỉnh độ rộng cột
  for (let c = 1; c <= lastCol; c++) {
    if (c === accInfo.spacerColIdx) {
      tonKhoSheet.setColumnWidth(c, 40);
    } else {
      tonKhoSheet.autoResizeColumn(c);
    }
  }
}

/**
 * Chuẩn bị sheet Tồn kho: clear dữ liệu cũ, thiết lập tiêu đề và ô Tìm kiếm
 * @private
 */
function _prepareTonKhoSheet(sheet, totalSummaryColIdx) {
  sheet.clear();
  sheet.clearFormats();

  // Dòng 1: Tiêu đề trang tính
  const sheetTitleRange = sheet.getRange(1, 1, 1, totalSummaryColIdx);
  sheetTitleRange.merge();
  sheetTitleRange.setValue("BÁO CÁO TỒN KHO HỆ THỐNG & TÌM KIẾM");
  sheetTitleRange
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  // Dòng 2: Ô tìm kiếm nhanh
  const searchLabelCell = sheet.getRange(2, 1);
  searchLabelCell
    .setValue("Tìm kiếm:")
    .setFontWeight("bold")
    .setHorizontalAlignment("right")
    .setBackground("#f1f3f4");

  const searchInputCell = sheet.getRange(2, 2);
  searchInputCell
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

  const searchTipRange = sheet.getRange(
    2,
    3,
    1,
    Math.max(1, totalSummaryColIdx - 2),
  );
  searchTipRange.merge();
  searchTipRange
    .setValue("Nhập B2 để tìm kiếm")
    .setFontStyle("italic")
    .setFontColor("#5f6368")
    .setBackground("#f1f3f4")
    .setHorizontalAlignment("left");

  // Dòng 3: Dòng spacer trống
  sheet.setRowHeight(3, 15);
}

/**
 * Xây dựng Bảng 1: Tổng số máy điện thoại của mỗi chi nhánh
 * @private
 * @return {number} Chỉ số dòng tiếp theo sau bảng 1
 */
function _buildSummaryTable(sheet, branches, brands) {
  const N = branches.length || 1;
  const B = brands.length || 1;
  const totalSummaryColIdx = 2 + N;

  const colThuongHieu = columnToLetter(COL_DT.THUONG_HIEU);
  const colChiNhanh = columnToLetter(COL_DT.CHI_NHANH);
  const colTrangThai = columnToLetter(COL_DT.TRANG_THAI_KHO);

  // Dòng 4: Tiêu đề Bảng 1
  const summaryTitleRange = sheet.getRange(4, 1, 1, totalSummaryColIdx);
  summaryTitleRange.merge();
  summaryTitleRange.setValue("TỔNG SỐ MÁY ĐIỆN THOẠI CỦA MỖI CHI NHÁNH");
  summaryTitleRange
    .setFontWeight("bold")
    .setFontSize(12)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  // Dòng 5: Headers Bảng 1
  const summaryHeaders = ["Thương hiệu"];
  branches.forEach(function (branch) {
    summaryHeaders.push(branch);
  });
  summaryHeaders.push("Tổng cộng");
  sheet.getRange(5, 1, 1, totalSummaryColIdx).setValues([summaryHeaders]);

  // Dòng 6 đến 5 + B: Ghi danh sách thương hiệu & công thức
  const brandValues = [];
  brands.forEach(function (brand) {
    brandValues.push([brand]);
  });
  sheet.getRange(6, 1, B, 1).setValues(brandValues);

  // Chèn công thức COUNTIFS cho các chi nhánh & SUM cho cột Tổng cộng
  for (let r = 6; r <= 5 + B; r++) {
    for (let c = 2; c <= 1 + N; c++) {
      const colLetter = columnToLetter(c);
      sheet
        .getRange(r, c)
        .setFormula(
          "=COUNTIFS('Điện thoại'!$" + colThuongHieu + ":$" + colThuongHieu + "; $A" +
            r +
            "; 'Điện thoại'!$" + colChiNhanh + ":$" + colChiNhanh + "; " +
            colLetter +
            "$5; 'Điện thoại'!$" + colTrangThai + ":$" + colTrangThai + "; \"Còn hàng\")",
        );
    }
    const lastBranchColLetter = columnToLetter(1 + N);
    sheet
      .getRange(r, 2 + N)
      .setFormula("=SUM(B" + r + ":" + lastBranchColLetter + r + ")");
  }

  // Dòng 6 + B: Dòng Tổng cộng của bảng 1
  const totalRowIdx = 6 + B;
  sheet.getRange(totalRowIdx, 1).setValue("Tổng cộng");
  for (let c = 2; c <= 1 + N; c++) {
    const colLetter = columnToLetter(c);
    sheet
      .getRange(totalRowIdx, c)
      .setFormula("=SUM(" + colLetter + "6:" + colLetter + (5 + B) + ")");
  }

  const lastBranchColLetter = columnToLetter(1 + N);
  sheet
    .getRange(totalRowIdx, 2 + N)
    .setFormula(
      "=SUM(B" + totalRowIdx + ":" + lastBranchColLetter + totalRowIdx + ")",
    );

  // Format Table 1
  const summaryHeaderRange = sheet.getRange(5, 1, 1, totalSummaryColIdx);
  summaryHeaderRange
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  const summaryTotalRowRange = sheet.getRange(
    totalRowIdx,
    1,
    1,
    totalSummaryColIdx,
  );
  summaryTotalRowRange
    .setBackground("#f1f3f4")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  // Dòng spacer trống là dòng 7 + B
  sheet.setRowHeight(7 + B, 15);

  return 8 + B; // Dòng bắt đầu của Bảng 2 và 3
}

/**
 * Xây dựng Bảng 2: Tồn kho điện thoại theo chi nhánh
 * @private
 * @return {number} Chỉ số cột Tổng tồn của điện thoại
 */
function _buildPhoneInventoryTable(sheet, startRow, branches) {
  const N = branches.length || 1;
  const totalPhoneColIdx = 3 + N + 1; // 3 cột thông tin + N chi nhánh + 1 cột tổng

  // Tiêu đề Bảng 2
  const phoneTitleRange = sheet.getRange(
    startRow,
    1,
    1,
    totalPhoneColIdx,
  );
  phoneTitleRange.merge();
  phoneTitleRange.setValue("TỒN KHO ĐIỆN THOẠI THEO CHI NHÁNH");
  phoneTitleRange
    .setFontWeight("bold")
    .setFontSize(13)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  const maxColLetter = columnToLetter(Math.max(
    COL_DT.MA_DT, COL_DT.TEN_SP, COL_DT.THUONG_HIEU, COL_DT.IMEI, COL_DT.IMEI_2,
    COL_DT.MAU_SAC, COL_DT.DUNG_LUONG, COL_DT.TRANG_THAI_KHO, COL_DT.CHI_NHANH
  ));

  // Công thức QUERY xoay chiều theo chi nhánh (có lọc theo ô Tìm kiếm B2)
  const queryFormula = _buildPhoneQueryFormula(maxColLetter, {
    colTenSP: "Col" + COL_DT.TEN_SP,
    colMauSac: "Col" + COL_DT.MAU_SAC,
    colDungLuong: "Col" + COL_DT.DUNG_LUONG,
    colMaDT: "Col" + COL_DT.MA_DT,
    colTrangThai: "Col" + COL_DT.TRANG_THAI_KHO,
    colThuongHieu: "Col" + COL_DT.THUONG_HIEU,
    colImei: "Col" + COL_DT.IMEI,
    colImei2: "Col" + COL_DT.IMEI_2,
    colChiNhanh: "Col" + COL_DT.CHI_NHANH
  });

  sheet
    .getRange(startRow + 1, 1)
    .setFormula(queryFormula);

  // Đầu cột Tổng tồn
  const phoneTotalHeaderCell = sheet.getRange(
    startRow + 1,
    totalPhoneColIdx,
  );
  phoneTotalHeaderCell
    .setValue("Tổng tồn")
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8");

  // Công thức tính Tổng tồn dùng MAP và OFFSET động từ dòng startRow + 2
  const phoneTotalFormulaCell = sheet.getRange(
    startRow + 2,
    totalPhoneColIdx,
  );
  phoneTotalFormulaCell.setFormula(
    "=MAP(A" +
      (startRow + 2) +
      ':A; LAMBDA(val; IF(val=""; ""; SUM(OFFSET(val; 0; 3; 1; ' +
      N +
      ")))))",
  );

  // Định dạng Header Dòng startRow + 1
  const phoneHeaderRange = sheet.getRange(
    startRow + 1,
    1,
    1,
    totalPhoneColIdx,
  );
  phoneHeaderRange
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  return totalPhoneColIdx;
}

/**
 * Xây dựng Bảng 3: Tồn kho phụ kiện theo chi nhánh
 * @private
 * @return {Object} { spacerColIdx, totalAccColIdx }
 */
function _buildAccessoryInventoryTable(sheet, startRow, branches, totalPhoneColIdx) {
  const N = branches.length || 1;
  const spacerColIdx = totalPhoneColIdx + 1; // Cột khoảng cách trống
  const startAccColIdx = spacerColIdx + 1; // Cột bắt đầu bảng Phụ kiện
  const totalAccColIdx = startAccColIdx + 4 + N; // 4 cột thông tin + N chi nhánh + 1 cột tổng

  const accTitleRange = sheet.getRange(
    startRow,
    startAccColIdx,
    1,
    4 + N + 1,
  );
  accTitleRange.merge();
  accTitleRange.setValue("TỒN KHO PHỤ KIỆN THEO CHI NHÁNH");
  accTitleRange
    .setFontWeight("bold")
    .setFontSize(13)
    .setFontColor("#ffffff")
    .setBackground("#1a73e8")
    .setHorizontalAlignment("left");

  const maxColPKLetter = columnToLetter(Math.max(
    COL_PK.MA_PK, COL_PK.TEN_SP, COL_PK.LOAI_PK, COL_PK.THUONG_HIEU,
    COL_PK.SO_LUONG_TON, COL_PK.TRANG_THAI, COL_PK.CHI_NHANH
  ));

  const startAccColLetter = columnToLetter(startAccColIdx);

  // Công thức QUERY xoay chiều theo chi nhánh (có lọc theo ô Tìm kiếm B2)
  const queryAccFormula = _buildAccessoryQueryFormula(maxColPKLetter, {
    colMaPK: columnToLetter(COL_PK.MA_PK),
    colTenPK: columnToLetter(COL_PK.TEN_SP),
    colLoaiPK: columnToLetter(COL_PK.LOAI_PK),
    colThuongHieuPK: columnToLetter(COL_PK.THUONG_HIEU),
    colSoLuongTon: columnToLetter(COL_PK.SO_LUONG_TON),
    colTrangThai: columnToLetter(COL_PK.TRANG_THAI),
    colChiNhanh: columnToLetter(COL_PK.CHI_NHANH)
  });

  sheet
    .getRange(startRow + 1, startAccColIdx)
    .setFormula(queryAccFormula);

  // Đầu cột Tổng tồn của Phụ kiện
  const accTotalHeaderCell = sheet.getRange(
    startRow + 1,
    totalAccColIdx,
  );
  accTotalHeaderCell
    .setValue("Tổng tồn")
    .setFontWeight("bold")
    .setHorizontalAlignment("left")
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8");

  // Công thức tính Tổng tồn Phụ kiện
  const accTotalFormulaCell = sheet.getRange(
    startRow + 2,
    totalAccColIdx,
  );
  accTotalFormulaCell.setFormula(
    "=MAP(" +
      startAccColLetter +
      (startRow + 2) +
      ":" +
      startAccColLetter +
      '; LAMBDA(val; IF(val=""; ""; SUM(OFFSET(val; 0; 4; 1; ' +
      N +
      ")))))",
  );

  // Định dạng Header Dòng startRow + 1
  const accHeaderRange = sheet.getRange(
    startRow + 1,
    startAccColIdx,
    1,
    4 + N + 1,
  );
  accHeaderRange
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setFontWeight("bold")
    .setHorizontalAlignment("left");

  return { spacerColIdx: spacerColIdx, totalAccColIdx: totalAccColIdx };
}

/**
 * Tạo công thức QUERY cho bảng tồn kho điện thoại
 * @private
 */
function _buildPhoneQueryFormula(maxColLetter, params) {
  const {
    colTenSP,
    colMauSac,
    colDungLuong,
    colMaDT,
    colTrangThai,
    colThuongHieu,
    colImei,
    colImei2,
    colChiNhanh,
    searchCell = "$B$2"
  } = params;

  const selectClause = `select ${colTenSP}, ${colMauSac}, ${colDungLuong}, count(${colMaDT})`;
  const whereBase = `where ${colMaDT} is not null and ${colTrangThai} = 'Còn hàng'`;
  
  const searchFilter = `and (lower(${colTenSP}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colMauSac}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colDungLuong}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colThuongHieu}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colImei}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colImei2}) contains '\"&LOWER(${searchCell})&\"')`;

  const groupPivotClause = `group by ${colTenSP}, ${colMauSac}, ${colDungLuong} pivot ${colChiNhanh}`;

  return `=QUERY(INDEX(TO_TEXT('Điện thoại'!A:${maxColLetter})); "${selectClause} ${whereBase} " & IF(ISBLANK(${searchCell}); ""; "${searchFilter}") & " ${groupPivotClause}"; 1)`;
}

/**
 * Tạo công thức QUERY cho bảng tồn kho phụ kiện
 * @private
 */
function _buildAccessoryQueryFormula(maxColPKLetter, params) {
  const {
    colMaPK,
    colTenPK,
    colLoaiPK,
    colThuongHieuPK,
    colSoLuongTon,
    colTrangThai,
    colChiNhanh,
    searchCell = "$B$2"
  } = params;

  const selectClause = `select ${colMaPK}, ${colTenPK}, ${colLoaiPK}, ${colThuongHieuPK}, sum(${colSoLuongTon})`;
  const whereBase = `where ${colMaPK} is not null and ${colTrangThai} = 'Đang bán' and ${colSoLuongTon} > 0`;

  const searchFilter = `and (lower(${colMaPK}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colTenPK}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colLoaiPK}) contains '\"&LOWER(${searchCell})&\"' ` +
                       `or lower(${colThuongHieuPK}) contains '\"&LOWER(${searchCell})&\"')`;

  const groupPivotClause = `group by ${colMaPK}, ${colTenPK}, ${colLoaiPK}, ${colThuongHieuPK} pivot ${colChiNhanh}`;

  return `=QUERY('Phụ kiện'!A:${maxColPKLetter}; "${selectClause} ${whereBase} " & IF(ISBLANK(${searchCell}); ""; "${searchFilter}") & " ${groupPivotClause}"; 1)`;
}
