/**
 * ============================================================
 * VanTran Mobile — DonHang.gs
 * Quản lý đơn hàng bán điện thoại & phụ kiện
 * ============================================================
 */

/**
 * Tạo đơn hàng mới
 *
 * @param {Object} data - {
 *   maKH, maSP, nguonSP ('Điện thoại'/'Phụ kiện'),
 *   soLuong, donGia, hinhThucBan ('Bán thẳng'/'Trả góp'),
 *   hinhThucThanhToan ('Tiền mặt'/'Chuyển khoản'/'Quẹt thẻ (POS)'),
 *   nguoiBan (MaNV), nguoiHoTro (MaNV), ghiChu, chiNhanh,
 *   // Quà tặng:
 *   maQuaTang, tenQuaTang, coNhanQua ('✓'/'✗'), tienGiamGia,
 *   // Nếu trả góp:
 *   traGop: {traTruoc, soKy, loaiTraGop ('Cửa hàng'/'Công ty tài chính'), congTyTC}
 * }
 * @return {string} Mã đơn hàng mới
 */
function taoDonHang(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (data.items && data.items.length > 0) {
    var rollbackActions = [];
    var createdMaDHs = [];
    try {
      // Tính tổng tiền các phụ kiện
      var sumAllThanhTien = 0;
      data.items.forEach(function(item) {
        sumAllThanhTien += item.soLuong * item.donGia;
      });
      var totalDiscount = Number(data.tienGiamGia) || 0;
      var netPayable = Math.max(0, sumAllThanhTien - totalDiscount);

      var count = data.items.length;
      var discountPerItem = Math.floor(totalDiscount / count);
      var remainingDiscount = totalDiscount - (discountPerItem * count);

      var remainingCK = Number(data.splitChuyenKhoan) || 0;
      var remainingTM = Number(data.splitTienMat) || 0;

      for (var i = 0; i < count; i++) {
        var item = data.items[i];
        var itemDiscount = discountPerItem + (i === 0 ? remainingDiscount : 0);
        var itemThanhTien = item.soLuong * item.donGia - itemDiscount;

        var itemCK = 0;
        var itemTM = 0;
        if (netPayable > 0) {
          var ratio = itemThanhTien / netPayable;
          itemCK = i === count - 1 ? remainingCK : Math.round((Number(data.splitChuyenKhoan) || 0) * ratio);
          itemTM = i === count - 1 ? remainingTM : Math.round((Number(data.splitTienMat) || 0) * ratio);
          remainingCK -= itemCK;
          remainingTM -= itemTM;
        }

        var singleData = {
          chiNhanh: data.chiNhanh,
          maKH: data.maKH,
          maSP: item.maSP,
          nguonSP: "Phụ kiện",
          soLuong: item.soLuong,
          donGia: item.donGia,
          tienGiamGia: itemDiscount,
          hinhThucBan: "Bán thẳng",
          hinhThucThanhToan: data.hinhThucThanhToan,
          nguoiBan: data.nguoiBan,
          nguoiHoTro: data.nguoiHoTro,
          ghiChu: data.ghiChu,
          coNhanQua: "✗"
        };
        
        if (data.hinhThucThanhToan === "Hỗn hợp") {
          singleData.splitChuyenKhoan = itemCK;
          singleData.splitTienMat = itemTM;
        }
        
        var maDH = _taoDonHangSingle(singleData, rollbackActions, ss);
        createdMaDHs.push(maDH);
      }
      
      showToast("✅ Đã tạo thành công " + createdMaDHs.length + " đơn hàng phụ kiện: " + createdMaDHs.join(", "));
      return createdMaDHs.join(", ");
    } catch (e) {
      // Chạy các hành động rollback theo thứ tự ngược
      for (var rIdx = rollbackActions.length - 1; rIdx >= 0; rIdx--) {
        try { rollbackActions[rIdx](); } catch (err) {
          Logger.log("Rollback failed: " + err.message);
        }
      }
      throw e;
    }
  }

  // Đơn hàng đơn lẻ (điện thoại hoặc phụ kiện đơn lẻ)
  var rollbackActions = [];
  try {
    var maDH = _taoDonHangSingle(data, rollbackActions, ss);
    return maDH;
  } catch (e) {
    for (var rIdx = rollbackActions.length - 1; rIdx >= 0; rIdx--) {
      try { rollbackActions[rIdx](); } catch (err) {
        Logger.log("Rollback failed: " + err.message);
      }
    }
    throw e;
  }
}

/**
 * Thực thi logic tạo một đơn hàng đơn lẻ và ghi nhận các hành động rollback
 * @private
 */
function _taoDonHangSingle(data, rollbackActions, ss) {
  var maDH = generateId("DH", SHEET_NAMES.DON_HANG);
  var chiNhanh = data.chiNhanh;

  if (!chiNhanh) {
    throw new Error("Vui lòng chọn chi nhánh tạo đơn hàng!");
  }

  // Lookup thông tin
  var tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, 1, data.maKH, 2) || "";
  var tenSP = "";
  var thuongHieu = "";
  var nguonSP = data.nguonSP || "Điện thoại";

  if (nguonSP === "Điện thoại") {
    tenSP = lookupValue(SHEET_NAMES.DIEN_THOAI, 1, data.maSP, 2) || "";
    thuongHieu = lookupValue(SHEET_NAMES.DIEN_THOAI, 1, data.maSP, 3) || "";
  } else {
    tenSP = lookupValue(SHEET_NAMES.PHU_KIEN, 1, data.maSP, 2) || "";
    thuongHieu = lookupValue(SHEET_NAMES.PHU_KIEN, 1, data.maSP, 4) || "";
  }

  var tenNguoiBan =
    lookupValue(SHEET_NAMES.NHAN_VIEN, 1, data.nguoiBan, 2) || "";
  var coQuyenXuatMay = kiemTraQuyenXuatMay(data.nguoiBan) ? "✓" : "✗";
  var tenNguoiHoTro = data.nguoiHoTro
    ? lookupValue(SHEET_NAMES.NHAN_VIEN, 1, data.nguoiHoTro, 2) || ""
    : "";

  var soLuong = Number(data.soLuong) || 1;
  var donGia = Number(data.donGia) || 0;
  var tienGiamGia = Number(data.tienGiamGia) || 0;
  var thanhTien = soLuong * donGia - tienGiamGia;

  if (nguonSP === "Phụ kiện" && data.hinhThucBan === "Trả góp") {
    throw new Error("Hình thức trả góp chỉ hỗ trợ cho Điện thoại, không áp dụng cho Phụ kiện!");
  }

  // Kiểm tra tồn kho trước khi bán
  if (nguonSP === "Điện thoại") {
    var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
    var phoneRow = -1;

    if (data.imei) {
      phoneRow = findRow(SHEET_NAMES.DIEN_THOAI, 4, data.imei);
    } else {
      // Tự động tìm máy có mã SP tại chi nhánh này và đang "Còn hàng"
      var dtData = dtSheet.getDataRange().getValues();
      for (var i = 1; i < dtData.length; i++) {
        if (String(dtData[i][0]) === data.maSP && 
            String(dtData[i][12]) === chiNhanh && 
            String(dtData[i][10]) === "Còn hàng") {
          phoneRow = i + 1;
          data.imei = String(dtData[i][3]); // Tự động điền IMEI tìm được
          break;
        }
      }
      if (phoneRow === -1) {
        phoneRow = findRow(SHEET_NAMES.DIEN_THOAI, 1, data.maSP);
      }
    }

    if (phoneRow === -1) {
      throw new Error("Không tìm thấy điện thoại " + (data.imei ? "IMEI: " + data.imei : data.maSP) + " trong hệ thống!");
    }

    var trangThaiKho = dtSheet.getRange(phoneRow, 11).getValue();
    if (trangThaiKho !== "Còn hàng") {
      throw new Error(
        "Điện thoại " +
          data.maSP +
          " không còn hàng (Trạng thái: " +
          trangThaiKho +
          ")",
      );
    }
    // Xác nhận máy nằm đúng chi nhánh
    var dtBranch = dtSheet.getRange(phoneRow, 13).getValue();
    if (dtBranch !== chiNhanh) {
      throw new Error(
        "Điện thoại " +
          data.maSP +
          " không thuộc chi nhánh này (Thuộc: " +
          dtBranch +
          ")",
      );
    }
  } else {
    var pkRow = findPhuKienRow(data.maSP, chiNhanh);
    if (pkRow === -1) {
      throw new Error(
        "Phụ kiện " + data.maSP + " không tồn tại ở chi nhánh " + chiNhanh,
      );
    }
    var pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
    var tonKho = Number(pkSheet.getRange(pkRow, 7).getValue()) || 0;
    if (tonKho < soLuong) {
      throw new Error(
        "Phụ kiện " +
          data.maSP +
          " không đủ tồn kho tại chi nhánh! Hiện tại: " +
          tonKho +
          ", cần: " +
          soLuong,
      );
    }
  }

  // Kiểm tra tồn kho quà tặng nếu nhận quà
  var coNhanQua = data.coNhanQua || "✗";
  var maQuaTang = data.maQuaTang || "";
  var tenQuaTang = "";
  if (coNhanQua === "✓" && maQuaTang) {
    var maQuaList = String(maQuaTang).split(",");
    var tenQuaList = [];
    var pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
    
    var codeCounts = {};
    for (var i = 0; i < maQuaList.length; i++) {
      var code = maQuaList[i].trim();
      if (!code) continue;
      codeCounts[code] = (codeCounts[code] || 0) + 1;
    }

    for (var code in codeCounts) {
      var quaRow = findPhuKienRow(code, chiNhanh);
      if (quaRow === -1) {
        throw new Error(
          "Quà tặng " + code + " không tồn tại ở chi nhánh " + chiNhanh,
        );
      }
      var tonQua = Number(pkSheet.getRange(quaRow, 7).getValue()) || 0;
      var requiredQty = codeCounts[code];
      if (tonQua < requiredQty) {
        throw new Error(
          "Quà tặng " + code + " không đủ số lượng tại chi nhánh này! Hiện tại: " + tonQua + ", cần: " + requiredQty,
        );
      }
      var giftName = pkSheet.getRange(quaRow, 2).getValue() || "";
      for (var k = 0; k < requiredQty; k++) {
        tenQuaList.push(giftName);
      }
    }
    tenQuaTang = tenQuaList.join(", ");
  }

  var rowData = [
    maDH,
    new Date(), // NgayBan
    data.maKH || "",
    tenKH,
    data.maSP || "",
    tenSP,
    nguonSP,
    thuongHieu,
    soLuong,
    donGia,
    (data.hinhThucThanhToan === "Hỗn hợp" && data.splitChuyenKhoan !== undefined)
      ? (data.splitChuyenKhoan + "," + data.splitTienMat)
      : thanhTien,
    data.hinhThucBan || "Bán thẳng",
    data.hinhThucThanhToan || "Tiền mặt",
    data.nguoiBan || "",
    tenNguoiBan,
    coQuyenXuatMay,
    data.nguoiHoTro || "",
    tenNguoiHoTro,
    "Hoàn thành",
    (data.ghiChu || "") + (data.imei ? " [IMEI: " + data.imei + "]" : ""),
    chiNhanh, // 21st column
    maQuaTang, // 22nd column
    tenQuaTang, // 23rd column
    coNhanQua, // 24th column
    tienGiamGia, // 25th column
  ];

  var donHangSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  var donHangRow = appendRow(SHEET_NAMES.DON_HANG, rowData);
  rollbackActions.push(function() {
    try {
      donHangSheet.deleteRow(donHangRow);
      clearSheetCache(SHEET_NAMES.DON_HANG);
    } catch (err) {
      Logger.log("Rollback failed to delete order row: " + err.message);
    }
  });

  // Cập nhật kho sản phẩm
  if (nguonSP === "Điện thoại") {
    var trangThaiMoi = "Đã bán";
    if (data.hinhThucBan === "Trả góp" && data.traGop && data.traGop.loaiTraGop === "Cửa hàng") {
      trangThaiMoi = "Đang trả góp";
    }
    updateTrangThaiKhoDT(data.imei || data.maSP, trangThaiMoi);
    rollbackActions.push(function() {
      try {
        updateTrangThaiKhoDT(data.imei || data.maSP, "Còn hàng");
      } catch (err) {
        Logger.log("Rollback failed to restore phone status: " + err.message);
      }
    });
  } else {
    updateTonKhoPhuKien(data.maSP, soLuong, "xuat", chiNhanh);
    rollbackActions.push(function() {
      try {
        updateTonKhoPhuKien(data.maSP, soLuong, "nhap", chiNhanh);
      } catch (err) {
        Logger.log("Rollback failed to restore accessory stock: " + err.message);
      }
    });
  }

  // Trừ kho quà tặng nếu có
  if (coNhanQua === "✓" && maQuaTang) {
    var maQuaList = String(maQuaTang).split(",");
    var updatedGifts = [];
    for (var i = 0; i < maQuaList.length; i++) {
      var code = maQuaList[i].trim();
      if (code) {
        updateTonKhoPhuKien(code, 1, "xuat", chiNhanh);
        updatedGifts.push(code);
      }
    }
    rollbackActions.push(function() {
      for (var j = 0; j < updatedGifts.length; j++) {
        try {
          updateTonKhoPhuKien(updatedGifts[j], 1, "nhap", chiNhanh);
        } catch (err) {
          Logger.log("Rollback failed to restore gift stock: " + err.message);
        }
      }
    });
  }

  // Nếu trả góp → tạo hợp đồng
  if (data.hinhThucBan === "Trả góp" && data.traGop) {
    var maTG = taoHopDongTraGop({
      maDH: maDH,
      maKH: data.maKH,
      tongTien: data.traGop.tongTien !== undefined ? Number(data.traGop.tongTien) : thanhTien,
      traTruoc: data.traGop.traTruoc,
      soKy: Number(data.traGop.soKy) || 1,
      loaiTraGop: data.traGop.loaiTraGop || "Cửa hàng",
      congTyTC: data.traGop.congTyTC || "",
      tienMoiKy: Number(data.traGop.tienMoiKy) || 0,
      chiNhanh: chiNhanh,
    });
    rollbackActions.push(function() {
      try {
        var tgSheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
        var tgRow = findRow(SHEET_NAMES.TRA_GOP, 1, maTG);
        if (tgRow !== -1) {
          tgSheet.deleteRow(tgRow);
          clearSheetCache(SHEET_NAMES.TRA_GOP);
          invalidateDropdownCache(SHEET_NAMES.TRA_GOP);
        }
        var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
        var lstgData = lstgSheet.getDataRange().getValues();
        for (var rowIdx = lstgData.length - 1; rowIdx >= 1; rowIdx--) {
          if (String(lstgData[rowIdx][1]) === maTG) {
            lstgSheet.deleteRow(rowIdx + 1);
          }
        }
        clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);
      } catch (err) {
        Logger.log("Rollback failed to delete installment contract: " + err.message);
      }
    });
  }

  showToast(
    "✅ Đã tạo đơn hàng: " +
      maDH +
      "\n" +
      tenSP +
      " - " +
      formatCurrency(thanhTien) +
      "đ",
  );
  return maDH;
}

/**
 * Lấy đơn hàng theo tháng/năm
 *
 * @param {number} thang - Tháng (1-12)
 * @param {number} nam - Năm
 * @return {Object[]}
 */
function getDonHangTheoThang(thang, nam) {
  var data = getAllData(SHEET_NAMES.DON_HANG);
  var result = [];

  data.forEach(function (row) {
    var ngayBan = row[1];
    if (ngayBan instanceof Date) {
      if (ngayBan.getMonth() + 1 === thang && ngayBan.getFullYear() === nam) {
        result.push({
          MaDH: String(row[0]),
          NgayBan: row[1],
          MaKH: String(row[2]),
          TenKH: String(row[3]),
          MaSP: String(row[4]),
          TenSP: String(row[5]),
          NguonSP: String(row[6]),
          ThuongHieu: String(row[7]),
          SoLuong: Number(row[8]) || 0,
          DonGia: Number(row[9]) || 0,
          ThanhTien: Number(row[10]) || 0,
          HinhThucBan: String(row[11]),
          HinhThucThanhToan: String(row[12]),
          NguoiBan: String(row[13]),
          TenNguoiBan: String(row[14]),
          CoQuyenXuatMay: String(row[15]),
          NguoiHoTro: String(row[16]),
          TenNguoiHoTro: String(row[17]),
          TrangThai: String(row[18]),
          GhiChu: String(row[19]),
          ChiNhanh: String(row[20] || ""),
          MaQuaTang: String(row[21] || ""),
          TenQuaTang: String(row[22] || ""),
          CoNhanQua: String(row[23] || "✗"),
          TienGiamGia: Number(row[24] || 0),
        });
      }
    }
  });

  return result;
}

/**
 * Huỷ đơn hàng
 *
 * @param {string} maDH - Mã đơn hàng
 * @return {boolean}
 */
function huyDonHang(maDH) {
  var row = findRow(SHEET_NAMES.DON_HANG, 1, maDH);
  if (row === -1) {
    showAlert("❌ Lỗi", "Không tìm thấy đơn hàng: " + maDH);
    return false;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  var currentTT = sheet.getRange(row, 19).getValue();

  if (currentTT === "Huỷ") {
    showAlert("⚠️ Cảnh báo", "Đơn hàng " + maDH + " đã bị huỷ trước đó.");
    return false;
  }

  // Hoàn kho sản phẩm chính
  var nguonSP = sheet.getRange(row, 7).getValue();
  var maSP = sheet.getRange(row, 5).getValue();
  var soLuong = Number(sheet.getRange(row, 9).getValue());
  var chiNhanh = sheet.getRange(row, 21).getValue();

  if (nguonSP === "Điện thoại") {
    var ghiChu = sheet.getRange(row, 20).getValue() || "";
    var imeiMatch = ghiChu.match(/\[IMEI:\s*([^\s\]]+)\]/);
    var imei = imeiMatch ? imeiMatch[1] : "";
    updateTrangThaiKhoDT(imei || maSP, "Còn hàng");
  } else {
    updateTonKhoPhuKien(maSP, soLuong, "nhap", chiNhanh);
  }

  // Hoàn kho quà tặng nếu có nhận quà
  var maQuaTang = sheet.getRange(row, 22).getValue();
  var coNhanQua = sheet.getRange(row, 24).getValue();
  if (coNhanQua === "✓" && maQuaTang) {
    var maQuaList = String(maQuaTang).split(",");
    for (var i = 0; i < maQuaList.length; i++) {
      var code = maQuaList[i].trim();
      if (code) {
        updateTonKhoPhuKien(code, 1, "nhap", chiNhanh);
      }
    }
  }

  // Đánh dấu huỷ
  sheet.getRange(row, 19).setValue("Huỷ");

  // Kiểm tra nếu có trả góp → cập nhật trạng thái hợp đồng và các kỳ liên quan
  var hinhThuc = sheet.getRange(row, 12).getValue();
  if (hinhThuc === "Trả góp") {
    huyHopDongTraGop(maDH);
  }

  showToast("✅ Đã huỷ đơn hàng: " + maDH + " và hoàn kho sản phẩm/quà tặng.");
  return true;
}

/**
 * Lấy tổng kết đơn hàng trong tháng
 *
 * @param {number} thang
 * @param {number} nam
 * @return {Object} {tongDonHang, tongDoanhThu, tongDienThoai, tongPhuKien}
 */
function getTongKetDonHang(thang, nam) {
  var donHangs = getDonHangTheoThang(thang, nam);
  var result = {
    tongDonHang: 0,
    tongDoanhThu: 0,
    tongDienThoai: 0,
    tongPhuKien: 0,
  };

  donHangs.forEach(function (dh) {
    if (dh.TrangThai !== "Huỷ" && dh.TrangThai !== "Đổi trả") {
      result.tongDonHang++;
      result.tongDoanhThu += Number(dh.ThanhTien) || 0;
      if (dh.NguonSP === "Điện thoại") {
        result.tongDienThoai++;
      } else {
        result.tongPhuKien += Number(dh.SoLuong) || 0;
      }
    }
  });

  return result;
}

/**
 * Lấy chi tiết một đơn hàng theo mã đơn hàng
 *
 * @param {string} maDH - Mã đơn hàng
 * @return {Object|null}
 */
function getDonHangDetails(maDH) {
  if (!maDH) return null;
  var row = findRow(SHEET_NAMES.DON_HANG, 1, maDH);
  if (row === -1) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  var values = sheet.getRange(row, 1, 1, 25).getValues()[0];

  return {
    maDH: String(values[0]),
    ngayBan: formatDate(
      values[1] instanceof Date ? values[1] : new Date(values[1]),
    ),
    maKH: String(values[2]),
    tenKH: String(values[3]),
    maSP: String(values[4]),
    tenSP: String(values[5]),
    nguonSP: String(values[6]),
    thuongHieu: String(values[7]),
    soLuong: Number(values[8]) || 0,
    donGia: Number(values[9]) || 0,
    thanhTien: Number(values[10]) || 0,
    hinhThucBan: String(values[11]),
    hinhThucThanhToan: String(values[12]),
    nguoiBan: String(values[13]),
    tenNguoiBan: String(values[14]),
    coQuyenXuatMay: String(values[15]),
    nguoiHoTro: String(values[16]),
    tenNguoiHoTro: String(values[17]),
    trangThai: String(values[18]),
    ghiChu: String(values[19]),
    chiNhanh: String(values[20] || ""),
    maQuaTang: String(values[21] || ""),
    tenQuaTang: String(values[22] || ""),
    coNhanQua: String(values[23] || "✗"),
    tienGiamGia: Number(values[24] || 0),
  };
}
