/**
 * ============================================================
 * VanTran Mobile — DonHang.gs
 * Quản lý đơn hàng bán điện thoại & phụ kiện
 * ============================================================
 */

/**
 * Lấy index cột (0-indexed) cho Đơn hàng
 * @private
 */
function _getDonHangIndices() {
  initializeColumnEnums();
  return {
    maDH: COL_DH.MA_DH - 1,
    ngayBan: COL_DH.NGAY_BAN - 1,
    maKH: COL_DH.MA_KH - 1,
    tenKH: COL_DH.TEN_KH - 1,
    maSP: COL_DH.MA_SP - 1,
    tenSP: COL_DH.TEN_SP - 1,
    nguonSP: COL_DH.NGUON_SP - 1,
    thuongHieu: COL_DH.THUONG_HIEU - 1,
    soLuong: COL_DH.SO_LUONG - 1,
    donGia: COL_DH.DON_GIA - 1,
    thanhTien: COL_DH.THANH_TIEN - 1,
    hinhThucBan: COL_DH.HINH_THUC_BAN - 1,
    hinhThucTT: COL_DH.HINH_THUC_TT - 1,
    nguoiBan: COL_DH.NGUOI_BAN - 1,
    tenNguoiBan: COL_DH.TEN_NGUOI_BAN - 1,
    coQuyenXuatMay: COL_DH.QUYEN_XUAT - 1,
    nguoiHoTro: COL_DH.NGUOI_HO_TRO - 1,
    tenNguoiHoTro: COL_DH.TEN_NGUOI_HO_TRO - 1,
    trangThai: COL_DH.TRANG_THAI - 1,
    ghiChu: COL_DH.GHI_CHU - 1,
    chiNhanh: COL_DH.CHI_NHANH - 1,
    maQuaTang: COL_DH.MA_QUA_TANG - 1,
    tenQuaTang: COL_DH.TEN_QUA_TANG - 1,
    coNhanQua: COL_DH.CO_NHAN_QUA - 1,
    tienGiamGia: COL_DH.TIEN_GIAM_GIA - 1,
  };
}

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
  return withDocumentLock(function () {
    clearSheetCache();
    initializeColumnEnums();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Validation thanh toán hỗn hợp ở Backend
    var netPayable = 0;
    if (data.items && data.items.length > 0) {
      var sumAllThanhTien = 0;
      data.items.forEach(function (item) {
        sumAllThanhTien += item.soLuong * item.donGia;
      });
      var totalDiscount = Number(data.tienGiamGia) || 0;
      netPayable = Math.max(0, sumAllThanhTien - totalDiscount);
    } else {
      var qty = Number(data.soLuong) || 1;
      var price = Number(data.donGia) || 0;
      var disc = Number(data.tienGiamGia) || 0;
      netPayable = qty * price - disc;
    }

    var expectedPaid = netPayable;
    if (data.hinhThucBan === "Trả góp" && data.traGop) {
      var ttVal = 0;
      if (
        typeof data.traGop.traTruoc === "string" &&
        data.traGop.traTruoc.indexOf(",") !== -1
      ) {
        var parts = data.traGop.traTruoc.split(",");
        ttVal = (Number(parts[0]) || 0) + (Number(parts[1]) || 0);
      } else {
        ttVal = Number(data.traGop.traTruoc) || 0;
      }
      expectedPaid = ttVal;
    }

    if (data.hinhThucThanhToan === "Hỗn hợp") {
      var tongThanhToan =
        (Number(data.splitChuyenKhoan) || 0) + (Number(data.splitTienMat) || 0);
      if (Math.abs(tongThanhToan - expectedPaid) > 1) {
        throw new Error(
          "Lỗi dữ liệu: Tổng tiền mặt (" +
            data.splitTienMat +
            ") và chuyển khoản (" +
            data.splitChuyenKhoan +
            ") không khớp với giá trị cần thanh toán (" +
            expectedPaid +
            ")!",
        );
      }
    }

    if (data.items && data.items.length > 0) {
      var rollbackActions = [];
      var createdMaDHs = [];
      try {
        // Tính tổng tiền các phụ kiện
        var sumAllThanhTien = 0;
        data.items.forEach(function (item) {
          sumAllThanhTien += item.soLuong * item.donGia;
        });
        var totalDiscount = Number(data.tienGiamGia) || 0;
        var netPayable = Math.max(0, sumAllThanhTien - totalDiscount);

        var count = data.items.length;
        var discountPerItem = Math.floor(totalDiscount / count);
        var remainingDiscount = totalDiscount - discountPerItem * count;

        var remainingCK = Number(data.splitChuyenKhoan) || 0;
        var remainingTM = Number(data.splitTienMat) || 0;

        for (var i = 0; i < count; i++) {
          var item = data.items[i];
          var itemDiscount =
            discountPerItem + (i === 0 ? remainingDiscount : 0);
          var itemThanhTien = item.soLuong * item.donGia - itemDiscount;

          var itemCK = 0;
          var itemTM = 0;
          if (netPayable > 0) {
            var ratio = itemThanhTien / netPayable;
            itemCK =
              i === count - 1
                ? remainingCK
                : Math.round((Number(data.splitChuyenKhoan) || 0) * ratio);
            itemTM =
              i === count - 1
                ? remainingTM
                : Math.round((Number(data.splitTienMat) || 0) * ratio);
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
            coNhanQua: "✗",
          };

          if (data.hinhThucThanhToan === "Hỗn hợp") {
            singleData.splitChuyenKhoan = itemCK;
            singleData.splitTienMat = itemTM;
          }

          var maDH = _taoDonHangSingle(singleData, rollbackActions, ss);
          createdMaDHs.push(maDH);
        }

        showToast(
          "✅ Đã tạo thành công " +
            createdMaDHs.length +
            " đơn hàng phụ kiện: " +
            createdMaDHs.join(", "),
        );
        return createdMaDHs.join(", ");
      } catch (e) {
        // Chạy các hành động rollback theo thứ tự ngược
        for (var rIdx = rollbackActions.length - 1; rIdx >= 0; rIdx--) {
          try {
            rollbackActions[rIdx]();
          } catch (err) {
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
        try {
          rollbackActions[rIdx]();
        } catch (err) {
          Logger.log("Rollback failed: " + err.message);
        }
      }
      throw e;
    }
  });
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

  // Lookup thông tin (Tự động thêm KH nếu chưa có)
  var tenKH = ensureKhachHangExists(data.maKH, data.tenKH);
  var tenSP = "";
  var thuongHieu = "";
  var nguonSP = data.nguonSP || "Điện thoại";

  if (nguonSP === "Điện thoại") {
    tenSP =
      lookupValue(
        SHEET_NAMES.DIEN_THOAI,
        COL_DT.MA_DT,
        data.maSP,
        COL_DT.TEN_SP,
      ) || "";
    thuongHieu =
      lookupValue(
        SHEET_NAMES.DIEN_THOAI,
        COL_DT.MA_DT,
        data.maSP,
        COL_DT.THUONG_HIEU,
      ) || "";
  } else {
    tenSP =
      lookupValue(
        SHEET_NAMES.PHU_KIEN,
        COL_PK.MA_PK,
        data.maSP,
        COL_PK.TEN_SP,
      ) || "";
    thuongHieu =
      lookupValue(
        SHEET_NAMES.PHU_KIEN,
        COL_PK.MA_PK,
        data.maSP,
        COL_PK.THUONG_HIEU,
      ) || "";
  }

  var tenNguoiBan = getNhanVienName(data.nguoiBan);
  var coQuyenXuatMay = kiemTraQuyenXuatMay(data.nguoiBan) ? "✓" : "✗";
  var tenNguoiHoTro = getNhanVienName(data.nguoiHoTro);

  var soLuong = Number(data.soLuong) || 1;
  var donGia = Number(data.donGia) || 0;
  var tienGiamGia = Number(data.tienGiamGia) || 0;
  var thanhTien = soLuong * donGia - tienGiamGia;

  if (nguonSP === "Phụ kiện" && data.hinhThucBan === "Trả góp") {
    throw new Error(
      "Hình thức trả góp chỉ hỗ trợ cho Điện thoại, không áp dụng cho Phụ kiện!",
    );
  }

  // Kiểm tra tồn kho trước khi bán
  if (nguonSP === "Điện thoại") {
    var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
    var phoneRow = -1;

    if (data.imei) {
      phoneRow = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.IMEI, data.imei);
      if (phoneRow === -1 && COL_DT.IMEI_2) {
        phoneRow = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.IMEI_2, data.imei);
      }
    } else {
      // Tự động tìm máy có mã SP tại chi nhánh này và đang "Còn hàng"
      var dtData = getAllData(SHEET_NAMES.DIEN_THOAI);
      var maDTIdx = COL_DT.MA_DT - 1;
      var chiNhanhIdx = COL_DT.CHI_NHANH - 1;
      var trangThaiKhoIdx = COL_DT.TRANG_THAI_KHO - 1;
      var imeiIdx = COL_DT.IMEI - 1;

      for (var i = 0; i < dtData.length; i++) {
        if (
          String(dtData[i][maDTIdx]) === data.maSP &&
          String(dtData[i][chiNhanhIdx]) === chiNhanh &&
          String(dtData[i][trangThaiKhoIdx]) === "Còn hàng"
        ) {
          phoneRow = i + 2;
          data.imei = String(dtData[i][imeiIdx]); // Tự động điền IMEI tìm được
          break;
        }
      }
      if (phoneRow === -1) {
        phoneRow = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, data.maSP);
      }
    }

    if (phoneRow === -1) {
      throw new Error(
        "Không tìm thấy điện thoại " +
          (data.imei ? "IMEI: " + data.imei : data.maSP) +
          " trong hệ thống!",
      );
    }

    var trangThaiKho = dtSheet
      .getRange(phoneRow, COL_DT.TRANG_THAI_KHO)
      .getValue();
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
    var dtBranch = dtSheet.getRange(phoneRow, COL_DT.CHI_NHANH).getValue();
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
    var tonKho =
      Number(pkSheet.getRange(pkRow, COL_PK.SO_LUONG_TON).getValue()) || 0;
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
      var tonQua =
        Number(pkSheet.getRange(quaRow, COL_PK.SO_LUONG_TON).getValue()) || 0;
      var requiredQty = codeCounts[code];
      if (tonQua < requiredQty) {
        throw new Error(
          "Quà tặng " +
            code +
            " không đủ số lượng tại chi nhánh này! Hiện tại: " +
            tonQua +
            ", cần: " +
            requiredQty,
        );
      }
      var giftName = pkSheet.getRange(quaRow, COL_PK.TEN_SP).getValue() || "";
      for (var k = 0; k < requiredQty; k++) {
        tenQuaList.push(giftName);
      }
    }
    tenQuaTang = tenQuaList.join(", ");
  }

  var rowData = [];
  rowData[COL_DH.MA_DH - 1] = maDH;
  rowData[COL_DH.NGAY_BAN - 1] = new Date();
  rowData[COL_DH.MA_KH - 1] = data.maKH || "";
  rowData[COL_DH.TEN_KH - 1] = tenKH;
  rowData[COL_DH.MA_SP - 1] = data.maSP || "";
  rowData[COL_DH.TEN_SP - 1] = tenSP;
  rowData[COL_DH.NGUON_SP - 1] = nguonSP;
  rowData[COL_DH.THUONG_HIEU - 1] = thuongHieu;
  rowData[COL_DH.SO_LUONG - 1] = soLuong;
  rowData[COL_DH.DON_GIA - 1] = donGia;
  rowData[COL_DH.THANH_TIEN - 1] = thanhTien;
  rowData[COL_DH.HINH_THUC_BAN - 1] = data.hinhThucBan || "Bán thẳng";
  rowData[COL_DH.HINH_THUC_TT - 1] = data.hinhThucThanhToan || "Tiền mặt";
  rowData[COL_DH.NGUOI_BAN - 1] = data.nguoiBan || "";
  rowData[COL_DH.TEN_NGUOI_BAN - 1] = tenNguoiBan;
  rowData[COL_DH.QUYEN_XUAT - 1] = coQuyenXuatMay;
  rowData[COL_DH.NGUOI_HO_TRO - 1] = data.nguoiHoTro || "";
  rowData[COL_DH.TEN_NGUOI_HO_TRO - 1] = tenNguoiHoTro;
  rowData[COL_DH.TRANG_THAI - 1] = "Hoàn thành";
  rowData[COL_DH.GHI_CHU - 1] =
    (data.ghiChu || "") + (data.imei ? " [IMEI: " + data.imei + "]" : "");
  rowData[COL_DH.CHI_NHANH - 1] = chiNhanh;
  rowData[COL_DH.MA_QUA_TANG - 1] = maQuaTang;
  rowData[COL_DH.TEN_QUA_TANG - 1] = tenQuaTang;
  rowData[COL_DH.CO_NHAN_QUA - 1] = coNhanQua;
  rowData[COL_DH.TIEN_GIAM_GIA - 1] = tienGiamGia;

  var paidAmount = (data.hinhThucBan === "Trả góp" && data.traGop) ? (Number(data.traGop.traTruoc) || 0) : thanhTien;
  var splitResult = calculatePaymentSplit(data, paidAmount);
  var tienMat = splitResult.tienMat;
  var chuyenKhoan = splitResult.chuyenKhoan;
  rowData[COL_DH.TIEN_MAT - 1] = tienMat;
  rowData[COL_DH.CHUYEN_KHOAN - 1] = chuyenKhoan;

  var donHangSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  var donHangRow = appendRow(SHEET_NAMES.DON_HANG, rowData);
  (function (capturedRow) {
    rollbackActions.push(function () {
      try {
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DON_HANG);
        sheet.deleteRow(capturedRow);
        clearSheetCache(SHEET_NAMES.DON_HANG);
      } catch (err) {
        Logger.log("Rollback failed to delete order row: " + err.message);
      }
    });
  })(donHangRow);

  // Cập nhật kho sản phẩm
  if (nguonSP === "Điện thoại") {
    var trangThaiMoi = "Đã bán";
    if (
      data.hinhThucBan === "Trả góp" &&
      data.traGop &&
      data.traGop.loaiTraGop === "Cửa hàng"
    ) {
      trangThaiMoi = "Đang trả góp";
    }
    updateTrangThaiKhoDT(data.imei || data.maSP, trangThaiMoi);
    (function (key) {
      rollbackActions.push(function () {
        try {
          updateTrangThaiKhoDT(key, "Còn hàng");
        } catch (err) {
          Logger.log("Rollback failed to restore phone status: " + err.message);
        }
      });
    })(data.imei || data.maSP);
  } else {
    updateTonKhoPhuKien(data.maSP, soLuong, "xuat", chiNhanh);
    (function (maSP, qty, branch) {
      rollbackActions.push(function () {
        try {
          updateTonKhoPhuKien(maSP, qty, "nhap", branch);
        } catch (err) {
          Logger.log(
            "Rollback failed to restore accessory stock: " + err.message,
          );
        }
      });
    })(data.maSP, soLuong, chiNhanh);
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
    rollbackActions.push(function () {
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
      tongTien:
        data.traGop.tongTien !== undefined
          ? Number(data.traGop.tongTien)
          : thanhTien,
      traTruoc: data.traGop.traTruoc,
      soKy: Number(data.traGop.soKy) || 1,
      loaiTraGop: data.traGop.loaiTraGop || "Cửa hàng",
      congTyTC: data.traGop.congTyTC || "",
      tienMoiKy: Number(data.traGop.tienMoiKy) || 0,
      chiNhanh: chiNhanh,
      hinhThucThanhToan: data.hinhThucThanhToan,
      splitTienMat: data.splitTienMat,
      splitChuyenKhoan: data.splitChuyenKhoan,
    });
    (function (capturedMaTG) {
      rollbackActions.push(function () {
        try {
          var ss = SpreadsheetApp.getActiveSpreadsheet();
          var tgSheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
          var tgRow = findRow(SHEET_NAMES.TRA_GOP, 1, capturedMaTG);
          if (tgRow !== -1) {
            tgSheet.deleteRow(tgRow);
            clearSheetCache(SHEET_NAMES.TRA_GOP);
            invalidateDropdownCache(SHEET_NAMES.TRA_GOP);
          }
          var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
          var lstgData = lstgSheet.getDataRange().getValues();
          for (var rowIdx = lstgData.length - 1; rowIdx >= 1; rowIdx--) {
            if (String(lstgData[rowIdx][1]) === capturedMaTG) {
              lstgSheet.deleteRow(rowIdx + 1);
            }
          }
          clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);
        } catch (err) {
          Logger.log(
            "Rollback failed to delete installment contract: " + err.message,
          );
        }
      });
    })(maTG);
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
  var c = _getDonHangIndices();
  var data = getAllData(SHEET_NAMES.DON_HANG);
  var result = [];

  data.forEach(function (row) {
    var ngayBan = row[c.ngayBan];
    if (isSameMonthYear(ngayBan, thang, nam)) {
        result.push({
          MaDH: String(row[c.maDH]),
          NgayBan: row[c.ngayBan],
          MaKH: String(row[c.maKH]),
          TenKH: String(row[c.tenKH]),
          MaSP: String(row[c.maSP]),
          TenSP: String(row[c.tenSP]),
          NguonSP: String(row[c.nguonSP]),
          ThuongHieu: String(row[c.thuongHieu]),
          SoLuong: Number(row[c.soLuong]) || 0,
          DonGia: Number(row[c.donGia]) || 0,
          ThanhTien: Number(row[c.thanhTien]) || 0,
          HinhThucBan: String(row[c.hinhThucBan]),
          HinhThucThanhToan: String(row[c.hinhThucTT]),
          NguoiBan: String(row[c.nguoiBan]),
          TenNguoiBan: String(row[c.tenNguoiBan]),
          CoQuyenXuatMay: String(row[c.coQuyenXuatMay]),
          NguoiHoTro: String(row[c.nguoiHoTro]),
          TenNguoiHoTro: String(row[c.tenNguoiHoTro]),
          TrangThai: String(row[c.trangThai]),
          GhiChu: String(row[c.ghiChu]),
          ChiNhanh: String(row[c.chiNhanh] || ""),
          MaQuaTang: String(row[c.maQuaTang] || ""),
          TenQuaTang: String(row[c.tenQuaTang] || ""),
          CoNhanQua: String(row[c.coNhanQua] || "✗"),
          TienGiamGia: Number(row[c.tienGiamGia] || 0),
        });
      }
    }
  );

  return result;
}

/**
 * Huỷ đơn hàng
 *
 * @param {string} maDH - Mã đơn hàng
 * @return {boolean}
 */
function huyDonHang(maDH) {
  return withDocumentLock(function () {
    initializeColumnEnums();
    var row = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH);
    if (row === -1) {
      showAlert("❌ Lỗi", "Không tìm thấy đơn hàng: " + maDH);
      return false;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
    var currentTT = sheet.getRange(row, COL_DH.TRANG_THAI).getValue();

    if (currentTT === "Huỷ") {
      showAlert("⚠️ Cảnh báo", "Đơn hàng " + maDH + " đã bị huỷ trước đó.");
      return false;
    }

    // Hoàn kho sản phẩm chính
    var nguonSP = sheet.getRange(row, COL_DH.NGUON_SP).getValue();
    var maSP = sheet.getRange(row, COL_DH.MA_SP).getValue();
    var soLuong = Number(sheet.getRange(row, COL_DH.SO_LUONG).getValue());
    var chiNhanh = sheet.getRange(row, COL_DH.CHI_NHANH).getValue();

    if (nguonSP === "Điện thoại") {
      var ghiChu = sheet.getRange(row, COL_DH.GHI_CHU).getValue() || "";
      var imeiMatch = ghiChu.match(/\[IMEI:\s*([^\s\]]+)\]/);
      var imei = imeiMatch ? imeiMatch[1] : "";
      updateTrangThaiKhoDT(imei || maSP, "Còn hàng");
    } else {
      updateTonKhoPhuKien(maSP, soLuong, "nhap", chiNhanh);
    }

    // Hoàn kho quà tặng nếu có nhận quà
    var maQuaTang = sheet.getRange(row, COL_DH.MA_QUA_TANG).getValue();
    var coNhanQua = sheet.getRange(row, COL_DH.CO_NHAN_QUA).getValue();
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
    sheet.getRange(row, COL_DH.TRANG_THAI).setValue("Huỷ");

    // Kiểm tra nếu có trả góp → cập nhật trạng thái hợp đồng và các kỳ liên quan
    var hinhThuc = sheet.getRange(row, COL_DH.HINH_THUC_BAN).getValue();
    if (hinhThuc === "Trả góp") {
      huyHopDongTraGop(maDH);
    }

    showToast("✅ Đã huỷ đơn hàng: " + maDH + " và hoàn kho sản phẩm/quà tặng.");
    return true;
  });
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
  var c = _getDonHangIndices();
  var row = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH);
  if (row === -1) return null;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  var values = sheet.getRange(row, 1, 1, COL_DH.CHUYEN_KHOAN).getValues()[0];

  return {
    maDH: String(values[c.maDH]),
    ngayBan: formatDate(
      values[c.ngayBan] instanceof Date ? values[c.ngayBan] : new Date(values[c.ngayBan]),
    ),
    maKH: String(values[c.maKH]),
    tenKH: String(values[c.tenKH]),
    maSP: String(values[c.maSP]),
    tenSP: String(values[c.tenSP]),
    nguonSP: String(values[c.nguonSP]),
    thuongHieu: String(values[c.thuongHieu]),
    soLuong: Number(values[c.soLuong]) || 0,
    donGia: Number(values[c.donGia]) || 0,
    thanhTien: Number(values[c.thanhTien]) || 0,
    hinhThucBan: String(values[c.hinhThucBan]),
    hinhThucThanhToan: String(values[c.hinhThucTT]),
    nguoiBan: String(values[c.nguoiBan]),
    tenNguoiBan: String(values[c.tenNguoiBan]),
    coQuyenXuatMay: String(values[c.coQuyenXuatMay]),
    nguoiHoTro: String(values[c.nguoiHoTro]),
    tenNguoiHoTro: String(values[c.tenNguoiHoTro]),
    trangThai: String(values[c.trangThai]),
    ghiChu: String(values[c.ghiChu]),
    chiNhanh: String(values[c.chiNhanh] || ""),
    maQuaTang: String(values[c.maQuaTang] || ""),
    tenQuaTang: String(values[c.tenQuaTang] || ""),
    coNhanQua: String(values[c.coNhanQua] || "✗"),
    tienGiamGia: Number(values[c.tienGiamGia] || 0),
  };
}
