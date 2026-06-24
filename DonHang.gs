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
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Validation thanh toán hỗn hợp ở Backend
    const netPayable = _calcNetPayable(data);

    let expectedPaid = netPayable;
    if (data.hinhThucBan === "Trả góp" && data.traGop) {
      let ttVal = 0;
      if (
        typeof data.traGop.traTruoc === "string" &&
        data.traGop.traTruoc.indexOf(",") !== -1
      ) {
        const parts = data.traGop.traTruoc.split(",");
        ttVal = (Number(parts[0]) || 0) + (Number(parts[1]) || 0);
      } else {
        ttVal = Number(data.traGop.traTruoc) || 0;
      }
      expectedPaid = ttVal;
    }

    if (data.hinhThucThanhToan === "Hỗn hợp") {
      const tongThanhToan =
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

    return _executeWithRollback(function (rollbackActions) {
      if (data.items && data.items.length > 0) {
        // Tính tổng tiền các phụ kiện
        const totalDiscount = Number(data.tienGiamGia) || 0;
        const netPayableVal = netPayable;

        const count = data.items.length;
        const discountPerItem = Math.floor(totalDiscount / count);
        const remainingDiscount = totalDiscount - discountPerItem * count;

        let remainingCK = Number(data.splitChuyenKhoan) || 0;
        let remainingTM = Number(data.splitTienMat) || 0;

        const createdMaDHs = [];
        for (let i = 0; i < count; i++) {
          const item = data.items[i];
          const itemDiscount =
            discountPerItem + (i === 0 ? remainingDiscount : 0);
          const itemThanhTien = item.soLuong * item.donGia - itemDiscount;

          let itemCK = 0;
          let itemTM = 0;
          if (netPayableVal > 0) {
            const ratio = itemThanhTien / netPayableVal;
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

          const singleData = {
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

          const maDH = _taoDonHangSingle(singleData, rollbackActions, ss);
          createdMaDHs.push(maDH);
        }

        showToast(
          "✅ Đã tạo thành công " +
            createdMaDHs.length +
            " đơn hàng phụ kiện: " +
            createdMaDHs.join(", "),
        );
        return createdMaDHs.join(", ");
      } else {
        // Đơn hàng đơn lẻ (điện thoại hoặc phụ kiện đơn lẻ)
        const maDH = _taoDonHangSingle(data, rollbackActions, ss);
        return maDH;
      }
    });
  });
}

/**
 * Tính toán số tiền thực nhận của đơn hàng (sau giảm giá)
 * @private
 */
function _calcNetPayable(data) {
  if (data.items && data.items.length > 0) {
    const total = data.items.reduce((sum, i) => sum + i.soLuong * i.donGia, 0);
    return Math.max(0, total - (Number(data.tienGiamGia) || 0));
  }
  return Math.max(
    0,
    (Number(data.soLuong) || 1) * (Number(data.donGia) || 0) -
      (Number(data.tienGiamGia) || 0),
  );
}

/**
 * Thực thi logic có bọc cơ chế rollback tự động khi xảy ra lỗi
 * @private
 */
function _executeWithRollback(fn) {
  const rollbackActions = [];
  try {
    return fn(rollbackActions);
  } catch (e) {
    for (let i = rollbackActions.length - 1; i >= 0; i--) {
      try {
        rollbackActions[i]();
      } catch (err) {
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
  const maDH = generateId("DH", SHEET_NAMES.DON_HANG);
  const chiNhanh = data.chiNhanh;

  if (!chiNhanh) {
    throw new Error("Vui lòng chọn chi nhánh tạo đơn hàng!");
  }

  // Lookup thông tin (Tự động thêm KH nếu chưa có)
  const tenKH = ensureKhachHangExists(data.maKH, data.tenKH);
  const nguonSP = data.nguonSP || "Điện thoại";
  const strategy = ProductStrategy[nguonSP];
  if (!strategy) {
    throw new Error("Loại sản phẩm không hợp lệ: " + nguonSP);
  }

  const prodInfo = strategy.lookup(data.maSP);
  const tenSP = prodInfo.tenSP;
  const thuongHieu = prodInfo.thuongHieu;

  const tenNguoiBan = getNhanVienName(data.nguoiBan);
  const coQuyenXuatMay = kiemTraQuyenXuatMay(data.nguoiBan) ? "✓" : "✗";
  const tenNguoiHoTro = getNhanVienName(data.nguoiHoTro);

  const soLuong = Number(data.soLuong) || 1;
  const donGia = Number(data.donGia) || 0;
  const tienGiamGia = Number(data.tienGiamGia) || 0;
  const thanhTien = soLuong * donGia - tienGiamGia;

  if (nguonSP === "Phụ kiện" && data.hinhThucBan === "Trả góp") {
    throw new Error(
      "Hình thức trả góp chỉ hỗ trợ cho Điện thoại, không áp dụng cho Phụ kiện!",
    );
  }

  // Kiểm tra tồn kho trước khi bán
  strategy.checkStock(data, chiNhanh, soLuong, ss);

  // Kiểm tra tồn kho quà tặng nếu nhận quà
  const coNhanQua = data.coNhanQua || "✗";
  const maQuaTang = data.maQuaTang || "";
  let tenQuaTang = "";
  if (coNhanQua === "✓" && maQuaTang) {
    const maQuaList = String(maQuaTang).split(",");
    const tenQuaList = [];
    const pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);

    const codeCounts = {};
    for (let i = 0; i < maQuaList.length; i++) {
      const code = maQuaList[i].trim();
      if (!code) continue;
      codeCounts[code] = (codeCounts[code] || 0) + 1;
    }

    for (const code in codeCounts) {
      const quaRow = findPhuKienRow(code, chiNhanh);
      if (quaRow === -1) {
        throw new Error(
          "Quà tặng " + code + " không tồn tại ở chi nhánh " + chiNhanh,
        );
      }
      const tonQua =
        Number(pkSheet.getRange(quaRow, COL_PK.SO_LUONG_TON).getValue()) || 0;
      const requiredQty = codeCounts[code];
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
      const giftName = pkSheet.getRange(quaRow, COL_PK.TEN_SP).getValue() || "";
      for (let k = 0; k < requiredQty; k++) {
        tenQuaList.push(giftName);
      }
    }
    tenQuaTang = tenQuaList.join(", ");
  }

  const rowData = [];
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

  const paidAmount = (data.hinhThucBan === "Trả góp" && data.traGop) ? (Number(data.traGop.traTruoc) || 0) : thanhTien;
  const splitResult = calculatePaymentSplit(data, paidAmount);
  const tienMat = splitResult.tienMat;
  const chuyenKhoan = splitResult.chuyenKhoan;
  rowData[COL_DH.TIEN_MAT - 1] = tienMat;
  rowData[COL_DH.CHUYEN_KHOAN - 1] = chuyenKhoan;

  const donHangSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  const donHangRow = appendRow(SHEET_NAMES.DON_HANG, rowData);
  (function (capturedRow) {
    rollbackActions.push(function () {
      try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.DON_HANG);
        sheet.deleteRow(capturedRow);
        clearSheetCache(SHEET_NAMES.DON_HANG);
      } catch (err) {
        Logger.log("Rollback failed to delete order row: " + err.message);
      }
    });
  })(donHangRow);

  // Cập nhật kho sản phẩm
  strategy.updateStock(data, chiNhanh, soLuong, rollbackActions, ss);

  // Trừ kho quà tặng nếu có
  if (coNhanQua === "✓" && maQuaTang) {
    const maQuaList = String(maQuaTang).split(",");
    const updatedGifts = [];
    for (let i = 0; i < maQuaList.length; i++) {
      const code = maQuaList[i].trim();
      if (code) {
        updateTonKhoPhuKien(code, 1, "xuat", chiNhanh);
        updatedGifts.push(code);
      }
    }
    rollbackActions.push(function () {
      for (let j = 0; j < updatedGifts.length; j++) {
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
    const maTG = taoHopDongTraGop({
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
          const ss = SpreadsheetApp.getActiveSpreadsheet();
          const tgSheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
          const tgRow = findRow(SHEET_NAMES.TRA_GOP, 1, capturedMaTG);
          if (tgRow !== -1) {
            tgSheet.deleteRow(tgRow);
            clearSheetCache(SHEET_NAMES.TRA_GOP);
            invalidateDropdownCache(SHEET_NAMES.TRA_GOP);
          }
          const lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
          const lstgData = lstgSheet.getDataRange().getValues();
          for (let rowIdx = lstgData.length - 1; rowIdx >= 1; rowIdx--) {
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
  const c = _getDonHangIndices();
  const data = getAllData(SHEET_NAMES.DON_HANG);
  const result = [];

  data.forEach(function (row) {
    const ngayBan = row[c.ngayBan];
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
  return withDocumentLock(function () {
    const row = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH);
    if (row === -1) {
      showAlert("❌ Lỗi", "Không tìm thấy đơn hàng: " + maDH);
      return false;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
    const maxCols = sheet.getLastColumn();
    const rowValues = sheet.getRange(row, 1, 1, maxCols).getValues()[0];

    const currentTT = rowValues[COL_DH.TRANG_THAI - 1];

    if (currentTT === "Huỷ") {
      showAlert("⚠️ Cảnh báo", "Đơn hàng " + maDH + " đã bị huỷ trước đó.");
      return false;
    }

    return _executeWithRollback(function (rollbackActions) {
      // Hoàn kho sản phẩm chính
      const nguonSP = rowValues[COL_DH.NGUON_SP - 1];
      const maSP = rowValues[COL_DH.MA_SP - 1];
      const soLuong = Number(rowValues[COL_DH.SO_LUONG - 1]);
      const chiNhanh = rowValues[COL_DH.CHI_NHANH - 1];

      if (nguonSP === "Điện thoại") {
        const ghiChu = rowValues[COL_DH.GHI_CHU - 1] || "";
        const imeiMatch = ghiChu.match(/\[IMEI:\s*([^\s\]]+)\]/);
        const imei = imeiMatch ? imeiMatch[1] : "";
        const targetPhoneKey = imei || maSP;

        let oldPhoneStatus = lookupValue(
          SHEET_NAMES.DIEN_THOAI,
          COL_DT.IMEI,
          targetPhoneKey,
          COL_DT.TRANG_THAI_KHO
        );
        if (oldPhoneStatus === null) {
          oldPhoneStatus = lookupValue(
            SHEET_NAMES.DIEN_THOAI,
            COL_DT.MA_DT,
            targetPhoneKey,
            COL_DT.TRANG_THAI_KHO
          );
        }

        updateTrangThaiKhoDT(targetPhoneKey, "Còn hàng");

        (function (key, status) {
          rollbackActions.push(function () {
            try {
              if (status !== null) {
                updateTrangThaiKhoDT(key, status);
              }
            } catch (err) {
              Logger.log("Rollback failed to restore phone status: " + err.message);
            }
          });
        })(targetPhoneKey, oldPhoneStatus);
      } else {
        updateTonKhoPhuKien(maSP, soLuong, "nhap", chiNhanh);

        (function (code, qty, branch) {
          rollbackActions.push(function () {
            try {
              updateTonKhoPhuKien(code, qty, "xuat", branch);
            } catch (err) {
              Logger.log("Rollback failed to restore accessory stock: " + err.message);
            }
          });
        })(maSP, soLuong, chiNhanh);
      }

      // Hoàn kho quà tặng nếu có nhận quà
      const maQuaTang = rowValues[COL_DH.MA_QUA_TANG - 1];
      const coNhanQua = rowValues[COL_DH.CO_NHAN_QUA - 1];
      if (coNhanQua === "✓" && maQuaTang) {
        const maQuaList = String(maQuaTang).split(",");
        const returnedGifts = [];
        for (let i = 0; i < maQuaList.length; i++) {
          const code = maQuaList[i].trim();
          if (code) {
            updateTonKhoPhuKien(code, 1, "nhap", chiNhanh);
            returnedGifts.push(code);
          }
        }
        (function (gifts, branch) {
          rollbackActions.push(function () {
            for (let i = 0; i < gifts.length; i++) {
              try {
                updateTonKhoPhuKien(gifts[i], 1, "xuat", branch);
              } catch (err) {
                Logger.log("Rollback failed to restore gift stock: " + err.message);
              }
            }
          });
        })(returnedGifts, chiNhanh);
      }

      // Đánh dấu huỷ
      updateCell(SHEET_NAMES.DON_HANG, row, COL_DH.TRANG_THAI, "Huỷ");

      (function (rowNum, oldStatus) {
        rollbackActions.push(function () {
          try {
            updateCell(SHEET_NAMES.DON_HANG, rowNum, COL_DH.TRANG_THAI, oldStatus);
          } catch (err) {
            Logger.log("Rollback failed to restore order status: " + err.message);
          }
        });
      })(row, currentTT);

      // Kiểm tra nếu có trả góp → cập nhật trạng thái hợp đồng và các kỳ liên quan
      const hinhThuc = rowValues[COL_DH.HINH_THUC_BAN - 1];
      if (hinhThuc === "Trả góp") {
        const maTG = lookupValue(SHEET_NAMES.TRA_GOP, COL_TG.MA_DH, maDH, COL_TG.MA_TG);
        let oldTGStatus = "";
        const oldLSTGStates = [];

        if (maTG) {
          const tgRow = findRow(SHEET_NAMES.TRA_GOP, COL_TG.MA_TG, maTG);
          if (tgRow !== -1) {
            const tgSheet = ss.getSheetByName(SHEET_NAMES.TRA_GOP);
            oldTGStatus = tgSheet.getRange(tgRow, COL_TG.TRANG_THAI).getValue();

            const lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
            if (lstgSheet) {
              const lastRow = lstgSheet.getLastRow();
              if (lastRow > 1) {
                const lstgData = lstgSheet.getRange(2, 1, lastRow - 1, lstgSheet.getLastColumn()).getValues();
                for (let i = 0; i < lstgData.length; i++) {
                  if (String(lstgData[i][COL_LSTG.MA_TG - 1]) === maTG) {
                    oldLSTGStates.push({
                      row: i + 2,
                      status: String(lstgData[i][COL_LSTG.TRANG_THAI - 1])
                    });
                  }
                }
              }
            }
          }
        }

        huyHopDongTraGop(maDH);

        (function (tgId, tgStatus, lstgStates) {
          rollbackActions.push(function () {
            try {
              if (tgId) {
                const tgRow = findRow(SHEET_NAMES.TRA_GOP, COL_TG.MA_TG, tgId);
                if (tgRow !== -1) {
                  updateCell(SHEET_NAMES.TRA_GOP, tgRow, COL_TG.TRANG_THAI, tgStatus);
                }
                const lstgSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
                if (lstgSheet) {
                  lstgStates.forEach(function (item) {
                    lstgSheet.getRange(item.row, COL_LSTG.TRANG_THAI).setValue(item.status);
                  });
                  clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);
                }
              }
            } catch (err) {
              Logger.log("Rollback failed to restore installment contract status: " + err.message);
            }
          });
        })(maTG, oldTGStatus, oldLSTGStates);
      }

      showToast("✅ Đã huỷ đơn hàng: " + maDH + " và hoàn kho sản phẩm/quà tặng.");
      return true;
    });
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
  const donHangs = getDonHangTheoThang(thang, nam);
  const result = {
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
  const c = _getDonHangIndices();
  const row = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH);
  if (row === -1) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  const values = sheet.getRange(row, 1, 1, COL_DH.CHUYEN_KHOAN).getValues()[0];

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
