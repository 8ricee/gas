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
  return withDocumentLock(function () {
    clearSheetCache();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Validation thanh toán hỗn hợp ở Backend
    const netPayable = _calcNetPayable(data);
    const deduction = Number(data.tradeInDeduction) || 0;
    const netToPay = Math.max(0, netPayable - deduction);

    let expectedPaid = netToPay;
    if (data.hinhThucBan === SALES_METHOD.INSTALLMENT && data.traGop) {
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

    assertPaymentMatches(data, expectedPaid);

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
          const itemThanhTien = calcThanhTien(item.soLuong, item.donGia, itemDiscount, 0);

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
            hinhThucBan: SALES_METHOD.DIRECT,
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
    return Math.max(0, calcThanhTien(1, total, data.tienGiamGia, 0));
  }
  return Math.max(
    0,
    calcThanhTien(data.soLuong || 1, data.donGia || 0, data.tienGiamGia || 0, 0),
  );
}



/**
 * Thực thi logic tạo một đơn hàng đơn lẻ và ghi nhận các hành động rollback
 * @private
 */
function _taoDonHangSingle(data, rollbackActions, ss) {
  const maDH = generateId("DH", SHEET_NAMES.DON_HANG);
  validateRequiredFields(data, [
    { key: "chiNhanh", label: "Chi nhánh" }
  ]);
  const chiNhanh = data.chiNhanh;

  // Lookup thông tin (Tự động thêm KH nếu chưa có)
  const tenKH = ensureKhachHangExists(data.maKH, data.tenKH);
  const nguonSP = data.nguonSP || PRODUCT_SOURCE.PHONE;
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
  const deduction = Number(data.tradeInDeduction) || 0;
  const thanhTien = calcThanhTien(soLuong, donGia, tienGiamGia, deduction);

  if (nguonSP === PRODUCT_SOURCE.ACCESSORY && data.hinhThucBan === SALES_METHOD.INSTALLMENT) {
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

  const paidAmount = (data.hinhThucBan === SALES_METHOD.INSTALLMENT && data.traGop) ? (Number(data.traGop.traTruoc) || 0) : thanhTien;
  const splitResult = calculatePaymentSplit(data, paidAmount);
  const tienMat = splitResult.tienMat;
  const chuyenKhoan = splitResult.chuyenKhoan;

  const rowData = buildRowData(COL_DH, {
    MA_DH: maDH,
    NGAY_BAN: new Date(),
    MA_KH: data.maKH || "",
    TEN_KH: tenKH,
    MA_SP: data.maSP || "",
    TEN_SP: tenSP,
    NGUON_SP: nguonSP,
    THUONG_HIEU: thuongHieu,
    SO_LUONG: soLuong,
    DON_GIA: donGia,
    THANH_TIEN: thanhTien,
    HINH_THUC_BAN: data.hinhThucBan || SALES_METHOD.DIRECT,
    HINH_THUC_TT: data.hinhThucThanhToan || PAYMENT_METHOD.CASH,
    NGUOI_BAN: data.nguoiBan || "",
    TEN_NGUOI_BAN: tenNguoiBan,
    QUYEN_XUAT: coQuyenXuatMay,
    NGUOI_HO_TRO: data.nguoiHoTro || "",
    TEN_NGUOI_HO_TRO: tenNguoiHoTro,
    TRANG_THAI: ORDER_STATUS.DONE,
    GHI_CHU: data.ghiChu || "",
    CHI_NHANH: chiNhanh,
    MA_QUA_TANG: maQuaTang,
    TEN_QUA_TANG: tenQuaTang,
    CO_NHAN_QUA: coNhanQua,
    TIEN_GIAM_GIA: tienGiamGia,
    IMEI: data.imei || "",
    TIEN_THU_MUA: deduction,
    TIEN_MAT: tienMat,
    CHUYEN_KHOAN: chuyenKhoan
  });

  const donHangSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
  const donHangRow = appendRow(SHEET_NAMES.DON_HANG, rowData);

  addRollback(rollbackActions, "Delete order record " + maDH, function () {
    const ssRollback = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ssRollback.getSheetByName(SHEET_NAMES.DON_HANG);
    const r = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH);
    if (r !== -1) {
      sheet.deleteRow(r);
      clearSheetCache(SHEET_NAMES.DON_HANG);
    }
  });

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
    addRollback(rollbackActions, "Restore gift stock for order " + maDH, function () {
      for (let j = 0; j < updatedGifts.length; j++) {
        updateTonKhoPhuKien(updatedGifts[j], 1, "nhap", chiNhanh);
      }
    });
  }

  // Nếu trả góp → tạo hợp đồng
  if (data.hinhThucBan === SALES_METHOD.INSTALLMENT && data.traGop) {
    const maTG = taoHopDongTraGop({
      maDH: maDH,
      maKH: data.maKH,
      tongTien:
        data.traGop.tongTien !== undefined
          ? Number(data.traGop.tongTien)
          : thanhTien,
      traTruoc: data.traGop.traTruoc,
      soKy: Number(data.traGop.soKy) || 1,
      loaiTraGop: data.traGop.loaiTraGop || INSTALLMENT_TYPE.STORE,
      congTyTC: data.traGop.congTyTC || "",
      tienMoiKy: Number(data.traGop.tienMoiKy) || 0,
      chiNhanh: chiNhanh,
      hinhThucThanhToan: data.hinhThucThanhToan,
      splitTienMat: data.splitTienMat,
      splitChuyenKhoan: data.splitChuyenKhoan,
    });
    addRollback(rollbackActions, "Delete installment contract " + maTG, function () {
      const ssRollback = SpreadsheetApp.getActiveSpreadsheet();
      const tgSheet = ssRollback.getSheetByName(SHEET_NAMES.TRA_GOP);
      const tgRow = findRow(SHEET_NAMES.TRA_GOP, COL_TG.MA_TG, maTG);
      if (tgRow !== -1) {
        tgSheet.deleteRow(tgRow);
        clearSheetCache(SHEET_NAMES.TRA_GOP);
        invalidateDropdownCache(SHEET_NAMES.TRA_GOP);
      }
      const lstgSheet = ssRollback.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
      const lstgData = lstgSheet.getDataRange().getValues();
      for (let rowIdx = lstgData.length - 1; rowIdx >= 1; rowIdx--) {
        if (String(lstgData[rowIdx][COL_LSTG.MA_TG - 1]) === maTG) {
          lstgSheet.deleteRow(rowIdx + 1);
        }
      }
      clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);
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
  const data = getAllData(SHEET_NAMES.DON_HANG);
  const result = [];

  data.forEach(function (row) {
    const obj = mapRowToObject(row, SHEET_NAMES.DON_HANG);
    const ngayBan = obj.NGAY_BAN;
    if (isSameMonthYear(ngayBan, thang, nam)) {
      result.push({
        MaDH: String(obj.MA_DH),
        NgayBan: obj.NGAY_BAN,
        MaKH: String(obj.MA_KH),
        TenKH: String(obj.TEN_KH),
        MaSP: String(obj.MA_SP),
        TenSP: String(obj.TEN_SP),
        NguonSP: String(obj.NGUON_SP),
        ThuongHieu: String(obj.THUONG_HIEU),
        SoLuong: Number(obj.SO_LUONG) || 0,
        DonGia: Number(obj.DON_GIA) || 0,
        ThanhTien: Number(obj.THANH_TIEN) || 0,
        HinhThucBan: String(obj.HINH_THUC_BAN),
        HinhThucThanhToan: String(obj.HINH_THUC_TT),
        NguoiBan: String(obj.NGUOI_BAN),
        TenNguoiBan: String(obj.TEN_NGUOI_BAN),
        CoQuyenXuatMay: String(obj.QUYEN_XUAT),
        NguoiHoTro: String(obj.NGUOI_HO_TRO),
        TenNguoiHoTro: String(obj.TEN_NGUOI_HO_TRO),
        TrangThai: String(obj.TRANG_THAI),
        GhiChu: String(obj.GHI_CHU),
        ChiNhanh: String(obj.CHI_NHANH || ""),
        MaQuaTang: String(obj.MA_QUA_TANG || ""),
        TenQuaTang: String(obj.TEN_QUA_TANG || ""),
        CoNhanQua: String(obj.CO_NHAN_QUA || "✗"),
        TienGiamGia: Number(obj.TIEN_GIAM_GIA || 0),
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

    if (isCancelStatus(currentTT)) {
      showAlert("⚠️ Cảnh báo", "Đơn hàng " + maDH + " đã bị huỷ trước đó.");
      return false;
    }

    return _executeWithRollback(function (rollbackActions) {
      // Hoàn kho sản phẩm chính
      const nguonSP = rowValues[COL_DH.NGUON_SP - 1];
      const maSP = rowValues[COL_DH.MA_SP - 1];
      const soLuong = Number(rowValues[COL_DH.SO_LUONG - 1]);
      const chiNhanh = rowValues[COL_DH.CHI_NHANH - 1];

      const strategy = ProductStrategy[nguonSP];
      if (!strategy) {
        throw new Error("Loại sản phẩm không hợp lệ: " + nguonSP);
      }

      const prodData = {
        maSP: maSP,
        imei: rowValues[COL_DH.IMEI - 1] || ""
      };

      let oldPhoneStatus = null;
      if (nguonSP === PRODUCT_SOURCE.PHONE) {
        const targetPhoneKey = prodData.imei || prodData.maSP;
        oldPhoneStatus = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.IMEI, targetPhoneKey, COL_DT.TRANG_THAI_KHO);
        if (oldPhoneStatus === null) {
          oldPhoneStatus = lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, targetPhoneKey, COL_DT.TRANG_THAI_KHO);
        }
      }

      strategy.restoreStock(prodData, chiNhanh, soLuong);

      if (nguonSP === PRODUCT_SOURCE.PHONE) {
        const key = prodData.imei || prodData.maSP;
        const status = oldPhoneStatus;
        addRollback(rollbackActions, "Restore phone status", function () {
          if (status !== null) {
            updateTrangThaiKhoDT(key, status);
          }
        });
      } else {
        const code = maSP;
        const qty = soLuong;
        const branch = chiNhanh;
        addRollback(rollbackActions, "Restore accessory stock", function () {
          updateTonKhoPhuKien(code, qty, "xuat", branch);
        });
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
        const gifts = returnedGifts;
        const branch = chiNhanh;
        addRollback(rollbackActions, "Restore gift stock", function () {
          for (let i = 0; i < gifts.length; i++) {
            updateTonKhoPhuKien(gifts[i], 1, "xuat", branch);
          }
        });
      }

      // Đánh dấu huỷ
      updateCell(SHEET_NAMES.DON_HANG, row, COL_DH.TRANG_THAI, ORDER_STATUS.CANCELLED);

      const rowNum = row;
      const oldStatus = currentTT;
      addRollback(rollbackActions, "Restore order status", function () {
        updateCell(SHEET_NAMES.DON_HANG, rowNum, COL_DH.TRANG_THAI, oldStatus);
      });

      // Kiểm tra nếu có trả góp → cập nhật trạng thái hợp đồng và các kỳ liên quan
      const hinhThuc = rowValues[COL_DH.HINH_THUC_BAN - 1];
      if (hinhThuc === SALES_METHOD.INSTALLMENT) {
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

        const tgId = maTG;
        const tgStatus = oldTGStatus;
        const lstgStates = oldLSTGStates;
        addRollback(rollbackActions, "Restore installment contract status", function () {
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
        });
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
    if (!isCancelStatus(dh.TrangThai) && dh.TrangThai !== ORDER_STATUS.EXCHANGED) {
      result.tongDonHang++;
      result.tongDoanhThu += Number(dh.ThanhTien) || 0;
      if (dh.NguonSP === PRODUCT_SOURCE.PHONE) {
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
  const row = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH);
  if (row === -1) return null;

  const values = getRowValues(SHEET_NAMES.DON_HANG, row);
  const dto = toDTO(values, SHEET_NAMES.DON_HANG, "DON_HANG_FULL");
  if (dto.ngayBan) {
    dto.ngayBan = formatDate(dto.ngayBan);
  }
  return dto;
}
