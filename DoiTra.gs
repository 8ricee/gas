/**
 * ============================================================
 * VanTran Mobile — DoiTra.gs
 * Nghiệp vụ khách hàng đổi máy & trả máy
 * ============================================================
 */

/**
 * Thực hiện giao dịch Đổi / Trả máy của khách
 *
 * @param {Object} data - {
 *   maDH, loaiGiaoDich ('Trả máy'/'Đổi máy'),
 *   maSP_Nhan, phiDoiTra, tienHoanTra,
 *   hinhThucThanhToan ('Tiền mặt'/'Chuyển khoản'),
 *   chiNhanh, nguoiThucHien (MaNV), ghiChu
 * }
 * @return {string} Mã đổi trả mới
 */
const DoiTraStrategy = {
  "Quà tặng kèm": {
    validate: function(data, context, ss) {
      const { dhRow, dhSheet, chiNhanh, maDH } = context;
      const maQuaCu = String(dhSheet.getRange(dhRow, COL_DH.MA_QUA_TANG).getValue() || "");
      const coNhanQua = dhSheet.getRange(dhRow, COL_DH.CO_NHAN_QUA).getValue();

      if (coNhanQua !== "✓" || !maQuaCu) {
        throw new Error(
          "Đơn hàng " + maDH + " không nhận quà tặng kèm trước đó!",
        );
      }

      const exchanges = data.giftExchanges || [];
      if (exchanges.length === 0) {
        const maQuaMoi = data.maSP_Nhan;
        if (!maQuaMoi) {
          throw new Error("Vui lòng chọn quà tặng mới để đổi!");
        }
        if (maQuaCu === maQuaMoi) {
          throw new Error("Quà tặng mới chọn trùng với quà tặng cũ đang nhận!");
        }

        const maQuaMoiList = String(maQuaMoi).split(",");
        const pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
        const codeCounts = {};
        for (let i = 0; i < maQuaMoiList.length; i++) {
          const code = maQuaMoiList[i].trim();
          if (!code) continue;
          codeCounts[code] = (codeCounts[code] || 0) + 1;
        }

        for (const code in codeCounts) {
          const newRow = findPhuKienRow(code, chiNhanh);
          if (newRow === -1) {
            throw new Error(
              "Quà tặng mới " + code + " không tồn tại ở chi nhánh " + chiNhanh,
            );
          }
          const tonQuaMoi =
            Number(pkSheet.getRange(newRow, COL_PK.SO_LUONG_TON).getValue()) ||
            0;
          const requiredQty = codeCounts[code];
          if (tonQuaMoi < requiredQty) {
            throw new Error(
              "Quà tặng mới " +
                code +
                " không đủ số lượng tại chi nhánh này! Hiện tại: " +
                tonQuaMoi +
                ", cần: " +
                requiredQty,
            );
          }
        }
        context.maQuaMoi = maQuaMoi;
        context.maQuaMoiList = maQuaMoiList;
      } else {
        const currentMaGifts = maQuaCu.split(",").map(function(c) { return c.trim().toLowerCase(); });
        const pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
        
        exchanges.forEach(function(item) {
          const oldCode = String(item.oldCode).trim();
          const newCode = String(item.newCode).trim();
          
          if (currentMaGifts.indexOf(oldCode.toLowerCase()) === -1) {
            throw new Error("Quà tặng cũ '" + oldCode + "' không thuộc đơn hàng " + maDH + "!");
          }
          
          const newRow = findPhuKienRow(newCode, chiNhanh);
          if (newRow === -1) {
            throw new Error("Quà tặng mới '" + newCode + "' không tồn tại ở chi nhánh " + chiNhanh);
          }
          
          const tonQuaMoi = Number(pkSheet.getRange(newRow, COL_PK.SO_LUONG_TON).getValue()) || 0;
          if (tonQuaMoi < 1) {
            throw new Error("Quà tặng mới '" + newCode + "' đã hết hàng tại chi nhánh " + chiNhanh + "!");
          }
        });
      }
      
      context.maQuaCu = maQuaCu;
    },
    execute: function(data, context, ss, rollbackActions) {
      const { dhRow, dhSheet, chiNhanh, maDH, maKH, tenKH, maDT, maQuaCu } = context;
      const originalBranch = dhSheet.getRange(dhRow, COL_DH.CHI_NHANH).getValue() || chiNhanh;
      
      const exchanges = data.giftExchanges || [];
      const pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
      
      let finalMaQua = maQuaCu;
      let finalTenQua = dhSheet.getRange(dhRow, COL_DH.TEN_QUA_TANG).getValue() || "";
      
      const returnedGiftsList = [];
      const issuedGiftsList = [];
      
      const trackingOldCodes = [];
      const trackingOldNames = [];
      const trackingNewCodes = [];
      const trackingNewNames = [];

      if (exchanges.length > 0) {
        exchanges.forEach(function(item) {
          const oldCode = String(item.oldCode).trim();
          const newCode = String(item.newCode).trim();
          
          const newRow = findPhuKienRow(newCode, chiNhanh);
          const newName = pkSheet.getRange(newRow, COL_PK.TEN_SP).getValue() || "";
          
          trackingOldCodes.push(oldCode);
          trackingOldNames.push(item.oldName || oldCode);
          trackingNewCodes.push(newCode);
          trackingNewNames.push(newName);
          
          if (data.congKho !== false) {
            updateTonKhoPhuKien(oldCode, 1, "nhap", originalBranch);
            returnedGiftsList.push(oldCode);
          }
          
          updateTonKhoPhuKien(newCode, 1, "xuat", chiNhanh);
          issuedGiftsList.push(newCode);
          
          let codesArr = finalMaQua.split(",").map(function(c) { return c.trim(); });
          let namesArr = finalTenQua.split(",").map(function(n) { return n.trim(); });
          
          let idx = -1;
          for (let k = 0; k < codesArr.length; k++) {
            if (codesArr[k].toLowerCase() === oldCode.toLowerCase()) {
              idx = k;
              break;
            }
          }
          if (idx !== -1) {
            codesArr[idx] = newCode;
            namesArr[idx] = newName;
          }
          finalMaQua = codesArr.join(", ");
          finalTenQua = namesArr.join(", ");
        });
        
        rollbackActions.push(function () {
          if (data.congKho !== false) {
            returnedGiftsList.forEach(function (code) {
              try {
                updateTonKhoPhuKien(code, 1, "xuat", originalBranch);
              } catch (err) {
                Logger.log("Rollback returned gift " + code + " failed: " + err.message);
              }
            });
          }
          issuedGiftsList.forEach(function (code) {
            try {
              updateTonKhoPhuKien(code, 1, "nhap", chiNhanh);
            } catch (err) {
              Logger.log("Rollback issued gift " + code + " failed: " + err.message);
            }
          });
        });
        
      } else {
        const maQuaMoi = context.maQuaMoi;
        const maQuaMoiList = context.maQuaMoiList;
        const tenQuaMoiList = [];
        
        for (let i = 0; i < maQuaMoiList.length; i++) {
          const code = maQuaMoiList[i].trim();
          if (!code) continue;
          const newRow = findPhuKienRow(code, chiNhanh);
          const giftName = pkSheet.getRange(newRow, COL_PK.TEN_SP).getValue() || "";
          tenQuaMoiList.push(giftName);
        }
        finalMaQua = maQuaMoi;
        finalTenQua = tenQuaMoiList.join(", ");

        trackingOldCodes.push(maQuaCu);
        trackingOldNames.push(dhSheet.getRange(dhRow, COL_DH.TEN_QUA_TANG).getValue() || "");
        trackingNewCodes.push(maQuaMoi);
        trackingNewNames.push(finalTenQua);

        if (data.congKho !== false) {
          const listCu = String(maQuaCu).split(",");
          for (let i = 0; i < listCu.length; i++) {
            const code = listCu[i].trim();
            if (code) {
              updateTonKhoPhuKien(code, 1, "nhap", originalBranch);
              returnedGiftsList.push(code);
            }
          }
          rollbackActions.push(function () {
            returnedGiftsList.forEach(function (code) {
              try {
                updateTonKhoPhuKien(code, 1, "xuat", originalBranch);
              } catch (err) {
                Logger.log("Rollback returned gift " + code + " failed: " + err.message);
              }
            });
          });
        }

        for (let i = 0; i < maQuaMoiList.length; i++) {
          const code = maQuaMoiList[i].trim();
          if (code) {
            updateTonKhoPhuKien(code, 1, "xuat", chiNhanh);
            issuedGiftsList.push(code);
          }
        }
        rollbackActions.push(function () {
          issuedGiftsList.forEach(function (code) {
            try {
              updateTonKhoPhuKien(code, 1, "nhap", chiNhanh);
            } catch (err) {
              Logger.log("Rollback issued gift " + code + " failed: " + err.message);
            }
          });
        });
      }

      const oldMaQua = dhSheet.getRange(dhRow, COL_DH.MA_QUA_TANG).getValue();
      const oldTenQua = dhSheet.getRange(dhRow, COL_DH.TEN_QUA_TANG).getValue();
      const oldNote = dhSheet.getRange(dhRow, COL_DH.GHI_CHU).getValue();

      updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.MA_QUA_TANG, finalMaQua);
      updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.TEN_QUA_TANG, finalTenQua);

      const dateStr = formatDate(new Date());
      const newNote =
        (oldNote || "") +
        " [Đổi quà ngày " +
        dateStr +
        ": " +
        trackingOldCodes.join(",") +
        " ➔ " +
        trackingNewCodes.join(",") +
        "]";
      updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.GHI_CHU, newNote);

      rollbackActions.push(function () {
        try {
          updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.MA_QUA_TANG, oldMaQua);
          updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.TEN_QUA_TANG, oldTenQua);
          updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.GHI_CHU, oldNote);
        } catch (err) {
          Logger.log("Rollback original order gift info failed: " + err.message);
        }
      });

      const rowData = [];
      rowData[COL_DT_TRA.MA_DT - 1] = maDT;
      rowData[COL_DT_TRA.NGAY_DT - 1] = new Date();
      rowData[COL_DT_TRA.MA_DH - 1] = maDH;
      rowData[COL_DT_TRA.MA_KH - 1] = maKH;
      rowData[COL_DT_TRA.TEN_KH - 1] = tenKH;
      rowData[COL_DT_TRA.LOAI_GD - 1] = "Đổi quà";
      rowData[COL_DT_TRA.MA_SP_TRA - 1] = trackingOldCodes.join(", ");
      rowData[COL_DT_TRA.TEN_SP_TRA - 1] = trackingOldNames.join(", ");
      rowData[COL_DT_TRA.IMEI_TRA - 1] = "";
      rowData[COL_DT_TRA.MA_SP_NHAN - 1] = trackingNewCodes.join(", ");
      rowData[COL_DT_TRA.TEN_SP_NHAN - 1] = trackingNewNames.join(", ");
      rowData[COL_DT_TRA.IMEI_NHAN - 1] = "";
      rowData[COL_DT_TRA.TIEN_HOAN_TRA - 1] = 0;
      rowData[COL_DT_TRA.PHI_DOI_TRA - 1] = Number(data.phiDoiTra) || 0;
      rowData[COL_DT_TRA.HINH_THUC_TT - 1] =
        data.hinhThucThanhToan || "Tiền mặt";
      rowData[COL_DT_TRA.CHI_NHANH - 1] = chiNhanh;
      rowData[COL_DT_TRA.NGUOI_THUC_HIEN - 1] = data.nguoiThucHien || "";
      rowData[COL_DT_TRA.TRANG_THAI - 1] = ORDER_STATUS.DONE;
      rowData[COL_DT_TRA.GHI_CHU - 1] = data.ghiChu || "";

      const totalPay = Number(data.phiDoiTra) || 0;
      const paymentSplit = calculatePaymentSplit(data, totalPay);
      rowData[COL_DT_TRA.TIEN_MAT - 1] = paymentSplit.tienMat;
      rowData[COL_DT_TRA.CHUYEN_KHOAN - 1] = paymentSplit.chuyenKhoan;

      const dtSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
      const dtRow = appendRow(SHEET_NAMES.DOI_TRA, rowData);
      (function (capturedMaDT) {
        rollbackActions.push(function () {
          try {
            const ssRollback = SpreadsheetApp.getActiveSpreadsheet();
            const sheet = ssRollback.getSheetByName(SHEET_NAMES.DOI_TRA);
            const r = findRow(SHEET_NAMES.DOI_TRA, COL_DT_TRA.MA_DT, capturedMaDT);
            if (r !== -1) {
              sheet.deleteRow(r);
              clearSheetCache(SHEET_NAMES.DOI_TRA);
            }
          } catch (err) {
            Logger.log("Rollback failed to delete exchange record " + capturedMaDT + ": " + err.message);
          }
        });
      })(maDT);

      showToast("✅ Đổi quà tặng thành công đơn " + maDH);
      return maDT;
    }
  },


  "Sản phẩm chính": {
    validate: function(data, context, ss) {
      const { dhRow, dhSheet, chiNhanh, loaiGD } = context;
      const maSP_Tra = dhSheet.getRange(dhRow, COL_DH.MA_SP).getValue();
      const tenSP_Tra = dhSheet.getRange(dhRow, COL_DH.TEN_SP).getValue();
      const nguonSP_Tra = dhSheet.getRange(dhRow, COL_DH.NGUON_SP).getValue();
      
      context.maSP_Tra = maSP_Tra;
      context.tenSP_Tra = tenSP_Tra;
      context.nguonSP_Tra = nguonSP_Tra;

      if (nguonSP_Tra === "Phụ kiện") {
        if (loaiGD === "Đổi hàng") {
          const maSP_Nhan = data.maSP_Nhan;
          if (!maSP_Nhan) {
            throw new Error("Vui lòng chọn phụ kiện mới để đổi!");
          }

          const pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
          const pkRow = findPhuKienRow(maSP_Nhan, chiNhanh);
          if (pkRow === -1) {
            throw new Error(
              "Phụ kiện mới " +
                maSP_Nhan +
                " không tồn tại ở chi nhánh " +
                chiNhanh,
            );
          }
          const tonMoi =
            Number(pkSheet.getRange(pkRow, COL_PK.SO_LUONG_TON).getValue()) ||
            0;
          if (tonMoi < 1) {
            throw new Error(
              "Phụ kiện mới " +
                maSP_Nhan +
                " đã hết hàng tại chi nhánh " +
                chiNhanh,
            );
          }
          
          context.maSP_Nhan = maSP_Nhan;
          context.tenSP_Nhan = pkSheet.getRange(pkRow, COL_PK.TEN_SP).getValue() || "";
          context.giaBanMoi = Number(pkSheet.getRange(pkRow, COL_PK.GIA_BAN).getValue()) || 0;
        }
      } else {
        // Điện thoại
        if (loaiGD === "Đổi máy") {
          const maSP_Nhan = data.maSP_Nhan;
          if (!maSP_Nhan) {
            throw new Error("Vui lòng chọn máy mới để đổi!");
          }
          context.maSP_Nhan = maSP_Nhan;
        }
      }
    },
    execute: function(data, context, ss, rollbackActions) {
      const { nguonSP_Tra } = context;
      if (nguonSP_Tra === "Phụ kiện") {
        return this._executePhuKien(data, context, ss, rollbackActions);
      } else {
        return this._executeDienThoai(data, context, ss, rollbackActions);
      }
    },
    
    _executePhuKien: function(data, context, ss, rollbackActions) {
      const { dhRow, dhSheet, chiNhanh, maDH, maKH, tenKH, maDT, loaiGD, maSP_Tra, tenSP_Tra, maSP_Nhan, tenSP_Nhan, giaBanMoi } = context;
      const soLuong = Number(dhSheet.getRange(dhRow, COL_DH.SO_LUONG).getValue()) || 1;

      // 1. Chuyển trạng thái đơn hàng gốc sang "Đổi trả"
      const oldDHStatus = dhSheet
        .getRange(dhRow, COL_DH.TRANG_THAI)
        .getValue();
      updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.TRANG_THAI, ORDER_STATUS.EXCHANGED);
      rollbackActions.push(function () {
        try {
          updateCell(
            SHEET_NAMES.DON_HANG,
            dhRow,
            COL_DH.TRANG_THAI,
            oldDHStatus,
          );
        } catch (err) {
          Logger.log("Rollback original order status failed: " + err.message);
        }
      });

      // 2. Hoàn trả kho phụ kiện cũ vào chi nhánh của đơn hàng gốc (nếu chọn cộng kho)
      if (data.congKho !== false) {
        const originalBranch = dhSheet
          .getRange(dhRow, COL_DH.CHI_NHANH)
          .getValue();
        updateTonKhoPhuKien(maSP_Tra, soLuong, "nhap", originalBranch);
        rollbackActions.push(function () {
          try {
            updateTonKhoPhuKien(maSP_Tra, soLuong, "xuat", originalBranch);
          } catch (err) {
            Logger.log("Rollback returned accessory stock failed: " + err.message);
          }
        });
      }

      // 3. Nếu là Đổi hàng -> Tạo đơn hàng mới cho phụ kiện mới nhận
      if (loaiGD === "Đổi hàng") {
        // Tạo đơn hàng mới cho phụ kiện nhận
        const maDHMoi = taoDonHang({
          maKH: maKH,
          maSP: maSP_Nhan,
          nguonSP: "Phụ kiện",
          soLuong: 1, // Đổi mặc định 1
          donGia: giaBanMoi,
          hinhThucBan: "Bán thẳng",
          hinhThucThanhToan: data.hinhThucThanhToan || "Tiền mặt",
          nguoiBan: data.nguoiThucHien,
          chiNhanh: chiNhanh,
          ghiChu: "Đơn hàng đổi phụ kiện, thay thế cho đơn gốc " + maDH,
        });

        rollbackActions.push(function () {
          try {
            huyDonHang(maDHMoi);
            const oSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
            const oRow = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDHMoi);
            if (oRow !== -1) {
              oSheet.deleteRow(oRow);
              clearSheetCache(SHEET_NAMES.DON_HANG);
            }
          } catch (err) {
            Logger.log("Rollback failed to delete new exchange order " + maDHMoi + ": " + err.message);
          }
        });

        data.ghiChu =
          (data.ghiChu || "") + " [Đổi sang đơn mới: " + maDHMoi + "]";
      }

      // 4. Ghi nhận giao dịch vào bảng DoiTra
      const tienHoanTra = Number(data.tienHoanTra) || 0;
      const phiDoiTra = Number(data.phiDoiTra) || 0;

      const rowData = [];
      rowData[COL_DT_TRA.MA_DT - 1] = maDT;
      rowData[COL_DT_TRA.NGAY_DT - 1] = new Date();
      rowData[COL_DT_TRA.MA_DH - 1] = maDH;
      rowData[COL_DT_TRA.MA_KH - 1] = maKH;
      rowData[COL_DT_TRA.TEN_KH - 1] = tenKH;
      rowData[COL_DT_TRA.LOAI_GD - 1] = loaiGD;
      rowData[COL_DT_TRA.MA_SP_TRA - 1] = maSP_Tra;
      rowData[COL_DT_TRA.TEN_SP_TRA - 1] = tenSP_Tra;
      rowData[COL_DT_TRA.IMEI_TRA - 1] = "";
      rowData[COL_DT_TRA.MA_SP_NHAN - 1] = maSP_Nhan || "";
      rowData[COL_DT_TRA.TEN_SP_NHAN - 1] = tenSP_Nhan || "";
      rowData[COL_DT_TRA.IMEI_NHAN - 1] = "";
      rowData[COL_DT_TRA.TIEN_HOAN_TRA - 1] = tienHoanTra;
      rowData[COL_DT_TRA.PHI_DOI_TRA - 1] = phiDoiTra;
      rowData[COL_DT_TRA.HINH_THUC_TT - 1] =
        data.hinhThucThanhToan || "Tiền mặt";
      rowData[COL_DT_TRA.CHI_NHANH - 1] = chiNhanh;
      rowData[COL_DT_TRA.NGUOI_THUC_HIEN - 1] = data.nguoiThucHien || "";
      rowData[COL_DT_TRA.TRANG_THAI - 1] = ORDER_STATUS.DONE;
      rowData[COL_DT_TRA.GHI_CHU - 1] = data.ghiChu || "";

      const totalPay = tienHoanTra > 0 ? tienHoanTra : phiDoiTra;
      const paymentSplit = calculatePaymentSplit(data, totalPay);
      rowData[COL_DT_TRA.TIEN_MAT - 1] = paymentSplit.tienMat;
      rowData[COL_DT_TRA.CHUYEN_KHOAN - 1] = paymentSplit.chuyenKhoan;

      const dtSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
      const dtRow = appendRow(SHEET_NAMES.DOI_TRA, rowData);
      (function (capturedMaDT) {
        rollbackActions.push(function () {
          try {
            const ssRollback = SpreadsheetApp.getActiveSpreadsheet();
            const sheet = ssRollback.getSheetByName(SHEET_NAMES.DOI_TRA);
            const r = findRow(SHEET_NAMES.DOI_TRA, COL_DT_TRA.MA_DT, capturedMaDT);
            if (r !== -1) {
              sheet.deleteRow(r);
              clearSheetCache(SHEET_NAMES.DOI_TRA);
            }
          } catch (err) {
            Logger.log("Rollback failed to delete exchange record " + capturedMaDT + ": " + err.message);
          }
        });
      })(maDT);

      showToast("✅ Ghi nhận Đổi trả phụ kiện thành công: " + maDT);
      return maDT;
    },

    _executeDienThoai: function(data, context, ss, rollbackActions) {
      const { dhRow, dhSheet, chiNhanh, maDH, maKH, tenKH, maDT, loaiGD, maSP_Tra, tenSP_Tra, maSP_Nhan } = context;
      
      const imei_Tra =
        dhSheet.getRange(dhRow, COL_DH.IMEI).getValue() ||
        lookupValue(
          SHEET_NAMES.DIEN_THOAI,
          COL_DT.MA_DT,
          maSP_Tra,
          COL_DT.IMEI,
        ) || "";

      // 1. Đổi trạng thái đơn hàng gốc sang "Đổi trả" để trừ doanh thu/hoa hồng
      const oldDHStatus = dhSheet
        .getRange(dhRow, COL_DH.TRANG_THAI)
        .getValue();
      updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.TRANG_THAI, ORDER_STATUS.EXCHANGED);
      rollbackActions.push(function () {
        try {
          updateCell(
            SHEET_NAMES.DON_HANG,
            dhRow,
            COL_DH.TRANG_THAI,
            oldDHStatus,
          );
        } catch (err) {
          Logger.log("Rollback original order status failed: " + err.message);
        }
      });

      // 2. Hoàn trả kho máy trả về trạng thái "Máy lỗi" (nếu chọn cộng kho)
      if (data.congKho !== false) {
        let oldPhoneStatus = lookupValue(
          SHEET_NAMES.DIEN_THOAI,
          COL_DT.IMEI,
          imei_Tra,
          COL_DT.TRANG_THAI_KHO,
        );
        if (oldPhoneStatus === null)
          oldPhoneStatus = lookupValue(
            SHEET_NAMES.DIEN_THOAI,
            COL_DT.MA_DT,
            maSP_Tra,
            COL_DT.TRANG_THAI_KHO,
          );
        updateTrangThaiKhoDT(imei_Tra || maSP_Tra, STOCK_STATUS.FAULTY);
        rollbackActions.push(function () {
          try {
            updateTrangThaiKhoDT(imei_Tra || maSP_Tra, oldPhoneStatus);
          } catch (err) {
            Logger.log("Rollback phone stock status failed: " + err.message);
          }
        });
      }

      // 3. Huỷ hợp đồng trả góp gốc nếu có
      const maTG = lookupValue(SHEET_NAMES.TRA_GOP, COL_TG.MA_DH, maDH, COL_TG.MA_TG);
      let oldTGStatus = "";
      const oldLSTGStates = [];
      if (maTG) {
        const tgRow = findRow(SHEET_NAMES.TRA_GOP, COL_TG.MA_TG, maTG);
        if (tgRow !== -1) {
          oldTGStatus = ss
            .getSheetByName(SHEET_NAMES.TRA_GOP)
            .getRange(tgRow, COL_TG.TRANG_THAI)
            .getValue();
          const lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
          const lastRow = lstgSheet.getLastRow();
          if (lastRow > 1) {
            const lstgData = lstgSheet
              .getRange(2, 1, lastRow - 1, lstgSheet.getLastColumn())
              .getValues();
            for (let i = 0; i < lstgData.length; i++) {
              if (String(lstgData[i][COL_LSTG.MA_TG - 1]) === maTG) {
                oldLSTGStates.push({
                  row: i + 2,
                  status: String(lstgData[i][COL_LSTG.TRANG_THAI - 1]),
                });
              }
            }
          }
        }
      }

      huyHopDongTraGop(maDH);

      rollbackActions.push(function () {
        try {
          if (maTG) {
            const tgRow = findRow(SHEET_NAMES.TRA_GOP, COL_TG.MA_TG, maTG);
            if (tgRow !== -1) {
              updateCell(SHEET_NAMES.TRA_GOP, tgRow, COL_TG.TRANG_THAI, oldTGStatus);
            }
            const lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
            oldLSTGStates.forEach(function (item) {
              lstgSheet.getRange(item.row, COL_LSTG.TRANG_THAI).setValue(item.status);
            });
            clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);
          }
        } catch (err) {
          Logger.log("Rollback installment contract status failed for TG " + maTG + ": " + err.message);
        }
      });

      let tenSP_Nhan = "";
      let imei_Nhan = "";

      // 4. Nếu là Đổi máy -> Tạo đơn hàng mới cho máy nhận
      if (loaiGD === "Đổi máy") {
        tenSP_Nhan =
          lookupValue(
            SHEET_NAMES.DIEN_THOAI,
            COL_DT.MA_DT,
            maSP_Nhan,
            COL_DT.TEN_SP,
          ) || "";

        // Tự động tìm máy có mã SP_Nhan tại chi nhánh này và đang "Còn hàng"
        const dtData = getAllData(SHEET_NAMES.DIEN_THOAI);
        const maDTIdx = COL_DT.MA_DT - 1;
        const chiNhanhIdx = COL_DT.CHI_NHANH - 1;
        const trangThaiKhoIdx = COL_DT.TRANG_THAI_KHO - 1;
        const imeiIdx = COL_DT.IMEI - 1;

        for (let i = 0; i < dtData.length; i++) {
          if (
            String(dtData[i][maDTIdx]) === maSP_Nhan &&
            String(dtData[i][chiNhanhIdx]) === chiNhanh &&
            String(dtData[i][trangThaiKhoIdx]) === STOCK_STATUS.IN_STOCK
          ) {
            imei_Nhan = String(dtData[i][imeiIdx]);
            break;
          }
        }
        
        if (!imei_Nhan) {
          imei_Nhan =
            lookupValue(
              SHEET_NAMES.DIEN_THOAI,
              COL_DT.MA_DT,
              maSP_Nhan,
              COL_DT.IMEI,
            ) || "";
        }

        const giaBanMoi =
          Number(
            lookupValue(
              SHEET_NAMES.DIEN_THOAI,
              COL_DT.MA_DT,
              maSP_Nhan,
              COL_DT.GIA_BAN,
            ),
          ) || 0;

        // Tạo đơn hàng mới cho máy nhận
        const maDHMoi = taoDonHang({
          maKH: maKH,
          maSP: maSP_Nhan,
          imei: imei_Nhan,
          nguonSP: "Điện thoại",
          soLuong: 1,
          donGia: giaBanMoi,
          hinhThucBan: "Bán thẳng",
          hinhThucThanhToan: data.hinhThucThanhToan || "Tiền mặt",
          nguoiBan: data.nguoiThucHien,
          chiNhanh: chiNhanh,
          ghiChu: "Đơn hàng đổi máy, thay thế cho đơn gốc " + maDH,
        });

        rollbackActions.push(function () {
          try {
            huyDonHang(maDHMoi);
            const oSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
            const oRow = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDHMoi);
            if (oRow !== -1) {
              oSheet.deleteRow(oRow);
              clearSheetCache(SHEET_NAMES.DON_HANG);
            }
          } catch (err) {
            Logger.log("Rollback failed to delete new exchange order " + maDHMoi + ": " + err.message);
          }
        });

        data.ghiChu =
          (data.ghiChu || "") + " [Đổi sang đơn mới: " + maDHMoi + "]";
      }

      // 5. Ghi nhận giao dịch vào bảng DoiTra
      const tienHoanTra = Number(data.tienHoanTra) || 0;
      const phiDoiTra = Number(data.phiDoiTra) || 0;

      const rowData = [];
      rowData[COL_DT_TRA.MA_DT - 1] = maDT;
      rowData[COL_DT_TRA.NGAY_DT - 1] = new Date();
      rowData[COL_DT_TRA.MA_DH - 1] = maDH;
      rowData[COL_DT_TRA.MA_KH - 1] = maKH;
      rowData[COL_DT_TRA.TEN_KH - 1] = tenKH;
      rowData[COL_DT_TRA.LOAI_GD - 1] = loaiGD;
      rowData[COL_DT_TRA.MA_SP_TRA - 1] = maSP_Tra;
      rowData[COL_DT_TRA.TEN_SP_TRA - 1] = tenSP_Tra;
      rowData[COL_DT_TRA.IMEI_TRA - 1] = imei_Tra;
      rowData[COL_DT_TRA.MA_SP_NHAN - 1] = maSP_Nhan || "";
      rowData[COL_DT_TRA.TEN_SP_NHAN - 1] = tenSP_Nhan || "";
      rowData[COL_DT_TRA.IMEI_NHAN - 1] = imei_Nhan || "";
      rowData[COL_DT_TRA.TIEN_HOAN_TRA - 1] = tienHoanTra;
      rowData[COL_DT_TRA.PHI_DOI_TRA - 1] = phiDoiTra;
      rowData[COL_DT_TRA.HINH_THUC_TT - 1] =
        data.hinhThucThanhToan || "Tiền mặt";
      rowData[COL_DT_TRA.CHI_NHANH - 1] = chiNhanh;
      rowData[COL_DT_TRA.NGUOI_THUC_HIEN - 1] = data.nguoiThucHien || "";
      rowData[COL_DT_TRA.TRANG_THAI - 1] = ORDER_STATUS.DONE;
      rowData[COL_DT_TRA.GHI_CHU - 1] = data.ghiChu || "";

      const totalPay = tienHoanTra > 0 ? tienHoanTra : phiDoiTra;
      const paymentSplit = calculatePaymentSplit(data, totalPay);
      rowData[COL_DT_TRA.TIEN_MAT - 1] = paymentSplit.tienMat;
      rowData[COL_DT_TRA.CHUYEN_KHOAN - 1] = paymentSplit.chuyenKhoan;

      const dtSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
      const dtRow = appendRow(SHEET_NAMES.DOI_TRA, rowData);
      (function (capturedMaDT) {
        rollbackActions.push(function () {
          try {
            const ssRollback = SpreadsheetApp.getActiveSpreadsheet();
            const sheet = ssRollback.getSheetByName(SHEET_NAMES.DOI_TRA);
            const r = findRow(SHEET_NAMES.DOI_TRA, COL_DT_TRA.MA_DT, capturedMaDT);
            if (r !== -1) {
              sheet.deleteRow(r);
              clearSheetCache(SHEET_NAMES.DOI_TRA);
            }
          } catch (err) {
            Logger.log("Rollback failed to delete exchange record " + capturedMaDT + ": " + err.message);
          }
        });
      })(maDT);

      showToast("✅ Ghi nhận Đổi trả thành công: " + maDT);
      return maDT;
    }
  }
};
function thucHienDoiTra(data) {
  return withDocumentLock(function () {
    // Backend validation for Hỗn hợp payments
    let expectedPaid = 0;
    const isHoanTra = (Number(data.tienHoanTra) || 0) > 0;
    if (isHoanTra) {
      expectedPaid = Number(data.tienHoanTra) || 0;
    } else {
      expectedPaid = Number(data.phiDoiTra) || 0;
    }

    assertPaymentMatches(data, expectedPaid);

    const maDT = generateId("DT", SHEET_NAMES.DOI_TRA);
    const maDH = data.maDH;
    const loaiGD = data.loaiGiaoDich;
    const chiNhanh = data.chiNhanh;
    const doiTuong = data.doiTuong || "Sản phẩm chính";

    if (!maDH || !loaiGD || !chiNhanh) {
      throw new Error(
        "Vui lòng điền đầy đủ Mã đơn hàng, Loại giao dịch và Chi nhánh!",
      );
    }

    // 1. Tìm đơn hàng gốc
    const dhRow = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH);
    if (dhRow === -1) {
      throw new Error("Không tìm thấy đơn hàng gốc: " + maDH);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dhSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);

    const dhStatus = dhSheet.getRange(dhRow, COL_DH.TRANG_THAI).getValue();
    if (isCancelStatus(dhStatus) || dhStatus === ORDER_STATUS.EXCHANGED) {
      throw new Error(
        "Đơn hàng " +
          maDH +
          " đã ở trạng thái: " +
          dhStatus +
          ", không thể đổi trả!",
      );
    }

    const maKH = dhSheet.getRange(dhRow, COL_DH.MA_KH).getValue();
    const tenKH = dhSheet.getRange(dhRow, COL_DH.TEN_KH).getValue();

    // Prepare strategy context
    const context = {
      maDT: maDT,
      maDH: maDH,
      loaiGD: loaiGD,
      chiNhanh: chiNhanh,
      dhRow: dhRow,
      dhSheet: dhSheet,
      maKH: maKH,
      tenKH: tenKH
    };

    const strategy = DoiTraStrategy[doiTuong];
    if (!strategy) {
      throw new Error("Đối tượng đổi trả không hợp lệ: " + doiTuong);
    }

    // Validate strategy requirements before transaction
    strategy.validate(data, context, ss);

    // Execute within transaction
    return _executeWithRollback(function (rollbackActions) {
      return strategy.execute(data, context, ss, rollbackActions);
    });
  });
}
