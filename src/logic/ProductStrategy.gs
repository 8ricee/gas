/**
 * ============================================================
 * VanTran Mobile — ProductStrategy.gs
 * Strategy pattern for handling product-specific rules and actions
 * ============================================================
 */

const ProductStrategy = {
  [PRODUCT_SOURCE.PHONE]: {
    lookup: function(maSP) {
      return {
        tenSP: lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.TEN_SP) || "",
        thuongHieu: lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.THUONG_HIEU) || ""
      };
    },
    checkStock: function(data, chiNhanh, soLuong, ss) {
      const normalizedQty = Number(soLuong) || 1;
      if (normalizedQty !== 1) {
        throw new Error("Điện thoại chỉ được bán với số lượng 1 cho mỗi IMEI/đơn hàng.");
      }

      const resolved = resolvePhoneRowForSale(data, chiNhanh);
      const phoneRow = resolved.row;
      if (phoneRow === -1) {
        throw new Error(
          "Không tìm thấy điện thoại " +
            (data.imei ? "IMEI: " + data.imei : data.maSP) +
            " còn hàng tại chi nhánh " + chiNhanh + "!"
        );
      }

      if (resolved.resolvedImei) {
        data.imei = resolved.resolvedImei;
      }

      const trangThaiKho = getCachedCellValue(SHEET_NAMES.DIEN_THOAI, phoneRow, COL_DT.TRANG_THAI_KHO);
      if (trangThaiKho !== STOCK_STATUS.IN_STOCK) {
        throw new Error(
          "Điện thoại " +
            data.maSP +
            " không còn hàng (Trạng thái: " +
            trangThaiKho +
            ")",
        );
      }

      const dtBranch = getCachedCellValue(SHEET_NAMES.DIEN_THOAI, phoneRow, COL_DT.CHI_NHANH);
      if (String(dtBranch).trim() !== String(chiNhanh).trim()) {
        throw new Error(
          "Điện thoại " +
            data.maSP +
            " không thuộc chi nhánh này (Thuộc: " +
            dtBranch +
            ")",
        );
      }
    },
    updateStock: function(data, chiNhanh, soLuong, rollbackActions, ss) {
      let trangThaiMoi = STOCK_STATUS.SOLD;
      if (
        data.hinhThucBan === SALES_METHOD.INSTALLMENT &&
        data.traGop &&
        data.traGop.loaiTraGop === INSTALLMENT_TYPE.STORE
      ) {
        trangThaiMoi = STOCK_STATUS.INSTALLMENT;
      }
      updateTrangThaiKhoDT(data.imei || data.maSP, trangThaiMoi);
      const key = data.imei || data.maSP;
      addRollback(rollbackActions, "Restore phone status to IN_STOCK", function () {
        updateTrangThaiKhoDT(key, STOCK_STATUS.IN_STOCK);
      });
    },
    restoreStock: function(data, chiNhanh, soLuong) {
      updateTrangThaiKhoDT(data.imei || data.maSP, STOCK_STATUS.IN_STOCK);
    }
  },
  [PRODUCT_SOURCE.ACCESSORY]: {
    lookup: function(maSP) {
      return {
        tenSP: lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.TEN_SP) || "",
        thuongHieu: lookupValue(SHEET_NAMES.PHU_KIEN, COL_PK.MA_PK, maSP, COL_PK.THUONG_HIEU) || ""
      };
    },
    checkStock: function(data, chiNhanh, soLuong, ss) {
      const pkRow = findPhuKienRow(data.maSP, chiNhanh);
      if (pkRow === -1) {
        throw new Error(
          "Phụ kiện " + data.maSP + " không tồn tại ở chi nhánh " + chiNhanh,
        );
      }
      const pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
      const tonKho = Number(getCachedCellValue(SHEET_NAMES.PHU_KIEN, pkRow, COL_PK.SO_LUONG_TON)) || 0;
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
    },
    updateStock: function(data, chiNhanh, soLuong, rollbackActions, ss) {
      updateTonKhoPhuKien(data.maSP, soLuong, "xuat", chiNhanh);
      const maSP = data.maSP;
      const qty = soLuong;
      const branch = chiNhanh;
      addRollback(rollbackActions, "Restore accessory stock", function () {
        updateTonKhoPhuKien(maSP, qty, "nhap", branch);
      });
    },
    restoreStock: function(data, chiNhanh, soLuong) {
      updateTonKhoPhuKien(data.maSP, soLuong, "nhap", chiNhanh);
    }
  }
};
