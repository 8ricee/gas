/**
 * ============================================================
 * VanTran Mobile — ProductStrategy.gs
 * Strategy pattern for handling product-specific rules and actions
 * ============================================================
 */

const ProductStrategy = {
  "Điện thoại": {
    lookup: function(maSP) {
      return {
        tenSP: lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.TEN_SP) || "",
        thuongHieu: lookupValue(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP, COL_DT.THUONG_HIEU) || ""
      };
    },
    checkStock: function(data, chiNhanh, soLuong, ss) {
      const dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
      let phoneRow = -1;

      if (data.imei) {
        phoneRow = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.IMEI, data.imei);
        if (phoneRow === -1 && COL_DT.IMEI_2) {
          phoneRow = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.IMEI_2, data.imei);
        }
      } else {
        // Tự động tìm máy có mã SP tại chi nhánh này và đang "Còn hàng"
        const dtData = getAllData(SHEET_NAMES.DIEN_THOAI);
        const maDTIdx = COL_DT.MA_DT - 1;
        const chiNhanhIdx = COL_DT.CHI_NHANH - 1;
        const trangThaiKhoIdx = COL_DT.TRANG_THAI_KHO - 1;
        const imeiIdx = COL_DT.IMEI - 1;

        for (let i = 0; i < dtData.length; i++) {
          if (
            String(dtData[i][maDTIdx]) === data.maSP &&
            String(dtData[i][chiNhanhIdx]) === chiNhanh &&
            String(dtData[i][trangThaiKhoIdx]) === STOCK_STATUS.IN_STOCK
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

      const trangThaiKho = dtSheet.getRange(phoneRow, COL_DT.TRANG_THAI_KHO).getValue();
      if (trangThaiKho !== STOCK_STATUS.IN_STOCK) {
        throw new Error(
          "Điện thoại " +
            data.maSP +
            " không còn hàng (Trạng thái: " +
            trangThaiKho +
            ")",
        );
      }
      
      const dtBranch = dtSheet.getRange(phoneRow, COL_DT.CHI_NHANH).getValue();
      if (dtBranch !== chiNhanh) {
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
        data.hinhThucBan === "Trả góp" &&
        data.traGop &&
        data.traGop.loaiTraGop === "Cửa hàng"
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
  "Phụ kiện": {
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
      const tonKho = Number(pkSheet.getRange(pkRow, COL_PK.SO_LUONG_TON).getValue()) || 0;
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
