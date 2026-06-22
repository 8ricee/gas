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
function thucHienDoiTra(data) {
  var maDT = generateId("DT", SHEET_NAMES.DOI_TRA);
  var maDH = data.maDH;
  var loaiGD = data.loaiGiaoDich;
  var chiNhanh = data.chiNhanh;
  var doiTuong = data.doiTuong || "Sản phẩm chính";

  if (!maDH || !loaiGD || !chiNhanh) {
    throw new Error(
      "Vui lòng điền đầy đủ Mã đơn hàng, Loại giao dịch và Chi nhánh!",
    );
  }

  // 1. Tìm đơn hàng gốc
  var dhRow = findRow(SHEET_NAMES.DON_HANG, 1, maDH);
  if (dhRow === -1) {
    throw new Error("Không tìm thấy đơn hàng gốc: " + maDH);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dhSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);

  var dhStatus = dhSheet.getRange(dhRow, 19).getValue();
  if (dhStatus === "Huỷ" || dhStatus === "Đổi trả") {
    throw new Error(
      "Đơn hàng " +
        maDH +
        " đã ở trạng thái: " +
        dhStatus +
        ", không thể đổi trả!",
    );
  }

  var maKH = dhSheet.getRange(dhRow, 3).getValue();
  var tenKH = dhSheet.getRange(dhRow, 4).getValue();

  // XỬ LÝ THEO ĐỐI TƯỢNG ĐỔI TRẢ
  var rollbackActions = [];
  try {
    if (doiTuong === "Quà tặng kèm") {
      var maQuaCu = dhSheet.getRange(dhRow, 22).getValue();
      var tenQuaCu = dhSheet.getRange(dhRow, 23).getValue() || "(Không tên)";
      var coNhanQua = dhSheet.getRange(dhRow, 24).getValue();

      if (coNhanQua !== "✓") {
        throw new Error("Đơn hàng " + maDH + " không nhận quà tặng kèm trước đó!");
      }

      var maQuaMoi = data.maSP_Nhan;
      if (!maQuaMoi) {
        throw new Error("Vui lòng chọn quà tặng mới để đổi!");
      }

      if (maQuaCu === maQuaMoi) {
        throw new Error("Quà tặng mới chọn trùng với quà tặng cũ đang nhận!");
      }

      // Kiểm tra tồn kho quà mới
      var maQuaMoiList = String(maQuaMoi).split(",");
      var tenQuaMoiList = [];
      var pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
      
      var codeCounts = {};
      for (var i = 0; i < maQuaMoiList.length; i++) {
        var code = maQuaMoiList[i].trim();
        if (!code) continue;
        codeCounts[code] = (codeCounts[code] || 0) + 1;
      }

      for (var code in codeCounts) {
        var newRow = findPhuKienRow(code, chiNhanh);
        if (newRow === -1) {
          throw new Error(
            "Quà tặng mới " + code + " không tồn tại ở chi nhánh " + chiNhanh,
          );
        }
        var tonQuaMoi = Number(pkSheet.getRange(newRow, 7).getValue()) || 0;
        var requiredQty = codeCounts[code];
        if (tonQuaMoi < requiredQty) {
          throw new Error(
            "Quà tặng mới " + code + " không đủ số lượng tại chi nhánh này! Hiện tại: " + tonQuaMoi + ", cần: " + requiredQty,
          );
        }
        var giftName = pkSheet.getRange(newRow, 2).getValue() || "";
        for (var k = 0; k < requiredQty; k++) {
          tenQuaMoiList.push(giftName);
        }
      }
      var tenQuaMoi = tenQuaMoiList.join(", ");

      // Hoàn kho quà cũ tại chi nhánh
      var returnedGifts = [];
      if (maQuaCu) {
        var listCu = String(maQuaCu).split(",");
        for (var i = 0; i < listCu.length; i++) {
          var code = listCu[i].trim();
          if (code) {
            updateTonKhoPhuKien(code, 1, "nhap", chiNhanh);
            returnedGifts.push(code);
          }
        }
        rollbackActions.push(function() {
          returnedGifts.forEach(function(code) {
            try { updateTonKhoPhuKien(code, 1, "xuat", chiNhanh); } catch(err){}
          });
        });
      }

      // Trừ kho quà mới tại chi nhánh
      var issuedGifts = [];
      for (var i = 0; i < maQuaMoiList.length; i++) {
        var code = maQuaMoiList[i].trim();
        if (code) {
          updateTonKhoPhuKien(code, 1, "xuat", chiNhanh);
          issuedGifts.push(code);
        }
      }
      rollbackActions.push(function() {
        issuedGifts.forEach(function(code) {
          try { updateTonKhoPhuKien(code, 1, "nhap", chiNhanh); } catch(err){}
        });
      });

      // Cập nhật thông tin đơn hàng gốc
      var oldMaQua = dhSheet.getRange(dhRow, 22).getValue();
      var oldTenQua = dhSheet.getRange(dhRow, 23).getValue();
      var oldNote = dhSheet.getRange(dhRow, 20).getValue();

      updateCell(SHEET_NAMES.DON_HANG, dhRow, 22, maQuaMoi);
      updateCell(SHEET_NAMES.DON_HANG, dhRow, 23, tenQuaMoi);

      // Ghi nhận vào ghi chú đơn hàng
      var dateStr = formatDate(new Date());
      var newNote =
        (oldNote || "") +
        " [Đổi quà ngày " +
        dateStr +
        ": " +
        maQuaCu +
        " ➔ " +
        maQuaMoi +
        "]";
      updateCell(SHEET_NAMES.DON_HANG, dhRow, 20, newNote);

      rollbackActions.push(function() {
        try {
          updateCell(SHEET_NAMES.DON_HANG, dhRow, 22, oldMaQua);
          updateCell(SHEET_NAMES.DON_HANG, dhRow, 23, oldTenQua);
          updateCell(SHEET_NAMES.DON_HANG, dhRow, 20, oldNote);
        } catch(err){}
      });

      // Ghi nhận giao dịch vào bảng DoiTra
      var rowData = [
        maDT,
        new Date(), // NgayDoiTra
        maDH,
        maKH,
        tenKH,
        "Đổi quà",
        maQuaCu,
        tenQuaCu,
        "", // imei_Tra
        maQuaMoi,
        tenQuaMoi,
        "", // imei_Nhan
        0, // tienHoanTra
        Number(data.phiDoiTra) || 0,
        data.hinhThucThanhToan || "Tiền mặt",
        chiNhanh,
        data.nguoiThucHien || "",
        "Hoàn thành",
        data.ghiChu || "",
      ];

      var dtSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
      var dtRow = appendRow(SHEET_NAMES.DOI_TRA, rowData);
      rollbackActions.push(function() {
        try {
          dtSheet.deleteRow(dtRow);
          clearSheetCache(SHEET_NAMES.DOI_TRA);
        } catch(err){}
      });

      showToast("✅ Đổi quà tặng thành công đơn " + maDH);
      return maDT;

    } else {
      // ĐỐI TƯỢNG LÀ SẢN PHẨM CHÍNH
      var maSP_Tra = dhSheet.getRange(dhRow, 5).getValue();
      var tenSP_Tra = dhSheet.getRange(dhRow, 6).getValue();
      var nguonSP_Tra = dhSheet.getRange(dhRow, 7).getValue();

      if (nguonSP_Tra === "Phụ kiện") {
        var soLuong = Number(dhSheet.getRange(dhRow, 9).getValue()) || 1;

        // 1. Chuyển trạng thái đơn hàng gốc sang "Đổi trả"
        var oldDHStatus = dhSheet.getRange(dhRow, 19).getValue();
        updateCell(SHEET_NAMES.DON_HANG, dhRow, 19, "Đổi trả");
        rollbackActions.push(function() {
          try { updateCell(SHEET_NAMES.DON_HANG, dhRow, 19, oldDHStatus); } catch(err){}
        });

        // 2. Hoàn trả kho phụ kiện cũ vào chi nhánh của đơn hàng gốc
        var originalBranch = dhSheet.getRange(dhRow, 21).getValue();
        updateTonKhoPhuKien(maSP_Tra, soLuong, "nhap", originalBranch);
        rollbackActions.push(function() {
          try { updateTonKhoPhuKien(maSP_Tra, soLuong, "xuat", originalBranch); } catch(err){}
        });

        var maSP_Nhan = "";
        var tenSP_Nhan = "";

        // 3. Nếu là Đổi hàng -> Tạo đơn hàng mới cho phụ kiện mới nhận
        if (loaiGD === "Đổi hàng") {
          maSP_Nhan = data.maSP_Nhan;
          if (!maSP_Nhan) {
            throw new Error("Vui lòng chọn phụ kiện mới để đổi!");
          }

          var pkSheet = ss.getSheetByName(SHEET_NAMES.PHU_KIEN);
          var pkRow = findPhuKienRow(maSP_Nhan, chiNhanh);
          if (pkRow === -1) {
            throw new Error("Phụ kiện mới " + maSP_Nhan + " không tồn tại ở chi nhánh " + chiNhanh);
          }
          var tonMoi = Number(pkSheet.getRange(pkRow, 7).getValue()) || 0;
          if (tonMoi < 1) {
            throw new Error("Phụ kiện mới " + maSP_Nhan + " đã hết hàng tại chi nhánh " + chiNhanh);
          }

          tenSP_Nhan = pkSheet.getRange(pkRow, 2).getValue() || "";
          var giaBanMoi = Number(pkSheet.getRange(pkRow, 6).getValue()) || 0;

          // Tạo đơn hàng mới cho phụ kiện nhận
          var maDHMoi = taoDonHang({
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

          rollbackActions.push(function() {
            try {
              huyDonHang(maDHMoi);
              var oSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
              var oRow = findRow(SHEET_NAMES.DON_HANG, 1, maDHMoi);
              if (oRow !== -1) {
                oSheet.deleteRow(oRow);
                clearSheetCache(SHEET_NAMES.DON_HANG);
              }
            } catch(err){}
          });

          data.ghiChu = (data.ghiChu || "") + " [Đổi sang đơn mới: " + maDHMoi + "]";
        }

        // 4. Ghi nhận giao dịch vào bảng DoiTra
        var rowData = [
          maDT,
          new Date(), // NgayDoiTra
          maDH,
          maKH,
          tenKH,
          loaiGD,
          maSP_Tra,
          tenSP_Tra,
          "", // imei_Tra
          maSP_Nhan,
          tenSP_Nhan,
          "", // imei_Nhan
          (data.hinhThucThanhToan === "Hỗn hợp" && data.splitChuyenKhoan !== undefined)
            ? (data.splitChuyenKhoan + "," + data.splitTienMat)
            : (Number(data.tienHoanTra) || 0),
          Number(data.phiDoiTra) || 0,
          data.hinhThucThanhToan || "Tiền mặt",
          chiNhanh,
          data.nguoiThucHien || "",
          "Hoàn thành",
          data.ghiChu || "",
        ];

        var dtSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
        var dtRow = appendRow(SHEET_NAMES.DOI_TRA, rowData);
        rollbackActions.push(function() {
          try {
            dtSheet.deleteRow(dtRow);
            clearSheetCache(SHEET_NAMES.DOI_TRA);
          } catch(err){}
        });

        showToast("✅ Ghi nhận Đổi trả phụ kiện thành công: " + maDT);
        return maDT;

      } else {
        // ĐỐI TƯỢNG LÀ ĐIỆN THOẠI (logic cũ)
        var ghiChuDH = dhSheet.getRange(dhRow, 20).getValue() || "";
        var imeiMatch = ghiChuDH.match(/\[IMEI:\s*([^\s\]]+)\]/);
        var imei_Tra = imeiMatch ? imeiMatch[1] : (lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP_Tra, 4) || "");

        // 1. Đổi trạng thái đơn hàng gốc sang "Đổi trả" để trừ doanh thu/hoa hồng
        var oldDHStatus = dhSheet.getRange(dhRow, 19).getValue();
        updateCell(SHEET_NAMES.DON_HANG, dhRow, 19, "Đổi trả");
        rollbackActions.push(function() {
          try { updateCell(SHEET_NAMES.DON_HANG, dhRow, 19, oldDHStatus); } catch(err){}
        });

        // 2. Hoàn trả kho máy trả về trạng thái "Còn hàng"
        var oldPhoneStatus = lookupValue(SHEET_NAMES.DIEN_THOAI, 4, imei_Tra, 11);
        if (oldPhoneStatus === null) oldPhoneStatus = lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP_Tra, 11);
        updateTrangThaiKhoDT(imei_Tra || maSP_Tra, "Còn hàng");
        rollbackActions.push(function() {
          try { updateTrangThaiKhoDT(imei_Tra || maSP_Tra, oldPhoneStatus); } catch(err){}
        });

        // 3. Huỷ hợp đồng trả góp gốc nếu có
        var maTG = lookupValue(SHEET_NAMES.TRA_GOP, 2, maDH, 1);
        var oldTGStatus = "";
        var oldLSTGStates = [];
        if (maTG) {
          var tgRow = findRow(SHEET_NAMES.TRA_GOP, 1, maTG);
          if (tgRow !== -1) {
            oldTGStatus = ss.getSheetByName(SHEET_NAMES.TRA_GOP).getRange(tgRow, 16).getValue();
            var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
            var lastRow = lstgSheet.getLastRow();
            if (lastRow > 1) {
              var lstgData = lstgSheet.getRange(2, 2, lastRow - 1, 8).getValues();
              for (var i = 0; i < lstgData.length; i++) {
                if (String(lstgData[i][0]) === maTG) {
                  oldLSTGStates.push({
                    row: i + 2,
                    status: String(lstgData[i][7])
                  });
                }
              }
            }
          }
        }

        huyHopDongTraGop(maDH);

        rollbackActions.push(function() {
          try {
            if (maTG) {
              var tgRow = findRow(SHEET_NAMES.TRA_GOP, 1, maTG);
              if (tgRow !== -1) {
                updateCell(SHEET_NAMES.TRA_GOP, tgRow, 16, oldTGStatus);
              }
              var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
              oldLSTGStates.forEach(function(item) {
                lstgSheet.getRange(item.row, 9).setValue(item.status);
              });
              clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);
            }
          } catch(err){}
        });

        var maSP_Nhan = "";
        var tenSP_Nhan = "";
        var imei_Nhan = "";

        // 4. Nếu là Đổi máy -> Tạo đơn hàng mới cho máy nhận
        if (loaiGD === "Đổi máy") {
          maSP_Nhan = data.maSP_Nhan;
          if (!maSP_Nhan) {
            throw new Error("Vui lòng chọn máy mới để đổi!");
          }

          tenSP_Nhan = lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP_Nhan, 2) || "";
          
          imei_Nhan = data.imei_Nhan;
          if (!imei_Nhan) {
            // Tự động tìm máy có mã SP_Nhan tại chi nhánh này và đang "Còn hàng"
            var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
            var dtData = dtSheet.getDataRange().getValues();
            for (var i = 1; i < dtData.length; i++) {
              if (String(dtData[i][0]) === maSP_Nhan && 
                  String(dtData[i][12]) === chiNhanh && 
                  String(dtData[i][10]) === "Còn hàng") {
                imei_Nhan = String(dtData[i][3]);
                break;
              }
            }
          }
          if (!imei_Nhan) {
            imei_Nhan = lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP_Nhan, 4) || "";
          }

          var giaBanMoi =
            Number(lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP_Nhan, 9)) || 0;

          // Tạo đơn hàng mới cho máy nhận
          var maDHMoi = taoDonHang({
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

          rollbackActions.push(function() {
            try {
              huyDonHang(maDHMoi);
              var oSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
              var oRow = findRow(SHEET_NAMES.DON_HANG, 1, maDHMoi);
              if (oRow !== -1) {
                oSheet.deleteRow(oRow);
                clearSheetCache(SHEET_NAMES.DON_HANG);
              }
            } catch(err){}
          });

          data.ghiChu = (data.ghiChu || "") + " [Đổi sang đơn mới: " + maDHMoi + "]";
        }

        // 5. Ghi nhận giao dịch vào bảng DoiTra
        var rowData = [
          maDT,
          new Date(), // NgayDoiTra
          maDH,
          maKH,
          tenKH,
          loaiGD,
          maSP_Tra,
          tenSP_Tra,
          imei_Tra,
          maSP_Nhan,
          tenSP_Nhan,
          imei_Nhan,
          (data.hinhThucThanhToan === "Hỗn hợp" && data.splitChuyenKhoan !== undefined)
            ? (data.splitChuyenKhoan + "," + data.splitTienMat)
            : (Number(data.tienHoanTra) || 0),
          Number(data.phiDoiTra) || 0,
          data.hinhThucThanhToan || "Tiền mặt",
          chiNhanh,
          data.nguoiThucHien || "",
          "Hoàn thành",
          data.ghiChu || "",
        ];

        var dtSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
        var dtRow = appendRow(SHEET_NAMES.DOI_TRA, rowData);
        rollbackActions.push(function() {
          try {
            dtSheet.deleteRow(dtRow);
            clearSheetCache(SHEET_NAMES.DOI_TRA);
          } catch(err){}
        });

        showToast("✅ Ghi nhận Đổi trả thành công: " + maDT);
        return maDT;
      }
    }
  } catch(e) {
    for (var rIdx = rollbackActions.length - 1; rIdx >= 0; rIdx--) {
      try { rollbackActions[rIdx](); } catch(err){}
    }
    throw e;
  }
}
