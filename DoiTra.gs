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
    initializeColumnEnums();

    // Backend validation for Hỗn hợp payments
    var expectedPaid = 0;
    var isHoanTra = (Number(data.tienHoanTra) || 0) > 0;
    if (isHoanTra) {
      expectedPaid = Number(data.tienHoanTra) || 0;
    } else {
      expectedPaid = Number(data.phiDoiTra) || 0;
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
            ") không khớp với số tiền cần thanh toán (" +
            expectedPaid +
            ")!",
        );
      }
    }

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
    var dhRow = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH);
    if (dhRow === -1) {
      throw new Error("Không tìm thấy đơn hàng gốc: " + maDH);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dhSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);

    var dhStatus = dhSheet.getRange(dhRow, COL_DH.TRANG_THAI).getValue();
    if (dhStatus === "Huỷ" || dhStatus === "Đổi trả") {
      throw new Error(
        "Đơn hàng " +
          maDH +
          " đã ở trạng thái: " +
          dhStatus +
          ", không thể đổi trả!",
      );
    }

    var maKH = dhSheet.getRange(dhRow, COL_DH.MA_KH).getValue();
    var tenKH = dhSheet.getRange(dhRow, COL_DH.TEN_KH).getValue();

    // XỬ LÝ THEO ĐỐI TƯỢNG ĐỔI TRẢ
    var rollbackActions = [];
    try {
      if (doiTuong === "Quà tặng kèm") {
        var maQuaCu = dhSheet.getRange(dhRow, COL_DH.MA_QUA_TANG).getValue();
        var tenQuaCu =
          dhSheet.getRange(dhRow, COL_DH.TEN_QUA_TANG).getValue() ||
          "(Không tên)";
        var coNhanQua = dhSheet.getRange(dhRow, COL_DH.CO_NHAN_QUA).getValue();

        if (coNhanQua !== "✓") {
          throw new Error(
            "Đơn hàng " + maDH + " không nhận quà tặng kèm trước đó!",
          );
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
          var tonQuaMoi =
            Number(pkSheet.getRange(newRow, COL_PK.SO_LUONG_TON).getValue()) ||
            0;
          var requiredQty = codeCounts[code];
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
          var giftName =
            pkSheet.getRange(newRow, COL_PK.TEN_SP).getValue() || "";
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
          rollbackActions.push(function () {
            returnedGifts.forEach(function (code) {
              try {
                updateTonKhoPhuKien(code, 1, "xuat", chiNhanh);
              } catch (err) {}
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
        rollbackActions.push(function () {
          issuedGifts.forEach(function (code) {
            try {
              updateTonKhoPhuKien(code, 1, "nhap", chiNhanh);
            } catch (err) {}
          });
        });

        // Cập nhật thông tin đơn hàng gốc
        var oldMaQua = dhSheet.getRange(dhRow, COL_DH.MA_QUA_TANG).getValue();
        var oldTenQua = dhSheet.getRange(dhRow, COL_DH.TEN_QUA_TANG).getValue();
        var oldNote = dhSheet.getRange(dhRow, COL_DH.GHI_CHU).getValue();

        updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.MA_QUA_TANG, maQuaMoi);
        updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.TEN_QUA_TANG, tenQuaMoi);

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
        updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.GHI_CHU, newNote);

        rollbackActions.push(function () {
          try {
            updateCell(
              SHEET_NAMES.DON_HANG,
              dhRow,
              COL_DH.MA_QUA_TANG,
              oldMaQua,
            );
            updateCell(
              SHEET_NAMES.DON_HANG,
              dhRow,
              COL_DH.TEN_QUA_TANG,
              oldTenQua,
            );
            updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.GHI_CHU, oldNote);
          } catch (err) {}
        });

        // Ghi nhận giao dịch vào bảng DoiTra
        var rowData = [];
        rowData[COL_DT_TRA.MA_DT - 1] = maDT;
        rowData[COL_DT_TRA.NGAY_DT - 1] = new Date();
        rowData[COL_DT_TRA.MA_DH - 1] = maDH;
        rowData[COL_DT_TRA.MA_KH - 1] = maKH;
        rowData[COL_DT_TRA.TEN_KH - 1] = tenKH;
        rowData[COL_DT_TRA.LOAI_GD - 1] = "Đổi quà";
        rowData[COL_DT_TRA.MA_SP_TRA - 1] = maQuaCu;
        rowData[COL_DT_TRA.TEN_SP_TRA - 1] = tenQuaCu;
        rowData[COL_DT_TRA.IMEI_TRA - 1] = "";
        rowData[COL_DT_TRA.MA_SP_NHAN - 1] = maQuaMoi;
        rowData[COL_DT_TRA.TEN_SP_NHAN - 1] = tenQuaMoi;
        rowData[COL_DT_TRA.IMEI_NHAN - 1] = "";
        rowData[COL_DT_TRA.TIEN_HOAN_TRA - 1] = 0;
        rowData[COL_DT_TRA.PHI_DOI_TRA - 1] = Number(data.phiDoiTra) || 0;
        rowData[COL_DT_TRA.HINH_THUC_TT - 1] =
          data.hinhThucThanhToan || "Tiền mặt";
        rowData[COL_DT_TRA.CHI_NHANH - 1] = chiNhanh;
        rowData[COL_DT_TRA.NGUOI_THUC_HIEN - 1] = data.nguoiThucHien || "";
        rowData[COL_DT_TRA.TRANG_THAI - 1] = "Hoàn thành";
        rowData[COL_DT_TRA.GHI_CHU - 1] = data.ghiChu || "";

        var tienMat = 0;
        var chuyenKhoan = 0;
        var totalPay = Number(data.phiDoiTra) || 0;
        if (data.hinhThucThanhToan === "Hỗn hợp") {
          tienMat = Number(data.splitTienMat) || 0;
          chuyenKhoan = Number(data.splitChuyenKhoan) || 0;
        } else if (data.hinhThucThanhToan === "Tiền mặt") {
          tienMat = totalPay;
        } else {
          chuyenKhoan = totalPay;
        }
        rowData[COL_DT_TRA.TIEN_MAT - 1] = tienMat;
        rowData[COL_DT_TRA.CHUYEN_KHOAN - 1] = chuyenKhoan;

        var dtSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
        var dtRow = appendRow(SHEET_NAMES.DOI_TRA, rowData);
        rollbackActions.push(function () {
          try {
            dtSheet.deleteRow(dtRow);
            clearSheetCache(SHEET_NAMES.DOI_TRA);
          } catch (err) {}
        });

        showToast("✅ Đổi quà tặng thành công đơn " + maDH);
        return maDT;
      } else {
        // ĐỐI TƯỢNG LÀ SẢN PHẨM CHÍNH
        var maSP_Tra = dhSheet.getRange(dhRow, COL_DH.MA_SP).getValue();
        var tenSP_Tra = dhSheet.getRange(dhRow, COL_DH.TEN_SP).getValue();
        var nguonSP_Tra = dhSheet.getRange(dhRow, COL_DH.NGUON_SP).getValue();

        if (nguonSP_Tra === "Phụ kiện") {
          var soLuong =
            Number(dhSheet.getRange(dhRow, COL_DH.SO_LUONG).getValue()) || 1;

          // 1. Chuyển trạng thái đơn hàng gốc sang "Đổi trả"
          var oldDHStatus = dhSheet
            .getRange(dhRow, COL_DH.TRANG_THAI)
            .getValue();
          updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.TRANG_THAI, "Đổi trả");
          rollbackActions.push(function () {
            try {
              updateCell(
                SHEET_NAMES.DON_HANG,
                dhRow,
                COL_DH.TRANG_THAI,
                oldDHStatus,
              );
            } catch (err) {}
          });

          // 2. Hoàn trả kho phụ kiện cũ vào chi nhánh của đơn hàng gốc
          var originalBranch = dhSheet
            .getRange(dhRow, COL_DH.CHI_NHANH)
            .getValue();
          updateTonKhoPhuKien(maSP_Tra, soLuong, "nhap", originalBranch);
          rollbackActions.push(function () {
            try {
              updateTonKhoPhuKien(maSP_Tra, soLuong, "xuat", originalBranch);
            } catch (err) {}
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
              throw new Error(
                "Phụ kiện mới " +
                  maSP_Nhan +
                  " không tồn tại ở chi nhánh " +
                  chiNhanh,
              );
            }
            var tonMoi =
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

            tenSP_Nhan =
              pkSheet.getRange(pkRow, COL_PK.TEN_SP).getValue() || "";
            var giaBanMoi =
              Number(pkSheet.getRange(pkRow, COL_PK.GIA_BAN).getValue()) || 0;

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

            rollbackActions.push(function () {
              try {
                huyDonHang(maDHMoi);
                var oSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
                var oRow = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDHMoi);
                if (oRow !== -1) {
                  oSheet.deleteRow(oRow);
                  clearSheetCache(SHEET_NAMES.DON_HANG);
                }
              } catch (err) {}
            });

            data.ghiChu =
              (data.ghiChu || "") + " [Đổi sang đơn mới: " + maDHMoi + "]";
          }

          // 4. Ghi nhận giao dịch vào bảng DoiTra
          var tienHoanTra = Number(data.tienHoanTra) || 0;
          var phiDoiTra = Number(data.phiDoiTra) || 0;

          var rowData = [];
          rowData[COL_DT_TRA.MA_DT - 1] = maDT;
          rowData[COL_DT_TRA.NGAY_DT - 1] = new Date();
          rowData[COL_DT_TRA.MA_DH - 1] = maDH;
          rowData[COL_DT_TRA.MA_KH - 1] = maKH;
          rowData[COL_DT_TRA.TEN_KH - 1] = tenKH;
          rowData[COL_DT_TRA.LOAI_GD - 1] = loaiGD;
          rowData[COL_DT_TRA.MA_SP_TRA - 1] = maSP_Tra;
          rowData[COL_DT_TRA.TEN_SP_TRA - 1] = tenSP_Tra;
          rowData[COL_DT_TRA.IMEI_TRA - 1] = "";
          rowData[COL_DT_TRA.MA_SP_NHAN - 1] = maSP_Nhan;
          rowData[COL_DT_TRA.TEN_SP_NHAN - 1] = tenSP_Nhan;
          rowData[COL_DT_TRA.IMEI_NHAN - 1] = "";
          rowData[COL_DT_TRA.TIEN_HOAN_TRA - 1] = tienHoanTra;
          rowData[COL_DT_TRA.PHI_DOI_TRA - 1] = phiDoiTra;
          rowData[COL_DT_TRA.HINH_THUC_TT - 1] =
            data.hinhThucThanhToan || "Tiền mặt";
          rowData[COL_DT_TRA.CHI_NHANH - 1] = chiNhanh;
          rowData[COL_DT_TRA.NGUOI_THUC_HIEN - 1] = data.nguoiThucHien || "";
          rowData[COL_DT_TRA.TRANG_THAI - 1] = "Hoàn thành";
          rowData[COL_DT_TRA.GHI_CHU - 1] = data.ghiChu || "";

          var tienMat = 0;
          var chuyenKhoan = 0;
          var totalPay = tienHoanTra > 0 ? tienHoanTra : phiDoiTra;
          if (data.hinhThucThanhToan === "Hỗn hợp") {
            tienMat = Number(data.splitTienMat) || 0;
            chuyenKhoan = Number(data.splitChuyenKhoan) || 0;
          } else if (data.hinhThucThanhToan === "Tiền mặt") {
            tienMat = totalPay;
          } else {
            chuyenKhoan = totalPay;
          }
          rowData[COL_DT_TRA.TIEN_MAT - 1] = tienMat;
          rowData[COL_DT_TRA.CHUYEN_KHOAN - 1] = chuyenKhoan;

          var dtSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
          var dtRow = appendRow(SHEET_NAMES.DOI_TRA, rowData);
          rollbackActions.push(function () {
            try {
              dtSheet.deleteRow(dtRow);
              clearSheetCache(SHEET_NAMES.DOI_TRA);
            } catch (err) {}
          });

          showToast("✅ Ghi nhận Đổi trả phụ kiện thành công: " + maDT);
          return maDT;
        } else {
          // ĐỐI TƯỢNG LÀ ĐIỆN THOẠI (logic cũ)
          var ghiChuDH =
            dhSheet.getRange(dhRow, COL_DH.GHI_CHU).getValue() || "";
          var imeiMatch = ghiChuDH.match(/\[IMEI:\s*([^\s\]]+)\]/);
          var imei_Tra = imeiMatch
            ? imeiMatch[1]
            : lookupValue(
                SHEET_NAMES.DIEN_THOAI,
                COL_DT.MA_DT,
                maSP_Tra,
                COL_DT.IMEI,
              ) || "";

          // 1. Đổi trạng thái đơn hàng gốc sang "Đổi trả" để trừ doanh thu/hoa hồng
          var oldDHStatus = dhSheet
            .getRange(dhRow, COL_DH.TRANG_THAI)
            .getValue();
          updateCell(SHEET_NAMES.DON_HANG, dhRow, COL_DH.TRANG_THAI, "Đổi trả");
          rollbackActions.push(function () {
            try {
              updateCell(
                SHEET_NAMES.DON_HANG,
                dhRow,
                COL_DH.TRANG_THAI,
                oldDHStatus,
              );
            } catch (err) {}
          });

          // 2. Hoàn trả kho máy trả về trạng thái "Còn hàng"
          var oldPhoneStatus = lookupValue(
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
          updateTrangThaiKhoDT(imei_Tra || maSP_Tra, "Còn hàng");
          rollbackActions.push(function () {
            try {
              updateTrangThaiKhoDT(imei_Tra || maSP_Tra, oldPhoneStatus);
            } catch (err) {}
          });

          // 3. Huỷ hợp đồng trả góp gốc nếu có
          var maTG = lookupValue(SHEET_NAMES.TRA_GOP, 2, maDH, 1);
          var oldTGStatus = "";
          var oldLSTGStates = [];
          if (maTG) {
            var tgRow = findRow(SHEET_NAMES.TRA_GOP, 1, maTG);
            if (tgRow !== -1) {
              oldTGStatus = ss
                .getSheetByName(SHEET_NAMES.TRA_GOP)
                .getRange(tgRow, 16)
                .getValue();
              var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
              var lastRow = lstgSheet.getLastRow();
              if (lastRow > 1) {
                var lstgData = lstgSheet
                  .getRange(2, 2, lastRow - 1, 8)
                  .getValues();
                for (var i = 0; i < lstgData.length; i++) {
                  if (String(lstgData[i][0]) === maTG) {
                    oldLSTGStates.push({
                      row: i + 2,
                      status: String(lstgData[i][7]),
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
                var tgRow = findRow(SHEET_NAMES.TRA_GOP, 1, maTG);
                if (tgRow !== -1) {
                  updateCell(SHEET_NAMES.TRA_GOP, tgRow, 16, oldTGStatus);
                }
                var lstgSheet = ss.getSheetByName(SHEET_NAMES.LICH_SU_TRA_GOP);
                oldLSTGStates.forEach(function (item) {
                  lstgSheet.getRange(item.row, 9).setValue(item.status);
                });
                clearSheetCache(SHEET_NAMES.LICH_SU_TRA_GOP);
              }
            } catch (err) {}
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

            tenSP_Nhan =
              lookupValue(
                SHEET_NAMES.DIEN_THOAI,
                COL_DT.MA_DT,
                maSP_Nhan,
                COL_DT.TEN_SP,
              ) || "";

            imei_Nhan = data.imei_Nhan;
            if (!imei_Nhan) {
              // Tự động tìm máy có mã SP_Nhan tại chi nhánh này và đang "Còn hàng"
              var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
              var dtData = getAllData(SHEET_NAMES.DIEN_THOAI);
              var maDTIdx = COL_DT.MA_DT - 1;
              var chiNhanhIdx = COL_DT.CHI_NHANH - 1;
              var trangThaiKhoIdx = COL_DT.TRANG_THAI_KHO - 1;
              var imeiIdx = COL_DT.IMEI - 1;

              for (var i = 0; i < dtData.length; i++) {
                if (
                  String(dtData[i][maDTIdx]) === maSP_Nhan &&
                  String(dtData[i][chiNhanhIdx]) === chiNhanh &&
                  String(dtData[i][trangThaiKhoIdx]) === "Còn hàng"
                ) {
                  imei_Nhan = String(dtData[i][imeiIdx]);
                  break;
                }
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

            var giaBanMoi =
              Number(
                lookupValue(
                  SHEET_NAMES.DIEN_THOAI,
                  COL_DT.MA_DT,
                  maSP_Nhan,
                  COL_DT.GIA_BAN,
                ),
              ) || 0;

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

            rollbackActions.push(function () {
              try {
                huyDonHang(maDHMoi);
                var oSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
                var oRow = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDHMoi);
                if (oRow !== -1) {
                  oSheet.deleteRow(oRow);
                  clearSheetCache(SHEET_NAMES.DON_HANG);
                }
              } catch (err) {}
            });

            data.ghiChu =
              (data.ghiChu || "") + " [Đổi sang đơn mới: " + maDHMoi + "]";
          }

          // 5. Ghi nhận giao dịch vào bảng DoiTra
          var tienHoanTra = Number(data.tienHoanTra) || 0;
          var phiDoiTra = Number(data.phiDoiTra) || 0;

          var rowData = [];
          rowData[COL_DT_TRA.MA_DT - 1] = maDT;
          rowData[COL_DT_TRA.NGAY_DT - 1] = new Date();
          rowData[COL_DT_TRA.MA_DH - 1] = maDH;
          rowData[COL_DT_TRA.MA_KH - 1] = maKH;
          rowData[COL_DT_TRA.TEN_KH - 1] = tenKH;
          rowData[COL_DT_TRA.LOAI_GD - 1] = loaiGD;
          rowData[COL_DT_TRA.MA_SP_TRA - 1] = maSP_Tra;
          rowData[COL_DT_TRA.TEN_SP_TRA - 1] = tenSP_Tra;
          rowData[COL_DT_TRA.IMEI_TRA - 1] = imei_Tra;
          rowData[COL_DT_TRA.MA_SP_NHAN - 1] = maSP_Nhan;
          rowData[COL_DT_TRA.TEN_SP_NHAN - 1] = tenSP_Nhan;
          rowData[COL_DT_TRA.IMEI_NHAN - 1] = imei_Nhan;
          rowData[COL_DT_TRA.TIEN_HOAN_TRA - 1] = tienHoanTra;
          rowData[COL_DT_TRA.PHI_DOI_TRA - 1] = phiDoiTra;
          rowData[COL_DT_TRA.HINH_THUC_TT - 1] =
            data.hinhThucThanhToan || "Tiền mặt";
          rowData[COL_DT_TRA.CHI_NHANH - 1] = chiNhanh;
          rowData[COL_DT_TRA.NGUOI_THUC_HIEN - 1] = data.nguoiThucHien || "";
          rowData[COL_DT_TRA.TRANG_THAI - 1] = "Hoàn thành";
          rowData[COL_DT_TRA.GHI_CHU - 1] = data.ghiChu || "";

          var tienMat = 0;
          var chuyenKhoan = 0;
          var totalPay = tienHoanTra > 0 ? tienHoanTra : phiDoiTra;
          if (data.hinhThucThanhToan === "Hỗn hợp") {
            tienMat = Number(data.splitTienMat) || 0;
            chuyenKhoan = Number(data.splitChuyenKhoan) || 0;
          } else if (data.hinhThucThanhToan === "Tiền mặt") {
            tienMat = totalPay;
          } else {
            chuyenKhoan = totalPay;
          }
          rowData[COL_DT_TRA.TIEN_MAT - 1] = tienMat;
          rowData[COL_DT_TRA.CHUYEN_KHOAN - 1] = chuyenKhoan;

          var dtSheet = ss.getSheetByName(SHEET_NAMES.DOI_TRA);
          var dtRow = appendRow(SHEET_NAMES.DOI_TRA, rowData);
          rollbackActions.push(function () {
            try {
              dtSheet.deleteRow(dtRow);
              clearSheetCache(SHEET_NAMES.DOI_TRA);
            } catch (err) {}
          });

          showToast("✅ Ghi nhận Đổi trả thành công: " + maDT);
          return maDT;
        }
      }
    } catch (e) {
      for (var rIdx = rollbackActions.length - 1; rIdx >= 0; rIdx--) {
        try {
          rollbackActions[rIdx]();
        } catch (err) {}
      }
      throw e;
    }
}
