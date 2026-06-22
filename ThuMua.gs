/**
 * ============================================================
 * VanTran Mobile — ThuMua.gs
 * Nghiệp vụ thu mua điện thoại của khách hàng & Thu cũ đổi mới
 * ============================================================
 */

/**
 * Thực hiện giao dịch Thu mua máy từ khách
 *
 * @param {Object} data - {
 *   maKH, soDienThoaiKH, loaiGiaoDich ('Bán thẳng'/'Thu cũ đổi mới'),
 *   tenSP_Thu, thuongHieu_Thu, imei_Thu, mauSac_Thu, dungLuong_Thu, tinhTrang_Thu,
 *   giaThuMua, tienHoTro, hinhThucThanhToan ('Tiền mặt'/'Chuyển khoản'),
 *   chiNhanh, nguoiThucHien (MaNV), ghiChu,
 *   // Nếu Thu cũ đổi mới:
 *   maSP_Moi
 * }
 * @return {string} Mã thu mua mới
 */
function thucHienThuMua(data) {
    initializeColumnEnums();
    var maTM = generateId("TM", SHEET_NAMES.THU_MUA);
    var chiNhanh = data.chiNhanh;
    var loaiGD = data.loaiGiaoDich;

    if (!data.maKH || !loaiGD || !data.imei_Thu || !chiNhanh) {
      throw new Error(
        "Vui lòng nhập đầy đủ Mã khách hàng, IMEI máy thu mua, Loại giao dịch và Chi nhánh!",
      );
    }

    var tenKH = ensureKhachHangExists(data.maKH, data.tenKH);
    if (!tenKH) {
      tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, 1, data.maKH, 2) || "";
    }
    var giaThuMua = Number(data.giaThuMua) || 0;
    var tienHoTro = Number(data.tienHoTro) || 0;
    var tongTienTraKhach = giaThuMua + tienHoTro;

    var rollbackActions = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    try {
      // 1. Tự động đưa máy thu mua vào kho Điện thoại
      var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
      var existingRow = findRow(
        SHEET_NAMES.DIEN_THOAI,
        COL_DT.IMEI,
        data.imei_Thu,
      ); // Tìm theo IMEI
      var maSP_Thu = "";

      if (existingRow !== -1) {
        // Máy đã từng tồn tại ở kho cửa hàng (khách mua ở tiệm rồi bán lại) -> Cập nhật trạng thái và giá nhập mới
        var oldGiaNhap = dtSheet
          .getRange(existingRow, COL_DT.GIA_NHAP)
          .getValue();
        var oldTrangThai = dtSheet
          .getRange(existingRow, COL_DT.TRANG_THAI_KHO)
          .getValue();
        var oldChiNhanh = dtSheet
          .getRange(existingRow, COL_DT.CHI_NHANH)
          .getValue();
        var oldTinhTrang = dtSheet
          .getRange(existingRow, COL_DT.TINH_TRANG)
          .getValue();

        maSP_Thu = dtSheet.getRange(existingRow, COL_DT.MA_DT).getValue();
        dtSheet
          .getRange(existingRow, COL_DT.GIA_NHAP)
          .setValue(tongTienTraKhach); // Giá nhập mới = Tổng giá trị thu mua
        dtSheet
          .getRange(existingRow, COL_DT.TRANG_THAI_KHO)
          .setValue("Còn hàng"); // Trở lại kho
        dtSheet.getRange(existingRow, COL_DT.CHI_NHANH).setValue(chiNhanh); // Cập nhật chi nhánh hiện tại
        dtSheet
          .getRange(existingRow, COL_DT.TINH_TRANG)
          .setValue(data.tinhTrang_Thu || "Đã qua sử dụng"); // Cập nhật tình trạng máy
        clearSheetCache(SHEET_NAMES.DIEN_THOAI);
        invalidateDropdownCache(SHEET_NAMES.DIEN_THOAI);

        rollbackActions.push(function () {
          try {
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
            dtSheet.getRange(existingRow, COL_DT.GIA_NHAP).setValue(oldGiaNhap);
            dtSheet
              .getRange(existingRow, COL_DT.TRANG_THAI_KHO)
              .setValue(oldTrangThai);
            dtSheet
              .getRange(existingRow, COL_DT.CHI_NHANH)
              .setValue(oldChiNhanh);
            dtSheet
              .getRange(existingRow, COL_DT.TINH_TRANG)
              .setValue(oldTinhTrang);
            clearSheetCache(SHEET_NAMES.DIEN_THOAI);
            invalidateDropdownCache(SHEET_NAMES.DIEN_THOAI);
          } catch (err) {
            Logger.log(
              "Rollback failed to restore phone values: " + err.message,
            );
          }
        });
        showToast("Cập nhật lại kho máy thu mua: " + maSP_Thu);
      } else {
        // Máy mới hoàn toàn -> Thêm mới vào danh mục
        maSP_Thu = addDienThoai({
          tenSP: data.tenSP_Thu,
          thuongHieu: data.thuongHieu_Thu,
          imei: data.imei_Thu,
          mauSac: data.mauSac_Thu,
          dungLuong: data.dungLuong_Thu,
          tinhTrang: data.tinhTrang_Thu || "Đã qua sử dụng",
          giaNhap: tongTienTraKhach,
          giaBan: 0, // Mặc định giá bán lẻ = 0, chủ tiệm tự sửa đổi sau
          giaTraGop: 0,
          chiNhanh: chiNhanh,
          ghiChu: "Thu mua từ khách hàng " + tenKH,
        });

        rollbackActions.push(function () {
          try {
            var ss = SpreadsheetApp.getActiveSpreadsheet();
            var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
            var row = findRow(SHEET_NAMES.DIEN_THOAI, COL_DT.MA_DT, maSP_Thu);
            if (row !== -1) {
              dtSheet.deleteRow(row);
              clearSheetCache(SHEET_NAMES.DIEN_THOAI);
              invalidateDropdownCache(SHEET_NAMES.DIEN_THOAI);
            }
          } catch (err) {
            Logger.log("Rollback failed to delete new phone: " + err.message);
          }
        });
      }

      var maDH_Moi = "";

      // 2. Nếu là Thu cũ đổi mới -> Tạo đơn hàng mua máy mới
      if (loaiGD === "Thu cũ đổi mới") {
        var maSP_Moi = data.maSP_Moi;
        if (!maSP_Moi) {
          throw new Error("Vui lòng chọn máy mới để đổi lên đời!");
        }

        var giaBanMoi =
          Number(
            lookupValue(
              SHEET_NAMES.DIEN_THOAI,
              COL_DT.MA_DT,
              maSP_Moi,
              COL_DT.GIA_BAN,
            ),
          ) || 0;

        var hinhThucBan = data.hinhThucBan || "Bán thẳng";
        var giaThuMua = Number(data.giaThuMua) || 0;
        var tienHoTro = Number(data.tienHoTro) || 0;
        var tienGiamGia = Number(data.tienGiamGia) || 0;

        // Tạo đơn hàng mua máy mới
        var donHangHTTT = data.hinhThucThanhToanPhu || "Tiền mặt";
        var donHangSplitTM = undefined;
        var donHangSplitCK = undefined;

        if (donHangHTTT === "Hỗn hợp") {
          donHangSplitTM = data.splitTienMatPhu;
          donHangSplitCK = data.splitChuyenKhoanPhu;
        }

        maDH_Moi = taoDonHang({
          maKH: data.maKH,
          maSP: maSP_Moi,
          imei: data.imei_Moi,
          nguonSP: "Điện thoại",
          soLuong: 1,
          donGia: giaBanMoi,
          hinhThucBan: hinhThucBan,
          hinhThucThanhToan: donHangHTTT,
          splitChuyenKhoan: donHangSplitCK,
          splitTienMat: donHangSplitTM,
          nguoiBan: data.nguoiThucHien,
          chiNhanh: chiNhanh,
          ghiChu:
            "Đơn hàng Thu cũ đổi mới, liên kết tới giao dịch thu mua " + maTM,
          coNhanQua: data.coNhanQua || "✗",
          maQuaTang: data.maQuaTang || "",
          tienGiamGia: tienGiamGia,
          traGop:
            hinhThucBan === "Trả góp" && data.traGop
              ? {
                  traTruoc: data.traGop.traTruoc,
                  soKy: data.traGop.soKy,
                  loaiTraGop: data.traGop.loaiTraGop,
                  congTyTC: data.traGop.congTyTC,
                  tienMoiKy: data.traGop.tienMoiKy,
                  tongTien: giaBanMoi - tienGiamGia - tienHoTro,
                }
              : undefined,
        });

        rollbackActions.push(function () {
          try {
            huyDonHang(maDH_Moi);
            var dhSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
            var dhRow = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH_Moi);
            if (dhRow !== -1) {
              dhSheet.deleteRow(dhRow);
              clearSheetCache(SHEET_NAMES.DON_HANG);
            }
          } catch (err) {
            Logger.log("Rollback failed to delete order: " + err.message);
          }
        });
      }

      // 3. Ghi nhận giao dịch vào bảng ThuMua
      var rowData = [];
      rowData[COL_TM.MA_TM - 1] = maTM;
      rowData[COL_TM.NGAY_TM - 1] = new Date();
      rowData[COL_TM.MA_KH - 1] = data.maKH;
      rowData[COL_TM.TEN_KH - 1] = tenKH;
      rowData[COL_TM.SDT_KH - 1] = data.soDienThoaiKH || "";
      rowData[COL_TM.TEN_SP_THU - 1] = data.tenSP_Thu || "";
      rowData[COL_TM.THUONG_HIEU_THU - 1] = data.thuongHieu_Thu || "";
      rowData[COL_TM.IMEI_THU - 1] = data.imei_Thu;
      rowData[COL_TM.MAU_SAC_THU - 1] = data.mauSac_Thu || "";
      rowData[COL_TM.DUNG_LUONG_THU - 1] = data.dungLuong_Thu || "";
      rowData[COL_TM.TINH_TRANG_THU - 1] =
        data.tinhTrang_Thu || "Đã qua sử dụng";
      rowData[COL_TM.GIA_THU_MUA - 1] = giaThuMua;
      rowData[COL_TM.LOAI_GD - 1] = loaiGD;
      rowData[COL_TM.MA_DH_MOI - 1] = maDH_Moi;
      rowData[COL_TM.TIEN_HO_TRO - 1] = tienHoTro;
      rowData[COL_TM.TONG_TIEN_TRA - 1] = tongTienTraKhach;
      rowData[COL_TM.HINH_THUC_TT - 1] = data.hinhThucThanhToan || "Tiền mặt";
      rowData[COL_TM.CHI_NHANH - 1] = chiNhanh;
      rowData[COL_TM.NGUOI_THUC_HIEN - 1] = data.nguoiThucHien || "";
      rowData[COL_TM.GHI_CHU - 1] = data.ghiChu || "";

      // Backend validation for Hỗn hợp buyback payment
      if (data.hinhThucThanhToan === "Hỗn hợp") {
        var splitTienMat = Number(data.splitTienMat) || 0;
        var splitChuyenKhoan = Number(data.splitChuyenKhoan) || 0;
        var totalNeeded = tongTienTraKhach;
        if (Math.abs(splitTienMat + splitChuyenKhoan - totalNeeded) > 1) {
          throw new Error(
            "Lỗi dữ liệu: Tổng tiền mặt (" +
              splitTienMat +
              ") và chuyển khoản (" +
              splitChuyenKhoan +
              ") không khớp với số tiền cần thanh toán (" +
              totalNeeded +
              ")!",
          );
        }
      }

      var tienMat = 0;
      var chuyenKhoan = 0;
      if (data.hinhThucThanhToan === "Hỗn hợp") {
        tienMat = Number(data.splitTienMat) || 0;
        chuyenKhoan = Number(data.splitChuyenKhoan) || 0;
      } else if (data.hinhThucThanhToan === "Tiền mặt") {
        tienMat = tongTienTraKhach;
      } else {
        chuyenKhoan = tongTienTraKhach;
      }
      rowData[COL_TM.TIEN_MAT - 1] = tienMat;
      rowData[COL_TM.CHUYEN_KHOAN - 1] = chuyenKhoan;

      var tmSheet = ss.getSheetByName(SHEET_NAMES.THU_MUA);
      var tmRow = appendRow(SHEET_NAMES.THU_MUA, rowData);
      rollbackActions.push(function () {
        try {
          var ss = SpreadsheetApp.getActiveSpreadsheet();
          var tmSheet = ss.getSheetByName(SHEET_NAMES.THU_MUA);
          tmSheet.deleteRow(tmRow);
          clearSheetCache(SHEET_NAMES.THU_MUA);
        } catch (err) {
          Logger.log("Rollback failed to delete buyback row: " + err.message);
        }
      });
    } catch (e) {
      for (var rIdx = rollbackActions.length - 1; rIdx >= 0; rIdx--) {
        rollbackActions[rIdx]();
      }
      throw e;
    }

    showToast("✅ Đã ghi nhận thu mua thành công: " + maTM);
    return maTM;
}
