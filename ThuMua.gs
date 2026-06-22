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
  var maTM = generateId("TM", SHEET_NAMES.THU_MUA);
  var chiNhanh = data.chiNhanh;
  var loaiGD = data.loaiGiaoDich;

  if (!data.maKH || !loaiGD || !data.imei_Thu || !chiNhanh) {
    throw new Error(
      "Vui lòng nhập đầy đủ Mã khách hàng, IMEI máy thu mua, Loại giao dịch và Chi nhánh!",
    );
  }

  var tenKH = lookupValue(SHEET_NAMES.KHACH_HANG, 1, data.maKH, 2) || "";
  var giaThuMua = Number(data.giaThuMua) || 0;
  var tienHoTro = Number(data.tienHoTro) || 0;
  var tongTienTraKhach = giaThuMua + tienHoTro;

  var rollbackActions = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // 1. Tự động đưa máy thu mua vào kho Điện thoại
    var dtSheet = ss.getSheetByName(SHEET_NAMES.DIEN_THOAI);
    var existingRow = findRow(SHEET_NAMES.DIEN_THOAI, 4, data.imei_Thu); // Tìm theo IMEI
    var maSP_Thu = "";

    if (existingRow !== -1) {
      // Máy đã từng tồn tại ở kho cửa hàng (khách mua ở tiệm rồi bán lại) -> Cập nhật trạng thái và giá nhập mới
      var oldGiaNhap = dtSheet.getRange(existingRow, 8).getValue();
      var oldTrangThai = dtSheet.getRange(existingRow, 11).getValue();
      var oldChiNhanh = dtSheet.getRange(existingRow, 13).getValue();
      var oldTinhTrang = dtSheet.getRange(existingRow, 7).getValue();

      maSP_Thu = dtSheet.getRange(existingRow, 1).getValue();
      dtSheet.getRange(existingRow, 8).setValue(tongTienTraKhach); // Giá nhập mới = Tổng giá trị thu mua
      dtSheet.getRange(existingRow, 11).setValue("Còn hàng"); // Trở lại kho
      dtSheet.getRange(existingRow, 13).setValue(chiNhanh); // Cập nhật chi nhánh hiện tại
      dtSheet
        .getRange(existingRow, 7)
        .setValue(data.tinhTrang_Thu || "Đã qua sử dụng"); // Cập nhật tình trạng máy
      clearSheetCache(SHEET_NAMES.DIEN_THOAI);
      invalidateDropdownCache(SHEET_NAMES.DIEN_THOAI);

      rollbackActions.push(function() {
        try {
          dtSheet.getRange(existingRow, 8).setValue(oldGiaNhap);
          dtSheet.getRange(existingRow, 11).setValue(oldTrangThai);
          dtSheet.getRange(existingRow, 13).setValue(oldChiNhanh);
          dtSheet.getRange(existingRow, 7).setValue(oldTinhTrang);
          clearSheetCache(SHEET_NAMES.DIEN_THOAI);
          invalidateDropdownCache(SHEET_NAMES.DIEN_THOAI);
        } catch (err) {
          Logger.log("Rollback failed to restore phone values: " + err.message);
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

      rollbackActions.push(function() {
        try {
          var row = findRow(SHEET_NAMES.DIEN_THOAI, 1, maSP_Thu);
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
        Number(lookupValue(SHEET_NAMES.DIEN_THOAI, 1, maSP_Moi, 9)) || 0;

      var hinhThucBan = data.hinhThucBan || "Bán thẳng";
      var giaThuMua = Number(data.giaThuMua) || 0;
      var tienHoTro = Number(data.tienHoTro) || 0;
      var tienGiamGia = Number(data.tienGiamGia) || 0;
      var chenhLech = (giaBanMoi - tienGiamGia) - tienHoTro - giaThuMua;

      // Tạo đơn hàng mua máy mới
      maDH_Moi = taoDonHang({
        maKH: data.maKH,
        maSP: maSP_Moi,
        imei: data.imei_Moi,
        nguonSP: "Điện thoại",
        soLuong: 1,
        donGia: giaBanMoi,
        hinhThucBan: hinhThucBan,
        hinhThucThanhToan: hinhThucBan === "Trả góp" ? (data.hinhThucThanhToanPhu || "Tiền mặt") : (data.hinhThucThanhToan || "Tiền mặt"),
        splitChuyenKhoan: data.splitChuyenKhoan,
        splitTienMat: data.splitTienMat,
        nguoiBan: data.nguoiThucHien,
        chiNhanh: chiNhanh,
        ghiChu: "Đơn hàng Thu cũ đổi mới, liên kết tới giao dịch thu mua " + maTM,
        coNhanQua: data.coNhanQua || "✗",
        maQuaTang: data.maQuaTang || "",
        tienGiamGia: tienGiamGia,
        traGop: hinhThucBan === "Trả góp" && data.traGop ? {
          traTruoc: data.traGop.traTruoc,
          soKy: data.traGop.soKy,
          loaiTraGop: data.traGop.loaiTraGop,
          congTyTC: data.traGop.congTyTC,
          tienMoiKy: data.traGop.tienMoiKy,
          tongTien: giaBanMoi - tienGiamGia - tienHoTro
        } : undefined
      });

      rollbackActions.push(function() {
        try {
          huyDonHang(maDH_Moi);
          var dhSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
          var dhRow = findRow(SHEET_NAMES.DON_HANG, 1, maDH_Moi);
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
    var rowData = [
      maTM,
      new Date(), // NgayThuMua
      data.maKH,
      tenKH,
      data.soDienThoaiKH || "",
      data.tenSP_Thu || "",
      data.thuongHieu_Thu || "",
      data.imei_Thu,
      data.mauSac_Thu || "",
      data.dungLuong_Thu || "",
      data.tinhTrang_Thu || "Đã qua sử dụng",
      giaThuMua,
      loaiGD,
      maDH_Moi,
      tienHoTro,
      (data.hinhThucThanhToan === "Hỗn hợp" && data.splitChuyenKhoan !== undefined)
        ? (data.splitChuyenKhoan + "," + data.splitTienMat)
        : tongTienTraKhach,
      data.hinhThucThanhToan || "Tiền mặt",
      chiNhanh,
      data.nguoiThucHien || "",
      data.ghiChu || "",
    ];

    var tmSheet = ss.getSheetByName(SHEET_NAMES.THU_MUA);
    var tmRow = appendRow(SHEET_NAMES.THU_MUA, rowData);
    rollbackActions.push(function() {
      try {
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
