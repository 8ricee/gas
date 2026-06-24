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
 *   chiNhanh, nguoiThucHien (MaNV), nguoiHoTro (MaNV), ghiChu,
 *   // Nếu Thu cũ đổi mới:
 *   maSP_Moi
 * }
 * @return {string} Mã thu mua mới
 */
function thucHienThuMua(data) {
  return withDocumentLock(function () {
    clearSheetCache();
    const maTM = generateId("TM", SHEET_NAMES.THU_MUA);
    const chiNhanh = data.chiNhanh;
    const loaiGD = data.loaiGiaoDich;

    if (!data.maKH || !loaiGD || !data.imei_Thu || !chiNhanh) {
      throw new Error(
        "Vui lòng nhập đầy đủ Mã khách hàng, IMEI máy thu mua, Loại giao dịch và Chi nhánh!",
      );
    }

    const tenKH = ensureKhachHangExists(data.maKH, data.tenKH);
    let giaThuMua = Number(data.giaThuMua) || 0;
    let tienHoTro = Number(data.tienHoTro) || 0;
    const tongTienTraKhach = giaThuMua + tienHoTro;

    const rollbackActions = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    try {
      // 1. Không cần tự động đưa máy thu mua vào kho Điện thoại theo yêu cầu của người dùng
      // Dữ liệu chỉ cần lưu lại ở sheet Thu mua.

      let maDH_Moi = "";

      // 2. Nếu là Thu cũ đổi mới -> Tạo đơn hàng mua máy mới
      if (loaiGD === "Thu cũ đổi mới") {
        const maSP_Moi = data.maSP_Moi;
        if (!maSP_Moi) {
          throw new Error("Vui lòng chọn máy mới để đổi lên đời!");
        }

        const giaBanMoi = Number(data.donGia_Moi) ||
          Number(
            lookupValue(
              SHEET_NAMES.DIEN_THOAI,
              COL_DT.MA_DT,
              maSP_Moi,
              COL_DT.GIA_BAN,
            ),
          ) || 0;

        const hinhThucBan = data.hinhThucBan || "Bán thẳng";
        giaThuMua = Number(data.giaThuMua) || 0;
        tienHoTro = Number(data.tienHoTro) || 0;
        const tienGiamGia = Number(data.tienGiamGia) || 0;

        // Tạo đơn hàng mua máy mới
        const donHangHTTT = data.hinhThucThanhToanPhu || "Tiền mặt";
        let donHangSplitTM = undefined;
        let donHangSplitCK = undefined;

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
          nguoiHoTro: data.nguoiHoTro,
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
            const dhSheet = ss.getSheetByName(SHEET_NAMES.DON_HANG);
            const dhRow = findRow(SHEET_NAMES.DON_HANG, COL_DH.MA_DH, maDH_Moi);
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
      const rowData = [];
      rowData[COL_TM.MA_TM - 1] = maTM;
      rowData[COL_TM.NGAY_TM - 1] = new Date();
      rowData[COL_TM.MA_KH - 1] = data.maKH;
      rowData[COL_TM.TEN_KH - 1] = tenKH;
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

      const splitResult = calculatePaymentSplit(data, tongTienTraKhach);
      rowData[COL_TM.TIEN_MAT - 1] = splitResult.tienMat;
      rowData[COL_TM.CHUYEN_KHOAN - 1] = splitResult.chuyenKhoan;

      const tmSheet = ss.getSheetByName(SHEET_NAMES.THU_MUA);
      const tmRow = appendRow(SHEET_NAMES.THU_MUA, rowData);
      rollbackActions.push(function () {
        try {
          const ssRollback = SpreadsheetApp.getActiveSpreadsheet();
          const tmSheetRollback = ssRollback.getSheetByName(SHEET_NAMES.THU_MUA);
          tmSheetRollback.deleteRow(tmRow);
          clearSheetCache(SHEET_NAMES.THU_MUA);
        } catch (err) {
          Logger.log("Rollback failed to delete buyback row: " + err.message);
        }
      });
    } catch (e) {
      for (let rIdx = rollbackActions.length - 1; rIdx >= 0; rIdx--) {
        rollbackActions[rIdx]();
      }
      throw e;
    }

    showToast("✅ Đã ghi nhận thu mua thành công: " + maTM);
    return maTM;
  });
}
