/**
 * ============================================================
 * VanTran Mobile — BaoCaoNgayCollector.gs
 * Chức năng thu thập dữ liệu giao dịch từ các phân hệ để làm báo cáo ngày
 * ============================================================
 */

/**
 * Thu thập giao dịch ĐƠN HÀNG (DON_HANG)
 * @private
 */
function _collectDonHang(targetDate, branchMetrics, fallbackBranch, transactions, ss) {
  const orders = getAllData(SHEET_NAMES.DON_HANG);

  orders.forEach(function (row) {
    const ngayBan = row[COL_DH.NGAY_BAN - 1];
    if (ngayBan instanceof Date && _isSameDay(ngayBan, targetDate)) {
      const status = String(row[COL_DH.TRANG_THAI - 1]);
      if (isCancelStatus(status) || status === ORDER_STATUS.EXCHANGED) return;

      const maDH = String(row[COL_DH.MA_DH - 1]);
      const tenKH = String(row[COL_DH.TEN_KH - 1]);
      const maSP = String(row[COL_DH.MA_SP - 1]);
      const tenSP = String(row[COL_DH.TEN_SP - 1]);
      const nguonSP = String(row[COL_DH.NGUON_SP - 1]);
      const sl = Number(row[COL_DH.SO_LUONG - 1]) || 0;
      const donGia = Number(row[COL_DH.DON_GIA - 1]) || 0;
      const thanhTien = parseAmountVal(row[COL_DH.THANH_TIEN - 1]);

      const tienThuMua = Number(row[COL_DH.TIEN_THU_MUA - 1]) || 0;
      const netRevenue = thanhTien + tienThuMua;

      const hinhThucBan = String(row[COL_DH.HINH_THUC_BAN - 1]);
      const hinhThucTT = String(row[COL_DH.HINH_THUC_TT - 1]);
      const branch = String(row[COL_DH.CHI_NHANH - 1] || "").trim();
      const maQua = String(row[COL_DH.MA_QUA_TANG - 1] || "");
      const coNhanQua = String(row[COL_DH.CO_NHAN_QUA - 1] || "✗");

      // Tra cứu giá nhập sản phẩm chính
      let giaNhapSP = 0;
      if (nguonSP === "Điện thoại") {
        giaNhapSP =
          Number(
            lookupValue(
              SHEET_NAMES.DIEN_THOAI,
              COL_DT.IMEI,
              maSP,
              COL_DT.GIA_NHAP,
            ),
          ) ||
          Number(
            lookupValue(
              SHEET_NAMES.DIEN_THOAI,
              COL_DT.MA_DT,
              maSP,
              COL_DT.GIA_NHAP,
            ),
          ) ||
          0;
      } else {
        const pkRow = findPhuKienRow(maSP, branch);
        if (pkRow !== -1) {
          giaNhapSP =
            Number(
              ss
                .getSheetByName(SHEET_NAMES.PHU_KIEN)
                .getRange(pkRow, COL_PK.GIA_NHAP)
                .getValue(),
            ) || 0;
        }
      }
      const costSP = giaNhapSP * sl;

      // Tra cứu giá nhập quà tặng (nếu nhận quà)
      let costQua = 0;
      if (coNhanQua === "✓" && maQua) {
        const maQuaList = String(maQua).split(",");
        for (let i = 0; i < maQuaList.length; i++) {
          const code = maQuaList[i].trim();
          if (!code) continue;
          const quaRow = findPhuKienRow(code, branch);
          if (quaRow !== -1) {
            costQua +=
              Number(
                ss
                  .getSheetByName(SHEET_NAMES.PHU_KIEN)
                  .getRange(quaRow, COL_PK.GIA_NHAP)
                  .getValue(),
              ) || 0;
          }
        }
      }

      const loiNhuan = netRevenue - costSP - costQua;

      // Tính dòng tiền
      let thuTM = 0;
      let thuCK = 0;
      let congNo = 0;

      const cellTM = row[COL_DH.TIEN_MAT - 1];
      const cellCK = row[COL_DH.CHUYEN_KHOAN - 1];

      // Check if we have new columns populated
      const hasNewColumns =
        (cellTM !== undefined && cellTM !== "") ||
        (cellCK !== undefined && cellCK !== "");

      if (hasNewColumns) {
        thuTM = Number(cellTM) || 0;
        thuCK = Number(cellCK) || 0;

        if (hinhThucBan === "Trả góp") {
          // Nợ (Công nợ) = Thành tiền đơn hàng - Số tiền trả trước đã thu (thuTM + thuCK)
          congNo += Math.max(0, thanhTien - (thuTM + thuCK));
        }
      } else {
        // Fallback cho dữ liệu cũ (chưa chia cột)
        let amountCollected = thanhTien;
        let rawAmount = row[COL_DH.THANH_TIEN - 1];

        if (hinhThucBan === "Trả góp") {
          const traTruocRaw = lookupValue(
            SHEET_NAMES.TRA_GOP,
            COL_TG.MA_DH,
            maDH,
            COL_TG.TRA_TRUOC,
          );
          const traTruoc = parseAmountVal(traTruocRaw);
          const conLai = Math.max(0, thanhTien - traTruoc);

          congNo += conLai;
          amountCollected = traTruoc;
          rawAmount = traTruocRaw;
        }

        const hybrid = parseHybridAmount(
          hinhThucTT === "Hỗn hợp" ? rawAmount : "",
        );
        if (hybrid) {
          thuTM += hybrid.tm;
          thuCK += hybrid.ck;
        } else {
          const payment = parseMixedPayment(hinhThucTT, amountCollected);
          thuTM += payment.tm;
          thuCK += payment.ck;
        }
      }

      const brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = {
          doanhSo: 0,
          loiNhuan: 0,
          thuTM: 0,
          chiTM: 0,
          thuCK: 0,
          chiCK: 0,
        };
      }
      branchMetrics[brKey].doanhSo += netRevenue;
      branchMetrics[brKey].loiNhuan += loiNhuan;
      branchMetrics[brKey].thuTM += thuTM;
      branchMetrics[brKey].thuCK += thuCK;
      branchMetrics[brKey].congNoCTTC += congNo;

      transactions.push({
        maGD: maDH,
        time: ngayBan,
        loaiGD: "Đơn hàng",
        khachHang: tenKH,
        branch: branch,
        thuTM: thuTM,
        congNo: congNo,
        chiTM: 0,
        thuCK: thuCK,
        chiCK: 0,
        loiNhuan: loiNhuan,
        detail:
          tenSP + " (SL: " + sl + ", " + hinhThucBan + " - " + hinhThucTT + ")" + 
          (tienThuMua > 0 ? " [Trừ máy thu: -" + formatCurrency(tienThuMua) + "đ]" : ""),
      });
    }
  });
}

/**
 * Thu thập giao dịch DỊCH VỤ (DICH_VU)
 * @private
 */
function _collectDichVu(targetDate, branchMetrics, fallbackBranch, transactions) {
  const services = getAllData(SHEET_NAMES.DICH_VU);
  services.forEach(function (row) {
    const ngayGD = row[COL_DV.NGAY_GD - 1];
    if (ngayGD instanceof Date && _isSameDay(ngayGD, targetDate)) {
      const status = String(row[COL_DV.TRANG_THAI - 1]);
      if (isCancelStatus(status)) return;

      const maDV = String(row[COL_DV.MA_DV - 1]);
      const loaiDV = String(row[COL_DV.LOAI_DV - 1]);
      const tenKH = String(row[COL_DV.TEN_KH - 1]);
      const rawSoTien = row[COL_DV.SO_TIEN_GD - 1];
      const soTien = parseAmountVal(rawSoTien);
      const phi = Number(row[COL_DV.PHI_DV - 1]) || 0;
      const hinhThucTT = String(row[COL_DV.HINH_THUC_TT - 1]);
      const branch = String(row[COL_DV.CHI_NHANH - 1] || "").trim();

      let thuTM = 0;
      let chiTM = 0;
      let thuCK = 0;
      let chiCK = 0;
      const loiNhuan = phi;

      const cellTM = row[COL_DV.TIEN_MAT - 1];
      const cellCK = row[COL_DV.CHUYEN_KHOAN - 1];

      const hasNewColumns =
        (cellTM !== undefined && cellTM !== "" && Number(cellTM) !== 0) ||
        (cellCK !== undefined && cellCK !== "" && Number(cellCK) !== 0);

      if (hasNewColumns) {
        thuTM = Number(cellTM) || 0;
        thuCK = Number(cellCK) || 0;

        if (loaiDV === "Chuyển khoản hộ" || loaiDV === "Nạp thẻ điện thoại") {
          chiCK = soTien;
        } else if (loaiDV === "Rút tiền mặt") {
          chiTM = soTien;
        }
      } else {
        // Fallback cho dữ liệu cũ (chưa chia cột)
        const hybrid = parseHybridAmount(
          hinhThucTT === "Hỗn hợp" ? rawSoTien : "",
        );
        if (loaiDV === "Chuyển khoản hộ") {
          if (hybrid) {
            thuTM = hybrid.tm;
            thuCK = hybrid.ck;
            chiCK = soTien - phi;
          } else {
            const payment = parseMixedPayment(hinhThucTT, soTien + phi);
            thuTM = payment.tm;
            thuCK = payment.ck;
            chiCK = soTien;
          }
        } else if (loaiDV === "Rút tiền mặt") {
          thuCK = soTien + phi;
          chiTM = soTien;
        } else if (loaiDV === "Nạp thẻ điện thoại") {
          if (hybrid) {
            thuTM = hybrid.tm;
            thuCK = hybrid.ck;
            chiCK = soTien - phi;
          } else {
            const payment = parseMixedPayment(hinhThucTT, soTien + phi);
            thuTM = payment.tm;
            thuCK = payment.ck;
            chiCK = soTien;
          }
        }
      }

      const brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = {
          doanhSo: 0,
          loiNhuan: 0,
          thuTM: 0,
          chiTM: 0,
          thuCK: 0,
          chiCK: 0,
        };
      }
      branchMetrics[brKey].loiNhuan += loiNhuan;
      branchMetrics[brKey].thuTM += thuTM;
      branchMetrics[brKey].chiTM += chiTM;
      branchMetrics[brKey].thuCK += thuCK;
      branchMetrics[brKey].chiCK += chiCK;

      transactions.push({
        maGD: maDV,
        time: ngayGD,
        loaiGD: "Dịch vụ",
        khachHang: tenKH,
        branch: branch,
        thuTM: thuTM,
        chiTM: chiTM,
        thuCK: thuCK,
        chiCK: chiCK,
        loiNhuan: loiNhuan,
        detail:
          loaiDV +
          " (" +
          formatCurrency(soTien) +
          "đ, Phí: " +
          formatCurrency(phi) +
          "đ)",
      });
    }
  });
}

/**
 * Thu thập giao dịch TRẢ GÓP (LICH_SU_TRA_GOP - repayments)
 * @private
 */
function _collectRepayments(targetDate, branchMetrics, fallbackBranch, transactions) {
  const repayments = getAllData(SHEET_NAMES.LICH_SU_TRA_GOP);
  repayments.forEach(function (row) {
    const ngayTra = row[COL_LSTG.NGAY_THUC_TRA - 1];
    const status = String(row[COL_LSTG.TRANG_THAI - 1]);
    if (
      ngayTra instanceof Date &&
      _isSameDay(ngayTra, targetDate) &&
      status === "Đã trả"
    ) {
      const maLS = String(row[COL_LSTG.MA_LS - 1]);
      const maTG = String(row[COL_LSTG.MA_TG - 1]);
      const kySo = String(row[COL_LSTG.KY_SO - 1]);
      const rawSoTienDaTra = row[COL_LSTG.SO_TIEN_DA_TRA - 1];
      const soTienDaTra = parseAmountVal(rawSoTienDaTra);
      const hinhThucTT = String(row[COL_LSTG.HINH_THUC_TT - 1]);

      const maKH = String(
        lookupValue(SHEET_NAMES.TRA_GOP, COL_TG.MA_TG, maTG, COL_TG.MA_KH),
      );
      const tenKH =
        String(lookupValue(SHEET_NAMES.KHACH_HANG, 1, maKH, 2)) ||
        "Khách trả góp";
      const branch = String(
        lookupValue(
          SHEET_NAMES.TRA_GOP,
          COL_TG.MA_TG,
          maTG,
          COL_TG.CHI_NHANH,
        ) || "",
      ).trim();

      let thuTM = 0;
      let thuCK = 0;

      const cellTM = row[COL_LSTG.TIEN_MAT - 1];
      const cellCK = row[COL_LSTG.CHUYEN_KHOAN - 1];

      const hasNewColumns =
        (cellTM !== undefined && cellTM !== "" && Number(cellTM) !== 0) ||
        (cellCK !== undefined && cellCK !== "" && Number(cellCK) !== 0);

      if (hasNewColumns) {
        thuTM = Number(cellTM) || 0;
        thuCK = Number(cellCK) || 0;
      } else {
        // Fallback cho dữ liệu cũ (chưa chia cột)
        const hybrid = parseHybridAmount(
          hinhThucTT === "Hỗn hợp" ? rawSoTienDaTra : "",
        );
        if (hybrid) {
          thuTM = hybrid.tm;
          thuCK = hybrid.ck;
        } else {
          const payment = parseMixedPayment(hinhThucTT, soTienDaTra);
          thuTM = payment.tm;
          thuCK = payment.ck;
        }
      }

      const brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = {
          doanhSo: 0,
          loiNhuan: 0,
          thuTM: 0,
          chiTM: 0,
          thuCK: 0,
          chiCK: 0,
        };
      }
      branchMetrics[brKey].thuTM += thuTM;
      branchMetrics[brKey].thuCK += thuCK;

      transactions.push({
        maGD: maLS,
        time: ngayTra,
        loaiGD: "Thu trả góp",
        khachHang: tenKH,
        branch: branch,
        thuTM: thuTM,
        chiTM: 0,
        thuCK: thuCK,
        chiCK: 0,
        loiNhuan: 0,
        detail: "Thanh toán HĐ " + maTG + " - Kỳ " + kySo,
      });
    }
  });
}

/**
 * Thu thập giao dịch NHẬP KHO (NHAP_KHO)
 * @private
 */
function _collectNhapKho(targetDate, branchMetrics, fallbackBranch, transactions) {
  const imports = getAllData(SHEET_NAMES.NHAP_KHO);
  imports.forEach(function (row) {
    const ngayNhap = row[1];
    if (ngayNhap instanceof Date && _isSameDay(ngayNhap, targetDate)) {
      const maNK = String(row[0]);
      const nguonNhap = String(row[2]);
      const tenSP = String(row[4]);
      const sl = Number(row[5]) || 0;
      const thanhTien = Number(row[7]) || 0;
      const nhaCC = String(row[8] || "");
      const branch = String(row[10] || "").trim();

      const chiCK = thanhTien;

      const brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = {
          doanhSo: 0,
          loiNhuan: 0,
          thuTM: 0,
          chiTM: 0,
          thuCK: 0,
          chiCK: 0,
        };
      }
      branchMetrics[brKey].chiCK += chiCK;

      transactions.push({
        maGD: maNK,
        time: ngayNhap,
        loaiGD: "Nhập kho",
        khachHang: nhaCC || "Nhà cung cấp",
        branch: branch,
        thuTM: 0,
        chiTM: 0,
        thuCK: 0,
        chiCK: chiCK,
        loiNhuan: 0,
        detail: "Nhập " + nguonNhap + ": " + tenSP + " (SL: " + sl + ")",
      });
    }
  });
}

/**
 * Thu thập giao dịch ĐỔI TRẢ (DOI_TRA)
 * @private
 */
function _collectDoiTra(targetDate, branchMetrics, fallbackBranch, transactions, ss) {
  const returns = getAllData(SHEET_NAMES.DOI_TRA);
  returns.forEach(function (row) {
    const ngayDT = row[COL_DT_TRA.NGAY_DT - 1];
    if (ngayDT instanceof Date && _isSameDay(ngayDT, targetDate)) {
      const status = String(row[COL_DT_TRA.TRANG_THAI - 1]);
      if (isCancelStatus(status)) return;

      const maDT = String(row[COL_DT_TRA.MA_DT - 1]);
      const maDH = String(row[COL_DT_TRA.MA_DH - 1]);
      const tenKH = String(row[COL_DT_TRA.TEN_KH - 1]);
      const loaiGD = String(row[COL_DT_TRA.LOAI_GD - 1]);
      const tenSPTra = String(row[COL_DT_TRA.TEN_SP_TRA - 1]);
      const tenSPNhan = String(row[COL_DT_TRA.TEN_SP_NHAN - 1]);
      const rawHoanTien = row[COL_DT_TRA.TIEN_HOAN_TRA - 1];
      const hoanTien = parseAmountVal(rawHoanTien);
      const phi = Number(row[COL_DT_TRA.PHI_DOI_TRA - 1]) || 0;
      const hinhThucTT = String(row[COL_DT_TRA.HINH_THUC_TT - 1]);
      const branch = String(row[COL_DT_TRA.CHI_NHANH - 1] || "").trim();

      let thuTM = 0;
      let chiTM = 0;
      let thuCK = 0;
      let chiCK = 0;

      const originalOrder = getDonHangDetails(maDH);
      let costOriginal = 0;
      let priceOriginal = 0;
      if (originalOrder) {
        priceOriginal = parseAmountVal(originalOrder.thanhTien);
        let costSP = 0;
        if (originalOrder.nguonSP === "Điện thoại") {
          costSP =
            Number(
              lookupValue(
                SHEET_NAMES.DIEN_THOAI,
                COL_DT.IMEI,
                originalOrder.maSP,
                COL_DT.GIA_NHAP,
              ),
            ) ||
            Number(
              lookupValue(
                SHEET_NAMES.DIEN_THOAI,
                COL_DT.MA_DT,
                originalOrder.maSP,
                COL_DT.GIA_NHAP,
              ),
            ) ||
            0;
        } else {
          const pkRow = findPhuKienRow(
            originalOrder.maSP,
            originalOrder.chiNhanh,
          );
          if (pkRow !== -1) {
            costSP =
              Number(
                ss
                  .getSheetByName(SHEET_NAMES.PHU_KIEN)
                  .getRange(pkRow, COL_PK.GIA_NHAP)
                  .getValue(),
              ) || 0;
          }
        }
        costOriginal = costSP * originalOrder.soLuong;

        if (originalOrder.coNhanQua === "✓" && originalOrder.maQuaTang) {
          const maQuaOriginalList = String(originalOrder.maQuaTang).split(",");
          for (let i = 0; i < maQuaOriginalList.length; i++) {
            const code = maQuaOriginalList[i].trim();
            if (!code) continue;
            const quaRow = findPhuKienRow(code, originalOrder.chiNhanh);
            if (quaRow !== -1) {
              costOriginal +=
                Number(
                  ss
                    .getSheetByName(SHEET_NAMES.PHU_KIEN)
                    .getRange(quaRow, COL_PK.GIA_NHAP)
                    .getValue(),
                ) || 0;
            }
          }
        }
      }

      let lossDoanhSo = 0;
      let lossLoiNhuan = 0;

      const cellTM = row[COL_DT_TRA.TIEN_MAT - 1];
      const cellCK = row[COL_DT_TRA.CHUYEN_KHOAN - 1];

      const hasNewColumns =
        (cellTM !== undefined && cellTM !== "" && Number(cellTM) !== 0) ||
        (cellCK !== undefined && cellCK !== "" && Number(cellCK) !== 0);

      if (hasNewColumns) {
        const valTM = Number(cellTM) || 0;
        const valCK = Number(cellCK) || 0;

        if (hoanTien >= 0) {
          chiTM = valTM;
          chiCK = valCK;
        } else {
          thuTM = valTM;
          thuCK = valCK;
        }
      } else {
        // Fallback cho dữ liệu cũ (chưa chia cột)
        const hybrid = parseHybridAmount(
          hinhThucTT === "Hỗn hợp" ? rawHoanTien : "",
        );
        if (loaiGD === "Trả máy") {
          if (hybrid) {
            chiTM = hybrid.tm;
            chiCK = hybrid.ck;
          } else {
            const payment = parseMixedPayment(hinhThucTT, hoanTien);
            chiTM = payment.tm;
            chiCK = payment.ck;
          }
        } else {
          if (hoanTien >= 0) {
            if (hybrid) {
              chiTM = hybrid.tm;
              chiCK = hybrid.ck;
            } else {
              const payment = parseMixedPayment(hinhThucTT, hoanTien);
              chiTM = payment.tm;
              chiCK = payment.ck;
            }
          } else {
            if (hybrid) {
              thuTM = hybrid.tm;
              thuCK = hybrid.ck;
            } else {
              const payment = parseMixedPayment(hinhThucTT, Math.abs(hoanTien));
              thuTM = payment.tm;
              thuCK = payment.ck;
            }
          }
        }
      }

      lossDoanhSo = -priceOriginal;
      lossLoiNhuan = -(priceOriginal - costOriginal);
      lossLoiNhuan += phi;

      const brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = {
          doanhSo: 0,
          loiNhuan: 0,
          thuTM: 0,
          chiTM: 0,
          thuCK: 0,
          chiCK: 0,
        };
      }
      branchMetrics[brKey].doanhSo += lossDoanhSo;
      branchMetrics[brKey].loiNhuan += lossLoiNhuan;
      branchMetrics[brKey].thuTM += thuTM;
      branchMetrics[brKey].chiTM += chiTM;
      branchMetrics[brKey].thuCK += thuCK;
      branchMetrics[brKey].chiCK += chiCK;

      transactions.push({
        maGD: maDT,
        time: ngayDT,
        loaiGD: "Đổi trả",
        khachHang: tenKH,
        branch: branch,
        thuTM: thuTM,
        chiTM: chiTM,
        thuCK: thuCK,
        chiCK: chiCK,
        loiNhuan: lossLoiNhuan,
        detail:
          loaiGD +
          ": " +
          tenSPTra +
          (loaiGD === "Đổi máy" ? " ➔ " + tenSPNhan : "") +
          " (Phí: " +
          formatCurrency(phi) +
          "đ)",
      });
    }
  });
}

/**
 * Thu thập giao dịch THU MUA (THU_MUA)
 * @private
 */
function _collectThuMua(targetDate, branchMetrics, fallbackBranch, transactions) {
  const buybacks = getAllData(SHEET_NAMES.THU_MUA);
  buybacks.forEach(function (row) {
    const ngayTM = row[COL_TM.NGAY_TM - 1];
    if (ngayTM instanceof Date && _isSameDay(ngayTM, targetDate)) {
      const status = String(row[COL_TM.TRANG_THAI - 1]);
      if (isCancelStatus(status)) return;

      const maTM = String(row[COL_TM.MA_TM - 1]);
      const tenKH = String(row[COL_TM.TEN_KH - 1]);
      const tenSPThu = String(row[COL_TM.TEN_SP_THU - 1]);
      const rawTongTienTra = row[COL_TM.TONG_TIEN_TRA - 1];
      const tongTienTra = parseAmountVal(rawTongTienTra);
      const hinhThucTT = String(row[COL_TM.HINH_THUC_TT - 1]);
      const branch = String(row[COL_TM.CHI_NHANH - 1] || "").trim();

      let chiTM = 0;
      let chiCK = 0;

      if (hinhThucTT === "Trừ vào đơn mới") {
        chiTM = 0;
        chiCK = 0;
      } else {
        const cellTM = row[COL_TM.TIEN_MAT - 1];
        const cellCK = row[COL_TM.CHUYEN_KHOAN - 1];

        const hasNewColumns =
          (cellTM !== undefined && cellTM !== "") ||
          (cellCK !== undefined && cellCK !== "");

        if (hasNewColumns) {
          chiTM = Number(cellTM) || 0;
          chiCK = Number(cellCK) || 0;
        } else {
          // Fallback cho dữ liệu cũ (chưa chia cột)
          const hybrid = parseHybridAmount(
            hinhThucTT === "Hỗn hợp" ? rawTongTienTra : "",
          );
          if (hybrid) {
            chiTM = hybrid.tm;
            chiCK = hybrid.ck;
          } else {
            const payment = parseMixedPayment(hinhThucTT, tongTienTra);
            chiTM = payment.tm;
            chiCK = payment.ck;
          }
        }
      }

      const brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = {
          doanhSo: 0,
          loiNhuan: 0,
          thuTM: 0,
          chiTM: 0,
          thuCK: 0,
          chiCK: 0,
        };
      }
      branchMetrics[brKey].chiTM += chiTM;
      branchMetrics[brKey].chiCK += chiCK;

      transactions.push({
        maGD: maTM,
        time: ngayTM,
        loaiGD: "Thu mua",
        khachHang: tenKH,
        branch: branch,
        thuTM: 0,
        chiTM: chiTM,
        thuCK: 0,
        chiCK: chiCK,
        loiNhuan: 0,
        detail:
          "Thu mua máy cũ: " +
          tenSPThu +
          " (Giá: " +
          formatCurrency(tongTienTra) +
          "đ)",
      });
    }
  });
}

/**
 * Thu thập giao dịch BẢO HÀNH & SỬA CHỮA (BAO_HANH)
 * @private
 */
function _collectBaoHanh(targetDate, branchMetrics, fallbackBranch, transactions) {
  const repairs = getAllData(SHEET_NAMES.BAO_HANH);
  repairs.forEach(function (row) {
    const ngayNhan = row[COL_BH.NGAY_NHAN - 1];
    if (ngayNhan instanceof Date && _isSameDay(ngayNhan, targetDate)) {
      const status = String(row[COL_BH.TRANG_THAI - 1]);
      if (isCancelStatus(status)) return;

      const maBH = String(row[COL_BH.MA_BH - 1]);
      const tenKH = String(row[COL_BH.TEN_KH - 1]);
      const tenSP = String(row[COL_BH.TEN_SP - 1]);
      const loaiDichVu = String(row[COL_BH.LOAI_DICH_VU - 1]);
      const rawPhiSuaChua = row[COL_BH.PHI_SUA_CHUA - 1];
      const phiSuaChua = parseAmountVal(rawPhiSuaChua);
      const branch = String(row[COL_BH.CHI_NHANH - 1] || "").trim();

      let thuTM = 0;
      let chiTM = 0;
      let thuCK = 0;
      let chiCK = 0;

      const cellTM = row[COL_BH.TIEN_MAT - 1];
      const cellCK = row[COL_BH.CHUYEN_KHOAN - 1];

      const hasNewColumns =
        (cellTM !== undefined && cellTM !== "" && Number(cellTM) !== 0) ||
        (cellCK !== undefined && cellCK !== "" && Number(cellCK) !== 0);

      if (hasNewColumns) {
        thuTM = Number(cellTM) || 0;
        thuCK = Number(cellCK) || 0;
      } else {
        const hinhThucTT = String(row[COL_BH.HINH_THUC_TT - 1]);
        const hybrid = parseHybridAmount(
          hinhThucTT === "Hỗn hợp" ? rawPhiSuaChua : "",
        );
        if (hybrid) {
          thuTM = hybrid.tm;
          thuCK = hybrid.ck;
        } else {
          const payment = parseMixedPayment(hinhThucTT, phiSuaChua);
          thuTM = payment.tm;
          thuCK = payment.ck;
        }
      }

      const loiNhuan = phiSuaChua;

      const brKey = branch || fallbackBranch;
      if (!branchMetrics[brKey]) {
        branchMetrics[brKey] = {
          doanhSo: 0,
          loiNhuan: 0,
          thuTM: 0,
          chiTM: 0,
          thuCK: 0,
          chiCK: 0,
        };
      }
      branchMetrics[brKey].thuTM += thuTM;
      branchMetrics[brKey].thuCK += thuCK;
      branchMetrics[brKey].loiNhuan += loiNhuan;
      branchMetrics[brKey].doanhSo += phiSuaChua;

      transactions.push({
        maGD: maBH,
        time: ngayNhan,
        loaiGD: loaiDichVu || "Sửa chữa",
        khachHang: tenKH,
        branch: branch,
        thuTM: thuTM,
        chiTM: chiTM,
        thuCK: thuCK,
        chiCK: chiCK,
        loiNhuan: loiNhuan,
        detail:
          (loaiDichVu || "Sửa chữa") +
          " máy " +
          tenSP +
          " (Phí: " +
          formatCurrency(phiSuaChua) +
          "đ)",
      });
    }
  });
}

/**
 * Kiểm tra xem hai ngày có cùng ngày, tháng, năm hay không
 * @private
 */
function _isSameDay(d1, d2) {
  if (!(d1 instanceof Date) || !(d2 instanceof Date)) return false;
  return (
    d1.getDate() === d2.getDate() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getFullYear() === d2.getFullYear()
  );
}
