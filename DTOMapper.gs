/**
 * ============================================================
 * VanTran Mobile — DTOMapper.gs
 * Thư viện khai báo DTO và map dữ liệu động (DRY + KISS)
 * ============================================================
 */

const DTO_DEFS = {
  DON_HANG_FULL: {
    maDH: ["MA_DH", "string"],
    ngayBan: ["NGAY_BAN", "date"],
    maKH: ["MA_KH", "string"],
    tenKH: ["TEN_KH", "string"],
    maSP: ["MA_SP", "string"],
    tenSP: ["TEN_SP", "string"],
    nguonSP: ["NGUON_SP", "string"],
    thuongHieu: ["THUONG_HIEU", "string"],
    soLuong: ["SO_LUONG", "number"],
    donGia: ["DON_GIA", "number"],
    thanhTien: ["THANH_TIEN", "number"],
    hinhThucBan: ["HINH_THUC_BAN", "string"],
    hinhThucThanhToan: ["HINH_THUC_TT", "string"],
    nguoiBan: ["NGUOI_BAN", "string"],
    tenNguoiBan: ["TEN_NGUOI_BAN", "string"],
    coQuyenXuatMay: ["QUYEN_XUAT", "string"],
    nguoiHoTro: ["NGUOI_HO_TRO", "string"],
    tenNguoiHoTro: ["TEN_NGUOI_HO_TRO", "string"],
    trangThai: ["TRANG_THAI", "string"],
    ghiChu: ["GHI_CHU", "string"],
    chiNhanh: ["CHI_NHANH", "string"],
    maQuaTang: ["MA_QUA_TANG", "string"],
    tenQuaTang: ["TEN_QUA_TANG", "string"],
    coNhanQua: ["CO_NHAN_QUA", "string", "✗"],
    tienGiamGia: ["TIEN_GIAM_GIA", "number"],
    imei: ["IMEI", "string"]
  }
};

/**
 * Chuyển đổi dòng thô hoặc object thô sang DTO theo định nghĩa
 * @param {Array|Object} row Dòng dữ liệu thô (mảng) hoặc Object thô từ mapRowToObject
 * @param {string} sheetName Tên sheet tương ứng
 * @param {string} defKey Khóa định nghĩa DTO trong DTO_DEFS
 * @return {Object} Đối tượng DTO sau khi đã chuẩn hóa kiểu dữ liệu
 */
function toDTO(row, sheetName, defKey) {
  const obj = Array.isArray(row) ? mapRowToObject(row, sheetName) : row;
  const def = DTO_DEFS[defKey];
  if (!def) throw new Error("Không tìm thấy định nghĩa DTO: " + defKey);
  
  const out = {};
  for (const field in def) {
    const [code, type, dft] = def[field];
    const raw = obj[code];
    
    if (type === "number") {
      out[field] = (raw === undefined || raw === null || raw === "") ? (dft ?? 0) : Number(raw);
      if (isNaN(out[field])) out[field] = dft ?? 0;
    } else if (type === "date") {
      out[field] = raw instanceof Date ? raw : (raw ? new Date(raw) : (dft ?? null));
    } else { // type === "string"
      out[field] = (raw === undefined || raw === null || raw === "") ? (dft ?? "") : String(raw);
    }
  }
  return out;
}
