/**
 * ============================================================
 * VanTran Mobile — NhanVien.gs
 * Quản lý nhân viên + quyền xuất máy
 * ============================================================
 */

/**
 * Lấy danh sách tất cả nhân viên đang làm
 *
 * @return {Object[]} Mảng objects nhân viên
 */
function getAllNhanVien() {
  const activeObjs = _getActiveNhanVienRows();
  
  return activeObjs.map(function(obj) {
    return {
      MaNV: String(obj.MA_NV),
      HoTen: String(obj.HO_TEN),
      SoDienThoai: String(obj.SO_DIEN_THOAI),
      Email: String(obj.EMAIL),
      VaiTro: String(obj.VAI_TRO),
      QuyenXuatMay: String(obj.QUYEN_XUAT),
      NgayVaoLam: obj.NGAY_VAO,
      TrangThai: String(obj.TRANG_THAI),
    };
  });
}

/**
 * Lấy danh sách NV để hiển thị trong dropdown (MaNV - HoTen)
 *
 * @return {Object[]} [{value: 'NV001', text: 'NV001 - Nguyễn Văn A'}, ...]
 */
function getNhanVienDropdown() {
  const activeObjs = _getActiveNhanVienRows();
  
  return activeObjs.map(function(obj) {
    return {
      value: String(obj.MA_NV),
      text: obj.MA_NV + " - " + obj.HO_TEN,
      vaiTro: String(obj.VAI_TRO),
      quyenXuatMay: String(obj.QUYEN_XUAT) === "✓",
    };
  });
}

/**
 * Thêm nhân viên mới
 *
 * @param {Object} data - {hoTen, soDienThoai, email, vaiTro, quyenXuatMay, ngayVaoLam, chiNhanh}
 * @return {string} Mã NV mới
 */
function addNhanVien(data) {
  const maNV = generateId("NV", SHEET_NAMES.NHAN_VIEN);

  const rowData = [];
  rowData[COL_NV.MA_NV - 1] = maNV;
  rowData[COL_NV.HO_TEN - 1] = data.hoTen || "";
  rowData[COL_NV.SO_DIEN_THOAI - 1] = data.soDienThoai || "";
  rowData[COL_NV.EMAIL - 1] = data.email || "";
  rowData[COL_NV.VAI_TRO - 1] = data.vaiTro || "Bán hàng";
  rowData[COL_NV.QUYEN_XUAT - 1] = data.quyenXuatMay ? "✓" : "✗";
  rowData[COL_NV.NGAY_VAO - 1] = data.ngayVaoLam ? new Date(data.ngayVaoLam) : new Date();
  rowData[COL_NV.TRANG_THAI - 1] = "Đang làm";

  appendRow(SHEET_NAMES.NHAN_VIEN, rowData);
  showToast("Đã thêm nhân viên: " + data.hoTen + " (" + maNV + ")");
  return maNV;
}

/**
 * Cập nhật thông tin nhân viên
 *
 * @param {string} maNV - Mã NV cần cập nhật
 * @param {Object} data - Dữ liệu cần cập nhật
 * @return {boolean} Thành công hay không
 */
function updateNhanVien(maNV, data) {
  const row = findRow(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNV);
  if (row === -1) {
    showAlert("❌ Lỗi", "Không tìm thấy nhân viên: " + maNV);
    return false;
  }

  const updates = {};
  if (data.hoTen !== undefined) updates[COL_NV.HO_TEN] = data.hoTen;
  if (data.soDienThoai !== undefined) updates[COL_NV.SO_DIEN_THOAI] = data.soDienThoai;
  if (data.email !== undefined) updates[COL_NV.EMAIL] = data.email;
  if (data.vaiTro !== undefined) updates[COL_NV.VAI_TRO] = data.vaiTro;
  if (data.quyenXuatMay !== undefined) updates[COL_NV.QUYEN_XUAT] = data.quyenXuatMay ? "✓" : "✗";
  if (data.trangThai !== undefined) updates[COL_NV.TRANG_THAI] = data.trangThai;

  saveRowData(SHEET_NAMES.NHAN_VIEN, row, updates);

  showToast("Đã cập nhật NV: " + maNV);
  return true;
}

/**
 * Tìm nhân viên theo mã
 *
 * @param {string} maNV - Mã NV
 * @return {Object|null} Thông tin NV
 */
function getNhanVienByMa(maNV) {
  const result = lookupMultipleValues(
    SHEET_NAMES.NHAN_VIEN,
    COL_NV.MA_NV,
    maNV,
    [
      COL_NV.MA_NV,
      COL_NV.HO_TEN,
      COL_NV.SO_DIEN_THOAI,
      COL_NV.EMAIL,
      COL_NV.VAI_TRO,
      COL_NV.QUYEN_XUAT,
      COL_NV.NGAY_VAO,
      COL_NV.TRANG_THAI
    ],
  );
  if (!result) return null;

  const obj = {
    MaNV: String(result[COL_NV.MA_NV]),
    HoTen: String(result[COL_NV.HO_TEN]),
    SoDienThoai: String(result[COL_NV.SO_DIEN_THOAI]),
    Email: String(result[COL_NV.EMAIL]),
    VaiTro: String(result[COL_NV.VAI_TRO]),
    QuyenXuatMay: String(result[COL_NV.QUYEN_XUAT]),
    NgayVaoLam: result[COL_NV.NGAY_VAO],
    TrangThai: String(result[COL_NV.TRANG_THAI]),
  };

  return obj;
}

/**
 * Kiểm tra NV có quyền xuất máy không
 *
 * @param {string} maNV - Mã NV
 * @return {boolean}
 */
function kiemTraQuyenXuatMay(maNV) {
  const quyen = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNV, COL_NV.QUYEN_XUAT);
  return String(quyen) === "✓";
}

/**
 * Lấy danh sách nhân viên đang hoạt động dưới dạng mapped objects
 * @private
 */
function _getActiveNhanVienRows() {
  const data = getAllData(SHEET_NAMES.NHAN_VIEN);
  
  return data.map(function(row) {
    return mapRowToObject(row, SHEET_NAMES.NHAN_VIEN);
  }).filter(function (obj) {
    return (
      obj.MA_NV &&
      String(obj.MA_NV).trim() !== "" &&
      String(obj.TRANG_THAI) !== "Nghỉ việc"
    );
  });
}

/**
 * Lấy tên nhân viên từ mã NV
 * 
 * @param {string} maNV Mã nhân viên cần tìm
 * @return {string} Tên nhân viên hoặc chuỗi rỗng
 */
function getNhanVienName(maNV) {
  if (!maNV) return "";
  return lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNV, COL_NV.HO_TEN) || "";
}
