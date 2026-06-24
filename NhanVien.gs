/**
 * ============================================================
 * VanTran Mobile — NhanVien.gs
 * Quản lý nhân viên + quyền xuất máy
 * ============================================================
 */

/**
 * Lấy index cột (0-indexed) cho Nhân viên
 * @private
 */
function _getNhanVienIndices() {
  initializeColumnEnums();
  return {
    maNV: COL_NV.MA_NV - 1,
    hoTen: COL_NV.HO_TEN - 1,
    soDienThoai: COL_NV.SO_DIEN_THOAI - 1,
    email: COL_NV.EMAIL - 1,
    vaiTro: COL_NV.VAI_TRO - 1,
    quyenXuat: COL_NV.QUYEN_XUAT - 1,
    ngayVao: COL_NV.NGAY_VAO - 1,
    trangThai: COL_NV.TRANG_THAI - 1,
  };
}

/**
 * Lấy danh sách tất cả nhân viên đang làm
 *
 * @return {Object[]} Mảng objects nhân viên
 */
function getAllNhanVien() {
  var activeRows = _getActiveNhanVienRows();
  var c = _getNhanVienIndices();
  
  return activeRows.map(function(row) {
    return {
      MaNV: String(row[c.maNV]),
      HoTen: String(row[c.hoTen]),
      SoDienThoai: String(row[c.soDienThoai]),
      Email: String(row[c.email]),
      VaiTro: String(row[c.vaiTro]),
      QuyenXuatMay: String(row[c.quyenXuat]),
      NgayVaoLam: row[c.ngayVao],
      TrangThai: String(row[c.trangThai]),
    };
  });
}

/**
 * Lấy danh sách NV để hiển thị trong dropdown (MaNV - HoTen)
 *
 * @return {Object[]} [{value: 'NV001', text: 'NV001 - Nguyễn Văn A'}, ...]
 */
function getNhanVienDropdown() {
  var activeRows = _getActiveNhanVienRows();
  var c = _getNhanVienIndices();
  
  return activeRows.map(function(row) {
    return {
      value: String(row[c.maNV]),
      text: row[c.maNV] + " - " + row[c.hoTen],
      vaiTro: String(row[c.vaiTro]),
      quyenXuatMay: String(row[c.quyenXuat]) === "✓",
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
  initializeColumnEnums();
  var maNV = generateId("NV", SHEET_NAMES.NHAN_VIEN);

  var rowData = [];
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
  initializeColumnEnums();
  var row = findRow(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNV);
  if (row === -1) {
    showAlert("❌ Lỗi", "Không tìm thấy nhân viên: " + maNV);
    return false;
  }

  var updates = {};
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
  initializeColumnEnums();
  var result = lookupMultipleValues(
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

  var obj = {
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
  initializeColumnEnums();
  var quyen = lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNV, COL_NV.QUYEN_XUAT);
  return String(quyen) === "✓";
}

/**
 * Lấy danh sách các dòng nhân viên đang hoạt động
 * @private
 */
function _getActiveNhanVienRows() {
  var data = getAllData(SHEET_NAMES.NHAN_VIEN);
  var c = _getNhanVienIndices();
  
  return data.filter(function (row) {
    if (row.length <= c.trangThai) return false;
    return (
      row[c.maNV] &&
      String(row[c.maNV]).trim() !== "" &&
      String(row[c.trangThai]) !== "Nghỉ việc"
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
  initializeColumnEnums();
  return lookupValue(SHEET_NAMES.NHAN_VIEN, COL_NV.MA_NV, maNV, COL_NV.HO_TEN) || "";
}
