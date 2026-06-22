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
  var data = getAllData(SHEET_NAMES.NHAN_VIEN);
  var result = [];

  data.forEach(function (row) {
    if (
      row[0] &&
      String(row[0]).trim() !== "" &&
      String(row[7]) !== "Nghỉ việc"
    ) {
      result.push({
        MaNV: String(row[0]),
        HoTen: String(row[1]),
        SoDienThoai: String(row[2]),
        Email: String(row[3]),
        VaiTro: String(row[4]),
        QuyenXuatMay: String(row[5]),
        NgayVaoLam: row[6],
        TrangThai: String(row[7]),
      });
    }
  });

  return result;
}

/**
 * Lấy danh sách NV để hiển thị trong dropdown (MaNV - HoTen)
 *
 * @return {Object[]} [{value: 'NV001', text: 'NV001 - Nguyễn Văn A'}, ...]
 */
function getNhanVienDropdown() {
  var data = getAllData(SHEET_NAMES.NHAN_VIEN);
  var result = [];

  data.forEach(function (row) {
    if (
      row[0] &&
      String(row[0]).trim() !== "" &&
      String(row[7]) !== "Nghỉ việc"
    ) {
      result.push({
        value: String(row[0]),
        text: row[0] + " - " + row[1],
        vaiTro: String(row[4]),
        quyenXuatMay: String(row[5]) === "✓",
      });
    }
  });

  return result;
}

/**
 * Thêm nhân viên mới
 *
 * @param {Object} data - {hoTen, soDienThoai, email, vaiTro, quyenXuatMay, ngayVaoLam, chiNhanh}
 * @return {string} Mã NV mới
 */
function addNhanVien(data) {
  var maNV = generateId("NV", SHEET_NAMES.NHAN_VIEN);

  var rowData = [
    maNV,
    data.hoTen || "",
    data.soDienThoai || "",
    data.email || "",
    data.vaiTro || "Bán hàng",
    data.quyenXuatMay ? "✓" : "✗",
    data.ngayVaoLam ? new Date(data.ngayVaoLam) : new Date(),
    "Đang làm",
  ];

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
  var row = findRow(SHEET_NAMES.NHAN_VIEN, 1, maNV);
  if (row === -1) {
    showAlert("❌ Lỗi", "Không tìm thấy nhân viên: " + maNV);
    return false;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.NHAN_VIEN);

  if (data.hoTen !== undefined) sheet.getRange(row, 2).setValue(data.hoTen);
  if (data.soDienThoai !== undefined)
    sheet.getRange(row, 3).setValue(data.soDienThoai);
  if (data.email !== undefined) sheet.getRange(row, 4).setValue(data.email);
  if (data.vaiTro !== undefined) sheet.getRange(row, 5).setValue(data.vaiTro);
  if (data.quyenXuatMay !== undefined)
    sheet.getRange(row, 6).setValue(data.quyenXuatMay ? "✓" : "✗");
  if (data.trangThai !== undefined)
    sheet.getRange(row, 8).setValue(data.trangThai);

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
  var result = lookupMultipleValues(
    SHEET_NAMES.NHAN_VIEN,
    1,
    maNV,
    [1, 2, 3, 4, 5, 6, 7, 8],
  );
  if (!result) return null;

  var obj = {
    MaNV: String(result[1]),
    HoTen: String(result[2]),
    SoDienThoai: String(result[3]),
    Email: String(result[4]),
    VaiTro: String(result[5]),
    QuyenXuatMay: String(result[6]),
    NgayVaoLam: result[7],
    TrangThai: String(result[8]),
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
  var quyen = lookupValue(SHEET_NAMES.NHAN_VIEN, 1, maNV, 6);
  return String(quyen) === "✓";
}
