/**
 * [HỖ TRỢ NHẬP LIỆU NHANH]
 * Tạo nhanh 1 dòng thiết bị mẫu ở cuối sheet "Danh mục Thiết bị".
 * Dùng để test quy trình tạo mã, cập nhật mua hàng, v.v.
 * Gợi ý: Sau khi tạo, chọn dòng và chạy "Tạo Mã & Xử lý Dòng TB Mới".
 */
function insertSampleEquipmentRow() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    if (!sheet) throw new Error(`Không tìm thấy sheet "${EQUIPMENT_SHEET_NAME}"`);
    // Tạo dòng mẫu với các trường cơ bản
    const sampleRow = [];
    sampleRow[COL_EQUIP_ID - 1] = ""; // Để trống để script tạo mã tự động
    sampleRow[COL_EQUIP_NAME - 1] = "TB Mẫu nhập nhanh";
    sampleRow[COL_EQUIP_TYPE - 1] = "QUATDL"; // Sửa lại cho đúng mã loại có trong cấu hình
    sampleRow[COL_EQUIP_PARENT_ID - 1] = ""; // Không có cha
    sampleRow[COL_EQUIP_PURCHASE_ID - 1] = ""; // Không có lô mua hàng
    sampleRow[COL_EQUIP_LOCATION - 1] = "KHO"; // Sửa lại cho đúng mã vị trí có trong cấu hình
    sampleRow[COL_EQUIP_STATUS - 1] = "Đang hoạt động";
    // ...bổ sung trường nếu cần
    sheet.appendRow(sampleRow);
    ui.alert('Đã tạo 1 dòng thiết bị mẫu ở cuối sheet. Hãy chọn dòng đó và chạy "Tạo Mã & Xử lý Dòng TB Mới".');
  } catch (e) {
    Logger.log(`Lỗi tạo thiết bị mẫu: ${e}`);
    ui.alert("Lỗi tạo thiết bị mẫu: " + e.message);
  }
}

/**
 * [HỖ TRỢ NHẬP LIỆU NHANH]
 * Tạo nhanh 1 dòng phiếu công việc mẫu ở cuối sheet "Phiếu Công Việc".
 * Gợi ý: Sau khi tạo, chọn dòng và chỉnh sửa bổ sung nếu cần.
 */
function insertSampleWorkOrderRow() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
    if (!sheet) throw new Error(`Không tìm thấy sheet "${SHEET_PHIEU_CONG_VIEC}"`);
    const sampleRow = [];
    sampleRow[COL_PCV_MA_PHIEU - 1] = ""; // Để trống để script tạo mã
    sampleRow[COL_PCV_NGAY_TAO - 1] = new Date();
    sampleRow[COL_PCV_NGUOI_TAO - 1] = Session.getActiveUser().getEmail();
    sampleRow[COL_PCV_DOI_TUONG - 1] = "QUATDL-001"; // Sửa lại cho đúng mã thiết bị thực tế
    sampleRow[COL_PCV_LOAI_CV - 1] = "Bảo trì Định kỳ";
    sampleRow[COL_PCV_TRANG_THAI - 1] = "Đã lên kế hoạch";
    sampleRow[COL_PCV_MO_TA_YC - 1] = "Phiếu mẫu kiểm thử quy trình";
    sheet.appendRow(sampleRow);
    ui.alert('Đã tạo 1 dòng phiếu công việc mẫu ở cuối sheet. Hãy chọn dòng đó và chỉnh sửa bổ sung nếu cần.');
  } catch (e) {
    Logger.log(`Lỗi tạo phiếu công việc mẫu: ${e}`);
    ui.alert("Lỗi tạo phiếu công việc mẫu: " + e.message);
  }
}

/**
 * [HỖ TRỢ NHẬP LIỆU NHANH]
 * Tạo nhanh 1 dòng lịch sử mẫu ở cuối sheet "Lịch sử Bảo trì / Sửa chữa".
 * Gợi ý: Sau khi tạo, chọn dòng và chạy "🆔 Tạo ID & Xử lý Dòng Lịch sử Mới".
 */
function insertSampleHistoryRow() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!sheet) throw new Error(`Không tìm thấy sheet "${HISTORY_SHEET_NAME}"`);
    const sampleRow = [];
    sampleRow[COL_HISTORY_ID - 1] = ""; // Để trống để script tạo mã
    sampleRow[COL_HISTORY_TARGET_CODE - 1] = "QUATDL-001"; // Sửa lại cho đúng mã thiết bị thực tế
    sampleRow[COL_HISTORY_EXEC_DATE - 1] = new Date();
    sampleRow[COL_HISTORY_WORK_TYPE - 1] = "Bảo trì Định kỳ";
    sampleRow[COL_HISTORY_DESCRIPTION - 1] = "Lịch sử bảo trì mẫu để test quy trình";
    sampleRow[COL_HISTORY_STATUS - 1] = "Hoàn thành";
    sheet.appendRow(sampleRow);
    ui.alert('Đã tạo 1 dòng lịch sử mẫu ở cuối sheet. Hãy chọn dòng đó và chạy "🆔 Tạo ID & Xử lý Dòng Lịch sử Mới".');
  } catch (e) {
    Logger.log(`Lỗi tạo lịch sử mẫu: ${e}`);
    ui.alert("Lỗi tạo lịch sử mẫu: " + e.message);
  }
}
