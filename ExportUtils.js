/**
 * Xuất sheet Danh mục Thiết bị ra file CSV và gửi link tải về cho người dùng.
 */
function exportEquipmentSheetToCsv() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    if (!sheet) throw new Error(`Không tìm thấy sheet "${EQUIPMENT_SHEET_NAME}"`);
    const data = sheet.getDataRange().getValues();
    let csv = "";
    data.forEach(row => {
      csv += row.map(cell => `"${cell !== null ? cell.toString().replace(/"/g, '""') : ""}"`).join(",") + "\r\n";
    });
    const blob = Utilities.newBlob(csv, "text/csv", "Danh_muc_Thiet_bi.csv");
    const file = DriveApp.createFile(blob);
    ui.alert("Đã xuất file CSV. Bạn có thể tải tại: " + file.getUrl());
  } catch (e) {
    Logger.log(`Lỗi xuất CSV: ${e}`);
    ui.alert("Lỗi xuất CSV: " + e.message);
  }
}

function showQrSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('QrSidebar')
    .setTitle('Tạo QR code Thiết bị')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 *  Hàm mở sidebar hướng dẫn sử dụng
 */
function showUserGuideSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('UserGuideSidebar')
    .setTitle('Hướng dẫn sử dụng SHT')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Sidebar Liên hệ hỗ trợ kỹ thuật
 */
function showSupportContactSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('SupportContactSidebar')
    .setTitle('Liên hệ hỗ trợ kỹ thuật')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Hàm kiểm tra phiên bản code & nhật ký cập nhật
 */
function showVersionInfo() {
  var html = HtmlService.createHtmlOutputFromFile('VersionInfoSidebar')
    .setTitle('Phiên bản & Nhật ký cập nhật')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Hàm tạo phiếu giao việc đầu ca cho đội kỹ thuật (Phiên bản A4)
 * Được gọi từ menu tùy chỉnh
 */
function generateTechnicianWorkOrderSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const workOrderSheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
    
    if (!workOrderSheet) {
      throw new Error("Không tìm thấy sheet Phiếu Công Việc");
    }
    
    // Lọc ra các công việc chưa hoàn thành và đã giao cho kỹ thuật
    const dataRange = workOrderSheet.getDataRange();
    const data = dataRange.getValues();
    
    const filteredWorkOrders = [];
    
    for (let i = 1; i < data.length; i++) {
      const status = data[i][COL_PCV_TRANG_THAI - 1];
      
      // Chỉ lấy các công việc có trạng thái "Đã giao" hoặc "Đã lên kế hoạch"
      if (status === "Đã giao" || status === "Đang thực hiện" || status === "Đã lên kế hoạch") {
        filteredWorkOrders.push({
          id: data[i][COL_PCV_MA_PHIEU - 1],
          target: data[i][COL_PCV_DOI_TUONG - 1],
          target_name: data[i][COL_PCV_TEN_DOI_TUONG - 1],
          location: data[i][COL_PCV_VI_TRI - 1],
          workType: data[i][COL_PCV_LOAI_CV - 1],
          description: truncateText(data[i][COL_PCV_MO_TA_YC - 1], 200), // Giới hạn độ dài mô tả
          priority: data[i][COL_PCV_UU_TIEN - 1],
          deadline: data[i][COL_PCV_HAN_HT - 1]
        });
      }
    }
    
    // Sắp xếp theo ưu tiên: Cao -> Trung bình -> Thấp
    filteredWorkOrders.sort((a, b) => {
      const priorityOrder = { "Cao": 0, "Trung bình": 1, "Thấp": 2 };
      return priorityOrder[a.priority] - priorityOrder[b.priority];
    });
    
    // Tạo HTML output từ template tối ưu cho A4
    const htmlTemplate = HtmlService.createTemplateFromFile('TechnicianWorkOrderA4Template');
    htmlTemplate.workOrders = filteredWorkOrders;
    
    const htmlOutput = htmlTemplate.evaluate()
      .setTitle('Phiếu Công Việc Đầu Ca - Đội Kỹ Thuật')
      .setWidth(800)  // Đủ rộng để hiển thị A4
      .setHeight(600); // Đủ cao để xem toàn bộ nội dung
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Phiếu Công Việc Đầu Ca - Đội Kỹ Thuật');
    
  } catch (error) {
    Logger.log("Lỗi khi tạo phiếu công việc kỹ thuật: " + error);
    SpreadsheetApp.getUi().alert("Lỗi: " + error.message);
  }
}

/**
 * Hàm cắt ngắn văn bản nếu quá dài
 */
function truncateText(text, maxLength) {
  if (!text) return "";
  text = text.toString();
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength) + "...";
}



