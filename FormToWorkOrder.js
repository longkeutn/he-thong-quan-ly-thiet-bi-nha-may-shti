/**
 * Trigger tự động khi có form submit (từ Google Form vào Google Sheets)
 * Xử lý dữ liệu từ form responses và tạo phiếu công việc trong sheet Phiếu Công Việc
 */
function onFormSubmit(e) {
  try {
    const responses = e.values;
    // Logger.log("Dữ liệu responses: " + JSON.stringify(responses));

    // Map theo đúng thứ tự cột bạn đã gửi
    const formData = {
      timestamp: new Date(responses[0]),                  // Dấu thời gian
      email: responses[1] || "",                          // Địa chỉ email (nếu có)
      targetCode: responses[2] || "",                     // Mã thiết bị/hệ thống
      faultDescription: responses[3] || "",               // Mô tả chi tiết lỗi/sự cố
      imageUrl: responses[4] || "",                       // Ảnh minh họa (link Google Drive)
      priority: responses[5] || "Trung bình",             // Mức độ khẩn cấp
      reporter: responses[6] || "",                       // Người báo hỏng
      phone: responses[7] || ""                           // Số điện thoại liên hệ
    };

    createWorkOrderFromFormData(formData);
    Logger.log("Đã xử lý thành công form submit.");
  } catch (error) {
    Logger.log(`Lỗi khi xử lý form: ${error}\nStack: ${error.stack}`);
    MailApp.sendEmail({
      to: EMAIL_ADMIN,
      subject: "Lỗi xử lý form báo hỏng thiết bị",
      body: "Có lỗi khi xử lý form báo hỏng: " + error.toString()
    });
  }
}

/**
 * Xử lý URL ảnh từ Google Form (cho phép nhiều ảnh)
 */
function processImageUrl(fileInfo) {
  if (!fileInfo) return "";
  try {
    if (Array.isArray(fileInfo) && fileInfo.length > 0) {
      return fileInfo.map(file => file.url).filter(Boolean).join(", ");
    }
    if (typeof fileInfo === 'string') return fileInfo;
    return "";
  } catch (e) {
    Logger.log(`Lỗi xử lý URL ảnh: ${e}`);
    return "";
  }
}

/**
 * Tạo phiếu công việc mới từ dữ liệu form báo hỏng
 * Đã loại bỏ việc điền dữ liệu vào cột D (Ngày YC/Phát sinh)
 * @param {Object} formData Dữ liệu từ form báo hỏng
 */
function createWorkOrderFromFormData(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const workOrderSheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
  if (!workOrderSheet) throw new Error(`Không tìm thấy sheet "${SHEET_PHIEU_CONG_VIEC}"`);

  // Tạo mã phiếu công việc trực tiếp
  const newWorkOrderId = generateWorkOrderId();

  // Chuẩn bị dữ liệu cho dòng mới (mảng đủ 23 phần tử cho đến cột W)
  const newRow = Array(COL_PCV_HINH_ANH).fill("");

  // Đặt mã phiếu công việc đã tạo vào cột A
  newRow[COL_PCV_MA_PHIEU - 1] = newWorkOrderId; // A: Mã Phiếu CV

  newRow[COL_PCV_NGAY_TAO - 1] = new Date(); // B: Ngày tạo
  newRow[COL_PCV_NGUOI_TAO - 1] = formData.email; // C: Người tạo (email từ form)

  // Bỏ đi phần điền dữ liệu vào cột D (Ngày YC/Phát sinh) để người dùng điền thủ công
  newRow[COL_PCV_NGAY_YC - 1] = ""; // D: Để trống

  newRow[COL_PCV_HAN_HT - 1] = ""; // E: Hạn hoàn thành
  newRow[COL_PCV_DOI_TUONG - 1] = formData.targetCode; // F: Đối tượng / Hệ thống
  newRow[COL_PCV_TEN_DOI_TUONG - 1] = ""; // G: Tên TB / Mô tả HT
  newRow[COL_PCV_VI_TRI - 1] = ""; // H: Vị trí
  newRow[COL_PCV_LOAI_CV - 1] = "Sửa chữa Đột xuất"; // I: Loại Công việc
  newRow[COL_PCV_TAN_SUAT_PM - 1] = ""; // J: Tần suất PM

  // K: Mô tả Yêu cầu / Vấn đề - ĐÃ BỎ TIMESTAMP
  let detailedDesc = formData.faultDescription;
  detailedDesc += "\n\n---\nNgười báo: " + formData.reporter;
  detailedDesc += "\nEmail: " + formData.email;
  if (formData.phone) detailedDesc += "\nSĐT: " + formData.phone;
  
  // Thêm nguồn nhập - form báo lỗi
  detailedDesc += "\n\nForm báo lỗi"; // Đánh dấu nguồn từ form
  
  newRow[COL_PCV_MO_TA_YC - 1] = detailedDesc;

  // L: Mức độ Ưu tiên
  let priority = "Trung bình";
  if (formData.priority) {
    if (formData.priority.match(/Cao/i)) priority = "Cao";
    else if (formData.priority.match(/Thấp/i)) priority = "Thấp";
  }
  newRow[COL_PCV_UU_TIEN - 1] = priority;

  newRow[COL_PCV_NGUOI_GIAO - 1] = ""; // M: Người/Nhóm được giao
  newRow[COL_PCV_TRANG_THAI - 1] = "Đã lên kế hoạch"; // N: Trạng thái Phiếu CV
  newRow[COL_PCV_CHI_TIET_NGOAI - 1] = ""; // O: Chi tiết ĐV Ngoài / Liên hệ
  newRow[COL_PCV_MO_TA_HT - 1] = ""; // P: Mô tả Hoàn thành / Kết quả
  newRow[COL_PCV_VAT_TU - 1] = ""; // Q: Vật tư sử dụng
  newRow[COL_PCV_NGAY_HT_THUC_TE - 1] = ""; // R: Ngày Hoàn thành Thực tế
  newRow[COL_PCV_TRANG_THAI_TB_SAU - 1] = ""; // S: Trạng thái TB sau HĐ
  newRow[COL_PCV_CHI_PHI - 1] = ""; // T: Chi phí (VND)
  newRow[COL_PCV_LINK_LS - 1] = ""; // U: Link Lịch sử Hoàn thành
  newRow[COL_PCV_GHI_CHU - 1] = "Tạo tự động từ Form Báo Hỏng."; // V: Ghi chú Phiếu CV
  newRow[COL_PCV_HINH_ANH - 1] = formData.imageUrl; // W: Hình ảnh từ form

  // Thêm dòng mới vào sheet
  workOrderSheet.appendRow(newRow);
  const newRowIndex = workOrderSheet.getLastRow();

  // Cập nhật thông tin thiết bị từ mã
  updateEquipmentDetailsSimple(newRowIndex, formData.targetCode, workOrderSheet);
}





/**
 * Phiên bản đơn giản của updateEquipmentDetails không thực hiện bất kỳ định dạng nào
 * Chỉ cập nhật giá trị cột G và H dựa trên mã thiết bị
 */
function updateEquipmentDetailsSimple(rowIndex, targetCode, sheet) {
  try {
    if (!targetCode) return;
    
    // Trích xuất mã thiết bị từ chuỗi đầy đủ
    let equipmentCode = targetCode;
    
    if (typeof targetCode === 'string' && targetCode.includes(" - ")) {
      equipmentCode = targetCode.split(" - ")[0].trim();
      Logger.log(`Đã trích xuất mã thiết bị "${equipmentCode}" từ chuỗi đầy đủ "${targetCode}"`);
    } else {
      Logger.log(`Sử dụng mã thiết bị trực tiếp: "${equipmentCode}"`);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipmentSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    const systemDefSheet = ss.getSheetByName(SHEET_DINH_NGHIA_HE_THONG);
    
    // Tạo map tra cứu thiết bị
    const equipmentMap = {};
    if (equipmentSheet && equipmentSheet.getLastRow() >= 2) {
      const lastEquipCol = Math.max(COL_EQUIP_ID, COL_EQUIP_NAME, COL_EQUIP_LOCATION);
      const equipData = equipmentSheet.getRange(2, 1, equipmentSheet.getLastRow() - 1, lastEquipCol).getValues();
      equipData.forEach(row => {
        const id = row[COL_EQUIP_ID - 1];
        if (id) {
          const idStr = id.toString().trim();
          if (idStr) {
            equipmentMap[idStr] = {
              name: row[COL_EQUIP_NAME - 1] || 'N/A',
              location: row[COL_EQUIP_LOCATION - 1] || 'N/A'
            };
          }
        }
      });
    }
    
    // Tạo map tra cứu hệ thống
    const systemMap = {};
    if (systemDefSheet && systemDefSheet.getLastRow() >= 2) {
      const systemData = systemDefSheet.getRange(2, COL_HT_MA, systemDefSheet.getLastRow() - 1, 2).getValues();
      systemData.forEach(row => {
        const code = row[COL_HT_MA - 1];
        if (code) {
          const codeStr = code.toString().trim();
          if (codeStr) {
            systemMap[codeStr] = row[COL_HT_MO_TA - 1] || "";
          }
        }
      });
    }
    
    Logger.log(`Đã tải ${Object.keys(equipmentMap).length} TB, ${Object.keys(systemMap).length} HT.`);
    
    // Tìm thông tin theo mã đã trích xuất
    let targetName = "";
    let targetLocation = "";

    if (equipmentMap[equipmentCode]) {
      targetName = equipmentMap[equipmentCode].name;
      targetLocation = equipmentMap[equipmentCode].location;
      Logger.log(`Tìm thấy TB. Tên="${targetName}", Vị trí="${targetLocation}"`);
    } else if (systemMap[equipmentCode]) {
      targetName = systemMap[equipmentCode];
      targetLocation = "N/A";
      Logger.log(`Tìm thấy HT. Mô tả="${targetName}"`);
    } else {
      targetName = "Mã không hợp lệ";
      targetLocation = "";
      Logger.log(`Không tìm thấy thông tin cho mã "${equipmentCode}"`);
    }

    // Cập nhật cột G và H - CHỈ DÙNG setValue, KHÔNG ĐỊNH DẠNG
    sheet.getRange(rowIndex, COL_PCV_TEN_DOI_TUONG).setValue(targetName); // Cột G
    sheet.getRange(rowIndex, COL_PCV_VI_TRI).setValue(targetLocation);    // Cột H
    
  } catch (error) {
    Logger.log(`Lỗi khi cập nhật thông tin thiết bị: ${error}\nStack: ${error.stack}`);
  }
}


 
/**
 * Gửi email thông báo cho team kỹ thuật (tùy chọn)
 */
function sendNotificationEmail(formData, rowIndex) {
  const subject = "📢 [SHT] Thông báo có phiếu báo hỏng thiết bị mới";
  let body = "Có phiếu báo hỏng thiết bị mới đã được tạo tự động từ form.\n\n";
  body += "Thiết bị/Hệ thống: " + formData.targetCode + "\n";
  body += "Mô tả lỗi: " + formData.faultDescription + "\n";
  body += "Mức độ ưu tiên: " + formData.priority + "\n";
  body += "Người báo: " + formData.reporter + " (" + formData.email + ")\n";
  body += "Thời gian báo: " + formData.timestamp.toLocaleString() + "\n\n";
  if (formData.imageUrl) body += "Ảnh minh họa: " + formData.imageUrl + "\n\n";
  body += "Link đến Phiếu Công Việc: " + SpreadsheetApp.getActiveSpreadsheet().getUrl() + "\n"; 
  body += "Dòng: " + rowIndex + "\n\n";
  body += "Vui lòng xử lý theo quy trình.\n";
  const recipientEmails = [EMAIL_ADMIN]; // Sửa lại nếu cần gửi cho nhiều người
  MailApp.sendEmail({
    to: recipientEmails.join(","),
    subject: subject,
    body: body
  });
}
