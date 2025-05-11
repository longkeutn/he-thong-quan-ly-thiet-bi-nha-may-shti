/**
 * Hàm chính để tạo prefill link và QR code cho các thiết bị từ sheet ValidationSource
 * Được gọi từ menu tùy chỉnh
 */
function generateQrCodesForEquipment() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Hiện thông báo
    ui.alert(
      'Tạo QR Code báo hỏng thiết bị',
      'Script sẽ đọc dữ liệu từ sheet ValidationSource, tạo link prefill và QR code cho từng thiết bị.\n\n' +
      'Kết quả sẽ được lưu trong sheet mới "QR_Codes_Equipment".\n\n' +
      'Bạn đã sẵn sàng tiếp tục?',
      ui.ButtonSet.OK_CANCEL
    );
    
    // Lấy ID form cần prefill
    const formId = getFormId_();
    if (!formId) {
      throw new Error("Cần nhập ID Google Form báo hỏng thiết bị");
    }
    
    // Lấy ID field trong form (mã thiết bị)
    const fieldId = getFormFieldId_();
    if (!fieldId) {
      throw new Error("Cần nhập ID trường Mã thiết bị trên Google Form");
    }
    
    // Đọc dữ liệu từ sheet ValidationSource
    const sourceData = readValidationSourceData_();
    if (!sourceData || sourceData.length === 0) {
      throw new Error("Không tìm thấy dữ liệu trong sheet ValidationSource");
    }
    
    // Tạo hoặc lấy sheet để lưu kết quả
    const resultSheet = createOrGetResultSheet_();
    if (!resultSheet) {
      throw new Error("Không thể tạo hoặc tìm sheet kết quả");
    }
    
    // Tạo link prefill và QR code cho từng thiết bị
    processDataAndGenerateQrCodes_(sourceData, formId, fieldId, resultSheet);
    
    // Thông báo hoàn tất
    ui.alert(
      'Hoàn thành',
      `Đã tạo ${sourceData.length} QR code và link prefill.\nKết quả được lưu trong sheet "QR_Codes_Equipment".`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log(`Lỗi: ${error}\nStack: ${error.stack}`);
    ui.alert('Lỗi', `Có lỗi xảy ra: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Lấy ID Google Form từ người dùng (chỉ cần nhập 1 lần)
 * @return {string} Form ID
 * @private
 */
function getFormId_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let formId = scriptProperties.getProperty('FORM_ID');
  
  if (!formId) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Nhập ID Google Form',
      'Nhập ID của Google Form báo hỏng thiết bị.\n\nID nằm trong URL (phần giữa /d/ và /viewform hoặc /edit):\n' +
      'https://docs.google.com/forms/d/[FORM_ID]/viewform',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      formId = response.getResponseText().trim();
      if (formId) {
        scriptProperties.setProperty('FORM_ID', formId);
      } else {
        throw new Error("ID Google Form không được để trống");
      }
    } else {
      return null;
    }
  }
  
  return formId;
}

/**
 * Lấy ID trường Mã thiết bị trong Google Form từ người dùng (chỉ cần nhập 1 lần)
 * @return {string} Field ID (dạng entry.123456789)
 * @private
 */
function getFormFieldId_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let fieldId = scriptProperties.getProperty('FIELD_ID');
  
  if (!fieldId) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Nhập ID Trường Form',
      'Nhập ID của trường "Mã thiết bị" trong Google Form.\n\n' +
      'ID có dạng "entry.123456789" và có thể lấy từ URL prefill của form.',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.OK) {
      fieldId = response.getResponseText().trim();
      if (fieldId && fieldId.startsWith('entry.')) {
        scriptProperties.setProperty('FIELD_ID', fieldId);
      } else {
        throw new Error("ID trường Form không hợp lệ. Phải bắt đầu bằng 'entry.'");
      }
    } else {
      return null;
    }
  }
  
  return fieldId;
}

/**
 * Đọc dữ liệu từ sheet ValidationSource (cột A)
 * @return {Array} Mảng các mã thiết bị
 * @private
 */
function readValidationSourceData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const validationSheet = ss.getSheetByName('ValidationSource');
  
  if (!validationSheet) {
    throw new Error("Không tìm thấy sheet ValidationSource");
  }
  
  const lastRow = validationSheet.getLastRow();
  if (lastRow < 2) {
    return []; // Không có dữ liệu
  }
  
  // Đọc cột A từ dòng 2 trở đi (bỏ qua header)
  const dataRange = validationSheet.getRange(2, 1, lastRow - 1, 1);
  const data = dataRange.getValues();
  
  // Lọc ra các dòng không trống
  return data.filter(row => row[0] && row[0].toString().trim() !== '')
           .map(row => ({
             original: row[0],                         // Giá trị gốc (có thể là "QUATDL-001 - Quạt trần nhà ăn")
             code: row[0].toString().split(" - ")[0].trim()  // Mã thiết bị (QUATDL-001)
           }));
}

/**
 * Tạo hoặc lấy sheet để lưu kết quả
 * @return {Sheet} Sheet kết quả
 * @private
 */
function createOrGetResultSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheetName = 'QR_Codes_Equipment';
  let resultSheet = ss.getSheetByName(resultSheetName);
  
  if (!resultSheet) {
    resultSheet = ss.insertSheet(resultSheetName);
    
    // Tạo header
    resultSheet.getRange('A1:E1').setValues([['Mã thiết bị', 'Tên thiết bị', 'Link Prefill', 'Link QR Code', 'QR Code']]);
    resultSheet.getRange('A1:E1').setFontWeight('bold');
    resultSheet.setFrozenRows(1);
    
    // Định dạng cột
    resultSheet.setColumnWidth(1, 150);  // Mã thiết bị
    resultSheet.setColumnWidth(2, 250);  // Tên thiết bị
    resultSheet.setColumnWidth(3, 300);  // Link Prefill
    resultSheet.setColumnWidth(4, 300);  // Link QR Code 
    resultSheet.setColumnWidth(5, 150);  // QR Code
  } else {
    // Xóa dữ liệu cũ (giữ header)
    const lastRow = resultSheet.getLastRow();
    if (lastRow > 1) {
      resultSheet.getRange(2, 1, lastRow - 1, 5).clearContent();
    }
  }
  
  return resultSheet;
}

/**
 * Xử lý dữ liệu và tạo QR code cho từng thiết bị
 * @param {Array} sourceData Dữ liệu nguồn
 * @param {string} formId ID Google Form
 * @param {string} fieldId ID trường trong form
 * @param {Sheet} resultSheet Sheet kết quả
 * @private
 */
function processDataAndGenerateQrCodes_(sourceData, formId, fieldId, resultSheet) {
  const data = [];
  
  sourceData.forEach((item, index) => {
    // Tách tên thiết bị nếu có
    let equipmentName = '';
    if (item.original.includes(" - ")) {
      equipmentName = item.original.split(" - ").slice(1).join(" - ").trim();
    }
    
    // Tạo link prefill với item.original thay vì item.code
    const prefillLink = createPrefillLink_(formId, fieldId, item.original);
    
    // Tạo QR code URL
    const qrCodeUrl = createQrCodeUrl_(prefillLink);
    
    // Tạo dữ liệu cho dòng mới
    data.push([
      item.code,
      equipmentName,
      prefillLink,
      qrCodeUrl,
      `=IMAGE("${qrCodeUrl}")`
    ]);
    
    // Báo cáo tiến độ mỗi 20 dòng
    if ((index + 1) % 20 === 0 || index === sourceData.length - 1) {
      Logger.log(`Đã xử lý ${index + 1}/${sourceData.length} thiết bị`);
    }
  });
  
  // Ghi dữ liệu vào sheet kết quả
  if (data.length > 0) {
    resultSheet.getRange(2, 1, data.length, 5).setValues(data);
  }
}

/**
 * Tạo link prefill cho Google Form
 * @param {string} formId ID Google Form
 * @param {string} fieldId ID trường trong form
 * @param {string} equipmentCode Mã thiết bị
 * @return {string} Link prefill
 * @private
 */
function createPrefillLink_(formId, fieldId, equipmentCode) {
  const encodedValue = encodeURIComponent(equipmentCode);
  return `https://docs.google.com/forms/d/e/${formId}/viewform?usp=pp_url&${fieldId}=${encodedValue}`;
}

/**
 * Tạo URL QR code cho link prefill
 * @param {string} prefillLink Link prefill cần tạo QR
 * @return {string} URL QR code
 * @private
 */
function createQrCodeUrl_(prefillLink) {
  const encodedLink = encodeURIComponent(prefillLink);
  return `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodedLink}`;
}

/**
 * Tạo menu tùy chỉnh khi mở spreadsheet
 */
function setupQrCodeMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🔄 QR Code Tools');
  menu.addItem('Tạo QR Code báo hỏng thiết bị', 'generateQrCodesForEquipment');
  menu.addItem('Đặt lại ID Form & ID Field', 'resetFormSettings');
  menu.addToUi();
}

/**
 * Đặt lại cài đặt ID Form và ID Field
 */
function resetFormSettings() {
  const ui = SpreadsheetApp.getUi();
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const response = ui.alert(
    'Đặt lại thông tin Form',
    'Bạn có chắc muốn đặt lại thông tin ID Form và ID Field? Bạn sẽ cần nhập lại chúng khi tạo QR code.',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    scriptProperties.deleteProperty('FORM_ID');
    scriptProperties.deleteProperty('FIELD_ID');
    ui.alert('Đã đặt lại thông tin Form thành công.');
  }
}
