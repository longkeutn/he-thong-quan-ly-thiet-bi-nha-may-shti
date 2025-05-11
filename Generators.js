// ==================================
// CÁC HÀM TẠO ID/MÃ DUY NHẤT CHO HỆ THỐNG 
// ==================================

/**
 * Tạo ID duy nhất cho bản ghi lịch sử bảo trì.
 * @return {string} ID mới theo định dạng yyyyMMddHHmmss-XXXX.
 */
function generateHistoryId() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss") + 
         '-' + Math.floor(Math.random() * 10000);
}

/**
 * Tạo Mã Thiết Bị duy nhất dựa trên Mã VT Loại TB và Bộ đếm theo loại TB.
 * @param {string} equipmentTypeAcronym Mã viết tắt Loại TB (VD: 'QUATDL').
 * @return {string} Mã Thiết Bị mới (VD: 'QUATDL-001') hoặc null nếu có lỗi.
 */
function generateEquipmentId(equipmentTypeAcronym) {
  // Kiểm tra đầu vào
  if (!equipmentTypeAcronym || typeof equipmentTypeAcronym !== 'string' || equipmentTypeAcronym.trim() === "") {
    Logger.log("Lỗi generateEquipmentId: Cần cung cấp Mã VT Loại TB hợp lệ.");
    SpreadsheetApp.getUi().alert("Lỗi tạo mã: Thiếu thông tin Loại Thiết Bị.");
    return null;
  }
  equipmentTypeAcronym = equipmentTypeAcronym.trim();

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    Logger.log(`Đã lấy khóa để tạo Mã TB cho Loại: ${equipmentTypeAcronym}.`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) {
      Logger.log(`Lỗi: Không tìm thấy Sheet "${SETTINGS_SHEET_NAME}"`);
      return null;
    }
    
    // Khóa bộ đếm chính là Mã VT Loại TB
    const counterKey = equipmentTypeAcronym;
    
    // Đọc dữ liệu bộ đếm từ sheet Cấu hình
    const counterDataRange = settingsSheet.getRange(2, COL_SETTINGS_COUNTER_KEY, settingsSheet.getLastRow() - 1, 2);
    const counterValues = counterDataRange.getValues();
    let nextNum = 1; // Số bắt đầu mặc định
    let rowIndex = -1; // Dòng tìm thấy bộ đếm

    // Tìm khóa bộ đếm trong sheet
    for (let i = 0; i < counterValues.length; i++) {
      if (counterValues[i][0] && counterValues[i][0].toString().trim() === counterKey) {
        nextNum = parseInt(counterValues[i][1], 10);
        if (isNaN(nextNum) || nextNum < 1) nextNum = 1;
        rowIndex = i + 2; // +2 vì đọc từ hàng 2
        Logger.log(`Tìm thấy bộ đếm "${counterKey}" tại dòng ${rowIndex}. Số hiện tại là ${nextNum}.`);
        break;
      }
    }

    // Định dạng số thứ tự (3 chữ số) và tạo ID mới
    const sequenceNumber = nextNum.toString().padStart(3, '0');
    const newEquipmentId = `${counterKey}-${sequenceNumber}`;
    const updatedNextNum = nextNum + 1;

    // Cập nhật bộ đếm trong sheet
    if (rowIndex !== -1) {
      // Nếu tìm thấy khóa, cập nhật số đếm
      settingsSheet.getRange(rowIndex, COL_SETTINGS_COUNTER_NEXT_NUM).setValue(updatedNextNum);
      Logger.log(`Bộ đếm "${counterKey}" được cập nhật thành ${updatedNextNum} tại dòng ${rowIndex}`);
    } else {
      // Nếu không tìm thấy khóa, thêm dòng mới
      let newRowData = Array(COL_SETTINGS_COUNTER_NEXT_NUM).fill('');
      newRowData[COL_SETTINGS_COUNTER_KEY - 1] = counterKey;
      newRowData[COL_SETTINGS_COUNTER_NEXT_NUM - 1] = updatedNextNum;
      settingsSheet.appendRow(newRowData);
      Logger.log(`Bộ đếm mới "${counterKey}" được tạo và đặt thành ${updatedNextNum} tại dòng mới.`);
    }

    SpreadsheetApp.flush();
    Logger.log(`Đã tạo Mã TB: ${newEquipmentId}.`);
    return newEquipmentId;

  } catch (e) {
    Logger.log(`Lỗi trong hàm generateEquipmentId: ${e}`);
    SpreadsheetApp.getUi().alert(`Lỗi khi tạo mã thiết bị: ${e}`);
    return null;
  } finally {
    lock.releaseLock();
    Logger.log("Đã giải phóng khóa tạo Mã Thiết Bị.");
  }
}

/**
 * Tạo Mã Lô Mua Hàng / ID Giao Dịch duy nhất.
 * @return {string} Mã mới theo định dạng PO-YYYYMMDD-NNN.
 */
function generatePurchaseId() {
  try {
    const today = new Date();
    const prefix = "PO-";
    const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    const scriptProperties = PropertiesService.getScriptProperties();
    const lastDateKey = 'lastPurchaseIdDate';
    const counterKey = 'purchaseIdCounter';

    let counter = 1;
    const lastDate = scriptProperties.getProperty(lastDateKey);

    if (lastDate === dateString) {
      let currentCounter = scriptProperties.getProperty(counterKey);
      counter = currentCounter ? parseInt(currentCounter, 10) + 1 : 1;
    } else {
      scriptProperties.setProperty(lastDateKey, dateString);
      counter = 1;
    }
    scriptProperties.setProperty(counterKey, counter.toString());

    const sequenceNumber = counter.toString().padStart(3, '0');
    const finalId = prefix + dateString + "-" + sequenceNumber;
    
    Logger.log(`Đã tạo Mã Lô Mua Hàng: ${finalId}`);
    return finalId;

  } catch (e) {
    Logger.log(`Lỗi trong hàm generatePurchaseId: ${e}`);
    return null;
  }
}

/**
 * Tạo Mã Phiếu Công Việc duy nhất.
 * @return {string} Mã Phiếu CV mới theo định dạng PCV-YYYYMMDD-NNN.
 */
function generateWorkOrderId() {
  try {
    const today = new Date();
    const prefix = "PCV-";
    const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");
    const scriptProperties = PropertiesService.getScriptProperties();
    const lastDateKey = 'lastWorkOrderIdDate';
    const counterKey = 'workOrderIdCounter';

    let counter = 1;
    const lastDate = scriptProperties.getProperty(lastDateKey);

    if (lastDate === dateString) {
      let currentCounter = scriptProperties.getProperty(counterKey);
      counter = currentCounter ? parseInt(currentCounter, 10) + 1 : 1;
    } else {
      scriptProperties.setProperty(lastDateKey, dateString);
      counter = 1;
    }
    scriptProperties.setProperty(counterKey, counter.toString());

    const sequenceNumber = counter.toString().padStart(3, '0');
    const finalId = prefix + dateString + "-" + sequenceNumber;
    
    Logger.log(`Đã tạo Mã Phiếu CV: ${finalId}`);
    return finalId;

  } catch (e) {
    Logger.log(`Lỗi trong hàm generateWorkOrderId: ${e}`);
    return null;
  }
}

