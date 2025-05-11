/**
 * Sao lưu dữ liệu các sheet chính sang các sheet có tiền tố "BACKUP_".
 * Ghi đè bản backup cũ nếu có.
 */
function backupMainSheets() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToBackup = [
      EQUIPMENT_SHEET_NAME,
      PURCHASE_SHEET_NAME,
      HISTORY_SHEET_NAME,
      SHEET_PHIEU_CONG_VIEC
    ];
    sheetsToBackup.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const backupName = "BACKUP_" + sheetName;
      // Xóa sheet backup cũ nếu có
      const oldBackup = ss.getSheetByName(backupName);
      if (oldBackup) ss.deleteSheet(oldBackup);
      // Tạo sheet backup mới
      sheet.copyTo(ss).setName(backupName);
    });
    ui.alert('Đã sao lưu các sheet chính thành công (BACKUP_...)');
  } catch (e) {
    Logger.log(`Lỗi sao lưu sheet: ${e}`);
    ui.alert('Lỗi khi sao lưu: ' + e.message);
  }
}

/**
 * Khôi phục các sheet chính từ bản backup (BACKUP_...).
 * Ghi đè dữ liệu hiện tại (cẩn thận!).
 */
function restoreBackupSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sheetsToRestore = [
      EQUIPMENT_SHEET_NAME,
      PURCHASE_SHEET_NAME,
      HISTORY_SHEET_NAME,
      SHEET_PHIEU_CONG_VIEC
    ];
    sheetsToRestore.forEach(sheetName => {
      const backupName = "BACKUP_" + sheetName;
      const backupSheet = ss.getSheetByName(backupName);
      const mainSheet = ss.getSheetByName(sheetName);
      if (backupSheet && mainSheet) {
        // Xóa dữ liệu trên sheet chính (trừ dòng tiêu đề)
        const lastRow = mainSheet.getLastRow();
        if (lastRow > 1) mainSheet.getRange(2, 1, lastRow - 1, mainSheet.getMaxColumns()).clearContent();
        // Copy dữ liệu từ backup
        const data = backupSheet.getDataRange().getValues();
        if (data.length > 1) {
          mainSheet.getRange(2, 1, data.length - 1, data[0].length).setValues(data.slice(1));
        }
      }
    });
    ui.alert('Đã khôi phục dữ liệu từ các bản backup thành công!');
  } catch (e) {
    Logger.log(`Lỗi khôi phục backup: ${e}`);
    ui.alert('Lỗi khi khôi phục backup: ' + e.message);
  }
}

/**
 * Reset bộ đếm mã thiết bị trong sheet Cấu hình (Cột O).
 */
function resetEquipmentCounter() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) throw new Error(`Không tìm thấy sheet "${SETTINGS_SHEET_NAME}"`);
    const lastRow = settingsSheet.getLastRow();
    if (lastRow < 2) return;
    // Reset tất cả bộ đếm về 1 (Cột O)
    settingsSheet.getRange(2, COL_SETTINGS_COUNTER_NEXT_NUM, lastRow - 1, 1).setValue(1);
    ui.alert('Đã reset bộ đếm mã thiết bị về 1!');
  } catch (e) {
    Logger.log(`Lỗi reset bộ đếm mã thiết bị: ${e}`);
    ui.alert('Lỗi: ' + e.message);
  }
}

/**
 * Reset bộ đếm mã phiếu công việc (trong PropertiesService).
 */
function resetWorkOrderCounter() {
  const ui = SpreadsheetApp.getUi();
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('workOrderIdCounter');
    scriptProperties.deleteProperty('lastWorkOrderIdDate');
    ui.alert('Đã reset bộ đếm mã Phiếu Công Việc!');
  } catch (e) {
    Logger.log(`Lỗi reset bộ đếm mã PCV: ${e}`);
    ui.alert('Lỗi: ' + e.message);
  }
}

/**
 * Reset bộ đếm mã lô mua hàng (trong PropertiesService).
 */
function resetPurchaseCounter() {
  const ui = SpreadsheetApp.getUi();
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('purchaseIdCounter');
    scriptProperties.deleteProperty('lastPurchaseIdDate');
    ui.alert('Đã reset bộ đếm mã Lô Mua Hàng!');
  } catch (e) {
    Logger.log(`Lỗi reset bộ đếm mã Lô MH: ${e}`);
    ui.alert('Lỗi: ' + e.message);
  }
}

