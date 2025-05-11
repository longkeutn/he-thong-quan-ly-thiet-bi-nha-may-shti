/**
 * Xử lý tự động cập nhật thông tin bảo hành khi thay đổi trạng thái trong sheet Lịch sử
 */
function onEditHistory(e) {
  try {
    // Chỉ xử lý trong sheet Lịch sử
    if (!e || !e.range || e.value === undefined) return;
    
    const sheet = e.range.getSheet();
    if (sheet.getName() !== HISTORY_SHEET_NAME) return;
    
    const col = e.range.getColumn();
    const row = e.range.getRow();
    
    // Nếu đánh dấu cột L (Kiểm tra theo dõi bảo hành)
    if (col === COL_HISTORY_WARRANTY_CHECK && row > 1) {
      const isWarranty = e.value === true;
      
      if (isWarranty) {
        sheet.getRange(row, COL_HISTORY_STATUS).setValue("Đang bảo hành");
        
        const warrantyStatus = sheet.getRange(row, COL_HISTORY_WARRANTY_REQ_STAT).getValue();
        if (!warrantyStatus) {
          sheet.getRange(row, COL_HISTORY_WARRANTY_REQ_STAT).setValue("Đã gửi yêu cầu");
        }
      }
    }
  } catch (error) {
    Logger.log(`Lỗi trong onEditHistory: ${error}`);
  }
}

/**
 * Tạo trigger cho onEditHistory
 */
function createHistoryEditTrigger() {
  try {
    // Xóa các trigger cũ
    const allTriggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getHandlerFunction() === 'onEditHistory') {
        ScriptApp.deleteTrigger(allTriggers[i]);
      }
    }
    
    // Tạo trigger mới
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('onEditHistory')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
    
    SpreadsheetApp.getUi().alert("Đã tạo trigger onEditHistory thành công!");
  } catch (error) {
    Logger.log(`Lỗi khi tạo trigger: ${error}`);
    SpreadsheetApp.getUi().alert("Lỗi khi tạo trigger: " + error);
  }
}
