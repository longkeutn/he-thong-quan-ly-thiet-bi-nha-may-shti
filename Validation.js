/**
 * Hàm chính để kiểm tra tính nhất quán dữ liệu trên các sheet.
 * Tạo sheet báo cáo lỗi, highlight dòng lỗi và tạo link điều hướng.
 * Được gọi từ Menu tùy chỉnh "🔍 Kiểm tra Tính nhất quán Dữ liệu".
 */
function checkDataConsistency() {
  const FUNCTION_NAME = "checkDataConsistency";
  Logger.log(`===== BẮT ĐẦU ${FUNCTION_NAME} =====`);
  
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportSheetName = "Báo cáo Lỗi Dữ liệu";
  const highlightColor = "#fce8e6"; // Màu hồng nhạt
  const headers = ["Sheet", "Dòng", "Ô Lỗi (Ước lượng)", "Link", "Mô tả Lỗi"];
  
  let reportSheet;
  let allIssues = [];
  let rangesToHighlight = {};
  
  try {
    // ---- Tạo hoặc xóa nội dung Sheet Báo cáo ----
    Logger.log(`[${FUNCTION_NAME}] Chuẩn bị sheet báo cáo "${reportSheetName}"...`);
    reportSheet = ss.getSheetByName(reportSheetName);
    
    if (reportSheet) {
      // Sheet đã tồn tại - xóa nội dung
      reportSheet.clearContents();
      try {
        ss.setActiveSheet(reportSheet, true);
      } catch(e) { 
        Logger.log(`[${FUNCTION_NAME}] Không thể activate report sheet (có thể đã ẩn): ${e}`);
      }
    } else {
      // Tạo sheet mới
      reportSheet = ss.insertSheet(reportSheetName);
      ss.setActiveSheet(reportSheet, true);
    }
    
    // Đặt tiêu đề cho sheet báo cáo
    reportSheet.getRange("A1:E1").setValues([headers]).setFontWeight("bold");
    reportSheet.setFrozenRows(1);
    
    // Cố gắng đặt định dạng cột Link (D) thành Automatic
    try {
      const linkColumn = reportSheet.getRange("D2:D");
      linkColumn.setNumberFormat("@"); // Text
      SpreadsheetApp.flush();
      Utilities.sleep(200);
      linkColumn.setNumberFormat("General"); // Automatic
      Logger.log(`[${FUNCTION_NAME}] Đã đặt định dạng cột Link của Sheet Báo cáo thành Automatic.`);
    } catch(formatErr) {
      Logger.log(`[${FUNCTION_NAME}] Lỗi khi đặt định dạng cột Link: ${formatErr}`);
    }
    
    SpreadsheetApp.flush(); // Đảm bảo sheet sẵn sàng
    
    // ---- Chạy các hàm kiểm tra ----
    Logger.log(`[${FUNCTION_NAME}] Bắt đầu chạy các kiểm tra...`);
    
    // Chạy các kiểm tra cơ bản
    allIssues = allIssues.concat(findMissingTypeLocation_());
    allIssues = allIssues.concat(findOrphanHistoryRecords_());
    
    // Chạy các kiểm tra mở rộng
    allIssues = allIssues.concat(findInconsistentMaintDates_()); // Kiểm tra ngày bảo trì TB
    allIssues = allIssues.concat(findInvalidPurchaseLinks_());   // Kiểm tra Mã Lô Mua Hàng trong TB
    
    Logger.log(`[${FUNCTION_NAME}] Đã hoàn thành tất cả kiểm tra. Tổng số vấn đề: ${allIssues.length}`);
    
    // ---- Xử lý kết quả ----
    if (allIssues.length > 0) {
      // Thu thập dữ liệu báo cáo
      Logger.log(`[${FUNCTION_NAME}] Đang xử lý ${allIssues.length} vấn đề để hiển thị trên báo cáo...`);
      const reportData = [];
      
      allIssues.forEach(issue => {
        Logger.log(`[${issue.sheetName}] Dòng ${issue.row}: ${issue.message}`);
        let rowDataForReport = [issue.sheetName || "Lỗi", issue.row || "?", "", "", issue.message];
        
        if (issue.sheetName !== "Hệ thống" && issue.row > 0) {
          const targetSheet = ss.getSheetByName(issue.sheetName);
          if (targetSheet) {
            const sheetId = targetSheet.getSheetId();
            const errorColumnIndex = issue.column || 1;
            const colLetter = columnToLetter_(errorColumnIndex);
            
            // Tạo notation và URL link sử dụng toán tử +
            const rangeNotation = colLetter + issue.row;
            const cellLinkUrl = "#gid=" + sheetId + "&range=" + rangeNotation;
            const cellLinkFormula = '=HYPERLINK("' + cellLinkUrl + '";"' + rangeNotation + '")';
            
            // Cập nhật dữ liệu báo cáo
            rowDataForReport = [issue.sheetName, issue.row, rangeNotation, cellLinkFormula, issue.message];
            
            // Chuẩn bị danh sách dòng cần highlight
            if (!rangesToHighlight[issue.sheetName]) {
              rangesToHighlight[issue.sheetName] = [];
            }
            rangesToHighlight[issue.sheetName].push(issue.row);
          } else {
            rowDataForReport[4] = `Lỗi: Không tìm thấy sheet '${issue.sheetName}' để tạo link. Lỗi gốc: ${issue.message}`;
          }
        }
        
        reportData.push(rowDataForReport);
      });
      
      // Ghi dữ liệu vào sheet báo cáo
      if (reportData.length > 0) {
        reportSheet.getRange(2, 1, reportData.length, headers.length).setValues(reportData);
        Logger.log(`[${FUNCTION_NAME}] Đã ghi ${reportData.length} dòng vào Sheet Báo cáo.`);
      }
      
      // Đảm bảo dữ liệu được ghi trước khi highlight
      SpreadsheetApp.flush();
      Utilities.sleep(500);
      
      // Highlight các dòng có lỗi
      Logger.log(`[${FUNCTION_NAME}] Bắt đầu thực hiện highlight...`);
      for (const sheetName in rangesToHighlight) {
        const targetSheet = ss.getSheetByName(sheetName);
        if (targetSheet) {
          const rowsToHighlight = rangesToHighlight[sheetName];
          const maxCols = targetSheet.getMaxColumns();
          Logger.log(`[${FUNCTION_NAME}] Highlighting sheet: ${sheetName}, Rows: ${rowsToHighlight.join(',')}, MaxCols: ${maxCols}`);
          
          rowsToHighlight.forEach(rowNum => {
            if (rowNum > 1 && maxCols > 0) {
              try {
                // Highlight cả dòng
                targetSheet.getRange(rowNum, 1, 1, maxCols).setBackground(highlightColor);
                Logger.log(`[${FUNCTION_NAME}] > Đã highlight dòng ${rowNum} sheet ${sheetName}`);
              } catch (highlightError) {
                Logger.log(`[${FUNCTION_NAME}] > Lỗi khi highlight dòng ${rowNum} sheet ${sheetName}: ${highlightError}`);
              }
            } else {
              Logger.log(`[${FUNCTION_NAME}] > Bỏ qua highlight dòng không hợp lệ ${rowNum} hoặc maxCols=0`);
            }
          });
        }
      }
      Logger.log(`[${FUNCTION_NAME}] Đã hoàn thành highlight.`);
      
      // Định dạng lại sheet báo cáo
      try {
        reportSheet.getDataRange().setVerticalAlignment("top");
        reportSheet.autoResizeColumns(1, headers.length);
        Logger.log(`[${FUNCTION_NAME}] Đã định dạng sheet báo cáo.`);
      } catch (resizeError) {
        Logger.log(`[${FUNCTION_NAME}] Lỗi khi tự động thay đổi kích thước cột báo cáo: ${resizeError}`);
      }
      
      // Thông báo cho người dùng
      ui.alert(`Tìm thấy ${allIssues.length} vấn đề. Xem sheet 'Báo cáo Lỗi Dữ liệu'.`);
      
    } else {
      // Không tìm thấy vấn đề nào
      Logger.log(`[${FUNCTION_NAME}] Không tìm thấy vấn đề nào.`);
      reportSheet.getRange("A2").setValue("Không tìm thấy vấn đề nào.");
      ui.alert("Kiểm tra hoàn tất: Không có vấn đề.");
    }
    
  } catch (e) {
    // Xử lý lỗi chung
    Logger.log(`[${FUNCTION_NAME}] Lỗi nghiêm trọng: ${e} \nStack: ${e.stack}`);
    try { 
      if (reportSheet) {
        reportSheet.getRange("A2").setValue(`Lỗi hệ thống: ${e.message}`);
      }
    } catch(err) {
      Logger.log(`[${FUNCTION_NAME}] Không thể ghi lỗi vào sheet báo cáo: ${err}`);
    }
    ui.alert("Lỗi nghiêm trọng khi kiểm tra. Xem Log để biết chi tiết.");
  }
  
  Logger.log(`===== KẾT THÚC ${FUNCTION_NAME} =====`);
}


// ==================================
// CÁC HÀM KIỂM TRA TÍNH NHẤT QUÁN DỮ LIỆU
// ==================================

/**
 * Hàm tiện ích chuyển đổi chỉ số cột thành chữ cái (1->A, 2->B, 27->AA).
 * @param {number} column Chỉ số cột (bắt đầu từ 1).
 * @return {string} Chữ cái tương ứng cho cột.
 * @private
 */
function columnToLetter_(column) {
  if (typeof column !== 'number' || column < 1) {
    return '';
  }
  
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = Math.floor((column - temp - 1) / 26);
  }
  return letter;
}

/**
 * Xóa màu nền highlight khỏi các sheet dữ liệu chính.
 * Hàm này được gọi từ Menu tùy chỉnh "🧹 Xóa Đánh dấu Lỗi".
 */
function clearErrorHighlights() {
  const FUNCTION_NAME = "clearErrorHighlights";
  const ui = SpreadsheetApp.getUi();
  let totalCleared = 0;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToClear = [EQUIPMENT_SHEET_NAME, HISTORY_SHEET_NAME]; // Thêm các sheet khác nếu cần
    
    sheetsToClear.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        try {
          const lastRow = sheet.getLastRow();
          if (lastRow > 1) {
            // Chỉ xóa nếu có dữ liệu ngoài header
            const rangeToChange = sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns());
            rangeToChange.setBackground(null);
            const rowsCleared = lastRow - 1;
            totalCleared += rowsCleared;
            Logger.log(`[${FUNCTION_NAME}] Đã xóa highlight ${rowsCleared} dòng khỏi sheet: ${sheetName}`);
          } else {
            Logger.log(`[${FUNCTION_NAME}] Sheet ${sheetName} chỉ có header hoặc trống. Bỏ qua.`);
          }
        } catch (clearError) {
          Logger.log(`[${FUNCTION_NAME}] Lỗi khi xóa highlight sheet ${sheetName}: ${clearError}`);
        }
      } else {
        Logger.log(`[${FUNCTION_NAME}] Không tìm thấy sheet: ${sheetName}`);
      }
    });
    
    const message = totalCleared > 0 
      ? `Đã xóa các đánh dấu lỗi (${totalCleared} dòng từ ${sheetsToClear.length} sheet).`
      : "Không có đánh dấu lỗi nào cần xóa.";
    ui.alert(message);
    
  } catch (e) {
    Logger.log(`[${FUNCTION_NAME}] Lỗi: ${e}\nStack: ${e.stack}`);
    ui.alert(`Có lỗi xảy ra khi xóa highlight: ${e.message}`);
  }
}

/**
 * [CHECK 1] Tìm các thiết bị thiếu Loại hoặc Vị trí.
 * @return {Array<Object>} Mảng các object lỗi { sheetName, row, column, message }.
 * @private
 */
function findMissingTypeLocation_() {
  const FUNCTION_NAME = "findMissingTypeLocation_";
  const errors = [];
  const sheetName = EQUIPMENT_SHEET_NAME;
  
  try {
    Logger.log(`[${FUNCTION_NAME}] Bắt đầu kiểm tra thiết bị thiếu Loại hoặc Vị trí...`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipmentSheet = ss.getSheetByName(sheetName);
    
    if (!equipmentSheet) {
      Logger.log(`[${FUNCTION_NAME}] Không tìm thấy sheet "${sheetName}"`);
      return [{ 
        sheetName: "Hệ thống", 
        row: 0, 
        message: `Lỗi: Không tìm thấy Sheet ${sheetName}.` 
      }];
    }

    const lastRow = equipmentSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log(`[${FUNCTION_NAME}] Sheet "${sheetName}" không có dữ liệu.`);
      return errors; // Trả về mảng rỗng
    }
    
    const lastColRead = Math.max(COL_EQUIP_ID, COL_EQUIP_TYPE, COL_EQUIP_LOCATION);
    const data = equipmentSheet.getRange(2, 1, lastRow - 1, lastColRead).getValues();
    
    Logger.log(`[${FUNCTION_NAME}] Quét ${data.length} thiết bị để kiểm tra thiếu Loại/Vị trí...`);
    
    let missingTypeCount = 0;
    let missingLocationCount = 0;

    for (let i = 0; i < data.length; i++) {
      const rowIndex = i + 2;
      const equipId = data[i][COL_EQUIP_ID - 1];
      
      // Chỉ kiểm tra các dòng có Mã TB
      if (equipId && equipId.toString().trim() !== "") {
        const equipType = data[i][COL_EQUIP_TYPE - 1];
        const location = data[i][COL_EQUIP_LOCATION - 1];
        const equipIdStr = equipId.toString().trim();

        // Kiểm tra thiếu Loại TB
        if (!equipType || equipType.toString().trim() === "") {
          errors.push({ 
            sheetName, 
            row: rowIndex, 
            column: COL_EQUIP_TYPE, 
            message: `Thiếu 'Loại Thiết Bị / Linh Kiện'. (TB: ${equipIdStr}) -> HD: Chọn Loại TB từ danh sách.` 
          });
          missingTypeCount++;
        }
        
        // Kiểm tra thiếu Vị trí
        if (!location || location.toString().trim() === "") {
          errors.push({ 
            sheetName, 
            row: rowIndex, 
            column: COL_EQUIP_LOCATION, 
            message: `Thiếu 'Vị Trí'. (TB: ${equipIdStr}) -> HD: Chọn Vị trí từ danh sách.` 
          });
          missingLocationCount++;
        }
      }
    }
    
    Logger.log(`[${FUNCTION_NAME}] Kết quả: ${missingTypeCount} TB thiếu Loại, ${missingLocationCount} TB thiếu Vị trí, tổng ${errors.length} vấn đề.`);
  } catch (e) {
    Logger.log(`[${FUNCTION_NAME}] Lỗi: ${e}\nStack: ${e.stack}`);
    errors.push({ 
      sheetName: "Hệ thống", 
      row: 0, 
      message: `Lỗi hệ thống khi kiểm tra thiếu Loại/Vị trí: ${e.message}` 
    });
  }
  
  return errors;
}



// ===== THAY THẾ HÀM NÀY TRONG FILE VALIDATION.GS =====

/**
 * [CHECK 4] Tìm các bản ghi Lịch sử có Mã Đối tượng/Hệ thống không hợp lệ.
 * Kiểm tra xem Mã ở Cột B có tồn tại trong Danh mục TB hoặc DinhNghiaHeThong không.
 * @return {Array<Object>} Mảng các object lỗi { sheetName, row, column, message }.
 * @private
 */
function findOrphanHistoryRecords_() {
  const errors = [];
  const sheetName = HISTORY_SHEET_NAME;
  Logger.log(`--- Bắt đầu Check 4: Kiểm tra Mã Đối tượng/Hệ thống không hợp lệ (Sheet: ${sheetName}) ---`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipmentSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    const historySheet = ss.getSheetByName(sheetName);
    const systemDefSheet = ss.getSheetByName(SHEET_DINH_NGHIA_HE_THONG);

    // Kiểm tra sự tồn tại của các sheet cần thiết
    if (!equipmentSheet || !historySheet || !systemDefSheet) {
      const missingSheet = !equipmentSheet ? EQUIPMENT_SHEET_NAME : 
                          (!historySheet ? sheetName : SHEET_DINH_NGHIA_HE_THONG);
      const message = `Lỗi Check 4: Không tìm thấy Sheet "${missingSheet}".`;
      errors.push({ sheetName: "Hệ thống", row: 0, column: 0, message });
      Logger.log(message);
      return errors;
    }

    // Tải danh sách Mã hợp lệ vào Sets để tra cứu nhanh
    const validEquipmentIdSet = new Set();
    const validSystemCodeSet = new Set();
    
    // 1. Mã Thiết Bị
    if (equipmentSheet.getLastRow() >= 2) {
      const equipmentIds = equipmentSheet.getRange(2, COL_EQUIP_ID, equipmentSheet.getLastRow() - 1, 1).getValues();
      equipmentIds.forEach(row => {
        if (row[0]) {
          const idStr = row[0].toString().trim();
          if (idStr) validEquipmentIdSet.add(idStr);
        }
      });
    }
    Logger.log(` > Check 4: Đã tải ${validEquipmentIdSet.size} Mã TB hợp lệ.`);

    // 2. Mã Hệ thống
    if (systemDefSheet.getLastRow() >= 2) {
      const systemCodes = systemDefSheet.getRange(2, COL_HT_MA, systemDefSheet.getLastRow() - 1, 1).getValues();
      systemCodes.forEach(row => {
        if (row[0]) {
          const codeStr = row[0].toString().trim();
          if (codeStr) validSystemCodeSet.add(codeStr);
        }
      });
    }
    Logger.log(` > Check 4: Đã tải ${validSystemCodeSet.size} Mã Hệ thống hợp lệ.`);

    // Đọc dữ liệu từ Sheet Lịch sử
    const lastHistoryRow = historySheet.getLastRow();
    if (lastHistoryRow < 2) {
      Logger.log(" > Check 4: Sheet Lịch sử không có dữ liệu.");
      return errors;
    }
    
    const historyData = historySheet.getRange(2, 1, lastHistoryRow - 1, 
                          Math.max(COL_HISTORY_ID, COL_HISTORY_TARGET_CODE, COL_HISTORY_EXEC_DATE)).getValues();
    Logger.log(` > Check 4: Đã đọc ${historyData.length} dòng từ sheet Lịch sử.`);

    // Lặp và kiểm tra từng dòng lịch sử
    for (let i = 0; i < historyData.length; i++) {
      const rowIndex = i + 2; // Số dòng thực tế trên sheet
      const historyId = historyData[i][COL_HISTORY_ID - 1];
      const rawTargetValue = historyData[i][COL_HISTORY_TARGET_CODE - 1];
      const execDate = historyData[i][COL_HISTORY_EXEC_DATE - 1];

      // Trích xuất Mã từ Cột B
      let targetCode = "";
      if (rawTargetValue && typeof rawTargetValue === 'string') {
        targetCode = rawTargetValue.split(" - ")[0].trim();
      } else if (rawTargetValue) {
        targetCode = rawTargetValue.toString().trim();
      }

      if (targetCode) {
        // Kiểm tra mã có hợp lệ không
        if (!validEquipmentIdSet.has(targetCode) && !validSystemCodeSet.has(targetCode)) {
          const message = `Mã Đối tượng/Hệ thống '${targetCode}' không tồn tại trong Danh mục TB hoặc Định nghĩa Hệ thống. HD: Kiểm tra lại mã, có thể chạy lại Data Validation hoặc sửa thủ công.`;
          errors.push({ sheetName, row: rowIndex, column: COL_HISTORY_TARGET_CODE, message });
          Logger.log(` >> Lỗi dòng ${rowIndex}: ${message}`);
        }
      } else if (historyId || execDate) {
        // Dòng có dữ liệu nhưng thiếu mã
        const message = `Thiếu thông tin 'Đối tượng / Hệ thống' (Cột B) trong khi dòng có dữ liệu khác (ID hoặc Ngày TH). HD: Bổ sung Mã hoặc xóa dòng nếu không cần thiết.`;
        errors.push({ sheetName, row: rowIndex, column: COL_HISTORY_TARGET_CODE, message });
        Logger.log(` >> Lỗi dòng ${rowIndex}: ${message}`);
      }
    }

    Logger.log(`--- Kết thúc Check 4: Tìm thấy ${errors.length} vấn đề.`);
    return errors;

  } catch (e) {
    Logger.log(`Lỗi nghiêm trọng trong findOrphanHistoryRecords_: ${e} \nStack: ${e.stack}`);
    errors.push({ 
      sheetName: "Hệ thống", 
      row: 0, 
      column: 0, 
      message: `Lỗi hệ thống khi kiểm tra Lịch sử mồ côi: ${e.message}` 
    });
    return errors;
  }
}

// ===== KẾT THÚC HÀM THAY THẾ =====

// ============================================================
// TẠO MÃ VIẾT TẮT (Mã VT) TỰ ĐỘNG CHO SHEET SETTINGS
// ============================================================

/**
 * Hàm trợ giúp để loại bỏ dấu tiếng Việt khỏi chuỗi.
 * @param {string} str Chuỗi đầu vào có dấu.
 * @return {string} Chuỗi đầu ra không dấu.
 * @private
 */
function removeVietnameseAccents_(str) {
  if (!str) return "";
  
  str = str.toString().toLowerCase();
  
  const replacements = {
    'à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ': 'a',
    'è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ': 'e',
    'ì|í|ị|ỉ|ĩ': 'i',
    'ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ': 'o',
    'ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ': 'u',
    'ỳ|ý|ỵ|ỷ|ỹ': 'y',
    'đ': 'd'
  };
  
  for (const [pattern, replacement] of Object.entries(replacements)) {
    str = str.replace(new RegExp(pattern, 'g'), replacement);
  }
  
  return str;
}

/**
 * Hàm trợ giúp để tạo Mã VT cơ bản từ tên.
 * Cố gắng giữ lại dấu '-' nếu có trong tên gốc.
 * @param {string} text Tên Loại TB / Vị trí / Bộ phận.
 * @return {string} Mã VT cơ bản được tạo ra.
 * @private
 */
function createAcronym_(text) {
  if (!text) return "";

  // Xử lý văn bản đầu vào
  let cleanText = removeVietnameseAccents_(text); // Loại bỏ dấu tiếng Việt
  cleanText = cleanText.toUpperCase();            // Chuyển thành chữ HOA

  // Chuẩn hóa văn bản
  cleanText = cleanText.replace(/[^A-Z0-9-]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!cleanText) return "";

  // Tách và xử lý từng phần
  const parts = cleanText.split(' ');
  let acronym = "";
  let lastPartWasWord = false;

  // Duyệt qua từng phần để tạo viết tắt
  for (const part of parts) {
    if (part === "-") {
      // Xử lý dấu gạch ngang đứng riêng
      if (lastPartWasWord) {
        acronym += "-";
        lastPartWasWord = false;
      }
    } else if (part.includes("-")) {
      // Xử lý từ ghép có gạch ngang
      const subParts = part.split('-').filter(Boolean);
      const subAcronym = subParts.map(sp => sp.length > 0 ? sp.charAt(0) : '').join('-');
      
      if (subAcronym) {
        if (lastPartWasWord) {
          acronym += "-";
        }
        acronym += subAcronym;
        lastPartWasWord = true;
      }
    } else if (part.length > 0) {
      // Xử lý từ thông thường
      acronym += part.charAt(0);
      lastPartWasWord = true;
    }
  }

  // Làm sạch kết quả cuối cùng
  return acronym
    .replace(/-{2,}/g, '-')
    .replace(/^-/, '')
    .replace(/-$/, '');
}


// ===== THAY THẾ HÀM NÀY TRONG FILE VALIDATION.GS =====

/**
 * Quét các cột Giá trị (Loại TB - A, Vị Trí - D, Bộ phận - I) trong sheet Cấu hình,
 * nếu cột Mã VT tương ứng (B, E, J) đang trống, sẽ tự động tạo Mã VT,
 * đảm bảo tính duy nhất trong từng cột, ghi lại VÀ ÁP DỤNG ĐỊNH DẠNG CHUẨN.
 * Được gọi từ Menu tùy chỉnh.
 */
function generateMissingAcronyms_Settings() {
  const FUNCTION_NAME = "generateMissingAcronyms_Settings";
  const ui = SpreadsheetApp.getUi();
  let totalUpdates = 0;
  
  try {
    Logger.log(`[${FUNCTION_NAME}] Bắt đầu tạo Mã VT thiếu trong sheet Cấu hình...`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    
    if (!settingsSheet) {
      throw new Error(`Không tìm thấy Sheet "${SETTINGS_SHEET_NAME}"`);
    }

    // Định nghĩa các cặp cột cần xử lý
    const columnPairs = [
      { nameCol: COL_SETTINGS_LOAI_TB_GIATRI, acronymCol: COL_SETTINGS_LOAI_TB_MA, setName: "Loại TB" },  // A -> B
      { nameCol: COL_SETTINGS_VITRI_GIATRI, acronymCol: COL_SETTINGS_VITRI_MA, setName: "Vị Trí" },      // D -> E
      { nameCol: COL_SETTINGS_BOPHAN_GIATRI, acronymCol: COL_SETTINGS_BOPHAN_MA, setName: "Bộ phận" }    // I -> J
    ];

    // Xác định phạm vi dữ liệu cần đọc
    let maxCol = 0;
    columnPairs.forEach(pair => {
      maxCol = Math.max(maxCol, pair.nameCol, pair.acronymCol);
    });

    const startRow = 2;  // Bắt đầu từ dòng 2 (sau tiêu đề)
    const lastRow = settingsSheet.getLastRow();
    
    if (lastRow < startRow) {
      ui.alert("Sheet Cấu hình không có dữ liệu để xử lý.");
      return;
    }

    // Đọc toàn bộ dữ liệu một lần để tối ưu hiệu suất
    const dataRange = settingsSheet.getRange(startRow, 1, lastRow - startRow + 1, maxCol);
    const dataValues = dataRange.getValues();
    Logger.log(`[${FUNCTION_NAME}] Đã đọc ${dataValues.length} dòng dữ liệu từ sheet Cấu hình`);

    // Xử lý từng cặp cột (Loại TB, Vị trí, Bộ phận)
    for (const pair of columnPairs) {
      Logger.log(`[${FUNCTION_NAME}] Đang xử lý: ${pair.setName} (Cột ${columnToLetter_(pair.nameCol)} -> ${columnToLetter_(pair.acronymCol)})`);
      
      // Tạo Set lưu các mã đã tồn tại để kiểm tra trùng lặp
      const existingAcronyms = new Set();
      const updates = [];
      
      // Thu thập các mã đã tồn tại
      dataValues.forEach(row => {
        const acronymValue = row[pair.acronymCol - 1];
        if (acronymValue && acronymValue.toString().trim() !== "") {
          existingAcronyms.add(acronymValue.toString().toUpperCase());
        }
      });
      Logger.log(`[${FUNCTION_NAME}] Tìm thấy ${existingAcronyms.size} Mã VT ${pair.setName} đã tồn tại`);

      // Xử lý từng dòng để tạo mã còn thiếu
      for (let i = 0; i < dataValues.length; i++) {
        const currentRowIndex = startRow + i;
        const nameValue = dataValues[i][pair.nameCol - 1];
        const currentAcronymValue = dataValues[i][pair.acronymCol - 1];

        // Chỉ xử lý các dòng có giá trị tên nhưng chưa có mã
        if (nameValue && nameValue.toString().trim() !== "" && 
            (!currentAcronymValue || currentAcronymValue.toString().trim() === "")) {
          
          Logger.log(`[${FUNCTION_NAME}] Đang xử lý dòng ${currentRowIndex}, Tên ${pair.setName}: "${nameValue}"`);
          
          // Tạo mã cơ bản từ tên
          let baseAcronym = createAcronym_(nameValue);
          if (!baseAcronym) {
            Logger.log(`[${FUNCTION_NAME}] Không thể tạo mã cho "${nameValue}". Bỏ qua.`);
            continue;
          }
          
          // Đảm bảo tính duy nhất của mã
          let uniqueAcronym = baseAcronym;
          let counter = 2;
          let collisionDetected = false;
          
          while (existingAcronyms.has(uniqueAcronym.toUpperCase())) {
            if (!collisionDetected) {
              Logger.log(`[${FUNCTION_NAME}] Mã "${baseAcronym}" cho "${nameValue}" bị trùng. Thêm số.`);
              collisionDetected = true;
            }
            uniqueAcronym = baseAcronym + counter;
            counter++;
          }
          
          // Thêm mã vừa tạo vào set để tránh trùng lặp trong cùng đợt xử lý
          existingAcronyms.add(uniqueAcronym.toUpperCase());
          updates.push({ rowIndex: currentRowIndex, acronym: uniqueAcronym });
          
          if (collisionDetected) {
            Logger.log(`[${FUNCTION_NAME}] Đã tạo mã "${uniqueAcronym}" (từ "${baseAcronym}") sau khi xử lý trùng lặp`);
          } else {
            Logger.log(`[${FUNCTION_NAME}] Đã tạo mã: "${uniqueAcronym}"`);
          }
        }
      }

      // Thực hiện cập nhật và định dạng cho tất cả các mã đã tạo
      if (updates.length > 0) {
        Logger.log(`[${FUNCTION_NAME}] Cập nhật ${updates.length} Mã VT ${pair.setName}...`);
        
        for (const update of updates) {
          try {
            const cell = settingsSheet.getRange(update.rowIndex, pair.acronymCol);
            cell.setValue(update.acronym);
            
            // Áp dụng định dạng chuẩn
            cell.setFontSize(12)
                .setVerticalAlignment("middle")
                .setHorizontalAlignment("center")
                .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
          } catch (e) {
            Logger.log(`[${FUNCTION_NAME}] Lỗi cập nhật mã "${update.acronym}" tại dòng ${update.rowIndex}: ${e}`);
          }
        }
        
        totalUpdates += updates.length;
        // Đảm bảo dữ liệu được ghi vào sheet sau mỗi cặp cột
        SpreadsheetApp.flush();
      } else {
        Logger.log(`[${FUNCTION_NAME}] Không có Mã VT ${pair.setName} nào cần tạo mới`);
      }
    }

    // Thông báo kết quả
    const message = totalUpdates > 0
      ? `Đã tạo, cập nhật và định dạng ${totalUpdates} Mã VT trong sheet Cấu hình.`
      : "Không tìm thấy Mã VT nào cần tạo mới trong sheet Cấu hình.";
    
    ui.alert(message);
    Logger.log(`[${FUNCTION_NAME}] Hoàn thành. ${message}`);

  } catch (e) {
    Logger.log(`[${FUNCTION_NAME}] Lỗi: ${e}\nStack: ${e.stack}`);
    ui.alert(`Đã xảy ra lỗi: ${e.message}. Vui lòng kiểm tra Log.`);
  }
}



// ===== Đồng bộ Hệ thống Cơ bản cho Vị trí mới vào DinhNghiaHeThong =====

// ===== THAY THẾ HOÀN TOÀN HÀM NÀY TRONG FILE VALIDATION.GS =====

/**
 * Đồng bộ Hệ thống Cơ bản vào sheet DinhNghiaHeThong dựa trên sheet Cấu hình.
 * 1. Xóa các mã hệ thống cơ bản "mồ côi" (có tiền tố cơ bản, có mã vị trí nhưng vị trí đó không còn tồn tại).
 * 2. Thêm các mã hệ thống cơ bản còn thiếu cho các vị trí hợp lệ hiện có, dựa trên "Loại Vị Trí" để loại trừ các hệ thống không phù hợp.
 * KHÔNG XÓA các mã được nhập thủ công (không khớp với mẫu cơ bản).
 * Được gọi từ Menu.
 */
function syncBasicSystemsForNewLocations() {
  const FUNCTION_NAME = "syncBasicSystemsForNewLocations";
  const ui = SpreadsheetApp.getUi();
  
  Logger.log(`===== Bắt đầu ${FUNCTION_NAME} =====`);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    const systemDefSheet = ss.getSheetByName(SHEET_DINH_NGHIA_HE_THONG);

    if (!configSheet || !systemDefSheet) {
      throw new Error(`Không tìm thấy sheet "${SETTINGS_SHEET_NAME}" hoặc "${SHEET_DINH_NGHIA_HE_THONG}".`);
    }
    
    // Kiểm tra hằng số Loại Vị Trí
    if (typeof COL_SETTINGS_VITRI_TYPE === 'undefined') {
      throw new Error("Lỗi cấu hình: Thiếu khai báo hằng số COL_SETTINGS_VITRI_TYPE trong Config.gs.");
    }

    // --- Định nghĩa các mẫu hệ thống cơ bản VÀ quy tắc loại trừ ---
    const basicSystems = [
      { prefix: "HT-DIEN-CS-", descTemplate: "Hệ thống Điện Chiếu sáng - ", excludeTypes: [] }, // Áp dụng cho mọi loại
      { prefix: "HT-DIEN-OC-", descTemplate: "Hệ thống Điện Ổ cắm - ", excludeTypes: [] }, // Áp dụng cho mọi loại
      { prefix: "HT-NUOC-CAP-", descTemplate: "Hệ thống Cấp nước - ", excludeTypes: ["Phòng thờ", "Kho", "Sân", "Ngoài trời"] }, // Ví dụ loại trừ
      { prefix: "HT-NUOC-THOAT-", descTemplate: "Hệ thống Thoát nước (Sàn/Chung) - ", excludeTypes: ["Phòng thờ", "Văn phòng", "Phòng họp"] }, // Ví dụ loại trừ
      { prefix: "TB-DENUV-", descTemplate: "Thiết bị Đèn UV - ", excludeTypes: ["Phòng thờ", "WC", "Hành lang", "Sân", "Ngoài trời", "Khu vực chung"] }, // Ví dụ loại trừ
      { prefix: "TB-DENCT-", descTemplate: "Thiết bị Đèn Diệt côn trùng - ", excludeTypes: ["Phòng thờ", "WC", "Phòng họp", "Văn phòng"] }, // Ví dụ loại trừ
      { prefix: "HT-PCCC-KV-", descTemplate: "Hệ thống PCCC Khu vực - ", excludeTypes: ["WC", "Phòng thờ"] }, // Ví dụ loại trừ
      { prefix: "HT-HVAC-THONGGIO-", descTemplate: "Hệ thống Thông gió - ", excludeTypes: ["Ngoài trời", "Sân"] }, // Ví dụ loại trừ
      { prefix: "HT-HVAC-HUTMUI-", descTemplate: "Hệ thống Hút mùi/Khói - ", excludeTypes: ["Văn phòng", "Phòng họp", "Hành lang", "Phòng thờ", "WC", "Ngoài trời", "Sân", "Khu vực chung"]} // Chỉ áp dụng cho nơi có khả năng phát sinh mùi/khói
    ];
    
    // Tạo một Set chứa các tiền tố cơ bản để kiểm tra khi xóa
    const basicPrefixes = new Set(basicSystems.map(sys => sys.prefix));

    // --- 1. Đọc dữ liệu và chuẩn bị ---
    // Đọc Vị trí hợp lệ (Tên, Mã VT, Loại) từ Cấu hình
    const locations = []; // [{ name: '...', acronym: '...', type: '...' }]
    const validLocationAcronyms = new Set();
    const lastConfigRow = configSheet.getLastRow();
    
    if (lastConfigRow >= 2) {
      const locationData = configSheet.getRange(2, 1, lastConfigRow - 1, Math.max(COL_SETTINGS_VITRI_GIATRI, COL_SETTINGS_VITRI_MA, COL_SETTINGS_VITRI_TYPE)).getValues();
      locationData.forEach(row => {
        const locName = row[COL_SETTINGS_VITRI_GIATRI - 1] ? row[COL_SETTINGS_VITRI_GIATRI - 1].toString().trim() : null;
        const locAcronym = row[COL_SETTINGS_VITRI_MA - 1] ? row[COL_SETTINGS_VITRI_MA - 1].toString().trim() : null;
        const locType = row[COL_SETTINGS_VITRI_TYPE - 1] ? row[COL_SETTINGS_VITRI_TYPE - 1].toString().trim() : null;
        
        if (locName && locAcronym) { // Chỉ lấy những vị trí có đủ Tên và Mã VT
          locations.push({ name: locName, acronym: locAcronym, type: locType || "" });
          validLocationAcronyms.add(locAcronym);
        }
      });
    }
    
    Logger.log(`[${FUNCTION_NAME}] Đã đọc ${locations.length} vị trí hợp lệ từ Cấu hình.`);

    // Đọc Mã Hệ thống và Mô tả hiện có từ DinhNghiaHeThong
    const systemDefData = []; // [{ code: '...', description: '...', rowNum: R }]
    const existingSystemCodes = new Set();
    const lastSystemRow = systemDefSheet.getLastRow();
    
    if (lastSystemRow >= 2) {
      const systemValues = systemDefSheet.getRange(2, 1, lastSystemRow - 1, 2).getValues();
      systemValues.forEach((row, index) => {
        const code = row[0] ? row[0].toString().trim() : null;
        if (code) {
          systemDefData.push({ code: code, description: row[1] || "", rowNum: index + 2 });
          existingSystemCodes.add(code);
        }
      });
    }
    
    Logger.log(`[${FUNCTION_NAME}] Đã đọc ${systemDefData.length} mã hệ thống hiện có từ DinhNghiaHeThong.`);

    // --- 2. Tìm và xử lý các mã hệ thống cơ bản "mồ côi" ---
    const rowsToDelete = []; // Chứa số thứ tự dòng cần xóa
    
    systemDefData.forEach(system => {
      // Kiểm tra xem mã có phải là mã hệ thống cơ bản không
      const isBasicSystem = Array.from(basicPrefixes).some(prefix => system.code.startsWith(prefix));
      
      if (isBasicSystem) {
        // Tìm vị trí trong mã (phần sau tiền tố)
        let locationCode = null;
        for (const prefix of basicPrefixes) {
          if (system.code.startsWith(prefix)) {
            locationCode = system.code.substring(prefix.length);
            break;
          }
        }
        
        // Nếu tìm được locationCode và nó không tồn tại trong danh sách vị trí hợp lệ
        if (locationCode && !validLocationAcronyms.has(locationCode)) {
          rowsToDelete.push(system.rowNum);
          Logger.log(`[${FUNCTION_NAME}] Đánh dấu để xóa mã hệ thống mồ côi: ${system.code} (dòng ${system.rowNum}) - Mã vị trí '${locationCode}' không còn hợp lệ.`);
        }
      }
    });
    
    // Xóa các dòng theo thứ tự từ dưới lên để tránh shift index
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // Sắp xếp giảm dần
      
      for (const rowNum of rowsToDelete) {
        systemDefSheet.deleteRow(rowNum);
      }
      
      Logger.log(`[${FUNCTION_NAME}] Đã xóa ${rowsToDelete.length} mã hệ thống mồ côi.`);
      SpreadsheetApp.flush();
    } else {
      Logger.log(`[${FUNCTION_NAME}] Không tìm thấy mã hệ thống mồ côi nào cần xóa.`);
    }
    
    // --- 3. Tạo và thêm các mã hệ thống cơ bản còn thiếu cho các vị trí hợp lệ ---
    const newSystemRows = []; // Mảng chứa các dòng sẽ thêm vào sheet
    
    // Lặp qua từng vị trí
    locations.forEach(location => {
      // Lặp qua từng mẫu hệ thống cơ bản
      basicSystems.forEach(system => {
        // Kiểm tra xem loại vị trí có bị loại trừ không
        const shouldExclude = system.excludeTypes.some(excludeType => 
          location.type.toLowerCase() === excludeType.toLowerCase());
        
        if (!shouldExclude) {
          // Tạo mã hệ thống mới
          const newSystemCode = system.prefix + location.acronym;
          
          // Kiểm tra xem mã này đã tồn tại chưa
          if (!existingSystemCodes.has(newSystemCode)) {
            // Tạo mô tả
            const newSystemDesc = system.descTemplate + location.name;
            
            // Thêm vào danh sách chờ
            newSystemRows.push([newSystemCode, newSystemDesc]);
            existingSystemCodes.add(newSystemCode); // Thêm vào set để tránh trùng lặp
            Logger.log(`[${FUNCTION_NAME}] Tạo mã hệ thống mới: ${newSystemCode} - ${newSystemDesc}`);
          }
        }
      });
    });
    
    // Thêm các dòng mới vào sheet
    if (newSystemRows.length > 0) {
      const lastRow = systemDefSheet.getLastRow();
      systemDefSheet.getRange(lastRow + 1, 1, newSystemRows.length, 2).setValues(newSystemRows);
      
      // Áp dụng định dạng cho các dòng mới
      try {
        const newRowsRange = systemDefSheet.getRange(lastRow + 1, 1, newSystemRows.length, 2);
        newRowsRange.setFontSize(11)
                   .setVerticalAlignment("middle");
      } catch (formatErr) {
        Logger.log(`[${FUNCTION_NAME}] Lỗi khi định dạng các dòng mới: ${formatErr}`);
      }
      
      Logger.log(`[${FUNCTION_NAME}] Đã thêm ${newSystemRows.length} mã hệ thống mới.`);
    } else {
      Logger.log(`[${FUNCTION_NAME}] Không có mã hệ thống mới nào cần thêm.`);
    }
    
    // --- 4. Sắp xếp sheet ---
    try {
      if (systemDefSheet.getLastRow() > 2) {
        systemDefSheet.getRange(2, 1, systemDefSheet.getLastRow() - 1, 2).sort({column: 1, ascending: true});
        Logger.log(`[${FUNCTION_NAME}] Đã sắp xếp lại các mã hệ thống theo thứ tự A-Z.`);
      }
    } catch (sortErr) {
      Logger.log(`[${FUNCTION_NAME}] Không thể sắp xếp sheet: ${sortErr}`);
    }
    
    // --- 5. Thông báo kết quả ---
    ui.alert(`Đồng bộ hoàn tất:\n- Đã xóa ${rowsToDelete.length} mã hệ thống mồ côi.\n- Đã thêm ${newSystemRows.length} mã hệ thống mới.`);
    
    Logger.log(`===== Kết thúc ${FUNCTION_NAME} =====`);
    
  } catch (e) {
    Logger.log(`[${FUNCTION_NAME}] Lỗi: ${e}\nStack: ${e.stack}`);
    ui.alert(`Đã xảy ra lỗi: ${e.message}`);
  }
}


// ===== THÊM CÁC HÀM MỚI NÀY VÀO CUỐI FILE VALIDATION.GS =====

/**
 * [CHECK MỚI] Tìm các dòng có Ngày BT Tiếp theo không nhất quán với Ngày BT cuối và Tần suất.
 * @return {Array<Object>} Mảng các object lỗi { sheetName, row, column, message }.
 * @private
 */
function findInconsistentMaintDates_() {
  const FUNCTION_NAME = "findInconsistentMaintDates_";
  const errors = [];
  const sheetName = EQUIPMENT_SHEET_NAME;
  
  Logger.log(`--- Bắt đầu ${FUNCTION_NAME}: Kiểm tra Ngày Bảo trì không nhất quán (Sheet: ${sheetName}) ---`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipmentSheet = ss.getSheetByName(sheetName);
    
    // Kiểm tra sheet có tồn tại không
    if (!equipmentSheet) {
      const errorMsg = `Lỗi ${FUNCTION_NAME}: Không tìm thấy Sheet "${sheetName}".`;
      errors.push({ 
        sheetName: "Hệ thống", 
        row: 0, 
        column: 0, 
        message: errorMsg
      });
      Logger.log(errorMsg);
      return errors;
    }
    
    // Kiểm tra hàm calculateNextMaintenanceDate có tồn tại không
    if (typeof calculateNextMaintenanceDate !== 'function') {
      const errorMsg = `Lỗi ${FUNCTION_NAME}: Không tìm thấy hàm 'calculateNextMaintenanceDate'. Cần kiểm tra file Calculations.gs.`;
      errors.push({ 
        sheetName: "Hệ thống", 
        row: 0, 
        column: 0, 
        message: errorMsg
      });
      Logger.log(errorMsg);
      return errors;
    }

    // Xác định phạm vi dữ liệu cần đọc
    const lastColNeeded = Math.max(COL_EQUIP_ID, COL_EQUIP_MAINT_FREQ, COL_EQUIP_MAINT_LAST, COL_EQUIP_MAINT_NEXT);
    const startRow = 2;
    const lastRow = equipmentSheet.getLastRow();

    // Kiểm tra sheet có dữ liệu không
    if (lastRow < startRow) {
      Logger.log(`${FUNCTION_NAME}: Sheet ${sheetName} không có dữ liệu để kiểm tra.`);
      return errors;
    }

    // Đọc dữ liệu từ sheet
    const data = equipmentSheet.getRange(startRow, 1, lastRow - startRow + 1, lastColNeeded).getValues();
    Logger.log(`${FUNCTION_NAME}: Đã đọc ${data.length} dòng từ sheet ${sheetName}.`);

    // Đếm số lượng để thống kê
    let inconsistentCount = 0;
    let missingCount = 0;
    let invalidFormatCount = 0;

    // Kiểm tra từng dòng dữ liệu
    for (let i = 0; i < data.length; i++) {
      const rowIndex = startRow + i;
      const equipId = data[i][COL_EQUIP_ID - 1];
      const maintFreq = data[i][COL_EQUIP_MAINT_FREQ - 1]; // Cột Q (17)
      const lastMaintDateRaw = data[i][COL_EQUIP_MAINT_LAST - 1]; // Cột R (18)
      const nextMaintDateRaw = data[i][COL_EQUIP_MAINT_NEXT - 1]; // Cột S (19)

      // Chỉ kiểm tra nếu có đủ dữ liệu cần thiết
      if (equipId && lastMaintDateRaw instanceof Date && maintFreq && maintFreq.toString().trim() !== "") {
        const lastMaintDate = new Date(lastMaintDateRaw);
        lastMaintDate.setHours(0, 0, 0, 0); // Chuẩn hóa về đầu ngày
        
        const nextMaintDate = (nextMaintDateRaw instanceof Date) ? new Date(nextMaintDateRaw) : null;
        if (nextMaintDate) nextMaintDate.setHours(0, 0, 0, 0); // Chuẩn hóa nếu không null

        // Tính ngày tiếp theo dự kiến
        let calculatedNextDate = null;
        try {
          calculatedNextDate = calculateNextMaintenanceDate(lastMaintDate, maintFreq.toString().trim());
          if (calculatedNextDate instanceof Date) {
            calculatedNextDate.setHours(0, 0, 0, 0); // Chuẩn hóa kết quả tính toán
          }
        } catch (calcError) {
          Logger.log(`${FUNCTION_NAME} >> Lỗi khi tính ngày BT tiếp theo cho TB ${equipId} dòng ${rowIndex}: ${calcError}`);
        }

        // Xử lý các trường hợp lỗi
        if (calculatedNextDate instanceof Date) {
          // TH1: Có thể tính được ngày dự kiến và đã có ngày trong cột S, nhưng không khớp nhau
          if (nextMaintDate instanceof Date && calculatedNextDate.getTime() !== nextMaintDate.getTime()) {
            const formattedCalcDate = Utilities.formatDate(calculatedNextDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
            const formattedNextDate = Utilities.formatDate(nextMaintDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
            
            const message = `Ngày BT Tiếp theo (${formattedNextDate}) không khớp với tính toán (${formattedCalcDate}) dựa trên Ngày BT cuối và Tần suất "${maintFreq}". HD: Chạy menu "🗓️ Tính & Cập nhật Ngày BT Tiếp theo (TB)" hoặc cập nhật thủ công.`;
            
            errors.push({ 
              sheetName, 
              row: rowIndex, 
              column: COL_EQUIP_MAINT_NEXT, 
              message 
            });
            
            inconsistentCount++;
            Logger.log(`${FUNCTION_NAME} >> Không khớp (TB: ${equipId}) dòng ${rowIndex}: ${message}`);
          } 
          // TH2: Có thể tính được ngày dự kiến nhưng cột S trống
          else if (!nextMaintDate) {
            const formattedCalcDate = Utilities.formatDate(calculatedNextDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
            
            const message = `Cột Ngày BT Tiếp theo đang trống trong khi có thể tính toán được: ${formattedCalcDate} (từ Ngày BT cuối và Tần suất "${maintFreq}"). HD: Chạy menu "🗓️ Tính & Cập nhật Ngày BT Tiếp theo (TB)".`;
            
            errors.push({ 
              sheetName, 
              row: rowIndex, 
              column: COL_EQUIP_MAINT_NEXT, 
              message 
            });
            
            missingCount++;
            Logger.log(`${FUNCTION_NAME} >> Thiếu ngày (TB: ${equipId}) dòng ${rowIndex}: Đề xuất ${formattedCalcDate}`);
          }
        } 
        // TH3: Có Ngày BT cuối và Tần suất nhưng không tính được ngày dự kiến
        else if (lastMaintDate && maintFreq) {
          const message = `Không thể tính Ngày BT Tiếp theo. Định dạng Tần suất "${maintFreq}" có thể không hợp lệ. HD: Kiểm tra lại chuỗi tần suất (vd: '3 tháng', '1 năm', 'Hàng tuần').`;
          
          errors.push({ 
            sheetName, 
            row: rowIndex, 
            column: COL_EQUIP_MAINT_FREQ, 
            message 
          });
          
          invalidFormatCount++;
          Logger.log(`${FUNCTION_NAME} >> Định dạng không hợp lệ (TB: ${equipId}) dòng ${rowIndex}: "${maintFreq}"`);
        }
      }
    }

    // Tổng kết kết quả kiểm tra
    const totalIssues = inconsistentCount + missingCount + invalidFormatCount;
    Logger.log(`--- Kết thúc ${FUNCTION_NAME}: Tìm thấy ${totalIssues} vấn đề.`);
    Logger.log(`   - ${inconsistentCount} không khớp`);
    Logger.log(`   - ${missingCount} thiếu ngày`);
    Logger.log(`   - ${invalidFormatCount} định dạng tần suất không hợp lệ`);
    
    return errors;

  } catch (e) {
    const errorMsg = `Lỗi nghiêm trọng trong ${FUNCTION_NAME}: ${e}\nStack: ${e.stack}`;
    Logger.log(errorMsg);
    errors.push({ 
      sheetName: "Hệ thống", 
      row: 0, 
      column: 0, 
      message: `Lỗi hệ thống khi kiểm tra Ngày Bảo trì: ${e.message}` 
    });
    return errors;
  }
}



/**
 * [CHECK MỚI] Tìm các dòng trong Danh mục TB có Mã Lô Mua Hàng không hợp lệ.
 * Kiểm tra xem Mã Lô MH ở Cột J (TB) có tồn tại trong Cột A (Mua Hàng) không.
 * @return {Array<Object>} Mảng các object lỗi { sheetName, row, column, message }.
 * @private
 */
function findInvalidPurchaseLinks_() {
  const FUNCTION_NAME = "findInvalidPurchaseLinks_";
  const errors = [];
  const equipSheetName = EQUIPMENT_SHEET_NAME;
  const purchaseSheetName = PURCHASE_SHEET_NAME;
  
  Logger.log(`--- Bắt đầu ${FUNCTION_NAME}: Kiểm tra Mã Lô Mua Hàng không hợp lệ (Sheet: ${equipSheetName}) ---`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipmentSheet = ss.getSheetByName(equipSheetName);
    const purchaseSheet = ss.getSheetByName(purchaseSheetName);

    // Kiểm tra sự tồn tại của các sheet cần thiết
    if (!equipmentSheet || !purchaseSheet) {
      const missingSheet = !equipmentSheet ? equipSheetName : purchaseSheetName;
      const errorMsg = `Lỗi ${FUNCTION_NAME}: Không tìm thấy Sheet "${missingSheet}".`;
      errors.push({ 
        sheetName: "Hệ thống", 
        row: 0, 
        column: 0, 
        message: errorMsg
      });
      Logger.log(errorMsg);
      return errors;
    }

    // Tạo Set chứa danh sách Mã Lô Mua Hàng hợp lệ
    const validPurchaseIdSet = new Set();
    let purchaseIdCount = 0;
    
    if (purchaseSheet.getLastRow() >= 2) {
      const purchaseIds = purchaseSheet.getRange(2, COL_PURCHASE_ID, purchaseSheet.getLastRow() - 1, 1).getValues();
      purchaseIds.forEach(row => {
        if (row[0]) {
          const idStr = row[0].toString().trim();
          if (idStr) {
            validPurchaseIdSet.add(idStr);
            purchaseIdCount++;
          }
        }
      });
    }
    
    Logger.log(`${FUNCTION_NAME}: Đã tải ${purchaseIdCount} Mã Lô MH hợp lệ từ sheet "${purchaseSheetName}".`);

    // Đọc dữ liệu từ Sheet Danh mục Thiết bị
    const lastEquipRow = equipmentSheet.getLastRow();
    if (lastEquipRow < 2) {
      Logger.log(`${FUNCTION_NAME}: Sheet ${equipSheetName} không có dữ liệu để kiểm tra.`);
      return errors;
    }
    
    // Đọc các cột cần thiết: ID(A) và Mã Lô MH(J)
    const equipData = equipmentSheet.getRange(2, 1, lastEquipRow - 1, COL_EQUIP_PURCHASE_ID).getValues();
    Logger.log(`${FUNCTION_NAME}: Đã đọc ${equipData.length} dòng từ sheet "${equipSheetName}".`);

    // Thống kê
    let checkedCount = 0;
    let invalidCount = 0;

    // Kiểm tra từng dòng thiết bị
    for (let i = 0; i < equipData.length; i++) {
      const rowIndex = i + 2;
      const equipId = equipData[i][COL_EQUIP_ID - 1]; // Dữ liệu cột A
      const purchaseIdRaw = equipData[i][COL_EQUIP_PURCHASE_ID - 1]; // Dữ liệu cột J

      // Chỉ kiểm tra khi có giá trị Mã Lô MH
      if (purchaseIdRaw && purchaseIdRaw.toString().trim() !== "") {
        const purchaseId = purchaseIdRaw.toString().trim();
        checkedCount++;
        
        // Kiểm tra tính hợp lệ của Mã Lô MH
        if (!validPurchaseIdSet.has(purchaseId)) {
          const message = `Mã Lô Mua Hàng '${purchaseId}' không tồn tại trong sheet '${purchaseSheetName}'. HD: Kiểm tra lại Mã Lô hoặc thông tin bên sheet Mua Hàng.`;
          errors.push({ 
            sheetName: equipSheetName, 
            row: rowIndex, 
            column: COL_EQUIP_PURCHASE_ID, 
            message 
          });
          
          invalidCount++;
          Logger.log(`${FUNCTION_NAME} >> TB ${equipId || 'Không có ID'} (dòng ${rowIndex}): Mã Lô MH '${purchaseId}' không hợp lệ.`);
        }
      }
    }

    // Tổng kết kết quả
    Logger.log(`--- Kết thúc ${FUNCTION_NAME}: Đã kiểm tra ${checkedCount} Mã Lô MH, tìm thấy ${invalidCount} vấn đề.`);
    return errors;

  } catch (e) {
    const errorMsg = `Lỗi nghiêm trọng trong ${FUNCTION_NAME}: ${e}\nStack: ${e.stack}`;
    Logger.log(errorMsg);
    errors.push({ 
      sheetName: "Hệ thống", 
      row: 0, 
      column: 0, 
      message: `Lỗi hệ thống khi kiểm tra Mã Lô Mua Hàng: ${e.message}` 
    });
    return errors;
  }
}

// ===== KẾT THÚC PHẦN THÊM HÀM MỚI =====


/**
 * Xóa các dòng thiết bị trùng mã trong sheet Danh mục Thiết bị (giữ lại dòng đầu tiên).
 */
function cleanDuplicateEquipmentRows() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    if (!sheet) throw new Error(`Không tìm thấy sheet "${EQUIPMENT_SHEET_NAME}"`);
    const data = sheet.getDataRange().getValues();
    const seen = new Set();
    let rowsToDelete = [];
    for (let i = 1; i < data.length; i++) { // Bỏ qua header
      const id = data[i][COL_EQUIP_ID - 1];
      if (id && seen.has(id)) {
        rowsToDelete.push(i + 1); // Dòng thực tế trên sheet
      } else if (id) {
        seen.add(id);
      }
    }
    // Xóa từ dưới lên
    rowsToDelete.reverse().forEach(row => sheet.deleteRow(row));
    ui.alert(`Đã xóa ${rowsToDelete.length} dòng thiết bị trùng mã.`);
  } catch (e) {
    Logger.log(`Lỗi làm sạch trùng lặp: ${e}`);
    ui.alert("Lỗi làm sạch: " + e.message);
  }
}
