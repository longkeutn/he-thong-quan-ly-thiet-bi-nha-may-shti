// ==================================
// MÃ.GS - FILE CHÍNH CHỨA CÁC HÀM XỬ LÝ NGHIỆP VỤ
// ==================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const mainMenu = ui.createMenu('✨ Tiện ích [SHT]');

  // --- MENU TIỆN ÍCH CHUNG (luôn hiển thị cho mọi user) ---
  mainMenu
    .addItem('⚙️ Tạo Mã & Xử lý Dòng TB Mới', 'processNewEquipmentRows')
    .addItem('🛒 Tạo Mã Lô Mua Hàng & Cập nhật Bảo hành', 'processPurchaseRow')
    .addItem('🆔 Tạo ID & Xử lý Dòng Lịch sử Mới', 'processNewHistoryRows')
    .addItem('📋 Tạo phiếu công việc cho đội kỹ thuật', 'generateTechnicianWorkOrderSheet')
    .addSeparator()
    .addItem('✅ Hoàn thành Phiếu CV & Lưu Lịch sử', 'showCompletionDialog')
    .addItem('🗓️ Tính & Cập nhật Ngày BT Tiếp theo (TB)', 'calculateAndUpdateNextMaintDates')
    .addSeparator()
    .addItem('🔎 Tra cứu Lịch sử Bảo trì / Sửa chữa', 'getTargetForHistorySearch')
    .addItem('🏢 Tra cứu theo Vị trí', 'showLocationSearchView')
    .addItem('👨‍👦 Tra cứu Thiết bị Con', 'showParentChildSearchView')
    .addSeparator()
    .addItem('🔍 Kiểm tra Tính nhất quán Dữ liệu', 'checkDataConsistency')
    .addItem('🧹 Xóa Đánh dấu Lỗi', 'clearErrorHighlights')
    .addSeparator();

  // --- MENU CON: Cài đặt & Cấu hình ---
  var settingsSubMenu = ui.createMenu('⚙️ Cài đặt & Cấu hình');
  settingsSubMenu.addItem('🆔 Tạo Mã VT còn thiếu (Settings)', 'generateMissingAcronyms_Settings');
  settingsSubMenu.addItem('🔄 Đồng bộ Hệ thống Cơ bản cho Vị trí', 'syncBasicSystemsForNewLocations');
  mainMenu.addSubMenu(settingsSubMenu);

  // --- MENU TIỆN ÍCH NÂNG CAO ---
  var advancedMenu = ui.createMenu('🛠️ Tiện ích nâng cao');
  advancedMenu.addItem('Xuất sheet thiết bị ra Excel/CSV', 'exportEquipmentSheetToCsv');
  advancedMenu.addItem('Tạo QR code cho mã thiết bị', 'showQrSidebar');
  advancedMenu.addItem('Làm sạch dữ liệu trùng lặp/thừa', 'cleanDuplicateEquipmentRows');
  mainMenu.addSubMenu(advancedMenu);

// --- MENU BẢO HÀNH - THỢ NGOÀI ---
var advancedMenu = ui.createMenu('🛠️ Bảo hành - Thợ ngoài');
advancedMenu.addItem('📋 Báo cáo thiết bị đang BH/thuê ngoài', 'createExternalServiceReport');
  advancedMenu.addItem('🔍 Kiểm tra bảo hành thiết bị hiện tại', 'checkCurrentEquipmentWarranty');
  advancedMenu.addItem('📊 Báo cáo thiết bị theo bảo hành', 'showWarrantyReport');
  advancedMenu.addItem('📋 Báo cáo thiết bị đang trong quá trình bảo hành', 'showDevicesInWarrantyProcess');
  mainMenu.addSubMenu(advancedMenu);

  // --- MENU HỖ TRỢ NHẬP LIỆU NHANH ---
  var quickInputMenu = ui.createMenu('📝 Hỗ trợ nhập liệu nhanh');
  quickInputMenu.addItem('Tạo nhanh thiết bị mẫu', 'insertSampleEquipmentRow');
  quickInputMenu.addItem('Tạo nhanh phiếu công việc mẫu', 'insertSampleWorkOrderRow');
  quickInputMenu.addItem('Tạo nhanh lịch sử mẫu', 'insertSampleHistoryRow');
  mainMenu.addSubMenu(quickInputMenu);

  // --- MENU BÁO CÁO & THỐNG KÊ ---
  var reportMenu = ui.createMenu('📊 Báo cáo & Thống kê')
    .addItem('Báo cáo thiết bị theo loại/vị trí', 'reportEquipmentByType')
    .addItem('Báo cáo phiếu công việc theo trạng thái', 'reportWorkOrderByStatus')
    .addItem('Báo cáo lịch sử bảo trì tháng/quý', 'reportMaintenanceHistory')
    .addItem('Báo cáo lỗi dữ liệu tổng hợp', 'reportDataErrors');
  mainMenu.addSubMenu(reportMenu);

  // --- MENU HƯỚNG DẪN & TRỢ GIÚP ---
  var helpMenu = ui.createMenu('❓ Hướng dẫn & Trợ giúp');
  helpMenu.addItem('Xem hướng dẫn sử dụng', 'showUserGuideSidebar');
  helpMenu.addItem('Liên hệ hỗ trợ kỹ thuật', 'showSupportContactSidebar');
  helpMenu.addItem('Kiểm tra phiên bản code & nhật ký cập nhật', 'showVersionInfo');
  mainMenu.addSubMenu(helpMenu);

  // --- MENU QUẢN TRỊ & BẢO MẬT (chỉ cho admin) ---
  if (typeof isCurrentUserAdmin === 'function' && isCurrentUserAdmin()) {
    var adminMenu = ui.createMenu('🛡️ Quản trị & Sao lưu');
    adminMenu.addItem('Sao lưu dữ liệu các sheet chính', 'backupMainSheets');
    adminMenu.addItem('Khôi phục dữ liệu từ bản sao lưu gần nhất', 'restoreBackupSheets');
    adminMenu.addItem('Reset bộ đếm mã thiết bị', 'resetEquipmentCounter');
    adminMenu.addItem('Reset bộ đếm mã phiếu CV', 'resetWorkOrderCounter');
    adminMenu.addItem('Reset bộ đếm mã lô mua hàng', 'resetPurchaseCounter');
    mainMenu.addSubMenu(adminMenu);

    var securityMenu = ui.createMenu('🛡️ Quản lý bảo mật & phân quyền');
    securityMenu.addItem('Khóa sheet Cấu hình', 'protectSettingsSheet');
    securityMenu.addItem('Mở khóa sheet Cấu hình', 'unprotectSettingsSheet');
    securityMenu.addItem('Khóa cột Mã TB', 'protectEquipmentIdColumn');
    securityMenu.addItem('Mở khóa cột Mã TB', 'unprotectEquipmentIdColumn');
    securityMenu.addItem('Khóa cột Ngày BT Cuối & Tiếp theo', 'protectMaintenanceDateColumns');
    securityMenu.addItem('Mở khóa cột Ngày BT Cuối & Tiếp theo', 'unprotectMaintenanceDateColumns');
    securityMenu.addItem('Khóa tất cả dòng tiêu đề', 'protectAllHeaderRows');
    mainMenu.addSubMenu(securityMenu);
  }
  mainMenu.addToUi();

  // HOẶC tạo menu QR riêng biệt
  const qrMenu = ui.createMenu('🔄 QR Code Tools');
  qrMenu.addItem('Tạo QR Code báo hỏng thiết bị', 'generateQrCodesForEquipment');
  qrMenu.addItem('Đặt lại ID Form & ID Field', 'resetFormSettings');
  qrMenu.addToUi();
}



// =============================================
// NHÓM CHỨC NĂNG: XỬ LÝ THIẾT BỊ MỚI
// =============================================

/**
 * Xử lý các dòng được chọn trong Sheet "Danh mục Thiết bị".
 * - CHỈ tạo Mã Thiết Bị mới (theo Loại - NNN) nếu ô Mã TB (A) trống.
 * - Cập nhật thông tin Mua hàng (K, L, M) nếu có Mã Lô MH (J).
 * - Áp dụng định dạng chuẩn cho các ô được ghi/cập nhật (Trừ cột S và các cột ngày L, M).
 * - KHÔNG tính toán hay ghi Ngày BT Tiếp theo (S).
 * Hàm này được gọi từ Menu "⚙️ Tạo Mã & Xử lý Dòng TB Mới".
 */
function processNewEquipmentRows() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipmentSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    if (!equipmentSheet) throw new Error(`Không tìm thấy Sheet "${EQUIPMENT_SHEET_NAME}"`);
    if (typeof generateEquipmentId !== 'function') throw new Error("Lỗi hệ thống: Không tìm thấy hàm 'generateEquipmentId'. Kiểm tra file Generators.gs.");
    if (typeof getAcronyms !== 'function') throw new Error("Lỗi hệ thống: Không tìm thấy hàm 'getAcronyms'. Kiểm tra file DataAccess.gs.");
    if (typeof getPurchaseInfo !== 'function') throw new Error("Lỗi hệ thống: Không tìm thấy hàm 'getPurchaseInfo'. Kiểm tra file DataAccess.gs.");
    
    const selectedRange = equipmentSheet.getActiveRange();
    if (!selectedRange) { ui.alert("Vui lòng chọn ít nhất một dòng cần xử lý."); return; }

    const startRow = selectedRange.getRow();
    const numRows = selectedRange.getNumRows();
    let idGeneratedCount = 0, purchaseInfoUpdatedCount = 0, errorCount = 0, formatErrorCount = 0;

    // Đọc các cột cần thiết
    const lastColRead = Math.max(COL_EQUIP_ID, COL_EQUIP_TYPE, COL_EQUIP_PURCHASE_ID, COL_EQUIP_LOCATION, COL_EQUIP_SUPPLIER, COL_EQUIP_PURCHASE_DATE, COL_EQUIP_WARRANTY_END);
    const selectedDataRange = equipmentSheet.getRange(startRow, 1, numRows, lastColRead);
    const selectedValues = selectedDataRange.getValues();

    for (let i = 0; i < numRows; i++) {
      const currentRowIndex = startRow + i;
      if (currentRowIndex === 1) continue;

      let rowData = selectedValues[i];
      let equipmentId = rowData[COL_EQUIP_ID - 1]; 
      const equipmentType = rowData[COL_EQUIP_TYPE - 1]; 
      const purchaseId = rowData[COL_EQUIP_PURCHASE_ID - 1]; 

      try {
        // BƯỚC 1: Tạo Mã TB nếu ô A trống
        if (!equipmentId || equipmentId.toString().trim() === "") {
          Logger.log(`Dòng ${currentRowIndex}: Ô Mã TB trống, tiến hành tạo mã...`);
          
          if (!equipmentType || equipmentType.toString().trim() === "") {
            Logger.log(` Lỗi dòng ${currentRowIndex}: Thiếu Loại Thiết Bị (Cột C) để tạo mã.`);
            const cell = equipmentSheet.getRange(currentRowIndex, COL_EQUIP_ID);
            cell.setValue("LỖI: Thiếu Loại TB"); try{ cell.setFontColor("red");} catch(e){}
            errorCount++;
            continue;
          }
          
          const acronyms = getAcronyms(equipmentType, null);
          if (!acronyms || !acronyms.type) {
             Logger.log(` Lỗi dòng ${currentRowIndex}: Không lấy được Mã VT Loại TB cho "${equipmentType}". Kiểm tra sheet Cấu hình.`);
             const cell = equipmentSheet.getRange(currentRowIndex, COL_EQUIP_ID);
             cell.setValue("LỖI: Mã VT Loại TB"); try{ cell.setFontColor("red");} catch(e){}
             errorCount++;
             continue;
          }

          // Tạo ID mới theo loại
          const newId = generateEquipmentId(acronyms.type);
          if (newId) {
            const idCell = equipmentSheet.getRange(currentRowIndex, COL_EQUIP_ID);
            idCell.setValue(newId);
            
            // Định dạng ô Mã TB
            try {
              idCell.setFontSize(12).setVerticalAlignment("middle").setHorizontalAlignment("center").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setFontColor(null);
            } catch (fmtErr) { 
              Logger.log(` >> LỖI ĐỊNH DẠNG ô Mã TB (A${currentRowIndex}): ${fmtErr}`); 
              formatErrorCount++; 
            }
            
            idGeneratedCount++;
            Logger.log(` > Đã tạo Mã TB mới: ${newId}`);
          } else {
            Logger.log(` Lỗi dòng ${currentRowIndex}: Không tạo được Mã TB từ Generators.gs.`);
             const cell = equipmentSheet.getRange(currentRowIndex, COL_EQUIP_ID);
             cell.setValue("LỖI TẠO MÃ"); try{ cell.setFontColor("red");} catch(e){}
            errorCount++;
            continue;
          }
        } else {
           Logger.log(`Dòng ${currentRowIndex}: Đã có Mã TB "${equipmentId}". Bỏ qua bước tạo mã.`);
        }

        // BƯỚC 2: Cập nhật thông tin mua hàng nếu có Mã Lô MH
        if (purchaseId && purchaseId.toString().trim() !== "") {
           const purchaseIdStr = purchaseId.toString().trim();
           const purchaseInfo = getPurchaseInfo(purchaseIdStr);
           if (purchaseInfo) {
               let updatesMade = false;
               
               // Cập nhật NCC (K)
               const supplierCell = equipmentSheet.getRange(currentRowIndex, COL_EQUIP_SUPPLIER);
               if (supplierCell.getValue() != purchaseInfo.supplier) {
                   supplierCell.setValue(purchaseInfo.supplier || "");
                   try { 
                     supplierCell.setFontSize(12).setVerticalAlignment("middle").setHorizontalAlignment("left").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); 
                     updatesMade = true; 
                   } catch (fmtErr) { 
                     Logger.log(` >> LỖI ĐỊNH DẠNG ô NCC (K${currentRowIndex}): ${fmtErr}`); 
                     formatErrorCount++;
                   }
               }
               
               // Cập nhật Ngày Mua (L)
               const purDateCell = equipmentSheet.getRange(currentRowIndex, COL_EQUIP_PURCHASE_DATE);
               if ((purDateCell.getValue() instanceof Date && purchaseInfo.purchaseDate instanceof Date && purDateCell.getValue().getTime() !== purchaseInfo.purchaseDate.getTime()) || (!purDateCell.getValue() && purchaseInfo.purchaseDate)) {
                   purDateCell.setValue(purchaseInfo.purchaseDate); 
                   updatesMade = true;
               }
               
               // Cập nhật Hạn BH (M)
               const warEndCell = equipmentSheet.getRange(currentRowIndex, COL_EQUIP_WARRANTY_END);
               if ((warEndCell.getValue() instanceof Date && purchaseInfo.warrantyEnd instanceof Date && warEndCell.getValue().getTime() !== purchaseInfo.warrantyEnd.getTime()) || (!warEndCell.getValue() && purchaseInfo.warrantyEnd)) {
                   warEndCell.setValue(purchaseInfo.warrantyEnd); 
                   updatesMade = true;
               }
               
               if (updatesMade) { 
                 purchaseInfoUpdatedCount++; 
                 Logger.log(` > Đã cập nhật TT Mua hàng cho dòng ${currentRowIndex} từ Mã Lô ${purchaseIdStr}.`); 
               }
           } else { 
             Logger.log(` > Không tìm thấy thông tin cho Mã Lô ${purchaseIdStr} khi cập nhật dòng ${currentRowIndex}.`); 
           }
        }
      } catch (procError) {
         Logger.log(`Lỗi xử lý dữ liệu dòng ${currentRowIndex}: ${procError}`);
         errorCount++;
         try { 
           equipmentSheet.getRange(currentRowIndex, COL_EQUIP_ID).setValue("LỖI XỬ LÝ"); 
         } catch(e){}
      }

       if ((idGeneratedCount > 0 || purchaseInfoUpdatedCount > 0) && (idGeneratedCount + purchaseInfoUpdatedCount) % 10 === 0) {
         SpreadsheetApp.flush();
       }
    }

    // Thông báo kết quả cuối cùng
    let message = `Hoàn thành:\n- Tạo mới ID cho ${idGeneratedCount} dòng.\n- Cập nhật TT Mua hàng cho ${purchaseInfoUpdatedCount} dòng.`;
    if (errorCount > 0) { message += `\n- Có ${errorCount} lỗi xử lý dòng.`; }
    if (formatErrorCount > 0) { message += `\n- Có ${formatErrorCount} lỗi định dạng.`; }
    ui.alert(message);

  } catch (e) {
    Logger.log(`Lỗi nghiêm trọng trong processNewEquipmentRows: ${e} \nStack: ${e.stack}`);
    ui.alert(`Đã xảy ra lỗi nghiêm trọng: ${e}. Vui lòng kiểm tra Nhật ký thực thi.`);
  }
}

/**
 * Xử lý các dòng được chọn trong Sheet "Lịch sử Bảo trì / Sửa chữa".
 * - Tạo ID Lịch sử mới nếu cột A trống
 * - Tự động điền và định dạng thông tin Tên và Hiển thị (Cột C, D)
 * - Cập nhật Ngày bảo trì gần nhất cho Thiết bị nếu là bảo trì định kỳ
 * Hàm này được gọi từ Menu hoặc từ saveHistoryFromDialog.
 */
function processNewHistoryRows() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    const equipmentSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    const htSheet = ss.getSheetByName(SHEET_DINH_NGHIA_HE_THONG);
    
    if (!historySheet) throw new Error(`Không tìm thấy Sheet "${HISTORY_SHEET_NAME}"`);
    if (!equipmentSheet) throw new Error(`Không tìm thấy Sheet "${EQUIPMENT_SHEET_NAME}"`);
    if (!htSheet) throw new Error(`Không tìm thấy Sheet "${SHEET_DINH_NGHIA_HE_THONG}"`);
    
    // Tạo Maps để tra cứu
    const equipmentMap = {};
    if (equipmentSheet.getLastRow() >= 2) {
      const lastEquipCol = Math.max(COL_EQUIP_ID, COL_EQUIP_NAME, COL_EQUIP_LOCATION);
      const equipData = equipmentSheet.getRange(2, 1, equipmentSheet.getLastRow() - 1, lastEquipCol).getValues();
      equipData.forEach(row => {
        const id = row[COL_EQUIP_ID - 1];
        if (id) {
          const idStr = id.toString().trim();
          if (idStr) {
            equipmentMap[idStr] = {
              name: row[COL_EQUIP_NAME - 1] || '',
              location: row[COL_EQUIP_LOCATION - 1] || ''
            };
          }
        }
      });
    }
    
    const systemMap = {};
    if (htSheet.getLastRow() >= 2) {
      const htData = htSheet.getRange(2, 1, htSheet.getLastRow() - 1, 2).getValues();
      htData.forEach(row => {
        const code = row[COL_HT_MA - 1];
        if (code) {
          const codeStr = code.toString().trim();
          if (codeStr) {
            systemMap[codeStr] = {
              name: row[COL_HT_MO_TA - 1] || '',
              location: 'N/A'
            };
          }
        }
      });
    }
    
    const selectedRange = historySheet.getActiveRange();
    if (!selectedRange) {
      ui.alert("Vui lòng chọn ít nhất một dòng lịch sử cần xử lý.");
      return;
    }
    
    const startRow = selectedRange.getRow();
    const numRows = selectedRange.getNumRows();
    let idGeneratedCount = 0, autoFilledCount = 0, maintDateUpdatedCount = 0, errorCount = 0;
    
    // Đọc dữ liệu các dòng được chọn
    const lastColRead = Math.max(COL_HISTORY_ID, COL_HISTORY_TARGET_CODE, COL_HISTORY_TARGET_NAME, 
                                 COL_HISTORY_DISPLAY_NAME, COL_HISTORY_EXEC_DATE, COL_HISTORY_WORK_TYPE);
    const selectedData = historySheet.getRange(startRow, 1, numRows, lastColRead).getValues();
    
    for (let i = 0; i < numRows; i++) {
      const currentRowIndex = startRow + i;
      if (currentRowIndex === 1) continue; // Bỏ qua dòng tiêu đề
      
      try {
        const rowData = selectedData[i];
        const historyId = rowData[COL_HISTORY_ID - 1];
        const targetCode = rowData[COL_HISTORY_TARGET_CODE - 1];
        const execDate = rowData[COL_HISTORY_EXEC_DATE - 1];
        const workType = rowData[COL_HISTORY_WORK_TYPE - 1];
        
        // Kiểm tra cột B (Target Code)
        if (!targetCode) {
          Logger.log(`Dòng ${currentRowIndex}: Thiếu Mã Đối tượng/Hệ thống (B). Bỏ qua.`);
          errorCount++;
          continue;
        }
        
        // Trích xuất mã từ chuỗi (nếu targetCode có định dạng "CODE - Name")
        let cleanTargetCode = "";
        if (typeof targetCode === 'string' && targetCode.includes(" - ")) {
          cleanTargetCode = targetCode.split(" - ")[0].trim();
        } else {
          cleanTargetCode = targetCode.toString().trim();
        }
        
        if (!cleanTargetCode) {
          Logger.log(`Dòng ${currentRowIndex}: Mã trích xuất từ "${targetCode}" không hợp lệ. Bỏ qua.`);
          errorCount++;
          continue;
        }
        
        // BƯỚC 1: Tạo ID nếu cột A trống
        if (!historyId) {
          const newId = generateHistoryId();
          if (newId) {
            const idCell = historySheet.getRange(currentRowIndex, COL_HISTORY_ID);
            idCell.setValue(newId);
            try {
              idCell.setFontSize(10).setVerticalAlignment("middle").setHorizontalAlignment("left");
            } catch (fmtErr) {
              Logger.log(`Lỗi định dạng ID mới ở dòng ${currentRowIndex}: ${fmtErr}`);
            }
            idGeneratedCount++;
            Logger.log(`Dòng ${currentRowIndex}: Đã tạo ID lịch sử ${newId}.`);
          } else {
            Logger.log(`Dòng ${currentRowIndex}: Lỗi tạo ID mới.`);
            errorCount++;
          }
        }
        
        // BƯỚC 2: Điền/cập nhật thông tin Tên và Tên hiển thị
        let targetInfo = equipmentMap[cleanTargetCode] || systemMap[cleanTargetCode];
        if (targetInfo) {
          const nameCell = historySheet.getRange(currentRowIndex, COL_HISTORY_TARGET_NAME);
          const displayNameCell = historySheet.getRange(currentRowIndex, COL_HISTORY_DISPLAY_NAME);
          
          nameCell.setValue(targetInfo.name);
          
          const displayText = targetInfo.location && targetInfo.location !== 'N/A' 
            ? `${cleanTargetCode} - ${targetInfo.name} (${targetInfo.location})`
            : `${cleanTargetCode} - ${targetInfo.name}`;
          
          displayNameCell.setValue(displayText);
          
          try {
            nameCell.setFontSize(10).setVerticalAlignment("middle");
            displayNameCell.setFontSize(10).setVerticalAlignment("middle");
          } catch (fmtErr) {
            Logger.log(`Lỗi định dạng tên ở dòng ${currentRowIndex}: ${fmtErr}`);
          }
          
          autoFilledCount++;
          Logger.log(`Dòng ${currentRowIndex}: Đã điền thông tin tên và hiển thị.`);
        } else {
          Logger.log(`Dòng ${currentRowIndex}: Không tìm thấy thông tin cho mã "${cleanTargetCode}". Bỏ qua điền thông tin.`);
        }
        
        // BƯỚC 3: Cập nhật Ngày Bảo trì gần nhất nếu là PM
        if (equipmentMap[cleanTargetCode] && execDate instanceof Date && 
            workType && workType.toString().trim().toLowerCase() === "bảo trì định kỳ") {
          
          const equipRows = equipmentSheet.getRange(2, COL_EQUIP_ID, equipmentSheet.getLastRow() - 1, 1).getValues();
          let equipRowIndex = -1;
          
          for (let j = 0; j < equipRows.length; j++) {
            if (equipRows[j][0] && equipRows[j][0].toString().trim() === cleanTargetCode) {
              equipRowIndex = j + 2; // +2 vì bắt đầu từ dòng 2
              break;
            }
          }
          
          if (equipRowIndex > 0) {
            const lastMaintCell = equipmentSheet.getRange(equipRowIndex, COL_EQUIP_MAINT_LAST);
            const currentLastMaint = lastMaintCell.getValue();
            
            // Chỉ cập nhật nếu chưa có ngày cũ hoặc ngày mới lớn hơn
            if (!currentLastMaint || 
                (currentLastMaint instanceof Date && execDate.getTime() > currentLastMaint.getTime())) {
              lastMaintCell.setValue(execDate);
              maintDateUpdatedCount++;
              Logger.log(`Đã cập nhật Ngày BT Gần nhất cho TB ${cleanTargetCode} (Dòng ${equipRowIndex}) thành ${execDate.toLocaleDateString()}.`);
            } else {
              Logger.log(`Không cập nhật Ngày BT Gần nhất cho TB ${cleanTargetCode} vì ngày hiện tại (${currentLastMaint instanceof Date ? currentLastMaint.toLocaleDateString() : 'null'}) mới hơn ngày thực hiện (${execDate.toLocaleDateString()}).`);
            }
          } else {
            Logger.log(`Không tìm thấy dòng TB ${cleanTargetCode} để cập nhật Ngày BT Gần nhất.`);
          }
        }
        
      } catch (rowError) {
        Logger.log(`Lỗi xử lý dòng ${currentRowIndex}: ${rowError}`);
        errorCount++;
      }
      
      // Flush định kỳ để tránh timeout
      if ((idGeneratedCount > 0 || autoFilledCount > 0 || maintDateUpdatedCount > 0) && 
          (idGeneratedCount + autoFilledCount + maintDateUpdatedCount) % 10 === 0) {
        SpreadsheetApp.flush();
      }
    }
    
    // Thông báo kết quả nếu được gọi từ Menu (không phải từ saveHistoryFromDialog)
    const callerFunction = (new Error()).stack.split('\n')[2].trim().split(' ')[1];
    if (callerFunction !== 'saveHistoryFromDialog') {
      let message = `Hoàn thành:\n` +
                    `- Tạo ID cho ${idGeneratedCount} dòng.\n` +
                    `- Điền thông tin tên/hiển thị cho ${autoFilledCount} dòng.\n` +
                    `- Cập nhật Ngày BT Gần nhất cho ${maintDateUpdatedCount} thiết bị.`;
      if (errorCount > 0) message += `\n- Có ${errorCount} lỗi xử lý dòng.`;
      ui.alert(message);
    }
    
    return {
      idGenerated: idGeneratedCount,
      autoFilled: autoFilledCount,
      maintDateUpdated: maintDateUpdatedCount,
      errors: errorCount
    };
    
  } catch (e) {
    Logger.log(`Lỗi nghiêm trọng trong processNewHistoryRows: ${e}\nStack: ${e.stack}`);
    if ((new Error()).stack.split('\n')[2].trim().split(' ')[1] !== 'saveHistoryFromDialog') {
      ui.alert(`Đã xảy ra lỗi: ${e}. Vui lòng kiểm tra Nhật ký thực thi.`);
    }
    return { errors: 1 };
  }
}

/**
 * Xử lý dòng được chọn trong Sheet "Chi tiết Mua Hàng & Nhà Cung Cấp".
 * - Tạo Mã Lô Mua Hàng mới nếu ô A trống.
 * - Tính Ngày Hạn BH nếu có Ngày bắt đầu BH và Thời hạn BH.
 * Hàm này được gọi từ Menu "🛒 Tạo Mã Lô Mua Hàng & Cập nhật Bảo hành".
 */
function processPurchaseRow() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const purchaseSheet = ss.getSheetByName(PURCHASE_SHEET_NAME);
    
    if (!purchaseSheet) throw new Error(`Không tìm thấy Sheet "${PURCHASE_SHEET_NAME}"`);
    
    const selectedRange = purchaseSheet.getActiveRange();
    if (!selectedRange) {
      ui.alert("Vui lòng chọn dòng Mua Hàng cần xử lý.");
      return;
    }
    
    const startRow = selectedRange.getRow();
    const numRows = selectedRange.getNumRows();
    let idGeneratedCount = 0, warrantyUpdatedCount = 0, errorCount = 0;
    
    for (let i = 0; i < numRows; i++) {
      const currentRowIndex = startRow + i;
      if (currentRowIndex === 1) continue; // Bỏ qua dòng tiêu đề
      
      try {
        // Kiểm tra Mã Lô Mua Hàng trống
        const idCell = purchaseSheet.getRange(currentRowIndex, COL_PURCHASE_ID);
        const currentId = idCell.getValue();
        
        if (!currentId || currentId.toString().trim() === "") {
          // Tạo Mã mới
          const newId = generatePurchaseId();
          if (newId) {
            idCell.setValue(newId);
            try {
              idCell.setFontSize(12)
                   .setVerticalAlignment("middle")
                   .setHorizontalAlignment("center")
                   .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
            } catch (fmtErr) {
              Logger.log(`Lỗi định dạng ID Mua Hàng mới ở dòng ${currentRowIndex}: ${fmtErr}`);
            }
            idGeneratedCount++;
            Logger.log(`Đã tạo Mã Lô Mua Hàng mới ${newId} tại dòng ${currentRowIndex}.`);
          } else {
            Logger.log(`Lỗi tạo ID Mua Hàng mới ở dòng ${currentRowIndex}.`);
            errorCount++;
          }
        }
        
        // Cập nhật Ngày Kết thúc Bảo hành
        const startDateCell = purchaseSheet.getRange(currentRowIndex, COL_PURCHASE_WARRANTY_START);
        const monthsCell = purchaseSheet.getRange(currentRowIndex, COL_PURCHASE_WARRANTY_MONTHS);
        const endDateCell = purchaseSheet.getRange(currentRowIndex, COL_PURCHASE_WARRANTY_END);
        
        const startDate = startDateCell.getValue();
        const months = monthsCell.getValue();
        
        if (startDate instanceof Date && !isNaN(startDate) && months && !isNaN(months)) {
          // Tính ngày kết thúc bảo hành
          const endDate = new Date(startDate);
          endDate.setMonth(endDate.getMonth() + parseInt(months));
          
          // Cập nhật ô
          endDateCell.setValue(endDate);
          warrantyUpdatedCount++;
          Logger.log(`Đã cập nhật Ngày Kết thúc BH tại dòng ${currentRowIndex}: ${endDate.toLocaleDateString()}.`);
        }
      } catch (rowError) {
        Logger.log(`Lỗi xử lý dòng Mua Hàng ${currentRowIndex}: ${rowError}`);
        errorCount++;
      }
    }
    
    // Thông báo kết quả
    let message = `Hoàn thành:\n- Tạo mới Mã Lô Mua Hàng cho ${idGeneratedCount} dòng.\n- Cập nhật Ngày Kết thúc BH cho ${warrantyUpdatedCount} dòng.`;
    if (errorCount > 0) message += `\n- Có ${errorCount} lỗi xử lý dòng.`;
    ui.alert(message);
    
  } catch (e) {
    Logger.log(`Lỗi nghiêm trọng trong processPurchaseRow: ${e}\nStack: ${e.stack}`);
    ui.alert(`Đã xảy ra lỗi: ${e}. Vui lòng kiểm tra Nhật ký thực thi.`);
  }
}

/**
 * Tính toán và cập nhật Ngày Bảo trì Tiếp theo (Cột S) cho các dòng được chọn
 * trong sheet Danh mục Thiết bị, dựa trên Ngày BT cuối (R) và Tần suất (Q).
 * KHÔNG tạo mã ID mới. Chỉ tính toán và ghi ngày.
 * Hàm này được gọi từ Menu "🗓️ Tính & Cập nhật Ngày BT Tiếp theo (TB)".
 */
function calculateAndUpdateNextMaintDates() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipmentSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    
    if (!equipmentSheet) throw new Error(`Không tìm thấy Sheet "${EQUIPMENT_SHEET_NAME}"`);
    
    if (typeof calculateNextMaintenanceDate !== 'function') {
      throw new Error("Lỗi hệ thống: Không tìm thấy hàm 'calculateNextMaintenanceDate'. Kiểm tra file Calculations.gs.");
    }

    const selectedRange = equipmentSheet.getActiveRange();
    if (!selectedRange) { 
      ui.alert("Vui lòng chọn ít nhất một dòng thiết bị cần tính Ngày BT Tiếp theo.");
      return;
    }

    const startRow = selectedRange.getRow();
    const numRows = selectedRange.getNumRows();
    let calculatedCount = 0, errorCount = 0, skippedCount = 0;

    // Đọc dữ liệu tần suất và ngày BT cuối
    const lastColRead = Math.max(COL_EQUIP_MAINT_LAST, COL_EQUIP_MAINT_FREQ);
    const dataRange = equipmentSheet.getRange(startRow, COL_EQUIP_MAINT_FREQ, numRows, 2);
    const dataValues = dataRange.getValues();
    const datesToWrite = [];

    Logger.log(`Bắt đầu tính Ngày BT Tiếp theo cho ${numRows} dòng từ ${startRow}...`);

    // Xử lý từng dòng
    for (let i = 0; i < numRows; i++) {
      const currentRowIndex = startRow + i;
      if (currentRowIndex === 1) { 
        datesToWrite.push([null]);
        continue;
      }

      const maintFreq = dataValues[i][0];
      const lastMaintDateRaw = dataValues[i][1];
      let nextMaintDate = null;

      const lastMaintDate = (lastMaintDateRaw instanceof Date) ? lastMaintDateRaw : null;
      const freqStr = maintFreq ? maintFreq.toString().trim() : "";

      Logger.log(` > Dòng ${currentRowIndex}: Ngày cuối='${lastMaintDateRaw}', Tần suất='${freqStr}'`);

      if (lastMaintDate && freqStr !== "") {
        try {
          nextMaintDate = calculateNextMaintenanceDate(lastMaintDate, freqStr);
          if (nextMaintDate instanceof Date) {
            Logger.log(`  >> Tính được Ngày tiếp theo: ${nextMaintDate.toLocaleDateString()}`);
            calculatedCount++;
          } else {
            Logger.log(`  >> Không tính được Ngày tiếp theo (Tần suất '${freqStr}' có thể không hợp lệ). Giữ nguyên/Xóa cột S.`);
            skippedCount++;
            nextMaintDate = null;
          }
        } catch (calcErr) { 
          Logger.log(`  >> Lỗi khi tính ngày: ${calcErr}`);
          errorCount++;
          nextMaintDate = null;
        }
      } else { 
        Logger.log(`  >> Thiếu Ngày cuối hoặc Tần suất. Bỏ qua tính toán.`);
        skippedCount++;
        nextMaintDate = null;
      }
      datesToWrite.push([nextMaintDate]);
    }

    // Xác định vùng ghi dữ liệu
    const firstDataRowIndexInLoop = (startRow === 1) ? 1 : 0;
    const finalDatesToWrite = datesToWrite.slice(firstDataRowIndexInLoop);
    const firstDataSheetRow = Math.max(2, startRow);
    const numDataRowsToWrite = finalDatesToWrite.length;

    // Ghi dữ liệu vào sheet
    if (numDataRowsToWrite > 0) {
      const targetRange = equipmentSheet.getRange(firstDataSheetRow, COL_EQUIP_MAINT_NEXT, numDataRowsToWrite, 1);
      targetRange.setValues(finalDatesToWrite);
      Logger.log(`Đã cập nhật ${numDataRowsToWrite} dòng cho Cột Ngày BT Tiếp theo (S).`);
    }

    // Thông báo kết quả
    let message = `Hoàn thành:\n- Tính và cập nhật Ngày BT Tiếp theo cho ${calculatedCount} dòng.`;
    if (skippedCount > 0) { 
      message += `\n- Bỏ qua ${skippedCount} dòng do thiếu thông tin hoặc tần suất không hợp lệ.`;
    }
    if (errorCount > 0) { 
      message += `\n- Gặp ${errorCount} lỗi khi tính toán.`;
    }
    ui.alert(message);

  } catch (e) {
    Logger.log(`Lỗi nghiêm trọng trong calculateAndUpdateNextMaintDates: ${e} \nStack: ${e.stack}`);
    ui.alert(`Đã xảy ra lỗi: ${e}. Vui lòng kiểm tra Nhật ký thực thi.`);
  }
}

// =============================================
// NHÓM CHỨC NĂNG: TRA CỨU VÀ HIỂN THỊ DỮ LIỆU
// =============================================

/**
 * Hiển thị Sidebar cho chức năng tra cứu Thiết bị Con theo Thiết bị Cha.
 * Hàm này được gọi từ Menu "👨‍👦 Tra cứu Thiết bị Con".
 */
function showParentChildSearchView() {
  try {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('ParentChildSearch')
        .setTitle('Tra cứu Thiết bị Cha-Con')
        .setWidth(400);
    
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch (e) {
    Logger.log(`Lỗi khi hiển thị Sidebar tra cứu Cha-Con: ${e}`);
    SpreadsheetApp.getUi().alert(`Không thể mở giao diện tra cứu cha-con: ${e.message}`);
  }
}

/**
 * Hiển thị Sidebar cho chức năng tìm kiếm theo Vị Trí.
 * Hàm này được gọi từ Menu "🏢 Tra cứu theo Vị trí".
 */
function showLocationSearchView() {
  try {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('LocationSearch')
        .setTitle('Tra cứu theo Vị trí')
        .setWidth(400);
    
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch (e) {
    Logger.log(`Lỗi khi hiển thị Sidebar tra cứu Vị trí: ${e}`);
    SpreadsheetApp.getUi().alert(`Không thể mở giao diện tra cứu Vị trí: ${e.message}`);
  }
}

/**
 * Hiển thị prompt để nhập mã cần tra cứu lịch sử.
 * Kiểm tra ô đang chọn trước, chỉ hiện prompt nếu ô trống.
 * Hàm này được gọi từ Menu "🔎 Tra cứu Lịch sử Bảo trì / Sửa chữa".
 */
function getTargetForHistorySearch() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Kiểm tra ô đang chọn trước
    const activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
    let targetCode = "";
    
    if (activeCell && activeCell.getValue()) {
      // Lấy giá trị từ ô đang chọn
      const cellValue = activeCell.getValue();
      
      // Trích xuất mã nếu cần
      if (typeof cellValue === 'string' && cellValue.includes(" - ")) {
        targetCode = cellValue.split(" - ")[0].trim();
      } else {
        targetCode = cellValue.toString().trim();
      }
      
      // Nếu có mã hợp lệ, hiển thị sidebar ngay
      if (targetCode) {
        showHistorySidebar(targetCode);
        return;
      }
    }
    
    // Nếu ô đang chọn trống hoặc không lấy được mã, hiện prompt
    const result = ui.prompt(
      'Tra cứu Lịch sử Bảo trì / Sửa chữa',
      'Nhập Mã Thiết Bị hoặc Mã Hệ thống:',
      ui.ButtonSet.OK_CANCEL
    );
    
    const button = result.getSelectedButton();
    targetCode = result.getResponseText().trim();
    
    if (button === ui.Button.OK && targetCode) {
      showHistorySidebar(targetCode);
    } else if (button === ui.Button.OK) {
      ui.alert("Vui lòng nhập Mã Thiết Bị hoặc Mã Hệ thống hợp lệ.");
    }
  } catch (e) {
    Logger.log(`Lỗi trong getTargetForHistorySearch: ${e}`);
    ui.alert(`Lỗi khi tìm kiếm lịch sử: ${e.message}`);
  }
}


/**
 * Hiển thị sidebar lịch sử cho mã được chỉ định.
 * @param {string} targetCode Mã Thiết Bị hoặc Mã Hệ thống.
 */
function showHistorySidebar(targetCode) {
  try {
    // Trích xuất mã từ chuỗi nếu cần
    let cleanCode = targetCode;
    if (typeof targetCode === 'string' && targetCode.includes(" - ")) {
      cleanCode = targetCode.split(" - ")[0].trim();
    }
    
    // Tạo HTML template và truyền mã thiết bị vào template
    const htmlTemplate = HtmlService.createTemplateFromFile('SidebarHistory');
    htmlTemplate.targetCode = cleanCode; // Truyền mã vào template
    
    // Tạo sidebar
    const htmlOutput = htmlTemplate.evaluate()
        .setTitle(`Lịch sử: ${cleanCode}`)
        .setWidth(450);
    
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    
  } catch (e) {
    Logger.log(`Lỗi trong showHistorySidebar: ${e}`);
    SpreadsheetApp.getUi().alert(`Không thể hiển thị lịch sử: ${e.message}`);
  }
}


/**
 * Lấy dữ liệu lịch sử cho sidebar.
 * @return {Array} Mảng chứa [dữ liệu lịch sử, mã đối tượng].
 */
function getHistoryForSidebar() {
  try {
    // Lấy sidebar hiện tại để xác định tiêu đề (chứa Mã TB/HT)
    const ui = SpreadsheetApp.getUi();
    const sidebar = HtmlService.createHtmlOutput().getTitle(); // Tiêu đề: "Lịch sử: XYZ"
    
    // Parse mã từ tiêu đề
    let targetCode = "";
    if (sidebar && sidebar.startsWith("Lịch sử: ")) {
      targetCode = sidebar.substring(9).trim();
    } else {
      // Nếu không lấy được từ tiêu đề, kiểm tra ô đang chọn
      const activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
      if (activeCell) {
        const value = activeCell.getValue();
        if (value) {
          if (typeof value === 'string' && value.includes(" - ")) {
            targetCode = value.split(" - ")[0].trim();
          } else {
            targetCode = value.toString().trim();
          }
        }
      }
    }
    
    // Nếu vẫn không có mã hợp lệ, trả về mảng rỗng
    if (!targetCode) {
      return [[], "Không xác định"];
    }
    
    // Lấy dữ liệu lịch sử từ hàm getMaintenanceHistory
    const historyData = getMaintenanceHistory(targetCode);
    
    // Trả về tuple [dữ liệu, mã]
    return [historyData, targetCode];
    
  } catch (e) {
    Logger.log(`Lỗi trong getHistoryForSidebar: ${e}`);
    return [[], "Lỗi: " + e.message];
  }
}

/**
 * Mở dialog hiển thị lịch sử cho mã được chỉ định.
 * @param {string} targetCode Mã Thiết Bị hoặc Mã Hệ thống.
 */
function openHistoryDialogForCode(targetCode) {
  try {
    // Lấy dữ liệu lịch sử
    const historyData = getMaintenanceHistory(targetCode);
    
    // Tạo template và truyền dữ liệu
    const htmlTemplate = HtmlService.createTemplateFromFile('HistoryDialogContent');
    htmlTemplate.historyData = historyData;
    
    // Render template và tạo dialog
    const htmlOutput = htmlTemplate.evaluate()
        .setWidth(800)
        .setHeight(500);
    
    // Hiển thị dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Lịch sử: ${targetCode}`);
    
    return { success: true };
    
  } catch (e) {
    Logger.log(`Lỗi trong openHistoryDialogForCode: ${e}`);
    return { success: false, error: e.message };
  }
}

// =============================================
// NHÓM CHỨC NĂNG: QUẢN LÝ PHIẾU CÔNG VIỆC
// =============================================

/**
 * Tạo Installable trigger cho chức năng onEditWithAuth (thay thế onEdit đơn giản).
 * Chỉ cần chạy một lần bởi admin.
 */
function createEditTrigger() {
  try {
    // Xóa các trigger cũ nếu có để tránh trùng lặp
    const allTriggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getHandlerFunction() === 'onEditWithAuth') {
        ScriptApp.deleteTrigger(allTriggers[i]);
      }
    }
    
    // Tạo installable trigger mới
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ScriptApp.newTrigger('onEditWithAuth')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
    
    // Thông báo thành công
    Logger.log("Đã tạo Installable trigger 'onEditWithAuth' thành công");
    SpreadsheetApp.getUi().alert("✅ Đã cài đặt trigger thành công! Giờ đây email người dùng sẽ được tự động điền vào cột Người tạo.");
    
    return "Success";
  } catch (error) {
    Logger.log("Lỗi khi tạo trigger: " + error);
    SpreadsheetApp.getUi().alert("❌ Lỗi: " + error + "\nVui lòng đảm bảo bạn có quyền Admin và thử lại.");
    return "Error: " + error;
  }
}

/**
 * Phiên bản onEdit có quyền đầy đủ (AuthMode.FULL)
 * Được gọi bởi installable trigger đã tạo.
 * @param {Object} e Đối tượng sự kiện onEdit
 */
function onEditWithAuth(e) {
  try {
    // Kiểm tra sự kiện hợp lệ
    if (!e || !e.range || e.value === undefined) {
      return;
    }

    const sheet = e.range.getSheet();
    const sheetName = sheet.getName();
    const editedCol = e.range.getColumn();
    const editedRow = e.range.getRow();

    // Chỉ xử lý khi chỉnh sửa cột F (Đối tượng/HT) của sheet Phiếu Công Việc, và không phải hàng tiêu đề
    if (sheetName === SHEET_PHIEU_CONG_VIEC && editedCol === COL_PCV_DOI_TUONG && editedRow > 1) {
      const targetCodeRaw = e.value;
      let targetCode = "";

      // Trích xuất Mã từ giá trị
      if (targetCodeRaw && typeof targetCodeRaw === 'string') {
        if (targetCodeRaw.includes(" - ")) {
          targetCode = targetCodeRaw.split(" - ")[0].trim();
        } else {
          targetCode = targetCodeRaw.trim();
        }
      } else if (targetCodeRaw) {
        targetCode = targetCodeRaw.toString().trim();
      }

      Logger.log(`onEditWithAuth: Xử lý ${sheetName}, ô ${e.range.getA1Notation()}. Giá trị gốc="${targetCodeRaw}", Mã trích xuất="${targetCode}"`);

      if (targetCode) {
        // Tải dữ liệu tra cứu
        Logger.log("onEditWithAuth: Đang tải dữ liệu tra cứu TB và HT...");
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
        
        Logger.log(`onEditWithAuth: Đã tải ${Object.keys(equipmentMap).length} TB, ${Object.keys(systemMap).length} HT.`);

        // Tra cứu thông tin
        let targetName = "";
        let targetLocation = "";

        if (equipmentMap[targetCode]) {
          targetName = equipmentMap[targetCode].name;
          targetLocation = equipmentMap[targetCode].location;
          Logger.log(`onEditWithAuth: Tìm thấy TB. Tên="${targetName}", Vị trí="${targetLocation}"`);
        } else if (systemMap[targetCode]) {
          targetName = systemMap[targetCode];
          targetLocation = "N/A";
          Logger.log(`onEditWithAuth: Tìm thấy HT. Mô tả="${targetName}"`);
        } else {
          targetName = "Mã không hợp lệ";
          targetLocation = "";
          Logger.log(`onEditWithAuth: Mã trích xuất "${targetCode}" không hợp lệ.`);
        }

        // Cập nhật cột G và H
        const nameCell = sheet.getRange(editedRow, COL_PCV_TEN_DOI_TUONG);
        const locationCell = sheet.getRange(editedRow, COL_PCV_VI_TRI);
        
        if (nameCell.getValue() != targetName) {
          nameCell.setValue(targetName);
          Logger.log(`onEditWithAuth: Đã cập nhật Cột G thành "${targetName}"`);
        }
        
        if (locationCell.getValue() != targetLocation) {
          locationCell.setValue(targetLocation);
          Logger.log(`onEditWithAuth: Đã cập nhật Cột H thành "${targetLocation}"`);
        }

        // Kiểm tra và tạo Mã Phiếu CV nếu cột A trống
        const woIdCell = sheet.getRange(editedRow, COL_PCV_MA_PHIEU);
        if (!woIdCell.getValue()) {
          const newWoId = generateWorkOrderId();
          if (newWoId) {
            woIdCell.setValue(newWoId);
            sheet.getRange(editedRow, COL_PCV_NGAY_TAO).setValue(new Date());
            
            // Đoạn code đã sửa - Tận dụng AuthMode.FULL để lấy email người dùng
            const userEmail = Session.getActiveUser().getEmail();
            sheet.getRange(editedRow, COL_PCV_NGUOI_TAO).setValue(userEmail);
            Logger.log(`onEditWithAuth: Email người tạo = "${userEmail}"`);
            
            try {
              woIdCell.setFontSize(12).setVerticalAlignment("middle").setHorizontalAlignment("center")
                  .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
              // Bỏ dòng định dạng cột B (Ngày tạo)
              sheet.getRange(editedRow, COL_PCV_NGUOI_TAO).setFontSize(12).setVerticalAlignment("middle")
                  .setHorizontalAlignment("left").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
            } catch(fmtErr) {
              Logger.log(`onEditWithAuth: Lỗi định dạng ô A, C dòng ${editedRow}: ${fmtErr}`);
            }
            
            Logger.log(`onEditWithAuth: Đã tạo Mã Phiếu CV "${newWoId}" và điền thông tin cột B, C cho dòng ${editedRow}`);
          } else {
            Logger.log(`onEditWithAuth: Lỗi khi tạo Mã Phiếu CV cho dòng ${editedRow}`);
            woIdCell.setValue("LỖI TẠO MÃ PCV");
          }
        }
      } else {
        // Xóa dữ liệu khi ô F trống
        sheet.getRange(editedRow, COL_PCV_TEN_DOI_TUONG).clearContent();
        sheet.getRange(editedRow, COL_PCV_VI_TRI).clearContent();
        Logger.log(`onEditWithAuth: Đã xóa trống cột G, H do cột F trống/không hợp lệ - Dòng ${editedRow}`);
      }
    }
  } catch (err) {
    Logger.log(`Lỗi trong onEditWithAuth trigger: ${err}\nStack: ${err.stack}\nEvent Object: ${JSON.stringify(e)}`);
  }
}



/**
 * Simple trigger onEdit - chuyển hướng xử lý sang onEditWithAuth
 * Giữ lại để tương thích ngược khi chưa cài đặt trigger
 * @param {Object} e Đối tượng sự kiện onEdit.
 */
function onEdit(e) {
  try {
    // Kiểm tra xem installable trigger đã được cài đặt chưa
    const allTriggers = ScriptApp.getProjectTriggers();
    let hasInstallableTrigger = false;
    
    for (let i = 0; i < allTriggers.length; i++) {
      if (allTriggers[i].getHandlerFunction() === 'onEditWithAuth') {
        hasInstallableTrigger = true;
        break;
      }
    }
    
    if (!hasInstallableTrigger) {
      // Nếu chưa có installable trigger, vẫn chạy logic cũ
      // nhưng không lấy email (vì có thể không có quyền)
      Logger.log("Chưa cài đặt installable trigger. Chạy onEdit với các tính năng giới hạn.");
      
      // Kiểm tra sự kiện hợp lệ
      if (!e || !e.range || e.value === undefined) {
        return;
      }

      const sheet = e.range.getSheet();
      const sheetName = sheet.getName();
      const editedCol = e.range.getColumn();
      const editedRow = e.range.getRow();

      // Chỉ xử lý khi chỉnh sửa cột F (Đối tượng/HT) của sheet Phiếu Công Việc, và không phải hàng tiêu đề
      if (sheetName === SHEET_PHIEU_CONG_VIEC && editedCol === COL_PCV_DOI_TUONG && editedRow > 1) {
        const targetCodeRaw = e.value;
        let targetCode = "";

        // Trích xuất Mã từ giá trị
        if (targetCodeRaw && typeof targetCodeRaw === 'string') {
          if (targetCodeRaw.includes(" - ")) {
            targetCode = targetCodeRaw.split(" - ")[0].trim();
          } else {
            targetCode = targetCodeRaw.trim();
          }
        } else if (targetCodeRaw) {
          targetCode = targetCodeRaw.toString().trim();
        }

        Logger.log(`onEdit: Xử lý ${sheetName}, ô ${e.range.getA1Notation()}. Giá trị gốc="${targetCodeRaw}", Mã trích xuất="${targetCode}"`);

        if (targetCode) {
          // Tải dữ liệu tra cứu
          Logger.log("onEdit: Đang tải dữ liệu tra cứu TB và HT...");
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
                if (idStr) equipmentMap[idStr] = { 
                  name: row[COL_EQUIP_NAME - 1] || 'N/A', 
                  location: row[COL_EQUIP_LOCATION - 1] || 'N/A' 
                };
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
                if (codeStr) systemMap[codeStr] = row[COL_HT_MO_TA - 1] || "";
              }
            });
          }
          
          Logger.log(`onEdit: Đã tải ${Object.keys(equipmentMap).length} TB, ${Object.keys(systemMap).length} HT.`);

          // Tra cứu thông tin
          let targetName = "";
          let targetLocation = "";

          if (equipmentMap[targetCode]) {
            targetName = equipmentMap[targetCode].name;
            targetLocation = equipmentMap[targetCode].location;
          } else if (systemMap[targetCode]) {
            targetName = systemMap[targetCode];
            targetLocation = "N/A";
          } else {
            targetName = "Mã không hợp lệ";
            targetLocation = "";
          }

          // Cập nhật cột G và H
          const nameCell = sheet.getRange(editedRow, COL_PCV_TEN_DOI_TUONG);
          const locationCell = sheet.getRange(editedRow, COL_PCV_VI_TRI);
          
          if (nameCell.getValue() != targetName) nameCell.setValue(targetName);
          if (locationCell.getValue() != targetLocation) locationCell.setValue(targetLocation);

          // Kiểm tra và tạo Mã Phiếu CV nếu cột A trống
          const woIdCell = sheet.getRange(editedRow, COL_PCV_MA_PHIEU);
          if (!woIdCell.getValue()) {
            const newWoId = generateWorkOrderId();
            if (newWoId) {
              woIdCell.setValue(newWoId);
              sheet.getRange(editedRow, COL_PCV_NGAY_TAO).setValue(new Date());
              
              // KHÁC BIỆT: KHÔNG cố gắng đặt email người dùng vì simple trigger có thể không có quyền
              // Thay vào đó, đặt giá trị mặc định
              sheet.getRange(editedRow, COL_PCV_NGUOI_TAO).setValue("Người dùng hệ thống");
              
              try {
                woIdCell.setFontSize(12).setVerticalAlignment("middle").setHorizontalAlignment("center")
                    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
                sheet.getRange(editedRow, COL_PCV_NGUOI_TAO).setFontSize(12).setVerticalAlignment("middle")
                    .setHorizontalAlignment("left").setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
              } catch(fmtErr) {
                Logger.log(`onEdit: Lỗi định dạng ô A, C dòng ${editedRow}: ${fmtErr}`);
              }
            } else {
              woIdCell.setValue("LỖI TẠO MÃ PCV");
            }
          }
        } else {
          // Xóa dữ liệu khi ô F trống
          sheet.getRange(editedRow, COL_PCV_TEN_DOI_TUONG).clearContent();
          sheet.getRange(editedRow, COL_PCV_VI_TRI).clearContent();
        }
      }
    } else {
      // Nếu đã có installable trigger, không làm gì cả
      // vì onEditWithAuth sẽ được gọi tự động
      return;
    }
  } catch (err) {
    Logger.log(`Lỗi trong onEdit trigger: ${err}\nStack: ${err.stack}`);
  }
}


/**
 * Hàm được gọi từ Menu "✅ Hoàn thành Phiếu CV & Lưu Lịch sử".
 * Lấy dữ liệu từ dòng Phiếu CV được chọn và hiển thị Dialog nhập chi tiết hoàn thành.
 * Đã bổ sung quản lý đơn vị ngoài/NCC.
 */
function showCompletionDialog() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const workOrderSheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);

    // Kiểm tra vị trí hiện tại
    if (!workOrderSheet || ss.getActiveSheet().getName() !== SHEET_PHIEU_CONG_VIEC) {
      throw new Error(`Vui lòng chọn dòng cần hoàn thành trên sheet "${SHEET_PHIEU_CONG_VIEC}".`);
    }

    // Kiểm tra vùng đang chọn
    const selectedRange = workOrderSheet.getActiveRange();
    if (!selectedRange || selectedRange.getNumRows() !== 1 || selectedRange.getRow() < 2) {
      throw new Error("Vui lòng chọn chính xác một dòng Phiếu Công Việc (không phải dòng tiêu đề) để hoàn thành.");
    }
    
    const rowIndex = selectedRange.getRow();
    Logger.log(`showCompletionDialog: Chuẩn bị hiển thị dialog cho dòng ${rowIndex}`);

    // Đọc dữ liệu từ dòng được chọn - bổ sung đọc cột Chi tiết ĐV Ngoài
    const lastColToReadInitial = Math.max(COL_PCV_CHI_PHI, COL_PCV_CHI_TIET_NGOAI);
    const rowData = workOrderSheet.getRange(rowIndex, 1, 1, lastColToReadInitial).getValues()[0];

    // Chuẩn bị dữ liệu cho dialog
    const initialData = {
      rowIndex: rowIndex,
      maPhieuCV: rowData[COL_PCV_MA_PHIEU - 1],
      doiTuong: rowData[COL_PCV_DOI_TUONG - 1],
      tenDoiTuong: rowData[COL_PCV_TEN_DOI_TUONG - 1],
      viTri: rowData[COL_PCV_VI_TRI - 1],
      loaiCV: rowData[COL_PCV_LOAI_CV - 1],
      moTaYC: rowData[COL_PCV_MO_TA_YC - 1],
      nguoiGiao: rowData[COL_PCV_NGUOI_GIAO - 1],
      // Thông tin hoàn thành đã nhập trước (nếu có)
      moTaHT: rowData[COL_PCV_MO_TA_HT - 1],
      vatTu: rowData[COL_PCV_VAT_TU - 1],
      ngayHTTT: rowData[COL_PCV_NGAY_HT_THUC_TE - 1],
      trangThaiTBSau: rowData[COL_PCV_TRANG_THAI_TB_SAU - 1],
      chiPhi: rowData[COL_PCV_CHI_PHI - 1],
      // THÊM: Thông tin đơn vị ngoài
      externalVendorDetails: rowData[COL_PCV_CHI_TIET_NGOAI - 1] || ""
    };

    // Lấy danh sách trạng thái TB sau HĐ từ sheet Cấu hình (giữ nguyên code hiện tại)
    const configSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    let assetStatusOptions = [];
    const statusColIndex = COL_SETTINGS_ASSET_POST_STATUS_LIST_COL;

    if (configSheet && statusColIndex > 0) {
      // [Code đọc danh sách trạng thái giữ nguyên]
      try {
        const lastRowConfig = configSheet.getLastRow();
        if (lastRowConfig >= 2) {
          const fullColumnValues = configSheet.getRange(2, statusColIndex, lastRowConfig - 1, 1).getValues();
          assetStatusOptions = fullColumnValues.flat()
            .filter(value => value && value.toString().trim() !== "");
          Logger.log(`Đã lấy được ${assetStatusOptions.length} tùy chọn Trạng thái TB sau HĐ: [${assetStatusOptions.join(', ')}]`);
        } else {
          Logger.log(`Sheet ${SETTINGS_SHEET_NAME} không có dữ liệu Trạng thái TB sau HĐ (từ hàng 2).`);
          assetStatusOptions = [];
        }
      } catch (e) {
        Logger.log("Lỗi khi lấy danh sách Trạng thái TB sau HĐ từ Cấu hình: " + e);
        assetStatusOptions = [];
      }
    } else {
      if (!configSheet) 
        Logger.log(`Không tìm thấy sheet "${SETTINGS_SHEET_NAME}" để lấy danh sách trạng thái.`);
      if (!statusColIndex || statusColIndex <= 0) 
        Logger.log("Hằng số COL_SETTINGS_ASSET_POST_STATUS_LIST_COL không hợp lệ trong Config.gs.");
      assetStatusOptions = [];
    }

    // THÊM: Lấy danh sách đơn vị ngoài
    let vendorOptions = "";
    try {
      vendorOptions = getVendorOptionsHtml();
      Logger.log("Đã lấy danh sách đơn vị ngoài cho dialog");
    } catch (vendorErr) {
      Logger.log(`Lỗi khi lấy danh sách đơn vị ngoài: ${vendorErr}`);
      // Nếu chưa có hàm getVendorOptionsHtml, tạo danh sách trống
      vendorOptions = "";
    }

    // Tạo HTML dialog từ template
    const htmlTemplate = HtmlService.createTemplateFromFile('CompleteWorkOrderDialog');
    htmlTemplate.workOrderData = initialData;
    htmlTemplate.statusOptions = assetStatusOptions;
    htmlTemplate.vendorOptions = vendorOptions; // THÊM: Truyền danh sách đơn vị ngoài

    const htmlOutput = htmlTemplate.evaluate()
          .setWidth(650)  // Tăng kích thước để hiển thị tốt hơn
          .setHeight(600);
    
    const title = `Hoàn thành & Lưu Lịch sử cho Phiếu CV: ${initialData.maPhieuCV || '(Chưa có mã)'}`;

    // Hiển thị dialog
    ui.showModalDialog(htmlOutput, title);
    Logger.log(`Đã hiển thị Dialog hoàn thành cho dòng ${rowIndex}.`);

  } catch (e) {
    Logger.log(`Lỗi trong showCompletionDialog: ${e}`);
    ui.alert(`Lỗi mở hộp thoại hoàn thành: ${e.message}`);
  }
}


/**
 * Hàm được gọi từ Dialog Hoàn thành để lưu dữ liệu vào Lịch sử và cập nhật Phiếu CV.
 * Đã kiểm tra trùng lặp PCV và cập nhật thông tin đơn vị ngoài/NCC.
 * @param {object} completionData Dữ liệu người dùng nhập từ Dialog.
 * @return {object} Đối tượng báo thành công hoặc lỗi {success: boolean, message: string}.
 */
function saveHistoryFromDialog(completionData) {
  Logger.log(`saveHistoryFromDialog: Nhận dữ liệu: ${JSON.stringify(completionData)}`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    const workOrderSheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);

    // Kiểm tra tính khả dụng của sheet
    if (!historySheet || !workOrderSheet) {
      throw new Error("Không tìm thấy sheet Lịch sử hoặc Phiếu Công Việc.");
    }
    
    // Kiểm tra dữ liệu đầu vào
    if (!completionData || !completionData.rowIndex) {
      throw new Error("Dữ liệu gửi lên không hợp lệ (thiếu chỉ số dòng Phiếu CV).");
    }

    const woRowIndex = parseInt(completionData.rowIndex, 10);
    if (isNaN(woRowIndex) || woRowIndex < 2) {
      throw new Error("Chỉ số dòng Phiếu Công Việc không hợp lệ.");
    }

    // Lấy mã thiết bị/phiếu CV
    const targetCode = completionData.targetCode || "";
    
    // THÊM MỚI: Kiểm tra xem PCV đã tồn tại trong sheet Lịch sử hay chưa
    const historyData = historySheet.getDataRange().getValues();
    let existingRowIndex = -1;
    
    // Tìm kiếm targetCode trong cột COL_HISTORY_TARGET_CODE
    for (let i = 1; i < historyData.length; i++) {
      if (historyData[i][COL_HISTORY_TARGET_CODE - 1] === targetCode) {
        existingRowIndex = i + 1; // +1 vì index trong sheet bắt đầu từ 1
        break;
      }
    }
    
    // Nếu PCV đã tồn tại, hiển thị thông báo và hỏi người dùng
    if (existingRowIndex > 0) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'PCV đã tồn tại trong Lịch sử',
        `Mã PCV "${targetCode}" đã có bản ghi trong Lịch sử (dòng ${existingRowIndex}).\n\nBạn muốn cập nhật bản ghi hiện có thay vì tạo mới?`,
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        // Người dùng chọn cập nhật bản ghi hiện có
        return updateExistingHistoryRecord(historySheet, existingRowIndex, completionData, workOrderSheet, woRowIndex);
      }
      // Nếu chọn NO, tiếp tục tạo bản ghi mới như bình thường
    }

    // Chuẩn bị dữ liệu cho dòng lịch sử mới
    const historyRowData = [];
    historyRowData[COL_HISTORY_ID - 1] = null; // ID sẽ được tạo trong processNewHistoryRows
    historyRowData[COL_HISTORY_TARGET_CODE - 1] = targetCode;
    historyRowData[COL_HISTORY_TARGET_NAME - 1] = ""; // Sẽ được điền tự động
    historyRowData[COL_HISTORY_DISPLAY_NAME - 1] = ""; // Sẽ được điền tự động

    // Xử lý ngày hoàn thành
    let completionDate = null;
    if (completionData.completionDateStr) {
      try {
        const parts = completionData.completionDateStr.split('/');
        if (parts.length === 3) {
          completionDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          if (isNaN(completionDate.getTime())) {
            completionDate = null;
            Logger.log(`Ngày hoàn thành nhập vào không hợp lệ: ${completionData.completionDateStr}`);
          }
        }
      } catch(dateErr) {
        Logger.log(`Lỗi chuyển đổi ngày hoàn thành: ${dateErr}`);
        completionDate = null;
      }
    }
    
    if (!completionDate) {
      throw new Error("Ngày hoàn thành thực tế không hợp lệ hoặc bị thiếu.");
    }

    historyRowData[COL_HISTORY_EXEC_DATE - 1] = completionDate;
    historyRowData[COL_HISTORY_WORK_TYPE - 1] = completionData.workType || "";
    historyRowData[COL_HISTORY_DESCRIPTION - 1] = completionData.completionDesc || "";
    historyRowData[COL_HISTORY_PERFORMER - 1] = completionData.performer || "";
    
    // Lưu thông tin đơn vị ngoài
    historyRowData[COL_HISTORY_EXTERNAL_DETAILS - 1] = completionData.externalVendorDetails || "";
    
    historyRowData[COL_HISTORY_COST - 1] = completionData.cost || 0;
    
    // Xử lý thông tin bảo hành
    if (completionData.warrantyCheck) {
      // Nếu là bảo hành, cập nhật trạng thái và thông tin liên quan
      historyRowData[COL_HISTORY_STATUS - 1] = "Đang bảo hành";
      historyRowData[COL_HISTORY_WARRANTY_CHECK - 1] = true;
      historyRowData[COL_HISTORY_WARRANTY_REQ_ID - 1] = completionData.warrantyReqId || "";
      historyRowData[COL_HISTORY_WARRANTY_REQ_STAT - 1] = completionData.warrantyStatus || "Đã gửi yêu cầu";
      
      // Thêm thông tin NCC vào ghi chú bảo hành
      let warrantyNoteText = completionData.warrantyNote || "";
      if (completionData.warrantyVendorName) {
        if (warrantyNoteText) warrantyNoteText = "NCC: " + completionData.warrantyVendorName + "\n" + warrantyNoteText;
        else warrantyNoteText = "NCC: " + completionData.warrantyVendorName;
      }
      historyRowData[COL_HISTORY_WARRANTY_REQ_NOTE - 1] = warrantyNoteText;
    } else {
      // Nếu không phải bảo hành
      historyRowData[COL_HISTORY_STATUS - 1] = "Hoàn thành";
      historyRowData[COL_HISTORY_WARRANTY_CHECK - 1] = false;
      historyRowData[COL_HISTORY_WARRANTY_REQ_ID - 1] = "";
      historyRowData[COL_HISTORY_WARRANTY_REQ_STAT - 1] = "";
      historyRowData[COL_HISTORY_WARRANTY_REQ_NOTE - 1] = "";
    }
    
    historyRowData[COL_HISTORY_ASSET_POST_STATUS - 1] = completionData.assetStatus || "";
    historyRowData[COL_HISTORY_DETAIL_NOTE - 1] = completionData.detailNote || "";

    // Đảm bảo đủ số phần tử
    while (historyRowData.length < COL_HISTORY_DETAIL_NOTE) {
      historyRowData.push("");
    }

    // Ghi dòng mới vào sheet Lịch sử
    historySheet.appendRow(historyRowData);
    const newHistoryRowIndex = historySheet.getLastRow();
    Logger.log(`Đã thêm dòng lịch sử mới tại hàng ${newHistoryRowIndex}.`);
    
    // Khôi phục dropdown cho dòng mới
    try {
      copyDataValidationToNewRow(historySheet, 2, newHistoryRowIndex);
    } catch (dvErr) {
      Logger.log(`Lỗi khi sao chép data validation: ${dvErr}`);
    }
    
    SpreadsheetApp.flush();

    // Gọi processNewHistoryRows để hoàn thiện dòng lịch sử
    const newHistoryRowRange = historySheet.getRange(newHistoryRowIndex, 1);
    historySheet.setActiveRange(newHistoryRowRange);
    processNewHistoryRows(); // Hàm này sẽ tạo ID, điền các thông tin tự động
    Logger.log(`Đã chạy processNewHistoryRows cho dòng lịch sử mới ${newHistoryRowIndex}.`);
    SpreadsheetApp.flush();

    // Cập nhật thông tin đơn vị ngoài vào Phiếu CV
    if (completionData.externalVendorDetails) {
      workOrderSheet.getRange(woRowIndex, COL_PCV_CHI_TIET_NGOAI).setValue(completionData.externalVendorDetails);
      Logger.log(`Đã cập nhật thông tin đơn vị ngoài cho Phiếu CV dòng ${woRowIndex}.`);
    }

    // Cập nhật lại sheet Phiếu Công Việc
    Logger.log(`Bắt đầu cập nhật lại Phiếu CV dòng ${woRowIndex}...`);

    // Đọc ID Lịch sử vừa tạo
    let newHistoryId = "";
    try {
      newHistoryId = historySheet.getRange(newHistoryRowIndex, COL_HISTORY_ID).getDisplayValue();
      if (!newHistoryId) {
        newHistoryId = "Xem LS";
      }
    } catch(readIdErr) {
      Logger.log(`Lỗi nhỏ khi đọc lại ID lịch sử mới tạo: ${readIdErr}`);
      newHistoryId = "Xem LS";
    }

    // Tạo URL fragment trỏ đến ô ID Lịch sử
    const historySheetId = historySheet.getSheetId();
    const historyLinkUrl = `#gid=${historySheetId}&range=A${newHistoryRowIndex}`;
    Logger.log(`Generated history link URL: ${historyLinkUrl}`);

    // Tạo công thức HYPERLINK
    const linkFormula = `=HYPERLINK("${historyLinkUrl}"; "${newHistoryId.replace(/"/g, '""')}")`;
    Logger.log(`Generated history link formula: ${linkFormula}`);

    // Cập nhật Phiếu CV với trạng thái và link
    const woStatusCell = workOrderSheet.getRange(woRowIndex, COL_PCV_TRANG_THAI);
    const woLinkCell = workOrderSheet.getRange(woRowIndex, COL_PCV_LINK_LS);

    woStatusCell.setValue("Đã Lưu LS");
    woLinkCell.setFormula(linkFormula);

    // Định dạng ô link
    try {
      woLinkCell.setFontColor("#1155cc")
                .setFontLine("underline")
                .setFontSize(12)
                .setVerticalAlignment("middle")
                .setHorizontalAlignment("left")
                .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    } catch (fmtLinkErr) {
      Logger.log(`Lỗi định dạng ô Link Lịch sử (T${woRowIndex}): ${fmtLinkErr}`);
    }

    Logger.log(`Đã cập nhật trạng thái và CÔNG THỨC link lịch sử cho Phiếu CV dòng ${woRowIndex}.`);

    return { success: true, message: "Đã lưu lịch sử thành công!" };

  } catch (e) {
    Logger.log(`Lỗi trong saveHistoryFromDialog: ${e} \nStack: ${e.stack}`);
    return { success: false, message: `Lỗi lưu lịch sử: ${e.message}` };
  }
}

/**
 * Cập nhật bản ghi hiện có trong lịch sử thay vì tạo mới
 * @param {Object} historySheet Sheet lịch sử
 * @param {number} rowIndex Chỉ số dòng cần cập nhật
 * @param {Object} completionData Dữ liệu mới
 * @param {Object} workOrderSheet Sheet phiếu công việc
 * @param {number} woRowIndex Chỉ số dòng phiếu công việc
 * @return {Object} Kết quả cập nhật {success, message}
 */
function updateExistingHistoryRecord(historySheet, rowIndex, completionData, workOrderSheet, woRowIndex) {
  try {
    Logger.log(`Cập nhật bản ghi lịch sử hiện có tại dòng ${rowIndex}`);
    
    // Xử lý ngày hoàn thành
    let completionDate = null;
    if (completionData.completionDateStr) {
      try {
        const parts = completionData.completionDateStr.split('/');
        if (parts.length === 3) {
          completionDate = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          if (isNaN(completionDate.getTime())) {
            completionDate = null;
          }
        }
      } catch(dateErr) {
        Logger.log(`Lỗi chuyển đổi ngày hoàn thành: ${dateErr}`);
        completionDate = null;
      }
    }
    
    if (!completionDate) {
      return { success: false, message: "Ngày hoàn thành thực tế không hợp lệ." };
    }

    // Cập nhật các ô trong dòng hiện có
    historySheet.getRange(rowIndex, COL_HISTORY_EXEC_DATE).setValue(completionDate);
    historySheet.getRange(rowIndex, COL_HISTORY_WORK_TYPE).setValue(completionData.workType || "");
    historySheet.getRange(rowIndex, COL_HISTORY_DESCRIPTION).setValue(completionData.completionDesc || "");
    historySheet.getRange(rowIndex, COL_HISTORY_PERFORMER).setValue(completionData.performer || "");
    historySheet.getRange(rowIndex, COL_HISTORY_EXTERNAL_DETAILS).setValue(completionData.externalVendorDetails || "");
    historySheet.getRange(rowIndex, COL_HISTORY_COST).setValue(completionData.cost || 0);
    historySheet.getRange(rowIndex, COL_HISTORY_ASSET_POST_STATUS).setValue(completionData.assetStatus || "");
    
    // Xử lý thông tin bảo hành
    if (completionData.warrantyCheck) {
      historySheet.getRange(rowIndex, COL_HISTORY_STATUS).setValue("Đang bảo hành");
      historySheet.getRange(rowIndex, COL_HISTORY_WARRANTY_CHECK).setValue(true);
      historySheet.getRange(rowIndex, COL_HISTORY_WARRANTY_REQ_ID).setValue(completionData.warrantyReqId || "");
      historySheet.getRange(rowIndex, COL_HISTORY_WARRANTY_REQ_STAT).setValue(completionData.warrantyStatus || "Đã gửi yêu cầu");
      
      // Thêm thông tin NCC vào ghi chú bảo hành
      let warrantyNoteText = completionData.warrantyNote || "";
      if (completionData.warrantyVendorName) {
        if (warrantyNoteText) warrantyNoteText = "NCC: " + completionData.warrantyVendorName + "\n" + warrantyNoteText;
        else warrantyNoteText = "NCC: " + completionData.warrantyVendorName;
      }
      historySheet.getRange(rowIndex, COL_HISTORY_WARRANTY_REQ_NOTE).setValue(warrantyNoteText);
    } else {
      historySheet.getRange(rowIndex, COL_HISTORY_STATUS).setValue("Hoàn thành");
    }
    
    // Thêm dấu thời gian vào ghi chú
    const now = new Date();
    const currentNote = historySheet.getRange(rowIndex, COL_HISTORY_DETAIL_NOTE).getValue();
    const newNote = completionData.detailNote || "";
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    const updatedNote = newNote + (newNote ? "\n\n" : "") + 
                        "Cập nhật: " + timestamp + 
                        (currentNote ? "\n\nGhi chú trước:\n" + currentNote : "");
    
    historySheet.getRange(rowIndex, COL_HISTORY_DETAIL_NOTE).setValue(updatedNote);
    
    // Cập nhật thông tin đơn vị ngoài vào Phiếu CV
    if (completionData.externalVendorDetails) {
      workOrderSheet.getRange(woRowIndex, COL_PCV_CHI_TIET_NGOAI).setValue(completionData.externalVendorDetails);
    }
    
    Logger.log(`Đã cập nhật thành công bản ghi lịch sử tại dòng ${rowIndex}`);
    return { 
      success: true, 
      message: "Đã cập nhật bản ghi lịch sử hiện có thành công!"
    };
    
  } catch (error) {
    Logger.log(`Lỗi khi cập nhật bản ghi lịch sử: ${error}`);
    return { 
      success: false, 
      message: `Lỗi khi cập nhật bản ghi lịch sử: ${error.message}`
    };
  }
}


/**
 * Hàm sao chép data validation từ dòng mẫu sang dòng mới trong sheet Lịch sử
 * @param {Sheet} historySheet Sheet Lịch sử
 * @param {number} templateRow Dòng mẫu (thường là 2)
 * @param {number} newRow Dòng mới vừa được thêm
 */
function copyDataValidationToNewRow(historySheet, templateRow, newRow) {
  try {
    // Các cột cần sao chép data validation (B, F, H, K, N, P)
    const columnsNeedValidation = [
      COL_HISTORY_EXEC_DATE,      // B: Ngày thực hiện
      COL_HISTORY_WORK_TYPE,      // F: Loại công việc
      COL_HISTORY_PERFORMER,      // H: Người thực hiện
      COL_HISTORY_STATUS,         // K: Trạng thái
      COL_HISTORY_WARRANTY_REQ_STAT, // N: Trạng thái yêu cầu bảo hành
      COL_HISTORY_ASSET_POST_STATUS // P: Trạng thái TB sau HĐ
    ];
    
    // Nếu các hằng số chưa được định nghĩa, dùng số cột trực tiếp
    if (typeof COL_HISTORY_EXEC_DATE === 'undefined') {
      columnsNeedValidation = [2, 6, 8, 11, 14, 16]; // B, F, H, K, N, P
    }
    
    // Sao chép data validation từ dòng mẫu sang dòng mới
    for (const col of columnsNeedValidation) {
      const sourceCell = historySheet.getRange(templateRow, col);
      const targetCell = historySheet.getRange(newRow, col);
      
      const validation = sourceCell.getDataValidation();
      if (validation) {
        targetCell.setDataValidation(validation);
        Logger.log(`Đã sao chép data validation từ ô ${sourceCell.getA1Notation()} đến ô ${targetCell.getA1Notation()}`);
      }
    }
    
    Logger.log(`Đã khôi phục dropdown cho dòng ${newRow} trong sheet Lịch sử`);
    
  } catch (error) {
    Logger.log("Lỗi khi sao chép data validation: " + error);
  }
}


// =============================================
// NHÓM CHỨC NĂNG: ĐỒNG BỘ & BẢO TRÌ HỆ THỐNG
// =============================================

/**
 * Đồng bộ Hệ thống Cơ bản vào sheet DinhNghiaHeThong dựa trên sheet Cấu hình.
 * 1. Xóa các mã hệ thống cơ bản "mồ côi" (có tiền tố cơ bản, có mã vị trí nhưng vị trí đó không còn tồn tại).
 * 2. Thêm các mã hệ thống cơ bản còn thiếu cho các vị trí hợp lệ hiện có, dựa trên "Loại Vị Trí" để loại trừ các hệ thống không phù hợp.
 * KHÔNG XÓA các mã được nhập thủ công (không khớp với mẫu cơ bản).
 * Được gọi từ Menu.
 */
function syncBasicSystemsForNewLocations() {
  const ui = SpreadsheetApp.getUi();
  Logger.log("===== Bắt đầu syncBasicSystemsForNewLocations =====");
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
                 locations.push({ name: locName, acronym: locAcronym, type: locType || "" }); // Lưu cả Loại VT, nếu trống thì là chuỗi rỗng
                 validLocationAcronyms.add(locAcronym);
             }
        });
    }
    Logger.log(`Đã đọc ${locations.length} vị trí hợp lệ từ Cấu hình. Các Mã VT hợp lệ: ${Array.from(validLocationAcronyms).join(', ')}`);


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
    Logger.log(`Đã đọc ${systemDefData.length} mã hệ thống hiện có từ DinhNghiaHeThong.`);

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
          Logger.log(`Đánh dấu để xóa mã hệ thống mồ côi: ${system.code} (dòng ${system.rowNum}) - Mã vị trí '${locationCode}' không còn hợp lệ.`);
        }
      }
    });
    
    // Xóa các dòng theo thứ tự từ dưới lên để tránh shift index
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // Sắp xếp giảm dần
      
      for (const rowNum of rowsToDelete) {
        systemDefSheet.deleteRow(rowNum);
      }
      
      Logger.log(`Đã xóa ${rowsToDelete.length} mã hệ thống mồ côi.`);
      SpreadsheetApp.flush();
    } else {
      Logger.log("Không tìm thấy mã hệ thống mồ côi nào cần xóa.");
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
            Logger.log(`Tạo mã hệ thống mới: ${newSystemCode} - ${newSystemDesc}`);
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
        Logger.log(`Lỗi khi định dạng các dòng mới: ${formatErr}`);
      }
      
      Logger.log(`Đã thêm ${newSystemRows.length} mã hệ thống mới.`);
    } else {
      Logger.log("Không có mã hệ thống mới nào cần thêm.");
    }
    
    // --- 4. Sắp xếp sheet ---
    try {
      if (systemDefSheet.getLastRow() > 2) {
        systemDefSheet.getRange(2, 1, systemDefSheet.getLastRow() - 1, 2).sort({column: 1, ascending: true});
        Logger.log("Đã sắp xếp lại các mã hệ thống theo thứ tự A-Z.");
      }
    } catch (sortErr) {
      Logger.log(`Không thể sắp xếp sheet: ${sortErr}`);
    }
    
    // --- 5. Thông báo kết quả ---
    ui.alert(`Đồng bộ hoàn tất:\n- Đã xóa ${rowsToDelete.length} mã hệ thống mồ côi.\n- Đã thêm ${newSystemRows.length} mã hệ thống mới.`);
    
    Logger.log("===== Kết thúc syncBasicSystemsForNewLocations =====");
    
  } catch (e) {
    Logger.log(`Lỗi nghiêm trọng trong syncBasicSystemsForNewLocations: ${e} \nStack: ${e.stack}`);
    ui.alert(`Đã xảy ra lỗi: ${e.message}`);
  }
}


/**
 * Tạo Phiếu Công Việc Bảo trì Định kỳ tự động dựa trên lịch sử.
 * Hàm này được thiết kế để chạy như một trigger định kỳ.
 */
function createScheduledPmWorkOrders_v2() {
  const FUNCTION_NAME = "createScheduledPmWorkOrders_v2";
  Logger.log(`===== [${FUNCTION_NAME}] Bắt đầu chạy (Logic dựa trên Lịch sử) =====`);
  const daysAhead = 15; // Số ngày quét trước
  const pmWorkType = "Bảo trì Định kỳ"; // Giá trị Loại CV cho PM trong Lịch sử
  const defaultInitialStatus = "Đã lên kế hoạch"; // Trạng thái ban đầu cho Phiếu CV mới
  const defaultPriority = "Trung bình"; // Mức ưu tiên mặc định

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    const workOrderSheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
    const detailSheet = ss.getSheetByName(SHEET_CHI_TIET_CV_BT);
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);

    // Kiểm tra các sheet và hằng số cần thiết
    if (!equipSheet || !workOrderSheet || !detailSheet || !historySheet) {
      throw new Error(`Không tìm thấy một hoặc nhiều sheet cần thiết.`);
    }
    if (typeof COL_EQUIP_STATUS === 'undefined' || typeof EMAIL_ADMIN === 'undefined' || typeof COL_PCV_TAN_SUAT_PM === 'undefined') {
      throw new Error("Lỗi cấu hình: Thiếu khai báo hằng số COL_EQUIP_STATUS, EMAIL_ADMIN hoặc COL_PCV_TAN_SUAT_PM trong Config.gs.");
    }

    // --- 1. Đọc dữ liệu cần thiết ---
    Logger.log(`[${FUNCTION_NAME}] Đang đọc dữ liệu...`);

    // 1.1. Đọc Định nghĩa Công việc PM
    const pmTasks = {}; // { loaiTB: { tanSuat: { months: number, tasks: string }, ... }, ... }
    if (detailSheet.getLastRow() >= 2) {
      const taskData = detailSheet.getRange(2, COL_CTCV_LOAI_TB, detailSheet.getLastRow() - 1, COL_CTCV_CONG_VIEC).getValues();
      taskData.forEach(row => {
        const loaiTB = row[COL_CTCV_LOAI_TB - 1]?.toString().trim();
        const tanSuat = row[COL_CTCV_TAN_SUAT - 1]?.toString().trim();
        const congViec = row[COL_CTCV_CONG_VIEC - 1]?.toString().trim();
        const months = parseFrequencyToMonths(tanSuat);
        if (loaiTB && tanSuat && congViec && months !== null) {
          if (!pmTasks[loaiTB]) { pmTasks[loaiTB] = {}; }
          pmTasks[loaiTB][tanSuat] = { months: months, tasks: congViec };
        } else if(loaiTB && tanSuat) { 
          Logger.log(`[${FUNCTION_NAME}] Cảnh báo: Bỏ qua định nghĩa PM cho ${loaiTB} - ${tanSuat} do thiếu công việc hoặc tần suất không hợp lệ.`); 
        }
      });
      Logger.log(`[${FUNCTION_NAME}] Đã đọc định nghĩa PM cho ${Object.keys(pmTasks).length} loại TB.`);
    } else { 
      Logger.log(`[${FUNCTION_NAME}] Sheet ${SHEET_CHI_TIET_CV_BT} trống.`); 
    }

    // 1.2. Đọc Phiếu CV đang mở
const openWorkOrders = {}; // { "MaTB_TanSuat": true, ... }
if (workOrderSheet.getLastRow() >= 2) {
  const lastWoColCheck = Math.max(COL_PCV_DOI_TUONG, COL_PCV_LOAI_CV, COL_PCV_TAN_SUAT_PM, COL_PCV_TRANG_THAI);
  const woCheckData = workOrderSheet.getRange(2, 1, workOrderSheet.getLastRow() - 1, lastWoColCheck).getValues();
  
  // In ra log toàn bộ dữ liệu để debug
  Logger.log(`[${FUNCTION_NAME}] Đang phân tích ${woCheckData.length} phiếu CV trong sheet ${SHEET_PHIEU_CONG_VIEC}`);
  
  for (let i = 0; i < woCheckData.length; i++) {
    const row = woCheckData[i];
    const rowNum = i + 2;
    const rawTarget = row[COL_PCV_DOI_TUONG - 1]; // Dữ liệu gốc từ cột F
    const workType = row[COL_PCV_LOAI_CV - 1]?.toString().trim();
    const pmFrequency = row[COL_PCV_TAN_SUAT_PM - 1]?.toString().trim();
    const status = row[COL_PCV_TRANG_THAI - 1]?.toString().trim();
    
    // Log chi tiết từng dòng để debug
    Logger.log(`[${FUNCTION_NAME}] Dòng ${rowNum}: F="${rawTarget}", I="${workType}", J="${pmFrequency}", N="${status}"`);
    
    // Trích xuất mã thiết bị từ chuỗi
    let target = "";
    if (rawTarget) {
      if (typeof rawTarget === 'string' && rawTarget.includes(" - ")) {
        target = rawTarget.split(" - ")[0].trim();
      } else {
        target = rawTarget.toString().trim();
      }
      Logger.log(`[${FUNCTION_NAME}] >> Mã trích xuất: "${target}"`);
    }
    
    // Danh sách trạng thái được coi là đã đóng/hoàn thành
    const closedStatuses = ["Đã Lưu LS", "Hủy", "Hoàn thành", "Đã hoàn thành"];
    
    // Lưu ý: Kiểm tra "PM" vẫn dựa trên giá trị pmWorkType (thường là "Bảo trì Định kỳ")
    if (target && workType === pmWorkType && pmFrequency && !closedStatuses.includes(status)) {
      const key = `${target}_${pmFrequency}`;
      openWorkOrders[key] = true;
      Logger.log(`[${FUNCTION_NAME}] >> PHIẾU MỞ HỢP LỆ: TB=${target}, Tần suất=${pmFrequency}, Trạng thái=${status}`);
    }
  }
  
  Logger.log(`[${FUNCTION_NAME}] Đã tìm thấy ${Object.keys(openWorkOrders).length} Phiếu CV PM đang mở.`);
}

    // 1.3. Đọc và Xử lý Dữ liệu Lịch sử PM
    const lastPmCompletionMap = {}; // { maTB: Date, ... }
    if (historySheet.getLastRow() >= 2) {
      const lastHistColCheck = Math.max(COL_HISTORY_TARGET_CODE, COL_HISTORY_EXEC_DATE, COL_HISTORY_WORK_TYPE);
      const historyData = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, lastHistColCheck).getValues();
      Logger.log(`[${FUNCTION_NAME}] Đã đọc ${historyData.length} dòng từ ${HISTORY_SHEET_NAME}. Đang lọc và sắp xếp...`);
      
      const filteredHistory = historyData.filter(row => {
        const targetCode = extractTargetCode_(row[COL_HISTORY_TARGET_CODE - 1]);
        const workType = row[COL_HISTORY_WORK_TYPE - 1]?.toString().trim();
        const execDate = row[COL_HISTORY_EXEC_DATE - 1];
        return targetCode && workType === pmWorkType && execDate instanceof Date && !isNaN(execDate);
      });

      // Sắp xếp theo Mã Đối tượng, sau đó Ngày hoàn thành giảm dần
      filteredHistory.sort((a, b) => {
        const targetA = extractTargetCode_(a[COL_HISTORY_TARGET_CODE - 1]);
        const targetB = extractTargetCode_(b[COL_HISTORY_TARGET_CODE - 1]);
        const dateA = a[COL_HISTORY_EXEC_DATE - 1].getTime();
        const dateB = b[COL_HISTORY_EXEC_DATE - 1].getTime();
        
        if (targetA < targetB) return -1;
        if (targetA > targetB) return 1;
        return dateB - dateA; // Ngày giảm dần
      });
      
      Logger.log(`[${FUNCTION_NAME}] Đã lọc và sắp xếp còn ${filteredHistory.length} bản ghi Lịch sử PM hợp lệ.`);

      // Tạo Map lấy ngày hoàn thành PM cuối cùng cho mỗi thiết bị
      filteredHistory.forEach(row => {
        const targetCode = extractTargetCode_(row[COL_HISTORY_TARGET_CODE - 1]);
        if (targetCode && !lastPmCompletionMap[targetCode]) {
          lastPmCompletionMap[targetCode] = new Date(row[COL_HISTORY_EXEC_DATE - 1]);
          lastPmCompletionMap[targetCode].setHours(0,0,0,0);
        }
      });
      
      Logger.log(`[${FUNCTION_NAME}] Đã xác định Ngày HT PM cuối cùng cho ${Object.keys(lastPmCompletionMap).length} đối tượng.`);
    } else { 
      Logger.log(`[${FUNCTION_NAME}] Sheet ${HISTORY_SHEET_NAME} trống.`); 
    }

    // 1.4. Đọc Danh mục Thiết bị (chỉ lấy TB đang hoạt động)
    const equipmentList = [];
    if (equipSheet.getLastRow() >= 2) {
      const lastEquipCol = Math.max(COL_EQUIP_ID, COL_EQUIP_NAME, COL_EQUIP_TYPE, COL_EQUIP_LOCATION, COL_EQUIP_STATUS);
      const equipData = equipSheet.getRange(2, 1, equipSheet.getLastRow() - 1, lastEquipCol).getValues();
      equipData.forEach(row => {
        const status = row[COL_EQUIP_STATUS - 1]?.toString().trim();
        const id = row[COL_EQUIP_ID - 1]?.toString().trim();
        if (id && status === "Đang hoạt động") {
          equipmentList.push({
            id: id,
            name: row[COL_EQUIP_NAME - 1]?.toString().trim(),
            type: row[COL_EQUIP_TYPE - 1]?.toString().trim(),
            location: row[COL_EQUIP_LOCATION - 1]?.toString().trim()
          });
        }
      });
      Logger.log(`[${FUNCTION_NAME}] Đã đọc ${equipmentList.length} thiết bị đang hoạt động.`);
    } else { 
      Logger.log(`[${FUNCTION_NAME}] Sheet ${EQUIPMENT_SHEET_NAME} trống.`); 
    }

    // --- 2. Tính toán và Tạo Phiếu CV ---
    Logger.log(`[${FUNCTION_NAME}] Bắt đầu tính toán và tạo Phiếu CV PM...`);
    const today = new Date(); 
    today.setHours(0, 0, 0, 0);
    const targetDate = new Date(today); 
    targetDate.setDate(today.getDate() + daysAhead);
    Logger.log(`[${FUNCTION_NAME}] Ngưỡng ngày quét: Từ ${formatDate_(today)} đến ${formatDate_(targetDate)}.`);

    const newWorkOrders = [];

    // Lặp qua danh sách Thiết bị đang hoạt động
    equipmentList.forEach(equip => {
      if (!equip.id || !equip.type) { return; }

      // Lấy ngày PM cuối cùng của thiết bị từ map đã tạo
      const lastPmDate = lastPmCompletionMap[equip.id];

      if (!lastPmDate) {
        Logger.log(`[${FUNCTION_NAME}] TB ${equip.id}: Chưa có lịch sử PM hoặc chưa được xử lý. Bỏ qua.`);
        return;
      }
      
      Logger.log(`[${FUNCTION_NAME}] TB ${equip.id} (${equip.type}), PM cuối ngày: ${formatDate_(lastPmDate)}. Kiểm tra các tần suất...`);

      // Lấy tất cả định nghĩa PM cho Loại TB này
      const definedPms = pmTasks[equip.type];
      if (!definedPms) {
        Logger.log(`[${FUNCTION_NAME}] Không tìm thấy định nghĩa PM nào cho Loại TB: ${equip.type}. Bỏ qua TB ${equip.id}.`);
        return;
      }

      // Lặp qua từng tần suất PM được định nghĩa
      for (const frequency in definedPms) {
        const pmDefinition = definedPms[frequency];
        const monthsToAdd = pmDefinition.months;
        const detailedTasks = pmDefinition.tasks;

        if (monthsToAdd === null || monthsToAdd <= 0) continue;

        let dueDate;
        try { 
          dueDate = addMonthsToDate(lastPmDate, monthsToAdd); 
          dueDate.setHours(0,0,0,0); 
        } catch (e) { 
          Logger.log(`[${FUNCTION_NAME}] Lỗi tính dueDate cho ${equip.id}, tần suất ${frequency}: ${e.message}`); 
          continue; 
        }

        Logger.log(`  - Tần suất: ${frequency} (${monthsToAdd} tháng) -> Hạn HT dự kiến: ${formatDate_(dueDate)}`);

        // Kiểm tra ngưỡng ngày và phiếu đang mở
        if (dueDate >= today && dueDate <= targetDate) {
          Logger.log(`    >> Nằm trong ngưỡng quét!`);
          const openWoKey = `${equip.id}_${frequency}`;
          
          if (!openWorkOrders[openWoKey]) {
            Logger.log(`    >> Chưa có Phiếu CV mở. ---> Cần tạo Phiếu CV!`);

            const workOrderId = generateWorkOrderId();
            if (!workOrderId) { 
              Logger.log(`!!! Lỗi tạo Mã Phiếu CV cho ${equip.id}. Bỏ qua.`); 
              continue; 
            }

            // Chuẩn bị dữ liệu dòng mới
            const newWoRow = Array(COL_PCV_GHI_CHU).fill("");
            newWoRow[COL_PCV_MA_PHIEU - 1] = workOrderId;
            newWoRow[COL_PCV_NGAY_TAO - 1] = new Date();
            newWoRow[COL_PCV_NGUOI_TAO - 1] = "Auto Script PM";
            newWoRow[COL_PCV_HAN_HT - 1] = dueDate;
            newWoRow[COL_PCV_DOI_TUONG - 1] = equip.id;
            newWoRow[COL_PCV_TEN_DOI_TUONG - 1] = equip.name;
            newWoRow[COL_PCV_VI_TRI - 1] = equip.location;
            newWoRow[COL_PCV_LOAI_CV - 1] = pmWorkType;
            newWoRow[COL_PCV_TAN_SUAT_PM - 1] = frequency;
            newWoRow[COL_PCV_MO_TA_YC - 1] = detailedTasks;
            newWoRow[COL_PCV_UU_TIEN - 1] = defaultPriority;
            newWoRow[COL_PCV_TRANG_THAI - 1] = defaultInitialStatus;

            newWorkOrders.push(newWoRow);
            openWorkOrders[openWoKey] = true; // Đánh dấu đã tạo
          } else { 
            Logger.log(`    >> Đã có Phiếu CV mở cho tần suất ${frequency}. Bỏ qua.`); 
          }
        }
      }
    });

    // --- 3. Ghi các Phiếu CV mới vào sheet ---
    if (newWorkOrders.length > 0) {
      try {
        Logger.log(`[${FUNCTION_NAME}] Chuẩn bị ghi ${newWorkOrders.length} Phiếu CV PM mới...`);
        
        // Xác định vị trí ghi an toàn
        const lastRow = Math.max(1, workOrderSheet.getLastRow());
        
        // Kiểm tra độ dài mảng dữ liệu
        const expectedCols = COL_PCV_GHI_CHU;
        for (let i = 0; i < newWorkOrders.length; i++) {
          while (newWorkOrders[i].length < expectedCols) {
            newWorkOrders[i].push(""); // Đảm bảo đủ số cột
          }
        }
        
        Logger.log(`[${FUNCTION_NAME}] Vị trí ghi: hàng ${lastRow + 1}, ${newWorkOrders.length} dòng x ${newWorkOrders[0].length} cột`);
        workOrderSheet.getRange(lastRow + 1, 1, newWorkOrders.length, newWorkOrders[0].length)
          .setValues(newWorkOrders);
        
        SpreadsheetApp.flush(); // Đảm bảo dữ liệu được ghi ngay
        Logger.log(`[${FUNCTION_NAME}] Đã ghi thành công ${newWorkOrders.length} Phiếu CV PM mới.`);
      } catch (writeErr) {
        Logger.log(`[${FUNCTION_NAME}] LỖI GHI DỮ LIỆU: ${writeErr}\nStack: ${writeErr.stack}`);
        throw writeErr; // Ném lỗi để xử lý ở catch bên ngoài
      }
    } else { 
      Logger.log(`[${FUNCTION_NAME}] Không có Phiếu CV PM mới nào cần tạo trong lần chạy này.`); 
    }

    Logger.log(`===== [${FUNCTION_NAME}] Kết thúc =====`);
  } catch (e) {
    Logger.log(`!!!!!! [${FUNCTION_NAME}] LỖI NGHIÊM TRỌNG: ${e} \nStack: ${e.stack}`);
    // Gửi email báo lỗi
    try { 
      MailApp.sendEmail(EMAIL_ADMIN, `[Lỗi] Script Tạo Phiếu CV PM Tự Động`, `Chi tiết lỗi: ${e}\nStack: ${e.stack}`); 
    } catch (mailErr) { 
      Logger.log(`Lỗi gửi mail thông báo: ${mailErr}`);
    }
  }
}

/**
 * Hàm trợ giúp để trích xuất mã đối tượng từ chuỗi có thể chứa thêm tên
 * @param {string|any} rawValue Giá trị từ cột đối tượng
 * @return {string} Mã đã trích xuất hoặc chuỗi rỗng nếu không hợp lệ
 * @private
 */
function extractTargetCode_(rawValue) {
  if (!rawValue) return "";
  
  if (typeof rawValue === 'string') { 
    return rawValue.split(" - ")[0].trim(); 
  }
  return rawValue.toString().trim();
}

/**
 * Hàm trợ giúp để định dạng ngày theo chuẩn dd/MM/yyyy
 * @param {Date} date Đối tượng Date cần định dạng
 * @return {string} Chuỗi ngày đã định dạng
 * @private
 */
function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

/**
 * Thêm số tháng vào ngày và trả về ngày mới
 * @param {Date} date Ngày gốc
 * @param {number} months Số tháng cần thêm
 * @return {Date} Ngày mới sau khi thêm tháng
 */
function addMonthsToDate(date, months) {
  const newDate = new Date(date);
  newDate.setMonth(newDate.getMonth() + parseInt(months, 10));
  return newDate;
}

/**
 * Phân tích chuỗi tần suất và chuyển đổi thành số tháng
 * @param {string} frequencyStr Chuỗi tần suất (vd: "3 tháng", "1 năm", "Hằng tháng")
 * @return {number|null} Số tháng tương ứng hoặc null nếu không phân tích được
 */
function parseFrequencyToMonths(frequencyStr) {
  if (!frequencyStr) return null;
  
  const str = frequencyStr.toString().toLowerCase().trim();
  
  // Các trường hợp đặc biệt
  if (str === "hàng ngày" || str === "hằng ngày" || str === "mỗi ngày") return 0.033; // ~ 1 ngày
  if (str === "hàng tuần" || str === "hằng tuần" || str === "mỗi tuần") return 0.25; // ~ 1 tuần
  if (str === "nửa tháng" || str === "2 tuần" || str === "hai tuần") return 0.5; // ~ 2 tuần
  if (str === "hàng tháng" || str === "hằng tháng" || str === "mỗi tháng" || str === "1 tháng" || str === "một tháng") return 1;
  if (str === "quý" || str === "hàng quý" || str === "mỗi quý") return 3;
  if (str === "nửa năm" || str === "6 tháng" || str === "sáu tháng") return 6;
  if (str === "hàng năm" || str === "mỗi năm" || str === "1 năm" || str === "một năm") return 12;
  if (str === "hai năm" || str === "2 năm") return 24;
  
  // Tìm số từ chuỗi
  const numberMatch = str.match(/(\d+)/);
  if (numberMatch) {
    const number = parseInt(numberMatch[1], 10);
    if (str.includes("năm")) return number * 12;
    if (str.includes("tháng")) return number;
    if (str.includes("tuần")) return number * 0.25;
    if (str.includes("ngày")) return number * 0.033;
  }
  
  return null; // Không thể phân tích được
}


/**
 * Hàm trợ giúp để trích xuất mã đối tượng từ chuỗi có thể chứa thêm tên
 * @param {string|any} rawValue Giá trị từ cột đối tượng
 * @return {string} Mã đã trích xuất hoặc chuỗi rỗng nếu không hợp lệ
 * @private
 */
function extractTargetCode_(rawValue) {
  if (!rawValue) return "";
  
  if (typeof rawValue === 'string') { 
    return rawValue.split(" - ")[0].trim(); 
  }
  return rawValue.toString().trim();
}

/**
 * Hàm trợ giúp để định dạng ngày theo chuẩn dd/MM/yyyy
 * @param {Date} date Đối tượng Date cần định dạng
 * @return {string} Chuỗi ngày đã định dạng
 * @private
 */
function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

/**
 * Kiểm tra bảo hành cho thiết bị trên dòng đang chọn
 */
function checkCurrentEquipmentWarranty() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const activeRange = activeSheet.getActiveRange();
    
    if (!activeRange) {
      ui.alert("Vui lòng chọn một dòng thiết bị trước khi kiểm tra bảo hành.");
      return;
    }
    
    const activeRow = activeRange.getRow();
    let targetCode;
    
    // Xác định sheet đang làm việc để lấy mã thiết bị phù hợp
    if (activeSheet.getName() === EQUIPMENT_SHEET_NAME) {
      targetCode = activeSheet.getRange(activeRow, COL_EQUIP_ID).getValue();
    } else if (activeSheet.getName() === SHEET_PHIEU_CONG_VIEC) {
      targetCode = activeSheet.getRange(activeRow, COL_PCV_DOI_TUONG).getValue();
      
      // Nếu là dạng "MÃ - Tên", chỉ lấy phần mã
      if (typeof targetCode === 'string' && targetCode.includes(" - ")) {
        targetCode = targetCode.split(" - ")[0].trim();
      }
    } else {
      ui.alert("Vui lòng chọn một dòng trong sheet Danh mục Thiết bị hoặc Phiếu Công Việc.");
      return;
    }
    
    if (!targetCode) {
      ui.alert("Không tìm thấy mã thiết bị trên dòng đã chọn.");
      return;
    }
    
    // Gọi hàm kiểm tra bảo hành
    const warrantyInfo = checkWarrantyStatus(targetCode);
    
    if (warrantyInfo) {
      // Tạo HTML hiển thị thông tin bảo hành
      let htmlContent = `
        <style>
          body { font-family: Arial, sans-serif; padding: 15px; }
          .info-card { border: 1px solid #ddd; padding: 15px; border-radius: 5px; }
          .header { font-size: 16px; font-weight: bold; margin-bottom: 15px; }
          .warranty-status { font-size: 14px; margin: 10px 0; padding: 8px; border-radius: 3px; }
          .in-warranty { background-color: #d4edda; color: #155724; }
          .out-warranty { background-color: #f8d7da; color: #721c24; }
          .info-row { margin: 8px 0; display: flex; }
          .label { min-width: 120px; font-weight: bold; }
          .value { flex-grow: 1; }
          .actions { margin-top: 20px; padding-top: 15px; border-top: 1px solid #eee; }
        </style>
        
        <div class="info-card">
          <div class="header">Thông tin bảo hành thiết bị</div>
          
          <div class="warranty-status ${warrantyInfo.status.includes("Còn bảo hành") ? "in-warranty" : "out-warranty"}">
            ${warrantyInfo.status}
          </div>
          
          <div class="info-row">
            <div class="label">Mã thiết bị:</div>
            <div class="value">${targetCode}</div>
          </div>
          
          <div class="info-row">
            <div class="label">Nhà cung cấp:</div>
            <div class="value">${warrantyInfo.supplier}</div>
          </div>
          
          <div class="info-row">
            <div class="label">Mã mua hàng:</div>
            <div class="value">${warrantyInfo.purchaseId}</div>
          </div>
          
          <div class="actions">
            <p><strong>Tiếp theo:</strong> ${warrantyInfo.status.includes("Còn bảo hành") ? 
              "Thiết bị còn trong thời gian bảo hành. Đề xuất chuyển sang quy trình bảo hành với NCC." : 
              "Thiết bị đã hết bảo hành. Đề xuất xem xét thuê đơn vị ngoài sửa chữa."}
            </p>
          </div>
        </div>
      `;
      
      const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
          .setWidth(450)
          .setHeight(300);
      
      ui.showModalDialog(htmlOutput, "Thông tin bảo hành thiết bị");
    } else {
      ui.alert("Không thể lấy thông tin bảo hành. Vui lòng thử lại sau.");
    }
  } catch (e) {
    Logger.log(`Lỗi khi kiểm tra bảo hành: ${e}`);
    ui.alert(`Lỗi khi kiểm tra bảo hành: ${e.message}`);
  }
}
/**
 * Cập nhật trạng thái phiếu công việc sang "Đang kiểm tra BH"
 * Được gọi từ Dialog Kiểm tra bảo hành
 * @param {number} rowIndex Chỉ số dòng trong sheet Phiếu Công Việc
 * @param {string} equipmentCode Mã thiết bị
 * @param {string} supplier Thông tin nhà cung cấp
 * @return {object} Kết quả cập nhật {success: boolean, message: string}
 */
function updateWorkOrderForWarranty(rowIndex, equipmentCode, supplier) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
    
    if (!sheet) throw new Error("Không tìm thấy sheet Phiếu Công Việc");
    if (isNaN(rowIndex) || rowIndex < 2) throw new Error("Chỉ số dòng không hợp lệ");
    
    // Cập nhật trạng thái
    sheet.getRange(rowIndex, COL_PCV_TRANG_THAI).setValue("Đang kiểm tra BH");
    
    // Cập nhật thông tin ĐV Ngoài/liên hệ
    const supplierInfo = "Theo diện bảo hành - " + supplier;
    sheet.getRange(rowIndex, COL_PCV_CHI_TIET_NGOAI).setValue(supplierInfo);
    
    // Thêm ghi chú
    const currentNotes = sheet.getRange(rowIndex, COL_PCV_GHI_CHU).getValue() || "";
    const newNote = currentNotes + "\n" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") + ": Chuyển sang quy trình bảo hành.";
    sheet.getRange(rowIndex, COL_PCV_GHI_CHU).setValue(newNote);
    
    Logger.log(`Đã cập nhật dòng ${rowIndex} sang trạng thái "Đang kiểm tra BH" và thêm thông tin NCC`);
    
    return {
      success: true,
      message: "Đã chuyển phiếu sang quy trình bảo hành thành công!"
    };
  } catch (e) {
    Logger.log(`Lỗi updateWorkOrderForWarranty: ${e}`);
    return {
      success: false,
      message: e.toString()
    };
  }
}

/**
 * Cập nhật trạng thái phiếu công việc sang "Chờ đơn vị ngoài"
 * Được gọi từ Dialog Kiểm tra bảo hành
 * @param {number} rowIndex Chỉ số dòng trong sheet Phiếu Công Việc
 * @return {object} Kết quả cập nhật {success: boolean, message: string}
 */
function updateWorkOrderForExternal(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
    
    if (!sheet) throw new Error("Không tìm thấy sheet Phiếu Công Việc");
    if (isNaN(rowIndex) || rowIndex < 2) throw new Error("Chỉ số dòng không hợp lệ");
    
    // Cập nhật trạng thái
    sheet.getRange(rowIndex, COL_PCV_TRANG_THAI).setValue("Chờ đơn vị ngoài");
    
    // Thêm ghi chú
    const currentNotes = sheet.getRange(rowIndex, COL_PCV_GHI_CHU).getValue() || "";
    const newNote = currentNotes + "\n" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") + ": Chuyển sang thuê đơn vị ngoài.";
    sheet.getRange(rowIndex, COL_PCV_GHI_CHU).setValue(newNote);
    
    Logger.log(`Đã cập nhật dòng ${rowIndex} sang trạng thái "Chờ đơn vị ngoài"`);
    
    return {
      success: true,
      message: "Đã chuyển phiếu sang thuê đơn vị ngoài thành công!"
    };
  } catch (e) {
    Logger.log(`Lỗi updateWorkOrderForExternal: ${e}`);
    return {
      success: false,
      message: e.toString()
    };
  }
}

/**
 * Hiển thị dialog kiểm tra bảo hành cho dòng hiện tại
 * Được gọi từ menu tiện ích
 */
function checkCurrentEquipmentWarranty() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const activeRange = sheet.getActiveRange();
    
    if (!activeRange) {
      ui.alert("Vui lòng chọn một dòng thiết bị trước khi kiểm tra bảo hành.");
      return;
    }
    
    const rowIndex = activeRange.getRow();
    let targetCode = "";
    let sheetType = "";
    
    // Xác định sheet đang làm việc
    if (sheet.getName() === EQUIPMENT_SHEET_NAME) {
      targetCode = sheet.getRange(rowIndex, COL_EQUIP_ID).getValue();
      sheetType = "equipment";
    } else if (sheet.getName() === SHEET_PHIEU_CONG_VIEC) {
      const rawValue = sheet.getRange(rowIndex, COL_PCV_DOI_TUONG).getValue();
      targetCode = typeof rawValue === 'string' && rawValue.includes(" - ") 
                  ? rawValue.split(" - ")[0].trim() : rawValue;
      sheetType = "workorder";
    } else {
      ui.alert("Vui lòng chọn một dòng trong sheet Danh mục Thiết bị hoặc Phiếu Công Việc.");
      return;
    }
    
    if (!targetCode) {
      ui.alert("Không tìm thấy mã thiết bị trên dòng đã chọn.");
      return;
    }
    
    // Kiểm tra bảo hành
    const warrantyInfo = checkWarrantyStatus(targetCode);
    
    // Chuẩn bị dữ liệu cho dialog
    const data = {
      equipmentCode: targetCode,
      status: warrantyInfo.status,
      supplier: warrantyInfo.supplier,
      purchaseId: warrantyInfo.purchaseId,
      contactInfo: warrantyInfo.contactInfo || "Không có thông tin",
      rowIndex: sheetType === "workorder" ? rowIndex : null
    };
    
    // Hiển thị dialog
    const htmlTemplate = HtmlService.createTemplateFromFile('WarrantyCheckDialog');
    htmlTemplate.data = data;
    
    const htmlOutput = htmlTemplate.evaluate()
        .setWidth(450)
        .setHeight(350);
    
    ui.showModalDialog(htmlOutput, "Thông tin bảo hành thiết bị");
    
  } catch (e) {
    Logger.log(`Lỗi checkCurrentEquipmentWarranty: ${e}`);
    ui.alert(`Lỗi khi kiểm tra bảo hành: ${e.message}`);
  }
}
