// ==================================
// CÁC HÀM TRUY XUẤT DỮ LIỆU
// ==================================

/**
 * Lấy mã viết tắt cho Loại Thiết Bị và Vị trí từ Sheet Cấu hình.
 * @param {string} equipmentTypeValue Giá trị Loại Thiết Bị cần tìm.
 * @param {string|null} locationValue Giá trị Vị trí cần tìm (có thể là null).
 * @return {object|null} Object {type: mã loại TB, location: mã vị trí} hoặc null nếu không tìm thấy type.
 */
function getAcronyms(equipmentTypeValue, locationValue) {
  // Chuẩn hóa giá trị đầu vào
  const typeValueToFind = equipmentTypeValue ? equipmentTypeValue.toString().trim() : null;
  const locValueToFind = locationValue ? locationValue.toString().trim() : null;
  
  if (!typeValueToFind) {
    Logger.log("Lỗi: Giá trị equipmentTypeValue đầu vào không hợp lệ.");
    return null;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) {
      Logger.log(`Lỗi: Không tìm thấy Sheet "${SETTINGS_SHEET_NAME}"`);
      return null;
    }

    // Đọc dữ liệu sheet Cấu hình để tránh gọi API nhiều lần
    const data = settingsSheet.getDataRange().getValues();
    let typeAcronym = null;
    let locationAcronym = null;

    // Tìm mã viết tắt cho Loại TB và Vị trí
    for (let i = 1; i < data.length; i++) {
      const typeValueInSheet = data[i][COL_SETTINGS_LOAI_TB_GIATRI - 1];
      const typeAcronymInSheet = data[i][COL_SETTINGS_LOAI_TB_MA - 1];
      const locValueInSheet = data[i][COL_SETTINGS_VITRI_GIATRI - 1];
      const locAcronymInSheet = data[i][COL_SETTINGS_VITRI_MA - 1];
      
      // Tìm mã viết tắt Loại TB
      if (!typeAcronym && typeValueInSheet && typeValueInSheet.toString().trim() === typeValueToFind) {
        if (typeAcronymInSheet && typeAcronymInSheet.toString().trim() !== "") {
          typeAcronym = typeAcronymInSheet.toString().trim();
        }
      }
      
      // Tìm mã viết tắt Vị trí (nếu cần)
      if (locValueToFind && !locationAcronym && locValueInSheet 
          && locValueInSheet.toString().trim() === locValueToFind) {
        if (locAcronymInSheet && locAcronymInSheet.toString().trim() !== "") {
          locationAcronym = locAcronymInSheet.toString().trim();
        }
      }
      
      // Thoát sớm nếu đã tìm đủ thông tin cần thiết
      if (typeAcronym && (!locValueToFind || locationAcronym)) {
        break;
      }
    }

    // Trả về kết quả
    if (typeAcronym) {
      return { type: typeAcronym, location: locationAcronym };
    }
    
    Logger.log(`Không tìm thấy Mã VT cho Loại TB "${typeValueToFind}"`);
    return null;
  } catch (e) {
    Logger.log(`Lỗi trong getAcronyms: ${e}`);
    return null;
  }
}

/**
 * Lấy thông tin Nhà cung cấp, Ngày mua, Hạn bảo hành từ Sheet Chi tiết Mua Hàng.
 * @param {string} purchaseId Mã Lô Mua Hàng / ID Giao Dịch.
 * @return {object} Object {supplier, purchaseDate, warrantyEnd} hoặc null.
 */
function getPurchaseInfo(purchaseId) {
  if (!purchaseId) return null;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const purchaseSheet = ss.getSheetByName(PURCHASE_SHEET_NAME);
    if (!purchaseSheet) {
      Logger.log(`Lỗi: Không tìm thấy Sheet "${PURCHASE_SHEET_NAME}"`);
      return null;
    }
    
    const data = purchaseSheet.getDataRange().getValues();

    // Tìm thông tin mua hàng theo mã
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL_PURCHASE_ID - 1] == purchaseId) {
        return {
          supplier: data[i][COL_PURCHASE_SUPPLIER - 1],
          purchaseDate: data[i][COL_PURCHASE_DATE - 1] instanceof Date ? data[i][COL_PURCHASE_DATE - 1] : null,
          warrantyEnd: data[i][COL_PURCHASE_WARRANTY_END - 1] instanceof Date ? data[i][COL_PURCHASE_WARRANTY_END - 1] : null
        };
      }
    }
    
    Logger.log(`Không tìm thấy thông tin cho Mã Lô Mua Hàng: ${purchaseId}`);
    return null;
  } catch (e) {
    Logger.log(`Lỗi trong getPurchaseInfo: ${e}`);
    return null;
  }
}

/**
 * Lấy danh sách lịch sử bảo trì cho một Mã Đối tượng cụ thể.
 * @param {string} targetCode Mã Thiết Bị hoặc Mã Hệ thống cần tra cứu.
 * @return {Array<Object>} Mảng các đối tượng lịch sử, hoặc mảng rỗng nếu không có dữ liệu.
 */
function getMaintenanceHistory(targetCode) {
  if (!targetCode || typeof targetCode !== 'string' || targetCode.trim() === "") {
    Logger.log("getMaintenanceHistory: Mã đối tượng không hợp lệ.");
    return [];
  }
  
  targetCode = targetCode.trim();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!historySheet) {
      Logger.log(`Lỗi: Không tìm thấy Sheet "${HISTORY_SHEET_NAME}"`);
      return [];
    }

    const lastRow = historySheet.getLastRow();
    if (lastRow < 2) return [];

    const dataRange = historySheet.getRange(2, 1, lastRow - 1, COL_HISTORY_DETAIL_NOTE);
    const dataValues = dataRange.getValues();
    const historyRecords = [];

    // Lọc và sắp xếp lịch sử
    for (let i = 0; i < dataValues.length; i++) {
      const rowData = dataValues[i];
      const rawTargetValue = rowData[COL_HISTORY_TARGET_CODE - 1];
      let parsedCode = "";

      if (rawTargetValue && typeof rawTargetValue === 'string') {
        parsedCode = rawTargetValue.split(" - ")[0].trim();
      } else if (rawTargetValue) {
        parsedCode = rawTargetValue.toString().trim();
      }

      if (parsedCode === targetCode) {
        const execDateRaw = rowData[COL_HISTORY_EXEC_DATE - 1];
        const formattedDate = (execDateRaw instanceof Date)
          ? Utilities.formatDate(execDateRaw, Session.getScriptTimeZone(), "dd/MM/yyyy")
          : (execDateRaw || "");

        historyRecords.push({
          id: rowData[COL_HISTORY_ID - 1] || "",
          date: formattedDate,
          workType: rowData[COL_HISTORY_WORK_TYPE - 1] || "",
          description: rowData[COL_HISTORY_DESCRIPTION - 1] || "",
          performer: rowData[COL_HISTORY_PERFORMER - 1] || "",
          externalDetails: rowData[COL_HISTORY_EXTERNAL_DETAILS - 1] || "",
          cost: rowData[COL_HISTORY_COST - 1],
          status: rowData[COL_HISTORY_STATUS - 1] || "",
          warrantyCheck: rowData[COL_HISTORY_WARRANTY_CHECK - 1],
          warrantyReqId: rowData[COL_HISTORY_WARRANTY_REQ_ID - 1] || "",
          warrantyReqStat: rowData[COL_HISTORY_WARRANTY_REQ_STAT - 1] || "",
          warrantyReqNote: rowData[COL_HISTORY_WARRANTY_REQ_NOTE - 1] || "",
          assetPostStatus: rowData[COL_HISTORY_ASSET_POST_STATUS - 1] || "",
          detailNote: rowData[COL_HISTORY_DETAIL_NOTE - 1] || ""
        });
      }
    }

    return historyRecords;
  } catch (e) {
    Logger.log(`Lỗi trong getMaintenanceHistory: ${e}`);
    return [];
  }
}

/**
 * Lấy danh sách các tên vị trí duy nhất từ sheet Thiết bị và Settings.
 * @return {Array<string>} Mảng các tên vị trí đã sắp xếp.
 */
function getLocationList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    let locations = new Set();

    // Lấy từ sheet Danh mục Thiết bị
    if (equipSheet && equipSheet.getLastRow() >= 2) {
      const equipLocations = equipSheet.getRange(2, COL_EQUIP_LOCATION, equipSheet.getLastRow() - 1, 1).getValues();
      equipLocations.forEach(row => {
        if (row[0] && row[0].toString().trim() !== "") {
          locations.add(row[0].toString().trim());
        }
      });
    }

    // Lấy từ sheet Cấu hình
    if (settingsSheet && settingsSheet.getLastRow() >= 2) {
      const settingsLocations = settingsSheet.getRange(2, COL_SETTINGS_VITRI_GIATRI, settingsSheet.getLastRow() - 1, 1).getValues();
      settingsLocations.forEach(row => {
        if (row[0] && row[0].toString().trim() !== "") {
          locations.add(row[0].toString().trim());
        }
      });
    }

    return Array.from(locations).sort();
  } catch (e) {
    Logger.log(`Lỗi trong getLocationList: ${e}`);
    return [];
  }
}

/**
 * Lấy danh sách thiết bị và hệ thống tại một vị trí cụ thể.
 * @param {string} locationName Tên vị trí cần tra cứu.
 * @return {object} {equipment: [{id, name, type, location, parentId}], systems: [{code, description}]}
 */
function getAssetsByLocation(locationName) {
  if (!locationName || typeof locationName !== 'string' || locationName.trim() === "") {
    throw new Error("Vui lòng chọn hoặc nhập tên vị trí hợp lệ.");
  }
  
  locationName = locationName.trim();
  let results = { equipment: [], systems: [] };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    const systemDefSheet = ss.getSheetByName(SHEET_DINH_NGHIA_HE_THONG);

    // Tìm thiết bị tại vị trí
    if (equipSheet && equipSheet.getLastRow() >= 2) {
      const lastEquipColNeeded = Math.max(COL_EQUIP_ID, COL_EQUIP_NAME, COL_EQUIP_TYPE, COL_EQUIP_PARENT_ID, COL_EQUIP_LOCATION);
      const equipData = equipSheet.getRange(2, 1, equipSheet.getLastRow() - 1, lastEquipColNeeded).getValues();
      
      equipData.forEach(row => {
        const currentLocation = row[COL_EQUIP_LOCATION - 1];
        if (currentLocation && currentLocation.toString().trim() === locationName) {
          results.equipment.push({
            id: row[COL_EQUIP_ID - 1],
            name: row[COL_EQUIP_NAME - 1] || "",
            type: row[COL_EQUIP_TYPE - 1] || "",
            location: currentLocation.toString().trim(),
            parentId: row[COL_EQUIP_PARENT_ID - 1] || ""
          });
        }
      });
    }

    // Tìm mã viết tắt của vị trí
    let locationAcronym = null;
    if (settingsSheet && settingsSheet.getLastRow() >= 2) {
      const settingsVitriValues = settingsSheet.getRange(2, COL_SETTINGS_VITRI_GIATRI, settingsSheet.getLastRow() - 1, 2).getValues();
      
      for (let i = 0; i < settingsVitriValues.length; i++) {
        const currentName = settingsVitriValues[i][0];
        if (currentName && currentName.toString().trim() === locationName) {
          locationAcronym = settingsVitriValues[i][1] ? settingsVitriValues[i][1].toString().trim() : null;
          break;
        }
      }
    }

    // Tìm hệ thống dựa trên mã vị trí
    if (locationAcronym && systemDefSheet && systemDefSheet.getLastRow() >= 2) {
      const searchSuffix = "-" + locationAcronym;
      const systemData = systemDefSheet.getRange(2, COL_HT_MA, systemDefSheet.getLastRow() - 1, 2).getValues();
      
      systemData.forEach(row => {
        const systemCode = row[COL_HT_MA - 1] ? row[COL_HT_MA - 1].toString().trim() : null;
        if (systemCode && systemCode.endsWith(searchSuffix)) {
          results.systems.push({
            code: systemCode,
            description: row[COL_HT_MO_TA - 1] || ""
          });
        }
      });
    }

    return results;
  } catch (e) {
    Logger.log(`Lỗi trong getAssetsByLocation: ${e}`);
    throw new Error(`Không thể lấy dữ liệu tài sản: ${e.message}`);
  }
}

/**
 * Lấy danh sách các thiết bị con dựa vào Mã Thiết bị Cha.
 * @param {string} parentId Mã của thiết bị Cha cần tra cứu.
 * @return {Array<Object>} Mảng các object thông tin thiết bị con.
 */
function getChildEquipment(parentId) {
  if (!parentId || typeof parentId !== 'string' || parentId.trim() === "") {
    return [];
  }
  
  parentId = parentId.trim();
  const children = [];

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    if (!equipSheet || equipSheet.getLastRow() < 2) {
      return [];
    }

    const lastColNeeded = Math.max(COL_EQUIP_ID, COL_EQUIP_NAME, COL_EQUIP_TYPE, COL_EQUIP_PARENT_ID, COL_EQUIP_LOCATION);
    const equipData = equipSheet.getRange(2, 1, equipSheet.getLastRow() - 1, lastColNeeded).getValues();

    equipData.forEach(row => {
      const currentParentId = row[COL_EQUIP_PARENT_ID - 1];
      if (currentParentId && currentParentId.toString().trim() === parentId) {
        children.push({
          id: row[COL_EQUIP_ID - 1],
          name: row[COL_EQUIP_NAME - 1] || "",
          type: row[COL_EQUIP_TYPE - 1] || "",
          location: row[COL_EQUIP_LOCATION - 1] || ""
        });
      }
    });

    return children;
  } catch (e) {
    Logger.log(`Lỗi trong getChildEquipment: ${e}`);
    return [];
  }
}

/**
 * Lấy danh sách tất cả các Mã Thiết Bị hợp lệ từ Cột A sheet Danh mục TB.
 * @return {Array<string>} Mảng các Mã Thiết Bị.
 */
function getAllEquipmentIds() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    if (!equipSheet || equipSheet.getLastRow() < 2) {
      return [];
    }
    
    const ids = equipSheet.getRange(2, COL_EQUIP_ID, equipSheet.getLastRow() - 1, 1)
                          .getValues()
                          .flat()
                          .map(id => id ? id.toString().trim() : null)
                          .filter(Boolean);
    
    return ids;
  } catch (e) {
    Logger.log(`Lỗi trong getAllEquipmentIds: ${e}`);
    return [];
  }
}
/**
 * Kiểm tra tình trạng bảo hành của thiết bị/hệ thống
 * Thêm thông tin vào Dialog khi mở
 */
function checkWarrantyStatus(targetCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipmentSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    
    // Tìm thiết bị
    const data = equipmentSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][COL_EQUIP_ID - 1] === targetCode) {
        const warrantyEnd = data[i][COL_EQUIP_WARRANTY_END - 1];
        const supplierName = data[i][COL_EQUIP_SUPPLIER - 1] || "Không có";
        const purchaseId = data[i][COL_EQUIP_PURCHASE_ID - 1] || "Không có";
        
        const today = new Date();
        let warrantyStatus = "Không có thông tin bảo hành";
        
        if (warrantyEnd instanceof Date) {
          if (warrantyEnd > today) {
            warrantyStatus = "Còn bảo hành đến: " + Utilities.formatDate(warrantyEnd, Session.getScriptTimeZone(), "dd/MM/yyyy");
          } else {
            warrantyStatus = "Hết bảo hành từ: " + Utilities.formatDate(warrantyEnd, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
        }
        
        return {
          status: warrantyStatus,
          supplier: supplierName,
          purchaseId: purchaseId
        };
      }
    }
    return { status: "Không tìm thấy thông tin thiết bị", supplier: "N/A", purchaseId: "N/A" };
  } catch (e) {
    Logger.log("Lỗi khi kiểm tra bảo hành: " + e);
    return { status: "Lỗi kiểm tra bảo hành", supplier: "N/A", purchaseId: "N/A" };
  }
}
