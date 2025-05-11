/**
 * VendorManagement.gs
 * Tập hợp các hàm quản lý đơn vị ngoài/NCC bảo hành
 */

/**
 * Tạo mã đơn vị ngoài tự động
 * Format: DV + số thứ tự 4 chữ số
 * @return {string} Mã đơn vị mới
 */
function generateExternalVendorId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vendorSheet = ss.getSheetByName("Danh sách Đơn vị Ngoài");
    
    if (!vendorSheet) {
      throw new Error("Không tìm thấy sheet Danh sách Đơn vị Ngoài");
    }
    
    // Lấy dữ liệu cột mã (cột A)
    const lastRow = vendorSheet.getLastRow();
    let maxId = 0;
    
    if (lastRow > 1) {
      const idRange = vendorSheet.getRange(2, 1, lastRow - 1, 1);
      const idValues = idRange.getValues();
      
      // Tìm số thứ tự lớn nhất
      idValues.forEach(row => {
        if (row[0] && typeof row[0] === 'string' && row[0].startsWith('DV')) {
          const idNum = parseInt(row[0].substring(2), 10);
          if (!isNaN(idNum) && idNum > maxId) {
            maxId = idNum;
          }
        }
      });
    }
    
    const nextId = maxId + 1;
    const newVendorId = 'DV' + nextId.toString().padStart(4, '0');
    return newVendorId;
  } catch (error) {
    Logger.log("Lỗi tạo mã đơn vị: " + error);
    return null;
  }
}

/**
 * Lấy danh sách đơn vị ngoài để hiển thị trong dropdown
 * @return {string} HTML options cho dropdown
 */
function getVendorOptionsHtml() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Kiểm tra nếu sheet Danh sách Đơn vị Ngoài tồn tại
    let vendorSheet = ss.getSheetByName("Danh sách Đơn vị Ngoài");
    
    // Nếu chưa có sheet, tạo mới
    if (!vendorSheet) {
      vendorSheet = ss.insertSheet("Danh sách Đơn vị Ngoài");
      // Tạo header
      const headers = ["Mã ĐV", "Tên đơn vị", "Loại đơn vị", "Địa chỉ", 
                      "Người liên hệ", "Điện thoại", "Email", "Website", 
                      "Số tài khoản", "Ghi chú", "Ngày cập nhật"];
      vendorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header
      vendorSheet.getRange(1, 1, 1, headers.length)
        .setBackground("#f3f3f3")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      
      // Điều chỉnh độ rộng các cột
      vendorSheet.setColumnWidth(1, 100); // Mã ĐV
      vendorSheet.setColumnWidth(2, 200); // Tên đơn vị
      vendorSheet.setColumnWidth(3, 150); // Loại đơn vị
      vendorSheet.setColumnWidth(4, 200); // Địa chỉ
      vendorSheet.setColumnWidth(5, 150); // Người liên hệ
      vendorSheet.setColumnWidth(6, 120); // Điện thoại
      vendorSheet.setColumnWidth(7, 150); // Email
      vendorSheet.setColumnWidth(10, 200); // Ghi chú
      
      // Định dạng cột ngày
      vendorSheet.getRange(2, 11, 999, 1).setNumberFormat("dd/MM/yyyy");
      
      Logger.log("Đã tạo sheet Danh sách Đơn vị Ngoài mới");
      return ""; // Không có dữ liệu ban đầu
    }
    
    // Nếu sheet tồn tại nhưng không có dữ liệu
    if (vendorSheet.getLastRow() < 2) {
      return "";
    }
    
    // Lấy dữ liệu: Mã, Tên, Loại đơn vị
    const vendors = vendorSheet.getRange(2, 1, vendorSheet.getLastRow() - 1, 3).getValues();
    let optionsHtml = "";
    
    vendors.forEach(vendor => {
      if (vendor[0] && vendor[1]) {
        optionsHtml += `<option value="${vendor[0]}">${vendor[1]} (${vendor[2] || "N/A"})</option>`;
      }
    });
    
    return optionsHtml;
  } catch (error) {
    Logger.log("Lỗi lấy danh sách đơn vị: " + error);
    return "";
  }
}

/**
 * Lưu đơn vị ngoài mới
 * @param {string} name Tên đơn vị
 * @param {string} type Loại đơn vị
 * @param {string} contact Thông tin liên hệ
 * @return {Object} Kết quả {success, id, message}
 */
function saveNewVendor(name, type, contact) {
  try {
    if (!name) {
      return {success: false, message: "Tên đơn vị không được để trống"};
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let vendorSheet = ss.getSheetByName("Danh sách Đơn vị Ngoài");
    
    // Nếu sheet chưa tồn tại, tạo mới
    if (!vendorSheet) {
      vendorSheet = ss.insertSheet("Danh sách Đơn vị Ngoài");
      // Tạo header
      const headers = ["Mã ĐV", "Tên đơn vị", "Loại đơn vị", "Địa chỉ", 
                       "Người liên hệ", "Điện thoại", "Email", "Website", 
                       "Số tài khoản", "Ghi chú", "Ngày cập nhật"];
      vendorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header tương tự như trong hàm getVendorOptionsHtml
    }
    
    // Tạo mã đơn vị mới
    const newId = generateExternalVendorId();
    if (!newId) {
      return {success: false, message: "Lỗi tạo mã đơn vị"};
    }
    
    // Kiểm tra xem đơn vị đã tồn tại chưa (dựa vào tên)
    const lastRow = vendorSheet.getLastRow();
    if (lastRow > 1) {
      const existingVendors = vendorSheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < existingVendors.length; i++) {
        if (existingVendors[i][0] && existingVendors[i][0].toString().trim().toLowerCase() === name.toLowerCase()) {
          return {
            success: false, 
            message: "Đơn vị với tên này đã tồn tại"
          };
        }
      }
    }
    
    // Thêm dòng mới
    const newRow = [
      newId,                              // Mã ĐV
      name,                               // Tên đơn vị
      type,                               // Loại đơn vị
      "",                                 // Địa chỉ
      "",                                 // Người liên hệ
      "",                                 // Điện thoại
      "",                                 // Email
      "",                                 // Website
      "",                                 // Số tài khoản
      contact,                            // Ghi chú
      new Date()                          // Ngày cập nhật
    ];
    
    vendorSheet.appendRow(newRow);
    
    return {
      success: true, 
      id: newId,
      message: "Đã thêm đơn vị mới: " + name
    };
    
  } catch (error) {
    Logger.log("Lỗi khi lưu đơn vị mới: " + error);
    return {success: false, message: "Lỗi: " + error.toString()};
  }
}

/**
 * Lấy chi tiết đơn vị ngoài
 * @param {string} vendorId Mã đơn vị
 * @return {string} Thông tin chi tiết về đơn vị
 */
function getVendorDetails(vendorId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const vendorSheet = ss.getSheetByName("Danh sách Đơn vị Ngoài");
    
    if (!vendorSheet) {
      return "Không tìm thấy dữ liệu đơn vị ngoài";
    }
    
    // Tìm đơn vị theo mã
    const data = vendorSheet.getDataRange().getValues();
    let vendorInfo = "Không tìm thấy thông tin đơn vị";
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === vendorId) {
        vendorInfo = `Mã: ${data[i][0]}\n`;
        vendorInfo += `Tên: ${data[i][1]}\n`;
        vendorInfo += `Loại: ${data[i][2] || "N/A"}\n`;
        vendorInfo += `Địa chỉ: ${data[i][3] || "N/A"}\n`;
        vendorInfo += `Người liên hệ: ${data[i][4] || "N/A"}\n`;
        vendorInfo += `Điện thoại: ${data[i][5] || "N/A"}\n`;
        vendorInfo += `Email: ${data[i][6] || "N/A"}\n`;
        vendorInfo += `Ghi chú: ${data[i][9] || "N/A"}`;
        break;
      }
    }
    
    return vendorInfo;
    
  } catch (error) {
    Logger.log("Lỗi khi lấy chi tiết đơn vị: " + error);
    return "Lỗi: " + error.toString();
  }
}

/**
 * Menu để quản lý đơn vị ngoài/NCC
 */
function showVendorManagement() {
  try {
    // Tạo HTML output từ file mới nếu có
    const htmlOutput = HtmlService.createHtmlOutputFromFile('VendorManagementView')
        .setTitle('Quản lý đơn vị ngoài/NCC')
        .setWidth(800)
        .setHeight(600);
    
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
  } catch (e) {
    // Nếu chưa có file HTML VendorManagementView, mở sheet trực tiếp
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let vendorSheet = ss.getSheetByName("Danh sách Đơn vị Ngoài");
    
    if (!vendorSheet) {
      // Tạo sheet nếu chưa có
      getVendorOptionsHtml(); // Hàm này sẽ tạo sheet nếu chưa có
      vendorSheet = ss.getSheetByName("Danh sách Đơn vị Ngoài");
    }
    
    if (vendorSheet) {
      ss.setActiveSheet(vendorSheet);
    } else {
      SpreadsheetApp.getUi().alert("Không thể mở hoặc tạo sheet Danh sách Đơn vị Ngoài.");
    }
  }
}

/**
 * Tìm thông tin nhà cung cấp từ sheet Danh mục Thiết bị
 * @param {string} fullEquipmentString Chuỗi đầy đủ từ cột F
 * @return {Object} Thông tin về nhà cung cấp
 */
function findVendorInfoFromDeviceCatalog(fullEquipmentString) {
  try {
    Logger.log("Tìm thông tin NCC cho thiết bị: " + fullEquipmentString);
    
    if (!fullEquipmentString) {
      return { error: "Không có thông tin thiết bị để tìm kiếm" };
    }
    
    // Trích xuất mã thiết bị từ chuỗi đầy đủ
    let equipmentCode = fullEquipmentString;
    if (fullEquipmentString.includes(" - ")) {
      equipmentCode = fullEquipmentString.split(" - ")[0].trim();
      Logger.log("Đã trích xuất mã thiết bị: " + equipmentCode);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deviceCatalogSheet = ss.getSheetByName("Danh mục Thiết bị");
    
    if (!deviceCatalogSheet) {
      return { error: "Không tìm thấy sheet Danh mục Thiết bị" };
    }
    
    const data = deviceCatalogSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { error: "Sheet Danh mục Thiết bị không có dữ liệu" };
    }
    
    // Lấy headers từ hàng đầu tiên
    const headers = data[0];
    
    // Tìm vị trí cột mã thiết bị và nhà cung cấp
    let equipCodeCol = -1;
    let vendorCol = -1;
    let lookupCol = -1; // Cột Tra cứu phụ (X)
    
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().trim().toLowerCase();
      
      if (header === "mã thiết bị" || header === "ma thiết bị" || header === "mã tb") {
        equipCodeCol = i;
      } else if (header === "nhà cung cấp" || header === "nha cung cap" || header === "ncc") {
        vendorCol = i;
      } else if (header.includes("tra cứu") || header.includes("lookup")) {
        lookupCol = i;
      }
    }
    
    if (equipCodeCol === -1 || vendorCol === -1) {
      return { error: "Không tìm thấy cột mã thiết bị hoặc nhà cung cấp trong sheet" };
    }
    
    Logger.log(`Tìm kiếm với mã: ${equipmentCode} trong cột ${equipCodeCol+1}`);
    
    // PHƯƠNG PHÁP 1: Tìm theo mã thiết bị trong cột A
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const currentCode = row[equipCodeCol];
      
      if (currentCode && currentCode.toString().trim() === equipmentCode) {
        const vendor = row[vendorCol];
        Logger.log("Tìm thấy NCC từ mã: " + vendor);
        
        return {
          vendor: vendor || "Không rõ",
          id: "NCC-" + (vendor || "").toString().replace(/[^a-zA-Z0-9]/g, ""),
          contact: "",
          found: true
        };
      }
    }
    
    // PHƯƠNG PHÁP 2: Nếu không tìm thấy, thử tìm trong cột Tra cứu phụ (X)
    if (lookupCol !== -1) {
      Logger.log(`Tìm kiếm trong cột Tra cứu phụ (${lookupCol+1})`);
      
      for (let i = 1; i < data.length; i++) {
        const lookupValue = data[i][lookupCol];
        
        if (lookupValue && lookupValue.toString().includes(equipmentCode)) {
          const vendor = data[i][vendorCol];
          Logger.log("Tìm thấy NCC từ cột tra cứu: " + vendor);
          
          return {
            vendor: vendor || "Không rõ",
            id: "NCC-" + (vendor || "").toString().replace(/[^a-zA-Z0-9]/g, ""),
            contact: "",
            found: true
          };
        }
      }
    }
    
    return { error: "Không tìm thấy thông tin nhà cung cấp cho thiết bị " + equipmentCode };
    
  } catch (error) {
    Logger.log("Lỗi khi tìm thông tin nhà cung cấp: " + error);
    return { error: "Lỗi khi tìm thông tin nhà cung cấp: " + error.toString() };
  }
}

