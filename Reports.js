/**
 * Hiển thị báo cáo số lượng thiết bị theo loại và vị trí.
 * Được gọi từ menu "Báo cáo & Thống kê".
 */
function reportEquipmentByType() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    if (!equipSheet) throw new Error(`Không tìm thấy sheet "${EQUIPMENT_SHEET_NAME}"`);
    
    // Đọc dữ liệu thiết bị
    const lastRow = equipSheet.getLastRow();
    if (lastRow < 2) throw new Error("Không có dữ liệu thiết bị để báo cáo");
    
    const data = equipSheet.getRange(2, 1, lastRow - 1, Math.max(COL_EQUIP_TYPE, COL_EQUIP_LOCATION, COL_EQUIP_STATUS)).getValues();
    
    // Tính toán thống kê theo loại thiết bị
    const typeStats = {};
    const locationStats = {};
    const statusStats = {};
    let totalEquipment = 0;
    
    data.forEach(row => {
      const equipType = row[COL_EQUIP_TYPE - 1]?.toString().trim() || "Không xác định";
      const location = row[COL_EQUIP_LOCATION - 1]?.toString().trim() || "Không xác định";
      const status = row[COL_EQUIP_STATUS - 1]?.toString().trim() || "Không xác định";
      
      // Chỉ đếm nếu có giá trị ở cột A (Mã thiết bị)
      if (row[COL_EQUIP_ID - 1]) {
        // Thống kê theo loại
        typeStats[equipType] = (typeStats[equipType] || 0) + 1;
        
        // Thống kê theo vị trí
        locationStats[location] = (locationStats[location] || 0) + 1;
        
        // Thống kê theo trạng thái
        statusStats[status] = (statusStats[status] || 0) + 1;
        
        totalEquipment++;
      }
    });
    
    // Tạo HTML để hiển thị báo cáo
    let html = '<h3>Báo cáo Thống kê Thiết bị</h3>';
    html += `<p>Tổng số thiết bị: <b>${totalEquipment}</b></p>`;
    
    // Báo cáo theo loại thiết bị
    html += '<h4>Theo Loại Thiết bị</h4>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Loại Thiết bị</th><th>Số lượng</th><th>Tỷ lệ</th></tr>';
    
    Object.keys(typeStats).sort().forEach(type => {
      const count = typeStats[type];
      const percentage = ((count / totalEquipment) * 100).toFixed(1);
      html += `<tr><td>${type}</td><td style="text-align: center">${count}</td><td style="text-align: center">${percentage}%</td></tr>`;
    });
    html += '</table>';
    
    // Báo cáo theo vị trí
    html += '<h4>Theo Vị trí</h4>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Vị trí</th><th>Số lượng</th><th>Tỷ lệ</th></tr>';
    
    Object.keys(locationStats).sort().forEach(location => {
      const count = locationStats[location];
      const percentage = ((count / totalEquipment) * 100).toFixed(1);
      html += `<tr><td>${location}</td><td style="text-align: center">${count}</td><td style="text-align: center">${percentage}%</td></tr>`;
    });
    html += '</table>';
    
    // Báo cáo theo trạng thái
    html += '<h4>Theo Trạng thái Hoạt động</h4>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Trạng thái</th><th>Số lượng</th><th>Tỷ lệ</th></tr>';
    
    Object.keys(statusStats).sort().forEach(status => {
      const count = statusStats[status];
      const percentage = ((count / totalEquipment) * 100).toFixed(1);
      html += `<tr><td>${status}</td><td style="text-align: center">${count}</td><td style="text-align: center">${percentage}%</td></tr>`;
    });
    html += '</table>';
    
    // Hiển thị báo cáo trong một hộp thoại
    const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(600)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Báo cáo Thống kê Thiết bị');
    
  } catch (e) {
    Logger.log(`Lỗi khi tạo báo cáo thiết bị: ${e}`);
    SpreadsheetApp.getUi().alert(`Lỗi khi tạo báo cáo: ${e.message}`);
  }
}


/**
 * Hiển thị báo cáo số lượng phiếu công việc theo trạng thái.
 * Được gọi từ menu "Báo cáo & Thống kê".
 */
function reportWorkOrderByStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const woSheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
    if (!woSheet) throw new Error(`Không tìm thấy sheet "${SHEET_PHIEU_CONG_VIEC}"`);
    
    // Đọc dữ liệu phiếu công việc
    const lastRow = woSheet.getLastRow();
    if (lastRow < 2) throw new Error("Không có dữ liệu phiếu công việc để báo cáo");
    
    const data = woSheet.getRange(2, 1, lastRow - 1, Math.max(COL_PCV_MA_PHIEU, COL_PCV_LOAI_CV, COL_PCV_TRANG_THAI, COL_PCV_NGAY_TAO)).getValues();
    
    // Tính toán thống kê
    const statusStats = {};
    const typeStats = {};
    const monthStats = {};
    let totalWorkOrders = 0;
    
    data.forEach(row => {
      // Chỉ đếm nếu có Mã Phiếu CV
      if (row[COL_PCV_MA_PHIEU - 1]) {
        const status = row[COL_PCV_TRANG_THAI - 1]?.toString().trim() || "Không xác định";
        const workType = row[COL_PCV_LOAI_CV - 1]?.toString().trim() || "Không xác định";
        
        // Thống kê theo trạng thái
        statusStats[status] = (statusStats[status] || 0) + 1;
        
        // Thống kê theo loại công việc
        typeStats[workType] = (typeStats[workType] || 0) + 1;
        
        // Thống kê theo tháng tạo
        const createDate = row[COL_PCV_NGAY_TAO - 1];
        if (createDate instanceof Date && !isNaN(createDate)) {
          const monthYear = Utilities.formatDate(createDate, Session.getScriptTimeZone(), "MM/yyyy");
          monthStats[monthYear] = (monthStats[monthYear] || 0) + 1;
        }
        
        totalWorkOrders++;
      }
    });
    
    // Tạo HTML để hiển thị báo cáo
    let html = '<h3>Báo cáo Thống kê Phiếu Công Việc</h3>';
    html += `<p>Tổng số phiếu: <b>${totalWorkOrders}</b></p>`;
    
    // Báo cáo theo trạng thái
    html += '<h4>Theo Trạng thái</h4>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Trạng thái</th><th>Số lượng</th><th>Tỷ lệ</th></tr>';
    
    Object.keys(statusStats).sort().forEach(status => {
      const count = statusStats[status];
      const percentage = ((count / totalWorkOrders) * 100).toFixed(1);
      html += `<tr><td>${status}</td><td style="text-align: center">${count}</td><td style="text-align: center">${percentage}%</td></tr>`;
    });
    html += '</table>';
    
    // Báo cáo theo loại công việc
    html += '<h4>Theo Loại Công Việc</h4>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Loại Công Việc</th><th>Số lượng</th><th>Tỷ lệ</th></tr>';
    
    Object.keys(typeStats).sort().forEach(type => {
      const count = typeStats[type];
      const percentage = ((count / totalWorkOrders) * 100).toFixed(1);
      html += `<tr><td>${type}</td><td style="text-align: center">${count}</td><td style="text-align: center">${percentage}%</td></tr>`;
    });
    html += '</table>';
    
    // Báo cáo theo tháng tạo (chỉ hiển thị 6 tháng gần nhất)
    html += '<h4>Theo Tháng Tạo (6 tháng gần nhất)</h4>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Tháng/Năm</th><th>Số lượng</th></tr>';
    
    // Sắp xếp tháng theo thứ tự mới nhất đến cũ nhất
    const sortedMonths = Object.keys(monthStats).sort((a, b) => {
      const [monthA, yearA] = a.split('/').map(Number);
      const [monthB, yearB] = b.split('/').map(Number);
      if (yearA !== yearB) return yearB - yearA;
      return monthB - monthA;
    });
    
    // Lấy 6 tháng gần nhất
    sortedMonths.slice(0, 6).forEach(month => {
      html += `<tr><td>${month}</td><td style="text-align: center">${monthStats[month]}</td></tr>`;
    });
    html += '</table>';
    
    // Hiển thị báo cáo trong một hộp thoại
    const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(600)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Báo cáo Thống kê Phiếu Công Việc');
    
  } catch (e) {
    Logger.log(`Lỗi khi tạo báo cáo phiếu công việc: ${e}`);
    SpreadsheetApp.getUi().alert(`Lỗi khi tạo báo cáo: ${e.message}`);
  }
}

/**
 * Hiển thị báo cáo lịch sử bảo trì trong khoảng thời gian.
 * Cho phép người dùng chọn số tháng gần nhất để phân tích.
 * Được gọi từ menu "Báo cáo & Thống kê".
 */
function reportMaintenanceHistory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    if (!historySheet) throw new Error(`Không tìm thấy sheet "${HISTORY_SHEET_NAME}"`);
    
    // Tạo giao diện chọn khoảng thời gian
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      'Báo cáo Lịch sử Bảo trì',
      'Vui lòng chọn khoảng thời gian (Định dạng: SỐ_THÁNG, ví dụ 3 cho 3 tháng gần nhất hoặc 1 cho tháng hiện tại):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() === ui.Button.CANCEL) return;
    
    const monthsInput = parseInt(response.getResponseText(), 10);
    if (isNaN(monthsInput) || monthsInput <= 0) {
      ui.alert('Số tháng không hợp lệ. Vui lòng nhập số dương.');
      return;
    }
    
    // Xác định ngày bắt đầu khoảng thời gian
    const endDate = new Date();
    let startDate = new Date();
    startDate.setMonth(startDate.getMonth() - monthsInput + 1);
    startDate.setDate(1);
    startDate.setHours(0, 0, 0, 0);
    
    // Format các ngày giới hạn
    const startDateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    const endDateStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    // Đọc dữ liệu lịch sử
    const lastRow = historySheet.getLastRow();
    if (lastRow < 2) throw new Error("Không có dữ liệu lịch sử để báo cáo");
    
    const data = historySheet.getRange(2, 1, lastRow - 1, Math.max(
      COL_HISTORY_ID, COL_HISTORY_TARGET_CODE, COL_HISTORY_EXEC_DATE, 
      COL_HISTORY_WORK_TYPE, COL_HISTORY_COST, COL_HISTORY_STATUS
    )).getValues();
    
    // Lọc dữ liệu trong khoảng thời gian
    const filteredData = data.filter(row => {
      const execDate = row[COL_HISTORY_EXEC_DATE - 1];
      return execDate instanceof Date && !isNaN(execDate) && 
             execDate >= startDate && execDate <= endDate;
    });
    
    // Tính toán thống kê
    const workTypeStats = {};
    const targetStats = {};
    const monthlyStats = {};
    let totalRecords = filteredData.length;
    let totalCost = 0;
    
    filteredData.forEach(row => {
      const target = row[COL_HISTORY_TARGET_CODE - 1]?.toString().trim() || "Không xác định";
      const workType = row[COL_HISTORY_WORK_TYPE - 1]?.toString().trim() || "Không xác định";
      const cost = typeof row[COL_HISTORY_COST - 1] === 'number' ? row[COL_HISTORY_COST - 1] : 0;
      const execDate = row[COL_HISTORY_EXEC_DATE - 1];
      
      // Thống kê theo loại công việc
      workTypeStats[workType] = (workTypeStats[workType] || 0) + 1;
      
      // Thống kê theo đối tượng
      targetStats[target] = (targetStats[target] || 0) + 1;
      
      // Thống kê theo tháng thực hiện
      if (execDate instanceof Date && !isNaN(execDate)) {
        const monthYear = Utilities.formatDate(execDate, Session.getScriptTimeZone(), "MM/yyyy");
        if (!monthlyStats[monthYear]) {
          monthlyStats[monthYear] = { count: 0, cost: 0 };
        }
        monthlyStats[monthYear].count++;
        monthlyStats[monthYear].cost += cost;
      }
      
      // Tổng chi phí
      totalCost += cost;
    });
    
    // Tạo HTML để hiển thị báo cáo
    let html = '<h3>Báo cáo Lịch sử Bảo trì</h3>';
    html += `<p>Khoảng thời gian: <b>${startDateStr}</b> đến <b>${endDateStr}</b> (${monthsInput} tháng)</p>`;
    html += `<p>Tổng số bản ghi: <b>${totalRecords}</b></p>`;
    html += `<p>Tổng chi phí: <b>${totalCost.toLocaleString('vi-VN')}</b> VND</p>`;
    
    // Báo cáo theo loại công việc
    html += '<h4>Theo Loại Công Việc</h4>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Loại Công Việc</th><th>Số lượng</th><th>Tỷ lệ</th></tr>';
    
    Object.keys(workTypeStats).sort().forEach(type => {
      const count = workTypeStats[type];
      const percentage = totalRecords > 0 ? ((count / totalRecords) * 100).toFixed(1) : "0.0";
      html += `<tr><td>${type}</td><td style="text-align: center">${count}</td><td style="text-align: center">${percentage}%</td></tr>`;
    });
    html += '</table>';
    
    // Báo cáo theo tháng thực hiện
    html += '<h4>Theo Tháng Thực hiện</h4>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Tháng/Năm</th><th>Số lượng</th><th>Chi phí (VND)</th></tr>';
    
    // Sắp xếp tháng theo thứ tự mới nhất đến cũ nhất
    const sortedMonths = Object.keys(monthlyStats).sort((a, b) => {
      const [monthA, yearA] = a.split('/').map(Number);
      const [monthB, yearB] = b.split('/').map(Number);
      if (yearA !== yearB) return yearB - yearA;
      return monthB - monthA;
    });
    
    sortedMonths.forEach(month => {
      html += `<tr>
        <td>${month}</td>
        <td style="text-align: center">${monthlyStats[month].count}</td>
        <td style="text-align: right">${monthlyStats[month].cost.toLocaleString('vi-VN')}</td>
      </tr>`;
    });
    html += '</table>';
    
    // Báo cáo top 10 đối tượng có nhiều bản ghi nhất
    html += '<h4>Top 10 Đối tượng có nhiều Bản ghi nhất</h4>';
    html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
    html += '<tr><th>Mã Đối tượng</th><th>Số lượng</th></tr>';
    
    // Sắp xếp và lấy top 10
    const sortedTargets = Object.entries(targetStats)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10);
    
    sortedTargets.forEach(([target, count]) => {
      html += `<tr><td>${target}</td><td style="text-align: center">${count}</td></tr>`;
    });
    html += '</table>';
    
    // Hiển thị báo cáo trong một hộp thoại
    const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(600)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Báo cáo Lịch sử Bảo trì');
    
  } catch (e) {
    Logger.log(`Lỗi khi tạo báo cáo lịch sử bảo trì: ${e}`);
    SpreadsheetApp.getUi().alert(`Lỗi khi tạo báo cáo: ${e.message}`);
  }
}

/**
 * Hiển thị báo cáo lỗi dữ liệu tổng hợp.
 * Sử dụng các hàm kiểm tra dữ liệu hiện có để tạo báo cáo.
 * Được gọi từ menu "Báo cáo & Thống kê".
 */
function reportDataErrors() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Thu thập lỗi từ các kiểm tra khác nhau
    const typeLocationErrors = findMissingTypeLocation_();
    const historyErrors = findOrphanHistoryRecords_();
    
    // Thêm các kiểm tra khác vào đây nếu cần
    // const otherErrors = findOtherErrors_();
    
    // Tổng hợp tất cả lỗi
    const allErrors = [...typeLocationErrors, ...historyErrors];
    
    // Phân loại lỗi theo sheet
    const errorsBySheet = {};
    allErrors.forEach(err => {
      if (!errorsBySheet[err.sheetName]) {
        errorsBySheet[err.sheetName] = [];
      }
      errorsBySheet[err.sheetName].push(err);
    });
    
    // Tạo HTML để hiển thị báo cáo
    let html = '<h3>Báo cáo Lỗi Dữ liệu Tổng hợp</h3>';
    html += `<p>Tổng số lỗi tìm thấy: <b>${allErrors.length}</b></p>`;
    
    if (allErrors.length === 0) {
      html += '<p style="color: green;"><b>Không tìm thấy lỗi dữ liệu. Hệ thống đang hoạt động tốt!</b></p>';
    } else {
      // Hiển thị lỗi theo từng sheet
      Object.keys(errorsBySheet).forEach(sheetName => {
        const sheetErrors = errorsBySheet[sheetName];
        html += `<h4>Sheet: ${sheetName} (${sheetErrors.length} lỗi)</h4>`;
        html += '<table border="1" style="border-collapse: collapse; width: 100%;">';
        html += '<tr><th>Dòng</th><th>Cột</th><th>Mô tả lỗi</th></tr>';
        
        sheetErrors.forEach(err => {
          const row = err.row || "N/A";
          const column = err.column ? columnToLetter_(err.column) : "N/A";
          html += `<tr>
            <td style="text-align: center">${row}</td>
            <td style="text-align: center">${column}</td>
            <td>${err.message}</td>
          </tr>`;
        });
        html += '</table>';
      });
      
      // Thêm gợi ý khắc phục
      html += '<h4>Gợi ý Khắc phục</h4>';
      html += '<ul>';
      html += '<li>Kiểm tra và bổ sung các thông tin còn thiếu.</li>';
      html += '<li>Đảm bảo các mã tham chiếu đến tồn tại trong sheet tương ứng.</li>';
      html += '<li>Sử dụng chức năng "Kiểm tra Tính nhất quán Dữ liệu" để xem chi tiết hơn và đánh dấu các lỗi trực tiếp trong sheet.</li>';
      html += '</ul>';
    }
    
    // Hiển thị báo cáo trong một hộp thoại
    const htmlOutput = HtmlService.createHtmlOutput(html)
        .setWidth(700)
        .setHeight(500);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Báo cáo Lỗi Dữ liệu Tổng hợp');
    
  } catch (e) {
    Logger.log(`Lỗi khi tạo báo cáo lỗi dữ liệu: ${e}`);
    SpreadsheetApp.getUi().alert(`Lỗi khi tạo báo cáo: ${e.message}`);
  }
}

/**
 * Tạo header CSS chung cho các báo cáo.
 * @return {string} Chuỗi CSS để thêm vào header báo cáo.
 * @private
 */
function getReportHeaderStyle_() {
  return `
    <style>
      body { font-family: Arial, sans-serif; margin: 10px; font-size: 13px; }
      h3 { color: #4285f4; margin-top: 0; border-bottom: 1px solid #eee; padding-bottom: 5px; }
      h4 { color: #333; margin-top: 20px; margin-bottom: 5px; background: #f5f5f5; padding: 5px; }
      table { border-collapse: collapse; width: 100%; margin-bottom: 15px; }
      th { background-color: #f0f0f0; padding: 6px; text-align: left; }
      td { padding: 5px; }
      tr:nth-child(even) { background-color: #f9f9f9; }
      p { margin: 5px 0; }
      .note { font-style: italic; color: #666; font-size: 12px; }
      .highlight { background-color: #fffde7; }
    </style>
  `;
}

/**
 * Xuất báo cáo bảo hành ra sheet mới
 * @param {string} reportType Loại báo cáo cần xuất (active, near, expired, unknown, all)
 * @return {boolean} Kết quả xuất báo cáo
 */
function exportWarrantyReport(reportType) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Tạo sheet mới với tên dựa trên loại báo cáo
    let reportSheetName = "Báo cáo Bảo hành";
    switch (reportType) {
      case "active": reportSheetName += " - Còn BH"; break;
      case "near": reportSheetName += " - Sắp hết BH"; break;
      case "expired": reportSheetName += " - Hết BH"; break;
      case "unknown": reportSheetName += " - Không có TT"; break;
      case "all": reportSheetName += " - Tất cả"; break;
    }
    
    // Thêm timestamp để tránh trùng tên
    reportSheetName += " " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmm");
    
    // Tạo sheet mới
    let reportSheet = ss.getSheetByName(reportSheetName);
    if (reportSheet) {
      reportSheet = reportSheet.clear();
    } else {
      reportSheet = ss.insertSheet(reportSheetName);
    }
    
    // Truy vấn dữ liệu thiết bị theo loại báo cáo
    const equipmentData = getWarrantyReportData(reportType);
    
    // Nếu không có dữ liệu, trả về lỗi
    if (!equipmentData || equipmentData.length === 0) {
      throw new Error("Không có dữ liệu thiết bị để xuất báo cáo");
    }
    
    // Tạo header
    const headers = ["Mã TB", "Tên thiết bị", "Loại TB", "Vị trí", "NCC", "Hạn bảo hành"];
    if (reportType === "all") {
      headers.push("Trạng thái");
    }
    
    reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Ghi dữ liệu
    reportSheet.getRange(2, 1, equipmentData.length, equipmentData[0].length).setValues(equipmentData);
    
    // Định dạng sheet
    reportSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    reportSheet.setFrozenRows(1);
    reportSheet.autoResizeColumns(1, headers.length);
    
    // Chuyển đến sheet báo cáo
    ss.setActiveSheet(reportSheet);
    
    return true;
    
  } catch (e) {
    Logger.log(`Lỗi khi xuất báo cáo: ${e}\nStack: ${e.stack}`);
    throw new Error(`Lỗi khi xuất báo cáo: ${e.message}`);
  }
}

/**
 * Lấy dữ liệu báo cáo bảo hành theo loại
 * @param {string} reportType Loại báo cáo
 * @return {Array} Mảng dữ liệu báo cáo
 */
function getWarrantyReportData(reportType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const equipmentSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
  
  // Đọc dữ liệu thiết bị
  const lastRow = equipmentSheet.getLastRow();
  const dataRange = equipmentSheet.getRange(2, 1, lastRow - 1, Math.max(
    COL_EQUIP_ID, COL_EQUIP_NAME, COL_EQUIP_TYPE, COL_EQUIP_LOCATION, 
    COL_EQUIP_SUPPLIER, COL_EQUIP_WARRANTY_END
  ));
  const data = dataRange.getValues();
  
  // Phân loại thiết bị theo tình trạng bảo hành
  const today = new Date();
  const nearExpiryDays = 30;
  const thirtyDaysLater = new Date(today);
  thirtyDaysLater.setDate(today.getDate() + nearExpiryDays);
  
  const result = [];
  const dateFormat = Session.getScriptTimeZone();
  
  // Xử lý từng thiết bị
  data.forEach(row => {
    const id = row[COL_EQUIP_ID - 1];
    const name = row[COL_EQUIP_NAME - 1] || "";
    const type = row[COL_EQUIP_TYPE - 1] || "";
    const location = row[COL_EQUIP_LOCATION - 1] || "";
    const supplier = row[COL_EQUIP_SUPPLIER - 1] || "";
    const warrantyEnd = row[COL_EQUIP_WARRANTY_END - 1];
    
    if (!id) return; // Bỏ qua dòng không có mã thiết bị
    
    let status = "Không có TT";
    let matchesFilter = false;
    
    if (warrantyEnd instanceof Date) {
      let warrantyEndFormat = Utilities.formatDate(warrantyEnd, dateFormat, "dd/MM/yyyy");
      
      if (warrantyEnd > thirtyDaysLater) {
        status = "Còn BH";
        matchesFilter = (reportType === "active" || reportType === "all");
      } else if (warrantyEnd > today) {
        status = "Sắp hết";
        matchesFilter = (reportType === "near" || reportType === "all");
      } else {
        status = "Hết BH";
        matchesFilter = (reportType === "expired" || reportType === "all");
      }
      
      if (matchesFilter) {
        if (reportType === "all") {
          result.push([id, name, type, location, supplier, warrantyEndFormat, status]);
        } else {
          result.push([id, name, type, location, supplier, warrantyEndFormat]);
        }
      }
    } else if (reportType === "unknown" || reportType === "all") {
      if (reportType === "all") {
        result.push([id, name, type, location, supplier, "Không có thông tin", status]);
      } else {
        result.push([id, name, type, location, supplier, "Không có thông tin"]);
      }
    }
  });
  
  return result;
}
