// ==================================
// TỰ ĐỘNG TẠO PHIẾU CÔNG VIỆC BẢO TRÌ
// ==================================

/**
 * === PHIÊN BẢN 2: Tự động tạo Phiếu CV PM dựa trên Lịch sử hoàn thành ===
 * Quét Lịch sử để tìm lần PM cuối cùng cho mỗi thiết bị.
 * Tính toán các mốc PM tiếp theo cho TẤT CẢ các tần suất được định nghĩa.
 * Tạo Phiếu CV chi tiết nếu đến hạn và chưa có phiếu mở.
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
      woCheckData.forEach(row => {
        const target = row[COL_PCV_DOI_TUONG - 1]?.toString().trim();
        const workType = row[COL_PCV_LOAI_CV - 1]?.toString().trim();
        const pmFrequency = row[COL_PCV_TAN_SUAT_PM - 1]?.toString().trim();
        const status = row[COL_PCV_TRANG_THAI - 1]?.toString().trim();
        if (target && workType === pmWorkType && pmFrequency && status !== "Đã Lưu LS" && status !== "Hủy") {
          const key = `${target}_${pmFrequency}`;
          openWorkOrders[key] = true;
        }
      });
      Logger.log(`[${FUNCTION_NAME}] Tìm thấy ${Object.keys(openWorkOrders).length} Phiếu CV PM đang mở.`);
    }

    // 1.3. Đọc và Xử lý Dữ liệu Lịch sử PM
    const lastPmCompletionMap = {}; // { maTB: Date, ... }
    if (historySheet.getLastRow() >= 2) {
      const lastHistColCheck = Math.max(COL_HISTORY_TARGET_CODE, COL_HISTORY_EXEC_DATE, COL_HISTORY_WORK_TYPE);
      const historyData = historySheet.getRange(2, 1, historySheet.getLastRow() - 1, lastHistColCheck).getValues();
      Logger.log(`[${FUNCTION_NAME}] Đã đọc ${historyData.length} dòng từ ${HISTORY_SHEET_NAME}. Đang lọc và sắp xếp...`);
      
      const filteredHistory = historyData.filter(row => {
        const targetCodeRaw = row[COL_HISTORY_TARGET_CODE - 1];
        let targetCode = "";
        if (targetCodeRaw && typeof targetCodeRaw === 'string') { 
          targetCode = targetCodeRaw.split(" - ")[0].trim(); 
        } else if (targetCodeRaw) { 
          targetCode = targetCodeRaw.toString().trim();
        }

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
      Logger.log(`[${FUNCTION_NAME}] Chuẩn bị ghi ${newWorkOrders.length} Phiếu CV PM mới...`);
      workOrderSheet.getRange(workOrderSheet.getLastRow() + 1, 1, newWorkOrders.length, newWorkOrders[0].length)
                  .setValues(newWorkOrders);
      Logger.log(`[${FUNCTION_NAME}] Đã ghi thành công ${newWorkOrders.length} Phiếu CV PM mới.`);
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
 * Tạo báo cáo theo dõi thiết bị đang xử lý qua BH/thuê ngoài
 */
function createExternalServiceReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const workOrderSheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
    const data = workOrderSheet.getDataRange().getValues();
    
    // Lọc các phiếu đang xử lý bên ngoài
    const pendingOrders = data.filter((row, index) => {
      if (index === 0) return false; // Bỏ qua header
      
      const status = row[COL_PCV_TRANG_THAI - 1];
      return status === "Vượt khả năng nội bộ" || 
             status === "Đang kiểm tra BH" || 
             status === "Chờ đơn vị BH" || 
             status === "Chờ đơn vị ngoài";
    });
    
    // Hiển thị báo cáo trong Dialog
    if (pendingOrders.length === 0) {
      SpreadsheetApp.getUi().alert("Không có phiếu nào đang xử lý bên ngoài");
      return;
    }
    
    let htmlContent = '<table style="width:100%; border-collapse:collapse;">' +
                     '<tr style="background:#f0f0f0;"><th>Mã Phiếu</th><th>Thiết bị</th>' +
                     '<th>Trạng thái</th><th>Thông tin BH/Ngoài</th></tr>';
    
    pendingOrders.forEach(row => {
      htmlContent += `<tr>
        <td>${row[COL_PCV_MA_PHIEU - 1]}</td>
        <td>${row[COL_PCV_DOI_TUONG - 1]} - ${row[COL_PCV_TEN_DOI_TUONG - 1]}</td>
        <td>${row[COL_PCV_TRANG_THAI - 1]}</td>
        <td>${row[COL_PCV_CHI_TIET_NGOAI - 1] || "N/A"}</td>
      </tr>`;
    });
    
    htmlContent += '</table>';
    
    // Hiển thị Dialog
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(600)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Báo cáo Thiết bị đang xử lý bên ngoài");
    
  } catch (e) {
    Logger.log("Lỗi khi tạo báo cáo: " + e);
    SpreadsheetApp.getUi().alert("Lỗi khi tạo báo cáo: " + e);
  }
}

/**
 * Kiểm tra và nhắc nhở các phiếu xử lý bên ngoài quá hạn
 * Gọi bởi trigger hàng ngày
 */
function checkExternalServiceReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
  const data = sheet.getDataRange().getValues();
  
  const today = new Date();
  const threeDaysAgo = new Date(today);
  threeDaysAgo.setDate(today.getDate() - 3);
  
  // Lọc các phiếu xử lý bên ngoài > 3 ngày
  const overdueOrders = [];
  
  for (let i = 1; i < data.length; i++) {
    const status = data[i][COL_PCV_TRANG_THAI - 1];
    const targetStatuses = ["Chờ đơn vị BH", "Chờ đơn vị ngoài"];
    
    if (targetStatuses.includes(status)) {
      const lastUpdateDate = data[i][COL_PCV_NGAY_TAO - 1];
      if (lastUpdateDate < threeDaysAgo) {
        overdueOrders.push({
          row: i + 1,
          id: data[i][COL_PCV_MA_PHIEU - 1],
          equipment: data[i][COL_PCV_DOI_TUONG - 1] + " - " + data[i][COL_PCV_TEN_DOI_TUONG - 1],
          status: status,
          days: Math.floor((today - lastUpdateDate) / (1000 * 60 * 60 * 24))
        });
      }
    }
  }
  
  // Gửi email nếu có phiếu quá hạn
  if (overdueOrders.length > 0) {
    let emailBody = "Các phiếu xử lý bên ngoài quá 3 ngày chưa có phản hồi:\n\n";
    
    overdueOrders.forEach(order => {
      emailBody += `- ${order.id}: ${order.equipment} (${order.status}, ${order.days} ngày)\n`;
    });
    
    MailApp.sendEmail({
      to: EMAIL_ADMIN,
      subject: "[SHT] Nhắc nhở phiếu xử lý bên ngoài quá hạn",
      body: emailBody
    });
  }
}
/**
 * Hiển thị báo cáo thiết bị theo tình trạng bảo hành
 * Được gọi từ menu tùy chỉnh
 */
function showWarrantyReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const equipmentSheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
    
    if (!equipmentSheet) {
      throw new Error(`Không tìm thấy sheet "${EQUIPMENT_SHEET_NAME}"`);
    }
    
    // Đọc dữ liệu thiết bị
    const lastRow = equipmentSheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("Không có dữ liệu thiết bị để báo cáo");
      return;
    }
    
    // Đọc các cột cần thiết
    const dataRange = equipmentSheet.getRange(2, 1, lastRow - 1, Math.max(
      COL_EQUIP_ID, COL_EQUIP_NAME, COL_EQUIP_TYPE, COL_EQUIP_LOCATION, 
      COL_EQUIP_SUPPLIER, COL_EQUIP_WARRANTY_END
    ));
    const data = dataRange.getValues();
    
    // Phân loại thiết bị theo tình trạng bảo hành
    const today = new Date();
    const nearExpiryDays = 30; // Thiết bị sẽ hết hạn bảo hành trong 30 ngày
    const thirtyDaysLater = new Date(today);
    thirtyDaysLater.setDate(today.getDate() + nearExpiryDays);
    
    const categorizedEquipment = {
      activeWarranty: [],    // Còn bảo hành
      nearExpiry: [],        // Sắp hết bảo hành (30 ngày)
      expired: [],           // Hết bảo hành
      noWarrantyInfo: []     // Không có thông tin bảo hành
    };
    
    // Xử lý từng thiết bị
    data.forEach(row => {
      const id = row[COL_EQUIP_ID - 1];
      const name = row[COL_EQUIP_NAME - 1] || "";
      const type = row[COL_EQUIP_TYPE - 1] || "";
      const location = row[COL_EQUIP_LOCATION - 1] || "";
      const supplier = row[COL_EQUIP_SUPPLIER - 1] || "";
      const warrantyEnd = row[COL_EQUIP_WARRANTY_END - 1];
      
      if (!id) return; // Bỏ qua dòng không có mã thiết bị
      
      const item = {
        id: id,
        name: name,
        type: type,
        location: location,
        supplier: supplier,
        warrantyEnd: warrantyEnd
      };
      
      // Phân loại theo tình trạng bảo hành
      if (warrantyEnd instanceof Date) {
        if (warrantyEnd > thirtyDaysLater) {
          categorizedEquipment.activeWarranty.push(item);
        } else if (warrantyEnd > today) {
          categorizedEquipment.nearExpiry.push(item);
        } else {
          categorizedEquipment.expired.push(item);
        }
      } else {
        categorizedEquipment.noWarrantyInfo.push(item);
      }
    });
    
    // Hiển thị kết quả trong dialog
    const htmlTemplate = HtmlService.createTemplateFromFile('WarrantyReportDialog');
    htmlTemplate.data = categorizedEquipment;
    htmlTemplate.today = today;
    htmlTemplate.nearExpiryDays = nearExpiryDays;
    
    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(800)
      .setHeight(600)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Báo cáo Thiết bị theo Bảo hành");
    
  } catch (e) {
    Logger.log(`Lỗi khi hiển thị báo cáo bảo hành: ${e}\nStack: ${e.stack}`);
    SpreadsheetApp.getUi().alert(`Lỗi khi hiển thị báo cáo: ${e.message}`);
  }
}

/**
 * Hiển thị báo cáo thiết bị đang trong quy trình bảo hành với NCC
 */
function showDevicesInWarrantyProcess() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const workOrderSheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
    
    if (!workOrderSheet) throw new Error("Không tìm thấy sheet Phiếu Công Việc");
    
    // Lấy dữ liệu từ sheet Phiếu Công Việc
    const lastRow = workOrderSheet.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert("Không có dữ liệu để báo cáo");
      return;
    }
    
    const dataRange = workOrderSheet.getRange(2, 1, lastRow - 1, 
      Math.max(COL_PCV_MA_PHIEU, COL_PCV_DOI_TUONG, COL_PCV_TEN_DOI_TUONG, 
               COL_PCV_TRANG_THAI, COL_PCV_CHI_TIET_NGOAI, COL_PCV_GHI_CHU));
    const data = dataRange.getValues();
    
    // Lọc các công việc đang trong quá trình bảo hành
    const warrantyItems = data.filter(row => {
      const status = row[COL_PCV_TRANG_THAI - 1] || "";
      const externalDetails = row[COL_PCV_CHI_TIET_NGOAI - 1] || "";
      return status.includes("BH") || externalDetails.includes("bảo hành");
    });
    
    if (warrantyItems.length === 0) {
      SpreadsheetApp.getUi().alert("Không có thiết bị nào đang trong quá trình bảo hành");
      return;
    }
    
    // Hiển thị report trong dialog
    let htmlContent = `
      <style>
        body { font-family: Arial, sans-serif; padding: 15px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        h3 { color: #333; }
        .status-badge {
          display: inline-block;
          padding: 3px 8px;
          border-radius: 3px;
          font-size: 12px;
          color: white;
        }
        .status-pending { background-color: #ffc107; }
        .status-inprogress { background-color: #17a2b8; }
        .status-completed { background-color: #28a745; }
      </style>
      <h3>Thiết bị đang trong quá trình bảo hành</h3>
      <table>
        <tr>
          <th>Mã Phiếu</th>
          <th>Thiết bị</th>
          <th>Trạng thái</th>
          <th>Chi tiết BH/NCC</th>
          <th>Thao tác</th>
        </tr>
    `;
    
    warrantyItems.forEach((item, index) => {
      const status = item[COL_PCV_TRANG_THAI - 1] || "";
      let statusClass = "status-pending";
      
      if (status.includes("xử lý")) {
        statusClass = "status-inprogress";
      } else if (status.includes("xong") || status.includes("hoàn thành")) {
        statusClass = "status-completed";
      }
      
      htmlContent += `
        <tr>
          <td>${item[COL_PCV_MA_PHIEU - 1] || ""}</td>
          <td>${item[COL_PCV_TEN_DOI_TUONG - 1] || ""} (${item[COL_PCV_DOI_TUONG - 1] || ""})</td>
          <td><span class="status-badge ${statusClass}">${status}</span></td>
          <td>${item[COL_PCV_CHI_TIET_NGOAI - 1] || ""}</td>
          <td>
            <button onclick="viewDetails(${index})">Xem chi tiết</button>
          </td>
        </tr>
      `;
    });
    
    htmlContent += `</table>
    <script>
      function viewDetails(index) {
        google.script.run.viewWarrantyDetails(${JSON.stringify(warrantyItems)}, index);
      }
    </script>`;
    
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(700)
        .setHeight(500);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Thiết bị đang trong quá trình bảo hành');
    
  } catch (e) {
    Logger.log(`Lỗi trong showDevicesInWarrantyProcess: ${e}`);
    SpreadsheetApp.getUi().alert(`Lỗi khi tạo báo cáo: ${e.message}`);
  }
}

/**
 * Hiển thị chi tiết thiết bị bảo hành
 * @param {Array} items Danh sách thiết bị đang bảo hành
 * @param {number} index Chỉ số thiết bị cần xem chi tiết
 */
function viewWarrantyDetails(items, index) {
  try {
    const item = items[index];
    if (!item) {
      SpreadsheetApp.getUi().alert("Không tìm thấy thông tin chi tiết thiết bị");
      return;
    }
    
    // Tạo HTML hiển thị thông tin chi tiết
    let htmlContent = `
      <style>
        body { font-family: Arial, sans-serif; padding: 15px; }
        .detail-card { border: 1px solid #ddd; padding: 15px; border-radius: 8px; }
        .detail-title { font-size: 16px; font-weight: bold; margin-bottom: 15px; }
        .detail-section { margin-bottom: 10px; }
        .detail-label { font-weight: bold; margin-bottom: 5px; }
        .detail-value { margin-left: 10px; }
        .button-group { margin-top: 20px; text-align: center; }
        .button { padding: 8px 15px; margin: 0 5px; }
      </style>
      
      <div class="detail-card">
        <div class="detail-title">Chi tiết thiết bị trong quá trình bảo hành</div>
        
        <div class="detail-section">
          <div class="detail-label">Mã phiếu:</div>
          <div class="detail-value">${item[COL_PCV_MA_PHIEU - 1] || ""}</div>
        </div>
        
        <div class="detail-section">
          <div class="detail-label">Thiết bị:</div>
          <div class="detail-value">${item[COL_PCV_TEN_DOI_TUONG - 1] || ""} (${item[COL_PCV_DOI_TUONG - 1] || ""})</div>
        </div>
        
        <div class="detail-section">
          <div class="detail-label">Trạng thái hiện tại:</div>
          <div class="detail-value">${item[COL_PCV_TRANG_THAI - 1] || ""}</div>
        </div>
        
        <div class="detail-section">
          <div class="detail-label">Chi tiết bảo hành/NCC:</div>
          <div class="detail-value">${item[COL_PCV_CHI_TIET_NGOAI - 1] || "Không có thông tin"}</div>
        </div>
        
        <div class="detail-section">
          <div class="detail-label">Ghi chú:</div>
          <div class="detail-value">${item[COL_PCV_GHI_CHU - 1] || "Không có ghi chú"}</div>
        </div>
        
        <div class="button-group">
          <button class="button" onclick="google.script.host.close()">Đóng</button>
        </div>
      </div>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(500)
        .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Chi tiết thiết bị bảo hành');
    
  } catch (e) {
    Logger.log(`Lỗi trong viewWarrantyDetails: ${e}`);
    SpreadsheetApp.getUi().alert(`Lỗi khi hiển thị chi tiết: ${e.message}`);
  }
}
