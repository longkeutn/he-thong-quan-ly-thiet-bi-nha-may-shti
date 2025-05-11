/**
 * File chứa các hàm quản lý bảo mật và phân quyền cho hệ thống [SHT]
 * Quản lý việc khóa/mở khóa sheet và dải ô
 * Phân quyền Admin/Editor
 * Bảo vệ dữ liệu nhạy cảm
 */

// Danh sách email quản trị viên (Admin) có quyền đặc biệt
const ADMIN_EMAILS = [
  "longkeutn@gmail.com", // Thay bằng email của bạn
  // Thêm email admin khác tại đây
];

/**
 * Kiểm tra xem người dùng hiện tại có phải là Admin không
 * @return {boolean} True nếu người dùng hiện tại là Admin
 */
function isCurrentUserAdmin() {
  const userEmail = Session.getActiveUser().getEmail();
  return ADMIN_EMAILS.includes(userEmail);
}
/**
 * Khóa sheet Cấu hình (Settings / Cấu hình) chỉ cho Admin chỉnh sửa.
 */
function protectSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Không tìm thấy sheet "${SETTINGS_SHEET_NAME}"`);
    return;
  }
  // Xóa tất cả bảo vệ cũ (nếu có)
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(p => p.remove());
  // Tạo bảo vệ mới
  const protection = sheet.protect().setDescription('Chỉ Admin được chỉnh sửa');
  protection.addEditors(ADMIN_EMAILS);
  protection.removeEditors(protection.getEditors().filter(e => !ADMIN_EMAILS.includes(e.getEmail())));
  protection.setWarningOnly(false);
  SpreadsheetApp.getUi().alert('Đã khóa sheet Cấu hình. Chỉ Admin có thể chỉnh sửa!');
}

/**
 * Mở khóa sheet Cấu hình cho mọi người.
 */
function unprotectSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Không tìm thấy sheet "${SETTINGS_SHEET_NAME}"`);
    return;
  }
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(p => p.remove());
  SpreadsheetApp.getUi().alert('Đã mở khóa sheet Cấu hình cho mọi người.');
}
/**
 * Khóa cột Mã Thiết Bị (A) chỉ cho Admin chỉnh sửa.
 */
function protectEquipmentIdColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Không tìm thấy sheet "${EQUIPMENT_SHEET_NAME}"`);
    return;
  }
  // Xóa tất cả bảo vệ cũ trên cột A (nếu có)
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => {
    if (p.getRange().getColumn() === COL_EQUIP_ID) p.remove();
  });
  // Tạo bảo vệ mới cho cột A (từ dòng 2 trở đi)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Không có dữ liệu để bảo vệ.');
    return;
  }
  const idRange = sheet.getRange(2, COL_EQUIP_ID, lastRow - 1, 1);
  const protection = idRange.protect().setDescription('Chỉ Admin được chỉnh sửa cột Mã TB');
  protection.addEditors(ADMIN_EMAILS);
  protection.removeEditors(protection.getEditors().filter(e => !ADMIN_EMAILS.includes(e.getEmail())));
  protection.setWarningOnly(false);
  SpreadsheetApp.getUi().alert('Đã khóa cột Mã TB. Chỉ Admin có thể chỉnh sửa!');
}

/**
 * Mở khóa cột Mã Thiết Bị (A) cho mọi người.
 */
function unprotectEquipmentIdColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Không tìm thấy sheet "${EQUIPMENT_SHEET_NAME}"`);
    return;
  }
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => {
    if (p.getRange().getColumn() === COL_EQUIP_ID) p.remove();
  });
  SpreadsheetApp.getUi().alert('Đã mở khóa cột Mã TB cho mọi người.');
}

/**
 * Khóa cột Ngày Bảo Trì Cuối và Tiếp theo (R, S) chỉ cho Admin và Giám sát chỉnh sửa.
 */
function protectMaintenanceDateColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();
  
  if (!sheet) {
    ui.alert(`Không tìm thấy sheet "${EQUIPMENT_SHEET_NAME}"`);
    return;
  }
  
  // Danh sách email được phép chỉnh sửa (Admin + Giám sát)
  const allowedEditors = [...ADMIN_EMAILS, 
    "email_giamsat@example.com" // Thêm email của Giám sát/HCNS vào đây
  ];

  try {
    // Xóa tất cả bảo vệ cũ cho các cột R và S
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    protections.forEach(p => {
      const range = p.getRange();
      const col = range.getColumn();
      if (col === COL_EQUIP_MAINT_LAST || col === COL_EQUIP_MAINT_NEXT) p.remove();
    });
    
    // Lấy phạm vi dữ liệu (từ hàng 2 đến cuối)
    const lastRow = Math.max(2, sheet.getLastRow());
    
    // Bảo vệ cột R - Ngày BT Cuối
    const lastMaintRange = sheet.getRange(2, COL_EQUIP_MAINT_LAST, lastRow - 1, 1);
    const lastMaintProtection = lastMaintRange.protect()
      .setDescription("Cột Ngày BT Cuối - Chỉ Admin/Giám sát được chỉnh sửa");
    lastMaintProtection.addEditors(allowedEditors);
    lastMaintProtection.removeEditors(
      lastMaintProtection.getEditors().filter(e => !allowedEditors.includes(e.getEmail()))
    );
    
    // Bảo vệ cột S - Ngày BT Tiếp theo
    const nextMaintRange = sheet.getRange(2, COL_EQUIP_MAINT_NEXT, lastRow - 1, 1);
    const nextMaintProtection = nextMaintRange.protect()
      .setDescription("Cột Ngày BT Tiếp theo - Chỉ Admin/Giám sát được chỉnh sửa");
    nextMaintProtection.addEditors(allowedEditors);
    nextMaintProtection.removeEditors(
      nextMaintProtection.getEditors().filter(e => !allowedEditors.includes(e.getEmail()))
    );
    
    ui.alert('Đã khóa cột Ngày BT Cuối (R) và Ngày BT Tiếp theo (S). Chỉ Admin và Giám sát có thể chỉnh sửa!');
  } catch (e) {
    Logger.log(`Lỗi khi bảo vệ cột Ngày bảo trì: ${e}`);
    ui.alert(`Lỗi: ${e.message}`);
  }
}

/**
 * Mở khóa cột Ngày Bảo Trì Cuối và Tiếp theo (R, S).
 */
function unprotectMaintenanceDateColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EQUIPMENT_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();
  
  if (!sheet) {
    ui.alert(`Không tìm thấy sheet "${EQUIPMENT_SHEET_NAME}"`);
    return;
  }
  
  // Xóa tất cả bảo vệ cho cột R và S
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => {
    const range = p.getRange();
    const col = range.getColumn();
    if (col === COL_EQUIP_MAINT_LAST || col === COL_EQUIP_MAINT_NEXT) p.remove();
  });
  
  ui.alert('Đã mở khóa cột Ngày BT Cuối (R) và Ngày BT Tiếp theo (S) cho mọi người.');
}

/**
 * Khóa tất cả các dòng tiêu đề của các sheet chính, chỉ cho Admin truy cập.
 * Được gọi từ menu "Quản lý bảo mật & phân quyền".
 */
function protectAllHeaderRows() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Danh sách các sheet cần bảo vệ tiêu đề
    const sheetsToProtect = [
      EQUIPMENT_SHEET_NAME,           // Danh mục Thiết bị
      PURCHASE_SHEET_NAME,            // Chi tiết Mua Hàng & NCC
      HISTORY_SHEET_NAME,             // Lịch sử Bảo trì / Sửa chữa
      SHEET_PHIEU_CONG_VIEC,          // Phiếu Công Việc
      SETTINGS_SHEET_NAME,            // Settings / Cấu hình
      SHEET_DINH_NGHIA_HE_THONG,      // DinhNghiaHeThong
      SHEET_CHI_TIET_CV_BT           // Chi tiết CV Bảo trì
    ];
    
    let protectedCount = 0;
    
    // Lặp qua từng sheet và bảo vệ dòng tiêu đề
    sheetsToProtect.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        Logger.log(`Sheet "${sheetName}" không tồn tại.`);
        return;
      }
      
      // Xóa tất cả các bảo vệ hiện có trên dòng 1
      const existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      existingProtections.forEach(p => {
        const range = p.getRange();
        if (range && range.getRow() === 1 && range.getHeight() === 1) {
          p.remove();
        }
      });
      
      // Tạo bảo vệ mới cho dòng tiêu đề
      const headerRange = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
      const protection = headerRange.protect().setDescription(`Dòng tiêu đề ${sheetName} - Chỉ Admin`);
      
      // Chỉ cho phép Admin chỉnh sửa
      protection.addEditors(ADMIN_EMAILS);
      protection.removeEditors(
        protection.getEditors().filter(e => !ADMIN_EMAILS.includes(e.getEmail()))
      );
      protection.setWarningOnly(false);
      
      protectedCount++;
      Logger.log(`Đã khóa thành công dòng tiêu đề sheet: ${sheetName}`);
    });
    
    ui.alert(`Đã khóa thành công ${protectedCount} dòng tiêu đề sheet.`);
  } 
  catch (e) {
    Logger.log(`Lỗi trong protectAllHeaderRows: ${e}\nStack: ${e.stack}`);
    ui.alert(`Đã xảy ra lỗi: ${e.message}`);
  }
}

/**
 * Khóa các cột quan trọng khác trong các sheet chính (ví dụ: sheet Mua Hàng, Lịch sử, Phiếu CV)
 */
function protectCriticalColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Sheet Chi tiết Mua Hàng & NCC
  const purchaseSheet = ss.getSheetByName(PURCHASE_SHEET_NAME);
  if (purchaseSheet) {
    // Khóa cột A (Mã Lô MH)
    const lastRow = purchaseSheet.getLastRow();
    if (lastRow >= 2) {
      const idRange = purchaseSheet.getRange(2, COL_PURCHASE_ID, lastRow - 1, 1);
      const idProtection = idRange.protect().setDescription("Cột Mã Lô MH - Chỉ Admin/Script");
      idProtection.addEditors(ADMIN_EMAILS);
      idProtection.removeEditors(idProtection.getEditors().filter(e => !ADMIN_EMAILS.includes(e.getEmail())));
    }
    // Khóa cột S (Ngày Kết thúc BH)
    const warrantyEndRange = purchaseSheet.getRange(2, COL_PURCHASE_WARRANTY_END, lastRow - 1, 1);
    const warrantyEndProtection = warrantyEndRange.protect().setDescription("Cột Ngày Kết thúc BH - Chỉ Admin/Script");
    warrantyEndProtection.addEditors(ADMIN_EMAILS);
    warrantyEndProtection.removeEditors(warrantyEndProtection.getEditors().filter(e => !ADMIN_EMAILS.includes(e.getEmail())));
  }

  // Sheet Lịch sử Bảo trì / Sửa chữa
  const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (historySheet) {
    const lastRow = historySheet.getLastRow();
    if (lastRow >= 2) {
      // Khóa cột A (ID Lịch sử)
      const idRange = historySheet.getRange(2, COL_HISTORY_ID, lastRow - 1, 1);
      const idProtection = idRange.protect().setDescription("Cột ID Lịch sử - Chỉ Admin/Script");
      idProtection.addEditors(ADMIN_EMAILS);
      idProtection.removeEditors(idProtection.getEditors().filter(e => !ADMIN_EMAILS.includes(e.getEmail())));
    }
  }

  // Sheet Phiếu Công Việc
  const workOrderSheet = ss.getSheetByName(SHEET_PHIEU_CONG_VIEC);
  if (workOrderSheet) {
    const lastRow = workOrderSheet.getLastRow();
    if (lastRow >= 2) {
      // Khóa cột A (Mã Phiếu CV)
      const idRange = workOrderSheet.getRange(2, COL_PCV_MA_PHIEU, lastRow - 1, 1);
      const idProtection = idRange.protect().setDescription("Cột Mã Phiếu CV - Chỉ Admin/Script");
      idProtection.addEditors(ADMIN_EMAILS);
      idProtection.removeEditors(idProtection.getEditors().filter(e => !ADMIN_EMAILS.includes(e.getEmail())));
      // Khóa cột U (Link Lịch sử)
      const linkRange = workOrderSheet.getRange(2, COL_PCV_LINK_LS, lastRow - 1, 1);
      const linkProtection = linkRange.protect().setDescription("Cột Link Lịch sử - Chỉ Admin/Script");
      linkProtection.addEditors(ADMIN_EMAILS);
      linkProtection.removeEditors(linkProtection.getEditors().filter(e => !ADMIN_EMAILS.includes(e.getEmail())));
    }
  }

  ui.alert("Đã bảo vệ các cột quan trọng trong các sheet chính!");
}
