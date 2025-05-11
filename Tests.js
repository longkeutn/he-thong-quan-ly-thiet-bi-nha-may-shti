// ==================================
// HÀM TIỆN ÍCH ĐỂ CHẠY THỬ (Optional)
// ==================================
function testGenerateId() {
  const type = "MTB20L"; // Lấy từ Cột B Settings
  const location = "NM-KTRON"; // Lấy từ Cột E Settings
  const newId = generateEquipmentId(type, location);
  if (newId) {
    Logger.log("ID thử nghiệm: " + newId);
  } else {
     Logger.log("Tạo ID thử nghiệm thất bại.");
  }
}

function testGetPurchaseInfo() {
  const purchaseId = "PO-TEST-001"; // Thay bằng Mã Lô Mua Hàng có thật trong sheet của bạn
  const info = getPurchaseInfo(purchaseId);
  if (info) {
    Logger.log(`NCC: ${info.supplier}, Ngày mua: ${info.purchaseDate}, Hạn BH: ${info.warrantyEnd}`);
  } else {
    Logger.log("Không tìm thấy thông tin mua hàng.");
  }
}

function testCalculateNextDate() {
  const last = new Date(2025, 0, 15); // 15 tháng 1 năm 2025
  const freq = "3 tháng";
  const next = calculateNextMaintenanceDate(last, freq);
   if (next) {
     Logger.log(`Ngày cuối: ${last.toLocaleDateString()}, Tần suất: ${freq}, Ngày tiếp: ${next.toLocaleDateString()}`);
   } else {
      Logger.log("Tính ngày tiếp theo thất bại.");
   }
}
// ==================================