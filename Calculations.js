// ==================================
// CÁC HÀM TÍNH TOÁN THỜI GIAN VÀ SỐ LIỆU
// ==================================

/**
 * Tính toán ngày bảo trì tiếp theo.
 * @param {Date} lastDate Ngày bảo trì cuối cùng (Đối tượng Date).
 * @param {string} frequency Chuỗi tần suất (VD: "3 tháng", "1 năm", "2 tuần", "Hàng tuần", "Hàng tháng").
 * @return {Date} Ngày bảo trì tiếp theo hoặc null nếu không tính được.
 */
function calculateNextMaintenanceDate(lastDate, frequency) {
  if (!(lastDate instanceof Date) || !frequency) {
    return null;
  }

  try {
    let nextDate = new Date(lastDate.getTime());
    frequency = frequency.toString().toLowerCase().trim();
    const parts = frequency.match(/(\d+)\s*(tuần|tháng|năm)/); // Tìm số và đơn vị

    if (frequency === "hàng tuần" || frequency === "1 tuần") {
      nextDate.setDate(nextDate.getDate() + 7);
    } else if (frequency === "hàng tháng" || frequency === "1 tháng") {
      nextDate.setMonth(nextDate.getMonth() + 1);
    } else if (frequency === "hàng năm" || frequency === "1 năm") {
      nextDate.setFullYear(nextDate.getFullYear() + 1);
    } else if (parts) {
      const number = parseInt(parts[1], 10);
      const unit = parts[2];
      if (!isNaN(number)) {
        if (unit === "tuần") {
          nextDate.setDate(nextDate.getDate() + number * 7);
        } else if (unit === "tháng") {
          nextDate.setMonth(nextDate.getMonth() + number);
        } else if (unit === "năm") {
          nextDate.setFullYear(nextDate.getFullYear() + number);
        } else {
          return null; // Đơn vị không được hỗ trợ
        }
      } else {
        return null; // Không tìm thấy số hợp lệ
      }
    } else {
      return null; // Định dạng tần suất không nhận dạng được
    }
    return nextDate;
  } catch (e) {
    Logger.log(`Lỗi trong hàm calculateNextMaintenanceDate: ${e}`);
    return null;
  }
}

/**
 * Chuyển đổi chuỗi tần suất thành số tháng.
 * @param {string} frequencyString Chuỗi tần suất (VD: "3 tháng").
 * @return {number|null} Số tháng tương ứng hoặc null nếu không nhận dạng được.
 */
function parseFrequencyToMonths(frequencyString) {
  if (!frequencyString || typeof frequencyString !== 'string') return null;
  
  const freqLower = frequencyString.toLowerCase().trim();
  const numMatch = freqLower.match(/^(\d+)/); // Tìm số ở đầu chuỗi
  const number = numMatch ? parseInt(numMatch[1], 10) : null;

  if (freqLower.includes("tháng") && number) {
    return number;
  } else if (freqLower.includes("năm") && number) {
    return number * 12;
  } else if (freqLower.includes("tuần") && number) {
    return number * (1/4.345); // Ước lượng tuần thành tháng
  } else if (freqLower.includes("ngày") && number) {
    return number * (1/30.437); // Ước lượng ngày thành tháng
  } else if (freqLower === "hàng tháng") {
    return 1;
  } else if (freqLower === "hàng năm") {
    return 12;
  } else if (freqLower === "hàng tuần") {
    return 1/4.345;
  }
  
  return null;
}

/**
 * Thêm một số tháng vào một ngày cụ thể.
 * @param {Date} date Ngày bắt đầu.
 * @param {number} months Số tháng cần thêm (có thể là số thập phân).
 * @return {Date} Ngày mới sau khi thêm tháng.
 */
function addMonthsToDate(date, months) {
  if (!(date instanceof Date) || isNaN(date) || typeof months !== 'number' || isNaN(months)) {
    throw new Error("Invalid input for addMonthsToDate");
  }
  
  const originalDate = new Date(date); // Tạo bản sao để không thay đổi ngày gốc
  
  // Xử lý phần nguyên của tháng
  const wholeMonths = Math.floor(months);
  originalDate.setMonth(originalDate.getMonth() + wholeMonths);
  
  // Xử lý phần thập phân (nếu có) bằng cách chuyển thành ngày
  const fractionalMonths = months - wholeMonths;
  if (fractionalMonths > 0) {
    // Ước tính số ngày trong một tháng (30.437)
    const daysToAdd = Math.round(fractionalMonths * 30.437);
    originalDate.setDate(originalDate.getDate() + daysToAdd);
  }
  
  return originalDate;
}
