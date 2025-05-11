// ==================================
// KHAI BÁO BIẾN TOÀN CỤC 
// ==================================

// === CÁC SHEET CHÍNH ===
const SETTINGS_SHEET_NAME = "Settings / Cấu hình";
const EQUIPMENT_SHEET_NAME = "Danh mục Thiết bị";
const PURCHASE_SHEET_NAME = "Chi tiết Mua Hàng & Nhà Cung Cấp";
const HISTORY_SHEET_NAME = "Lịch sử Bảo trì / Sửa chữa";
const SHEET_PHIEU_CONG_VIEC = "Phiếu Công Việc";
const SHEET_DINH_NGHIA_HE_THONG = "DinhNghiaHeThong";
const SHEET_CHI_TIET_CV_BT = "Chi tiết CV Bảo trì";

// === SHEET CẤU HÌNH (SETTINGS) ===
// Cột Loại TB
const COL_SETTINGS_LOAI_TB_GIATRI = 1;   // Cột A: Loại TB - Giá Trị
const COL_SETTINGS_LOAI_TB_MA = 2;       // Cột B: Loại TB - Mã Viết Tắt

// Cột Vị Trí
const COL_SETTINGS_VITRI_GIATRI = 4;     // Cột D: Vị Trí - Giá Trị
const COL_SETTINGS_VITRI_MA = 5;         // Cột E: Vị Trí - Mã Viết Tắt
const COL_SETTINGS_VITRI_TYPE = 6;       // Cột F: Loại Vị Trí

// Cột Bộ phận
const COL_SETTINGS_BOPHAN_GIATRI = 9;    // Cột I: Bộ phận
const COL_SETTINGS_BOPHAN_MA = 10;       // Cột J: Mã VT (Bộ phận)

// Cột Bộ đếm
const COL_SETTINGS_COUNTER_KEY = 14;     // Cột N: Khóa Bộ Đếm
const COL_SETTINGS_COUNTER_NEXT_NUM = 15;// Cột O: Số Thứ Tự Tiếp Theo

// Cột Hệ thống
const COL_SETTINGS_SYSTEM_CODE = 20;     // Cột T: Mã Hệ thống / Hạng mục
const COL_SETTINGS_SYSTEM_DESC = 21;     // Cột U: Mô tả Hệ thống / Hạng mục

// Danh sách dropdown
const COL_SETTINGS_ASSET_POST_STATUS_LIST_COL = 26; // Cột Z: Danh sách Trạng thái TB sau HĐ

// === SHEET DANH MỤC THIẾT BỊ ===
const COL_EQUIP_ID = 1;                  // Cột A: Mã Thiết Bị
const COL_EQUIP_NAME = 2;                // Cột B: Tên Thiết Bị / Linh Kiện
const COL_EQUIP_TYPE = 3;                // Cột C: Loại Thiết Bị / Linh Kiện
const COL_EQUIP_PARENT_ID = 4;           // Cột D: Mã Thiết bị Cha
const COL_EQUIP_PURCHASE_ID = 10;        // Cột J: Mã Lô Mua Hàng / ID Giao Dịch
const COL_EQUIP_SUPPLIER = 11;           // Cột K: Nhà Cung Cấp
const COL_EQUIP_PURCHASE_DATE = 12;      // Cột L: Ngày Mua
const COL_EQUIP_WARRANTY_END = 13;       // Cột M: Hạn Bảo Hành
const COL_EQUIP_LOCATION = 14;           // Cột N: Vị Trí
const COL_EQUIP_STATUS = 16;             // Cột P: Trạng Thái Hiện Tại
const COL_EQUIP_MAINT_FREQ = 17;         // Cột Q: Tần Suất Bảo Trì Định Kỳ Chính
const COL_EQUIP_MAINT_LAST = 18;         // Cột R: Ngày Bảo Trì Định Kỳ Gần Nhất
const COL_EQUIP_MAINT_NEXT = 19;         // Cột S: Ngày Bảo Trì Định Kỳ Tiếp Theo

// === SHEET MUA HÀNG ===
const COL_PURCHASE_ID = 1;               // Cột A: Mã Lô Mua Hàng / ID Giao Dịch
const COL_PURCHASE_DATE = 2;             // Cột B: Ngày Mua Hàng / Ngày Hợp Đồng
const COL_PURCHASE_CONTRACT_NUM = 3;     // Cột C: Số Hợp Đồng
const COL_PURCHASE_CONTRACT_SIGN_DATE = 4;// Cột D: Ngày Ký Hợp Đồng
const COL_PURCHASE_SUPPLIER = 5;         // Cột E: Tên Nhà Cung Cấp
const COL_PURCHASE_SUPPLIER_CONTACT = 6; // Cột F: Liên Hệ NCC (BH/KT)
const COL_PURCHASE_SUPPLIER_PHONE = 7;   // Cột G: SĐT NCC
const COL_PURCHASE_SUPPLIER_EMAIL = 8;   // Cột H: Email NCC
const COL_PURCHASE_CONTRACT_VALUE = 9;   // Cột I: Tổng Giá Trị Hợp Đồng
const COL_PURCHASE_CURRENCY = 10;        // Cột J: Tiền tệ HĐ
const COL_PURCHASE_PAYMENT_TERM = 11;    // Cột K: Điều khoản Thanh toán Chính
const COL_PURCHASE_DELIVERY_EXPECTED = 12;// Cột L: Ngày Giao hàng Dự kiến
const COL_PURCHASE_ACCEPTANCE_EXPECTED = 13;// Cột M: Ngày Nghiệm thu Dự kiến
const COL_PURCHASE_CONTRACT_STATUS = 14; // Cột N: Trạng thái Hợp đồng
const COL_PURCHASE_PAYMENT_STATUS = 15;  // Cột O: Tình trạng Thanh toán HĐ
const COL_PURCHASE_WARRANTY_TERM_TEXT = 16;// Cột P: Điều Khoản Bảo Hành (Text)
const COL_PURCHASE_WARRANTY_MONTHS = 17; // Cột Q: Thời hạn BH (tháng)
const COL_PURCHASE_WARRANTY_START = 18;  // Cột R: Ngày Bắt Đầu Bảo Hành
const COL_PURCHASE_WARRANTY_END = 19;    // Cột S: Ngày Kết Thúc Bảo Hành
const COL_PURCHASE_WARRANTY_CONTACT = 20;// Cột T: Liên Hệ Bảo Hành NCC
const COL_PURCHASE_WARRANTY_LINK = 21;   // Cột U: Link Quy trình/Yêu cầu BH
const COL_PURCHASE_DOC_LINK = 22;        // Cột V: Link Hóa Đơn / Hợp Đồng Mua
const COL_PURCHASE_NOTES = 23;           // Cột W: Ghi chú Hợp đồng/TT/TĐ

// === SHEET LỊCH SỬ BẢO TRÌ ===
const COL_HISTORY_ID = 1;                // Cột A: ID Lịch Sử
const COL_HISTORY_TARGET_CODE = 2;       // Cột B: Đối tượng / Hệ thống
const COL_HISTORY_TARGET_NAME = 3;       // Cột C: Tên Thiết bị / Hệ thống
const COL_HISTORY_DISPLAY_NAME = 4;      // Cột D: Tên Hiển Thị (TB/HT - Vị trí)
const COL_HISTORY_EXEC_DATE = 5;         // Cột E: Ngày thực hiện / Hoàn thành
const COL_HISTORY_WORK_TYPE = 6;         // Cột F: Loại Công việc
const COL_HISTORY_DESCRIPTION = 7;       // Cột G: Mô tả Công việc / Vấn đề
const COL_HISTORY_PERFORMER = 8;         // Cột H: Người thực hiện / Đơn vị
const COL_HISTORY_EXTERNAL_DETAILS = 9;  // Cột I: Chi tiết ĐV Ngoài / Liên hệ
const COL_HISTORY_COST = 10;             // Cột J: Chi phí (VND)
const COL_HISTORY_STATUS = 11;           // Cột K: Trạng thái CV
const COL_HISTORY_WARRANTY_CHECK = 12;   // Cột L: Theo Bảo Hành? (Checkbox)
const COL_HISTORY_WARRANTY_REQ_ID = 13;  // Cột M: Mã Yêu cầu BH NCC
const COL_HISTORY_WARRANTY_REQ_STAT = 14;// Cột N: Trạng thái Yêu cầu BH
const COL_HISTORY_WARRANTY_REQ_NOTE = 15;// Cột O: Ghi chú Yêu cầu BH
const COL_HISTORY_ASSET_POST_STATUS = 16;// Cột P: Trạng Thái Thiết Bị Sau Hoạt Động
const COL_HISTORY_DETAIL_NOTE = 17;      // Cột Q: Ghi Chú Chi Tiết

// === SHEET ĐỊNH NGHĨA HỆ THỐNG ===
const COL_HT_MA = 1;                     // Cột A: Mã Hệ thống / Hạng mục
const COL_HT_MO_TA = 2;                  // Cột B: Mô tả Hệ thống / Hạng mục

// === SHEET PHIẾU CÔNG VIỆC ===
const COL_PCV_MA_PHIEU = 1;              // A: Mã Phiếu CV
const COL_PCV_NGAY_TAO = 2;              // B: Ngày tạo
const COL_PCV_NGUOI_TAO = 3;             // C: Người tạo
const COL_PCV_NGAY_YC = 4;               // D: Ngày YC/Phát sinh
const COL_PCV_HAN_HT = 5;                // E: Hạn hoàn thành
const COL_PCV_DOI_TUONG = 6;             // F: Đối tượng / Hệ thống
const COL_PCV_TEN_DOI_TUONG = 7;         // G: Tên TB / Mô tả HT
const COL_PCV_VI_TRI = 8;                // H: Vị trí
const COL_PCV_LOAI_CV = 9;               // I: Loại Công việc
const COL_PCV_TAN_SUAT_PM = 10;          // J: Tần suất PM
const COL_PCV_MO_TA_YC = 11;             // K: Mô tả Yêu cầu / Vấn đề
const COL_PCV_UU_TIEN = 12;              // L: Mức độ Ưu tiên
const COL_PCV_NGUOI_GIAO = 13;           // M: Người/Nhóm được giao
const COL_PCV_TRANG_THAI = 14;           // N: Trạng thái Phiếu CV
const COL_PCV_CHI_TIET_NGOAI = 15;       // O: Chi tiết ĐV Ngoài / Liên hệ
const COL_PCV_MO_TA_HT = 16;             // P: Mô tả Hoàn thành / Kết quả
const COL_PCV_VAT_TU = 17;               // Q: Vật tư sử dụng
const COL_PCV_NGAY_HT_THUC_TE = 18;      // R: Ngày Hoàn thành Thực tế
const COL_PCV_TRANG_THAI_TB_SAU = 19;    // S: Trạng thái TB sau HĐ
const COL_PCV_CHI_PHI = 20;              // T: Chi phí (VND)
const COL_PCV_LINK_LS = 21;              // U: Link Lịch sử Hoàn thành
const COL_PCV_GHI_CHU = 22;              // V: Ghi chú Phiếu CV
const COL_PCV_HINH_ANH = 23;             // W: Hình ảnh (từ form báo hỏng)


// === SHEET CHI TIẾT CÔNG VIỆC BẢO TRÌ ===
const COL_CTCV_LOAI_TB = 1;              // Cột A: Loại Thiết Bị
const COL_CTCV_TAN_SUAT = 2;             // Cột B: Tần suất
const COL_CTCV_CONG_VIEC = 3;            // Cột C: Danh sách Công việc

// === CẤU HÌNH EMAIL ADMIN ===
const EMAIL_ADMIN = "longkeutn@gmail.com"; // Email nhận thông báo lỗi
