/**
 * Nhận bài nộp của trang "Bài 18. Tính chất chung của kim loại" và ghi vào Google Sheet.
 * Cách cài: xem hướng dẫn giáo viên. Chỉ cần sửa MAT_KHAU bên dưới.
 */
const MAT_KHAU = "doi-mat-khau-nay";   // ← ĐỔI thành mật khẩu riêng của thầy/cô (dùng để xem bảng điểm trên trang)
const TEN_TRANG_TINH = "KetQua";

const TIEU_DE = ["Thời gian nộp", "Họ và tên", "Lớp", "Lần làm", "Điểm (/10)",
  "Câu 1", "Câu 2", "Câu 3", "Câu 4", "Câu 5", "Câu 6", "Câu 7", "Câu 8", "Câu 9", "Câu 10",
  "Mã đề", "Thời gian làm (phút)", "Mã kết quả"];

function layTrangTinh_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TEN_TRANG_TINH);
  if (!sh) {
    sh = ss.insertSheet(TEN_TRANG_TINH);
    sh.appendRow(TIEU_DE);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, TIEU_DE.length).setFontWeight("bold");
  }
  return sh;
}

function traVe_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Học sinh nộp bài
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const per = Array.isArray(d.per) ? d.per.slice(0, 10) : [];
    while (per.length < 10) per.push("");
    layTrangTinh_().appendRow([new Date(), String(d.name || "").slice(0, 100), String(d.cls || "").slice(0, 30),
      d.n, d.total].concat(per).concat([d.made, d.minutes, String(d.code || "").slice(0, 5000)]));
    return traVe_({ ok: true });
  } catch (err) {
    return traVe_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Giáo viên xem bảng điểm (cần mật khẩu)
function doGet(e) {
  if (String((e.parameter || {}).key || "") !== MAT_KHAU) return traVe_({ ok: false, error: "sai_mat_khau" });
  const values = layTrangTinh_().getDataRange().getValues();
  const codes = values.slice(1).map(function (r) { return r[TIEU_DE.length - 1]; }).filter(String);
  return traVe_({ ok: true, codes: codes });
}

// Chạy hàm này một lần trong trình soạn thảo để cấp quyền và tạo trang tính "KetQua"
function caiDat() {
  layTrangTinh_();
}
