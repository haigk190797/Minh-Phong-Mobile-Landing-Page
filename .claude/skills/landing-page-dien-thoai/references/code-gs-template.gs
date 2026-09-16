/**
 * Template Google Apps Script nhận dữ liệu lead (khách để lại SĐT) từ landing page
 * và ghi vào Google Sheet đang gắn script này.
 *
 * Cách dùng: dán toàn bộ nội dung này vào Apps Script (Tiện ích mở rộng → Apps Script)
 * của Google Sheet, sau đó Triển khai làm "Ứng dụng web" với:
 *   - Thực thi với quyền của: Tôi
 *   - Ai có quyền truy cập: Bất kỳ ai
 *
 * Đổi tên cột / thêm trường (vd chi nhánh, tên khách...) thì sửa mảng appendRow bên dưới,
 * và nhớ thêm input tương ứng trong form HTML (name="..." khớp với e.parameter.<name>).
 */
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Leads');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Leads');
    sheet.appendRow(['Thời gian', 'Số điện thoại', 'Sản phẩm']);
  }

  var phone = e.parameter.phone || '';
  var product = e.parameter.product || '';

  sheet.appendRow([new Date(), phone, product]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}
