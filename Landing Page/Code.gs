/**
 * Google Apps Script nhận dữ liệu khách hàng từ Landing Page (index.html)
 * và ghi vào Google Sheet đang gắn script này.
 *
 * Cách cài đặt — xem hướng dẫn chi tiết trong tin nhắn chat.
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
