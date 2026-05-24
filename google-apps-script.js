/**
 * KOKO Air Service — Google Apps Script
 *
 * วิธี deploy:
 * 1. เปิด Google Sheet: https://docs.google.com/spreadsheets/d/1ClaoMZkJYdHAGBXGcxUrmdnAk5kgLB0V1oUhAEFxLH8
 * 2. เมนู Extensions > Apps Script
 * 3. วางโค้ดนี้ทับโค้ดเดิมทั้งหมด แล้วกด Save
 * 4. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy deployment URL ไปวางใน index.html บรรทัด APPS_SCRIPT_URL
 */

const SHEET_ID = '1ClaoMZkJYdHAGBXGcxUrmdnAk5kgLB0V1oUhAEFxLH8';

const HEADERS = [
  'หมายเลขจอง',
  'วันที่จอง',
  'ชื่อลูกค้า',
  'เบอร์โทร',
  'ที่อยู่',
  'วันที่บริการ',
  'เวลา',
  'รายการบริการ',
  'ระยะทาง',
  'ค่าบริการ (฿)',
  'ค่าเดินทาง (฿)',
  'ยอดรวม (฿)',
];

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();

    if (sheet.getLastRow() === 0) {
      const headerRow = sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#1A1A1A')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    const d = JSON.parse(e.postData.contents);
    sheet.appendRow([
      d.bookingNo,
      d.bookingDate,
      d.name,
      d.phone,
      d.address,
      d.serviceDate,
      d.serviceTime,
      d.services,
      d.distance,
      d.serviceSum,
      d.travelFee,
      d.total,
    ]);

    sheet.autoResizeColumns(1, HEADERS.length);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', bookingNo: d.bookingNo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
