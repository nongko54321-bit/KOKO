/**
 * KOKO Air Service — Google Apps Script
 *
 * วิธี deploy / อัปเดต:
 * 1. เปิด Google Sheet → Extensions > Apps Script
 * 2. วางโค้ดนี้ทับโค้ดเดิมทั้งหมด → Save (Ctrl+S)
 * 3. Deploy > Manage deployments > ✏️ แก้ไข > Version: New version > Deploy
 */

const SHEET_ID          = '1ClaoMZkJYdHAGBXGcxUrmdnAk5kgLB0V1oUhAEFxLH8';
const BOOKING_SHEET     = 'การจอง';

const HEADERS = [
  'หมายเลขจอง',
  'วันที่จอง',
  'ชื่อ-นามสกุล',
  'เบอร์โทร',
  'ที่อยู่',
  'วันที่บริการ',
  'เวลา',
  'รายการบริการ',
  'ระยะทาง',
  'ค่าบริการรวม (฿)',
  'ค่าเดินทาง (฿)',
  'ยอดรวมทั้งหมด (฿)',
];

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    // หา sheet "การจอง" ถ้ายังไม่มีให้สร้างใหม่
    let sheet = ss.getSheetByName(BOOKING_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(BOOKING_SHEET);
    }

    // เพิ่ม header row ถ้า sheet ยังว่าง
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold')
        .setBackground('#1A1A1A')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, HEADERS.length);
    }

    const raw = (e.parameter && e.parameter.payload)
      ? e.parameter.payload
      : e.postData.contents;
    const d = JSON.parse(raw);

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
