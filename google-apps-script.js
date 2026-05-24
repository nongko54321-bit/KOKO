/**
 * KOKO Air Service — Google Apps Script
 * วิธี deploy: Extensions > Apps Script > วางโค้ด > Save > Deploy > Manage > New version > Deploy
 */

const SHEET_ID      = '1ClaoMZkJYdHAGBXGcxUrmdnAk5kgLB0V1oUhAEFxLH8';
const BOOKING_SHEET = 'การจอง';

/* ── Column definitions ────────────────────────────────────────────
   group A  (grey)  : ข้อมูลการจอง   col 1–2
   group B  (blue)  : ข้อมูลลูกค้า   col 3–7
   group C  (green) : รายการบริการ    col 8–11
   group D  (amber) : ระยะทาง        col 12
   group E  (red)   : การเงิน        col 13–15
──────────────────────────────────────────────────────────────────── */
const HEADERS = [
  'หมายเลขจอง',        // A1  ─ group A
  'วันที่จอง',          // B1
  'ชื่อ-นามสกุล',      // C1  ─ group B
  'เบอร์โทร',          // D1
  'ที่อยู่',            // E1
  'วันที่บริการ',       // F1
  'เวลา',              // G1
  'รายการบริการ',      // H1  ─ group C
  'ล้างแอร์',          // I1
  'ซ่อมแอร์',          // J1
  'ติดตั้งแอร์',       // K1
  'ระยะทาง',           // L1  ─ group D
  'ค่าบริการ (฿)',     // M1  ─ group E
  'ค่าเดินทาง (฿)',    // N1
  'ยอดรวม (฿)',        // O1
];

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);

    let sheet = ss.getSheetByName(BOOKING_SHEET);
    if (!sheet) { sheet = ss.insertSheet(BOOKING_SHEET); }

    if (sheet.getLastRow() === 0) {
      _setupHeader(sheet);
    }

    const raw = e.postData && e.postData.contents ? e.postData.contents : '{}';
    const d   = JSON.parse(raw);

    sheet.appendRow([
      d.bookingNo    || '',
      d.bookingDate  || '',
      d.name         || '',
      d.phone        || '',
      d.address      || '',
      d.serviceDate  || '',
      d.serviceTime  || '',
      d.services     || '',
      Number(d.washCount)    || 0,
      Number(d.repairCount)  || 0,
      Number(d.installCount) || 0,
      d.distance     || '',
      Number(d.serviceSum)   || 0,
      Number(d.travelFee)    || 0,
      Number(d.total)        || 0,
    ]);

    /* format money columns on new row */
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 13, 1, 3)
      .setNumberFormat('#,##0');

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

function _setupHeader(sheet) {
  sheet.appendRow(HEADERS);
  const n = HEADERS.length;

  /* color groups */
  const groups = [
    { start: 1,  end: 2,  bg: '#3F3F46', fg: '#FFFFFF' },   // booking  (zinc)
    { start: 3,  end: 7,  bg: '#1D4ED8', fg: '#FFFFFF' },   // customer (blue)
    { start: 8,  end: 11, bg: '#065F46', fg: '#FFFFFF' },   // service  (green)
    { start: 12, end: 12, bg: '#92400E', fg: '#FFFFFF' },   // distance (amber)
    { start: 13, end: 15, bg: '#991B1B', fg: '#FFFFFF' },   // finance  (red)
  ];
  groups.forEach(g => {
    sheet.getRange(1, g.start, 1, g.end - g.start + 1)
      .setBackground(g.bg)
      .setFontColor(g.fg)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  });

  /* freeze header + set row height */
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 32);

  /* number format for money columns */
  sheet.getRange(2, 13, 1000, 3).setNumberFormat('#,##0');

  /* set sensible column widths */
  const widths = [130,140,150,110,200,150,70,220,80,80,90,100,120,120,120];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  /* add alternating row banding */
  if (!sheet.getBandings().length) {
    sheet.getRange(2, 1, 1000, n).applyRowBanding(
      SpreadsheetApp.BandingTheme.LIGHT_GREY
    );
  }
}
