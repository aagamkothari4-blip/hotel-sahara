// ================================================================
//  HOTEL SAHARA — Google Apps Script Backend
//  Paste this entire file into Google Apps Script editor
//  Then run setupSheets() once, then deploy as Web App
// ================================================================

const CONFIG = {
  HOTEL_EMAIL:   'bookingsahara@rediffmail.com',
  HOTEL_NAME:    'Hotel Sahara, Pune',
  HOTEL_PHONE:   '020-25655405/6/8/9',
  HOTEL_WHATSAPP:'9822393889',
  TIMEZONE:      'Asia/Kolkata',
};

// ── ROOM MASTER ──────────────────────────────────────────────────
// Update "type" to match your actual room layout
// status: Available | Permanent | Non-Existent | Maintenance
const ROOMS = [
  // Floor 1
  {no:101, floor:1, type:'Executive', status:'Permanent',     note:'Permanently occupied'},
  {no:102, floor:1, type:'Executive', status:'Available',     note:''},
  {no:103, floor:1, type:'Executive', status:'Available',     note:''},
  {no:104, floor:1, type:'Deluxe',    status:'Available',     note:''},
  {no:105, floor:1, type:'Executive', status:'Available',     note:''},
  {no:106, floor:1, type:'Executive', status:'Available',     note:''},
  {no:107, floor:1, type:'Executive', status:'Available',     note:''},
  {no:108, floor:1, type:'Deluxe',    status:'Available',     note:''},
  {no:109, floor:1, type:'Executive', status:'Available',     note:''},
  {no:110, floor:1, type:'Executive', status:'Available',     note:''},
  {no:111, floor:1, type:'N/A',       status:'Non-Existent',  note:'Does not exist'},
  {no:112, floor:1, type:'Deluxe',    status:'Available',     note:''},
  // Floor 2
  {no:201, floor:2, type:'Executive', status:'Available', note:''},
  {no:202, floor:2, type:'Executive', status:'Available', note:''},
  {no:203, floor:2, type:'Executive', status:'Available', note:''},
  {no:204, floor:2, type:'Deluxe',    status:'Available', note:''},
  {no:205, floor:2, type:'Executive', status:'Available', note:''},
  {no:206, floor:2, type:'Executive', status:'Available', note:''},
  {no:207, floor:2, type:'Executive', status:'Available', note:''},
  {no:208, floor:2, type:'Deluxe',    status:'Available', note:''},
  {no:209, floor:2, type:'Executive', status:'Available', note:''},
  {no:210, floor:2, type:'Executive', status:'Available', note:''},
  {no:211, floor:2, type:'Executive', status:'Available', note:''},
  {no:212, floor:2, type:'Deluxe',    status:'Available', note:''},
  // Floor 3
  {no:301, floor:3, type:'Executive', status:'Available', note:''},
  {no:302, floor:3, type:'Executive', status:'Available', note:''},
  {no:303, floor:3, type:'Executive', status:'Available', note:''},
  {no:304, floor:3, type:'Deluxe',    status:'Available', note:''},
  {no:305, floor:3, type:'Executive', status:'Available', note:''},
  {no:306, floor:3, type:'Executive', status:'Available', note:''},
  {no:307, floor:3, type:'Executive', status:'Available', note:''},
  {no:308, floor:3, type:'Deluxe',    status:'Available', note:''},
  {no:309, floor:3, type:'Executive', status:'Available', note:''},
  {no:310, floor:3, type:'Executive', status:'Available', note:''},
  {no:311, floor:3, type:'Executive', status:'Available', note:''},
  {no:312, floor:3, type:'Deluxe',    status:'Available', note:''},
];

// ================================================================
//  RECEIVE BOOKING FROM WEBSITE (HTTP POST)
// ================================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    const bookingId = data.bookingId || generateId();
    addToBookingsSheet(ss, data, bookingId);
    sendHotelEmail(data, bookingId);
    refreshAvailability(ss);

    return jsonResponse({success: true, bookingId: bookingId});
  } catch(err) {
    return jsonResponse({success: false, error: err.message});
  }
}

// Allow CORS preflight
function doGet(e) {
  return jsonResponse({status: 'Hotel Sahara Booking System — running'});
}

// ================================================================
//  ADD BOOKING ROW
// ================================================================
function addToBookingsSheet(ss, data, bookingId) {
  const sheet = ss.getSheetByName('📋 Bookings');
  const now   = new Date();
  const fmt   = d => d ? Utilities.formatDate(new Date(d+'T00:00:00'), CONFIG.TIMEZONE, 'dd MMM yyyy') : '';

  sheet.appendRow([
    bookingId,
    Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd MMM yyyy HH:mm'),
    data.guestName   || '',
    data.guestPhone  || '',
    data.guestEmail  || '',
    cap(data.roomType) + ' AC',
    cap(data.occupancy),
    '',                                       // Room No — hotel assigns at check-in
    fmt(data.checkIn),
    fmt(data.checkOut),
    data.nights      || '',
    data.extraPersons|| 0,
    '₹' + num(data.roomTotal),
    '₹' + num(data.gst),
    '₹' + num(data.totalAmount),
    data.paymentId   || 'DEMO',
    'Website',
    'Confirmed',
    data.requests    || '',
  ]);

  // Highlight new row in light green
  const row = sheet.getLastRow();
  sheet.getRange(row, 1, 1, 19)
    .setBackground('#DCFCE7')
    .setBorder(true, true, true, true, false, false, '#BBF7D0', SpreadsheetApp.BorderStyle.SOLID);
}

// ================================================================
//  SEND EMAIL TO HOTEL
// ================================================================
function sendHotelEmail(data, bookingId) {
  const fmt = d => {
    if (!d) return '—';
    const date = new Date(d + 'T00:00:00');
    return Utilities.formatDate(date, CONFIG.TIMEZONE, 'EEE, d MMM yyyy');
  };

  const subject = `🏨 New Booking: ${bookingId} | ${data.guestName} | Check-in ${fmt(data.checkIn)}`;

  const htmlBody = `
<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
  <div style="background:#1B2A4A;padding:24px;text-align:center">
    <h2 style="color:#C9A84C;margin:0;font-size:22px">🏨 Hotel Sahara</h2>
    <p style="color:rgba(255,255,255,0.7);margin:4px 0 0;font-size:13px">New Booking Received</p>
  </div>
  <div style="background:#F8F5EF;padding:24px">
    <table style="width:100%;border-collapse:collapse">
      <tr><td colspan="2" style="background:#1B2A4A;color:#C9A84C;font-weight:bold;padding:10px 14px;font-size:13px;letter-spacing:1px">BOOKING ID: ${bookingId}</td></tr>
      <tr style="background:#fff"><td style="padding:10px 14px;color:#666;font-size:13px;width:40%">Received At</td><td style="padding:10px 14px;font-size:13px;font-weight:bold">${new Date().toLocaleString('en-IN',{timeZone:'Asia/Kolkata'})}</td></tr>

      <tr><td colspan="2" style="background:#E5E7EB;color:#374151;font-weight:bold;padding:8px 14px;font-size:12px;letter-spacing:1px">GUEST</td></tr>
      <tr style="background:#fff"><td style="padding:10px 14px;color:#666;font-size:13px">Name</td><td style="padding:10px 14px;font-size:13px;font-weight:bold">${data.guestName}</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:10px 14px;color:#666;font-size:13px">Phone</td><td style="padding:10px 14px;font-size:13px"><a href="tel:${data.guestPhone}">${data.guestPhone}</a></td></tr>
      <tr style="background:#fff"><td style="padding:10px 14px;color:#666;font-size:13px">Email</td><td style="padding:10px 14px;font-size:13px"><a href="mailto:${data.guestEmail}">${data.guestEmail}</a></td></tr>

      <tr><td colspan="2" style="background:#E5E7EB;color:#374151;font-weight:bold;padding:8px 14px;font-size:12px;letter-spacing:1px">BOOKING DETAILS</td></tr>
      <tr style="background:#fff"><td style="padding:10px 14px;color:#666;font-size:13px">Room Type</td><td style="padding:10px 14px;font-size:13px;font-weight:bold">${cap(data.roomType)} AC — ${cap(data.occupancy)} Occupancy</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:10px 14px;color:#666;font-size:13px">Check-In</td><td style="padding:10px 14px;font-size:13px;font-weight:bold">${fmt(data.checkIn)} &nbsp;at 1:00 PM</td></tr>
      <tr style="background:#fff"><td style="padding:10px 14px;color:#666;font-size:13px">Check-Out</td><td style="padding:10px 14px;font-size:13px;font-weight:bold">${fmt(data.checkOut)} &nbsp;at 11:00 AM</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:10px 14px;color:#666;font-size:13px">Nights</td><td style="padding:10px 14px;font-size:13px">${data.nights}</td></tr>
      <tr style="background:#fff"><td style="padding:10px 14px;color:#666;font-size:13px">Extra Persons</td><td style="padding:10px 14px;font-size:13px">${data.extraPersons || 0}</td></tr>

      <tr><td colspan="2" style="background:#E5E7EB;color:#374151;font-weight:bold;padding:8px 14px;font-size:12px;letter-spacing:1px">PAYMENT</td></tr>
      <tr style="background:#fff"><td style="padding:10px 14px;color:#666;font-size:13px">Room Charges</td><td style="padding:10px 14px;font-size:13px">₹${num(data.roomTotal)}</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:10px 14px;color:#666;font-size:13px">GST</td><td style="padding:10px 14px;font-size:13px">₹${num(data.gst)}</td></tr>
      <tr style="background:#DCFCE7"><td style="padding:12px 14px;color:#166534;font-size:14px;font-weight:bold">Total Paid</td><td style="padding:12px 14px;font-size:18px;font-weight:bold;color:#166534">₹${num(data.totalAmount)}</td></tr>
      <tr style="background:#fff"><td style="padding:10px 14px;color:#666;font-size:13px">Payment ID</td><td style="padding:10px 14px;font-size:12px;font-family:monospace">${data.paymentId || 'DEMO'}</td></tr>

      ${data.requests ? `<tr><td colspan="2" style="background:#FEF3C7;padding:12px 14px;font-size:13px"><strong>Special Requests:</strong> ${data.requests}</td></tr>` : ''}
    </table>

    <div style="margin-top:20px;padding:16px;background:#fff;border-left:4px solid #C9A84C;border-radius:4px">
      <p style="margin:0;color:#374151;font-size:13px">
        <strong>⚠️ Action Required:</strong> Please assign a room number in the
        <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" style="color:#1B2A4A">Bookings Sheet</a>
        before the guest's check-in date.
      </p>
    </div>
  </div>
  <div style="background:#1B2A4A;padding:16px;text-align:center">
    <p style="color:rgba(255,255,255,0.6);font-size:12px;margin:0">
      Hotel Sahara · Senapati Bapat Road, Pune · ${CONFIG.HOTEL_PHONE}
    </p>
  </div>
</div>`;

  GmailApp.sendEmail(CONFIG.HOTEL_EMAIL, subject, '', {
    htmlBody:  htmlBody,
    name:      'Hotel Sahara Booking System',
    replyTo:   data.guestEmail || CONFIG.HOTEL_EMAIL,
  });
}

// ================================================================
//  REFRESH AVAILABILITY GRID
// ================================================================
function refreshAvailability(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookingSheet = ss.getSheetByName('📋 Bookings');
  const availSheet   = ss.getSheetByName('📅 Availability');

  const allBookings = bookingSheet.getDataRange().getValues().slice(1); // skip header
  // bookingId(0), date(1), name(2), phone(3), email(4), roomType(5), occ(6),
  // roomNo(7), checkIn(8), checkOut(9), nights(10)... status(17)

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Get date headers from row 1 (columns 4 onwards)
  const dateHeaders = availSheet.getRange(1, 4, 1, 90).getValues()[0];

  // Get room list from column A (rows 2+)
  const roomRows = availSheet.getRange(2, 1, availSheet.getLastRow()-1, 3).getValues();

  roomRows.forEach((roomRow, rIdx) => {
    const roomNo = String(roomRow[0]);
    if (!roomNo || roomNo === '111') return;

    dateHeaders.forEach((dateHeader, dIdx) => {
      const col  = dIdx + 4;
      const row  = rIdx + 2;
      const cell = availSheet.getRange(row, col);

      if (cell.getValue() === 'PERMANENT') return;

      const cellDate = new Date(today);
      cellDate.setDate(today.getDate() + dIdx);

      // Find any confirmed booking for this room on this date
      const booking = allBookings.find(b => {
        const bRoomNo = String(b[7]).trim();
        const status  = String(b[17]);
        if (status === 'Cancelled' || bRoomNo === '' || bRoomNo === 'TBD') return false;
        if (bRoomNo !== roomNo) return false;

        const checkIn  = parseDateStr(b[8]);
        const checkOut = parseDateStr(b[9]);
        if (!checkIn || !checkOut) return false;
        return cellDate >= checkIn && cellDate < checkOut;
      });

      if (booking) {
        const guestName = booking[2] ? booking[2].split(' ')[0] : 'BOOKED';
        cell.setValue(guestName)
            .setBackground('#FEE2E2')
            .setFontColor('#991B1B')
            .setFontWeight('bold');
      } else {
        cell.setValue('')
            .setBackground('#D1FAE5')
            .setFontColor('#166534')
            .setFontWeight('normal');
      }
    });
  });
}

// ================================================================
//  MANUAL BOOKING TRIGGER (run from Bookings sheet after manual entry)
//  Select the row and run this function, or just type and save
// ================================================================
function onBookingManualEntry() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 Bookings');
  const row = sheet.getActiveCell().getRow();
  if (row < 2) return;

  const data = sheet.getRange(row, 1, 1, 19).getValues()[0];
  if (!data[0]) {
    // Auto-generate booking ID if empty
    sheet.getRange(row, 1).setValue(generateId());
  }
  if (!data[1]) {
    // Set received timestamp
    sheet.getRange(row, 2).setValue(
      Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd MMM yyyy HH:mm')
    );
  }
  if (!data[16]) {
    sheet.getRange(row, 17).setValue('Manual'); // Source
  }
  if (!data[17]) {
    sheet.getRange(row, 18).setValue('Confirmed'); // Status
  }

  refreshAvailability(ss);
  SpreadsheetApp.getUi().alert('✅ Row saved and availability updated!');
}

// ================================================================
//  ONE-TIME SETUP — Run this once to create all sheets
// ================================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('🏨 Hotel Sahara — Booking System');

  createBookingsSheet(ss);
  createRoomsSheet(ss);
  createAvailabilitySheet(ss);
  createTodaySheet(ss);

  // Remove default 'Sheet1' if it exists
  const def = ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);

  // Add menu
  addMenu();

  SpreadsheetApp.getUi().alert(
    '✅ Setup complete!\n\n' +
    'Next steps:\n' +
    '1. Go to Extensions → Apps Script\n' +
    '2. Click Deploy → New Deployment\n' +
    '3. Type: Web App | Execute as: Me | Who has access: Anyone\n' +
    '4. Copy the Web App URL\n' +
    '5. Paste it into js/booking.js as googleSheetsUrl'
  );
}

function createBookingsSheet(ss) {
  let sheet = ss.getSheetByName('📋 Bookings');
  if (!sheet) sheet = ss.insertSheet('📋 Bookings', 0);
  else sheet.clearContents();

  const headers = [
    'Booking ID', 'Received At', 'Guest Name', 'Phone', 'Email',
    'Room Type', 'Occupancy', '🔑 Room No', 'Check-In', 'Check-Out',
    'Nights', 'Extra', 'Room Charges', 'GST', 'Total Paid',
    'Payment ID', 'Source', 'Status', 'Special Requests'
  ];

  const hRow = sheet.getRange(1, 1, 1, headers.length);
  hRow.setValues([headers])
      .setBackground('#1B2A4A')
      .setFontColor('#C9A84C')
      .setFontWeight('bold')
      .setFontSize(10);

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 165);   // Booking ID
  sheet.setColumnWidth(2, 145);   // Date
  sheet.setColumnWidth(3, 175);   // Guest name
  sheet.setColumnWidth(4, 130);   // Phone
  sheet.setColumnWidth(5, 200);   // Email
  sheet.setColumnWidth(6, 115);   // Room type
  sheet.setColumnWidth(7, 95);    // Occupancy
  sheet.setColumnWidth(8, 95);    // Room No
  sheet.setColumnWidth(9, 110);   // Check-in
  sheet.setColumnWidth(10, 110);  // Check-out
  sheet.setColumnWidth(11, 65);   // Nights
  sheet.setColumnWidth(12, 60);   // Extra
  sheet.setColumnWidth(13, 115);  // Charges
  sheet.setColumnWidth(14, 90);   // GST
  sheet.setColumnWidth(15, 115);  // Total
  sheet.setColumnWidth(16, 175);  // Payment ID
  sheet.setColumnWidth(17, 90);   // Source
  sheet.setColumnWidth(18, 100);  // Status
  sheet.setColumnWidth(19, 250);  // Requests

  // Add data validation for Status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Confirmed','Pending','Checked-In','Checked-Out','Cancelled','No-Show'], true)
    .build();
  sheet.getRange(2, 18, 200, 1).setDataValidation(statusRule);

  // Add data validation for Source column
  const sourceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Website','Walk-in','Phone','OTA - MakeMyTrip','OTA - OYO','OTA - Booking.com','Other'], true)
    .build();
  sheet.getRange(2, 17, 200, 1).setDataValidation(sourceRule);
}

function createRoomsSheet(ss) {
  let sheet = ss.getSheetByName('🏨 Rooms');
  if (!sheet) sheet = ss.insertSheet('🏨 Rooms', 1);
  else sheet.clearContents();

  const headers = ['Room No', 'Floor', 'Type', 'Status', 'Tariff (Single)', 'Tariff (Double)', 'Notes'];
  sheet.getRange(1, 1, 1, headers.length)
       .setValues([headers])
       .setBackground('#1B2A4A')
       .setFontColor('#C9A84C')
       .setFontWeight('bold');
  sheet.setFrozenRows(1);

  const TARIFF = { 'Executive': [2300, 2800], 'Deluxe': [2500, 3000], 'N/A': [0, 0] };

  const rows = ROOMS.map(r => [
    r.no, r.floor, r.type, r.status,
    r.type !== 'N/A' ? '₹' + TARIFF[r.type][0] : '',
    r.type !== 'N/A' ? '₹' + TARIFF[r.type][1] : '',
    r.note
  ]);
  sheet.getRange(2, 1, rows.length, 7).setValues(rows);

  // Color coding
  ROOMS.forEach((r, i) => {
    const bg = r.status === 'Permanent'    ? '#FEE2E2' :
               r.status === 'Non-Existent' ? '#E5E7EB' :
               r.type   === 'Deluxe'       ? '#FEF9C3' : '#F0F9FF';
    sheet.getRange(i+2, 1, 1, 7).setBackground(bg);
  });

  sheet.setColumnWidths(1, 7, 120);
  sheet.setColumnWidth(7, 220);

  // Add Floor separator rows
  [1, 13, 25].forEach(divRow => {
    const floorLabel = ['Floor 1 (101–112)', 'Floor 2 (201–212)', 'Floor 3 (301–312)'];
    // Already separated by natural row ordering, just color floor groups
  });
}

function createAvailabilitySheet(ss) {
  let sheet = ss.getSheetByName('📅 Availability');
  if (!sheet) sheet = ss.insertSheet('📅 Availability', 2);
  else sheet.clearContents();

  // Build 90-day date headers
  const today = new Date();
  const headers = ['Room', 'Type', 'Floor'];
  for (let i = 0; i < 90; i++) {
    const d = new Date(today);
    d.setDate(today.getDate() + i);
    headers.push(Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd MMM'));
  }

  sheet.getRange(1, 1, 1, headers.length)
       .setValues([headers])
       .setBackground('#1B2A4A')
       .setFontColor('#C9A84C')
       .setFontWeight('bold')
       .setFontSize(9);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);
  sheet.setColumnWidth(1, 65);
  sheet.setColumnWidth(2, 85);
  sheet.setColumnWidth(3, 70);
  for (let c = 4; c <= headers.length; c++) sheet.setColumnWidth(c, 62);

  // Room rows (skip non-existent 111)
  const bookable = ROOMS.filter(r => r.status !== 'Non-Existent');
  bookable.forEach((r, i) => {
    const row = i + 2;
    sheet.getRange(row, 1).setValue(r.no);
    sheet.getRange(row, 2).setValue(r.type);
    sheet.getRange(row, 3).setValue('F' + r.floor);

    const typeBg = r.type === 'Deluxe' ? '#FEF9C3' : '#F0F9FF';
    sheet.getRange(row, 1, 1, 3).setBackground(typeBg);

    if (r.status === 'Permanent') {
      sheet.getRange(row, 4, 1, 90)
           .setValue('PERMANENT')
           .setBackground('#FEE2E2')
           .setFontColor('#991B1B')
           .setFontWeight('bold');
    } else {
      // Default all cells to green (available)
      sheet.getRange(row, 4, 1, 90)
           .setBackground('#D1FAE5')
           .setFontColor('#166534');
    }
  });

  // Add legend
  const legendRow = bookable.length + 3;
  sheet.getRange(legendRow, 1).setValue('LEGEND:');
  sheet.getRange(legendRow, 2).setValue('✅ Green = Available').setBackground('#D1FAE5');
  sheet.getRange(legendRow, 3).setValue('🔴 Red = Booked/Name').setBackground('#FEE2E2');
  sheet.getRange(legendRow, 4).setValue('🟡 Yellow = Deluxe Room').setBackground('#FEF9C3');
  sheet.getRange(legendRow, 5).setValue('🔵 Blue = Executive Room').setBackground('#F0F9FF');
}

function createTodaySheet(ss) {
  let sheet = ss.getSheetByName('🗓️ Today');
  if (!sheet) sheet = ss.insertSheet('🗓️ Today', 3);
  else sheet.clearContents();

  updateTodaySheet();
}

function updateTodaySheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('🗓️ Today');
  const bSheet= ss.getSheetByName('📋 Bookings');
  if (!sheet || !bSheet) return;

  sheet.clearContents();
  const today    = new Date();
  const todayStr = Utilities.formatDate(today, CONFIG.TIMEZONE, 'dd MMM yyyy');

  sheet.getRange(1, 1).setValue(`🗓️ Daily Summary — ${todayStr}`)
       .setFontSize(14).setFontWeight('bold').setFontColor('#1B2A4A');
  sheet.getRange(2, 1).setValue('Last updated: ' + new Date().toLocaleTimeString())
       .setFontColor('#9CA3AF').setFontSize(10);

  const bookings = bSheet.getDataRange().getValues().slice(1);
  const todayDate = new Date(today); todayDate.setHours(0,0,0,0);

  const checkIns  = bookings.filter(b => {
    const ci = parseDateStr(b[8]);
    return ci && ci.getTime() === todayDate.getTime() && b[17] !== 'Cancelled';
  });
  const checkOuts = bookings.filter(b => {
    const co = parseDateStr(b[9]);
    return co && co.getTime() === todayDate.getTime() && b[17] !== 'Cancelled';
  });
  const stayOvers = bookings.filter(b => {
    const ci = parseDateStr(b[8]);
    const co = parseDateStr(b[9]);
    return ci && co && ci < todayDate && co > todayDate && b[17] !== 'Cancelled';
  });

  // Check-Ins Section
  sheet.getRange(4, 1).setValue(`⬇️ CHECK-INS TODAY (${checkIns.length})`).setFontWeight('bold').setBackground('#DCFCE7');
  const ciHeaders = ['Booking ID', 'Guest Name', 'Phone', 'Room Type', 'Room No', 'Occupancy', 'Nights', 'Total Paid'];
  sheet.getRange(5, 1, 1, 8).setValues([ciHeaders]).setBackground('#1B2A4A').setFontColor('#FFFFFF').setFontWeight('bold');
  if (checkIns.length > 0) {
    const ciData = checkIns.map(b => [b[0], b[2], b[3], b[5], b[7]||'TBD', b[6], b[10], b[14]]);
    sheet.getRange(6, 1, ciData.length, 8).setValues(ciData).setBackground('#F0FFF4');
  } else {
    sheet.getRange(6, 1).setValue('No check-ins today').setFontColor('#9CA3AF');
  }

  const offset = 7 + Math.max(checkIns.length, 1) + 1;

  // Check-Outs Section
  sheet.getRange(offset, 1).setValue(`⬆️ CHECK-OUTS TODAY (${checkOuts.length})`).setFontWeight('bold').setBackground('#FEF3C7');
  sheet.getRange(offset+1, 1, 1, 8).setValues([ciHeaders]).setBackground('#92400E').setFontColor('#FFFFFF').setFontWeight('bold');
  if (checkOuts.length > 0) {
    const coData = checkOuts.map(b => [b[0], b[2], b[3], b[5], b[7]||'TBD', b[6], b[10], b[14]]);
    sheet.getRange(offset+2, 1, coData.length, 8).setValues(coData).setBackground('#FFFBEB');
  } else {
    sheet.getRange(offset+2, 1).setValue('No check-outs today').setFontColor('#9CA3AF');
  }

  const offset2 = offset + 3 + Math.max(checkOuts.length, 1) + 1;
  sheet.getRange(offset2, 1).setValue(`🏨 STAYING OVER (${stayOvers.length})`).setFontWeight('bold').setBackground('#DBEAFE');
  if (stayOvers.length > 0) {
    sheet.getRange(offset2+1, 1, 1, 8).setValues([ciHeaders]).setBackground('#1E40AF').setFontColor('#FFFFFF').setFontWeight('bold');
    const soData = stayOvers.map(b => [b[0], b[2], b[3], b[5], b[7]||'TBD', b[6], b[10], b[14]]);
    sheet.getRange(offset2+2, 1, soData.length, 8).setValues(soData).setBackground('#EFF6FF');
  }

  sheet.setColumnWidths(1, 8, 140);
}

// ================================================================
//  MENU SETUP (added to Google Sheets toolbar)
// ================================================================
function onOpen() {
  addMenu();
}

function addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🏨 Hotel Sahara')
    .addItem('🔄 Refresh Availability', 'refreshAvailability')
    .addItem('🗓️ Update Today\'s Summary', 'updateTodaySheet')
    .addSeparator()
    .addItem('✅ Save Manual Booking (current row)', 'onBookingManualEntry')
    .addSeparator()
    .addItem('🔧 Re-run Setup', 'setupSheets')
    .addToUi();
}

// ================================================================
//  UTILITIES
// ================================================================
function parseDateStr(dateStr) {
  if (!dateStr) return null;
  const s = String(dateStr).trim();
  if (!s) return null;
  // Handle "dd MMM yyyy" format
  const months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  const parts = s.split(' ');
  if (parts.length === 3) {
    const d = new Date(parseInt(parts[2]), months[parts[1]], parseInt(parts[0]));
    return isNaN(d) ? null : d;
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function generateId() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth()+1).padStart(2,'0');
  const d = String(now.getDate()).padStart(2,'0');
  return `SHR-${y}${m}${d}-${Math.floor(1000+Math.random()*9000)}`;
}

function cap(str) {
  return str ? str.charAt(0).toUpperCase() + str.slice(1).toLowerCase() : '';
}

function num(val) {
  return parseFloat(val || 0).toLocaleString('en-IN', {maximumFractionDigits: 0});
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
