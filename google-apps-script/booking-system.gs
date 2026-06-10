// ============================================================
// HOTEL SAHARA — Google Apps Script Booking System
// ============================================================
// HOW TO USE:
// 1. Go to script.google.com → New Project → paste this entire file
// 2. Run setupSpreadsheet() ONCE to create all sheets
// 3. Deploy as Web App: Deploy → New Deployment → Web App
//    → Execute as: Me | Who has access: Anyone → Deploy
// 4. Copy the Web App URL into js/booking.js (GOOGLE_SCRIPT_URL)
// ============================================================

// ---- CONFIGURATION — Edit these ----
const HOTEL_EMAIL   = 'bookingsahara@rediffmail.com';
const HOTEL_NAME    = 'Hotel Sahara, Pune';
const HOTEL_PHONE   = '020-25655405/6/8/9';
const SHEET_ID      = SpreadsheetApp.getActiveSpreadsheet().getId();

// ---- ROOM DIRECTORY ----
// ⚠️  VERIFY: Confirm which 3 rooms per floor are Deluxe before going live
// Current assumption: highest-numbered active rooms = Deluxe
const ROOM_DIRECTORY = [
  // Floor 1 (101–112, 111 = N/A, 101 = Permanently Booked)
  { no:'101', floor:1, type:'Executive', status:'Permanent', note:'Permanently occupied' },
  { no:'102', floor:1, type:'Executive', status:'Active',    note:'' },
  { no:'103', floor:1, type:'Executive', status:'Active',    note:'' },
  { no:'104', floor:1, type:'Executive', status:'Active',    note:'' },
  { no:'105', floor:1, type:'Executive', status:'Active',    note:'' },
  { no:'106', floor:1, type:'Executive', status:'Active',    note:'' },
  { no:'107', floor:1, type:'Executive', status:'Active',    note:'' },
  { no:'108', floor:1, type:'Executive', status:'Active',    note:'' },
  { no:'109', floor:1, type:'Deluxe',    status:'Active',    note:'⚠️ Confirm if Deluxe' },
  { no:'110', floor:1, type:'Deluxe',    status:'Active',    note:'⚠️ Confirm if Deluxe' },
  { no:'111', floor:1, type:'N/A',       status:'N/A',       note:'Room does not exist' },
  { no:'112', floor:1, type:'Deluxe',    status:'Active',    note:'⚠️ Confirm if Deluxe' },
  // Floor 2 (201–212)
  { no:'201', floor:2, type:'Executive', status:'Active',    note:'' },
  { no:'202', floor:2, type:'Executive', status:'Active',    note:'' },
  { no:'203', floor:2, type:'Executive', status:'Active',    note:'' },
  { no:'204', floor:2, type:'Executive', status:'Active',    note:'' },
  { no:'205', floor:2, type:'Executive', status:'Active',    note:'' },
  { no:'206', floor:2, type:'Executive', status:'Active',    note:'' },
  { no:'207', floor:2, type:'Executive', status:'Active',    note:'' },
  { no:'208', floor:2, type:'Executive', status:'Active',    note:'' },
  { no:'209', floor:2, type:'Executive', status:'Active',    note:'' },
  { no:'210', floor:2, type:'Deluxe',    status:'Active',    note:'⚠️ Confirm if Deluxe' },
  { no:'211', floor:2, type:'Deluxe',    status:'Active',    note:'⚠️ Confirm if Deluxe' },
  { no:'212', floor:2, type:'Deluxe',    status:'Active',    note:'⚠️ Confirm if Deluxe' },
  // Floor 3 (301–312)
  { no:'301', floor:3, type:'Executive', status:'Active',    note:'' },
  { no:'302', floor:3, type:'Executive', status:'Active',    note:'' },
  { no:'303', floor:3, type:'Executive', status:'Active',    note:'' },
  { no:'304', floor:3, type:'Executive', status:'Active',    note:'' },
  { no:'305', floor:3, type:'Executive', status:'Active',    note:'' },
  { no:'306', floor:3, type:'Executive', status:'Active',    note:'' },
  { no:'307', floor:3, type:'Executive', status:'Active',    note:'' },
  { no:'308', floor:3, type:'Executive', status:'Active',    note:'' },
  { no:'309', floor:3, type:'Executive', status:'Active',    note:'' },
  { no:'310', floor:3, type:'Deluxe',    status:'Active',    note:'⚠️ Confirm if Deluxe' },
  { no:'311', floor:3, type:'Deluxe',    status:'Active',    note:'⚠️ Confirm if Deluxe' },
  { no:'312', floor:3, type:'Deluxe',    status:'Active',    note:'⚠️ Confirm if Deluxe' },
];

// ============================================================
// WEB APP ENTRY POINT
// Receives booking data from the website (POST request)
// ============================================================
function doPost(e) {
  try {
    const raw  = e.postData ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    // Save to "All Bookings" sheet
    appendBooking(data, 'Website');

    // Send email to hotel
    sendHotelAlert(data);

    return buildResponse({ status: 'success', message: 'Booking saved' });

  } catch(err) {
    Logger.log('doPost error: ' + err.toString());
    return buildResponse({ status: 'error', message: err.toString() });
  }
}

// Health check (GET)
function doGet(e) {
  return buildResponse({ status: 'ok', message: 'Hotel Sahara Booking System is live' });
}

function buildResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SAVE BOOKING TO "All Bookings" SHEET
// ============================================================
function appendBooking(data, source) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 All Bookings');
  const now   = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MM-yyyy HH:mm');

  const ROOM_NAMES = { executive:'Executive AC', deluxe:'Deluxe AC' };
  const OCC_NAMES  = { single:'Single', double:'Double' };

  sheet.appendRow([
    data.bookingId      || generateId(),
    now,
    source              || 'Manual',
    data.guestName      || '',
    data.guestPhone     || '',
    data.guestEmail     || '',
    '',                                      // Room No — hotel assigns this
    ROOM_NAMES[data.roomType] || data.roomType || '',
    OCC_NAMES[data.occupancy] || data.occupancy || '',
    data.checkIn        || '',
    data.checkOut       || '',
    data.nights         || '',
    data.extraPersons   || 0,
    data.baseRate       || '',
    data.roomTotal      || '',
    data.gst            || '',
    data.totalAmount    || '',
    data.paymentId      || 'Pending',
    source === 'Website' ? 'Razorpay' : 'Cash/Other',
    'Pending Room Assignment',               // Status
    data.requests       || '',               // Special requests
  ]);
}

// ============================================================
// SEND HOTEL EMAIL ALERT
// ============================================================
function sendHotelAlert(data) {
  const ROOM_NAMES = { executive:'Executive AC', deluxe:'Deluxe AC' };
  const OCC_NAMES  = { single:'Single Occupancy', double:'Double Occupancy' };

  const subject = `🏨 New Online Booking: ${data.bookingId} | ${ROOM_NAMES[data.roomType]} | Check-in ${data.checkIn}`;

  const body = `
═══════════════════════════════════
  HOTEL SAHARA — NEW ONLINE BOOKING
═══════════════════════════════════

Booking ID : ${data.bookingId}
Booked On  : ${new Date().toLocaleString('en-IN', {timeZone:'Asia/Kolkata'})}
Source     : Website (Razorpay Payment)

───────────────────────────────────
GUEST DETAILS
───────────────────────────────────
Name    : ${data.guestName}
Phone   : ${data.guestPhone}
Email   : ${data.guestEmail}

───────────────────────────────────
BOOKING DETAILS
───────────────────────────────────
Room Type    : ${ROOM_NAMES[data.roomType] || data.roomType}
Occupancy    : ${OCC_NAMES[data.occupancy] || data.occupancy}
Check-In     : ${data.checkIn}  (1:00 PM)
Check-Out    : ${data.checkOut}  (11:00 AM)
Nights       : ${data.nights}
Extra Persons: ${data.extraPersons || 0}

───────────────────────────────────
PAYMENT
───────────────────────────────────
Rate/Night   : ₹${data.baseRate}
Subtotal     : ₹${data.roomTotal}
GST (12%)    : ₹${data.gst}
TOTAL PAID   : ₹${data.totalAmount}
Payment ID   : ${data.paymentId}
Status       : ✅ PAYMENT CONFIRMED

───────────────────────────────────
${data.requests ? 'SPECIAL REQUESTS:\n' + data.requests + '\n───────────────────────────────────\n' : ''}
⚠️  ACTION REQUIRED: Please assign a room number in the Google Sheet.

View Bookings Sheet → [Your Google Sheet URL]
Admin Dashboard    → https://hotelsaharapune.com/admin.html

${HOTEL_NAME} | ${HOTEL_PHONE}
`;

  MailApp.sendEmail({ to: HOTEL_EMAIL, subject: subject, body: body });
}

// ============================================================
// MANUAL BOOKING (call this from the sheet menu)
// Or hotel staff types directly into the sheet
// ============================================================
function addManualBooking() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.prompt('Manual Booking', 'Enter: GuestName,Phone,RoomNo,CheckIn(dd-mm-yyyy),CheckOut(dd-mm-yyyy)', ui.ButtonSet.OK_CANCEL);

  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const parts = resp.getResponseText().split(',').map(s => s.trim());

  const data = {
    bookingId:  generateId(),
    guestName:  parts[0] || '',
    guestPhone: parts[1] || '',
    roomType:   getRoomType(parts[2]),
    checkIn:    parts[3] || '',
    checkOut:   parts[4] || '',
  };
  appendBooking(data, 'Manual');
  ui.alert('✅ Booking added: ' + data.bookingId);
}

function getRoomType(roomNo) {
  const r = ROOM_DIRECTORY.find(r => r.no === String(roomNo));
  return r ? r.type.toLowerCase() : 'executive';
}

function generateId() {
  const now    = new Date();
  const date   = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyyMMdd');
  const random = Math.floor(1000 + Math.random() * 9000);
  return `SHR-${date}-${random}`;
}

// ============================================================
// SETUP — RUN ONCE to create all sheets
// ============================================================
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setName('Hotel Sahara — Booking System');

  // Clear all existing sheets and rebuild
  createBookingsSheet(ss);
  createRoomDirectorySheet(ss);
  createTodaySheet(ss);
  createCalendarSheet(ss);

  // Remove default "Sheet1" if it exists
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1) ss.deleteSheet(sheet1);

  SpreadsheetApp.getUi().alert('✅ Setup complete! All sheets created.\n\nNext: Deploy as Web App and paste the URL into js/booking.js');
}

// ---- Sheet 1: All Bookings ----
function createBookingsSheet(ss) {
  let sheet = ss.getSheetByName('📋 All Bookings');
  if (!sheet) sheet = ss.insertSheet('📋 All Bookings', 0);
  sheet.clearContents();

  const headers = [
    'Booking ID','Booked On','Source','Guest Name','Phone','Email',
    'Room No','Room Type','Occupancy','Check-In','Check-Out','Nights',
    'Extra Persons','Rate/Night','Room Subtotal','GST','Total Paid',
    'Payment ID','Payment Method','Status','Special Requests'
  ];

  // Header row styling
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#1B2A4A').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10)
    .setHorizontalAlignment('center');

  // Column widths
  const widths = [160,130,80,150,110,180,75,110,80,95,95,60,80,85,95,65,85,160,110,140,200];
  widths.forEach((w, i) => sheet.setColumnWidth(i+1, w));

  // Freeze header
  sheet.setFrozenRows(1);

  // Add example row (manual entry template)
  const exRow = [
    '← Auto-filled','← Auto-filled','Manual','Guest Name','9876543210',
    'guest@email.com','102','Executive AC','Double','12-06-2026','14-06-2026',
    2, 0, 2300, 4600, 552, 5152,'Cash','Cash','Confirmed',''
  ];
  sheet.getRange(2, 1, 1, exRow.length).setValues([exRow])
    .setBackground('#F5F5F5').setFontColor('#999999').setFontStyle('italic');
  sheet.getRange(2,1).setValue('← Example row — delete before use');

  // Status dropdown for column 20
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending Room Assignment','Confirmed','Checked In','Checked Out','Cancelled','No Show'], true)
    .build();
  sheet.getRange(2, 20, 1000, 1).setDataValidation(statusRule);

  Logger.log('✅ All Bookings sheet created');
}

// ---- Sheet 2: Room Directory ----
function createRoomDirectorySheet(ss) {
  let sheet = ss.getSheetByName('🏨 Room Directory');
  if (!sheet) sheet = ss.insertSheet('🏨 Room Directory', 1);
  sheet.clearContents();

  const headers = ['Room No','Floor','Room Type','Status','Tariff (Single)','Tariff (Double)','Notes'];
  sheet.getRange(1,1,1,headers.length).setValues([headers])
    .setBackground('#1B2A4A').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  const RATES = { Executive:[2300,2800], Deluxe:[2500,3000] };

  const rows = ROOM_DIRECTORY.map(r => [
    r.no, r.floor,
    r.status === 'N/A' ? '—' : r.type + ' AC',
    r.status,
    r.status === 'Active' && RATES[r.type] ? '₹'+RATES[r.type][0] : '—',
    r.status === 'Active' && RATES[r.type] ? '₹'+RATES[r.type][1] : '—',
    r.note
  ]);

  sheet.getRange(2, 1, rows.length, 7).setValues(rows);

  // Colour coding
  rows.forEach((row, i) => {
    const bg = row[3] === 'Permanent' ? '#FEF3C7' :
               row[3] === 'N/A'       ? '#E5E7EB' :
               row[1]  == 1           ? '#EFF6FF' :
               row[1]  == 2           ? '#F0FDF4' : '#FFF7F0';
    sheet.getRange(i+2, 1, 1, 7).setBackground(bg);
  });

  // Highlight Deluxe rooms in gold
  rows.forEach((row, i) => {
    if (row[2].includes('Deluxe')) {
      sheet.getRange(i+2, 3).setBackground('#FEF9C3').setFontWeight('bold');
    }
  });

  sheet.setColumnWidths(1, 7, [75, 55, 110, 100, 120, 120, 220]);
  sheet.setFrozenRows(1);
  Logger.log('✅ Room Directory sheet created');
}

// ---- Sheet 3: Today's View ----
function createTodaySheet(ss) {
  let sheet = ss.getSheetByName("📅 Today's View");
  if (!sheet) sheet = ss.insertSheet("📅 Today's View", 2);
  sheet.clearContents();

  sheet.getRange('A1').setValue("📅 TODAY'S HOTEL STATUS")
    .setFontWeight('bold').setFontSize(14).setFontColor('#1B2A4A');
  sheet.getRange('A2').setFormula('=TEXT(TODAY(),"dddd, d mmmm yyyy")').setFontStyle('italic').setFontColor('#6B7280');
  sheet.getRange('A4').setValue('⬇  Refresh by running: refreshTodayView() from Extensions → Apps Script')
    .setFontColor('#9CA3AF').setFontSize(9);

  // This sheet is auto-populated by refreshTodayView()
  refreshTodayView();
  Logger.log("✅ Today's View sheet created");
}

// Called by a daily trigger OR manually from menu
function refreshTodayView() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const sheet   = ss.getSheetByName("📅 Today's View");
  const bookings = ss.getSheetByName('📋 All Bookings');
  if (!sheet || !bookings) return;

  // Clear previous data (below row 5)
  if (sheet.getLastRow() > 5) sheet.getRange(5, 1, sheet.getLastRow()-4, 6).clearContent();

  const today    = new Date();
  const todayStr = Utilities.formatDate(today, 'Asia/Kolkata', 'dd-MM-yyyy');
  const allData  = bookings.getDataRange().getValues().slice(2); // skip header + example

  const arrivals    = [];
  const departures  = [];
  const inHouse     = [];

  allData.forEach(row => {
    const status   = row[19]; // Status column
    const checkIn  = String(row[9]);
    const checkOut = String(row[10]);
    const cancelled = status === 'Cancelled' || status === 'No Show';
    if (cancelled) return;

    if (checkIn === todayStr)  arrivals.push(row);
    if (checkOut === todayStr) departures.push(row);
    if (isInHouseToday(checkIn, checkOut)) inHouse.push(row);
  });

  let r = 5;
  const write = (label, color, rows) => {
    sheet.getRange(r, 1, 1, 6).merge()
      .setValue(label + ' (' + rows.length + ')')
      .setBackground(color).setFontWeight('bold').setFontColor('#FFFFFF').setFontSize(11);
    r++;
    if (rows.length === 0) {
      sheet.getRange(r, 1).setValue('None').setFontColor('#9CA3AF');
      r++;
    } else {
      const subHdr = ['Booking ID','Guest Name','Phone','Room No','Check-In','Check-Out'];
      sheet.getRange(r, 1, 1, 6).setValues([subHdr]).setFontWeight('bold').setBackground('#F3F4F6');
      r++;
      rows.forEach(row => {
        sheet.getRange(r, 1, 1, 6).setValues([[row[0],row[3],row[4],row[6],row[9],row[10]]]);
        r++;
      });
    }
    r++;
  };

  write('🟡 ARRIVALS TODAY', '#D97706', arrivals);
  write('🔴 DEPARTURES TODAY', '#DC2626', departures);
  write('🟢 CURRENTLY IN HOUSE', '#15803D', inHouse);
  sheet.autoResizeColumns(1, 6);
}

function isInHouseToday(checkInStr, checkOutStr) {
  try {
    const parts = s => s.split('-').map(Number);
    const ci = parts(checkInStr);  // [dd,mm,yyyy]
    const co = parts(checkOutStr);
    const today = new Date();
    const checkIn  = new Date(ci[2], ci[1]-1, ci[0]);
    const checkOut = new Date(co[2], co[1]-1, co[0]);
    return checkIn < today && checkOut > today;
  } catch(e) { return false; }
}

// ---- Sheet 4: 60-Day Availability Calendar ----
function createCalendarSheet(ss) {
  let sheet = ss.getSheetByName('📊 60-Day Calendar');
  if (!sheet) sheet = ss.insertSheet('📊 60-Day Calendar', 3);
  sheet.clearContents();

  const ACTIVE_ROOMS = ROOM_DIRECTORY
    .filter(r => r.status === 'Active')
    .map(r => r.no);

  // Header
  sheet.getRange(1,1).setValue('Room ↓ / Date →')
    .setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#FFFFFF');

  // Date headers (next 60 days)
  const today = new Date();
  for (let d = 0; d < 60; d++) {
    const date = new Date(today);
    date.setDate(today.getDate() + d);
    const label = Utilities.formatDate(date, 'Asia/Kolkata', 'dd/MM');
    sheet.getRange(1, d+2).setValue(label)
      .setFontWeight('bold').setBackground('#1B2A4A').setFontColor('#C9A84C')
      .setHorizontalAlignment('center');
    sheet.setColumnWidth(d+2, 46);
  }

  // Room rows
  ACTIVE_ROOMS.forEach((room, i) => {
    const roomInfo = ROOM_DIRECTORY.find(r => r.no === room);
    const bg = roomInfo && roomInfo.type === 'Deluxe' ? '#FEF9C3' : '#FFFFFF';
    sheet.getRange(i+2, 1).setValue('  ' + room + (roomInfo ? '  ' + roomInfo.type[0] : ''))
      .setBackground(bg).setFontWeight('bold');
    // Cells start as Available
    for (let d = 0; d < 60; d++) {
      sheet.getRange(i+2, d+2).setValue('').setBackground('#D1FAE5'); // green = available
    }
  });

  sheet.setColumnWidth(1, 90);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  sheet.getRange(1, 63).setValue('🟢 Available   🔴 Occupied   🟡 Pending')
    .setFontColor('#6B7280').setFontSize(9);

  Logger.log('✅ Calendar sheet created');
}

// ============================================================
// DAILY TRIGGER — set up once
// Run installDailyTrigger() to auto-refresh Today's View at 6AM
// ============================================================
function installDailyTrigger() {
  // Delete existing triggers first
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('refreshTodayView')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .inTimezone('Asia/Kolkata')
    .create();

  SpreadsheetApp.getUi().alert('✅ Daily trigger set. Today\'s View will auto-refresh every morning at 6 AM IST.');
}

// ============================================================
// CUSTOM MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🏨 Hotel Sahara')
    .addItem('🔄 Refresh Today\'s View', 'refreshTodayView')
    .addItem('➕ Add Manual Booking', 'addManualBooking')
    .addItem('⏰ Set Daily Auto-Refresh (6AM)', 'installDailyTrigger')
    .addSeparator()
    .addItem('⚙️  Run Initial Setup', 'setupSpreadsheet')
    .addToUi();
}
