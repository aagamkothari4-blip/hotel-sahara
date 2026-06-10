// ================================================================
//  HOTEL SAHARA — Google Apps Script Backend  v2.0
//  Features: Auto room allocation, room blocking, availability grid
//  Paste into Apps Script editor → run setupSheets() → deploy as Web App
// ================================================================

const CONFIG = {
  HOTEL_EMAIL:    'bookingsahara@rediffmail.com',
  HOTEL_NAME:     'Hotel Sahara, Pune',
  HOTEL_PHONE:    '020-25655405/6/8/9',
  HOTEL_WHATSAPP: '9822393889',
  TIMEZONE:       'Asia/Kolkata',
};

// ── ROOM MASTER ──────────────────────────────────────────────────
// ⚠️  Update Deluxe room numbers to match your actual layout
const ROOMS = [
  {no:101,floor:1,type:'Executive',status:'Permanent',   note:'Permanently occupied'},
  {no:102,floor:1,type:'Executive',status:'Available',   note:''},
  {no:103,floor:1,type:'Executive',status:'Available',   note:''},
  {no:104,floor:1,type:'Deluxe',   status:'Available',   note:''},
  {no:105,floor:1,type:'Executive',status:'Available',   note:''},
  {no:106,floor:1,type:'Executive',status:'Available',   note:''},
  {no:107,floor:1,type:'Executive',status:'Available',   note:''},
  {no:108,floor:1,type:'Deluxe',   status:'Available',   note:''},
  {no:109,floor:1,type:'Executive',status:'Available',   note:''},
  {no:110,floor:1,type:'Executive',status:'Available',   note:''},
  {no:111,floor:1,type:'N/A',      status:'Non-Existent',note:'Does not exist'},
  {no:112,floor:1,type:'Deluxe',   status:'Available',   note:''},
  {no:201,floor:2,type:'Executive',status:'Available',note:''},
  {no:202,floor:2,type:'Executive',status:'Available',note:''},
  {no:203,floor:2,type:'Executive',status:'Available',note:''},
  {no:204,floor:2,type:'Deluxe',   status:'Available',note:''},
  {no:205,floor:2,type:'Executive',status:'Available',note:''},
  {no:206,floor:2,type:'Executive',status:'Available',note:''},
  {no:207,floor:2,type:'Executive',status:'Available',note:''},
  {no:208,floor:2,type:'Deluxe',   status:'Available',note:''},
  {no:209,floor:2,type:'Executive',status:'Available',note:''},
  {no:210,floor:2,type:'Executive',status:'Available',note:''},
  {no:211,floor:2,type:'Executive',status:'Available',note:''},
  {no:212,floor:2,type:'Deluxe',   status:'Available',note:''},
  {no:301,floor:3,type:'Executive',status:'Available',note:''},
  {no:302,floor:3,type:'Executive',status:'Available',note:''},
  {no:303,floor:3,type:'Executive',status:'Available',note:''},
  {no:304,floor:3,type:'Deluxe',   status:'Available',note:''},
  {no:305,floor:3,type:'Executive',status:'Available',note:''},
  {no:306,floor:3,type:'Executive',status:'Available',note:''},
  {no:307,floor:3,type:'Executive',status:'Available',note:''},
  {no:308,floor:3,type:'Deluxe',   status:'Available',note:''},
  {no:309,floor:3,type:'Executive',status:'Available',note:''},
  {no:310,floor:3,type:'Executive',status:'Available',note:''},
  {no:311,floor:3,type:'Executive',status:'Available',note:''},
  {no:312,floor:3,type:'Deluxe',   status:'Available',note:''},
];

// ================================================================
//  HTTP ENDPOINTS
// ================================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const bookingId = data.bookingId || generateId();

    // Auto-allocate a room number
    const roomType  = (data.roomType || '').toLowerCase().includes('deluxe') ? 'Deluxe' : 'Executive';
    const allocated = autoAllocateRoom(ss, roomType, data.checkIn, data.checkOut);

    addToBookingsSheet(ss, data, bookingId, allocated);
    sendHotelEmail(data, bookingId, allocated);
    refreshAvailability(ss);

    return jsonResponse({success:true, bookingId:bookingId, roomAllocated: allocated || 'TBD'});
  } catch(err) {
    return jsonResponse({success:false, error:err.message});
  }
}

function doGet(e) {
  return jsonResponse({status:'Hotel Sahara Booking System v2 — running'});
}

// ================================================================
//  AUTO ROOM ALLOCATION
//  Finds the first available room of the right type for the dates
// ================================================================
function autoAllocateRoom(ss, roomType, checkIn, checkOut) {
  if (!checkIn || !checkOut) return null;

  const ciDate = new Date(checkIn  + 'T00:00:00');
  const coDate = new Date(checkOut + 'T00:00:00');

  // Get all rooms of this type that are bookable
  const candidates = ROOMS.filter(r =>
    r.type === roomType &&
    r.status === 'Available'
  ).map(r => String(r.no));

  // Get already-occupied room numbers for these dates
  const occupied = getOccupiedRooms(ss, ciDate, coDate);

  // Return first candidate not in occupied list
  const available = candidates.filter(r => !occupied.includes(r));
  return available.length > 0 ? available[0] : null;
}

// Returns array of room numbers (as strings) occupied between ciDate and coDate
function getOccupiedRooms(ss, ciDate, coDate) {
  const occupied = [];

  // 1. From confirmed bookings
  const bSheet   = ss.getSheetByName('📋 Bookings');
  const bookings = bSheet.getDataRange().getValues().slice(1);
  bookings.forEach(b => {
    const roomNo = String(b[7]).trim();
    const status = String(b[17]);
    if (!roomNo || roomNo === 'TBD' || roomNo === '' || status === 'Cancelled') return;
    const bCI = parseDateStr(b[8]);
    const bCO = parseDateStr(b[9]);
    if (!bCI || !bCO) return;
    // Overlaps if CI < bCO AND CO > bCI
    if (ciDate < bCO && coDate > bCI) occupied.push(roomNo);
  });

  // 2. From blocked rooms sheet
  const blkSheet = ss.getSheetByName('🚫 Blocked');
  if (blkSheet) {
    const blocks = blkSheet.getDataRange().getValues().slice(1);
    blocks.forEach(b => {
      const roomNo = String(b[0]).trim();
      if (!roomNo || roomNo === '') return;
      const bCI = parseDateStr(b[1]);
      const bCO = parseDateStr(b[2]);
      if (!bCI || !bCO) return;
      if (ciDate < bCO && coDate > bCI) occupied.push(roomNo);
    });
  }

  return occupied;
}

// ================================================================
//  ADD BOOKING TO SHEET
// ================================================================
function addToBookingsSheet(ss, data, bookingId, roomNo) {
  const sheet = ss.getSheetByName('📋 Bookings');
  const fmt   = d => d ? fmtDate(new Date(d + 'T00:00:00')) : '';

  sheet.appendRow([
    bookingId,
    fmtNow(),
    data.guestName    || '',
    data.guestPhone   || '',
    data.guestEmail   || '',
    cap(data.roomType) + ' AC',
    cap(data.occupancy),
    roomNo || 'TBD',
    fmt(data.checkIn),
    fmt(data.checkOut),
    data.nights       || '',
    data.extraPersons || 0,
    '₹' + num(data.roomTotal),
    '₹' + num(data.gst),
    '₹' + num(data.totalAmount),
    data.paymentId    || 'DEMO',
    'Website',
    'Confirmed',
    data.requests     || '',
  ]);

  const row = sheet.getLastRow();
  sheet.getRange(row, 1, 1, 19)
    .setBackground(roomNo ? '#DCFCE7' : '#FEF9C3') // green if allocated, yellow if TBD
    .setBorder(true,true,true,true,false,false,'#BBF7D0',SpreadsheetApp.BorderStyle.SOLID);

  // Bold the allocated room number cell
  if (roomNo) sheet.getRange(row, 8).setFontWeight('bold').setFontColor('#166534');
}

// ================================================================
//  HOTEL NOTIFICATION EMAIL
//  Sent via GmailApp — uses the Google account that owns this script
//  Free, unlimited, no SMTP setup needed
// ================================================================
function sendHotelEmail(data, bookingId, roomNo) {
  const fmt = d => d ? fmtDate(new Date(d + 'T00:00:00'), 'EEE d MMM yyyy') : '—';

  const roomAllocated = roomNo
    ? `<span style="color:#166534;font-weight:bold;font-size:16px">Room ${roomNo} ✅</span>`
    : `<span style="color:#92400E">TBD — no ${cap(data.roomType)} rooms free (please check manually)</span>`;

  const subject = `🏨 New Booking ${bookingId} | ${data.guestName} | ${fmt(data.checkIn)}${roomNo ? ' | Rm '+roomNo : ''}`;

  const html = `
<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
  <div style="background:#1B2A4A;padding:24px;text-align:center">
    <h2 style="color:#C9A84C;margin:0;font-size:22px">🏨 Hotel Sahara — New Booking</h2>
    <p style="color:rgba(255,255,255,0.7);margin:6px 0 0;font-size:13px">Received: ${fmtNow()}</p>
  </div>

  <div style="background:#F8F5EF;padding:24px">
    <!-- Room Allocation Banner -->
    <div style="background:#fff;border-radius:8px;padding:16px 20px;margin-bottom:20px;border-left:5px solid ${roomNo?'#16A34A':'#D97706'}">
      <div style="font-size:12px;color:#6B7280;margin-bottom:4px;text-transform:uppercase;letter-spacing:1px">Auto-Allocated Room</div>
      <div>${roomAllocated}</div>
    </div>

    <table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden">
      <tr><td colspan="2" style="background:#1B2A4A;color:#C9A84C;font-weight:bold;padding:10px 16px;font-size:12px;letter-spacing:1px">BOOKING: ${bookingId}</td></tr>

      <tr><td colspan="2" style="background:#F3F4F6;color:#374151;font-weight:600;padding:8px 16px;font-size:11px;letter-spacing:1px">GUEST DETAILS</td></tr>
      <tr><td style="padding:10px 16px;color:#6B7280;font-size:13px;width:38%">Name</td><td style="padding:10px 16px;font-weight:bold">${data.guestName}</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:10px 16px;color:#6B7280;font-size:13px">Phone</td><td style="padding:10px 16px"><a href="tel:${data.guestPhone}">${data.guestPhone}</a></td></tr>
      <tr><td style="padding:10px 16px;color:#6B7280;font-size:13px">Email</td><td style="padding:10px 16px"><a href="mailto:${data.guestEmail}">${data.guestEmail}</a></td></tr>

      <tr><td colspan="2" style="background:#F3F4F6;color:#374151;font-weight:600;padding:8px 16px;font-size:11px;letter-spacing:1px">STAY DETAILS</td></tr>
      <tr><td style="padding:10px 16px;color:#6B7280;font-size:13px">Room Type</td><td style="padding:10px 16px;font-weight:bold">${cap(data.roomType)} AC — ${cap(data.occupancy)} Occupancy</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:10px 16px;color:#6B7280;font-size:13px">Check-In</td><td style="padding:10px 16px;font-weight:bold">${fmt(data.checkIn)} at 1:00 PM</td></tr>
      <tr><td style="padding:10px 16px;color:#6B7280;font-size:13px">Check-Out</td><td style="padding:10px 16px;font-weight:bold">${fmt(data.checkOut)} at 11:00 AM</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:10px 16px;color:#6B7280;font-size:13px">Nights</td><td style="padding:10px 16px">${data.nights}</td></tr>
      <tr><td style="padding:10px 16px;color:#6B7280;font-size:13px">Extra Persons</td><td style="padding:10px 16px">${data.extraPersons || 0}</td></tr>

      <tr><td colspan="2" style="background:#F3F4F6;color:#374151;font-weight:600;padding:8px 16px;font-size:11px;letter-spacing:1px">PAYMENT</td></tr>
      <tr><td style="padding:10px 16px;color:#6B7280;font-size:13px">Room Charges</td><td style="padding:10px 16px">₹${num(data.roomTotal)}</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:10px 16px;color:#6B7280;font-size:13px">GST</td><td style="padding:10px 16px">₹${num(data.gst)}</td></tr>
      <tr style="background:#DCFCE7"><td style="padding:12px 16px;color:#166534;font-weight:bold">Total Paid</td><td style="padding:12px 16px;font-size:18px;font-weight:bold;color:#166534">₹${num(data.totalAmount)}</td></tr>
      <tr><td style="padding:10px 16px;color:#6B7280;font-size:13px">Payment ID</td><td style="padding:10px 16px;font-family:monospace;font-size:12px">${data.paymentId || 'DEMO'}</td></tr>

      ${data.requests ? `<tr style="background:#FEF3C7"><td colspan="2" style="padding:12px 16px;font-size:13px"><strong>Special Requests:</strong> ${data.requests}</td></tr>` : ''}
    </table>

    <div style="margin-top:16px;text-align:center">
      <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}"
         style="display:inline-block;background:#1B2A4A;color:#C9A84C;padding:12px 28px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:13px">
        Open Booking Sheet →
      </a>
    </div>
  </div>

  <div style="background:#1B2A4A;padding:14px;text-align:center">
    <p style="color:rgba(255,255,255,0.5);font-size:11px;margin:0">
      Hotel Sahara · Senapati Bapat Road, Pune · ${CONFIG.HOTEL_PHONE}
    </p>
  </div>
</div>`;

  GmailApp.sendEmail(CONFIG.HOTEL_EMAIL, subject, '', {
    htmlBody: html,
    name:     'Hotel Sahara Booking System',
    replyTo:  data.guestEmail || CONFIG.HOTEL_EMAIL,
  });
}

// ================================================================
//  REFRESH AVAILABILITY GRID
//  Reads Bookings + Blocked sheets, colours 90-day grid
// ================================================================
function refreshAvailability(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const availSheet = ss.getSheetByName('📅 Availability');
  if (!availSheet) return;

  const today = new Date(); today.setHours(0,0,0,0);
  const bookable = ROOMS.filter(r => r.status !== 'Non-Existent');

  // Build occupied map: roomNo → [array of occupied date strings YYYY-MM-DD]
  const occupiedMap = {};
  bookable.forEach(r => { occupiedMap[String(r.no)] = new Set(); });

  // From bookings
  const bSheet   = ss.getSheetByName('📋 Bookings');
  const bookings = bSheet.getDataRange().getValues().slice(1);
  bookings.forEach(b => {
    const roomNo = String(b[7]).trim();
    const status = String(b[17]);
    if (!roomNo || roomNo === 'TBD' || roomNo === '' || status === 'Cancelled') return;
    const ci = parseDateStr(b[8]);
    const co = parseDateStr(b[9]);
    if (!ci || !co) return;
    const guestFirst = String(b[2]).split(' ')[0] || 'BOOKED';
    markDates(occupiedMap, roomNo, ci, co, guestFirst);
  });

  // From blocked rooms
  const blkSheet = ss.getSheetByName('🚫 Blocked');
  if (blkSheet && blkSheet.getLastRow() > 1) {
    const blocks = blkSheet.getDataRange().getValues().slice(1);
    blocks.forEach(b => {
      const roomNo = String(b[0]).trim();
      if (!roomNo || roomNo === '') return;
      const ci = parseDateStr(b[1]);
      const co = parseDateStr(b[2]);
      const reason = String(b[3] || 'BLOCKED');
      if (!ci || !co) return;
      markDates(occupiedMap, roomNo, ci, co, reason);
    });
  }

  // Write to grid
  bookable.forEach((room, rIdx) => {
    const row   = rIdx + 2;
    const rNo   = String(room.no);
    if (room.status === 'Permanent') return; // Already set as PERMANENT

    for (let d = 0; d < 90; d++) {
      const col  = d + 4;
      const date = new Date(today);
      date.setDate(today.getDate() + d);
      const dateKey = toDateKey(date);
      const cell  = availSheet.getRange(row, col);
      const label = occupiedMap[rNo] ? [...occupiedMap[rNo]]
        .find(entry => entry.startsWith(dateKey + '|')) : null;

      if (label) {
        const guestLabel = label.split('|')[1] || 'BOOKED';
        const isBlocked  = guestLabel === guestLabel.toUpperCase() && guestLabel.length < 12;
        cell.setValue(isBlocked ? '🚫 ' + guestLabel : guestLabel)
            .setBackground(isBlocked ? '#FEF3C7' : '#FEE2E2')
            .setFontColor(isBlocked ? '#92400E'  : '#991B1B')
            .setFontWeight('bold')
            .setFontSize(8);
      } else {
        cell.setValue('')
            .setBackground('#D1FAE5')
            .setFontColor('#166534')
            .setFontWeight('normal')
            .setFontSize(9);
      }
    }
  });
}

function markDates(map, roomNo, ci, co, label) {
  if (!map[roomNo]) return;
  const d = new Date(ci);
  while (d < co) {
    map[roomNo].add(toDateKey(d) + '|' + label);
    d.setDate(d.getDate() + 1);
  }
}

// ================================================================
//  BLOCK A ROOM (manual via menu)
// ================================================================
function blockRoomDialog() {
  const ui   = SpreadsheetApp.getUi();
  const room = ui.prompt('Block a Room', 'Enter room number (e.g. 205):', ui.ButtonSet.OK_CANCEL);
  if (room.getSelectedButton() !== ui.Button.OK) return;

  const fromDate = ui.prompt('From Date', 'Check-in to block (YYYY-MM-DD or "dd Mon yyyy"):', ui.ButtonSet.OK_CANCEL);
  if (fromDate.getSelectedButton() !== ui.Button.OK) return;

  const toDate = ui.prompt('To Date', 'Check-out to unblock (YYYY-MM-DD or "dd Mon yyyy"):', ui.ButtonSet.OK_CANCEL);
  if (toDate.getSelectedButton() !== ui.Button.OK) return;

  const reason = ui.prompt('Reason', 'Reason (e.g. OTA-MakeMyTrip, Walk-in, Maintenance):', ui.ButtonSet.OK_CANCEL);
  if (reason.getSelectedButton() !== ui.Button.OK) return;

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const blk = ss.getSheetByName('🚫 Blocked');

  // Auto-generate a block ID
  const blockId = 'BLK-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd') +
                  '-' + Math.floor(1000 + Math.random()*9000);

  blk.appendRow([
    room.getResponseText().trim(),
    fromDate.getResponseText().trim(),
    toDate.getResponseText().trim(),
    reason.getResponseText().trim(),
    blockId,
    fmtNow(),
    Session.getActiveUser().getEmail(),
  ]);

  // Highlight new row
  blk.getRange(blk.getLastRow(), 1, 1, 7).setBackground('#FEF3C7');

  refreshAvailability(ss);
  ui.alert(`✅ Room ${room.getResponseText()} blocked!\n\nBlock ID: ${blockId}\nAvailability grid updated.`);
}

// ================================================================
//  MANUAL BOOKING SAVE (run from row in Bookings sheet)
// ================================================================
function onBookingManualEntry() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 Bookings');
  const row   = sheet.getActiveCell().getRow();
  if (row < 2) { SpreadsheetApp.getUi().alert('Please click on a booking row first.'); return; }

  const cells = sheet.getRange(row, 1, 1, 19).getValues()[0];

  // Auto-fill missing fields
  if (!cells[0]) sheet.getRange(row, 1).setValue(generateId());
  if (!cells[1]) sheet.getRange(row, 2).setValue(fmtNow());
  if (!cells[16]) sheet.getRange(row, 17).setValue('Manual');
  if (!cells[17]) sheet.getRange(row, 18).setValue('Confirmed');

  // Auto-allocate room if missing
  if (!cells[7] || cells[7] === 'TBD') {
    const roomType   = String(cells[5]).toLowerCase().includes('deluxe') ? 'Deluxe' : 'Executive';
    const checkInStr = String(cells[8]);
    const checkOutStr= String(cells[9]);
    const ci = parseDateStr(checkInStr);
    const co = parseDateStr(checkOutStr);
    if (ci && co) {
      const ciISO = toDateKey(ci);
      const coISO = toDateKey(co);
      const allocated = autoAllocateRoom(ss, roomType, ciISO, coISO);
      if (allocated) {
        sheet.getRange(row, 8).setValue(allocated).setFontWeight('bold').setFontColor('#166534');
        SpreadsheetApp.getUi().alert(`✅ Room ${allocated} auto-allocated!\n\nAvailability grid updated.`);
      } else {
        SpreadsheetApp.getUi().alert(`⚠️ No ${roomType} rooms available for those dates.\nPlease assign a room manually.`);
      }
    }
  }

  sheet.getRange(row, 1, 1, 19).setBackground('#DCFCE7');
  refreshAvailability(ss);
}

// ================================================================
//  ONE-TIME SETUP
// ================================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('🏨 Hotel Sahara — Booking System');

  createBookingsSheet(ss);
  createRoomsSheet(ss);
  createBlockedSheet(ss);
  createAvailabilitySheet(ss);
  createTodaySheet(ss);

  const def = ss.getSheetByName('Sheet1');
  if (def) try { ss.deleteSheet(def); } catch(e) {}

  addMenu();
  SpreadsheetApp.getUi().alert(
    '✅ Setup complete!\n\n' +
    'Tabs created:\n' +
    '📋 Bookings  — all bookings (auto + manual)\n' +
    '🏨 Rooms     — room master list\n' +
    '🚫 Blocked   — rooms blocked from OTAs/walk-ins\n' +
    '📅 Availability — 90-day visual grid\n' +
    '🗓️ Today     — daily check-in/out summary\n\n' +
    'Now deploy as Web App and paste the URL into booking.js'
  );
}

function createBookingsSheet(ss) {
  let s = ss.getSheetByName('📋 Bookings') || ss.insertSheet('📋 Bookings', 0);
  s.clearContents();
  const h = ['Booking ID','Received At','Guest Name','Phone','Email',
             'Room Type','Occupancy','🔑 Room No','Check-In','Check-Out',
             'Nights','Extra','Room Charges','GST','Total Paid',
             'Payment ID','Source','Status','Special Requests'];
  s.getRange(1,1,1,h.length).setValues([h])
   .setBackground('#1B2A4A').setFontColor('#C9A84C').setFontWeight('bold').setFontSize(10);
  s.setFrozenRows(1);
  [165,145,180,130,200,115,95,90,110,110,65,60,115,90,115,175,90,105,250]
    .forEach((w,i) => s.setColumnWidth(i+1, w));

  // Status dropdown
  s.getRange(2,18,500,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Confirmed','Pending','Checked-In','Checked-Out','Cancelled','No-Show'],true).build()
  );
  // Source dropdown
  s.getRange(2,17,500,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Website','Walk-in','Phone','OTA - MakeMyTrip','OTA - OYO','OTA - Booking.com','Other'],true).build()
  );
}

function createRoomsSheet(ss) {
  let s = ss.getSheetByName('🏨 Rooms') || ss.insertSheet('🏨 Rooms', 1);
  s.clearContents();
  s.getRange(1,1,1,7).setValues([['Room No','Floor','Type','Status','Single Tariff','Double Tariff','Notes']])
   .setBackground('#1B2A4A').setFontColor('#C9A84C').setFontWeight('bold');
  s.setFrozenRows(1);
  const T = {Executive:[2300,2800],Deluxe:[2500,3000],'N/A':[0,0]};
  const rows = ROOMS.map(r=>[r.no,r.floor,r.type,r.status,
    r.type!=='N/A'?'₹'+T[r.type][0]:'',r.type!=='N/A'?'₹'+T[r.type][1]:'',r.note]);
  s.getRange(2,1,rows.length,7).setValues(rows);
  ROOMS.forEach((r,i)=>{
    const bg = r.status==='Permanent'?'#FEE2E2':r.status==='Non-Existent'?'#E5E7EB':r.type==='Deluxe'?'#FEF9C3':'#F0F9FF';
    s.getRange(i+2,1,1,7).setBackground(bg);
  });
  s.setColumnWidths(1,7,120); s.setColumnWidth(7,220);
}

function createBlockedSheet(ss) {
  let s = ss.getSheetByName('🚫 Blocked') || ss.insertSheet('🚫 Blocked', 2);
  s.clearContents();
  s.getRange(1,1,1,7).setValues([['Room No','Check-In','Check-Out','Reason / Source','Block ID','Blocked On','Blocked By']])
   .setBackground('#92400E').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  s.setFrozenRows(1);
  [90,110,110,200,165,145,180].forEach((w,i)=>s.setColumnWidth(i+1,w));

  // Add example row so staff understand the format
  s.getRange(2,1,1,7).setValues([
    ['205','01 Jan 2026','03 Jan 2026','OTA - MakeMyTrip','BLK-EXAMPLE','(example — delete this row)','staff@hotel']
  ]).setBackground('#FEF9C3').setFontColor('#92400E').setFontStyle('italic');

  // Reason dropdown
  s.getRange(2,4,500,1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['OTA - MakeMyTrip','OTA - OYO','OTA - Booking.com','Walk-in','Phone Booking','Maintenance','Inspection','Other'],true).build()
  );
}

function createAvailabilitySheet(ss) {
  let s = ss.getSheetByName('📅 Availability') || ss.insertSheet('📅 Availability', 3);
  s.clearContents();

  const today = new Date();
  const headers = ['Room','Type','Floor'];
  for (let i=0;i<90;i++) {
    const d = new Date(today); d.setDate(today.getDate()+i);
    headers.push(Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd MMM'));
  }
  s.getRange(1,1,1,headers.length).setValues([headers])
   .setBackground('#1B2A4A').setFontColor('#C9A84C').setFontWeight('bold').setFontSize(9);
  s.setFrozenRows(1); s.setFrozenColumns(3);
  s.setColumnWidth(1,65); s.setColumnWidth(2,85); s.setColumnWidth(3,60);
  for (let c=4;c<=headers.length;c++) s.setColumnWidth(c,62);

  const bookable = ROOMS.filter(r=>r.status!=='Non-Existent');
  bookable.forEach((r,i)=>{
    const row=i+2;
    s.getRange(row,1).setValue(r.no);
    s.getRange(row,2).setValue(r.type);
    s.getRange(row,3).setValue('F'+r.floor);
    const typeBg = r.type==='Deluxe'?'#FEF9C3':'#F0F9FF';
    s.getRange(row,1,1,3).setBackground(typeBg);
    if (r.status==='Permanent') {
      s.getRange(row,4,1,90).setValue('PERM').setBackground('#FEE2E2').setFontColor('#991B1B').setFontWeight('bold').setFontSize(8);
    } else {
      s.getRange(row,4,1,90).setBackground('#D1FAE5');
    }
  });

  // Legend row
  const leg = bookable.length + 3;
  s.getRange(leg,1).setValue('LEGEND →');
  s.getRange(leg,2).setValue('🟢 Available').setBackground('#D1FAE5');
  s.getRange(leg,3).setValue('🔴 Booked (website)').setBackground('#FEE2E2');
  s.getRange(leg,4).setValue('🟡 Blocked (OTA/walk-in)').setBackground('#FEF3C7');
  s.getRange(leg,5).setValue('🔵 Deluxe room').setBackground('#FEF9C3');
  s.getRange(leg,6).setValue('⚪ Executive room').setBackground('#F0F9FF');
}

function createTodaySheet(ss) {
  let s = ss.getSheetByName('🗓️ Today') || ss.insertSheet('🗓️ Today', 4);
  s.clearContents();
  updateTodaySheet();
}

function updateTodaySheet() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName('🗓️ Today');
  const bSheet = ss.getSheetByName('📋 Bookings');
  if (!sheet||!bSheet) return;
  sheet.clearContents();

  const today    = new Date(); today.setHours(0,0,0,0);
  const todayStr = fmtDate(today, 'EEEE, d MMMM yyyy');
  sheet.getRange(1,1).setValue('🗓️  Daily Summary — '+todayStr).setFontSize(14).setFontWeight('bold').setFontColor('#1B2A4A');
  sheet.getRange(2,1).setValue('Last updated: '+fmtNow()).setFontColor('#9CA3AF').setFontSize(10);

  const rows   = bSheet.getDataRange().getValues().slice(1);
  const cols   = ['Booking ID','Guest Name','Phone','Room Type','Room No','Check-In','Check-Out','Nights','Total Paid'];
  const checkIns  = rows.filter(r=>{ const d=parseDateStr(r[8]); return d&&d.getTime()===today.getTime()&&r[17]!=='Cancelled'; });
  const checkOuts = rows.filter(r=>{ const d=parseDateStr(r[9]); return d&&d.getTime()===today.getTime()&&r[17]!=='Cancelled'; });
  const stayovers = rows.filter(r=>{ const ci=parseDateStr(r[8]),co=parseDateStr(r[9]); return ci&&co&&ci<today&&co>today&&r[17]!=='Cancelled'; });

  const writeSection = (startRow, emoji, label, color, headerBg, data) => {
    sheet.getRange(startRow,1).setValue(`${emoji} ${label} (${data.length})`).setFontWeight('bold').setBackground(color).setFontSize(11);
    sheet.getRange(startRow+1,1,1,cols.length).setValues([cols]).setBackground(headerBg).setFontColor('#fff').setFontWeight('bold');
    if (data.length) {
      const mapped = data.map(r=>[r[0],r[2],r[3],r[5],r[7]||'TBD',r[8],r[9],r[10],r[14]]);
      sheet.getRange(startRow+2,1,mapped.length,cols.length).setValues(mapped);
    } else {
      sheet.getRange(startRow+2,1).setValue('— None today —').setFontColor('#9CA3AF');
    }
    return startRow + 3 + Math.max(data.length,1) + 1;
  };

  let r = 4;
  r = writeSection(r, '⬇️','CHECK-INS TODAY',   '#DCFCE7','#166534', checkIns);
  r = writeSection(r, '⬆️','CHECK-OUTS TODAY',  '#FEF3C7','#92400E', checkOuts);
  r = writeSection(r, '🛏️','STAYING OVER',      '#DBEAFE','#1E40AF', stayovers);
  sheet.setColumnWidths(1,cols.length,150);
}

// ================================================================
//  MENU
// ================================================================
function onOpen() { addMenu(); }

function addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🏨 Hotel Sahara')
    .addItem('🚫 Block a Room',                  'blockRoomDialog')
    .addItem('✅ Save & Auto-Allocate (current row)', 'onBookingManualEntry')
    .addSeparator()
    .addItem('🔄 Refresh Availability Grid',     'refreshAvailability')
    .addItem('🗓️  Update Today\'s Summary',       'updateTodaySheet')
    .addSeparator()
    .addItem('🔧 Re-run Full Setup',             'setupSheets')
    .addToUi();
}

// ================================================================
//  UTILITIES
// ================================================================
function parseDateStr(v) {
  if (!v) return null;
  const s = String(v).trim();
  if (!s) return null;
  const months={Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  const p = s.split(' ');
  if (p.length===3 && months[p[1]]!==undefined) {
    return new Date(parseInt(p[2]), months[p[1]], parseInt(p[0]));
  }
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function toDateKey(d) {
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function fmtDate(d, pattern) {
  if (!d) return '';
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), CONFIG.TIMEZONE, pattern||'dd MMM yyyy');
}

function fmtNow() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd MMM yyyy HH:mm');
}

function generateId() {
  const now=new Date();
  return `SHR-${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}-${Math.floor(1000+Math.random()*9000)}`;
}

function cap(s) { return s ? s.charAt(0).toUpperCase()+s.slice(1).toLowerCase() : ''; }

function num(v) { return parseFloat(v||0).toLocaleString('en-IN',{maximumFractionDigits:0}); }

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
