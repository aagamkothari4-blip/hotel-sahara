// ================================================================
//  HOTEL SAHARA — Google Apps Script Backend  v2.1
//  FIXED: parseDateStr now handles Date objects (Sheets returns
//  date cells as JS Date objects, not strings)
// ================================================================

const CONFIG = {
  HOTEL_EMAIL:    'bookingsahara@rediffmail.com',
  HOTEL_NAME:     'Hotel Sahara, Pune',
  HOTEL_PHONE:    '020-25655405/6/8/9',
  HOTEL_WHATSAPP: '9822393889',
  TIMEZONE:       'Asia/Kolkata',
};

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
    const data       = JSON.parse(e.postData.contents);
    const ss         = SpreadsheetApp.getActiveSpreadsheet();
    const bookingId  = data.bookingId || generateId();
    const roomType   = String(data.roomType||'').toLowerCase().includes('deluxe') ? 'Deluxe' : 'Executive';
    const allocated  = autoAllocateRoom(ss, roomType, data.checkIn, data.checkOut);

    addToBookingsSheet(ss, data, bookingId, allocated);
    sendHotelEmail(data, bookingId, allocated);
    refreshAvailability(ss);

    return jsonResponse({success:true, bookingId:bookingId, roomAllocated: allocated || 'TBD'});
  } catch(err) {
    Logger.log('doPost error: ' + err.message + '\n' + err.stack);
    return jsonResponse({success:false, error:err.message});
  }
}

function doGet(e) {
  return jsonResponse({status:'Hotel Sahara Booking System v2.1 OK', time: new Date().toISOString()});
}

// ================================================================
//  AUTO ROOM ALLOCATION
// ================================================================
function autoAllocateRoom(ss, roomType, checkInStr, checkOutStr) {
  if (!checkInStr || !checkOutStr) return null;

  // checkIn/checkOut come from website as "YYYY-MM-DD"
  const ciDate = safeDate(checkInStr);
  const coDate = safeDate(checkOutStr);
  if (!ciDate || !coDate) {
    Logger.log('autoAllocate: could not parse dates: ' + checkInStr + ' / ' + checkOutStr);
    return null;
  }

  // All bookable rooms of requested type
  const candidates = ROOMS
    .filter(r => r.type === roomType && r.status === 'Available')
    .map(r => String(r.no));

  if (candidates.length === 0) return null;

  // Rooms already occupied in this date range
  const occupied = getOccupiedRooms(ss, ciDate, coDate);
  Logger.log('autoAllocate: candidates=' + candidates + ' occupied=' + occupied);

  const available = candidates.filter(r => !occupied.includes(r));
  return available.length > 0 ? available[0] : null;
}

// Returns array of room number strings occupied between ciDate and coDate
function getOccupiedRooms(ss, ciDate, coDate) {
  const occupied = [];

  // From Bookings sheet
  const bSheet = ss.getSheetByName('📋 Bookings');
  if (bSheet && bSheet.getLastRow() > 1) {
    bSheet.getDataRange().getValues().slice(1).forEach(row => {
      const roomNo = String(row[7]).trim();
      const status = String(row[17]).trim();
      if (!roomNo || roomNo === 'TBD' || roomNo === '' || status === 'Cancelled') return;

      const bCI = safeDate(row[8]);
      const bCO = safeDate(row[9]);
      if (!bCI || !bCO) return;

      // Overlap: newCI < existCO  AND  newCO > existCI
      if (ciDate < bCO && coDate > bCI) {
        occupied.push(roomNo);
      }
    });
  }

  // From Blocked sheet
  const blkSheet = ss.getSheetByName('🚫 Blocked');
  if (blkSheet && blkSheet.getLastRow() > 1) {
    blkSheet.getDataRange().getValues().slice(1).forEach(row => {
      const roomNo = String(row[0]).trim();
      if (!roomNo || roomNo === '') return;
      const bCI = safeDate(row[1]);
      const bCO = safeDate(row[2]);
      if (!bCI || !bCO) return;
      if (ciDate < bCO && coDate > bCI) occupied.push(roomNo);
    });
  }

  return occupied;
}

// ================================================================
//  ADD BOOKING ROW
// ================================================================
function addToBookingsSheet(ss, data, bookingId, roomNo) {
  const sheet = ss.getSheetByName('📋 Bookings');
  const fmtD  = s => { const d = safeDate(s); return d ? fmtDate(d) : String(s||''); };

  sheet.appendRow([
    bookingId,
    fmtNow(),
    data.guestName    || '',
    data.guestPhone   || '',
    data.guestEmail   || '',
    cap(data.roomType) + ' AC',
    cap(data.occupancy),
    roomNo || 'TBD',
    fmtD(data.checkIn),
    fmtD(data.checkOut),
    Number(data.nights)    || '',
    Number(data.extraPersons) || 0,
    '₹' + num(data.roomTotal),
    '₹' + num(data.gst),
    '₹' + num(data.totalAmount),
    data.paymentId    || 'DEMO',
    'Website',
    'Confirmed',
    data.requests     || '',
  ]);

  const lastRow = sheet.getLastRow();
  const rowRange = sheet.getRange(lastRow, 1, 1, 19);
  rowRange.setBackground(roomNo ? '#DCFCE7' : '#FEF9C3');

  if (roomNo) {
    sheet.getRange(lastRow, 8).setFontWeight('bold').setFontColor('#166534');
  } else {
    sheet.getRange(lastRow, 8).setFontColor('#DC2626').setFontWeight('bold');
  }
}

// ================================================================
//  HOTEL EMAIL
// ================================================================
function sendHotelEmail(data, bookingId, roomNo) {
  const fmtD = s => { const d = safeDate(s); return d ? fmtDate(d, 'EEE d MMM yyyy') : String(s||''); };

  const roomBanner = roomNo
    ? `<div style="background:#DCFCE7;border:1px solid #86EFAC;border-radius:6px;padding:14px 18px;margin-bottom:18px">
        <div style="font-size:11px;color:#166534;letter-spacing:1px;margin-bottom:4px">AUTO-ALLOCATED ROOM</div>
        <div style="font-size:22px;font-weight:bold;color:#166534">Room ${roomNo} ✅</div>
       </div>`
    : `<div style="background:#FEF3C7;border:1px solid #FCD34D;border-radius:6px;padding:14px 18px;margin-bottom:18px">
        <div style="font-size:11px;color:#92400E;letter-spacing:1px;margin-bottom:4px">ROOM ALLOCATION</div>
        <div style="font-size:15px;color:#92400E">⚠️ No ${cap(data.roomType)} rooms free — please assign manually</div>
       </div>`;

  const subject = `🏨 New Booking ${bookingId} | ${data.guestName} | ${fmtD(data.checkIn)}${roomNo?' | Rm '+roomNo:''}`;

  const html = `<div style="font-family:Arial,sans-serif;max-width:580px;margin:0 auto">
  <div style="background:#1B2A4A;padding:20px;text-align:center">
    <div style="color:#C9A84C;font-size:20px;font-weight:bold">🏨 Hotel Sahara — New Booking</div>
    <div style="color:rgba(255,255,255,0.6);font-size:12px;margin-top:4px">${fmtNow()}</div>
  </div>
  <div style="padding:20px;background:#F8F5EF">
    ${roomBanner}
    <table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;font-size:13px">
      <tr><td colspan="2" style="background:#1B2A4A;color:#C9A84C;padding:9px 14px;font-weight:bold;letter-spacing:1px;font-size:11px">BOOKING: ${bookingId}</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:9px 14px;color:#6B7280;width:38%">Guest</td><td style="padding:9px 14px;font-weight:bold">${data.guestName}</td></tr>
      <tr><td style="padding:9px 14px;color:#6B7280">Phone</td><td style="padding:9px 14px"><a href="tel:${data.guestPhone}">${data.guestPhone}</a></td></tr>
      <tr style="background:#F9FAFB"><td style="padding:9px 14px;color:#6B7280">Email</td><td style="padding:9px 14px"><a href="mailto:${data.guestEmail}">${data.guestEmail}</a></td></tr>
      <tr><td style="padding:9px 14px;color:#6B7280">Room</td><td style="padding:9px 14px;font-weight:bold">${cap(data.roomType)} AC — ${cap(data.occupancy)}</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:9px 14px;color:#6B7280">Check-In</td><td style="padding:9px 14px;font-weight:bold">${fmtD(data.checkIn)} at 1:00 PM</td></tr>
      <tr><td style="padding:9px 14px;color:#6B7280">Check-Out</td><td style="padding:9px 14px;font-weight:bold">${fmtD(data.checkOut)} at 11:00 AM</td></tr>
      <tr style="background:#F9FAFB"><td style="padding:9px 14px;color:#6B7280">Nights</td><td style="padding:9px 14px">${data.nights}</td></tr>
      <tr style="background:#DCFCE7"><td style="padding:12px 14px;color:#166534;font-weight:bold;font-size:14px">Total Paid</td><td style="padding:12px 14px;font-size:20px;font-weight:bold;color:#166534">₹${num(data.totalAmount)}</td></tr>
      <tr><td style="padding:9px 14px;color:#6B7280">Payment ID</td><td style="padding:9px 14px;font-family:monospace;font-size:11px">${data.paymentId||'DEMO'}</td></tr>
      ${data.requests ? `<tr style="background:#FEF3C7"><td colspan="2" style="padding:10px 14px"><strong>Requests:</strong> ${data.requests}</td></tr>` : ''}
    </table>
    <div style="margin-top:14px;text-align:center">
      <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" style="background:#1B2A4A;color:#C9A84C;padding:11px 24px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:13px;display:inline-block">Open Booking Sheet →</a>
    </div>
  </div>
  <div style="background:#1B2A4A;padding:12px;text-align:center;font-size:11px;color:rgba(255,255,255,0.4)">Hotel Sahara · Senapati Bapat Road, Pune · ${CONFIG.HOTEL_PHONE}</div>
</div>`;

  GmailApp.sendEmail(CONFIG.HOTEL_EMAIL, subject, '', {
    htmlBody: html, name: 'Hotel Sahara Booking System', replyTo: data.guestEmail || CONFIG.HOTEL_EMAIL
  });
}

// ================================================================
//  REFRESH AVAILABILITY GRID  (the key function — fully rewritten)
// ================================================================
function refreshAvailability(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const availSheet = ss.getSheetByName('📅 Availability');
  if (!availSheet) { Logger.log('refreshAvailability: sheet not found'); return; }

  const today = new Date();
  today.setHours(0,0,0,0);

  // Build a flat map: "roomNo|YYYY-MM-DD" → display label
  const occupiedMap = {};

  // --- 1. Confirmed bookings ---
  const bSheet = ss.getSheetByName('📋 Bookings');
  if (bSheet && bSheet.getLastRow() > 1) {
    bSheet.getDataRange().getValues().slice(1).forEach(row => {
      const roomNo = String(row[7]).trim();
      const status = String(row[17]).trim();
      if (!roomNo || roomNo === 'TBD' || roomNo === '' || status === 'Cancelled') return;

      const ci = safeDate(row[8]);
      const co = safeDate(row[9]);
      if (!ci || !co) { Logger.log('Bad dates in row: ' + JSON.stringify(row.slice(0,10))); return; }

      const guestLabel = String(row[2]||'Guest').split(' ')[0];
      markRange(occupiedMap, roomNo, ci, co, guestLabel);
    });
  }

  // --- 2. Blocked rooms ---
  const blkSheet = ss.getSheetByName('🚫 Blocked');
  if (blkSheet && blkSheet.getLastRow() > 1) {
    blkSheet.getDataRange().getValues().slice(1).forEach(row => {
      const roomNo = String(row[0]).trim();
      if (!roomNo || roomNo === '') return;
      const ci = safeDate(row[1]);
      const co = safeDate(row[2]);
      if (!ci || !co) return;
      const reason = String(row[3]||'BLOCKED').replace('OTA - ','').substring(0,10);
      markRange(occupiedMap, roomNo, ci, co, '🚫' + reason);
    });
  }

  Logger.log('occupiedMap entries: ' + Object.keys(occupiedMap).length);

  // --- 3. Write to grid ---
  const bookable = ROOMS.filter(r => r.status !== 'Non-Existent');

  bookable.forEach((room, rIdx) => {
    const row = rIdx + 2;
    const rNo = String(room.no);
    if (room.status === 'Permanent') return;

    const updates = [];
    const bgColors = [];
    const fontColors = [];
    const fontWeights = [];
    const fontSizes = [];

    for (let d = 0; d < 90; d++) {
      const cellDate = new Date(today);
      cellDate.setDate(today.getDate() + d);
      const key = rNo + '|' + dateKey(cellDate);
      const label = occupiedMap[key];

      if (label) {
        const isBlocked = label.startsWith('🚫');
        updates.push([label.substring(0, 10)]);
        bgColors.push([isBlocked ? '#FEF3C7' : '#FEE2E2']);
        fontColors.push([isBlocked ? '#92400E' : '#991B1B']);
        fontWeights.push(['bold']);
        fontSizes.push([8]);
      } else {
        updates.push(['']);
        bgColors.push(['#D1FAE5']);
        fontColors.push(['#166534']);
        fontWeights.push(['normal']);
        fontSizes.push([9]);
      }
    }

    // Batch write for the entire row (much faster than cell-by-cell)
    if (updates.length > 0) {
      const range = availSheet.getRange(row, 4, 1, 90);
      range.setValues([updates.map(u => u[0])]);
      range.setBackgrounds([bgColors.map(b => b[0])]);
      range.setFontColors([fontColors.map(c => c[0])]);
      range.setFontWeights([fontWeights.map(w => w[0])]);
      range.setFontSizes([fontSizes.map(s => s[0])]);
    }
  });

  Logger.log('refreshAvailability: done');
}

// Mark every date from ci (inclusive) to co (exclusive) in occupiedMap
function markRange(map, roomNo, ci, co, label) {
  const d = new Date(ci);
  d.setHours(0,0,0,0);
  const end = new Date(co);
  end.setHours(0,0,0,0);
  while (d < end) {
    map[roomNo + '|' + dateKey(d)] = label;
    d.setDate(d.getDate() + 1);
  }
}

// ================================================================
//  BLOCK A ROOM (manual via menu)
// ================================================================
function blockRoomDialog() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Block Room — Step 1/4', 'Room number to block (e.g. 205):', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const r2 = ui.prompt('Block Room — Step 2/4', 'Check-in date to block from (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const r3 = ui.prompt('Block Room — Step 3/4', 'Check-out date to unblock (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  const r4 = ui.prompt('Block Room — Step 4/4', 'Reason (OTA - MakeMyTrip / Walk-in / Maintenance etc.):', ui.ButtonSet.OK_CANCEL);
  if (r4.getSelectedButton() !== ui.Button.OK) return;

  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const blk = ss.getSheetByName('🚫 Blocked');
  if (!blk) { ui.alert('🚫 Blocked sheet not found. Run "Add Blocked Sheet" from the menu first.'); return; }

  const blockId = 'BLK-' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd') + '-' + Math.floor(1000+Math.random()*9000);
  blk.appendRow([
    r1.getResponseText().trim(),
    r2.getResponseText().trim(),
    r3.getResponseText().trim(),
    r4.getResponseText().trim(),
    blockId, fmtNow(),
    Session.getActiveUser().getEmail()
  ]);
  blk.getRange(blk.getLastRow(),1,1,7).setBackground('#FEF3C7');

  refreshAvailability(ss);
  ui.alert('✅ Room ' + r1.getResponseText() + ' blocked!\nID: ' + blockId + '\n\nAvailability grid updated.');
}

// ================================================================
//  MANUAL BOOKING SAVE (run from Bookings sheet menu)
// ================================================================
function onBookingManualEntry() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 Bookings');
  const row   = sheet.getActiveCell().getRow();
  const ui    = SpreadsheetApp.getUi();

  if (row < 2) { ui.alert('Click on a booking data row first (not the header).'); return; }

  const cells = sheet.getRange(row, 1, 1, 19).getValues()[0];

  if (!cells[0]) sheet.getRange(row,1).setValue(generateId());
  if (!cells[1]) sheet.getRange(row,2).setValue(fmtNow());
  if (!cells[16]) sheet.getRange(row,17).setValue('Manual');
  if (!cells[17]) sheet.getRange(row,18).setValue('Confirmed');

  // Auto-allocate room if blank or TBD
  const assignedRoom = String(cells[7]||'').trim();
  if (!assignedRoom || assignedRoom === 'TBD') {
    const roomType = String(cells[5]||'').toLowerCase().includes('deluxe') ? 'Deluxe' : 'Executive';
    const ci = safeDate(cells[8]);
    const co = safeDate(cells[9]);
    if (ci && co) {
      const ciStr = dateKey(ci);
      const coStr = dateKey(co);
      const allocated = autoAllocateRoom(ss, roomType, ciStr, coStr);
      if (allocated) {
        sheet.getRange(row,8).setValue(allocated).setFontWeight('bold').setFontColor('#166534');
        refreshAvailability(ss);
        ui.alert('✅ Room ' + allocated + ' assigned!\nAvailability updated.');
        return;
      } else {
        ui.alert('⚠️ No ' + roomType + ' rooms available for those dates.\nPlease assign a room number manually in column H.');
      }
    }
  }
  sheet.getRange(row,1,1,19).setBackground('#DCFCE7');
  refreshAvailability(ss);
}

// ================================================================
//  SAFE FIRST-TIME SETUP  (won't wipe existing data)
// ================================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('🏨 Hotel Sahara — Booking System');

  // Create sheets only if they don't exist, never clear existing data
  ensureBookingsSheet(ss);
  ensureRoomsSheet(ss);
  ensureBlockedSheet(ss);
  ensureAvailabilitySheet(ss);
  ensureTodaySheet(ss);

  try { const d = ss.getSheetByName('Sheet1'); if(d) ss.deleteSheet(d); } catch(e) {}
  addMenu();

  SpreadsheetApp.getUi().alert('✅ Setup complete!\n\n' +
    '📋 Bookings — all bookings\n' +
    '🏨 Rooms    — room master\n' +
    '🚫 Blocked  — OTA/walk-in blocks\n' +
    '📅 Availability — 90-day grid\n' +
    '🗓️ Today   — daily summary\n\n' +
    'Existing data was NOT deleted.');
}

// Run this if you only need to add the Blocked sheet to an existing setup
function addBlockedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureBlockedSheet(ss);
  SpreadsheetApp.getUi().alert('✅ 🚫 Blocked sheet added (or already existed).');
}

function ensureBookingsSheet(ss) {
  let s = ss.getSheetByName('📋 Bookings');
  const isNew = !s;
  if (isNew) s = ss.insertSheet('📋 Bookings', 0);
  // Only write header if sheet is truly empty
  if (s.getLastRow() === 0) {
    const h = ['Booking ID','Received At','Guest Name','Phone','Email','Room Type','Occupancy',
               '🔑 Room No','Check-In','Check-Out','Nights','Extra','Room Charges','GST',
               'Total Paid','Payment ID','Source','Status','Special Requests'];
    s.getRange(1,1,1,h.length).setValues([h])
     .setBackground('#1B2A4A').setFontColor('#C9A84C').setFontWeight('bold').setFontSize(10);
    s.setFrozenRows(1);
  }
  // Always ensure dropdowns (safe to re-apply)
  s.getRange(2,18,500,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(
      ['Confirmed','Pending','Checked-In','Checked-Out','Cancelled','No-Show'],true).build());
  s.getRange(2,17,500,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(
      ['Website','Walk-in','Phone','OTA - MakeMyTrip','OTA - OYO','OTA - Booking.com','Other'],true).build());
}

function ensureRoomsSheet(ss) {
  if (ss.getSheetByName('🏨 Rooms')) return; // Already exists — don't touch
  const s = ss.insertSheet('🏨 Rooms', 1);
  s.getRange(1,1,1,7).setValues([['Room No','Floor','Type','Status','Single Tariff','Double Tariff','Notes']])
   .setBackground('#1B2A4A').setFontColor('#C9A84C').setFontWeight('bold');
  s.setFrozenRows(1);
  const T = {Executive:[2300,2800],Deluxe:[2500,3000],'N/A':[0,0]};
  const rows = ROOMS.map(r=>[r.no,r.floor,r.type,r.status,
    r.type!=='N/A'?'₹'+T[r.type][0]:'',r.type!=='N/A'?'₹'+T[r.type][1]:'',r.note]);
  s.getRange(2,1,rows.length,7).setValues(rows);
  ROOMS.forEach((r,i)=>{
    const bg=r.status==='Permanent'?'#FEE2E2':r.status==='Non-Existent'?'#E5E7EB':r.type==='Deluxe'?'#FEF9C3':'#F0F9FF';
    s.getRange(i+2,1,1,7).setBackground(bg);
  });
}

function ensureBlockedSheet(ss) {
  if (ss.getSheetByName('🚫 Blocked')) return; // Already exists
  const s = ss.insertSheet('🚫 Blocked');
  s.getRange(1,1,1,7).setValues([['Room No','Check-In','Check-Out','Reason / Source','Block ID','Blocked On','Blocked By']])
   .setBackground('#92400E').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  s.setFrozenRows(1);
  [90,110,110,200,165,145,180].forEach((w,i)=>s.setColumnWidth(i+1,w));
  s.getRange(2,4,500,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(
      ['OTA - MakeMyTrip','OTA - OYO','OTA - Booking.com','Walk-in','Phone Booking','Maintenance','Other'],true).build());
  // Example row
  s.getRange(2,1,1,7).setValues([['205','2026-01-01','2026-01-03','OTA - MakeMyTrip',
    'BLK-EXAMPLE',fmtNow(),'example — delete this row']])
   .setBackground('#FEF9C3').setFontColor('#92400E').setFontStyle('italic');
}

function ensureAvailabilitySheet(ss) {
  let s = ss.getSheetByName('📅 Availability');
  if (!s) s = ss.insertSheet('📅 Availability');
  s.clearContents(); // Safe to clear — this sheet is auto-generated, not user data

  const today = new Date();
  const headers = ['Room','Type','Floor'];
  for (let i=0;i<90;i++) {
    const d = new Date(today); d.setDate(today.getDate()+i);
    headers.push(Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd MMM'));
  }
  s.getRange(1,1,1,headers.length).setValues([headers])
   .setBackground('#1B2A4A').setFontColor('#C9A84C').setFontWeight('bold').setFontSize(9);
  s.setFrozenRows(1); s.setFrozenColumns(3);
  s.setColumnWidth(1,65); s.setColumnWidth(2,80); s.setColumnWidth(3,52);
  for (let c=4;c<=headers.length;c++) s.setColumnWidth(c,60);

  const bookable = ROOMS.filter(r=>r.status!=='Non-Existent');
  bookable.forEach((r,i)=>{
    const row=i+2;
    s.getRange(row,1).setValue(r.no);
    s.getRange(row,2).setValue(r.type);
    s.getRange(row,3).setValue('F'+r.floor);
    s.getRange(row,1,1,3).setBackground(r.type==='Deluxe'?'#FEF9C3':'#F0F9FF');
    if (r.status==='Permanent') {
      s.getRange(row,4,1,90).setValue('PERM').setBackground('#FEE2E2').setFontColor('#991B1B').setFontWeight('bold').setFontSize(8);
    } else {
      s.getRange(row,4,1,90).setBackground('#D1FAE5');
    }
  });
  // Legend
  const leg = bookable.length + 3;
  s.getRange(leg,1).setValue('LEGEND →');
  [['🟢 Available','#D1FAE5'],['🔴 Guest name (booked)','#FEE2E2'],['🟡 🚫 Blocked/OTA','#FEF3C7'],['🟡 Deluxe room','#FEF9C3'],['🔵 Executive room','#F0F9FF']]
    .forEach(([t,bg],j) => s.getRange(leg,j+2).setValue(t).setBackground(bg));
}

function ensureTodaySheet(ss) {
  let s = ss.getSheetByName('🗓️ Today');
  if (!s) s = ss.insertSheet('🗓️ Today');
  updateTodaySheet();
}

// ================================================================
//  TODAY SHEET
// ================================================================
function updateTodaySheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('🗓️ Today');
  const bSheet= ss.getSheetByName('📋 Bookings');
  if (!sheet||!bSheet) return;
  sheet.clearContents();

  const today    = new Date(); today.setHours(0,0,0,0);
  const todayStr = Utilities.formatDate(today, CONFIG.TIMEZONE, 'EEEE, d MMMM yyyy');
  sheet.getRange(1,1).setValue('🗓️  Daily Summary — ' + todayStr).setFontSize(14).setFontWeight('bold');
  sheet.getRange(2,1).setValue('Updated: ' + fmtNow()).setFontColor('#9CA3AF').setFontSize(10);

  const rows = bSheet.getLastRow() > 1 ? bSheet.getDataRange().getValues().slice(1) : [];
  const checkIns  = rows.filter(r => { const d=safeDate(r[8]); return d&&dateKey(d)===dateKey(today)&&r[17]!=='Cancelled'; });
  const checkOuts = rows.filter(r => { const d=safeDate(r[9]); return d&&dateKey(d)===dateKey(today)&&r[17]!=='Cancelled'; });
  const stayovers = rows.filter(r => {
    const ci=safeDate(r[8]),co=safeDate(r[9]);
    return ci&&co&&ci<today&&co>today&&r[17]!=='Cancelled';
  });

  const cols = ['Booking ID','Guest Name','Phone','Room Type','Room No','Check-In','Check-Out','Nights','Total Paid'];
  const mapRow = r => [r[0],r[2],r[3],r[5],r[7]||'TBD',String(r[8]||''),String(r[9]||''),r[10],r[14]];

  let startRow = 4;
  const sections = [
    {label:'⬇  CHECK-INS TODAY',  bg:'#DCFCE7',hBg:'#166534', data:checkIns},
    {label:'⬆  CHECK-OUTS TODAY', bg:'#FEF3C7',hBg:'#92400E', data:checkOuts},
    {label:'🛏  STAYING OVER',     bg:'#DBEAFE',hBg:'#1E40AF', data:stayovers},
  ];

  sections.forEach(sec => {
    sheet.getRange(startRow,1).setValue(sec.label + ' (' + sec.data.length + ')')
      .setFontWeight('bold').setBackground(sec.bg).setFontSize(11);
    sheet.getRange(startRow+1,1,1,cols.length).setValues([cols]).setBackground(sec.hBg).setFontColor('#FFFFFF').setFontWeight('bold');
    if (sec.data.length > 0) {
      sheet.getRange(startRow+2,1,sec.data.length,cols.length).setValues(sec.data.map(mapRow));
    } else {
      sheet.getRange(startRow+2,1).setValue('— None today —').setFontColor('#9CA3AF');
    }
    startRow += 3 + Math.max(sec.data.length, 1) + 1;
  });

  sheet.setColumnWidths(1, cols.length, 145);
}

// ================================================================
//  MENU
// ================================================================
function onOpen() { addMenu(); }

function addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🏨 Hotel Sahara')
    .addItem('🚫 Block a Room',                       'blockRoomDialog')
    .addItem('✅ Save & Auto-Allocate (current row)', 'onBookingManualEntry')
    .addSeparator()
    .addItem('🔄 Refresh Availability Grid',          'refreshAvailability')
    .addItem('🗓️  Update Today\'s Summary',            'updateTodaySheet')
    .addSeparator()
    .addItem('➕ Add Blocked Sheet (if missing)',      'addBlockedSheet')
    .addItem('🔧 Full Setup (safe — keeps data)',     'setupSheets')
    .addToUi();
}

// ================================================================
//  *** CORE FIX: safeDate handles BOTH strings AND Date objects ***
//  Google Sheets returns date cells as JS Date objects, not strings.
//  parseDateStr v1 only handled strings → broke allocation & grid.
// ================================================================
function safeDate(v) {
  if (!v) return null;
  // Already a JS Date (how Sheets returns date-formatted cells)
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    const d = new Date(v.getFullYear(), v.getMonth(), v.getDate());
    return d;
  }
  const s = String(v).trim();
  if (!s || s === 'undefined' || s === 'null') return null;

  // YYYY-MM-DD  (from website)
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return new Date(s + 'T00:00:00');
  }
  // dd MMM yyyy  e.g. "15 Jun 2026"
  const months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  const p = s.split(' ');
  if (p.length === 3 && months[p[1]] !== undefined) {
    return new Date(parseInt(p[2]), months[p[1]], parseInt(p[0]));
  }
  // Fallback — try native parse
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// Returns "YYYY-MM-DD" string for a Date
function dateKey(d) {
  return Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

// ================================================================
//  UTILITIES
// ================================================================
function fmtDate(d, pattern) {
  if (!d) return '';
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), CONFIG.TIMEZONE, pattern||'dd MMM yyyy');
}
function fmtNow() { return fmtDate(new Date(), 'dd MMM yyyy HH:mm'); }
function generateId() {
  const n=new Date();
  return 'SHR-'+n.getFullYear()+String(n.getMonth()+1).padStart(2,'0')+String(n.getDate()).padStart(2,'0')+'-'+Math.floor(1000+Math.random()*9000);
}
function cap(s) { return s?s.charAt(0).toUpperCase()+s.slice(1).toLowerCase():''; }
function num(v) { return parseFloat(v||0).toLocaleString('en-IN',{maximumFractionDigits:0}); }
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
