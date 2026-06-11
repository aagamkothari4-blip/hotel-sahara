// ================================================================
//  HOTEL SAHARA — Google Apps Script Backend  v3.0
//  NEW: Block by selecting cells on Availability grid (multi-room)
//       Multi-room allocation dialog
//       Corrected Deluxe room list
// ================================================================

const CONFIG = {
  HOTEL_EMAIL:    'bookingsahara@rediffmail.com',
  HOTEL_NAME:     'Hotel Sahara, Pune',
  HOTEL_PHONE:    '020-25655405/6/8/9',
  TIMEZONE:       'Asia/Kolkata',
};

// ================================================================
//  ROOM MASTER  — updated with correct Deluxe rooms
//  Deluxe: 105,106,112 | 201,205,206,212 | 301,305,306,312
// ================================================================
const DELUXE = new Set(['105','106','112','201','205','206','212','301','305','306','312']);

const ROOMS = [
  // Floor 1
  {no:101,floor:1,type:'Executive',status:'Permanent',   note:'Permanently occupied'},
  {no:102,floor:1,type:'Executive',status:'Available',   note:''},
  {no:103,floor:1,type:'Executive',status:'Available',   note:''},
  {no:104,floor:1,type:'Executive',status:'Available',   note:''},
  {no:105,floor:1,type:'Deluxe',   status:'Available',   note:''},
  {no:106,floor:1,type:'Deluxe',   status:'Available',   note:''},
  {no:107,floor:1,type:'Executive',status:'Available',   note:''},
  {no:108,floor:1,type:'Executive',status:'Available',   note:''},
  {no:109,floor:1,type:'Executive',status:'Available',   note:''},
  {no:110,floor:1,type:'Executive',status:'Available',   note:''},
  {no:111,floor:1,type:'N/A',      status:'Non-Existent',note:'Does not exist'},
  {no:112,floor:1,type:'Deluxe',   status:'Available',   note:''},
  // Floor 2
  {no:201,floor:2,type:'Deluxe',   status:'Available',note:''},
  {no:202,floor:2,type:'Executive',status:'Available',note:''},
  {no:203,floor:2,type:'Executive',status:'Available',note:''},
  {no:204,floor:2,type:'Executive',status:'Available',note:''},
  {no:205,floor:2,type:'Deluxe',   status:'Available',note:''},
  {no:206,floor:2,type:'Deluxe',   status:'Available',note:''},
  {no:207,floor:2,type:'Executive',status:'Available',note:''},
  {no:208,floor:2,type:'Executive',status:'Available',note:''},
  {no:209,floor:2,type:'Executive',status:'Available',note:''},
  {no:210,floor:2,type:'Executive',status:'Available',note:''},
  {no:211,floor:2,type:'Executive',status:'Available',note:''},
  {no:212,floor:2,type:'Deluxe',   status:'Available',note:''},
  // Floor 3
  {no:301,floor:3,type:'Deluxe',   status:'Available',note:''},
  {no:302,floor:3,type:'Executive',status:'Available',note:''},
  {no:303,floor:3,type:'Executive',status:'Available',note:''},
  {no:304,floor:3,type:'Executive',status:'Available',note:''},
  {no:305,floor:3,type:'Deluxe',   status:'Available',note:''},
  {no:306,floor:3,type:'Deluxe',   status:'Available',note:''},
  {no:307,floor:3,type:'Executive',status:'Available',note:''},
  {no:308,floor:3,type:'Executive',status:'Available',note:''},
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
    const data      = JSON.parse(e.postData.contents);
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const bookingId = data.bookingId || generateId();
    const roomType  = String(data.roomType||'').toLowerCase().includes('deluxe') ? 'Deluxe' : 'Executive';
    const allocated = autoAllocateRoom(ss, roomType, data.checkIn, data.checkOut);

    addToBookingsSheet(ss, data, bookingId, allocated);
    sendHotelEmail(data, bookingId, allocated);
    refreshAvailability(ss);

    return jsonResponse({success:true, bookingId:bookingId, roomAllocated: allocated||'TBD'});
  } catch(err) {
    Logger.log('doPost error: '+err.message+'\n'+err.stack);
    return jsonResponse({success:false, error:err.message});
  }
}

function doGet(e) {
  return jsonResponse({status:'Hotel Sahara v3.0 OK', time:new Date().toISOString()});
}

// ================================================================
//  AUTO ALLOCATE — single room
// ================================================================
function autoAllocateRoom(ss, roomType, checkInStr, checkOutStr) {
  if (!checkInStr || !checkOutStr) return null;
  const ci = safeDate(checkInStr);
  const co = safeDate(checkOutStr);
  if (!ci || !co) return null;

  const candidates = ROOMS.filter(r => r.type === roomType && r.status === 'Available').map(r => String(r.no));
  const occupied   = getOccupiedRooms(ss, ci, co);
  const available  = candidates.filter(r => !occupied.includes(r));
  return available.length > 0 ? available[0] : null;
}

// ================================================================
//  ALLOCATE MULTIPLE ROOMS (menu dialog)
//  For group bookings — returns a list of available rooms
// ================================================================
function allocateMultipleRooms() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const r1 = ui.prompt('Allocate Rooms — 1/4',
    'Room type?\n  D = Deluxe AC\n  E = Executive AC', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;

  const r2 = ui.prompt('Allocate Rooms — 2/4',
    'How many rooms needed?', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;

  const r3 = ui.prompt('Allocate Rooms — 3/4',
    'Check-in date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;

  const r4 = ui.prompt('Allocate Rooms — 4/4',
    'Check-out date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (r4.getSelectedButton() !== ui.Button.OK) return;

  const typeIn   = r1.getResponseText().trim().toLowerCase();
  const roomType = (typeIn === 'd' || typeIn.includes('deluxe')) ? 'Deluxe' : 'Executive';
  const count    = Math.max(1, parseInt(r2.getResponseText()) || 1);
  const ci       = safeDate(r3.getResponseText().trim());
  const co       = safeDate(r4.getResponseText().trim());

  if (!ci || !co) { ui.alert('Invalid dates. Use YYYY-MM-DD (e.g. 2026-07-15).'); return; }
  if (co <= ci)   { ui.alert('Check-out must be after check-in.'); return; }

  const occupied  = getOccupiedRooms(ss, ci, co);
  const candidates= ROOMS.filter(r => r.type === roomType && r.status === 'Available').map(r => String(r.no));
  const available = candidates.filter(r => !occupied.includes(r));

  if (available.length === 0) {
    ui.alert('❌ No ' + roomType + ' rooms available\n' + fmtDate(ci) + ' → ' + fmtDate(co));
    return;
  }

  const toAllocate = available.slice(0, Math.min(count, available.length));
  const shortfall  = count - toAllocate.length;

  let msg = '✅ ' + roomType + ' AC — ' + toAllocate.length + ' room(s) allocated\n\n'
    + 'Check-in:  ' + fmtDate(ci) + '\n'
    + 'Check-out: ' + fmtDate(co) + '\n\n'
    + 'Rooms assigned:  ' + toAllocate.join(', ') + '\n';

  if (shortfall > 0) {
    msg += '\n⚠️ Could only find ' + toAllocate.length + ' (needed ' + count + ').\n'
         + 'Check if other ' + roomType + ' rooms can be freed up.';
  }

  msg += '\n\nPlease enter booking rows in 📋 Bookings with these room numbers, or type them in column H.';

  ui.alert(msg);

  // Highlight the allocated rooms in the availability sheet for the date range
  highlightAllocatedRooms(ss, toAllocate, ci, co);
}

function highlightAllocatedRooms(ss, roomNos, ci, co) {
  const availSheet = ss.getSheetByName('📅 Availability');
  if (!availSheet) return;
  const today = new Date(); today.setHours(0,0,0,0);
  const bookable = ROOMS.filter(r => r.status !== 'Non-Existent');

  roomNos.forEach(rNo => {
    const rIdx = bookable.findIndex(r => String(r.no) === rNo);
    if (rIdx < 0) return;
    const row = rIdx + 2;
    for (let d = 0; d < 90; d++) {
      const cellDate = new Date(today); cellDate.setDate(today.getDate() + d);
      if (cellDate >= ci && cellDate < co) {
        availSheet.getRange(row, d + 4)
          .setBackground('#DBEAFE').setFontColor('#1E40AF')
          .setValue('HOLD').setFontWeight('bold').setFontSize(8);
      }
    }
  });
}

// ================================================================
//  BLOCK SELECTED CELLS ON AVAILABILITY SHEET
//  1. Switch to 📅 Availability tab
//  2. Select any cells across rooms AND dates
//  3. Click menu → Block Selected Cells
// ================================================================
function blockSelectedCells() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();

  if (sheet.getName() !== '📅 Availability') {
    ui.alert('⚠️ Please switch to the 📅 Availability sheet,\nselect the cells you want to block, then run this again.');
    return;
  }

  const sel      = sheet.getActiveRange();
  const startRow = sel.getRow();
  const endRow   = sel.getLastRow();
  const startCol = sel.getColumn();
  const endCol   = sel.getLastColumn();

  if (startRow < 2) { ui.alert('Please do not include the header row in your selection.'); return; }
  if (startCol < 4) { ui.alert('Please select only date columns (from column 4 onwards).'); return; }

  // Collect room numbers from column A of selected rows
  const roomNums = [];
  for (let r = startRow; r <= endRow; r++) {
    const rNo = String(sheet.getRange(r, 1).getValue()).trim();
    if (rNo && rNo !== '' && rNo !== 'Room' && rNo !== 'LEGEND →') roomNums.push(rNo);
  }
  if (roomNums.length === 0) { ui.alert('No valid room numbers in selection.'); return; }

  // Get dates from header row
  const headerVals = sheet.getRange(1, startCol, 1, endCol - startCol + 1).getValues()[0];
  const today      = new Date(); today.setHours(0,0,0,0);

  const parsedDates = headerVals
    .map(h => parseDdMon(String(h).trim(), today))
    .filter(d => d !== null);

  if (parsedDates.length === 0) { ui.alert('Could not read dates from column headers.'); return; }

  const ciDate = parsedDates[0];
  const coDate = new Date(parsedDates[parsedDates.length - 1]);
  coDate.setDate(coDate.getDate() + 1); // checkout = day AFTER last night

  // Prompt for reason
  const resp = ui.prompt(
    '🚫 Block ' + roomNums.length + ' Room(s)',
    'Rooms: ' + roomNums.join(', ') + '\n'
      + 'From:  ' + fmtDate(ciDate) + '\n'
      + 'To:    ' + fmtDate(coDate) + ' (check-out)\n\n'
      + 'Reason / Source:',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const blkSheet = ss.getSheetByName('🚫 Blocked');
  if (!blkSheet) {
    ui.alert('🚫 Blocked sheet not found. Run "Add Blocked Sheet" from the menu first.');
    return;
  }

  const ciStr  = dateKey(ciDate);
  const coStr  = dateKey(coDate);
  const reason = resp.getResponseText().trim() || 'Blocked';

  roomNums.forEach(rNo => {
    const blockId = 'BLK-' + dateKey(today).replace(/-/g,'') + '-' + Math.floor(1000+Math.random()*9000);
    blkSheet.appendRow([rNo, ciStr, coStr, reason, blockId, fmtNow(), Session.getActiveUser().getEmail()]);
    blkSheet.getRange(blkSheet.getLastRow(), 1, 1, 7).setBackground('#FEF3C7');
  });

  refreshAvailability(ss);
  ui.alert('✅ Blocked ' + roomNums.length + ' room(s): ' + roomNums.join(', ')
    + '\n' + fmtDate(ciDate) + ' → ' + fmtDate(coDate)
    + '\n\nAvailability grid updated.');
}

// Parse "dd MMM" header string → Date (year chosen so it's not far in the past)
function parseDdMon(s, today) {
  const months = {Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  const p = s.split(' ');
  if (p.length !== 2 || months[p[1]] === undefined) return null;
  const day = parseInt(p[0]);
  const mon = months[p[1]];
  let d = new Date(today.getFullYear(), mon, day);
  // If the date is more than 7 days in the past, it must belong to next year
  const threshold = new Date(today); threshold.setDate(today.getDate() - 7);
  if (d < threshold) d = new Date(today.getFullYear() + 1, mon, day);
  return d;
}

// ================================================================
//  UNBLOCK — delete a block entry from the Blocked sheet
// ================================================================
function unblockRoomDialog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const blk = ss.getSheetByName('🚫 Blocked');
  if (!blk || blk.getLastRow() < 2) { ui.alert('No blocks found.'); return; }

  const r1 = ui.prompt('Unblock Room',
    'Enter the Block ID to remove (e.g. BLK-20260615-1234).\n\n'
    + 'You can find it in the 🚫 Blocked sheet column E.',
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;

  const blockId = r1.getResponseText().trim();
  const data    = blk.getDataRange().getValues();

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][4]).trim() === blockId) {
      blk.deleteRow(i + 1);
      refreshAvailability(ss);
      ui.alert('✅ Block ' + blockId + ' removed.\nAvailability grid updated.');
      return;
    }
  }
  ui.alert('Block ID not found: ' + blockId);
}

// ================================================================
//  MANUAL BOOKING SAVE (menu → current row)
// ================================================================
function onBookingManualEntry() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('📋 Bookings');
  const row   = sheet.getActiveCell().getRow();
  const ui    = SpreadsheetApp.getUi();
  if (row < 2) { ui.alert('Click on a booking data row first.'); return; }

  const cells = sheet.getRange(row, 1, 1, 19).getValues()[0];
  if (!cells[0]) sheet.getRange(row,1).setValue(generateId());
  if (!cells[1]) sheet.getRange(row,2).setValue(fmtNow());
  if (!cells[16]) sheet.getRange(row,17).setValue('Manual');
  if (!cells[17]) sheet.getRange(row,18).setValue('Confirmed');

  const assignedRoom = String(cells[7]||'').trim();
  if (!assignedRoom || assignedRoom === 'TBD') {
    const roomType = String(cells[5]||'').toLowerCase().includes('deluxe') ? 'Deluxe' : 'Executive';
    const ci = safeDate(cells[8]);
    const co = safeDate(cells[9]);
    if (ci && co) {
      const alloc = autoAllocateRoom(ss, roomType, dateKey(ci), dateKey(co));
      if (alloc) {
        sheet.getRange(row,8).setValue(alloc).setFontWeight('bold').setFontColor('#166534');
        sheet.getRange(row,1,1,19).setBackground('#DCFCE7');
        refreshAvailability(ss);
        ui.alert('✅ Room ' + alloc + ' assigned!');
        return;
      } else {
        ui.alert('⚠️ No ' + roomType + ' rooms free for those dates.\nPlease assign a room number manually in column H.');
      }
    }
  }
  sheet.getRange(row,1,1,19).setBackground('#DCFCE7');
  refreshAvailability(ss);
}

// ================================================================
//  ADD BOOKING TO BOOKINGS SHEET
// ================================================================
function addToBookingsSheet(ss, data, bookingId, roomNo) {
  const sheet = ss.getSheetByName('📋 Bookings');
  const fmtD  = s => { const d=safeDate(s); return d?fmtDate(d):String(s||''); };

  sheet.appendRow([
    bookingId, fmtNow(),
    data.guestName||'', data.guestPhone||'', data.guestEmail||'',
    cap(data.roomType)+' AC', cap(data.occupancy),
    roomNo||'TBD',
    fmtD(data.checkIn), fmtD(data.checkOut),
    Number(data.nights)||'', Number(data.extraPersons)||0,
    '₹'+num(data.roomTotal), '₹'+num(data.gst), '₹'+num(data.totalAmount),
    data.paymentId||'DEMO', 'Website', 'Confirmed', data.requests||'',
  ]);

  const last = sheet.getLastRow();
  sheet.getRange(last,1,1,19).setBackground(roomNo?'#DCFCE7':'#FEF9C3');
  if (roomNo) sheet.getRange(last,8).setFontWeight('bold').setFontColor('#166534');
  else        sheet.getRange(last,8).setFontColor('#DC2626').setFontWeight('bold');
}

// ================================================================
//  HOTEL EMAIL
// ================================================================
function sendHotelEmail(data, bookingId, roomNo) {
  const fmtD = s => { const d=safeDate(s); return d?fmtDate(d,'EEE d MMM yyyy'):String(s||''); };
  const banner = roomNo
    ? `<div style="background:#DCFCE7;border:1px solid #86EFAC;border-radius:6px;padding:14px;margin-bottom:16px"><div style="font-size:11px;color:#166534;letter-spacing:1px">AUTO-ALLOCATED ROOM</div><div style="font-size:24px;font-weight:bold;color:#166534">Room ${roomNo} ✅</div></div>`
    : `<div style="background:#FEF3C7;border:1px solid #FCD34D;border-radius:6px;padding:14px;margin-bottom:16px"><div style="font-size:11px;color:#92400E;letter-spacing:1px">ROOM ALLOCATION</div><div style="font-size:15px;color:#92400E">⚠️ No ${cap(data.roomType)} rooms free — assign manually</div></div>`;

  const subject = `🏨 ${bookingId} | ${data.guestName} | ${fmtD(data.checkIn)}${roomNo?' | Rm '+roomNo:''}`;

  const html = `<div style="font-family:Arial,sans-serif;max-width:560px;margin:0 auto">
  <div style="background:#1B2A4A;padding:18px;text-align:center">
    <div style="color:#C9A84C;font-size:19px;font-weight:bold">🏨 Hotel Sahara — New Booking</div>
    <div style="color:rgba(255,255,255,0.55);font-size:11px;margin-top:3px">${fmtNow()}</div>
  </div>
  <div style="padding:18px;background:#F8F5EF">${banner}
  <table style="width:100%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden;font-size:13px">
    <tr><td colspan="2" style="background:#1B2A4A;color:#C9A84C;padding:8px 14px;font-weight:bold;font-size:11px;letter-spacing:1px">BOOKING: ${bookingId}</td></tr>
    <tr><td style="padding:8px 14px;color:#6B7280;width:36%">Guest</td><td style="padding:8px 14px;font-weight:bold">${data.guestName}</td></tr>
    <tr style="background:#F9FAFB"><td style="padding:8px 14px;color:#6B7280">Phone</td><td style="padding:8px 14px"><a href="tel:${data.guestPhone}">${data.guestPhone}</a></td></tr>
    <tr><td style="padding:8px 14px;color:#6B7280">Email</td><td style="padding:8px 14px"><a href="mailto:${data.guestEmail}">${data.guestEmail}</a></td></tr>
    <tr style="background:#F9FAFB"><td style="padding:8px 14px;color:#6B7280">Room</td><td style="padding:8px 14px;font-weight:bold">${cap(data.roomType)} AC — ${cap(data.occupancy)}</td></tr>
    <tr><td style="padding:8px 14px;color:#6B7280">Check-In</td><td style="padding:8px 14px;font-weight:bold">${fmtD(data.checkIn)} at 1:00 PM</td></tr>
    <tr style="background:#F9FAFB"><td style="padding:8px 14px;color:#6B7280">Check-Out</td><td style="padding:8px 14px;font-weight:bold">${fmtD(data.checkOut)} at 11:00 AM</td></tr>
    <tr><td style="padding:8px 14px;color:#6B7280">Nights</td><td style="padding:8px 14px">${data.nights}</td></tr>
    <tr style="background:#DCFCE7"><td style="padding:10px 14px;color:#166534;font-weight:bold;font-size:13px">Total Paid</td><td style="padding:10px 14px;font-size:19px;font-weight:bold;color:#166534">₹${num(data.totalAmount)}</td></tr>
    <tr><td style="padding:8px 14px;color:#6B7280">Payment ID</td><td style="padding:8px 14px;font-family:monospace;font-size:11px">${data.paymentId||'DEMO'}</td></tr>
    ${data.requests?`<tr style="background:#FEF3C7"><td colspan="2" style="padding:8px 14px"><strong>Requests:</strong> ${data.requests}</td></tr>`:''}
  </table>
  <div style="margin-top:12px;text-align:center">
    <a href="${SpreadsheetApp.getActiveSpreadsheet().getUrl()}" style="background:#1B2A4A;color:#C9A84C;padding:10px 22px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:13px;display:inline-block">Open Booking Sheet →</a>
  </div></div>
  <div style="background:#1B2A4A;padding:10px;text-align:center;font-size:11px;color:rgba(255,255,255,0.4)">Hotel Sahara · Senapati Bapat Road, Pune · ${CONFIG.HOTEL_PHONE}</div>
</div>`;

  GmailApp.sendEmail(CONFIG.HOTEL_EMAIL, subject, '', {htmlBody:html, name:'Hotel Sahara Booking System', replyTo:data.guestEmail||CONFIG.HOTEL_EMAIL});
}

// ================================================================
//  REFRESH AVAILABILITY GRID
// ================================================================
function refreshAvailability(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const availSheet = ss.getSheetByName('📅 Availability');
  if (!availSheet) return;

  const today = new Date(); today.setHours(0,0,0,0);

  // Build flat occupied map: "roomNo|YYYY-MM-DD" → label
  const oMap = {};

  // From bookings
  const bSheet = ss.getSheetByName('📋 Bookings');
  if (bSheet && bSheet.getLastRow() > 1) {
    bSheet.getDataRange().getValues().slice(1).forEach(row => {
      const rNo   = String(row[7]).trim();
      const status= String(row[17]).trim();
      if (!rNo||rNo==='TBD'||rNo===''||status==='Cancelled') return;
      const ci = safeDate(row[8]);
      const co = safeDate(row[9]);
      if (!ci||!co) return;
      const label = String(row[2]||'Guest').split(' ')[0];
      markRange(oMap, rNo, ci, co, label);
    });
  }

  // From blocked
  const blkSheet = ss.getSheetByName('🚫 Blocked');
  if (blkSheet && blkSheet.getLastRow() > 1) {
    blkSheet.getDataRange().getValues().slice(1).forEach(row => {
      const rNo = String(row[0]).trim();
      if (!rNo||rNo==='') return;
      const ci = safeDate(row[1]);
      const co = safeDate(row[2]);
      if (!ci||!co) return;
      const reason = String(row[3]||'BLOCKED').replace('OTA - ','').substring(0,10);
      markRange(oMap, rNo, ci, co, '🚫'+reason);
    });
  }

  // Write to grid in batches
  const bookable = ROOMS.filter(r => r.status !== 'Non-Existent');
  bookable.forEach((room, rIdx) => {
    const row = rIdx + 2;
    if (room.status === 'Permanent') return;
    const rNo = String(room.no);
    const vals=[], bgs=[], fcs=[], fws=[], fss=[];
    for (let d=0;d<90;d++) {
      const cd = new Date(today); cd.setDate(today.getDate()+d);
      const key = rNo+'|'+dateKey(cd);
      const lbl = oMap[key];
      if (lbl) {
        const isBlk = lbl.startsWith('🚫');
        const isHold= lbl === 'HOLD';
        vals.push(lbl.substring(0,10));
        bgs.push(isHold?'#DBEAFE':isBlk?'#FEF3C7':'#FEE2E2');
        fcs.push(isHold?'#1E40AF':isBlk?'#92400E':'#991B1B');
        fws.push('bold'); fss.push(8);
      } else {
        vals.push(''); bgs.push('#D1FAE5'); fcs.push('#166534'); fws.push('normal'); fss.push(9);
      }
    }
    const rng = availSheet.getRange(row,4,1,90);
    rng.setValues([vals]).setBackgrounds([bgs]).setFontColors([fcs]).setFontWeights([fws]).setFontSizes([fss]);
  });
}

function markRange(map, rNo, ci, co, label) {
  const d   = new Date(ci); d.setHours(0,0,0,0);
  const end = new Date(co); end.setHours(0,0,0,0);
  while (d < end) { map[rNo+'|'+dateKey(d)] = label; d.setDate(d.getDate()+1); }
}

// ================================================================
//  SETUP (safe — never wipes existing data)
// ================================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename('🏨 Hotel Sahara — Booking System');
  ensureBookingsSheet(ss);
  ensureRoomsSheet(ss);
  ensureBlockedSheet(ss);
  ensureAvailabilitySheet(ss);
  ensureTodaySheet(ss);
  try { const d=ss.getSheetByName('Sheet1'); if(d) ss.deleteSheet(d); } catch(e) {}
  addMenu();
  SpreadsheetApp.getUi().alert('✅ Setup complete! Existing data was NOT deleted.');
}

function addBlockedSheet() {
  ensureBlockedSheet(SpreadsheetApp.getActiveSpreadsheet());
  SpreadsheetApp.getUi().alert('✅ 🚫 Blocked sheet ready.');
}

function ensureBookingsSheet(ss) {
  let s = ss.getSheetByName('📋 Bookings');
  if (!s) s = ss.insertSheet('📋 Bookings', 0);
  if (s.getLastRow() === 0) {
    const h=['Booking ID','Received At','Guest Name','Phone','Email','Room Type','Occupancy',
             '🔑 Room No','Check-In','Check-Out','Nights','Extra','Room Charges','GST',
             'Total Paid','Payment ID','Source','Status','Special Requests'];
    s.getRange(1,1,1,h.length).setValues([h]).setBackground('#1B2A4A').setFontColor('#C9A84C').setFontWeight('bold').setFontSize(10);
    s.setFrozenRows(1);
  }
  s.getRange(2,18,500,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Confirmed','Pending','Checked-In','Checked-Out','Cancelled','No-Show'],true).build());
  s.getRange(2,17,500,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Website','Walk-in','Phone','OTA - MakeMyTrip','OTA - OYO','OTA - Booking.com','Other'],true).build());
}

function ensureRoomsSheet(ss) {
  if (ss.getSheetByName('🏨 Rooms')) return;
  const s = ss.insertSheet('🏨 Rooms', 1);
  s.getRange(1,1,1,7).setValues([['Room No','Floor','Type','Status','Single Tariff','Double Tariff','Notes']]).setBackground('#1B2A4A').setFontColor('#C9A84C').setFontWeight('bold');
  s.setFrozenRows(1);
  const T={Executive:[2300,2800],Deluxe:[2500,3000],'N/A':[0,0]};
  s.getRange(2,1,ROOMS.length,7).setValues(ROOMS.map(r=>[r.no,r.floor,r.type,r.status,r.type!=='N/A'?'₹'+T[r.type][0]:'',r.type!=='N/A'?'₹'+T[r.type][1]:'',r.note]));
  ROOMS.forEach((r,i)=>{const bg=r.status==='Permanent'?'#FEE2E2':r.status==='Non-Existent'?'#E5E7EB':r.type==='Deluxe'?'#FEF9C3':'#F0F9FF';s.getRange(i+2,1,1,7).setBackground(bg);});
}

function ensureBlockedSheet(ss) {
  if (ss.getSheetByName('🚫 Blocked')) return;
  const s = ss.insertSheet('🚫 Blocked');
  s.getRange(1,1,1,7).setValues([['Room No','Check-In','Check-Out','Reason / Source','Block ID','Blocked On','Blocked By']]).setBackground('#92400E').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(10);
  s.setFrozenRows(1);
  [90,110,110,200,165,145,180].forEach((w,i)=>s.setColumnWidth(i+1,w));
  s.getRange(2,4,500,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['OTA - MakeMyTrip','OTA - OYO','OTA - Booking.com','Walk-in','Phone Booking','Maintenance','Other'],true).build());
}

function ensureAvailabilitySheet(ss) {
  let s = ss.getSheetByName('📅 Availability') || ss.insertSheet('📅 Availability');
  s.clearContents();
  const today=new Date();
  const headers=['Room','Type','Floor'];
  for(let i=0;i<90;i++){const d=new Date(today);d.setDate(today.getDate()+i);headers.push(Utilities.formatDate(d,CONFIG.TIMEZONE,'dd MMM'));}
  s.getRange(1,1,1,headers.length).setValues([headers]).setBackground('#1B2A4A').setFontColor('#C9A84C').setFontWeight('bold').setFontSize(9);
  s.setFrozenRows(1);s.setFrozenColumns(3);
  s.setColumnWidth(1,65);s.setColumnWidth(2,80);s.setColumnWidth(3,52);
  for(let c=4;c<=headers.length;c++)s.setColumnWidth(c,60);
  const bookable=ROOMS.filter(r=>r.status!=='Non-Existent');
  bookable.forEach((r,i)=>{
    const row=i+2;
    s.getRange(row,1).setValue(r.no);s.getRange(row,2).setValue(r.type);s.getRange(row,3).setValue('F'+r.floor);
    s.getRange(row,1,1,3).setBackground(r.type==='Deluxe'?'#FEF9C3':'#F0F9FF');
    if(r.status==='Permanent')s.getRange(row,4,1,90).setValue('PERM').setBackground('#FEE2E2').setFontColor('#991B1B').setFontWeight('bold').setFontSize(8);
    else s.getRange(row,4,1,90).setBackground('#D1FAE5');
  });
  const leg=bookable.length+3;
  s.getRange(leg,1).setValue('LEGEND →');
  [['🟢 Available','#D1FAE5'],['🔴 Guest name (booked)','#FEE2E2'],['🟡 🚫 Blocked/OTA','#FEF3C7'],['🔵 HOLD (multi-alloc)','#DBEAFE']].forEach(([t,bg],j)=>s.getRange(leg,j+2).setValue(t).setBackground(bg));
}

function ensureTodaySheet(ss) {
  if (!ss.getSheetByName('🗓️ Today')) ss.insertSheet('🗓️ Today');
  updateTodaySheet();
}

function updateTodaySheet() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('🗓️ Today'),bSheet=ss.getSheetByName('📋 Bookings');
  if(!sheet||!bSheet)return;
  sheet.clearContents();
  const today=new Date();today.setHours(0,0,0,0);
  sheet.getRange(1,1).setValue('🗓️  Daily Summary — '+Utilities.formatDate(today,CONFIG.TIMEZONE,'EEEE, d MMMM yyyy')).setFontSize(14).setFontWeight('bold');
  sheet.getRange(2,1).setValue('Updated: '+fmtNow()).setFontColor('#9CA3AF').setFontSize(10);
  const rows=bSheet.getLastRow()>1?bSheet.getDataRange().getValues().slice(1):[];
  const tk=dateKey(today);
  const ci=rows.filter(r=>{const d=safeDate(r[8]);return d&&dateKey(d)===tk&&r[17]!=='Cancelled';});
  const co=rows.filter(r=>{const d=safeDate(r[9]);return d&&dateKey(d)===tk&&r[17]!=='Cancelled';});
  const so=rows.filter(r=>{const a=safeDate(r[8]),b=safeDate(r[9]);return a&&b&&a<today&&b>today&&r[17]!=='Cancelled';});
  const cols=['Booking ID','Guest Name','Phone','Room Type','Room No','Check-In','Check-Out','Nights','Total Paid'];
  const mapR=r=>[r[0],r[2],r[3],r[5],r[7]||'TBD',String(r[8]||''),String(r[9]||''),r[10],r[14]];
  let sr=4;
  [{label:'⬇  CHECK-INS TODAY',bg:'#DCFCE7',hBg:'#166534',data:ci},
   {label:'⬆  CHECK-OUTS TODAY',bg:'#FEF3C7',hBg:'#92400E',data:co},
   {label:'🛏  STAYING OVER',bg:'#DBEAFE',hBg:'#1E40AF',data:so}].forEach(sec=>{
    sheet.getRange(sr,1).setValue(sec.label+' ('+sec.data.length+')').setFontWeight('bold').setBackground(sec.bg).setFontSize(11);
    sheet.getRange(sr+1,1,1,cols.length).setValues([cols]).setBackground(sec.hBg).setFontColor('#FFFFFF').setFontWeight('bold');
    if(sec.data.length>0)sheet.getRange(sr+2,1,sec.data.length,cols.length).setValues(sec.data.map(mapR));
    else sheet.getRange(sr+2,1).setValue('— None today —').setFontColor('#9CA3AF');
    sr+=3+Math.max(sec.data.length,1)+1;
  });
  sheet.setColumnWidths(1,cols.length,145);
}

// ================================================================
//  MENU
// ================================================================
function onOpen() { addMenu(); }

function addMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🏨 Hotel Sahara')
    .addItem('🚫 Block Selected Cells (on Availability sheet)', 'blockSelectedCells')
    .addItem('🔓 Unblock a Room (by Block ID)',                 'unblockRoomDialog')
    .addSeparator()
    .addItem('🛏  Allocate Multiple Rooms',                     'allocateMultipleRooms')
    .addItem('✅ Save & Auto-Allocate (current Bookings row)',  'onBookingManualEntry')
    .addSeparator()
    .addItem('🔄 Refresh Availability Grid',                    'refreshAvailability')
    .addItem('🗓️  Update Today\'s Summary',                     'updateTodaySheet')
    .addSeparator()
    .addItem('➕ Add Blocked Sheet (if missing)',               'addBlockedSheet')
    .addItem('🔧 Full Setup (safe — keeps existing data)',      'setupSheets')
    .addToUi();
}

// ================================================================
//  CORE UTILITIES
// ================================================================
function safeDate(v) {
  if (!v) return null;
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  const s = String(v).trim();
  if (!s||s==='undefined'||s==='null') return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s+'T00:00:00');
  const months={Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11};
  const p=s.split(' ');
  if (p.length===3&&months[p[1]]!==undefined) return new Date(parseInt(p[2]),months[p[1]],parseInt(p[0]));
  const d=new Date(s);
  return isNaN(d.getTime())?null:new Date(d.getFullYear(),d.getMonth(),d.getDate());
}

function getOccupiedRooms(ss, ciDate, coDate) {
  const occupied=[];
  const bSheet=ss.getSheetByName('📋 Bookings');
  if (bSheet&&bSheet.getLastRow()>1) {
    bSheet.getDataRange().getValues().slice(1).forEach(row=>{
      const rNo=String(row[7]).trim(),status=String(row[17]).trim();
      if(!rNo||rNo==='TBD'||rNo===''||status==='Cancelled')return;
      const ci=safeDate(row[8]),co=safeDate(row[9]);
      if(!ci||!co)return;
      if(ciDate<co&&coDate>ci)occupied.push(rNo);
    });
  }
  const blkSheet=ss.getSheetByName('🚫 Blocked');
  if (blkSheet&&blkSheet.getLastRow()>1) {
    blkSheet.getDataRange().getValues().slice(1).forEach(row=>{
      const rNo=String(row[0]).trim();
      if(!rNo||rNo==='')return;
      const ci=safeDate(row[1]),co=safeDate(row[2]);
      if(!ci||!co)return;
      if(ciDate<co&&coDate>ci)occupied.push(rNo);
    });
  }
  return occupied;
}

function dateKey(d) { return Utilities.formatDate(d,CONFIG.TIMEZONE,'yyyy-MM-dd'); }
function fmtDate(d,pat){if(!d)return '';return Utilities.formatDate(d instanceof Date?d:new Date(d),CONFIG.TIMEZONE,pat||'dd MMM yyyy');}
function fmtNow(){return fmtDate(new Date(),'dd MMM yyyy HH:mm');}
function generateId(){const n=new Date();return 'SHR-'+n.getFullYear()+String(n.getMonth()+1).padStart(2,'0')+String(n.getDate()).padStart(2,'0')+'-'+Math.floor(1000+Math.random()*9000);}
function cap(s){return s?s.charAt(0).toUpperCase()+s.slice(1).toLowerCase():'';}
function num(v){return parseFloat(v||0).toLocaleString('en-IN',{maximumFractionDigits:0});}
function jsonResponse(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
