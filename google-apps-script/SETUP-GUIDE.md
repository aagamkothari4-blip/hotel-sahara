# Hotel Sahara — Google Sheets Booking System Setup

## What This Creates
- **📋 All Bookings** — master log of every booking (website + manual)
- **🏨 Room Directory** — all 34 rooms with types and rates
- **📅 Today's View** — auto-refreshes daily: arrivals, departures, in-house
- **📊 60-Day Calendar** — colour-coded availability grid

---

## Step 1 — Create the Google Sheet

1. Go to **https://sheets.google.com** → click **+** New spreadsheet
2. Name it "Hotel Sahara — Booking System"

---

## Step 2 — Open Apps Script

1. In the sheet: **Extensions → Apps Script**
2. Delete all existing code in the editor
3. Paste the entire contents of **`booking-system.gs`** (this folder)
4. Click **💾 Save** (Ctrl+S)

---

## Step 3 — Run Initial Setup

1. In Apps Script, select function **`setupSpreadsheet`** from the dropdown
2. Click ▶ **Run**
3. Grant permissions when prompted (it needs access to Sheets + Gmail)
4. All 4 sheets will be created automatically ✅

---

## Step 4 — Deploy as Web App

1. Click **Deploy → New Deployment**
2. Click ⚙️ gear icon → **Web app**
3. Settings:
   - **Description:** Hotel Sahara Booking API
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy**
5. **Copy the Web App URL** — looks like:
   `https://script.google.com/macros/s/AKfycb.../exec`

---

## Step 5 — Connect to Website

Open `js/booking.js` and replace:
```javascript
googleScriptUrl: 'YOUR_GOOGLE_APPS_SCRIPT_URL',
```
with your copied URL:
```javascript
googleScriptUrl: 'https://script.google.com/macros/s/AKfycb.../exec',
```

Then run `push.bat` — website bookings will now flow directly to the sheet.

---

## Step 6 — Set Up Daily Auto-Refresh

1. Back in Apps Script, run function **`installDailyTrigger`**
2. "Today's View" sheet will auto-refresh every morning at 6 AM IST ✅

---

## How to Add Manual Bookings (Phone / Walk-in)

**Option A — Type directly into the sheet:**
Open "📋 All Bookings" → scroll to the next empty row → fill in:
- Booking ID (leave blank, auto-fills if blank)
- Source: **Manual**
- Guest Name, Phone, Room No, Room Type, Check-In, Check-Out, etc.

**Option B — Use the menu:**
Open the sheet → **🏨 Hotel Sahara menu** → **➕ Add Manual Booking**
→ Enter: `GuestName, Phone, RoomNo, CheckIn(dd-mm-yyyy), CheckOut`

---

## ⚠️ Action Required: Confirm Deluxe Room Numbers

Current assumption (update in `booking-system.gs` if wrong):

| Floor | Deluxe Rooms (assumed) | Executive Rooms |
|-------|------------------------|-----------------|
| 1     | 109, 110, 112          | 102–108         |
| 2     | 210, 211, 212          | 201–209         |
| 3     | 310, 311, 312          | 301–309         |

To change: edit the `ROOM_DIRECTORY` array at the top of `booking-system.gs`
and re-run `setupSpreadsheet()`.

---

## Room Reference

| Room | Status |
|------|--------|
| 101  | Permanently Occupied |
| 111  | Does Not Exist |
| All others | Active |
