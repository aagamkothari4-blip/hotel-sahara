# 🏨 Hotel Sahara — Website Setup Guide

## Running on Localhost

1. Open Terminal / Command Prompt
2. Navigate to this folder:
   ```
   cd hotel-sahara
   ```
3. Start a local server:
   ```
   python3 -m http.server 8000
   ```
4. Open your browser and go to:
   ```
   http://localhost:8000
   ```

> **Windows users:** Use `python -m http.server 8000` if `python3` doesn't work.

---

## Adding Your Hotel Photos

Place your real photos inside the `images/` folder with these exact filenames:

| Filename               | What it is              |
|------------------------|-------------------------|
| `exterior.jpg`         | Hotel building outside  |
| `exterior-2.jpg`       | Another exterior shot   |
| `reception.jpg`        | Front desk / lobby      |
| `corridor.jpg`         | Hotel corridor          |
| `room-executive.jpg`   | Executive AC room       |
| `room-deluxe.jpg`      | Deluxe AC room          |
| `room-amenities.jpg`   | Room TV / amenities     |
| `room-wardrobe.jpg`    | Wardrobe / dressing     |
| `room-single.jpg`      | Single room view        |
| `bathroom.jpg`         | Bathroom                |
| `restaurant.jpg`       | Restaurant / dining     |
| `lounge.jpg`           | Lobby lounge area       |

> **Tip:** Resize photos to max 1200×900px (JPG, 80% quality) to keep the site fast.
> Until you add your own photos, the site shows beautiful placeholder images from the internet.

---

## Admin Dashboard

- **URL:** `http://localhost:8000/admin.html`
- **Username:** `admin`
- **Password:** `Sahara@2024`

To change the password, open `admin.html` and find:
```javascript
const ADMIN_PASSWORD = 'Sahara@2024';
```
Change it to your preferred password and save.

---

## Setting Up Online Payments (Razorpay)

1. Create an account at https://razorpay.com
2. Go to **Settings → API Keys → Generate Test Key**
3. Copy your **Key ID** (starts with `rzp_test_...`)
4. Open `js/booking.js` and replace:
   ```javascript
   razorpayKeyId: 'YOUR_RAZORPAY_KEY_ID',
   ```
   with your actual key:
   ```javascript
   razorpayKeyId: 'rzp_test_xxxxxxxxxxxxxxxx',
   ```

> **Test cards for Razorpay:**
> - Card: `4111 1111 1111 1111`
> - Expiry: Any future date
> - CVV: Any 3 digits

When ready to go live, generate a **Live Key** and replace `rzp_test_...` with `rzp_live_...`

---

## Setting Up Email Confirmations (EmailJS)

EmailJS sends automated booking confirmation emails to guests — **no backend needed.**

### Step 1: Create EmailJS Account
1. Go to https://www.emailjs.com and sign up (free tier: 200 emails/month)
2. Go to **Email Services** → Add New Service → Gmail (or Outlook)
3. Note your **Service ID** (e.g. `service_abc1234`)

### Step 2: Create Email Template
1. Go to **Email Templates** → Create New Template
2. Use these variables in your template:

```
Subject: Booking Confirmed — {{booking_id}} | Hotel Sahara

Dear {{guest_name}},

Your booking at Hotel Sahara is CONFIRMED!

Booking ID:    {{booking_id}}
Room:          {{room_type}}
Check-In:      {{check_in}} at 1:00 PM
Check-Out:     {{check_out}} at 11:00 AM
Nights:        {{nights}}
Total Paid:    {{total_amount}} (incl. GST)
Payment ID:    {{payment_id}}

Complimentary buffet breakfast is included.
Please carry a valid photo ID at check-in.

For any queries:
📞 {{hotel_phone}}
📧 {{hotel_email}}
💬 WhatsApp: {{hotel_whatsapp}}

We look forward to welcoming you!
Hotel Sahara, Pune
```

3. Note your **Template ID** (e.g. `template_xyz5678`)
4. Go to **Account** → note your **Public Key**

### Step 3: Update `js/booking.js`
```javascript
emailjsServiceId:  'service_abc1234',
emailjsTemplateId: 'template_xyz5678',
emailjsPublicKey:  'YOUR_PUBLIC_KEY',
```

---

## Deploying to hotelsaharapune.com

### Option A: cPanel / File Manager (most common)
1. Log in to your hosting cPanel
2. Go to **File Manager** → `public_html`
3. Upload all files from the `hotel-sahara/` folder
4. Your site is live at `hotelsaharapune.com`

### Option B: FTP (FileZilla)
1. Download FileZilla: https://filezilla-project.org
2. Connect with your hosting FTP credentials
3. Upload everything to the `public_html` folder

### Option C: Netlify (Free hosting, very easy)
1. Go to https://netlify.com and sign up
2. Drag and drop the `hotel-sahara/` folder onto Netlify
3. Set your custom domain: `hotelsaharapune.com`

---

## Before Going Live Checklist

- [ ] Add real hotel photos to `images/` folder
- [ ] Replace Razorpay test key with live key in `js/booking.js`
- [ ] Configure EmailJS (service ID, template ID, public key) in `js/booking.js`
- [ ] Update the Google Maps embed in `index.html` and `contact.html` with the exact hotel location
- [ ] Change admin password in `admin.html`
- [ ] Test a complete booking flow end-to-end

---

## File Structure

```
hotel-sahara/
├── index.html              Homepage
├── rooms.html              Rooms & Booking
├── gallery.html            Photo Gallery
├── about.html              About Hotel Sahara
├── contact.html            Contact
├── admin.html              Admin Dashboard
├── booking-success.html    Booking Confirmation
├── css/
│   └── style.css           All styles
├── js/
│   ├── main.js             Navbar, animations, toast
│   └── booking.js          Booking engine (Razorpay + EmailJS)
└── images/
    └── (put hotel photos here)
```

---

## Support

For any issues with the website, contact your web developer.
For hotel queries: bookingsahara@rediffmail.com | 020-25655405/6/8/9
