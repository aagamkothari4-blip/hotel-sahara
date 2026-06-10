/* ============================================================
   Hotel Sahara — booking.js
   Booking engine: price calc, Razorpay payment, EmailJS, localStorage
   ============================================================ */

// ============================================================
// CONFIGURATION — Fill these in before going live
// ============================================================
const CONFIG = {
  razorpayKeyId:    'YOUR_RAZORPAY_KEY_ID',
  emailjsServiceId: 'YOUR_EMAILJS_SERVICE_ID',
  emailjsTemplateId:'YOUR_EMAILJS_TEMPLATE_ID',
  emailjsPublicKey: 'YOUR_EMAILJS_PUBLIC_KEY',
  // Paste your Google Apps Script Web App URL here after deploying
  googleScriptUrl:  'YOUR_GOOGLE_APPS_SCRIPT_URL',
  hotelName:        'Hotel Sahara',
  hotelPhone:       '020-25655405/6/8/9',
  hotelEmail:       'bookingsahara@rediffmail.com',
  hotelWhatsApp:    '+91 98223 93889',
};

// ============================================================
// ROOM RATES (per night in ₹)
// ============================================================
const ROOM_RATES = {
  'executive-single': 2300,
  'executive-double': 2800,
  'deluxe-single':    2500,
  'deluxe-double':    3000,
};
const ROOM_NAMES = {
  executive: 'Executive AC',
  deluxe:    'Deluxe AC',
};
const OCC_NAMES = {
  single: 'Single Occupancy',
  double: 'Double Occupancy',
};
const EXTRA_PERSON_RATE = 600;
const GST_RATE          = 0.12;

// ============================================================
// MODAL OPEN / CLOSE
// ============================================================
function openBookingModal(roomType = 'executive'){
  const modal = document.getElementById('bookingModal');
  if(!modal) return;

  // Pre-set room type
  const roomSelect = document.getElementById('modalRoomType');
  if(roomSelect) roomSelect.value = roomType;

  // Init Flatpickr date pickers (safe to call multiple times)
  if(typeof initModalDatePickers === 'function') initModalDatePickers();

  // Pre-fill from URL if present and pickers not already set
  const params = new URLSearchParams(window.location.search);
  if(params.get('checkin') && window.checkinPicker && !window.checkinPicker.selectedDates.length){
    window.checkinPicker.setDate(params.get('checkin'));
  }
  if(params.get('checkout') && window.checkoutPicker && !window.checkoutPicker.selectedDates.length){
    window.checkoutPicker.setDate(params.get('checkout'));
  }

  updateModalLabel();
  calculatePrice();

  modal.classList.add('open');
  document.body.style.overflow = 'hidden';
}

function closeBookingModal(){
  const modal = document.getElementById('bookingModal');
  if(modal) modal.classList.remove('open');
  document.body.style.overflow = '';
}

function updateModalLabel(){
  const roomType = document.getElementById('modalRoomType')?.value;
  const occ      = document.getElementById('modalOccupancy')?.value;
  const label    = document.getElementById('modalRoomLabel');
  if(label && roomType && occ){
    label.textContent = `${ROOM_NAMES[roomType]} — ${OCC_NAMES[occ]}`;
  }
}

// ============================================================
// PRICE CALCULATION
// ============================================================
function calculatePrice(){
  updateModalLabel();

  const roomType    = document.getElementById('modalRoomType')?.value;
  const occupancy   = document.getElementById('modalOccupancy')?.value;
  const checkin     = document.getElementById('modalCheckin')?.value;
  const checkout    = document.getElementById('modalCheckout')?.value;
  const extraPersons= parseInt(document.getElementById('modalExtraPersons')?.value || '0');

  // Enforce checkout > checkin
  if(checkin && checkout && checkout <= checkin){
    const next = new Date(checkin);
    next.setDate(next.getDate() + 1);
    document.getElementById('modalCheckout').value = next.toISOString().split('T')[0];
    return;
  }

  if(!roomType || !occupancy || !checkin || !checkout){
    document.getElementById('priceBreakdown')?.classList.remove('show');
    document.getElementById('footerTotal').textContent = '₹ —';
    return;
  }

  const nights = Math.max(1, Math.ceil(
    (new Date(checkout) - new Date(checkin)) / (1000 * 60 * 60 * 24)
  ));

  const rateKey    = `${roomType}-${occupancy}`;
  const baseRate   = ROOM_RATES[rateKey] || 0;
  const roomTotal  = baseRate * nights;
  const extraTotal = extraPersons * EXTRA_PERSON_RATE * nights;
  const subtotal   = roomTotal + extraTotal;
  const gst        = Math.round(subtotal * GST_RATE);
  const grandTotal = subtotal + gst;

  // Update breakdown UI
  document.getElementById('bdRoom').textContent      = `${ROOM_NAMES[roomType]} (${OCC_NAMES[occupancy]})`;
  document.getElementById('bdRoomRate').textContent  = `₹${baseRate.toLocaleString('en-IN')}/night`;
  document.getElementById('bdNights').textContent    = `${nights} Night${nights > 1 ? 's' : ''}`;
  document.getElementById('bdRoomTotal').textContent = `₹${roomTotal.toLocaleString('en-IN')}`;

  const extraRow = document.getElementById('bdExtraRow');
  if(extraPersons > 0){
    extraRow.style.display = 'flex';
    document.getElementById('bdExtra').textContent      = `Extra ${extraPersons} Person${extraPersons > 1 ? 's' : ''} × ${nights} night${nights > 1 ? 's' : ''}`;
    document.getElementById('bdExtraTotal').textContent = `₹${extraTotal.toLocaleString('en-IN')}`;
  } else {
    extraRow.style.display = 'none';
  }

  document.getElementById('bdGst').textContent   = `₹${gst.toLocaleString('en-IN')}`;
  document.getElementById('bdTotal').innerHTML   = `<strong>₹${grandTotal.toLocaleString('en-IN')}</strong>`;
  document.getElementById('footerTotal').textContent = `₹${grandTotal.toLocaleString('en-IN')}`;

  document.getElementById('priceBreakdown')?.classList.add('show');

  return { nights, baseRate, roomTotal, extraTotal, subtotal, gst, grandTotal };
}

// ============================================================
// FORM VALIDATION
// ============================================================
function validateBookingForm(){
  const fields = {
    'modalCheckin':  'Check-in date',
    'modalCheckout': 'Check-out date',
    'guestName':     'Full name',
    'guestEmail':    'Email address',
    'guestPhone':    'Mobile number',
  };

  for(const [id, label] of Object.entries(fields)){
    const el = document.getElementById(id);
    if(!el || !el.value.trim()){
      showToast(`Please enter your ${label}.`, 'error');
      el?.focus();
      return false;
    }
  }

  // Email format
  const email = document.getElementById('guestEmail').value;
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)){
    showToast('Please enter a valid email address.', 'error');
    document.getElementById('guestEmail').focus();
    return false;
  }

  // Phone format (10 digits minimum)
  const phone = document.getElementById('guestPhone').value.replace(/\D/g,'');
  if(phone.length < 10){
    showToast('Please enter a valid 10-digit mobile number.', 'error');
    document.getElementById('guestPhone').focus();
    return false;
  }

  // Date check
  const checkin  = document.getElementById('modalCheckin').value;
  const checkout = document.getElementById('modalCheckout').value;
  if(checkout <= checkin){
    showToast('Check-out must be after check-in.', 'error');
    return false;
  }

  return true;
}

// ============================================================
// PROCEED TO PAYMENT
// ============================================================
function proceedToPayment(){
  if(!validateBookingForm()) return;

  const pricing = calculatePrice();
  if(!pricing){
    showToast('Please select your dates first.', 'error');
    return;
  }

  const roomType    = document.getElementById('modalRoomType').value;
  const occupancy   = document.getElementById('modalOccupancy').value;
  const checkin     = document.getElementById('modalCheckin').value;
  const checkout    = document.getElementById('modalCheckout').value;
  const extra       = parseInt(document.getElementById('modalExtraPersons').value || '0');
  const guestName   = document.getElementById('guestName').value.trim();
  const guestEmail  = document.getElementById('guestEmail').value.trim();
  const guestPhone  = document.getElementById('guestPhone').value.trim();
  const requests    = document.getElementById('guestRequests')?.value?.trim() || '';

  const bookingData = {
    roomType, occupancy, checkin, checkout,
    extraPersons: extra,
    nights:       pricing.nights,
    baseRate:     pricing.baseRate,
    roomTotal:    pricing.roomTotal,
    extraTotal:   pricing.extraTotal,
    gst:          pricing.gst,
    grandTotal:   pricing.grandTotal,
    guestName, guestEmail, guestPhone, requests,
  };

  // ---- Check if Razorpay key is configured ----
  if(CONFIG.razorpayKeyId === 'YOUR_RAZORPAY_KEY_ID'){
    // DEMO MODE: simulate payment for localhost testing
    console.warn('⚠️  Razorpay key not configured. Running in demo mode.');
    simulateDemoPayment(bookingData);
    return;
  }

  // ---- REAL Razorpay Payment ----
  initRazorpayPayment(bookingData);
}

// ============================================================
// RAZORPAY PAYMENT
// ============================================================
function initRazorpayPayment(bookingData){
  const payBtn = document.getElementById('payBtn');
  payBtn.disabled = true;
  payBtn.innerHTML = '<i class="fa fa-spinner fa-spin"></i>&nbsp; Opening Payment...';

  const options = {
    key:         CONFIG.razorpayKeyId,
    amount:      Math.round(bookingData.grandTotal * 100), // paise
    currency:    'INR',
    name:        CONFIG.hotelName,
    description: `${ROOM_NAMES[bookingData.roomType]} — ${bookingData.nights} Night${bookingData.nights > 1 ? 's' : ''}`,
    image:       '/images/logo.png',
    prefill: {
      name:    bookingData.guestName,
      email:   bookingData.guestEmail,
      contact: bookingData.guestPhone,
    },
    notes: {
      checkin:   bookingData.checkin,
      checkout:  bookingData.checkout,
      roomType:  bookingData.roomType,
      occupancy: bookingData.occupancy,
    },
    theme: { color: '#C9A84C' },
    modal: {
      ondismiss: () => {
        payBtn.disabled = false;
        payBtn.innerHTML = '<i class="fa fa-lock"></i>&nbsp; Proceed to Pay';
        showToast('Payment cancelled.', 'error');
      }
    },
    handler: function(response){
      // Payment successful
      const bookingId = generateBookingId();
      const booking   = {
        bookingId,
        paymentId:   response.razorpay_payment_id,
        roomType:    bookingData.roomType,
        occupancy:   bookingData.occupancy,
        checkIn:     bookingData.checkin,
        checkOut:    bookingData.checkout,
        nights:      bookingData.nights,
        extraPersons:bookingData.extraPersons,
        totalAmount: bookingData.grandTotal,
        guestName:   bookingData.guestName,
        guestEmail:  bookingData.guestEmail,
        guestPhone:  bookingData.guestPhone,
        requests:    bookingData.requests,
        status:      'confirmed',
        createdAt:   new Date().toISOString(),
      };

      saveBooking(booking);
      sendConfirmationEmail(booking);
      closeBookingModal();
      window.location.href = `booking-success.html?id=${bookingId}`;
    },
  };

  const rzp = new Razorpay(options);
  rzp.open();

  payBtn.disabled = false;
  payBtn.innerHTML = '<i class="fa fa-lock"></i>&nbsp; Proceed to Pay';
}

// ============================================================
// DEMO MODE (when Razorpay key not yet set — localhost testing)
// ============================================================
function simulateDemoPayment(bookingData){
  const payBtn = document.getElementById('payBtn');
  payBtn.disabled = true;
  payBtn.innerHTML = '<i class="fa fa-spinner fa-spin"></i>&nbsp; Processing (Demo)...';

  setTimeout(() => {
    const bookingId = generateBookingId();
    const booking   = {
      bookingId,
      paymentId:   'DEMO_' + Date.now(),
      roomType:    bookingData.roomType,
      occupancy:   bookingData.occupancy,
      checkIn:     bookingData.checkin,
      checkOut:    bookingData.checkout,
      nights:      bookingData.nights,
      extraPersons:bookingData.extraPersons,
      totalAmount: bookingData.grandTotal,
      guestName:   bookingData.guestName,
      guestEmail:  bookingData.guestEmail,
      guestPhone:  bookingData.guestPhone,
      requests:    bookingData.requests,
      status:      'pending', // pending in demo mode
      createdAt:   new Date().toISOString(),
    };

    saveBooking(booking);
    closeBookingModal();
    showToast('✅ Demo booking saved! (No payment charged)', 'success');

    setTimeout(() => {
      window.location.href = `booking-success.html?id=${bookingId}`;
    }, 800);
  }, 1500);
}

// ============================================================
// SAVE BOOKING TO LOCALSTORAGE + GOOGLE SHEETS
// ============================================================
function saveBooking(booking){
  // 1. Always save locally (fallback)
  const bookings = JSON.parse(localStorage.getItem('saharaBookings') || '[]');
  bookings.unshift(booking);
  localStorage.setItem('saharaBookings', JSON.stringify(bookings));

  // 2. Send to Google Sheets (primary backend)
  sendToGoogleSheets(booking);

  console.log('Booking saved:', booking.bookingId);
}

// ============================================================
// SEND TO GOOGLE SHEETS
// ============================================================
function sendToGoogleSheets(booking) {
  if (!CONFIG.googleScriptUrl || CONFIG.googleScriptUrl === 'YOUR_GOOGLE_APPS_SCRIPT_URL') {
    console.warn('⚠️  Google Sheets URL not configured. Booking saved locally only.');
    return;
  }

  const payload = {
    bookingId:    booking.bookingId,
    guestName:    booking.guestName,
    guestPhone:   booking.guestPhone,
    guestEmail:   booking.guestEmail,
    roomType:     booking.roomType,
    occupancy:    booking.occupancy,
    checkIn:      booking.checkIn,
    checkOut:     booking.checkOut,
    nights:       booking.nights,
    extraPersons: booking.extraPersons || 0,
    baseRate:     booking.baseRate,
    roomTotal:    booking.roomTotal,
    gst:          booking.gst,
    totalAmount:  booking.totalAmount,
    paymentId:    booking.paymentId,
    requests:     booking.requests || '',
  };

  // Use no-cors mode — Google Apps Script handles CORS
  fetch(CONFIG.googleScriptUrl, {
    method:  'POST',
    mode:    'no-cors',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(payload),
  }).then(() => {
    console.log('✅ Booking sent to Google Sheets');
  }).catch(err => {
    console.error('⚠️  Google Sheets sync failed (booking still saved locally):', err);
  });
}

// ============================================================
// SEND CONFIRMATION EMAIL VIA EMAILJS
// ============================================================
function sendConfirmationEmail(booking){
  if(CONFIG.emailjsPublicKey === 'YOUR_EMAILJS_PUBLIC_KEY'){
    console.warn('⚠️  EmailJS not configured. Skipping email.');
    return;
  }

  emailjs.init(CONFIG.emailjsPublicKey);

  const checkInDate  = formatDateLong(booking.checkIn);
  const checkOutDate = formatDateLong(booking.checkOut);

  emailjs.send(CONFIG.emailjsServiceId, CONFIG.emailjsTemplateId, {
    booking_id:    booking.bookingId,
    guest_name:    booking.guestName,
    guest_email:   booking.guestEmail,
    guest_phone:   booking.guestPhone,
    room_type:     ROOM_NAMES[booking.roomType] + ' (' + OCC_NAMES[booking.occupancy] + ')',
    check_in:      checkInDate,
    check_out:     checkOutDate,
    nights:        booking.nights,
    extra_persons: booking.extraPersons || 0,
    total_amount:  `₹${parseFloat(booking.totalAmount).toLocaleString('en-IN', {minimumFractionDigits: 2})}`,
    payment_id:    booking.paymentId,
    hotel_name:    CONFIG.hotelName,
    hotel_phone:   CONFIG.hotelPhone,
    hotel_email:   CONFIG.hotelEmail,
    hotel_whatsapp:CONFIG.hotelWhatsApp,
    checkin_time:  '1:00 PM',
    checkout_time: '11:00 AM',
  }).then(() => {
    console.log('Confirmation email sent to', booking.guestEmail);
  }).catch(err => {
    console.error('EmailJS error:', err);
  });
}

// ============================================================
// GENERATE BOOKING ID
// ============================================================
function generateBookingId(){
  const now    = new Date();
  const year   = now.getFullYear();
  const month  = String(now.getMonth() + 1).padStart(2, '0');
  const day    = String(now.getDate()).padStart(2, '0');
  const random = Math.floor(1000 + Math.random() * 9000);
  return `SHR-${year}${month}${day}-${random}`;
}

// ============================================================
// UTILS
// ============================================================
function formatDateLong(ds){
  if(!ds) return '—';
  const d = new Date(ds + 'T00:00:00');
  return d.toLocaleDateString('en-IN', {
    weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
  });
}
