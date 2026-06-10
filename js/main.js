/* ============================================================
   Hotel Sahara — main.js
   Shared across all pages: navbar, scroll animations, toast
   ============================================================ */

// ---- Navbar scroll effect ----
const navbar = document.getElementById('navbar');
if(navbar){
  // Pages with transparent hero start with no 'scrolled' class
  // Inner pages already have 'scrolled' class in HTML
  window.addEventListener('scroll', () => {
    if(window.scrollY > 60){
      navbar.classList.add('scrolled');
    } else {
      // Only remove on homepage (no 'scrolled' in initial HTML)
      if(!navbar.classList.contains('force-scrolled')){
        navbar.classList.remove('scrolled');
      }
    }
  });
}

// ---- Mobile nav toggle ----
function toggleNav(){
  const links  = document.getElementById('navLinks');
  const toggle = document.getElementById('navToggle');
  if(!links || !toggle) return;
  links.classList.toggle('open');
  toggle.classList.toggle('open');
}

// Close nav when a link is clicked
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('#navLinks a').forEach(a => {
    a.addEventListener('click', () => {
      document.getElementById('navLinks')?.classList.remove('open');
      document.getElementById('navToggle')?.classList.remove('open');
    });
  });

  // Hero background loaded class (ken burns start)
  const heroBg = document.getElementById('heroBg');
  if(heroBg){
    setTimeout(() => heroBg.classList.add('loaded'), 100);
  }

  // Init scroll animations
  initScrollAnimations();

  // Set active nav link
  setActiveNavLink();

  // Set min dates for any date inputs
  initDateInputs();
});

// ---- Scroll animations (IntersectionObserver) ----
function initScrollAnimations(){
  const elements = document.querySelectorAll('.fade-up');
  if(!elements.length) return;

  const observer = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if(entry.isIntersecting){
        entry.target.classList.add('visible');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.12, rootMargin: '0px 0px -40px 0px' });

  elements.forEach((el, i) => {
    // Staggered delay for siblings
    el.style.transitionDelay = `${(i % 4) * 0.1}s`;
    observer.observe(el);
  });
}

// ---- Set active nav link based on current page ----
function setActiveNavLink(){
  const page = window.location.pathname.split('/').pop() || 'index.html';
  document.querySelectorAll('.navbar-links a').forEach(a => {
    const href = a.getAttribute('href')?.split('?')[0];
    if(href === page || (page === '' && href === 'index.html')){
      a.classList.add('active');
    } else {
      a.classList.remove('active');
    }
  });
}

// ---- Initialize date inputs ----
function initDateInputs(){
  const today = new Date().toISOString().split('T')[0];
  document.querySelectorAll('input[type="date"]').forEach(input => {
    if(!input.min) input.min = today;
  });
}

// ---- Toast notification ----
let toastTimeout;
function showToast(message, type = 'success'){
  let toast = document.getElementById('globalToast');
  if(!toast){
    toast = document.createElement('div');
    toast.id = 'globalToast';
    toast.className = 'toast';
    document.body.appendChild(toast);
  }

  toast.className = `toast ${type}`;
  toast.innerHTML = `<i class="fa fa-${type === 'success' ? 'check-circle' : 'exclamation-circle'}"></i> ${message}`;

  clearTimeout(toastTimeout);
  setTimeout(() => toast.classList.add('show'), 10);

  toastTimeout = setTimeout(() => {
    toast.classList.remove('show');
  }, 3500);
}

// ---- Close modal on overlay click ----
document.addEventListener('click', (e) => {
  if(e.target.classList.contains('modal-overlay')){
    e.target.classList.remove('open');
    document.body.style.overflow = '';
  }
});

// Close lightbox on background click
document.addEventListener('click', (e) => {
  const lb = document.getElementById('lightbox');
  if(lb && e.target === lb) closeLightbox?.();
});
