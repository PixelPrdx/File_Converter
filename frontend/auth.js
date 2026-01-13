const apiBase = CONFIG.API_BASE + '/auth';

// Şifremi Unuttum Form Toggles
document.getElementById('show-register').addEventListener('click', (e) => { e.preventDefault(); toggleForms('register') });
document.getElementById('show-login').addEventListener('click', (e) => { e.preventDefault(); toggleForms('login') });
document.getElementById('show-forgot').addEventListener('click', (e) => { e.preventDefault(); toggleForms('forgot') });
document.getElementById('show-login-from-forgot').addEventListener('click', (e) => { e.preventDefault(); toggleForms('login') });
document.getElementById('show-login-from-reset').addEventListener('click', (e) => { e.preventDefault(); toggleForms('login') });

function toggleForms(formName) {
  document.getElementById('login-form').style.display = 'none';
  document.getElementById('register-form').style.display = 'none';
  document.getElementById('forgot-form').style.display = 'none';
  document.getElementById('reset-form').style.display = 'none';
  document.getElementById(formName + '-form').style.display = 'block';
  document.getElementById('status').textContent = '';
}

// Giriş
document.getElementById('login-form').addEventListener('submit', async e => {
  e.preventDefault();
  const f = e.target;
  const email = f.email.value; const password = f.password.value;
  if (!email || !password) { setStatus('Email ve şifre gerekli', 'error'); return; }
  const res = await fetch(`${apiBase}/login`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ email, password }) });
  if (!res.ok) {
    const err = await res.text();
    setStatus(err || 'Giriş başarısız', 'error');
    return;
  }
  const data = await res.json();
  localStorage.setItem('token', data.token);
  localStorage.setItem('email', data.user.email);
  localStorage.setItem('userId', data.user.id);
  setStatus('Giriş başarılı! ✓', 'success');
  setTimeout(() => { document.location.href = '/converter.html'; }, 1000);
});

// Şifremi Unuttum
document.getElementById('forgot-form').addEventListener('submit', async e => {
  e.preventDefault();
  const f = e.target;
  const email = f.email.value;
  if (!email) { setStatus('Email gerekli', 'error'); return; }
  const res = await fetch(`${apiBase}/forgot-password`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ email }) });
  if (!res.ok) { setStatus('İstek başarısız', 'error'); return; }
  const data = await res.json();
  setStatus(`✓ Sıfırlama kodu: ${data.resetToken} (Test için)`, 'info');
  toggleForms('reset');
});

// Load countries on page load
document.addEventListener('DOMContentLoaded', async () => {
  try {
    const response = await fetch('countries.json');
    const countries = await response.json();
    const select = document.getElementById('country-code');

    // Sort countries by name (optional) or use Turkey as default top
    // Find Turkey first to default select it

    countries.forEach(country => {
      const option = document.createElement('option');
      option.value = country.dial_code;
      // Format: 🇹🇷 +90
      option.text = `${country.flag} ${country.dial_code}`;
      if (country.code === 'TR') option.selected = true;
      select.appendChild(option);
    });
  } catch (error) {
    console.error('Error loading countries:', error);
  }

  // Handle URL mode parameter
  const params = new URLSearchParams(window.location.search);
  const mode = params.get('mode');
  const resetToken = params.get('resetToken');

  if (resetToken) {
    toggleForms('reset');
    document.querySelector('#reset-form [name="token"]').value = resetToken;
  } else if (mode === 'register') {
    toggleForms('register');
  } else if (mode === 'login') {
    toggleForms('login');
  }
});

// Kayıt
document.getElementById('register-form').addEventListener('submit', async e => {
  e.preventDefault();
  const f = e.target;
  const firstName = f.firstName.value.trim();
  const lastName = f.lastName.value.trim();
  const email = f.email.value.trim();
  const password = f.password.value.trim();
  const confirmPassword = f.confirmPassword.value.trim();
  const countryCode = f.countryCode.value;
  const phone = f.phone.value.trim();

  if (!firstName || !lastName || !email || !password || !confirmPassword) { setStatus('Tüm alanlar gerekli', 'error'); return; }
  if (password !== confirmPassword) { setStatus('Şifreler eşleşmiyor', 'error'); return; }

  // Send country code and phone number separately
  // Note: Backend expects generic "PhoneNumber" field which we map to the raw number input

  const res = await fetch(`${apiBase}/register`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ email, password, firstName, lastName, CountryCode: countryCode, PhoneNumber: phone }) });
  const txt = await res.json();

  if (res.ok) {
    let msg = 'Kayıt başarılı! Lütfen emailinize gelen doğrulama linkine tıklayın.';
    if (txt.verifyLink) {
      msg += `<br><br><a href="${txt.verifyLink}" style="color: white; font-weight: bold; text-decoration: underline;">Doğrulama Linki (Tıklayın)</a>`;
    }
    setStatus(msg, 'success');
    f.reset();
    setTimeout(() => toggleForms('login'), 8000); // Increased time for user to click
  } else {
    setStatus(`Hata: ${txt.message || 'Kayıt başarısız'}`, 'error');
  }
});

// Şifre Sıfırla
document.getElementById('reset-form').addEventListener('submit', async e => {
  e.preventDefault();
  const f = e.target;
  const token = f.token.value;
  const newPassword = f.newPassword.value;
  if (!token || !newPassword) { setStatus('Tüm alanlar gerekli', 'error'); return; }
  const res = await fetch(`${apiBase}/reset-password`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ token, newPassword }) });
  setStatus(res.ok ? '✓ Şifre değiştirildi. Giriş yapabilirsiniz.' : 'Hata: geçersiz kod', res.ok ? 'success' : 'error');
  if (res.ok) setTimeout(() => toggleForms('login'), 1500);
});

// Password Toggle Logic
document.querySelectorAll('.toggle-password').forEach(btn => {
  btn.addEventListener('click', () => {
    const input = document.getElementById(btn.dataset.target);
    const type = input.getAttribute('type') === 'password' ? 'text' : 'password';
    input.setAttribute('type', type);

    // Update SVG icon (optional but nice)
    // Simple slash for hidden state
    if (type === 'text') {
      btn.innerHTML = `<svg viewBox="0 0 24 24"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle><line x1="1" y1="1" x2="23" y2="23" stroke="currentColor" stroke-width="2"></line></svg>`;
    } else {
      btn.innerHTML = `<svg viewBox="0 0 24 24"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>`;
    }
  });
});

function setStatus(msg, type = 'info') {
  const el = document.getElementById('status');
  el.innerHTML = msg; // Changed to innerHTML to support links
  el.className = 'status-' + type;
}
