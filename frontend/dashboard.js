const apiBase = CONFIG.API_BASE;

// Check authentication on page load
document.addEventListener('DOMContentLoaded', async () => {
    const token = localStorage.getItem('token');
    const email = localStorage.getItem('email');

    if (!token || !email) {
        // Guest mode
        document.body.classList.remove('auth-mode');
        document.body.classList.add('guest-mode');
        setupGuestMode();
    } else {
        // Logged-in mode
        document.body.classList.remove('guest-mode');
        document.body.classList.add('auth-mode');
        initializeAuthUser(email);
    }

    // Setup event listeners
    setupEventListeners();
});

function setupGuestMode() {
    console.log('Running in Guest Mode');
}

async function initializeAuthUser(email) {
    // Display user info
    document.getElementById('user-email').textContent = email;
    const userName = email.split('@')[0];
    document.getElementById('user-name').textContent = userName;

    // Set avatar initials
    const initials = userName.substring(0, 2).toUpperCase();
    document.getElementById('avatar-initials').textContent = initials;

    // Load user profile
    await loadUserProfile();
}

// Setup all event listeners
function setupEventListeners() {
    // Navigation items
    document.querySelectorAll('.nav-item').forEach(item => {
        item.addEventListener('click', (e) => {
            e.preventDefault();
            const page = item.getAttribute('data-page');

            // Update active nav
            document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
            item.classList.add('active');

            // Show page
            if (page === 'settings') {
                showPage('settings');
            } else if (page === 'history') {
                showPage('history');
                loadHistory();
            } else if (page === 'support') {
                showPage('support');
            } else {
                // If clicking "New Conversion", reset the state
                if (page === 'converter' && item.innerText.includes('Yeni Dönüştürme')) {
                    resetConverterState();
                }
                showPage('converter');
            }
        });
    });

    // Settings button in header
    document.getElementById('open-settings').addEventListener('click', () => {
        showPage('settings');
        document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
        document.getElementById('nav-settings').classList.add('active');
    });

    // Back button (General)
    document.getElementById('back-btn').addEventListener('click', () => {
        showPage('converter');
        document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
        document.querySelector('.nav-item[data-page="converter"]').classList.add('active');
    });

    // Support Back button
    const supportBackBtn = document.getElementById('support-back-btn');
    if (supportBackBtn) {
        supportBackBtn.addEventListener('click', () => {
            showPage('converter');
            document.querySelectorAll('.nav-item').forEach(i => i.classList.remove('active'));
            document.querySelector('.nav-item[data-page="converter"]').classList.add('active');
        });
    }

    // Logout
    document.getElementById('logout-btn').addEventListener('click', logout);

    // Password form
    document.getElementById('password-form').addEventListener('submit', handlePasswordChange);

    // Toggle password visibility
    document.querySelectorAll('.toggle-password').forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.getAttribute('data-target');
            const input = document.getElementById(targetId);
            if (input.type === 'password') {
                input.type = 'text';
                btn.innerHTML = `<svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"></path><line x1="1" y1="1" x2="23" y2="23"></line></svg>`;
            } else {
                input.type = 'password';
                btn.innerHTML = `<svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>`;
            }
        });
    });

    // Password confirmation check
    if (document.getElementById('confirm-password')) {
        document.getElementById('confirm-password').addEventListener('input', checkPasswordMatch);
    }
    if (document.getElementById('new-password')) {
        document.getElementById('new-password').addEventListener('input', checkPasswordMatch);
    }

    // Profile form
    if (document.getElementById('profile-form')) {
        document.getElementById('profile-form').addEventListener('submit', handleProfileUpdate);
    }

    // File upload
    setupFileUpload();

    // Start Conversion Button
    const startBtn = document.getElementById('start-conversion-btn');
    if (startBtn) {
        startBtn.addEventListener('click', startConversion);
    }

    // Support Request form
    const supportForm = document.getElementById('support-form');
    if (supportForm) {
        supportForm.addEventListener('submit', handleSubmitSupportRequest);
    }
}

// Show page
function showPage(pageName) {
    document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
    document.getElementById(`${pageName}-page`).classList.add('active');
}

// File upload setup
function setupFileUpload() {
    const uploadZone = document.getElementById('upload-zone');
    const fileInput = document.getElementById('file-input');

    uploadZone.addEventListener('click', () => fileInput.click());

    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    });

    uploadZone.addEventListener('dragleave', () => {
        uploadZone.classList.remove('dragover');
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
        handleFiles(e.dataTransfer.files);
    });

    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
        // Clear the input value so the same file can be selected again
        fileInput.value = '';
    });
}

// Handle files
let selectedFiles = [];

function handleFiles(files) {
    // Append new files to existing list
    const newFiles = Array.from(files);
    selectedFiles = [...selectedFiles, ...newFiles];
    updateFilesList();

    // Show action area if we have files
    if (selectedFiles.length > 0) {
        document.getElementById('action-area').style.display = 'block';
    }
}

// Reset converter state
function resetConverterState() {
    selectedFiles = [];
    updateFilesList();
    const startBtn = document.getElementById('start-conversion-btn');
    if (startBtn) {
        startBtn.disabled = false;
        startBtn.textContent = 'Dönüştürmeyi Başlat';
    }
}

function updateFilesList() {
    const container = document.getElementById('files-list');

    if (selectedFiles.length === 0) {
        container.innerHTML = '<div class="empty-state"><p>Henüz işlem yok</p></div>';
        document.getElementById('action-area').style.display = 'none';
        return;
    }

    container.innerHTML = selectedFiles.map((file, index) => `
    <div class="file-item" id="file-item-${index}">
      <div class="file-item-icon">
        <svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"></path><polyline points="13 2 13 9 20 9"></polyline></svg>
      </div>
      <div class="file-item-info">
        <div class="file-item-name">${file.name}</div>
        <div class="file-item-progress">
          <div class="file-item-progress-bar" style="width: 0%"></div>
        </div>
        <div class="file-item-meta" id="status-${index}">${formatFileSize(file.size)} • Dönüştürmeye Hazır (0%)</div>
      </div>
      <div class="file-item-actions" id="actions-${index}">
        <button onclick="removeFile(${index})" class="btn-icon btn-action-delete" title="Kaldır">
            <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
        </button>
      </div>
    </div>
  `).join('');
}

// Convert Files
// Update progress UI
function updateProgressUI(index, percent, statusText) {
    const fileItem = document.getElementById(`file-item-${index}`);
    if (!fileItem) return;

    const progressBar = fileItem.querySelector('.file-item-progress-bar');
    const statusEl = document.getElementById(`status-${index}`);

    if (progressBar) {
        progressBar.style.width = `${percent}%`;
    }

    if (statusEl) {
        const file = selectedFiles[index];
        const fileSize = file ? formatFileSize(file.size) : '';

        if (percent >= 100) {
            statusEl.innerHTML = `${fileSize} • Tamamlandı`;
            statusEl.style.color = 'green';
        } else {
            statusEl.innerHTML = `${fileSize} • ${statusText} (${percent}%)`;
            statusEl.style.color = ''; // Reset color
        }
    }
}

// Simulate progress
function simulateProgress(index) {
    let percent = 0;
    updateProgressUI(index, 0, 'Dönüştürülüyor...');

    const interval = setInterval(() => {
        // Variable increment for realism
        let increment = Math.random() * 10;
        if (percent > 60) increment = Math.random() * 3;
        if (percent > 85) increment = Math.random() * 1;

        percent = Math.min(percent + increment, 95); // Cap at 95 until done
        updateProgressUI(index, Math.round(percent), 'Dönüştürülüyor...');
    }, 400);

    return interval;
}

// Convert Files
async function startConversion() {
    const btn = document.getElementById('start-conversion-btn');
    btn.disabled = true;
    btn.textContent = 'İşleniyor...';

    const sourceFormat = document.getElementById('source-format').value;
    const targetFormat = document.getElementById('target-format').value;

    for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        const statusEl = document.getElementById(`status-${i}`);
        const actionsEl = document.getElementById(`actions-${i}`);

        // Skip if already processed
        if (statusEl.textContent.includes('Tamamlandı')) continue;

        // Start simulation
        const progressInterval = simulateProgress(i);

        try {
            const formData = new FormData();
            formData.append('file', file);
            formData.append('sourceFormat', sourceFormat);
            formData.append('targetFormat', targetFormat);

            const token = localStorage.getItem('token');
            const headers = {};
            if (token) {
                headers['Authorization'] = `Bearer ${token}`;
            }

            const res = await fetch(`${apiBase}/conversion/convert`, {
                method: 'POST',
                headers: headers,
                body: formData
            });

            clearInterval(progressInterval);

            if (res.ok) {
                const blob = await res.blob();
                const url = window.URL.createObjectURL(blob);
                const fileName = file.name.replace(/\.[^/.]+$/, "") + ".pdf";

                updateProgressUI(i, 100, 'Tamamlandı');

                // Add individual download button
                // Modern SVG icons
                actionsEl.innerHTML = `
                    <a href="${url}" download="${fileName}" class="btn-icon btn-action-download" title="İndir">
                        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                    </a>
                    <button onclick="removeFile(${i})" class="btn-icon btn-action-delete" title="Kaldır">
                        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><line x1="10" y1="11" x2="10" y2="17"></line><line x1="14" y1="11" x2="14" y2="17"></line></svg>
                    </button>
                `;

                // Show confirmation modal
                await askToDownload(fileName, url);

            } else {
                const txt = await res.text();
                statusEl.innerHTML = `Hata: ${txt}`; // Use innerHTML to overwrite
                statusEl.style.color = 'red';
                const progressBar = document.getElementById(`file-item-${i}`).querySelector('.file-item-progress-bar');
                if (progressBar) progressBar.style.backgroundColor = 'red';
            }
        } catch (err) {
            console.error(err);
            clearInterval(progressInterval);
            statusEl.innerHTML = 'Bağlantı Hatası';
            statusEl.style.color = 'red';
        }
    }

    btn.disabled = false;
    btn.textContent = '⚡ Dönüştürmeyi Başlat';
}

// Modal Logic
function askToDownload(fileName, url) {
    return new Promise((resolve) => {
        const modal = document.getElementById('download-modal');
        const msg = document.getElementById('download-message');
        const confirmBtn = document.getElementById('modal-confirm-btn');
        const cancelBtn = document.getElementById('modal-cancel-btn');

        msg.textContent = `${fileName} hazır. İndirmek ister misiniz?`;
        modal.style.display = 'flex';

        // Clean previous listeners to avoid duplicates
        const newConfirm = confirmBtn.cloneNode(true);
        confirmBtn.parentNode.replaceChild(newConfirm, confirmBtn);

        const newCancel = cancelBtn.cloneNode(true);
        cancelBtn.parentNode.replaceChild(newCancel, cancelBtn);

        newConfirm.addEventListener('click', () => {
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);

            modal.style.display = 'none';
            resolve(true);
        });

        newCancel.addEventListener('click', () => {
            modal.style.display = 'none';
            resolve(false);
        });
    });
}

function removeFile(index) {
    selectedFiles.splice(index, 1);
    updateFilesList();
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// Check password match
function checkPasswordMatch() {
    const newPassword = document.getElementById('new-password').value;
    const confirmPassword = document.getElementById('confirm-password').value;
    const errorEl = document.getElementById('password-error');

    if (confirmPassword && newPassword !== confirmPassword) {
        errorEl.textContent = '⚠ Şifreler eşleşmiyor';
    } else {
        errorEl.textContent = '';
    }
}

// Load user profile from API
async function loadUserProfile() {
    const token = localStorage.getItem('token');

    try {
        const res = await fetch(`${apiBase}/user/profile`, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!res.ok) {
            if (res.status === 401) {
                logout();
                return;
            }
            return;
        }

        const profile = await res.json();

        // Update profile display
        if (document.getElementById('profile-email')) {
            document.getElementById('profile-email').value = profile.email;
        }

        if (profile.firstName) {
            document.getElementById('user-name').textContent = profile.firstName;
            if (document.getElementById('profile-name')) {
                document.getElementById('profile-name').value = profile.firstName;
            }
            const initials = profile.firstName.substring(0, 2).toUpperCase();
            document.getElementById('avatar-initials').textContent = initials;
        }

    } catch (error) {
        console.error('Profile load error:', error);
    }
}

// Handle password change
async function handlePasswordChange(e) {
    e.preventDefault();

    const form = e.target;
    const oldPassword = form.oldPassword.value;
    const newPassword = form.newPassword.value;
    const confirmPassword = form.confirmPassword.value;

    // Validation
    if (!oldPassword || !newPassword || !confirmPassword) {
        showStatus('password-status', 'Tüm alanlar gerekli', 'error');
        return;
    }

    if (newPassword !== confirmPassword) {
        showStatus('password-status', 'Yeni şifreler eşleşmiyor', 'error');
        return;
    }

    if (newPassword.length < 6) {
        showStatus('password-status', 'Yeni şifre en az 6 karakter olmalı', 'error');
        return;
    }

    const token = localStorage.getItem('token');

    try {
        const res = await fetch(`${apiBase}/user/change-password`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify({
                oldPassword,
                newPassword
            })
        });

        const text = await res.text();
        let data;
        try {
            data = JSON.parse(text);
        } catch {
            data = { message: text };
        }

        if (!res.ok) {
            showStatus('password-status', data.message || 'Şifre değiştirilemedi', 'error');
            return;
        }

        showStatus('password-status', 'Şifre değiştirme işlemi başarılı', 'success');
        showToast('Şifre başarıyla güncellendi');
        form.reset();

    } catch (error) {
        console.error('Password change error:', error);
        showStatus('password-status', 'Bir hata oluştu', 'error');
    }
}

// Show toast notification
function showToast(message) {
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.textContent = '✓ ' + message;
    document.body.appendChild(toast);

    setTimeout(() => {
        toast.remove();
    }, 3000);
}

// Logout function
function logout() {
    localStorage.removeItem('token');
    localStorage.removeItem('email');
    localStorage.removeItem('userId');
    window.location.href = '/index.html';
}



// Load user profile from API
async function loadUserProfile() {
    const token = localStorage.getItem('token');

    try {
        const res = await fetch(`${apiBase}/user/profile`, {
            headers: {
                'Authorization': `Bearer ${token}`
            }
        });

        if (!res.ok) {
            if (res.status === 401) {
                logout();
                return;
            }
            return;
        }

        const profile = await res.json();

        // Update profile display
        if (document.getElementById('profile-email')) {
            document.getElementById('profile-email').value = profile.email;
        }

        if (document.getElementById('profile-name')) {
            document.getElementById('profile-name').value = profile.firstName || '';
        }

        if (document.getElementById('profile-lastname')) {
            document.getElementById('profile-lastname').value = profile.lastName || '';
        }

        if (document.getElementById('profile-phone')) {
            document.getElementById('profile-phone').value = profile.phoneNumber || '';
        }

        if (profile.firstName) {
            document.getElementById('user-name').textContent = profile.firstName;
            const initials = profile.firstName.substring(0, 2).toUpperCase();
            document.getElementById('avatar-initials').textContent = initials;
        }

    } catch (error) {
        console.error('Profile load error:', error);
    }
}

// Handle profile update
async function handleProfileUpdate(e) {
    e.preventDefault();
    const btn = e.target.querySelector('button[type="submit"]');
    const originalText = btn.textContent;
    btn.textContent = 'Kaydediliyor...';
    btn.disabled = true;

    const firstName = document.getElementById('profile-name').value;
    const lastName = document.getElementById('profile-lastname').value;
    const phoneNumber = document.getElementById('profile-phone').value;
    const token = localStorage.getItem('token');

    try {
        const res = await fetch(`${apiBase}/user/profile`, {
            method: 'PUT',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${token}`
            },
            body: JSON.stringify({ firstName, lastName, phoneNumber })
        });

        const data = await res.json();

        if (!res.ok) {
            showStatus('profile-status', data.message || 'Güncelleme başarısız', 'error');
            return;
        }

        showStatus('profile-status', 'Profil bilgileri başarıyla güncellendi', 'success');
        showToast('Profil bilgileri güncellendi');

        // Update sidebar info
        if (firstName) {
            document.getElementById('user-name').textContent = firstName;
            document.getElementById('avatar-initials').textContent = firstName.substring(0, 2).toUpperCase();
        }

    } catch (error) {
        console.error('Profile update error:', error);
        showStatus('profile-status', 'Bir hata oluştu', 'error');
    } finally {
        btn.textContent = originalText;
        btn.disabled = false;
    }
}

// History Functions
async function loadHistory() {
    const tableBody = document.getElementById('history-table-body');
    if (!tableBody) return;

    tableBody.innerHTML = '<tr><td colspan="4" style="text-align:center; padding: 40px; color: #64748b;">Yükleniyor...</td></tr>';

    const token = localStorage.getItem('token');
    try {
        const res = await fetch(`${apiBase}/conversion/history`, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!res.ok) throw new Error('Geçmiş yüklenemedi');

        const history = await res.json();

        if (history.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="4" style="text-align:center; padding: 40px; color: #64748b;">Henüz işlem geçmişi yok</td></tr>';
            return;
        }

        tableBody.innerHTML = history.map(item => `
            <tr style="border-bottom: 1px solid #f1f5f9; transition: background 0.2s;" onmouseover="this.style.background='#f8fafc'" onmouseout="this.style.background='transparent'">
                <td style="padding: 16px 24px;">
                    <div style="display:flex; align-items:center; gap:12px;">
                        <div style="background: #eef2ff; color: #4f46e5; padding: 8px; border-radius: 8px; font-weight: 700; font-size: 11px; width: 32px; height: 32px; display: flex; align-items: center; justify-content: center;">
                            PDF
                        </div>
                        <div style="display: flex; flex-direction: column;">
                            <span style="font-weight: 500; color: #1e293b; font-size: 14px;">${item.originalFileName}</span>
                            <span style="font-size: 12px; color: #64748b;">Dönüştürüldü</span>
                        </div>
                    </div>
                </td>
                <td style="padding: 16px 24px; color: #64748b; font-size: 14px;">${formatFileSize(item.fileSize)}</td>
                <td style="padding: 16px 24px; color: #64748b; font-size: 14px;">${new Date(item.createdAt).toLocaleDateString('tr-TR', { day: 'numeric', month: 'long', year: 'numeric', hour: '2-digit', minute: '2-digit' })}</td>
                <td style="padding: 16px 24px; text-align: right;">
                    <button onclick="downloadHistoryFile(${item.id}, '${item.originalFileName}')" class="btn-icon" title="İndir" style="display:inline-flex; align-items:center; justify-content:center; color: #4f46e5; background: #eef2ff; border: none; padding: 8px; border-radius: 8px; cursor: pointer; transition: all 0.2s;">
                        <svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                    </button>
                </td>
            </tr>
        `).join('');

    } catch (err) {
        console.error(err);
        tableBody.innerHTML = '<tr><td colspan="4" style="text-align:center; color:red; padding: 40px;">Veriler yüklenirken bir hata oluştu.</td></tr>';
    }
}

async function downloadHistoryFile(id, originalFileName) {
    const token = localStorage.getItem('token');
    const btn = event.currentTarget || event.target.closest('button');
    const originalContent = btn.innerHTML;

    try {
        btn.disabled = true;
        btn.innerHTML = '<svg class="spinner" viewBox="0 0 24 24" width="18" height="18" style="animation: rotate 2s linear infinite;"><circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="3" style="stroke-dasharray: 31.4, 31.4;"></circle></svg>';

        const res = await fetch(`${apiBase}/conversion/download/${id}`, {
            headers: { 'Authorization': `Bearer ${token}` }
        });

        if (!res.ok) throw new Error('İndirme başarısız');

        const blob = await res.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;

        // Ensure name ends with pdf
        let fileName = originalFileName;
        if (!fileName.toLowerCase().endsWith('.pdf')) {
            const pos = fileName.lastIndexOf('.');
            fileName = (pos > 0 ? fileName.substring(0, pos) : fileName) + ".pdf";
        }

        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

    } catch (err) {
        console.error(err);
        alert('Dosya indirilirken bir hata oluştu. Dosya sunucudan silinmiş olabilir.');
    } finally {
        btn.disabled = false;
        btn.innerHTML = originalContent;
    }
}

// Show status message
function showStatus(elementId, message, type) {
    const el = document.getElementById(elementId);
    if (!el) return;

    el.textContent = message;
    el.className = 'status-message ' + (type === 'success' ? 'status-success' : 'status-error');

    // Clear after 3 seconds if success
    if (type === 'success') {
        setTimeout(() => {
            el.textContent = '';
            el.className = 'status-message';
        }, 3000);
    }
}

async function handleSubmitSupportRequest(e) {
    e.preventDefault();
    const subject = document.getElementById('support-subject').value;
    const message = document.getElementById('support-message').value;
    const statusEl = document.getElementById('support-status');
    const submitBtn = document.getElementById('support-submit-btn');
    const originalBtnText = submitBtn.textContent;

    const token = localStorage.getItem('token');
    const headers = { 'Content-Type': 'application/json' };
    if (token) {
        headers['Authorization'] = `Bearer ${token}`;
    }

    try {
        submitBtn.disabled = true;
        submitBtn.textContent = 'Gönderiliyor...';

        const res = await fetch(`${apiBase}/support/submit`, {
            method: 'POST',
            headers: headers,
            body: JSON.stringify({ subject, message })
        });

        if (res.ok) {
            const data = await res.json();
            showStatus('support-status', data.message || 'Talebiniz başarıyla gönderildi.', 'success');
            document.getElementById('support-form').reset();
        } else {
            const errorData = await res.json().catch(() => ({}));
            throw new Error(errorData.message || 'Talep gönderilemedi.');
        }
    } catch (err) {
        console.error(err);
        showStatus('support-status', `Hata: ${err.message}`, 'error');
    } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = originalBtnText;
    }
}
