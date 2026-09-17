'use strict';

// ── State ─────────────────────────────────────────────────────────────────────
let parsedRecipients = [];
let sendMode = 'all';

// ── DOM refs ──────────────────────────────────────────────────────────────────
const excelFileInput = document.getElementById('excelFile');
const fileUploadRow  = document.getElementById('fileUploadRow');
const fileBadge      = document.getElementById('fileBadge');
const fileNameLabel  = document.getElementById('fileNameLabel');
const countPill      = document.getElementById('countPill');
const clearFileBtn   = document.getElementById('clearFileBtn');
const personField    = document.getElementById('personField');
const personDropdown = document.getElementById('personDropdown');
const subjectInput   = document.getElementById('subject');
const messageInput   = document.getElementById('message');
const sendBtn        = document.getElementById('sendBtn');
const sendBtnText    = document.getElementById('sendBtnText');
const spinner        = document.getElementById('spinner');
const resultsPanel   = document.getElementById('resultsPanel');
const toast          = document.getElementById('toast');

// ── File Upload Listener ──────────────────────────────────────────────────────

excelFileInput.addEventListener('change', function () {
  const file = this.files[0];
  if (file) handleFile(file);
});

clearFileBtn.addEventListener('click', resetFile);

// ── handleFile ────────────────────────────────────────────────────────────────

function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) {
    showToast('❌ Please upload a .xlsx or .xls file.', 'error');
    return;
  }

  const reader = new FileReader();

  reader.onerror = function () {
    showToast('❌ Could not read the file. Try again.', 'error');
  };

  reader.onload = function (e) {
    try {
      const data     = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      if (!workbook.SheetNames.length) {
        showToast('❌ Excel file has no sheets.', 'error');
        return;
      }

      const sheetName = workbook.SheetNames[0];
      const sheet     = workbook.Sheets[sheetName];

      const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });

      console.log(`[Excel] Sheet: "${sheetName}" | Rows: ${rows.length}`);

      if (!rows.length) {
        showToast('❌ The Excel sheet is empty.', 'error');
        return;
      }

      const allKeys = Object.keys(rows[0]);
      console.log('[Excel] Columns found:', allKeys);

      const nameKey = allKeys.find(k => k.trim().toLowerCase().includes('name'));
      const emailKey = allKeys.find(k => {
        const lower = k.trim().toLowerCase().replace(/[^a-z]/g, '');
        return lower.includes('email') || lower.includes('mail');
      });

      console.log(`[Excel] Detected → Name column: "${nameKey}" | Email column: "${emailKey}"`);

      if (!nameKey && !emailKey) {
        showToast(`❌ Could not find Name or Email columns. Found: ${allKeys.join(', ')}`, 'error');
        return;
      }
      if (!nameKey) {
        showToast(`❌ No "Name" column found. Your columns: ${allKeys.join(', ')}`, 'error');
        return;
      }
      if (!emailKey) {
        showToast(`❌ No "Email" column found. Your columns: ${allKeys.join(', ')}`, 'error');
        return;
      }

      const allRecipients = rows.map((row) => {
        const obj = {};
        allKeys.forEach(k => { obj[k] = String(row[k] ?? '').trim(); });
        obj['Name']  = String(row[nameKey]  ?? '').trim();
        obj['Email'] = String(row[emailKey] ?? '').trim();
        return obj;
      });

      parsedRecipients = allRecipients.filter((r, idx) => {
        const valid = r.Email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(r.Email);
        if (!valid) console.warn(`[Excel] Row ${idx + 2} skipped — invalid email: "${r.Email}"`);
        return valid;
      });

      console.log(`[Excel] Valid recipients: ${parsedRecipients.length}`);

      if (!parsedRecipients.length) {
        showToast('❌ No valid email addresses found. Check the Email column contains real emails.', 'error');
        return;
      }

      fileNameLabel.textContent       = file.name;
      countPill.textContent           = `${parsedRecipients.length} recipient${parsedRecipients.length !== 1 ? 's' : ''}`;
      fileUploadRow.style.display     = 'none';
      fileBadge.style.display         = 'flex';

      populateDropdown();
      showToast(`✅ Loaded ${parsedRecipients.length} recipients from "${sheetName}"`, 'success');

    } catch (err) {
      console.error('[Excel] Parse error:', err);
      showToast(`❌ Failed to read Excel: ${err.message}`, 'error');
    }
  };

  reader.readAsArrayBuffer(file);
}

// ── Reset ─────────────────────────────────────────────────────────────────────

function resetFile() {
  parsedRecipients            = [];
  excelFileInput.value        = '';
  fileBadge.style.display     = 'none';
  fileUploadRow.style.display = 'flex';
  resultsPanel.style.display  = 'none';
  personDropdown.innerHTML    = '<option value="">— Upload an Excel file first —</option>';
}

// ── Populate dropdown ─────────────────────────────────────────────────────────

function populateDropdown() {
  personDropdown.innerHTML = '<option value="">— Select a person —</option>';
  parsedRecipients.forEach((r, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = `${r.Name || '(no name)'}  —  ${r.Email}`;
    personDropdown.appendChild(opt);
  });
}

// ── Send Mode Toggle ──────────────────────────────────────────────────────────

function getSendBtnLabel() {
  return sendMode === 'all' ? 'Send to All' : 'Send Mail';
}

function setSendMode(mode) {
  sendMode = mode;
  document.getElementById('btnAll').classList.toggle('active', mode === 'all');
  document.getElementById('btnOne').classList.toggle('active', mode === 'one');
  personField.style.display = mode === 'one' ? 'flex' : 'none';

  // ── UPDATE SEND BUTTON LABEL ──
  sendBtnText.textContent = getSendBtnLabel();
}

// ── Personalization ───────────────────────────────────────────────────────────

function personalize(text, recipient) {
  return text.replace(/\{\{(\w[\w\s]*)\}\}/g, function (match, key) {
    const trimmedKey = key.trim();
    if (Object.prototype.hasOwnProperty.call(recipient, trimmedKey)) {
      return recipient[trimmedKey];
    }
    const found = Object.keys(recipient).find(
      k => k.toLowerCase() === trimmedKey.toLowerCase()
    );
    return found ? recipient[found] : match;
  });
}

// ── Send Mails ────────────────────────────────────────────────────────────────

async function sendMails() {
  if (!parsedRecipients.length) {
    showToast('❌ Please upload an Excel file first.', 'error');
    return;
  }

  const subject = subjectInput.value.trim();
  const message = messageInput.value.trim();

  if (!subject) { showToast('❌ Subject is required.', 'error'); return; }
  if (!message) { showToast('❌ Message body is required.', 'error'); return; }

  let recipients = [];

  if (sendMode === 'all') {
    recipients = parsedRecipients.map(r => ({
      name:    r.Name,
      email:   r.Email,
      subject: personalize(subject, r),
      message: personalize(message, r),
    }));
  } else {
    const idx = personDropdown.value;
    if (idx === '') {
      showToast('❌ Please select a person from the dropdown.', 'error');
      return;
    }
    const r = parsedRecipients[parseInt(idx, 10)];
    recipients = [{
      name:    r.Name,
      email:   r.Email,
      subject: personalize(subject, r),
      message: personalize(message, r),
    }];
  }

  setLoading(true);
  resultsPanel.style.display = 'none';

  try {
    const res = await fetch('/send-mails', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ recipients }),
    });

    const data = await res.json();

    if (!res.ok) {
      showToast(data.error || '❌ Server error. Is the backend running?', 'error');
      return;
    }

    showResults(data);
    showToast(
      data.failureCount === 0
        ? `✅ All ${data.successCount} email(s) sent successfully!`
        : `⚠️ ${data.successCount} sent, ${data.failureCount} failed.`,
      data.failureCount === 0 ? 'success' : 'error'
    );

  } catch (err) {
    console.error('[Fetch] Error:', err);
    showToast('❌ Cannot reach the server. Run: node server.js', 'error');
  } finally {
    setLoading(false);
  }
}

// ── UI Helpers ────────────────────────────────────────────────────────────────

function setLoading(on) {
  sendBtn.disabled        = on;
  spinner.style.display   = on ? 'inline-block' : 'none';
  // Restore correct label based on current mode
  sendBtnText.textContent = on ? 'Sending…' : getSendBtnLabel();
}

function showResults(data) {
  document.getElementById('statTotal').textContent   = data.total        ?? 0;
  document.getElementById('statSuccess').textContent = data.successCount ?? 0;
  document.getElementById('statFailed').textContent  = data.failureCount ?? 0;

  const failedBlock = document.getElementById('failedBlock');
  const failedList  = document.getElementById('failedList');
  failedList.innerHTML = '';

  if (data.failed && data.failed.length) {
    data.failed.forEach(f => {
      const li = document.createElement('li');
      li.textContent = `${f.email} — ${f.error}`;
      failedList.appendChild(li);
    });
    failedBlock.style.display = 'block';
  } else {
    failedBlock.style.display = 'none';
  }

  resultsPanel.style.display = 'block';
  resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

let toastTimer;
function showToast(msg, type = '') {
  clearTimeout(toastTimer);
  toast.textContent = msg;
  toast.className   = `toast ${type} show`;
  toastTimer = setTimeout(() => { toast.className = 'toast'; }, 4500);
}