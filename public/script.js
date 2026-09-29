'use strict';

// ── State ─────────────────────────────────────────────────────────────────────
let parsedRecipients = [];   // ALL rows from Excel (valid + invalid) — nothing is dropped here anymore
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

// ── Column detection helpers ─────────────────────────────────────────────────
// FIX: previously this picked the FIRST column whose name merely *contained*
// "mail" or "name" anywhere (e.g. "Alternate Mail", "Domain Name",
// "Voicemail" would all match "mail"/"name" and could be picked by mistake).
// Now we look for exact / well-known header names first, and only fall back
// to loose substring matching if nothing exact is found. Every candidate is
// logged to the console so mismatches are easy to diagnose.

function normalizeHeader(k) {
  return k.trim().toLowerCase().replace(/[^a-z]/g, '');
}

function findEmailColumn(allKeys) {
  const exactNames = ['email', 'emailaddress', 'emailid', 'email id', 'e-mail', 'mailid', 'mail'];
  // 1) Exact (normalized) match against known email header names
  for (const candidate of exactNames) {
    const norm = normalizeHeader(candidate);
    const hit = allKeys.find(k => normalizeHeader(k) === norm);
    if (hit) return hit;
  }
  // 2) Fallback: loose substring match (old behaviour), but log every match
  //    found so a wrong pick is visible instead of silent.
  const looseMatches = allKeys.filter(k => {
    const lower = normalizeHeader(k);
    return lower.includes('email') || lower.includes('mail');
  });
  if (looseMatches.length > 1) {
    console.warn('[Excel] Multiple possible email columns found:', looseMatches, '→ using first:', looseMatches[0]);
  }
  return looseMatches[0];
}

function findNameColumn(allKeys) {
  const exactNames = ['name', 'fullname', 'full name', 'employeename', 'employee name'];
  for (const candidate of exactNames) {
    const norm = normalizeHeader(candidate);
    const hit = allKeys.find(k => normalizeHeader(k) === norm);
    if (hit) return hit;
  }
  const looseMatches = allKeys.filter(k => normalizeHeader(k).includes('name'));
  if (looseMatches.length > 1) {
    console.warn('[Excel] Multiple possible name columns found:', looseMatches, '→ using first:', looseMatches[0]);
  }
  return looseMatches[0];
}

// ── Email format check (used only for UI hints — server does real validation) ─
function looksLikeEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

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

      const nameKey  = findNameColumn(allKeys);
      const emailKey = findEmailColumn(allKeys);

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

      // FIX: we no longer silently filter out rows with a malformed email
      // here. Every row is kept, tagged with its validity, and sent through
      // to the backend so it shows up properly in the final failed count.
      const allRecipients = rows.map((row, idx) => {
        const obj = {};
        allKeys.forEach(k => { obj[k] = String(row[k] ?? '').trim(); });
        obj['Name']    = String(row[nameKey]  ?? '').trim();
        obj['Email']   = String(row[emailKey] ?? '').trim();
        obj['_row']    = idx + 2; // Excel row number (1 = header)
        obj['_valid']  = looksLikeEmail(obj['Email']);
        return obj;
      });

      const invalidCount = allRecipients.filter(r => !r['_valid']).length;
      if (invalidCount > 0) {
        console.warn(`[Excel] ${invalidCount} row(s) have a missing/invalid email — they will still be listed in results as "Failed".`);
      }

      parsedRecipients = allRecipients;

      if (!parsedRecipients.length) {
        showToast('❌ No rows found in the sheet.', 'error');
        return;
      }

      fileNameLabel.textContent = file.name;
      countPill.textContent = invalidCount > 0
        ? `${parsedRecipients.length} rows (${invalidCount} invalid email)`
        : `${parsedRecipients.length} recipient${parsedRecipients.length !== 1 ? 's' : ''}`;
      fileUploadRow.style.display = 'none';
      fileBadge.style.display     = 'flex';

      populateDropdown();

      showToast(
        invalidCount > 0
          ? `⚠️ Loaded ${parsedRecipients.length} rows from "${sheetName}" — ${invalidCount} have an invalid/missing email and will show as failed.`
          : `✅ Loaded ${parsedRecipients.length} recipients from "${sheetName}"`,
        invalidCount > 0 ? 'error' : 'success'
      );

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
    const flag = r['_valid'] ? '' : '  ⚠️ invalid email';
    opt.textContent = `${r.Name || '(no name)'}  —  ${r.Email || '(missing)'}${flag}`;
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
    // FIX: send ALL rows (including ones with an invalid email) so the
    // backend can report every row's outcome. Previously invalid rows were
    // filtered out here and simply vanished from the results.
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
        : `⚠️ ${data.successCount} sent, ${data.failureCount} failed (see details below).`,
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
  // FIX: total / success / failed are always shown and always add up
  // (total should equal successCount + failureCount every time now).
  const total   = data.total ?? 0;
  const success = data.successCount ?? 0;
  const failed  = data.failureCount ?? 0;

  document.getElementById('statTotal').textContent   = total;
  document.getElementById('statSuccess').textContent = success;
  document.getElementById('statFailed').textContent  = failed;

  if (total !== success + failed) {
    console.warn(`[Results] Mismatch: total=${total} but success(${success})+failed(${failed})=${success + failed}`);
  }

  const failedBlock = document.getElementById('failedBlock');
  const failedList  = document.getElementById('failedList');
  failedList.innerHTML = '';

  if (data.failed && data.failed.length) {
    data.failed.forEach(f => {
      const li = document.createElement('li');
      const rowPrefix = f.row ? `[${f.row}] ` : '';
      const namePart  = f.name ? `${f.name} — ` : '';
      li.textContent = `${rowPrefix}${namePart}${f.email} — ${f.error}`;
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
git remote -v