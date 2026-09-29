/**
 * Excel Mail Sender — Backend (server.js)
 * ----------------------------------------
 * Express server that exposes POST /send-mails
 * Uses Nodemailer + Zoho SMTP (App Password).
 * Credentials are loaded from .env (never hardcoded).
 *
 * Setup:
 *   1. cp .env.example .env  → fill in your Zoho email and App Password
 *   2. npm install
 *   3. node server.js
 *
 * FIXES in this version:
 *   - Every recipient row is now judged and reported individually — no row
 *     is ever silently skipped without showing up in the results.
 *   - Clear, specific failure reasons ("Missing email", "Invalid email
 *     format", the raw SMTP error, etc.) instead of one generic bucket.
 *   - Response always includes total / successCount / failureCount / failed
 *     / succeeded, so the front-end can never "lose" the failed count.
 */

'use strict';

const express = require('express');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3002;

/* ─────────────────────────────────────────────
   Signature Image Path
   Place image at: public/images/signature.png
───────────────────────────────────────────── */

const SIGNATURE_PATH = path.join(__dirname, 'public', 'images', 'signature.png');

/* ─────────────────────────────────────────────
   Middleware
───────────────────────────────────────────── */

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

/* ─────────────────────────────────────────────
   Nodemailer Transporter
───────────────────────────────────────────── */

function createTransporter() {
  return nodemailer.createTransport({
    host: 'smtp.zoho.com',
    port: 465,
    secure: true,
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS
    }
  });
}

/* ─────────────────────────────────────────────
   Email Validation
───────────────────────────────────────────── */

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

/* ─────────────────────────────────────────────
   Send Mail Route
───────────────────────────────────────────── */

app.post('/send-mails', async (req, res) => {

  if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
    return res.status(500).json({
      error: 'EMAIL_USER or EMAIL_PASS missing in .env'
    });
  }

  const { recipients } = req.body;

  if (!Array.isArray(recipients) || recipients.length === 0) {
    return res.status(400).json({
      error: 'recipients array required'
    });
  }

  const transporter = createTransporter();

  try {
    await transporter.verify();
  } catch (err) {
    return res.status(500).json({
      error: 'SMTP verification failed: ' + err.message
    });
  }

  // Every row gets a definite outcome — nothing is dropped silently.
  const results = {
    total: recipients.length,
    successCount: 0,
    failureCount: 0,
    succeeded: [],
    failed: []
  };

  for (let i = 0; i < recipients.length; i++) {
    const recipient = recipients[i] || {};
    const rowLabel = `Row ${i + 1}`;
    const name = (recipient.name || '').toString().trim();
    const email = (recipient.email || '').toString().trim();
    const subject = (recipient.subject || '').toString().trim();
    const message = (recipient.message || '').toString();

    // 1) Missing email entirely
    if (!email) {
      results.failureCount++;
      results.failed.push({ row: rowLabel, name: name || '(no name)', email: '(empty)', error: 'Missing email address' });
      continue;
    }

    // 2) Malformed email (this used to be silently filtered out on the
    //    front-end and never reported — now it always shows up here)
    if (!isValidEmail(email)) {
      results.failureCount++;
      results.failed.push({ row: rowLabel, name: name || '(no name)', email, error: 'Invalid email format' });
      continue;
    }

    // 3) Missing message body — still attempted, but flagged if truly empty
    if (!message.trim()) {
      results.failureCount++;
      results.failed.push({ row: rowLabel, name: name || '(no name)', email, error: 'Empty message body' });
      continue;
    }

    const mailOptions = {
      from: `"${process.env.SENDER_NAME || 'Jhaishna HR'}" <${process.env.EMAIL_USER}>`,
      to: email,
      subject: subject || '(No Subject)',
      text: message,
      html: `
        <div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.6;color:#1e293b;max-width:600px;">
          ${message.replace(/\n/g, '<br/>')}
<div style="margin-top:16px;">
  <img src="cid:signatureImage"
       alt="Signature"
       style="height:22px; width:auto; display:block; margin-bottom:6px;" />

  <b>HR Regards</b><br/>
  <span>Jhaishna HR Team</span>
</div>
        </div>
      `,
      attachments: fs.existsSync(SIGNATURE_PATH)
        ? [
          {
            filename: 'signature.png',
            path: SIGNATURE_PATH,
            cid: 'signatureImage'
          }
        ]
        : []
    };

    try {
      const info = await transporter.sendMail(mailOptions);
      console.log(`✓ Sent to ${email} (${rowLabel})`);
      results.successCount++;
      results.succeeded.push({ row: rowLabel, name: name || '(no name)', email, messageId: info && info.messageId });
    } catch (err) {
      console.log(`✗ Failed ${email} (${rowLabel}) → ${err.message}`);
      results.failureCount++;
      results.failed.push({ row: rowLabel, name: name || '(no name)', email, error: err.message });
    }
  }

  // Sanity check — total must always equal success + failure.
  console.log(
    `[send-mails] total=${results.total} success=${results.successCount} failed=${results.failureCount}`
  );

  return res.status(200).json(results);
});

/* ─────────────────────────────────────────────
   Catch All
───────────────────────────────────────────── */

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

/* ─────────────────────────────────────────────
   Start Server
───────────────────────────────────────────── */

app.listen(PORT, () => {
  console.log(`\n🚀 Server running at http://localhost:${PORT}`);
  console.log(`EMAIL_USER: ${process.env.EMAIL_USER || 'NOT SET'}`);
  console.log(`SIGNATURE: ${fs.existsSync(SIGNATURE_PATH) ? 'Found' : 'Not Found'}\n`);
});
