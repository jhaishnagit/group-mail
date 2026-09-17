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
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email));
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

  const results = {
    total: recipients.length,
    successCount: 0,
    failureCount: 0,
    failed: []
  };

  for (const recipient of recipients) {

    const { name, email, subject, message } = recipient;

    if (!email || !isValidEmail(email)) {
      results.failureCount++;
      results.failed.push({ email: email || '(empty)', error: 'Invalid email' });
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
      await transporter.sendMail(mailOptions);
      console.log(`✓ Sent to ${email}`);
      results.successCount++;
    } catch (err) {
      console.log(`✗ Failed ${email} → ${err.message}`);
      results.failureCount++;
      results.failed.push({ email, error: err.message });
    }
  }

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