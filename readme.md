
# 📧 Group Mails & Compensation Letter Automation

## 📌 Project Overview
This project is a **Group Mail and Compensation Letter Automation System** developed for **Jhaishna Technologies Pvt. Ltd.**  
It enables HR to upload an Excel sheet with employee details and automatically generate, preview, and send personalized appraisal/compensation letters.

## 🚀 Key Features
- Upload employee Excel (.xlsx) file
- Auto-detect Name, Email, Gender, CTC details
- Gender-based salutation (Mr / Ms)
- Preview & edit letters before sending
- Send single or bulk emails
- Generate individual PDFs
- Download all PDFs as a ZIP file
- Success/error notifications and loading indicators

## 🛠️ Tech Stack
### Frontend
- HTML5
- CSS3
- JavaScript
- jsPDF
- html2canvas
- JSZip

### Backend
- Node.js
- Express.js
- Multer
- XLSX
- Nodemailer
- pdf-lib
- dotenv
- cors

## 📦 Dependencies Installation

Run the following command:
```bash
npm install express multer xlsx nodemailer pdf-lib cors dotenv
```

Optional (recommended for development):
```bash
npm install --save-dev nodemon
```

## 📄 Environment Variables
Create a `.env` file in the root directory:

```env
PORT=3000
EMAIL_USER=yourgmail@gmail.com
EMAIL_PASS=your_gmail_app_password
HR_REPLY_EMAIL=hr@jhaishna.com
```

> ⚠️ Use **Gmail App Password**, not your normal Gmail password.


# Install the nodemon 
# Nodemon is a development tool used in Node.js that automatically restarts the server whenever changes are made to the code. It helps save time and makes development faster.

 npm install -g nodemon
 nodemon -v

# In package.json write this after {},

"scripts": {
  "start": "node app.js",
  "dev": "nodemon app.js"
}



## ▶️ How to Run
'''
node server.js (or)

nodemon server.js

Open in browser:
```
http://localhost:3000
```

## 📂 Project Structure
```
project-root/
│
├── public/
│   ├── index.html
│   ├── style.css
│   ├── script.js
│   └── images/
│
├── uploads/
├── server.js
├── .env
├── package.json
└── README.md
```

## ✅ Conclusion
This application automates HR appraisal communication, reduces manual work, and ensures accurate and efficient group mailing with PDF generation.


https://hzlnz792-3001.inc1.devtunnels.ms/