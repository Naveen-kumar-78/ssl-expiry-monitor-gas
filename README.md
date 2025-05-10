# SSL Certificate Expiry Monitor 📆🔐

This Google Apps Script automates the monitoring of SSL certificate expiry dates from a Google Sheet and sends email alerts for certificates expiring soon. It helps teams proactively renew certificates before they expire, preventing downtime or security warnings.

---

## 🚀 Features

- Auto-calculates days remaining for SSL certificate expiry.
- Color-coded visual status indicators (Active, Warning, Expiring Soon, Expired).
- Sends automated email alerts to defined recipients when certificates are nearing expiry.
- Prevents multiple alert emails per day using timestamp tracking.
- Easily extendable and customizable.

---

## 📊 Sheet Structure

| Column | Name              | Description                              |
|--------|-------------------|------------------------------------------|
| B      | Project           | Project Name                              |
| C      | Website URL       | Domain or Website                        |
| G      | SSL Expires On    | Expiry Date (manually filled or fetched) |
| H      | Days Left         | Auto-filled by script                    |
| I      | Status            | Auto-filled by script                    |

---

## 🛠️ Setup Instructions

1. Open a new or existing Google Sheet.
2. Go to **Extensions > Apps Script**.
3. Copy and paste the code from `Code.gs` into the script editor.
4. Update the email addresses in the `sendExpiryAlert()` function.
5. Save and authorize permissions.
6. Set a time-based trigger to run `updateSSLCertExpiry()` daily.

---

## ✉️ Email Alert Example

```text
Subject: ⚠️ SSL Expiry Alert – Project A, Project B – Action Required

Dear Team,

The following SSL certificates are expiring soon:

🔹 Project: Project A  
🔹 Website URL: https://example.com  
🔹 Expires In: 12 days  
🔹 Expiry Date: Wed May 22 2025

📌 Click here to view the full SSL certificate report:
➡️ https://docs.google.com/spreadsheets/d/...

Best Regards,  
SSL Monitoring Team
