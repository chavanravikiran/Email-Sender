# Email-Sender
Email-Sender

# Bulk Email Sender - Spring Boot Application

This Spring Boot application automatically reads an Excel file and sends bulk emails with a resume attachment.

## ✅ Features

- Reads company data from Excel (`.xlsx`)
- Sends personalized emails using `JavaMailSender`
- Attaches your resume automatically
- Supports multiple emails per company (newline-separated in Excel)

---

## 📂 Excel Format

| Company | Location | District | Email                          | Phone     |
|---------|----------|----------|--------------------------------|-----------|
| TCS     | Baner    | Pune     | example1@tcs.com               | 1234567890 |
| Infosys | Hinjewadi| Pune     | hr@infosys.com<br>careers@in.com | 9876543210 |

> 📌 *Multiple email IDs in the `Email` column must be separated by pressing `Alt + Enter` (newline).*

---
