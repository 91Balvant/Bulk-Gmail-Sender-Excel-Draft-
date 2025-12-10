<div align="center">

# 📧 **Mail Merge Pro – Bulk Gmail Sender**

A modern desktop tool to send **personalized bulk emails** using  
**Excel data + Gmail Drafts + Google OAuth**.

<img src="https://komarev.com/ghpvc/?username=MailMergePro&label=VISITORS&color=blue&style=for-the-badge" />

<br><br>

<img src="https://github.com/91Balvant/Bulk-Gmail-Sender-Excel-Draft-/blob/main/preview.png?raw=true" width="720" />

</div>

---

# 👋 Welcome to Mail Merge Pro!
Send personalized bulk emails directly from your Gmail Drafts using Excel data.

---

# 🚀 Key Features

### 🔐 Google OAuth 2.0 Login  
Automatically authenticates your Google account — no manual token setup.

### 📊 Excel Integration  
Import `.xlsx` files with variables like **Name**, **Email**, **Company**, etc.

### 📝 Gmail Draft Templates  
Use Gmail Drafts as your email template with placeholders.

### 📎 Attachments  
Send attachments automatically — or conditionally by Excel rules.

### ▶️ Start / Stop / Resume  
Safely manage long sending queues.

### 📌 Real-Time Logs & Progress Bar  
Get live status while emails are being sent.

---

# 📘 Complete User Guide

## 1️⃣ Excel File Setup

Your Excel file **must include**:

| Column | Description |
|--------|-------------|
| **Email** | Required |
| **Name**, **Company**, etc. | Optional placeholders |
| **CC**, **BCC** | Optional (comma-separated) |
| **Send Attachments** | Optional (Yes/No) |

Example:

| Email | Name | Company | CC | Send Attachments |
|-------|------|---------|-----|------------------|
| john@xyz.com | John | XYZ Ltd | mark@abc.com | Yes |
| amy@abc.com | Amy | ABC Corp | | No |

---

## 2️⃣ Creating Gmail Draft Templates

Create a Draft inside Gmail and use placeholders that match Excel headers:
Hi {{Name}}, attached is the report for {{Company}}.

✨ Subject line can also use placeholders:


---

## 3️⃣ Sending Process

**Step 1:** Login with Google  
**Step 2:** Select your Excel file  
**Step 3:** Select your Gmail Draft  
**Step 4:** Click **Start New** to begin sending  

You can **Stop** anytime and **Resume** later.

---

## 4️⃣ Conditional Attachments

To send attachments only for specific users:

1. Uncheck **Send Attachments** in the app  
2. Add a column in Excel → `Send Attachments`  
3. Behavior:  
   - **Yes** (or empty): Send attachments  
   - **No**: Remove attachments for that user  

Great for campaigns with different requirements.

---

# 🛠️ Setup Instructions

## 1️⃣ Get `credentials.json`

1. Open **Google Cloud Console**  
2. Create a project  
3. Enable **Gmail API**  
4. Go to **Credentials → Create Credentials → OAuth Client ID**  
5. Select **Desktop App**  
6. Download JSON → rename to **credentials.json**

---

## 2️⃣ Install the Application

```bash
git clone https://github.com/91Balvant/Bulk-Gmail-Sender-Excel-Draft-.git
cd Bulk-Gmail-Sender-Excel-Draft-
pip install -r requirements.txt
