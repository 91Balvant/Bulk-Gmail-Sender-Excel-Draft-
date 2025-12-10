# 📧 Mail Merge Pro - Bulk Gmail Sender

![Mail Merge Pro Preview](https://github.com/91Balvant/Bulk-Gmail-Sender-Excel-Draft-/blob/main/preview.png?raw=true)

**Mail Merge Pro** is a robust desktop application designed to automate bulk email sending directly through your Gmail account. By combining the power of Excel for data management and Gmail Drafts for template creation, this tool makes sending personalized mass emails effortless.

## 🚀 Key Features

* **🔒 Secure Google Authentication:** Log in safely using your own Google account credentials (OAuth 2.0).
* **📄 Excel Integration:** Import recipient lists and custom data variables directly from `.xlsx` files.
* **📝 Gmail Draft Support:** No need to code HTML templates! Design your email in Gmail (including formatting and signatures), save it as a draft, and select it within the app.
* **📎 Attachment Support:** Option to include attachments with your bulk emails.
* **⏯️ Control Flow:** Includes **Start**, **Stop**, and **Resume** functionality to manage large sending queues.
* **📊 Real-Time Status:** Live progress bar and activity logs to track sending status, authentication, and errors.
* **🔄 Auto-Refresh:** Capabilities to refresh Excel data and Draft lists without restarting the app.

---

## 🛠️ How It Works

1.  **Authentication:** The app authenticates with the Gmail API to access your drafts and sending capabilities.
2.  **Data Loading:** You upload an Excel file containing your recipient list (e.g., Email, Name, Company).
3.  **Template Selection:** The app fetches your current Gmail Drafts. Select the one you want to use as a template.
4.  **Mail Merge:** The app replaces placeholders in your draft with data from the Excel file (if configured) and sends the emails one by one.

---

## 💻 Installation & Setup

### Prerequisites
* Python 3.x installed.
* A Google Cloud Project with the Gmail API enabled.
* `credentials.json` file from Google Cloud Console.

### Steps

1.  **Clone the Repository**
    ```bash
    git clone [https://github.com/91Balvant/Bulk-Gmail-Sender-Excel-Draft-.git](https://github.com/91Balvant/Bulk-Gmail-Sender-Excel-Draft-.git)
    cd Bulk-Gmail-Sender-Excel-Draft-
    ```

2.  **Install Dependencies**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Setup Google Credentials**
    * Place your `credentials.json` file in the root directory of the project.
    * *Note: On the first run, a browser window will open asking you to authorize the application.*

4.  **Run the Application**
    ```bash
    python main.py
    ```
    *(Note: Replace `main.py` with the actual name of your entry script if different).*

---

## 📖 Usage Guide

1.  **Sign In:** Click the **Sign Out/In** button to authenticate with your Google Account.
2.  **Load Data:** Click **Select Excel File** to choose your recipient list.
3.  **Choose Draft:** The dropdown will populate with your Gmail Drafts. Select the draft you wish to send.
4.  **Configure:** Check **Send Attachments** if required.
5.  **Launch:** Click **Start New** to begin the mailing process.
6.  **Monitor:** Watch the logs and progress bar. You can **Pause** or **Stop** the process at any time.

---

## 📸 Screenshots

| Login & Draft Selection | Sending Process |
|:---:|:---:|
| *(You can add more screenshots here)* | *(You can add more screenshots here)* |

---

## ⚠️ Important Notes

* **Daily Limits:** Be aware of Gmail's daily sending limits (usually 500 emails/day for free accounts, 2000/day for Workspace).
* **Draft ID:** The application automatically detects the unique Draft ID (e.g., `r-4320...`) as shown in the interface.

---

## 📜 License & Copyright

**© 2025 Balvant Sharma. All Rights Reserved.**

This project is intended for personal or educational use. Please respect anti-spam laws and regulations when using bulk email tools.

---

## 👨‍💻 Author

**Balvant Sharma**
* [GitHub Profile](https://github.com/91Balvant)
