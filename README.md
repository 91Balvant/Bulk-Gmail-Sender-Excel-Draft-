<!-- ============================= -->
<!-- 🌟 GLASSY EFFECT README THEME 🌟 -->
<!-- ============================= -->

<style>
/* -------- GLOBAL -------- */
body {
  background: linear-gradient(135deg, #0f172a, #1e293b, #0f172a);
  font-family: "Segoe UI", sans-serif;
  color: #e2e8f0;
  margin: 0;
  padding: 0;
}

/* Center Wrapper */
.glass-container {
  max-width: 900px;
  margin: 40px auto;
  padding: 40px;
  background: rgba(255, 255, 255, 0.08);
  backdrop-filter: blur(14px);
  border-radius: 18px;
  border: 1px solid rgba(255, 255, 255, 0.18);
  box-shadow: 0 8px 25px rgba(0, 0, 0, 0.45);
  animation: fadeIn 1.2s ease;
}

/* Section Headers */
h1, h2, h3 {
  color: #ffffff;
  text-shadow: 0 0 10px rgba(255, 255, 255, 0.25);
}

h1 {
  font-size: 2.6rem;
  letter-spacing: 1px;
  margin-top: 0;
}

/* Glow underline */
.glow-title {
  padding-bottom: 8px;
  border-bottom: 2px solid #38bdf8;
  box-shadow: 0 6px 15px -6px #38bdf8;
}

/* Feature cards */
.feature-box {
  background: rgba(255, 255, 255, 0.1);
  padding: 18px;
  margin: 14px 0;
  border-radius: 12px;
  border-left: 4px solid #38bdf8;
  transition: 0.3s;
}
.feature-box:hover {
  background: rgba(56, 189, 248, 0.18);
  transform: translateX(4px);
}

/* Tables */
table {
  width: 100%;
  border-collapse: collapse;
  background: rgba(255,255,255,0.07);
  backdrop-filter: blur(10px);
  border-radius: 10px;
  overflow: hidden;
}

th, td {
  padding: 12px;
  border-bottom: 1px solid rgba(255,255,255,0.12);
}

th {
  background: rgba(255,255,255,0.12);
  color: #fff;
}

/* Code Block */
pre {
  background: rgba(0,0,0,0.55);
  color: #00eaff;
  padding: 16px;
  border-radius: 10px;
  overflow: auto;
  border-left: 4px solid #00eaff;
}

/* Links */
a {
  color: #38bdf8;
  text-decoration: none;
  transition: 0.3s;
}
a:hover {
  color: #7dd3fc;
}

/* Fade animation */
@keyframes fadeIn {
  from { opacity: 0; transform: translateY(25px); }
  to { opacity: 1; transform: translateY(0); }
}

/* Center images */
.center {
  text-align: center;
}

img {
  border-radius: 10px;
  box-shadow: 0 0 20px rgba(56, 189, 248, 0.2);
}
</style>


<div class="glass-container">

<div class="center">
<h1 class="glow-title">📧 Mail Merge Pro – Bulk Gmail Sender</h1>

<p>A modern desktop tool to send <b>personalized bulk emails</b> using  
Excel data + Gmail Drafts + Google OAuth.</p>

<img src="https://komarev.com/ghpvc/?username=MailMergePro&label=VISITORS&color=blue&style=for-the-badge" />

<br><br>

<img src="https://github.com/91Balvant/Bulk-Gmail-Sender-Excel-Draft-/blob/main/preview.png?raw=true" width="720" />
</div>


<hr>


<h1>👋 Welcome to Mail Merge Pro!</h1>
<p>Send personalized bulk emails directly from your Gmail Drafts using Excel data.</p>


<hr>


<h1>🚀 Key Features</h1>

<div class="feature-box">🔐 <b>Google OAuth 2.0 Login</b><br>Automatically authenticates your Google account — no manual token setup.</div>

<div class="feature-box">📊 <b>Excel Integration</b><br>Import <code>.xlsx</code> files with variables like Name, Email, Company, etc.</div>

<div class="feature-box">📝 <b>Gmail Draft Templates</b><br>Use Gmail Drafts as your email template with placeholders.</div>

<div class="feature-box">📎 <b>Attachments</b><br>Send attachments automatically — or conditionally by Excel rules.</div>

<div class="feature-box">▶️ <b>Start / Stop / Resume</b><br>Safely manage long sending queues.</div>

<div class="feature-box">📌 <b>Real-Time Logs & Progress Bar</b><br>Track sending live.</div>


<hr>


<h1>📘 Complete User Guide</h1>

<h2>1️⃣ Excel File Setup</h2>

<p>Your Excel file <b>must include</b>:</p>

<table>
<tr><th>Column</th><th>Description</th></tr>
<tr><td>Email</td><td>Required</td></tr>
<tr><td>Name, Company, etc.</td><td>Optional placeholders</td></tr>
<tr><td>CC, BCC</td><td>Optional (comma-separated)</td></tr>
<tr><td>Send Attachments</td><td>Optional (Yes/No)</td></tr>
</table>

<br>

<p><b>Example:</b></p>

<table>
<tr><th>Email</th><th>Name</th><th>Company</th><th>CC</th><th>Send Attachments</th></tr>
<tr><td>john@xyz.com</td><td>John</td><td>XYZ Ltd</td><td>mark@abc.com</td><td>Yes</td></tr>
<tr><td>amy@abc.com</td><td>Amy</td><td>ABC Corp</td><td></td><td>No</td></tr>
</table>


<hr>


<h2>2️⃣ Creating Gmail Draft Templates</h2>
<p>Create a Draft inside Gmail and use placeholders that match Excel headers:</p>

<pre>Hi {{Name}}, attached is the report for {{Company}}.</pre>

<p>✨ Subject lines also support placeholders.</p>


<hr>


<h2>3️⃣ Sending Process</h2>

<p><b>Step 1:</b> Login with Google<br>
<b>Step 2:</b> Select your Excel file<br>
<b>Step 3:</b> Select your Gmail Draft<br>
<b>Step 4:</b> Click <b>Start New</b> to begin sending</p>

<p>You can <b>Stop</b> anytime and <b>Resume</b> later.</p>


<hr>


<h2>4️⃣ Conditional Attachments</h2>

<p>To send attachments only for specific users:</p>

<ol>
<li>Uncheck <b>Send Attachments</b> in the app</li>
<li>Add a column in Excel → <code>Send Attachments</code></li>
<li><b>Behavior:</b><br>
&nbsp;&nbsp;• Yes (or empty): Send attachments<br>
&nbsp;&nbsp;• No: Remove attachments for that user</li>
</ol>

<p>Great for campaigns with different requirements.</p>


<hr>


<h1>🛠️ Setup Instructions</h1>

<h2>1️⃣ Get <code>credentials.json</code></h2>

<ol>
<li>Open <b>Google Cloud Console</b></li>
<li>Create a project</li>
<li>Enable <b>Gmail API</b></li>
<li>Go to <b>Credentials → Create Credentials → OAuth Client ID</b></li>
<li>Select <b>Desktop App</b></li>
<li>Download JSON → rename to <code>credentials.json</code></li>
</ol>


<hr>


<h2>2️⃣ Install the Application</h2>

<pre>
git clone https://github.com/91Balvant/Bulk-Gmail-Sender-Excel-Draft-.git
cd Bulk-Gmail-Sender-Excel-Draft-
pip install -r requirements.txt
</pre>

</div>
