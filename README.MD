# Mass Mail Sender for Job Applications  

🚨 **IMPORTANT NOTE**  
!!! DO NOT SPAM MAILS TO HRs, USE IT WISELY… !!!  
This project is intended to help freshers and job seekers send **personalized, professional emails** to multiple HR contacts efficiently. Misuse of this tool (like spamming or bulk cold emailing without relevance) can hurt your reputation and even block your email account. Please use responsibly.  

---

## 📌 Features  
- Reads HR email IDs directly from an Excel (`.xlsx`) file.  
- Sends personalized emails to all listed recipients.  
- Option to attach your **resume (PDF)**.  
- Uses **Gmail SMTP** with App Password for secure login.  

---

## ⚙️ Requirements  

1. **Python 3.9+** (recommended: 3.10–3.12 for stability)  
   > ⚠️ Python 3.14 (alpha) may cause issues, avoid unstable releases.  

2. Install dependencies:  
   ```bash
   pip install openpyxl

🔑 Gmail Setup – App Password

Since Google no longer allows "less secure apps," you need an App Password:

Watch this guide: How to generate Gmail App Password

Use that generated password instead of your normal Gmail password in the script.

📂 Data Source

HR email list is maintained in this Google Sheet:
Google Drive Link

⚠️ The sheet will be updated from time to time. Please check for the latest version before running the script.

🚀 How to Use

Clone/download this project.

Place your resume.pdf in the same folder as the script.

Update script.py with:

YOUR_EMAIL = "yourmail@gmail.com"
YOUR_PASSWORD = "your_app_password_here"


Run the script:

python script.py
