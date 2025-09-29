# Mass Mail Sender for Job Applications  

🚨 **IMPORTANT NOTE**  
!!! DO NOT SPAM MAILS TO HRs, USE IT WISELY… !!!  
This project is intended to help freshers and job seekers send **personalized, professional emails** to multiple HR contacts efficiently. Misuse of this tool (like spamming or bulk cold emailing without relevance) can hurt your reputation and even block your email account. Please use responsibly.  

---

## ⚙️ Requirements  

1. **Python 3.9+** (recommended: 3.10–3.12 for stability)  
   > ⚠️ Python 3.14 (alpha) may cause issues, avoid unstable releases.  

2. Install dependencies:  
   ```bash
   pip install openpyxl


## 🔑 Gmail Setup – App Password  

Since Google no longer allows "less secure apps," you need an **App Password**:  

- Watch this guide: [How to generate Gmail App Password](https://www.youtube.com/watch?v=MkLX85XU5rU)  
- Use that generated password instead of your normal Gmail password in the script.  

---

## 📂 Data Source  

HR email list is maintained in this **Google Sheet**:  
[Google Drive Link](https://docs.google.com/spreadsheets/d/18wOmINfhsbU4V4qgaexSYotE11mXf3TqJ01RboG9Gm8/edit?usp=sharing)  

⚠️ **The sheet will be updated from time to time.** Please check for the latest version before running the script.  

Copy the emails in your exel file save it and drop it into the folder along with script and your resume in pdf format.  
**Exel file name- "hr_contacts.xlsx" (change in code according to your file name)
**Resume file name- "resume.pdf" (change in code according to your file name)
---

## 🚀 How to Use  

1. Clone/download this project.  
2. Place your **resume.pdf** in the same folder as the script.  
3. Update `script.py` with your credentials:  

   ```python
   YOUR_EMAIL = "yourmail@gmail.com"
   YOUR_PASSWORD = "your_app_password_here"
4. Run the python file
   ```python
   python script.py

