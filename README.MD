Mass Mail Sender for Job Applications
🚨 IMPORTANT NOTE: Use This Tool Wisely!
!!! DO NOT SPAM MAILS TO HRs. USE IT WISELY… !!!

This project is intended to help freshers and job seekers send personalized, professional emails to multiple HR contacts efficiently. Misuse of this tool (like spamming or bulk cold emailing without relevance) can hurt your reputation and even block your email account. Please use responsibly.

📌 Features
This script streamlines your job application process by offering the following capabilities:

Reads HR email IDs directly from an Excel (.xlsx) file.

Sends a professional, personalized email to every listed recipient.

Option to securely attach your resume (PDF) to each email.

Uses Gmail SMTP with a secure App Password for reliable sending.

⚙️ Requirements
1. Python
You need Python 3.9+ installed on your system.

⚠️ Recommendation: Python versions 3.10–3.12 are recommended for stability. Avoid unstable releases like Python 3.14 alpha.

2. Dependencies
Install the necessary Python library for reading Excel files:

pip install openpyxl

🔑 Gmail Setup – App Password
Since Google no longer allows simple authentication for third-party apps, you must generate an App Password for use in this script.

Watch a guide on how to generate a Gmail App Password (Search online for the latest instructions).

Use the generated 16-character password instead of your normal Gmail password in the script configuration.

📂 Data Source
Your list of HR contacts must be in an Excel file.

The required HR email list format is currently maintained in this Google Sheet: [Google Drive Link](https://docs.google.com/spreadsheets/d/18wOmINfhsbU4V4qgaexSYotE11mXf3TqJ01RboG9Gm8/edit?usp=sharing)

⚠️ Important: The content and format of the sheet may be updated from time to time. Please download the latest version as an .xlsx file before running the script.

🚀 How to Use
Follow these steps to successfully run the application:

Step 1: Prepare Files
Clone or download this project folder.

Place your resume file named resume.pdf in the same folder as the main script (e.g., script.py).

Ensure your HR email data is saved as an .xlsx file in the same directory.

Step 2: Configure Credentials
Open the main Python script (script.py) and update the following variables with your secure information:

YOUR_EMAIL = "yourmail@gmail.com"
YOUR_PASSWORD = "your_app_password_here" # Use the Gmail App Password here!

Step 3: Run the Script
Execute the script from your terminal:

python script.py
