# Paytime-USB-Python-Code
💡 One-Liner (Short Version) A Python-based attendance report automation tool that transforms biometric .TXT logs into monthly Excel and CSV reports — employee-wise, date-wise, and status-wise.
# 🕒 Attendance Report Generator
> Built & maintained by [@maulin18203](https://github.com/maulin18203)

A powerful Python tool to process raw biometric `.TXT` attendance logs and generate well-formatted Excel/CSV reports — employee-wise, day-wise, and status-wise.

---

## 📦 Features

- ✅ Parses Paytime biometric `.TXT` logs (supports various encodings)
- 🧼 Cleans & formats raw logs
- 🧠 In-Time, Out-Time, Status tracking for every employee/day
- 📊 Summary reports with total days, presents, absents, %
- 🧮 Interactive or command-line month selection
- 🪵 Detailed logging for debugging
- 📁 Exports Excel & CSV files per month

---

## 📥 How to Get the `.TXT` File from Your Paytime Biometric Device

### 📌 Option 1: Using USB Drive (Most Common)
1. Format a USB drive to **FAT32**
2. Plug into the biometric machine
3. Navigate: `Menu → USB Management → Download Attendance Logs`
4. Select date range if needed
5. Files exported:
   - ✅ `AGL_0001.TXT` — attendance logs
   - Optional: `ENROLLDB.DAT`, `DEVICEINFO.TXT`
6. Remove USB → plug into your PC → copy `.TXT` to your project folder

### 💻 Option 2: Using Paytime/ZK Software
1. Open Paytime Bio Attendance software
2. Select connected device (via LAN/IP)
3. Go to **Download Logs**
4. Export to `.TXT` format

---

## 🧑‍💻 How to Use This Script

### ✅ 1. Git Clone the Project

```bash
git clone https://github.com/maulin18203/attendance-report-generator.git
cd attendance-report-generator
✅ 2. Install Python (if not already installed)
Download Python:
👉 https://www.python.org/downloads/

During installation:
✅ Make sure to check the box that says "Add Python to PATH" before clicking install.

After installation, verify by opening a terminal or command prompt:

bash
Copy
Edit
python --version
You should see something like:

bash
Copy
Edit
Python 3.10.12
✅ 3. Install Required Libraries
Inside the cloned project folder, run:

bash
Copy
Edit
pip install -r requirements.txt
This will install:

pandas

openpyxl

⚠️ If pip doesn’t work, try:
python -m pip install -r requirements.txt

✅ 4. Run the Python Script
▶️ Option A: Interactive Month Selection
bash
Copy
Edit
python code.py --file AGL_0001.TXT --interactive
You’ll get a terminal menu to select specific months.

▶️ Option B: Auto-run All Months
bash
Copy
Edit
python code.py --file AGL_0001.TXT --months all
Processes all months in the file without prompts.

▶️ Option C: Specific Months Only
bash
Copy
Edit
python code.py --file AGL_0001.TXT --months 2023-10,2023-11
Replace the dates with the months you want, in YYYY-MM format.

📁 Output Files
📘 attendance_report.xlsx — one sheet per month + summary

📄 CSVs like report_2023-10.csv, summary_2023-10.csv

📂 Output folder created with a timestamp

🪵 Logs saved to attendance_processor.log

❗ Troubleshooting
Problem	Solution
ModuleNotFoundError	Run pip install -r requirements.txt again
python not recognized	Use python3 or check that Python is added to PATH
Garbled or missing data	Re-export fresh .TXT file from biometric machine
Date/Time errors	Ensure your .TXT file has valid DateTime entries
