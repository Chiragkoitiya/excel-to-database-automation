# excel-to-database-automation
📌 Excel to Database Automation (Python + MySQL + Tkinter)
📖 Overview

Excel to Database Automation is a real-world Python project that automates the process of:

Reading multiple monthly Excel billing files

Cleaning & validating the data

Removing duplicates

Merging all data into one consolidated dataset

Saving it into a MySQL database

Exporting a yearly summary Excel report

Providing an easy-to-use GUI application for non-technical users

This project was originally built for a jewelry shop’s billing workflow but works for any business that wants to automate Excel → Database operations.

🚀 Features
🔹 1. Automated Excel File Processing

Reads all .xlsx files from a selected folder

Combines & cleans them automatically

Removes duplicates and invalid rows

Handles missing values intelligently

🔹 2. GUI Application (Tkinter)

User-friendly, interactive interface

Folder selection

Database configuration screen

Data preview mode

Status logs

One-click automation

🔹 3. MySQL Database Integration

Auto-creates database & tables if not present

Inserts new records

Updates existing ones

Ensures Bill_No uniqueness

Stores clean, structured billing records

🔹 4. Yearly Excel Report Generator

Exports a consolidated yearly Excel file

Includes:

All transactions

Monthly summaries

Top customers

Revenue totals

🔹 5. Synthetic Dataset Generator

Creates 12 months of realistic jewelry billing data

Adds random missing values & duplicate rows

Useful for testing, demos, and training ML/ETL models

🛠 Tech Stack
Component	Technology
Language	Python
GUI	Tkinter
Data Processing	Pandas
Database	MySQL (mysql-connector-python)
File Handling	OpenPyXL
Scripting	Python (OOP-based classes)
📁 Project Structure
excel-to-database-automation/
│── app/
│   ├── billing_automation.py
│   ├── dataset_generator.py
│   └── config.pkl
│
│── sample_data/
│   └── monthly_billing_data/
│
│── requirements.txt
│── README.md
│── .gitignore

▶️ How to Run the Application
1️⃣ Install required dependencies
pip install -r requirements.txt

2️⃣ (Optional) Generate sample monthly Excel files
python app/dataset_generator.py


This creates realistic test data.

3️⃣ Run the GUI application
python app/billing_automation.py
