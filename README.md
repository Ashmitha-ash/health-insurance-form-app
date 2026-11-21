# 🛡️ Health Insurance Registration & Analytics System

A desktop GUI application built with **Python**, **Tkinter**, **OpenPyXL**, and **Matplotlib** for insurance applicant registration, Excel database management, and interactive data visualization.

---

## 📌 Project Overview

This project provides a complete workflow for registering insurance policy applicants, storing customer records in a structured Excel spreadsheet, and visualizing applicant distribution across insurance plans.

Developed as a Data Science mini-project under the guidance of **Dr. Sivalingan**, Head of the Department of Data Science at **Providence College for Women (Autonomous)**.

---

## 🚀 Key Features

- **🖥️ Desktop GUI Interface (Tkinter)**: Clean, user-friendly form designed with centered window layout, clear input fields, radio buttons, and dropdown menus.
- **📁 Automated Excel Database (OpenPyXL)**: Seamlessly creates, formats, and updates `health_insurance_data.xlsx` with custom column dimensions.
- **🔒 Input Validation & Data Integrity**:
  - **Phone Validation**: Ensures phone number contains exactly 10 digits.
  - **Email Format**: Validates email structure.
  - **DOB & Age Check**: Parses date of birth and ensures applicants are at least 18 years old.
  - **Mandatory Fields**: Confirms all details and terms agreement are provided before submission.
- **📊 Interactive Visual Analytics (Matplotlib)**:
  - **Pie Chart**: Visualizes market share across insurance products.
  - **Bar Chart**: Shows applicant counts per insurance category with value labels.
  - **Line Plot**: Visualizes applicant distribution trends.
- **⚡ Quick Actions**: One-click **Submit**, **Clear**, and **Exit** controls.

---

## 🛠️ Tech Stack

- **Programming Language**: Python 3.x
- **GUI Framework**: Tkinter (Standard Python GUI library)
- **Database / File Storage**: [OpenPyXL](https://openpyxl.readthedocs.io/) (Excel Workbook Engine)
- **Data Visualization**: [Matplotlib](https://matplotlib.org/)

---

## 📁 Repository Structure

```text
├── app.py                      # Main Tkinter application and logic
├── health_insurance_data.xlsx  # Excel workbook database (auto-generated if missing)
├── requirements.txt            # Python dependencies (openpyxl, matplotlib)
├── .gitignore                  # Git ignore rules for virtualenvs & cache
└── README.md                   # Project documentation
```

---

## 🧪 Installation & How to Run

### 1. Clone the repository
```bash
git clone https://github.com/Ashmitha-ash/health-insurance-form-app.git
cd health-insurance-form-app
```

### 2. (Optional) Create a virtual environment
```bash
python -m venv .venv
# On Windows:
.venv\Scripts\activate
# On macOS/Linux:
source .venv/bin/activate
```

### 3. Install required packages
```bash
pip install -r requirements.txt
```

### 4. Run the application
```bash
python app.py
```

---

## 👤 Author & Acknowledgments

- **Ashmitha A** – Department of Data Science, Providence College for Women (Autonomous), Coonoor, The Nilgiris.
- **Faculty Guide**: Dr. Sivalingan (MSc, MPhil, PhD), Head & Assistant Professor, Department of Data Science.
