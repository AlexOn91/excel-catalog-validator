# Excel Catalog Validator

A complete desktop application for validating Excel product catalogs, built in **Python** with **Tkinter**, **Pandas**, and **OpenPyXL**.  
The app allows users to upload Excel files, run multiple data integrity checks, and generate structured error reports.

---

## ✨ Features
- **User Interface (Tkinter)** for file selection and sheet mapping.
- **Validation checks**:
  - Mandatory field completeness.
  - Uniqueness on key columns (e.g., ProductID, SKU).
  - Detection of hidden rows and columns.
  - Invalid characters, Excel formulas, HTML tags.
  - URL format validation (text and hyperlinks).
- **Error Reporting**:
  - Export structured fail reports into a separate Excel file.
  - Errors grouped by validation rule.
  - Includes cell references for easier review.
- **Performance**: tested with catalogs up to **65,000+ rows**.
- **Packaged executable**: built with **PyInstaller** (Windows `.exe`).

---

## 📂 Project Structure
excel-catalog-validator/
│── data/
│ └── products_demo.xlsx # Demo Excel file with fake product data
│── downloadfailreport.py # Fail report generator
│── main.py # Entry point / launcher
│── offline_app.py # Tkinter GUI
│── validator.py # Validation logic
│── README.md # Project documentation
│── requirements.txt # Dependencies
│── LICENSE # MIT License

## 🛠 Tech Stack
- Python 3.10+
- Tkinter (GUI)
- Pandas (data validation & processing)
- NumPy (data handling, array operations)
- OpenPyXL (Excel file integration)
- PyInstaller (packaging into Windows executable)
