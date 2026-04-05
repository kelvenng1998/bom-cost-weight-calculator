# bom-cost-weight-calculator
# BOM Cost & Weight Calculator

## 📌 Overview
This project is a Python-based automation tool designed to process BOM (Bill of Materials) Excel files and generate:

- Cost estimation
- Weight calculation
- Material optimization (bar nesting)

It was built to reduce manual engineering workload and improve consistency in manufacturing workflows.

---

## ⚙️ Features

- ✅ Automated BOM processing
- 💰 Cost calculation based on material database
- ⚖️ Weight calculation with unit handling
- 📏 Bar nesting optimization (minimize material waste)
- 📂 Batch processing of multiple BOM files
- ⚠️ Missing data detection (cost / weight / files)

---

## 🧱 Project Structure

```
├── main.py                # Main script to process BOMs
├── requirements.txt       # List dependencies (pandas, openpyxl)
│
├── sample_data/           # Input data folder
│   ├── input.xlsx         # Defines BOM files to process
│   ├── raw_material_data.xlsx # Material database
│   └── database/          # Individual BOM files
│       ├── sample_part_a.xlsx
│       └── sample_part_b.xlsx
│
└── output/                # Generated results folder


---

## ▶️ How to Run

### 1. Install dependencies

pip install -r requirements.txt

### 2. Run the program

python main.py


---

## 📊 Example Workflow

input.xlsx → database BOM files → processing → output reports


---

## 📁 Output

The system generates:

- Bar nesting reports (Excel)
- Weight summary
- Cost summary
- Missing data reports (if any)

---

## 💡 Why This Project

Manual BOM processing is:
- repetitive
- error-prone
- time-consuming

This tool automates the process and can significantly reduce workload while improving accuracy.

---

## 📈 Impact

This tool automates BOM processing workflows and can significantly reduce manual engineering effort, improve consistency, and minimize material waste.

## ⚠️ Notes

- All data in this repository is for demonstration purposes only
- Built to replace repetitive manual BOM calculations with an automated, scalable solution.

---

## 🚀 Future Improvements

- GUI interface
- Database integration (SQL)
- API integration
- Performance optimization

---

## 👤 Author

Kel Ven Ng  
