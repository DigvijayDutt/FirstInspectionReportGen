````markdown
# 🧾 Insurance Claim Report Generator

This repository contains two Python scripts that automate the generation of **Word documents** from insurance claim data stored in an **Excel file**. It uses `pandas`, `openpyxl`, and `python-docx` to read structured data and generate styled `.docx` reports, complete with headings, photographs, and company branding.

---

## 📂 Repository Contents

### 1. `datagen.py`
Generates an Excel file (`data.xlsx`) with pre-filled sample insurance claim data.

- Uses `openpyxl` to write structured records to a sheet named `"Claims"`.
- Each record contains:
  - Insured name, address, insurer, adjuster
  - Dates (loss, inspection, report)
  - Loss type, cause, scope of work

### 2. `main.py`
Reads data from the Excel sheet and generates Word reports (`new0.docx`, `new1.docx`, etc.).

- Uses `pandas` and `python-docx`
- Each document includes:
  - A logo and heading
  - Claim data sections (title, address, type, etc.)
  - A photo section categorized by room (Living Room, Kitchen, etc.)

---

## 🛠 Dependencies

Install the following Python libraries:

```bash
pip install pandas openpyxl python-docx
```
````

---

## 🗂 Folder Structure

```
.
├── data.xlsx                   # Auto-generated Excel file
├── logo.jpg                    # Logo image to be placed in Word header
├── images/
│   ├── home.jpg                # Image used in all reports
│   ├── LIVING ROOM/
│   │   ├── 1.jpg
│   │   └── ... (up to 4.jpg)
│   ├── BEDROOM/
│   ├── KITCHEN/
│   └── STORAGE/
├── datagen.py            # Script to generate Excel data
├── main.py   # Script to generate Word reports
└── README.md                  # This file
```

---

## ▶️ How to Use

### Step 1: Generate Sample Claim Data

```bash
python datagen.py
```

### Step 2: Generate Word Reports from Excel

Ensure all images and `logo.jpg` exist as described, then run:

```bash
python main.py
```

Each `.docx` file will be saved in the current directory.

---

## 📌 Notes

* This project uses `.xlsx` format for compatibility with `openpyxl`.
* Update the image folders and logo as per your organization’s assets.
* Ensure room folders contain **at least 4 images** for the generator to work without error (or handle exceptions).

---


## 👨‍💻 Author

Built by \Digvijay Dutt – Python developer & automation enthusiast.

Feel free to contribute or raise issues!


