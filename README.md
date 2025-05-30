# 📊 Streamlit Excel Comparator Tool

This tool helps automate Excel-based analysis by comparing inventory rows with reference data using **fuzzy matching and block extraction**.

---

## 🔍 What It Does

- Upload a **base Excel file** (your inventory or analysis sheet)
- Upload a **source Excel file** (reference blocks in multiple sheets)
- The tool:
  - Fuzzy matches each row’s sheet name with actual sheet in source
  - Finds the **weight group**
  - Extracts a 12x8 data block from the matched sheet
  - Pastes it into the base file
  - Adds conditional formatting and formulas to compare values

---

## 🛠️ Tech Stack

| Tool         | Purpose                     |
|--------------|-----------------------------|
| Python       | Core logic                  |
| Streamlit    | Web UI                      |
| pandas       | DataFrame reading (optional)|
| openpyxl     | Excel manipulation          |
| fuzzywuzzy   | Smart sheet name matching   |

---

## ⚙️ How to Run

1. Clone the repo:
   ```bash
   git clone https://github.com/yourusername/streamlit-excel-comparator.git
   cd streamlit-excel-comparator
   ```
2. Install dependencies:
    ```bash
    pip install -r requirements.txt
    ```
3. Run the Streamlit app:
    ```bash
    streamlit run streamlit_app.py
    ```

---
## 📂 Excel Format (Expected)
  
  Column	Meaning
  B	Shape
  F	Weight group
  M	Sheet name (to match)
  
  Output block will be inserted starting from Column N
  
  Formatting and formulas will appear in row 13 (below each block)

---

## 🪪 License
MIT License
