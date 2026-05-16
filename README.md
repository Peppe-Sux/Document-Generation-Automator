# 📄 Document Generation Automator - Automatic Document Generator

## 📌 Overview

Document Generation Automator is a Python tool for the automatic generation of customized documents (Word and PDF) from:

- a Word template (.docx)
- an Excel file (.xlsx) with structured data

The system automatically replaces placeholders in the template with data from each row of the Excel file, generating individual documents in batch mode.

The goal is to automate repetitive processes such as:

- contracts
- invoices
- certificates
- personalized documents

---

## ⚙️ Key Features

### 1. Document generation from template

```bash
[nome]
{nome}
<nome>
{{nome}}
```

---

### 2. Secure placeholder replacement

The replacement occurs at the paragraph level to avoid the "run" fragmentation issue typical of Word.

This guarantees:

- document stability  
- formatting preservation  
- reliable replacements  

---

### 3. Data mapping from Excel

Each row of the Excel file is converted into a structured dictionary with:

- first name  
- last name  
- date (formatted)  
- amount (formatted as currency)  

---

### 4. Batch generation

A separate document is generated for each row of the dataset.

The files are:

- automatically named  
- saved in an output folder  

---

### 5. PDF export (headless)

Supports automatic DOCX → PDF conversion via LibreOffice in headless mode.

Compatible with:

- Windows  
- Linux  
- macOS  

If LibreOffice is unavailable, the system still saves the file in Word format as a fallback.

---

### 6. Dynamic file naming

Example:

```bash
Contratto_Mario_Rossi.pdf
```

---

## 📊 Automation Preview (Before & After)

Here is a practical example of how the pipeline takes a single row from the Excel spreadsheet, maps it into the Word template, and outputs a perfectly formatted, personalized document.

### 1. The Source Data (`data_source.xlsx`)
| first_name | last_name | date | amount |
| :--- | :--- | :--- | :--- |
| mario | rossi | 2026-05-15 | 1200.50 |
| luigi | verdi | 2026-05-16 | 1500.00 |

### 2. The Word Template (`template_braces_demo.docx`)
> ### 📄 DEMO COMMERCIAL AGREEMENT
> ---
> I, the undersigned **{first_name} {last_name}**, hereby declare this agreement valid on **{date}** for a total balance value of **{amount}**.

### 3. The Generated Output (`Output_Documents/Document_Mario_Rossi.pdf`)
> ### 📄 DEMO COMMERCIAL AGREEMENT
> ---
> I, the undersigned **Mario Rossi**, hereby declare this agreement valid on **15/05/2026** for a total balance value of **€ 1.200,50**.

### 💡 Smart Data Formatting Under the Hood
Notice how the raw, unformatted data from Excel was automatically transformed during the batch generation process to match professional document standards:
1. **Case Normalization:** `mario` and `rossi` were normalized to *Proper Case* (**Mario Rossi**) inside the document body.
2. **Temporal Formatting:** The ISO date `2026-05-15` was automatically converted into a human-readable local format (**15/05/2026**).
3. **Financial Formatting:** The raw float `1200.50` was dynamically cast into a properly localized currency string with the Euro symbol and standard decimals (**€ 1.200,50**).
4. **Dynamic Naming:** The pipeline combined the context variables to save the file with a clean, identifiable name (e.g., `Contratto_mario_rossi.pdf`).

---

## 🧠 Workflow

```bash
Load Excel file
Read Word template
Extract data row by row
Create data context
Replace placeholders in the document
Save temporary file
Convert to PDF (optional)
Final save in the output folder
```

---

## 📂 Input Requirements

### Excel File

```bash
nome
cognome
data
importo
```

---

### Word Template

```bash
[nome]
{nome}
<nome>
{{nome}}
```

---

## 📤 Output

For each row of the Excel file, the system generates:

- a Word or PDF document  
- automatically named  
- saved in the `/Output` folder  

---

## 🚀 Usage Example

```python
process_batch(
    template_file="template.docx",
    excel_file="dati.xlsx",
    force_pdf=True
)
```

---

## 🔧 Main Components

### format_context()

Converts an Excel row into a clean, formatted dictionary.

Handles:

- names  
- dates  
- amounts  

---

### replace_placeholders_safely()

Replaces placeholders in the Word document securely, avoiding text fragmentation issues.

---

### convert_docx_to_pdf()

Converts DOCX files to PDF using LibreOffice in headless mode.

Includes a fallback in case the software is unavailable.

---

### process_batch()

Main function that manages the entire workflow:

- Excel reading  
- document generation  
- data replacement  
- PDF/DOCX export  
- output management  

---

## 🧩 Technologies Used

```bash
pandas → Excel data management
python-docx → Word document manipulation
subprocess → external PDF conversion
LibreOffice → document rendering
os / shutil → file and directory management
```

---

## 📌 Real-World Use Cases

This system is useful for:

- automatic contract generation  
- personalized invoice issuance  
- certificate creation  
- HR documents  
- customized reports  
- corporate document automation  

---

## ⚖️ Technical Specifications

- batch processing architecture  
- document generation from template  
- cross-platform PDF support  
- robust fallback system  
- dynamic file name management  

---

## 🧱 Current Limitations

- requires structured Excel files  
- dependency on LibreOffice for PDF  
- template must comply with defined placeholders  
- no graphical interface (script only)  

---

## 🔄 Potential Extensions

```bash
web interface (Flask / FastAPI)
template editor
database integration
automatic email delivery of documents
advanced operation logging
JSON template configuration system
```

---

## 🎯 Project Goal

The goal is to automate the creation of repetitive documents by transforming structured data into ready-to-use files with minimal manual intervention.
