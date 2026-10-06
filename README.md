# Certificate Automation System (PPTX & PDF)

An automated certificate generation workflow built with Python, designed for standalone batch processing and seamless integration with **n8n** automation workflows.

It populates recipient names from **CSV** or **Excel (.xlsx)** into a PowerPoint (`.pptx`) template while **strictly preserving original typography, font sizes, colors, and formatting**, with built-in export to **PDF**.

---

## Features

- **Batch Generation**: Process entire lists of names from CSV or Excel in a single command.
- **Style & Font Preservation**: Unlike standard text replacement, this replaces text at the run level, preserving font family (`Times New Roman MT`), size (`13.56pt`), weight (`Bold`), alignment, and color (`#2B232E`).
- **Dual Format Export**: Generates `.pptx` presentations and automatically converts them to `.pdf`.
- **Cross-Environment Ready**:
  - **Local Windows**: Automatically utilizes Microsoft PowerPoint COM automation for high-fidelity PDF export.
  - **Linux / Docker (n8n)**: Automatically resolves paths (e.g. `/files/...`) and supports headless LibreOffice conversion.
- **Flexible CLI**: Generate for a single student, a comma-separated list, or entire spreadsheet datasets.

---

## Project Structure

```text
cer/
├── input/
│   └── name.xlsx                   # Excel source of recipient names
├── name.csv                        # CSV source of recipient names
├── template/
│   └── Certificate_template.pptx   # PowerPoint template containing {{NAME}} placeholder
├── output/
│   └── pptx/                       # Generated PowerPoint certificates
├── ouput_pdf/                      # Generated PDF certificates
├── scripts/
│   ├── generate_certificate.py     # Core certificate generator (single or multi-name)
│   ├── batch_generate.py           # Batch orchestrator (reads CSV/XLSX)
│   └── convert_to_pdf.py           # Cross-platform PPTX to PDF converter
├── run_batch.py                    # Root convenience launcher
├── requirements.txt                # Python dependencies
├── .gitignore                      # Git ignore rules for outputs and caches
└── README.md                       # Project documentation
```

---

## Installation

### 1. Clone or Open the Repository
```bash
cd d:/cer
```

### 2. Install Python Dependencies
```bash
pip install -r requirements.txt
```

### 3. PDF Conversion Requirements
- **Windows**: Microsoft PowerPoint is already supported out of the box.
- **Linux / Docker (n8n)**: Install LibreOffice for headless conversion:
  ```bash
  apt-get update && apt-get install -y libreoffice
  ```

---

## How to Use

### 1. Quick Batch Generation (All Names)
Automatically detects `name.csv` or `input/name.xlsx` and creates certificates for everyone in the list:

```bash
# Generate PPTX only:
python run_batch.py

# Generate both PPTX and PDF:
python run_batch.py --pdf
```

### 2. Generate from a Specific File
Specify a custom CSV or Excel file with `--file` (or `-f`):

```bash
# From CSV:
python scripts/batch_generate.py --file name.csv --pdf

# From Excel:
python scripts/batch_generate.py --file input/name.xlsx --pdf
```

### 3. Generate for Single / Multiple Names Directly
Use `scripts/generate_certificate.py` for targeted generation:

```bash
# Single recipient:
python scripts/generate_certificate.py "John Doe" --pdf

# Multiple recipients:
python scripts/generate_certificate.py "Alice" "Bob" "Charlie" --pdf

# Comma-separated list:
python scripts/generate_certificate.py --names "Alice, Bob, Charlie" --pdf
```

### 4. Convert Existing PPTX Files to PDF
If you already have `.pptx` certificates inside `output/pptx/` and want to convert them into PDF:

```bash
# Convert all files in output/pptx/ to ouput_pdf/
python scripts/convert_to_pdf.py

# Convert a specific file:
python scripts/convert_to_pdf.py output/pptx/abi.pptx
```

---

## n8n Integration Guide

This repository is optimized for **n8n** automation nodes (running either locally or inside Docker).

### Scenario A: Batch Execution in n8n
To run the generation in a single step inside an **Execute Command** node:

```bash
python /files/scripts/batch_generate.py --file /files/name.csv --pdf
```

### Scenario B: Per-Item Loop in n8n
If your n8n workflow fetches items one-by-one (e.g. from Google Sheets, Airtable, Typeform, or Webhooks):

Configure the **Execute Command** node with:
```bash
python /files/scripts/generate_certificate.py "{{ $json.name }}" --pdf
```

### Docker Volume Mounting Note
Ensure your n8n Docker container mounts the project folder to `/files`:
```yaml
volumes:
  - d:/cer:/files
```
The scripts automatically detect the `/files` path inside Docker containers and fallback to local paths when executed on Windows.

---

## Template Customization

To customize the certificate template:
1. Open [template/Certificate_template.pptx](template/Certificate_template.pptx) in PowerPoint.
2. Edit graphics, fonts, dates, signatures, or college headers.
3. Place `{{NAME}}` in any text box where the recipient's name should appear.
4. Save the file. The scripts will automatically populate the name while maintaining the exact font, size, and color defined in that text box.
