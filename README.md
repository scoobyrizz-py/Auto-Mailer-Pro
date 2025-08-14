  <p align="center">
    <img src="Logo.png" alt="Auto Mailer Pro Logo" width="300">
  </p>

  # Auto Mailer Pro v6.0

  **Automate personalized insurance mail campaigns with ease.**

  Auto Mailer Pro is a Python-based tool developed for Jones Insurance Advisors, Inc. to generate personalized letters, envelopes, mailing labels, and CRM-ready CSV files for personal and commercial property owners. Featuring a user-friendly GUI, it supports targeted mail campaigns for owner-occupied homes (personal lines) or businesses (commercial lines) based on sales data, with customizable letter content and subject lines.

  **Version**: 6.0  
  **Author**: Kyle Padilla  
  **Title**: Insurance Agent, Producer, Developer   
  **Company**: Jones Insurance Advisors, Inc.  
  **Contact**: scooby_rizz@proton.me  
  **Last Updated**: August 13, 2025

  ## Features

  - **Dual Modes**:
    - **Personal Lines**: Targets owner-occupied properties using `Owner Name` and address matching.
    - **Commercial Lines**: Targets businesses with valid `Business Type` (e.g., Retail, Office).
  - **Graphical User Interface (GUI)**:
    - Select mode (Personal or Commercial) via radio buttons.
    - Upload Excel sales data files with a file picker.
    - Pre-filled, editable subject line that updates with mode selection.
    - Toggle default/custom letter content with a checkbox.
    - Large, scrollable text areas for letter content and output logs.
    - Brand logo at the top and credits at the bottom.
    - Larger window (1000x800) for better visibility.
  - **Output Generation**:
    - Personalized letters (`all_letters.docx`) with `[Name]` and `[County]` placeholders.
    - Envelopes (`all_envelopes.docx`) formatted for printing.
    - Mailing labels (`mailing_labels.docx`) in a 3-column layout.
    - CRM-ready CSV (`crm_personal_occupied.csv` or `crm_commercial_occupied.csv`).
  - **Customization**:
    - Custom subject line (bolded in letters).
    - Default templates for personal/commercial modes.
    - Optional signature image (`signature_brian.png`) in letters.
  - **Data Processing**:
    - Cleans names (e.g., `John Smith||Jane Smith` to `John & Jane Smith`).
    - Maps ZIP codes to city/state using `zip_lookup.csv`.
    - Filters non-owner-occupied properties or invalid business types.

  ## Prerequisites

  - **Python**: Version 3.13 or higher (for source code; not needed for executable).
  - **Operating System**: Tested on Windows (compatible with macOS/Linux with path adjustments).
  - **Dependencies** (for source code):
    ```bash
    pip install pandas python-docx fuzzywuzzy python-Levenshtein
    ```
  - **Optional** (for JPEG logo support or icon conversion):
    ```bash
    pip install Pillow
    ```

  ## Installation

  ### Setup (for Source Code)
  1. Clone or download the repository to a local directory (e.g., `C:\Users\Your

<!-- 
# Auto Mailer Pro
**Author:** Kyle Padilla  
**Contact:** scooby_rizz@protonmail.com  
**GitHub:** [scoobyrizz-py](https://github.com/scoobyrizz-py)
---


## 📚 Table of Contents
- [Overview](#overview)
- [Credits and Disclosure](#credits-and-disclosure)
- [Requirements](#requirements)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Files and Dependencies](#files-and-dependencies)
- [Notes](#notes)

---

## 📖 Overview
This Python script processes sales data from an Excel file to generate personalized letters, envelopes, and mailing labels for owner-occupied properties.

It is designed for businesses to automate outreach to homeowners, using public data such as:

- Property addresses
- Sale dates
- Sale prices

**Key Features:**
- Cleans and formats owner names.
- Checks if a property is owner-occupied using fuzzy string matching.
- Generates:
  - Letters (`.docx`)
  - Envelopes (`.docx`)
  - Mailing Labels (`.docx`)
- Exports owner-occupied data to a CRM-compatible CSV file.

---

## 🤝 Credits and Disclosure
This script was developed by **Kyle Padilla** with assistance from AI tools such as **Grok (xAI)** and **ChatGPT (Model GPT-5)** for code structure, documentation, and optimization.  
While AI provided support, the core logic and customizations were implemented by the author.

---

## ⚙️ Requirements
**Python Version:** 3.6+  

**Required Libraries:**
```bash
pip install pandas python-docx fuzzywuzzy python-Levenshtein
```

**Built-in Modules Used:**
- (`os`)
- (`datetime`)
- (`csv`)

## 💻 Installation
1. Install Python
2. Install Dependencies
```bash
pip install pandas python-docx fuzzywuzzy python-Levenshtein
```
3. Prepare Input Files
- Excel file with sales data
- ZIP code lookup CSV file
- Optional images (e.g., signature, logo)

---

## 🛠 Configuration
Edit the CONFIG section at the top of the script to set:

| Variable          | Description                  | Example                          |
|-------------------|------------------------------|---------------------------------|
| `EXCEL_FILE`      | Path to Excel sales data     | `"sales_data.xlsx"`             |
| `ZIP_LOOKUP_FILE` | ZIP to city/state CSV file   | `"zip_lookup.csv"`              |
| `LOGO_PATH`       | Path to logo image           | `"logo.png"`                   |
| `YOUR_NAME`       | Your full name               | `"Brian Jones"`                |
| `YOUR_TITLE`      | Your job title               | `"Vice President"`             |
| `YOUR_CO`         | Your company name            | `"Jones Insurance Advisors, Inc"` |
| `YOUR_PHONE`      | Your phone number            | `"(772) 569-6802"`             |
| `YOUR_EMAIL`      | Your email address           | `"Brian@jonesia.com"`          |
| `YOUR_ADDRESS`    | Your full mailing address    | `"3885 20th Street,\nVero Beach, FL 32960"` |
| `YOUR_WEB`        | Your website URL             | `"www.jonesinsuranceadvisors.com"` |
| `YOUR_RETURN_ADDRESS` | Return address formatted from above | `f"{YOUR_NAME}\n{YOUR_ADDRESS}"` |

---

## Usage

1. **Prepare Input Files**

   - Your Excel file (e.g., `sales_data.xlsx`) must contain these columns:
     - `Owner Name`
     - `Address`
     - `Site Zip Code`
     - `Mailing Address`
     - `Sale Date`
     - `Sale Price`
   
   - Your ZIP lookup CSV file (e.g., `zip_lookup.csv`) must have these columns:
     - `zip`
     - `city`
     - `state`
   
   - Place optional images like `signature_brian.png` or `logo.png` in the script directory.

2. **Run the Script**

   Open a terminal or command prompt and navigate to the folder containing the script:

   ```bash
   cd path/to/your/script
   python generate_correspondence.py
The script will:

- Filter out non-owner-occupied properties
- Generate all output files

---

### Review Output Files

| File                    | Description                      |
|-------------------------|--------------------------------|
| `all_letters.docx`       | Personalized letters            |
| `all_envelopes.docx`     | Envelopes with formatted addresses |
| `mailing_labels.docx`    | Avery label sheet format        |
| `crm_owner_occupied.csv` | Data for CRM import             |

---

## Files and Dependencies

### Input Files

- `sales_data.xlsx` — Sales property data
- `zip_lookup.csv` — ZIP to city/state mapping
- `signature_brian.png` — Signature image for letters (optional)
- `logo.png` — Company logo (optional)

### Output Files

- `all_letters.docx`
- `all_envelopes.docx`
- `mailing_labels.docx`
- `crm_owner_occupied.csv`

---

## Notes

- Always double-check file paths in the CONFIG section.
- Envelopes and labels are formatted for standard #10 envelopes and Avery 5160 labels.
- To change spacing or formatting, adjust the layout settings in the script.
