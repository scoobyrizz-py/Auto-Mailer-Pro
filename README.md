 # Auto Mailer Pro
 
 Automate personalized insurance mail campaigns with a polished desktop interface. Auto Mailer Pro generates tailored letters, envelopes, mailing labels, and CRM-ready CSV exports for both personal and commercial property prospects. The tool was developed for Jones Insurance Advisors, Inc. and is designed to accelerate follow-up on recent property sales while preventing duplicate outreach to existing clients.
 
 ---
 
 ## ✨ Key Features
 - **Dual Campaign Modes** – Target *personal* (owner-occupied) or *commercial* audiences with dedicated filtering rules.
 - **Guided GUI Workflow** – Select your data file, campaign mode, templates, and signature branding from a themed Tkinter interface.
 - **Dynamic Letter Content** – Use default market-ready templates or supply your own subject line and body copy per campaign.
 - **Document Automation** – Produce Microsoft Word documents for letters, #10 envelopes, and Avery 5160 mailing labels in one run.
- **CRM Export** – Create a cleaned CSV (`crm_<mode>_occupied.csv`) ready for import into your CRM.
- **Automated Campaign Log** – Each CRM export is appended to `data/campaign_history.db` for long-term tracking.
 - **Data Hygiene** – Cleans owner names, verifies owner-occupancy, validates business types, maps ZIP codes to city/state, and skips existing clients found in the `master_client_list.xlsx` file.
 
 ---
 
 ## 📦 Repository Contents
 | Path | Description |
 | --- | --- |
 | `run.py` | GUI launcher for Auto Mailer Pro. |
 | `AutoMailerPro.py` | Core campaign generation logic. Can be imported or run via the GUI. |
| `AutoMailerPro.exe` | Packaged Windows executable of the GUI (no Python installation required). |
| `assets/` | Branding assets referenced by the GUI and generated documents. |
| `assets/signatures/` | Signature images selectable within the GUI. |
| `data/zip_lookup.csv` | ZIP-to-city/state reference data used during processing. |
| `data/master_client_list.xlsx` | Reference list of existing clients to exclude from outreach. |
| `data/campaign_history.db` | SQLite database that accumulates every CRM export for dashboarding. |
| `output/` | Auto-created workspace where timestamped campaign folders are stored. |
| `archive/`, `dist/` | Build artifacts for distributing the Windows executable. |
 ---
 
 ## ✅ Prerequisites
 
 ### Using the packaged executable (recommended for Windows)
 - Windows 10 or later.
 - No additional dependencies—Python and libraries are bundled inside `AutoMailerPro.exe`.
 
 ### Running from source
 - Python 3.10 or later (tested up to 3.13).
 - Recommended virtual environment (e.g., `python -m venv .venv`).
 - Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   If a `requirements.txt` file is not present, install the libraries manually:
   ```bash
   pip install pandas python-docx fuzzywuzzy python-Levenshtein ttkthemes pillow openpyxl
   ```
 
 ---
 
 ## 🚀 Quick Start
 
### Option 1: Launch the Windows executable
 1. Copy `AutoMailerPro.exe` along with the `assets/` and `data/` folders into the same directory.
 2. Double-click `AutoMailerPro.exe`.
 3. Follow the on-screen workflow (see [Using the Application](#-using-the-application)).
 
 ### Option 2: Run the GUI from source
 1. Clone this repository and open a terminal in the project directory.
 2. Create and activate a virtual environment (optional but recommended).
 3. Install dependencies (see [Prerequisites](#-prerequisites)).
 4. Start the GUI:
    ```bash
    python run.py
    ```
 
 ### Option 3: Automate via Python
 If you need to integrate Auto Mailer Pro into another Python workflow, import and call `AutoMailerPro_v5_1.main`:
 ```python
 from AutoMailerPro_v5_1 import main
 
 main(
     mode="personal",                     # or "commercial"
     file_path="/path/to/sales_data.xlsx",
     content=None,                         # defaults to built-in template per mode
     subject_line="Custom subject here",
     signature_name="Brian Jones",
     signature_title="Vice President",
     signature_image="assets/signatures/signature_brian.png",
 )
 ```
 
 ---
 ### 🛠️ Building the Windows executable yourself
If you need to refresh the packaged application or customize it with your own assets, you can rebuild the Windows executable using [PyInstaller](https://pyinstaller.org/).

1. Ensure Python 3.10+ is installed on your Windows machine.
2. Install the project dependencies and PyInstaller:
   ```bash
   pip install pandas python-docx fuzzywuzzy python-Levenshtein ttkthemes pillow openpyxl
   pip install pyinstaller
   ```
3. From the repository root, run:
   ```bash
   pyinstaller run.spec
   ```
4. After the build completes, the executable and required assets will be located in `dist/AutoMailerPro/`.
5. Distribute the entire folder (or create a zip) so that `AutoMailerPro.exe`, the `assets/` directory, and the `data/` directory stay together.

> ℹ️ PyInstaller must be executed on Windows to produce a Windows `.exe`. Running the build on Linux or macOS generates platform-specific binaries instead.

---

 ## 🗂 Preparing Your Data
 Auto Mailer Pro expects an Excel file exported from your sales source. The required columns vary by campaign mode:
 
 | Column | Personal Mode | Commercial Mode |
 | --- | --- | --- |
 | `Owner Name` | ✅ | – |
 | `Address` | ✅ | ✅ |
 | `Mailing Address` | ✅ | – *(uses `Address` in new commercial format)* |
 | `Site Zip Code` | ✅ | – |
 | `ZIP Code` | – | ✅ *(new commercial format)* |
 | `Sale Date` | ✅ | Optional |
 | `Sale Price` | ✅ | Optional |
 | `Business Type` | – | ✅ *(legacy commercial format)* |
 | `Executive First Name` / `Executive Last Name` | – | ✅ *(new commercial format)* |
 
    Additional reference files:
    - **`data/zip_lookup.csv`** – Maps 5-digit ZIP codes to city/state values written to output documents.
    - **`data/master_client_list.xlsx`** – Existing client roster. Any record with a matching name and mailing address is automatically skipped.
 
 > 💡 Tip: Place your sales Excel file in the same directory as the application for easier browsing.
 
 ---
 
 ## 🧭 Using the Application
 1. **Launch the GUI** via the executable or `python run.py`.
 2. **Select Campaign Mode** (`Personal` or `Commercial`). The subject line and default letter template update automatically.
 3. **Choose a Letter Template**:
    - `Indian River` and `St. Lucie` presets for both personal and commercial audiences.
    - `Custom` to edit the body copy directly in the GUI.
 4. **Pick a Signature Profile** to apply a name, title, email address, and signature image.
 5. **Load Sales Data** by clicking **Browse**, then selecting your Excel file.
 6. **Adjust Subject Line** if desired. If you type in the subject box, the value stays locked even when switching modes.
 7. **Review Letter Content** in the scrollable preview. Custom content is fully editable.
 8. Click **Run Campaign**. Progress updates appear in the output console at the bottom of the window.
 9. When processing completes, a timestamped folder (e.g., `output/031224_1430_Personal_Mailing_Campaign`) is created with all generated files.
 
 ---
 
 ## 📄 Output Files
 Each campaign generates the following assets inside the timestamped output folder:
 
 | File | Description |
 | --- | --- |
 | `all_letters.docx` | Personalized letter for each qualified recipient. Subject line is bolded at the top. |
 | `all_envelopes.docx` | #10 envelope layout, one per recipient. |
 | `mailing_labels.docx` | Avery 5160-compatible 3×10 sheet of labels. |
 | `crm_<mode>_occupied.csv` | Filtered and cleaned contact list for CRM import. |
 | `data/campaign_history.db` | Consolidated log of every contact mailed, updated after each run. |
 | `processing_log.txt` *(optional)* | Console output when redirected via GUI (copy from output panel if needed). |
---

## 📊 Campaign History Database

Every successful campaign automatically appends its CRM-ready rows to the SQLite file at `data/campaign_history.db`. The `campaign_contacts` table includes the campaign folder name (`campaign_id`), mode, send timestamp, and the cleaned contact fields. Connect the database to Excel, Google Data Studio, Metabase, or any BI tool to blend in response/conversion outcomes without manually merging CSV exports.

---

## 🛠 Configuration & Customization
 - **Templates** – Edit the predefined templates within `run.py` or pass a custom string to `main(content=...)`.
 - **Signatures** – Add new entries to the `signature_profiles` dictionary in `run.py`, pointing to PNG files stored under `assets/signatures/`.
 - **Branding** – Replace `Logo.png` or `logo.ico` to update visuals shown in the GUI and exported letters.
 - **Data Rules** – Advanced logic (name cleaning, filtering, CRM export) resides in `AutoMailerPro_v5_1.py`. Adjust the helper functions there for bespoke workflows.
 
 ---
 
 ## 🧪 Troubleshooting
 | Issue | Resolution |
 | --- | --- |
 | *Excel file not found* | Confirm the path shown in the GUI, or move the Excel file into the project directory. |
 | *Missing columns* | Ensure your spreadsheet headers exactly match the column names listed in [Preparing Your Data](#-preparing-your-data). |
 | *Existing clients still receiving letters* | Update `master_client_list.xlsx` with the latest client roster. Matching is case-insensitive on name and mailing address. |
 | *Signature image not loading* | Verify the PNG file exists and is listed in `signature_profiles`. |
 | *GUI fails to start due to missing theme assets* | Install `ttkthemes` and `pillow`, then re-run `python run.py`. |
 
 ---
 
 ## 🤝 Support
 For assistance, feature requests, or to report issues, contact **Kyle Padilla** at [scooby_rizz@proton.me](mailto:scooby_rizz@proton.me).
 
 ---
 
 © 2025 Jones Insurance Advisors, Inc. All rights reserved.

