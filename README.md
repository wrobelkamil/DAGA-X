![DAGA X Logo](DAGAX.png)

# DAGA X

**Automate and streamline technical specification workflows in Adobe InDesign.**

DAGA X (Dowolne Automatyczne Generowanie Artykułów) is a Python-based desktop application that simplifies the creation, management, and merging of technical specifications in Adobe InDesign. From folder setup to CSV-driven data merges and barcode generation, DAGA X integrates multiple modules into a single, user-friendly interface.

---

## Table of Contents
- [Features](#features)
- [Screenshots](#screenshots)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [Modules](#modules)
- [Dependencies](#dependencies)
- [Contributing](#contributing)
- [License](#license)

---

## Features

- **Folder Creation**: Automatically create project directories with predefined subfolders and copy template files. Relink special assets (e.g., AI links) via ExtendScript for InDesign.
- **DAGA Module (CSV → InDesign)**: Generate and import CSV data merges directly into InDesign documents. Supports custom fields, multiple color groups, image previews, and quantity calculations.
- **KeyNoteCoder Module**:
  - **Note Generation**: Build aligned TXT notebooks for color lists, excluding specific keywords.
  - **Barcode Generation**: Produce EAN‑13 barcodes in JPG or PDF formats from text input or files, with automated cropping and margin adjustments.
- **Excel Integration**: Browse and open project paths listed in an Excel workbook. Live search and refresh data without restarting the app.
- **Context Menus & Shortcuts**: Right-click paste menus, keyboard shortcuts for common actions (Ctrl‑E for export, Ctrl‑I for import, etc.).
- **Dark Theme**: Modern dark UI powered by `sv_ttk` and custom icons.

---

## Screenshots

<img src="screenshots/folder_creation.png" alt="Folder Creation Tab" width="400" />
<img src="screenshots/daga_module.png" alt="DAGA Module" width="400" />

---

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/daga-x.git
   cd daga-x
   ```
2. (Optional) Create and activate a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # macOS/Linux
   venv\\Scripts\\activate   # Windows
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

**Note**: This application is tested on Windows 10/11 with Adobe InDesign installed and Python 3.8+.

---

## Usage

```bash
python main.py
```

- Use the **FOLDER** tab to configure and create project directories.
- Switch to the **DAGA** tab for CSV data merges with InDesign.
- Open the **CODE** tab for note and barcode generation workflows.
- Access **PROJEKTY** to browse Excel-managed project paths.

---

## Configuration

Edit the following constants in `main.py` to match your environment:

```python
BASE_IMPORT_PATH    = r"S:\\_KOLEKCJE_\\IMPORT"
TEMPLATE_FILE       = r"S:\\Graficy Specyfikacje\\Techniczne\\szablony_specyfikacje różne\\SPECYFIKACJA_IMPORT.indd"
EXCEL_FILE          = r"G:\\_Projekty\\projekty.xlsx"
EXCEL_ICON_PATH     = r"S:\\Graficy Specyfikacje\\Techniczne\\icons\\icons8-excel-32.png"
```

---

## Modules

### Folder Creation
- Scans `BASE_IMPORT_PATH` for season and employee folders.
- Creates `WIZ`, `CODE`, and optionally a relinked template INDD.
- Runs ExtendScript via COM to update AI links.

### DAGA (CSV → InDesign)
- Define up to 6 color groups, image paths, and metadata.
- Export to CSV with proper quoting and headers.
- Optional direct data merge into InDesign via COM.

### KeyNoteCoder
- Generate aligned TXT files for color lists.
- Extract EAN‑13 codes from text or files and fetch barcodes using BWIP‑JS API.
- Crop, margin-adjust, and save as JPG or PDF.

### ExcelTab
- Loads an Excel workbook (`openpyxl`) and displays project names with navigation buttons.

---

## Dependencies

- Python 3.8+
- Tkinter (built-in)
- `tkcalendar`
- `sv_ttk`
- `Pillow`
- `openpyxl`
- `pywin32`
- `requests`

Install all via:
```bash
pip install -r requirements.txt
```

---

## Contributing

1. Fork the repo.
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Commit your changes: `git commit -am 'Add feature'
4. Push to the branch: `git push origin feature/my-feature`
5. Open a Pull Request.

Please follow the existing code style and add tests where applicable.

---

## License

Copyright (c) 2025 Kamil Wróbel

All rights reserved.

This software is proprietary and confidential.  
Unauthorized copying, distribution, modification, or commercial use is strictly prohibited without explicit written permission from the author.

---

*Developed with ❤️ by Kamil Wróbel*

