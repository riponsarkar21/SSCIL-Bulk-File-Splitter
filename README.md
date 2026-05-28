
### 1. Repository Description (Short Blurb)

> **Excel VBA Bulk Carrier Splitter**
>
> A powerful automation tool for logistics and data entry teams. This Excel VBA macro simplifies the process of splitting sequential bulk carrier loading records into separate, brand-specific files (BM, AM, Gold). It features a user-friendly "Helper" interface, intelligent column detection, and ensures no data is lost or modified in the original source files.

---

### 2. README.md


```markdown
# Bulk Carrier Data Splitter (Excel VBA)

An efficient Excel VBA tool designed to automate the separation of bulk carrier loading records. Instead of manually filtering and copying data into different files, this macro allows you to import a master sequential file and instantly split it into three separate files based on brand (BM, AM, Gold).

## 🚀 Features

- **One-Click Splitting:** Automatically separates mixed data into individual `.xlsx` files.
- **Smart Column Detection:** Intelligently finds the "Brand" column by searching for headers (Type, Cargo, Brand) or data values.
- **Source Protection:** Opens source files in "Read-Only" mode, preventing accidental data modification.
- **Auto-Setup:** Automatically creates required sheets (`Info`, `Main`, `BM`, `AM`, `Gold`) if they are missing.
- **User-Friendly Interface:** Simple control panel on an `Info` sheet with button-based macros.

## 📋 Requirements

- Microsoft Excel (Windows).
- Macros enabled (VBA).
- Basic understanding of how to insert VBA code into a module.

## 🛠️ Installation & Setup

1. **Create the Helper File:**
   - Open a new (or existing) Excel Workbook.
   - Save it as an **Excel Macro-Enabled Workbook (.xlsm)**.

2. **Insert the Code:**
   - Press `Alt + F11` to open the VBA Editor.
   - Go to `Insert > Module`.
   - Copy the code from `Module.bas` in this repository and paste it into the module window.
   - Close the VBA Editor.

3. **Create Control Buttons:**
   - Go to the `Info` sheet (the macro will create this automatically if you run it once).
   - Enable the "Developer" tab in Excel if not already visible.
   - Go to `Developer > Insert > Button (Form Control)`.
   - Draw three buttons and assign the following macros:
     - Button 1: `SelectSourceFile`
     - Button 2: `ImportSelectedSheet`
     - Button 3: `SplitDataToFiles`

## 📖 Usage Guide

### Step 1: Select Source File
1. Click the **Select File** button.
2. A file dialog will appear. Navigate to and select your master source file (containing mixed BM/AM/Gold data).
3. The macro will list the sheets found in the source file in Column **E**.
4. **Important:** Enter the number `1` in Column **F** next to the specific sheet name you wish to import.

### Step 2: Get Data
1. Click the **Get Data** button.
2. The macro will verify you have selected only one sheet.
3. It will copy all data (formats, formulas, values) from the source file into the `Main` sheet of your Helper file.

### Step 3: Split Data
1. Click the **Split Data** button.
2. The macro will scan the `Main` sheet to identify the Brand column.
3. It will create three new Excel files in the same folder as your Helper file:
   - `Bulk_Record_BM_[Date].xlsx`
   - `Bulk_Record_AM_[Date].xlsx`
   - `Bulk_Record_Gold_[Date].xlsx`

## ⚠️ Troubleshooting

- **"Variable not defined" Error:** Ensure `Option Explicit` is present at the very top of your code module.
- **Macros not running:** Check your Excel Trust Center settings (`File > Options > Trust Center > Macro Settings`) and ensure "Enable VBA macros" is selected.
- **Brand not found:** Ensure your data has a column header named "Type", "Cargo", "Brand", or contains the exact values "BM", "AM", or "Gold" in the rows.

## 📝 License

This project is open source and available for educational and commercial use.

## 🤝 Contributing

Suggestions and improvements are welcome! Please feel free to submit a Pull Request.
```

---

### 3. .gitignore
*Save this content as `.gitignore` to prevent temporary Excel files from cluttering your repo.*

```gitignore
# Excel Temporary Files
~$*
*.tmp
*.temp

# OS Generated Files
.DS_Store
.DS_Store?
._*
.Spotlight-V100
.Trashes
ehthumbs.db
Thumbs.db

# User specific files
*.xlsm
# Uncomment the line above ONLY if you do not want to track the macro file itself.
# Usually, you want to track the .xlsm if this is a code repository.
```

---

### 4. LICENSE (MIT)
*Save this content as `LICENSE`.*

```markdown
MIT License

Copyright (c) 2023 [Your Name/Organization]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

### 5. CHANGELOG.md
*Save this content as `CHANGELOG.md` to track updates.*

```markdown
# Changelog

All notable changes to the "Bulk Carrier Data Splitter" project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2023-10-27

### Added
- Initial release of the Bulk Carrier Splitter.
- `SelectSourceFile` macro to open and list source file sheets.
- `ImportSelectedSheet` macro to copy data to the Helper `Main` sheet.
- `SplitDataToFiles` macro to separate records into BM, AM, and Gold files.
- Auto-setup feature to generate required sheets (`Info`, `Main`, `BM`, `AM`, `Gold`).
- Smart Brand Column detection logic.
- Error handling for "Variable not defined" and missing sheets.
```

https://chat.z.ai/s/72329f29-1876-4811-a5c8-d2b351e3cd78 for more details
