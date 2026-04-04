# Setup Guide for New Members

## Prerequisites Installation Guide

This guide will help you set up Python and pip on your Windows machine to run the project.

---

## Step 1: Install Python

### Option A: Download from Official Website (Recommended)

1. Visit: https://www.python.org/downloads/
2. Click the **"Download Python 3.12"** button (or latest version)
3. Run the installer (.exe file)
4. **IMPORTANT**: Check the box **"Add Python to PATH"** ✓
5. Click **"Install Now"** or **"Customize installation"**
6. Wait for installation to complete
7. Click **"Close"**

### Option B: Using Windows Package Manager (If available)

```powershell
winget install Python.Python.3.12
```

---

## Step 2: Verify Python Installation

Open **PowerShell** or **Command Prompt** and run:

```powershell
python --version
```

You should see something like:
```
Python 3.12.x
```

---

## Step 3: Verify pip Installation

pip comes with Python automatically. Verify it:

```powershell
pip --version
```

You should see something like:
```
pip 24.x.x from C:\Users\...\Python3xx\lib\site-packages\pip (python 3.12)
```

---

## Step 4: Upgrade pip (Optional but Recommended)

```powershell
python -m pip install --upgrade pip
```

---

## Step 5: Clone or Download the Project

Navigate to your project directory:

```powershell
cd "C:\Backend Engineer\agent-ai\word"
```

---

## Step 6: Install Project Dependencies

Install all required libraries from requirements.txt:

```powershell
pip install -r requirements.txt
```

This will install:
- pandas - Data manipulation
- docxtpl - Word template generation
- python-docx - Word document handling
- xlrd - Excel file reading (.xls)
- openpyxl - Excel file reading (.xlsx)

---

## Step 7: Test the Installation

Run the main script to verify everything works:

```powershell
python main.py
```

You should see output like:
```
046184002665
046171002162
✅ Done! Đã tạo file Word hàng loạt từ CSV.
```

---

## Troubleshooting

### Issue: "python" command not found

**Solution**: Python is not in PATH. Reinstall Python and **make sure to check "Add Python to PATH"** during installation.

### Issue: "pip" command not found

**Solution**: Run using:
```powershell
python -m pip install -r requirements.txt
```

### Issue: Permission Denied

**Solution**: Run PowerShell as Administrator:
1. Right-click PowerShell
2. Select "Run as Administrator"
3. Try again

### Issue: Module not found after pip install

**Solution**: Make sure you're in the correct directory and using the same Python:
```powershell
python main.py
```

---

## Project Structure

```
word/
├── main.py                  # Main script
├── template2.docx          # Word template
├── text.xls                # Input Excel file
├── requirements.txt        # Python dependencies
├── output/                 # Generated Word documents
└── README.md              # This file
```

---

## File Descriptions

- **main.py**: Main Python script that reads Excel data and generates Word documents
- **template2.docx**: Template file used for document generation
- **text.xls**: Input Excel file with customer data
- **requirements.txt**: List of Python packages needed
- **output/**: Folder where generated Word documents are saved

---

## Quick Start

After installation, to run the project:

```powershell
cd "C:\Backend Engineer\agent-ai\word"
python main.py
```

Generated documents will appear in the `output/` folder.

---

## Need Help?

If you encounter any issues:
1. Check the Troubleshooting section above
2. Verify Python is installed: `python --version`
3. Verify pip is installed: `pip --version`
4. Make sure you're in the correct directory
5. Contact the project lead

---

**Last Updated**: April 4, 2026
