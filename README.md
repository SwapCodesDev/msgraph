# Microsoft Graph Excel Updater (Python CLI & Library)

A lightweight, robust, headless Python tool and module that uses the **Microsoft Graph API** to update, append, and manage Excel workbooks (`.xlsx`) hosted in **OneDrive** or **SharePoint** (including dedicated project folders such as `my-app/data.xlsx`).

---

## 🌟 Key Features

- **Dual Interfaces**:
  - **Headless CLI & Library**: For automated scripts, terminal workflows, and CI/CD cron jobs.
  - **Single-File Desktop GUI (`gui.py`)**: Intuitive desktop interface with a live spreadsheet grid, one-click login dialog, and visual data entry tools.
- **Dual Authentication Modes**:
  - **Device Code Flow (Delegated)**: Sign in once with any web browser (no local server/ports needed); tokens are securely cached for seamless recurring runs. Works with both **personal Microsoft accounts** (`@outlook.com`, `@hotmail.com`) and **Work/School accounts** without requiring tenant admin privileges.
  - **Client Credentials Flow (Application)**: True headless daemon authentication with Client ID and Secret for corporate servers and automated daemons.
- **Excel Operations Supported**:
  - **Update Range**: Modify specific cell ranges (e.g., `A1:C2`) with custom 2D arrays or JSON files.
  - **Append Row**: Dynamically append rows to the bottom of any worksheet (auto-detects the used range) or directly to an **Excel Table**.
  - **Read Range**: View data from specific cell addresses or read the full used range.
  - **Worksheet Management**: List and create new sheets.
  - **Table Management**: List tables and convert cell ranges into named Excel Tables.
  - **Auto-Provisioning**: Automatically creates and initializes a blank `.xlsx` workbook on OneDrive/SharePoint if the target file does not yet exist.

---

## 📁 Project Structure

```
d:\Projects\msgraph\
├── .env.example              # Sample environment configuration
├── .gitignore                # Git ignore rules for venv, tokens, and secrets
├── requirements.txt          # Python dependencies (msal, requests, python-dotenv)
├── config.py                 # Configuration loader and validation
├── auth.py                   # MSAL authentication handler with token cache
├── graph_client.py           # MS Graph Excel API wrapper & minimal XLSX generator
├── main.py                   # Headless CLI interface
├── gui.py                    # Single-file Desktop GUI interface
├── example_usage.py          # Python programmatic usage script
├── AZURE_SETUP_GUIDE.md      # Detailed step-by-step Azure portal registration guide
└── README.md                 # Project documentation
```

---

## 🚀 Quick Start & Installation

### 1. Prerequisites
- **Python 3.9+** (Tested on Python 3.13)
- An active Microsoft account (Personal or Work/School)

### 2. Set Up Virtual Environment

Open PowerShell or Command Prompt in the project folder:

```powershell
# Create virtual environment
python -m venv venv

# Activate virtual environment (Windows PowerShell)
.\venv\Scripts\Activate.ps1

# (Alternative: Windows Command Prompt cmd.exe)
# .\venv\Scripts\activate.bat
```

### 3. Install Dependencies

```powershell
pip install -r requirements.txt
```

---

## 🔑 Azure Entra ID / API Configuration

Before running the program, you need to register an application in the Azure Portal to get your API credentials.

> 📖 **Read the full step-by-step tutorial**: [AZURE_SETUP_GUIDE.md](file:///d:/Projects/msgraph/AZURE_SETUP_GUIDE.md)

### Summary of Azure Registration:
1. Go to [Microsoft Entra Admin Center](https://entra.microsoft.com/) or [Azure Portal](https://portal.azure.com/).
2. Register a new App Registration (select multi-tenant + personal accounts if using a personal account).
3. Under **Authentication** -> **Advanced Settings**, enable **Allow public client flows = Yes**.
4. Under **API Permissions**, add Microsoft Graph delegated permissions: `Files.ReadWrite`, `User.Read`, `offline_access`.
5. Copy your **Application (client) ID**.

### Configure `.env`:
Create your `.env` file from the example:

```powershell
Copy-Item .env.example .env
```

Edit `.env` with your values:
```ini
# Flow: "device_code" (Recommended for personal/interactive) or "client_credentials" (Daemon)
AUTH_FLOW=device_code

# Your Application (Client) ID from Azure Portal
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

# Tenant ID ("common" for personal accounts or your Tenant GUID for corporate)
AZURE_TENANT_ID=common

# Target Excel path in OneDrive (e.g. "my-app/data.xlsx")
EXCEL_FILE_PATH=my-app/data.xlsx

# Drive type: "me" (current user), "user" (specific user), or "site" (SharePoint)
DRIVE_TARGET=me
```

---

## 🖥️ Graphical User Interface (GUI)

If you prefer a visual interface rather than running terminal commands, launch the single-file desktop GUI:

```powershell
python gui.py
```

### GUI Features:
- **Live Spreadsheet Grid**: Browse worksheets, inspect cell values, and auto-detect used ranges in an interactive table view.
- **One-Click Sign In**: Prompts the Device Code popup with a 1-click **"📋 Copy Code"** button and direct link to Microsoft's login page.
- **Visual Data Updates & Appends**: Form-based tools to append rows or write 2D arrays directly into cells.
- **Direct Web Access**: Click **"🌐 Open in Excel Online"** to jump immediately to the live workbook in Microsoft 365.
- **Built-in Settings Manager**: Edit and save Azure credentials to `.env` without touching configuration files manually.
- **Non-blocking Threading**: Background workers ensure the interface remains smooth and never freezes during network operations.

---

## 💻 CLI Usage Guide

Run all commands using `python main.py <command>`.

### 1. Check Configuration & Connection
Checks your credentials, drive connection, and target file metadata:
```powershell
python main.py info
```

### 2. Ensure File Exists (Create if Missing)
Verifies if the remote file exists in OneDrive/SharePoint. If missing, it automatically creates a fresh blank workbook at that path:
```powershell
python main.py ensure-file
```

### 3. List Worksheets
```powershell
python main.py list-sheets
```

### 4. Add a New Worksheet
```powershell
python main.py add-sheet --name "Q1_Metrics"
```

### 5. Update a Range of Cells
Update cells with a 2D JSON array:
```powershell
python main.py update-range --sheet Sheet1 --address A1:C2 --values '[["Product", "Category", "Price"], ["Laptop", "Electronics", 1200]]'
```

Or load values from a JSON file:
```powershell
python main.py update-range --sheet Sheet1 --address A1:C10 --json-file my_data.json
```

### 6. Append Rows Dynamically
Append a row to the end of the sheet's used range (automatically detects the next available row):
```powershell
python main.py append-row --sheet Sheet1 --values '["2026-11-24", "my-app Automated Run", 99.8]'
```

### 7. Create and Append to Excel Tables
Convert a range into a formatted Excel Table:
```powershell
python main.py create-table --sheet Sheet1 --address A1:C10 --name "SalesTable"
```

Append a row directly into the Excel Table:
```powershell
python main.py append-row --table SalesTable --values '["2026-11-24", "Desktop PC", 850]'
```

### 8. Read Range Data
Read an explicit cell range:
```powershell
python main.py read-range --sheet Sheet1 --address A1:C5
```
Or read all active cells (used range):
```powershell
python main.py read-range --sheet Sheet1
```

### 9. Clear Cells
Clear contents and formats:
```powershell
python main.py clear-range --sheet Sheet1 --address A2:C50 --apply-to Contents
```

### 10. Authentication Management
Trigger one-time login explicitly:
```powershell
python main.py auth-login
```

Clear cached tokens:
```powershell
python main.py auth-logout
```

---

## 🐍 Using as a Python Library in Your Scripts

You can also import and use `GraphExcelClient` directly in your custom Python automation scripts without the CLI.

```python
from datetime import datetime
from graph_client import GraphExcelClient

# Initialize client (loads settings from .env)
client = GraphExcelClient()

target_path = "my-app/data.xlsx"

# 1. Ensure file exists in OneDrive
client.ensure_file_exists(target_path)

# 2. Write headers
client.update_range(
    sheet_name="Sheet1",
    address="A1:C1",
    values=[["Timestamp", "Event", "Metric"]],
    file_path=target_path
)

# 3. Append dynamic data rows
now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
client.append_row_to_range(
    row_values=[now, "Backup Complete", 100],
    sheet_name="Sheet1",
    file_path=target_path
)

# 4. Read back the updated sheet
data = client.read_used_range(sheet_name="Sheet1", file_path=target_path)
print("Updated values:", data.get("values"))
```

---

## 🔒 Security Best Practices

1. **Never commit `.env` or `token_cache.bin`**: The `.gitignore` file is configured to keep these files private.
2. **Principle of Least Privilege**:
   - Use `Files.ReadWrite` for user-scoped access instead of tenant-wide `Files.ReadWrite.All` whenever possible.
3. **Secrets Management**: If using `client_credentials`, rotate Client Secrets regularly in Azure Entra ID.

---

## ❓ Frequently Asked Questions

**Q: Where is the Excel file stored?**  
A: The file is stored in Microsoft Cloud (OneDrive Personal, OneDrive for Business, or SharePoint Document Library). The default path `my-app/data.xlsx` will appear in your OneDrive under a `my-app` folder.

**Q: Can I open and edit the file while the script runs?**  
A: Yes! You can have the file open in Excel Online or the Excel desktop app. The updates made by MS Graph API will sync in real time (co-authoring).

**Q: Do I need a paid Azure subscription?**  
A: **No.** Azure App Registrations and Microsoft Graph API access for OneDrive personal and Microsoft 365 developer tenants are completely free to register and use.
