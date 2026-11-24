# Complete Guide: Microsoft Graph API & Azure Entra ID Setup

This guide walks you step-by-step through setting up Microsoft Graph API credentials in **Microsoft Entra ID** (formerly Azure Active Directory) so your Python application can read and update Excel workbooks stored in OneDrive or SharePoint.

---

## 1. Choose Your Authentication Flow

Microsoft Graph supports two primary ways to authenticate:

| Feature | Option A: Delegated (Device Code Flow) **[Recommended for Most Users]** | Option B: Application (Client Credentials Flow) |
| :--- | :--- | :--- |
| **Best For** | Personal Microsoft accounts (`@outlook.com`, `@hotmail.com`), work accounts, developer testing | Fully headless daemons, scheduled tasks, background server workers |
| **User Sign-in** | One-time sign-in in browser (subsequent runs use cached tokens) | Completely non-interactive (no user prompts) |
| **Admin Consent** | Usually **not required** (users consent for themselves) | **Requires Tenant Admin Consent** |
| **Account Type** | Personal accounts, Work/School accounts | Work/School organizational tenants only |
| **Secret Needed** | No client secret needed | Client Secret required |

---

## 2. Step-by-Step Azure Portal Setup

### Step 1: Open the Azure Portal
1. Navigate to the [Microsoft Entra Admin Center](https://entra.microsoft.com/) or the [Azure Portal](https://portal.azure.com/).
2. Sign in with your Microsoft account (personal or work/school).

---

### Step 2: Register a New Application
1. In the search bar at the top, search for **Microsoft Entra ID** (or **Azure Active Directory**).
2. In the left navigation menu, click **App registrations**.
3. Click the **+ New registration** button at the top.
4. Fill in the registration form:
   - **Name**: Enter a descriptive name (e.g., `MyApp-Excel-Updater` or `MSGraph-Excel-CLI`).
   - **Supported account types**:
     - If you have a **Personal Microsoft account** (e.g. `@gmail.com`, `@outlook.com`, `@hotmail.com`), choose:
       > **Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**
     - If you are in a **Company / School tenant**, choose:
       > **Accounts in this organizational directory only** (Single tenant)
   - **Redirect URI (optional)**:
     - Select platform: **Public client/native (mobile & desktop)**
     - URI: Enter `https://login.microsoftonline.com/common/oauth2/nativeclient`
5. Click **Register** at the bottom.

---

### Step 3: Copy Your IDs
After the app is created, you will be on the **Overview** page:
1. Copy the **Application (client) ID** (a GUID like `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`).
2. Copy the **Directory (tenant) ID**:
   - For **Work/School** single-tenant accounts: copy the GUID.
   - For **Personal** Microsoft accounts: you can use `common` or `consumers` in your `.env`.

---

### Step 4: Configure Authentication Settings (Crucial for Device Code Flow)
1. In the left menu of your app registration, click **Authentication**.
2. Scroll down to the section titled **Advanced settings**.
3. Under **Allow public client flows**:
   - Set **Enable the following mobile and desktop flows** to **Yes**.
   - *(This allows the script to use the Device Code Flow without opening web ports).*
4. Click **Save** at the top.

---

### Step 5: Configure API Permissions

1. In the left menu, click **API permissions**.
2. Click **+ Add a permission**.
3. Select **Microsoft Graph**.

#### For Option A: Delegated Permissions (Recommended for Device Code Flow)
1. Select **Delegated permissions**.
2. Search and check the following permissions:
   - `Files.ReadWrite` (Allows reading and writing files in the user's OneDrive)
   - `User.Read` (Allows reading the user's basic profile)
   - `offline_access` (Allows refreshing tokens silently without re-login)
3. Click **Add permissions**.
4. *(Optional / If in enterprise)*: If you are an admin and want to bypass user consent prompts, click **Grant admin consent for [Your Organization]**.

#### For Option B: Application Permissions (For Unattended Client Credentials Daemons)
1. Select **Application permissions**.
2. Search and check:
   - `Files.ReadWrite.All` (Allows the app to read and write files in any drive without a signed-in user).
3. Click **Add permissions**.
4. **Mandatory**: Click **Grant admin consent for [Your Organization]** and approve.

---

### Step 6: Create a Client Secret (Only for Option B: Client Credentials)
> [!NOTE]
> If you are using **Option A (Device Code Flow)**, you can **skip this step**.

1. In the left menu, click **Certificates & secrets**.
2. Under the **Client secrets** tab, click **+ New client secret**.
3. Add a description (e.g. `ExcelAutomationKey`) and select an expiration period (e.g. 180 days).
4. Click **Add**.
5. **CRITICAL**: Copy the **Value** immediately (NOT the Secret ID). Microsoft hides this value permanently once you leave the page.

---

## 3. Configuring the `.env` File

In the root of your project (`d:\Projects\msgraph`), create a `.env` file by copying the template:

```powershell
Copy-Item .env.example .env
```

### Scenario 1: Personal OneDrive or Interactive Login (Delegated / Device Code)

```ini
AUTH_FLOW=device_code
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=
AZURE_TENANT_ID=common
EXCEL_FILE_PATH=my-app/data.xlsx
DRIVE_TARGET=me
```

### Scenario 2: Corporate Microsoft 365 Daemon (Application / Client Credentials)

```ini
AUTH_FLOW=client_credentials
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=your_secret_value_here
AZURE_TENANT_ID=your-tenant-guid-here
EXCEL_FILE_PATH=my-app/data.xlsx
DRIVE_TARGET=user
TARGET_USER_ID=user@yourdomain.com
```

### Scenario 3: SharePoint Document Library

```ini
AUTH_FLOW=client_credentials
AZURE_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
AZURE_CLIENT_SECRET=your_secret_value_here
AZURE_TENANT_ID=your-tenant-guid-here
EXCEL_FILE_PATH=my-app/data.xlsx
DRIVE_TARGET=site
TARGET_SITE_ID=yourcompany.sharepoint.com,xxxxxxxx-xxxx,yyyyyyyy-yyyy
```

---

## 4. Understanding How MS Graph Finds Your Excel File

Microsoft Graph organizes files by **Drives** and **driveItems**:

1. **Path convention**:
   - `EXCEL_FILE_PATH=my-app/data.xlsx` will look for a folder named `my-app` in your OneDrive, containing `data.xlsx`.
   - If the folder or file doesn't exist yet, simply run:
     ```powershell
     python main.py ensure-file
     ```
     The tool will automatically create a valid blank Excel workbook at that path!
2. **Accessing Excel Online**:
   - Once created, you can open `my-app/data.xlsx` directly in Excel Online or desktop Excel.
   - Any updates made through this script will be reflected in real time.

---

## 5. Troubleshooting & Diagnostics

| Error Code / Message | Cause | Resolution |
| :--- | :--- | :--- |
| **`AADSTS700016: Application ... not found`** | `AZURE_CLIENT_ID` is wrong or typed incorrectly. | Check your App Registration Overview page and copy the exact Application ID. |
| **`AADSTS7000218: The request body must contain client_secret`** | App has not enabled public client flows for `device_code`. | Go to **Authentication** -> **Advanced Settings** -> enable **Allow public client flows = Yes**. |
| **`AADSTS65001: The user or admin has not consented`** | The account has not granted consent to the permissions. | Sign in during device code flow and click **Accept**, or have your tenant admin grant consent. |
| **`ItemNotFound (404)`** | File or worksheet path does not exist in OneDrive. | Run `python main.py ensure-file` or verify the file path. |
| **`AccessDenied (403)`** | App lacks `Files.ReadWrite` or `Files.ReadWrite.All` permission. | Check **API permissions** in Azure portal and make sure permissions are added and consented. |
| **`ResourceLocked / 423`** | File is being edited with an exclusive lock. | Close active local desktop Excel instances or wait 30 seconds for the web session to release the lock. |
