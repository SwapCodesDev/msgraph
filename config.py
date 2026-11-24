"""Configuration loader for Microsoft Graph Excel integration."""

import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env file from project root
BASE_DIR = Path(__file__).resolve().parent
ENV_FILE = BASE_DIR / ".env"
if ENV_FILE.exists():
    load_dotenv(ENV_FILE)
else:
    load_dotenv()


class Config:
    """Application configuration loaded from environment variables."""

    # Authentication settings
    # Flow type: "device_code" (Delegated/User) or "client_credentials" (App-only/Daemon)
    AUTH_FLOW = os.getenv("AUTH_FLOW", "device_code").lower().strip()

    AZURE_CLIENT_ID = os.getenv("AZURE_CLIENT_ID", "").strip()
    AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET", "").strip()
    AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID", "common").strip()

    # Target Excel file settings in OneDrive / SharePoint
    # e.g., "my-app/workbook.xlsx" or "data.xlsx"
    EXCEL_FILE_PATH = os.getenv("EXCEL_FILE_PATH", "my-app/data.xlsx").strip()

    # Target drive options:
    # "me" -> /me/drive (Delegated flow: current signed-in user)
    # "user" -> /users/{TARGET_USER_ID}/drive (Daemon flow accessing a specific user's OneDrive)
    # "site" -> /sites/{TARGET_SITE_ID}/drive (SharePoint Document Library)
    # "drive" -> /drives/{TARGET_DRIVE_ID} (Direct Drive ID)
    DRIVE_TARGET = os.getenv("DRIVE_TARGET", "me").lower().strip()

    TARGET_USER_ID = os.getenv("TARGET_USER_ID", "").strip()
    TARGET_SITE_ID = os.getenv("TARGET_SITE_ID", "").strip()
    TARGET_DRIVE_ID = os.getenv("TARGET_DRIVE_ID", "").strip()

    # Cache file for delegated flow token persistence
    TOKEN_CACHE_FILE = BASE_DIR / "token_cache.bin"

    # Default Graph scopes
    DELEGATED_SCOPES = ["Files.ReadWrite", "User.Read", "offline_access"]
    APP_SCOPES = ["https://graph.microsoft.com/.default"]

    @classmethod
    def validate(cls):
        """Validates critical configuration parameters."""
        errors = []
        if not cls.AZURE_CLIENT_ID:
            errors.append("AZURE_CLIENT_ID is not set. Please provide it in .env file.")

        if cls.AUTH_FLOW == "client_credentials":
            if not cls.AZURE_CLIENT_SECRET:
                errors.append("AZURE_CLIENT_SECRET is required when AUTH_FLOW='client_credentials'.")
            if cls.AZURE_TENANT_ID.lower() in ("common", "consumers"):
                errors.append(
                    f"AZURE_TENANT_ID cannot be '{cls.AZURE_TENANT_ID}' for client_credentials flow. "
                    "Specify your Azure Directory (tenant) ID or domain."
                )
            if cls.DRIVE_TARGET == "me":
                errors.append(
                    "DRIVE_TARGET cannot be 'me' in client_credentials flow (app-only context has no 'me'). "
                    "Set DRIVE_TARGET='user' and provide TARGET_USER_ID, or use DRIVE_TARGET='site' / 'drive'."
                )
            if cls.DRIVE_TARGET == "user" and not cls.TARGET_USER_ID:
                errors.append("TARGET_USER_ID is required when DRIVE_TARGET='user'.")
            if cls.DRIVE_TARGET == "site" and not cls.TARGET_SITE_ID:
                errors.append("TARGET_SITE_ID is required when DRIVE_TARGET='site'.")
            if cls.DRIVE_TARGET == "drive" and not cls.TARGET_DRIVE_ID:
                errors.append("TARGET_DRIVE_ID is required when DRIVE_TARGET='drive'.")

        elif cls.AUTH_FLOW == "device_code":
            if cls.DRIVE_TARGET == "user" and not cls.TARGET_USER_ID:
                errors.append("TARGET_USER_ID is required when DRIVE_TARGET='user'.")
            if cls.DRIVE_TARGET == "site" and not cls.TARGET_SITE_ID:
                errors.append("TARGET_SITE_ID is required when DRIVE_TARGET='site'.")
            if cls.DRIVE_TARGET == "drive" and not cls.TARGET_DRIVE_ID:
                errors.append("TARGET_DRIVE_ID is required when DRIVE_TARGET='drive'.")
        else:
            errors.append(f"Invalid AUTH_FLOW '{cls.AUTH_FLOW}'. Must be 'device_code' or 'client_credentials'.")

        return errors
