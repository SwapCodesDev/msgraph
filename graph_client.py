"""Microsoft Graph API client specialized for Excel workbook operations."""

import io
import urllib.parse
import zipfile
from typing import Any, Dict, List, Optional
import requests
from auth import GraphAuthenticator
from config import Config


class GraphAPIError(Exception):
    """Custom exception for Microsoft Graph API errors with status and error message."""

    def __init__(self, status_code: int, code: str, message: str, raw_response: Any = None):
        super().__init__(f"Graph API Error [{status_code}] {code}: {message}")
        self.status_code = status_code
        self.code = code
        self.message = message
        self.raw_response = raw_response


def create_minimal_xlsx() -> bytes:
    """Generates bytes for a valid, minimal blank Excel (.xlsx) file without external dependencies."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
            '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
            '  <Default Extension="xml" ContentType="application/xml"/>\n'
            '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n'
            '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
            "</Types>",
        )
        z.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
            "</Relationships>",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>\n'
            "</Relationships>",
        )
        z.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
            '  <sheets>\n'
            '    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>\n'
            "  </sheets>\n"
            "</workbook>",
        )
        z.writestr(
            "xl/worksheets/sheet1.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
            "  <sheetData/>\n"
            "</worksheet>",
        )
    return buf.getvalue()


class GraphExcelClient:
    """Client for updating and interacting with Excel files hosted in Microsoft 365 / OneDrive / SharePoint."""

    GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"

    def __init__(self, authenticator: Optional[GraphAuthenticator] = None, device_code_callback=None):
        self.auth = authenticator or GraphAuthenticator()
        self.device_code_callback = device_code_callback
        self.access_token: Optional[str] = None
        self._refresh_token()

    def _refresh_token(self):
        """Fetches or refreshes the access token."""
        self.access_token = self.auth.get_access_token(device_code_callback=self.device_code_callback)

    def _get_headers(self, content_type: str = "application/json") -> Dict[str, str]:
        """Builds HTTP request headers."""
        headers = {
            "Authorization": f"Bearer {self.access_token}",
        }
        if content_type:
            headers["Content-Type"] = content_type
        return headers

    def _request(
        self,
        method: str,
        endpoint: str,
        json_data: Any = None,
        raw_data: Any = None,
        content_type: str = "application/json",
        retry_on_401: bool = True,
    ) -> requests.Response:
        """Executes HTTP request to Microsoft Graph with error handling and 401 retry."""
        url = endpoint if endpoint.startswith("http") else f"{self.GRAPH_BASE_URL}/{endpoint.lstrip('/')}"
        headers = self._get_headers(content_type)

        resp = requests.request(
            method=method,
            url=url,
            headers=headers,
            json=json_data if json_data is not None else None,
            data=raw_data if raw_data is not None else None,
        )

        if resp.status_code == 401 and retry_on_401:
            # Token might have expired, refresh and retry once
            self._refresh_token()
            headers = self._get_headers(content_type)
            resp = requests.request(
                method=method,
                url=url,
                headers=headers,
                json=json_data if json_data is not None else None,
                data=raw_data if raw_data is not None else None,
            )

        if not resp.ok:
            error_code = "Unknown"
            error_message = resp.text
            try:
                err_json = resp.json().get("error", {})
                error_code = err_json.get("code", "Error")
                error_message = err_json.get("message", resp.text)
            except Exception:
                pass
            raise GraphAPIError(resp.status_code, error_code, error_message, resp)

        return resp

    # -------------------------------------------------------------------------
    # Drive and Item Path Resolution
    # -------------------------------------------------------------------------

    def get_drive_base_path(self) -> str:
        """Returns the base API path for the targeted drive."""
        target = Config.DRIVE_TARGET
        if target == "me":
            return "me/drive"
        elif target == "user":
            return f"users/{Config.TARGET_USER_ID}/drive"
        elif target == "site":
            return f"sites/{Config.TARGET_SITE_ID}/drive"
        elif target == "drive":
            return f"drives/{Config.TARGET_DRIVE_ID}"
        else:
            raise ValueError(f"Unsupported DRIVE_TARGET: {target}")

    def _normalize_file_path(self, file_path: Optional[str] = None) -> str:
        """Sanitizes and normalizes the target Excel file path."""
        path = (file_path or Config.EXCEL_FILE_PATH).strip()
        path = path.replace("\\", "/").strip("/")
        if not path.lower().endswith(".xlsx"):
            path = f"{path}.xlsx"
        return path

    def get_file_item_endpoint(self, file_path: Optional[str] = None) -> str:
        """Returns the Graph API endpoint for the file driveItem."""
        path = self._normalize_file_path(file_path)
        encoded_path = urllib.parse.quote(path)
        return f"{self.get_drive_base_path()}/root:/{encoded_path}"

    def get_workbook_endpoint(self, file_path: Optional[str] = None) -> str:
        """Returns the Graph API workbook root endpoint."""
        file_endpoint = self.get_file_item_endpoint(file_path)
        return f"{file_endpoint}:/workbook"

    # -------------------------------------------------------------------------
    # File Existence and Creation
    # -------------------------------------------------------------------------

    def get_file_metadata(self, file_path: Optional[str] = None) -> Optional[Dict[str, Any]]:
        """Gets driveItem metadata for the target Excel file, or None if not found."""
        endpoint = self.get_file_item_endpoint(file_path)
        try:
            resp = self._request("GET", endpoint)
            return resp.json()
        except GraphAPIError as e:
            if e.status_code == 404:
                return None
            raise

    def upload_file(self, content_bytes: bytes, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Uploads file content to the target file path in OneDrive/SharePoint."""
        endpoint = f"{self.get_file_item_endpoint(file_path)}:/content"
        resp = self._request(
            method="PUT",
            endpoint=endpoint,
            raw_data=content_bytes,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        return resp.json()

    def ensure_file_exists(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Ensures the target Excel file exists. If missing, creates a blank workbook."""
        meta = self.get_file_metadata(file_path)
        if meta is None:
            normalized_path = self._normalize_file_path(file_path)
            print(f"[Info] File '{normalized_path}' does not exist on drive. Initializing blank Excel workbook...")
            blank_bytes = create_minimal_xlsx()
            meta = self.upload_file(blank_bytes, file_path)
            print(f"[Success] Created new workbook: {meta.get('webUrl', normalized_path)}")
        return meta

    # -------------------------------------------------------------------------
    # Worksheets Operations
    # -------------------------------------------------------------------------

    def list_worksheets(self, file_path: Optional[str] = None) -> List[Dict[str, Any]]:
        """Lists all worksheets in the target workbook."""
        endpoint = f"{self.get_workbook_endpoint(file_path)}/worksheets"
        resp = self._request("GET", endpoint)
        return resp.json().get("value", [])

    def add_worksheet(self, name: str, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Adds a new worksheet with the specified name to the workbook."""
        endpoint = f"{self.get_workbook_endpoint(file_path)}/worksheets"
        resp = self._request("POST", endpoint, json_data={"name": name})
        return resp.json()

    # -------------------------------------------------------------------------
    # Range Read / Write Operations
    # -------------------------------------------------------------------------

    def read_range(
        self,
        sheet_name: str = "Sheet1",
        address: str = "A1:D10",
        file_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Reads cell data (values, text, formulas) for a given address range."""
        quoted_sheet = urllib.parse.quote(sheet_name)
        quoted_address = urllib.parse.quote(address)
        endpoint = f"{self.get_workbook_endpoint(file_path)}/worksheets('{quoted_sheet}')/range(address='{quoted_address}')"
        resp = self._request("GET", endpoint)
        return resp.json()

    def read_used_range(
        self,
        sheet_name: str = "Sheet1",
        values_only: bool = True,
        file_path: Optional[str] = None,
    ) -> Optional[Dict[str, Any]]:
        """Reads the bounding box of used cells in the worksheet."""
        quoted_sheet = urllib.parse.quote(sheet_name)
        param = "?valuesOnly=true" if values_only else ""
        endpoint = f"{self.get_workbook_endpoint(file_path)}/worksheets('{quoted_sheet}')/usedRange{param}"
        try:
            resp = self._request("GET", endpoint)
            return resp.json()
        except GraphAPIError as e:
            if "ItemNotFound" in e.code or e.status_code == 404:
                return None
            raise

    def update_range(
        self,
        sheet_name: str = "Sheet1",
        address: str = "A1:B2",
        values: Optional[List[List[Any]]] = None,
        formulas: Optional[List[List[str]]] = None,
        file_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Updates values or formulas for a specified cell range in the worksheet.

        Example:
            update_range('Sheet1', 'A1:B2', values=[['Item', 'Price'], ['Laptop', 1200]])
        """
        if values is None and formulas is None:
            raise ValueError("Either 'values' or 'formulas' must be provided to update_range.")

        payload: Dict[str, Any] = {}
        if values is not None:
            payload["values"] = values
        if formulas is not None:
            payload["formulas"] = formulas

        quoted_sheet = urllib.parse.quote(sheet_name)
        quoted_address = urllib.parse.quote(address)
        endpoint = f"{self.get_workbook_endpoint(file_path)}/worksheets('{quoted_sheet}')/range(address='{quoted_address}')"
        resp = self._request("PATCH", endpoint, json_data=payload)
        return resp.json()

    def clear_range(
        self,
        sheet_name: str = "Sheet1",
        address: str = "A1:Z100",
        apply_to: str = "All",
        file_path: Optional[str] = None,
    ) -> None:
        """Clears range contents or formatting. apply_to options: 'All', 'Formats', 'Contents'."""
        quoted_sheet = urllib.parse.quote(sheet_name)
        quoted_address = urllib.parse.quote(address)
        endpoint = f"{self.get_workbook_endpoint(file_path)}/worksheets('{quoted_sheet}')/range(address='{quoted_address}')/clear"
        self._request("POST", endpoint, json_data={"applyTo": apply_to})

    def append_row_to_range(
        self,
        row_values: List[Any],
        sheet_name: str = "Sheet1",
        start_col: str = "A",
        file_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Appends a row to the end of the worksheet's used range.

        If the sheet is completely empty, writes to row 1.
        Otherwise, calculates the next available row index and writes the values.
        """
        used = self.read_used_range(sheet_name=sheet_name, values_only=True, file_path=file_path)

        num_cols = len(row_values)
        if num_cols == 0:
            raise ValueError("row_values cannot be empty.")

        def col_letter(index: int) -> str:
            """Converts 0-based index to Excel column letter (0 -> A, 1 -> B, etc.)."""
            result = ""
            while index >= 0:
                result = chr(index % 26 + ord("A")) + result
                index = index // 26 - 1
            return result

        def parse_col_letter(col_str: str) -> int:
            """Converts column letter like 'A' or 'AA' to 0-based index."""
            num = 0
            for char in col_str.upper():
                num = num * 26 + (ord(char) - ord("A") + 1)
            return num - 1

        start_col_idx = parse_col_letter(start_col)
        end_col_idx = start_col_idx + num_cols - 1
        end_col = col_letter(end_col_idx)

        if not used or not used.get("values"):
            # Empty sheet: start at row 1
            target_row = 1
        else:
            row_index = used.get("rowIndex", 0)  # 0-based
            row_count = used.get("rowCount", 0)
            target_row = row_index + row_count + 1

        target_address = f"{start_col}{target_row}:{end_col}{target_row}"
        return self.update_range(
            sheet_name=sheet_name,
            address=target_address,
            values=[row_values],
            file_path=file_path,
        )

    # -------------------------------------------------------------------------
    # Excel Tables Operations
    # -------------------------------------------------------------------------

    def list_tables(
        self,
        sheet_name: Optional[str] = None,
        file_path: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Lists tables in the workbook or within a specific worksheet."""
        if sheet_name:
            quoted_sheet = urllib.parse.quote(sheet_name)
            endpoint = f"{self.get_workbook_endpoint(file_path)}/worksheets('{quoted_sheet}')/tables"
        else:
            endpoint = f"{self.get_workbook_endpoint(file_path)}/tables"
        resp = self._request("GET", endpoint)
        return resp.json().get("value", [])

    def create_table(
        self,
        address: str,
        table_name: Optional[str] = None,
        sheet_name: str = "Sheet1",
        has_headers: bool = True,
        file_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Creates an Excel table over a given range address."""
        quoted_sheet = urllib.parse.quote(sheet_name)
        endpoint = f"{self.get_workbook_endpoint(file_path)}/worksheets('{quoted_sheet}')/tables/add"
        payload: Dict[str, Any] = {
            "address": address,
            "hasHeaders": has_headers,
        }
        resp = self._request("POST", endpoint, json_data=payload)
        table_data = resp.json()

        # Rename table if a custom name is desired
        if table_name and table_data.get("id"):
            table_id = table_data["id"]
            patch_endpoint = f"{self.get_workbook_endpoint(file_path)}/tables('{table_id}')"
            patch_resp = self._request("PATCH", patch_endpoint, json_data={"name": table_name})
            table_data = patch_resp.json()

        return table_data

    def append_table_row(
        self,
        table_name_or_id: str,
        row_values: List[Any],
        file_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Appends a new row of data to an existing Excel Table."""
        quoted_table = urllib.parse.quote(table_name_or_id)
        endpoint = f"{self.get_workbook_endpoint(file_path)}/tables('{quoted_table}')/rows"
        payload = {"values": [row_values]}
        resp = self._request("POST", endpoint, json_data=payload)
        return resp.json()
