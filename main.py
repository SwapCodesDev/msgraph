"""Command-line interface (CLI) for updating Excel files using Microsoft Graph API."""

import argparse
import json
import sys
from typing import Any, List
from auth import GraphAuthenticator
from config import Config
from graph_client import GraphAPIError, GraphExcelClient


def print_table(headers: List[str], rows: List[List[Any]]):
    """Renders a simple ASCII table to stdout."""
    all_rows = [headers] + [[str(c) if c is not None else "" for c in r] for r in rows]
    col_widths = [max(len(str(row[i])) for row in all_rows) for i in range(len(headers))]

    separator = "+" + "+".join("-" * (w + 2) for w in col_widths) + "+"
    print(separator)
    header_line = "| " + " | ".join(f"{str(headers[i]).ljust(col_widths[i])}" for i in range(len(headers))) + " |"
    print(header_line)
    print(separator)
    for r in rows:
        row_line = "| " + " | ".join(f"{str(r[i] if i < len(r) and r[i] is not None else '').ljust(col_widths[i])}" for i in range(len(headers))) + " |"
        print(row_line)
    print(separator)


def parse_json_or_list(value_str: str) -> List[Any]:
    """Parses a string that can either be JSON or comma-separated values."""
    value_str = value_str.strip()
    try:
        data = json.loads(value_str)
        if isinstance(data, list):
            return data
        return [data]
    except json.JSONDecodeError:
        # Fallback to comma-separated values, stripping quotes
        parts = [p.strip().strip("'\"") for p in value_str.split(",")]
        # Attempt simple type coercion (int, float)
        parsed = []
        for p in parts:
            try:
                parsed.append(int(p))
            except ValueError:
                try:
                    parsed.append(float(p))
                except ValueError:
                    parsed.append(p)
        return parsed


def cmd_info(client: GraphExcelClient, args: argparse.Namespace):
    """Displays drive metadata and file information."""
    print("=" * 60)
    print("Microsoft Graph Configuration & Target Status")
    print("=" * 60)
    print(f"Auth Flow       : {Config.AUTH_FLOW}")
    print(f"Azure Client ID : {Config.AZURE_CLIENT_ID[:8]}... (configured)" if Config.AZURE_CLIENT_ID else "Client ID: Not set")
    print(f"Tenant ID       : {Config.AZURE_TENANT_ID}")
    print(f"Drive Target    : {Config.DRIVE_TARGET}")
    print(f"Target File     : {args.file or Config.EXCEL_FILE_PATH}")
    print("-" * 60)

    print("Checking remote file status...")
    meta = client.get_file_metadata(args.file)
    if not meta:
        print("[Status] File does not exist in OneDrive/SharePoint yet.")
        print("Run 'python main.py ensure-file' to create an empty workbook.")
        return

    print("[Status] File found successfully!")
    print(f"Name            : {meta.get('name')}")
    print(f"ID              : {meta.get('id')}")
    print(f"Size            : {meta.get('size', 0):,} bytes")
    print(f"Last Modified   : {meta.get('lastModifiedDateTime')}")
    print(f"Web URL         : {meta.get('webUrl')}")


def cmd_ensure_file(client: GraphExcelClient, args: argparse.Namespace):
    """Ensures file exists; uploads blank Excel workbook if missing."""
    meta = client.ensure_file_exists(args.file)
    print(f"Target Excel file ready: {meta.get('name')} (ID: {meta.get('id')})")
    if meta.get("webUrl"):
        print(f"Direct link: {meta.get('webUrl')}")


def cmd_list_sheets(client: GraphExcelClient, args: argparse.Namespace):
    """Lists worksheets in the target Excel file."""
    sheets = client.list_worksheets(args.file)
    if not sheets:
        print("No worksheets found.")
        return
    headers = ["Name", "ID", "Position", "Visibility"]
    rows = [[s.get("name"), s.get("id"), s.get("position"), s.get("visibility")] for s in sheets]
    print(f"Worksheets in '{args.file or Config.EXCEL_FILE_PATH}':")
    print_table(headers, rows)


def cmd_add_sheet(client: GraphExcelClient, args: argparse.Namespace):
    """Adds a new worksheet to the workbook."""
    result = client.add_worksheet(name=args.name, file_path=args.file)
    print(f"[Success] Worksheet '{args.name}' created with ID: {result.get('id')}")


def cmd_read_range(client: GraphExcelClient, args: argparse.Namespace):
    """Reads and displays a cell range or used range from a worksheet."""
    sheet = args.sheet or "Sheet1"
    if args.address:
        print(f"Reading range '{sheet}!{args.address}'...")
        res = client.read_range(sheet_name=sheet, address=args.address, file_path=args.file)
    else:
        print(f"Reading used range in sheet '{sheet}'...")
        res = client.read_used_range(sheet_name=sheet, file_path=args.file)
        if not res:
            print(f"Worksheet '{sheet}' is empty.")
            return

    address = res.get("address", "")
    values = res.get("values", [])
    print(f"Range: {address}")
    if not values:
        print("No cell values found in this range.")
        return

    # Print raw grid
    for r_idx, row in enumerate(values):
        print(f"Row {r_idx + 1:02d}: {row}")


def cmd_update_range(client: GraphExcelClient, args: argparse.Namespace):
    """Updates a specific cell range with provided values."""
    sheet = args.sheet or "Sheet1"
    address = args.address

    if args.json_file:
        with open(args.json_file, "r", encoding="utf-8") as f:
            values = json.load(f)
    elif args.values:
        parsed = json.loads(args.values)
        if not isinstance(parsed, list):
            raise ValueError("--values must be a 2D JSON array (e.g. '[[\"A\", \"B\"]]')")
        if parsed and not isinstance(parsed[0], list):
            values = [parsed]
        else:
            values = parsed
    else:
        print("[Error] Please specify --values (2D JSON array) or --json-file.")
        sys.exit(1)

    print(f"Updating '{sheet}!{address}' with {len(values)} row(s)...")
    res = client.update_range(sheet_name=sheet, address=address, values=values, file_path=args.file)
    print(f"[Success] Updated range '{res.get('address')}'.")


def cmd_append_row(client: GraphExcelClient, args: argparse.Namespace):
    """Appends a row to the end of the worksheet or an Excel Table."""
    sheet = args.sheet or "Sheet1"

    if args.json_file:
        with open(args.json_file, "r", encoding="utf-8") as f:
            row_values = json.load(f)
    elif args.values:
        row_values = parse_json_or_list(args.values)
    else:
        print("[Error] Please specify --values or --json-file.")
        sys.exit(1)

    if not isinstance(row_values, list):
        row_values = [row_values]

    if args.table:
        print(f"Appending row to table '{args.table}'...")
        res = client.append_table_row(table_name_or_id=args.table, row_values=row_values, file_path=args.file)
        print(f"[Success] Row appended to table '{args.table}'. Values: {row_values}")
    else:
        print(f"Appending row to sheet '{sheet}'...")
        res = client.append_row_to_range(
            row_values=row_values,
            sheet_name=sheet,
            start_col=args.start_col or "A",
            file_path=args.file,
        )
        print(f"[Success] Row appended at address '{res.get('address')}'. Values: {row_values}")


def cmd_create_table(client: GraphExcelClient, args: argparse.Namespace):
    """Converts a cell range into an Excel Table."""
    sheet = args.sheet or "Sheet1"
    print(f"Creating Excel Table on '{sheet}!{args.address}'...")
    res = client.create_table(
        address=args.address,
        table_name=args.name,
        sheet_name=sheet,
        has_headers=not args.no_headers,
        file_path=args.file,
    )
    print(f"[Success] Table created. Name: '{res.get('name')}', ID: {res.get('id')}")


def cmd_list_tables(client: GraphExcelClient, args: argparse.Namespace):
    """Lists tables in the workbook."""
    tables = client.list_tables(sheet_name=args.sheet, file_path=args.file)
    if not tables:
        print("No tables found.")
        return
    headers = ["Name", "ID", "Show Headers", "Style"]
    rows = [[t.get("name"), t.get("id"), t.get("showHeaders"), t.get("style")] for t in tables]
    print_table(headers, rows)


def cmd_clear_range(client: GraphExcelClient, args: argparse.Namespace):
    """Clears cells in a specified range."""
    sheet = args.sheet or "Sheet1"
    print(f"Clearing range '{sheet}!{args.address}' ({args.apply_to})...")
    client.clear_range(
        sheet_name=sheet,
        address=args.address,
        apply_to=args.apply_to,
        file_path=args.file,
    )
    print(f"[Success] Range '{sheet}!{args.address}' cleared.")


def cmd_auth_login(auth: GraphAuthenticator):
    """Triggers authentication to populate the token cache."""
    print("Initiating authentication check...")
    token = auth.get_access_token()
    print("[Success] Authentication token acquired successfully.")


def cmd_auth_logout(auth: GraphAuthenticator):
    """Clears cached tokens."""
    auth.clear_cache()
    print("[Success] Local authentication cache has been removed.")


def build_parser() -> argparse.ArgumentParser:
    """Configures the command-line argument parser."""
    common_file_parser = argparse.ArgumentParser(add_help=False)
    common_file_parser.add_argument(
        "-f", "--file",
        dest="file",
        help="Target Excel file path in OneDrive/SharePoint (overrides EXCEL_FILE_PATH in .env)",
        default=None,
    )

    parser = argparse.ArgumentParser(
        prog="python main.py",
        description="Headless CLI to update and query Excel files in OneDrive/SharePoint via MS Graph API.",
        parents=[common_file_parser],
    )

    subparsers = parser.add_subparsers(dest="command", help="Available subcommands")

    # info
    subparsers.add_parser("info", parents=[common_file_parser], help="Display drive and target file information")

    # ensure-file
    subparsers.add_parser("ensure-file", parents=[common_file_parser], help="Verify target Excel file exists or create a blank one")

    # list-sheets
    subparsers.add_parser("list-sheets", parents=[common_file_parser], help="List all worksheets in the workbook")

    # add-sheet
    p_add_sheet = subparsers.add_parser("add-sheet", parents=[common_file_parser], help="Add a new worksheet")
    p_add_sheet.add_argument("--name", "-n", required=True, help="Name of the new worksheet")

    # read-range
    p_read = subparsers.add_parser("read-range", parents=[common_file_parser], help="Read data from a worksheet range")
    p_read.add_argument("--sheet", "-s", default="Sheet1", help="Worksheet name (default: Sheet1)")
    p_read.add_argument("--address", "-a", default=None, help="Cell address (e.g. A1:D10). If omitted, reads usedRange.")

    # update-range
    p_update = subparsers.add_parser("update-range", parents=[common_file_parser], help="Update values in a specified cell range")
    p_update.add_argument("--sheet", "-s", default="Sheet1", help="Worksheet name (default: Sheet1)")
    p_update.add_argument("--address", "-a", required=True, help="Target range address (e.g. A1:B2)")
    p_update.add_argument("--values", "-v", help="2D JSON array of values (e.g. '[[\"Name\", \"Score\"], [\"Alice\", 95]]')")
    p_update.add_argument("--json-file", help="Path to local JSON file containing 2D array of values")

    # append-row
    p_append = subparsers.add_parser("append-row", parents=[common_file_parser], help="Append a row to the end of the sheet or table")
    p_append.add_argument("--sheet", "-s", default="Sheet1", help="Worksheet name (default: Sheet1)")
    p_append.add_argument("--values", "-v", required=True, help="JSON list or comma-separated values (e.g. '[\"2026-11-24\", \"Metric\", 42]')")
    p_append.add_argument("--start-col", default="A", help="Starting column letter for sheet append (default: A)")
    p_append.add_argument("--table", "-t", default=None, help="Excel Table name/ID (if appending to an existing table)")
    p_append.add_argument("--json-file", help="Path to local JSON file containing list of row values")

    # list-tables
    p_ltables = subparsers.add_parser("list-tables", parents=[common_file_parser], help="List tables in the workbook")
    p_ltables.add_argument("--sheet", "-s", default=None, help="Optional filter by worksheet")

    # create-table
    p_ctable = subparsers.add_parser("create-table", parents=[common_file_parser], help="Create an Excel Table over a range")
    p_ctable.add_argument("--sheet", "-s", default="Sheet1", help="Worksheet name (default: Sheet1)")
    p_ctable.add_argument("--address", "-a", required=True, help="Range address (e.g. A1:D20)")
    p_ctable.add_argument("--name", "-n", default=None, help="Custom table name (e.g. 'SalesData')")
    p_ctable.add_argument("--no-headers", action="store_true", help="Set if range does NOT contain a header row")

    # clear-range
    p_clear = subparsers.add_parser("clear-range", parents=[common_file_parser], help="Clear cell contents or formats in a range")
    p_clear.add_argument("--sheet", "-s", default="Sheet1", help="Worksheet name (default: Sheet1)")
    p_clear.add_argument("--address", "-a", required=True, help="Cell range to clear (e.g. A2:D100)")
    p_clear.add_argument("--apply-to", choices=["All", "Contents", "Formats"], default="All", help="Clear type (default: All)")

    # auth commands
    subparsers.add_parser("auth-login", help="Authenticate with Microsoft Graph and populate token cache")
    subparsers.add_parser("auth-logout", help="Clear local token cache")

    return parser


def main():
    """Main entrypoint for CLI execution."""
    parser = build_parser()
    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(0)

    auth = GraphAuthenticator()

    if args.command == "auth-login":
        cmd_auth_login(auth)
        return
    elif args.command == "auth-logout":
        cmd_auth_logout(auth)
        return

    # All other commands require the Graph client
    try:
        client = GraphExcelClient(authenticator=auth)
    except Exception as e:
        print(f"[Error] Failed to initialize Graph client: {e}")
        sys.exit(1)

    try:
        if args.command == "info":
            cmd_info(client, args)
        elif args.command == "ensure-file":
            cmd_ensure_file(client, args)
        elif args.command == "list-sheets":
            cmd_list_sheets(client, args)
        elif args.command == "add-sheet":
            cmd_add_sheet(client, args)
        elif args.command == "read-range":
            cmd_read_range(client, args)
        elif args.command == "update-range":
            cmd_update_range(client, args)
        elif args.command == "append-row":
            cmd_append_row(client, args)
        elif args.command == "list-tables":
            cmd_list_tables(client, args)
        elif args.command == "create-table":
            cmd_create_table(client, args)
        elif args.command == "clear-range":
            cmd_clear_range(client, args)
        else:
            parser.print_help()
    except GraphAPIError as gerr:
        print(f"\n[Graph API Error] Status Code {gerr.status_code}")
        print(f"Error Code    : {gerr.code}")
        print(f"Message       : {gerr.message}")
        sys.exit(1)
    except Exception as exc:
        print(f"\n[Unexpected Error] {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
