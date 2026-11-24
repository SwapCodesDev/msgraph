"""Example showing how to use GraphExcelClient programmatically in Python."""

from datetime import datetime
from graph_client import GraphExcelClient

def main():
    # Initialize client (uses .env settings)
    client = GraphExcelClient()

    target_file = "my-app/data.xlsx"
    sheet_name = "Sheet1"

    print(f"Ensuring file exists at '{target_file}'...")
    client.ensure_file_exists(target_file)

    # 1. Update Header Row
    headers = [["Timestamp", "Source", "Status", "Value"]]
    print("Writing headers...")
    client.update_range(sheet_name=sheet_name, address="A1:D1", values=headers, file_path=target_file)

    # 2. Append new log rows dynamically
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    sample_record = [now_str, "my-app Worker", "SUCCESS", 98.7]
    print(f"Appending row: {sample_record}...")
    client.append_row_to_range(row_values=sample_record, sheet_name=sheet_name, file_path=target_file)

    # 3. Read back the used range
    used = client.read_used_range(sheet_name=sheet_name, file_path=target_file)
    print("\nCurrent Sheet Data:")
    for row in used.get("values", []):
        print(" ", row)

if __name__ == "__main__":
    main()
