"""Unit tests for MS Graph Excel integration components."""

import io
import unittest
import zipfile
from config import Config
from graph_client import GraphExcelClient, create_minimal_xlsx
from main import build_parser, parse_json_or_list


class TestGraphIntegration(unittest.TestCase):
    """Test suite for config, parser, client helpers, and XLSX generator."""

    def test_minimal_xlsx_structure(self):
        """Verifies that create_minimal_xlsx generates a valid zip with required OpenXML files."""
        xlsx_bytes = create_minimal_xlsx()
        self.assertGreater(len(xlsx_bytes), 0)

        buf = io.BytesIO(xlsx_bytes)
        with zipfile.ZipFile(buf, "r") as z:
            file_names = z.namelist()
            self.assertIn("[Content_Types].xml", file_names)
            self.assertIn("_rels/.rels", file_names)
            self.assertIn("xl/workbook.xml", file_names)
            self.assertIn("xl/worksheets/sheet1.xml", file_names)
            self.assertIn("xl/_rels/workbook.xml.rels", file_names)

    def test_parse_json_or_list(self):
        """Verifies parsing of JSON strings and comma-separated lists."""
        # JSON array with numbers and strings
        json_input = '["2026-11-24", "Metric", 42, 99.5]'
        parsed = parse_json_or_list(json_input)
        self.assertEqual(parsed, ["2026-11-24", "Metric", 42, 99.5])

        # Comma-separated strings
        csv_input = "Alice, 100, 45.5"
        parsed_csv = parse_json_or_list(csv_input)
        self.assertEqual(parsed_csv, ["Alice", 100, 45.5])

    def test_endpoint_resolution(self):
        """Tests endpoint building logic in GraphExcelClient."""
        # Mock client without network auth
        client = GraphExcelClient.__new__(GraphExcelClient)
        normalized = client._normalize_file_path("my-app\\folder\\test.xlsx")
        self.assertEqual(normalized, "my-app/folder/test.xlsx")

        # Automatically appends .xlsx if omitted
        normalized_no_ext = client._normalize_file_path("my_folder/report")
        self.assertEqual(normalized_no_ext, "my_folder/report.xlsx")

    def test_argument_parser(self):
        """Verifies CLI argument parsing across subcommands."""
        parser = build_parser()

        # info command
        args = parser.parse_args(["info", "--file", "my_path/book.xlsx"])
        self.assertEqual(args.command, "info")
        self.assertEqual(args.file, "my_path/book.xlsx")

        # append-row command
        args = parser.parse_args(["append-row", "--values", '["foo", "bar"]', "--sheet", "Sheet2"])
        self.assertEqual(args.command, "append-row")
        self.assertEqual(args.values, '["foo", "bar"]')
        self.assertEqual(args.sheet, "Sheet2")

        # update-range command
        args = parser.parse_args(["update-range", "-a", "A1:B2", "-v", "[[1, 2]]"])
        self.assertEqual(args.command, "update-range")
        self.assertEqual(args.address, "A1:B2")
        self.assertEqual(args.values, "[[1, 2]]")

        # create-table command
        args = parser.parse_args(["create-table", "-a", "A1:C10", "-n", "LogsTable"])
        self.assertEqual(args.command, "create-table")
        self.assertEqual(args.address, "A1:C10")
        self.assertEqual(args.name, "LogsTable")


if __name__ == "__main__":
    unittest.main()
