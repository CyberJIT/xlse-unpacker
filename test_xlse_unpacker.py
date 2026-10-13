#!/usr/bin/env python3
"""
test_xlse_unpacker.py

Unit tests for xlse_unpacker: directory traversal, file discovery, delta sync,
MS-OVBA decompression, Power Query M parsing, and OpenXML structure deconstruction.
"""

import io
import os
import shutil
import struct
import tempfile
import unittest
import zipfile

from xlse_unpacker import (
    find_excel_files,
    matches_patterns,
    unpack_excel_file,
    compute_sha256,
)
from vba_parser import decompress_msovba, export_vba_to_directory
from mashup_parser import (
    parse_m_queries,
    extract_datamashup_package,
    export_power_query_to_directory,
)
from openxml_parser import (
    parse_workbook_xml,
    parse_worksheet_formulas,
    deconstruct_workbook_structure,
)


class TestXlseUnpackerCore(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="xlse_test_")

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def _create_mock_zip(self, file_path: str, entries: dict):
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with zipfile.ZipFile(file_path, "w") as zf:
            for arcname, data in entries.items():
                if isinstance(data, str):
                    data = data.encode("utf-8")
                zf.writestr(arcname, data)

    def test_find_excel_files_and_filtering(self):
        sub1 = os.path.join(self.test_dir, "sub1")
        sub2 = os.path.join(self.test_dir, "ignore_dir")
        os.makedirs(sub1)
        os.makedirs(sub2)

        f1 = os.path.join(self.test_dir, "report.xlsx")
        f2 = os.path.join(self.test_dir, "macro.xlsm")
        f3 = os.path.join(self.test_dir, "binary.xlsb")
        f4 = os.path.join(self.test_dir, "notes.txt")
        f5 = os.path.join(sub1, "nested.xlsx")
        f6 = os.path.join(sub2, "skip_me.xlsx")

        for f in [f1, f2, f3, f4, f5, f6]:
            with open(f, "w") as fp:
                fp.write("dummy")

        # Non-recursive
        res = find_excel_files(self.test_dir, recurse=False)
        self.assertEqual(len(res), 3)

        # Recursive
        res_rec = find_excel_files(self.test_dir, recurse=True)
        self.assertEqual(len(res_rec), 5)

        # Filter in
        res_filtered = find_excel_files(self.test_dir, recurse=True, filter_in=["*report*"])
        self.assertEqual(res_filtered, [f1])

        # Filter out dir
        res_no_dir = find_excel_files(self.test_dir, recurse=True, dir_filter_out=["*ignore*"])
        self.assertEqual(len(res_no_dir), 4)
        self.assertNotIn(f6, res_no_dir)

    def test_unpack_and_delta_sync(self):
        xlsx_file = os.path.join(self.test_dir, "Workbook1.xlsx")
        target_dir = os.path.join(self.test_dir, "Workbook1")

        self._create_mock_zip(xlsx_file, {
            "[Content_Types].xml": "<Types/>",
            "xl/workbook.xml": "<workbook><sheets><sheet name=\"Sheet1\" sheetId=\"1\"/></sheets></workbook>",
        })

        # First run: should create 2 files
        c, u, uc = unpack_excel_file(xlsx_file, quiet=True)
        self.assertEqual(c, 2)
        self.assertEqual(u, 0)
        self.assertEqual(uc, 0)
        self.assertTrue(os.path.exists(os.path.join(target_dir, "[Content_Types].xml")))

        # Second run without changes: 0 created, 0 updated, 2 unchanged
        c, u, uc = unpack_excel_file(xlsx_file, quiet=True)
        self.assertEqual(c, 0)
        self.assertEqual(u, 0)
        self.assertEqual(uc, 2)

    def test_msovba_decompression_uncompressed_chunk(self):
        payload = b"Sub Test()\nEnd Sub\n"
        chunk_header = (len(payload) + 2 - 3) | (3 << 12)
        compressed = b"\x01" + struct.pack("<H", chunk_header) + payload

        decompressed = decompress_msovba(compressed)
        self.assertEqual(decompressed, payload)

    def test_export_vba_to_directory(self):
        vba_data = {
            "project_info": {"project_name": "TestVBAProject"},
            "references": ["stdole", "Office"],
            "modules": {
                "Module1": {"type": "Module", "code": "Sub Run()\nEnd Sub\n"},
                "ThisWorkbook": {"type": "Document", "code": "Private Sub Workbook_Open()\nEnd Sub\n"},
            },
        }
        counts = export_vba_to_directory(vba_data, self.test_dir)
        self.assertEqual(counts["modules"], 1)
        self.assertEqual(counts["documents"], 1)

        mod1_path = os.path.join(self.test_dir, "VBA_Extracted", "Modules", "Module1.bas")
        doc_path = os.path.join(self.test_dir, "VBA_Extracted", "Document", "ThisWorkbook.cls")
        ref_path = os.path.join(self.test_dir, "VBA_Extracted", "References.txt")

        self.assertTrue(os.path.exists(mod1_path))
        self.assertTrue(os.path.exists(doc_path))
        self.assertTrue(os.path.exists(ref_path))

    def test_parse_m_queries_and_export(self):
        section_m = '''
section Section1;

shared #"Customer Query" = let
    Source = Csv.Document(File.Contents("C:\\data.csv")),
    #"Promoted Headers" = Table.PromoteHeaders(Source)
in
    #"Promoted Headers";

shared CleanSales = let
    Source = Sql.Database("server", "db")
in
    Source;
'''
        queries = parse_m_queries(section_m)
        self.assertIn("Customer Query", queries)
        self.assertIn("CleanSales", queries)
        self.assertTrue("Csv.Document" in queries["Customer Query"])

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("Formulas/Section1.m", section_m.encode("utf-8"))
            zf.writestr("Config/Package.xml", "<Package/>")

        mashup_bytes = b"\x00\x00\x00\x00" + buf.getvalue()
        res = extract_datamashup_package(mashup_bytes)
        self.assertEqual(len(res["queries"]), 2)

        count = export_power_query_to_directory(res, self.test_dir)
        self.assertEqual(count, 2)
        q1_file = os.path.join(self.test_dir, "PowerQuery_Extracted", "Queries", "Customer Query.m")
        self.assertTrue(os.path.exists(q1_file))

    def test_openxml_structure_deconstruction(self):
        workbook_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheets>
    <sheet name="Summary" sheetId="1" state="visible"/>
    <sheet name="Data" sheetId="2" state="visible"/>
  </sheets>
  <definedNames>
    <definedName name="TaxRate">0.21</definedName>
    <definedName name="TotalRevenue">Summary!$B$10</definedName>
  </definedNames>
</workbook>
"""
        sheet1_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1"><v>100</v></c>
      <c r="B1"><f>A1*TaxRate</f><v>21</v></c>
    </row>
  </sheetData>
</worksheet>
"""
        xl_dir = os.path.join(self.test_dir, "xl")
        os.makedirs(os.path.join(xl_dir, "worksheets"), exist_ok=True)
        with open(os.path.join(xl_dir, "workbook.xml"), "w") as f:
            f.write(workbook_xml)
        with open(os.path.join(xl_dir, "worksheets", "sheet1.xml"), "w") as f:
            f.write(sheet1_xml)

        summary = deconstruct_workbook_structure(self.test_dir)
        self.assertEqual(summary["sheets_count"], 2)
        self.assertEqual(summary["defined_names_count"], 2)
        self.assertEqual(summary["total_formulas"], 1)

        sheets_json = os.path.join(self.test_dir, "Workbook_Structure", "Sheets.json")
        formulas_json = os.path.join(self.test_dir, "Workbook_Structure", "Formulas_Summary.json")
        names_json = os.path.join(self.test_dir, "Workbook_Structure", "DefinedNames.json")

        self.assertTrue(os.path.exists(sheets_json))
        self.assertTrue(os.path.exists(formulas_json))
        self.assertTrue(os.path.exists(names_json))


if __name__ == "__main__":
    unittest.main()
