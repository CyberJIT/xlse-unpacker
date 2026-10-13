#!/usr/bin/env python3
"""
test_xlse_unpacker.py

Unit tests for xlse_unpacker: directory traversal, file discovery, delta sync,
and pure stdlib MS-OVBA decompression.
"""

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
            "xl/workbook.xml": "<workbook/>",
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
        # Build an uncompressed MS-OVBA chunk:
        # byte 0: 0x01 (sig)
        # bytes 1-2: chunk header: (size-3) | (sig=0b011 << 12) | (compressed=0 << 15)
        # payload: "Sub Test()\nEnd Sub\n" (20 bytes) -> chunk_size = 20 + 2 = 22 -> header = (22-3) | (3<<12) = 19 | 0x3000 = 0x3013
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


if __name__ == "__main__":
    unittest.main()
