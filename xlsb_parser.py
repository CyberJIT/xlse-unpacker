"""
xlsb_parser.py

Zero-external-dependency parser for Excel Binary Workbook (.xlsb) structures.
Parses Biff12 variable-length record streams (Record Type + Record Length)
to extract workbook sheet names, states, and structure without third-party libraries.
"""

import json
import os
import struct
from typing import Any, Dict, List, Tuple


def read_biff12_record_header(data: bytes, offset: int) -> Tuple[int, int, int]:
    """
    Read Biff12 record type and length at offset using variable-length integer encoding.
    Returns:
        (record_type, record_length, bytes_consumed_by_header)
    """
    total = len(data)
    if offset >= total:
        return (-1, 0, 0)

    curr = offset

    # 1. Read Record Type (1 or 2 bytes)
    b = data[curr]
    curr += 1
    rec_type = b & 0x7F
    if b & 0x80:
        if curr >= total:
            return (-1, 0, 0)
        b2 = data[curr]
        curr += 1
        rec_type |= (b2 & 0x7F) << 7

    # 2. Read Record Length (1 to 4 bytes)
    rec_len = 0
    shift = 0
    for _ in range(4):
        if curr >= total:
            return (-1, 0, 0)
        lb = data[curr]
        curr += 1
        rec_len |= (lb & 0x7F) << shift
        shift += 7
        if not (lb & 0x80):
            break

    header_len = curr - offset
    return (rec_type, rec_len, header_len)


# Common Biff12 Record Types [MS-XLSB]
# BrtBundleSh = 0x009C (Sheet definition)
# BrtBeginBook = 0x0089
# BrtEndBook = 0x008A
# BrtDefName = 0x0027 (Defined Name)
RECORD_BrtBundleSh = 0x009C
RECORD_BrtDefName = 0x0027


def parse_xlsb_workbook_bin(workbook_bin_path: str) -> Dict[str, Any]:
    """
    Parse xl/workbook.bin to extract sheet names and basic properties.
    """
    res: Dict[str, Any] = {"sheets": [], "defined_names": []}
    if not os.path.exists(workbook_bin_path):
        return res

    try:
        with open(workbook_bin_path, "rb") as f:
            data = f.read()

        pos = 0
        total = len(data)

        while pos < total:
            rec_type, rec_len, header_len = read_biff12_record_header(data, pos)
            if rec_type == -1 or pos + header_len + rec_len > total:
                break

            rec_data_offset = pos + header_len
            rec_data = data[rec_data_offset : rec_data_offset + rec_len]
            pos = rec_data_offset + rec_len

            # BrtBundleSh Record:
            # 4 bytes hsState (0 = visible, 1 = hidden, 2 = veryHidden)
            # 4 bytes iTabID
            # 4 bytes strRelID length, followed by utf-16LE characters
            # followed by sheet name length and utf-16LE characters
            if rec_type == RECORD_BrtBundleSh and len(rec_data) >= 8:
                hs_state, tab_id = struct.unpack_from("<II", rec_data, 0)
                state_map = {0: "visible", 1: "hidden", 2: "veryHidden"}
                state_str = state_map.get(hs_state, "visible")

                curr_rec = 8
                # Read rId string (4-byte length + UTF-16LE bytes)
                if curr_rec + 4 <= len(rec_data):
                    rel_id_char_len = struct.unpack_from("<I", rec_data, curr_rec)[0]
                    curr_rec += 4
                    rel_id_byte_len = rel_id_char_len * 2
                    r_id = rec_data[curr_rec : curr_rec + rel_id_byte_len].decode("utf-16le", errors="replace")
                    curr_rec += rel_id_byte_len

                    # Read sheet name string
                    if curr_rec + 4 <= len(rec_data):
                        name_char_len = struct.unpack_from("<I", rec_data, curr_rec)[0]
                        curr_rec += 4
                        name_byte_len = name_char_len * 2
                        sheet_name = rec_data[curr_rec : curr_rec + name_byte_len].decode("utf-16le", errors="replace")
                    else:
                        sheet_name = f"Sheet{tab_id}"
                else:
                    r_id = ""
                    sheet_name = f"Sheet{tab_id}"

                res["sheets"].append({
                    "name": sheet_name,
                    "sheet_id": str(tab_id),
                    "r_id": r_id,
                    "state": state_str,
                })

    except Exception:
        pass

    return res


def deconstruct_xlsb_structure(unpacked_dir: str) -> Dict[str, Any]:
    """
    Deconstruct unpacked XLSB binary files in unpacked_dir into Workbook_Structure/.
    """
    struct_out = os.path.join(unpacked_dir, "Workbook_Structure")
    os.makedirs(struct_out, exist_ok=True)

    xl_dir = os.path.join(unpacked_dir, "xl")
    wb_bin = os.path.join(xl_dir, "workbook.bin")

    wb_data = parse_xlsb_workbook_bin(wb_bin)
    sheets = wb_data.get("sheets", [])

    with open(os.path.join(struct_out, "Sheets.json"), "w", encoding="utf-8") as f:
        json.dump(sheets, f, indent=2, ensure_ascii=False)

    return {
        "sheets_count": len(sheets),
        "format": "XLSB",
    }
