"""
openxml_parser.py

Zero-external-dependency parser for Excel OpenXML structures:
Extracts sheet definitions, cell formulas, defined names (named ranges), structured tables,
and pivot cache summaries into searchable human-readable JSON files.
"""

import json
import os
import re
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional


def _strip_ns(tag: str) -> str:
    """Remove XML namespace prefix from tag."""
    return tag.split("}", 1)[1] if "}" in tag else tag


def parse_workbook_xml(workbook_xml_path: str) -> Dict[str, Any]:
    """
    Parse xl/workbook.xml to extract sheets and defined names.
    Returns:
        Dict with 'sheets' (list) and 'defined_names' (list).
    """
    res: Dict[str, Any] = {"sheets": [], "defined_names": []}
    if not os.path.exists(workbook_xml_path):
        return res

    try:
        tree = ET.parse(workbook_xml_path)
        root = tree.getroot()

        # Sheets
        for elem in root.iter():
            tag = _strip_ns(elem.tag)
            if tag == "sheet":
                name = elem.attrib.get("name", "")
                sheet_id = elem.attrib.get("sheetId", "")
                r_id = elem.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                state = elem.attrib.get("state", "visible")
                res["sheets"].append({
                    "name": name,
                    "sheet_id": sheet_id,
                    "r_id": r_id,
                    "state": state,
                })
            elif tag == "definedName":
                d_name = elem.attrib.get("name", "")
                local_sheet_id = elem.attrib.get("localSheetId")
                comment = elem.attrib.get("comment", "")
                formula = elem.text.strip() if elem.text else ""
                res["defined_names"].append({
                    "name": d_name,
                    "local_sheet_id": local_sheet_id,
                    "formula": formula,
                    "comment": comment,
                })
    except Exception:
        pass

    return res


def parse_worksheet_formulas(worksheet_xml_path: str) -> Dict[str, str]:
    """
    Scan a worksheet XML file for cells containing formulas (<f> tags).
    Returns:
        Dict mapping cell_ref (e.g. 'B5') -> formula string (e.g. 'SUM(A1:A4)').
    """
    formulas: Dict[str, str] = {}
    if not os.path.exists(worksheet_xml_path):
        return formulas

    try:
        # Use iterative parsing to avoid loading huge sheets entirely into memory
        context = ET.iterparse(worksheet_xml_path, events=("end",))
        for event, elem in context:
            tag = _strip_ns(elem.tag)
            if tag == "c":  # Cell
                cell_ref = elem.attrib.get("r", "")
                # Find child <f> tag
                f_elem = None
                for child in elem:
                    if _strip_ns(child.tag) == "f":
                        f_elem = child
                        break
                if f_elem is not None and f_elem.text:
                    formula_text = f_elem.text.strip()
                    formulas[cell_ref] = formula_text

                elem.clear()
    except Exception:
        pass

    return formulas


def parse_tables(tables_dir: str) -> List[Dict[str, Any]]:
    """Parse table definitions in xl/tables/table*.xml."""
    tables = []
    if not os.path.exists(tables_dir):
        return tables

    for f in sorted(os.listdir(tables_dir)):
        if f.lower().endswith(".xml"):
            full_path = os.path.join(tables_dir, f)
            try:
                tree = ET.parse(full_path)
                root = tree.getroot()
                tbl_name = root.attrib.get("name", "")
                display_name = root.attrib.get("displayName", "")
                ref = root.attrib.get("ref", "")
                columns = []
                for col in root.iter():
                    if _strip_ns(col.tag) == "tableColumn":
                        columns.append(col.attrib.get("name", ""))
                tables.append({
                    "name": tbl_name,
                    "display_name": display_name,
                    "ref": ref,
                    "columns": columns,
                    "file": f,
                })
            except Exception:
                continue

    return tables


def parse_pivot_caches(pivot_dir: str) -> List[Dict[str, Any]]:
    """Parse pivot cache definitions in xl/pivotCache/pivotCacheDefinition*.xml."""
    caches = []
    if not os.path.exists(pivot_dir):
        return caches

    for f in sorted(os.listdir(pivot_dir)):
        if f.lower().endswith(".xml"):
            full_path = os.path.join(pivot_dir, f)
            try:
                tree = ET.parse(full_path)
                root = tree.getroot()
                fields = []
                for f_elem in root.iter():
                    if _strip_ns(f_elem.tag) == "cacheField":
                        fields.append(f_elem.attrib.get("name", ""))
                caches.append({
                    "file": f,
                    "field_count": len(fields),
                    "fields": fields,
                })
            except Exception:
                continue

    return caches


def deconstruct_workbook_structure(unpacked_dir: str) -> Dict[str, Any]:
    """
    Deconstruct the unpacked OpenXML directory into Workbook_Structure/.
    Returns a summary dictionary of extracted structure items.
    """
    struct_out = os.path.join(unpacked_dir, "Workbook_Structure")
    os.makedirs(struct_out, exist_ok=True)

    xl_dir = os.path.join(unpacked_dir, "xl")
    wb_xml = os.path.join(xl_dir, "workbook.xml")

    # 1. Workbook and Sheet metadata
    wb_data = parse_workbook_xml(wb_xml)
    sheets = wb_data.get("sheets", [])
    defined_names = wb_data.get("defined_names", [])

    with open(os.path.join(struct_out, "Sheets.json"), "w", encoding="utf-8") as f:
        json.dump(sheets, f, indent=2, ensure_ascii=False)

    with open(os.path.join(struct_out, "DefinedNames.json"), "w", encoding="utf-8") as f:
        json.dump(defined_names, f, indent=2, ensure_ascii=False)

    # 2. Sheet Formulas
    sheets_dir = os.path.join(xl_dir, "worksheets")
    formulas_summary: Dict[str, Any] = {}
    total_formulas = 0

    # Build map from sheet filename to sheet name
    sheet_name_map = {}
    for idx, s in enumerate(sheets, start=1):
        sheet_name_map[f"sheet{idx}.xml"] = s["name"]

    if os.path.exists(sheets_dir):
        for s_file in sorted(os.listdir(sheets_dir)):
            if s_file.lower().endswith(".xml"):
                full_s_path = os.path.join(sheets_dir, s_file)
                f_map = parse_worksheet_formulas(full_s_path)
                sheet_label = sheet_name_map.get(s_file, s_file)
                if f_map:
                    formulas_summary[sheet_label] = {
                        "formula_count": len(f_map),
                        "formulas": f_map,
                    }
                    total_formulas += len(f_map)

    with open(os.path.join(struct_out, "Formulas_Summary.json"), "w", encoding="utf-8") as f:
        json.dump(formulas_summary, f, indent=2, ensure_ascii=False)

    # 3. Tables
    tables_dir = os.path.join(xl_dir, "tables")
    tables = parse_tables(tables_dir)
    with open(os.path.join(struct_out, "Tables.json"), "w", encoding="utf-8") as f:
        json.dump(tables, f, indent=2, ensure_ascii=False)

    # 4. Pivot caches
    pivot_dir = os.path.join(xl_dir, "pivotCache")
    pivot_caches = parse_pivot_caches(pivot_dir)
    with open(os.path.join(struct_out, "PivotTables_Summary.json"), "w", encoding="utf-8") as f:
        json.dump(pivot_caches, f, indent=2, ensure_ascii=False)

    return {
        "sheets_count": len(sheets),
        "defined_names_count": len(defined_names),
        "total_formulas": total_formulas,
        "tables_count": len(tables),
        "pivot_caches_count": len(pivot_caches),
    }
