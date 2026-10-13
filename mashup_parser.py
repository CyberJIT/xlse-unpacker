"""
mashup_parser.py

Module to parse and extract Microsoft Power Query DataMashup binary streams (MS-QDEFF specification)
and customXml Power Query definitions within Excel workbooks into plain text Power Query M code,
individual query files, and XML/JSON configuration settings.
"""

import base64
import hashlib
import io
import json
import os
import re
import struct
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Dict, List, Optional, Tuple


def pretty_format_xml(xml_bytes_or_str) -> str:
    """Format XML string or bytes with readable indentations."""
    try:
        raw_str = xml_bytes_or_str.decode("utf-8") if isinstance(xml_bytes_or_str, bytes) else xml_bytes_or_str
        elem = ET.fromstring(raw_str)
        rough = ET.tostring(elem, encoding="utf-8")
        reparsed = minidom.parseString(rough)
        return reparsed.toprettyxml(indent="  ")
    except Exception:
        if isinstance(xml_bytes_or_str, bytes):
            return xml_bytes_or_str.decode("utf-8", errors="replace")
        return str(xml_bytes_or_str)


def parse_m_queries(section_text: str) -> Dict[str, str]:
    """
    Decompose a Power Query M Section document (such as Formulas/Section1.m)
    into individual shared query expressions.
    Returns:
        Dict mapping query_name -> M expression string.
    """
    queries: Dict[str, str] = {}
    if not section_text or not section_text.strip():
        return queries

    pattern = re.compile(
        r"(?:shared\s+)?(?:#\"([^\"]+)\"|([a-zA-Z0-9_\.\-]+))\s*=\s*(.+?)(?=(?:;\s*(?:shared\s+)?(?:#\"[^\"]+\"|[a-zA-Z0-9_\.\-]+)\s*=)|\Z)",
        re.DOTALL,
    )

    for match in pattern.finditer(section_text):
        quoted_name, bare_name, expression = match.groups()
        query_name = quoted_name if quoted_name is not None else bare_name
        clean_expr = expression.strip()
        if clean_expr.endswith(";"):
            clean_expr = clean_expr[:-1].rstrip()
        queries[query_name] = clean_expr

    return queries


def extract_datamashup_package(mashup_bytes: bytes) -> Dict[str, Any]:
    """
    Extract embedded components from a Microsoft DataMashup binary package (MS-QDEFF).
    Returns a dict with extracted items:
        'package_zip_entries': dict of {filename: bytes}
        'queries': dict of {query_name: m_expression_code}
        'permissions': str (pretty xml/json if present)
        'metadata': str (pretty xml if present)
    """
    result: Dict[str, Any] = {
        "package_zip_entries": {},
        "queries": {},
        "permissions": None,
        "metadata": None,
    }

    if not mashup_bytes or len(mashup_bytes) < 8:
        return result

    # Find inner zip package
    zip_offset = mashup_bytes.find(b"PK\x03\x04")
    if zip_offset == -1:
        return result

    zip_data = mashup_bytes[zip_offset:]

    try:
        with zipfile.ZipFile(io.BytesIO(zip_data), "r") as zf:
            for name in zf.namelist():
                content = zf.read(name)
                result["package_zip_entries"][name] = content

                if name.lower() == "formulas/section1.m":
                    try:
                        m_text = content.decode("utf-8-sig")
                    except UnicodeDecodeError:
                        m_text = content.decode("utf-8", errors="replace")
                    result["queries"] = parse_m_queries(m_text)
                    result["section1_raw"] = m_text

                elif name.lower() == "config/package.xml":
                    result["package_xml"] = pretty_format_xml(content)
                elif name.lower() == "permissions.xml":
                    result["permissions"] = pretty_format_xml(content)
                elif name.lower() == "metadata.xml":
                    result["metadata"] = pretty_format_xml(content)

    except Exception:
        pass

    return result


def find_and_extract_power_query(unpacked_dir: str) -> Optional[Dict[str, Any]]:
    """
    Scan unpacked Excel files (customXml, xl/customData, etc.) for Power Query mashup data.
    """
    # 1. Search in customXml/item*.xml
    custom_xml_dir = os.path.join(unpacked_dir, "customXml")
    if os.path.exists(custom_xml_dir):
        for entry in os.listdir(custom_xml_dir):
            if entry.lower().endswith(".xml"):
                xml_path = os.path.join(custom_xml_dir, entry)
                try:
                    with open(xml_path, "rb") as f:
                        data = f.read()

                    # Check for base64 encoded DataMashup or direct DataMashup
                    if b"DataMashup" in data or b"PK\x03\x04" in data:
                        # Attempt to locate zip signature directly
                        res = extract_datamashup_package(data)
                        if res.get("queries"):
                            return res

                        # Or base64 embedded tag
                        m = re.search(rb">([A-Za-z0-9+/=]{100,})<", data)
                        if m:
                            try:
                                decoded = base64.b64decode(m.group(1))
                                res = extract_datamashup_package(decoded)
                                if res.get("queries"):
                                    return res
                            except Exception:
                                pass
                except Exception:
                    continue

    # 2. Check for binary stream in xl/customData/ or customXml/item*.bin
    for root, _, files in os.walk(unpacked_dir):
        for f in files:
            if f.lower().endswith(".bin") or "mashup" in f.lower():
                bin_path = os.path.join(root, f)
                try:
                    with open(bin_path, "rb") as fp:
                        raw = fp.read()
                    if b"PK\x03\x04" in raw:
                        res = extract_datamashup_package(raw)
                        if res.get("queries"):
                            return res
                except Exception:
                    continue

    return None


def export_power_query_to_directory(mashup_data: Dict[str, Any], target_dir: str) -> int:
    """
    Save decompiled Power Query queries and configs to target_dir/PowerQuery_Extracted/.
    Returns count of queries extracted.
    """
    pq_out = os.path.join(target_dir, "PowerQuery_Extracted")
    os.makedirs(pq_out, exist_ok=True)

    queries_dir = os.path.join(pq_out, "Queries")
    os.makedirs(queries_dir, exist_ok=True)

    queries = mashup_data.get("queries", {})
    for q_name, q_code in queries.items():
        safe_name = re.sub(r'[\\/*?:"<>|]', "_", q_name)
        q_path = os.path.join(queries_dir, f"{safe_name}.m")
        with open(q_path, "w", encoding="utf-8") as f:
            f.write(q_code)

    # Save summary
    summary_path = os.path.join(pq_out, "Queries_Summary.json")
    with open(summary_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "query_count": len(queries),
                "query_names": list(queries.keys()),
            },
            f,
            indent=2,
        )

    # Save package xmls if present
    if mashup_data.get("package_xml"):
        with open(os.path.join(pq_out, "Package.xml"), "w", encoding="utf-8") as f:
            f.write(mashup_data["package_xml"])

    if mashup_data.get("metadata"):
        with open(os.path.join(pq_out, "Metadata.xml"), "w", encoding="utf-8") as f:
            f.write(mashup_data["metadata"])

    if mashup_data.get("permissions"):
        with open(os.path.join(pq_out, "Permissions.xml"), "w", encoding="utf-8") as f:
            f.write(mashup_data["permissions"])

    return len(queries)
