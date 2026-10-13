"""
vba_parser.py

Zero-external-dependency parser for Microsoft Compound File Binary Format (CFBF/OLE)
and Microsoft Office VBA compression algorithm ([MS-OVBA] section 2.4.1.3).
Extracts VBA modules, classes, forms, documents, and project metadata directly from
vbaProject.bin without requiring olefile or oletools.
"""

import io
import json
import os
import re
import struct
from typing import Dict, List, Optional, Tuple

# Header signature for Compound File Binary Format: \xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1
CFB_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"


class CFBDirectoryEntry:
    """Represents a directory entry inside a CFBF compound file."""

    def __init__(self, raw: bytes, entry_id: int):
        self.entry_id = entry_id
        # Name: 64 bytes UTF-16LE, null terminated
        name_bytes = raw[:64]
        name_len = struct.unpack_from("<H", raw, 64)[0]
        if name_len > 0:
            name_raw = name_bytes[: max(0, name_len - 2)]
            self.name = name_raw.decode("utf-16le", errors="replace")
        else:
            self.name = ""

        self.obj_type = raw[66]  # 1 = storage, 2 = stream, 5 = root
        self.color = raw[67]
        self.left_sibling = struct.unpack_from("<I", raw, 68)[0]
        self.right_sibling = struct.unpack_from("<I", raw, 72)[0]
        self.child = struct.unpack_from("<I", raw, 76)[0]
        self.start_sector = struct.unpack_from("<I", raw, 116)[0]
        self.size = struct.unpack_from("<Q", raw, 120)[0]


class CFBReader:
    """Reader for Microsoft Compound File Binary Format (MS-CFB)."""

    def __init__(self, data: bytes):
        if not data.startswith(CFB_SIGNATURE):
            raise ValueError("Not a valid Compound File Binary Format (CFBF) file")

        self.data = data
        self.sector_size = 1 << struct.unpack_from("<H", data, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from("<H", data, 32)[0]
        self.num_sat_sectors = struct.unpack_from("<I", data, 44)[0]
        self.dir_first_sector = struct.unpack_from("<I", data, 48)[0]
        self.mini_stream_cutoff = struct.unpack_from("<I", data, 56)[0]
        self.mini_fat_first_sector = struct.unpack_from("<I", data, 60)[0]
        self.num_mini_fat_sectors = struct.unpack_from("<I", data, 64)[0]
        self.difat_first_sector = struct.unpack_from("<I", data, 68)[0]
        self.num_difat_sectors = struct.unpack_from("<I", data, 72)[0]

        # Read DIFAT
        difat = list(struct.unpack_from("<109I", data, 76))
        # Handle additional DIFAT sectors if file is large (> 6.8MB)
        curr_difat = self.difat_first_sector
        while curr_difat not in (0xFFFFFFFE, 0xFFFFFFFF) and len(difat) < self.num_sat_sectors + 109:
            sec_offset = (curr_difat + 1) * self.sector_size
            sec_data = data[sec_offset : sec_offset + self.sector_size]
            entries = struct.unpack(f"<{self.sector_size // 4}I", sec_data)
            difat.extend(entries[:-1])
            curr_difat = entries[-1]

        # Read FAT (SAT)
        self.fat: List[int] = []
        for sat_sec in difat[: self.num_sat_sectors]:
            if sat_sec in (0xFFFFFFFE, 0xFFFFFFFF):
                continue
            sec_offset = (sat_sec + 1) * self.sector_size
            sec_data = data[sec_offset : sec_offset + self.sector_size]
            self.fat.extend(struct.unpack(f"<{self.sector_size // 4}I", sec_data))

        # Read Directory Stream
        dir_bytes = self._read_chain(self.dir_first_sector)
        self.entries: List[CFBDirectoryEntry] = []
        num_entries = len(dir_bytes) // 128
        for i in range(num_entries):
            raw_entry = dir_bytes[i * 128 : (i + 1) * 128]
            self.entries.append(CFBDirectoryEntry(raw_entry, i))

        # Read Mini-FAT
        if self.mini_fat_first_sector not in (0xFFFFFFFE, 0xFFFFFFFF):
            mini_fat_bytes = self._read_chain(self.mini_fat_first_sector)
            self.mini_fat = list(struct.unpack(f"<{len(mini_fat_bytes) // 4}I", mini_fat_bytes))
        else:
            self.mini_fat = []

        # Root entry contains mini stream
        if self.entries and self.entries[0].obj_type == 5:
            self.mini_stream_data = self._read_chain(self.entries[0].start_sector)
        else:
            self.mini_stream_data = b""

    def _read_chain(self, start_sector: int) -> bytes:
        """Read standard sector chain from FAT."""
        chunks = []
        curr = start_sector
        while curr not in (0xFFFFFFFE, 0xFFFFFFFF) and curr < len(self.fat):
            sec_offset = (curr + 1) * self.sector_size
            chunks.append(self.data[sec_offset : sec_offset + self.sector_size])
            curr = self.fat[curr]
        return b"".join(chunks)

    def _read_mini_chain(self, start_sector: int, size: int) -> bytes:
        """Read mini sector chain from Mini-FAT."""
        chunks = []
        curr = start_sector
        while curr not in (0xFFFFFFFE, 0xFFFFFFFF) and curr < len(self.mini_fat):
            sec_offset = curr * self.mini_sector_size
            chunks.append(self.mini_stream_data[sec_offset : sec_offset + self.mini_sector_size])
            curr = self.mini_fat[curr]
        return b"".join(chunks)[:size]

    def read_stream(self, entry: CFBDirectoryEntry) -> bytes:
        """Extract stream bytes for a given directory entry."""
        if entry.obj_type != 2:
            return b""
        if entry.size < self.mini_stream_cutoff:
            return self._read_mini_chain(entry.start_sector, entry.size)
        else:
            return self._read_chain(entry.start_sector)[: entry.size]

    def list_streams(self) -> Dict[str, bytes]:
        """Return a mapping of stream paths/names to byte content."""
        streams = {}
        for entry in self.entries:
            if entry.obj_type == 2 and entry.name:
                streams[entry.name] = self.read_stream(entry)
        return streams


def decompress_msovba(compressed_data: bytes) -> bytes:
    """
    Decompresses a byte stream compressed using the Microsoft Office VBA compression algorithm
    ([MS-OVBA] section 2.4.1.3 Decompression Algorithm).
    """
    if not compressed_data:
        return b""

    # Verify signature byte: 0x01
    if compressed_data[0] != 0x01:
        # Not a compressed container, return as raw or best-effort
        return compressed_data

    decompressed = bytearray()
    pos = 1
    total_len = len(compressed_data)

    while pos < total_len:
        if pos + 2 > total_len:
            break
        # Chunk header: 16-bit unsigned integer
        chunk_header = struct.unpack_from("<H", compressed_data, pos)[0]
        pos += 2

        # 12-bit size (size - 3), 3-bit signature (0b011 = 3), 1-bit compressed flag (1 = compressed)
        chunk_size = (chunk_header & 0x0FFF) + 3
        chunk_signature = (chunk_header >> 12) & 0x07
        chunk_compressed = (chunk_header >> 15) & 0x01

        chunk_end = min(pos + chunk_size - 2, total_len)

        if chunk_compressed == 0:
            # Uncompressed chunk: copy 4096 bytes or remainder directly
            decompressed.extend(compressed_data[pos:chunk_end])
            pos = chunk_end
            continue

        chunk_decompressed_start = len(decompressed)

        while pos < chunk_end:
            flag_byte = compressed_data[pos]
            pos += 1

            for bit in range(8):
                if pos >= chunk_end:
                    break

                flag = (flag_byte >> bit) & 1
                if flag == 0:
                    # Literal token
                    decompressed.append(compressed_data[pos])
                    pos += 1
                else:
                    # Copy token
                    if pos + 2 > chunk_end:
                        pos = chunk_end
                        break

                    token = struct.unpack_from("<H", compressed_data, pos)[0]
                    pos += 2

                    # Dynamic bit allocation based on current offset within chunk
                    offset_in_chunk = len(decompressed) - chunk_decompressed_start
                    # Maximum offset is 4096
                    bit_count = max(4, (offset_in_chunk - 1).bit_length()) if offset_in_chunk > 0 else 4
                    length_mask = 0xFFFF >> bit_count
                    offset_mask = ~length_mask & 0xFFFF

                    length = (token & length_mask) + 3
                    offset = ((token & offset_mask) >> (16 - bit_count)) + 1

                    copy_start = len(decompressed) - offset
                    for _ in range(length):
                        if 0 <= copy_start < len(decompressed):
                            decompressed.append(decompressed[copy_start])
                            copy_start += 1

    return bytes(decompressed)


def parse_vba_project(vba_bin_data: bytes) -> Dict[str, any]:
    """
    Parse a vbaProject.bin buffer, extract all modules, code, references, and project info.
    Returns:
        Dict with keys:
          'project_info': dict of PROJECT properties
          'modules': dict mapping module_name -> {'code': str, 'type': str}
          'references': list of reference strings
    """
    cfb = CFBReader(vba_bin_data)
    streams = cfb.list_streams()

    result = {
        "project_info": {},
        "modules": {},
        "references": [],
    }

    # 1. Parse 'dir' stream (compressed VBA directory stream containing module list & offsets)
    dir_stream = streams.get("dir")
    modules_meta = {}
    if dir_stream:
        try:
            decomp_dir = decompress_msovba(dir_stream)
            modules_meta = _parse_dir_stream(decomp_dir)
            result["project_info"]["project_name"] = modules_meta.get("project_name", "")
            result["references"] = modules_meta.get("references", [])
        except Exception as e:
            result["project_info"]["dir_parse_error"] = str(e)

    # 2. Parse PROJECT stream (text ini-like format)
    project_stream = streams.get("PROJECT")
    if project_stream:
        try:
            proj_text = project_stream.decode("latin-1", errors="replace")
            for line in proj_text.splitlines():
                if "=" in line:
                    k, v = line.split("=", 1)
                    k = k.strip()
                    v = v.strip()
                    if k.startswith("Document") or k.startswith("Module") or k.startswith("Class") or k.startswith("BaseClass"):
                        mod_type = k.split("/")[0] if "/" in k else k
                        mod_name = v.split("/")[0] if "/" in v else v
                        if mod_name not in modules_meta.get("modules", {}):
                            modules_meta.setdefault("modules", {})[mod_name] = {"type": mod_type, "offset": 0}
                    else:
                        result["project_info"][k] = v
        except Exception:
            pass

    # 3. Extract code from module streams in VBA storage
    vba_modules = modules_meta.get("modules", {})
    for stream_name, stream_bytes in streams.items():
        if stream_name in ("dir", "PROJECT", "PROJECTwm", "_VBA_PROJECT"):
            continue

        meta = vba_modules.get(stream_name, {})
        text_offset = meta.get("offset", 0)

        # Module stream begins with compiled P-code up to text_offset, followed by compressed source text
        source_bytes = stream_bytes[text_offset:] if text_offset < len(stream_bytes) else stream_bytes
        decomp_code = decompress_msovba(source_bytes)

        if decomp_code:
            code_text = _clean_vba_code(decomp_code)
            mod_type = meta.get("type", "Module")
            result["modules"][stream_name] = {
                "name": stream_name,
                "type": mod_type,
                "code": code_text,
            }

    return result


def _parse_dir_stream(data: bytes) -> Dict[str, any]:
    """Parse uncompressed dir stream record structures [MS-OVBA] 2.3.4.2."""
    pos = 0
    total = len(data)
    info = {"project_name": "", "references": [], "modules": {}}
    curr_module = None

    while pos + 6 <= total:
        rec_id = struct.unpack_from("<H", data, pos)[0]
        rec_len = struct.unpack_from("<I", data, pos + 2)[0]
        pos += 6

        rec_data = data[pos : pos + rec_len]
        pos += rec_len

        # PROJECTNAME Record = 0x0004
        if rec_id == 0x0004:
            info["project_name"] = rec_data.decode("latin-1", errors="replace")

        # REFERENCENAME Record = 0x0016
        elif rec_id == 0x0016:
            ref_name = rec_data.decode("latin-1", errors="replace")
            info["references"].append(ref_name)

        # MODULEMODULENAME Record = 0x0019
        elif rec_id == 0x0019:
            mod_name = rec_data.decode("latin-1", errors="replace")
            curr_module = {"name": mod_name, "type": "Module", "offset": 0}
            info["modules"][mod_name] = curr_module

        # MODULEOFFSET Record = 0x0031
        elif rec_id == 0x0031 and curr_module is not None:
            if len(rec_data) >= 4:
                curr_module["offset"] = struct.unpack("<I", rec_data[:4])[0]

        # MODULETYPE (Standard / Doc) = 0x0021 / 0x0022
        elif rec_id == 0x0021 and curr_module is not None:
            curr_module["type"] = "Procedural"
        elif rec_id == 0x0022 and curr_module is not None:
            curr_module["type"] = "Document"

    return info


def _clean_vba_code(raw_decompressed: bytes) -> str:
    """Decode and clean decompressed VBA code, removing binary attributes if necessary."""
    # Attempt utf-8 first, fallback to cp1252/latin-1
    try:
        text = raw_decompressed.decode("utf-8")
    except UnicodeDecodeError:
        text = raw_decompressed.decode("latin-1", errors="replace")

    # Normalize CRLF
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text


def export_vba_to_directory(vba_data: Dict[str, any], target_dir: str) -> Dict[str, int]:
    """
    Write extracted VBA modules and metadata to target_dir/VBA_Extracted/.
    Returns counts of exported items: {'modules': count, 'classes': count, ...}
    """
    vba_out = os.path.join(target_dir, "VBA_Extracted")
    os.makedirs(vba_out, exist_ok=True)

    counts = {"modules": 0, "classes": 0, "forms": 0, "documents": 0}

    # Save Project_Info.json
    proj_info_path = os.path.join(vba_out, "Project_Info.json")
    with open(proj_info_path, "w", encoding="utf-8") as f:
        json.dump(vba_data.get("project_info", {}), f, indent=2, ensure_ascii=False)

    # Save References.txt
    refs = vba_data.get("references", [])
    if refs:
        refs_path = os.path.join(vba_out, "References.txt")
        with open(refs_path, "w", encoding="utf-8") as f:
            for r in refs:
                f.write(f"{r}\n")

    for mod_name, mod in vba_data.get("modules", {}).items():
        m_type = mod.get("type", "Module")
        code = mod.get("code", "")

        # Categorize subfolder & extension
        if m_type == "Document" or "ThisWorkbook" in mod_name or "Sheet" in mod_name:
            sub = "Document"
            ext = ".cls"
            counts["documents"] += 1
        elif "Class" in m_type or m_type == "Class":
            sub = "Classes"
            ext = ".cls"
            counts["classes"] += 1
        elif "Form" in m_type or m_type == "BaseClass":
            sub = "Forms"
            ext = ".frm"
            counts["forms"] += 1
        else:
            sub = "Modules"
            ext = ".bas"
            counts["modules"] += 1

        sub_path = os.path.join(vba_out, sub)
        os.makedirs(sub_path, exist_ok=True)
        file_path = os.path.join(sub_path, f"{mod_name}{ext}")

        with open(file_path, "w", encoding="utf-8") as f:
            f.write(code)

    return counts
