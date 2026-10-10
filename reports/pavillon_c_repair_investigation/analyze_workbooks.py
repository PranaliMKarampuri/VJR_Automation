from __future__ import annotations

import hashlib
import json
import posixpath
import re
import sys
import warnings
import zipfile
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

from openpyxl import load_workbook


GENERATED = Path(sys.argv[1])
DOWNLOADED = Path(sys.argv[2])
TEMPLATE = Path(sys.argv[3])
OUT = Path(__file__).parent


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def inspect(path: Path) -> dict:
    info = {"path": str(path), "size": path.stat().st_size, "sha256": sha(path)}
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        names = set(z.namelist())
        info["zip_member_count"] = len(names)
        info["zip_testzip_bad_member"] = bad
        info["members"] = sorted(names)
        xml_problems, invalid_mc, rel_problems = [], [], []
        for name in sorted(names):
            if name.endswith((".xml", ".rels")):
                data = z.read(name)
                try:
                    ET.fromstring(data)
                except Exception as e:
                    xml_problems.append({"part": name, "error": str(e)})
                if name.endswith(".xml"):
                    root_start = data.find(b"<worksheet")
                    root_end = data.find(b">", root_start) + 1 if root_start >= 0 else 0
                    root = data[root_start:root_end].decode("utf-8", "replace") if root_end else ""
                    decl = set(re.findall(r"xmlns:([\w.-]+)=", root))
                    ign = re.search(r"(?:\w+:)?Ignorable=[\"']([^\"']*)", root)
                    if ign:
                        missing = [p for p in ign.group(1).split() if p not in decl]
                        if missing:
                            invalid_mc.append({"part": name, "ignorable": ign.group(1), "undeclared": missing, "root": root})
        # Resolve OPC relationship targets against their source part.
        for relname in sorted(n for n in names if n.endswith(".rels")):
            root = ET.fromstring(z.read(relname))
            if relname == "_rels/.rels":
                base = PurePosixPath("")
            else:
                p = PurePosixPath(relname)
                # xl/_rels/workbook.xml.rels -> xl/workbook.xml
                base = p.parent.parent
            for rel in root:
                if rel.attrib.get("TargetMode") == "External":
                    continue
                target = rel.attrib.get("Target", "")
                norm = posixpath.normpath(posixpath.join(base.as_posix(), target))
                if norm not in names:
                    rel_problems.append({"rels": relname, "target": target, "resolved": norm})
        info["xml_parse_problems"] = xml_problems
        info["mc_ignorable_undeclared_prefixes"] = invalid_mc
        info["relationship_target_problems"] = rel_problems
        # Structural package component inventory.
        info["part_categories"] = {
            "drawings": sorted(n for n in names if n.startswith("xl/drawings/")),
            "media": sorted(n for n in names if n.startswith("xl/media/")),
            "tables": sorted(n for n in names if n.startswith("xl/tables/")),
            "comments_vml": sorted(n for n in names if "comment" in n.lower() or "vml" in n.lower()),
        }
        info["extension_markers"] = {}
        info["worksheet_features"] = {}
        for n in sorted(names):
            if n.endswith(".xml") and (n.startswith("xl/worksheets/") or n == "xl/workbook.xml"):
                raw = z.read(n)
                data = raw.decode("utf-8", "replace")
                marks = [x for x in ("<extLst", "x14:dataValidations", "x14:dataValidation", "<dataValidations", "<mergeCells", "<drawing", "<tableParts") if x in data]
                if marks:
                    info["extension_markers"][n] = marks
                if n.startswith("xl/worksheets/"):
                    root = ET.fromstring(raw)
                    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"}
                    info["worksheet_features"][n] = {
                        "merge_cells": len(root.findall("m:mergeCells/m:mergeCell", ns)),
                        "standard_data_validations": len(root.findall("m:dataValidations/m:dataValidation", ns)),
                        "x14_data_validations": len(root.findall(".//x14:dataValidations/x14:dataValidation", ns)),
                        "extension_list_count": len(root.findall("m:extLst", ns)),
                        "drawing_refs": len(root.findall("m:drawing", ns)),
                        "table_parts_refs": len(root.findall("m:tableParts/m:tablePart", ns)),
                    }
        info["package_content_types"] = z.read("[Content_Types].xml").decode("utf-8", "replace")
    # Capture openpyxl read-only warnings/errors.
    info["openpyxl"] = {}
    for read_only in (True, False):
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            try:
                wb = load_workbook(path, read_only=read_only, data_only=False)
                info["openpyxl"]["read_only" if read_only else "normal_mode"] = {"loaded": True, "sheetnames": wb.sheetnames, "defined_names": list(wb.defined_names)}
                wb.close()
            except Exception as e:
                info["openpyxl"]["read_only" if read_only else "normal_mode"] = {"loaded": False, "error": repr(e)}
            info["openpyxl"]["read_only" if read_only else "normal_mode"]["warnings"] = [str(w.message) for w in caught]
    return info


items = {"generated": inspect(GENERATED), "downloaded": inspect(DOWNLOADED), "template": inspect(TEMPLATE)}
with zipfile.ZipFile(GENERATED) as gz, zipfile.ZipFile(TEMPLATE) as tz:
    gm, tm = set(gz.namelist()), set(tz.namelist())
    changed, generated_only, template_only = [], sorted(gm - tm), sorted(tm - gm)
    for n in sorted(gm & tm):
        if gz.read(n) != tz.read(n):
            changed.append(n)
    items["template_comparison"] = {
        "generated_only_parts": generated_only,
        "template_only_parts": template_only,
        "byte_different_parts": changed,
    }
with zipfile.ZipFile(GENERATED) as gz, zipfile.ZipFile(DOWNLOADED) as dz:
    gm, dm = set(gz.namelist()), set(dz.namelist())
    items["generated_downloaded_comparison"] = {
        "same_sha256": items["generated"]["sha256"] == items["downloaded"]["sha256"],
        "same_size": items["generated"]["size"] == items["downloaded"]["size"],
        "same_member_names": gm == dm,
        "byte_different_parts": [n for n in sorted(gm & dm) if gz.read(n) != dz.read(n)],
        "generated_only_parts": sorted(gm - dm),
        "downloaded_only_parts": sorted(dm - gm),
    }
report = OUT / "workbook_package_analysis.json"
report.write_text(json.dumps(items, indent=2), encoding="utf-8")
print(json.dumps({
    "generated": {k: items["generated"][k] for k in ("path", "size", "sha256", "zip_member_count", "zip_testzip_bad_member", "xml_parse_problems", "mc_ignorable_undeclared_prefixes", "relationship_target_problems", "openpyxl", "extension_markers", "part_categories")},
    "downloaded": {k: items["downloaded"][k] for k in ("path", "size", "sha256", "zip_member_count", "zip_testzip_bad_member", "xml_parse_problems", "mc_ignorable_undeclared_prefixes", "relationship_target_problems", "openpyxl")},
    "template": {k: items["template"][k] for k in ("path", "size", "sha256", "zip_member_count", "zip_testzip_bad_member", "xml_parse_problems", "mc_ignorable_undeclared_prefixes", "relationship_target_problems", "openpyxl", "extension_markers", "part_categories")},
    "template_comparison": items["template_comparison"],
    "generated_downloaded_comparison": items["generated_downloaded_comparison"],
    "report": str(report),
}, indent=2))
