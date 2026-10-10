from __future__ import annotations

import hashlib
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from manual_to_vjr import audit_shared_formula_package
from stage1_app_backend import generate_vjr

HERE = Path(__file__).resolve().parent
OUTPUT_ROOT = ROOT / "output" / "shared_formula_fix_validation" / "fresh_jobs"
PAVILLON_ARCHIVE = Path.home() / "Downloads" / "TAKEOFF_Campus MIL Pavillon C.zip"


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def pavillon_input() -> Path:
    input_dir = HERE / "source_inputs"
    input_dir.mkdir(parents=True, exist_ok=True)
    destination = input_dir / "TAKEOFF_Plans_Pav_C_source.xlsx"
    with zipfile.ZipFile(PAVILLON_ARCHIVE) as archive:
        candidates = [name for name in archive.namelist()
                      if name.lower().endswith(".xlsx") and "_vjr" not in Path(name).stem.casefold()]
        if len(candidates) != 1:
            raise RuntimeError(f"Expected one Pavillon C source workbook in archive; found {candidates!r}")
        destination.write_bytes(archive.read(candidates[0]))
    return destination


projects = [
    ("Pavillon C", pavillon_input()),
    ("KATIVIK", ROOT / "source_files" / "TAKEOFF_CS KATIVIK - CENTRE D'ÉDUCATION AUX ADULTES.xlsx"),
    ("53310TT-B", ROOT / "test_input" / "TAKEOFF_Plans_ingenierie_53310TT-B.xlsx"),
    ("RTL LOT 1 canonical CSV", ROOT / "output" / "rtl_lot1_pressure_fix" / "rtl_lot1_normalized.csv"),
]

results = []
for label, source in projects:
    result = generate_vjr(source, output_dir=OUTPUT_ROOT)
    audit = audit_shared_formula_package(Path(result["output"]))
    if audit:
        raise RuntimeError(f"{label} output has shared-formula issues: {audit}")
    results.append({"project": label, "source": str(source), "source_sha256": sha(source),
                    "output": str(result["output"]), "output_sha256": result["output_sha256"],
                    "records_processed": result["records_processed"], "cells_written": result["cells_written"],
                    "shared_formula_audit": "PASS", "report": str(result["report"])})

summary = {"projects": results, "pavillon_archive": str(PAVILLON_ARCHIVE),
           "pavillon_archive_sha256": sha(PAVILLON_ARCHIVE),
           "note": "RTL LOT 1 source workbook was not available in the workspace; its previously normalized canonical CSV was used."}
out = HERE / "project_validation_results.json"
out.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
print(json.dumps(summary, indent=2, ensure_ascii=False))
