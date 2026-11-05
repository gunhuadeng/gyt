from __future__ import annotations

import json
import os
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from office_export import build_residential_docx
from report_template import build_template_report
from residential_calculator import calculate_residential_comparison
from xlsx_template import build_filled_workbook


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def validate_zip(path: Path, required: list[str]) -> zipfile.ZipFile:
    package = zipfile.ZipFile(path)
    assert package.testzip() is None
    for name in package.namelist():
        if name.endswith((".xml", ".rels")):
            ET.fromstring(package.read(name))
    for name in required:
        assert name in package.namelist(), f"{path.name} missing {name}"
    return package


def main() -> None:
    reference_value = os.environ.get("REFERENCE_XLSX")
    if not reference_value:
        raise SystemExit("Set REFERENCE_XLSX to the source workbook path")
    reference = Path(reference_value)
    if not reference.exists():
        raise SystemExit(f"Reference workbook not found: {reference}")
    project = json.loads((ROOT / "sample-project.json").read_text(encoding="utf-8"))
    result = calculate_residential_comparison(project)
    assert result["estimated_unit_price"] == "2779"
    assert result["estimated_total"] == "337454"

    output_dir = ROOT / "samples"
    output_dir.mkdir(exist_ok=True)
    workbook_path = output_dir / "residential-workbook-filled.xlsx"
    report_path = output_dir / "residential-report-draft.docx"
    workbook_path.write_bytes(
        build_filled_workbook(reference.read_bytes(), project, result)
    )
    report_path.write_bytes(build_residential_docx(project, result))
    template_path = ROOT / "templates" / "formal-report-template.docx"
    formal_report_path = output_dir / "formal-report-from-saved-template.docx"
    if template_path.exists():
        formal_report_path.write_bytes(
            build_template_report(template_path.read_bytes(), project, result)
        )

    with validate_zip(
        workbook_path,
        ["xl/workbook.xml", "xl/worksheets/sheet5.xml", "xl/sharedStrings.xml"],
    ) as workbook:
        media = [name for name in workbook.namelist() if name.startswith("xl/media/")]
        workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
        sheets = workbook_root.findall(f".//{{{SHEET_NS}}}sheet")
        hidden = [item.get("name") for item in sheets if item.get("state") == "hidden"]
        assert len(media) >= 10
        assert "税务局调取案例信息" in hidden
        assert b'fullCalcOnLoad="1"' in workbook.read("xl/workbook.xml")
        print(f"workbook: {workbook_path.name}, sheets={len(sheets)}, hidden={len(hidden)}, images={len(media)}")

    with validate_zip(report_path, ["word/document.xml", "word/styles.xml"]) as report:
        document = ET.fromstring(report.read("word/document.xml"))
        text = "".join(node.text or "" for node in document.findall(f".//{{{WORD_NS}}}t"))
        assert "2,779" in text
        assert "337,454" in text
        assert "18/19（地下一层、地上18层、顶层）" in text
        print(f"report: {report_path.name}, characters={len(text)}")

    if template_path.exists():
        with validate_zip(
            formal_report_path,
            ["word/document.xml", "word/styles.xml"],
        ) as report:
            media = [name for name in report.namelist() if name.startswith("word/media/")]
            document = ET.fromstring(report.read("word/document.xml"))
            text = "".join(node.text or "" for node in document.findall(f".//{{{WORD_NS}}}t"))
            assert len(media) == 17
            assert "测试产权人" in text
            assert "18/19（地下一层、地上18层、顶层）" in text
            assert "337,454" in text
            print(f"formal template report: {formal_report_path.name}, images={len(media)}")

    print("office package verification: OK")


if __name__ == "__main__":
    main()
