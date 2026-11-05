from __future__ import annotations

import io
import os
import posixpath
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from residential_calculator import calculate_residential_comparison
from tests.test_residential_calculator import reference_project


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
reference_value = os.environ.get("REFERENCE_XLSX")
REFERENCE_XLSX = Path(reference_value) if reference_value else Path("__missing_reference__.xlsx")


def rename_sheet(workbook_bytes: bytes, old_name: str, new_name: str) -> bytes:
    source = zipfile.ZipFile(io.BytesIO(workbook_bytes))
    target_buffer = io.BytesIO()
    with zipfile.ZipFile(target_buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for name in source.namelist():
            content = source.read(name)
            if name == "xl/workbook.xml":
                root = ET.fromstring(content)
                for sheet in root.findall(f".//{{{SHEET_NS}}}sheet"):
                    if sheet.get("name") == old_name:
                        sheet.set("name", new_name)
                content = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            target.writestr(name, content)
    source.close()
    return target_buffer.getvalue()


def sheet_xml_and_cells(workbook_bytes: bytes, sheet_name: str) -> tuple[ET.Element, dict[str, str]]:
    with zipfile.ZipFile(io.BytesIO(workbook_bytes)) as package:
        workbook = ET.fromstring(package.read("xl/workbook.xml"))
        rels = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            item.get("Id"): item.get("Target")
            for item in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
        }
        target = None
        for sheet in workbook.findall(f".//{{{SHEET_NS}}}sheet"):
            if sheet.get("name") == sheet_name:
                target = rel_map[sheet.get(f"{{{REL_NS}}}id")]
                break
        if target is None:
            raise AssertionError(f"Sheet not found: {sheet_name}")
        path = posixpath.normpath(posixpath.join("xl", target))
        root = ET.fromstring(package.read(path))
        values: dict[str, str] = {}
        for cell in root.findall(f".//{{{SHEET_NS}}}c"):
            reference = cell.get("r")
            inline = cell.find(f"{{{SHEET_NS}}}is/{{{SHEET_NS}}}t")
            number = cell.find(f"{{{SHEET_NS}}}v")
            formula = cell.find(f"{{{SHEET_NS}}}f")
            values[reference] = (
                inline.text
                if inline is not None
                else (number.text if number is not None else (formula.text if formula is not None else ""))
            )
        return root, values


@unittest.skipUnless(REFERENCE_XLSX.exists(), "Reference residential workbook is unavailable")
class ResidentialTemplateTests(unittest.TestCase):
    def test_preserves_hidden_sheets_and_images(self):
        from xlsx_template import build_filled_workbook

        project = reference_project()
        result = calculate_residential_comparison(project)
        original = REFERENCE_XLSX.read_bytes()
        output = build_filled_workbook(original, project, result)
        before = set(zipfile.ZipFile(io.BytesIO(original)).namelist())
        after = set(zipfile.ZipFile(io.BytesIO(output)).namelist())

        self.assertTrue(after.issuperset(before - {"xl/workbook.xml", "xl/worksheets/sheet5.xml", "xl/calcChain.xml"}))
        self.assertGreaterEqual(len([name for name in after if name.startswith("xl/media/")]), 10)

    def test_writes_non_neutral_p1_and_p2(self):
        from xlsx_template import build_filled_workbook

        project = reference_project()
        project["comparables"][0]["transaction_condition_factor"] = 0.98
        project["comparables"][0]["market_condition_factor"] = 1.03
        result = calculate_residential_comparison(project)
        output = build_filled_workbook(REFERENCE_XLSX.read_bytes(), project, result)
        _, values = sheet_xml_and_cells(output, "1-1 测算表")

        self.assertEqual(values["C54"], "0.98")
        self.assertEqual(values["C55"], "1.03")
        self.assertEqual(values["D9"], "18/18（顶层）")
        self.assertEqual(values["D6"], "示例市滨江区示范路88号示例花园5幢1单元1801室")
        self.assertEqual(values["D7"], "121.43")
        self.assertEqual(values["E32"], "3341")
        self.assertEqual(values["D16"], "二等")
        self.assertEqual(values["D30"], "三等")
        self.assertEqual(values["B43"], "100")
        self.assertEqual(values["C43"], "102")
        self.assertEqual(values["B45"], "99.4")
        self.assertEqual(values["C45"], "102")
        self.assertEqual(values["C50"], "0.965")
        self.assertEqual(values["C53"], "1")

    def test_rejects_template_without_calculation_sheet(self):
        from xlsx_template import TemplateFormatError, build_filled_workbook

        broken = rename_sheet(REFERENCE_XLSX.read_bytes(), "1-1 测算表", "错误的测算表")
        project = reference_project()
        result = calculate_residential_comparison(project)

        with self.assertRaisesRegex(TemplateFormatError, "1-1 测算表"):
            build_filled_workbook(broken, project, result)


if __name__ == "__main__":
    unittest.main()
