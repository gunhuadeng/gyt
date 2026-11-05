from __future__ import annotations

import io
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from residential_calculator import calculate_residential_comparison
from tests.test_residential_calculator import reference_project


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
TEMPLATE = Path(__file__).resolve().parents[1] / "templates" / "formal-report-template.docx"


def text_from_docx(content: bytes) -> tuple[str, set[str]]:
    with zipfile.ZipFile(io.BytesIO(content)) as package:
        assert package.testzip() is None
        root = ET.fromstring(package.read("word/document.xml"))
        text = "".join(node.text or "" for node in root.findall(f".//{{{W}}}t"))
        names = set(package.namelist())
    return text, names


def signature_media_names(content: bytes) -> set[str]:
    with zipfile.ZipFile(io.BytesIO(content)) as package:
        root = ET.fromstring(package.read("word/document.xml"))
        signature_table = next(
            (
                table
                for table in root.findall(f".//{{{W}}}tbl")
                if "3220070051" in "".join(
                    node.text or "" for node in table.findall(f".//{{{W}}}t")
                )
            ),
            None,
        )
        if signature_table is None:
            return set()
        relationship_ids = {
            blip.get(f"{{{R}}}embed")
            for blip in signature_table.findall(f".//{{{A}}}blip")
            if blip.get(f"{{{R}}}embed")
        }
        relationships = ET.fromstring(package.read("word/_rels/document.xml.rels"))
        targets = {
            relationship.get("Target")
            for relationship in relationships.findall(f"{{{P}}}Relationship")
            if relationship.get("Id") in relationship_ids
        }
        return {
            ("word/" + target.replace("\\", "/")).lstrip("/")
            for target in targets
            if target
        }


@unittest.skipUnless(TEMPLATE.exists(), "Formal report template is unavailable")
class FormalReportTemplateTests(unittest.TestCase):
    def test_fills_rights_location_result_and_floor_note(self):
        from report_template import build_template_report

        project = reference_project()
        result = calculate_residential_comparison(project)
        content = build_template_report(TEMPLATE.read_bytes(), project, result)
        text, names = text_from_docx(content)

        self.assertIn("测试产权人", text)
        self.assertIn("苏（2026）示例不动产权第TEST001号", text)
        self.assertIn("东临测试路，西临示例路，南临规划路，北临公园路。", text)
        self.assertIn("估价对象位于测试居住片区，周边交通和生活配套完善。", text)
        self.assertIn("测试区域居住环境较好，区位条件对估价对象价值产生积极影响。", text)
        self.assertIn("18/19（地下一层、地上18层、顶层）", text)
        self.assertIn("本次住宅层次修正按地上18/18层计算，层次修正系数为99.4%", text)
        self.assertIn("337,454", text)
        self.assertIn("2,779", text)
        self.assertEqual(len([name for name in names if name.startswith("word/media/")]), 17)

    def test_preserves_styles_and_media_package(self):
        from report_template import build_template_report

        project = reference_project()
        result = calculate_residential_comparison(project)
        template = TEMPLATE.read_bytes()
        output = build_template_report(template, project, result)

        with zipfile.ZipFile(io.BytesIO(template)) as source, zipfile.ZipFile(io.BytesIO(output)) as changed:
            self.assertEqual(source.read("word/styles.xml"), changed.read("word/styles.xml"))
            source_media = {name for name in source.namelist() if name.startswith("word/media/")}
            changed_media = {name for name in changed.namelist() if name.startswith("word/media/")}
            self.assertEqual(source_media - signature_media_names(template), changed_media)

    def test_writes_non_neutral_p1_and_p2_into_calculation_table(self):
        from report_template import build_template_report

        project = reference_project()
        project["comparables"][0]["transaction_condition_factor"] = 0.98
        project["comparables"][0]["transaction_condition_note"] = "测试交易情况说明"
        project["comparables"][0]["market_condition_factor"] = 1.03
        project["comparables"][0]["market_condition_note"] = "测试市场状况说明"
        result = calculate_residential_comparison(project)
        text, _ = text_from_docx(build_template_report(TEMPLATE.read_bytes(), project, result))

        self.assertIn("0.980", text)
        self.assertIn("1.030", text)
        self.assertIn("测试交易情况说明", text)
        self.assertIn("测试市场状况说明", text)

    def test_updates_valuation_date_in_template(self):
        from report_template import build_template_report

        project = reference_project()
        project["subject"]["valuation_date"] = "2026-09-30"
        result = calculate_residential_comparison(project)
        text, _ = text_from_docx(build_template_report(TEMPLATE.read_bytes(), project, result))

        self.assertIn("2026年09月30日", text)
        self.assertNotIn("2026年08月28日", text)

    def test_removes_preexisting_signature_images(self):
        from report_template import build_template_report

        project = reference_project()
        result = calculate_residential_comparison(project)
        template = TEMPLATE.read_bytes()
        signature_media = signature_media_names(template)
        output = build_template_report(template, project, result)
        _, names = text_from_docx(output)

        self.assertTrue(signature_media)
        self.assertFalse(signature_media & names)
        self.assertFalse(signature_media_names(output))

    def test_keeps_package_relationships_in_default_namespace(self):
        from report_template import build_template_report

        project = reference_project()
        result = calculate_residential_comparison(project)
        output = build_template_report(TEMPLATE.read_bytes(), project, result)

        with zipfile.ZipFile(io.BytesIO(output)) as package:
            relationships = package.read("word/_rels/document.xml.rels")
        self.assertIn(b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">', relationships)
        self.assertNotIn(b"ns0:Relationships", relationships)


if __name__ == "__main__":
    unittest.main()
