from __future__ import annotations

import io
import unittest
import zipfile
from xml.etree import ElementTree as ET

from residential_calculator import calculate_residential_comparison
from tests.test_residential_calculator import reference_project


WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def document_text(package: zipfile.ZipFile) -> str:
    root = ET.fromstring(package.read("word/document.xml"))
    return "".join(node.text or "" for node in root.findall(f".//{{{WORD_NS}}}t"))


class ResidentialWordTests(unittest.TestCase):
    def test_docx_contains_same_result_as_calculator(self):
        from office_export import build_residential_docx

        project = reference_project()
        result = calculate_residential_comparison(project)
        content = build_residential_docx(project, result)

        self.assertTrue(content.startswith(b"PK"))
        with zipfile.ZipFile(io.BytesIO(content)) as package:
            self.assertIsNone(package.testzip())
            text = document_text(package)
        self.assertIn("2,779", text)
        self.assertIn("337,454", text)
        self.assertIn("18/19（地下一层、地上18层、顶层）", text)
        self.assertIn("人民币叁拾叁万柒仟肆佰伍拾肆元整", text)

    def test_docx_includes_non_neutral_p1_p2_explanations(self):
        from office_export import build_residential_docx

        project = reference_project()
        project["comparables"][0]["transaction_condition_factor"] = 0.98
        project["comparables"][0]["transaction_condition_note"] = "成交条件略低于正常市场"
        project["comparables"][0]["market_condition_factor"] = 1.03
        project["comparables"][0]["market_condition_note"] = "按近期价格指数修正"
        result = calculate_residential_comparison(project)
        content = build_residential_docx(project, result)

        with zipfile.ZipFile(io.BytesIO(content)) as package:
            text = document_text(package)
        self.assertIn("0.980", text)
        self.assertIn("成交条件略低于正常市场", text)
        self.assertIn("1.030", text)
        self.assertIn("按近期价格指数修正", text)


if __name__ == "__main__":
    unittest.main()
