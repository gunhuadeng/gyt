from __future__ import annotations

import io
import posixpath
import re
import zipfile
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Any
from xml.etree import ElementTree as ET

from residential_rules import REGIONAL_FACTORS


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
TARGET_SHEET = "1-1 测算表"
GRADE_LABELS = ("一等", "二等", "三等", "四等", "五等")


class TemplateFormatError(ValueError):
    pass


def _excel_date(value: Any) -> int | None:
    if value in (None, ""):
        return None
    try:
        parsed = date.fromisoformat(str(value))
    except ValueError as exc:
        raise TemplateFormatError(f"日期格式无效：{value}") from exc
    return (parsed - date(1899, 12, 30)).days


def _number(value: Any) -> float | int:
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError) as exc:
        raise TemplateFormatError(f"数值无效：{value}") from exc
    if number == number.to_integral_value():
        return int(number)
    return float(number)


def _floor_display(item: dict[str, Any]) -> str:
    floor = int(item.get("floor") or 0)
    total = int(item.get("above_ground_total_floors") or 0)
    return f"{floor}/{total}（顶层）" if floor == total else f"{floor}/{total}"


def _find_target_sheet(package: zipfile.ZipFile) -> str:
    try:
        workbook = ET.fromstring(package.read("xl/workbook.xml"))
        rels = ET.fromstring(package.read("xl/_rels/workbook.xml.rels"))
    except (KeyError, ET.ParseError) as exc:
        raise TemplateFormatError("Excel 原表缺少工作簿结构") from exc
    relationship_map = {
        item.get("Id"): item.get("Target")
        for item in rels.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    for sheet in workbook.findall(f".//{{{SHEET_NS}}}sheet"):
        if sheet.get("name") == TARGET_SHEET:
            rel_id = sheet.get(f"{{{REL_NS}}}id")
            target = relationship_map.get(rel_id)
            if not target:
                break
            if target.startswith("/"):
                return target.lstrip("/")
            return posixpath.normpath(posixpath.join("xl", target))
    raise TemplateFormatError(f"Excel 原表缺少“{TARGET_SHEET}”工作表")


def _set_cell(root: ET.Element, reference: str, value: Any, is_number: bool = False) -> None:
    row_number = re.match(r"\d+", reference[1:])
    if not row_number:
        raise TemplateFormatError(f"单元格引用无效：{reference}")
    row_ref = reference[1 : 1 + len(row_number.group())]
    sheet_data = root.find(f"{{{SHEET_NS}}}sheetData")
    if sheet_data is None:
        raise TemplateFormatError(f"工作表缺少单元格区域：{reference}")
    row = next((item for item in sheet_data.findall(f"{{{SHEET_NS}}}row") if item.get("r") == row_ref), None)
    if row is None:
        row = ET.SubElement(sheet_data, f"{{{SHEET_NS}}}row", {"r": row_ref})
    cell = next((item for item in row.findall(f"{{{SHEET_NS}}}c") if item.get("r") == reference), None)
    if cell is None:
        cell = ET.SubElement(row, f"{{{SHEET_NS}}}c", {"r": reference})
    for child in list(cell):
        cell.remove(child)
    if is_number:
        cell.attrib.pop("t", None)
        ET.SubElement(cell, f"{{{SHEET_NS}}}v").text = str(value)
    else:
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, f"{{{SHEET_NS}}}is")
        text = ET.SubElement(inline, f"{{{SHEET_NS}}}t")
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = "" if value is None else str(value)


def _input_cells(project: dict[str, Any], result: dict[str, Any]) -> dict[str, tuple[Any, bool]]:
    subject = project.get("subject") or {}
    comparables = list(project.get("comparables") or [])
    if len(comparables) != 3:
        raise TemplateFormatError("项目必须包含三个可比实例")
    entities = [subject, *comparables]
    cells: dict[str, tuple[Any, bool]] = {}

    column_values = (
        ("D", subject),
        ("E", comparables[0]),
        ("F", comparables[1]),
        ("G", comparables[2]),
    )
    for column, item in column_values:
        cells[f"{column}6"] = (item.get("address") or "", False)
        cells[f"{column}7"] = (_number(item.get("area")), True)
        cells[f"{column}8"] = (item.get("structure") or "", False)
        cells[f"{column}9"] = (_floor_display(item), False)
        cells[f"{column}10"] = (_number(item.get("build_year")), True)
        cells[f"{column}11"] = (item.get("orientation") or "", False)
        cells[f"{column}12"] = (item.get("east_west") or "", False)
        cells[f"{column}14"] = (item.get("appurtenance_description") or "/", False)
        cells[f"{column}15"] = (item.get("property_nature") or "", False)
        transaction_date = subject.get("valuation_date") if column == "D" else item.get("transaction_date")
        excel_date = _excel_date(transaction_date)
        cells[f"{column}31"] = (excel_date, True)
        if column == "D":
            cells[f"{column}32"] = ("/", False)
        else:
            cells[f"{column}32"] = (_number(item.get("transaction_unit_price")), True)

    regional_grades = project.get("regional_grades") or {}
    grade_sets = [
        regional_grades.get("subject") or {},
        *(regional_grades.get("comparables") or []),
    ]
    for grade_index, factor in enumerate(REGIONAL_FACTORS):
        row = 16 + grade_index
        for column_index, column in enumerate(("D", "E", "F", "G")):
            grade = int(grade_sets[column_index][factor["key"]])
            cells[f"{column}{row}"] = (GRADE_LABELS[grade - 1], False)

    subject_result = result["subject"]
    case_results = result["comparables"]
    factor_columns = ("B", "C", "D", "E")
    result_sets = [subject_result, *case_results]
    for index, column in enumerate(factor_columns):
        item = entities[index]
        computed = result_sets[index]
        cells[f"{column}43"] = (_number(item.get("area_factor_pct")), True)
        cells[f"{column}45"] = (_number(computed["factor_breakdown"]["level"]), True)
        cells[f"{column}48"] = (_number(computed["factor_breakdown"]["east_west"]), True)
        cells[f"{column}51"] = (_number(item.get("decoration_unit_price", 0)), True)
        cells[f"{column}52"] = (_number(item.get("decoration_newness", 0)), True)
        # P4 is stored on each comparable result; the subject baseline is 1.000.
        cells[f"{column}53"] = (
            1 if index == 0 else _number(case_results[index - 1]["p4"]),
            True,
        )
        cells[f"{column}54"] = (
            1 if index == 0 else _number(item.get("transaction_condition_factor", 1)),
            True,
        )
        cells[f"{column}55"] = (
            1 if index == 0 else _number(item.get("market_condition_factor", 1)),
            True,
        )
        cells[f"{column}58"] = (_number(item.get("appurtenance_unit_price", 0)), True)
    return cells


def _update_workbook_calculation(workbook_xml: bytes) -> bytes:
    text = workbook_xml.decode("utf-8")
    if re.search(r"<(?:\w+:)?calcPr\b", text):
        text = re.sub(
            r"(<(?:\w+:)?calcPr\b[^>]*?)(/?>)",
            lambda match: (
                re.sub(r'\s+fullCalcOnLoad="[^"]*"', "", match.group(1))
                + ' fullCalcOnLoad="1"'
                + match.group(2)
            ),
            text,
            count=1,
        )
    else:
        text = re.sub(
            r"(</(?:\w+:)?workbook>)",
            r'<calcPr calcId="0" fullCalcOnLoad="1"/>\1',
            text,
            count=1,
        )
    return text.encode("utf-8")


def build_filled_workbook(
    template_bytes: bytes,
    project: dict[str, Any],
    result: dict[str, Any],
) -> bytes:
    try:
        source = zipfile.ZipFile(io.BytesIO(template_bytes))
    except zipfile.BadZipFile as exc:
        raise TemplateFormatError("Excel 原表不是有效的 XLSX 文件") from exc

    sheet_path = _find_target_sheet(source)
    try:
        sheet_xml = source.read(sheet_path)
    except KeyError as exc:
        raise TemplateFormatError(f"Excel 原表缺少工作表文件：{sheet_path}") from exc
    ET.register_namespace("", SHEET_NS)
    root = ET.fromstring(sheet_xml)
    for reference, (value, is_number) in _input_cells(project, result).items():
        _set_cell(root, reference, value, is_number)
    new_sheet = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    target = io.BytesIO()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as package:
        for info in source.infolist():
            if info.filename in {"xl/calcChain.xml", "xl/calcChain.xml.rels"}:
                continue
            content = source.read(info.filename)
            if info.filename == sheet_path:
                content = new_sheet
            elif info.filename == "xl/workbook.xml":
                content = _update_workbook_calculation(content)
            package.writestr(info, content)
    source.close()
    return target.getvalue()
