from __future__ import annotations

import copy
import io
import json
import zipfile
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
ET.register_namespace("w", W)

BASELINE_PATH = Path(__file__).resolve().parent / "templates" / "template_baseline.json"


class ReportTemplateError(ValueError):
    pass


def _load_baseline() -> dict[str, str]:
    try:
        payload = json.loads(BASELINE_PATH.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError) as exc:
        raise ReportTemplateError(
            "正式报告模板缺少私有基准配置 templates/template_baseline.json"
        ) from exc
    if not isinstance(payload, dict):
        raise ReportTemplateError("正式报告模板基准配置格式不正确")
    return {str(key): str(value) for key, value in payload.items()}


def _chinese_small_number(number: int) -> str:
    digits = "零一二三四五六七八九十"
    if 0 <= number <= 10:
        return digits[number]
    return str(number)


def _parse_iso_date(value: Any) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(str(value))
    except ValueError:
        return None


def _paragraph_text(paragraph: ET.Element) -> str:
    return "".join(node.text or "" for node in paragraph.findall(f".//{{{W}}}t"))


def _replace_once(paragraph: ET.Element, old: str, new: str) -> bool:
    text_nodes = list(paragraph.findall(f".//{{{W}}}t"))
    full_text = "".join(node.text or "" for node in text_nodes)
    start = full_text.find(old)
    if start < 0:
        return False
    end = start + len(old)
    position = 0
    first_node = None
    last_node = None
    first_offset = 0
    last_offset = 0
    for node in text_nodes:
        value = node.text or ""
        node_start = position
        node_end = position + len(value)
        if first_node is None and start < node_end:
            first_node = node
            first_offset = start - node_start
        if start < node_end and end <= node_end:
            last_node = node
            last_offset = end - node_start
            break
        position = node_end
    if first_node is None or last_node is None:
        return False
    first_value = first_node.text or ""
    last_value = last_node.text or ""
    prefix = first_value[:first_offset]
    suffix = last_value[last_offset:]
    if first_node is last_node:
        first_node.text = prefix + new + suffix
        return True
    first_node.text = prefix + new
    current = False
    for node in text_nodes:
        if node is first_node:
            current = True
            continue
        if not current:
            continue
        if node is last_node:
            node.text = suffix
            break
        node.text = ""
    return True


def _replace_all_in_paragraph(paragraph: ET.Element, replacements: list[tuple[str, str]]) -> bool:
    changed = False
    for old, new in sorted(replacements, key=lambda item: len(item[0]), reverse=True):
        if not old or old == new:
            continue
        while _replace_once(paragraph, old, new):
            changed = True
            if old in new:
                break
    return changed


def _replace_in_part(xml: bytes, replacements: list[tuple[str, str]]) -> bytes:
    root = ET.fromstring(xml)
    for paragraph in root.findall(f".//{{{W}}}p"):
        _replace_all_in_paragraph(paragraph, replacements)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _cell_text(cell: ET.Element) -> str:
    return "".join(_paragraph_text(p) for p in cell.findall(f".//{{{W}}}p")).strip()


def _set_cell_text(cell: ET.Element, value: str) -> None:
    paragraphs = cell.findall(f".//{{{W}}}p")
    if not paragraphs:
        paragraph = ET.SubElement(cell, f"{{{W}}}p")
    else:
        paragraph = paragraphs[0]
        for extra in paragraphs[1:]:
            cell.remove(extra)
    runs = paragraph.findall(f".//{{{W}}}r")
    run_properties = None
    if runs:
        current = runs[0].find(f"{{{W}}}rPr")
        if current is not None:
            run_properties = copy.deepcopy(current)
    for child in list(paragraph):
        if child.tag != f"{{{W}}}pPr":
            paragraph.remove(child)
    run = ET.SubElement(paragraph, f"{{{W}}}r")
    if run_properties is not None:
        run.append(run_properties)
    text = ET.SubElement(run, f"{{{W}}}t")
    text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = value


def _update_calculation_rows(root: ET.Element, result: dict[str, Any]) -> None:
    comparables = result["comparables"]
    for table in root.findall(f".//{{{W}}}tbl"):
        for row in table.findall(f"{{{W}}}tr"):
            cells = row.findall(f"{{{W}}}tc")
            if not cells:
                continue
            label = _cell_text(cells[0])
            data = cells[1:4]
            if label == "交易单价":
                for cell, item in zip(data, comparables):
                    _set_cell_text(cell, item["transaction_unit_price"])
            elif label == "P1交易情况修正":
                for cell, item in zip(data, comparables):
                    _set_cell_text(cell, item["p1"])
            elif label == "P2交易日期修正":
                for cell, item in zip(data, comparables):
                    _set_cell_text(cell, item["p2"])
            elif label == "P3实体因素修正":
                subject_physical = result["subject"]["factor_breakdown"]["physical"]
                for cell, item in zip(data, comparables):
                    _set_cell_text(
                        cell,
                        f"{subject_physical}/{item['factor_breakdown']['physical']}",
                    )
            elif label == "P4区位状况修正":
                for cell, item in zip(data, comparables):
                    _set_cell_text(cell, item["p4"])
            elif label == "P1×P2×P3×P4结果":
                for cell, item in zip(data, comparables):
                    _set_cell_text(cell, item["combined_factor"])
            elif label.startswith("Vs×P1×P2×P3×P4结果"):
                for cell, item in zip(data, comparables):
                    _set_cell_text(cell, f"{Decimal(item['adjusted_amount']):,.0f}")
            elif label == "单位面积装修重置价":
                for cell, item in zip(data, comparables):
                    _set_cell_text(cell, f"{Decimal(item['decoration_deduction']):,.0f}")
            elif label == "单位面积附属物单价":
                for cell, item in zip(data, comparables):
                    _set_cell_text(cell, f"{Decimal(item['appurtenance_deduction']):,.0f}")
            elif label == "比准价格（元/㎡）":
                for cell, item in zip(data, comparables):
                    _set_cell_text(cell, f"{Decimal(item['comparable_unit_price']):,.0f}")
            elif label == "评估单价（元/㎡）":
                _set_cell_text(data[0], f"{Decimal(result['estimated_unit_price']):,.0f}")
            elif label == "评估总价（元）":
                _set_cell_text(data[0], f"{Decimal(result['estimated_total']):,.0f}")
            elif label == "P4区位状况修正结果":
                _set_cell_text(cells[1], "100%")
                for cell, item in zip(cells[2:5], comparables):
                    _set_cell_text(cell, f"{Decimal(item['p4']) * 100:.1f}%")
            elif label == "P1交易情况修正结果":
                _set_cell_text(cells[1], "100%")
                for cell, item in zip(cells[2:5], comparables):
                    _set_cell_text(cell, f"{Decimal(item['p1']) * 100:.1f}%")
            elif label == "P2交易日期修正结果":
                _set_cell_text(cells[1], "100%")
                for cell, item in zip(cells[2:5], comparables):
                    _set_cell_text(cell, f"{Decimal(item['p2']) * 100:.1f}%")


def _insert_floor_note(root: ET.Element, project: dict[str, Any], result: dict[str, Any]) -> None:
    subject = project.get("subject") or {}
    above = int(subject.get("above_ground_total_floors") or 0)
    floor = int(subject.get("floor") or 0)
    below = int(subject.get("below_ground_floors") or 0)
    note = (
        f"注：估价对象楼层口径为“{subject.get('certificate_floor_display') or ''}"
        f"（地下{_chinese_small_number(below)}层、地上{_chinese_small_number(above)}层、"
        f"{'顶层' if floor == above else '非顶层'}）”；实际位于地上第{_chinese_small_number(floor)}层，"
        f"属{'顶层' if floor == above else '非顶层'}。本次住宅层次修正按地上"
        f"{floor}/{above}层计算，层次修正系数为{float(result['subject']['level_factor_pct']):.1f}%。"
    )
    for table in root.findall(f".//{{{W}}}tbl"):
        table_text = "".join(_paragraph_text(p) for p in table.findall(f".//{{{W}}}p"))
        if "建筑结构及所在层数" not in table_text or "钢混结构" not in table_text:
            continue
        body = root.find(f"{{{W}}}body")
        if body is None:
            return
        candidates = [
            paragraph
            for paragraph in root.findall(f".//{{{W}}}p")
            if "估价对象范围包含证载面积房地产" in _paragraph_text(paragraph)
        ]
        template_paragraph = candidates[0] if candidates else None
        if template_paragraph is None:
            return
        paragraph = copy.deepcopy(template_paragraph)
        runs = paragraph.findall(f".//{{{W}}}r")
        if not runs:
            return
        first_run = runs[0]
        for child in list(first_run):
            if child.tag != f"{{{W}}}rPr":
                first_run.remove(child)
        text = ET.SubElement(first_run, f"{{{W}}}t")
        text.text = note
        for extra in runs[1:]:
            parent = next((parent for parent in paragraph.iter() if extra in list(parent)), None)
            if parent is not None:
                parent.remove(extra)
        children = list(body)
        table_index = children.index(table)
        body.insert(table_index + 1, paragraph)
        return


def _remove_signature_images(root: ET.Element) -> tuple[set[str], set[str]]:
    relationship_ids: set[str] = set()
    for table in root.findall(f".//{{{W}}}tbl"):
        text = "".join(node.text or "" for node in table.findall(f".//{{{W}}}t"))
        if "3220070051" not in text or "签名" not in text:
            continue
        for blip in table.findall(f".//{{{A}}}blip"):
            relationship_id = blip.get(f"{{{R}}}embed")
            if relationship_id:
                relationship_ids.add(relationship_id)
        for drawing in table.findall(f".//{{{W}}}drawing"):
            for parent in table.iter():
                if drawing in list(parent):
                    parent.remove(drawing)
                    break
    return relationship_ids, set()


def _remove_relationship_targets(
    relationships_xml: bytes,
    relationship_ids: set[str],
) -> tuple[bytes, set[str]]:
    root = ET.fromstring(relationships_xml)
    targets: set[str] = set()
    for relationship in list(root.findall(f"{{{P}}}Relationship")):
        if relationship.get("Id") not in relationship_ids:
            continue
        target = relationship.get("Target") or ""
        if target:
            targets.add(("word/" + target.replace("\\", "/")).lstrip("/"))
        root.remove(relationship)
    ET.register_namespace("", P)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), targets


def build_template_report(
    template_bytes: bytes,
    project: dict[str, Any],
    result: dict[str, Any],
) -> bytes:
    try:
        source = zipfile.ZipFile(io.BytesIO(template_bytes))
    except zipfile.BadZipFile as exc:
        raise ReportTemplateError("Word 模板不是有效的 DOCX 文件") from exc
    if "word/document.xml" not in source.namelist():
        raise ReportTemplateError("Word 模板缺少正文")

    baseline = _load_baseline()
    subject = project.get("subject") or {}
    full_address = str(subject.get("address_full") or subject.get("address") or "")
    short_address = str(
        subject.get("address_short")
        or (full_address.removeprefix("南京市") if full_address.startswith("南京市") else full_address)
    )
    total_value = f"{Decimal(result['estimated_total']):,.0f}"
    unit_value = f"{Decimal(result['estimated_unit_price']):,.0f}"
    uppercase = result["total_uppercase"].removeprefix("人民币").removesuffix("整")
    replacements = [
        (baseline.get("address_full", ""), full_address),
        (baseline.get("address_short", ""), short_address),
        (baseline.get("certificate_no", ""), str(subject.get("certificate_no") or baseline.get("certificate_no", ""))),
        (baseline.get("owner", ""), str(subject.get("owner") or baseline.get("owner", ""))),
        ("单独所有", str(subject.get("co_ownership") or "单独所有")),
        (baseline.get("unit_no", ""), str(subject.get("real_estate_unit_no") or baseline.get("unit_no", ""))),
        ("国有建设用地使用权/房屋所有权", str(subject.get("rights_type") or "国有建设用地使用权/房屋所有权")),
        ("出让/易地搬迁安置房", str(subject.get("rights_nature") or "出让/易地搬迁安置房")),
        ("城镇住宅用地/成套住宅", f"{subject.get('land_use') or '城镇住宅用地'}/{subject.get('use_type') or '成套住宅'}"),
        ("32543.20", f"{Decimal(str(subject.get('land_area') or '32543.20')):,.2f}"),
        ("钢混结构", str(subject.get("structure") or "钢混结构")),
        (baseline.get("rights_term", ""), str(subject.get("rights_term") or baseline.get("rights_term", ""))),
        (baseline.get("parcel_no", ""), str(subject.get("parcel_no") or baseline.get("parcel_no", ""))),
        (baseline.get("boundaries", ""), str(subject.get("location_boundaries") or baseline.get("boundaries", ""))),
        (baseline.get("location_description", ""), str(subject.get("location_description") or baseline.get("location_description", ""))),
        (baseline.get("location_analysis", ""), str(subject.get("location_analysis") or baseline.get("location_analysis", ""))),
        ("337,454", total_value),
        ("叁拾叁万柒仟肆佰伍拾肆元", uppercase),
        ("2,779", unit_value),
    ]
    if subject.get("report_number"):
        replacements.append((baseline.get("report_number", ""), str(subject["report_number"])))
    if subject.get("dispatch_number"):
        replacements.append(("LS-ZY20260120", str(subject["dispatch_number"])))
    if subject.get("report_title"):
        replacements.append(
            (
                baseline.get("report_title", ""),
                str(subject["report_title"]),
            )
        )
    comparables = project.get("comparables") or []
    p1_values = [str(item.get("transaction_condition_factor", 1)) for item in comparables]
    p2_values = [str(item.get("market_condition_factor", 1)) for item in comparables]
    if any(Decimal(value) != Decimal("1") for value in p1_values):
        details = "；".join(
            f"{item.get('name') or f'可比实例{chr(65 + index)}'}为"
            f"{Decimal(str(item.get('transaction_condition_factor', 1))):.3f}"
            f"（{item.get('transaction_condition_note') or '应补充修正依据'}）"
            for index, item in enumerate(comparables)
        )
        replacements.append(
            (baseline.get("transaction_note", ""), f"本次交易情况修正如下：{details}。")
        )
    if any(Decimal(value) != Decimal("1") for value in p2_values):
        details = "；".join(
            f"{item.get('name') or f'可比实例{chr(65 + index)}'}为"
            f"{Decimal(str(item.get('market_condition_factor', 1))):.3f}"
            f"（{item.get('market_condition_note') or '应补充修正依据'}）"
            for index, item in enumerate(comparables)
        )
        replacements.append(
            (baseline.get("market_note", ""), f"本次市场状况修正如下：{details}。")
        )
    valuation_date = _parse_iso_date(subject.get("valuation_date"))
    if valuation_date:
        replacements.extend(
            [
                (
                    "2026年08月28日",
                    f"{valuation_date.year}年{valuation_date.month:02d}月{valuation_date.day:02d}日",
                ),
                (
                    "2026/08/28",
                    f"{valuation_date.year}/{valuation_date.month:02d}/{valuation_date.day:02d}",
                ),
            ]
        )

    document_xml = _replace_in_part(source.read("word/document.xml"), replacements)
    root = ET.fromstring(document_xml)
    _update_calculation_rows(root, result)
    _insert_floor_note(root, project, result)
    signature_relationship_ids, _ = _remove_signature_images(root)
    document_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    relationships_path = "word/_rels/document.xml.rels"
    updated_relationships, removed_media = _remove_relationship_targets(
        source.read(relationships_path),
        signature_relationship_ids,
    )

    target = io.BytesIO()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as package:
        for info in source.infolist():
            if info.filename in removed_media:
                continue
            content = source.read(info.filename)
            if info.filename == "word/document.xml":
                content = document_xml
            elif info.filename == relationships_path:
                content = updated_relationships
            elif info.filename.startswith("word/header") or info.filename.startswith("word/footer"):
                if info.filename.endswith(".xml"):
                    content = _replace_in_part(content, replacements)
            package.writestr(info, content)
    source.close()
    return target.getvalue()
