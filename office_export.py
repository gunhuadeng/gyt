from __future__ import annotations

import io
import zipfile
from datetime import datetime, timezone
from decimal import Decimal
from html import escape
from typing import Any
from xml.etree import ElementTree as ET


SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPE_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"


def _xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _zip_bytes(parts: dict[str, bytes]) -> bytes:
    target = io.BytesIO()
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as package:
        for path, content in parts.items():
            package.writestr(path, content)
    return target.getvalue()


def _core_properties(title: str) -> bytes:
    ET.register_namespace("cp", CORE_NS)
    ET.register_namespace("dc", DC_NS)
    ET.register_namespace("dcterms", DCTERMS_NS)
    ET.register_namespace("xsi", XSI_NS)
    root = ET.Element(f"{{{CORE_NS}}}coreProperties")
    ET.SubElement(root, f"{{{DC_NS}}}title").text = title
    ET.SubElement(root, f"{{{DC_NS}}}creator").text = "房地产估价报告原型"
    ET.SubElement(root, f"{{{CORE_NS}}}lastModifiedBy").text = "房地产估价报告原型"
    created = ET.SubElement(root, f"{{{DCTERMS_NS}}}created")
    created.set(f"{{{XSI_NS}}}type", "dcterms:W3CDTF")
    created.text = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return _xml_bytes(root)


def _column_name(index: int) -> str:
    text = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        text = chr(65 + remainder) + text
    return text


def _xlsx_cell(value: Any, style: int = 0, formula: str | None = None) -> str:
    style_attr = f' s="{style}"' if style else ""
    if formula:
        cached = escape(str(value), quote=True)
        return f'<c r="{{ref}}"{style_attr}><f>{escape(formula)}</f><v>{cached}</v></c>'
    if value is None or value == "":
        return f'<c r="{{ref}}"{style_attr}/>'
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        return f'<c r="{{ref}}"{style_attr}><v>{escape(str(value))}</v></c>'
    text = escape(str(value), quote=True)
    return f'<c r="{{ref}}" t="inlineStr"{style_attr}><is><t xml:space="preserve">{text}</t></is></c>'


def _worksheet_xml(
    rows: list[list[tuple[Any, int] | tuple[Any, int, str]]],
    merges: list[str] | None = None,
    widths: dict[str, float] | None = None,
    freeze_rows: int = 0,
) -> bytes:
    ET.register_namespace("", SHEET_NS)
    root = ET.Element(f"{{{SHEET_NS}}}worksheet")
    ET.SubElement(root, f"{{{SHEET_NS}}}sheetPr")
    ET.SubElement(root, f"{{{SHEET_NS}}}dimension", {"ref": f"A1:{_column_name(max(len(row) for row in rows))}{len(rows)}"})

    if freeze_rows:
        views = ET.SubElement(root, f"{{{SHEET_NS}}}sheetViews")
        view = ET.SubElement(views, f"{{{SHEET_NS}}}sheetView", {"workbookViewId": "0"})
        pane = ET.SubElement(
            view,
            f"{{{SHEET_NS}}}pane",
            {
                "ySplit": str(freeze_rows),
                "topLeftCell": f"A{freeze_rows + 1}",
                "activePane": "bottomLeft",
                "state": "frozen",
            },
        )
        pane.text = ""

    ET.SubElement(root, f"{{{SHEET_NS}}}sheetFormatPr", {"defaultRowHeight": "19"})

    if widths:
        columns = ET.SubElement(root, f"{{{SHEET_NS}}}cols")
        for column, width in widths.items():
            ET.SubElement(
                columns,
                f"{{{SHEET_NS}}}col",
                {"min": column, "max": column, "width": str(width), "customWidth": "1"},
            )

    sheet_data = ET.SubElement(root, f"{{{SHEET_NS}}}sheetData")
    for row_index, row in enumerate(rows, start=1):
        row_element = ET.SubElement(sheet_data, f"{{{SHEET_NS}}}row", {"r": str(row_index)})
        for column_index, cell in enumerate(row, start=1):
            value, style = cell[:2]
            formula = cell[2] if len(cell) > 2 else None
            reference = f"{_column_name(column_index)}{row_index}"
            cell_xml = _xlsx_cell(value, style, formula).replace("{ref}", reference)
            row_element.append(ET.fromstring(cell_xml))

    if merges:
        merge_cells = ET.SubElement(root, f"{{{SHEET_NS}}}mergeCells", {"count": str(len(merges))})
        for merge in merges:
            ET.SubElement(merge_cells, f"{{{SHEET_NS}}}mergeCell", {"ref": merge})

    ET.SubElement(root, f"{{{SHEET_NS}}}pageMargins", {"left": "0.4", "right": "0.4", "top": "0.6", "bottom": "0.6", "header": "0.2", "footer": "0.2"})
    return _xml_bytes(root)


def _xlsx_styles() -> bytes:
    ET.register_namespace("", SHEET_NS)
    root = ET.Element(f"{{{SHEET_NS}}}styleSheet")
    fonts = ET.SubElement(root, f"{{{SHEET_NS}}}fonts", {"count": "3"})
    for name, size, bold, color in (
        ("宋体", 11, False, "1F2937"),
        ("宋体", 11, True, "FFFFFF"),
        ("宋体", 16, True, "0F172A"),
    ):
        font = ET.SubElement(fonts, f"{{{SHEET_NS}}}font")
        ET.SubElement(font, f"{{{SHEET_NS}}}name", {"val": name})
        ET.SubElement(font, f"{{{SHEET_NS}}}sz", {"val": str(size)})
        ET.SubElement(font, f"{{{SHEET_NS}}}color", {"rgb": "FF" + color})
        if bold:
            ET.SubElement(font, f"{{{SHEET_NS}}}b")

    fills = ET.SubElement(root, f"{{{SHEET_NS}}}fills", {"count": "4"})
    for pattern, foreground in (("none", None), ("gray125", None), ("solid", "FF0F766E"), ("solid", "FFDDF4EF")):
        fill = ET.SubElement(fills, f"{{{SHEET_NS}}}fill")
        pattern_fill = ET.SubElement(fill, f"{{{SHEET_NS}}}patternFill", {"patternType": pattern})
        if foreground:
            ET.SubElement(pattern_fill, f"{{{SHEET_NS}}}fgColor", {"rgb": foreground})
            ET.SubElement(pattern_fill, f"{{{SHEET_NS}}}bgColor", {"indexed": "64"})

    borders = ET.SubElement(root, f"{{{SHEET_NS}}}borders", {"count": "2"})
    empty_border = ET.SubElement(borders, f"{{{SHEET_NS}}}border")
    for side in ("left", "right", "top", "bottom", "diagonal"):
        ET.SubElement(empty_border, f"{{{SHEET_NS}}}{side}")
    border = ET.SubElement(borders, f"{{{SHEET_NS}}}border")
    for side in ("left", "right", "top", "bottom"):
        element = ET.SubElement(border, f"{{{SHEET_NS}}}{side}", {"style": "thin"})
        ET.SubElement(element, f"{{{SHEET_NS}}}color", {"rgb": "FFD1D5DB"})
    ET.SubElement(border, f"{{{SHEET_NS}}}diagonal")

    cell_xfs = ET.SubElement(root, f"{{{SHEET_NS}}}cellXfs", {"count": "7"})
    configurations = (
        (0, 0, False, False),
        (1, 2, True, True),
        (0, 3, False, True),
        (0, 0, True, True),
        (2, 0, True, False),
        (0, 0, False, True),
        (3, 3, True, True),
    )
    for font_id, fill_id, wrap, border_id in configurations:
        xf = ET.SubElement(
            cell_xfs,
            f"{{{SHEET_NS}}}xf",
            {
                "numFmtId": "4" if fill_id == 3 else "0",
                "fontId": str(font_id),
                "fillId": str(fill_id),
                "borderId": "1" if border_id else "0",
                "xfId": "0",
                "applyAlignment": "1",
                "applyBorder": "1" if border_id else "0",
            },
        )
        ET.SubElement(
            xf,
            f"{{{SHEET_NS}}}alignment",
            {"vertical": "center", "wrapText": "1" if wrap else "0"},
        )
    ET.SubElement(root, f"{{{SHEET_NS}}}cellStyles", {"count": "1"})
    ET.SubElement(root, f"{{{SHEET_NS}}}dxfs", {"count": "0"})
    ET.SubElement(root, f"{{{SHEET_NS}}}tableStyles", {"count": "0", "defaultTableStyle": "TableStyleMedium2"})
    return _xml_bytes(root)


def build_xlsx(payload: dict[str, Any], result: dict[str, Any]) -> bytes:
    subject = payload.get("subject") or {}
    title = str(subject.get("report_title") or "房地产估价报告测算底稿")
    comps = result["comparables"]

    overview_rows: list[list[tuple[Any, int] | tuple[Any, int, str]]] = [
        [(title, 4), ("", 4), ("", 4), ("", 4)],
        [("估价对象", 1), (str(subject.get("owner") or ""), 5), ("估价目的", 1), (str(subject.get("purpose") or ""), 5)],
        [("坐落", 1), (str(subject.get("address") or ""), 5), ("估价基准日", 1), (str(subject.get("valuation_date") or ""), 5)],
        [("建筑面积（㎡）", 1), (float(subject.get("area") or 0), 3), ("估算单价（元/㎡）", 1), (float(result["weighted_unit_price"]), 3)],
        [("估算总价（元）", 1), (float(result["subject_value"]), 3), ("取整单位（元）", 1), (float(result["rounding_step"]), 3)],
        [("总价大写", 1), (result["subject_value_uppercase"], 5), ("结果离散度", 1), (result["spread_pct"] + "%", 5)],
        [("测算方法", 2), ("市场比较法：案例单价经交易时间、区位、实物、权益和其他因素综合修正后，按权重加权确定估价对象价值。", 4), ("", 4), ("", 4)],
        [("复核提示", 2), ("本工作簿为软件自动生成草稿。正式报告须由具备资质的估价专业人员复核数据、参数、修正幅度和最终结论。", 4), ("", 4), ("", 4)],
    ]
    for warning in result.get("warnings", []):
        overview_rows.append([(warning, 4), ("", 4), ("", 4), ("", 4)])

    detail_headers = [
        "案例",
        "成交日期",
        "成交总价（元）",
        "建筑面积（㎡）",
        "成交单价（元/㎡）",
        "交易时间",
        "区位状况",
        "实物状况",
        "权益状况",
        "其他因素",
        "综合修正",
        "修正后单价（元/㎡）",
        "权重",
    ]
    detail_rows: list[list[tuple[Any, int] | tuple[Any, int, str]]] = [
        [("市场比较法测算", 4)] + [("", 4) for _ in range(12)],
        [(header, 1) for header in detail_headers],
    ]
    for comp in comps:
        detail_rows.append(
            [
                (comp["name"], 5),
                (comp["transaction_date"], 5),
                (float(comp["transaction_total"]), 3),
                (float(comp["area"]), 3),
                (float(comp["base_unit_price"]), 3),
                (comp["adjustments"][0]["percent"] + "%", 5),
                (comp["adjustments"][1]["percent"] + "%", 5),
                (comp["adjustments"][2]["percent"] + "%", 5),
                (comp["adjustments"][3]["percent"] + "%", 5),
                (comp["adjustments"][4]["percent"] + "%", 5),
                (comp["total_adjustment_pct"] + "%", 5),
                (float(comp["adjusted_unit_price"]), 3),
                (comp["normalized_weight_pct"] + "%", 5),
            ]
        )
    detail_rows.extend(
        [
            [("加权平均单价（元/㎡）", 1), (float(result["weighted_unit_price"]), 3)] + [("", 5) for _ in range(11)],
            [("估价对象建筑面积（㎡）", 1), (float(subject.get("area") or 0), 3)] + [("", 5) for _ in range(11)],
            [("未取整总价（元）", 1), (float(result["raw_subject_value"]), 3)] + [("", 5) for _ in range(11)],
            [("最终取整总价（元）", 1), (float(result["subject_value"]), 3)] + [("", 5) for _ in range(11)],
        ]
    )

    adjustment_labels = ["交易时间", "区位状况", "实物状况", "权益状况", "其他因素"]
    adjustment_rows: list[list[tuple[Any, int] | tuple[Any, int, str]]] = [
        [("修正因素明细", 4), ("", 4), ("", 4), ("", 4)],
        *[[(label, 1), ("向估价对象修正百分比", 1), ("说明", 1), ("", 1)] for label in []],
        [("因素", 1), ("案例", 1), ("修正百分比", 1), ("方向说明", 1)],
    ]
    for comp in comps:
        for item in comp["adjustments"]:
            percent = Decimal(item["percent"])
            direction = "向上修正" if percent > 0 else ("向下修正" if percent < 0 else "不修正")
            adjustment_rows.append([(item["label"], 5), (comp["name"], 5), (item["percent"] + "%", 5), (direction, 5)])

    content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CONTENT_TYPE_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>""".encode("utf-8")

    root_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>""".encode("utf-8")

    workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{SHEET_NS}" xmlns:r="{REL_NS}">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="24000" windowHeight="14000"/></bookViews>
  <sheets>
    <sheet name="项目概览" sheetId="1" r:id="rId1"/>
    <sheet name="比较法测算" sheetId="2" r:id="rId2"/>
    <sheet name="修正明细" sheetId="3" r:id="rId3"/>
  </sheets>
  <calcPr calcId="0" fullCalcOnLoad="1"/>
</workbook>""".encode("utf-8")

    workbook_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""".encode("utf-8")

    parts = {
        "[Content_Types].xml": content_types,
        "_rels/.rels": root_rels,
        "docProps/core.xml": _core_properties(title),
        "xl/workbook.xml": workbook_xml,
        "xl/_rels/workbook.xml.rels": workbook_rels,
        "xl/styles.xml": _xlsx_styles(),
        "xl/worksheets/sheet1.xml": _worksheet_xml(
            overview_rows,
            merges=["A1:D1", "B7:D7", "B8:D8"],
            widths={"1": 22, "2": 34, "3": 22, "4": 34},
        ),
        "xl/worksheets/sheet2.xml": _worksheet_xml(
            detail_rows,
            merges=["A1:M1", "A12:B12", "A13:B13", "A14:B14", "A15:B15"],
            widths={"1": 17, "2": 14, "3": 18, "4": 16, "5": 18, **{str(i): 12 for i in range(6, 14)}},
            freeze_rows=2,
        ),
        "xl/worksheets/sheet3.xml": _worksheet_xml(
            adjustment_rows,
            merges=["A1:D1"],
            widths={"1": 18, "2": 16, "3": 18, "4": 18},
            freeze_rows=2,
        ),
    }
    return _zip_bytes(parts)


def _w_run(text: str, bold: bool = False, size: int = 24, color: str = "1F2937") -> ET.Element:
    run = ET.Element(f"{{{WORD_NS}}}r")
    props = ET.SubElement(run, f"{{{WORD_NS}}}rPr")
    fonts = ET.SubElement(props, f"{{{WORD_NS}}}rFonts")
    fonts.set(f"{{{WORD_NS}}}ascii", "Arial")
    fonts.set(f"{{{WORD_NS}}}hAnsi", "Arial")
    fonts.set(f"{{{WORD_NS}}}eastAsia", "宋体")
    ET.SubElement(props, f"{{{WORD_NS}}}color", {f"{{{WORD_NS}}}val": color})
    ET.SubElement(props, f"{{{WORD_NS}}}sz", {f"{{{WORD_NS}}}val": str(size)})
    if bold:
        ET.SubElement(props, f"{{{WORD_NS}}}b")
    text_element = ET.SubElement(run, f"{{{WORD_NS}}}t")
    text_element.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_element.text = text
    return run


def _w_paragraph(
    text: str = "",
    *,
    style: str | None = None,
    bold: bool = False,
    size: int = 24,
    color: str = "1F2937",
    align: str | None = None,
    space_after: int = 120,
) -> ET.Element:
    paragraph = ET.Element(f"{{{WORD_NS}}}p")
    props = ET.SubElement(paragraph, f"{{{WORD_NS}}}pPr")
    if style:
        ET.SubElement(props, f"{{{WORD_NS}}}pStyle", {f"{{{WORD_NS}}}val": style})
    if align:
        ET.SubElement(props, f"{{{WORD_NS}}}jc", {f"{{{WORD_NS}}}val": align})
    ET.SubElement(props, f"{{{WORD_NS}}}spacing", {f"{{{WORD_NS}}}after": str(space_after), f"{{{WORD_NS}}}line": "360"})
    if text:
        paragraph.append(_w_run(text, bold=bold, size=size, color=color))
    return paragraph


def _w_table(headers: list[str], rows: list[list[str]], widths: list[int]) -> ET.Element:
    table = ET.Element(f"{{{WORD_NS}}}tbl")
    props = ET.SubElement(table, f"{{{WORD_NS}}}tblPr")
    ET.SubElement(props, f"{{{WORD_NS}}}tblW", {f"{{{WORD_NS}}}w": "0", f"{{{WORD_NS}}}type": "auto"})
    borders = ET.SubElement(props, f"{{{WORD_NS}}}tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge = ET.SubElement(borders, f"{{{WORD_NS}}}{side}", {f"{{{WORD_NS}}}val": "single", f"{{{WORD_NS}}}sz": "4", f"{{{WORD_NS}}}color": "D1D5DB"})
        edge.text = ""
    layout = ET.SubElement(props, f"{{{WORD_NS}}}tblLayout", {f"{{{WORD_NS}}}type": "fixed"})
    grid = ET.SubElement(table, f"{{{WORD_NS}}}tblGrid")
    for width in widths:
        ET.SubElement(grid, f"{{{WORD_NS}}}gridCol", {f"{{{WORD_NS}}}w": str(width)})

    def add_row(values: list[str], header: bool = False) -> None:
        row = ET.SubElement(table, f"{{{WORD_NS}}}tr")
        for index, value in enumerate(values):
            cell = ET.SubElement(row, f"{{{WORD_NS}}}tc")
            cell_props = ET.SubElement(cell, f"{{{WORD_NS}}}tcPr")
            ET.SubElement(cell_props, f"{{{WORD_NS}}}tcW", {f"{{{WORD_NS}}}w": str(widths[index]), f"{{{WORD_NS}}}type": "dxa"})
            shading = ET.SubElement(cell_props, f"{{{WORD_NS}}}shd", {f"{{{WORD_NS}}}fill": "0F766E" if header else "F8FAFC"})
            shading.text = ""
            paragraph = _w_paragraph(
                value,
                bold=header,
                size=20,
                align="center" if header else None,
                space_after=0,
            )
            cell.append(paragraph)

    add_row(headers, header=True)
    for row in rows:
        add_row(row)
    return table


def _docx_styles() -> bytes:
    ET.register_namespace("w", WORD_NS)
    root = ET.Element(f"{{{WORD_NS}}}styles")
    normal = ET.SubElement(root, f"{{{WORD_NS}}}style", {f"{{{WORD_NS}}}type": "paragraph", f"{{{WORD_NS}}}default": "1", f"{{{WORD_NS}}}styleId": "Normal"})
    ET.SubElement(normal, f"{{{WORD_NS}}}name", {f"{{{WORD_NS}}}val": "Normal"})
    ET.SubElement(normal, f"{{{WORD_NS}}}qFormat")
    heading = ET.SubElement(root, f"{{{WORD_NS}}}style", {f"{{{WORD_NS}}}type": "paragraph", f"{{{WORD_NS}}}styleId": "Heading1"})
    ET.SubElement(heading, f"{{{WORD_NS}}}name", {f"{{{WORD_NS}}}val": "heading 1"})
    ET.SubElement(heading, f"{{{WORD_NS}}}basedOn", {f"{{{WORD_NS}}}val": "Normal"})
    ET.SubElement(heading, f"{{{WORD_NS}}}next", {f"{{{WORD_NS}}}val": "Normal"})
    ET.SubElement(heading, f"{{{WORD_NS}}}qFormat")
    return _xml_bytes(root)


def build_docx(payload: dict[str, Any], result: dict[str, Any]) -> bytes:
    subject = payload.get("subject") or {}
    title = str(subject.get("report_title") or "房地产估价报告")
    ET.register_namespace("w", WORD_NS)
    document = ET.Element(f"{{{WORD_NS}}}document")
    body = ET.SubElement(document, f"{{{WORD_NS}}}body")

    body.append(_w_paragraph(title, bold=True, size=36, align="center", space_after=240))
    body.append(_w_paragraph("房地产估价报告草稿", bold=True, size=26, align="center", color="0F766E"))
    body.append(_w_paragraph("（软件生成原型，仅供专业估价人员复核）", size=20, align="center", color="B45309", space_after=320))

    body.append(_w_paragraph("一、估价对象概况", style="Heading1", bold=True, size=28))
    overview_rows = [
        ["产权人/委托方", str(subject.get("owner") or "—"), "估价目的", str(subject.get("purpose") or "—")],
        ["坐落", str(subject.get("address") or "—"), "估价基准日", str(subject.get("valuation_date") or "—")],
        ["用途", str(subject.get("use_type") or "—"), "权利状况", str(subject.get("rights_status") or "—")],
        ["建筑面积", f"{Decimal(str(subject.get('area') or 0)):,.2f} ㎡", "报告编号", str(subject.get("report_number") or "—")],
    ]
    body.append(_w_table(["项目", "内容", "项目", "内容"], overview_rows, [1800, 2900, 1800, 2900]))

    body.append(_w_paragraph("二、估价结果", style="Heading1", bold=True, size=28))
    body.append(
        _w_paragraph(
            f"采用市场比较法测算，估价对象于 {subject.get('valuation_date') or '估价基准日'} 的市场价值为：",
            size=24,
        )
    )
    body.append(_w_paragraph(f"{Decimal(result['subject_value']):,.2f} 元", bold=True, size=34, align="center", color="0F766E"))
    body.append(_w_paragraph(result["subject_value_uppercase"], bold=True, size=24, align="center"))
    body.append(_w_paragraph(f"加权估算单价：{Decimal(result['weighted_unit_price']):,.2f} 元/㎡。", size=24))

    body.append(_w_paragraph("三、市场比较法测算", style="Heading1", bold=True, size=28))
    comp_headers = ["案例", "成交总价（元）", "面积（㎡）", "成交单价（元/㎡）", "综合修正", "修正后单价（元/㎡）", "权重"]
    comp_rows = [
        [
            comp["name"],
            f"{Decimal(comp['transaction_total']):,.2f}",
            f"{Decimal(comp['area']):,.2f}",
            f"{Decimal(comp['base_unit_price']):,.2f}",
            f"{Decimal(comp['total_adjustment_pct']):+.2f}%",
            f"{Decimal(comp['adjusted_unit_price']):,.2f}",
            f"{Decimal(comp['normalized_weight_pct']):.2f}%",
        ]
        for comp in result["comparables"]
    ]
    body.append(_w_table(comp_headers, comp_rows, [1000, 1400, 1100, 1500, 1000, 1600, 900]))

    body.append(_w_paragraph("四、测算说明", style="Heading1", bold=True, size=28))
    body.append(
        _w_paragraph(
            "成交单价由案例成交总价除以建筑面积得到；综合修正系数为各修正因素系数连乘，"
            "修正后单价为成交单价乘以综合修正系数；估价对象价格由修正后单价按案例权重加权，"
            f"乘以估价对象建筑面积，并按 {Decimal(result['rounding_step']):,.0f} 元取整。",
            size=24,
        )
    )
    if result.get("warnings"):
        body.append(_w_paragraph("复核提示", bold=True, size=24, color="B45309"))
        for warning in result["warnings"]:
            body.append(_w_paragraph("• " + warning, size=22, color="B45309"))

    body.append(_w_paragraph("五、使用条件与声明", style="Heading1", bold=True, size=28))
    body.append(
        _w_paragraph(
            "1. 本报告为软件根据用户录入数据自动生成的草稿，不构成由注册房地产估价师签署的正式估价报告。",
            size=22,
        )
    )
    body.append(
        _w_paragraph(
            "2. 交易案例、修正幅度、权重及估价对象信息应逐项核验来源和适用性；软件不替代专业判断。",
            size=22,
        )
    )
    body.append(
        _w_paragraph(
            "3. 正式使用前，应由具备相应资质的专业人员复核计算、检查文字与数据一致性，并按适用规范完成签章流程。",
            size=22,
        )
    )

    section = ET.SubElement(body, f"{{{WORD_NS}}}sectPr")
    ET.SubElement(section, f"{{{WORD_NS}}}pgSz", {f"{{{WORD_NS}}}w": "11906", f"{{{WORD_NS}}}h": "16838"})
    margins = ET.SubElement(section, f"{{{WORD_NS}}}pgMar")
    margins.set(f"{{{WORD_NS}}}top", "1440")
    margins.set(f"{{{WORD_NS}}}right", "1440")
    margins.set(f"{{{WORD_NS}}}bottom", "1440")
    margins.set(f"{{{WORD_NS}}}left", "1440")
    margins.set(f"{{{WORD_NS}}}header", "720")
    margins.set(f"{{{WORD_NS}}}footer", "720")
    margins.set(f"{{{WORD_NS}}}gutter", "0")

    content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CONTENT_TYPE_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>""".encode("utf-8")

    root_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>""".encode("utf-8")

    document_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""".encode("utf-8")

    return _zip_bytes(
        {
            "[Content_Types].xml": content_types,
            "_rels/.rels": root_rels,
            "docProps/core.xml": _core_properties(title),
            "word/document.xml": _xml_bytes(document),
            "word/styles.xml": _docx_styles(),
            "word/_rels/document.xml.rels": document_rels,
        }
    )


def build_residential_docx(project: dict[str, Any], result: dict[str, Any]) -> bytes:
    subject = project.get("subject") or {}
    comparables = project.get("comparables") or []
    title = str(subject.get("report_title") or "住宅房地产市场价值估价报告")
    report_number = str(subject.get("report_number") or "—")
    certificate_floor = str(subject.get("certificate_floor_display") or "—")
    below = int(subject.get("below_ground_floors") or 0)
    above = int(subject.get("above_ground_total_floors") or 0)
    floor = int(subject.get("floor") or 0)
    floor_explanation = (
        f"{certificate_floor}（地下一层、地上{above}层"
        f"{'、顶层' if floor == above else ''}）"
        if below
        else f"{certificate_floor}（地上{above}层{'、顶层' if floor == above else ''}）"
    )

    ET.register_namespace("w", WORD_NS)
    document = ET.Element(f"{{{WORD_NS}}}document")
    body = ET.SubElement(document, f"{{{WORD_NS}}}body")

    body.append(_w_paragraph(title, bold=True, size=36, align="center", space_after=220))
    body.append(_w_paragraph("住宅房地产市场价值估价报告草稿", bold=True, size=26, align="center", color="0F766E"))
    body.append(_w_paragraph(f"报告编号：{report_number}", size=21, align="center", color="475569"))
    body.append(_w_paragraph("（软件生成草稿，须由注册房地产估价师复核）", size=20, align="center", color="B45309", space_after=300))

    body.append(_w_paragraph("一、估价结果", style="Heading1", bold=True, size=28))
    body.append(
        _w_paragraph(
            f"估价对象于 {subject.get('valuation_date') or '价值时点'} 的市场价值为：",
            size=23,
        )
    )
    body.append(
        _w_paragraph(
            f"{Decimal(result['estimated_total']):,.0f} 元",
            bold=True,
            size=34,
            align="center",
            color="0F766E",
        )
    )
    body.append(_w_paragraph(result["total_uppercase"], bold=True, size=24, align="center"))
    body.append(
        _w_paragraph(
            f"评估单价：{Decimal(result['estimated_unit_price']):,.0f} 元/㎡；"
            f"建筑面积：{Decimal(str(subject.get('area') or 0)):,.2f} ㎡。",
            size=23,
        )
    )

    body.append(_w_paragraph("二、估价对象和楼层口径", style="Heading1", bold=True, size=28))
    body.append(
        _w_paragraph(
            f"坐落：{subject.get('address') or '—'}；建筑面积："
            f"{Decimal(str(subject.get('area') or 0)):,.2f} ㎡；"
            f"建筑结构：{subject.get('structure') or '—'}。",
            size=23,
        )
    )
    body.append(
        _w_paragraph(
            f"楼层：{floor_explanation}。层次修正按地上总层数计算，"
            f"本次修正口径为 {floor}/{above}，层次系数为 {result['subject']['level_factor_pct']}%。",
            size=23,
        )
    )

    body.append(_w_paragraph("三、权益状况与区位状况", style="Heading1", bold=True, size=28))
    rights_rows = [
        ["不动产权证号", str(subject.get("certificate_no") or "—"), "产权人", str(subject.get("owner") or "—")],
        ["共有情况", str(subject.get("co_ownership") or "—"), "不动产单元号", str(subject.get("real_estate_unit_no") or "—")],
        ["权利类型", str(subject.get("rights_type") or "—"), "权利性质", str(subject.get("rights_nature") or "—")],
        ["土地用途", str(subject.get("land_use") or "—"), "宗地面积", f"{Decimal(str(subject.get('land_area') or 0)):,.2f} ㎡"],
        ["使用期限", str(subject.get("rights_term") or "—"), "丘权号", str(subject.get("parcel_no") or "—")],
    ]
    body.append(_w_table(["权益项目", "内容", "权益项目", "内容"], rights_rows, [1500, 3150, 1500, 3150]))
    if subject.get("location_boundaries"):
        body.append(_w_paragraph(f"四至：{subject['location_boundaries']}", size=22))
    if subject.get("location_description"):
        body.append(_w_paragraph(f"区位状况：{subject['location_description']}", size=22))
    if subject.get("location_analysis"):
        body.append(_w_paragraph(f"区位分析：{subject['location_analysis']}", size=22))

    body.append(_w_paragraph("四、可比实例与修正参数", style="Heading1", bold=True, size=28))
    body.append(
        _w_table(
            ["案例", "坐落", "面积（㎡）", "成交单价", "P1", "P2"],
            [
                [
                    str(item.get("name") or ""),
                    str(item.get("address") or ""),
                    f"{Decimal(str(item.get('area') or 0)):,.2f}",
                    f"{Decimal(str(item.get('transaction_unit_price') or 0)):,.0f}",
                    str(calculated["p1"]),
                    str(calculated["p2"]),
                ]
                for item, calculated in zip(comparables, result["comparables"])
            ],
            [900, 3300, 1100, 1200, 800, 800],
        )
    )
    body.append(
        _w_paragraph(
            "P1 为交易情况修正，P2 为市场状况修正；两项默认均为1.000，允许估价师依据市场调查独立调整。",
            size=21,
        )
    )

    body.append(_w_paragraph("五、比较法测算明细", style="Heading1", bold=True, size=28))
    body.append(
        _w_table(
            ["案例", "P3", "P4", "综合系数", "修正金额", "装修扣减", "附属物", "比准单价"],
            [
                [
                    item["name"],
                    item["p3"],
                    item["p4"],
                    item["combined_factor"],
                    f"{Decimal(item['adjusted_amount']):,.0f}",
                    f"{Decimal(item['decoration_deduction']):,.0f}",
                    f"{Decimal(item['appurtenance_deduction']):,.0f}",
                    f"{Decimal(item['comparable_unit_price']):,.0f}",
                ]
                for item in result["comparables"]
            ],
            [1050, 800, 800, 1100, 1300, 1100, 900, 1300],
        )
    )
    body.append(
        _w_paragraph(
            "评估单价为三个比准价格的算术平均值；评估总价为评估单价乘以估价对象建筑面积后取整。",
            size=22,
        )
    )
    comparable_prices = "、".join(
        f"{Decimal(item['comparable_unit_price']):,.0f}"
        for item in result["comparables"]
    )
    body.append(
        _w_paragraph(
            f"三个比准价格为 {comparable_prices} 元/㎡；"
            f"评估单价为 {Decimal(result['estimated_unit_price']):,.0f} 元/㎡；"
            f"评估总价为 {Decimal(result['estimated_total']):,.0f} 元。",
            bold=True,
            size=23,
        )
    )

    notes = []
    for item in comparables:
        name = str(item.get("name") or "可比实例")
        if Decimal(str(item.get("transaction_condition_factor", 1))) != Decimal("1"):
            note = str(item.get("transaction_condition_note") or "未填写具体依据，应补充复核")
            notes.append(f"{name} 交易情况修正：{note}。")
        if Decimal(str(item.get("market_condition_factor", 1))) != Decimal("1"):
            note = str(item.get("market_condition_note") or "未填写具体依据，应补充复核")
            notes.append(f"{name} 市场状况修正：{note}。")
    if result.get("warnings"):
        notes.extend(str(item) for item in result["warnings"])
    if notes:
        body.append(_w_paragraph("六、复核说明", style="Heading1", bold=True, size=28))
        for note in notes:
            body.append(_w_paragraph("• " + note, size=22, color="B45309"))

    body.append(_w_paragraph("七、使用限制", style="Heading1", bold=True, size=28))
    body.append(
        _w_paragraph(
            "1. 本报告为软件根据用户录入数据自动生成的测算草稿，不是由注册房地产估价师签署的正式报告。",
            size=21,
        )
    )
    body.append(
        _w_paragraph(
            "2. 产权信息、案例来源、区域因素、修正系数及市场价格指数应由估价专业人员复核。",
            size=21,
        )
    )
    body.append(
        _w_paragraph(
            "3. 正式使用前，应完成现场查勘、报告文字核对、金额一致性检查及签章程序。",
            size=21,
        )
    )

    section = ET.SubElement(body, f"{{{WORD_NS}}}sectPr")
    ET.SubElement(section, f"{{{WORD_NS}}}pgSz", {f"{{{WORD_NS}}}w": "11906", f"{{{WORD_NS}}}h": "16838"})
    margins = ET.SubElement(section, f"{{{WORD_NS}}}pgMar")
    for side in ("top", "right", "bottom", "left"):
        margins.set(f"{{{WORD_NS}}}{side}", "1440")
    margins.set(f"{{{WORD_NS}}}header", "720")
    margins.set(f"{{{WORD_NS}}}footer", "720")
    margins.set(f"{{{WORD_NS}}}gutter", "0")

    content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CONTENT_TYPE_NS}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>""".encode("utf-8")
    root_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>""".encode("utf-8")
    document_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{PKG_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""".encode("utf-8")
    return _zip_bytes(
        {
            "[Content_Types].xml": content_types,
            "_rels/.rels": root_rels,
            "docProps/core.xml": _core_properties(title),
            "word/document.xml": _xml_bytes(document),
            "word/styles.xml": _docx_styles(),
            "word/_rels/document.xml.rels": document_rels,
        }
    )
