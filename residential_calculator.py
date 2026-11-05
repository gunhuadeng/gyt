from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from typing import Any

from residential_rules import (
    EAST_WEST_FACTORS,
    ELEVATOR_MULTI_STORY_LEVELS,
    MULTI_STORY_LEVELS,
    ORIENTATION_FACTORS,
    REGIONAL_FACTORS,
    STRUCTURE_FACTORS,
    high_rise_level_factor,
)


CN_DIGITS = "零壹贰叁肆伍陆柒捌玖"
CN_UNITS = ("", "拾", "佰", "仟")
CN_GROUP_UNITS = ("", "万", "亿", "万亿")


class ResidentialCalculationError(ValueError):
    pass


def _decimal(value: Any, label: str) -> Decimal:
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError) as exc:
        raise ResidentialCalculationError(f"{label}必须是有效数字") from exc
    if not number.is_finite():
        raise ResidentialCalculationError(f"{label}必须是有效数字")
    return number


def _rounded(value: Decimal, places: int) -> Decimal:
    quantum = Decimal("1").scaleb(-places)
    return value.quantize(quantum, rounding=ROUND_HALF_UP)


def _plain(value: Decimal, places: int = 0) -> str:
    return f"{_rounded(value, places):.{places}f}"


def _four_digit_upper(group: int) -> str:
    text = ""
    pending_zero = False
    for position in range(3, -1, -1):
        divisor = 10**position
        digit = group // divisor % 10
        if digit == 0:
            if text and group % divisor:
                pending_zero = True
            continue
        if pending_zero and not text.endswith("零"):
            text += "零"
        pending_zero = False
        text += CN_DIGITS[digit] + CN_UNITS[position]
    return text


def _integer_upper(number: int) -> str:
    if number == 0:
        return "零"
    groups: list[int] = []
    while number:
        groups.append(number % 10000)
        number //= 10000
    text = ""
    for index in range(len(groups) - 1, -1, -1):
        group = groups[index]
        if group == 0:
            if text and not text.endswith("零"):
                text += "零"
            continue
        if text and group < 1000 and not text.endswith("零"):
            text += "零"
        text += _four_digit_upper(group) + CN_GROUP_UNITS[index]
    return text.rstrip("零") or "零"


def rmb_uppercase(value: Any) -> str:
    amount = _decimal(value, "金额")
    if amount < 0:
        raise ResidentialCalculationError("金额不能小于 0")
    cents = int(_rounded(amount * 100, 0))
    yuan, remainder = divmod(cents, 100)
    jiao, fen = divmod(remainder, 10)
    text = "人民币" + _integer_upper(yuan) + "元"
    if not jiao and not fen:
        return text + "整"
    if jiao:
        text += CN_DIGITS[jiao] + "角"
    elif fen:
        text += "零"
    if fen:
        text += CN_DIGITS[fen] + "分"
    return text


def _regional_score(grades: dict[str, Any], label: str) -> Decimal:
    score = Decimal("0")
    for factor in REGIONAL_FACTORS:
        key = factor["key"]
        try:
            grade = int(grades[key])
        except (KeyError, TypeError, ValueError) as exc:
            raise ResidentialCalculationError(f"{label}{factor['label']}的等级无效") from exc
        if grade < 1 or grade > 5:
            raise ResidentialCalculationError(f"{label}{factor['label']}的等级必须是1至5")
        score += Decimal(str(factor["scores"][grade - 1]))
    return score


def _property_factors(item: dict[str, Any], label: str) -> dict[str, Decimal]:
    structure = str(item.get("structure") or "")
    if structure not in STRUCTURE_FACTORS:
        raise ResidentialCalculationError(f"{label}结构类型不在修正规则内")
    orientation = str(item.get("orientation") or "")
    if orientation not in ORIENTATION_FACTORS:
        raise ResidentialCalculationError(f"{label}朝向不在修正规则内")
    floor = int(item.get("floor") or 0)
    total = int(item.get("above_ground_total_floors") or 0)
    building_type = str(item.get("building_type") or "high_rise")
    if building_type == "high_rise":
        try:
            level_factor = high_rise_level_factor(floor, total)
        except ValueError as exc:
            raise ResidentialCalculationError(f"{label}楼层数据无效") from exc
    elif building_type == "multi_story":
        level_factor = MULTI_STORY_LEVELS.get((floor, total))
        if level_factor is None:
            raise ResidentialCalculationError(f"{label}楼层组合不在普通多层修正表内")
    elif building_type == "elevator_multi_story":
        level_factor = ELEVATOR_MULTI_STORY_LEVELS.get((floor, total))
        if level_factor is None:
            raise ResidentialCalculationError(f"{label}楼层组合不在电梯多层修正表内")
    else:
        raise ResidentialCalculationError(f"{label}建筑类型不在修正规则内")
    east_west = str(item.get("east_west") or "")
    if east_west not in EAST_WEST_FACTORS:
        raise ResidentialCalculationError(f"{label}东西至不在修正规则内")
    area_factor = _decimal(item.get("area_factor_pct", 100), f"{label}建筑面积修正")
    build_year = _decimal(item.get("build_year"), f"{label}建成年代")
    subject_year = _decimal(item.get("_subject_build_year"), f"{label}估价对象建成年代")
    return {
        "area": area_factor,
        "structure": Decimal(STRUCTURE_FACTORS[structure]),
        "level": Decimal(str(level_factor)),
        "age": Decimal("100") + (build_year - subject_year) * Decimal("0.5"),
        "orientation": Decimal(ORIENTATION_FACTORS[orientation]),
        "east_west": Decimal("100"),
    }


def _physical_percentage(factors: dict[str, Decimal]) -> Decimal:
    product = Decimal("1")
    for value in factors.values():
        product *= value
    return _rounded(product / Decimal("10000000000"), 2)


def calculate_residential_comparison(project: dict[str, Any]) -> dict[str, Any]:
    subject = dict(project.get("subject") or {})
    comparables = list(project.get("comparables") or [])
    regional = project.get("regional_grades") or {}
    if len(comparables) != 3:
        raise ResidentialCalculationError("必须填写 3 个可比实例")
    case_regional_grades = regional.get("comparables") or []
    if len(case_regional_grades) != 3:
        missing_index = min(len(case_regional_grades), 2)
        case_name = str(
            (comparables[missing_index] if missing_index < len(comparables) else {}).get("name")
            or f"可比实例{chr(65 + missing_index)}"
        )
        raise ResidentialCalculationError(f"{case_name}的区域因素等级缺失或数量不正确")

    subject_area = _decimal(subject.get("area"), "估价对象建筑面积")
    if subject_area <= 0:
        raise ResidentialCalculationError("估价对象建筑面积必须大于 0")
    subject["_subject_build_year"] = subject.get("build_year")
    subject_factors = _property_factors(subject, "估价对象")
    subject_physical_pct = _physical_percentage(subject_factors)
    subject_score = _regional_score(regional.get("subject") or {}, "估价对象")
    certificate_display = str(subject.get("certificate_floor_display") or "")
    above_total = int(subject.get("above_ground_total_floors") or 0)
    subject_floor = int(subject.get("floor") or 0)

    prepared: list[dict[str, Any]] = []
    for index, raw in enumerate(comparables):
        item = dict(raw)
        label = str(item.get("name") or f"可比实例{index + 1}")
        item["_subject_build_year"] = subject.get("build_year")
        factors = _property_factors(item, label)
        physical_pct = _physical_percentage(factors)
        p3 = _rounded(subject_physical_pct / physical_pct, 3)
        case_score = _regional_score(
            case_regional_grades[index],
            label,
        )
        p4 = _rounded(subject_score / case_score, 3)
        p1 = _decimal(item.get("transaction_condition_factor", 1), f"{label}交易情况修正系数")
        p2 = _decimal(item.get("market_condition_factor", 1), f"{label}市场状况修正系数")
        if p1 <= 0:
            raise ResidentialCalculationError(f"{label}交易情况修正系数必须大于 0")
        if p2 <= 0:
            raise ResidentialCalculationError(f"{label}市场状况修正系数必须大于 0")
        combined = _rounded(p1 * p2 * p3 * p4, 3)
        transaction_unit = _decimal(item.get("transaction_unit_price"), f"{label}成交单价")
        if transaction_unit <= 0:
            raise ResidentialCalculationError(f"{label}成交单价必须大于 0")
        comparable_area = _decimal(item.get("area"), f"{label}建筑面积")
        if comparable_area <= 0:
            raise ResidentialCalculationError(f"{label}建筑面积必须大于 0")
        adjusted_amount = _rounded(transaction_unit * combined, 0)
        decoration_delta = _rounded(
            _decimal(item.get("decoration_unit_price", 0), f"{label}装修重置单价")
            * _decimal(item.get("decoration_newness", 0), f"{label}装修成新率")
            - _decimal(subject.get("decoration_unit_price", 0), "估价对象装修重置单价")
            * _decimal(subject.get("decoration_newness", 0), "估价对象装修成新率"),
            0,
        )
        appurtenance_delta = _rounded(
            _decimal(item.get("appurtenance_unit_price", 0), f"{label}附属物扣减"),
            0,
        )
        comparable_unit = adjusted_amount - decoration_delta - appurtenance_delta
        prepared.append(
            {
                "name": label,
                "area": _plain(comparable_area, 2),
                "p1": _plain(p1, 3),
                "p2": _plain(p2, 3),
                "p3": _plain(p3, 3),
                "p4": _plain(p4, 3),
                "combined_factor": _plain(combined, 3),
                "transaction_unit_price": _plain(transaction_unit, 0),
                "adjusted_amount": _plain(adjusted_amount, 0),
                "decoration_deduction": _plain(decoration_delta, 0),
                "appurtenance_deduction": _plain(appurtenance_delta, 0),
                "comparable_unit_price": _plain(comparable_unit, 0),
                "factor_breakdown": {
                    "area": _plain(factors["area"], 2),
                    "structure": _plain(factors["structure"], 2),
                    "level": _plain(factors["level"], 2),
                    "age": _plain(factors["age"], 2),
                    "orientation": _plain(factors["orientation"], 2),
                    "east_west": _plain(factors["east_west"], 2),
                    "physical": _plain(physical_pct, 2),
                },
            }
        )

    estimated_unit = _rounded(
        sum(Decimal(item["comparable_unit_price"]) for item in prepared) / Decimal("3"),
        0,
    )
    estimated_total = _rounded(estimated_unit * subject_area, 0)
    minimum = min(Decimal(item["comparable_unit_price"]) for item in prepared)
    maximum = max(Decimal(item["comparable_unit_price"]) for item in prepared)
    spread = _rounded((maximum / minimum - 1) * 100, 2)
    warnings: list[str] = []
    if spread > 10:
        warnings.append("比较价格相差最好不超过10%")
    for item, calculated in zip(comparables, prepared):
        name = calculated["name"]
        if Decimal(calculated["p1"]) != Decimal("1") and not str(
            item.get("transaction_condition_note") or ""
        ).strip():
            warnings.append(f"{name}交易情况修正未填写依据")
        if Decimal(calculated["p2"]) != Decimal("1") and not str(
            item.get("market_condition_note") or ""
        ).strip():
            warnings.append(f"{name}市场状况修正未填写依据")
    return {
        "subject": {
            "certificate_floor_display": certificate_display,
            "above_ground_total_floors": above_total,
            "floor": subject_floor,
            "floor_correction_display": f"{subject_floor}/{above_total}",
            "level_factor_pct": _plain(subject_factors["level"], 2),
            "physical_factor_pct": _plain(subject_physical_pct, 2),
            "regional_score": _plain(subject_score, 2),
            "factor_breakdown": {
                "area": _plain(subject_factors["area"], 2),
                "structure": _plain(subject_factors["structure"], 2),
                "level": _plain(subject_factors["level"], 2),
                "age": _plain(subject_factors["age"], 2),
                "orientation": _plain(subject_factors["orientation"], 2),
                "east_west": _plain(subject_factors["east_west"], 2),
                "physical": _plain(subject_physical_pct, 2),
            },
        },
        "comparables": prepared,
        "estimated_unit_price": _plain(estimated_unit, 0),
        "estimated_total": _plain(estimated_total, 0),
        "total_uppercase": rmb_uppercase(estimated_total),
        "spread_pct": _plain(spread, 2),
        "warnings": warnings,
    }
