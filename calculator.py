from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from typing import Any


MONEY_UNIT = Decimal("0.01")
ADJUSTMENT_FIELDS = (
    ("time_adjustment_pct", "交易时间修正"),
    ("location_adjustment_pct", "区位状况修正"),
    ("physical_adjustment_pct", "实物状况修正"),
    ("rights_adjustment_pct", "权益状况修正"),
    ("other_adjustment_pct", "其他因素修正"),
)
CN_DIGITS = "零壹贰叁肆伍陆柒捌玖"
CN_UNITS = ("", "拾", "佰", "仟")
CN_GROUP_UNITS = ("", "万", "亿", "万亿")


class CalculationError(ValueError):
    pass


def _decimal(value: Any, field_name: str) -> Decimal:
    try:
        number = Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError) as exc:
        raise CalculationError(f"{field_name}必须是有效数字") from exc
    if not number.is_finite():
        raise CalculationError(f"{field_name}必须是有效数字")
    return number


def _money(value: Decimal) -> str:
    return str(value.quantize(MONEY_UNIT, rounding=ROUND_HALF_UP))


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
        groups.append(number % 10_000)
        number //= 10_000

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
    amount = _decimal(value, "金额").quantize(MONEY_UNIT, rounding=ROUND_HALF_UP)
    if amount == 0:
        return "人民币零元整"
    if amount < 0:
        return "人民币负" + rmb_uppercase(-amount).removeprefix("人民币")

    cents = int(amount * 100)
    yuan, remainder = divmod(cents, 100)
    jiao, fen = divmod(remainder, 10)

    text = "人民币" + _integer_upper(yuan) + "元"
    if jiao == 0 and fen == 0:
        return text + "整"
    if jiao:
        text += CN_DIGITS[jiao] + "角"
    elif fen:
        text += "零"
    if fen:
        text += CN_DIGITS[fen] + "分"
    return text


def calculate_market_comparison(payload: dict[str, Any]) -> dict[str, Any]:
    subject = payload.get("subject") or {}
    comparables = payload.get("comparables") or []

    subject_area = _decimal(subject.get("area"), "估价对象建筑面积")
    if subject_area <= 0:
        raise CalculationError("估价对象建筑面积必须大于 0")
    if not comparables:
        raise CalculationError("至少需要填写一个比较案例")

    total_weight = Decimal("0")
    prepared: list[dict[str, Any]] = []
    for index, item in enumerate(comparables, start=1):
        name = str(item.get("name") or f"比较案例{index}").strip()
        transaction_total = _decimal(item.get("transaction_total"), f"{name}成交总价")
        comparable_area = _decimal(item.get("area"), f"{name}建筑面积")
        weight = _decimal(item.get("weight", 1), f"{name}权重")
        if transaction_total <= 0:
            raise CalculationError(f"{name}成交总价必须大于 0")
        if comparable_area <= 0:
            raise CalculationError(f"{name}建筑面积必须大于 0")
        if weight < 0:
            raise CalculationError(f"{name}权重不能小于 0")

        adjustments: list[dict[str, Any]] = []
        factor = Decimal("1")
        for key, label in ADJUSTMENT_FIELDS:
            percent = _decimal(item.get(key, 0), f"{name}{label}")
            if percent <= -100:
                raise CalculationError(f"{name}{label}不能小于等于 -100%")
            factor *= Decimal("1") + percent / Decimal("100")
            adjustments.append({"key": key, "label": label, "percent": _money(percent)})

        base_unit_price = transaction_total / comparable_area
        adjusted_unit_price = base_unit_price * factor
        total_weight += weight
        prepared.append(
            {
                "name": name,
                "transaction_date": str(item.get("transaction_date") or ""),
                "transaction_total": _money(transaction_total),
                "area": _money(comparable_area),
                "base_unit_price": _money(base_unit_price),
                "adjustment_factor": str(factor.quantize(Decimal("0.000001"))),
                "total_adjustment_pct": _money((factor - 1) * 100),
                "adjusted_unit_price": _money(adjusted_unit_price),
                "adjustments": adjustments,
                "weight": _money(weight),
                "_weight": weight,
                "_adjusted_unit_price": adjusted_unit_price,
            }
        )

    if total_weight <= 0:
        raise CalculationError("至少一个比较案例的权重必须大于 0")

    weighted_unit_price = Decimal("0")
    for item in prepared:
        normalized_weight = item["_weight"] / total_weight
        item["normalized_weight_pct"] = _money(normalized_weight * 100)
        weighted_unit_price += item["_adjusted_unit_price"] * normalized_weight

    adjusted_prices = [item["_adjusted_unit_price"] for item in prepared]
    minimum_price = min(adjusted_prices)
    maximum_price = max(adjusted_prices)
    raw_subject_value = weighted_unit_price * subject_area

    rounding_step = _decimal(payload.get("rounding_step", 1000), "取整单位")
    if rounding_step <= 0:
        raise CalculationError("取整单位必须大于 0")
    subject_value = (
        raw_subject_value / rounding_step
    ).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * rounding_step

    for item in prepared:
        item.pop("_weight", None)
        item.pop("_adjusted_unit_price", None)

    spread_pct = (
        (maximum_price - minimum_price) / weighted_unit_price * 100
        if weighted_unit_price
        else Decimal("0")
    )
    warnings: list[str] = []
    if len(prepared) < 3:
        warnings.append("比较案例少于 3 个，正式报告通常需要补充更多案例。")
    if spread_pct > Decimal("20"):
        warnings.append("比较案例修正后单价离散度超过 20%，建议复核案例或修正幅度。")

    return {
        "comparables": prepared,
        "total_weight": _money(total_weight),
        "weighted_unit_price": _money(weighted_unit_price),
        "raw_subject_value": _money(raw_subject_value),
        "subject_value": _money(subject_value),
        "subject_value_uppercase": rmb_uppercase(subject_value),
        "minimum_adjusted_unit_price": _money(minimum_price),
        "maximum_adjusted_unit_price": _money(maximum_price),
        "spread_pct": _money(spread_pct),
        "rounding_step": _money(rounding_step),
        "warnings": warnings,
    }
