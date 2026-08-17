#!/usr/bin/env python3
"""Exact, registry-driven scalar normalization for the V0.4.1 P2a contract."""

from __future__ import annotations

import re
import unicodedata
from copy import deepcopy
from decimal import Decimal, InvalidOperation
from fractions import Fraction
from typing import Any, Callable


DECIMAL_PATTERN = re.compile(r"^-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?$")


class ValueNormalizationError(ValueError):
    """Raised when a scalar cannot be normalized without losing exact meaning."""


def _decimal_fraction(value: Any) -> Fraction:
    if not isinstance(value, str) or not DECIMAL_PATTERN.fullmatch(value):
        raise ValueNormalizationError("Decimal values must use non-exponent JSON strings.")
    try:
        decimal_value = Decimal(value)
    except InvalidOperation as exc:
        raise ValueNormalizationError(f"Invalid exact decimal value: {value}") from exc
    return Fraction(decimal_value)


def _round_half_even(value: Fraction) -> int:
    sign = -1 if value < 0 else 1
    absolute = abs(value)
    quotient, remainder = divmod(absolute.numerator, absolute.denominator)
    doubled = remainder * 2
    if doubled > absolute.denominator or (
        doubled == absolute.denominator and quotient % 2 == 1
    ):
        quotient += 1
    return sign * quotient


def _format_fixed(integer: int, places: int) -> str:
    if places == 0:
        return str(integer)
    sign = "-" if integer < 0 else ""
    digits = str(abs(integer)).zfill(places + 1)
    result = f"{sign}{digits[:-places]}.{digits[-places:]}"
    return "0." + ("0" * places) if integer == 0 else result


def _format_exact_decimal(value: Fraction) -> str:
    denominator = value.denominator
    twos = 0
    fives = 0
    while denominator % 2 == 0:
        denominator //= 2
        twos += 1
    while denominator % 5 == 0:
        denominator //= 5
        fives += 1
    if denominator != 1:
        raise ValueNormalizationError(
            "Exact decimal output requires comparison_precision for a repeating fraction."
        )
    places = max(twos, fives)
    scaled = value * (10**places)
    if scaled.denominator != 1:
        raise ValueNormalizationError("Exact decimal conversion did not terminate.")
    result = _format_fixed(scaled.numerator, places)
    if "." in result:
        result = result.rstrip("0").rstrip(".")
    return "0" if result in {"-0", ""} else result


def _unit_index(registry: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {item["unit_id"]: item for item in registry["units"]}


def normalize_identity(value: Any, _entry: dict[str, Any], _registry: dict[str, Any]) -> Any:
    if isinstance(value, str):
        return unicodedata.normalize("NFC", value)
    if isinstance(value, float):
        raise ValueNormalizationError("Binary floating-point values are not permitted.")
    return deepcopy(value)


def normalize_decimal(
    value: Any,
    entry: dict[str, Any],
    registry: dict[str, Any],
    *,
    source_unit_id: str | None,
) -> tuple[Any, str | None]:
    if entry["data_type_id"] == "integer":
        if isinstance(value, bool) or not isinstance(value, int):
            raise ValueNormalizationError("Integer normalization requires an integer value.")
        fraction = Fraction(value)
    else:
        fraction = _decimal_fraction(value)
    canonical_unit_id = entry.get("canonical_unit_id")
    if canonical_unit_id is None:
        if source_unit_id is not None:
            raise ValueNormalizationError("Unitless property cannot receive a unit.")
    else:
        if source_unit_id not in entry.get("allowed_unit_ids", []):
            raise ValueNormalizationError(f"Unit is not allowed for {entry['property_id']}.")
        unit = _unit_index(registry).get(source_unit_id)
        if unit is None:
            raise ValueNormalizationError(f"Unregistered unit: {source_unit_id}")
        if unit["canonical_unit_id"] != canonical_unit_id:
            raise ValueNormalizationError("Unit belongs to a different canonical dimension.")
        fraction *= Fraction(
            unit["to_canonical_numerator"], unit["to_canonical_denominator"]
        )

    if entry["data_type_id"] == "integer":
        if fraction.denominator != 1:
            raise ValueNormalizationError(
                "Integer unit conversion must remain an exact integer."
            )
        return fraction.numerator, canonical_unit_id

    precision = entry.get("comparison_precision")
    if precision is not None:
        places = precision["value"]
        rounded = _round_half_even(fraction * (10**places))
        normalized = _format_fixed(rounded, places)
    else:
        normalized = _format_exact_decimal(fraction)
    return normalized, canonical_unit_id


def compare_exact(left: Any, right: Any) -> bool:
    return left == right


def compare_decimal(left: Any, right: Any) -> bool:
    return _decimal_fraction(left) == _decimal_fraction(right)


NORMALIZER_IMPLEMENTATIONS: dict[str, Callable[..., Any]] = {
    "normalizer.identity": normalize_identity,
    "normalizer.decimal": normalize_decimal,
}
COMPARATOR_IMPLEMENTATIONS: dict[str, Callable[..., bool]] = {
    "comparator.exact": compare_exact,
    "comparator.decimal": compare_decimal,
}


def normalize_property_binding(
    binding: dict[str, Any], registry: dict[str, Any]
) -> dict[str, Any]:
    properties = {item["property_id"]: item for item in registry["properties"]}
    entry = properties.get(binding.get("property_id"))
    if entry is None:
        raise ValueNormalizationError(f"Unregistered property: {binding.get('property_id')}")
    typed = binding.get("value", {})
    if typed.get("type") != entry["data_type_id"]:
        raise ValueNormalizationError("Property binding uses the wrong data type.")
    normalizer_id = entry["normalizer_id"]
    implementation = NORMALIZER_IMPLEMENTATIONS.get(normalizer_id)
    if implementation is None:
        raise ValueNormalizationError(f"No built-in normalizer for {normalizer_id}.")

    if normalizer_id == "normalizer.decimal":
        normalized_value, normalized_unit = implementation(
            typed.get("value"),
            entry,
            registry,
            source_unit_id=binding.get("unit_id"),
        )
    else:
        normalized_value = implementation(typed.get("value"), entry, registry)
        normalized_unit = binding.get("unit_id")
        if normalized_unit != entry.get("canonical_unit_id"):
            if not (normalized_unit is None and entry.get("canonical_unit_id") is None):
                raise ValueNormalizationError("Identity-normalized binding is not canonical.")

    result = deepcopy(binding)
    result["value"]["value"] = normalized_value
    result["unit_id"] = normalized_unit
    return result


def compare_property_bindings(
    left: dict[str, Any], right: dict[str, Any], registry: dict[str, Any]
) -> bool:
    normalized_left = normalize_property_binding(left, registry)
    normalized_right = normalize_property_binding(right, registry)
    if normalized_left["property_id"] != normalized_right["property_id"]:
        return False
    entry = {
        item["property_id"]: item for item in registry["properties"]
    }[normalized_left["property_id"]]
    comparator = COMPARATOR_IMPLEMENTATIONS.get(entry["comparator_id"])
    if comparator is None:
        raise ValueNormalizationError(
            f"No built-in comparator for {entry['comparator_id']}."
        )
    return normalized_left["unit_id"] == normalized_right["unit_id"] and comparator(
        normalized_left["value"]["value"], normalized_right["value"]["value"]
    )
