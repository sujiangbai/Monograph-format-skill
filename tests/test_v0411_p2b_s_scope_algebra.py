#!/usr/bin/env python3
"""V0.4.1 P2b-S deterministic scope algebra contract tests."""

from __future__ import annotations

import itertools
import json
import random
import sys
import unicodedata
import unittest
from copy import deepcopy
from functools import lru_cache
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "format-monograph" / "scripts"
sys.path.insert(0, str(SCRIPTS))

import profile_v2_scope as scope_module  # noqa: E402
from profile_v2_scope import (  # noqa: E402
    SELECTOR_KINDS,
    ScopeAlgebraError,
    ScopeContractError,
    derive_scope_id,
    normalize_scope,
    normalized_property_scope_key,
    scope_difference,
    scope_disjoint,
    scope_equal,
    scope_intersection,
    scope_overlap_state,
    scope_partition,
    scope_subset,
)


SEED = 0x04115005
GENERATED_ITERATIONS = 2048
KINDS = (
    "document",
    "section",
    "chapter",
    "semantic_role",
    "object",
    "property",
    "rule",
    "conflict",
)
GROUP_COUNTS = {"NORM": 18, "REL": 24, "DIFF": 30, "PROP": 24}
ASSERTION_IDS = tuple(
    f"T411-S-{group}-{index:03d}"
    for group, count in GROUP_COUNTS.items()
    for index in range(1, count + 1)
)
OTHER = ("__OTHER__",)


def raw_scope(
    selectors: dict[str, list[str]],
    exclusions: dict[str, list[str]] | None = None,
    conditions: list[dict[str, object]] | None = None,
) -> dict[str, object]:
    return {
        "selectors": [
            {"selector_kind": kind, "selector_ids": ids}
            for kind, ids in selectors.items()
        ],
        "exclusions": [
            {"selector_kind": kind, "selector_ids": ids}
            for kind, ids in (exclusions or {}).items()
        ],
        "mutually_exclusive_conditions": conditions or [],
    }


def conditional_scope(document_id: str = "document:one") -> dict[str, object]:
    return raw_scope(
        {"document": [document_id]},
        conditions=[
            {
                "condition_id": "condition:one",
                "condition_kind": "requires_selector",
                "target": {
                    "selector_kind": "chapter",
                    "selector_ids": ["chapter:one"],
                },
            }
        ],
    )


def _oracle_maps(
    scope: dict[str, object], field: str
) -> dict[str, set[str]]:
    result: dict[str, set[str]] = {}
    for selector in scope[field]:  # type: ignore[index]
        kind = selector["selector_kind"]
        result[kind] = {
            unicodedata.normalize("NFC", value) for value in selector["selector_ids"]
        }
    return result


def _oracle_universe(
    scopes: list[dict[str, object]],
) -> tuple[list[str], dict[str, tuple[object, ...]]]:
    kinds: set[str] = set()
    ids: dict[str, set[str]] = {}
    for scope in scopes:
        for field in ("selectors", "exclusions"):
            for selector in scope[field]:  # type: ignore[index]
                kind = selector["selector_kind"]
                kinds.add(kind)
                ids.setdefault(kind, set()).update(
                    unicodedata.normalize("NFC", value)
                    for value in selector["selector_ids"]
                )
    ordered_kinds = sorted(kinds)
    atoms = {
        kind: tuple(sorted(ids.get(kind, set()))) + (OTHER,)
        for kind in ordered_kinds
    }
    return ordered_kinds, atoms


def _oracle_members(
    scope: dict[str, object],
    kinds: list[str],
    atoms: dict[str, tuple[object, ...]],
) -> set[tuple[object, ...]]:
    selectors = _oracle_maps(scope, "selectors")
    exclusions = _oracle_maps(scope, "exclusions")
    members: set[tuple[object, ...]] = set()
    for point in itertools.product(*(atoms[kind] for kind in kinds)):
        accepted = True
        for kind, atom in zip(kinds, point):
            if kind in selectors and atom not in selectors[kind]:
                accepted = False
            if kind in exclusions and atom in exclusions[kind]:
                accepted = False
        if accepted:
            members.add(point)
    return members


def _oracle_partition_sets(
    source: dict[str, object],
    cut: dict[str, object],
    result: dict[str, object],
) -> tuple[set[tuple[object, ...]], set[tuple[object, ...]], list[set[tuple[object, ...]]]]:
    component_scopes = [source, cut, *result["residual_scopes"]]  # type: ignore[list-item]
    kinds, atoms = _oracle_universe(component_scopes)
    source_members = _oracle_members(source, kinds, atoms)
    cut_members = _oracle_members(cut, kinds, atoms)
    residual_members = [
        _oracle_members(residual, kinds, atoms)
        for residual in result["residual_scopes"]  # type: ignore[index]
    ]
    return source_members, cut_members, residual_members


def _generated_scope_pairs() -> tuple[tuple[dict[str, object], dict[str, object]], ...]:
    rng = random.Random(SEED)
    pairs = []
    for iteration in range(GENERATED_ITERATIONS):
        kind_count = rng.randint(1, 4)
        active = rng.sample(list(KINDS), kind_count)
        source_selectors: dict[str, list[str]] = {}
        for kind in active:
            size = rng.randint(1, 3)
            values = [f"{kind}:{iteration}:{index}" for index in range(size)]
            rng.shuffle(values)
            source_selectors[kind] = values
        cut_selectors = deepcopy(source_selectors)
        restrictable = [kind for kind in active if len(cut_selectors[kind]) > 1]
        remaining = [kind for kind in KINDS if kind not in active]
        if restrictable and (not remaining or rng.choice((True, False))):
            kind = rng.choice(restrictable)
            cut_selectors[kind] = [rng.choice(cut_selectors[kind])]
        else:
            kind = rng.choice(remaining)
            cut_selectors[kind] = [f"{kind}:{iteration}:cut"]
        source_items = list(source_selectors.items())
        cut_items = list(cut_selectors.items())
        rng.shuffle(source_items)
        rng.shuffle(cut_items)
        pairs.append((raw_scope(dict(source_items)), raw_scope(dict(cut_items))))
    return tuple(pairs)


@lru_cache(maxsize=1)
def _generated_evidence() -> tuple[
    tuple[
        dict[str, object],
        dict[str, object],
        dict[str, object],
        set[tuple[object, ...]],
        set[tuple[object, ...]],
        list[set[tuple[object, ...]]],
    ],
    ...,
]:
    evidence = []
    for iteration, (source, cut) in enumerate(_generated_scope_pairs()):
        result = scope_partition(source, cut)
        source_members, cut_members, residual_members = _oracle_partition_sets(
            source, cut, result
        )
        valid = (
            result["status"] == "partitioned"
            and cut_members < source_members
            and all(part <= source_members for part in residual_members)
            and all(part.isdisjoint(cut_members) for part in residual_members)
            and all(
                left.isdisjoint(right)
                for offset, left in enumerate(residual_members)
                for right in residual_members[offset + 1 :]
            )
            and set().union(cut_members, *residual_members) == source_members
            and [item["scope_id"] for item in result["residual_scopes"]]
            == sorted(item["scope_id"] for item in result["residual_scopes"])
        )
        if not valid:
            minimal = _shrink_counterexample(
                source,
                cut,
                lambda candidate_source, candidate_cut: not _generated_case_holds(
                    candidate_source, candidate_cut
                ),
            )
            raise AssertionError(
                "scope algebra generated counterexample: "
                + json.dumps(
                    {
                        "seed": SEED,
                        "iteration": iteration,
                        "minimal": minimal,
                    },
                    ensure_ascii=False,
                    sort_keys=True,
                    separators=(",", ":"),
                )
            )
        evidence.append(
            (
                source,
                cut,
                result,
                source_members,
                cut_members,
                residual_members,
            )
        )
    return tuple(evidence)


def _generated_case_holds(
    source: dict[str, object], cut: dict[str, object]
) -> bool:
    try:
        result = scope_partition(source, cut)
        if result["status"] != "partitioned":
            return False
        source_members, cut_members, residual_members = _oracle_partition_sets(
            source, cut, result
        )
    except (ScopeContractError, KeyError, TypeError, ValueError):
        return False
    return (
        cut_members < source_members
        and all(part <= source_members for part in residual_members)
        and all(part.isdisjoint(cut_members) for part in residual_members)
        and all(
            left.isdisjoint(right)
            for offset, left in enumerate(residual_members)
            for right in residual_members[offset + 1 :]
        )
        and set().union(cut_members, *residual_members) == source_members
    )


def _candidate_reductions(
    source: dict[str, object], cut: dict[str, object]
) -> list[tuple[dict[str, object], dict[str, object]]]:
    reductions = []
    targets = (("source", source, cut), ("cut", cut, source))
    for target_name, original, other in targets:
        values = original["selectors"]  # type: ignore[index]
        if len(values) > 1:
            candidate = deepcopy(original)
            candidate["selectors"] = values[:-1]
            reductions.append(
                (candidate, deepcopy(other))
                if target_name == "source"
                else (deepcopy(other), candidate)
            )
    for target_name, original, other in targets:
        for selector_index, selector in enumerate(original["selectors"]):  # type: ignore[index]
            if len(selector["selector_ids"]) > 1:
                candidate = deepcopy(original)
                candidate["selectors"][selector_index]["selector_ids"] = selector[
                    "selector_ids"
                ][:-1]
                reductions.append(
                    (candidate, deepcopy(other))
                    if target_name == "source"
                    else (deepcopy(other), candidate)
                )
    for field in ("exclusions", "mutually_exclusive_conditions"):
        for target_name, original, other in targets:
            values = original[field]  # type: ignore[index]
            if values:
                candidate = deepcopy(original)
                candidate[field] = values[:-1]
                reductions.append(
                    (candidate, deepcopy(other))
                    if target_name == "source"
                    else (deepcopy(other), candidate)
                )
    return reductions


def _shrink_counterexample(
    source: dict[str, object],
    cut: dict[str, object],
    still_fails: object,
) -> tuple[dict[str, object], dict[str, object]]:
    current = (deepcopy(source), deepcopy(cut))
    while True:
        changed = False
        for candidate in _candidate_reductions(*current):
            try:
                if still_fails(*candidate):  # type: ignore[operator]
                    current = candidate
                    changed = True
                    break
            except (KeyError, TypeError, ValueError):
                continue
        if not changed:
            return current


class ScopeAlgebraContracts(unittest.TestCase):
    maxDiff = None

    def test_t411_s_assertion_manifest_is_exact(self) -> None:
        self.assertEqual(96, len(ASSERTION_IDS))
        self.assertEqual(96, len(set(ASSERTION_IDS)))
        self.assertEqual(set(KINDS), SELECTOR_KINDS)
        for group, count in GROUP_COUNTS.items():
            expected = {
                f"T411-S-{group}-{index:03d}" for index in range(1, count + 1)
            }
            self.assertEqual(
                expected, {item for item in ASSERTION_IDS if f"-{group}-" in item}
            )

    def test_v0411_s_normalize_001(self) -> None:
        for index in range(1, GROUP_COUNTS["NORM"] + 1):
            assertion_id = f"T411-S-NORM-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                self._assert_normalize(index)

    def _assert_normalize(self, index: int) -> None:
        if index <= 8:
            kind = KINDS[index - 1]
            normalized = normalize_scope(raw_scope({kind: [f"{kind}:one"]}))
            self.assertEqual(kind, normalized["selectors"][0]["selector_kind"])
            self.assertEqual(derive_scope_id(normalized), normalized["scope_id"])
            return
        if index == 9:
            composed = normalize_scope(raw_scope({"document": ["document:Caf\u00e9"]}))
            decomposed = normalize_scope(raw_scope({"document": ["document:Cafe\u0301"]}))
            self.assertEqual(composed, decomposed)
        elif index == 10:
            left = raw_scope({"document": ["document:one"], "chapter": ["chapter:one"]})
            right = raw_scope({"chapter": ["chapter:one"], "document": ["document:one"]})
            self.assertEqual(normalize_scope(left), normalize_scope(right))
        elif index == 11:
            left = raw_scope({"document": ["document:two", "document:one"]})
            right = raw_scope({"document": ["document:one", "document:two"]})
            self.assertEqual(normalize_scope(left), normalize_scope(right))
        elif index == 12:
            value = normalize_scope(
                raw_scope(
                    {"object": ["object:one"]},
                    {"object": ["object:two"]},
                )
            )
            self.assertEqual("object:two", value["exclusions"][0]["selector_ids"][0])
        elif index == 13:
            value = normalize_scope(
                raw_scope(
                    {"object": ["object:one", "object:two"]},
                    {"object": ["object:two"]},
                )
            )
            self.assertEqual(2, len(value["selectors"][0]["selector_ids"]))
        elif index == 14:
            with self.assertRaises(ScopeContractError):
                normalize_scope(raw_scope({"document": ["document:one", "document:one"]}))
        elif index == 15:
            value = raw_scope({"document": ["document:one"]})
            value["selectors"].append(  # type: ignore[index]
                {"selector_kind": "document", "selector_ids": ["document:two"]}
            )
            with self.assertRaises(ScopeContractError):
                normalize_scope(value)
        elif index == 16:
            with self.assertRaises(ScopeContractError):
                normalize_scope(raw_scope({}))
        elif index == 17:
            value = raw_scope({"document": ["document:one"]})
            value["scope_id"] = "scope:" + ("0" * 64)
            with self.assertRaises(ScopeContractError):
                scope_partition(value, raw_scope({"document": ["document:one"]}))
        else:
            value = normalize_scope(raw_scope({"document": ["document:test"]}))
            self.assertEqual(
                "scope:d06ce74ec1d6e69d08a70667c7314dec9b1debe6d7696610d98bec36e4ca8ef9",
                value["scope_id"],
            )
            self.assertEqual(
                ("paragraph", "font.size", value["scope_id"]),
                normalized_property_scope_key("paragraph", "font.size", value),
            )

    def test_v0411_s_relation_002(self) -> None:
        for index in range(1, GROUP_COUNTS["REL"] + 1):
            assertion_id = f"T411-S-REL-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                self._assert_relation(index)

    def _assert_relation(self, index: int) -> None:
        if index <= 8:
            scope = raw_scope({KINDS[index - 1]: [f"{KINDS[index - 1]}:one"]})
            self.assertTrue(scope_equal(scope, deepcopy(scope)))
            self.assertTrue(scope_subset(scope, deepcopy(scope)))
            return
        conjunctions = (
            ("document", "chapter"),
            ("section", "semantic_role"),
            ("object", "property"),
            ("rule", "conflict"),
            ("document", "object", "property"),
            ("chapter", "semantic_role", "rule"),
        )
        if 9 <= index <= 14:
            kinds = conjunctions[index - 9]
            source = raw_scope({kinds[0]: [f"{kinds[0]}:one"]})
            cut = raw_scope({kind: [f"{kind}:one"] for kind in kinds})
            self.assertTrue(scope_subset(cut, source))
            self.assertEqual("overlap", scope_overlap_state(source, cut))
        elif index == 15:
            source = raw_scope({"document": ["document:one"]})
            cut = raw_scope(
                {"document": ["document:one"], "chapter": ["chapter:one"]}
            )
            self.assertFalse(scope_subset(source, cut))
        elif index == 16:
            left = raw_scope({"document": ["document:one"]})
            right = raw_scope({"document": ["document:two"]})
            self.assertTrue(scope_disjoint(left, right))
        elif index == 17:
            source = raw_scope({"document": ["document:one"]})
            cut = raw_scope(
                {"document": ["document:one"], "chapter": ["chapter:one"]}
            )
            self.assertEqual("overlap", scope_overlap_state(source, cut))
        elif index == 18:
            left = raw_scope({"document": ["document:one", "document:two"]})
            right = raw_scope({"document": ["document:two", "document:three"]})
            self.assertEqual("overlap", scope_overlap_state(left, right))
        elif index == 19:
            value = conditional_scope()
            self.assertTrue(scope_equal(value, deepcopy(value)))
            self.assertEqual(normalize_scope(value), scope_intersection(value, value))
        elif index == 20:
            self.assertEqual(
                "unknown",
                scope_overlap_state(
                    conditional_scope(), raw_scope({"document": ["document:one"]})
                ),
            )
        elif index == 21:
            self.assertEqual(
                "disjoint",
                scope_overlap_state(
                    conditional_scope(), raw_scope({"document": ["document:two"]})
                ),
            )
        elif index == 22:
            partial = raw_scope(
                {
                    "document": ["document:one", "document:two"],
                    "chapter": ["chapter:one"],
                },
                {"document": ["document:two"]},
            )
            other = raw_scope({"document": ["document:one"]})
            self.assertEqual("unknown", scope_overlap_state(partial, other))
            other_axis = raw_scope(
                {
                    "document": ["document:one"],
                    "chapter": ["chapter:other"],
                }
            )
            self.assertEqual("disjoint", scope_overlap_state(partial, other_axis))
        elif index == 23:
            source = raw_scope({"document": ["document:one"]})
            cut = raw_scope(
                {"document": ["document:one"], "chapter": ["chapter:one"]}
            )
            self.assertEqual(normalize_scope(cut), scope_intersection(source, cut))
        else:
            source = raw_scope(
                {"document": ["document:one"], "chapter": ["chapter:one"]}
            )
            cut = raw_scope({"document": ["document:one"]})
            self.assertEqual(normalize_scope(source), scope_intersection(source, cut))

    def test_v0411_s_difference_003(self) -> None:
        for index in range(1, GROUP_COUNTS["DIFF"] + 1):
            assertion_id = f"T411-S-DIFF-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                self._assert_difference(index)

    def _assert_difference(self, index: int) -> None:
        source = raw_scope({"document": ["document:one"]})
        narrower = raw_scope(
            {"document": ["document:one"], "chapter": ["chapter:one"]}
        )
        disjoint = raw_scope({"document": ["document:other"]})
        crossing_source = raw_scope(
            {"document": ["document:one", "document:two"]}
        )
        crossing_cut = raw_scope(
            {"document": ["document:two", "document:three"]}
        )
        partial = raw_scope(
            {"document": ["document:one", "document:two"]},
            {"document": ["document:two"]},
        )
        if index == 1:
            self.assertEqual([], scope_difference(source, deepcopy(source)))
        elif index == 2:
            self.assertEqual([normalize_scope(source)], scope_difference(source, disjoint))
        elif index == 3:
            finite_source = raw_scope(
                {"document": ["document:one", "document:two", "document:three"]}
            )
            finite_cut = raw_scope({"document": ["document:two"]})
            residual = scope_difference(finite_source, finite_cut)
            self.assertEqual(
                ["document:one", "document:three"],
                residual[0]["selectors"][0]["selector_ids"],
            )
        elif index == 4:
            residual = scope_difference(source, narrower)
            self.assertEqual("chapter", residual[0]["exclusions"][0]["selector_kind"])
        elif index == 5:
            cut = raw_scope(
                {
                    "document": ["document:one"],
                    "chapter": ["chapter:one"],
                    "object": ["object:one"],
                }
            )
            self.assertEqual(2, len(scope_difference(source, cut)))
        elif index == 6:
            with self.assertRaises(ScopeAlgebraError) as raised:
                scope_difference(narrower, source)
            self.assertEqual("unexpressible_difference", raised.exception.code)
            self.assertEqual("strict_superset", raised.exception.details["relation"])
        elif index == 7:
            with self.assertRaises(ScopeAlgebraError) as raised:
                scope_difference(crossing_source, crossing_cut)
            self.assertEqual("crossing_overlap", raised.exception.code)
        elif index == 8:
            with self.assertRaises(ScopeAlgebraError) as raised:
                scope_difference(source, conditional_scope())
            self.assertEqual("unprovable_condition", raised.exception.code)
        elif index == 9:
            with self.assertRaises(ScopeAlgebraError) as raised:
                scope_difference(partial, source)
            self.assertEqual("unexpressible_difference", raised.exception.code)
        elif index == 10:
            self.assertEqual("equal", scope_partition(source, deepcopy(source))["status"])
        elif index == 11:
            self.assertEqual("disjoint", scope_partition(source, disjoint)["status"])
        elif index == 12:
            result = scope_partition(source, narrower)
            self.assertEqual("partitioned", result["status"])
            self.assertEqual(normalize_scope(narrower), result["intersection"])
        elif index == 13:
            result = scope_partition(narrower, source)
            self.assertEqual(("blocked", "unexpressible_difference"), (result["status"], result["code"]))
            self.assertEqual("strict_superset", result["evidence"]["relation"])
        elif index == 14:
            result = scope_partition(crossing_source, crossing_cut)
            self.assertEqual(("blocked", "crossing_overlap"), (result["status"], result["code"]))
        elif index == 15:
            result = scope_partition(source, conditional_scope())
            self.assertEqual("unprovable_condition", result["code"])
        elif index == 16:
            result = scope_partition(partial, source)
            self.assertEqual("unexpressible_difference", result["code"])
        elif index == 17:
            with mock.patch.object(scope_module, "_prove_partition_conservation", return_value=False):
                with self.assertRaises(ScopeAlgebraError) as raised:
                    scope_difference(source, narrower)
            self.assertEqual("partition_conservation_failure", raised.exception.code)
        elif index == 18:
            with mock.patch.object(scope_module, "_prove_partition_conservation", return_value=False):
                result = scope_partition(source, narrower)
            self.assertEqual(("blocked", "partition_conservation_failure"), (result["status"], result["code"]))
        elif index == 19:
            cut = raw_scope(
                {
                    "document": ["document:one"],
                    "chapter": ["chapter:one"],
                    "object": ["object:one"],
                }
            )
            ids = [item["scope_id"] for item in scope_difference(source, cut)]
            self.assertEqual(sorted(ids), ids)
        elif index == 20:
            permuted = deepcopy(narrower)
            permuted["selectors"].reverse()
            self.assertEqual(scope_difference(source, narrower), scope_difference(source, permuted))
        elif index == 21:
            composed = raw_scope({"document": ["document:Caf\u00e9"]})
            decomposed = raw_scope({"document": ["document:Cafe\u0301"], "chapter": ["chapter:one"]})
            expected = scope_difference(composed, raw_scope({"document": ["document:Caf\u00e9"], "chapter": ["chapter:one"]}))
            self.assertEqual(expected, scope_difference(composed, decomposed))
        elif index == 22:
            residuals = scope_difference(source, narrower)
            self.assertTrue(all(item["selectors"] for item in residuals))
        elif index == 23:
            cut = raw_scope(
                {"document": ["document:one"], "chapter": ["chapter:one"], "object": ["object:one"]}
            )
            ids = [item["scope_id"] for item in scope_difference(source, cut)]
            self.assertEqual(len(ids), len(set(ids)))
        elif index == 24:
            wide = raw_scope({"document": ["document:one"]}, {"chapter": ["chapter:zero"]})
            cut = raw_scope({"document": ["document:one"], "chapter": ["chapter:one"]})
            residual = scope_difference(wide, cut)[0]
            self.assertEqual(["chapter:one", "chapter:zero"], residual["exclusions"][0]["selector_ids"])
        elif index == 25:
            wide = raw_scope({"document": ["document:one"]}, {"chapter": ["chapter:zero"]})
            cut = raw_scope({"document": ["document:one"]}, {"chapter": ["chapter:zero", "chapter:one"]})
            residual = scope_difference(wide, cut)[0]
            self.assertEqual(["chapter:one"], residual["selectors"][0]["selector_ids"])
        elif index == 26:
            wide = raw_scope({"document": ["document:one", "document:two"]})
            cut = raw_scope({"document": ["document:one"]})
            self.assertEqual(["document:two"], scope_difference(wide, cut)[0]["selectors"][0]["selector_ids"])
        elif index == 27:
            malformed = deepcopy(source)
            malformed["scope_id"] = "scope:" + ("0" * 64)
            with self.assertRaises(ScopeContractError):
                scope_partition(malformed, narrower)
        elif index == 28:
            result = scope_partition(source, narrower)
            self.assertEqual(result, json.loads(json.dumps(result, ensure_ascii=False)))
            self.assertEqual(
                {
                    "status",
                    "code",
                    "source_scope_id",
                    "cut_scope_id",
                    "intersection",
                    "residual_scopes",
                    "evidence",
                },
                set(result),
            )
        elif index == 29:
            self.assertIsNone(scope_intersection(source, disjoint))
        else:
            error = ScopeAlgebraError("unknown_overlap", {"relation": "unknown"})
            with mock.patch.object(scope_module, "_scope_relation_normalized", side_effect=error):
                result = scope_partition(source, narrower)
            self.assertEqual(("blocked", "unknown_overlap"), (result["status"], result["code"]))

    def test_v0411_s_conservation_004(self) -> None:
        for index in range(1, GROUP_COUNTS["PROP"] + 1):
            assertion_id = f"T411-S-PROP-{index:03d}"
            with self.subTest(assertion_id=assertion_id):
                self._assert_property(index)

    def _assert_property(self, index: int) -> None:
        evidence = _generated_evidence()
        if index == 1:
            self.assertEqual(GENERATED_ITERATIONS, len(evidence))
            self.assertTrue(all(item[2]["status"] == "partitioned" for item in evidence))
        elif index == 2:
            self.assertTrue(all(item[4] < item[3] for item in evidence))
        elif index == 3:
            self.assertTrue(all(all(part <= item[3] for part in item[5]) for item in evidence))
        elif index == 4:
            self.assertTrue(
                all(
                    all(left.isdisjoint(right) for offset, left in enumerate(item[5]) for right in item[5][offset + 1 :])
                    for item in evidence
                )
            )
        elif index == 5:
            self.assertTrue(all(all(part.isdisjoint(item[4]) for part in item[5]) for item in evidence))
        elif index == 6:
            self.assertTrue(all(set().union(item[4], *item[5]) == item[3] for item in evidence))
        elif index == 7:
            for source, cut, result, *_ in evidence[:128]:
                source = deepcopy(source)
                cut = deepcopy(cut)
                source["selectors"].reverse()
                cut["selectors"].reverse()
                self.assertEqual(result, scope_partition(source, cut))
        elif index == 8:
            for source, cut, result, *_ in evidence[:128]:
                source = dict(reversed(list(source.items())))
                cut = dict(reversed(list(cut.items())))
                self.assertEqual(result, scope_partition(source, cut))
        elif index == 9:
            left = raw_scope({"document": ["document:Caf\u00e9"]})
            right = raw_scope({"document": ["document:Cafe\u0301"], "chapter": ["chapter:one"]})
            self.assertEqual("partitioned", scope_partition(left, right)["status"])
        elif index == 10:
            for kind in KINDS:
                source = raw_scope({kind: [f"{kind}:one", f"{kind}:two"]})
                cut = raw_scope({kind: [f"{kind}:one"]})
                self.assertEqual("partitioned", scope_partition(source, cut)["status"])
        elif index == 11:
            conjunctions = (
                ("document", "chapter"),
                ("section", "semantic_role"),
                ("object", "property"),
                ("rule", "conflict"),
                ("document", "object", "property"),
                ("chapter", "semantic_role", "rule"),
            )
            for kinds in conjunctions:
                source = raw_scope({kinds[0]: [f"{kinds[0]}:one"]})
                cut = raw_scope({kind: [f"{kind}:one"] for kind in kinds})
                self.assertEqual("partitioned", scope_partition(source, cut)["status"])
        elif index == 12:
            for source, cut, *_ in evidence[:128]:
                result = scope_partition(cut, source)
                self.assertEqual(("blocked", "unexpressible_difference"), (result["status"], result["code"]))
                self.assertEqual("strict_superset", result["evidence"]["relation"])
        elif index == 13:
            result = scope_partition(
                raw_scope({"document": ["document:one", "document:two"]}),
                raw_scope({"document": ["document:two", "document:three"]}),
            )
            self.assertEqual("crossing_overlap", result["code"])
        elif index == 14:
            result = scope_partition(raw_scope({"document": ["document:one"]}), conditional_scope())
            self.assertEqual("unprovable_condition", result["code"])
        elif index == 15:
            result = scope_partition(
                raw_scope({"document": ["document:one", "document:two"]}, {"document": ["document:two"]}),
                raw_scope({"document": ["document:one"]}),
            )
            self.assertEqual("unexpressible_difference", result["code"])
        elif index == 16:
            self.assertEqual(_generated_scope_pairs(), _generated_scope_pairs())
        elif index == 17:
            source = raw_scope(
                {"document": ["document:one", "document:two"], "chapter": ["chapter:one"]},
                {"object": ["object:excluded"]},
                conditional_scope()["mutually_exclusive_conditions"],
            )
            cut = raw_scope({"document": ["document:one"], "chapter": ["chapter:one"]})
            minimal = _shrink_counterexample(source, cut, lambda *_: True)
            self.assertLess(len(json.dumps(minimal, sort_keys=True)), len(json.dumps((source, cut), sort_keys=True)))
        elif index == 18:
            names = _oracle_members.__code__.co_names
            self.assertFalse(any(name.startswith("scope_") or name.startswith("_proof") for name in names))
        elif index == 19:
            source, cut = _generated_scope_pairs()[0]
            first = json.dumps(scope_partition(source, cut), sort_keys=True, ensure_ascii=False)
            second = json.dumps(scope_partition(source, cut), sort_keys=True, ensure_ascii=False)
            self.assertEqual(first, second)
        elif index == 20:
            source = raw_scope({"document": ["document:one"]})
            cut = raw_scope({"document": ["document:one"], "chapter": ["chapter:one"]})
            result = scope_partition(source, cut)
            source_members, _, residual_members = _oracle_partition_sets(source, cut, result)
            self.assertTrue(source_members)
            self.assertTrue(set().union(*residual_members) < source_members)
        elif index == 21:
            self.assertTrue(
                all(
                    [item["scope_id"] for item in result[2]["residual_scopes"]]
                    == sorted(item["scope_id"] for item in result[2]["residual_scopes"])
                    for result in evidence
                )
            )
        elif index == 22:
            blocked_cases = (
                scope_partition(
                    raw_scope({"document": ["document:one", "document:two"]}),
                    raw_scope({"document": ["document:two", "document:three"]}),
                ),
                scope_partition(raw_scope({"document": ["document:one"]}), conditional_scope()),
            )
            self.assertTrue(all(not item["residual_scopes"] and item["intersection"] is None for item in blocked_cases))
        elif index == 23:
            source, cut = _generated_scope_pairs()[1]
            before = deepcopy((source, cut))
            scope_partition(source, cut)
            self.assertEqual(before, (source, cut))
        else:
            result = scope_partition(*_generated_scope_pairs()[2])
            encoded = json.dumps(result, sort_keys=True)
            self.assertNotIn("message", encoded)
            self.assertNotIn("platform", encoded)
            self.assertNotIn("path", encoded)
            self.assertNotIn("time", encoded)


if __name__ == "__main__":
    unittest.main()
