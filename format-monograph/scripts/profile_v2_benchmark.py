#!/usr/bin/env python3
"""Pure P3a-C2 benchmark contract helpers. No runner or runtime integration."""
from __future__ import annotations
import copy, hashlib, json, math, re
from decimal import Decimal, ROUND_HALF_EVEN
from pathlib import PurePosixPath
from statistics import median
from typing import Any, Iterable, Mapping

class BenchmarkContractError(ValueError): pass

SCALES=("0.5x","1.0x","1.5x","2.0x")
SCENARIOS=("disjoint","subset-chain","dense-crossing","mixed-conflict-approval")
DETERMINISM_SCALES=("1.5x","2.0x")
PERFORMANCE_CELLS=(("0.5x","mixed-conflict-approval"),("1.0x","mixed-conflict-approval"),*((s,c) for s in DETERMINISM_SCALES for c in SCENARIOS))
TIMING_STAGES=("synthetic_generation","schema_registry_validation","compose","approval_generation","apply","canonical_serialization","end_to_end")
C1_METRIC_FIELDS=("input_asset_count","input_rule_count","input_binding_count","expected_key_count","candidate_count","candidate_group_count","partition_count","conflict_count","proposal_count","blocker_count","max_candidates_per_key","max_partition_width","max_repartition_depth")
STOP_REASON_ORDER=("threshold_exceeded","wall_ratio_exceeded","rss_ratio_exceeded","output_ratio_exceeded","timeout","process_crash","rss_unavailable","coverage_conservation_failure","stable_id_drift","canonical_nondeterminism","fingerprint_nondeterminism","terminal_state_mismatch","subject_stale","contract_error","reference_budget_exceeded")
STOP_REASON_INDEX={v:i for i,v in enumerate(STOP_REASON_ORDER)}
SHA256_RE=re.compile(r"^sha256:[a-f0-9]{64}$"); COMMIT_RE=re.compile(r"^[a-f0-9]{40}$")
FROZEN_RSS_PROTOCOL={"measurement_method":"external_supervisor_child_peak_working_set","sampling_interval_seconds":0.05,"child_process_scope":"benchmark_child_process_only","record_baseline":True,"record_delta":True,"unavailable_policy":"fail_closed","delta_rounding":"mib_3_decimal_places_half_even"}
FROZEN_THRESHOLDS={"scale_limits":[{"scale_id":"0.5x","max_wall_seconds":60,"max_peak_rss_mib":512},{"scale_id":"1.0x","max_wall_seconds":60,"max_peak_rss_mib":512},{"scale_id":"1.5x","max_wall_seconds":60,"max_peak_rss_mib":512},{"scale_id":"2.0x","max_wall_seconds":120,"max_peak_rss_mib":1024}],"wall_median_ratio_2x_to_1x":6,"rss_ratio_2x_to_1x":3,"output_json_ratio_2x_to_1x":3}

def _finite(v):
    if isinstance(v,float) and not math.isfinite(v): raise BenchmarkContractError("Canonical JSON forbids NaN/Infinity")
    if isinstance(v,Mapping):
        for x in v.values(): _finite(x)
    elif isinstance(v,(list,tuple)):
        for x in v: _finite(x)

def canonical_json_bytes(v):
    _finite(v)
    try: return json.dumps(v,ensure_ascii=False,sort_keys=True,separators=(",",":"),allow_nan=False).encode()
    except (TypeError,ValueError) as e: raise BenchmarkContractError(f"not canonical-JSON safe: {e}") from e

def canonical_sha256(v): return "sha256:"+hashlib.sha256(canonical_json_bytes(v)).hexdigest()
def _self_digest(d,key):
    x=copy.deepcopy(dict(d)); x.pop(key,None); return canonical_sha256(x)
def recompute_config_digest(d): return _self_digest(d,"config_digest")
def recompute_envelope_digest(d): return _self_digest(d,"envelope_digest")
def recompute_result_digest(d): return _self_digest(d,"result_digest")

def recompute_subject_digest(manifest):
    paths=[]; files=[]
    for e in manifest:
        p=e.get("path"); h=e.get("sha256")
        if not isinstance(p,str) or not p or "\\" in p or PurePosixPath(p).is_absolute() or ".." in PurePosixPath(p).parts or re.match(r"^[A-Za-z]:",p): raise BenchmarkContractError("invalid subject path")
        if not isinstance(h,str) or not SHA256_RE.fullmatch(h): raise BenchmarkContractError("invalid subject digest")
        paths.append(p); files.append({"path":p,"sha256":h})
    if len(paths)!=len(set(paths)) or paths!=sorted(paths): raise BenchmarkContractError("subject paths must be unique and sorted")
    return canonical_sha256({"basis":"canonical_subject_manifest_v1","files":files})

def expected_coverage_cells(): return {(s,c) for s in SCALES for c in SCENARIOS}
def expected_performance_cells(): return set(PERFORMANCE_CELLS)
def expected_determinism_cells(seeds): return {(s,c,int(seed)) for s in DETERMINISM_SCALES for c in SCENARIOS for seed in seeds}
def _exact(actual,expected,label):
    if len(actual)!=len(set(actual)) or set(actual)!=expected: raise BenchmarkContractError(f"{label} differs from frozen matrix")

def validate_benchmark_config_semantics(d):
    seeds=list(d["generation"]["permutation_seeds"])
    if len(seeds)<5 or len(seeds)!=len(set(seeds)): raise BenchmarkContractError("five unique permutation seeds required")
    if tuple(x["scale_id"] for x in d["scales"])!=SCALES or tuple(x["scenario_id"] for x in d["scenarios"])!=SCENARIOS: raise BenchmarkContractError("frozen scale/scenario set mismatch")
    if any(tuple(x["scenario_ids"])!=SCENARIOS for x in d["scales"]): raise BenchmarkContractError("coverage scenario set mismatch")
    m=d["matrices"]
    _exact([(x["scale_id"],x["scenario_id"]) for x in m["coverage_cells"]],expected_coverage_cells(),"coverage")
    _exact([(x["scale_id"],x["scenario_id"]) for x in m["performance_cells"]],expected_performance_cells(),"performance")
    _exact([(x["scale_id"],x["scenario_id"],x["permutation_seed"]) for x in m["determinism_cells"]],expected_determinism_cells(seeds),"determinism")
    if d["repetitions"]!={"coverage_min_runs_per_cell":1,"performance_warmup_runs_per_cell":1,"performance_measured_runs_per_cell":3,"determinism_runs_per_seed":1}: raise BenchmarkContractError("repetition policy mismatch")
    if d["rss_protocol"]!=FROZEN_RSS_PROTOCOL or d["thresholds"]!=FROZEN_THRESHOLDS or d["total_reference_budget_hours"]!=3: raise BenchmarkContractError("frozen protocol/threshold mismatch")
    if d["config_digest"]!=recompute_config_digest(d): raise BenchmarkContractError("config digest mismatch")

def _zero(c): return all(v==0 for v in c.values())
def validate_projected_envelope_semantics(d):
    entries=list(d["entries"]); ids=[x["decision_id"] for x in entries]
    if len(entries)!=171 or len(ids)!=len(set(ids)): raise BenchmarkContractError("envelope must have 171 unique decision ids")
    counts={"V0.4.1":0,"V0.4.2":0,"V0.4.3":0}
    for e in entries:
        counts[e["primary_version"]]+=1; c=e["projected_counts"]
        if not e["source_locator"].strip() or not e["derivation_formula"].strip() or any(isinstance(v,bool) or not isinstance(v,int) or v<0 for v in c.values()): raise BenchmarkContractError("invalid projection evidence")
        if e["primary_version"]!="V0.4.1":
            ok=_zero(c) and e["implementation_owner"]=="future_primary" and e["requirement_kind"]=="protected_boundary" and e["scope_topology"]=="unmodeled" and e["conflict_approval_assumption"]=="unmodeled" and e["zero_load_reason"]=="protected_boundary" and not e["blocked_projection"]
            if not ok: raise BenchmarkContractError("protected boundary mismatch")
        elif e["blocked_projection"]:
            if not _zero(c) or e["zero_load_reason"]!="none" or e["scope_topology"]!="unmodeled" or e["conflict_approval_assumption"]!="unmodeled": raise BenchmarkContractError("blocked projection mismatch")
        elif _zero(c) and e["zero_load_reason"]=="none": raise BenchmarkContractError("zero load reason required")
        elif not _zero(c) and e["zero_load_reason"]!="none": raise BenchmarkContractError("nonzero projection cannot have zero-load reason")
    actual={"v041_primary":counts["V0.4.1"],"v042_protected":counts["V0.4.2"],"v043_protected":counts["V0.4.3"]}
    if actual!={"v041_primary":150,"v042_protected":20,"v043_protected":1} or d["decision_population_summary"]!=actual: raise BenchmarkContractError("150/20/1 summary mismatch")
    if d["projection_kind"]=="synthetic_contract_fixture" and not all(i.startswith("V040-SYN-") for i in ids): raise BenchmarkContractError("synthetic ids must be explicit")
    if d["envelope_digest"]!=recompute_envelope_digest(d): raise BenchmarkContractError("envelope digest mismatch")

def _round(v,q): return Decimal(str(v)).quantize(Decimal(q),rounding=ROUND_HALF_EVEN)
def _ratio(n,d):
    d=Decimal(str(d))
    if d<=0: raise BenchmarkContractError("ratio baseline must be positive")
    return (Decimal(str(n))/d).quantize(Decimal("0.000001"),rounding=ROUND_HALF_EVEN)
def _limit(d): return next(x for x in d["thresholds"]["scale_limits"] if x["scale_id"]==d["parameters"]["scale_id"])
def _expected(s):
    if s in ("disjoint","subset-chain"): return ("final","not_applicable","final",True)
    if s=="dense-crossing": return ("unresolvable","not_applicable","unresolvable",False)
    return ("awaiting_approval","final","final",True)
def _terminal_bad(d):
    pre,post,term,final=_expected(d["parameters"]["scenario_id"])
    for r in d["runs"]:
        if r["run_status"]!="completed": continue
        t=r["terminal_trace"]
        if (t["pre_approval"],t["post_approval"],r["terminal_state"],r["final_profile_present"])!=(pre,post,term,final): return True
        if final!=(isinstance(r["final_profile_fingerprint"],str) and bool(SHA256_RE.fullmatch(str(r["final_profile_fingerprint"])))): return True
    return d["execution_status"]=="completed" and d["composer_terminal_state"]!=term

def derive_stop_reasons(d):
    out=set(); st=d["subject_digest_status"]; limit=_limit(d)
    if st["state"]=="stale": out.add("subject_stale")
    for r in d["runs"]:
        status=r["run_status"]
        if status=="timeout": out.add("timeout")
        if status=="process_crash": out.add("process_crash")
        if r["rss"]["status"]=="unavailable": out.add("rss_unavailable")
        for field,bad,reason in (("coverage_conservation","failed","coverage_conservation_failure"),("stable_id_status","drift","stable_id_drift"),("canonical_determinism","mismatched","canonical_nondeterminism"),("fingerprint_determinism","mismatched","fingerprint_nondeterminism"),("contract_status","error","contract_error")):
            if r[field]==bad: out.add(reason)
        if r["run_kind"]=="performance_measured" and status=="completed":
            e=r["timings"]["end_to_end"]
            if e["status"]!="measured": out.add("contract_error")
            elif Decimal(str(e["wall_seconds"]))>Decimal(str(limit["max_wall_seconds"])): out.add("threshold_exceeded")
            if r["rss"]["status"]=="available" and Decimal(str(r["rss"]["peak_rss_mib"]))>Decimal(str(limit["max_peak_rss_mib"])): out.add("threshold_exceeded")
    if _terminal_bad(d): out.add("terminal_state_mismatch")
    for k,threshold,reason in (("wall","wall_median_ratio_2x_to_1x","wall_ratio_exceeded"),("rss","rss_ratio_2x_to_1x","rss_ratio_exceeded"),("output_json","output_json_ratio_2x_to_1x","output_ratio_exceeded")):
        e=d["ratio_evidence"][k]
        if e["status"]=="measured" and _ratio(e["observed_2x"],e["baseline_1x"])>Decimal(str(d["thresholds"][threshold])): out.add(reason)
    if Decimal(str(d["reference_budget"]["elapsed_hours"]))>Decimal(str(d["reference_budget"]["limit_hours"])): out.add("reference_budget_exceeded")
    return sorted(out,key=STOP_REASON_INDEX.__getitem__)

def _subject_ok(d):
    observed=recompute_subject_digest(list(d["subject_manifest"])); st=d["subject_digest_status"]
    if st["observed_subject_digest"]!=observed: raise BenchmarkContractError("observed subject digest mismatch")
    current=observed==d["benchmark_subject_digest"]
    if current!=(st["state"]=="current") or st["revalidation_required"] is current: raise BenchmarkContractError("subject status mismatch")
def _runs_ok(d):
    runs=list(d["runs"]); kind=d["parameters"]["measurement_kind"]; seed=d["parameters"]["permutation_seed"]
    if [r["run_index"] for r in runs]!=list(range(1,len(runs)+1)): raise BenchmarkContractError("run indices invalid")
    if kind=="determinism":
        if isinstance(seed,bool) or not isinstance(seed,int) or len(runs)>1 or any(r["run_kind"]!="determinism" for r in runs) or (d["execution_status"]=="completed" and len(runs)!=1): raise BenchmarkContractError("determinism run shape invalid")
    elif seed is not None: raise BenchmarkContractError("non-determinism result cannot bind permutation seed")
    if kind=="performance":
        expected=["performance_warmup",*(["performance_measured"]*3)]; actual=[r["run_kind"] for r in runs]
        if actual!=expected[:len(actual)] or (d["execution_status"]=="completed" and actual!=expected): raise BenchmarkContractError("performance run shape invalid")
    if kind=="coverage" and (any(r["run_kind"]!="coverage" for r in runs) or (d["execution_status"]=="completed" and not runs)): raise BenchmarkContractError("coverage run shape invalid")
def _run_evidence_ok(d):
    for r in d["runs"]:
        if tuple(r["timings"].keys())!=TIMING_STAGES or r["timings"]["end_to_end"]["status"]!="measured": raise BenchmarkContractError("seven-stage timing contract invalid")
        rss=r["rss"]
        if rss["status"]=="available":
            b,p,x=map(lambda k:Decimal(str(rss[k])),("baseline_rss_mib","peak_rss_mib","delta_peak_rss_mib"))
            if p<b or x!=_round(p-b,"0.001"): raise BenchmarkContractError("RSS delta mismatch")
        m=r["metrics"]
        if m is None:
            if r["run_status"]=="completed" and r["terminal_state"]!="not_reached": raise BenchmarkContractError("completed composer run requires C1 metrics")
        elif tuple(m.keys())!=C1_METRIC_FIELDS or any(isinstance(v,bool) or not isinstance(v,int) or v<0 for v in m.values()): raise BenchmarkContractError("C1 metrics invalid")
def _summary_ok(d):
    s=d["summary"]
    if d["parameters"]["measurement_kind"]!="performance" or d["execution_status"]!="completed":
        if s is not None: raise BenchmarkContractError("summary only allowed for completed performance")
        return
    r=[x for x in d["runs"] if x["run_kind"]=="performance_measured"]
    expected={"median_wall_seconds":median([x["timings"]["end_to_end"]["wall_seconds"] for x in r]),"max_wall_seconds":max(x["timings"]["end_to_end"]["wall_seconds"] for x in r),"median_peak_rss_mib":median([x["rss"]["peak_rss_mib"] for x in r]),"max_peak_rss_mib":max(x["rss"]["peak_rss_mib"] for x in r),"median_output_json_bytes":int(median([x["output_json_bytes"] for x in r])),"max_output_json_bytes":max(x["output_json_bytes"] for x in r)}
    if s!=expected: raise BenchmarkContractError("summary does not reconstruct")
def validate_benchmark_result_semantics(d):
    if not COMMIT_RE.fullmatch(d["benchmark_subject_commit"]): raise BenchmarkContractError("invalid subject commit")
    if d["benchmark_subject_digest_basis"]!="canonical_subject_manifest_v1" or d["rss_protocol"]!=FROZEN_RSS_PROTOCOL or d["thresholds"]!=FROZEN_THRESHOLDS or d["reference_budget"]["limit_hours"]!=3: raise BenchmarkContractError("frozen result contract mismatch")
    _subject_ok(d); _runs_ok(d); _run_evidence_ok(d); _summary_ok(d)
    for e in d["ratio_evidence"].values():
        if e["status"]=="measured" and Decimal(str(e["ratio"]))!=_ratio(e["observed_2x"],e["baseline_1x"]): raise BenchmarkContractError("ratio evidence mismatch")
    reasons=derive_stop_reasons(d)
    if list(d["stop_reasons"])!=reasons: raise BenchmarkContractError("stop reasons not evidence-derived")
    if d["overall_gate"]=="go" and (d["execution_status"]!="completed" or reasons): raise BenchmarkContractError("GO state inconsistent")
    if d["overall_gate"]=="stop" and not reasons: raise BenchmarkContractError("STOP requires reason")
    if d["execution_status"]=="completed" and any(r["run_status"]!="completed" for r in d["runs"]): raise BenchmarkContractError("completed execution contains interrupted run")
    if d["result_digest"]!=recompute_result_digest(d): raise BenchmarkContractError("result digest mismatch")
