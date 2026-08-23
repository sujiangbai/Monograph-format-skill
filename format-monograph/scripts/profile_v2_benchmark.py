#!/usr/bin/env python3
"""Pure P3a-C2 benchmark contract helpers; no runner/runtime integration."""
from __future__ import annotations
import copy, hashlib, json, math, re
from decimal import Decimal, ROUND_HALF_EVEN
from pathlib import Path
from statistics import median
from collections.abc import Mapping

class BenchmarkContractError(ValueError): pass
SCALES=("0.5x","1.0x","1.5x","2.0x")
SCENARIOS=("disjoint","subset-chain","dense-crossing","mixed-conflict-approval")
DETERMINISM_SCALES=("1.5x","2.0x")
PERFORMANCE_CELLS={("0.5x","mixed-conflict-approval"),("1.0x","mixed-conflict-approval"),*((s,c) for s in DETERMINISM_SCALES for c in SCENARIOS)}
TIMING_STAGES=("synthetic_generation","schema_registry_validation","compose","approval_generation","apply","canonical_serialization","end_to_end")
C1_METRIC_FIELDS=("input_asset_count","input_rule_count","input_binding_count","expected_key_count","candidate_count","candidate_group_count","partition_count","conflict_count","proposal_count","blocker_count","max_candidates_per_key","max_partition_width","max_repartition_depth")
STOP_REASON_ORDER=("threshold_exceeded","wall_ratio_exceeded","rss_ratio_exceeded","output_ratio_exceeded","timeout","process_crash","rss_unavailable","coverage_conservation_failure","stable_id_drift","canonical_nondeterminism","fingerprint_nondeterminism","terminal_state_mismatch","subject_stale","contract_error","reference_budget_exceeded")
STOP_REASON_INDEX={v:i for i,v in enumerate(STOP_REASON_ORDER)}
SHA256_RE=re.compile(r"^sha256:[a-f0-9]{64}$"); COMMIT_RE=re.compile(r"^[a-f0-9]{40}$"); DRIVE_RE=re.compile(r"^[A-Za-z]:")
FROZEN_SCALE_FACTORS={"0.5x":.5,"1.0x":1.0,"1.5x":1.5,"2.0x":2.0}
FROZEN_SCENARIO_SEMANTICS={
 "disjoint":{"pre_approval_terminal":"final","post_approval_terminal":"not_applicable","terminal_state":"final","final_requirement":"required"},
 "subset-chain":{"pre_approval_terminal":"final","post_approval_terminal":"not_applicable","terminal_state":"final","final_requirement":"required"},
 "dense-crossing":{"pre_approval_terminal":"unresolvable","post_approval_terminal":"not_applicable","terminal_state":"unresolvable","final_requirement":"forbidden"},
 "mixed-conflict-approval":{"pre_approval_terminal":"awaiting_approval","post_approval_terminal":"final","terminal_state":"final","final_requirement":"conditional"}}
FROZEN_RSS_PROTOCOL={"measurement_method":"external_supervisor_child_peak_working_set","sampling_interval_seconds":.05,"child_process_scope":"benchmark_child_process_only","record_baseline":True,"record_delta":True,"unavailable_policy":"fail_closed","delta_rounding":"mib_3_decimal_places_half_even"}
FROZEN_THRESHOLDS={"scale_limits":[{"scale_id":"0.5x","max_wall_seconds":60,"max_peak_rss_mib":512},{"scale_id":"1.0x","max_wall_seconds":60,"max_peak_rss_mib":512},{"scale_id":"1.5x","max_wall_seconds":60,"max_peak_rss_mib":512},{"scale_id":"2.0x","max_wall_seconds":120,"max_peak_rss_mib":1024}],"wall_median_ratio_2x_to_1x":6,"rss_ratio_2x_to_1x":3,"output_json_ratio_2x_to_1x":3}
FROZEN_REPETITIONS={"coverage_min_runs_per_cell":1,"performance_warmup_runs_per_cell":1,"performance_measured_runs_per_cell":3,"determinism_runs_per_seed":1}
ABSOLUTE_GATE_RUN_KINDS={"coverage","performance_measured","determinism"}
RATIO_SUMMARY_FIELDS={"wall":"median_wall_seconds","rss":"median_peak_rss_mib","output_json":"median_output_json_bytes"}
RATIO_THRESHOLD_FIELDS={"wall":"wall_median_ratio_2x_to_1x","rss":"rss_ratio_2x_to_1x","output_json":"output_json_ratio_2x_to_1x"}
RATIO_STOP_REASONS={"wall":"wall_ratio_exceeded","rss":"rss_ratio_exceeded","output_json":"output_ratio_exceeded"}
PROJECTION_COUNT_FIELDS=("rule_fragment","binding","key","candidate")
REQUIRED_UNMODELED_DIMENSIONS={"full_production_property_registry","multi_property_production_distribution","real_monograph_base_assets","docx_runtime"}
REPRESENTATIVENESS_SCOPE="non_production_single_core_probe"
_SCHEMA_FILES={
 "envelope":"projected-envelope.schema.json",
 "config":"benchmark-config.schema.json",
 "result":"benchmark-result.schema.json",
}

def _finite(v):
 if isinstance(v,float) and not math.isfinite(v): raise BenchmarkContractError("non-finite JSON")
 if isinstance(v,Mapping):
  for x in v.values(): _finite(x)
 elif isinstance(v,(list,tuple)):
  for x in v: _finite(x)
def canonical_json_bytes(v): _finite(v); return json.dumps(v,ensure_ascii=False,sort_keys=True,separators=(",",":"),allow_nan=False).encode()
def canonical_sha256(v): return "sha256:"+hashlib.sha256(canonical_json_bytes(v)).hexdigest()
def _self_digest(d,k): x=copy.deepcopy(dict(d)); x.pop(k,None); return canonical_sha256(x)
def recompute_config_digest(d): return _self_digest(d,"config_digest")
def recompute_envelope_digest(d): return _self_digest(d,"envelope_digest")
def recompute_result_digest(d): return _self_digest(d,"result_digest")

def _validate_contract_schema(kind, document):
 """Apply the committed C2A JSON Schema before any suite semantics."""
 try:
  from jsonschema import Draft202012Validator
 except ImportError as exc:
  raise BenchmarkContractError("jsonschema dependency unavailable") from exc
 try:
  filename=_SCHEMA_FILES[kind]
 except KeyError as exc:
  raise BenchmarkContractError("unknown contract schema") from exc
 schema_path=Path(__file__).resolve().parents[1] / "references" / "benchmarks" / "v0412" / "p3a-c2" / filename
 try:
  schema=json.loads(schema_path.read_text(encoding="utf-8"))
  errors=sorted(Draft202012Validator(schema).iter_errors(document),key=lambda error:list(error.absolute_path))
 except (OSError,json.JSONDecodeError) as exc:
  raise BenchmarkContractError("contract schema unavailable") from exc
 if errors: raise BenchmarkContractError("%s schema invalid" % kind)

def canonical_subject_path(p):
 if not isinstance(p,str) or not p or "\\" in p or DRIVE_RE.match(p) or p.startswith("/") or p.endswith("/") or "//" in p or any(x in {"",".",".."} for x in p.split("/")): raise BenchmarkContractError("noncanonical subject path")
 return p
def recompute_subject_digest(m):
 files=[]; paths=[]
 for e in m:
  p=canonical_subject_path(e.get("path")); h=e.get("sha256")
  if not isinstance(h,str) or not SHA256_RE.fullmatch(h): raise BenchmarkContractError("bad subject digest")
  paths.append(p); files.append({"path":p,"sha256":h})
 if paths!=sorted(paths) or len(paths)!=len(set(paths)): raise BenchmarkContractError("subject paths unsorted/duplicate")
 return canonical_sha256({"basis":"canonical_subject_manifest_v1","files":files})
def expected_coverage_cells(): return {(s,c) for s in SCALES for c in SCENARIOS}
def expected_performance_cells(): return set(PERFORMANCE_CELLS)
def expected_determinism_cells(seeds): return {(s,c,int(seed)) for s in DETERMINISM_SCALES for c in SCENARIOS for seed in seeds}
def _exact(a,e,label):
 if len(a)!=len(set(a)) or set(a)!=e: raise BenchmarkContractError(label+" matrix mismatch")

def validate_benchmark_config_semantics(d):
 seeds=list(d["generation"]["permutation_seeds"])
 if len(seeds)<5 or len(seeds)!=len(set(seeds)): raise BenchmarkContractError("seed contract")
 scales={x["scale_id"]:x for x in d["scales"]}; scenarios={x["scenario_id"]:x for x in d["scenarios"]}
 if set(scales)!=set(SCALES) or len(d["scales"])!=4 or set(scenarios)!=set(SCENARIOS) or len(d["scenarios"])!=4: raise BenchmarkContractError("inventory")
 for k,f in FROZEN_SCALE_FACTORS.items():
  if scales[k]["factor"]!=f or set(scales[k]["scenario_ids"])!=set(SCENARIOS) or len(scales[k]["scenario_ids"])!=4: raise BenchmarkContractError("scale semantics")
 for k,e in FROZEN_SCENARIO_SEMANTICS.items():
  if any(scenarios[k][x]!=e[x] for x in ("pre_approval_terminal","post_approval_terminal","final_requirement")): raise BenchmarkContractError("scenario semantics")
 m=d["matrices"]; _exact([(x["scale_id"],x["scenario_id"]) for x in m["coverage_cells"]],expected_coverage_cells(),"coverage"); _exact([(x["scale_id"],x["scenario_id"]) for x in m["performance_cells"]],expected_performance_cells(),"performance"); _exact([(x["scale_id"],x["scenario_id"],x["permutation_seed"]) for x in m["determinism_cells"]],expected_determinism_cells(seeds),"determinism")
 if d["repetitions"]!=FROZEN_REPETITIONS or d["rss_protocol"]!=FROZEN_RSS_PROTOCOL or d["thresholds"]!=FROZEN_THRESHOLDS or d["total_reference_budget_hours"]!=3: raise BenchmarkContractError("frozen config")
 b=d.get("projection_binding")
 if not isinstance(b,Mapping) or b.get("base_scale_id")!="1.0x" or b.get("compose_projection_strategy")!="single_core_probe" or b.get("representativeness_scope")!=REPRESENTATIVENESS_SCOPE or b.get("revalidation_required") is not True: raise BenchmarkContractError("projection binding")
 c=b.get("aggregate_counts")
 if not isinstance(c,Mapping) or set(c)!=set(PROJECTION_COUNT_FIELDS) or any(isinstance(v,bool) or not isinstance(v,int) or v<0 for v in c.values()): raise BenchmarkContractError("aggregate counts")
 if d["config_digest"]!=recompute_config_digest(d): raise BenchmarkContractError("config digest")

def _zero(c): return all(v==0 for v in c.values())
def _synthetic(e): t=e["derivation_formula"].lower(); return e["decision_id"].startswith("V040-SYN-") and e["source_locator"].startswith("plan:synthetic_") and "synthetic" in t and ("fixture" in t or "not a production" in t)
def _formal_bad(e): t=e["derivation_formula"].lower(); return e["decision_id"].startswith("V040-SYN-") or e["source_locator"].startswith("plan:synthetic_") or "synthetic" in t or "not a production decision mapping" in t
def validate_projected_envelope_semantics(d):
 es=list(d["entries"]); ids=[e["decision_id"] for e in es]
 if len(es)!=171 or len(ids)!=len(set(ids)): raise BenchmarkContractError("171 IDs")
 versions={"V0.4.1":0,"V0.4.2":0,"V0.4.3":0}
 for e in es:
  versions[e["primary_version"]]+=1; c=e["projected_counts"]
  if any(isinstance(v,bool) or not isinstance(v,int) or v<0 for v in c.values()) or not e["source_locator"].strip() or not e["derivation_formula"].strip(): raise BenchmarkContractError("projection evidence")
  if e["primary_version"]!="V0.4.1":
   if not(_zero(c) and e["implementation_owner"]=="future_primary" and e["requirement_kind"]=="protected_boundary" and e["scope_topology"]=="unmodeled" and e["conflict_approval_assumption"]=="unmodeled" and e["zero_load_reason"]=="protected_boundary" and not e["blocked_projection"]): raise BenchmarkContractError("protected boundary")
  elif e["blocked_projection"]:
   if not _zero(c) or e["zero_load_reason"]!="none" or e["scope_topology"]!="unmodeled" or e["conflict_approval_assumption"]!="unmodeled": raise BenchmarkContractError("blocked projection")
  elif _zero(c):
   if not(e["zero_load_reason"]=="no_base_appearance" and e["implementation_owner"] in {"p3a_c","p3a_r","p3b_b","p3b_o","p4","p5","p6","p7"} and e["requirement_kind"] in {"base_appearance","system_projection"}): raise BenchmarkContractError("zero-load V041 entry")
  elif e["zero_load_reason"]!="none": raise BenchmarkContractError("zero-load reason")
 actual={"v041_primary":versions["V0.4.1"],"v042_protected":versions["V0.4.2"],"v043_protected":versions["V0.4.3"]}
 if actual!={"v041_primary":150,"v042_protected":20,"v043_protected":1} or d["decision_population_summary"]!=actual: raise BenchmarkContractError("population")
 aggregate={k:sum(e["projected_counts"][k] for e in es if e["primary_version"]=="V0.4.1" and not e["blocked_projection"]) for k in PROJECTION_COUNT_FIELDS}
 if any(value<=0 for value in aggregate.values()): raise BenchmarkContractError("nonblocking V041 aggregate")
 if d["projection_kind"]=="synthetic_contract_fixture" and not all(_synthetic(e) for e in es): raise BenchmarkContractError("synthetic namespace")
 if d["projection_kind"]=="formal_planning_projection" and any(_formal_bad(e) for e in es): raise BenchmarkContractError("formal namespace")
 if d["compose_projection_strategy"]!="single_core_probe": raise BenchmarkContractError("strategy")
 if d["envelope_digest"]!=recompute_envelope_digest(d): raise BenchmarkContractError("envelope digest")
def aggregate_projected_envelope_counts(e):
 validate_projected_envelope_semantics(e); out={k:0 for k in PROJECTION_COUNT_FIELDS}
 for x in e["entries"]:
  for k in out: out[k]+=x["projected_counts"][k]
 return out
def validate_benchmark_config_against_envelope(c,e):
 validate_benchmark_config_semantics(c); validate_projected_envelope_semantics(e); b=c["projection_binding"]
 if b["projected_envelope_digest"]!=recompute_envelope_digest(e) or b["projection_kind"]!=e["projection_kind"] or b["aggregate_counts"]!=aggregate_projected_envelope_counts(e) or b["compose_projection_strategy"]!=e["compose_projection_strategy"] or b["base_scale_id"]!="1.0x" or not REQUIRED_UNMODELED_DIMENSIONS<=set(e["unmodeled_dimensions"]): raise BenchmarkContractError("config/envelope binding")

def _round(v,q): return Decimal(str(v)).quantize(Decimal(q),rounding=ROUND_HALF_EVEN)
def _ratio(n,d):
 d=Decimal(str(d))
 if d<=0: raise BenchmarkContractError("ratio baseline")
 return (Decimal(str(n))/d).quantize(Decimal("0.000001"),rounding=ROUND_HALF_EVEN)
def _limit(d): return next(x for x in d["thresholds"]["scale_limits"] if x["scale_id"]==d["parameters"]["scale_id"])
def _terminal_bad(d):
 e=FROZEN_SCENARIO_SEMANTICS[d["parameters"]["scenario_id"]]
 for r in d["runs"]:
  if r["run_status"]!="completed": continue
  t=r["terminal_trace"]; final=e["terminal_state"]=="final"; fp=isinstance(r["final_profile_fingerprint"],str) and bool(SHA256_RE.fullmatch(str(r["final_profile_fingerprint"])))
  if (t["pre_approval"],t["post_approval"],r["terminal_state"],r["final_profile_present"],fp)!=(e["pre_approval_terminal"],e["post_approval_terminal"],e["terminal_state"],final,final): return True
 return d["execution_status"]=="completed" and d["composer_terminal_state"]!=e["terminal_state"]
def derive_stop_reasons(d):
 out=set(); limit=_limit(d)
 if d["subject_digest_status"]["state"]=="stale": out.add("subject_stale")
 for r in d["runs"]:
  if r["run_status"] in {"timeout","process_crash"}: out.add(r["run_status"])
  if r["rss"]["status"]=="unavailable": out.add("rss_unavailable")
  for f,bad,reason in (("coverage_conservation","failed","coverage_conservation_failure"),("stable_id_status","drift","stable_id_drift"),("canonical_determinism","mismatched","canonical_nondeterminism"),("fingerprint_determinism","mismatched","fingerprint_nondeterminism"),("contract_status","error","contract_error")):
   if r[f]==bad: out.add(reason)
  if r["run_kind"] in ABSOLUTE_GATE_RUN_KINDS and r["run_status"]=="completed":
   e=r["timings"]["end_to_end"]
   if e["status"]!="measured": out.add("contract_error")
   elif Decimal(str(e["wall_seconds"]))>Decimal(str(limit["max_wall_seconds"])): out.add("threshold_exceeded")
   if r["rss"]["status"]=="available" and Decimal(str(r["rss"]["peak_rss_mib"]))>Decimal(str(limit["max_peak_rss_mib"])): out.add("threshold_exceeded")
 if _terminal_bad(d): out.add("terminal_state_mismatch")
 for m,tf in RATIO_THRESHOLD_FIELDS.items():
  e=d["ratio_evidence"][m]
  if e["status"]=="measured" and _ratio(e["observed_2x"],e["baseline_1x"])>Decimal(str(d["thresholds"][tf])): out.add(RATIO_STOP_REASONS[m])
 if Decimal(str(d["reference_budget"]["elapsed_hours"]))>Decimal(str(d["reference_budget"]["limit_hours"])): out.add("reference_budget_exceeded")
 return sorted(out,key=STOP_REASON_INDEX.__getitem__)
def _subject_ok(d):
 obs=recompute_subject_digest(d["subject_manifest"]); st=d["subject_digest_status"]; cur=obs==d["benchmark_subject_digest"]
 if st["observed_subject_digest"]!=obs or cur!=(st["state"]=="current") or st["revalidation_required"]!=(not cur): raise BenchmarkContractError("subject status")
def _runs_ok(d):
 rs=list(d["runs"]); k=d["parameters"]["measurement_kind"]; seed=d["parameters"]["permutation_seed"]
 if [r["run_index"] for r in rs]!=list(range(1,len(rs)+1)): raise BenchmarkContractError("run indices")
 intr=[i for i,r in enumerate(rs) if r["run_status"] in {"timeout","process_crash"}]
 if intr and (intr!=[len(rs)-1] or d["execution_status"]!="stopped"): raise BenchmarkContractError("interrupt sequence")
 if k=="determinism" and (isinstance(seed,bool) or not isinstance(seed,int) or len(rs)>1 or any(r["run_kind"]!="determinism" for r in rs) or (d["execution_status"]=="completed" and len(rs)!=1)): raise BenchmarkContractError("det run shape")
 if k!="determinism" and seed is not None: raise BenchmarkContractError("unexpected seed")
 if k=="performance":
  exp=["performance_warmup","performance_measured","performance_measured","performance_measured"]; act=[r["run_kind"] for r in rs]
  if act!=exp[:len(act)] or (d["execution_status"]=="completed" and act!=exp): raise BenchmarkContractError("perf run shape")
 if k=="coverage" and (any(r["run_kind"]!="coverage" for r in rs) or (d["execution_status"]=="completed" and not rs)): raise BenchmarkContractError("coverage run shape")
def _timing_statuses(r): return {x:r["timings"][x]["status"] for x in TIMING_STAGES}
def _interrupted_terminal_expectation(r,scenario):
 st=_timing_statuses(r)
 if st["compose"]!="measured": return ("not_reached",{"pre_approval":"not_reached","post_approval":"not_reached"},False)
 e=FROZEN_SCENARIO_SEMANTICS[scenario]
 if scenario=="mixed-conflict-approval" and st["approval_generation"]!="measured": return ("awaiting_approval",{"pre_approval":"awaiting_approval","post_approval":"not_reached"},False)
 return (e["terminal_state"],{"pre_approval":e["pre_approval_terminal"],"post_approval":e["post_approval_terminal"]},e["terminal_state"]=="final")
def _validate_interrupted_timing(r,scenario):
 st=_timing_statuses(r)
 if st["end_to_end"]!="measured": raise BenchmarkContractError("interrupt elapsed timing")
 if scenario=="mixed-conflict-approval":
  if st["approval_generation"] not in {"measured","not_reached"}: raise BenchmarkContractError("mixed interrupt approval timing")
 elif st["approval_generation"]!="not_applicable": raise BenchmarkContractError("interrupt approval NA")
 for x in ("synthetic_generation","schema_registry_validation","compose","canonical_serialization"):
  if st[x]=="not_applicable": raise BenchmarkContractError("applicable interrupt stage marked NA")
 if scenario!="dense-crossing" and st["apply"]=="not_applicable": raise BenchmarkContractError("applicable interrupt apply marked NA")
 seen=False
 for x in TIMING_STAGES[:-1]:
  status=st[x]
  if status not in {"measured","not_applicable","not_reached"}: raise BenchmarkContractError("interrupt timing status")
  if status=="not_reached": seen=True
  elif status=="measured" and seen: raise BenchmarkContractError("interrupt timing resumed")
def _validate_metrics(m):
 if not isinstance(m,Mapping) or set(m)!=set(C1_METRIC_FIELDS) or any(isinstance(v,bool) or not isinstance(v,int) or v<0 for v in m.values()): raise BenchmarkContractError("metrics")
def _validate_bytes(v,label,required):
 if v is None and not required: return
 if isinstance(v,bool) or not isinstance(v,int) or v<0: raise BenchmarkContractError(label+" byte evidence")
def _run_evidence_ok(d):
 scenario=d["parameters"]["scenario_id"]; kind=d["parameters"]["measurement_kind"]
 for r in d["runs"]:
  if set(r["timings"])!=set(TIMING_STAGES): raise BenchmarkContractError("timing keys")
  rss=r["rss"]
  if rss["status"]=="available" and (Decimal(str(rss["peak_rss_mib"]))<Decimal(str(rss["baseline_rss_mib"])) or Decimal(str(rss["delta_peak_rss_mib"]))!=_round(Decimal(str(rss["peak_rss_mib"]))-Decimal(str(rss["baseline_rss_mib"])),"0.001")): raise BenchmarkContractError("rss delta")
  st=_timing_statuses(r)
  if r["run_status"]=="completed":
   if any(st[x]!="measured" for x in ("synthetic_generation","schema_registry_validation","compose","canonical_serialization","end_to_end")): raise BenchmarkContractError("timing evidence")
   if r["final_profile_present"]:
    if st["apply"]!="measured": raise BenchmarkContractError("apply timing")
   elif st["apply"] not in {"measured","not_applicable"}: raise BenchmarkContractError("non-final apply timing")
   if scenario=="mixed-conflict-approval" and st["approval_generation"]!="measured": raise BenchmarkContractError("approval timing")
   if scenario!="mixed-conflict-approval" and st["approval_generation"]!="not_applicable": raise BenchmarkContractError("approval NA")
   _validate_metrics(r["metrics"]); _validate_bytes(r["input_json_bytes"],"input",True); _validate_bytes(r["output_json_bytes"],"output",True)
   if r["coverage_conservation"] not in {"passed","failed"} or r["stable_id_status"] not in {"stable","drift"}: raise BenchmarkContractError("mandatory evidence")
   if kind=="determinism" and (r["canonical_determinism"] not in {"matched","mismatched"} or r["fingerprint_determinism"] not in {"matched","mismatched"}): raise BenchmarkContractError("det evidence")
   if kind!="determinism" and (r["canonical_determinism"]!="not_applicable" or r["fingerprint_determinism"]!="not_applicable"): raise BenchmarkContractError("det NA")
  else:
   _validate_interrupted_timing(r,scenario); terminal,trace,final=_interrupted_terminal_expectation(r,scenario); fp=isinstance(r["final_profile_fingerprint"],str) and bool(SHA256_RE.fullmatch(str(r["final_profile_fingerprint"])))
   if r["terminal_state"]!=terminal or r["terminal_trace"]!=trace or r["final_profile_present"]!=final or fp!=final or d["composer_terminal_state"]!=terminal: raise BenchmarkContractError("interrupted terminal evidence")
   _validate_bytes(r["input_json_bytes"],"input",False); _validate_bytes(r["output_json_bytes"],"output",False)
   if st["compose"]=="measured":
    _validate_metrics(r["metrics"])
    if r["coverage_conservation"] not in {"passed","failed"} or r["stable_id_status"] not in {"stable","drift"}: raise BenchmarkContractError("reached interrupted evidence")
   elif r["metrics"] is not None or r["coverage_conservation"]!="not_reached" or r["stable_id_status"]!="not_reached": raise BenchmarkContractError("future interrupted evidence")
   if kind=="determinism":
    reached=st["canonical_serialization"]=="measured"
    if reached and (r["canonical_determinism"] not in {"matched","mismatched"} or r["fingerprint_determinism"] not in {"matched","mismatched"}): raise BenchmarkContractError("reached determinism evidence")
    if not reached and (r["canonical_determinism"]!="not_reached" or r["fingerprint_determinism"]!="not_reached"): raise BenchmarkContractError("future determinism evidence")
   elif r["canonical_determinism"]!="not_applicable" or r["fingerprint_determinism"]!="not_applicable": raise BenchmarkContractError("interrupted det NA")
def _summary_ok(d):
 s=d["summary"]
 if d["parameters"]["measurement_kind"]!="performance" or d["execution_status"]!="completed":
  if s is not None: raise BenchmarkContractError("summary applicability")
  return
 r=[x for x in d["runs"] if x["run_kind"]=="performance_measured"]
 if any(x["rss"]["status"]!="available" for x in r): raise BenchmarkContractError("summary RSS")
 exp={"median_wall_seconds":median([x["timings"]["end_to_end"]["wall_seconds"] for x in r]),"max_wall_seconds":max(x["timings"]["end_to_end"]["wall_seconds"] for x in r),"median_peak_rss_mib":median([x["rss"]["peak_rss_mib"] for x in r]),"max_peak_rss_mib":max(x["rss"]["peak_rss_mib"] for x in r),"median_output_json_bytes":int(median([x["output_json_bytes"] for x in r])),"max_output_json_bytes":max(x["output_json_bytes"] for x in r)}
 if s!=exp: raise BenchmarkContractError("summary reconstruct")
def validate_benchmark_result_semantics(d):
 if not COMMIT_RE.fullmatch(d["benchmark_subject_commit"]) or d["benchmark_subject_digest_basis"]!="canonical_subject_manifest_v1" or d["rss_protocol"]!=FROZEN_RSS_PROTOCOL or d["thresholds"]!=FROZEN_THRESHOLDS or d["reference_budget"]["limit_hours"]!=3: raise BenchmarkContractError("result frozen fields")
 _subject_ok(d); _runs_ok(d); _run_evidence_ok(d); _summary_ok(d)
 if d["execution_status"]!="completed" and any(e["status"]=="measured" for e in d["ratio_evidence"].values()): raise BenchmarkContractError("stopped result cannot claim ratio evidence")
 for e in d["ratio_evidence"].values():
  if e["status"]=="measured" and Decimal(str(e["ratio"]))!=_ratio(e["observed_2x"],e["baseline_1x"]): raise BenchmarkContractError("ratio arithmetic")
 reasons=derive_stop_reasons(d)
 if list(d["stop_reasons"])!=reasons or (d["overall_gate"]=="go" and (d["execution_status"]!="completed" or reasons)) or (d["overall_gate"]=="stop" and not reasons) or (d["execution_status"]=="completed" and any(r["run_status"]!="completed" for r in d["runs"])) or d["result_digest"]!=recompute_result_digest(d): raise BenchmarkContractError("result semantic state")
def _matrix_sets(c):
 m=c["matrices"]; return ({(e["scale_id"],e["scenario_id"]) for e in m["coverage_cells"]},{(e["scale_id"],e["scenario_id"]) for e in m["performance_cells"]},{(e["scale_id"],e["scenario_id"],e["permutation_seed"]) for e in m["determinism_cells"]})
def validate_benchmark_result_against_config(r,c):
 validate_benchmark_config_semantics(c); validate_benchmark_result_semantics(r)
 if r["benchmark_config_digest"]!=recompute_config_digest(c) or r["parameters"]["generation_seed"]!=c["generation"]["generation_seed"] or r["rss_protocol"]!=c["rss_protocol"] or r["thresholds"]!=c["thresholds"] or r["output_json_bytes_basis"]!=c["output_json_bytes_basis"] or r["input_json_bytes_basis"]!=c["input_json_bytes_basis"] or r["reference_budget"]["limit_hours"]!=c["total_reference_budget_hours"]: raise BenchmarkContractError("result/config frozen fields")
 cov,perf,det=_matrix_sets(c); p=r["parameters"]; cell=(p["scale_id"],p["scenario_id"]); k=p["measurement_kind"]
 if k=="coverage" and cell not in cov or k=="performance" and cell not in perf: raise BenchmarkContractError("matrix cell")
 if k=="determinism" and (p["permutation_seed"] not in set(c["generation"]["permutation_seeds"]) or p["scale_id"] not in DETERMINISM_SCALES or (p["scale_id"],p["scenario_id"],p["permutation_seed"]) not in det): raise BenchmarkContractError("det matrix")
 if k=="performance":
  exp=["performance_warmup"]*c["repetitions"]["performance_warmup_runs_per_cell"]+["performance_measured"]*c["repetitions"]["performance_measured_runs_per_cell"]; act=[x["run_kind"] for x in r["runs"]]
  if act!=exp[:len(act)] or (r["execution_status"]=="completed" and act!=exp): raise BenchmarkContractError("perf repetitions")
 if k=="determinism" and r["execution_status"]=="completed" and len(r["runs"])!=c["repetitions"]["determinism_runs_per_seed"]: raise BenchmarkContractError("det repetitions")
 if k=="coverage" and r["execution_status"]=="completed" and len(r["runs"])<c["repetitions"]["coverage_min_runs_per_cell"]: raise BenchmarkContractError("coverage repetitions")
def benchmark_comparison_identity(r):
 _subject_ok(r); return {"benchmark_subject_commit":r["benchmark_subject_commit"],"benchmark_subject_digest":r["benchmark_subject_digest"],"benchmark_subject_digest_basis":r["benchmark_subject_digest_basis"],"canonical_subject_identity":recompute_subject_digest(r["subject_manifest"]),"environment":copy.deepcopy(r["environment"]),"command_template":r["command_template"],"benchmark_config_digest":r["benchmark_config_digest"],"generation_seed":r["parameters"]["generation_seed"],"rss_protocol":copy.deepcopy(r["rss_protocol"]),"thresholds":copy.deepcopy(r["thresholds"])}
def validate_ratio_evidence(one,two,c):
 if one is None or two is None: raise BenchmarkContractError("missing ratio pair")
 validate_benchmark_result_against_config(one,c); validate_benchmark_result_against_config(two,c); a=one["parameters"]; bb=two["parameters"]
 if a["measurement_kind"]!="performance" or bb["measurement_kind"]!="performance" or a["scale_id"]!="1.0x" or bb["scale_id"]!="2.0x" or a["scenario_id"]!=bb["scenario_id"] or one["execution_status"]!="completed" or two["execution_status"]!="completed" or benchmark_comparison_identity(one)!=benchmark_comparison_identity(two): raise BenchmarkContractError("comparison identity/shape")
 perf=_matrix_sets(c)[1]
 if ("1.0x",a["scenario_id"]) not in perf or ("2.0x",a["scenario_id"]) not in perf: raise BenchmarkContractError("ratio matrix")
 expected={m:(one["summary"][f],two["summary"][f],_ratio(two["summary"][f],one["summary"][f])) for m,f in RATIO_SUMMARY_FIELDS.items()}
 for m,(base,obs,ratio) in expected.items():
  e=two["ratio_evidence"][m]
  if e["status"]!="measured" or Decimal(str(e["baseline_1x"]))!=Decimal(str(base)) or Decimal(str(e["observed_2x"]))!=Decimal(str(obs)) or Decimal(str(e["ratio"]))!=ratio: raise BenchmarkContractError("ratio binding")
 expected_reasons={RATIO_STOP_REASONS[m] for m,(_,_,ratio) in expected.items() if ratio>Decimal(str(c["thresholds"][RATIO_THRESHOLD_FIELDS[m]]))}; actual=set(two["stop_reasons"])&set(RATIO_STOP_REASONS.values())
 if actual!=expected_reasons or (expected_reasons and two["overall_gate"]!="stop") or (not expected_reasons and two["overall_gate"]=="stop" and not(set(two["stop_reasons"])-set(RATIO_STOP_REASONS.values()))): raise BenchmarkContractError("ratio gate")
def _logical_key(r):
 p=r["parameters"]; return (p["measurement_kind"],p["scale_id"],p["scenario_id"],p["permutation_seed"])
def validate_benchmark_result_set(results,c):
 """Validate partial collection consistency; this is not a completeness gate."""
 validate_benchmark_config_semantics(c); idx={}
 for r in results:
  validate_benchmark_result_against_config(r,c); k=_logical_key(r)
  if k in idx: raise BenchmarkContractError("duplicate result key")
  idx[k]=r
 perf=_matrix_sets(c)[1]
 for s in {x for x in SCENARIOS if ("1.0x",x) in perf and ("2.0x",x) in perf}:
  one=idx.get(("performance","1.0x",s,None)); two=idx.get(("performance","2.0x",s,None))
  if (one is None)!=(two is None): raise BenchmarkContractError("missing pair")
  if one is not None:
   if one["execution_status"]=="completed" and two["execution_status"]=="completed": validate_ratio_evidence(one,two,c)
   elif any(e["status"]=="measured" for e in two["ratio_evidence"].values()): raise BenchmarkContractError("ratio evidence requires completed pair")
def validate_complete_benchmark_suite(results,c,e):
 # This is the one public closure point for C2B: structure precedes digests,
 # semantics, collection completeness, and the suite gate.
 _validate_contract_schema("envelope",e)
 _validate_contract_schema("config",c)
 rs=list(results)
 for r in rs: _validate_contract_schema("result",r)
 validate_benchmark_config_against_envelope(c,e); validate_benchmark_result_set(rs,c)
 if rs:
  identity=benchmark_comparison_identity(rs[0])
  for r in rs[1:]:
   if benchmark_comparison_identity(r)!=identity: raise BenchmarkContractError("benchmark suite campaign identity mismatch")
 idx={_logical_key(r):r for r in rs}; cov,perf,det=_matrix_sets(c)
 required={("coverage",s,sc,None) for s,sc in cov}|{("performance",s,sc,None) for s,sc in perf}|{("determinism",s,sc,seed) for s,sc,seed in det}
 if set(idx)!=required: raise BenchmarkContractError("benchmark suite incomplete")
 for sc in {x for x in SCENARIOS if ("1.0x",x) in perf and ("2.0x",x) in perf}:
  one=idx[("performance","1.0x",sc,None)]; two=idx[("performance","2.0x",sc,None)]
  if one["execution_status"]=="completed" and two["execution_status"]=="completed": validate_ratio_evidence(one,two,c)
  elif any(e["status"]=="measured" for e in two["ratio_evidence"].values()): raise BenchmarkContractError("incomplete comparison cannot carry ratio evidence")
 return {"structurally_complete":True,"overall_gate":"stop" if any(r["overall_gate"]=="stop" for r in rs) else "go"}
def validate_benchmark_result_context(r,c,e):
 validate_benchmark_config_against_envelope(c,e); validate_benchmark_result_against_config(r,c)
 return {"projection_kind":e["projection_kind"],"compose_projection_strategy":e["compose_projection_strategy"],"unmodeled_dimensions":tuple(sorted(e["unmodeled_dimensions"])),"base_scale_id":c["projection_binding"]["base_scale_id"],"aggregate_counts":copy.deepcopy(c["projection_binding"]["aggregate_counts"]),"representativeness_scope":REPRESENTATIVENESS_SCOPE,"production_representative":False,"revalidation_required":True}
