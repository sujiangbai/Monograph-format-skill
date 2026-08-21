import ast, copy, importlib.util, json, math, unittest
from pathlib import Path
from jsonschema import Draft202012Validator
ROOT=Path(__file__).resolve().parents[1]; BENCH=ROOT/'format-monograph/references/benchmarks/v0412/p3a-c2'; FIX=ROOT/'tests/fixtures/v0412/p3a_c2'; MOD=ROOT/'format-monograph/scripts/profile_v2_benchmark.py'; COMP=ROOT/'format-monograph/scripts/profile_v2_composer.py'
spec=importlib.util.spec_from_file_location('profile_v2_benchmark',MOD); b=importlib.util.module_from_spec(spec); spec.loader.exec_module(b)
SCHEMAS={k:json.loads((BENCH/f).read_text()) for k,f in {'ENVELOPE':'projected-envelope.schema.json','CONFIG':'benchmark-config.schema.json','RESULT':'benchmark-result.schema.json'}.items()}; V={k:Draft202012Validator(v) for k,v in SCHEMAS.items()}
R={'STATE':12,'MATRIX':12,'TIMING':8,'RSS':10,'METRICS':8,'ENVELOPE':12,'RESULT':14,'BOUNDARY':10}

def load(n): return json.loads((FIX/n).read_text())
def stamp(d,k='result_digest'):
 fn={'result_digest':b.recompute_result_digest,'config_digest':b.recompute_config_digest,'envelope_digest':b.recompute_envelope_digest}[k]; d[k]=fn(d); return d
def sem(fn,d):
 try: fn(d); return True
 except b.BenchmarkContractError: return False
def schema(k,d): return not list(V[k].iter_errors(d))
def env_fixture(name):
 r=load(name); q=r['synthetic_recipe']; entries=[]
 for i in range(1,q['v041_primary']+1):
  blocked=i==q['blocked_v041_index']; entries.append({'decision_id':f'V040-SYN-V041-{i:03d}','source_locator':f'plan:synthetic_v041_{i:03d}','primary_version':'V0.4.1','implementation_owner':'p3a_c','requirement_kind':'system_projection','projected_counts':{'rule_fragment':0 if blocked else 1,'binding':0 if blocked else 2,'key':0 if blocked else 1,'candidate':0 if blocked else 2},'scope_topology':'unmodeled' if blocked else 'single_core_probe','conflict_approval_assumption':'unmodeled' if blocked else 'none','derivation_formula':'synthetic contract fixture only; not a production decision mapping','confidence':'low','zero_load_reason':'none','blocked_projection':blocked})
 for ver,n,label in [('V0.4.2',q['v042_protected'],'V042'),('V0.4.3',q['v043_protected'],'V043')]:
  for i in range(1,n+1): entries.append({'decision_id':f'V040-SYN-{label}-{i:03d}','source_locator':f'plan:synthetic_{label.lower()}_{i:03d}','primary_version':ver,'implementation_owner':'future_primary','requirement_kind':'protected_boundary','projected_counts':{'rule_fragment':0,'binding':0,'key':0,'candidate':0},'scope_topology':'unmodeled','conflict_approval_assumption':'unmodeled','derivation_formula':'synthetic protected-boundary fixture only; zero benchmark load','confidence':'low','zero_load_reason':'protected_boundary','blocked_projection':False})
 if q.get('duplicate_decision_id'): entries[1]['decision_id']=entries[0]['decision_id']
 d={k:copy.deepcopy(v) for k,v in r.items() if k not in {'fixture_kind','synthetic_recipe'}}; d['entries']=entries; d['envelope_digest']='sha256:'+'0'*64; return stamp(d,'envelope_digest')
def coverage(v,scenario='disjoint'):
 d=copy.deepcopy(v); d['parameters'].update({'scenario_id':scenario,'measurement_kind':'coverage','permutation_seed':None}); r=copy.deepcopy(d['runs'][0]); r.update({'run_index':1,'run_kind':'coverage','canonical_determinism':'not_applicable','fingerprint_determinism':'not_applicable'}); d.update({'runs':[r],'summary':None,'execution_status':'completed','overall_gate':'go','stop_reasons':[]})
 if scenario=='dense-crossing': r.update({'terminal_state':'unresolvable','terminal_trace':{'pre_approval':'unresolvable','post_approval':'not_applicable'},'final_profile_present':False,'final_profile_fingerprint':None}); d['composer_terminal_state']='unresolvable'
 elif scenario=='mixed-conflict-approval': r['terminal_trace']={'pre_approval':'awaiting_approval','post_approval':'final'}; r['timings']['approval_generation']={'status':'measured','wall_seconds':0.03}; d['composer_terminal_state']='final'
 else: r['terminal_trace']={'pre_approval':'final','post_approval':'not_applicable'}; d['composer_terminal_state']='final'
 return stamp(d)
def stop(v,reason):
 d=copy.deepcopy(v); r=copy.deepcopy(d['runs'][0]); d.update({'runs':[r],'execution_status':'stopped','composer_terminal_state':'not_reached','overall_gate':'stop','summary':None})
 if reason in ('timeout','process_crash'):
  r.update({'run_status':reason,'input_json_bytes':None,'output_json_bytes':None,'metrics':None,'terminal_state':'not_reached','terminal_trace':{'pre_approval':'not_reached','post_approval':'not_reached'},'coverage_conservation':'not_reached','stable_id_status':'not_reached','canonical_determinism':'not_reached','fingerprint_determinism':'not_reached','final_profile_present':False,'final_profile_fingerprint':None})
 elif reason=='rss_unavailable': r['rss']={'status':'unavailable'}; d['composer_terminal_state']='final'
 else: d['parameters']['measurement_kind']='coverage'; r['run_kind']='coverage'; r['contract_status']='error'; d['composer_terminal_state']='final'
 d['stop_reasons']=b.derive_stop_reasons(d); return stamp(d)
def stale(v):
 d=copy.deepcopy(v); d['subject_manifest'][1]['sha256']='sha256:'+'9'*64; obs=b.recompute_subject_digest(d['subject_manifest']); d['subject_digest_status']={'state':'stale','observed_subject_digest':obs,'revalidation_required':True}; d.update({'runs':[],'summary':None,'execution_status':'stopped','composer_terminal_state':'not_reached','overall_gate':'stop'}); d['stop_reasons']=b.derive_stop_reasons(d); return stamp(d)

class C2(unittest.TestCase):
 def chk(self,f,checks):
  self.assertEqual(R[f],len(checks)); ids=[f'T412-C2A-{f}-{i:03d}' for i in range(1,len(checks)+1)]; self.assertEqual(len(ids),len(set(ids)))
  for aid,ok in zip(ids,checks): self.assertTrue(ok,aid)
 def test_state(self):
  v=load('benchmark-result.valid.json'); ts=stop(v,'timeout'); cr=stop(v,'process_crash'); rs=stop(v,'rss_unavailable'); st=stale(v); ce=stop(v,'contract_error'); th=copy.deepcopy(v); th['runs'][3]['timings']['end_to_end']['wall_seconds']=61.; th['summary']['max_wall_seconds']=61.; th['overall_gate']='stop'; th['stop_reasons']=b.derive_stop_reasons(th); stamp(th); tm=coverage(v); tm['runs'][0].update({'terminal_state':'unresolvable','terminal_trace':{'pre_approval':'unresolvable','post_approval':'not_applicable'},'final_profile_present':False,'final_profile_fingerprint':None}); tm['composer_terminal_state']='unresolvable'; tm['overall_gate']='stop'; tm['stop_reasons']=b.derive_stop_reasons(tm); stamp(tm); dense=coverage(v,'dense-crossing'); nr=copy.deepcopy(ts); nr['stop_reasons']=[]; stamp(nr); gwr=load('benchmark-result.invalid.json')
  self.chk('STATE',[schema('RESULT',v) and sem(b.validate_benchmark_result_semantics,v),schema('RESULT',ts) and sem(b.validate_benchmark_result_semantics,ts),sem(b.validate_benchmark_result_semantics,cr),sem(b.validate_benchmark_result_semantics,rs),sem(b.validate_benchmark_result_semantics,st),sem(b.validate_benchmark_result_semantics,ce),sem(b.validate_benchmark_result_semantics,th) and th['execution_status']=='completed',sem(b.validate_benchmark_result_semantics,tm),sem(b.validate_benchmark_result_semantics,dense) and dense['overall_gate']=='go',not sem(b.validate_benchmark_result_semantics,nr),schema('RESULT',gwr) and not sem(b.validate_benchmark_result_semantics,gwr),set(b.STOP_REASON_ORDER)>={'timeout','process_crash','rss_unavailable','subject_stale','contract_error','terminal_state_mismatch','threshold_exceeded'}])
 def test_matrix(self):
  v=load('benchmark-config.valid.json'); four=load('benchmark-config.invalid.json'); dup=copy.deepcopy(v); dup['generation']['permutation_seeds'][-1]=dup['generation']['permutation_seeds'][0]; stamp(dup,'config_digest'); wp=copy.deepcopy(v); wp['matrices']['performance_cells'][0]={'scale_id':'0.5x','scenario_id':'disjoint'}; stamp(wp,'config_digest'); wd=copy.deepcopy(v); wd['matrices']['determinism_cells'][0]['permutation_seed']=999; stamp(wd,'config_digest'); wr=copy.deepcopy(v); wr['repetitions']['determinism_runs_per_seed']=2; stamp(wr,'config_digest')
  self.chk('MATRIX',[schema('CONFIG',v) and sem(b.validate_benchmark_config_semantics,v),tuple(x['scale_id'] for x in v['scales'])==b.SCALES,tuple(x['scenario_id'] for x in v['scenarios'])==b.SCENARIOS,len(v['matrices']['coverage_cells'])==16,len(v['matrices']['performance_cells'])==10,len(v['generation']['permutation_seeds'])>=5,len(v['matrices']['determinism_cells'])==40,{x['scale_id'] for x in v['matrices']['determinism_cells']}=={'1.5x','2.0x'},not schema('CONFIG',four),not schema('CONFIG',dup),schema('CONFIG',wp) and not sem(b.validate_benchmark_config_semantics,wp) and schema('CONFIG',wd) and not sem(b.validate_benchmark_config_semantics,wd),not schema('CONFIG',wr)])
 def test_timing(self):
  v=load('benchmark-result.valid.json'); t=v['runs'][0]['timings']; miss=copy.deepcopy(v); del miss['runs'][0]['timings']['compose']; stamp(miss); bad=copy.deepcopy(v); bad['runs'][0]['timings']['compose']={'status':'measured'}; stamp(bad); e2e=copy.deepcopy(v); e2e['runs'][0]['timings']['end_to_end']={'status':'not_applicable'}; stamp(e2e)
  self.chk('TIMING',[tuple(t)==b.TIMING_STAGES,len(t)==7,all(x in t for x in b.TIMING_STAGES),t['approval_generation']=={'status':'not_applicable'},all('wall_seconds' in x for x in t.values() if x['status']=='measured'),not schema('RESULT',miss),not schema('RESULT',bad),schema('RESULT',e2e) and not sem(b.validate_benchmark_result_semantics,e2e)])
 def test_rss(self):
  c=load('benchmark-config.valid.json'); v=load('benchmark-result.valid.json'); r=v['runs'][1]['rss']; delta=copy.deepcopy(v); delta['runs'][1]['rss']['delta_peak_rss_mib']=999; stamp(delta); un=stop(v,'rss_unavailable'); go=copy.deepcopy(un); go['overall_gate']='go'; go['stop_reasons']=[]; stamp(go); bp=copy.deepcopy(c); bp['rss_protocol']['sampling_interval_seconds']=1; stamp(bp,'config_digest')
  self.chk('RSS',[c['rss_protocol']['measurement_method']=='external_supervisor_child_peak_working_set',c['rss_protocol']['sampling_interval_seconds']==.05,c['rss_protocol']['child_process_scope']=='benchmark_child_process_only',c['rss_protocol']['record_baseline'] and c['rss_protocol']['record_delta'],c['rss_protocol']['unavailable_policy']=='fail_closed',set(r)=={'status','baseline_rss_mib','peak_rss_mib','delta_peak_rss_mib'},math.isclose(r['delta_peak_rss_mib'],r['peak_rss_mib']-r['baseline_rss_mib']),schema('RESULT',delta) and not sem(b.validate_benchmark_result_semantics,delta),sem(b.validate_benchmark_result_semantics,un) and not sem(b.validate_benchmark_result_semantics,go),not schema('CONFIG',bp)])
 def test_metrics(self):
  v=load('benchmark-result.valid.json'); m=v['runs'][1]['metrics']; miss=copy.deepcopy(v); del miss['runs'][1]['metrics']['candidate_count']; stamp(miss); fl=copy.deepcopy(v); fl['runs'][1]['metrics']['candidate_count']=1.5; stamp(fl); tree=ast.parse(COMP.read_text()); fields=next(tuple(x.target.id for x in n.body if isinstance(x,ast.AnnAssign)) for n in tree.body if isinstance(n,ast.ClassDef) and n.name=='IntentCompositionMetrics')
  self.chk('METRICS',[tuple(m)==b.C1_METRIC_FIELDS,fields==b.C1_METRIC_FIELDS,all(isinstance(x,int) and not isinstance(x,bool) for x in m.values()),b.canonical_json_bytes(m)==b.canonical_json_bytes(copy.deepcopy(m)),not schema('RESULT',miss),not schema('RESULT',fl),'metrics' not in SCHEMAS['CONFIG']['properties'] and 'metrics' not in SCHEMAS['ENVELOPE']['properties'],'semantic_fingerprint' not in SCHEMAS['RESULT']['properties']])
 def test_envelope(self):
  v=env_fixture('projected-envelope.valid.json'); inv=env_fixture('projected-envelope.invalid.json'); summ=copy.deepcopy(v); summ['decision_population_summary']['v041_primary']=149; stamp(summ,'envelope_digest'); blk=copy.deepcopy(v); blk['entries'][149]['projected_counts']['candidate']=1; stamp(blk,'envelope_digest'); changed=copy.deepcopy(v); changed['entries'][0]['confidence']='medium'; prot=[x for x in v['entries'] if x['primary_version']!='V0.4.1']; counts={z:sum(x['primary_version']==z for x in v['entries']) for z in ('V0.4.1','V0.4.2','V0.4.3')}
  self.chk('ENVELOPE',[schema('ENVELOPE',v) and sem(b.validate_projected_envelope_semantics,v),len(v['entries'])==171,counts=={'V0.4.1':150,'V0.4.2':20,'V0.4.3':1},len({x['decision_id'] for x in v['entries']})==171,schema('ENVELOPE',inv) and not sem(b.validate_projected_envelope_semantics,inv),not schema('ENVELOPE',summ),schema('ENVELOPE',blk) and not sem(b.validate_projected_envelope_semantics,blk),all(x['implementation_owner']=='future_primary' for x in prot),all(all(y==0 for y in x['projected_counts'].values()) for x in prot),v['projection_kind']=='synthetic_contract_fixture' and all(x['decision_id'].startswith('V040-SYN-') for x in v['entries']),v['envelope_digest']==b.recompute_envelope_digest(v),schema('ENVELOPE',changed) and not sem(b.validate_projected_envelope_semantics,changed)])
 def test_result(self):
  v=load('benchmark-result.valid.json'); mixed=coverage(v,'mixed-conflict-approval'); dense=coverage(v,'dense-crossing'); st=stale(v); py=copy.deepcopy(v); py['environment']['python_version']='3.11.9'; stamp(py); only=copy.deepcopy(v); only['result_digest']='sha256:'+'e'*64; ch=copy.deepcopy(v); ch['runs'][1]['output_json_bytes']+=1; reorder=json.loads(json.dumps(v,sort_keys=True))
  def nonfinite():
   for x in (float('nan'),float('inf'),float('-inf')):
    try: b.canonical_json_bytes({'x':x}); return False
    except b.BenchmarkContractError: pass
   return True
  self.chk('RESULT',[v['result_digest']==b.recompute_result_digest(v),b.recompute_result_digest(v)==b.recompute_result_digest(only),b.recompute_result_digest(v)!=b.recompute_result_digest(ch),b.recompute_result_digest(v)==b.recompute_result_digest(reorder),not sem(b.validate_benchmark_result_semantics,only),not schema('RESULT',py),v['benchmark_subject_digest']==b.recompute_subject_digest(v['subject_manifest']),sem(b.validate_benchmark_result_semantics,st) and st['stop_reasons']==['subject_stale'],mixed['runs'][0]['terminal_trace']=={'pre_approval':'awaiting_approval','post_approval':'final'} and sem(b.validate_benchmark_result_semantics,mixed),dense['composer_terminal_state']=='unresolvable' and dense['overall_gate']=='go' and sem(b.validate_benchmark_result_semantics,dense),'evidence_commit' not in SCHEMAS['RESULT']['properties'] and 'evidence_commit' not in v,v['output_json_bytes_basis'].startswith('canonical_composition_report') and v['input_json_bytes_basis']=='canonical_benchmark_input_json_v1',not sem(b.validate_benchmark_result_semantics,ch),nonfinite()])
 def test_boundary(self):
  forbidden={'artifact_kind','semantic_fingerprint','input_fingerprints','delivery','final_ready','runtime_eligible','execution_eligibility'}; roots=set().union(*(set(x['properties']) for x in SCHEMAS.values())); refs=[]
  def walk(x):
   if isinstance(x,dict):
    for k,v in x.items():
     if k=='$ref': refs.append(v)
     walk(v)
   elif isinstance(x,list):
    for v in x: walk(v)
  for s in SCHEMAS.values(): walk(s)
  src=MOD.read_text(); fixture='\n'.join(p.read_text() for p in FIX.glob('*.json'))
  self.chk('BOUNDARY',[len(SCHEMAS)==3,all(not Draft202012Validator.check_schema(s) for s in SCHEMAS.values()),all(r.startswith('#') for r in refs),not(forbidden&roots),all(k not in SCHEMAS['RESULT']['properties'] for k in forbidden),'subprocess' not in src and 'run_benchmark' not in src and 'def main(' not in src,FIX.name=='p3a_c2' and not(FIX.parent/'p3a_c2a').exists(),'hostname' not in fixture and '/home/' not in fixture and '/Users/' not in fixture,'runtime_eligible' not in roots,'evidence_commit' not in roots])
if __name__=='__main__': unittest.main()
