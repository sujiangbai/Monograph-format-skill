import copy, hashlib, importlib.util, json, re, unittest
from decimal import Decimal
from pathlib import Path
from statistics import median
from jsonschema import Draft202012Validator
ROOT=Path(__file__).resolve().parents[1]; B=ROOT/'format-monograph/references/benchmarks/v0412/p3a-c2'; F=ROOT/'tests/fixtures/v0412/p3a_c2'; M=ROOT/'format-monograph/scripts/profile_v2_benchmark.py'
spec=importlib.util.spec_from_file_location('b',M); b=importlib.util.module_from_spec(spec); spec.loader.exec_module(b)
S={k:json.loads((B/f).read_text()) for k,f in {'C':'benchmark-config.schema.json','R':'benchmark-result.schema.json','E':'projected-envelope.schema.json'}.items()}; V={k:Draft202012Validator(v) for k,v in S.items()}
XR=tuple(f'T412-C2A-XR-{i:03d}' for i in range(1,111)); PM=tuple(f'T412-C2A-PM2-{i:03d}' for i in range(1,37)); IR=tuple(f'T412-C2A-IR3-{i:03d}' for i in range(1,49))
def load(n): return json.loads((F/n).read_text())
def schema(k,d): return not list(V[k].iter_errors(d))
def ok(fn,*a):
 try: fn(*a); return True
 except b.BenchmarkContractError: return False
def dig(d,k):
 x=copy.deepcopy(d); x.pop(k,None); return 'sha256:'+hashlib.sha256(json.dumps(x,sort_keys=True,separators=(',',':'),ensure_ascii=False,allow_nan=False).encode()).hexdigest()
def stamp(d,k): d[k]=dig(d,k); return d
def settle(d): d['stop_reasons']=b.derive_stop_reasons(d); d['overall_gate']='stop' if d['stop_reasons'] else 'go'; return stamp(d,'result_digest')
def envelope():
 r=load('projected-envelope.valid.json'); es=[]
 for i in range(1,151):
  z=i==150; es.append({'decision_id':f'V040-SYN-V041-{i:03d}','source_locator':f'plan:synthetic_v041_{i:03d}','primary_version':'V0.4.1','implementation_owner':'p3a_c','requirement_kind':'system_projection','projected_counts':{'rule_fragment':0 if z else 1,'binding':0 if z else 2,'key':0 if z else 1,'candidate':0 if z else 2},'scope_topology':'unmodeled' if z else 'single_core_probe','conflict_approval_assumption':'unmodeled' if z else 'none','derivation_formula':'synthetic contract fixture only; not a production decision mapping','confidence':'low','zero_load_reason':'none','blocked_projection':z})
 for ver,n,label in [('V0.4.2',20,'V042'),('V0.4.3',1,'V043')]:
  for i in range(1,n+1): es.append({'decision_id':f'V040-SYN-{label}-{i:03d}','source_locator':f'plan:synthetic_{label.lower()}_{i:03d}','primary_version':ver,'implementation_owner':'future_primary','requirement_kind':'protected_boundary','projected_counts':{'rule_fragment':0,'binding':0,'key':0,'candidate':0},'scope_topology':'unmodeled','conflict_approval_assumption':'unmodeled','derivation_formula':'synthetic protected-boundary fixture only; zero benchmark load','confidence':'low','zero_load_reason':'protected_boundary','blocked_projection':False})
 d={k:v for k,v in r.items() if k not in {'fixture_kind','synthetic_recipe'}}; d['entries']=es; return stamp(d,'envelope_digest')
def scenario_run(r,s):
 e=b.FROZEN_SCENARIO_SEMANTICS[s]; final=e['terminal_state']=='final'; r['terminal_state']=e['terminal_state']; r['terminal_trace']={'pre_approval':e['pre_approval_terminal'],'post_approval':e['post_approval_terminal']}; r['final_profile_present']=final; r['final_profile_fingerprint']='sha256:'+'d'*64 if final else None; r['timings']['approval_generation']={'status':'measured','wall_seconds':.03} if s=='mixed-conflict-approval' else {'status':'not_applicable'}; r['timings']['apply']={'status':'measured','wall_seconds':.04} if final else {'status':'not_applicable'}
def perf(scale='1.5x',scenario='disjoint'):
 d=load('benchmark-result.valid.json'); d['parameters'].update({'scale_id':scale,'scenario_id':scenario,'measurement_kind':'performance','permutation_seed':None}); d['composer_terminal_state']=b.FROZEN_SCENARIO_SEMANTICS[scenario]['terminal_state']
 for i,r in enumerate(d['runs']): r['run_index']=i+1; r['run_kind']='performance_warmup' if i==0 else 'performance_measured'; r['canonical_determinism']=r['fingerprint_determinism']='not_applicable'; r['coverage_conservation']='passed'; r['stable_id_status']='stable'; scenario_run(r,scenario)
 ms=d['runs'][1:]; d['summary']={'median_wall_seconds':median([x['timings']['end_to_end']['wall_seconds'] for x in ms]),'max_wall_seconds':max(x['timings']['end_to_end']['wall_seconds'] for x in ms),'median_peak_rss_mib':median([x['rss']['peak_rss_mib'] for x in ms]),'max_peak_rss_mib':max(x['rss']['peak_rss_mib'] for x in ms),'median_output_json_bytes':int(median([x['output_json_bytes'] for x in ms])),'max_output_json_bytes':max(x['output_json_bytes'] for x in ms)}; d['ratio_evidence']={x:{'status':'not_applicable'} for x in ('wall','rss','output_json')}; return settle(d)
def coverage(scale='0.5x',scenario='disjoint'):
 d=perf('1.5x',scenario); r=copy.deepcopy(d['runs'][0]); r['run_index']=1; r['run_kind']='coverage'; d['runs']=[r]; d['summary']=None; d['parameters'].update({'scale_id':scale,'measurement_kind':'coverage'}); return settle(d)
def determinism(scale='1.5x',seed=7,scenario='disjoint'):
 d=coverage(scale,scenario); r=d['runs'][0]; r['run_kind']='determinism'; r['canonical_determinism']=r['fingerprint_determinism']='matched'; d['parameters'].update({'measurement_kind':'determinism','permutation_seed':seed}); return settle(d)
def pair():
 one=perf('1.0x','mixed-conflict-approval'); two=perf('2.0x','mixed-conflict-approval')
 for m,f in b.RATIO_SUMMARY_FIELDS.items():
  x=one['summary'][f]; y=two['summary'][f]; two['ratio_evidence'][m]={'status':'measured','baseline_1x':x,'observed_2x':y,'ratio':float((Decimal(str(y))/Decimal(str(x))).quantize(Decimal('0.000001')))}
 return one,settle(two)
def interrupted(base,status='timeout',point='early'):
 d=copy.deepcopy(base); r=copy.deepcopy(d['runs'][0]); d['runs']=[r]; r['run_index']=1; r['run_status']=status; scenario=d['parameters']['scenario_id']; kind=d['parameters']['measurement_kind']; d['execution_status']='stopped'; d['summary']=None; d['ratio_evidence']={x:{'status':'not_applicable'} for x in ('wall','rss','output_json')}
 if point=='early':
  r['timings']['compose']={'status':'not_reached'}; r['timings']['approval_generation']={'status':'not_reached'} if scenario=='mixed-conflict-approval' else {'status':'not_applicable'}; r['timings']['apply']={'status':'not_applicable'} if scenario=='dense-crossing' else {'status':'not_reached'}; r['timings']['canonical_serialization']={'status':'not_reached'}; r['terminal_state']='not_reached'; r['terminal_trace']={'pre_approval':'not_reached','post_approval':'not_reached'}; r['final_profile_present']=False; r['final_profile_fingerprint']=None; r['metrics']=None; r['input_json_bytes']=None; r['output_json_bytes']=None; r['coverage_conservation']='not_reached'; r['stable_id_status']='not_reached'; r['canonical_determinism']=r['fingerprint_determinism']='not_reached' if kind=='determinism' else 'not_applicable'; d['composer_terminal_state']='not_reached'
 else:
  if scenario!='mixed-conflict-approval': r['timings']['approval_generation']={'status':'not_applicable'}
  if point=='terminal': r['timings']['apply']={'status':'not_reached'}; r['timings']['canonical_serialization']={'status':'not_reached'}; r['output_json_bytes']=None
  elif point=='apply': r['timings']['apply']={'status':'measured','wall_seconds':.04}; r['timings']['canonical_serialization']={'status':'not_reached'}; r['output_json_bytes']=None
  elif point=='canonical': r['timings']['apply']={'status':'measured','wall_seconds':.04}; r['timings']['canonical_serialization']={'status':'measured','wall_seconds':.02}; r['output_json_bytes']=None
  d['composer_terminal_state']=b.FROZEN_SCENARIO_SEMANTICS[scenario]['terminal_state']
 r['timings']['end_to_end']={'status':'measured','wall_seconds':30.0}; return settle(d)
def complete_suite():
 c=load('benchmark-config.valid.json'); out=[coverage(s,sc) for s in b.SCALES for sc in b.SCENARIOS]; ps={(s,sc):perf(s,sc) for s,sc in b.PERFORMANCE_CELLS}; one,two=pair(); ps[('1.0x','mixed-conflict-approval')]=one; ps[('2.0x','mixed-conflict-approval')]=two; out.extend(ps.values()); out.extend(determinism(s,seed,sc) for s in b.DETERMINISM_SCALES for sc in b.SCENARIOS for seed in c['generation']['permutation_seeds']); return out
def placeholder_count():
 src=Path(__file__).read_text(); pats=(r'\[\s*True\s*\]\s*\*\s*\d+',r'assertTrue\(\s*True\s*\)',r'\bor\s+True\b'); return sum(len(re.findall(p,src)) for p in pats)
class Base(unittest.TestCase):
 def checks(self,ids,vals):
  self.assertEqual(len(ids),len(vals));
  for i,v in zip(ids,vals): self.assertTrue(v,i)
class XRTests(Base):
 seen=set()
 def run10(self,start,vals):
  ids=XR[start-1:start+9]; self.checks(ids,vals); type(self).seen.update(ids)
 @classmethod
 def tearDownClass(cls): assert cls.seen==set(XR)
 def test_01(self):
  c=load('benchmark-config.valid.json'); v=load('benchmark-result.valid.json'); bad=copy.deepcopy(v); bad['parameters']['scale_id']='0.5x'; stamp(bad,'result_digest'); d=determinism('1.0x'); u=determinism(seed=999)
  self.run10(1,[schema('R',bad),ok(b.validate_benchmark_result_semantics,bad),not ok(b.validate_benchmark_result_against_config,bad,c),not ok(b.validate_benchmark_result_against_config,d,c),not ok(b.validate_benchmark_result_against_config,u,c),ok(b.validate_benchmark_result_against_config,v,c),v['benchmark_config_digest']==c['config_digest'],ok(b.validate_benchmark_result_against_config,coverage(),c),len(c['matrices']['coverage_cells'])==16,len(c['matrices']['performance_cells'])==10])
 def test_02(self):
  d=determinism(); d['runs'][0]['canonical_determinism']='not_applicable'; stamp(d,'result_digest'); x=coverage(); x['runs'][0]['coverage_conservation']='not_reached'; stamp(x,'result_digest'); y=coverage(); y['runs'][0]['stable_id_status']='not_reached'; stamp(y,'result_digest')
  self.run10(11,[not ok(b.validate_benchmark_result_semantics,d),not ok(b.validate_benchmark_result_semantics,x),not ok(b.validate_benchmark_result_semantics,y),determinism()['runs'][0]['canonical_determinism']=='matched',coverage()['runs'][0]['coverage_conservation']=='passed',coverage()['runs'][0]['stable_id_status']=='stable',set(b.C1_METRIC_FIELDS)==set(load('benchmark-result.valid.json')['runs'][0]['metrics']),len(b.TIMING_STAGES)==7,b.FROZEN_REPETITIONS['determinism_runs_per_seed']==1,b.FROZEN_REPETITIONS['performance_measured_runs_per_cell']==3])
 def test_03(self):
  c=load('benchmark-config.valid.json'); one,two=pair(); bad=copy.deepcopy(two); bad['ratio_evidence']['wall']['ratio']=.1; stamp(bad,'result_digest')
  self.run10(21,[ok(b.validate_ratio_evidence,one,two,c),not ok(b.validate_ratio_evidence,one,bad,c),ok(b.validate_benchmark_result_set,[one,two],c),not ok(b.validate_benchmark_result_set,[two],c),b.RATIO_SUMMARY_FIELDS['rss']=='median_peak_rss_mib',two['ratio_evidence']['wall']['status']=='measured',one['parameters']['scale_id']=='1.0x',two['parameters']['scale_id']=='2.0x',one['parameters']['scenario_id']==two['parameters']['scenario_id'],two['overall_gate']=='go'])
 def test_04(self):
  early=interrupted(perf(),'timeout','early'); ef=copy.deepcopy(early); ef['runs'][0]['final_profile_present']=True; ef['runs'][0]['final_profile_fingerprint']='sha256:'+'d'*64; stamp(ef,'result_digest'); cf=copy.deepcopy(interrupted(perf(),'process_crash','early')); cf['runs'][0]['final_profile_present']=True; cf['runs'][0]['final_profile_fingerprint']='sha256:'+'d'*64; stamp(cf,'result_digest'); later=copy.deepcopy(early); x=copy.deepcopy(perf()['runs'][0]); x['run_index']=2; later['runs'].append(x); stamp(later,'result_digest'); completed=copy.deepcopy(early); completed['execution_status']='completed'; stamp(completed,'result_digest'); late=interrupted(perf(),'timeout','terminal'); after=interrupted(perf(),'timeout','apply'); resume=copy.deepcopy(early); resume['runs'][0]['timings']['apply']={'status':'measured','wall_seconds':.1}; stamp(resume,'result_digest'); crash=interrupted(perf(),'process_crash','canonical')
  self.run10(31,[not ok(b.validate_benchmark_result_semantics,ef),not ok(b.validate_benchmark_result_semantics,cf),not ok(b.validate_benchmark_result_semantics,later),not ok(b.validate_benchmark_result_semantics,completed),ok(b.validate_benchmark_result_semantics,early) and early['overall_gate']=='stop',ok(b.validate_benchmark_result_semantics,late),late['runs'][0]['timings']['approval_generation']['status']=='not_applicable' and late['runs'][0]['timings']['apply']['status']=='not_reached',not ok(b.validate_benchmark_result_semantics,resume),after['composer_terminal_state']=='final' and ok(b.validate_benchmark_result_semantics,after),ok(b.validate_benchmark_result_semantics,crash) and crash['stop_reasons']==['process_crash']])
 def test_05(self):
  c=load('benchmark-config.valid.json'); x=coverage(); x['runs'][0]['timings']['end_to_end']['wall_seconds']=61; settle(x); d=determinism(); d['runs'][0]['rss']['peak_rss_mib']=600; d['runs'][0]['rss']['delta_peak_rss_mib']=580; settle(d)
  self.run10(41,[ok(b.validate_benchmark_result_against_config,x,c),x['overall_gate']=='stop','threshold_exceeded' in x['stop_reasons'],ok(b.validate_benchmark_result_against_config,d,c),d['overall_gate']=='stop',b.FROZEN_THRESHOLDS['scale_limits'][-1]['max_wall_seconds']==120,b.FROZEN_THRESHOLDS['scale_limits'][-1]['max_peak_rss_mib']==1024,b.ABSOLUTE_GATE_RUN_KINDS=={'coverage','performance_measured','determinism'},coverage()['overall_gate']=='go',determinism()['overall_gate']=='go'])
 def test_06(self):
  c=load('benchmark-config.valid.json'); f=copy.deepcopy(c); f['scales'][0]['factor']=2; stamp(f,'config_digest'); s=copy.deepcopy(c); s['scenarios'][2]['final_requirement']='required'; stamp(s,'config_digest')
  self.run10(51,[not ok(b.validate_benchmark_config_semantics,f),not ok(b.validate_benchmark_config_semantics,s),ok(b.validate_benchmark_config_semantics,c),b.FROZEN_SCALE_FACTORS['0.5x']==.5,b.FROZEN_SCALE_FACTORS['2.0x']==2,b.FROZEN_SCENARIO_SEMANTICS['dense-crossing']['terminal_state']=='unresolvable',b.FROZEN_SCENARIO_SEMANTICS['mixed-conflict-approval']['pre_approval_terminal']=='awaiting_approval',len(c['generation']['permutation_seeds'])>=5,c['config_digest']==dig(c,'config_digest'),c['projection_binding']['base_scale_id']=='1.0x'])
 def test_07(self):
  v=load('benchmark-result.valid.json'); t=copy.deepcopy(v); t['runs'][0]['timings']=dict(reversed(list(t['runs'][0]['timings'].items()))); stamp(t,'result_digest'); m=copy.deepcopy(v); m['runs'][0]['metrics']=dict(reversed(list(m['runs'][0]['metrics'].items()))); stamp(m,'result_digest'); miss=copy.deepcopy(v); del miss['runs'][0]['timings']['compose']; stamp(miss,'result_digest'); extra=copy.deepcopy(v); extra['runs'][0]['metrics']['extra']=1; stamp(extra,'result_digest'); sr=copy.deepcopy(v); sr['subject_manifest'][0]={'sha256':sr['subject_manifest'][0]['sha256'],'path':sr['subject_manifest'][0]['path']}; stamp(sr,'result_digest')
  self.run10(61,[ok(b.validate_benchmark_result_semantics,t),ok(b.validate_benchmark_result_semantics,m),dig(v,'result_digest')==dig(json.loads(json.dumps(v,sort_keys=True)),'result_digest'),set(v['runs'][0]['timings'])==set(b.TIMING_STAGES),set(v['runs'][0]['metrics'])==set(b.C1_METRIC_FIELDS),not schema('R',miss),not ok(b.validate_benchmark_result_semantics,miss),not schema('R',extra),not ok(b.validate_benchmark_result_semantics,extra),ok(b.validate_benchmark_result_semantics,sr)])
 def test_08(self):
  e=envelope(); renamed=copy.deepcopy(e); renamed['projection_kind']='formal_planning_projection'; stamp(renamed,'envelope_digest')
  self.run10(71,[schema('E',e),ok(b.validate_projected_envelope_semantics,e),not ok(b.validate_projected_envelope_semantics,renamed),len(e['entries'])==171,len({x['decision_id'] for x in e['entries']})==171,e['decision_population_summary']=={'v041_primary':150,'v042_protected':20,'v043_protected':1},e['compose_projection_strategy']=='single_core_probe',e['envelope_digest']==dig(e,'envelope_digest'),all(x['decision_id'].startswith('V040-SYN-') for x in e['entries']),all(x['source_locator'].startswith('plan:synthetic_') for x in e['entries'])])
 def test_09(self):
  v=load('benchmark-result.valid.json'); reordered=[{'sha256':x['sha256'],'path':x['path']} for x in v['subject_manifest']]; dup=[copy.deepcopy(v['subject_manifest'][0]),copy.deepcopy(v['subject_manifest'][0])]
  self.run10(81,[b.canonical_subject_path('format-monograph/scripts/profile_v2_benchmark.py').endswith('profile_v2_benchmark.py'),not ok(b.canonical_subject_path,'a/./b'),not ok(b.canonical_subject_path,'a//b'),not ok(b.canonical_subject_path,'a/b/'),not ok(b.canonical_subject_path,'../a'),not ok(b.canonical_subject_path,'/a'),not ok(b.canonical_subject_path,'C:\\a'),v['subject_manifest']==sorted(v['subject_manifest'],key=lambda x:x['path']),b.recompute_subject_digest(v['subject_manifest'])==b.recompute_subject_digest(reordered),not ok(b.recompute_subject_digest,dup)])
 def test_10(self):
  roots=set().union(*(set(x['properties']) for x in S.values())); forbidden={'artifact_kind','semantic_fingerprint','input_fingerprints','delivery','final_ready','runtime_eligible','execution_eligibility','evidence_commit'}
  self.run10(91,[all(Draft202012Validator.check_schema(x) is None for x in S.values()),'subprocess' not in M.read_text(),'def main(' not in M.read_text(),schema('C',load('benchmark-config.valid.json')),schema('R',load('benchmark-result.valid.json')),schema('E',envelope()),len(XR)==110,XR[0].endswith('001'),XR[-1].endswith('110'),not(forbidden&roots)])
 def test_11(self):
  c=load('benchmark-config.valid.json'); v=load('benchmark-result.valid.json')
  self.run10(101,[v['rss_protocol']==c['rss_protocol'],v['thresholds']==c['thresholds'],v['output_json_bytes_basis']==c['output_json_bytes_basis'],v['input_json_bytes_basis']==c['input_json_bytes_basis'],v['parameters']['generation_seed']==c['generation']['generation_seed'],v['benchmark_config_digest']==c['config_digest'],ok(b.validate_benchmark_result_against_config,v,c),v['result_digest']==dig(v,'result_digest'),c['config_digest']==dig(c,'config_digest'),v['reference_budget']['limit_hours']==c['total_reference_budget_hours']])
class PM2Tests(Base):
 seen=set()
 def run12(self,start,vals): ids=PM[start-1:start+11]; self.checks(ids,vals); type(self).seen.update(ids)
 @classmethod
 def tearDownClass(cls): assert cls.seen==set(PM)
 def test_pm2_001(self):
  c=load('benchmark-config.valid.json'); one,two=pair(); variants=[]
  for mut in ('commit','subject','os','ram','python','command'):
   x=copy.deepcopy(two)
   if mut=='commit': x['benchmark_subject_commit']='f'*40
   elif mut=='subject': x['subject_manifest'][0]['sha256']='sha256:'+'9'*64; h=b.recompute_subject_digest(x['subject_manifest']); x['benchmark_subject_digest']=h; x['subject_digest_status']={'state':'current','observed_subject_digest':h,'revalidation_required':False}
   elif mut=='os': x['environment']['os_family']='linux'
   elif mut=='ram': x['environment']['ram_tier']='32_to_64gib'
   elif mut=='python': x['environment']['python_version']='3.12.14'
   else: x['command_template']='python -m internal_benchmark --config alternate.json'
   stamp(x,'result_digest'); variants.append(x)
  vals=[]
  for x in variants: vals.extend([ok(b.validate_benchmark_result_against_config,x,c),not ok(b.validate_ratio_evidence,one,x,c)])
  self.run12(1,vals)
 def test_pm2_002(self):
  c=load('benchmark-config.valid.json'); e=envelope(); counts=b.aggregate_projected_envelope_counts(e); changed=copy.deepcopy(e); changed['entries'][0]['projected_counts']['candidate']+=1; stamp(changed,'envelope_digest'); bad=copy.deepcopy(c); bad['projection_binding']['aggregate_counts']['candidate']+=1; stamp(bad,'config_digest'); wrong=copy.deepcopy(c); wrong['projection_binding']['projected_envelope_digest']='sha256:'+'e'*64; stamp(wrong,'config_digest')
  self.run12(13,[ok(b.validate_benchmark_config_against_envelope,c,e),counts=={'rule_fragment':149,'binding':298,'key':149,'candidate':298},not ok(b.validate_benchmark_config_against_envelope,c,changed),not ok(b.validate_benchmark_config_against_envelope,bad,e),not ok(b.validate_benchmark_config_against_envelope,wrong,e),c['projection_binding']['base_scale_id']=='1.0x',c['projection_binding']['projection_kind']=='synthetic_contract_fixture',c['projection_binding']['projected_envelope_digest']==e['envelope_digest'],c['projection_binding']['aggregate_counts']==counts,e['projection_kind']=='synthetic_contract_fixture',schema('C',c),ok(b.validate_benchmark_config_semantics,c)])
 def test_pm2_003(self):
  c=load('benchmark-config.valid.json'); e=envelope(); r=load('benchmark-result.valid.json'); ctx=b.validate_benchmark_result_context(r,c,e); bad=copy.deepcopy(e); bad['compose_projection_strategy']='other'; stamp(bad,'envelope_digest'); miss=copy.deepcopy(e); miss['unmodeled_dimensions'].remove('docx_runtime'); stamp(miss,'envelope_digest')
  self.run12(25,[ctx['compose_projection_strategy']=='single_core_probe',ctx['production_representative'] is False,ctx['revalidation_required'] is True,ctx['representativeness_scope']=='non_production_single_core_probe',set(ctx['unmodeled_dimensions'])>=b.REQUIRED_UNMODELED_DIMENSIONS,ctx['base_scale_id']=='1.0x',ctx['aggregate_counts']==c['projection_binding']['aggregate_counts'],not ok(b.validate_benchmark_result_context,r,c,bad),not ok(b.validate_benchmark_result_context,r,c,miss),'hostname' not in r['environment'],'machine_id' not in r['environment'],ok(b.validate_benchmark_result_context,r,c,e)])
class IR3Tests(Base):
 seen=set()
 def run8(self,start,vals): ids=IR[start-1:start+7]; self.checks(ids,vals); type(self).seen.update(ids)
 @classmethod
 def tearDownClass(cls): assert cls.seen==set(IR)
 def test_ir3_001_completed_applicability_and_bytes(self):
  c=load('benchmark-config.valid.json'); dense=coverage('1.5x','dense-crossing'); bad=copy.deepcopy(dense); bad['runs'][0]['timings']['apply']={'status':'not_reached'}; stamp(bad,'result_digest'); final=coverage('1.5x','disjoint'); ni=copy.deepcopy(final); ni['runs'][0]['input_json_bytes']=None; stamp(ni,'result_digest'); no=copy.deepcopy(final); no['runs'][0]['output_json_bytes']=None; stamp(no,'result_digest'); nd=determinism(); nd['runs'][0]['input_json_bytes']=None; stamp(nd,'result_digest'); np=perf(); np['runs'][1]['output_json_bytes']=None; stamp(np,'result_digest')
  self.run8(1,[schema('R',bad) and not ok(b.validate_benchmark_result_against_config,bad,c),ok(b.validate_benchmark_result_against_config,dense,c),final['runs'][0]['timings']['apply']['status']=='measured' and ok(b.validate_benchmark_result_against_config,final,c),schema('R',ni) and not ok(b.validate_benchmark_result_against_config,ni,c),schema('R',no) and not ok(b.validate_benchmark_result_against_config,no,c),schema('R',nd) and not ok(b.validate_benchmark_result_against_config,nd,c),schema('R',np) and not ok(b.validate_benchmark_result_against_config,np,c),all(isinstance(final['runs'][0][k],int) for k in ('input_json_bytes','output_json_bytes'))])
 def test_ir3_002_interrupted_progression(self):
  c=load('benchmark-config.valid.json'); early=interrupted(perf(),'timeout','early'); late=interrupted(perf(),'timeout','terminal'); after=interrupted(perf(),'timeout','apply'); crash=interrupted(perf(),'process_crash','canonical'); resume=copy.deepcopy(early); resume['runs'][0]['timings']['apply']={'status':'measured','wall_seconds':.1}; stamp(resume,'result_digest')
  self.run8(9,[schema('R',early) and ok(b.validate_benchmark_result_against_config,early,c) and early['overall_gate']=='stop',ok(b.validate_benchmark_result_against_config,late,c) and late['composer_terminal_state']=='final',ok(b.validate_benchmark_result_against_config,after,c),ok(b.validate_benchmark_result_against_config,crash,c),not ok(b.validate_benchmark_result_against_config,resume,c),late['runs'][0]['timings']['approval_generation']['status']=='not_applicable' and late['runs'][0]['timings']['apply']['status']=='not_reached',after['runs'][0]['timings']['approval_generation']['status']=='not_applicable' and after['runs'][0]['timings']['apply']['status']=='measured',crash['runs'][0]['terminal_state']=='final' and crash['stop_reasons']==['process_crash']])
 def test_ir3_003_stopped_ratio(self):
  c=load('benchmark-config.valid.json'); timeout=interrupted(perf('2.0x','mixed-conflict-approval'),'timeout','early'); crash=interrupted(perf('2.0x','mixed-conflict-approval'),'process_crash','early'); one=perf('1.0x','mixed-conflict-approval'); two=perf('2.0x','mixed-conflict-approval'); fake=copy.deepcopy(timeout); fake['ratio_evidence']['wall']={'status':'measured','baseline_1x':1,'observed_2x':2,'ratio':2}; settle(fake); good1,good2=pair()
  self.run8(17,[ok(b.validate_benchmark_result_against_config,timeout,c),all(x['status']=='not_applicable' for x in timeout['ratio_evidence'].values()),timeout['summary'] is None and timeout['stop_reasons']==['timeout'],ok(b.validate_benchmark_result_against_config,crash,c),ok(b.validate_benchmark_result_against_config,two,c),not ok(b.validate_benchmark_result_set,[one,two],c),ok(b.validate_ratio_evidence,good1,good2,c),not ok(b.validate_benchmark_result_against_config,fake,c)])
 def test_ir3_004_suite_completeness(self):
  c=load('benchmark-config.valid.json'); full=complete_suite(); miss_cov=full[1:]; miss_perf=[x for x in full if not(x['parameters']['measurement_kind']=='performance' and x['parameters']['scale_id']=='0.5x')]; target=next(x for x in full if x['parameters']['measurement_kind']=='determinism'); miss_det=[x for x in full if x is not target]; dup=full+[copy.deepcopy(full[0])]; outcome=b.validate_complete_benchmark_suite(full,c)
  self.run8(25,[not ok(b.validate_complete_benchmark_suite,[],c),not ok(b.validate_complete_benchmark_suite,miss_cov,c),not ok(b.validate_complete_benchmark_suite,miss_perf,c),not ok(b.validate_complete_benchmark_suite,miss_det,c),not ok(b.validate_complete_benchmark_suite,dup,c),outcome['structurally_complete'] is True,outcome['overall_gate']=='go',len(full)==16+10+40])
 def test_ir3_005_complete_suite_stop(self):
  c=load('benchmark-config.valid.json'); timed=complete_suite(); timed[0]=interrupted(timed[0],'timeout','early'); to=b.validate_complete_benchmark_suite(timed,c); threshold=complete_suite(); threshold[0]['runs'][0]['timings']['end_to_end']['wall_seconds']=61; settle(threshold[0]); th=b.validate_complete_benchmark_suite(threshold,c)
  self.run8(33,[to['structurally_complete'] is True,to['overall_gate']=='stop','timeout' in timed[0]['stop_reasons'],th['structurally_complete'] is True,th['overall_gate']=='stop','threshold_exceeded' in threshold[0]['stop_reasons'],ok(b.validate_benchmark_result_set,[],c),callable(b.validate_complete_benchmark_suite)])
 def test_ir3_006_integrity_and_closed_contracts(self):
  c=load('benchmark-config.valid.json'); e=envelope(); r=load('benchmark-result.valid.json'); ctx=b.validate_benchmark_result_context(r,c,e); one,two=pair(); other=copy.deepcopy(two); other['environment']['os_family']='linux'; stamp(other,'result_digest')
  self.run8(41,[placeholder_count()==0,len(XR)==110,len(PM)==36,len(IR)==48,ok(b.validate_benchmark_config_against_envelope,c,e) and ctx['production_representative'] is False and ctx['compose_projection_strategy']=='single_core_probe',not ok(b.validate_ratio_evidence,one,other,c),ok(b.validate_benchmark_result_against_config,r,c),'subprocess' not in M.read_text() and 'def main(' not in M.read_text()])
if __name__=='__main__': unittest.main()
