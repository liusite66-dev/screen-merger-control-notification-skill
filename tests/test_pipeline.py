import json, subprocess, sys
from pathlib import Path

ROOT=Path(__file__).parents[1]
def test_prepare_report(tmp_path):
 tx=tmp_path/'tx.json'; tx.write_text(json.dumps({'title':'甲乙股权收购','transaction_type':'取得控制权','closing_status':'已签约未交割','operators':[{'operator_id':'A','name':'甲公司'}]}),encoding='utf-8')
 ws=tmp_path/'ws'; out=tmp_path/'out'
 r=subprocess.run([sys.executable,str(ROOT/'scripts/run_pipeline.py'),'prepare','--transaction',str(tx),'--workspace',str(ws),'--processing-environment','local'],capture_output=True,text=True); assert r.returncode==0, r.stderr
 analysis={'operators':[{'operator_id':'A','name':'甲公司','role':'收购方','source_locator':'tx.json'}],'control_findings':[],'turnover_records':[],'legal_bases':[],'procedure_assessment':[{'item_id':'P1','route':'待核验','status':'待核验','reason':'缺少营业额','source_locator':'tx.json'}],'material_checklist':[],'review_items':[{'item_id':'R1','topic':'营业额','risk_level':'待核验','finding':'资料不足','recommendation':'补充审计口径营业额','source_locator':'tx.json'}],'processing_records':[]}
 ap=tmp_path/'a.json'; ap.write_text(json.dumps(analysis,ensure_ascii=False),encoding='utf-8')
 r=subprocess.run([sys.executable,str(ROOT/'scripts/run_pipeline.py'),'report','--bundle',str(ws/'merger-bundle.json'),'--analysis-json',str(ap),'--output-dir',str(out),'--cleanup'],capture_output=True,text=True); assert r.returncode==0, r.stderr
 assert (out/'甲乙股权收购_经营者集中申报审查.xlsx').exists(); assert (out/'甲乙股权收购_经营者集中申报审查.html').exists(); assert not ws.exists()
