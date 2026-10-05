from pathlib import Path
from datetime import date
import argparse,json,sys
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'tools'))
from distribution_gate import signed_policy,gate_source
p=argparse.ArgumentParser();p.add_argument('--out',type=Path,required=True);args=p.parse_args()
args.out.mkdir(exist_ok=False)
source=(ROOT/'src/vba/core/NxDistributionGate.bas').read_text(encoding='utf-8')
policy=signed_policy('v0.01_r86','internal-xlam',date(2028,1,1))
fixtures={name:policy for name in ['active','expires_today','expired','tampered_metadata']}
fixtures['tampered_date']=dict(policy,payload=policy['payload'][:-10]+'2099-12-31')
fixtures['tampered_signature']=dict(policy,signature_hex=('00' if policy['signature_hex'][:2]!='00' else '01')+policy['signature_hex'][2:])
for name,item in fixtures.items():
    now={'expires_today':'DateSerial(2028, 1, 1)','expired':'DateSerial(2028, 1, 2)'}.get(name,'DateSerial(2027, 12, 31)')
    text=gate_source(source,item)
    assert text.count('If Date >= mBlockFrom Then')==1
    (args.out/(name+'.bas')).write_text(text.replace('If Date >= mBlockFrom Then',f'If {now} >= mBlockFrom Then'),encoding='utf-8')
(args.out/'fixture-info.json').write_text(json.dumps({'valid_until':'2027-12-31','block_from':'2028-01-01','system_date_changed':False,'names':list(fixtures)}),encoding='utf-8')
