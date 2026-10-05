from pathlib import Path
from datetime import date, timedelta
import argparse,json,sys
ROOT=Path(__file__).resolve().parents[2]
sys.path.insert(0,str(ROOT/'tools'))
from distribution_gate import signed_policy, gate_source

parser=argparse.ArgumentParser()
parser.add_argument('--out',type=Path,required=True)
args=parser.parse_args()
args.out.mkdir(exist_ok=False)
source=(ROOT/'src/vba/core/NxDistributionGate.bas').read_text(encoding='utf-8')
active=signed_policy('v0.01_r76','internal-xlam',date.today()+timedelta(days=1))
fixtures={'active':active, 'tampered_metadata':active,
          'expires_today':signed_policy('v0.01_r76','internal-xlam',date.today()),
          'expired':signed_policy('v0.01_r76','internal-xlam',date.today()-timedelta(days=1)),
          'tampered_date':dict(active,payload=active['payload'][:-10]+'2099-12-31'),
          'tampered_signature':dict(active,signature_hex=('00' if active['signature_hex'][:2]!='00' else '01')+active['signature_hex'][2:])}
for name,policy in fixtures.items():
    (args.out/(name+'.bas')).write_text(gate_source(source,policy),encoding='utf-8')
(args.out/'fixture-info.json').write_text(json.dumps({'date':str(date.today()),'names':list(fixtures),'system_date_changed':False}),encoding='utf-8')
