"""Read-only release inclusion checks; file hashes are not publisher authentication."""
from pathlib import Path
import re, zipfile, xml.etree.ElementTree as ET
from pyopenvba import ExcelFile
from distribution_notice import audit_distribution
from distribution_source_profile import verify_enhanced_source
from distribution_gate import MANAGEMENT_CONTROLS

def audit_release(path: Path, profile: str, host_hashes=None):
    if profile not in ('internal-xlam','enhanced-dll'):raise ValueError('Unknown profile')
    result=audit_distribution(path)
    failures=[]
    if not result['zip_integrity_ok'] or not result['vba_password_present']:failures.append('Invalid or unprotected XLAM')
    if result['notified_module_count'] != result['module_count'] or not result['custom_notice_present'] or not result['core_notice_present']:failures.append('Missing usage notice')
    if result['web_extension_count']:failures.append('Unexpected web extension')
    with zipfile.ZipFile(path) as archive:
        names=archive.namelist()
        props={p.get('name'): ''.join(p.itertext()) for p in ET.fromstring(archive.read('docProps/custom.xml'))}
        for ribbon in [n for n in names if n.startswith('customUI/') and n.endswith('.xml')]:
            for control in ET.fromstring(archive.read(ribbon)).iter():
                if control.tag.rsplit('}',1)[-1] not in {'button','toggleButton','menu','dynamicMenu','splitButton','gallery','dropDown','comboBox','checkBox'}:continue
                if control.get('id') in MANAGEMENT_CONTROLS:continue
                if control.get('getEnabled') not in {'NxRibbonGetEnabled','NxRibbonGetDistributionEnabled'}:
                    failures.append('Ungated distribution control: '+str(control.attrib))
        if len(names)!=len(set(n.lower() for n in names)):failures.append('Duplicate package member')
        for name in names:
            if name.startswith('/') or '..' in name.split('/') or '\\' in name:failures.append('Unsafe package member')
            if re.search(r'(^|/)(?:tests?|tools|src|contracts)/|\.(?:pdb|py|ps1|cs|bas|cls)$',name,re.I):failures.append('Development file in XLAM: '+name)
            if name.endswith('.rels'):
                for relationship in ET.fromstring(archive.read(name)):
                    if relationship.get('TargetMode')=='External':failures.append('External relationship: '+name)
    with ExcelFile(path) as book:
        gate=book.get_module('NxDistributionGate')
        if 'NX_DISTRIBUTION_ENFORCE As Boolean = True' not in gate:failures.append('Expiry enforcement disabled')
        policy={}
        for field,key in [('PAYLOAD','payload'),('PUBLIC','public_blob_hex'),('SIGNATURE','signature_hex')]:
            match=re.search(r'NX_POLICY_'+field+r' As String = "([^"]+)"',gate)
            if match:policy[key]=match.group(1)
        if len(policy)!=3:
            failures.append('Signed expiry policy missing')
        else:
            from datetime import date, timedelta
            from distribution_gate import verify_policy_signature
            fields=policy['payload'].split('|')
            if len(fields)!=4 or fields[0]!='NXL1' or fields[2]!=profile:failures.append('Invalid signed policy identity')
            elif (props.get('DistributionBlockFrom')!=fields[3] or props.get('DistributionProfile')!=profile
                  or props.get('DistributionValidUntil')!=(date.fromisoformat(fields[3])-timedelta(days=1)).isoformat()
                  or props.get('DistributionVersion')!='내엑셀 '+fields[1]+'_배포본'):
                failures.append('Policy metadata differs from signed payload')
            if not verify_policy_signature(policy):failures.append('Invalid expiry signature')
        result['expiry_signature_verified']=not any('policy' in f.lower() or 'expiry' in f.lower() or 'metadata' in f.lower() for f in failures)
        for name in book.module_names():
            if name.startswith(('T_','Test_')) or name=='NxTestHarness':failures.append('Test module: '+name)
        if profile=='enhanced-dll':
            verify_enhanced_source(book.get_module('NxWorkbookCompare'))
            guard=book.get_module('NxHostIntegrity')
            if not re.search(r'NX_HOST_VERIFY_FILES\s+As Boolean\s*=\s*True',guard,re.I):failures.append('Runtime DLL validation disabled')
            for arch in ('x86','x64'):
                digest=(host_hashes or {}).get(arch,'')
                if len(digest)!=64 or digest not in guard:failures.append('Runtime DLL hash mismatch: '+arch)
            if 'resultJson = NxPowerShellRunHwpx' in book.get_module('NxHwpxController'):failures.append('Enhanced HWPX fallback still present')
    result.update(profile=profile,failures=failures,status='PASS' if not failures else 'FAIL',publisher_authentication='NOT_RUN')
    if failures:raise ValueError('; '.join(failures))
    return result
