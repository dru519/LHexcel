import ctypes
import sys
import unittest
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / 'tools'))
from distribution_gate import gate_ribbon, gate_source, signed_policy, verify_policy_signature, MANAGEMENT_CONTROLS
from distribution_notice import NOTICE_TEXT, LEGACY_NOTICE_TEXT, with_vba_notice, without_vba_notice, _block


def cryptoapi_verify(policy):
    api = ctypes.WinDLL('advapi32', use_last_error=True)
    pointer, dword = ctypes.c_size_t, ctypes.c_ulong
    signatures = {
        'CryptAcquireContextW': [ctypes.POINTER(pointer), ctypes.c_void_p, ctypes.c_void_p, dword, dword],
        'CryptImportKey': [pointer, ctypes.c_void_p, dword, pointer, dword, ctypes.POINTER(pointer)],
        'CryptCreateHash': [pointer, dword, pointer, dword, ctypes.POINTER(pointer)],
        'CryptHashData': [pointer, ctypes.c_void_p, dword, dword],
        'CryptVerifySignatureW': [pointer, ctypes.c_void_p, dword, pointer, ctypes.c_void_p, dword],
        'CryptDestroyHash': [pointer], 'CryptDestroyKey': [pointer],
        'CryptReleaseContext': [pointer, dword],
    }
    for name, args in signatures.items():
        getattr(api, name).argtypes = args
        getattr(api, name).restype = ctypes.c_int
    provider, key, hash_handle = pointer(), pointer(), pointer()
    blob = bytes.fromhex(policy['public_blob_hex'])
    signature = bytes.fromhex(policy['signature_hex'])
    payload = policy['payload'].encode('ascii')
    try:
        if not api.CryptAcquireContextW(ctypes.byref(provider), None, None, 24, 0xf0000000): return False
        if not api.CryptImportKey(provider, blob, len(blob), 0, 0, ctypes.byref(key)): return False
        if not api.CryptCreateHash(provider, 0x800c, 0, 0, ctypes.byref(hash_handle)): return False
        if not api.CryptHashData(hash_handle, payload, len(payload), 0): return False
        return bool(api.CryptVerifySignatureW(hash_handle, signature, len(signature), key, None, 0))
    finally:
        if hash_handle.value: api.CryptDestroyHash(hash_handle)
        if key.value: api.CryptDestroyKey(key)
        if provider.value: api.CryptReleaseContext(provider, 0)


class DistributionGateTests(unittest.TestCase):
    def test_every_shipped_control_has_gate_and_only_management_is_exempt(self):
        root = ET.fromstring(gate_ribbon((ROOT/'src/ribbon/customUI14.xml').read_bytes()))
        kinds = {'button','toggleButton','menu','dynamicMenu','splitButton','gallery','dropDown','comboBox','checkBox'}
        for e in root.iter():
            if e.tag.rsplit('}',1)[-1] in kinds:
                self.assertTrue(e.get('getEnabled') or e.get('id') in MANAGEMENT_CONTROLS, e.attrib)

    def test_ribbon_transformation_is_idempotent(self):
        source = (ROOT/'src/ribbon/customUI14.xml').read_bytes()
        self.assertEqual(gate_ribbon(source), gate_ribbon(gate_ribbon(source)))

    def test_prior_notice_upgrades_without_duplicate_or_code_changes(self):
        source = 'Option Explicit\n'
        upgraded = with_vba_notice(_block('\n', LEGACY_NOTICE_TEXT) + source)
        self.assertEqual(without_vba_notice(upgraded), source)
        self.assertIn('사용기한과 보호장치의 임의 수정', upgraded)
        self.assertIn('AI나 자동 분석', NOTICE_TEXT)

    @unittest.skipUnless(sys.platform == 'win32', 'Windows CryptoAPI required')
    def test_independent_cryptoapi_rejects_changed_date_signature_and_key(self):
        policy = signed_policy('v0.01_r86', 'internal-xlam', date(2027,6,30))
        self.assertTrue(cryptoapi_verify(policy))
        self.assertTrue(verify_policy_signature(policy))
        self.assertFalse(verify_policy_signature(dict(policy, payload=policy['payload'].replace('2027','2028'))))
        self.assertFalse(cryptoapi_verify(dict(policy, payload=policy['payload'].replace('2027','2028'))))
        sig = policy['signature_hex']
        self.assertFalse(cryptoapi_verify(dict(policy, signature_hex=('00' if sig[:2]!='00' else '01')+sig[2:])))
        blob = policy['public_blob_hex']
        self.assertFalse(cryptoapi_verify(dict(policy, public_blob_hex=blob[:-2]+('00' if blob[-2:]!='00' else '01'))))
        source = gate_source((ROOT/'src/vba/core/NxDistributionGate.bas').read_text(encoding='utf-8'), policy)
        self.assertIn('NX_DISTRIBUTION_ENFORCE As Boolean = True', source)
        self.assertLessEqual(max(map(len, source.splitlines())), 1023, 'VBA physical line limit')
        with self.assertRaises(ValueError): gate_source(source, policy)

if __name__ == '__main__':
    unittest.main()
