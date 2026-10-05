"""Build-only signing and distribution-only VBA/Ribbon transformation.

The Windows user key container is never copied to source or delivery packages.
This detects policy edits; it is not Authenticode or protection against a
recipient replacing the verifier itself.
"""
from __future__ import annotations
import json
import hashlib
import re
import struct
import subprocess
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

MANAGEMENT_CONTROLS = {'NX-HOME-MANAGEMENT', 'NX-PROD-MANAGEMENT'}

def verify_policy_signature(policy: dict) -> bool:
    """Read-only RSA/SHA256 verification of the public CryptoAPI blob."""
    try:
        blob = bytes.fromhex(policy['public_blob_hex'])
        signature = bytes.fromhex(policy['signature_hex'])
        kind, version, reserved, algorithm, magic, bits, exponent = struct.unpack('<BBHIIII', blob[:20])
        if (kind, version, reserved, algorithm, magic) != (6, 2, 0, 0x2400, 0x31415352):
            return False
        if bits != 3072 or exponent != 65537 or len(blob) != 20 + bits // 8 or len(signature) != bits // 8:
            return False
        modulus = int.from_bytes(blob[20:], 'little')
        value = int.from_bytes(signature, 'little')
        if modulus.bit_length() != bits or value >= modulus:
            return False
        decoded = pow(value, exponent, modulus).to_bytes(bits // 8, 'big')
        digest_info = bytes.fromhex('3031300d060960864801650304020105000420') + hashlib.sha256(policy['payload'].encode('ascii')).digest()
        expected = b'\x00\x01' + b'\xff' * (bits // 8 - len(digest_info) - 3) + b'\x00' + digest_info
        return decoded == expected
    except (ValueError, KeyError, TypeError, struct.error, UnicodeError):
        return False

def signed_policy(version: str, profile: str, block_from: date) -> dict:
    payload = f'NXL1|{version}|{profile}|{block_from.isoformat()}'
    script = Path(__file__).resolve().parents[1] / 'build/Sign-DistributionPolicy.ps1'
    result = subprocess.run(
        ['powershell.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', str(script), '-Payload', payload],
        check=True, capture_output=True, text=True, encoding='utf-8')
    policy = json.loads(result.stdout.strip())
    if policy['payload'] != payload:
        raise ValueError('Signed policy payload mismatch')
    for key in ('public_blob_hex', 'signature_hex'):
        if not re.fullmatch(r'[0-9a-f]+', policy[key]) or len(policy[key]) % 2:
            raise ValueError('Invalid signature encoding')
    return policy

def gate_source(source: str, policy: dict) -> str:
    replacements = {
        'NX_DISTRIBUTION_ENFORCE As Boolean = False': 'NX_DISTRIBUTION_ENFORCE As Boolean = True',
        'NX_POLICY_PAYLOAD As String = ""': f'NX_POLICY_PAYLOAD As String = "{policy["payload"]}"',
        'NX_POLICY_PUBLIC As String = ""': f'NX_POLICY_PUBLIC As String = "{policy["public_blob_hex"]}"',
        'NX_POLICY_SIGNATURE As String = ""': f'NX_POLICY_SIGNATURE As String = "{policy["signature_hex"]}"',
    }
    for old, new in replacements.items():
        if source.count(old) != 1:
            raise ValueError('Distribution gate template changed')
        source = source.replace(old, new)
    return source

def gate_ribbon(data: bytes) -> bytes:
    # Every product control, including nested built-in Office controls, is gated.
    # Do not gate TabHome or groups: Home's clipboard and Excel controls are owned
    # by Excel, and the two management menus must remain usable after expiry.
    root = ET.fromstring(data)
    for element in root.iter():
        kind = element.tag.rsplit('}', 1)[-1]
        if kind not in {'button', 'toggleButton', 'menu', 'dynamicMenu', 'splitButton', 'gallery', 'dropDown', 'comboBox', 'checkBox'}:
            continue
        if element.get('id') in MANAGEMENT_CONTROLS:
            continue
        if not element.get('getEnabled'):
            element.attrib.pop('enabled', None)
            element.set('getEnabled', 'NxRibbonGetDistributionEnabled')
    ET.register_namespace('', root.tag.split('}')[0][1:])
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)
