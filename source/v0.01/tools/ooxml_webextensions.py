"""Remove only the four reviewed, hidden inherited Office Store extensions."""
from __future__ import annotations
import os
import posixpath
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

KNOWN_STORE_IDS = {'wa200003696', 'wa200005271', 'wa200005502', 'wa200009190'}
BASE = 'xl/webextensions/'
REL_TYPE = 'http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes'

def clean_members(members: dict[str, bytes]) -> tuple[dict[str, bytes], list[str]]:
    if any('vbaprojectsignature' in n.lower() or n.lower().startswith('_xmlsignatures/') for n in members):
        raise ValueError('signed OOXML input must not be rewritten')
    removed = sorted(n for n in members if n.startswith(BASE))
    if not removed:
        return dict(members), []
    expected = {BASE+'taskpanes.xml', BASE+'_rels/taskpanes.xml.rels'}
    expected.update(BASE+f'webextension{i}.xml' for i in range(1, 5))
    if set(removed) != expected:
        raise ValueError('unexpected webextension parts; manual review required')
    panes = ET.fromstring(members[BASE+'taskpanes.xml'])
    for pane in panes:
        if pane.get('visibility') != '0':
            raise ValueError('visible webextension pane requires review')
    ids = []
    for i in range(1, 5):
        extension = ET.fromstring(members[BASE+f'webextension{i}.xml'])
        refs = [e for e in extension.iter() if e.tag.rsplit('}', 1)[-1] == 'reference']
        identities = {ref.get('id', '').lower() for ref in refs}
        if len(identities) != 1 or not identities.issubset(KNOWN_STORE_IDS):
            raise ValueError('unknown Office Store extension requires review')
        ids.append(refs[0].get('id').lower())
    if set(ids) != KNOWN_STORE_IDS:
        raise ValueError('unexpected duplicate Office Store identity')
    result = {n: b for n, b in members.items() if n not in expected}
    for name, raw in list(result.items()):
        if not name.endswith('.rels'):
            continue
        root = ET.fromstring(raw)
        changed = False
        origin = '' if name == '_rels/.rels' else posixpath.dirname(posixpath.dirname(name))
        for rel in list(root):
            if rel.get('TargetMode') == 'External':
                continue
            target = rel.get('Target', '')
            resolved = posixpath.normpath(posixpath.join(origin, target)).lstrip('/')
            if resolved in expected:
                if rel.get('Type') != REL_TYPE or resolved != BASE+'taskpanes.xml':
                    raise ValueError('unexpected inbound webextension relationship')
                root.remove(rel)
                changed = True
        if changed:
            result[name] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    types = ET.fromstring(result['[Content_Types].xml'])
    for child in list(types):
        if child.get('PartName', '').lstrip('/') in expected:
            types.remove(child)
    result['[Content_Types].xml'] = ET.tostring(types, encoding='utf-8', xml_declaration=True)
    return result, removed

def clean_file(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        if len(archive.namelist()) != len(set(archive.namelist())):
            raise ValueError('duplicate ZIP entry rejected')
        original = {n: archive.read(n) for n in archive.namelist()}
    cleaned, removed = clean_members(original)
    if not removed:
        return []
    fd, temporary = tempfile.mkstemp(prefix=path.name+'.', dir=path.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as archive:
            for name, data in cleaned.items():
                archive.writestr(name, data)
        with zipfile.ZipFile(temporary) as archive:
            if archive.testzip() is not None:
                raise ValueError('rewritten ZIP validation failed')
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)
    return removed
