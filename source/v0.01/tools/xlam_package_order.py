"""Keep workbook identification records at the front of unsigned XLAM packages."""
import os
import tempfile
import zipfile
from pathlib import Path

def ordered_names(names):
    names = list(names)
    if len(names) != len(set(names)):
        raise ValueError('Duplicate ZIP entry')
    preferred = ['[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml',
                 'xl/_rels/workbook.xml.rels']
    preferred += sorted(n for n in names if n.startswith('xl/worksheets/sheet'))
    preferred += ['xl/theme/theme1.xml', 'xl/styles.xml', 'xl/sharedStrings.xml',
                  'xl/vbaProject.bin', 'docProps/core.xml', 'docProps/app.xml', 'docProps/custom.xml']
    result = [n for n in preferred if n in names]
    return result + sorted(n for n in names if n not in result)

def normalize(path):
    path = Path(path)
    with zipfile.ZipFile(path) as source:
        infos = source.infolist()
        names = [i.filename for i in infos]
        if any('signature' in n.lower() for n in names):
            raise ValueError('Signed package requires separate re-signing workflow')
        order = ordered_names(names)
        if order == names:
            return
        if source.testzip() is not None:
            raise ValueError('ZIP integrity failure')
        members = {i.filename: (i, source.read(i.filename)) for i in infos}
        comment = source.comment
    fd, temp = tempfile.mkstemp(prefix=path.name+'.order-', dir=path.parent)
    os.close(fd)
    try:
        with zipfile.ZipFile(temp, 'w') as target:
            target.comment = comment
            for name in order:
                info, data = members[name]
                target.writestr(info, data)
        with zipfile.ZipFile(temp) as target:
            if target.testzip() is not None or target.namelist() != order:
                raise ValueError('ZIP order verification failure')
            if any(target.read(n) != members[n][1] for n in order):
                raise ValueError('ZIP content changed')
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)
