"""One-shot mechanical hook insertion into active forms with a preview text box."""
import json
from pathlib import Path

root = Path(__file__).resolve().parents[2]
manifest = json.loads((root / 'build/manifests/Product.json').read_text(encoding='utf-8'))
for binding in manifest['form_bindings']:
    layout = json.loads((root / binding['layout']).read_text(encoding='utf-8-sig'))
    controls = [c for c in layout['controls'] if c['type'] == 'TextBox' and c['name'] == 'txtPreview']
    if not controls:
        continue
    path = root / binding['code']
    text = path.read_text(encoding='utf-8-sig')
    if 'Private Sub txtPreview_Change()' in text:
        continue
    text += '\nPrivate Sub txtPreview_Change()\n    NxNativePreviewTextChanged Me, txtPreview\nEnd Sub\n'
    path.write_text(text, encoding='utf-8', newline='\n')
    print(path.name)
