"""Apply the approved single-entry template navigation."""
import json
from pathlib import Path

root = Path(__file__).resolve().parents[1]
path = root / "contracts/feature-contract.json"
data = json.loads(path.read_text(encoding="utf-8-sig"))
for feature in data["features"]:
    if feature["owner"] != "template":
        continue
    visible = feature["id"] == "NX-TPL-LIST"
    feature["catalog_visible"] = visible
    feature["ribbon"]["exposed"] = visible
    if visible:
        feature["ribbon"]["label"] = "템플릿"
        feature["catalog_path"] = ["템플릿", "템플릿"]
path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
path = root / "contracts/ribbon-shell-contract.json"
data = json.loads(path.read_text(encoding="utf-8-sig"))
group = next(g for g in data["surfaces"]["product"]["groups"] if g["id"] == "NX-GRP-TEMPLATE")
group["controls"] = [dict(kind="button", id="NX-TEMPLATE-OPEN", feature_id="NX-TPL-LIST",
                          label="템플릿", image_mso="FileNew", keytip="T", size="large", control_type=2)]
path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
