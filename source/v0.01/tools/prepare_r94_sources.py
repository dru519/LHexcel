"""Update owned release labels and deterministic template test inventory."""
import hashlib
import json
from pathlib import Path

root = Path(__file__).resolve().parents[1]
for rel in ("src/vba/ui/NxProductAbout.bas", "contracts/build-profile-contract.json",
            "tools/verify_build_profile.py", "tools/build_enhanced_distribution.py",
            "build/Verify-NxEnhancedPackage.ps1"):
    path = root / rel
    text = path.read_text(encoding="utf-8-sig")
    path.write_text(text.replace("r93", "r94"), encoding="utf-8")
path = root / "tests/manifests/Template.json"
data = json.loads(path.read_text(encoding="utf-8-sig"))
for rel in ("src/vba/features/template/NxTemplateLibrary.bas",
            "src/vba/features/template/CNxTemplateHeaderSizer.cls"):
    if not any(item["path"] == rel for item in data["modules"]):
        data["modules"].append(dict(path=rel, sha256="", owner_phase="template"))
for item in data["modules"]:
    item["sha256"] = hashlib.sha256((root / item["path"]).read_bytes()).hexdigest()
path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
path = root / "contracts/feature-contract.json"
data = json.loads(path.read_text(encoding="utf-8"))
for feature in data["features"]:
    if feature["id"] in {"NX-TPL-LIST", "NX-TPL-LOAD", "NX-TPL-RENAME", "NX-TPL-DELETE"}:
        feature.update(dialog_id="NX-DLG-TPL-MANAGER", dialog_template="FNxTemplateManager", ui_surface="medium")
        if feature["id"] == "NX-TPL-LIST":
            feature["ribbon"]["label"] = "템플릿"
            feature["catalog_path"] = ["템플릿", "템플릿"]
            feature["description_ko"] = "등록한 엑셀 양식을 검색하고 순서를 조정합니다. 파일·시트 등록, 상세 편집, 사용과 삭제를 제공합니다."
        if feature["id"] == "NX-TPL-DELETE":
            feature["description_ko"] = "선택한 양식을 보관함 목록에서 삭제합니다. 등록에 사용한 원본 파일은 변경하지 않습니다."
        if feature["id"] == "NX-TPL-LOAD":
            feature["privacy"]["local_file_creation"] = True
path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
