"""Import the user-approved design originals without modifying them."""
import csv
import hashlib
import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DESIGN = ROOT.parents[1] / "08_디자인소스/260927 18;09 LHexcel 내엑셀 확정 디자인셋"


def main():
    target = ROOT / "src/ribbon/images/brand"
    target.mkdir(parents=True, exist_ok=True)
    catalog = json.loads((DESIGN / "07_엑셀컬러아이콘/catalog.json").read_text(encoding="utf-8-sig"))
    images = {}
    for item in catalog:
        for size in (16, 32):
            key = item["key"]
            source = DESIGN / f"07_엑셀컬러아이콘/png/{size}/{key}.png"
            relative = f"images/brand/{key}-{size}.png"
            shutil.copyfile(source, ROOT / "src/ribbon" / relative)
            images[f"NxBrand_{key.replace('-', '_')}_{size}"] = {
                "path": relative, "sha256": hashlib.sha256(source.read_bytes()).hexdigest()
            }
    for size in (16, 32):
        source = DESIGN / f"01_브랜드/app-{size}.png"
        relative = f"images/brand/product-{size}.png"
        shutil.copyfile(source, ROOT / "src/ribbon" / relative)
        images[f"NxBrand_product_{size}"] = {"path": relative, "sha256": hashlib.sha256(source.read_bytes()).hexdigest()}
    with (DESIGN / "07_엑셀컬러아이콘/기능대응표.csv").open(encoding="utf-8-sig", newline="") as stream:
        tags = {row["tag"]: row["icon"] for row in csv.DictReader(stream)}
    manifest = {"source": DESIGN.name, "tags": tags,
                "labels": {item["label"]: item["key"] for item in catalog}, "images": images}
    manifest["labels"].update({"내엑셀": "product", "전체기능": "all", "저장": "save", "인쇄": "print",
        "복붙": "copy", "삽입": "image", "파일관리": "folder", "데이터": "normalize", "데이터 비교": "compare-file",
        "함수": "formula-map", "셀맞춤": "resize", "시트·보기": "view", "스타일": "style", "스타일 조정": "style",
        "표시 형식": "decimal", "배치·병합": "merge", "추가기능": "menu", "정보진단": "info"})
    (ROOT / "src/ribbon/brand-images.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    setup = ROOT / "src/dotnet/NxSetup/Assets"
    setup.mkdir(exist_ok=True)
    shutil.copyfile(DESIGN / "01_브랜드/lhexcel.ico", setup / "lhexcel.ico")
    shutil.copyfile(DESIGN / "01_브랜드/app-48.png", setup / "brand-48.png")
    host = ROOT / "src/dotnet/NxHost/Assets"
    host.mkdir(exist_ok=True)
    for theme in ("ink", "reverse"):
        for key in ("refresh", "chevron"):
            shutil.copyfile(DESIGN / f"02_아이콘/png/{theme}/48/{key}.png", host / f"{key}-{theme}.png")
    print(f"Imported {len(images)} ribbon images and installer brand assets")


if __name__ == "__main__":
    main()
