"""One-time mechanical layout update for the approved template library."""
import json
from pathlib import Path

root = Path(__file__).resolve().parents[1]
path = root / "src/vba/features/template/ui/FNxTemplateManager.form.json"
form = json.loads(path.read_text(encoding="utf-8-sig"))
form.update(caption="내엑셀 - 템플릿 보관함", width=650, height=500, client_width=650, client_height=500, default_control="cmdUse")
controls = {c["name"]: c for c in form["controls"]}
positions = {
    "lblTitle": (12, 12, 300, 18), "lblFolder": (12, 50, 66, 14),
    "txtFolder": (78, 46, 436, 22), "cmdOpenFolder": (524, 46, 112, 22),
    "lstTemplates": (12, 138, 508, 180),
    "cmdRegisterFile": (532, 138, 104, 26),
    "cmdRegisterSheet": (532, 172, 104, 26), "cmdUse": (414, 462, 106, 26),
    "cmdDelete": (532, 284, 104, 26), "cmdCatalog": (422, 326, 98, 26),
    "cmdRefresh": (336, 326, 80, 26), "lblStatus": (12, 412, 335, 28),
    "cmdCancel": (530, 462, 106, 26),
}
for name, (left, top, width, height) in positions.items():
    controls[name].update(left=left, top=top, width=width, height=height)
controls["lblTitle"]["caption"] = "템플릿 보관함"
controls["lstTemplates"]["column_widths"] = "200;90;90;114;0"
controls["cmdRegisterFile"]["caption"] = "엑셀파일 등록"
controls["cmdRegisterSheet"]["caption"] = "엑셀시트 등록"
added = {"cmdOpenCopy", "cmdRegisterWorkbook", "txtSearch", "cboFilter", "cboSort", "lblName", "txtName", "lblCategory",
         "txtCategory", "lblDescription", "txtDescription", "cmdEdit", "cboOutput", "chkDeleted", "lblSearch", "lblOutput",
         "lblList", "cmdUp", "cmdDown", "header1", "header2", "header3", "header4", "grip1", "grip2", "grip3"}
form["controls"] = [c for c in form["controls"] if c["name"] not in added]
def add(kind, name, x, y, width, height, caption=None, **extra):
    c = dict(type=kind, name=name, left=x, top=y, width=width, height=height,
             tab_index=len(form["controls"]), tab_stop=kind != "Label", enabled=True,
             font_name="맑은 고딕", font_size=9)
    if caption is not None: c["caption"] = caption
    if kind == "ComboBox":
        choices = {"cboFilter": ["전체 분류"], "cboSort": ["사용자 순서", "최근 수정순", "최근 사용순", "이름순"],
                   "cboOutput": ["현재 문서에 새 시트", "새 통합문서"]}
        c.update(style="drop_down_list", match_required=True, items=choices[name], selected_index=0)
    c.update(extra)
    form["controls"].append(c)
add("Label", "lblSearch", 12, 76, 220, 14, "이름·설명·분류 검색")
add("TextBox", "txtSearch", 12, 92, 340, 22, value="")
add("ComboBox", "cboFilter", 360, 92, 140, 22)
add("ComboBox", "cboSort", 508, 92, 128, 22)
add("CommandButton", "cmdUp", 532, 216, 104, 26, "위로")
add("CommandButton", "cmdDown", 532, 250, 104, 26, "아래로")
add("Label", "lblName", 12, 364, 90, 14, "양식 이름")
add("TextBox", "txtName", 12, 384, 190, 22, value="")
add("Label", "lblCategory", 210, 364, 90, 14, "분류")
add("TextBox", "txtCategory", 210, 384, 142, 22, value="")
add("Label", "lblDescription", 360, 364, 90, 14, "설명")
add("TextBox", "txtDescription", 360, 384, 166, 48, value="", multi_line=True)
add("CommandButton", "cmdEdit", 536, 384, 100, 26, "상세 저장")
add("Label", "lblOutput", 12, 444, 100, 14, "사용 위치")
add("ComboBox", "cboOutput", 12, 462, 390, 26)
cursor = 12
for index, (title, width) in enumerate(zip(["양식 이름", "분류", "등록 방식", "수정일"], [200,90,90,114]), 1):
    add("Label", f"header{index}", cursor, 120, width-2, 18, title)
    cursor += width
    if index < 4:
        add("Label", f"grip{index}", cursor-2, 120, 4, 18, "│")
path.write_text(json.dumps(form, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

path = root / "src/vba/features/template/ui/FNxTemplateRegister.form.json"
form = json.loads(path.read_text(encoding="utf-8-sig"))
if not any(c["name"] == "txtCategory" for c in form["controls"]):
    for key in ("height", "client_height"):
        if key in form: form[key] += 34
    for control in form["controls"]:
        if control["top"] >= 112: control["top"] += 34
    add("Label", "lblCategory", 12, 112, 68, 16, "분류")
    add("TextBox", "txtCategory", 84, 108, 384, 22, value="")
    path.write_text(json.dumps(form, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
