"""Update the approved comparison controls without altering other form metadata."""
import json
from pathlib import Path
root = Path(__file__).resolve().parents[1]
path = root / "src/vba/features/file/ui/FNxWorkbookCompare.form.json"
form = json.loads(path.read_text(encoding="utf-8"))
if not any(c["name"] == "cboBaseSheet" for c in form["controls"]):
    form["height"] += 86
    form["client_height"] += 86
    for c in form["controls"]:
        if c["top"] >= 116:
            c["top"] += 86
    def combo(name, left, top, width, index, items):
        return dict(type="ComboBox", name=name, left=left, top=top, width=width,
                    height=24, tab_index=index, tab_stop=True, enabled=True,
                    items=items, selected_index=0, style="drop_down_list", match_required=True)
    form["controls"] += [
        combo("cboBaseSheet", 16, 112, 172, 12, ["전체 시트"]),
        combo("cboCompareSheet", 198, 112, 172, 13, ["전체 시트"]),
        dict(type="CommandButton",name="cmdLoadSheets",left=380,top=112,width=84,height=24,
             tab_index=14,tab_stop=True,enabled=True,caption="시트 목록"),
        dict(type="Label",name="lblOutput",left=16,top=152,width=78,height=20,
             tab_index=-1,tab_stop=False,enabled=True,caption="결과 위치"),
        combo("cboOutput",96,148,274,15,["새 창(새 통합문서)","현재 파일의 새 시트"])
    ]
    next(c for c in form["controls"] if c["name"]=="lblTitle")["caption"]="파일과 시트를 선택하세요. 기준·비교 사본과 차이 목록을 만듭니다."
    preview=next(c for c in form["controls"] if c["name"]=="txtPreview")
    preview["value"]="전체 시트 또는 시트 목록에서 양쪽 시트를 선택한 뒤 바로 비교할 수 있습니다."
    preview["back_color"]=16777215
    path.write_text(json.dumps(form,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")

# Keep both sheet selectors labelled and aligned with their file fields.
controls = {c["name"]: c for c in form["controls"]}
form["height"] = form["client_height"] = 344
for name, top in (("cboBaseSheet",112),("cboCompareSheet",146),("cboOutput",180)):
    controls[name].update(left=96, top=top, width=272)
controls["lblOutput"]["top"] = 184
controls["chkCompareFormats"]["top"] = 210
controls["txtPreview"]["top"] = 240
for name in ("cmdResults", "cmdPreview", "cmdExecute", "cmdCancel"):
    controls[name]["top"] = 302
for name, top, caption in (("lblBaseSheet",116,"기준 시트"),("lblCompareSheet",150,"비교 시트")):
    if name not in controls:
        form["controls"].append(dict(type="Label",name=name,left=16,top=top,width=78,height=20,
                                     tab_index=-1,tab_stop=False,enabled=True,caption=caption))
path.write_text(json.dumps(form,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
