"""Apply the approved formula notes/report UI and navigator icon contracts."""
import json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]

def read(path): return json.loads((ROOT/path).read_text(encoding="utf-8"))
def write(path, value): (ROOT/path).write_text(json.dumps(value, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")

path=ROOT/"src/vba/ui/NxCommandHandlers.bas"
text=path.read_text(encoding="utf-8")
start=text.index("Private Sub NxAddFormulaComments(")
end=text.index("Private Sub NxCycleEnterDirection()", start)
text=text[:start]+'''Private Sub NxAddFormulaComments(ByVal target As Excel.Range)
    NxFormulaToolsOpen target, False
End Sub

Public Sub NxFormulaCommentsApply(ByVal target As Excel.Range, ByVal showNotes As Boolean)
    NxFormulaNotesApply target, showNotes
End Sub

'''+text[end:]
start=text.index("Private Sub NxCreatePrecedentsIndex(")
end=text.index("End Sub",start)+len("End Sub")
text=text[:start]+'''Private Sub NxCreatePrecedentsIndex(ByVal target As Excel.Range)
    NxFormulaToolsOpen target, True
End Sub'''+text[end:]
path.write_text(text,encoding="utf-8")
path=ROOT/"src/vba/ui/NxCommandRegistry.bas"
text=path.read_text(encoding="utf-8")
line='    If definition.ArgumentKey = "RB_EDIT_MEMO_ADD_LHEXCELFORMULA" Then NxFormulaNotesUndoArm'
if line not in text: text=text.replace('    If definition.Handler = "NxCmdStyle" And Left$',line+'\n    If definition.Handler = "NxCmdStyle" And Left$')
path.write_text(text,encoding="utf-8")

contract=read("contracts/feature-contract.json")
def walk(node):
    if isinstance(node,dict):
        if node.get("id")=="NX-UTIL-DOCUMENT-NAVIGATOR": node["ribbon"]["image_mso"]="FindDialog"
        for value in node.values(): walk(value)
    elif isinstance(node,list):
        for value in node: walk(value)
walk(contract);write("contracts/feature-contract.json",contract)
commands=read("src/resources/commands.ko-KR.json")
for command in commands["commands"]:
    if "RB_FORMULA_PRECEDENTS_LIST" in command.get("args",[]):
        command["label_ko"]="수식 참조표 만들기"
        command["description_ko"]="선택한 수식의 참조 위치와 값을 한글 표로 정리합니다. 새 시트 또는 새 통합문서를 선택하며, 확인하지 못한 참조는 확인 상태에 표시합니다."
        command["navigation"]["launch_surface"]="dialog"
        command["navigation"]["feedback_mode"]="inline"
    if "RB_EDIT_MEMO_ADD_LHEXCELFORMULA" in command.get("args",[]):
        command["description_ko"]="기존 메모를 유지하고 수식을 추가하거나 덮어씁니다. 메모 표시 여부를 선택하고 마지막 기록을 되돌릴 수 있습니다."
        command["navigation"]["feedback_mode"]="inline"
write("src/resources/commands.ko-KR.json",commands)

controls=[]
def control(kind,name,x,y,w,h,**kwargs):
    controls.append(dict(type=kind,name=name,left=x,top=y,width=w,height=h,enabled=True,tab_index=len(controls),tab_stop=kind!="Label",font_name="맑은 고딕",font_size=9,**kwargs))
control("Label","lblRange",14,12,372,38,caption="선택 범위")
control("Label","lblOption",14,60,72,20,caption="결과 위치")
control("ComboBox","cboOption",90,56,296,24,style="drop_down_list",match_required=True,items=[],selected_index=-1)
control("CheckBox","chkShow",90,90,270,20,caption="기록 후 메모 표시",value=False)
control("Label","lblGuide",14,120,372,52,caption="")
control("CommandButton","cmdExecute",234,184,72,26,caption="적용")
control("CommandButton","cmdCancel",314,184,72,26,caption="취소")
write("src/vba/ui/FNxFormulaTools.form.json",dict(schema_version=2,name="FNxFormulaTools",caption="내엑셀 - 수식",width=400,height=224,client_width=400,client_height=224,start_up_position=1,code_path="src/vba/ui/FNxFormulaTools.vba",default_control="cmdExecute",cancel_control="cmdCancel",controls=controls))
