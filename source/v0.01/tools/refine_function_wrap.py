"""Apply the approved compact Template/Function ribbon and wrapper commands."""
import copy
import json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1]
def read(path): return json.loads((ROOT/path).read_text(encoding='utf-8'))
def write(path, value): (ROOT/path).write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
shell=read('contracts/ribbon-shell-contract.json')
groups=shell['surfaces']['product']['groups']
template=next(g for g in groups if g['id']=='NX-GRP-TEMPLATE')
data=next(g for g in groups if g['id']=='NX-GRP-DATA')
groups.remove(data); groups.insert(groups.index(template),data)
template['label']=' '
button={'kind':'button','id':'NX-TEMPLATE-OPEN','feature_id':'NX-TPL-LIST','label':'템플릿','image_mso':'FileNew','keytip':'T','size':'normal','control_type':2}
template['controls']=[{'kind':'box','id':'NX-TEMPLATE-FUNCTION-VERTICAL','box_style':'vertical','box':[
    button, {'kind':'menu','id':'NX-FUNCTION-MENU','label':'함수','image_mso':'FunctionWizard','control_type':2,
    'menu':[{'kind':'button','id':'NX-FUNCTION-'+key,'command_id':'NX-CMD-FUNCTION-'+key} for key in ['ROUND','IFERROR']]}]}]
for category in shell['navigation_categories']:
    if category['id']=='NX-GRP-FORMULA': category['label']='함수';category['image_mso']='FunctionWizard'
write('contracts/ribbon-shell-contract.json',shell)
source=read('src/resources/commands.ko-KR.json')
base=next(r for r in source['commands'] if r['handler']=='NxCmdFormula')
for index,key in enumerate(['ROUND','IFERROR']):
    command=copy.deepcopy(base)
    command.update(id='NX-CMD-FUNCTION-'+key,label_ko=key+' 감싸기',aliases_ko=[],
        description_ko=('기존 값과 수식을 ROUND·ROUNDUP·ROUNDDOWN으로 감쌉니다. 오류는 그대로 유지합니다.' if key=='ROUND' else '기존 수식이나 숫자를 IFERROR로 감싸 오류일 때 공란·숫자·문자를 반환합니다.'),
        image_mso='FunctionWizard',search_terms_ko=[key,'함수','감싸기'],sort_order=1100+index,
        input_context='range_selection',active_when='excel_ready',mutates_document=True,
        recovery_policy='command_scoped_guard',ribbon_visible=True,taskpane_visible=True,
        favorite_eligible=True,shortcut_eligible=True,args=['NX_FUNCTION_'+key],legacy_ids=[],legacy_entrypoints=[])
    command['navigation'].update(ribbon_group_id='NX-GRP-FORMULA',purpose_tags=[key+' 감싸기'],
        launch_surface='dialog',mutation_scope='document',selection_mode='required',feedback_mode='inline')
    source['commands']=[r for r in source['commands'] if r['id']!=command['id']]+[command]
for group in source['groups']:
    if group['id']==base['group_id']: group['raw_count']=4
source['normalized_command_count']=len(source['commands'])
source['raw_command_count']=sum(g['raw_count'] for g in source['groups'])
write('src/resources/commands.ko-KR.json',source)
