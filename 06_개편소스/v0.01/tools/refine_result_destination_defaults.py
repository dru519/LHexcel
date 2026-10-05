"""Apply the user's new-workbook default to existing result selectors."""
import json
from pathlib import Path
root = Path(__file__).resolve().parents[1]
forms = {
    "src/vba/features/data/ui/FNxDataNormalize.form.json": ("cboOutputMode", "새 통합문서"),
    "src/vba/features/data/ui/FNxDataSpecial.form.json": ("cboOutputMode", "새 통합문서"),
    "src/vba/features/data/ui/FNxAgeCalculator.form.json": ("cboOutputMode", "새 통합문서"),
    "src/vba/features/data/ui/FNxDataAnalyze.form.json": ("cboOutputMode", "새창으로 보기"),
    "src/vba/features/file/ui/FNxDataCompare.form.json": ("cboOutput", "새 창(새 통합문서)"),
    "src/vba/features/file/ui/FNxWorkbookCompare.form.json": ("cboOutput", "새 창(새 통합문서)"),
    "src/vba/features/template/ui/FNxTemplateManager.form.json": ("cboOutput", "새 통합문서"),
}
for relative, (name, default) in forms.items():
    path = root / relative
    form = json.loads(path.read_text(encoding="utf-8"))
    control = next(c for c in form["controls"] if c["name"] == name)
    control["items"] = [default] + [x for x in control["items"] if x != default]
    control["selected_index"] = 0
    path.write_text(json.dumps(form, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")

replacements = {
    "tests/python/test_r57_data_about_sheetindex_contract.py": [
        ('["새 시트 만들기", "새 통합문서", "오른쪽 새 열 삽입", "원본 변경"]', '["새 통합문서", "새 시트 만들기", "오른쪽 새 열 삽입", "원본 변경"]')],
    "tests/python/test_r69_output_workflow.py": [
        ('["새 시트 만들기", "새 통합문서", "오른쪽 새 열 삽입", "원본 변경"]', '["새 통합문서", "새 시트 만들기", "오른쪽 새 열 삽입", "원본 변경"]')],
    "tests/python/test_r62_data_analysis.py": [
        ("['새 시트 만들기', '새창으로 보기']", "['새창으로 보기', '새 시트 만들기']")],
    "tests/python/test_r62_data_special.py": [
        ('self.assertEqual("새 시트 만들기", choices["cboOutputMode"]["items"][0])', 'self.assertEqual("새 통합문서", choices["cboOutputMode"]["items"][0])')],
    "tests/python/test_r62_data_name_restore.py": [
        ('self.assertEqual("새 시트 만들기", next(', 'self.assertEqual("새 통합문서", next('),
        ("기본값은 새 시트 만들기입니다.", "기본값은 새 통합문서입니다.")],
    "contracts/feature-contract.json": [
        ("기본값은 새 시트 만들기입니다.", "기본값은 새 통합문서입니다.")],
    "tests/vba/product/T_R69OutputFlows.bas": [
        ('panel.cboOutputMode.Value = "새 시트 만들기", "Date default"', 'panel.cboOutputMode.Value = "새 통합문서", "Date default"')],
    "tests/vba/product/T_R69UserFlows.bas": [
        ('panel.cboOutputMode.Value = "새 시트 만들기", "Default output changed"', 'panel.cboOutputMode.Value = "새 통합문서", "Default output changed"')],
    "tests/vba/product/T_R94FormulaTools.bas": [
        ('Require report.Parent Is book, "popup default new sheet"', 'Require Not report.Parent Is book, "popup default new workbook"')],
    "tools/refine_compare_output_layout.py": [
        ('["현재 파일의 새 시트","새 창(새 통합문서)"]', '["새 창(새 통합문서)","현재 파일의 새 시트"]')],
}
for relative, pairs in replacements.items():
    path = root / relative
    text = path.read_text(encoding="utf-8")
    for old, new in pairs:
        text = text.replace(old, new)
    path.write_text(text, encoding="utf-8")
