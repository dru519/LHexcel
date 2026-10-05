Attribute VB_Name = "NxGeneratedNavigationCatalog"
Option Explicit

Private mNavigationItems As Collection
Private mNavigationRouteIndex As Object

Public Function NxGeneratedNavigationItems() As Collection
    Dim items As Collection
    If mNavigationItems Is Nothing Then
        Set items = New Collection
        NxGeneratedNavigationAddItems01 items
        NxGeneratedNavigationAddItems02 items
        NxGeneratedNavigationAddItems03 items
        Set mNavigationItems = items
    End If
    Set NxGeneratedNavigationItems = mNavigationItems
End Function

Private Sub NxGeneratedNavigationAddItems01(ByRef items As Collection)
    items.Add Array("feature:NX-FILE-MANNER-SAVE", "feature", "NX-FILE-MANNER-SAVE", "매너 저장", "모든 시트의 커서를 A1로 이동하고 첫 번째 시트의 A1을 선택한 상태로 저장합니다. 파일을 다시 열면 첫 번째 시트의 A1에서 시작합니다.", "NX-GRP-SAVE", "저장", "FileSave", "active_workbook", "none", "direct", "document", "inline", "guarded", "NX-FILE-MANNER-SAVE | 매너 저장 | 저장", "workbook,filesystem")
    items.Add Array("feature:NX-FILE-SHEET-COPY-SAVE", "feature", "NX-FILE-SHEET-COPY-SAVE", "현재 시트 사본 저장", "현재 워크시트를 별도 Excel 파일로 저장합니다. 결과 경로를 확인하며 원본 문서를 덮어쓰지 않습니다.", "NX-GRP-SAVE", "저장", "FileSaveAs", "active_workbook", "none", "direct", "filesystem", "completion", "planned", "NX-FILE-SHEET-COPY-SAVE | 현재 시트 사본 저장 | 저장", "workbook,filesystem")
    items.Add Array("feature:NX-FILE-RANGE-COPY-SAVE", "feature", "NX-FILE-RANGE-COPY-SAVE", "선택 범위 사본 저장", "선택한 셀 영역으로 별도 Excel 파일을 만듭니다. 전체 시트 저장과 달리 지정한 범위가 결과 대상입니다.", "NX-GRP-SAVE", "저장", "FileSaveAs", "range_selection", "required", "direct", "filesystem", "completion", "planned", "NX-FILE-RANGE-COPY-SAVE | 선택 범위 사본 저장 | 저장", "selection,filesystem")
    items.Add Array("feature:NX-FILE-PDF-CURRENT-SHEET", "feature", "NX-FILE-PDF-CURRENT-SHEET", "현재 시트 PDF 저장", "현재 워크시트를 한 PDF 파일로 저장합니다. 인쇄영역과 페이지 설정을 확인한 뒤 결과 위치를 선택하세요.", "NX-GRP-SAVE", "저장", "FileSaveAsPdfOrXps", "active_workbook", "none", "dialog", "filesystem", "completion", "planned", "NX-FILE-PDF-CURRENT-SHEET | 현재 시트 PDF 저장 | 저장", "workbook,filesystem")
    items.Add Array("feature:NX-FILE-PDF-ALL-COMBINED", "feature", "NX-FILE-PDF-ALL-COMBINED", "문서 PDF 저장", "통합문서의 출력 대상 시트를 순서대로 한 PDF에 담습니다. 숨김 시트 등 제외 조건은 저장 화면에서 확인하세요.", "NX-GRP-SAVE", "저장", "FileSaveAsPdfOrXps", "active_workbook", "none", "dialog", "filesystem", "completion", "planned", "NX-FILE-PDF-ALL-COMBINED | 문서 PDF 저장 | 저장", "workbook,filesystem")
    items.Add Array("feature:NX-FILE-RANGE-PNG", "feature", "NX-FILE-RANGE-PNG", "선택 범위 그림 저장", "선택 범위를 문서 폴더에 PNG로 바로 저장합니다. 미저장 문서는 Excel 기본 저장 폴더를 사용하며 같은 이름은 새 번호로 구분합니다. 오류가 있을 때만 알립니다.", "NX-GRP-SAVE", "저장", "FileSaveAs", "range_selection", "required", "direct", "filesystem", "inline", "planned", "NX-FILE-RANGE-PNG | 선택 범위 그림 저장 | 저장", "selection,filesystem")
    items.Add Array("feature:NX-FILE-CHART-PNG", "feature", "NX-FILE-CHART-PNG", "차트 그림 저장", "선택한 차트를 그림 파일로 내보냅니다. 저장 형식과 위치를 선택하며 문서 안의 차트는 그대로 둡니다.", "NX-GRP-SAVE", "저장", "FileSaveAs", "shape_selection", "required", "dialog", "filesystem", "completion", "planned", "NX-FILE-CHART-PNG | 차트 그림 저장 | 저장", "shape,filesystem")
    items.Add Array("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK", "command", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK", "선택 범위 인쇄 미리보기", "선택 범위를 인쇄 대상으로 지정하고 세로 방향·너비 한 쪽으로 미리 봅니다. 이 단계에서 프린터로 출력하지 않습니다.", "NX-GRP-PRINT", "인쇄", "PrintPreviewAndPrint", "active_worksheet", "optional", "direct", "document", "completion", "guarded", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK | 선택 범위 인쇄 미리보기 | 인쇄", "worksheet")
    items.Add Array("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT", "command", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT", "세로 방향 · 너비 1쪽", "선택 범위를 세로 방향으로 배치합니다. 가로는 한 쪽에 맞추고 세로는 내용 길이에 따라 나눕니다.", "NX-GRP-PRINT", "인쇄", "PrintPreviewAndPrint", "active_worksheet", "optional", "direct", "document", "completion", "guarded", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT | 세로 방향 · 너비 1쪽 | 인쇄", "worksheet")
    items.Add Array("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE", "command", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE", "가로 방향 · 너비 1쪽", "선택 범위를 가로 방향으로 배치합니다. 가로는 한 쪽에 맞추고 세로는 내용 길이에 따라 나눕니다.", "NX-GRP-PRINT", "인쇄", "PrintPreviewAndPrint", "active_worksheet", "optional", "direct", "document", "completion", "guarded", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE | 가로 방향 · 너비 1쪽 | 인쇄", "worksheet")
    items.Add Array("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT-1BY1", "command", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT-1BY1", "세로 방향 · 한 장에 맞춤", "선택 범위 전체가 세로 방향 용지 한 장에 들어가도록 인쇄 배율을 설정합니다. 범위가 크면 글자가 작아질 수 있습니다.", "NX-GRP-PRINT", "인쇄", "PrintPreviewAndPrint", "active_worksheet", "optional", "direct", "document", "completion", "guarded", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT-1BY1 | 세로 방향 · 한 장에 맞춤 | 인쇄", "worksheet")
    items.Add Array("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE-1BY1", "command", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE-1BY1", "가로 방향 · 한 장에 맞춤", "선택 범위 전체가 가로 방향 용지 한 장에 들어가도록 인쇄 배율을 설정합니다. 범위가 크면 글자가 작아질 수 있습니다.", "NX-GRP-PRINT", "인쇄", "PrintPreviewAndPrint", "active_worksheet", "optional", "direct", "document", "completion", "guarded", "NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE-1BY1 | 가로 방향 · 한 장에 맞춤 | 인쇄", "worksheet")
    items.Add Array("command:NX-CMD-PRINT-RB-PRINT-SETUP-REPEAT", "command", "NX-CMD-PRINT-RB-PRINT-SETUP-REPEAT", "선택 행을 인쇄 제목으로", "선택한 셀이 속한 행을 각 인쇄 페이지의 위쪽에 반복할 제목 행으로 지정합니다.", "NX-GRP-PRINT", "인쇄", "PrintTitles", "active_worksheet", "optional", "direct", "document", "completion", "guarded", "NX-CMD-PRINT-RB-PRINT-SETUP-REPEAT | 선택 행을 인쇄 제목으로 | 인쇄", "worksheet")
    items.Add Array("feature:NX-DATA-COPY-VISIBLE", "feature", "NX-DATA-COPY-VISIBLE", "보이는 셀만 복사", "숨김·필터로 가려진 셀을 제외하고 보이는 셀의 값을 복사합니다. 붙여넣을 범위 크기를 함께 확인하세요.", "NX-GRP-COPY", "복붙", "Copy", "range_selection", "required", "direct", "none", "inline", "fast", "NX-DATA-COPY-VISIBLE | 보이는 셀만 복사 | 복붙", "selection,clipboard")
    items.Add Array("feature:NX-DATA-PASTE-VISIBLE-VALUES", "feature", "NX-DATA-PASTE-VISIBLE-VALUES", "보이는 셀에 값 붙여넣기", "복사한 값을 대상의 보이는 셀에 입력합니다. 숨김 셀은 건너뛰며 대상 크기가 맞는지 실행 전에 확인합니다.", "NX-GRP-COPY", "복붙", "Paste", "range_selection", "required", "direct", "document", "completion", "planned", "NX-DATA-PASTE-VISIBLE-VALUES | 보이는 셀에 값 붙여넣기 | 복붙", "selection,clipboard")
    items.Add Array("command:NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYFORMULA", "command", "NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYFORMULA", "첫 셀 수식 복사", "선택 범위의 첫 셀에서 수식을 텍스트로 가져옵니다. 첫 셀에 수식이 있어야 합니다.", "NX-GRP-COPY", "복붙", "Copy", "selection_or_clipboard", "optional", "direct", "none", "inline", "fast", "NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYFORMULA | 첫 셀 수식 복사 | 복붙", "selection,clipboard")
    items.Add Array("command:NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYREFERENCE", "command", "NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYREFERENCE", "첫 셀 주소 복사", "선택 범위의 첫 셀 주소를 문서·시트 이름과 함께 클립보드에 담습니다.", "NX-GRP-COPY", "복붙", "NameDefine", "selection_or_clipboard", "optional", "direct", "none", "inline", "fast", "NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYREFERENCE | 첫 셀 주소 복사 | 복붙", "selection,clipboard")
    items.Add Array("command:NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYTEXT", "command", "NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYTEXT", "보이는 값 복사", "화면에 표시되는 셀 내용을 탭으로 구분한 텍스트로 가져옵니다. 수식 자체는 포함하지 않습니다.", "NX-GRP-COPY", "복붙", "Copy", "selection_or_clipboard", "optional", "direct", "none", "inline", "fast", "NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYTEXT | 보이는 값 복사 | 복붙", "selection,clipboard")
    items.Add Array("command:NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMULAS", "command", "NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMULAS", "보이는 셀에 수식 붙여넣기", "대상에서 숨겨지지 않은 셀만 골라 Excel 클립보드의 수식을 넣습니다. 복사 범위와 대상 구성을 확인하세요.", "NX-GRP-COPY", "복붙", "PasteFormulas", "selection_or_clipboard", "optional", "direct", "document", "completion", "guarded", "NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMULAS | 보이는 셀에 수식 붙여넣기 | 복붙", "selection,clipboard")
    items.Add Array("command:NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMATS", "command", "NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMATS", "보이는 셀에 서식 붙여넣기", "대상에서 숨겨지지 않은 셀에만 복사한 서식을 적용합니다. 숨김 셀의 서식은 바꾸지 않습니다.", "NX-GRP-COPY", "복붙", "PasteFormatting", "selection_or_clipboard", "optional", "direct", "document", "completion", "guarded", "NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMATS | 보이는 셀에 서식 붙여넣기 | 복붙", "selection,clipboard")
    items.Add Array("feature:NX-DRAW-INSERT-PICTURE", "feature", "NX-DRAW-INSERT-PICTURE", "그림 삽입", "새 그림 파일을 여러 개 삽입하거나 선택한 그림을 셀 범위에 맞춥니다. 대상 범위·크기 방식·여백을 미리 확인한 뒤 적용합니다.", "NX-GRP-INSERT", "삽입", "PictureInsertFromFile", "selection_optional", "optional", "dialog", "document", "completion", "planned", "NX-DRAW-INSERT-PICTURE | 그림 삽입 | 삽입", "selection,shape")
    items.Add Array("command:NX-CMD-STRUCTURE-RB-EDIT-CELL-MAKEGROUP-HIDDEN-ROWCOLUMN", "command", "NX-CMD-STRUCTURE-RB-EDIT-CELL-MAKEGROUP-HIDDEN-ROWCOLUMN", "선택 행·열 그룹 만들기", "가로로 넓은 선택은 열 그룹, 세로로 긴 선택은 행 그룹으로 묶습니다. 숨겨진 항목을 자동 탐색하는 기능은 아닙니다.", "NX-GRP-INSERT", "삽입", "PropertySheet", "active_worksheet", "optional", "direct", "document", "completion", "guarded", "NX-CMD-STRUCTURE-RB-EDIT-CELL-MAKEGROUP-HIDDEN-ROWCOLUMN | 선택 행·열 그룹 만들기 | 삽입", "worksheet,selection")
    items.Add Array("feature:NX-FILE-FOLDER-CREATE", "feature", "NX-FILE-FOLDER-CREATE", "폴더 일괄 생성", "셀의 문구를 폴더 이름으로 사용합니다. 행별 계층 경로와 저장 위치를 확인한 뒤 필요한 폴더를 만듭니다.", "NX-GRP-FILE", "파일관리", "Folder", "selection_optional", "optional", "dialog", "filesystem", "completion", "planned", "NX-FILE-FOLDER-CREATE | 폴더 일괄 생성 | 파일관리", "selection,filesystem")
    items.Add Array("feature:NX-FILE-CONSOLIDATE", "feature", "NX-FILE-CONSOLIDATE", "파일 통합", "여러 문서의 시트를 모으거나 데이터를 한 시트로 이어 붙입니다. 입력 순서와 결과 구성을 확인하고 새 파일로 저장합니다.", "NX-GRP-FILE", "파일관리", "FileOpen", "active_workbook", "none", "dialog", "filesystem", "completion", "planned", "NX-FILE-CONSOLIDATE | 파일 통합 | 파일관리", "workbook,filesystem")
    items.Add Array("feature:NX-FILE-BATCH-RENAME", "feature", "NX-FILE-BATCH-RENAME", "파일 이름 일괄 변경", "선택한 파일에 이름 규칙을 순서대로 적용합니다. 예정 이름과 충돌을 확인한 뒤 바꾸며 마지막 작업 복구를 지원합니다.", "NX-GRP-FILE", "파일관리", "FileProperties", "active_workbook", "none", "dialog", "filesystem", "completion", "planned", "NX-FILE-BATCH-RENAME | 파일 이름 일괄 변경 | 파일관리", "filesystem")
    items.Add Array("feature:NX-TPL-LIST", "feature", "NX-TPL-LIST", "템플릿", "등록한 엑셀 양식을 검색하고 순서를 조정합니다. 파일·시트 등록, 상세 편집, 사용과 삭제를 제공합니다.", "NX-GRP-TEMPLATE", "템플릿", "PropertySheet", "active_workbook", "none", "dialog", "local_store", "completion", "fast", "NX-TPL-LIST | 템플릿", "template_store,workbook")
    items.Add Array("feature:NX-FILE-WORKBOOK-COMPARE", "feature", "NX-FILE-WORKBOOK-COMPARE", "범위 비교", "두 범위의 왼쪽 위 셀을 맞춰 비교하고 새 통합문서 두 시트에 차이를 연한 빨강으로 표시합니다. 원본은 변경하지 않습니다.", "NX-GRP-DATA", "데이터", "WindowsArrangeAll", "active_workbook", "none", "dialog", "none", "inline", "planned", "NX-FILE-WORKBOOK-COMPARE | 범위 비교 | 데이터", "workbook,range")
    items.Add Array("feature:NX-FILE-SHEET-COMPARE", "feature", "NX-FILE-SHEET-COMPARE", "시트 비교", "두 시트의 같은 셀 주소끼리 비교하고 새 통합문서 두 시트에 차이를 연한 빨강으로 표시합니다. 원본은 변경하지 않습니다.", "NX-GRP-DATA", "데이터", "WindowsArrangeAll", "active_workbook", "none", "dialog", "none", "inline", "planned", "NX-FILE-SHEET-COMPARE | 시트 비교 | 데이터", "workbook,range")
    items.Add Array("feature:NX-FILE-FILE-COMPARE", "feature", "NX-FILE-FILE-COMPARE", "파일 비교", "두 Excel 파일의 시트를 비교해 차이와 이동 링크가 있는 새 보고서를 엽니다. 원본은 읽기 전용으로 열며 결과 저장은 직접 선택합니다.", "NX-GRP-DATA", "데이터", "FileOpen", "active_workbook", "none", "dialog", "none", "inline", "planned", "NX-FILE-FILE-COMPARE | 파일 비교 | 데이터", "workbook,range")
    items.Add Array("feature:NX-DATA-NORMALIZE", "feature", "NX-DATA-NORMALIZE", "데이터 정규화", "공백·문자·숫자·날짜의 표현을 정리합니다. 결과 위치와 옵션을 고른 뒤 바로 실행하며 미리보기는 선택 사항입니다.", "NX-GRP-DATA", "데이터", "TextDirectionLeftToRight", "range_selection", "required", "dialog", "document", "completion", "fast", "NX-DATA-NORMALIZE | 데이터 정규화 | 데이터", "selection,worksheet")
    items.Add Array("feature:NX-DATA-UNIQUE-COUNT", "feature", "NX-DATA-UNIQUE-COUNT", "고유값 개수", "같은 값을 모아 등장 횟수와 비율을 계산합니다. 제목 행·대소문자·빈값 처리를 선택하고 별도 결과로 확인합니다.", "NX-GRP-DATA", "데이터", "NumberingGallery", "range_selection", "required", "dialog", "document", "completion", "fast", "NX-DATA-UNIQUE-COUNT | 고유값 개수 | 데이터", "selection,worksheet")
    items.Add Array("feature:NX-DATA-DUPLICATE-LIST", "feature", "NX-DATA-DUPLICATE-LIST", "중복 목록", "두 목록을 비교해 함께 있는 값이나 한쪽에만 있는 값을 찾습니다. 비교 기준과 결과 위치를 선택하며 원본은 유지합니다.", "NX-GRP-DATA", "데이터", "TableInsert", "range_selection", "required", "dialog", "document", "completion", "fast", "NX-DATA-DUPLICATE-LIST | 중복 목록 | 데이터", "selection,worksheet")
End Sub

Private Sub NxGeneratedNavigationAddItems02(ByRef items As Collection)
    items.Add Array("feature:NX-DATA-AGE", "feature", "NX-DATA-AGE", "만나이 계산", "생년월일과 기준일로 만 나이를 계산합니다. 만·세 표기와 결과 위치를 선택합니다. 기본값은 새 통합문서입니다.", "NX-GRP-DATA", "데이터", "DateAndTimeInsert", "range_selection", "required", "dialog", "document", "completion", "guarded", "NX-DATA-AGE | 만나이 계산 | 데이터", "selection,worksheet")
    items.Add Array("feature:NX-DATA-KOREAN-MONEY", "feature", "NX-DATA-KOREAN-MONEY", "한글 금액 변환", "숫자 금액을 한글 표현으로 만듭니다. 앞뒤 문구·숫자 혼합·0 처리와 결과 위치를 선택합니다.", "NX-GRP-DATA", "데이터", "AccountingFormat", "range_selection", "required", "dialog", "document", "completion", "guarded", "NX-DATA-KOREAN-MONEY | 한글 금액 변환 | 데이터", "selection,worksheet")
    items.Add Array("feature:NX-DATA-PRIVACY-MASK", "feature", "NX-DATA-PRIVACY-MASK", "개인정보 마스킹", "이름·연락처 등 선택한 유형의 일부 문자를 가립니다. 가림 방식과 결과 위치를 확인한 뒤 적용합니다.", "NX-GRP-DATA", "데이터", "ProtectDocument", "range_selection", "required", "dialog", "document", "completion", "guarded", "NX-DATA-PRIVACY-MASK | 개인정보 마스킹 | 데이터", "selection,worksheet")
    items.Add Array("command:NX-CMD-DATA-DATE-CONVERT", "command", "NX-CMD-DATA-DATE-CONVERT", "날짜 변환", "240617, 20240617, 점·하이픈 날짜, 주민번호형 생년월일을 날짜로 변환합니다. 표시 형식과 결과 위치를 선택하세요.", "NX-GRP-DATA", "데이터", "DateAndTimeInsert", "range_selection", "required", "dialog", "none", "inline", "fast", "NX-CMD-DATA-DATE-CONVERT | 날짜 변환 | 데이터", "worksheet,selection")
    items.Add Array("command:NX-CMD-SIZE-RB-EDIT-CELL-RESIZE", "command", "NX-CMD-SIZE-RB-EDIT-CELL-RESIZE", "셀·글자·도형 크기 조정", "열 너비·행 높이·글자 크기·범위 안 도형 배율 중 필요한 항목을 선택해 조정합니다. 셀 값은 유지하며 오류 시 이전 크기로 복원합니다.", "NX-GRP-CELL-FIT", "셀맞춤", "PropertySheet", "range_selection", "required", "dialog", "document", "completion", "guarded", "NX-CMD-SIZE-RB-EDIT-CELL-RESIZE | 셀·글자·도형 크기 조정 | 셀맞춤", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-EDIT-MEMO-ADD-LHEXCELFORMULA", "command", "NX-CMD-STYLE-RB-EDIT-MEMO-ADD-LHEXCELFORMULA", "셀수식 메모 기록", "기존 메모를 유지하고 수식을 추가하거나 덮어씁니다. 메모 표시 여부를 선택하고 마지막 기록을 되돌릴 수 있습니다.", "NX-GRP-FORMULA", "함수", "ReviewShowAllComments", "range_selection", "required", "dialog", "document", "inline", "guarded", "NX-CMD-STYLE-RB-EDIT-MEMO-ADD-LHEXCELFORMULA | 셀수식 메모 기록 | 함수", "selection")
    items.Add Array("command:NX-CMD-FORMULA-RB-FORMULA-PRECEDENTS-LIST", "command", "NX-CMD-FORMULA-RB-FORMULA-PRECEDENTS-LIST", "수식 참조표 만들기", "선택한 수식의 참조 위치와 값을 한글 표로 정리합니다. 새 시트 또는 새 통합문서를 선택하며, 확인하지 못한 참조는 확인 상태에 표시합니다.", "NX-GRP-FORMULA", "함수", "NameDefine", "range_selection", "required", "dialog", "document", "inline", "guarded", "NX-CMD-FORMULA-RB-FORMULA-PRECEDENTS-LIST | 수식 참조표 만들기 | 함수", "selection")
    items.Add Array("command:NX-CMD-FUNCTION-ROUND", "command", "NX-CMD-FUNCTION-ROUND", "ROUND 감싸기", "기존 값과 수식을 ROUND·ROUNDUP·ROUNDDOWN으로 감쌉니다. 오류는 그대로 유지합니다.", "NX-GRP-FORMULA", "함수", "FunctionWizard", "range_selection", "required", "dialog", "document", "inline", "guarded", "NX-CMD-FUNCTION-ROUND | ROUND 감싸기 | 함수", "selection")
    items.Add Array("command:NX-CMD-FUNCTION-IFERROR", "command", "NX-CMD-FUNCTION-IFERROR", "IFERROR 감싸기", "기존 수식이나 숫자를 IFERROR로 감싸 오류일 때 공란·숫자·문자를 반환합니다.", "NX-GRP-FORMULA", "함수", "FunctionWizard", "range_selection", "required", "dialog", "document", "inline", "guarded", "NX-CMD-FUNCTION-IFERROR | IFERROR 감싸기 | 함수", "selection")
    items.Add Array("command:NX-CMD-SHEET-RB-SHEET-SELECT-HOME", "command", "NX-CMD-SHEET-RB-SHEET-SELECT-HOME", "첫번째 시트로 이동", "문서의 첫 워크시트를 활성화합니다. 시트 탭 순서를 변경하는 명령과 구별하세요.", "NX-GRP-SHEET", "시트·보기", "SelectAll", "active_workbook", "none", "direct", "none", "inline", "fast", "NX-CMD-SHEET-RB-SHEET-SELECT-HOME | 첫번째 시트로 이동 | 시트·보기", "workbook,worksheet")
    items.Add Array("command:NX-CMD-SHEET-RB-SHEET-SELECT-END", "command", "NX-CMD-SHEET-RB-SHEET-SELECT-END", "마지막 시트로 이동", "문서의 마지막 워크시트를 활성화합니다. 시트 탭 순서를 변경하는 명령과 구별하세요.", "NX-GRP-SHEET", "시트·보기", "SelectAll", "active_workbook", "none", "direct", "none", "inline", "fast", "NX-CMD-SHEET-RB-SHEET-SELECT-END | 마지막 시트로 이동 | 시트·보기", "workbook,worksheet")
    items.Add Array("feature:NX-UTIL-DOCUMENT-NAVIGATOR", "feature", "NX-UTIL-DOCUMENT-NAVIGATOR", "문서·시트 탐색", "열린 문서와 시트를 찾아 전환합니다. 목록을 보기 위해 원본 문서에 보고서 시트를 추가하지 않습니다.", "NX-GRP-SHEET", "시트·보기", "FindDialog", "excel_ready", "none", "direct", "none", "inline", "fast", "NX-UTIL-DOCUMENT-NAVIGATOR | 문서·시트 탐색 | 시트·보기", "application")
    items.Add Array("feature:NX-FILE-SHEET-BATCH-RENAME", "feature", "NX-FILE-SHEET-BATCH-RENAME", "시트 이름 일괄 변경", "문서의 시트 이름에 규칙을 적용하거나 예정 이름을 직접 편집합니다. 이름 충돌을 확인한 뒤 적용하며 시트 내용은 유지합니다.", "NX-GRP-SHEET", "시트·보기", "NameDefine", "active_worksheet", "none", "dialog", "document", "completion", "planned", "NX-FILE-SHEET-BATCH-RENAME | 시트 이름 일괄 변경 | 시트·보기", "workbook,worksheet")
    items.Add Array("command:NX-CMD-VIEW-RB-WINDOWS-VIEW-FULLSCREEN", "command", "NX-CMD-VIEW-RB-WINDOWS-VIEW-FULLSCREEN", "전체 화면 켜기/끄기", "Excel의 전체 화면 표시를 전환합니다. 화면 표시만 바꾸며 문서 내용과 인쇄 설정은 유지합니다.", "NX-GRP-SHEET", "시트·보기", "ViewFullScreenView", "active_window", "none", "direct", "none", "inline", "fast", "NX-CMD-VIEW-RB-WINDOWS-VIEW-FULLSCREEN | 전체 화면 켜기/끄기 | 시트·보기", "window")
    items.Add Array("command:NX-CMD-VIEW-PRESET-1", "command", "NX-CMD-VIEW-PRESET-1", "화면 프리셋 1", "저장한 화면 배율을 적용합니다. 같은 프리셋을 다시 실행하면 이전 배율로 돌아갑니다.", "NX-GRP-SHEET", "시트·보기", "ViewNormalViewExcel", "active_window", "none", "direct", "none", "inline", "fast", "NX-CMD-VIEW-PRESET-1 | 화면 프리셋 1 | 시트·보기", "window")
    items.Add Array("command:NX-CMD-VIEW-PRESET-2", "command", "NX-CMD-VIEW-PRESET-2", "화면 프리셋 2", "저장한 화면 배율을 적용합니다. 같은 프리셋을 다시 실행하면 이전 배율로 돌아갑니다.", "NX-GRP-SHEET", "시트·보기", "ViewNormalViewExcel", "active_window", "none", "direct", "none", "inline", "fast", "NX-CMD-VIEW-PRESET-2 | 화면 프리셋 2 | 시트·보기", "window")
    items.Add Array("command:NX-CMD-VIEW-PRESET-3", "command", "NX-CMD-VIEW-PRESET-3", "화면 프리셋 3", "저장한 화면 배율을 적용합니다. 같은 프리셋을 다시 실행하면 이전 배율로 돌아갑니다.", "NX-GRP-SHEET", "시트·보기", "ViewNormalViewExcel", "active_window", "none", "direct", "none", "inline", "fast", "NX-CMD-VIEW-PRESET-3 | 화면 프리셋 3 | 시트·보기", "window")
    items.Add Array("command:NX-CMD-VIEW-PRESET-4", "command", "NX-CMD-VIEW-PRESET-4", "화면 프리셋 4", "저장한 화면 배율을 적용합니다. 같은 프리셋을 다시 실행하면 이전 배율로 돌아갑니다.", "NX-GRP-SHEET", "시트·보기", "ViewNormalViewExcel", "active_window", "none", "direct", "none", "inline", "fast", "NX-CMD-VIEW-PRESET-4 | 화면 프리셋 4 | 시트·보기", "window")
    items.Add Array("command:NX-CMD-VIEW-PRESET-5", "command", "NX-CMD-VIEW-PRESET-5", "화면 프리셋 5", "저장한 화면 배율을 적용합니다. 같은 프리셋을 다시 실행하면 이전 배율로 돌아갑니다.", "NX-GRP-SHEET", "시트·보기", "ViewNormalViewExcel", "active_window", "none", "direct", "none", "inline", "fast", "NX-CMD-VIEW-PRESET-5 | 화면 프리셋 5 | 시트·보기", "window")
    items.Add Array("command:NX-CMD-VIEW-PRESET-SETTINGS", "command", "NX-CMD-VIEW-PRESET-SETTINGS", "화면 프리셋 설정", "화면 프리셋 1~5의 배율을 10~400%로 지정하고 저장합니다.", "NX-GRP-SHEET", "시트·보기", "ViewNormalViewExcel", "active_window", "none", "dialog", "none", "inline", "fast", "NX-CMD-VIEW-PRESET-SETTINGS | 화면 프리셋 설정 | 시트·보기", "window")
    items.Add Array("feature:NX-DRAW-TITLE-TABLE", "feature", "NX-DRAW-TITLE-TABLE", "타이틀표 그리기", "선택 범위를 제목 영역으로 꾸밉니다. 테두리·채우기 옵션을 미리 보고 적용하며 셀 값은 유지합니다.", "NX-GRP-STYLE", "스타일", "BorderOutside", "range_selection", "required", "dialog", "document", "completion", "fast", "NX-DRAW-TITLE-TABLE | 타이틀표 그리기 | 스타일", "selection,worksheet")
    items.Add Array("feature:NX-DRAW-BUSINESS-TABLE", "feature", "NX-DRAW-BUSINESS-TABLE", "표 그리기", "선택 범위에 표의 테두리와 머리글 구성을 적용합니다. 범위를 다시 지정하고 원하는 모양을 확인한 뒤 실행합니다.", "NX-GRP-STYLE", "스타일", "TableAutoFormat", "range_selection", "required", "dialog", "document", "completion", "fast", "NX-DRAW-BUSINESS-TABLE | 표 그리기 | 스타일", "selection,worksheet")
    items.Add Array("feature:NX-DRAW-ROLE-STYLE", "feature", "NX-DRAW-ROLE-STYLE", "내엑셀 스타일", "제목·본문·합계 등 역할에 맞는 셀 서식을 선택합니다. 적용 축과 상대 위치를 확인하고 원본 값은 유지합니다.", "NX-GRP-STYLE", "스타일", "WordArtInsertGallery", "range_selection", "required", "dialog", "document", "completion", "fast", "NX-DRAW-ROLE-STYLE | 내엑셀 스타일 | 스타일", "selection,worksheet")
    items.Add Array("feature:NX-DRAW-CLEAR-INNER", "feature", "NX-DRAW-CLEAR-INNER", "내부선 제거", "선택 범위 안쪽의 가로·세로 테두리를 지웁니다. 범위 바깥 둘레의 테두리는 유지합니다.", "NX-GRP-STYLE", "스타일", "BorderOutside", "range_selection", "required", "direct", "document", "completion", "fast", "NX-DRAW-CLEAR-INNER | 내부선 제거 | 스타일", "selection,worksheet")
    items.Add Array("feature:NX-HANGUL-TABLE-SEND", "feature", "NX-HANGUL-TABLE-SEND", "아래한글 표 전송", "선택 범위를 아래한글에서 사용할 표로 전달합니다. 글꼴·글자 크기·표 모양 설정을 반영하며 원본 Excel 셀은 유지합니다.", "NX-GRP-STYLE", "스타일", "FileSendAsAttachment", "range_selection", "required", "direct", "external_document", "completion", "guarded", "NX-HANGUL-TABLE-SEND | 아래한글 표 전송 | 스타일", "selection,hangul_document")
    items.Add Array("feature:NX-HANGUL-PICTURE-SEND", "feature", "NX-HANGUL-PICTURE-SEND", "아래한글 그림 전송", "그림을 제목순 또는 촬영시간순으로 정렬해 아래한글 표의 셀 배경에 문서 포함 방식으로 넣습니다. 그림 셀 크기·쪽당 배치·제목행을 지정할 수 있습니다.", "NX-GRP-STYLE", "스타일", "PictureInsertFromFile", "selection_optional", "optional", "dialog", "external_document", "inline", "guarded", "NX-HANGUL-PICTURE-SEND | 아래한글 그림 전송 | 스타일", "file,hangul_document")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-RESETWORKBOOKSTYLES", "command", "NX-CMD-STYLE-RB-LHESTYLE-RESETWORKBOOKSTYLES", "사용자 스타일 정리", "현재 문서의 사용자 지정 셀 스타일을 삭제합니다. Excel 기본 스타일은 남으며, 적용 중인 사용자 스타일도 정리 대상입니다.", "NX-GRP-STYLE", "스타일", "DataFormDeleteRecord", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-RESETWORKBOOKSTYLES | 사용자 스타일 정리 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TITLE", "command", "NX-CMD-STYLE-RB-LHESTYLE-MONO-TITLE", "흑백 · 제목", "선택 범위에 흑백 계열의 문서의 큰 제목에 맞는 글자와 채우기 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "FontColorPicker", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-MONO-TITLE | 흑백 · 제목 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-SUBTITLE", "command", "NX-CMD-STYLE-RB-LHESTYLE-MONO-SUBTITLE", "흑백 · 소제목", "선택 범위에 흑백 계열의 제목 아래의 구분 문구에 맞는 글자와 채우기 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "BorderBottom", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-MONO-SUBTITLE | 흑백 · 소제목 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEHEADER", "command", "NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEHEADER", "흑백 · 표 머리글", "선택 범위에 흑백 계열의 표의 열 제목을 구분하는 글자·채우기·테두리 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "TableInsert", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEHEADER | 흑백 · 표 머리글 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEBODY", "command", "NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEBODY", "흑백 · 표 본문", "선택 범위에 흑백 계열의 표의 일반 내용 영역에 맞는 기본 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "TableAutoFormat", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEBODY | 흑백 · 표 본문 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-EMPHASISCELL", "command", "NX-CMD-STYLE-RB-LHESTYLE-MONO-EMPHASISCELL", "흑백 · 강조 셀", "선택 범위에 흑백 계열의 눈에 띄게 표시할 셀의 글자와 채우기 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "CellFillColorPicker", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-MONO-EMPHASISCELL | 흑백 · 강조 셀 | 스타일", "selection")
End Sub

Private Sub NxGeneratedNavigationAddItems03(ByRef items As Collection)
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TOTALROW", "command", "NX-CMD-STYLE-RB-LHESTYLE-MONO-TOTALROW", "흑백 · 합계 행", "선택 범위에 흑백 계열의 합계 영역을 구분하는 글자·채우기·테두리 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "AutoSum", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-MONO-TOTALROW | 흑백 · 합계 행 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TITLE", "command", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-TITLE", "컬러 · 제목", "선택 범위에 컬러 계열의 문서의 큰 제목에 맞는 글자와 채우기 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "FontColorPicker", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-TITLE | 컬러 · 제목 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-SUBTITLE", "command", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-SUBTITLE", "컬러 · 소제목", "선택 범위에 컬러 계열의 제목 아래의 구분 문구에 맞는 글자와 채우기 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "BorderBottom", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-SUBTITLE | 컬러 · 소제목 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEHEADER", "command", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEHEADER", "컬러 · 표 머리글", "선택 범위에 컬러 계열의 표의 열 제목을 구분하는 글자·채우기·테두리 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "TableInsert", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEHEADER | 컬러 · 표 머리글 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEBODY", "command", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEBODY", "컬러 · 표 본문", "선택 범위에 컬러 계열의 표의 일반 내용 영역에 맞는 기본 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "TableAutoFormat", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEBODY | 컬러 · 표 본문 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-EMPHASISCELL", "command", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-EMPHASISCELL", "컬러 · 강조 셀", "선택 범위에 컬러 계열의 눈에 띄게 표시할 셀의 글자와 채우기 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "CellFillColorPicker", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-EMPHASISCELL | 컬러 · 강조 셀 | 스타일", "selection")
    items.Add Array("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TOTALROW", "command", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-TOTALROW", "컬러 · 합계 행", "선택 범위에 컬러 계열의 합계 영역을 구분하는 글자·채우기·테두리 서식을 적용합니다. 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "AutoSum", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-STYLE-RB-LHESTYLE-COLOR-TOTALROW | 컬러 · 합계 행 | 스타일", "selection")
    items.Add Array("command:NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-PERCENTAGE", "command", "NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-PERCENTAGE", "백분율 표시 · 자릿수 선택", "선택한 숫자를 백분율로 표시합니다. 소수 자릿수를 0~15자리로 선택하며 실제 셀 값은 유지합니다.", "NX-GRP-STYLE", "스타일", "AccountingFormat", "range_selection", "required", "dialog", "document", "completion", "guarded", "NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-PERCENTAGE | 백분율 표시 · 자릿수 선택 | 스타일", "selection")
    items.Add Array("command:NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-DECIMAL", "command", "NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-DECIMAL", "숫자 표시 · 소수 자릿수 선택", "천 단위 구분과 소수 자릿수를 지정합니다. 0~15자리를 선택하며 실제 셀 값은 유지합니다.", "NX-GRP-STYLE", "스타일", "NumberingGallery", "range_selection", "required", "dialog", "document", "completion", "guarded", "NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-DECIMAL | 숫자 표시 · 소수 자릿수 선택 | 스타일", "selection")
    items.Add Array("command:NX-CMD-NUMBER-EMPHASIS", "command", "NX-CMD-NUMBER-EMPHASIS", "숫자 강조(▲▼)", "양수는 빨간색 ▲, 음수는 파란색 ▼로 표시합니다. 소수 자릿수와 백분율 표시를 선택하며 원래 값과 수식은 유지합니다.", "NX-GRP-STYLE", "스타일", "ConditionalFormattingHighlightCellsMenu", "range_selection", "required", "dialog", "document", "completion", "guarded", "NX-CMD-NUMBER-EMPHASIS | 숫자 강조(▲▼) | 스타일", "selection")
    items.Add Array("command:NX-CMD-ALIGN-RB-EDIT-ALIGN-CENTER-OVERCELLS", "command", "NX-CMD-ALIGN-RB-EDIT-ALIGN-CENTER-OVERCELLS", "모형병합 실행/취소", "셀을 실제로 합치지 않고 하나의 영역처럼 보이게 정렬합니다. 범위와 표시 방식을 선택하며 값 이동 여부를 확인한 뒤 적용합니다.", "NX-GRP-STYLE", "스타일", "AlignCenter", "range_selection", "required", "direct", "document", "completion", "guarded", "NX-CMD-ALIGN-RB-EDIT-ALIGN-CENTER-OVERCELLS | 모형병합 실행/취소 | 스타일", "selection")
    items.Add Array("feature:NX-DATA-FOCUS-CELL", "feature", "NX-DATA-FOCUS-CELL", "포커스셀", "현재 선택 범위의 행과 열을 강조해 위치를 쉽게 구분합니다. 내부망은 조건부서식 기반으로, DLL 환경은 오버레이로 표시합니다.", "NX-GRP-UTIL", "추가기능", "TableStyleBandedColumns", "selection_optional", "optional", "hybrid", "none", "inline", "fast", "NX-DATA-FOCUS-CELL | 포커스셀 | 추가기능", "selection,window")
    items.Add Array("feature:NX-UTIL-CALCULATOR", "feature", "NX-UTIL-CALCULATOR", "계산기", "숫자와 연산을 계산하고 결과를 복사하거나 선택 셀에 넣습니다. 선택 셀 값 불러오기와 최근 계산 기록을 제공합니다.", "NX-GRP-UTIL", "추가기능", "CalculateNow", "selection_optional", "optional", "direct", "none", "inline", "fast", "NX-UTIL-CALCULATOR | 계산기 | 추가기능", "application,selection")
    items.Add Array("feature:NX-UTIL-SYMBOLS", "feature", "NX-UTIL-SYMBOLS", "기호표", "기호를 골라 입력 문구를 구성하고 선택한 셀들에 넣습니다. 목록은 휠로 탐색하며 원래 셀 값은 입력 시 교체됩니다.", "NX-GRP-UTIL", "추가기능", "SymbolInsert", "selection_optional", "optional", "direct", "document", "completion", "fast", "NX-UTIL-SYMBOLS | 기호표 | 추가기능", "application,selection")
    items.Add Array("feature:NX-UTIL-NAVIGATOR", "feature", "NX-UTIL-NAVIGATOR", "탐색창", "내엑셀 기능을 분류별로 찾고 실행합니다. 내부망 기본 화면과 DLL 확장 화면은 같은 기능 카탈로그를 사용합니다.", "NX-GRP-UTIL", "추가기능", "ControlProperties", "active_workbook", "none", "direct", "none", "inline", "fast", "NX-UTIL-NAVIGATOR | 탐색창 | 추가기능", "application,workbook")
    items.Add Array("feature:NX-DATA-PRIVACY-SCAN", "feature", "NX-DATA-PRIVACY-SCAN", "개인정보 의심 항목 점검", "숨김 시트를 포함한 문서의 셀에서 개인정보로 의심되는 패턴과 위치를 찾습니다. 실제 외부 유출 여부를 판정하는 기능은 아닙니다.", "NX-GRP-INFO", "정보진단", "FileCheckOut", "active_workbook", "none", "direct", "none", "inline", "guarded", "NX-DATA-PRIVACY-SCAN | 개인정보 의심 항목 점검 | 정보진단", "workbook")
    items.Add Array("command:NX-CMD-INFO-RB-APP-NAMEUNHIDE", "command", "NX-CMD-INFO-RB-APP-NAMEUNHIDE", "숨김 이름 정의 표시", "현재 문서의 숨겨진 이름 정의를 이름 관리자에서 볼 수 있게 바꿉니다. 이름을 삭제하거나 참조식을 바꾸지는 않습니다.", "NX-GRP-INFO", "정보진단", "NameDefine", "active_workbook", "none", "direct", "document", "completion", "guarded", "NX-CMD-INFO-RB-APP-NAMEUNHIDE | 숨김 이름 정의 표시 | 정보진단", "workbook")
    items.Add Array("command:NX-CMD-INFO-RB-INFO-MEMO-LIST", "command", "NX-CMD-INFO-RB-INFO-MEMO-LIST", "문서 메모 목록 만들기", "현재 문서의 셀 메모 위치와 내용을 새 보고서 시트에 모읍니다. 원래 셀의 메모는 유지합니다.", "NX-GRP-INFO", "정보진단", "ReviewShowAllComments", "active_workbook", "none", "direct", "document", "completion", "guarded", "NX-CMD-INFO-RB-INFO-MEMO-LIST | 문서 메모 목록 만들기 | 정보진단", "workbook")
    items.Add Array("command:NX-CMD-INFO-RB-INFO-SHEET-LIST", "command", "NX-CMD-INFO-RB-INFO-SHEET-LIST", "문서 시트 목록 만들기", "현재 문서의 워크시트 이름과 표시·숨김 상태를 새 시트에 정리합니다. 탐색창과 달리 문서에 결과 시트가 추가됩니다.", "NX-GRP-INFO", "정보진단", "CellsInsertDialog", "active_workbook", "none", "direct", "document", "completion", "guarded", "NX-CMD-INFO-RB-INFO-SHEET-LIST | 문서 시트 목록 만들기 | 정보진단", "workbook")
End Sub

Private Function NxGeneratedNavigationRouteIndex() As Object
    Dim item As Variant
    Dim routeIndex As Object
    If mNavigationRouteIndex Is Nothing Then
        Set routeIndex = CreateObject("Scripting.Dictionary")
        routeIndex.CompareMode = vbBinaryCompare
        For Each item In NxGeneratedNavigationItems()
            routeIndex.Add CStr(item(0)), item
        Next item
        Set mNavigationRouteIndex = routeIndex
    End If
    Set NxGeneratedNavigationRouteIndex = mNavigationRouteIndex
End Function

Public Function NxGeneratedNavigationRouteExists(ByVal routeKey As String) As Boolean
    routeKey = NxCanonicalRouteKey(routeKey)
    NxGeneratedNavigationRouteExists = NxGeneratedNavigationRouteIndex().Exists(routeKey)
End Function

Public Function NxGeneratedNavigationRouteField(ByVal routeKey As String, ByVal fieldName As String) As String
    Dim item As Variant
    Dim routeIndex As Object
    Dim fieldIndex As Long
    routeKey = NxCanonicalRouteKey(routeKey)
    Select Case fieldName
        Case "route_key": fieldIndex = 0
        Case "item_type": fieldIndex = 1
        Case "id": fieldIndex = 2
        Case "label_ko": fieldIndex = 3
        Case "description_ko": fieldIndex = 4
        Case "category_id": fieldIndex = 5
        Case "category_label": fieldIndex = 6
        Case "image_mso": fieldIndex = 7
        Case "input_context": fieldIndex = 8
        Case "selection_mode": fieldIndex = 9
        Case "launch_surface": fieldIndex = 10
        Case "mutation_scope": fieldIndex = 11
        Case "feedback_mode": fieldIndex = 12
        Case "execution_grade": fieldIndex = 13
        Case "search_text": fieldIndex = 14
        Case "object_scope": fieldIndex = 15
        Case Else: NxRaiseContractError "Unknown navigation route field"
    End Select
    Set routeIndex = NxGeneratedNavigationRouteIndex()
    If Not routeIndex.Exists(routeKey) Then NxRaiseContractError "Unknown navigation route key"
    item = routeIndex.Item(routeKey)
    NxGeneratedNavigationRouteField = CStr(item(fieldIndex))
End Function

Public Function NxGeneratedNavigationMenuXml() As String
    Dim xml As String
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_SAVE()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_PRINT()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_COPY()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_INSERT()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_FILE()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_TEMPLATE()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_DATA()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_CELL_FIT()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_FORMULA()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_SHEET()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_STYLE()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_UTIL()
    xml = xml & NxGeneratedNavigationCategory_NX_GRP_INFO()
    NxGeneratedNavigationMenuXml = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_SAVE() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_SAVE"" label=""저장"" imageMso=""FileSave"">"
    xml = xml & "<button id=""nav_NX_FILE_MANNER_SAVE"" label=""매너 저장"" imageMso=""FileSave"" tag=""nx1|feature|NX-FILE-MANNER-SAVE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-MANNER-SAVE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-MANNER-SAVE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_SHEET_COPY_SAVE"" label=""현재 시트 사본 저장"" imageMso=""FileSaveAs"" tag=""nx1|feature|NX-FILE-SHEET-COPY-SAVE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-SHEET-COPY-SAVE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-SHEET-COPY-SAVE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_RANGE_COPY_SAVE"" label=""선택 범위 사본 저장"" imageMso=""FileSaveAs"" tag=""nx1|feature|NX-FILE-RANGE-COPY-SAVE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-RANGE-COPY-SAVE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-RANGE-COPY-SAVE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_PDF_CURRENT_SHEET"" label=""현재 시트 PDF 저장"" imageMso=""FileSaveAsPdfOrXps"" tag=""nx1|feature|NX-FILE-PDF-CURRENT-SHEET"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-PDF-CURRENT-SHEET") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-PDF-CURRENT-SHEET")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_PDF_ALL_COMBINED"" label=""문서 PDF 저장"" imageMso=""FileSaveAsPdfOrXps"" tag=""nx1|feature|NX-FILE-PDF-ALL-COMBINED"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-PDF-ALL-COMBINED") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-PDF-ALL-COMBINED")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_RANGE_PNG"" label=""선택 범위 그림 저장"" imageMso=""FileSaveAs"" tag=""nx1|feature|NX-FILE-RANGE-PNG"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-RANGE-PNG") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-RANGE-PNG")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_CHART_PNG"" label=""차트 그림 저장"" imageMso=""FileSaveAs"" tag=""nx1|feature|NX-FILE-CHART-PNG"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-CHART-PNG") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-CHART-PNG")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_SAVE = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_PRINT() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_PRINT"" label=""인쇄"" imageMso=""PrintPreviewAndPrint"">"
    xml = xml & "<button id=""nav_NX_CMD_PRINT_RB_PRINT_SETUP_QUICK"" label=""선택 범위 인쇄 미리보기"" imageMso=""PrintPreviewAndPrint"" tag=""nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_PRINT_RB_PRINT_SETUP_QUICK_PORTRAIT"" label=""세로 방향 · 너비 1쪽"" imageMso=""PrintPreviewAndPrint"" tag=""nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_PRINT_RB_PRINT_SETUP_QUICK_LANDSCAPE"" label=""가로 방향 · 너비 1쪽"" imageMso=""PrintPreviewAndPrint"" tag=""nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_PRINT_RB_PRINT_SETUP_QUICK_PORTRAIT_1BY1"" label=""세로 방향 · 한 장에 맞춤"" imageMso=""PrintPreviewAndPrint"" tag=""nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT-1BY1"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT-1BY1") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT-1BY1")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_PRINT_RB_PRINT_SETUP_QUICK_LANDSCAPE_1BY1"" label=""가로 방향 · 한 장에 맞춤"" imageMso=""PrintPreviewAndPrint"" tag=""nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE-1BY1"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE-1BY1") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE-1BY1")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_PRINT_RB_PRINT_SETUP_REPEAT"" label=""선택 행을 인쇄 제목으로"" imageMso=""PrintTitles"" tag=""nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-REPEAT"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-PRINT-RB-PRINT-SETUP-REPEAT") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-PRINT-RB-PRINT-SETUP-REPEAT")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_PRINT = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_COPY() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_COPY"" label=""복붙"" imageMso=""Copy"">"
    xml = xml & "<button id=""nav_NX_DATA_COPY_VISIBLE"" label=""보이는 셀만 복사"" imageMso=""Copy"" tag=""nx1|feature|NX-DATA-COPY-VISIBLE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-COPY-VISIBLE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-COPY-VISIBLE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_DATA_PASTE_VISIBLE_VALUES"" label=""보이는 셀에 값 붙여넣기"" imageMso=""Paste"" tag=""nx1|feature|NX-DATA-PASTE-VISIBLE-VALUES"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-PASTE-VISIBLE-VALUES") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-PASTE-VISIBLE-VALUES")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_CLIPBOARD_RB_CLIPBOARD_COPY_BYFORMULA"" label=""첫 셀 수식 복사"" imageMso=""Copy"" tag=""nx1|command|NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYFORMULA"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYFORMULA") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYFORMULA")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_CLIPBOARD_RB_CLIPBOARD_COPY_BYREFERENCE"" label=""첫 셀 주소 복사"" imageMso=""NameDefine"" tag=""nx1|command|NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYREFERENCE"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYREFERENCE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYREFERENCE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_CLIPBOARD_RB_CLIPBOARD_COPY_BYTEXT"" label=""보이는 값 복사"" imageMso=""Copy"" tag=""nx1|command|NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYTEXT"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYTEXT") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYTEXT")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_CLIPBOARD_RB_LHECLIPBOARD_PASTEVISIBLEFORMULAS"" label=""보이는 셀에 수식 붙여넣기"" imageMso=""PasteFormulas"" tag=""nx1|command|NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMULAS"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMULAS") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMULAS")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_CLIPBOARD_RB_LHECLIPBOARD_PASTEVISIBLEFORMATS"" label=""보이는 셀에 서식 붙여넣기"" imageMso=""PasteFormatting"" tag=""nx1|command|NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMATS"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMATS") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMATS")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_COPY = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_INSERT() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_INSERT"" label=""삽입"" imageMso=""PictureInsertFromFile"">"
    xml = xml & "<button id=""nav_NX_DRAW_INSERT_PICTURE"" label=""그림 삽입"" imageMso=""PictureInsertFromFile"" tag=""nx1|feature|NX-DRAW-INSERT-PICTURE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DRAW-INSERT-PICTURE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DRAW-INSERT-PICTURE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STRUCTURE_RB_EDIT_CELL_MAKEGROUP_HIDDEN_ROWCOLUMN"" label=""선택 행·열 그룹 만들기"" imageMso=""PropertySheet"" tag=""nx1|command|NX-CMD-STRUCTURE-RB-EDIT-CELL-MAKEGROUP-HIDDEN-ROWCOLUMN"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STRUCTURE-RB-EDIT-CELL-MAKEGROUP-HIDDEN-ROWCOLUMN") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STRUCTURE-RB-EDIT-CELL-MAKEGROUP-HIDDEN-ROWCOLUMN")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_INSERT = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_FILE() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_FILE"" label=""파일관리"" imageMso=""FileOpen"">"
    xml = xml & "<button id=""nav_NX_FILE_FOLDER_CREATE"" label=""폴더 일괄 생성"" imageMso=""Folder"" tag=""nx1|feature|NX-FILE-FOLDER-CREATE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-FOLDER-CREATE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-FOLDER-CREATE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_CONSOLIDATE"" label=""파일 통합"" imageMso=""FileOpen"" tag=""nx1|feature|NX-FILE-CONSOLIDATE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-CONSOLIDATE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-CONSOLIDATE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_BATCH_RENAME"" label=""파일 이름 일괄 변경"" imageMso=""FileProperties"" tag=""nx1|feature|NX-FILE-BATCH-RENAME"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-BATCH-RENAME") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-BATCH-RENAME")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_FILE = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_TEMPLATE() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_TEMPLATE"" label=""템플릿"" imageMso=""FileNew"">"
    xml = xml & "<button id=""nav_NX_TPL_LIST"" label=""템플릿"" imageMso=""PropertySheet"" tag=""nx1|feature|NX-TPL-LIST"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-TPL-LIST") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-TPL-LIST")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_TEMPLATE = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_DATA() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_DATA"" label=""데이터"" imageMso=""TableInsert"">"
    xml = xml & "<menu id=""nav_grp_NX_NAV_GRP_DATA_COMPARE"" label=""데이터 비교"" imageMso=""WindowsArrangeAll"">"
    xml = xml & "<button id=""nav_NX_FILE_WORKBOOK_COMPARE"" label=""범위 비교"" imageMso=""WindowsArrangeAll"" tag=""nx1|feature|NX-FILE-WORKBOOK-COMPARE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-WORKBOOK-COMPARE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-WORKBOOK-COMPARE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_SHEET_COMPARE"" label=""시트 비교"" imageMso=""WindowsArrangeAll"" tag=""nx1|feature|NX-FILE-SHEET-COMPARE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-SHEET-COMPARE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-SHEET-COMPARE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_FILE_COMPARE"" label=""파일 비교"" imageMso=""FileOpen"" tag=""nx1|feature|NX-FILE-FILE-COMPARE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-FILE-COMPARE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-FILE-COMPARE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "</menu>"
    xml = xml & "<button id=""nav_NX_DATA_NORMALIZE"" label=""데이터 정규화"" imageMso=""TextDirectionLeftToRight"" tag=""nx1|feature|NX-DATA-NORMALIZE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-NORMALIZE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-NORMALIZE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_DATA_UNIQUE_COUNT"" label=""고유값 개수"" imageMso=""NumberingGallery"" tag=""nx1|feature|NX-DATA-UNIQUE-COUNT"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-UNIQUE-COUNT") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-UNIQUE-COUNT")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_DATA_DUPLICATE_LIST"" label=""중복 목록"" imageMso=""TableInsert"" tag=""nx1|feature|NX-DATA-DUPLICATE-LIST"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-DUPLICATE-LIST") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-DUPLICATE-LIST")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_DATA_AGE"" label=""만나이 계산"" imageMso=""DateAndTimeInsert"" tag=""nx1|feature|NX-DATA-AGE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-AGE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-AGE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_DATA_KOREAN_MONEY"" label=""한글 금액 변환"" imageMso=""AccountingFormat"" tag=""nx1|feature|NX-DATA-KOREAN-MONEY"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-KOREAN-MONEY") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-KOREAN-MONEY")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_DATA_PRIVACY_MASK"" label=""개인정보 마스킹"" imageMso=""ProtectDocument"" tag=""nx1|feature|NX-DATA-PRIVACY-MASK"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-PRIVACY-MASK") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-PRIVACY-MASK")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_DATA_DATE_CONVERT"" label=""날짜 변환"" imageMso=""DateAndTimeInsert"" tag=""nx1|command|NX-CMD-DATA-DATE-CONVERT"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-DATA-DATE-CONVERT") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-DATA-DATE-CONVERT")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_DATA = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_CELL_FIT() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_CELL_FIT"" label=""셀맞춤"" imageMso=""RowHeight"">"
    xml = xml & "<button id=""nav_NX_CMD_SIZE_RB_EDIT_CELL_RESIZE"" label=""셀·글자·도형 크기 조정"" imageMso=""PropertySheet"" tag=""nx1|command|NX-CMD-SIZE-RB-EDIT-CELL-RESIZE"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-SIZE-RB-EDIT-CELL-RESIZE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-SIZE-RB-EDIT-CELL-RESIZE")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_CELL_FIT = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_FORMULA() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_FORMULA"" label=""함수"" imageMso=""FunctionWizard"">"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_EDIT_MEMO_ADD_LHEXCELFORMULA"" label=""셀수식 메모 기록"" imageMso=""ReviewShowAllComments"" tag=""nx1|command|NX-CMD-STYLE-RB-EDIT-MEMO-ADD-LHEXCELFORMULA"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-EDIT-MEMO-ADD-LHEXCELFORMULA") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-EDIT-MEMO-ADD-LHEXCELFORMULA")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_FORMULA_RB_FORMULA_PRECEDENTS_LIST"" label=""수식 참조표 만들기"" imageMso=""NameDefine"" tag=""nx1|command|NX-CMD-FORMULA-RB-FORMULA-PRECEDENTS-LIST"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-FORMULA-RB-FORMULA-PRECEDENTS-LIST") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-FORMULA-RB-FORMULA-PRECEDENTS-LIST")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_FUNCTION_ROUND"" label=""ROUND 감싸기"" imageMso=""FunctionWizard"" tag=""nx1|command|NX-CMD-FUNCTION-ROUND"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-FUNCTION-ROUND") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-FUNCTION-ROUND")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_FUNCTION_IFERROR"" label=""IFERROR 감싸기"" imageMso=""FunctionWizard"" tag=""nx1|command|NX-CMD-FUNCTION-IFERROR"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-FUNCTION-IFERROR") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-FUNCTION-IFERROR")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_FORMULA = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_SHEET() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_SHEET"" label=""시트·보기"" imageMso=""ViewNormalViewExcel"">"
    xml = xml & "<button id=""nav_NX_CMD_SHEET_RB_SHEET_SELECT_HOME"" label=""첫번째 시트로 이동"" imageMso=""SelectAll"" tag=""nx1|command|NX-CMD-SHEET-RB-SHEET-SELECT-HOME"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-SHEET-RB-SHEET-SELECT-HOME") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-SHEET-RB-SHEET-SELECT-HOME")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_SHEET_RB_SHEET_SELECT_END"" label=""마지막 시트로 이동"" imageMso=""SelectAll"" tag=""nx1|command|NX-CMD-SHEET-RB-SHEET-SELECT-END"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-SHEET-RB-SHEET-SELECT-END") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-SHEET-RB-SHEET-SELECT-END")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_UTIL_DOCUMENT_NAVIGATOR"" label=""문서·시트 탐색"" imageMso=""FindDialog"" tag=""nx1|feature|NX-UTIL-DOCUMENT-NAVIGATOR"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-UTIL-DOCUMENT-NAVIGATOR") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-UTIL-DOCUMENT-NAVIGATOR")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_FILE_SHEET_BATCH_RENAME"" label=""시트 이름 일괄 변경"" imageMso=""NameDefine"" tag=""nx1|feature|NX-FILE-SHEET-BATCH-RENAME"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-FILE-SHEET-BATCH-RENAME") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-FILE-SHEET-BATCH-RENAME")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_VIEW_RB_WINDOWS_VIEW_FULLSCREEN"" label=""전체 화면 켜기/끄기"" imageMso=""ViewFullScreenView"" tag=""nx1|command|NX-CMD-VIEW-RB-WINDOWS-VIEW-FULLSCREEN"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-VIEW-RB-WINDOWS-VIEW-FULLSCREEN") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-VIEW-RB-WINDOWS-VIEW-FULLSCREEN")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<menu id=""nav_grp_NX_CMD_GRP_VIEW_WINDOW"" label=""화면 프리셋"" imageMso=""ViewNormalViewExcel"">"
    xml = xml & "<button id=""nav_NX_CMD_VIEW_PRESET_1"" label=""화면 프리셋 1"" imageMso=""ViewNormalViewExcel"" tag=""nx1|command|NX-CMD-VIEW-PRESET-1"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-VIEW-PRESET-1") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-VIEW-PRESET-1")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_VIEW_PRESET_2"" label=""화면 프리셋 2"" imageMso=""ViewNormalViewExcel"" tag=""nx1|command|NX-CMD-VIEW-PRESET-2"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-VIEW-PRESET-2") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-VIEW-PRESET-2")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_VIEW_PRESET_3"" label=""화면 프리셋 3"" imageMso=""ViewNormalViewExcel"" tag=""nx1|command|NX-CMD-VIEW-PRESET-3"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-VIEW-PRESET-3") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-VIEW-PRESET-3")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_VIEW_PRESET_4"" label=""화면 프리셋 4"" imageMso=""ViewNormalViewExcel"" tag=""nx1|command|NX-CMD-VIEW-PRESET-4"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-VIEW-PRESET-4") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-VIEW-PRESET-4")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_VIEW_PRESET_5"" label=""화면 프리셋 5"" imageMso=""ViewNormalViewExcel"" tag=""nx1|command|NX-CMD-VIEW-PRESET-5"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-VIEW-PRESET-5") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-VIEW-PRESET-5")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_VIEW_PRESET_SETTINGS"" label=""화면 프리셋 설정"" imageMso=""ViewNormalViewExcel"" tag=""nx1|command|NX-CMD-VIEW-PRESET-SETTINGS"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-VIEW-PRESET-SETTINGS") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-VIEW-PRESET-SETTINGS")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "</menu>"
    NxGeneratedNavigationCategory_NX_GRP_SHEET = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_STYLE() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_STYLE"" label=""스타일"" imageMso=""TableAutoFormat"">"
    xml = xml & "<button id=""nav_NX_DRAW_TITLE_TABLE"" label=""타이틀표 그리기"" imageMso=""BorderOutside"" tag=""nx1|feature|NX-DRAW-TITLE-TABLE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DRAW-TITLE-TABLE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DRAW-TITLE-TABLE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_DRAW_BUSINESS_TABLE"" label=""표 그리기"" imageMso=""TableAutoFormat"" tag=""nx1|feature|NX-DRAW-BUSINESS-TABLE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DRAW-BUSINESS-TABLE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DRAW-BUSINESS-TABLE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_DRAW_ROLE_STYLE"" label=""내엑셀 스타일"" imageMso=""WordArtInsertGallery"" tag=""nx1|feature|NX-DRAW-ROLE-STYLE"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DRAW-ROLE-STYLE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DRAW-ROLE-STYLE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_DRAW_CLEAR_INNER"" label=""내부선 제거"" imageMso=""BorderOutside"" tag=""nx1|feature|NX-DRAW-CLEAR-INNER"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DRAW-CLEAR-INNER") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DRAW-CLEAR-INNER")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_HANGUL_TABLE_SEND"" label=""아래한글 표 전송"" imageMso=""FileSendAsAttachment"" tag=""nx1|feature|NX-HANGUL-TABLE-SEND"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-HANGUL-TABLE-SEND") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-HANGUL-TABLE-SEND")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_HANGUL_PICTURE_SEND"" label=""아래한글 그림 전송"" imageMso=""PictureInsertFromFile"" tag=""nx1|feature|NX-HANGUL-PICTURE-SEND"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-HANGUL-PICTURE-SEND") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-HANGUL-PICTURE-SEND")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<menu id=""nav_grp_NX_CMD_GRP_STYLE"" label=""스타일 조정"" imageMso=""DataFormDeleteRecord"">"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_RESETWORKBOOKSTYLES"" label=""사용자 스타일 정리"" imageMso=""DataFormDeleteRecord"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-RESETWORKBOOKSTYLES"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-RESETWORKBOOKSTYLES") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-RESETWORKBOOKSTYLES")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_MONO_TITLE"" label=""흑백 · 제목"" imageMso=""FontColorPicker"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-TITLE"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TITLE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TITLE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_MONO_SUBTITLE"" label=""흑백 · 소제목"" imageMso=""BorderBottom"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-SUBTITLE"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-SUBTITLE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-SUBTITLE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_MONO_TABLEHEADER"" label=""흑백 · 표 머리글"" imageMso=""TableInsert"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEHEADER"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEHEADER") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEHEADER")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_MONO_TABLEBODY"" label=""흑백 · 표 본문"" imageMso=""TableAutoFormat"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEBODY"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEBODY") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEBODY")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_MONO_EMPHASISCELL"" label=""흑백 · 강조 셀"" imageMso=""CellFillColorPicker"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-EMPHASISCELL"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-EMPHASISCELL") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-EMPHASISCELL")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_MONO_TOTALROW"" label=""흑백 · 합계 행"" imageMso=""AutoSum"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-TOTALROW"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TOTALROW") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-MONO-TOTALROW")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_COLOR_TITLE"" label=""컬러 · 제목"" imageMso=""FontColorPicker"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-TITLE"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TITLE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TITLE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_COLOR_SUBTITLE"" label=""컬러 · 소제목"" imageMso=""BorderBottom"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-SUBTITLE"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-SUBTITLE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-SUBTITLE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_COLOR_TABLEHEADER"" label=""컬러 · 표 머리글"" imageMso=""TableInsert"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEHEADER"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEHEADER") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEHEADER")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_COLOR_TABLEBODY"" label=""컬러 · 표 본문"" imageMso=""TableAutoFormat"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEBODY"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEBODY") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEBODY")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_COLOR_EMPHASISCELL"" label=""컬러 · 강조 셀"" imageMso=""CellFillColorPicker"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-EMPHASISCELL"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-EMPHASISCELL") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-EMPHASISCELL")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_STYLE_RB_LHESTYLE_COLOR_TOTALROW"" label=""컬러 · 합계 행"" imageMso=""AutoSum"" tag=""nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-TOTALROW"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TOTALROW") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-STYLE-RB-LHESTYLE-COLOR-TOTALROW")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "</menu>"
    xml = xml & "<menu id=""nav_grp_NX_CMD_GRP_NUMBER_FORMAT"" label=""표시 형식"" imageMso=""AccountingFormat"">"
    xml = xml & "<button id=""nav_NX_CMD_NUMBER_RB_EDIT_NUMBERFORMAT_PERCENTAGE"" label=""백분율 표시 · 자릿수 선택"" imageMso=""AccountingFormat"" tag=""nx1|command|NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-PERCENTAGE"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-PERCENTAGE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-PERCENTAGE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_NUMBER_RB_EDIT_NUMBERFORMAT_DECIMAL"" label=""숫자 표시 · 소수 자릿수 선택"" imageMso=""NumberingGallery"" tag=""nx1|command|NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-DECIMAL"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-DECIMAL") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-DECIMAL")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_NUMBER_EMPHASIS"" label=""숫자 강조(▲▼)"" imageMso=""ConditionalFormattingHighlightCellsMenu"" tag=""nx1|command|NX-CMD-NUMBER-EMPHASIS"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-NUMBER-EMPHASIS") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-NUMBER-EMPHASIS")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "</menu>"
    xml = xml & "<menu id=""nav_grp_NX_CMD_GRP_ALIGNMENT_MERGE"" label=""정렬·병합"" imageMso=""AlignCenter"">"
    xml = xml & "<button id=""nav_NX_CMD_ALIGN_RB_EDIT_ALIGN_CENTER_OVERCELLS"" label=""모형병합 실행/취소"" imageMso=""AlignCenter"" tag=""nx1|command|NX-CMD-ALIGN-RB-EDIT-ALIGN-CENTER-OVERCELLS"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-ALIGN-RB-EDIT-ALIGN-CENTER-OVERCELLS") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-ALIGN-RB-EDIT-ALIGN-CENTER-OVERCELLS")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "</menu>"
    NxGeneratedNavigationCategory_NX_GRP_STYLE = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_UTIL() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_UTIL"" label=""추가기능"" imageMso=""AddInManager"">"
    xml = xml & "<button id=""nav_NX_DATA_FOCUS_CELL"" label=""포커스셀"" imageMso=""TableStyleBandedColumns"" tag=""nx1|feature|NX-DATA-FOCUS-CELL"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-FOCUS-CELL") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-FOCUS-CELL")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_UTIL_CALCULATOR"" label=""계산기"" imageMso=""CalculateNow"" tag=""nx1|feature|NX-UTIL-CALCULATOR"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-UTIL-CALCULATOR") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-UTIL-CALCULATOR")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_UTIL_SYMBOLS"" label=""기호표"" imageMso=""SymbolInsert"" tag=""nx1|feature|NX-UTIL-SYMBOLS"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-UTIL-SYMBOLS") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-UTIL-SYMBOLS")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_UTIL_NAVIGATOR"" label=""탐색창"" imageMso=""ControlProperties"" tag=""nx1|feature|NX-UTIL-NAVIGATOR"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-UTIL-NAVIGATOR") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-UTIL-NAVIGATOR")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_UTIL = xml & "</menu>"
End Function

Private Function NxGeneratedNavigationCategory_NX_GRP_INFO() As String
    Dim xml As String
    xml = "<menu id=""nav_cat_NX_GRP_INFO"" label=""정보진단"" imageMso=""Info"">"
    xml = xml & "<button id=""nav_NX_DATA_PRIVACY_SCAN"" label=""개인정보 의심 항목 점검"" imageMso=""FileCheckOut"" tag=""nx1|feature|NX-DATA-PRIVACY-SCAN"" enabled=""" & NxRouteAvailabilityRibbonValue("feature:NX-DATA-PRIVACY-SCAN") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("feature:NX-DATA-PRIVACY-SCAN")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_INFO_RB_APP_NAMEUNHIDE"" label=""숨김 이름 정의 표시"" imageMso=""NameDefine"" tag=""nx1|command|NX-CMD-INFO-RB-APP-NAMEUNHIDE"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-INFO-RB-APP-NAMEUNHIDE") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-INFO-RB-APP-NAMEUNHIDE")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_INFO_RB_INFO_MEMO_LIST"" label=""문서 메모 목록 만들기"" imageMso=""ReviewShowAllComments"" tag=""nx1|command|NX-CMD-INFO-RB-INFO-MEMO-LIST"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-INFO-RB-INFO-MEMO-LIST") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-INFO-RB-INFO-MEMO-LIST")) & """ onAction=""NxRibbonExecute""/>"
    xml = xml & "<button id=""nav_NX_CMD_INFO_RB_INFO_SHEET_LIST"" label=""문서 시트 목록 만들기"" imageMso=""CellsInsertDialog"" tag=""nx1|command|NX-CMD-INFO-RB-INFO-SHEET-LIST"" enabled=""" & NxRouteAvailabilityRibbonValue("command:NX-CMD-INFO-RB-INFO-SHEET-LIST") & """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip("command:NX-CMD-INFO-RB-INFO-SHEET-LIST")) & """ onAction=""NxRibbonExecute""/>"
    NxGeneratedNavigationCategory_NX_GRP_INFO = xml & "</menu>"
End Function

Public Function NxNavigationDisplayText(ByVal token As String) As String
    Select Case token
        Case "excel_ready": NxNavigationDisplayText = "선택 없이 사용 가능"
        Case "active_data_region": NxNavigationDisplayText = "현재 데이터 표"
        Case "active_window": NxNavigationDisplayText = "현재 Excel 창"
        Case "active_workbook": NxNavigationDisplayText = "현재 통합문서"
        Case "active_worksheet": NxNavigationDisplayText = "현재 시트"
        Case "range_or_shape": NxNavigationDisplayText = "선택 셀 또는 그림"
        Case "range_selection": NxNavigationDisplayText = "선택한 셀"
        Case "selection_optional": NxNavigationDisplayText = "선택 없이 사용 가능"
        Case "selection_or_clipboard": NxNavigationDisplayText = "선택 셀 또는 복사한 내용"
        Case "shape_selection": NxNavigationDisplayText = "선택한 그림"
        Case "fast": NxNavigationDisplayText = "바로 실행"
        Case "guarded": NxNavigationDisplayText = "확인 후 실행"
        Case "planned": NxNavigationDisplayText = "계획 확인"
        Case Else: NxNavigationDisplayText = "입력 조건 확인"
    End Select
End Function


Public Function NxGeneratedBrandImage(ByVal tag As String) As String
    Select Case tag
        Case "nx1|command|NX-CMD-ALIGN-RB-EDIT-ALIGN-CENTER-OVERCELLS": NxGeneratedBrandImage = "NxBrand_merge_16"
        Case "nx1|command|NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYFORMULA": NxGeneratedBrandImage = "NxBrand_copy_formula_16"
        Case "nx1|command|NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYREFERENCE": NxGeneratedBrandImage = "NxBrand_copy_address_16"
        Case "nx1|command|NX-CMD-CLIPBOARD-RB-CLIPBOARD-COPY-BYTEXT": NxGeneratedBrandImage = "NxBrand_copy_values_16"
        Case "nx1|command|NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMATS": NxGeneratedBrandImage = "NxBrand_paste_format_16"
        Case "nx1|command|NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEFORMULAS": NxGeneratedBrandImage = "NxBrand_paste_formula_16"
        Case "nx1|command|NX-CMD-DATA-DATE-CONVERT": NxGeneratedBrandImage = "NxBrand_date_16"
        Case "nx1|command|NX-CMD-FORMULA-RB-FORMULA-PRECEDENTS-LIST": NxGeneratedBrandImage = "NxBrand_formula_map_16"
        Case "nx1|command|NX-CMD-FUNCTION-IFERROR": NxGeneratedBrandImage = "NxBrand_iferror_16"
        Case "nx1|command|NX-CMD-FUNCTION-ROUND": NxGeneratedBrandImage = "NxBrand_round_16"
        Case "nx1|command|NX-CMD-INFO-RB-APP-NAMEUNHIDE": NxGeneratedBrandImage = "NxBrand_names_16"
        Case "nx1|command|NX-CMD-INFO-RB-INFO-MEMO-LIST": NxGeneratedBrandImage = "NxBrand_notes_16"
        Case "nx1|command|NX-CMD-INFO-RB-INFO-SHEET-LIST": NxGeneratedBrandImage = "NxBrand_sheet_list_16"
        Case "nx1|command|NX-CMD-NUMBER-EMPHASIS": NxGeneratedBrandImage = "NxBrand_number_emphasis_16"
        Case "nx1|command|NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-DECIMAL": NxGeneratedBrandImage = "NxBrand_decimal_16"
        Case "nx1|command|NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-NUMBER": NxGeneratedBrandImage = "NxBrand_comma_16"
        Case "nx1|command|NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-PERCENTAGE": NxGeneratedBrandImage = "NxBrand_percent_16"
        Case "nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK": NxGeneratedBrandImage = "NxBrand_print_16"
        Case "nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE": NxGeneratedBrandImage = "NxBrand_landscape_16"
        Case "nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-LANDSCAPE-1BY1": NxGeneratedBrandImage = "NxBrand_fit_page_16"
        Case "nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT": NxGeneratedBrandImage = "NxBrand_portrait_16"
        Case "nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-QUICK-PORTRAIT-1BY1": NxGeneratedBrandImage = "NxBrand_fit_page_16"
        Case "nx1|command|NX-CMD-PRINT-RB-PRINT-SETUP-REPEAT": NxGeneratedBrandImage = "NxBrand_print_title_16"
        Case "nx1|command|NX-CMD-SHEET-RB-SHEET-SELECT-END": NxGeneratedBrandImage = "NxBrand_last_sheet_16"
        Case "nx1|command|NX-CMD-SHEET-RB-SHEET-SELECT-HOME": NxGeneratedBrandImage = "NxBrand_first_sheet_16"
        Case "nx1|command|NX-CMD-SIZE-RB-EDIT-CELL-RESIZE": NxGeneratedBrandImage = "NxBrand_resize_16"
        Case "nx1|command|NX-CMD-STRUCTURE-RB-EDIT-CELL-MAKEGROUP-HIDDEN-ROWCOLUMN": NxGeneratedBrandImage = "NxBrand_group_16"
        Case "nx1|command|NX-CMD-STYLE-FILL-GRAY": NxGeneratedBrandImage = "NxBrand_fill_16"
        Case "nx1|command|NX-CMD-STYLE-FILL-LIGHT-GREEN": NxGeneratedBrandImage = "NxBrand_fill_16"
        Case "nx1|command|NX-CMD-STYLE-FILL-LIGHT-RED": NxGeneratedBrandImage = "NxBrand_fill_16"
        Case "nx1|command|NX-CMD-STYLE-FILL-LIGHT-YELLOW": NxGeneratedBrandImage = "NxBrand_fill_16"
        Case "nx1|command|NX-CMD-STYLE-FONT-COLOR-BLACK": NxGeneratedBrandImage = "NxBrand_font_color_16"
        Case "nx1|command|NX-CMD-STYLE-FONT-COLOR-BLUE": NxGeneratedBrandImage = "NxBrand_font_color_16"
        Case "nx1|command|NX-CMD-STYLE-FONT-COLOR-GREEN": NxGeneratedBrandImage = "NxBrand_font_color_16"
        Case "nx1|command|NX-CMD-STYLE-FONT-COLOR-RED": NxGeneratedBrandImage = "NxBrand_font_color_16"
        Case "nx1|command|NX-CMD-STYLE-FONT-SIZE-10": NxGeneratedBrandImage = "NxBrand_type_10_16"
        Case "nx1|command|NX-CMD-STYLE-FONT-SIZE-12": NxGeneratedBrandImage = "NxBrand_type_12_16"
        Case "nx1|command|NX-CMD-STYLE-FONT-SIZE-15": NxGeneratedBrandImage = "NxBrand_type_15_16"
        Case "nx1|command|NX-CMD-STYLE-FONT-SIZE-9": NxGeneratedBrandImage = "NxBrand_type_9_16"
        Case "nx1|command|NX-CMD-STYLE-RB-EDIT-MEMO-ADD-LHEXCELFORMULA": NxGeneratedBrandImage = "NxBrand_formula_note_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-EMPHASISCELL": NxGeneratedBrandImage = "NxBrand_role_emphasis_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-SUBTITLE": NxGeneratedBrandImage = "NxBrand_role_subtitle_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEBODY": NxGeneratedBrandImage = "NxBrand_role_body_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-TABLEHEADER": NxGeneratedBrandImage = "NxBrand_role_header_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-TITLE": NxGeneratedBrandImage = "NxBrand_role_title_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-COLOR-TOTALROW": NxGeneratedBrandImage = "NxBrand_role_total_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-EMPHASISCELL": NxGeneratedBrandImage = "NxBrand_role_emphasis_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-SUBTITLE": NxGeneratedBrandImage = "NxBrand_role_subtitle_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEBODY": NxGeneratedBrandImage = "NxBrand_role_body_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-TABLEHEADER": NxGeneratedBrandImage = "NxBrand_role_header_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-TITLE": NxGeneratedBrandImage = "NxBrand_role_title_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-MONO-TOTALROW": NxGeneratedBrandImage = "NxBrand_role_total_16"
        Case "nx1|command|NX-CMD-STYLE-RB-LHESTYLE-RESETWORKBOOKSTYLES": NxGeneratedBrandImage = "NxBrand_style_clean_16"
        Case "nx1|command|NX-CMD-VIEW-PRESET-1": NxGeneratedBrandImage = "NxBrand_view_1_16"
        Case "nx1|command|NX-CMD-VIEW-PRESET-2": NxGeneratedBrandImage = "NxBrand_view_2_16"
        Case "nx1|command|NX-CMD-VIEW-PRESET-3": NxGeneratedBrandImage = "NxBrand_view_3_16"
        Case "nx1|command|NX-CMD-VIEW-PRESET-4": NxGeneratedBrandImage = "NxBrand_view_4_16"
        Case "nx1|command|NX-CMD-VIEW-PRESET-5": NxGeneratedBrandImage = "NxBrand_view_5_16"
        Case "nx1|command|NX-CMD-VIEW-PRESET-SETTINGS": NxGeneratedBrandImage = "NxBrand_settings_16"
        Case "nx1|command|NX-CMD-VIEW-RB-WINDOWS-VIEW-FULLSCREEN": NxGeneratedBrandImage = "NxBrand_fullscreen_16"
        Case "nx1|entry|NX-ENTRY-AI": NxGeneratedBrandImage = "NxBrand_ai_16"
        Case "nx1|entry|NX-MGMT-SAVED-FOLDER": NxGeneratedBrandImage = "NxBrand_folder_16"
        Case "nx1|entry|NX-MGMT-SHORTCUTS": NxGeneratedBrandImage = "NxBrand_shortcuts_16"
        Case "nx1|feature|NX-DATA-AGE": NxGeneratedBrandImage = "NxBrand_age_16"
        Case "nx1|feature|NX-DATA-COPY-VISIBLE": NxGeneratedBrandImage = "NxBrand_copy_visible_16"
        Case "nx1|feature|NX-DATA-DUPLICATE-LIST": NxGeneratedBrandImage = "NxBrand_duplicates_16"
        Case "nx1|feature|NX-DATA-FOCUS-CELL": NxGeneratedBrandImage = "NxBrand_focus_16"
        Case "nx1|feature|NX-DATA-KOREAN-MONEY": NxGeneratedBrandImage = "NxBrand_currency_16"
        Case "nx1|feature|NX-DATA-NORMALIZE": NxGeneratedBrandImage = "NxBrand_normalize_16"
        Case "nx1|feature|NX-DATA-PASTE-VISIBLE-VALUES": NxGeneratedBrandImage = "NxBrand_paste_values_16"
        Case "nx1|feature|NX-DATA-PRIVACY-MASK": NxGeneratedBrandImage = "NxBrand_mask_16"
        Case "nx1|feature|NX-DATA-PRIVACY-SCAN": NxGeneratedBrandImage = "NxBrand_privacy_16"
        Case "nx1|feature|NX-DATA-UNIQUE-COUNT": NxGeneratedBrandImage = "NxBrand_unique_16"
        Case "nx1|feature|NX-DRAW-BUSINESS-TABLE": NxGeneratedBrandImage = "NxBrand_table_16"
        Case "nx1|feature|NX-DRAW-CLEAR-INNER": NxGeneratedBrandImage = "NxBrand_clear_inner_16"
        Case "nx1|feature|NX-DRAW-INSERT-PICTURE": NxGeneratedBrandImage = "NxBrand_image_16"
        Case "nx1|feature|NX-DRAW-ROLE-STYLE": NxGeneratedBrandImage = "NxBrand_style_16"
        Case "nx1|feature|NX-DRAW-TITLE-TABLE": NxGeneratedBrandImage = "NxBrand_title_table_16"
        Case "nx1|feature|NX-FILE-BATCH-RENAME": NxGeneratedBrandImage = "NxBrand_file_rename_16"
        Case "nx1|feature|NX-FILE-CHART-PNG": NxGeneratedBrandImage = "NxBrand_save_chart_16"
        Case "nx1|feature|NX-FILE-CONSOLIDATE": NxGeneratedBrandImage = "NxBrand_file_merge_16"
        Case "nx1|feature|NX-FILE-FILE-COMPARE": NxGeneratedBrandImage = "NxBrand_compare_file_16"
        Case "nx1|feature|NX-FILE-FOLDER-CREATE": NxGeneratedBrandImage = "NxBrand_folder_add_16"
        Case "nx1|feature|NX-FILE-MANNER-SAVE": NxGeneratedBrandImage = "NxBrand_save_16"
        Case "nx1|feature|NX-FILE-PDF-ALL-COMBINED": NxGeneratedBrandImage = "NxBrand_save_pdf_all_16"
        Case "nx1|feature|NX-FILE-PDF-CURRENT-SHEET": NxGeneratedBrandImage = "NxBrand_save_pdf_16"
        Case "nx1|feature|NX-FILE-RANGE-COPY-SAVE": NxGeneratedBrandImage = "NxBrand_save_range_16"
        Case "nx1|feature|NX-FILE-RANGE-PNG": NxGeneratedBrandImage = "NxBrand_save_image_16"
        Case "nx1|feature|NX-FILE-SHEET-BATCH-RENAME": NxGeneratedBrandImage = "NxBrand_sheet_rename_16"
        Case "nx1|feature|NX-FILE-SHEET-COMPARE": NxGeneratedBrandImage = "NxBrand_compare_sheet_16"
        Case "nx1|feature|NX-FILE-SHEET-COPY-SAVE": NxGeneratedBrandImage = "NxBrand_save_sheet_16"
        Case "nx1|feature|NX-FILE-WORKBOOK-COMPARE": NxGeneratedBrandImage = "NxBrand_compare_range_16"
        Case "nx1|feature|NX-HANGUL-PICTURE-SEND": NxGeneratedBrandImage = "NxBrand_hangul_image_16"
        Case "nx1|feature|NX-HANGUL-TABLE-SEND": NxGeneratedBrandImage = "NxBrand_hangul_table_16"
        Case "nx1|feature|NX-TPL-LIST": NxGeneratedBrandImage = "NxBrand_template_16"
        Case "nx1|feature|NX-UTIL-CALCULATOR": NxGeneratedBrandImage = "NxBrand_calculator_16"
        Case "nx1|feature|NX-UTIL-DOCUMENT-NAVIGATOR": NxGeneratedBrandImage = "NxBrand_navigate_16"
        Case "nx1|feature|NX-UTIL-NAVIGATOR": NxGeneratedBrandImage = "NxBrand_navigate_16"
        Case "nx1|feature|NX-UTIL-SYMBOLS": NxGeneratedBrandImage = "NxBrand_symbols_16"
    End Select
End Function
