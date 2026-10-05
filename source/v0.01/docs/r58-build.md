# 내엑셀 v0.01_r58

- profile: enhanced-dll
- product: Product.xlam
- native sidecars: NxHost32.dll, NxHost64.dll
- internal-network fallback: Product.xlam 단독 실행 및 조건부서식 기반 포커스셀
- candidate status: PASS_TARGETED_NATIVE
- native Excel: x64, Windows 실사용 200%
- x86 DLL: 빌드 및 PE 구조 확인; x86 Excel 실행은 NOT_RUN
- Authenticode: NOT_RUN
- VBA signature: NOT_RUN
- formal signed release: HOLD

r58은 홈/내엑셀 탭 아이콘·표기 일치, 안전한 파일 열기와 보고서, 이름 변경/PDF 복구, 데이터 전후 미리보기와 탐색 상태 유지를 보완한다.
닫힌 파일의 안전 열기는 검증 가능한 .xlsx/.xlsm 사본을 사용한다. 구형/암호화/XLM 입력은 자동 변환하지 않는다.
문서 탐색 CTP 후보는 실제 Excel COM 연결 검증 미통과로 포함하지 않는다.

2026-09-06 변경 영역 검증: ProductRibbon/ProductUi PASS, ShortcutKeyboard 13/13,
ShortcutManager 20/20, SafeOpen 13/13, Convergence 10/10. 정적 시험 1,082개 PASS.
단축키는 실제 조합키 입력·등록·호출과 Excel 재실행을 검증했다.
비교는 생성 함수 반환 뒤 같은 Excel에서 보고서를 읽기 전용으로 열고 원래 값·수식·병합·서식 검사를 유지했다.
모든 기능의 네이티브 전수 검사를 반복한 결과가 아니다. 배포 보호본의 별도 실행 결과는 출시 기록을 따른다.
배포본은 2027-06-29까지 사용 가능하며 2027-06-30부터 차단한다. 생일 안내 및 기존 VBA 보호 규칙을 유지한다.
