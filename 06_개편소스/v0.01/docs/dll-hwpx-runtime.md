# DLL HWPX 표 생성 경로

r63 기반 DLL 1차 고도화의 개발 계약입니다. 정식 버전 승격이나 배포 승인을 뜻하지 않습니다.

## 구성

- Excel UI 스레드: 선택 범위·표시값·병합·숨김 행·열 및 설정을 기존 직렬화기로 JSON에 고정합니다.
- `LH.NxHost.HwpxExportService`: 제품명, nx-hwpx-v1, bridge/feature/command 해시와 Excel 비트수를 확인한 뒤 단일 비동기 작업을 수락합니다.
- DLL 작업 스레드: Excel COM 객체 없이 고정된 JSON과 해시 검증된 자체 HWPX 템플릿으로 표를 생성합니다.
- 임시 .part 파일을 완성한 뒤 같은 폴더의 GUID.hwpx로 게시합니다. 기존 파일은 덮어쓰지 않습니다.
- DLL 연결 불가·호환 불일치일 때만 기존 내장 PowerShell 경로를 사용합니다. DLL Start 이후 오류·취소에는 중복 실행하지 않습니다.

## 인터페이스 및 제한

ProgID: LH.NxHost.HwpxExportService  
Class GUID: 1434C649-18AA-4435-B0D2-2BD81B548A01  
Interface GUID: 2C1FC9B6-2471-4A82-9080-10331E120801

메서드: Connect → Start → GetStatus → GetResult/GetError, 취소는 Cancel. 설정 오류의 상세 내부 예외는 외부로 노출하지 않습니다.

JSON 최대 16,777,216자, 16,384셀, 셀 텍스트 32,767자, 동시 작업 1개, 제한시간 60초입니다. 경로는 로컬 드라이브만 허용하고 UNC, ADS, 확인된 reparse 경로를 거부합니다. 임의 출력 경로나 Shell 명령을 받지 않습니다.

기존 생성기의 글꼴 4종, 9~24pt, 공문형·간결형, 제목행, 병합, 폭·높이 설정을 유지합니다. 기본 글꼴은 모든 스크립트 슬롯에서 중고딕/HFT입니다. Excel 상태표시줄을 변경하지 않으며 EnableCancelKey는 복원합니다.

## 재검증

- `python -m unittest tests.python.test_hwpx_native_writer -v`: 환경변수 NX_HWPX_TEST_DLL로 검사할 x64 DLL을 지정할 수 있습니다. 컴파일된 HwpxProbe.exe가 필요합니다.
- `tests/windows/NxHwpxServiceProbe.cs`: x86/x64 별도 빌드하여 연결 불일치·반복 생성·취소·작업 소유권·잔류 파일을 검사합니다.
- 프로젝트의 `lhexcel_windows_ascii_build.py prepare --version v0.01_r63 --attempt <새 이름> --suite dll-hwpx --no-clean`으로 만든 BAT를 Windows에서 실행합니다. 검증 하네스는 열려 있는 Excel이 있으면 거부합니다.
- `verify --version v0.01_r63 --attempt <기존 이름> --suite dll-hwpx --tag <새 이름>`은 기존 빌드 바이트를 유지한 재검증입니다.

BAT 시험은 새 HWPX COM 클래스만 임시 등록하고 해제합니다. 기존 설치된 Bridge/AddIn/Navigator 등록을 변경하지 않습니다. 취소 검사 중에는 저장하지 않는 시험용 XLAM 사본에 오류 18을 주입하므로 물리 키보드 Esc 검증과 구별해야 합니다.

## 남은 경계

2026-09-08 후속 검증에서 한컴 실제 열기와 한컴 PDF 출력(1쪽, 한글·특수문자·병합·테두리)을 확인했습니다. 동일 제품 바이트로 만든 개발 검증 EXE의 실제 설치 → Excel 정상 시작 시 XLAM 자동 로드 → DLL 호환 연결 → HWPX COM 등록 → 정상 종료 → 제거도 통과했습니다. 설치 안내와 개발 검증본 경고를 포함한 ZIP/EXE이며 정식 배포 보호와 서명은 적용하지 않았습니다.

2026-09-08 탐색창 후속 수정은 Windows 200% / x64 Excel에서 검증했습니다. 이전 본문 300×460px를 600×920px로 바로잡고, 모니터 가장자리가 아니라 Excel 창과 작업 영역의 교집합 안에 배치합니다. UI Automation에서 실제 창/본문 경계를 읽어 크기 및 창 축소 후 위치를 확인했습니다. 검색 목록에는 관측한 목록 HWND로 Home 키 메시지를 전달하고 UIA 선택/텍스트로 결과를 검증했습니다. 포커스셀 OFF 상태에서 보호 시트 실행 불가 → 보호 해제 실행 가능 갱신, 숨김·재표시 후 HWND와 선택 유지가 PASS입니다. 이 검사는 물리 키보드 또는 마우스 조작 전수 검사가 아닙니다.

초기 CTP 진단은 factory_available → create_ctp → ctp_failed(HRESULT 80004005) → fallback_shown이었습니다. 후속 x86/x64 명시 프로세스의 QueryInterface에서 OLE 인터페이스는 성공했지만 IDispatch만 E_NOINTERFACE였습니다. NavigatorPane에 빈 명시적 COM 기본 IDispatch 인터페이스를 추가한 뒤 실제 Excel CreateCTP가 성공했습니다. ClassInterface(None)을 유지하며 전체 Control 멤버를 자동 노출하지 않습니다. 보안 정책은 변경하지 않았습니다. Windows ARM64 PowerShell의 BadImageFormat과 x64 Excel 호스팅 오류는 서로 다른 검증 단계입니다.

도킹창은 Excel SDI HWND별 NavigatorSession으로 소유합니다. 숨김/재표시 시 동일 내용 컨트롤과 선택을 유지하며, 다른 통합문서 창에는 별도 탐색 상태를 만듭니다. 닫힌 창의 CTP를 삭제하고 잔류 HWND를 검사합니다. Office.Core의 원시 CTP 폭은 네이티브 측정에 맞춰 물리 픽셀로 지정하고, CTP와 Form 대체 경로의 DPI 소유자를 분리합니다. 작은 Office 클라이언트 영역에서도 하단 컨트롤이 보이도록 CTP의 최소 크기를 해제합니다.

Windows 200% / x64 Excel에서 실제 CTP 자식 HWND, 600px 도킹 폭, 창 축소 후 모든 직접 자식의 경계, 검색/선택, Focus-OFF 보호 상태 갱신, 복수 통합문서 창의 독립성, 다른 창 종료 후 원래 창 유지 및 Excel 자연 종료를 확인했습니다. 진단은 LHEXCEL_NXHOST_DIAGNOSTICS=1 및 기존 절대경로 LHEXCEL_PROFILE_ROOT가 있을 때만 기록하며 문서 내용·예외 메시지는 남기지 않습니다.

남은 수용 검증: 다른 DPI 및 혼합 DPI 모니터 이동, x86 Excel 실사용, 같은 통합문서의 새 창 및 별도 Excel 프로세스 간 이동, 실물 인쇄, 활성 기존 설치에서의 버전 갱신. 기존 완성본·배포본은 변경하지 않았습니다. 최신 DLL의 설치 패키지 내용/해시 검증과 이전 DLL의 실제 설치 시험은 서로 구분하며, 새 설치 검증은 해당 패키지의 영수증을 기준으로 합니다.

재실행: `verify --version v0.01_r63 --attempt <기존 이름> --suite dll-visual --tag <새 이름>`은 15분 제한의 시험 통합문서 세션을 만듭니다. 01-protect부터 05-stop까지 순서가 고정된 .request 파일로 시험 상태만 바꿉니다. UI 관찰은 별도로 수행하며 세션 PASS를 UI PASS로 해석하지 않습니다. `setup-acceptance --setup-exe <시험 EXE> --tag <새 이름>`은 실제 설치·자동 로드·연결·제거 BAT를 만듭니다.

`verify --version v0.01_r63 --attempt <기존 이름> --suite dll-navigator --tag <새 이름>`은 기존 XLAM을 유지하고 현재 NxHost 소스를 별도 snapshot에 복사하여 x86/x64 DLL을 새로 빌드한 뒤 위 UIA 회귀 검사를 수행하는 BAT를 생성합니다. 증거는 긴 경로 제한을 피하기 위해 `%LOCALAPPDATA%/LHExcel/verification/<새 이름>`에 저장합니다. 정상 종료, 실행 중 Excel 0, 레지스트리 원상복원 및 XLAM 해시 보존을 요구합니다. 이미 종료됐지만 핸들 때문에 잠시 열거되는 프로세스는 exited_process_entries에 별도 기록하며 실행 중으로 숨기지 않습니다.
