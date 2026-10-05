# 내엑셀 v0.01_r63 검증 후보

r62 출시 동결 소스 기반 안정화 후보입니다. 출시·설치 성공 기록이 아닙니다.

- Excel 프로세스 비트수 판정, 명령 상태 복원 오류 전달, HWPX 파일 해시의 Windows CNG 재사용.
- 미사용 Office Store Web Extension 4개와 관련 OOXML 참조 정리.
- 보호 배포본 이용 고지 및 설치 안내 자동 포함.
- 고도화 DLL ZIP과 동일 payload를 사용하는 현재 사용자 설치 EXE 병행. 설치 전 확인, HKCU/XLSTART 등록, 실행 정책 자동 변경 없음.

내부망은 DLL 없이 기존 XLAM 공통 기능을 유지합니다. 고도화는 NxHost32.dll/NxHost64.dll을 사용하며 설치 이후 Excel 시작 시 자동 로드합니다.

사용기한은 2027-06-29까지, 2027-06-30부터 차단합니다.

서명은 이번 범위에서 제외합니다. Authenticode/VBA signature는 NOT_RUN, 조직 신뢰 승인은 HOLD입니다.
실제 Excel 자동 로드·DLL 연결, 설치/제거 복원, x86/x64 수용은 개별 네이티브 기록을 확인해야 합니다.
첫 휠 입력 누락과 조건부서식 포커스셀의 복사/Undo 중 갱신 유예는 기존 수용 제한입니다.
