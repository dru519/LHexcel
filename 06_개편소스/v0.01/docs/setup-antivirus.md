# 설치 EXE의 로컬 보안 검사

`build/Build-NxSetup.ps1`은 Windows PowerShell 5.1에서 실행한다. Microsoft 서명이 유효한 로컬 컴파일러로 빌드한 뒤, EXE 실행 전과 패키지 확인 후에 Defender 검사를 수행한다. 검사와 해시 확인을 통과한 바이트만 요청한 출력 경로로 복사한다. 기존 파일은 덮어쓰지 않는다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build\Build-NxSetup.ps1 -PackageZip "C:\release\package.zip" -OutputPath "C:\release\review\Setup.exe" -Version v0.01_r95
```

검사 조건은 Defender·실시간 보호 활성화, 48시간 이내 보안 정의, 검사기 Microsoft 서명, 명시적인 무탐지 결과, 검사 전후 동일한 SHA-256, 검사 중 해당 경로의 추가 탐지 없음이다. 정의 갱신은 조직에서 승인한 기존 경로를 사용한다. 정의가 오래됐거나 검사 결과를 해석할 수 없으면 빌드는 중단된다.

검사의 `-DisableRemediation` 옵션은 해당 사용자 지정 검사에서 제외 경로와 압축파일도 검사하고 탐지 내용을 출력한다. 백신 설정이나 실시간 보호를 끄는 옵션이 아니다. 이 진단 검사 자체는 치료하지 않으므로 탐지된 후보는 실행하거나 배포하지 않고 보안 담당자의 처리 대상으로 보존한다. 현재 구현은 영문 `found no threats.` 완료 문구를 인식한다. 다른 언어·새 출력 형식은 정상으로 추정하지 않고 중단한다.

빌드 결과의 `evidence_directory`에 두 검사 JSON과 `build-receipt.json`이 남는다. 영수증에는 EXE·패키지·소스·컴파일러 해시와 검사 정의를 기록한다. 코드 서명 뒤 바이트가 변경되거나 다른 파일을 재패키징한 경우에는 최종 파일을 다시 검사해야 한다. 기존 영수증을 새 파일에 재사용하지 않는다.

`LOCAL_DEFENDER_PASS`는 그 PC와 그 시점 정의의 검사 결과다. 공인 게시자 신뢰, 다른 백신의 판정, 설치·제거의 실제 동작을 보증하지 않는다. 설치 수용 검사는 별도로 진행한다.

## r95 EXE 보류

2026-09-23에 공개한 r95 EXE의 SHA-256은 `1e985598accc2ce650e8b7af79157a708ec120765c726be5d75cfc9a73fb155b`이며 Defender가 `Trojan:Win32/Bearfoos.B!ml`로 격리했다. 같은 소스·같은 ZIP으로 다시 컴파일한 후보의 로컬 검사가 통과했어도 원본 판정의 원인이 해결됐다는 뜻은 아니다. 원본 EXE의 재배포는 보류한다. 탐지를 피하려는 반복 재컴파일, 파일명 변경, 백신 예외 등록은 해결 절차로 사용하지 않는다.

공식 명령 문서: [Microsoft Defender MpCmdRun](https://learn.microsoft.com/en-us/defender-endpoint/command-line-arguments-microsoft-defender-antivirus). 세부 진단 및 증거는 프로젝트 작업 기록의 r95 설치기 Defender 항목을 참조한다.

## r95 재출시 결과

2026-09-23 사용자 승인으로 후보 `dcde4a916c0d7bf77895dd3c7d69f3b2edb56b1227fd7f43dce820ae6ed489f2`를 별도 재출시 폴더에 게시했다. 같은 Windows PC에서 설치·복구·제거, Excel 자동실행·DLL 연결, 사용자 설정 보존을 확인했고 최종 공유 경로의 Defender 검사도 통과했다. 기존 격리본과 다른 해시의 후보에 대한 결과이며, 원본 탐지 해제나 다른 PC의 무탐지를 보증하지 않는다. 상세 결과는 프로젝트 작업 기록 `260923 22;30 r95 재출시와 실제 설치 검증.md`를 따른다.
