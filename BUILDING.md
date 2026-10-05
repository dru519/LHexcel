# 내엑셀 공개 소스 빌드

기능 기준은 v0.01_r105입니다. 빌드는 Windows에서 수행하며, Mac의 정적 검사로 Excel 빌드나 실행 검증을 대신하지 않습니다.

## 준비

Windows PowerShell 5.1, Windows용 Excel, .NET Framework의 C# 컴파일러와 Office COM 상호운용 환경, Python 3 환경이 필요합니다. Python 도구는 pyopenvba·olefile·oletools·Pillow를 사용합니다. 이 저장소는 해당 런타임이나 외부 라이브러리의 구현을 동봉하지 않습니다.

중요한 원본을 보존하고 Excel을 종료한 전용 빌드 환경에서 작업하세요. VBA 프로젝트 접근이 조직 정책에서 허용된 환경이어야 합니다. 조직의 정책을 우회하여 설정하지 마세요.

## BAT 준비 및 실행

저장소 루트에서 실행합니다.

```powershell
python "development/scripts/lhexcel_windows_ascii_build.py" prepare --version v0.01_r105 --attempt local01
```

출력된 `windows_launcher_bat`를 **Windows 안에서** 실행합니다. 출력과 검증 기록은 저장소의 `workspace/tmp/lhexcel_v001r105_direct_run_ascii_local01` 아래에 생성됩니다. 기존 출력이 있으면 새 attempt 이름을 사용하세요.

빌드는 `Build-Xlam.ps1`로 XLAM을 구성하고 `Build-NxHost.ps1 -ProtectedLoader`로 x86/x64 DLL을 만듭니다. 제품 manifest와 생성된 심벌 목록의 일치 여부를 확인합니다. Excel과 COM 검증의 성공 여부는 생성된 증거 파일을 확인해야 합니다.

## 배포본 구성

`tools/build_r47_distribution.py`는 완성 XLAM에서 내부망판 또는 DLL판의 보호 배포본을 만듭니다. `--source-xlam`, `--destination-xlam`, `--protection-seed`, `--profile`을 명시하세요. 보호 seed는 본인이 권한을 가진 보호 XLAM이어야 합니다. 이 저장소에 암호나 개인 서명키는 포함되어 있지 않습니다.

배포 정책 서명은 `build/Sign-DistributionPolicy.ps1`이 Windows의 로컬 키 컨테이너를 사용합니다. 키 저장소를 공유하거나 복사하지 마세요. 이 정책 서명은 Authenticode 전자서명과 별개입니다. 다른 환경에서 만든 빌드를 공식 배포본으로 표시하지 마세요.

`tools/build_enhanced_distribution.py --help`에서 DLL ZIP 생성 옵션을 확인할 수 있습니다. ZIP은 XLAM, NxHost/NxCore 32·64비트 DLL, 해시 목록, 설치·제거 스크립트 및 README를 포함합니다. 설치 EXE는 `build/Build-NxSetup.ps1 -PackageZip <ZIP> -OutputPath <새 EXE> -Version v0.01_r105`로 만듭니다. 컴파일·내장 패키지 검증과 로컬 Defender 검사를 통과해야 게시합니다. 설치·제거 전 과정은 별도의 검증 범위입니다.

## 공개 사본 검사

```powershell
python "source/v0.01/tools/build_product_manifest.py" --check
python "source/v0.01/tools/generate_symbol_catalog.py" --check
python -m unittest discover -s "source/v0.01/tests/python" -p "test_*.py"
```

공개 기본값은 개인 생일 알림을 비활성화합니다. 사용기한과 보호장치를 유지해야 합니다. 소속 조직의 자체 업무를 위한 빌드가 허용되는 범위와 수정본의 외부 배포 제한은 LICENSE를 따릅니다. 컴파일러·Office·로컬 정책 서명과 패키지 메타데이터가 다를 수 있으므로 빌드가 공식 파일과 같은 해시가 된다고 보장하지 않습니다.
