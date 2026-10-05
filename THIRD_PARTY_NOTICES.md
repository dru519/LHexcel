# 외부 구성요소 안내

제품은 Windows와 Microsoft Excel/Office의 COM·.NET 환경을 사용합니다. Windows, Office, .NET Framework의 실행 파일이나 Microsoft 설치 패키지는 이 소스 저장소에 포함하지 않습니다.

Python 빌드·검사 도구는 pyopenvba, olefile, oletools, Pillow를 사용합니다. 해당 구현과 Python 런타임은 별도 설치 환경에서 제공되며, 각 구성요소의 이용 조건을 따릅니다.

소스의 출처 구분은 `06_개편소스/v0.01/provenance`에 보존합니다. `user-approved-v3x-lineage` 표시는 기존 내엑셀 계열의 사용자가 승인한 이식 이력을 뜻합니다. 자체 DLL·설치기·리본·폼의 공개는 내엑셀 LICENSE를 따르며, 외부 제품의 이름이나 기능 참조가 해당 제품의 코드를 포함하거나 이용 권한을 부여한다는 뜻은 아닙니다.
