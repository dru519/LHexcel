#!/usr/bin/env python3
"""Stage and promote Windows Excel COM builds through an ASCII-only runner.

This script captures the successful Parallels workflow used for v2.9:
1. Copy the source XLAM, exported VBA source, and lhexcel_dev.ps1 into an
   ASCII-only temp directory.
2. Generate a CRLF run.bat and ASCII build.ps1 that Windows CMD/PowerShell 5.1
   can execute without Korean-path encoding issues.
3. After the user runs run.bat in Windows, promote the timestamped output back
   to the project candidate/final locations and archive the logs.

Execution contract: use the generated BAT files inside Windows. Do not use
prlctl exec or a Mac .app launcher for this project; the installed Parallels
edition blocks CLI guest execution and the WinApp launcher path is unreliable.
"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import re
import shutil
import subprocess
import sys
import unicodedata
import zipfile
from dataclasses import dataclass, replace
from pathlib import Path

def find_project_root(start: Path) -> Path:
    for path in (start, *start.parents):
        if (path / "source/v0.01/contracts").is_dir() and (path / "development/scripts").is_dir():
            return path
        if (path / "01_완성본").exists() and (path / "source").exists():
            return path
    raise RuntimeError(f"project root not found from {start}")


ROOT = find_project_root(Path(__file__).resolve())
SCRIPT_DIR = ROOT / "development" / "scripts"
WORKSPACE = ROOT / "workspace"
TMP_ROOT = WORKSPACE / "tmp"


@dataclass(frozen=True)
class VersionPaths:
    version: str
    token: str
    version_dir: Path
    source_xlam: Path
    source_dir: Path
    candidate_xlam: Path
    final_xlam: Path
    reports_dir: Path
    tmp_dir: Path
    builder: Path
    launcher_bat: Path


@dataclass(frozen=True)
class V001ReleasePaths:
    version: str
    token: str
    source_root: Path
    tmp_dir: Path
    launcher_bat: Path


def normalize_version(value: str) -> str:
    raw = value.strip()
    if not raw:
        raise SystemExit("version is required")
    if not raw.lower().startswith("v"):
        raw = "v" + raw
    return raw


def version_token(version: str) -> str:
    return "v" + re.sub(r"[^0-9A-Za-z]", "", version[1:])


def uses_strict_lh_case(version: str) -> bool:
    match = re.match(r"^v(\d+)(?:\.(\d+))?", version, re.I)
    if not match:
        return True
    major = int(match.group(1))
    minor = int(match.group(2) or "0")
    return (major, minor) >= (3, 1)


def is_v001_release(version: str) -> bool:
    return re.fullmatch(r"v0\.01_r[0-9]+", version, re.I) is not None


def resolve_v001_release_paths(version_arg: str) -> V001ReleasePaths:
    version = normalize_version(version_arg)
    if not is_v001_release(version):
        raise SystemExit(f"unsupported v0.01 release version: {version}")
    source_root = ROOT / "source" / "v0.01"
    if not source_root.is_dir():
        raise SystemExit(f"v0.01 source root not found: {source_root}")
    token = version_token(version)
    return V001ReleasePaths(
        version=version,
        token=token,
        source_root=source_root,
        tmp_dir=TMP_ROOT / f"lhexcel_{token}_direct_run_ascii",
        launcher_bat=TMP_ROOT / f"run_{token}_build.bat",
    )


def resolve_paths(version_arg: str) -> VersionPaths:
    version = normalize_version(version_arg)
    token = version_token(version)
    version_dir = ROOT / "source" / version
    if not version_dir.exists():
        raise SystemExit(f"version directory not found: {version_dir}")

    source_candidates = [
        version_dir / "base" / f"LHExcel_{token}_excel_open_source.xlam",
        version_dir / "candidate" / f"lhexcel_{token}_static_candidate.xlam",
        version_dir / "base" / f"LHExcel_{token}_package_source.xlam",
        version_dir / "base" / f"내엑셀 {version}_source.xlam",
    ]
    source_xlam = next((path for path in source_candidates if path.exists()), None)
    if source_xlam is None:
        matches = sorted((version_dir / "base").glob("*.xlam"))
        if not matches:
            raise SystemExit(f"source XLAM not found under: {version_dir / 'base'}")
        source_xlam = matches[0]

    source_dir = version_dir / "candidate" / "src"
    if not source_dir.is_dir():
        raise SystemExit(f"candidate source dir not found: {source_dir}")

    candidate_xlam = version_dir / "candidate" / f"lhexcel_{token}_candidate.xlam"
    final_xlam = ROOT / "01_완성본" / f"내엑셀 {version}.xlam"
    reports_dir = version_dir / "reports"
    tmp_dir = TMP_ROOT / f"lhexcel_{token}_direct_run_ascii"
    builder = SCRIPT_DIR / "lhexcel_dev.ps1"
    if not builder.exists():
        raise SystemExit(f"builder script not found: {builder}")
    launcher_bat = TMP_ROOT / f"run_{token}_build.bat"

    return VersionPaths(
        version=version,
        token=token,
        version_dir=version_dir,
        source_xlam=source_xlam,
        source_dir=source_dir,
        candidate_xlam=candidate_xlam,
        final_xlam=final_xlam,
        reports_dir=reports_dir,
        tmp_dir=tmp_dir,
        builder=builder,
        launcher_bat=launcher_bat,
    )


def to_windows_path(path: Path) -> str:
    resolved = path.resolve()
    if sys.platform == "win32":
        return str(resolved).replace("/", "\\")
    try:
        rel = resolved.relative_to(Path.home())
    except ValueError:
        return str(resolved).replace("/", "\\")
    return "C:\\Mac\\Home\\" + str(rel).replace("/", "\\")


def sanitize_tag(tag: str) -> str:
    clean = re.sub(r"[^0-9A-Za-z가-힣_.-]+", "-", tag.strip())
    return clean.strip("-") or "direct-run"


def write_ascii(path: Path, text: str, *, crlf: bool = False) -> None:
    if crlf:
        text = re.sub(r"\r?\n", "\r\n", text)
    path.write_text(text, encoding="ascii")


def read_vba_source_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp949"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"unsupported VBA source encoding: {path}")


def ensure_exported_class_header(text: str) -> str:
    if re.match(r"^\s*VERSION\s+1\.0\s+CLASS\b", text, re.I):
        return text
    return "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\n" + text


def write_windows_vba_source(source: Path, destination: Path) -> None:
    text = unicodedata.normalize("NFC", read_vba_source_text(source))
    if source.suffix.lower() == ".cls":
        text = ensure_exported_class_header(text)
    text = re.sub(r"\r?\n", "\r\n", text)
    destination.write_bytes(text.encode("cp949", errors="strict"))


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sync_generated_sources(paths: VersionPaths, *, check: bool) -> None:
    sync_script = SCRIPT_DIR / f"lhexcel_{paths.token}_sync_embedded_support.py"
    if not sync_script.exists():
        return
    mode = "--check" if check else "--write"
    run_checked(["python3", str(sync_script), mode], ROOT)


def stage_source_tree(source_dir: Path, destination_dir: Path) -> None:
    if destination_dir.exists():
        shutil.rmtree(destination_dir)
    destination_dir.mkdir(parents=True)
    for source in sorted(source_dir.iterdir()):
        destination = destination_dir / source.name
        if source.is_dir():
            shutil.copytree(source, destination)
        elif source.suffix.lower() in {".bas", ".cls", ".frm"}:
            write_windows_vba_source(source, destination)
        else:
            shutil.copy2(source, destination)


def v001_release_build_ps1_text(version: str) -> str:
    if not is_v001_release(version):
        raise ValueError(f"unsupported v0.01 release version: {version}")
    script = f'''$ErrorActionPreference = "Stop"
Import-Module Microsoft.PowerShell.Utility -Global -ErrorAction Stop
[void](Get-Command Get-FileHash -ErrorAction Stop)
Set-StrictMode -Version Latest

$root = [IO.Path]::GetFullPath((Split-Path -Parent $MyInvocation.MyCommand.Path))
$sourceRoot = Join-Path $root "source"
$outRoot = Join-Path $root "out"
$completeRoot = Join-Path $outRoot "complete"
$evidenceRoot = Join-Path $root "evidence"
$logRoot = Join-Path $root "logs"
$runLog = Join-Path $root "windows-build-direct-run.log"
$latest = Join-Path $root "latest-output.txt"
$python = (Get-Command python -ErrorAction Stop).Source
$runId = [Guid]::NewGuid().ToString().ToLowerInvariant()

if (Test-Path -LiteralPath $outRoot) {{ throw "Existing v0.01 build output is rejected; run prepare again" }}
if (Test-Path -LiteralPath $evidenceRoot) {{ throw "Existing v0.01 evidence is rejected; run prepare again" }}
[void](New-Item -ItemType Directory -Path $completeRoot -Force)
[void](New-Item -ItemType Directory -Path $evidenceRoot -Force)
[void](New-Item -ItemType Directory -Path $logRoot -Force)
"$(Get-Date -Format o) START {version} BAT-only build" | Set-Content -LiteralPath $runLog -Encoding UTF8
"run_id=$runId" | Add-Content -LiteralPath $runLog -Encoding UTF8

function Invoke-CheckedPowerShell([string]$Name, [string]$ScriptPath, [string[]]$Arguments, [int[]]$AcceptedCodes = @(0)) {{
    $stdout = Join-Path $logRoot ($Name + ".stdout.log")
    $stderr = Join-Path $logRoot ($Name + ".stderr.log")
    $priorErrorAction = $ErrorActionPreference
    try {{
        $ErrorActionPreference = "Continue"
        & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $ScriptPath @Arguments 1> $stdout 2> $stderr
    }} finally {{ $ErrorActionPreference = $priorErrorAction }}
    $code = [int]$LASTEXITCODE
    "$Name=$code" | Add-Content -LiteralPath $runLog -Encoding UTF8
    if ($AcceptedCodes -notcontains $code) {{
        $detail = if (Test-Path -LiteralPath $stderr) {{ [IO.File]::ReadAllText($stderr) }} else {{ "" }}
        throw ($Name + " failed with exit " + $code + ": " + $detail)
    }}
}}

$product = Join-Path $completeRoot "Product.xlam"
& (Join-Path $sourceRoot "build/Build-Xlam.ps1") -DataRoot $sourceRoot -OutputPath $product `
    -ManifestPath (Join-Path $sourceRoot "build/manifests/Product.json") `
    -RibbonPath (Join-Path $sourceRoot "src/ribbon/customUI14.xml") *>> $runLog
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $product -PathType Leaf)) {{ throw "Build-Xlam failed" }}

$hostRoot = Join-Path $outRoot "NxHost"
& (Join-Path $sourceRoot "build/Build-NxHost.ps1") -OutputRoot $hostRoot -Configuration Release -ProtectedLoader *>> $runLog
if ($LASTEXITCODE -ne 0) {{ throw "Build-NxHost failed" }}
Copy-Item -LiteralPath (Join-Path $hostRoot "x86/NxHost32.dll") -Destination (Join-Path $completeRoot "NxHost32.dll")
Copy-Item -LiteralPath (Join-Path $hostRoot "x64/NxHost64.dll") -Destination (Join-Path $completeRoot "NxHost64.dll")
Copy-Item -LiteralPath (Join-Path $hostRoot "x86/NxCore32.dll") -Destination (Join-Path $completeRoot "NxCore32.dll")
Copy-Item -LiteralPath (Join-Path $hostRoot "x64/NxCore64.dll") -Destination (Join-Path $completeRoot "NxCore64.dll")

$docs = Join-Path $completeRoot "docs"
[void](New-Item -ItemType Directory -Path $docs)
Copy-Item -LiteralPath (Join-Path $sourceRoot "docs/r57-build.md") -Destination (Join-Path $docs "r57-build.md")
$manifestByRelative = @{{}}
$manifestNames = @(
    Get-ChildItem -LiteralPath $completeRoot -Recurse -File |
        Where-Object {{ $_.Name -ne "SHA256SUMS" }} |
        ForEach-Object {{
            $relative = $_.FullName.Substring($completeRoot.Length + 1).Replace([char]92, [char]47)
            $manifestByRelative[$relative] = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
            $relative
        }}
)
[Array]::Sort($manifestNames, [StringComparer]::Ordinal)
$manifestRows = @($manifestNames | ForEach-Object {{ $manifestByRelative[$_] + "  " + $_ }})
[IO.File]::WriteAllText((Join-Path $completeRoot "SHA256SUMS"), (($manifestRows -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))
& $python -X utf8 (Join-Path $sourceRoot "tools/verify_build_profile.py") --profile enhanced-dll --payload-root $completeRoot *>> $runLog
if ($LASTEXITCODE -ne 0) {{ throw "enhanced-dll payload verification failed" }}
$sourceDigest = (Get-FileHash -LiteralPath (Join-Path $sourceRoot "build/manifests/Product.json") -Algorithm SHA256).Hash.ToLowerInvariant()
$snapshotDigest = (Get-FileHash -LiteralPath $product -Algorithm SHA256).Hash.ToLowerInvariant()

Invoke-CheckedPowerShell "product-ui" (Join-Path $sourceRoot "tests/windows/Run-ProductUiSuite.ps1") @(
    "-DataRoot", $sourceRoot, "-EvidenceRoot", (Join-Path $evidenceRoot "product-ui"),
    "-ArtifactPath", $product, "-RunId", $runId
)
Invoke-CheckedPowerShell "focus-conditional" (Join-Path $sourceRoot "tests/windows/Run-FocusConditionalProductSuite.ps1") @(
    "-DataRoot", $sourceRoot, "-EvidenceRoot", (Join-Path $evidenceRoot "focus-conditional"), "-RunId", $runId
)
Invoke-CheckedPowerShell "focus-viewport" (Join-Path $sourceRoot "tests/windows/Run-FocusViewportSyncSuite.ps1") @(
    "-EvidenceRoot", (Join-Path $evidenceRoot "focus-viewport"), "-ArtifactPath", $product, "-NxHostRoot", $completeRoot
)
Invoke-CheckedPowerShell "shortcuts" (Join-Path $sourceRoot "tests/windows/Run-ShortcutManagerSuite.ps1") @(
    "-ProductXlam", $product, "-EvidenceRoot", (Join-Path $evidenceRoot "shortcuts"), "-RunId", $runId
)
Invoke-CheckedPowerShell "r57-usability" (Join-Path $sourceRoot "tests/windows/Run-R57UsabilitySuite.ps1") @(
    "-ProductXlam", $product, "-EvidenceRoot", (Join-Path $evidenceRoot "r57-usability"), "-RunId", $runId
)

$vbaEvidence = Join-Path $evidenceRoot "symbols-vba"
Invoke-CheckedPowerShell "symbols-red-probe" (Join-Path $sourceRoot "tests/windows/Run-VbaSuite.ps1") @(
    "-Suite", "Core", "-Mode", "RedProbe", "-ExpectedCompileFailure", "CNxStateGuard",
    "-DataRoot", $sourceRoot, "-EvidenceRoot", $vbaEvidence, "-RunId", $runId,
    "-SourceDigest", $sourceDigest, "-SnapshotDigest", $snapshotDigest
) @(1)
$redProbe = Get-Content -LiteralPath (Join-Path $vbaEvidence "Core.RedProbe.json") -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$redProbe.classification -cne "EXPECTED_RED") {{ throw "Symbols prerequisite red probe was not expected" }}
Invoke-CheckedPowerShell "symbols" (Join-Path $sourceRoot "tests/windows/Run-VbaSuite.ps1") @(
    "-Suite", "Symbols", "-Mode", "Green", "-DataRoot", $sourceRoot,
    "-EvidenceRoot", $vbaEvidence, "-ExtensionManifest", (Join-Path $sourceRoot "tests/manifests/Symbols.json"), "-RunId", $runId,
    "-SourceDigest", $sourceDigest, "-SnapshotDigest", $snapshotDigest
)

$product | Set-Content -LiteralPath $latest -Encoding ASCII
"$(Get-Date -Format o) PASS {version}" | Add-Content -LiteralPath $runLog -Encoding UTF8
Write-Output ("PASS|{version}|" + $completeRoot)
exit 0
'''
    revision = version.rsplit("_", 1)[1]
    if int(revision[1:]) >= 58:
        script = script.replace("r57-build.md", f"{revision}-build.md")
        start = script.index('Invoke-CheckedPowerShell "focus-conditional"')
        end = script.index('$product | Set-Content', start)
        script = script[:start] + '''Invoke-CheckedPowerShell "r58-convergence" (Join-Path $sourceRoot "tests/windows/Run-R58ConvergenceSuite.ps1") @(
    "-ProductXlam", $product, "-EvidenceRoot", (Join-Path $evidenceRoot "r58-convergence"), "-RunId", $runId
)
Invoke-CheckedPowerShell "r58-navigator-model" (Join-Path $sourceRoot "tests/windows/Run-R58NavigatorModelSuite.ps1") @(
    "-EvidenceRoot", (Join-Path $evidenceRoot "r58-navigator-model")
)
Invoke-CheckedPowerShell "shortcuts" (Join-Path $sourceRoot "tests/windows/Run-ShortcutManagerSuite.ps1") @(
    "-ProductXlam", $product, "-EvidenceRoot", (Join-Path $evidenceRoot "shortcuts"), "-RunId", $runId
)

''' + script[end:]
    if int(revision[1:]) >= 59:
        start = script.index('Invoke-CheckedPowerShell "r58-convergence"')
        end = script.index('$product | Set-Content', start)
        script = script[:start] + '''Invoke-CheckedPowerShell "r59-manner" (Join-Path $sourceRoot "tests/windows/Run-R59MannerSuite.ps1") @(
    "-ProductXlam", $product, "-EvidenceRoot", (Join-Path $evidenceRoot "r59-manner"), "-RunId", $runId
)

''' + script[end:]
    return script


def v001_release_run_bat_text(paths: V001ReleasePaths) -> str:
    win_dir = "%~dp0"
    return f'''@echo off
setlocal
pushd "{win_dir}" || exit /b 1
echo %DATE% %TIME% > runner-started.txt
echo START {paths.version} BAT-only build and targeted validation
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File build.ps1
set "RC=%ERRORLEVEL%"
echo EXIT CODE=%RC%
echo LOG={win_dir}\windows-build-direct-run.log
if not "%LHEXCEL_NO_PAUSE%"=="1" pause
popd
endlocal & exit /b %RC%
'''


def prepare_v001_release(paths: V001ReleasePaths, *, clean: bool, suite: str = "all") -> None:
    if clean and paths.tmp_dir.exists():
        shutil.rmtree(paths.tmp_dir)
    if paths.tmp_dir.exists() and any(paths.tmp_dir.iterdir()):
        raise SystemExit(f"refusing non-empty v0.01 runner directory: {paths.tmp_dir}")
    paths.tmp_dir.mkdir(parents=True, exist_ok=True)
    staged_source = paths.tmp_dir / "source"
    shutil.copytree(
        paths.source_root,
        staged_source,
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".pytest_cache", "out"),
    )
    script = v001_release_build_ps1_text(paths.version)
    if suite == "manner":
        start = script.index('Invoke-CheckedPowerShell "product-ui"')
        end = script.index('Invoke-CheckedPowerShell "r59-manner"', start)
        script = script[:start] + script[end:]
    elif suite == "r68-product":
        # Targeted development verification only; the default release suites remain unchanged.
        start = script.index('Invoke-CheckedPowerShell "product-ui"')
        end = script.index('$product | Set-Content', start)
        script = script[:start] + '''Invoke-CheckedPowerShell "r68-table" (Join-Path $sourceRoot "tests/windows/Run-R68ProductSuite.ps1") @(
    "-ProductXlam", $product, "-SourceRoot", $sourceRoot,
    "-EvidenceRoot", (Join-Path $evidenceRoot "r68-table"), "-Suites", "table"
)

''' + script[end:]
    elif suite == "dll-hwpx":
        script = "Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop\n" + script
        start = script.index('Invoke-CheckedPowerShell "product-ui"')
        end = script.index('$product | Set-Content', start)
        script = script[:start] + '''Invoke-CheckedPowerShell "dll-hwpx" (Join-Path $sourceRoot "tests/windows/Run-NxHwpxNativeSuite.ps1") @(
    "-ProductXlam", $product, "-EvidenceRoot", (Join-Path $evidenceRoot "dll-hwpx"),
    "-NxHostRoot", $completeRoot, "-RunId", $runId
)

''' + script[end:]
    write_ascii(paths.tmp_dir / "build.ps1", script)
    write_ascii(paths.tmp_dir / "run.bat", v001_release_run_bat_text(paths), crlf=True)
    write_ascii(paths.launcher_bat, launcher_bat_text(paths), crlf=True)
    info = {
        "version": paths.version,
        "token": paths.token,
        "source_root": str(paths.source_root),
        "staged_source_root": str(staged_source),
        "tmp_dir": str(paths.tmp_dir),
        "windows_run_bat": to_windows_path(paths.tmp_dir / "run.bat"),
        "windows_launcher_bat": to_windows_path(paths.launcher_bat),
        "execution_contract": "BAT_ONLY_INSIDE_WINDOWS",
        "targeted_suites": (["R68Product.table"] if suite == "r68-product" else ["NxHwpxNative"] if suite == "dll-hwpx" else ["R59Manner"] if suite == "manner" else ["ProductUi", "R59Manner"] if int(paths.version.rsplit("_r", 1)[1]) >= 59 else ["ProductUi", "R58Convergence", "R58NavigatorModel", "ShortcutManager"] if int(paths.version.rsplit("_r", 1)[1]) >= 58 else [
            "ProductUi",
            "FocusConditionalProduct",
            "FocusViewportSync",
            "ShortcutManager",
            "R57Usability",
            "Symbols",
        ]),
        "complete_root": str(paths.tmp_dir / "out" / "complete"),
    }
    (paths.tmp_dir / "runner-info.json").write_text(
        json.dumps(info, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(info, ensure_ascii=False, indent=2))


def prepare_v001_verification(paths: V001ReleasePaths, tag: str, suite: str = "all", cases: str = "", artifact: Path | None = None, host_root: Path | None = None) -> None:
    """Resume changed-path tests without rebuilding or erasing prior evidence."""
    if not re.fullmatch(r"[a-z0-9-]+", tag):
        raise SystemExit("verification tag must be lowercase ASCII letters, digits or hyphens")
    product = artifact.resolve() if artifact else paths.tmp_dir / "out/complete/Product.xlam"
    if (artifact or host_root) and suite not in ("dll-navigator", "compare-benchmark", "focus-speed"):
        raise SystemExit("artifact/host overrides are scoped to navigator verification")
    if host_root and cases == "fallback":
        raise SystemExit("cannot fault-inject shipping DLLs")
    if not product.is_file():
        raise SystemExit("verification requires an existing built Product.xlam")
    ps_path = paths.tmp_dir / f"verify-{tag}.ps1"
    bat_path = paths.tmp_dir / f"verify-{tag}.bat"
    if ps_path.exists() or bat_path.exists() or (paths.tmp_dir / f"evidence-{tag}").exists():
        raise SystemExit("existing verification runner/evidence is rejected")
    if suite == "compare-reopen":
        fixture = Path(cases).resolve()
        if not fixture.is_file() or fixture.suffix.lower() != ".xlsx":
            raise SystemExit("compare-reopen requires an existing XLSX path in --cases")
        snapshot = paths.tmp_dir / f"source-{tag}"
        for relative in ("tests/windows/Run-R101CompareReopen.ps1", "build/Excel-ProcessLifecycle.ps1"):
            destination = snapshot / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(paths.source_root / relative, destination)
        shutil.copy2(fixture, snapshot / "report.xlsx")
        script = f'''$ErrorActionPreference = "Stop"
$source=Join-Path $PSScriptRoot "source-{tag}"
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source "tests/windows/Run-R101CompareReopen.ps1") -InputFile (Join-Path $source "report.xlsx") -EvidenceRoot (Join-Path $PSScriptRoot "evidence-{tag}")
exit $LASTEXITCODE
'''
        write_ascii(ps_path, script)
        write_ascii(bat_path, v001_release_run_bat_text(paths).replace("-File build.ps1", f"-File verify-{tag}.ps1"), crlf=True)
        print(json.dumps({"verification_bat": str(bat_path), "fixture_sha256": file_sha256(fixture)}))
        return
    if suite == "hangul-transfer":
        if cases not in ("", "picker"):
            raise SystemExit("unknown hangul transfer case")
        picker_flag = " -UseRangePicker" if cases == "picker" else ""
        snapshot = paths.tmp_dir / f"source-{tag}"
        shutil.copytree(paths.source_root, snapshot, ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".pytest_cache", "out", "*.dll", "*.xlam"))
        script = f'''$ErrorActionPreference = "Stop"
$source=Join-Path $PSScriptRoot "source-{tag}"
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source "tests/windows/Run-HangulTableSuite.ps1") -DataRoot $source -ArtifactPath (Join-Path $PSScriptRoot "out/complete/Product.xlam") -EvidenceRoot (Join-Path $PSScriptRoot "evidence-{tag}"){picker_flag}
exit $LASTEXITCODE
'''
        write_ascii(ps_path, script)
        write_ascii(bat_path, v001_release_run_bat_text(paths).replace("-File build.ps1", f"-File verify-{tag}.ps1"), crlf=True)
        print(json.dumps({"verification_bat": str(bat_path), "built_product_sha256": file_sha256(product)}))
        return
    if suite in ("frame-draw", "frame-product", "frame-product-all", "compare-results", "picture-state", "r68-product"):
        snapshot = paths.tmp_dir / f"source-{tag}"
        shutil.copytree(paths.source_root, snapshot, ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".pytest_cache", "out", "tmp", "*.dll", "*.xlam"))
        script = f'''$ErrorActionPreference = "Stop"
$source=Join-Path $PSScriptRoot "source-{tag}"
$evidence=Join-Path $PSScriptRoot "evidence-{tag}"
$run=[Guid]::NewGuid().ToString()
$digest=(Get-FileHash -LiteralPath (Join-Path $source "build/manifests/Product.json")).Hash.ToLowerInvariant()
$product=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot "out/complete/Product.xlam")).Hash.ToLowerInvariant()
foreach($suite in @("Frame","Draw")){{
    & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source "tests/windows/Run-VbaSuite.ps1") -Suite $suite -Mode Green -DataRoot $source -EvidenceRoot $evidence -RunId $run -SourceDigest $digest -SnapshotDigest $product -ExtensionManifest (Join-Path $source ("tests/manifests/"+$suite+".json"))
    if($LASTEXITCODE -ne 0){{throw ($suite+" verification failed")}}
}}
exit 0
'''
        if suite in ("frame-product", "frame-product-all", "compare-results"):
            full_frame = " -FullFrame" if suite == "frame-product-all" else ""
            if suite == "compare-results":
                full_frame = " -CompareResults"
            if cases == "registry-override":
                full_frame += " -RegistryOverride"
            script = f'''$ErrorActionPreference = "Stop"
$source=Join-Path $PSScriptRoot "source-{tag}"
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source "tests/windows/Run-R66FrameProbe.ps1") -SourceRoot $source -ProductXlam (Join-Path $PSScriptRoot "out/complete/Product.xlam") -EvidenceRoot (Join-Path $PSScriptRoot "evidence-{tag}"){full_frame}
exit $LASTEXITCODE
'''
        if suite == "r68-product":
            requested = cases or "performance,table,shortcuts,quickformat"
            baseline = requested == "baseline"
            override = requested == "performance-override"
            if baseline:
                requested = "performance"
            if override:
                requested = "performance"
            case_filter = ""
            if re.fullmatch(r"(?:compare-consolidate-links|consolidate-list|r94-template|table|output-flows|popup-layout)\.[A-Za-z0-9_]+", requested):
                requested, case_filter = requested.split(".", 1)
            if any(value not in {"r76-format-backup", "r76-policy", "r75-guard", "r75-journal", "r74-folder", "r74-features", "performance", "table", "shortcuts", "quickformat", "visual", "benchmark", "runner-benchmark", "runner-safety", "repaint-benchmark", "snapshot-probe", "compare-internal", "compare-enhanced", "userflows", "popup-layout", "output-flows", "large-range", "alignment", "serialization-probe", "fast-snapshot", "stage-profile", "range-properties"} for value in requested.split(",")):
                if requested not in ("compare-consolidate-links", "consolidate-list", "unit-symbols", "privacy-flow", "r100-focus-preview", "r97-preview", "r96-ribbon", "r94-compare-output", "r94-formulatools", "r94-functions", "r94-template", "r93-output", "r93-compare", "r77-normalization", "r78-files", "r79-table", "r80-table", "r81-table", "r82-file", "r83-files", "r84-ai", "r85-data", "r86-data", "r86-policy"):
                    raise SystemExit("unknown r68 product case group")
            flag = " -ObserveBaseline" if baseline else ""
            if case_filter:
                flag += f' -CaseName "{case_filter}"'
            if override:
                flag += " -OverridePerformance"
            script = f'''$ErrorActionPreference = "Stop"
$source=Join-Path $PSScriptRoot "source-{tag}"
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source "tests/windows/Run-R68ProductSuite.ps1") -SourceRoot $source -ProductXlam (Join-Path $PSScriptRoot "out/complete/Product.xlam") -EvidenceRoot (Join-Path $PSScriptRoot "evidence-{tag}") -Suites "{requested}"{flag}
exit $LASTEXITCODE
'''
            if requested in {"compare-internal", "compare-enhanced"}:
                invocation = script.splitlines()[2]
                revision = paths.version.split("_")[-1]
                script = f'''$ErrorActionPreference = "Stop"
$source=Join-Path $PSScriptRoot "source-{tag}"
$inputs=Join-Path $PSScriptRoot "out/complete"
$state=Join-Path $env:LOCALAPPDATA "LHexcel/NxHost/state/{revision}/registry-state.json"
$runtime=Join-Path $env:LOCALAPPDATA "LHexcel/NxHost/{revision}"
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -or (Test-Path -LiteralPath $runtime)) {{throw "Existing Excel/runtime preserved"}}
$registered=$false
try {{
    & (Join-Path $source "build/Register-NxHost.ps1") -Action Register -OptIn -X86Dll (Join-Path $inputs "NxHost32.dll") -X64Dll (Join-Path $inputs "NxHost64.dll")
    $registered=$true
    $before=(Get-Content -LiteralPath $state -Raw | ConvertFrom-Json).snapshot_sha256
    {invocation}
    $code=$LASTEXITCODE
}} finally {{
    if($registered) {{
        & (Join-Path $source "build/Register-NxHost.ps1") -Action Unregister -X86Dll (Join-Path $inputs "NxHost32.dll") -X64Dll (Join-Path $inputs "NxHost64.dll")
        $after=(Get-Content -LiteralPath $state -Raw | ConvertFrom-Json).restored_snapshot_sha256
        if($before -ne $after) {{throw "Registry restore mismatch"}}
    }}
}}
exit $code
'''
        if suite == "picture-state":
            script = f'''$ErrorActionPreference = "Stop"
$source=Join-Path $PSScriptRoot "source-{tag}"
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source "tests/windows/Run-R62WorkflowSuite.ps1") -ProductXlam (Join-Path $PSScriptRoot "out/complete/Product.xlam") -EvidenceRoot (Join-Path $PSScriptRoot "evidence-{tag}") -CaseNames "picture.picker_state"
exit $LASTEXITCODE
'''
        write_ascii(ps_path, script)
        write_ascii(bat_path, v001_release_run_bat_text(paths).replace("-File build.ps1", f"-File verify-{tag}.ps1"), crlf=True)
        print(json.dumps({"verification_bat": str(bat_path), "built_product_sha256": file_sha256(product)}))
        return
    if suite == "compare-benchmark":
        snapshot = paths.tmp_dir / f"source-{tag}"
        for relative in ("tests/windows/Run-R65CompareBenchmark.ps1", "tests/vba/product/CNxR66CompareProgress.cls", "build/Excel-ProcessLifecycle.ps1", "build/Register-NxHost.ps1"):
            destination = snapshot / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(paths.source_root / relative, destination)
        inputs = snapshot / "inputs"
        inputs.mkdir()
        shutil.copy2(product, inputs / "Product.xlam")
        dll_source = host_root.resolve() if host_root else paths.tmp_dir / "out/complete"
        for name in ("NxHost32.dll", "NxHost64.dll", "NxCore32.dll", "NxCore64.dll"):
            shutil.copy2(dll_source / name, inputs / name)
        revision = paths.version.rsplit("_", 1)[1]
        script = f'''$ErrorActionPreference = "Stop"
$source=Join-Path $PSScriptRoot "source-{tag}"
$inputs=Join-Path $source "inputs"
$state=Join-Path $env:LOCALAPPDATA "LHexcel/NxHost/state/{revision}/registry-state.json"
$runtime=Join-Path $env:LOCALAPPDATA "LHexcel/NxHost/{revision}"
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -or (Test-Path -LiteralPath $runtime)) {{throw "Existing Excel/runtime preserved"}}
$registered=$false
try {{
    & (Join-Path $source "build/Register-NxHost.ps1") -Action Register -OptIn -X86Dll (Join-Path $inputs "NxHost32.dll") -X64Dll (Join-Path $inputs "NxHost64.dll")
    $registered=$true
    $before=(Get-Content -LiteralPath $state -Raw | ConvertFrom-Json).snapshot_sha256
    & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source "tests/windows/Run-R65CompareBenchmark.ps1") -ProductXlam (Join-Path $inputs "Product.xlam") -EvidenceRoot (Join-Path $PSScriptRoot "evidence-{tag}")
    if($LASTEXITCODE -ne 0){{throw "Native comparison benchmark failed"}}
}} finally {{
    if($registered){{
        & (Join-Path $source "build/Register-NxHost.ps1") -Action Unregister -X86Dll (Join-Path $inputs "NxHost32.dll") -X64Dll (Join-Path $inputs "NxHost64.dll")
        $after=(Get-Content -LiteralPath $state -Raw | ConvertFrom-Json).restored_snapshot_sha256
        if($before -ne $after){{throw "Registration restore mismatch"}}
    }}
}}
exit 0
'''
        write_ascii(ps_path, script)
        write_ascii(bat_path, v001_release_run_bat_text(paths).replace("-File build.ps1", f"-File verify-{tag}.ps1"), crlf=True)
        print(json.dumps({"verification_bat": str(bat_path), "built_product_sha256": file_sha256(product)}))
        return
    if suite in ("dll-hwpx", "dll-visual", "dll-navigator"):
        visual = suite in ("dll-visual", "dll-navigator")
        harness = "Run-NxHostVisualSession.ps1" if visual else "Run-NxHwpxNativeSuite.ps1"
        files = (f"tests/windows/{harness}", "build/Excel-ProcessLifecycle.ps1", "build/Register-NxHost.ps1") if visual else (f"tests/windows/{harness}", "tests/vba/hwpx/T_NxHwpxNative.bas")
        snapshot = paths.tmp_dir / (f"source-{tag}" if suite == "dll-navigator" else "source")
        if suite == "dll-navigator" and snapshot.exists():
            raise SystemExit("existing verification source snapshot is rejected")
        if suite == "dll-navigator":
            files += ("tests/windows/Test-NxNavigatorUi.ps1", "tests/windows/Test-NxNavigatorComContract.ps1", "tests/windows/NxNavigatorComProbe.cs", "build/Build-NxHost.ps1")
            shutil.copytree(paths.source_root / "src/dotnet/NxHost", snapshot / "src/dotnet/NxHost")
            if cases not in ("", "fallback", "enhancements", "document", "services", "services-isolated", "services-no-other-addins", "manual", "picture-manual"):
                raise SystemExit("unknown navigator test case")
            if cases in ("enhancements", "document", "services", "services-isolated", "services-no-other-addins"):
                files += ("tests/windows/Test-R65EnhancementUi.ps1", "tests/windows/R65PictureWatcher.cs", "tests/vba/product/T_R65Enhancement.bas", "tests/vba/product/CNxR65CancelCompare.cls")
            if cases in ("services", "services-isolated", "services-no-other-addins"):
                files += ("tests/fixtures/r60-document-navigator/simple.xlsx",)
            if cases == "fallback":
                # Fault injection is confined to this disposable test snapshot.
                # No production toggle, registry policy change, or shipping DLL is modified.
                host_source = snapshot / "src/dotnet/NxHost/NxHostAddIn.cs"
                text = host_source.read_text(encoding="utf-8-sig")
                needle = "ctpFactory.CreateCTP(NavigatorProgId, NavigatorTitle, activeWindow)"
                if text.count(needle) != 1:
                    raise SystemExit("fallback fixture injection target changed")
                host_source.write_text(text.replace(needle, 'ctpFactory.CreateCTP(NavigatorProgId + ".MissingFixture", NavigatorTitle, activeWindow)'), encoding="utf-8")
        for relative in files:
            destination = snapshot / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(paths.source_root / relative, destination)
        host_path = '"out/complete"'
        evidence_path = f'(Join-Path $PSScriptRoot "evidence-{tag}")'
        build = ""
        extra = ""
        if suite == "dll-navigator":
            host_path = f'"host-{tag}"'
            # Registry backup basenames are long; keep native PS 5.1 paths below MAX_PATH.
            evidence_path = f'(Join-Path $env:LOCALAPPDATA "LHExcel/verification/{tag}")'
            # Manual observation keeps fixture/lifecycle checks, but leaves UI
            # inputs to the active computer-use session instead of a second driver.
            extra = "" if cases in ("manual", "picture-manual") else " -Automated"
            if cases == "picture-manual":
                extra += " -PictureFixture"
            if cases == "fallback":
                extra += " -ExpectFallback"
            if cases == "enhancements":
                extra += " -Enhancements"
            if cases == "document":
                extra += " -DocumentOnly"
            if cases in ("services", "services-isolated", "services-no-other-addins"):
                extra += " -Enhancements -ServicesOnly"
            if cases == "services-isolated":
                extra += " -SafeMode"
            if cases == "services-no-other-addins":
                extra += " -IsolateOtherAddins"
            build = f'''$hostRoot = Join-Path $PSScriptRoot {host_path}
if (Test-Path -LiteralPath $hostRoot) {{ throw "Fresh DLL output required" }}
& (Join-Path $PSScriptRoot "source-{tag}/build/Build-NxHost.ps1") -OutputRoot $hostRoot -ProtectedLoader
Copy-Item -LiteralPath (Join-Path $hostRoot "x86/NxHost32.dll") -Destination $hostRoot
Copy-Item -LiteralPath (Join-Path $hostRoot "x64/NxHost64.dll") -Destination $hostRoot
Copy-Item -LiteralPath (Join-Path $hostRoot "x86/NxCore32.dll") -Destination $hostRoot
Copy-Item -LiteralPath (Join-Path $hostRoot "x64/NxCore64.dll") -Destination $hostRoot
'''
        script = f'''$ErrorActionPreference = "Stop"
{build}& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "{snapshot.name}/tests/windows/{harness}") -ProductXlam (Join-Path $PSScriptRoot "out/complete/Product.xlam") -NxHostRoot (Join-Path $PSScriptRoot {host_path}) -EvidenceRoot {evidence_path}{extra}
exit $LASTEXITCODE
'''
        if artifact or host_root:
            shutil.copy2(product, paths.tmp_dir / f"product-{tag}.xlam")
            script = script.replace('"out/complete/Product.xlam"', f'"product-{tag}.xlam"')
        if host_root:
            exact_host = paths.tmp_dir / f"exact-host-{tag}"
            exact_host.mkdir(exist_ok=False)
            for name in ("NxHost32.dll", "NxHost64.dll", "NxCore32.dll", "NxCore64.dll"):
                shutil.copy2(host_root / name, exact_host / name)
            script = script.replace(build, "").replace(host_path, f'"{exact_host.name}"')
        write_ascii(ps_path, script)
        write_ascii(bat_path, v001_release_run_bat_text(paths).replace("-File build.ps1", f"-File verify-{tag}.ps1"), crlf=True)
        print(json.dumps({"verification_bat": str(bat_path), "built_product_sha256": file_sha256(product)}))
        return
    if suite == "data-compare":
        snapshot = paths.tmp_dir / f"source-{tag}"
        for relative in ("tests/windows/Run-DataCompareSuite.ps1", "tests/vba/product/T_DataCompare.bas", "tests/vba/product/CNxDataCompareCancelProbe.cls", "build/Excel-ProcessLifecycle.ps1"):
            destination = snapshot / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(paths.source_root / relative, destination)
        script = f'''$ErrorActionPreference = "Stop"
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "source-{tag}/tests/windows/Run-DataCompareSuite.ps1") -ProductXlam (Join-Path $PSScriptRoot "out/complete/Product.xlam") -EvidenceRoot (Join-Path $PSScriptRoot "evidence-{tag}")
exit $LASTEXITCODE
'''
        write_ascii(ps_path, script)
        write_ascii(bat_path, v001_release_run_bat_text(paths).replace("-File build.ps1", f"-File verify-{tag}.ps1"), crlf=True)
        print(json.dumps({"verification_bat": str(bat_path), "built_product_sha256": file_sha256(product)}))
        return
    if suite == "focus-speed":
        if cases not in ("baseline", "duplicate", "visible", "combined", "all", "product", "product-quick", "paint-batch-quick", "name-batch-quick", "product-visual"):
            raise SystemExit("unknown focus speed variant")
        speed_variant = cases.removesuffix("-quick")
        speed_options = " -MoveCount 24 -DuplicateCount 12" if cases.endswith("-quick") else ""
        if cases == "product-visual":
            speed_variant = "product"
            speed_options = " -MoveCount 20 -DuplicateCount 10 -PauseForVisual"
        snapshot = paths.tmp_dir / f"source-{tag}"
        for relative in ("tests/windows/Run-FocusSpeedSuite.ps1", "build/Excel-ProcessLifecycle.ps1",
                         "tests/fixtures/r60-document-navigator/simple.xlsx",
                         "src/vba/features/data/focus/CNxFocusRuleEngine.cls", "src/vba/features/data/focus/NxFocusController.bas"):
            destination = snapshot / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(paths.source_root / relative, destination)
        target_product = paths.tmp_dir / f"product-{tag}.xlam"
        shutil.copy2(product, target_product)
        script = f'''$ErrorActionPreference = "Stop"
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "source-{tag}/tests/windows/Run-FocusSpeedSuite.ps1") -ProductXlam (Join-Path $PSScriptRoot "product-{tag}.xlam") -EvidenceRoot (Join-Path $PSScriptRoot "evidence-{tag}") -Variant "{speed_variant}"{speed_options}
exit $LASTEXITCODE
'''
        write_ascii(ps_path, script)
        write_ascii(bat_path, v001_release_run_bat_text(paths).replace("-File build.ps1", f"-File verify-{tag}.ps1"), crlf=True)
        print(json.dumps({"verification_bat": str(bat_path), "built_product_sha256": file_sha256(product)}))
        return
    if suite == "focus-copy":
        if cases not in ("", "accepted-suspension", "memory_cache", "memory_no_calculate", "memory_no_calculate_visual", "memory_no_calculate_repaint_visual"):
            raise SystemExit("unknown focus-copy experiment")
        extra = " -MemoryCacheProbe" if cases.startswith("memory_") else ""
        if cases == "accepted-suspension":
            extra += " -AcceptClipboardSuspension"
        if cases.startswith("memory_no_calculate"):
            extra += " -NoCalculateProbe"
        if cases.endswith("_visual"):
            extra += " -PauseForVisual"
        if "_repaint_" in cases:
            extra += " -RepaintProbe"
        relative = "tests/windows/Run-FocusCopyPasteSuite.ps1"
        shutil.copy2(paths.source_root / relative, paths.tmp_dir / "source" / relative)
        shutil.copy2(paths.source_root / "build/Excel-ProcessLifecycle.ps1", paths.tmp_dir / "source/build/Excel-ProcessLifecycle.ps1")
        script = f'''$ErrorActionPreference = "Stop"
& powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "source/{relative}") -ProductXlam (Join-Path $PSScriptRoot "out/complete/Product.xlam") -EvidenceRoot (Join-Path $PSScriptRoot "evidence-{tag}"){extra}
exit $LASTEXITCODE
'''
        write_ascii(ps_path, script)
        write_ascii(bat_path, v001_release_run_bat_text(paths).replace("-File build.ps1", f"-File verify-{tag}.ps1"), crlf=True)
        print(json.dumps({"verification_bat": str(bat_path), "built_product_sha256": file_sha256(product)}))
        return
    for relative in ("tests/windows/Run-ProductRibbonSuite.ps1", "tests/windows/Run-ProductUiSuite.ps1",
                     "tools/inspect_windows_return.py", "tests/windows/Run-R59MannerSuite.ps1",
                     "tests/vba/product/T_R59MannerSave.bas", "tests/vba/product/CNxR59CancelSave.cls",
                     "tests/vba/product/T_R59SaveExceptions.bas"):
        shutil.copy2(paths.source_root / relative, paths.tmp_dir / "source" / relative)
    script = v001_release_build_ps1_text(paths.version)
    script = script.replace('if (Test-Path -LiteralPath $outRoot) { throw "Existing v0.01 build output is rejected; run prepare again" }', '')
    start = script.index('& (Join-Path $sourceRoot "build/Build-Xlam.ps1")')
    end = script.index('$sourceDigest =', start)
    script = script[:start] + script[end:]
    if suite == "manner":
        start = script.index('Invoke-CheckedPowerShell "product-ui"')
        end = script.index('Invoke-CheckedPowerShell "r59-manner"', start)
        script = script[:start] + script[end:]
    elif suite == "ui":
        start = script.index('Invoke-CheckedPowerShell "r59-manner"')
        end = script.index('$product | Set-Content', start)
        script = script[:start] + script[end:]
    if cases:
        if not re.fullmatch(r"[a-z_,]+", cases):
            raise SystemExit("case names must be lowercase ASCII")
        script = script.replace('"-ProductXlam", $product,', f'"-CaseNames", "{cases}", "-ProductXlam", $product,')
    script = script.replace('"evidence"', f'"evidence-{tag}"').replace('"logs"', f'"logs-{tag}"')
    script = script.replace('"windows-build-direct-run.log"', f'"windows-verify-{tag}.log"')
    # Guard the built source independently of the instrumented test copies.
    script = script.replace('$sourceDigest =', '$beforeProductHash = (Get-FileHash -LiteralPath $product -Algorithm SHA256).Hash\n$sourceDigest =', 1)
    script = script.replace('$product | Set-Content', 'if ((Get-FileHash -LiteralPath $product -Algorithm SHA256).Hash -ne $beforeProductHash) { throw "Verification changed built Product.xlam" }\n$product | Set-Content', 1)
    write_ascii(ps_path, script)
    bat = v001_release_run_bat_text(paths).replace('-File build.ps1', f'-File verify-{tag}.ps1')
    write_ascii(bat_path, bat, crlf=True)
    print(json.dumps({"verification_bat": str(bat_path), "built_product_sha256": file_sha256(product)}))


def prepare_v001_distribution_smoke(paths: V001ReleasePaths, tag: str, internal: Path, enhanced: Path, defer_artifact_check: bool = False, expected_policy: str = "active", skip_no_dll_focus: bool = False) -> None:
    """Verify the two protected artifacts without installing or registering them."""
    if not re.fullmatch(r"[a-z0-9-]+", tag):
        raise SystemExit("distribution tag must be lowercase ASCII")
    if expected_policy not in ("active", "expired", "invalid"):
        raise SystemExit("Unknown policy state")
    policy_flags = "-VerifyNoDllFocus" if expected_policy == "active" else "-ExpectedPolicy " + expected_policy
    if expected_policy == "active" and skip_no_dll_focus:
        policy_flags = ""
    artifacts = [internal.resolve(), enhanced.resolve()]
    if any(path.suffix.lower() != ".xlam" for path in artifacts):
        raise SystemExit("both protected distribution paths must name XLAM files")
    if not defer_artifact_check and any(not path.is_file() for path in artifacts):
        raise SystemExit("both protected distribution artifacts must exist")
    evidence = paths.tmp_dir / f"distribution-{tag}"
    ps_path = paths.tmp_dir / f"distribution-{tag}.ps1"
    bat_path = paths.tmp_dir / f"distribution-{tag}.bat"
    if evidence.exists() or ps_path.exists() or bat_path.exists():
        raise SystemExit("existing distribution runner/evidence is rejected")
    relative = "tests/windows/Run-R57DistributionSmokeSuite.ps1"
    shutil.copy2(paths.source_root / relative, paths.tmp_dir / "source" / relative)
    encoded = [base64.b64encode(str(path).encode("utf-8")).decode("ascii") for path in artifacts]
    script = f'''$ErrorActionPreference = "Stop"
$root = $PSScriptRoot
$sourceRoot = Join-Path $root "source"
$evidenceRoot = Join-Path $root "distribution-{tag}"
if (Test-Path -LiteralPath $evidenceRoot) {{ throw "Existing distribution evidence rejected" }}
if (Get-Process EXCEL -ErrorAction SilentlyContinue) {{ throw "Close user Excel first" }}
[void](New-Item -ItemType Directory -Path $evidenceRoot)
$encoded = @("{encoded[0]}", "{encoded[1]}")
$names = @("internal", "enhanced")
$profiles = @("internal-xlam", "enhanced-dll")
for ($index = 0; $index -lt 2; $index++) {{
    $artifact = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encoded[$index]))
    if (-not (Test-Path -LiteralPath $artifact -PathType Leaf)) {{ throw "Protected distribution artifact is missing" }}
    $caseEvidence = Join-Path $evidenceRoot $names[$index]
    & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $sourceRoot "{relative}") -EvidenceRoot $caseEvidence -ArtifactPath $artifact -ExpectedVersion "{paths.version}" -ExpectedProfile $profiles[$index] {policy_flags}
    if ($LASTEXITCODE -ne 0) {{ throw ("Protected distribution failed: " + $names[$index]) }}
}}
Write-Output "PASS|ProtectedDistribution|2/2"
exit 0
'''
    write_ascii(ps_path, script)
    write_ascii(bat_path, v001_release_run_bat_text(paths).replace("-File build.ps1", f"-File distribution-{tag}.ps1"), crlf=True)
    print(json.dumps({"distribution_bat": str(bat_path), "artifact_check_deferred": defer_artifact_check, "source_hashes": [file_sha256(path) if path.is_file() else None for path in artifacts]}))


def status_v001_release(paths: V001ReleasePaths) -> None:
    files = [
        paths.tmp_dir / "run.bat",
        paths.tmp_dir / "build.ps1",
        paths.tmp_dir / "windows-build-direct-run.log",
        paths.tmp_dir / "out" / "complete" / "Product.xlam",
        paths.tmp_dir / "out" / "complete" / "NxHost32.dll",
        paths.tmp_dir / "out" / "complete" / "NxHost64.dll",
        paths.tmp_dir / "latest-output.txt",
    ]
    for path in files:
        if path.exists():
            print(f"OK {path.stat().st_mtime:.0f} {path.stat().st_size:>10} {path}")
        else:
            print(f"-- {'':>10} {path}")


def build_ps1_text(
    *,
    run_focus_probe: bool,
    run_shortcut_probe: bool,
    run_style_probe: bool,
    run_batch_rename_probe: bool,
    run_folder_create_probe: bool,
    run_batch_tools_ui_probe: bool,
    run_picture_fit_probe: bool,
    run_hangul_probe: bool,
) -> str:
    text = """$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceXlam = Join-Path $root "source.xlam"
$sourceDir = Join-Path $root "src"
$outDir = Join-Path $root "out"
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$candidate = Join-Path $outDir ("lhexcel_candidate_" + $stamp + ".xlam")
$builder = Join-Path $root "lhexcel_dev.ps1"
$focusProbe = Join-Path $root "lhexcel_focus_runtime_probe.ps1"
$shortcutProbe = Join-Path $root "lhexcel_shortcut_runtime_probe.ps1"
$styleProbe = Join-Path $root "lhexcel_style_runtime_probe.ps1"
$batchRenameProbe = Join-Path $root "lhexcel_batch_rename_runtime_probe.ps1"
$folderCreateProbe = Join-Path $root "lhexcel_folder_create_runtime_probe.ps1"
$batchToolsUiProbe = Join-Path $root "lhexcel_batch_tools_ui_runtime_probe.ps1"
$pictureFitProbe = Join-Path $root "lhexcel_picture_fit_runtime_probe.ps1"
$hangulProbe = Join-Path $root "lhexcel_hangul_runtime_probe.ps1"
$dialogAction = Join-Path $root "lhexcel_dialog_action.ps1"
$log = Join-Path $root "windows-build-direct-run.log"
$focusLog = Join-Path $root "windows-focus-runtime.log"
$shortcutLog = Join-Path $root "windows-shortcut-runtime.log"
$styleLog = Join-Path $root "windows-style-runtime.log"
$styleJson = Join-Path $root "windows-style-runtime.json"
$batchRenameLog = Join-Path $root "windows-batch-rename-runtime.log"
$folderCreateLog = Join-Path $root "windows-folder-create-runtime.log"
$batchToolsUiLog = Join-Path $root "windows-batch-tools-ui-runtime.log"
$pictureFitLog = Join-Path $root "windows-picture-fit-runtime.log"
$hangulLog = Join-Path $root "windows-hangul-runtime.log"
$latest = Join-Path $root "latest-output.txt"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Remove-Item -LiteralPath $latest -Force -ErrorAction SilentlyContinue
$env:LHEXCEL_SKIP_FORM_REPAIR = "1"

"$(Get-Date -Format o) START ascii direct build" | Set-Content -Path $log -Encoding UTF8
"root=$root" | Add-Content -Path $log -Encoding UTF8
"sourceXlam=$sourceXlam" | Add-Content -Path $log -Encoding UTF8
"sourceDir=$sourceDir" | Add-Content -Path $log -Encoding UTF8
"candidate=$candidate" | Add-Content -Path $log -Encoding UTF8
"skipFormRepair=$env:LHEXCEL_SKIP_FORM_REPAIR" | Add-Content -Path $log -Encoding UTF8

$code = 0
try {
    & $builder -Mode Build -SourceXlam $sourceXlam -SourceDir $sourceDir -CandidateXlam $candidate -RemoveMissingModules *>> $log
    if ($null -ne $LASTEXITCODE) {
        $code = [int]$LASTEXITCODE
    }
} catch {
    $code = 1
    "ERROR: $($_.Exception.Message)" | Add-Content -Path $log -Encoding UTF8
    $_ | Format-List * -Force | Out-String | Add-Content -Path $log -Encoding UTF8
}

if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $focusProbe `
        -Source $candidate -LogPath $focusLog -DialogActionScript $dialogAction *> $null
    $focusCode = [int]$LASTEXITCODE
    "focusRuntimeAttempt=1 exit=$focusCode" | Add-Content -Path $log -Encoding UTF8
    if ($focusCode -ne 0) {
        if (Test-Path -LiteralPath $focusLog) {
            Copy-Item -LiteralPath $focusLog -Destination ($focusLog + ".attempt1.log") -Force
        }
        Start-Sleep -Seconds 3
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $focusProbe `
            -Source $candidate -LogPath $focusLog -DialogActionScript $dialogAction *> $null
        $focusCode = [int]$LASTEXITCODE
        "focusRuntimeAttempt=2 exit=$focusCode" | Add-Content -Path $log -Encoding UTF8
    }
    "focusRuntimeExit=$focusCode" | Add-Content -Path $log -Encoding UTF8
    if ($focusCode -ne 0) { $code = 2 }
}

if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $shortcutProbe `
        -Source $candidate -LogPath $shortcutLog *> $null
    $shortcutCode = [int]$LASTEXITCODE
    "shortcutRuntimeExit=$shortcutCode" | Add-Content -Path $log -Encoding UTF8
    if ($shortcutCode -ne 0) { $code = 3 }
}

if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $styleProbe `
        -Source $candidate -LogPath $styleLog -JsonPath $styleJson *> $null
    $styleCode = [int]$LASTEXITCODE
    "styleRuntimeExit=$styleCode" | Add-Content -Path $log -Encoding UTF8
    if ($styleCode -ne 0) { $code = 7 }
}

if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $batchRenameProbe `
        -Source $candidate -LogPath $batchRenameLog *> $null
    $batchRenameCode = [int]$LASTEXITCODE
    "batchRenameRuntimeExit=$batchRenameCode" | Add-Content -Path $log -Encoding UTF8
    if ($batchRenameCode -ne 0) { $code = 4 }
}

if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $folderCreateProbe `
        -Source $candidate -LogPath $folderCreateLog *> $null
    $folderCreateCode = [int]$LASTEXITCODE
    "folderCreateRuntimeExit=$folderCreateCode" | Add-Content -Path $log -Encoding UTF8
    if ($folderCreateCode -ne 0) { $code = 5 }
}

if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $batchToolsUiProbe `
        -Source $candidate -LogPath $batchToolsUiLog *> $null
    $batchToolsUiCode = [int]$LASTEXITCODE
    "batchToolsUiRuntimeExit=$batchToolsUiCode" | Add-Content -Path $log -Encoding UTF8
    if ($batchToolsUiCode -ne 0) { $code = 6 }
}

if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $pictureFitProbe `
        -Source $candidate -LogPath $pictureFitLog *> $null
    $pictureFitCode = [int]$LASTEXITCODE
    "pictureFitRuntimeExit=$pictureFitCode" | Add-Content -Path $log -Encoding UTF8
    if ($pictureFitCode -ne 0) { $code = 8 }
}

if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    $hangulArguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $hangulProbe + '" -Source "' + $candidate + '" -LogPath "' + $hangulLog + '"'
    $hangulProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $hangulArguments -WindowStyle Hidden -PassThru
    if (-not $hangulProcess.WaitForExit(120000)) {
        & taskkill.exe /PID $hangulProcess.Id /T /F *> $null
        "HANGUL_TIMEOUT seconds=120" | Add-Content -Path $hangulLog -Encoding UTF8
        $hangulCode = 124
    } else {
        $hangulCode = [int]$hangulProcess.ExitCode
    }
    "hangulRuntimeExit=$hangulCode" | Add-Content -Path $log -Encoding UTF8
    if ($hangulCode -ne 0) { $code = 9 }
}

"$(Get-Date -Format o) EXIT code=$code" | Add-Content -Path $log -Encoding UTF8
if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    $candidate | Set-Content -Path $latest -Encoding UTF8
}
exit $code
"""
    if not run_focus_probe:
        focus_block = """if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $focusProbe `
        -Source $candidate -LogPath $focusLog -DialogActionScript $dialogAction *> $null
    $focusCode = [int]$LASTEXITCODE
    "focusRuntimeAttempt=1 exit=$focusCode" | Add-Content -Path $log -Encoding UTF8
    if ($focusCode -ne 0) {
        if (Test-Path -LiteralPath $focusLog) {
            Copy-Item -LiteralPath $focusLog -Destination ($focusLog + ".attempt1.log") -Force
        }
        Start-Sleep -Seconds 3
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $focusProbe `
            -Source $candidate -LogPath $focusLog -DialogActionScript $dialogAction *> $null
        $focusCode = [int]$LASTEXITCODE
        "focusRuntimeAttempt=2 exit=$focusCode" | Add-Content -Path $log -Encoding UTF8
    }
    "focusRuntimeExit=$focusCode" | Add-Content -Path $log -Encoding UTF8
    if ($focusCode -ne 0) { $code = 2 }
}

"""
        text = text.replace(focus_block, "")
    if not run_shortcut_probe:
        shortcut_block = """if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $shortcutProbe `
        -Source $candidate -LogPath $shortcutLog *> $null
    $shortcutCode = [int]$LASTEXITCODE
    "shortcutRuntimeExit=$shortcutCode" | Add-Content -Path $log -Encoding UTF8
    if ($shortcutCode -ne 0) { $code = 3 }
}

"""
        text = text.replace(shortcut_block, "")
    if not run_style_probe:
        style_block = """if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $styleProbe `
        -Source $candidate -LogPath $styleLog -JsonPath $styleJson *> $null
    $styleCode = [int]$LASTEXITCODE
    "styleRuntimeExit=$styleCode" | Add-Content -Path $log -Encoding UTF8
    if ($styleCode -ne 0) { $code = 7 }
}

"""
        text = text.replace(style_block, "")
    if not run_batch_rename_probe:
        batch_rename_block = """if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $batchRenameProbe `
        -Source $candidate -LogPath $batchRenameLog *> $null
    $batchRenameCode = [int]$LASTEXITCODE
    "batchRenameRuntimeExit=$batchRenameCode" | Add-Content -Path $log -Encoding UTF8
    if ($batchRenameCode -ne 0) { $code = 4 }
}

"""
        text = text.replace(batch_rename_block, "")
    if not run_folder_create_probe:
        folder_create_block = """if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $folderCreateProbe `
        -Source $candidate -LogPath $folderCreateLog *> $null
    $folderCreateCode = [int]$LASTEXITCODE
    "folderCreateRuntimeExit=$folderCreateCode" | Add-Content -Path $log -Encoding UTF8
    if ($folderCreateCode -ne 0) { $code = 5 }
}

"""
        text = text.replace(folder_create_block, "")
    if not run_batch_tools_ui_probe:
        batch_tools_ui_block = """if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $batchToolsUiProbe `
        -Source $candidate -LogPath $batchToolsUiLog *> $null
    $batchToolsUiCode = [int]$LASTEXITCODE
    "batchToolsUiRuntimeExit=$batchToolsUiCode" | Add-Content -Path $log -Encoding UTF8
    if ($batchToolsUiCode -ne 0) { $code = 6 }
}

"""
        text = text.replace(batch_tools_ui_block, "")
    if not run_picture_fit_probe:
        picture_fit_block = """if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $pictureFitProbe `
        -Source $candidate -LogPath $pictureFitLog *> $null
    $pictureFitCode = [int]$LASTEXITCODE
    "pictureFitRuntimeExit=$pictureFitCode" | Add-Content -Path $log -Encoding UTF8
    if ($pictureFitCode -ne 0) { $code = 8 }
}

"""
        text = text.replace(picture_fit_block, "")
    if not run_hangul_probe:
        hangul_block = """if ($code -eq 0 -and (Test-Path -LiteralPath $candidate)) {
    $hangulArguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $hangulProbe + '" -Source "' + $candidate + '" -LogPath "' + $hangulLog + '"'
    $hangulProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $hangulArguments -WindowStyle Hidden -PassThru
    if (-not $hangulProcess.WaitForExit(120000)) {
        & taskkill.exe /PID $hangulProcess.Id /T /F *> $null
        "HANGUL_TIMEOUT seconds=120" | Add-Content -Path $hangulLog -Encoding UTF8
        $hangulCode = 124
    } else {
        $hangulCode = [int]$hangulProcess.ExitCode
    }
    "hangulRuntimeExit=$hangulCode" | Add-Content -Path $log -Encoding UTF8
    if ($hangulCode -ne 0) { $code = 9 }
}

"""
        text = text.replace(hangul_block, "")
    return text


def run_bat_text(tmp_dir: Path) -> str:
    win_dir = to_windows_path(tmp_dir)
    return f"""@echo off
cd /d {win_dir}
echo %DATE% %TIME% > runner-started.txt
echo START LHExcel direct build
echo WORKDIR=%CD%
powershell.exe -NoProfile -ExecutionPolicy Bypass -File build.ps1
set RC=%ERRORLEVEL%
echo.
echo EXIT CODE=%RC%
echo LOG={win_dir}\\windows-build-direct-run.log
echo LATEST={win_dir}\\latest-output.txt
echo.
pause
exit /b %RC%
"""


def launcher_bat_text(paths: VersionPaths) -> str:
    win_dir = "%~dp0" + paths.tmp_dir.relative_to(paths.launcher_bat.parent).as_posix().replace("/", "\\")
    return f"""@echo off
setlocal
set "TARGET={win_dir}"
set "LAUNCH_LOG=%~dp0run_{paths.token}_launcher.log"
> "%LAUNCH_LOG%" echo %DATE% %TIME% START LHExcel launcher
>> "%LAUNCH_LOG%" echo TARGET=%TARGET%
if not exist "%TARGET%\\run.bat" (
  echo LHExcel build runner was not found.
  echo TARGET=%TARGET%
  echo Check that the Mac home folder is shared with Windows.
  >> "%LAUNCH_LOG%" echo ERROR runner_not_found
  echo.
  pause
  endlocal
  exit /b 2
)
call "%TARGET%\\run.bat"
set "RC=%ERRORLEVEL%"
>> "%LAUNCH_LOG%" echo EXIT code=%RC%
endlocal & exit /b %RC%
"""


def prepare(paths: VersionPaths, *, clean: bool) -> None:
    from lhexcel_static_ribbon_catalog import inject_static_catalog

    sync_generated_sources(paths, check=False)
    if paths.version == "v3.4":
        paths.reports_dir.mkdir(parents=True, exist_ok=True)
        run_checked(
            ["python3", str(SCRIPT_DIR / "lhexcel_shortcut_key_contract.py"), "--source", str(paths.source_dir / "mdShortcutManager.bas")],
            ROOT,
            paths.reports_dir / "shortcut-key-contract-v34-prepare.json",
        )
        run_checked(
            [
                "python3",
                str(SCRIPT_DIR / "lhexcel_v34_style_system_contract.py"),
                "--source",
                str(paths.source_dir),
            ],
            ROOT,
            paths.reports_dir / "style-system-contract-v34-prepare.json",
        )
        run_checked(
            ["python3", str(SCRIPT_DIR / "lhexcel_v34_picture_fit_contract.py")],
            ROOT,
            paths.reports_dir / "picture-fit-contract-v34-prepare.json",
        )
        run_checked(
            [
                "python3",
                str(SCRIPT_DIR / "lhexcel_v34_hangul_export_contract.py"),
                "--source",
                str(paths.source_dir),
            ],
            ROOT,
            paths.reports_dir / "hangul-export-contract-v34-prepare.json",
        )
    if clean and paths.tmp_dir.exists():
        shutil.rmtree(paths.tmp_dir)
    paths.tmp_dir.mkdir(parents=True, exist_ok=True)
    out_dir = paths.tmp_dir / "out"
    out_dir.mkdir(exist_ok=True)

    staged_source = paths.tmp_dir / "source.xlam"
    catalog_source = paths.source_xlam
    prepared_base = paths.tmp_dir / "prepared-base.xlam"
    if paths.version in {"v3.3", "v3.4"}:
        if paths.version == "v3.4":
            from lhexcel_v34_prepare import clean_package as cleaner
        else:
            from lhexcel_v33_prepare import clean_package as cleaner
        cleaner(paths.source_xlam, prepared_base, replace_custom_ui=True)
        catalog_source = prepared_base
    static_catalog_report = inject_static_catalog(catalog_source, staged_source, paths.source_dir)
    prepared_base.unlink(missing_ok=True)
    (paths.tmp_dir / "static-ribbon-catalog.json").write_text(
        json.dumps(static_catalog_report, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    stage_source_tree(paths.source_dir, paths.tmp_dir / "src")
    shutil.copy2(paths.builder, paths.tmp_dir / "lhexcel_dev.ps1")
    run_focus_probe = paths.version in {"v3.3", "v3.4"}
    run_shortcut_probe = paths.version == "v3.4"
    run_style_probe = paths.version == "v3.4"
    run_batch_rename_probe = paths.version == "v3.4"
    run_folder_create_probe = paths.version == "v3.4"
    run_batch_tools_ui_probe = paths.version == "v3.4"
    run_picture_fit_probe = paths.version == "v3.4"
    run_hangul_probe = paths.version == "v3.4"
    if run_focus_probe:
        shutil.copy2(SCRIPT_DIR / "lhexcel_focus_runtime_probe.ps1", paths.tmp_dir / "lhexcel_focus_runtime_probe.ps1")
        shutil.copy2(SCRIPT_DIR / "lhexcel_dialog_action.ps1", paths.tmp_dir / "lhexcel_dialog_action.ps1")
    if run_shortcut_probe:
        shutil.copy2(SCRIPT_DIR / "lhexcel_shortcut_runtime_probe.ps1", paths.tmp_dir / "lhexcel_shortcut_runtime_probe.ps1")
    if run_style_probe:
        shutil.copy2(SCRIPT_DIR / "lhexcel_style_runtime_probe.ps1", paths.tmp_dir / "lhexcel_style_runtime_probe.ps1")
    if run_batch_rename_probe:
        shutil.copy2(
            SCRIPT_DIR / "lhexcel_batch_rename_runtime_probe.ps1",
            paths.tmp_dir / "lhexcel_batch_rename_runtime_probe.ps1",
        )
    if run_folder_create_probe:
        shutil.copy2(
            SCRIPT_DIR / "lhexcel_folder_create_runtime_probe.ps1",
            paths.tmp_dir / "lhexcel_folder_create_runtime_probe.ps1",
        )
    if run_batch_tools_ui_probe:
        shutil.copy2(
            SCRIPT_DIR / "lhexcel_batch_tools_ui_runtime_probe.ps1",
            paths.tmp_dir / "lhexcel_batch_tools_ui_runtime_probe.ps1",
        )
    if run_picture_fit_probe:
        shutil.copy2(
            SCRIPT_DIR / "lhexcel_picture_fit_runtime_probe.ps1",
            paths.tmp_dir / "lhexcel_picture_fit_runtime_probe.ps1",
        )
    if run_hangul_probe:
        shutil.copy2(
            SCRIPT_DIR / "lhexcel_hangul_runtime_probe.ps1",
            paths.tmp_dir / "lhexcel_hangul_runtime_probe.ps1",
        )

    write_ascii(
        paths.tmp_dir / "build.ps1",
        build_ps1_text(
            run_focus_probe=run_focus_probe,
            run_shortcut_probe=run_shortcut_probe,
            run_style_probe=run_style_probe,
            run_batch_rename_probe=run_batch_rename_probe,
            run_folder_create_probe=run_folder_create_probe,
            run_batch_tools_ui_probe=run_batch_tools_ui_probe,
            run_picture_fit_probe=run_picture_fit_probe,
            run_hangul_probe=run_hangul_probe,
        ),
    )
    write_ascii(paths.tmp_dir / "run.bat", run_bat_text(paths.tmp_dir), crlf=True)
    write_ascii(paths.launcher_bat, launcher_bat_text(paths), crlf=True)

    info = {
        "version": paths.version,
        "token": paths.token,
        "tmp_dir": str(paths.tmp_dir),
        "windows_run_bat": to_windows_path(paths.tmp_dir / "run.bat"),
        "windows_launcher_bat": to_windows_path(paths.launcher_bat),
        "execution_contract": "BAT_ONLY_INSIDE_WINDOWS",
        "do_not_use": ["prlctl exec", "Mac .app launcher"],
        "source_xlam": str(paths.source_xlam),
        "staged_source_xlam": str(staged_source),
        "source_dir": str(paths.source_dir),
        "static_catalog_action_count": static_catalog_report["action_count"],
        "static_catalog_sha256": static_catalog_report["destination_sha256"],
        "staged_source_encoding": "cp949",
        "staged_cls_export_header": True,
        "automatic_focus_runtime_probe": run_focus_probe,
        "automatic_shortcut_runtime_probe": run_shortcut_probe,
        "automatic_style_runtime_probe": run_style_probe,
        "automatic_batch_rename_runtime_probe": run_batch_rename_probe,
        "automatic_folder_create_runtime_probe": run_folder_create_probe,
        "automatic_batch_tools_ui_runtime_probe": run_batch_tools_ui_probe,
        "automatic_picture_fit_runtime_probe": run_picture_fit_probe,
        "automatic_hangul_runtime_probe": run_hangul_probe,
        "candidate_xlam": str(paths.candidate_xlam),
        "final_xlam": str(paths.final_xlam),
    }
    (paths.tmp_dir / "runner-info.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(info, ensure_ascii=False, indent=2))


def read_latest_output(paths: VersionPaths) -> Path:
    latest = paths.tmp_dir / "latest-output.txt"
    if not latest.exists():
        raise SystemExit(f"latest-output.txt not found. Run Windows batch first: {to_windows_path(paths.tmp_dir / 'run.bat')}")
    raw = latest.read_text(encoding="utf-8-sig", errors="replace").strip()
    if not raw:
        raise SystemExit(f"latest-output.txt is empty: {latest}")
    candidate = paths.tmp_dir / "out" / Path(raw.replace("\\", "/")).name
    if not candidate.exists():
        raise SystemExit(f"built candidate not found: {candidate}")
    return candidate


def assert_success_logs(paths: VersionPaths) -> None:
    console_log = paths.tmp_dir / "windows-build-direct-run.log"
    dev_log = paths.tmp_dir / "lhexcel_dev.log"
    if not console_log.exists():
        raise SystemExit(f"console log not found: {console_log}")
    if not dev_log.exists():
        raise SystemExit(f"Excel build log not found: {dev_log}")
    console_text = console_log.read_text(encoding="utf-8-sig", errors="replace")
    dev_text = dev_log.read_text(encoding="utf-8-sig", errors="replace")
    required_console = "EXIT code=0"
    required_dev = ("VERIFIED vba_compile=pass", "SUCCESS Build")
    missing = [item for item in required_dev if item not in dev_text]
    if missing:
        raise SystemExit(f"Excel build log is missing success markers {missing}: {dev_log}")
    if required_console not in console_text:
        focus_log = paths.tmp_dir / "windows-focus-runtime.log"
        focus_ok = focus_log.exists() and "SUCCESS" in focus_log.read_text(encoding="utf-8-sig", errors="replace")
        latest_ok = (paths.tmp_dir / "latest-output.txt").exists()
        if not (focus_ok and latest_ok):
            raise SystemExit(f"Windows batch did not finish cleanly: {console_log}")
    if paths.version == "v3.4":
        shortcut_log = paths.tmp_dir / "windows-shortcut-runtime.log"
        shortcut_ok = shortcut_log.exists() and "SUCCESS" in shortcut_log.read_text(
            encoding="utf-8-sig", errors="replace"
        )
        if not shortcut_ok:
            raise SystemExit(f"Windows shortcut runtime probe did not pass: {shortcut_log}")
        style_log = paths.tmp_dir / "windows-style-runtime.log"
        style_ok = style_log.exists() and "SUCCESS" in style_log.read_text(
            encoding="utf-8-sig", errors="replace"
        )
        if not style_ok:
            raise SystemExit(f"Windows style runtime probe did not pass: {style_log}")
        batch_rename_log = paths.tmp_dir / "windows-batch-rename-runtime.log"
        batch_rename_ok = batch_rename_log.exists() and "SUCCESS" in batch_rename_log.read_text(
            encoding="utf-8-sig", errors="replace"
        )
        if not batch_rename_ok:
            raise SystemExit(f"Windows batch rename runtime probe did not pass: {batch_rename_log}")
        folder_create_log = paths.tmp_dir / "windows-folder-create-runtime.log"
        folder_create_ok = folder_create_log.exists() and "SUCCESS" in folder_create_log.read_text(
            encoding="utf-8-sig", errors="replace"
        )
        if not folder_create_ok:
            raise SystemExit(f"Windows folder create runtime probe did not pass: {folder_create_log}")
        batch_tools_ui_log = paths.tmp_dir / "windows-batch-tools-ui-runtime.log"
        batch_tools_ui_ok = batch_tools_ui_log.exists() and "SUCCESS" in batch_tools_ui_log.read_text(
            encoding="utf-8-sig", errors="replace"
        )
        if not batch_tools_ui_ok:
            raise SystemExit(f"Windows batch tools UI runtime probe did not pass: {batch_tools_ui_log}")
        picture_fit_log = paths.tmp_dir / "windows-picture-fit-runtime.log"
        picture_fit_text = (
            picture_fit_log.read_text(encoding="utf-8-sig", errors="replace")
            if picture_fit_log.exists()
            else ""
        )
        picture_fit_ok = (
            "PICTURE SUCCESS|" in picture_fit_text
            and "PICTURE_PROBE_EXIT=0" in picture_fit_text
            and "pictureFitRuntimeExit=0" in console_text
        )
        if not picture_fit_ok:
            raise SystemExit(f"Windows picture fit runtime probe did not pass: {picture_fit_log}")
        hangul_log = paths.tmp_dir / "windows-hangul-runtime.log"
        hangul_text = (
            hangul_log.read_text(encoding="utf-8-sig", errors="replace")
            if hangul_log.exists()
            else ""
        )
        hangul_ok = (
            "HANGUL SUCCESS|" in hangul_text
            and "HANGUL_PROBE_EXIT=0" in hangul_text
            and "hangulRuntimeExit=0" in console_text
        )
        if not hangul_ok:
            raise SystemExit(f"Windows Hangul runtime probe did not pass: {hangul_log}")


def run_checked(command: list[str], cwd: Path, stdout: Path | None = None) -> None:
    if stdout is None:
        subprocess.run(command, cwd=cwd, check=True)
        return
    with stdout.open("w", encoding="utf-8") as handle:
        subprocess.run(command, cwd=cwd, check=True, stdout=handle, stderr=subprocess.STDOUT)


def duplicate_report(xlam: Path, report: Path) -> None:
    with zipfile.ZipFile(xlam) as archive:
        names = archive.namelist()
    seen: set[str] = set()
    duplicates: list[str] = []
    for name in names:
        if name in seen and name not in duplicates:
            duplicates.append(name)
        seen.add(name)
    report.write_text("\n".join(duplicates) if duplicates else "duplicate_count=0\n", encoding="utf-8")


def promote_validated_pair(validated: Path, targets: tuple[Path, Path], *, tag: str) -> None:
    """Stage both destinations first and roll back both if either replace fails."""
    stages: dict[Path, Path] = {}
    backups: dict[Path, Path] = {}
    existed: dict[Path, bool] = {}
    expected_hash = file_sha256(validated)

    for target in targets:
        target.parent.mkdir(parents=True, exist_ok=True)
        stage = target.with_name(f".{target.name}.{tag}.new")
        backup = target.with_name(f".{target.name}.{tag}.bak")
        stage.unlink(missing_ok=True)
        backup.unlink(missing_ok=True)
        shutil.copy2(validated, stage)
        if file_sha256(stage) != expected_hash:
            raise SystemExit(f"staged promotion hash mismatch: {stage}")
        stages[target] = stage
        backups[target] = backup
        existed[target] = target.exists()

    for target in targets:
        if existed[target]:
            shutil.copy2(target, backups[target])

    try:
        for target in targets:
            stages[target].replace(target)
    except Exception:
        for target in targets:
            backup = backups[target]
            if backup.exists():
                backup.replace(target)
            elif not existed[target]:
                target.unlink(missing_ok=True)
        raise
    finally:
        for path in (*stages.values(), *backups.values()):
            path.unlink(missing_ok=True)


def promote(paths: VersionPaths, *, tag: str) -> None:
    built = read_latest_output(paths)
    assert_success_logs(paths)
    sync_generated_sources(paths, check=True)
    paths.reports_dir.mkdir(parents=True, exist_ok=True)
    tag = sanitize_tag(tag)

    console_report = paths.reports_dir / f"windows-build-{paths.token}-{tag}-console.log"
    build_report = paths.reports_dir / f"windows-build-{paths.token}-{tag}.log"
    shutil.copy2(paths.tmp_dir / "windows-build-direct-run.log", console_report)
    shutil.copy2(paths.tmp_dir / "lhexcel_dev.log", build_report)
    focus_runtime_log = paths.tmp_dir / "windows-focus-runtime.log"
    if focus_runtime_log.exists():
        shutil.copy2(focus_runtime_log, paths.reports_dir / f"windows-focus-runtime-{paths.token}-{tag}.log")
    shortcut_runtime_log = paths.tmp_dir / "windows-shortcut-runtime.log"
    if shortcut_runtime_log.exists():
        shutil.copy2(shortcut_runtime_log, paths.reports_dir / f"windows-shortcut-runtime-{paths.token}-{tag}.log")
    style_runtime_log = paths.tmp_dir / "windows-style-runtime.log"
    if style_runtime_log.exists():
        shutil.copy2(style_runtime_log, paths.reports_dir / f"windows-style-runtime-{paths.token}-{tag}.log")
    style_runtime_json = paths.tmp_dir / "windows-style-runtime.json"
    if style_runtime_json.exists():
        shutil.copy2(style_runtime_json, paths.reports_dir / f"windows-style-runtime-{paths.token}-{tag}.json")
    batch_rename_runtime_log = paths.tmp_dir / "windows-batch-rename-runtime.log"
    if batch_rename_runtime_log.exists():
        shutil.copy2(
            batch_rename_runtime_log,
            paths.reports_dir / f"windows-batch-rename-runtime-{paths.token}-{tag}.log",
        )
    folder_create_runtime_log = paths.tmp_dir / "windows-folder-create-runtime.log"
    if folder_create_runtime_log.exists():
        shutil.copy2(
            folder_create_runtime_log,
            paths.reports_dir / f"windows-folder-create-runtime-{paths.token}-{tag}.log",
        )
    batch_tools_ui_runtime_log = paths.tmp_dir / "windows-batch-tools-ui-runtime.log"
    if batch_tools_ui_runtime_log.exists():
        shutil.copy2(
            batch_tools_ui_runtime_log,
            paths.reports_dir / f"windows-batch-tools-ui-runtime-{paths.token}-{tag}.log",
        )
    picture_fit_runtime_log = paths.tmp_dir / "windows-picture-fit-runtime.log"
    if picture_fit_runtime_log.exists():
        shutil.copy2(
            picture_fit_runtime_log,
            paths.reports_dir / f"windows-picture-fit-runtime-{paths.token}-{tag}.log",
        )
    hangul_runtime_log = paths.tmp_dir / "windows-hangul-runtime.log"
    if hangul_runtime_log.exists():
        shutil.copy2(
            hangul_runtime_log,
            paths.reports_dir / f"windows-hangul-runtime-{paths.token}-{tag}.log",
        )

    validated = paths.tmp_dir / "validated-candidate.xlam"
    validated.unlink(missing_ok=True)
    shutil.copy2(built, validated)

    sanitizer_script = SCRIPT_DIR / f"lhexcel_{paths.token}_sanitize_package.py"
    if sanitizer_script.exists():
        run_checked(
            [
                "python3",
                str(sanitizer_script),
                "--xlam",
                str(validated),
                "--in-place",
                "--json",
                str(paths.reports_dir / f"package-sanitize-{paths.token}-{tag}-validated.json"),
            ],
            ROOT,
        )

    package_policy_script = SCRIPT_DIR / f"lhexcel_{paths.token}_package_policy.py"
    if package_policy_script.exists():
        run_checked(
            [
                "python3",
                str(package_policy_script),
                "--xlam",
                str(validated),
                "--json",
                str(paths.reports_dir / f"{paths.token}-package-policy-{tag}-validated.json"),
            ],
            ROOT,
        )

    contract_script = SCRIPT_DIR / f"lhexcel_{paths.token}_contract.py"
    if contract_script.exists():
        run_checked(
            ["python3", str(contract_script), "--source", str(paths.source_dir), "--xlam", str(validated)],
            ROOT,
            paths.reports_dir / f"contract-{paths.token}-{tag}-validated.json",
        )
    else:
        (paths.reports_dir / f"contract-{paths.token}-{tag}.skipped.txt").write_text(
            f"contract script not found: {contract_script}\n",
            encoding="utf-8",
        )

    if paths.version == "v3.4":
        run_checked(
            ["python3", str(SCRIPT_DIR / "lhexcel_shortcut_key_contract.py"), "--source", str(paths.source_dir / "mdShortcutManager.bas")],
            ROOT,
            paths.reports_dir / f"shortcut-key-contract-{paths.token}-{tag}.json",
        )
        run_checked(
            [
                "python3",
                str(SCRIPT_DIR / "lhexcel_v34_style_system_contract.py"),
                "--source",
                str(paths.source_dir),
                "--xlam",
                str(validated),
            ],
            ROOT,
            paths.reports_dir / f"style-system-contract-{paths.token}-{tag}-validated.json",
        )
        run_checked(
            [
                "python3",
                str(SCRIPT_DIR / "lhexcel_shortcut_icon_contract.py"),
                "--source",
                str(paths.source_dir),
                "--xlam",
                str(validated),
                "--allow-reset-migration-removed",
            ],
            ROOT,
            paths.reports_dir / f"shortcut-icon-contract-{paths.token}-{tag}-validated.json",
        )
        run_checked(
            [
                "python3",
                str(SCRIPT_DIR / "lhexcel_v34_hangul_export_contract.py"),
                "--source",
                str(paths.source_dir),
            ],
            ROOT,
            paths.reports_dir / f"hangul-export-contract-{paths.token}-{tag}-validated.json",
        )

    callback_contract_cmd = [
        "python3",
        str(SCRIPT_DIR / "lhexcel_contract_check.py"),
        "--xlam",
        str(validated),
        "--source",
        str(paths.source_dir),
    ]
    if not uses_strict_lh_case(paths.version):
        callback_contract_cmd.append("--allow-legacy-lh-case")
    run_checked(
        callback_contract_cmd,
        ROOT,
        paths.reports_dir / f"callback-contract-{paths.token}-{tag}.json",
    )
    feature_contract_script = SCRIPT_DIR / f"lhexcel_{paths.token}_feature_contract.py"
    if feature_contract_script.exists():
        run_checked(
            ["python3", str(feature_contract_script)],
            ROOT,
            paths.reports_dir / f"feature-contract-{paths.token}-{tag}.log",
        )
    run_checked(["unzip", "-t", str(validated)], ROOT, paths.reports_dir / f"unzip-{paths.token}-{tag}-validated.log")
    duplicate_path = paths.reports_dir / f"zip-duplicate-{paths.token}-{tag}-validated.txt"
    duplicate_report(validated, duplicate_path)
    if duplicate_path.read_text(encoding="utf-8").strip() != "duplicate_count=0":
        raise SystemExit(f"duplicate ZIP members found: {duplicate_path}")

    validated_hash = file_sha256(validated)
    promote_validated_pair(validated, (paths.candidate_xlam, paths.final_xlam), tag=tag)
    candidate_hash = file_sha256(paths.candidate_xlam)
    final_hash = file_sha256(paths.final_xlam)
    if candidate_hash != validated_hash or final_hash != validated_hash:
        raise SystemExit("candidate/final SHA-256 mismatch after promotion")

    result = {
        "version": paths.version,
        "built": str(built),
        "candidate_xlam": str(paths.candidate_xlam),
        "final_xlam": str(paths.final_xlam),
        "build_log": str(build_report),
        "console_log": str(console_report),
        "validated_xlam": str(validated),
        "validated_sha256": validated_hash,
        "candidate_sha256": candidate_hash,
        "final_sha256": final_hash,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


def status(paths: VersionPaths) -> None:
    files = [
        paths.tmp_dir / "run.bat",
        paths.tmp_dir / "build.ps1",
        paths.launcher_bat,
        paths.tmp_dir / "windows-build-direct-run.log",
        paths.tmp_dir / "lhexcel_dev.log",
        paths.tmp_dir / "latest-output.txt",
        paths.candidate_xlam,
        paths.final_xlam,
    ]
    for path in files:
        if path.exists():
            print(f"OK {path.stat().st_mtime:.0f} {path.stat().st_size:>10} {path}")
        else:
            print(f"-- {'':>10} {path}")


def prepare_v001_refinement(paths: V001ReleasePaths, tag: str, suite: str = "all", cases: str = "") -> None:
    """Generate a BAT for changed workflows against the exact built XLAM."""
    if not re.fullmatch(r"[a-z0-9-]+", tag):
        raise SystemExit("refinement tag must be lowercase ASCII")
    if cases and (suite not in {"refinement", "workflow"} or not re.fullmatch(r"[A-Za-z0-9_.,]+", cases)):
        raise SystemExit("selected cases require refinement/workflow and ASCII case IDs")
    product = paths.tmp_dir / "out/complete/Product.xlam"
    source = paths.tmp_dir / "source"
    if not product.is_file():
        raise SystemExit("refinement requires an existing release Product.xlam")
    suites = {"convergence": ("r58-convergence", "Run-R58ConvergenceSuite.ps1"),
              "refinement": ("r61-refinement", "Run-R61RefinementSuite.ps1"),
              "workflow": ("r62-workflow", "Run-R62WorkflowSuite.ps1"),
              "compare-navigation": ("r62-compare-navigation", "Run-R62CompareNavigationSuite.ps1")}
    defaults = ("convergence", "refinement", "workflow", "compare-navigation") if int(paths.version.rsplit("_r", 1)[1]) >= 62 else ("convergence", "refinement")
    selected = [suites[key] for key in defaults] if suite == "all" else [suites[suite]]
    for _, name in selected:
        if not (source / "tests/windows" / name).is_file():
            raise SystemExit("built source is missing " + name)
    for name, _ in selected:
        if (paths.tmp_dir / ("evidence-" + tag) / name).exists():
            raise SystemExit("existing refinement evidence is preserved: " + name)
    ps_path = paths.tmp_dir / ("refinement-" + tag + ".ps1")
    bat_path = paths.tmp_dir / ("refinement-" + tag + ".bat")
    if ps_path.exists() or bat_path.exists():
        raise SystemExit("existing refinement runner is preserved")
    script = r'''Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close user Excel before refinement checks' }
$source = Join-Path $PSScriptRoot 'source'
$product = Join-Path $PSScriptRoot 'out/complete/Product.xlam'
foreach ($case in @(__CASES__)) {
    $evidence = Join-Path $PSScriptRoot ('evidence-__TAG__/' + $case[0])
    & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source ('tests/windows/' + $case[1])) -ProductXlam $product -EvidenceRoot $evidence
    if ($LASTEXITCODE -ne 0) { throw ('Refinement native suite failed: ' + $case[0]) }
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Excel remained after refinement suite' }
}
'''
    case_rows = ", ".join("@('" + name + "','" + runner + "')" for name, runner in selected)
    # Unary comma retains the single case tuple when PowerShell enumerates it.
    if len(selected) == 1:
        case_rows = "," + case_rows
    script = script.replace("__CASES__", case_rows).replace("__TAG__", tag)
    if cases:
        script = script.replace("-EvidenceRoot $evidence", "-EvidenceRoot $evidence -CaseNames '" + cases + "'")
    write_ascii(ps_path, script)
    write_ascii(bat_path, '@echo off\nsetlocal\ncd /d "%~dp0"\npowershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0' + ps_path.name + '"\nexit /b %ERRORLEVEL%\n', crlf=True)
    print(json.dumps({"bat": str(bat_path), "artifact_sha256": hashlib.sha256(product.read_bytes()).hexdigest(), "execution_contract": "BAT_ONLY_INSIDE_WINDOWS"}))


def prepare_document_navigator(tag: str, visual_hold: bool = False, artifact: Path | None = None) -> None:
    """Build a fresh no-DLL diagnostic candidate; never publish or replace releases."""
    if not re.fullmatch(r"[a-z0-9-]+", tag):
        raise SystemExit("navigator tag must be lowercase ASCII")
    directory = TMP_ROOT / ("r61-document-navigator-" + tag)
    directory.mkdir(parents=True, exist_ok=False)
    source = ROOT / "source/v0.01"
    shutil.copytree(source, directory / "source",
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".pytest_cache", "out", "*.dll", "*.xlam"))
    if artifact is not None:
        if not artifact.is_file() or artifact.suffix.lower() != '.xlam':
            raise SystemExit('existing diagnostic XLAM is required')
        (directory / 'out').mkdir()
        shutil.copy2(artifact, directory / 'out/Product.xlam')
    script = r'''Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close user Excel before navigator checks' }
$source = Join-Path $PSScriptRoot 'source'
$product = Join-Path $PSScriptRoot 'out/Product.xlam'
$previousProfile = $env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT = Join-Path $PSScriptRoot 'profile-build'
try {
    & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source 'build/Build-Xlam.ps1') -DataRoot $source -OutputPath $product -ManifestPath (Join-Path $source 'build/manifests/Product.json') -RibbonPath (Join-Path $source 'src/ribbon/customUI14.xml') *> (Join-Path $PSScriptRoot 'build.log')
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $product -PathType Leaf)) { throw 'Navigator candidate build failed; see build.log' }
    & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $source 'tests/windows/Run-R60DocumentNavigatorSuite.ps1') -ProductXlam $product -EvidenceRoot (Join-Path $PSScriptRoot 'evidence')__VISUAL__
    if ($LASTEXITCODE -ne 0) { throw 'Navigator native suite failed; see evidence' }
} finally { $env:LHEXCEL_PROFILE_ROOT = $previousProfile }
    Write-Output 'PASS|R60_DOCUMENT_NAVIGATOR|candidate only; no release promotion'
'''.replace('__VISUAL__', ' -VisualHold' if visual_hold else '')
    if artifact is not None:
        start = script.index('    & powershell.exe')
        end = script.index("    & powershell.exe", start + 1)
        script = script[:start] + script[end:]
    write_ascii(directory / 'run.ps1', script)
    write_ascii(directory / 'run.bat', '@echo off\nsetlocal\ncd /d "%~dp0"\npowershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0run.ps1"\nexit /b %ERRORLEVEL%\n', crlf=True)
    info = {"bat": str(directory / 'run.bat'), "source_manifest_sha256": hashlib.sha256((source / 'build/manifests/Product.json').read_bytes()).hexdigest(),
            "candidate_only": True, "dll_included": False, "security_settings_changed": False,
            "reused_artifact": str(artifact) if artifact is not None else None,
            "reused_artifact_sha256": hashlib.sha256(artifact.read_bytes()).hexdigest() if artifact is not None else None}
    (directory / 'runner-info.json').write_text(json.dumps(info, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(info))


def prepare_ctp_diagnostics(tag: str) -> None:
    if not re.fullmatch(r"[a-z0-9-]+", tag):
        raise SystemExit("check tag must be lowercase ASCII")
    directory = TMP_ROOT / ("r60-ctp-diagnostics-" + tag)
    directory.mkdir(parents=True, exist_ok=False)
    source = ROOT / "source/v0.01/tests/windows"
    for name in ("Test-R60CtpDiagnostics.ps1", "Run-R60CtpCreation.ps1", "Run-R60RegisteredInventory.ps1"):
        shutil.copy2(source / name, directory / name)
    (directory / "run.bat").write_text('@echo off\nsetlocal\ncd /d "%~dp0"\npowershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-R60CtpDiagnostics.ps1" -EvidenceRoot "%~dp0evidence"\nexit /b %ERRORLEVEL%\n', encoding="ascii", newline="\r\n")
    print(json.dumps({"bat": str(directory / "run.bat"), "excel_or_registry_execution": False}))


def prepare_internal_checks(tag: str, artifact: Path) -> None:
    import ast
    from datetime import date
    if not re.fullmatch(r"[a-z0-9-]+", tag):
        raise SystemExit("check tag must be lowercase ASCII")
    directory = TMP_ROOT / ("r61-internal-checks-" + tag)
    directory.mkdir(parents=True, exist_ok=False)
    source = ROOT / "source" / "v0.01"
    script = (source / "tools/build_r47_distribution.py").read_text(encoding="utf-8")
    parsed = ast.parse(script)
    function = next(node for node in parsed.body if isinstance(node, ast.FunctionDef) and node.name == "distribution_policy_module")
    namespace = {"date": date, "VERSION_LABEL": "v0.01_r61"}
    exec(compile(ast.Module(body=[function], type_ignores=[]), "policy fixture", "exec"), namespace)
    policy = namespace["distribution_policy_module"](date(2027, 6, 30), date(2027, 6, 29))
    policy = re.sub(r'^Attribute VB_Name = .*\n', '', policy)
    policy = re.sub(r'\bDate\b', 'PolicyTestDate()', policy)
    policy, count = re.subn(r'    MsgBox .*?vbExclamation, "[^"\n]+"', '    PolicyTestNotice "expiry"', policy, count=1, flags=re.S)
    if count != 1:
        raise SystemExit("expiry notice fixture mismatch")
    birthday = '    MsgBox LHE_BIRTHDAY_MESSAGE, vbInformation, "내엑셀"'
    if policy.count(birthday) != 1:
        raise SystemExit("birthday notice fixture mismatch")
    policy = policy.replace(birthday, '    PolicyTestNotice "birthday"')
    (directory / "policy.bas").write_text(policy, encoding="utf-8")
    for item in (source / "tests/vba/product/R60PolicyProbe.bas", source / "tests/windows/Run-R60InternalChecks.ps1", source / "build/Excel-ProcessLifecycle.ps1"):
        shutil.copy2(item, directory / item.name)
    shutil.copy2(artifact, directory / "Internal.xlam")
    (directory / "run.bat").write_text('@echo off\nsetlocal\ncd /d "%~dp0"\npowershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Run-R60InternalChecks.ps1" -EvidenceRoot "%~dp0evidence"\nexit /b %ERRORLEVEL%\n', encoding="ascii", newline="\r\n")
    print(json.dumps({"bat": str(directory / "run.bat"), "system_date_changed": False}))


def prepare_com_inventory(tag: str, candidate: Path | None = None, build: Path | None = None) -> None:
    if not re.fullmatch(r"[a-z0-9-]+", tag):
        raise SystemExit("inventory tag must be lowercase ASCII")
    directory = TMP_ROOT / ("r60-com-inventory-" + tag)
    directory.mkdir(parents=True, exist_ok=False)
    source = ROOT / "source" / "v0.01"
    for item in (source / "tests/windows/Run-R60ComInventory.ps1", source / "build/Excel-ProcessLifecycle.ps1"):
        shutil.copy2(item, directory / item.name)
    runner = "Run-R60ComInventory.ps1"
    if candidate is not None:
        if build is None:
            raise SystemExit("registered inventory requires --probe-build")
        runner = "Run-R60RegisteredInventory.ps1"
        shutil.copy2(source / "tests/windows" / runner, directory / runner)
        shutil.copy2(source / "tests/windows/Run-R60CtpCreation.ps1", directory)
        shutil.copy2(source / "tests/windows/Run-R60PaneActivation.ps1", directory)
        registration = candidate / "tests/windows/Register-DocumentNavigatorCtpProbe.ps1"
        shutil.copy2(registration, directory / registration.name)
        target = directory / "build/x64"
        target.mkdir(parents=True)
        shutil.copy2(build / "x64/DocumentNavigatorCtpProbe64.dll", target)
    bat = directory / "run.bat"
    bat.write_text('@echo off\nsetlocal\ncd /d "%~dp0"\n'
                   f'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0{runner}" -EvidenceRoot "%~dp0evidence"\n'
                   'exit /b %ERRORLEVEL%\n', encoding="ascii", newline="\r\n")
    print(json.dumps({"bat": str(bat), "registry_writes": candidate is not None}))


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    setup_parser = sub.add_parser("setup-acceptance", help="Generate native installed acceptance BAT")
    setup_parser.add_argument("--setup-exe", type=Path, required=True)
    setup_parser.add_argument("--previous-setup-exe", type=Path)
    setup_parser.add_argument("--version", default="v0.01_r65")
    setup_parser.add_argument("--previous-version", default="v0.01_r64")
    setup_parser.add_argument("--tag", required=True)
    setup_parser.add_argument("--already-installed", action="store_true")
    setup_parser.add_argument("--keep-installed", action="store_true")
    setup_parser.add_argument("--interactive-ui", action="store_true")
    navigator_parser = sub.add_parser("document-navigator", help="Build and test an isolated no-DLL navigator candidate")
    navigator_parser.add_argument("--tag", required=True)
    navigator_parser.add_argument("--visual-hold", action="store_true")
    navigator_parser.add_argument("--artifact", type=Path, help="Retest an unchanged existing diagnostic build")
    diagnostics_parser = sub.add_parser("ctp-diagnostics", help="Generate isolated PowerShell diagnostic regression BAT")
    diagnostics_parser.add_argument("--tag", required=True)
    internal_parser = sub.add_parser("internal-checks", help="Generate scoped no-DLL preview and policy BAT")
    internal_parser.add_argument("--tag", required=True)
    internal_parser.add_argument("--artifact", type=Path, required=True)
    inventory_parser = sub.add_parser("com-inventory", help="Generate read-only native COM discovery BAT")
    inventory_parser.add_argument("--version", choices=("v0.01_r60",), required=True)
    inventory_parser.add_argument("--tag", required=True)
    inventory_parser.add_argument("--probe-candidate", type=Path)
    inventory_parser.add_argument("--probe-build", type=Path)

    prepare_parser = sub.add_parser("prepare", help="Create ASCII Windows runner")
    prepare_parser.add_argument("--version", required=True)
    prepare_parser.add_argument("--no-clean", action="store_true", help="Do not remove an existing temp runner first")
    prepare_parser.add_argument("--attempt", default="", help="Preserve earlier evidence in a separate named attempt")
    prepare_parser.add_argument("--suite", choices=("all", "manner", "dll-hwpx", "r68-product"), default="all")

    verify_parser = sub.add_parser("verify", help="Create a BAT for changed-path verification of an existing v0.01 build")
    verify_parser.add_argument("--version", required=True)
    verify_parser.add_argument("--tag", required=True)
    verify_parser.add_argument("--suite", choices=("all", "manner", "ui", "focus-copy", "focus-speed", "data-compare", "dll-hwpx", "dll-visual", "dll-navigator", "compare-benchmark", "compare-reopen", "frame-draw", "frame-product", "frame-product-all", "compare-results", "picture-state", "r68-product", "hangul-transfer"), default="all")
    verify_parser.add_argument("--cases", default="")
    verify_parser.add_argument("--attempt", default="")
    verify_parser.add_argument("--artifact", type=Path)
    verify_parser.add_argument("--host-root", type=Path)

    refinement_parser = sub.add_parser("refinement", help="Generate BAT-only changed workflow checks")
    refinement_parser.add_argument("--version", required=True)
    refinement_parser.add_argument("--attempt", default="")
    refinement_parser.add_argument("--tag", required=True)
    refinement_parser.add_argument("--suite", choices=("all", "convergence", "refinement", "workflow", "compare-navigation"), default="all")
    refinement_parser.add_argument("--cases", default="")

    distribution_parser = sub.add_parser("distribution-smoke", help="Generate BAT-only protected-distribution checks")
    distribution_parser.add_argument("--version", required=True)
    distribution_parser.add_argument("--attempt", default="")
    distribution_parser.add_argument("--tag", required=True)
    distribution_parser.add_argument("--internal", type=Path, required=True)
    distribution_parser.add_argument("--enhanced", type=Path, required=True)
    distribution_parser.add_argument("--defer-artifact-check", action="store_true", help="Prepare the BAT before protected staging; execution still requires both XLAM files")
    distribution_parser.add_argument("--expected-policy", choices=("active", "expired", "invalid"), default="active")
    distribution_parser.add_argument("--skip-no-dll-focus", action="store_true", help="Keep focus validation NOT_RUN when only changed non-focus paths are in scope")

    promote_parser = sub.add_parser("promote", help="Promote successful Windows runner output")
    promote_parser.add_argument("--version", required=True)
    promote_parser.add_argument("--tag", default="direct-run-final")

    status_parser = sub.add_parser("status", help="Show runner and output status")
    status_parser.add_argument("--version", required=True)

    args = parser.parse_args(argv)
    if args.command == "setup-acceptance":
        if not all(re.fullmatch(r"v0\.01_r[0-9]+", value) for value in (args.version, args.previous_version)):
            parser.error("invalid setup version")
        if not re.fullmatch(r"[a-z0-9-]+", args.tag):
            parser.error("invalid tag")
        directory = TMP_ROOT / ("r63-setup-" + args.tag)
        directory.mkdir(parents=True, exist_ok=False)
        setup = directory / "Setup.exe"
        shutil.copy2(args.setup_exe, setup)
        source = ROOT / "source/v0.01"
        for relative in ("tests/windows/Run-R63SetupSuite.ps1", "tests/windows/Test-R75InstalledIntegrity.ps1", "tests/windows/Run-NxSetupUpgradeSuite.ps1", "tests/windows/Run-R97InstallerUi.ps1", "build/Excel-ProcessLifecycle.ps1"):
            destination = directory / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source / relative, destination)
        suite = "Run-R63SetupSuite.ps1"
        extra = f' -ExpectedVersion "{args.version}"'
        if args.interactive_ui:
            if int(args.version.rsplit("_r", 1)[1]) < 97 or args.already_installed or args.keep_installed or args.previous_setup_exe:
                parser.error("interactive-ui requires r97 or later and manages its own lifecycle")
            suite = "Run-R97InstallerUi.ps1"
            extra = f' -ExpectedVersion "{args.version}"'
        if args.already_installed:
            extra += " -AlreadyInstalled"
        if args.keep_installed:
            extra += " -KeepInstalled"
        if args.previous_setup_exe:
            shutil.copy2(args.previous_setup_exe, directory / "PreviousSetup.exe")
            suite = "Run-NxSetupUpgradeSuite.ps1"
            extra += f' -PreviousSetupExe "%~dp0PreviousSetup.exe" -PreviousVersion "{args.previous_version}"'
        write_ascii(directory / "run.bat", '@echo off\nsetlocal\npushd "%~dp0" || exit /b 1\npowershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0tests/windows/' + suite + '" -SetupExe "%~dp0Setup.exe" -EvidenceRoot "%~dp0evidence"' + extra + '\nset "NX_SETUP_EXIT=%ERRORLEVEL%"\npopd\nexit /b %NX_SETUP_EXIT%\n', crlf=True)
        print(json.dumps({"bat": str(directory / "run.bat"), "setup_sha256": file_sha256(setup)}))
        return 0
    if args.command == "document-navigator":
        prepare_document_navigator(args.tag, args.visual_hold, args.artifact)
        return 0
    if args.command == "ctp-diagnostics":
        prepare_ctp_diagnostics(args.tag)
        return 0
    if args.command == "internal-checks":
        prepare_internal_checks(args.tag, args.artifact)
        return 0
    if args.command == "com-inventory":
        prepare_com_inventory(args.tag, args.probe_candidate, args.probe_build)
        return 0
    normalized = normalize_version(args.version)
    if is_v001_release(normalized):
        paths = resolve_v001_release_paths(normalized)
        attempt = getattr(args, "attempt", "")
        if attempt:
            if not re.fullmatch(r"[a-z0-9-]+", attempt):
                parser.error("attempt must be lowercase ASCII")
            paths = replace(paths, tmp_dir=paths.tmp_dir.with_name(paths.tmp_dir.name + "_" + attempt),
                            launcher_bat=paths.launcher_bat.with_name(paths.launcher_bat.stem + "_" + attempt + ".bat"))
        if args.command == "prepare":
            prepare_v001_release(paths, clean=not args.no_clean, suite=args.suite)
        elif args.command == "status":
            status_v001_release(paths)
        elif args.command == "verify":
            prepare_v001_verification(paths, args.tag, args.suite, args.cases, args.artifact, args.host_root)
        elif args.command == "refinement":
            prepare_v001_refinement(paths, args.tag, args.suite, args.cases)
        elif args.command == "distribution-smoke":
            prepare_v001_distribution_smoke(paths, args.tag, args.internal, args.enhanced, args.defer_artifact_check, args.expected_policy, args.skip_no_dll_focus)
        else:
            parser.error("v0.01 release runners publish through the verified release pipeline; promote is unsupported")
        return 0

    paths = resolve_paths(normalized)
    if args.command == "prepare":
        prepare(paths, clean=not args.no_clean)
    elif args.command == "promote":
        promote(paths, tag=args.tag)
    elif args.command == "status":
        status(paths)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
