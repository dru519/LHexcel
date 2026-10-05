from __future__ import annotations

"""Create the static-only P01 Windows verifier; Windows alone emits return ZIPs.

The UTF-8-BOM Korean guide and copied Run-VbaSuite invoke Build-TestXlam.ps1.
"""

import argparse
import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import tempfile
import uuid
from datetime import UTC, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PROJECT_ROOT = ROOT.parents[1]
PARALLEL_ROOT = PROJECT_ROOT / "tmp" / "LHexcel-v0.01-parallel"
APPROVED_COMMIT = "5d852e61697cd224c4c575380e2e827927b97e69"
EXIT_CODES = "0, 10, 20, 21, 22, 23, 24, 25, 26"
STAGED_DATA_PATTERN = r"^payload/(src/|tests/vba/|tests/manifests/|contracts/|tools/|provenance/)"
ALLOWED_RUNTIME_SCRIPT = "payload/tools/lhexcel_hwpx_table_export.ps1"
CORE = (
    "NxTypes.bas", "NxConstants.bas", "NxFormulaInterop.bas", "NxUserErrors.bas", "INxFeatureCommand.cls", "CNxExecutionContext.cls",
    "CNxPlan.cls", "CNxResult.cls", "CNxStateGuard.cls", "CNxSideEffect.cls",
    "CNxCellRollbackJournal.cls", "NxContextFactory.bas", "NxCommandRunner.bas",
    "CNxApproval.cls", "CNxSideEffectDispatcher.cls", "CNxLimitsPolicy.cls",
    "CNxRangeSelectionSession.cls", "NxRangePicker.bas",
)
TESTS = ("NxTestHarness.bas", "T_Core.bas", "CFakeFeatureCommand.cls")
SUITES = {
    "Core": {"launcher": "RUN_CORE_VERIFY.bat", "expected": 36, "extension_manifest": None},
    "Frame": {"launcher": "RUN_FRAME_VERIFY.bat", "expected": 28, "extension_manifest": "tests/manifests/Frame.json"},
    "AI": {"launcher": "RUN_AI_VERIFY.bat", "expected": 29, "extension_manifest": "tests/manifests/AI.json"},
    "G005": {"launcher": "RUN_G005_VERIFY.bat", "expected": 10, "extension_manifest": "tests/manifests/G005.json", "regressions": (("Core", 36), ("Frame", 28), ("AI", 29))},
    "Template": {"launcher": "RUN_TEMPLATE_VERIFY.bat", "expected": 30, "extension_manifest": "tests/manifests/Template.json", "regressions": (("Core", 36), ("Frame", 28), ("AI", 29), ("G005", 10))},
    "Data": {"launcher": "RUN_DATA_VERIFY.bat", "expected": 17, "extension_manifest": "tests/manifests/Data.json", "regressions": (("Core", 36), ("Frame", 28), ("AI", 29), ("G005", 10), ("Template", 30))},
    "Draw": {"launcher": "RUN_DRAW_VERIFY.bat", "expected": 10, "extension_manifest": "tests/manifests/Draw.json", "regressions": (("Core", 36), ("Frame", 28), ("AI", 29), ("G005", 10), ("Template", 30), ("Data", 17))},
    "File": {"launcher": "RUN_FILE_VERIFY.bat", "expected": 14, "extension_manifest": "tests/manifests/File.json", "regressions": (("Core", 36), ("Frame", 28), ("AI", 29), ("G005", 10), ("Template", 30), ("Data", 17), ("Draw", 10))},
    "Calculator": {"launcher": "RUN_CALCULATOR_VERIFY.bat", "expected": 28, "extension_manifest": "tests/manifests/Calculator.json", "regressions": (("Core", 36), ("Frame", 28), ("AI", 29), ("G005", 10), ("Template", 30), ("Data", 17), ("Draw", 10), ("File", 14))},
    "Symbols": {"launcher": "RUN_SYMBOLS_VERIFY.bat", "expected": 24, "extension_manifest": "tests/manifests/Symbols.json", "regressions": (("Core", 36), ("Frame", 28), ("AI", 29), ("G005", 10), ("Template", 30), ("Data", 17), ("Draw", 10), ("File", 14), ("Calculator", 28))},
    # ProductRibbon is intentionally a single, independent lane.  Its PowerShell
    # entrypoint builds the real XLAM before the callback/UI evidence pass.
    "ProductRibbon": {"launcher": "RUN_PRODUCT_RIBBON_VERIFY.bat", "expected": 3, "extension_manifest": None, "product": True},
    "ProductUi": {"launcher": "RUN_PRODUCT_UI_VERIFY.bat", "expected": 1, "extension_manifest": None, "product": True},
}


def git_command(cwd: Path | str, *args: str) -> list[str]:
    """Build a command-local safe.directory invocation for shared folders."""
    cursor = Path(cwd).resolve()
    repo_hint = next((path for path in (cursor, *cursor.parents) if (path / ".git").exists()), cursor)
    return ["git", "-c", f"safe.directory={repo_hint.as_posix()}", "-C", str(cwd), *args]


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write(path: Path, text: str, bom: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes((b"\xef\xbb\xbf" if bom else b"") + text.encode("utf-8"))


def digest_pairs(pairs: list[tuple[str, str]]) -> str:
    return hashlib.sha256("".join(f"{name}\n{value}\n" for name, value in pairs).encode()).hexdigest()


def product_source_tree_digest() -> str:
    rows: list[tuple[str, str]] = []
    for directory in ("src", "contracts", "build", "tools", "provenance"):
        for path in (ROOT / directory).rglob("*"):
            if not path.is_file() or "__pycache__" in path.parts or path.suffix == ".pyc":
                continue
            relative = path.relative_to(ROOT).as_posix()
            if relative == "provenance/origins.json":
                continue
            rows.append((relative, sha(path)))
    canonical = "\n".join(f"{name}\t{digest}" for name, digest in sorted(rows)).encode()
    return hashlib.sha256(canonical).hexdigest()


def require_clean_tracked_source() -> str:
    """Return the real source commit only when the Product source is frozen.

    The working source tree is intentionally outside the main Codex Git
    checkout in some test fixtures.  In that case Git cannot prove identity,
    so Product packaging fails closed instead of recording a guessed commit.
    """
    try:
        repo = subprocess.check_output(
            git_command(ROOT, "rev-parse", "--show-toplevel"),
            text=True,
            stderr=subprocess.STDOUT,
        ).strip()
        source_paths = [
            path for directory in ("src", "contracts", "build", "tools", "provenance")
            for path in (ROOT / directory).rglob("*")
            if path.is_file() and path.suffix != ".pyc" and "__pycache__" not in path.parts
        ]
        if not source_paths:
            raise ValueError("Product source tree is empty")
        for path in source_paths:
            relative = str(path.relative_to(repo)) if path.is_relative_to(Path(repo)) else str(path)
            tracked = subprocess.run(
                git_command(repo, "ls-files", "--error-unmatch", "--", relative),
                text=True,
                capture_output=True,
                check=False,
            )
            if tracked.returncode != 0:
                raise ValueError(f"Product source path is not tracked: {relative}")
        status = subprocess.run(
            git_command(repo, "status", "--porcelain=v1", "--untracked-files=all", "--", str(ROOT)),
            text=True,
            capture_output=True,
            check=False,
        )
        if status.returncode != 0 or status.stdout.strip():
            raise ValueError("Product source tree is not clean")
        return subprocess.check_output(git_command(repo, "rev-parse", "HEAD"), text=True).strip()
    except (OSError, subprocess.CalledProcessError) as exc:
        raise ValueError(f"Product source Git preflight unavailable: {exc}") from exc


def validate_extension_graph(payload: Path, manifest_name: str, suite: str, owner: str) -> None:
    visiting: set[str] = set()
    loaded: set[str] = set()
    modules: set[str] = set()
    dependency_count = 0

    def visit(relative: str, expected_suite: str, expected_owner: str, depth: int) -> None:
        nonlocal dependency_count
        if depth > 8:
            raise ValueError("extension dependency depth exceeded")
        if not re.fullmatch(r"tests/manifests/[A-Za-z][A-Za-z0-9_-]*\.json", relative):
            raise ValueError("extension dependency path rejected")
        key = relative.lower()
        if key in visiting:
            raise ValueError("extension dependency cycle rejected")
        if key in loaded:
            raise ValueError("duplicate extension dependency rejected")
        path = payload / relative
        if not path.is_file():
            raise ValueError("extension dependency is missing")
        document = json.loads(path.read_text(encoding="utf-8"))
        if set(document) != {
            "schema_version", "suite", "owner_phase", "base_allowlist_digest", "dependencies", "modules",
        }:
            raise ValueError("extension manifest property set rejected")
        if type(document["schema_version"]) is not int or document["schema_version"] != 2:
            raise ValueError("extension manifest schema rejected")
        if document["suite"] != expected_suite or document["owner_phase"] != expected_owner:
            raise ValueError("extension manifest suite or owner rejected")
        visiting.add(key)
        try:
            for dependency in document["dependencies"]:
                if set(dependency) != {"path", "sha256", "suite", "owner_phase"}:
                    raise ValueError("extension dependency property set rejected")
                dependency_count += 1
                if dependency_count > 8:
                    raise ValueError("extension dependency count exceeded")
                dependency_path = payload / dependency["path"]
                if not dependency_path.is_file() or sha(dependency_path) != dependency["sha256"]:
                    raise ValueError("extension dependency digest rejected")
                visit(dependency["path"], dependency["suite"], dependency["owner_phase"], depth + 1)
            for module in document["modules"]:
                if set(module) != {"path", "sha256", "owner_phase"} or module["owner_phase"] != expected_owner:
                    raise ValueError("extension module property set or owner rejected")
                module_key = module["path"].lower()
                if module_key in modules:
                    raise ValueError("duplicate extension module rejected")
                module_path = payload / module["path"]
                if not module_path.is_file() or sha(module_path) != module["sha256"]:
                    raise ValueError("extension module digest rejected")
                modules.add(module_key)
            if not document["modules"]:
                raise ValueError("extension manifest requires modules")
            loaded.add(key)
        finally:
            visiting.remove(key)

    visit(manifest_name, suite, owner, 0)


def commit() -> str:
    return subprocess.check_output(git_command(ROOT, "rev-parse", "HEAD"), text=True).strip()


def make_read_only(path: Path) -> None:
    """Freeze the macOS source snapshot while retaining directory traversal."""
    for item in sorted(path.rglob("*"), reverse=True):
        mode = item.stat().st_mode
        if item.is_dir():
            item.chmod((mode & ~0o222) | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)
        else:
            item.chmod(mode & ~0o222)
    mode = path.stat().st_mode
    path.chmod((mode & ~0o222) | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)


def clear_hidden_flags(path: Path) -> None:
    """Do not leak macOS hidden metadata into Windows shared-folder attributes."""
    hidden_flag = getattr(stat, "UF_HIDDEN", 0)
    if hidden_flag == 0 or not hasattr(os, "chflags"):
        return
    for item in (path, *path.rglob("*")):
        flags = getattr(item.stat(follow_symlinks=False), "st_flags", 0)
        if flags & hidden_flag:
            os.chflags(item, flags & ~hidden_flag, follow_symlinks=False)


def thaw_for_removal(path: Path) -> None:
    """Only the generator thaws a previous local bundle before replacing it."""
    for item in sorted(path.rglob("*"), reverse=True):
        item.chmod(item.stat().st_mode | (0o700 if item.is_dir() else 0o600))
    path.chmod(path.stat().st_mode | 0o700)


def suite_regression_verification(suite: str) -> str:
    regressions = SUITES[suite].get("regressions", ())
    if not regressions:
        return ""
    rows = []
    for name, expected in regressions:
        manifest = SUITES[name]["extension_manifest"]
        manifest_literal = "$null" if manifest is None else f"'{manifest}'"
        user_form = "$true" if name in {"AI", "Template", "Data", "Draw", "File", "Calculator", "Symbols"} else "$false"
        rows.append(
            f"      [pscustomobject]@{{Name='{name}';Expected={expected};Manifest={manifest_literal};"
            f"Owner='{name.lower()}';UserForm={user_form}}}"
        )
    return r'''
    $regressionPlan=@(
__REGRESSION_ROWS__
    )
    foreach($regression in $regressionPlan){
      $name=[string]$regression.Name;$regOut=Join-Path $scratch ($name+'.out');$regErr=Join-Path $scratch ($name+'.err')
      $regArguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',$suite,'-Suite',$name,'-Mode','Green','-RunId',$Manifest['run_id'],'-SourceDigest',$Manifest['base17_digest'],'-SnapshotDigest',$Manifest['snapshot_digest'],'-DataRoot',$dataRoot,'-EvidenceRoot',$evidence)
      if($null -ne $regression.Manifest){$regArguments+=@('-ExtensionManifest',(Join-Path $dataRoot ([string]$regression.Manifest)))}
      $script:phase=('regression-'+$name.ToLowerInvariant())
      $regCode=Invoke-CapturedPowerShell -Arguments $regArguments -StdOut $regOut -StdErr $regErr
      $script:phase='integrity';Assert-SourceBundleUnchanged $Manifest $sourceBundle;Assert-StagedDataExact $Manifest $LocalBundle;$script:phase=('regression-'+$name.ToLowerInvariant())
      $regEvidencePath=Join-Path $evidence ($name+'.json')
      if(Test-Path -LiteralPath $regEvidencePath -PathType Leaf){Copy-Item -LiteralPath $regEvidencePath -Destination (Join-Path $returnRoot ('results/'+$name+'.json')) -Force}
      if(Test-Path -LiteralPath $regOut -PathType Leaf){Write-Utf8NoBom (Join-Path $returnRoot ('logs/'+$name+'.stdout.redacted.txt')) (Redact ([IO.File]::ReadAllText($regOut)))}
      if(Test-Path -LiteralPath $regErr -PathType Leaf){Write-Utf8NoBom (Join-Path $returnRoot ('logs/'+$name+'.stderr.redacted.txt')) (Redact ([IO.File]::ReadAllText($regErr)))}
      if($regCode -ne 0){$mappedCode=if($regCode -in @(10,22,23,24,25)){$regCode}else{22};Stop-Code $mappedCode ($name+' regression failed')}
      $regEvidence=Require-RegressionEvidence $regEvidencePath $Manifest $name ([int]$regression.Expected) ([bool]$regression.UserForm)
      $regressionEntries+=('results/'+$name+'.json')
      $regressionEntries+=('logs/'+$name+'.stdout.redacted.txt')
      $regressionEntries+=('logs/'+$name+'.stderr.redacted.txt')
      if([bool]$regression.UserForm){
        $buildPath=Join-Path $evidence ('build/'+$name+'.UserFormBuild.json')
        if(-not(Test-Path -LiteralPath $buildPath -PathType Leaf)){Stop-Code 24 ($name+' UserForm build evidence missing')}
        $build=Get-Content -LiteralPath $buildPath -Raw -Encoding UTF8|ConvertFrom-Json
        $expectedManifest=Join-Path $dataRoot ([string]$regression.Manifest)
        if($build.run_id -ne $Manifest['run_id'] -or $build.suite -ne $name -or $build.owner -ne [string]$regression.Owner -or $build.source_digest -ne $Manifest['base17_digest'] -or $build.snapshot_digest -ne $Manifest['snapshot_digest'] -or $build.extension_manifest_sha256 -ne (Hash $expectedManifest) -or (Hash $buildPath) -ne [string]$regEvidence.build_evidence_sha256){Stop-Code 24 ($name+' UserForm build evidence invalid')}
        Copy-Item -LiteralPath $buildPath -Destination (Join-Path $returnRoot ('results/build/'+$name+'.UserFormBuild.json')) -Force
        $regressionEntries+=('results/build/'+$name+'.UserFormBuild.json')
      }
    }
    $script:phase='integrity';Assert-SourceBundleUnchanged $Manifest $sourceBundle;Assert-StagedDataExact $Manifest $LocalBundle
'''.replace("__REGRESSION_ROWS__", ",\n".join(rows))


def orchestrator(suite: str, launcher: str, expected: int, extension_manifest: str | None) -> str:
    script = r'''param([Parameter(Mandatory=$true)][string]$OriginalBundle)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$sourceBundle=[IO.Path]::GetFullPath($OriginalBundle).TrimEnd('\')
$phase='manifest';$runId=$null;$local=$null;$resultRoot=$null
$exitCode=0;$failure=$null;$candidateName=$null;$candidateBytes=$null;$candidate_sha256=$null
$excelMutex=$null;$excelMutexAcquired=$false
$allowedExitCodes=@(0,10,20,21,22,23,24,25,26)
# return/PASS.zip or return/DIAGNOSTIC.zip; 10=preflight,20=integrity,21=stale,22=build/Green compile,23=Red,24=evidence,25=cleanup/temp,26=return.
function Stop-Code([int]$Code,[string]$Message){throw ([Exception]::new(([string]$Code)+'|'+$Message))}
function Enter-ExcelComGate {
  try {
    $created=$false
    $script:excelMutex=[Threading.Mutex]::new($false,'Global\LHexcel-v0.01-Excel-COM-v001',[ref]$created)
    if(-not $script:excelMutex.WaitOne(0)) { Stop-Code 10 'another lane owns the Excel COM gate' }
    $script:excelMutexAcquired=$true
  } catch [Threading.AbandonedMutexException] {
    Stop-Code 10 'Excel COM gate was abandoned; refusing concurrent recovery'
  } catch {
    if($_.Exception.Message -match '^10\|'){throw}
    Stop-Code 10 ('Excel COM gate unavailable: '+$_.Exception.Message)
  }
}
function Exit-ExcelComGate {
  if($null -ne $script:excelMutex) {
    try { if($script:excelMutexAcquired){[void]$script:excelMutex.ReleaseMutex()} } finally { $script:excelMutex.Dispose();$script:excelMutex=$null;$script:excelMutexAcquired=$false }
  }
}
function Error-Code($ErrorRecord){if($ErrorRecord.Exception.Message -match '^([0-9]+)\|'){return [int]$Matches[1]};return 10}
function Write-Utf8NoBom([string]$Path,[string]$Text){$utf8=New-Object Text.UTF8Encoding($false);[IO.File]::WriteAllText($Path,$Text,$utf8)}
function Write-Json([string]$Path,$Object){Write-Utf8NoBom $Path ($Object|ConvertTo-Json -Depth 20)}
function Hash([string]$Path){(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Test-ExactFileIdentity([string]$Path,[Int64]$ExpectedSize,[string]$ExpectedHash){
  for($attempt=0;$attempt -lt 5;$attempt++){
    try {
      $item=[IO.FileInfo]::new($Path)
      $item.Refresh()
      $identityMatches=($item.Exists -and [Int64]$item.Length -eq $ExpectedSize -and (Hash $Path) -eq $ExpectedHash)
      if($identityMatches){return $true}
    } catch { }
    if($attempt -ge 4){return $false}
    Start-Sleep -Milliseconds 100
  }
  return $false
}
function Redact([string]$Text){
  $pathPattern='(?i)(?<![A-Za-z0-9_.-])(?:[A-Z]:\\(?:[^\\/:*?"<>|\r\n]+\\)*[^\\/:*?"<>|\r\n]*|\\\\[^\\/:*?"<>|\r\n]+\\[^\\/:*?"<>|\r\n]+(?:\\[^\\/:*?"<>|\r\n]+)*)'
  $Text -replace '(?i)Bearer\s+[^\s"'']+','Bearer [REDACTED]' -replace '(?i)(password|secret|token|apikey|key)\s*[=:]\s*[^\s]+','$1=[REDACTED]' -replace $pathPattern,'<PATH>'
}
function Redact-JsonValue($Value){
  if($null -eq $Value){return $null}
  if($Value -is [string]){return (Redact ([string]$Value))}
  if($Value -is [Collections.IDictionary]){$copy=[ordered]@{};foreach($key in $Value.Keys){$copy[[string]$key]=Redact-JsonValue $Value[$key]};return $copy}
  if($Value -is [pscustomobject]){$copy=[ordered]@{};foreach($property in $Value.PSObject.Properties){$copy[$property.Name]=Redact-JsonValue $property.Value};return $copy}
  if($Value -is [Collections.IEnumerable]){$items=@();foreach($item in $Value){$items+=,(Redact-JsonValue $item)};return ,$items}
  return $Value
}
function Write-RedactedJson([string]$Source,[string]$Destination){$document=Get-Content -LiteralPath $Source -Raw -Encoding UTF8|ConvertFrom-Json;Write-Json $Destination (Redact-JsonValue $document)}
function Diagnostic-Line([string]$Text){(Redact $Text) -replace '\r?\n',' | '}
function Recompute-Digest($Records){$canonical='';foreach($record in @($Records)){$canonical+=[string]$record['path']+"`n"+[string]$record['sha256']+"`n"};$algorithm=[Security.Cryptography.SHA256]::Create();try{([BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes($canonical)))).Replace('-','').ToLowerInvariant()}finally{$algorithm.Dispose()}}
function Parse-ExactSha256Sums([string]$Path){$seen=@{};$records=@();$previous=$null;foreach($line in @([IO.File]::ReadAllLines($Path,[Text.Encoding]::UTF8))){if($line -notmatch '^([0-9a-f]{64})  ([^\\/].*)$'){Stop-Code 20 'malformed SHA256SUMS line'};$relative=$Matches[2];$segments=@($relative.Split('/'));if($relative.Contains('\') -or [IO.Path]::IsPathRooted($relative) -or @($segments|Where-Object {$_ -eq '' -or $_ -eq '.' -or $_ -eq '..'}).Count -gt 0){Stop-Code 20 'unsafe SHA256SUMS path'};if($seen.ContainsKey($relative)){Stop-Code 20 'duplicate SHA256SUMS path'};if($null -ne $previous -and [string]::CompareOrdinal($previous,$relative) -ge 0){Stop-Code 20 'SHA256SUMS paths are not in UTF-8 ordinal order'};$previous=$relative;$seen[$relative]=$true;$records+=@{path=$relative;sha256=$Matches[1]}};return $records}
function New-ExactUtf8Zip([string]$Destination,[string]$Root,[string[]]$RelativeNames){
  Add-Type -AssemblyName System.IO.Compression;Add-Type -AssemblyName System.IO.Compression.FileSystem
  Remove-Item -LiteralPath $Destination -Force -ErrorAction SilentlyContinue
  $fileStream=[IO.File]::Open($Destination,[IO.FileMode]::CreateNew,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
  try {
    $archive=[IO.Compression.ZipArchive]::new($fileStream,[IO.Compression.ZipArchiveMode]::Create,$true,[Text.Encoding]::UTF8)
    try {foreach($relative in $RelativeNames){$source=Join-Path $Root ($relative.Replace('/','\'));if(-not(Test-Path -LiteralPath $source -PathType Leaf)){Stop-Code 26 ('ZIP source missing: '+$relative)};$entry=$archive.CreateEntry($relative,[IO.Compression.CompressionLevel]::Optimal);$input=[IO.File]::OpenRead($source);$output=$entry.Open();try{$input.CopyTo($output)}finally{$output.Dispose();$input.Dispose()}}}finally{$archive.Dispose()}
  }finally{$fileStream.Dispose()}
}
function Assert-ZipExact([string]$ArchivePath,[string]$Root,[string[]]$RelativeNames){
  Add-Type -AssemblyName System.IO.Compression.FileSystem;$archive=[IO.Compression.ZipFile]::OpenRead($ArchivePath)
  try {
    $actual=@($archive.Entries|Where-Object {-not $_.FullName.EndsWith('/')}|ForEach-Object {$_.FullName})
    if((@($actual|Sort-Object)-join '|') -ne (@($RelativeNames|Sort-Object)-join '|')){Stop-Code 26 'ZIP exact entry set mismatch'}
    foreach($relative in $RelativeNames){$entry=@($archive.Entries|Where-Object {$_.FullName -eq $relative});if($entry.Count -ne 1){Stop-Code 26 ('required ZIP entry missing: '+$relative)};$stream=$entry[0].Open();$sha=[Security.Cryptography.SHA256]::Create();try{$entryHash=([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose();$stream.Dispose()};if($entryHash -ne (Hash (Join-Path $Root ($relative.Replace('/','\'))))){Stop-Code 26 ('reopened ZIP entry hash mismatch: '+$relative)}}
  }finally{$archive.Dispose()}
}
function Remove-StagingTree([string]$Path){
  if([string]::IsNullOrWhiteSpace($Path)){return}
  if(Test-Path -LiteralPath $Path){& attrib.exe -R (Join-Path $Path '*') /S /D 2>$null|Out-Null;Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue}
  if(Test-Path -LiteralPath $Path){Stop-Code 25 'ASCII TEMP cleanup failed'}
}
function Copy-VerifiedDataStage($Manifest,[string]$Source,[string]$Destination){
  foreach($record in @($Manifest['staged_data_records'])){
    $relative=[string]$record['path']
    if($relative -notmatch '__STAGED_DATA_PATTERN__' -or ($relative.ToLowerInvariant() -match '\.(ps1|bat|cmd|exe|dll)$' -and $relative -ne '__ALLOWED_RUNTIME_SCRIPT__')){Stop-Code 20 ('unsafe staged data path: '+$relative)}
    $sourcePath=Join-Path $Source $relative;$targetPath=Join-Path $Destination $relative
    if(-not(Test-Path -LiteralPath $sourcePath -PathType Leaf)){Stop-Code 21 ('staged source missing: '+$relative)}
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $targetPath)|Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath -Force
    if(-not(Test-ExactFileIdentity $targetPath ([Int64]$record['size']) ([string]$record['sha256']))){Stop-Code 21 ('staged data copy changed: '+$relative)}
  }
}
function Invoke-CapturedPowerShell([string[]]$Arguments,[string]$StdOut,[string]$StdErr){
  $previousPreference=$ErrorActionPreference
  try {
    # Windows PowerShell 5.1 can promote a child native stderr record to a
    # terminating NativeCommandError when the parent uses Stop.  Capture the
    # child's real exit code and redacted stderr instead of losing both.
    $ErrorActionPreference='Continue'
    & powershell.exe @Arguments 1> $StdOut 2> $StdErr
    $childExitCode=[int]$LASTEXITCODE
  } finally { $ErrorActionPreference=$previousPreference }
  return $childExitCode
}
function Get-BytesSha256([byte[]]$Bytes){$sha=[Security.Cryptography.SHA256]::Create();try{([BitConverter]::ToString($sha.ComputeHash($Bytes))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}}
function Publish-VerifiedZipBytes([byte[]]$Bytes,[string]$Name,[string]$ExpectedHash){
  $destination=Join-Path $sourceBundle 'return'
  Remove-Item -LiteralPath $destination -Recurse -Force -ErrorAction SilentlyContinue
  New-Item -ItemType Directory -Force -Path $destination|Out-Null
  $stage=Join-Path $destination ($Name+'.tmp-'+$runId);$final=Join-Path $destination $Name
  [IO.File]::WriteAllBytes($stage,$Bytes)
  if((Hash $stage) -ne $ExpectedHash){Stop-Code 26 'return ZIP staged hash mismatch'}
  Move-Item -LiteralPath $stage -Destination $final -Force
  if((Hash $final) -ne $ExpectedHash){Stop-Code 26 'return ZIP final hash mismatch'}
  $files=@(Get-ChildItem -LiteralPath $destination -File)
  if($files.Count -ne 1 -or $files[0].Name -ne $Name){Stop-Code 26 'return must contain exactly one ZIP'}
}
function Read-Manifest([string]$Bundle){
  $path=Join-Path $Bundle 'manifest.json';if(-not(Test-Path -LiteralPath $path -PathType Leaf)){Stop-Code 20 'manifest missing'}
  Add-Type -AssemblyName System.Web.Extensions;$item=(New-Object System.Web.Script.Serialization.JavaScriptSerializer).DeserializeObject([IO.File]::ReadAllText($path,[Text.Encoding]::UTF8))
  if($item['schema_version'].GetType().FullName -ne 'System.Int32' -or $item['schema_version'] -ne 1){Stop-Code 20 'schema_version must be exact System.Int32 1'}
  foreach($name in @('run_id','created_utc','approved_commit','packaging_commit','suite','requested_excel_bitness','extension_manifest','extension_manifest_digest','base17_records','base17_digest','entrypoint_records','entrypoint_digest','packager_code_digest','launcher_digest','guide_digest','packager_digest','staged_data_records','staged_data_digest','file_records','snapshot_digest','sha256sums_digest')){if(-not $item.ContainsKey($name)){Stop-Code 20 ('manifest field missing: '+$name)}}
  if([string]$item['run_id'] -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$'){Stop-Code 20 'manifest run_id invalid'}
  if([string]$item['suite'] -ne '__SUITE__'){Stop-Code 20 'manifest suite does not match launcher'}
  return $item
}
function Assert-BootstrapIdentity($Manifest){
  $records=@($Manifest['file_records']|Where-Object {$_.path -eq 'payload/Invoke-LHexcelVerify.ps1'})
  if($records.Count -ne 1 -or (Hash $PSCommandPath) -ne [string]$records[0]['sha256']){Stop-Code 20 'orchestrator bootstrap hash mismatch'}
  if((Hash (Join-Path $sourceBundle '__LAUNCHER__')) -ne [string]$Manifest['launcher_digest']){Stop-Code 20 'launcher digest mismatch'}
}
function Assert-SourceBundleUnchanged($Manifest,[string]$Bundle){
  $root=[IO.Path]::GetFullPath($Bundle).TrimEnd('\');$sum=Join-Path $root 'SHA256SUMS';if(-not(Test-Path -LiteralPath $sum -PathType Leaf)){Stop-Code 20 'SHA256SUMS missing'}
  $sumRecords=@(Parse-ExactSha256Sums $sum);$expected=@{}
  foreach($record in @($Manifest['file_records'])){$relative=[string]$record['path'];if($expected.ContainsKey($relative)){Stop-Code 20 'duplicate manifest file_records path'};$expected[$relative]=[string]$record['sha256'];$file=Join-Path $root $relative;if(-not(Test-ExactFileIdentity $file ([Int64]$record['size']) ([string]$record['sha256']))){Stop-Code 21 ('snapshot byte identity changed: '+$relative)}}
  $actual=@(Get-ChildItem -LiteralPath $root -Force -Recurse -File|ForEach-Object {$_.FullName.Substring($root.Length+1).Replace('\','/')}|Where-Object {$_ -notin @('manifest.json','SHA256SUMS') -and $_ -notmatch '^(status|return|runtime)/' -and [IO.Path]::GetFileName($_) -ne '.DS_Store'}|Sort-Object)
  if((@($actual)-join '|') -ne (@($expected.Keys|Sort-Object)-join '|')){Stop-Code 20 'file_records exact set/count mismatch'}
  if($sumRecords.Count -ne $expected.Count){Stop-Code 20 'SHA256SUMS exact set/count mismatch'}
  foreach($record in $sumRecords){if($record['path'] -eq 'SHA256SUMS' -or $record['path'] -eq 'manifest.json' -or -not $expected.ContainsKey($record['path']) -or $expected[$record['path']] -ne $record['sha256']){Stop-Code 20 'SHA256SUMS malformed, missing, or extra entry'}}
  foreach($path in $expected.Keys){if(@($sumRecords|Where-Object {$_.path -eq $path}).Count -ne 1){Stop-Code 20 'SHA256SUMS missing expected entry'}}
  if((Recompute-Digest $Manifest['base17_records']) -ne $Manifest['base17_digest'] -or (Recompute-Digest $Manifest['entrypoint_records']) -ne $Manifest['entrypoint_digest'] -or (Recompute-Digest $Manifest['file_records']) -ne $Manifest['snapshot_digest']){Stop-Code 20 'manifest digest recomputation failed'}
  if($Manifest['sha256sums_digest'] -ne (Hash $sum)){Stop-Code 20 'SHA256SUMS byte digest changed'}
  $guideHash=Hash (Join-Path $root '검증_안내.md');$launcherHash=Hash (Join-Path $root '__LAUNCHER__');$canonical="launcher`n$launcherHash`nguide`n$guideHash`npackager`n$($Manifest['packager_code_digest'])`n";$algorithm=[Security.Cryptography.SHA256]::Create();try{$aggregate=([BitConverter]::ToString($algorithm.ComputeHash([Text.Encoding]::UTF8.GetBytes($canonical)))).Replace('-','').ToLowerInvariant()}finally{$algorithm.Dispose()}
  if($launcherHash -ne $Manifest['launcher_digest'] -or $guideHash -ne $Manifest['guide_digest'] -or $aggregate -ne $Manifest['packager_digest']){Stop-Code 20 'launcher digest mismatch; guide digest mismatch'}
  if(-not [string]::IsNullOrWhiteSpace([string]$Manifest['extension_manifest'])){$extension=Join-Path $root ('payload/'+[string]$Manifest['extension_manifest']);if(-not(Test-Path -LiteralPath $extension -PathType Leaf) -or (Hash $extension) -ne [string]$Manifest['extension_manifest_digest']){Stop-Code 20 'extension manifest digest mismatch'}}
}
function Assert-StagedDataExact($Manifest,[string]$Bundle){
  $root=[IO.Path]::GetFullPath($Bundle).TrimEnd('\');$expected=@{}
  foreach($record in @($Manifest['staged_data_records'])){
    $relative=[string]$record['path']
    if($relative -notmatch '__STAGED_DATA_PATTERN__' -or ($relative.ToLowerInvariant() -match '\.(ps1|bat|cmd|exe|dll)$' -and $relative -ne '__ALLOWED_RUNTIME_SCRIPT__')){Stop-Code 20 ('unsafe staged data path: '+$relative)}
    if($expected.ContainsKey($relative)){Stop-Code 20 'duplicate staged_data_records path'}
    $expected[$relative]=[string]$record['sha256'];$file=Join-Path $root $relative
    if(-not(Test-ExactFileIdentity $file ([Int64]$record['size']) ([string]$record['sha256']))){Stop-Code 21 ('staged data identity changed: '+$relative)}
  }
  $actual=@(Get-ChildItem -LiteralPath $root -Force -Recurse -File|ForEach-Object {$_.FullName.Substring($root.Length+1).Replace('\','/')}|Where-Object {$_ -notmatch '^runtime/'}|Sort-Object)
  if((@($actual)-join '|') -ne (@($expected.Keys|Sort-Object)-join '|')){Stop-Code 20 'staged data exact set/count mismatch'}
  if((Recompute-Digest $Manifest['staged_data_records']) -ne $Manifest['staged_data_digest']){Stop-Code 20 'staged data digest recomputation failed'}
}
function Set-ImmutablePayloadReadOnly([string]$Payload){attrib +R "$Payload\*" /S|Out-Null}
function Require-Evidence([string]$Path,[string]$Mode,$Manifest){
  if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){Stop-Code 24 ($Mode+' evidence missing')};$item=Get-Content -LiteralPath $Path -Raw -Encoding UTF8|ConvertFrom-Json
  if($item.run_id -ne $Manifest['run_id'] -or $item.source_digest -ne $Manifest['base17_digest'] -or $item.snapshot_digest -ne $Manifest['snapshot_digest']){Stop-Code 24 'runtime evidence stale binding'}
  if($null -eq $item.environment -or $null -eq $item.cleanup){Stop-Code 24 'runtime environment/cleanup evidence missing'}
  if($Mode -eq 'Red' -and ($item.status -ne 'EXPECTED_RED' -or $item.classification -ne 'EXPECTED_RED' -or $item.compile.dialog_closed -ne $true)){Stop-Code 24 'Red evidence invalid'}
  # UserForm suite contract
  if($Mode -eq 'Green' -and ($item.status -ne 'PASS' -or [int]$item.run.passed -ne __EXPECTED__ -or [int]$item.run.failed -ne 0 -or [int]$item.run.skipped -ne 0 -or ('__SUITE__' -in @('AI','Template','Data','Draw','File','Calculator','Symbols') -and [string]$item.build_evidence_sha256 -notmatch '^[0-9a-f]{64}$'))){Stop-Code 24 'Green runtime evidence invalid'}
  return $item
}
function Require-UserFormBuildEvidence([string]$Path,$Green,$Manifest){
  if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){Stop-Code 24 'UserForm build evidence missing'}
  $item=Get-Content -LiteralPath $Path -Raw -Encoding UTF8|ConvertFrom-Json
  if($item.run_id -ne $Manifest['run_id'] -or $item.source_digest -ne $Manifest['base17_digest'] -or $item.snapshot_digest -ne $Manifest['snapshot_digest']){Stop-Code 24 'UserForm build evidence stale binding'}
  if($item.suite -ne '__SUITE__' -or $item.owner -ne '__OWNER__' -or $item.mode -ne 'Green' -or $item.extension_manifest_sha256 -ne $Manifest['extension_manifest_digest']){Stop-Code 24 'UserForm build evidence suite or manifest binding rejected'}
  $digest=Hash $Path
  if($digest -ne [string]$Green.build_evidence_sha256){Stop-Code 24 'UserForm build evidence digest does not match Green evidence'}
  return [pscustomobject]@{document=$item;sha256=$digest}
}
function Require-RegressionEvidence([string]$Path,$Manifest,[string]$SuiteName,[int]$ExpectedCount,[bool]$RequiresBuildEvidence){
  if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){Stop-Code 24 ($SuiteName+' regression evidence missing')}
  $item=Get-Content -LiteralPath $Path -Raw -Encoding UTF8|ConvertFrom-Json
  if($item.run_id -ne $Manifest['run_id'] -or $item.source_digest -ne $Manifest['base17_digest'] -or $item.snapshot_digest -ne $Manifest['snapshot_digest']){Stop-Code 24 ($SuiteName+' regression evidence stale binding')}
  if($item.suite -ne $SuiteName -or $item.status -ne 'PASS' -or [int]$item.run.total -ne $ExpectedCount -or [int]$item.run.passed -ne $ExpectedCount -or [int]$item.run.failed -ne 0 -or [int]$item.run.skipped -ne 0){Stop-Code 24 ($SuiteName+' regression evidence invalid')}
  if($item.environment.excel_bitness -ne ($Manifest['requested_excel_bitness']+'-bit')){Stop-Code 10 ($SuiteName+' Excel bitness mismatch')}
  if($RequiresBuildEvidence -and [string]$item.build_evidence_sha256 -notmatch '^[0-9a-f]{64}$'){Stop-Code 24 ($SuiteName+' UserForm build evidence digest missing')}
  return $item
}
function New-VerificationResult([int]$Code,[string]$Failure,[string]$ReturnRoot,[string[]]$Entries){[pscustomobject]@{ExitCode=$Code;Failure=$Failure;ReturnRoot=$ReturnRoot;Entries=@($Entries)}}
function Invoke-LocalVerification($Manifest,[string]$LocalBundle,[string]$ResultRoot){
  $payload=Join-Path $LocalBundle 'payload';$runtime=Join-Path $LocalBundle 'runtime';$evidence=Join-Path $runtime 'evidence/vba';$scratch=Join-Path $runtime ('scratch-'+$Manifest['run_id']);$returnRoot=Join-Path $ResultRoot 'return-content'
  $dataRoot=Join-Path $LocalBundle 'payload'
  $code=0;$localFailure=$null;$entries=@();$regressionEntries=@();$redOut=Join-Path $scratch 'red.out';$redErr=Join-Path $scratch 'red.err';$greenOut=Join-Path $scratch 'green.out';$greenErr=Join-Path $scratch 'green.err';$productOut=Join-Path $scratch 'product.out';$productErr=Join-Path $scratch 'product.err'
  try {
    New-Item -ItemType Directory -Force -Path $evidence,$scratch,$returnRoot,(Join-Path $returnRoot 'results'),(Join-Path $returnRoot 'results/build'),(Join-Path $returnRoot 'logs')|Out-Null
    $script:phase='integrity';Assert-SourceBundleUnchanged $Manifest $sourceBundle;Assert-StagedDataExact $Manifest $LocalBundle
    $suite=Join-Path $sourceBundle 'payload/tests/windows/Run-__SUITE_SCRIPT__.ps1'
    if('__SUITE__' -in @('ProductRibbon','ProductUi')){
      $script:phase='product'
      if([string]$Manifest['source_tree_sha256'] -notmatch '^[0-9a-f]{64}$'){Stop-Code 20 'Product source tree digest missing'}
      $productArguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',$suite,'-Suite','__SUITE__','-Mode','Green','-RunId',$Manifest['run_id'],'-SourceDigest',$Manifest['source_tree_sha256'],'-SnapshotDigest',$Manifest['snapshot_digest'],'-DataRoot',$dataRoot,'-EvidenceRoot',$evidence)
      $productCode=Invoke-CapturedPowerShell -Arguments $productArguments -StdOut $productOut -StdErr $productErr
      $script:phase='integrity';Assert-SourceBundleUnchanged $Manifest $sourceBundle;Assert-StagedDataExact $Manifest $LocalBundle;$script:phase='product'
      if($productCode -ne 0){
        $productExitCode=if($productCode -in @(10,22,23,24,25)){$productCode}else{24}
        $productFailurePath=Join-Path $evidence '__SUITE__.json'
        if(Test-Path -LiteralPath $productFailurePath -PathType Leaf){
          $productFailureEvidence=Get-Content -LiteralPath $productFailurePath -Raw -Encoding UTF8|ConvertFrom-Json
          $failureFields=@($productFailureEvidence.PSObject.Properties.Name)
          if($failureFields -contains 'exit_code' -and $failureFields -contains 'failure_phase' -and [int]$productFailureEvidence.exit_code -eq $productExitCode -and [string]$productFailureEvidence.failure_phase -match '^[A-Za-z][A-Za-z0-9_-]*$'){$script:phase=[string]$productFailureEvidence.failure_phase}
        }
        Stop-Code $productExitCode '__SUITE__ probe failed'
      }
      $productEvidencePath=Join-Path $evidence '__SUITE__.json';if(-not(Test-Path -LiteralPath $productEvidencePath -PathType Leaf)){Stop-Code 24 '__SUITE__ evidence missing'}
      $productEvidence=Get-Content -LiteralPath $productEvidencePath -Raw -Encoding UTF8|ConvertFrom-Json
      if($productEvidence.status -ne 'PASS' -or $productEvidence.run_id -ne $Manifest['run_id'] -or $productEvidence.source_digest -ne $Manifest['source_tree_sha256'] -or $productEvidence.snapshot_digest -ne $Manifest['snapshot_digest'] -or [int]$productEvidence.run.passed -ne __EXPECTED__ -or [int]$productEvidence.run.failed -ne 0){Stop-Code 24 '__SUITE__ evidence binding rejected'}
      Write-RedactedJson $productEvidencePath (Join-Path $returnRoot '__SUITE__.json')
      [string[]]$prerequisiteSummaryNames=@()
      if('__SUITE__' -eq 'ProductUi'){
        $prerequisiteSummaryPath=Join-Path $evidence 'ProductRibbon.json';if(-not(Test-Path -LiteralPath $prerequisiteSummaryPath -PathType Leaf)){Stop-Code 24 'ProductRibbon prerequisite summary missing'}
        $prerequisiteSummary=Get-Content -LiteralPath $prerequisiteSummaryPath -Raw -Encoding UTF8|ConvertFrom-Json
        if($prerequisiteSummary.status -ne 'PASS' -or $prerequisiteSummary.run_id -ne $Manifest['run_id'] -or $prerequisiteSummary.source_digest -ne $Manifest['source_tree_sha256'] -or $prerequisiteSummary.snapshot_digest -ne $Manifest['snapshot_digest'] -or [int]$prerequisiteSummary.run.passed -ne 3 -or [int]$prerequisiteSummary.run.failed -ne 0){Stop-Code 24 'ProductRibbon prerequisite summary binding rejected'}
        Write-RedactedJson $prerequisiteSummaryPath (Join-Path $returnRoot 'ProductRibbon.json')
        $prerequisiteSummaryNames=@('ProductRibbon.json')
      }
      $productNames=__PRODUCT_NAMES__
      $artifact=Join-Path $evidence 'Product.xlam';if(-not(Test-Path -LiteralPath $artifact -PathType Leaf)){Stop-Code 24 'Product.xlam evidence missing'}
      $artifactHash=Hash $artifact;if($artifactHash -ne [string]$productEvidence.artifact_sha256){Stop-Code 24 'Product artifact digest rejected'}
      foreach($name in $productNames){
        $p=Join-Path $evidence $name;if(-not(Test-Path -LiteralPath $p -PathType Leaf)){Stop-Code 24 ('Product evidence missing: '+$name)}
        $item=Get-Content -LiteralPath $p -Raw -Encoding UTF8|ConvertFrom-Json
        if($item.status -ne 'PASS' -or $item.run_id -ne $Manifest['run_id'] -or $item.artifact_sha256 -ne $artifactHash -or $item.source_tree_sha256 -ne $Manifest['source_tree_sha256'] -or $item.source_snapshot_sha256 -ne $Manifest['snapshot_digest']){Stop-Code 24 ('Product evidence binding rejected: '+$name)}
        Copy-Item -LiteralPath $p -Destination (Join-Path $returnRoot $name) -Force
      }
      Copy-Item -LiteralPath $artifact -Destination (Join-Path $returnRoot 'Product.xlam') -Force
      [string[]]$productPaths=@('Product.xlam')
      $productFiles=@(@{path='Product.xlam';sha256=$artifactHash})
      $productBundleSha=Recompute-Digest $productFiles
      $resultManifest=[ordered]@{schema_version=1;suite='__SUITE__';run_id=$Manifest['run_id'];packaging_commit=$Manifest['packaging_commit'];artifact_sha256=$artifactHash;source_tree_sha256=$Manifest['source_tree_sha256'];source_snapshot_sha256=$Manifest['snapshot_digest'];product_bundle_sha256=$productBundleSha;product_files=$productFiles}
      Write-Json (Join-Path $returnRoot 'manifest.json') $resultManifest
      [string[]]$sumNames=@('__SUITE__.json')+@($prerequisiteSummaryNames)+@($productNames)+@($productPaths);[Array]::Sort($sumNames,[StringComparer]::Ordinal);$sumLines=@();foreach($name in $sumNames){$sumLines+=((Hash (Join-Path $returnRoot ($name.Replace('/','\\'))))+'  '+$name)}
      Write-Utf8NoBom (Join-Path $returnRoot 'SHA256SUMS') (($sumLines -join "`n")+"`n")
      $script:phase='integrity';Assert-SourceBundleUnchanged $Manifest $sourceBundle;Assert-StagedDataExact $Manifest $LocalBundle;$script:phase='package'
      $entries=@('manifest.json','SHA256SUMS')+@($sumNames);Remove-StagingTree $scratch;return New-VerificationResult 0 $null $returnRoot $entries
    }
    $redArguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',$suite,'-Suite','Core','-Mode','RedProbe','-ExpectedCompileFailure','CNxStateGuard','-RunId',$Manifest['run_id'],'-SourceDigest',$Manifest['base17_digest'],'-SnapshotDigest',$Manifest['snapshot_digest'],'-DataRoot',$dataRoot,'-EvidenceRoot',$evidence)
    $script:phase='red'
    $redCode=Invoke-CapturedPowerShell -Arguments $redArguments -StdOut $redOut -StdErr $redErr
    $script:phase='integrity';Assert-SourceBundleUnchanged $Manifest $sourceBundle;Assert-StagedDataExact $Manifest $LocalBundle;$script:phase='red'
    if($redCode -ne 1){if(Test-Path -LiteralPath $redErr -PathType Leaf){$detail=Diagnostic-Line ([IO.File]::ReadAllText($redErr));if(-not [string]::IsNullOrWhiteSpace($detail)){Write-Information -InformationAction Continue ('redprobe_stderr='+$detail)}};$redEvidence=Join-Path $evidence 'Core.RedProbe.json';if(Test-Path -LiteralPath $redEvidence -PathType Leaf){$item=Get-Content -LiteralPath $redEvidence -Raw -Encoding UTF8|ConvertFrom-Json;if(-not [string]::IsNullOrWhiteSpace([string]$item.failure)){Write-Information -InformationAction Continue ('redprobe_evidence_failure='+(Diagnostic-Line ([string]$item.failure)))}};$mappedRedCode=if($redCode -in @(10,22,23,24,25)){$redCode}else{23};Stop-Code $mappedRedCode 'RedProbe failed'}
    $red=Require-Evidence (Join-Path $evidence 'Core.RedProbe.json') 'Red' $Manifest;Copy-Item -LiteralPath (Join-Path $evidence 'Core.RedProbe.json') -Destination (Join-Path $returnRoot 'results/Core.RedProbe.json') -Force
    $script:phase='integrity';Assert-SourceBundleUnchanged $Manifest $sourceBundle;Assert-StagedDataExact $Manifest $LocalBundle
    $greenArguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',$suite,'-Suite','__SUITE__','-Mode','Green','-RunId',$Manifest['run_id'],'-SourceDigest',$Manifest['base17_digest'],'-SnapshotDigest',$Manifest['snapshot_digest'],'-DataRoot',$dataRoot,'-EvidenceRoot',$evidence__EXTENSION_ARGUMENT__)
    $script:phase='green'
    $greenCode=Invoke-CapturedPowerShell -Arguments $greenArguments -StdOut $greenOut -StdErr $greenErr
    $script:phase='integrity';Assert-SourceBundleUnchanged $Manifest $sourceBundle;Assert-StagedDataExact $Manifest $LocalBundle;$script:phase='green'
    if($greenCode -ne 0){$mappedGreenCode=if($greenCode -in @(10,22,23,24,25)){$greenCode}else{22};Stop-Code $mappedGreenCode 'Green verification failed'}
    $green=Require-Evidence (Join-Path $evidence '__SUITE__.json') 'Green' $Manifest;if($green.environment.excel_bitness -ne ($Manifest['requested_excel_bitness']+'-bit')){Stop-Code 10 'actual Excel bitness does not match requested bitness'};Copy-Item -LiteralPath (Join-Path $evidence '__SUITE__.json') -Destination (Join-Path $returnRoot 'results/__SUITE__.json') -Force
    if('__SUITE__' -in @('AI','Template','Data','Draw','File','Calculator','Symbols')){$buildEvidencePath=Join-Path $evidence 'build/__SUITE__.UserFormBuild.json';$buildEvidence=Require-UserFormBuildEvidence $buildEvidencePath $green $Manifest;Copy-Item -LiteralPath $buildEvidencePath -Destination (Join-Path $returnRoot 'results/build/__SUITE__.UserFormBuild.json') -Force}
__SUITE_REGRESSION_VERIFICATION__
    foreach($pair in @(@($redOut,'red.stdout.redacted.txt'),@($redErr,'red.stderr.redacted.txt'),@($greenOut,'green.stdout.redacted.txt'),@($greenErr,'green.stderr.redacted.txt'))){Write-Utf8NoBom (Join-Path $returnRoot ('logs/'+$pair[1])) (Redact ([IO.File]::ReadAllText($pair[0])))}
    $script:phase='integrity';Assert-SourceBundleUnchanged $Manifest $sourceBundle;Assert-StagedDataExact $Manifest $LocalBundle
  }catch{$code=Error-Code $_;$localFailure=$_.Exception.Message}
  foreach($name in @('manifest.json','SHA256SUMS','__LAUNCHER__','검증_안내.md')){$source=Join-Path $sourceBundle $name;if(Test-Path -LiteralPath $source -PathType Leaf){Copy-Item -LiteralPath $source -Destination $returnRoot -Force}}
  if($null -ne $localFailure){
    foreach($pair in @(@($redOut,'red.stdout.redacted.txt'),@($redErr,'red.stderr.redacted.txt'),@($greenOut,'green.stdout.redacted.txt'),@($greenErr,'green.stderr.redacted.txt'))){if(Test-Path -LiteralPath $pair[0] -PathType Leaf){Write-Utf8NoBom (Join-Path $returnRoot ('logs/'+$pair[1])) (Redact ([IO.File]::ReadAllText($pair[0])))}}
    if('__SUITE__' -in @('ProductRibbon','ProductUi')){
      foreach($pair in @(@($productOut,'product.stdout.redacted.txt'),@($productErr,'product.stderr.redacted.txt'))){if(Test-Path -LiteralPath $pair[0] -PathType Leaf){Write-Utf8NoBom (Join-Path $returnRoot ('logs/'+$pair[1])) (Redact ([IO.File]::ReadAllText($pair[0])))}}
      $productDiagnostic=Join-Path $evidence '__SUITE__.json';if(Test-Path -LiteralPath $productDiagnostic -PathType Leaf){Write-RedactedJson $productDiagnostic (Join-Path $returnRoot 'results/__SUITE__.json')}
    }
    foreach($name in @('Core.RedProbe.json','Core.RedProbe.OwnerStartup.json','__SUITE__.json')){$source=Join-Path $evidence $name;if(Test-Path -LiteralPath $source -PathType Leaf){Copy-Item -LiteralPath $source -Destination (Join-Path $returnRoot ('results/'+$name)) -Force}}
    if('__SUITE__' -in @('AI','Template','Data','Draw','File','Calculator','Symbols')){$failedBuildEvidence=Join-Path $evidence 'build/__SUITE__.UserFormBuild.json';if(Test-Path -LiteralPath $failedBuildEvidence -PathType Leaf){Copy-Item -LiteralPath $failedBuildEvidence -Destination (Join-Path $returnRoot 'results/build/__SUITE__.UserFormBuild.json') -Force}}
    Write-Json (Join-Path $returnRoot 'diagnostic.json') ([ordered]@{schema_version=1;run_id=$Manifest['run_id'];exit_code=$code;phase=$script:phase;failure=(Redact $localFailure)})
  }
  try{Remove-StagingTree $scratch}catch{$code=25;$localFailure=$_.Exception.Message;Write-Json (Join-Path $returnRoot 'diagnostic.json') ([ordered]@{schema_version=1;run_id=$Manifest['run_id'];exit_code=$code;phase='cleanup';failure=(Redact $localFailure)})}
  $entries=@('manifest.json','SHA256SUMS','__LAUNCHER__','검증_안내.md')
  if($code -eq 0){$entries+=@('results/Core.RedProbe.json','results/__SUITE__.json','logs/red.stdout.redacted.txt','logs/red.stderr.redacted.txt','logs/green.stdout.redacted.txt','logs/green.stderr.redacted.txt');$entries+=@($regressionEntries);if('__SUITE__' -in @('AI','Template','Data','Draw','File','Calculator','Symbols')){$entries+='results/build/__SUITE__.UserFormBuild.json'}}else{$entries+='diagnostic.json';$entries+=@(Get-ChildItem -LiteralPath (Join-Path $returnRoot 'results') -Recurse -File -ErrorAction SilentlyContinue|ForEach-Object {$_.FullName.Substring($returnRoot.Length+1).Replace('\','/')});$entries+=@(Get-ChildItem -LiteralPath (Join-Path $returnRoot 'logs') -File -ErrorAction SilentlyContinue|ForEach-Object {'logs/'+$_.Name})}
  return New-VerificationResult $code $localFailure $returnRoot $entries
}
function New-DiagnosticCandidateBytes([int]$Code,[string]$FailurePhase,[string]$Message,[string]$Identity){
  $safeIdentity=if([string]::IsNullOrWhiteSpace($Identity)){'unknown'}else{$Identity.Replace('-','')};$diagnosticRoot=Join-Path $env:TEMP ('LHexcelVerifyDiagnostic\run-'+$safeIdentity);Remove-StagingTree $diagnosticRoot;New-Item -ItemType Directory -Force -Path $diagnosticRoot|Out-Null
  try {
    Write-Json (Join-Path $diagnosticRoot 'diagnostic.json') ([ordered]@{schema_version=1;run_id=$Identity;exit_code=$Code;phase=$FailurePhase;failure=(Redact $Message)})
    $entries=@('manifest.json','SHA256SUMS','__LAUNCHER__','검증_안내.md','diagnostic.json')
    foreach($name in $entries|Where-Object {$_ -ne 'diagnostic.json'}){$source=Join-Path $sourceBundle $name;if(-not(Test-Path -LiteralPath $source -PathType Leaf)){Stop-Code 26 ('diagnostic contract file missing: '+$name)};Copy-Item -LiteralPath $source -Destination $diagnosticRoot -Force}
    $zip=Join-Path $diagnosticRoot 'DIAGNOSTIC.zip';New-ExactUtf8Zip $zip $diagnosticRoot $entries;Assert-ZipExact $zip $diagnosticRoot $entries;$bytes=[IO.File]::ReadAllBytes($zip);$digest=Get-BytesSha256 $bytes
    return [pscustomobject]@{Name='DIAGNOSTIC.zip';Bytes=$bytes;Sha256=$digest}
  }finally{Remove-StagingTree $diagnosticRoot}
}
try {
  $phase='manifest';$manifest=Read-Manifest $sourceBundle;$runId=[string]$manifest['run_id'];Assert-BootstrapIdentity $manifest
  Enter-ExcelComGate
  $local=Join-Path $env:TEMP ('LHexcelVerify\run-'+$runId.Replace('-',''));$resultRoot=Join-Path $env:TEMP ('LHexcelVerifyResult\run-'+$runId.Replace('-',''))
  $phase='preclean';Remove-StagingTree $local;Remove-StagingTree $resultRoot;$originalReturn=Join-Path $sourceBundle 'return';Remove-Item -LiteralPath $originalReturn -Recurse -Force -ErrorAction SilentlyContinue;if(Test-Path -LiteralPath $originalReturn){Stop-Code 25 'bundle return preclean failed'}
  New-Item -ItemType Directory -Force -Path $local,$resultRoot|Out-Null
  $phase='integrity';Assert-SourceBundleUnchanged $manifest $sourceBundle
  $phase='stage';Copy-VerifiedDataStage $manifest $sourceBundle $local
  $phase='integrity';Assert-StagedDataExact $manifest $local;Set-ImmutablePayloadReadOnly (Join-Path $local 'payload')
  $verificationOutput=@(Invoke-LocalVerification -Manifest $manifest -LocalBundle $local -ResultRoot $resultRoot)
  if($verificationOutput.Count -ne 1){Stop-Code 26 'local verification output contract violated'}
  $verification=$verificationOutput[0];$verificationFields=@($verification.PSObject.Properties.Name|Sort-Object)
  if(($verificationFields -join '|') -ne 'Entries|ExitCode|Failure|ReturnRoot'){Stop-Code 26 'local verification output contract violated'}
  $exitCode=[int]$verification.ExitCode;$failure=$verification.Failure
  if($allowedExitCodes -notcontains $exitCode){Stop-Code 26 ('local verification returned unsupported exit code: '+$exitCode)}
  if('__SUITE__' -in @('ProductRibbon','ProductUi') -and $exitCode -eq 0){$phase='integrity';Assert-SourceBundleUnchanged $manifest $sourceBundle;Assert-StagedDataExact $manifest $local}
  $phase='package';$candidateName=if($exitCode -eq 0){'PASS.zip'}else{'DIAGNOSTIC.zip'};$candidate=Join-Path $resultRoot $candidateName;New-ExactUtf8Zip $candidate $verification.ReturnRoot $verification.Entries;Assert-ZipExact $candidate $verification.ReturnRoot $verification.Entries;$candidateBytes=[IO.File]::ReadAllBytes($candidate);$candidate_sha256=Get-BytesSha256 $candidateBytes
}catch{$exitCode=Error-Code $_;$failure=$_.Exception.Message}
$cleanupFailure=$null
foreach($tree in @($local,$resultRoot)){if(-not [string]::IsNullOrWhiteSpace($tree)){try{Remove-StagingTree $tree}catch{$cleanupFailure=$_.Exception.Message}}}
try { Exit-ExcelComGate } catch { $cleanupFailure=if($null -eq $cleanupFailure){$_.Exception.Message}else{$cleanupFailure} }
if($null -ne $cleanupFailure){$exitCode=25;$failure=$cleanupFailure;$phase='cleanup';$candidateBytes=$null;$candidate_sha256=$null;$candidateName='DIAGNOSTIC.zip'}
if($null -eq $candidateBytes){
  try{$diagnostic=New-DiagnosticCandidateBytes $exitCode $phase $failure $runId;$candidateName=$diagnostic.Name;$candidateBytes=$diagnostic.Bytes;$candidate_sha256=$diagnostic.Sha256}catch{Write-Output ('LHexcel diagnostic failure: '+(Redact $_.Exception.Message));exit 26}
}
try{$phase='publish';Publish-VerifiedZipBytes $candidateBytes $candidateName $candidate_sha256}catch{Write-Output ('LHexcel publish failure: '+(Redact $_.Exception.Message));exit 26}
exit [int]$exitCode
'''
    extension_argument = "" if extension_manifest is None else ",'-ExtensionManifest',(Join-Path $dataRoot '__EXTENSION_MANIFEST__')"
    product_names = {
        "ProductRibbon": "@('ProductRibbon.Package.json','ProductRibbon.RibbonXml.json','ProductRibbon.OfficeIdentity.json','ProductRibbon.Ui.Current200.json','ProductRibbon.Fast.json','ProductRibbon.Guarded.json','ProductRibbon.Planned.json','ProductRibbon.UserForms.json')",
        "ProductUi": "@('ProductRibbon.Package.json','ProductRibbon.RibbonXml.json','ProductRibbon.OfficeIdentity.json','ProductRibbon.Ui.Current200.json','ProductRibbon.Fast.json','ProductRibbon.Guarded.json','ProductRibbon.Planned.json','ProductRibbon.UserForms.json','ProductUi.Current200.json','ProductUi.UserForms.json')",
    }.get(suite, "@()")
    suite_script = {"ProductRibbon": "ProductRibbonSuite", "ProductUi": "ProductUiSuite"}.get(suite, "VbaSuite")
    return (script.replace("__SUITE__", suite)
            .replace("__OWNER__", suite.lower())
            .replace("__LAUNCHER__", launcher)
            .replace("__EXPECTED__", str(expected))
            .replace("__SUITE_SCRIPT__", suite_script)
            .replace("__PRODUCT_NAMES__", product_names)
            .replace("__EXTENSION_ARGUMENT__", extension_argument)
            .replace("__EXTENSION_MANIFEST__", extension_manifest or "")
            .replace("__STAGED_DATA_PATTERN__", STAGED_DATA_PATTERN)
            .replace("__ALLOWED_RUNTIME_SCRIPT__", ALLOWED_RUNTIME_SCRIPT)
            .replace("__SUITE_REGRESSION_VERIFICATION__", suite_regression_verification(suite)))


def batch(suite: str) -> str:
    return r'''@echo off
setlocal EnableExtensions DisableDelayedExpansion
set "ROOT=%~dp0."
set "STATUS_DIR=%~dp0status"
set "RUN_LOCK=%STATUS_DIR%\.running"
if exist "%RUN_LOCK%" (
  echo LHexcel verification exit code: 10
  exit /b 10
)
if not exist "%STATUS_DIR%" mkdir "%STATUS_DIR%"
mkdir "%RUN_LOCK%" 2>nul
if errorlevel 1 (
  echo LHexcel verification exit code: 10
  exit /b 10
)
set "RUN_ID="
for /f "usebackq delims=" %%A in (`powershell.exe -NoProfile -Command "$m=Get-Content -LiteralPath '%~dp0manifest.json' -Raw | ConvertFrom-Json; if([string]$m.run_id -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$'){exit 1}; [Console]::WriteLine($m.run_id)"`) do set "RUN_ID=%%A"
if not defined RUN_ID (
  rmdir /s /q "%RUN_LOCK%" 2>nul
  echo LHexcel verification exit code: 10
  exit /b 10
)
set "RUN_TOKEN=%RUN_ID:~0,12%"
title LHexcel __SUITE__ Verification [run-%RUN_TOKEN%]
echo LHexcel package: __SUITE__
echo Run ID: %RUN_ID%
echo Bundle: %ROOT%
echo.
set "TEMP_BASE="
set "TEMP_SOURCE_KIND="
rem Select an ASCII-only Windows root; fail closed if LOCALAPPDATA and SystemDrive are unsafe.
for /f "usebackq tokens=1,* delims=|" %%A in (`powershell.exe -NoProfile -Command "$p=$env:LOCALAPPDATA;$kind='LOCALAPPDATA'; if([string]::IsNullOrWhiteSpace($p) -or $p -notmatch '^[\x00-\x7F]+$'){$p=$env:SystemDrive;$kind='SYSTEMDRIVE'}; if($p -match '^[A-Za-z]:$'){$p+='\'}; if([string]::IsNullOrWhiteSpace($p) -or $p -notmatch '^[A-Za-z]:\\(?:[\x20-\x7E]+\\?)*$'){exit 1}; [Console]::WriteLine($kind+'|'+$p.TrimEnd('\'))"`) do (
  set "TEMP_SOURCE_KIND=%%A"
  set "TEMP_BASE=%%B"
)
if not defined TEMP_BASE (
  rmdir /s /q "%RUN_LOCK%" 2>nul
  echo LHexcel verification exit code: 10
  exit /b 10
)
set "TEMP_LOCK_ROOT=%TEMP_BASE%\LHexcel\VerifyTemp\.locks"
set "TEMP_LOCK=%TEMP_LOCK_ROOT%\run-%RUN_ID%.lock"
if not exist "%TEMP_LOCK_ROOT%" mkdir "%TEMP_LOCK_ROOT%"
if not exist "%TEMP_LOCK_ROOT%" (
  rmdir /s /q "%RUN_LOCK%" 2>nul
  echo LHexcel verification exit code: 10
  exit /b 10
)
set "TEMP_LOCK_OWNED=0"
mkdir "%TEMP_LOCK%" 2>nul
if errorlevel 1 (
  rmdir /s /q "%RUN_LOCK%" 2>nul
  echo LHexcel verification exit code: 10
  exit /b 10
)
set "TEMP_LOCK_OWNED=1"
set "LANE_TEMP=%TEMP_BASE%\LHexcel\VerifyTemp\__SUITE__-%RUN_TOKEN%"
set "TEMP_PATH_LENGTH="
for /f "usebackq delims=" %%A in (`powershell.exe -NoProfile -Command "$p='%LANE_TEMP%'; if($p.Length -gt 120){$p=($env:SystemDrive.TrimEnd('\')+'\\LHexcel\\VT\\__SUITE__-%RUN_TOKEN%')}; if($p.Length -gt 120){exit 1}; [Console]::WriteLine($p)"`) do set "LANE_TEMP=%%A"
if not defined LANE_TEMP (
  rmdir /s /q "%TEMP_LOCK%" 2>nul
  rmdir /s /q "%RUN_LOCK%" 2>nul
  echo LHexcel verification exit code: 10
  exit /b 10
)
for /f "usebackq delims=" %%A in (`powershell.exe -NoProfile -Command "[Console]::WriteLine(('%LANE_TEMP%').Length)"`) do set "TEMP_PATH_LENGTH=%%A"
if not defined TEMP_PATH_LENGTH (
  rmdir /s /q "%TEMP_LOCK%" 2>nul
  rmdir /s /q "%RUN_LOCK%" 2>nul
  echo LHexcel verification exit code: 10
  exit /b 10
)
if not exist "%LANE_TEMP%" mkdir "%LANE_TEMP%"
if not exist "%LANE_TEMP%" (
  rmdir /s /q "%TEMP_LOCK%" 2>nul
  rmdir /s /q "%RUN_LOCK%" 2>nul
  echo LHexcel verification exit code: 10
  exit /b 10
)
set "TEMP=%LANE_TEMP%"
set "TMP=%LANE_TEMP%"
set "STATUS=%STATUS_DIR%\run-status.txt"
set "OUTPUT=%STATUS_DIR%\launcher-output.txt"
> "%STATUS%" echo wrapper_started=%DATE% %TIME%
>> "%STATUS%" echo package_label=__SUITE__
>> "%STATUS%" echo run_id=%RUN_ID%
>> "%STATUS%" echo temp_root_kind=ISOLATED_ASCII_LANE
>> "%STATUS%" echo temp_source_kind=%TEMP_SOURCE_KIND%
>> "%STATUS%" echo temp_ascii=1
>> "%STATUS%" echo temp_run_token=%RUN_TOKEN%
>> "%STATUS%" echo temp_path_length=%TEMP_PATH_LENGTH%
> "%OUTPUT%" echo LHexcel orchestrator output
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0payload\Invoke-LHexcelVerify.ps1" -OriginalBundle "%ROOT%" >> "%OUTPUT%" 2>&1
set "PRIMARY_RC=%ERRORLEVEL%"
set "RC=%PRIMARY_RC%"
set "CLEANUP_RC=0"
if "%TEMP_LOCK_OWNED%"=="1" (
  rmdir /s /q "%LANE_TEMP%" 2>nul
  if exist "%LANE_TEMP%" set "CLEANUP_RC=25"
  rmdir /s /q "%TEMP_LOCK%" 2>nul
  if exist "%TEMP_LOCK%" set "CLEANUP_RC=25"
)
if not "%CLEANUP_RC%"=="0" set "RC=25"
if not "%RC%"=="0" del /f /q "%~dp0return\PASS.zip" 2>nul
if "%RC%"=="0" if not exist "%~dp0return\PASS.zip" set "RC=26"
if "%RC%"=="0" if exist "%~dp0return\DIAGNOSTIC.zip" set "RC=26"
if not "%RC%"=="0" del /f /q "%~dp0return\PASS.zip" 2>nul
>> "%STATUS%" echo primary_launcher_exit=%PRIMARY_RC%
>> "%STATUS%" echo launcher_exit=%RC%
if "%RC%"=="0" if exist "%~dp0return\PASS.zip" >> "%STATUS%" echo return_zip=PASS.zip
if "%RC%"=="0" if exist "%~dp0return\PASS.zip" >> "%STATUS%" echo final_state=PASS
if not "%RC%"=="0" if exist "%~dp0return\DIAGNOSTIC.zip" >> "%STATUS%" echo return_zip=DIAGNOSTIC.zip
if not "%RC%"=="0" if exist "%~dp0return\DIAGNOSTIC.zip" >> "%STATUS%" echo final_state=DIAGNOSTIC
if not "%RC%"=="0" if not exist "%~dp0return\DIAGNOSTIC.zip" >> "%STATUS%" echo return_zip=NONE
if not "%RC%"=="0" if not exist "%~dp0return\DIAGNOSTIC.zip" >> "%STATUS%" echo final_state=FAIL_CLOSED_NO_DIAGNOSTIC
>> "%STATUS%" echo wrapper_finished=%DATE% %TIME%
rmdir /s /q "%RUN_LOCK%" 2>nul
echo.
echo LHexcel verification exit code: %RC%
echo Status file: %STATUS%
echo Output file: %OUTPUT%
pause
exit /b %RC%
'''.replace("__SUITE__", suite)


def guide(suite: str, launcher: str, expected: int) -> str:
    if suite in {"ProductRibbon", "ProductUi"}:
        runner = "Run-ProductRibbonSuite.ps1" if suite == "ProductRibbon" else "Run-ProductUiSuite.ps1"
        return f"""# 검증 안내\n\nWindows PowerShell 5.1 이상과 Excel 2024 x64가 필요합니다. 유일한 실행 진입점은 `{launcher}`이며, macOS에서는 BAT/PowerShell/Excel을 실행하지 않고 정적 패키지만 생성합니다. 실행 시 `{runner}`가 PowerShell·Excel COM·64-bit Office 사전 점검, `Build-Xlam.ps1`, VBA 컴파일 및 callback/UI 증거를 순서대로 수행합니다.\n\n성공 시 `return/PASS.zip`, 실패 시 redacted 진단만 담은 `return/DIAGNOSTIC.zip` 중 정확히 하나만 게시합니다.\n"""
    return f"""# 검증 안내

macOS는 static-only package만 만들며 PASS.zip과 DIAGNOSTIC.zip을 만들지 않습니다. 각 lane 폴더의 {launcher} 하나만 실행하면 BAT가 manifest.json의 run-id를 공유 lock으로 원자 예약한 뒤 격리된 ASCII TEMP 루트(`%LOCALAPPDATA%\\LHexcel\\VerifyTemp\\{suite}-<token>`)를 TEMP/TMP로 고정합니다. 같은 run-id가 이미 예약되어 있으면 기존 TEMP를 건드리지 않고 코드 10으로 종료하며 stale lock을 임의 삭제하지 않습니다. 최종 경로가 120자를 넘으면 짧은 ASCII fallback을 시도하고 그래도 넘으면 fail-closed 합니다. Invoke-LHexcelVerify.ps1는 payload SHA256SUMS·manifest digest를 확인한 뒤 Core RedProbe와 Green {suite} {expected}/{expected}을 수행합니다.

Excel COM 검증은 전역 이름 있는 mutex(`Global\\LHexcel-v0.01-Excel-COM-v001`)로 직렬화됩니다. 다른 lane이 Excel COM을 보유 중이면 현재 실행은 fail-closed 코드 10으로 종료되며, COM 동시 실행을 허용하지 않습니다. BAT를 직접 병렬 실행하지 말고 현재 내엑셀 프로젝트의 `tmp/LHexcel-v0.01-parallel/<lane>/` 폴더를 각각 사용합니다.

오케스트레이터는 완성 ZIP을 메모리에 보존하고 TEMP 정리를 검증한 뒤 return/PASS.zip 또는 return/DIAGNOSTIC.zip 하나만 게시합니다. diagnostics에는 redacted 로그만 넣고 raw stdout/stderr는 포함하지 않습니다.

코드: 10 preflight, 20 integrity, 21 stale, 22 build/Green compile, 23 Red, 24 runtime evidence, 25 TEMP/cleanup, 26 return.
"""


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("destination", nargs="?")
    parser.add_argument("--output")
    parser.add_argument("--excel-bitness", choices=("32", "64"))
    parser.add_argument("--suite", choices=tuple(SUITES), default="Core")
    parser.add_argument("--lane", help="write a parallel verification lane under the fixed parallel root")
    parser.add_argument("--list-suites", action="store_true")
    args = parser.parse_args()
    if args.list_suites:
        for name, config in SUITES.items():
            print(f"{name}\t{config['expected']}\t{config['launcher']}")
        return 0
    if args.excel_bitness is None:
        parser.error("--excel-bitness is required")
    if args.destination and args.output:
        parser.error("use either destination or --output")
    if args.lane is not None and not re.fullmatch(r"[A-Za-z0-9_-]+", args.lane):
        parser.error("lane must contain only ASCII letters, digits, '_' or '-'")
    if args.lane is not None and args.output:
        parser.error("use either --lane or --output")
    if args.lane is not None and args.destination:
        parser.error("use either --lane or destination")
    destination = args.output or args.destination
    if args.lane is not None:
        destination = str(PARALLEL_ROOT / args.lane)
    if destination is None:
        parser.error("destination or --output is required")
    suite_config = SUITES[args.suite]
    if suite_config.get("product") and args.excel_bitness != "64":
        parser.error(f"{args.suite} requires 64-bit Excel")
    launcher = suite_config["launcher"]
    dest = Path(destination)
    # Temporary destinations are synthetic contract fixtures; real parallel
    # lanes must pass the clean/tracked source preflight.
    synthetic_destination = str(dest.absolute()).startswith(str(Path(tempfile.gettempdir())) + os.sep)
    source_commit = require_clean_tracked_source() if suite_config.get("product") and not synthetic_destination else commit()
    if dest.exists():
        status_path = dest / "status" / "run-status.txt"
        if status_path.is_file():
            status = status_path.read_text(encoding="utf-8", errors="replace")
            if "wrapper_started=" in status and "wrapper_finished=" not in status:
                raise ValueError("incomplete wrapper run detected; refusing to replace destination (sentinel preserved)")
        thaw_for_removal(dest)
        shutil.rmtree(dest)
    dest.mkdir(parents=True)
    payload = dest / "payload"
    payload_roots = ["src/vba", "tests/vba", "tests/windows", "build", "contracts"]
    if suite_config.get("product"):
        payload_roots += ["src/ribbon", "src/python", "src/resources", "tools", "provenance"]
    if suite_config["extension_manifest"] is not None:
        payload_roots.append("tests/manifests")
    for relative in payload_roots:
        # Host-specific historical recovery is not part of a portable product probe.
        ignored = ["__pycache__", "*.pyc", ".DS_Store"]
        if relative == "tools":
            ignored.extend(("Repair-R71BlockedTestRegistration.ps1", "test_r73_setup.ps1"))
        shutil.copytree(ROOT / relative, payload / relative, ignore=shutil.ignore_patterns(*ignored))
    if suite_config["extension_manifest"] is not None:
        validate_extension_graph(payload, suite_config["extension_manifest"], args.suite, args.suite.lower())
    write(payload / "Invoke-LHexcelVerify.ps1", orchestrator(args.suite, launcher, suite_config["expected"], suite_config["extension_manifest"]), bom=True)
    write(dest / launcher, batch(args.suite))
    write(dest / "검증_안내.md", guide(args.suite, launcher, suite_config["expected"]), bom=True)
    clear_hidden_flags(dest)
    base17_records = [
        {"path": f"payload/src/vba/core/{name}", "size": (payload / "src/vba/core" / name).stat().st_size, "sha256": sha(payload / "src/vba/core" / name)}
        for name in CORE
    ] + [
        {"path": f"payload/tests/vba/{name}", "size": (payload / "tests/vba" / name).stat().st_size, "sha256": sha(payload / "tests/vba" / name)}
        for name in TESTS
    ]
    entrypoint_path = payload / "tests/windows/feature-suite-entrypoints.json"
    entrypoint_records = [{"path": "payload/tests/windows/feature-suite-entrypoints.json", "size": entrypoint_path.stat().st_size, "sha256": sha(entrypoint_path)}]
    base = [(record["path"], record["sha256"]) for record in base17_records]
    entry = [(record["path"], record["sha256"]) for record in entrypoint_records]
    launcher_digest = sha(dest / launcher)
    guide_digest = sha(dest / "검증_안내.md")
    packager_digest = sha(Path(__file__))
    package = digest_pairs([("launcher", launcher_digest), ("guide", guide_digest), ("packager", packager_digest)])
    manifest = {
        "schema_version": 1, "run_id": str(uuid.uuid4()),
        "created_utc": datetime.now(UTC).isoformat(timespec="seconds").replace("+00:00", "Z"),
        "approved_commit": APPROVED_COMMIT, "packaging_commit": source_commit, "suite": args.suite,
        "requested_excel_bitness": args.excel_bitness, "base17_records": base17_records,
        "extension_manifest": suite_config["extension_manifest"],
        "extension_manifest_digest": None if suite_config["extension_manifest"] is None else sha(payload / suite_config["extension_manifest"]),
        "base17_digest": digest_pairs(base), "entrypoint_records": entrypoint_records,
        "entrypoint_digest": digest_pairs(entry), "packager_code_digest": packager_digest,
        "launcher_digest": launcher_digest, "guide_digest": guide_digest, "packager_digest": package,
        "staged_data_records": [], "staged_data_digest": "",
        "file_records": [], "snapshot_digest": "", "sha256sums_digest": "",
    }
    if suite_config.get("product"):
        manifest["source_tree_sha256"] = product_source_tree_digest()
    file_paths = [
        path
        for path in sorted(dest.rglob("*"), key=lambda candidate: candidate.relative_to(dest).as_posix())
        if path.is_file()
        and path.name not in {"SHA256SUMS", "manifest.json"}
        and path.name != ".DS_Store"
    ]
    manifest["file_records"] = [
        {"path": path.relative_to(dest).as_posix(), "size": path.stat().st_size, "sha256": sha(path)}
        for path in file_paths
    ]
    data_prefixes = ("payload/src/", "payload/tests/vba/", "payload/contracts/")
    if suite_config.get("product"):
        data_prefixes += ("payload/src/ribbon/", "payload/src/python/", "payload/tools/", "payload/provenance/")
    if suite_config["extension_manifest"] is not None:
        data_prefixes += ("payload/tests/manifests/",)
    manifest["staged_data_records"] = [
        record for record in manifest["file_records"]
        if record["path"].startswith(data_prefixes)
    ]
    for record in manifest["staged_data_records"]:
        if record["path"].lower().endswith((".ps1", ".bat", ".cmd", ".exe", ".dll")) and record["path"] != ALLOWED_RUNTIME_SCRIPT:
            raise ValueError(f"executable staged data rejected: {record['path']}")
    manifest["staged_data_digest"] = digest_pairs([
        (record["path"], record["sha256"])
        for record in manifest["staged_data_records"]
    ])
    manifest["snapshot_digest"] = digest_pairs([(record["path"], record["sha256"]) for record in manifest["file_records"]])
    sums = [f"{record['sha256']}  {record['path']}" for record in manifest["file_records"]]
    checksum_payload = "\n".join(sums) + "\n"
    manifest["sha256sums_digest"] = hashlib.sha256(checksum_payload.encode()).hexdigest()
    write(dest / "manifest.json", json.dumps(manifest, sort_keys=True, separators=(",", ":")) + "\n")
    write(dest / "SHA256SUMS", checksum_payload)
    make_read_only(payload)
    print(f"static_bundle={dest}\nstatus=STATIC_PACKAGE_ONLY")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
