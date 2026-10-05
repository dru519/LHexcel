param(
    [string]$Suite = 'Core',
    [string]$Filter = '',
    [ValidateSet('RedProbe','Green')][string]$Mode = 'Green',
    [string]$ExpectedCompileFailure = '',
    [string]$ExtensionManifest = '',
    [string]$RunId = '',
    [string]$SourceDigest = '',
    [string]$SnapshotDigest = '',
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [string]$EvidenceRoot = '',
    [switch]$OwnerCompile,
    [string]$OwnerStartup = '',
    [string]$OwnerExcelIdentity = '',
    [string]$OwnerArtifact = '',
    [string]$OwnerHandshake = '',
    [string]$OwnerGo = '',
    [string]$OwnerResult = '',
    [switch]$OwnerCase,
    [string]$OwnerCaseRequest = ''
)

function Release-ComObject([object]$Value) { if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } }
function Find-CommandBarControlById([object]$Controls, [int]$ControlId, [int]$Depth) {
    if ($null -eq $Controls -or $Depth -gt 6) { return $null }
    $count = 0
    try { $count = [int]$Controls.Count } catch { return $null }
    for ($index = 1; $index -le $count; $index++) {
        $candidate = $null; $children = $null; $found = $null
        try {
            $candidate = $Controls.Item($index)
            if ([int]$candidate.Id -eq $ControlId) { $result = $candidate; $candidate = $null; return $result }
            try { $children = $candidate.Controls } catch { $children = $null }
            if ($null -ne $children -and [int]$children.Count -gt 0) {
                $found = Find-CommandBarControlById $children $ControlId ($Depth + 1)
                if ($null -ne $found) { $result = $found; $found = $null; return $result }
            }
        } finally {
            Release-ComObject $found; Release-ComObject $children; Release-ComObject $candidate
        }
    }
    return $null
}
function Find-VbeCompileControl([object]$Bars) {
    $barCount = 0
    try { $barCount = [int]$Bars.Count } catch { return $null }
    for ($barIndex = 1; $barIndex -le $barCount; $barIndex++) {
        $bar = $null; $controls = $null; $found = $null
        try {
            $bar = $Bars.Item($barIndex); $controls = $bar.Controls
            $found = Find-CommandBarControlById $controls 578 0
            if ($null -ne $found) { $result = $found; $found = $null; return $result }
        } finally {
            Release-ComObject $found; Release-ComObject $controls; Release-ComObject $bar
        }
    }
    return $null
}
Set-StrictMode -Version Latest
if ($PSVersionTable.PSVersion.Major -lt 5 -or ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1)) { throw 'Windows PowerShell 5.1 or later is required' }
function Get-Sha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant() }
function Get-StringSha256([string]$Text) { $sha = [Security.Cryptography.SHA256]::Create(); try { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant() } finally { $sha.Dispose() } }
function Get-RangeFormulaForDigest([object]$Range) {
    try { return [string]$Range.PSObject.Properties['Formula2'].Value } catch {}
    try { return [string]$Range.Formula } catch { return '<formula-unreadable>' }
}
function ConvertTo-ValidUnicode([string]$Text) { if($null -eq $Text){return $null};$utf8=New-Object Text.UTF8Encoding($false,$false);$utf8.GetString($utf8.GetBytes($Text)) }
function Write-AtomicEvidence([string]$Path, [object]$Evidence) {
    $directory = Split-Path -Parent $Path; [void](New-Item -ItemType Directory -Path $directory -Force)
    $temporary = Join-Path $directory (([IO.Path]::GetFileName($Path)) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $backup = Join-Path $directory (([IO.Path]::GetFileName($Path)) + '.' + [Guid]::NewGuid().ToString('N') + '.bak')
    $writer = New-Object IO.StreamWriter($temporary, $false, (New-Object Text.UTF8Encoding($false)))
    try { $writer.Write(($Evidence | ConvertTo-Json -Depth 16 -Compress)); $writer.Flush(); $writer.BaseStream.Flush($true) } finally { $writer.Close() }
    try {
        if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($temporary, $Path, $backup) } else { [IO.File]::Move($temporary, $Path) }
    } finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $backup) { Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue }
    }
    if (Test-Path -LiteralPath $temporary) { throw 'atomic evidence temporary residue' }
    if (Test-Path -LiteralPath $backup) { throw 'atomic evidence backup residue' }
}
function Get-BoundedDiagnosticText([string]$Path, [string[]]$SensitiveValues) {
    if([string]::IsNullOrWhiteSpace($Path) -or -not(Test-Path -LiteralPath $Path -PathType Leaf)){return $null}
    try {
        $text=[IO.File]::ReadAllText($Path,[Text.Encoding]::UTF8)
        foreach($value in $SensitiveValues){if(-not [string]::IsNullOrWhiteSpace($value)){$text=$text.Replace($value,'<path>')}}
        $text=[regex]::Replace($text,'[\u0000-\u0008\u000B\u000C\u000E-\u001F]',' ')
        if($text.Length -gt 2048){$text=$text.Substring($text.Length-2048)}
        return ConvertTo-ValidUnicode $text
    } catch { return $null }
}
function Assert-ExactObjectProperties([object]$Value, [string[]]$Expected, [string]$Context) {
    if ($null -eq $Value) { throw ($Context + ' is missing') }
    $actual = @($Value.PSObject.Properties.Name | Sort-Object)
    $wanted = @($Expected | Sort-Object)
    if (($actual -join '|') -ne ($wanted -join '|')) { throw ($Context + ' property set rejected') }
}
function Assert-UserFormBuildEvidence([object]$TestHost,[string]$RunIdentity,[string]$ExpectedSourceDigest,[string]$ExpectedSnapshotDigest,[string]$Root,[string]$SuiteName,[string]$Owner,[string]$ManifestPath) {
    # Evidence names remain stable even when an owner intentionally has no UserForm.
    if($SuiteName -notin @('AI','Template','Data','Draw','File','Calculator','Symbols')){return $null}
    $expectedEvidenceName=$SuiteName+'.UserFormBuild.json'
    if($null -eq $TestHost -or [string]::IsNullOrWhiteSpace([string]$TestHost.BuildEvidencePath) -or [IO.Path]::GetFileName([string]$TestHost.BuildEvidencePath) -ne $expectedEvidenceName -or -not(Test-Path -LiteralPath $TestHost.BuildEvidencePath -PathType Leaf)){throw 'UserForm build evidence missing'}
    $document=[IO.File]::ReadAllText([string]$TestHost.BuildEvidencePath,[Text.Encoding]::UTF8)|ConvertFrom-Json
    Assert-ExactObjectProperties $document @('schema_version','run_id','suite','owner','mode','source_digest','snapshot_digest','extension_manifest_path','extension_manifest_sha256','artifact_path','artifact_sha256','forms','created_utc') 'UserFormBuildEvidence'
    if($document.schema_version -is [bool] -or -not($document.schema_version -is [int]) -or $document.schema_version -ne 1 -or $document.run_id -ne $RunIdentity -or $document.source_digest -ne $ExpectedSourceDigest -or $document.snapshot_digest -ne $ExpectedSnapshotDigest -or $document.suite -ne $SuiteName -or $document.owner -ne $Owner -or $document.mode -ne 'Green'){throw 'UserForm build evidence stale binding'}
    $manifestFull=[IO.Path]::GetFullPath($ManifestPath)
    $rootPrefix=[IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar,[IO.Path]::AltDirectorySeparatorChar)+[IO.Path]::DirectorySeparatorChar
    $expectedManifestRelative=$manifestFull.Substring($rootPrefix.Length).Replace('\','/')
    if($document.extension_manifest_path -ne $expectedManifestRelative -or $document.extension_manifest_sha256 -ne (Get-Sha256 $manifestFull)){throw 'UserForm build evidence manifest binding rejected'}
    if([IO.Path]::GetFullPath([string]$document.artifact_path) -ne [IO.Path]::GetFullPath([string]$TestHost.ArtifactPath) -or $document.artifact_sha256 -ne $TestHost.HostSha256 -or (Get-Sha256 $TestHost.ArtifactPath) -ne $document.artifact_sha256){throw 'UserForm build evidence artifact binding rejected'}
    $expectedFormName=if($SuiteName -eq 'AI'){'FNxAi'}elseif($SuiteName -eq 'Template'){'FNxTemplate'}elseif($SuiteName -eq 'Data'){'FNxData'}elseif($SuiteName -eq 'Draw'){'FNxDraw'}elseif($SuiteName -eq 'Calculator'){'FNxCalculator'}elseif($SuiteName -eq 'Symbols'){'FNxSymbols'}else{'FNxFile'}
    $forms=@($document.forms|Where-Object {$_.name -eq $expectedFormName})
    if($forms.Count -ne 1){throw ('UserForm build evidence '+$expectedFormName+' exact form rejected')}
    $form=$forms[0]
    Assert-ExactObjectProperties $form @('name','layout_path','layout_sha256','code_path','code_sha256','controls_created','default_control','cancel_control','comboboxes','listboxes') 'UserForm build evidence form'
    if($SuiteName -eq 'AI'){
        if($form.layout_path -ne 'src/vba/features/ai/ui/FNxAi.form.json' -or $form.code_path -ne 'src/vba/features/ai/ui/FNxAi.vba' -or [int]$form.controls_created -ne 24 -or $form.default_control -ne 'cmdExecute' -or $form.cancel_control -ne 'cmdCancel'){throw 'UserForm build evidence FNxAi contract rejected'}
    }elseif($SuiteName -eq 'Template'){
        if($form.layout_path -ne 'src/vba/features/template/ui/FNxTemplate.form.json' -or $form.code_path -ne 'src/vba/features/template/ui/FNxTemplate.vba' -or [int]$form.controls_created -ne 19 -or $form.default_control -ne 'cmdPreview' -or $form.cancel_control -ne 'cmdCancel'){throw 'UserForm build evidence FNxTemplate contract rejected'}
    }elseif($SuiteName -eq 'Data'){
        if($form.layout_path -ne 'src/vba/features/data/ui/FNxData.form.json' -or $form.code_path -ne 'src/vba/features/data/ui/FNxData.vba' -or [int]$form.controls_created -ne 14 -or $form.default_control -ne 'cmdPreview' -or $form.cancel_control -ne 'cmdCancel'){throw 'UserForm build evidence FNxData contract rejected'}
    }elseif($SuiteName -eq 'Draw'){
        if($form.layout_path -ne 'src/vba/features/draw/ui/FNxDraw.form.json' -or $form.code_path -ne 'src/vba/features/draw/ui/FNxDraw.vba' -or [int]$form.controls_created -ne 14 -or $form.default_control -ne 'cmdPreview' -or $form.cancel_control -ne 'cmdCancel'){throw 'UserForm build evidence FNxDraw contract rejected'}
    }elseif($SuiteName -eq 'Calculator'){
        if($form.layout_path -ne 'src/vba/features/calculator/ui/FNxCalculator.form.json' -or $form.code_path -ne 'src/vba/features/calculator/ui/FNxCalculator.vba' -or [int]$form.controls_created -ne 39 -or $form.default_control -ne 'cmdEquals' -or $null -ne $form.cancel_control){throw 'UserForm build evidence FNxCalculator contract rejected'}
    }elseif($SuiteName -eq 'Symbols'){
        if($form.layout_path -ne 'src/vba/features/symbols/ui/FNxSymbols.form.json' -or $form.code_path -ne 'src/vba/features/symbols/ui/FNxSymbols.vba' -or [int]$form.controls_created -ne 10 -or $form.default_control -ne 'cmdInsertCell' -or $form.cancel_control -ne 'cmdClose'){throw 'UserForm build evidence FNxSymbols contract rejected'}
    }else{
        if($form.layout_path -ne 'src/vba/features/file/ui/FNxFile.form.json' -or $form.code_path -ne 'src/vba/features/file/ui/FNxFile.vba' -or [int]$form.controls_created -ne 15 -or $form.default_control -ne 'cmdPreview' -or $form.cancel_control -ne 'cmdCancel'){throw 'UserForm build evidence FNxFile contract rejected'}
    }
    if($form.layout_sha256 -ne (Get-Sha256 (Join-Path $Root $form.layout_path)) -or $form.code_sha256 -ne (Get-Sha256 (Join-Path $Root $form.code_path))){throw 'UserForm build evidence source binding rejected'}
    $expectedCombos=if($SuiteName -eq 'AI'){@(
        [pscustomobject]@{name='cboMode';items=@('요약/분석','데이터 정리','수식 도우미','문서 작성','이미지 활용')},
        [pscustomobject]@{name='cboOutput';items=@('간결한 요약','표 또는 목록','단계별 설명','업무 문안','사용자 지정')}
    )}elseif($SuiteName -eq 'Template'){@(
        [pscustomobject]@{name='cboSourceKind';items=@('선택 범위','현재 시트 사용 영역')},
        [pscustomobject]@{name='cboHiddenPolicy';items=@('숨김 행·열 제외','숨김 행·열 포함')},
        [pscustomobject]@{name='cboRiskConsent';items=@('동의하지 않음','위험을 확인하고 동의함')},
        [pscustomobject]@{name='cboTemplates';items=@('등록된 템플릿 없음')}
    )}elseif($SuiteName -eq 'Data'){@(
        [pscustomobject]@{name='cboFeature';items=@('고유값 카운팅','중복 리스트')},
        [pscustomobject]@{name='cboHeader';items=@('제목행 있음','제목행 없음')},
        [pscustomobject]@{name='cboKeyColumn';items=@('첫 번째 열')},
        [pscustomobject]@{name='cboTrim';items=@('앞뒤 공백 유지','앞뒤 공백 제거')},
        [pscustomobject]@{name='cboCase';items=@('대소문자 구분 안 함','대소문자 구분')},
        [pscustomobject]@{name='cboHidden';items=@('숨김·필터 행 제외','숨김·필터 행 포함')}
    )}elseif($SuiteName -eq 'Draw'){@(
        [pscustomobject]@{name='cboFeature';items=@('제목표','실무표','내부선 제거','외곽선','그림 셀 맞춤')}
    )}elseif($SuiteName -in @('Calculator','Symbols')){@()}else{@(
        [pscustomobject]@{name='cboFeature';items=@('여러 파일 통합','매너 저장','선택 범위 PNG 저장','차트 PNG 저장')}
    )}
    # PowerShell 5.1 unwraps a singleton conditional result; normalize both singleton and multi-item branches.
    $expectedCombos=@($expectedCombos)
    $combos=@($form.comboboxes)
    if($combos.Count -ne $expectedCombos.Count){throw 'UserForm build evidence ComboBox count rejected'}
    $separator=[string][char]31
    for($index=0;$index -lt $expectedCombos.Count;$index++){
        $combo=$combos[$index];$expected=$expectedCombos[$index]
        Assert-ExactObjectProperties $combo @('name','prog_id','items','selected_index','style_value','match_required','set_sequence') 'UserForm build evidence ComboBox'
        if($combo.name -ne $expected.name -or $combo.prog_id -ne 'Forms.ComboBox.1' -or (@($combo.items)-join $separator) -ne (@($expected.items)-join $separator) -or [int]$combo.selected_index -ne 0 -or [int]$combo.style_value -ne 2 -or $combo.match_required -ne $true -or (@($combo.set_sequence)-join '|') -ne 'style|match_required|items|selected_index'){throw 'UserForm build evidence ComboBox exact values rejected'}
    }
    $expectedLists=if($SuiteName -eq 'File'){@(
        [pscustomobject]@{name='lstInputFiles';items=@();selected_index=-1;multi_select=$false;multi_select_value=0;set_sequence=@('multi_select')},
        [pscustomobject]@{name='lstDecisions';items=@();selected_index=-1;multi_select=$false;multi_select_value=0;set_sequence=@()}
    )}elseif($SuiteName -eq 'Symbols'){@(
        [pscustomobject]@{name='lstCategories';items=@();selected_index=-1;multi_select=$false;multi_select_value=0;set_sequence=@('multi_select')}
    )}else{@()}
    $expectedLists=@($expectedLists)
    $lists=@($form.listboxes)
    if($lists.Count -ne $expectedLists.Count){throw 'UserForm build evidence ListBox count rejected'}
    for($index=0;$index -lt $expectedLists.Count;$index++){
        $list=$lists[$index];$expected=$expectedLists[$index]
        Assert-ExactObjectProperties $list @('name','prog_id','items','selected_index','multi_select','multi_select_value','set_sequence') 'UserForm build evidence ListBox'
        if($list.name -ne $expected.name -or $list.prog_id -ne 'Forms.ListBox.1' -or (@($list.items)-join $separator) -ne (@($expected.items)-join $separator) -or [int]$list.selected_index -ne [int]$expected.selected_index -or [bool]$list.multi_select -ne [bool]$expected.multi_select -or [int]$list.multi_select_value -ne [int]$expected.multi_select_value -or (@($list.set_sequence)-join '|') -ne (@($expected.set_sequence)-join '|')){throw 'UserForm build evidence ListBox exact values rejected'}
    }
    $evidenceSha256=Get-Sha256 $TestHost.BuildEvidencePath
    if($evidenceSha256 -ne [string]$TestHost.BuildEvidenceSha256){throw 'UserForm build evidence SHA-256 binding rejected'}
    return [pscustomobject]@{status='VERIFIED_EXACT_V1';run_id=$RunIdentity;build_evidence_sha256=$evidenceSha256}
}
function Initialize-NativeWindowApi {
    if (-not ('NxTask3.NativeWindowApi' -as [type])) {
        Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
namespace NxTask3 { public static class NativeWindowApi {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int max);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr w, IntPtr l);
  public static string[] DialogsForPid(uint wanted) { var result=new List<string>(); EnumWindows((h,l)=>{uint pid; GetWindowThreadProcessId(h,out pid); if(pid==wanted && IsWindowVisible(h)){var b=new StringBuilder(2048); GetWindowText(h,b,b.Capacity); if(b.Length>0) result.Add(b.ToString());} return true;},IntPtr.Zero); return result.ToArray(); }
}}
'@
    }
}
function Get-ExcelPid([object]$Excel) { Initialize-NativeWindowApi; [uint32]$excelProcessId = 0; [void][NxTask3.NativeWindowApi]::GetWindowThreadProcessId([IntPtr]$Excel.Hwnd, [ref]$excelProcessId); return $excelProcessId }
. (Join-Path $PSScriptRoot '../../build/Excel-ProcessLifecycle.ps1')
function Start-CompileDialogWatcher([uint32]$ExcelPid, [string]$RunId, [Int64]$ExcelHwnd, [Int64]$VbeHwnd) {
    # A genuinely concurrent owner process watches visible own-PID windows while the parent owns Excel COM.
    Start-Job -ArgumentList $ExcelPid,$RunId,$ExcelHwnd,$VbeHwnd -ScriptBlock {
        param($excelProcessId,$runId,$excelHwnd,$vbeHwnd)
        function Get-BodySha256([string]$Body) { $sha=[Security.Cryptography.SHA256]::Create(); try { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Body)))).Replace('-','').ToLowerInvariant() } finally { $sha.Dispose() } }
        Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
        $uiaClientPath=[System.Windows.Automation.AutomationElement].Assembly.Location
        $uiaTypesPath=[System.Windows.Automation.AutomationPattern].Assembly.Location
        $compileErrors=@()
        Add-Type -ReferencedAssemblies @($uiaClientPath,$uiaTypesPath) -ErrorVariable compileErrors -ErrorAction SilentlyContinue -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices; using System.Windows.Automation;
public static class NxDialogWatcher {
  [UnmanagedFunctionPointer(CallingConvention.Winapi)] public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc p,IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h,EnumWindowsProc p,IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h,StringBuilder b,int n);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h,StringBuilder b,int n);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  static void Add(List<string> r,string value){if(!String.IsNullOrWhiteSpace(value))r.Add(value);}
  static string HwndText(IntPtr h){var b=new StringBuilder(4096);GetWindowText(h,b,b.Capacity);return b.ToString();}
  static string Body(IntPtr h){var parts=new List<string>();Add(parts,HwndText(h));try{var root=AutomationElement.FromHandle(h);Action<AutomationElement> add=(e)=>{if(e==null)return;Add(parts,e.Current.Name);object p;if(e.TryGetCurrentPattern(TextPattern.Pattern,out p))Add(parts,((TextPattern)p).DocumentRange.GetText(-1));if(e.TryGetCurrentPattern(ValuePattern.Pattern,out p))Add(parts,((ValuePattern)p).Current.Value);};add(root);foreach(AutomationElement e in root.FindAll(TreeScope.Descendants,Condition.TrueCondition))add(e);}catch{}EnumChildWindows(h,(child,l)=>{Add(parts,HwndText(child));return true;},IntPtr.Zero);return String.Join("\n",parts.ToArray());}
  public static string[] PollOwnDialogs(uint wanted,long excelRoot,long vbeRoot){var r=new List<string>();EnumWindows((h,l)=>{uint pid;GetWindowThreadProcessId(h,out pid);var c=new StringBuilder(256);GetClassName(h,c,c.Capacity);if(pid==wanted&&IsWindowVisible(h)&&h.ToInt64()!=excelRoot&&h.ToInt64()!=vbeRoot){var body=Body(h);if(!String.IsNullOrWhiteSpace(body))r.Add(h.ToInt64()+"|"+c.ToString()+"|"+body);}return true;},IntPtr.Zero);return r.ToArray();}
  public static bool VerifyDialogDisappeared(long h){return !IsWindow(new IntPtr(h));}
  static bool InvokeClose(AutomationElement root){foreach(AutomationElement e in root.FindAll(TreeScope.Descendants,Condition.TrueCondition)){object p;if(e.Current.ControlType==ControlType.Button&&e.TryGetCurrentPattern(InvokePattern.Pattern,out p)){((InvokePattern)p).Invoke();return true;}}return false;}
  public static bool CloseOnlyOwnDialog(long h,uint wanted){uint pid;var handle=new IntPtr(h);GetWindowThreadProcessId(handle,out pid);if(pid!=wanted)return false;try{var root=AutomationElement.FromHandle(handle);InvokeClose(root);}catch{}for(int i=0;i<10;i++){if(VerifyDialogDisappeared(h))return true;System.Threading.Thread.Sleep(50);}PostMessage(handle,0x0010,IntPtr.Zero,IntPtr.Zero);for(int i=0;i<10;i++){if(VerifyDialogDisappeared(h))return true;System.Threading.Thread.Sleep(50);}return false;}
}
'@
        if(-not ('NxDialogWatcher' -as [type])){
            $localCompilerErrors=@()
            foreach($compileError in $compileErrors){
                $target=$compileError.TargetObject
                $diagnosticText=([string]$compileError.Exception.Message)+'|'+([string]$target)+'|'+([string]($compileError|Out-String))
                $errorNumber=if($null -ne $target -and $null -ne $target.PSObject.Properties['ErrorNumber']){[string]$target.ErrorNumber}else{([regex]::Match($diagnosticText,'(?i)(CS\d+)')).Groups[1].Value.ToUpperInvariant()}
                $errorLine=if($null -ne $target -and $null -ne $target.PSObject.Properties['Line']){[int]$target.Line}else{0}
                $errorColumn=if($null -ne $target -and $null -ne $target.PSObject.Properties['Column']){[int]$target.Column}else{0}
                $errorText=if($null -ne $target -and $null -ne $target.PSObject.Properties['ErrorText']){[string]$target.ErrorText}else{$diagnosticText}
                $localCompilerErrors += [pscustomobject]@{error_number=$errorNumber;line=$errorLine;column=$errorColumn;target_type=if($null -ne $target){$target.GetType().FullName}else{$null};error_text_sha256=(Get-BodySha256 $errorText)}
            }
            [pscustomobject]@{run_id=$runId;status='WATCHER_COMPILE_ERROR';class=$null;body=$null;body_sha256=$null;closed=$false;stable_polls=0;hwnd=0;owner_pid=$excelProcessId;disappeared=$false;candidate_count=0;max_candidate_count=0;candidates=@();watcher_error=[pscustomobject]@{count=$compileErrors.Count;compiler_errors=$localCompilerErrors}}
            return
        }
        $found = @(); $previousBody = $null; $stablePolls = 0; $maxCandidateCount = 0; $lastCandidates = @()
        # Accept exactly one own-PID dialog only after two equal nonempty polls; close that hwnd and verify disappearance.
        for($i=0;$i -lt 40;$i++){
            $poll=@([NxDialogWatcher]::PollOwnDialogs([uint32]$excelProcessId,$excelHwnd,$vbeHwnd))
            if($poll.Count -gt $maxCandidateCount){$maxCandidateCount=$poll.Count}
            $currentCandidates=@($poll|ForEach-Object{$candidateParts=$_.Split('|',3);$candidateBody=$candidateParts[2];[pscustomobject]@{hwnd=[Int64]$candidateParts[0];class=$candidateParts[1];body_length=$candidateBody.Length;body_sha256=(Get-BodySha256 $candidateBody)}})
            if($currentCandidates.Count -gt 0){$lastCandidates=$currentCandidates}
            if($poll.Count -eq 1){
                $parts=$poll[0].Split('|',3);$class=$parts[1];$body=$parts[2]
                if(-not [string]::IsNullOrWhiteSpace($body) -and $body -eq $previousBody){$stablePolls=$stablePolls+1}else{$stablePolls=1}
                $previousBody=$body
                if($stablePolls -ge 2){
                    $closed=[NxDialogWatcher]::CloseOnlyOwnDialog([Int64]$parts[0],[uint32]$excelProcessId)
                    $found += [pscustomobject]@{run_id=$runId;status='CAPTURED';class=$class;body=$body;body_sha256=(Get-BodySha256 $body);closed=$closed;stable_polls=$stablePolls;hwnd=[Int64]$parts[0];owner_pid=$excelProcessId;disappeared=([NxDialogWatcher]::VerifyDialogDisappeared([Int64]$parts[0]));candidate_count=$poll.Count;max_candidate_count=$maxCandidateCount;candidates=$lastCandidates;watcher_error=$null}
                    break
                }
            }else{$stablePolls=0;$previousBody=$null}
            Start-Sleep -Milliseconds 100
        }
        if($found.Count -eq 0){$found += [pscustomobject]@{run_id=$runId;status='NO_STABLE_DIALOG';class=$null;body=$null;body_sha256=$null;closed=$false;stable_polls=$stablePolls;hwnd=0;owner_pid=$excelProcessId;disappeared=$false;candidate_count=$lastCandidates.Count;max_candidate_count=$maxCandidateCount;candidates=$lastCandidates;watcher_error=$null}}
        $found
    }
}
function Invoke-CompileWithWatcher([object]$Excel, [object]$Control, [uint32]$ExcelPid, [Int64]$VbeHwnd, [string]$RunId) {
    Initialize-NativeWindowApi
    $dialog = $null; $started = [DateTime]::UtcNow.ToString('o'); $run_id = if([string]::IsNullOrWhiteSpace($RunId)){[Guid]::NewGuid().ToString('N')}else{$RunId}; $executeError = $null; $paneError = $null; $endLine = 0; $endColumn = 0; $stablePolls = 0; $watcher = Start-CompileDialogWatcher $ExcelPid $run_id ([Int64]$Excel.Hwnd) $VbeHwnd
    try { $Control.Execute() } catch { $executeError = $_.Exception.Message }
    [void](Wait-Job -Job $watcher); $watcherJobState=[string]$watcher.State; $watcherReceiveErrors=@()
    $watcherBody = @(Receive-Job -Job $watcher -Wait -AutoRemoveJob -ErrorVariable watcherReceiveErrors -ErrorAction SilentlyContinue | Select-Object -First 1)
    $watcherError = if($watcherBody.Count -eq 1 -and $null -ne $watcherBody[0].watcher_error){$watcherBody[0].watcher_error}elseif($watcherReceiveErrors.Count -gt 0){
        $fingerprints=@($watcherReceiveErrors|Group-Object FullyQualifiedErrorId|ForEach-Object{
            $diagnosticParts=@()
            foreach($errorRecord in $_.Group){
                $diagnosticParts += [string]$errorRecord.Exception.Message
                if($null -ne $errorRecord.ErrorDetails){$diagnosticParts += [string]$errorRecord.ErrorDetails.Message}
                if($null -ne $errorRecord.TargetObject){$diagnosticParts += [string]$errorRecord.TargetObject}
                $diagnosticParts += [string]($errorRecord|Out-String)
            }
            $diagnosticText=$diagnosticParts -join '|'
            $compilerCodes=@([regex]::Matches($diagnosticText,'(?i)(CS\d+)')|ForEach-Object{$_.Groups[1].Value.ToUpperInvariant()}|Sort-Object -Unique)
            $sourcePositions=@([regex]::Matches($diagnosticText,'\((\d+)\s*,\s*(\d+)\)')|ForEach-Object{$_.Groups[1].Value+':'+$_.Groups[2].Value}|Sort-Object -Unique)
            [pscustomobject]@{error_id=$_.Name;count=$_.Count;exception_types=@($_.Group|ForEach-Object{$_.Exception.GetType().FullName}|Sort-Object -Unique);compiler_codes=$compilerCodes;source_positions=$sourcePositions;message_sha256=(Get-StringSha256 $diagnosticText)}
        })
        [pscustomobject]@{count=$watcherReceiveErrors.Count;fingerprints=$fingerprints}
    }else{[pscustomobject]@{count=0;status='job_returned_no_record'}}
    for ($index = 0; $index -lt 40; $index++) {
        $dialogs = [NxTask3.NativeWindowApi]::DialogsForPid($ExcelPid)
        if ($dialogs.Count -gt 0) { $dialog = $dialogs[0]; break }
        Start-Sleep -Milliseconds 100
    }
    $pane = $null; $module = $null; $moduleName = $null; $line = 0; $sourceLine = $null; $caretColumn = 0
    try { $pane = $Excel.VBE.ActiveCodePane; if ($pane) { $module = $pane.CodeModule; $moduleName = $module.Name; $pane.GetSelection([ref]$line, [ref]$caretColumn, [ref]$endLine, [ref]$endColumn); $sourceLine = $module.Lines($line, 1) } } catch { $paneError = $_.Exception.Message } finally { Release-ComObject $module; Release-ComObject $pane }
    if($watcherBody.Count -gt 0){$dialog=ConvertTo-ValidUnicode ([string]$watcherBody[0].body);$stablePolls=$watcherBody[0].stable_polls}
    # UIA TextPattern/ValuePattern child-body capture requires two consecutive identical nonempty polls for exactly one own-PID dialog.
    [pscustomobject]@{ run_id=$run_id; control_id = 578; watcher_started_utc = $started; dialog_text = $dialog; dialog_sha256 = if ($null -ne $dialog) { Get-StringSha256 $dialog } else { $null }; dialog_closed = if($watcherBody.Count -eq 1){$watcherBody[0].closed}else{$false}; dialog_disappeared = if($watcherBody.Count -eq 1){$watcherBody[0].disappeared}else{$false}; stable_polls = $stablePolls; dialog_owner_pid = $ExcelPid; watcher_status=if($watcherBody.Count -eq 1){$watcherBody[0].status}else{'JOB_NO_RESULT'}; watcher_job_state=$watcherJobState; watcher_error=$watcherError; candidate_count=if($watcherBody.Count -eq 1){$watcherBody[0].candidate_count}else{0}; max_candidate_count=if($watcherBody.Count -eq 1){$watcherBody[0].max_candidate_count}else{0}; candidates=if($watcherBody.Count -eq 1){@($watcherBody[0].candidates)}else{@()}; active_module = $moduleName; line = $line; caret_column = $caretColumn; source_line_sha256 = if ($null -ne $sourceLine) { Get-StringSha256 $sourceLine } else { $null }; execute_error = $executeError; pane_error = $paneError }
}
function Test-Preflight([object]$Excel, [object]$Book) {
    $items = @()
    $items += [pscustomobject]@{ name = 'interactive'; passed = [Environment]::UserInteractive }
    $items += [pscustomobject]@{ name = 'excel_com'; passed = ($null -ne $Excel) }
    try { $security = $Excel.AutomationSecurity; $items += [pscustomobject]@{ name = 'macro_policy'; passed = $true; value = $security } } catch { $items += [pscustomobject]@{ name = 'macro_policy'; passed = $false; value = $_.Exception.Message } }
    $project = $null
    try { $project = $Book.VBProject; $items += [pscustomobject]@{ name = 'VBOM'; passed = $true } } catch { $items += [pscustomobject]@{ name = 'VBOM'; passed = $false; value = $_.Exception.Message } } finally { Release-ComObject $project }
    $vbe = $null; $bars = $null; $control = $null
    try { $vbe = $Excel.VBE; $bars = $vbe.CommandBars; $control = Find-VbeCompileControl $bars; $items += [pscustomobject]@{ name = 'VBE_compile_578'; passed = ($null -ne $control) } } catch { $items += [pscustomobject]@{ name = 'VBE_compile_578'; passed = $false; value = $_.Exception.Message } } finally { Release-ComObject $control; Release-ComObject $bars; Release-ComObject $vbe }
    return $items
}
function Get-ExpectedRedLocation([string]$Root) {
    $path = Join-Path $Root 'tests/vba/T_Core.bas'; $match = Select-String -LiteralPath $path -Pattern 'CNxStateGuard' | Select-Object -First 1
    if ($null -eq $match) { throw 'Expected CNxStateGuard probe line is absent' }
    $lines = [IO.File]::ReadAllLines($path)
    $moduleLine = @($lines[0..($match.LineNumber - 1)] | Where-Object { $_ -notmatch '^Attribute ' }).Count
    [pscustomobject]@{ module = 'T_Core'; line = $moduleLine; source_line_sha256 = Get-StringSha256 $match.Line }
}
function Get-EnvironmentEvidence([object]$Excel) {
    $excelPath = $Excel.Path; $excel_bitness = if($excelPath -match '\\Program Files \(x86\)\\'){'32-bit'}else{'64-bit'}
    $vba7 = $false; $win64 = $false
    try { $vba7 = ([double]$Excel.VBE.Version -ge 7) } catch { $vba7 = $null }
    try { $win64 = ($excel_bitness -eq '64-bit') } catch { $win64 = $null }
    [pscustomobject]@{ os = [Environment]::OSVersion.VersionString; os_version = [Environment]::OSVersion.Version.ToString(); os_build = [Environment]::OSVersion.Version.Build; excel_version = $Excel.Version; excel_path=$excelPath; excel_bitness = $excel_bitness; VBA7 = $vba7; Win64 = $win64; macro_policy=$Excel.AutomationSecurity; powershell_version = $PSVersionTable.PSVersion.ToString(); interactive = [Environment]::UserInteractive }
}
function Invoke-OwnerCompile([string]$Artifact, [string]$StartupPath, [string]$ExcelIdentityPath, [string]$HandshakePath, [string]$GoPath, [string]$ResultPath, [string]$RunIdentity) {
    # This Windows PowerShell 5.1 process is the sole owner of Excel COM, VBE,
    # compile execution, selection capture, and COM/process cleanup.
    $excel = $null; $books = $null; $book = $null; $vbe = $null; $bars = $null; $control = $null; $excelProcess = $null; $excelPid = 0; $cleanup = @(); $ownerResult = $null; $ownerPreflight = @()
    Write-AtomicEvidence $StartupPath ([ordered]@{schema_version=1;run_id=$RunIdentity;stage='OWNER_ENTRY';owner_pid=$PID;created_utc=[DateTime]::UtcNow.ToString('o')})
    try {
        Write-AtomicEvidence $StartupPath ([ordered]@{schema_version=1;run_id=$RunIdentity;stage='EXCEL_CREATE_START';owner_pid=$PID;created_utc=[DateTime]::UtcNow.ToString('o')})
        $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excelPid = Get-ExcelPid $excel; $excelHwnd=[Int64]$excel.Hwnd
        $excelProcess=Get-Process -Id $excelPid -ErrorAction Stop;[void]$excelProcess.Handle
        [uint32]$excelHwndPid=0;[void][NxTask3.NativeWindowApi]::GetWindowThreadProcessId([IntPtr]$excelHwnd,[ref]$excelHwndPid)
        if($excelHwndPid -ne $excelPid){throw 'owner compile Excel instance is not exact-bound'}
        $excelStartedUtc=$excelProcess.StartTime.ToUniversalTime().ToString('o');$excelExecutable=[IO.Path]::GetFullPath($excelProcess.MainModule.FileName)
        Write-AtomicEvidence $ExcelIdentityPath ([ordered]@{schema_version=1;run_id=$RunIdentity;stage='EXCEL_BOUND';owner_pid=$PID;excel_pid=$excelPid;excel_hwnd=$excelHwnd;excel_started_utc=$excelStartedUtc;excel_executable=$excelExecutable;created_utc=[DateTime]::UtcNow.ToString('o')})
        Write-AtomicEvidence $StartupPath ([ordered]@{schema_version=1;run_id=$RunIdentity;stage='EXCEL_BOUND';owner_pid=$PID;created_utc=[DateTime]::UtcNow.ToString('o')})
        $books = $excel.Workbooks
        Write-AtomicEvidence $StartupPath ([ordered]@{schema_version=1;run_id=$RunIdentity;stage='ARTIFACT_OPEN_START';owner_pid=$PID;created_utc=[DateTime]::UtcNow.ToString('o')})
        $book = $books.Open($Artifact)
        Write-AtomicEvidence $StartupPath ([ordered]@{schema_version=1;run_id=$RunIdentity;stage='ARTIFACT_OPENED';owner_pid=$PID;created_utc=[DateTime]::UtcNow.ToString('o')})
        $ownerPreflight = Test-Preflight $excel $book
        Write-AtomicEvidence $StartupPath ([ordered]@{schema_version=1;run_id=$RunIdentity;stage='PREFLIGHT_COMPLETE';owner_pid=$PID;created_utc=[DateTime]::UtcNow.ToString('o')})
        $vbe = $excel.VBE; $bars = $vbe.CommandBars; $control = Find-VbeCompileControl $bars
        $vbeHwnd = 0; try { $vbeHwnd = [Int64]$vbe.MainWindow.HWnd } catch { $vbeHwnd = 0 }
        Write-AtomicEvidence $StartupPath ([ordered]@{schema_version=1;run_id=$RunIdentity;stage='VBE_READY';owner_pid=$PID;created_utc=[DateTime]::UtcNow.ToString('o')})
        Write-AtomicEvidence $HandshakePath ([ordered]@{ run_id=$RunIdentity; owner_pid=$PID; excel_pid=$excelPid; excel_hwnd=$excelHwnd; excel_started_utc=$excelStartedUtc; excel_executable=$excelExecutable; vbe_hwnd=$vbeHwnd; preflight=$ownerPreflight; environment=(Get-EnvironmentEvidence $excel) })
        Write-AtomicEvidence $StartupPath ([ordered]@{schema_version=1;run_id=$RunIdentity;stage='HANDSHAKE_PUBLISHED';owner_pid=$PID;created_utc=[DateTime]::UtcNow.ToString('o')})
        for($i=0;$i -lt 600 -and -not (Test-Path -LiteralPath $GoPath);$i++){Start-Sleep -Milliseconds 100}
        if(-not (Test-Path -LiteralPath $GoPath)){throw 'Owner compile handshake timed out before parent release'}
        $compile = Invoke-CompileWithWatcher $excel $control $excelPid $vbeHwnd $RunIdentity
        $ownerResult = [ordered]@{ run_id=$RunIdentity; owner_pid=$PID; excel_pid=$excelPid; excel_hwnd=$excelHwnd; excel_started_utc=$excelStartedUtc; excel_executable=$excelExecutable; vbe_hwnd=$vbeHwnd; watcher_started_utc=$compile.watcher_started_utc; control_id=578; dialog_text=$compile.dialog_text; dialog_sha256=$compile.dialog_sha256; dialog_closed=$compile.dialog_closed; dialog_disappeared=$compile.dialog_disappeared; stable_polls=$compile.stable_polls; dialog_owner_pid=$compile.dialog_owner_pid; watcher_status=$compile.watcher_status; watcher_job_state=$compile.watcher_job_state; watcher_error=$compile.watcher_error; candidate_count=$compile.candidate_count; max_candidate_count=$compile.max_candidate_count; candidates=$compile.candidates; active_module=$compile.active_module; line=$compile.line; caret_column=$compile.caret_column; source_line_sha256=$compile.source_line_sha256; execute_error=$compile.execute_error; pane_error=$compile.pane_error }
    } catch { $ownerResult = [ordered]@{run_id=$RunIdentity;owner_pid=$PID;excel_pid=$excelPid;failure=$_.Exception.Message} } finally {
        try {if($book){$book.Close($false)}} catch {$cleanup += $_.Exception.Message}; try {if($excel){$excel.Quit()}} catch {$cleanup += $_.Exception.Message}
        try {Release-ComObject $control;Release-ComObject $bars;Release-ComObject $vbe;Release-ComObject $book;Release-ComObject $books;Release-ComObject $excel} catch {$cleanup += $_.Exception.Message}
        [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
        if($null -ne $excelProcess){$excelCleanup=Stop-ExactProcessAfterGrace $excelProcess 'owner compile Excel';if(-not [string]::IsNullOrWhiteSpace($excelCleanup)){$cleanup += $excelCleanup};try{$excelProcess.Dispose()}catch{$cleanup += $_.Exception.Message};$excelProcess=$null}
        if($null -eq $ownerResult){$ownerResult=[ordered]@{run_id=$RunIdentity;owner_pid=$PID;excel_pid=$excelPid;failure='owner compile produced no result'}}
        $ownerResult['cleanup_failures']=$cleanup
        if($cleanup.Count -gt 0){$ownerResult['failure']=if($ownerResult.Contains('failure')){$ownerResult['failure'] + ' | owner cleanup failed: ' + ($cleanup -join ' | ')}else{'owner cleanup failed: ' + ($cleanup -join ' | ')}}
        Write-AtomicEvidence $ResultPath $ownerResult
    }
}
function Get-OwnerCompileEarlyExcelProcess([string]$IdentityPath,[string]$RunIdentity,[int]$OwnerPid) {
    $excelProcess=$null
    try {
        $excelIdentityDocument=Get-Content -LiteralPath $IdentityPath -Raw -Encoding UTF8|ConvertFrom-Json
        Assert-ExactObjectProperties $excelIdentityDocument @('schema_version','run_id','stage','owner_pid','excel_pid','excel_hwnd','excel_started_utc','excel_executable','created_utc') 'owner compile early Excel identity'
        if($excelIdentityDocument.schema_version -ne 1 -or $excelIdentityDocument.run_id -ne $RunIdentity -or $excelIdentityDocument.stage -ne 'EXCEL_BOUND' -or [int]$excelIdentityDocument.owner_pid -ne $OwnerPid -or [int]$excelIdentityDocument.excel_pid -le 0 -or [Int64]$excelIdentityDocument.excel_hwnd -eq 0){throw 'identity fields rejected'}
        Initialize-NativeWindowApi
        $excelProcess=Get-Process -Id ([int]$excelIdentityDocument.excel_pid) -ErrorAction Stop;[void]$excelProcess.Handle
        [uint32]$earlyHwndPid=0;[void][NxTask3.NativeWindowApi]::GetWindowThreadProcessId([IntPtr][Int64]$excelIdentityDocument.excel_hwnd,[ref]$earlyHwndPid)
        if($earlyHwndPid -ne [int]$excelIdentityDocument.excel_pid -or $excelProcess.StartTime.ToUniversalTime().ToString('o') -ne $excelIdentityDocument.excel_started_utc -or [IO.Path]::GetFullPath($excelProcess.MainModule.FileName) -ne [IO.Path]::GetFullPath([string]$excelIdentityDocument.excel_executable)){throw 'process binding rejected'}
        return $excelProcess
    } catch {
        if($null -ne $excelProcess){try{$excelProcess.Dispose()}catch{};$excelProcess=$null}
        throw
    }
}
function Invoke-OwnerCompileProbe([object]$BuildHost, [string]$RunIdentity, [string]$EvidenceRoot, [string]$DataRoot) {
    $id=[Guid]::NewGuid().ToString('N');$root=Join-Path ([IO.Path]::GetTempPath()) ('LHexcelOwner-'+$id);[void](New-Item -ItemType Directory -Force -Path $root)
    $startup=Join-Path $root 'startup.json';$excelIdentity=Join-Path $root 'excel-identity.json';$handshake=Join-Path $root 'handshake.json';$go=Join-Path $root 'go.signal';$result=Join-Path $root 'result.json'
    $ownerStdout=Join-Path $root 'owner.stdout.txt';$ownerStderr=Join-Path $root 'owner.stderr.txt';$startupEvidencePath=Join-Path $EvidenceRoot 'Core.RedProbe.OwnerStartup.json'
    $owner=$null;$excelProcess=$null;$probeCleanup=@();$probeFailure=$null;$probeResult=$null;$hello=$null;$waitStartedUtc=$null;$waitElapsedMs=0
    $ownerStartupTimeoutMs=180000;$handshakeObserved=$false;$handshakePublishedAtTimeout=$false;$startupTimer=[Diagnostics.Stopwatch]::StartNew()
    try {
        $owner=Start-Process -FilePath powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',$PSCommandPath,'-OwnerCompile','-OwnerStartup',$startup,'-OwnerExcelIdentity',$excelIdentity,'-OwnerArtifact',$BuildHost.ArtifactPath,'-OwnerHandshake',$handshake,'-OwnerGo',$go,'-OwnerResult',$result,'-RunId',$RunIdentity,'-DataRoot',$DataRoot) -RedirectStandardOutput $ownerStdout -RedirectStandardError $ownerStderr -PassThru
        while($startupTimer.ElapsedMilliseconds -lt $ownerStartupTimeoutMs){
            if(Test-Path -LiteralPath $handshake -PathType Leaf){$handshakeObserved=$true;break}
            if($owner.HasExited){break}
            Start-Sleep -Milliseconds 100
        }
        $handshakePublishedAtTimeout=Test-Path -LiteralPath $handshake -PathType Leaf
        if($handshakePublishedAtTimeout){$handshakeObserved=$true}
        $startupTimer.Stop()
        if(-not $handshakeObserved){
            if($owner.HasExited){throw 'Owner process exited before atomic handshake'}
            if(Test-Path -LiteralPath $startup){throw 'Owner process entered but did not publish atomic handshake'}
            throw 'Owner process did not publish owner-entry evidence'
        }
        $hello=Get-Content -LiteralPath $handshake -Raw -Encoding UTF8|ConvertFrom-Json;if([int]$hello.owner_pid -ne $owner.Id -or [int]$hello.excel_pid -le 0 -or [Int64]$hello.excel_hwnd -eq 0 -or [Int64]$hello.vbe_hwnd -eq 0){throw 'Owner handshake is incomplete or not owner-bound'}
        Initialize-NativeWindowApi
        $excelProcess=Get-Process -Id ([int]$hello.excel_pid) -ErrorAction Stop;[void]$excelProcess.Handle
        [uint32]$helloHwndPid=0;[void][NxTask3.NativeWindowApi]::GetWindowThreadProcessId([IntPtr][Int64]$hello.excel_hwnd,[ref]$helloHwndPid)
        if($helloHwndPid -ne [int]$hello.excel_pid -or $excelProcess.StartTime.ToUniversalTime().ToString('o') -ne $hello.excel_started_utc -or [IO.Path]::GetFullPath($excelProcess.MainModule.FileName) -ne [IO.Path]::GetFullPath([string]$hello.excel_executable)){throw 'owner compile Excel instance is not exact-bound'}
        [IO.File]::WriteAllText($go,'go',(New-Object Text.UTF8Encoding($false)))
        $waitStartedUtc=[DateTime]::UtcNow.ToString('o');$waitTimer=[Diagnostics.Stopwatch]::StartNew();$ownerExited=$owner.WaitForExit(60000);$waitTimer.Stop();$waitElapsedMs=$waitTimer.ElapsedMilliseconds
        if(-not $ownerExited){$owner.Kill();[void]$owner.WaitForExit(5000);throw ('Owner compile timeout exceeded 60000ms; owner_pid='+$owner.Id+'; excel_pid='+[int]$hello.excel_pid+'; wait_started_utc='+$waitStartedUtc+'; wait_elapsed_ms='+$waitElapsedMs)}
        if(-not(Test-Path -LiteralPath $result)){throw 'Owner process did not publish compile result'}
        $selection=Get-Content -LiteralPath $result -Raw -Encoding UTF8|ConvertFrom-Json
        if($null -ne $selection.PSObject.Properties['failure'] -and -not [string]::IsNullOrWhiteSpace([string]$selection.failure)){throw ('Owner compile failed: '+[string]$selection.failure)}
        if([int]$selection.owner_pid -ne $owner.Id -or [int]$selection.excel_pid -ne [int]$hello.excel_pid -or $selection.excel_started_utc -ne $hello.excel_started_utc -or [IO.Path]::GetFullPath([string]$selection.excel_executable) -ne [IO.Path]::GetFullPath([string]$hello.excel_executable) -or [int]$selection.dialog_owner_pid -ne [int]$hello.excel_pid){throw 'Owner compile result is not exact owner evidence'}
        if(@($selection.cleanup_failures).Count -gt 0){throw 'owner cleanup failed'}
        $probeResult=[pscustomobject]@{run_id=$RunIdentity;owner_pid=$hello.owner_pid;excel_pid=$hello.excel_pid;excel_hwnd=$hello.excel_hwnd;excel_started_utc=$hello.excel_started_utc;excel_executable=$hello.excel_executable;vbe_hwnd=$hello.vbe_hwnd;owner_wait_started_utc=$waitStartedUtc;owner_wait_elapsed_ms=$waitElapsedMs;watcher_started_utc=$selection.watcher_started_utc;control_id=$selection.control_id;dialog_text=$selection.dialog_text;dialog_sha256=$selection.dialog_sha256;dialog_closed=$selection.dialog_closed;dialog_disappeared=$selection.dialog_disappeared;stable_polls=$selection.stable_polls;dialog_owner_pid=$selection.dialog_owner_pid;watcher_status=$selection.watcher_status;watcher_job_state=$selection.watcher_job_state;watcher_error=$selection.watcher_error;candidate_count=$selection.candidate_count;max_candidate_count=$selection.max_candidate_count;candidates=$selection.candidates;active_module=$selection.active_module;line=$selection.line;caret_column=$selection.caret_column;source_line_sha256=$selection.source_line_sha256;execute_error=$selection.execute_error;environment=$hello.environment;preflight=$hello.preflight}
    } catch {
        $probeFailure=$_.Exception.Message
    } finally {
        if($startupTimer.IsRunning){$startupTimer.Stop()}
        if(-not $handshakeObserved){
            $ownerExited=$false;$ownerExitCode=$null;$ownerResultFailure=$null;$stdoutSha=$null;$stderrSha=$null;$startupOwnerPid=$null;$startupStage=$null
            if($null -ne $owner){try{$ownerExited=$owner.HasExited;if($ownerExited){$ownerExitCode=[int]$owner.ExitCode}}catch{}}
            if(Test-Path -LiteralPath $startup -PathType Leaf){try{$startupDocument=Get-Content -LiteralPath $startup -Raw -Encoding UTF8|ConvertFrom-Json;$startupOwnerPid=[int]$startupDocument.owner_pid;$startupStage=[string]$startupDocument.stage}catch{}}
            if(Test-Path -LiteralPath $result -PathType Leaf){try{$ownerResultDocument=Get-Content -LiteralPath $result -Raw -Encoding UTF8|ConvertFrom-Json;if($null -ne $ownerResultDocument.PSObject.Properties['failure']){$ownerResultFailure=[string]$ownerResultDocument.failure}}catch{}}
            $earlyExcelPid=$null
            if(Test-Path -LiteralPath $excelIdentity -PathType Leaf){
                try{
                    $excelProcess=Get-OwnerCompileEarlyExcelProcess $excelIdentity $RunIdentity $owner.Id
                    $earlyExcelPid=$excelProcess.Id
                }catch{
                    $probeCleanup += ('owner compile early Excel identity rejected: '+$_.Exception.GetType().FullName)
                }
            }
            if($null -ne $owner -and -not $owner.HasExited){$ownerCleanup=Stop-ExactProcessAfterGrace $owner 'owner compile PowerShell';if(-not [string]::IsNullOrWhiteSpace($ownerCleanup)){$probeCleanup += $ownerCleanup}}
            if($null -ne $owner){try{$ownerExited=$owner.HasExited;if($ownerExited){$ownerExitCode=[int]$owner.ExitCode}}catch{}}
            $sensitiveValues=@($root,$DataRoot,[string]$BuildHost.ArtifactPath,$PSCommandPath,$env:USERPROFILE,$env:TEMP)
            if(-not [string]::IsNullOrWhiteSpace($ownerResultFailure)){foreach($value in $sensitiveValues){if(-not [string]::IsNullOrWhiteSpace($value)){$ownerResultFailure=$ownerResultFailure.Replace($value,'<path>')}};if($ownerResultFailure.Length -gt 1024){$ownerResultFailure=$ownerResultFailure.Substring($ownerResultFailure.Length-1024)}}
            try{if(Test-Path -LiteralPath $ownerStdout -PathType Leaf){$stdoutSha=Get-Sha256 $ownerStdout}}catch{}
            try{if(Test-Path -LiteralPath $ownerStderr -PathType Leaf){$stderrSha=Get-Sha256 $ownerStderr}}catch{}
            $handshakePublishedAfterTimeout=Test-Path -LiteralPath $handshake -PathType Leaf
            Write-AtomicEvidence $startupEvidencePath ([ordered]@{schema_version=1;run_id=$RunIdentity;status='HANDSHAKE_TIMEOUT';owner_pid=if($null -ne $owner){$owner.Id}else{$null};startup_published=(Test-Path -LiteralPath $startup -PathType Leaf);startup_owner_pid=$startupOwnerPid;startup_stage=$startupStage;excel_identity_published=(Test-Path -LiteralPath $excelIdentity -PathType Leaf);early_excel_pid=$earlyExcelPid;handshake_published=$handshakePublishedAfterTimeout;handshake_published_at_timeout=$handshakePublishedAtTimeout;handshake_published_after_timeout=$handshakePublishedAfterTimeout;result_published=(Test-Path -LiteralPath $result -PathType Leaf);owner_exited=$ownerExited;owner_exit_code=$ownerExitCode;startup_wait_elapsed_ms=$startupTimer.ElapsedMilliseconds;owner_startup_timeout_ms=$ownerStartupTimeoutMs;owner_stdout_sha256=$stdoutSha;owner_stderr_sha256=$stderrSha;owner_stdout_tail=(Get-BoundedDiagnosticText $ownerStdout $sensitiveValues);owner_stderr_tail=(Get-BoundedDiagnosticText $ownerStderr $sensitiveValues);owner_result_failure=$ownerResultFailure;failure=$probeFailure;created_utc=[DateTime]::UtcNow.ToString('o')})
        }
        if($null -ne $excelProcess){$excelCleanup=Stop-ExactProcessAfterGrace $excelProcess 'owner compile Excel';if(-not [string]::IsNullOrWhiteSpace($excelCleanup)){$probeCleanup += $excelCleanup};try{$excelProcess.Dispose()}catch{$probeCleanup += $_.Exception.Message};$excelProcess=$null}
        if($null -ne $owner -and -not $owner.HasExited){$ownerCleanup=Stop-ExactProcessAfterGrace $owner 'owner compile PowerShell';if(-not [string]::IsNullOrWhiteSpace($ownerCleanup)){$probeCleanup += $ownerCleanup}}
        if($null -ne $owner){try{$owner.Dispose()}catch{$probeCleanup += $_.Exception.Message};$owner=$null}
        Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
        if(Test-Path -LiteralPath $root){$probeCleanup += 'Owner temp residue'}
    }
    if($probeCleanup.Count -gt 0){$cleanupFailure='Owner compile exact cleanup failed: '+($probeCleanup -join ' | ');$probeFailure=if([string]::IsNullOrWhiteSpace($probeFailure)){$cleanupFailure}else{$probeFailure+' | '+$cleanupFailure}}
    if(-not [string]::IsNullOrWhiteSpace($probeFailure)){throw $probeFailure}
    return $probeResult
}
function Resolve-SuiteMapping([string]$Root, [string]$RequestedSuite) {
    $inventoryPath = Join-Path $PSScriptRoot 'feature-suite-entrypoints.json'
    $inventory = Get-Content -LiteralPath $inventoryPath -Raw | ConvertFrom-Json
    if ($inventory.schema_version.GetType().FullName -ne 'System.Int32' -or $inventory.schema_version -ne 1) { throw 'Suite inventory schema rejected' }
    $exactMappings = @{
        "Core" = @{ owner='core'; module='T_Core'; list='T_Core.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest=$null }
        "Frame" = @{ owner='frame'; module='T_Frame'; list='T_Frame.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest='tests/manifests/Frame.json' }
        "AI" = @{ owner='ai'; module='T_Ai'; list='T_Ai.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest='tests/manifests/AI.json' }
        "G005" = @{ owner='g005'; module='T_G005'; list='T_G005.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest='tests/manifests/G005.json' }
        "Template" = @{ owner='template'; module='T_Template'; list='T_Template.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest='tests/manifests/Template.json' }
        "Data" = @{ owner='data'; module='T_Data'; list='T_Data.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest='tests/manifests/Data.json' }
        "Draw" = @{ owner='draw'; module='T_Draw'; list='T_Draw.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest='tests/manifests/Draw.json' }
        "File" = @{ owner='file'; module='T_File'; list='T_File.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest='tests/manifests/File.json' }
        "Calculator" = @{ owner='calculator'; module='T_Calculator'; list='T_Calculator.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest='tests/manifests/Calculator.json' }
        "Symbols" = @{ owner='symbols'; module='T_Symbols'; list='T_Symbols.TestNames'; run='NxTestHarness.RunSuite'; extension_manifest='tests/manifests/Symbols.json' }
    }
    if(@($inventory.suites.PSObject.Properties).Count -ne $exactMappings.Count){throw 'Suite inventory exact mapping rejected'}
    foreach($suiteName in $exactMappings.Keys){
        $candidate=$inventory.suites.PSObject.Properties[$suiteName]
        if($null -eq $candidate){throw 'Suite inventory exact mapping rejected'}
        $candidateFields=@($candidate.Value.PSObject.Properties.Name|Sort-Object)
        if(($candidateFields -join '|') -ne 'extension_manifest|list|module|owner|run'){throw 'Suite inventory exact mapping rejected'}
        $expectedRecord=$exactMappings[$suiteName]
        foreach($field in @('owner','module','list','run','extension_manifest')){
            if($candidate.Value.$field -ne $expectedRecord[$field]){throw 'Suite inventory exact mapping rejected'}
        }
    }
    $property = $inventory.suites.PSObject.Properties[$RequestedSuite]
    if ($null -eq $property) { throw 'Suite inventory has no requested suite' }
    $record = $property.Value
    $actual = @($record.PSObject.Properties.Name | Sort-Object)
    $expected = @('extension_manifest','list','module','owner','run')
    if (($actual -join '|') -ne ($expected -join '|')) { throw 'Suite inventory record schema rejected' }
    foreach ($name in @('owner','module','list','run')) { if ($record.$name -isnot [string] -or [string]::IsNullOrWhiteSpace([string]$record.$name)) { throw 'Suite inventory field rejected: ' + $name } }
    return $record
}
function Write-OwnerCaseProgress([string]$Path, [string]$Marker) {
    try {
        if(-not [string]::IsNullOrWhiteSpace($Path)){
            [IO.File]::AppendAllText($Path,$Marker+[Environment]::NewLine,[Text.Encoding]::ASCII)
        }
    } catch {}
}
function Invoke-OwnerGreenCase([string]$Artifact, [string]$HandshakePath, [string]$ResultPath, [string]$RunIdentity, [string]$CaseIdentity, [string]$SuiteName, [string]$CaseName, [string]$RunEntry, [int]$ExpectedCount, [string]$BuildEvidenceSha256) {
    $caseExcel=$null;$caseBooks=$null;$caseBook=$null
    $dataBook=$null;$dataSheets=$null;$dataSheet=$null;$dataRange=$null
    $excelPid=0;$excelStartedUtc=$null;$excelExecutable=$null;$cleanup=@();$raw=$null;$pre=$null;$post=$null;$failure=$null
    $originalTemp=$env:TEMP;$originalTmp=$env:TMP;$ownerTempRoot=$null;$remainingBook=$null
    $progressPath=[IO.Path]::ChangeExtension($ResultPath,'.progress.txt')
    try {
        Write-OwnerCaseProgress $progressPath 'owner-00-entry'
        $ownerTempRoot=Split-Path -Parent $ResultPath
        [void](New-Item -ItemType Directory -Force -Path $ownerTempRoot)
        $env:TEMP=$ownerTempRoot
        $env:TMP=$ownerTempRoot
        $env:LHEXCEL_OWNER_CASE_PROGRESS=$progressPath
        if($CaseName -eq 'TestAiComboChoicesAreFixed'){
            if($BuildEvidenceSha256 -notmatch '^[0-9a-f]{64}$'){throw 'AI ComboBox build evidence attestation missing'}
            $env:LHEXCEL_AI_COMBO_EVIDENCE_STATUS='VERIFIED_EXACT_V1'
            $env:LHEXCEL_AI_COMBO_EVIDENCE_SHA256=$BuildEvidenceSha256
            $env:LHEXCEL_AI_COMBO_EVIDENCE_RUN_ID=$RunIdentity
        }
        Write-OwnerCaseProgress $progressPath 'owner-00-before-excel'
        $caseExcel = New-Object -ComObject Excel.Application
        Write-OwnerCaseProgress $progressPath 'owner-00-after-excel'
        $caseExcel.Visible = $false
        $excelPid = Get-ExcelPid $caseExcel
        $excelIdentity=Get-Process -Id $excelPid -ErrorAction Stop
        try{[void]$excelIdentity.Handle;$excelStartedUtc=$excelIdentity.StartTime.ToUniversalTime().ToString('o');$excelExecutable=[IO.Path]::GetFullPath($excelIdentity.MainModule.FileName)}finally{$excelIdentity.Dispose()}
        Write-AtomicEvidence $HandshakePath ([ordered]@{schema_version=1;run_id=$RunIdentity;case_id=$CaseIdentity;case_name=$CaseName;suite=$SuiteName;owner_pid=$PID;excel_pid=$excelPid;excel_hwnd=[Int64]$caseExcel.Hwnd;excel_started_utc=$excelStartedUtc;excel_executable=$excelExecutable;created_utc=[DateTime]::UtcNow.ToString('o')})
        $caseBooks = $caseExcel.Workbooks
        $caseBook = $caseBooks.Open($Artifact)
        $dataBook = $caseBooks.Add()
        $dataSheets = $dataBook.Worksheets
        $dataSheet = $dataSheets.Item(1)
        [void]$dataSheet.Activate()
        $dataRange = $dataSheet.Range("A1")
        [void]$dataRange.Select()
        $pre = [pscustomobject]@{application=$caseExcel.Version;workbook=$dataBook.Name;selection=$dataRange.Address();source_digest=(Get-StringSha256 ((Get-RangeFormulaForDigest $dataRange) + '|' + [string]$dataRange.Value2));workbook_count=[int]$caseBooks.Count}
        Write-OwnerCaseProgress $progressPath 'owner-01-before-run'
        $raw = $caseExcel.Run("'" + $caseBook.Name + "'!" + $RunEntry, $SuiteName, $CaseName) | ConvertFrom-Json
        Write-OwnerCaseProgress $progressPath 'owner-02-after-run'
        $post = [pscustomobject]@{application=$caseExcel.Version;workbook=$dataBook.Name;selection=$dataRange.Address();source_digest=(Get-StringSha256 ((Get-RangeFormulaForDigest $dataRange) + '|' + [string]$dataRange.Value2));workbook_count=[int]$caseBooks.Count}
        Write-OwnerCaseProgress $progressPath 'owner-03-post-snapshot'
        $passedDetail=@($raw.tests|Where-Object {$_.status -eq 'passed'})
        if ([int]$raw.total -ne $ExpectedCount -or $raw.tests.Count -ne $ExpectedCount -or [int]$raw.failed -ne 0 -or [int]$raw.passed -ne 1 -or [int]$raw.skipped -ne ($ExpectedCount - 1) -or $passedDetail.Count -ne 1 -or $passedDetail[0].name -ne $CaseName -or @($raw.tests|Where-Object {$_.status -notin @('passed','skipped')}).Count -ne 0) {
            $failedDetail = @($raw.tests | Where-Object { $_.status -eq 'failed' } | Select-Object -First 1)
            $failureDetail = if($failedDetail.Count -eq 1){[string]$failedDetail[0].error}else{'missing failed-case detail'}
            throw ('owner case failed: ' + $CaseName + ' | ' + $failureDetail)
        }
        if([int]$post.workbook_count -ne [int]$pre.workbook_count){throw ('owner case workbook inventory changed: '+[int]$pre.workbook_count+' -> '+[int]$post.workbook_count)}
    } catch {$failure=$_.Exception.Message} finally {
        Write-OwnerCaseProgress $progressPath 'owner-04-finally-enter'
        try { if($caseExcel){$caseExcel.DisplayAlerts=$false} } catch {$cleanup += $_.Exception.Message}
        try { if($dataBook){[void]$dataBook.Close($false)} } catch {$cleanup += $_.Exception.Message}
        Write-OwnerCaseProgress $progressPath 'owner-05-data-closed'
        try { if($caseBook){[void]$caseBook.Close($false)} } catch {$cleanup += $_.Exception.Message}
        Write-OwnerCaseProgress $progressPath 'owner-06-host-closed'
        try {
            $remainingCount=if($caseBooks){[int]$caseBooks.Count}else{0}
            for($index=$remainingCount;$index -ge 1;$index--){
                $remainingBook=$null
                try {$remainingBook=$caseBooks.Item($index);[void]$remainingBook.Close($false)} catch {$cleanup += $_.Exception.Message} finally {Release-ComObject $remainingBook;$remainingBook=$null}
            }
        } catch {$cleanup += $_.Exception.Message}
        try { if($caseExcel){[void]$caseExcel.Quit()} } catch {$cleanup += $_.Exception.Message}
        Write-OwnerCaseProgress $progressPath 'owner-07-excel-quit'
        try { Release-ComObject $dataRange } catch {$cleanup += $_.Exception.Message}
        try { Release-ComObject $dataSheet } catch {$cleanup += $_.Exception.Message}
        try { Release-ComObject $dataSheets } catch {$cleanup += $_.Exception.Message}
        try { Release-ComObject $dataBook } catch {$cleanup += $_.Exception.Message}
        try { Release-ComObject $caseBook } catch {$cleanup += $_.Exception.Message}
        try { Release-ComObject $caseBooks } catch {$cleanup += $_.Exception.Message}
        try { Release-ComObject $caseExcel } catch {$cleanup += $_.Exception.Message}
        $dataRange=$null;$dataSheet=$null;$dataSheets=$null;$dataBook=$null
        $caseBook=$null;$caseBooks=$null;$caseExcel=$null
        Write-OwnerCaseProgress $progressPath 'owner-08-com-released'
        Remove-Item Env:LHEXCEL_AI_COMBO_EVIDENCE_STATUS -ErrorAction SilentlyContinue
        Remove-Item Env:LHEXCEL_AI_COMBO_EVIDENCE_SHA256 -ErrorAction SilentlyContinue
        Remove-Item Env:LHEXCEL_AI_COMBO_EVIDENCE_RUN_ID -ErrorAction SilentlyContinue
        Remove-Item Env:LHEXCEL_OWNER_CASE_PROGRESS -ErrorAction SilentlyContinue
        try{$env:TEMP=$originalTemp;$env:TMP=$originalTmp}catch{$cleanup += $_.Exception.Message}
        $ownerStatus=if($null -eq $failure -and $cleanup.Count -eq 0){'PASS'}else{'FAIL'}
        Write-OwnerCaseProgress $progressPath 'owner-09-before-result'
        Write-AtomicEvidence $ResultPath ([ordered]@{schema_version=1;run_id=$RunIdentity;case_id=$CaseIdentity;case_name=$CaseName;suite=$SuiteName;owner_pid=$PID;excel_pid=$excelPid;excel_started_utc=$excelStartedUtc;excel_executable=$excelExecutable;status=$ownerStatus;raw_result=$raw;pre=$pre;post=$post;cleanup_failures=@($cleanup);failure=$failure;finished_utc=[DateTime]::UtcNow.ToString('o')})
        Write-OwnerCaseProgress $progressPath 'owner-10-result-written'
        [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
        Write-OwnerCaseProgress $progressPath 'owner-11-gc-complete'
    }
}
function Invoke-OwnerGreenCaseProbe([object]$TestHost, [object]$suiteMapping, [string]$CaseName, [string]$EvidenceRoot, [string]$SuiteName, [int]$ExpectedCount, [string]$RunIdentity, [string]$BuildEvidenceSha256) {
    $caseId=[Guid]::NewGuid().ToString('N');$caseFolder=$null;$caseRoot=$null;$caseRootLength=0;$pathAscii=$false;$owner=$null;$ownerPid=0
    $hello=$null;$selection=$null;$excelProcess=$null;$excelPid=0;$verifiedExcelPid=0
    $caseFailure=$null;$cleanupFailures=@();$processResidue=@();$casePre=$null;$casePost=$null;$progressPath=$null;$caseProgress=@();$ownerExitMode=$null;$excelExitMode=$null
    $ownerStartupTimeoutMs=180000;$startupTimer=$null;$caseExecutionTimeoutMs=180000;$resultTimer=$null;$excelPidsBeforeOwner=@();$ownerStartedUtc=$null
    try {
        $tempRoot=[IO.Path]::GetFullPath([IO.Path]::GetTempPath())
        if($tempRoot -match '[^\x20-\x7e]'){throw 'owner case ASCII temp root unavailable'}
        $caseFolder='c-'+$caseId.Substring(0,16)
        $caseRoot=Join-Path $tempRoot $caseFolder
        if($caseRoot -match '[^\x20-\x7e]'){throw 'owner case ASCII temp root unavailable'}
        $pathAscii=$true
        if(Test-Path -LiteralPath $caseRoot){throw 'owner case compact temp collision'}
        $hostCopy=Join-Path $caseRoot 'host.xlam';$handshake=Join-Path $caseRoot 'handshake.json';$result=Join-Path $caseRoot 'result.json';$request=Join-Path $caseRoot 'request.json';$progressPath=[IO.Path]::ChangeExtension($result,'.progress.txt')
        [void](New-Item -ItemType Directory -Path $caseRoot)
        Copy-Item -LiteralPath $TestHost.ArtifactPath -Destination $hostCopy -Force
        $ownerDataRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $caseEvidenceSha256=if($CaseName -eq 'TestAiComboChoicesAreFixed'){$BuildEvidenceSha256}else{$null}
        Write-AtomicEvidence $request ([ordered]@{schema_version=1;artifact=$hostCopy;handshake_path=$handshake;result_path=$result;run_id=$RunIdentity;case_id=$caseId;case_name=$CaseName;suite=$SuiteName;run_entry=[string]$suiteMapping.run;expected_count=$ExpectedCount;build_evidence_sha256=$caseEvidenceSha256})
        $scriptLiteral=$PSCommandPath.Replace("'","''");$requestLiteral=$request.Replace("'","''");$dataRootLiteral=$ownerDataRoot.Replace("'","''")
        $ownerCommand="& '"+$scriptLiteral+"' -OwnerCase -OwnerCaseRequest '"+$requestLiteral+"' -DataRoot '"+$dataRootLiteral+"'; exit "+'$LASTEXITCODE'
        $encodedCommand=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($ownerCommand))
        $excelPidsBeforeOwner=@(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction SilentlyContinue|ForEach-Object {[uint32]$_.ProcessId})
        $owner=Start-Process -FilePath powershell.exe -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',$encodedCommand) -PassThru
        $ownerPid=$owner.Id
        $ownerStartedUtc=$owner.StartTime.ToUniversalTime()
        $startupTimer=[Diagnostics.Stopwatch]::StartNew()
        while($startupTimer.ElapsedMilliseconds -lt $ownerStartupTimeoutMs){
            if(Test-Path -LiteralPath $handshake -PathType Leaf){break}
            if($owner.HasExited){break}
            Start-Sleep -Milliseconds 100
        }
        $startupTimer.Stop()
        if(-not(Test-Path -LiteralPath $handshake -PathType Leaf)){
            if($owner.HasExited){throw 'owner case process exited before handshake'}
            throw ('owner case handshake timed out after '+$startupTimer.ElapsedMilliseconds+'ms')
        }
        $hello=Get-Content -LiteralPath $handshake -Raw -Encoding UTF8|ConvertFrom-Json
        Assert-ExactObjectProperties $hello @('schema_version','run_id','case_id','case_name','suite','owner_pid','excel_pid','excel_hwnd','excel_started_utc','excel_executable','created_utc') 'owner case handshake'
        if($hello.schema_version -ne 1 -or [int]$hello.owner_pid -ne $owner.Id -or [int]$hello.excel_pid -le 0 -or [Int64]$hello.excel_hwnd -eq 0 -or $hello.run_id -ne $RunIdentity -or $hello.case_id -ne $caseId -or $hello.case_name -ne $CaseName -or $hello.suite -ne $SuiteName){throw 'owner case handshake is not exact-bound'}
        $excelPid=[uint32]$hello.excel_pid
        $excelProcess=Get-Process -Id $excelPid -ErrorAction Stop;[void]$excelProcess.Handle
        [uint32]$hwndPid=0;[void][NxTask3.NativeWindowApi]::GetWindowThreadProcessId([IntPtr][Int64]$hello.excel_hwnd,[ref]$hwndPid)
        if($hwndPid -ne $excelPid -or $excelProcess.StartTime.ToUniversalTime().ToString('o') -ne $hello.excel_started_utc -or [IO.Path]::GetFullPath($excelProcess.MainModule.FileName) -ne [IO.Path]::GetFullPath([string]$hello.excel_executable)){throw 'owner case Excel instance is not exact-bound'}
        $resultTimer=[Diagnostics.Stopwatch]::StartNew()
        while($resultTimer.ElapsedMilliseconds -lt $caseExecutionTimeoutMs){
            if(Test-Path -LiteralPath $result -PathType Leaf){break}
            if($owner.HasExited){break}
            Start-Sleep -Milliseconds 100
        }
        $resultTimer.Stop()
        $resultPublished=Test-Path -LiteralPath $result -PathType Leaf
        if(-not $resultPublished){
            if($owner.HasExited){throw 'owner case result missing after process exit'}
            throw ('owner case result timeout exceeded '+$caseExecutionTimeoutMs+'ms')
        }
        $selection=Get-Content -LiteralPath $result -Raw -Encoding UTF8|ConvertFrom-Json
        $casePre=$selection.pre;$casePost=$selection.post
        Assert-ExactObjectProperties $selection @('schema_version','run_id','case_id','case_name','suite','owner_pid','excel_pid','excel_started_utc','excel_executable','status','raw_result','pre','post','cleanup_failures','failure','finished_utc') 'owner case result'
        if($selection.schema_version -ne 1 -or [int]$selection.owner_pid -ne $owner.Id -or [int]$selection.excel_pid -ne $excelPid -or $selection.excel_started_utc -ne $hello.excel_started_utc -or [IO.Path]::GetFullPath([string]$selection.excel_executable) -ne [IO.Path]::GetFullPath([string]$hello.excel_executable) -or $selection.run_id -ne $RunIdentity -or $selection.case_id -ne $caseId -or $selection.case_name -ne $CaseName -or $selection.suite -ne $SuiteName){throw 'owner case result is not exact-bound'}
        if($selection.status -ne 'PASS' -or @($selection.cleanup_failures).Count -ne 0 -or -not [string]::IsNullOrWhiteSpace([string]$selection.failure)){throw ('owner case failed: '+$CaseName+' | '+[string]$selection.failure)}
        if($owner.WaitForExit(15000)){
            $ownerExitMode='NATURAL'
        }else{
            $ownerLifecycle=Stop-ExactProcessAfterGrace -Process $owner -Label 'owner case PowerShell' -GraceMs 10000 -Detailed
            $ownerExitMode=[string]$ownerLifecycle.exit_mode
            if(-not [string]::IsNullOrWhiteSpace([string]$ownerLifecycle.failure)){throw ('owner case process containment failed after atomic result: '+[string]$ownerLifecycle.failure)}
        }
        if($ownerExitMode -ne 'FORCED_CONTAINED' -and $owner.ExitCode -ne 0){throw ('owner case process exit code rejected: '+$owner.ExitCode)}
        $verifiedExcelPid=$excelPid
    } catch {$caseFailure=$_.Exception.Message} finally {
        if($null -ne $startupTimer -and $startupTimer.IsRunning){$startupTimer.Stop()}
        if($null -ne $resultTimer -and $resultTimer.IsRunning){$resultTimer.Stop()}
        if($null -ne $owner -and -not $owner.HasExited){
            $ownerLifecycle=Stop-ExactProcessAfterGrace -Process $owner -Label 'owner case PowerShell' -GraceMs 0 -Detailed
            $ownerExitMode=[string]$ownerLifecycle.exit_mode
            if(-not [string]::IsNullOrWhiteSpace([string]$ownerLifecycle.failure)){$cleanupFailures += [string]$ownerLifecycle.failure}
        }elseif($null -ne $owner -and [string]::IsNullOrWhiteSpace($ownerExitMode)){
            $ownerExitMode='NATURAL'
        }
        if($null -eq $excelProcess -and $excelPid -eq 0 -and $ownerPid -gt 0){
            try {
                $ownerExcelCandidates=@(Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" -ErrorAction Stop|Where-Object {$excelPidsBeforeOwner -notcontains [uint32]$_.ProcessId -and [string]$_.CommandLine -match '(?i)EXCEL\.EXE"?\s+/automation\s+-Embedding\s*$'})
                if($ownerExcelCandidates.Count -eq 1){
                    $candidatePid=[uint32]$ownerExcelCandidates[0].ProcessId
                    $candidateProcess=Get-Process -Id $candidatePid -ErrorAction Stop
                    [void]$candidateProcess.Handle
                    if($null -ne $ownerStartedUtc -and $candidateProcess.StartTime.ToUniversalTime() -lt $ownerStartedUtc){
                        $candidateProcess.Dispose()
                        $processResidue += ('owner case pre-handshake Excel candidate predates owner: '+$candidatePid)
                    }else{
                        $excelPid=$candidatePid
                        $excelProcess=$candidateProcess
                    }
                }elseif($ownerExcelCandidates.Count -gt 1){
                    $processResidue += ('owner case pre-handshake Excel identity ambiguous for owner '+$ownerPid)
                }
            } catch {$cleanupFailures += ('owner case pre-handshake Excel recovery failed: '+$_.Exception.Message)}
        }
        if($null -ne $excelProcess){
            $excelLifecycle=Stop-ExactProcessAfterGrace -Process $excelProcess -Label 'owner case Excel' -GraceMs 10000 -Detailed
            $excelExitMode=[string]$excelLifecycle.exit_mode
            if(-not [string]::IsNullOrWhiteSpace([string]$excelLifecycle.failure)){$processResidue += ('owner case Excel cleanup failed: '+[string]$excelLifecycle.failure)}
            try{$excelProcess.Dispose()}catch{$cleanupFailures += $_.Exception.Message}
            $excelProcess=$null
        }elseif($excelPid -gt 0){$processResidue += ('owner case Excel identity unavailable for cleanup: '+$excelPid)}
        if($null -ne $owner){try{$owner.Dispose()}catch{$cleanupFailures += $_.Exception.Message};$owner=$null}
        if(-not [string]::IsNullOrWhiteSpace($progressPath) -and (Test-Path -LiteralPath $progressPath)){
            try{$caseProgress=@(Get-Content -LiteralPath $progressPath -Encoding UTF8|Where-Object {-not [string]::IsNullOrWhiteSpace([string]$_)})}catch{$cleanupFailures += ('owner case progress read failed: '+$_.Exception.Message)}
        }
        if(-not [string]::IsNullOrWhiteSpace($caseRoot)){
            Remove-Item -LiteralPath $caseRoot -Recurse -Force -ErrorAction SilentlyContinue
            if(Test-Path -LiteralPath $caseRoot){$cleanupFailures += 'owner case temp residue'}
        }
    }
    $cleanupSummary=@($cleanupFailures+$processResidue)
    if($caseProgress.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($caseFailure)){$caseFailure += ' | owner case progress: '+($caseProgress -join '>')}
    if($cleanupSummary.Count -gt 0){$caseFailure=if([string]::IsNullOrWhiteSpace($caseFailure)){'owner case cleanup failed: '+($cleanupSummary -join ' | ')}else{$caseFailure+' | cleanup: '+($cleanupSummary -join ' | ')}}
    $caseStatus=if([string]::IsNullOrWhiteSpace($caseFailure)){'PASS'}else{'FAIL'}
    $excelPidSet=if($excelPid -gt 0){@($excelPid)}else{@()}
    $caseRootLength=if([string]::IsNullOrWhiteSpace($caseRoot)){0}else{$caseRoot.Length}
    [pscustomobject]@{name=$CaseName;temp_dir=$caseRoot;path_evidence=([pscustomobject]@{case_root_length=$caseRootLength;owner_temp_root_length=$caseRootLength;case_folder=$caseFolder;ascii=$pathAscii});owner_pid=$ownerPid;excel_pid_set=$excelPidSet;pre=$casePre;post=$casePost;restore_failures=@($cleanupFailures);process_residue=@($processResidue);operation_status=$caseStatus;owner_exit_mode=$ownerExitMode;excel_exit_mode=$excelExitMode;status=$caseStatus;failure=$caseFailure}
}
function Get-RetryableOwnerFailureCategory([object]$Case) {
    if($null -eq $Case -or $Case.status -eq 'PASS'){return $null}
    $failure=[string]$Case.failure
    if($failure -match '(?i)0x800706BA|0x800706BE'){return 'transport'}
    if($failure -match '(?i)owner case process exited before handshake|owner case handshake timed out'){return 'startup'}
    return $null
}
function Invoke-IsolatedGreenCases([object]$TestHost, [object]$suiteMapping, [object[]]$Names, [string]$EvidenceRoot, [string]$Suite, [string]$RunIdentity, [string]$BuildEvidenceSha256) {
    $cases=@();$passed=0;$retryHistory=@()
    foreach($caseName in $Names){
        $case=Invoke-OwnerGreenCaseProbe $TestHost $suiteMapping ([string]$caseName) $EvidenceRoot $Suite $Names.Count $RunIdentity $BuildEvidenceSha256
        $retryCategory=Get-RetryableOwnerFailureCategory $case
        if(-not [string]::IsNullOrWhiteSpace([string]$retryCategory)){
            $retryHistory+=([pscustomobject]@{name=[string]$caseName;category=$retryCategory;failure=[string]$case.failure})
            Start-Sleep -Milliseconds 1000
            $case=Invoke-OwnerGreenCaseProbe $TestHost $suiteMapping ([string]$caseName) $EvidenceRoot $Suite $Names.Count $RunIdentity $BuildEvidenceSha256
        }
        $cases+=$case
        if($case.status -ne 'PASS'){break}
        $passed++
    }
    $failed=@($cases|Where-Object {$_.status -eq 'FAIL'}).Count;$skipped=$Names.Count-$cases.Count
    $transportRetryHistory=@($retryHistory|Where-Object {$_.category -eq 'transport'})
    [pscustomobject]@{ total=$Names.Count; passed=$passed; failed=$failed; skipped=$skipped; ordered_names=@($Names); tests=$cases; isolation=@{case_count=$cases.Count;fresh_excel_per_case=$true;owner_process_per_case=$true;unique_ascii_temp_host=$true;separate_data_workbook_per_case=$true;infrastructure_retry_limit=1;infrastructure_retries_used=$retryHistory.Count;infrastructure_retry_history=@($retryHistory);transport_retry_limit=1;transport_retries_used=$transportRetryHistory.Count;transport_retry_history=@($transportRetryHistory)} }
}

if ($OwnerCompile) { Invoke-OwnerCompile $OwnerArtifact $OwnerStartup $OwnerExcelIdentity $OwnerHandshake $OwnerGo $OwnerResult $RunId; exit 0 }
if ($OwnerCase) {
    $ownerCaseText=[IO.File]::ReadAllText($OwnerCaseRequest,[Text.Encoding]::UTF8);$ownerCasePayload=$ownerCaseText|ConvertFrom-Json
    Assert-ExactObjectProperties $ownerCasePayload @('schema_version','artifact','handshake_path','result_path','run_id','case_id','case_name','suite','run_entry','expected_count','build_evidence_sha256') 'owner case request'
    if($ownerCasePayload.schema_version -ne 1){throw 'owner case request schema rejected'}
    Invoke-OwnerGreenCase $ownerCasePayload.artifact $ownerCasePayload.handshake_path $ownerCasePayload.result_path $ownerCasePayload.run_id $ownerCasePayload.case_id $ownerCasePayload.suite $ownerCasePayload.case_name $ownerCasePayload.run_entry ([int]$ownerCasePayload.expected_count) ([string]$ownerCasePayload.build_evidence_sha256)
    exit 0
}

if ($Mode -eq 'RedProbe' -and ($Suite -ne 'Core' -or $ExpectedCompileFailure -ne 'CNxStateGuard')) { throw 'RedProbe is Core-only with CNxStateGuard' }
if ($Mode -eq 'Green' -and -not [string]::IsNullOrWhiteSpace($ExpectedCompileFailure)) { throw 'Green does not accept expected compile failure' }
$root = [IO.Path]::GetFullPath($DataRoot)
if (-not (Test-Path -LiteralPath $root -PathType Container)) { throw 'DataRoot directory missing' }
$suiteMapping = Resolve-SuiteMapping $root $Suite
$evidenceBase = if([string]::IsNullOrWhiteSpace($EvidenceRoot)) { Join-Path $root 'out/evidence/vba' } else { $EvidenceRoot }
$evidencePath = Join-Path $evidenceBase $(if ($Mode -eq 'RedProbe') { 'Core.RedProbe.json' } else { $Suite + '.json' })
$builder = Join-Path $PSScriptRoot 'Build-TestXlam.ps1'
$excel = $null; $workbooks = $null; $book = $null; $vbe = $null; $bars = $null; $control = $null; $excelProcess = $null
$cleanupFailures = @(); $preflight = @(); $compile = $null; $run = $null; $environment = $null; $failure = $null; $excelPid = 0; $testHost = $null; $buildAttestation=$null; $terminalCode = 10
try {
    $testHost = & $builder -Mode $Mode -Suite $Suite -SuiteOwner $suiteMapping.owner -ExpectedCompileFailure $ExpectedCompileFailure -ExtensionManifest $ExtensionManifest -DataRoot $root -RunId $RunId -SourceDigest $SourceDigest -SnapshotDigest $SnapshotDigest -EvidenceRoot $evidenceBase
        if($Mode -eq 'Green' -and $Suite -in @('AI','Template','Data','Draw','File','Calculator','Symbols')){$buildAttestation=Assert-UserFormBuildEvidence $testHost $RunId $SourceDigest $SnapshotDigest $root $Suite ([string]$suiteMapping.owner) $ExtensionManifest}
    if ($Mode -eq 'RedProbe') {
        $terminalCode = 22; $compile = Invoke-OwnerCompileProbe $testHost $RunId $evidenceBase $root; $excelPid = [uint32]$compile.excel_pid; $environment = $compile.environment; $preflight = @($compile.preflight)
    } else {
        $excel = New-Object -ComObject Excel.Application; $excel.Visible = $false; $excelPid = Get-ExcelPid $excel; $excelHwnd=[Int64]$excel.Hwnd
        $excelProcess=Get-Process -Id $excelPid -ErrorAction Stop;[void]$excelProcess.Handle
        [uint32]$excelHwndPid=0;[void][NxTask3.NativeWindowApi]::GetWindowThreadProcessId([IntPtr]$excelHwnd,[ref]$excelHwndPid)
        if($excelHwndPid -ne $excelPid){throw 'compile-only Excel instance is not exact-bound'}
        $excelStartedUtc=$excelProcess.StartTime.ToUniversalTime().ToString('o');$excelExecutable=[IO.Path]::GetFullPath($excelProcess.MainModule.FileName)
        $workbooks = $excel.Workbooks; $book = $workbooks.Open($testHost.ArtifactPath); $environment = Get-EnvironmentEvidence $excel; $preflight = Test-Preflight $excel $book
        if (@($preflight | Where-Object { -not $_.passed }).Count -gt 0) { $terminalCode = 10; throw 'ENVIRONMENT/INCOMPLETE preflight failure' }
        $vbe = $excel.VBE; $bars = $vbe.CommandBars; $control = Find-VbeCompileControl $bars; $terminalCode = 22; $compile = Invoke-CompileWithWatcher $excel $control $excelPid ([Int64]$vbe.MainWindow.HWnd) $RunId
        $compile | Add-Member -NotePropertyName excel_started_utc -NotePropertyValue $excelStartedUtc
        $compile | Add-Member -NotePropertyName excel_executable -NotePropertyValue $excelExecutable
        $compile | Add-Member -NotePropertyName excel_hwnd -NotePropertyValue $excelHwnd
    }
    if (@($preflight | Where-Object { -not $_.passed }).Count -gt 0) { $terminalCode = 10; throw 'ENVIRONMENT/INCOMPLETE preflight failure' }
    if ($Mode -eq 'Green') {
        if (-not [string]::IsNullOrWhiteSpace($compile.dialog_text)) { $terminalCode = 22; throw 'Green compile produced a dialog' }
        $terminalCode = 24
        $names = $excel.Application.Run("'" + $book.Name + "'!" + $suiteMapping.list)
        $caseNames = @($names | ForEach-Object { [string]$_ })
        if ($Suite -eq 'Core' -and $caseNames.Count -ne 36) { throw 'Core suite mapping did not resolve exact 36 names' }
        if ($Suite -eq 'Frame' -and $caseNames.Count -ne 28) { throw 'Frame suite mapping did not resolve exact 28 names' }
        if ($Suite -eq 'AI' -and $caseNames.Count -ne 29) { throw 'AI suite mapping did not resolve exact 29 names' }
        if ($Suite -eq 'G005' -and $caseNames.Count -ne 10) { throw 'G005 suite mapping did not resolve exact 10 names' }
        if ($Suite -eq 'Template' -and $caseNames.Count -ne 30) { throw 'Template suite mapping did not resolve exact 30 names' }
        if ($Suite -eq 'Data' -and $caseNames.Count -ne 17) { throw 'Data suite mapping did not resolve exact 17 names' }
        if ($Suite -eq 'Draw' -and $caseNames.Count -ne 10) { throw 'Draw suite mapping did not resolve exact 10 names' }
        if ($Suite -eq 'File' -and $caseNames.Count -ne 11) { throw 'File suite mapping did not resolve exact 11 names' }
        if ($Suite -eq 'Calculator' -and $caseNames.Count -ne 28) { throw 'Calculator suite mapping did not resolve exact 28 names' }
        if ($Suite -eq 'Symbols' -and $caseNames.Count -ne 25) { throw 'Symbols suite mapping did not resolve exact 25 names' }
        # Compile-only Excel must close before the first isolated case starts.
        $book.Close($false); $excel.Quit(); Release-ComObject $control; Release-ComObject $bars; Release-ComObject $vbe; Release-ComObject $book; Release-ComObject $workbooks; Release-ComObject $excel
        $control=$null;$bars=$null;$vbe=$null;$book=$null;$workbooks=$null;$excel=$null
        [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
        $compileOnlyCleanup=Stop-ExactProcessAfterGrace $excelProcess 'compile-only Excel';try{$excelProcess.Dispose()}catch{$cleanupFailures += $_.Exception.Message};$excelProcess=$null;$excelPid=0
        if(-not [string]::IsNullOrWhiteSpace($compileOnlyCleanup)){throw ('compile-only Excel process residue: '+$compileOnlyCleanup)}
        $run = Invoke-IsolatedGreenCases $testHost $suiteMapping $caseNames (Join-Path $evidenceBase 'isolated') $Suite $RunId $(if($null -eq $buildAttestation){$null}else{[string]$buildAttestation.build_evidence_sha256})
        if([int]$run.failed -ne 0){$failedCase=@($run.tests|Where-Object {$_.status -eq 'FAIL'}|Select-Object -First 1);throw ('VBA suite failed: '+[string]$failedCase[0].name+' | '+[string]$failedCase[0].failure)}
        if ($null -eq $run -or [int]$run.total -ne $caseNames.Count -or [int]$run.failed -ne 0 -or [int]$run.passed -ne $caseNames.Count -or [int]$run.skipped -ne 0) { throw ('VBA suite must report exact ' + $caseNames.Count + '/' + $caseNames.Count) }
    }
} catch { $failure = $_.Exception.Message } finally {
    try { if ($book) { $book.Close($false) } } catch { $cleanupFailures += $_.Exception.Message }
    try { if ($excel) { $excel.Quit() } } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $control } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $bars } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $vbe } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $book } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $workbooks } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $excel } catch { $cleanupFailures += $_.Exception.Message }
    $control = $null; $bars = $null; $vbe = $null; $book = $null; $workbooks = $null; $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers(); Start-Sleep -Milliseconds 400
}
if($null -ne $excelProcess){$exactCleanupFailure=Stop-ExactProcessAfterGrace $excelProcess 'compile-only Excel';if(-not [string]::IsNullOrWhiteSpace($exactCleanupFailure)){$cleanupFailures += $exactCleanupFailure};try{$excelProcess.Dispose()}catch{$cleanupFailures += $_.Exception.Message};$excelProcess=$null;$excelPid=0}
$classification = 'ENVIRONMENT/INCOMPLETE'; $status = 'INCOMPLETE'
if ($cleanupFailures.Count -gt 0) {
    $terminalCode = 25
    $failure = if ($null -eq $failure) { 'Cleanup blocked PASS: ' + ($cleanupFailures -join ' | ') } else { $failure + ' | Cleanup blocked PASS: ' + ($cleanupFailures -join ' | ') }
} elseif ($Mode -eq 'RedProbe' -and $null -eq $failure) {
    $expected = Get-ExpectedRedLocation $root
    if (-not [string]::IsNullOrWhiteSpace($compile.dialog_text) -and $compile.dialog_sha256 -eq (Get-StringSha256 $compile.dialog_text) -and $compile.dialog_closed -eq $true -and $compile.dialog_disappeared -eq $true -and $compile.stable_polls -ge 2 -and $compile.active_module -eq $expected.module -and $compile.line -eq $expected.line -and $compile.source_line_sha256 -eq $expected.source_line_sha256) { $classification = 'EXPECTED_RED'; $status = 'EXPECTED_RED' } else { $terminalCode = 23; $failure = 'RedProbe stable dialog body/hash or exact CNxStateGuard code-pane selection did not match' }
} elseif ($Mode -eq 'Green' -and $null -eq $failure) { $classification = 'PASS'; $status = 'PASS' }
$redPath = Join-Path $evidenceBase 'Core.RedProbe.json'
$evidence = [ordered]@{ schema_version = 1; suite = $Suite; mode = $Mode; status = $status; classification = $classification; run_id = if($compile){$compile.run_id}elseif(-not [string]::IsNullOrWhiteSpace($RunId)){$RunId}else{$null}; started_utc = if ($compile) { $compile.watcher_started_utc } else { $null }; finished_utc = [DateTime]::UtcNow.ToString('o'); source_allowlist = if ($testHost) { $testHost.SourceAllowlist }; source_allowlist_digest = if ($testHost) { $testHost.SourceAllowlistDigest }; source_digest = if(-not [string]::IsNullOrWhiteSpace($SourceDigest)){$SourceDigest}elseif($testHost){$testHost.SourceAllowlistDigest}else{$null}; snapshot_digest = if(-not [string]::IsNullOrWhiteSpace($SnapshotDigest)){$SnapshotDigest}elseif($testHost){$testHost.HostSha256}else{$null}; test_host = if ($testHost) { @{ sha256 = $testHost.HostSha256; file_format = $testHost.HostFileFormat } }; build_evidence_sha256=if($null -eq $buildAttestation){$null}else{[string]$buildAttestation.build_evidence_sha256}; environment = $environment; preflight = $preflight; compile = $compile; run = $run; state_baseline = @{ restored = ($cleanupFailures.Count -eq 0); restore_failures = @() }; cleanup = @{ failures = $cleanupFailures; excel_pid = $excelPid }; redprobe_sha256 = if ($Mode -eq 'Green' -and (Test-Path -LiteralPath $redPath)) { Get-Sha256 $redPath } else { $null }; failure = $failure }
if ($Mode -eq 'Green' -and ($status -ne 'PASS' -or [string]::IsNullOrWhiteSpace($evidence.redprobe_sha256))) { $terminalCode = 24; $evidence.status = 'INCOMPLETE'; $evidence.classification = 'ENVIRONMENT/INCOMPLETE' }
Write-AtomicEvidence $evidencePath ([pscustomobject]$evidence)
if ($Mode -eq 'RedProbe' -and $status -eq 'EXPECTED_RED') { exit 1 }
if ($evidence.status -ne 'PASS') { exit $terminalCode }
