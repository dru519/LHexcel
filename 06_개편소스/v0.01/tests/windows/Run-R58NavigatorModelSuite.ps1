param([Parameter(Mandatory=$true)][string]$EvidenceRoot,[string]$RunId='')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop';$SourceRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'));$EvidenceRoot=[IO.Path]::GetFullPath($EvidenceRoot)
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Existing R58 navigator-model evidence root is rejected'};if([string]::IsNullOrWhiteSpace($RunId)){$RunId=[Guid]::NewGuid().ToString().ToLowerInvariant()};[void](New-Item -ItemType Directory -Path $EvidenceRoot)
function Write-Json([string]$Path,[object]$Value){[IO.File]::WriteAllText($Path,(($Value|ConvertTo-Json -Depth 12)+"`n"),(New-Object Text.UTF8Encoding($false)))}
$runtime=[Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory();$csc=Join-Path $runtime 'csc.exe';if(-not(Test-Path $csc)){throw 'NET Framework csc.exe unavailable'}
$pane=Join-Path $SourceRoot 'src/dotnet/NxHost/NavigatorPane.cs';$catalog=Join-Path $SourceRoot 'src/dotnet/NxHost/NxGeneratedNavigatorCatalog.cs';$harness=Join-Path $EvidenceRoot 'R58NavigatorModelHarness.cs'
$code=@'
using System;using System.Reflection;using System.Windows.Forms;using LH.NxHost;
sealed class H:INavigatorRequestHandler{public int ExecuteCount,CloseCount,AvailabilityCount;public bool ExecuteRequest(string f,string c,string p){ExecuteCount++;return true;}public string GetAvailability(string r){AvailabilityCount++;return "available|\uC900\uBE44\uB428";}public void CloseNavigator(){CloseCount++;}}
static class P{static object F(object o,string n){return o.GetType().GetField(n,BindingFlags.Instance|BindingFlags.NonPublic).GetValue(o);}static void Key(object p,Keys k){p.GetType().GetMethod("KeyPressed",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(p,new object[]{p,new KeyEventArgs(k)});}static string Route(object item){return (string)item.GetType().GetProperty("RouteKey",BindingFlags.Instance|BindingFlags.NonPublic).GetValue(item,null);} [STAThread]static int Main(){try{var p=new NavigatorPane();p.CreateControl();var h=new H();p.SetRequestHandler(h);var r=(ListBox)F(p,"results");r.SelectedIndex=Math.Min(20,r.Items.Count-1);r.TopIndex=Math.Min(10,r.Items.Count-1);var route=Route(r.SelectedItem);var top=r.TopIndex;var before=h.AvailabilityCount;for(int i=0;i<500;i++)p.RefreshPending();for(int i=0;i<500;i++)Application.DoEvents();if(Route(r.SelectedItem)!=route||r.TopIndex!=top)throw new Exception("selection/top-index not preserved");if(h.AvailabilityCount-before>2)throw new Exception("refresh was not coalesced");var korean=NxGeneratedNavigatorCatalog.DisplayText("active_workbook");if(korean=="active_workbook"||korean.IndexOf('\uD1B5')<0)throw new Exception("DisplayText leaked raw token");Key(p,Keys.Enter);Key(p,Keys.Escape);if(h.ExecuteCount!=1||h.CloseCount!=1)throw new Exception("keyboard did not route once");p.Dispose();p.RefreshPending();Console.WriteLine("PASS");return 0;}catch(Exception e){Console.Error.WriteLine(e);return 1;}}}
'@
[IO.File]::WriteAllText($harness,$code,(New-Object Text.UTF8Encoding($false)));$results=@()
foreach($platform in @('x86','anycpu')){
    $exe=Join-Path $EvidenceRoot ('R58NavigatorModel.'+$platform+'.exe')
    $compile=& $csc /nologo /codepage:65001 /target:exe /platform:$platform /out:$exe /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll $pane $catalog $harness 2>&1
    $compileExitCode=$LASTEXITCODE;$compiled=($compileExitCode -eq 0);$run=$null;$runExitCode=$null
    $runStatus=if($platform -eq 'anycpu'){'NOT_RUN'}else{'NOT_REQUIRED'}
    if($compiled -and $platform -eq 'anycpu'){
        $run=& $exe 2>&1;$runExitCode=$LASTEXITCODE
        $runStatus=if($runExitCode -eq 0 -and ($run -join "`n") -match 'PASS'){'PASS'}else{'FAIL'}
    }
    $results+=[pscustomobject][ordered]@{platform=$platform;compile_status=if($compiled){'PASS'}else{'FAIL'};compile_exit_code=$compileExitCode;run_status=$runStatus;run_exit_code=$runExitCode;output=@($compile+$run)}
}
$anyCpu=@($results|Where-Object platform -eq 'anycpu'|Select-Object -First 1)
$status=if(@($results|Where-Object compile_status -ne 'PASS').Count -eq 0 -and $anyCpu.Count -eq 1 -and $anyCpu[0].run_status -eq 'PASS'){'PASS'}else{'DIAGNOSTIC'}
Write-Json (Join-Path $EvidenceRoot 'R58NavigatorModel.json') ([ordered]@{schema_version=1;suite='R58NavigatorModel';status=$status;run_id=$RunId;machine_architecture=$env:PROCESSOR_ARCHITECTURE;machine_architecture_wow64=$env:PROCESSOR_ARCHITEW6432;is_64bit_os=[Environment]::Is64BitOperatingSystem;navigator_sha256=(Get-FileHash $pane -Algorithm SHA256).Hash.ToLowerInvariant();catalog_sha256=(Get-FileHash $catalog -Algorithm SHA256).Hash.ToLowerInvariant();results=$results;completed_utc=[DateTime]::UtcNow.ToString('o')})
if ($status -eq 'PASS'){Write-Output 'PASS|R58NavigatorModel';exit 0};Write-Output 'DIAGNOSTIC|R58NavigatorModel';throw ('R58 navigator model failed; receipt='+ (Join-Path $EvidenceRoot 'R58NavigatorModel.json'))
