function Test-R75InstalledIntegrity([object]$Application,[string]$Macro,[string]$EvidenceRoot) {
    $proof=[ordered]@{normal='NOT_RUN';missing='NOT_RUN';changed='NOT_RUN';restored='NOT_RUN';hwpx_export='NOT_RUN'}
    $root=Join-Path $EvidenceRoot 'runtime-integrity'
    [void](New-Item -ItemType Directory -Path $root)
    # The loader intentionally rejects UNC CodeBase paths before inspecting a DLL.
    # Use local candidates so these cases exercise missing-file and hash rejection.
    $candidateRoot=Join-Path ([IO.Path]::GetTempPath()) ('nx-integrity-'+[Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $candidateRoot)
    $view=if([string]$Application.OperatingSystem -match '64'){[Microsoft.Win32.RegistryView]::Registry64}else{[Microsoft.Win32.RegistryView]::Registry32}
    $registry=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
    $key=$registry.OpenSubKey('Software\Classes\CLSID\{1434C649-18AA-4435-B0D2-2BD81B548A01}\InprocServer32',$true)
    if($null -eq $key){$registry.Dispose();throw 'Owned installed HWPX registration missing'}
    $original=[string]$key.GetValue('CodeBase');$kind=$key.GetValueKind('CodeBase')
    $dll=([Uri]$original).LocalPath
    $before=(Get-FileHash -LiteralPath $dll).Hash
    $proof.dll_sha256=$before
    $service=$null;$book=$null;$sheet=$null;$range=$null;$books=$null
    try {
        foreach($factory in @('NxHostCreateHwpxService','NxHostCreatePicturePreview','NxHostCreateWorkbookCompare')) {
            $service=$Application.Run($Macro+$factory)
            if($null -eq $service){throw ('Protected service connection failed: '+$factory+'; '+[string]$Application.Run($Macro+'NxHostIntegrityLastError'))}
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($service);$service=$null
        }
        $proof.normal='PASS'
        foreach($case in @('missing','changed')) {
            $candidate=Join-Path $candidateRoot ($case+'.dll')
            if($case -eq 'changed') {
                $bytes=[IO.File]::ReadAllBytes($dll);$bytes[$bytes.Length-1]=$bytes[$bytes.Length-1] -bxor 1
                [IO.File]::WriteAllBytes($candidate,$bytes)
            }
            # Change only this test installation's CodeBase, restoring it in finally.
            $key.SetValue('CodeBase',([Uri]$candidate).AbsoluteUri,$kind)
            $service=$Application.Run($Macro+'NxHostCreateHwpxService')
            if($null -ne $service){throw ('Invalid DLL connected: '+$case)}
            $detail=[string]$Application.Run($Macro+'NxHostIntegrityLastError')
            $proof[($case+'_candidate')]=$candidate
            $proof[($case+'_detail')]=$detail
            if(-not $detail.Contains($candidate)){throw 'DLL rejection omitted the affected path'}
            if(-not [bool]$Application.Ready){throw 'Excel did not remain usable after DLL rejection'}
            $proof[$case]='PASS'
        }
        $key.SetValue('CodeBase',$original,$kind)
        $service=$Application.Run($Macro+'NxHostCreateHwpxService')
        if($null -eq $service){throw 'Restored DLL connection failed'}
        $proof.restored='PASS'
        # Produce a real native HWPX using an Excel-captured fixture, without opening Hancom.
        $books=$Application.Workbooks;$book=$books.Add();$sheet=$book.Worksheets.Item(1)
        $range=$sheet.Range('A1:B2');$range.Value2='r75';$book.Saved=$true
        $payload=[string]$Application.Run($Macro+'NxHwpxSerialize',$range,$true,$false,'display')
        $template=[string]$Application.Run($Macro+'NxHwpxEnsureEmbeddedTemplate')
        $job=[string]$service.Start($payload,$template)
        $watch=[Diagnostics.Stopwatch]::StartNew()
        do {
            $status=[string]$service.GetStatus($job)
            if($status -ne 'running'){break}
            Start-Sleep -Milliseconds 25
        } while($watch.Elapsed.TotalSeconds -lt 60)
        if($status -ne 'succeeded'){$service.Cancel($job);throw ('Installed HWPX export failed: '+[string]$service.GetError($job))}
        $output=[string]$service.GetResult($job)
        if(-not (Test-Path -LiteralPath $output -PathType Leaf) -or -not [bool]$book.Saved -or [string]$range.Cells.Item(1,1).Value2 -ne 'r75'){throw 'HWPX output or source preservation failed'}
        Copy-Item -LiteralPath $output -Destination (Join-Path $root 'native.hwpx')
        $proof.hwpx_export='PASS';$proof.hwpx_sha256=(Get-FileHash -LiteralPath $output).Hash
    } finally {
        $key.SetValue('CodeBase',$original,$kind);$key.Dispose();$registry.Dispose()
        if($null -ne $book){$book.Close($false)}
        foreach($item in @($range,$sheet,$book,$books,$service)) {
            if($null -ne $item -and [Runtime.InteropServices.Marshal]::IsComObject($item)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($item)}
        }
        if((Get-FileHash -LiteralPath $dll).Hash -ne $before){throw 'Installed DLL changed during probe'}
        [IO.File]::WriteAllText((Join-Path $root 'integrity.json'),($proof|ConvertTo-Json -Depth 6),(New-Object Text.UTF8Encoding($false)))
    }
    return $proof
}
