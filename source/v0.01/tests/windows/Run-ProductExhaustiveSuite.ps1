param(
    [Parameter(Mandatory=$true)][string]$SourceRoot,
    [Parameter(Mandatory=$true)][string]$ArtifactPath,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$RunId,
    [Parameter(Mandatory=$true)][string]$SourceDigest,
    [Parameter(Mandatory=$true)][string]$SnapshotDigest,
    [string[]]$RouteIds,
    [string]$NxHostRoot = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }

$SourceRoot = [IO.Path]::GetFullPath($SourceRoot)
$ArtifactPath = [IO.Path]::GetFullPath($ArtifactPath)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

$terminalCode = 0
$failureMessage = $null
$failurePhase = 'preflight'
$cleanupFailures = @()
$cleanupExitMode = 'NOT_STARTED'
$artifactSha = $null
$verifierSha = $null
$results = New-Object 'Collections.Generic.List[object]'
$caseProcessReceipts = New-Object 'Collections.Generic.List[object]'
$excel = $null; $books = $null; $addin = $null; $binding = $null; $process = $null
$interactiveLaunchProcess = $null
$comAddins = $null; $hostAddin = $null
$clipboardSnapshot = $null; $clipboardSnapshotCaptured = $false
$profileOverrideWasPresent = Test-Path Env:LHEXCEL_PROFILE_ROOT
$originalProfileOverride = if ($profileOverrideWasPresent) { [string]$env:LHEXCEL_PROFILE_ROOT } else { $null }
$isolatedProfileRoot = $null
$nativeContractPath = Join-Path $SourceRoot 'contracts/exhaustive-native-contract.json'
$expectedRouteCount = [int]((Get-Content -LiteralPath $nativeContractPath -Raw -Encoding UTF8 | ConvertFrom-Json).route_count)
$routeSelectionRequested = $PSBoundParameters.ContainsKey('RouteIds')
$runScope = 'FULL'
$requestedRouteIds = @()
$selectedCases = @()
$executionRouteCount = $expectedRouteCount
$resultFileName = 'ProductExhaustive.json'
$formalReleaseGate = 'NOT_EVALUATED'
if ($routeSelectionRequested) {
    $runScope = 'TARGETED'
    $resultFileName = 'ProductExhaustive.Targeted.json'
    $executionRouteCount = @($RouteIds).Count
}
$stateSampleLimit = 32
$stateBulkCellLimit = 4096
$fixtureBase = $null; $workRoot = $null; $fixtureRootCleanup = 'NOT_STARTED'
$probeSource = Join-Path $SourceRoot 'tests/windows/NxExhaustiveUiProbe.cs'
$navigatorHostRequired = $false
$navigatorHostRegistrationAttempted = $false
$navigatorHostRegistered = $false
$navigatorHostConnected = $false
$navigatorHostRegisterScript = $null
$navigatorHostStatePath = $null
$navigatorHostX86 = $null
$navigatorHostX64 = $null
$navigatorHostReceipt = [ordered]@{
    required=$false;connected=$false;nxhost32_sha256=$null;nxhost64_sha256=$null
    registry_snapshot_before=$null;registry_snapshot_after=$null;status='NOT_REQUIRED'
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
    }
}

function Get-Sha256([string]$Path) {
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
}

function Get-TextSha256([string]$Value) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value)))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Write-AtomicJson([string]$Path, [object]$Value) {
    $directory = Split-Path -Parent $Path
    [void](New-Item -ItemType Directory -Path $directory -Force)
    $temporary = Join-Path $directory (([IO.Path]::GetFileName($Path)) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $backup = $Path + '.previous'
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($temporary, ($Value | ConvertTo-Json -Depth 50), $utf8)
    if (Test-Path -LiteralPath $backup) { Remove-Item -LiteralPath $backup -Force }
    if (Test-Path -LiteralPath $Path) { [IO.File]::Replace($temporary, $Path, $backup, $true); Remove-Item -LiteralPath $backup -Force }
    else { [IO.File]::Move($temporary, $Path) }
}

function Assert-Condition([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Wait-NoExcelProcesses([string]$Label, [int]$TimeoutMs = 10000) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $lastProcessIds = @()
    do {
        $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
        try {
            $lastProcessIds = @($processes | ForEach-Object { [int]$_.Id })
            if ($lastProcessIds.Count -eq 0) { return }
        } finally {
            foreach ($candidate in $processes) { try { $candidate.Dispose() } catch { } }
        }
        if ($watch.ElapsedMilliseconds -lt $TimeoutMs) { Start-Sleep -Milliseconds 100 }
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    throw ('ENVIRONMENT_MISMATCH: EXCEL remained ' + $Label + '; pids=' + ($lastProcessIds -join ','))
}

function Resolve-ExcelExecutable {
    $subKey = 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64,[Microsoft.Win32.RegistryView]::Registry32)) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
        $key = $null
        try {
            $key = $base.OpenSubKey($subKey, $false)
            if ($null -ne $key) {
                $candidate = [string]$key.GetValue('')
                if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
                    return [IO.Path]::GetFullPath($candidate)
                }
            }
        } finally {
            if ($null -ne $key) { $key.Dispose() }
            $base.Dispose()
        }
    }
    throw 'Excel executable is not registered in App Paths'
}

function Resolve-NxHostDll([string]$Root, [string]$Architecture, [string]$FileName) {
    foreach ($candidate in @(
        (Join-Path $Root $FileName),
        (Join-Path $Root (Join-Path $Architecture $FileName)),
        (Join-Path $Root (Join-Path 'nxhost' (Join-Path $Architecture $FileName)))
    )) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return [IO.Path]::GetFullPath($candidate) }
    }
    throw ('NxHost DLL is missing: ' + $Architecture + '/' + $FileName)
}

function Get-AddinMacro([object]$Addin, [string]$Procedure) {
    $escapedName = ([string]$Addin.Name).Replace("'", "''")
    return ("'{0}'!{1}" -f $escapedName, $Procedure)
}

function Get-CommandMap([object]$Contract) {
    $map = @{}
    foreach ($row in $Contract.commands) { $map[[string]$row.id] = $row }
    return $map
}

function Get-FeatureMap([object]$Contract) {
    $map = @{}
    foreach ($row in $Contract.features) { $map[[string]$row.id] = $row }
    return $map
}

function Read-ClipboardUnicode() {
    return [NxExhaustiveUiProbe]::ReadClipboardText()
}

function Write-ClipboardUnicode([string]$Value) {
    [NxExhaustiveUiProbe]::WriteClipboardText($Value)
}

function Set-Cell([object]$Sheet, [string]$Address, [object]$Value) {
    $cell = $null
    try {
        $cell = $Sheet.Range($Address)
        if ($Value -is [DateTime]) {
            $cell.Value2 = [double]([DateTime]$Value).ToOADate()
        } elseif ($Value -is [ValueType]) {
            $cell.Value2 = [double]$Value
        } elseif ($null -eq $Value) {
            $cell.Value2 = $null
        } else {
            $cell.Value2 = [string]$Value
        }
    }
    finally { Release-ComObject $cell }
}

function Set-Formula([object]$Sheet, [string]$Address, [string]$Formula) {
    $cell = $null
    try { $cell = $Sheet.Range($Address); $cell.Formula = $Formula }
    finally { Release-ComObject $cell }
}

function Select-Range([object]$Sheet, [string]$Address) {
    $range = $null
    try { [void]$Sheet.Activate(); $range = $Sheet.Range($Address); [void]$range.Select() }
    finally { Release-ComObject $range }
}

function Clear-Range([object]$Sheet, [string]$Address) {
    $range = $null
    try { $range = $Sheet.Range($Address); [void]$range.ClearContents() }
    finally { Release-ComObject $range }
}

function Initialize-Fixture([object]$Excel, [object]$Books, [string]$CaseDirectory, [string]$FixtureKind) {
    [void](New-Item -ItemType Directory -Path $CaseDirectory -Force)
    $book = $Books.Add()
    while ([int]$book.Worksheets.Count -lt 3) { [void]$book.Worksheets.Add() }
    $sheet = $book.Worksheets.Item(1); $sheet.Name = 'Data'
    $second = $book.Worksheets.Item(2); $second.Name = 'Second'
    $last = $book.Worksheets.Item(3); $last.Name = 'Last'
    try {
        Set-Cell $sheet 'A1' 'Category'; Set-Cell $sheet 'B1' 'Value'; Set-Cell $sheet 'C1' 'Date'; Set-Cell $sheet 'D1' 'Text'
        Set-Cell $sheet 'A2' 'A'; Set-Cell $sheet 'B2' 10; Set-Cell $sheet 'C2' ([datetime]'1990-01-01'); Set-Cell $sheet 'D2' 'alpha'
        Set-Cell $sheet 'A3' 'B'; Set-Cell $sheet 'B3' 20; Set-Cell $sheet 'C3' ([datetime]'2000-02-02'); Set-Cell $sheet 'D3' 'beta'
        Set-Cell $sheet 'A4' 'A'; Set-Cell $sheet 'B4' 30; Set-Cell $sheet 'C4' ([datetime]'2010-03-03'); Set-Cell $sheet 'D4' 'alpha'
        Set-Cell $sheet 'A5' 'C'; Set-Cell $sheet 'B5' 40; Set-Cell $sheet 'C5' ([datetime]'2020-04-04'); Set-Cell $sheet 'D5' 'gamma'
        Set-Formula $sheet 'E2' '=B2*2'; Set-Formula $sheet 'E3' '=B3*2'; Set-Formula $sheet 'E4' '=SUM(B2:B3)'
        Set-Cell $sheet 'F2' '900101-1234567'; Set-Cell $sheet 'G2' '1,234'
        Set-Cell $sheet 'T1' 101; Set-Formula $sheet 'T2' '=T1+1'; Set-Cell $sheet 'T3' 103
        $sheet.Rows.Item(6).Hidden = $true
        $sheet.Columns.Item(10).Hidden = $true
        $name = $book.Names.Add('NxHiddenName', '=Data!$A$1'); $name.Visible = $false; Release-ComObject $name
        $commentCell = $sheet.Range('D2'); [void]$commentCell.AddComment('native memo'); Release-ComObject $commentCell
        try { $style = $book.Styles.Add('NxNativeStyle'); $style.Font.Bold = $true; Release-ComObject $style } catch { }
        if ($FixtureKind -ceq 'drawing') {
            Add-Type -AssemblyName System.Drawing
            $picturePath = Join-Path $CaseDirectory 'native-picture.png'
            $bitmap = $null; $graphics = $null
            try {
                $bitmap = New-Object System.Drawing.Bitmap 8, 8
                $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
                $graphics.Clear([System.Drawing.Color]::FromArgb(31, 111, 235))
                $bitmap.Save($picturePath, [System.Drawing.Imaging.ImageFormat]::Png)
            } finally {
                if ($null -ne $graphics) { $graphics.Dispose() }
                if ($null -ne $bitmap) { $bitmap.Dispose() }
            }
            $shapes = $sheet.Shapes
            $shape = $shapes.AddPicture($picturePath, 0, -1, 80, 80, 60, 30)
            $shape.Name = 'NxNativeShape'
            Release-ComObject $shape; Release-ComObject $shapes
        }
        if ($FixtureKind -in @('drawing','file')) {
            $chartObjects = $sheet.ChartObjects()
            $chartObject = $chartObjects.Add(220, 40, 180, 120)
            $chart = $chartObject.Chart
            $chartRange = $sheet.Range('A1:B5')
            [void]$chart.SetSourceData($chartRange)
            Release-ComObject $chartRange; Release-ComObject $chart; Release-ComObject $chartObject; Release-ComObject $chartObjects
        }
        Set-Cell $second 'A1' 'second'; Set-Cell $last 'A1' 'last'
        $path = Join-Path $CaseDirectory 'fixture.xlsx'
        $book.SaveAs($path, 51)
        Select-Range $sheet 'B2:D4'
        return [pscustomobject]@{book=$book;sheet=$sheet;second=$second;last=$last;path=$path}
    } catch {
        Release-ComObject $last; Release-ComObject $second; Release-ComObject $sheet
        try { $book.Close($false) } catch { }
        Release-ComObject $book
        throw
    }
}

function Select-PivotDetailCell([object]$Fixture) {
    $source = $null; $caches = $null; $cache = $null; $destination = $null
    $table = $null; $rowField = $null; $valueField = $null; $dataField = $null
    $body = $null; $cell = $null
    try {
        $source = $Fixture.sheet.Range('A1:B5')
        $caches = $Fixture.book.PivotCaches()
        $cache = $caches.Create(1, $source)
        $destination = $Fixture.last.Range('A1')
        $table = $cache.CreatePivotTable($destination, 'NxNativePivot')
        $rowField = $table.PivotFields('Category')
        $rowField.Orientation = 1
        $rowField.Position = 1
        $valueField = $table.PivotFields('Value')
        $dataField = $table.AddDataField($valueField, 'Sum of Value', -4157)
        [void]$table.RefreshTable()
        $body = $table.DataBodyRange
        if ($null -eq $body) { throw 'native pivot fixture has no detail-enabled data cell' }
        $cell = $body.Cells.Item(1, 1)
        [void]$Fixture.last.Activate()
        [void]$cell.Select()
    } finally {
        Release-ComObject $cell; Release-ComObject $body; Release-ComObject $dataField
        Release-ComObject $valueField; Release-ComObject $rowField; Release-ComObject $table
        Release-ComObject $destination; Release-ComObject $cache; Release-ComObject $caches
        Release-ComObject $source
    }
}

function Get-RouteInput([string]$Key) {
    switch ($Key) {
        'RB_SELECTION_CELL_TEXT_SPECIFIC' { return 'alpha' }
        'RB_EDIT_CELL_RESIZE' { return 'width=12;height=18;font=10' }
        'RB_FILTER_AUTOFILTER_FILTERING_OPTIONAL' { return 'A' }
        'RB_EDIT_ALIGN_CENTER_OVERCELLS' { return '$B$2:$D$2' }
        default { return 'native-test' }
    }
}

function Get-ExpectedVisibleClipboardText([object]$Range) {
    $rows = New-Object 'Collections.Generic.List[string]'
    for ($rowIndex=1; $rowIndex -le [int]$Range.Rows.Count; $rowIndex++) {
        $values = New-Object 'Collections.Generic.List[string]'
        for ($columnIndex=1; $columnIndex -le [int]$Range.Columns.Count; $columnIndex++) {
            $cell = $null; $entireRow = $null; $entireColumn = $null
            try {
                $cell = $Range.Cells.Item($rowIndex,$columnIndex)
                $entireRow = $cell.EntireRow; $entireColumn = $cell.EntireColumn
                if (-not [bool]$entireRow.Hidden -and -not [bool]$entireColumn.Hidden) {
                    [void]$values.Add([string]$cell.Text)
                }
            } finally { Release-ComObject $entireColumn; Release-ComObject $entireRow; Release-ComObject $cell }
        }
        if ($values.Count -gt 0) { [void]$rows.Add(($values -join "`t")) }
    }
    return ($rows -join "`r`n")
}

function Resolve-RouteOracleExpected([object]$Oracle, [object]$Fixture) {
    $operator = [string]$Oracle.operator
    if ($operator -ceq 'fixture-reference') {
        $cell = $null
        try {
            $cell = $Fixture.sheet.Range([string]$Oracle.expected)
            return [string]$cell.Address($true,$true,1,$true)
        } finally { Release-ComObject $cell }
    }
    if ($operator -ceq 'fixture-visible-text') {
        $range = $null
        try {
            $range = $Fixture.sheet.Range([string]$Oracle.expected)
            return Get-ExpectedVisibleClipboardText $range
        } finally { Release-ComObject $range }
    }
    return $Oracle.expected
}

function Test-RangeContains([object]$Excel, [object]$Outer, [object]$Inner) {
    $outerRows = $null; $outerColumns = $null; $innerRows = $null; $innerColumns = $null
    try {
        if ($null -eq $Outer -or $null -eq $Inner) { return $false }
        $outerRows = $Outer.Rows; $outerColumns = $Outer.Columns
        $innerRows = $Inner.Rows; $innerColumns = $Inner.Columns
        $outerLastRow = [int64]$Outer.Row + [int64]$outerRows.Count - 1
        $outerLastColumn = [int64]$Outer.Column + [int64]$outerColumns.Count - 1
        $innerLastRow = [int64]$Inner.Row + [int64]$innerRows.Count - 1
        $innerLastColumn = [int64]$Inner.Column + [int64]$innerColumns.Count - 1
        return (
            [int64]$Inner.Row -ge [int64]$Outer.Row -and
            [int64]$Inner.Column -ge [int64]$Outer.Column -and
            $innerLastRow -le $outerLastRow -and
            $innerLastColumn -le $outerLastColumn
        )
    } catch { return $false }
    finally {
        Release-ComObject $innerColumns; Release-ComObject $innerRows
        Release-ComObject $outerColumns; Release-ComObject $outerRows
    }
}

function Get-RouteOracleValue([object]$Excel, [object]$Fixture, [object]$Oracle) {
    $property = [string]$Oracle.property
    switch ($property) {
        'clipboard.text' { return Read-ClipboardUnicode }
        'application.move_after_return_direction' { return [int]$Excel.MoveAfterReturnDirection }
        'application.fullscreen' { return [bool]$Excel.DisplayFullScreen }
        'active_sheet.name' { return [string]$Excel.ActiveSheet.Name }
        'selection.address' {
            $selection = $null
            try { $selection = $Excel.Selection; return [string]$selection.Address($true,$true,1,$false) }
            finally { Release-ComObject $selection }
        }
        'window.scroll_row' { return [int]$Excel.ActiveWindow.ScrollRow }
        'window.scroll_column' { return [int]$Excel.ActiveWindow.ScrollColumn }
        'window.zoom' { return [int]$Excel.ActiveWindow.Zoom }
        'worksheet.filter_mode' { return [bool]$Fixture.sheet.FilterMode }
        'selection.font_size' {
            $selection=$null; $format=$null
            try {$selection=$Excel.Selection; $format=$selection.Font; return $format.Size}
            finally {Release-ComObject $format; Release-ComObject $selection}
        }
        'selection.font_color' {
            $selection=$null; $format=$null
            try {$selection=$Excel.Selection; $format=$selection.Font; return $format.Color}
            finally {Release-ComObject $format; Release-ComObject $selection}
        }
        'selection.fill_color' {
            $selection=$null; $format=$null
            try {$selection=$Excel.Selection; $format=$selection.Interior; return $format.Color}
            finally {Release-ComObject $format; Release-ComObject $selection}
        }
        'worksheet.left_margin_points' {
            $pageSetup = $null
            try { $pageSetup=$Fixture.sheet.PageSetup; return [math]::Round([double]$pageSetup.LeftMargin,2) }
            finally { Release-ComObject $pageSetup }
        }
        'worksheet.print_title_rows' {
            $pageSetup = $null
            try { $pageSetup=$Fixture.sheet.PageSetup; return [string]$pageSetup.PrintTitleRows }
            finally { Release-ComObject $pageSetup }
        }
        'selection.first_column_outline_level' {
            $selection = $null; $cell = $null; $column = $null
            try { $selection=$Excel.Selection; $cell=$selection.Cells.Item(1,1); $column=$cell.EntireColumn; return [int]$column.OutlineLevel }
            finally { Release-ComObject $column; Release-ComObject $cell; Release-ComObject $selection }
        }
        default { throw ('unsupported oracle property: ' + $property) }
    }
}

function Test-RouteOracle([object]$Excel, [object]$Fixture, [object]$Oracle) {
    if ([string]$Oracle.type -cne 'property') { throw ('unsupported oracle type: ' + [string]$Oracle.type) }
    $expected = Resolve-RouteOracleExpected $Oracle $Fixture
    $actual = Get-RouteOracleValue $Excel $Fixture $Oracle
    $operator = [string]$Oracle.operator
    if ($operator -notin @('equals','approx','fixture-reference','fixture-visible-text')) {
        throw ('unsupported oracle operator: ' + $operator)
    }
    if ($operator -ceq 'approx') {
        $toleranceProperty = $Oracle.PSObject.Properties['tolerance']
        if ($null -eq $toleranceProperty -or [double]$toleranceProperty.Value -lt 0) { throw 'approx oracle tolerance is invalid' }
        $passed = [math]::Abs([double]$actual - [double]$expected) -le [double]$toleranceProperty.Value
    } elseif ($expected -is [string] -or $actual -is [string]) {
        $passed = [string]::Equals([string]$expected,[string]$actual,[StringComparison]::Ordinal)
    } else { $passed = ($actual -eq $expected) }
    $reason = if ($passed) { $null } else { 'oracle: expected=' + [string]$expected + '; actual=' + [string]$actual }
    return [pscustomobject]@{passed=[bool]$passed;reason=$reason;expected=$expected;actual=$actual}
}

function Test-TextListContains([object[]]$Values, [string]$Needle) {
    foreach ($value in $Values) {
        if (-not [string]::IsNullOrWhiteSpace([string]$value) -and ([string]$value).IndexOf($Needle,[StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
}

function Test-OwnedWindowOracle([object[]]$Windows, [object]$Oracle) {
    if ([string]$Oracle.type -cne 'owned-window') { throw ('unsupported oracle type: ' + [string]$Oracle.type) }
    if ($Windows.Count -eq 0) { return [pscustomobject]@{passed=$false;reason='popup: no owned window was observed'} }
    foreach ($window in $Windows) {
        foreach ($forbidden in @($Oracle.forbidden_titles)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$forbidden) -and ([string]$window.Title).IndexOf([string]$forbidden,[StringComparison]::OrdinalIgnoreCase) -ge 0) {
                return [pscustomobject]@{passed=$false;reason=('popup: forbidden window observed: ' + [string]$window.Title)}
            }
        }
    }
    $candidate = $null
    foreach ($window in $Windows) {
        $titleMatch = @($Oracle.expected_titles) -ccontains [string]$window.Title
        if (-not $titleMatch) { continue }
        $classMatch = $false
        foreach ($expectedClass in @($Oracle.expected_classes)) {
            if (([string]$window.ClassName).StartsWith([string]$expectedClass,[StringComparison]::OrdinalIgnoreCase)) { $classMatch=$true; break }
        }
        if ($classMatch) { $candidate=$window; break }
    }
    if ($null -eq $candidate) {
        $seen = @($Windows | ForEach-Object { ([string]$_.Title) + '|' + ([string]$_.ClassName) }) -join ', '
        return [pscustomobject]@{passed=$false;reason=('popup: expected owned window identity was absent; seen=' + $seen)}
    }
    $surface = @([string]$candidate.Title) + @($candidate.ChildTexts) + @($candidate.ChildClasses)
    $any = @($Oracle.required_controls_any)
    if ($any.Count -gt 0) {
        $anyMatched = $false
        foreach ($needle in $any) { if (Test-TextListContains $surface ([string]$needle)) { $anyMatched=$true; break } }
        if (-not $anyMatched) { return [pscustomobject]@{passed=$false;reason='popup: no required signature control was observed'} }
    }
    foreach ($needle in @($Oracle.required_controls_all)) {
        if (-not (Test-TextListContains $surface ([string]$needle))) {
            return [pscustomobject]@{passed=$false;reason=('popup: required route discriminator was absent: ' + [string]$needle)}
        }
    }
    return [pscustomobject]@{passed=$true;reason=$null}
}

function Get-ExecutionResultFailure([object[]]$Windows) {
    $failureLabels = @('환경 오류','부분 실패','입력 오류','취소','Environment Error','Partial Failure','Input Error','Cancelled')
    foreach ($window in @($Windows)) {
        if ([string]$window.Title -cne '내엑셀 - 실행결과') { continue }
        $surface = @([string]$window.Title) + @($window.ChildTexts)
        foreach ($failureLabel in $failureLabels) {
            if (Test-TextListContains $surface $failureLabel) { return $failureLabel }
        }
    }
    return $null
}

function Get-CellValue2([object]$Sheet, [string]$Address) {
    $cell = $null
    try { $cell = $Sheet.Range($Address); return $cell.Value2 }
    finally { Release-ComObject $cell }
}

function Get-CellFormula([object]$Sheet, [string]$Address) {
    $cell = $null
    try { $cell = $Sheet.Range($Address); return [string]$cell.Formula }
    finally { Release-ComObject $cell }
}

function Test-CopySaveOutputOracle([object]$Excel, [string]$RouteId, [object[]]$Outputs) {
    $temporary = @($Outputs | Where-Object { $_.Name -match '\.nx-copy-(save|source)\.tmp\.' })
    if ($temporary.Count -gt 0) { return [pscustomobject]@{passed=$false;reason='copy-save output left a temporary file'} }
    $workbooks = @($Outputs | Where-Object { $_.Extension -ieq '.xlsx' })
    if ($workbooks.Count -ne 1) { return [pscustomobject]@{passed=$false;reason='copy-save output must contain exactly one XLSX file'} }
    $outputBook = $null; $outputSheet = $null; $charts = $null; $row = $null; $column = $null
    $commentCell = $null; $comment = $null
    try {
        $outputBook = $Excel.Workbooks.Open($workbooks[0].FullName, 0, $true)
        if ($RouteId -ceq 'NX-FILE-SHEET-COPY-SAVE') {
            if ([int]$outputBook.Sheets.Count -ne 1) { return [pscustomobject]@{passed=$false;reason='copy-save sheet output must contain exactly one sheet'} }
            $outputSheet = $outputBook.Worksheets.Item(1)
            if ([string]$outputSheet.Name -cne 'Data' -or [string](Get-CellValue2 $outputSheet 'A1') -cne 'Category' -or [double](Get-CellValue2 $outputSheet 'B2') -ne 10 -or [string](Get-CellFormula $outputSheet 'E2') -cne '=B2*2') {
                return [pscustomobject]@{passed=$false;reason='copy-save sheet values or formulas were not preserved'}
            }
            $charts = $outputSheet.ChartObjects()
            if ([int]$charts.Count -ne 1) { return [pscustomobject]@{passed=$false;reason='copy-save sheet chart was not preserved'} }
            Release-ComObject $charts; $charts = $null
            $commentCell = $outputSheet.Range('D2'); $comment = $commentCell.Comment
            if ($null -eq $comment -or [string]$comment.Text() -cne 'native memo') { return [pscustomobject]@{passed=$false;reason='copy-save sheet note was not preserved'} }
            Release-ComObject $comment; $comment = $null; Release-ComObject $commentCell; $commentCell = $null
            $row = $outputSheet.Rows.Item(6); $column = $outputSheet.Columns.Item(10)
            if (-not [bool]$row.Hidden -or -not [bool]$column.Hidden) { return [pscustomobject]@{passed=$false;reason='copy-save sheet hidden dimensions were not preserved'} }
        } elseif ($RouteId -ceq 'NX-FILE-RANGE-COPY-SAVE') {
            $outputSheet = $outputBook.Worksheets.Item(1)
            $rangeValuesPreserved = (
                [double](Get-CellValue2 $outputSheet 'A1') -eq 10 -and
                [string](Get-CellValue2 $outputSheet 'C1') -ceq 'alpha' -and
                [double](Get-CellValue2 $outputSheet 'A3') -eq 30 -and
                [string](Get-CellValue2 $outputSheet 'C3') -ceq 'alpha'
            )
            if (-not $rangeValuesPreserved) { return [pscustomobject]@{passed=$false;reason='copy-save range values were not preserved'} }
        }
        return [pscustomobject]@{passed=$true;reason=$null}
    } catch {
        return [pscustomobject]@{passed=$false;reason=('copy-save output inspection failed: ' + $_.Exception.Message)}
    } finally {
        Release-ComObject $comment; Release-ComObject $commentCell; Release-ComObject $column; Release-ComObject $row; Release-ComObject $charts; Release-ComObject $outputSheet
        if ($null -ne $outputBook) { try { $outputBook.Close($false) } catch { } }
        Release-ComObject $outputBook
    }
}

function Get-CaseOutputTarget([object]$Case, [string]$CaseDirectory) {
    $safeId = ([string]$Case.route_id) -replace '[^A-Za-z0-9._-]','_'
    switch ([string]$Case.oracle.format) {
        'xlsx' { return Join-Path $CaseDirectory ('output-' + $safeId + '.xlsx') }
        'png' { return Join-Path $CaseDirectory ('output-' + $safeId + '.png') }
        'pdf' { return $CaseDirectory }
        'settings' { return $CaseDirectory }
        default { throw ('unsupported file-output format: ' + [string]$Case.oracle.format) }
    }
}

function Get-CaseOutputFiles([object]$Case, [string]$CaseDirectory) {
    if ([string]$Case.oracle.format -ceq 'settings') {
        $settingsPath = Join-Path $env:LHEXCEL_PROFILE_ROOT 'Settings/pdf-v1.cfg'
        if (Test-Path -LiteralPath $settingsPath -PathType Leaf) { return @((Get-Item -LiteralPath $settingsPath)) }
        return @()
    }
    return @(Get-ChildItem -LiteralPath $CaseDirectory -File |
        Where-Object { $_.Name -cne 'fixture.xlsx' -and $_.Name -notlike '~$*' } |
        Sort-Object Name)
}

function Get-UInt32BigEndian([byte[]]$Bytes, [int]$Offset) {
    return [uint32]((([uint32]$Bytes[$Offset]) -shl 24) -bor
        (([uint32]$Bytes[$Offset + 1]) -shl 16) -bor
        (([uint32]$Bytes[$Offset + 2]) -shl 8) -bor
        ([uint32]$Bytes[$Offset + 3]))
}

function Test-PngSignatureAndDimensions([string]$Path) {
    try {
        $bytes = [IO.File]::ReadAllBytes($Path)
        $signature = [byte[]](0x89,0x50,0x4e,0x47,0x0d,0x0a,0x1a,0x0a)
        if ($bytes.Length -lt 24) { return [pscustomobject]@{passed=$false;reason='PNG output is truncated'} }
        for ($index=0; $index -lt $signature.Length; $index++) {
            if ($bytes[$index] -ne $signature[$index]) { return [pscustomobject]@{passed=$false;reason='PNG signature is invalid'} }
        }
        $width = Get-UInt32BigEndian $bytes 16
        $height = Get-UInt32BigEndian $bytes 20
        if ($width -lt 1 -or $height -lt 1) { return [pscustomobject]@{passed=$false;reason='PNG dimensions are invalid'} }
        return [pscustomobject]@{passed=$true;reason=$null}
    } catch { return [pscustomobject]@{passed=$false;reason=('PNG inspection failed: ' + $_.Exception.Message)} }
}

function Test-PdfSignature([object[]]$Outputs, [int]$ExpectedFiles) {
    if ($Outputs.Count -ne $ExpectedFiles) {
        return [pscustomobject]@{passed=$false;reason=('PDF output count mismatch: expected=' + $ExpectedFiles + '; actual=' + $Outputs.Count)}
    }
    foreach ($output in $Outputs) {
        if ([string]$output.Extension -ine '.pdf' -or [int64]$output.Length -le 5) {
            return [pscustomobject]@{passed=$false;reason=('PDF output is invalid: ' + [string]$output.Name)}
        }
        $stream = $null
        try {
            $stream = [IO.File]::Open($output.FullName, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
            $buffer = New-Object byte[] 5
            if ($stream.Read($buffer, 0, 5) -ne 5 -or [Text.Encoding]::ASCII.GetString($buffer) -cne '%PDF-') {
                return [pscustomobject]@{passed=$false;reason=('PDF signature is invalid: ' + [string]$output.Name)}
            }
        } finally { if ($null -ne $stream) { $stream.Dispose() } }
    }
    return [pscustomobject]@{passed=$true;reason=$null}
}

function Test-MannerSaveOutputOracle([object]$Excel, [object[]]$Outputs) {
    $workbooks = @($Outputs | Where-Object { $_.Extension -ieq '.xlsx' })
    if ($Outputs.Count -ne 1 -or $workbooks.Count -ne 1) { return [pscustomobject]@{passed=$false;reason='manner-save output must contain exactly one XLSX file'} }
    $header = [IO.File]::ReadAllBytes($workbooks[0].FullName)
    if ($header.Length -lt 4 -or $header[0] -ne 0x50 -or $header[1] -ne 0x4b) { return [pscustomobject]@{passed=$false;reason='manner-save XLSX package signature is invalid'} }
    $outputBook = $null; $outputSheet = $null; $commentCell = $null; $comment = $null
    try {
        $outputBook = $Excel.Workbooks.Open($workbooks[0].FullName, 0, $true)
        if ([int]$outputBook.Worksheets.Count -ne 3) { return [pscustomobject]@{passed=$false;reason='manner-save sheet count was not preserved'} }
        $outputSheet = $outputBook.Worksheets.Item('Data')
        if ([string](Get-CellValue2 $outputSheet 'A1') -cne 'Category' -or [double](Get-CellValue2 $outputSheet 'B2') -ne 10 -or [string](Get-CellFormula $outputSheet 'E2') -cne '=B2*2') {
            return [pscustomobject]@{passed=$false;reason='manner-save values or formulas were not preserved'}
        }
        $commentCell = $outputSheet.Range('D2'); $comment = $commentCell.Comment
        if ($null -ne $comment) { return [pscustomobject]@{passed=$false;reason='manner-save did not remove the approved note'} }
        return [pscustomobject]@{passed=$true;reason=$null}
    } catch { return [pscustomobject]@{passed=$false;reason=('manner-save output inspection failed: ' + $_.Exception.Message)} }
    finally {
        Release-ComObject $comment; Release-ComObject $commentCell; Release-ComObject $outputSheet
        if ($null -ne $outputBook) { try { $outputBook.Close($false) } catch { } }
        Release-ComObject $outputBook
    }
}

function Test-PdfSettingsReadback([string]$SettingsPath, [string]$ExpectedFolder) {
    if (-not (Test-Path -LiteralPath $SettingsPath -PathType Leaf)) { return [pscustomobject]@{passed=$false;reason='PDF settings file was not created'} }
    $settingsFolder = Split-Path -Parent $SettingsPath
    $residue = @(Get-ChildItem -LiteralPath $settingsFolder -File | Where-Object { $_.Name -match '\.(tmp|bak)$' })
    if ($residue.Count -gt 0) { return [pscustomobject]@{passed=$false;reason='PDF settings left temporary or backup residue'} }
    $lines = [IO.File]::ReadAllLines($SettingsPath, [Text.Encoding]::Default)
    $expected = @(
        ('folder=' + $ExpectedFolder),
        'base=fixture',
        'quality=표준 품질'
    )
    if ($lines.Count -ne $expected.Count) { return [pscustomobject]@{passed=$false;reason='PDF settings line count is invalid'} }
    for ($index=0; $index -lt $expected.Count; $index++) {
        if ([string]$lines[$index] -cne [string]$expected[$index]) { return [pscustomobject]@{passed=$false;reason=('PDF settings mismatch at line ' + ($index + 1))} }
    }
    return [pscustomobject]@{passed=$true;reason=$null}
}

function Test-FileOutputOracle([object]$Excel, [object]$Case, [object[]]$Outputs, [string]$CaseDirectory) {
    $format = [string]$Case.oracle.format
    $expectedFiles = [int]$Case.oracle.expected_files
    if ($Outputs.Count -ne $expectedFiles) {
        return [pscustomobject]@{passed=$false;reason=('file-output count mismatch: expected=' + $expectedFiles + '; actual=' + $Outputs.Count)}
    }
    switch ($format) {
        'xlsx' {
            if ([string]$Case.route_id -ceq 'NX-FILE-MANNER-SAVE') { return Test-MannerSaveOutputOracle $Excel $Outputs }
            return Test-CopySaveOutputOracle $Excel ([string]$Case.route_id) $Outputs
        }
        'png' { return Test-PngSignatureAndDimensions $Outputs[0].FullName }
        'pdf' { return Test-PdfSignature $Outputs $expectedFiles }
        'settings' { return Test-PdfSettingsReadback $Outputs[0].FullName $CaseDirectory }
        default { return [pscustomobject]@{passed=$false;reason=('unsupported file-output format: ' + $format)} }
    }
}

function Copy-FileOutputEvidence([string]$SafeId, [object[]]$Outputs) {
    $destinationRoot = Join-Path $EvidenceRoot (Join-Path 'outputs' $SafeId)
    [void](New-Item -ItemType Directory -Path $destinationRoot -Force)
    $evidence = @()
    foreach ($output in $Outputs) {
        $destination = Join-Path $destinationRoot ([string]$output.Name)
        Copy-Item -LiteralPath $output.FullName -Destination $destination
        $copy = Get-Item -LiteralPath $destination
        $evidence += [ordered]@{name=[string]$copy.Name;length=[int64]$copy.Length;sha256=Get-Sha256 $copy.FullName;relative_path=('outputs/' + $SafeId + '/' + [string]$copy.Name)}
    }
    return $evidence
}

function Prepare-Route([object]$Excel, [object]$Addin, [object]$Fixture, [object]$Case, [object]$Command) {
    $sheet = $Fixture.sheet
    $key = if ($null -eq $Command) { '' } else { [string]$Command.args[0] }
    Select-Range $sheet 'B2:D4'

    if ([string]$Case.route_id -in @('NX-AI-IMAGE','NX-DRAW-FIT-PICTURE')) {
        $shape = $sheet.Shapes.Item('NxNativeShape'); [void]$shape.Select(); Release-ComObject $shape
    }
    if ([string]$Case.route_id -ceq 'NX-FILE-CHART-PNG') {
        $charts = $null; $chartObject = $null
        try { $charts=$sheet.ChartObjects(); $chartObject=$charts.Item(1); [void]$chartObject.Select() }
        finally { Release-ComObject $chartObject; Release-ComObject $charts }
    }
    if ([string]$Case.route_id -ceq 'NX-DATA-AGE') { Clear-Range $sheet 'D2:D5'; Select-Range $sheet 'C2:C5' }
    if ([string]$Case.route_id -ceq 'NX-DATA-KOREAN-MONEY') { Clear-Range $sheet 'C2:C5'; Select-Range $sheet 'B2:B5' }
    if ([string]$Case.route_id -ceq 'NX-DATA-PRIVACY-MASK') { Clear-Range $sheet 'G2:G2'; Select-Range $sheet 'F2:F2' }
    if ([string]$Case.route_id -ceq 'NX-DATA-PRIVACY-SCAN') { Select-Range $sheet 'F2:F2' }
    if ([string]$Case.route_id -ceq 'NX-DRAW-CLEAR-INNER') { Select-Range $sheet 'A1:D5' }
    if ([string]$Case.route_id -ceq 'NX-DATA-PASTE-VISIBLE-VALUES') {
        $write = Get-AddinMacro $Addin 'NxClipboardWriteUnicode'; [void]$Excel.Run($write, "native-1`r`nnative-2`r`nnative-3")
        Select-Range $sheet 'K2:K4'
    }
    if ([string]$Case.route_id -ceq 'NX-HANGUL-TABLE-SEND') { Select-Range $sheet 'A1:D5' }
    if ([string]$Case.route_id -ceq 'NX-FILE-PDF-SELECTED-COMBINED') {
        [void]$Fixture.sheet.Select()
        [void]$Fixture.second.Select($false)
    }

    switch ($key) {
        'RB_CLIPBOARD_COPY_BYFORMULA' { Select-Range $sheet 'E2' }
        'RB_CLIPBOARD_COPY_BYREFERENCE' { Select-Range $sheet 'B2' }
        'RB_CLIPBOARD_COPY_BYTEXT' { Select-Range $sheet 'D2:D5' }
        'RB_LHECLIPBOARD_COPYVISIBLE' {
            Set-Cell $sheet 'A6' 'HIDDEN_ROW'; Set-Cell $sheet 'B6' 999; Set-Cell $sheet 'J1' 'HIDDEN_COLUMN'
            Select-Range $sheet 'A1:J6'
        }
        'RB_LHECLIPBOARD_PASTEVISIBLEVALUES' { $write=Get-AddinMacro $Addin 'NxClipboardWriteUnicode'; [void]$Excel.Run($write,"201`r`n202`r`n203"); Select-Range $sheet 'K2:K4' }
        'RB_LHECLIPBOARD_PASTEVISIBLEFORMULAS' { $source=$sheet.Range('T1:T3'); [void]$source.Copy(); Release-ComObject $source; Select-Range $sheet 'K2:K4' }
        'RB_LHECLIPBOARD_PASTEVISIBLEFORMATS' { $source=$sheet.Range('T1:T3'); $source.NumberFormat='0.000'; [void]$source.Copy(); Release-ComObject $source; Select-Range $sheet 'K2:K4' }
        'RB_SELECTION_CELL_MOVE_LEFT' { Select-Range $sheet 'D4' }
        'RB_SELECTION_CELL_MOVE_RIGHT' { Select-Range $sheet 'B4' }
        'RB_SELECTION_CELL_MOVE_DOWN' { Select-Range $sheet 'B2' }
        'RB_SELECTION_CELL_MOVE_UP' { Select-Range $sheet 'B4' }
        'RB_SELECTION_CELL_USEDRANGE' { Set-Cell $sheet 'Z20' 'USED_RANGE_END'; Select-Range $sheet 'B2:D4' }
        'RB_APP_OPTION_EDITDIRECTION_CHANGE' { $Excel.MoveAfterReturnDirection=-4121 }
        'RB_EDIT_ALIGN_CENTER_OVERCELLS' { Select-Range $sheet 'B2:D2' }
        'RB_PRINT_SETUP_REPEAT' { Select-Range $sheet 'B2:D4' }
        'RB_LHEPRINT_ADD1MMSPACER' { $pageSetup=$sheet.PageSetup; $pageSetup.LeftMargin=72; Release-ComObject $pageSetup; Select-Range $sheet 'B2:D4' }
        'RB_EDIT_CELL_MAKEGROUP_HIDDEN_ROWCOLUMN' { Select-Range $sheet 'B2:D4' }
        'RB_FILTER_AUTOFILTER_CANCEL' { $range=$sheet.Range('A1:D5'); [void]$range.AutoFilter(1,'A'); Release-ComObject $range; Select-Range $sheet 'A2' }
        'RB_FILTER_AUTOFILTER_SHOWALL' { $range=$sheet.Range('A1:D5'); [void]$range.AutoFilter(1,'A'); Release-ComObject $range; Select-Range $sheet 'A2' }
        'RB_FILTER_AUTOFILTER_FILTERING' { Select-Range $sheet 'A2' }
        'RB_FILTER_AUTOFILTER_FILTERING_OPTIONAL' { Select-Range $sheet 'A2' }
        'RB_DATA_SORT_ASCENDING' { Select-Range $sheet 'A2' }
        'RB_DATA_SORT_DESCENDING' { Select-Range $sheet 'A2' }
        'RB_DATA_SHOW_SOURCE' { Select-PivotDetailCell $Fixture }
        'RB_FORMULA_PRECEDENTS_LIST' { Select-Range $sheet 'E2:E4' }
        'RB_EDIT_MEMO_ADD_LHEXCELFORMULA' { Select-Range $sheet 'E2:E4' }
        'RB_SHEET_SELECT_HOME' { [void]$Fixture.second.Activate() }
        'RB_SHEET_SELECT_END' { [void]$Fixture.second.Activate() }
        'RB_WINDOWS_VIEW_FULLSCREEN' { $Excel.DisplayFullScreen=$false }
        'RB_WINDOWS_VIEW_SCROLL_DOWN' { $Excel.ActiveWindow.ScrollRow=1 }
        'RB_WINDOWS_VIEW_SCROLL_UP' { $Excel.ActiveWindow.ScrollRow=20 }
        'RB_WINDOWS_VIEW_SCROLL_RIGHT' { $Excel.ActiveWindow.ScrollColumn=1 }
        'RB_WINDOWS_VIEW_SCROLL_LEFT' { $column=$sheet.Columns.Item(10); $column.Hidden=$false; Release-ComObject $column; $Excel.ActiveWindow.ScrollColumn=10 }
        default { }
    }
}

function Convert-StateValueWire([object]$Value) {
    if ($null -eq $Value) { return 'NULL' }
    if ($Value -is [Array]) {
        $items = New-Object 'Collections.Generic.List[string]'
        foreach ($item in $Value) { [void]$items.Add((Convert-StateValueWire $item)) }
        return ('ARRAY[' + ($items -join ',') + ']')
    }
    $typeName = $Value.GetType().FullName
    if ($Value -is [IFormattable]) { $text = $Value.ToString($null,[Globalization.CultureInfo]::InvariantCulture) }
    else { $text = [string]$Value }
    return ($typeName + ':' + $text)
}

function Get-BulkRangeState([object]$Range) {
    if ($null -eq $Range) { return 'NO_RANGE' }
    $identity = [string]$Range.Address($true,$true,1,$true) + '|count=' + [string]$Range.CountLarge
    if ([double]$Range.CountLarge -gt [double]$stateBulkCellLimit) { return ($identity + '|bulk=SKIPPED') }
    return ($identity +
        '|value=' + (Convert-StateValueWire $Range.Value2) +
        '|formula=' + (Convert-StateValueWire $Range.Formula) +
        '|format=' + (Convert-StateValueWire $Range.NumberFormat))
}

function Get-BoundedRangeSample([object]$Range) {
    $samples = New-Object 'Collections.Generic.List[string]'
    if ($null -eq $Range) { return $samples.ToArray() }
    $sampleCount = [math]::Min($stateSampleLimit,[int][math]::Min([double]$Range.CountLarge,[double][int]::MaxValue))
    for ($index=1; $index -le $sampleCount; $index++) {
        $cell = $null; $row = $null; $column = $null; $font = $null
        try {
            $cell = $Range.Cells.Item($index)
            $row = $cell.EntireRow; $column = $cell.EntireColumn
            $font = $cell.Font
            [void]$samples.Add(([string]$cell.Address($true,$true,1,$true) +
                '|formula=' + (Convert-StateValueWire $cell.Formula) +
                '|format=' + (Convert-StateValueWire $cell.NumberFormat) +
                '|align=' + (Convert-StateValueWire $cell.HorizontalAlignment) +
                '|font=' + (Convert-StateValueWire $font.Size) +
                '|row_hidden=' + [string][bool]$row.Hidden +
                '|column_hidden=' + [string][bool]$column.Hidden))
        } finally { Release-ComObject $font; Release-ComObject $column; Release-ComObject $row; Release-ComObject $cell }
    }
    return $samples.ToArray()
}

function Get-WorkbookStateSnapshot([object]$Excel, [object]$Book, [string]$CaseDirectory, [switch]$ExcludeCaseFiles) {
    $sheets = @(); $sheetStates = @(); $selectionAddress = ''; $selectionState = $null; $selectionType = 'none'; $activeSheet = ''; $windowState = @{}
    for ($sheetIndex=1; $sheetIndex -le [int]$Book.Worksheets.Count; $sheetIndex++) {
        $sheet = $null; $usedRange = $null; $shapes = $null; $charts = $null
        try {
            $sheet = $Book.Worksheets.Item($sheetIndex)
            $sheets += [string]$sheet.Name
            $usedRange = $sheet.UsedRange
            $shapes = $sheet.Shapes
            $charts = $sheet.ChartObjects()
            $sheetStates += ([string]$sheet.Name + '|' + (Get-BulkRangeState $usedRange) +
                '|shapes=' + [string][int]$shapes.Count + '|charts=' + [string][int]$charts.Count)
        } finally { Release-ComObject $charts; Release-ComObject $shapes; Release-ComObject $usedRange; Release-ComObject $sheet }
    }
    try { $activeSheet = [string]$Excel.ActiveSheet.Name } catch { }
    $selection = $null
    try {
        $selection = $Excel.Selection
        if ($null -ne $selection) {
            try {
                $selectionAddress = [string]$selection.Address($true,$true,1,$true)
                $selectionType = 'range'
                $selectionState = [ordered]@{bulk=Get-BulkRangeState $selection;sample=Get-BoundedRangeSample $selection}
            } catch {
                $selectionType = 'non-range'
                try { $selectionAddress = [string]$selection.Name } catch { }
            }
        }
    } catch { }
    finally { Release-ComObject $selection }
    try {
        $windowState = [ordered]@{count=[int]$Book.Windows.Count;zoom=[string]$Excel.ActiveWindow.Zoom;grid=[bool]$Excel.ActiveWindow.DisplayGridlines;freeze=[bool]$Excel.ActiveWindow.FreezePanes;scroll_row=[int]$Excel.ActiveWindow.ScrollRow;scroll_column=[int]$Excel.ActiveWindow.ScrollColumn;fullscreen=[bool]$Excel.DisplayFullScreen}
    } catch { }
    $nameStates = @()
    for ($nameIndex=1; $nameIndex -le [int]$Book.Names.Count; $nameIndex++) {
        $name = $null
        try { $name=$Book.Names.Item($nameIndex); $nameStates += ([string]$name.Name + '|' + [string][bool]$name.Visible + '|' + [string]$name.RefersTo) }
        finally { Release-ComObject $name }
    }
    $files = @()
    if (-not $ExcludeCaseFiles -and (Test-Path -LiteralPath $CaseDirectory)) { $files = @(Get-ChildItem -LiteralPath $CaseDirectory -File | Sort-Object Name | ForEach-Object { $_.Name + '|' + $_.Length }) }
    $document = [ordered]@{sheets=$sheets;sheet_states=$sheetStates;active_sheet=$activeSheet;selection_type=$selectionType;selection=$selectionAddress;selection_state=$selectionState;window=$windowState;styles=[int]$Book.Styles.Count;names=$nameStates;saved=[bool]$Book.Saved;files=$files;move_direction=[int]$Excel.MoveAfterReturnDirection}
    return [pscustomobject]@{
        digest=Get-TextSha256 ($document | ConvertTo-Json -Depth 10 -Compress)
        document=$document
    }
}

function Get-WorkbookState([object]$Excel, [object]$Book, [string]$CaseDirectory, [switch]$ExcludeCaseFiles) {
    return [string](Get-WorkbookStateSnapshot $Excel $Book $CaseDirectory -ExcludeCaseFiles:$ExcludeCaseFiles).digest
}

function Get-WorkbookStateDelta([Collections.IDictionary]$Before, [Collections.IDictionary]$After) {
    $keys = New-Object 'Collections.Generic.List[string]'
    foreach ($key in $Before.Keys) { [void]$keys.Add([string]$key) }
    foreach ($key in $After.Keys) { if (-not $keys.Contains([string]$key)) { [void]$keys.Add([string]$key) } }
    $changes = New-Object 'Collections.Generic.List[object]'
    foreach ($key in $keys) {
        $beforeWire = $Before[$key] | ConvertTo-Json -Depth 10 -Compress
        $afterWire = $After[$key] | ConvertTo-Json -Depth 10 -Compress
        if ($beforeWire -cne $afterWire) {
            [void]$changes.Add([ordered]@{field=[string]$key;before=$Before[$key];after=$After[$key]})
        }
    }
    return $changes.ToArray()
}

function Get-ProbeMode([object]$Case, [string]$Key) {
    if ([string]$Case.route_id -ceq 'NX-UTIL-NAVIGATOR') { return 'navigator' }
    if ([string]$Case.expected -ceq 'window-open') { return 'popup' }
    if ([string]$Case.expected -ceq 'file-output') { return 'file-output' }
    if ($Key -ceq 'RB_PRINT_SETUP_QUICK') { return 'print-preview' }
    return 'approve'
}

function Close-UserWorkbooks([object]$Books, [object]$Addin) {
    for ($index=[int]$Books.Count; $index -ge 1; $index--) {
        $book = $null
        try {
            $book = $Books.Item($index)
            if ($null -eq $Addin -or [string]$book.Name -cne [string]$Addin.Name) { $book.Close($false) }
        } catch { }
        finally { Release-ComObject $book }
    }
}

function Get-HancomProcessIds() {
    return @(Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -match '^(Hwp|Hwp64|Hancom)' } |
        ForEach-Object { [int]$_.Id })
}

function Get-NewHancomProcessIds([int[]]$BaselineIds) {
    $baseline = New-Object 'Collections.Generic.HashSet[int]'
    foreach ($baselineId in @($BaselineIds)) { [void]$baseline.Add([int]$baselineId) }
    return @(Get-HancomProcessIds | Where-Object { -not $baseline.Contains([int]$_) })
}

function Wait-NewHancomDocumentWindow(
    [int[]]$BaselineIds,
    [string]$ExpectedPath,
    [int]$TimeoutMs = 30000
) {
    $baseline = New-Object 'Collections.Generic.HashSet[int]'
    foreach ($baselineId in @($BaselineIds)) { [void]$baseline.Add([int]$baselineId) }
    $expectedName = if ([string]::IsNullOrWhiteSpace($ExpectedPath)) { '' } else { [IO.Path]::GetFileNameWithoutExtension($ExpectedPath) }
    $wait = [Diagnostics.Stopwatch]::StartNew()
    do {
        $observed = New-Object 'Collections.Generic.List[int]'
        foreach ($hancomProcess in @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -match '^(Hwp|Hwp64|Hancom)' })) {
            try {
                if ($baseline.Contains([int]$hancomProcess.Id)) { continue }
                $hancomProcess.Refresh()
                if ([int64]$hancomProcess.MainWindowHandle -eq 0) { continue }
                if (-not [string]::IsNullOrWhiteSpace($expectedName) -and
                    ([string]$hancomProcess.MainWindowTitle).IndexOf($expectedName,[StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
                [void]$observed.Add([int]$hancomProcess.Id)
            } catch { }
            finally { try { $hancomProcess.Dispose() } catch { } }
        }
        if ($observed.Count -gt 0) { return @($observed.ToArray()) }
        Start-Sleep -Milliseconds 100
    } while ($wait.ElapsedMilliseconds -lt $TimeoutMs)
    return @()
}

function Close-OwnedHancomProcesses(
    [int[]]$OwnedIds,
    [int[]]$ProtectedIds,
    [int]$TimeoutMs = 30000
) {
    $protected = New-Object 'Collections.Generic.HashSet[int]'
    foreach ($protectedId in @($ProtectedIds)) { [void]$protected.Add([int]$protectedId) }
    $requested = New-Object 'Collections.Generic.HashSet[int]'
    $failures = New-Object 'Collections.Generic.List[string]'
    $uniqueOwned = New-Object 'Collections.Generic.HashSet[int]'
    foreach ($ownedExternalPid in @($OwnedIds)) {
        $ownedExternalPid = [int]$ownedExternalPid
        if (-not $uniqueOwned.Add($ownedExternalPid)) { continue }
        if ($protected.Contains($ownedExternalPid)) {
            [void]$failures.Add('Protected Hancom PID was incorrectly marked owned: ' + $ownedExternalPid)
        }
    }

    $lastCloseFailure = @{}
    $wait = [Diagnostics.Stopwatch]::StartNew()
    do {
        $activeCount = 0
        foreach ($ownedExternalPid in @($uniqueOwned)) {
            $ownedExternalPid = [int]$ownedExternalPid
            if ($protected.Contains($ownedExternalPid)) { continue }
            $ownedProcess = $null
            try {
                $ownedProcess = Get-Process -Id $ownedExternalPid -ErrorAction SilentlyContinue
                if ($null -eq $ownedProcess) {
                    [void]$lastCloseFailure.Remove($ownedExternalPid)
                    continue
                }
                $activeCount++
                if ([string]$ownedProcess.ProcessName -notmatch '^(Hwp|Hwp64|Hancom)') {
                    $lastCloseFailure[$ownedExternalPid] = 'Owned PID identity changed before Hancom cleanup'
                    continue
                }
                $closeRequested = [bool]$ownedProcess.CloseMainWindow()
                if ($closeRequested) { [void]$requested.Add($ownedExternalPid) }
                $remainingMs = [math]::Max(0, $TimeoutMs - [int]$wait.ElapsedMilliseconds)
                $waitSlice = [math]::Min(250, $remainingMs)
                if ($waitSlice -gt 0 -and $ownedProcess.WaitForExit($waitSlice)) {
                    [void]$lastCloseFailure.Remove($ownedExternalPid)
                }
            } catch {
                $lastCloseFailure[$ownedExternalPid] = $_.Exception.Message
            } finally {
                if ($null -ne $ownedProcess) { try { $ownedProcess.Dispose() } catch { } }
            }
        }
        if ($activeCount -gt 0 -and $wait.ElapsedMilliseconds -lt $TimeoutMs) { Start-Sleep -Milliseconds 100 }
    } while ($activeCount -gt 0 -and $wait.ElapsedMilliseconds -lt $TimeoutMs)
    $wait.Stop()

    $remaining = New-Object 'Collections.Generic.List[int]'
    foreach ($ownedExternalPid in @($uniqueOwned)) {
        $ownedExternalPid = [int]$ownedExternalPid
        if ($protected.Contains($ownedExternalPid)) { continue }
        if ($null -ne (Get-Process -Id $ownedExternalPid -ErrorAction SilentlyContinue)) {
            [void]$remaining.Add($ownedExternalPid)
            $detail = if ($lastCloseFailure.ContainsKey($ownedExternalPid)) { ': ' + [string]$lastCloseFailure[$ownedExternalPid] } else { '' }
            [void]$failures.Add('Hancom process did not exit after graceful close retries for PID ' + $ownedExternalPid + $detail)
        }
    }
    $status = if ($remaining.Count -eq 0 -and $failures.Count -eq 0) { 'PASS' } else { 'FAIL' }
    return [ordered]@{
        status=$status
        requested_pids=@($requested | Sort-Object)
        remaining_pids=$remaining.ToArray()
        failures=$failures.ToArray()
    }
}

function Remove-FixtureRootWithRetry([string]$Path, [string]$Boundary, [int]$TimeoutMs = 10000) {
    $resolvedBoundary = [IO.Path]::GetFullPath($Boundary).TrimEnd('\')
    $resolvedPath = [IO.Path]::GetFullPath($Path)
    if (-not $resolvedPath.StartsWith($resolvedBoundary + '\',[StringComparison]::OrdinalIgnoreCase)) {
        throw 'Local fixture cleanup escaped its boundary'
    }
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $lastFailure = 'fixture cleanup was not attempted'
    do {
        try {
            if (Test-Path -LiteralPath $resolvedPath) {
                Remove-Item -LiteralPath $resolvedPath -Recurse -Force -ErrorAction Stop
            }
            if (-not (Test-Path -LiteralPath $resolvedPath)) { return }
            $lastFailure = 'Local fixture root remained after cleanup'
        } catch { $lastFailure = $_.Exception.Message }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    throw ('Local fixture root remained after bounded cleanup: ' + $lastFailure)
}

function Invoke-NativeCase([object]$Case, [object]$Command, [object]$Feature) {
    $caseId = [string]$Case.route_id
    $safeId = $caseId -replace '[^A-Za-z0-9._-]','_'
    $caseDirectory = Join-Path $workRoot $safeId
    if (Test-Path -LiteralPath $caseDirectory) { Remove-Item -LiteralPath $caseDirectory -Recurse -Force }
    $fixture = $null; $probeToken = $null; $probeFailure = $null; $windows = @(); $before = $null; $after = $null
    $beforeState = $null; $afterState = $null; $stateDelta = @()
    $routeError = $null; $status = 'PASS'; $reason = $null
    $focusBefore = $null; $focusAfter = $null
    $oracleExpected = $null; $oracleActual = $null
    $key = if ($null -eq $Command) { '' } else { [string]$Command.args[0] }
    $isFileOutput = ([string]$Case.expected -ceq 'file-output')
    $outputPath = if ($isFileOutput) { Get-CaseOutputTarget $Case $caseDirectory } else { Join-Path $caseDirectory ('output-' + $safeId + '.xlsx') }
    $evidenceOutputs = @()
    $externalProcessesBefore = @()
    $ownedExternalProcessIds = @()
    $externalOutputPath = $null
    $externalWindowObserved = $false
    $externalCleanup = [ordered]@{status='NOT_APPLICABLE';requested_pids=@();remaining_pids=@();failures=@()}
    $watch = $null
    $skipExecution = $false
    try {
        Write-Host ('exhaustive_progress route=' + $caseId + ' stage=start')
        if ([string]$Case.mode -ceq 'external') {
            $externalProcessesBefore = @(Get-HancomProcessIds)
            if ($externalProcessesBefore.Count -gt 0) {
                $status = 'HOLD'
                $reason = 'external-precondition: existing Hancom process is protected'
                $externalCleanup.status = 'PROTECTED_PREEXISTING'
                $skipExecution = $true
            }
        }
        if (-not $skipExecution) {
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $fixture = Initialize-Fixture $excel $books $caseDirectory ([string]$Case.fixture)
            $excel.Visible = $true
            [void]$fixture.book.Activate()
            $excel.WindowState = -4143
            Write-Host ('exhaustive_progress route=' + $caseId + ' stage=fixture-ready')
            Prepare-Route $excel $addin $fixture $Case $Command
            $beforeState = Get-WorkbookStateSnapshot $excel $fixture.book $caseDirectory -ExcludeCaseFiles:$isFileOutput
            $before = [string]$beforeState.digest
            Write-Host ('exhaustive_progress route=' + $caseId + ' stage=before-ready')
            if ($caseId -ceq 'NX-DATA-FOCUS-CELL') {
                $focusStateMacro = Get-AddinMacro $addin 'NxFocusIsEnabled'
                $focusBefore = [bool]$excel.Run($focusStateMacro)
            }
            $probeMode = Get-ProbeMode $Case $key
            $input = Get-RouteInput $key
            $probeToken = [NxExhaustiveUiProbe]::Start([int]$binding.pid, [int64]$excel.Hwnd, $probeMode, $input, $outputPath)
            $macro = if ([string]$Case.kind -ceq 'feature') { Get-AddinMacro $addin 'NxRouteFeature' } else { Get-AddinMacro $addin 'NxRouteCommand' }
            try { [void]$excel.Run($macro, $caseId) } catch { $routeError = $_.Exception.Message }
            Write-Host ('exhaustive_progress route=' + $caseId + ' stage=route-returned')
            if ([string]$Case.mode -ceq 'external') {
                $lastPathMacro = Get-AddinMacro $addin 'NxHangulLastOutputPath'
                $externalOutputPath = [string]$excel.Run($lastPathMacro)
            }
            $deadline = [DateTime]::UtcNow.AddMilliseconds([math]::Min(5000,[int]$Case.timeout_ms))
            if ([string]$Case.expected -ceq 'window-open') {
                $expectedWindowTitles = [string[]]@($Case.oracle.expected_titles)
                while ([NxExhaustiveUiProbe]::ObservedTitleCount($probeToken, $expectedWindowTitles) -eq 0 -and [DateTime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
            } elseif ([string]$Case.expected -ceq 'external-open') {
                $ownedExternalProcessIds = @(Wait-NewHancomDocumentWindow $externalProcessesBefore $externalOutputPath 30000)
                $externalWindowObserved = $ownedExternalProcessIds.Count -gt 0
            } else { Start-Sleep -Milliseconds 250 }
            $probeFailure = [NxExhaustiveUiProbe]::FailureMessage($probeToken)
            $windows = @([NxExhaustiveUiProbe]::Stop($probeToken, 2000)); $probeToken = $null
            Write-Host ('exhaustive_progress route=' + $caseId + ' stage=probe-stopped')
            $afterState = Get-WorkbookStateSnapshot $excel $fixture.book $caseDirectory -ExcludeCaseFiles:$isFileOutput
            $after = [string]$afterState.digest
            $stateDelta = @(Get-WorkbookStateDelta $beforeState.document $afterState.document)
            Write-Host ('exhaustive_progress route=' + $caseId + ' stage=after-ready')
            if ($caseId -ceq 'NX-DATA-FOCUS-CELL') { $focusAfter = [bool]$excel.Run($focusStateMacro) }

            if (-not [string]::IsNullOrWhiteSpace($probeFailure)) { $status='FAIL'; $reason='verifier-ui: ' + $probeFailure }
            elseif (-not [string]::IsNullOrWhiteSpace($routeError)) { $status='FAIL'; $reason='runtime: ' + $routeError }
            elseif ([string]$Case.expected -ceq 'window-open') {
                $oracleResult = Test-OwnedWindowOracle $windows $Case.oracle
                if (-not [bool]$oracleResult.passed) { $status='FAIL'; $reason=[string]$oracleResult.reason }
            }
            elseif ([string]$Case.expected -ceq 'oracle') {
                $oracleResult = Test-RouteOracle $excel $fixture $Case.oracle
                $oracleExpected = $oracleResult.expected; $oracleActual = $oracleResult.actual
                if (-not [bool]$oracleResult.passed) { $status='FAIL'; $reason=[string]$oracleResult.reason }
            }
            elseif ($caseId -ceq 'NX-DATA-FOCUS-CELL' -and $focusBefore -eq $focusAfter) { $status='FAIL'; $reason='runtime: focus enabled state did not toggle' }
            elseif ([string]$Case.expected -ceq 'state-delta' -and $caseId -cne 'NX-DATA-FOCUS-CELL' -and $before -ceq $after) { $status='FAIL'; $reason='runtime: expected state delta was absent' }
            elseif ([string]$Case.expected -ceq 'file-output') {
                $executionResultFailure = Get-ExecutionResultFailure $windows
                if (-not [string]::IsNullOrWhiteSpace([string]$executionResultFailure)) {
                    $status='FAIL'; $reason='runtime: execution result reported failure: ' + [string]$executionResultFailure
                } elseif ($before -cne $after) {
                    $status='FAIL'; $reason='runtime: source workbook, selection, Saved state, or display state changed'
                } else {
                    $outputs = @(Get-CaseOutputFiles $Case $caseDirectory)
                    $fileOutputOracle = Test-FileOutputOracle $excel $Case $outputs $caseDirectory
                    if (-not [bool]$fileOutputOracle.passed) { $status='FAIL'; $reason='runtime: ' + [string]$fileOutputOracle.reason }
                    else { $evidenceOutputs = @(Copy-FileOutputEvidence $safeId $outputs) }
                }
            } elseif ([string]$Case.expected -ceq 'external-open') {
                $ownedExternalProcessIds = @(Get-NewHancomProcessIds $externalProcessesBefore)
                if ([string]::IsNullOrWhiteSpace($externalOutputPath) -or -not (Test-Path -LiteralPath $externalOutputPath) -or -not $externalWindowObserved) {
                    $status='HOLD'; $reason='external-precondition: HWPX output and owned Hancom open were not both observed'
                }
            }
        }
    } catch {
        Write-Host ('exhaustive_progress route=' + $caseId + ' stage=exception')
        $status = 'FAIL'
        $reason = 'fixture: ' + $_.Exception.Message
    } finally {
        if ($null -ne $probeToken) { try { $windows=@([NxExhaustiveUiProbe]::Stop($probeToken,1000)) } catch { } }
        if ([string]$Case.mode -ceq 'external') {
            $ownedExternalProcessIds = @(Get-NewHancomProcessIds $externalProcessesBefore)
            if ($ownedExternalProcessIds.Count -gt 0) {
                $externalCleanup = Close-OwnedHancomProcesses $ownedExternalProcessIds $externalProcessesBefore 30000
                if ([string]$externalCleanup.status -cne 'PASS') {
                    $status = 'FAIL'
                    $cleanupReason = 'external-cleanup: ' + (@($externalCleanup.failures) -join '; ')
                    if (@($externalCleanup.remaining_pids).Count -gt 0) {
                        $cleanupReason += '; remaining=' + (@($externalCleanup.remaining_pids) -join ',')
                    }
                    $reason = if ([string]::IsNullOrWhiteSpace($reason)) { $cleanupReason } else { $reason + ' | ' + $cleanupReason }
                }
            } elseif ($externalProcessesBefore.Count -eq 0) {
                $externalCleanup.status = 'NOT_OBSERVED'
            }
        }
        if ($caseId -ceq 'NX-DATA-FOCUS-CELL' -and $null -ne $addin) {
            try { $disable = Get-AddinMacro $addin 'NxFocusDisable'; [void]$excel.Run($disable) } catch { }
        }
        if ($caseId -ceq 'NX-UTIL-NAVIGATOR' -and $null -ne $addin) { try { $hide=Get-AddinMacro $addin 'NxHostHideNavigator'; [void]$excel.Run($hide) } catch { } }
        if ($null -ne $addin) { try { $releaseClipboard=Get-AddinMacro $addin 'NxClipboardReleaseExcelCopyMode'; [void]$excel.Run($releaseClipboard) } catch { } }
        if ($null -ne $fixture) {
            Release-ComObject $fixture.last; Release-ComObject $fixture.second; Release-ComObject $fixture.sheet
            try { $fixture.book.Close($false) } catch { }
            Release-ComObject $fixture.book
        }
        try { Close-UserWorkbooks $books $addin } catch { }
        try { $excel.DisplayFullScreen=$false; $excel.DisplayAlerts=$false; $excel.MoveAfterReturnDirection=-4121 } catch { }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
    if ($null -ne $watch) { $watch.Stop() }
    $durationMs = if ($null -eq $watch) { 0 } else { [int]$watch.ElapsedMilliseconds }
    Write-Host ('exhaustive_progress route=' + $caseId + ' stage=complete status=' + $status)
    return [ordered]@{
        route_id=$caseId;kind=[string]$Case.kind;mode=[string]$Case.mode;fixture=[string]$Case.fixture;expected=[string]$Case.expected
        status=$status;reason=$reason;before=$before;after=$after;state_delta=$stateDelta;focus_before=$focusBefore;focus_after=$focusAfter
        oracle_expected=$oracleExpected;oracle_actual=$oracleActual;windows=$windows;duration_ms=$durationMs
        evidence_outputs=$evidenceOutputs
        owned_external_process_ids=@($ownedExternalProcessIds);external_output_path=$externalOutputPath;external_window_observed=$externalWindowObserved;external_cleanup=$externalCleanup
    }
}

try {
    [void](New-Item -ItemType Directory -Path $EvidenceRoot -Force)
    $isolatedProfileRoot = Join-Path $EvidenceRoot 'profile/LHexcel'
    Assert-Condition (-not (Test-Path -LiteralPath $isolatedProfileRoot)) 'Isolated LHexcel profile already exists'
    [void](New-Item -ItemType Directory -Path $isolatedProfileRoot -Force)
    $env:LHEXCEL_PROFILE_ROOT = $isolatedProfileRoot
    $verifierSha = Get-Sha256 $PSCommandPath
    Assert-Condition (Test-Path -LiteralPath $SourceRoot -PathType Container) 'SourceRoot is missing'
    Assert-Condition (Test-Path -LiteralPath $ArtifactPath -PathType Leaf) 'ArtifactPath is missing'
    Assert-Condition ([IO.Path]::GetExtension($ArtifactPath) -ieq '.xlam') 'ArtifactPath must be XLAM'
    Assert-Condition (Test-Path -LiteralPath $probeSource -PathType Leaf) 'UI probe source is missing'
    Assert-Condition ($RunId -match '^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$') 'RunId is invalid'
    Assert-Condition ($SourceDigest -match '^[0-9a-f]{64}$') 'source_tree_sha256 is invalid'
    Assert-Condition ($SnapshotDigest -match '^[0-9a-f]{64}$') 'source_snapshot_sha256 is invalid'
    Assert-Condition (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) 'LOCALAPPDATA is unavailable'
    $fixtureBase = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'LHexcel\ProductExhaustiveFixtures'))
    $workRoot = [IO.Path]::GetFullPath((Join-Path $fixtureBase $RunId))
    Assert-Condition ($workRoot.StartsWith($fixtureBase + '\',[StringComparison]::OrdinalIgnoreCase)) 'Local fixture root escaped its boundary'
    Assert-Condition (-not (Test-Path -LiteralPath $workRoot)) 'Local fixture root already exists'
    if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -ne 0) { throw 'ENVIRONMENT_MISMATCH: close all user Excel windows before ProductExhaustive verification' }

    $nativeContract = Get-Content -LiteralPath $nativeContractPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $commandContract = Get-Content -LiteralPath (Join-Path $SourceRoot 'contracts/command-contract.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $featureContract = Get-Content -LiteralPath (Join-Path $SourceRoot 'contracts/feature-contract.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-Condition ([int]$nativeContract.route_count -eq $expectedRouteCount) ('expected_routes=' + $expectedRouteCount + ' contract mismatch')
    Assert-Condition (@($nativeContract.cases).Count -eq $expectedRouteCount) 'Native case inventory is incomplete'
    foreach ($nativeCase in $nativeContract.cases) {
        Assert-Condition ($null -ne $nativeCase.oracle -and -not [string]::IsNullOrWhiteSpace([string]$nativeCase.oracle.type)) ('Native case oracle is missing: ' + [string]$nativeCase.route_id)
        $oracleType = [string]$nativeCase.oracle.type
        Assert-Condition ($oracleType -in @('property','owned-window','state-delta','file-output','external-open')) ('unsupported oracle type: ' + $oracleType)
        if ($oracleType -ceq 'file-output') {
            Assert-Condition ([string]$nativeCase.oracle.format -in @('xlsx','png','pdf','settings')) ('unsupported file-output format: ' + [string]$nativeCase.route_id)
            Assert-Condition ([int]$nativeCase.oracle.expected_files -ge 1) ('invalid file-output count: ' + [string]$nativeCase.route_id)
        }
    }
    if ($routeSelectionRequested) {
        Assert-Condition ($null -ne $RouteIds -and @($RouteIds).Count -gt 0) 'Targeted route selection is empty'
        $requestedSet = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($requestedRouteIdValue in @($RouteIds)) {
            $requestedRouteId = [string]$requestedRouteIdValue
            Assert-Condition (-not [string]::IsNullOrWhiteSpace($requestedRouteId)) 'Targeted route selection is empty'
            Assert-Condition ($requestedSet.Add($requestedRouteId)) ('Duplicate targeted route id: ' + $requestedRouteId)
            $requestedRouteIds += $requestedRouteId
        }
        $knownRouteIds = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        foreach ($nativeCase in $nativeContract.cases) { [void]$knownRouteIds.Add([string]$nativeCase.route_id) }
        foreach ($requestedRouteId in $requestedRouteIds) {
            Assert-Condition ($knownRouteIds.Contains([string]$requestedRouteId)) ('Unknown targeted route id: ' + [string]$requestedRouteId)
        }
        $selectedCases = @()
        foreach ($requestedRouteId in $requestedRouteIds) {
            $selectedCases += @($nativeContract.cases | Where-Object { [string]$_.route_id -ceq [string]$requestedRouteId })
        }
        Assert-Condition ($selectedCases.Count -eq $requestedRouteIds.Count) 'Targeted route selection did not resolve exactly'
        $executionRouteCount = $selectedCases.Count
    } else {
        $selectedCases = @($nativeContract.cases)
        $executionRouteCount = $expectedRouteCount
    }
    $navigatorHostRequired = @($selectedCases | Where-Object { [string]$_.route_id -ceq 'NX-UTIL-NAVIGATOR' }).Count -gt 0
    $navigatorHostReceipt.required = $navigatorHostRequired
    if ($navigatorHostRequired) {
        $resolvedNxHostRoot = if ([string]::IsNullOrWhiteSpace($NxHostRoot)) { [IO.Path]::GetDirectoryName($ArtifactPath) } else { [IO.Path]::GetFullPath($NxHostRoot) }
        Assert-Condition (Test-Path -LiteralPath $resolvedNxHostRoot -PathType Container) 'NxHostRoot is missing'
        $sourceX86 = Resolve-NxHostDll $resolvedNxHostRoot 'x86' 'NxHost32.dll'
        $sourceX64 = Resolve-NxHostDll $resolvedNxHostRoot 'x64' 'NxHost64.dll'
        $runtimeRoot = Join-Path $EvidenceRoot 'nxhost-runtime'
        $x86Root = Join-Path $runtimeRoot 'x86'; $x64Root = Join-Path $runtimeRoot 'x64'
        [void](New-Item -ItemType Directory -Path $x86Root,$x64Root -Force)
        $navigatorHostX86 = Join-Path $x86Root 'NxHost32.dll'
        $navigatorHostX64 = Join-Path $x64Root 'NxHost64.dll'
        Copy-Item -LiteralPath $sourceX86 -Destination $navigatorHostX86
        Copy-Item -LiteralPath $sourceX64 -Destination $navigatorHostX64
        Assert-Condition ((Get-Sha256 $sourceX86) -ceq (Get-Sha256 $navigatorHostX86)) 'NxHost32 evidence copy mismatch'
        Assert-Condition ((Get-Sha256 $sourceX64) -ceq (Get-Sha256 $navigatorHostX64)) 'NxHost64 evidence copy mismatch'
        $navigatorHostReceipt.nxhost32_sha256 = Get-Sha256 $navigatorHostX86
        $navigatorHostReceipt.nxhost64_sha256 = Get-Sha256 $navigatorHostX64
        $navigatorHostRegisterScript = Join-Path $SourceRoot 'build/Register-NxHost.ps1'
        Assert-Condition (Test-Path -LiteralPath $navigatorHostRegisterScript -PathType Leaf) 'Register-NxHost.ps1 is missing'
        $navigatorHostStatePath = Join-Path $env:LOCALAPPDATA 'LHexcel\NxHost\state\r57\registry-state.json'
        $navigatorHostRegistrationAttempted = $true
        $registerOutput = @(& $navigatorHostRegisterScript -Action Register -OptIn -X86Dll $navigatorHostX86 -X64Dll $navigatorHostX64)
        Assert-Condition (@($registerOutput | Where-Object { [string]$_ -ceq 'PASS|Register' }).Count -gt 0) 'NxHost registration did not pass'
        Assert-Condition (Test-Path -LiteralPath $navigatorHostStatePath -PathType Leaf) 'NxHost registry state receipt is missing'
        $registryState = Get-Content -LiteralPath $navigatorHostStatePath -Raw -Encoding UTF8 | ConvertFrom-Json
        $navigatorHostReceipt.registry_snapshot_before = [string]$registryState.snapshot_sha256
        Assert-Condition ([string]$navigatorHostReceipt.registry_snapshot_before -match '^[0-9a-f]{64}$') 'NxHost registry baseline is invalid'
        $navigatorHostRegistered = $true
        $navigatorHostReceipt.status = 'REGISTERED'
    }
    $commandMap = Get-CommandMap $commandContract; $featureMap = Get-FeatureMap $featureContract
    $artifactSha = Get-Sha256 $ArtifactPath
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    Add-Type -AssemblyName Accessibility
    $automationReferences = @([System.Windows.Automation.AutomationElement].Assembly.Location,[System.Windows.Automation.AutomationIdentifier].Assembly.Location,[Accessibility.IAccessible].Assembly.Location) | Select-Object -Unique
    Add-Type -TypeDefinition (Get-Content -LiteralPath $probeSource -Raw -Encoding UTF8) -Language CSharp -ReferencedAssemblies $automationReferences
    $clipboardSnapshot = Read-ClipboardUnicode; $clipboardSnapshotCaptured = $true

    $failurePhase = 'native-execution'
    [void](New-Item -ItemType Directory -Path $workRoot -Force)
    foreach ($case in $selectedCases) {
        $caseId = [string]$case.route_id
        $caseCleanupFailures = @()
        $caseCleanupExitMode = 'NOT_STARTED'
        $caseStartedUtc = [DateTime]::UtcNow
        $caseProfileName = ('{0:D3}-{1}' -f ($caseProcessReceipts.Count + 1), ($caseId -replace '[^A-Za-z0-9._-]','_'))
        $caseProfileRoot = Join-Path $isolatedProfileRoot $caseProfileName
        $caseNavigatorRequired = $caseId -ceq 'NX-UTIL-NAVIGATOR'
        $caseExcelPid = 0
        $interactiveLaunchProcess = $null
        try {
            Assert-Condition (-not (Test-Path -LiteralPath $caseProfileRoot)) ('Isolated case profile already exists: ' + $caseId)
            [void](New-Item -ItemType Directory -Path $caseProfileRoot -Force)
            $env:LHEXCEL_PROFILE_ROOT = $caseProfileRoot
            Assert-Condition (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -eq 0) ('ENVIRONMENT_MISMATCH: EXCEL remained before isolated case ' + $caseId)
            $baseline = Get-ExcelProcessBaseline
            if ($caseNavigatorRequired) {
                $excelExecutable = Resolve-ExcelExecutable
                $interactiveLaunchProcess = Start-Process -FilePath $excelExecutable -ArgumentList '/x' -PassThru
                $caseExcelPid = [int]$interactiveLaunchProcess.Id
                $rotDeadline = [DateTime]::UtcNow.AddSeconds(30)
                do {
                    try { $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { $excel = $null }
                    if ($null -eq $excel) { Start-Sleep -Milliseconds 100 }
                } while ($null -eq $excel -and [DateTime]::UtcNow -lt $rotDeadline)
                Assert-Condition ($null -ne $excel) 'Interactive Excel ROT object was unavailable'
            } else {
                $excel = New-Object -ComObject Excel.Application
            }
            $binding = Get-ExactExcelProcessOwnership $excel $baseline ('ProductExhaustive case ' + $caseId)
            Assert-Condition ([bool]$binding.owned) ('ProductExhaustive Excel ownership failed for ' + $caseId + ': ' + [string]$binding.failure)
            $process = $binding.process
            $caseExcelPid = [int]$binding.pid
            if ($caseNavigatorRequired) {
                Assert-Condition ([int]$interactiveLaunchProcess.Id -eq [int]$binding.pid) 'Interactive Excel process ownership mismatch'
                $interactiveLaunchProcess.Dispose()
                $interactiveLaunchProcess = $null
            }
            $excel.Visible = $true; $excel.DisplayAlerts = $false; $excel.WindowState = -4143
            if ($caseNavigatorRequired) {
                $comAddins = $excel.COMAddIns
                $addinDeadline = [DateTime]::UtcNow.AddSeconds(15)
                do {
                    try {
                        [void]$comAddins.Update()
                        $hostAddin = $comAddins.Item('LH.NxHost.Connect')
                    } catch {
                        $hostAddin = $null
                        Start-Sleep -Milliseconds 100
                    }
                } while ($null -eq $hostAddin -and [DateTime]::UtcNow -lt $addinDeadline)
                Assert-Condition ($null -ne $hostAddin) 'NxHost COM add-in was not discovered by interactive Excel'
                $hostAddin.Connect=$true
                $navigatorHostConnected = [bool]$hostAddin.Connect
                Assert-Condition $navigatorHostConnected 'NxHost COM add-in did not connect'
                $navigatorHostReceipt.connected = $navigatorHostConnected
                $navigatorHostReceipt.status = 'CONNECTED'
            }
            $books = $excel.Workbooks
            $addin = $books.Open($ArtifactPath, $false, $true)
            Assert-Condition ([bool]$addin.IsAddin -and [string]$addin.Name -ceq [IO.Path]::GetFileName($ArtifactPath)) ('ProductExhaustive add-in identity failed for ' + $caseId)
            $command = if ([string]$case.kind -ceq 'command') { $commandMap[$caseId] } else { $null }
            $feature = if ([string]$case.kind -ceq 'feature') { $featureMap[$caseId] } else { $null }
            [void]$results.Add((Invoke-NativeCase $case $command $feature))
        } finally {
            if ($caseNavigatorRequired -and $null -ne $addin -and $null -ne $excel) {
                try { $hideNavigator = Get-AddinMacro $addin 'NxHostHideNavigator'; [void]$excel.Run($hideNavigator) } catch { }
                try { $resetBridge = Get-AddinMacro $addin 'NxHostBridgeReset'; [void]$excel.Run($resetBridge) } catch { }
            }
            if ($null -ne $hostAddin) {
                try { if ([bool]$hostAddin.Connect) { $hostAddin.Connect=$false } } catch { $caseCleanupFailures += ('NxHost disconnect failed for ' + $caseId + ': ' + $_.Exception.Message) }
            }
            Release-ComObject $hostAddin; Release-ComObject $comAddins
            if ($null -ne $addin) {
                try { $addin.Close($false) } catch { $caseCleanupFailures += ('Add-in close failed for ' + $caseId + ': ' + $_.Exception.Message) }
                Release-ComObject $addin
            }
            Release-ComObject $books
            if ($null -ne $excel) {
                try { $excel.DisplayFullScreen=$false; $excel.Quit() } catch { $caseCleanupFailures += ('Excel quit failed for ' + $caseId + ': ' + $_.Exception.Message) }
                Release-ComObject $excel
            }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            if ($null -ne $process) {
                try {
                    $caseCleanup = Stop-ExactProcessAfterGrace -Process $process -Label ('ProductExhaustive case ' + $caseId) -GraceMs 10000 -Detailed
                    $caseCleanupExitMode = [string]$caseCleanup.exit_mode
                    if ($caseCleanup.failure) { $caseCleanupFailures += [string]$caseCleanup.failure }
                } catch { $caseCleanupFailures += ('Excel process cleanup failed for ' + $caseId + ': ' + $_.Exception.Message) }
                try { $process.Dispose() } catch { }
            }
            if ($null -ne $interactiveLaunchProcess) {
                try {
                    $launchCleanup = Stop-ExactProcessAfterGrace -Process $interactiveLaunchProcess -Label ('ProductExhaustive interactive launch ' + $caseId) -GraceMs 10000 -Detailed
                    if ($caseCleanupExitMode -ceq 'NOT_STARTED') { $caseCleanupExitMode = [string]$launchCleanup.exit_mode }
                    if ($launchCleanup.failure) { $caseCleanupFailures += [string]$launchCleanup.failure }
                } catch { $caseCleanupFailures += ('Interactive Excel process cleanup failed for ' + $caseId + ': ' + $_.Exception.Message) }
                try { $interactiveLaunchProcess.Dispose() } catch { }
            }
            [void]$caseProcessReceipts.Add([ordered]@{
                route_id=$caseId;excel_pid=$caseExcelPid;started_utc=$caseStartedUtc.ToString('o');completed_utc=[DateTime]::UtcNow.ToString('o')
                profile_root=$caseProfileRoot;exit_mode=$caseCleanupExitMode;cleanup_failures=$caseCleanupFailures
            })
            if ($caseCleanupExitMode -ceq 'FORCED_FAILED') { $cleanupExitMode = 'FORCED_FAILED' }
            elseif ($caseCleanupExitMode -ceq 'FORCED_CONTAINED' -and $cleanupExitMode -cne 'FORCED_FAILED') { $cleanupExitMode = 'FORCED_CONTAINED' }
            elseif ($cleanupExitMode -ceq 'NOT_STARTED' -and $caseCleanupExitMode -ceq 'NATURAL') { $cleanupExitMode = 'NATURAL' }
            $excel = $null; $books = $null; $addin = $null; $binding = $null; $process = $null
            $interactiveLaunchProcess = $null
            $comAddins = $null; $hostAddin = $null
        }
        if ($caseCleanupFailures.Count -gt 0) {
            foreach ($caseCleanupFailure in $caseCleanupFailures) { $cleanupFailures += [string]$caseCleanupFailure }
            throw ('Isolated Excel cleanup failed for ' + $caseId)
        }
        Wait-NoExcelProcesses ('after isolated case ' + $caseId) 10000
        Assert-Condition ((Get-Sha256 $ArtifactPath) -ceq $artifactSha) ('Artifact changed during isolated case ' + $caseId)
    }
    Assert-Condition ((Get-Sha256 $ArtifactPath) -ceq $artifactSha) 'Artifact changed during ProductExhaustive verification'
} catch {
    $terminalCode = if ($failurePhase -ceq 'preflight') { 10 } else { 24 }
    $failureMessage = $_.Exception.Message
} finally {
    if ($navigatorHostRequired -and $null -ne $addin -and $null -ne $excel) {
        try { $hideNavigator = Get-AddinMacro $addin 'NxHostHideNavigator'; [void]$excel.Run($hideNavigator) } catch { }
        try { $resetBridge = Get-AddinMacro $addin 'NxHostBridgeReset'; [void]$excel.Run($resetBridge) } catch { }
    }
    if ($null -ne $hostAddin) {
        try { if ([bool]$hostAddin.Connect) { $hostAddin.Connect=$false } } catch { $cleanupFailures += ('NxHost disconnect failed: ' + $_.Exception.Message) }
    }
    Release-ComObject $hostAddin; Release-ComObject $comAddins
    if ($null -ne $addin) { try { $addin.Close($false) } catch { $cleanupFailures += $_.Exception.Message }; Release-ComObject $addin }
    Release-ComObject $books
    if ($null -ne $excel) { try { $excel.DisplayFullScreen=$false; $excel.Quit() } catch { $cleanupFailures += $_.Exception.Message }; Release-ComObject $excel }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($null -ne $process) {
        try {
            $cleanup = Stop-ExactProcessAfterGrace -Process $process -Label 'ProductExhaustive Excel' -GraceMs 10000 -Detailed
            $cleanupExitMode = [string]$cleanup.exit_mode
            if ($cleanup.failure) { $cleanupFailures += [string]$cleanup.failure }
        } catch { $cleanupFailures += $_.Exception.Message }
        try { $process.Dispose() } catch { }
    }
    if ($navigatorHostRegistrationAttempted -and -not [string]::IsNullOrWhiteSpace($navigatorHostStatePath) -and (Test-Path -LiteralPath $navigatorHostStatePath -PathType Leaf)) {
        $unregisterRequired = $navigatorHostRegistered
        try {
            $registryState = Get-Content -LiteralPath $navigatorHostStatePath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ([string]$registryState.status -ceq 'ACTIVE') { $unregisterRequired = $true }
            elseif ([string]$registryState.status -ceq 'RESTORED') { $navigatorHostReceipt.registry_snapshot_after = [string]$registryState.restored_snapshot_sha256 }
        } catch { $cleanupFailures += ('NxHost registry state inspection failed: ' + $_.Exception.Message) }
        if ($unregisterRequired) {
            try {
                $unregisterOutput = @(& $navigatorHostRegisterScript -Action Unregister -X86Dll $navigatorHostX86 -X64Dll $navigatorHostX64)
                Assert-Condition (@($unregisterOutput | Where-Object { [string]$_ -ceq 'PASS|Unregister' }).Count -gt 0) 'NxHost unregistration did not pass'
                $registryState = Get-Content -LiteralPath $navigatorHostStatePath -Raw -Encoding UTF8 | ConvertFrom-Json
                $navigatorHostReceipt.registry_snapshot_after = [string]$registryState.restored_snapshot_sha256
            } catch { $cleanupFailures += ('NxHost registry restore failed: ' + $_.Exception.Message) }
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$navigatorHostReceipt.registry_snapshot_before) -and
            [string]$navigatorHostReceipt.registry_snapshot_after -ceq [string]$navigatorHostReceipt.registry_snapshot_before) {
            $navigatorHostReceipt.status = 'PASS'
        } else {
            $navigatorHostReceipt.status = 'FAIL'
            $cleanupFailures += 'NxHost registry snapshot was not restored exactly'
        }
    } elseif ($navigatorHostRequired) {
        $navigatorHostReceipt.status = 'FAIL'
        $cleanupFailures += 'NxHost registry recovery state was unavailable'
    }
    if (-not [string]::IsNullOrWhiteSpace($workRoot) -and (Test-Path -LiteralPath $workRoot)) {
        try {
            Remove-FixtureRootWithRetry $workRoot $fixtureBase 10000
            $fixtureRootCleanup = 'PASS'
        } catch {
            $fixtureRootCleanup = 'FAIL'
            $cleanupFailures += ('Local fixture cleanup failed: ' + $_.Exception.Message)
        }
    } elseif (-not [string]::IsNullOrWhiteSpace($workRoot)) {
        $fixtureRootCleanup = 'PASS'
    }
    try {
        if ($profileOverrideWasPresent) { $env:LHEXCEL_PROFILE_ROOT = $originalProfileOverride }
        else { Remove-Item Env:LHEXCEL_PROFILE_ROOT -ErrorAction SilentlyContinue }
    } catch { $cleanupFailures += ('LHexcel profile override restore failed: ' + $_.Exception.Message) }
    if ($clipboardSnapshotCaptured) { try { Write-ClipboardUnicode ([string]$clipboardSnapshot) } catch { $cleanupFailures += ('Clipboard restore failed: ' + $_.Exception.Message) } }
    $externalCleanupFailures = @($results | Where-Object { $null -ne $_.external_cleanup -and [string]$_.external_cleanup.status -ceq 'FAIL' })
    $externalRemainingPids = @(foreach ($caseResult in $results) {
        if ($null -ne $caseResult.external_cleanup) { @($caseResult.external_cleanup.remaining_pids) }
    })
    $externalRemainingPids = @($externalRemainingPids | Sort-Object -Unique)
    if ($externalCleanupFailures.Count -gt 0) {
        $cleanupFailures += ('Owned Hancom cleanup failed for ' + $externalCleanupFailures.Count + ' route(s)')
    }
    if ($cleanupFailures.Count -gt 0) { $terminalCode=25; if ([string]::IsNullOrWhiteSpace($failureMessage)) { $failureMessage='ProductExhaustive cleanup failed' } }

    $passed = @($results | Where-Object { $_.status -ceq 'PASS' }).Count
    $held = @($results | Where-Object { $_.status -ceq 'HOLD' }).Count
    $failed = @($results | Where-Object { $_.status -ceq 'FAIL' }).Count
    if ($terminalCode -eq 0 -and $results.Count -ne $executionRouteCount) { $terminalCode=24; $failureMessage='Executed route count is incomplete' }
    if ($terminalCode -eq 0 -and $failed -gt 0) { $terminalCode=24; $failureMessage='One or more native routes failed' }
    $status='PASS'
    if ($terminalCode -ne 0 -or $failed -gt 0) { $status='FAIL' }
    elseif ($held -gt 0) { $status='HOLD'; $terminalCode=26; $failureMessage='One or more native routes are held' }
    if ($runScope -ceq 'FULL') { $formalReleaseGate = $status }
    Write-AtomicJson (Join-Path $EvidenceRoot $resultFileName) ([ordered]@{
        schema_version=1;suite='ProductExhaustive';scope=$runScope;status=$status;formal_release_gate=$formalReleaseGate;run_id=$RunId
        artifact_sha256=$artifactSha;source_tree_sha256=$SourceDigest;source_snapshot_sha256=$SnapshotDigest;verifier_sha256=$verifierSha
        contract_routes=$expectedRouteCount;requested_routes=$requestedRouteIds;expected_routes=$executionRouteCount;executed_routes=$results.Count;passed_routes=$passed;held_routes=$held;failed_routes=$failed
        cases=$results.ToArray();failure_phase=if($terminalCode -eq 0){$null}else{$failurePhase};failure=$failureMessage
        case_processes=$caseProcessReceipts.ToArray()
        navigator_host=$navigatorHostReceipt
        cleanup=[ordered]@{status=if($cleanupFailures.Count -eq 0){'PASS'}else{'FAIL'};exit_mode=$cleanupExitMode;fixture_root_cleanup=$fixtureRootCleanup;external_remaining_pids=$externalRemainingPids;failures=$cleanupFailures}
    })
    exit $terminalCode
}
