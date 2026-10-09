param(
    [Parameter(Mandatory = $true)][string]$OldDrawing,
    [Parameter(Mandatory = $true)][string]$NewDrawing,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [string]$PythonExe,
    [string]$CoreConsole
)

$ErrorActionPreference = 'Stop'
$scriptRoot = $PSScriptRoot
. (Join-Path $scriptRoot 'AutoCADProfileGuard.ps1')

function Find-CoreConsole {
    if ($CoreConsole -and (Test-Path -LiteralPath $CoreConsole)) { return (Resolve-Path -LiteralPath $CoreConsole).Path }
    $candidates = 2027..2020 | ForEach-Object { "C:\Program Files\Autodesk\AutoCAD $_\accoreconsole.exe" }
    $found = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $found) { throw 'AutoCAD Core Console (accoreconsole.exe) was not found. This workflow does not fall back to DXF.' }
    return $found
}

function Find-Python {
    if ($PythonExe -and (Test-Path -LiteralPath $PythonExe)) { return (Resolve-Path -LiteralPath $PythonExe).Path }
    $command = Get-Command python -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    throw 'Python was not found on PATH. Pass -PythonExe with an explicit path (must have openpyxl installed).'
}

function AcadPath([string]$Path) { return ((Resolve-Path -LiteralPath $Path).Path -replace '\\','/') }

function Invoke-Core([string]$Drawing, [string]$Script, [string]$Log) {
    Invoke-AutoCADCoreGuarded -CoreConsole $core -Drawing $Drawing -Script $Script -Log $Log
}

foreach ($path in @($OldDrawing, $NewDrawing)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Required file not found: $path" }
}
if (Test-Path -LiteralPath $OutputDirectory) {
    $existing = Get-ChildItem -LiteralPath $OutputDirectory -Force -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($existing) { throw "Output directory is not empty: $OutputDirectory" }
}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$core = Find-CoreConsole
$python = Find-Python
$work = Join-Path $OutputDirectory 'work'
New-Item -ItemType Directory -Force -Path $work | Out-Null
$oldCopy = Join-Path $work 'old.dwg'
$newCopy = Join-Path $work 'new.dwg'
Copy-Item -LiteralPath $OldDrawing -Destination $oldCopy
Copy-Item -LiteralPath $NewDrawing -Destination $newCopy

$exportLisp = AcadPath (Join-Path $scriptRoot 'export_dwg_facts.lsp')
$oldFacts = Join-Path $work 'old_facts.ndjson'
$newFacts = Join-Path $work 'new_facts.ndjson'
foreach ($job in @(
    @{ Drawing = $oldCopy; Facts = $oldFacts; Name = 'old' },
    @{ Drawing = $newCopy; Facts = $newFacts; Name = 'new' }
)) {
    $factsAcad = ($job.Facts -replace '\\','/')
    $scr = Join-Path $work ("export_{0}.scr" -f $job.Name)
    $lines = @(
        '(setq cdr:old-filedia (getvar "FILEDIA") cdr:old-secureload (getvar "SECURELOAD"))',
        '(setvar "FILEDIA" 0)', '(setvar "SECURELOAD" 0)',
        "(load `"$exportLisp`")", "(cdr:export-facts `"$factsAcad`")",
        '(setvar "FILEDIA" cdr:old-filedia)', '(setvar "SECURELOAD" cdr:old-secureload)',
        '_.QUIT'
    )
    [IO.File]::WriteAllLines($scr, $lines, [Text.UTF8Encoding]::new($false))
    Invoke-Core $job.Drawing $scr (Join-Path $OutputDirectory ("export_{0}.log" -f $job.Name))
    if (-not (Test-Path -LiteralPath $job.Facts)) { throw "DWG fact extraction did not create $($job.Facts)" }
}

$candidatesJson = Join-Path $OutputDirectory 'change_candidates.json'
$changesJson = Join-Path $OutputDirectory 'changes.json'
& $python (Join-Path $scriptRoot 'compare_facts.py') $oldFacts $newFacts --output $candidatesJson --max-items 1000
if ($LASTEXITCODE -ne 0) { throw "DWG fact comparison failed with exit code $LASTEXITCODE" }
& $python (Join-Path $scriptRoot 'curate_changes.py') $candidatesJson $newFacts --output $changesJson
if ($LASTEXITCODE -ne 0) { throw "Deliverable change curation failed with exit code $LASTEXITCODE" }
$quality = Get-Content -LiteralPath $changesJson -Raw -Encoding utf8 | ConvertFrom-Json
if (-not $quality.meta.delivery_ready) {
    $warnings = ($quality.meta.quality_warnings -join '; ')
    throw "The near-deliverable quality gate failed. Review changes.json before delivery. $warnings"
}

$rawWork = Join-Path $work 'compare_snapshot_raw.dwg'
$compareScript = Join-Path $work 'compare.scr'
$oldAcad = AcadPath $oldCopy
$rawAcad = ($rawWork -replace '\\','/')
$compareLines = @(
    '(setq cdr:old-filedia (getvar "FILEDIA") cdr:old-cmdecho (getvar "CMDECHO"))',
    '(setvar "FILEDIA" 0)', '(setvar "CMDECHO" 1)',
    '_.-COMPARE', $oldAcad, '_.COMPAREEXPORT', '', $rawAcad,
    '(setvar "FILEDIA" cdr:old-filedia)', '(setvar "CMDECHO" cdr:old-cmdecho)',
    '_.QUIT','_N'
)
[IO.File]::WriteAllLines($compareScript, $compareLines, [Text.UTF8Encoding]::new($false))
Invoke-Core $newCopy $compareScript (Join-Path $OutputDirectory 'compare.log')
if (-not (Test-Path -LiteralPath $rawWork)) { throw 'COMPAREEXPORT did not create the native comparison snapshot.' }
$rawFinal = Join-Path $OutputDirectory 'CAD圖面比較快照_原始.dwg'
Copy-Item -LiteralPath $rawWork -Destination $rawFinal

$deliverableArgs = @{
    ChangesJson = $changesJson; RawSnapshot = $rawFinal; OutputDirectory = $OutputDirectory
    PythonExe = $python; CoreConsole = $core
    OldName = (Split-Path -Leaf $OldDrawing); NewName = (Split-Path -Leaf $NewDrawing)
}
& (Join-Path $scriptRoot 'New-DwgRevisionDeliverables.ps1') @deliverableArgs | Out-Null

$manifest = [ordered]@{
    OldDrawing = (Resolve-Path -LiteralPath $OldDrawing).Path
    NewDrawing = (Resolve-Path -LiteralPath $NewDrawing).Path
    CoreConsole = $core; ChangesJson = $changesJson
    CandidateChangesJson = $candidatesJson
    Workbook = (Join-Path $OutputDirectory 'DWG圖面修改彙整.xlsx')
    RawSnapshot = $rawFinal
    AnnotatedSnapshot = (Join-Path $OutputDirectory 'CAD圖面比較快照_項次標示.dwg')
}
$manifestPath = Join-Path $OutputDirectory 'manifest.json'
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding utf8
$manifest | ConvertTo-Json -Depth 4
