param(
    [Parameter(Mandatory = $true)][string]$ChangesJson,
    [Parameter(Mandatory = $true)][string]$RawSnapshot,
    [Parameter(Mandatory = $true)][string]$OutputDirectory,
    [string]$PythonExe,
    [string]$CoreConsole,
    [string]$OldName = '舊版 DWG',
    [string]$NewName = '新版 DWG'
)

$ErrorActionPreference = 'Stop'
$scriptRoot = $PSScriptRoot
. (Join-Path $scriptRoot 'AutoCADProfileGuard.ps1')

function Find-CoreConsole {
    if ($CoreConsole -and (Test-Path -LiteralPath $CoreConsole)) { return (Resolve-Path -LiteralPath $CoreConsole).Path }
    $candidates = 2027..2020 | ForEach-Object { "C:\Program Files\Autodesk\AutoCAD $_\accoreconsole.exe" }
    $found = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if (-not $found) { throw 'AutoCAD Core Console (accoreconsole.exe) was not found.' }
    return $found
}

function Find-Python {
    if ($PythonExe -and (Test-Path -LiteralPath $PythonExe)) { return (Resolve-Path -LiteralPath $PythonExe).Path }
    $command = Get-Command python -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    throw 'Python was not found on PATH. Pass -PythonExe with an explicit path (must have openpyxl installed).'
}

function AcadPath([string]$Path) { return ((Resolve-Path -LiteralPath $Path).Path -replace '\\','/') }

$core = Find-CoreConsole
$python = Find-Python
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$xlsx = Join-Path $OutputDirectory 'DWG圖面修改彙整.xlsx'
$annotatedWork = Join-Path $OutputDirectory 'compare_snapshot_annotated.dwg'
$annotatedFinal = Join-Path $OutputDirectory 'CAD圖面比較快照_項次標示.dwg'
$annotationLisp = Join-Path $OutputDirectory 'annotation_generated.lsp'
$annotationScript = Join-Path $OutputDirectory 'annotation_generated.scr'
$annotationLog = Join-Path $OutputDirectory 'annotation.log'

& $python (Join-Path $scriptRoot 'build_excel.py') $ChangesJson $xlsx --old-name $OldName --new-name $NewName
if ($LASTEXITCODE -ne 0) { throw "Excel generation failed with exit code $LASTEXITCODE" }

& $python (Join-Path $scriptRoot 'generate_annotation_lsp.py') $ChangesJson $annotationLisp
if ($LASTEXITCODE -ne 0) { throw "Annotation generation failed with exit code $LASTEXITCODE" }

Copy-Item -LiteralPath $RawSnapshot -Destination $annotatedWork -Force
$lispAcad = AcadPath $annotationLisp
$scriptLines = @(
    '(setq cdr:old-filedia (getvar "FILEDIA") cdr:old-secureload (getvar "SECURELOAD"))',
    '(setvar "FILEDIA" 0)', '(setvar "SECURELOAD" 0)',
    "(load `"$lispAcad`")", '_.CDRANNOTATE', '_.QSAVE',
    '(setvar "FILEDIA" cdr:old-filedia)', '(setvar "SECURELOAD" cdr:old-secureload)',
    '_.QUIT'
)
[IO.File]::WriteAllLines($annotationScript, $scriptLines, [Text.UTF8Encoding]::new($false))
Invoke-AutoCADCoreGuarded -CoreConsole $core -Drawing $annotatedWork -Script $annotationScript -Log $annotationLog
if (-not (Test-Path -LiteralPath $annotatedWork)) { throw 'Annotated snapshot was not created.' }
Copy-Item -LiteralPath $annotatedWork -Destination $annotatedFinal -Force

[pscustomobject]@{ Workbook = $xlsx; AnnotatedSnapshot = $annotatedFinal; AnnotationLog = $annotationLog } | ConvertTo-Json -Depth 3
