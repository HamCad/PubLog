#requires -Version 5.1
<#
.SYNOPSIS
  FLIS bulk download -> sustainment.db, end to end. Drop the raw FLIS .txt
  files in one folder and run this.

.DESCRIPTION
  1. Schema    applies any pending migrations\NNN_*.sql (refuses a legacy DB)
  2. Raw       FlisIngest.ps1: identifies every file by name, checks headers
               against flis_registry.psd1, lands fb_* tables (batch-stamped)
  3. Typed     FlisTransform.ps1: fb_* -> sustainment tables per flis_model.psd1
               (dates to ISO, numbers, NIIN/CAGE checks, history dedupe,
               rolling stock/forecast windows)
  4. Coverage  cross-file orphan and gap checks
  Console output is grouped and capped (see -ConsoleDetail); every finding is
  in fb_issue_log and the report file under -ReportFolder. Ends by throwing if
  anything was an ERROR, after reporting everything.

  The PUB LOG pipeline is separate: Update-PubLogDatabase.ps1.

.EXAMPLE
  .\Invoke-FlisPipeline.ps1 -DropFolder D:\FLIS\drop
.EXAMPLE
  .\Invoke-FlisPipeline.ps1 -DropFolder D:\FLIS\drop -AsOfMonth 2026-10 -UpdatePriorMonths 2
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DropFolder,
    [string]$DatabasePath = (Join-Path $PSScriptRoot 'sustainment.db'),
    [string]$ReportFolder = (Join-Path $PSScriptRoot 'logs'),
    # Rolling files: which month is "current" (yyyy-MM). Default: the month of
    # the newest data file's timestamp in the drop folder (the pull date).
    [string]$AsOfMonth,
    # Rolling files: how many months before the current one newer pulls may overwrite.
    [int]$UpdatePriorMonths = 1,
    # DLAStockOnHand carries columns past the current month; by default they are not stored.
    [switch]$StoreFutureStockMonths,
    # Only when the drop has no Batch{N}Details.txt.
    [long]$BatchNumber = 0,
    [switch]$RequireAllFiles,
    [ValidateSet('Summary','Files','All')][string]$ConsoleDetail = 'Summary',
    [int]$MaxConsoleGroups = 40,
    [switch]$NoThrow
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SQLiteInterop.ps1')
. (Join-Path $PSScriptRoot 'FlisIngest.ps1')
. (Join-Path $PSScriptRoot 'FlisTransform.ps1')

$dbPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DatabasePath)
$sw = [Diagnostics.Stopwatch]::StartNew()

# ---- 1. schema (before anything is written) -------------------------------
$db = Open-SqliteDb -Path $dbPath
try {
    $schemaRun = New-FlisRun -ConsoleDetail $ConsoleDetail
    $applied = @(Initialize-FlisSchema -Database $db -Run $schemaRun)
    if ($applied.Count) { Write-Host ("Schema: applied {0}" -f ($applied -join ', ')) -ForegroundColor Cyan }
} finally { Close-SqliteDb -Database $db }

# ---- 2. raw landing ----------------------------------------------------------
Write-Host '== FLIS raw ingest ==' -ForegroundColor Cyan
$ingest = Invoke-FlisIngest -DropFolder $DropFolder -DatabasePath $dbPath -ReportFolder $ReportFolder -BatchNumber $BatchNumber `
    -RequireAllFiles:$RequireAllFiles -ConsoleDetail $ConsoleDetail -MaxConsoleGroups $MaxConsoleGroups -NoThrow
$landedCount = @($ingest.Files | Where-Object Status -eq 'LANDED').Count

# ---- 3 + 4. typed layer and coverage ----------------------------------------
$tRun = New-FlisRun -ReportFolder $ReportFolder -ConsoleDetail $ConsoleDetail -MaxConsoleGroups $MaxConsoleGroups
$tRun.RunId = $ingest.RunId
$tRun.Batch = $ingest.Batch
if (-not $landedCount) {
    Add-FlisIssue $tRun ERROR '(pipeline)' 'NothingLanded' 'No file landed in the raw layer, so the typed tables were not touched. Fix the raw-ingest errors above and run again.'
} else {
    Write-Host '== FLIS typed tables ==' -ForegroundColor Cyan
    $db = Open-SqliteDb -Path $dbPath
    try {
        # The as-of month for the rolling files.
        if ($AsOfMonth) {
            $asOf = [datetime]::ParseExact($AsOfMonth, 'yyyy-MM', [Globalization.CultureInfo]::InvariantCulture)
            Add-FlisIssue $tRun INFO '(pipeline)' 'AsOf' ("As-of month {0:yyyy-MM} given on the command line." -f $asOf)
        } else {
            $p = (Invoke-SqliteQuery -Database $db -Sql 'SELECT pulled_at FROM fb_batch WHERE batch_number = ?;' -Params @($ingest.Batch))
            $asOf = if ($p.Count -and $p[0].pulled_at) { [datetime]::Parse($p[0].pulled_at, [Globalization.CultureInfo]::InvariantCulture) } else { Get-Date }
            $asOf = Get-Date -Year $asOf.Year -Month $asOf.Month -Day 1
            Add-FlisIssue $tRun INFO '(pipeline)' 'AsOf' ("As-of month {0:yyyy-MM} taken from the newest file timestamp in the drop. If the files were copied (timestamps reset), pass -AsOfMonth yyyy-MM." -f $asOf)
        }
        $model = Get-FlisModel
        $registry = Get-FlisRegistry
        Invoke-FlisTransform -Database $db -Run $tRun -Batch $ingest.Batch -RunId $ingest.RunId -Model $model -Registry $registry `
            -AsOfMonth $asOf -PriorMonths $UpdatePriorMonths -StoreFutureStockMonths:$StoreFutureStockMonths
        Write-Host '== Coverage checks ==' -ForegroundColor Cyan
        Invoke-FlisCoverage -Database $db -Run $tRun -Batch $ingest.Batch -Model $model

        Invoke-SqliteExec -Database $db -Sql 'BEGIN;'
        foreach ($i in $tRun.Issues) {
            Invoke-SqliteNonQuery -Database $db -Sql 'INSERT INTO fb_issue_log (run_id, batch_number, file_stem, severity, code, line, column_name, message) VALUES (?,?,?,?,?,?,?,?);' `
                -Params @($tRun.RunId, $tRun.Batch, $i.File, $i.Severity, $i.Code, $(if ($i.Line) { $i.Line }), $i.Column, $i.Message) | Out-Null
        }
        Invoke-SqliteExec -Database $db -Sql 'COMMIT;'
    } finally { Close-SqliteDb -Database $db }
}

# ---- summary ------------------------------------------------------------------
$tErr = @($tRun.Issues | Where-Object Severity -eq 'ERROR').Count
$tWarn = @($tRun.Issues | Where-Object Severity -eq 'WARN').Count
Write-Host '== Typed-layer summary ==' -ForegroundColor Cyan
if ($tErr -or $tWarn) { Write-FlisDigest -Run $tRun -Issues $tRun.Issues -Max $MaxConsoleGroups }
if ($ConsoleDetail -ne 'Summary') { Write-FlisDigest -Run $tRun -Issues @($tRun.Issues | Where-Object Severity -eq 'INFO') -IncludeInfo -Max $MaxConsoleGroups }
if ($ingest.ReportPath) {
    # Append the typed-layer findings to the same report file.
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(''); $lines.Add("== Typed layer and coverage (run $($tRun.RunId)) ==")
    foreach ($g in @(Get-FlisIssueGroups -Issues $tRun.Issues)) { $lines.Add((Format-FlisGroup $g)) }
    $lines.Add(''); $lines.Add('-- full typed-layer messages --')
    foreach ($i in $tRun.Issues) { $lines.Add(("  [{0,-5}] {1} {2}: {3}" -f $i.Severity, $i.File, $i.Code, $i.Message)) }
    [IO.File]::AppendAllLines($ingest.ReportPath, $lines)
}
$allErr = $ingest.Errors + $tErr
$allWarn = $ingest.Warnings + $tWarn
$color = if ($allErr) { 'Red' } elseif ($allWarn) { 'Yellow' } else { 'Green' }
Write-Host ("FLIS pipeline batch {0}: {1} error(s), {2} warning(s), {3:N1}s. Database: {4}. Report: {5}" -f $ingest.Batch, $allErr, $allWarn, $sw.Elapsed.TotalSeconds, $dbPath, $ingest.ReportPath) -ForegroundColor $color

$result = [PSCustomObject]@{ RunId = $ingest.RunId; Batch = $ingest.Batch; Errors = $allErr; Warnings = $allWarn; Ingest = $ingest; TransformIssues = $tRun.Issues; ReportPath = $ingest.ReportPath; DatabasePath = $dbPath }
if ($allErr -and -not $NoThrow) { throw ("FLIS pipeline finished with {0} error(s) -- see above, fb_issue_log (run {1}) or {2}." -f $allErr, $ingest.RunId, $ingest.ReportPath) }
$result
