#requires -Version 5.1
<#
.SYNOPSIS
  Regenerates the single self-contained sustainment dashboard HTML from
  sustainment.db, optionally running the FLIS pipeline first.

.DESCRIPTION
  1. (with -DropFolder) Runs Invoke-FlisPipeline.ps1: raw FLIS .txt files in
     the drop folder -> sustainment.db (schema migrations, raw landing,
     typed tables, coverage checks; problems reported on the console, in
     fb_issue_log and in logs\). Stops here if the pipeline reports errors,
     unless -ContinueOnPipelineErrors.
  2. Queries the dashboard views and writes ONE html file with the data
     embedded -- no fetch(), no web server; it opens straight off disk.

  The legacy CSV loader is gone: the FLIS bulk download is the only feeder
  of sustainment.db. The PUB LOG reference database is maintained separately
  by Update-PubLogDatabase.ps1.

.NOTES
  Dashboard rework (insufficient-data state, completeness banner, lead-time
  flag) is tracked in Gitea #17; until then the views in
  migrations\002_dashboard_views.sql feed the existing template.
#>

[CmdletBinding()]
param(
    [string]$DatabasePath = (Join-Path $PSScriptRoot 'sustainment.db'),
    [string]$DropFolder,
    [string]$TemplatePath = (Join-Path $PSScriptRoot 'dashboard_template.html'),
    [string]$OutputHtml   = (Join-Path $PSScriptRoot 'sustainment_dashboard.html'),
    [string]$AsOfMonth,
    [switch]$ContinueOnPipelineErrors,
    [switch]$SkipSelfTest
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SQLiteInterop.ps1')

if (-not $SkipSelfTest) {
    Write-Host "== Step 0: verifying winsqlite3.dll ==" -ForegroundColor Cyan
    if (-not (Test-SqliteSetup)) {
        throw "SQLite self-test failed -- fix this before loading real data (see messages above)."
    }
}

if ($DropFolder) {
    Write-Host "== Step 1: FLIS pipeline from $DropFolder ==" -ForegroundColor Cyan
    $pp = @{ DropFolder = $DropFolder; DatabasePath = $DatabasePath; NoThrow = $true }
    if ($AsOfMonth) { $pp.AsOfMonth = $AsOfMonth }
    $pipeline = & (Join-Path $PSScriptRoot 'Invoke-FlisPipeline.ps1') @pp
    if ($pipeline.Errors -and -not $ContinueOnPipelineErrors) {
        throw ("The FLIS pipeline reported {0} error(s); the dashboard was NOT regenerated. Fix them (see {1}) or rerun with -ContinueOnPipelineErrors." -f $pipeline.Errors, $pipeline.ReportPath)
    }
}

$dbFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DatabasePath)
if (-not (Test-Path -LiteralPath $dbFull)) {
    throw "Database not found: $dbFull -- run with -DropFolder <folder with the FLIS .txt files> to build it."
}
$database = Open-SqliteDb -Path $dbFull
$hasSchema = (Invoke-SqliteQuery -Database $database -Sql "SELECT 1 FROM sqlite_master WHERE name = 'schema_migrations';").Count -gt 0
if (-not $hasSchema) {
    Close-SqliteDb -Database $database
    throw "$dbFull was not built by the FLIS pipeline (no schema_migrations table) -- it is probably a legacy database. Run with -DropFolder to build a new one."
}

Write-Host "== Step 2: querying dashboard data ==" -ForegroundColor Cyan
$items     = Invoke-SqliteQuery -Database $database -Sql 'SELECT * FROM vw_item_summary;'
$suppliers = Invoke-SqliteQuery -Database $database -Sql 'SELECT * FROM vw_supplier_base_risk;'
$gaps      = Invoke-SqliteQuery -Database $database -Sql 'SELECT * FROM vw_procurement_gap;'
$prices    = Invoke-SqliteQuery -Database $database -Sql 'SELECT * FROM vw_price_history;'
$noAward   = Invoke-SqliteQuery -Database $database -Sql 'SELECT * FROM vw_solicitation_no_award;'
$monthly   = Invoke-SqliteQuery -Database $database -Sql 'SELECT * FROM vw_monthly_procurement ORDER BY niin, report_year, report_month;'
$stock     = Invoke-SqliteQuery -Database $database -Sql 'SELECT niin, report_year, report_month, stock_on_hand_qty FROM stock_on_hand_monthly ORDER BY niin, report_year, report_month;'
$xref      = Invoke-SqliteQuery -Database $database -Sql 'SELECT x.niin, x.part_number, x.cage_code, v.vendor_name FROM item_part_number_xref x LEFT JOIN vendor_cage v ON v.cage_code = x.cage_code ORDER BY x.niin;'
$awards    = Invoke-SqliteQuery -Database $database -Sql 'SELECT c.niin, c.contract_number, c.cage_code, v.vendor_name, c.award_date, c.quantity, c.unit_price FROM contract_award_line c LEFT JOIN vendor_cage v ON v.cage_code = c.cage_code ORDER BY c.niin, c.award_date;'
Write-Host "  items=$($items.Count) suppliers=$($suppliers.Count) gaps=$($gaps.Count) prices=$($prices.Count) noAward=$($noAward.Count) monthly=$($monthly.Count) stock=$($stock.Count) xref=$($xref.Count) awards=$($awards.Count)"

Close-SqliteDb -Database $database

Write-Host "== Step 3: writing dashboard ==" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "Dashboard template not found: $TemplatePath"
}

# Build the payload as one hashtable, then serialize with ConvertTo-Json's
# -InputObject parameter rather than the pipe. Piping a PowerShell array
# into ConvertTo-Json makes it enumerate the array one item at a time, so a
# result set with exactly one row silently collapses from a JSON array
# into a bare JSON object -- -InputObject passes the array by reference
# instead and serializes it correctly at any length, including zero.
$payload = [ordered]@{
    generatedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    items       = @($items)
    suppliers   = @($suppliers)
    gaps        = @($gaps)
    prices      = @($prices)
    noAward     = @($noAward)
    monthly     = @($monthly)
    stock       = @($stock)
    xref        = @($xref)
    awards      = @($awards)
}
$json = ConvertTo-Json -InputObject $payload -Depth 6 -Compress

$template = Get-Content -LiteralPath $TemplatePath -Raw
$needle = '"__DASHBOARD_DATA_JSON__"'
if ($template -notmatch [regex]::Escape($needle)) {
    throw "Template is missing the $needle placeholder -- did the template get edited?"
}
$html = $template.Replace($needle, $json)
Set-Content -LiteralPath $OutputHtml -Value $html -Encoding UTF8

Write-Host "Done. Open $OutputHtml directly in any browser -- no server and no internet connection needed." -ForegroundColor Green
