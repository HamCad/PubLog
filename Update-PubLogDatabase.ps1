#requires -Version 5.1
<#
.SYNOPSIS
  PUB LOG / FED LOG monthly download -> publog.db (separate reference
  database). Independent of the FLIS pipeline and of sustainment.db.

.DESCRIPTION
  Decodes the .TAB tables straight out of PublogDVD.zip (or an extracted
  folder) into pl_* tables, loads the month's NIIN/CAGE change lists, and
  rebuilds the part-number join helpers, then (re)builds publog_alt.db for
  the Alternate Part Finder (docs/ALTERNATE_PARTS.md; -SkipAltPartIndex to
  skip). Tables and the index already built from the same cut are skipped
  unless -Force. See PubLogEtl.ps1 and docs/PUBLOG_*.md.

.EXAMPLE
  .\Update-PubLogDatabase.ps1 -Source D:\Downloads\PublogDVD.zip
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Source,
    [string]$DatabasePath = (Join-Path $PSScriptRoot 'publog.db'),
    [string[]]$Tables,
    [switch]$IncludeSearchPicklists,
    [switch]$Force,
    # Skip (re)building publog_alt.db for the Alternate Part Finder.
    [switch]$SkipAltPartIndex
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PubLogEtl.ps1')

$sw = [Diagnostics.Stopwatch]::StartNew()
Write-Host "== PUB LOG: decoding $Source into $DatabasePath ==" -ForegroundColor Cyan
$params = @{ Source = $Source; DatabasePath = $DatabasePath; IncludeSearchPicklists = $IncludeSearchPicklists; Force = $Force }
if ($Tables) { $params.Tables = $Tables }
$loaded = @(Import-PubLogRaw @params)

$db = Open-SqliteDb -Path (Get-FullPath $DatabasePath)
$hasLinks = (Invoke-SqliteQuery -Database $db -Sql "SELECT 1 FROM sqlite_master WHERE name = 'pl_part_norm';").Count -gt 0
Close-SqliteDb -Database $db
if ($loaded.Count -gt 0 -or -not $hasLinks -or $Force) {
    Write-Host '== PUB LOG: change lists ==' -ForegroundColor Cyan
    Import-PubLogChanges -Source $Source -DatabasePath $DatabasePath
    Write-Host '== PUB LOG: part-number join helpers ==' -ForegroundColor Cyan
    Build-PubLogLinks -DatabasePath $DatabasePath
} else {
    Write-Host '  Nothing new loaded; change lists and join helpers are current.'
}
if (-not $SkipAltPartIndex) {
    . (Join-Path $PSScriptRoot 'PubLogAlternates.ps1')
    Build-AltPartIndex -PubLogDatabase $DatabasePath -Force:$Force
}
Write-Host ("PUB LOG done in {0:N1} min: {1} table(s) loaded this run." -f $sw.Elapsed.TotalMinutes, $loaded.Count) -ForegroundColor Green
