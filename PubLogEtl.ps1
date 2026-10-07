#requires -Version 5.1
<#
.SYNOPSIS
  PUB LOG / FED LOG monthly ETL: raw .TAB tables -> publog.db (full catalog)
  -> sustainment.db (the dashboard warehouse, scoped to a NIIN watchlist).

.DESCRIPTION
  Stage 1  Import-PubLogRaw
           Every selected .TAB is decoded (PubLogReader.ps1) and bulk-loaded
           into publog.db as pl_<table> -- one TEXT column per FLIS column,
           physical column order, empty fields as NULL. Reads straight from
           PublogDVD.zip or an extracted folder. Each table is loaded into a
           staging table, row-count-checked against the .TAB header, then
           swapped in -- a failed or interrupted month never replaces a good
           one. Re-running is cheap: a table whose source (pub date + size)
           matches pl_load_log is skipped unless -Force.

  Stage 2  Import-PubLogChanges
           The monthly CHANGES\NIIN.zip / CAGE.zip key lists (adds, changes,
           deletes) go into pl_change_log, so "what changed for my items this
           month" is a join, not a diff.

  Stage 3  Build-PubLogLinks
           Engineering-join helpers in publog.db: pl_part_norm (every FLIS
           part number with a punctuation-insensitive match key) and the
           pl_vw_nsn / pl_vw_part views. This is what CAD, drawing and IPB
           databases join against.

  Stage 4  Update-SustainmentFromPubLog  -- RETIRED (not called by the
           build). Kept only for its FLIS<->PUBLOG column mappings, used to
           cross-validate the real FLIS files (Gitea #18). Originally:
           refreshed the dashboard tables in sustainment.db (item_niin,
           vendor_cage, item_part_number_xref, management, MOE, phrases,
           characteristics, standardization, freight, packaging, code
           tables) for every NIIN in niin_watchlist. Watchlist sources:
           data\niin_watchlist.csv, NIINs already in item_niin, and
           part numbers resolved with Resolve-PubLogPartNumber.

  Everything is plain PowerShell 5.1 + SQLiteInterop.ps1 (winsqlite3.dll)
  + PubLogReader.ps1. Nothing to install, no vendor executables.
#>

# Always dot-source (functions are per-scope); a repeated identical Add-Type
# is a no-op, and PubLogReader.ps1 guards its own.
. (Join-Path $PSScriptRoot 'SQLiteInterop.ps1')
. (Join-Path $PSScriptRoot 'PubLogReader.ps1')
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# SQLite resolves relative paths against the process working directory, not
# PowerShell's current location -- always hand it a full path.
function Get-FullPath { param([string]$Path) $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path) }

# Loaded by default: every view table plus the picklists that carry data no
# view has (P_FLIS_NSN is the only source of an NSN's FSC and item name).
# The three big search picklists are derivable from the views and are only
# loaded with -IncludeSearchPicklists.
$Script:PubLogDefaultTables = @(
    'P_FLIS_NSN','P_CAGE','P_HELP','P_H2_PICK','P_H3_PICK','P_H6_PICK',
    'V_FLIS_IDENTIFICATION','V_FLIS_STANDARDIZATION','V_FLIS_CANCELLED_NIIN',
    'V_FLIS_MANAGEMENT','V_FLIS_MANAGEMENT_FUTURE','V_FLIS_PHRASE','V_MOE_RULE',
    'V_MGMT_AIR_FORCE','V_MGMT_ARMY','V_MGMT_COAST_GUARD','V_MGMT_MARINE_CORPS','V_MGMT_NAVY','V_SOCOM_MANAGEMENT',
    'V_FLIS_PART','V_FREIGHT','V_FLIS_PACKAGING_1','V_FLIS_PACKAGING_2','V_FLIS_PACKAGING_3',
    'V_CHARACTERISTICS','V_CAGE_ADDRESS',
    'V_H2_FSG','V_H2_FSC','V_H2_FSC_INC','V_H3_AMMUNITION','V_H6_NAME_INC','V_H6_MODIFIER','V_H6_COLLOQUIAL','V_FSC_IMM','V_COLLOQUIAL_NAME',
    'V_ITEM_IDENTIFICATION_HISTORY','V_MANAGEMENT_HISTORY','V_REFERENCE_NUMBER_HISTORY'
)
$Script:PubLogSearchPicklists = @('P_PART_PICK','P_CHARACTERISTICS_PICK','P_HISTORY_PICK')

# Extra indexes beyond the primary key, for the joins people actually run.
$Script:PubLogExtraIndexes = @{
    'V_FLIS_PART'           = @(@('cage_code','part_number'), @('part_number'))
    'V_CAGE_ADDRESS'        = @()
    'P_HELP'                = @(@('code'))
    'V_FLIS_STANDARDIZATION'= @(@('related_nsn'))
    'V_FLIS_CANCELLED_NIIN' = @(@('cancelled_niin'))
    'V_REFERENCE_NUMBER_HISTORY' = @(@('cage_code','part_number'))
}

# ---------------------------------------------------------------------------
# Source access: a folder of .TAB files or PublogDVD.zip, same interface.
# ---------------------------------------------------------------------------

function Open-PubLogSource {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    $src = [PSCustomObject]@{ Path = $full; Zip = $null; Tables = @{}; ChangeFiles = @{} }
    if (Test-Path -LiteralPath $full -PathType Container) {
        foreach ($f in Get-ChildItem -LiteralPath $full -Filter *.TAB) {
            $src.Tables[$f.BaseName.ToUpperInvariant()] = [PSCustomObject]@{ Name = $f.Name; Size = $f.Length; File = $f.FullName; Entry = $null }
        }
        foreach ($n in 'NIIN','CAGE') {
            $c = Join-Path $full "CHANGES\$n.zip"
            if (Test-Path -LiteralPath $c) { $src.ChangeFiles[$n] = $c }
        }
    }
    else {
        $src.Zip = [System.IO.Compression.ZipFile]::OpenRead($full)
        foreach ($e in $src.Zip.Entries) {
            if ($e.FullName -match '^([^/\\]+)\.TAB$') {
                $src.Tables[$matches[1].ToUpperInvariant()] = [PSCustomObject]@{ Name = $e.Name; Size = $e.Length; File = $null; Entry = $e }
            }
            elseif ($e.FullName -match '^CHANGES[/\\](NIIN|CAGE)\.zip$') { $src.ChangeFiles[$matches[1]] = $e }
        }
    }
    if ($src.Tables.Count -eq 0) { throw "No .TAB files found in $full" }
    $src
}

function Close-PubLogSource {
    param($Source)
    if ($Source -and $Source.Zip) { $Source.Zip.Dispose() }
}

function Read-EntryBytes {
    param($Entry)
    $ms = New-Object System.IO.MemoryStream ([int]$Entry.Length)
    $s = $Entry.Open()
    try { $s.CopyTo($ms) } finally { $s.Dispose() }
    , $ms.ToArray()   # leading comma: an empty entry must stay an empty byte[], not become `$null
}

function Get-PubLogHeaderFromSource {
    # Decrypts just the 8 KB header -- no decompression of the table body.
    param($Source, [string]$Table)
    $t = $Source.Tables[$Table.ToUpperInvariant()]
    if (-not $t) { throw "Table $Table not found in $($Source.Path)" }
    $buf = New-Object byte[] 8192
    $s = if ($t.File) { [IO.File]::OpenRead($t.File) } else { $t.Entry.Open() }
    try {
        $got = 0
        while ($got -lt 8192) { $n = $s.Read($buf, $got, 8192 - $got); if ($n -le 0) { break }; $got += $n }
    } finally { $s.Dispose() }
    [PubLog.ImdTable]::HeaderOnly($buf, $t.Name)
}

function Open-PubLogTableFromSource {
    param($Source, [string]$Table)
    $t = $Source.Tables[$Table.ToUpperInvariant()]
    if (-not $t) { throw "Table $Table not found in $($Source.Path)" }
    if ($t.File) { return [PubLog.ImdTable]::Open($t.File) }
    [PubLog.ImdTable]::FromBytes((Read-EntryBytes $t.Entry), "$($Source.Path)!$($t.Name)")
}

# ---------------------------------------------------------------------------
# Stage 1 -- raw load
# ---------------------------------------------------------------------------

function Initialize-PubLogDb {
    param([IntPtr]$Database)
    Invoke-SqliteExec -Database $Database -Sql @'
CREATE TABLE IF NOT EXISTS pl_load_log (
    table_name   TEXT NOT NULL,
    pub_date     TEXT,
    rows_declared INTEGER,
    rows_loaded  INTEGER,
    source_name  TEXT,
    source_size  INTEGER,
    columns      TEXT,
    seconds      REAL,
    loaded_at    TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS pl_change_log (
    pub_date    TEXT NOT NULL,
    key_type    TEXT NOT NULL,     -- NIIN or CAGE
    change_type TEXT NOT NULL,     -- add / chg / del
    key_value   TEXT NOT NULL,
    related     TEXT               -- CAGE files list associated CAGEs after the key
);
CREATE INDEX IF NOT EXISTS ix_pl_change_log_key ON pl_change_log (key_type, key_value);
'@
}

function Get-PubLogRawTableName { param([string]$Table) 'pl_' + $Table.ToLowerInvariant() }

function Import-PubLogRaw {
    # Decode + bulk-load .TAB tables into publog.db.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$DatabasePath,
        [string[]]$Tables,
        [switch]$IncludeSearchPicklists,
        [switch]$Force,
        [int]$BatchRows = 500000
    )
    if (-not $Tables) {
        $Tables = $Script:PubLogDefaultTables
        if ($IncludeSearchPicklists) { $Tables += $Script:PubLogSearchPicklists }
    }
    $src = Open-PubLogSource -Path $Source
    $db = Open-SqliteDb -Path (Get-FullPath $DatabasePath)
    try {
        # Bulk-load settings: this database is a rebuildable mirror, so trade
        # crash durability during the load for speed. The staging/swap design
        # means an interrupted load leaves the previous month intact.
        Invoke-SqliteExec -Database $db -Sql 'PRAGMA journal_mode = WAL; PRAGMA synchronous = OFF; PRAGMA temp_store = MEMORY; PRAGMA cache_size = -262144;'
        Initialize-PubLogDb -Database $db
        $results = New-Object System.Collections.Generic.List[object]
        foreach ($table in $Tables) {
            $table = $table.ToUpperInvariant()
            $entry = $src.Tables[$table]
            if (-not $entry) { Write-Warning "  $table not in source, skipping"; continue }
            $raw = Get-PubLogRawTableName $table
            $sw = [Diagnostics.Stopwatch]::StartNew()
            if (-not $Force) {
                $h = Get-PubLogHeaderFromSource -Source $src -Table $table
                $prev = Invoke-SqliteQuery -Database $db -Sql 'SELECT rows_loaded FROM pl_load_log WHERE table_name = ? AND pub_date = ? AND source_size = ? ORDER BY loaded_at DESC LIMIT 1;' -Params @($table, $h.PubDate, [long]$entry.Size)
                $exists = Invoke-SqliteQuery -Database $db -Sql "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?;" -Params @($raw)
                if ($prev.Count -gt 0 -and $exists.Count -gt 0) {
                    Write-Host ("  {0,-30} already loaded for {1} ({2:N0} rows) -- skipped" -f $table, $h.PubDate, $prev[0].rows_loaded)
                    continue
                }
            }
            $t = Open-PubLogTableFromSource -Source $src -Table $table
            try {
                $cols = @($t.ColumnNames | ForEach-Object { $_.ToLowerInvariant() })
                $stage = "${raw}__staging"
                $colDefs = ($cols | ForEach-Object { "`"$_`" TEXT" }) -join ', '
                Invoke-SqliteExec -Database $db -Sql "DROP TABLE IF EXISTS `"$stage`"; CREATE TABLE `"$stage`" ($colDefs);"

                $n = [PubLog.SqliteBulk]::Load($db, $t, $stage, $BatchRows)
                if ($t.DeclaredRows -ge 0 -and $n -ne $t.DeclaredRows) {
                    Invoke-SqliteExec -Database $db -Sql "DROP TABLE IF EXISTS `"$stage`";"
                    throw "$table decoded $n rows but its header declares $($t.DeclaredRows) -- not swapping it in. Previous month's $raw is untouched."
                }

                # Swap in, then index under the final name.
                # legacy_alter_table: newer SQLite re-validates every view on RENAME and
                # would refuse while pl_vw_* points at the table just dropped.
                Invoke-SqliteExec -Database $db -Sql "PRAGMA legacy_alter_table = ON; BEGIN; DROP TABLE IF EXISTS `"$raw`"; ALTER TABLE `"$stage`" RENAME TO `"$raw`"; COMMIT; PRAGMA legacy_alter_table = OFF;"
                $key = $t.PrimaryKey
                $indexSets = @()
                if ($key -and $key -ne 'None') { $indexSets += ,@($key.ToLowerInvariant()) }
                if ($Script:PubLogExtraIndexes.ContainsKey($table)) { $indexSets += $Script:PubLogExtraIndexes[$table] }
                foreach ($set in $indexSets) {
                    $ixName = "ix_${raw}_" + ($set -join '_')
                    $ixCols = ($set | ForEach-Object { "`"$_`"" }) -join ', '
                    Invoke-SqliteExec -Database $db -Sql "CREATE INDEX IF NOT EXISTS `"$ixName`" ON `"$raw`" ($ixCols);"
                }

                $secs = [math]::Round($sw.Elapsed.TotalSeconds, 1)
                Invoke-SqliteNonQuery -Database $db -Sql 'INSERT INTO pl_load_log (table_name, pub_date, rows_declared, rows_loaded, source_name, source_size, columns, seconds) VALUES (?,?,?,?,?,?,?,?);' `
                    -Params @($table, $t.PubDate, $t.DeclaredRows, $n, $entry.Name, [long]$entry.Size, ($cols -join '|'), $secs) | Out-Null
                Write-Host ("  {0,-30} {1,12:N0} rows  {2,7:N1}s  -> {3}" -f $table, $n, $secs, $raw)
                $results.Add([PSCustomObject]@{ Table = $table; Rows = $n; Seconds = $secs; PubDate = $t.PubDate })
            }
            finally { $t.Dispose(); [GC]::Collect() }
        }
        Invoke-SqliteExec -Database $db -Sql 'PRAGMA wal_checkpoint(TRUNCATE); PRAGMA optimize;'
        $results
    }
    finally {
        Close-SqliteDb -Database $db
        Close-PubLogSource $src
    }
}

# ---------------------------------------------------------------------------
# Stage 2 -- monthly change lists
# ---------------------------------------------------------------------------

function Import-PubLogChanges {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Source,
        [Parameter(Mandatory)][string]$DatabasePath
    )
    $src = Open-PubLogSource -Path $Source
    $db = Open-SqliteDb -Path (Get-FullPath $DatabasePath)
    try {
        Initialize-PubLogDb -Database $db
        $pubRow = Invoke-SqliteQuery -Database $db -Sql "SELECT pub_date FROM pl_load_log ORDER BY loaded_at DESC LIMIT 1;"
        $pub = if ($pubRow.Count) { $pubRow[0].pub_date } else { (Get-Date).ToString('MMM yyyy').ToUpperInvariant() }
        Invoke-SqliteNonQuery -Database $db -Sql 'DELETE FROM pl_change_log WHERE pub_date = ?;' -Params @($pub) | Out-Null
        foreach ($kind in 'NIIN','CAGE') {
            $cf = $src.ChangeFiles[$kind]
            if (-not $cf) { Write-Warning "  CHANGES\$kind.zip not in source"; continue }
            $bytes = if ($cf -is [string]) { [IO.File]::ReadAllBytes($cf) } else { Read-EntryBytes $cf }
            $inner = New-Object System.IO.Compression.ZipArchive (New-Object System.IO.MemoryStream (,$bytes))
            try {
                Invoke-SqliteExec -Database $db -Sql 'BEGIN;'
                $count = 0
                foreach ($e in $inner.Entries) {
                    if ($e.Name -notmatch '\.(add|chg|del)$') { continue }
                    $count += [PubLog.SqliteBulk]::LoadChangeList($db, (Read-EntryBytes $e), $pub, $kind, $matches[1])
                }
                Invoke-SqliteExec -Database $db -Sql 'COMMIT;'
                Write-Host ("  pl_change_log <- {0} {1:N0} keys ({2})" -f $kind, $count, $pub)
            }
            catch { Invoke-SqliteExec -Database $db -Sql 'ROLLBACK;'; throw }
            finally { $inner.Dispose() }
        }
    }
    finally {
        Close-SqliteDb -Database $db
        Close-PubLogSource $src
    }
}

# ---------------------------------------------------------------------------
# Stage 3 -- engineering join helpers
# ---------------------------------------------------------------------------

# Part-number match key: upper-case with common punctuation and spaces
# removed, so "MS 21042-3", "MS21042-3" and "ms21042 3" all meet. Built with
# nested replace() because winsqlite3 has no regex.
function Get-PartNormSql {
    param([string]$Expr)
    $s = "upper($Expr)"
    foreach ($ch in @('-',' ','/','.',',','_','#','(',')',"''",'"','*','+',':',';')) { $s = "replace($s,'$ch','')" }
    $s
}

function Build-PubLogLinks {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$DatabasePath)
    $db = Open-SqliteDb -Path (Get-FullPath $DatabasePath)
    try {
        $norm = Get-PartNormSql 'part_number'
        Invoke-SqliteExec -Database $db -Sql @"
PRAGMA journal_mode = WAL; PRAGMA synchronous = OFF; PRAGMA temp_store = MEMORY;
DROP TABLE IF EXISTS pl_part_norm;
CREATE TABLE pl_part_norm AS
    SELECT $norm AS pn_norm, part_number, cage_code, niin, rncc, rnvc
    FROM pl_v_flis_part WHERE part_number IS NOT NULL;
CREATE INDEX ix_pl_part_norm_pn ON pl_part_norm (pn_norm);
CREATE INDEX ix_pl_part_norm_cage_pn ON pl_part_norm (cage_code, pn_norm);

DROP VIEW IF EXISTS pl_vw_nsn;
CREATE VIEW pl_vw_nsn AS
SELECT n.fsc || n.niin AS nsn, n.fsc, n.niin, n.item_name, n.inc, n.sos, n.end_item_name,
       i.dmil AS demil, i.crit_cd, i.hmic, i.esd_emi, i.pmic, i.niin_asgmt, i.schedule_b,
       c.fsc_title
FROM pl_p_flis_nsn n
LEFT JOIN pl_v_flis_identification i ON i.niin = n.niin
LEFT JOIN pl_v_h2_fsc c ON c.fsc = n.fsc;

DROP VIEW IF EXISTS pl_vw_part;
CREATE VIEW pl_vw_part AS
SELECT p.niin, n.fsc || p.niin AS nsn, n.item_name, p.part_number, p.cage_code,
       a.company_name, p.rncc, p.rnvc, p.cage_status, p.msds
FROM pl_v_flis_part p
LEFT JOIN pl_p_flis_nsn n ON n.niin = p.niin
LEFT JOIN pl_v_cage_address a ON a.cage_code = p.cage_code;
"@
        $c = Invoke-SqliteQuery -Database $db -Sql 'SELECT COUNT(*) AS n FROM pl_part_norm;'
        Write-Host ("  pl_part_norm: {0:N0} part numbers indexed for engineering joins" -f $c[0].n)
    }
    finally { Close-SqliteDb -Database $db }
}

function Resolve-PubLogPartNumber {
    # Look up engineering part numbers (optionally with CAGE) in FLIS.
    # -InputCsv needs a part_number column, cage_code optional. Returns one
    # row per match; unmatched inputs come back with niin = $null so a
    # report shows coverage gaps, not just hits.
    [CmdletBinding(DefaultParameterSetName = 'One')]
    param(
        [Parameter(Mandatory)][string]$PubLogDb,
        [Parameter(ParameterSetName = 'One', Mandatory)][string]$PartNumber,
        [Parameter(ParameterSetName = 'One')][string]$CageCode,
        [Parameter(ParameterSetName = 'Csv', Mandatory)][string]$InputCsv
    )
    $inputs = if ($PSCmdlet.ParameterSetName -eq 'Csv') { @(Import-Csv -LiteralPath $InputCsv) } else { @([PSCustomObject]@{ part_number = $PartNumber; cage_code = $CageCode }) }
    $db = Open-SqliteDb -Path (Get-FullPath $PubLogDb)
    try {
        $normOf = { param($s) ($s.ToUpperInvariant() -replace '[-\s/.,_#()''"*+:;]', '') }
        foreach ($in in $inputs) {
            $pn = [string]$in.part_number; if (-not $pn) { continue }
            $cage = [string]$in.cage_code
            $key = & $normOf $pn
            $rows = if ($cage) {
                Invoke-SqliteQuery -Database $db -Sql 'SELECT niin, part_number, cage_code, rncc, rnvc FROM pl_part_norm WHERE cage_code = ? AND pn_norm = ?;' -Params @($cage.ToUpperInvariant(), $key)
            } else {
                Invoke-SqliteQuery -Database $db -Sql 'SELECT niin, part_number, cage_code, rncc, rnvc FROM pl_part_norm WHERE pn_norm = ?;' -Params @($key)
            }
            if ($rows.Count -eq 0) {
                [PSCustomObject]@{ input_part_number = $pn; input_cage_code = $cage; niin = $null; part_number = $null; cage_code = $null; rncc = $null; rnvc = $null; match = 'none' }
            }
            foreach ($r in $rows) {
                $exact = ($r.part_number -eq $pn) -and (-not $cage -or $r.cage_code -eq $cage)
                [PSCustomObject]@{ input_part_number = $pn; input_cage_code = $cage; niin = $r.niin; part_number = $r.part_number; cage_code = $r.cage_code; rncc = $r.rncc; rnvc = $r.rnvc; match = $(if ($exact) { 'exact' } else { 'normalized' }) }
            }
        }
    }
    finally { Close-SqliteDb -Database $db }
}

# ---------------------------------------------------------------------------
# Stage 4 -- refresh the dashboard warehouse for watchlisted NIINs
# ---------------------------------------------------------------------------

# FLIS dates: DD-MON-YYYY, DD-MON-YY, or YYDDD (Julian, packaging SPI date).
# Two-digit years pivot one year past today: '67' -> 1967, '22' -> 2022.
function Get-FlisDateSql {
    param([string]$Expr)
    $pivot = ((Get-Date).Year % 100) + 1
    $mm = "printf('%02d', (instr('JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC', upper(substr($Expr,4,3))) + 2) / 3)"
    @"
(CASE
  WHEN $Expr GLOB '[0-9][0-9]-[A-Za-z][A-Za-z][A-Za-z]-[0-9][0-9][0-9][0-9]'
    THEN substr($Expr,8,4) || '-' || $mm || '-' || substr($Expr,1,2)
  WHEN $Expr GLOB '[0-9][0-9]-[A-Za-z][A-Za-z][A-Za-z]-[0-9][0-9]'
    THEN (CASE WHEN CAST(substr($Expr,8,2) AS INTEGER) > $pivot THEN '19' ELSE '20' END) || substr($Expr,8,2) || '-' || $mm || '-' || substr($Expr,1,2)
  ELSE NULL END)
"@
}

function Get-JulianDateSql {
    param([string]$Expr)
    $pivot = ((Get-Date).Year % 100) + 1
    "(CASE WHEN $Expr GLOB '[0-9][0-9][0-9][0-9][0-9]' AND CAST(substr($Expr,3,3) AS INTEGER) BETWEEN 1 AND 366 THEN date((CASE WHEN CAST(substr($Expr,1,2) AS INTEGER) > $pivot THEN '19' ELSE '20' END) || substr($Expr,1,2) || '-01-01', '+' || (CAST(substr($Expr,3,3) AS INTEGER) - 1) || ' days') ELSE NULL END)"
}

function Initialize-SustainmentWatchlist {
    param([IntPtr]$Database)
    Invoke-SqliteExec -Database $Database -Sql @'
CREATE TABLE IF NOT EXISTS niin_watchlist (
    niin      VARCHAR(9) PRIMARY KEY,
    source    VARCHAR(40),              -- csv / item_niin / part_resolve / manual
    note      TEXT,
    added_at  TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);
'@
}

function Add-NiinWatchlist {
    # Adds NIINs (or 13-digit NSNs) to the watchlist. Accepts a CSV with a
    # niin or nsn column, or values on the pipeline.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SustainmentDb,
        [string]$Csv,
        [Parameter(ValueFromPipeline)][string[]]$Niin,
        [string]$Source = 'manual'
    )
    begin { $all = New-Object System.Collections.Generic.List[string] }
    process { foreach ($n in $Niin) { if ($n) { $all.Add($n) } } }
    end {
        if ($Csv) {
            if (-not (Test-Path -LiteralPath $Csv)) { Write-Warning "Watchlist CSV not found: $Csv"; return }
            foreach ($r in Import-Csv -LiteralPath $Csv) {
                $v = if ($r.PSObject.Properties['niin']) { $r.niin } elseif ($r.PSObject.Properties['nsn']) { $r.nsn } else { $null }
                if ($v) { $all.Add($v) }
            }
        }
        $db = Open-SqliteDb -Path (Get-FullPath $SustainmentDb)
        try {
            Initialize-SustainmentWatchlist -Database $db
            $added = 0
            Invoke-SqliteExec -Database $db -Sql 'BEGIN;'
            foreach ($v in $all) {
                $digits = ($v -replace '[^0-9]', '')
                if ($digits.Length -eq 13) { $digits = $digits.Substring(4) }   # NSN -> NIIN
                if ($digits.Length -ne 9) { Write-Warning "  not a NIIN/NSN: '$v'"; continue }
                $added += Invoke-SqliteNonQuery -Database $db -Sql 'INSERT OR IGNORE INTO niin_watchlist (niin, source) VALUES (?, ?);' -Params @($digits, $Source)
            }
            Invoke-SqliteExec -Database $db -Sql 'COMMIT;'
            Write-Host "  niin_watchlist: $added added ($($all.Count) supplied)"
        }
        finally { Close-SqliteDb -Database $db }
    }
}

# Every raw column Update-SustainmentFromPubLog reads.
$Script:PubLogTransformColumns = [ordered]@{
    pl_p_flis_nsn            = 'niin','fsc','item_name','inc'
    pl_p_cage                = 'cage_code','cage_status'
    pl_p_help                = 'code','literal','description'
    pl_p_h6_pick             = 'inc','item_name'
    pl_v_h2_fsc              = 'fsc','fsc_title'
    pl_v_flis_identification = 'niin','crit_cd','ii','adp','dmil','dmil_int_cd','hmic','esd_emi','pmic','hcc','schedule_b','enac','niin_asgmt'
    pl_v_flis_standardization= 'niin','isc','niin_stat_cd','related_nsn','orig_stdzn_dec','dt_stdzn_dec'
    pl_v_cage_address        = 'cage_code','company_name','street_address_1','street_address_2','city','state','country','zip'
    pl_v_flis_part           = 'niin','part_number','cage_code','rncc','rnvc','rnfc','rnjc','rnsc','rnaac','sadc','dac','hcc','msds'
    pl_v_flis_management     = 'niin','moe','sos','mgmt_ctl','ui','ui_conv_fac','qup','aac','slc','rep_rec_code','ciic','usc','unit_price','effective_date'
    pl_v_flis_phrase         = 'niin','moe','phrs_cd','phrase_statement'
    pl_v_moe_rule            = 'niin','moe_rl','moe_cd','pica','dsor','amc','amsc','nimsc','imc','imca','aac','auth_collab','supp_collab','dt_asgnd'
    pl_v_characteristics     = 'niin','mrc','requirements_statement','clear_text_reply'
    pl_v_freight             = 'niin','acty_cd','integ','nmf_desc','rvc','nmfc','nmfc_sub','ufc','hmc','ltl','wcc','shc','adc','acc','tcc'
    pl_v_flis_packaging_1    = 'niin','pkg_data_source','pica_sica','clng_drying','cush_dun','thk','icq','tos','mop','pres_mat','pkg_cat','unit_cont','unpkg_item_length','unpkg_item_width','unpkg_item_height','unpkg_item_weight','wrap_mat'
    pl_v_flis_packaging_2    = 'niin','pica_sica','inter_cont','lvl_a','lvl_b','lvl_c','opi','spc_mkg','ucl','unit_pack_cube','up_sz_maxl','up_sz_maxb','up_sz_maxh','unit_pack_weight','supplemental_instructions'
    pl_v_flis_packaging_3    = 'niin','pica_sica','cont_nsn','pkg_design_acty','spi_date','spi_no','spi_rev'
}

function Test-PubLogTransformColumns {
    # Returns 'table.column' for every column the transform needs that the
    # loaded PUB LOG tables don't have (empty = all present).
    param([IntPtr]$Database, [string]$Schema = 'main')
    foreach ($table in $Script:PubLogTransformColumns.Keys) {
        $have = @{}
        foreach ($c in (Invoke-SqliteQuery -Database $Database -Sql "PRAGMA $Schema.table_info(`"$table`");")) { $have[$c.name] = $true }
        if ($have.Count -eq 0) { "$table (table not loaded)"; continue }
        foreach ($col in $Script:PubLogTransformColumns[$table]) { if (-not $have[$col]) { "$table.$col" } }
    }
}

function Update-SustainmentFromPubLog {
    # RETIRED -- see the header. PUB LOG must not feed sustainment.db.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SustainmentDb,
        [Parameter(Mandatory)][string]$PubLogDb
    )
    Write-Warning 'Update-SustainmentFromPubLog is retired: PUB LOG does not feed sustainment.db (docs/flis handoff section 0). Running it anyway for cross-validation only.'
    $db = Open-SqliteDb -Path (Get-FullPath $SustainmentDb)
    try {
        Initialize-SustainmentWatchlist -Database $db
        $plPath = (Resolve-Path -LiteralPath $PubLogDb).ProviderPath
        Invoke-SqliteNonQuery -Database $db -Sql 'ATTACH DATABASE ? AS pl;' -Params @($plPath) | Out-Null
        try {
            $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            $effDate  = Get-FlisDateSql 'm.effective_date'
            $asgDate  = Get-FlisDateSql 'i.niin_asgmt'
            $moeDate  = Get-FlisDateSql 'r.dt_asgnd'
            $stdDate  = Get-FlisDateSql 's.dt_stdzn_dec'
            $spiDate  = Get-JulianDateSql 'p3.spi_date'

            # PUB LOG's columns change month to month (see RELNOTES.TXT on each
            # cut). Fail with a precise list instead of a SQL error halfway in.
            $missingCols = @(Test-PubLogTransformColumns -Database $db -Schema 'pl')
            if ($missingCols.Count -gt 0) {
                throw ("PUB LOG columns this transform needs are missing (format change in this cut? check RELNOTES.TXT):`n  " + ($missingCols -join "`n  "))
            }

            # NIINs already tracked from earlier loads join the watchlist, so
            # switching sources never silently drops an item.
            Invoke-SqliteExec -Database $db -Sql "INSERT OR IGNORE INTO niin_watchlist (niin, source) SELECT niin, 'item_niin' FROM item_niin;"
            $w = Invoke-SqliteQuery -Database $db -Sql 'SELECT COUNT(*) AS n FROM niin_watchlist;'
            Write-Host "  watchlist: $($w[0].n) NIINs"

            $sql = @"
BEGIN;
CREATE TEMP TABLE w AS SELECT niin FROM niin_watchlist;
CREATE UNIQUE INDEX temp.ix_w ON w (niin);

-- Code tables are small: refresh them in full.
INSERT OR IGNORE INTO supply_class (fsc_code, fsc_title) SELECT fsc, fsc_title FROM pl.pl_v_h2_fsc;
UPDATE supply_class SET fsc_title = (SELECT f.fsc_title FROM pl.pl_v_h2_fsc f WHERE f.fsc = supply_class.fsc_code)
 WHERE fsc_code IN (SELECT fsc FROM pl.pl_v_h2_fsc);

INSERT OR IGNORE INTO item_name_code (inc, approved_item_name) SELECT inc, item_name FROM pl.pl_p_h6_pick;
UPDATE item_name_code SET approved_item_name = (SELECT h.item_name FROM pl.pl_p_h6_pick h WHERE h.inc = item_name_code.inc)
 WHERE inc IN (SELECT inc FROM pl.pl_p_h6_pick);

-- P_HELP: CODE is 'DOMAIN' (a code family) or 'DOMAIN:VALUE' (one code).
INSERT OR IGNORE INTO ref_code_domain (domain_code, domain_name, description)
  SELECT code, coalesce(literal, code), description FROM pl.pl_p_help WHERE instr(code, ':') = 0;
INSERT OR IGNORE INTO ref_code_domain (domain_code, domain_name)
  SELECT DISTINCT substr(code, 1, instr(code, ':') - 1), substr(code, 1, instr(code, ':') - 1) FROM pl.pl_p_help WHERE instr(code, ':') > 0;
INSERT OR REPLACE INTO ref_code_value (domain_code, code_value, code_description)
  SELECT substr(code, 1, instr(code, ':') - 1), substr(code, instr(code, ':') + 1), coalesce(literal, description, '')
  FROM pl.pl_p_help WHERE instr(code, ':') > 0
  GROUP BY substr(code, 1, instr(code, ':') - 1), substr(code, instr(code, ':') + 1);

-- Item master (one row per NIIN).
CREATE TEMP TABLE src_item AS
SELECT n.niin, n.fsc, n.item_name, n.inc, i.crit_cd, s.isc, s.niin_stat_cd, i.ii, i.adp, i.dmil, i.dmil_int_cd,
       i.hmic, i.esd_emi, i.pmic, i.hcc, i.schedule_b, i.enac, $asgDate AS assign_date
FROM w CROSS JOIN pl.pl_p_flis_nsn n ON n.niin = w.niin
LEFT JOIN pl.pl_v_flis_identification i ON i.niin = n.niin
LEFT JOIN (SELECT niin, MIN(isc) AS isc, MIN(niin_stat_cd) AS niin_stat_cd FROM pl.pl_v_flis_standardization
           WHERE niin IN (SELECT niin FROM w) GROUP BY niin) s ON s.niin = n.niin;
-- winsqlite3 enforces foreign keys: make sure every parent code exists first
-- (some INCs/FSCs on items are absent from the H2/H6 handbook tables).
INSERT OR IGNORE INTO supply_class (fsc_code) SELECT DISTINCT fsc FROM src_item WHERE fsc IS NOT NULL;
INSERT OR IGNORE INTO item_name_code (inc, approved_item_name) SELECT inc, MIN(item_name) FROM src_item WHERE inc IS NOT NULL GROUP BY inc;
UPDATE item_niin SET
  fsc_code = (SELECT fsc FROM src_item x WHERE x.niin = item_niin.niin),
  item_name = (SELECT item_name FROM src_item x WHERE x.niin = item_niin.niin),
  inc = (SELECT inc FROM src_item x WHERE x.niin = item_niin.niin),
  crit_code = (SELECT crit_cd FROM src_item x WHERE x.niin = item_niin.niin),
  isc_code = (SELECT isc FROM src_item x WHERE x.niin = item_niin.niin),
  niin_status_code = (SELECT niin_stat_cd FROM src_item x WHERE x.niin = item_niin.niin),
  tiic_code = (SELECT ii FROM src_item x WHERE x.niin = item_niin.niin),
  adpe_code = (SELECT adp FROM src_item x WHERE x.niin = item_niin.niin),
  demil_code = (SELECT dmil FROM src_item x WHERE x.niin = item_niin.niin),
  demil_integrity_code = (SELECT dmil_int_cd FROM src_item x WHERE x.niin = item_niin.niin),
  hmic_code = (SELECT hmic FROM src_item x WHERE x.niin = item_niin.niin),
  esd_emi_code = (SELECT esd_emi FROM src_item x WHERE x.niin = item_niin.niin),
  precious_metals_code = (SELECT pmic FROM src_item x WHERE x.niin = item_niin.niin),
  hcc_code = (SELECT hcc FROM src_item x WHERE x.niin = item_niin.niin),
  hts_schedule_b_code = (SELECT schedule_b FROM src_item x WHERE x.niin = item_niin.niin),
  enac_code = (SELECT enac FROM src_item x WHERE x.niin = item_niin.niin),
  niin_assign_date = (SELECT assign_date FROM src_item x WHERE x.niin = item_niin.niin),
  source_loaded_at = '$now'
WHERE niin IN (SELECT niin FROM src_item);
INSERT INTO item_niin (niin, fsc_code, item_name, inc, crit_code, isc_code, niin_status_code, tiic_code, adpe_code,
  demil_code, demil_integrity_code, hmic_code, esd_emi_code, precious_metals_code, hcc_code, hts_schedule_b_code,
  enac_code, niin_assign_date, source_loaded_at)
SELECT niin, fsc, item_name, inc, crit_cd, isc, niin_stat_cd, ii, adp, dmil, dmil_int_cd, hmic, esd_emi, pmic, hcc,
  schedule_b, enac, assign_date, '$now'
FROM src_item WHERE niin NOT IN (SELECT niin FROM item_niin);

-- Vendors: every CAGE cross-referenced to a watchlisted NIIN.
CREATE TEMP TABLE src_cage AS
SELECT a.cage_code, a.company_name, a.street_address_1, a.street_address_2, a.city,
       coalesce(a.state, a.country) AS state, a.zip, pc.cage_status
FROM (SELECT DISTINCT p.cage_code FROM w CROSS JOIN pl.pl_v_flis_part p ON p.niin = w.niin WHERE p.cage_code IS NOT NULL) c
CROSS JOIN pl.pl_v_cage_address a ON a.cage_code = c.cage_code
LEFT JOIN pl.pl_p_cage pc ON pc.cage_code = a.cage_code;
UPDATE vendor_cage SET
  vendor_name = (SELECT company_name FROM src_cage x WHERE x.cage_code = vendor_cage.cage_code),
  address1 = (SELECT street_address_1 FROM src_cage x WHERE x.cage_code = vendor_cage.cage_code),
  address2 = (SELECT street_address_2 FROM src_cage x WHERE x.cage_code = vendor_cage.cage_code),
  city = (SELECT city FROM src_cage x WHERE x.cage_code = vendor_cage.cage_code),
  state = (SELECT state FROM src_cage x WHERE x.cage_code = vendor_cage.cage_code),
  postal_code = (SELECT zip FROM src_cage x WHERE x.cage_code = vendor_cage.cage_code),
  company_status_code = (SELECT cage_status FROM src_cage x WHERE x.cage_code = vendor_cage.cage_code),
  source_loaded_at = '$now'
WHERE cage_code IN (SELECT cage_code FROM src_cage);
INSERT INTO vendor_cage (cage_code, vendor_name, address1, address2, city, state, postal_code, company_status_code, source_loaded_at)
SELECT cage_code, company_name, street_address_1, street_address_2, city, state, zip, cage_status, '$now'
FROM src_cage WHERE cage_code NOT IN (SELECT cage_code FROM vendor_cage);
-- CAGEs cross-referenced in FLIS but with no address record still need a parent row.
INSERT OR IGNORE INTO vendor_cage (cage_code, source_loaded_at)
SELECT DISTINCT p.cage_code, '$now' FROM w CROSS JOIN pl.pl_v_flis_part p ON p.niin = w.niin WHERE p.cage_code IS NOT NULL;

-- Child segments: FLIS is the system of record, so for watchlisted NIINs
-- these become this month's snapshot.
DELETE FROM item_part_number_xref WHERE niin IN (SELECT niin FROM w);
INSERT OR IGNORE INTO item_part_number_xref (niin, fsc_code, part_number, cage_code, rncc_code, rnvc_code, rnfc_code,
  rnjc_code, rnsc_code, rnaac_code, sadc_code, dac_code, hcc_code, msds_id, source_loaded_at)
SELECT p.niin, n.fsc, p.part_number, p.cage_code, p.rncc, p.rnvc, p.rnfc, p.rnjc, p.rnsc, p.rnaac, p.sadc, p.dac, p.hcc, p.msds, '$now'
FROM w CROSS JOIN pl.pl_v_flis_part p ON p.niin = w.niin LEFT JOIN pl.pl_p_flis_nsn n ON n.niin = p.niin
WHERE p.part_number IS NOT NULL;

INSERT OR IGNORE INTO service_agency (sa_code)
  SELECT m.moe FROM w CROSS JOIN pl.pl_v_flis_management m ON m.niin = w.niin WHERE m.moe IS NOT NULL
  UNION SELECT f.moe FROM w CROSS JOIN pl.pl_v_flis_phrase f ON f.niin = w.niin WHERE f.moe IS NOT NULL
  UNION SELECT r.moe_cd FROM w CROSS JOIN pl.pl_v_moe_rule r ON r.niin = w.niin WHERE r.moe_cd IS NOT NULL;
DELETE FROM item_management_data WHERE niin IN (SELECT niin FROM w);
INSERT OR IGNORE INTO item_management_data (niin, fsc_code, sa_code, sos_code, mcd, ui_code, uicf, qup, aac_code, slc_code,
  rc_code, ciic_code, usc_code, unit_price, effective_date, source_loaded_at)
SELECT m.niin, n.fsc, m.moe, m.sos, m.mgmt_ctl, m.ui, CAST(m.ui_conv_fac AS REAL), m.qup, m.aac, m.slc, m.rep_rec_code,
  m.ciic, m.usc, CAST(m.unit_price AS REAL), $effDate, '$now'
FROM w CROSS JOIN pl.pl_v_flis_management m ON m.niin = w.niin LEFT JOIN pl.pl_p_flis_nsn n ON n.niin = m.niin;

DELETE FROM item_management_phrase WHERE niin IN (SELECT niin FROM w);
INSERT INTO item_management_phrase (niin, fsc_code, sa_code, phrase_code, phrase_statement)
SELECT f.niin, n.fsc, f.moe, f.phrs_cd, f.phrase_statement
FROM w CROSS JOIN pl.pl_v_flis_phrase f ON f.niin = w.niin LEFT JOIN pl.pl_p_flis_nsn n ON n.niin = f.niin;

DELETE FROM item_moe_data WHERE niin IN (SELECT niin FROM w);
INSERT INTO item_moe_data (niin, fsc_code, moe_rule, sa_code, pica_code, dsor_code, amc_code, amsc_code, nimsc_code,
  imc_code, imca_code, aac_code, collaborators, receivers, moe_eff_date, source_loaded_at)
SELECT r.niin, n.fsc, r.moe_rl, r.moe_cd, r.pica, r.dsor, r.amc, r.amsc, r.nimsc, r.imc, r.imca, r.aac,
  r.auth_collab, r.supp_collab, $moeDate, '$now'
FROM w CROSS JOIN pl.pl_v_moe_rule r ON r.niin = w.niin LEFT JOIN pl.pl_p_flis_nsn n ON n.niin = r.niin;

INSERT OR IGNORE INTO characteristic_code (mrc_code, characteristic_label)
SELECT c.mrc, MIN(c.requirements_statement) FROM w CROSS JOIN pl.pl_v_characteristics c ON c.niin = w.niin WHERE c.mrc IS NOT NULL GROUP BY c.mrc;
DELETE FROM item_characteristic WHERE niin IN (SELECT niin FROM w);
INSERT OR IGNORE INTO item_characteristic (niin, fsc_code, mrc_code, char_value, source_loaded_at)
SELECT c.niin, n.fsc, c.mrc, c.clear_text_reply, '$now'
FROM w CROSS JOIN pl.pl_v_characteristics c ON c.niin = w.niin LEFT JOIN pl.pl_p_flis_nsn n ON n.niin = c.niin
WHERE c.mrc IS NOT NULL;

DELETE FROM item_related_niin WHERE niin IN (SELECT niin FROM w);
INSERT INTO item_related_niin (niin, fsc_code, replaced_niin, isc_code, standardization_originator, standardization_date, status_code)
SELECT s.niin, n.fsc, s.related_nsn, s.isc, s.orig_stdzn_dec, $stdDate, s.niin_stat_cd
FROM w CROSS JOIN pl.pl_v_flis_standardization s ON s.niin = w.niin LEFT JOIN pl.pl_p_flis_nsn n ON n.niin = s.niin
WHERE s.related_nsn IS NOT NULL;

DELETE FROM freight_data WHERE niin IN (SELECT niin FROM w);
INSERT INTO freight_data (niin, fsc_code, originating_activity, integrity_code, freight_description, rail_variance_code,
  nmfc_item_number, nmfc_sub_item_number, uniform_freight_class, hazardous_material_code, less_than_carload_code,
  water_commodity_code, water_commodity_handling_code, air_dimension_code, air_commodity_handling_code, cargo_type_code)
SELECT f.niin, n.fsc, f.acty_cd, f.integ, f.nmf_desc, f.rvc, f.nmfc, f.nmfc_sub, f.ufc, f.hmc, f.ltl, f.wcc, f.shc, f.adc, f.acc, f.tcc
FROM w CROSS JOIN pl.pl_v_freight f ON f.niin = w.niin LEFT JOIN pl.pl_p_flis_nsn n ON n.niin = f.niin;

DELETE FROM packaging_data WHERE niin IN (SELECT niin FROM w);
INSERT INTO packaging_data (niin, fsc_code, data_source_code, pica_sica_indicator, cleaning_drying_procedure_code,
  container_nsn, cushioning_dunnage_material_code, cushioning_dunnage_thickness_code, intermediate_container_code,
  intermediate_container_qty, item_type_storage_code, level_a_pkg_requirement_code, level_b_pkg_requirement_code,
  level_c_pkg_requirement_code, preservation_method_code, preservation_material_code, optional_procedure_indicator,
  packaging_category_code, packaging_design_activity_cage, special_marking_code, special_pkg_instruction_date,
  special_pkg_instruction_number, special_pkg_instruction_revision, unit_container_code, unit_container_level_code,
  unpackaged_item_dimensions, unpackaged_item_weight, unit_pack_cube, unit_pack_size, unit_pack_weight,
  wrapping_material_code, supplemental_instructions)
SELECT p1.niin, n.fsc, p1.pkg_data_source, p1.pica_sica, p1.clng_drying, p3.cont_nsn, p1.cush_dun, p1.thk, p2.inter_cont,
  CAST(p1.icq AS INTEGER), p1.tos, p2.lvl_a, p2.lvl_b, p2.lvl_c, p1.mop, p1.pres_mat, p2.opi, p1.pkg_cat, p3.pkg_design_acty,
  p2.spc_mkg, $spiDate, p3.spi_no, p3.spi_rev, p1.unit_cont, p2.ucl,
  -- FLIS writes 0.00000 for "not recorded": keep those as NULL.
  CASE WHEN CAST(p1.unpkg_item_length AS REAL) > 0 THEN p1.unpkg_item_length || ' x ' || p1.unpkg_item_width || ' x ' || p1.unpkg_item_height END,
  NULLIF(CAST(p1.unpkg_item_weight AS REAL), 0), NULLIF(CAST(p2.unit_pack_cube AS REAL), 0),
  CASE WHEN CAST(p2.up_sz_maxl AS REAL) > 0 THEN p2.up_sz_maxl || ' x ' || p2.up_sz_maxb || ' x ' || p2.up_sz_maxh END,
  NULLIF(CAST(p2.unit_pack_weight AS REAL), 0), p1.wrap_mat, p2.supplemental_instructions
FROM w CROSS JOIN pl.pl_v_flis_packaging_1 p1 ON p1.niin = w.niin
LEFT JOIN pl.pl_v_flis_packaging_2 p2 ON p2.niin = p1.niin AND p2.pica_sica IS p1.pica_sica
LEFT JOIN pl.pl_v_flis_packaging_3 p3 ON p3.niin = p1.niin AND p3.pica_sica IS p1.pica_sica
LEFT JOIN pl.pl_p_flis_nsn n ON n.niin = p1.niin;

DROP TABLE temp.src_item; DROP TABLE temp.src_cage; DROP TABLE temp.w;
COMMIT;
"@
            # One statement at a time (statements end with ';' at end of line),
            # so -Verbose can name any step that gets slow.
            $statements = [regex]::Split($sql, ';[ \t]*\r?\n') | Where-Object { $_.Trim() -and ($_ -replace '(?m)^\s*--.*$', '').Trim() }
            try {
                foreach ($st in $statements) {
                    $sw = [Diagnostics.Stopwatch]::StartNew()
                    Invoke-SqliteExec -Database $db -Sql ($st + ';')
                    if ($sw.Elapsed.TotalSeconds -ge 1) {
                        $first = (($st -replace '(?m)^\s*--.*$', '').Trim() -split "`n")[0]
                        Write-Verbose ("  {0,6:N1}s  {1}" -f $sw.Elapsed.TotalSeconds, $first.Trim())
                    }
                }
            }
            catch { try { Invoke-SqliteExec -Database $db -Sql 'ROLLBACK;' } catch { }; throw }

            $counts = Invoke-SqliteQuery -Database $db -Sql @'
SELECT (SELECT COUNT(*) FROM item_niin) AS items, (SELECT COUNT(*) FROM vendor_cage) AS cages,
       (SELECT COUNT(*) FROM item_part_number_xref) AS xref, (SELECT COUNT(*) FROM item_management_data) AS mgmt,
       (SELECT COUNT(*) FROM item_characteristic) AS chars, (SELECT COUNT(*) FROM ref_code_value) AS codes;
'@
            $c = $counts[0]
            Write-Host "  sustainment.db now: items=$($c.items) cages=$($c.cages) xref=$($c.xref) mgmt=$($c.mgmt) chars=$($c.chars) codes=$($c.codes)"
            $missing = Invoke-SqliteQuery -Database $db -Sql 'SELECT w.niin FROM niin_watchlist w WHERE NOT EXISTS (SELECT 1 FROM pl.pl_p_flis_nsn n WHERE n.niin = w.niin);'
            if ($missing.Count -gt 0) {
                Write-Warning ("  {0} watchlisted NIIN(s) not in this PUB LOG cut (cancelled, restricted, or typo): {1}" -f $missing.Count, (($missing | Select-Object -First 10 | ForEach-Object niin) -join ', '))
            }
        }
        finally { Invoke-SqliteExec -Database $db -Sql 'DETACH DATABASE pl;' }
    }
    finally { Close-SqliteDb -Database $db }
}
