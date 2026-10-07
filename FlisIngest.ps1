#requires -Version 5.1
<#
.SYNOPSIS
  FLIS bulk-download ingest: raw *.txt files from a drop folder -> fb_* raw
  landing tables in sustainment.db, with loud, itemized problem reporting.

.DESCRIPTION
  The FLIS bulk download is the only feeder of sustainment.db. This script
  lands it verbatim (the "raw layer"); normalized tables are built from the
  fb_* tables afterwards.

  Nothing here fails silently:
    * Every file in the drop folder is accounted for: expected and present,
      expected and MISSING, recognized-but-not-landed (BatchDetails, the
      forecast disclaimer), or UNKNOWN.
    * Every header is checked against flis_registry.psd1 (generated from
      docs/flis/FLIS-source_file_headers_schema.md). A missing, extra or
      renamed column QUARANTINES the file -- it is not landed, and the
      previous batch stays in its fb_* table.
    * Every record is parsed by a strict CSV reader that reports, with line
      numbers: wrong field counts (row rejected), broken quoting, encoding
      problems. Columns are always addressed by header name.
    * Key columns are sanity-checked (NIIN = 9 digits, CAGE = 5 characters),
      exact duplicate rows are counted, per-column fill rates are recorded
      and compared with the previous batch, and NIINs added/dropped since the
      previous batch go to fb_change_log.
    * All findings go to the console (colored by severity), to fb_issue_log,
      and to a plain-text report under -ReportFolder. A run with any ERROR
      ends by throwing, after everything has been reported.

  Landing pattern (same as publog.db): fb_<FileStem> with the file's
  verbatim column names, all TEXT, empty -> NULL, plus fb_batch and
  fb_line; loaded into a staging table, row-count-verified, then swapped in.

.EXAMPLE
  . .\FlisIngest.ps1
  Invoke-FlisIngest -DropFolder D:\FLIS\drop -DatabasePath .\sustainment.db
#>

if (-not (Get-Command Open-SqliteDb -ErrorAction SilentlyContinue)) { . (Join-Path $PSScriptRoot 'SQLiteInterop.ps1') }

if (-not ('Flis.CsvFile' -as [type])) {
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Flis
{
    public class CsvProblem
    {
        public int Line;          // physical line where the record starts (header = 1)
        public string Code;       // FieldCount, Quote, Encoding, Empty, ...
        public string Severity;   // ERROR / WARN / INFO
        public string Message;
    }

    // Strict RFC-4180-style reader: comma separated, '"' quoting, '""' escape,
    // quoted fields may contain commas and line breaks. Anything irregular is
    // recorded as a CsvProblem instead of being silently "fixed".
    public class CsvFile
    {
        public string Path;
        public string EncodingName;
        public bool HadBom;
        public string LineEndings;              // CRLF / LF / CR / mixed / none
        public string[] Headers = new string[0];
        public List<string[]> Rows = new List<string[]>();
        public List<int> RowLines = new List<int>();
        public int RecordsRead;                 // data records seen (accepted + rejected)
        public int RejectedRecords;
        public int MultiLineRecords;
        public List<CsvProblem> Problems = new List<CsvProblem>();

        void Add(int line, string code, string sev, string msg)
        {
            Problems.Add(new CsvProblem { Line = line, Code = code, Severity = sev, Message = msg });
        }

        public static CsvFile Read(string path)
        {
            var f = new CsvFile();
            f.Path = path;
            byte[] bytes = File.ReadAllBytes(path);
            string text = f.Decode(bytes);
            f.Parse(text);
            return f;
        }

        string Decode(byte[] b)
        {
            if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) { HadBom = true; EncodingName = "UTF-8 (BOM)"; return new UTF8Encoding(false, true).GetString(b, 3, b.Length - 3); }
            if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) { HadBom = true; EncodingName = "UTF-16LE (BOM)"; return Encoding.Unicode.GetString(b, 2, b.Length - 2); }
            if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) { HadBom = true; EncodingName = "UTF-16BE (BOM)"; return Encoding.BigEndianUnicode.GetString(b, 2, b.Length - 2); }
            bool ascii = true;
            for (int i = 0; i < b.Length; i++) if (b[i] >= 0x80) { ascii = false; break; }
            if (ascii) { EncodingName = "ASCII"; return Encoding.ASCII.GetString(b); }
            try { string s = new UTF8Encoding(false, true).GetString(b); EncodingName = "UTF-8"; return s; }
            catch (DecoderFallbackException)
            {
                EncodingName = "Windows-1252";
                Add(0, "Encoding", "WARN", "File is not valid UTF-8; decoded as Windows-1252. Check any non-ASCII characters (accented names, symbols) in the landed data.");
                return Encoding.GetEncoding(1252).GetString(b);
            }
        }

        void Parse(string t)
        {
            int crlf = 0, lf = 0, cr = 0;
            for (int i = 0; i < t.Length; i++)
            {
                if (t[i] == '\r') { if (i + 1 < t.Length && t[i + 1] == '\n') { crlf++; i++; } else cr++; }
                else if (t[i] == '\n') lf++;
            }
            int kinds = (crlf > 0 ? 1 : 0) + (lf > 0 ? 1 : 0) + (cr > 0 ? 1 : 0);
            LineEndings = kinds == 0 ? "none" : kinds > 1 ? "mixed" : crlf > 0 ? "CRLF" : lf > 0 ? "LF" : "CR";
            if (kinds > 1) Add(0, "LineEndings", "WARN", string.Format("Mixed line endings (CRLF={0}, LF={1}, CR={2}) -- file may have been edited or concatenated.", crlf, lf, cr));

            int pos = 0, line = 1;
            bool first = true;
            while (pos < t.Length)
            {
                int startLine = line;
                bool multi;
                List<string> rec = ReadRecord(t, ref pos, ref line, startLine, out multi);
                if (rec == null) break;
                if (rec.Count == 1 && rec[0].Length == 0)
                {
                    // blank line: harmless at end of file, suspicious in the middle
                    if (pos < t.Length) Add(startLine, "BlankLine", "WARN", "Blank line inside the file (skipped).");
                    continue;
                }
                if (first)
                {
                    Headers = rec.ToArray();
                    first = false;
                    continue;
                }
                RecordsRead++;
                if (multi) MultiLineRecords++;
                if (rec.Count != Headers.Length)
                {
                    RejectedRecords++;
                    string preview = string.Join(",", rec.ToArray());
                    if (preview.Length > 120) preview = preview.Substring(0, 120) + "...";
                    Add(startLine, "FieldCount", "ERROR", string.Format("Row has {0} fields, header has {1} -- row REJECTED (not landed). Starts: {2}", rec.Count, Headers.Length, preview));
                    continue;
                }
                Rows.Add(rec.ToArray());
                RowLines.Add(startLine);
            }
            if (first) Add(0, "Empty", "ERROR", "File is empty -- no header row.");
            if (MultiLineRecords > 0) Add(0, "MultiLine", "INFO", MultiLineRecords + " record(s) contain line breaks inside quoted values (kept as-is).");
        }

        List<string> ReadRecord(string t, ref int pos, ref int line, int startLine, out bool multi)
        {
            multi = false;
            var fields = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false, wasQuoted = false, afterClose = false;
            if (pos >= t.Length) return null;
            while (pos < t.Length)
            {
                char c = t[pos];
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (pos + 1 < t.Length && t[pos + 1] == '"') { sb.Append('"'); pos += 2; continue; }
                        inQuotes = false; afterClose = true; pos++; continue;
                    }
                    if (c == '\r' || c == '\n')
                    {
                        multi = true;
                        if (c == '\r' && pos + 1 < t.Length && t[pos + 1] == '\n') { sb.Append("\r\n"); pos += 2; } else { sb.Append(c); pos++; }
                        line++;
                        continue;
                    }
                    sb.Append(c); pos++; continue;
                }
                if (c == ',') { fields.Add(sb.ToString()); sb.Length = 0; wasQuoted = false; afterClose = false; pos++; continue; }
                if (c == '\r' || c == '\n')
                {
                    if (c == '\r' && pos + 1 < t.Length && t[pos + 1] == '\n') pos += 2; else pos++;
                    line++;
                    fields.Add(sb.ToString());
                    return fields;
                }
                if (c == '"')
                {
                    if (sb.Length == 0 && !wasQuoted) { inQuotes = true; wasQuoted = true; pos++; continue; }
                    Add(startLine, "Quote", "WARN", "Stray double quote inside an unquoted value (field " + (fields.Count + 1) + "); kept literally.");
                    sb.Append(c); pos++; continue;
                }
                if (afterClose)
                {
                    Add(startLine, "Quote", "WARN", "Text after a closing quote in field " + (fields.Count + 1) + "; kept literally.");
                    afterClose = false;
                }
                sb.Append(c); pos++;
            }
            if (inQuotes) Add(startLine, "Quote", "ERROR", "Unterminated quoted value at end of file -- the last record is probably truncated.");
            fields.Add(sb.ToString());
            return fields;
        }
    }

    // Generic bulk insert of string rows through winsqlite3.dll. Empty -> NULL.
    public static class Sqlite
    {
        const string DLL = "winsqlite3.dll";
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int nBytes, out IntPtr stmt, IntPtr tail);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_bind_text(IntPtr stmt, int index, byte[] value, int nBytes, IntPtr destructor);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_bind_null(IntPtr stmt, int index);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_bind_int64(IntPtr stmt, int index, long value);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_step(IntPtr stmt);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_reset(IntPtr stmt);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_finalize(IntPtr stmt);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr sqlite3_errmsg(IntPtr db);
        static readonly IntPtr TRANSIENT = new IntPtr(-1);

        static byte[] Z(string s) { var b = Encoding.UTF8.GetBytes(s); var z = new byte[b.Length + 1]; Buffer.BlockCopy(b, 0, z, 0, b.Length); return z; }
        static string Err(IntPtr db)
        {
            IntPtr p = sqlite3_errmsg(db); if (p == IntPtr.Zero) return "?";
            int n = 0; while (Marshal.ReadByte(p, n) != 0) n++;
            var b = new byte[n]; Marshal.Copy(p, b, 0, n); return Encoding.UTF8.GetString(b);
        }

        // INSERT INTO table VALUES (row..., batch, line) for every row.
        // Caller owns the transaction.
        public static int InsertRows(IntPtr db, string table, int columnCount, List<string[]> rows, List<int> lines, long batch)
        {
            var ph = new StringBuilder();
            for (int c = 0; c < columnCount + 2; c++) ph.Append(c == 0 ? "?" : ",?");
            IntPtr stmt;
            if (sqlite3_prepare_v2(db, Z("INSERT INTO \"" + table.Replace("\"", "\"\"") + "\" VALUES (" + ph + ")"), -1, out stmt, IntPtr.Zero) != 0)
                throw new Exception("prepare failed: " + Err(db));
            int n = 0;
            try
            {
                for (int r = 0; r < rows.Count; r++)
                {
                    string[] row = rows[r];
                    for (int c = 0; c < columnCount; c++)
                    {
                        string v = c < row.Length ? row[c] : null;
                        if (string.IsNullOrEmpty(v)) sqlite3_bind_null(stmt, c + 1);
                        else { byte[] b = Encoding.UTF8.GetBytes(v); sqlite3_bind_text(stmt, c + 1, b, b.Length, TRANSIENT); }
                    }
                    sqlite3_bind_int64(stmt, columnCount + 1, batch);
                    sqlite3_bind_int64(stmt, columnCount + 2, lines[r]);
                    if (sqlite3_step(stmt) != 101) throw new Exception("insert failed at source line " + lines[r] + ": " + Err(db));
                    sqlite3_reset(stmt);
                    n++;
                }
            }
            finally { sqlite3_finalize(stmt); }
            return n;
        }
    }
}
"@
}

# ---------------------------------------------------------------------------
# Issue collection: console + fb_issue_log + text report
# ---------------------------------------------------------------------------

function New-FlisRun {
    param([string]$ReportFolder, [string]$ConsoleDetail = 'Summary', [int]$MaxConsoleGroups = 40)
    $run = [PSCustomObject]@{
        ConsoleDetail = $ConsoleDetail
        MaxConsoleGroups = $MaxConsoleGroups
        RunId      = (Get-Date).ToString('yyyyMMdd-HHmmss')
        Started    = Get-Date
        Batch      = $null
        Issues     = New-Object System.Collections.Generic.List[object]
        Files      = New-Object System.Collections.Generic.List[object]
        Report     = New-Object System.Collections.Generic.List[string]
        ReportPath = $null
    }
    if ($ReportFolder) {
        New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null
        $run.ReportPath = Join-Path $ReportFolder ("flis_ingest_{0}.txt" -f $run.RunId)
    }
    $run
}

function Add-FlisIssue {
    # Severity: ERROR (something was not landed / cannot be trusted),
    # WARN (landed, but look at it), INFO (context).
    # Every issue is recorded (fb_issue_log + report). The console shows them
    # grouped (see Write-FlisDigest) unless -ConsoleDetail All, so one bad
    # pattern repeated over thousands of rows can't bury the useful output.
    param($Run, [ValidateSet('ERROR','WARN','INFO')][string]$Severity, [string]$File, [string]$Code, [string]$Message, [int]$Line = 0, [string]$Column)
    $Run.Issues.Add([PSCustomObject]@{ Severity = $Severity; File = $File; Code = $Code; Line = $Line; Column = $Column; Message = $Message })
    if ($Run.ConsoleDetail -eq 'All') {
        $where = if ($Line -gt 0) { "${File}:$Line" } elseif ($File) { $File } else { '(batch)' }
        $text = "  [{0,-5}] {1}: {2}" -f $Severity, $where, $Message
        switch ($Severity) { 'ERROR' { Write-Host $text -ForegroundColor Red } 'WARN' { Write-Host $text -ForegroundColor Yellow } default { Write-Host $text -ForegroundColor Gray } }
    }
}

# One line per (file, code): severity, count, first example. Used for the
# console digest and as the index at the top of the report.
function Get-FlisIssueGroups {
    param($Issues, [string[]]$Severities = @('ERROR','WARN','INFO'))
    $rank = @{ ERROR = 0; WARN = 1; INFO = 2 }
    $Issues | Where-Object { $Severities -contains $_.Severity } |
        Group-Object { "{0}|{1}|{2}" -f $_.Severity, $_.File, $_.Code } |
        ForEach-Object {
            $f = $_.Group[0]
            [PSCustomObject]@{ Severity = $f.Severity; File = $f.File; Code = $f.Code; Count = $_.Count; First = $f; Rank = $rank[$f.Severity] }
        } | Sort-Object Rank, File, Code
}

function Format-FlisGroup {
    param($G)
    $where = if ($G.File) { $G.File } else { '(batch)' }
    $at = if ($G.First.Line -gt 0) { " (first at line $($G.First.Line))" } else { '' }
    $msg = $G.First.Message
    if ($msg.Length -gt 160) { $msg = $msg.Substring(0, 157) + '...' }
    $cnt = if ($G.Count -gt 1) { (" x{0:N0}" -f $G.Count) } else { '' }
    "  [{0,-5}] {1} {2}{3}: {4}{5}" -f $G.Severity, $where, $G.Code, $cnt, $msg, $at
}

function Write-FlisDigest {
    # Prints grouped WARN/ERROR lines (capped), INFO only in Files/All mode.
    param($Run, $Issues, [switch]$IncludeInfo, [int]$Max = 0)
    $sev = if ($IncludeInfo) { 'ERROR','WARN','INFO' } else { 'ERROR','WARN' }
    $groups = @(Get-FlisIssueGroups -Issues $Issues -Severities $sev)
    $shown = 0
    foreach ($g in $groups) {
        if ($Max -gt 0 -and $shown -ge $Max) { break }
        $color = switch ($g.Severity) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } default { 'Gray' } }
        Write-Host (Format-FlisGroup $g) -ForegroundColor $color
        $shown++
    }
    if ($Max -gt 0 -and $groups.Count -gt $Max) {
        Write-Host ("  ... {0} more issue group(s) not shown -- every one is in the report and fb_issue_log." -f ($groups.Count - $Max)) -ForegroundColor Yellow
    }
}

function Write-FlisReportLine {
    # Section text for the report; echoed to the console in Files/All mode,
    # or always with -Always (headers, final summary).
    param($Run, [string]$Text, [string]$Color = 'White', [switch]$Always)
    $Run.Report.Add($Text)
    if ($Always -or $Run.ConsoleDetail -ne 'Summary') { Write-Host $Text -ForegroundColor $Color }
}

# Caps per-row findings so a systematically bad file reports clearly
# ("1,204 rows have a bad NIIN; first 10: ...") instead of 1,204 lines.
function Add-FlisIssueSample {
    param($Run, [string]$Severity, [string]$File, [string]$Code, [string]$What, [System.Collections.Generic.List[object]]$Hits, [int]$Total, [string]$Column, [int]$Show = 10)
    if ($Hits.Count -eq 0) { return }
    $sample = ($Hits | Select-Object -First $Show | ForEach-Object { "line $($_.Line) '$($_.Value)'" }) -join '; '
    $more = if ($Hits.Count -gt $Show) { " (+$($Hits.Count - $Show) more)" } else { '' }
    Add-FlisIssue $Run $Severity $File $Code ("{0:N0} of {1:N0} rows {2}. {3}{4}" -f $Hits.Count, $Total, $What, $sample, $more) -Column $Column
}

# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

function Get-FlisRegistry {
    [CmdletBinding()]
    param([string]$Path = (Join-Path $PSScriptRoot 'flis_registry.psd1'))
    if (-not (Test-Path -LiteralPath $Path)) { throw "FLIS registry not found: $Path" }
    $reg = Import-PowerShellDataFile -LiteralPath $Path
    foreach ($stem in $reg.Files.Keys) {
        $f = $reg.Files[$stem]
        foreach ($k in 'Kind','FileName','Headers') { if (-not $f.ContainsKey($k)) { throw "Registry entry '$stem' is missing '$k'." } }
        if ($f.Kind -notin 'Exact','RollingStock','RollingForecast','BatchDetails') { throw "Registry entry '$stem' has unknown Kind '$($f.Kind)'." }
    }
    $reg
}

# ---------------------------------------------------------------------------
# Header validation
# ---------------------------------------------------------------------------

$Script:FlisMonthNum = @{
    January = 1; February = 2; March = 3; April = 4; May = 5; June = 6; July = 7; August = 8; September = 9; October = 10; November = 11; December = 12
    Jan = 1; Feb = 2; Mar = 3; Apr = 4; Jun = 6; Jul = 7; Aug = 8; Sep = 9; Oct = 10; Nov = 11; Dec = 12
}

function Test-FlisHeader {
    # Returns @{ Ok = bool; Months = @(...) } and records issues. Comparison
    # is exact (case-sensitive); a case-only difference is reported as such.
    param($Run, [string]$Stem, $Entry, [string[]]$Actual)
    $ok = $true
    $months = @()
    $dups = $Actual | Group-Object | Where-Object Count -gt 1
    foreach ($d in $dups) { Add-FlisIssue $Run ERROR $Stem 'HeaderDuplicate' "Column '$($d.Name)' appears $($d.Count) times in the header."; $ok = $false }
    $blank = @($Actual | Where-Object { -not $_ -or -not $_.Trim() })
    if ($blank.Count) { Add-FlisIssue $Run ERROR $Stem 'HeaderBlank' "$($blank.Count) blank column name(s) in the header."; $ok = $false }

    if ($Entry.Kind -eq 'RollingStock' -or $Entry.Kind -eq 'RollingForecast') {
        $lead = @($Entry.LeadHeaders); $trail = @($Entry.TrailHeaders)
        $expectedFixed = $lead + $trail
        $monthCols = New-Object System.Collections.Generic.List[object]
        $other = New-Object System.Collections.Generic.List[string]
        foreach ($h in $Actual) {
            if ($expectedFixed -ccontains $h) { continue }
            $m = [regex]::Match($h, $Entry.MonthPattern)
            if ($m.Success) {
                $yr = [int]$m.Groups[2].Value; if ($yr -lt 100) { $yr += 2000 }
                $monthCols.Add([PSCustomObject]@{ Column = $h; Year = $yr; Month = $Script:FlisMonthNum[$m.Groups[1].Value] })
            } else { $other.Add($h) }
        }
        foreach ($h in $expectedFixed) {
            if ($Actual -cnotcontains $h) {
                $ci = $Actual | Where-Object { $_ -ieq $h } | Select-Object -First 1
                if ($ci) { Add-FlisIssue $Run ERROR $Stem 'HeaderCase' "Expected column '$h' is present only as '$ci' (case differs)." -Column $h }
                else { Add-FlisIssue $Run ERROR $Stem 'HeaderMissing' "Expected column '$h' is missing." -Column $h }
                $ok = $false
            }
        }
        foreach ($h in $other) { Add-FlisIssue $Run ERROR $Stem 'HeaderUnexpected' "Column '$h' is neither a known fixed column nor a month column matching $($Entry.MonthPattern)." -Column $h; $ok = $false }
        if ($monthCols.Count -eq 0) { Add-FlisIssue $Run ERROR $Stem 'HeaderNoMonths' "No month columns found."; $ok = $false }
        else {
            $sorted = @($monthCols | Sort-Object Year, Month)
            $first = $sorted[0]; $last = $sorted[-1]
            $span = ($last.Year - $first.Year) * 12 + ($last.Month - $first.Month) + 1
            Add-FlisIssue $Run INFO $Stem 'Months' ("{0} month columns, {1} through {2}." -f $monthCols.Count, $first.Column, $last.Column)
            if ($span -ne $monthCols.Count) {
                Add-FlisIssue $Run WARN $Stem 'MonthGap' ("Month columns are not contiguous: {0} columns span {1} months ({2}..{3})." -f $monthCols.Count, $span, $first.Column, $last.Column)
            }
            $dupMonths = $monthCols | Group-Object { "$($_.Year)-$($_.Month)" } | Where-Object Count -gt 1
            foreach ($d in $dupMonths) { Add-FlisIssue $Run ERROR $Stem 'MonthDuplicate' "Month $($d.Name) appears in $($d.Count) columns."; $ok = $false }
            $months = $sorted
        }
        return @{ Ok = $ok; Months = $months }
    }

    $expected = @($Entry.Headers)
    $missing = @($expected | Where-Object { $Actual -cnotcontains $_ })
    $extra = @($Actual | Where-Object { $expected -cnotcontains $_ })
    foreach ($m in $missing) {
        $ci = $extra | Where-Object { $_ -ieq $m } | Select-Object -First 1
        if ($ci) { Add-FlisIssue $Run ERROR $Stem 'HeaderCase' "Expected column '$m' is present only as '$ci' (case differs)." -Column $m }
        else { Add-FlisIssue $Run ERROR $Stem 'HeaderMissing' "Expected column '$m' is missing (renamed or dropped?)." -Column $m }
    }
    foreach ($x in $extra) {
        if ($missing -icontains $x) { continue }   # already reported as a case difference
        Add-FlisIssue $Run ERROR $Stem 'HeaderUnexpected' "Unexpected column '$x' (not in the source-of-truth schema)." -Column $x
    }
    if ($missing.Count -or $extra.Count -or $dups) { $ok = $false }
    elseif (($Actual -join '|') -cne ($expected -join '|')) {
        Add-FlisIssue $Run WARN $Stem 'HeaderOrder' "All expected columns are present but in a different order. Landing by column name (safe), but the source format changed -- worth reporting."
    }
    @{ Ok = $ok; Months = @() }
}

# ---------------------------------------------------------------------------
# Database objects
# ---------------------------------------------------------------------------

function Initialize-FlisIngestDb {
    param([IntPtr]$Database)
    Invoke-SqliteExec -Database $Database -Sql @'
CREATE TABLE IF NOT EXISTS fb_load_log (
    run_id        TEXT NOT NULL,
    batch_number  INTEGER,
    file_stem     TEXT NOT NULL,
    file_name     TEXT,
    status        TEXT NOT NULL,      -- LANDED / QUARANTINED / MISSING / RECOGNIZED / UNKNOWN
    file_size     INTEGER,
    file_sha256   TEXT,
    encoding      TEXT,
    line_endings  TEXT,
    columns       TEXT,
    rows_read     INTEGER,
    rows_rejected INTEGER,
    rows_landed   INTEGER,
    duplicate_rows INTEGER,
    errors        INTEGER,
    warnings      INTEGER,
    loaded_at     TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS fb_issue_log (
    run_id       TEXT NOT NULL,
    batch_number INTEGER,
    file_stem    TEXT,
    severity     TEXT NOT NULL,
    code         TEXT NOT NULL,
    line         INTEGER,
    column_name  TEXT,
    message      TEXT NOT NULL,
    logged_at    TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS fb_fill_rate (
    run_id       TEXT NOT NULL,
    batch_number INTEGER,
    file_stem    TEXT NOT NULL,
    column_name  TEXT NOT NULL,
    non_empty    INTEGER NOT NULL,
    total        INTEGER NOT NULL,
    pct          REAL NOT NULL
);
CREATE TABLE IF NOT EXISTS fb_change_log (
    batch_number      INTEGER NOT NULL,
    previous_batch    INTEGER,
    file_stem         TEXT NOT NULL,
    niin              TEXT NOT NULL,
    change            TEXT NOT NULL     -- added / dropped (vs the previous landed batch of this file)
);
CREATE TABLE IF NOT EXISTS fb_batch (
    batch_number   INTEGER PRIMARY KEY,
    run_id         TEXT NOT NULL,
    batch_file     TEXT,
    batch_ids_in_file TEXT,             -- distinct BatchId values inside the file (may differ from the name)
    pulled_at      TEXT,                -- newest LastWriteTime of the data files
    loaded_at      TEXT NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS ix_fb_issue_log_run ON fb_issue_log (run_id);
CREATE INDEX IF NOT EXISTS ix_fb_change_log_niin ON fb_change_log (niin);
'@
}

function ConvertTo-SqlIdent { param([string]$Name) '"' + $Name.Replace('"', '""') + '"' }

# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

function Invoke-FlisIngest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DropFolder,
        [Parameter(Mandatory)][string]$DatabasePath,
        [string]$RegistryPath = (Join-Path $PSScriptRoot 'flis_registry.psd1'),
        [string]$ReportFolder = (Join-Path $PSScriptRoot 'logs'),
        # Use only when the drop has no Batch{N}Details.txt; recorded as a warning.
        [long]$BatchNumber = 0,
        # Treat a missing expected file as an ERROR instead of a WARN.
        [switch]$RequireAllFiles,
        # Thresholds for batch-over-batch comparisons.
        [double]$RowDropWarnPct = 20,
        [double]$FillDropWarnPoints = 25,
        # Console verbosity. Summary (default): header, one line per file with
        # problems, grouped issues (capped at -MaxConsoleGroups), final summary.
        # Files: also every file's status and its grouped issues incl. INFO.
        # All: every individual issue as it happens (can be very long).
        # Full detail always goes to fb_issue_log and the report file.
        [ValidateSet('Summary','Files','All')][string]$ConsoleDetail = 'Summary',
        [int]$MaxConsoleGroups = 40,
        # Report problems but do not throw at the end (for scripted callers).
        [switch]$NoThrow
    )
    $run = New-FlisRun -ReportFolder $ReportFolder -ConsoleDetail $ConsoleDetail -MaxConsoleGroups $MaxConsoleGroups
    $reg = Get-FlisRegistry -Path $RegistryPath
    $drop = (Resolve-Path -LiteralPath $DropFolder).ProviderPath
    $dbPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DatabasePath)

    Write-FlisReportLine $run ("FLIS ingest run {0}" -f $run.RunId) Cyan -Always
    Write-FlisReportLine $run ("  drop folder : {0}" -f $drop) -Always
    Write-FlisReportLine $run ("  database    : {0}" -f $dbPath) -Always
    Write-FlisReportLine $run ("  registry    : {0} ({1} files)" -f $RegistryPath, $reg.Files.Count) -Always

    # ---- 1. account for every file -----------------------------------------
    Write-FlisReportLine $run "== 1. File inventory ==" Cyan
    $present = @(Get-ChildItem -LiteralPath $drop -File)
    $byName = @{}
    foreach ($f in $present) { $byName[$f.Name.ToLowerInvariant()] = $f }
    $claimed = @{}
    $plan = New-Object System.Collections.Generic.List[object]

    $batchFiles = @($present | Where-Object { $_.Name -match '^Batch(\d+)Details\.txt$' })
    foreach ($bf in $batchFiles) { $claimed[$bf.Name.ToLowerInvariant()] = $true }
    foreach ($ig in $reg.IgnoredFiles) {
        $f = $byName[$ig.ToLowerInvariant()]
        if ($f) { $claimed[$f.Name.ToLowerInvariant()] = $true; Add-FlisIssue $run INFO $f.Name 'Ignored' 'Known non-CSV file (prose); recognized and skipped.' }
    }
    foreach ($stem in ($reg.Files.Keys | Sort-Object)) {
        $e = $reg.Files[$stem]
        if ($e.Kind -eq 'BatchDetails') { continue }
        $f = $byName[$e.FileName.ToLowerInvariant()]
        if ($f) {
            $claimed[$f.Name.ToLowerInvariant()] = $true
            if ($f.Name -cne $e.FileName) { Add-FlisIssue $run WARN $stem 'FileNameCase' "File is named '$($f.Name)'; the source of truth says '$($e.FileName)' (case differs). Accepted." }
            $plan.Add([PSCustomObject]@{ Stem = $stem; Entry = $e; File = $f })
        }
        else {
            $sev = if ($RequireAllFiles) { 'ERROR' } else { 'WARN' }
            Add-FlisIssue $run $sev $stem 'FileMissing' "Expected file '$($e.FileName)' is NOT in the drop folder. Its fb_ table keeps the previous batch (if any)."
            $run.Files.Add([PSCustomObject]@{ Stem = $stem; Status = 'MISSING'; File = $e.FileName })
        }
    }
    foreach ($f in $present) {
        if (-not $claimed[$f.Name.ToLowerInvariant()]) {
            Add-FlisIssue $run WARN $f.Name 'FileUnknown' "Unrecognized file in the drop folder -- not landed. If this is a new FLIS file, add it to the source-of-truth schema and the registry."
            $run.Files.Add([PSCustomObject]@{ Stem = $f.Name; Status = 'UNKNOWN'; File = $f.Name })
        }
    }
    Write-FlisReportLine $run ("  {0} files present, {1} expected data files found, {2} missing." -f $present.Count, $plan.Count, @($run.Files | Where-Object Status -eq 'MISSING').Count)

    # ---- 2. batch number ---------------------------------------------------
    Write-FlisReportLine $run "== 2. Batch ==" Cyan
    $batchIds = $null; $batchFileName = $null
    if ($batchFiles.Count -gt 1) {
        Add-FlisIssue $run ERROR '(batch)' 'BatchMultiple' ("{0} batch files in one drop ({1}) -- files from different downloads are mixed. Nothing will be landed." -f $batchFiles.Count, (($batchFiles | ForEach-Object Name) -join ', '))
    }
    elseif ($batchFiles.Count -eq 1) {
        $batchFileName = $batchFiles[0].Name
        $fromName = [long]([regex]::Match($batchFileName, '^Batch(\d+)Details\.txt$').Groups[1].Value)
        if ($BatchNumber -and $BatchNumber -ne $fromName) {
            Add-FlisIssue $run WARN '(batch)' 'BatchOverride' "-BatchNumber $BatchNumber overrides the file name's $fromName."
        } else { $BatchNumber = $fromName }
        try {
            $bcsv = [Flis.CsvFile]::Read($batchFiles[0].FullName)
            $ix = [array]::IndexOf($bcsv.Headers, 'BatchId')
            if ($ix -ge 0) {
                $batchIds = (@($bcsv.Rows | ForEach-Object { $_[$ix] } | Where-Object { $_ } | Sort-Object -Unique) -join ',')
                Add-FlisIssue $run INFO $batchFileName 'BatchId' "BatchId value(s) inside the file: $(if ($batchIds) { $batchIds } else { '(none)' })."
                if ($batchIds -and $batchIds -ne "$BatchNumber") { Add-FlisIssue $run INFO $batchFileName 'BatchIdDiffers' "In-file BatchId differs from the file-name batch number $BatchNumber (expected per the owner; logged only)." }
            } else { Add-FlisIssue $run INFO $batchFileName 'BatchId' 'No BatchId column in the batch file (not needed; logged only).' }
        } catch { Add-FlisIssue $run WARN $batchFileName 'BatchRead' "Could not read the batch file ($($_.Exception.Message)); batch number taken from its name." }
        $run.Files.Add([PSCustomObject]@{ Stem = 'BatchDetails'; Status = 'RECOGNIZED'; File = $batchFileName })
    }
    elseif ($BatchNumber) {
        Add-FlisIssue $run WARN '(batch)' 'BatchManual' "No Batch{N}Details.txt in the drop; using -BatchNumber $BatchNumber as given."
    }
    else {
        Add-FlisIssue $run ERROR '(batch)' 'BatchMissing' 'No Batch{N}Details.txt in the drop and no -BatchNumber given -- cannot stamp rows with a batch. Nothing will be landed.'
    }
    $run.Batch = $BatchNumber
    $canLand = -not ($run.Issues | Where-Object { $_.Severity -eq 'ERROR' -and $_.Code -like 'Batch*' })
    if ($BatchNumber) { Write-FlisReportLine $run ("  batch number: {0}" -f $BatchNumber) }

    $db = Open-SqliteDb -Path $dbPath
    try {
        Initialize-FlisIngestDb -Database $db
        $fk = (Invoke-SqliteQuery -Database $db -Sql 'PRAGMA foreign_keys;')[0].foreign_keys
        $ver = [NativeSqlite]::Utf8PtrToString([NativeSqlite]::sqlite3_libversion())
        Add-FlisIssue $run INFO '(database)' 'Sqlite' "winsqlite3 $ver; PRAGMA foreign_keys = $fk on this connection."
        if ($canLand) {
            $prior = Invoke-SqliteQuery -Database $db -Sql 'SELECT MAX(batch_number) AS b FROM fb_batch;'
            if ($prior.Count -and $prior[0].b -and [long]$prior[0].b -gt $BatchNumber) {
                Add-FlisIssue $run WARN '(batch)' 'BatchOlder' "Batch $BatchNumber is older than the newest batch already loaded ($($prior[0].b)). Landing it anyway -- check this is intended."
            }
            $pulled = ($plan | ForEach-Object { $_.File.LastWriteTime } | Sort-Object | Select-Object -Last 1)
            Invoke-SqliteNonQuery -Database $db -Sql 'INSERT OR REPLACE INTO fb_batch (batch_number, run_id, batch_file, batch_ids_in_file, pulled_at) VALUES (?,?,?,?,?);' `
                -Params @($BatchNumber, $run.RunId, $batchFileName, $batchIds, $(if ($pulled) { $pulled.ToString('yyyy-MM-dd HH:mm:ss') } else { $null })) | Out-Null
        }

        # ---- 3. each data file ----------------------------------------------
        Write-FlisReportLine $run "== 3. Files ==" Cyan
        foreach ($p in $plan) {
            $stem = $p.Stem; $e = $p.Entry; $file = $p.File
            $errBefore = @($run.Issues | Where-Object { $_.File -eq $stem -and $_.Severity -eq 'ERROR' }).Count
            $fileNo = $plan.IndexOf($p) + 1
            Write-Progress -Activity "FLIS ingest, batch $BatchNumber" -Status ("{0} ({1} of {2})" -f $file.Name, $fileNo, $plan.Count) -PercentComplete (100 * ($fileNo - 1) / [math]::Max(1, $plan.Count))
            $issuesBefore = $run.Issues.Count
            Write-FlisReportLine $run ("-- {0} ({1:N0} bytes)" -f $file.Name, $file.Length) White
            $sha = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            $csv = $null
            try { $csv = [Flis.CsvFile]::Read($file.FullName) }
            catch { Add-FlisIssue $run ERROR $stem 'ReadFailed' "Could not read the file: $($_.Exception.Message)" }
            $status = 'QUARANTINED'; $landed = 0; $dupCount = 0
            if ($csv) {
                # Every parser finding is recorded individually (fb_issue_log has
                # them all); the console groups them per file and type.
                foreach ($pr in $csv.Problems) { Add-FlisIssue $run $pr.Severity $stem $pr.Code $pr.Message -Line $pr.Line }
                Add-FlisIssue $run INFO $stem 'Read' ("{0:N0} data rows read, {1:N0} rejected; encoding {2}; line endings {3}." -f $csv.RecordsRead, $csv.RejectedRecords, $csv.EncodingName, $csv.LineEndings)
                $h = Test-FlisHeader -Run $run -Stem $stem -Entry $e -Actual $csv.Headers
                if ($h.Ok -and ($csv.Problems | Where-Object { $_.Code -eq 'Quote' -and $_.Severity -eq 'ERROR' })) {
                    Add-FlisIssue $run ERROR $stem 'Quarantined' 'File ends inside a quoted value -- probably truncated during download or copy. File QUARANTINED (not landed); re-copy it from the original download.'
                    $h.Ok = $false
                }
                elseif (-not $h.Ok) {
                    Add-FlisIssue $run ERROR $stem 'Quarantined' 'Header does not match the source-of-truth schema -- file QUARANTINED (not landed). Report this to the data owner.'
                }
                elseif ($csv.Rows.Count -eq 0) {
                    Add-FlisIssue $run WARN $stem 'NoRows' 'Header is valid but the file has no data rows. Landing an empty table for this batch.'
                }
                if ($h.Ok -and $canLand) {
                    $cols = $csv.Headers
                    $colIx = @{}; for ($i = 0; $i -lt $cols.Count; $i++) { $colIx[$cols[$i]] = $i }

                    # Key-column checks (by name).
                    if ($e.NiinColumn) {
                        $ix = $colIx[$e.NiinColumn]
                        $bad = New-Object System.Collections.Generic.List[object]; $blank = New-Object System.Collections.Generic.List[object]; $ws = New-Object System.Collections.Generic.List[object]
                        for ($r = 0; $r -lt $csv.Rows.Count; $r++) {
                            $v = $csv.Rows[$r][$ix]
                            if (-not $v) { $blank.Add([PSCustomObject]@{ Line = $csv.RowLines[$r]; Value = '' }) }
                            elseif ($v -cnotmatch '^\d{9}$') {
                                if ($v.Trim() -cmatch '^\d{9}$') { $ws.Add([PSCustomObject]@{ Line = $csv.RowLines[$r]; Value = $v }) }
                                else { $bad.Add([PSCustomObject]@{ Line = $csv.RowLines[$r]; Value = $v }) }
                            }
                        }
                        Add-FlisIssueSample $run WARN $stem 'NiinBlank' "have a blank $($e.NiinColumn)" $blank $csv.Rows.Count $e.NiinColumn
                        Add-FlisIssueSample $run WARN $stem 'NiinSpaces' "have spaces around $($e.NiinColumn) (landed verbatim; normalization will trim)" $ws $csv.Rows.Count $e.NiinColumn
                        Add-FlisIssueSample $run WARN $stem 'NiinFormat' "have a $($e.NiinColumn) that is not 9 digits (leading zeros lost in Excel?)" $bad $csv.Rows.Count $e.NiinColumn
                    }
                    foreach ($rc in @($e.NiinRefColumns)) {
                        if (-not $rc) { continue }
                        $ix = $colIx[$rc]; $bad = New-Object System.Collections.Generic.List[object]
                        for ($r = 0; $r -lt $csv.Rows.Count; $r++) { $v = $csv.Rows[$r][$ix]; if ($v -and $v.Trim() -cnotmatch '^\d{9}$') { $bad.Add([PSCustomObject]@{ Line = $csv.RowLines[$r]; Value = $v }) } }
                        Add-FlisIssueSample $run WARN $stem 'NiinRefFormat' "have a $rc that is not 9 digits" $bad $csv.Rows.Count $rc
                    }
                    if ($e.CageColumn) {
                        $ix = $colIx[$e.CageColumn]; $bad = New-Object System.Collections.Generic.List[object]
                        for ($r = 0; $r -lt $csv.Rows.Count; $r++) { $v = $csv.Rows[$r][$ix]; if ($v -and $v.Trim() -cnotmatch '^[0-9A-Z]{5}$') { $bad.Add([PSCustomObject]@{ Line = $csv.RowLines[$r]; Value = $v }) } }
                        Add-FlisIssueSample $run WARN $stem 'CageFormat' "have a $($e.CageColumn) that is not 5 upper-case letters/digits" $bad $csv.Rows.Count $e.CageColumn
                    }

                    # Exact duplicate rows (landed verbatim, but counted).
                    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
                    $dupHits = New-Object System.Collections.Generic.List[object]
                    for ($r = 0; $r -lt $csv.Rows.Count; $r++) {
                        if (-not $seen.Add(($csv.Rows[$r] -join [char]31))) { $dupHits.Add([PSCustomObject]@{ Line = $csv.RowLines[$r]; Value = 'duplicate of an earlier row' }) }
                    }
                    $dupCount = $dupHits.Count
                    Add-FlisIssueSample $run WARN $stem 'DuplicateRows' 'are exact duplicates of an earlier row (landed as-is; dedupe happens in normalization)' $dupHits $csv.Rows.Count '' 5

                    # Previous batch of this file, for comparisons.
                    $prev = Invoke-SqliteQuery -Database $db -Sql "SELECT batch_number, rows_landed FROM fb_load_log WHERE file_stem = ? AND status = 'LANDED' AND batch_number <> ? ORDER BY loaded_at DESC LIMIT 1;" -Params @($stem, $BatchNumber)

                    # Land: staging -> verify -> swap.
                    $raw = "fb_$stem"; $stage = "${raw}__staging"
                    $colDefs = (($cols | ForEach-Object { (ConvertTo-SqlIdent $_) + ' TEXT' }) + 'fb_batch INTEGER NOT NULL', 'fb_line INTEGER NOT NULL') -join ', '
                    Invoke-SqliteExec -Database $db -Sql "DROP TABLE IF EXISTS $(ConvertTo-SqlIdent $stage); CREATE TABLE $(ConvertTo-SqlIdent $stage) ($colDefs);"
                    Invoke-SqliteExec -Database $db -Sql 'BEGIN;'
                    try {
                        $landed = [Flis.Sqlite]::InsertRows($db, $stage, $cols.Count, $csv.Rows, $csv.RowLines, $BatchNumber)
                        Invoke-SqliteExec -Database $db -Sql 'COMMIT;'
                    } catch { Invoke-SqliteExec -Database $db -Sql 'ROLLBACK;'; throw }
                    $check = (Invoke-SqliteQuery -Database $db -Sql "SELECT COUNT(*) AS n FROM $(ConvertTo-SqlIdent $stage);")[0].n
                    if ($check -ne $csv.Rows.Count) {
                        Add-FlisIssue $run ERROR $stem 'LandCount' "Staging holds $check rows but $($csv.Rows.Count) were accepted -- NOT swapped in; previous batch kept."
                        Invoke-SqliteExec -Database $db -Sql "DROP TABLE IF EXISTS $(ConvertTo-SqlIdent $stage);"
                    }
                    else {
                        # NIIN set change vs the previous landed batch of this file.
                        if ($e.NiinColumn -and $prev.Count) {
                            $nc = ConvertTo-SqlIdent $e.NiinColumn
                            $hasOld = (Invoke-SqliteQuery -Database $db -Sql "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?;" -Params @($raw)).Count -gt 0
                            if ($hasOld) {
                                Invoke-SqliteNonQuery -Database $db -Sql @"
INSERT INTO fb_change_log (batch_number, previous_batch, file_stem, niin, change)
SELECT ?, ?, ?, n, 'added'   FROM (SELECT DISTINCT trim($nc) AS n FROM $(ConvertTo-SqlIdent $stage) WHERE $nc IS NOT NULL
                                    EXCEPT SELECT DISTINCT trim($nc) FROM $(ConvertTo-SqlIdent $raw) WHERE $nc IS NOT NULL)
UNION ALL
SELECT ?, ?, ?, n, 'dropped' FROM (SELECT DISTINCT trim($nc) AS n FROM $(ConvertTo-SqlIdent $raw) WHERE $nc IS NOT NULL
                                    EXCEPT SELECT DISTINCT trim($nc) FROM $(ConvertTo-SqlIdent $stage) WHERE $nc IS NOT NULL);
"@ -Params @($BatchNumber, $prev[0].batch_number, $stem, $BatchNumber, $prev[0].batch_number, $stem) | Out-Null
                                $chg = Invoke-SqliteQuery -Database $db -Sql "SELECT change, COUNT(*) AS n FROM fb_change_log WHERE batch_number = ? AND file_stem = ? GROUP BY change;" -Params @($BatchNumber, $stem)
                                $added = ($chg | Where-Object change -eq 'added').n; $dropped = ($chg | Where-Object change -eq 'dropped').n
                                if ($added -or $dropped) { Add-FlisIssue $run INFO $stem 'NiinChange' ("NIINs vs batch {0}: {1} added, {2} dropped (fb_change_log)." -f $prev[0].batch_number, [int]$added, [int]$dropped) }
                            }
                        }
                        Invoke-SqliteExec -Database $db -Sql "BEGIN; DROP TABLE IF EXISTS $(ConvertTo-SqlIdent $raw); ALTER TABLE $(ConvertTo-SqlIdent $stage) RENAME TO $(ConvertTo-SqlIdent $raw); COMMIT;"
                        $status = 'LANDED'

                        # Row-count trend.
                        if ($prev.Count -and $prev[0].rows_landed -gt 0) {
                            $pct = 100.0 * ($prev[0].rows_landed - $landed) / $prev[0].rows_landed
                            if ($pct -ge $RowDropWarnPct) { Add-FlisIssue $run WARN $stem 'RowDrop' ("Row count fell {0:N0}% vs batch {1} ({2:N0} -> {3:N0})." -f $pct, $prev[0].batch_number, $prev[0].rows_landed, $landed) }
                        }

                        # Fill rates, and drops vs the previous batch.
                        $total = $csv.Rows.Count
                        $prevFill = @{}
                        if ($prev.Count) {
                            foreach ($pf in (Invoke-SqliteQuery -Database $db -Sql 'SELECT column_name, pct FROM fb_fill_rate WHERE file_stem = ? AND batch_number = ?;' -Params @($stem, $prev[0].batch_number))) { $prevFill[$pf.column_name] = [double]$pf.pct }
                        }
                        $emptyCols = New-Object System.Collections.Generic.List[string]
                        # Empty values landed as NULL, so COUNT(col) = non-empty count.
                        $countSql = 'SELECT ' + (($cols | ForEach-Object -Begin { $k = 0 } -Process { "COUNT($(ConvertTo-SqlIdent $_)) AS c$k"; $k++ }) -join ', ') + " FROM $(ConvertTo-SqlIdent $raw);"
                        $counts = (Invoke-SqliteQuery -Database $db -Sql $countSql)[0]
                        Invoke-SqliteExec -Database $db -Sql 'BEGIN;'
                        for ($c = 0; $c -lt $cols.Count; $c++) {
                            $ne = [long]$counts."c$c"
                            $pctFill = if ($total) { [math]::Round(100.0 * $ne / $total, 2) } else { 0 }
                            Invoke-SqliteNonQuery -Database $db -Sql 'INSERT INTO fb_fill_rate (run_id, batch_number, file_stem, column_name, non_empty, total, pct) VALUES (?,?,?,?,?,?,?);' -Params @($run.RunId, $BatchNumber, $stem, $cols[$c], $ne, $total, $pctFill) | Out-Null
                            if ($total -and $ne -eq 0) { $emptyCols.Add($cols[$c]) }
                            if ($prevFill.ContainsKey($cols[$c]) -and ($prevFill[$cols[$c]] - $pctFill) -ge $FillDropWarnPoints) {
                                Add-FlisIssue $run WARN $stem 'FillDrop' ("Column '{0}' fill rate fell from {1:N1}% to {2:N1}% vs batch {3}." -f $cols[$c], $prevFill[$cols[$c]], $pctFill, $prev[0].batch_number) -Column $cols[$c]
                            }
                        }
                        Invoke-SqliteExec -Database $db -Sql 'COMMIT;'
                        if ($emptyCols.Count -and $total) {
                            $list = ($emptyCols | Select-Object -First 15) -join ', '
                            if ($emptyCols.Count -gt 15) { $list += ", +$($emptyCols.Count - 15) more" }
                            Add-FlisIssue $run INFO $stem 'EmptyColumns' ("{0} of {1} columns are completely empty in this batch: {2}" -f $emptyCols.Count, $cols.Count, $list)
                        }
                    }
                }
            }
            $errs = @($run.Issues | Where-Object { $_.File -eq $stem -and $_.Severity -eq 'ERROR' }).Count - $errBefore
            $warns = @($run.Issues | Where-Object { $_.File -eq $stem -and $_.Severity -eq 'WARN' }).Count
            $fileIssues = @($run.Issues | Select-Object -Skip $issuesBefore)
            $fErr = @($fileIssues | Where-Object Severity -eq 'ERROR').Count
            $fWarn = @($fileIssues | Where-Object Severity -eq 'WARN').Count
            if ($status -ne 'LANDED' -and -not $canLand) { $status = 'NOT LANDED' }
            $line = switch ($status) {
                'LANDED'      { "   LANDED {0:N0} rows into fb_{1}" -f $landed, $stem }
                'QUARANTINED' { "   QUARANTINED -- fb_{0} still holds its previous batch (if any)" -f $stem }
                default       { '   checked only -- batch problems prevent landing' }
            }
            Write-FlisReportLine $run $line $(if ($status -eq 'LANDED') { 'Green' } else { 'Red' })
            if ($run.ConsoleDetail -eq 'Summary' -and ($fErr -or $fWarn -or $status -ne 'LANDED')) {
                $c = if ($fErr -or $status -ne 'LANDED') { 'Red' } else { 'Yellow' }
                Write-Host ("  {0,-12} {1,-36} {2,9:N0} rows   {3} error(s), {4} warning(s)" -f $status, $stem, $landed, $fErr, $fWarn) -ForegroundColor $c
            }
            elseif ($run.ConsoleDetail -eq 'Files') { Write-FlisDigest -Run $run -Issues $fileIssues -IncludeInfo }
            $run.Files.Add([PSCustomObject]@{ Stem = $stem; Status = $status; File = $file.Name })
            Invoke-SqliteNonQuery -Database $db -Sql @'
INSERT INTO fb_load_log (run_id, batch_number, file_stem, file_name, status, file_size, file_sha256, encoding, line_endings, columns,
  rows_read, rows_rejected, rows_landed, duplicate_rows, errors, warnings) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);
'@ -Params @($run.RunId, $BatchNumber, $stem, $file.Name, $status, [long]$file.Length, $sha,
               $(if ($csv) { $csv.EncodingName }), $(if ($csv) { $csv.LineEndings }), $(if ($csv) { $csv.Headers -join '|' }),
               $(if ($csv) { $csv.RecordsRead } else { 0 }), $(if ($csv) { $csv.RejectedRecords } else { 0 }), $landed, $dupCount, $errs, $warns) | Out-Null
        }

        foreach ($f in ($run.Files | Where-Object { $_.Status -in 'MISSING','UNKNOWN','RECOGNIZED' })) {
            Invoke-SqliteNonQuery -Database $db -Sql 'INSERT INTO fb_load_log (run_id, batch_number, file_stem, file_name, status) VALUES (?,?,?,?,?);' -Params @($run.RunId, $BatchNumber, $f.Stem, $f.File, $f.Status) | Out-Null
        }
        # Persist every issue.
        Invoke-SqliteExec -Database $db -Sql 'BEGIN;'
        foreach ($i in $run.Issues) {
            Invoke-SqliteNonQuery -Database $db -Sql 'INSERT INTO fb_issue_log (run_id, batch_number, file_stem, severity, code, line, column_name, message) VALUES (?,?,?,?,?,?,?,?);' `
                -Params @($run.RunId, $BatchNumber, $i.File, $i.Severity, $i.Code, $(if ($i.Line) { $i.Line }), $i.Column, $i.Message) | Out-Null
        }
        Invoke-SqliteExec -Database $db -Sql 'COMMIT;'
    }
    finally { Close-SqliteDb -Database $db }

    # ---- 4. summary --------------------------------------------------------
    Write-Progress -Activity 'FLIS ingest' -Completed
    $nErr = @($run.Issues | Where-Object Severity -eq 'ERROR').Count
    $nWarn = @($run.Issues | Where-Object Severity -eq 'WARN').Count
    Write-FlisReportLine $run "== Summary ==" Cyan -Always
    foreach ($g in ($run.Files | Group-Object Status | Sort-Object Name)) {
        $names = if ($g.Name -eq 'LANDED' -and $run.ConsoleDetail -eq 'Summary') { '' } else { ($g.Group | ForEach-Object Stem | Sort-Object) -join ', ' }
        Write-FlisReportLine $run ("  {0,-12} {1,3}  {2}" -f $g.Name, $g.Count, $names) -Always
    }
    if ($nErr -or $nWarn) {
        Write-Host '  Issues (grouped by file and type; counts in x...):' -ForegroundColor Cyan
        Write-FlisDigest -Run $run -Issues $run.Issues -Max $run.MaxConsoleGroups
    }
    $color = if ($nErr) { 'Red' } elseif ($nWarn) { 'Yellow' } else { 'Green' }
    Write-FlisReportLine $run ("  batch {0}: {1} error(s), {2} warning(s). Full detail: fb_issue_log (run {3}){4}" -f $run.Batch, $nErr, $nWarn, $run.RunId, $(if ($run.ReportPath) { " and $($run.ReportPath)" } else { '' })) $color -Always
    if ($run.ReportPath) {
        # Report: grouped index first, then every issue (capped per group at
        # 200 lines; fb_issue_log has all of them), then the run narrative.
        $out = New-Object System.Collections.Generic.List[string]
        $out.Add("FLIS ingest report -- run $($run.RunId), batch $($run.Batch): $nErr error(s), $nWarn warning(s)")
        $out.Add(''); $out.Add('== Issue index (one line per file + type) ==')
        $groups = @(Get-FlisIssueGroups -Issues $run.Issues)
        foreach ($g in $groups) { $out.Add((Format-FlisGroup $g)) }
        $out.Add(''); $out.Add('== All issues ==')
        foreach ($g in $groups) {
            $out.Add(("-- {0} {1} {2} ({3:N0})" -f $g.Severity, $(if ($g.File) { $g.File } else { '(batch)' }), $g.Code, $g.Count))
            $items = @($run.Issues | Where-Object { $_.Severity -eq $g.Severity -and $_.File -eq $g.File -and $_.Code -eq $g.Code })
            foreach ($i in ($items | Select-Object -First 200)) { $out.Add(("   {0}{1}" -f $(if ($i.Line -gt 0) { "line $($i.Line): " } else { '' }), $i.Message)) }
            if ($items.Count -gt 200) { $out.Add(("   ... {0:N0} more in fb_issue_log (run {1})" -f ($items.Count - 200), $run.RunId)) }
        }
        $out.Add(''); $out.Add('== Run narrative ==')
        foreach ($l in $run.Report) { $out.Add($l) }
        [IO.File]::WriteAllLines($run.ReportPath, $out)
    }

    $result = [PSCustomObject]@{ RunId = $run.RunId; Batch = $run.Batch; Errors = $nErr; Warnings = $nWarn; Files = $run.Files; Issues = $run.Issues; ReportPath = $run.ReportPath }
    if ($nErr -and -not $NoThrow) {
        throw ("FLIS ingest finished with {0} error(s) -- see the messages above, fb_issue_log, or {1}." -f $nErr, $run.ReportPath)
    }
    $result
}
