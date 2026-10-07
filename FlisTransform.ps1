#requires -Version 5.1
<#
.SYNOPSIS
  FLIS normalized layer: schema migrations, and fb_* raw tables -> typed
  sustainment tables per flis_model.psd1, with every conversion problem
  reported. Used by Invoke-FlisPipeline.ps1.

.DESCRIPTION
  Initialize-FlisSchema   applies migrations\NNN_*.sql in order (tracked in
                          schema_migrations); refuses to touch a database that
                          still has the legacy guessed schema.
  Invoke-FlisTransform    for every model target whose source file LANDED in
                          the given ingest run: converts columns (text, ISO
                          date, number, NIIN, CAGE, normalized numbers) in
                          compiled code, then merges by the target's strategy
                          (ItemMaster, Upsert, Snapshot, ReplaceAll, History,
                          RollingStock, RollingForecast).
  Invoke-FlisCoverage     cross-file checks: NIINs not in the item master,
                          item-master NIINs with no data per key table, CAGEs
                          not in vendor_cage (stubbed and reported), PartsList
                          components not in the item master.

  Problems are collected on the same run object FlisIngest.ps1 uses (grouped
  on the console, every occurrence in fb_issue_log and the report).
#>

if (-not (Get-Command Open-SqliteDb -ErrorAction SilentlyContinue)) { . (Join-Path $PSScriptRoot 'SQLiteInterop.ps1') }
if (-not (Get-Command Add-FlisIssue -ErrorAction SilentlyContinue)) { . (Join-Path $PSScriptRoot 'FlisIngest.ps1') }

if (-not ('Flis.Transform' -as [type])) {
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace Flis
{
    public class Finding
    {
        public string Column;     // target column
        public string Code;       // DateUnparseable, NumberUnparseable, NiinFormat, CageFormat, DateZero
        public int Count;
        public List<string> Samples = new List<string>();   // "line N 'value'"
    }

    public class TransformResult
    {
        public int RowsRead;
        public int RowsWritten;
        public List<Finding> Findings = new List<Finding>();
    }

    public static class Transform
    {
        const string DLL = "winsqlite3.dll";
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int nBytes, out IntPtr stmt, IntPtr tail);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_step(IntPtr stmt);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_reset(IntPtr stmt);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_finalize(IntPtr stmt);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr sqlite3_column_text(IntPtr stmt, int col);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_column_bytes(IntPtr stmt, int col);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_bind_text(IntPtr stmt, int index, byte[] value, int nBytes, IntPtr destructor);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_bind_double(IntPtr stmt, int index, double value);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_bind_int64(IntPtr stmt, int index, long value);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_bind_null(IntPtr stmt, int index);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr sqlite3_errmsg(IntPtr db);
        static readonly IntPtr TRANSIENT = new IntPtr(-1);
        const int SQLITE_ROW = 100, SQLITE_DONE = 101;

        static byte[] Z(string s) { var b = Encoding.UTF8.GetBytes(s); var z = new byte[b.Length + 1]; Buffer.BlockCopy(b, 0, z, 0, b.Length); return z; }
        static string Err(IntPtr db)
        {
            IntPtr p = sqlite3_errmsg(db); if (p == IntPtr.Zero) return "?";
            int n = 0; while (Marshal.ReadByte(p, n) != 0) n++;
            var b = new byte[n]; Marshal.Copy(p, b, 0, n); return Encoding.UTF8.GetString(b);
        }
        static string Col(IntPtr stmt, int c)
        {
            IntPtr p = sqlite3_column_text(stmt, c);
            if (p == IntPtr.Zero) return null;
            int n = sqlite3_column_bytes(stmt, c);
            var b = new byte[n]; if (n > 0) Marshal.Copy(p, b, 0, n);
            return Encoding.UTF8.GetString(b);
        }
        static string Q(string ident) { return "\"" + ident.Replace("\"", "\"\"") + "\""; }

        // ---- converters (public so PowerShell and tests can call them) ----
        static readonly string[] Mon = { "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC" };
        public static int PivotTwoDigitYear = (DateTime.Now.Year % 100) + 1;

        static int Year2(int yy) { return yy > PivotTwoDigitYear ? 1900 + yy : 2000 + yy; }
        static string Iso(int y, int m, int d)
        {
            if (y < 1900 || y > 2200 || m < 1 || m > 12 || d < 1 || d > DateTime.DaysInMonth(y, m)) return null;
            return string.Format("{0:D4}-{1:D2}-{2:D2}", y, m, d);
        }
        static readonly Regex ReIso = new Regex(@"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$");
        static readonly Regex ReUs = new Regex(@"^(\d{1,2})/(\d{1,2})/(\d{4}|\d{2})$");
        static readonly Regex ReMon = new Regex(@"^(\d{1,2})[- ]([A-Za-z]{3})[- ](\d{4}|\d{2})$");
        static readonly Regex ReYmd = new Regex(@"^(\d{4})(\d{2})(\d{2})$");
        static readonly Regex ReJul = new Regex(@"^(\d{2})(\d{3})$");

        // Returns ISO date, or null with code = "DateUnparseable" / "DateZero".
        public static string ParseDate(string v, out string code)
        {
            code = null;
            if (v == null) return null;
            string s = v.Trim();
            if (s.Length == 0) return null;
            // drop a trailing time part: "3/15/2024 12:00:00 AM", "2024-03-15T00:00:00", "2024-03-15 00:00"
            int t = s.IndexOfAny(new[] { 'T', ' ' });
            if (t > 0 && Regex.IsMatch(s.Substring(t + 1).Trim(), @"^\d{1,2}:\d{2}")) s = s.Substring(0, t);
            Match m;
            string r = null;
            if ((m = ReIso.Match(s)).Success) r = Iso(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), int.Parse(m.Groups[3].Value));
            else if ((m = ReUs.Match(s)).Success)
            {
                int y = int.Parse(m.Groups[3].Value); if (m.Groups[3].Value.Length == 2) y = Year2(y);
                r = Iso(y, int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value));   // US month/day
            }
            else if ((m = ReMon.Match(s)).Success)
            {
                int mi = Array.IndexOf(Mon, m.Groups[2].Value.ToUpperInvariant());
                int y = int.Parse(m.Groups[3].Value); if (m.Groups[3].Value.Length == 2) y = Year2(y);
                if (mi >= 0) r = Iso(y, mi + 1, int.Parse(m.Groups[1].Value));
            }
            else if ((m = ReYmd.Match(s)).Success) r = Iso(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), int.Parse(m.Groups[3].Value));
            else if ((m = ReJul.Match(s)).Success)
            {
                if (s == "00000") { code = "DateZero"; return null; }
                int y = Year2(int.Parse(m.Groups[1].Value)), doy = int.Parse(m.Groups[2].Value);
                if (doy >= 1 && doy <= (DateTime.IsLeapYear(y) ? 366 : 365)) r = new DateTime(y, 1, 1).AddDays(doy - 1).ToString("yyyy-MM-dd");
            }
            if (r == null) code = "DateUnparseable";
            return r;
        }

        // Accepts 1234, -1234.5, 1,234.50, $1,234.50 . Returns null + code otherwise.
        public static bool TryParseNumber(string v, out double d, out string code)
        {
            d = 0; code = null;
            if (v == null) return false;
            string s = v.Trim();
            if (s.Length == 0) return false;
            string c = s.Replace("$", "").Replace(" ", "");
            if (Regex.IsMatch(c, @"^[+-]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$|^[+-]?\.\d+$"))
            {
                if (double.TryParse(c.Replace(",", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return true;
            }
            code = "NumberUnparseable";
            return false;
        }

        static readonly Regex ReNiin = new Regex(@"^\d{9}$");
        static readonly Regex ReCage = new Regex(@"^[0-9A-Z]{5}$");
        public static string NormNumber(string v) { if (v == null) return null; var s = Regex.Replace(v.ToUpperInvariant(), @"[-\s]", ""); return s.Length == 0 ? null : s; }
        public static string NormPart(string v) { if (v == null) return null; var s = Regex.Replace(v.ToUpperInvariant(), @"[-\s/.,_#()'""*+:;]", ""); return s.Length == 0 ? null : s; }

        // Copies srcTable -> dstTable, converting each column by type.
        // srcCols[i] -> dstCols[i] with types[i]; extra (name, value) pairs are
        // bound as constants (e.g. source_batch). Caller owns the transaction.
        public static TransformResult CopyConvert(IntPtr db, string srcTable, string[] srcCols, string dstTable, string[] dstCols, string[] types,
                                                  string[] constNames, string[] constValues)
        {
            var res = new TransformResult();
            var findings = new Dictionary<string, Finding>();
            var sel = new StringBuilder("SELECT ");
            for (int i = 0; i < srcCols.Length; i++) { if (i > 0) sel.Append(','); sel.Append(Q(srcCols[i])); }
            sel.Append(", fb_line FROM ").Append(Q(srcTable)).Append(" ORDER BY fb_line");
            var ins = new StringBuilder("INSERT INTO ").Append(Q(dstTable)).Append(" (");
            var ph = new StringBuilder();
            int nOut = dstCols.Length + constNames.Length;
            for (int i = 0; i < dstCols.Length; i++) { if (i > 0) { ins.Append(','); ph.Append(','); } ins.Append(Q(dstCols[i])); ph.Append('?'); }
            for (int i = 0; i < constNames.Length; i++) { ins.Append(',').Append(Q(constNames[i])); ph.Append(",?"); }
            ins.Append(") VALUES (").Append(ph).Append(')');

            IntPtr s1, s2;
            if (sqlite3_prepare_v2(db, Z(sel.ToString()), -1, out s1, IntPtr.Zero) != 0) throw new Exception("select failed: " + Err(db) + " -- " + sel);
            if (sqlite3_prepare_v2(db, Z(ins.ToString()), -1, out s2, IntPtr.Zero) != 0) { sqlite3_finalize(s1); throw new Exception("insert prepare failed: " + Err(db) + " -- " + ins); }
            try
            {
                while (sqlite3_step(s1) == SQLITE_ROW)
                {
                    res.RowsRead++;
                    string line = Col(s1, srcCols.Length);
                    for (int i = 0; i < srcCols.Length; i++)
                    {
                        string v = Col(s1, i);
                        string code = null;
                        string outText = null; bool isNum = false; double num = 0;
                        switch (types[i])
                        {
                            case "RAW": outText = string.IsNullOrEmpty(v) ? null : v; break;
                            case "D": outText = ParseDate(v, out code); break;
                            case "N": isNum = TryParseNumber(v, out num, out code); break;
                            case "NIIN":
                                outText = v == null ? null : v.Trim(); if (outText == "") outText = null;
                                if (outText != null && !ReNiin.IsMatch(outText)) code = "NiinFormat";
                                break;
                            case "CAGE":
                                outText = v == null ? null : v.Trim().ToUpperInvariant(); if (outText == "") outText = null;
                                if (outText != null && !ReCage.IsMatch(outText)) code = "CageFormat";
                                break;
                            case "NORM": outText = NormNumber(v == null ? null : v.Trim()); break;
                            case "PNNORM": outText = NormPart(v == null ? null : v.Trim()); break;
                            default: outText = v == null ? null : v.Trim(); if (outText == "") outText = null; break;
                        }
                        if (code != null)
                        {
                            string k = dstCols[i] + "|" + code;
                            Finding f;
                            if (!findings.TryGetValue(k, out f)) { f = new Finding { Column = dstCols[i], Code = code }; findings[k] = f; res.Findings.Add(f); }
                            f.Count++;
                            if (f.Samples.Count < 10) f.Samples.Add("line " + line + " '" + v + "'");
                        }
                        if (isNum) sqlite3_bind_double(s2, i + 1, num);
                        else if (outText == null) sqlite3_bind_null(s2, i + 1);
                        else { var b = Encoding.UTF8.GetBytes(outText); sqlite3_bind_text(s2, i + 1, b, b.Length, TRANSIENT); }
                    }
                    for (int c = 0; c < constNames.Length; c++)
                    {
                        if (constValues[c] == null) sqlite3_bind_null(s2, dstCols.Length + c + 1);
                        else { var b = Encoding.UTF8.GetBytes(constValues[c]); sqlite3_bind_text(s2, dstCols.Length + c + 1, b, b.Length, TRANSIENT); }
                    }
                    if (sqlite3_step(s2) != SQLITE_DONE) throw new Exception("insert into " + dstTable + " failed at source line " + line + ": " + Err(db));
                    sqlite3_reset(s2);
                    res.RowsWritten++;
                }
            }
            finally { sqlite3_finalize(s1); sqlite3_finalize(s2); }
            return res;
        }
    }
}
"@
}

function ConvertTo-FlisIdent { param([string]$Name) '"' + $Name.Replace('"', '""') + '"' }

# ---------------------------------------------------------------------------
# Schema migrations
# ---------------------------------------------------------------------------

# NOTE: Invoke-SqliteQuery returns its rows with a leading comma (so a
# 1-row result stays an array). Never wrap it in @(...) -- that nests the
# array and an empty result then counts as one row. Use (Invoke-SqliteQuery ...)
# or pipe it in parentheses: @((Invoke-SqliteQuery ...) | ForEach-Object name).

function Initialize-FlisSchema {
    # Applies migrations\NNN_name.sql not yet recorded in schema_migrations,
    # each in its own transaction. Returns the list applied.
    param([IntPtr]$Database, $Run, [string]$MigrationsFolder = (Join-Path $PSScriptRoot 'migrations'))
    $tables = @((Invoke-SqliteQuery -Database $Database -Sql "SELECT name FROM sqlite_master WHERE type='table';") | ForEach-Object name)
    if ($tables -notcontains 'schema_migrations') {
        $legacy = @('qualified_product_list','solicitation','contract','supply_class','ref_code_value') | Where-Object { $tables -contains $_ }
        if ($legacy.Count) {
            throw ("This database has the LEGACY (guessed) schema (tables: {0}) and no migration history. The FLIS pipeline will not modify it. Move or rename the file and run again to build a fresh database from the real FLIS files." -f ($legacy -join ', '))
        }
        Invoke-SqliteExec -Database $Database -Sql "CREATE TABLE schema_migrations (version INTEGER PRIMARY KEY, name TEXT NOT NULL, applied_at TEXT NOT NULL DEFAULT (datetime('now')));"
    }
    $done = @{}
    foreach ($r in (Invoke-SqliteQuery -Database $Database -Sql 'SELECT version FROM schema_migrations;')) { $done[[int]$r.version] = $true }
    $applied = @()
    foreach ($f in (Get-ChildItem -LiteralPath $MigrationsFolder -Filter '*.sql' | Sort-Object Name)) {
        if ($f.Name -notmatch '^(\d{3})_(.+)\.sql$') { if ($Run) { Add-FlisIssue $Run WARN '(schema)' 'MigrationName' "Ignoring '$($f.Name)': migration files must be named NNN_name.sql." }; continue }
        $v = [int]$matches[1]
        if ($done[$v]) { continue }
        $sql = [IO.File]::ReadAllText($f.FullName)
        Invoke-SqliteExec -Database $Database -Sql 'BEGIN;'
        try {
            Invoke-SqliteExec -Database $Database -Sql $sql
            Invoke-SqliteNonQuery -Database $Database -Sql 'INSERT INTO schema_migrations (version, name) VALUES (?, ?);' -Params @($v, $matches[2]) | Out-Null
            Invoke-SqliteExec -Database $Database -Sql 'COMMIT;'
        } catch {
            try { Invoke-SqliteExec -Database $Database -Sql 'ROLLBACK;' } catch { }
            throw "Migration $($f.Name) failed and was rolled back: $($_.Exception.Message)"
        }
        $applied += $f.Name
        if ($Run) { Add-FlisIssue $Run INFO '(schema)' 'Migration' "Applied migration $($f.Name)." }
    }
    $applied
}

# ---------------------------------------------------------------------------
# Model
# ---------------------------------------------------------------------------

function Get-FlisModel {
    param([string]$Path = (Join-Path $PSScriptRoot 'flis_model.psd1'))
    $m = Import-PowerShellDataFile -LiteralPath $Path
    foreach ($t in $m.Targets) {
        $t.Specs = @(foreach ($c in $t.Columns) {
            if ($c -notmatch '^([^=]+)=([a-z0-9_]+):(T|RAW|D|N|NIIN|CAGE|NORM|PNNORM)$') { throw "flis_model.psd1: bad column spec '$c' in target $($t.Target)." }
            [PSCustomObject]@{ Src = $matches[1]; Dst = $matches[2]; Type = $matches[3] }
        })
    }
    $m
}

function Add-FlisFindings {
    # Turns engine findings into grouped issues. NIIN/CAGE format problems are
    # already reported (per row) by the raw ingest, so here they are INFO.
    param($Run, [string]$Target, $Result)
    foreach ($f in $Result.Findings) {
        $sev = switch ($f.Code) { 'DateUnparseable' { 'WARN' } 'NumberUnparseable' { 'WARN' } 'DateZero' { 'INFO' } default { 'INFO' } }
        $what = switch ($f.Code) {
            'DateUnparseable'   { 'are not a recognizable date (stored as NULL)' }
            'NumberUnparseable' { 'are not a number (stored as NULL)' }
            'DateZero'          { "are the 'no date' value 00000 (stored as NULL)" }
            'NiinFormat'        { 'are not 9-digit NIINs (kept as delivered; see the raw-ingest warning)' }
            'CageFormat'        { 'are not 5-character CAGEs (kept as delivered)' }
            default             { $f.Code }
        }
        $more = if ($f.Count -gt $f.Samples.Count) { " (+$($f.Count - $f.Samples.Count) more)" } else { '' }
        Add-FlisIssue $Run $sev $Target $f.Code ("{0:N0} of {1:N0} values in {2} {3}. {4}{5}" -f $f.Count, $Result.RowsRead, $f.Column, $what, ($f.Samples -join '; '), $more) -Column $f.Column
    }
}

# Stage a target's rows from one fb_ table into a temp table shaped like the target.
function Copy-FlisToStaging {
    param([IntPtr]$Database, $Run, $Target, [string]$SourceStem, [string]$Staging, [string[]]$ConstNames = @(), [string[]]$ConstValues = @())
    $fb = "fb_$SourceStem"
    $have = @{}
    foreach ($c in (Invoke-SqliteQuery -Database $Database -Sql "PRAGMA table_info($(ConvertTo-FlisIdent $fb));")) { $have[$c.name] = $true }
    $missing = @($Target.Specs | Where-Object { -not $have[$_.Src] } | ForEach-Object Src | Select-Object -Unique)
    if ($missing.Count) {
        Add-FlisIssue $Run ERROR $Target.Target 'ModelMismatch' ("{0} lacks column(s) {1} that flis_model.psd1 maps -- target not refreshed. The registry and model disagree; regenerate the model." -f $fb, ($missing -join ', '))
        return $null
    }
    $r = [Flis.Transform]::CopyConvert($Database, $fb, [string[]]@($Target.Specs | ForEach-Object Src), $Staging,
        [string[]]@($Target.Specs | ForEach-Object Dst), [string[]]@($Target.Specs | ForEach-Object Type), [string[]]$ConstNames, [string[]]$ConstValues)
    Add-FlisFindings -Run $Run -Target $Target.Target -Result $r
    $r
}

function Get-FlisDstColumns { param($Target) @($Target.Specs | ForEach-Object Dst | Select-Object -Unique) }

function Remove-FlisStagingDuplicateKeys {
    # Keeps the first row per key; reports blank keys and duplicates.
    param([IntPtr]$Database, $Run, [string]$TargetName, [string]$Staging, [string]$Key, [string]$Context)
    $k = ConvertTo-FlisIdent $Key
    $blank = (Invoke-SqliteQuery -Database $Database -Sql "SELECT COUNT(*) AS n FROM $Staging WHERE $k IS NULL;")[0].n
    if ($blank) {
        Add-FlisIssue $Run WARN $TargetName 'KeyBlank' ("{0:N0} row(s) have no {1} and were skipped ({2})." -f $blank, $Key, $Context)
        Invoke-SqliteExec -Database $Database -Sql "DELETE FROM $Staging WHERE $k IS NULL;"
    }
    $dups = (Invoke-SqliteQuery -Database $Database -Sql "SELECT $k AS k, COUNT(*) AS n FROM $Staging GROUP BY $k HAVING COUNT(*) > 1 ORDER BY n DESC;")
    if ($dups.Count) {
        $ex = ($dups | Select-Object -First 10 | ForEach-Object { "$($_.k) x$($_.n)" }) -join '; '
        Add-FlisIssue $Run WARN $TargetName 'KeyDuplicate' ("{0:N0} {1} value(s) appear more than once in {2}; kept the first row of each. {3}" -f $dups.Count, $Key, $Context, $ex)
        Invoke-SqliteExec -Database $Database -Sql "DELETE FROM $Staging WHERE rowid NOT IN (SELECT MIN(rowid) FROM $Staging GROUP BY $k);"
    }
}

# ---------------------------------------------------------------------------
# Strategies
# ---------------------------------------------------------------------------

function Invoke-FlisUpsert {
    param([IntPtr]$Database, $Run, $Target, [long]$Batch, [string]$SourceStem, [switch]$ItemMaster)
    $t = $Target.Target; $key = $Target.Key; $cols = Get-FlisDstColumns $Target
    Invoke-SqliteExec -Database $Database -Sql "DROP TABLE IF EXISTS temp.tx; CREATE TEMP TABLE tx AS SELECT $(($cols | ForEach-Object { ConvertTo-FlisIdent $_ }) -join ', ') FROM $t WHERE 0;"
    $r = Copy-FlisToStaging -Database $Database -Run $Run -Target $Target -SourceStem $SourceStem -Staging 'tx'
    if (-not $r) { return }
    Remove-FlisStagingDuplicateKeys -Database $Database -Run $Run -TargetName $t -Staging 'tx' -Key $key -Context "fb_$SourceStem"
    $now = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $k = ConvertTo-FlisIdent $key
    $set = (($cols | Where-Object { $_ -ne $key }) | ForEach-Object { "$(ConvertTo-FlisIdent $_) = (SELECT x.$(ConvertTo-FlisIdent $_) FROM tx x WHERE x.$k = $t.$k)" }) -join ",`n  "
    $life = "source_batch = $Batch, last_seen_batch = $Batch, updated_at = '$now'" + $(if ($ItemMaster) { ', in_latest_batch = 1' } else { ', is_stub = 0' })
    $before = (Invoke-SqliteQuery -Database $Database -Sql "SELECT COUNT(*) AS n FROM $t;")[0].n
    Invoke-SqliteExec -Database $Database -Sql 'BEGIN;'
    try {
        $updated = Invoke-SqliteNonQuery -Database $Database -Sql "UPDATE $t SET`n  $set,`n  $life`nWHERE $k IN (SELECT $k FROM tx);"
        $colList = ($cols | ForEach-Object { ConvertTo-FlisIdent $_ }) -join ', '
        $lifeCols = 'source_batch, first_seen_batch, last_seen_batch, updated_at' + $(if ($ItemMaster) { ', in_latest_batch' } else { '' })
        $lifeVals = "$Batch, $Batch, $Batch, '$now'" + $(if ($ItemMaster) { ', 1' } else { '' })
        $inserted = Invoke-SqliteNonQuery -Database $Database -Sql "INSERT INTO $t ($colList, $lifeCols) SELECT $colList, $lifeVals FROM tx WHERE $k NOT IN (SELECT $k FROM $t);"
        $dropped = 0
        if ($ItemMaster) {
            $dropped = Invoke-SqliteNonQuery -Database $Database -Sql "UPDATE $t SET in_latest_batch = 0 WHERE in_latest_batch = 1 AND $k NOT IN (SELECT $k FROM tx);"
        }
        Invoke-SqliteExec -Database $Database -Sql 'COMMIT;'
    } catch { try { Invoke-SqliteExec -Database $Database -Sql 'ROLLBACK;' } catch { }; throw }
    Add-FlisIssue $Run INFO $t 'Merged' ("{0:N0} rows from fb_{1}: {2:N0} updated, {3:N0} new." -f $r.RowsRead, $SourceStem, $updated, $inserted)
    if ($ItemMaster) {
        if ($dropped) {
            $sev = if ($before -and $dropped -ge [math]::Max(10, 0.1 * $before)) { 'WARN' } else { 'INFO' }
            Add-FlisIssue $Run $sev $t 'NiinDropped' ("{0:N0} NIIN(s) are no longer in SegmentA (in_latest_batch set to 0; their history is kept)." -f $dropped)
        }
        if ($inserted -and $before) { Add-FlisIssue $Run INFO $t 'NiinNew' ("{0:N0} NIIN(s) are new in this batch." -f $inserted) }
    }
}

function Invoke-FlisSnapshot {
    param([IntPtr]$Database, $Run, $Target, [long]$Batch, [string]$SourceStem, [switch]$ReplaceAll, [bool]$ItemMasterRefreshed)
    $t = $Target.Target; $cols = Get-FlisDstColumns $Target
    $colList = ($cols | ForEach-Object { ConvertTo-FlisIdent $_ }) -join ', '
    Invoke-SqliteExec -Database $Database -Sql "DROP TABLE IF EXISTS temp.tx; CREATE TEMP TABLE tx AS SELECT $colList, source_batch FROM $t WHERE 0;"
    $r = Copy-FlisToStaging -Database $Database -Run $Run -Target $Target -SourceStem $SourceStem -Staging 'tx' -ConstNames @('source_batch') -ConstValues @("$Batch")
    if (-not $r) { return }
    Invoke-SqliteExec -Database $Database -Sql 'BEGIN;'
    try {
        if ($ReplaceAll -or -not $Target.Key) {
            $deleted = Invoke-SqliteNonQuery -Database $Database -Sql "DELETE FROM $t;"
        } else {
            $k = ConvertTo-FlisIdent $Target.Key
            # Replace rows for NIINs in this file, plus NIINs active in this
            # batch's item master (an active NIIN with no rows here now has none).
            # NIINs that dropped out of the batch keep their last rows.
            $active = if ($ItemMasterRefreshed) { " OR $k IN (SELECT niin FROM item_niin WHERE in_latest_batch = 1)" } else { '' }
            $deleted = Invoke-SqliteNonQuery -Database $Database -Sql "DELETE FROM $t WHERE $k IN (SELECT $k FROM tx)$active;"
            $blank = (Invoke-SqliteQuery -Database $Database -Sql "SELECT COUNT(*) AS n FROM tx WHERE $k IS NULL;")[0].n
            if ($blank) { Add-FlisIssue $Run WARN $t 'KeyBlank' ("{0:N0} row(s) from fb_{1} have no {2}; landed anyway but they can't be tied to an item." -f $blank, $SourceStem, $Target.Key) }
        }
        $ins = Invoke-SqliteNonQuery -Database $Database -Sql "INSERT INTO $t ($colList, source_batch) SELECT $colList, source_batch FROM tx;"
        if ($Target.PostSql) { Invoke-SqliteExec -Database $Database -Sql $Target.PostSql }
        Invoke-SqliteExec -Database $Database -Sql 'COMMIT;'
    } catch { try { Invoke-SqliteExec -Database $Database -Sql 'ROLLBACK;' } catch { }; throw }
    Add-FlisIssue $Run INFO $t 'Merged' ("{0:N0} rows from fb_{1} (replaced {2:N0} older rows)." -f $ins, $SourceStem, $deleted)
}

function Invoke-FlisHistory {
    # Append-only on dedup_key across all landed sources of the target.
    param([IntPtr]$Database, $Run, $Target, [long]$Batch, [string[]]$SourceStems)
    $t = $Target.Target; $cols = Get-FlisDstColumns $Target
    $colList = ($cols | ForEach-Object { ConvertTo-FlisIdent $_ }) -join ', '
    Invoke-SqliteExec -Database $Database -Sql "DROP TABLE IF EXISTS temp.tx; CREATE TEMP TABLE tx AS SELECT $colList FROM $t WHERE 0; ALTER TABLE tx ADD COLUMN src TEXT; ALTER TABLE tx ADD COLUMN dedup_key TEXT;"
    $read = 0
    foreach ($s in $SourceStems) {
        $r = Copy-FlisToStaging -Database $Database -Run $Run -Target $Target -SourceStem $s -Staging 'tx' -ConstNames @('src') -ConstValues @($s)
        if (-not $r) { return }
        $read += $r.RowsRead
    }
    $keyExpr = ($Target.KeyColumns | ForEach-Object { "coalesce(CAST($(ConvertTo-FlisIdent $_) AS TEXT), '')" }) -join " || '|' || "
    Invoke-SqliteExec -Database $Database -Sql "UPDATE tx SET dedup_key = $keyExpr;"
    $nonKey = @($cols | Where-Object { $Target.KeyColumns -notcontains $_ })
    $rowSig = ($cols | ForEach-Object { "coalesce(CAST($(ConvertTo-FlisIdent $_) AS TEXT), '')" }) -join " || '|' || "

    # Same key, different non-key values inside this delivery = the key is not unique enough.
    $coll = (Invoke-SqliteQuery -Database $Database -Sql "SELECT dedup_key AS k, COUNT(DISTINCT $rowSig) AS v FROM tx GROUP BY dedup_key HAVING COUNT(DISTINCT $rowSig) > 1;")
    if ($coll.Count) {
        Add-FlisIssue $Run WARN $t 'KeyCollision' ("{0:N0} dedup key(s) ({1}) cover rows with DIFFERENT values in this delivery; kept the first of each. The dedup key may be too narrow (Gitea #15). Examples: {2}" -f $coll.Count, ($Target.KeyColumns -join '+'), (($coll | Select-Object -First 5 | ForEach-Object k) -join '; '))
    }
    $distinct = (Invoke-SqliteQuery -Database $Database -Sql 'SELECT COUNT(DISTINCT dedup_key) AS n FROM tx;')[0].n
    if ($read -gt $distinct) {
        $both = (Invoke-SqliteQuery -Database $Database -Sql "SELECT COUNT(*) AS n FROM (SELECT dedup_key FROM tx GROUP BY dedup_key HAVING COUNT(DISTINCT src) > 1);")[0].n
        Add-FlisIssue $Run INFO $t 'Deduplicated' ("{0:N0} rows read from {1}; {2:N0} distinct after dedup ({3:N0} keys present in more than one file)." -f $read, ($SourceStems -join ' + '), $distinct, $both)
    }
    Invoke-SqliteExec -Database $Database -Sql @"
DROP TABLE IF EXISTS temp.txu;
CREATE TEMP TABLE txu AS SELECT * FROM tx WHERE rowid IN (SELECT MIN(rowid) FROM tx GROUP BY dedup_key);
CREATE UNIQUE INDEX temp.ix_txu ON txu (dedup_key);
UPDATE txu SET src = (SELECT group_concat(s, ',') FROM (SELECT DISTINCT src AS s FROM tx WHERE tx.dedup_key = txu.dedup_key ORDER BY s));
"@
    Invoke-SqliteExec -Database $Database -Sql 'BEGIN;'
    try {
        # Seen again: record any non-key value changes, then apply them.
        $changed = 0
        foreach ($c in $nonKey) {
            $ci = ConvertTo-FlisIdent $c
            $changed += Invoke-SqliteNonQuery -Database $Database -Sql "INSERT INTO history_value_change (table_name, dedup_key, column_name, old_value, new_value, batch_number) SELECT '$t', h.dedup_key, '$c', CAST(h.$ci AS TEXT), CAST(x.$ci AS TEXT), $Batch FROM $t h JOIN txu x ON x.dedup_key = h.dedup_key WHERE h.$ci IS NOT x.$ci;"
        }
        $set = ($nonKey | ForEach-Object { "$(ConvertTo-FlisIdent $_) = (SELECT x.$(ConvertTo-FlisIdent $_) FROM txu x WHERE x.dedup_key = $t.dedup_key)" }) -join ",`n  "
        $seen = Invoke-SqliteNonQuery -Database $Database -Sql "UPDATE $t SET`n  $set,`n  last_seen_batch = $Batch,`n  source_files = (SELECT x.src FROM txu x WHERE x.dedup_key = $t.dedup_key)`nWHERE dedup_key IN (SELECT dedup_key FROM txu);"
        $new = Invoke-SqliteNonQuery -Database $Database -Sql "INSERT INTO $t ($colList, dedup_key, source_files, first_seen_batch, last_seen_batch) SELECT $colList, dedup_key, src, $Batch, $Batch FROM txu WHERE dedup_key NOT IN (SELECT dedup_key FROM $t);"
        Invoke-SqliteExec -Database $Database -Sql 'COMMIT;'
    } catch { try { Invoke-SqliteExec -Database $Database -Sql 'ROLLBACK;' } catch { }; throw }
    Add-FlisIssue $Run INFO $t 'Merged' ("{0:N0} new history row(s), {1:N0} seen again (history is never deleted)." -f $new, $seen)
    if ($changed) { Add-FlisIssue $Run INFO $t 'ValuesChanged' ("{0:N0} value change(s) on previously delivered rows (e.g. contract status) -- applied; old values in history_value_change." -f $changed) }
}

$Script:FlisMonthNames = @{ January = 1; February = 2; March = 3; April = 4; May = 5; June = 6; July = 7; August = 8; September = 9; October = 10; November = 11; December = 12
                           Jan = 1; Feb = 2; Mar = 3; Apr = 4; Jun = 6; Jul = 7; Aug = 8; Sep = 9; Oct = 10; Nov = 11; Dec = 12 }

function Invoke-FlisRolling {
    # Wide month columns -> long rows with the windowed update rule:
    #   months in [current - PriorMonths, current]  overwrite with the latest pull
    #   older months                                insert only if missing
    #   a blank cell never overwrites a stored value
    #   stock: months after the current month are skipped (unless -StoreFutureMonths)
    #   forecast: future months are the point -- they follow the overwrite rule
    param([IntPtr]$Database, $Run, $Target, [long]$Batch, [string]$SourceStem, $Registry,
          [datetime]$AsOfMonth, [int]$PriorMonths = 1, [switch]$StoreFutureMonths)
    $t = $Target.Target
    $isStock = $Target.Strategy -eq 'RollingStock'
    $entry = $Registry.Files[$SourceStem]
    $fb = "fb_$SourceStem"
    $fbCols = @((Invoke-SqliteQuery -Database $Database -Sql "PRAGMA table_info($(ConvertTo-FlisIdent $fb));") | ForEach-Object name)
    $months = @(foreach ($c in $fbCols) {
        $m = [regex]::Match($c, $entry.MonthPattern)
        if ($m.Success) { $y = [int]$m.Groups[2].Value; if ($y -lt 100) { $y += 2000 }; [PSCustomObject]@{ Col = $c; Y = $y; M = $Script:FlisMonthNames[$m.Groups[1].Value] } }
    })
    $cur = $AsOfMonth.Year * 12 + $AsOfMonth.Month
    $winStart = $cur - $PriorMonths
    Add-FlisIssue $Run INFO $t 'Window' ("As-of month {0:yyyy-MM}; months {1} through {0:yyyy-MM} take the latest pull, older months are only filled if missing{2}." -f $AsOfMonth, (Get-Date -Year ([math]::Floor(($winStart - 1) / 12)) -Month ((($winStart - 1) % 12) + 1) -Day 1).ToString('yyyy-MM'), $(if ($isStock -and -not $StoreFutureMonths) { '; future months skipped' } else { '' }))

    # Attributes (NIIN, FSC, names, price) via the converter engine.
    $attrTable = if ($isStock) { $null } else { 'dla_forecast_item' }
    Invoke-SqliteExec -Database $Database -Sql "DROP TABLE IF EXISTS temp.ta; CREATE TEMP TABLE ta ($((Get-FlisDstColumns $Target | ForEach-Object { (ConvertTo-FlisIdent $_) + ' ' + $(if ($_ -eq 'latest_mlc_price') { 'REAL' } else { 'TEXT' }) }) -join ', '), fb_line INTEGER);"
    $r = Copy-FlisToStaging -Database $Database -Run $Run -Target $Target -SourceStem $SourceStem -Staging 'ta'
    if (-not $r) { return }

    # Unpivot to long, validating each cell.
    Invoke-SqliteExec -Database $Database -Sql "DROP TABLE IF EXISTS temp.tl; CREATE TEMP TABLE tl (niin TEXT, y INTEGER, m INTEGER, raw TEXT, qty REAL, fb_line INTEGER);"
    Invoke-SqliteExec -Database $Database -Sql 'BEGIN;'
    foreach ($mo in $months) {
        $c = ConvertTo-FlisIdent $mo.Col
        Invoke-SqliteExec -Database $Database -Sql "INSERT INTO tl (niin, y, m, raw, fb_line) SELECT trim(NIIN), $($mo.Y), $($mo.M), trim($c), fb_line FROM $(ConvertTo-FlisIdent $fb) WHERE trim(coalesce($c, '')) <> '';"
    }
    Invoke-SqliteExec -Database $Database -Sql 'COMMIT;'
    # Numeric check: digits with optional sign, thousands commas and decimals.
    Invoke-SqliteExec -Database $Database -Sql "UPDATE tl SET qty = CAST(replace(raw, ',', '') AS REAL) WHERE replace(replace(raw, ',', ''), '-', '') NOT GLOB '*[^0-9.]*' AND raw GLOB '*[0-9]*' AND length(raw) - length(replace(raw, '.', '')) <= 1;"
    $bad = (Invoke-SqliteQuery -Database $Database -Sql "SELECT niin, y, m, raw, fb_line FROM tl WHERE qty IS NULL ORDER BY fb_line LIMIT 10;")
    $badN = (Invoke-SqliteQuery -Database $Database -Sql 'SELECT COUNT(*) AS n FROM tl WHERE qty IS NULL;')[0].n
    if ($badN) {
        Add-FlisIssue $Run WARN $t 'NumberUnparseable' ("{0:N0} month cell(s) are not numbers and were skipped. {1}" -f $badN, (($bad | ForEach-Object { "line $($_.fb_line) $($_.y)-$($_.m) '$($_.raw)'" }) -join '; '))
        Invoke-SqliteExec -Database $Database -Sql 'DELETE FROM tl WHERE qty IS NULL;'
    }
    $blankNiin = (Invoke-SqliteQuery -Database $Database -Sql "SELECT COUNT(*) AS n FROM tl WHERE niin IS NULL OR niin = '';")[0].n
    if ($blankNiin) { Add-FlisIssue $Run WARN $t 'KeyBlank' "$blankNiin month cell(s) belong to rows with no NIIN; skipped."; Invoke-SqliteExec -Database $Database -Sql "DELETE FROM tl WHERE niin IS NULL OR niin = '';" }
    $dupCells = (Invoke-SqliteQuery -Database $Database -Sql 'SELECT COUNT(*) AS n FROM (SELECT niin, y, m FROM tl GROUP BY niin, y, m HAVING COUNT(*) > 1);')[0].n
    if ($dupCells) {
        Add-FlisIssue $Run WARN $t 'KeyDuplicate' ("{0:N0} NIIN/month combination(s) appear on more than one row of fb_{1}; kept the first." -f $dupCells, $SourceStem)
        Invoke-SqliteExec -Database $Database -Sql 'DELETE FROM tl WHERE rowid NOT IN (SELECT MIN(rowid) FROM tl GROUP BY niin, y, m);'
    }
    if ($isStock -and -not $StoreFutureMonths) {
        $fut = Invoke-SqliteNonQuery -Database $Database -Sql "DELETE FROM tl WHERE y * 12 + m > $cur;"
        if ($fut) { Add-FlisIssue $Run INFO $t 'FutureMonths' ("{0:N0} non-blank stock cell(s) are for months after {1:yyyy-MM} and were not stored (use -StoreFutureStockMonths to keep them)." -f $fut, $AsOfMonth) }
    }

    if ($isStock) { $yc = 'report_year'; $mc = 'report_month'; $qc = 'stock_on_hand_qty' } else { $yc = 'forecast_year'; $mc = 'forecast_month'; $qc = 'forecast_qty' }
    $join = "x.niin = $t.niin AND x.y = $t.$yc AND x.m = $t.$mc"
    Invoke-SqliteExec -Database $Database -Sql 'BEGIN;'
    try {
        $differOld = (Invoke-SqliteQuery -Database $Database -Sql "SELECT COUNT(*) AS n FROM tl x JOIN $t ON $join WHERE x.y * 12 + x.m < $winStart AND $t.$qc IS NOT x.qty;")[0].n
        # Inside the window every re-delivered month takes this pull's value and
        # batch (value_batch also records "confirmed by the latest pull").
        $over = (Invoke-SqliteQuery -Database $Database -Sql "SELECT COUNT(*) AS n FROM tl x JOIN $t ON $join WHERE x.y * 12 + x.m >= $winStart AND $t.$qc IS NOT x.qty;")[0].n
        Invoke-SqliteNonQuery -Database $Database -Sql "UPDATE $t SET $qc = (SELECT x.qty FROM tl x WHERE $join), value_batch = $Batch WHERE ($yc * 12 + $mc) >= $winStart AND EXISTS (SELECT 1 FROM tl x WHERE $join);" | Out-Null
        $attrCols = if ($isStock) { ', fsc_code, item_name' } else { '' }
        $attrVals = if ($isStock) { ', (SELECT a.fsc_code FROM ta a WHERE trim(a.niin) = x.niin LIMIT 1), (SELECT a.item_name FROM ta a WHERE trim(a.niin) = x.niin LIMIT 1)' } else { '' }
        $new = Invoke-SqliteNonQuery -Database $Database -Sql "INSERT INTO $t (niin, $yc, $mc, $qc, first_batch, value_batch$attrCols) SELECT x.niin, x.y, x.m, x.qty, $Batch, $Batch$attrVals FROM tl x WHERE NOT EXISTS (SELECT 1 FROM $t WHERE $join);"
        if (-not $isStock) {
            Invoke-SqliteExec -Database $Database -Sql @"
DELETE FROM dla_forecast_item WHERE niin IN (SELECT trim(niin) FROM ta);
INSERT OR IGNORE INTO dla_forecast_item (niin, fsc_code, supply_chain, item_description, unit_of_issue, latest_mlc_price, source_batch)
SELECT trim(niin), fsc_code, supply_chain, item_description, unit_of_issue, latest_mlc_price, $Batch FROM ta WHERE niin IS NOT NULL ORDER BY fb_line;
"@
        }
        Invoke-SqliteExec -Database $Database -Sql 'COMMIT;'
    } catch { try { Invoke-SqliteExec -Database $Database -Sql 'ROLLBACK;' } catch { }; throw }
    Add-FlisIssue $Run INFO $t 'Merged' ("{0} month columns, {1:N0} non-blank cells: {2:N0} new month value(s), {3:N0} changed inside the window." -f $months.Count, ((Invoke-SqliteQuery -Database $Database -Sql 'SELECT COUNT(*) AS n FROM tl;')[0].n), $new, $over)
    if ($differOld) { Add-FlisIssue $Run INFO $t 'HistoryDiffers' ("{0:N0} older month value(s) in this pull differ from what is stored; stored values kept (outside the update window)." -f $differOld) }

    if (-not $isStock -and $fbCols -contains 'Total') {
        $mcols = ($months | ForEach-Object { "coalesce(CAST(replace(trim($(ConvertTo-FlisIdent $_.Col)), ',', '') AS REAL), 0)" }) -join ' + '
        $mis = (Invoke-SqliteQuery -Database $Database -Sql "SELECT NIIN AS niin, fb_line, Total AS total, ($mcols) AS s FROM $(ConvertTo-FlisIdent $fb) WHERE trim(coalesce(Total, '')) <> '' AND abs(CAST(replace(trim(Total), ',', '') AS REAL) - ($mcols)) > 0.5 LIMIT 10;")
        if ($mis.Count) { Add-FlisIssue $Run WARN $t 'TotalMismatch' ("Rows whose Total is not the sum of their month columns (first 10): {0}" -f (($mis | ForEach-Object { "line $($_.fb_line) NIIN $($_.niin) Total=$($_.total) sum=$($_.s)" }) -join '; ')) }
    }
}

# ---------------------------------------------------------------------------
# Orchestration of the transform for one ingest run
# ---------------------------------------------------------------------------

function Invoke-FlisTransform {
    param([IntPtr]$Database, $Run, [long]$Batch, [string]$RunId, $Model, $Registry,
          [datetime]$AsOfMonth, [int]$PriorMonths = 1, [switch]$StoreFutureStockMonths)
    $landed = @{}
    foreach ($r in (Invoke-SqliteQuery -Database $Database -Sql "SELECT file_stem FROM fb_load_log WHERE run_id = ? AND status = 'LANDED';" -Params @($RunId))) { $landed[$r.file_stem] = $true }
    $itemMasterRefreshed = [bool]$landed['SegmentA']
    if (-not $itemMasterRefreshed) {
        Add-FlisIssue $Run WARN 'item_niin' 'ItemMasterNotRefreshed' 'SegmentA did not land in this run, so the item master (and NIIN lifecycle) was NOT refreshed. Other tables are updated for the NIINs in their own files only.'
    }
    foreach ($tg in $Model.Targets) {
        $srcs = @($tg.Sources | Where-Object { $landed[$_] })
        if (-not $srcs.Count) {
            Add-FlisIssue $Run INFO $tg.Target 'NotRefreshed' ("Source {0} did not land in this run; {1} keeps its previous contents." -f ($tg.Sources -join ' / '), $tg.Target)
            continue
        }
        if ($srcs.Count -lt $tg.Sources.Count) {
            Add-FlisIssue $Run WARN $tg.Target 'PartialSources' ("Only {0} of {1} landed; deduplication this run covers just that file." -f ($srcs -join ', '), ($tg.Sources -join ' + '))
        }
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try {
            switch ($tg.Strategy) {
                'ItemMaster'      { Invoke-FlisUpsert -Database $Database -Run $Run -Target $tg -Batch $Batch -SourceStem $srcs[0] -ItemMaster }
                'Upsert'          { Invoke-FlisUpsert -Database $Database -Run $Run -Target $tg -Batch $Batch -SourceStem $srcs[0] }
                'Snapshot'        { Invoke-FlisSnapshot -Database $Database -Run $Run -Target $tg -Batch $Batch -SourceStem $srcs[0] -ItemMasterRefreshed $itemMasterRefreshed }
                'ReplaceAll'      { Invoke-FlisSnapshot -Database $Database -Run $Run -Target $tg -Batch $Batch -SourceStem $srcs[0] -ReplaceAll }
                'History'         { Invoke-FlisHistory -Database $Database -Run $Run -Target $tg -Batch $Batch -SourceStems $srcs }
                'RollingStock'    { Invoke-FlisRolling -Database $Database -Run $Run -Target $tg -Batch $Batch -SourceStem $srcs[0] -Registry $Registry -AsOfMonth $AsOfMonth -PriorMonths $PriorMonths -StoreFutureMonths:$StoreFutureStockMonths }
                'RollingForecast' { Invoke-FlisRolling -Database $Database -Run $Run -Target $tg -Batch $Batch -SourceStem $srcs[0] -Registry $Registry -AsOfMonth $AsOfMonth -PriorMonths $PriorMonths }
                default           { Add-FlisIssue $Run ERROR $tg.Target 'UnknownStrategy' "flis_model.psd1 strategy '$($tg.Strategy)' is not implemented." }
            }
        } catch {
            Add-FlisIssue $Run ERROR $tg.Target 'TransformFailed' ("{0} was not refreshed: {1}" -f $tg.Target, $_.Exception.Message)
        }
        Write-Verbose ("{0}: {1:N1}s" -f $tg.Target, $sw.Elapsed.TotalSeconds)
    }
    Invoke-SqliteExec -Database $Database -Sql 'DROP TABLE IF EXISTS temp.tx; DROP TABLE IF EXISTS temp.txu; DROP TABLE IF EXISTS temp.tl; DROP TABLE IF EXISTS temp.ta;'
}

# ---------------------------------------------------------------------------
# Coverage / orphans (Gitea #13)
# ---------------------------------------------------------------------------

function Add-FlisCoverage {
    param([IntPtr]$Database, $Run, [long]$Batch, [string]$Check, [string]$Table, [string]$Severity, [string]$CountSql, [string]$ExampleSql, [string]$Message)
    $n = [long](Invoke-SqliteQuery -Database $Database -Sql $CountSql)[0].n
    if (-not $n) { return }
    $ex = (@((Invoke-SqliteQuery -Database $Database -Sql $ExampleSql)) | ForEach-Object { $_.v }) -join ', '
    Invoke-SqliteNonQuery -Database $Database -Sql 'INSERT INTO coverage_report (batch_number, check_name, file_stem, severity, count, examples) VALUES (?,?,?,?,?,?);' -Params @($Batch, $Check, $Table, $Severity, $n, $ex) | Out-Null
    Add-FlisIssue $Run $Severity $Table $Check (($Message -f $n) + $(if ($ex) { " Examples: $ex" } else { '' }))
}

function Invoke-FlisCoverage {
    param([IntPtr]$Database, $Run, [long]$Batch, $Model)
    Invoke-SqliteNonQuery -Database $Database -Sql 'DELETE FROM coverage_report WHERE batch_number = ?;' -Params @($Batch) | Out-Null
    $hasItems = (Invoke-SqliteQuery -Database $Database -Sql 'SELECT COUNT(*) AS n FROM item_niin;')[0].n -gt 0
    if (-not $hasItems) { Add-FlisIssue $Run WARN '(coverage)' 'NoItemMaster' 'item_niin is empty -- coverage checks skipped (SegmentA has never landed).'; return }

    # 1. NIINs in a table but not in the item master.
    foreach ($tg in $Model.Targets) {
        if ($tg.Target -eq 'item_niin') { continue }
        $niinCol = if ($tg.Target -eq 'parts_list') { 'source_niin' } elseif ((Get-FlisDstColumns $tg) -contains 'niin' -or $tg.Strategy -like 'Rolling*') { 'niin' } else { $null }
        if (-not $niinCol) { continue }
        $c = ConvertTo-FlisIdent $niinCol
        Add-FlisCoverage -Database $Database -Run $Run -Batch $Batch -Check 'NiinNotInItemMaster' -Table $tg.Target -Severity 'WARN' `
            -CountSql "SELECT COUNT(DISTINCT $c) AS n FROM $($tg.Target) WHERE $c IS NOT NULL AND $c NOT IN (SELECT niin FROM item_niin);" `
            -ExampleSql "SELECT DISTINCT $c AS v FROM $($tg.Target) WHERE $c IS NOT NULL AND $c NOT IN (SELECT niin FROM item_niin) LIMIT 10;" `
            -Message "{0:N0} NIIN(s) in $($tg.Target) are not in the item master (SegmentA)."
    }
    # 2. Active item-master NIINs with no data in the tables the dashboard depends on.
    foreach ($pair in @(@('item_part_number_xref','part numbers'), @('item_management_data','management data'), @('contract_award_line','procurement history'), @('solicitation_line','solicitations'), @('stock_on_hand_monthly','stock on hand'), @('dla_forecast_monthly','DLA forecast'))) {
        Add-FlisCoverage -Database $Database -Run $Run -Batch $Batch -Check 'ItemWithoutData' -Table $pair[0] -Severity 'INFO' `
            -CountSql "SELECT COUNT(*) AS n FROM item_niin i WHERE i.in_latest_batch = 1 AND NOT EXISTS (SELECT 1 FROM $($pair[0]) x WHERE x.niin = i.niin);" `
            -ExampleSql "SELECT i.niin AS v FROM item_niin i WHERE i.in_latest_batch = 1 AND NOT EXISTS (SELECT 1 FROM $($pair[0]) x WHERE x.niin = i.niin) LIMIT 5;" `
            -Message "{0:N0} active item(s) have no $($pair[1]) at all (the dashboard must show these as insufficient data, not healthy)."
    }
    # 3. CAGEs referenced but missing from vendor_cage: report, then stub (is_stub = 1) from the referencing row's company name.
    $cageSources = @(@('item_part_number_xref','company'), @('contract_award_line','vendor_name'), @('item_nicn_xref','vendor'), @('spmig',$null), @('usn_p2300_next_higher_assembly',$null))
    foreach ($cs in $cageSources) {
        $tbl = $cs[0]; $nameCol = $cs[1]
        Add-FlisCoverage -Database $Database -Run $Run -Batch $Batch -Check 'CageNotInVendors' -Table $tbl -Severity 'WARN' `
            -CountSql "SELECT COUNT(DISTINCT cage_code) AS n FROM $tbl WHERE cage_code IS NOT NULL AND cage_code NOT IN (SELECT cage_code FROM vendor_cage);" `
            -ExampleSql "SELECT DISTINCT cage_code AS v FROM $tbl WHERE cage_code IS NOT NULL AND cage_code NOT IN (SELECT cage_code FROM vendor_cage) LIMIT 10;" `
            -Message "{0:N0} CAGE(s) in $tbl are not in VendorInformation; added to vendor_cage as stubs (is_stub = 1)."
        $nm = if ($nameCol) { "MIN($nameCol)" } else { 'NULL' }
        Invoke-SqliteExec -Database $Database -Sql "INSERT OR IGNORE INTO vendor_cage (cage_code, vendor_name, is_stub, first_seen_batch, last_seen_batch, source_batch) SELECT cage_code, $nm, 1, $Batch, $Batch, $Batch FROM $tbl WHERE cage_code IS NOT NULL AND cage_code NOT IN (SELECT cage_code FROM vendor_cage) GROUP BY cage_code;"
    }
    # 4. Referenced NIINs (assembly components, NHA, replacements, interchangeables) outside the item master: informational.
    foreach ($ref in @(@('parts_list','component_niin'), @('next_higher_assembly','next_higher_assembly_niin'), @('snud','replacement_niin'), @('item_interchangeability','is_niin'), @('usn_p2300_next_higher_assembly','alternate_niin'), @('usn_p2300_next_higher_assembly','related_niin'))) {
        Add-FlisCoverage -Database $Database -Run $Run -Batch $Batch -Check 'ReferencedNiinOutsideItemMaster' -Table "$($ref[0]).$($ref[1])" -Severity 'INFO' `
            -CountSql "SELECT COUNT(DISTINCT $($ref[1])) AS n FROM $($ref[0]) WHERE $($ref[1]) IS NOT NULL AND $($ref[1]) NOT IN (SELECT niin FROM item_niin);" `
            -ExampleSql "SELECT DISTINCT $($ref[1]) AS v FROM $($ref[0]) WHERE $($ref[1]) IS NOT NULL AND $($ref[1]) NOT IN (SELECT niin FROM item_niin) LIMIT 5;" `
            -Message "{0:N0} NIIN(s) referenced in $($ref[0]).$($ref[1]) are outside the item master (normal for components/replacements; listed for completeness)."
    }
}
