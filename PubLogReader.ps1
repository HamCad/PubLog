#requires -Version 5.1
<#
.SYNOPSIS
  Zero-dependency reader for PUB LOG / FED LOG ".TAB" tables (the DLA
  "IMD" monthly data files), for Windows PowerShell 5.1.

.DESCRIPTION
  Reads the raw .TAB files straight out of PublogDVD.zip (or a FED LOG
  download) -- no DECOMP.EXE, no IMD viewer, no vendor DLL, no NuGet,
  nothing to install. Like SQLiteInterop.ps1, it uses Add-Type to compile a
  small C# class with the .NET compiler already on every Windows machine;
  the hot decode loop has to be compiled code, because a pure-script byte
  loop over ~3.7 GB / ~230 million rows would take days.

  Two pieces of the format matter for anyone maintaining this:
    * The first 8 KB of every .TAB is an INI-style header ([FileInformation]
      + one [ColumnNNN] section per column) encrypted with AES-256/ECB under
      a fixed key. It is decrypted with System.Security.Cryptography, which
      ships with .NET Framework.
    * Everything after the header is NOT encrypted, only compressed:
      a word list (when FrequentWords=Yes), a dictionary of distinct
      "value tuples" (every non-key column joined with '|'), keyed data
      blocks (BlockSize from the header, usually 8 KB) that pair each key with a pointer into that dictionary, a sparse
      key index, and a 16-byte trailer [dataOffset, dataLength, indexOffset,
      indexLength]. Format notes: docs/PUBLOG_TAB_FORMAT.md.

  Output is byte-for-byte identical to DECOMP.EXE's "select * from T where
  KEY='*'" for every table this was validated against (see
  Test-PubLogDecoder.ps1).

.EXAMPLE
  . .\PubLogReader.ps1
  Get-PubLogTableInfo -Path D:\publog\V_FLIS_PART.TAB
  Export-PubLogTable -Path D:\publog\V_H2_FSC.TAB -OutFile .\V_H2_FSC.txt
  Read-PubLogTable -Path D:\publog\V_H2_FSG.TAB -First 5
#>

if (-not ('PubLog.ImdTable' -as [type])) {
Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace PubLog
{
    public class ImdColumn
    {
        public string Name;
        public Dictionary<string, string> Properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public override string ToString() { return Name; }
    }

    // Decoded dictionary entries packed into large byte chunks (an entry
    // never straddles two chunks), addressed by the file offset the data
    // blocks point at. Avoids millions of small .NET strings for tables
    // like V_CAGE_ADDRESS whose dictionary expands to hundreds of MB.
    internal class ValueStore
    {
        const int ChunkSize = 1 << 26;   // 64 MB
        List<byte[]> chunks = new List<byte[]>();
        int used = ChunkSize;
        List<int> offsets = new List<int>();
        List<long> locs = new List<long>();      // (chunk << 32) | start
        List<int> lens = new List<int>();
        int[] offArr;

        public void Add(int fileOffset, byte[] buf, int len)
        {
            if (used + len > ChunkSize) { chunks.Add(new byte[Math.Max(ChunkSize, len)]); used = 0; }
            int c = chunks.Count - 1;
            Buffer.BlockCopy(buf, 0, chunks[c], used, len);
            offsets.Add(fileOffset); locs.Add(((long)c << 32) | (uint)used); lens.Add(len);
            used += len;
        }
        public void Seal() { offArr = offsets.ToArray(); offsets = null; }
        public int Count { get { return offArr != null ? offArr.Length : offsets.Count; } }

        public bool TryGet(int fileOffset, out byte[] chunk, out int start, out int len)
        {
            int k = Array.BinarySearch(offArr, fileOffset);
            if (k < 0) { chunk = null; start = 0; len = 0; return false; }
            long loc = locs[k];
            chunk = chunks[(int)(loc >> 32)]; start = (int)(loc & 0xffffffff); len = lens[k];
            return true;
        }
    }

    public class ImdTable : IDisposable
    {
        static readonly byte[] HeaderKey = Encoding.ASCII.GetBytes("Kab1r__M1ke__RicKab1r__M1ke__Ric");
        const int HeaderDecryptBytes = 8192;   // the INI text always fits in the first 8 KB
        int Block = 8192;                      // per table: BlockSize= in the header (P_CHARACTERISTICS_PICK uses 10240)

        public string Path;
        public Dictionary<string, string> Info = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public List<ImdColumn> Columns = new List<ImdColumn>();
        public string HeaderText;
        public int DataOffset, DataLength, IndexOffset, IndexLength;

        byte[] d;              // whole file; largest PUB LOG table is ~550 MB
        byte[][] words;
        int wordHi;            // word-token high bytes are 0..wordHi-1 (0 = no word list)
        int valueStart;

        public string TableName { get { return Get("TableName"); } }
        public string FileType { get { return Get("FileType"); } }
        public string PrimaryKey { get { return Get("PrimaryKey"); } }
        public string PubDate { get { return Get("Date"); } }
        public long DeclaredRows { get { long n; return long.TryParse(Get("Rows"), out n) ? n : -1; } }
        public bool IsPicklist { get { return string.Equals(FileType, "Picklist", StringComparison.OrdinalIgnoreCase); } }
        public string[] ColumnNames
        {
            get { var a = new string[Columns.Count]; for (int i = 0; i < a.Length; i++) a[i] = Columns[i].Name; return a; }
        }
        string Get(string k) { string v; return Info.TryGetValue(k, out v) ? v : null; }

        public static ImdTable Open(string path) { return FromBytes(File.ReadAllBytes(path), path); }

        // For reading straight out of PublogDVD.zip without extracting it.
        public static ImdTable FromBytes(byte[] bytes, string displayName)
        {
            var t = new ImdTable();
            t.Path = displayName;
            t.d = bytes;
            t.ParseHeader();
            t.ParseTrailer();
            t.ParseWords();
            return t;
        }
        // Header only (first 8 KB): table name, pub date, rows, columns.
        // Cheap enough to call on every table to decide what needs loading.
        public static ImdTable HeaderOnly(byte[] firstBlock, string displayName)
        {
            var t = new ImdTable();
            t.Path = displayName;
            t.d = firstBlock;
            t.ParseHeader();
            t.d = null;
            return t;
        }
        public void Dispose() { d = null; words = null; }

        public static string DecryptHeader(byte[] fileBytes)
        {
            using (var aes = new AesManaged())
            {
                aes.Key = HeaderKey; aes.Mode = CipherMode.ECB; aes.Padding = PaddingMode.None;
                using (var dec = aes.CreateDecryptor())
                {
                    byte[] plain = dec.TransformFinalBlock(fileBytes, 0, HeaderDecryptBytes);
                    int n = Array.IndexOf(plain, (byte)0); if (n < 0) n = plain.Length;
                    return Encoding.GetEncoding(28591).GetString(plain, 0, n);
                }
            }
        }

        void ParseHeader()
        {
            HeaderText = DecryptHeader(d);
            if (!HeaderText.StartsWith("[FileInformation]"))
                throw new InvalidDataException(Path + ": header did not decrypt to [FileInformation] -- not an IMD .TAB file, or a format change.");
            Dictionary<string, string> cur = null;
            foreach (string raw in HeaderText.Split('\n'))
            {
                string line = raw.Trim();
                if (line.StartsWith("["))
                {
                    if (line.StartsWith("[Column", StringComparison.OrdinalIgnoreCase)) { var c = new ImdColumn(); Columns.Add(c); cur = c.Properties; }
                    else cur = Info;
                    continue;
                }
                int eq = line.IndexOf('=');
                if (eq > 0 && cur != null) cur[line.Substring(0, eq)] = line.Substring(eq + 1);
            }
            foreach (var c in Columns) { string n; c.Name = c.Properties.TryGetValue("Name", out n) ? n : "?"; }
            int bs; if (int.TryParse(Get("BlockSize"), out bs) && bs >= 4096 && bs % 16 == 0) Block = bs;
        }

        void ParseTrailer()
        {
            int p = d.Length - 16;
            DataOffset = BitConverter.ToInt32(d, p); DataLength = BitConverter.ToInt32(d, p + 4);
            IndexOffset = BitConverter.ToInt32(d, p + 8); IndexLength = BitConverter.ToInt32(d, p + 12);
        }

        void ParseWords()
        {
            valueStart = Block;
            if (!string.Equals(Get("FrequentWords"), "Yes", StringComparison.OrdinalIgnoreCase)) return;
            int start = Block, end = -1;
            for (int i = start; i + 3 < d.Length; i++)
                if (d[i] == 13 && d[i + 1] == 10 && d[i + 2] == 13 && d[i + 3] == 10) { end = i; break; }
            if (end < 0) throw new InvalidDataException(Path + ": FrequentWords=Yes but no word list terminator found.");
            var list = new List<byte[]>();
            int s = start;
            for (int i = start; i <= end; i++)
                if (i == end || (d[i] == 13 && d[i + 1] == 10)) { var w = new byte[i - s]; Buffer.BlockCopy(d, s, w, 0, w.Length); list.Add(w); s = i + 2; i++; }
            words = list.ToArray();
            wordHi = (words.Length + 255) / 256;
            valueStart = ((end + 4 + Block - 1) / Block) * Block;
        }

        // ---- the string codec ------------------------------------------------
        // Byte c, low 7 bits v, high bit h:
        //   v <  wordHi        word token: words[(v<<8)|next], plus ' ' if h
        //   0x1C <= v <= 0x1F  little-endian integer of (v-0x1B) bytes, written
        //                      in decimal, plus ' ' if h
        //   v <  0x1C          copy v bytes from the previous string, at the
        //                      same position; ends the string if h
        //   otherwise          literal character v; ends the string if h
        int ReadString(int i, byte[] prev, int prevLen, ref byte[] buf, out int len)
        {
            len = 0;
            while (true)
            {
                int c = d[i++];
                int v = c & 0x7f;
                bool hi = (c & 0x80) != 0;
                if (v < wordHi)
                {
                    byte[] w = words[(v << 8) | d[i++]];
                    Ensure(ref buf, len + w.Length + 1);
                    Buffer.BlockCopy(w, 0, buf, len, w.Length); len += w.Length;
                    if (hi) buf[len++] = 32;
                    continue;
                }
                if (v >= 0x1C && v <= 0x1F)
                {
                    int n = v - 0x1B; long num = 0;
                    for (int k = 0; k < n; k++) num |= (long)d[i + k] << (8 * k);
                    i += n;
                    string s = num.ToString();
                    Ensure(ref buf, len + s.Length + 1);
                    for (int k = 0; k < s.Length; k++) buf[len++] = (byte)s[k];
                    if (hi) buf[len++] = 32;
                    continue;
                }
                if (v < 0x1C)
                {
                    int take = Math.Max(0, Math.Min(v, prevLen - len));
                    Ensure(ref buf, len + take);
                    if (take > 0) Buffer.BlockCopy(prev, len, buf, len, take);
                    len += take;
                }
                else { Ensure(ref buf, len + 1); buf[len++] = (byte)v; }
                if (hi) return i;
            }
        }

        // Keys: a lone 0x80..0x9F byte n means "same as previous key" when n
        // equals the previous key's length, else "previous key + n" (added to
        // its trailing digit run, or to the last character when it has none).
        int ReadKey(int i, byte[] prev, int prevLen, ref byte[] buf, out int len)
        {
            int c = d[i];
            if (c >= 0x80 && c < 0xA0 && prevLen > 0)
            {
                int n = c & 0x7f;
                Ensure(ref buf, prevLen + 4);
                if (n == prevLen) { Buffer.BlockCopy(prev, 0, buf, 0, prevLen); len = prevLen; return i + 1; }
                int j = prevLen;
                while (j > 0 && prev[j - 1] >= 48 && prev[j - 1] <= 57) j--;
                Buffer.BlockCopy(prev, 0, buf, 0, j);
                if (j == prevLen) { buf[prevLen - 1] = (byte)(prev[prevLen - 1] + n); len = prevLen; return i + 1; }
                long tail = 0;
                for (int k = j; k < prevLen; k++) tail = tail * 10 + (prev[k] - 48);
                string s = (tail + n).ToString().PadLeft(prevLen - j, '0');
                Ensure(ref buf, j + s.Length);
                for (int k = 0; k < s.Length; k++) buf[j + k] = (byte)s[k];
                len = j + s.Length;
                return i + 1;
            }
            return ReadString(i, prev, prevLen, ref buf, out len);
        }

        static void Ensure(ref byte[] buf, int need)
        {
            if (need > buf.Length) { var nb = new byte[Math.Max(need, buf.Length * 2)]; Buffer.BlockCopy(buf, 0, nb, 0, buf.Length); buf = nb; }
        }

        // A block ends at the first record/entry boundary where every byte
        // left in the block is zero (padding). A lone zero byte is NOT an end
        // marker: in data blocks it introduces a multi-row group, and in
        // FrequentWords dictionaries 00 xx is a word token.
        bool RestIsZero(int i, int end)
        {
            for (int k = i; k < end; k++) if (d[k] != 0) return false;
            return true;
        }

        // Callback receives (buffer, length) for each decoded string, in file
        // order; the buffer is reused between calls.
        public delegate void RowSink(byte[] buf, int len);

        void ForEachDictEntry(Action<int, byte[], int> sink)
        {
            int stop = DataOffset;
            if (stop <= 0 || stop > d.Length - 16) stop = d.Length - 16;
            byte[] cur = new byte[256], prev = new byte[256];
            for (int blk = valueStart; blk < stop; blk += Block)
            {
                int end = Math.Min(blk + Block, d.Length);
                int i = blk, prevLen = 0, len;
                while (i < end && !(d[i] == 0 && RestIsZero(i, end)))
                {
                    int off = i;
                    i = ReadString(i, prev, prevLen, ref cur, out len);
                    sink(off, cur, len);
                    var t = prev; prev = cur; cur = t; prevLen = len;
                    if (cur.Length < prev.Length) cur = new byte[prev.Length];
                }
            }
        }

        // Streams every row as one '|'-joined byte string (no header line),
        // exactly as DECOMP.EXE writes it. Returns the row count.
        public long ReadRows(RowSink sink)
        {
            long rows = 0;
            if (IsPicklist)
            {
                // Picklist "dictionary" blocks already hold complete rows in key
                // order; the data blocks only map keys to those blocks.
                ForEachDictEntry((off, b, n) => { sink(b, n); rows++; });
                return rows;
            }

            var store = new ValueStore();
            ForEachDictEntry((off, b, n) => store.Add(off, b, n));
            store.Seal();

            int plen = int.Parse(Get("DataLength"));
            byte[] key = new byte[64], prevKey = new byte[64], row = new byte[1024];
            for (int blk = DataOffset; blk < DataOffset + DataLength; blk += Block)
            {
                int end = Math.Min(blk + Block, d.Length);
                int i = blk, prevLen = 0, klen;
                while (i < end && !(d[i] == 0 && RestIsZero(i, end)))
                {
                    int cnt = 1;
                    if (d[i] == 0) { cnt = d[i + 1] + 1; i += 2; }
                    i = ReadKey(i, prevKey, prevLen, ref key, out klen);
                    var t = prevKey; prevKey = key; key = t; prevLen = klen;
                    if (key.Length < prevKey.Length) key = new byte[prevKey.Length];
                    for (int r = 0; r < cnt; r++)
                    {
                        int ptr = 0;
                        for (int k = 0; k < plen; k++) ptr |= d[i + k] << (8 * k);
                        i += plen;
                        byte[] chunk; int vs, vl;
                        if (!store.TryGet(ptr, out chunk, out vs, out vl))
                            throw new InvalidDataException(string.Format("{0}: record at 0x{1:X} points to 0x{2:X}, which is not a dictionary entry.", Path, i - plen, ptr));
                        Ensure(ref row, klen + 1 + vl);
                        Buffer.BlockCopy(prevKey, 0, row, 0, klen);
                        row[klen] = (byte)'|';
                        Buffer.BlockCopy(chunk, vs, row, klen + 1, vl);
                        sink(row, klen + 1 + vl);
                        rows++;
                    }
                }
            }
            return rows;
        }

        // Writes DECOMP-identical output: header line, then rows, CRLF endings.
        public long ExportPipe(string outFile)
        {
            using (var fs = new FileStream(outFile, FileMode.Create, FileAccess.Write, FileShare.Read, 1 << 20))
            {
                byte[] hdr = Encoding.GetEncoding(28591).GetBytes(Get("Header") + "\r\n");
                fs.Write(hdr, 0, hdr.Length);
                byte[] crlf = new byte[] { 13, 10 };
                return ReadRows((b, n) => { fs.Write(b, 0, n); fs.Write(crlf, 0, 2); });
            }
        }

        // SHA-256 of exactly what ExportPipe would write, without writing it.
        // Comparable with Get-FileHash of a DECOMP.EXE dump of the same table.
        public string HashPipe(out long rows)
        {
            using (var sha = SHA256.Create())
            {
                byte[] hdr = Encoding.GetEncoding(28591).GetBytes(Get("Header") + "\r\n");
                sha.TransformBlock(hdr, 0, hdr.Length, null, 0);
                byte[] crlf = new byte[] { 13, 10 };
                rows = ReadRows((b, n) => { sha.TransformBlock(b, 0, n, null, 0); sha.TransformBlock(crlf, 0, 2, null, 0); });
                sha.TransformFinalBlock(new byte[0], 0, 0);
                return BitConverter.ToString(sha.Hash).Replace("-", "");
            }
        }

        // Convenience for small tables / spot checks: rows as string arrays.
        public List<string[]> ReadAll(int max)
        {
            var list = new List<string[]>();
            var enc = Encoding.GetEncoding(28591);
            try
            {
                ReadRows((b, n) =>
                {
                    if (max > 0 && list.Count >= max) throw new StopIteration();
                    list.Add(enc.GetString(b, 0, n).Split('|'));
                });
            }
            catch (StopIteration) { }
            return list;
        }
        class StopIteration : Exception { }
    }

    // Bulk loader into SQLite through winsqlite3.dll (built into Windows).
    // Takes a db handle from SQLiteInterop.ps1's Open-SqliteDb -- same DLL,
    // same process, so the handle is shared directly.
    public static class SqliteBulk
    {
        const string DLL = "winsqlite3.dll";
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_prepare_v2(IntPtr db, byte[] sql, int nBytes, out IntPtr stmt, IntPtr tail);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_bind_text(IntPtr stmt, int index, byte[] value, int nBytes, IntPtr destructor);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_bind_null(IntPtr stmt, int index);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_step(IntPtr stmt);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_reset(IntPtr stmt);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_finalize(IntPtr stmt);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern int sqlite3_exec(IntPtr db, byte[] sql, IntPtr cb, IntPtr arg, out IntPtr err);
        [DllImport(DLL, CallingConvention = CallingConvention.Cdecl)] static extern IntPtr sqlite3_errmsg(IntPtr db);
        static readonly IntPtr SQLITE_TRANSIENT = new IntPtr(-1);
        const int SQLITE_DONE = 101;

        static byte[] Z(string s) { var b = Encoding.UTF8.GetBytes(s); var z = new byte[b.Length + 1]; Buffer.BlockCopy(b, 0, z, 0, b.Length); return z; }
        static string Err(IntPtr db)
        {
            IntPtr p = sqlite3_errmsg(db); if (p == IntPtr.Zero) return "?";
            int n = 0; while (Marshal.ReadByte(p, n) != 0) n++;
            var b = new byte[n]; Marshal.Copy(p, b, 0, n); return Encoding.UTF8.GetString(b);
        }
        public static void Exec(IntPtr db, string sql)
        {
            IntPtr err;
            if (sqlite3_exec(db, Z(sql), IntPtr.Zero, IntPtr.Zero, out err) != 0) throw new Exception("sqlite: " + Err(db) + " in: " + sql);
        }

        // CHANGES\NIIN.zip / CAGE.zip members: one key per line, CAGE lines
        // followed by associated CAGEs. Millions of lines, hence compiled.
        public static long LoadChangeList(IntPtr db, byte[] content, string pubDate, string keyType, string changeType)
        {
            IntPtr stmt;
            if (sqlite3_prepare_v2(db, Z("INSERT INTO pl_change_log (pub_date, key_type, change_type, key_value, related) VALUES (?,?,?,?,?)"), -1, out stmt, IntPtr.Zero) != 0)
                throw new Exception("prepare failed: " + Err(db));
            byte[] p = Z(pubDate), k = Z(keyType), c = Z(changeType);
            long n = 0;
            try
            {
                int i = 0, len = content.Length;
                while (i < len)
                {
                    int s = i; while (i < len && content[i] != 10) i++;
                    int e = i; i++;
                    while (e > s && (content[e - 1] == 13 || content[e - 1] == 32)) e--;
                    while (s < e && content[s] == 32) s++;
                    if (e <= s) continue;
                    int sp = s; while (sp < e && content[sp] != 32) sp++;
                    sqlite3_bind_text(stmt, 1, p, p.Length - 1, SQLITE_TRANSIENT);
                    sqlite3_bind_text(stmt, 2, k, k.Length - 1, SQLITE_TRANSIENT);
                    sqlite3_bind_text(stmt, 3, c, c.Length - 1, SQLITE_TRANSIENT);
                    var key = new byte[sp - s]; Buffer.BlockCopy(content, s, key, 0, key.Length);
                    sqlite3_bind_text(stmt, 4, key, key.Length, SQLITE_TRANSIENT);
                    int rs = sp; while (rs < e && content[rs] == 32) rs++;
                    if (rs < e) { var rel = new byte[e - rs]; Buffer.BlockCopy(content, rs, rel, 0, rel.Length); sqlite3_bind_text(stmt, 5, rel, rel.Length, SQLITE_TRANSIENT); }
                    else sqlite3_bind_null(stmt, 5);
                    if (sqlite3_step(stmt) != SQLITE_DONE) throw new Exception("insert failed: " + Err(db));
                    sqlite3_reset(stmt);
                    n++;
                }
            }
            finally { sqlite3_finalize(stmt); }
            return n;
        }

        // Inserts every row of the table into sqlTable (which must already
        // exist with one column per IMD column, in physical order). Empty
        // fields become NULL. Commits every batchRows rows. Returns row count.
        public static long Load(IntPtr db, ImdTable t, string sqlTable, int batchRows)
        {
            int ncol = t.Columns.Count;
            var ph = new StringBuilder();
            for (int c = 0; c < ncol; c++) ph.Append(c == 0 ? "?" : ",?");
            IntPtr stmt;
            if (sqlite3_prepare_v2(db, Z("INSERT INTO \"" + sqlTable + "\" VALUES (" + ph + ")"), -1, out stmt, IntPtr.Zero) != 0)
                throw new Exception("prepare failed: " + Err(db));
            long n = 0;
            byte[] field = new byte[4096];
            Exec(db, "BEGIN");
            try
            {
                t.ReadRows((b, len) =>
                {
                    int col = 0, start = 0;
                    for (int i = 0; i <= len && col < ncol; i++)
                    {
                        if (i == len || b[i] == (byte)'|')
                        {
                            int flen = i - start;
                            if (flen == 0) sqlite3_bind_null(stmt, col + 1);
                            else
                            {
                                // Latin-1 -> UTF-8 (FLIS text is almost entirely ASCII).
                                int need = flen * 2; if (field.Length < need) field = new byte[need * 2];
                                int o = 0;
                                for (int k = start; k < i; k++)
                                {
                                    byte x = b[k];
                                    if (x < 0x80) field[o++] = x;
                                    else { field[o++] = (byte)(0xC0 | (x >> 6)); field[o++] = (byte)(0x80 | (x & 0x3F)); }
                                }
                                sqlite3_bind_text(stmt, col + 1, field, o, SQLITE_TRANSIENT);
                            }
                            col++; start = i + 1;
                        }
                    }
                    for (; col < ncol; col++) sqlite3_bind_null(stmt, col + 1);
                    if (sqlite3_step(stmt) != SQLITE_DONE) throw new Exception("insert failed at row " + (n + 1) + ": " + Err(db));
                    sqlite3_reset(stmt);
                    n++;
                    if (batchRows > 0 && n % batchRows == 0) { Exec(db, "COMMIT"); Exec(db, "BEGIN"); }
                });
                Exec(db, "COMMIT");
            }
            catch { try { Exec(db, "ROLLBACK"); } catch { } throw; }
            finally { sqlite3_finalize(stmt); }
            return n;
        }
    }
}
"@
}

function Get-PubLogTableInfo {
    # Header facts for one .TAB (or every .TAB in a folder): table name,
    # declared rows, publication date, physical column order.
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $files = if (Test-Path -LiteralPath $Path -PathType Container) {
        Get-ChildItem -LiteralPath $Path -Filter *.TAB | Sort-Object Name
    } else { Get-Item -LiteralPath $Path }
    foreach ($f in $files) {
        $t = [PubLog.ImdTable]::Open($f.FullName)
        try {
            [PSCustomObject]@{
                Table      = $t.TableName
                FileType   = $t.FileType
                PrimaryKey = $t.PrimaryKey
                Rows       = $t.DeclaredRows
                PubDate    = $t.PubDate
                Columns    = $t.ColumnNames -join '|'
                SizeMB     = [math]::Round($f.Length / 1MB, 1)
                Path       = $f.FullName
            }
        } finally { $t.Dispose() }
    }
}

function Export-PubLogTable {
    # Decodes a .TAB to the same pipe-delimited text DECOMP.EXE produces.
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$OutFile
    )
    $t = [PubLog.ImdTable]::Open((Resolve-Path -LiteralPath $Path).ProviderPath)
    try {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $outFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)
        $n = $t.ExportPipe($outFull)
        if ($t.DeclaredRows -ge 0 -and $n -ne $t.DeclaredRows) {
            Write-Warning "$($t.TableName): decoded $n rows but the header declares $($t.DeclaredRows)."
        }
        Write-Verbose ("{0}: {1:N0} rows in {2:N1}s" -f $t.TableName, $n, $sw.Elapsed.TotalSeconds)
        $n
    } finally { $t.Dispose() }
}

function Read-PubLogTable {
    # Small tables / spot checks only: returns PSCustomObjects. For bulk work
    # use Export-PubLogTable or Import-PubLogTable (PubLogEtl.ps1).
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$First = 0
    )
    $t = [PubLog.ImdTable]::Open((Resolve-Path -LiteralPath $Path).ProviderPath)
    try {
        $names = $t.ColumnNames
        foreach ($r in $t.ReadAll($First)) {
            $o = [ordered]@{}
            for ($c = 0; $c -lt $names.Count; $c++) { $o[$names[$c]] = if ($c -lt $r.Count) { $r[$c] } else { '' } }
            [PSCustomObject]$o
        }
    } finally { $t.Dispose() }
}
