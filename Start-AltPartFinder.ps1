#requires -Version 5.1
<#
.SYNOPSIS
  Alternate Part Finder: a local drill-down page over PUB LOG characteristics.

.DESCRIPTION
  Start from a NIIN, NSN, part number (or CAGE:PN) or an item name. The page
  lists the item's characteristics; lock the ones that must match, give
  ranges for the ones that may vary (length, thickness, ...), and it lists
  the other items with the same item name that satisfy them, ranked by how
  closely they match the starting item. It also shows FLIS's own
  relationships (standardization decisions, cancelled/replacement NIINs,
  phrases, shared reference numbers) and a data-quality view of the
  characteristics being compared.

  Runs entirely on this machine: PowerShell 5.1, winsqlite3.dll and a small
  HTTP listener bound to 127.0.0.1 (no admin rights, nothing installed,
  nothing reachable from the network). Stop it with Ctrl+C.

  Needs publog.db (Update-PubLogDatabase.ps1). Builds publog_alt.db on first
  run (~20-30 minutes for a full PUB LOG cut) and warns when it is older than
  publog.db.

.EXAMPLE
  .\Start-AltPartFinder.ps1
.EXAMPLE
  .\Start-AltPartFinder.ps1 -PubLogDatabase D:\publog\publog.db -Port 8800 -NoBrowser
#>
[CmdletBinding()]
param(
    [string]$PubLogDatabase = (Join-Path $PSScriptRoot 'publog.db'),
    [string]$AltDatabase,
    [int]$Port = 8765,
    [switch]$NoBrowser,
    [switch]$NoBuild,
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'PubLogAlternates.ps1')

$pl = Get-AltFullPath $PubLogDatabase
if (-not (Test-Path -LiteralPath $pl)) { throw "PUB LOG database not found: $pl. Run Update-PubLogDatabase.ps1 first, or pass -PubLogDatabase." }
if (-not $AltDatabase) { $AltDatabase = Join-Path (Split-Path $pl -Parent) 'publog_alt.db' }
$alt = Get-AltFullPath $AltDatabase
$html = Join-Path $PSScriptRoot 'altparts.html'
if (-not (Test-Path -LiteralPath $html)) { throw "altparts.html is missing next to this script ($html)." }

if (-not (Test-Path -LiteralPath $alt)) {
    if ($NoBuild) { throw "No alternate-part index at $alt. Run without -NoBuild (or Build-AltPartIndex) to create it." }
    Write-Host "No alternate-part index yet; building it now (one-off, roughly 20-30 minutes for a full PUB LOG cut)." -ForegroundColor Yellow
    Build-AltPartIndex -PubLogDatabase $pl -AltDatabase $alt
} else {
    Build-AltPartIndex -PubLogDatabase $pl -AltDatabase $alt -WhatIfStale
}

$engine = New-Object AltParts.Engine $alt, $pl
$server = $null
for ($p = $Port; $p -lt $Port + 20; $p++) {
    try { $server = New-Object AltParts.Server $engine, $html, $p; break }
    catch [System.Net.Sockets.SocketException] { continue }
    catch { if ($_.Exception.InnerException -is [System.Net.Sockets.SocketException]) { continue } else { throw } }
}
if (-not $server) { $engine.Dispose(); throw "No free port between $Port and $($Port + 19)." }
if (-not $Quiet) { $server.Log = [Action[string]] { param($m) Write-Host $m -ForegroundColor DarkGray } }

Write-Host ("Alternate Part Finder: {0}   (Ctrl+C to stop)" -f $server.Url) -ForegroundColor Green
Write-Host ("  index {0}`n  publog {1}" -f $alt, $pl)
if (-not $NoBrowser) { Start-Process $server.Url }
try {
    while ($true) { [void]$server.ServeOne(250) }
}
finally {
    $server.Dispose()
    $engine.Dispose()
    Write-Host 'Alternate Part Finder stopped.'
}
