# node-pruefen.ps1 - prüft Node.js, npm/npx, winget, Zertifikat und Add-in-Server
$ErrorActionPreference = "SilentlyContinue"
function Ok($t)   { Write-Host "  [OK]    $t" -ForegroundColor Green }
function Warn($t) { Write-Host "  [FEHLT] $t" -ForegroundColor Yellow }
function Info($t) { Write-Host "          $t" -ForegroundColor Gray }

Write-Host "`n=== Prüfung für das Claude Mail Add-in ===`n" -ForegroundColor Cyan

# 1. Node.js im PATH
Write-Host "1. Node.js im PATH"
$node = Get-Command node -ErrorAction SilentlyContinue
if ($node) { Ok "node gefunden: $($node.Source)"; Info "Version: $(& node -v)" }
else       { Warn "node ist nicht im PATH" }

# 2. Standard-Installationsordner
Write-Host "`n2. Installationsordner"
$dirs = @("$env:ProgramFiles\nodejs", "${env:ProgramFiles(x86)}\nodejs", "$env:LOCALAPPDATA\Programs\nodejs")
$found = @($dirs | Where-Object { Test-Path (Join-Path $_ "node.exe") })
if ($found.Count) { foreach ($d in $found) { Ok "node.exe in $d"; Info "Version: $(& (Join-Path $d 'node.exe') -v)" } }
else        { Warn "kein node.exe in den Standardordnern" }
if ($found.Count -and -not $node) { Info "Hinweis: Node ist installiert, aber dieses Fenster kennt den PATH noch nicht -> PowerShell neu öffnen." }

# 3. npm / npx
Write-Host "`n3. npm und npx"
foreach ($cmd in "npm", "npx") {
  $c = Get-Command "$cmd.cmd" -ErrorAction SilentlyContinue
  if (-not $c -and $found.Count) { $p = Join-Path $found[0] "$cmd.cmd"; if (Test-Path $p) { $c = @{ Source = $p } } }
  if ($c) { Ok "$cmd : $($c.Source)"; Info "Version: $(& $c.Source -v)" } else { Warn "$cmd nicht gefunden" }
}

# 4. winget
Write-Host "`n4. winget (für automatische Installation)"
$wg = Get-Command winget -ErrorAction SilentlyContinue
if ($wg) { Ok "winget vorhanden: $(& winget --version)" } else { Warn "winget fehlt -> Node.js LTS manuell von https://nodejs.org installieren" }

# 5. localhost-Zertifikat
Write-Host "`n5. localhost-Zertifikat"
$certDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
if ((Test-Path "$certDir\localhost.crt") -and (Test-Path "$certDir\localhost.key")) { Ok "Zertifikatsdateien in $certDir" }
else { Warn "Zertifikatsdateien fehlen -> einmalig-zertifikat-installieren.cmd ausführen" }
$trusted = Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like "*Developer CA for Microsoft Office Add-ins*" }
if ($trusted) { Ok "Entwickler-Zertifikat ist als vertrauenswürdig eingetragen" } else { Warn "Entwickler-Zertifikat nicht im Stammzertifikatspeicher" }

# 6. Add-in-Server
Write-Host "`n6. Add-in-Server auf Port 3000"
$listen = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue
if ($listen) { Ok "Server läuft (https://localhost:3000)" } else { Warn "Server läuft nicht -> start-addin.cmd doppelklicken" }

Write-Host "`n=== Fertig ===`n" -ForegroundColor Cyan
