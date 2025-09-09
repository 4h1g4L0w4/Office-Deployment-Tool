<# 
installer.ps1 (usa enlace directo a ODT)
#>

# ---------- Config ----------
$WorkDir        = "C:\ODT"
$TimeoutMin     = 20
$OdtDirectUrl   = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19029-20136.exe"
$OdtExePath     = Join-Path $WorkDir "officedeploymenttool.exe"
$SetupExePath   = Join-Path $WorkDir "setup.exe"
$DestXml        = Join-Path $WorkDir "configuration.xml"
$OctUrl         = "https://config.office.com/deploymentsettings"
$Downloads      = Join-Path $env:USERPROFILE "Downloads"

$Browsers = @(
  "$env:ProgramFiles(x86)\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
  "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
  "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
)

# ---------- Helpers ----------
function Ensure-Folder($p) { if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p | Out-Null } }

function Enable-Tls12 { try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {} }

function Download-File($url, $dest) {
  Write-Host "Descargando: $url" -ForegroundColor Cyan
  try {
    # BITS primero
    Start-BitsTransfer -Source $url -Destination $dest -ErrorAction Stop
  } catch {
    Write-Host "BITS falló, probando con Invoke-WebRequest…" -ForegroundColor Yellow
    Invoke-WebRequest -Uri $url -OutFile $dest -MaximumRedirection 10 -UseBasicParsing -ErrorAction Stop
  }
  if (-not (Test-Path -LiteralPath $dest)) { throw "No se pudo descargar $url" }
}

function Open-Url($url) {
  $browser = $Browsers | Where-Object { Test-Path $_ } | Select-Object -First 1
  if ($browser) { Start-Process -FilePath $browser -ArgumentList $url | Out-Null }
  else { Start-Process $url | Out-Null }
}

# ---------- Flujo ----------
Enable-Tls12
Ensure-Folder $WorkDir

# 1) Descargar ODT desde TU enlace directo
if (-not (Test-Path -LiteralPath $OdtExePath)) {
  Download-File -url $OdtDirectUrl -dest $OdtExePath
} else {
  Write-Host "ODT ya existe: $OdtExePath" -ForegroundColor DarkGray
}

# 2) Extraer ODT
Write-Host "Extrayendo ODT en $WorkDir…" -ForegroundColor Cyan
& $OdtExePath /quiet /extract:$WorkDir
Start-Sleep -Seconds 2
if (-not (Test-Path -LiteralPath $SetupExePath)) { throw "No se encontró setup.exe tras extraer ODT." }

# 3) Abrir el configurador para que generes tu XML
Write-Host "Abriendo Office Customization Tool…" -ForegroundColor Cyan
Open-Url $OctUrl
Write-Host "Descargá el XML (Export) y guardalo. Voy a monitorear tu carpeta: $Downloads" -ForegroundColor Yellow

# 4) Esperar un XML NUEVO en Descargas y moverlo a C:\ODT\configuration.xml
$startTime = Get-Date
$deadline  = $startTime.AddMinutes($TimeoutMin)

function Get-NewXml {
  $files = Get-ChildItem -Path $Downloads -Filter *.xml -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime
  if (-not $files) { return $null }
  $latest = $files[-1]
  if ($latest.LastWriteTime -ge $startTime.AddSeconds(-5)) { return $latest } # margen por reloj
  return $null
}

$newXml = $null
while (-not $newXml -and (Get-Date) -lt $deadline) {
  Start-Sleep -Seconds 3
  $newXml = Get-NewXml
}

if (-not $newXml) { throw "No encontré un XML nuevo en $Downloads dentro de $TimeoutMin minutos." }

Write-Host "XML detectado: $($newXml.FullName)" -ForegroundColor Green
Copy-Item -LiteralPath $newXml.FullName -Destination $DestXml -Force

# 5) Instalar con ODT
Write-Host "Instalando Office con: `"$SetupExePath`" /configure `"$DestXml`"" -ForegroundColor Cyan
Push-Location $WorkDir
& $SetupExePath /configure $DestXml
$code = $LASTEXITCODE
Pop-Location

if ($code -eq 0) {
  Write-Host "Instalación finalizada OK." -ForegroundColor Green
} else {
  Write-Host "Instalación finalizada con código $code. Revisá logs en %temp% y el XML." -ForegroundColor Yellow
}
