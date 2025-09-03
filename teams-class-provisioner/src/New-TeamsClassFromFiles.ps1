param(
  [int]$ChannelAddDelaySec = 8,
  [int]$ChannelAddRetries  = 6,
  [int]$ChannelNameMaxLen  = 50
)

# ============ helpers ============
function Ensure-Module {
  param([string]$Name)
  if (-not (Get-Module -ListAvailable -Name $Name)) {
    Write-Host "Installing module $Name ..."
    Install-Module $Name -Scope CurrentUser -Force -ErrorAction Stop
  }
  Import-Module $Name -ErrorAction Stop
}

function Get-FilePath($title, $filter) {
  if ($IsWindows) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = $filter
    $dlg.Title  = $title
    if ($dlg.ShowDialog() -eq "OK") { return $dlg.FileName } else { throw "No seleccionaste ningun archivo" }
  } elseif ($IsMacOS) {
    $osa = "POSIX path of (choose file with prompt ""$title"")"
    $p = & osascript -e $osa 2>$null
    if ([string]::IsNullOrWhiteSpace($p)) { throw "No seleccionaste ningun archivo" }
    return $p.Trim()
  } else {
    $p = Read-Host "$title (ruta completa)"
    if (-not (Test-Path $p)) { throw "No se encontro el archivo: $p" }
    return $p
  }
}

function Read-KeyValueConfig([string]$path) {
  $dic = @{}
  Get-Content -Path $path | ForEach-Object {
    $line = $_.Trim()
    if (-not $line) { return }
    if ($line -match '^\s*#') { return }
    if ($line -match '^\s*([^=:#]+)\s*=\s*(.*)\s*$') {
      $k = $matches[1].Trim()
      $v = $matches[2].Trim()
      $dic[$k.ToLowerInvariant()] = $v
    }
  }
  return $dic
}

# normalizacion ascii + title case
function Normalize-AsciiTitle {
  param([string]$s)
  if ([string]::IsNullOrWhiteSpace($s)) { return $s }
  $lower = $s.ToLowerInvariant()
  $formD = $lower.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object System.Text.StringBuilder
  foreach ($ch in $formD.ToCharArray()) {
    $uc = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
    if ($uc -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
      switch ($ch) { 'ñ' { [void]$sb.Append('n') } default { [void]$sb.Append($ch) } }
    }
  }
  $noDia = $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
  $ti = (New-Object System.Globalization.CultureInfo("en-US")).TextInfo
  $title = $ti.ToTitleCase($noDia)
  $title = ([regex]::Replace($title, "\s+", " ")).Trim()
  return $title
}

function Build-ChannelName {
  param([string]$fullName, [int]$maxLen)
  $prefix = "Canal privado de entrega de "
  $name = Normalize-AsciiTitle $fullName
  function Mk([string]$n){ return ($prefix + $n) }

  $candidate = Mk $name
  if ($candidate.Length -le $maxLen) { return $candidate }

  $parts = $name.Split(" ")
  if ($parts.Count -ge 3) {
    $name1 = ($parts[0..($parts.Count-2)] -join " ")
    $cand1 = Mk $name1
    if ($cand1.Length -le $maxLen) { return $cand1 }
  }
  if ($parts.Count -ge 2) {
    $name2 = ($parts[0] + " " + $parts[1])
    $cand2 = Mk $name2
    if ($cand2.Length -le $maxLen) { return $cand2 }
  }
  if ($parts.Count -ge 2) {
    $name3 = ($parts[0].Substring(0,1) + ". " + $parts[1])
    $cand3 = Mk $name3
    if ($cand3.Length -le $maxLen) { return $cand3 }
  }
  return (Mk $name).Substring(0, $maxLen)
}

function Ensure-TeamMember {
  param([string]$GroupId, [string]$Email)
  $emailNorm = $Email.Trim().ToLowerInvariant()
  $tu = Get-TeamUser -GroupId $GroupId -ErrorAction SilentlyContinue |
        Where-Object { $_.User -and ($_.User.ToLowerInvariant() -eq $emailNorm) }
  if (-not $tu) {
    try {
      Add-TeamUser -GroupId $GroupId -User $emailNorm -Role Member -ErrorAction Stop
      Start-Sleep -Seconds 4
    } catch { Write-Warning "Could not add $emailNorm to team: $_" }
    $tu = Get-TeamUser -GroupId $GroupId -ErrorAction SilentlyContinue |
          Where-Object { $_.User -and ($_.User.ToLowerInvariant() -eq $emailNorm) }
  }
  if ($tu) { return $tu.User } else { return $null }
}

function Retry-AddChannelMember {
  param([string]$GroupId, [string]$ChannelName, [string]$UserUpn, [int]$Tries, [int]$Delay)
  for ($t=1; $t -le $Tries; $t++) {
    try {
      Start-Sleep -Seconds $Delay
      Add-TeamChannelUser -GroupId $GroupId -DisplayName $ChannelName -User $UserUpn -ErrorAction Stop
      Write-Host "Channel member added: $UserUpn (try $t)"
      return $true
    } catch { Write-Warning "Add member failed on try $t for $UserUpn in '$ChannelName': $_" }
  }
  return $false
}

# ============ modulos ============
try { Ensure-Module -Name MicrosoftTeams } catch { throw $_ }
try { Ensure-Module -Name ImportExcel }  catch { throw $_ }

# ============ seleccionar archivos ============
$configPath = Get-FilePath -title "Selecciona config.txt" -filter "Text files (*.txt)|*.txt|All files (*.*)|*.*"
$xlsxPath   = Get-FilePath -title "Selecciona el archivo XLSX" -filter "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"

# ============ leer config ============
$cfg = Read-KeyValueConfig -path $configPath

# requeridos
$teamName = $cfg['teamname']
$ownerUpn = $cfg['ownerupn']
if ([string]::IsNullOrWhiteSpace($teamName)) { throw "TeamName (en config.txt) es obligatorio. Abortando." }
if ([string]::IsNullOrWhiteSpace($ownerUpn)) { throw "OwnerUPN (en config.txt) es obligatorio. Abortando." }
$ownerUpn = $ownerUpn.Trim().ToLowerInvariant()

# parsear TeamName: "CSTI-1900-5337 - Asignatura"
$pattern = '^[^-]+-(\d+)-(\d+)\s*-\s*(.+)$'
if ($teamName -match $pattern) {
  $quarter = $matches[1]
  $groupNo = $matches[2]
  $course  = $matches[3].Trim()
} else {
  Write-Warning "TeamName no coincide con el formato esperado 'CSTI-1900-5337 - Curso'."
  $quarter = "19XX"
  $groupNo = ""
  $course  = $teamName
}

# career desde config o prompt (visibility se ignora: EDU siempre privado)
$career = $cfg['career']
if ([string]::IsNullOrWhiteSpace($career)) {
  $career = Read-Host "Carrera (default ITT/ICC)"
  if ([string]::IsNullOrWhiteSpace($career)) { $career = "ITT/ICC" }
}

# description desde config o molde
$description = $cfg['description']
if ([string]::IsNullOrWhiteSpace($description)) {
  $description = "Grupo de Teams de {Course} ({TeamName}) del cuatrimestre {Quarter} de la carrera {Career}."
}
$description = $description.Replace("{TeamName}", $teamName).
                            Replace("{Quarter}", $quarter).
                            Replace("{Career}", $career).
                            Replace("{Course}", $course)

# overrides de ejecucion
if ($cfg['channeladddelaysec']) { [int]$ChannelAddDelaySec = $cfg['channeladddelaysec'] }
if ($cfg['channeladdretries'])  { [int]$ChannelAddRetries  = $cfg['channeladdretries']  }
if ($cfg['channelnamemaxlen'])  { [int]$ChannelNameMaxLen  = $cfg['channelnamemaxlen']  }

Write-Host "Config leida:"
Write-Host " TeamName: $teamName"
Write-Host " Quarter: $quarter"
Write-Host " Group#:  $groupNo"
Write-Host " Course:  $course"
Write-Host " Career:  $career"
Write-Host " OwnerUPN: $ownerUpn"
Write-Host " Description: $description"
Write-Host " Delay/Tries/MaxLen: $ChannelAddDelaySec / $ChannelAddRetries / $ChannelNameMaxLen"

# ============ leer XLSX por posicion A..F, sin headers ============
$rows = Import-Excel -Path $xlsxPath -NoHeader
$participants = @()
$idx = 0
foreach ($r in $rows) {
  $idx++
  if ($idx -eq 1) { continue }  # saltar encabezados
  $vals = $r.PSObject.Properties.Value
  if ($vals.Count -lt 5) { continue }
  $email  = [string]$vals[2]  # C
  $nombre = [string]$vals[3]  # D
  $apell  = [string]$vals[4]  # E
  if ([string]::IsNullOrWhiteSpace($email) -or ($email -notmatch '@')) { continue }
  $full = ((($nombre + " " + $apell) -replace '\s+', ' ').Trim())
  $participants += [pscustomobject]@{ Email = $email.Trim().ToLowerInvariant(); Name = $full }
}
if ($participants.Count -eq 0) { throw "El XLSX no tiene filas validas (columna C vacia?). Abortando." }

# ============ conectar y CREAR CLASS TEAM ============
Write-Host "Connecting to Microsoft Teams..."
Connect-MicrosoftTeams

$teamId = $null
try {
  Write-Host "Creating EDU_Class team (sin visibility)..."
  $t = New-Team -DisplayName $teamName -Description $description -Template "EDU_Class" -ErrorAction Stop
  $teamId = $t.GroupId
  Write-Host ""
  Write-Host "============================================="
  Write-Host "  Clase creada: $($t.DisplayName)"
  Write-Host "  IMPORTANTE: abre Teams y pulsa 'Activar'."
  Write-Host "============================================="
  Write-Host ""
} catch {
  throw "Fallo creando el Class Team (EDU_Class): $_"
}

# ============ pausa para ACTIVACION manual ============
Read-Host "Presiona Enter cuando YA este ACTIVADO el team en la UI de Teams"

# ============ verificar que ya acepta miembros (pequeno loop) ============
$ready = $false
for ($i=1; $i -le 6 -and -not $ready; $i++) {
  try {
    # intento idempotente: solo para probar que ya acepta cambios
    Add-TeamUser -GroupId $teamId -User $ownerUpn -Role Owner -ErrorAction Stop
    $ready = $true
  } catch {
    Write-Warning "Aun no acepta miembros (intento $i). Esperando 8s..."
    Start-Sleep -Seconds 8
  }
}

if (-not $ready) {
  Write-Warning "No se pudo confirmar estado 'activo'. Continuare igualmente."
}

# ============ agregar owner (asegurar) + miembros ============
Write-Host "Adding owner and members to team..."
try { Add-TeamUser -GroupId $teamId -User $ownerUpn -Role Owner -ErrorAction Stop } catch { Write-Warning "Owner add warning: $_" }
foreach ($p in $participants) {
  try {
    if ($p.Email -eq $ownerUpn) { continue }
    Add-TeamUser -GroupId $teamId -User $p.Email -Role Member -ErrorAction Stop
  } catch { Write-Warning "Could not add $($p.Email) to team: $_" }
}
Start-Sleep -Seconds 6

# ============ crear canales privados y agregar estudiante ============
$usedNames = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($p in $participants) {
  if ($p.Email -eq $ownerUpn) { continue }  # no crear canal para el owner
  $channelName = Build-ChannelName -fullName $p.Name -maxLen $ChannelNameMaxLen
  $base=$channelName; $j=1
  while (-not $usedNames.Add($channelName)) {
    $suffix=" ($j)"; if (($base.Length+$suffix.Length) -gt $ChannelNameMaxLen){$base=$base.Substring(0,$ChannelNameMaxLen-$suffix.Length)}
    $channelName=$base+$suffix; $j++
  }
  try {
    New-TeamChannel -GroupId $teamId -DisplayName $channelName -MembershipType Private -ErrorAction Stop
    Write-Host "Private channel created: $channelName -> $($p.Email)"
    $resolved = Ensure-TeamMember -GroupId $teamId -Email $p.Email
    if (-not $resolved) { Write-Warning "Skipping: $($p.Email) no resuelto como miembro del team."; continue }
    $ok = Retry-AddChannelMember -GroupId $teamId -ChannelName $channelName -UserUpn $resolved -Tries $ChannelAddRetries -Delay $ChannelAddDelaySec
    if (-not $ok) { Write-Warning "No se pudo agregar $resolved a '$channelName' despues de reintentos." }
  } catch {
    Write-Warning "No se pudo crear/poblar el canal '$channelName': $_"
  }
}

Write-Host "Done."
