# elm2madrix_full.ps1 — Compatible Windows PowerShell 5.1 et PowerShell 7+
# Objectif:
# - Sortie base 0 (univers & canaux) par défaut
# - Option -ChannelOffset -1 pour décaler toutes les adresses d’un canal vers le bas
# - Emprunt/carry inter‑univers pour éviter les 0 illégaux quand base 1 (désactivable)
# - Display Name = Group + SubGroup + compteur local
# - Export CSV en UTF‑8 avec BOM

[CmdletBinding()]
param(
  # Fichiers d’entrée (wildcards OK)
  [Parameter(ValueFromRemainingArguments = $true)]
  [string[]] $Paths,
  [switch] $Recurse,

  # Valeurs de sortie par défaut (Madrix)
  [string] $DisplayTemplate = 'GRB_light #{0:D4}',
  [string] $Mode           = '1 pixel',
  [string] $Manufacturer   = '!generic',
  [string] $Type           = 'DMX',
  [string] $Rotation       = '0',
  [string] $Product        = '!generic RGB Light 1 pixel',

  # CSV sortie
  [string] $OutDelimiter   = ',',

  # Dossier de sortie
  [string] $OutDir = '',
  [string] $OutSubfolder = 'Export vers Madrix 5.5A',
  [switch] $NoSubfolder,

  # Détection/forçage du séparateur d’entrée (0 = auto)
  [char] $InDelimiter = [char]0,

  # Format DMX "Universe{sep}Channel"
  [ValidatePattern('^.$')][string] $DmxSeparator = '.',

  # Bases/offsets ELM -> Madrix
  [int] $UniverseBaseIn  = 0,
  [int] $ChannelBaseIn   = 1,
  [int] $UniverseBaseOut = 0,  # base 0 demandée
  [int] $ChannelBaseOut  = 0,  # base 0 demandée
  [int] $UniverseOffset  = 0,
  [int] $ChannelOffset   = 0,

  # Emprunt/carry inter‑univers lorsque canal déborde
  [switch] $BorrowAcrossUniverse = $true,

  # Colonnes Group/SubGroup
  [string] $DisplayJoiner = ' - ',
  [string[]] $GroupCandidates    = @('Group','Groupe','Group Name','Fixture Group'),
  [string[]] $SubGroupCandidates = @('Sub Group','SubGroup','Sous Groupe','Sous-Groupe','Subgroup','Group 2')
)

# ---------- Logs ----------
function Write-Info([string]$m){ Write-Host "[INFO ] $m"  -ForegroundColor Cyan }
function Write-Ok  ([string]$m){ Write-Host "[ OK  ] $m"  -ForegroundColor Green }
function Write-Warn([string]$m){ Write-Host "[WARN ] $m"  -ForegroundColor Yellow }
function Write-Err ([string]$m){ Write-Host "[ERROR] $m" -ForegroundColor Red }

# ---------- Utilitaires ----------
function Detect-Delimiter {
  param([string]$Path)
  $sample = Get-Content -Path $Path -TotalCount 5 -ErrorAction Stop | Out-String
  $cComma = ([regex]::Matches($sample, ',')).Count
  $cSemi  = ([regex]::Matches($sample, ';')).Count
  $cTab   = ([regex]::Matches($sample, "`t")).Count
  if ($cTab  -ge $cComma -and $cTab  -ge $cSemi ) { return "`t" }
  if ($cSemi -ge $cComma -and $cSemi -ge $cTab  ) { return ';' }
  return ','
}

function Import-CsvSafe {
  param([string]$Path, [char]$Delimiter)
  try {
    if ($Delimiter -eq 0) {
      $d = Detect-Delimiter -Path $Path
      Write-Info "Séparateur détecté: '$d' pour $Path"
    } else { $d = $Delimiter }
    $rows = Import-Csv -Path $Path -Delimiter $d -ErrorAction Stop
    return @{ Ok=$true; Rows=$rows; Delim=$d }
  } catch {
    return @{ Ok=$false; Error=$_.Exception.Message }
  }
}

function Export-CsvUtf8Bom {
  param([object[]]$Rows, [string]$Path, [char]$Delimiter)
  $tmp = [System.IO.Path]::GetTempFileName()
  try {
    $Rows | Export-Csv -Path $tmp -NoTypeInformation -Encoding UTF8 -Delimiter $Delimiter
    $bytes = [System.Text.Encoding]::UTF8.GetBytes([System.IO.File]::ReadAllText($tmp))
    $bom   = [byte[]](0xEF,0xBB,0xBF)
    [System.IO.File]::WriteAllBytes($Path, $bom + $bytes)
  } finally {
    Remove-Item -Path $tmp -ErrorAction SilentlyContinue
  }
}

function Find-ColumnName {
  param([hashtable]$Map, [string[]]$Candidates)
  foreach($c in $Candidates){
    if ($Map.ContainsKey($c)) { return $c }
    foreach($k in $Map.Keys){ if ($k -ieq $c) { return $k } }
  }
  return $null
}

function Get-Field {
  param([hashtable]$Map, [psobject]$Row, [string[]]$Candidates, [object]$Default=$null)
  $name = Find-ColumnName -Map $Map -Candidates $Candidates
  if ($null -ne $name) { return $Row.$name }
  return $Default
}

# Normalisation DMX avec emprunt/carry inter‑univers
function Normalize-Dmx {
  param(
    [int]$uIn, [int]$cIn,
    [int]$UniverseBaseIn, [int]$ChannelBaseIn,
    [int]$UniverseBaseOut, [int]$ChannelBaseOut,
    [int]$UniverseOffset, [int]$ChannelOffset,
    [switch]$BorrowAcrossUniverse,
    [string]$Sep
  )
  # 1) Convertir en 0‑based absolu
  $uAbs = $uIn - $UniverseBaseIn
  $cAbs = $cIn - $ChannelBaseIn

  # 2) Appliquer offsets
  $uAbs += $UniverseOffset
  $cAbs += $ChannelOffset

  # 3) Revenir à la base de sortie
  $uOut = $uAbs + $UniverseBaseOut
  $cOut = $cAbs + $ChannelBaseOut

  # 4) Définir les bornes de canal selon la base sortie
  $chMin = 0; $chMax = 0
  if ($ChannelBaseOut -eq 0) { $chMin = 0; $chMax = 511 } else { $chMin = 1; $chMax = 512 }

  if ($BorrowAcrossUniverse.IsPresent) {
    $span = $chMax - $chMin + 1
    while ($cOut -lt $chMin) { $cOut += $span; $uOut -= 1 }
    while ($cOut -gt $chMax) { $cOut -= $span; $uOut += 1 }
  } else {
    if     ($cOut -lt $chMin) { $cOut = $chMin }
    elseif ($cOut -gt $chMax) { $cOut = $chMax }
  }

  if ($uOut -lt 0) { $uOut = 0 }

  # 5) Format "UUU{sep}CCC"
  $uStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:D3}", $uOut)
  $cStr = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:D3}", $cOut)
  return @{ U=$uOut; C=$cOut; S="$uStr$Sep$cStr" }
}

function Convert-ElmToMadrix {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory)][object[]]$rows,
    [Parameter(Mandatory)][string]$Product,
    [Parameter(Mandatory)][string]$DisplayTemplate,
    [Parameter(Mandatory)][string]$Mode,
    [Parameter(Mandatory)][string]$Manufacturer,
    [Parameter(Mandatory)][string]$Type,
    [Parameter(Mandatory)][string]$Rotation,
    [Parameter(Mandatory)][string]$DmxSeparator,

    [Parameter(Mandatory)][int]$UniverseBaseIn,
    [Parameter(Mandatory)][int]$ChannelBaseIn,
    [Parameter(Mandatory)][int]$UniverseBaseOut,
    [Parameter(Mandatory)][int]$ChannelBaseOut,
    [Parameter(Mandatory)][int]$UniverseOffset,
    [Parameter(Mandatory)][int]$ChannelOffset,

    [Parameter(Mandatory)][string]$DisplayJoiner,
    [Parameter(Mandatory)][string[]]$GroupCandidates,
    [Parameter(Mandatory)][string[]]$SubGroupCandidates,
    [switch]$BorrowAcrossUniverse
  )

  $out = New-Object System.Collections.Generic.List[object]
  $mapNames = @{}
  if ($rows.Count -gt 0) {
    $first = $rows[0] | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
    foreach($n in $first){ $mapNames[$n] = $n }
  }

  $U_candidates = @('Universe','DMX Universe','ArtNet Universe','ELM Universe','U','Univers')
  $C_candidates = @('Channel','DMX Channel','Start Channel','Address','C','Canal')
  $X_candidates = @('X','Pos X','Position X')
  $Y_candidates = @('Y','Pos Y','Position Y')

  $counters = @{}
  $i = 0

  foreach($r in $rows){
    $i++

    $uIn = [int](Get-Field -Map $mapNames -Row $r -Candidates $U_candidates -Default 0)
    $cIn = [int](Get-Field -Map $mapNames -Row $r -Candidates $C_candidates -Default 0)

    $norm = Normalize-Dmx -uIn $uIn -cIn $cIn `
      -UniverseBaseIn $UniverseBaseIn -ChannelBaseIn $ChannelBaseIn `
      -UniverseBaseOut $UniverseBaseOut -ChannelBaseOut $ChannelBaseOut `
      -UniverseOffset $UniverseOffset -ChannelOffset $ChannelOffset `
      -BorrowAcrossUniverse:$BorrowAcrossUniverse -Sep $DmxSeparator

    $x = Get-Field -Map $mapNames -Row $r -Candidates $X_candidates -Default $null
    $y = Get-Field -Map $mapNames -Row $r -Candidates $Y_candidates -Default $null

    $grp = (Get-Field -Map $mapNames -Row $r -Candidates $GroupCandidates -Default '').ToString().Trim()
    $sub = (Get-Field -Map $mapNames -Row $r -Candidates $SubGroupCandidates -Default '').ToString().Trim()

    if (-not $grp -and -not $sub) {
      $disp = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, $DisplayTemplate, $i)
    } else {
      $key = "$grp|$sub"
      if (-not $counters.ContainsKey($key)) { $counters[$key] = 0 }
      $counters[$key] = [int]$counters[$key] + 1
      $n = $counters[$key]
      if ($grp -and $sub) { $disp = "$grp$DisplayJoiner$sub$DisplayJoiner$n" }
      elseif ($grp)       { $disp = "$grp$DisplayJoiner$n" }
      else                { $disp = "$sub$DisplayJoiner$n" }
    }

    $fixId = [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, "{0:D6}", $i)

    $out.Add([PSCustomObject]@{
      'Fixture ID'   = "$fixId"
      'Product'      = "$Product"
      'Display Name' = "$disp"
      'Position X'   = $x
      'Position Y'   = $y
      'DMX Address'  = $norm.S
    })
  }

  return ,$out
}

# ---------- Programme principal ----------
$script:__HadError = $false
try {
  if (-not $Paths -or $Paths.Count -eq 0) {
    Write-Err "Aucun chemin fourni. Exemple: -Paths 'C:\Data\*.csv'"
    exit 1
  }

  # Résoudre les chemins
  $expanded = New-Object System.Collections.Generic.List[string]
  foreach($p in $Paths){
    if ($Recurse) {
      Get-ChildItem -Path $p -Recurse -File -ErrorAction SilentlyContinue | ForEach-Object { $expanded.Add($_.FullName) }
    } else {
      Get-ChildItem -Path $p -File -ErrorAction SilentlyContinue | ForEach-Object { $expanded.Add($_.FullName) }
    }
  }

  if ($expanded.Count -eq 0) { Write-Err "Aucun fichier trouvé."; exit 1 }

  foreach($csv in $expanded){
    try {
      Write-Info "Traitement: $csv"
      $imp = Import-CsvSafe -Path $csv -Delimiter $InDelimiter
      if (-not $imp.Ok) { throw $imp.Error }
      $rows = $imp.Rows

      $converted = Convert-ElmToMadrix -rows $rows `
        -Product $Product -DisplayTemplate $DisplayTemplate -Mode $Mode -Manufacturer $Manufacturer -Type $Type -Rotation $Rotation -DmxSeparator $DmxSeparator `
        -UniverseBaseIn $UniverseBaseIn -ChannelBaseIn $ChannelBaseIn -UniverseBaseOut $UniverseBaseOut -ChannelBaseOut $ChannelBaseOut `
        -UniverseOffset $UniverseOffset -ChannelOffset $ChannelOffset `
        -DisplayJoiner $DisplayJoiner -GroupCandidates $GroupCandidates -SubGroupCandidates $SubGroupCandidates `
        -BorrowAcrossUniverse:$BorrowAcrossUniverse

      # Dossier de sortie
      $inFull  = [System.IO.Path]::GetFullPath($csv)
      $dirIn   = [System.IO.Path]::GetDirectoryName($inFull)
      $stem    = [System.IO.Path]::GetFileNameWithoutExtension($inFull)

      if ($OutDir -and $OutDir.Trim().Length -gt 0) {
        $targetDir = [System.IO.Path]::GetFullPath($OutDir)
      } elseif ($NoSubfolder) {
        $targetDir = $dirIn
      } else {
        $targetDir = [System.IO.Path]::Combine($dirIn, $OutSubfolder)
      }
      if (-not (Test-Path -Path $targetDir)) {
        Write-Info "Création du dossier: $targetDir"
        New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
      }

      $outCsv  = [System.IO.Path]::Combine($targetDir, "$stem`_madrix.csv")

      # Export UTF‑8 BOM
      $outDelimChar = [char]$OutDelimiter[0]
      Export-CsvUtf8Bom -Rows $converted -Path $outCsv -Delimiter $outDelimChar

      Write-Ok "Exporté: $outCsv"
      Write-Host "Script de conversion patch ELM vers Madrix 5.5A, par Brandon Lainé, Eklipse LED." -ForegroundColor Magenta
    } catch {
      $script:__HadError = $true
      Write-Err ("Échec sur '{0}' : {1}" -f $csv, $_.Exception.Message)
    }
  }

} catch {
  $script:__HadError = $true
  Write-Err $_.Exception.Message
}

exit ([int]([bool]$script:__HadError))
