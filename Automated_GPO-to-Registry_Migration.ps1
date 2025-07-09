<#
  Export Office-2016 GPO ► inject registry items into
  Preferences ▸ Windows Settings ▸ Registry  (same GPO)
  OUTPUT under C:\Temp:
     • LGPO\LGPO.exe
     • GPOExport-<name>-yyyymmdd_HHmmss\user.txt / machine.txt / gpreport.xml
#>

param(
    [Parameter(Mandatory)][string]$GpoName,
    [string]$TempRoot = 'C:\Temp',
    [string]$LgpoDir  = 'C:\Temp\LGPO'   # LGPO.exe lives here
)

# ── TLS 1.2 for Invoke-WebRequest (old servers default to TLS 1.0) ─────────
if (-not ([Net.ServicePointManager]::SecurityProtocol -band [Net.SecurityProtocolType]::Tls12)) {
    [Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls12
}

# ── mini-fn: download LGPO.exe once ────────────────────────────────────────
function Get-LGPOExe {
    param([string]$Dir)
    $exe = Join-Path $Dir 'LGPO.exe'
    if (Test-Path $exe) { return $exe }

    Write-Host 'LGPO.exe not found – downloading …'
    $zipURL = 'https://download.microsoft.com/download/8/5/c/85c25433-a1b0-4ffa-9429-7e023e7da8d8/LGPO.zip'
    $zip    = Join-Path $Dir 'LGPO.zip'
    New-Item -ItemType Directory -Path $Dir -Force | Out-Null
    Invoke-WebRequest -Uri $zipURL -OutFile $zip
    Expand-Archive -Path $zip -DestinationPath $Dir -Force
    Remove-Item $zip
    if (-not (Test-Path $exe)) { throw 'LGPO.exe missing after extraction.' }
    Write-Host "✓ LGPO.exe saved to $exe"
    return $exe
}

# ── mini-fn: parse new LGPO /parse 4-line blocks ───────────────────────────
function Parse-LGPOText {
    param([string]$Path,[string]$Hive)
    if (-not (Test-Path $Path)) { return @() }
    $rows   = Get-Content $Path
    $blocks = @()
    $buf    = @()
    foreach ($l in $rows) {
        if ($l.Trim()) { $buf += $l.Trim() }
        else {
            if ($buf.Count -ge 4) { $blocks += ,$buf }; $buf=@()
        }
    }
    if ($buf.Count -ge 4) { $blocks += ,$buf }

    foreach ($b in $blocks) {
        $keyPath   = $b[1]
        $valueName = $b[2]
        $pair      = $b[3] -split ':',2
        $rawType   = $pair[0].Trim().ToUpper()   # SZ / EXSZ / DWORD
        $data      = $pair[1].Trim()

        switch ($rawType) {
            'SZ'     { $type = 'String' }
            'EXSZ'   { $type = 'ExpandString' }
            'DWORD'  { $type = 'Dword' ; $data =[int]$data }
            default  { $type = 'String' }
        }

        [pscustomobject]@{
            Hive  = $Hive          # HKCU / HKLM
            Key   = $keyPath       # without hive prefix
            Name  = $valueName
            Type  = $type
            Data  = $data
        }
    }
}

# ── 0 · prerequisites ──────────────────────────────────────────────────────
Import-Module GroupPolicy -ErrorAction Stop
try { $null = Get-GPO -Name $GpoName -ErrorAction Stop }
catch { throw "GPO '$GpoName' not found." }

$lgpo   = Get-LGPOExe $LgpoDir
$stamp  = Get-Date -Format yyyyMMdd_HHmmss
$export = Join-Path $TempRoot ("GPOExport-$($GpoName -replace '\s','_')-$stamp")
New-Item -ItemType Directory -Path $export -Force | Out-Null

# ── 1 · Backup-GPO ─────────────────────────────────────────────────────────
Backup-GPO -Name $GpoName -Path $export | Out-Null
$guidDir = (Get-ChildItem -Path $export -Directory | Select-Object -First 1).FullName
Copy-Item -Path (Join-Path $guidDir 'gpreport.xml') -Destination $export -Force
Write-Host "✓ GPO backup stored in $guidDir"

$userPol = Join-Path $guidDir 'DomainSysvol\GPO\User\registry.pol'
$machPol = Join-Path $guidDir 'DomainSysvol\GPO\Machine\registry.pol'
if (-not (Test-Path $userPol) -and -not (Test-Path $machPol)) {
    Write-Warning 'No registry.pol files found – nothing to process.' ; exit
}

# ── 2 · LGPO /parse → txt → objects ───────────────────────────────────────
$settings = @()

if (Test-Path $userPol) {
    $uTxt = Join-Path $export 'user.txt'
    & $lgpo /parse /u $userPol > $uTxt
    Write-Host "✓ USER registry.pol dumped → user.txt"
    $settings += Parse-LGPOText -Path $uTxt -Hive 'HKCU'
}

if (Test-Path $machPol) {
    $mTxt = Join-Path $export 'machine.txt'
    & $lgpo /parse /m $machPol > $mTxt
    Write-Host "✓ MACHINE registry.pol dumped → machine.txt"
    $settings += Parse-LGPOText -Path $mTxt -Hive 'HKLM'
}

if (-not $settings) { Write-Warning 'Parsed text contained no registry rows.'; exit }

# ── 3 · Push each row into Preferences ▸ Registry ──────────────────────────
$created = 0
foreach ($s in $settings) {
    $ctx      = if ($s.Hive -eq 'HKCU') { 'User' } else { 'Computer' }
    $fullKey  = "$($s.Hive)\$($s.Key)"

    Set-GPPrefRegistryValue -Name $GpoName `
        -Context   $ctx `
        -Action    Create `
        -Key       $fullKey `
        -ValueName $s.Name `
        -Type      $s.Type `
        -Value     $s.Data `
        -ErrorAction Stop

    $created++
}

Write-Host "`n✓ Added $created registry preference items to '$GpoName'."
Write-Host "   Reopen the GPO ► Preferences ► Windows Settings ► Registry to confirm."
Write-Host "`nArtifacts live under:  $export"
