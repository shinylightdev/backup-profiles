# $sourceDir = 'C:\Users\Dan\AppData\Local\Google\Chrome\User Data'


$sevenZipPath = 'C:\Program Files\7-Zip'

$sourceDir = 'C:\Users\Dan\AppData\Local\Google\Chrome\User Data\Default'
$destZip  = 'C:\temp\save\1.zip'

$sevenZipExe = Join-Path $sevenZipPath '7z.exe'
if (-not (Test-Path $sevenZipExe)) {
    Write-Error "7z.exe not found at $sevenZipExe"
    return
}

# ensure destination folder exists
$destFolder = Split-Path $destZip
if (-not (Test-Path $destFolder)) { New-Item -Path $destFolder -ItemType Directory -Force | Out-Null }

function Get-FreeDriveLetter {
    for ($i = [int][char]'Z'; $i -ge [int][char]'D'; $i--) {
        $letter = [char]$i
        if (-not (Test-Path "$letter`:\")) { return $letter }
    }
    return $null
}

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# build 7z arguments function (so we can reuse for snapshot path)
function Build-7zArgs($pathToAdd, $destZipPath) {
    $a = @(
        'a'
        '-tzip'
        $destZipPath
        (Join-Path $pathToAdd '*')
        '-mx=9'
        '-r'
    )

    # include empty directories explicitly
    $hasAnyFiles = Get-ChildItem -Path $pathToAdd -Recurse -Force -File -ErrorAction SilentlyContinue
    if (-not $hasAnyFiles) {
        $a += $pathToAdd
    } else {
        $emptyDirs = Get-ChildItem -Path $pathToAdd -Directory -Recurse -Force -ErrorAction SilentlyContinue |
                     Where-Object { -not (Get-ChildItem -Path $_.FullName -Recurse -Force -File -ErrorAction SilentlyContinue) }
        foreach ($d in $emptyDirs) { $a += $d.FullName }
    }
    return $a
}

# try normal compression first
$args = Build-7zArgs -pathToAdd $sourceDir -destZipPath $destZip
& $sevenZipExe @args
if ($LASTEXITCODE -eq 0) { return }

# if we get here, some files may be locked -> attempt VSS snapshot approach
if (-not (Test-IsAdmin)) {
    Write-Error "Administrative privileges are required to create a VSS snapshot. Rerun elevated."
    return
}

$driveRoot = [IO.Path]::GetPathRoot($sourceDir)       # e.g. "C:\"
$vol = $driveRoot.TrimEnd('\')                       # e.g. "C:"
$freeLetter = Get-FreeDriveLetter
if (-not $freeLetter) {
    Write-Error "No free drive letter found to expose snapshot."
    return
}
$exposedDrive = "$freeLetter`:"

# prepare DiskShadow scripts
$tmpCreate = Join-Path $env:TEMP 'create_vss.dsh'
$tmpDelete = Join-Path $env:TEMP 'delete_vss.dsh'
$alias = 'MySnap'

$createContent = @(
    "SET CONTEXT PERSISTENT"
    "ADD VOLUME $vol ALIAS $alias"
    "CREATE"
    "EXPOSE %$alias% $freeLetter`:"
)
$deleteContent = @(
    "SET CONTEXT PERSISTENT"
    "DELETE SHADOWS SET %$alias%"
)

$createContent -join "`r`n" | Set-Content -Path $tmpCreate -Encoding ASCII
$deleteContent -join "`r`n" | Set-Content -Path $tmpDelete -Encoding ASCII

$created = $false
try {
    # create & expose snapshot (diskshadow must be available)
    $p = Start-Process -FilePath 'diskshadow.exe' -ArgumentList "/s `"$tmpCreate`"" -Wait -NoNewWindow -PassThru -ErrorAction Stop
    if ($p.ExitCode -ne 0) { throw "diskshadow create failed (exit $($p.ExitCode))" }
    $created = $true

    # map source path into snapshot namespace
    $relative = $sourceDir.Substring($driveRoot.Length).TrimStart('\')
    $snapshotSource = Join-Path $exposedDrive $relative

    if (-not (Test-Path $snapshotSource)) {
        throw "Snapshot path not accessible: $snapshotSource"
    }

    # run 7z against the snapshot path
    $args = Build-7zArgs -pathToAdd $snapshotSource -destZipPath $destZip
    & $sevenZipExe @args
    if ($LASTEXITCODE -ne 0) { throw "7-Zip failed inside snapshot (exit $LASTEXITCODE)" }
}
finally {
    if ($created) {
        # attempt to delete the snapshot (best effort)
        Start-Process -FilePath 'diskshadow.exe' -ArgumentList "/s `"$tmpDelete`"" -Wait -NoNewWindow -ErrorAction SilentlyContinue | Out-Null
    }
    # cleanup temp scripts
    Remove-Item -Path $tmpCreate -ErrorAction SilentlyContinue
    Remove-Item -Path $tmpDelete -ErrorAction SilentlyContinue
}
