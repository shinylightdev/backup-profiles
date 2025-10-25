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

# Use wildcard to add folder contents (helps 7z include files correctly)
$pathToAdd = (Join-Path $sourceDir '*')

# build args: single invocation, max compression, multithread, try open-for-write files
$args = @(
    'a'
    '-tzip'
    $destZip
    $pathToAdd
    '-mx=9'
    '-mmt=on'
    '-ssw'
    '-r'
    '-bd'
    '-y'
)

# normalize/dedupe
$args = $args | ForEach-Object { if ($_ -is [string]) { $_.TrimEnd('\') } else { $_ } } | Select-Object -Unique

# ensure a valid (empty) zip exists so 7z won't leave nothing if the first add fails
if (-not (Test-Path $destZip)) {
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
        $z = [IO.Compression.ZipFile]::Open($destZip, [IO.Compression.ZipArchiveMode]::Create)
        $z.Dispose()
    } catch {
        New-Item -Path $destZip -ItemType File -Force | Out-Null
    }
}

# run 7z and capture output
$logFile = Join-Path $env:TEMP '7z_run_log.txt'
$output = & $sevenZipExe @args 2>&1
$exit = $LASTEXITCODE
$output | Out-File -FilePath $logFile -Encoding UTF8 -Force

# check whether archive contains entries
$zipHasEntries = $false
try {
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $zr = [IO.Compression.ZipFile]::OpenRead($destZip)
    $zipHasEntries = ($zr.Entries.Count -gt 0)
    $zr.Dispose()
} catch {
    $zipHasEntries = $false
}

if ($zipHasEntries) {
    if ($exit -ne 0) { Write-Warning "7-Zip finished with errors but archive was created: $destZip. See $logFile" }
    return
}

# fallback: create fast robocopy mirror that skips locked files, then compress the mirror
$mirror = Join-Path $env:TEMP ("zip_source_copy_{0:yyyyMMddHHmmss}" -f (Get-Date))
New-Item -Path $mirror -ItemType Directory -Force | Out-Null

$robocopyArgs = @(
    ('"{0}"' -f $sourceDir),
    ('"{0}"' -f $mirror),
    '/E',        # include empty dirs
    '/R:0',      # no retries
    '/W:0',      # no wait
    '/MT:16',    # multithreaded copy
    '/NFL','/NDL' # suppress listing
)

$rc = Start-Process -FilePath 'robocopy' -ArgumentList $robocopyArgs -NoNewWindow -Wait -PassThru -ErrorAction SilentlyContinue
$rcExit = if ($rc) { $rc.ExitCode } else { 1 }

# robocopy exit codes < 16 are acceptable (files copied, some skipped). >=16 = failure
if ($rcExit -ge 16) {
    Write-Error "Robocopy failed (exit $rcExit). See robocopy output. Aborting."
    Remove-Item -LiteralPath $mirror -Recurse -Force -ErrorAction SilentlyContinue
    return
}

# run 7z against the mirror (single, fast pass)
$args2 = @('a','-tzip',$destZip, (Join-Path $mirror '*'), '-mx=9','-mmt=on','-r','-bd','-y')
$args2 = $args2 | ForEach-Object { if ($_ -is [string]) { $_.TrimEnd('\') } else { $_ } } | Select-Object -Unique
$output2 = & $sevenZipExe @args2 2>&1
$exit2 = $LASTEXITCODE
$output2 | Out-File -FilePath $logFile -Append -Encoding UTF8

# cleanup mirror
Remove-Item -LiteralPath $mirror -Recurse -Force -ErrorAction SilentlyContinue

# final check
try {
    $zr = [IO.Compression.ZipFile]::OpenRead($destZip)
    $has = ($zr.Entries.Count -gt 0)
    $zr.Dispose()
} catch {
    $has = $false
}

if ($has) {
    if ($exit2 -ne 0) { Write-Warning "Archive created with warnings (exit $exit2). See $logFile" }
} else {
    Write-Error "Failed to create archive. See $logFile for 7z output."
}
