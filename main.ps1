# ...existing code...

$sourceDir = 'C:\Users\Dan\AppData\Local\Google\Chrome\User Data'
$destZip  = 'C:\temp\save\1.zip'

# ensure destination folder exists
$destFolder = Split-Path $destZip
if (-not (Test-Path $destFolder)) { New-Item -Path $destFolder -ItemType Directory -Force | Out-Null }

# remove existing zip if present (overwrite)
if (Test-Path $destZip) { Remove-Item $destZip -Force }

Add-Type -AssemblyName System.IO.Compression.FileSystem

$sourceFull = [IO.Path]::GetFullPath($sourceDir)
$files = Get-ChildItem -Path $sourceDir -Recurse -File

# total input bytes (used to compute progress)
$totalBytes = ($files | Measure-Object -Property Length -Sum).Sum
if (-not $totalBytes) { $totalBytes = 0 }

$zip = [IO.Compression.ZipFile]::Open($destZip, [IO.Compression.ZipArchiveMode]::Create)
try {
    $cumulative = 0
    $lastPercent = -1
    foreach ($file in $files) {
        $relative = $file.FullName.Substring($sourceFull.Length).TrimStart('\','/')

        $added = $false
        try {
            # normal fast path
            [IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $zip,
                $file.FullName,
                $relative,
                [IO.Compression.CompressionLevel]::Optimal
            )
            $added = $true
        } catch [System.IO.IOException] {
            # try opening with shared read (may succeed if file allows shared reads)
            try {
                $fs = [System.IO.File]::Open($file.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                try {
                    $entry = $zip.CreateEntry($relative, [IO.Compression.CompressionLevel]::Optimal)
                    $dest = $entry.Open()
                    try { $fs.CopyTo($dest) } finally { $dest.Close() }
                    $added = $true
                } finally { $fs.Close() }
            } catch {
                # still locked or another error — skip file silently
                $added = $false
            }
        } catch {
            # other errors — skip silently
            $added = $false
        }

        # update progress based on original file sizes (count skipped files so progress completes)
        $cumulative += $file.Length
        $percent = if ($totalBytes -gt 0) { [int](($cumulative * 100) / $totalBytes) } else { 100 }
        if ($percent -ne $lastPercent) {
            Write-Host "`r$percent%" -NoNewline
            $lastPercent = $percent
        }
    }

    if ($lastPercent -lt 100) { Write-Host "`r100%" -NoNewline }
}
finally { $zip.Dispose() }

Write-Output "Created zip: $destZip (Compression: Optimal)"
# ...existing code...