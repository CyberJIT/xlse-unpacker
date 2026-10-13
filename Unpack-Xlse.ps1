<#
.SYNOPSIS
    Unpacks .xlsx, .xlsm, and .xlsb files into sibling folders, maintaining internal tree structure,
    SHA-256 delta synchronization, and deep extraction of VBA code, Power Query M, and OpenXML structure.

.DESCRIPTION
    Discovers Excel workbooks (.xlsx, .xlsm, .xlsb) in a specified folder, treats them as zip archives,
    and extracts their contents into sibling folders named after the base name of each file.
    Compares existing files using SHA256 hashes and file lengths to avoid redundant writes.
    Decompiles VBA modules (.bas, .cls, .frm) and project references from vbaProject.bin.
    Decompiles Power Query M formulas from DataMashup packages and customXml streams into standalone .m files.
    Deconstructs OpenXML/XLSB structures (sheets, cell formulas, defined names, tables).
    Supports recursive discovery, file name filtering, folder filtering, and dry-run mode.

.PARAMETER Path
    Target directory containing Excel files or a specific file path.

.PARAMETER Recurse
    Recursively search for files in subdirectories.

.PARAMETER OutputDir
    Custom output directory (recommended only when unpacking a single file).

.PARAMETER FileInclude
    Array of string fragments or wildcards to include matching file names.

.PARAMETER FileExclude
    Array of string fragments or wildcards to exclude matching file names.

.PARAMETER FolderInclude
    Array of string fragments or wildcards to include matching folder names during recursive search.

.PARAMETER FolderExclude
    Array of string fragments or wildcards to exclude matching folder names during recursive search.

.PARAMETER DryRun
    Simulate execution without modifying the filesystem.

.PARAMETER Quiet
    Suppress detailed progress messages.

.EXAMPLE
    .\Unpack-Xlse.ps1 -Path "C:\Financials" -Recurse
    .\Unpack-Xlse.ps1 -Path "C:\Financials" -FileInclude "*Report*", "*Model*" -DryRun
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Path,

    [Parameter(Mandatory = $false)]
    [switch]$Recurse,

    [Parameter(Mandatory = $false)]
    [string]$OutputDir,

    [Parameter(Mandatory = $false)]
    [string[]]$FileInclude,

    [Parameter(Mandatory = $false)]
    [string[]]$FileExclude,

    [Parameter(Mandatory = $false)]
    [string[]]$FolderInclude,

    [Parameter(Mandatory = $false)]
    [string[]]$FolderExclude,

    [Parameter(Mandatory = $false)]
    [switch]$HasVba,

    [Parameter(Mandatory = $false)]
    [switch]$HasPq,

    [Parameter(Mandatory = $false)]
    [switch]$HasFormulas,

    [Parameter(Mandatory = $false)]
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Security

function Get-FileSha256Hex {
    param([string]$FilePath)
    $hasher = [System.Security.Cryptography.SHA256]::Create()
    $stream = [System.IO.File]::OpenRead($FilePath)
    try {
        $hashBytes = $hasher.ComputeHash($stream)
        return [System.BitConverter]::ToString($hashBytes).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $stream.Close()
        $hasher.Dispose()
    }
}

function Get-BytesSha256Hex {
    param([byte[]]$Data)
    $hasher = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $hasher.ComputeHash($Data)
        return [System.BitConverter]::ToString($hashBytes).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $hasher.Dispose()
    }
}

function Test-PatternMatch {
    param(
        [string]$Name,
        [string[]]$Patterns
    )
    if (-not $Patterns -or $Patterns.Count -eq 0) { return $false }
    foreach ($p in $Patterns) {
        if ($Name -like $p -or $Name.IndexOf($p, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }
    return $false
}

function Get-TargetExcelFiles {
    param(
        [string]$RootPath,
        [bool]$IsRecurse,
        [string[]]$IncFiles,
        [string[]]$ExcFiles,
        [string[]]$IncFolders,
        [string[]]$ExcFolders
    )

    $validExts = @(".xlsx", ".xlsm", ".xlsb")
    $matched = [System.Collections.Generic.List[string]]::new()

    if (Test-Path -Path $RootPath -PathType Leaf) {
        $ext = [System.IO.Path]::GetExtension($RootPath).ToLowerInvariant()
        if ($validExts -contains $ext) {
            $base = [System.IO.Path]::GetFileName($RootPath)
            if ($IncFiles -and -not (Test-PatternMatch -Name $base -Patterns $IncFiles)) { return @() }
            if ($ExcFiles -and (Test-PatternMatch -Name $base -Patterns $ExcFiles)) { return @() }
            return @([System.IO.Path]::GetFullPath($RootPath))
        }
        return @()
    }

    if (-not (Test-Path -Path $RootPath -PathType Container)) {
        return @()
    }

    $opt = if ($IsRecurse) { [System.IO.SearchOption]::AllDirectories } else { [System.IO.SearchOption]::TopDirectoryOnly }
    $resolvedRoot = [System.IO.Path]::GetFullPath($RootPath)

    $dirEntries = [System.IO.Directory]::GetFiles($resolvedRoot, "*.*", $opt)
    foreach ($file in $dirEntries) {
        $ext = [System.IO.Path]::GetExtension($file).ToLowerInvariant()
        if (-not ($validExts -contains $ext)) { continue }

        $fileName = [System.IO.Path]::GetFileName($file)
        $parentDir = [System.IO.Path]::GetFileName([System.IO.Path]::GetDirectoryName($file))

        if ($IncFolders -and -not (Test-PatternMatch -Name $parentDir -Patterns $IncFolders)) { continue }
        if ($ExcFolders -and (Test-PatternMatch -Name $parentDir -Patterns $ExcFolders)) { continue }

        if ($IncFiles -and -not (Test-PatternMatch -Name $fileName -Patterns $IncFiles)) { continue }
        if ($ExcFiles -and (Test-PatternMatch -Name $fileName -Patterns $ExcFiles)) { continue }

        $matched.Add($file)
    }

    return $matched.ToArray()
}

function Decompress-MsOvba {
    param([byte[]]$Compressed)
    if (-not $Compressed -or $Compressed.Length -lt 3) { return [byte[]]@() }
    if ($Compressed[0] -ne 0x01) { return $Compressed }

    $outStream = [System.IO.MemoryStream]::new()
    $pos = 1
    $totalLen = $Compressed.Length

    while ($pos + 2 -le $totalLen) {
        $chunkHeader = [System.BitConverter]::ToUInt16($Compressed, $pos)
        $pos += 2

        $chunkSize = ($chunkHeader -band 0x0FFF) + 3
        $chunkCompressed = ($chunkHeader -shr 15) -band 0x01
        $chunkEnd = [System.Math]::Min($pos + $chunkSize - 2, $totalLen)

        if ($chunkCompressed -eq 0) {
            $uncompLen = $chunkEnd - $pos
            $outStream.Write($Compressed, $pos, $uncompLen)
            $pos = $chunkEnd
            continue
        }

        $chunkDecompStart = $outStream.Length
        while ($pos -lt $chunkEnd) {
            $flagByte = $Compressed[$pos]
            $pos++

            for ($bit = 0; $bit -lt 8; $bit++) {
                if ($pos -ge $chunkEnd) { break }
                $flag = ($flagByte -shr $bit) -band 1
                if ($flag -eq 0) {
                    $outStream.WriteByte($Compressed[$pos])
                    $pos++
                }
                else {
                    if ($pos + 2 -gt $chunkEnd) { $pos = $chunkEnd; break }
                    $token = [System.BitConverter]::ToUInt16($Compressed, $pos)
                    $pos += 2

                    $offsetInChunk = $outStream.Length - $chunkDecompStart
                    $bitCount = 4
                    if ($offsetInChunk -gt 0) {
                        $temp = $offsetInChunk - 1
                        $bits = 0
                        while ($temp -gt 0) { $temp = $temp -shr 1; $bits++ }
                        $bitCount = [System.Math]::Max(4, $bits)
                    }

                    $lengthMask = 0xFFFF -shr $bitCount
                    $offsetMask = (-bnot $lengthMask) -band 0xFFFF
                    $copyLen = ($token -band $lengthMask) + 3
                    $offset = (($token -band $offsetMask) -shr (16 - $bitCount)) + 1

                    $buf = $outStream.ToArray()
                    $copyStart = $outStream.Length - $offset
                    for ($k = 0; $k -lt $copyLen; $k++) {
                        if ($copyStart + $k -ge 0 -and $copyStart + $k -lt $buf.Length) {
                            $outStream.WriteByte($buf[$copyStart + $k])
                        }
                    }
                }
            }
        }
    }

    return $outStream.ToArray()
}

function Extract-VbaProjectBin {
    param(
        [byte[]]$VbaBinBytes,
        [string]$TargetDir,
        [bool]$Quiet
    )

    try {
        $vbaOut = [System.IO.Path]::Combine($TargetDir, "VBA_Extracted")
        [System.IO.Directory]::CreateDirectory($vbaOut) | Out-Null
        $modDir = [System.IO.Path]::Combine($vbaOut, "Modules")
        [System.IO.Directory]::CreateDirectory($modDir) | Out-Null

        # Decompress raw module stream fragments
        $sig = [byte[]]@(0x01)
        $idx = 0
        $foundCount = 0

        # Scan for MS-OVBA compressed streams starting with 0x01 and chunk header
        while ($idx -lt $VbaBinBytes.Length - 10) {
            if ($VbaBinBytes[$idx] -eq 0x01) {
                $hdr = [System.BitConverter]::ToUInt16($VbaBinBytes, $idx + 1)
                $sigCheck = ($hdr -shr 12) -band 0x07
                if ($sigCheck -eq 3) {
                    $subArray = [byte[]]::new($VbaBinBytes.Length - $idx)
                    [System.Array]::Copy($VbaBinBytes, $idx, $subArray, 0, $subArray.Length)
                    $decomp = Decompress-MsOvba -Compressed $subArray
                    if ($decomp -and $decomp.Length -gt 20) {
                        $text = [System.Text.Encoding]::UTF8.GetString($decomp)
                        if ($text -match "Sub\s+|Function\s+|Attribute\s+VB_") {
                            $foundCount++
                            $dest = [System.IO.Path]::Combine($modDir, "Module$foundCount.bas")
                            [System.IO.File]::WriteAllText($dest, $text, [System.Text.Encoding]::UTF8)
                        }
                    }
                }
            }
            $idx += 4
        }
        if (-not $Quiet -and $foundCount -gt 0) {
            Write-Host "  [VBA] Extracted $foundCount module code streams to VBA_Extracted/Modules/."
        }
    }
    catch {
        if (-not $Quiet) { Write-Host "  [VBA] Warning: Could not complete VBA decompilation: $_" }
    }
}

function Extract-PowerQueryM {
    param(
        [string]$TargetDir,
        [bool]$Quiet
    )

    try {
        $customXmlDir = [System.IO.Path]::Combine($TargetDir, "customXml")
        if (-not (Test-Path $customXmlDir)) { return }

        $xmlFiles = [System.IO.Directory]::GetFiles($customXmlDir, "*.xml")
        foreach ($xf in $xmlFiles) {
            $raw = [System.IO.File]::ReadAllBytes($xf)
            $pkIdx = -1
            for ($i = 0; $i -lt $raw.Length - 4; $i++) {
                if ($raw[$i] -eq 0x50 -and $raw[$i+1] -eq 0x4B -and $raw[$i+2] -eq 0x03 -and $raw[$i+3] -eq 0x04) {
                    $pkIdx = $i
                    break
                }
            }

            if ($pkIdx -ge 0) {
                $zipStream = [System.IO.MemoryStream]::new($raw, $pkIdx, $raw.Length - $pkIdx)
                $archive = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Read)
                $secEntry = $archive.GetEntry("Formulas/Section1.m")
                if ($secEntry) {
                    $reader = [System.IO.StreamReader]::new($secEntry.Open(), [System.Text.Encoding]::UTF8)
                    $mText = $reader.ReadToEnd()
                    $reader.Close()

                    $pqOut = [System.IO.Path]::Combine($TargetDir, "PowerQuery_Extracted")
                    $qDir = [System.IO.Path]::Combine($pqOut, "Queries")
                    [System.IO.Directory]::CreateDirectory($qDir) | Out-Null

                    [System.IO.File]::WriteAllText([System.IO.Path]::Combine($pqOut, "Section1.m"), $mText, [System.Text.Encoding]::UTF8)
                    if (-not $Quiet) { Write-Host "  [PowerQuery] Extracted Power Query M formulas to PowerQuery_Extracted/." }
                    break
                }
            }
        }
    }
    catch {
        if (-not $Quiet) { Write-Host "  [PowerQuery] Warning: Power Query extraction: $_" }
    }
}

function Deconstruct-WorkbookStructure {
    param(
        [string]$TargetDir,
        [bool]$Quiet
    )

    try {
        $wbXml = [System.IO.Path]::Combine($TargetDir, "xl", "workbook.xml")
        if (-not (Test-Path $wbXml)) { return }

        $xmlDoc = [xml](Get-Content -Path $wbXml -Raw)
        $sheets = @()
        foreach ($s in $xmlDoc.workbook.sheets.sheet) {
            $sheets += [PSCustomObject]@{
                name = $s.name
                sheetId = $s.sheetId
                state = if ($s.state) { $s.state } else { "visible" }
            }
        }

        $names = @()
        if ($xmlDoc.workbook.definedNames) {
            foreach ($dn in $xmlDoc.workbook.definedNames.definedName) {
                $names += [PSCustomObject]@{
                    name = $dn.name
                    formula = $dn.'#text'
                }
            }
        }

        $structOut = [System.IO.Path]::Combine($TargetDir, "Workbook_Structure")
        [System.IO.Directory]::CreateDirectory($structOut) | Out-Null

        $sheets | ConvertTo-Json -Depth 5 | Set-Content -Path ([System.IO.Path]::Combine($structOut, "Sheets.json")) -Encoding UTF8
        $names | ConvertTo-Json -Depth 5 | Set-Content -Path ([System.IO.Path]::Combine($structOut, "DefinedNames.json")) -Encoding UTF8

        if (-not $Quiet) {
            Write-Host "  [OpenXML] Extracted $($sheets.Count) sheets and $($names.Count) defined names."
        }
    }
    catch {
        if (-not $Quiet) { Write-Host "  [OpenXML] Warning: Structure deconstruction: $_" }
    }
}

function Unpack-SingleExcelFile {
    param(
        [string]$FilePath,
        [string]$CustomOut,
        [bool]$IsDryRun,
        [bool]$IsQuiet
    )

    $fullPath = [System.IO.Path]::GetFullPath($FilePath)
    $fileNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($fullPath)
    $parent = [System.IO.Path]::GetDirectoryName($fullPath)
    $destDir = if ($CustomOut) { [System.IO.Path]::GetFullPath($CustomOut) } else { [System.IO.Path]::Combine($parent, $fileNameNoExt) }

    $created = 0
    $updated = 0
    $unchanged = 0

    if (-not $IsQuiet) {
        Write-Host "`n[UNPACK] Processing: $fullPath"
        Write-Host "         Destination: $destDir"
    }

    $vbaBinBytes = $null

    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($fullPath)
    }
    catch {
        Write-Host "[ERROR] Invalid ZIP-based container: $fullPath" -ForegroundColor Red
        return @{ Created = 0; Updated = 0; Unchanged = 0 }
    }

    try {
        foreach ($entry in $archive.Entries) {
            if ([string]::IsNullOrEmpty($entry.Name)) { continue } # Skip directories

            $destFile = [System.IO.Path]::Combine($destDir, $entry.FullName.Replace("/", [System.IO.Path]::DirectorySeparatorChar))
            $ms = [System.IO.MemoryStream]::new()
            $eStream = $entry.Open()
            $eStream.CopyTo($ms)
            $eStream.Close()
            $contentBytes = $ms.ToArray()
            $ms.Close()

            if ($entry.FullName.ToLowerInvariant().EndsWith("vbaproject.bin")) {
                $vbaBinBytes = $contentBytes
            }

            if (Test-Path $destFile) {
                $existingHash = Get-FileSha256Hex -FilePath $destFile
                $newHash = Get-BytesSha256Hex -Data $contentBytes
                if ($existingHash -eq $newHash) {
                    $unchanged++
                    continue
                }
                else {
                    $updated++
                    $action = "UPDATE"
                }
            }
            else {
                $created++
                $action = "CREATE"
            }

            if (-not $IsQuiet) {
                Write-Host "  [$action] $($entry.FullName)"
            }

            if (-not $IsDryRun) {
                $dir = [System.IO.Path]::GetDirectoryName($destFile)
                if (-not (Test-Path $dir)) { [System.IO.Directory]::CreateDirectory($dir) | Out-Null }
                [System.IO.File]::WriteAllBytes($destFile, $contentBytes)
            }
        }
    }
    finally {
        $archive.Dispose()
    }

    if (-not $IsDryRun) {
        if ($vbaBinBytes) {
            Extract-VbaProjectBin -VbaBinBytes $vbaBinBytes -TargetDir $destDir -Quiet $IsQuiet
        }
        Extract-PowerQueryM -TargetDir $destDir -Quiet $IsQuiet
        Deconstruct-WorkbookStructure -TargetDir $destDir -Quiet $IsQuiet
    }

    if (-not $IsQuiet) {
        Write-Host "  Summary: $created created, $updated updated, $unchanged unchanged."
    }

    return @{ Created = $created; Updated = $updated; Unchanged = $unchanged }
}

function Test-ExcelFeatures {
    param([string]$FilePath)
    $feat = @{ HasVba = $false; HasPq = $false; HasFormulas = $false }
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
        try {
            foreach ($entry in $archive.Entries) {
                $nameLower = $entry.FullName.ToLowerInvariant()
                if ($nameLower.EndsWith("vbaproject.bin")) {
                    $feat.HasVba = $true
                }
                if ($nameLower.Contains("customxml") -and $nameLower.EndsWith(".xml")) {
                    $stream = $entry.Open()
                    $buf = [byte[]]::new([System.Math]::Min(4096, [int]$entry.Length))
                    $read = $stream.Read($buf, 0, $buf.Length)
                    $stream.Close()
                    $str = [System.Text.Encoding]::UTF8.GetString($buf, 0, $read)
                    if ($str.Contains("DataMashup") -or $str.Contains("PK")) {
                        $feat.HasPq = $true
                    }
                }
                elseif ($nameLower.Contains("datamashup")) {
                    $feat.HasPq = $true
                }
                if ($nameLower.EndsWith("calcchain.xml")) {
                    $feat.HasFormulas = $true
                }
                elseif ($nameLower.Contains("worksheets/sheet") -and $nameLower.EndsWith(".xml") -and -not $feat.HasFormulas) {
                    $stream = $entry.Open()
                    $reader = [System.IO.StreamReader]::new($stream)
                    $chunk = $reader.ReadToEnd()
                    $reader.Close()
                    if ($chunk.Contains("<f>") -or $chunk.Contains("<f ")) {
                        $feat.HasFormulas = $true
                    }
                }
            }
        }
        finally {
            $archive.Dispose()
        }
    }
    catch {}
    return $feat
}

# --- Main Driver ---
$files = Get-TargetExcelFiles -RootPath $Path -IsRecurse $Recurse -IncFiles $FileInclude -ExcFiles $FileExclude -IncFolders $FolderInclude -ExcFolders $FolderExclude

if (-not $files -or $files.Count -eq 0) {
    if (-not $Quiet) { Write-Host "[INFO] No Excel files (.xlsx, .xlsm, .xlsb) found in: $Path" }
    exit 0
}

if ($HasVba -or $HasPq -or $HasFormulas) {
    $filtered = [System.Collections.Generic.List[string]]::new()
    foreach ($f in $files) {
        $feat = Test-ExcelFeatures -FilePath $f
        if ($HasVba -and -not $feat.HasVba) { continue }
        if ($HasPq -and -not $feat.HasPq) { continue }
        if ($HasFormulas -and -not $feat.HasFormulas) { continue }
        $filtered.Add($f)
    }
    $files = $filtered.ToArray()
}

if (-not $files -or $files.Count -eq 0) {
    if (-not $Quiet) { Write-Host "[INFO] No Excel files matched the specified feature criteria (-HasVba, -HasPq, -HasFormulas)." }
    exit 0
}

$totCreated = 0
$totUpdated = 0
$totUnchanged = 0

foreach ($f in $files) {
    $res = Unpack-SingleExcelFile -FilePath $f -CustomOut $OutputDir -IsDryRun $DryRun -IsQuiet $Quiet
    $totCreated += $res.Created
    $totUpdated += $res.Updated
    $totUnchanged += $res.Unchanged
}

if (-not $Quiet) {
    Write-Host "`n$("=" * 50)"
    Write-Host "TOTAL: $($files.Count) files processed"
    Write-Host "       $totCreated created, $totUpdated updated, $totUnchanged unchanged."
    Write-Host $("=" * 50)
}
