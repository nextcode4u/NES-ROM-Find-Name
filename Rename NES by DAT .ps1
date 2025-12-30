<#
Rename NES by DAT .ps1
ScriptVersion: 2025-12-29i (FAST + RECOMMENDED + END PROMPT)

Features:
- Loads all .dat/.xml in the folder.
- Two-pass CRC per ROM:
    - headered (full file)
    - deheadered (skip 16-byte iNES header if present)
- Optional fallback: strip 4-digit prefix ONLY for unmatched (won't touch "007 GoldenEye")
- Exports:
    - planned-renames.csv
    - unmatched.txt
- End-of-scan prompt:
    - If -Apply not provided, asks to proceed with renames (Y/N).
- Verbose logging + Press ENTER to close.

Usage:
  .\Rename` NES` by` DAT`.ps1
  .\Rename` NES` by` DAT`.ps1 -Strip4DigitPrefixForUnmatched
  .\Rename` NES` by` DAT`.ps1 -Apply
  .\Rename` NES` by` DAT`.ps1 -Apply -Strip4DigitPrefixForUnmatched
#>

param(
    [switch]$Apply,
    [switch]$DebugBytes,
    [switch]$Strip4DigitPrefixForUnmatched,
    [string[]]$Extensions = @(".nes")
)

$ErrorActionPreference = "Stop"
$Root = Get-Location
$LogPath = Join-Path $Root "rename-log.txt"
$UnmatchedPath = Join-Path $Root "unmatched.txt"
$CsvPlanPath = Join-Path $Root "planned-renames.csv"
$ScriptVersion = "2025-12-29i"
$ScriptPath = $PSCommandPath

function Log {
    param([string]$Msg)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
    $line = "[$ts] $Msg"
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}

function Pause-End {
    Write-Host ""
    Read-Host "Done. Press ENTER to close"
}

function HexPreview {
    param([byte[]]$Bytes, [int]$Count = 32)
    if (-not $Bytes) { return "<null>" }
    $n = [Math]::Min($Count, $Bytes.Length)
    return ($Bytes[0..($n-1)] | ForEach-Object { $_.ToString("X2") }) -join " "
}

function Safe-FileName {
    param([string]$Name)
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) {
        $Name = $Name.Replace($c, "_")
    }
    return $Name.Trim()
}

# ----------------------------
# FAST CRC32 (compiled C#)
# ----------------------------
if (-not ("FastCrc32" -as [type])) {
    Add-Type -TypeDefinition @"
using System;

public static class FastCrc32
{
    private static readonly UInt32[] Table = new UInt32[256];

    static FastCrc32()
    {
        const UInt32 poly = 0xEDB88320u;
        for (UInt32 i = 0; i < 256; i++)
        {
            UInt32 c = i;
            for (int j = 0; j < 8; j++)
            {
                c = ((c & 1) != 0) ? (poly ^ (c >> 1)) : (c >> 1);
            }
            Table[i] = c;
        }
    }

    public static UInt32 Compute(byte[] bytes)
    {
        UInt32 crc = 0xFFFFFFFFu;
        for (int i = 0; i < bytes.Length; i++)
        {
            byte b = bytes[i];
            crc = Table[(crc ^ b) & 0xFF] ^ (crc >> 8);
        }
        return crc ^ 0xFFFFFFFFu;
    }
}
"@ | Out-Null
}

function Get-Crc32 {
    param([byte[]]$Bytes)
    $val = [FastCrc32]::Compute($Bytes)
    return ('{0:X8}' -f $val)
}

function CrcSelfTest {
    $bytes = [System.Text.Encoding]::ASCII.GetBytes("123456789")
    $crc = Get-Crc32 $bytes
    Log ("DEBUG CRC SelfTest: CRC32('123456789') = {0} (expected CBF43926)" -f $crc)
    if ($crc -ne "CBF43926") { Log "WARNING: CRC self-test did not match expected value." }
    else { Log "DEBUG CRC self-test passed." }
}

function Has-INesHeader {
    param([byte[]]$Bytes)
    if (-not $Bytes -or $Bytes.Length -lt 16) { return $false }
    return ($Bytes[0] -eq 0x4E -and $Bytes[1] -eq 0x45 -and $Bytes[2] -eq 0x53 -and $Bytes[3] -eq 0x1A)
}

function Strip-4DigitPrefix {
    param([string]$BaseName)
    $m = [regex]::Match($BaseName, '^(?<num>\d{4})[\s\-_\.]+(?<rest>.+)$')
    if ($m.Success) { return $m.Groups["rest"].Value.Trim() }
    return $BaseName
}

function Confirm-YesNo {
    param(
        [string]$Prompt = "Proceed? (Y/N)"
    )
    while ($true) {
        $ans = Read-Host $Prompt
        if (-not $ans) { continue }
        $ans = $ans.Trim().ToUpperInvariant()
        if ($ans -in @("Y","YES")) { return $true }
        if ($ans -in @("N","NO")) { return $false }
        Write-Host "Please enter Y or N."
    }
}

try {
    Set-Content -LiteralPath $LogPath -Value ("=== SCRIPT START {0} ===" -f (Get-Date))
    Set-Content -LiteralPath $UnmatchedPath -Value ""
    if (Test-Path -LiteralPath $CsvPlanPath) { Remove-Item -LiteralPath $CsvPlanPath -Force }

    $psv = $PSVersionTable.PSVersion
    $pse = $PSVersionTable.PSEdition
    $procBits = if ([Environment]::Is64BitProcess) { "64" } else { "32" }
    $osBits   = if ([Environment]::Is64BitOperatingSystem) { "64" } else { "32" }

    Log "ScriptVersion: $ScriptVersion"
    Log "ScriptPath: $ScriptPath"
    Log "Working folder: $Root"
    Log ("PowerShell: {0}  Edition: {1}" -f $psv, $pse)
    Log ("Process bitness: {0}-bit" -f $procBits)
    Log ("OS bitness     : {0}-bit" -f $osBits)
    Log "Apply mode: $Apply"
    Log "DebugBytes: $DebugBytes"
    Log "Strip4DigitPrefixForUnmatched: $Strip4DigitPrefixForUnmatched"
    Log ("Extensions: " + ($Extensions -join ", "))

    CrcSelfTest

    # ----------------------------
    # Load DAT/XML files
    # ----------------------------
    $datFiles = Get-ChildItem -LiteralPath $Root -File |
        Where-Object { $_.Extension -in ".dat", ".xml" } |
        Sort-Object Name

    if (-not $datFiles -or $datFiles.Count -eq 0) {
        Log "ERROR: No .dat or .xml files found in this folder."
        Pause-End
        return
    }

    Log ("Found {0} DAT/XML file(s): {1}" -f $datFiles.Count, (($datFiles | Select-Object -ExpandProperty Name) -join "; "))

    # Build master lookup: CRC -> { Title, SourceDat } (first DAT wins)
    $lookup = @{}

    foreach ($df in $datFiles) {
        Log "Loading DAT: $($df.Name)"
        [xml]$dat = Get-Content -LiteralPath $df.FullName

        $games =
            if ($dat.datafile.game) { $dat.datafile.game }
            elseif ($dat.datafile.machine) { $dat.datafile.machine }
            else {
                Log "WARNING: Unsupported DAT structure in $($df.Name) (no <game> or <machine>). Skipping."
                continue
            }

        $added = 0
        $seen  = 0

        foreach ($g in $games) {
            $title = $g.name
            if (-not $title) { continue }

            foreach ($r in $g.rom) {
                if (-not $r.crc) { continue }
                $seen++
                $crcKey = $r.crc.ToString().ToUpper()
                if (-not $lookup.ContainsKey($crcKey)) {
                    $lookup[$crcKey] = [PSCustomObject]@{ Title = $title; Source = $df.Name }
                    $added++
                }
            }
        }

        Log ("DAT loaded: {0} (rom nodes seen: {1}, new CRC entries added: {2})" -f $df.Name, $seen, $added)
    }

    Log ("Master CRC table size: {0}" -f $lookup.Count)

    # ----------------------------
    # Find ROMs
    # ----------------------------
    $roms = @()
    foreach ($ext in $Extensions) {
        $roms += Get-ChildItem -LiteralPath $Root -File | Where-Object { $_.Extension -ieq $ext }
    }
    $roms = $roms | Sort-Object Name

    if (-not $roms -or $roms.Count -eq 0) {
        Log "No ROM files found for selected extensions."
        Pause-End
        return
    }

    Log ("Found {0} ROM file(s) to process." -f $roms.Count)

    $plan = New-Object System.Collections.Generic.List[object]
    $matched = 0
    $unmatched = 0

    # ----------------------------
    # Process ROMs
    # ----------------------------
    $i = 0
    foreach ($rom in $roms) {
        $i++
        Log "----"
        Log ("[{0}/{1}] Reading: {2}" -f $i, $roms.Count, $rom.Name)

        $bytesFull = [IO.File]::ReadAllBytes($rom.FullName)

        Log ("FileSize: {0} bytes" -f $bytesFull.Length)
        if ($DebugBytes) {
            Log ("DEBUG First32 : {0}" -f (HexPreview -Bytes $bytesFull -Count 32))
        }

        $hasHeader = Has-INesHeader $bytesFull
        Log ("iNES header detected: {0}" -f $hasHeader)

        $crcHeadered = Get-Crc32 $bytesFull
        Log ("CRC32 (headered): {0}" -f $crcHeadered)

        $crcDeheadered = $null
        if ($hasHeader -and $bytesFull.Length -gt 16) {
            $bytesNoHdr = $bytesFull[16..($bytesFull.Length - 1)]
            $crcDeheadered = Get-Crc32 $bytesNoHdr
            Log ("CRC32 (deheadered): {0}" -f $crcDeheadered)
        } else {
            Log "CRC32 (deheadered): <skipped>"
        }

        $matchCrc = $null
        $matchMode = $null
        if ($lookup.ContainsKey($crcHeadered)) {
            $matchCrc = $crcHeadered
            $matchMode = "headered"
        } elseif ($crcDeheadered -and $lookup.ContainsKey($crcDeheadered)) {
            $matchCrc = $crcDeheadered
            $matchMode = "deheadered"
        }

        if ($matchCrc) {
            $entry = $lookup[$matchCrc]
            $newBase = Safe-FileName $entry.Title
            $newName = $newBase + $rom.Extension

            Log ("MATCH ({0}): '{1}' => '{2}' (from {3})" -f $matchMode, $rom.Name, $entry.Title, $entry.Source)

            if ($newName -ne $rom.Name) {
                $plan.Add([PSCustomObject]@{
                    Action      = "DAT_CRC_$matchMode"
                    OldFullPath  = $rom.FullName
                    OldName      = $rom.Name
                    NewName      = $newName
                    CRCMatched   = $matchCrc
                    SourceDat    = $entry.Source
                })
                Log ("Planned rename: {0} -> {1}" -f $rom.Name, $newName)
            } else {
                Log "Already named correctly. Skipping rename."
            }
            $matched++
        }
        else {
            $unmatched++
            Add-Content -LiteralPath $UnmatchedPath -Value $rom.Name
            Log "NO MATCH in loaded DATs."

            if ($Strip4DigitPrefixForUnmatched) {
                $base = [IO.Path]::GetFileNameWithoutExtension($rom.Name)
                $ext  = [IO.Path]::GetExtension($rom.Name)
                $stripped = Strip-4DigitPrefix $base

                if ($stripped -ne $base) {
                    $newName = (Safe-FileName $stripped) + $ext
                    if ($newName -ne $rom.Name) {
                        $plan.Add([PSCustomObject]@{
                            Action      = "STRIP_4DIGIT_PREFIX"
                            OldFullPath  = $rom.FullName
                            OldName      = $rom.Name
                            NewName      = $newName
                            CRCMatched   = ""
                            SourceDat    = ""
                        })
                        Log ("Planned fallback rename (unmatched): {0} -> {1}" -f $rom.Name, $newName)
                    }
                } else {
                    Log "Fallback prefix strip: no 4-digit prefix detected."
                }
            }
        }
    }

    Log "===="
    Log ("Scan complete. Matched: {0}, Unmatched: {1}, Planned renames (total): {2}" -f $matched, $unmatched, $plan.Count)
    Log ("Unmatched list written to: {0}" -f $UnmatchedPath)

    if ($plan.Count -gt 0) {
        Log ("Exporting planned renames CSV: {0}" -f $CsvPlanPath)
        $plan | Select-Object Action, OldName, NewName, CRCMatched, SourceDat |
            Export-Csv -LiteralPath $CsvPlanPath -NoTypeInformation -Encoding UTF8

        Log "Planned renames (table):"
        $plan | Select-Object Action, OldName, NewName, CRCMatched, SourceDat | Format-Table -AutoSize | Out-String |
            ForEach-Object { $_.TrimEnd() } |
            ForEach-Object { if ($_ -ne "") { Log $_ } }
    } else {
        Log "No renames planned."
    }

    # ----------------------------
    # Decide whether to apply
    # ----------------------------
    $doApply = $Apply
    if (-not $doApply -and $plan.Count -gt 0) {
        Write-Host ""
        Write-Host "Planned renames: $($plan.Count)"
        Write-Host "CSV: $CsvPlanPath"
        Write-Host "Unmatched: $UnmatchedPath"
        Write-Host ""
        $doApply = Confirm-YesNo "Rename files now? (Y/N)"
        Log ("User prompt result: doApply={0}" -f $doApply)
    }

    if (-not $doApply) {
        Log "No changes applied."
        Pause-End
        return
    }

    if ($plan.Count -eq 0) {
        Log "Nothing to rename."
        Pause-End
        return
    }

    # ----------------------------
    # Apply renames
    # ----------------------------
    Log "APPLY MODE: Starting renames..."

    foreach ($item in $plan) {
        $oldPath = $item.OldFullPath
        $targetName = $item.NewName
        $targetPath = Join-Path $Root $targetName

        if (Test-Path -LiteralPath $targetPath) {
            $base = [IO.Path]::GetFileNameWithoutExtension($targetName)
            $ext  = [IO.Path]::GetExtension($targetName)
            $n = 2
            do {
                $candidate = "{0} ({1}){2}" -f $base, $n, $ext
                $targetPath = Join-Path $Root $candidate
                $n++
            } while (Test-Path -LiteralPath $targetPath)

            $targetName = [IO.Path]::GetFileName($targetPath)
            Log ("Collision detected; using: {0}" -f $targetName)
        }

        Rename-Item -LiteralPath $oldPath -NewName $targetName
        Log ("RENAMED ({0}): {1} -> {2}" -f $item.Action, $item.OldName, $targetName)
    }

    Log "All renames complete."
    Pause-End
}
catch {
    try { Log ("FATAL ERROR: {0}" -f $_.Exception.Message) } catch { Write-Host ("FATAL ERROR: {0}" -f $_.Exception.Message) }
    Pause-End
}
