<#
.SYNOPSIS
  Read-only inventory of a movie library (for example a QNAP NAS share).

.DESCRIPTION
  Walks the movie folder, groups files into movies (DVD VIDEO_TS folders,
  Blu-ray BDMV folders, ISO images and single video files) and writes two
  files to your Documents folder:
    MovieInventory_<date>.csv          one row per movie, opens in Excel
    MovieInventory_<date>_summary.txt  totals by format and genre, problems
  Nothing on the NAS is changed, moved or deleted.

.PARAMETER Root
  The movies folder, e.g. \\NAS\Movies or M:\. Asked for if left out.

.PARAMETER OutDir
  Where to save the report. Defaults to your Documents folder.

.PARAMETER NoOpen
  Don't open the spreadsheet when finished.

.EXAMPLE
  .\MovieInventory.ps1 -Root "\\NAS\Movies"
#>
[CmdletBinding()]
param(
    [string]$Root,
    [string]$OutDir = [Environment]::GetFolderPath('MyDocuments'),
    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'

$VideoExt = @('.mkv', '.mp4', '.m4v', '.avi', '.mov', '.wmv', '.mpg', '.mpeg', '.m2ts', '.mts',
              '.ts', '.divx', '.xvid', '.flv', '.webm', '.ogm', '.rmvb', '.vob')
$ImageExt = @('.jpg', '.jpeg', '.png', '.tbn')
$OldFormats = @('AVI', 'WMV', 'FLV', 'RMVB', 'DIVX', 'XVID', 'OGM')

# NAS system folders (QNAP recycle bin, snapshots, thumbnails) and Windows ones.
$SkipDirNames = @('$RECYCLE.BIN', 'System Volume Information', '#recycle', '.streams', 'lost+found')

# Release-name junk that marks the end of a title, e.g. "Alien.1979.1080p.BluRay".
$TagPattern = '(?i)\b(480p|576p|720p|1080[pi]|2160p|4k|uhd|dvdrip|dvdscr|dvd5|dvd9|dvd|bdrip|brrip|' +
              'bluray|blu-ray|web-?dl|webrip|hdrip|hdtv|x264|x265|h264|h265|hevc|avc|xvid|divx|' +
              'ac3|dts|aac|pal|ntsc|remux|proper|repack|extended|unrated)\b'

# File names that say nothing about the movie, so the folder name is used instead.
$GenericNamePattern = '(?i)^(movie|video|film|main|feature|title ?\d*|vts_\d+_\d+|video_ts|\d{1,3})$'

function Test-SkipDir([string]$name) {
    if ($name.StartsWith('@') -or $name.StartsWith('.@')) { return $true }
    foreach ($s in $SkipDirNames) { if ($name -ieq $s) { return $true } }
    return $false
}

function Get-RelPath([string]$path) {
    if ($path.Length -le $RootFull.Length) { return '' }
    return $path.Substring($RootFull.Length).Trim('\')
}

function Get-Genre([string]$containingDir) {
    $rel = Get-RelPath $containingDir
    if (-not $rel) { return '(root)' }
    return $rel.Split('\')[0]
}

function Get-TitleYear([string]$name) {
    $s = ($name -replace '[._]+', ' ') -replace '\s+', ' '
    $title = $s
    $year = ''
    # Use the last plausible year that isn't at the start, so "2001 A Space Odyssey (1968)"
    # gives 1968 and a movie called "2012" keeps its title.
    $ms = [regex]::Matches($s, '(?<![0-9])(19[1-9][0-9]|20[0-4][0-9])(?![0-9])')
    for ($i = $ms.Count - 1; $i -ge 0; $i--) {
        if ($ms[$i].Index -gt 0) {
            $year = $ms[$i].Value
            $title = $s.Substring(0, $ms[$i].Index)
            break
        }
    }
    $tag = [regex]::Match($title, $TagPattern)
    if ($tag.Success -and $tag.Index -gt 0) { $title = $title.Substring(0, $tag.Index) }
    $title = ($title -replace '\[[^\]]*\]', '') -replace '\{[^}]*\}', ''
    $title = ($title.Trim(' -(['.ToCharArray()) -replace '\s+', ' ').Trim()
    if (-not $title) { $title = $name }
    return [pscustomobject]@{ Title = $title; Year = $year }
}

# True if $fileName belongs to the movie whose file name starts with $prefix,
# e.g. "alien (1979)_sheet.jpg" for "alien (1979)", but not "aliens.jpg" for "alien".
function Test-PrefixMatch([string]$fileName, [string]$prefix) {
    if (-not $fileName.StartsWith($prefix)) { return $false }
    if ($fileName.Length -eq $prefix.Length) { return $true }
    return '._- '.IndexOf($fileName[$prefix.Length]) -ge 0
}

function Get-Art([string[]]$dirs, [string]$prefix) {
    $r = @{ Sheet = $false; Poster = $false; Nfo = $false }
    foreach ($d in $dirs) {
        $k = $d.ToLowerInvariant()
        if ($images.ContainsKey($k)) {
            foreach ($n in $images[$k]) {
                $rest = $n
                if ($prefix) {
                    if (-not (Test-PrefixMatch $n $prefix)) { continue }
                    $rest = $n.Substring($prefix.Length)
                }
                if ($rest -match 'sheet') { $r.Sheet = $true }
                elseif ($rest -notmatch 'fanart|backdrop|background|banner|thumb|logo|clearart') { $r.Poster = $true }
            }
        }
        if ($nfos.ContainsKey($k)) {
            foreach ($n in $nfos[$k]) {
                if (-not $prefix -or (Test-PrefixMatch $n $prefix)) { $r.Nfo = $true }
            }
        }
    }
    return $r
}

function Add-ToIndex([hashtable]$index, [string]$dir, [string]$name) {
    $k = $dir.ToLowerInvariant()
    if (-not $index.ContainsKey($k)) { $index[$k] = New-Object System.Collections.Generic.List[string] }
    $index[$k].Add($name.ToLowerInvariant())
}

function YesNo([bool]$b) { if ($b) { 'Yes' } else { 'No' } }

# ---------------------------------------------------------------- where to scan

if (-not $Root) {
    Write-Host ''
    Write-Host 'Enter the path to your movies folder on the NAS.'
    Write-Host 'Examples:  \\NAS\Movies   or   M:\    (you can also drag the folder onto this window)'
    $Root = Read-Host 'Movies folder'
}
$Root = $Root.Trim().Trim('"')
if (-not $Root -or -not (Test-Path -LiteralPath $Root -PathType Container)) {
    Write-Host "Can't find that folder: $Root" -ForegroundColor Red
    Write-Host 'Check the NAS is on, and that you can open the folder in File Explorer.'
    exit 1
}
$RootDir = New-Object System.IO.DirectoryInfo ((Resolve-Path -LiteralPath $Root).ProviderPath)
$RootFull = $RootDir.FullName.TrimEnd('\')

# ---------------------------------------------------------------- scan

Write-Host ''
Write-Host "Scanning $($RootDir.FullName) ..."
Write-Host '(Read-only. Over the network this can take a few minutes for a big library.)'

$files = New-Object System.Collections.Generic.List[System.IO.FileInfo]
$unreadable = New-Object System.Collections.Generic.List[string]
$stack = New-Object System.Collections.Generic.Stack[System.IO.DirectoryInfo]
$stack.Push($RootDir)
$dirCount = 0
while ($stack.Count -gt 0) {
    $dir = $stack.Pop()
    $dirCount++
    if ($dirCount % 25 -eq 0) {
        Write-Progress -Activity 'Scanning movie folders' -Status "$dirCount folders, $($files.Count) files"
    }
    try {
        foreach ($f in $dir.GetFiles()) { $files.Add($f) }
        foreach ($d in $dir.GetDirectories()) {
            if (Test-SkipDir $d.Name) { continue }
            if ($d.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
            $stack.Push($d)
        }
    } catch {
        $unreadable.Add("$($dir.FullName)  ($($_.Exception.Message))")
    }
}
Write-Progress -Activity 'Scanning movie folders' -Completed

# ---------------------------------------------------------------- group files into movies

$images = @{}
$nfos = @{}
$discs = @{}    # movie folder (lower case) -> DVD or Blu-ray details
$singles = New-Object System.Collections.Generic.List[System.IO.FileInfo]

foreach ($f in $files) {
    $ext = $f.Extension.ToLowerInvariant()
    if ($ImageExt -contains $ext) { Add-ToIndex $images $f.DirectoryName $f.Name; continue }
    if ($ext -eq '.nfo') { Add-ToIndex $nfos $f.DirectoryName $f.Name; continue }

    $discKind = $null
    $bdIdx = $f.FullName.IndexOf('\BDMV\', [System.StringComparison]::OrdinalIgnoreCase)
    if ($f.Name -match '(?i)^(VIDEO_TS|VTS_\d{2}_\d)\.(IFO|BUP|VOB)$') {
        $discKind = 'DVD'
        $inVideoTs = $f.Directory.Name -ieq 'VIDEO_TS'
        if ($inVideoTs -and $f.Directory.Parent) { $movieDir = $f.Directory.Parent.FullName } else { $movieDir = $f.DirectoryName }
    } elseif ($bdIdx -ge 0) {
        $discKind = 'BD'
        $inVideoTs = $true
        $movieDir = $f.FullName.Substring(0, $bdIdx)
    }

    if ($discKind) {
        $k = $movieDir.ToLowerInvariant()
        if (-not $discs.ContainsKey($k)) {
            $discs[$k] = @{ Kind = $discKind; Dir = $movieDir; Bytes = [long]0; Files = 0; HasIfo = $false; Structured = $true }
        }
        $e = $discs[$k]
        $e.Bytes += $f.Length
        $e.Files++
        if ($f.Name -ieq 'VIDEO_TS.IFO') { $e.HasIfo = $true }
        if (-not $inVideoTs) { $e.Structured = $false }
        continue
    }

    if ($ext -eq '.iso' -or $VideoExt -contains $ext) { $singles.Add($f) }
}

# How many movies sit in each folder: artwork in a shared folder must match the file name.
$moviesPerDir = @{}
foreach ($f in $singles) { $k = $f.DirectoryName.ToLowerInvariant(); $moviesPerDir[$k] = 1 + [int]$moviesPerDir[$k] }
foreach ($e in $discs.Values) { $k = $e.Dir.ToLowerInvariant(); $moviesPerDir[$k] = 1 + [int]$moviesPerDir[$k] }

$rows = New-Object System.Collections.Generic.List[object]

foreach ($e in $discs.Values) {
    $notes = New-Object System.Collections.Generic.List[string]
    $dirInfo = New-Object System.IO.DirectoryInfo $e.Dir
    $nameDir = $dirInfo
    if ($dirInfo.Name -match '(?i)^(disc|disk|dvd|cd)\s*[-_]?\s*(\d+)$' -and $dirInfo.Parent) {
        $notes.Add("Disc $($Matches[2]) of a multi-disc set")
        $nameDir = $dirInfo.Parent
    }
    if ($e.Dir.TrimEnd('\') -ieq $RootFull) { $notes.Add('Disc files are at the top of the share') }
    elseif ((Get-RelPath $nameDir.FullName).Split('\').Count -le 1) {
        $notes.Add('Disc files sit directly in what may be a genre folder - title guess unreliable')
    }
    if ($e.Kind -eq 'DVD') {
        $format = 'DVD (VIDEO_TS)'
        if (-not $e.HasIfo) { $notes.Add('No VIDEO_TS.IFO - DVD may be incomplete') }
        if (-not $e.Structured) { $notes.Add('VOB files not inside a VIDEO_TS folder') }
    } else {
        $format = 'Blu-ray (BDMV)'
    }

    $ty = Get-TitleYear $nameDir.Name
    if ($moviesPerDir[$e.Dir.ToLowerInvariant()] -gt 1) { $prefix = $nameDir.Name.ToLowerInvariant() } else { $prefix = '' }
    $art = Get-Art @($e.Dir, (Join-Path $e.Dir 'VIDEO_TS')) $prefix
    $genreDir = if ($nameDir.Parent) { $nameDir.Parent.FullName } else { $RootFull }

    $rows.Add([pscustomobject]@{
        Genre = Get-Genre $genreDir; Title = $ty.Title; Year = $ty.Year; Format = $format
        Files = $e.Files; SizeGB = [math]::Round($e.Bytes / 1GB, 2)
        HasMoviesheet = YesNo $art.Sheet; HasPoster = YesNo $art.Poster; HasNfo = YesNo $art.Nfo
        Notes = $notes; Path = $e.Dir
    })
}

foreach ($f in $singles) {
    $notes = New-Object System.Collections.Generic.List[string]
    $ext = $f.Extension.ToLowerInvariant()
    $base = $f.BaseName
    if ($ext -eq '.iso') { $format = 'ISO image' }
    elseif ($ext -eq '.vob') { $format = 'VOB (loose)'; $notes.Add('Loose VOB file - check what this is') }
    else { $format = $ext.TrimStart('.').ToUpperInvariant() }
    if ($OldFormats -contains $format) { $notes.Add('Older format - may need converting') }

    $ty = Get-TitleYear $base
    $folderRel = Get-RelPath $f.DirectoryName
    $folderIsMovieFolder = $folderRel -and $folderRel.Split('\').Count -ge 2
    if ($folderIsMovieFolder) {
        $fromFolder = Get-TitleYear $f.Directory.Name
        if ($base -match $GenericNamePattern -or (-not $ty.Year -and $fromFolder.Year)) { $ty = $fromFolder }
    }

    if ($base -match '(?i)[\s._-](cd|disc|disk|part|pt)\s*[-_]?\s*\d+$') { $notes.Add('Multi-part file') }
    if ($base -match '(?i)\b(sample|trailer|featurette|behind the scenes|deleted scenes?|extras?)\b') {
        $notes.Add('Looks like an extra, trailer or sample')
    } elseif ($ext -ne '.iso' -and $f.Length -lt 150MB) {
        $notes.Add('Small file (under 150 MB) - extra or sample?')
    }

    if ($moviesPerDir[$f.DirectoryName.ToLowerInvariant()] -gt 1) { $prefix = $base.ToLowerInvariant() } else { $prefix = '' }
    $art = Get-Art @($f.DirectoryName) $prefix

    $rows.Add([pscustomobject]@{
        Genre = Get-Genre $f.DirectoryName; Title = $ty.Title; Year = $ty.Year; Format = $format
        Files = 1; SizeGB = [math]::Round($f.Length / 1GB, 2)
        HasMoviesheet = YesNo $art.Sheet; HasPoster = YesNo $art.Poster; HasNfo = YesNo $art.Nfo
        Notes = $notes; Path = $f.FullName
    })
}

# Possible duplicates: same title and year in more than one place.
$dupGroups = $rows | Group-Object { (($_.Title.ToLowerInvariant()) -replace '[^a-z0-9]', '') + '|' + $_.Year } |
    Where-Object { $_.Count -gt 1 }
foreach ($g in $dupGroups) {
    foreach ($r in $g.Group) { $r.Notes.Add("Possible duplicate ($($g.Count) copies)") }
}

$out = $rows | Sort-Object Genre, Title, Year | Select-Object Genre, Title, Year, Format, Files, SizeGB,
    HasMoviesheet, HasPoster, HasNfo, @{ n = 'Notes'; e = { $_.Notes -join '; ' } }, Path

# ---------------------------------------------------------------- write report

if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$csvPath = Join-Path $OutDir "MovieInventory_$stamp.csv"
$txtPath = Join-Path $OutDir "MovieInventory_${stamp}_summary.txt"
$out | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8

$total = @($out).Count
$totalGB = ($out | Measure-Object SizeGB -Sum).Sum
if (-not $totalGB) { $totalGB = 0 }
function Count-Where($cond) { @($out | Where-Object $cond).Count }

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("Movie inventory of $($RootDir.FullName)")
$lines.Add("Made $(Get-Date -Format 'd MMM yyyy HH:mm'), $dirCount folders scanned")
$lines.Add('')
$lines.Add(('Movies found:        {0}' -f $total))
$lines.Add(('Total size:          {0:N1} GB ({1:N2} TB)' -f $totalGB, ($totalGB / 1024)))
$lines.Add(('With a moviesheet:   {0}' -f (Count-Where { $_.HasMoviesheet -eq 'Yes' })))
$lines.Add(('With a poster:       {0}' -f (Count-Where { $_.HasPoster -eq 'Yes' })))
$lines.Add(('With an NFO file:    {0}' -f (Count-Where { $_.HasNfo -eq 'Yes' })))
$lines.Add(('Possible duplicates: {0}' -f (Count-Where { $_.Notes -like '*duplicate*' })))
$lines.Add(('Rows with notes:     {0}  (filter the Notes column in Excel)' -f (Count-Where { $_.Notes })))
$lines.Add('')
$lines.Add('By format')
foreach ($g in ($out | Group-Object Format | Sort-Object Count -Descending)) {
    $lines.Add(('  {0,-18} {1,6} movies  {2,10:N1} GB' -f $g.Name, $g.Count, ($g.Group | Measure-Object SizeGB -Sum).Sum))
}
$lines.Add('')
$lines.Add('By genre folder')
foreach ($g in ($out | Group-Object Genre | Sort-Object Name)) {
    $lines.Add(('  {0,-30} {1,6}' -f $g.Name, $g.Count))
}
if ($unreadable.Count -gt 0) {
    $lines.Add('')
    $lines.Add("Folders that couldn't be read ($($unreadable.Count))")
    foreach ($u in $unreadable) { $lines.Add("  $u") }
}
$lines | Out-File -LiteralPath $txtPath -Encoding UTF8

Write-Host ''
$lines | ForEach-Object { Write-Host $_ }
Write-Host ''
Write-Host "Spreadsheet: $csvPath" -ForegroundColor Green
Write-Host "Summary:     $txtPath" -ForegroundColor Green

if (-not $NoOpen) {
    try { Invoke-Item -LiteralPath $csvPath } catch { }
}
