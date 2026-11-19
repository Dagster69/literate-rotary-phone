<#
.SYNOPSIS
  Converts DVD folders (VIDEO_TS) to MKV files with FFmpeg, no re-encoding.

.DESCRIPTION
  Scans a movie folder for DVD folders, shows them in a window, and converts the
  ones you tick. Each MKV keeps all video, audio, subtitle and chapter data. The
  original DVD folders are never touched, moved or deleted.

  - Finished movies are skipped, so it can be stopped and restarted any time.
  - Each file is written as .partial and only renamed once it passes checks
    (sensible length, has audio, size close to the DVD).
  - Subtitle packets with impossible timestamps (over 6 hours) are dropped, which
    stops the 13-hour length problem seen on DVD Cloner copies.
  - A log is saved in your Documents folder.

  Needs FFmpeg 7.1 or newer:  winget install Gyan.FFmpeg

.PARAMETER Source
  The folder to scan, e.g. \\NASDE4EB7\Multimedia\Danes vids

.PARAMETER Output
  Where the MKV files go. The genre folder structure is copied.

.PARAMETER NoGui
  Run in the console without a window: converts everything still waiting.

.PARAMETER Limit
  With -NoGui, stop after this many movies (0 = no limit).
#>
[CmdletBinding()]
param(
    [string]$Source = '\\NASDE4EB7\Multimedia\Danes vids',
    [string]$Output = '\\NASDE4EB7\Multimedia\MKV Movies',
    [switch]$NoGui,
    [int]$Limit = 0
)

$ErrorActionPreference = 'Stop'

if (-not ('DiskInfo' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class DiskInfo {
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool GetDiskFreeSpaceEx(string dir, out ulong avail, out ulong total, out ulong free);
    public static long FreeBytes(string dir) {
        ulong a, t, f;
        if (!GetDiskFreeSpaceEx(dir, out a, out t, out f)) return -1;
        return (long)a;
    }
}
'@
}

$script:Gui = $false
$script:StopRequested = $false
$script:Running = $false
$script:Bulk = $false
$script:FreeGB = -1
$script:GenreNames = @()
$script:ResultsFile = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'DvdToMkv_results.csv'
$script:LogFile = Join-Path ([Environment]::GetFolderPath('MyDocuments')) ("DvdToMkv_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd'))
$script:Ffmpeg = $null
$script:Ffprobe = $null

# ---------------------------------------------------------------- helpers

function Write-Log([string]$msg) {
    $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $msg
    try { Add-Content -LiteralPath $script:LogFile -Value $line } catch { }
    if ($script:Gui) { $script:txtLog.AppendText($line + "`r`n") } else { Write-Host $line }
}

function Pump {
    if ($script:Gui) { [System.Windows.Forms.Application]::DoEvents() }
}

function Set-UiStatus([string]$text) {
    if ($script:Gui) { $script:lblStatus.Text = $text; Pump } else { Write-Host $text }
}

function Format-Span([double]$sec) {
    if ($sec -lt 0 -or [double]::IsNaN($sec) -or [double]::IsInfinity($sec)) { return '...' }
    $ts = [TimeSpan]::FromSeconds($sec)
    if ($ts.TotalHours -ge 1) { return ('{0}h {1:00}m' -f [int][math]::Floor($ts.TotalHours), $ts.Minutes) }
    return ('{0}m {1:00}s' -f $ts.Minutes, $ts.Seconds)
}

function Find-Tool([string]$name) {
    $c = Get-Command $name -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    $root = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
    if (Test-Path -LiteralPath $root) {
        $p = Get-ChildItem -LiteralPath $root -Recurse -Filter "$name.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($p) { return $p.FullName }
    }
    return $null
}

function Test-Tools {
    $script:Ffmpeg = Find-Tool 'ffmpeg'
    $script:Ffprobe = Find-Tool 'ffprobe'
    if (-not $script:Ffmpeg -or -not $script:Ffprobe) {
        return 'FFmpeg was not found. Open Command Prompt and run: winget install Gyan.FFmpeg  (then start this again).'
    }
    $ErrorActionPreference = 'Continue'
    $demuxers = & $script:Ffmpeg -hide_banner -demuxers 2>$null | Out-String
    if ($demuxers -notmatch 'dvdvideo') {
        return 'This FFmpeg cannot read DVD folders (it needs version 7.1 or newer). Run: winget upgrade Gyan.FFmpeg'
    }
    return $null
}

function Get-MkvDuration([string]$file) {
    $ErrorActionPreference = 'Continue'
    try {
        $o = & $script:Ffprobe -v error -f matroska -show_entries format=duration -of 'default=nw=1:nk=1' $file 2>$null
        $o = $o | Select-Object -First 1
        return [double]::Parse(([string]$o).Trim(), [Globalization.CultureInfo]::InvariantCulture)
    } catch { return 0 }
}

function Get-MkvAudioCount([string]$file) {
    $ErrorActionPreference = 'Continue'
    try {
        $o = & $script:Ffprobe -v error -f matroska -select_streams a -show_entries stream=index -of 'csv=p=0' $file 2>$null
        return @($o | Where-Object { $_ }).Count
    } catch { return 0 }
}

function Get-MkvSubCount([string]$file) {
    $ErrorActionPreference = 'Continue'
    try {
        $o = & $script:Ffprobe -v error -f matroska -select_streams s -show_entries stream=index -of 'csv=p=0' $file 2>$null
        return @($o | Where-Object { $_ }).Count
    } catch { return 0 }
}

# Length of the movie as FFmpeg reads it from the DVD's own index files; 0 if it can't tell.
function Get-DvdDuration([string]$folder) {
    $ErrorActionPreference = 'Continue'
    try {
        $o = & $script:Ffprobe -v error -f dvdvideo -show_entries format=duration -of 'default=nw=1:nk=1' $folder 2>$null
        $o = $o | Select-Object -First 1
        $v = [double]::Parse(([string]$o).Trim(), [Globalization.CultureInfo]::InvariantCulture)
        if ($v -gt 0) { return $v } else { return 0 }
    } catch { return 0 }
}

function Get-FreeGB([string]$path) {
    $p = $path.Trim().Trim('"')
    while ($p -and -not (Test-Path -LiteralPath $p)) {
        $parent = Split-Path $p -Parent
        if ($parent -eq $p) { break }
        $p = $parent
    }
    if (-not $p) { return -1 }
    $f = [DiskInfo]::FreeBytes($p.TrimEnd('\') + '\')
    if ($f -lt 0) { return -1 }
    return [math]::Round([double]$f / 1GB)
}

function Get-Prop($o, [string]$name, $default) {
    $p = $o.PSObject.Properties[$name]
    if ($p) { return $p.Value } else { return $default }
}

function Add-Result($it, $r, [string]$outcome, [int]$secs) {
    $row = [pscustomobject]@{
        Time = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Genre = $it.Genre; Name = $it.Name; Result = $outcome
        Note = (Get-Prop $r 'Note' ''); DvdGB = [math]::Round([double]$it.Bytes / 1GB, 2)
        MkvGB = [math]::Round([double](Get-Prop $r 'Size' 0) / 1GB, 2)
        MkvLength = (Format-Span ([double](Get-Prop $r 'Duration' 0)))
        DvdLength = (Format-Span ([double](Get-Prop $r 'SrcDuration' 0)))
        AudioTracks = (Get-Prop $r 'Audio' ''); SubtitleTracks = (Get-Prop $r 'Subs' '')
        Method = (Get-Prop $r 'Method' ''); Seconds = $secs; Input = $it.Input; Output = $it.Out
    }
    try { $row | Export-Csv -LiteralPath $script:ResultsFile -Append -NoTypeInformation -Encoding UTF8 } catch { }
}

function Update-Selection {
    if ($script:Bulk) { return }
    $n = 0; $bytes = [double]0
    foreach ($lvi in $script:lv.Items) {
        $it = $lvi.Tag
        if ($lvi.Checked -and ($it.Status -eq 'Waiting' -or $it.Status -like 'FAILED*')) { $n++; $bytes += [double]$it.Bytes }
    }
    $gb = $bytes / 1GB
    $text = ('Ticked: {0} movies, {1:N0} GB' -f $n, $gb)
    $script:lblSel.ForeColor = [System.Drawing.Color]::FromArgb(230, 232, 238)
    if ($script:FreeGB -ge 0) {
        $text += ('     Free on the output drive: {0:N0} GB' -f $script:FreeGB)
        if ($n -gt 0 -and ($gb * 1.05 + 20) -gt $script:FreeGB) {
            $text += '     NOT ENOUGH ROOM - untick some'
            $script:lblSel.ForeColor = [System.Drawing.Color]::FromArgb(240, 110, 110)
        }
    }
    $script:lblSel.Text = $text
}

# ---------------------------------------------------------------- scanning

function Get-Dvds([string]$src, [string]$out) {
    $src = $src.TrimEnd('\')
    $list = New-Object System.Collections.ArrayList
    $stack = New-Object System.Collections.Generic.Stack[System.IO.DirectoryInfo]
    $stack.Push((New-Object System.IO.DirectoryInfo $src))
    $n = 0
    while ($stack.Count -gt 0) {
        $d = $stack.Pop()
        $n++
        if ($n % 20 -eq 0) { Set-UiStatus "Scanning... $n folders looked at, $($list.Count) DVDs found" }
        try { $subs = $d.GetDirectories() } catch { continue }
        foreach ($s in $subs) {
            if ($s.Name.StartsWith('@') -or $s.Name.StartsWith('.') -or $s.Name -ieq '#recycle') { continue }
            if ($s.Name -ieq 'VIDEO_TS') {
                $bytes = [long]0
                try { foreach ($f in $s.GetFiles()) { $bytes += $f.Length } } catch { }
                $rel = $d.FullName.Substring($src.Length).TrimStart('\')
                if ($rel) { $parts = $rel.Split('\'); $genre = $parts[0] } else { $genre = '(root)' }
                $outFile = Join-Path $out ($rel + '\' + $d.Name + '.mkv')
                $hasIfo = Test-Path -LiteralPath (Join-Path $s.FullName 'VIDEO_TS.IFO')
                if (Test-Path -LiteralPath $outFile) { $status = 'Done' }
                elseif (-not $hasIfo) { $status = 'Skipped: no VIDEO_TS.IFO' }
                elseif ($bytes -lt 500MB) { $status = 'Skipped: too small / empty' }
                else { $status = 'Waiting' }
                [void]$list.Add([pscustomobject]@{
                    Genre = $genre; Name = $d.Name; Rel = $rel; Input = $d.FullName
                    Bytes = $bytes; Out = $outFile; Status = $status; LvItem = $null
                })
            } else {
                $stack.Push($s)
            }
        }
    }
    return @($list)
}

# ---------------------------------------------------------------- converting

function Set-ItemStatus($item, [string]$status) {
    $item.Status = $status
    if ($item.LvItem) {
        $item.LvItem.SubItems[3].Text = $status
        if ($status -like 'Done - CHECK*') { $item.LvItem.ForeColor = [System.Drawing.Color]::FromArgb(240, 160, 70) }
        elseif ($status -like 'Done*') { $item.LvItem.ForeColor = [System.Drawing.Color]::FromArgb(110, 210, 140) }
        elseif ($status -like 'FAILED*') { $item.LvItem.ForeColor = [System.Drawing.Color]::FromArgb(240, 110, 110) }
        elseif ($status -like 'Converting*') {
            $item.LvItem.ForeColor = [System.Drawing.Color]::FromArgb(240, 200, 90)
            $item.LvItem.EnsureVisible()
        }
        elseif ($status -like 'Skipped*') { $item.LvItem.ForeColor = [System.Drawing.Color]::Gray }
    }
    Pump
}

function Update-Progress($item, [long]$curBytes) {
    $pct = 0.0
    if ($item.Bytes -gt 0) { $pct = [math]::Min([double]100, [math]::Round([double]$curBytes * 100.0 / [double]$item.Bytes)) }
    $done = [double]$script:DoneBytes + [math]::Min([double]$curBytes, [double]$item.Bytes)
    $elapsed = ((Get-Date) - $script:BatchStart).TotalSeconds
    if ($done -gt 0 -and $elapsed -gt 20) {
        $left = ($script:TotalBytes - $done) / ($done / $elapsed)
        $eta = Format-Span $left
    } else { $eta = 'working it out...' }
    if ($script:Gui) {
        $script:pbCur.Value = [int]$pct
        if ($script:TotalBytes -gt 0) { $script:pbAll.Value = [int][math]::Min([double]100, [math]::Round($done * 100.0 / [double]$script:TotalBytes)) }
        $script:lblCur.Text = ('Now: {0}\{1}   {2}%' -f $item.Genre, $item.Name, $pct)
        $script:lblAll.Text = ('Overall: {0} of {1} movies, {2:N0} of {3:N0} GB, about {4} left' -f ($script:DoneCount + 1), $script:TotalCount, ($done / 1GB), ($script:TotalBytes / 1GB), $eta)
    }
}

function Invoke-Ffmpeg([string]$argStr, $item, [string]$part) {
    $errFile = Join-Path $env:TEMP ('dvdtomkv_err_{0}.txt' -f [guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $part) { Remove-Item -LiteralPath $part -Force }
    $p = Start-Process -FilePath $script:Ffmpeg -ArgumentList $argStr -NoNewWindow -PassThru -RedirectStandardError $errFile
    $null = $p.Handle

    # Watchdog: if the output file has not grown for 5 minutes, FFmpeg is hung (for example on an
    # unreadable file on the NAS), so stop it and let the caller treat it as a failure.
    $lastSize = [long]-1
    $lastChange = Get-Date
    $stalled = $false
    while (-not $p.HasExited) {
        Start-Sleep -Milliseconds 400
        Pump
        if ($script:StopRequested) { try { $p.Kill() } catch { }; break }
        $cur = [long]0
        if (Test-Path -LiteralPath $part) { try { $cur = (Get-Item -LiteralPath $part).Length } catch { } }
        Update-Progress $item $cur
        if ($cur -ne $lastSize) {
            $lastSize = $cur
            $lastChange = Get-Date
        } elseif (((Get-Date) - $lastChange).TotalSeconds -gt 300) {
            $stalled = $true
            try { $p.Kill() } catch { }
            break
        }
    }
    [void]$p.WaitForExit(15000)

    $tail = ''
    if (Test-Path -LiteralPath $errFile) {
        $tail = (Get-Content -LiteralPath $errFile -Tail 3 -ErrorAction SilentlyContinue) -join ' | '
        Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue
    }
    $code = -1
    if ($p.HasExited) { $code = $p.ExitCode }
    return [pscustomobject]@{ ExitCode = $code; Stopped = [bool]$script:StopRequested; Stalled = $stalled; Tail = $tail }
}

# Checks a finished .partial file and compares its length with the DVD's own length.
function Test-Mkv([string]$part, $item) {
    $size = (Get-Item -LiteralPath $part).Length
    $dur = Get-MkvDuration $part
    $aud = Get-MkvAudioCount $part
    $subs = Get-MkvSubCount $part
    $ratio = [double]$size / [math]::Max([double]1, [double]$item.Bytes)
    $problems = @()
    if ($dur -le 0) { $problems += 'no length found' }
    elseif ($dur -lt 600) { $problems += ('only {0} long' -f (Format-Span $dur)) }
    elseif ($dur -gt 21600) { $problems += ('length is {0} (timestamp fault)' -f (Format-Span $dur)) }
    if ($aud -lt 1) { $problems += 'no audio track' }
    if ($ratio -lt 0.7 -or $ratio -gt 1.3) { $problems += ('size is {0:P0} of the DVD' -f $ratio) }

    $srcDur = Get-DvdDuration $item.Input
    $short = $false
    $flag = ''
    if ($srcDur -gt 0 -and $dur -gt 0) {
        $diff = $dur - $srcDur
        $tol = [math]::Max([double]90, $srcDur * 0.03)
        if ([math]::Abs($diff) -gt $tol) {
            $flag = ('MKV is {0}, DVD says {1}' -f (Format-Span $dur), (Format-Span $srcDur))
            if ($diff -lt 0) { $short = $true }
        }
    }
    return [pscustomobject]@{ Size = $size; Dur = $dur; Aud = $aud; Subs = $subs; Problems = $problems; SrcDur = $srcDur; Flag = $flag; Short = $short }
}

# Builds the FFmpeg arguments for the fallback: join the VOB files of the main title set and
# read them as a plain video stream, ignoring the DVD index files. Returns $null if there are none.
function Get-VobArgs($item, [string]$part) {
    $ErrorActionPreference = 'Continue'
    $dir = Join-Path $item.Input 'VIDEO_TS'
    if (-not (Test-Path -LiteralPath $dir)) { return $null }

    $sets = @{}
    foreach ($f in (Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -match '^VTS_(\d\d)_([1-9]\d*)\.VOB$') {
            $k = $Matches[1]
            if (-not $sets.ContainsKey($k)) { $sets[$k] = New-Object System.Collections.ArrayList }
            [void]$sets[$k].Add([pscustomobject]@{ N = [int]$Matches[2]; Path = $f.FullName; Len = $f.Length })
        }
    }
    if ($sets.Count -eq 0) { return $null }

    $best = $null
    $bestLen = [long]-1
    foreach ($k in $sets.Keys) {
        $len = [long]0
        foreach ($v in $sets[$k]) { $len += $v.Len }
        if ($len -gt $bestLen) { $bestLen = $len; $best = $sets[$k] }
    }
    $paths = @($best | Sort-Object N | ForEach-Object { $_.Path })
    $concat = 'concat:' + ($paths -join '|')

    # Put the audio tracks in DVD stream order (0x80, 0x81, ...) and keep the first as default.
    $audio = @()
    try {
        $json = (& $script:Ffprobe -v error -probesize 200M -analyzeduration 200M -show_entries stream=index,codec_type,codec_name,id -of json $concat 2>$null | Out-String) | ConvertFrom-Json
        foreach ($s in $json.streams) {
            if ($s.codec_type -eq 'audio') {
                $sid = 0
                try { $sid = [Convert]::ToInt32(([string]$s.id).Replace('0x', ''), 16) } catch { }
                $audio += [pscustomobject]@{ Index = [int]$s.index; Id = $sid; Codec = [string]$s.codec_name }
            }
        }
    } catch { }

    $maps = '-map 0:v:0'
    $over = ''
    $disp = ''
    if ($audio.Count -gt 0) {
        $k = 0
        foreach ($a in ($audio | Sort-Object Id, Index)) {
            $maps += (' -map 0:{0}' -f $a.Index)
            if ($a.Codec -eq 'pcm_dvd') { $over += (' -c:a:{0} pcm_s24le' -f $k) }
            if ($k -eq 0) { $disp += ' -disposition:a:0 default' } else { $disp += (' -disposition:a:{0} 0' -f $k) }
            $k++
        }
    } else {
        $maps += ' -map 0:a'
    }
    $maps += ' -map 0:s?'

    return ('-y -nostdin -nostats -loglevel warning -fflags +genpts -probesize 200M -analyzeduration 200M -i "{0}" {1} -c copy{2}{3} -bsf:s "noise=drop=floor(pts*tb/21600)" -f matroska "{4}"' -f $concat, $maps, $over, $disp, $part)
}

function Convert-One($item) {
    $outDir = Split-Path $item.Out -Parent
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    $part = $item.Out + '.partial'

    $free = [DiskInfo]::FreeBytes($outDir.TrimEnd('\') + '\')
    if ($free -ge 0 -and $free -lt ($item.Bytes * 1.1 + 5GB)) {
        return [pscustomobject]@{ Ok = $false; Fatal = $true; Note = 'Not enough free space on the output drive' }
    }

    # First try: let FFmpeg read the DVD through its own navigation (keeps chapters and language labels).
    $dvdArgs = '-y -nostdin -nostats -loglevel warning -f dvdvideo -i "{0}" -map 0 -c copy -bsf:s "noise=drop=floor(pts*tb/21600)" -f matroska "{1}"' -f $item.Input, $part
    $run = Invoke-Ffmpeg $dvdArgs $item $part
    if ($run.Stopped) {
        Start-Sleep -Milliseconds 500
        Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
        return [pscustomobject]@{ Ok = $false; Fatal = $false; Stopped = $true; Note = 'Stopped' }
    }

    $why = ''
    $method = 'DVD'
    $res = $null
    if ($run.Stalled) { $why = 'Stalled: no progress for 5 minutes' }
    elseif ($run.ExitCode -ne 0) { $why = "FFmpeg error $($run.ExitCode): $($run.Tail)" }
    elseif (-not (Test-Path -LiteralPath $part)) { $why = 'FFmpeg produced no file' }
    else {
        $res = Test-Mkv $part $item
        if ($res.Problems.Count -gt 0) { $why = 'Check failed: ' + ($res.Problems -join ', ') }
        elseif ($res.Short) { $why = 'Too short: ' + $res.Flag }
    }

    if ($why) {
        Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
        Write-Log ('       {0}: normal method failed ({1}). Trying the direct VOB method...' -f $item.Name, $why)
        $vobArgs = Get-VobArgs $item $part
        if (-not $vobArgs) {
            return [pscustomobject]@{ Ok = $false; Fatal = $false; Note = $why + ' (no VOB files for the fallback)' }
        }
        $run = Invoke-Ffmpeg $vobArgs $item $part
        if ($run.Stopped) {
            Start-Sleep -Milliseconds 500
            Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
            return [pscustomobject]@{ Ok = $false; Fatal = $false; Stopped = $true; Note = 'Stopped' }
        }
        if ($run.Stalled) {
            Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
            return [pscustomobject]@{ Ok = $false; Fatal = $false; Note = "$why; VOB method also stalled: no progress for 5 minutes (file may be unreadable on the NAS)" }
        }
        if ($run.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $part)) {
            Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
            return [pscustomobject]@{ Ok = $false; Fatal = $false; Note = "$why; VOB method also failed: FFmpeg error $($run.ExitCode): $($run.Tail)" }
        }
        $res = Test-Mkv $part $item
        if ($res.Problems.Count -gt 0) {
            Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
            return [pscustomobject]@{ Ok = $false; Fatal = $false; Note = "$why; VOB method: Check failed: " + ($res.Problems -join ', ') }
        }
        $method = 'VOB'
    }

    Move-Item -LiteralPath $part -Destination $item.Out -Force
    return [pscustomobject]@{
        Ok = $true; Fatal = $false; Flag = $res.Flag; Method = $method
        Note = ('{0}, {1:N2} GB, {2} audio, {3} subtitle tracks, {4} method' -f (Format-Span $res.Dur), ($res.Size / 1GB), $res.Aud, $res.Subs, $method)
        Duration = $res.Dur; SrcDuration = $res.SrcDur; Size = $res.Size; Audio = $res.Aud; Subs = $res.Subs
    }
}

function Start-Batch($items, [string]$outRoot) {
    $script:StopRequested = $false
    $todo = @($items)
    $script:TotalCount = $todo.Count
    $script:TotalBytes = [long]0
    foreach ($t in $todo) { $script:TotalBytes += $t.Bytes }
    $script:DoneBytes = [long]0
    $script:DoneCount = 0
    $script:BatchStart = Get-Date
    $ok = 0; $failed = 0; $flagged = 0

    Write-Log ("Starting: {0} movies, {1:N0} GB, output {2}" -f $todo.Count, ($script:TotalBytes / 1GB), $outRoot)
    foreach ($it in $todo) {
        if ($script:StopRequested) { break }
        $it.Out = Join-Path $outRoot ($it.Rel + '\' + $it.Name + '.mkv')
        if (Test-Path -LiteralPath $it.Out) {
            Set-ItemStatus $it 'Done'
            $script:DoneBytes += $it.Bytes; $script:DoneCount++
            continue
        }
        Set-ItemStatus $it 'Converting...'
        Write-Log ("START  {0}\{1}" -f $it.Genre, $it.Name)
        $t0 = Get-Date
        try { $r = Convert-One $it }
        catch {
            Remove-Item -LiteralPath ($it.Out + '.partial') -Force -ErrorAction SilentlyContinue
            $r = [pscustomobject]@{ Ok = $false; Fatal = $false; Note = 'Error: ' + $_.Exception.Message }
        }
        $secs = [int]((Get-Date) - $t0).TotalSeconds

        if ($r.Ok) {
            $ok++
            if ($r.Flag) {
                $flagged++
                Set-ItemStatus $it ('Done - CHECK: ' + $r.Flag)
                Write-Log ("CHECK  {0}\{1}  {2}  ({3})" -f $it.Genre, $it.Name, $r.Flag, $r.Note)
                Add-Result $it $r 'CHECK' $secs
            } else {
                Set-ItemStatus $it 'Done'
                Write-Log ("OK     {0}\{1}  ({2}, took {3}s)" -f $it.Genre, $it.Name, $r.Note, $secs)
                Add-Result $it $r 'OK' $secs
            }
        } elseif ($r.PSObject.Properties['Stopped'] -and $r.Stopped) {
            Set-ItemStatus $it 'Waiting'
            Write-Log ("STOP   {0}\{1}" -f $it.Genre, $it.Name)
            break
        } else {
            $failed++
            Set-ItemStatus $it ('FAILED: ' + $r.Note)
            Write-Log ("FAILED {0}\{1}  {2}" -f $it.Genre, $it.Name, $r.Note)
            Add-Result $it $r 'FAILED' $secs
            if ($r.Fatal) { Write-Log 'Stopping: free up space on the output drive and start again.'; break }
        }
        $script:DoneBytes += $it.Bytes
        $script:DoneCount++
    }
    Write-Log ("Finished: {0} converted ({1} to check), {2} failed. Results: {3}" -f $ok, $flagged, $failed, $script:ResultsFile)
    return [pscustomobject]@{ Ok = $ok; Failed = $failed; Flagged = $flagged }
}

# ---------------------------------------------------------------- window

function Show-Gui {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $script:Gui = $true

    $bg = [System.Drawing.Color]::FromArgb(24, 26, 32)
    $panel = [System.Drawing.Color]::FromArgb(36, 39, 48)
    $fg = [System.Drawing.Color]::FromArgb(230, 232, 238)
    $accent = [System.Drawing.Color]::FromArgb(64, 132, 245)
    $font = New-Object System.Drawing.Font('Segoe UI', 10)

    $script:Items = @()

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'DVD to MKV'
    $form.Size = New-Object System.Drawing.Size(1020, 760)
    $form.MinimumSize = New-Object System.Drawing.Size(860, 600)
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = $bg
    $form.ForeColor = $fg
    $form.Font = $font

    function New-Label($text, $x, $y, $w) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x, $y); $l.Size = New-Object System.Drawing.Size($w, 24)
        $l.ForeColor = $fg; $l.BackColor = [System.Drawing.Color]::Transparent
        $form.Controls.Add($l); return $l
    }
    function New-Box($text, $x, $y, $w) {
        $t = New-Object System.Windows.Forms.TextBox
        $t.Text = $text; $t.Location = New-Object System.Drawing.Point($x, $y); $t.Size = New-Object System.Drawing.Size($w, 26)
        $t.BackColor = $panel; $t.ForeColor = $fg; $t.BorderStyle = 'FixedSingle'; $t.Anchor = 'Top,Left,Right'
        $form.Controls.Add($t); return $t
    }
    function New-Button($text, $x, $y, $w, $primary) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text; $b.Location = New-Object System.Drawing.Point($x, $y); $b.Size = New-Object System.Drawing.Size($w, 30)
        $b.FlatStyle = 'Flat'; $b.FlatAppearance.BorderSize = 0; $b.ForeColor = [System.Drawing.Color]::White
        if ($primary) { $b.BackColor = $accent } else { $b.BackColor = [System.Drawing.Color]::FromArgb(70, 75, 90) }
        $form.Controls.Add($b); return $b
    }

    [void](New-Label 'Movies folder (DVD folders)' 16 14 300)
    $script:txtSrc = New-Box $Source 16 38 800
    $btnSrc = New-Button 'Browse' 828 36 80 $false
    $btnSrc.Anchor = 'Top,Right'
    [void](New-Label 'Save MKV files to' 16 74 300)
    $script:txtOut = New-Box $Output 16 98 800
    $btnOut = New-Button 'Browse' 828 96 80 $false
    $btnOut.Anchor = 'Top,Right'

    $script:btnScan = New-Button 'Scan' 16 140 90 $true
    $btnAll = New-Button 'Tick all' 116 140 90 $false
    $btnNone = New-Button 'Untick all' 216 140 100 $false
    $script:cmbGenre = New-Object System.Windows.Forms.ComboBox
    $script:cmbGenre.Location = New-Object System.Drawing.Point(16, 182)
    $script:cmbGenre.Size = New-Object System.Drawing.Size(300, 28)
    $script:cmbGenre.DropDownStyle = 'DropDownList'
    $script:cmbGenre.BackColor = $panel; $script:cmbGenre.ForeColor = $fg; $script:cmbGenre.FlatStyle = 'Flat'
    $form.Controls.Add($script:cmbGenre)
    $btnGenre = New-Button 'Tick genre' 326 180 100 $false
    $btnRetry = New-Button 'Tick retries' 436 180 110 $false
    $script:btnStart = New-Button 'Start' 760 140 100 $true
    $script:btnStart.Anchor = 'Top,Right'
    $script:btnStop = New-Button 'Stop' 868 140 100 $false
    $script:btnStop.Anchor = 'Top,Right'
    $script:btnStop.Enabled = $false

    $script:lv = New-Object System.Windows.Forms.ListView
    $script:lv.Location = New-Object System.Drawing.Point(16, 220)
    $script:lv.Size = New-Object System.Drawing.Size(970, 262)
    $script:lv.Anchor = 'Top,Left,Right,Bottom'
    $script:lv.View = 'Details'; $script:lv.CheckBoxes = $true; $script:lv.FullRowSelect = $true
    $script:lv.BackColor = $panel; $script:lv.ForeColor = $fg; $script:lv.BorderStyle = 'None'
    [void]$script:lv.Columns.Add('Genre', 120)
    [void]$script:lv.Columns.Add('Movie', 330)
    [void]$script:lv.Columns.Add('GB', 70)
    [void]$script:lv.Columns.Add('Status', 420)
    $form.Controls.Add($script:lv)

    $script:lblSel = New-Label 'Ticked: 0 movies' 16 488 970
    $script:lblSel.Anchor = 'Bottom,Left,Right'

    $script:lblStatus = New-Label 'Press Scan to look for DVDs.' 16 522 970
    $script:lblStatus.Anchor = 'Bottom,Left,Right'
    $script:lblCur = New-Label 'Now: -' 16 550 970
    $script:lblCur.Anchor = 'Bottom,Left,Right'
    $script:pbCur = New-Object System.Windows.Forms.ProgressBar
    $script:pbCur.Location = New-Object System.Drawing.Point(16, 576); $script:pbCur.Size = New-Object System.Drawing.Size(970, 16)
    $script:pbCur.Anchor = 'Bottom,Left,Right'
    $form.Controls.Add($script:pbCur)
    $script:lblAll = New-Label 'Overall: -' 16 598 970
    $script:lblAll.Anchor = 'Bottom,Left,Right'
    $script:pbAll = New-Object System.Windows.Forms.ProgressBar
    $script:pbAll.Location = New-Object System.Drawing.Point(16, 624); $script:pbAll.Size = New-Object System.Drawing.Size(970, 16)
    $script:pbAll.Anchor = 'Bottom,Left,Right'
    $form.Controls.Add($script:pbAll)

    $script:txtLog = New-Object System.Windows.Forms.TextBox
    $script:txtLog.Location = New-Object System.Drawing.Point(16, 650); $script:txtLog.Size = New-Object System.Drawing.Size(970, 60)
    $script:txtLog.Multiline = $true; $script:txtLog.ReadOnly = $true; $script:txtLog.ScrollBars = 'Vertical'
    $script:txtLog.BackColor = $panel; $script:txtLog.ForeColor = [System.Drawing.Color]::FromArgb(170, 175, 190)
    $script:txtLog.BorderStyle = 'None'; $script:txtLog.Anchor = 'Bottom,Left,Right'
    $form.Controls.Add($script:txtLog)

    $script:browse = {
        param($box)
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($box.Text -and (Test-Path -LiteralPath $box.Text)) { $d.SelectedPath = $box.Text }
        if ($d.ShowDialog() -eq 'OK') { $box.Text = $d.SelectedPath }
    }
    $btnSrc.Add_Click({ & $script:browse $script:txtSrc })
    $btnOut.Add_Click({ & $script:browse $script:txtOut })

    $btnAll.Add_Click({
        $script:Bulk = $true
        foreach ($lvi in $script:lv.Items) { $s = $lvi.Tag.Status; if ($s -eq 'Waiting' -or $s -like 'FAILED*') { $lvi.Checked = $true } }
        $script:Bulk = $false
        Update-Selection
    })
    $btnNone.Add_Click({
        $script:Bulk = $true
        foreach ($lvi in $script:lv.Items) { $lvi.Checked = $false }
        $script:Bulk = $false
        Update-Selection
    })
    $btnGenre.Add_Click({
        $i = $script:cmbGenre.SelectedIndex
        if ($i -lt 0) { return }
        $g = $script:GenreNames[$i]
        $script:Bulk = $true
        foreach ($lvi in $script:lv.Items) {
            $it = $lvi.Tag
            if ($it.Genre -eq $g -and ($it.Status -eq 'Waiting' -or $it.Status -like 'FAILED*')) { $lvi.Checked = $true }
        }
        $script:Bulk = $false
        Update-Selection
    })
    # Ticks movies whose latest result in DvdToMkv_results.csv was FAILED or CHECK and that have no MKV now.
    $btnRetry.Add_Click({
        if (-not (Test-Path -LiteralPath $script:ResultsFile)) {
            [void][System.Windows.Forms.MessageBox]::Show('There is no results file yet.', 'DVD to MKV'); return
        }
        $latest = @{}
        foreach ($row in (Import-Csv -LiteralPath $script:ResultsFile)) { $latest[($row.Genre + '|' + $row.Name)] = $row.Result }
        $script:Bulk = $true
        $n = 0
        foreach ($lvi in $script:lv.Items) {
            $it = $lvi.Tag
            $res = $latest[($it.Genre + '|' + $it.Name)]
            if (($res -eq 'FAILED' -or $res -eq 'CHECK') -and ($it.Status -eq 'Waiting' -or $it.Status -like 'FAILED*')) { $lvi.Checked = $true; $n++ }
        }
        $script:Bulk = $false
        Update-Selection
        $script:lblStatus.Text = ('{0} movies that failed or were flagged earlier are now ticked.' -f $n)
    })
    $script:lv.Add_ItemChecked({ Update-Selection })
    $script:txtOut.Add_Leave({
        $script:FreeGB = Get-FreeGB $script:txtOut.Text
        Update-Selection
    })

    $script:btnScan.Add_Click({
        $src = $script:txtSrc.Text.Trim().Trim('"')
        $out = $script:txtOut.Text.Trim().Trim('"')
        if (-not (Test-Path -LiteralPath $src -PathType Container)) {
            [void][System.Windows.Forms.MessageBox]::Show("Can't open the movies folder:`r`n$src", 'DVD to MKV'); return
        }
        $problem = Test-Tools
        if ($problem) { [void][System.Windows.Forms.MessageBox]::Show($problem, 'DVD to MKV'); return }
        $script:btnScan.Enabled = $false; $script:btnStart.Enabled = $false
        $script:lv.Items.Clear()
        $script:Items = @(Get-Dvds $src $out | Sort-Object Genre, Name)
        $script:lv.BeginUpdate()
        foreach ($it in $script:Items) {
            $lvi = New-Object System.Windows.Forms.ListViewItem($it.Genre)
            [void]$lvi.SubItems.Add($it.Name)
            [void]$lvi.SubItems.Add(('{0:N2}' -f ($it.Bytes / 1GB)))
            [void]$lvi.SubItems.Add($it.Status)
            $lvi.Tag = $it
            $lvi.Checked = $false
            $it.LvItem = $lvi
            [void]$script:lv.Items.Add($lvi)
            Set-ItemStatus $it $it.Status
        }
        $script:lv.EndUpdate()
        $script:cmbGenre.Items.Clear()
        $script:GenreNames = @()
        foreach ($g in ($script:Items | Group-Object Genre | Sort-Object Name)) {
            $w = @($g.Group | Where-Object { $_.Status -eq 'Waiting' })
            if ($w.Count -eq 0) { continue }
            $b = [double]0
            foreach ($x in $w) { $b += [double]$x.Bytes }
            [void]$script:cmbGenre.Items.Add(('{0}  ({1} movies, {2:N0} GB)' -f $g.Name, $w.Count, ($b / 1GB)))
            $script:GenreNames += $g.Name
        }
        if ($script:cmbGenre.Items.Count -gt 0) { $script:cmbGenre.SelectedIndex = 0 }
        $script:FreeGB = Get-FreeGB $out
        Update-Selection
        $waiting = @($script:Items | Where-Object { $_.Status -eq 'Waiting' })
        $gb = 0; foreach ($w in $waiting) { $gb += $w.Bytes }
        $nDone = @($script:Items | Where-Object { $_.Status -eq 'Done' }).Count
        $nSkip = @($script:Items | Where-Object { $_.Status -like 'Skipped*' }).Count
        $script:lblStatus.Text = ('{0} DVDs found: {1} waiting ({2:N0} GB), {3} done, {4} skipped. Tick a few and press Start.' -f $script:Items.Count, $waiting.Count, ($gb / 1GB), $nDone, $nSkip)
        Write-Log ('Scanned {0}: {1} DVDs, {2} waiting' -f $src, $script:Items.Count, $waiting.Count)
        $script:btnScan.Enabled = $true; $script:btnStart.Enabled = $true
    })

    $script:btnStart.Add_Click({
        $sel = @()
        foreach ($lvi in $script:lv.Items) {
            if ($lvi.Checked) { $it = $lvi.Tag; if ($it.Status -eq 'Waiting' -or $it.Status -like 'FAILED*') { $sel += $it } }
        }
        if ($sel.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show('Tick at least one movie that is Waiting.', 'DVD to MKV'); return
        }
        $out = $script:txtOut.Text.Trim().Trim('"')
        if (-not $out) { [void][System.Windows.Forms.MessageBox]::Show('Choose where to save the MKV files.', 'DVD to MKV'); return }
        $selBytes = [double]0
        foreach ($s in $sel) { $selBytes += [double]$s.Bytes }
        $free = Get-FreeGB $out
        if ($free -ge 0 -and (($selBytes / 1GB) * 1.05 + 20) -gt $free) {
            $msg = ('You ticked about {0:N0} GB but the output drive has {1:N0} GB free. It will stop when it runs out of room. Continue anyway?' -f ($selBytes / 1GB), $free)
            $ans = [System.Windows.Forms.MessageBox]::Show($msg, 'DVD to MKV', 'YesNo', 'Warning')
            if ($ans -ne 'Yes') { return }
        }
        $script:btnScan.Enabled = $false; $script:btnStart.Enabled = $false; $script:btnStop.Enabled = $true
        $script:Running = $true
        try { $r = Start-Batch $sel $out } finally { $script:Running = $false }
        $script:btnStop.Enabled = $false; $script:btnScan.Enabled = $true; $script:btnStart.Enabled = $true
        $script:pbCur.Value = 0
        $script:pbAll.Value = 0
        $script:lblCur.Text = 'Now: -'
        $script:lblAll.Text = 'Overall: -'
        $script:FreeGB = Get-FreeGB $out
        Update-Selection
        $script:lblStatus.Text = ('Finished: {0} converted ({1} marked CHECK), {2} failed. Details are in DvdToMkv_results.csv in Documents.' -f $r.Ok, $r.Flagged, $r.Failed)
    })

    $script:btnStop.Add_Click({
        $script:StopRequested = $true
        $script:btnStop.Enabled = $false
        $script:lblStatus.Text = 'Stopping after the current step...'
    })

    $form.Add_FormClosing({
        if ($script:Running) {
            $args[1].Cancel = $true
            $script:StopRequested = $true
            $script:lblStatus.Text = 'Stopping... close the window again once it has stopped.'
        }
    })
    [void]$form.ShowDialog()
}

# ---------------------------------------------------------------- go

if ($NoGui) {
    $problem = Test-Tools
    if ($problem) { Write-Host $problem -ForegroundColor Red; exit 1 }
    $items = @(Get-Dvds $Source $Output | Where-Object { $_.Status -eq 'Waiting' } | Sort-Object Genre, Name)
    if ($Limit -gt 0) { $items = @($items | Select-Object -First $Limit) }
    Write-Host ("{0} movies to convert" -f $items.Count)
    if ($items.Count -gt 0) { [void](Start-Batch $items $Output) }
} else {
    Show-Gui
}
