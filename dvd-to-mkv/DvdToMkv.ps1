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
        if ($status -eq 'Done') { $item.LvItem.ForeColor = [System.Drawing.Color]::FromArgb(110, 210, 140) }
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
    if ($item.Bytes -gt 0) { $pct = [math]::Min(100, [math]::Round($curBytes * 100.0 / $item.Bytes)) } else { $pct = 0 }
    $done = $script:DoneBytes + [math]::Min($curBytes, $item.Bytes)
    $elapsed = ((Get-Date) - $script:BatchStart).TotalSeconds
    if ($done -gt 0 -and $elapsed -gt 20) {
        $left = ($script:TotalBytes - $done) / ($done / $elapsed)
        $eta = Format-Span $left
    } else { $eta = 'working it out...' }
    if ($script:Gui) {
        $script:pbCur.Value = [int]$pct
        if ($script:TotalBytes -gt 0) { $script:pbAll.Value = [int][math]::Min(100, [math]::Round($done * 100.0 / $script:TotalBytes)) }
        $script:lblCur.Text = ('Now: {0}\{1}   {2}%' -f $item.Genre, $item.Name, $pct)
        $script:lblAll.Text = ('Overall: {0} of {1} movies, {2:N0} of {3:N0} GB, about {4} left' -f ($script:DoneCount + 1), $script:TotalCount, ($done / 1GB), ($script:TotalBytes / 1GB), $eta)
    }
}

function Convert-One($item) {
    $outDir = Split-Path $item.Out -Parent
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    $part = $item.Out + '.partial'
    $errFile = Join-Path $env:TEMP ('dvdtomkv_err_{0}.txt' -f [guid]::NewGuid().ToString('N'))
    if (Test-Path -LiteralPath $part) { Remove-Item -LiteralPath $part -Force }

    $free = [DiskInfo]::FreeBytes($outDir.TrimEnd('\') + '\')
    if ($free -ge 0 -and $free -lt ($item.Bytes * 1.1 + 5GB)) {
        return [pscustomobject]@{ Ok = $false; Fatal = $true; Note = 'Not enough free space on the output drive' }
    }

    $argStr = '-y -nostdin -nostats -loglevel warning -f dvdvideo -i "{0}" -map 0 -c copy -bsf:s "noise=drop=floor(pts*tb/21600)" -f matroska "{1}"' -f $item.Input, $part
    $p = Start-Process -FilePath $script:Ffmpeg -ArgumentList $argStr -NoNewWindow -PassThru -RedirectStandardError $errFile
    $null = $p.Handle

    while (-not $p.HasExited) {
        Start-Sleep -Milliseconds 400
        Pump
        if ($script:StopRequested) { try { $p.Kill() } catch { }; break }
        $cur = [long]0
        if (Test-Path -LiteralPath $part) { try { $cur = (Get-Item -LiteralPath $part).Length } catch { } }
        Update-Progress $item $cur
    }
    $p.WaitForExit()

    $tail = ''
    if (Test-Path -LiteralPath $errFile) {
        $tail = (Get-Content -LiteralPath $errFile -Tail 3 -ErrorAction SilentlyContinue) -join ' | '
        Remove-Item -LiteralPath $errFile -Force -ErrorAction SilentlyContinue
    }

    if ($script:StopRequested) {
        Start-Sleep -Milliseconds 500
        Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
        return [pscustomobject]@{ Ok = $false; Fatal = $false; Stopped = $true; Note = 'Stopped' }
    }
    if ($p.ExitCode -ne 0) {
        Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
        return [pscustomobject]@{ Ok = $false; Fatal = $false; Note = "FFmpeg error $($p.ExitCode): $tail" }
    }

    $size = (Get-Item -LiteralPath $part).Length
    $dur = Get-MkvDuration $part
    $aud = Get-MkvAudioCount $part
    $ratio = $size / [math]::Max(1, $item.Bytes)
    $problems = @()
    if ($dur -le 0) { $problems += 'no length found' }
    elseif ($dur -lt 600) { $problems += ('only {0} long' -f (Format-Span $dur)) }
    elseif ($dur -gt 21600) { $problems += ('length is {0} (timestamp fault)' -f (Format-Span $dur)) }
    if ($aud -lt 1) { $problems += 'no audio track' }
    if ($ratio -lt 0.7 -or $ratio -gt 1.3) { $problems += ('size is {0:P0} of the DVD' -f $ratio) }
    if ($problems.Count -gt 0) {
        Remove-Item -LiteralPath $part -Force -ErrorAction SilentlyContinue
        return [pscustomobject]@{ Ok = $false; Fatal = $false; Note = 'Check failed: ' + ($problems -join ', ') }
    }

    Move-Item -LiteralPath $part -Destination $item.Out -Force
    return [pscustomobject]@{ Ok = $true; Fatal = $false; Note = ('{0}, {1:N2} GB, {2} audio tracks' -f (Format-Span $dur), ($size / 1GB), $aud); Duration = $dur }
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
    $ok = 0; $failed = 0

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
        catch { $r = [pscustomobject]@{ Ok = $false; Fatal = $false; Note = 'Error: ' + $_.Exception.Message } }
        $secs = [int]((Get-Date) - $t0).TotalSeconds

        if ($r.Ok) {
            $ok++
            Set-ItemStatus $it 'Done'
            Write-Log ("OK     {0}\{1}  ({2}, took {3}s)" -f $it.Genre, $it.Name, $r.Note, $secs)
        } elseif ($r.PSObject.Properties['Stopped'] -and $r.Stopped) {
            Set-ItemStatus $it 'Waiting'
            Write-Log ("STOP   {0}\{1}" -f $it.Genre, $it.Name)
            break
        } else {
            $failed++
            Set-ItemStatus $it ('FAILED: ' + $r.Note)
            Write-Log ("FAILED {0}\{1}  {2}" -f $it.Genre, $it.Name, $r.Note)
            if ($r.Fatal) { Write-Log 'Stopping: free up space on the output drive and start again.'; break }
        }
        $script:DoneBytes += $it.Bytes
        $script:DoneCount++
    }
    Write-Log ("Finished: {0} converted, {1} failed, log: {2}" -f $ok, $failed, $script:LogFile)
    return [pscustomobject]@{ Ok = $ok; Failed = $failed }
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
    $script:btnStart = New-Button 'Start' 760 140 100 $true
    $script:btnStart.Anchor = 'Top,Right'
    $script:btnStop = New-Button 'Stop' 868 140 100 $false
    $script:btnStop.Anchor = 'Top,Right'
    $script:btnStop.Enabled = $false

    $script:lv = New-Object System.Windows.Forms.ListView
    $script:lv.Location = New-Object System.Drawing.Point(16, 182)
    $script:lv.Size = New-Object System.Drawing.Size(970, 330)
    $script:lv.Anchor = 'Top,Left,Right,Bottom'
    $script:lv.View = 'Details'; $script:lv.CheckBoxes = $true; $script:lv.FullRowSelect = $true
    $script:lv.BackColor = $panel; $script:lv.ForeColor = $fg; $script:lv.BorderStyle = 'None'
    [void]$script:lv.Columns.Add('Genre', 120)
    [void]$script:lv.Columns.Add('Movie', 330)
    [void]$script:lv.Columns.Add('GB', 70)
    [void]$script:lv.Columns.Add('Status', 420)
    $form.Controls.Add($script:lv)

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
        foreach ($lvi in $script:lv.Items) { $s = $lvi.Tag.Status; if ($s -eq 'Waiting' -or $s -like 'FAILED*') { $lvi.Checked = $true } }
    })
    $btnNone.Add_Click({ foreach ($lvi in $script:lv.Items) { $lvi.Checked = $false } })

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
        $script:btnScan.Enabled = $false; $script:btnStart.Enabled = $false; $script:btnStop.Enabled = $true
        $script:Running = $true
        try { $r = Start-Batch $sel $out } finally { $script:Running = $false }
        $script:btnStop.Enabled = $false; $script:btnScan.Enabled = $true; $script:btnStart.Enabled = $true
        $script:pbCur.Value = 0
        $script:lblCur.Text = 'Now: -'
        $script:lblStatus.Text = ('Finished: {0} converted, {1} failed. Failed ones can be ticked and retried.' -f $r.Ok, $r.Failed)
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
