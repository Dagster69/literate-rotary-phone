<#
.SYNOPSIS
  Deletes the original DVD folders of movies that were converted and verified.

.DESCRIPTION
  Reads DvdToMkv_results.csv and, for each movie whose LATEST result is OK, deletes the
  original VIDEO_TS folder. It first shows a plan and waits for you to type DELETE.

  A movie is only included if ALL of these are true:
    - its latest result is OK (CHECK and FAILED are never touched)
    - the original is inside the movies folder (Danes vids)
    - the MKV is inside the MKV Movies folder and the file exists
    - the MKV is at least 70% of the DVD's size
    - the original VIDEO_TS folder still exists

  Only the VIDEO_TS folder is deleted. The movie folder itself is removed only if it is then
  empty, so moviesheets, posters and other files are left alone.

  The plan is saved as DvdToMkv_delete_plan.csv next to the results file, and a log is
  written to Documents. After deleting, empty the NAS recycle bin (@Recycle) to free the space.
#>
[CmdletBinding()]
param(
    [string]$Results = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'DvdToMkv_results.csv'),
    [string[]]$SourceRoots = @('\\NASDE4EB7\Multimedia\Danes vids', 'Z:\Danes vids'),
    [string[]]$OutputRoots = @('\\NASDE4EB7\Multimedia\MKV Movies', 'Z:\MKV Movies')
)

$ErrorActionPreference = 'Stop'

function Test-Under([string]$path, [string[]]$roots) {
    foreach ($r in $roots) {
        $rr = $r.TrimEnd('\') + '\'
        if ($path.StartsWith($rr, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

$logFile = Join-Path ([Environment]::GetFolderPath('MyDocuments')) ("DvdToMkv_delete_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd'))
function Write-Log([string]$msg) {
    try { Add-Content -LiteralPath $logFile -Value ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $msg) } catch { }
}

if (-not (Test-Path -LiteralPath $Results)) {
    Write-Host "Results file not found: $Results" -ForegroundColor Red
    exit 1
}

# Latest result for each original folder (a later retry replaces an earlier failure).
$latest = @{}
foreach ($row in (Import-Csv -LiteralPath $Results)) { $latest[$row.Input] = $row }

Write-Host ''
Write-Host 'Checking each converted movie...'
$plan = New-Object System.Collections.ArrayList
$kept = 0
foreach ($key in $latest.Keys) {
    $row = $latest[$key]
    if ($row.Result -ne 'OK') { continue }
    $why = ''
    $mkvBytes = [double]0
    if (-not (Test-Under $row.Input $SourceRoots)) { $why = 'the original is outside the movies folder' }
    elseif (-not (Test-Under $row.Output $OutputRoots)) { $why = 'the MKV is not in the MKV Movies folder' }
    elseif (-not (Test-Path -LiteralPath $row.Output -PathType Leaf)) { $why = 'the MKV file is missing' }
    elseif (-not (Test-Path -LiteralPath (Join-Path $row.Input 'VIDEO_TS') -PathType Container)) { $why = 'the original is already gone' }
    else {
        $mkvBytes = [double](Get-Item -LiteralPath $row.Output).Length
        $dvdBytes = [double]$row.DvdGB * 1GB
        if ($dvdBytes -gt 0 -and ($mkvBytes / $dvdBytes) -lt 0.7) { $why = 'the MKV is under 70% of the DVD size' }
    }
    if ($why) {
        $kept++
        Write-Host ('  keeping {0}: {1}' -f $row.Name, $why) -ForegroundColor Yellow
        continue
    }
    [void]$plan.Add([pscustomobject]@{
        Genre = $row.Genre; Name = $row.Name; Input = $row.Input; Output = $row.Output
        DvdGB = [double]$row.DvdGB; MkvGB = [math]::Round($mkvBytes / 1GB, 2)
    })
}

if ($plan.Count -eq 0) {
    Write-Host 'Nothing to delete.'
    exit 0
}

$planFile = Join-Path (Split-Path $Results -Parent) 'DvdToMkv_delete_plan.csv'
$plan | Sort-Object Genre, Name | Export-Csv -LiteralPath $planFile -NoTypeInformation -Encoding UTF8

$totalGB = [double]0
foreach ($p in $plan) { $totalGB += $p.DvdGB }

Write-Host ''
Write-Host ('{0} original DVD folders are ready to delete ({1:N0} GB).' -f $plan.Count, $totalGB) -ForegroundColor Cyan
Write-Host ('{0} converted movies are being kept because a check did not pass (listed above).' -f $kept)
Write-Host ''
Write-Host 'By genre:'
foreach ($g in ($plan | Group-Object Genre | Sort-Object Name)) {
    $gb = [double]0
    foreach ($x in $g.Group) { $gb += $x.DvdGB }
    Write-Host ('  {0,-20} {1,4} folders  {2,7:N0} GB' -f $g.Name, $g.Count, $gb)
}
Write-Host ''
Write-Host "The full list is saved here, so you can open it in Excel first:"
Write-Host "  $planFile"
Write-Host ''
Write-Host 'Deleting cannot be undone. The movies must already play properly from the MKV files.' -ForegroundColor Yellow
$ans = Read-Host 'Type DELETE (in capitals) to delete these folders, or just press Enter to cancel'
if ($ans -cne 'DELETE') {
    Write-Host 'Nothing was deleted.'
    exit 0
}

$ok = 0
$fail = 0
$i = 0
foreach ($p in ($plan | Sort-Object Genre, Name)) {
    $i++
    Write-Progress -Activity 'Deleting original DVD folders' -Status ('{0} of {1}: {2}' -f $i, $plan.Count, $p.Name) -PercentComplete ([int](100 * $i / $plan.Count))
    try {
        Remove-Item -LiteralPath (Join-Path $p.Input 'VIDEO_TS') -Recurse -Force
        $left = @(Get-ChildItem -LiteralPath $p.Input -Force -ErrorAction SilentlyContinue)
        if ($left.Count -eq 0) { Remove-Item -LiteralPath $p.Input -Force }
        $ok++
        Write-Log ('DELETED {0}\{1}  ({2})' -f $p.Genre, $p.Name, $p.Input)
    } catch {
        $fail++
        Write-Host ('  could not delete {0}: {1}' -f $p.Name, $_.Exception.Message) -ForegroundColor Red
        Write-Log ('FAILED  {0}\{1}  {2}' -f $p.Genre, $p.Name, $_.Exception.Message)
    }
}
Write-Progress -Activity 'Deleting original DVD folders' -Completed

Write-Host ''
Write-Host ('Deleted {0} folders, {1} failed. Log: {2}' -f $ok, $fail, $logFile) -ForegroundColor Green
Write-Host 'Now empty the NAS recycle bin (the @Recycle folder in the Multimedia share, or in the QNAP page'
Write-Host 'under File Station), otherwise the space is not freed.'
