<#
.SYNOPSIS
  Reads the current MakeMKV beta key from the MakeMKV forum and registers it.

.DESCRIPTION
  Fetches the "MakeMKV is free while in Beta" announcement, picks out the key
  (it starts with T-), copies it to the clipboard and saves it where MakeMKV
  looks for it (HKCU\Software\MakeMKV, value app_Key). Close MakeMKV first.
  If MakeMKV still says the key is invalid, paste it in by hand:
  Help > Register.
#>
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$url = 'https://forum.makemkv.com/forum/viewtopic.php?t=1053'

Write-Host "Reading $url ..."
try {
    $html = (Invoke-WebRequest -Uri $url -UseBasicParsing -UserAgent 'Mozilla/5.0').Content
} catch {
    Write-Host "Couldn't reach the MakeMKV forum: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Open the link above in a browser, copy the key from the first post,'
    Write-Host 'then paste it into MakeMKV under Help > Register.'
    exit 1
}

$m = [regex]::Match($html, 'T-[A-Za-z0-9@_]{40,}')
if (-not $m.Success) {
    Write-Host "Reached the forum but couldn't find a key on the page." -ForegroundColor Red
    Write-Host 'Open the link above in a browser and copy it by hand.'
    exit 1
}
$key = $m.Value

if (Get-Process -Name makemkv -ErrorAction SilentlyContinue) {
    Write-Host 'MakeMKV is running. Close it and run this again.' -ForegroundColor Yellow
    exit 1
}

New-Item -Path 'HKCU:\Software\MakeMKV' -Force | Out-Null
Set-ItemProperty -Path 'HKCU:\Software\MakeMKV' -Name 'app_Key' -Value $key
try { Set-Clipboard -Value $key } catch { }

Write-Host ''
Write-Host 'Key found and saved:' -ForegroundColor Green
Write-Host $key
Write-Host ''
Write-Host 'It is also on the clipboard. Open MakeMKV now. If it still asks for a key,'
Write-Host 'go to Help > Register and press Ctrl+V.'
