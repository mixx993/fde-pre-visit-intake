# Read-only snapshot of this Windows PC for the pre-visit form. Changes nothing, sends nothing.
# Reads: Windows version, CPU, memory, graphics, disk space, installed programs (names and
# versions from the registry), programs running right now (name and folder), network adapters,
# and whether python/ffmpeg/node exist. Optional -ScanFolder: count and sizes of video files only.
# It does NOT read documents, chat history, browser data, passwords or account files.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File collect_env.ps1 -Out env.json [-ScanFolder D:\videos]
# This file is kept ASCII-only on purpose (PowerShell 5.1 reads BOM-less scripts as ANSI).
param(
    [Parameter(Mandatory = $true)][string]$Out,
    [string]$ScanFolder = ''
)
$ErrorActionPreference = 'SilentlyContinue'
$utf8 = New-Object System.Text.UTF8Encoding($false)

$os = Get-CimInstance Win32_OperatingSystem
$cs = Get-CimInstance Win32_ComputerSystem
$info = [ordered]@{
    collected_at = (Get-Date).ToString('yyyy-MM-dd HH:mm')
    os = [ordered]@{ name = $os.Caption; version = $os.Version; build = $os.BuildNumber; arch = $os.OSArchitecture }
    cpu = @(Get-CimInstance Win32_Processor | ForEach-Object { $_.Name.Trim() })
    memory_gb = [math]::Round($cs.TotalPhysicalMemory / 1GB)
    gpu = @(Get-CimInstance Win32_VideoController | ForEach-Object { [ordered]@{ name = $_.Name; driver = $_.DriverVersion } })
    disks = @(Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | ForEach-Object {
        [ordered]@{ drive = $_.DeviceID; size_gb = [math]::Round($_.Size / 1GB); free_gb = [math]::Round($_.FreeSpace / 1GB) } })
    network = @(Get-CimInstance Win32_NetworkAdapter -Filter 'NetEnabled=True' | ForEach-Object {
        [ordered]@{ name = $_.NetConnectionID; adapter = $_.Name; link_mbps = [math]::Round($_.Speed / 1e6) } })
    tools = [ordered]@{
        python = [bool](Get-Command python -ErrorAction SilentlyContinue)
        ffmpeg = [bool](Get-Command ffmpeg -ErrorAction SilentlyContinue)
        node   = [bool](Get-Command node -ErrorAction SilentlyContinue)
    }
}

$keys = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
$info.installed = @(Get-ItemProperty $keys | Where-Object { $_.DisplayName -and -not $_.SystemComponent } |
    Sort-Object DisplayName -Unique | ForEach-Object {
        [ordered]@{ name = $_.DisplayName; version = $_.DisplayVersion; publisher = $_.Publisher } })

# Running programs outside the Windows folder: catches "green" (unzip-and-run) software
# that never appears in the installed list.
$winDir = $env:WINDIR.ToLower()
$info.running = @(Get-Process | Where-Object { $_.Path -and -not $_.Path.ToLower().StartsWith($winDir) } |
    Group-Object Path | ForEach-Object {
        $p = $_.Group[0]
        $ver = ''
        try { $ver = $p.MainModule.FileVersionInfo.ProductVersion } catch {}
        [ordered]@{ name = $p.ProcessName; folder = (Split-Path $p.Path); version = $ver; count = $_.Count } } |
    Sort-Object { $_.name })

if ($ScanFolder -ne '' -and (Test-Path -LiteralPath $ScanFolder)) {
    $videos = @(Get-ChildItem -LiteralPath $ScanFolder -File | Where-Object { $_.Extension -match '^\.(mp4|mov|mkv|avi|flv)$' } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 30)
    $info.scan = [ordered]@{
        folder = $ScanFolder
        video_count = $videos.Count
        size_mb_recent = @($videos | ForEach-Object { [math]::Round($_.Length / 1MB) })
    }
}

$json = $info | ConvertTo-Json -Depth 5
[IO.File]::WriteAllText([IO.Path]::GetFullPath($Out), $json, $utf8)
Write-Output ("saved: " + [IO.Path]::GetFullPath($Out))
