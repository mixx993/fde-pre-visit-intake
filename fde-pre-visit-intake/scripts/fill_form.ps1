# Fill the pre-visit form template from answers.json. Windows PowerShell 5.1+, no extra installs.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File fill_form.ps1 -Answers answers.json -Out filled.xlsx
# This file is kept ASCII-only on purpose (PowerShell 5.1 reads BOM-less scripts as ANSI).
param(
    [Parameter(Mandatory = $true)][string]$Answers,
    [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$utf8 = New-Object System.Text.UTF8Encoding($false)

$templateDir = Join-Path $PSScriptRoot '..\template'
$meta = [IO.File]::ReadAllText((Join-Path $templateDir 'fields.json'), $utf8) | ConvertFrom-Json
$where = @{}
foreach ($f in $meta.fields) { $where[$f.id] = $f }

$flat = New-Object System.Collections.ArrayList
function Add-Flat($obj, [string]$prefix) {
    if ($null -eq $obj) { return }
    if ($obj -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $obj.PSObject.Properties) { Add-Flat $p.Value ($prefix + $p.Name + '.') }
    } elseif (($obj -is [System.Array]) -or ($obj -is [System.Collections.IList])) {
        for ($i = 0; $i -lt $obj.Count; $i++) { Add-Flat $obj[$i] ($prefix + $i + '.') }
    } else {
        $s = ([string]$obj).Trim()
        if ($s -ne '') { [void]$flat.Add(@($prefix.TrimEnd('.'), $s)) }
    }
}
$ans = [IO.File]::ReadAllText((Resolve-Path -LiteralPath $Answers).Path, $utf8) | ConvertFrom-Json
Add-Flat $ans ''

$edits = @{}
$unknown = New-Object System.Collections.ArrayList
foreach ($kv in $flat) {
    $key = $kv[0]; $val = $kv[1]
    if ($key.StartsWith('_')) { continue }
    if ($where.ContainsKey($key)) {
        $sheet = [int]$where[$key].sheet_index
        if (-not $edits.ContainsKey($sheet)) { $edits[$sheet] = New-Object System.Collections.ArrayList }
        [void]$edits[$sheet].Add(@($where[$key].cell, $val))
    } else { [void]$unknown.Add($key) }
}

$outPath = [IO.Path]::GetFullPath($Out)
Copy-Item -LiteralPath (Join-Path $templateDir $meta.template) -Destination $outPath -Force
$zip = [IO.Compression.ZipFile]::Open($outPath, [IO.Compression.ZipArchiveMode]::Update)
$filled = 0
$missed = New-Object System.Collections.ArrayList
try {
    foreach ($sheet in $edits.Keys) {
        $name = "xl/worksheets/sheet$sheet.xml"
        $entry = $zip.GetEntry($name)
        $reader = New-Object IO.StreamReader($entry.Open(), $utf8)
        $xml = $reader.ReadToEnd(); $reader.Close()
        foreach ($e in $edits[$sheet]) {
            $cell = $e[0]
            $text = [Security.SecurityElement]::Escape($e[1]).Replace('$', '$$')
            $pattern = '<c r="' + $cell + '"((?: s="\d+")?)[^>]*?(?:/>|>.*?</c>)'
            $rx = New-Object Text.RegularExpressions.Regex($pattern, [Text.RegularExpressions.RegexOptions]::Singleline)
            if ($rx.IsMatch($xml)) {
                $xml = $rx.Replace($xml, '<c r="' + $cell + '"$1 t="inlineStr"><is><t xml:space="preserve">' + $text + '</t></is></c>', 1)
                $filled++
            } else { [void]$missed.Add($cell) }
        }
        $entry.Delete()
        $new = $zip.CreateEntry($name)
        $writer = New-Object IO.StreamWriter($new.Open(), $utf8)
        $writer.Write($xml); $writer.Close()
    }
} finally { $zip.Dispose() }

$result = [ordered]@{ out = $outPath; filled = $filled; unknown_keys = @($unknown); cells_not_found = @($missed) }
$json = $result | ConvertTo-Json -Compress
try { [Console]::OutputEncoding = $utf8 } catch {}
Write-Output $json
if ($missed.Count -gt 0) { exit 1 }
