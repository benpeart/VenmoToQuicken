[CmdletBinding()]
param(
    [string]$InputCsv = $null,
    [string]$OutputCsv = $null,
    [string]$Account = "Venmo",
    [string]$DateFormat = "MM/dd/yyyy"
)

# Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Accept first raw argument if -InputCsv wasn't bound or is empty/whitespace
if ([string]::IsNullOrWhiteSpace($InputCsv) -and $args.Count -gt 0) {
    $InputCsv = $args[0]
}

# File picker if no InputCsv provided after all fallbacks
if ([string]::IsNullOrWhiteSpace($InputCsv)) {
    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $dlg.InitialDirectory = (Get-Location).Path
        if ($dlg.ShowDialog() -eq "OK") {
            $InputCsv = $dlg.FileName
        } else { exit 1 }
    } catch {
        $InputCsv = Read-Host "Enter Venmo CSV path"
        if ([string]::IsNullOrWhiteSpace($InputCsv)) { exit 1 }
    }
}

if (-not (Test-Path -LiteralPath $InputCsv)) {
    throw "InputCsv file not found: $InputCsv"
}

if (-not $OutputCsv) {
    $dir  = Split-Path -Path $InputCsv -Parent
    $base = [IO.Path]::GetFileNameWithoutExtension($InputCsv)
    $OutputCsv = Join-Path $dir ($base + "_for_Quicken.csv")
}

function Parse-Date {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { throw "Missing Datetime value." }
    [datetime]::Parse($s, [System.Globalization.CultureInfo]::InvariantCulture)
}

function Parse-Amount {
    param([string]$raw)
    if ([string]::IsNullOrWhiteSpace($raw)) { throw "Missing Amount (total) value." }
    $sign = 1
    $trim = $raw.Trim()
    if ($trim.StartsWith('-')) { $sign = -1 }
    $clean = ($trim -replace '[\+\-$\s,]', '')
    [decimal]$val = [decimal]::Parse($clean, [System.Globalization.CultureInfo]::InvariantCulture)
    $val * $sign
}

# Read and normalize headers
$raw = Get-Content -LiteralPath $InputCsv -Raw -Encoding UTF8
$lines = $raw -split "`r?`n"

$headerRegex = '^(?:\s*,)?ID,Datetime,Type,Status,Note,From,To,Amount \(total\)'
$headerMatch = $lines | Select-String -Pattern $headerRegex | Select-Object -First 1
if (-not $headerMatch) { throw "Could not find transaction header row in $InputCsv" }
$headerIndex = $headerMatch.LineNumber - 1

$headerLine = $lines[$headerIndex]
$headerCells = $headerLine -split ','
for ($i = 0; $i -lt $headerCells.Length; $i++) {
    if ([string]::IsNullOrWhiteSpace($headerCells[$i])) {
        $headerCells[$i] = "Ignore$($i+1)"
    }
}
$lines[$headerIndex] = ($headerCells -join ',')

# Parse CSV
$payload = ($lines[$headerIndex..($lines.Length - 1)] -join "`n")
$rows = $payload | ConvertFrom-Csv
if (-not $rows) { throw "No rows parsed" }

$out = New-Object System.Collections.Generic.List[object]
[int]$skippedBalance = 0

foreach ($r in $rows) {
    $propsNonEmpty = @($r.PSObject.Properties | Where-Object {
        $_.Value -and $_.Value.ToString().Trim() -ne ""
    })
    $nonEmptyCount = $propsNonEmpty.Count
    $ignorePopulatedCount = @($propsNonEmpty | Where-Object { $_.Name -like 'Ignore*' }).Count

    if ($nonEmptyCount -eq 0) { continue }
    if ($nonEmptyCount -eq 1 -and $ignorePopulatedCount -eq 1) { $skippedBalance++; continue }
    if ($nonEmptyCount -eq 1 -and $r.'Amount (total)') { $skippedBalance++; continue }
    if (-not $r.Datetime) { continue }

    $dt  = Parse-Date ([string]$r.Datetime)
    $amt = Parse-Amount ([string]$r.'Amount (total)')

    # Determine payee based on amount and available fields
    $payee = $null
    if     ($amt -lt 0 -and $r.To)   { $payee = [string]$r.To }
    elseif ($amt -ge 0 -and $r.From) { $payee = [string]$r.From }
    if (-not $payee) {
        if     ($r.From) { $payee = [string]$r.From }
        elseif ($r.To)   { $payee = [string]$r.To }
        elseif ($r.Note) { $payee = [string]$r.Note }
        else             { $payee = "Venmo" }
    }

    $memoParts = @()
    if ($r.Note)              { $memoParts += "Note: $($r.Note)" }
    if ($r.Type)              { $memoParts += "Type: $($r.Type)" }
    if ($r.Status)            { $memoParts += "Status: $($r.Status)" }
    if ($r.From)              { $memoParts += "From: $($r.From)" }
    if ($r.To)                { $memoParts += "To: $($r.To)" }
    if ($r.'Amount (fee)')    { $memoParts += "Fee: $($r.'Amount (fee)')" }
    if ($r.'Funding Source')  { $memoParts += "Funding: $($r.'Funding Source')" }
    if ($r.Destination)       { $memoParts += "Destination: $($r.Destination)" }
    $memo = $memoParts -join ' | '

    $out.Add([pscustomobject]@{
        'Date'         = $dt.ToString($DateFormat, [System.Globalization.CultureInfo]::InvariantCulture)
        'Payee'        = $payee
        'FI Payee'     = ''                        # blank
        'Amount'       = ('{0:N2}' -f $amt)
        'Debit/Credit' = ''                        # blank
        'Category'     = ''                        # blank
        'Account'      = $Account
        'Tag'          = ''                        # blank
        'Memo'         = $memo
        'Chknum'       = ''                        # blank        
    }) | Out-Null
}

$encoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$out | Export-Csv -LiteralPath $OutputCsv -NoTypeInformation -Encoding $encoding

Write-Host ("Done. Wrote {0} transactions to {1}" -f $out.Count, $OutputCsv) -ForegroundColor Green
if ($skippedBalance -gt 0) {
    Write-Host ("Skipped {0} balance summary line(s)" -f $skippedBalance) -ForegroundColor Yellow
}
