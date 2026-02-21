<#
.SYNOPSIS
    Chase UK PDF -> CSV Extractor

.DESCRIPTION
    Parses Chase (UK) PDF statements using Poppler `pdftotext`, extracts
    transaction rows and balances, classifies transactions, and exports:
      - Full transactions CSV
      - Monthly summary CSV
      - UK tax-year summary CSV (includes TotalInterest)
      - Calendar-year summary CSV

.REQUIREMENTS
    - Poppler `pdftotext` installed and the path assigned to `$popplerPath`.
    - Run in PowerShell with appropriate ExecutionPolicy.

.AUTHOR
    Craig Horsfield

.VERSION
    1.2.0

.CHANGELOG
    2026-02-21 v1.0   Initial stable extractor (basic parsing + CSV exports)
    2026-02-21 v1.1   Relaxed parsing rules; fixed date parsing issues
    2026-02-21 v1.2   Added `TotalInterest` to tax-year summary and console synopsis

.NOTES
    See `$popplerPath` variable near the top of this file to configure `pdftotext`.
#>

$popplerPath = "C:\Program Files\Poppler\poppler-25.12.0\Library\bin\pdftotext.exe"

if (-not (Test-Path $popplerPath)) {
    throw "pdftotext not found at '$popplerPath'. Update the path if needed."
}

# Select folder containing bank PDFs (default to G:\My Drive\Statements)
$defaultBrowse = 'G:\My Drive\Statements'
if (-not (Test-Path $defaultBrowse)) { $defaultBrowse = 0 }
$folder = (New-Object -ComObject Shell.Application).BrowseForFolder(0, "Select folder containing bank PDF statements", 0, $defaultBrowse)
if (-not $folder) { throw "No folder selected." }
$rootPath = $folder.Self.Path

# Combined outputs go into the central statements folder (use default if available)
$combinedRoot = 'G:\My Drive\Statements'
if (-not (Test-Path $combinedRoot)) { $combinedRoot = $rootPath }

# Timestamped filenames
$stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$outAllCsv          = Join-Path $combinedRoot "Bank_AllAccounts_$stamp.csv"
$outMonthlyCsv      = Join-Path $combinedRoot "Bank_Summary_Monthly_$stamp.csv"
$outTaxYearCsv      = Join-Path $combinedRoot "Bank_Summary_TaxYear_$stamp.csv"
$outCalendarYearCsv = Join-Path $combinedRoot "Bank_Summary_CalendarYear_$stamp.csv"

# Convert money string to decimal
function Convert-ToAmount {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $neg = $s.Trim().StartsWith("-")
    $clean = $s -replace "[^0-9\.]", ""
    if (-not [decimal]::TryParse($clean, [ref]$null)) { return $null }
    $val = [decimal]$clean
    if ($neg) { $val = -$val }
    return $val
}

# Hybrid categorisation
function Get-Category {
    param([string]$description, [string]$transactionType)
    $d = $description.ToLower()

    if ($d -match "netflix") { return "Subscriptions" }
    if ($d -match "shell") { return "Transport/Fuel" }
    if ($d -match "leeds city council") { return "Parking/Council" }
    if ($d -match "mercadona|supermercado") { return "Groceries" }
    if ($d -match "café balzar|cafe balzar") { return "Food & Drink" }
    if ($d -match "hotel|benidorm centre") { return "Accommodation" }
    if ($d -match "easyjet") { return "Travel" }
    if ($d -match "world duty free") { return "Shopping" }
    if ($d -match "bar|pub|flamingo|jolly roger|burning bar|winners fun") { return "Bars & Nightlife" }
    if ($d -match "barbers") { return "Personal Care" }

    switch ($transactionType) {
        "Cash Withdrawal" { return "Cash Withdrawal" }
        "Transfer" { return "Transfer" }
        "Payment" { return "Payment" }
        default { return "Other" }
    }
}

# Transaction type detection
function Get-TransactionType {
    param([string]$description, [decimal]$amountGBP)
    $d = $description.ToLower()

    if ($d -match "cash withdrawal") { return "Cash Withdrawal" }
    if ($d -match "^from " -or $d -match "^to ") { return "Transfer" }
    if ($d -match "payment") { return "Payment" }
    if ($amountGBP -gt 0) { return "Transfer" }
    return "Purchase"
}

# UK tax year
function Get-UkTaxYear {
    param([datetime]$date)
    $year = $date.Year
    $taxYearStart = Get-Date -Year $year -Month 4 -Day 6
    if ($date -lt $taxYearStart) { $startYear = $year - 1 } else { $startYear = $year }
    $endYearShort = ($startYear + 1).ToString().Substring(2)
    return "{0}-{1}" -f $startYear, $endYearShort
}

# Robust date parsing helper
function Get-ParsedDate {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return Get-Date }
    $formats = @('d MMM yyyy','dd MMM yyyy','d/MMM/yyyy','dd/MMM/yyyy','d MMM yy','dd MMM yy')
    foreach ($f in $formats) {
        try {
            return [datetime]::ParseExact($s, $f, $null)
        } catch { }
    }
    try { return [datetime]::Parse($s) } catch { return Get-Date }
}

$allTransactions = @()
$skippedPdfs = @()
$processedCount = 0

# Process PDFs
Get-ChildItem -Path $rootPath -Filter *.pdf | ForEach-Object {
    $pdf = $_.FullName
    Write-Host "Processing $pdf"

    $tempTxt = [System.IO.Path]::GetTempFileName()

    # Force UTF-8 output
    & $popplerPath -layout -enc UTF-8 $pdf $tempTxt 2>$null

    $lines = Get-Content -LiteralPath $tempTxt -Encoding UTF8

    # Extract header info
    $accountNumber = $null
    $sortCode      = $null
    $accountName   = $null

    foreach ($line in $lines) {
        if (-not $accountNumber -and $line -match "Account number:\s*(\d+)") { $accountNumber = $matches[1] }
        if (-not $sortCode -and $line -match "Sort code:\s*([0-9\-]+)") { $sortCode = $matches[1] }
        if (-not $accountName -and $line -match "(.+?) statement") { $accountName = $matches[1].Trim() }
    }

    # Detect bank by scanning page text or filename
    $bankName = 'Unknown'
    if ($lines -match '(?i)lloyd') { $bankName = 'Lloyds' }
    elseif ($lines -match '(?i)chase') { $bankName = 'Chase' }
    elseif ($pdf -match '(?i)lloyd') { $bankName = 'Lloyds' }
    elseif ($pdf -match '(?i)chase') { $bankName = 'Chase' }

    if (-not $accountNumber) { $accountNumber = "Unknown" }
    if (-not $sortCode) { $sortCode = "N/A" }
    if (-not $accountName) { $accountName = [System.IO.Path]::GetFileNameWithoutExtension($pdf) }

    # Find transaction table (try several common header patterns)
    $headerPatterns = @(
        '^Date\s+Transaction',
        '^Date\s+Transaction details',
        '^Date\s+Details',
        '^Date\s+Description',
        '^Date\s+Narrative',
        '^Date\s+\-'
    )
    $startIndex = $null
    foreach ($pat in $headerPatterns) {
        $m = $lines | Select-String -Pattern $pat | Select-Object -First 1
        if ($m) { $startIndex = $m.LineNumber; break }
    }
    if (-not $startIndex) {
        Write-Host ("Header not found in {0}; falling back to scanning entire file" -f $pdf) -ForegroundColor Yellow
        $startIndex = 0
    } else {
        if ($startIndex -gt 1) { $startIndex-- } else { $startIndex = 0 }
    }

    $endIndex = ($lines | Select-String -Pattern "^Some useful information" | Select-Object -First 1).LineNumber
    if ($endIndex) { $endIndex -= 2 } else { $endIndex = $lines.Count - 1 }

    $txLines = $lines[$startIndex..$endIndex]

    $currentTx = $null

    for ($i = 0; $i -lt $txLines.Count; $i++) {
        $line = $txLines[$i].TrimEnd()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # Skip non-date lines (allow 1-2 digit day, 2- or 4-digit year)
        if ($line -notmatch "^\d{1,2}\s+\w{3}\s+\d{2,4}") { continue }

        # Match date + rest (allow 1 or 2 digit day, 2- or 4-digit year)
        if ($line -match "^(?<Date>\d{1,2}\s+\w{3}\s+\d{2,4})(?<Rest>.*)$") {
            $dateStr = $matches['Date'].Trim()
            $rest    = $matches['Rest'].Trim()

            # Skip lines that have no amount-like token (be currency-agnostic)
            # allow odd leading chars (e.g. Â£) by looking for a number with decimals
            if ($rest -notmatch "[^\d\-+]*[-+]?\d[\d,]*\.\d{2}") { continue }

            # Flush previous
            if ($currentTx) { $allTransactions += $currentTx }

            # Parse date robustly
            $date = Get-ParsedDate $dateStr

            # Extract description, amount, balance
            $desc = $rest.Trim()
            $amountStr = $null
            $balanceStr = $null

            if ($rest -match "^(?<Desc>.+?)\s+(?<Amt>[-+]?£?[\d,]+\.\d{2})\s+(?<Bal>£?[\d,]+\.\d{2})$") {
                $desc       = $matches['Desc'].Trim()
                $amountStr  = $matches['Amt']
                $balanceStr = $matches['Bal']
            }
            elseif ($rest -match "^(?<Desc>.+?)\s+(?<Bal>£?[\d,]+\.\d{2})$") {
                $desc       = $matches['Desc'].Trim()
                $balanceStr = $matches['Bal']
            }

            $amountGBP    = Convert-ToAmount $amountStr
            $balanceAfter = Convert-ToAmount $balanceStr
            $txType       = Get-TransactionType -description $desc -amountGBP $amountGBP

            $currentTx = [PSCustomObject]@{
                BankName        = $bankName
                AccountNumber   = $accountNumber
                SortCode        = $sortCode
                AccountName     = $accountName
                Date            = $date
                Description     = $desc
                AmountGBP       = $amountGBP
                AmountFX        = $null
                FXCurrency      = $null
                FXRate          = $null
                BalanceAfter    = $balanceAfter
                TransactionType = $txType
                Category        = $null
                YearMonth       = $date.ToString('yyyy-MM')
                TaxYear         = Get-UkTaxYear $date
                CalendarYear    = $date.Year
                RawText         = $line
            }

            continue
        }
    }

    if ($currentTx) { $allTransactions += $currentTx }

    Remove-Item $tempTxt -Force
    $processedCount++
}

# Categorise
foreach ($tx in $allTransactions) {
    $tx.Category = Get-Category -description $tx.Description -transactionType $tx.TransactionType
}



# Export full CSV
if ($allTransactions.Count -eq 0) {
    Write-Host "No transactions extracted from PDFs." -ForegroundColor Yellow
    # create header-only CSVs so user can see files were created
    $headers = 'BankName,AccountNumber,SortCode,AccountName,Date,Description,AmountGBP,AmountFX,FXCurrency,FXRate,BalanceAfter,TransactionType,Category,RawText'
    $headers | Out-File -FilePath $outAllCsv -Encoding UTF8
    Write-Host "Created header-only file: $outAllCsv" -ForegroundColor Yellow
} else {
    try {
        $allTransactions |
            Select-Object BankName, AccountNumber, SortCode, AccountName,
                          @{Name="Date";Expression={$_.Date.ToString("yyyy-MM-dd")}},
                          Description, AmountGBP, AmountFX, FXCurrency, FXRate,
                          BalanceAfter, TransactionType, Category, RawText |
            Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outAllCsv -Force
        Write-Host "Full transactions exported to: $outAllCsv"
    } catch {
        Write-Host "Failed to export full transactions: $_" -ForegroundColor Red
    }
}

# ===== Summaries =====

# Monthly summary
$monthlySummary = $allTransactions |
    Group-Object BankName, AccountNumber, SortCode, AccountName, YearMonth |
    ForEach-Object {
        $group = $_.Group
        $key   = $_.Name -split ',' | ForEach-Object { $_.Trim() }

        [PSCustomObject]@{
            BankName          = $key[0]
            AccountNumber     = $key[1]
            SortCode          = $key[2]
            AccountName       = $key[3]
            YearMonth         = $key[4]
            TotalSpent        = ($group | Where-Object { $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalReceived     = ($group | Where-Object { $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalFxSpent      = ($group | Where-Object { $_.AmountFX -ne $null -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalCashWithdraw = ($group | Where-Object { $_.TransactionType -eq "Cash Withdrawal" } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersIn  = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersOut = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            NetMovement       = ($group | Measure-Object AmountGBP -Sum).Sum
        }
    }

try {
    if ($monthlySummary -and $monthlySummary.Count -gt 0) {
        $monthlySummary | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outMonthlyCsv -Force
    } else {
        'BankName,AccountNumber,SortCode,AccountName,YearMonth,TotalSpent,TotalReceived,TotalFxSpent,TotalCashWithdraw,TotalTransfersIn,TotalTransfersOut,NetMovement' | Out-File -FilePath $outMonthlyCsv -Encoding UTF8
    }
    Write-Host "Monthly summary exported to: $outMonthlyCsv"
} catch {
    Write-Host "Failed to export monthly summary: $_" -ForegroundColor Red
}

# UK tax-year summary
$taxYearSummary = $allTransactions |
    Group-Object BankName, AccountNumber, SortCode, AccountName, TaxYear |
    ForEach-Object {
        $group = $_.Group
        $key   = $_.Name -split ',' | ForEach-Object { $_.Trim() }

        [PSCustomObject]@{
            BankName          = $key[0]
            AccountNumber     = $key[1]
            SortCode          = $key[2]
            AccountName       = $key[3]
            TaxYear           = $key[4]
            TotalSpent        = ($group | Where-Object { $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalReceived     = ($group | Where-Object { $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalInterest     = ($group | Where-Object { $_.Description -match 'interest' } | Measure-Object AmountGBP -Sum).Sum
            TotalFxSpent      = ($group | Where-Object { $_.AmountFX -ne $null -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalCashWithdraw = ($group | Where-Object { $_.TransactionType -eq "Cash Withdrawal" } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersIn  = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersOut = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            NetMovement       = ($group | Measure-Object AmountGBP -Sum).Sum
        }
    }

# Export tax-year summary with monetary fields formatted to 2 dp to avoid '83.2' style values
$taxYearSummary |
    Select-Object BankName, AccountNumber, SortCode, AccountName, TaxYear,
        @{Name='TotalSpent';Expression={ if ($_.TotalSpent -ne $null) { '{0:F2}' -f [decimal]$_.TotalSpent } else { '' } }},
        @{Name='TotalReceived';Expression={ if ($_.TotalReceived -ne $null) { '{0:F2}' -f [decimal]$_.TotalReceived } else { '' } }},
        @{Name='TotalInterest';Expression={ if ($_.TotalInterest -ne $null) { '{0:F2}' -f [decimal]$_.TotalInterest } else { '' } }},
        @{Name='TotalFxSpent';Expression={ if ($_.TotalFxSpent -ne $null) { '{0:F2}' -f [decimal]$_.TotalFxSpent } else { '' } }},
        @{Name='TotalCashWithdraw';Expression={ if ($_.TotalCashWithdraw -ne $null) { '{0:F2}' -f [decimal]$_.TotalCashWithdraw } else { '' } }},
        @{Name='TotalTransfersIn';Expression={ if ($_.TotalTransfersIn -ne $null) { '{0:F2}' -f [decimal]$_.TotalTransfersIn } else { '' } }},
        @{Name='TotalTransfersOut';Expression={ if ($_.TotalTransfersOut -ne $null) { '{0:F2}' -f [decimal]$_.TotalTransfersOut } else { '' } }},
        @{Name='NetMovement';Expression={ if ($_.NetMovement -ne $null) { '{0:F2}' -f [decimal]$_.NetMovement } else { '' } }} |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outTaxYearCsv
Write-Host "UK tax-year summary exported to: $outTaxYearCsv"
try {
    if ($taxYearSummary -and $taxYearSummary.Count -gt 0) {
        # already exported above with formatting
        Write-Host "Tax-year CSV written: $outTaxYearCsv"
    } else {
        'BankName,AccountNumber,SortCode,AccountName,TaxYear,TotalSpent,TotalReceived,TotalInterest,TotalFxSpent,TotalCashWithdraw,TotalTransfersIn,TotalTransfersOut,NetMovement' | Out-File -FilePath $outTaxYearCsv -Encoding UTF8
        Write-Host "Created header-only tax-year CSV: $outTaxYearCsv" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Failed to ensure tax-year CSV: $_" -ForegroundColor Red
}

# Print overall tax-year summary to screen
$overallByTaxYear = $taxYearSummary | Group-Object TaxYear | Sort-Object Name
Write-Host ""
Write-Host "Overall UK Tax-Year Summary (all accounts):" -ForegroundColor Cyan

# Build a small table in console
$pad = @{TaxYear=10;Interest=12;Received=12;Spent=12;Net=12}
$hdr = ('{0,-10} {1,12} {2,12} {3,12} {4,12}' -f 'TaxYear','Interest','Received','Spent','Net')
Write-Host $hdr
Write-Host ('-' * 62)
foreach ($grp in $overallByTaxYear) {
    $ty = $grp.Name
    $sumSpent = ($grp.Group | Measure-Object TotalSpent -Sum).Sum
    $sumReceived = ($grp.Group | Measure-Object TotalReceived -Sum).Sum
    $sumInterest = ($grp.Group | Measure-Object TotalInterest -Sum).Sum
    $net = ($grp.Group | Measure-Object NetMovement -Sum).Sum
    Write-Host ('{0,-10} {1,12:N2} {2,12:N2} {3,12:N2} {4,12:N2}' -f $ty, $sumInterest, $sumReceived, $sumSpent, $net)
}

# Also write a Markdown report summarising results
$outSummaryMd = Join-Path $rootPath "Bank_Summary_Report_$stamp.md"
$lines = @()
$lines += '# Bank Statements Summary'
$lines += ''
$lines += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
$lines += ''
$lines += '## Overall UK Tax-Year Summary (all accounts)'
$lines += ''
$lines += '| TaxYear | Interest | Received | Spent | Net |'
$lines += '|---:|---:|---:|---:|---:|'

foreach ($grp in $overallByTaxYear) {
    $ty = $grp.Name
    $sumSpent = ($grp.Group | Measure-Object TotalSpent -Sum).Sum
    $sumReceived = ($grp.Group | Measure-Object TotalReceived -Sum).Sum
    $sumInterest = ($grp.Group | Measure-Object TotalInterest -Sum).Sum
    $net = ($grp.Group | Measure-Object NetMovement -Sum).Sum
    $lines += "| $ty | $([math]::Round($sumInterest,2)) | $([math]::Round($sumReceived,2)) | $([math]::Round($sumSpent,2)) | $([math]::Round($net,2)) |"
}

$lines += ''
$lines += '## Per-account Tax-Year Summary'
$lines += ''
$lines += '| Bank | AccountNumber | AccountName | TaxYear | TotalInterest | TotalReceived | TotalSpent | Net |'
$lines += '|---|---|---|---:|---:|---:|---:|---:|'

foreach ($row in $taxYearSummary | Sort-Object BankName,AccountNumber,TaxYear) {
    $lines += "| $($row.BankName) | $($row.AccountNumber) | $($row.AccountName) | $($row.TaxYear) | $([math]::Round($row.TotalInterest,2)) | $([math]::Round($row.TotalReceived,2)) | $([math]::Round($row.TotalSpent,2)) | $([math]::Round($row.NetMovement,2)) |"
}

$lines += ''
$lines += '---'
$lines += 'Generated by ChasePDFParsing.ps1'

$lines | Out-File -FilePath $outSummaryMd -Encoding UTF8
Write-Host "Markdown summary written to: $outSummaryMd" -ForegroundColor Green

# Calendar-year summary
$calendarSummary = $allTransactions |
    Group-Object BankName, AccountNumber, SortCode, AccountName, CalendarYear |
    ForEach-Object {
        $group = $_.Group
        $key   = $_.Name -split ',' | ForEach-Object { $_.Trim() }
        [PSCustomObject]@{
            BankName          = $key[0]
            AccountNumber     = $key[1]
            SortCode          = $key[2]
            AccountName       = $key[3]
            CalendarYear      = $key[4]
            TotalSpent        = ($group | Where-Object { $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalReceived     = ($group | Where-Object { $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalFxSpent      = ($group | Where-Object { $_.AmountFX -ne $null -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalCashWithdraw = ($group | Where-Object { $_.TransactionType -eq "Cash Withdrawal" } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersIn  = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersOut = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            NetMovement       = ($group | Measure-Object AmountGBP -Sum).Sum
        }
    }

try {
    if ($calendarSummary -and $calendarSummary.Count -gt 0) {
        $calendarSummary | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outCalendarYearCsv -Force
    } else {
        'BankName,AccountNumber,SortCode,AccountName,CalendarYear,TotalSpent,TotalReceived,TotalFxSpent,TotalCashWithdraw,TotalTransfersIn,TotalTransfersOut,NetMovement' | Out-File -FilePath $outCalendarYearCsv -Encoding UTF8
    }
    Write-Host "Calendar-year summary exported to: $outCalendarYearCsv"
} catch {
    Write-Host "Failed to export calendar-year summary: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "Processed $processedCount PDFs. Skipped $($skippedPdfs.Count) PDFs." -ForegroundColor Cyan
if ($skippedPdfs.Count -gt 0) {
    Write-Host "Skipped files:" -ForegroundColor Yellow
    $skippedPdfs | ForEach-Object { Write-Host " - $_" }
}

Write-Host "Combined outputs written to: $combinedRoot" -ForegroundColor Green
Write-Host "Done."
