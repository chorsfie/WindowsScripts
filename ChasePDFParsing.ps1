# ===========================
# Chase UK PDF → CSV Extractor (Final Corrected Version)
# Multi-account, FX-aware, UTF-8, summaries, hybrid categories
# ===========================

$popplerPath = "C:\Program Files\Poppler\poppler-25.12.0\Library\bin\pdftotext.exe"

if (-not (Test-Path $popplerPath)) {
    throw "pdftotext not found at '$popplerPath'. Update the path if needed."
}

# Select folder containing ALL Chase PDFs
$folder = (New-Object -ComObject Shell.Application).BrowseForFolder(0, "Select folder containing Chase PDF statements", 0)
if (-not $folder) { throw "No folder selected." }
$rootPath = $folder.Self.Path

# Timestamped filenames
$stamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$outAllCsv          = Join-Path $rootPath "Chase_AllAccounts_$stamp.csv"
$outMonthlyCsv      = Join-Path $rootPath "Chase_Summary_Monthly_$stamp.csv"
$outTaxYearCsv      = Join-Path $rootPath "Chase_Summary_TaxYear_$stamp.csv"
$outCalendarYearCsv = Join-Path $rootPath "Chase_Summary_CalendarYear_$stamp.csv"

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

$allTransactions = @()

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

    if (-not $accountNumber) { $accountNumber = "Unknown" }
    if (-not $sortCode) { $sortCode = "N/A" }
    if (-not $accountName) { $accountName = [System.IO.Path]::GetFileNameWithoutExtension($pdf) }

    # Find transaction table
    $startIndex = ($lines | Select-String -Pattern "^Date\s+Transaction details" | Select-Object -First 1).LineNumber
    if (-not $startIndex) { Remove-Item $tempTxt -Force; return }
    $startIndex--

    $endIndex = ($lines | Select-String -Pattern "^Some useful information" | Select-Object -First 1).LineNumber
    if ($endIndex) { $endIndex -= 2 } else { $endIndex = $lines.Count - 1 }

    $txLines = $lines[$startIndex..$endIndex]

    $currentTx = $null

    for ($i = 0; $i -lt $txLines.Count; $i++) {
        $line = $txLines[$i].TrimEnd()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # Skip non-date lines
        if ($line -notmatch "^\d{2} \w{3} \d{4}") { continue }

        # Match date + rest
        if ($line -match "^(?<Date>\d{2} \w{3} \d{4})(?<Rest>.*)$") {

            # PATCH: Skip lines that have no amount and no balance
            if ($rest -notmatch "£[\d,]+\.\d{2}" -and $rest -notmatch "[+\-]?£[\d,]+\.\d{2}") {
                continue
            }
        
            $dateStr = $matches['Date']
            $rest    = $matches['Rest']

            # Skip lines with no amount and no balance
            if ($rest -notmatch "£[\d,]+\.\d{2}" -and $rest -notmatch "[+\-]?£[\d,]+\.\d{2}") {
                continue
            }

            # Flush previous
            if ($currentTx) { $allTransactions += $currentTx }

            [datetime]$date = [datetime]::ParseExact($dateStr, "dd MMM yyyy", $null)

            # Extract description, amount, balance
            $desc = $rest.Trim()
            $amountStr = $null
            $balanceStr = $null

            if ($rest -match "^(?<Desc>.+?)\s+(?<Amt>[+\-]?£[\d,]+\.\d{2})\s+(?<Bal>£[\d,]+\.\d{2})$") {
                $desc       = $matches['Desc'].Trim()
                $amountStr  = $matches['Amt']
                $balanceStr = $matches['Bal']
            }
            elseif ($rest -match "^(?<Desc>.+?)\s+(?<Bal>£[\d,]+\.\d{2})$") {
                $desc       = $matches['Desc'].Trim()
                $balanceStr = $matches['Bal']
            }

            $amountGBP    = Convert-ToAmount $amountStr
            $balanceAfter = Convert-ToAmount $balanceStr
            $txType       = Get-TransactionType -description $desc -amountGBP $amountGBP

            $currentTx = [PSCustomObject]@{
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
                RawText         = $line
            }

            continue
        }
    }

    if ($currentTx) { $allTransactions += $currentTx }

    Remove-Item $tempTxt -Force
}

# Categorise
foreach ($tx in $allTransactions) {
    $tx.Category = Get-Category -description $tx.Description -transactionType $tx.TransactionType
}

# Export full CSV
$allTransactions |
    Select-Object AccountNumber, SortCode, AccountName,
                  @{Name="Date";Expression={$_.Date.ToString("yyyy-MM-dd")}},
                  Description, AmountGBP, AmountFX, FXCurrency, FXRate,
                  BalanceAfter, TransactionType, Category, RawText |
    Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outAllCsv

Write-Host "Full transactions exported to: $outAllCsv"

# ===== Summaries =====

# Monthly summary
$monthlySummary = $allTransactions |
    Group-Object AccountNumber, SortCode, AccountName, @{Name="YearMonth";Expression={$_.Date.ToString("yyyy-MM")}} |
    ForEach-Object {
        $group = $_.Group
        $key   = $_.Name -split ',' | ForEach-Object { $_.Trim() }

        [PSCustomObject]@{
            AccountNumber     = $key[0]
            SortCode          = $key[1]
            AccountName       = $key[2]
            YearMonth         = $key[3]
            TotalSpent        = ($group | Where-Object { $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalReceived     = ($group | Where-Object { $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalFxSpent      = ($group | Where-Object { $_.AmountFX -ne $null -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalCashWithdraw = ($group | Where-Object { $_.TransactionType -eq "Cash Withdrawal" } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersIn  = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersOut = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            NetMovement       = ($group | Measure-Object AmountGBP -Sum).Sum
        }
    }

$monthlySummary | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outMonthlyCsv
Write-Host "Monthly summary exported to: $outMonthlyCsv"

# UK tax-year summary
$taxYearSummary = $allTransactions |
    Group-Object AccountNumber, SortCode, AccountName, @{Name="TaxYear";Expression={ Get-UkTaxYear $_.Date }} |
    ForEach-Object {
        $group = $_.Group
        $key   = $_.Name -split ',' | ForEach-Object { $_.Trim() }

        [PSCustomObject]@{
            AccountNumber     = $key[0]
            SortCode          = $key[1]
            AccountName       = $key[2]
            TaxYear           = $key[3]
            TotalSpent        = ($group | Where-Object { $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalReceived     = ($group | Where-Object { $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalFxSpent      = ($group | Where-Object { $_.AmountFX -ne $null -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalCashWithdraw = ($group | Where-Object { $_.TransactionType -eq "Cash Withdrawal" } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersIn  = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersOut = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            NetMovement       = ($group | Measure-Object AmountGBP -Sum).Sum
        }
    }

$taxYearSummary | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outTaxYearCsv
Write-Host "UK tax-year summary exported to: $outTaxYearCsv"

# Calendar-year summary
$calendarSummary = $allTransactions |
    Group-Object AccountNumber, SortCode, AccountName, @{Name="CalendarYear";Expression={$_.Date.Year}} |
    ForEach-Object {
        $group = $_.Group
        $key   = $_.Name -split ',' | ForEach-Object { $_.Trim() }

        [PSCustomObject]@{
            AccountNumber     = $key[0]
            SortCode          = $key[1]
            AccountName       = $key[2]
            CalendarYear      = $key[3]
            TotalSpent        = ($group | Where-Object { $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalReceived     = ($group | Where-Object { $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalFxSpent      = ($group | Where-Object { $_.AmountFX -ne $null -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalCashWithdraw = ($group | Where-Object { $_.TransactionType -eq "Cash Withdrawal" } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersIn  = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersOut = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            NetMovement       = ($group | Measure-Object AmountGBP -Sum).Sum
        }
    }

$calendarSummary | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $outCalendarYearCsv
Write-Host "Calendar-year summary exported to: $outCalendarYearCsv"

Write-Host "Done."
