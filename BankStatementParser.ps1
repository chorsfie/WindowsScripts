<#
.SYNOPSIS
    Bank Statement PDF -> CSV Extractor (Chase, Lloyds, etc.)
.DESCRIPTION
    Parses UK bank PDF statements using Poppler `pdftotext`, extracts
    transaction rows and balances, classifies transactions, and exports:
      - Full transactions CSV
      - Monthly summary CSV
      - UK tax-year summary CSV (includes TotalInterest)
      - Calendar-year summary CSV
      - Per-bank summaries (screen and Markdown)
.REQUIREMENTS
    - Poppler `pdftotext` installed and the path assigned to `$popplerPath`.
    - Run in PowerShell with appropriate ExecutionPolicy.
.AUTHOR
    Craig Horsfield
.VERSION
    2.0.0
# .CHANGELOG
#     2026-02-21 v1.0   Initial stable extractor (basic parsing + CSV exports)
#     2026-02-21 v1.1   Relaxed parsing rules; fixed date parsing issues
#     2026-02-21 v1.2   Added `TotalInterest` to tax-year summary and console synopsis
#     2026-02-21 v2.0   Generic parser, per-bank summaries, renamed script
#>

function Convert-ToAmount {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    
    $s = $s.Trim()
    $isNegative = $false
    
    if ($s.StartsWith('-')) {
        $isNegative = $true
        $s = $s.Substring(1).Trim()
    } elseif ($s.StartsWith('+')) {
        $s = $s.Substring(1).Trim()
    }
    
    while ($s.Length -gt 0 -and -not [char]::IsDigit($s[0])) {
        $s = $s.Substring(1)
    }
    
    $s = $s.Replace(' ', '').Trim()
    
    if ($s -match '^([0-9]{1,3}(?:,[0-9]{3})*|[0-9]+)\.([0-9]{2})$') {
        $num = $matches[1].Replace(',', '')
        $dec = $matches[2]
        try {
            $val = [decimal]::Parse($num + '.' + $dec)
            if ($isNegative) { $val = -$val }
            return $val
        } catch {
            return $null
        }
    }
    return $null
}

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

function Get-TransactionType {
    param([string]$description, [decimal]$amountGBP)
    $d = $description.ToLower()
    if ($d -match "(?i)interest") { return "Interest" }
    if ($d -match "cash withdrawal") { return "Cash Withdrawal" }
    if ($d -match "^from " -or $d -match "^to ") { return "Transfer" }
    if ($d -match "payment") { return "Payment" }
    if ($amountGBP -gt 0) { return "Transfer" }
    return "Purchase"
}

function Get-UkTaxYear {
    param([datetime]$date)
    $year = $date.Year
    $taxYearStart = Get-Date -Year $year -Month 4 -Day 6
    if ($date -lt $taxYearStart) { $startYear = $year - 1 } else { $startYear = $year }
    $endYearShort = ($startYear + 1).ToString().Substring(2)
    return "{0}-{1}" -f $startYear, $endYearShort
}

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

# Folder selection and PDF discovery logic

$popplerPath = "C:\Program Files\Poppler\poppler-25.12.0\Library\bin\pdftotext.exe"
if (-not (Test-Path $popplerPath)) {
    throw "pdftotext not found at '$popplerPath'. Update the path if needed."
}

# Always prompt for folder selection
$defaultBrowse = 'G:\My Drive\Statements'
if (-not (Test-Path $defaultBrowse)) { $defaultBrowse = 0 }
$folder = $null
while (-not $folder) {
    $folder = (New-Object -ComObject Shell.Application).BrowseForFolder(0, "Select folder containing bank PDF statements", 0, $defaultBrowse)
    if (-not $folder) { Write-Host "No folder selected. Please try again." -ForegroundColor Yellow }
}
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
$outSummaryMd       = Join-Path $combinedRoot "Bank_Summary_Report_$stamp.md"

# PDF discovery and diagnostic
$pdfFiles = Get-ChildItem -Path $rootPath -Filter *.pdf -Recurse | Where-Object { $_.FullName -notmatch '(Archive|FirstDirect|British Gas)' }
if ($pdfFiles.Count -eq 0) {
    Write-Host "No PDF files found in selected folder (after exclusions)." -ForegroundColor Red
}

$allTransactions = @()
$skippedPdfs = @()
$processedCount = 0
$VerboseMode = $false
if ($args -contains '--dry-run' -or $args -contains '-v') { $VerboseMode = $true }

Get-ChildItem -Path $rootPath -Filter *.pdf -Recurse |
    Where-Object { $_.FullName -notmatch '(Archive|FirstDirect|British Gas)' } |
    ForEach-Object {
        $pdf = $_.FullName
        Write-Host "Processing $pdf"
        $tempTxt = [System.IO.Path]::GetTempFileName()
        & $popplerPath -layout -enc UTF-8 $pdf $tempTxt 2>$null
        $lines = Get-Content -LiteralPath $tempTxt -Encoding UTF8

        # Diagnostic: print first 10 lines of extracted text
        Write-Host "[DIAGNOSTIC] First 10 lines of extracted text:" -ForegroundColor Cyan
        $lines | Select-Object -First 10 | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }

        $accountNumber = $null
        $sortCode      = $null
        $accountName   = $null
        foreach ($line in $lines) {
            if (-not $accountNumber -and $line -match "(?i)account\s+number[:\s]+(\d+)") { $accountNumber = $matches[1] }
            if (-not $sortCode -and $line -match "(?i)sort\s+code[:\s]+([0-9\-]+)") { $sortCode = $matches[1] }
            if (-not $accountName -and $line -match "(.+?) statement") { $accountName = $matches[1].Trim() }
        }
        $bankName = 'Unknown'
        $lloydsText = ($lines | Select-String -Pattern '(?i)lloyds')
        $lloydsFile = ($pdf -match '(?i)lloyds')
        $chaseText = ($lines | Select-String -Pattern '(?i)chase')
        $chaseFile = ($pdf -match '(?i)chase|rain.*day|holiday')
        if ($lloydsText -or $lloydsFile) {
            $bankName = 'Lloyds'
        } elseif ($chaseText -or $chaseFile) {
            $bankName = 'Chase'
        } else {
            Write-Host "[DIAGNOSTIC] Skipping file: $pdf (bank not detected)" -ForegroundColor Yellow
            continue
        }
        if (-not $accountNumber) { $accountNumber = "Unknown" }
        if (-not $sortCode) { $sortCode = "N/A" }
        if (-not $accountName) { $accountName = [System.IO.Path]::GetFileNameWithoutExtension($pdf) }

        if ($bankName -eq 'Chase') {
            # --- Chase parsing logic ---
            Write-Host "[DIAGNOSTIC] Parsing Chase file: $pdf" -ForegroundColor Cyan
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
                    if ($rest -notmatch "[^\d\-+]*[-+]?\d[\d,]*\.\d{2}") { continue }
                    # Flush previous
                    if ($currentTx) { $allTransactions += $currentTx }
                    $date = Get-ParsedDate $dateStr
                    
                    # Chase PDFs: Extract amounts by finding all currency values in the line
                    # Match amounts: +/- optional, currency symbol optional, digits with optional commas, decimal point, 2 digits
                    $amounts = [regex]::Matches($rest, '[-+]?[\d,]+\.\d{2}') | ForEach-Object { $_.Value }
                    
                    # If no matches, try with currency symbol
                    if ($amounts.Count -eq 0) {
                        $amounts = [regex]::Matches($rest, '[-+]?[\d,]+\.\d{2}') | ForEach-Object { $_.Value }
                    }
                    
                    $amountStr = $null
                    $balanceStr = $null
                    $desc = $rest
                    
                    if ($amounts.Count -ge 2) {
                        # Two or more amounts: second-to-last is transaction amount, last is balance
                        $balanceStr = $amounts[-1]
                        $amountStr  = $amounts[-2]
                        # Remove amounts from description using simple replace
                        $desc = $rest
                        foreach ($amt in $amounts) {
                            $desc = $desc -replace [regex]::Escape($amt) + '\s*'
                        }
                    } elseif ($amounts.Count -eq 1) {
                        # Single amount: could be balance only or transaction without balance
                        # Check if description suggests it's balance-only (Opening/Closing balance lines)
                        if ($rest -match "(?i)(opening|closing)\s+balance") {
                            $balanceStr = $amounts[0]
                        } else {
                            $amountStr  = $amounts[0]
                        }
                        # Remove amount from description
                        if ($amounts.Count -gt 0) {
                            $desc = $rest
                            foreach ($amt in $amounts) {
                                $desc = $desc -replace [regex]::Escape($amt) + '\s*'
                            }
                        }
                    }
                    
                    $desc = $desc.Trim()
                    $amountGBP    = Convert-ToAmount $amountStr
                    $balanceAfter = Convert-ToAmount $balanceStr
                    $txType       = Get-TransactionType -description $desc -amountGBP $amountGBP
                    # Force interest transactions to always be positive
                    if ($desc -match '(?i)interest' -and $amountGBP -lt 0) {
                        $amountGBP = -$amountGBP
                    }
                    if ($amountGBP -eq $null -and $balanceAfter -eq $null) {
                        # Skip lines where we couldn't extract any amounts
                    }
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
        } elseif ($bankName -eq 'Lloyds') {
            # --- Lloyds parsing logic for tabular format ---
            $startIdx = $null
            foreach ($idx in 0..($lines.Count - 1)) {
                if ($lines[$idx] -match '(?i)Your Transactions') {
                    $startIdx = $idx + 1
                    break
                }
            }
            if (-not $startIdx) { $startIdx = 0 }
            
            # Lloyds transactions: lines starting with date (e.g., "01 Nov 24")
            # Each transaction is multiple lines:
            # - Line 1: Date [spaces] Description [spaces] Type [spaces] Amount [spaces] Balance
            # - Line 2: "Money In (£)" or "Money Out (£)" label
            # - Line 3: "blank." for the opposite field
            $datePattern = '^\s*(\d{1,2})\s+([A-Za-z]{3})\s+(\d{2,4})'
            
            for ($i = $startIdx; $i -lt $lines.Count; $i++) {
                $line = $lines[$i]
                if ($line -match $datePattern) {
                    # This is a transaction line
                    # Extract all contiguous non-whitespace blocks
                    $fields = @($line -split '\s{2,}' | Where-Object { $_.Length -gt 0 })
                    
                    if ($fields.Count -ge 3) {
                        $dateStr = $fields[0].Trim()
                        $desc = if ($fields.Count -gt 1) { $fields[1].Trim() -replace '\s+', ' ' } else { '' }
                        
                        # Determine if this has a Type field or not
                        # If field[2] is an amount (has decimals), then no Type field (e.g., interest)
                        # If field[2] is a type code (DD, BGC, DEB, etc.), then we have standard format
                        $hasTypeField = if ($fields.Count -gt 2) {
                            $field2 = $fields[2].Trim()
                            -not ($field2 -match '^[\d,]+\.\d{2}$')  # True if NOT an amount
                        } else {
                            $false
                        }
                        
                        # Extract type and amounts based on format
                        $type = ''
                        $startAmountIdx = 3
                        if ($hasTypeField) {
                            $type = $fields[2].Trim()
                            $startAmountIdx = 3
                        } else {
                            # No explicit type field, amounts start at field[2]
                            $startAmountIdx = 2
                        }
                        
                        # Extract monetary amounts from remaining fields
                        $amounts = @()
                        for ($f = $startAmountIdx; $f -lt $fields.Count; $f++) {
                            $field = $fields[$f].Trim()
                            # Match currency amounts: 1,234.56 or 1234.56
                            if ($field -match '^[\d,]+\.\d{2}$') {
                                $amounts += $field
                            }
                        }
                        
                        $transactionAmount = $null
                        $balanceAfter = if ($amounts.Count -gt 0) { $amounts[-1] } else { $null }
                        
                        # Check the next line to determine Money In vs Money Out
                        $isMoneyOut = $false
                        if ($i + 1 -lt $lines.Count) {
                            $nextLine = $lines[$i + 1]
                            if ($nextLine -match '(?i)Money\s+Out') {
                                $isMoneyOut = $true
                            }
                        }
                        
                        # The transaction amount is the one before the balance
                        if ($amounts.Count -ge 2) {
                            $txAmountStr = $amounts[-2]
                        } elseif ($amounts.Count -eq 1) {
                            $txAmountStr = $amounts[0]
                            $balanceAfter = $null
                        } else {
                            $txAmountStr = $null
                        }
                        
                        # Parse the amount with proper sign handling
                        if ([string]::IsNullOrWhiteSpace($txAmountStr)) {
                            $transactionAmount = $null
                        } else {
                            try {
                                $cleanAmount = [decimal]::Parse(($txAmountStr -replace '[^\d\.]', ''))
                                # Interest is always positive income, regardless of Money In/Out labels
                                if ($desc -match '(?i)interest') {
                                    $transactionAmount = $cleanAmount
                                } elseif ($isMoneyOut) {
                                    $transactionAmount = -$cleanAmount
                                } else {
                                    $transactionAmount = $cleanAmount
                                }
                            } catch { }
                        }
                        
                        # Parse balance
                        $balanceValue = $null
                        if ($balanceAfter -and $balanceAfter -ne 'blank') {
                            try {
                                $balanceValue = [decimal]::Parse(($balanceAfter -replace '[^\d\.]', ''))
                            } catch { }
                        }
                        
                        # Create transaction if we have valid data
                        if ($transactionAmount -ne $null -and -not [string]::IsNullOrWhiteSpace($dateStr)) {
                            $date = Get-ParsedDate $dateStr
                            if ($desc -match '(?i)interest') {
                                $txType = 'Interest'
                            } else {
                                $txType = Get-TransactionType -description $desc -amountGBP $transactionAmount
                            }
                            
                            $allTransactions += [PSCustomObject]@{
                                BankName        = $bankName
                                AccountNumber   = $accountNumber
                                SortCode        = $sortCode
                                AccountName     = $accountName
                                Date            = $date
                                Description     = $desc
                                AmountGBP       = $transactionAmount
                                AmountFX        = $null
                                FXCurrency      = $null
                                FXRate          = $null
                                BalanceAfter    = $balanceValue
                                TransactionType = $txType
                                Category        = $null
                                YearMonth       = $date.ToString('yyyy-MM')
                                TaxYear         = Get-UkTaxYear $date
                                CalendarYear    = $date.Year
                                RawText         = $desc
                            }
                        }
                    }
                }
            }
        }
        # End Lloyds parsing block
        Remove-Item $tempTxt -Force
        $processedCount++
    }
foreach ($tx in $allTransactions) {
    $tx.Category = Get-Category -description $tx.Description -transactionType $tx.TransactionType
}
if ($allTransactions.Count -eq 0) {
    Write-Host "No transactions extracted from PDFs." -ForegroundColor Yellow
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
            TotalInterest     = ($group | Where-Object { $_.Description -match '(?i)interest' } | Measure-Object AmountGBP -Sum).Sum
            TotalFxSpent      = ($group | Where-Object { $_.AmountFX -ne $null -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalCashWithdraw = ($group | Where-Object { $_.TransactionType -eq "Cash Withdrawal" } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersIn  = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -gt 0 } | Measure-Object AmountGBP -Sum).Sum
            TotalTransfersOut = ($group | Where-Object { $_.TransactionType -eq "Transfer" -and $_.AmountGBP -lt 0 } | Measure-Object AmountGBP -Sum).Sum
            NetMovement       = ($group | Measure-Object AmountGBP -Sum).Sum
        }
    }
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
        Write-Host "Tax-year CSV written: $outTaxYearCsv"
    } else {
        'BankName,AccountNumber,SortCode,AccountName,TaxYear,TotalSpent,TotalReceived,TotalInterest,TotalFxSpent,TotalCashWithdraw,TotalTransfersIn,TotalTransfersOut,NetMovement' | Out-File -FilePath $outTaxYearCsv -Encoding UTF8
        Write-Host "Created header-only tax-year CSV: $outTaxYearCsv" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Failed to ensure tax-year CSV: $_" -ForegroundColor Red
}
$overallByTaxYear = $taxYearSummary | Group-Object TaxYear | Sort-Object Name
Write-Host ""
Write-Host "Overall UK Tax-Year Summary (all accounts):" -ForegroundColor Cyan
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
$lines += 'Generated by BankStatementParser.ps1'
$lines | Out-File -FilePath $outSummaryMd -Encoding UTF8
Write-Host "Markdown summary written to: $outSummaryMd" -ForegroundColor Green
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
