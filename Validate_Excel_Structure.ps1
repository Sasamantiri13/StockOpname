# Validate Excel Structure and Calculations
param(
    [string]$ExcelFile = "Stock_Opname_DSS_Template_AHP.xlsx"
)

# Load Excel COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

try {
    $Workbook = $Excel.Workbooks.Open((Resolve-Path $ExcelFile).Path)
    
    Write-Host "=== EXCEL FILE STRUCTURE VALIDATION ===" -ForegroundColor Green
    Write-Host "File: $ExcelFile" -ForegroundColor Yellow
    Write-Host ""
    
    # Get all worksheets
    $worksheets = @()
    foreach ($worksheet in $Workbook.Worksheets) {
        $worksheets += $worksheet.Name
    }
    
    Write-Host "Current Sheets ($($worksheets.Count)):" -ForegroundColor Cyan
    foreach ($sheet in $worksheets) {
        Write-Host "  - $sheet" -ForegroundColor Blue
    }
    
    # Required sheets for export
    $requiredSheets = @('Recommendations', 'AHP Urgency Ranking', 'Summary_Dashboard', 'Data Produk')
    $currentSheets = $worksheets
    
    Write-Host ""
    Write-Host "=== SHEET REQUIREMENTS ANALYSIS ===" -ForegroundColor Green
    Write-Host "Required for Export: Recommendations, AHP Urgency Ranking, Summary_Dashboard, Data Produk" -ForegroundColor Yellow
    
    foreach ($required in $requiredSheets) {
        $exists = $currentSheets -contains $required
        $status = if ($exists) { "FOUND" } else { "MISSING" }
        $color = if ($exists) { "Green" } else { "Red" }
        Write-Host "  $required`: $status" -ForegroundColor $color
    }
    
    # Check for extra sheets
    $extraSheets = $currentSheets | Where-Object { $requiredSheets -notcontains $_ }
    if ($extraSheets) {
        Write-Host ""
        Write-Host "Extra sheets (not needed for export):" -ForegroundColor Yellow
        foreach ($extra in $extraSheets) {
            Write-Host "  - $extra" -ForegroundColor Yellow
        }
    }
    
    Write-Host ""
    Write-Host "=== DETAILED SHEET ANALYSIS ===" -ForegroundColor Green
    
    # Analyze each sheet
    foreach ($sheetName in $currentSheets) {
        $sheet = $Workbook.Worksheets.Item($sheetName)
        Write-Host ""
        Write-Host "Sheet: $sheetName" -ForegroundColor Yellow
        Write-Host "----------------------------------------"
        
        $usedRange = $sheet.UsedRange
        if ($usedRange) {
            $rows = $usedRange.Rows.Count
            $cols = $usedRange.Columns.Count
            Write-Host "  Dimensions: $rows rows x $cols columns"
            
            # Get headers (first row)
            Write-Host "  Headers:"
            for ($col = 1; $col -le [Math]::Min($cols, 15); $col++) {
                $header = $sheet.Cells.Item(1, $col).Text
                if ($header) {
                    # Check for formulas in header row
                    $formula = $sheet.Cells.Item(1, $col).Formula
                    $hasFormula = $formula -and $formula.StartsWith("=")
                    $formulaIndicator = if ($hasFormula) { " (FORMULA)" } else { "" }
                    Write-Host "    Col $col`: $header$formulaIndicator"
                }
            }
            
            # Check for formulas in data rows (sample first few)
            if ($rows -gt 1) {
                Write-Host "  Formula Analysis (first 3 data rows):"
                for ($row = 2; $row -le [Math]::Min(4, $rows); $row++) {
                    $formulaCount = 0
                    for ($col = 1; $col -le $cols; $col++) {
                        $formula = $sheet.Cells.Item($row, $col).Formula
                        if ($formula -and $formula.StartsWith("=")) {
                            $formulaCount++
                        }
                    }
                    Write-Host "    Row $row`: $formulaCount formulas found"
                }
            }
            
        } else {
            Write-Host "  No data found"
        }
    }
    
    Write-Host ""
    Write-Host "=== WEB APPLICATION ALIGNMENT CHECK ===" -ForegroundColor Green
    
    # Check AHP Urgency Ranking sheet structure for web alignment
    $ahpSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "AHP Urgency Ranking" }
    if ($ahpSheet) {
        Write-Host ""
        Write-Host "AHP Urgency Ranking Sheet Analysis:" -ForegroundColor Cyan
        
        $expectedHeaders = @(
            "No", "Kode Produk", "Nama Produk", "Kategori", "Status Stok", "Kelas ABC",
            "Stok Saat Ini", "Nilai Inventori", "Tingkat Stok (45%)", "Dampak Finansial (30%)",
            "Kritisitas Permintaan (15%)", "Risiko Lead Time (10%)", "Skor AHP Komposit",
            "Peringkat", "Level Urgensi", "Alasan", "Tindakan", "Jangka Waktu",
            "Stok Sistem", "Stok Aktual", "Harga Satuan", "Min Stock", "Max Stock", "Lead Time", "Avg Demand"
        )
        
        $ahpUsedRange = $ahpSheet.UsedRange
        $ahpCols = $ahpUsedRange.Columns.Count
        
        Write-Host "  Expected headers vs Actual:"
        for ($i = 0; $i -lt $expectedHeaders.Length; $i++) {
            $col = $i + 1
            $expected = $expectedHeaders[$i]
            $actual = if ($col -le $ahpCols) { $ahpSheet.Cells.Item(1, $col).Text } else { "MISSING" }
            $match = $expected -eq $actual
            $status = if ($match) { "MATCH" } else { "MISMATCH" }
            $color = if ($match) { "Green" } else { "Red" }
            Write-Host "    Col $col`: Expected='$expected', Actual='$actual' - $status" -ForegroundColor $color
        }
    }
    
    # Check Summary_Dashboard structure
    $summarySheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "Summary_Dashboard" }
    if ($summarySheet) {
        Write-Host ""
        Write-Host "Summary_Dashboard Sheet Analysis:" -ForegroundColor Cyan
        
        $summaryUsedRange = $summarySheet.UsedRange
        if ($summaryUsedRange) {
            $summaryRows = $summaryUsedRange.Rows.Count
            Write-Host "  Dashboard metrics found: $summaryRows items"
            
            # Check for key metrics
            $keyMetrics = @("Total Produk", "Total Nilai Inventory", "Tingkat Akurasi", "Item Low Stock", "Item Overstock")
            Write-Host "  Key metrics check:"
            for ($row = 1; $row -le $summaryRows; $row++) {
                $metric = $summarySheet.Cells.Item($row, 1).Text
                if ($metric -and $keyMetrics -contains $metric) {
                    $value = $summarySheet.Cells.Item($row, 2).Text
                    Write-Host "    $metric`: $value" -ForegroundColor Green
                }
            }
        }
    }
    
    Write-Host ""
    Write-Host "=== RECOMMENDATIONS ===" -ForegroundColor Magenta
    
    # Missing 'Data Produk' sheet check
    $dataProdukExists = $currentSheets -contains "Data Produk"
    if (-not $dataProdukExists) {
        Write-Host "1. CREATE 'Data Produk' sheet for import compatibility" -ForegroundColor Red
        Write-Host "   - This sheet should contain base product data for reimport"
    }
    
    # Check if DSS-SPK_Analysis should be renamed or merged
    $dssAnalysisExists = $currentSheets -contains "DSS-SPK_Analysis"
    if ($dssAnalysisExists) {
        Write-Host "2. Consider renaming 'DSS-SPK_Analysis' to 'Data Produk' if it contains base data" -ForegroundColor Yellow
    }
    
    # Formula vs Values recommendation
    Write-Host "3. Ensure export contains VALUES only, not formulas" -ForegroundColor Yellow
    Write-Host "4. Verify all calculations match web application logic" -ForegroundColor Yellow
    Write-Host "5. Test import functionality with exported file" -ForegroundColor Yellow
    
    Write-Host ""
    Write-Host "=== NEXT STEPS ===" -ForegroundColor Green
    Write-Host "1. Run formula-to-values conversion if needed"
    Write-Host "2. Create/rename 'Data Produk' sheet"
    Write-Host "3. Verify web export produces exact structure"
    Write-Host "4. Test round-trip import/export"
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($Workbook) { $Workbook.Close($false) }
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

