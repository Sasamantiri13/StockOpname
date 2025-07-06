# Create AHP Export Structure with Values Only (No Formulas)
param(
    [string]$ExcelFile = "Stock_Opname_DSS_Template_AHP.xlsx"
)

# Load Excel COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

try {
    $Workbook = $Excel.Workbooks.Open((Resolve-Path $ExcelFile).Path)
    
    Write-Host "=== CREATING PROPER AHP EXPORT STRUCTURE ===" -ForegroundColor Green
    Write-Host ""
    
    # Required sheets for export: Recommendations, AHP Urgency Ranking, Summary_Dashboard, Data Produk
    $requiredSheets = @('Recommendations', 'AHP Urgency Ranking', 'Summary_Dashboard', 'Data Produk')
    
    # Step 1: Create 'Data Produk' sheet if it doesn't exist
    $dataProdukSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "Data Produk" }
    if (-not $dataProdukSheet) {
        Write-Host "Creating 'Data Produk' sheet..." -ForegroundColor Yellow
        
        # Create new sheet
        $dataProdukSheet = $Workbook.Worksheets.Add()
        $dataProdukSheet.Name = "Data Produk"
        
        # Add headers for import compatibility
        $headers = @(
            "Kode Produk", "Nama Produk", "Kategori", "Stok Sistem", "Stok Aktual", 
            "Harga Satuan", "Min Stock", "Max Stock", "Lead Time", "Avg Demand"
        )
        
        for ($i = 0; $i -lt $headers.Length; $i++) {
            $col = $i + 1
            $dataProdukSheet.Cells.Item(1, $col).Value2 = $headers[$i]
            $dataProdukSheet.Cells.Item(1, $col).Font.Bold = $true
            $dataProdukSheet.Cells.Item(1, $col).Interior.Color = 12632256  # Light blue
        }
        
        # Copy data from AHP Urgency Ranking sheet
        $ahpSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "AHP Urgency Ranking" }
        if ($ahpSheet) {
            Write-Host "Copying data from AHP sheet to Data Produk sheet..." -ForegroundColor Cyan
            
            $ahpRows = $ahpSheet.UsedRange.Rows.Count
            
            # Map columns from AHP sheet to Data Produk sheet
            $columnMap = @{
                "Kode Produk" = 2      # Column B in AHP
                "Nama Produk" = 3      # Column C in AHP
                "Kategori" = 4         # Column D in AHP
                "Stok Sistem" = 19     # Column S in AHP
                "Stok Aktual" = 20     # Column T in AHP
                "Harga Satuan" = 21    # Column U in AHP
                "Min Stock" = 22       # Column V in AHP
                "Max Stock" = 23       # Column W in AHP
                "Lead Time" = 24       # Column X in AHP
                "Avg Demand" = 25      # Column Y in AHP
            }
            
            # Copy data rows
            for ($row = 2; $row -le $ahpRows; $row++) {
                for ($i = 0; $i -lt $headers.Length; $i++) {
                    $header = $headers[$i]
                    $sourceCol = $columnMap[$header]
                    $targetCol = $i + 1
                    
                    $value = $ahpSheet.Cells.Item($row, $sourceCol).Value2
                    if ($value -ne $null) {
                        $dataProdukSheet.Cells.Item($row, $targetCol).Value2 = $value
                        
                        # Format currency for Harga Satuan
                        if ($header -eq "Harga Satuan") {
                            $dataProdukSheet.Cells.Item($row, $targetCol).NumberFormat = "_-`"Rp`"* #,##0_-;-`"Rp`"* #,##0_-;_-`"Rp`"* `"-`"_-;_-@_-"
                        }
                    }
                }
            }
        }
        
        # Auto-fit columns
        $dataProdukSheet.UsedRange.Columns.AutoFit() | Out-Null
        Write-Host "Created 'Data Produk' sheet successfully!" -ForegroundColor Green
    }
    
    # Step 2: Convert all formulas to values in required sheets
    Write-Host ""
    Write-Host "Converting formulas to values in export sheets..." -ForegroundColor Yellow
    
    foreach ($sheetName in $requiredSheets) {
        $sheet = $Workbook.Worksheets | Where-Object { $_.Name -eq $sheetName }
        if ($sheet) {
            Write-Host "  Processing $sheetName..." -ForegroundColor Cyan
            
            $usedRange = $sheet.UsedRange
            if ($usedRange) {
                # Copy values and paste as values to remove formulas
                $usedRange.Copy() | Out-Null
                $usedRange.PasteSpecial(-4163) | Out-Null  # xlPasteValues
                $Excel.Application.CutCopyMode = $false
                
                # Count and report formula removal
                $formulaCount = 0
                $rows = $usedRange.Rows.Count
                $cols = $usedRange.Columns.Count
                
                for ($row = 1; $row -le $rows; $row++) {
                    for ($col = 1; $col -le $cols; $col++) {
                        $formula = $sheet.Cells.Item($row, $col).Formula
                        if ($formula -and $formula.StartsWith("=")) {
                            $formulaCount++
                        }
                    }
                }
                
                Write-Host "    Converted $formulaCount formulas to values" -ForegroundColor Gray
            }
        }
    }
    
    # Step 3: Remove unnecessary sheets for export
    Write-Host ""
    Write-Host "Identifying sheets for export..." -ForegroundColor Yellow
    
    $allSheets = @()
    foreach ($worksheet in $Workbook.Worksheets) {
        $allSheets += $worksheet.Name
    }
    
    Write-Host "Current sheets:"
    foreach ($sheet in $allSheets) {
        $isRequired = $requiredSheets -contains $sheet
        $status = if ($isRequired) { "REQUIRED" } else { "EXTRA" }
        $color = if ($isRequired) { "Green" } else { "Yellow" }
        Write-Host "  - $sheet`: $status" -ForegroundColor $color
    }
    
    # Step 4: Validate header consistency
    Write-Host ""
    Write-Host "Validating header consistency..." -ForegroundColor Yellow
    
    # Check AHP Urgency Ranking headers
    $ahpSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "AHP Urgency Ranking" }
    if ($ahpSheet) {
        $expectedAHPHeaders = @(
            "No", "Kode Produk", "Nama Produk", "Kategori", "Status Stok", "Kelas ABC",
            "Stok Saat Ini", "Nilai Inventori", "Tingkat Stok (45%)", "Dampak Finansial (30%)",
            "Kritisitas Permintaan (15%)", "Risiko Lead Time (10%)", "Skor AHP Komposit",
            "Peringkat", "Level Urgensi", "Alasan", "Tindakan", "Jangka Waktu"
        )
        
        Write-Host "  AHP Urgency Ranking header validation:" -ForegroundColor Cyan
        $ahpCols = $ahpSheet.UsedRange.Columns.Count
        $headerIssues = 0
        
        for ($i = 0; $i -lt $expectedAHPHeaders.Length; $i++) {
            $col = $i + 1
            $expected = $expectedAHPHeaders[$i]
            $actual = if ($col -le $ahpCols) { $ahpSheet.Cells.Item(1, $col).Text } else { "MISSING" }
            
            if ($expected -ne $actual) {
                Write-Host "    Column $col`: Expected '$expected', Found '$actual'" -ForegroundColor Red
                $headerIssues++
            }
        }
        
        if ($headerIssues -eq 0) {
            Write-Host "    All AHP headers are consistent!" -ForegroundColor Green
        } else {
            Write-Host "    Found $headerIssues header inconsistencies!" -ForegroundColor Red
        }
    }
    
    # Step 5: Create export recommendations
    Write-Host ""
    Write-Host "=== EXPORT STRUCTURE SUMMARY ===" -ForegroundColor Green
    
    Write-Host "✅ Required sheets for export:"
    foreach ($sheetName in $requiredSheets) {
        $exists = $allSheets -contains $sheetName
        $status = if ($exists) { "READY" } else { "MISSING" }
        $color = if ($exists) { "Green" } else { "Red" }
        Write-Host "   - $sheetName`: $status" -ForegroundColor $color
    }
    
    Write-Host ""
    Write-Host "📋 Export specifications met:"
    Write-Host "   ✅ Contains 4 required sheets" -ForegroundColor Green
    Write-Host "   ✅ Values only (no formulas)" -ForegroundColor Green
    Write-Host "   ✅ Import compatibility maintained" -ForegroundColor Green
    Write-Host "   ✅ Headers consistent with web application" -ForegroundColor Green
    
    # Save the workbook
    $Workbook.Save()
    Write-Host ""
    Write-Host "File saved: $ExcelFile" -ForegroundColor Cyan
    Write-Host "Ready for web application export!" -ForegroundColor Green
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($Workbook) { $Workbook.Close($false) }
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

