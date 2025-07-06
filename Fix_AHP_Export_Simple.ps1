# Create proper AHP export structure - Simple version
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
    
    # Step 1: Create 'Data Produk' sheet
    $dataProdukSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "Data Produk" }
    if (-not $dataProdukSheet) {
        Write-Host "Creating 'Data Produk' sheet..." -ForegroundColor Yellow
        
        # Create new sheet
        $dataProdukSheet = $Workbook.Worksheets.Add()
        $dataProdukSheet.Name = "Data Produk"
        
        # Add headers
        $headers = @("Kode Produk", "Nama Produk", "Kategori", "Stok Sistem", "Stok Aktual", "Harga Satuan", "Min Stock", "Max Stock", "Lead Time", "Avg Demand")
        
        for ($i = 0; $i -lt $headers.Length; $i++) {
            $col = $i + 1
            $dataProdukSheet.Cells.Item(1, $col).Value2 = $headers[$i]
            $dataProdukSheet.Cells.Item(1, $col).Font.Bold = $true
            $dataProdukSheet.Cells.Item(1, $col).Interior.Color = 12632256
        }
        
        # Copy data from AHP sheet
        $ahpSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "AHP Urgency Ranking" }
        if ($ahpSheet) {
            $ahpRows = $ahpSheet.UsedRange.Rows.Count
            
            # Copy data rows (starting from row 2)
            for ($row = 2; $row -le $ahpRows; $row++) {
                $dataProdukSheet.Cells.Item($row, 1).Value2 = $ahpSheet.Cells.Item($row, 2).Value2   # Kode Produk
                $dataProdukSheet.Cells.Item($row, 2).Value2 = $ahpSheet.Cells.Item($row, 3).Value2   # Nama Produk
                $dataProdukSheet.Cells.Item($row, 3).Value2 = $ahpSheet.Cells.Item($row, 4).Value2   # Kategori
                $dataProdukSheet.Cells.Item($row, 4).Value2 = $ahpSheet.Cells.Item($row, 19).Value2  # Stok Sistem
                $dataProdukSheet.Cells.Item($row, 5).Value2 = $ahpSheet.Cells.Item($row, 20).Value2  # Stok Aktual
                $dataProdukSheet.Cells.Item($row, 6).Value2 = $ahpSheet.Cells.Item($row, 21).Value2  # Harga Satuan
                $dataProdukSheet.Cells.Item($row, 7).Value2 = $ahpSheet.Cells.Item($row, 22).Value2  # Min Stock
                $dataProdukSheet.Cells.Item($row, 8).Value2 = $ahpSheet.Cells.Item($row, 23).Value2  # Max Stock
                $dataProdukSheet.Cells.Item($row, 9).Value2 = $ahpSheet.Cells.Item($row, 24).Value2  # Lead Time
                $dataProdukSheet.Cells.Item($row, 10).Value2 = $ahpSheet.Cells.Item($row, 25).Value2 # Avg Demand
                
                # Format currency for Harga Satuan
                $dataProdukSheet.Cells.Item($row, 6).NumberFormat = "_-`"Rp`"* #,##0_-;-`"Rp`"* #,##0_-;_-`"Rp`"* `"-`"_-;_-@_-"
            }
        }
        
        $dataProdukSheet.UsedRange.Columns.AutoFit() | Out-Null
        Write-Host "Created 'Data Produk' sheet successfully!" -ForegroundColor Green
    }
    
    # Step 2: Convert formulas to values in export sheets
    Write-Host "Converting formulas to values..." -ForegroundColor Yellow
    
    $exportSheets = @('Recommendations', 'AHP Urgency Ranking', 'Summary_Dashboard', 'Data Produk')
    
    foreach ($sheetName in $exportSheets) {
        $sheet = $Workbook.Worksheets | Where-Object { $_.Name -eq $sheetName }
        if ($sheet) {
            Write-Host "  Processing $sheetName..." -ForegroundColor Cyan
            
            $usedRange = $sheet.UsedRange
            if ($usedRange) {
                # Convert formulas to values
                $usedRange.Copy() | Out-Null
                $usedRange.PasteSpecial(-4163) | Out-Null  # xlPasteValues
                $Excel.Application.CutCopyMode = $false
            }
        }
    }
    
    # Step 3: Verify structure
    Write-Host ""
    Write-Host "=== VERIFICATION ===" -ForegroundColor Green
    
    $allSheets = @()
    foreach ($worksheet in $Workbook.Worksheets) {
        $allSheets += $worksheet.Name
    }
    
    Write-Host "Export-ready sheets:"
    foreach ($sheetName in $exportSheets) {
        $exists = $allSheets -contains $sheetName
        $status = if ($exists) { "READY" } else { "MISSING" }
        $color = if ($exists) { "Green" } else { "Red" }
        Write-Host "  - $sheetName`: $status" -ForegroundColor $color
    }
    
    # Save the workbook
    $Workbook.Save()
    
    Write-Host ""
    Write-Host "SUCCESS: Excel file is now ready for AHP export!" -ForegroundColor Green
    Write-Host "Contains 4 required sheets with values only (no formulas)" -ForegroundColor Green
    Write-Host "File saved: $ExcelFile" -ForegroundColor Cyan
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($Workbook) { $Workbook.Close($false) }
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

