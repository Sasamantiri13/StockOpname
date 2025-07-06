# Create Data Produk sheet for AHP export compatibility
param(
    [string]$ExcelFile = "Stock_Opname_DSS_Template_AHP.xlsx"
)

# Load Excel COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

try {
    $Workbook = $Excel.Workbooks.Open((Resolve-Path $ExcelFile).Path)
    
    Write-Host "=== CREATING DATA PRODUK SHEET ===" -ForegroundColor Green
    
    # Check if Data Produk sheet already exists
    $dataProdukSheet = $null
    foreach ($sheet in $Workbook.Worksheets) {
        if ($sheet.Name -eq "Data Produk") {
            $dataProdukSheet = $sheet
            break
        }
    }
    
    if ($dataProdukSheet) {
        Write-Host "Data Produk sheet already exists. Updating..." -ForegroundColor Yellow
        # Clear existing content
        $dataProdukSheet.UsedRange.Clear() | Out-Null
    } else {
        Write-Host "Creating new Data Produk sheet..." -ForegroundColor Yellow
        $dataProdukSheet = $Workbook.Worksheets.Add()
        $dataProdukSheet.Name = "Data Produk"
    }
    
    # Add headers
    $headers = @("Kode Produk", "Nama Produk", "Kategori", "Stok Sistem", "Stok Aktual", "Harga Satuan", "Min Stock", "Max Stock", "Lead Time", "Avg Demand")
    
    for ($i = 0; $i -lt $headers.Length; $i++) {
        $col = $i + 1
        $dataProdukSheet.Cells.Item(1, $col).Value2 = $headers[$i]
        $dataProdukSheet.Cells.Item(1, $col).Font.Bold = $true
        $dataProdukSheet.Cells.Item(1, $col).Interior.Color = 12632256
    }
    
    # Add sample data for testing
    Write-Host "Adding sample data..." -ForegroundColor Cyan
    
    $sampleData = @(
        @("PRD001", "Laptop Dell Inspiron", "Electronics", 50, 48, 8500000, 10, 100, 7, 5),
        @("PRD002", "Mouse Wireless", "Electronics", 150, 145, 250000, 20, 200, 3, 10),
        @("PRD003", "Keyboard Mechanical", "Electronics", 75, 78, 750000, 15, 150, 5, 8)
    )
    
    for ($row = 0; $row -lt $sampleData.Length; $row++) {
        $dataRow = $sampleData[$row]
        for ($col = 0; $col -lt $dataRow.Length; $col++) {
            $cellRow = $row + 2  # Start from row 2 (after headers)
            $cellCol = $col + 1  # Start from column 1
            
            $dataProdukSheet.Cells.Item($cellRow, $cellCol).Value2 = $dataRow[$col]
            
            # Format currency for Harga Satuan (column 6)
            if ($cellCol -eq 6) {
                $dataProdukSheet.Cells.Item($cellRow, $cellCol).NumberFormat = "_-`"Rp`"* #,##0_-;-`"Rp`"* #,##0_-;_-`"Rp`"* `"-`"_-;_-@_-"
            }
        }
    }
    
    # Auto-fit columns
    $dataProdukSheet.UsedRange.Columns.AutoFit() | Out-Null
    
    Write-Host "Data Produk sheet created successfully!" -ForegroundColor Green
    
    # Now convert all export sheets to values only
    Write-Host ""
    Write-Host "Converting formulas to values in export sheets..." -ForegroundColor Yellow
    
    $exportSheets = @('Recommendations', 'AHP Urgency Ranking', 'Summary_Dashboard', 'Data Produk')
    
    foreach ($sheetName in $exportSheets) {
        $sheet = $null
        foreach ($ws in $Workbook.Worksheets) {
            if ($ws.Name -eq $sheetName) {
                $sheet = $ws
                break
            }
        }
        
        if ($sheet) {
            Write-Host "  Converting $sheetName..." -ForegroundColor Cyan
            
            $usedRange = $sheet.UsedRange
            if ($usedRange) {
                try {
                    # Create a copy of the values
                    $values = $usedRange.Value2
                    # Paste back as values
                    $usedRange.Value2 = $values
                } catch {
                    Write-Host "    Note: Some cells in $sheetName may still contain formulas" -ForegroundColor Yellow
                }
            }
        }
    }
    
    # Save the workbook
    $Workbook.Save()
    
    Write-Host ""
    Write-Host "=== SUMMARY ===" -ForegroundColor Green
    Write-Host "✅ Created 'Data Produk' sheet with import-compatible structure" -ForegroundColor Green
    Write-Host "✅ Converted export sheets to values only" -ForegroundColor Green
    Write-Host "✅ Ready for AHP export functionality" -ForegroundColor Green
    Write-Host ""
    Write-Host "Required export sheets:"
    
    foreach ($sheetName in $exportSheets) {
        $exists = $false
        foreach ($ws in $Workbook.Worksheets) {
            if ($ws.Name -eq $sheetName) {
                $exists = $true
                break
            }
        }
        $status = if ($exists) { "READY" } else { "MISSING" }
        $color = if ($exists) { "Green" } else { "Red" }
        Write-Host "  - $sheetName`: $status" -ForegroundColor $color
    }
    
    Write-Host ""
    Write-Host "File saved: $ExcelFile" -ForegroundColor Cyan
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($Workbook) { $Workbook.Close($false) }
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

