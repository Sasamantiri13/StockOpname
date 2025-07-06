# Simple PowerShell Script to Add Price Column to AHP Excel

Write-Host "Adding Price column to AHP Excel..." -ForegroundColor Green

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true  # Make visible for debugging
    $excel.DisplayAlerts = $false
    
    $workbookPath = "C:\Users\SASA\P-SPK-StockOpname-TSX\Stock_Opname_DSS_Template_AHP.xlsx"
    $workbook = $excel.Workbooks.Open($workbookPath)
    
    # Find AHP sheet
    $targetSheet = $workbook.Worksheets("AHP Urgency Ranking")
    
    Write-Host "Manually adding price column..." -ForegroundColor Cyan
    
    # Add price header in column 8
    $targetSheet.Cells.Item(1, 8) = "Harga Satuan"
    
    # Add sample price data based on product types
    $targetSheet.Cells.Item(2, 8) = 8500000   # Laptop Dell
    $targetSheet.Cells.Item(3, 8) = 150000    # Mouse Wireless  
    $targetSheet.Cells.Item(4, 8) = 750000    # Keyboard Mechanical
    
    # Update Nilai Inventori calculation (column 9)
    $targetSheet.Cells.Item(1, 9) = "Nilai Inventori"
    for ($row = 2; $row -le 4; $row++) {
        $targetSheet.Cells.Item($row, 9).Formula = "=G$row*H$row"
    }
    
    # Format price column as currency
    $priceRange = $targetSheet.Range("H:H")
    $priceRange.NumberFormat = "_-Rp* #,##0_-;-Rp* #,##0_-;_-Rp* ""-""_-;_-@_-"
    
    # Format inventory value as currency  
    $inventoryRange = $targetSheet.Range("I:I")
    $inventoryRange.NumberFormat = "_-Rp* #,##0_-;-Rp* #,##0_-;_-Rp* ""-""_-;_-@_-"
    
    $targetSheet.Columns.AutoFit() | Out-Null
    
    Write-Host "Price column added successfully" -ForegroundColor Green
    
    # Save and show
    $workbook.Save()
    
    Write-Host "File saved. Please check the Excel file to verify the changes." -ForegroundColor Yellow
    Write-Host "Press Enter to close Excel..." -ForegroundColor Cyan
    Read-Host
    
    $workbook.Close()
    $excel.Quit()
    
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($workbook) { $workbook.Close() }
    if ($excel) { $excel.Quit() }
}

