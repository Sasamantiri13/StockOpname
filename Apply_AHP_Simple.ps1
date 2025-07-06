# Simple PowerShell Script to Apply AHP Excel Formulas

Write-Host "Starting AHP Excel formula application..." -ForegroundColor Green

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $workbookPath = "C:\Users\SASA\P-SPK-StockOpname-TSX\Stock_Opname_DSS_Template_AHP.xlsx"
    $workbook = $excel.Workbooks.Open($workbookPath)
    
    Write-Host "Excel file opened successfully" -ForegroundColor Cyan
    
    # Find AHP sheet
    $targetSheet = $null
    foreach ($sheet in $workbook.Worksheets) {
        if ($sheet.Name -eq "AHP Urgency Ranking") {
            $targetSheet = $sheet
            break
        }
    }
    
    if ($targetSheet -eq $null) {
        Write-Host "Creating AHP sheet..." -ForegroundColor Yellow
        $targetSheet = $workbook.Worksheets.Add()
        $targetSheet.Name = "AHP Urgency Ranking"
        
        # Add headers
        $targetSheet.Cells.Item(1, 1) = "No"
        $targetSheet.Cells.Item(1, 2) = "Kode Produk"
        $targetSheet.Cells.Item(1, 3) = "Nama Produk"
        $targetSheet.Cells.Item(1, 4) = "Kategori"
        $targetSheet.Cells.Item(1, 5) = "Status Stok"
        $targetSheet.Cells.Item(1, 6) = "Kelas ABC"
        $targetSheet.Cells.Item(1, 7) = "Stok Saat Ini"
        $targetSheet.Cells.Item(1, 8) = "Nilai Inventori"
        $targetSheet.Cells.Item(1, 9) = "Tingkat Stok (45%)"
        $targetSheet.Cells.Item(1, 10) = "Dampak Finansial (30%)"
        $targetSheet.Cells.Item(1, 11) = "Kritisitas Permintaan (15%)"
        $targetSheet.Cells.Item(1, 12) = "Risiko Lead Time (10%)"
        $targetSheet.Cells.Item(1, 13) = "Skor AHP Komposit"
        $targetSheet.Cells.Item(1, 14) = "Peringkat"
        $targetSheet.Cells.Item(1, 15) = "Level Urgensi"
        
        # Format headers
        $headerRange = $targetSheet.Range("A1:O1")
        $headerRange.Font.Bold = $true
        
        # Add sample data
        $targetSheet.Cells.Item(2, 1) = 1
        $targetSheet.Cells.Item(2, 2) = "ELC001"
        $targetSheet.Cells.Item(2, 3) = "MacBook Pro 16 M3"
        $targetSheet.Cells.Item(2, 4) = "Electronics"
        $targetSheet.Cells.Item(2, 5) = "Reorder"
        $targetSheet.Cells.Item(2, 6) = "A"
        $targetSheet.Cells.Item(2, 7) = 23
        $targetSheet.Cells.Item(2, 8) = 805000000
        
        $targetSheet.Cells.Item(3, 1) = 2
        $targetSheet.Cells.Item(3, 2) = "ELC002"
        $targetSheet.Cells.Item(3, 3) = "iPhone 15 Pro Max"
        $targetSheet.Cells.Item(3, 4) = "Electronics"
        $targetSheet.Cells.Item(3, 5) = "Normal"
        $targetSheet.Cells.Item(3, 6) = "A"
        $targetSheet.Cells.Item(3, 7) = 42
        $targetSheet.Cells.Item(3, 8) = 840000000
        
        $targetSheet.Cells.Item(4, 1) = 3
        $targetSheet.Cells.Item(4, 2) = "OFF002"
        $targetSheet.Cells.Item(4, 3) = "Document Camera"
        $targetSheet.Cells.Item(4, 4) = "Office"
        $targetSheet.Cells.Item(4, 5) = "Out of Stock"
        $targetSheet.Cells.Item(4, 6) = "C"
        $targetSheet.Cells.Item(4, 7) = 0
        $targetSheet.Cells.Item(4, 8) = 0
    }
    
    Write-Host "Applying AHP formulas..." -ForegroundColor Cyan
    
    # Apply formulas to rows 2-10
    for ($row = 2; $row -le 10; $row++) {
        
        # Stock Level Score (Column I)
        $stockFormula = "=IF(E$row=""Out of Stock"",100,IF(E$row=""Reorder"",90,IF(E$row=""Low Stock"",80,IF(E$row=""Overstock"",60,50))))"
        $targetSheet.Cells.Item($row, 9).Formula = $stockFormula
        
        # Financial Impact Score (Column J)
        $financialFormula = "=IF(MAX(H:H)>0,(H$row/MAX(H:H))*100,0)"
        $targetSheet.Cells.Item($row, 10).Formula = $financialFormula
        
        # Demand Criticality Score (Column K)
        $demandFormula = "=IF(F$row=""A"",90,IF(F$row=""B"",70,40))"
        $targetSheet.Cells.Item($row, 11).Formula = $demandFormula
        
        # Lead Time Risk Score (Column L)
        $leadTimeFormula = "=RANDBETWEEN(20,80)"
        $targetSheet.Cells.Item($row, 12).Formula = $leadTimeFormula
        
        # Composite AHP Score (Column M)
        $compositeFormula = "=(I$row*0.45)+(J$row*0.30)+(K$row*0.15)+(L$row*0.10)"
        $targetSheet.Cells.Item($row, 13).Formula = $compositeFormula
        
        # Ranking (Column N)
        $rankFormula = "=RANK(M$row,M$2:M$10,0)"
        $targetSheet.Cells.Item($row, 14).Formula = $rankFormula
        
        # Urgency Level (Column O)
        $urgencyFormula = "=IF(M$row>=83,""KRITIS"",IF(M$row>=58,""TINGGI"",IF(M$row>=33,""SEDANG"",""RENDAH"")))"
        $targetSheet.Cells.Item($row, 15).Formula = $urgencyFormula
    }
    
    # Auto-fit columns
    $targetSheet.Columns.AutoFit() | Out-Null
    
    Write-Host "AHP formulas applied successfully" -ForegroundColor Green
    
    # Save and close
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "AHP Excel formulas applied successfully!" -ForegroundColor Green
    Write-Host "File updated: $workbookPath" -ForegroundColor Cyan
    
} catch {
    Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($workbook) { $workbook.Close() }
    if ($excel) { $excel.Quit() }
}

