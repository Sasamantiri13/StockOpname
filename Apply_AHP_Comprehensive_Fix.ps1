# Comprehensive PowerShell Script to Fix AHP Excel and Status Stok Column

Write-Host "Starting comprehensive AHP Excel fix..." -ForegroundColor Green

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
        Write-Host "AHP sheet not found!" -ForegroundColor Red
        return
    }
    
    Write-Host "Fixing column headers and formulas..." -ForegroundColor Cyan
    
    # Fix missing column headers first
    $targetSheet.Cells.Item(1, 13) = "Skor AHP Komposit"
    $targetSheet.Cells.Item(1, 14) = "Peringkat" 
    $targetSheet.Cells.Item(1, 15) = "Level Urgensi"
    $targetSheet.Cells.Item(1, 16) = "Alasan"
    $targetSheet.Cells.Item(1, 17) = "Tindakan"
    $targetSheet.Cells.Item(1, 18) = "Jangka Waktu"
    
    # Clear existing data in formula columns
    $targetSheet.Range("E:R").ClearContents()
    
    # Re-add headers
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
    $targetSheet.Cells.Item(1, 16) = "Alasan"
    $targetSheet.Cells.Item(1, 17) = "Tindakan"
    $targetSheet.Cells.Item(1, 18) = "Jangka Waktu"
    
    Write-Host "Linking data from DSS-SPK_Analysis sheet..." -ForegroundColor Yellow
    
    # Apply formulas to rows 2-10 linking to DSS sheet data
    for ($row = 2; $row -le 10; $row++) {
        $dssRow = $row  # Assuming same row mapping
        
        # Link basic data from DSS-SPK_Analysis sheet
        $targetSheet.Cells.Item($row, 1).Formula = "=$dssRow"  # Row number
        $targetSheet.Cells.Item($row, 2).Formula = "='DSS-SPK_Analysis'!A$dssRow"  # Product_Code
        $targetSheet.Cells.Item($row, 3).Formula = "='DSS-SPK_Analysis'!B$dssRow"  # Product_Name
        $targetSheet.Cells.Item($row, 4).Formula = "='DSS-SPK_Analysis'!C$dssRow"  # Category
        
        # Status Stok - Use the same logic as DSS sheet
        $statusFormula = "=IF('DSS-SPK_Analysis'!E$dssRow<=0,""Out of Stock"",IF('DSS-SPK_Analysis'!E$dssRow<='DSS-SPK_Analysis'!L$dssRow,""Reorder"",""Normal""))"
        $targetSheet.Cells.Item($row, 5).Formula = $statusFormula
        
        $targetSheet.Cells.Item($row, 6).Formula = "='DSS-SPK_Analysis'!P$dssRow"  # ABC_Class
        $targetSheet.Cells.Item($row, 7).Formula = "='DSS-SPK_Analysis'!E$dssRow"  # Actual_Stock
        $targetSheet.Cells.Item($row, 8).Formula = "='DSS-SPK_Analysis'!I$dssRow"  # Inventory_Value
        
        # AHP Criteria Scores
        # Stock Level Score (Column I) - Based on Status Stok
        $stockFormula = "=IF(E$row=""Out of Stock"",100,IF(E$row=""Reorder"",90,IF(E$row=""Low Stock"",80,IF(E$row=""Overstock"",60,50))))"
        $targetSheet.Cells.Item($row, 9).Formula = $stockFormula
        
        # Financial Impact Score (Column J) - Normalized inventory value
        $financialFormula = "=IF(MAX(H`$2:H`$10)>0,(H$row/MAX(H`$2:H`$10))*100,0)"
        $targetSheet.Cells.Item($row, 10).Formula = $financialFormula
        
        # Demand Criticality Score (Column K) - Based on ABC Class
        $demandFormula = "=IF(F$row=""A"",90,IF(F$row=""B"",70,40))"
        $targetSheet.Cells.Item($row, 11).Formula = $demandFormula
        
        # Lead Time Risk Score (Column L) - Based on Category
        $leadTimeFormula = "=IF(D$row=""Electronics"",80,IF(D$row=""Office"",70,60))"
        $targetSheet.Cells.Item($row, 12).Formula = $leadTimeFormula
        
        # Composite AHP Score (Column M) - Weighted average
        $compositeFormula = "=(I$row*0.45)+(J$row*0.30)+(K$row*0.15)+(L$row*0.10)"
        $targetSheet.Cells.Item($row, 13).Formula = $compositeFormula
        
        # Ranking (Column N) - Proper ranking
        $rankFormula = "=RANK(M$row,M`$2:M`$10,0)"
        $targetSheet.Cells.Item($row, 14).Formula = $rankFormula
        
        # Level Urgensi (Column O) - Based on composite score
        $urgencyFormula = "=IF(M$row>=83,""KRITIS"",IF(M$row>=65,""TINGGI"",IF(M$row>=40,""SEDANG"",""RENDAH"")))"
        $targetSheet.Cells.Item($row, 15).Formula = $urgencyFormula
        
        # Alasan (Column P) - Reasoning based on status and score
        $reasonFormula = "=IF(E$row=""Out of Stock"",""Stok habis - memerlukan pemesanan darurat"",IF(E$row=""Reorder"",""Mencapai titik pemesanan kembali"",IF(M$row>=83,""Skor AHP tinggi - prioritas kritis"",IF(M$row>=65,""Skor AHP sedang-tinggi - perlu perhatian"",""Kondisi stok dalam batas normal""))))"
        $targetSheet.Cells.Item($row, 16).Formula = $reasonFormula
        
        # Tindakan (Column Q) - Actions based on urgency
        $actionFormula = "=IF(O$row=""KRITIS"",""Buat pesanan berdasarkan EOQ"",IF(O$row=""TINGGI"",""Terapkan strategi sesuai kondisi"",IF(O$row=""SEDANG"",""Pantau dan rencanakan pemesanan"",""Pantau berkala"")))"
        $targetSheet.Cells.Item($row, 17).Formula = $actionFormula
        
        # Jangka Waktu (Column R) - Timeline based on urgency
        $timelineFormula = "=IF(O$row=""KRITIS"",""Segera"",IF(O$row=""TINGGI"",""Dalam 1-10 hari"",IF(O$row=""SEDANG"",""Dalam 1-4 minggu"",""Bulanan"")))"
        $targetSheet.Cells.Item($row, 18).Formula = $timelineFormula
    }
    
    # Format headers
    $headerRange = $targetSheet.Range("A1:R1")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 12632256  # Light gray color
    
    # Auto-fit columns
    $targetSheet.Columns.AutoFit() | Out-Null
    
    # Force recalculation
    $excel.Calculate()
    
    Write-Host "AHP formulas and Status Stok fixed successfully" -ForegroundColor Green
    
    # Save and close
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "Comprehensive AHP Excel fix completed successfully!" -ForegroundColor Green
    Write-Host "File updated: $workbookPath" -ForegroundColor Cyan
    
} catch {
    Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($workbook) { $workbook.Close() }
    if ($excel) { $excel.Quit() }
}

