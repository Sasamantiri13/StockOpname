# PowerShell Script to Add Price/Harga Column to AHP Excel Template

Write-Host "Adding Price/Harga column to AHP Excel template..." -ForegroundColor Green

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
    
    Write-Host "Adding Harga/Unit Cost column to AHP sheet..." -ForegroundColor Cyan
    
    # Shift columns to make room for Harga column after Stok Saat Ini (column 7)
    # Insert new column at position 8 (between Stok Saat Ini and Nilai Inventori)
    $targetSheet.Columns("H:H").Insert()
    
    # Update headers to include Harga column
    $targetSheet.Cells.Item(1, 8) = "Harga Satuan (Unit Cost)"
    $targetSheet.Cells.Item(1, 9) = "Nilai Inventori"
    $targetSheet.Cells.Item(1, 10) = "Tingkat Stok (45%)"
    $targetSheet.Cells.Item(1, 11) = "Dampak Finansial (30%)"
    $targetSheet.Cells.Item(1, 12) = "Kritisitas Permintaan (15%)"
    $targetSheet.Cells.Item(1, 13) = "Risiko Lead Time (10%)"
    $targetSheet.Cells.Item(1, 14) = "Skor AHP Komposit"
    $targetSheet.Cells.Item(1, 15) = "Peringkat"
    $targetSheet.Cells.Item(1, 16) = "Level Urgensi"
    $targetSheet.Cells.Item(1, 17) = "Alasan"
    $targetSheet.Cells.Item(1, 18) = "Tindakan"
    $targetSheet.Cells.Item(1, 19) = "Jangka Waktu"
    
    # Apply formulas to rows 2-10 with updated column references
    for ($row = 2; $row -le 10; $row++) {
        
        # Skip if row is empty
        if ([string]::IsNullOrEmpty($targetSheet.Cells.Item($row, 2).Value)) {
            continue
        }
        
        # Link basic data from DSS-SPK_Analysis sheet
        $targetSheet.Cells.Item($row, 1).Formula = "=$row"  # Row number
        $targetSheet.Cells.Item($row, 2).Formula = "='DSS-SPK_Analysis'!A$row"  # Product_Code
        $targetSheet.Cells.Item($row, 3).Formula = "='DSS-SPK_Analysis'!B$row"  # Product_Name
        $targetSheet.Cells.Item($row, 4).Formula = "='DSS-SPK_Analysis'!C$row"  # Category
        
        # Status Stok - Use the same logic as DSS sheet
        $statusFormula = "=IF('DSS-SPK_Analysis'!E$row<=0,""Out of Stock"",IF('DSS-SPK_Analysis'!E$row<='DSS-SPK_Analysis'!L$row,""Reorder"",""Normal""))"
        $targetSheet.Cells.Item($row, 5).Formula = $statusFormula
        
        $targetSheet.Cells.Item($row, 6).Formula = "='DSS-SPK_Analysis'!P$row"  # ABC_Class
        $targetSheet.Cells.Item($row, 7).Formula = "='DSS-SPK_Analysis'!E$row"  # Actual_Stock
        
        # NEW: Harga Satuan (Unit Cost) - Link from Input_Data sheet
        $targetSheet.Cells.Item($row, 8).Formula = "='Input_Data'!F$row"  # Unit_Cost
        
        # Nilai Inventori - Now calculated using the price from column H
        $targetSheet.Cells.Item($row, 9).Formula = "=G$row*H$row"  # Actual_Stock * Unit_Cost
        
        # AHP Criteria Scores (Updated column references)
        # Stock Level Score (Column J) - Based on Status Stok
        $stockFormula = "=IF(E$row=""Out of Stock"",100,IF(E$row=""Reorder"",90,IF(E$row=""Low Stock"",80,IF(E$row=""Overstock"",60,50))))"
        $targetSheet.Cells.Item($row, 10).Formula = $stockFormula
        
        # Financial Impact Score (Column K) - Normalized inventory value
        $financialFormula = "=IF(MAX(I`$2:I`$10)>0,(I$row/MAX(I`$2:I`$10))*100,0)"
        $targetSheet.Cells.Item($row, 11).Formula = $financialFormula
        
        # Demand Criticality Score (Column L) - Based on ABC Class
        $demandFormula = "=IF(F$row=""A"",90,IF(F$row=""B"",70,40))"
        $targetSheet.Cells.Item($row, 12).Formula = $demandFormula
        
        # Lead Time Risk Score (Column M) - Based on Category
        $leadTimeFormula = "=IF(D$row=""Electronics"",80,IF(D$row=""Office"",70,60))"
        $targetSheet.Cells.Item($row, 13).Formula = $leadTimeFormula
        
        # Composite AHP Score (Column N) - Weighted average
        $compositeFormula = "=(J$row*0.45)+(K$row*0.30)+(L$row*0.15)+(M$row*0.10)"
        $targetSheet.Cells.Item($row, 14).Formula = $compositeFormula
        
        # Ranking (Column O) - Proper ranking
        $rankFormula = "=RANK(N$row,N`$2:N`$10,0)"
        $targetSheet.Cells.Item($row, 15).Formula = $rankFormula
        
        # Level Urgensi (Column P) - Based on composite score
        $urgencyFormula = "=IF(N$row>=83,""KRITIS"",IF(N$row>=65,""TINGGI"",IF(N$row>=40,""SEDANG"",""RENDAH"")))"
        $targetSheet.Cells.Item($row, 16).Formula = $urgencyFormula
        
        # Alasan (Column Q) - Reasoning based on status and score
        $reasonFormula = "=IF(E$row=""Out of Stock"",""Stok habis - memerlukan pemesanan darurat"",IF(E$row=""Reorder"",""Mencapai titik pemesanan kembali"",IF(N$row>=83,""Skor AHP tinggi - prioritas kritis"",IF(N$row>=65,""Skor AHP sedang-tinggi - perlu perhatian"",""Kondisi stok dalam batas normal""))))"
        $targetSheet.Cells.Item($row, 17).Formula = $reasonFormula
        
        # Tindakan (Column R) - Actions based on urgency
        $actionFormula = "=IF(P$row=""KRITIS"",""Buat pesanan berdasarkan EOQ"",IF(P$row=""TINGGI"",""Terapkan strategi sesuai kondisi"",IF(P$row=""SEDANG"",""Pantau dan rencanakan pemesanan"",""Pantau berkala"")))"
        $targetSheet.Cells.Item($row, 18).Formula = $actionFormula
        
        # Jangka Waktu (Column S) - Timeline based on urgency
        $timelineFormula = "=IF(P$row=""KRITIS"",""Segera"",IF(P$row=""TINGGI"",""Dalam 1-10 hari"",IF(P$row=""SEDANG"",""Dalam 1-4 minggu"",""Bulanan"")))"
        $targetSheet.Cells.Item($row, 19).Formula = $timelineFormula
    }
    
    # Format headers
    $headerRange = $targetSheet.Range("A1:S1")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 12632256  # Light gray color
    
    # Format Harga column as currency
    $priceRange = $targetSheet.Range("H:H")
    $priceRange.NumberFormat = "_-Rp* #,##0_-;-Rp* #,##0_-;_-Rp* ""-""_-;_-@_-"
    
    # Format Nilai Inventori column as currency
    $inventoryRange = $targetSheet.Range("I:I")
    $inventoryRange.NumberFormat = "_-Rp* #,##0_-;-Rp* #,##0_-;_-Rp* ""-""_-;_-@_-"
    
    # Auto-fit columns
    $targetSheet.Columns.AutoFit() | Out-Null
    
    # Force recalculation
    $excel.Calculate()
    
    Write-Host "Price/Harga column added successfully" -ForegroundColor Green
    
    # Save and close
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "Price/Harga column integration completed successfully!" -ForegroundColor Green
    Write-Host "File updated: $workbookPath" -ForegroundColor Cyan
    Write-Host "" -ForegroundColor White
    Write-Host "Changes made:" -ForegroundColor Yellow
    Write-Host "✓ Added 'Harga Satuan (Unit Cost)' column" -ForegroundColor Green
    Write-Host "✓ Updated Nilai Inventori calculation to use price" -ForegroundColor Green
    Write-Host "✓ Applied proper currency formatting" -ForegroundColor Green
    Write-Host "✓ Updated all AHP formula references" -ForegroundColor Green
    
} catch {
    Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    
    if ($workbook) { $workbook.Close() }
    if ($excel) { $excel.Quit() }
}

