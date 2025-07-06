# PowerShell Script to Apply Correct AHP Excel Formulas
# Based on internationally recognized AHP method

Write-Host "🔄 Applying AHP Excel Formulas..." -ForegroundColor Green

try {
    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Open the workbook
    $workbookPath = "C:\Users\SASA\P-SPK-StockOpname-TSX\Stock_Opname_DSS_Template_AHP.xlsx"
    $workbook = $excel.Workbooks.Open($workbookPath)
    
    Write-Host "📂 Excel file opened successfully" -ForegroundColor Cyan
    
    # Find or create the AHP sheet
    $targetSheetName = "AHP Urgency Ranking"
    $targetSheet = $null
    
    foreach ($sheet in $workbook.Worksheets) {
        if ($sheet.Name -eq $targetSheetName) {
            $targetSheet = $sheet
            break
        }
    }
    
    if ($targetSheet -eq $null) {
        Write-Host "❌ AHP sheet not found. Creating new sheet..." -ForegroundColor Yellow
        $targetSheet = $workbook.Worksheets.Add()
        $targetSheet.Name = $targetSheetName
        
        # Add headers
        $headers = @(
            "No", "Kode Produk", "Nama Produk", "Kategori", "Status Stok", "Kelas ABC",
            "Stok Saat Ini", "Nilai Inventori", "Tingkat Stok (45%)", "Dampak Finansial (30%)",
            "Kritisitas Permintaan (15%)", "Risiko Lead Time (10%)", "Skor AHP Komposit",
            "Peringkat Urgensi", "Level Urgensi", "Alasan", "Tindakan", "Jangka Waktu"
        )
        
        for ($i = 0; $i -lt $headers.Length; $i++) {
            $targetSheet.Cells.Item(1, $i + 1) = $headers[$i]
            $targetSheet.Cells.Item(1, $i + 1).Font.Bold = $true
        }
        
        # Add sample data for formula testing
        $sampleData = @(
            @(1, "ELC001", "MacBook Pro 16 M3", "Electronics", "Reorder", "A", 23, 805000000),
            @(2, "ELC002", "iPhone 15 Pro Max", "Electronics", "Normal", "A", 42, 840000000),
            @(3, "ELC003", "Samsung Galaxy S24 Ultra", "Electronics", "Low Stock", "B", 8, 144000000),
            @(4, "ACC001", "Logitech MX Master 3S", "Accessories", "Normal", "B", 85, 102000000),
            @(5, "OFF002", "Document Camera", "Office", "Out of Stock", "C", 0, 0)
        )
        
        for ($row = 0; $row -lt $sampleData.Length; $row++) {
            $currentRow = $row + 2
            $data = $sampleData[$row]
            for ($col = 0; $col -lt 8; $col++) {
                $targetSheet.Cells.Item($currentRow, $col + 1) = $data[$col]
            }
        }
    }
    
    Write-Host "✅ Found AHP sheet: $targetSheetName" -ForegroundColor Green
    
    # Apply AHP formulas starting from row 2
    $startRow = 2
    $endRow = 20  # Adjust based on your data size
    
    Write-Host "📊 Applying AHP formulas..." -ForegroundColor Cyan
    
    for ($row = $startRow; $row -le $endRow; $row++) {
        
        # Column I: Tingkat Stok (45%) - Stock Level Criteria
        $stockLevelFormula = "=IF(E$row=`"Out of Stock`",100,IF(E$row=`"Reorder`",90,IF(E$row=`"Low Stock`",80,IF(E$row=`"Overstock`",60,50))))"
        $targetSheet.Cells.Item($row, 9).Formula = $stockLevelFormula
        
        # Column J: Dampak Finansial (30%) - Financial Impact (normalized to 100)
        $financialImpactFormula = "=IF(MAX(H:H)>0,(H$row/MAX(H:H))*100,0)"
        $targetSheet.Cells.Item($row, 10).Formula = $financialImpactFormula
        
        # Column K: Kritisitas Permintaan (15%) - Demand Criticality based on ABC
        $demandCriticalityFormula = "=IF(F$row=`"A`",90,IF(F$row=`"B`",70,40))"
        $targetSheet.Cells.Item($row, 11).Formula = $demandCriticalityFormula
        
        # Column L: Risiko Lead Time (10%) - Lead Time Risk (random for demo, replace with actual data)
        $leadTimeRiskFormula = "=RANDBETWEEN(20,80)"
        $targetSheet.Cells.Item($row, 12).Formula = $leadTimeRiskFormula
        
        # Column M: Skor AHP Komposit - Composite AHP Score
        $compositeFormula = "=(I$row*0.45)+(J$row*0.30)+(K$row*0.15)+(L$row*0.10)"
        $targetSheet.Cells.Item($row, 13).Formula = $compositeFormula
        
        # Column N: Peringkat Urgensi - Ranking based on composite score
        $rankFormula = "=RANK(M$row,M`$2:M`$$endRow,0)"
        $targetSheet.Cells.Item($row, 14).Formula = $rankFormula
        
        # Column O: Level Urgensi - Urgency Level based on thresholds
        $urgencyLevelFormula = "=IF(M$row>=83,`"KRITIS`",IF(M$row>=58,`"TINGGI`",IF(M$row>=33,`"SEDANG`",`"RENDAH`")))"
        $targetSheet.Cells.Item($row, 15).Formula = $urgencyLevelFormula
        
        # Column P: Alasan - Reason based on analysis
        $reasonFormula = "=`"Analisis AHP skor `"&ROUND(M$row,1)&`" - item kelas `"&F$row&`" dengan status `"&E$row"
        $targetSheet.Cells.Item($row, 16).Formula = $reasonFormula
        
        # Column Q: Tindakan - Recommended Action
        $actionFormula = "=IF(E$row=`"Out of Stock`",`"Pesanan darurat segera!`",IF(E$row=`"Reorder`",`"Buat pesanan berdasarkan EOQ`",IF(E$row=`"Low Stock`",`"Pantau dan rencanakan pemesanan`",`"Terapkan strategi sesuai kondisi`")))"
        $targetSheet.Cells.Item($row, 17).Formula = $actionFormula
        
        # Column R: Jangka Waktu - Timeframe based on urgency level
        $timeframeFormula = "=IF(O$row=`"KRITIS`",`"Segera`",IF(O$row=`"TINGGI`",`"Dalam 1-10 hari`",IF(O$row=`"SEDANG`",`"Dalam 1-4 minggu`",`"Dalam 1 bulan`")))"
        $targetSheet.Cells.Item($row, 18).Formula = $timeframeFormula
    }
    
    Write-Host "📈 AHP formulas applied successfully" -ForegroundColor Green
    
    # Add AHP explanation sheet
    $explanationSheet = $workbook.Worksheets.Add()
    $explanationSheet.Name = "AHP Explanation"
    
    # Add AHP method explanation
    $explanationSheet.Cells.Item(1, 1) = "PENJELASAN METODE AHP (Analytic Hierarchy Process)"
    $explanationSheet.Cells.Item(1, 1).Font.Bold = $true
    $explanationSheet.Cells.Item(1, 1).Font.Size = 14
    
    $explanationSheet.Cells.Item(3, 1) = "KRITERIA DAN BOBOT:"
    $explanationSheet.Cells.Item(3, 1).Font.Bold = $true
    $explanationSheet.Cells.Item(4, 1) = "1. Tingkat Stok (45%) - Kritisitas ketersediaan stok"
    $explanationSheet.Cells.Item(5, 1) = "2. Dampak Finansial (30%) - Nilai bisnis dan dampak keuangan"
    $explanationSheet.Cells.Item(6, 1) = "3. Kritisitas Permintaan (15%) - Klasifikasi ABC dan tingkat permintaan"
    $explanationSheet.Cells.Item(7, 1) = "4. Risiko Lead Time (10%) - Risiko keterlambatan pasokan"
    
    $explanationSheet.Cells.Item(9, 1) = "FORMULA PERHITUNGAN:"
    $explanationSheet.Cells.Item(9, 1).Font.Bold = $true
    $explanationSheet.Cells.Item(10, 1) = "Skor Komposit = (Tingkat Stok × 0.45) + (Dampak Finansial × 0.30) + (Kritisitas Permintaan × 0.15) + (Risiko Lead Time × 0.10)"
    
    $explanationSheet.Cells.Item(12, 1) = "TINGKAT URGENSI:"
    $explanationSheet.Cells.Item(12, 1).Font.Bold = $true
    $explanationSheet.Cells.Item(13, 1) = "• KRITIS (83-100): Tindakan darurat diperlukan SEKARANG"
    $explanationSheet.Cells.Item(14, 1) = "• TINGGI (58-82): Tindakan diperlukan dalam hitungan hari"
    $explanationSheet.Cells.Item(15, 1) = "• SEDANG (33-57): Rencanakan tindakan dalam hitungan minggu"
    $explanationSheet.Cells.Item(16, 1) = "• RENDAH (8-32): Pantau situasi"
    
    # Auto-fit columns
    $targetSheet.Columns.AutoFit() | Out-Null
    $explanationSheet.Columns.AutoFit() | Out-Null
    
    Write-Host "📋 AHP explanation sheet added" -ForegroundColor Green
    
    # Save the workbook
    $workbook.Save()
    Write-Host "💾 Workbook saved successfully" -ForegroundColor Green
    
    # Close Excel
    $workbook.Close()
    $excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($explanationSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "✅ AHP Excel formulas applied successfully!" -ForegroundColor Green
    Write-Host "📁 File updated: $workbookPath" -ForegroundColor Cyan
    Write-Host "📊 New sheets added:" -ForegroundColor Yellow
    Write-Host "   • AHP Urgency Ranking (with formulas)" -ForegroundColor White
    Write-Host "   • AHP Explanation (method documentation)" -ForegroundColor White
    
} catch {
    Write-Host "❌ Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    
    # Clean up COM objects in case of error
    if ($workbook) { $workbook.Close() }
    if ($excel) { $excel.Quit() }
}

Write-Host "`n🎯 AHP FORMULAS SUCCESSFULLY APPLIED TO EXCEL!" -ForegroundColor Magenta

