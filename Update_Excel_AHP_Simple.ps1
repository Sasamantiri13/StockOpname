# Simplified PowerShell Script to Update Excel with AHP Calculations

Write-Host "Starting Excel AHP Update Process..." -ForegroundColor Green

try {
    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Open the workbook
    $workbookPath = "C:\Users\SASA\P-SPK-StockOpname-TSX\Stock_Opname_DSS_Template_AHP.xlsx"
    $workbook = $excel.Workbooks.Open($workbookPath)
    
    Write-Host "Excel file opened successfully" -ForegroundColor Cyan
    
    # Create new sheet for AHP
    $targetSheetName = "AHP Urgency Ranking"
    $targetSheet = $workbook.Worksheets.Add()
    $targetSheet.Name = $targetSheetName
    
    Write-Host "New sheet created: $targetSheetName" -ForegroundColor Yellow
    
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
    $targetSheet.Cells.Item(1, 14) = "Level Urgensi"
    $targetSheet.Cells.Item(1, 15) = "Alasan"
    $targetSheet.Cells.Item(1, 16) = "Tindakan"
    $targetSheet.Cells.Item(1, 17) = "Jangka Waktu"
    
    # Format headers
    $headerRange = $targetSheet.Range("A1:Q1")
    $headerRange.Font.Bold = $true
    
    # Sample data
    $sampleData = @(
        @("ELC001", "MacBook Pro 16 M3", "Electronics", "Reorder", "A", 23, 805000000),
        @("ELC002", "iPhone 15 Pro Max", "Electronics", "Normal", "A", 42, 840000000),
        @("ELC003", "Samsung Galaxy S24 Ultra", "Electronics", "Low Stock", "B", 8, 144000000),
        @("ACC001", "Logitech MX Master 3S", "Accessories", "Normal", "B", 85, 102000000),
        @("OFF002", "Document Camera", "Office", "Out of Stock", "C", 0, 0)
    )
    
    for ($row = 0; $row -lt $sampleData.Length; $row++) {
        $currentRow = $row + 2
        $data = $sampleData[$row]
        
        # Basic data
        $targetSheet.Cells.Item($currentRow, 1) = $row + 1
        $targetSheet.Cells.Item($currentRow, 2) = $data[0]
        $targetSheet.Cells.Item($currentRow, 3) = $data[1]
        $targetSheet.Cells.Item($currentRow, 4) = $data[2]
        $targetSheet.Cells.Item($currentRow, 5) = $data[3]
        $targetSheet.Cells.Item($currentRow, 6) = $data[4]
        $targetSheet.Cells.Item($currentRow, 7) = $data[5]
        $targetSheet.Cells.Item($currentRow, 8) = $data[6]
        
        # AHP Calculations
        # Tingkat Stok Score
        $stockScore = 50
        switch ($data[3]) {
            "Out of Stock" { $stockScore = 100 }
            "Low Stock" { $stockScore = 80 }
            "Reorder" { $stockScore = 90 }
            "Overstock" { $stockScore = 60 }
            default { $stockScore = 50 }
        }
        $targetSheet.Cells.Item($currentRow, 9) = $stockScore
        
        # Dampak Finansial Score
        $financialScore = if ($data[6] -gt 0) { ($data[6] / 840000000) * 100 } else { 0 }
        $targetSheet.Cells.Item($currentRow, 10) = [math]::Round($financialScore, 2)
        
        # Kritisitas Permintaan Score
        $demandScore = 40
        switch ($data[4]) {
            "A" { $demandScore = 90 }
            "B" { $demandScore = 70 }
            default { $demandScore = 40 }
        }
        $targetSheet.Cells.Item($currentRow, 11) = $demandScore
        
        # Risiko Lead Time Score
        $leadTimeScore = Get-Random -Minimum 20 -Maximum 80
        $targetSheet.Cells.Item($currentRow, 12) = $leadTimeScore
        
        # Composite AHP Score
        $compositeScore = ($stockScore * 0.45) + ($financialScore * 0.30) + ($demandScore * 0.15) + ($leadTimeScore * 0.10)
        $targetSheet.Cells.Item($currentRow, 13) = [math]::Round($compositeScore, 2)
        
        # Urgency Level
        $urgencyLevel = "RENDAH"
        if ($compositeScore -ge 83) { $urgencyLevel = "KRITIS" }
        elseif ($compositeScore -ge 58) { $urgencyLevel = "TINGGI" }
        elseif ($compositeScore -ge 33) { $urgencyLevel = "SEDANG" }
        $targetSheet.Cells.Item($currentRow, 14) = $urgencyLevel
        
        # Reason
        $reason = "Analisis AHP skor $([math]::Round($compositeScore,1)) - kelas $($data[4]) status $($data[3])"
        $targetSheet.Cells.Item($currentRow, 15) = $reason
        
        # Action
        $action = "Terapkan strategi sesuai kondisi"
        switch ($data[3]) {
            "Out of Stock" { $action = "Pesanan darurat segera!" }
            "Reorder" { $action = "Buat pesanan berdasarkan EOQ" }
            "Low Stock" { $action = "Pantau dan rencanakan pemesanan" }
        }
        $targetSheet.Cells.Item($currentRow, 16) = $action
        
        # Timeframe
        $timeframe = "Dalam 1 bulan"
        switch ($urgencyLevel) {
            "KRITIS" { $timeframe = "Segera" }
            "TINGGI" { $timeframe = "Dalam 1-10 hari" }
            "SEDANG" { $timeframe = "Dalam 1-4 minggu" }
        }
        $targetSheet.Cells.Item($currentRow, 17) = $timeframe
    }
    
    # Auto-fit columns
    $targetSheet.Columns.AutoFit() | Out-Null
    
    Write-Host "AHP calculations added successfully" -ForegroundColor Green
    
    # Save and close
    $workbook.Save()
    $workbook.Close()
    $excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "AHP Excel update completed successfully!" -ForegroundColor Green
    Write-Host "File updated: $workbookPath" -ForegroundColor Cyan
    
} catch {
    Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    
    # Clean up
    if ($workbook) { $workbook.Close() }
    if ($excel) { $excel.Quit() }
}

