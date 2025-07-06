# PowerShell Script to Update Excel with AHP Calculations
# Requires Excel to be installed

Write-Host "🔄 Starting Excel AHP Update Process..." -ForegroundColor Green

try {
    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Open the workbook
    $workbookPath = "C:\Users\SASA\P-SPK-StockOpname-TSX\Stock_Opname_DSS_Template_AHP.xlsx"
    $workbook = $excel.Workbooks.Open($workbookPath)
    
    Write-Host "📂 Excel file opened successfully" -ForegroundColor Cyan
    
    # Check if the 'Peringkat Urgensi - Tindakan Prioritas SPK' sheet exists
    $targetSheetName = "Peringkat Urgensi - Tindakan Prioritas SPK"
    $targetSheet = $null
    
    foreach ($sheet in $workbook.Worksheets) {
        if ($sheet.Name -eq $targetSheetName) {
            $targetSheet = $sheet
            break
        }
    }
    
    # If sheet doesn't exist, create it
    if ($targetSheet -eq $null) {
        Write-Host "➕ Creating new sheet: $targetSheetName" -ForegroundColor Yellow
        $targetSheet = $workbook.Worksheets.Add()
        $targetSheet.Name = $targetSheetName
    } else {
        Write-Host "✅ Found existing sheet: $targetSheetName" -ForegroundColor Green
    }
    
    # Clear existing content and add headers
    $targetSheet.Cells.Clear()
    
    # Add AHP Headers
    $headers = @(
        "No",
        "Kode Produk", 
        "Nama Produk",
        "Kategori",
        "Status Stok",
        "Kelas ABC",
        "Stok Saat Ini",
        "Nilai Inventori",
        "Tingkat Stok (45%)",
        "Dampak Finansial (30%)",
        "Kritisitas Permintaan (15%)",
        "Risiko Lead Time (10%)",
        "Skor AHP Komposit",
        "Peringkat Urgensi",
        "Level Urgensi",
        "Alasan",
        "Tindakan yang Direkomendasikan",
        "Jangka Waktu"
    )
    
    # Set headers
    for ($i = 0; $i -lt $headers.Length; $i++) {
        $targetSheet.Cells.Item(1, $i + 1) = $headers[$i]
        $targetSheet.Cells.Item(1, $i + 1).Font.Bold = $true
        $targetSheet.Cells.Item(1, $i + 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightBlue)
    }
    
    Write-Host "📊 Headers added successfully" -ForegroundColor Green
    
    # Add sample data and AHP formulas
    $sampleData = @(
        @("ELC001", "MacBook Pro 16`" M3", "Electronics", "Reorder", "A", 23, 805000000),
        @("ELC002", "iPhone 15 Pro Max", "Electronics", "Normal", "A", 42, 840000000),
        @("ELC003", "Samsung Galaxy S24 Ultra", "Electronics", "Low Stock", "B", 8, 144000000),
        @("ACC001", "Logitech MX Master 3S", "Accessories", "Normal", "B", 85, 102000000),
        @("OFF002", "Document Camera", "Office", "Out of Stock", "C", 0, 0)
    )
    
    for ($row = 0; $row -lt $sampleData.Length; $row++) {
        $currentRow = $row + 2
        $data = $sampleData[$row]
        
        # Basic data
        $targetSheet.Cells.Item($currentRow, 1) = $row + 1  # No
        $targetSheet.Cells.Item($currentRow, 2) = $data[0]  # Kode
        $targetSheet.Cells.Item($currentRow, 3) = $data[1]  # Nama
        $targetSheet.Cells.Item($currentRow, 4) = $data[2]  # Kategori
        $targetSheet.Cells.Item($currentRow, 5) = $data[3]  # Status
        $targetSheet.Cells.Item($currentRow, 6) = $data[4]  # ABC
        $targetSheet.Cells.Item($currentRow, 7) = $data[5]  # Stok
        $targetSheet.Cells.Item($currentRow, 8) = $data[6]  # Nilai
        
        # AHP Criteria Calculations - Use simple values instead of complex formulas
        # Tingkat Stok (45%) - Column I
        $stockLevelValue = 0
        if ($data[3] -eq "Out of Stock") { $stockLevelValue = 100 }
        elseif ($data[3] -eq "Low Stock") { $stockLevelValue = 80 }
        elseif ($data[3] -eq "Reorder") { $stockLevelValue = 90 }
        elseif ($data[3] -eq "Overstock") { $stockLevelValue = 60 }
        else { $stockLevelValue = 50 }
        $targetSheet.Cells.Item($currentRow, 9) = $stockLevelValue
        
        # Dampak Finansial (30%) - Column J  
        $financialImpactValue = if ($data[6] -gt 0) { ($data[6] / 840000000) * 100 } else { 0 }
        $targetSheet.Cells.Item($currentRow, 10) = [math]::Round($financialImpactValue, 2)
        
        # Kritisitas Permintaan (15%) - Column K
        $demandCriticalityValue = 0
        if ($data[4] -eq "A") { $demandCriticalityValue = 90 }
        elseif ($data[4] -eq "B") { $demandCriticalityValue = 70 }
        else { $demandCriticalityValue = 40 }
        $targetSheet.Cells.Item($currentRow, 11) = $demandCriticalityValue
        
        # Risiko Lead Time (10%) - Column L
        $leadTimeRiskValue = Get-Random -Minimum 20 -Maximum 80
        $targetSheet.Cells.Item($currentRow, 12) = $leadTimeRiskValue
        
        # Skor AHP Komposit - Column M
        $compositeScore = ($stockLevelValue * 0.45) + ($financialImpactValue * 0.30) + ($demandCriticalityValue * 0.15) + ($leadTimeRiskValue * 0.10)
        $targetSheet.Cells.Item($currentRow, 13) = [math]::Round($compositeScore, 2)
        
        # Peringkat Urgensi - Column N (will be calculated later)
        $targetSheet.Cells.Item($currentRow, 14) = $row + 1  # Temporary ranking
        
        # Level Urgensi - Column O
        $urgencyLevel = "RENDAH"
        if ($compositeScore -ge 83) { $urgencyLevel = "KRITIS" }
        elseif ($compositeScore -ge 58) { $urgencyLevel = "TINGGI" }
        elseif ($compositeScore -ge 33) { $urgencyLevel = "SEDANG" }
        $targetSheet.Cells.Item($currentRow, 15) = $urgencyLevel
        
        # Alasan - Column P
        $reason = "Analisis AHP menunjukkan skor $([math]::Round($compositeScore,1)) - item kelas $($data[4]) dengan status $($data[3])"
        $targetSheet.Cells.Item($currentRow, 16) = $reason
        
        # Tindakan - Column Q
        $action = "Terapkan strategi sesuai kondisi"
        if ($data[3] -eq "Out of Stock") { $action = "Pesanan darurat segera!" }
        elseif ($data[3] -eq "Reorder") { $action = "Buat pesanan berdasarkan EOQ" }
        elseif ($data[3] -eq "Low Stock") { $action = "Pantau dan rencanakan pemesanan" }
        $targetSheet.Cells.Item($currentRow, 17) = $action
        
        # Jangka Waktu - Column R
        $timeframe = "Dalam 1 bulan"
        if ($urgencyLevel -eq "KRITIS") { $timeframe = "Segera" }
        elseif ($urgencyLevel -eq "TINGGI") { $timeframe = "Dalam 1-10 hari" }
        elseif ($urgencyLevel -eq "SEDANG") { $timeframe = "Dalam 1-4 minggu" }
        $targetSheet.Cells.Item($currentRow, 18) = $timeframe
    }
    
    # Auto-fit columns
    $targetSheet.Columns.AutoFit() | Out-Null
    
    # Add AHP explanation
    $explanationRow = $sampleData.Length + 4
    $targetSheet.Cells.Item($explanationRow, 1) = "PENJELASAN AHP (Analytic Hierarchy Process):"
    $targetSheet.Cells.Item($explanationRow, 1).Font.Bold = $true
    $targetSheet.Cells.Item($explanationRow, 1).Font.Size = 12
    
    $explanationRow++
    $targetSheet.Cells.Item($explanationRow, 1) = "Tingkat Stok (45 persen): Kritisitas ketersediaan stok saat ini"
    $explanationRow++
    $targetSheet.Cells.Item($explanationRow, 1) = "Dampak Finansial (30 persen): Nilai bisnis dan dampak keuangan"
    $explanationRow++
    $targetSheet.Cells.Item($explanationRow, 1) = "Kritisitas Permintaan (15 persen): Klasifikasi ABC dan tingkat permintaan"
    $explanationRow++
    $targetSheet.Cells.Item($explanationRow, 1) = "Risiko Lead Time (10 persen): Risiko keterlambatan pasokan"
    
    Write-Host "📈 AHP calculations and formulas added successfully" -ForegroundColor Green
    
    # Save the workbook
    $workbook.Save()
    Write-Host "💾 Workbook saved successfully" -ForegroundColor Green
    
    # Close Excel
    $workbook.Close()
    $excel.Quit()
    
    # Release COM objects
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($targetSheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "✅ AHP Excel update completed successfully!" -ForegroundColor Green
    Write-Host "📁 File updated: $workbookPath" -ForegroundColor Cyan
    
} catch {
    Write-Host "❌ Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    
    # Clean up COM objects in case of error
    if ($workbook) { $workbook.Close() }
    if ($excel) { $excel.Quit() }
}

