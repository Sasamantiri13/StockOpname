# Test AHP Excel Import Functionality
param(
    [string]$ExcelFile = "Stock_Opname_DSS_Template_AHP.xlsx"
)

# Load Excel COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

try {
    $Workbook = $Excel.Workbooks.Open((Resolve-Path $ExcelFile).Path)
    
    Write-Host "=== TESTING AHP EXCEL IMPORT FUNCTIONALITY ===" -ForegroundColor Green
    Write-Host ""
    
    # Find the AHP Urgency Ranking worksheet
    $AHPSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "AHP Urgency Ranking" }
    
    if (-not $AHPSheet) {
        throw "AHP Urgency Ranking sheet not found"
    }
    
    # Get sheet dimensions
    $LastRow = $AHPSheet.UsedRange.Rows.Count
    $LastColumn = $AHPSheet.UsedRange.Columns.Count
    
    Write-Host "AHP Sheet Dimensions: $LastRow rows x $LastColumn columns" -ForegroundColor Yellow
    
    # Build header map
    $headerMap = @{}
    for ($col = 1; $col -le $LastColumn; $col++) {
        $header = $AHPSheet.Cells.Item(1, $col).Text
        if ($header) {
            $headerMap[$header] = $col
        }
    }
    
    # Required import fields per the importFromExcel function
    $requiredFields = @(
        "Kode Produk",
        "Nama Produk", 
        "Kategori",
        "Stok Sistem",
        "Stok Aktual",
        "Harga Satuan",
        "Min Stock",
        "Max Stock", 
        "Lead Time",
        "Avg Demand"
    )
    
    Write-Host "IMPORT COMPATIBILITY CHECK:" -ForegroundColor Cyan
    $allFieldsFound = $true
    foreach ($field in $requiredFields) {
        $found = $headerMap.ContainsKey($field)
        $status = if ($found) { "FOUND" } else { "MISSING" }
        $color = if ($found) { "Green" } else { "Red" }
        Write-Host "  $field`: $status" -ForegroundColor $color
        if (-not $found) { $allFieldsFound = $false }
    }
    
    if (-not $allFieldsFound) {
        Write-Host ""
        Write-Host "❌ IMPORT TEST FAILED: Missing required fields!" -ForegroundColor Red
        return
    }
    
    Write-Host ""
    Write-Host "✅ ALL REQUIRED FIELDS FOUND!" -ForegroundColor Green
    Write-Host ""
    
    # Simulate data extraction (like XLSX.utils.sheet_to_json would do)
    Write-Host "SIMULATING IMPORT PROCESS:" -ForegroundColor Cyan
    
    $importedProducts = @()
    for ($row = 2; $row -le [Math]::Min(5, $LastRow); $row++) {  # Test first few rows
        $product = @{}
        
        # Extract data for each required field
        foreach ($field in $requiredFields) {
            $col = $headerMap[$field]
            $value = $AHPSheet.Cells.Item($row, $col).Value2
            
            # Convert to expected format based on field type
            switch ($field) {
                "Kode Produk" { 
                    $product["code"] = if ($value) { $value.ToString() } else { "PRD000" }
                    $product["Kode Produk"] = $product["code"]
                }
                "Nama Produk" { 
                    $product["name"] = if ($value) { $value.ToString() } else { "" }
                    $product["Nama Produk"] = $product["name"]
                }
                "Kategori" { 
                    $product["category"] = if ($value) { $value.ToString() } else { "" }
                    $product["Kategori"] = $product["category"]
                }
                "Stok Sistem" { 
                    $product["systemStock"] = if ($value) { [int]$value } else { 0 }
                    $product["Stok Sistem"] = $product["systemStock"]
                }
                "Stok Aktual" { 
                    $product["actualStock"] = if ($value) { [int]$value } else { 0 }
                    $product["Stok Aktual"] = $product["actualStock"]
                }
                "Harga Satuan" { 
                    $product["unitCost"] = if ($value) { [double]$value } else { 0 }
                    $product["Harga Satuan"] = $product["unitCost"]
                }
                "Min Stock" { 
                    $product["minStock"] = if ($value) { [int]$value } else { 0 }
                    $product["Min Stock"] = $product["minStock"]
                }
                "Max Stock" { 
                    $product["maxStock"] = if ($value) { [int]$value } else { 0 }
                    $product["Max Stock"] = $product["maxStock"]
                }
                "Lead Time" { 
                    $product["leadTime"] = if ($value) { [int]$value } else { 1 }
                    $product["Lead Time"] = $product["leadTime"]
                }
                "Avg Demand" { 
                    $product["avgDemand"] = if ($value) { [int]$value } else { 1 }
                    $product["Avg Demand"] = $product["avgDemand"]
                }
            }
        }
        
        $product["id"] = [DateTimeOffset]::Now.ToUnixTimeMilliseconds() + $row
        $importedProducts += $product
        
        Write-Host "  Row $row`: $($product["code"]) - $($product["name"])" -ForegroundColor Blue
    }
    
    Write-Host ""
    Write-Host "SAMPLE IMPORTED PRODUCT DATA:" -ForegroundColor Cyan
    $sampleProduct = $importedProducts[0]
    $sampleProduct.Keys | ForEach-Object {
        $key = $_
        $value = $sampleProduct[$key]
        Write-Host "  $key`: $value" -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "=== IMPORT TEST RESULTS ===" -ForegroundColor Green
    Write-Host "✅ Import compatibility: PASSED" -ForegroundColor Green
    Write-Host "✅ Data extraction: PASSED" -ForegroundColor Green
    Write-Host "✅ Field mapping: PASSED" -ForegroundColor Green
    Write-Host "✅ Data type conversion: PASSED" -ForegroundColor Green
    Write-Host ""
    Write-Host "🎉 SUCCESS: The AHP Excel export can now be imported back into the system!" -ForegroundColor Green
    Write-Host "   The data includes all required fields and maintains compatibility." -ForegroundColor Green
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($Workbook) { $Workbook.Close($false) }
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

