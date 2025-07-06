# Make AHP Excel Import Compatible
param(
    [string]$ExcelFile = "Stock_Opname_DSS_Template_AHP.xlsx"
)

# Load Excel COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

try {
    $Workbook = $Excel.Workbooks.Open((Resolve-Path $ExcelFile).Path)
    
    Write-Host "=== MAKING AHP EXCEL IMPORT COMPATIBLE ===" -ForegroundColor Green
    Write-Host ""
    
    # Find the AHP Urgency Ranking worksheet
    $AHPSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "AHP Urgency Ranking" }
    $InputSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "Input_Data" }
    
    if (-not $AHPSheet) {
        throw "AHP Urgency Ranking sheet not found"
    }
    
    if (-not $InputSheet) {
        throw "Input_Data sheet not found"
    }
    
    Write-Host "Found AHP Urgency Ranking sheet" -ForegroundColor Yellow
    Write-Host "Found Input_Data sheet" -ForegroundColor Yellow
    
    # Get the last column and row in AHP sheet
    $LastRow = $AHPSheet.UsedRange.Rows.Count
    $LastColumn = $AHPSheet.UsedRange.Columns.Count
    
    Write-Host "Current AHP sheet dimensions: $LastRow rows x $LastColumn columns"
    
    # Map of existing headers to find positions
    $headerMap = @{}
    for ($col = 1; $col -le $LastColumn; $col++) {
        $header = $AHPSheet.Cells.Item(1, $col).Text
        if ($header) {
            $headerMap[$header] = $col
        }
    }
    
    # Required fields for import and their sources
    $importFields = @{
        "Stok Sistem" = @{ "source" = "Input_Data"; "column" = "System_Stock"; "ahp_source" = $null }
        "Stok Aktual" = @{ "source" = "Input_Data"; "column" = "Actual_Stock"; "ahp_source" = $null }
        "Harga Satuan" = @{ "source" = "Input_Data"; "column" = "Unit_Cost"; "ahp_source" = $null }
        "Min Stock" = @{ "source" = "Input_Data"; "column" = "Min_Stock"; "ahp_source" = $null }
        "Max Stock" = @{ "source" = "Input_Data"; "column" = "Max_Stock"; "ahp_source" = $null }
        "Lead Time" = @{ "source" = "Input_Data"; "column" = "Lead_Time_Days"; "ahp_source" = $null }
        "Avg Demand" = @{ "source" = "Input_Data"; "column" = "Avg_Daily_Demand"; "ahp_source" = $null }
    }
    
    # Get Input_Data headers and data
    $InputLastRow = $InputSheet.UsedRange.Rows.Count
    $InputLastColumn = $InputSheet.UsedRange.Columns.Count
    
    $inputHeaderMap = @{}
    for ($col = 1; $col -le $InputLastColumn; $col++) {
        $header = $InputSheet.Cells.Item(1, $col).Text.Trim()
        if ($header) {
            $inputHeaderMap[$header] = $col
        }
    }
    
    Write-Host "Input_Data headers found:"
    $inputHeaderMap.Keys | ForEach-Object { Write-Host "  - $_" }
    
    # Check which import fields need to be added
    $fieldsToAdd = @()
    foreach ($field in $importFields.Keys) {
        if (-not $headerMap.ContainsKey($field)) {
            $fieldsToAdd += $field
        }
    }
    
    if ($fieldsToAdd.Count -eq 0) {
        Write-Host "All required import fields already exist!" -ForegroundColor Green
        return
    }
    
    Write-Host ""
    Write-Host "Adding missing import fields:" -ForegroundColor Cyan
    $fieldsToAdd | ForEach-Object { Write-Host "  + $_" -ForegroundColor Blue }
    
    # Add headers for missing fields
    $currentColumn = $LastColumn + 1
    foreach ($field in $fieldsToAdd) {
        $AHPSheet.Cells.Item(1, $currentColumn).Value2 = $field
        $AHPSheet.Cells.Item(1, $currentColumn).Font.Bold = $true
        $AHPSheet.Cells.Item(1, $currentColumn).Interior.Color = 15849925  # Light blue color
        
        Write-Host "Added header '$field' in column $currentColumn"
        $headerMap[$field] = $currentColumn
        $currentColumn++
    }
    
    # Build lookup table from Input_Data using Product_Code
    $inputData = @{}
    if ($InputLastRow -gt 1) {
        $productCodeCol = $inputHeaderMap["Product_Code"]
        for ($row = 2; $row -le $InputLastRow; $row++) {
            $productCode = $InputSheet.Cells.Item($row, $productCodeCol).Text.Trim()
            if ($productCode) {
                $rowData = @{}
                foreach ($header in $inputHeaderMap.Keys) {
                    $col = $inputHeaderMap[$header]
                    $value = $InputSheet.Cells.Item($row, $col).Value2
                    $rowData[$header] = $value
                }
                $inputData[$productCode] = $rowData
            }
        }
    }
    
    Write-Host ""
    Write-Host "Filling data for missing fields..." -ForegroundColor Yellow
    
    # Get the product code column in AHP sheet
    $ahpProductCodeCol = $headerMap["Kode Produk"]
    
    # Fill data for each row
    for ($row = 2; $row -le $LastRow; $row++) {
        $productCodeValue = $AHPSheet.Cells.Item($row, $ahpProductCodeCol).Value2
        $productCode = if ($productCodeValue) { $productCodeValue.ToString().Trim() } else { "" }
        
        if ($productCode -and $inputData.ContainsKey($productCode)) {
            $sourceData = $inputData[$productCode]
            
            foreach ($field in $fieldsToAdd) {
                $sourceColumn = $importFields[$field]["column"]
                $targetColumn = $headerMap[$field]
                
                if ($sourceData.ContainsKey($sourceColumn)) {
                    $value = $sourceData[$sourceColumn]
                    if ($value -ne $null) {
                        $AHPSheet.Cells.Item($row, $targetColumn).Value2 = $value
                    }
                    
                    # Format currency for Harga Satuan
                    if ($field -eq "Harga Satuan") {
                        $AHPSheet.Cells.Item($row, $targetColumn).NumberFormat = "_-`"Rp`"* #,##0_-;-`"Rp`"* #,##0_-;_-`"Rp`"* `"-`"_-;_-@_-"
                    }
                }
            }
        }
    }
    
    # Auto-fit columns
    $AHPSheet.UsedRange.Columns.AutoFit() | Out-Null
    
    # Save the workbook
    $Workbook.Save()
    
    Write-Host ""
    Write-Host "=== IMPORT COMPATIBILITY VERIFICATION ===" -ForegroundColor Green
    
    # Verify all required fields are now present
    $requiredFields = @(
        "Kode Produk", "Nama Produk", "Kategori",
        "Stok Sistem", "Stok Aktual", "Harga Satuan",
        "Min Stock", "Max Stock", "Lead Time", "Avg Demand"
    )
    
    $allFieldsPresent = $true
    foreach ($field in $requiredFields) {
        $exists = $headerMap.ContainsKey($field)
        $status = if ($exists) { "FOUND" } else { "MISSING" }
        $color = if ($exists) { "Green" } else { "Red" }
        Write-Host "  $status $field" -ForegroundColor $color
        if (-not $exists) { $allFieldsPresent = $false }
    }
    
    Write-Host ""
    if ($allFieldsPresent) {
        Write-Host "SUCCESS: AHP Excel file is now fully compatible with import functionality!" -ForegroundColor Green
        Write-Host "The exported AHP data can now be imported back into the system." -ForegroundColor Green
    } else {
        Write-Host "WARNING: Some required fields are still missing!" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Updated file saved: $ExcelFile" -ForegroundColor Cyan
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
} finally {
    if ($Workbook) { $Workbook.Close($false) }
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

