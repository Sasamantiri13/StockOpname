# Simple fix to make AHP Excel Import Compatible
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
    
    # Find the AHP Urgency Ranking worksheet
    $AHPSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "AHP Urgency Ranking" }
    $InputSheet = $Workbook.Worksheets | Where-Object { $_.Name -eq "Input_Data" }
    
    if (-not $AHPSheet) {
        throw "AHP Urgency Ranking sheet not found"
    }
    
    if (-not $InputSheet) {
        throw "Input_Data sheet not found"
    }
    
    # Get the last column in AHP sheet
    $LastRow = $AHPSheet.UsedRange.Rows.Count
    $LastColumn = $AHPSheet.UsedRange.Columns.Count
    
    Write-Host "Adding required import fields..."
    
    # Add missing headers manually
    $newHeaders = @("Stok Sistem", "Stok Aktual", "Harga Satuan", "Min Stock", "Max Stock", "Lead Time", "Avg Demand")
    $startColumn = $LastColumn + 1
    
    for ($i = 0; $i -lt $newHeaders.Length; $i++) {
        $col = $startColumn + $i
        $AHPSheet.Cells.Item(1, $col).Value2 = $newHeaders[$i]
        $AHPSheet.Cells.Item(1, $col).Font.Bold = $true
        $AHPSheet.Cells.Item(1, $col).Interior.Color = 15849925  # Light blue
        Write-Host "  Added: $($newHeaders[$i]) in column $col"
    }
    
    # Add sample data for the new columns based on Input_Data
    Write-Host "Adding sample data from Input_Data sheet..."
    
    # Map the columns
    $mappings = @{
        "Stok Sistem" = 22    # Column V
        "Stok Aktual" = 23    # Column W  
        "Harga Satuan" = 24   # Column X
        "Min Stock" = 25      # Column Y
        "Max Stock" = 26      # Column Z
        "Lead Time" = 27      # Column AA
        "Avg Demand" = 28     # Column AB
    }
    
    # Input data mapping
    $inputMappings = @{
        "System_Stock" = 4
        "Actual_Stock" = 5
        "Unit_Cost" = 6
        "Min_Stock" = 7
        "Max_Stock" = 8
        "Lead_Time_Days" = 9
        "Avg_Daily_Demand" = 10
    }
    
    # Copy data from Input_Data sheet
    for ($row = 2; $row -le $LastRow; $row++) {
        try {
            # Get product code to match
            $productCodeValue = $AHPSheet.Cells.Item($row, 2).Value2
            $productCode = if ($productCodeValue) { $productCodeValue.ToString() } else { "" }
            
            # Find matching row in Input_Data
            $inputRow = -1
            for ($inputRowNum = 2; $inputRowNum -le $InputSheet.UsedRange.Rows.Count; $inputRowNum++) {
                $inputCodeValue = $InputSheet.Cells.Item($inputRowNum, 1).Value2
                $inputProductCode = if ($inputCodeValue) { $inputCodeValue.ToString() } else { "" }
                if ($inputProductCode -eq $productCode) {
                    $inputRow = $inputRowNum
                    break
                }
            }
            
            if ($inputRow -gt 0) {
                # Copy the data
                $AHPSheet.Cells.Item($row, 22).Value2 = $InputSheet.Cells.Item($inputRow, 4).Value2  # System_Stock
                $AHPSheet.Cells.Item($row, 23).Value2 = $InputSheet.Cells.Item($inputRow, 5).Value2  # Actual_Stock
                $AHPSheet.Cells.Item($row, 24).Value2 = $InputSheet.Cells.Item($inputRow, 6).Value2  # Unit_Cost
                $AHPSheet.Cells.Item($row, 25).Value2 = $InputSheet.Cells.Item($inputRow, 7).Value2  # Min_Stock
                $AHPSheet.Cells.Item($row, 26).Value2 = $InputSheet.Cells.Item($inputRow, 8).Value2  # Max_Stock
                $AHPSheet.Cells.Item($row, 27).Value2 = $InputSheet.Cells.Item($inputRow, 9).Value2  # Lead_Time_Days
                $AHPSheet.Cells.Item($row, 28).Value2 = $InputSheet.Cells.Item($inputRow, 10).Value2 # Avg_Daily_Demand
                
                # Format currency for Harga Satuan
                $AHPSheet.Cells.Item($row, 24).NumberFormat = "_-`"Rp`"* #,##0_-;-`"Rp`"* #,##0_-;_-`"Rp`"* `"-`"_-;_-@_-"
            }
        } catch {
            Write-Host "  Warning: Could not process row $row - $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    
    # Auto-fit columns
    $AHPSheet.UsedRange.Columns.AutoFit() | Out-Null
    
    # Save the workbook
    $Workbook.Save()
    
    Write-Host ""
    Write-Host "SUCCESS: AHP Excel file is now compatible with import functionality!" -ForegroundColor Green
    Write-Host "Added fields: Stok Sistem, Stok Aktual, Harga Satuan, Min Stock, Max Stock, Lead Time, Avg Demand" -ForegroundColor Green
    Write-Host "File saved: $ExcelFile" -ForegroundColor Cyan
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($Workbook) { $Workbook.Close($false) }
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

