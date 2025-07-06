# Check AHP Excel Structure for Import Compatibility
param(
    [string]$ExcelFile = "Stock_Opname_DSS_Template_AHP.xlsx"
)

# Load Excel COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $false
$Excel.DisplayAlerts = $false

try {
    $Workbook = $Excel.Workbooks.Open((Resolve-Path $ExcelFile).Path)
    
    Write-Host "=== CHECKING AHP EXCEL STRUCTURE FOR IMPORT COMPATIBILITY ===" -ForegroundColor Green
    Write-Host ""
    
    # Check each worksheet
    foreach ($Worksheet in $Workbook.Worksheets) {
        Write-Host "Sheet: $($Worksheet.Name)" -ForegroundColor Yellow
        
        # Get used range
        $UsedRange = $Worksheet.UsedRange
        if ($UsedRange) {
            $LastRow = $UsedRange.Rows.Count
            $LastColumn = $UsedRange.Columns.Count
            
            Write-Host "  Dimensions: $LastRow rows x $LastColumn columns"
            
            # Get headers (first row)
            if ($LastRow -gt 0) {
                Write-Host "  Headers:"
                for ($col = 1; $col -le $LastColumn; $col++) {
                    $header = $Worksheet.Cells.Item(1, $col).Text
                    if ($header) {
                        Write-Host "    Column $col`: $header"
                    }
                }
            }
            
            # Check if this is the AHP sheet and analyze import compatibility
            if ($Worksheet.Name -eq "AHP Urgency Ranking") {
                Write-Host ""
                Write-Host "  === IMPORT COMPATIBILITY ANALYSIS ===" -ForegroundColor Cyan
                
                # Check for required import fields
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
                
                $availableFields = @()
                for ($col = 1; $col -le $LastColumn; $col++) {
                    $header = $Worksheet.Cells.Item(1, $col).Text
                    if ($header) {
                        $availableFields += $header
                    }
                }
                
                Write-Host "  Required fields for import:"
                foreach ($field in $requiredFields) {
                    $exists = $availableFields -contains $field
                    $status = if ($exists) { "FOUND" } else { "MISSING" }
                    $color = if ($exists) { "Green" } else { "Red" }
                    Write-Host "    $field`: $status" -ForegroundColor $color
                }
                
                Write-Host ""
                Write-Host "  Additional fields in AHP sheet:"
                foreach ($field in $availableFields) {
                    if ($requiredFields -notcontains $field) {
                        Write-Host "    + $field" -ForegroundColor Blue
                    }
                }
                
                # Check sample data
                if ($LastRow -gt 1) {
                    Write-Host ""
                    Write-Host "  Sample data (first data row):"
                    for ($col = 1; $col -le [Math]::Min(10, $LastColumn); $col++) {
                        $header = $Worksheet.Cells.Item(1, $col).Text
                        $value = $Worksheet.Cells.Item(2, $col).Text
                        if ($header) {
                            Write-Host "    $header`: $value"
                        }
                    }
                }
            }
        }
        Write-Host ""
    }
    
    Write-Host "=== RECOMMENDATIONS ===" -ForegroundColor Magenta
    Write-Host "1. Ensure all required import fields are present in the AHP sheet"
    Write-Host "2. Verify data formatting matches import expectations"
    Write-Host "3. Test the import functionality with the current AHP export"
    Write-Host ""
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    if ($Workbook) { $Workbook.Close($false) }
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}

