# PowerShell script to fix Excel Total Variance Value formula
# This script corrects the formula to match the web application calculation

try {
    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Open the workbook
    $workbook = $excel.Workbooks.Open("$PWD\Stock_Opname_DSS_Template.xlsx")
    
    Write-Host "=== Fixing Excel Total Variance Value Formula ==="
    
    # Get the Summary Dashboard sheet
    $summarySheet = $workbook.Worksheets["Summary_Dashboard"]
    
    # Find and fix the Total Variance Value formula
    $found = $false
    for ($row = 1; $row -le 20; $row++) {
        for ($col = 1; $col -le 10; $col++) {
            $cell = $summarySheet.Cells($row, $col)
            if ($cell.Value -and $cell.Value.ToString().Contains("Total Variance")) {
                Write-Host "Found 'Total Variance Value' at row $row, column $col"
                
                # The formula should be in the next column
                $formulaCell = $summarySheet.Cells($row, $col + 1)
                $oldFormula = $formulaCell.Formula
                Write-Host "Old formula: $oldFormula"
                
                # Fix the formula to use proper range and SUMPRODUCT for ABS
                $newFormula = "=SUMPRODUCT(ABS('DSS-SPK_Analysis'!H2:H11))"
                $formulaCell.Formula = $newFormula
                Write-Host "New formula: $newFormula"
                
                $found = $true
                break
            }
        }
        if ($found) { break }
    }
    
    if (-not $found) {
        Write-Host "Total Variance Value cell not found, adding it..."
        # Add it at row 5 if not found
        $summarySheet.Cells(5, 1).Value = "Total Variance Value"
        $summarySheet.Cells(5, 2).Formula = "=SUMPRODUCT(ABS('DSS-SPK_Analysis'!H2:H11))"
        $summarySheet.Cells(5, 1).Font.Bold = $true
    }
    
    # Also fix the range to be more dynamic (handle up to 100 products)
    Write-Host "`nUpdating to handle dynamic product count..."
    $summarySheet.Cells(5, 2).Formula = "=SUMPRODUCT(ABS('DSS-SPK_Analysis'!H2:H101))"
    
    # Format the cell as currency
    $summarySheet.Cells(5, 2).NumberFormat = "[$Rp-421] #,##0"
    
    Write-Host "`n=== Additional Formula Improvements ==="
    
    # Also check and improve other formulas in Summary Dashboard
    $improvements = @(
        @{Row=2; Col=1; Label="Total Products"; Formula="=COUNTA('DSS-SPK_Analysis'!A2:A101)-COUNTBLANK('DSS-SPK_Analysis'!A2:A101)"},
        @{Row=3; Col=1; Label="Total Inventory Value"; Formula="=SUMPRODUCT('DSS-SPK_Analysis'!I2:I101)"},
        @{Row=4; Col=1; Label="Accuracy Rate %"; Formula="=100-((COUNTIFS('DSS-SPK_Analysis'!G2:G101,\">5\")+COUNTIFS('DSS-SPK_Analysis'!G2:G101,\"<-5\"))/COUNTA('DSS-SPK_Analysis'!A2:A101)*100)"}
    )
    
    foreach ($improvement in $improvements) {
        $summarySheet.Cells($improvement.Row, 1).Value = $improvement.Label
        $summarySheet.Cells($improvement.Row, 2).Formula = $improvement.Formula
        $summarySheet.Cells($improvement.Row, 1).Font.Bold = $true
        if ($improvement.Label.Contains("Value")) {
            $summarySheet.Cells($improvement.Row, 2).NumberFormat = "[$Rp-421] #,##0"
        }
        Write-Host "Updated: $($improvement.Label)"
    }
    
    # Save the workbook
    $workbook.Save()
    Write-Host "`n=== Excel file updated successfully! ==="
    
    # Close and cleanup
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Host "`n=== Testing the fix ==="
    Write-Host "Expected Total Variance Value: Rp 17,750,000"
    Write-Host "Please check the Excel file Summary Dashboard to verify the calculation matches."
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

