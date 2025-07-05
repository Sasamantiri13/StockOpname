# PowerShell script to create DSS Excel template with formulas
# Requires Excel to be installed

try {
    # Create Excel application
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Create new workbook
    $workbook = $excel.Workbooks.Add()
    
    # Sheet 1: Input Data
    $inputSheet = $workbook.Worksheets.Item(1)
    $inputSheet.Name = "Input_Data"
    
    # Headers for input data
    $headers = @(
        "Product_Code", "Product_Name", "Category", "System_Stock", "Actual_Stock", 
        "Unit_Cost", "Min_Stock", "Max_Stock", "Lead_Time_Days", "Avg_Daily_Demand",
        "Ordering_Cost", "Holding_Cost_Rate"
    )
    
    for ($i = 0; $i -lt $headers.Length; $i++) {
        $inputSheet.Cells.Item(1, $i + 1) = $headers[$i]
        $inputSheet.Cells.Item(1, $i + 1).Font.Bold = $true
        $inputSheet.Cells.Item(1, $i + 1).Interior.Color = 15917529  # Light blue
    }
    
    # Sample data
    $sampleData = @(
        @("PRD001", "Laptop Dell Inspiron", "Electronics", 50, 48, 8500000, 10, 100, 7, 8, 50000, 0.25),
        @("PRD002", "Mouse Wireless", "Accessories", 120, 125, 150000, 20, 200, 3, 15, 50000, 0.25),
        @("PRD003", "Keyboard Mechanical", "Accessories", 80, 75, 750000, 15, 150, 5, 12, 50000, 0.25)
    )
    
    for ($row = 0; $row -lt $sampleData.Length; $row++) {
        for ($col = 0; $col -lt $sampleData[$row].Length; $col++) {
            $inputSheet.Cells.Item($row + 2, $col + 1) = $sampleData[$row][$col]
        }
    }
    
    # Auto-fit columns
    $inputSheet.Columns.AutoFit() | Out-Null
    
    # Sheet 2: DSS Analysis
    $analysisSheet = $workbook.Worksheets.Add()
    $analysisSheet.Name = "DSS_Analysis"
    
    # Analysis headers
    $analysisHeaders = @(
        "Product_Code", "Product_Name", "Category", "System_Stock", "Actual_Stock",
        "Variance", "Variance_Pct", "Variance_Value", "Inventory_Value", 
        "Annual_Demand", "Safety_Stock", "Reorder_Point", "EOQ", 
        "Stock_Status", "Turnover_Ratio", "ABC_Class", "Cumulative_Pct"
    )
    
    for ($i = 0; $i -lt $analysisHeaders.Length; $i++) {
        $analysisSheet.Cells.Item(1, $i + 1) = $analysisHeaders[$i]
        $analysisSheet.Cells.Item(1, $i + 1).Font.Bold = $true
        $analysisSheet.Cells.Item(1, $i + 1).Interior.Color = 15917529
    }
    
    # Formulas for analysis (starting from row 2)
    $startRow = 2
    $endRow = 4  # Adjust based on sample data
    
    for ($row = $startRow; $row -le $endRow; $row++) {
        # Product Code (A)
        $analysisSheet.Cells.Item($row, 1).Formula = "=Input_Data!A$row"
        
        # Product Name (B)
        $analysisSheet.Cells.Item($row, 2).Formula = "=Input_Data!B$row"
        
        # Category (C)
        $analysisSheet.Cells.Item($row, 3).Formula = "=Input_Data!C$row"
        
        # System Stock (D)
        $analysisSheet.Cells.Item($row, 4).Formula = "=Input_Data!D$row"
        
        # Actual Stock (E)
        $analysisSheet.Cells.Item($row, 5).Formula = "=Input_Data!E$row"
        
        # Variance (F) = Actual - System
        $analysisSheet.Cells.Item($row, 6).Formula = "=E$row-D$row"
        
        # Variance % (G) = (Variance/System)*100
        $analysisSheet.Cells.Item($row, 7).Formula = "=IF(D$row<>0,F$row/D$row*100,0)"
        
        # Variance Value (H) = Variance * Unit Cost
        $analysisSheet.Cells.Item($row, 8).Formula = "=F$row*Input_Data!F$row"
        
        # Inventory Value (I) = Actual Stock * Unit Cost
        $analysisSheet.Cells.Item($row, 9).Formula = "=E$row*Input_Data!F$row"
        
        # Annual Demand (J) = Daily Demand * 365
        $analysisSheet.Cells.Item($row, 10).Formula = "=Input_Data!J$row*365"
        
        # Safety Stock (K) = Daily Demand * SQRT(Lead Time) * Service Level Factor (1.65 for 95%)
        $analysisSheet.Cells.Item($row, 11).Formula = "=CEILING(Input_Data!J$row*SQRT(Input_Data!I$row)*1.65,1)"
        
        # Reorder Point (L) = (Daily Demand * Lead Time) + Safety Stock
        $analysisSheet.Cells.Item($row, 12).Formula = "=(Input_Data!J$row*Input_Data!I$row)+K$row"
        
        # EOQ (M) = SQRT((2 * Annual Demand * Ordering Cost) / (Unit Cost * Holding Rate))
        $analysisSheet.Cells.Item($row, 13).Formula = "=CEILING(SQRT((2*J$row*Input_Data!K$row)/(Input_Data!F$row*Input_Data!L$row)),1)"
        
        # Stock Status (N)
        $analysisSheet.Cells.Item($row, 14).Formula = "=IF(E$row<=Input_Data!G$row,""Low Stock"",IF(E$row>=Input_Data!H$row,""Overstock"",IF(E$row<=L$row,""Reorder"",""Normal"")))"
        
        # Turnover Ratio (O) = Annual Demand / Actual Stock
        $analysisSheet.Cells.Item($row, 15).Formula = "=IF(E$row<>0,J$row/E$row,0)"
        
        # ABC Class (P) - Will be calculated after sorting
        $analysisSheet.Cells.Item($row, 16).Formula = "=IF(Q$row<=80,""A"",IF(Q$row<=95,""B"",""C""))"
        
        # Cumulative % (Q) - Placeholder, will need manual calculation or VBA
        $analysisSheet.Cells.Item($row, 17).Formula = "=SUMIF(I:I,"">=""&I$row,I:I)/SUM(I:I)*100"
    }
    
    # Auto-fit columns
    $analysisSheet.Columns.AutoFit() | Out-Null
    
    # Sheet 3: Summary Dashboard
    $summarySheet = $workbook.Worksheets.Add()
    $summarySheet.Name = "Summary_Dashboard"
    
    # Summary metrics
    $summarySheet.Cells.Item(1, 1) = "STOCK OPNAME DSS SUMMARY"
    $summarySheet.Cells.Item(1, 1).Font.Bold = $true
    $summarySheet.Cells.Item(1, 1).Font.Size = 16
    
    $summaryMetrics = @(
        @("Total Products", "=COUNTA(DSS_Analysis!A:A)-1"),
        @("Total Inventory Value", "=SUM(DSS_Analysis!I:I)"),
        @("Total Variance Value", "=SUM(ABS(DSS_Analysis!H:H))"),
        @("Accuracy Rate (%)", "=COUNTIF(DSS_Analysis!G:G,""<5"")/COUNTA(DSS_Analysis!G:G)*100"),
        @("Low Stock Items", "=COUNTIF(DSS_Analysis!N:N,""Low Stock"")"),
        @("Overstock Items", "=COUNTIF(DSS_Analysis!N:N,""Overstock"")"),
        @("Items Need Reorder", "=COUNTIF(DSS_Analysis!N:N,""Reorder"")"),
        @("ABC Class A Items", "=COUNTIF(DSS_Analysis!P:P,""A"")"),
        @("ABC Class B Items", "=COUNTIF(DSS_Analysis!P:P,""B"")"),
        @("ABC Class C Items", "=COUNTIF(DSS_Analysis!P:P,""C"")")
    )
    
    for ($i = 0; $i -lt $summaryMetrics.Length; $i++) {
        $summarySheet.Cells.Item($i + 3, 1) = $summaryMetrics[$i][0]
        $summarySheet.Cells.Item($i + 3, 1).Font.Bold = $true
        $summarySheet.Cells.Item($i + 3, 2).Formula = $summaryMetrics[$i][1]
        
        # Format currency for value fields
        if ($summaryMetrics[$i][0] -like "*Value*") {
            $summarySheet.Cells.Item($i + 3, 2).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
        }
        elseif ($summaryMetrics[$i][0] -like "*Rate*") {
            $summarySheet.Cells.Item($i + 3, 2).NumberFormat = "0.0%"
        }
    }
    
    # Sheet 4: Recommendations
    $recoSheet = $workbook.Worksheets.Add()
    $recoSheet.Name = "Recommendations"
    
    $recoSheet.Cells.Item(1, 1) = "AUTOMATED RECOMMENDATIONS"
    $recoSheet.Cells.Item(1, 1).Font.Bold = $true
    $recoSheet.Cells.Item(1, 1).Font.Size = 16
    
    # Headers for recommendations
    $recoHeaders = @("Priority", "Product", "Issue", "Recommendation", "Action Required")
    for ($i = 0; $i -lt $recoHeaders.Length; $i++) {
        $recoSheet.Cells.Item(3, $i + 1) = $recoHeaders[$i]
        $recoSheet.Cells.Item(3, $i + 1).Font.Bold = $true
        $recoSheet.Cells.Item(3, $i + 1).Interior.Color = 15917529
    }
    
    # Sample recommendations (manual entry - can be enhanced with complex formulas)
    $recoSheet.Cells.Item(4, 1) = "HIGH"
    $recoSheet.Cells.Item(4, 2) = "Items with Stock Status = Low Stock"
    $recoSheet.Cells.Item(4, 3) = "Stock below minimum threshold"
    $recoSheet.Cells.Item(4, 4) = "Order EOQ quantity immediately"
    $recoSheet.Cells.Item(4, 5) = "Place purchase order"
    
    $recoSheet.Cells.Item(5, 1) = "MEDIUM"
    $recoSheet.Cells.Item(5, 2) = "Items with |Variance %| > 10%"
    $recoSheet.Cells.Item(5, 3) = "High variance between system and actual"
    $recoSheet.Cells.Item(5, 4) = "Conduct detailed audit"
    $recoSheet.Cells.Item(5, 5) = "Schedule inventory audit"
    
    $recoSheet.Cells.Item(6, 1) = "LOW"
    $recoSheet.Cells.Item(6, 2) = "Class A items with low turnover"
    $recoSheet.Cells.Item(6, 3) = "High value items not moving fast"
    $recoSheet.Cells.Item(6, 4) = "Review inventory strategy"
    $recoSheet.Cells.Item(6, 5) = "Analyze demand patterns"
    
    # Auto-fit all columns in all sheets
    foreach ($sheet in $workbook.Worksheets) {
        $sheet.Columns.AutoFit() | Out-Null
    }
    
    # Save the workbook
    $filePath = "C:\Users\SASA\P-SPK-StockOpname-TSX\Stock_Opname_DSS_Template.xlsx"
    $workbook.SaveAs($filePath)
    
    Write-Host "Excel DSS template created successfully at: $filePath"
    
    # Clean up
    $workbook.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
} catch {
    Write-Error "Error creating Excel file: $($_.Exception.Message)"
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

