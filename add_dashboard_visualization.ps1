# PowerShell script to add visualizations to Excel DSS Dashboard
# This script adds various charts to enhance the Summary_Dashboard sheet

# Create Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    # Open the existing workbook
    $workbook = $excel.Workbooks.Open("C:\Users\SASA\P-SPK-StockOpname-TSX\Stock_Opname_DSS_Template.xlsx")
    
    # Get the Summary_Dashboard worksheet
    $dashboard = $workbook.Worksheets.Item("Summary_Dashboard")
    
    # Get the DSS Analysis worksheet for chart data
    $analysis = $workbook.Worksheets.Item("DSS-SPK_Analysis")
    
    Write-Host "Adding visualizations to dashboard..." -ForegroundColor Green
    
    # ============================================
    # 1. ABC Classification Pie Chart
    # ============================================
    Write-Host "Creating ABC Classification Pie Chart..." -ForegroundColor Yellow
    
    # Create pie chart for ABC classification
    $abcChart = $dashboard.ChartObjects().Add(450, 50, 300, 200)
    $abcChart.Chart.ChartType = 5  # xlPie
    $abcChart.Chart.SetSourceData($dashboard.Range("A15:B17"))  # ABC data range
    $abcChart.Chart.HasTitle = $true
    $abcChart.Chart.ChartTitle.Text = "ABC Classification Distribution"
    $abcChart.Chart.HasLegend = $true
    $abcChart.Chart.Legend.Position = -4107  # xlLegendPositionRight
    
    # Format pie chart colors
    $abcChart.Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 99, 132))   # Red for A
    $abcChart.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 205, 86))  # Yellow for B
    $abcChart.Chart.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(75, 192, 192))  # Green for C
    
    # ============================================
    # 2. Stock Status Bar Chart
    # ============================================
    Write-Host "Creating Stock Status Bar Chart..." -ForegroundColor Yellow
    
    $statusChart = $dashboard.ChartObjects().Add(50, 280, 350, 200)
    $statusChart.Chart.ChartType = 51  # xlColumnClustered
    $statusChart.Chart.SetSourceData($dashboard.Range("A10:B13"))  # Stock status data
    $statusChart.Chart.HasTitle = $true
    $statusChart.Chart.ChartTitle.Text = "Stock Status Distribution"
    $statusChart.Chart.HasLegend = $false
    
    # Format bar chart colors
    $statusChart.Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(54, 162, 235))   # Blue for Low Stock
    $statusChart.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 99, 132))   # Red for Overstock
    $statusChart.Chart.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 205, 86))  # Yellow for Reorder
    $statusChart.Chart.SeriesCollection(1).Points(4).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(75, 192, 192))  # Green for Normal
    
    # ============================================
    # 3. Inventory Value by Product (Top 5)
    # ============================================
    Write-Host "Creating Top 5 Inventory Value Chart..." -ForegroundColor Yellow
    
    # First, let's add labels and data for top 5 products by inventory value
    $dashboard.Cells.Item(20, 1).Value2 = "Top 5 Products by Inventory Value"
    $dashboard.Cells.Item(20, 1).Font.Bold = $true
    $dashboard.Cells.Item(20, 1).Font.Size = 12
    
    # Add sample data for top 5 products (in real scenario, this would be dynamic)
    $dashboard.Cells.Item(21, 1).Value2 = "Product"
    $dashboard.Cells.Item(21, 2).Value2 = "Inventory Value"
    $dashboard.Cells.Item(22, 1).Value2 = "=INDEX('DSS-SPK_Analysis'!A:A,2)"
    $dashboard.Cells.Item(22, 2).Value2 = "=INDEX('DSS-SPK_Analysis'!I:I,2)"
    $dashboard.Cells.Item(23, 1).Value2 = "=INDEX('DSS-SPK_Analysis'!A:A,3)"
    $dashboard.Cells.Item(23, 2).Value2 = "=INDEX('DSS-SPK_Analysis'!I:I,3)"
    $dashboard.Cells.Item(24, 1).Value2 = "=INDEX('DSS-SPK_Analysis'!A:A,4)"
    $dashboard.Cells.Item(24, 2).Value2 = "=INDEX('DSS-SPK_Analysis'!I:I,4)"
    $dashboard.Cells.Item(25, 1).Value2 = "=INDEX('DSS-SPK_Analysis'!A:A,5)"
    $dashboard.Cells.Item(25, 2).Value2 = "=INDEX('DSS-SPK_Analysis'!I:I,5)"
    $dashboard.Cells.Item(26, 1).Value2 = "=INDEX('DSS-SPK_Analysis'!A:A,6)"
    $dashboard.Cells.Item(26, 2).Value2 = "=INDEX('DSS-SPK_Analysis'!I:I,6)"
    
    $inventoryChart = $dashboard.ChartObjects().Add(450, 280, 350, 200)
    $inventoryChart.Chart.ChartType = 57  # xlBarClustered (horizontal bar)
    $inventoryChart.Chart.SetSourceData($dashboard.Range("A22:B26"))
    $inventoryChart.Chart.HasTitle = $true
    $inventoryChart.Chart.ChartTitle.Text = "Top 5 Products by Inventory Value"
    $inventoryChart.Chart.HasLegend = $false
    
    # ============================================
    # 4. Variance Analysis Chart
    # ============================================
    Write-Host "Creating Variance Analysis Chart..." -ForegroundColor Yellow
    
    $dashboard.Cells.Item(28, 1).Value2 = "Variance Analysis"
    $dashboard.Cells.Item(28, 1).Font.Bold = $true
    $dashboard.Cells.Item(28, 1).Font.Size = 12
    
    # Add variance categories
    $dashboard.Cells.Item(29, 1).Value2 = "Positive Variance"
    $dashboard.Cells.Item(29, 2).Value2 = "=SUMPRODUCT(('DSS-SPK_Analysis'!G2:G11>0)*'DSS-SPK_Analysis'!H2:H11)"
    $dashboard.Cells.Item(30, 1).Value2 = "Negative Variance"
    $dashboard.Cells.Item(30, 2).Value2 = "=SUMPRODUCT(('DSS-SPK_Analysis'!G2:G11<0)*ABS('DSS-SPK_Analysis'!H2:H11))"
    $dashboard.Cells.Item(31, 1).Value2 = "Zero Variance"
    $dashboard.Cells.Item(31, 2).Value2 = "=SUMPRODUCT(('DSS-SPK_Analysis'!G2:G11=0)*'DSS-SPK_Analysis'!H2:H11)"
    
    $varianceChart = $dashboard.ChartObjects().Add(50, 520, 350, 200)
    $varianceChart.Chart.ChartType = 5  # xlPie
    $varianceChart.Chart.SetSourceData($dashboard.Range("A29:B31"))
    $varianceChart.Chart.HasTitle = $true
    $varianceChart.Chart.ChartTitle.Text = "Variance Distribution"
    $varianceChart.Chart.HasLegend = $true
    $varianceChart.Chart.Legend.Position = -4107  # xlLegendPositionRight
    
    # ============================================
    # 5. Inventory Turnover Analysis
    # ============================================
    Write-Host "Creating Inventory Turnover Chart..." -ForegroundColor Yellow
    
    $dashboard.Cells.Item(33, 1).Value2 = "Inventory Turnover Analysis"
    $dashboard.Cells.Item(33, 1).Font.Bold = $true
    $dashboard.Cells.Item(33, 1).Font.Size = 12
    
    # Turnover categories
    $dashboard.Cells.Item(34, 1).Value2 = "High Turnover (>6)"
    $dashboard.Cells.Item(34, 2).Value2 = "=COUNTIFS('DSS-SPK_Analysis'!N2:N11,"">6"")"
    $dashboard.Cells.Item(35, 1).Value2 = "Medium Turnover (3-6)"
    $dashboard.Cells.Item(35, 2).Value2 = "=COUNTIFS('DSS-SPK_Analysis'!N2:N11,"">=3"",'DSS-SPK_Analysis'!N2:N11,""<=6"")"
    $dashboard.Cells.Item(36, 1).Value2 = "Low Turnover (<3)"
    $dashboard.Cells.Item(36, 2).Value2 = "=COUNTIFS('DSS-SPK_Analysis'!N2:N11,""<3"")"
    
    $turnoverChart = $dashboard.ChartObjects().Add(450, 520, 350, 200)
    $turnoverChart.Chart.ChartType = 51  # xlColumnClustered
    $turnoverChart.Chart.SetSourceData($dashboard.Range("A34:B36"))
    $turnoverChart.Chart.HasTitle = $true
    $turnoverChart.Chart.ChartTitle.Text = "Inventory Turnover Categories"
    $turnoverChart.Chart.HasLegend = $false
    
    # Format turnover chart colors
    $turnoverChart.Chart.SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(75, 192, 192))  # Green for High
    $turnoverChart.Chart.SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 205, 86))  # Yellow for Medium
    $turnoverChart.Chart.SeriesCollection(1).Points(3).Format.Fill.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(255, 99, 132))   # Red for Low
    
    # ============================================
    # 6. Add KPI Indicators with Visual Formatting
    # ============================================
    Write-Host "Adding KPI visual indicators..." -ForegroundColor Yellow
    
    # Add accuracy rate indicator with conditional formatting
    $dashboard.Cells.Item(38, 1).Value2 = "Accuracy Rate Indicator"
    $dashboard.Cells.Item(38, 1).Font.Bold = $true
    $dashboard.Cells.Item(39, 1).Value2 = "Status:"
    $dashboard.Cells.Item(39, 2).Value2 = "=IF(B6>95%,""Excellent"",IF(B6>90%,""Good"",IF(B6>80%,""Fair"",""Poor"")))"
    
    # Add conditional formatting for accuracy status
    $accuracyRange = $dashboard.Range("B39")
    $condition1 = $accuracyRange.FormatConditions.Add(1, 1, "=""Excellent""")  # xlCellValue, xlEqual
    $condition1.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)
    
    $condition2 = $accuracyRange.FormatConditions.Add(1, 1, "=""Good""")
    $condition2.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightBlue)
    
    $condition3 = $accuracyRange.FormatConditions.Add(1, 1, "=""Fair""")
    $condition3.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Yellow)
    
    $condition4 = $accuracyRange.FormatConditions.Add(1, 1, "=""Poor""")
    $condition4.Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightCoral)
    
    # ============================================
    # 7. Enhanced Dashboard Formatting
    # ============================================
    Write-Host "Applying enhanced formatting..." -ForegroundColor Yellow
    
    # Add borders and styling to chart sections
    $chartSections = @(
        $dashboard.Range("A1:D18"),    # Main KPIs
        $dashboard.Range("A20:B26"),   # Top 5 products
        $dashboard.Range("A28:B31"),   # Variance analysis
        $dashboard.Range("A33:B36"),   # Turnover analysis
        $dashboard.Range("A38:B39")    # Accuracy indicator
    )
    
    foreach ($section in $chartSections) {
        $section.Borders.LineStyle = 1  # xlContinuous
        $section.Borders.Weight = 2     # xlThin
    }
    
    # Auto-fit columns
    $dashboard.Columns("A:D").AutoFit()
    
    Write-Host "All visualizations added successfully!" -ForegroundColor Green
    
    # Save the workbook
    $workbook.Save()
    Write-Host "Excel file saved with dashboard visualizations!" -ForegroundColor Green
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
} finally {
    # Clean up
    if ($workbook) { $workbook.Close() }
    if ($excel) { 
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host "`nDashboard visualization enhancement completed!" -ForegroundColor Cyan
Write-Host "The following charts have been added to your Summary_Dashboard:" -ForegroundColor White
Write-Host "1. ABC Classification Pie Chart" -ForegroundColor Yellow
Write-Host "2. Stock Status Bar Chart" -ForegroundColor Yellow
Write-Host "3. Top 5 Products by Inventory Value" -ForegroundColor Yellow
Write-Host "4. Variance Distribution Pie Chart" -ForegroundColor Yellow
Write-Host "5. Inventory Turnover Categories" -ForegroundColor Yellow
Write-Host "6. KPI Status Indicators with Conditional Formatting" -ForegroundColor Yellow

