# Stock Opname Decision Support System (DSS) - Excel Template Documentation

## Overview
This Excel template implements a comprehensive Decision Support System (SPK/DSS) for stock opname analysis based on international standards. The system provides automated calculations for inventory management, variance analysis, and strategic recommendations.

## International DSS Standards Implemented

### 1. ABC Analysis (Pareto Analysis)
**Purpose**: Classify inventory items based on their value contribution
**Formula**: Items are sorted by inventory value, then classified:
- **Class A**: Top 80% of total inventory value (typically 20% of items)
- **Class B**: Next 15% of total inventory value (typically 30% of items)  
- **Class C**: Remaining 5% of total inventory value (typically 50% of items)

**Excel Formula**: `=IF(Cumulative%<=80,"A",IF(Cumulative%<=95,"B","C"))`

### 2. Economic Order Quantity (EOQ)
**Purpose**: Determine optimal order quantity to minimize total inventory costs
**Formula**: EOQ = √((2 × Annual Demand × Ordering Cost) / (Unit Cost × Holding Cost Rate))

**Excel Formula**: `=CEILING(SQRT((2*AnnualDemand*OrderingCost)/(UnitCost*HoldingRate)),1)`

**Components**:
- Annual Demand = Daily Demand × 365
- Ordering Cost = Fixed cost per order (typically $50-$100)
- Holding Cost Rate = 20-30% of unit cost annually

### 3. Safety Stock Calculation
**Purpose**: Buffer stock to protect against demand variability and lead time uncertainty
**Formula**: Safety Stock = Daily Demand × √(Lead Time) × Service Level Factor

**Excel Formula**: `=CEILING(DailyDemand*SQRT(LeadTime)*1.65,1)`

**Service Level Factors**:
- 90% service level: 1.28
- 95% service level: 1.65 (recommended)
- 99% service level: 2.33

### 4. Reorder Point (ROP)
**Purpose**: Determine when to place new orders
**Formula**: ROP = (Daily Demand × Lead Time) + Safety Stock

**Excel Formula**: `=(DailyDemand*LeadTime)+SafetyStock`

### 5. Inventory Turnover Ratio
**Purpose**: Measure how efficiently inventory is managed
**Formula**: Turnover Ratio = Annual Demand / Average Inventory

**Excel Formula**: `=IF(ActualStock<>0,AnnualDemand/ActualStock,0)`

**Interpretation**:
- High turnover (>6): Fast-moving, efficient inventory
- Medium turnover (3-6): Moderate efficiency
- Low turnover (<3): Slow-moving, potential excess inventory

### 6. Variance Analysis
**Purpose**: Measure accuracy of inventory records
**Formulas**:
- Variance = Actual Stock - System Stock
- Variance % = (Variance / System Stock) × 100
- Variance Value = Variance × Unit Cost

**Excel Formulas**:
- `=ActualStock-SystemStock`
- `=IF(SystemStock<>0,Variance/SystemStock*100,0)`
- `=Variance*UnitCost`

## Sheet Structure

### Sheet 1: Input_Data
**Purpose**: Primary data entry sheet
**Columns**:
1. Product_Code: Unique identifier
2. Product_Name: Item description
3. Category: Product classification
4. System_Stock: Stock per system records
5. Actual_Stock: Physical count results
6. Unit_Cost: Cost per unit
7. Min_Stock: Minimum stock threshold
8. Max_Stock: Maximum stock threshold
9. Lead_Time_Days: Supplier lead time
10. Avg_Daily_Demand: Average daily consumption
11. Ordering_Cost: Cost per purchase order
12. Holding_Cost_Rate: Annual holding cost percentage

### Sheet 2: DSS_Analysis
**Purpose**: Automated calculations and analysis
**Key Formulas**:

```excel
# Variance Calculation
=E2-D2  # Actual - System

# Variance Percentage
=IF(D2<>0,F2/D2*100,0)

# Inventory Value
=E2*Input_Data!F2  # Actual Stock × Unit Cost

# Annual Demand
=Input_Data!J2*365  # Daily Demand × 365

# Safety Stock (95% service level)
=CEILING(Input_Data!J2*SQRT(Input_Data!I2)*1.65,1)

# Reorder Point
=(Input_Data!J2*Input_Data!I2)+K2

# EOQ
=CEILING(SQRT((2*J2*Input_Data!K2)/(Input_Data!F2*Input_Data!L2)),1)

# Stock Status
=IF(E2<=Input_Data!G2,"Low Stock",IF(E2>=Input_Data!H2,"Overstock",IF(E2<=L2,"Reorder","Normal")))

# Turnover Ratio
=IF(E2<>0,J2/E2,0)

# ABC Classification
=IF(Q2<=80,"A",IF(Q2<=95,"B","C"))
```

### Sheet 3: Summary_Dashboard
**Purpose**: Executive summary and KPIs
**Key Metrics**:
- Total Products: `=COUNTA(DSS_Analysis!A:A)-1`
- Total Inventory Value: `=SUM(DSS_Analysis!I:I)`
- Accuracy Rate: `=COUNTIF(DSS_Analysis!G:G,"<5")/COUNTA(DSS_Analysis!G:G)*100`
- Low Stock Items: `=COUNTIF(DSS_Analysis!N:N,"Low Stock")`

### Sheet 4: Recommendations
**Purpose**: Automated decision support recommendations

## Decision Rules and Thresholds

### Stock Status Classification
1. **Low Stock**: Actual Stock ≤ Minimum Stock
2. **Overstock**: Actual Stock ≥ Maximum Stock  
3. **Reorder**: Actual Stock ≤ Reorder Point
4. **Normal**: All other cases

### Variance Analysis Thresholds
- **Acceptable**: |Variance %| < 5%
- **Attention Required**: 5% ≤ |Variance %| < 10%
- **Investigation Required**: |Variance %| ≥ 10%

### ABC Analysis Guidelines
- **Class A Items**: Require tight control, frequent reviews
- **Class B Items**: Moderate control, periodic reviews
- **Class C Items**: Simple control, basic monitoring

## Usage Instructions

### 1. Data Entry
1. Open the Excel file
2. Go to "Input_Data" sheet
3. Enter your product data in the respective columns
4. Ensure all required fields are completed

### 2. Analysis Review
1. Switch to "DSS_Analysis" sheet
2. Review calculated metrics for each product
3. Pay attention to items with:
   - High variance percentages
   - Low stock status
   - Class A classification

### 3. Dashboard Monitoring
1. Check "Summary_Dashboard" for overall KPIs
2. Monitor accuracy rates and exception counts
3. Track inventory value trends

### 4. Action Implementation
1. Review "Recommendations" sheet
2. Prioritize actions based on urgency
3. Implement suggested corrective measures

## Best Practices

### Data Quality
- Ensure accurate physical counts
- Verify system stock accuracy
- Maintain current cost information
- Update demand patterns regularly

### Analysis Frequency
- Daily: Monitor critical items (Class A)
- Weekly: Review exception reports
- Monthly: Full DSS analysis
- Quarterly: Strategy review

### Continuous Improvement
- Track accuracy improvement over time
- Adjust safety stock levels based on performance
- Refine ABC classifications annually
- Update economic parameters regularly

## Formulas Reference Card

| Metric | Formula | Excel Implementation |
|--------|---------|---------------------|
| Variance | Actual - System | `=E2-D2` |
| Variance % | (Variance/System)*100 | `=IF(D2<>0,F2/D2*100,0)` |
| EOQ | √((2×D×S)/(H×P)) | `=CEILING(SQRT((2*J2*K2)/(F2*L2)),1)` |
| Safety Stock | d×√LT×z | `=CEILING(J2*SQRT(I2)*1.65,1)` |
| ROP | (d×LT)+SS | `=(J2*I2)+K2` |
| Turnover | D/Inventory | `=IF(E2<>0,J2/E2,0)` |

Where:
- D = Annual Demand
- S = Ordering Cost  
- H = Holding Cost Rate
- P = Unit Price
- d = Daily Demand
- LT = Lead Time
- z = Service Level Factor
- SS = Safety Stock

## Validation and Testing

### Formula Verification
1. Test with known data sets
2. Compare with manual calculations
3. Validate against industry benchmarks
4. Cross-check with existing systems

### Sensitivity Analysis
- Test different service levels
- Vary lead times and demand patterns
- Adjust holding cost assumptions
- Evaluate ABC threshold impacts

## Integration with Existing Systems

### Data Import
- Export data from ERP/WMS systems
- Use standardized templates
- Implement data validation rules
- Automate regular updates

### Report Export
- Generate management reports
- Create exception dashboards
- Automate email notifications
- Integrate with BI tools

This DSS template provides a comprehensive, standards-based approach to stock opname analysis and inventory management decision support.

