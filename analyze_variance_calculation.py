#!/usr/bin/env python3
"""
Analysis of Total Variance Value calculation differences between Web App and Excel Template
"""

# Sample data from the web application (default data)
products = [
    {
        'code': 'PRD001',
        'name': 'Laptop Dell Inspiron',
        'category': 'Electronics',
        'systemStock': 50,
        'actualStock': 48,
        'unitCost': 8500000,
        'minStock': 10,
        'maxStock': 100,
        'leadTime': 7,
        'avgDemand': 8
    },
    {
        'code': 'PRD002',
        'name': 'Mouse Wireless',  
        'category': 'Accessories',
        'systemStock': 120,
        'actualStock': 125,
        'unitCost': 150000,
        'minStock': 20,
        'maxStock': 200,
        'leadTime': 3,
        'avgDemand': 15
    }
]

print("=== Analysis of Total Variance Value Calculation ===\n")

print("Sample Data:")
for i, product in enumerate(products, 1):
    print(f"{i}. {product['name']}")
    print(f"   System Stock: {product['systemStock']:,}")
    print(f"   Actual Stock: {product['actualStock']:,}")
    print(f"   Unit Cost: Rp {product['unitCost']:,}")
    print()

print("Calculations:")
total_variance_web = 0
for i, product in enumerate(products, 1):
    variance = product['actualStock'] - product['systemStock']
    variance_value = variance * product['unitCost']
    abs_variance_value = abs(variance_value)
    
    print(f"{i}. {product['name']}")
    print(f"   Variance: {variance:,}")
    print(f"   Variance Value: Rp {variance_value:,}")
    print(f"   Absolute Variance Value: Rp {abs_variance_value:,}")
    
    total_variance_web += abs_variance_value
    print()

print(f"=== WEB APPLICATION RESULT ===")
print(f"Total Variance Value: Rp {total_variance_web:,}")
print()

print("=== EXCEL TEMPLATE ANALYSIS ===")
print("Current Excel Formula: =SUM(ABS('DSS-SPK_Analysis'!H1:H10))")
print("Issues identified:")
print("1. Range includes H1 (header), should start from H2")
print("2. Range H1:H10 assumes 10 rows of data, but we only have 2 sample products")
print("3. The ABS function should be applied to individual cells, not the entire range")
print()

print("=== CORRECTED EXCEL FORMULA ===")
print("Should be: =SUMPRODUCT(ABS('DSS-SPK_Analysis'!H2:H11))")
print("Or for dynamic range: =SUMPRODUCT(ABS(INDIRECT(\"DSS-SPK_Analysis!H2:H\"&(ROW()-1+COUNTA(DSS-SPK_Analysis!A:A)))))")
print()

print("=== RECOMMENDATIONS ===")
print("1. Fix Excel formula to exclude header row")
print("2. Use SUMPRODUCT(ABS(...)) instead of SUM(ABS(...)) for proper array handling")
print("3. Make the range dynamic to handle variable number of products")
print("4. Ensure both web and Excel use the same absolute value logic")

