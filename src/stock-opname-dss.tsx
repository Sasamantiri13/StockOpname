import React, { useState, useEffect, useCallback } from 'react';
import { Download, Upload, Plus, Trash2, Calculator, AlertTriangle, CheckCircle, TrendingUp, Package, BarChart3 } from 'lucide-react';
import * as XLSX from 'xlsx';
import { ResponsiveContainer, PieChart, Pie, Cell, XAxis, YAxis, Tooltip, BarChart, Bar } from 'recharts';

// Charts availability flag
const CHARTS_ENABLED = false; // Temporarily disabled to fix white screen

// Create fallback components
const ChartFallback = ({ title, data }: { title: string; data?: any }) => (
  <div className="h-64 flex items-center justify-center bg-gray-50 rounded">
    <div className="text-center">
      <BarChart3 className="h-12 w-12 text-gray-400 mx-auto mb-2" />
      <h3 className="font-medium text-gray-600">{title}</h3>
      <p className="text-sm text-gray-500">Chart visualization available soon</p>
    </div>
  </div>
);

// Definisikan tipe data untuk produk
interface Product {
  id: number;
  code: string;
  name: string;
  category: string;
  systemStock: number;
  actualStock: number;
  unitCost: number;
  minStock: number;
  maxStock: number;
  leadTime: number;
  avgDemand: number;
}

// Definisikan tipe data untuk hasil analisis
interface Analysis extends Product {
  variance: number;
  variancePercentage: number;
  varianceValue: number;
  inventoryValue: number;
  safetyStock: number;
  reorderPoint: number;
  eoq: number;
  stockStatus: string;
  turnoverRatio: number;
  annualDemand: number;
  abcClass: string;
  cumulativePercentage: number;
}

// Definisikan tipe data untuk ringkasan
interface Summary {
  totalProducts?: number;
  totalInventoryValue?: number;
  totalVarianceValue?: number;
  accuracyRate?: number;
  lowStockItems?: number;
  overstockItems?: number;
  reorderItems?: number;
}

// Definisikan tipe data untuk rekomendasi
interface Recommendation {
  type: 'urgent' | 'warning' | 'audit' | 'optimization';
  product: string;
  message: string;
  action: string;
}

// Definisikan tipe data untuk urgency ranking
interface UrgencyItem {
  id: number;
  code: string;
  name: string;
  category: string;
  stockStatus: string;
  abcClass: string;
  urgencyScore: number;
  urgencyLevel: 'CRITICAL' | 'HIGH' | 'MEDIUM' | 'LOW';
  daysUntilStockout?: number;
  businessImpact: number;
  reason: string;
  recommendedAction: string;
  timeframe: string;
}

const StockOpnameDSS: React.FC = () => {
  const [products, setProducts] = useState<Product[]>([
    // High-value items (Class A) - Electronics
    {
      id: 1,
      code: 'ELC001',
      name: 'MacBook Pro 16" M3',
      category: 'Electronics',
      systemStock: 25,
      actualStock: 23,
      unitCost: 35000000,
      minStock: 5,
      maxStock: 40,
      leadTime: 14,
      avgDemand: 3
    },
    {
      id: 2,
      code: 'ELC002',
      name: 'iPhone 15 Pro Max',
      category: 'Electronics',
      systemStock: 40,
      actualStock: 42,
      unitCost: 20000000,
      minStock: 10,
      maxStock: 60,
      leadTime: 10,
      avgDemand: 5
    },
    {
      id: 3,
      code: 'ELC003',
      name: 'Samsung Galaxy S24 Ultra',
      category: 'Electronics',
      systemStock: 30,
      actualStock: 8, // LOW STOCK scenario
      unitCost: 18000000,
      minStock: 15,
      maxStock: 50,
      leadTime: 7,
      avgDemand: 6
    },
    {
      id: 4,
      code: 'ELC004',
      name: 'Dell XPS 13 Laptop',
      category: 'Electronics',
      systemStock: 20,
      actualStock: 18,
      unitCost: 15000000,
      minStock: 8,
      maxStock: 35,
      leadTime: 12,
      avgDemand: 4
    },
    // Medium-value items (Class B) - Accessories & Peripherals
    {
      id: 5,
      code: 'ACC001',
      name: 'Logitech MX Master 3S',
      category: 'Accessories',
      systemStock: 80,
      actualStock: 85,
      unitCost: 1200000,
      minStock: 20,
      maxStock: 120,
      leadTime: 5,
      avgDemand: 12
    },
    {
      id: 6,
      code: 'ACC002',
      name: 'Mechanical Keyboard RGB',
      category: 'Accessories',
      systemStock: 60,
      actualStock: 15, // REORDER scenario
      unitCost: 800000,
      minStock: 25,
      maxStock: 100,
      leadTime: 7,
      avgDemand: 8
    },
    {
      id: 7,
      code: 'ACC003',
      name: 'Webcam 4K Pro',
      category: 'Accessories',
      systemStock: 45,
      actualStock: 110, // OVERSTOCK scenario
      unitCost: 2500000,
      minStock: 15,
      maxStock: 70,
      leadTime: 6,
      avgDemand: 7
    },
    {
      id: 8,
      code: 'ACC004',
      name: 'Wireless Headphones',
      category: 'Accessories',
      systemStock: 90,
      actualStock: 88,
      unitCost: 1500000,
      minStock: 30,
      maxStock: 150,
      leadTime: 4,
      avgDemand: 15
    },
    // Low-value items (Class C) - Basic accessories
    {
      id: 9,
      code: 'BSC001',
      name: 'USB Cable Type-C',
      category: 'Cables',
      systemStock: 200,
      actualStock: 195,
      unitCost: 50000,
      minStock: 50,
      maxStock: 300,
      leadTime: 3,
      avgDemand: 25
    },
    {
      id: 10,
      code: 'BSC002',
      name: 'Basic Mouse Pad',
      category: 'Accessories',
      systemStock: 150,
      actualStock: 25, // LOW STOCK scenario
      unitCost: 25000,
      minStock: 40,
      maxStock: 200,
      leadTime: 2,
      avgDemand: 20
    },
    {
      id: 11,
      code: 'BSC003',
      name: 'Screen Cleaning Kit',
      category: 'Maintenance',
      systemStock: 100,
      actualStock: 350, // MAJOR OVERSTOCK scenario
      unitCost: 75000,
      minStock: 30,
      maxStock: 150,
      leadTime: 3,
      avgDemand: 12
    },
    {
      id: 12,
      code: 'BSC004',
      name: 'Laptop Stand Adjustable',
      category: 'Accessories',
      systemStock: 75,
      actualStock: 72,
      unitCost: 300000,
      minStock: 20,
      maxStock: 120,
      leadTime: 5,
      avgDemand: 10
    },
    // Office supplies (Class C)
    {
      id: 13,
      code: 'OFF001',
      name: 'Wireless Presenter Remote',
      category: 'Office',
      systemStock: 30,
      actualStock: 8, // REORDER scenario
      unitCost: 450000,
      minStock: 12,
      maxStock: 50,
      leadTime: 4,
      avgDemand: 5
    },
    {
      id: 14,
      code: 'OFF002',
      name: 'Document Camera',
      category: 'Office',
      systemStock: 15,
      actualStock: 0, // OUT OF STOCK scenario
      unitCost: 3500000,
      minStock: 3,
      maxStock: 25,
      leadTime: 14,
      avgDemand: 2
    },
    {
      id: 15,
      code: 'OFF003',
      name: 'Portable Projector',
      category: 'Office',
      systemStock: 20,
      actualStock: 22,
      unitCost: 4500000,
      minStock: 5,
      maxStock: 30,
      leadTime: 10,
      avgDemand: 3
    },
    // Gaming category (Mixed classes)
    {
      id: 16,
      code: 'GAM001',
      name: 'Gaming Monitor 27" 144Hz',
      category: 'Gaming',
      systemStock: 35,
      actualStock: 33,
      unitCost: 8500000,
      minStock: 8,
      maxStock: 50,
      leadTime: 9,
      avgDemand: 6
    },
    {
      id: 17,
      code: 'GAM002',
      name: 'Gaming Chair Ergonomic',
      category: 'Gaming',
      systemStock: 25,
      actualStock: 12, // LOW STOCK scenario
      unitCost: 3200000,
      minStock: 15,
      maxStock: 40,
      leadTime: 14,
      avgDemand: 4
    },
    {
      id: 18,
      code: 'GAM003',
      name: 'RGB Gaming Mousepad',
      category: 'Gaming',
      systemStock: 80,
      actualStock: 180, // OVERSTOCK scenario
      unitCost: 400000,
      minStock: 25,
      maxStock: 120,
      leadTime: 3,
      avgDemand: 18
    },
    // Network equipment (High value)
    {
      id: 19,
      code: 'NET001',
      name: 'Enterprise Router Cisco',
      category: 'Networking',
      systemStock: 10,
      actualStock: 9,
      unitCost: 25000000,
      minStock: 3,
      maxStock: 15,
      leadTime: 21,
      avgDemand: 1
    },
    {
      id: 20,
      code: 'NET002',
      name: 'Managed Switch 24-Port',
      category: 'Networking',
      systemStock: 20,
      actualStock: 2, // REORDER scenario
      unitCost: 12000000,
      minStock: 5,
      maxStock: 30,
      leadTime: 14,
      avgDemand: 3
    }
  ]);

  const [analysis, setAnalysis] = useState<Analysis[]>([]);
  const [summary, setSummary] = useState<Summary>({});
  const [recommendations, setRecommendations] = useState<Recommendation[]>([]);
  const [urgencyRankings, setUrgencyRankings] = useState<UrgencyItem[]>([]);
  
  // Filter states for Hasil Analisis SPK
  const [filters, setFilters] = useState({
    searchTerm: '',
    stockStatus: 'All',
    abcClass: 'All',
    category: 'All',
    varianceType: 'All'
  });

  // Fungsi perhitungan SPK
  const calculateAnalysis = useCallback(() => {
    try {
      const analysisData: Analysis[] = products.map(product => {
        // Safety checks for all numeric values
        const safeSystemStock = Math.max(0, product.systemStock || 0);
        const safeActualStock = Math.max(0, product.actualStock || 0);
        const safeUnitCost = Math.max(0, product.unitCost || 0);
        const safeMinStock = Math.max(1, product.minStock || 1);
        const safeMaxStock = Math.max(safeMinStock + 1, product.maxStock || safeMinStock + 10);
        const safeLeadTime = Math.max(1, product.leadTime || 1);
        const safeAvgDemand = Math.max(1, product.avgDemand || 1);
        
        const variance = safeActualStock - safeSystemStock;
        const variancePercentage = safeSystemStock > 0 ? (variance / safeSystemStock) * 100 : 0;
        const varianceValue = variance * safeUnitCost;
      
        // ABC Analysis berdasarkan nilai inventory (using safe values)
        const inventoryValue = safeActualStock * safeUnitCost;
        
        // Safety Stock calculation (using safe values)
        const safetyStock = Math.ceil(safeAvgDemand * Math.sqrt(safeLeadTime) * 1.65); // 95% service level
        
        // Reorder Point (using safe values)
        const reorderPoint = (safeAvgDemand * safeLeadTime) + safetyStock;
        
        // Economic Order Quantity (EOQ) - simplified (using safe values)
        const annualDemand = safeAvgDemand * 365;
        const orderingCost = 50000; // Rp 50,000 per order
        const holdingCostRate = 0.25; // 25% of unit cost
        const holdingCost = Math.max(1, safeUnitCost * holdingCostRate); // Prevent division by zero
        const eoq = Math.sqrt((2 * annualDemand * orderingCost) / holdingCost);
      
      // Stock Status - Fixed logic with proper priority and validation
      let stockStatus = 'Normal';
      
      // Ensure we have valid threshold values
      const validMinStock = product.minStock > 0 ? product.minStock : 0;
      const validMaxStock = product.maxStock > validMinStock ? product.maxStock : validMinStock + 50;
      const validReorderPoint = reorderPoint > validMinStock ? reorderPoint : validMinStock + 5;
      
      // Priority logic (most critical first)
      if (product.actualStock <= 0) {
        stockStatus = 'Out of Stock';
      } else if (product.actualStock <= validMinStock) {
        stockStatus = 'Low Stock';
      } else if (product.actualStock >= validMaxStock) {
        stockStatus = 'Overstock';
      } else if (product.actualStock <= validReorderPoint && validReorderPoint > validMinStock) {
        stockStatus = 'Reorder';
      }
      
      // Turnover ratio
      const turnoverRatio = product.actualStock > 0 ? annualDemand / product.actualStock : 0;
      
    
      return {
        ...product,
        variance,
        variancePercentage,
        varianceValue,
        inventoryValue,
        safetyStock,
        reorderPoint,
        eoq: Math.ceil(eoq),
        stockStatus,
        turnoverRatio,
        annualDemand,
        abcClass: '', // Initial value
        cumulativePercentage: 0 // Initial value
      };
    });

    // ABC Classification
    const sortedByValue = [...analysisData].sort((a, b) => b.inventoryValue - a.inventoryValue);
    const totalValue = sortedByValue.reduce((sum, item) => sum + item.inventoryValue, 0);
    
    let cumulativeValue = 0;
    const classifiedData = sortedByValue.map(item => {
      cumulativeValue += item.inventoryValue;
      const cumulativePercentage = totalValue > 0 ? (cumulativeValue / totalValue) * 100 : 0;
      
      let abcClass = 'C';
      if (cumulativePercentage <= 80) abcClass = 'A';
      else if (cumulativePercentage <= 95) abcClass = 'B';
      
      return { ...item, abcClass, cumulativePercentage };
    });

    // Sort back to original order
    const finalAnalysis = analysisData.map(item => {
      const classified = classifiedData.find(c => c.id === item.id);
      return { ...item, abcClass: classified!.abcClass, cumulativePercentage: classified!.cumulativePercentage };
    });

    setAnalysis(finalAnalysis);
    
    // Calculate summary
    const totalVarianceValue = finalAnalysis.reduce((sum, item) => sum + Math.abs(item.varianceValue), 0);
    const totalInventoryValue = finalAnalysis.reduce((sum, item) => sum + item.inventoryValue, 0);
    const accuracyRate = finalAnalysis.length > 0 ? ((finalAnalysis.length - finalAnalysis.filter(item => Math.abs(item.variancePercentage) > 5).length) / finalAnalysis.length) * 100 : 100;
    
    setSummary({
      totalProducts: finalAnalysis.length,
      totalInventoryValue,
      totalVarianceValue,
      accuracyRate,
      lowStockItems: finalAnalysis.filter(item => item.stockStatus === 'Low Stock').length,
      overstockItems: finalAnalysis.filter(item => item.stockStatus === 'Overstock').length,
      reorderItems: finalAnalysis.filter(item => item.stockStatus === 'Reorder').length
    });

    // Generate recommendations
    const recs: Recommendation[] = [];
    finalAnalysis.forEach(item => {
      if (item.stockStatus === 'Low Stock') {
        recs.push({
          type: 'urgent',
          product: item.name,
          message: `Stock rendah! Segera order ${item.eoq} unit (EOQ) untuk ${item.name}`,
          action: 'immediate_order'
        });
      }
      if (item.stockStatus === 'Overstock') {
        recs.push({
          type: 'warning',
          product: item.name,
          message: `Overstock detected untuk ${item.name}. Pertimbangkan promosi atau redistribusi`,
          action: 'reduce_stock'
        });
      }
      if (Math.abs(item.variancePercentage) > 10) {
        recs.push({
          type: 'audit',
          product: item.name,
          message: `Selisih besar (${item.variancePercentage.toFixed(1)}%) pada ${item.name}. Perlu audit mendalam`,
          action: 'investigate'
        });
      }
      if (item.abcClass === 'A' && item.turnoverRatio < 4) {
        recs.push({
          type: 'optimization',
          product: item.name,
          message: `Item kelas A dengan turnover rendah: ${item.name}. Evaluasi strategi inventory`,
          action: 'optimize_strategy'
        });
      }
    });
    
setRecommendations(recs);

    // Calculate urgency ranking
    const urgencyRankings: UrgencyItem[] = finalAnalysis.filter(item => item.stockStatus !== 'Normal').map(item => {
      // Base urgency score on stock status and ABC class
      let urgencyScore = 0;
      let urgencyLevel: 'CRITICAL' | 'HIGH' | 'MEDIUM' | 'LOW' = 'LOW';
      let reason = '';
      let recommendedAction = '';
      let timeframe = 'Within 1 month';
      let daysUntilStockout;
      let businessImpact = item.inventoryValue || 0;

      switch (item.stockStatus) {
        case 'Out of Stock':
          urgencyScore += 10;
          urgencyLevel = 'CRITICAL';
          reason = 'Out of stock! Immediate action required.';
          recommendedAction = `Order at least ${item.eoq} units now!`;
          timeframe = 'Immediate';
          break;
        case 'Low Stock':
          urgencyScore += 6;
          urgencyLevel = 'HIGH';
          reason = 'Low stock, approaching minimum levels.';
          recommendedAction = `Consider ordering ${item.eoq} units.`;
          timeframe = 'Within 1 week';
          daysUntilStockout = Math.floor(item.actualStock / item.avgDemand);
          break;
        case 'Reorder':
          urgencyScore += 8;
          urgencyLevel = 'HIGH';
          reason = 'Stock at reorder level, plan immediate restock.';
          recommendedAction = `Place an order for ${item.eoq} units.`;
          timeframe = 'Within 10 days';
          break;
        case 'Overstock':
          urgencyScore += 4;
          urgencyLevel = 'MEDIUM';
          reason = 'Overstock situation, reduce holding costs.';
          recommendedAction = 'Run promotions or redistribute inventory.';
          timeframe = 'Within 1 month';
          break;
      }

      // Intensify urgency based on ABC classification
      if (item.abcClass === 'A') {
        urgencyScore += 2;
        businessImpact *= 1.5;
      } else if (item.abcClass === 'B') {
        urgencyScore += 1;
        businessImpact *= 1.2;
      }
  
      return { 
        id: item.id,
        code: item.code,
        name: item.name,
        category: item.category,
        stockStatus: item.stockStatus,
        abcClass: item.abcClass,
        urgencyScore,
        urgencyLevel,
        daysUntilStockout,
        businessImpact,
        reason,
        recommendedAction,
        timeframe
      };
    });

    // Sort urgency rankings by urgency score (highest first) and save to state
    const sortedUrgencyRankings = urgencyRankings.sort((a, b) => b.urgencyScore - a.urgencyScore);
    setUrgencyRankings(sortedUrgencyRankings);
    
    console.table(sortedUrgencyRankings);
    
    } catch (error) {
      console.error('Error in calculateAnalysis:', error);
      // Set safe defaults if calculation fails
      setAnalysis([]);
      setSummary({
        totalProducts: products.length,
        totalInventoryValue: 0,
        totalVarianceValue: 0,
        accuracyRate: 100,
        lowStockItems: 0,
        overstockItems: 0,
        reorderItems: 0
      });
      setRecommendations([]);
    }
  }, [products]);

  useEffect(() => {
    calculateAnalysis();
  }, [calculateAnalysis]);

  const addProduct = () => {
    try {
      const newProduct: Product = {
        id: Date.now(),
        code: `PRD${String(products.length + 1).padStart(3, '0')}`,
        name: '',
        category: '',
        systemStock: 0,
        actualStock: 0,
        unitCost: 0,
        minStock: 1, // Minimum value to prevent division by zero
        maxStock: 10, // Default max stock
        leadTime: 1, // Minimum lead time to prevent zero division
        avgDemand: 1 // Minimum demand to prevent zero division
      };
      setProducts([...products, newProduct]);
    } catch (error) {
      console.error('Error adding product:', error);
      alert('Error adding product. Please try again.');
    }
  };

  const deleteProduct = (id: number) => {
    setProducts(products.filter(p => p.id !== id));
  };

  // Function to show action alerts for urgency ranking
  const showActionAlert = (item: Analysis) => {
    let message = '';
    
    switch (item.stockStatus) {
      case 'Out of Stock':
        message = `🚨 EMERGENCY ORDER REQUIRED\n\nProduct: ${item.name} (${item.code})\nStatus: OUT OF STOCK - No inventory available!\nRecommended Order Quantity (EOQ): ${item.eoq} units\nEstimated Lead Time: ${item.leadTime} days\n\nIMMEDIATE ACTIONS:\n1. Contact supplier NOW\n2. Place emergency order for ${item.eoq} units\n3. Check if customers can wait\n4. Consider substitute products\n\nThis is a CRITICAL situation affecting sales!`;
        break;
      case 'Reorder':
        message = `🔄 REORDER ACTION REQUIRED\n\nProduct: ${item.name} (${item.code})\nCurrent Stock: ${item.actualStock} units\nReorder Point: ${item.reorderPoint} units\nRecommended Order Quantity (EOQ): ${item.eoq} units\nLead Time: ${item.leadTime} days\n\nACTIONS NEEDED:\n1. Place order for ${item.eoq} units within 10 days\n2. Monitor daily consumption\n3. Contact supplier to confirm delivery\n4. Update inventory forecasts`;
        break;
      case 'Low Stock':
        message = `⚡ LOW STOCK WARNING\n\nProduct: ${item.name} (${item.code})\nCurrent Stock: ${item.actualStock} units\nMinimum Stock: ${item.minStock} units\nShortfall: ${item.minStock - item.actualStock} units\nEstimated Days Until Stockout: ~${Math.floor(item.actualStock / item.avgDemand)} days\n\nRECOMMENDED ACTIONS:\n1. Monitor consumption closely\n2. Consider ordering ${item.eoq} units soon\n3. Set up automatic alerts\n4. Review demand patterns\n5. Check supplier availability`;
        break;
      case 'Overstock':
        const overstock = item.actualStock - item.maxStock;
        message = `⚠️ OVERSTOCK SITUATION\n\nProduct: ${item.name} (${item.code})\nCurrent Stock: ${item.actualStock} units\nMaximum Stock: ${item.maxStock} units\nOverstock Amount: ${overstock} units\nTied Capital: ${formatCurrency(overstock * item.unitCost)}\n\nRECOMMENDED ACTIONS:\n1. Run promotional campaign (20-30% discount)\n2. Transfer inventory to other locations\n3. Bundle with complementary products\n4. Contact sales team for bulk deals\n5. Reduce future order quantities\n6. Consider liquidation if aging`;
        break;
      default:
        message = `Action details for ${item.name} (${item.code})`;
    }
    
    alert(message);
  };

  const updateProduct = (id: number, field: keyof Product, value: string | number) => {
    setProducts(products.map(p => p.id === id ? { ...p, [field]: value } : p));
  };

  const exportToExcel = () => {
    const wsData = analysis.map(item => ({
      'Kode Produk': item.code,
      'Nama Produk': item.name,
      'Kategori': item.category,
      'Stok Sistem': item.systemStock,
      'Stok Aktual': item.actualStock,
      'Unit Cost': item.unitCost, // Renamed for consistency
      'Min Stock': item.minStock, // Added for completeness
      'Max Stock': item.maxStock, // Added for completeness
      'Lead Time': item.leadTime, // Added for completeness
      'Avg Demand': item.avgDemand, // Added for completeness
      'Safety Stock': item.safetyStock,
      'Reorder Point': item.reorderPoint,
      'EOQ': item.eoq
    }));

    const summaryData = [
      ['RINGKASAN ANALISIS'],
      ['Total Produk', summary.totalProducts],
      ['Total Nilai Inventory', summary.totalInventoryValue],
      ['Total Nilai Selisih', summary.totalVarianceValue],
      ['Tingkat Akurasi (%)', `${summary.accuracyRate?.toFixed(2)}%`],
      ['Item Low Stock', summary.lowStockItems],
      ['Item Overstock', summary.overstockItems],
      ['Item Perlu Reorder', summary.reorderItems],
      [],
      ['REKOMENDASI TINDAKAN'],
      ...recommendations.map(rec => [rec.type.toUpperCase(), rec.message])
    ];

    const wb = XLSX.utils.book_new();
    const wsAnalysis = XLSX.utils.json_to_sheet(wsData);
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    
    XLSX.utils.book_append_sheet(wb, wsAnalysis, 'Analisis Detail');
    XLSX.utils.book_append_sheet(wb, wsSummary, 'Ringkasan & Rekomendasi');
    
    XLSX.writeFile(wb, `Stock_Opname_Analysis_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const importFromExcel = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const result = e.target?.result;
      if (!result) return;
      const workbook = XLSX.read(result, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data: any[] = XLSX.utils.sheet_to_json(worksheet);
      
      const importedProducts: Product[] = data.map((row, index) => ({
        id: Date.now() + index,
        code: row['Kode Produk'] || `PRD${String(index + 1).padStart(3, '0')}`,
        name: row['Nama Produk'] || '',
        category: row['Kategori'] || '',
        systemStock: Number(row['Stok Sistem']) || 0,
        actualStock: Number(row['Stok Aktual']) || 0,
        unitCost: Number(row['Unit Cost']) || Number(row['Harga Satuan']) || 0, // Support both new and old format
        minStock: Number(row['Min Stock']) || 1,
        maxStock: Number(row['Max Stock']) || 10,
        leadTime: Number(row['Lead Time']) || 1,
        avgDemand: Number(row['Avg Demand']) || 1
      }));
      
      setProducts(importedProducts);
    };
    reader.readAsBinaryString(file);
    event.target.value = '';
  };

  const formatCurrency = (value: number) => {
    return new Intl.NumberFormat('id-ID', {
      style: 'currency',
      currency: 'IDR',
      minimumFractionDigits: 0,
      maximumFractionDigits: 0
    }).format(value);
  };

  const getStatusColor = (status: string) => {
    switch (status) {
      case 'Out of Stock': return 'text-white bg-red-600';
      case 'Low Stock': return 'text-red-600 bg-red-100';
      case 'Overstock': return 'text-orange-600 bg-orange-100';
      case 'Reorder': return 'text-yellow-600 bg-yellow-100';
      default: return 'text-green-600 bg-green-100';
    }
  };


  const getABCColor = (abcClass: string) => {
    switch (abcClass) {
      case 'A': return 'text-red-600 bg-red-100 font-bold';
      case 'B': return 'text-orange-600 bg-orange-100 font-semibold';
      case 'C': return 'text-green-600 bg-green-100';
      default: return 'text-gray-600 bg-gray-100';
    }
  };

  // Filter function for analysis results
  const getFilteredAnalysis = () => {
    return analysis.filter(item => {
      // Search term filter
      const searchMatch = !filters.searchTerm || 
        item.name.toLowerCase().includes(filters.searchTerm.toLowerCase()) ||
        item.code.toLowerCase().includes(filters.searchTerm.toLowerCase()) ||
        item.category.toLowerCase().includes(filters.searchTerm.toLowerCase());
      
      // Stock status filter
      const statusMatch = filters.stockStatus === 'All' || item.stockStatus === filters.stockStatus;
      
      // ABC class filter
      const abcMatch = filters.abcClass === 'All' || item.abcClass === filters.abcClass;
      
      // Category filter
      const categoryMatch = filters.category === 'All' || item.category === filters.category;
      
      // Variance type filter
      const varianceMatch = filters.varianceType === 'All' || 
        (filters.varianceType === 'Positive' && item.variance > 0) ||
        (filters.varianceType === 'Negative' && item.variance < 0) ||
        (filters.varianceType === 'Zero' && item.variance === 0) ||
        (filters.varianceType === 'High Variance' && Math.abs(item.variancePercentage) > 10);
      
      return searchMatch && statusMatch && abcMatch && categoryMatch && varianceMatch;
    });
  };

  // Get unique categories for filter dropdown
  const uniqueCategories = [...new Set(analysis.map(item => item.category))].filter(Boolean).sort();

  return (
    <div className="min-h-screen bg-gray-50 p-4 sm:p-6">
      <div className="max-w-screen-xl mx-auto">
        {/* Header */}
        <header className="bg-white rounded-lg shadow-lg p-4 sm:p-6 mb-6">
          <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="flex items-center space-x-3">
              <Package className="h-10 w-10 text-blue-600" />
              <div>
                <h1 className="text-2xl sm:text-3xl font-bold text-gray-900">Sistem Penunjang Keputusan Stock Opname</h1>
                <p className="text-sm sm:text-base text-gray-600">Analisis mendalam dengan metode ABC, EOQ, Safety Stock dan Reorder Point</p>
              </div>
            </div>
            <div className="flex items-center space-x-2 sm:space-x-3">
              <label className="bg-green-600 text-white px-3 py-2 sm:px-4 sm:py-2 rounded-lg hover:bg-green-700 cursor-pointer flex items-center space-x-2 text-sm sm:text-base">
                <Upload className="h-4 w-4" />
                <span>Import</span>
                <input type="file" accept=".xlsx,.xls" onChange={importFromExcel} className="hidden" />
              </label>
              <button
                onClick={exportToExcel}
                className="bg-blue-600 text-white px-3 py-2 sm:px-4 sm:py-2 rounded-lg hover:bg-blue-700 flex items-center space-x-2 text-sm sm:text-base"
              >
                <Download className="h-4 w-4" />
                <span>Export</span>
              </button>
            </div>
          </div>
        </header>

        {/* Visualizations */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6 mb-6">
          {/* ABC Classification */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">ABC Classification</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-red-50 rounded">
                <span className="font-medium">Class A (High Value)</span>
                <span className="bg-red-100 text-red-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.abcClass === 'A').length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-yellow-50 rounded">
                <span className="font-medium">Class B (Medium Value)</span>
                <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.abcClass === 'B').length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">Class C (Low Value)</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.abcClass === 'C').length} items
                </span>
              </div>
            </div>
          </div>
          
          {/* Stock Status Distribution */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Stock Status Distribution</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">Normal</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.stockStatus === 'Normal').length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-yellow-50 rounded">
                <span className="font-medium">Low Stock</span>
                <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.stockStatus === 'Low Stock').length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-red-50 rounded">
                <span className="font-medium">Overstock</span>
                <span className="bg-red-100 text-red-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.stockStatus === 'Overstock').length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-orange-50 rounded">
                <span className="font-medium">Reorder</span>
                <span className="bg-orange-100 text-orange-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.stockStatus === 'Reorder').length} items
                </span>
              </div>
            </div>
          </div>

          {/* Top Products by Value */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Top Products by Inventory Value</h2>
            <div className="space-y-2">
              {analysis.sort((a, b) => b.inventoryValue - a.inventoryValue).slice(0, 5).map((item, index) => (
                <div key={item.id} className="flex justify-between items-center p-2 bg-gray-50 rounded">
                  <div>
                    <span className="font-medium text-sm">{item.code}</span>
                    <p className="text-xs text-gray-600 truncate">{item.name}</p>
                  </div>
                  <span className="text-sm font-medium text-blue-600">
                    {formatCurrency(item.inventoryValue)}
                  </span>
                </div>
              ))}
            </div>
          </div>

        </section>

        {/* Additional Charts Row */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6 mb-6">
          {/* Variance Analysis */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Variance Analysis</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">Positive Variance</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.variance > 0).length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-red-50 rounded">
                <span className="font-medium">Negative Variance</span>
                <span className="bg-red-100 text-red-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.variance < 0).length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-gray-50 rounded">
                <span className="font-medium">No Variance</span>
                <span className="bg-gray-100 text-gray-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.variance === 0).length} items
                </span>
              </div>
            </div>
          </div>

          {/* Turnover Ratio Analysis */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Inventory Turnover</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">High (&gt;6)</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.turnoverRatio > 6).length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-yellow-50 rounded">
                <span className="font-medium">Medium (3-6)</span>
                <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.turnoverRatio >= 3 && item.turnoverRatio <= 6).length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-red-50 rounded">
                <span className="font-medium">Low (&lt;3)</span>
                <span className="bg-red-100 text-red-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.turnoverRatio < 3).length} items
                </span>
              </div>
            </div>
          </div>

          {/* EOQ vs Current Stock Comparison */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Stock vs EOQ Recommendations</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-blue-50 rounded">
                <span className="font-medium">Stock Below EOQ</span>
                <span className="bg-blue-100 text-blue-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.actualStock < item.eoq).length} items
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">Stock Above EOQ</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.actualStock >= item.eoq).length} items
                </span>
              </div>
              <div className="text-xs text-gray-600 mt-2">
                <p>EOQ = Economic Order Quantity</p>
                <p>Optimal order size for cost efficiency</p>
              </div>
            </div>
          </div>

        </section>


        {/* Summary Dashboard */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 mb-6">
          {/* Card: Total Produk */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg flex items-center justify-between">
            <div>
              <p className="text-gray-600 text-sm">Total Produk</p>
              <p className="text-2xl font-bold text-gray-900">{summary.totalProducts || 0}</p>
            </div>
            <Package className="h-8 w-8 text-blue-600" />
          </div>
          {/* Card: Nilai Inventory */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg flex items-center justify-between">
            <div>
              <p className="text-gray-600 text-sm">Nilai Inventory</p>
              <p className="text-2xl font-bold text-gray-900">{formatCurrency(summary.totalInventoryValue || 0)}</p>
            </div>
            <TrendingUp className="h-8 w-8 text-green-600" />
          </div>
          {/* Card: Akurasi Stock */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg flex items-center justify-between">
            <div>
              <p className="text-gray-600 text-sm">Akurasi Stock</p>
              <p className="text-2xl font-bold text-gray-900">{(summary.accuracyRate || 0).toFixed(1)}%</p>
            </div>
            <CheckCircle className="h-8 w-8 text-green-600" />
          </div>
          {/* Card: Item Perlu Perhatian */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg flex items-center justify-between">
            <div>
              <p className="text-gray-600 text-sm">Item Perlu Perhatian</p>
              <p className="text-2xl font-bold text-red-600">{(summary.lowStockItems || 0) + (summary.overstockItems || 0)}</p>
            </div>
            <AlertTriangle className="h-8 w-8 text-red-600" />
          </div>
        </section>

        {/* Recommendations Panel */}
        {recommendations.length > 0 && (
          <section className="bg-white rounded-lg shadow-lg p-4 sm:p-6 mb-6">
            <h2 className="text-xl font-bold text-gray-900 mb-4 flex items-center">
              <BarChart3 className="h-5 w-5 mr-2" />
              Rekomendasi Sistem
            </h2>
            <div className="space-y-3">
              {recommendations.slice(0, 5).map((rec, index) => (
                <div key={index} className={`p-3 rounded-lg border-l-4 ${
                  rec.type === 'urgent' ? 'border-red-500 bg-red-50' :
                  rec.type === 'warning' ? 'border-yellow-500 bg-yellow-50' :
                  rec.type === 'audit' ? 'border-blue-500 bg-blue-50' :
                  'border-green-500 bg-green-50'
                }`}>
                  <p className="text-sm font-medium text-gray-900">{rec.message}</p>
                </div>
              ))}
            </div>
          </section>
        )}

        {/* Urgency Ranking - Priority Action Dashboard */}
        {urgencyRankings.length > 0 && (
          <section className="bg-white rounded-lg shadow-lg p-4 sm:p-6 mb-6">
            <h2 className="text-xl font-bold text-gray-900 mb-4 flex items-center">
              <AlertTriangle className="h-5 w-5 mr-2 text-red-600" />
              🚨 Urgency Ranking - SPK Priority Actions
            </h2>
            <p className="text-sm text-gray-600 mb-6">
              Items ranked by urgency score based on stock status, ABC classification, and business impact
            </p>
            
            <div className="space-y-4">
              {urgencyRankings.map((item, index) => (
                <div key={item.id} className={`p-4 rounded-lg border-l-4 ${
                  item.urgencyLevel === 'CRITICAL' ? 'border-red-600 bg-red-50' :
                  item.urgencyLevel === 'HIGH' ? 'border-orange-500 bg-orange-50' :
                  item.urgencyLevel === 'MEDIUM' ? 'border-yellow-500 bg-yellow-50' :
                  'border-blue-500 bg-blue-50'
                }`}>
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div className="flex-1">
                      <div className="flex items-center gap-3 mb-2">
                        <span className={`text-2xl font-bold ${
                          item.urgencyLevel === 'CRITICAL' ? 'text-red-600' :
                          item.urgencyLevel === 'HIGH' ? 'text-orange-600' :
                          item.urgencyLevel === 'MEDIUM' ? 'text-yellow-600' :
                          'text-blue-600'
                        }`}>
                          #{index + 1}
                        </span>
                        <div>
                          <h3 className="font-bold text-gray-900">{item.name}</h3>
                          <p className="text-sm text-gray-600">{item.code} • {item.category}</p>
                        </div>
                        <div className="flex gap-2">
                          <span className={`px-2 py-1 text-xs rounded-full font-medium ${
                            item.urgencyLevel === 'CRITICAL' ? 'bg-red-100 text-red-800' :
                            item.urgencyLevel === 'HIGH' ? 'bg-orange-100 text-orange-800' :
                            item.urgencyLevel === 'MEDIUM' ? 'bg-yellow-100 text-yellow-800' :
                            'bg-blue-100 text-blue-800'
                          }`}>
                            {item.urgencyLevel}
                          </span>
                          <span className={`px-2 py-1 text-xs rounded-full ${getABCColor(item.abcClass)}`}>
                            Class {item.abcClass}
                          </span>
                          <span className={`px-2 py-1 text-xs rounded-full ${getStatusColor(item.stockStatus)}`}>
                            {item.stockStatus}
                          </span>
                        </div>
                      </div>
                      
                      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-3">
                        <div>
                          <p className="text-xs text-gray-500">Urgency Score</p>
                          <div className="flex items-center gap-2">
                            <div className={`w-16 h-3 bg-gray-200 rounded-full overflow-hidden`}>
                              <div 
                                className={`h-full rounded-full ${
                                  item.urgencyLevel === 'CRITICAL' ? 'bg-red-600' :
                                  item.urgencyLevel === 'HIGH' ? 'bg-orange-500' :
                                  item.urgencyLevel === 'MEDIUM' ? 'bg-yellow-500' :
                                  'bg-blue-500'
                                }`}
                                style={{ width: `${(item.urgencyScore / 12) * 100}%` }}
                              ></div>
                            </div>
                            <span className="font-bold text-sm">{Math.round((item.urgencyScore / 12) * 100)}%</span>
                          </div>
                          <p className="text-xs text-gray-600 mt-1">{item.urgencyScore} out of 12</p>
                        </div>
                        <div>
                          <p className="text-xs text-gray-500">Business Impact</p>
                          <p className="font-medium">{formatCurrency(item.businessImpact)}</p>
                        </div>
                        <div>
                          <p className="text-xs text-gray-500">Timeframe</p>
                          <p className="font-medium text-red-600">{item.timeframe}</p>
                        </div>
                        {item.daysUntilStockout && (
                          <div>
                            <p className="text-xs text-gray-500">Days Until Stockout</p>
                            <p className="font-bold text-red-600">~{item.daysUntilStockout} days</p>
                          </div>
                        )}
                      </div>
                      
                      <div className="mb-3">
                        <p className="text-sm text-gray-700 mb-1"><strong>Reason:</strong> {item.reason}</p>
                        <p className="text-sm text-blue-700"><strong>Recommended Action:</strong> {item.recommendedAction}</p>
                      </div>
                    </div>
                    
                    <div className="flex gap-2">
                      {item.stockStatus === 'Out of Stock' && (
                        <button
                          onClick={() => showActionAlert(analysis.find(a => a.id === item.id)!)}
                          className="bg-red-600 text-white px-4 py-2 rounded-lg hover:bg-red-700 transition-colors font-medium flex items-center gap-2"
                        >
                          🚨 EMERGENCY ORDER
                        </button>
                      )}
                      {item.stockStatus === 'Reorder' && (
                        <button
                          onClick={() => showActionAlert(analysis.find(a => a.id === item.id)!)}
                          className="bg-orange-500 text-white px-4 py-2 rounded-lg hover:bg-orange-600 transition-colors font-medium flex items-center gap-2"
                        >
                          🔄 REORDER NOW
                        </button>
                      )}
                      {item.stockStatus === 'Low Stock' && (
                        <button
                          onClick={() => showActionAlert(analysis.find(a => a.id === item.id)!)}
                          className="bg-yellow-500 text-white px-4 py-2 rounded-lg hover:bg-yellow-600 transition-colors font-medium flex items-center gap-2"
                        >
                          👁️ MONITOR
                        </button>
                      )}
                      {item.stockStatus === 'Overstock' && (
                        <button
                          onClick={() => showActionAlert(analysis.find(a => a.id === item.id)!)}
                          className="bg-red-500 text-white px-4 py-2 rounded-lg hover:bg-red-600 transition-colors font-medium flex items-center gap-2"
                        >
                          📉 REDUCE STOCK
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
            
            <div className="mt-6 p-4 bg-gray-50 rounded-lg">
              <h3 className="font-bold text-gray-900 mb-2">📊 How to Read Urgency Score:</h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-sm">
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-4 h-2 bg-red-600 rounded"></div>
                    <span className="font-medium text-red-600">83-100% (CRITICAL)</span>
                  </div>
                  <p className="text-gray-600">🚨 Emergency action needed NOW!</p>
                  <p className="text-xs text-gray-500">Out of stock or high-value items</p>
                </div>
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-4 h-2 bg-orange-500 rounded"></div>
                    <span className="font-medium text-orange-600">58-82% (HIGH)</span>
                  </div>
                  <p className="text-gray-600">⚡ Action needed within days</p>
                  <p className="text-xs text-gray-500">Reorder level or low stock</p>
                </div>
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-4 h-2 bg-yellow-500 rounded"></div>
                    <span className="font-medium text-yellow-600">33-57% (MEDIUM)</span>
                  </div>
                  <p className="text-gray-600">⚠️ Plan action within weeks</p>
                  <p className="text-xs text-gray-500">Overstock or approaching limits</p>
                </div>
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-4 h-2 bg-blue-500 rounded"></div>
                    <span className="font-medium text-blue-600">8-32% (LOW)</span>
                  </div>
                  <p className="text-gray-600">👁️ Monitor situation</p>
                  <p className="text-xs text-gray-500">Non-urgent monitoring</p>
                </div>
              </div>
              <div className="mt-4 p-3 bg-blue-50 rounded border-l-4 border-blue-400">
                <p className="text-sm text-blue-800">
                  <strong>💡 Tip:</strong> Higher percentages mean more urgent action needed. The bar color matches the urgency level.
                </p>
              </div>
            </div>
          </section>
        )}



        {/* Analysis Results */}
        <section className="bg-white rounded-lg shadow-lg p-4 sm:p-6">
          <h2 className="text-xl font-bold text-gray-900 mb-4 flex items-center">
            <Calculator className="h-5 w-5 mr-2" />
            Hasil Analisis SPK
          </h2>
          
          {/* Filter Controls */}
          <div className="bg-gray-50 p-4 rounded-lg mb-6">
            <h3 className="font-semibold text-gray-900 mb-3">🔍 Filter & Search Options</h3>
            
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-4">
              {/* Search Input */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Search Product</label>
                <input
                  type="text"
                  placeholder="Name, code, or category"
                  value={filters.searchTerm}
                  onChange={(e) => setFilters({...filters, searchTerm: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                />
              </div>
              
              {/* Stock Status Filter */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Stock Status</label>
                <select
                  value={filters.stockStatus}
                  onChange={(e) => setFilters({...filters, stockStatus: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="All">All Status</option>
                  <option value="Normal">Normal</option>
                  <option value="Low Stock">Low Stock</option>
                  <option value="Overstock">Overstock</option>
                  <option value="Reorder">Reorder</option>
                  <option value="Out of Stock">Out of Stock</option>
                </select>
              </div>
              
              {/* ABC Class Filter */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">ABC Class</label>
                <select
                  value={filters.abcClass}
                  onChange={(e) => setFilters({...filters, abcClass: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="All">All Classes</option>
                  <option value="A">Class A (High Value)</option>
                  <option value="B">Class B (Medium Value)</option>
                  <option value="C">Class C (Low Value)</option>
                </select>
              </div>
              
              {/* Category Filter */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Category</label>
                <select
                  value={filters.category}
                  onChange={(e) => setFilters({...filters, category: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="All">All Categories</option>
                  {uniqueCategories.map(category => (
                    <option key={category} value={category}>{category}</option>
                  ))}
                </select>
              </div>
              
              {/* Variance Type Filter */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Variance Type</label>
                <select
                  value={filters.varianceType}
                  onChange={(e) => setFilters({...filters, varianceType: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="All">All Variance</option>
                  <option value="Positive">Positive (+)</option>
                  <option value="Negative">Negative (-)</option>
                  <option value="Zero">Zero (0)</option>
                  <option value="High Variance">High Variance (&gt;10%)</option>
                </select>
              </div>
              
              {/* Clear Filters Button */}
              <div className="flex items-end">
                <button
                  onClick={() => setFilters({
                    searchTerm: '',
                    stockStatus: 'All',
                    abcClass: 'All',
                    category: 'All',
                    varianceType: 'All'
                  })}
                  className="w-full bg-gray-600 text-white px-3 py-2 rounded-md text-sm hover:bg-gray-700 transition-colors"
                >
                  Clear Filters
                </button>
              </div>
            </div>
            
            {/* Filter Results Summary */}
            <div className="mt-3 flex items-center justify-between">
              <p className="text-sm text-gray-600">
                Showing <span className="font-medium">{getFilteredAnalysis().length}</span> of <span className="font-medium">{analysis.length}</span> items
                {filters.searchTerm && (
                  <span className="ml-1">• Search: "{filters.searchTerm}"</span>
                )}
                {filters.stockStatus !== 'All' && (
                  <span className="ml-1">• Status: {filters.stockStatus}</span>
                )}
                {filters.abcClass !== 'All' && (
                  <span className="ml-1">• Class: {filters.abcClass}</span>
                )}
                {filters.category !== 'All' && (
                  <span className="ml-1">• Category: {filters.category}</span>
                )}
                {filters.varianceType !== 'All' && (
                  <span className="ml-1">• Variance: {filters.varianceType}</span>
                )}
              </p>
            </div>
          </div>
          
          <div className="overflow-x-auto">
            <table className="min-w-full table-auto text-sm">
              <thead className="bg-gray-50">
                <tr>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Produk</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Selisih</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Selisih %</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Nilai Selisih</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">ABC</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Status</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Safety Stock</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Reorder Point</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">EOQ</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Turnover</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Action</th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {getFilteredAnalysis().map(item => (
                  <tr key={item.id}>
                    <td className="px-4 py-2 whitespace-nowrap">
                      <div className="font-medium text-gray-900">{item.name}</div>
                      <div className="text-sm text-gray-500">{item.code}</div>
                    </td>
                    <td className="px-4 py-2"><span className={`font-medium ${item.variance >= 0 ? 'text-green-600' : 'text-red-600'}`}>{item.variance >= 0 ? '+' : ''}{item.variance}</span></td>
                    <td className="px-4 py-2"><span className={`font-medium ${Math.abs(item.variancePercentage) > 10 ? 'text-red-600' : 'text-gray-900'}`}>{item.variancePercentage.toFixed(1)}%</span></td>
                    <td className="px-4 py-2"><span className={`text-sm ${item.varianceValue >= 0 ? 'text-green-600' : 'text-red-600'}`}>{formatCurrency(item.varianceValue)}</span></td>
                    <td className="px-4 py-2"><span className={`px-2 py-1 text-xs rounded-full ${getABCColor(item.abcClass)}`}>{item.abcClass}</span></td>
                    <td className="px-4 py-2"><span className={`px-2 py-1 text-xs rounded-full ${getStatusColor(item.stockStatus)}`}>{item.stockStatus}</span></td>
                    <td className="px-4 py-2 text-sm">{item.safetyStock}</td>
                    <td className="px-4 py-2 text-sm">{item.reorderPoint}</td>
                    <td className="px-4 py-2 text-sm font-medium">{item.eoq}</td>
                    <td className="px-4 py-2 text-sm">{item.turnoverRatio.toFixed(1)}x</td>
                    <td className="px-4 py-2">
                      {item.stockStatus === 'Reorder' && (
                        <button
                          onClick={() => {
                            const message = `🔄 REORDER ACTION REQUIRED\n\nProduct: ${item.name} (${item.code})\nCurrent Stock: ${item.actualStock}\nReorder Point: ${item.reorderPoint}\nRecommended Order Quantity (EOQ): ${item.eoq} units\n\nAction: Place order for ${item.eoq} units immediately!`;
                            alert(message);
                          }}
                          className="bg-orange-500 text-white px-3 py-1 rounded text-xs hover:bg-orange-600 transition-colors"
                        >
                          Reorder Now
                        </button>
                      )}
                      {item.stockStatus === 'Overstock' && (
                        <button
                          onClick={() => {
                            const overstock = item.actualStock - item.maxStock;
                            const message = `⚠️ OVERSTOCK ACTION REQUIRED\n\nProduct: ${item.name} (${item.code})\nCurrent Stock: ${item.actualStock}\nMax Stock: ${item.maxStock}\nOverstock Amount: ${overstock} units\n\nRecommended Actions:\n• Run promotional campaign\n• Transfer to other locations\n• Reduce future orders\n• Consider bulk discounts`;
                            alert(message);
                          }}
                          className="bg-red-500 text-white px-3 py-1 rounded text-xs hover:bg-red-600 transition-colors"
                        >
                          Reduce Stock
                        </button>
                      )}
                      {item.stockStatus === 'Low Stock' && (
                        <button
                          onClick={() => {
                            const shortfall = item.minStock - item.actualStock;
                            const message = `⚡ LOW STOCK WARNING\n\nProduct: ${item.name} (${item.code})\nCurrent Stock: ${item.actualStock}\nMinimum Stock: ${item.minStock}\nShortfall: ${shortfall} units\n\nAction: Consider ordering ${item.eoq} units to maintain optimal levels.`;
                            alert(message);
                          }}
                          className="bg-yellow-500 text-white px-3 py-1 rounded text-xs hover:bg-yellow-600 transition-colors"
                        >
                          Monitor
                        </button>
                      )}
                      {item.stockStatus === 'Normal' && (
                        <span className="text-green-600 text-xs font-medium">✓ OK</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>

        {/* Product Management - Moved to bottom */}
        <section className="bg-white rounded-lg shadow-lg p-4 sm:p-6 mt-6">
          <div className="flex flex-col sm:flex-row items-center justify-between mb-4 gap-3">
            <h2 className="text-xl font-bold text-gray-900">Data Produk</h2>
            <button
              onClick={addProduct}
              className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 flex items-center space-x-2 self-start sm:self-center"
            >
              <Plus className="h-4 w-4" />
              <span>Tambah Produk</span>
            </button>
          </div>
          
          <div className="overflow-x-auto">
            <table className="min-w-full table-auto text-sm">
              <thead className="bg-gray-50">
                <tr>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Kode</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Nama Produk</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Kategori</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Stok Sistem</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Stok Aktual</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Harga</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Min/Max</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Lead Time</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Avg Demand</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Aksi</th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {products.map(product => (
                  <tr key={product.id}>
                    <td className="px-4 py-2"><input type="text" value={product.code} onChange={(e) => updateProduct(product.id, 'code', e.target.value)} className="w-24 p-1 border rounded text-sm" /></td>
                    <td className="px-4 py-2"><input type="text" value={product.name} onChange={(e) => updateProduct(product.id, 'name', e.target.value)} className="w-40 p-1 border rounded text-sm" /></td>
                    <td className="px-4 py-2"><input type="text" value={product.category} onChange={(e) => updateProduct(product.id, 'category', e.target.value)} className="w-28 p-1 border rounded text-sm" /></td>
                    <td className="px-4 py-2"><input type="number" value={product.systemStock} onChange={(e) => updateProduct(product.id, 'systemStock', parseInt(e.target.value) || 0)} className="w-20 p-1 border rounded text-sm" /></td>
                    <td className="px-4 py-2"><input type="number" value={product.actualStock} onChange={(e) => updateProduct(product.id, 'actualStock', parseInt(e.target.value) || 0)} className="w-20 p-1 border rounded text-sm" /></td>
                    <td className="px-4 py-2"><input type="number" value={product.unitCost} onChange={(e) => updateProduct(product.id, 'unitCost', parseInt(e.target.value) || 0)} className="w-28 p-1 border rounded text-sm" /></td>
                    <td className="px-4 py-2">
                      <div className="flex space-x-1">
                        <input type="number" value={product.minStock} onChange={(e) => updateProduct(product.id, 'minStock', parseInt(e.target.value) || 0)} className="w-16 p-1 border rounded text-sm" placeholder="Min" />
                        <input type="number" value={product.maxStock} onChange={(e) => updateProduct(product.id, 'maxStock', parseInt(e.target.value) || 0)} className="w-16 p-1 border rounded text-sm" placeholder="Max" />
                      </div>
                    </td>
                    <td className="px-4 py-2"><input type="number" value={product.leadTime} onChange={(e) => updateProduct(product.id, 'leadTime', parseInt(e.target.value) || 1)} className="w-20 p-1 border rounded text-sm" /></td>
                    <td className="px-4 py-2"><input type="number" value={product.avgDemand} onChange={(e) => updateProduct(product.id, 'avgDemand', parseInt(e.target.value) || 1)} className="w-20 p-1 border rounded text-sm" /></td>
                    <td className="px-4 py-2"><button onClick={() => deleteProduct(product.id)} className="text-red-600 hover:text-red-800"><Trash2 className="h-4 w-4" /></button></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </div>
  );
};

export default StockOpnameDSS;
