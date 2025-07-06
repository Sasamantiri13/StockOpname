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

// Definisikan tipe data untuk urgency ranking dengan AHP
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
  // AHP specific fields
  ahpCriteria: {
    stockLevel: number;
    financialImpact: number;
    demandCriticality: number;
    leadTimeRisk: number;
  };
  ahpWeights: {
    stockLevel: number;
    financialImpact: number;
    demandCriticality: number;
    leadTimeRisk: number;
  };
  ahpCompositeScore: number;
}

// AHP: Tentukan Bobot Kriteria Strategis (Step 1)
// Strategic criteria weights based on Saaty Scale analysis
const AHP_STRATEGIC_WEIGHTS = {
  // Kriteria Strategis Utama
  stockCriticality: 0.50,   // Kritisitas stok (emergency level)
  businessValue: 0.30,      // Nilai bisnis dan dampak finansial
  operationalRisk: 0.20     // Risiko operasional (lead time, demand)
};

// Sub-kriteria untuk setiap kriteria strategis
const AHP_SUB_CRITERIA = {
  stockCriticality: {
    stockoutRisk: 0.60,      // Risiko kehabisan stok
    stockLevel: 0.40         // Level stok saat ini
  },
  businessValue: {
    inventoryValue: 0.70,    // Nilai inventory
    abcClassification: 0.30   // Klasifikasi ABC
  },
  operationalRisk: {
    demandVariability: 0.50, // Variabilitas permintaan
    leadTimeRisk: 0.50       // Risiko lead time
  }
};

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

    // Calculate urgency ranking using AHP (Analytic Hierarchy Process)
    // Following the schema: AHP → Prioritisasi Item → Analisis ABC berbasis Bobot AHP
    const urgencyRankings: UrgencyItem[] = finalAnalysis.filter(item => item.stockStatus !== 'Normal').map(item => {
      
      // Step 1: AHP Kriteria Strategis Calculation (0-100 scale)
      
      // A. Stock Criticality (50% weight)
      // A1. Stockout Risk (60% of Stock Criticality)
      let stockoutRiskScore = 0;
      if (item.stockStatus === 'Out of Stock') {
        stockoutRiskScore = 100; // Critical - no stock available
      } else if (item.stockStatus === 'Reorder') {
        stockoutRiskScore = 85; // High risk - at reorder point
      } else if (item.stockStatus === 'Low Stock') {
        stockoutRiskScore = 70; // Medium-high risk - below minimum
      } else if (item.stockStatus === 'Overstock') {
        stockoutRiskScore = 20; // Low risk - excess stock
      }
      
      // A2. Stock Level (40% of Stock Criticality)
      let stockLevelScore = 0;
      if (item.actualStock <= 0) {
        stockLevelScore = 100;
      } else if (item.actualStock <= item.minStock) {
        stockLevelScore = 80;
      } else if (item.actualStock >= item.maxStock) {
        stockLevelScore = 30;
      } else {
        // Normal range - calculate relative position
        const range = item.maxStock - item.minStock;
        const position = (item.actualStock - item.minStock) / range;
        stockLevelScore = 50 + (position * 30); // 50-80 range
      }
      
      // Stock Criticality Composite Score
      const stockCriticalityScore = (
        stockoutRiskScore * AHP_SUB_CRITERIA.stockCriticality.stockoutRisk +
        stockLevelScore * AHP_SUB_CRITERIA.stockCriticality.stockLevel
      );
      
      // B. Business Value (30% weight)
      // B1. Inventory Value (70% of Business Value)
      const maxInventoryValue = Math.max(...finalAnalysis.map(i => i.inventoryValue));
      const inventoryValueScore = maxInventoryValue > 0 ? (item.inventoryValue / maxInventoryValue) * 100 : 0;
      
      // B2. ABC Classification (30% of Business Value)
      let abcScore = 0;
      if (item.abcClass === 'A') {
        abcScore = 100; // Highest priority
      } else if (item.abcClass === 'B') {
        abcScore = 60;  // Medium priority
      } else {
        abcScore = 30;  // Lower priority
      }
      
      // Business Value Composite Score
      const businessValueScore = (
        inventoryValueScore * AHP_SUB_CRITERIA.businessValue.inventoryValue +
        abcScore * AHP_SUB_CRITERIA.businessValue.abcClassification
      );
      
      // C. Operational Risk (20% weight)
      // C1. Demand Variability (50% of Operational Risk)
      let demandVariabilityScore = 0;
      if (item.turnoverRatio > 6) {
        demandVariabilityScore = 80; // High variability - fast moving
      } else if (item.turnoverRatio >= 3) {
        demandVariabilityScore = 50; // Medium variability
      } else {
        demandVariabilityScore = 30; // Low variability - slow moving
      }
      
      // C2. Lead Time Risk (50% of Operational Risk)
      const maxLeadTime = Math.max(...finalAnalysis.map(i => i.leadTime));
      const leadTimeRiskScore = maxLeadTime > 0 ? (item.leadTime / maxLeadTime) * 100 : 0;
      
      // Operational Risk Composite Score
      const operationalRiskScore = (
        demandVariabilityScore * AHP_SUB_CRITERIA.operationalRisk.demandVariability +
        leadTimeRiskScore * AHP_SUB_CRITERIA.operationalRisk.leadTimeRisk
      );
      
      // Final AHP Composite Score (Strategic Level)
      const ahpCompositeScore = (
        stockCriticalityScore * AHP_STRATEGIC_WEIGHTS.stockCriticality +
        businessValueScore * AHP_STRATEGIC_WEIGHTS.businessValue +
        operationalRiskScore * AHP_STRATEGIC_WEIGHTS.operationalRisk
      );
      
      // Convert AHP score to 12-point urgency score for compatibility
      const urgencyScore = Math.round((ahpCompositeScore / 100) * 12);
      
      // Determine urgency level based on AHP composite score
      let urgencyLevel: 'CRITICAL' | 'HIGH' | 'MEDIUM' | 'LOW' = 'LOW';
      if (ahpCompositeScore >= 83) {
        urgencyLevel = 'CRITICAL';
      } else if (ahpCompositeScore >= 58) {
        urgencyLevel = 'HIGH';
      } else if (ahpCompositeScore >= 33) {
        urgencyLevel = 'MEDIUM';
      }
      
      // Set reason, action, and timeframe based on urgency level and stock status
      let reason = '';
      let recommendedAction = '';
      let timeframe = 'Within 1 month';
      let daysUntilStockout;
      
      switch (item.stockStatus) {
        case 'Out of Stock':
          reason = `Kritis! Stok habis dengan skor prioritas AHP tinggi (${ahpCompositeScore.toFixed(1)})`;
          recommendedAction = `Pesanan darurat: ${item.eoq} unit segera!`;
          timeframe = 'Segera';
          break;
        case 'Low Stock':
          reason = `Peringatan stok rendah dengan skor AHP ${ahpCompositeScore.toFixed(1)} - item kelas ${item.abcClass}`;
          recommendedAction = `Pesan ${item.eoq} unit dalam jangka waktu prioritas`;
          timeframe = urgencyLevel === 'CRITICAL' ? 'Dalam 1 hari' : 'Dalam 1 minggu';
          daysUntilStockout = Math.floor(item.actualStock / item.avgDemand);
          break;
        case 'Reorder':
          reason = `Titik reorder tercapai - analisis AHP menunjukkan prioritas ${urgencyLevel === 'CRITICAL' ? 'kritis' : urgencyLevel === 'HIGH' ? 'tinggi' : urgencyLevel === 'MEDIUM' ? 'sedang' : 'rendah'}`;
          recommendedAction = `Buat pesanan untuk ${item.eoq} unit berdasarkan rekomendasi AHP`;
          timeframe = urgencyLevel === 'CRITICAL' ? 'Dalam 2 hari' : 'Dalam 10 hari';
          break;
        case 'Overstock':
          reason = `Situasi stok berlebih - skor AHP: ${ahpCompositeScore.toFixed(1)} menunjukkan prioritas tindakan`;
          recommendedAction = 'Terapkan strategi permintaan: promosi, redistribusi, atau strategi tahan';
          timeframe = urgencyLevel === 'HIGH' ? 'Dalam 2 minggu' : 'Dalam 1 bulan';
          break;
      }
      
      // Business impact with AHP weighting
      const businessImpact = item.inventoryValue * (1 + (ahpCompositeScore / 100));
      
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
        timeframe,
        // AHP specific fields
        ahpCriteria: {
          stockLevel: stockLevelScore,
          financialImpact: inventoryValueScore,
          demandCriticality: demandVariabilityScore,
          leadTimeRisk: leadTimeRiskScore
        },
        ahpWeights: {
          stockLevel: AHP_STRATEGIC_WEIGHTS.stockCriticality,
          financialImpact: AHP_STRATEGIC_WEIGHTS.businessValue,
          demandCriticality: AHP_SUB_CRITERIA.operationalRisk.demandVariability,
          leadTimeRisk: AHP_SUB_CRITERIA.operationalRisk.leadTimeRisk
        },
        ahpCompositeScore
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
      alert('Gagal menambahkan produk. Silakan coba lagi.');
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
        message = `🚨 PESANAN DARURAT DIPERLUKAN\n\nProduk: ${item.name} (${item.code})\nStatus: HABIS STOK - Tidak ada inventori tersedia!\nKuantitas Pesanan yang Direkomendasikan (EOQ): ${item.eoq} unit\nEstimasi Lead Time: ${item.leadTime} hari\n\nTINDAKAN SEGERA:\n1. Hubungi pemasok SEKARANG\n2. Buat pesanan darurat untuk ${item.eoq} unit\n3. Periksa apakah pelanggan dapat menunggu\n4. Pertimbangkan produk pengganti\n\nIni adalah situasi KRITIS yang mempengaruhi penjualan!`;
        break;
      case 'Reorder':
        message = `🔄 TINDAKAN REORDER DIPERLUKAN\n\nProduk: ${item.name} (${item.code})\nStok Saat Ini: ${item.actualStock} unit\nTitik Reorder: ${item.reorderPoint} unit\nKuantitas Pesanan yang Direkomendasikan (EOQ): ${item.eoq} unit\nLead Time: ${item.leadTime} hari\n\nTINDAKAN YANG DIPERLUKAN:\n1. Buat pesanan untuk ${item.eoq} unit dalam 10 hari\n2. Pantau konsumsi harian\n3. Hubungi pemasok untuk konfirmasi pengiriman\n4. Perbarui perkiraan inventori`;
        break;
      case 'Low Stock':
        message = `⚡ PERINGATAN STOK RENDAH\n\nProduk: ${item.name} (${item.code})\nStok Saat Ini: ${item.actualStock} unit\nStok Minimum: ${item.minStock} unit\nKekurangan: ${item.minStock - item.actualStock} unit\nEstimasi Hari Hingga Habis: ~${Math.floor(item.actualStock / item.avgDemand)} hari\n\nTINDAKAN YANG DIREKOMENDASIKAN:\n1. Pantau konsumsi dengan cermat\n2. Pertimbangkan untuk memesan ${item.eoq} unit segera\n3. Atur peringatan otomatis\n4. Tinjau pola permintaan\n5. Periksa ketersediaan pemasok`;
        break;
      case 'Overstock':
        const overstock = item.actualStock - item.maxStock;
        message = `⚠️ SITUASI STOK BERLEBIH\n\nProduk: ${item.name} (${item.code})\nStok Saat Ini: ${item.actualStock} unit\nStok Maksimum: ${item.maxStock} unit\nJumlah Stok Berlebih: ${overstock} unit\nModal Tertahan: ${formatCurrency(overstock * item.unitCost)}\n\nTINDAKAN YANG DIREKOMENDASIKAN:\n1. Jalankan kampanye promosi (diskon 20-30%)\n2. Transfer inventori ke lokasi lain\n3. Bundling dengan produk pelengkap\n4. Hubungi tim penjualan untuk penawaran grosir\n5. Kurangi kuantitas pesanan masa depan\n6. Pertimbangkan likuidasi jika produk menua`;
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
                <h1 className="text-2xl sm:text-3xl font-bold text-gray-900">Sistem Pendukung Keputusan Stok Opname</h1>
                <p className="text-sm sm:text-base text-gray-600">Sistem Terintegrasi Pendukung Keputusan Inventori Berbasis AHP, Dengan Penerapan ABC, EOQ, Safety Stock, dan Reorder Point</p>
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
            <h2 className="text-lg font-bold text-gray-900 mb-4">Klasifikasi ABC</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-red-50 rounded">
                <span className="font-medium">Kelas A (Nilai Tinggi)</span>
                <span className="bg-red-100 text-red-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.abcClass === 'A').length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-yellow-50 rounded">
                <span className="font-medium">Kelas B (Nilai Sedang)</span>
                <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.abcClass === 'B').length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">Kelas C (Nilai Rendah)</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.abcClass === 'C').length} item
                </span>
              </div>
            </div>
          </div>
          
          {/* Status Stok Distribution */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Distribusi Status Stok</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">Normal</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.stockStatus === 'Normal').length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-yellow-50 rounded">
                <span className="font-medium">Stok Rendah</span>
                <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.stockStatus === 'Low Stock').length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-red-50 rounded">
                <span className="font-medium">Stok Berlebih</span>
                <span className="bg-red-100 text-red-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.stockStatus === 'Overstock').length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-orange-50 rounded">
                <span className="font-medium">Perlu Reorder</span>
                <span className="bg-orange-100 text-orange-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.stockStatus === 'Reorder').length} item
                </span>
              </div>
            </div>
          </div>

          {/* Produk Teratas berdasarkan Nilai */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Produk Teratas berdasarkan Nilai Inventori</h2>
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
          {/* Analisis Selisih */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Analisis Selisih</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">Selisih Positif</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.variance > 0).length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-red-50 rounded">
                <span className="font-medium">Selisih Negatif</span>
                <span className="bg-red-100 text-red-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.variance < 0).length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-gray-50 rounded">
                <span className="font-medium">Tanpa Selisih</span>
                <span className="bg-gray-100 text-gray-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.variance === 0).length} item
                </span>
              </div>
            </div>
          </div>

          {/* Analisis Rasio Perputaran */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Perputaran Inventori</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">Tinggi (&gt;6)</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.turnoverRatio > 6).length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-yellow-50 rounded">
                <span className="font-medium">Sedang (3-6)</span>
                <span className="bg-yellow-100 text-yellow-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.turnoverRatio >= 3 && item.turnoverRatio <= 6).length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-red-50 rounded">
                <span className="font-medium">Rendah (&lt;3)</span>
                <span className="bg-red-100 text-red-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.turnoverRatio < 3).length} item
                </span>
              </div>
            </div>
          </div>

          {/* Perbandingan Stok vs EOQ */}
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg">
            <h2 className="text-lg font-bold text-gray-900 mb-4">Stok vs Rekomendasi EOQ</h2>
            <div className="space-y-2">
              <div className="flex justify-between items-center p-2 bg-blue-50 rounded">
                <span className="font-medium">Stok di Bawah EOQ</span>
                <span className="bg-blue-100 text-blue-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.actualStock < item.eoq).length} item
                </span>
              </div>
              <div className="flex justify-between items-center p-2 bg-green-50 rounded">
                <span className="font-medium">Stok di Atas EOQ</span>
                <span className="bg-green-100 text-green-800 px-2 py-1 rounded text-sm">
                  {analysis.filter(item => item.actualStock >= item.eoq).length} item
                </span>
              </div>
              <div className="text-xs text-gray-600 mt-2">
                <p>EOQ = Kuantitas Pesanan Ekonomis</p>
                <p>Ukuran pesanan optimal untuk efisiensi biaya</p>
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
              🚨 Peringkat Urgensi - Tindakan Prioritas SPK - Berdasarkan Metode AHP
            </h2>
            <p className="text-sm text-gray-600 mb-6">
              Item diurutkan berdasarkan skor urgensi menggunakan analisis multi-kriteria AHP (Analytic Hierarchy Process):
              Tingkat Stok (45%), Dampak Finansial (30%), Kritisitas Permintaan (15%), Risiko Lead Time (10%)
            </p>
            
            {/* Cara Membaca Skor Urgensi - Moved to top */}
            <div className="mb-6 p-4 bg-gray-50 rounded-lg">
              <h3 className="font-bold text-gray-900 mb-2">📊 Cara Membaca Skor Urgensi:</h3>
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-sm">
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-4 h-2 bg-red-600 rounded"></div>
                    <span className="font-medium text-red-600">83-100% (KRITIS)</span>
                  </div>
                  <p className="text-gray-600">🚨 Tindakan darurat diperlukan SEKARANG!</p>
                  <p className="text-xs text-gray-500">Stok habis atau item bernilai tinggi</p>
                </div>
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-4 h-2 bg-orange-500 rounded"></div>
                    <span className="font-medium text-orange-600">58-82% (TINGGI)</span>
                  </div>
                  <p className="text-gray-600">⚡ Tindakan diperlukan dalam hitungan hari</p>
                  <p className="text-xs text-gray-500">Level reorder atau stok rendah</p>
                </div>
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-4 h-2 bg-yellow-500 rounded"></div>
                    <span className="font-medium text-yellow-600">33-57% (SEDANG)</span>
                  </div>
                  <p className="text-gray-600">⚠️ Rencanakan tindakan dalam hitungan minggu</p>
                  <p className="text-xs text-gray-500">Stok berlebih atau mendekati batas</p>
                </div>
                <div>
                  <div className="flex items-center gap-2 mb-1">
                    <div className="w-4 h-2 bg-blue-500 rounded"></div>
                    <span className="font-medium text-blue-600">8-32% (RENDAH)</span>
                  </div>
                  <p className="text-gray-600">👁️ Pantau situasi</p>
                  <p className="text-xs text-gray-500">Pemantauan non-urgent</p>
                </div>
              </div>
              <div className="mt-4 p-3 bg-blue-50 rounded border-l-4 border-blue-400">
                <p className="text-sm text-blue-800">
                  <strong>💡 Tips:</strong> Persentase yang lebih tinggi berarti tindakan yang lebih mendesak diperlukan. Warna bar sesuai dengan tingkat urgensi.
                </p>
              </div>
            </div>
            
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
                            {item.urgencyLevel === 'CRITICAL' ? 'KRITIS' :
                             item.urgencyLevel === 'HIGH' ? 'TINGGI' :
                             item.urgencyLevel === 'MEDIUM' ? 'SEDANG' : 'RENDAH'}
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
                          <p className="text-xs text-gray-500">Skor Urgensi</p>
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
                          <p className="text-xs text-gray-600 mt-1">{item.urgencyScore} dari 12</p>
                        </div>
                        <div>
                          <p className="text-xs text-gray-500">Dampak Bisnis</p>
                          <p className="font-medium">{formatCurrency(item.businessImpact)}</p>
                        </div>
                        <div>
                          <p className="text-xs text-gray-500">Jangka Waktu</p>
                          <p className="font-medium text-red-600">{item.timeframe}</p>
                        </div>
                        {item.daysUntilStockout && (
                          <div>
                            <p className="text-xs text-gray-500">Hari Hingga Habis</p>
                            <p className="font-bold text-red-600">~{item.daysUntilStockout} hari</p>
                          </div>
                        )}
                      </div>
                      
                      <div className="mb-3">
                        <p className="text-sm text-gray-700 mb-1"><strong>Alasan:</strong> {item.reason}</p>
                        <p className="text-sm text-blue-700"><strong>Tindakan yang Direkomendasikan:</strong> {item.recommendedAction}</p>
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
                <label className="block text-xs font-medium text-gray-700 mb-1">Cari Produk</label>
                <input
                  type="text"
                  placeholder="Nama, kode, atau kategori"
                  value={filters.searchTerm}
                  onChange={(e) => setFilters({...filters, searchTerm: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                />
              </div>
              
              {/* Stock Status Filter */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Status Stok</label>
                <select
                  value={filters.stockStatus}
                  onChange={(e) => setFilters({...filters, stockStatus: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="All">Semua Status</option>
                  <option value="Normal">Normal</option>
                  <option value="Low Stock">Stok Rendah</option>
                  <option value="Overstock">Stok Berlebih</option>
                  <option value="Reorder">Perlu Reorder</option>
                  <option value="Out of Stock">Habis Stok</option>
                </select>
              </div>
              
              {/* ABC Class Filter */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Kelas ABC</label>
                <select
                  value={filters.abcClass}
                  onChange={(e) => setFilters({...filters, abcClass: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="All">Semua Kelas</option>
                  <option value="A">Kelas A (Nilai Tinggi)</option>
                  <option value="B">Kelas B (Nilai Sedang)</option>
                  <option value="C">Kelas C (Nilai Rendah)</option>
                </select>
              </div>
              
              {/* Category Filter */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Kategori</label>
                <select
                  value={filters.category}
                  onChange={(e) => setFilters({...filters, category: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="All">Semua Kategori</option>
                  {uniqueCategories.map(category => (
                    <option key={category} value={category}>{category}</option>
                  ))}
                </select>
              </div>
              
              {/* Variance Type Filter */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1">Jenis Selisih</label>
                <select
                  value={filters.varianceType}
                  onChange={(e) => setFilters({...filters, varianceType: e.target.value})}
                  className="w-full p-2 border border-gray-300 rounded-md text-sm focus:ring-blue-500 focus:border-blue-500"
                >
                  <option value="All">Semua Selisih</option>
                  <option value="Positive">Positif (+)</option>
                  <option value="Negative">Negatif (-)</option>
                  <option value="Zero">Nol (0)</option>
                  <option value="High Variance">Selisih Tinggi (&gt;10%)</option>
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
                  Hapus Filter
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
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Stok Pengaman</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Titik Reorder</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">EOQ</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Perputaran</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Aksi</th>
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
                            const message = `🔄 TINDAKAN REORDER DIPERLUKAN\n\nProduk: ${item.name} (${item.code})\nStok Saat Ini: ${item.actualStock}\nTitik Reorder: ${item.reorderPoint}\nKuantitas Pesanan yang Direkomendasikan (EOQ): ${item.eoq} unit\n\nTindakan: Buat pesanan untuk ${item.eoq} unit segera!`;
                            alert(message);
                          }}
                          className="bg-orange-500 text-white px-3 py-1 rounded text-xs hover:bg-orange-600 transition-colors"
                        >
                          Pesan Sekarang
                        </button>
                      )}
                      {item.stockStatus === 'Overstock' && (
                        <button
                          onClick={() => {
                            const overstock = item.actualStock - item.maxStock;
                            const message = `⚠️ TINDAKAN STOK BERLEBIH DIPERLUKAN\n\nProduk: ${item.name} (${item.code})\nStok Saat Ini: ${item.actualStock}\nStok Maksimum: ${item.maxStock}\nJumlah Stok Berlebih: ${overstock} unit\n\nTindakan yang Direkomendasikan:\n• Jalankan kampanye promosi\n• Transfer ke lokasi lain\n• Kurangi pesanan masa depan\n• Pertimbangkan diskon grosir`;
                            alert(message);
                          }}
                          className="bg-red-500 text-white px-3 py-1 rounded text-xs hover:bg-red-600 transition-colors"
                        >
                          Kurangi Stok
                        </button>
                      )}
                      {item.stockStatus === 'Low Stock' && (
                        <button
                          onClick={() => {
                            const shortfall = item.minStock - item.actualStock;
                            const message = `⚡ PERINGATAN STOK RENDAH\n\nProduk: ${item.name} (${item.code})\nStok Saat Ini: ${item.actualStock}\nStok Minimum: ${item.minStock}\nKekurangan: ${shortfall} unit\n\nTindakan: Pertimbangkan untuk memesan ${item.eoq} unit untuk menjaga level optimal.`;
                            alert(message);
                          }}
                          className="bg-yellow-500 text-white px-3 py-1 rounded text-xs hover:bg-yellow-600 transition-colors"
                        >
                          Pantau
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
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Min/Maks</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Waktu Tunggu</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Permintaan Rata2</th>
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
