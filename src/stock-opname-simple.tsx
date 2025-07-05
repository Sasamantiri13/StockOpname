import React, { useState, useEffect, useCallback } from 'react';
import { Download, Upload, Plus, Trash2, Calculator, AlertTriangle, CheckCircle, TrendingUp, Package, BarChart3 } from 'lucide-react';
import * as XLSX from 'xlsx';

// Type definitions
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

interface Summary {
  totalProducts?: number;
  totalInventoryValue?: number;
  totalVarianceValue?: number;
  accuracyRate?: number;
  lowStockItems?: number;
  overstockItems?: number;
  reorderItems?: number;
}

const StockOpnameSimple: React.FC = () => {
  const [products, setProducts] = useState<Product[]>([
    {
      id: 1,
      code: 'PRD001',
      name: 'Laptop Dell Inspiron',
      category: 'Electronics',
      systemStock: 50,
      actualStock: 48,
      unitCost: 8500000,
      minStock: 10,
      maxStock: 100,
      leadTime: 7,
      avgDemand: 8
    },
    {
      id: 2,
      code: 'PRD002',
      name: 'Mouse Wireless',
      category: 'Accessories',
      systemStock: 120,
      actualStock: 125,
      unitCost: 150000,
      minStock: 20,
      maxStock: 200,
      leadTime: 3,
      avgDemand: 15
    }
  ]);

  const [analysis, setAnalysis] = useState<Analysis[]>([]);
  const [summary, setSummary] = useState<Summary>({});

  // Calculation function
  const calculateAnalysis = useCallback(() => {
    const analysisData: Analysis[] = products.map(product => {
      const variance = product.actualStock - product.systemStock;
      const variancePercentage = product.systemStock > 0 ? (variance / product.systemStock) * 100 : 0;
      const varianceValue = variance * product.unitCost;
      
      const inventoryValue = product.actualStock * product.unitCost;
      const safetyStock = Math.ceil(product.avgDemand * Math.sqrt(product.leadTime) * 1.65);
      const reorderPoint = (product.avgDemand * product.leadTime) + safetyStock;
      
      const annualDemand = product.avgDemand * 365;
      const orderingCost = 50000;
      const holdingCostRate = 0.25;
      const holdingCost = product.unitCost * holdingCostRate;
      const eoq = Math.sqrt((2 * annualDemand * orderingCost) / holdingCost);
      
      // Stock Status - Fixed logic
      let stockStatus = 'Normal';
      const validMinStock = product.minStock > 0 ? product.minStock : 0;
      const validMaxStock = product.maxStock > validMinStock ? product.maxStock : validMinStock + 50;
      const validReorderPoint = reorderPoint > validMinStock ? reorderPoint : validMinStock + 5;
      
      if (product.actualStock <= 0) {
        stockStatus = 'Out of Stock';
      } else if (product.actualStock <= validMinStock) {
        stockStatus = 'Low Stock';
      } else if (product.actualStock >= validMaxStock) {
        stockStatus = 'Overstock';
      } else if (product.actualStock <= validReorderPoint && validReorderPoint > validMinStock) {
        stockStatus = 'Reorder';
      }
      
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
        abcClass: 'A', // Simplified
        cumulativePercentage: 0
      };
    });

    setAnalysis(analysisData);
    
    const totalVarianceValue = analysisData.reduce((sum, item) => sum + Math.abs(item.varianceValue), 0);
    const totalInventoryValue = analysisData.reduce((sum, item) => sum + item.inventoryValue, 0);
    const accuracyRate = analysisData.length > 0 ? ((analysisData.length - analysisData.filter(item => Math.abs(item.variancePercentage) > 5).length) / analysisData.length) * 100 : 100;
    
    setSummary({
      totalProducts: analysisData.length,
      totalInventoryValue,
      totalVarianceValue,
      accuracyRate,
      lowStockItems: analysisData.filter(item => item.stockStatus === 'Low Stock').length,
      overstockItems: analysisData.filter(item => item.stockStatus === 'Overstock').length,
      reorderItems: analysisData.filter(item => item.stockStatus === 'Reorder').length
    });
  }, [products]);

  useEffect(() => {
    calculateAnalysis();
  }, [calculateAnalysis]);

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

  const showActionAlert = (item: Analysis) => {
    const currentDate = new Date().toLocaleDateString('id-ID');
    
    if (item.stockStatus === 'Reorder') {
      const daysUntilStockout = Math.floor(item.actualStock / item.avgDemand);
      alert(`🔄 URGENT REORDER NOTIFICATION\n\nDate: ${currentDate}\nProduct: ${item.name} (${item.code})\n\n📊 CURRENT SITUATION:\n• Current Stock: ${item.actualStock} units\n• Daily Demand: ${item.avgDemand} units/day\n• Reorder Point: ${item.reorderPoint} units\n• Days until stockout: ~${daysUntilStockout} days\n\n🎯 REQUIRED ACTION:\n• Order Quantity (EOQ): ${item.eoq} units\n• Lead Time: ${item.leadTime} days\n\n⚡ URGENCY: HIGH - Order immediately to avoid stockout!`);
    } else if (item.stockStatus === 'Overstock') {
      const overstock = item.actualStock - item.maxStock;
      const overstockValue = overstock * item.unitCost;
      alert(`⚠️ OVERSTOCK ACTION REQUIRED\n\nDate: ${currentDate}\nProduct: ${item.name} (${item.code})\n\n📊 OVERSTOCK ANALYSIS:\n• Current Stock: ${item.actualStock} units\n• Maximum Stock: ${item.maxStock} units\n• Excess Quantity: ${overstock} units\n• Excess Value: ${formatCurrency(overstockValue)}\n\n🎯 RECOMMENDED ACTIONS:\n• Run promotional campaign (20-30% discount)\n• Transfer to other locations\n• Bundle with complementary products`);
    } else if (item.stockStatus === 'Low Stock') {
      const shortfall = item.minStock - item.actualStock;
      alert(`⚡ LOW STOCK WARNING\n\nDate: ${currentDate}\nProduct: ${item.name} (${item.code})\n\n📊 STOCK SITUATION:\n• Current Stock: ${item.actualStock} units\n• Minimum Stock: ${item.minStock} units\n• Shortfall: ${shortfall} units\n\n🎯 MONITORING ACTIONS:\n• Track consumption rate\n• Monitor demand patterns\n• Consider ordering ${item.eoq} units`);
    }
  };

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
                <p className="text-sm sm:text-base text-gray-600">Analisis mendalam dengan metode ABC, EOQ, dan Safety Stock</p>
              </div>
            </div>
          </div>
        </header>

        {/* Quick Action Dashboard - Items Requiring Immediate Attention */}
        {(analysis.filter(item => item.stockStatus === 'Reorder' || item.stockStatus === 'Overstock' || item.stockStatus === 'Low Stock').length > 0) && (
          <section className="bg-white rounded-lg shadow-lg p-4 sm:p-6 mb-6">
            <h2 className="text-xl font-bold text-red-600 mb-4 flex items-center">
              <AlertTriangle className="h-5 w-5 mr-2" />
              ⚡ Action Required - Immediate Attention Needed
            </h2>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              {analysis.filter(item => item.stockStatus === 'Reorder' || item.stockStatus === 'Overstock' || item.stockStatus === 'Low Stock').map(item => (
                <div key={item.id} className={`p-4 rounded-lg border-l-4 ${
                  item.stockStatus === 'Reorder' ? 'border-orange-500 bg-orange-50' :
                  item.stockStatus === 'Overstock' ? 'border-red-500 bg-red-50' :
                  'border-yellow-500 bg-yellow-50'
                }`}>
                  <div className="flex justify-between items-start mb-2">
                    <div>
                      <h3 className="font-bold text-gray-900">{item.name}</h3>
                      <p className="text-sm text-gray-600">{item.code}</p>
                    </div>
                    <span className={`px-2 py-1 text-xs rounded-full ${getStatusColor(item.stockStatus)}`}>
                      {item.stockStatus}
                    </span>
                  </div>
                  <div className="text-sm text-gray-700 mb-3">
                    <p>Current Stock: <span className="font-medium">{item.actualStock}</span></p>
                    {item.stockStatus === 'Reorder' && (
                      <p>Reorder Point: <span className="font-medium">{item.reorderPoint}</span></p>
                    )}
                    {item.stockStatus === 'Overstock' && (
                      <p>Max Stock: <span className="font-medium">{item.maxStock}</span></p>
                    )}
                    {item.stockStatus === 'Low Stock' && (
                      <p>Min Stock: <span className="font-medium">{item.minStock}</span></p>
                    )}
                  </div>
                  <div className="flex gap-2">
                    {item.stockStatus === 'Reorder' && (
                      <button
                        onClick={() => showActionAlert(item)}
                        className="bg-orange-500 text-white px-3 py-2 rounded text-sm hover:bg-orange-600 transition-colors font-medium flex-1"
                      >
                        🔄 Order {item.eoq} Units
                      </button>
                    )}
                    {item.stockStatus === 'Overstock' && (
                      <button
                        onClick={() => showActionAlert(item)}
                        className="bg-red-500 text-white px-3 py-2 rounded text-sm hover:bg-red-600 transition-colors font-medium flex-1"
                      >
                        📉 Reduce Stock
                      </button>
                    )}
                    {item.stockStatus === 'Low Stock' && (
                      <button
                        onClick={() => showActionAlert(item)}
                        className="bg-yellow-500 text-white px-3 py-2 rounded text-sm hover:bg-yellow-600 transition-colors font-medium flex-1"
                      >
                        👁️ Monitor Closely
                      </button>
                    )}
                  </div>
                </div>
              ))}
            </div>
          </section>
        )}

        {/* Summary Dashboard */}
        <section className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6 mb-6">
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg flex items-center justify-between">
            <div>
              <p className="text-gray-600 text-sm">Total Produk</p>
              <p className="text-2xl font-bold text-gray-900">{summary.totalProducts || 0}</p>
            </div>
            <Package className="h-8 w-8 text-blue-600" />
          </div>
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg flex items-center justify-between">
            <div>
              <p className="text-gray-600 text-sm">Nilai Inventory</p>
              <p className="text-2xl font-bold text-gray-900">{formatCurrency(summary.totalInventoryValue || 0)}</p>
            </div>
            <TrendingUp className="h-8 w-8 text-green-600" />
          </div>
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg flex items-center justify-between">
            <div>
              <p className="text-gray-600 text-sm">Akurasi Stock</p>
              <p className="text-2xl font-bold text-gray-900">{(summary.accuracyRate || 0).toFixed(1)}%</p>
            </div>
            <CheckCircle className="h-8 w-8 text-green-600" />
          </div>
          <div className="bg-white p-4 sm:p-6 rounded-lg shadow-lg flex items-center justify-between">
            <div>
              <p className="text-gray-600 text-sm">Item Perlu Perhatian</p>
              <p className="text-2xl font-bold text-red-600">{(summary.lowStockItems || 0) + (summary.overstockItems || 0)}</p>
            </div>
            <AlertTriangle className="h-8 w-8 text-red-600" />
          </div>
        </section>

        {/* Analysis Results */}
        <section className="bg-white rounded-lg shadow-lg p-4 sm:p-6">
          <h2 className="text-xl font-bold text-gray-900 mb-4 flex items-center">
            <Calculator className="h-5 w-5 mr-2" />
            Hasil Analisis DSS
          </h2>
          
          <div className="overflow-x-auto">
            <table className="min-w-full table-auto text-sm">
              <thead className="bg-gray-50">
                <tr>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Produk</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Status</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Stock</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">EOQ</th>
                  <th className="px-4 py-2 text-left text-xs font-medium text-gray-500 uppercase">Action</th>
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {analysis.map(item => (
                  <tr key={item.id}>
                    <td className="px-4 py-2 whitespace-nowrap">
                      <div className="font-medium text-gray-900">{item.name}</div>
                      <div className="text-sm text-gray-500">{item.code}</div>
                    </td>
                    <td className="px-4 py-2">
                      <span className={`px-2 py-1 text-xs rounded-full ${getStatusColor(item.stockStatus)}`}>
                        {item.stockStatus}
                      </span>
                    </td>
                    <td className="px-4 py-2 text-sm">{item.actualStock}</td>
                    <td className="px-4 py-2 text-sm font-medium">{item.eoq}</td>
                    <td className="px-4 py-2">
                      {(item.stockStatus === 'Reorder' || item.stockStatus === 'Overstock' || item.stockStatus === 'Low Stock') && (
                        <button
                          onClick={() => showActionAlert(item)}
                          className={`px-3 py-1 rounded text-xs transition-colors ${
                            item.stockStatus === 'Reorder' ? 'bg-orange-500 hover:bg-orange-600 text-white' :
                            item.stockStatus === 'Overstock' ? 'bg-red-500 hover:bg-red-600 text-white' :
                            'bg-yellow-500 hover:bg-yellow-600 text-white'
                          }`}
                        >
                          {item.stockStatus === 'Reorder' ? 'Reorder' : 
                           item.stockStatus === 'Overstock' ? 'Reduce' : 'Monitor'}
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
      </div>
    </div>
  );
};

export default StockOpnameSimple;

