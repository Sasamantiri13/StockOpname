import React, { useState, useEffect, useCallback } from 'react';
import { Download, Upload, Plus, Trash2, Calculator, AlertTriangle, CheckCircle, TrendingUp, Package, BarChart3 } from 'lucide-react';
import * as XLSX from 'xlsx';

const StockOpnameDSS = () => {
  const [products, setProducts] = useState([
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

  const [analysis, setAnalysis] = useState([]);
  const [summary, setSummary] = useState({});
  const [recommendations, setRecommendations] = useState([]);

  // Fungsi perhitungan DSS
  const calculateAnalysis = useCallback(() => {
    const analysisData = products.map(product => {
      const variance = product.actualStock - product.systemStock;
      const variancePercentage = product.systemStock > 0 ? (variance / product.systemStock) * 100 : 0;
      const varianceValue = variance * product.unitCost;
      
      // ABC Analysis berdasarkan nilai inventory
      const inventoryValue = product.actualStock * product.unitCost;
      
      // Safety Stock calculation
      const safetyStock = Math.ceil(product.avgDemand * Math.sqrt(product.leadTime) * 1.65); // 95% service level
      
      // Reorder Point
      const reorderPoint = (product.avgDemand * product.leadTime) + safetyStock;
      
      // Economic Order Quantity (EOQ) - simplified
      const annualDemand = product.avgDemand * 365;
      const orderingCost = 50000; // Rp 50,000 per order
      const holdingCostRate = 0.25; // 25% of unit cost
      const holdingCost = product.unitCost * holdingCostRate;
      const eoq = Math.sqrt((2 * annualDemand * orderingCost) / holdingCost);
      
      // Stock Status
      let stockStatus = 'Normal';
      if (product.actualStock <= product.minStock) stockStatus = 'Low Stock';
      else if (product.actualStock >= product.maxStock) stockStatus = 'Overstock';
      else if (product.actualStock <= reorderPoint) stockStatus = 'Reorder';
      
      // Turnover ratio
      const turnoverRatio = annualDemand / product.actualStock;
      
    
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
        annualDemand
      };
    });

    // ABC Classification
    const sortedByValue = [...analysisData].sort((a, b) => b.inventoryValue - a.inventoryValue);
    const totalValue = sortedByValue.reduce((sum, item) => sum + item.inventoryValue, 0);
    
    let cumulativeValue = 0;
    const classifiedData = sortedByValue.map(item => {
      cumulativeValue += item.inventoryValue;
      const cumulativePercentage = (cumulativeValue / totalValue) * 100;
      
      let abcClass = 'C';
      if (cumulativePercentage <= 80) abcClass = 'A';
      else if (cumulativePercentage <= 95) abcClass = 'B';
      
      return { ...item, abcClass, cumulativePercentage };
    });

    // Sort back to original order
    const finalAnalysis = analysisData.map(item => {
      const classified = classifiedData.find(c => c.id === item.id);
      return { ...item, abcClass: classified.abcClass, cumulativePercentage: classified.cumulativePercentage };
    });

    setAnalysis(finalAnalysis);
    
    // Calculate summary
    const totalVarianceValue = finalAnalysis.reduce((sum, item) => sum + Math.abs(item.varianceValue), 0);
    const totalInventoryValue = finalAnalysis.reduce((sum, item) => sum + item.inventoryValue, 0);
    const accuracyRate = ((finalAnalysis.length - finalAnalysis.filter(item => Math.abs(item.variancePercentage) > 5).length) / finalAnalysis.length) * 100;
    
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
    const recs = [];
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
  }, [products]);

  useEffect(() => {
    calculateAnalysis();
  }, [calculateAnalysis]);

  const addProduct = () => {
    const newProduct = {
      id: Date.now(),
      code: `PRD${String(products.length + 1).padStart(3, '0')}`,
      name: '',
      category: '',
      systemStock: 0,
      actualStock: 0,
      unitCost: 0,
      minStock: 0,
      maxStock: 0,
      leadTime: 1,
      avgDemand: 1
    };
    setProducts([...products, newProduct]);
  };

  const deleteProduct = (id) => {
    setProducts(products.filter(p => p.id !== id));
  };

  const updateProduct = (id, field, value) => {
    setProducts(products.map(p => p.id === id ? { ...p, [field]: value } : p));
  };

  const exportToExcel = () => {
    const wsData = analysis.map(item => ({
      'Kode Produk': item.code,
      'Nama Produk': item.name,
      'Kategori': item.category,
      'Stok Sistem': item.systemStock,
      'Stok Aktual': item.actualStock,
      'Selisih': item.variance,
      'Selisih (%)': `${item.variancePercentage.toFixed(2)}%`,
      'Harga Satuan': item.unitCost,
      'Nilai Selisih': item.varianceValue,
      'Nilai Inventory': item.inventoryValue,
      'Kelas ABC': item.abcClass,
      'Status Stok': item.stockStatus,
      'Safety Stock': item.safetyStock,
      'Reorder Point': item.reorderPoint,
      'EOQ': item.eoq,
      'Turnover Ratio': item.turnoverRatio.toFixed(2),
      'Demand Tahunan': item.annualDemand
    }));

    const summaryData = [
      ['RINGKASAN ANALISIS'],
      ['Total Produk', summary.totalProducts],
      ['Total Nilai Inventory', summary.totalInventoryValue],
      ['Total Nilai Selisih', summary.totalVarianceValue],
      ['Tingkat Akurasi (%)', `${summary.accuracyRate.toFixed(2)}%`],
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

  const importFromExcel = (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      const workbook = XLSX.read(e.target.result, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet);
      
      const importedProducts = data.map((row, index) => ({
        id: Date.now() + index,
        code: row['Kode Produk'] || `PRD${String(index + 1).padStart(3, '0')}`,
        name: row['Nama Produk'] || '',
        category: row['Kategori'] || '',
        systemStock: Number(row['Stok Sistem']) || 0,
        actualStock: Number(row['Stok Aktual']) || 0,
        unitCost: Number(row['Harga Satuan']) || 0,
        minStock: Number(row['Min Stock']) || 0,
        maxStock: Number(row['Max Stock']) || 0,
        leadTime: Number(row['Lead Time']) || 1,
        avgDemand: Number(row['Avg Demand']) || 1
      }));
      
      setProducts(importedProducts);
    };
    reader.readAsBinaryString(file);
    event.target.value = '';
  };

  const formatCurrency = (value) => {
    return new Intl.NumberFormat('id-ID', {
      style: 'currency',
      currency: 'IDR'
    }).format(value);
  };

  const getStatusColor = (status) => {
    switch (status) {
      case 'Low Stock': return 'text-red-600 bg-red-100';
      case 'Overstock': return 'text-orange-600 bg-orange-100';
      case 'Reorder': return 'text-yellow-600 bg-yellow-100';
      default: return 'text-green-600 bg-green-100';
    }
  };

  const getABCColor = (abcClass) => {
    switch (abcClass) {
      case 'A': return 'text-red-600 bg-red-100 font-bold';
      case 'B': return 'text-orange-600 bg-orange-100 font-semibold';
      case 'C': return 'text-green-600 bg-green-100';
      default: return 'text-gray-600 bg-gray-100';
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        {/* Header */}
        <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <Package className="h-8 w-8 text-blue-600" />
              <div>
                <h1 className="text-3xl font-bold text-gray-900">Sistem Penunjang Keputusan Stock Opname</h1>
                <p className="text-gray-600">Analisis mendalam dengan metode ABC, EOQ, dan Safety Stock</p>
              </div>
            </div>
            <div className="flex space-x-3">
              <label className="bg-green-600 text-white px-4 py-2 rounded-lg hover:bg-green-700 cursor-pointer flex items-center space-x-2">
                <Upload className="h-4 w-4" />
                <span>Import Excel</span>
                <input type="file" accept=".xlsx,.xls" onChange={importFromExcel} className="hidden" />
              </label>
              <button
                onClick={exportToExcel}
                className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 flex items-center space-x-2"
              >
                <Download className="h-4 w-4" />
                <span>Export Excel</span>
              </button>
            </div>
          </div>
        </div>

        {/* Summary Dashboard */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-6 mb-6">
          <div className="bg-white p-6 rounded-lg shadow-lg">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-gray-600 text-sm">Total Produk</p>
                <p className="text-2xl font-bold text-gray-900">{summary.totalProducts || 0}</p>
              </div>
              <Package className="h-8 w-8 text-blue-600" />
            </div>
          </div>
          <div className="bg-white p-6 rounded-lg shadow-lg">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-gray-600 text-sm">Nilai Inventory</p>
                <p className="text-2xl font-bold text-gray-900">{formatCurrency(summary.totalInventoryValue || 0)}</p>
              </div>
              <TrendingUp className="h-8 w-8 text-green-600" />
            </div>
          </div>
          <div className="bg-white p-6 rounded-lg shadow-lg">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-gray-600 text-sm">Akurasi Stock</p>
                <p className="text-2xl font-bold text-gray-900">{(summary.accuracyRate || 0).toFixed(1)}%</p>
              </div>
              <CheckCircle className="h-8 w-8 text-green-600" />
            </div>
          </div>
          <div className="bg-white p-6 rounded-lg shadow-lg">
            <div className="flex items-center justify-between">
              <div>
                <p className="text-gray-600 text-sm">Item Perlu Perhatian</p>
                <p className="text-2xl font-bold text-red-600">{(summary.lowStockItems || 0) + (summary.overstockItems || 0)}</p>
              </div>
              <AlertTriangle className="h-8 w-8 text-red-600" />
            </div>
          </div>
        </div>

        {/* Recommendations Panel */}
        {recommendations.length > 0 && (
          <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
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
          </div>
        )}

        {/* Product Management */}
        <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-xl font-bold text-gray-900">Data Produk</h2>
            <button
              onClick={addProduct}
              className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 flex items-center space-x-2"
            >
              <Plus className="h-4 w-4" />
              <span>Tambah Produk</span>
            </button>
          </div>
          
          <div className="overflow-x-auto">
            <table className="min-w-full table-auto">
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
                    <td className="px-4 py-2">
                      <input
                        type="text"
                        value={product.code}
                        onChange={(e) => updateProduct(product.id, 'code', e.target.value)}
                        className="w-20 p-1 border rounded text-sm"
                      />
                    </td>
                    <td className="px-4 py-2">
                      <input
                        type="text"
                        value={product.name}
                        onChange={(e) => updateProduct(product.id, 'name', e.target.value)}
                        className="w-32 p-1 border rounded text-sm"
                      />
                    </td>
                    <td className="px-4 py-2">
                      <input
                        type="text"
                        value={product.category}
                        onChange={(e) => updateProduct(product.id, 'category', e.target.value)}
                        className="w-24 p-1 border rounded text-sm"
                      />
                    </td>
                    <td className="px-4 py-2">
                      <input
                        type="number"
                        value={product.systemStock}
                        onChange={(e) => updateProduct(product.id, 'systemStock', parseInt(e.target.value) || 0)}
                        className="w-16 p-1 border rounded text-sm"
                      />
                    </td>
                    <td className="px-4 py-2">
                      <input
                        type="number"
                        value={product.actualStock}
                        onChange={(e) => updateProduct(product.id, 'actualStock', parseInt(e.target.value) || 0)}
                        className="w-16 p-1 border rounded text-sm"
                      />
                    </td>
                    <td className="px-4 py-2">
                      <input
                        type="number"
                        value={product.unitCost}
                        onChange={(e) => updateProduct(product.id, 'unitCost', parseInt(e.target.value) || 0)}
                        className="w-24 p-1 border rounded text-sm"
                      />
                    </td>
                    <td className="px-4 py-2">
                      <div className="flex space-x-1">
                        <input
                          type="number"
                          value={product.minStock}
                          onChange={(e) => updateProduct(product.id, 'minStock', parseInt(e.target.value) || 0)}
                          className="w-12 p-1 border rounded text-sm"
                          placeholder="Min"
                        />
                        <input
                          type="number"
                          value={product.maxStock}
                          onChange={(e) => updateProduct(product.id, 'maxStock', parseInt(e.target.value) || 0)}
                          className="w-12 p-1 border rounded text-sm"
                          placeholder="Max"
                        />
                      </div>
                    </td>
                    <td className="px-4 py-2">
                      <input
                        type="number"
                        value={product.leadTime}
                        onChange={(e) => updateProduct(product.id, 'leadTime', parseInt(e.target.value) || 1)}
                        className="w-16 p-1 border rounded text-sm"
                      />
                    </td>
                    <td className="px-4 py-2">
                      <input
                        type="number"
                        value={product.avgDemand}
                        onChange={(e) => updateProduct(product.id, 'avgDemand', parseInt(e.target.value) || 1)}
                        className="w-16 p-1 border rounded text-sm"
                      />
                    </td>
                    <td className="px-4 py-2">
                      <button
                        onClick={() => deleteProduct(product.id)}
                        className="text-red-600 hover:text-red-800"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Analysis Results */}
        <div className="bg-white rounded-lg shadow-lg p-6">
          <div className="flex items-center justify-between mb-4">
            <h2 className="text-xl font-bold text-gray-900 flex items-center">
              <Calculator className="h-5 w-5 mr-2" />
              Hasil Analisis DSS
            </h2>
          </div>
          
          <div className="overflow-x-auto">
            <table className="min-w-full table-auto">
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
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {analysis.map(item => (
                  <tr key={item.id}>
                    <td className="px-4 py-2">
                      <div>
                        <div className="font-medium text-gray-900">{item.name}</div>
                        <div className="text-sm text-gray-500">{item.code}</div>
                      </div>
                    </td>
                    <td className="px-4 py-2">
                      <span className={`font-medium ${item.variance > 0 ? 'text-green-600' : 'text-red-600'}`}>
                        {item.variance > 0 ? '+' : ''}{item.variance}
                      </span>
                    </td>
                    <td className="px-4 py-2">
                      <span className={`font-medium ${Math.abs(item.variancePercentage) > 10 ? 'text-red-600' : 'text-gray-900'}`}>
                        {item.variancePercentage.toFixed(1)}%
                      </span>
                    </td>
                    <td className="px-4 py-2">
                      <span className={`text-sm ${item.varianceValue > 0 ? 'text-green-600' : 'text-red-600'}`}>
                        {formatCurrency(item.varianceValue)}
                      </span>
                    </td>
                    <td className="px-4 py-2">
                      <span className={`px-2 py-1 text-xs rounded-full ${getABCColor(item.abcClass)}`}>
                        {item.abcClass}
                      </span>
                    </td>
                    <td className="px-4 py-2">
                      <span className={`px-2 py-1 text-xs rounded-full ${getStatusColor(item.stockStatus)}`}>
                        {item.stockStatus}
                      </span>
                    </td>
                    <td className="px-4 py-2 text-sm">{item.safetyStock}</td>
                    <td className="px-4 py-2 text-sm">{item.reorderPoint}</td>
                    <td className="px-4 py-2 text-sm font-medium">{item.eoq}</td>
                    <td className="px-4 py-2 text-sm">{item.turnoverRatio.toFixed(1)}x</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};

export default StockOpnameDSS;