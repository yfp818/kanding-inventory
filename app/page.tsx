'use client';

import React, { useState, useEffect, useRef } from 'react';
import { supabase } from '../lib/supabase';
import * as XLSX from 'xlsx';
import { Toaster, toast } from 'react-hot-toast';

import { 
  Bell, Layers, Home, Settings, Loader2, CircleUserRound, Zap, 
  Plus, Minus, AlertTriangle, Search, X, Database, Store, Pencil, 
  PlusCircle, Save, CheckCircle2, ArrowRight, Upload, ClipboardList,
  ChevronLeft, BarChart3, Wallet, PackageOpen, TrendingDown, Check
} from 'lucide-react';

import { InventoryItem, ExcelRow } from '../types';

const extractErrorMessage = (err: unknown): string => {
  if (err instanceof TypeError && err.message === 'Type error') {
    return '網路連線異常。請確認 .env.local 內的網址正確，並重啟伺服器。';
  }
  if (err instanceof Error) return err.message;
  if (typeof err === 'object' && err !== null) {
    const obj = err as Record<string, unknown>;
    if (obj.message) return String(obj.message);
    return JSON.stringify(obj);
  }
  return String(err);
};

export default function KanDingLiveStocktake() {
  const [items, setItems] = useState<InventoryItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  
  // 底部導覽列切換視角
  const [activeTab, setActiveTab] = useState<string>('home');
  
  // ==========================================
  // 批次盤點草稿狀態 (Batch Draft System)
  // 解決按鈕沒反應的問題：先在前端記憶修改，最後一次送出
  // ==========================================
  const [draftStocks, setDraftStocks] = useState<Record<number, number>>({});
  const [isSavingBatch, setIsSavingBatch] = useState<boolean>(false);
  const hasDraftChanges = Object.keys(draftStocks).length > 0;

  const [searchQuery, setSearchQuery] = useState<string>('');
  const [selectedCategory, setSelectedCategory] = useState<string>('全部');
  const [homeFilter, setHomeFilter] = useState<'all' | 'low'>('all'); 
  
  const [activeSupplier, setActiveSupplier] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [isSavingItem, setIsSavingItem] = useState<boolean>(false);
  const [editForm, setEditForm] = useState<Partial<InventoryItem>>({});

  const dynamicCategories = ['全部', ...Array.from(new Set(items.map(i => i.category).filter(Boolean)))].sort();

  const totalItems = items.length;
  const lowStockItems = items.filter(item => {
    const current = draftStocks[item.id] ?? item.current_stock;
    return current < item.min_stock;
  }).length;
  const syncProgress = totalItems > 0 ? 100 : 0;
  
  const totalAssetValue = items.reduce((sum, item) => sum + ((draftStocks[item.id] ?? item.current_stock) * (item.price || 0)), 0);
  
  const totalLowStockValue = items
    .filter(item => (draftStocks[item.id] ?? item.current_stock) < item.min_stock)
    .reduce((sum, item) => {
      const current = draftStocks[item.id] ?? item.current_stock;
      return sum + ((item.min_stock - current) * (item.price || 0));
    }, 0);

  const logSystemError = async (action: string, errorMessage: string) => {
    console.warn(`[System Error - ${action}]:`, errorMessage);
    try {
      await supabase.from('system_errors').insert([{ action, error_message: errorMessage, created_at: new Date().toISOString() }]);
    } catch (logErr) {
      console.warn('系統日誌寫入失敗', logErr);
    }
  };

  const fetchInventory = async () => {
    setLoading(true);
    try {
      const { data, error } = await supabase.from('kanding_inventory').select('*').order('id', { ascending: true });
      if (error) throw error;
      if (data) setItems(data as InventoryItem[]);
      setDraftStocks({}); // 重新載入時清空草稿
    } catch (err: unknown) {
      logSystemError('fetchInventory', extractErrorMessage(err));
      toast.error('無法載入庫存資料，請檢查網路連線。');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { fetchInventory(); }, []);

  // ------------------------------------------
  // 前端即時草稿修改 (不打 API，保證 UI 順暢)
  // ------------------------------------------
  const handleDraftChange = (id: number, newValue: number, originalValue: number) => {
    const finalValue = Math.max(0, newValue);
    setDraftStocks(prev => {
      const newDraft = { ...prev };
      if (finalValue === originalValue) {
        delete newDraft[id]; // 若改回原值，則移出草稿
      } else {
        newDraft[id] = finalValue;
      }
      return newDraft;
    });
  };

  // ------------------------------------------
  // 批次寫入資料庫 (Batch Upsert)
  // ------------------------------------------
  const submitBatchChanges = async () => {
    if (!hasDraftChanges) return;
    setIsSavingBatch(true);
    const toastId = toast.loading('正在同步盤點數據至雲端...');

    try {
      const updates = Object.entries(draftStocks).map(([id, newStock]) => ({
        id: Number(id),
        current_stock: newStock,
        updated_at: new Date()
      }));

      // 使用 Promise.all 並行寫入以提高速度
      const updatePromises = updates.map(updateData => 
        supabase.from('kanding_inventory').update({ 
          current_stock: updateData.current_stock, 
          updated_at: updateData.updated_at 
        }).eq('id', updateData.id)
      );

      await Promise.all(updatePromises);
      
      toast.success(`成功更新 ${updates.length} 項商品庫存！`, { id: toastId });
      setDraftStocks({});
      await fetchInventory();
    } catch (err: unknown) {
      logSystemError('submitBatchChanges', extractErrorMessage(err));
      toast.error(`同步失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsSavingBatch(false);
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    if (!process.env.NEXT_PUBLIC_SUPABASE_URL) {
      toast.error('請確認環境變數 NEXT_PUBLIC_SUPABASE_URL 設定。');
      return;
    }

    setIsUploading(true);
    const toastId = toast.loading('解析檔案中...');

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]]; 
      const jsonData = XLSX.utils.sheet_to_json<ExcelRow>(worksheet, { range: 2 });

      if (jsonData.length === 0) throw new Error('解析無資料，請確認檔案格式。');

      const getSafeNumber = (val: unknown): number => {
        if (val === undefined || val === null || val === '') return 0;
        const parsed = parseFloat(String(val));
        return isNaN(parsed) ? 0 : parsed;
      };

      const formattedData = jsonData.map((row) => ({
        name: String(row['品項名稱'] ?? row['品名'] ?? row['名稱'] ?? '未命名品項'),
        category: String(row['類別'] ?? row['分類'] ?? '未分類'),
        supplier: String(row['供應商'] ?? '未指定'),
        price: getSafeNumber(row['單價']),
        current_stock: getSafeNumber(row['盤點數量'] ?? row['目前庫存']),
        min_stock: getSafeNumber(row['安全存量'] ?? row['安全水位']),
        unit: String(row['單位'] ?? row['規格'] ?? '個'),
      }));

      toast.loading(`寫入 ${formattedData.length} 筆資料...`, { id: toastId });
      const { error } = await supabase.from('kanding_inventory').insert(formattedData);
      if (error) throw error;
      
      toast.success(`匯入完成！共 ${formattedData.length} 筆`, { id: toastId });
      fetchInventory(); 
      setActiveTab('home');
    } catch (err: unknown) {
      toast.error(`處理失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const openAddModal = (presets?: { name?: string, supplier?: string }) => {
    setEditForm({
      name: presets?.name || '',
      category: selectedCategory === '全部' ? '未分類' : selectedCategory,
      supplier: presets?.supplier || '未指定',
      unit: '個',
      price: 0,
      current_stock: 0,
      min_stock: 0
    });
    setIsModalOpen(true);
  };

  const openEditModal = (item: InventoryItem) => {
    setEditForm({
      ...item,
      current_stock: draftStocks[item.id] ?? item.current_stock 
    });
    setIsModalOpen(true);
  };

  const handleSaveItem = async () => {
    if (!editForm.name) return toast.error('名稱不可為空');
    setIsSavingItem(true);
    const toastId = toast.loading('處理中...');

    try {
      const payload = {
        name: editForm.name,
        category: editForm.category || '未分類',
        supplier: editForm.supplier || '未指定',
        unit: editForm.unit || '個',
        price: Number(editForm.price) || 0,
        current_stock: Number(editForm.current_stock) || 0,
        min_stock: Number(editForm.min_stock) || 0,
        updated_at: new Date()
      };

      if (editForm.id) {
        const { error } = await supabase.from('kanding_inventory').update(payload).eq('id', editForm.id);
        if (error) throw error;
        
        // 若修改成功，清除該項目的本地草稿，避免衝突
        setDraftStocks(prev => {
          const newDraft = { ...prev };
          delete newDraft[editForm.id as number];
          return newDraft;
        });
        toast.success('更新成功', { id: toastId });
      } else {
        const { error } = await supabase.from('kanding_inventory').insert([payload]);
        if (error) throw error;
        toast.success('新增成功', { id: toastId });
        setSearchQuery(''); 
      }
      
      await fetchInventory();
      setIsModalOpen(false);
    } catch (err: unknown) {
      logSystemError('SaveItem', extractErrorMessage(err));
      toast.error(`失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsSavingItem(false);
    }
  };

  const getFilteredItems = (baseItems: InventoryItem[]) => {
    return baseItems.filter(item => {
      const currentStock = draftStocks[item.id] ?? item.current_stock;
      if (homeFilter === 'low' && currentStock >= item.min_stock) return false;
      const matchCategory = selectedCategory === '全部' || item.category === selectedCategory;
      const matchSearch = item.name.toLowerCase().includes(searchQuery.toLowerCase());
      return matchCategory && matchSearch;
    }).sort((a, b) => {
      const aCurrent = draftStocks[a.id] ?? a.current_stock;
      const bCurrent = draftStocks[b.id] ?? b.current_stock;
      const aIsLow = aCurrent < a.min_stock;
      const bIsLow = bCurrent < b.min_stock;
      if (aIsLow && !bIsLow) return -1;
      if (!aIsLow && bIsLow) return 1;
      if (a.category !== b.category) return a.category.localeCompare(b.category, 'zh-TW');
      return a.name.localeCompare(b.name, 'zh-TW');
    });
  };

  const homeProcessedItems = getFilteredItems(items);

  const supplierItems = items
    .filter(item => item.supplier === activeSupplier)
    .sort((a, b) => {
      const aCurrent = draftStocks[a.id] ?? a.current_stock;
      const bCurrent = draftStocks[b.id] ?? b.current_stock;
      const aIsLow = aCurrent < a.min_stock;
      const bIsLow = bCurrent < b.min_stock;
      if (aIsLow && !bIsLow) return -1;
      if (!aIsLow && bIsLow) return 1;
      return a.name.localeCompare(b.name, 'zh-TW');
    });

  const suppliersList = Array.from(new Set(items.map(i => i.supplier || '未指定'))).sort();
  
  const getSupplierStats = (sup: string) => {
    const supItems = items.filter(i => i.supplier === sup);
    const lowCount = supItems.filter(i => (draftStocks[i.id] ?? i.current_stock) < i.min_stock).length;
    const totalValue = supItems.reduce((acc, curr) => acc + ((draftStocks[curr.id] ?? curr.current_stock) * (curr.price || 0)), 0);
    return { count: supItems.length, lowCount, totalValue };
  };

  return (
    <div className="min-h-screen bg-[#F5F5F7] pb-40 font-sans selection:bg-gray-200">
      <Toaster position="top-center" toastOptions={{ className: 'rounded-[16px] text-sm font-medium shadow-lg' }} />
      
      {/* 頂部 Header */}
      <div className="px-6 pt-12 pb-6 flex justify-between items-center sticky top-0 z-30 bg-[#F5F5F7]/80 backdrop-blur-xl">
        <div className="flex items-center gap-4">
          <div className="w-[48px] h-[48px] bg-white rounded-full flex items-center justify-center text-gray-800 shadow-[0_4px_20px_rgba(0,0,0,0.04)]">
            <CircleUserRound size={26} strokeWidth={1.5} />
          </div>
          <div>
            <h1 className="text-[20px] font-bold text-gray-900 tracking-tight">崁頂總店庫存</h1>
            <p className="text-[12px] font-bold text-gray-400 mt-0.5">
              {activeTab === 'home' ? '現場作業模式' : activeTab === 'dashboard' ? '財務決策模式' : '系統維護模式'}
            </p>
          </div>
        </div>
        <button 
          onClick={fetchInventory} 
          disabled={loading || isSavingBatch}
          className="w-[48px] h-[48px] bg-white rounded-full flex items-center justify-center shadow-[0_4px_20px_rgba(0,0,0,0.04)] relative text-gray-800 transition-all active:scale-95 disabled:opacity-50"
        >
          <Zap size={22} strokeWidth={1.5} className={loading ? "animate-pulse text-gray-400" : ""} />
        </button>
      </div>

      <div className="px-6 space-y-6">
        
        {/* ========================================== */}
        {/* 員工視角：日常盤點與查貨 (Task-Oriented) */}
        {/* ========================================== */}
        {activeTab === 'home' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-6">
            
            <div className="relative flex items-center w-full h-[56px] bg-white rounded-[20px] shadow-[0_8px_30px_rgb(0,0,0,0.04)] px-5 transition-all focus-within:ring-2 ring-gray-200">
              <Search size={20} className="text-gray-400" />
              <input 
                type="text" 
                placeholder="輸入條碼或品項名稱搜尋..." 
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full h-full bg-transparent outline-none px-3 text-[15px] font-medium text-gray-900 placeholder:text-gray-400"
              />
              {searchQuery && (
                <button onClick={() => setSearchQuery('')} className="p-1.5 bg-gray-100 rounded-full text-gray-500 hover:bg-gray-200">
                  <X size={14} strokeWidth={3} />
                </button>
              )}
            </div>

            {searchQuery && homeProcessedItems.length === 0 && (
              <div 
                onClick={() => openAddModal({ name: searchQuery })}
                className="bg-white border border-gray-100 rounded-[24px] p-5 flex items-center justify-between cursor-pointer hover:shadow-md transition-all shadow-[0_8px_30px_rgb(0,0,0,0.04)]"
              >
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 bg-gray-50 text-gray-600 rounded-full flex items-center justify-center">
                    <PlusCircle size={24} strokeWidth={1.5} />
                  </div>
                  <div>
                    <h4 className="text-[16px] font-bold text-gray-900">建立 "{searchQuery}"</h4>
                    <p className="text-[13px] font-medium text-gray-400">點擊新增至資料庫</p>
                  </div>
                </div>
                <ArrowRight size={20} className="text-gray-400" />
              </div>
            )}

            <div className="flex gap-2 overflow-x-auto hide-scrollbar">
              {dynamicCategories.map(cat => (
                <button 
                  key={cat} 
                  onClick={() => setSelectedCategory(cat)} 
                  className={`px-5 py-3 rounded-[16px] text-[14px] font-bold whitespace-nowrap transition-all duration-300 ${
                    selectedCategory === cat 
                      ? 'bg-gray-900 text-white shadow-lg' 
                      : 'bg-white text-gray-500 shadow-[0_4px_20px_rgb(0,0,0,0.03)]'
                  }`}
                >
                  {cat}
                </button>
              ))}
            </div>

            <div className="space-y-4">
              {loading ? (
                 <div className="py-20 flex justify-center"><Loader2 className="animate-spin text-gray-300" size={32} /></div>
              ) : homeProcessedItems.length === 0 && !searchQuery ? (
                 <div className="py-20 text-center text-gray-400 font-medium text-[15px]">查無商品資訊</div>
              ) : (
                homeProcessedItems.map(item => {
                  const currentStock = draftStocks[item.id] ?? item.current_stock;
                  const isLow = currentStock < item.min_stock;
                  const isModified = draftStocks[item.id] !== undefined;

                  return (
                    <div key={item.id} className={`bg-white rounded-[24px] p-5 transition-all duration-300 border ${isModified ? 'border-gray-900 shadow-md' : 'border-transparent shadow-[0_8px_30px_rgb(0,0,0,0.03)]'}`}>
                      <div className="flex items-start justify-between mb-4">
                        <div className="flex-1 pr-4">
                          <div className="flex items-center gap-2 mb-1.5">
                             {isLow && <span className="bg-red-50 text-red-600 text-[11px] px-2 py-0.5 rounded-[8px] font-bold flex items-center gap-1"><AlertTriangle size={12}/> 低庫存</span>}
                             {isModified && <span className="bg-gray-100 text-gray-700 text-[11px] px-2 py-0.5 rounded-[8px] font-bold">待儲存</span>}
                          </div>
                          <h2 className="text-[17px] font-bold text-gray-900 leading-tight flex items-center gap-2">
                            {item.name}
                            <button onClick={() => openEditModal(item)} className="text-gray-300 hover:text-gray-600 p-1 transition-colors">
                              <Pencil size={14} strokeWidth={2.5} />
                            </button>
                          </h2>
                          <p className="text-[13px] font-medium text-gray-400 mt-1">
                            {item.category} • 安全水位: {item.min_stock} {item.unit}
                          </p>
                        </div>
                      </div>
                      
                      <div className="bg-[#F5F5F7] rounded-[18px] p-1.5 flex items-center justify-between border border-gray-100">
                        <button 
                          onClick={() => handleDraftChange(item.id, currentStock - 1, item.current_stock)} 
                          className="w-12 h-12 rounded-[14px] bg-white text-gray-600 flex items-center justify-center shadow-sm active:scale-90 transition-transform"
                        >
                          <Minus size={20} strokeWidth={2.5} />
                        </button>
                        <div className="flex-1 flex justify-center">
                          <span className={`text-[22px] font-bold ${isModified ? 'text-gray-900' : 'text-gray-600'}`}>
                            {currentStock}
                          </span>
                        </div>
                        <button 
                          onClick={() => handleDraftChange(item.id, currentStock + 1, item.current_stock)} 
                          className="w-12 h-12 rounded-[14px] bg-white text-gray-600 flex items-center justify-center shadow-sm active:scale-90 transition-transform"
                        >
                          <Plus size={20} strokeWidth={2.5} />
                        </button>
                      </div>
                    </div>
                  );
                })
              )}
            </div>
          </div>
        )}

        {/* ========================================== */}
        {/* 供應商視角：進銷存矩陣 (Supplier Matrix) */}
        {/* ========================================== */}
        {activeTab === 'suppliers' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
            {!activeSupplier && (
              <div className="space-y-4">
                 <div className="flex items-center justify-between px-1 mb-4">
                    <h2 className="text-[20px] font-bold text-gray-900 tracking-tight">供應鏈管理</h2>
                 </div>
                 
                 {suppliersList.map(sup => {
                   const stats = getSupplierStats(sup);
                   return (
                     <div 
                       key={sup} 
                       onClick={() => setActiveSupplier(sup)}
                       className="bg-white rounded-[28px] p-6 shadow-[0_8px_30px_rgba(0,0,0,0.03)] flex items-center justify-between cursor-pointer hover:shadow-lg transition-all active:scale-[0.98]"
                     >
                        <div className="flex items-center gap-4">
                          <div className={`w-14 h-14 rounded-[20px] flex items-center justify-center ${stats.lowCount > 0 ? 'bg-red-50 text-red-500' : 'bg-gray-50 text-gray-500'}`}>
                            <Store size={26} strokeWidth={1.5} />
                          </div>
                          <div>
                            <h3 className="text-[18px] font-bold text-gray-900">{sup}</h3>
                            <p className="text-[13px] font-bold text-gray-400 mt-1">共 {stats.count} 項商品</p>
                          </div>
                        </div>
                        <div className="text-right">
                           {stats.lowCount > 0 ? (
                             <span className="block text-[15px] font-bold text-red-500 mb-1">{stats.lowCount} 項需叫貨</span>
                           ) : (
                             <span className="block text-[15px] font-bold text-gray-400 mb-1 flex items-center gap-1 justify-end"><Check size={16}/> 庫存充足</span>
                           )}
                           <ArrowRight size={20} className="text-gray-300 inline-block mt-1" />
                        </div>
                     </div>
                   );
                 })}
              </div>
            )}

            {activeSupplier && (
              <div className="space-y-4 animate-in slide-in-from-right-4 duration-300">
                <button 
                  onClick={() => setActiveSupplier(null)}
                  className="flex items-center gap-1.5 text-gray-500 font-bold text-[15px] mb-4 hover:text-gray-900 transition-colors bg-white px-4 py-2 rounded-full shadow-sm"
                >
                  <ChevronLeft size={20} /> 返回
                </button>
                
                <div className="bg-white rounded-[32px] p-6 shadow-[0_8px_30px_rgba(0,0,0,0.03)]">
                  <div className="flex items-center justify-between mb-8">
                    <div>
                      <h2 className="text-[24px] font-bold text-gray-900 tracking-tight">{activeSupplier}</h2>
                      <p className="text-[14px] font-medium text-gray-400 mt-1">專屬商品明細</p>
                    </div>
                  </div>

                  <div className="space-y-6">
                    {supplierItems.map(item => {
                      const currentStock = draftStocks[item.id] ?? item.current_stock;
                      const isLow = currentStock < item.min_stock;
                      return (
                        <div key={item.id} className="flex flex-col gap-3 pb-6 border-b border-gray-100 last:border-0 last:pb-0">
                          <div className="flex justify-between items-start">
                            <div>
                              <div className="flex items-center gap-2 mb-1">
                                {isLow && <span className="bg-red-50 text-red-600 text-[11px] px-2 py-0.5 rounded-[6px] font-bold flex items-center gap-1"><AlertTriangle size={12}/> 低庫存</span>}
                              </div>
                              <h2 className="text-[17px] font-bold text-gray-900 flex items-center gap-2">
                                {item.name}
                                <button onClick={() => openEditModal(item)} className="text-gray-300 hover:text-gray-600 p-1">
                                  <Pencil size={14} strokeWidth={2.5} />
                                </button>
                              </h2>
                              <p className="text-[13px] font-medium text-gray-400 mt-0.5">單價: ${item.price} • {item.unit}</p>
                            </div>
                            <div className="text-right">
                              <span className="block text-[12px] font-bold text-gray-400 mb-0.5">目前庫存</span>
                              <span className={`text-[22px] font-bold ${isLow ? 'text-red-500' : 'text-gray-900'}`}>{currentStock}</span>
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ========================================== */}
        {/* 老闆視角：決策總表 (Financial Dashboard) */}
        {/* ========================================== */}
        {activeTab === 'dashboard' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-6">
            <h2 className="text-[20px] font-bold text-gray-900 tracking-tight px-1">財務與風險總表</h2>

            <div className="bg-gray-900 rounded-[36px] p-8 text-white relative overflow-hidden shadow-2xl">
              <div className="relative z-10">
                <div className="w-14 h-14 bg-white/10 backdrop-blur-md rounded-[20px] flex items-center justify-center mb-6 border border-white/10">
                  <Wallet size={26} strokeWidth={1.5} className="text-white" />
                </div>
                <span className="text-[15px] font-medium text-gray-400 mb-1 block">目前壓倉庫存總值</span>
                <h2 className="text-[44px] font-bold leading-none tracking-tight">${totalAssetValue.toLocaleString()}</h2>
              </div>
            </div>

            <div className="bg-white rounded-[32px] p-7 shadow-[0_8px_30px_rgba(0,0,0,0.03)] flex items-center justify-between">
               <div>
                  <span className="text-[14px] font-bold text-gray-400 block mb-1">叫貨資金缺口 (補齊至安全水位)</span>
                  <span className="text-[28px] font-bold text-gray-900">${totalLowStockValue.toLocaleString()}</span>
               </div>
               <div className="w-14 h-14 bg-gray-50 rounded-full flex items-center justify-center text-gray-600">
                  <TrendingDown size={26} strokeWidth={2} />
               </div>
            </div>

            <div className="bg-white rounded-[36px] p-7 shadow-[0_8px_30px_rgba(0,0,0,0.03)]">
              <h3 className="text-[18px] font-bold text-gray-900 mb-6 flex items-center gap-2">
                <PackageOpen size={20} className="text-gray-400"/> 資金積壓分佈
              </h3>
              <div className="space-y-5">
                {dynamicCategories.filter(c => c !== '全部').map(cat => {
                  const catItems = items.filter(i => i.category === cat);
                  const catValue = catItems.reduce((sum, item) => sum + ((draftStocks[item.id] ?? item.current_stock) * (item.price || 0)), 0);
                  const percent = totalAssetValue === 0 ? 0 : Math.round((catValue / totalAssetValue) * 100);
                  
                  return (
                    <div key={cat}>
                      <div className="flex justify-between items-end mb-2">
                        <span className="text-[14px] font-bold text-gray-600">{cat}</span>
                        <span className="text-[15px] font-bold text-gray-900">${catValue.toLocaleString()} <span className="text-[13px] text-gray-400 font-medium">({percent}%)</span></span>
                      </div>
                      <div className="w-full bg-gray-100 rounded-full h-2 overflow-hidden">
                        <div className="bg-gray-800 h-2 rounded-full" style={{ width: `${percent}%` }}></div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* ========================================== */}
        {/* 設定：系統初始化 (Settings) */}
        {/* ========================================== */}
        {activeTab === 'settings' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-6">
            <h2 className="text-[20px] font-bold text-gray-900 tracking-tight px-1">資料庫維護</h2>
            
            <div className="bg-white rounded-[36px] p-8 relative overflow-hidden shadow-[0_8px_30px_rgba(0,0,0,0.03)] border border-gray-100">
              <div className="relative z-10">
                <div className="w-14 h-14 bg-gray-50 rounded-[20px] flex items-center justify-center mb-5 text-gray-600">
                  <Database size={26} strokeWidth={1.5} />
                </div>
                <h2 className="text-[22px] font-bold text-gray-900 mb-2">初始化與批次匯入</h2>
                <p className="text-[14px] text-gray-500 font-medium leading-relaxed">
                  上傳 Excel 盤點表以大批量建立庫存資料。此操作建議僅由管理員執行。
                </p>
              </div>
              
              <input type="file" accept=".xlsx, .xls" ref={fileInputRef} onChange={handleFileUpload} className="hidden" />
              
              <div className="mt-8 relative z-10">
                <button 
                  onClick={() => fileInputRef.current?.click()} 
                  disabled={isUploading}
                  className="h-14 px-8 bg-gray-900 rounded-full flex items-center gap-3 hover:bg-black active:scale-95 transition-transform disabled:opacity-50 shadow-lg"
                >
                  {isUploading ? <Loader2 size={20} className="animate-spin text-white" /> : <Upload size={20} className="text-white" />}
                  <span className="text-[15px] font-bold text-white">{isUploading ? '處理中...' : '上傳 Excel 總表'}</span>
                </button>
              </div>
            </div>
          </div>
        )}

      </div>

      {/* ========================================== */}
      {/* 盤點提交按鈕 (Batch Submit FAB) */}
      {/* 當員工修改了任何數字，這個按鈕才會浮現 */}
      {/* ========================================== */}
      {hasDraftChanges && activeTab === 'home' && (
        <div className="fixed bottom-[100px] left-1/2 -translate-x-1/2 w-[90%] max-w-[380px] z-40 animate-in slide-in-from-bottom-10 fade-in duration-300">
          <button 
            onClick={submitBatchChanges}
            disabled={isSavingBatch}
            className="w-full h-16 bg-gray-900 rounded-[24px] flex items-center justify-center gap-3 shadow-[0_20px_40px_rgba(0,0,0,0.2)] hover:bg-black active:scale-[0.98] transition-all disabled:opacity-70"
          >
            {isSavingBatch ? <Loader2 size={24} className="animate-spin text-white" /> : <Save size={24} className="text-white" />}
            <span className="text-[17px] font-bold text-white">
              {isSavingBatch ? '同步中...' : `確認送出 (${Object.keys(draftStocks).length} 項異動)`}
            </span>
          </button>
        </div>
      )}

      {/* 底部導覽列 (純淨風格) */}
      <div className="fixed bottom-6 left-1/2 -translate-x-1/2 bg-white/80 backdrop-blur-xl border border-gray-200/50 rounded-full px-6 py-4 flex justify-between items-center w-[90%] max-w-[380px] shadow-[0_20px_40px_rgba(0,0,0,0.08)] z-50">
        <button onClick={() => setActiveTab('home')} className={`relative p-2 rounded-full transition-all duration-300 flex flex-col items-center gap-1 ${activeTab === 'home' ? 'text-gray-900' : 'text-gray-400 hover:text-gray-600'}`}>
          <Home size={24} strokeWidth={activeTab === 'home' ? 2.5 : 2} />
        </button>
        <button onClick={() => {setActiveTab('suppliers'); setActiveSupplier(null);}} className={`relative p-2 rounded-full transition-all duration-300 flex flex-col items-center gap-1 ${activeTab === 'suppliers' ? 'text-gray-900' : 'text-gray-400 hover:text-gray-600'}`}>
          <Store size={24} strokeWidth={activeTab === 'suppliers' ? 2.5 : 2} />
        </button>
        <button onClick={() => setActiveTab('dashboard')} className={`relative p-2 rounded-full transition-all duration-300 flex flex-col items-center gap-1 ${activeTab === 'dashboard' ? 'text-gray-900' : 'text-gray-400 hover:text-gray-600'}`}>
          <BarChart3 size={24} strokeWidth={activeTab === 'dashboard' ? 2.5 : 2} />
        </button>
        <button onClick={() => setActiveTab('settings')} className={`relative p-2 rounded-full transition-all duration-300 flex flex-col items-center gap-1 ${activeTab === 'settings' ? 'text-gray-900' : 'text-gray-400 hover:text-gray-600'}`}>
          <Settings size={24} strokeWidth={activeTab === 'settings' ? 2.5 : 2} />
        </button>
      </div>

      {/* ========================================== */}
      {/* 編輯與新增抽屜 (Bottom Sheet) */}
      {/* ========================================== */}
      {isModalOpen && (
        <div className="fixed inset-0 z-[60] flex items-end justify-center bg-gray-900/30 backdrop-blur-sm sm:items-center">
          <div className="bg-white w-full max-w-md rounded-t-[36px] sm:rounded-[36px] p-7 pb-12 sm:pb-7 animate-in slide-in-from-bottom-full duration-300 shadow-2xl relative">
            <button onClick={() => setIsModalOpen(false)} className="absolute top-6 right-6 w-8 h-8 bg-gray-100 rounded-full flex items-center justify-center text-gray-500 hover:bg-gray-200 transition-colors">
              <X size={16} strokeWidth={3} />
            </button>
            
            <h3 className="text-[20px] font-bold text-gray-900 mb-6 flex items-center gap-2">
              {editForm.id ? <Pencil size={20} className="text-gray-700" /> : <PlusCircle size={20} className="text-gray-700" />}
              {editForm.id ? '編輯品項明細' : '建立新品項'}
            </h3>

            <div className="space-y-4">
              <div>
                <label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">品項名稱</label>
                <input type="text" value={editForm.name || ''} onChange={e => setEditForm({...editForm, name: e.target.value})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-gray-200 border border-transparent focus:border-gray-300 transition-all" placeholder="例如：蝦子" />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">分類</label>
                  <input type="text" value={editForm.category || ''} onChange={e => setEditForm({...editForm, category: e.target.value})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-gray-200" />
                </div>
                <div>
                  <label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">供應商</label>
                  <input type="text" value={editForm.supplier || ''} onChange={e => setEditForm({...editForm, supplier: e.target.value})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-gray-200" placeholder="輸入廠商名稱" />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">單位 (如：包)</label>
                  <input type="text" value={editForm.unit || ''} onChange={e => setEditForm({...editForm, unit: e.target.value})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-gray-200" />
                </div>
                <div>
                  <label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">單價 ($)</label>
                  <input type="number" min="0" value={editForm.price ?? 0} onChange={e => setEditForm({...editForm, price: Number(e.target.value)})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-gray-200" />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[13px] font-bold text-gray-900 ml-2 mb-1.5 block">目前庫存</label>
                  <input type="number" min="0" value={editForm.current_stock ?? 0} onChange={e => setEditForm({...editForm, current_stock: Number(e.target.value)})} className="w-full h-14 bg-gray-100 rounded-[20px] px-5 text-[20px] font-bold text-gray-900 outline-none focus:ring-2 ring-gray-300" />
                </div>
                <div>
                  <label className="text-[13px] font-bold text-red-500 ml-2 mb-1.5 block">安全水位</label>
                  <input type="number" min="0" value={editForm.min_stock ?? 0} onChange={e => setEditForm({...editForm, min_stock: Number(e.target.value)})} className="w-full h-14 bg-red-50 rounded-[20px] px-5 text-[20px] font-bold text-red-600 outline-none focus:ring-2 ring-red-200" />
                </div>
              </div>
            </div>

            <button 
              onClick={handleSaveItem} 
              disabled={isSavingItem}
              className="w-full h-16 mt-8 bg-gray-900 rounded-[24px] flex items-center justify-center gap-2 hover:bg-black active:scale-[0.98] transition-all disabled:opacity-70 text-white shadow-lg"
            >
              {isSavingItem ? <Loader2 size={24} className="animate-spin" /> : <Check size={24} />}
              <span className="text-[17px] font-bold">{isSavingItem ? '處理中...' : '確認儲存'}</span>
            </button>
          </div>
        </div>
      )}

      <style dangerouslySetInnerHTML={{__html: `
        .hide-scrollbar::-webkit-scrollbar { display: none; }
        .hide-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
      `}} />
    </div>
  );
}