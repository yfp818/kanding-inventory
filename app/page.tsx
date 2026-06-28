'use client';

import React, { useState, useEffect, useRef } from 'react';
import { supabase } from '../lib/supabase';
import * as XLSX from 'xlsx';
import { Toaster, toast } from 'react-hot-toast';

import { 
  Bell, Layers, Home, Settings, Loader2, CircleUserRound, Zap, 
  Plus, Minus, AlertTriangle, Search, X, Database, Store, Pencil, 
  PlusCircle, Save, CheckCircle2, ArrowRight, Upload, ClipboardList,
  ChevronLeft, BarChart3, Wallet, PackageOpen, TrendingDown, Check,
  Lock // 👈 新增了解鎖畫面的鎖頭圖示
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
  // ==========================================
  // 安全登入系統狀態 (Security Auth System)
  // ==========================================
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [isAuthError, setIsAuthError] = useState<boolean>(false);

  const [items, setItems] = useState<InventoryItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  
  const [activeTab, setActiveTab] = useState<string>('home');
  
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
      setDraftStocks({}); 
    } catch (err: unknown) {
      logSystemError('fetchInventory', extractErrorMessage(err));
      toast.error('無法載入庫存資料，請檢查網路連線。');
    } finally {
      setLoading(false);
    }
  };

  // 🛡️ 防禦機制：只有在輸入正確密碼後，才會去資料庫抓資料
  useEffect(() => { 
    if (isAuthenticated) {
      fetchInventory(); 
    }
  }, [isAuthenticated]);

  // ------------------------------------------
  // 密碼驗證邏輯
  // ------------------------------------------
  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    
    // 優先讀取 Vercel 後台的環境變數，如果沒設定，就使用預設的加密字串
    // 'MTY4ODg=' 是 '16888' 的 Base64 編碼，防止工程師直接在原始碼看到數字
    const envPin = process.env.NEXT_PUBLIC_SYSTEM_PIN;
    const isValid = envPin ? passwordInput === envPin : btoa(passwordInput) === 'MTY4ODg=';

    if (isValid) {
      setIsAuthenticated(true);
      toast.success('系統已解鎖', { icon: '🔓' });
    } else {
      setIsAuthError(true);
      toast.error('密碼錯誤');
      setPasswordInput('');
      setTimeout(() => setIsAuthError(false), 500); // 震動特效時間
    }
  };

  const handleDraftChange = (id: number, newValue: number, originalValue: number) => {
    const finalValue = Math.max(0, newValue);
    setDraftStocks(prev => {
      const newDraft = { ...prev };
      if (finalValue === originalValue) delete newDraft[id];
      else newDraft[id] = finalValue;
      return newDraft;
    });
  };

  const submitBatchChanges = async () => {
    if (!hasDraftChanges) return;
    setIsSavingBatch(true);
    const toastId = toast.loading('正在同步盤點數據至雲端...');

    try {
      const updates = Object.entries(draftStocks).map(([id, newStock]) => ({
        id: Number(id), current_stock: newStock, updated_at: new Date()
      }));

      const updatePromises = updates.map(updateData => 
        supabase.from('kanding_inventory').update({ 
          current_stock: updateData.current_stock, updated_at: updateData.updated_at 
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
    if (!process.env.NEXT_PUBLIC_SUPABASE_URL) return toast.error('請確認環境變數設定。');

    setIsUploading(true);
    const toastId = toast.loading('解析檔案中...');

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const jsonData = XLSX.utils.sheet_to_json<ExcelRow>(workbook.Sheets[workbook.SheetNames[0]], { range: 2 });
      if (jsonData.length === 0) throw new Error('解析無資料');

      const getSafeNumber = (val: unknown) => { const p = parseFloat(String(val)); return isNaN(p) ? 0 : p; };
      const formattedData = jsonData.map((row) => ({
        name: String(row['品項名稱'] ?? row['品名'] ?? row['名稱'] ?? '未命名品項'),
        category: String(row['類別'] ?? row['分類'] ?? '未分類'),
        supplier: String(row['供應商'] ?? '未指定'),
        price: getSafeNumber(row['單價']),
        current_stock: getSafeNumber(row['盤點數量'] ?? row['目前庫存']),
        min_stock: getSafeNumber(row['安全存量'] ?? row['安全水位']),
        unit: String(row['單位'] ?? row['規格'] ?? '個'),
      }));

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
      unit: '個', price: 0, current_stock: 0, min_stock: 0
    });
    setIsModalOpen(true);
  };

  const openEditModal = (item: InventoryItem) => {
    setEditForm({ ...item, current_stock: draftStocks[item.id] ?? item.current_stock });
    setIsModalOpen(true);
  };

  const handleSaveItem = async () => {
    if (!editForm.name) return toast.error('名稱不可為空');
    setIsSavingItem(true);
    const toastId = toast.loading('處理中...');

    try {
      const payload = { ...editForm, price: Number(editForm.price), current_stock: Number(editForm.current_stock), min_stock: Number(editForm.min_stock), updated_at: new Date() };
      
      if (editForm.id) {
        const { error } = await supabase.from('kanding_inventory').update(payload).eq('id', editForm.id);
        if (error) throw error;
        setDraftStocks(prev => { const d = { ...prev }; delete d[editForm.id as number]; return d; });
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
      toast.error(`失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsSavingItem(false);
    }
  };

  const homeProcessedItems = items.filter(item => {
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

  const suppliersList = Array.from(new Set(items.map(i => i.supplier || '未指定'))).sort();
  const getSupplierStats = (sup: string) => {
    const supItems = items.filter(i => i.supplier === sup);
    const lowCount = supItems.filter(i => (draftStocks[i.id] ?? i.current_stock) < i.min_stock).length;
    const totalValue = supItems.reduce((acc, curr) => acc + ((draftStocks[curr.id] ?? curr.current_stock) * (curr.price || 0)), 0);
    return { count: supItems.length, lowCount, totalValue };
  };

  // ==========================================
  // 渲染：未登入時顯示鎖定畫面 (Lock Screen)
  // ==========================================
  if (!isAuthenticated) {
    return (
      <div className="min-h-screen bg-[#F5F5F7] flex flex-col items-center justify-center px-6 selection:bg-gray-200">
         <Toaster position="top-center" toastOptions={{ className: 'rounded-[16px] text-sm font-medium shadow-lg' }} />
         
         <div className="w-20 h-20 bg-white rounded-[24px] flex items-center justify-center shadow-[0_8px_30px_rgb(0,0,0,0.04)] mb-8">
           <Lock size={32} strokeWidth={1.5} className="text-gray-900" />
         </div>
         
         <h1 className="text-[24px] font-bold text-gray-900 mb-2 tracking-tight">系統安全鎖</h1>
         <p className="text-[14px] text-gray-500 font-medium mb-10">請輸入管理員密碼以存取庫存資料</p>

         <form onSubmit={handleLogin} className="w-full max-w-[320px] space-y-4">
            <input 
              type="password" 
              inputMode="numeric"
              placeholder="輸入密碼"
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              className={`w-full h-16 bg-white rounded-[20px] px-6 text-center text-[24px] tracking-[0.5em] font-bold outline-none transition-all shadow-sm ${isAuthError ? 'border-2 border-red-500 animate-shake' : 'border border-transparent focus:ring-2 ring-gray-200'}`}
              autoFocus
            />
            <button 
              type="submit" 
              className="w-full h-16 bg-gray-900 rounded-[20px] flex items-center justify-center text-white font-bold text-[17px] active:scale-[0.98] transition-transform shadow-lg hover:bg-black"
            >
              解鎖系統
            </button>
         </form>

         {/* 注入錯誤震動特效 */}
         <style dangerouslySetInnerHTML={{__html: `
           @keyframes shake {
             0%, 100% { transform: translateX(0); }
             25% { transform: translateX(-8px); }
             75% { transform: translateX(8px); }
           }
           .animate-shake { animation: shake 0.2s ease-in-out 0s 2; }
         `}} />
      </div>
    );
  }

  // ==========================================
  // 渲染：已登入顯示主系統 (Main App)
  // ==========================================
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
              {activeTab === 'home' ? '現場盤點模式' : activeTab === 'dashboard' ? '財務決策模式' : '廠商與系統維護'}
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
        {/* 員工視角：日常盤點與查貨 */}
        {/* ========================================== */}
        {activeTab === 'home' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-6">
            
            {/* 智慧過濾卡片 */}
            <div className="grid grid-cols-2 gap-4">
              <div 
                onClick={() => setHomeFilter('all')}
                className={`rounded-[32px] p-6 shadow-[0_8px_24px_rgba(0,0,0,0.03)] flex flex-col items-start relative overflow-hidden transition-all duration-300 cursor-pointer ${homeFilter === 'all' ? 'bg-[#292A32] text-white scale-105' : 'bg-white'}`}
              >
                <div className={`w-[52px] h-[52px] rounded-[20px] rounded-bl-[8px] flex items-center justify-center mb-4 transition-colors ${homeFilter === 'all' ? 'bg-[#3F404A] text-[#FFAD7A]' : 'bg-[#FFAD7A] text-white'}`}>
                  <ClipboardList size={26} strokeWidth={2} />
                </div>
                <span className={`text-[14px] font-bold mb-1 ${homeFilter === 'all' ? 'text-gray-300' : 'text-gray-400'}`}>監控品項</span>
                <span className={`text-[32px] font-bold leading-none ${homeFilter === 'all' ? 'text-white' : 'text-gray-900'}`}>{totalItems}</span>
              </div>
              
              <div 
                onClick={() => setHomeFilter('low')}
                className={`rounded-[32px] p-6 shadow-[0_8px_24px_rgba(0,0,0,0.03)] flex flex-col items-start relative overflow-hidden transition-all duration-300 cursor-pointer ${homeFilter === 'low' ? 'bg-[#FF6B6B] text-white scale-105 shadow-[0_12px_30px_rgba(255,107,107,0.3)]' : 'bg-white'}`}
              >
                <div className={`w-[52px] h-[52px] rounded-[20px] rounded-bl-[8px] flex items-center justify-center mb-4 transition-colors ${homeFilter === 'low' ? 'bg-white/20 text-white' : 'bg-[#FF6B6B]/10 text-[#FF6B6B]'}`}>
                  <AlertTriangle size={26} strokeWidth={2} />
                </div>
                <span className={`text-[14px] font-bold mb-1 ${homeFilter === 'low' ? 'text-white/80' : 'text-gray-400'}`}>急需叫貨</span>
                <span className={`text-[32px] font-bold leading-none ${homeFilter === 'low' ? 'text-white' : 'text-gray-900'}`}>{lowStockItems}</span>
              </div>
            </div>

            <div className="bg-white rounded-[36px] p-6 shadow-[0_8px_24px_rgba(0,0,0,0.03)]">
              <div className="flex items-center justify-between mb-5">
                <h3 className="text-[18px] font-bold text-gray-900 tracking-tight">
                  {homeFilter === 'low' ? '🚨 待叫貨清單' : '盤點與搜尋'}
                </h3>
              </div>

              {/* 搜尋列 */}
              <div className="relative flex items-center w-full h-[52px] bg-[#F4F5F9] rounded-full px-5 mb-5 transition-colors focus-within:ring-2 ring-[#8780F2]/30">
                <Search size={20} className="text-gray-400" />
                <input 
                  type="text" 
                  placeholder="搜尋或新增品項..." 
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full h-full bg-transparent outline-none px-3 text-[15px] font-medium text-gray-900 placeholder:text-gray-400"
                />
                {searchQuery && (
                  <button onClick={() => setSearchQuery('')} className="p-1.5 bg-gray-200 rounded-full text-gray-500 hover:bg-gray-300">
                    <X size={14} strokeWidth={3} />
                  </button>
                )}
              </div>

              {/* 搜尋即新增按鈕 */}
              {searchQuery && homeProcessedItems.length === 0 && (
                <div 
                  onClick={() => openAddModal({ name: searchQuery })}
                  className="mb-5 bg-[#8780F2]/10 border border-[#8780F2]/20 rounded-[20px] p-4 flex items-center justify-between cursor-pointer hover:bg-[#8780F2]/20 transition-colors"
                >
                  <div className="flex items-center gap-3">
                    <div className="w-10 h-10 bg-[#8780F2] text-white rounded-full flex items-center justify-center">
                      <PlusCircle size={20} />
                    </div>
                    <div>
                      <h4 className="text-[15px] font-bold text-[#8780F2]">找不到 "{searchQuery}"</h4>
                      <p className="text-[12px] font-medium text-[#8780F2]/70">點擊立即新增此品項</p>
                    </div>
                  </div>
                  <ArrowRight size={20} className="text-[#8780F2]" />
                </div>
              )}

              {/* 橫向膠囊動態分類 */}
              <div className="flex gap-2 overflow-x-auto hide-scrollbar mb-6">
                {dynamicCategories.map(cat => (
                  <button 
                    key={cat} 
                    onClick={() => setSelectedCategory(cat)} 
                    className={`px-5 py-2.5 rounded-full text-[14px] font-bold whitespace-nowrap transition-all duration-300 ${
                      selectedCategory === cat 
                        ? 'bg-[#FFAD7A] text-white shadow-[0_6px_16px_rgba(255,173,122,0.4)]' 
                        : 'bg-[#F4F5F9] text-gray-500 hover:bg-gray-100'
                    }`}
                  >
                    {cat}
                  </button>
                ))}
              </div>

              <div className="space-y-4">
                {loading ? (
                   <div className="py-10 flex justify-center"><Loader2 className="animate-spin text-[#8780F2]" size={32} /></div>
                ) : homeProcessedItems.length === 0 ? (
                   <div className="py-10 text-center text-gray-400 font-medium text-[14px]">
                     {homeFilter === 'low' ? '目前沒有缺貨品項！ 🎉' : '無品項資料'}
                   </div>
                ) : (
                  homeProcessedItems.map(item => {
                    const currentStock = draftStocks[item.id] ?? item.current_stock;
                    const isLow = currentStock < item.min_stock;
                    const isModified = draftStocks[item.id] !== undefined;

                    return (
                      <div key={item.id} className={`bg-white rounded-[24px] p-5 transition-all duration-300 border ${isModified ? 'border-[#8780F2] shadow-md' : 'border-transparent shadow-[0_4px_20px_rgb(0,0,0,0.03)]'}`}>
                        <div className="flex-1 pr-4 mb-4">
                          <div className="flex items-center gap-2 mb-1">
                             {isLow && <span className="flex items-center gap-1 bg-[#FF6B6B]/10 text-[#FF6B6B] text-[11px] px-2 py-0.5 rounded-[6px] font-bold"><AlertTriangle size={12}/> 需叫貨</span>}
                             {isModified && <span className="bg-gray-100 text-gray-700 text-[11px] px-2 py-0.5 rounded-[8px] font-bold">待儲存</span>}
                          </div>
                          <div className="flex items-center gap-2">
                            <h2 className="text-[16px] font-bold text-gray-900 leading-tight">{item.name}</h2>
                            <button onClick={() => openEditModal(item)} className="text-gray-300 hover:text-[#8780F2] transition-colors p-1">
                              <Pencil size={14} strokeWidth={2.5} />
                            </button>
                          </div>
                          <p className="text-[12px] font-semibold text-gray-400 mt-0.5">
                            {item.supplier} • 水位: {item.min_stock}{item.unit}
                          </p>
                        </div>
                        
                        <div className="bg-[#F4F5F9] rounded-[18px] p-1.5 flex items-center justify-between">
                          <button onClick={() => handleDraftChange(item.id, currentStock - 1, item.current_stock)} className="w-12 h-12 rounded-[14px] bg-white text-gray-600 flex items-center justify-center shadow-sm active:scale-90 transition-transform">
                            <Minus size={20} strokeWidth={2.5} />
                          </button>
                          <span className={`text-[22px] font-bold ${isModified ? 'text-[#8780F2]' : 'text-gray-900'}`}>{currentStock}</span>
                          <button onClick={() => handleDraftChange(item.id, currentStock + 1, item.current_stock)} className="w-12 h-12 rounded-[14px] bg-white text-gray-600 flex items-center justify-center shadow-sm active:scale-90 transition-transform">
                            <Plus size={20} strokeWidth={2.5} />
                          </button>
                        </div>
                      </div>
                    );
                  })
                )}
              </div>
            </div>
          </div>
        )}

        {/* ========================================== */}
        {/* 供應商視角：進銷存矩陣 */}
        {/* ========================================== */}
        {activeTab === 'suppliers' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
            {!activeSupplier && (
              <div className="space-y-4">
                 <div className="flex items-center justify-between px-2 mb-2">
                    <h2 className="text-[20px] font-bold text-gray-900 tracking-tight">供應商管理</h2>
                    <button onClick={() => openAddModal()} className="flex items-center gap-1.5 text-[13px] font-bold text-[#8780F2] bg-[#8780F2]/10 px-4 py-2 rounded-full hover:bg-[#8780F2]/20 transition-colors">
                      <Plus size={16} /> 新增廠商/品項
                    </button>
                 </div>
                 
                 {suppliersList.map(sup => {
                   const supItems = items.filter(i => i.supplier === sup);
                   const lowCount = supItems.filter(i => (draftStocks[i.id] ?? i.current_stock) < i.min_stock).length;
                   return (
                     <div key={sup} onClick={() => setActiveSupplier(sup)} className="bg-white rounded-[28px] p-5 shadow-[0_8px_24px_rgba(0,0,0,0.03)] border border-gray-50 flex items-center justify-between cursor-pointer hover:shadow-md transition-all active:scale-[0.98]">
                        <div className="flex items-center gap-4">
                          <div className={`w-12 h-12 rounded-[16px] flex items-center justify-center ${lowCount > 0 ? 'bg-[#FF6B6B]/10 text-[#FF6B6B]' : 'bg-[#F4F5F9] text-gray-500'}`}>
                            <Store size={24} strokeWidth={2} />
                          </div>
                          <div>
                            <h3 className="text-[17px] font-bold text-gray-900">{sup}</h3>
                            <p className="text-[13px] font-medium text-gray-400 mt-0.5">供應 {supItems.length} 項商品</p>
                          </div>
                        </div>
                        <div className="text-right">
                           {lowCount > 0 ? <span className="block text-[14px] font-bold text-[#FF6B6B] mb-1">{lowCount} 項缺貨</span> : <span className="block text-[14px] font-bold text-emerald-500 mb-1 flex items-center gap-1"><CheckCircle2 size={14}/> 庫存充足</span>}
                           <ArrowRight size={18} className="text-gray-300 mt-1" />
                        </div>
                     </div>
                   );
                 })}
              </div>
            )}

            {activeSupplier && (
              <div className="space-y-4 animate-in slide-in-from-right-4 duration-300">
                <button onClick={() => setActiveSupplier(null)} className="flex items-center gap-1 text-gray-500 font-bold text-[14px] mb-2 hover:text-gray-800 transition-colors">
                  <ChevronLeft size={20} /> 返回廠商列表
                </button>
                <div className="bg-white rounded-[32px] p-6 shadow-[0_8px_24px_rgba(0,0,0,0.03)] border border-gray-50">
                  <div className="flex items-center justify-between mb-6">
                    <div>
                      <h2 className="text-[22px] font-bold text-gray-900 tracking-tight">{activeSupplier}</h2>
                      <p className="text-[13px] font-medium text-gray-400 mt-1">廠商貨品盤點清單</p>
                    </div>
                    <button onClick={() => openAddModal({ supplier: activeSupplier })} className="w-10 h-10 bg-[#8780F2] text-white rounded-full flex items-center justify-center shadow-lg hover:bg-indigo-600 transition-colors">
                      <Plus size={20} strokeWidth={2.5} />
                    </button>
                  </div>

                  <div className="space-y-4">
                    {items.filter(item => item.supplier === activeSupplier).map(item => {
                        const currentStock = draftStocks[item.id] ?? item.current_stock;
                        const isLow = currentStock < item.min_stock;
                        return (
                          <div key={item.id} className="flex items-center justify-between pb-4 border-b border-gray-100 last:border-0 last:pb-0">
                            <div className="flex-1 pr-4">
                              <div className="flex items-center gap-2 mb-1">
                                {isLow && <span className="flex items-center gap-1 bg-[#FF6B6B]/10 text-[#FF6B6B] text-[11px] px-2 py-0.5 rounded-[6px] font-bold"><AlertTriangle size={12}/> 需叫貨</span>}
                              </div>
                              <div className="flex items-center gap-2">
                                <h2 className="text-[16px] font-bold text-gray-900 leading-tight">{item.name}</h2>
                                <button onClick={() => openEditModal(item)} className="text-gray-300 hover:text-[#8780F2] transition-colors p-1"><Pencil size={14} strokeWidth={2.5} /></button>
                              </div>
                              <p className="text-[12px] font-semibold text-gray-400 mt-0.5">單價: ${item.price} • 分類: {item.category}</p>
                            </div>
                            
                            <div className="flex items-center gap-3">
                              <button onClick={() => handleDraftChange(item.id, currentStock - 1, item.current_stock)} className="w-9 h-9 rounded-full bg-[#F4F5F9] text-gray-600 flex items-center justify-center active:scale-90 transition-transform"><Minus size={16} strokeWidth={2.5} /></button>
                              <span className="w-6 text-center text-[18px] font-bold text-gray-900">{currentStock}</span>
                              <button onClick={() => handleDraftChange(item.id, currentStock + 1, item.current_stock)} className="w-9 h-9 rounded-full bg-[#292A32] text-white flex items-center justify-center active:scale-90 transition-transform"><Plus size={16} strokeWidth={2.5} /></button>
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
        {/* 老闆視角：決策總覽 */}
        {/* ========================================== */}
        {activeTab === 'dashboard' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-6">
            <h2 className="text-[20px] font-bold text-gray-900 tracking-tight px-1">財務與營運總值</h2>

            <div className="bg-[#292A32] rounded-[36px] p-7 text-white relative overflow-hidden shadow-2xl">
              <div className="relative z-10">
                <div className="w-12 h-12 bg-white/10 backdrop-blur-md rounded-[16px] flex items-center justify-center mb-6">
                  <Wallet size={24} strokeWidth={2} className="text-[#FFAD7A]" />
                </div>
                <span className="text-[14px] font-medium text-white/60 mb-1 block">目前壓倉庫存總資產</span>
                <h2 className="text-[40px] font-bold leading-none tracking-tight">${totalAssetValue.toLocaleString()}</h2>
              </div>
              <div className="absolute right-0 bottom-0 opacity-10 pointer-events-none transform translate-x-4 translate-y-4">
                 <BarChart3 size={160} strokeWidth={1} />
              </div>
            </div>

            <div className="bg-white rounded-[32px] p-6 shadow-[0_8px_24px_rgba(0,0,0,0.03)] flex items-center justify-between border border-gray-50">
               <div>
                  <span className="text-[13px] font-bold text-gray-400 block mb-1">缺貨補齊金流缺口 (預估成本)</span>
                  <span className="text-[24px] font-bold text-[#FF6B6B]">${totalLowStockValue.toLocaleString()}</span>
               </div>
               <div className="w-12 h-12 bg-[#FF6B6B]/10 rounded-full flex items-center justify-center text-[#FF6B6B]">
                  <TrendingDown size={24} strokeWidth={2.5} />
               </div>
            </div>

            <div className="bg-white rounded-[36px] p-6 shadow-[0_8px_24px_rgba(0,0,0,0.03)] border border-gray-50">
              <h3 className="text-[17px] font-bold text-gray-900 mb-5 flex items-center gap-2">
                <PackageOpen size={18} className="text-[#8780F2]"/> 依品項分類資產比例
              </h3>
              <div className="space-y-4">
                {dynamicCategories.filter(c => c !== '全部').map(cat => {
                  const catItems = items.filter(i => i.category === cat);
                  const catValue = catItems.reduce((sum, item) => sum + ((draftStocks[item.id] ?? item.current_stock) * (item.price || 0)), 0);
                  const percent = totalAssetValue === 0 ? 0 : Math.round((catValue / totalAssetValue) * 100);
                  
                  return (
                    <div key={cat} className="mb-2">
                      <div className="flex justify-between items-end mb-1.5">
                        <span className="text-[14px] font-bold text-gray-700">{cat}</span>
                        <span className="text-[14px] font-bold text-gray-900">${catValue.toLocaleString()} <span className="text-[12px] text-gray-400 font-medium">({percent}%)</span></span>
                      </div>
                      <div className="w-full bg-[#F4F5F9] rounded-full h-2.5 overflow-hidden">
                        <div className="bg-[#8780F2] h-2.5 rounded-full" style={{ width: `${percent}%` }}></div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* ========================================== */}
        {/* 後台批次維護 (Settings) */}
        {/* ========================================== */}
        {activeTab === 'settings' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-6">
            <h2 className="text-[20px] font-bold text-gray-900 tracking-tight px-1">資料庫維護</h2>
            
            <div className="bg-[#8780F2] rounded-[36px] p-8 text-white relative overflow-hidden shadow-[0_16px_32px_-12px_rgba(135,128,242,0.6)]">
              <div className="relative z-10 w-3/4">
                <div className="w-12 h-12 bg-white/20 backdrop-blur-md rounded-[16px] flex items-center justify-center mb-4">
                  <Database size={24} strokeWidth={2} className="text-white" />
                </div>
                <h2 className="text-[24px] font-bold leading-tight mb-2">批次匯入總表</h2>
                <p className="text-[13px] text-white/90 leading-relaxed font-medium">
                  上傳最新的 Excel 盤點表以大批量更新或初始化庫存。
                </p>
              </div>
              
              <input type="file" accept=".xlsx, .xls" ref={fileInputRef} onChange={handleFileUpload} className="hidden" />
              
              <div className="mt-8 relative z-10">
                <button 
                  onClick={() => fileInputRef.current?.click()} 
                  disabled={isUploading}
                  className="h-14 px-6 bg-[#292A32] rounded-full flex items-center gap-3 hover:scale-105 active:scale-95 transition-transform disabled:opacity-50 shadow-xl"
                >
                  {isUploading ? <Loader2 size={20} className="animate-spin text-white" /> : <Upload size={20} className="text-white" />}
                  <span className="text-[15px] font-bold text-white">{isUploading ? '雲端寫入中...' : '選擇 Excel 檔案'}</span>
                </button>
              </div>
              <div className="absolute -right-6 -bottom-6 opacity-10 pointer-events-none">
                 <Upload size={220} strokeWidth={1.5} />
              </div>
            </div>
          </div>
        )}

      </div>

      {/* ========================================== */}
      {/* 盤點提交按鈕 (Batch Submit FAB) */}
      {/* ========================================== */}
      {hasDraftChanges && activeTab === 'home' && (
        <div className="fixed bottom-[100px] left-1/2 -translate-x-1/2 w-[90%] max-w-[380px] z-40 animate-in slide-in-from-bottom-10 fade-in duration-300">
          <button 
            onClick={submitBatchChanges}
            disabled={isSavingBatch}
            className="w-full h-16 bg-[#292A32] rounded-[24px] flex items-center justify-center gap-3 shadow-[0_20px_40px_rgba(0,0,0,0.3)] hover:bg-black active:scale-[0.98] transition-all disabled:opacity-70"
          >
            {isSavingBatch ? <Loader2 size={24} className="animate-spin text-white" /> : <Save size={24} className="text-white" />}
            <span className="text-[17px] font-bold text-white">
              {isSavingBatch ? '同步中...' : `確認送出 (${Object.keys(draftStocks).length} 項異動)`}
            </span>
          </button>
        </div>
      )}

      {/* ========================================== */}
      {/* 發光漸層底部導覽列 */}
      {/* ========================================== */}
      <div className="fixed bottom-8 left-1/2 -translate-x-1/2 bg-[#292A32] rounded-full px-6 py-4 flex justify-between items-center w-[90%] max-w-[380px] shadow-[0_24px_40px_-12px_rgba(0,0,0,0.5)] z-40">
        <button onClick={() => setActiveTab('home')} className={`relative p-2 rounded-full transition-all duration-300 ${activeTab === 'home' ? 'text-white' : 'text-gray-500 hover:text-gray-300'}`}>
          {activeTab === 'home' && <span className="absolute inset-0 bg-[#8780F2] rounded-full -z-10 shadow-[0_0_15px_rgba(135,128,242,0.5)]"></span>}
          <Home size={24} strokeWidth={2.5} />
        </button>
        <button onClick={() => {setActiveTab('suppliers'); setActiveSupplier(null);}} className={`relative p-2 rounded-full transition-all duration-300 ${activeTab === 'suppliers' ? 'text-white' : 'text-gray-500 hover:text-gray-300'}`}>
          {activeTab === 'suppliers' && <span className="absolute inset-0 bg-[#FFAD7A] rounded-full -z-10 shadow-[0_0_15px_rgba(255,173,122,0.5)]"></span>}
          <Store size={24} strokeWidth={2.5} />
        </button>
        <button onClick={() => setActiveTab('dashboard')} className={`relative p-2 rounded-full transition-all duration-300 ${activeTab === 'dashboard' ? 'text-white' : 'text-gray-500 hover:text-gray-300'}`}>
          {activeTab === 'dashboard' && <span className="absolute inset-0 bg-[#4ADE80] rounded-full -z-10 shadow-[0_0_15px_rgba(74,222,128,0.4)]"></span>}
          <BarChart3 size={24} strokeWidth={2.5} />
        </button>
        <button onClick={() => setActiveTab('settings')} className={`relative p-2 rounded-full transition-all duration-300 ${activeTab === 'settings' ? 'text-white' : 'text-gray-500 hover:text-gray-300'}`}>
          {activeTab === 'settings' && <span className="absolute inset-0 bg-gray-600 rounded-full -z-10 shadow-[0_0_15px_rgba(75,85,99,0.5)]"></span>}
          <Settings size={24} strokeWidth={2.5} />
        </button>
      </div>

      {/* ========================================== */}
      {/* 編輯與新增抽屜 Modal (Bottom Sheet) */}
      {/* ========================================== */}
      {isModalOpen && (
        <div className="fixed inset-0 z-[60] flex items-end justify-center bg-[#292A32]/40 backdrop-blur-sm sm:items-center">
          <div className="bg-white w-full max-w-md rounded-t-[36px] sm:rounded-[36px] p-7 pb-12 sm:pb-7 animate-in slide-in-from-bottom-full duration-300 shadow-2xl relative">
            <button onClick={() => setIsModalOpen(false)} className="absolute top-6 right-6 w-8 h-8 bg-gray-100 rounded-full flex items-center justify-center text-gray-500 hover:bg-gray-200 transition-colors">
              <X size={16} strokeWidth={3} />
            </button>
            
            <h3 className="text-[20px] font-bold text-gray-900 mb-6 flex items-center gap-2">
              {editForm.id ? <Pencil size={20} className="text-[#8780F2]" /> : <PlusCircle size={20} className="text-[#FFAD7A]" />}
              {editForm.id ? '編輯品項明細' : '建立新品項'}
            </h3>

            <div className="space-y-4">
              <div>
                <label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">品項名稱</label>
                <input type="text" value={editForm.name || ''} onChange={e => setEditForm({...editForm, name: e.target.value})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" placeholder="例如：蝦子" />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">分類</label>
                  <input type="text" value={editForm.category || ''} onChange={e => setEditForm({...editForm, category: e.target.value})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" />
                </div>
                <div>
                  <label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">供應商</label>
                  <input type="text" value={editForm.supplier || ''} onChange={e => setEditForm({...editForm, supplier: e.target.value})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" placeholder="輸入廠商名稱" />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">單位 (如：包、個)</label>
                  <input type="text" value={editForm.unit || ''} onChange={e => setEditForm({...editForm, unit: e.target.value})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" />
                </div>
                <div>
                  <label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">單價 ($)</label>
                  <input type="number" min="0" value={editForm.price ?? 0} onChange={e => setEditForm({...editForm, price: Number(e.target.value)})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[12px] font-bold text-[#8780F2] ml-1 mb-1 block">目前庫存</label>
                  <input type="number" min="0" value={editForm.current_stock ?? 0} onChange={e => setEditForm({...editForm, current_stock: Number(e.target.value)})} className="w-full h-12 bg-[#8780F2]/10 text-[#8780F2] rounded-[16px] px-4 text-[18px] font-bold outline-none focus:ring-2 ring-[#8780F2]/50" />
                </div>
                <div>
                  <label className="text-[12px] font-bold text-[#FF6B6B] ml-1 mb-1 block">安全水位</label>
                  <input type="number" min="0" value={editForm.min_stock ?? 0} onChange={e => setEditForm({...editForm, min_stock: Number(e.target.value)})} className="w-full h-12 bg-[#FF6B6B]/10 text-[#FF6B6B] rounded-[16px] px-4 text-[18px] font-bold outline-none focus:ring-2 ring-[#FF6B6B]/50" />
                </div>
              </div>
            </div>

            <button 
              onClick={handleSaveItem} 
              disabled={isSavingItem}
              className="w-full h-14 mt-8 bg-[#292A32] rounded-[20px] flex items-center justify-center gap-2 hover:bg-black active:scale-95 transition-transform disabled:opacity-50 text-white shadow-xl"
            >
              {isSavingItem ? <Loader2 size={20} className="animate-spin" /> : <Save size={20} />}
              <span className="text-[16px] font-bold">{isSavingItem ? '處理中...' : '儲存變更'}</span>
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