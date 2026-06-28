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
    return '網路連線異常。請確認 .env.local 內的網址正確。';
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
  
  // 視角切換：'home' | 'suppliers' | 'dashboard' | 'settings'
  const [activeTab, setActiveTab] = useState<string>('home');
  
  // 批次盤點草稿系統
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

  const fetchInventory = async () => {
    setLoading(true);
    try {
      const { data, error } = await supabase.from('kanding_inventory').select('*').order('id', { ascending: true });
      if (error) throw error;
      if (data) setItems(data as InventoryItem[]);
      setDraftStocks({});
    } catch (err: unknown) {
      toast.error('無法載入庫存資料');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { fetchInventory(); }, []);

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
    const toastId = toast.loading('同步至雲端...');

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
      toast.error(`同步失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsSavingBatch(false);
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setIsUploading(true);
    const toastId = toast.loading('解析檔案中...');

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const jsonData = XLSX.utils.sheet_to_json<ExcelRow>(workbook.Sheets[workbook.SheetNames[0]], { range: 2 });
      if (jsonData.length === 0) throw new Error('檔案為空');

      const getSafeNum = (val: unknown) => { const p = parseFloat(String(val)); return isNaN(p) ? 0 : p; };
      const formattedData = jsonData.map((row) => ({
        name: String(row['品項名稱'] ?? row['品名'] ?? row['名稱'] ?? '未命名'),
        category: String(row['類別'] ?? row['分類'] ?? '未分類'),
        supplier: String(row['供應商'] ?? '未指定'),
        price: getSafeNum(row['單價']),
        current_stock: getSafeNum(row['盤點數量'] ?? row['目前庫存']),
        min_stock: getSafeNum(row['安全存量'] ?? row['安全水位']),
        unit: String(row['單位'] ?? row['規格'] ?? '個'),
      }));

      const { error } = await supabase.from('kanding_inventory').insert(formattedData);
      if (error) throw error;
      
      toast.success(`匯入完成！共 ${formattedData.length} 筆`, { id: toastId });
      fetchInventory(); setActiveTab('home');
    } catch (err: unknown) {
      toast.error(`處理失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const openAddModal = (presets?: { name?: string, supplier?: string }) => {
    setEditForm({ name: presets?.name || '', category: selectedCategory === '全部' ? '未分類' : selectedCategory, supplier: presets?.supplier || '未指定', unit: '個', price: 0, current_stock: 0, min_stock: 0 });
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
      const payload = { ...editForm, updated_at: new Date() };
      if (editForm.id) {
        const { error } = await supabase.from('kanding_inventory').update(payload).eq('id', editForm.id);
        if (error) throw error;
        setDraftStocks(prev => { const d = {...prev}; delete d[editForm.id as number]; return d; });
        toast.success('更新成功', { id: toastId });
      } else {
        const { error } = await supabase.from('kanding_inventory').insert([payload]);
        if (error) throw error;
        toast.success('新增成功', { id: toastId });
        setSearchQuery(''); 
      }
      await fetchInventory(); setIsModalOpen(false);
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
    const aIsLow = (draftStocks[a.id] ?? a.current_stock) < a.min_stock;
    const bIsLow = (draftStocks[b.id] ?? b.current_stock) < b.min_stock;
    if (aIsLow && !bIsLow) return -1;
    if (!aIsLow && bIsLow) return 1;
    if (a.category !== b.category) return a.category.localeCompare(b.category, 'zh-TW');
    return a.name.localeCompare(b.name, 'zh-TW');
  });

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
              {activeTab === 'home' ? '現場盤點作業' : activeTab === 'dashboard' ? '財務決策報表' : '系統後台維護'}
            </p>
          </div>
        </div>
        <button onClick={fetchInventory} disabled={loading || isSavingBatch} className="w-[48px] h-[48px] bg-white rounded-full flex items-center justify-center shadow-[0_4px_20px_rgba(0,0,0,0.04)] active:scale-95 transition-transform">
          <Zap size={22} strokeWidth={1.5} className={loading ? "animate-pulse text-gray-400" : ""} />
        </button>
      </div>

      <div className="px-6 space-y-6">
        
        {/* 員工視角：首頁盤點 */}
        {activeTab === 'home' && (
          <div className="animate-in fade-in space-y-6">
            <div className="relative flex items-center w-full h-[56px] bg-white rounded-[20px] shadow-sm px-5 focus-within:ring-2 ring-gray-200">
              <Search size={20} className="text-gray-400" />
              <input type="text" placeholder="搜尋或新增品項..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full h-full bg-transparent outline-none px-3 text-[15px] font-medium text-gray-900" />
              {searchQuery && <button onClick={() => setSearchQuery('')} className="p-1.5 bg-gray-100 rounded-full text-gray-500"><X size={14} /></button>}
            </div>

            {searchQuery && homeProcessedItems.length === 0 && (
              <div onClick={() => openAddModal({ name: searchQuery })} className="bg-white border border-gray-100 rounded-[24px] p-5 flex items-center justify-between cursor-pointer shadow-sm">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 bg-gray-50 text-gray-600 rounded-full flex items-center justify-center"><PlusCircle size={24} /></div>
                  <div><h4 className="text-[16px] font-bold">建立 "{searchQuery}"</h4><p className="text-[13px] text-gray-400">點擊新增至資料庫</p></div>
                </div>
                <ArrowRight size={20} className="text-gray-400" />
              </div>
            )}

            <div className="flex gap-2 overflow-x-auto hide-scrollbar">
              {dynamicCategories.map(cat => (
                <button key={cat} onClick={() => setSelectedCategory(cat)} className={`px-5 py-3 rounded-[16px] text-[14px] font-bold whitespace-nowrap transition-all ${selectedCategory === cat ? 'bg-gray-900 text-white shadow-lg' : 'bg-white text-gray-500 shadow-sm'}`}>{cat}</button>
              ))}
            </div>

            <div className="space-y-4">
              {loading ? <div className="py-20 flex justify-center"><Loader2 className="animate-spin text-gray-300" size={32} /></div> : homeProcessedItems.map(item => {
                const currentStock = draftStocks[item.id] ?? item.current_stock;
                const isLow = currentStock < item.min_stock;
                const isModified = draftStocks[item.id] !== undefined;

                return (
                  <div key={item.id} className={`bg-white rounded-[24px] p-5 transition-all ${isModified ? 'border border-gray-900 shadow-md' : 'shadow-sm border border-transparent'}`}>
                    <div className="flex justify-between mb-4">
                      <div>
                        <div className="flex gap-2 mb-1.5">
                           {isLow && <span className="bg-red-50 text-red-600 text-[11px] px-2 py-0.5 rounded-[8px] font-bold">低庫存</span>}
                           {isModified && <span className="bg-gray-100 text-gray-700 text-[11px] px-2 py-0.5 rounded-[8px] font-bold">待儲存</span>}
                        </div>
                        <h2 className="text-[17px] font-bold text-gray-900 flex items-center gap-2">
                          {item.name}
                          <button onClick={() => openEditModal(item)} className="text-gray-300 hover:text-gray-600 p-1"><Pencil size={14} /></button>
                        </h2>
                        <p className="text-[13px] text-gray-400 mt-1">{item.category} • 水位: {item.min_stock}</p>
                      </div>
                    </div>
                    <div className="bg-[#F5F5F7] rounded-[18px] p-1.5 flex justify-between items-center">
                      <button onClick={() => handleDraftChange(item.id, currentStock - 1, item.current_stock)} className="w-12 h-12 rounded-[14px] bg-white shadow-sm flex items-center justify-center active:scale-90"><Minus size={20} /></button>
                      <span className={`text-[22px] font-bold ${isModified ? 'text-gray-900' : 'text-gray-600'}`}>{currentStock}</span>
                      <button onClick={() => handleDraftChange(item.id, currentStock + 1, item.current_stock)} className="w-12 h-12 rounded-[14px] bg-white shadow-sm flex items-center justify-center active:scale-90"><Plus size={20} /></button>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {/* 供應商視角 */}
        {activeTab === 'suppliers' && (
          <div className="animate-in fade-in space-y-4">
            {!activeSupplier ? Array.from(new Set(items.map(i => i.supplier || '未指定'))).map(sup => {
              const supItems = items.filter(i => i.supplier === sup);
              const lowCount = supItems.filter(i => (draftStocks[i.id] ?? i.current_stock) < i.min_stock).length;
              return (
                <div key={sup} onClick={() => setActiveSupplier(sup)} className="bg-white rounded-[28px] p-6 shadow-sm flex justify-between cursor-pointer active:scale-95">
                  <div className="flex gap-4 items-center">
                    <div className={`w-14 h-14 rounded-[20px] flex justify-center items-center ${lowCount > 0 ? 'bg-red-50 text-red-500' : 'bg-gray-50 text-gray-500'}`}><Store size={26} /></div>
                    <div><h3 className="text-[18px] font-bold">{sup}</h3><p className="text-[13px] text-gray-400">共 {supItems.length} 項商品</p></div>
                  </div>
                  <div className="text-right">
                    {lowCount > 0 ? <span className="text-[15px] font-bold text-red-500 block">{lowCount} 項需叫貨</span> : <span className="text-[15px] font-bold text-emerald-500 block flex items-center gap-1 justify-end"><CheckCircle2 size={16}/> 充足</span>}
                    <ArrowRight size={20} className="text-gray-300 mt-1 inline-block" />
                  </div>
                </div>
              );
            }) : (
              <div>
                <button onClick={() => setActiveSupplier(null)} className="flex items-center gap-1.5 text-gray-500 font-bold mb-4 bg-white px-4 py-2 rounded-full shadow-sm"><ChevronLeft size={20} /> 返回</button>
                <div className="bg-white rounded-[32px] p-6 shadow-sm space-y-6">
                  <div className="flex justify-between"><h2 className="text-[24px] font-bold">{activeSupplier}</h2><button onClick={() => openAddModal({ supplier: activeSupplier })} className="w-10 h-10 bg-gray-900 text-white rounded-full flex justify-center items-center"><Plus size={20} /></button></div>
                  {items.filter(i => i.supplier === activeSupplier).map(item => {
                    const currentStock = draftStocks[item.id] ?? item.current_stock;
                    return (
                      <div key={item.id} className="flex justify-between items-center pb-4 border-b border-gray-100 last:border-0 last:pb-0">
                        <div>
                          <h2 className="text-[16px] font-bold flex gap-2">{item.name} <button onClick={() => openEditModal(item)} className="text-gray-300"><Pencil size={14} /></button></h2>
                          <p className="text-[13px] text-gray-400 mt-0.5">單價: ${item.price} • {item.unit}</p>
                        </div>
                        <div className="text-right"><span className="text-[12px] text-gray-400 block">目前庫存</span><span className={`text-[20px] font-bold ${currentStock < item.min_stock ? 'text-red-500' : 'text-gray-900'}`}>{currentStock}</span></div>
                      </div>
                    );
                  })}
                </div>
              </div>
            )}
          </div>
        )}

        {/* 財務報表視角 */}
        {activeTab === 'dashboard' && (
          <div className="animate-in fade-in space-y-6">
            <div className="bg-gray-900 rounded-[36px] p-8 text-white relative shadow-2xl">
              <div className="w-14 h-14 bg-white/10 backdrop-blur-md rounded-[20px] flex justify-center items-center mb-6"><Wallet size={26} /></div>
              <span className="text-[15px] text-gray-400 mb-1 block">總資產積壓</span>
              <h2 className="text-[44px] font-bold tracking-tight">${totalAssetValue.toLocaleString()}</h2>
            </div>
            <div className="bg-white rounded-[32px] p-7 shadow-sm flex justify-between items-center">
               <div><span className="text-[14px] font-bold text-gray-400 block mb-1">叫貨資金缺口</span><span className="text-[28px] font-bold text-gray-900">${totalLowStockValue.toLocaleString()}</span></div>
               <div className="w-14 h-14 bg-gray-50 rounded-full flex justify-center items-center text-gray-600"><TrendingDown size={26} /></div>
            </div>
            <div className="bg-white rounded-[36px] p-7 shadow-sm">
              <h3 className="text-[18px] font-bold mb-6 flex items-center gap-2"><PackageOpen size={20} className="text-gray-400"/> 資產分佈</h3>
              <div className="space-y-5">
                {dynamicCategories.filter(c => c !== '全部').map(cat => {
                  const catValue = items.filter(i => i.category === cat).reduce((sum, i) => sum + ((draftStocks[i.id] ?? i.current_stock) * (i.price || 0)), 0);
                  const percent = totalAssetValue === 0 ? 0 : Math.round((catValue / totalAssetValue) * 100);
                  return (
                    <div key={cat}>
                      <div className="flex justify-between mb-2"><span className="text-[14px] font-bold text-gray-600">{cat}</span><span className="text-[15px] font-bold">${catValue.toLocaleString()} ({percent}%)</span></div>
                      <div className="w-full bg-gray-100 rounded-full h-2"><div className="bg-gray-800 h-2 rounded-full" style={{ width: `${percent}%` }}></div></div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* 設定視角 */}
        {activeTab === 'settings' && (
          <div className="animate-in fade-in bg-white rounded-[36px] p-8 shadow-sm">
            <div className="w-14 h-14 bg-gray-50 rounded-[20px] flex justify-center items-center mb-5"><Database size={26} /></div>
            <h2 className="text-[22px] font-bold mb-2">資料庫維護</h2>
            <p className="text-[14px] text-gray-500 font-medium mb-8">上傳 Excel 總表以大批量建立庫存資料。</p>
            <input type="file" accept=".xlsx, .xls" ref={fileInputRef} onChange={handleFileUpload} className="hidden" />
            <button onClick={() => fileInputRef.current?.click()} disabled={isUploading} className="h-14 px-8 bg-gray-900 rounded-full flex gap-3 items-center text-white font-bold w-full justify-center">
              {isUploading ? <Loader2 size={20} className="animate-spin" /> : <Upload size={20} />} {isUploading ? '處理中...' : '上傳 Excel 總表'}
            </button>
          </div>
        )}
      </div>

      {/* 批次送出按鈕 */}
      {hasDraftChanges && activeTab === 'home' && (
        <div className="fixed bottom-[100px] left-1/2 -translate-x-1/2 w-[90%] max-w-[380px] z-40 animate-in slide-in-from-bottom-10 fade-in">
          <button onClick={submitBatchChanges} disabled={isSavingBatch} className="w-full h-16 bg-gray-900 rounded-[24px] flex items-center justify-center gap-3 shadow-lg text-white font-bold text-[17px] active:scale-95">
            {isSavingBatch ? <Loader2 size={24} className="animate-spin" /> : <Save size={24} />} {isSavingBatch ? '同步中...' : `確認送出 (${Object.keys(draftStocks).length} 項異動)`}
          </button>
        </div>
      )}

      {/* 懸浮導覽列 */}
      <div className="fixed bottom-6 left-1/2 -translate-x-1/2 bg-white/90 backdrop-blur-xl border border-gray-100 rounded-full px-6 py-4 flex justify-between items-center w-[90%] max-w-[380px] shadow-lg z-50">
        <button onClick={() => setActiveTab('home')} className={`p-2 transition-all flex flex-col gap-1 ${activeTab === 'home' ? 'text-gray-900 scale-110' : 'text-gray-400'}`}><Home size={24} strokeWidth={2.5} /></button>
        <button onClick={() => {setActiveTab('suppliers'); setActiveSupplier(null);}} className={`p-2 transition-all flex flex-col gap-1 ${activeTab === 'suppliers' ? 'text-gray-900 scale-110' : 'text-gray-400'}`}><Store size={24} strokeWidth={2.5} /></button>
        <button onClick={() => setActiveTab('dashboard')} className={`p-2 transition-all flex flex-col gap-1 ${activeTab === 'dashboard' ? 'text-gray-900 scale-110' : 'text-gray-400'}`}><BarChart3 size={24} strokeWidth={2.5} /></button>
        <button onClick={() => setActiveTab('settings')} className={`p-2 transition-all flex flex-col gap-1 ${activeTab === 'settings' ? 'text-gray-900 scale-110' : 'text-gray-400'}`}><Settings size={24} strokeWidth={2.5} /></button>
      </div>

      {/* 新增/編輯 Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 z-[60] flex items-end sm:items-center justify-center bg-gray-900/40 backdrop-blur-sm">
          <div className="bg-white w-full max-w-md rounded-t-[36px] sm:rounded-[36px] p-7 pb-12 sm:pb-7 animate-in slide-in-from-bottom-full relative">
            <button onClick={() => setIsModalOpen(false)} className="absolute top-6 right-6 w-8 h-8 bg-gray-100 rounded-full flex justify-center items-center text-gray-500"><X size={16} strokeWidth={3} /></button>
            <h3 className="text-[20px] font-bold mb-6 flex items-center gap-2">{editForm.id ? <Pencil size={20} /> : <PlusCircle size={20} />} {editForm.id ? '編輯品項' : '建立新品項'}</h3>
            <div className="space-y-4">
              <div><label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">品項名稱</label><input type="text" value={editForm.name || ''} onChange={e => setEditForm({...editForm, name: e.target.value})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 font-bold outline-none focus:ring-2 ring-gray-200" /></div>
              <div className="grid grid-cols-2 gap-4">
                <div><label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">分類</label><input type="text" value={editForm.category || ''} onChange={e => setEditForm({...editForm, category: e.target.value})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 font-bold outline-none focus:ring-2 ring-gray-200" /></div>
                <div><label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">供應商</label><input type="text" value={editForm.supplier || ''} onChange={e => setEditForm({...editForm, supplier: e.target.value})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 font-bold outline-none focus:ring-2 ring-gray-200" /></div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div><label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">單位</label><input type="text" value={editForm.unit || ''} onChange={e => setEditForm({...editForm, unit: e.target.value})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 font-bold outline-none focus:ring-2 ring-gray-200" /></div>
                <div><label className="text-[13px] font-bold text-gray-400 ml-2 mb-1.5 block">單價 ($)</label><input type="number" value={editForm.price ?? 0} onChange={e => setEditForm({...editForm, price: Number(e.target.value)})} className="w-full h-14 bg-gray-50 rounded-[20px] px-5 font-bold outline-none focus:ring-2 ring-gray-200" /></div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div><label className="text-[13px] font-bold text-gray-900 ml-2 mb-1.5 block">目前庫存</label><input type="number" value={editForm.current_stock ?? 0} onChange={e => setEditForm({...editForm, current_stock: Number(e.target.value)})} className="w-full h-14 bg-gray-100 rounded-[20px] px-5 text-[20px] font-bold outline-none focus:ring-2 ring-gray-300" /></div>
                <div><label className="text-[13px] font-bold text-red-500 ml-2 mb-1.5 block">安全水位</label><input type="number" value={editForm.min_stock ?? 0} onChange={e => setEditForm({...editForm, min_stock: Number(e.target.value)})} className="w-full h-14 bg-red-50 text-red-600 rounded-[20px] px-5 text-[20px] font-bold outline-none focus:ring-2 ring-red-200" /></div>
              </div>
            </div>
            <button onClick={handleSaveItem} disabled={isSavingItem} className="w-full h-16 mt-8 bg-gray-900 rounded-[24px] flex items-center justify-center gap-2 text-white font-bold active:scale-95 transition-all shadow-lg">
              {isSavingItem ? <Loader2 size={24} className="animate-spin" /> : <Check size={24} />} {isSavingItem ? '處理中...' : '確認儲存'}
            </button>
          </div>
        </div>
      )}

      <style dangerouslySetInnerHTML={{__html: `.hide-scrollbar::-webkit-scrollbar { display: none; } .hide-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }`}} />
    </div>
  );
}