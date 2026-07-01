'use client';

import React, { useState, useEffect, useRef } from 'react';
import { supabase } from '../lib/supabase';
import * as XLSX from 'xlsx';
import { Toaster, toast } from 'react-hot-toast';

import { 
  Settings, Loader2, CircleUserRound, Zap, 
  Plus, Minus, AlertTriangle, Search, X, Database, Store, Pencil, 
  PlusCircle, Save, CheckCircle2, ArrowRight, Upload, 
  ChevronLeft, BarChart3, Wallet, TrendingDown, Clock, Trash2, 
  ClipboardCopy, Download, Lock, History, Eraser
} from 'lucide-react';

import { InventoryItem, ExcelRow } from '../types';

const extractErrorMessage = (err: unknown): string => {
  if (err instanceof TypeError && err.message === 'Type error') return '網路連線異常。請確認 .env.local 內的網址正確。';
  if (err instanceof Error) return err.message;
  return String(err);
};

interface ActionLog {
  id: string;
  timestamp: string;
  action: string;
  detail: string;
}

interface BackupRecord {
  id: string;
  date: string;
  data: InventoryItem[];
}

export default function KanDingLiveStocktake() {
  const [isAuthenticated, setIsAuthenticated] = useState<boolean>(false);
  const [passwordInput, setPasswordInput] = useState<string>('');
  const [isAuthError, setIsAuthError] = useState<boolean>(false);

  const [items, setItems] = useState<InventoryItem[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [isUploading, setIsUploading] = useState<boolean>(false);
  const [isClearingDB, setIsClearingDB] = useState<boolean>(false);
  
  // ★ 將預設首頁改為 'suppliers'
  const [activeTab, setActiveTab] = useState<string>('suppliers');
  const [draftStocks, setDraftStocks] = useState<Record<number, number>>({});
  const [isSavingBatch, setIsSavingBatch] = useState<boolean>(false);
  const hasDraftChanges = Object.keys(draftStocks).length > 0;

  const [orderNeeds, setOrderNeeds] = useState<Record<number, number>>({});
  const [countedItems, setCountedItems] = useState<Record<number, number>>({});

  const [searchQuery, setSearchQuery] = useState<string>('');
  const [activeSupplier, setActiveSupplier] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [isModalOpen, setIsModalOpen] = useState<boolean>(false);
  const [isSavingItem, setIsSavingItem] = useState<boolean>(false);
  const [isDeletingItem, setIsDeletingItem] = useState<boolean>(false);
  const [editForm, setEditForm] = useState<Partial<InventoryItem>>({});

  const [actionLogs, setActionLogs] = useState<ActionLog[]>([]);
  const [backups, setBackups] = useState<BackupRecord[]>([]);

  const totalAssetValue = items.reduce((sum, item) => sum + ((draftStocks[item.id] ?? item.current_stock) * (item.price || 0)), 0);
  const totalLowStockValue = items
    .filter(item => (draftStocks[item.id] ?? item.current_stock) < item.min_stock)
    .reduce((sum, item) => {
      const current = draftStocks[item.id] ?? item.current_stock;
      return sum + ((item.min_stock - current) * (item.price || 0));
    }, 0);

  useEffect(() => {
    const savedLogs = localStorage.getItem('kanding_logs');
    if (savedLogs) {
      try { setActionLogs(JSON.parse(savedLogs)); } catch(e) {}
    }

    const savedBackups = localStorage.getItem('kanding_backups');
    if (savedBackups) {
      try {
        const parsedBackups: BackupRecord[] = JSON.parse(savedBackups);
        const now = Date.now();
        const sixtyDaysMs = 60 * 24 * 60 * 60 * 1000;
        const validBackups = parsedBackups.filter(b => now - Number(b.id) < sixtyDaysMs);
        if (validBackups.length !== parsedBackups.length) {
           localStorage.setItem('kanding_backups', JSON.stringify(validBackups));
        }
        setBackups(validBackups);
      } catch(e) {}
    }
  }, []);

  const addLog = (action: string, detail: string) => {
    const newLog: ActionLog = {
      id: Date.now().toString(),
      timestamp: new Date().toLocaleString('zh-TW', { hour12: false }),
      action,
      detail
    };
    setActionLogs(prev => {
      const updated = [newLog, ...prev].slice(0, 150); 
      localStorage.setItem('kanding_logs', JSON.stringify(updated));
      return updated;
    });
  };

  const fetchInventory = async () => {
    setLoading(true);
    try {
      const { data, error } = await supabase.from('kanding_inventory').select('*').order('id', { ascending: true });
      if (error) throw error;
      if (data) setItems(data as InventoryItem[]);
      setDraftStocks({}); 
    } catch (err: unknown) {
      toast.error('無法載入庫存資料，請檢查連線。');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { if (isAuthenticated) fetchInventory(); }, [isAuthenticated]);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    const envPin = process.env.NEXT_PUBLIC_SYSTEM_PIN;
    const isValid = envPin ? passwordInput === envPin : btoa(passwordInput) === 'MTY4ODg=';

    if (isValid) {
      setIsAuthenticated(true);
      toast.success('系統已解鎖', { icon: '🔓' });
      addLog('系統登入', '管理員成功登入系統');
    } else {
      setIsAuthError(true);
      toast.error('密碼錯誤');
      setPasswordInput('');
      setTimeout(() => setIsAuthError(false), 500); 
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

    const item = items.find(i => i.id === id);
    if (item) {
       let need = item.min_stock - finalValue;
       if (need <= 0) need = Math.abs(finalValue - originalValue);
       if (need === 0) need = 1; 
       setOrderNeeds(prev => ({...prev, [id]: need}));
       setCountedItems(prev => ({...prev, [id]: finalValue}));
    }
  };

  useEffect(() => {
    if (!hasDraftChanges) return;
    const timer = setTimeout(() => {
      submitBatchChanges(false); 
    }, 3000);
    return () => clearTimeout(timer);
  }, [draftStocks]);

  const submitBatchChanges = async (isManual = true) => {
    if (!hasDraftChanges) return;
    setIsSavingBatch(true);
    const toastId = isManual ? toast.loading('正在手動同步...') : undefined;

    try {
      const updates = Object.entries(draftStocks).map(([id, newStock]) => ({
        id: Number(id), current_stock: newStock, updated_at: new Date()
      }));

      let logDetails = '';
      updates.forEach(u => {
        const item = items.find(i => i.id === u.id);
        if (item) {
          const diff = u.current_stock - item.current_stock;
          const sign = diff > 0 ? '+' : '';
          logDetails += `${item.name} (${sign}${diff}) `;
        }
      });

      const updatePromises = updates.map(updateData => 
        supabase.from('kanding_inventory').update({ 
          current_stock: updateData.current_stock, updated_at: updateData.updated_at 
        }).eq('id', updateData.id)
      );

      await Promise.all(updatePromises);
      
      if (isManual) {
        toast.success(`成功更新 ${updates.length} 項商品！`, { id: toastId });
      } else {
        toast.success(`已自動儲存 ${updates.length} 項異動`, { icon: '💾', position: 'bottom-center' });
      }
      
      addLog(isManual ? '手動盤點儲存' : '自動盤點儲存', `異動 ${updates.length} 項：${logDetails}`);
      
      setDraftStocks({});
      await fetchInventory();
    } catch (err: unknown) {
      if (isManual) toast.error(`同步失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsSavingBatch(false);
    }
  };

  const handleCopySupplierStock = (supplierName: string) => {
    const supplierItems = items.filter(i => i.supplier === supplierName);
    if (supplierItems.length === 0) return toast.error('此廠商目前沒有商品資料。');

    const today = new Date();
    const dateStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}`;
    
    let text = `【崁頂親子寵物餐廳 - 盤點回報】\n日期：${dateStr}\n廠商：${supplierName}\n\n`;
    
    supplierItems.forEach(item => {
      const currentStock = draftStocks[item.id] ?? item.current_stock;
      text += `${item.name}：${currentStock} ${item.unit}\n`;
    });

    navigator.clipboard.writeText(text.trim());
    toast.success(`已複製 ${supplierName} 的完整盤點紀錄！`, { icon: '📋', duration: 3000 });
    addLog('複製盤點單', `匯出了廠商「${supplierName}」的完整庫存紀錄`);
  };

  const handleCopySupplierLowStock = (supplierName: string) => {
    const supplierItems = items.filter(i => i.supplier === supplierName);
    const lowItems = supplierItems.filter(item => {
      const currentStock = draftStocks[item.id] ?? item.current_stock;
      return currentStock < item.min_stock;
    });

    if (lowItems.length === 0) return toast.error('此廠商目前庫存充足，無需叫貨。', { icon: '✨' });

    const today = new Date();
    const dateStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}`;
    
    let text = `【崁頂親子寵物餐廳 - 叫貨單】\n日期：${dateStr}\n廠商：${supplierName}\n\n`;
    
    lowItems.forEach(item => {
      const currentStock = draftStocks[item.id] ?? item.current_stock;
      let need = item.min_stock - currentStock;
      if (need <= 0) need = 1; 
      text += `${item.name}：${need} ${item.unit}\n`;
    });

    navigator.clipboard.writeText(text.trim());
    toast.success(`已複製 ${supplierName} 的缺貨叫貨單！`, { icon: '📋', duration: 3000 });
    addLog('複製叫貨單', `匯出了廠商「${supplierName}」的缺貨清單`);
  };

  const handleDownloadExcel = () => {
    const lowItems = items.filter(i => (draftStocks[i.id] ?? i.current_stock) < i.min_stock);
    if (lowItems.length === 0) return toast.error('目前沒有低於安全水位的商品可下載。');

    const exportData = lowItems.map(item => {
      const current = draftStocks[item.id] ?? item.current_stock;
      const need = item.min_stock - current;
      return {
        '供應商': item.supplier || '未指定',
        '品項名稱': item.name,
        '分類': item.category,
        '單位': item.unit,
        '單價': item.price || 0,
        '目前庫存': current,
        '安全水位': item.min_stock,
        '建議叫貨量': need,
        '預估補貨成本': need * (item.price || 0),
      };
    }).sort((a, b) => a['供應商'].localeCompare(b['供應商'], 'zh-TW'));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "待叫貨明細清單");
    XLSX.writeFile(workbook, `崁頂親子_缺貨明細表_${new Date().toISOString().slice(0,10)}.xlsx`);
    addLog('下載報表', '下載了 Excel 格式的全店缺貨明細表');
  };

  const handleDownloadBackup = (backup: BackupRecord) => {
    const exportData = backup.data.map(item => ({
        '供應商': item.supplier || '未指定',
        '品項名稱': item.name,
        '分類': item.category,
        '單位': item.unit,
        '單價': item.price || 0,
        '目前庫存': item.current_stock,
        '安全水位': item.min_stock,
        '庫存總值': item.current_stock * (item.price || 0),
    }));
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "歷史備份清單");
    XLSX.writeFile(workbook, `備份檔_${backup.date.replace(/[\/:\s]/g, '')}.xlsx`);
    addLog('下載備份', `下載了 ${backup.date} 的資料庫歷史備份`);
  };

  const handleDeleteBackup = (id: string) => {
    if (!window.confirm('確定要刪除這份歷史紀錄檔嗎？刪除後將無法還原。')) return;
    setBackups(prev => {
        const updated = prev.filter(b => b.id !== id);
        localStorage.setItem('kanding_backups', JSON.stringify(updated));
        return updated;
    });
    toast.success('歷史紀錄檔已刪除');
  };

  const handleClearDatabase = async () => {
    const confirmText = window.prompt('⚠️ 警告：這將永久刪除雲端資料庫中的「所有品項」！\n為防止誤觸，如果您確定要清空，請在下方輸入「確認清空」：');
    
    if (confirmText !== '確認清空') {
      if (confirmText !== null) toast.error('輸入錯誤，已取消清空作業。');
      return;
    }

    setIsClearingDB(true);
    const toastId = toast.loading('正在徹底清空雲端資料庫...');

    try {
      const finalBackup: BackupRecord = {
        id: Date.now().toString(),
        date: new Date().toLocaleString('zh-TW', { hour12: false }),
        data: [...items] 
      };
      setBackups(prev => {
        const updated = [finalBackup, ...prev];
        localStorage.setItem('kanding_backups', JSON.stringify(updated));
        return updated;
      });

      const { error } = await supabase.from('kanding_inventory').delete().gte('id', 0);
      if (error) throw error;

      setItems([]);
      setDraftStocks({});
      addLog('清空資料庫', '🚨 管理員執行了「全店資料庫清空」作業 (已自動建立最終備份)');
      toast.success('資料庫已徹底清空！', { id: toastId, duration: 4000 });
    } catch (err: unknown) {
      toast.error(`清空失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsClearingDB(false);
    }
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!process.env.NEXT_PUBLIC_SUPABASE_URL) return toast.error('請確認環境變數設定。');

    setIsUploading(true);
    const toastId = toast.loading('分析 Excel 並執行資料庫同步中...');

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const jsonData = XLSX.utils.sheet_to_json<ExcelRow>(workbook.Sheets[workbook.SheetNames[0]], { range: 2 });
      if (jsonData.length === 0) throw new Error('解析無資料');

      const newBackup: BackupRecord = {
        id: Date.now().toString(),
        date: new Date().toLocaleString('zh-TW', { hour12: false }),
        data: [...items] 
      };
      setBackups(prev => {
        const updated = [newBackup, ...prev];
        localStorage.setItem('kanding_backups', JSON.stringify(updated));
        return updated;
      });

      const getSafeNumber = (val: unknown) => { const p = parseFloat(String(val)); return isNaN(p) ? 0 : p; };
      let updateCount = 0;
      let insertCount = 0;
      const excelKeys = new Set();

      const formattedData = jsonData.map((row) => {
        const name = String(row['品項名稱'] ?? row['品名'] ?? row['名稱'] ?? '未命名品項');
        const supplier = String(row['供應商'] ?? '未指定');
        excelKeys.add(`${name}_${supplier}`); 

        const existingItem = items.find(i => i.name === name && i.supplier === supplier);
        const payload: any = {
          name,
          category: String(row['類別'] ?? row['分類'] ?? '未分類'),
          supplier,
          price: getSafeNumber(row['單價']),
          current_stock: getSafeNumber(row['盤點數量'] ?? row['目前庫存']),
          min_stock: getSafeNumber(row['安全存量'] ?? row['安全水位']),
          unit: String(row['單位'] ?? row['規格'] ?? '個'),
        };

        if (existingItem) {
          payload.id = existingItem.id;
          payload.updated_at = new Date();
          updateCount++;
        } else {
          insertCount++;
        }
        return payload;
      });

      const idsToDelete = items.filter(i => !excelKeys.has(`${i.name}_${i.supplier}`)).map(i => i.id);
      if (idsToDelete.length > 0) {
         const { error: deleteError } = await supabase.from('kanding_inventory').delete().in('id', idsToDelete);
         if (deleteError) throw deleteError;
      }

      const { error: upsertError } = await supabase.from('kanding_inventory').upsert(formattedData);
      if (upsertError) throw upsertError;
      
      toast.success(`同步完成！更新 ${updateCount} 筆，新增 ${insertCount} 筆，刪除下架商品 ${idsToDelete.length} 筆。`, { id: toastId, duration: 5000 });
      addLog('資料庫總表同步', `Excel 同步：更新 ${updateCount} 筆，新增 ${insertCount} 筆，刪除 ${idsToDelete.length} 筆。已建立覆蓋前備份。`);
      
      fetchInventory(); 
      setActiveTab('suppliers'); // ★ 匯入完成後直接回到廠商列表
    } catch (err: unknown) {
      toast.error(`處理失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const openAddModal = (presets?: { name?: string, supplier?: string }) => {
    setEditForm({
      name: presets?.name || '', category: '未分類',
      supplier: presets?.supplier || '未指定', unit: '個', price: 0, current_stock: 0, min_stock: 0
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
        addLog('修改品項', `更新了「${editForm.name}」的明細資料`);
        toast.success('更新成功', { id: toastId });
      } else {
        const { error } = await supabase.from('kanding_inventory').insert([payload]);
        if (error) throw error;
        addLog('新增品項', `建立了新品項「${editForm.name}」`);
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

  const handleDeleteItem = async () => {
    if (!editForm.id) return;
    if (!window.confirm(`確定要永久刪除「${editForm.name}」嗎？此動作無法復原。`)) return;

    setIsDeletingItem(true);
    const toastId = toast.loading('正在刪除品項...');

    try {
      const { error } = await supabase.from('kanding_inventory').delete().eq('id', editForm.id);
      if (error) throw error;

      setDraftStocks(prev => { const d = { ...prev }; delete d[editForm.id as number]; return d; });
      addLog('刪除品項', `永久刪除了「${editForm.name}」`);
      toast.success('品項已刪除', { id: toastId });
      setIsModalOpen(false);
      await fetchInventory(); 
    } catch (err: unknown) {
      toast.error(`刪除失敗: ${extractErrorMessage(err)}`, { id: toastId });
    } finally {
      setIsDeletingItem(false);
    }
  };

  const suppliersList = Array.from(new Set(items.map(i => i.supplier || '未指定'))).sort();
  const getSupplierStats = (sup: string) => {
    const supItems = items.filter(i => i.supplier === sup);
    const lowCount = supItems.filter(i => (draftStocks[i.id] ?? i.current_stock) < i.min_stock).length;
    const totalValue = supItems.reduce((acc, curr) => acc + ((draftStocks[curr.id] ?? curr.current_stock) * (curr.price || 0)), 0);
    return { count: supItems.length, lowCount, totalValue };
  };

  // 全域搜尋過濾邏輯
  const globalSearchItems = items.filter(item => 
    item.name.toLowerCase().includes(searchQuery.toLowerCase()) || 
    (item.supplier && item.supplier.toLowerCase().includes(searchQuery.toLowerCase()))
  ).sort((a, b) => a.name.localeCompare(b.name, 'zh-TW'));

  const getHeaderSubtext = () => {
    if (activeTab === 'suppliers') return '廠商進銷存管理';
    if (activeTab === 'dashboard') return '財務決策報表';
    if (activeTab === 'history') return '進出貨歷史紀錄';
    return '系統與備份管理';
  };

  // 封裝商品卡片，維持系統風格一致
  const renderItemCard = (item: InventoryItem) => {
    const currentStock = draftStocks[item.id] ?? item.current_stock;
    const isLow = currentStock < item.min_stock;
    const isModified = draftStocks[item.id] !== undefined;

    return (
      <div key={item.id} className={`bg-white rounded-[24px] p-5 transition-all duration-300 border ${isModified ? 'border-[#8780F2] shadow-md' : 'border-transparent shadow-[0_4px_20px_rgb(0,0,0,0.02)]'}`}>
        <div className="flex-1 pr-4 mb-4">
          <div className="flex items-center gap-2 mb-1">
             {isLow && <span className="flex items-center gap-1 bg-[#FF6B6B]/10 text-[#FF6B6B] text-[11px] px-2 py-0.5 rounded-[6px] font-bold"><AlertTriangle size={12}/> 需叫貨</span>}
             {isModified && <span className="bg-gray-100 text-gray-700 text-[11px] px-2 py-0.5 rounded-[8px] font-bold">自動儲存中...</span>}
          </div>
          <div className="flex items-center gap-2">
            <h2 className="text-[17px] font-bold text-gray-900 leading-tight">{item.name}</h2>
            <button onClick={() => openEditModal(item)} className="text-gray-300 hover:text-[#8780F2] transition-colors p-1"><Pencil size={14} strokeWidth={2.5} /></button>
          </div>
          <p className="text-[12px] font-semibold text-gray-400 mt-0.5">單價: ${item.price} • 分類: {item.category}</p>
        </div>
        
        <div className="bg-[#F4F5F9] rounded-[18px] p-1.5 flex items-center justify-between">
          <button onClick={() => handleDraftChange(item.id, currentStock - 1, item.current_stock)} className="w-12 h-12 rounded-[14px] bg-white text-gray-600 flex items-center justify-center shadow-sm active:scale-90 transition-transform"><Minus size={20} strokeWidth={2.5} /></button>
          <span className={`text-[22px] font-bold ${isModified ? 'text-[#8780F2]' : 'text-gray-900'}`}>{currentStock}</span>
          <button onClick={() => handleDraftChange(item.id, currentStock + 1, item.current_stock)} className="w-12 h-12 rounded-[14px] bg-white text-gray-600 flex items-center justify-center shadow-sm active:scale-90 transition-transform"><Plus size={20} strokeWidth={2.5} /></button>
        </div>
      </div>
    );
  };

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
            <input type="password" inputMode="numeric" placeholder="輸入密碼" value={passwordInput} onChange={(e) => setPasswordInput(e.target.value)} className={`w-full h-16 bg-white rounded-[20px] px-6 text-center text-[24px] tracking-[0.5em] font-bold outline-none transition-all shadow-sm ${isAuthError ? 'border-2 border-red-500 animate-shake' : 'border border-transparent focus:ring-2 ring-gray-200'}`} autoFocus />
            <button type="submit" className="w-full h-16 bg-gray-900 rounded-[20px] flex items-center justify-center text-white font-bold text-[17px] active:scale-[0.98] transition-transform shadow-lg hover:bg-black">解鎖系統</button>
         </form>
         <style dangerouslySetInnerHTML={{__html: `@keyframes shake { 0%, 100% { transform: translateX(0); } 25% { transform: translateX(-8px); } 75% { transform: translateX(8px); } } .animate-shake { animation: shake 0.2s ease-in-out 0s 2; }`}} />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#F5F5F7] pb-40 font-sans selection:bg-gray-200">
      <Toaster position="top-center" toastOptions={{ className: 'rounded-[16px] text-sm font-medium shadow-lg' }} />
      
      <div className="px-6 pt-12 pb-6 flex justify-between items-center sticky top-0 z-30 bg-[#F5F5F7]/80 backdrop-blur-xl">
        <div className="flex items-center gap-4">
          <div className="w-[48px] h-[48px] bg-white rounded-full flex items-center justify-center text-gray-800 shadow-[0_4px_20px_rgba(0,0,0,0.04)]">
            <CircleUserRound size={26} strokeWidth={1.5} />
          </div>
          <div>
            <h1 className="text-[20px] font-bold text-gray-900 tracking-tight">崁頂親子寵物餐廳</h1>
            <p className="text-[12px] font-bold text-gray-400 mt-0.5">
              {getHeaderSubtext()}
            </p>
          </div>
        </div>
        <button onClick={fetchInventory} disabled={loading || isSavingBatch} className="w-[48px] h-[48px] bg-white rounded-full flex items-center justify-center shadow-[0_4px_20px_rgba(0,0,0,0.04)] active:scale-95 disabled:opacity-50 transition-transform">
          <Zap size={22} strokeWidth={1.5} className={loading ? "animate-pulse text-gray-400" : ""} />
        </button>
      </div>

      <div className="px-6 space-y-6">

        {/* ========================================== */}
        {/* 廠商與搜尋頁面 (已成為落地第一頁) */}
        {/* ========================================== */}
        {activeTab === 'suppliers' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500">
            
            {/* 跨廠商全域搜尋列 */}
            <div className="relative flex items-center w-full h-[56px] bg-white shadow-[0_8px_24px_rgba(0,0,0,0.03)] rounded-full px-5 focus-within:ring-2 ring-[#8780F2]/30 transition-shadow mb-6">
              <Search size={20} className="text-gray-400" />
              <input type="text" placeholder="跨廠商搜尋或新增品項..." value={searchQuery} onChange={(e) => setSearchQuery(e.target.value)} className="w-full h-full bg-transparent outline-none px-3 text-[15px] font-medium text-gray-900 placeholder:text-gray-400" />
              {searchQuery && <button onClick={() => setSearchQuery('')} className="p-1.5 bg-gray-100 rounded-full text-gray-500 hover:bg-gray-200"><X size={14} strokeWidth={3} /></button>}
            </div>

            {searchQuery ? (
              <div className="space-y-4 animate-in fade-in duration-300">
                {globalSearchItems.length === 0 && (
                  <div onClick={() => openAddModal({ name: searchQuery })} className="bg-[#8780F2]/10 border border-[#8780F2]/20 rounded-[20px] p-4 flex items-center justify-between cursor-pointer hover:bg-[#8780F2]/20 transition-colors">
                    <div className="flex items-center gap-3"><div className="w-10 h-10 bg-[#8780F2] text-white rounded-full flex items-center justify-center"><PlusCircle size={20} /></div><div><h4 className="text-[15px] font-bold text-[#8780F2]">找不到 "{searchQuery}"</h4><p className="text-[12px] font-medium text-[#8780F2]/70">點擊立即新增此品項</p></div></div><ArrowRight size={20} className="text-[#8780F2]" />
                  </div>
                )}
                <div className="space-y-3">
                  {globalSearchItems.map(renderItemCard)}
                </div>
              </div>
            ) : activeSupplier ? (
              <div className="space-y-4 animate-in slide-in-from-right-4 duration-300">
                <button onClick={() => setActiveSupplier(null)} className="flex items-center gap-1 text-gray-500 font-bold text-[14px] mb-2 hover:text-gray-800 transition-colors"><ChevronLeft size={20} /> 返回廠商列表</button>
                
                {/* 雙按鈕的廠商一鍵匯出卡片 */}
                <div className="bg-[#8780F2] rounded-[28px] p-5 flex flex-col gap-4 shadow-[0_8px_30px_rgba(135,128,242,0.3)] text-white mb-2 animate-in slide-in-from-top-4 duration-300">
                  <div className="flex items-center gap-3">
                    <div className="w-12 h-12 bg-white/20 rounded-[16px] flex items-center justify-center">
                      <ClipboardCopy size={24} className="text-white" />
                    </div>
                    <div>
                      <h3 className="text-[17px] font-bold">一鍵匯出 ({activeSupplier})</h3>
                      <p className="text-[13px] text-white/80 mt-0.5">選擇複製完整盤點，或只叫缺貨商品</p>
                    </div>
                  </div>
                  <div className="flex gap-3">
                    <button onClick={() => handleCopySupplierLowStock(activeSupplier)} className="flex-1 py-3 bg-[#FFAD7A] text-white rounded-[16px] text-[14px] font-bold shadow-sm active:scale-95 transition-transform flex items-center justify-center gap-1.5 border border-white/20">
                      <AlertTriangle size={16} /> 複製缺貨 (叫貨)
                    </button>
                    <button onClick={() => handleCopySupplierStock(activeSupplier)} className="flex-1 py-3 bg-white text-[#8780F2] rounded-[16px] text-[14px] font-bold shadow-sm active:scale-95 transition-transform flex items-center justify-center gap-1.5">
                      <ClipboardCopy size={16} /> 複製全部 (盤點)
                    </button>
                  </div>
                </div>

                <div className="bg-white rounded-[32px] p-6 shadow-sm border border-gray-50">
                  <div className="flex items-center justify-between mb-6">
                    <div>
                      <h2 className="text-[22px] font-bold text-gray-900 tracking-tight">{activeSupplier}</h2>
                      <p className="text-[13px] font-medium text-gray-400 mt-1">廠商貨品清單</p>
                    </div>
                    <button onClick={() => openAddModal({ supplier: activeSupplier })} className="w-10 h-10 bg-[#8780F2] text-white rounded-full flex items-center justify-center shadow-lg hover:bg-indigo-600 transition-colors"><Plus size={20} strokeWidth={2.5} /></button>
                  </div>
                  <div className="space-y-4">
                    {items.filter(item => item.supplier === activeSupplier).map(renderItemCard)}
                  </div>
                </div>
              </div>
            ) : (
              <div className="space-y-6 animate-in fade-in duration-300">
                 <div className="flex items-center justify-between px-2 mb-2 mt-2">
                    <h2 className="text-[18px] font-bold text-gray-900 tracking-tight">依廠商檢視與新增</h2>
                    <button onClick={() => openAddModal()} className="flex items-center gap-1.5 text-[13px] font-bold text-[#8780F2] bg-[#8780F2]/10 px-4 py-2 rounded-full hover:bg-[#8780F2]/20 transition-colors"><Plus size={16} /> 新增品項</button>
                 </div>
                 
                 {suppliersList.map(sup => {
                   const stats = getSupplierStats(sup);
                   return (
                     <div key={sup} onClick={() => setActiveSupplier(sup)} className="bg-white rounded-[28px] p-5 shadow-sm flex items-center justify-between cursor-pointer hover:shadow-md transition-all active:scale-[0.98]">
                        <div className="flex items-center gap-4">
                          <div className={`w-12 h-12 rounded-[16px] flex items-center justify-center ${stats.lowCount > 0 ? 'bg-[#FF6B6B]/10 text-[#FF6B6B]' : 'bg-[#F4F5F9] text-gray-500'}`}>
                            <Store size={24} strokeWidth={2} />
                          </div>
                          <div>
                            <h3 className="text-[17px] font-bold text-gray-900">{sup}</h3>
                            <p className="text-[13px] font-medium text-gray-400 mt-0.5">供應 {stats.count} 項商品</p>
                          </div>
                        </div>
                        <div className="text-right">
                           {stats.lowCount > 0 ? <span className="block text-[14px] font-bold text-[#FF6B6B] mb-1">{stats.lowCount} 項缺貨</span> : <span className="block text-[14px] font-bold text-emerald-500 mb-1 flex items-center gap-1"><CheckCircle2 size={14}/> 充足</span>}
                           <ArrowRight size={18} className="text-gray-300 mt-1" />
                        </div>
                     </div>
                   );
                 })}
              </div>
            )}
          </div>
        )}

        {/* ========================================== */}
        {/* 財務與營運報表 */}
        {/* ========================================== */}
        {activeTab === 'dashboard' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-6">
            <h2 className="text-[20px] font-bold text-gray-900 tracking-tight px-1">財務與營運總值</h2>
            <div className="bg-[#292A32] rounded-[36px] p-7 text-white relative overflow-hidden shadow-2xl">
              <div className="relative z-10">
                <div className="w-12 h-12 bg-white/10 backdrop-blur-md rounded-[16px] flex items-center justify-center mb-6"><Wallet size={24} strokeWidth={2} className="text-[#FFAD7A]" /></div>
                <span className="text-[14px] font-medium text-white/60 mb-1 block">目前壓倉庫存總資產</span>
                <h2 className="text-[40px] font-bold leading-none tracking-tight">${totalAssetValue.toLocaleString()}</h2>
              </div>
              <div className="absolute right-0 bottom-0 opacity-10 pointer-events-none transform translate-x-4 translate-y-4"><BarChart3 size={160} strokeWidth={1} /></div>
            </div>
            <div className="bg-white rounded-[32px] p-6 shadow-sm flex items-center justify-between border border-gray-50">
               <div><span className="text-[13px] font-bold text-gray-400 block mb-1">缺貨補齊金流缺口 (預估成本)</span><span className="text-[24px] font-bold text-[#FF6B6B]">${totalLowStockValue.toLocaleString()}</span></div>
               <div className="w-12 h-12 bg-[#FF6B6B]/10 rounded-full flex items-center justify-center text-[#FF6B6B]"><TrendingDown size={24} strokeWidth={2.5} /></div>
            </div>
          </div>
        )}

        {/* ========================================== */}
        {/* 歷史紀錄 */}
        {/* ========================================== */}
        {activeTab === 'history' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-4">
            <h2 className="text-[20px] font-bold text-gray-900 tracking-tight px-1 mb-4">系統操作紀錄</h2>
            {actionLogs.length === 0 ? (
              <div className="bg-white rounded-[32px] py-16 flex flex-col items-center justify-center shadow-sm border border-gray-50">
                <div className="w-16 h-16 bg-gray-50 rounded-full flex items-center justify-center mb-4"><Clock size={28} className="text-gray-300" /></div>
                <p className="text-gray-400 font-bold text-[15px]">目前尚無任何操作紀錄</p>
              </div>
            ) : (
              <div className="space-y-3">
                {actionLogs.map(log => (
                  <div key={log.id} className="bg-white rounded-[24px] p-5 shadow-[0_4px_20px_rgba(0,0,0,0.02)] border border-gray-50 relative overflow-hidden">
                    <div className="absolute left-0 top-0 bottom-0 w-1.5 bg-[#8780F2]"></div>
                    <div className="flex justify-between items-start mb-2 pl-2">
                      <span className="bg-[#8780F2]/10 text-[#8780F2] text-[12px] px-2.5 py-0.5 rounded-[8px] font-bold">{log.action}</span>
                      <span className="text-[12px] font-bold text-gray-400 flex items-center gap-1"><Clock size={12}/> {log.timestamp}</span>
                    </div>
                    <p className="text-[14px] font-medium text-gray-700 leading-relaxed pl-2">{log.detail}</p>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* ========================================== */}
        {/* 設定與匯出 */}
        {/* ========================================== */}
        {activeTab === 'settings' && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-500 space-y-8">
            <div>
              <h2 className="text-[18px] font-bold text-gray-900 tracking-tight px-1 mb-4 flex items-center gap-2">📊 報表與明細匯出</h2>
              <button onClick={handleDownloadExcel} className="w-full bg-white rounded-[24px] p-5 flex items-center justify-center gap-3 shadow-sm border border-gray-50 active:scale-[0.98] transition-transform">
                <div className="w-10 h-10 bg-[#4ADE80]/10 rounded-full flex items-center justify-center"><Download size={20} className="text-[#4ADE80]" /></div>
                <span className="text-[15px] font-bold text-gray-900">下載全店缺貨/補貨明細表 (Excel)</span>
              </button>
            </div>
            <div>
              <h2 className="text-[18px] font-bold text-gray-900 tracking-tight px-1 mb-4 flex items-center gap-2">🗄️ 總表匯入與同步 (覆蓋舊檔)</h2>
              <div className="bg-[#8780F2] rounded-[32px] p-7 text-white relative overflow-hidden shadow-[0_16px_32px_-12px_rgba(135,128,242,0.6)]">
                <div className="relative z-10 w-full">
                  <div className="w-12 h-12 bg-white/20 backdrop-blur-md rounded-[16px] flex items-center justify-center mb-4"><Database size={24} strokeWidth={2} className="text-white" /></div>
                  <h2 className="text-[22px] font-bold leading-tight mb-2">上傳最新 Excel 總表</h2>
                  <p className="text-[13px] text-white/90 leading-relaxed font-medium">系統會自動比對並「覆蓋舊資料與刪除下架商品」。覆蓋前會自動建立歷史備份檔。</p>
                </div>
                <input type="file" accept=".xlsx, .xls" ref={fileInputRef} onChange={handleFileUpload} className="hidden" />
                <div className="mt-6 relative z-10">
                  <button onClick={() => fileInputRef.current?.click()} disabled={isUploading || isClearingDB} className="w-full h-14 bg-[#292A32] rounded-[20px] flex items-center justify-center gap-3 hover:bg-black active:scale-[0.98] transition-transform disabled:opacity-50 shadow-xl">
                    {isUploading ? <Loader2 size={20} className="animate-spin text-white" /> : <Upload size={20} className="text-white" />}<span className="text-[15px] font-bold text-white">{isUploading ? '雲端覆蓋中...' : '選擇 Excel 檔案'}</span>
                  </button>
                </div>
                <div className="absolute -right-6 -bottom-6 opacity-10 pointer-events-none"><Upload size={180} strokeWidth={1.5} /></div>
              </div>
            </div>
            <div>
              <div className="flex justify-between items-end px-1 mb-4">
                 <h2 className="text-[18px] font-bold text-gray-900 tracking-tight flex items-center gap-2">🕰️ 歷史備份檔案</h2>
                 <span className="text-[12px] font-bold text-gray-400">60天自動刪除</span>
              </div>
              {backups.length === 0 ? (
                 <div className="bg-white rounded-[24px] py-10 flex flex-col items-center justify-center shadow-sm border border-gray-50">
                   <p className="text-gray-400 font-bold text-[14px]">目前尚無備份紀錄</p>
                 </div>
              ) : (
                 <div className="space-y-3">
                   {backups.map(b => (
                     <div key={b.id} className="bg-white rounded-[24px] p-5 flex items-center justify-between shadow-sm border border-gray-50">
                       <div>
                         <h4 className="text-[15px] font-bold text-gray-900">{b.date}</h4>
                         <p className="text-[12px] font-medium text-gray-400 mt-1">匯入/清空前備份 • 共 {b.data.length} 項</p>
                       </div>
                       <div className="flex items-center gap-2">
                         <button onClick={() => handleDownloadBackup(b)} className="w-10 h-10 bg-[#4ADE80]/10 rounded-full flex items-center justify-center text-[#4ADE80] active:scale-90 transition-transform" title="下載備份 Excel"><Download size={18} strokeWidth={2.5}/></button>
                         <button onClick={() => handleDeleteBackup(b.id)} className="w-10 h-10 bg-red-50 rounded-full flex items-center justify-center text-red-500 active:scale-90 transition-transform" title="刪除此備份"><Trash2 size={18} strokeWidth={2.5}/></button>
                       </div>
                     </div>
                   ))}
                 </div>
              )}
            </div>
            <div className="pt-4 pb-8">
              <h2 className="text-[18px] font-bold text-red-500 tracking-tight px-1 mb-4 flex items-center gap-2">⚠️ 危險操作區 (Danger Zone)</h2>
              <div className="bg-red-50/50 rounded-[32px] p-6 border border-red-100/50 shadow-sm">
                <div className="flex items-center gap-4 mb-5">
                  <div className="w-12 h-12 bg-red-100 text-red-500 rounded-[16px] flex items-center justify-center"><AlertTriangle size={24} strokeWidth={2} /></div>
                  <div><h3 className="text-[17px] font-bold text-red-600">清空全部資料庫</h3><p className="text-[12px] font-medium text-red-400 mt-0.5">刪除所有品項，此動作需打字確認</p></div>
                </div>
                <button onClick={handleClearDatabase} disabled={isClearingDB || isUploading} className="w-full h-14 bg-white text-red-500 border border-red-200 rounded-[20px] flex items-center justify-center gap-2 font-bold active:scale-[0.98] transition-all shadow-sm hover:bg-red-50 disabled:opacity-50">
                  {isClearingDB ? <Loader2 size={20} className="animate-spin" /> : <Trash2 size={20} />}<span className="text-[15px]">{isClearingDB ? '正在執行清空...' : '永久清空資料庫'}</span>
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* 底部居中的藥丸型「自動儲存倒數」提示 */}
      {hasDraftChanges && activeTab === 'suppliers' && (
        <div className="fixed bottom-[100px] left-1/2 -translate-x-1/2 z-40 animate-in slide-in-from-bottom-5 fade-in duration-300 w-auto whitespace-nowrap">
          <button 
            onClick={() => submitBatchChanges(true)}
            className="px-6 h-14 bg-[#292A32] rounded-full flex items-center justify-center gap-3 shadow-[0_20px_40px_rgba(0,0,0,0.3)] hover:bg-black active:scale-[0.98] transition-all"
          >
            <Loader2 size={20} className="animate-spin text-[#8780F2]" />
            <span className="text-[14px] font-bold text-white tracking-wide">
              3秒後自動儲存 ({Object.keys(draftStocks).length} 項)
            </span>
            <div className="w-[1px] h-5 bg-white/20 ml-1"></div>
            <span className="text-[14px] font-bold text-[#8780F2]">立即儲存</span>
          </button>
        </div>
      )}

      {/* 新版極簡導覽列：拔除 Home 圖示 */}
      <div className="fixed bottom-8 left-1/2 -translate-x-1/2 bg-[#292A32] rounded-full px-5 py-3.5 flex justify-between items-center w-[92%] max-w-[320px] shadow-[0_24px_40px_-12px_rgba(0,0,0,0.5)] z-40">
        <button onClick={() => {setActiveTab('suppliers'); setActiveSupplier(null); setSearchQuery('');}} className={`relative p-2.5 rounded-full transition-all duration-300 ${activeTab === 'suppliers' ? 'text-white' : 'text-gray-500 hover:text-gray-300'}`}>
          {activeTab === 'suppliers' && <span className="absolute inset-0 bg-[#FFAD7A] rounded-full -z-10 shadow-[0_0_15px_rgba(255,173,122,0.5)]"></span>}<Store size={22} strokeWidth={2.5} />
        </button>
        <button onClick={() => {setActiveTab('dashboard'); setSearchQuery('');}} className={`relative p-2.5 rounded-full transition-all duration-300 ${activeTab === 'dashboard' ? 'text-white' : 'text-gray-500 hover:text-gray-300'}`}>
          {activeTab === 'dashboard' && <span className="absolute inset-0 bg-[#4ADE80] rounded-full -z-10 shadow-[0_0_15px_rgba(74,222,128,0.4)]"></span>}<BarChart3 size={22} strokeWidth={2.5} />
        </button>
        <button onClick={() => {setActiveTab('history'); setSearchQuery('');}} className={`relative p-2.5 rounded-full transition-all duration-300 ${activeTab === 'history' ? 'text-white' : 'text-gray-500 hover:text-gray-300'}`}>
          {activeTab === 'history' && <span className="absolute inset-0 bg-[#38BDF8] rounded-full -z-10 shadow-[0_0_15px_rgba(56,189,248,0.4)]"></span>}<History size={22} strokeWidth={2.5} />
        </button>
        <button onClick={() => {setActiveTab('settings'); setSearchQuery('');}} className={`relative p-2.5 rounded-full transition-all duration-300 ${activeTab === 'settings' ? 'text-white' : 'text-gray-500 hover:text-gray-300'}`}>
          {activeTab === 'settings' && <span className="absolute inset-0 bg-gray-600 rounded-full -z-10 shadow-[0_0_15px_rgba(75,85,99,0.5)]"></span>}<Settings size={22} strokeWidth={2.5} />
        </button>
      </div>

      {isModalOpen && (
        <div className="fixed inset-0 z-[60] flex items-end justify-center bg-[#292A32]/40 backdrop-blur-sm sm:items-center">
          <div className="bg-white w-full max-w-md rounded-t-[36px] sm:rounded-[36px] p-7 pb-12 sm:pb-7 animate-in slide-in-from-bottom-full duration-300 shadow-2xl relative">
            <button onClick={() => setIsModalOpen(false)} className="absolute top-6 right-6 w-8 h-8 bg-gray-100 rounded-full flex items-center justify-center text-gray-500 hover:bg-gray-200 transition-colors"><X size={16} strokeWidth={3} /></button>
            <h3 className="text-[20px] font-bold text-gray-900 mb-6 flex items-center gap-2">{editForm.id ? <Pencil size={20} className="text-[#8780F2]" /> : <PlusCircle size={20} className="text-[#FFAD7A]" />} {editForm.id ? '編輯品項明細' : '建立新品項'}</h3>
            
            <div className="space-y-4">
              <div><label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">品項名稱</label><input type="text" value={editForm.name || ''} onChange={e => setEditForm({...editForm, name: e.target.value})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" /></div>
              <div className="grid grid-cols-2 gap-4">
                <div><label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">分類</label><input type="text" value={editForm.category || ''} onChange={e => setEditForm({...editForm, category: e.target.value})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" /></div>
                <div><label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">供應商</label><input type="text" value={editForm.supplier || ''} onChange={e => setEditForm({...editForm, supplier: e.target.value})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" /></div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div><label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">單位</label><input type="text" value={editForm.unit || ''} onChange={e => setEditForm({...editForm, unit: e.target.value})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" /></div>
                <div><label className="text-[12px] font-bold text-gray-400 ml-1 mb-1 block">單價 ($)</label><input type="number" min="0" value={editForm.price ?? 0} onChange={e => setEditForm({...editForm, price: Number(e.target.value)})} className="w-full h-12 bg-[#F4F5F9] rounded-[16px] px-4 text-[15px] font-bold text-gray-900 outline-none focus:ring-2 ring-[#8780F2]/50" /></div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div><label className="text-[12px] font-bold text-[#8780F2] ml-1 mb-1 block">目前庫存</label><input type="number" min="0" value={editForm.current_stock ?? 0} onChange={e => setEditForm({...editForm, current_stock: Number(e.target.value)})} className="w-full h-12 bg-[#8780F2]/10 text-[#8780F2] rounded-[16px] px-4 text-[18px] font-bold outline-none focus:ring-2 ring-[#8780F2]/50" /></div>
                <div><label className="text-[12px] font-bold text-[#FF6B6B] ml-1 mb-1 block">安全水位</label><input type="number" min="0" value={editForm.min_stock ?? 0} onChange={e => setEditForm({...editForm, min_stock: Number(e.target.value)})} className="w-full h-12 bg-[#FF6B6B]/10 text-[#FF6B6B] rounded-[16px] px-4 text-[18px] font-bold outline-none focus:ring-2 ring-[#FF6B6B]/50" /></div>
              </div>
            </div>
            
            <div className="flex gap-3 mt-8">
              {editForm.id && (
                <button onClick={handleDeleteItem} disabled={isDeletingItem || isSavingItem} className="w-14 h-14 bg-red-50 text-red-500 rounded-[20px] flex items-center justify-center hover:bg-red-100 active:scale-95 transition-all shadow-sm" title="永久刪除">
                  {isDeletingItem ? <Loader2 size={20} className="animate-spin" /> : <Trash2 size={20} />}
                </button>
              )}
              <button onClick={handleSaveItem} disabled={isSavingItem || isDeletingItem} className="flex-1 h-14 bg-[#292A32] rounded-[20px] flex items-center justify-center gap-2 text-white active:scale-95 transition-transform shadow-lg">
                {isSavingItem ? <Loader2 size={20} className="animate-spin" /> : <Save size={20} />} <span className="text-[16px] font-bold">{isSavingItem ? '處理中...' : '儲存變更'}</span>
              </button>
            </div>
          </div>
        </div>
      )}
      <style dangerouslySetInnerHTML={{__html: `.hide-scrollbar::-webkit-scrollbar { display: none; } .hide-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }`}} />
    </div>
  );
}