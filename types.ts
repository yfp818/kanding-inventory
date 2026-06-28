// src/types.ts

export interface InventoryItem {
  id: number;
  name: string;
  category: string;
  supplier: string;      // 新增：供應商
  price: number;         // 新增：單價
  current_stock: number;
  min_stock: number;
  unit: string;
}

export interface SystemErrorLog {
  action: string;
  error_message: string;
  created_at: string;
}

export interface ExcelRow {
  品項名稱?: string;
  品名?: string;
  名稱?: string;
  商品名稱?: string;
  類別?: string;
  分類?: string;
  商品分類?: string;
  供應商?: string;        // 新增：Excel 對接
  單價?: string | number; // 新增：Excel 對接
  盤點數量?: string | number;
  目前庫存?: string | number;
  庫存?: string | number;
  安全存量?: string | number;
  安全水位?: string | number;
  安全庫存?: string | number;
  單位?: string;
  規格?: string;
  [key: string]: unknown;
}