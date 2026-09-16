import React, { useState, useMemo } from 'react';
import {
  Sparkles,
  ClipboardPaste,
  Plus,
  Trash2,
  FileSpreadsheet,
  ArrowDownToLine,
  CheckCircle2,
  AlertCircle,
  Settings2,
  RotateCcw,
  Layers,
  ChevronDown,
  ChevronUp,
} from 'lucide-react';
import { CatalogItem, MasterRowData } from '../types';
import { lookupCatalogBySapCode } from '../utils/catalogManager';

export interface SapInputEntry {
  id: string;
  sapName: string;
  tag?: string;
  yard?: string;
  supplier?: string;
  poNumber?: string;
}

interface SapBatchInputTableProps {
  catalogItems: CatalogItem[];
  activeShipName: string;
  onApplyToMaster: (rows: MasterRowData[], mode: 'replace' | 'append') => void;
  onExportDirectExcel: (rows: MasterRowData[]) => void;
  onOpenCatalogSettings?: () => void;
  currentMasterCount: number;
}

export function SapBatchInputTable({
  catalogItems,
  activeShipName,
  onApplyToMaster,
  onExportDirectExcel,
  onOpenCatalogSettings,
  currentMasterCount,
}: SapBatchInputTableProps) {
  // Initial list of SAP Names
  const [entries, setEntries] = useState<SapInputEntry[]>([
    { id: '1', sapName: 'A4010320608F', tag: 'V-01', yard: 'Yard A' },
    { id: '2', sapName: 'A401032060PF', tag: 'V-02', yard: 'Yard A' },
    { id: '3', sapName: 'A5051220602A', tag: 'V-03', yard: 'Yard A' },
    { id: '4', sapName: 'V5518319100B', tag: 'V-04', yard: 'Yard A' },
  ]);

  const [pasteModalOpen, setPasteModalOpen] = useState(false);
  const [pasteText, setPasteText] = useState('');
  const [appendMode, setAppendMode] = useState<'replace' | 'append'>('replace');
  const [commonYard, setCommonYard] = useState('Yard A');
  const [commonSupplier, setCommonSupplier] = useState('');
  const [commonPo, setCommonPo] = useState('');
  const [isBatchConfigOpen, setIsBatchConfigOpen] = useState(false);

  // Evaluate matches in real-time
  const evaluatedEntries = useMemo(() => {
    return entries.map((entry, index) => {
      const trimmed = entry.sapName.trim();
      const lookup = trimmed ? lookupCatalogBySapCode(trimmed, catalogItems) : { found: false, armatureFields: {} };
      return {
        ...entry,
        index: index + 1,
        trimmed,
        isMatched: lookup.found,
        matchedItem: lookup.matchedItem,
        armatureFields: lookup.armatureFields,
      };
    });
  }, [entries, catalogItems]);

  const matchedCount = useMemo(
    () => evaluatedEntries.filter((e) => e.trimmed && e.isMatched).length,
    [evaluatedEntries]
  );
  const totalEntered = useMemo(
    () => evaluatedEntries.filter((e) => e.trimmed).length,
    [evaluatedEntries]
  );

  // Build MasterRowData objects
  const buildMasterRows = (): MasterRowData[] => {
    return evaluatedEntries
      .filter((e) => e.trimmed)
      .map((e, idx) => {
        const specFields = e.armatureFields || {};
        const rowId = `sap-import-${Date.now()}-${idx}-${Math.random().toString(36).substring(2, 6)}`;

        const newRow: MasterRowData = {
          _id: rowId,
          _sourceFile: `SAP_Batch_Import`,
          _rowNumber: idx + 1,
          TAG: e.tag?.trim() || `VALVE-${String(idx + 1).padStart(3, '0')}`,
          'ACTUATOR TAG': '',
          'P.O. NUMBER': e.poNumber?.trim() || commonPo.trim() || '',
          'SUPPLIER': e.supplier?.trim() || specFields['SUPPLIER'] || commonSupplier.trim() || '',
          'DESTINATION (YARD)': e.yard?.trim() || commonYard.trim() || 'Yard A',
          ...specFields,
        };

        return newRow;
      });
  };

  // Handlers
  const handleUpdateEntry = (id: string, field: keyof SapInputEntry, value: string) => {
    setEntries((prev) =>
      prev.map((e) => (e.id === id ? { ...e, [field]: value } : e))
    );
  };

  const handleAddRow = () => {
    const newId = `row-${Date.now()}-${Math.random().toString(36).substring(2, 5)}`;
    setEntries((prev) => [
      ...prev,
      {
        id: newId,
        sapName: '',
        tag: `V-${String(prev.length + 1).padStart(2, '0')}`,
        yard: commonYard,
        supplier: commonSupplier,
        poNumber: commonPo,
      },
    ]);
  };

  const handleDeleteRow = (id: string) => {
    setEntries((prev) => prev.filter((e) => e.id !== id));
  };

  const handleClearAll = () => {
    setEntries([]);
  };

  const handleLoadSampleCodes = () => {
    setEntries([
      { id: 's1', sapName: 'A4010320608F', tag: 'V-101', yard: 'Yard A' },
      { id: 's2', sapName: 'A401032060PF', tag: 'V-102', yard: 'Yard A' },
      { id: 's3', sapName: 'A4010320708T', tag: 'V-103', yard: 'Yard A' },
      { id: 's4', sapName: 'A401032070PT', tag: 'V-104', yard: 'Yard A' },
      { id: 's5', sapName: 'A5051220602A', tag: 'V-105', yard: 'Yard A' },
      { id: 's6', sapName: 'V5518319100B', tag: 'V-106', yard: 'Yard A' },
    ]);
  };

  // Handle Batch Paste from Clipboard
  const handleApplyPasteText = () => {
    if (!pasteText.trim()) {
      setPasteModalOpen(false);
      return;
    }

    // Split by newlines or commas/semicolons
    const rawLines = pasteText.split(/\r?\n/);
    const newItems: SapInputEntry[] = [];

    rawLines.forEach((line, idx) => {
      const cleanLine = line.trim();
      if (!cleanLine) return;

      // Check if tab-delimited (e.g. copied directly from Excel with multiple columns: SAP Name \t TAG \t Yard...)
      const parts = cleanLine.split('\t');
      const sapName = parts[0]?.trim() || '';
      if (!sapName) return;

      const tag = parts[1]?.trim() || `V-${String(newItems.length + 1).padStart(2, '0')}`;
      const yard = parts[2]?.trim() || commonYard;
      const supplier = parts[3]?.trim() || commonSupplier;
      const poNumber = parts[4]?.trim() || commonPo;

      newItems.push({
        id: `paste-${Date.now()}-${idx}`,
        sapName,
        tag,
        yard,
        supplier,
        poNumber,
      });
    });

    if (newItems.length > 0) {
      setEntries((prev) => [...prev.filter((e) => e.sapName.trim() !== ''), ...newItems]);
    }

    setPasteText('');
    setPasteModalOpen(false);
  };

  // Handle cell paste inside table
  const handleCellPaste = (
    e: React.ClipboardEvent<HTMLInputElement>,
    rowIndex: number
  ) => {
    const pastedText = e.clipboardData.getData('text');
    if (!pastedText) return;

    const lines = pastedText.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (lines.length > 1) {
      e.preventDefault();
      // Multi-line paste from Excel
      const newEntries = [...entries];
      lines.forEach((line, offset) => {
        const parts = line.split('\t');
        const code = parts[0]?.trim();
        if (!code) return;

        const targetIdx = rowIndex + offset;
        const entryData: SapInputEntry = {
          id: `p-${Date.now()}-${offset}`,
          sapName: code,
          tag: parts[1]?.trim() || `V-${String(targetIdx + 1).padStart(2, '0')}`,
          yard: parts[2]?.trim() || commonYard,
          supplier: parts[3]?.trim() || commonSupplier,
          poNumber: parts[4]?.trim() || commonPo,
        };

        if (targetIdx < newEntries.length) {
          newEntries[targetIdx] = { ...newEntries[targetIdx], ...entryData };
        } else {
          newEntries.push(entryData);
        }
      });
      setEntries(newEntries);
    }
  };

  // Actions
  const handleExportToArmatureList = () => {
    const rows = buildMasterRows();
    if (rows.length === 0) return;
    onApplyToMaster(rows, appendMode);
  };

  const handleExportDirect = () => {
    const rows = buildMasterRows();
    if (rows.length === 0) return;
    onExportDirectExcel(rows);
  };

  return (
    <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden mb-5">
      {/* Header bar */}
      <div className="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 px-4 py-3.5 text-white flex flex-col md:flex-row md:items-center md:justify-between gap-3">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 rounded-lg bg-blue-500/20 border border-blue-400/30 flex items-center justify-center shrink-0">
            <Layers className="w-5 h-5 text-blue-300" />
          </div>
          <div>
            <div className="flex items-center gap-2">
              <h2 className="text-sm font-bold tracking-wide uppercase">
                Bảng Nhập Hàng Loạt SAP Name &rarr; Xuất Armature List
              </h2>
              <span className="text-[11px] px-2 py-0.5 rounded-full bg-blue-500/30 border border-blue-400/40 text-blue-200 font-mono">
                {totalEntered} mã đã nhập
              </span>
            </div>
            <p className="text-xs text-slate-300 mt-0.5 flex items-center gap-2">
              <span>Đang đối chiếu với Catalog tàu:</span>
              <span className="font-semibold text-amber-300 bg-amber-950/60 px-2 py-0.5 rounded border border-amber-500/30">
                {activeShipName} ({catalogItems.length} mã van trong kho)
              </span>
            </p>
          </div>
        </div>

        {/* Top Control Buttons */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Quick Paste Button */}
          <button
            id="btn-open-paste-modal"
            onClick={() => setPasteModalOpen(true)}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-indigo-600 hover:bg-indigo-500 text-white shadow-sm transition-all cursor-pointer active:scale-95"
            title="Mở khung dán nhanh nhiều mã SAP từ Excel hoặc Notepad"
          >
            <ClipboardPaste className="w-3.5 h-3.5" />
            <span>Dán Nhanh Danh Sách SAP</span>
          </button>

          {/* Sample Codes */}
          <button
            id="btn-load-sample-sap"
            onClick={handleLoadSampleCodes}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors cursor-pointer"
            title="Nạp 6 mã van SAP mẫu có sẵn trong Catalog"
          >
            <Sparkles className="w-3.5 h-3.5 text-amber-300" />
            <span className="hidden sm:inline">Nạp 6 Mã Mẫu</span>
          </button>

          {/* Add Row Button */}
          <button
            id="btn-add-sap-row"
            onClick={handleAddRow}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors cursor-pointer"
            title="Thêm một dòng nhập mới"
          >
            <Plus className="w-3.5 h-3.5 text-blue-400" />
            <span>Thêm Dòng</span>
          </button>

          {/* Clear Button */}
          {entries.length > 0 && (
            <button
              id="btn-clear-sap-entries"
              onClick={handleClearAll}
              className="inline-flex items-center gap-1 px-2 py-1.5 rounded-lg text-xs font-medium text-rose-300 hover:text-white hover:bg-rose-900/40 transition-colors cursor-pointer"
              title="Xóa sạch danh sách đang nhập"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Xóa hết</span>
            </button>
          )}

          {/* Catalog Settings (Hidden by default, openable) */}
          {onOpenCatalogSettings && (
            <button
              id="btn-open-catalog-settings"
              onClick={onOpenCatalogSettings}
              className="p-1.5 rounded-lg text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors cursor-pointer ml-1"
              title="Cấu hình / Quản lý file Catalog gốc (Hiện đang chạy ẩn)"
            >
              <Settings2 className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* Batch defaults config accordion (optional) */}
      <div className="bg-slate-50 dark:bg-slate-850 border-b border-slate-200 dark:border-slate-800 px-4 py-2">
        <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
          <button
            onClick={() => setIsBatchConfigOpen(!isBatchConfigOpen)}
            className="flex items-center gap-1.5 font-medium text-slate-600 dark:text-slate-300 hover:text-blue-600 cursor-pointer"
          >
            <span className="font-semibold">Thiết lập chung cho Functional Design (Yard, P.O, Nhà cung cấp):</span>
            {isBatchConfigOpen ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
          </button>

          <div className="flex items-center gap-3">
            <span className="text-slate-500">Trạng thái tra cứu Catalog:</span>
            <span className="font-bold text-emerald-600 dark:text-emerald-400 flex items-center gap-1">
              <CheckCircle2 className="w-3.5 h-3.5" /> {matchedCount} khớp
            </span>
            {totalEntered - matchedCount > 0 && (
              <span className="font-bold text-amber-600 dark:text-amber-400 flex items-center gap-1">
                <AlertCircle className="w-3.5 h-3.5" /> {totalEntered - matchedCount} chưa khớp
              </span>
            )}
          </div>
        </div>

        {isBatchConfigOpen && (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-3 pb-1 mt-2 border-t border-slate-200 dark:border-slate-800">
            <div>
              <label className="block text-[11px] font-semibold text-slate-500 uppercase mb-1">
                DESTINATION (YARD) MẶC ĐỊNH
              </label>
              <input
                type="text"
                value={commonYard}
                onChange={(e) => setCommonYard(e.target.value)}
                placeholder="VD: Yard A / Xưởng đóng tàu"
                className="w-full px-2.5 py-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100"
              />
            </div>
            <div>
              <label className="block text-[11px] font-semibold text-slate-500 uppercase mb-1">
                SUPPLIER (NHÀ CUNG CẤP) MẶC ĐỊNH
              </label>
              <input
                type="text"
                value={commonSupplier}
                onChange={(e) => setCommonSupplier(e.target.value)}
                placeholder="VD: Econosto / Danfoss"
                className="w-full px-2.5 py-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100"
              />
            </div>
            <div>
              <label className="block text-[11px] font-semibold text-slate-500 uppercase mb-1">
                P.O. NUMBER MẶC ĐỊNH
              </label>
              <input
                type="text"
                value={commonPo}
                onChange={(e) => setCommonPo(e.target.value)}
                placeholder="VD: PO-2026-09"
                className="w-full px-2.5 py-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100"
              />
            </div>
          </div>
        )}
      </div>

      {/* Main SAP Input Table */}
      <div className="overflow-x-auto max-h-[380px] overflow-y-auto">
        <table className="w-full text-left text-xs border-collapse">
          <thead className="sticky top-0 z-10 bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-200 border-b border-slate-300 dark:border-slate-700 text-[11px]">
            <tr>
              <th className="py-2 px-3 w-10 text-center text-slate-400 font-mono">#</th>
              <th className="py-2 px-3 font-bold text-blue-700 dark:text-blue-400 min-w-[190px]">
                SAP NAME (NHẬP MÃ / PASTE VÀO ĐÂY) *
              </th>
              <th className="py-2 px-3 min-w-[130px]">TRẠNG THÁI CATALOG</th>
              <th className="py-2 px-3 min-w-[180px]">DESCRIPTION TỰ ĐỘNG TỪ CATALOG</th>
              <th className="py-2 px-3 min-w-[120px]">NORMALE N° / SPEC</th>
              <th className="py-2 px-3 min-w-[110px]">SIZE / RATING</th>
              <th className="py-2 px-3 min-w-[110px]">TAG (FUNCTIONAL)</th>
              <th className="py-2 px-3 min-w-[100px]">YARD</th>
              <th className="py-2 px-2 w-12 text-center">XÓA</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100 dark:divide-slate-800 bg-white dark:bg-slate-900">
            {evaluatedEntries.length === 0 ? (
              <tr>
                <td colSpan={9} className="py-8 text-center text-slate-400">
                  <div className="max-w-md mx-auto space-y-2">
                    <p className="font-semibold text-slate-600 dark:text-slate-300">
                      Chưa có mã SAP nào trong bảng.
                    </p>
                    <p className="text-xs text-slate-400">
                      Bấm vào nút <strong>"Dán Nhanh Danh Sách SAP"</strong> hoặc <strong>"Nạp 6 Mã Mẫu"</strong> hoặc <strong>"Thêm Dòng"</strong> để bắt đầu.
                    </p>
                  </div>
                </td>
              </tr>
            ) : (
              evaluatedEntries.map((row, rIdx) => {
                const spec = row.armatureFields || {};
                return (
                  <tr
                    key={row.id}
                    className={`hover:bg-blue-50/40 dark:hover:bg-slate-800/60 transition-colors ${
                      !row.trimmed
                        ? 'opacity-60'
                        : row.isMatched
                        ? 'bg-emerald-50/15 dark:bg-emerald-950/10'
                        : 'bg-amber-50/20 dark:bg-amber-950/10'
                    }`}
                  >
                    <td className="py-1.5 px-3 text-center text-slate-400 font-mono text-[11px]">
                      {row.index}
                    </td>

                    {/* SAP Name Input */}
                    <td className="py-1.5 px-3">
                      <input
                        type="text"
                        value={row.sapName}
                        onChange={(e) => handleUpdateEntry(row.id, 'sapName', e.target.value)}
                        onPaste={(e) => handleCellPaste(e, rIdx)}
                        placeholder="Nhập hoặc dán mã SAP..."
                        className="w-full px-2.5 py-1 text-xs font-mono font-bold text-blue-900 dark:text-blue-300 bg-slate-50 dark:bg-slate-950 border border-slate-300 dark:border-slate-700 rounded focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none"
                      />
                    </td>

                    {/* Catalog Match Status */}
                    <td className="py-1.5 px-3 whitespace-nowrap">
                      {!row.trimmed ? (
                        <span className="text-slate-400 text-[11px] italic">Chưa nhập mã</span>
                      ) : row.isMatched ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-100 text-emerald-800 dark:bg-emerald-950/60 dark:text-emerald-300 border border-emerald-300 dark:border-emerald-800">
                          <CheckCircle2 className="w-3 h-3" /> Đã tìm thấy
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-amber-100 text-amber-800 dark:bg-amber-950/60 dark:text-amber-300 border border-amber-300 dark:border-amber-800">
                          <AlertCircle className="w-3 h-3" /> Chưa có trong Catalog
                        </span>
                      )}
                    </td>

                    {/* Auto Description */}
                    <td className="py-1.5 px-3 text-slate-800 dark:text-slate-200">
                      {spec['DESCRIPTION'] ? (
                        <span className="font-medium text-slate-900 dark:text-slate-100">
                          {spec['DESCRIPTION']}
                        </span>
                      ) : (
                        <span className="text-slate-400 italic text-[11px]">
                          {row.trimmed ? '-' : ''}
                        </span>
                      )}
                    </td>

                    {/* Auto Normale Nr */}
                    <td className="py-1.5 px-3 font-mono text-slate-700 dark:text-slate-300">
                      {spec['STD DRW NORMALE N°'] ? (
                        <span className="bg-slate-100 dark:bg-slate-800 px-1.5 py-0.5 rounded">
                          {spec['STD DRW NORMALE N°']}
                        </span>
                      ) : (
                        <span className="text-slate-400 italic text-[11px]">-</span>
                      )}
                    </td>

                    {/* Auto Size / Rating */}
                    <td className="py-1.5 px-3 text-slate-700 dark:text-slate-300">
                      {spec['SIZE'] || spec['PRESSURE RATING'] ? (
                        <span>
                          {spec['SIZE'] || '-'} / {spec['PRESSURE RATING'] || '-'}
                        </span>
                      ) : (
                        <span className="text-slate-400 italic text-[11px]">-</span>
                      )}
                    </td>

                    {/* TAG input */}
                    <td className="py-1.5 px-3">
                      <input
                        type="text"
                        value={row.tag || ''}
                        onChange={(e) => handleUpdateEntry(row.id, 'tag', e.target.value)}
                        placeholder="VD: V-01"
                        className="w-full px-2 py-1 text-xs bg-transparent border border-slate-200 dark:border-slate-800 rounded hover:border-slate-400 focus:border-blue-500 outline-none"
                      />
                    </td>

                    {/* Yard input */}
                    <td className="py-1.5 px-3">
                      <input
                        type="text"
                        value={row.yard || ''}
                        onChange={(e) => handleUpdateEntry(row.id, 'yard', e.target.value)}
                        placeholder="Yard A"
                        className="w-full px-2 py-1 text-xs bg-transparent border border-slate-200 dark:border-slate-800 rounded hover:border-slate-400 focus:border-blue-500 outline-none"
                      />
                    </td>

                    {/* Delete row */}
                    <td className="py-1.5 px-2 text-center">
                      <button
                        onClick={() => handleDeleteRow(row.id)}
                        className="p-1 text-slate-400 hover:text-rose-500 rounded transition-colors cursor-pointer"
                        title="Xóa dòng này"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Action Footer Bar */}
      <div className="bg-slate-50 dark:bg-slate-850 p-3.5 border-t border-slate-200 dark:border-slate-800 flex flex-col md:flex-row md:items-center md:justify-between gap-3">
        <div className="flex items-center gap-3 text-xs">
          <div className="flex items-center gap-2">
            <span className="text-slate-600 dark:text-slate-300 font-semibold">Chế độ đưa vào Master:</span>
            <select
              value={appendMode}
              onChange={(e) => setAppendMode(e.target.value as 'replace' | 'append')}
              className="px-2.5 py-1 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-200 cursor-pointer font-medium"
            >
              <option value="replace">Ghi đè hoàn toàn (Tạo mới Armature List)</option>
              <option value="append">Thêm nối tiếp vào Armature List hiện có</option>
            </select>
          </div>
          {currentMasterCount > 0 && (
            <span className="text-slate-400 text-[11px]">
              (Hiện tại Master đang có {currentMasterCount} hàng)
            </span>
          )}
        </div>

        {/* Big Action Buttons */}
        <div className="flex items-center gap-2.5">
          {/* Export directly to Excel */}
          <button
            id="btn-direct-export-excel"
            onClick={handleExportDirect}
            disabled={totalEntered === 0}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold transition-all shadow-sm ${
              totalEntered > 0
                ? 'bg-slate-800 hover:bg-slate-700 text-white border border-slate-700 cursor-pointer active:scale-95'
                : 'bg-slate-200 text-slate-400 dark:bg-slate-800 dark:text-slate-600 cursor-not-allowed'
            }`}
            title="Tạo và tải về ngay file Excel Armature Master chuẩn 2 tầng từ danh sách SAP Name này"
          >
            <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
            <span>Xuất Trực Tiếp Excel (.xlsx)</span>
          </button>

          {/* Generate & Apply to Master Table */}
          <button
            id="btn-apply-sap-to-master"
            onClick={handleExportToArmatureList}
            disabled={totalEntered === 0}
            className={`inline-flex items-center gap-2 px-5 py-2 rounded-lg text-xs font-bold transition-all shadow-md ${
              totalEntered > 0
                ? 'bg-blue-600 hover:bg-blue-500 text-white cursor-pointer active:scale-95 ring-2 ring-blue-400/30'
                : 'bg-slate-200 text-slate-400 dark:bg-slate-800 dark:text-slate-600 cursor-not-allowed'
            }`}
            title="Chuyển toàn bộ danh sách SAP Name thành các dòng Armature Master bên dưới"
          >
            <ArrowDownToLine className="w-4 h-4" />
            <span>XUẤT RA ARMATURE LIST BÊN DƯỚI &darr;</span>
          </button>
        </div>
      </div>

      {/* Paste Modal */}
      {pasteModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
          <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-2xl w-full shadow-2xl border border-slate-200 dark:border-slate-800 overflow-hidden">
            <div className="p-4 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <ClipboardPaste className="w-5 h-5 text-indigo-600 dark:text-indigo-400" />
                <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100 uppercase">
                  Dán Danh Sách Hàng Loạt Mã SAP Name
                </h3>
              </div>
              <button
                onClick={() => setPasteModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 text-lg cursor-pointer"
              >
                &times;
              </button>
            </div>

            <div className="p-4 space-y-3">
              <p className="text-xs text-slate-500">
                Copy danh sách cột <strong>SAP Name</strong> từ file Excel hoặc văn bản của bạn rồi dán trực tiếp vào ô bên dưới (mỗi mã SAP trên một dòng):
              </p>

              <textarea
                rows={9}
                value={pasteText}
                onChange={(e) => setPasteText(e.target.value)}
                placeholder={`A4010320608F\nA401032060PF\nA4010320708T\nA401032070PT\nA5051220602A\nV5518319100B`}
                className="w-full p-3 font-mono text-xs border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
              />

              <div className="flex items-center justify-between text-xs text-slate-400 pt-1">
                <span>Hỗ trợ cả dán 1 cột hoặc nhiều cột cách nhau bởi phím Tab từ Excel.</span>
                <span className="font-mono font-semibold text-blue-600 dark:text-blue-400">
                  {pasteText.split(/\r?\n/).filter((l) => l.trim().length > 0).length} dòng
                </span>
              </div>
            </div>

            <div className="p-4 bg-slate-50 dark:bg-slate-800 border-t border-slate-200 dark:border-slate-700 flex justify-end gap-2">
              <button
                onClick={() => setPasteModalOpen(false)}
                className="px-4 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-700 text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 cursor-pointer"
              >
                Hủy
              </button>
              <button
                onClick={handleApplyPasteText}
                disabled={!pasteText.trim()}
                className={`px-5 py-1.5 text-xs font-bold rounded-lg text-white transition-all cursor-pointer ${
                  pasteText.trim()
                    ? 'bg-blue-600 hover:bg-blue-500 shadow-sm'
                    : 'bg-slate-400 opacity-60 cursor-not-allowed'
                }`}
              >
                Đưa Vào Bảng Nhập
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
