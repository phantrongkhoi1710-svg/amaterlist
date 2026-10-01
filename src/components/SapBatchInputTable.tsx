import React, { useState, useMemo, useRef } from 'react';
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
  Upload,
  Check,
} from 'lucide-react';
import * as XLSX from 'xlsx';
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
  const [uploadNotification, setUploadNotification] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement | null>(null);

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

  // Build MasterRowData objects with the complete new 26-column specification structure
  const buildMasterRows = (): MasterRowData[] => {
    return evaluatedEntries
      .filter((e) => e.trimmed)
      .map((e, idx) => {
        const specFields = e.armatureFields || {};
        const rowId = `sap-import-${Date.now()}-${idx}-${Math.random().toString(36).substring(2, 6)}`;

        const newRow: MasterRowData = {
          _id: rowId,
          _sourceFile: `SAP_Batch_Import`,
          _importedAt: new Date().toLocaleString('vi-VN'),
          _revisionStatus: 'normal',
          _rowNumber: idx + 1,
          TAG: e.tag?.trim() || `V-${String(idx + 1).padStart(3, '0')}`,
          'ACTUATOR TAG': '',
          'P.O. NUMBER': e.poNumber?.trim() || commonPo.trim() || '',
          'SUPPLIER': e.supplier?.trim() || specFields['SUPPLIER'] || commonSupplier.trim() || '',
          'DESTINATION (YARD)': e.yard?.trim() || commonYard.trim() || 'Yard A',
          'SAP CODE': specFields['SAP CODE'] || e.trimmed,
          'STD DRW NORMALE N°': specFields['STD DRW NORMALE N°'] || '',
          'EXECUTION': specFields['EXECUTION'] || '',
          'NRF N°': specFields['NRF N°'] || '',
          'SFI': specFields['SFI'] || '',
          'DESCRIPTION': specFields['DESCRIPTION'] || '',
          'SIZE': specFields['SIZE'] || '',
          'CONNECTION': specFields['CONNECTION'] || '',
          'PRESSURE RATING': specFields['PRESSURE RATING'] || '',
          'HOUSING /BODY': specFields['HOUSING /BODY'] || '',
          'TYPE': specFields['TYPE'] || '',
          'PIPE CLASS': specFields['PIPE CLASS'] || '',
          'CLASS CERTIFICATE': specFields['CLASS CERTIFICATE'] || '',
          'REMARKS': specFields['REMARKS'] || '',
          'SIGN TYPE': specFields['SIGN TYPE'] || '',
          'SIGN TEXT': '',
          'INPUT - SIGN TEXT': '',
          'REV. HIS.': '',
          'DATE': '',
          'REV. DESCRIPTION': '',
          'SIGNATURE': '',
          ...specFields,
        };

        return newRow;
      });
  };

  // Upload Excel or CSV file with SAP Names or Catalog
  const handleFileUpload = async (file: File) => {
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
        setUploadNotification('Excel file does not contain any sheets.');
        return;
      }

      const ws = workbook.Sheets[workbook.SheetNames[0]];
      const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
      if (rawRows.length === 0) {
        setUploadNotification('Excel file contains no data.');
        return;
      }

      // Detect header row and column positions
      let headerRowIdx = 0;
      let sapColIdx = -1;
      let tagColIdx = -1;
      let yardColIdx = -1;
      let supplierColIdx = -1;
      let poColIdx = -1;

      for (let r = 0; r < Math.min(rawRows.length, 20); r++) {
        const row = rawRows[r];
        if (!Array.isArray(row)) continue;
        for (let c = 0; c < row.length; c++) {
          const val = String(row[c] || '').trim().toUpperCase().replace(/[\r\n\t_.\-]/g, ' ');
          if (
            val.includes('SAP') ||
            val === 'NAME' ||
            val.includes('MATERIAL') ||
            val.includes('PART NUMBER') ||
            val.includes('VALVE CODE')
          ) {
            if (sapColIdx === -1) {
              sapColIdx = c;
              headerRowIdx = r;
            }
          }
          if (val === 'TAG' || val.includes('TAG NO') || val.includes('TAG NUMBER') || val.includes('VALVE TAG')) {
            tagColIdx = c;
          }
          if (val.includes('YARD') || val.includes('DESTINATION')) {
            yardColIdx = c;
          }
          if (val.includes('SUPPLIER') || val.includes('MANUFACTURER') || val.includes('MAKER')) {
            supplierColIdx = c;
          }
          if (val.includes('PO') || val.includes('PURCHASE')) {
            poColIdx = c;
          }
        }
        if (sapColIdx !== -1) break;
      }

      if (sapColIdx === -1) {
        sapColIdx = 0;
        headerRowIdx = 0;
      }

      const newEntries: SapInputEntry[] = [];
      const startR =
        headerRowIdx === 0 &&
        !String(rawRows[0][sapColIdx]).toUpperCase().includes('SAP') &&
        !String(rawRows[0][sapColIdx]).toUpperCase().includes('NAME')
          ? 0
          : headerRowIdx + 1;

      for (let r = startR; r < rawRows.length; r++) {
        const row = rawRows[r];
        if (!Array.isArray(row)) continue;
        const code = String(row[sapColIdx] || '').trim();
        if (!code || code.toUpperCase() === 'SAP NAME' || code.toUpperCase() === 'SAP CODE') continue;

        const tag =
          tagColIdx >= 0 && row[tagColIdx]
            ? String(row[tagColIdx]).trim()
            : `V-${String(newEntries.length + 1).padStart(2, '0')}`;
        const yard = yardColIdx >= 0 && row[yardColIdx] ? String(row[yardColIdx]).trim() : commonYard;
        const supplier = supplierColIdx >= 0 && row[supplierColIdx] ? String(row[supplierColIdx]).trim() : commonSupplier;
        const poNumber = poColIdx >= 0 && row[poColIdx] ? String(row[poColIdx]).trim() : commonPo;

        newEntries.push({
          id: `up-${Date.now()}-${r}`,
          sapName: code,
          tag,
          yard,
          supplier,
          poNumber,
        });
      }

      if (newEntries.length > 0) {
        setEntries(newEntries);
        setUploadNotification(`Successfully loaded ${newEntries.length} SAP codes from file [${file.name}]!`);
        setTimeout(() => setUploadNotification(null), 6000);
      } else {
        setUploadNotification(`No SAP codes could be extracted from file [${file.name}].`);
      }
    } catch (e: any) {
      console.error(e);
      setUploadNotification(`Error reading file: ${e.message}`);
    }
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
                SAP Batch Entry &rarr; Export Armature List
              </h2>
              <span className="text-[11px] px-2 py-0.5 rounded-full bg-blue-500/30 border border-blue-400/40 text-blue-200 font-mono">
                {totalEntered} entries
              </span>
            </div>
            <p className="text-xs text-slate-300 mt-0.5 flex items-center gap-2">
              <span className="font-semibold text-amber-300 bg-amber-950/60 px-2 py-0.5 rounded border border-amber-500/30">
                Catalog: {activeShipName} ({catalogItems.length} items)
              </span>
            </p>
          </div>
        </div>

        {/* Top Control Buttons */}
        <div className="flex flex-wrap items-center gap-2">
          {/* Upload Excel Button */}
          <button
            id="btn-upload-sap-file"
            type="button"
            onClick={() => fileInputRef.current?.click()}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-emerald-600 hover:bg-emerald-500 text-white shadow-sm transition-all cursor-pointer active:scale-95"
            title="Upload Excel or CSV file containing SAP Names or Catalog"
          >
            <Upload className="w-3.5 h-3.5" />
            <span>Upload SAP Excel File</span>
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.xlsm,.csv"
            onChange={(e) => {
              if (e.target.files && e.target.files.length > 0) {
                handleFileUpload(e.target.files[0]);
                e.target.value = '';
              }
            }}
            className="hidden"
          />

          {/* Quick Paste Button */}
          <button
            id="btn-open-paste-modal"
            type="button"
            onClick={() => setPasteModalOpen(true)}
            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-indigo-600 hover:bg-indigo-500 text-white shadow-sm transition-all cursor-pointer active:scale-95"
            title="Open quick paste dialog for SAP codes"
          >
            <ClipboardPaste className="w-3.5 h-3.5" />
            <span>Quick Paste SAP List</span>
          </button>

          {/* Sample Codes */}
          <button
            id="btn-load-sample-sap"
            type="button"
            onClick={handleLoadSampleCodes}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors cursor-pointer"
            title="Load 6 sample SAP codes from Catalog"
          >
            <Sparkles className="w-3.5 h-3.5 text-amber-300" />
            <span className="hidden sm:inline">Load 6 Sample Codes</span>
          </button>

          {/* Add Row Button */}
          <button
            id="btn-add-sap-row"
            type="button"
            onClick={handleAddRow}
            className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors cursor-pointer"
            title="Add a new entry row"
          >
            <Plus className="w-3.5 h-3.5 text-blue-400" />
            <span>Add Row</span>
          </button>

          {/* Clear Button */}
          {entries.length > 0 && (
            <button
              id="btn-clear-sap-entries"
              type="button"
              onClick={handleClearAll}
              className="inline-flex items-center gap-1 px-2 py-1.5 rounded-lg text-xs font-medium text-rose-300 hover:text-white hover:bg-rose-900/40 transition-colors cursor-pointer"
              title="Clear all entered rows"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              <span>Clear All</span>
            </button>
          )}

          {/* Catalog Settings */}
          {onOpenCatalogSettings && (
            <button
              id="btn-open-catalog-settings"
              type="button"
              onClick={onOpenCatalogSettings}
              className="p-1.5 rounded-lg text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors cursor-pointer ml-1"
              title="Manage equipment catalog"
            >
              <Settings2 className="w-4 h-4" />
            </button>
          )}
        </div>
      </div>

      {/* Upload file notification banner */}
      {uploadNotification && (
        <div className="bg-emerald-600/15 dark:bg-emerald-950/40 text-emerald-800 dark:text-emerald-200 border-b border-emerald-300 dark:border-emerald-800 px-4 py-2.5 text-xs flex items-center justify-between animate-in fade-in duration-200">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0" />
            <span className="font-medium">{uploadNotification}</span>
          </div>
          <button
            type="button"
            onClick={() => setUploadNotification(null)}
            className="text-emerald-600 hover:text-emerald-900 dark:text-emerald-400 dark:hover:text-white text-base leading-none cursor-pointer px-1"
          >
            &times;
          </button>
        </div>
      )}

      {/* Batch defaults config accordion */}
      <div className="bg-slate-50 dark:bg-slate-850 border-b border-slate-200 dark:border-slate-800 px-4 py-2">
        <div className="flex flex-wrap items-center justify-between gap-2 text-xs">
          <button
            type="button"
            onClick={() => setIsBatchConfigOpen(!isBatchConfigOpen)}
            className="flex items-center gap-1.5 font-medium text-slate-600 dark:text-slate-300 hover:text-blue-600 cursor-pointer"
          >
            <span className="font-semibold">Batch Functional Settings (Yard, P.O, Supplier):</span>
            {isBatchConfigOpen ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
          </button>

          <div className="flex items-center gap-3">
            <span className="text-slate-500">Catalog Lookup Status:</span>
            <span className="font-bold text-emerald-600 dark:text-emerald-400 flex items-center gap-1">
              <CheckCircle2 className="w-3.5 h-3.5" /> {matchedCount} matched
            </span>
            {totalEntered - matchedCount > 0 && (
              <span className="font-bold text-amber-600 dark:text-amber-400 flex items-center gap-1">
                <AlertCircle className="w-3.5 h-3.5" /> {totalEntered - matchedCount} unmatched
              </span>
            )}
          </div>
        </div>

        {isBatchConfigOpen && (
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-3 pb-1 mt-2 border-t border-slate-200 dark:border-slate-800">
            <div>
              <label className="block text-[11px] font-semibold text-slate-500 uppercase mb-1">
                DEFAULT DESTINATION (YARD)
              </label>
              <input
                type="text"
                value={commonYard}
                onChange={(e) => setCommonYard(e.target.value)}
                placeholder="e.g. Yard A"
                className="w-full px-2.5 py-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100"
              />
            </div>
            <div>
              <label className="block text-[11px] font-semibold text-slate-500 uppercase mb-1">
                DEFAULT SUPPLIER
              </label>
              <input
                type="text"
                value={commonSupplier}
                onChange={(e) => setCommonSupplier(e.target.value)}
                placeholder="e.g. Econosto / Danfoss"
                className="w-full px-2.5 py-1.5 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-900 dark:text-slate-100"
              />
            </div>
            <div>
              <label className="block text-[11px] font-semibold text-slate-500 uppercase mb-1">
                DEFAULT P.O. NUMBER
              </label>
              <input
                type="text"
                value={commonPo}
                onChange={(e) => setCommonPo(e.target.value)}
                placeholder="e.g. PO-2026-09"
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
                SAP NAME (ENTER / PASTE CODE) *
              </th>
              <th className="py-2 px-3 min-w-[130px]">CATALOG MATCH</th>
              <th className="py-2 px-3 min-w-[180px]">CATALOG DESCRIPTION</th>
              <th className="py-2 px-3 min-w-[120px]">NORMALE N° / SPEC</th>
              <th className="py-2 px-3 min-w-[110px]">SIZE / RATING</th>
              <th className="py-2 px-3 min-w-[110px]">VALVE TAG</th>
              <th className="py-2 px-3 min-w-[100px]">DESTINATION (YARD)</th>
              <th className="py-2 px-2 w-12 text-center">DELETE</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100 dark:divide-slate-800 bg-white dark:bg-slate-900">
            {evaluatedEntries.length === 0 ? (
              <tr>
                <td colSpan={9} className="py-8 text-center text-slate-400">
                  <div className="max-w-md mx-auto space-y-2">
                    <p className="font-semibold text-slate-600 dark:text-slate-300">
                      No SAP codes entered yet.
                    </p>
                    <p className="text-xs text-slate-400">
                      Click <strong>"Upload SAP Excel File"</strong> or <strong>"Quick Paste SAP List"</strong> or <strong>"Load 6 Sample Codes"</strong> to start.
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
                        placeholder="Enter or paste SAP code..."
                        className="w-full px-2.5 py-1 text-xs font-mono font-bold text-blue-900 dark:text-blue-300 bg-slate-50 dark:bg-slate-950 border border-slate-300 dark:border-slate-700 rounded focus:border-blue-500 focus:ring-1 focus:ring-blue-500 outline-none"
                      />
                    </td>

                    {/* Catalog Match Status */}
                    <td className="py-1.5 px-3 whitespace-nowrap">
                      {!row.trimmed ? (
                        <span className="text-slate-400 text-[11px] italic">Not entered</span>
                      ) : row.isMatched ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-emerald-100 text-emerald-800 dark:bg-emerald-950/60 dark:text-emerald-300 border border-emerald-300 dark:border-emerald-800">
                          <CheckCircle2 className="w-3 h-3" /> Matched
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[11px] font-semibold bg-amber-100 text-amber-800 dark:bg-amber-950/60 dark:text-amber-300 border border-amber-300 dark:border-amber-800">
                          <AlertCircle className="w-3 h-3" /> Not in Catalog
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
                        placeholder="e.g. V-01"
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
                        type="button"
                        onClick={() => handleDeleteRow(row.id)}
                        className="p-1 text-slate-400 hover:text-rose-500 rounded transition-colors cursor-pointer"
                        title="Delete this entry"
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
            <span className="text-slate-600 dark:text-slate-300 font-semibold">Armature List Mode:</span>
            <select
              value={appendMode}
              onChange={(e) => setAppendMode(e.target.value as 'replace' | 'append')}
              className="px-2.5 py-1 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-200 cursor-pointer font-medium"
            >
              <option value="replace">Replace (Create New Armature List)</option>
              <option value="append">Append (Add to Existing List)</option>
            </select>
          </div>
          {currentMasterCount > 0 && (
            <span className="text-slate-400 text-[11px]">
              (Currently {currentMasterCount} rows in Master Table)
            </span>
          )}
        </div>

        {/* Big Action Buttons */}
        <div className="flex items-center gap-2.5">
          {/* Export directly to Excel */}
          <button
            id="btn-direct-export-excel"
            type="button"
            onClick={handleExportDirect}
            disabled={totalEntered === 0}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold transition-all shadow-sm ${
              totalEntered > 0
                ? 'bg-slate-800 hover:bg-slate-700 text-white border border-slate-700 cursor-pointer active:scale-95'
                : 'bg-slate-200 text-slate-400 dark:bg-slate-800 dark:text-slate-600 cursor-not-allowed'
            }`}
            title="Create and download Excel Armature Master Table directly from this list"
          >
            <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
            <span>Direct Export Excel (.xlsx)</span>
          </button>

          {/* Generate & Apply to Master Table */}
          <button
            id="btn-apply-sap-to-master"
            type="button"
            onClick={handleExportToArmatureList}
            disabled={totalEntered === 0}
            className={`inline-flex items-center gap-2 px-5 py-2 rounded-lg text-xs font-bold transition-all shadow-md ${
              totalEntered > 0
                ? 'bg-blue-600 hover:bg-blue-500 text-white cursor-pointer active:scale-95 ring-2 ring-blue-400/30'
                : 'bg-slate-200 text-slate-400 dark:bg-slate-800 dark:text-slate-600 cursor-not-allowed'
            }`}
            title="Convert all entered SAP items into Armature Master rows below"
          >
            <ArrowDownToLine className="w-4 h-4" />
            <span>EXPORT TO ARMATURE LIST BELOW &darr;</span>
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
                  Quick Paste SAP Code List
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setPasteModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 text-lg cursor-pointer"
              >
                &times;
              </button>
            </div>

            <div className="p-4 space-y-3">
              <p className="text-xs text-slate-500">
                Copy a column of <strong>SAP Names</strong> from Excel or Notepad and paste directly below (one per line):
              </p>

              <textarea
                rows={9}
                value={pasteText}
                onChange={(e) => setPasteText(e.target.value)}
                placeholder={`A4010320608F\nA401032060PF\nA4010320708T\nA401032070PT\nA5051220602A\nV5518319100B`}
                className="w-full p-3 font-mono text-xs border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 text-slate-900 dark:text-slate-100 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none"
              />

              <div className="flex items-center justify-between text-xs text-slate-400 pt-1">
                <span>Supports single or tab-delimited multi-column text copied from Excel.</span>
                <span className="font-mono font-semibold text-blue-600 dark:text-blue-400">
                  {pasteText.split(/\r?\n/).filter((l) => l.trim().length > 0).length} lines
                </span>
              </div>
            </div>

            <div className="p-4 bg-slate-50 dark:bg-slate-850 border-t border-slate-200 dark:border-slate-700 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setPasteModalOpen(false)}
                className="px-4 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-700 text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleApplyPasteText}
                disabled={!pasteText.trim()}
                className={`px-5 py-1.5 text-xs font-bold rounded-lg text-white transition-all cursor-pointer ${
                  pasteText.trim()
                    ? 'bg-blue-600 hover:bg-blue-500 shadow-sm'
                    : 'bg-slate-400 opacity-60 cursor-not-allowed'
                }`}
              >
                Import into Entry Table
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
