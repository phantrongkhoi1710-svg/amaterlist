import React, { useState, useEffect, useMemo } from 'react';
import { Header } from './components/Header';
import { FileUploader } from './components/FileUploader';
import { MasterTable } from './components/MasterTable';
import { MappingViewer } from './components/MappingViewer';
import { ImportLogView } from './components/ImportLogView';
import { CatalogViewer } from './components/CatalogViewer';
import { SapBatchInputTable } from './components/SapBatchInputTable';
import { HelpModal } from './components/HelpModal';
import { AliasConfigModal } from './components/AliasConfigModal';
import { GithubDeployModal } from './components/GithubDeployModal';
import { EditRowModal } from './components/EditRowModal';
import { OutlookEmailModal } from './components/OutlookEmailModal';
import { OutlookEmailContext, RowEditChange } from './utils/outlookMailer';
import {
  DEFAULT_MASTER_COLUMNS,
  DEFAULT_HEADER_ALIASES,
} from './utils/headerNormalizer';
import { parseSourceFile } from './utils/excelParser';
import { exportToExcel, exportToCSV } from './utils/excelExporter';
import {
  generateSampleExcelFile,
  generateImageArmatureSampleRows,
} from './utils/sampleData';
import {
  DEFAULT_SAMPLE_CATALOG_ITEMS,
  autoFillMasterFromCatalog,
} from './utils/catalogManager';
import {
  MasterRowData,
  SourceFileInfo,
  ImportLogRow,
  ImportMode,
  CatalogItem,
  CatalogProfile,
} from './types';
import { CheckCircle2, AlertTriangle, XCircle, Info, FileUp, X, BookOpen } from 'lucide-react';

const STORAGE_KEY_MASTER_ROWS = 'armature_tool_master_rows';
const STORAGE_KEY_LOGS = 'armature_tool_logs';
const STORAGE_KEY_HEADERS = 'armature_tool_headers';
const STORAGE_KEY_ALIASES = 'armature_tool_aliases';
const STORAGE_KEY_CATALOGS = 'armature_tool_catalogs';
const STORAGE_KEY_ACTIVE_CATALOG = 'armature_tool_active_catalog';

export default function App() {
  const [masterHeaders, setMasterHeaders] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_HEADERS);
      if (saved) {
        const parsed = JSON.parse(saved);
        // Exclude any legacy 'SAP CODE' or 'SAP' from masterHeaders as SAP Name belongs strictly to Catalog
        return parsed.filter(
          (h: string) =>
            h.toUpperCase().replace(/\s+/g, '') !== 'SAPCODE' &&
            h.toUpperCase() !== 'SAP' &&
            h.toUpperCase() !== 'SAP NAME'
        );
      }
      return DEFAULT_MASTER_COLUMNS;
    } catch {
      return DEFAULT_MASTER_COLUMNS;
    }
  });

  const [aliases, setAliases] = useState<Record<string, string[]>>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_ALIASES);
      return saved ? JSON.parse(saved) : DEFAULT_HEADER_ALIASES;
    } catch {
      return DEFAULT_HEADER_ALIASES;
    }
  });

  const [masterRows, setMasterRows] = useState<MasterRowData[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_MASTER_ROWS);
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const [logs, setLogs] = useState<ImportLogRow[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_LOGS);
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const [files, setFiles] = useState<SourceFileInfo[]>([]);
  const [importMode, setImportMode] = useState<ImportMode>('replace');
  const [isProcessing, setIsProcessing] = useState(false);
  const [isFileUploaderOpen, setIsFileUploaderOpen] = useState(false);
  const [isLogsModalOpen, setIsLogsModalOpen] = useState(false);
  const [isCatalogModalOpen, setIsCatalogModalOpen] = useState(false);

  // Catalog profiles state
  const [catalogProfiles, setCatalogProfiles] = useState<CatalogProfile[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_CATALOGS);
      if (saved) return JSON.parse(saved);
    } catch (e) {
      console.warn(e);
    }
    return [
      {
        id: 'cat-profile-default',
        shipName: 'Tàu Mẫu (Sample Ship)',
        versionName: 'Rev 01',
        uploadedAt: new Date().toLocaleString('vi-VN'),
        sourceFileName: 'Catalog_Armature_Goc.xlsx',
        items: DEFAULT_SAMPLE_CATALOG_ITEMS,
      },
    ];
  });

  const [activeProfileId, setActiveProfileId] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_ACTIVE_CATALOG);
      if (saved) return saved;
    } catch {}
    return 'cat-profile-default';
  });

  const activeProfile = useMemo(
    () => catalogProfiles.find((p) => p.id === activeProfileId) || catalogProfiles[0],
    [catalogProfiles, activeProfileId]
  );
  const activeCatalogItems = useMemo(
    () => (activeProfile ? activeProfile.items : []),
    [activeProfile]
  );

  const [isHelpOpen, setIsHelpOpen] = useState(false);
  const [isAliasModalOpen, setIsAliasModalOpen] = useState(false);
  const [isGithubModalOpen, setIsGithubModalOpen] = useState(false);
  const [editingRow, setEditingRow] = useState<MasterRowData | null>(null);
  const [isOutlookMailOpen, setIsOutlookMailOpen] = useState(false);
  const [outlookContext, setOutlookContext] = useState<OutlookEmailContext | null>(null);
  const [toastMessage, setToastMessage] = useState<{
    type: 'success' | 'warning' | 'error' | 'info';
    text: string;
  } | null>(null);

  // Sync to local storage
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_MASTER_ROWS, JSON.stringify(masterRows));
    } catch (e) {
      console.warn('Storage limit reached for masterRows', e);
    }
  }, [masterRows]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_LOGS, JSON.stringify(logs));
    } catch (e) {
      console.warn('Storage limit reached for logs', e);
    }
  }, [logs]);

  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY_CATALOGS, JSON.stringify(catalogProfiles));
    } catch (e) {
      console.warn('Storage limit reached for catalogProfiles', e);
    }
  }, [catalogProfiles]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY_ACTIVE_CATALOG, activeProfileId);
  }, [activeProfileId]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY_HEADERS, JSON.stringify(masterHeaders));
  }, [masterHeaders]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY_ALIASES, JSON.stringify(aliases));
  }, [aliases]);

  // Handlers for Catalog
  const handleSaveCatalogProfile = (profile: CatalogProfile) => {
    setCatalogProfiles((prev) => {
      const index = prev.findIndex((p) => p.id === profile.id);
      if (index >= 0) {
        const updated = [...prev];
        updated[index] = profile;
        return updated;
      }
      return [...prev, profile];
    });
    setActiveProfileId(profile.id);
    showToast('success', `Đã lưu Catalog cho ${profile.shipName} (${profile.items.length} mã van)!`);
  };

  const handleDeleteCatalogProfile = (id: string) => {
    setCatalogProfiles((prev) => {
      const filtered = prev.filter((p) => p.id !== id);
      if (filtered.length === 0) {
        return [
          {
            id: 'cat-profile-default',
            shipName: 'Tàu Mẫu (Sample Ship)',
            versionName: 'Rev 01',
            uploadedAt: new Date().toLocaleString('vi-VN'),
            sourceFileName: 'Catalog_Armature_Goc.xlsx',
            items: DEFAULT_SAMPLE_CATALOG_ITEMS,
          },
        ];
      }
      return filtered;
    });
    setActiveProfileId((prevId) => {
      if (prevId === id) {
        const remaining = catalogProfiles.filter((p) => p.id !== id);
        return remaining[0]?.id || 'cat-profile-default';
      }
      return prevId;
    });
    showToast('info', 'Đã xoá Catalog.');
  };

  const handleAutoFillMaster = (catalogItems: CatalogItem[]) => {
    if (masterRows.length === 0) {
      showToast('warning', 'Bảng Master hiện chưa có dữ liệu để đối chiếu thông số từ Catalog.');
      return;
    }
    const result = autoFillMasterFromCatalog(masterRows, catalogItems);
    if (result.matchedCount === 0) {
      showToast('warning', 'Không tìm thấy dòng nào trong Master khớp với TAG/Model/Normale Nr của Catalog.');
      return;
    }
    setMasterRows(result.updatedRows);
    showToast(
      'success',
      `Đã đối chiếu thành công ${result.matchedCount} mã van và bổ sung ${result.fieldsUpdatedCount} trường thông số kỹ thuật (Thân van, áp suất, tiêu chuẩn...) vào Master!`
    );
  };

  // Show toast notification
  const showToast = (type: 'success' | 'warning' | 'error' | 'info', text: string) => {
    setToastMessage({ type, text });
    setTimeout(() => {
      setToastMessage(null);
    }, 4500);
  };

  // Handle Apply from SAP Batch Input Table -> Master Table
  const handleApplySapBatch = (rows: MasterRowData[], mode: 'replace' | 'append') => {
    if (mode === 'replace') {
      setMasterRows(rows);
      showToast('success', `Đã xuất thành công ${rows.length} hàng từ danh sách SAP Name ra bảng Armature List!`);
    } else {
      setMasterRows((prev) => [...prev, ...rows]);
      showToast('success', `Đã thêm nối tiếp ${rows.length} hàng từ danh sách SAP Name vào bảng Armature List!`);
    }

    // Scroll to Master Table
    const el = document.getElementById('master-table-container');
    if (el) {
      el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  };

  // Direct Excel Export from SAP Batch Input Table
  const handleExportDirectExcelFromSap = async (rows: MasterRowData[]) => {
    try {
      await exportToExcel(
        masterHeaders,
        rows,
        logs,
        `Armature_List_From_SAP_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
      showToast('success', `Đã xuất file Excel Armature List cho ${rows.length} mã SAP thành công!`);
    } catch (err: any) {
      showToast('error', `Lỗi xuất Excel: ${err.message}`);
    }
  };

  // Upload & process files
  const handleUploadFiles = async (newFiles: File[], mode: ImportMode) => {
    if (newFiles.length === 0) return;
    setIsProcessing(true);

    try {
      let currentMasterData = mode === 'replace' ? [] : [...masterRows];
      const newFilesInfo: SourceFileInfo[] = [];
      const newLogRows: ImportLogRow[] = [];
      let totalImportedThisBatch = 0;

      for (const file of newFiles) {
        const result = await parseSourceFile(file, masterHeaders, aliases);
        newFilesInfo.push(result.fileInfo);
        newLogRows.push(result.logRow);

        if (result.importedData.length > 0) {
          currentMasterData = [...currentMasterData, ...result.importedData];
          totalImportedThisBatch += result.importedData.length;
        }
      }

      setMasterRows(currentMasterData);
      setFiles((prev) => (mode === 'replace' ? newFilesInfo : [...prev, ...newFilesInfo]));
      setLogs((prev) => [...newLogRows, ...prev]);

      showToast(
        'success',
        `Đã xử lý xong ${newFiles.length} file. Đã nhập ${totalImportedThisBatch} dòng dữ liệu vào bảng Master!`
      );
    } catch (err: any) {
      showToast('error', `Lỗi xử lý file: ${err.message || 'Không xác định'}`);
    } finally {
      setIsProcessing(false);
    }
  };

  // Load sample demo files
  const handleLoadSampleData = async () => {
    setIsProcessing(true);
    try {
      const file1 = generateSampleExcelFile('Armature_Vendor_Emerson', 'vendor_a');
      const file2 = generateSampleExcelFile('Armature_Vendor_Kitz', 'vendor_b');
      const file3 = generateSampleExcelFile('Armature_Vendor_Cameron', 'subsea_c');

      await handleUploadFiles([file1, file2, file3], 'replace');
      showToast(
        'info',
        'Đã nạp 3 file mẫu test thành công với các bí danh cột PO NO, ACTUATOR, PRESSURE, BODY...'
      );
    } catch (err: any) {
      showToast('error', `Lỗi tạo dữ liệu mẫu: ${err.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  // Clear Master Data
  const handleClearMaster = () => {
    if (window.confirm('Bạn có chắc chắn muốn xoá toàn bộ dữ liệu trong bảng Master (Total)?')) {
      setMasterRows([]);
      showToast('info', 'Đã làm trống bảng Master.');
    }
  };

  // Delete individual row
  const handleDeleteRow = (rowId: string) => {
    setMasterRows((prev) => prev.filter((r) => r._id !== rowId));
  };

  // Update individual row
  const handleUpdateRow = (updatedRow: MasterRowData) => {
    setMasterRows((prev) =>
      prev.map((r) => (r._id === updatedRow._id ? updatedRow : r))
    );
    showToast(
      'success',
      `Đã cập nhật van [${updatedRow.TAG || updatedRow['STD DRW NORMALE N°'] || 'vật tư'}]!`
    );
  };

  // Add new rows
  const handleAddRows = (newRows: MasterRowData[]) => {
    setMasterRows((prev) => [...newRows, ...prev]);
  };

  // Load the 3 sample rows from user's image
  const handleLoadImageSample = () => {
    const sampleRows = generateImageArmatureSampleRows();
    setMasterRows(sampleRows);
    showToast(
      'success',
      'Đã nạp 3 dòng vật tư Armature theo cấu trúc ảnh mẫu!'
    );
  };

  // Add new manual valve row
  const handleAddNewRow = (newRowData?: MasterRowData) => {
    const newId = `row-manual-${Date.now()}`;
    const row: MasterRowData = newRowData || {
      _id: newId,
      _sourceFile: 'Thủ công',
      _sourceSheet: 'Nhập tay',
      _importedAt: new Date().toLocaleString('vi-VN'),
      TAG: `V-${Math.floor(100 + Math.random() * 900)}`,
      'ACTUATOR TAG': '',
      'P.O. NUMBER': '',
      'SUPPLIER': '',
      'DESTINATION (YARD)': 'Yard A',
    };
    setMasterRows((prev) => [row, ...prev]);
    showToast('success', `Đã thêm van mới (${row.TAG}) vào bảng Armature!`);
  };

  // Edit individual row
  const handleStartEditRow = (row: MasterRowData) => {
    setEditingRow(row);
  };

  // Save edited row & auto trigger Outlook draft
  const handleSaveRow = (
    updatedRow: MasterRowData,
    changes: RowEditChange[],
    autoComposeMail: boolean
  ) => {
    setMasterRows((prev) =>
      prev.map((r) => (r._id === updatedRow._id ? updatedRow : r))
    );
    setEditingRow(null);

    if (changes.length > 0) {
      showToast(
        'success',
        `Đã lưu cập nhật van ${updatedRow.TAG || 'được chọn'} (${changes.length} thay đổi)!`
      );
    } else {
      showToast('info', 'Đã lưu thông tin.');
    }

    // Automatically trigger Outlook draft if requested
    if (autoComposeMail) {
      setOutlookContext({
        type: 'row_edited',
        rowTag: updatedRow.TAG,
        supplier: updatedRow.SUPPLIER,
        sourceFile: updatedRow._sourceFile,
        changes,
        rowData: updatedRow,
        masterHeaders,
        allRows: masterRows.map((r) => (r._id === updatedRow._id ? updatedRow : r)),
        logs,
      });
      setIsOutlookMailOpen(true);
    }
  };

  // Mail single row
  const handleMailRow = (row: MasterRowData) => {
    setOutlookContext({
      type: 'row_edited',
      rowTag: row.TAG,
      supplier: row.SUPPLIER,
      sourceFile: row._sourceFile,
      changes: [],
      rowData: row,
      masterHeaders,
      allRows: masterRows,
      logs,
    });
    setIsOutlookMailOpen(true);
  };

  // Batch Mail Outlook (from header or table toolbar)
  const handleOpenBatchMail = () => {
    const uniqueSuppliers = Array.from(
      new Set(masterRows.map((r) => r.SUPPLIER).filter(Boolean))
    ) as string[];

    setOutlookContext({
      type: 'batch_report',
      summaryStats: {
        totalRows: masterRows.length,
        totalFiles: files.length,
        suppliers: uniqueSuppliers,
      },
      masterHeaders,
      allRows: masterRows,
      logs,
    });
    setIsOutlookMailOpen(true);
  };

  // Remove single file
  const handleRemoveFile = (fileId: string) => {
    setFiles((prev) => prev.filter((f) => f.id !== fileId));
  };

  // Clear all files
  const handleClearAllFiles = () => {
    setFiles([]);
  };

  // Clear all logs
  const handleClearLogs = () => {
    if (window.confirm('Xoá toàn bộ nhật ký import?')) {
      setLogs([]);
    }
  };

  // Export Excel
  const handleExportExcel = async () => {
    if (masterRows.length === 0) {
      alert('Không có dữ liệu trong bảng Master để xuất!');
      return;
    }
    try {
      await exportToExcel(masterHeaders, masterRows, logs);
      showToast('success', 'Đã tạo và tải file Excel (Total + Import_Log) với màu tiêu đề và cách dòng rộng rãi!');
    } catch (err: any) {
      showToast('error', `Lỗi xuất file Excel: ${err.message}`);
    }
  };

  // Export CSV
  const handleExportCSV = () => {
    if (masterRows.length === 0) {
      alert('Không có dữ liệu trong bảng Master để xuất!');
      return;
    }
    exportToCSV(masterHeaders, masterRows);
    showToast('success', 'Đã tải file CSV của bảng Master thành công!');
  };

  // Active catalog count
  const activeCatalogCount = useMemo(() => {
    const activeProf =
      catalogProfiles.find((p) => p.id === activeProfileId) || catalogProfiles[0];
    return activeProf?.items?.length || 0;
  }, [catalogProfiles, activeProfileId]);

  return (
    <div className="min-h-screen bg-slate-100/70 dark:bg-slate-950 text-slate-900 dark:text-slate-100 flex flex-col font-sans antialiased">
      {/* Top App Header */}
      <Header
        totalRows={masterRows.length}
        totalFiles={files.length}
        catalogCount={activeCatalogCount}
        activeShipName={activeProfile?.shipName || 'Tàu Mẫu'}
        onExportExcel={handleExportExcel}
        onExportCSV={handleExportCSV}
        onLoadSampleData={handleLoadSampleData}
        onClearData={handleClearMaster}
        onOpenAliases={() => setIsAliasModalOpen(true)}
        onOpenHelp={() => setIsHelpOpen(true)}
        onOpenGithubDeploy={() => setIsGithubModalOpen(true)}
        onOpenOutlookMail={handleOpenBatchMail}
        onOpenCatalogManager={() => setIsCatalogModalOpen(true)}
      />

      {/* Main Container - Full Width Expanded */}
      <main className="flex-1 w-full px-3 sm:px-5 lg:px-6 py-3 space-y-4">
        {/* Toast Notification */}
        {toastMessage && (
          <div className="animate-in fade-in slide-in-from-top-2 duration-200">
            <div
              className={`p-3 rounded-xl border flex items-center justify-between text-xs font-medium shadow-md ${
                toastMessage.type === 'success'
                  ? 'bg-emerald-50 text-emerald-800 border-emerald-200 dark:bg-emerald-950/60 dark:text-emerald-200 dark:border-emerald-800'
                  : toastMessage.type === 'error'
                  ? 'bg-rose-50 text-rose-800 border-rose-200 dark:bg-rose-950/60 dark:text-rose-200 dark:border-rose-800'
                  : toastMessage.type === 'warning'
                  ? 'bg-amber-50 text-amber-800 border-amber-200 dark:bg-amber-950/60 dark:text-amber-200 dark:border-amber-800'
                  : 'bg-blue-50 text-blue-800 border-blue-200 dark:bg-blue-950/60 dark:text-blue-200 dark:border-blue-800'
              }`}
            >
              <div className="flex items-center gap-2">
                {toastMessage.type === 'success' && <CheckCircle2 className="w-4 h-4 text-emerald-500" />}
                {toastMessage.type === 'error' && <XCircle className="w-4 h-4 text-rose-500" />}
                {toastMessage.type === 'warning' && <AlertTriangle className="w-4 h-4 text-amber-500" />}
                {toastMessage.type === 'info' && <Info className="w-4 h-4 text-blue-500" />}
                <span>{toastMessage.text}</span>
              </div>
              <button
                onClick={() => setToastMessage(null)}
                className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 ml-4 cursor-pointer"
              >
                &times;
              </button>
            </div>
          </div>
        )}

        {/* 1. BẢNG NHẬP HÀNG LOẠT SAP NAME -> TỰ ĐỘNG TRA CỨU CATALOG -> XUẤT RA ARMATURE LIST */}
        <SapBatchInputTable
          catalogItems={activeCatalogItems}
          activeShipName={activeProfile?.shipName || 'Tàu Mẫu'}
          onApplyToMaster={handleApplySapBatch}
          onExportDirectExcel={handleExportDirectExcelFromSap}
          onOpenCatalogSettings={() => setIsCatalogModalOpen(true)}
          currentMasterCount={masterRows.length}
        />

        {/* Embedded / Collapsible File Uploader (VBA Macro Parser) */}
        {isFileUploaderOpen && (
          <div className="bg-white dark:bg-slate-900 rounded-xl border border-blue-200 dark:border-blue-900/60 p-4 shadow-sm animate-in fade-in duration-200">
            <div className="flex items-center justify-between pb-3 mb-3 border-b border-slate-200 dark:border-slate-800">
              <div className="flex items-center gap-2">
                <FileUp className="w-4 h-4 text-blue-600 dark:text-blue-400" />
                <h3 className="text-xs font-bold text-slate-900 dark:text-slate-100 uppercase tracking-wider">
                  Nhập File Nguồn Excel (VBA Macro Parser)
                </h3>
              </div>
              <button
                onClick={() => setIsFileUploaderOpen(false)}
                className="text-xs text-slate-500 hover:text-slate-800 dark:hover:text-slate-200 px-2.5 py-1 rounded-md bg-slate-100 dark:bg-slate-800 cursor-pointer font-medium"
              >
                Thu gọn ▲
              </button>
            </div>
            <FileUploader
              files={files}
              onUploadFiles={handleUploadFiles}
              onRemoveFile={handleRemoveFile}
              onClearAllFiles={handleClearAllFiles}
              isProcessing={isProcessing}
              importMode={importMode}
              setImportMode={setImportMode}
            />
          </div>
        )}

        {/* 2. BẢNG ARMATURE LIST (MASTER TABLE ĐẦY ĐỦ 2 TẦNG CHUẨN PIPE SPECIFICATION) */}
        <div id="master-table-container">
          <MasterTable
            data={masterRows}
            masterHeaders={masterHeaders}
            onDeleteRow={handleDeleteRow}
            onEditRow={handleStartEditRow}
            onUpdateRow={handleUpdateRow}
            onAddRow={handleAddNewRow}
            onLoadImageSample={handleLoadImageSample}
            onMailRow={handleMailRow}
            onOpenBatchMail={handleOpenBatchMail}
            onExportExcel={handleExportExcel}
            onExportCSV={handleExportCSV}
            onToggleFileUploader={() => setIsFileUploaderOpen((prev) => !prev)}
            isFileUploaderOpen={isFileUploaderOpen}
            totalFilesCount={files.length}
            onOpenAliases={() => setIsAliasModalOpen(true)}
            onOpenLogs={() => setIsLogsModalOpen(true)}
            logsCount={logs.length}
          />
        </div>
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 py-4 text-center text-xs text-slate-400">
        <div className="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-2">
          <span>Armature Import Web Tool &bull; Chuyển đổi từ VBA Excel Macro &amp; Tra cứu Catalog</span>
          <div className="flex items-center gap-3">
            <button
              onClick={() => setIsCatalogModalOpen(true)}
              className="text-amber-500 hover:underline cursor-pointer"
            >
              Quản lý Catalog nền ({activeCatalogCount} mã)
            </button>
            <span className="text-slate-300 dark:text-slate-700">&bull;</span>
            <button
              onClick={() => setIsHelpOpen(true)}
              className="text-blue-500 hover:underline cursor-pointer"
            >
              Tài liệu & Đối chiếu thuật toán VBA
            </button>
            <span className="text-slate-300 dark:text-slate-700">&bull;</span>
            <button
              onClick={() => setIsGithubModalOpen(true)}
              className="text-slate-600 dark:text-slate-300 hover:text-blue-500 flex items-center gap-1 cursor-pointer"
            >
              Deploy GitHub Pages
            </button>
          </div>
        </div>
      </footer>

      {/* MODALS */}
      {/* Hidden Catalog Manager Modal (Runs in background, opened on demand) */}
      {isCatalogModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/70 backdrop-blur-xs">
          <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-6xl w-full max-h-[92vh] flex flex-col shadow-2xl border border-slate-200 dark:border-slate-800 overflow-hidden">
            <div className="p-4 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between bg-slate-900 text-white">
              <div className="flex items-center gap-2">
                <BookOpen className="w-5 h-5 text-amber-400" />
                <h3 className="text-sm font-bold tracking-wide">
                  QUẢN LÝ CATALOG THIẾT BỊ NỀN THEO TÀU (BACKGROUND DATABASE)
                </h3>
              </div>
              <button
                onClick={() => setIsCatalogModalOpen(false)}
                className="text-slate-400 hover:text-white p-1 rounded-lg text-lg cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-4 overflow-y-auto flex-1 bg-slate-50 dark:bg-slate-950">
              <CatalogViewer
                catalogProfiles={catalogProfiles}
                activeProfileId={activeProfileId}
                setActiveProfileId={setActiveProfileId}
                onSaveProfile={handleSaveCatalogProfile}
                onDeleteProfile={handleDeleteCatalogProfile}
              />
            </div>
          </div>
        </div>
      )}

      {/* Import Log Modal */}
      {isLogsModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
          <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-5xl w-full max-h-[90vh] flex flex-col shadow-2xl border border-slate-200 dark:border-slate-800 overflow-hidden">
            <div className="p-4 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
              <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100 flex items-center gap-2">
                <span>Nhật Ký Quá Trình Import (Import_Log)</span>
                <span className="text-xs font-normal text-slate-500">({logs.length} sự kiện)</span>
              </h3>
              <button
                onClick={() => setIsLogsModalOpen(false)}
                className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 p-1.5 rounded-lg text-lg leading-none cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-4 overflow-y-auto flex-1">
              <ImportLogView
                logs={logs}
                onClearLogs={handleClearLogs}
                onExportExcel={handleExportExcel}
              />
            </div>
          </div>
        </div>
      )}

      <HelpModal isOpen={isHelpOpen} onClose={() => setIsHelpOpen(false)} />
      <AliasConfigModal
        isOpen={isAliasModalOpen}
        onClose={() => setIsAliasModalOpen(false)}
        aliases={aliases}
        setAliases={setAliases}
        masterHeaders={masterHeaders}
      />
      <GithubDeployModal
        isOpen={isGithubModalOpen}
        onClose={() => setIsGithubModalOpen(false)}
      />
      <EditRowModal
        isOpen={Boolean(editingRow)}
        onClose={() => setEditingRow(null)}
        row={editingRow}
        masterHeaders={masterHeaders}
        onSave={handleSaveRow}
      />
      <OutlookEmailModal
        isOpen={isOutlookMailOpen}
        onClose={() => setIsOutlookMailOpen(false)}
        context={outlookContext}
      />
    </div>
  );
}
