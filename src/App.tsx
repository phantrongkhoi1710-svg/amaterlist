import React, { useState, useEffect } from 'react';
import { Header } from './components/Header';
import { FileUploader } from './components/FileUploader';
import { MasterTable } from './components/MasterTable';
import { MappingViewer } from './components/MappingViewer';
import { ImportLogView } from './components/ImportLogView';
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
import { generateSampleExcelFile } from './utils/sampleData';
import {
  MasterRowData,
  SourceFileInfo,
  ImportLogRow,
  ImportMode,
} from './types';
import { CheckCircle2, AlertTriangle, XCircle, Info } from 'lucide-react';

const STORAGE_KEY_MASTER_ROWS = 'armature_tool_master_rows';
const STORAGE_KEY_LOGS = 'armature_tool_logs';
const STORAGE_KEY_HEADERS = 'armature_tool_headers';
const STORAGE_KEY_ALIASES = 'armature_tool_aliases';

export default function App() {
  const [masterHeaders, setMasterHeaders] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_HEADERS);
      return saved ? JSON.parse(saved) : DEFAULT_MASTER_COLUMNS;
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
  const [activeTab, setActiveTab] = useState<'table' | 'files' | 'mapping' | 'logs'>('table');

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
    localStorage.setItem(STORAGE_KEY_HEADERS, JSON.stringify(masterHeaders));
  }, [masterHeaders]);

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY_ALIASES, JSON.stringify(aliases));
  }, [aliases]);

  // Show toast notification
  const showToast = (type: 'success' | 'warning' | 'error' | 'info', text: string) => {
    setToastMessage({ type, text });
    setTimeout(() => {
      setToastMessage(null);
    }, 4500);
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
      setActiveTab('table');
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

  return (
    <div className="min-h-screen bg-slate-100/70 dark:bg-slate-950 text-slate-900 dark:text-slate-100 flex flex-col font-sans antialiased">
      {/* Top App Header */}
      <Header
        totalRows={masterRows.length}
        totalFiles={files.length}
        onExportExcel={handleExportExcel}
        onExportCSV={handleExportCSV}
        onLoadSampleData={handleLoadSampleData}
        onClearData={handleClearMaster}
        onOpenAliases={() => setIsAliasModalOpen(true)}
        onOpenHelp={() => setIsHelpOpen(true)}
        onOpenGithubDeploy={() => setIsGithubModalOpen(true)}
        onOpenOutlookMail={handleOpenBatchMail}
        activeTab={activeTab}
        setActiveTab={setActiveTab}
      />

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {/* Toast Notification */}
        {toastMessage && (
          <div className="mb-4 animate-in fade-in slide-in-from-top-2 duration-200">
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
                className="text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 ml-4"
              >
                &times;
              </button>
            </div>
          </div>
        )}

        {/* Tab 1: Master Table */}
        {activeTab === 'table' && (
          <MasterTable
            data={masterRows}
            masterHeaders={masterHeaders}
            onDeleteRow={handleDeleteRow}
            onEditRow={handleStartEditRow}
            onMailRow={handleMailRow}
            onOpenBatchMail={handleOpenBatchMail}
            onExportExcel={handleExportExcel}
            onExportCSV={handleExportCSV}
          />
        )}

        {/* Tab 2: File Uploader & Management */}
        {activeTab === 'files' && (
          <FileUploader
            files={files}
            onUploadFiles={handleUploadFiles}
            onRemoveFile={handleRemoveFile}
            onClearAllFiles={handleClearAllFiles}
            isProcessing={isProcessing}
            importMode={importMode}
            setImportMode={setImportMode}
          />
        )}

        {/* Tab 3: Mapping & Aliases */}
        {activeTab === 'mapping' && (
          <MappingViewer
            masterHeaders={masterHeaders}
            setMasterHeaders={setMasterHeaders}
            aliases={aliases}
            setAliases={setAliases}
          />
        )}

        {/* Tab 4: Logs */}
        {activeTab === 'logs' && (
          <ImportLogView
            logs={logs}
            onClearLogs={handleClearLogs}
            onExportExcel={handleExportExcel}
          />
        )}
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900 py-4 text-center text-xs text-slate-400">
        <div className="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-2">
          <span>Armature Import Web Tool &bull; Chuyển đổi từ VBA Excel Macro</span>
          <div className="flex items-center gap-3">
            <button
              onClick={() => setIsHelpOpen(true)}
              className="text-blue-500 hover:underline"
            >
              Tài liệu & Đối chiếu thuật toán VBA
            </button>
            <span className="text-slate-300 dark:text-slate-700">&bull;</span>
            <button
              onClick={() => setIsGithubModalOpen(true)}
              className="text-slate-600 dark:text-slate-300 hover:text-blue-500 flex items-center gap-1"
            >
              Hướng dẫn Deploy GitHub Pages
            </button>
          </div>
        </div>
      </footer>

      {/* Modals */}
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
