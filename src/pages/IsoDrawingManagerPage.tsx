import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Link } from 'react-router-dom';
import {
  FileSpreadsheet,
  FolderOpen,
  FolderSync,
  Play,
  Download,
  Trash2,
  AlertTriangle,
  Search,
  Upload,
  ArrowLeft,
  LayoutGrid,
  RefreshCw,
  Layers,
  Eye,
  Check,
  X,
  Sparkles,
  Clock,
  Archive,
  SlidersHorizontal,
} from 'lucide-react';
import * as XLSX from 'xlsx';
import {
  IsoDrawingRow,
  parseIsoFileName,
  matchTemplateFolder,
  formatToDDMMYYYY,
  markOldRevisionsInMatrix,
} from '../utils/isoDrawingLogic';
import {
  isFileSystemAccessSupported,
  pickDirectory,
  scanCadDirectory,
  scanCadFileList,
  extractTemplateSubfoldersFromFiles,
  listSubfolderNames,
  copyCadFilesToOutput,
  saveHandleToIDB,
  getHandleFromIDB,
  createZipFallback,
  ScannedFileMapItem,
  CopyProgress,
} from '../utils/fileSystemAccess';
import {
  generateIssuedHistoryExcel,
  getIssuedHistoryFileName,
} from '../utils/isoExcelExport';

export const IsoDrawingManagerPage: React.FC = () => {
  const [activeTab, setActiveTab] = useState<'main' | 'released' | 'oldrev'>('main');

  // Hidden File Inputs for universal directory selection
  const sourceInputRef = useRef<HTMLInputElement>(null);
  const templateInputRef = useRef<HTMLInputElement>(null);
  const relParentInputRef = useRef<HTMLInputElement>(null);

  // In-app Alert / Toast notification
  const [noticeMessage, setNoticeMessage] = useState<string | null>(null);

  // Setup State
  const [projectName, setProjectName] = useState<string>('VARD 9 80 1005');
  const [updatedBy, setUpdatedBy] = useState<string>('');
  const [logoBase64, setLogoBase64] = useState<string | null>(null);

  // Folder Handles & Names
  const [sourceHandle, setSourceHandle] = useState<FileSystemDirectoryHandle | null>(null);
  const [sourceName, setSourceName] = useState<string>('');
  const [templateHandle, setTemplateHandle] = useState<FileSystemDirectoryHandle | null>(null);
  const [templateName, setTemplateName] = useState<string>('');
  const [templateSubfolders, setTemplateSubfolders] = useState<string[]>([]);
  const [outputHandle, setOutputHandle] = useState<FileSystemDirectoryHandle | null>(null);
  const [outputName, setOutputName] = useState<string>('');

  // Scanned Raw Data & File Map
  const [fileMap, setFileMap] = useState<Map<string, ScannedFileMapItem>>(new Map());
  const [rows, setRows] = useState<IsoDrawingRow[]>([]);
  const [isScanning, setIsScanning] = useState<boolean>(false);

  // Table Filter & Search
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'valid' | 'invalid'>('all');
  const [sortField, setSortField] = useState<keyof IsoDrawingRow>('stt');
  const [sortAsc, setSortAsc] = useState<boolean>(true);

  // Execution & ZIP State
  const [isRunning, setIsRunning] = useState<boolean>(false);
  const [isZipExporting, setIsZipExporting] = useState<boolean>(false);
  const [runProgress, setRunProgress] = useState<CopyProgress | null>(null);
  const [executionLogs, setExecutionLogs] = useState<string[]>([]);
  const [isLogsModalOpen, setIsLogsModalOpen] = useState<boolean>(false);
  const [isResetConfirmOpen, setIsResetConfirmOpen] = useState<boolean>(false);

  // ISO Released State
  const [relParentHandle, setRelParentHandle] = useState<FileSystemDirectoryHandle | null>(null);
  const [relParentName, setRelParentName] = useState<string>('');
  const [relSubfolders, setRelSubfolders] = useState<{ name: string; checked: boolean }[]>([]);
  const [relIsProcessing, setRelIsProcessing] = useState<boolean>(false);
  const [relLogs, setRelLogs] = useState<string[]>([]);

  // Update Old Rev State
  const [revRawMatrix, setRevRawMatrix] = useState<unknown[][] | null>(null);
  const [revSheetNames, setRevSheetNames] = useState<string[]>([]);
  const [revActiveSheet, setRevActiveSheet] = useState<string>('');
  const [revKeyCol, setRevKeyCol] = useState<number>(11); // L
  const [revRevCol, setRevRevCol] = useState<number>(15); // P
  const [revMarkCol, setRevMarkCol] = useState<number>(18); // S
  const [revStartRow, setRevStartRow] = useState<number>(1);
  const [revPreviewCount, setRevPreviewCount] = useState<number | null>(null);
  const [revProcessedBlob, setRevProcessedBlob] = useState<Blob | null>(null);

  const logsEndRef = useRef<HTMLDivElement>(null);

  const showToast = (msg: string, durationMs: number = 3500) => {
    setNoticeMessage(msg);
    setTimeout(() => {
      setNoticeMessage((curr) => (curr === msg ? null : curr));
    }, durationMs);
  };

  useEffect(() => {
    const savedUser = localStorage.getItem('iso_updated_by');
    if (savedUser) setUpdatedBy(savedUser);
    const savedProject = localStorage.getItem('iso_project_name');
    if (savedProject) setProjectName(savedProject);
    getHandleFromIDB<string>('iso_logo_base64').then((res) => {
      if (res) setLogoBase64(res);
    });
  }, []);

  const handleUpdatedByChange = (val: string) => {
    setUpdatedBy(val);
    localStorage.setItem('iso_updated_by', val);
  };

  const handleProjectNameChange = (val: string) => {
    setProjectName(val);
    localStorage.setItem('iso_project_name', val);
  };

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const b64 = evt.target?.result as string;
      setLogoBase64(b64);
      saveHandleToIDB('iso_logo_base64', b64);
      showToast('Đã lưu Logo');
    };
    reader.readAsDataURL(file);
  };

  const handleSelectSource = async () => {
    if (!isFileSystemAccessSupported()) {
      sourceInputRef.current?.click();
      return;
    }
    try {
      const handle = await pickDirectory('iso-source');
      if (handle) {
        setSourceHandle(handle);
        setSourceName(handle.name);
        await triggerScanFromHandle(handle);
      }
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        sourceInputRef.current?.click();
      }
    }
  };

  const triggerScanFromHandle = async (handle: FileSystemDirectoryHandle) => {
    setIsScanning(true);
    try {
      const map = await scanCadDirectory(handle);
      setFileMap(map);
      populateRowsFromMap(map);
    } catch (err) {
      showToast(`Lỗi quét thư mục: ${(err as Error).message}`);
    } finally {
      setIsScanning(false);
    }
  };

  const handleSourceFilesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    const firstPath = (files[0] as unknown as { webkitRelativePath?: string }).webkitRelativePath || '';
    const folderName = firstPath ? firstPath.split('/')[0] : 'CAD_FOLDER';
    setSourceName(folderName);
    setSourceHandle(null);

    const map = scanCadFileList(files);
    setFileMap(map);
    populateRowsFromMap(map);
    showToast(`Đã nạp ${map.size} bản vẽ từ "${folderName}"`);
  };

  const populateRowsFromMap = (map: Map<string, ScannedFileMapItem>) => {
    let sttCounter = 1;
    const parsedRows: IsoDrawingRow[] = [];

    map.forEach((item) => {
      const parsed = parseIsoFileName(item.baseName);
      const targetSysFolder = matchTemplateFolder(parsed.systemShort, templateSubfolders);

      parsedRows.push({
        id: `iso_${sttCounter}_${item.baseName}`,
        stt: sttCounter++,
        fileName: item.baseName,
        pipeNumber: parsed.pipeNumber,
        zone: parsed.zone,
        system: parsed.system,
        pipeSpool: parsed.pipeSpool,
        rev: parsed.rev,
        issuedDate: parsed.issuedDate,
        remark: '',
        status: parsed.status,
        systemShort: parsed.systemShort,
        hasDWG: Boolean(item.dwgHandle || item.dwgFile),
        hasDXF: Boolean(item.dxfHandle || item.dxfFile),
        checkStatus: parsed.isValid ? 'Valid' : (parsed.errorMessage || 'Invalid'),
        isValid: parsed.isValid,
        targetFolder: targetSysFolder,
      });
    });

    setRows(parsedRows);
  };

  const handleSelectTemplate = async () => {
    if (!isFileSystemAccessSupported()) {
      templateInputRef.current?.click();
      return;
    }
    try {
      const handle = await pickDirectory('iso-template');
      if (handle) {
        setTemplateHandle(handle);
        setTemplateName(handle.name);
        const sub = await listSubfolderNames(handle);
        setTemplateSubfolders(sub);
        updateRowsWithTemplateSubfolders(sub);
        showToast(`Template: ${sub.length} subfolders`);
      }
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        templateInputRef.current?.click();
      }
    }
  };

  const handleTemplateFilesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    const firstPath = (files[0] as unknown as { webkitRelativePath?: string }).webkitRelativePath || '';
    const folderName = firstPath ? firstPath.split('/')[0] : 'TEMPLATE_FOLDER';
    setTemplateName(folderName);
    setTemplateHandle(null);

    const sub = extractTemplateSubfoldersFromFiles(files);
    setTemplateSubfolders(sub);
    updateRowsWithTemplateSubfolders(sub);
    showToast(`Template: ${sub.length} subfolders`);
  };

  const updateRowsWithTemplateSubfolders = (sub: string[]) => {
    setRows((prev) =>
      prev.map((r) => ({
        ...r,
        targetFolder: matchTemplateFolder(r.systemShort, sub),
      })),
    );
  };

  const handleSelectOutput = async () => {
    if (!isFileSystemAccessSupported()) {
      showToast('Đang chạy iFrame: Bạn có thể dùng nút "Tải Gói ZIP" tiện lợi!');
      return;
    }
    try {
      const handle = await pickDirectory('iso-output');
      if (handle) {
        setOutputHandle(handle);
        setOutputName(handle.name);
        showToast(`Đã chọn đích: "${handle.name}"`);
      }
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        showToast('Trình duyệt hạn chế chọn thư mục đích. Hãy dùng tính năng tải ZIP!');
      }
    }
  };

  const handleUpdateRow = (id: string, field: keyof IsoDrawingRow, value: string) => {
    setRows((prev) =>
      prev.map((r) => {
        if (r.id !== id) return r;
        const updated = { ...r, [field]: value };
        if (field === 'system') {
          const sysShort = matchTemplateFolder(value, templateSubfolders);
          updated.systemShort = sysShort;
          updated.targetFolder = sysShort;
        }
        if (updated.zone && updated.system && updated.pipeSpool) {
          updated.isValid = true;
          updated.checkStatus = 'Valid';
        }
        return updated;
      }),
    );
  };

  const filteredRows = useMemo(() => {
    return rows.filter((r) => {
      const q = searchQuery.toLowerCase().trim();
      const matchSearch =
        !q ||
        r.fileName.toLowerCase().includes(q) ||
        r.pipeNumber.toLowerCase().includes(q) ||
        r.zone.toLowerCase().includes(q) ||
        r.system.toLowerCase().includes(q) ||
        r.pipeSpool.toLowerCase().includes(q) ||
        r.systemShort.toLowerCase().includes(q);

      const matchStatus =
        statusFilter === 'all' ||
        (statusFilter === 'valid' && r.isValid) ||
        (statusFilter === 'invalid' && !r.isValid);

      return matchSearch && matchStatus;
    });
  }, [rows, searchQuery, statusFilter]);

  const sortedRows = useMemo(() => {
    return [...filteredRows].sort((a, b) => {
      const valA = a[sortField];
      const valB = b[sortField];
      if (valA === valB) return 0;
      if (valA === undefined || valA === null) return 1;
      if (valB === undefined || valB === null) return -1;
      const res = valA < valB ? -1 : 1;
      return sortAsc ? res : -res;
    });
  }, [filteredRows, sortField, sortAsc]);

  const stats = useMemo(() => {
    const total = rows.length;
    const valid = rows.filter((r) => r.isValid).length;
    const invalid = total - valid;
    const withDwg = rows.filter((r) => r.hasDWG).length;
    const withDxf = rows.filter((r) => r.hasDXF).length;
    return { total, valid, invalid, withDwg, withDxf };
  }, [rows]);

  const handleExecuteCopy = async (dryRun: boolean) => {
    if (rows.length === 0) {
      showToast('Chưa có dữ liệu bản vẽ.');
      return;
    }
    if (!dryRun && !outputHandle) {
      await handleDownloadZip();
      return;
    }

    setIsRunning(true);
    setRunProgress(null);
    setExecutionLogs([]);
    setIsLogsModalOpen(true);

    try {
      await copyCadFilesToOutput(
        rows,
        fileMap,
        outputHandle!,
        templateSubfolders,
        dryRun,
        (prog) => setRunProgress(prog),
        (msg) => setExecutionLogs((prev) => [...prev, msg]),
      );
      if (logsEndRef.current) {
        logsEndRef.current.scrollIntoView({ behavior: 'smooth' });
      }
    } catch (err) {
      showToast(`Lỗi: ${(err as Error).message}`);
    } finally {
      setIsRunning(false);
    }
  };

  const handleDownloadZip = async () => {
    if (rows.length === 0) {
      showToast('Chưa có dữ liệu bản vẽ.');
      return;
    }

    setIsZipExporting(true);
    setIsLogsModalOpen(true);
    setExecutionLogs((prev) => [...prev, '[ZIP] Đang nén file theo cây thư mục S{ZONE}/{System}/...']);

    try {
      const blob = await createZipFallback(rows, fileMap, (pct) => {
        setRunProgress({
          total: rows.length,
          current: Math.round((pct / 100) * rows.length),
          percentage: Math.round(pct),
          currentFileName: 'Đang nén...',
          statusText: `Đang nén: ${Math.round(pct)}%`,
        });
      });

      const zipName = `${(projectName || '1005').slice(-4)}_${sourceName || 'CAD_RELEASE'}_Tree.zip`;
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = zipName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      setExecutionLogs((prev) => [...prev, `[XONG] Tải file ZIP: ${zipName}`]);
      showToast(`Đã xuất ZIP: ${zipName}`);
    } catch (err) {
      showToast(`Lỗi ZIP: ${(err as Error).message}`);
    } finally {
      setIsZipExporting(false);
      setRunProgress(null);
    }
  };

  const handleExportExcel = async () => {
    if (rows.length === 0) {
      showToast('Chưa có dữ liệu bản vẽ.');
      return;
    }
    try {
      const blob = await generateIssuedHistoryExcel(rows, {
        projectName,
        updatedBy,
        sourceFolderName: sourceName || 'CAD_FILES',
        logoBase64: logoBase64 || undefined,
      });

      const fileName = getIssuedHistoryFileName(projectName, sourceName);
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      showToast(`Đã xuất Excel: ${fileName}`);
    } catch (err) {
      showToast(`Lỗi xuất Excel: ${(err as Error).message}`);
    }
  };

  const handleConfirmReset = () => {
    setRows([]);
    setFileMap(new Map());
    setRunProgress(null);
    setExecutionLogs([]);
    setIsResetConfirmOpen(false);
    showToast('Đã xóa dữ liệu.');
  };

  // ISO Released Handlers
  const handleSelectRelParent = async () => {
    if (!isFileSystemAccessSupported()) {
      relParentInputRef.current?.click();
      return;
    }
    try {
      const handle = await pickDirectory('iso-rel-parent');
      if (handle) {
        setRelParentHandle(handle);
        setRelParentName(handle.name);
        const sub = await listSubfolderNames(handle);
        setRelSubfolders(sub.map((name) => ({ name, checked: true })));
      }
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        relParentInputRef.current?.click();
      }
    }
  };

  const handleRelParentFilesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    const firstPath = (files[0] as unknown as { webkitRelativePath?: string }).webkitRelativePath || '';
    const folderName = firstPath ? firstPath.split('/')[0] : 'PARENT_FOLDER';
    setRelParentName(folderName);

    const sub = extractTemplateSubfoldersFromFiles(files);
    setRelSubfolders(sub.map((name) => ({ name, checked: true })));
  };

  const handleExecuteIsoReleased = async () => {
    const selectedFolders = relSubfolders.filter((f) => f.checked);
    if (selectedFolders.length === 0) {
      showToast('Vui lòng chọn ít nhất 1 thư mục con!');
      return;
    }

    const d = new Date();
    const yyyymmdd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
    const targetFolderName = `Iso released ${yyyymmdd}`;

    setRelIsProcessing(true);
    setRelLogs([`Tạo thư mục ${targetFolderName}...`]);

    try {
      if (outputHandle) {
        const targetDirHandle = await outputHandle.getDirectoryHandle(targetFolderName, { create: true });
        let totalCopied = 0;
        let totalOverwritten = 0;

        for (const item of selectedFolders) {
          const subHandle = await relParentHandle!.getDirectoryHandle(item.name);

          for await (const [fName, fHandle] of (subHandle as unknown as AsyncIterable<[string, FileSystemHandle]>)) {
            if (fHandle.kind !== 'file') continue;
            let exists = false;
            try {
              await targetDirHandle.getFileHandle(fName, { create: false });
              exists = true;
            } catch {
              exists = false;
            }

            if (exists) totalOverwritten++;
            else totalCopied++;

            const srcFile = await (fHandle as FileSystemFileHandle).getFile();
            const destHandle = await targetDirHandle.getFileHandle(fName, { create: true });
            const writable = await (destHandle as unknown as { createWritable: () => Promise<FileSystemWritableFileStream> }).createWritable();
            await writable.write(srcFile);
            await writable.close();
          }
        }
        setRelLogs((prev) => [...prev, `[XONG] File mới: ${totalCopied} | Ghi đè: ${totalOverwritten}`]);
      } else {
        setRelLogs((prev) => [...prev, 'Cần chọn thư mục đầu ra ở Bước 1.']);
      }
    } catch (err) {
      setRelLogs((prev) => [...prev, `Lỗi: ${(err as Error).message}`]);
    } finally {
      setRelIsProcessing(false);
    }
  };

  // Update Old Rev Handlers
  const handleUploadRevExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        setRevSheetNames(workbook.SheetNames);
        const firstSheet = workbook.SheetNames[0];
        setRevActiveSheet(firstSheet);
        const sheet = workbook.Sheets[firstSheet];
        const matrix = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1 });
        setRevRawMatrix(matrix);
        setRevPreviewCount(null);
        setRevProcessedBlob(null);
        showToast(`Đã nạp file (${matrix.length} dòng)`);
      } catch (err) {
        showToast(`Lỗi: ${(err as Error).message}`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleExecuteMarkOldRev = () => {
    if (!revRawMatrix || revRawMatrix.length === 0) return;
    const { updatedMatrix, markedCount } = markOldRevisionsInMatrix(
      revRawMatrix,
      revStartRow,
      revKeyCol,
      revRevCol,
      revMarkCol,
    );
    setRevPreviewCount(markedCount);
    const newWb = XLSX.utils.book_new();
    const newWs = XLSX.utils.aoa_to_sheet(updatedMatrix);
    XLSX.utils.book_append_sheet(newWb, newWs, revActiveSheet || 'Sheet1');
    const outBuf = XLSX.write(newWb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([outBuf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    setRevProcessedBlob(blob);
    showToast(`Đã đánh dấu "old rev" ${markedCount} dòng`);
  };

  const handleDownloadMarkedExcel = () => {
    if (!revProcessedBlob) return;
    const url = URL.createObjectURL(revProcessedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Updated_OldRev_${formatToDDMMYYYY().replace(/\//g, '')}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen w-full bg-slate-950 text-slate-100 flex flex-col font-sans selection:bg-purple-600 selection:text-white">
      {/* Hidden File Inputs */}
      <input ref={sourceInputRef} type="file" webkitdirectory="" multiple onChange={handleSourceFilesChange} className="hidden" />
      <input ref={templateInputRef} type="file" webkitdirectory="" multiple onChange={handleTemplateFilesChange} className="hidden" />
      <input ref={relParentInputRef} type="file" webkitdirectory="" multiple onChange={handleRelParentFilesChange} className="hidden" />

      {/* Floating Notification */}
      {noticeMessage && (
        <div className="fixed top-4 right-6 z-50 px-4 py-2.5 rounded-xl bg-purple-900 border border-purple-400 text-white shadow-2xl text-xs sm:text-sm font-semibold flex items-center gap-2.5 animate-in fade-in slide-in-from-top-2">
          <Sparkles className="w-4 h-4 text-purple-300 shrink-0" />
          <span>{noticeMessage}</span>
          <button onClick={() => setNoticeMessage(null)} className="text-purple-300 hover:text-white ml-2 font-bold cursor-pointer">
            &times;
          </button>
        </div>
      )}

      {/* FULL-WIDTH HEADER */}
      <header className="sticky top-0 z-40 bg-slate-900/95 backdrop-blur-md border-b border-slate-800 w-full">
        <div className="w-full px-4 sm:px-6 lg:px-8 xl:px-10 h-16 flex items-center justify-between gap-4">
          <div className="flex items-center space-x-3">
            <Link
              to="/"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 text-xs sm:text-sm font-bold transition-all shadow-xs"
              title="Quay lại Hub"
            >
              <ArrowLeft className="w-4 h-4 text-slate-400" />
              <LayoutGrid className="w-4 h-4 text-cyan-400" />
              <span>Tool Hub</span>
            </Link>

            <div className="w-9 h-9 rounded-xl bg-gradient-to-tr from-purple-600 to-indigo-600 flex items-center justify-center text-white shadow-md shadow-purple-600/30">
              <FolderSync className="w-5 h-5 text-white" />
            </div>

            <div className="flex items-center gap-2.5">
              <h1 className="text-base sm:text-lg font-black tracking-tight text-white uppercase">ISO DRAWING MANAGER</h1>
              <span className="text-xs font-bold px-2.5 py-0.5 rounded-full bg-purple-500/20 text-purple-300 border border-purple-500/30">
                {stats.total} CAD Files
              </span>
            </div>
          </div>

          {/* Header Actions */}
          <div className="flex items-center gap-2.5">
            <button
              type="button"
              onClick={() => setIsLogsModalOpen(true)}
              className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-xl text-xs sm:text-sm font-bold bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 cursor-pointer transition-colors"
            >
              <Clock className="w-4 h-4 text-purple-400" />
              <span>Nhật ký</span>
              {executionLogs.length > 0 && (
                <span className="px-2 py-0.2 text-xs rounded-full bg-purple-600 text-white font-bold">
                  {executionLogs.length}
                </span>
              )}
            </button>

            <button
              type="button"
              onClick={() => setIsResetConfirmOpen(true)}
              className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-xl text-xs sm:text-sm font-bold bg-rose-950/40 hover:bg-rose-900/60 text-rose-300 border border-rose-800/60 cursor-pointer transition-colors"
            >
              <Trash2 className="w-4 h-4 text-rose-400" />
              <span>Reset</span>
            </button>
          </div>
        </div>
      </header>

      {/* FULL-WIDTH TABS */}
      <div className="bg-slate-900 border-b border-slate-800 w-full px-4 sm:px-6 lg:px-8 xl:px-10">
        <div className="w-full flex items-center gap-2 py-2">
          <button
            onClick={() => setActiveTab('main')}
            className={`px-4 py-2 rounded-xl text-xs sm:text-sm font-bold transition-all cursor-pointer ${
              activeTab === 'main'
                ? 'bg-purple-600 text-white shadow-md shadow-purple-600/30'
                : 'text-slate-400 hover:text-white'
            }`}
          >
            Quản Lý &amp; Bóc Tách Bản Vẽ
          </button>
          <button
            onClick={() => setActiveTab('released')}
            className={`px-4 py-2 rounded-xl text-xs sm:text-sm font-bold transition-all cursor-pointer ${
              activeTab === 'released'
                ? 'bg-purple-600 text-white shadow-md shadow-purple-600/30'
                : 'text-slate-400 hover:text-white'
            }`}
          >
            ISO Released (Gộp Folder)
          </button>
          <button
            onClick={() => setActiveTab('oldrev')}
            className={`px-4 py-2 rounded-xl text-xs sm:text-sm font-bold transition-all cursor-pointer ${
              activeTab === 'oldrev'
                ? 'bg-purple-600 text-white shadow-md shadow-purple-600/30'
                : 'text-slate-400 hover:text-white'
            }`}
          >
            Update Old Rev (Excel)
          </button>
        </div>
      </div>

      {/* MAIN CONTAINER: FULL WIDTH & FULL HEIGHT */}
      <main className="w-full px-4 sm:px-6 lg:px-8 xl:px-10 py-4 flex-1 flex flex-col space-y-3.5">
        {activeTab === 'main' && (
          <div className="w-full flex-1 flex flex-col space-y-3.5">
            {/* EXPANDED CONFIG & FOLDERS TOOLBAR */}
            <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-4 shadow-sm flex flex-wrap items-center justify-between gap-4 text-xs sm:text-sm">
              {/* Inputs & Logo */}
              <div className="flex flex-wrap items-center gap-3.5">
                <div className="flex items-center gap-2">
                  <span className="text-slate-400 font-bold">Dự án:</span>
                  <input
                    type="text"
                    value={projectName}
                    onChange={(e) => handleProjectNameChange(e.target.value)}
                    placeholder="Tên dự án..."
                    className="w-48 sm:w-56 px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-xs sm:text-sm text-white focus:outline-none focus:ring-2 focus:ring-purple-500 font-bold"
                  />
                </div>

                <div className="flex items-center gap-2">
                  <span className="text-slate-400 font-bold">Cập nhật:</span>
                  <input
                    type="text"
                    value={updatedBy}
                    onChange={(e) => handleUpdatedByChange(e.target.value)}
                    placeholder="Người cập nhật..."
                    className="w-40 sm:w-48 px-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-xs sm:text-sm text-white focus:outline-none focus:ring-2 focus:ring-purple-500"
                  />
                </div>

                {/* Logo upload */}
                <label className="cursor-pointer inline-flex items-center gap-2 px-3 py-1.5 bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 rounded-xl text-xs sm:text-sm font-bold transition-colors">
                  <Upload className="w-4 h-4 text-purple-400" />
                  <span>{logoBase64 ? 'Đổi Logo' : 'Logo B3'}</span>
                  <input type="file" accept="image/png,image/jpeg" onChange={handleLogoUpload} className="hidden" />
                </label>
                {logoBase64 && (
                  <img src={logoBase64} alt="Logo" className="h-8 max-w-24 object-contain rounded-lg bg-slate-800 p-1 border border-slate-700" />
                )}
              </div>

              {/* 3 Folder Selector Buttons */}
              <div className="flex flex-wrap items-center gap-2.5">
                <button
                  type="button"
                  onClick={handleSelectSource}
                  className="inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs sm:text-sm font-bold bg-purple-600 hover:bg-purple-500 text-white cursor-pointer transition-all shadow-md shadow-purple-600/30 active:scale-95"
                  title="Chọn thư mục chứa file .dwg, .dxf"
                >
                  <FolderOpen className="w-4 h-4" />
                  <span>{sourceName || '1. Chọn Thư Mục CAD'}</span>
                  {rows.length > 0 && <span className="bg-purple-900/90 px-2 py-0.5 rounded-full text-xs font-extrabold">{rows.length}</span>}
                </button>

                <button
                  type="button"
                  onClick={handleSelectTemplate}
                  className="inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs sm:text-sm font-bold bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 cursor-pointer transition-colors"
                  title="Chọn thư mục template mẫu"
                >
                  <Layers className="w-4 h-4 text-amber-400" />
                  <span>{templateName ? `Mẫu: ${templateName}` : '2. Thư Mục Template'}</span>
                  {templateSubfolders.length > 0 && (
                    <span className="text-amber-400 font-extrabold">({templateSubfolders.length})</span>
                  )}
                </button>

                <button
                  type="button"
                  onClick={handleSelectOutput}
                  className="inline-flex items-center gap-2 px-3.5 py-2 rounded-xl text-xs sm:text-sm font-bold bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 cursor-pointer transition-colors"
                  title="Chọn thư mục đầu ra"
                >
                  <FolderSync className="w-4 h-4 text-cyan-400" />
                  <span>{outputName ? `Đích: ${outputName}` : '3. Thư Mục Đích'}</span>
                </button>
              </div>
            </div>

            {/* ACTION TOOLBAR: Clean, enlarged buttons */}
            <div className="bg-slate-900/90 border border-slate-800 rounded-2xl p-3 shadow-sm flex flex-wrap items-center justify-between gap-3 text-xs sm:text-sm">
              {/* Left actions: Run dry, copy, zip, excel */}
              <div className="flex flex-wrap items-center gap-2.5">
                <button
                  type="button"
                  onClick={() => handleExecuteCopy(true)}
                  disabled={isRunning || rows.length === 0}
                  className="inline-flex items-center gap-2 px-4 py-2 rounded-xl font-bold bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 cursor-pointer disabled:opacity-40 transition-colors"
                >
                  <Eye className="w-4 h-4 text-cyan-400" />
                  <span>Chạy Thử</span>
                </button>

                <button
                  type="button"
                  onClick={() => handleExecuteCopy(false)}
                  disabled={isRunning || rows.length === 0}
                  className="inline-flex items-center gap-2 px-5 py-2 rounded-xl font-bold bg-gradient-to-r from-purple-600 via-indigo-600 to-cyan-500 hover:from-purple-500 hover:to-cyan-400 text-white shadow-md shadow-purple-600/30 cursor-pointer disabled:opacity-40 transition-all active:scale-95"
                >
                  <Play className="w-4 h-4 fill-white" />
                  <span>{isRunning ? 'Đang Chạy...' : 'Copy Vào Thư Mục Đích'}</span>
                </button>

                <button
                  type="button"
                  onClick={handleDownloadZip}
                  disabled={isZipExporting || rows.length === 0}
                  className="inline-flex items-center gap-2 px-4 py-2 rounded-xl font-bold bg-indigo-950/60 hover:bg-indigo-900/80 text-indigo-300 border border-indigo-700/60 cursor-pointer disabled:opacity-40 transition-colors"
                >
                  <Archive className="w-4 h-4 text-indigo-400" />
                  <span>{isZipExporting ? 'Đang Nén...' : 'Tải Gói ZIP'}</span>
                </button>

                <button
                  type="button"
                  onClick={handleExportExcel}
                  disabled={rows.length === 0}
                  className="inline-flex items-center gap-2 px-4 py-2 rounded-xl font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-md shadow-emerald-600/30 cursor-pointer disabled:opacity-40 transition-all active:scale-95"
                >
                  <Download className="w-4 h-4" />
                  <span>Xuất Excel (.xlsx)</span>
                </button>
              </div>

              {/* Right: Quick filter & search */}
              <div className="flex items-center gap-3">
                <div className="relative w-64 sm:w-80">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Tìm theo Zone, System, Spool, File..."
                    className="w-full pl-9 pr-3 py-1.5 bg-slate-950 border border-slate-700 rounded-xl text-xs sm:text-sm text-white focus:outline-none focus:ring-2 focus:ring-purple-500"
                  />
                </div>

                <select
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value as 'all' | 'valid' | 'invalid')}
                  className="bg-slate-950 border border-slate-700 rounded-xl px-3 py-1.5 text-xs sm:text-sm text-slate-300 font-semibold"
                >
                  <option value="all">Tất cả ({rows.length})</option>
                  <option value="valid">Hợp lệ ({stats.valid})</option>
                  <option value="invalid">Lỗi định dạng ({stats.invalid})</option>
                </select>
              </div>
            </div>

            {/* EXPANDED FULL-HEIGHT DATA TABLE */}
            <div className="bg-slate-900/90 border border-slate-800 rounded-2xl overflow-hidden shadow-lg flex-1 flex flex-col min-h-[500px]">
              <div className="overflow-auto flex-1 w-full">
                <table className="w-full text-left border-collapse">
                  <thead className="bg-slate-900 sticky top-0 z-10 border-b border-slate-800 text-xs font-bold text-slate-300 uppercase tracking-wider">
                    <tr>
                      <th className="py-3 px-3 text-center w-12">STT</th>
                      <th className="py-3 px-4 min-w-[240px]">FILE NAME</th>
                      <th className="py-3 px-4 min-w-[200px]">PIPE NUMBER</th>
                      <th className="py-3 px-3 text-center min-w-[80px]">ZONE</th>
                      <th className="py-3 px-3 text-center min-w-[110px]">SYSTEM</th>
                      <th className="py-3 px-3 text-center min-w-[120px]">PIPE SPOOL</th>
                      <th className="py-3 px-3 text-center min-w-[60px]">REV</th>
                      <th className="py-3 px-3 text-center min-w-[110px]">ISSUED DATE</th>
                      <th className="py-3 px-4 min-w-[150px]">REMARK</th>
                      <th className="py-3 px-3 text-center min-w-[90px]">STATUS</th>
                      <th className="py-3 px-3 text-center min-w-[100px]">SHORT</th>
                      <th className="py-3 px-3 text-center min-w-[100px]">CAD</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60 font-mono text-xs sm:text-sm">
                    {sortedRows.length === 0 ? (
                      <tr>
                        <td colSpan={12} className="py-24 text-center font-sans">
                          <div className="flex flex-col items-center justify-center space-y-4">
                            <div className="w-16 h-16 rounded-2xl bg-purple-600/20 text-purple-400 flex items-center justify-center border border-purple-500/30">
                              <FolderOpen className="w-8 h-8" />
                            </div>
                            <h3 className="text-base sm:text-lg font-bold text-slate-200">
                              {rows.length === 0 ? 'Chưa có dữ liệu bản vẽ CAD' : 'Không tìm thấy kết quả phù hợp'}
                            </h3>
                            <p className="text-xs sm:text-sm text-slate-400 max-w-md">
                              {rows.length === 0
                                ? 'Bấm nút bên dưới để chọn thư mục chứa các file .dwg & .dxf. Hệ thống sẽ tự động bóc tách và phân nhóm Zone/System.'
                                : 'Hãy thử thay đổi từ khóa tìm kiếm hoặc đặt lại bộ lọc trạng thái.'}
                            </p>
                            {rows.length === 0 && (
                              <button
                                type="button"
                                onClick={handleSelectSource}
                                className="px-5 py-2.5 rounded-xl text-xs sm:text-sm font-bold bg-purple-600 hover:bg-purple-500 text-white shadow-lg shadow-purple-600/30 cursor-pointer transition-all active:scale-95"
                              >
                                Chọn Thư Mục CAD Nguồn
                              </button>
                            )}
                          </div>
                        </td>
                      </tr>
                    ) : (
                      sortedRows.map((r) => {
                        const isInvalid = !r.isValid;
                        return (
                          <tr key={r.id} className={`hover:bg-slate-800/50 transition-colors ${isInvalid ? 'bg-rose-950/25 text-rose-200' : ''}`}>
                            <td className="py-2.5 px-3 text-center text-slate-500 font-semibold">{r.stt}</td>
                            <td className="py-2.5 px-4 font-semibold text-slate-100 truncate max-w-[280px]" title={r.fileName}>
                              {r.fileName}
                            </td>
                            <td className="py-2.5 px-4">
                              <input
                                type="text"
                                value={r.pipeNumber}
                                onChange={(e) => handleUpdateRow(r.id, 'pipeNumber', e.target.value)}
                                className="w-full bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none text-slate-300 py-0.5"
                              />
                            </td>
                            <td className="py-2.5 px-3 text-center font-bold text-amber-300">
                              <input
                                type="text"
                                value={r.zone}
                                onChange={(e) => handleUpdateRow(r.id, 'zone', e.target.value)}
                                className="w-16 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none py-0.5"
                              />
                            </td>
                            <td className="py-2.5 px-3 text-center text-cyan-300 font-semibold">
                              <input
                                type="text"
                                value={r.system}
                                onChange={(e) => handleUpdateRow(r.id, 'system', e.target.value)}
                                className="w-24 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none py-0.5"
                              />
                            </td>
                            <td className="py-2.5 px-3 text-center text-emerald-300 font-semibold">
                              <input
                                type="text"
                                value={r.pipeSpool}
                                onChange={(e) => handleUpdateRow(r.id, 'pipeSpool', e.target.value)}
                                className="w-28 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none py-0.5"
                              />
                            </td>
                            <td className="py-2.5 px-3 text-center font-black text-white">
                              <input
                                type="text"
                                value={r.rev}
                                onChange={(e) => handleUpdateRow(r.id, 'rev', e.target.value)}
                                className="w-10 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none py-0.5"
                              />
                            </td>
                            <td className="py-2.5 px-3 text-center text-slate-400">
                              <input
                                type="text"
                                value={r.issuedDate}
                                onChange={(e) => handleUpdateRow(r.id, 'issuedDate', e.target.value)}
                                className="w-24 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none py-0.5"
                              />
                            </td>
                            <td className="py-2.5 px-4">
                              <input
                                type="text"
                                value={r.remark}
                                placeholder="Ghi chú..."
                                onChange={(e) => handleUpdateRow(r.id, 'remark', e.target.value)}
                                className="w-full bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none text-slate-400 text-xs py-0.5"
                              />
                            </td>
                            <td className="py-2.5 px-3 text-center">
                              <span className="px-2 py-0.5 rounded-md text-xs font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                                {r.status}
                              </span>
                            </td>
                            <td className="py-2.5 px-3 text-center font-bold text-purple-300">
                              {r.systemShort || '-'}
                            </td>
                            <td className="py-2.5 px-3 text-center">
                              <div className="flex items-center justify-center gap-1.5">
                                {r.hasDWG && <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-blue-500/20 text-blue-300 border border-blue-500/30">DWG</span>}
                                {r.hasDXF && <span className="px-1.5 py-0.5 rounded text-[10px] font-bold bg-purple-500/20 text-purple-300 border border-purple-500/30">DXF</span>}
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: ISO RELEASED TOOL */}
        {activeTab === 'released' && (
          <div className="w-full max-w-4xl mx-auto bg-slate-900/90 border border-slate-800 rounded-2xl p-6 shadow-md space-y-5 text-xs sm:text-sm">
            <div className="flex items-center justify-between border-b border-slate-800 pb-4">
              <h2 className="text-base font-bold text-white flex items-center gap-2.5">
                <Archive className="w-5 h-5 text-purple-400" />
                <span>ISO Released (Gộp Bản Vẽ)</span>
              </h2>
              <button
                type="button"
                onClick={handleSelectRelParent}
                className="px-4 py-2 rounded-xl font-bold bg-purple-600 hover:bg-purple-500 text-white cursor-pointer transition-colors shadow-xs"
              >
                {relParentName ? `Thư mục: ${relParentName}` : 'Chọn Thư Mục Mẹ'}
              </button>
            </div>

            {relSubfolders.length > 0 && (
              <div className="space-y-4">
                <div className="flex items-center justify-between text-xs text-slate-400">
                  <span className="font-semibold">Chọn các thư mục con cần copy ({relSubfolders.filter((f) => f.checked).length}/{relSubfolders.length})</span>
                  <div className="flex gap-2.5">
                    <button onClick={() => setRelSubfolders((prev) => prev.map((item) => ({ ...item, checked: true })))} className="text-purple-400 hover:underline cursor-pointer font-semibold">
                      Chọn tất cả
                    </button>
                    <span>|</span>
                    <button onClick={() => setRelSubfolders((prev) => prev.map((item) => ({ ...item, checked: false })))} className="text-slate-400 hover:underline cursor-pointer font-semibold">
                      Bỏ chọn
                    </button>
                  </div>
                </div>

                <div className="p-3 bg-slate-950 border border-slate-800 rounded-xl max-h-72 overflow-y-auto grid grid-cols-1 sm:grid-cols-2 gap-2 font-mono text-xs">
                  {relSubfolders.map((f, idx) => (
                    <label key={idx} className="flex items-center gap-2.5 p-2 rounded-lg hover:bg-slate-900 cursor-pointer border border-transparent hover:border-slate-800">
                      <input
                        type="checkbox"
                        checked={f.checked}
                        onChange={(e) => {
                          const checked = e.target.checked;
                          setRelSubfolders((prev) => prev.map((item, i) => (i === idx ? { ...item, checked } : item)));
                        }}
                        className="rounded border-slate-700 text-purple-600 w-4 h-4"
                      />
                      <span className="truncate text-slate-200">{f.name}</span>
                    </label>
                  ))}
                </div>

                <button
                  type="button"
                  onClick={handleExecuteIsoReleased}
                  disabled={relIsProcessing}
                  className="w-full py-3 rounded-xl font-bold bg-purple-600 hover:bg-purple-500 text-white cursor-pointer disabled:opacity-40 transition-all shadow-md shadow-purple-600/30"
                >
                  {relIsProcessing ? 'Đang Sao Chép...' : 'Tạo Thư Mục "Iso released yyyymmdd" & Copy Files'}
                </button>
              </div>
            )}

            {relLogs.length > 0 && (
              <div className="p-3 bg-slate-950 border border-slate-800 rounded-xl max-h-48 overflow-y-auto font-mono text-xs space-y-1 text-slate-400">
                {relLogs.map((m, i) => (
                  <div key={i}>{m}</div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* TAB 3: UPDATE OLD REV TOOL */}
        {activeTab === 'oldrev' && (
          <div className="w-full max-w-4xl mx-auto bg-slate-900/90 border border-slate-800 rounded-2xl p-6 shadow-md space-y-5 text-xs sm:text-sm">
            <div className="flex items-center justify-between border-b border-slate-800 pb-4">
              <h2 className="text-base font-bold text-white flex items-center gap-2.5">
                <FileSpreadsheet className="w-5 h-5 text-cyan-400" />
                <span>Update Old Rev</span>
              </h2>
              <label className="cursor-pointer px-4 py-2 rounded-xl font-bold bg-purple-600 hover:bg-purple-500 text-white transition-colors shadow-xs">
                <span>Upload File Excel (.xlsx)</span>
                <input type="file" accept=".xlsx, .xls" onChange={handleUploadRevExcel} className="hidden" />
              </label>
            </div>

            {revRawMatrix && (
              <div className="space-y-4">
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                  <div>
                    <label className="text-slate-400 block mb-1 font-semibold">Sheet:</label>
                    <select
                      value={revActiveSheet}
                      onChange={(e) => setRevActiveSheet(e.target.value)}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl p-2 text-white font-semibold"
                    >
                      {revSheetNames.map((s) => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="text-slate-400 block mb-1 font-semibold">Cột Key (L):</label>
                    <input
                      type="number"
                      value={revKeyCol + 1}
                      onChange={(e) => setRevKeyCol(Math.max(0, Number(e.target.value) - 1))}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl p-2 text-white font-semibold"
                    />
                  </div>

                  <div>
                    <label className="text-slate-400 block mb-1 font-semibold">Cột Rev (P):</label>
                    <input
                      type="number"
                      value={revRevCol + 1}
                      onChange={(e) => setRevRevCol(Math.max(0, Number(e.target.value) - 1))}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl p-2 text-white font-semibold"
                    />
                  </div>

                  <div>
                    <label className="text-slate-400 block mb-1 font-semibold">Cột Đánh Dấu (S):</label>
                    <input
                      type="number"
                      value={revMarkCol + 1}
                      onChange={(e) => setRevMarkCol(Math.max(0, Number(e.target.value) - 1))}
                      className="w-full bg-slate-950 border border-slate-700 rounded-xl p-2 text-white font-semibold"
                    />
                  </div>
                </div>

                <div className="flex gap-3 pt-2">
                  <button
                    type="button"
                    onClick={handleExecuteMarkOldRev}
                    className="flex-1 py-2.5 rounded-xl font-bold bg-cyan-600 hover:bg-cyan-500 text-white cursor-pointer transition-colors shadow-xs"
                  >
                    Đánh Dấu "old rev"
                  </button>

                  {revPreviewCount !== null && (
                    <button
                      type="button"
                      onClick={handleDownloadMarkedExcel}
                      className="flex-1 py-2.5 rounded-xl font-bold bg-emerald-600 hover:bg-emerald-500 text-white cursor-pointer flex items-center justify-center gap-2 transition-colors shadow-xs"
                    >
                      <Download className="w-4 h-4" />
                      <span>Tải Excel ({revPreviewCount} dòng cũ)</span>
                    </button>
                  )}
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      {/* POPUP MODAL: NHẬT KÝ THỰC THI (EXECUTION LOGS) */}
      {isLogsModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/75 backdrop-blur-xs animate-in fade-in duration-150">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-2xl w-full p-5 space-y-4 shadow-2xl flex flex-col max-h-[85vh]">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <div className="flex items-center gap-2.5">
                <Clock className="w-5 h-5 text-purple-400" />
                <h3 className="text-sm sm:text-base font-bold text-white uppercase tracking-wider">Nhật Ký Thực Thi</h3>
                <span className="text-xs font-mono px-2.5 py-0.5 rounded-full bg-slate-800 text-slate-300 font-bold">
                  {executionLogs.length} mục
                </span>
              </div>
              <button
                onClick={() => setIsLogsModalOpen(false)}
                className="text-slate-400 hover:text-white p-1 rounded-xl hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Progress bar inside popup */}
            {runProgress && (
              <div className="space-y-1.5 text-xs">
                <div className="flex justify-between font-semibold text-purple-300">
                  <span>{runProgress.statusText}</span>
                  <span className="font-bold">{runProgress.percentage}%</span>
                </div>
                <div className="w-full h-2 bg-slate-800 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-gradient-to-r from-purple-500 to-cyan-400 transition-all duration-150"
                    style={{ width: `${runProgress.percentage}%` }}
                  />
                </div>
              </div>
            )}

            {/* Log Stream */}
            <div className="flex-1 min-h-[250px] max-h-[420px] p-3.5 bg-slate-950 border border-slate-800 rounded-xl overflow-y-auto font-mono text-xs space-y-1 text-slate-300">
              {executionLogs.length === 0 ? (
                <p className="text-slate-600 italic">Chưa có nhật ký ghi nhận.</p>
              ) : (
                executionLogs.map((logMsg, idx) => (
                  <div
                    key={idx}
                    className={`${
                      logMsg.includes('[LỖI]')
                        ? 'text-rose-400 font-bold'
                        : logMsg.includes('[HOÀN THÀNH]') || logMsg.includes('[XONG]')
                        ? 'text-emerald-400 font-bold'
                        : logMsg.includes('[THIẾU]')
                        ? 'text-amber-400'
                        : logMsg.includes('[BỎ QUA')
                        ? 'text-slate-500'
                        : 'text-cyan-300'
                    }`}
                  >
                    {logMsg}
                  </div>
                ))
              )}
              <div ref={logsEndRef} />
            </div>

            <div className="flex items-center justify-between pt-1 text-xs">
              <button
                onClick={() => setExecutionLogs([])}
                className="text-slate-400 hover:text-white text-xs underline cursor-pointer"
              >
                Xóa nhật ký
              </button>
              <button
                onClick={() => setIsLogsModalOpen(false)}
                className="px-5 py-2 rounded-xl font-bold bg-slate-800 hover:bg-slate-750 text-white cursor-pointer transition-colors"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {/* CONFIRM RESET MODAL */}
      {isResetConfirmOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/75 backdrop-blur-xs animate-in fade-in duration-150">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-sm w-full p-5 space-y-3.5 shadow-2xl">
            <h3 className="text-base font-bold text-white">Xóa trắng bảng dữ liệu?</h3>
            <p className="text-xs text-slate-400 leading-relaxed">
              Các bản vẽ trên bảng sẽ bị xóa. Cấu hình thư mục và tên dự án được giữ nguyên.
            </p>
            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                onClick={() => setIsResetConfirmOpen(false)}
                className="px-3.5 py-1.5 rounded-xl text-xs font-semibold text-slate-400 hover:bg-slate-800 cursor-pointer"
              >
                Hủy
              </button>
              <button
                onClick={handleConfirmReset}
                className="px-4 py-1.5 rounded-xl text-xs font-bold bg-rose-600 hover:bg-rose-500 text-white cursor-pointer transition-colors"
              >
                Xóa sạch
              </button>
            </div>
          </div>
        </div>
      )}

      {/* FULL-WIDTH COMPACT FOOTER */}
      <footer className="border-t border-slate-800/80 bg-slate-950 py-3 text-xs text-slate-500 text-center w-full">
        <div className="w-full px-4 sm:px-6 lg:px-8 xl:px-10 flex items-center justify-between text-xs">
          <span>WebToolRush &bull; ISO Drawing Manager</span>
          <span className="font-mono text-slate-400">100% Client-Side In-Memory Execution</span>
        </div>
      </footer>
    </div>
  );
};
