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
      showToast('Đang chạy iFrame: Bạn có thể dùng nút "Tải ZIP" tiện lợi!');
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
    <div className="min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans selection:bg-purple-600 selection:text-white">
      {/* Hidden File Inputs */}
      <input ref={sourceInputRef} type="file" webkitdirectory="" multiple onChange={handleSourceFilesChange} className="hidden" />
      <input ref={templateInputRef} type="file" webkitdirectory="" multiple onChange={handleTemplateFilesChange} className="hidden" />
      <input ref={relParentInputRef} type="file" webkitdirectory="" multiple onChange={handleRelParentFilesChange} className="hidden" />

      {/* Floating Notification */}
      {noticeMessage && (
        <div className="fixed top-3 right-4 z-50 px-3.5 py-2 rounded-xl bg-purple-900 border border-purple-400 text-white shadow-xl text-xs font-semibold flex items-center gap-2 animate-in fade-in slide-in-from-top-2">
          <Sparkles className="w-3.5 h-3.5 text-purple-300 shrink-0" />
          <span>{noticeMessage}</span>
          <button onClick={() => setNoticeMessage(null)} className="text-purple-300 hover:text-white ml-1 font-bold cursor-pointer">
            &times;
          </button>
        </div>
      )}

      {/* COMPACT HEADER */}
      <header className="sticky top-0 z-40 bg-slate-900/95 backdrop-blur-md border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-14 flex items-center justify-between gap-3">
          <div className="flex items-center space-x-2.5">
            <Link
              to="/"
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white border border-slate-700 text-xs font-bold transition-all shadow-xs"
              title="Quay lại Hub"
            >
              <ArrowLeft className="w-3.5 h-3.5 text-slate-400" />
              <LayoutGrid className="w-3.5 h-3.5 text-cyan-400" />
              <span className="hidden sm:inline">Hub</span>
            </Link>

            <div className="w-7 h-7 rounded-lg bg-gradient-to-tr from-purple-600 to-indigo-600 flex items-center justify-center text-white shadow-xs">
              <FolderSync className="w-4 h-4 text-white" />
            </div>

            <div className="flex items-center gap-2">
              <h1 className="text-sm font-black tracking-tight text-white uppercase">ISO DRAWING MANAGER</h1>
              <span className="text-[10px] font-bold px-2 py-0.2 rounded-full bg-purple-500/20 text-purple-300 border border-purple-500/30">
                {stats.total} CAD
              </span>
            </div>
          </div>

          {/* Header Actions */}
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setIsLogsModalOpen(true)}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-750 text-slate-300 border border-slate-700 cursor-pointer"
            >
              <Clock className="w-3.5 h-3.5 text-purple-400" />
              <span>Nhật ký</span>
              {executionLogs.length > 0 && (
                <span className="px-1.5 py-0.2 text-[10px] rounded-full bg-purple-600 text-white font-bold">
                  {executionLogs.length}
                </span>
              )}
            </button>

            <button
              type="button"
              onClick={() => setIsResetConfirmOpen(true)}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-rose-950/40 hover:bg-rose-900/60 text-rose-300 border border-rose-800/60 cursor-pointer"
            >
              <Trash2 className="w-3.5 h-3.5 text-rose-400" />
              <span>Reset</span>
            </button>
          </div>
        </div>
      </header>

      {/* COMPACT TABS */}
      <div className="bg-slate-900 border-b border-slate-800 px-4 sm:px-6 lg:px-8">
        <div className="max-w-7xl mx-auto flex items-center gap-2 py-1.5">
          <button
            onClick={() => setActiveTab('main')}
            className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'main'
                ? 'bg-purple-600 text-white shadow-xs'
                : 'text-slate-400 hover:text-white'
            }`}
          >
            Quản Lý &amp; Bóc Tách Bản Vẽ
          </button>
          <button
            onClick={() => setActiveTab('released')}
            className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'released'
                ? 'bg-purple-600 text-white shadow-xs'
                : 'text-slate-400 hover:text-white'
            }`}
          >
            ISO Released (Gộp Folder)
          </button>
          <button
            onClick={() => setActiveTab('oldrev')}
            className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'oldrev'
                ? 'bg-purple-600 text-white shadow-xs'
                : 'text-slate-400 hover:text-white'
            }`}
          >
            Update Old Rev (Excel)
          </button>
        </div>
      </div>

      {/* MAIN CONTAINER */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex-1 w-full space-y-4">
        {activeTab === 'main' && (
          <>
            {/* COMPACT CONFIG & FOLDERS TOOLBAR */}
            <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-3 shadow-xs flex flex-wrap items-center justify-between gap-3 text-xs">
              {/* Inputs & Logo */}
              <div className="flex flex-wrap items-center gap-2.5">
                <div className="flex items-center gap-1.5">
                  <span className="text-slate-400 font-semibold">Dự án:</span>
                  <input
                    type="text"
                    value={projectName}
                    onChange={(e) => handleProjectNameChange(e.target.value)}
                    placeholder="Tên dự án..."
                    className="w-36 px-2.5 py-1 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500 font-bold"
                  />
                </div>

                <div className="flex items-center gap-1.5">
                  <span className="text-slate-400 font-semibold">Cập nhật:</span>
                  <input
                    type="text"
                    value={updatedBy}
                    onChange={(e) => handleUpdatedByChange(e.target.value)}
                    placeholder="Người cập nhật..."
                    className="w-32 px-2.5 py-1 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500"
                  />
                </div>

                {/* Compact Logo upload */}
                <label className="cursor-pointer inline-flex items-center gap-1.5 px-2.5 py-1 bg-slate-800 hover:bg-slate-750 text-slate-300 border border-slate-700 rounded-lg text-xs font-semibold">
                  <Upload className="w-3 h-3 text-purple-400" />
                  <span>{logoBase64 ? 'Đổi Logo' : 'Logo B3'}</span>
                  <input type="file" accept="image/png,image/jpeg" onChange={handleLogoUpload} className="hidden" />
                </label>
                {logoBase64 && (
                  <img src={logoBase64} alt="Logo" className="h-6 max-w-16 object-contain rounded bg-slate-800 p-0.5 border border-slate-700" />
                )}
              </div>

              {/* 3 Folder Selector Buttons */}
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={handleSelectSource}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-purple-600 hover:bg-purple-500 text-white cursor-pointer transition-colors shadow-xs"
                  title="Chọn thư mục chứa file .dwg, .dxf"
                >
                  <FolderOpen className="w-3.5 h-3.5" />
                  <span>{sourceName || '1. Chọn Thư Mục CAD'}</span>
                  {rows.length > 0 && <span className="bg-purple-800 px-1.5 py-0.2 rounded text-[10px]">{rows.length}</span>}
                </button>

                <button
                  type="button"
                  onClick={handleSelectTemplate}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-750 text-slate-300 border border-slate-700 cursor-pointer"
                  title="Chọn thư mục template mẫu"
                >
                  <Layers className="w-3.5 h-3.5 text-amber-400" />
                  <span>{templateName ? `Mẫu: ${templateName}` : '2. Thư Mục Template'}</span>
                  {templateSubfolders.length > 0 && (
                    <span className="text-amber-400 text-[10px]">({templateSubfolders.length})</span>
                  )}
                </button>

                <button
                  type="button"
                  onClick={handleSelectOutput}
                  className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-750 text-slate-300 border border-slate-700 cursor-pointer"
                  title="Chọn thư mục đầu ra"
                >
                  <FolderSync className="w-3.5 h-3.5 text-cyan-400" />
                  <span>{outputName ? `Đích: ${outputName}` : '3. Thư Mục Đích'}</span>
                </button>
              </div>
            </div>

            {/* ACTION TOOLBAR: Clean, compact buttons */}
            <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-2.5 shadow-xs flex flex-wrap items-center justify-between gap-2.5 text-xs">
              {/* Left actions: Run dry, copy, zip, excel */}
              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={() => handleExecuteCopy(true)}
                  disabled={isRunning || rows.length === 0}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg font-bold bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 cursor-pointer disabled:opacity-40"
                >
                  <Eye className="w-3.5 h-3.5 text-cyan-400" />
                  <span>Chạy Thử</span>
                </button>

                <button
                  type="button"
                  onClick={() => handleExecuteCopy(false)}
                  disabled={isRunning || rows.length === 0}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg font-bold bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-500 hover:to-indigo-500 text-white shadow-xs cursor-pointer disabled:opacity-40"
                >
                  <Play className="w-3.5 h-3.5 fill-white" />
                  <span>{isRunning ? 'Đang Chạy...' : 'Copy Vào Thư Mục Đích'}</span>
                </button>

                <button
                  type="button"
                  onClick={handleDownloadZip}
                  disabled={isZipExporting || rows.length === 0}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg font-bold bg-indigo-950/60 hover:bg-indigo-900/80 text-indigo-300 border border-indigo-700/60 cursor-pointer disabled:opacity-40"
                >
                  <Archive className="w-3.5 h-3.5 text-indigo-400" />
                  <span>{isZipExporting ? 'Đang Nén...' : 'Tải Gói ZIP'}</span>
                </button>

                <button
                  type="button"
                  onClick={handleExportExcel}
                  disabled={rows.length === 0}
                  className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-xs cursor-pointer disabled:opacity-40"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Xuất Excel (.xlsx)</span>
                </button>
              </div>

              {/* Right: Quick filter & search */}
              <div className="flex items-center gap-2">
                <div className="relative w-44 sm:w-56">
                  <Search className="w-3 h-3 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    placeholder="Lọc Zone, System, Spool..."
                    className="w-full pl-7 pr-2.5 py-1 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500"
                  />
                </div>

                <select
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value as 'all' | 'valid' | 'invalid')}
                  className="bg-slate-950 border border-slate-700 rounded-lg px-2 py-1 text-xs text-slate-300"
                >
                  <option value="all">Tất cả ({rows.length})</option>
                  <option value="valid">Hợp lệ ({stats.valid})</option>
                  <option value="invalid">Lỗi ({stats.invalid})</option>
                </select>
              </div>
            </div>

            {/* DATA TABLE */}
            <div className="bg-slate-900/80 border border-slate-800 rounded-xl overflow-hidden shadow-xs">
              <div className="overflow-x-auto max-h-[calc(100vh-280px)] min-h-[300px]">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-900 sticky top-0 z-10 border-b border-slate-800 text-[11px] font-bold text-slate-300">
                    <tr>
                      <th className="py-2 px-2.5 text-center w-10">STT</th>
                      <th className="py-2 px-2.5 min-w-[200px]">FILE NAME</th>
                      <th className="py-2 px-2.5 min-w-[160px]">PIPE NUMBER</th>
                      <th className="py-2 px-2.5 text-center min-w-[65px]">ZONE</th>
                      <th className="py-2 px-2.5 text-center min-w-[90px]">SYSTEM</th>
                      <th className="py-2 px-2.5 text-center min-w-[100px]">PIPE SPOOL</th>
                      <th className="py-2 px-2.5 text-center min-w-[45px]">REV</th>
                      <th className="py-2 px-2.5 text-center min-w-[95px]">ISSUED DATE</th>
                      <th className="py-2 px-2.5 min-w-[120px]">REMARK</th>
                      <th className="py-2 px-2.5 text-center min-w-[70px]">STATUS</th>
                      <th className="py-2 px-2.5 text-center min-w-[80px]">SHORT</th>
                      <th className="py-2 px-2.5 text-center min-w-[80px]">CAD</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60 font-mono text-[11px]">
                    {sortedRows.length === 0 ? (
                      <tr>
                        <td colSpan={12} className="py-12 text-center text-slate-500 font-sans">
                          {rows.length === 0
                            ? 'Bấm "1. Chọn Thư Mục CAD" ở trên để nạp danh sách file .dwg & .dxf.'
                            : 'Không tìm thấy dòng khớp với bộ lọc.'}
                        </td>
                      </tr>
                    ) : (
                      sortedRows.map((r) => {
                        const isInvalid = !r.isValid;
                        return (
                          <tr key={r.id} className={`hover:bg-slate-850/60 ${isInvalid ? 'bg-rose-950/25 text-rose-200' : ''}`}>
                            <td className="py-1.5 px-2.5 text-center text-slate-500">{r.stt}</td>
                            <td className="py-1.5 px-2.5 font-semibold text-slate-200 truncate max-w-[220px]" title={r.fileName}>
                              {r.fileName}
                            </td>
                            <td className="py-1.5 px-2.5">
                              <input
                                type="text"
                                value={r.pipeNumber}
                                onChange={(e) => handleUpdateRow(r.id, 'pipeNumber', e.target.value)}
                                className="w-full bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none text-slate-300"
                              />
                            </td>
                            <td className="py-1.5 px-2.5 text-center font-bold text-amber-300">
                              <input
                                type="text"
                                value={r.zone}
                                onChange={(e) => handleUpdateRow(r.id, 'zone', e.target.value)}
                                className="w-12 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none"
                              />
                            </td>
                            <td className="py-1.5 px-2.5 text-center text-cyan-300">
                              <input
                                type="text"
                                value={r.system}
                                onChange={(e) => handleUpdateRow(r.id, 'system', e.target.value)}
                                className="w-20 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none"
                              />
                            </td>
                            <td className="py-1.5 px-2.5 text-center text-emerald-300">
                              <input
                                type="text"
                                value={r.pipeSpool}
                                onChange={(e) => handleUpdateRow(r.id, 'pipeSpool', e.target.value)}
                                className="w-24 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none"
                              />
                            </td>
                            <td className="py-1.5 px-2.5 text-center font-bold text-white">
                              <input
                                type="text"
                                value={r.rev}
                                onChange={(e) => handleUpdateRow(r.id, 'rev', e.target.value)}
                                className="w-8 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none"
                              />
                            </td>
                            <td className="py-1.5 px-2.5 text-center text-slate-400">
                              <input
                                type="text"
                                value={r.issuedDate}
                                onChange={(e) => handleUpdateRow(r.id, 'issuedDate', e.target.value)}
                                className="w-20 text-center bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none"
                              />
                            </td>
                            <td className="py-1.5 px-2.5">
                              <input
                                type="text"
                                value={r.remark}
                                placeholder="Ghi chú..."
                                onChange={(e) => handleUpdateRow(r.id, 'remark', e.target.value)}
                                className="w-full bg-transparent border-b border-transparent hover:border-slate-700 focus:border-purple-500 focus:outline-none text-slate-400 text-[10px]"
                              />
                            </td>
                            <td className="py-1.5 px-2.5 text-center">
                              <span className="px-1.5 py-0.2 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                                {r.status}
                              </span>
                            </td>
                            <td className="py-1.5 px-2.5 text-center font-bold text-purple-300">
                              {r.systemShort || '-'}
                            </td>
                            <td className="py-1.5 px-2.5 text-center">
                              <div className="flex items-center justify-center gap-1">
                                {r.hasDWG && <span className="px-1 py-0.2 rounded text-[9px] font-bold bg-blue-500/20 text-blue-300">DWG</span>}
                                {r.hasDXF && <span className="px-1 py-0.2 rounded text-[9px] font-bold bg-purple-500/20 text-purple-300">DXF</span>}
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
          </>
        )}

        {/* TAB 2: ISO RELEASED TOOL (Simplified) */}
        {activeTab === 'released' && (
          <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-5 shadow-xs max-w-3xl mx-auto space-y-4 text-xs">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h2 className="text-sm font-bold text-white flex items-center gap-2">
                <Archive className="w-4 h-4 text-purple-400" />
                <span>ISO Released (Gộp Bản Vẽ)</span>
              </h2>
              <button
                type="button"
                onClick={handleSelectRelParent}
                className="px-3 py-1.5 rounded-lg font-bold bg-purple-600 hover:bg-purple-500 text-white cursor-pointer"
              >
                {relParentName ? `Thư mục: ${relParentName}` : 'Chọn Thư Mục Mẹ'}
              </button>
            </div>

            {relSubfolders.length > 0 && (
              <div className="space-y-3">
                <div className="flex items-center justify-between text-[11px] text-slate-400">
                  <span>Chọn các thư mục con cần copy ({relSubfolders.filter((f) => f.checked).length}/{relSubfolders.length})</span>
                  <div className="flex gap-2">
                    <button onClick={() => setRelSubfolders((prev) => prev.map((item) => ({ ...item, checked: true })))} className="text-purple-400 hover:underline cursor-pointer">
                      Chọn tất cả
                    </button>
                    <span>|</span>
                    <button onClick={() => setRelSubfolders((prev) => prev.map((item) => ({ ...item, checked: false })))} className="text-slate-400 hover:underline cursor-pointer">
                      Bỏ chọn
                    </button>
                  </div>
                </div>

                <div className="p-2.5 bg-slate-950 border border-slate-800 rounded-lg max-h-56 overflow-y-auto grid grid-cols-1 sm:grid-cols-2 gap-1.5 font-mono text-[11px]">
                  {relSubfolders.map((f, idx) => (
                    <label key={idx} className="flex items-center gap-2 p-1.5 rounded hover:bg-slate-900 cursor-pointer">
                      <input
                        type="checkbox"
                        checked={f.checked}
                        onChange={(e) => {
                          const checked = e.target.checked;
                          setRelSubfolders((prev) => prev.map((item, i) => (i === idx ? { ...item, checked } : item)));
                        }}
                        className="rounded border-slate-700 text-purple-600"
                      />
                      <span className="truncate text-slate-300">{f.name}</span>
                    </label>
                  ))}
                </div>

                <button
                  type="button"
                  onClick={handleExecuteIsoReleased}
                  disabled={relIsProcessing}
                  className="w-full py-2.5 rounded-lg font-bold bg-purple-600 hover:bg-purple-500 text-white cursor-pointer disabled:opacity-40"
                >
                  {relIsProcessing ? 'Đang Sao Chép...' : 'Tạo Thư Mục "Iso released yyyymmdd" & Copy Files'}
                </button>
              </div>
            )}

            {relLogs.length > 0 && (
              <div className="p-2.5 bg-slate-950 border border-slate-800 rounded-lg max-h-36 overflow-y-auto font-mono text-[11px] space-y-0.5 text-slate-400">
                {relLogs.map((m, i) => (
                  <div key={i}>{m}</div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* TAB 3: UPDATE OLD REV TOOL (Simplified) */}
        {activeTab === 'oldrev' && (
          <div className="bg-slate-900/90 border border-slate-800 rounded-xl p-5 shadow-xs max-w-3xl mx-auto space-y-4 text-xs">
            <div className="flex items-center justify-between border-b border-slate-800 pb-3">
              <h2 className="text-sm font-bold text-white flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-cyan-400" />
                <span>Update Old Rev</span>
              </h2>
              <label className="cursor-pointer px-3 py-1.5 rounded-lg font-bold bg-purple-600 hover:bg-purple-500 text-white">
                <span>Upload File Excel (.xlsx)</span>
                <input type="file" accept=".xlsx, .xls" onChange={handleUploadRevExcel} className="hidden" />
              </label>
            </div>

            {revRawMatrix && (
              <div className="space-y-3">
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2.5">
                  <div>
                    <label className="text-slate-400 block mb-1">Sheet:</label>
                    <select
                      value={revActiveSheet}
                      onChange={(e) => setRevActiveSheet(e.target.value)}
                      className="w-full bg-slate-950 border border-slate-700 rounded-lg p-1.5 text-white"
                    >
                      {revSheetNames.map((s) => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label className="text-slate-400 block mb-1">Cột Key (L):</label>
                    <input
                      type="number"
                      value={revKeyCol + 1}
                      onChange={(e) => setRevKeyCol(Math.max(0, Number(e.target.value) - 1))}
                      className="w-full bg-slate-950 border border-slate-700 rounded-lg p-1.5 text-white"
                    />
                  </div>

                  <div>
                    <label className="text-slate-400 block mb-1">Cột Rev (P):</label>
                    <input
                      type="number"
                      value={revRevCol + 1}
                      onChange={(e) => setRevRevCol(Math.max(0, Number(e.target.value) - 1))}
                      className="w-full bg-slate-950 border border-slate-700 rounded-lg p-1.5 text-white"
                    />
                  </div>

                  <div>
                    <label className="text-slate-400 block mb-1">Cột Đánh Dấu (S):</label>
                    <input
                      type="number"
                      value={revMarkCol + 1}
                      onChange={(e) => setRevMarkCol(Math.max(0, Number(e.target.value) - 1))}
                      className="w-full bg-slate-950 border border-slate-700 rounded-lg p-1.5 text-white"
                    />
                  </div>
                </div>

                <div className="flex gap-2 pt-1">
                  <button
                    type="button"
                    onClick={handleExecuteMarkOldRev}
                    className="flex-1 py-2 rounded-lg font-bold bg-cyan-600 hover:bg-cyan-500 text-white cursor-pointer"
                  >
                    Đánh Dấu "old rev"
                  </button>

                  {revPreviewCount !== null && (
                    <button
                      type="button"
                      onClick={handleDownloadMarkedExcel}
                      className="flex-1 py-2 rounded-lg font-bold bg-emerald-600 hover:bg-emerald-500 text-white cursor-pointer flex items-center justify-center gap-1.5"
                    >
                      <Download className="w-3.5 h-3.5" />
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
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-xl w-full p-4 space-y-3 shadow-2xl flex flex-col max-h-[85vh]">
            <div className="flex items-center justify-between border-b border-slate-800 pb-2.5">
              <div className="flex items-center gap-2">
                <Clock className="w-4 h-4 text-purple-400" />
                <h3 className="text-sm font-bold text-white uppercase tracking-wider">Nhật Ký Thực Thi</h3>
                <span className="text-[10px] font-mono px-2 py-0.2 rounded-full bg-slate-800 text-slate-300">
                  {executionLogs.length} mục
                </span>
              </div>
              <button
                onClick={() => setIsLogsModalOpen(false)}
                className="text-slate-400 hover:text-white p-1 rounded-lg hover:bg-slate-800 cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            {/* Progress bar inside popup */}
            {runProgress && (
              <div className="space-y-1 text-xs">
                <div className="flex justify-between font-semibold text-purple-300 text-[11px]">
                  <span>{runProgress.statusText}</span>
                  <span>{runProgress.percentage}%</span>
                </div>
                <div className="w-full h-1.5 bg-slate-800 rounded-full overflow-hidden">
                  <div
                    className="h-full bg-gradient-to-r from-purple-500 to-cyan-400 transition-all duration-150"
                    style={{ width: `${runProgress.percentage}%` }}
                  />
                </div>
              </div>
            )}

            {/* Log Stream */}
            <div className="flex-1 min-h-[220px] max-h-[380px] p-3 bg-slate-950 border border-slate-800 rounded-xl overflow-y-auto font-mono text-[11px] space-y-1 text-slate-300">
              {executionLogs.length === 0 ? (
                <p className="text-slate-600 italic">Chưa có nhật ký ghi nhận.</p>
              ) : (
                executionLogs.map((logMsg, idx) => (
                  <div
                    key={idx}
                    className={`${
                      logMsg.includes('[LỖI]')
                        ? 'text-rose-400'
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
                className="px-4 py-1.5 rounded-lg font-bold bg-slate-800 hover:bg-slate-700 text-white cursor-pointer"
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
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-sm w-full p-5 space-y-3 shadow-2xl">
            <h3 className="text-sm font-bold text-white">Xóa trắng bảng dữ liệu?</h3>
            <p className="text-xs text-slate-400 leading-relaxed">
              Các bản vẽ trên bảng sẽ bị xóa. Cấu hình thư mục và tên dự án được giữ nguyên.
            </p>
            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                onClick={() => setIsResetConfirmOpen(false)}
                className="px-3 py-1.5 rounded-lg text-xs font-semibold text-slate-400 hover:bg-slate-800 cursor-pointer"
              >
                Hủy
              </button>
              <button
                onClick={handleConfirmReset}
                className="px-3.5 py-1.5 rounded-lg text-xs font-bold bg-rose-600 hover:bg-rose-500 text-white cursor-pointer"
              >
                Xóa sạch
              </button>
            </div>
          </div>
        </div>
      )}

      {/* COMPACT FOOTER */}
      <footer className="border-t border-slate-800/80 bg-slate-950 py-2.5 text-xs text-slate-500 text-center">
        <div className="max-w-7xl mx-auto px-4 flex items-center justify-between text-[11px]">
          <span>WebToolRush &bull; ISO Drawing Manager</span>
          <span className="font-mono text-slate-400">100% Client-Side In-Memory Execution</span>
        </div>
      </footer>
    </div>
  );
};
