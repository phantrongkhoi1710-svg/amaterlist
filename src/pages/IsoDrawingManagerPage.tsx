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
  CheckCircle2,
  FileCheck,
  Search,
  Upload,
  ArrowLeft,
  LayoutGrid,
  RefreshCw,
  Sliders,
  ShieldCheck,
  Layers,
  FileCode,
  FileWarning,
  Eye,
  Check,
  X,
  Sparkles,
  Info,
  Clock,
  Archive,
  ExternalLink,
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
  isInIframe,
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
  // Navigation tabs
  const [activeTab, setActiveTab] = useState<'main' | 'released' | 'oldrev'>('main');

  // Environment checks
  const inIframe = useMemo(() => isInIframe(), []);
  const [hasFSAccess, setHasFSAccess] = useState<boolean>(true);

  // Hidden File Inputs for universal directory selection
  const sourceInputRef = useRef<HTMLInputElement>(null);
  const templateInputRef = useRef<HTMLInputElement>(null);
  const relParentInputRef = useRef<HTMLInputElement>(null);

  // In-app Alert / Toast notification
  const [noticeMessage, setNoticeMessage] = useState<string | null>(null);

  // STEP 1: Setup State
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
  const [scanProgressCount, setScanProgressCount] = useState<number>(0);

  // STEP 2: Table Filter & Search
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [statusFilter, setStatusFilter] = useState<'all' | 'valid' | 'invalid'>('all');
  const [sortField, setSortField] = useState<keyof IsoDrawingRow>('stt');
  const [sortAsc, setSortAsc] = useState<boolean>(true);

  // STEP 3: Execution & ZIP Fallback State
  const [isRunning, setIsRunning] = useState<boolean>(false);
  const [isZipExporting, setIsZipExporting] = useState<boolean>(false);
  const [runProgress, setRunProgress] = useState<CopyProgress | null>(null);
  const [executionLogs, setExecutionLogs] = useState<string[]>([]);
  const [isDryRun, setIsDryRun] = useState<boolean>(false);

  // Modal confirm reset
  const [isResetConfirmOpen, setIsResetConfirmOpen] = useState<boolean>(false);

  // TOOL 6: ISO Released State
  const [relParentHandle, setRelParentHandle] = useState<FileSystemDirectoryHandle | null>(null);
  const [relParentName, setRelParentName] = useState<string>('');
  const [relSubfolders, setRelSubfolders] = useState<{ name: string; checked: boolean }[]>([]);
  const [relIsProcessing, setRelIsProcessing] = useState<boolean>(false);
  const [relLogs, setRelLogs] = useState<string[]>([]);

  // TOOL 7: Update Old Rev State
  const [revRawMatrix, setRevRawMatrix] = useState<unknown[][] | null>(null);
  const [revSheetNames, setRevSheetNames] = useState<string[]>([]);
  const [revActiveSheet, setRevActiveSheet] = useState<string>('');
  const [revKeyCol, setRevKeyCol] = useState<number>(11); // Col L is 11 (0-indexed)
  const [revRevCol, setRevRevCol] = useState<number>(15); // Col P is 15 (0-indexed)
  const [revMarkCol, setRevMarkCol] = useState<number>(18); // Col S is 18 (0-indexed)
  const [revStartRow, setRevStartRow] = useState<number>(1); // Row 2 is index 1
  const [revPreviewCount, setRevPreviewCount] = useState<number | null>(null);
  const [revProcessedBlob, setRevProcessedBlob] = useState<Blob | null>(null);

  const logsEndRef = useRef<HTMLDivElement>(null);

  const showToast = (msg: string, durationMs: number = 4000) => {
    setNoticeMessage(msg);
    setTimeout(() => {
      setNoticeMessage((curr) => (curr === msg ? null : curr));
    }, durationMs);
  };

  // Initialize capability & load local configs
  useEffect(() => {
    const supported = isFileSystemAccessSupported();
    setHasFSAccess(supported);

    const savedUser = localStorage.getItem('iso_updated_by');
    if (savedUser) setUpdatedBy(savedUser);

    const savedProject = localStorage.getItem('iso_project_name');
    if (savedProject) setProjectName(savedProject);

    // Try reading logo from IDB
    getHandleFromIDB<string>('iso_logo_base64').then((res) => {
      if (res) setLogoBase64(res);
    });
  }, []);

  // Persist updatedBy
  const handleUpdatedByChange = (val: string) => {
    setUpdatedBy(val);
    localStorage.setItem('iso_updated_by', val);
  };

  // Persist projectName
  const handleProjectNameChange = (val: string) => {
    setProjectName(val);
    localStorage.setItem('iso_project_name', val);
  };

  // Handle Logo Upload
  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const b64 = evt.target?.result as string;
      setLogoBase64(b64);
      saveHandleToIDB('iso_logo_base64', b64);
      showToast('Đã lưu Logo dự án thành công!');
    };
    reader.readAsDataURL(file);
  };

  // Universal Source Folder Selection Handler
  const handleSelectSource = async () => {
    // If in iframe or no direct showDirectoryPicker support, trigger webkitdirectory directly
    if (!isFileSystemAccessSupported()) {
      sourceInputRef.current?.click();
      return;
    }

    try {
      const handle = await pickDirectory('iso-source');
      if (handle) {
        setSourceHandle(handle);
        setSourceName(handle.name);
        showToast(`Đã chọn thư mục: "${handle.name}". Đang quét file CAD...`);
        await triggerScanFromHandle(handle);
      }
    } catch (err) {
      // If blocked by cross-origin iframe security or permission error, seamlessly trigger file picker
      if ((err as Error).name !== 'AbortError') {
        sourceInputRef.current?.click();
      }
    }
  };

  // Scanner from Handle
  const triggerScanFromHandle = async (handle: FileSystemDirectoryHandle) => {
    setIsScanning(true);
    setScanProgressCount(0);
    try {
      const map = await scanCadDirectory(handle, (count) => {
        setScanProgressCount(count);
      });
      setFileMap(map);
      populateRowsFromMap(map);
    } catch (err) {
      showToast(`Lỗi quét thư mục: ${(err as Error).message}`);
    } finally {
      setIsScanning(false);
    }
  };

  // Webkitdirectory fallback handler for Source Folder
  const handleSourceFilesChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const firstPath = (files[0] as unknown as { webkitRelativePath?: string }).webkitRelativePath || '';
    const folderName = firstPath ? firstPath.split('/')[0] : 'CAD_FOLDER';
    setSourceName(folderName);
    setSourceHandle(null);

    // Immediately scan and map files in browser memory
    const map = scanCadFileList(files);
    setFileMap(map);
    populateRowsFromMap(map);

    showToast(`Đã nạp ${map.size} bản vẽ từ thư mục "${folderName}"!`);
  };

  // Populate rows from File Map
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

  // Universal Template Folder Selection Handler
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
        showToast(`Đã nhận diện ${sub.length} subfolders từ Template "${handle.name}"!`);
      }
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        templateInputRef.current?.click();
      }
    }
  };

  // Webkitdirectory fallback handler for Template Folder
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

    showToast(`Đã nhận diện ${sub.length} subfolders từ Template "${folderName}"!`);
  };

  const updateRowsWithTemplateSubfolders = (sub: string[]) => {
    setRows((prev) =>
      prev.map((r) => ({
        ...r,
        targetFolder: matchTemplateFolder(r.systemShort, sub),
      })),
    );
  };

  // Universal Output Folder Selection Handler
  const handleSelectOutput = async () => {
    if (!isFileSystemAccessSupported()) {
      showToast('Đang chạy trong iFrame: Hệ thống hỗ trợ nút "Tải Gói ZIP (Cấu Trúc S{ZONE}/{System}/)" ở Bước 3 tiện lợi!');
      return;
    }

    try {
      const handle = await pickDirectory('iso-output');
      if (handle) {
        setOutputHandle(handle);
        setOutputName(handle.name);
        showToast(`Đã chọn thư mục đầu ra: "${handle.name}"`);
      }
    } catch (err) {
      if ((err as Error).name !== 'AbortError') {
        showToast('Trình duyệt hạn chế chọn thư mục đầu ra trực tiếp trong iFrame. Hãy dùng tính năng tải ZIP tự động!');
      }
    }
  };

  // Manual Trigger Scan Source
  const handleManualScanSource = async () => {
    if (sourceHandle) {
      await triggerScanFromHandle(sourceHandle);
    } else {
      sourceInputRef.current?.click();
    }
  };

  // Inline row edits
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

  // Filtered & sorted table data
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

  // RUN: Execute File Copy
  const handleExecuteCopy = async (dryRun: boolean) => {
    if (rows.length === 0) {
      showToast('Chưa có dữ liệu bản vẽ. Vui lòng chọn Thư mục CAD nguồn trước.');
      return;
    }

    if (!dryRun && !outputHandle) {
      showToast('Thư mục đầu ra chưa được chọn trực tiếp. Đang chuẩn bị gói ZIP tự động...');
      await handleDownloadZip();
      return;
    }

    setIsRunning(true);
    setIsDryRun(dryRun);
    setRunProgress(null);
    setExecutionLogs([]);

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
      showToast(`Lỗi thực thi: ${(err as Error).message}`);
    } finally {
      setIsRunning(false);
    }
  };

  // ZIP Download Fallback (Creates S{ZONE}/{System}/ tree inside zip)
  const handleDownloadZip = async () => {
    if (rows.length === 0) {
      showToast('Chưa có dữ liệu bản vẽ để xuất ZIP. Vui lòng quét thư mục nguồn trước.');
      return;
    }

    setIsZipExporting(true);
    setExecutionLogs((prev) => [...prev, '[ZIP] Bắt đầu đóng gói toàn bộ bản vẽ vào cây thư mục S{ZONE}/{System}/...']);

    try {
      const blob = await createZipFallback(rows, fileMap, (pct) => {
        setRunProgress({
          total: rows.length,
          current: Math.round((pct / 100) * rows.length),
          percentage: Math.round(pct),
          currentFileName: 'Đang nén file...',
          statusText: `Đang nén gói ZIP: ${Math.round(pct)}%`,
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

      setExecutionLogs((prev) => [
        ...prev,
        `[HOÀN THÀNH] Đã tải về thành công gói ZIP: ${zipName}`,
      ]);
      showToast(`Đã xuất gói ZIP thành công: ${zipName}`);
    } catch (err) {
      showToast(`Lỗi tạo file ZIP: ${(err as Error).message}`);
    } finally {
      setIsZipExporting(false);
      setRunProgress(null);
    }
  };

  // EXPORT EXCEL ISSUED HISTORY
  const handleExportExcel = async () => {
    if (rows.length === 0) {
      showToast('Chưa có dữ liệu bản vẽ để xuất Excel.');
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

      showToast(`Đã xuất file Excel: ${fileName}`);
    } catch (err) {
      showToast(`Lỗi xuất file Excel: ${(err as Error).message}`);
    }
  };

  // RESET ALL
  const handleConfirmReset = () => {
    setRows([]);
    setFileMap(new Map());
    setRunProgress(null);
    setExecutionLogs([]);
    setIsResetConfirmOpen(false);
    showToast('Đã xóa sạch bảng dữ liệu bản vẽ.');
  };

  // TOOL 6: ISO Released Actions
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
    showToast(`Đã nhận diện ${sub.length} thư mục con từ "${folderName}"!`);
  };

  const handleToggleAllRel = (checked: boolean) => {
    setRelSubfolders((prev) => prev.map((item) => ({ ...item, checked })));
  };

  const handleExecuteIsoReleased = async () => {
    if (relSubfolders.length === 0) {
      showToast('Vui lòng chọn Thư mục mẹ chứa các thư mục con!');
      return;
    }

    const selectedFolders = relSubfolders.filter((f) => f.checked);
    if (selectedFolders.length === 0) {
      showToast('Vui lòng chọn ít nhất một thư mục con!');
      return;
    }

    const d = new Date();
    const yyyymmdd = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
    const targetFolderName = `Iso released ${yyyymmdd}`;

    setRelIsProcessing(true);
    setRelLogs([]);
    const log = (msg: string) => setRelLogs((prev) => [...prev, msg]);

    log(`[BẮT ĐẦU] Tạo thư mục ${targetFolderName}...`);

    try {
      if (outputHandle) {
        const targetDirHandle = await outputHandle.getDirectoryHandle(targetFolderName, { create: true });
        let totalCopied = 0;
        let totalOverwritten = 0;

        for (const item of selectedFolders) {
          log(`[QUÉT] Đang sao chép từ ${item.name}...`);
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

            if (exists) {
              totalOverwritten++;
              log(`  -> [GHI ĐÈ] ${fName}`);
            } else {
              totalCopied++;
              log(`  -> [MỚI] ${fName}`);
            }

            const srcFile = await (fHandle as FileSystemFileHandle).getFile();
            const destHandle = await targetDirHandle.getFileHandle(fName, { create: true });
            const writable = await (destHandle as unknown as { createWritable: () => Promise<FileSystemWritableFileStream> }).createWritable();
            await writable.write(srcFile);
            await writable.close();
          }
        }

        log(`[HOÀN THÀNH] Đã hoàn tất gộp bản vẽ vào ${targetFolderName}!`);
        log(`Tổng file mới: ${totalCopied} | File ghi đè: ${totalOverwritten}`);
      } else {
        log(`[THÔNG BÁO] Không có thư mục đích trên đĩa, vui lòng sử dụng tính năng tải file ZIP.`);
      }
    } catch (err) {
      log(`[LỖI] ${(err as Error).message}`);
    } finally {
      setRelIsProcessing(false);
    }
  };

  // TOOL 7: Update Old Rev Actions
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
        showToast(`Đã nạp file Excel với ${matrix.length} dòng.`);
      } catch (err) {
        showToast(`Không thể đọc file Excel: ${(err as Error).message}`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleExecuteMarkOldRev = () => {
    if (!revRawMatrix || revRawMatrix.length === 0) {
      showToast('Vui lòng upload file Excel trước.');
      return;
    }

    const { updatedMatrix, markedCount, totalProcessed } = markOldRevisionsInMatrix(
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
    showToast(`Đã đánh dấu "old rev" cho ${markedCount} dòng!`);
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
      {/* Hidden File Inputs for Full Browser Compatibility */}
      <input
        ref={sourceInputRef}
        type="file"
        webkitdirectory=""
        multiple
        onChange={handleSourceFilesChange}
        className="hidden"
      />
      <input
        ref={templateInputRef}
        type="file"
        webkitdirectory=""
        multiple
        onChange={handleTemplateFilesChange}
        className="hidden"
      />
      <input
        ref={relParentInputRef}
        type="file"
        webkitdirectory=""
        multiple
        onChange={handleRelParentFilesChange}
        className="hidden"
      />

      {/* Top Floating Notification / Toast */}
      {noticeMessage && (
        <div className="fixed top-4 right-4 z-50 p-4 rounded-2xl bg-purple-900/95 border border-purple-400 text-white shadow-2xl text-xs font-semibold flex items-center gap-2.5 animate-in fade-in slide-in-from-top-3 max-w-md">
          <Sparkles className="w-4 h-4 text-purple-300 shrink-0" />
          <span className="flex-1">{noticeMessage}</span>
          <button
            onClick={() => setNoticeMessage(null)}
            className="text-purple-300 hover:text-white ml-2 text-base font-bold cursor-pointer"
          >
            &times;
          </button>
        </div>
      )}

      {/* Top Navigation Bar */}
      <header className="sticky top-0 z-40 bg-slate-900/95 backdrop-blur-md border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between gap-4">
          <div className="flex items-center space-x-3">
            <Link
              to="/"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white border border-slate-700 text-xs font-bold transition-all shadow-sm active:scale-95 cursor-pointer group"
              title="Quay lại WebToolRush Hub"
            >
              <ArrowLeft className="w-3.5 h-3.5 text-slate-400 group-hover:-translate-x-0.5 transition-transform" />
              <LayoutGrid className="w-3.5 h-3.5 text-cyan-400" />
              <span className="hidden md:inline font-semibold text-slate-200">Tool Hub</span>
            </Link>

            <div className="w-8 h-8 rounded-lg bg-gradient-to-tr from-purple-600 via-indigo-600 to-cyan-500 flex items-center justify-center text-white shadow-md shadow-purple-600/30">
              <FolderSync className="w-5 h-5 text-white" />
            </div>

            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base font-black tracking-tight text-white uppercase">
                  ISO DRAWING MANAGER
                </h1>
                <span className="text-[10px] font-bold px-2 py-0.5 rounded-full bg-purple-500/20 text-purple-300 border border-purple-500/30">
                  CAD Release Suite
                </span>
              </div>
              <p className="text-[11px] text-slate-400 hidden sm:block">
                Tự động hóa bóc tách tên bản vẽ CAD .dwg/.dxf, phân nhóm Zone/System và xuất Issued History
              </p>
            </div>
          </div>

          {/* Quick Metrics */}
          <div className="flex items-center gap-2">
            <div className="hidden lg:flex items-center bg-slate-800/90 rounded-lg px-3 py-1.5 border border-slate-700/60 text-xs">
              <span className="text-slate-400 mr-1.5">Project:</span>
              <span className="font-bold text-cyan-400">{projectName}</span>
              <span className="mx-2 text-slate-600">|</span>
              <span className="text-slate-400 mr-1.5">Bản vẽ:</span>
              <span className="font-bold text-emerald-400">{stats.total} files</span>
            </div>

            <button
              type="button"
              onClick={() => setIsResetConfirmOpen(true)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-rose-950/40 hover:bg-rose-900/60 text-rose-300 border border-rose-800/60 transition-colors cursor-pointer"
              title="Xóa trắng bảng dữ liệu để làm việc với mẻ mới"
            >
              <Trash2 className="w-3.5 h-3.5 text-rose-400" />
              <span>Reset</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Tabs Navigation */}
      <div className="bg-slate-900 border-b border-slate-800">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex items-center gap-2 py-2 overflow-x-auto scrollbar-none">
          <button
            onClick={() => setActiveTab('main')}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'main'
                ? 'bg-gradient-to-r from-purple-600 to-indigo-600 text-white shadow-md shadow-purple-600/30'
                : 'bg-slate-800/80 text-slate-300 hover:bg-slate-800 hover:text-white'
            }`}
          >
            <FolderSync className="w-4 h-4" />
            <span>1. Quy Trình Bóc Tách &amp; Copy CAD</span>
          </button>

          <button
            onClick={() => setActiveTab('released')}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'released'
                ? 'bg-gradient-to-r from-purple-600 to-indigo-600 text-white shadow-md shadow-purple-600/30'
                : 'bg-slate-800/80 text-slate-300 hover:bg-slate-800 hover:text-white'
            }`}
          >
            <Archive className="w-4 h-4" />
            <span>2. Công Cụ ISO Released (Gộp Folder)</span>
          </button>

          <button
            onClick={() => setActiveTab('oldrev')}
            className={`inline-flex items-center gap-2 px-4 py-2 rounded-xl text-xs font-bold transition-all cursor-pointer ${
              activeTab === 'oldrev'
                ? 'bg-gradient-to-r from-purple-600 to-indigo-600 text-white shadow-md shadow-purple-600/30'
                : 'bg-slate-800/80 text-slate-300 hover:bg-slate-800 hover:text-white'
            }`}
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>3. Cập Nhật Old Rev (Excel)</span>
          </button>
        </div>
      </div>

      {/* Main Content Area */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-1 w-full space-y-6">
        {/* TAB 1: MAIN WORKFLOW */}
        {activeTab === 'main' && (
          <div className="space-y-6">
            {/* STEP 1: SETUP PANEL */}
            <section className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-sm space-y-4">
              <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                <div className="flex items-center gap-2.5">
                  <span className="w-6 h-6 rounded-full bg-purple-600 text-white text-xs font-black flex items-center justify-center">
                    1
                  </span>
                  <h2 className="text-sm font-bold text-white uppercase tracking-wider">
                    Cấu Hình Dự Án &amp; Chọn Thư Mục Làm Việc
                  </h2>
                </div>
                <span className="text-xs text-slate-400 font-medium">
                  Client-side 100% &bull; Không tải file lên server
                </span>
              </div>

              {/* Form Grid */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                {/* Project Name */}
                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1.5">
                    Tên Dự Án (Project Name):
                  </label>
                  <input
                    type="text"
                    value={projectName}
                    onChange={(e) => handleProjectNameChange(e.target.value)}
                    placeholder="VD: VARD 9 80 1005"
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-xs text-white focus:outline-none focus:ring-2 focus:ring-purple-500"
                  />
                  <p className="text-[10px] text-slate-500 mt-1">
                    4 ký tự cuối sẽ được dùng để đặt tên file Excel xuất ra.
                  </p>
                </div>

                {/* Updated by */}
                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1.5">
                    Người Cập Nhật (Updated by):
                  </label>
                  <input
                    type="text"
                    value={updatedBy}
                    onChange={(e) => handleUpdatedByChange(e.target.value)}
                    placeholder="Nhập tên kỹ sư / người phụ trách..."
                    className="w-full px-3 py-2 bg-slate-950 border border-slate-700 rounded-xl text-xs text-white focus:outline-none focus:ring-2 focus:ring-purple-500"
                  />
                  <p className="text-[10px] text-slate-500 mt-1">
                    Tự động lưu vào trình duyệt (localStorage) cho các lần sau.
                  </p>
                </div>

                {/* Logo Upload */}
                <div>
                  <label className="block text-xs font-semibold text-slate-300 mb-1.5">
                    Logo Dự Án (Xuất vào ô B3 của Excel):
                  </label>
                  <div className="flex items-center gap-2">
                    <label className="flex-1 cursor-pointer flex items-center justify-center gap-2 px-3 py-2 bg-slate-800 hover:bg-slate-750 text-slate-300 border border-slate-700 rounded-xl text-xs font-semibold transition-colors">
                      <Upload className="w-3.5 h-3.5 text-purple-400" />
                      <span>{logoBase64 ? 'Đổi Logo Khác' : 'Chọn Logo (PNG/JPG)'}</span>
                      <input
                        type="file"
                        accept="image/png,image/jpeg"
                        onChange={handleLogoUpload}
                        className="hidden"
                      />
                    </label>
                    {logoBase64 && (
                      <div className="w-10 h-9 rounded-lg bg-slate-800 border border-slate-700 p-1 flex items-center justify-center shrink-0">
                        <img src={logoBase64} alt="Logo" className="max-h-full max-w-full object-contain" />
                      </div>
                    )}
                  </div>
                  <p className="text-[10px] text-slate-500 mt-1">
                    Chiều rộng khuyến nghị 200px. Nếu không có, Excel sẽ hiển thị "Logo not found".
                  </p>
                </div>
              </div>

              {/* Folder Selector Strip */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 pt-2 border-t border-slate-800/80">
                {/* 1. Source Folder */}
                <div className="p-3.5 rounded-xl bg-slate-950/70 border border-slate-800 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-purple-300 flex items-center gap-1.5">
                      <FolderOpen className="w-4 h-4 text-purple-400" />
                      1. Thư mục CAD nguồn
                    </span>
                    <button
                      type="button"
                      onClick={handleSelectSource}
                      className="px-3 py-1.5 rounded-lg text-xs font-bold bg-purple-600 hover:bg-purple-500 text-white transition-all shadow-md shadow-purple-600/30 cursor-pointer active:scale-95"
                    >
                      {sourceName ? 'Đổi Thư Mục' : 'Chọn Thư Mục'}
                    </button>
                  </div>
                  <p className="text-xs font-mono text-slate-300 truncate" title={sourceName}>
                    {sourceName ? (
                      <span className="text-emerald-400 font-bold">{sourceName}</span>
                    ) : (
                      'Bấm nút để chọn thư mục chứa file .dwg, .dxf'
                    )}
                  </p>
                </div>

                {/* 2. Template Folder */}
                <div className="p-3.5 rounded-xl bg-slate-950/70 border border-slate-800 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-amber-300 flex items-center gap-1.5">
                      <Layers className="w-4 h-4 text-amber-400" />
                      2. Thư mục Template mẫu
                    </span>
                    <button
                      type="button"
                      onClick={handleSelectTemplate}
                      className="px-3 py-1.5 rounded-lg text-xs font-bold bg-amber-600 hover:bg-amber-500 text-white transition-all shadow-md shadow-amber-600/30 cursor-pointer active:scale-95"
                    >
                      {templateName ? 'Đổi Thư Mục' : 'Chọn Thư Mục'}
                    </button>
                  </div>
                  <p className="text-xs font-mono text-slate-300 truncate" title={templateName}>
                    {templateName ? (
                      <span className="text-amber-400 font-bold">
                        {templateName} ({templateSubfolders.length} subfolders)
                      </span>
                    ) : (
                      'Chưa chọn (Dùng để map tên System folder)'
                    )}
                  </p>
                </div>

                {/* 3. Output Folder */}
                <div className="p-3.5 rounded-xl bg-slate-950/70 border border-slate-800 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-bold text-cyan-300 flex items-center gap-1.5">
                      <FolderSync className="w-4 h-4 text-cyan-400" />
                      3. Thư mục đầu ra (Output)
                    </span>
                    <button
                      type="button"
                      onClick={handleSelectOutput}
                      className="px-3 py-1.5 rounded-lg text-xs font-bold bg-cyan-600 hover:bg-cyan-500 text-white transition-all shadow-md shadow-cyan-600/30 cursor-pointer active:scale-95"
                    >
                      {outputName ? 'Đổi Thư Mục' : 'Chọn Thư Mục'}
                    </button>
                  </div>
                  <p className="text-xs font-mono text-slate-300 truncate" title={outputName}>
                    {outputName ? (
                      <span className="text-cyan-400 font-bold">{outputName}</span>
                    ) : (
                      'Nơi lưu hoặc bấm "Tải Gói ZIP" ở Bước 3'
                    )}
                  </p>
                </div>
              </div>

              {/* Action Scan Button */}
              <div className="pt-2 flex items-center justify-between">
                <div className="text-xs text-slate-400">
                  {isScanning && (
                    <span className="text-purple-400 font-semibold animate-pulse">
                      Đang đọc danh mục file: {scanProgressCount} files...
                    </span>
                  )}
                  {rows.length > 0 && !isScanning && (
                    <span className="text-emerald-400 font-medium">
                      &check; Đã nhận diện và bóc tách thành công {rows.length} bản vẽ.
                    </span>
                  )}
                </div>
                <button
                  type="button"
                  onClick={handleManualScanSource}
                  disabled={isScanning}
                  className="inline-flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs font-bold bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-500 hover:to-indigo-500 text-white shadow-lg shadow-purple-600/30 disabled:opacity-50 disabled:cursor-not-allowed transition-all cursor-pointer active:scale-95"
                >
                  <RefreshCw className={`w-4 h-4 ${isScanning ? 'animate-spin' : ''}`} />
                  <span>{isScanning ? 'Đang Quét File CAD...' : 'Quét Lại Thư Mục CAD'}</span>
                </button>
              </div>
            </section>

            {/* STEP 2: DATA TABLE PANEL */}
            <section className="bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-sm space-y-4">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-slate-800 pb-3">
                <div className="flex items-center gap-2.5">
                  <span className="w-6 h-6 rounded-full bg-indigo-600 text-white text-xs font-black flex items-center justify-center">
                    2
                  </span>
                  <div>
                    <h2 className="text-sm font-bold text-white uppercase tracking-wider">
                      Bảng Dữ Liệu Bản Vẽ ISO (Editable &bull; Filterable)
                    </h2>
                    <p className="text-[11px] text-slate-400">
                      Tự động gộp cặp file .dwg &amp; .dxf trùng tên. Bấm trực tiếp vào các ô để chỉnh sửa nếu cần.
                    </p>
                  </div>
                </div>

                {/* Search & Status Filter */}
                <div className="flex items-center gap-2">
                  <div className="relative w-48 sm:w-64">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      value={searchQuery}
                      onChange={(e) => setSearchQuery(e.target.value)}
                      placeholder="Tìm theo Zone, System, Spool..."
                      className="w-full pl-8 pr-3 py-1.5 bg-slate-950 border border-slate-700 rounded-lg text-xs text-white focus:outline-none focus:ring-1 focus:ring-purple-500"
                    />
                  </div>

                  <select
                    value={statusFilter}
                    onChange={(e) => setStatusFilter(e.target.value as 'all' | 'valid' | 'invalid')}
                    className="bg-slate-950 border border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-300 font-medium"
                  >
                    <option value="all">Tất cả ({rows.length})</option>
                    <option value="valid">Hợp lệ ({stats.valid})</option>
                    <option value="invalid">Lỗi định dạng ({stats.invalid})</option>
                  </select>
                </div>
              </div>

              {/* Quick Status Bar */}
              <div className="flex flex-wrap items-center justify-between text-xs text-slate-400 gap-2 px-1">
                <div className="flex items-center gap-3">
                  <span>
                    Tổng: <strong className="text-white">{stats.total}</strong> bản vẽ
                  </span>
                  <span>&bull;</span>
                  <span className="text-emerald-400">
                    Hợp lệ: <strong>{stats.valid}</strong>
                  </span>
                  <span>&bull;</span>
                  <span className="text-rose-400">
                    Lỗi định dạng: <strong>{stats.invalid}</strong>
                  </span>
                  <span>&bull;</span>
                  <span>
                    Có DWG: <strong className="text-cyan-400">{stats.withDwg}</strong> | Có DXF:{' '}
                    <strong className="text-purple-400">{stats.withDxf}</strong>
                  </span>
                </div>
                <span>
                  Đang hiển thị: <strong>{sortedRows.length}</strong> dòng
                </span>
              </div>

              {/* Data Table */}
              <div className="border border-slate-800 rounded-xl overflow-x-auto max-h-[460px] bg-slate-950">
                <table className="w-full text-left text-xs border-collapse">
                  <thead className="bg-slate-900 sticky top-0 z-10 border-b border-slate-800 text-[11px] font-bold text-slate-300">
                    <tr>
                      <th className="py-2.5 px-3 text-center w-12">STT</th>
                      <th className="py-2.5 px-3 min-w-[200px]">FILE NAME</th>
                      <th className="py-2.5 px-3 min-w-[180px]">PIPE NUMBER</th>
                      <th className="py-2.5 px-3 text-center min-w-[70px]">ZONE</th>
                      <th className="py-2.5 px-3 text-center min-w-[100px]">SYSTEM</th>
                      <th className="py-2.5 px-3 text-center min-w-[110px]">PIPE SPOOL</th>
                      <th className="py-2.5 px-3 text-center min-w-[50px]">REV</th>
                      <th className="py-2.5 px-3 text-center min-w-[100px]">ISSUED DATE</th>
                      <th className="py-2.5 px-3 min-w-[140px]">REMARK</th>
                      <th className="py-2.5 px-3 text-center min-w-[80px]">Status</th>
                      <th className="py-2.5 px-3 text-center min-w-[90px]">SHORT CODE</th>
                      <th className="py-2.5 px-3 text-center min-w-[90px]">CAD FORMAT</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60 font-mono text-[11px]">
                    {sortedRows.length === 0 ? (
                      <tr>
                        <td colSpan={12} className="py-12 text-center text-slate-500 font-sans">
                          {rows.length === 0
                            ? 'Chưa có bản vẽ nào. Hãy bấm "Chọn Thư Mục" ở Bước 1 để chọn thư mục chứa file .dwg, .dxf.'
                            : 'Không tìm thấy dòng nào khớp với từ khóa tìm kiếm.'}
                        </td>
                      </tr>
                    ) : (
                      sortedRows.map((r) => {
                        const isInvalid = !r.isValid;

                        return (
                          <tr
                            key={r.id}
                            className={`hover:bg-slate-900/60 transition-colors ${
                              isInvalid ? 'bg-rose-950/20 text-rose-200' : ''
                            }`}
                          >
                            <td className="py-2 px-3 text-center text-slate-500">{r.stt}</td>

                            {/* File Name */}
                            <td className="py-2 px-3 font-semibold text-slate-200 truncate max-w-[240px]" title={r.fileName}>
                              {r.fileName}
                            </td>

                            {/* Pipe Number */}
                            <td className="py-2 px-3">
                              <input
                                type="text"
                                value={r.pipeNumber}
                                onChange={(e) => handleUpdateRow(r.id, 'pipeNumber', e.target.value)}
                                className="w-full bg-transparent border-b border-transparent hover:border-slate-600 focus:border-purple-500 focus:outline-none px-1 text-slate-300"
                              />
                            </td>

                            {/* Zone */}
                            <td className="py-2 px-3 text-center font-bold text-amber-300">
                              <input
                                type="text"
                                value={r.zone}
                                onChange={(e) => handleUpdateRow(r.id, 'zone', e.target.value)}
                                className="w-14 text-center bg-transparent border-b border-transparent hover:border-slate-600 focus:border-purple-500 focus:outline-none"
                              />
                            </td>

                            {/* System */}
                            <td className="py-2 px-3 text-center text-cyan-300">
                              <input
                                type="text"
                                value={r.system}
                                onChange={(e) => handleUpdateRow(r.id, 'system', e.target.value)}
                                className="w-20 text-center bg-transparent border-b border-transparent hover:border-slate-600 focus:border-purple-500 focus:outline-none"
                              />
                            </td>

                            {/* Pipe Spool */}
                            <td className="py-2 px-3 text-center text-emerald-300">
                              <input
                                type="text"
                                value={r.pipeSpool}
                                onChange={(e) => handleUpdateRow(r.id, 'pipeSpool', e.target.value)}
                                className="w-24 text-center bg-transparent border-b border-transparent hover:border-slate-600 focus:border-purple-500 focus:outline-none"
                              />
                            </td>

                            {/* Rev */}
                            <td className="py-2 px-3 text-center font-bold text-white">
                              <input
                                type="text"
                                value={r.rev}
                                onChange={(e) => handleUpdateRow(r.id, 'rev', e.target.value)}
                                className="w-10 text-center bg-transparent border-b border-transparent hover:border-slate-600 focus:border-purple-500 focus:outline-none"
                              />
                            </td>

                            {/* Issued Date */}
                            <td className="py-2 px-3 text-center text-slate-400">
                              <input
                                type="text"
                                value={r.issuedDate}
                                onChange={(e) => handleUpdateRow(r.id, 'issuedDate', e.target.value)}
                                className="w-20 text-center bg-transparent border-b border-transparent hover:border-slate-600 focus:border-purple-500 focus:outline-none"
                              />
                            </td>

                            {/* Remark */}
                            <td className="py-2 px-3">
                              <input
                                type="text"
                                value={r.remark}
                                placeholder="Ghi chú..."
                                onChange={(e) => handleUpdateRow(r.id, 'remark', e.target.value)}
                                className="w-full bg-transparent border-b border-transparent hover:border-slate-600 focus:border-purple-500 focus:outline-none px-1 text-slate-400 text-[10px]"
                              />
                            </td>

                            {/* Status */}
                            <td className="py-2 px-3 text-center">
                              <span className="px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                                {r.status}
                              </span>
                            </td>

                            {/* Short Code */}
                            <td className="py-2 px-3 text-center font-bold text-purple-300">
                              {r.systemShort || '-'}
                            </td>

                            {/* CAD Formats */}
                            <td className="py-2 px-3 text-center">
                              <div className="flex items-center justify-center gap-1">
                                {r.hasDWG && (
                                  <span className="px-1.5 py-0.2 rounded text-[9px] font-bold bg-blue-500/20 text-blue-300 border border-blue-500/30">
                                    DWG
                                  </span>
                                )}
                                {r.hasDXF && (
                                  <span className="px-1.5 py-0.2 rounded text-[9px] font-bold bg-purple-500/20 text-purple-300 border border-purple-500/30">
                                    DXF
                                  </span>
                                )}
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </section>

            {/* STEP 3 & 4: PREVIEW FOLDER TREE, RUN & EXPORT */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              {/* Preview Folder Structure & Action Buttons */}
              <section className="lg:col-span-6 bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-sm space-y-4">
                <div className="flex items-center gap-2.5 border-b border-slate-800 pb-3">
                  <span className="w-6 h-6 rounded-full bg-cyan-600 text-white text-xs font-black flex items-center justify-center">
                    3
                  </span>
                  <h2 className="text-sm font-bold text-white uppercase tracking-wider">
                    Xem Trước Cấu Trúc Thư Mục &amp; Thực Hiện
                  </h2>
                </div>

                {/* Folder Path Summary */}
                <div className="p-3.5 rounded-xl bg-slate-950 border border-slate-800 space-y-2 text-xs">
                  <p className="font-semibold text-slate-300">Cấu trúc thư mục đầu ra:</p>
                  <p className="font-mono text-cyan-400 bg-slate-900 p-2 rounded-lg border border-slate-800 truncate">
                    {outputName || 'Output_Folder'}/S{'{ZONE}'}/{'{System_Folder_Name}'}/[file].dwg &amp; [file].dxf
                  </p>
                  <p className="text-[11px] text-slate-400">
                    * Các thư mục con trong Template chứa mã System Short Code sẽ tự động được đặt tên chuẩn kỹ thuật.
                  </p>
                </div>

                {/* Progress Bar */}
                {runProgress && (
                  <div className="p-3.5 rounded-xl bg-purple-950/30 border border-purple-800/40 space-y-2 text-xs">
                    <div className="flex justify-between font-semibold text-purple-200">
                      <span>{runProgress.statusText}</span>
                      <span>{runProgress.percentage}%</span>
                    </div>
                    <div className="w-full h-2 bg-slate-800 rounded-full overflow-hidden">
                      <div
                        className="h-full bg-gradient-to-r from-purple-500 to-cyan-400 transition-all duration-150"
                        style={{ width: `${runProgress.percentage}%` }}
                      />
                    </div>
                  </div>
                )}

                {/* Action Buttons */}
                <div className="flex flex-wrap items-center gap-3 pt-2">
                  <button
                    type="button"
                    onClick={() => handleExecuteCopy(true)}
                    disabled={isRunning || rows.length === 0}
                    className="flex-1 inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold bg-slate-800 hover:bg-slate-750 text-slate-200 border border-slate-700 transition-all cursor-pointer disabled:opacity-50"
                  >
                    <Eye className="w-4 h-4 text-cyan-400" />
                    <span>Chạy Thử (Dry-Run)</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => handleExecuteCopy(false)}
                    disabled={isRunning || rows.length === 0}
                    className="flex-1 inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold bg-gradient-to-r from-purple-600 via-indigo-600 to-cyan-500 hover:from-purple-500 hover:to-cyan-400 text-white shadow-lg shadow-purple-600/30 transition-all cursor-pointer active:scale-95 disabled:opacity-50"
                  >
                    <Play className="w-4 h-4" />
                    <span>{isRunning ? 'Đang Xử Lý...' : 'Copy Vào Thư Mục Đích'}</span>
                  </button>
                </div>

                {/* Download Zip Fallback Button */}
                <div className="pt-2">
                  <button
                    type="button"
                    onClick={handleDownloadZip}
                    disabled={isZipExporting || rows.length === 0}
                    className="w-full inline-flex items-center justify-center gap-2 px-4 py-2.5 rounded-xl text-xs font-bold bg-indigo-600/30 hover:bg-indigo-600/50 text-indigo-200 border border-indigo-500/40 transition-all cursor-pointer active:scale-95 disabled:opacity-50"
                  >
                    <Archive className="w-4 h-4 text-indigo-300" />
                    <span>{isZipExporting ? 'Đang Nén Gói ZIP...' : 'Tải Gói ZIP (Cấu Trúc S{ZONE}/{System}/)'}</span>
                  </button>
                </div>

                {/* Export Excel Button */}
                <div className="pt-2 border-t border-slate-800">
                  <button
                    type="button"
                    onClick={handleExportExcel}
                    disabled={rows.length === 0}
                    className="w-full inline-flex items-center justify-center gap-2 px-4 py-3 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-md shadow-emerald-600/30 transition-all cursor-pointer active:scale-95 disabled:opacity-50"
                  >
                    <Download className="w-4 h-4" />
                    <span>Xuất File Báo Cáo "Issued History" (.xlsx)</span>
                  </button>
                  <p className="text-[10px] text-slate-500 text-center mt-1.5">
                    Tên file: "{getIssuedHistoryFileName(projectName, sourceName)}" (chuẩn font Calibri, logo B3, màu cột A..S)
                  </p>
                </div>
              </section>

              {/* Execution Real-Time Log Panel */}
              <section className="lg:col-span-6 bg-slate-900/80 border border-slate-800 rounded-2xl p-5 shadow-sm space-y-3 flex flex-col justify-between">
                <div>
                  <div className="flex items-center justify-between border-b border-slate-800 pb-3">
                    <h2 className="text-sm font-bold text-white uppercase tracking-wider flex items-center gap-2">
                      <Clock className="w-4 h-4 text-purple-400" />
                      Nhật Ký Thực Thi (Execution Logs)
                    </h2>
                    <span className="text-[10px] font-mono text-slate-400">
                      {executionLogs.length} sự kiện
                    </span>
                  </div>

                  <div className="mt-3 p-3 bg-slate-950 border border-slate-800 rounded-xl h-64 overflow-y-auto font-mono text-[11px] space-y-1 text-slate-300">
                    {executionLogs.length === 0 ? (
                      <p className="text-slate-600 italic">
                        Chưa có hoạt động. Hãy bấm "Chạy Thử", "Copy Vào Thư Mục Đích", hoặc "Tải Gói ZIP" để theo dõi log chi tiết tại đây.
                      </p>
                    ) : (
                      executionLogs.map((logMsg, idx) => (
                        <div
                          key={idx}
                          className={`${
                            logMsg.includes('[LỖI]')
                              ? 'text-rose-400'
                              : logMsg.includes('[HOÀN THÀNH]')
                              ? 'text-emerald-400 font-bold'
                              : logMsg.includes('[THIẾU]')
                              ? 'text-amber-400'
                              : logMsg.includes('[BỎ QUA')
                              ? 'text-slate-400'
                              : 'text-cyan-300'
                          }`}
                        >
                          {logMsg}
                        </div>
                      ))
                    )}
                    <div ref={logsEndRef} />
                  </div>
                </div>

                <div className="pt-2 text-[11px] text-slate-500 flex items-center justify-between">
                  <span>* Tự động bỏ qua file đã tồn tại để tránh ghi đè dữ liệu.</span>
                  {executionLogs.length > 0 && (
                    <button
                      onClick={() => setExecutionLogs([])}
                      className="text-slate-400 hover:text-white underline cursor-pointer text-xs"
                    >
                      Xóa logs
                    </button>
                  )}
                </div>
              </section>
            </div>
          </div>
        )}

        {/* TAB 2: ISO RELEASED TOOL */}
        {activeTab === 'released' && (
          <section className="bg-slate-900/80 border border-slate-800 rounded-2xl p-6 shadow-sm space-y-6 max-w-4xl mx-auto">
            <div className="border-b border-slate-800 pb-4">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-purple-600/20 text-purple-400 flex items-center justify-center border border-purple-500/30">
                  <Archive className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-base font-bold text-white">
                    Công Cụ ISO Released (Gộp Bản Vẽ Từ Các Thư Mục Con)
                  </h2>
                  <p className="text-xs text-slate-400">
                    Tạo thư mục "Iso released yyyymmdd" trong thư mục đầu ra và copy toàn bộ file từ các thư mục con được chọn.
                  </p>
                </div>
              </div>
            </div>

            {/* Select Parent Folder */}
            <div className="p-4 rounded-xl bg-slate-950 border border-slate-800 space-y-3">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <span className="text-xs font-bold text-slate-200">1. Chọn Thư mục mẹ chứa các thư mục con:</span>
                  <p className="text-xs font-mono text-purple-400 mt-1">
                    {relParentName || 'Chưa chọn thư mục mẹ'}
                  </p>
                </div>
                <button
                  type="button"
                  onClick={handleSelectRelParent}
                  className="px-4 py-2 rounded-xl text-xs font-bold bg-purple-600 hover:bg-purple-500 text-white transition-colors cursor-pointer"
                >
                  Chọn Thư Mục Mẹ
                </button>
              </div>
            </div>

            {/* Subfolders Checklist */}
            {relSubfolders.length > 0 && (
              <div className="space-y-3">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-bold text-slate-300">
                    2. Chọn các thư mục con cần copy ({relSubfolders.filter((f) => f.checked).length}/{relSubfolders.length}):
                  </span>
                  <div className="flex gap-2">
                    <button
                      type="button"
                      onClick={() => handleToggleAllRel(true)}
                      className="text-purple-400 hover:underline cursor-pointer"
                    >
                      Chọn tất cả
                    </button>
                    <span className="text-slate-600">|</span>
                    <button
                      type="button"
                      onClick={() => handleToggleAllRel(false)}
                      className="text-slate-400 hover:underline cursor-pointer"
                    >
                      Bỏ chọn tất cả
                    </button>
                  </div>
                </div>

                <div className="p-3 bg-slate-950 border border-slate-800 rounded-xl max-h-56 overflow-y-auto grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs">
                  {relSubfolders.map((f, idx) => (
                    <label
                      key={idx}
                      className="flex items-center gap-2 p-2 rounded-lg hover:bg-slate-900 cursor-pointer border border-transparent hover:border-slate-800"
                    >
                      <input
                        type="checkbox"
                        checked={f.checked}
                        onChange={(e) => {
                          const checked = e.target.checked;
                          setRelSubfolders((prev) =>
                            prev.map((item, i) => (i === idx ? { ...item, checked } : item)),
                          );
                        }}
                        className="rounded border-slate-700 text-purple-600 focus:ring-purple-500"
                      />
                      <span className="truncate text-slate-300 font-mono text-[11px]">{f.name}</span>
                    </label>
                  ))}
                </div>

                {/* Warning on Overwrite */}
                <div className="p-3 rounded-xl bg-amber-500/10 border border-amber-500/20 text-amber-300 text-xs flex items-center gap-2">
                  <AlertTriangle className="w-4 h-4 shrink-0" />
                  <span>
                    Lưu ý: Nếu trong thư mục đích đã có file cùng tên, hệ thống sẽ ghi đè và ghi nhận vào log.
                  </span>
                </div>

                {/* Execute Button */}
                <button
                  type="button"
                  onClick={handleExecuteIsoReleased}
                  disabled={relIsProcessing}
                  className="w-full py-3 rounded-xl text-xs font-bold bg-gradient-to-r from-purple-600 to-indigo-600 hover:from-purple-500 hover:to-indigo-500 text-white shadow-lg shadow-purple-600/30 transition-all cursor-pointer disabled:opacity-50"
                >
                  {relIsProcessing
                    ? 'Đang Sao Chép Vào Iso Released...'
                    : 'Thực Hiện Gom File Vào Thư Mục "Iso released yyyymmdd"'}
                </button>
              </div>
            )}

            {/* Logs */}
            {relLogs.length > 0 && (
              <div className="p-3 bg-slate-950 border border-slate-800 rounded-xl h-44 overflow-y-auto font-mono text-[11px] space-y-1 text-slate-300">
                {relLogs.map((m, i) => (
                  <div key={i}>{m}</div>
                ))}
              </div>
            )}
          </section>
        )}

        {/* TAB 3: UPDATE OLD REV TOOL */}
        {activeTab === 'oldrev' && (
          <section className="bg-slate-900/80 border border-slate-800 rounded-2xl p-6 shadow-sm space-y-6 max-w-4xl mx-auto">
            <div className="border-b border-slate-800 pb-4">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-cyan-600/20 text-cyan-400 flex items-center justify-center border border-cyan-500/30">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-base font-bold text-white">
                    Công Cụ Update Old Rev (Đánh Dấu Bản Vẽ Bản Cũ)
                  </h2>
                  <p className="text-xs text-slate-400">
                    Tìm Revision cao nhất cho từng nhóm bản vẽ (theo Group Key), tự động điền chữ "old rev" vào các dòng bản vẽ cũ hơn.
                  </p>
                </div>
              </div>
            </div>

            {/* Upload Excel File */}
            <div className="p-4 rounded-xl bg-slate-950 border border-slate-800 space-y-3">
              <span className="text-xs font-bold text-slate-200 block">1. Tải lên file Excel (.xlsx) cần xử lý:</span>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={handleUploadRevExcel}
                className="block w-full text-xs text-slate-400 file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-xs file:font-semibold file:bg-purple-600 file:text-white hover:file:bg-purple-500 cursor-pointer"
              />
            </div>

            {/* Configuration for Matrix */}
            {revRawMatrix && (
              <div className="space-y-4">
                <div className="p-4 rounded-xl bg-slate-950 border border-slate-800 space-y-3">
                  <span className="text-xs font-bold text-slate-200 block">2. Cấu hình các cột (Theo mẫu VBA chuẩn):</span>

                  <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 text-xs">
                    {/* Sheet selector */}
                    <div>
                      <label className="block text-slate-400 mb-1">Sheet làm việc:</label>
                      <select
                        value={revActiveSheet}
                        onChange={(e) => setRevActiveSheet(e.target.value)}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg p-1.5 text-xs text-white"
                      >
                        {revSheetNames.map((s) => (
                          <option key={s} value={s}>
                            {s}
                          </option>
                        ))}
                      </select>
                    </div>

                    {/* Group Key Col (L = index 11) */}
                    <div>
                      <label className="block text-slate-400 mb-1">Cột Group Key (Mặc định L):</label>
                      <input
                        type="number"
                        min={0}
                        value={revKeyCol + 1}
                        onChange={(e) => setRevKeyCol(Math.max(0, Number(e.target.value) - 1))}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg p-1.5 text-xs text-white"
                      />
                      <span className="text-[10px] text-slate-500">Cột số {revKeyCol + 1}</span>
                    </div>

                    {/* Revision Col (P = index 15) */}
                    <div>
                      <label className="block text-slate-400 mb-1">Cột Revision (Mặc định P):</label>
                      <input
                        type="number"
                        min={0}
                        value={revRevCol + 1}
                        onChange={(e) => setRevRevCol(Math.max(0, Number(e.target.value) - 1))}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg p-1.5 text-xs text-white"
                      />
                      <span className="text-[10px] text-slate-500">Cột số {revRevCol + 1}</span>
                    </div>

                    {/* Mark Col (S = index 18) */}
                    <div>
                      <label className="block text-slate-400 mb-1">Cột Đánh dấu (Mặc định S):</label>
                      <input
                        type="number"
                        min={0}
                        value={revMarkCol + 1}
                        onChange={(e) => setRevMarkCol(Math.max(0, Number(e.target.value) - 1))}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg p-1.5 text-xs text-white"
                      />
                      <span className="text-[10px] text-slate-500">Cột số {revMarkCol + 1}</span>
                    </div>
                  </div>
                </div>

                {/* Actions */}
                <div className="flex flex-col sm:flex-row items-center gap-3">
                  <button
                    type="button"
                    onClick={handleExecuteMarkOldRev}
                    className="flex-1 py-3 rounded-xl text-xs font-bold bg-cyan-600 hover:bg-cyan-500 text-white shadow-md transition-colors cursor-pointer"
                  >
                    Xử Lý &amp; Đánh Dấu "old rev"
                  </button>

                  {revPreviewCount !== null && (
                    <button
                      type="button"
                      onClick={handleDownloadMarkedExcel}
                      className="flex-1 py-3 rounded-xl text-xs font-bold bg-emerald-600 hover:bg-emerald-500 text-white shadow-md transition-colors cursor-pointer flex items-center justify-center gap-2"
                    >
                      <Download className="w-4 h-4" />
                      <span>Tải Về File Đã Cập Nhật ({revPreviewCount} dòng đã đánh dấu)</span>
                    </button>
                  )}
                </div>
              </div>
            )}
          </section>
        )}
      </main>

      {/* CONFIRM RESET MODAL */}
      {isResetConfirmOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs animate-in fade-in duration-150">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-md w-full p-6 space-y-4 shadow-2xl">
            <div className="flex items-start gap-3">
              <div className="w-10 h-10 rounded-xl bg-rose-500/20 text-rose-400 flex items-center justify-center shrink-0 border border-rose-500/30">
                <Trash2 className="w-5 h-5" />
              </div>
              <div className="space-y-1">
                <h3 className="text-sm font-bold text-white">Xác nhận làm mới toàn bộ bảng dữ liệu?</h3>
                <p className="text-xs text-slate-400 leading-relaxed">
                  Tất cả các bản vẽ đang hiển thị trong bảng dữ liệu sẽ bị xóa. Cấu hình thư mục và tên dự án vẫn được giữ nguyên.
                </p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setIsResetConfirmOpen(false)}
                className="px-4 py-2 rounded-xl text-xs font-semibold text-slate-400 hover:bg-slate-800 transition-colors cursor-pointer"
              >
                Hủy bỏ
              </button>
              <button
                type="button"
                onClick={handleConfirmReset}
                className="px-4 py-2 rounded-xl text-xs font-bold bg-rose-600 hover:bg-rose-500 text-white transition-colors cursor-pointer"
              >
                Đồng ý xóa sạch
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Footer */}
      <footer className="border-t border-slate-800/80 bg-slate-950 py-4 text-xs text-slate-500 text-center">
        <div className="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row items-center justify-between gap-2">
          <span>&copy; 2026 WebToolRush &bull; ISO Drawing Manager Suite.</span>
          <span className="font-mono text-[11px] text-slate-400">100% Client-Side File Processing</span>
        </div>
      </footer>
    </div>
  );
};
