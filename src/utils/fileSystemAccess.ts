/**
 * Utilities for File System Access API with IndexedDB handle persistence
 * and web-safe fallback support (JSZip).
 */

import JSZip from 'jszip';
import { IsoDrawingRow } from './isoDrawingLogic';

export interface ScannedFileMapItem {
  baseName: string;
  dwgHandle?: FileSystemFileHandle;
  dxfHandle?: FileSystemFileHandle;
  dwgFile?: File;
  dxfFile?: File;
}

export interface CopyProgress {
  total: number;
  current: number;
  percentage: number;
  currentFileName: string;
  statusText: string;
}

export interface CopySummary {
  foldersCreated: number;
  filesCopied: number;
  filesSkipped: number;
  filesMissing: number;
  logs: string[];
}

/**
 * Check if the browser supports File System Access API
 */
export function isFileSystemAccessSupported(): boolean {
  if (typeof window === 'undefined') return false;
  // If running inside a cross-origin iframe, showDirectoryPicker is forbidden by browser security policy
  try {
    if (window.self !== window.top) {
      return false; // Safely default to webkitdirectory inside iframe
    }
  } catch {
    return false;
  }
  return 'showDirectoryPicker' in window;
}

export function isInIframe(): boolean {
  try {
    return typeof window !== 'undefined' && window.self !== window.top;
  } catch {
    return true;
  }
}

/**
 * Scan a FileList or File[] (from <input webkitdirectory>) for .dwg and .dxf files
 */
export function scanCadFileList(files: FileList | File[]): Map<string, ScannedFileMapItem> {
  const map = new Map<string, ScannedFileMapItem>();
  const fileArray = Array.from(files);

  for (const file of fileArray) {
    const lowerName = file.name.toLowerCase();
    const isDWG = lowerName.endsWith('.dwg');
    const isDXF = lowerName.endsWith('.dxf');
    if (!isDWG && !isDXF) continue;

    const ext = isDWG ? '.dwg' : '.dxf';
    const baseName = file.name.slice(0, -ext.length);
    const lowerKey = baseName.toLowerCase();

    let entry = map.get(lowerKey);
    if (!entry) {
      entry = { baseName };
      map.set(lowerKey, entry);
    }

    if (isDWG) {
      entry.dwgFile = file;
    } else {
      entry.dxfFile = file;
    }
  }

  return map;
}

/**
 * Extract template subfolder names from files chosen via <input webkitdirectory>
 */
export function extractTemplateSubfoldersFromFiles(files: FileList | File[]): string[] {
  const set = new Set<string>();
  const fileArray = Array.from(files);

  for (const file of fileArray) {
    const rel = (file as unknown as { webkitRelativePath?: string }).webkitRelativePath || '';
    if (rel) {
      const parts = rel.split('/');
      if (parts.length > 2 && parts[1]) {
        set.add(parts[1]);
      }
    }
  }

  return Array.from(set);
}

// Simple Native IndexedDB Key-Value Helper for Directory Handles
const DB_NAME = 'iso_drawing_manager_db';
const STORE_NAME = 'handles_store';

function openIDB(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

export async function saveHandleToIDB(key: string, val: unknown): Promise<void> {
  try {
    const db = await openIDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      tx.objectStore(STORE_NAME).put(val, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn('Could not save handle to IndexedDB:', err);
  }
}

export async function getHandleFromIDB<T>(key: string): Promise<T | null> {
  try {
    const db = await openIDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const req = tx.objectStore(STORE_NAME).get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    console.warn('Could not read handle from IndexedDB:', err);
    return null;
  }
}

/**
 * Prompt user to pick a folder using showDirectoryPicker
 */
export async function pickDirectory(idHint?: string): Promise<FileSystemDirectoryHandle | null> {
  if (!isFileSystemAccessSupported()) {
    throw new Error('Trình duyệt của bạn không hỗ trợ File System Access API. Vui lòng sử dụng Chrome hoặc Edge trên Desktop.');
  }
  try {
    const handle = await (window as unknown as { showDirectoryPicker: (opts?: unknown) => Promise<FileSystemDirectoryHandle> }).showDirectoryPicker({
      id: idHint,
      mode: 'readwrite',
    });
    return handle;
  } catch (err) {
    if ((err as Error).name === 'AbortError') {
      return null;
    }
    throw err;
  }
}

/**
 * Scan a directory for all .dwg and .dxf files (non-recursive).
 * Deduplicates by base name (case-insensitive).
 */
export async function scanCadDirectory(
  dirHandle: FileSystemDirectoryHandle,
  onProgress?: (count: number) => void,
): Promise<Map<string, ScannedFileMapItem>> {
  const map = new Map<string, ScannedFileMapItem>();
  let scannedCount = 0;

  for await (const [name, handle] of (dirHandle as unknown as AsyncIterable<[string, FileSystemHandle]>)) {
    if (handle.kind !== 'file') continue;

    const lowerName = name.toLowerCase();
    const isDWG = lowerName.endsWith('.dwg');
    const isDXF = lowerName.endsWith('.dxf');

    if (!isDWG && !isDXF) continue;

    const ext = isDWG ? '.dwg' : '.dxf';
    const baseName = name.slice(0, -ext.length);
    const lowerKey = baseName.toLowerCase();

    let entry = map.get(lowerKey);
    if (!entry) {
      entry = { baseName };
      map.set(lowerKey, entry);
    }

    if (isDWG) {
      entry.dwgHandle = handle as FileSystemFileHandle;
    } else {
      entry.dxfHandle = handle as FileSystemFileHandle;
    }

    scannedCount++;
    if (scannedCount % 50 === 0 && onProgress) {
      onProgress(scannedCount);
      // Yield to UI loop
      await new Promise((r) => setTimeout(r, 0));
    }
  }

  if (onProgress) onProgress(scannedCount);
  return map;
}

/**
 * List subfolder names in a template folder handle
 */
export async function listSubfolderNames(dirHandle: FileSystemDirectoryHandle): Promise<string[]> {
  const subfolders: string[] = [];
  for await (const [name, handle] of (dirHandle as unknown as AsyncIterable<[string, FileSystemHandle]>)) {
    if (handle.kind === 'directory') {
      subfolders.push(name);
    }
  }
  return subfolders;
}

/**
 * Copy CAD files into destination hierarchy S{ZONE}/{SystemFolder}/
 * large-folder safe with progress callback.
 */
export async function copyCadFilesToOutput(
  rows: IsoDrawingRow[],
  fileMap: Map<string, ScannedFileMapItem>,
  outputDirHandle: FileSystemDirectoryHandle,
  templateSubfolders: string[],
  isDryRun: boolean = false,
  onProgress?: (progress: CopyProgress) => void,
  onLog?: (msg: string) => void,
): Promise<CopySummary> {
  const summary: CopySummary = {
    foldersCreated: 0,
    filesCopied: 0,
    filesSkipped: 0,
    filesMissing: 0,
    logs: [],
  };

  const log = (msg: string) => {
    summary.logs.push(msg);
    if (onLog) onLog(msg);
  };

  const total = rows.length;
  const createdZoneFolders = new Set<string>();
  const createdSysFolders = new Set<string>();

  log(`[START] Bắt đầu xử lý ${total} bản vẽ (Chế độ: ${isDryRun ? 'DRY-RUN / PREVIEW ONLY' : 'CHÍNH THỨC'})`);

  for (let i = 0; i < total; i++) {
    const row = rows[i];
    const baseKey = row.fileName.toLowerCase();
    const fileEntry = fileMap.get(baseKey);

    const percent = Math.round(((i + 1) / total) * 100);
    if (onProgress) {
      onProgress({
        total,
        current: i + 1,
        percentage: percent,
        currentFileName: row.fileName,
        statusText: `Đang xử lý ${i + 1}/${total}: ${row.fileName}`,
      });
    }

    if (!row.isValid) {
      log(`[BỎ QUA] STT ${row.stt} (${row.fileName}): Lỗi định dạng tên.`);
      summary.filesMissing++;
      continue;
    }

    const zoneFolderName = `S${row.zone}`;
    const sysFolderName = row.targetFolder || row.systemShort || 'MISC';

    if (!fileEntry || (!fileEntry.dwgHandle && !fileEntry.dxfHandle)) {
      log(`[THIẾU] STT ${row.stt} (${row.fileName}): Không tìm thấy file CAD gốc trong thư mục nguồn.`);
      summary.filesMissing++;
      continue;
    }

    if (isDryRun) {
      log(`[DRY-RUN] Sẽ sao chép: ${row.fileName} -> ${zoneFolderName}/${sysFolderName}/`);
      summary.filesCopied++;
      continue;
    }

    try {
      // 1. Get or create S{ZONE}
      const zoneHandle = await outputDirHandle.getDirectoryHandle(zoneFolderName, { create: true });
      if (!createdZoneFolders.has(zoneFolderName)) {
        createdZoneFolders.add(zoneFolderName);
        summary.foldersCreated++;
      }

      // 2. Get or create {SystemFolder}
      const sysHandle = await zoneHandle.getDirectoryHandle(sysFolderName, { create: true });
      const sysKey = `${zoneFolderName}/${sysFolderName}`;
      if (!createdSysFolders.has(sysKey)) {
        createdSysFolders.add(sysKey);
        summary.foldersCreated++;
      }

      // 3. Copy .dwg if present
      if (fileEntry.dwgHandle) {
        const destName = `${row.fileName}.dwg`;
        let alreadyExists = false;
        try {
          await sysHandle.getFileHandle(destName, { create: false });
          alreadyExists = true;
        } catch {
          alreadyExists = false;
        }

        if (alreadyExists) {
          log(`[BỎ QUA TỒN TẠI] ${destName} đã tồn tại trong ${sysKey}. Không ghi đè.`);
          summary.filesSkipped++;
        } else {
          const srcFile = await fileEntry.dwgHandle.getFile();
          const destHandle = await sysHandle.getFileHandle(destName, { create: true });
          const writable = await (destHandle as unknown as { createWritable: () => Promise<FileSystemWritableFileStream> }).createWritable();
          await writable.write(srcFile);
          await writable.close();
          summary.filesCopied++;
        }
      }

      // 4. Copy .dxf if present
      if (fileEntry.dxfHandle) {
        const destName = `${row.fileName}.dxf`;
        let alreadyExists = false;
        try {
          await sysHandle.getFileHandle(destName, { create: false });
          alreadyExists = true;
        } catch {
          alreadyExists = false;
        }

        if (alreadyExists) {
          log(`[BỎ QUA TỒN TẠI] ${destName} đã tồn tại trong ${sysKey}. Không ghi đè.`);
          summary.filesSkipped++;
        } else {
          const srcFile = await fileEntry.dxfHandle.getFile();
          const destHandle = await sysHandle.getFileHandle(destName, { create: true });
          const writable = await (destHandle as unknown as { createWritable: () => Promise<FileSystemWritableFileStream> }).createWritable();
          await writable.write(srcFile);
          await writable.close();
          summary.filesCopied++;
        }
      }
    } catch (err) {
      log(`[LỖI] STT ${row.stt} (${row.fileName}): ${(err as Error).message}`);
    }

    // Yield every 20 files to keep UI responsive
    if (i % 20 === 0) {
      await new Promise((r) => setTimeout(r, 0));
    }
  }

  log(`[HOÀN THÀNH] Đã tạo: ${summary.foldersCreated} thư mục | Đã sao chép: ${summary.filesCopied} files | Bỏ qua: ${summary.filesSkipped} | Thiếu: ${summary.filesMissing}`);
  return summary;
}

/**
 * Fallback ZIP exporter for browsers without File System Access API
 */
export async function createZipFallback(
  rows: IsoDrawingRow[],
  fileMap: Map<string, ScannedFileMapItem>,
  onProgress?: (percent: number) => void,
): Promise<Blob> {
  const zip = new JSZip();
  const total = rows.length;

  for (let i = 0; i < total; i++) {
    const row = rows[i];
    if (!row.isValid) continue;

    const baseKey = row.fileName.toLowerCase();
    const fileEntry = fileMap.get(baseKey);
    if (!fileEntry) continue;

    const zoneFolderName = `S${row.zone}`;
    const sysFolderName = row.targetFolder || row.systemShort || 'MISC';
    const folderPath = `${zoneFolderName}/${sysFolderName}`;

    if (fileEntry.dwgHandle || fileEntry.dwgFile) {
      const file = fileEntry.dwgFile || (await fileEntry.dwgHandle!.getFile());
      zip.folder(folderPath)?.file(`${row.fileName}.dwg`, file);
    }
    if (fileEntry.dxfHandle || fileEntry.dxfFile) {
      const file = fileEntry.dxfFile || (await fileEntry.dxfHandle!.getFile());
      zip.folder(folderPath)?.file(`${row.fileName}.dxf`, file);
    }

    if (onProgress && i % 10 === 0) {
      onProgress(Math.round(((i + 1) / total) * 100));
      await new Promise((r) => setTimeout(r, 0));
    }
  }

  return await zip.generateAsync({ type: 'blob' }, (metadata) => {
    if (onProgress) onProgress(metadata.percent);
  });
}
