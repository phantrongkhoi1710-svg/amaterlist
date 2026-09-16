import React, { useState, useRef } from 'react';
import {
  Upload,
  FileCheck,
  AlertTriangle,
  XCircle,
  FileSpreadsheet,
  Trash2,
  Play,
  RotateCcw,
  CheckCircle2,
  ChevronDown,
  ChevronUp,
  Layers,
} from 'lucide-react';
import { SourceFileInfo, ImportMode } from '../types';

interface FileUploaderProps {
  files: SourceFileInfo[];
  onUploadFiles: (files: File[], mode: ImportMode) => Promise<void>;
  onRemoveFile: (fileId: string) => void;
  onClearAllFiles: () => void;
  isProcessing: boolean;
  importMode: ImportMode;
  setImportMode: (mode: ImportMode) => void;
}

export const FileUploader: React.FC<FileUploaderProps> = ({
  files,
  onUploadFiles,
  onRemoveFile,
  onClearAllFiles,
  isProcessing,
  importMode,
  setImportMode,
}) => {
  const [isDragOver, setIsDragOver] = useState(false);
  const [expandedFileId, setExpandedFileId] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const selectedFiles = Array.from(e.dataTransfer.files);
      await onUploadFiles(selectedFiles, importMode);
    }
  };

  const handleFileInputChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const selectedFiles = Array.from(e.target.files);
      await onUploadFiles(selectedFiles, importMode);
      e.target.value = '';
    }
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
  };

  const toggleExpand = (id: string) => {
    setExpandedFileId((prev) => (prev === id ? null : id));
  };

  return (
    <div className="space-y-6">
      {/* Import Mode Selector */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 p-3.5 shadow-sm">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <h3 className="text-xs font-semibold text-slate-800 dark:text-slate-100 flex items-center gap-2">
            <Layers className="w-4 h-4 text-blue-500" />
            Chế độ Import dữ liệu
          </h3>

          <div className="inline-flex rounded-lg p-1 bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700">
            <button
              id="mode-replace-btn"
              type="button"
              onClick={() => setImportMode('replace')}
              className={`px-3 py-1 text-xs font-medium rounded-md transition-colors ${
                importMode === 'replace'
                  ? 'bg-blue-600 text-white shadow-sm'
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
              }`}
            >
              Làm mới hoàn toàn (Replace)
            </button>
            <button
              id="mode-append-btn"
              type="button"
              onClick={() => setImportMode('append')}
              className={`px-3 py-1 text-xs font-medium rounded-md transition-colors ${
                importMode === 'append'
                  ? 'bg-blue-600 text-white shadow-sm'
                  : 'text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-white'
              }`}
            >
              Ghi tiếp nối (Append)
            </button>
          </div>
        </div>
      </div>

      {/* Drag and Drop Zone */}
      <div
        id="drop-zone"
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
        className={`relative border-2 border-dashed rounded-2xl p-8 text-center cursor-pointer transition-all duration-200 ${
          isDragOver
            ? 'border-blue-500 bg-blue-50/60 dark:bg-blue-950/20 scale-[1.01]'
            : 'border-slate-300 dark:border-slate-700 hover:border-blue-400 dark:hover:border-blue-500 bg-slate-50/50 dark:bg-slate-900/40'
        }`}
      >
        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept=".xlsx,.xls,.xlsm,.xlsb,.csv"
          onChange={handleFileInputChange}
          className="hidden"
          id="file-input"
        />

        <div className="flex flex-col items-center justify-center space-y-3">
          <div className="w-14 h-14 rounded-full bg-blue-100 dark:bg-blue-900/40 text-blue-600 dark:text-blue-400 flex items-center justify-center shadow-inner">
            <Upload className="w-7 h-7" />
          </div>
          <div>
            <p className="text-base font-medium text-slate-800 dark:text-slate-200">
              Kéo thả các file Excel Armature vào đây, hoặc{' '}
              <span className="text-blue-600 dark:text-blue-400 font-semibold hover:underline">
                duyệt từ máy tính
              </span>
            </p>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
              Định dạng: <strong>.xlsx, .xls, .xlsm, .xlsb, .csv</strong> (Hỗ trợ chọn nhiều file)
            </p>
          </div>
        </div>

        {isProcessing && (
          <div className="absolute inset-0 bg-white/80 dark:bg-slate-900/80 backdrop-blur-xs rounded-2xl flex flex-col items-center justify-center gap-2 z-10">
            <div className="w-8 h-8 border-3 border-blue-600 border-t-transparent rounded-full animate-spin"></div>
            <p className="text-xs font-semibold text-slate-700 dark:text-slate-200">
              Đang phân tích và nhập dữ liệu từ file...
            </p>
          </div>
        )}
      </div>

      {/* Uploaded / Processed Files List */}
      {files.length > 0 && (
        <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
          <div className="px-5 py-3.5 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <FileSpreadsheet className="w-4 h-4 text-blue-500" />
              <h3 className="text-sm font-semibold text-slate-800 dark:text-slate-200">
                Danh sách File Nguồn ({files.length} file)
              </h3>
            </div>
            <button
              id="btn-clear-files"
              onClick={onClearAllFiles}
              className="text-xs font-medium text-rose-500 hover:text-rose-600 dark:hover:text-rose-400 flex items-center gap-1 transition-colors"
            >
              <Trash2 className="w-3.5 h-3.5" />
              Xoá danh sách file
            </button>
          </div>

          <div className="divide-y divide-slate-100 dark:divide-slate-800">
            {files.map((f) => (
              <div key={f.id} className="p-4 transition-colors hover:bg-slate-50/60 dark:hover:bg-slate-800/40">
                <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                  <div className="flex items-start gap-3">
                    <div className="w-9 h-9 rounded-lg bg-emerald-50 dark:bg-emerald-950/40 text-emerald-600 dark:text-emerald-400 flex items-center justify-center flex-shrink-0 mt-0.5">
                      <FileSpreadsheet className="w-5 h-5" />
                    </div>
                    <div>
                      <div className="flex flex-wrap items-center gap-2">
                        <span className="text-sm font-medium text-slate-900 dark:text-slate-100">
                          {f.name}
                        </span>
                        <span className="text-xs text-slate-400">
                          ({formatFileSize(f.size)})
                        </span>

                        {/* Status Badge */}
                        {f.status === 'success' && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-emerald-100 text-emerald-700 dark:bg-emerald-900/30 dark:text-emerald-300">
                            <CheckCircle2 className="w-3 h-3" /> Thành công
                          </span>
                        )}
                        {f.status === 'warning' && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-amber-100 text-amber-700 dark:bg-amber-900/30 dark:text-amber-300">
                            <AlertTriangle className="w-3 h-3" /> Thiếu một số trường
                          </span>
                        )}
                        {f.status === 'error' && (
                          <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium bg-rose-100 text-rose-700 dark:bg-rose-900/30 dark:text-rose-300">
                            <XCircle className="w-3 h-3" /> Lỗi
                          </span>
                        )}
                      </div>

                      {/* File Details Line */}
                      <div className="flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-500 dark:text-slate-400 mt-1">
                        {f.detectedSheet && (
                          <span>
                            Sheet: <strong>{f.detectedSheet}</strong>
                          </span>
                        )}
                        {f.headerRow && (
                          <span>
                            Dòng tiêu đề: <strong>Dòng {f.headerRow}</strong>
                          </span>
                        )}
                        {f.importedRows !== undefined && (
                          <span>
                            Nhập: <strong className="text-emerald-600 dark:text-emerald-400">{f.importedRows}</strong> hàng
                          </span>
                        )}
                        {f.skippedRows !== undefined && f.skippedRows > 0 && (
                          <span>
                            Bỏ qua: <span className="text-slate-400">{f.skippedRows} dòng trống</span>
                          </span>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Actions */}
                  <div className="flex items-center gap-2 self-end md:self-center">
                    <button
                      type="button"
                      onClick={() => toggleExpand(f.id)}
                      className="px-2.5 py-1 text-xs rounded-md bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 flex items-center gap-1"
                    >
                      {expandedFileId === f.id ? (
                        <>
                          <span>Thu gọn</span>
                          <ChevronUp className="w-3.5 h-3.5" />
                        </>
                      ) : (
                        <>
                          <span>Chi tiết cột</span>
                          <ChevronDown className="w-3.5 h-3.5" />
                        </>
                      )}
                    </button>
                    <button
                      type="button"
                      onClick={() => onRemoveFile(f.id)}
                      className="p-1 text-slate-400 hover:text-rose-500 transition-colors"
                      title="Xoá file này"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {/* Expandable Column Mapping Detail */}
                {expandedFileId === f.id && (
                  <div className="mt-3 pt-3 border-t border-slate-100 dark:border-slate-800/60 bg-slate-50/50 dark:bg-slate-950/40 rounded-lg p-3 text-xs">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                      <div>
                        <h4 className="font-semibold text-slate-700 dark:text-slate-300 mb-1.5 flex items-center gap-1.5">
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500" />
                          Các cột đã ánh xạ ({f.matchedColumns?.length || 0}):
                        </h4>
                        <div className="flex flex-wrap gap-1.5 max-h-36 overflow-y-auto">
                          {f.matchedColumns && f.matchedColumns.length > 0 ? (
                            f.matchedColumns.map((mc, idx) => (
                              <span
                                key={idx}
                                className="px-2 py-0.5 rounded bg-emerald-50 text-emerald-800 dark:bg-emerald-950/50 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 text-[11px]"
                              >
                                <strong>{mc.master}</strong> &larr; {mc.source} (Cột {mc.colIndex + 1})
                              </span>
                            ))
                          ) : (
                            <span className="text-slate-400 italic">Không có cột nào</span>
                          )}
                        </div>
                      </div>

                      <div>
                        <h4 className="font-semibold text-slate-700 dark:text-slate-300 mb-1.5 flex items-center gap-1.5">
                          <AlertTriangle className="w-3.5 h-3.5 text-amber-500" />
                          Trường Master không có trong file này ({f.missingFields?.length || 0}):
                        </h4>
                        <div className="flex flex-wrap gap-1.5 max-h-36 overflow-y-auto">
                          {f.missingFields && f.missingFields.length > 0 ? (
                            f.missingFields.map((field, idx) => (
                              <span
                                key={idx}
                                className="px-2 py-0.5 rounded bg-amber-50 text-amber-800 dark:bg-amber-950/50 dark:text-amber-300 border border-amber-200 dark:border-amber-800 text-[11px]"
                              >
                                {field}
                              </span>
                            ))
                          ) : (
                            <span className="text-emerald-600 dark:text-emerald-400 font-medium">
                              Đầy đủ tất cả trường master!
                            </span>
                          )}
                        </div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};
