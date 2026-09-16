import React, { useState, useMemo } from 'react';
import {
  Search,
  ArrowUpDown,
  ArrowUp,
  ArrowDown,
  Download,
  Filter,
  FileSpreadsheet,
  Trash2,
  ChevronLeft,
  ChevronRight,
  Sparkles,
  Edit3,
  Mail,
  PlusCircle,
  FileUp,
  Clock,
  SlidersHorizontal,
} from 'lucide-react';
import { MasterRowData } from '../types';

interface MasterTableProps {
  data: MasterRowData[];
  masterHeaders: string[];
  onDeleteRow: (rowId: string) => void;
  onEditRow: (row: MasterRowData) => void;
  onUpdateRow?: (updatedRow: MasterRowData) => void;
  onAddRow?: (newRow?: MasterRowData) => void;
  onLoadImageSample?: () => void;
  onMailRow: (row: MasterRowData) => void;
  onOpenBatchMail: () => void;
  onExportExcel: () => void;
  onExportCSV: () => void;
  onToggleFileUploader?: () => void;
  isFileUploaderOpen?: boolean;
  totalFilesCount?: number;
  onOpenAliases?: () => void;
  onOpenLogs?: () => void;
  logsCount?: number;
}

export const MasterTable: React.FC<MasterTableProps> = ({
  data,
  masterHeaders,
  onDeleteRow,
  onEditRow,
  onUpdateRow,
  onAddRow,
  onLoadImageSample,
  onMailRow,
  onOpenBatchMail,
  onExportExcel,
  onExportCSV,
  onToggleFileUploader,
  isFileUploaderOpen = false,
  totalFilesCount = 0,
  onOpenAliases,
  onOpenLogs,
  logsCount = 0,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedFileFilter, setSelectedFileFilter] = useState<string>('all');
  const [sortColumn, setSortColumn] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);

  // Editing cell state (for direct inline cell changes)
  const [editingCell, setEditingCell] = useState<{ rowId: string; header: string } | null>(null);
  const [editingValue, setEditingValue] = useState<string>('');

  // Determine split index for two groups:
  // Group 1: TO BE COMPLETED BY FUNCTIONAL DESIGN (Columns: TAG, ACTUATOR TAG, P.O. NUMBER, SUPPLIER, DESTINATION (YARD))
  // Group 2: ARMATURE INFO FROM PIPE SPECIFICATION (STD DRW NORMALE N° and onwards)
  const functionalDesignColsCount = useMemo(() => {
    const idx = masterHeaders.findIndex(
      (h) =>
        h.toUpperCase().includes('STD DRW') ||
        h.toUpperCase().includes('NORMALE') ||
        h.toUpperCase() === 'EXECUTION'
    );
    return idx !== -1 ? idx : 5;
  }, [masterHeaders]);

  // Get unique source files
  const sourceFiles = useMemo(() => {
    const files = new Set<string>();
    data.forEach((r) => {
      if (r._sourceFile) files.add(r._sourceFile);
    });
    return Array.from(files);
  }, [data]);

  // Handle inline commit
  const handleCommitInline = (row: MasterRowData, header: string, newVal: string) => {
    setEditingCell(null);
    if (row[header] === newVal) return;
    if (onUpdateRow) {
      const updated: MasterRowData = {
        ...row,
        [header]: newVal,
      };
      onUpdateRow(updated);
    }
  };

  // Filter data
  const filteredData = useMemo(() => {
    return data.filter((row) => {
      if (selectedFileFilter !== 'all' && row._sourceFile !== selectedFileFilter) {
        return false;
      }

      if (searchTerm.trim() !== '') {
        const query = searchTerm.toLowerCase();
        const matchesField = Object.entries(row).some(([key, value]) => {
          if (key.startsWith('_')) return false;
          return String(value || '').toLowerCase().includes(query);
        });
        if (!matchesField) return false;
      }

      return true;
    });
  }, [data, selectedFileFilter, searchTerm]);

  // Sort data
  const sortedData = useMemo(() => {
    if (!sortColumn) return filteredData;

    return [...filteredData].sort((a, b) => {
      const valA = a[sortColumn];
      const valB = b[sortColumn];

      if (valA === valB) return 0;
      if (valA === undefined || valA === null) return 1;
      if (valB === undefined || valB === null) return -1;

      const strA = String(valA).toLowerCase();
      const strB = String(valB).toLowerCase();

      return sortDirection === 'asc' ? strA.localeCompare(strB) : strB.localeCompare(strA);
    });
  }, [filteredData, sortColumn, sortDirection]);

  // Pagination
  const totalPages = Math.ceil(sortedData.length / pageSize) || 1;
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return sortedData.slice(start, start + pageSize);
  }, [sortedData, currentPage, pageSize]);

  const handleSort = (column: string) => {
    if (sortColumn === column) {
      if (sortDirection === 'asc') {
        setSortDirection('desc');
      } else {
        setSortColumn(null);
        setSortDirection('asc');
      }
    } else {
      setSortColumn(column);
      setSortDirection('asc');
    }
  };

  return (
    <div className="space-y-3.5">
      {/* SECTION 1 TOOLBAR */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 p-3.5 shadow-sm">
        <div className="flex flex-col xl:flex-row xl:items-center xl:justify-between gap-3">
          {/* Left: Search & Filter & Sample */}
          <div className="flex flex-wrap items-center gap-2.5 flex-1">
            {/* Search */}
            <div className="relative flex-1 min-w-[220px] max-w-sm">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                id="table-search-input"
                type="text"
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                placeholder="Tìm theo TAG, SAP Code, Quy cách, Kích cỡ..."
                className="w-full pl-9 pr-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg text-xs text-slate-900 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all"
              />
            </div>

            {/* Source File Filter */}
            {sourceFiles.length > 1 && (
              <div className="flex items-center gap-1.5 text-xs text-slate-500">
                <Filter className="w-3.5 h-3.5 text-slate-400" />
                <select
                  id="source-file-filter-select"
                  value={selectedFileFilter}
                  onChange={(e) => {
                    setSelectedFileFilter(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-2.5 py-2 text-xs text-slate-800 dark:text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="all">Tất cả file nguồn ({data.length} hàng)</option>
                  {sourceFiles.map((file) => (
                    <option key={file} value={file}>
                      {file}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* Load Image Sample 3 Rows */}
            {onLoadImageSample && (
              <button
                id="btn-load-image-sample"
                onClick={onLoadImageSample}
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-bold rounded-lg bg-yellow-400 hover:bg-yellow-500 text-slate-950 shadow-xs transition-colors"
                title="Nạp ngay 3 dòng vật tư mẫu như ảnh: A5051220602A, gạch ngang, và V5518319100B"
              >
                <Sparkles className="w-3.5 h-3.5 text-slate-950" />
                <span>Nạp 3 dòng mẫu như ảnh</span>
              </button>
            )}

            {/* Add new valve row */}
            {onAddRow && (
              <button
                id="btn-add-valve-row"
                onClick={() => onAddRow()}
                className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-800 dark:text-slate-200 border border-slate-300 dark:border-slate-700 transition-colors"
                title="Thêm một dòng van mới vào bảng Armature"
              >
                <PlusCircle className="w-3.5 h-3.5 text-blue-600" />
                <span>+ Thêm van mới</span>
              </button>
            )}
          </div>

          {/* Right: Actions */}
          <div className="flex flex-wrap items-center gap-2">
            {/* Toggle File Import Panel */}
            {onToggleFileUploader && (
              <button
                id="btn-toggle-file-uploader"
                onClick={onToggleFileUploader}
                className={`inline-flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold border transition-all ${
                  isFileUploaderOpen
                    ? 'bg-blue-50 text-blue-700 border-blue-300 dark:bg-blue-950/60 dark:text-blue-300 dark:border-blue-800'
                    : 'bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 border-slate-200 dark:border-slate-700 hover:bg-slate-50'
                }`}
                title="Mở bảng kéo thả và quản lý file Excel nguồn để import vào Master"
              >
                <FileUp className="w-3.5 h-3.5 text-blue-600" />
                <span>Nhập File Excel</span>
                {totalFilesCount > 0 && (
                  <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-blue-100 dark:bg-blue-900 text-blue-800 dark:text-blue-200 font-bold">
                    {totalFilesCount}
                  </span>
                )}
              </button>
            )}

            {/* View Import Logs */}
            {onOpenLogs && (
              <button
                id="btn-open-logs-modal"
                onClick={onOpenLogs}
                className="inline-flex items-center gap-1.5 px-2.5 py-2 rounded-lg bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-slate-700 hover:bg-slate-50 text-xs font-medium transition-colors"
                title="Xem lịch sử và nhật ký Import_Log"
              >
                <Clock className="w-3.5 h-3.5 text-slate-500" />
                <span>Nhật ký</span>
                {logsCount > 0 && (
                  <span className="text-[10px] text-slate-400">({logsCount})</span>
                )}
              </button>
            )}

            {/* Column Aliases Mapping */}
            {onOpenAliases && (
              <button
                id="btn-open-aliases-modal"
                onClick={onOpenAliases}
                className="inline-flex items-center gap-1.5 px-2.5 py-2 rounded-lg bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-slate-700 hover:bg-slate-50 text-xs font-medium transition-colors"
                title="Cấu hình từ đồng nghĩa cho tên cột"
              >
                <SlidersHorizontal className="w-3.5 h-3.5 text-slate-500" />
                <span className="hidden sm:inline">Ánh xạ cột</span>
              </button>
            )}

            {/* Outlook Mail */}
            <button
              id="btn-table-batch-mail"
              onClick={onOpenBatchMail}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg bg-sky-50 dark:bg-sky-950/40 text-sky-700 dark:text-sky-300 hover:bg-sky-100 dark:hover:bg-sky-900/50 text-xs font-semibold border border-sky-200 dark:border-sky-800 transition-colors"
              title="Soạn thảo và gửi email báo cáo qua Outlook"
            >
              <Mail className="w-3.5 h-3.5 text-sky-600" />
              <span>Gửi Mail ({data.length})</span>
            </button>

            {/* Export Excel (.xlsx) */}
            <button
              id="btn-table-export-excel"
              onClick={onExportExcel}
              disabled={data.length === 0}
              className={`inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-xs font-semibold shadow-xs transition-colors ${
                data.length > 0
                  ? 'bg-emerald-600 hover:bg-emerald-700 text-white cursor-pointer'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-400 border border-slate-200 dark:border-slate-700 cursor-not-allowed opacity-60'
              }`}
              title="Xuất file Excel gồm 2 Sheet: Total & Import_Log theo cấu trúc chuẩn 2 tầng (Vàng / Xám)"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Xuất Excel (.xlsx)</span>
            </button>

            {/* Export CSV */}
            <button
              id="btn-table-export-csv"
              onClick={onExportCSV}
              disabled={data.length === 0}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-750 text-slate-700 dark:text-slate-300 text-xs font-medium transition-colors"
              title="Xuất dữ liệu bảng Armature Master sang định dạng CSV"
            >
              <span>CSV</span>
            </button>
          </div>
        </div>
      </div>

      {/* Table Empty State */}
      {data.length === 0 ? (
        <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 p-12 text-center shadow-sm">
          <div className="w-16 h-16 rounded-2xl bg-amber-50 dark:bg-amber-950/40 text-amber-500 mx-auto flex items-center justify-center mb-4">
            <FileSpreadsheet className="w-8 h-8" />
          </div>
          <h3 className="text-base font-semibold text-slate-900 dark:text-slate-100">
            Chưa có dữ liệu trong bảng Armature Master
          </h3>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 max-w-md mx-auto">
            Bấm <strong>&quot;Nhập File Excel&quot;</strong> để tải các file Excel nguồn, hoặc bấm <strong>&quot;Nạp 3 dòng mẫu như ảnh&quot;</strong> để xem ngay cấu trúc 2 tầng chuẩn Pipe Specification.
          </p>
          <div className="mt-4 flex flex-wrap items-center justify-center gap-3">
            {onLoadImageSample && (
              <button
                onClick={onLoadImageSample}
                className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-yellow-400 hover:bg-yellow-500 text-slate-950 font-bold text-xs shadow-sm transition-all cursor-pointer"
              >
                <Sparkles className="w-4 h-4" />
                <span>Nạp 3 dòng mẫu như ảnh</span>
              </button>
            )}
            {onToggleFileUploader && (
              <button
                onClick={onToggleFileUploader}
                className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-semibold text-xs shadow-sm transition-all cursor-pointer"
              >
                <FileUp className="w-4 h-4" />
                <span>Nhập File Excel Nguồn</span>
              </button>
            )}
          </div>
        </div>
      ) : (
        /* Table Container - 2-Tier Visual Header */
        <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-300 dark:border-slate-800 shadow-sm overflow-hidden flex flex-col flex-1">
          <div className="overflow-x-auto max-h-[calc(100vh-230px)] min-h-[500px]">
            <table className="w-full text-left border-collapse text-xs">
              {/* STICKY 2-TIER HEADER */}
              <thead className="sticky top-0 z-20 shadow-xs">
                {/* LEVEL 1: SUPER HEADERS (Yellow & Gray Groups) */}
                <tr className="border-b border-slate-300 dark:border-slate-700">
                  {/* Empty space for index and source file */}
                  <th
                    colSpan={2}
                    className="bg-slate-200 dark:bg-slate-850 px-3 py-2 text-center text-[10px] font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400 border-r border-slate-300 dark:border-slate-700"
                  >
                    NGUỒN DỮ LIỆU
                  </th>

                  {/* GROUP 1: TO BE COMPLETED BY FUNCTIONAL DESIGN (Màu Vàng #FFFF00) */}
                  {functionalDesignColsCount > 0 && (
                    <th
                      colSpan={functionalDesignColsCount}
                      style={{ backgroundColor: '#FFFF00', color: '#000000' }}
                      className="px-4 py-2.5 text-center text-xs font-black uppercase tracking-wider border-r border-yellow-400 shadow-inner"
                    >
                      TO BE COMPLETED BY FUNCTIONAL DESIGN
                    </th>
                  )}

                  {/* GROUP 2: ARMATURE INFO FROM PIPE SPECIFICATION (Màu Xám Kim Loại #BFBFBF) */}
                  {masterHeaders.length > functionalDesignColsCount && (
                    <th
                      colSpan={masterHeaders.length - functionalDesignColsCount}
                      style={{ backgroundColor: '#BFBFBF', color: '#000000' }}
                      className="px-4 py-2.5 text-center text-xs font-black uppercase tracking-wider border-r border-slate-400 shadow-inner"
                    >
                      ARMATURE INFO FROM PIPE SPECIFICATION
                    </th>
                  )}

                  {/* Actions Column */}
                  <th className="bg-slate-200 dark:bg-slate-850 px-3 py-2 text-center text-[11px] font-bold uppercase tracking-wider text-slate-600 dark:text-slate-400">
                    THAO TÁC
                  </th>
                </tr>

                {/* LEVEL 2: INDIVIDUAL COLUMN HEADERS WITH SORT INDICATOR */}
                <tr className="border-b border-slate-300 dark:border-slate-700 font-semibold text-[11px]">
                  <th className="py-2.5 px-3 w-10 text-center text-slate-400 font-mono bg-slate-100 dark:bg-slate-800 border-r border-slate-200 dark:border-slate-700">
                    #
                  </th>
                  <th className="py-2.5 px-3 text-slate-600 dark:text-slate-400 font-mono text-[11px] whitespace-nowrap bg-slate-100 dark:bg-slate-800 border-r border-slate-200 dark:border-slate-700">
                    File Nguồn
                  </th>

                  {masterHeaders.map((header, colIdx) => {
                    const isSorted = sortColumn === header;
                    const isFunctionalGroup = colIdx < functionalDesignColsCount;

                    return (
                      <th
                        key={header}
                        onClick={() => handleSort(header)}
                        className={`py-2 px-3 cursor-pointer transition-colors whitespace-nowrap select-none border-r border-slate-300 dark:border-slate-700 ${
                          isFunctionalGroup
                            ? 'bg-yellow-200/90 dark:bg-yellow-950/70 hover:bg-yellow-300/80 text-slate-900 dark:text-yellow-200'
                            : 'bg-slate-200/90 dark:bg-slate-800 hover:bg-slate-300/80 text-slate-800 dark:text-slate-200 font-semibold'
                        }`}
                        title={`Nhấp để sắp xếp theo ${header}`}
                      >
                        <div className="flex items-center justify-between gap-2">
                          <span className="truncate">{header}</span>
                          {isSorted ? (
                            sortDirection === 'asc' ? (
                              <ArrowUp className="w-3.5 h-3.5 text-blue-700 dark:text-blue-400 shrink-0 font-bold" />
                            ) : (
                              <ArrowDown className="w-3.5 h-3.5 text-blue-700 dark:text-blue-400 shrink-0 font-bold" />
                            )
                          ) : (
                            <ArrowUpDown className="w-3 h-3 text-slate-500 opacity-40 hover:opacity-100 shrink-0" />
                          )}
                        </div>
                      </th>
                    );
                  })}

                  <th className="py-2 px-3 text-center w-28 whitespace-nowrap bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400">
                    Thao tác
                  </th>
                </tr>
              </thead>

              {/* TABLE BODY */}
              <tbody className="divide-y divide-slate-200 dark:divide-slate-800 font-normal">
                {paginatedData.map((row, idx) => {
                  const rowIndex = (currentPage - 1) * pageSize + idx + 1;

                  return (
                    <tr
                      key={row._id}
                      className="hover:bg-amber-50/40 dark:hover:bg-slate-850/60 transition-colors"
                    >
                      <td className="py-2.5 px-3 text-center text-slate-400 font-mono text-[11px] border-r border-slate-100 dark:border-slate-800">
                        {rowIndex}
                      </td>
                      <td className="py-2.5 px-3 text-slate-500 text-[11px] font-mono whitespace-nowrap border-r border-slate-100 dark:border-slate-800">
                        <span className="truncate max-w-[130px] block" title={row._sourceFile}>
                          {row._sourceFile}
                        </span>
                      </td>

                      {masterHeaders.map((header, colIdx) => {
                        const val = row[header];
                        const isEmpty = val === undefined || val === null || String(val).trim() === '';
                        const isFunctionalGroup = colIdx < functionalDesignColsCount;
                        const isTag = header === 'TAG';
                        const isEditingThis =
                          editingCell?.rowId === row._id && editingCell?.header === header;

                        if (isEditingThis) {
                          return (
                            <td
                              key={header}
                              className="py-1 px-1.5 whitespace-nowrap text-xs border-r border-slate-200 dark:border-slate-800 bg-blue-50/40 dark:bg-blue-950/20"
                            >
                              <input
                                autoFocus
                                type="text"
                                value={editingValue}
                                onChange={(e) => setEditingValue(e.target.value)}
                                onBlur={() => handleCommitInline(row, header, editingValue)}
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') handleCommitInline(row, header, editingValue);
                                  if (e.key === 'Escape') setEditingCell(null);
                                }}
                                className="px-2 py-1 text-xs font-mono bg-white dark:bg-slate-800 border-2 border-blue-500 rounded focus:outline-none w-full min-w-[100px]"
                              />
                            </td>
                          );
                        }

                        return (
                          <td
                            key={header}
                            onClick={() => {
                              if (isTag) {
                                onEditRow(row);
                              } else {
                                setEditingCell({ rowId: row._id, header });
                                setEditingValue(String(val || ''));
                              }
                            }}
                            className={`py-2.5 px-3.5 whitespace-nowrap text-xs border-r border-slate-100 dark:border-slate-800 cursor-pointer ${
                              isFunctionalGroup ? 'bg-amber-50/20 dark:bg-yellow-950/5' : ''
                            } ${
                              isEmpty
                                ? 'text-slate-300 dark:text-slate-600 italic'
                                : isTag
                                ? 'font-bold text-blue-600 dark:text-blue-400 hover:underline'
                                : 'text-slate-800 dark:text-slate-200'
                            }`}
                            title={
                              isTag
                                ? 'Nhấp để chỉnh sửa van này'
                                : 'Nhấp để sửa giá trị ô này'
                            }
                          >
                            {isEmpty ? '—' : String(val)}
                          </td>
                        );
                      })}

                      {/* Action buttons per row */}
                      <td className="py-2.5 px-3 text-center">
                        <div className="flex items-center justify-center gap-1">
                          <button
                            onClick={() => onEditRow(row)}
                            className="text-slate-500 hover:text-blue-600 dark:text-slate-400 dark:hover:text-blue-400 transition-colors p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer"
                            title="Chỉnh sửa thông số van"
                          >
                            <Edit3 className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => onMailRow(row)}
                            className="text-slate-500 hover:text-sky-600 dark:text-slate-400 dark:hover:text-sky-400 transition-colors p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer"
                            title="Soạn mail Outlook về van này"
                          >
                            <Mail className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => onDeleteRow(row._id)}
                            className="text-slate-400 hover:text-rose-500 transition-colors p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer"
                            title="Xoá hàng này"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {/* Footer & Pagination */}
          <div className="px-4 py-3 bg-slate-50 dark:bg-slate-850 border-t border-slate-200 dark:border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-600 dark:text-slate-400">
            <div className="flex items-center gap-2">
              <span>
                Hiển thị{' '}
                <strong>
                  {Math.min((currentPage - 1) * pageSize + 1, sortedData.length)} -{' '}
                  {Math.min(currentPage * pageSize, sortedData.length)}
                </strong>{' '}
                trong tổng số <strong>{sortedData.length}</strong> hàng
                {filteredData.length !== data.length && ` (Lọc từ ${data.length})`}
              </span>

              <span className="mx-2 text-slate-300 dark:text-slate-700">|</span>

              <span>Mỗi trang:</span>
              <select
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setCurrentPage(1);
                }}
                className="px-2 py-1 rounded border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-xs"
              >
                <option value={10}>10</option>
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
              </select>
            </div>

            <div className="flex items-center gap-1.5">
              <button
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="p-1.5 rounded-md border border-slate-200 dark:border-slate-700 disabled:opacity-40 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors cursor-pointer"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <span className="px-2 font-medium">
                Trang {currentPage} / {totalPages}
              </span>
              <button
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="p-1.5 rounded-md border border-slate-200 dark:border-slate-700 disabled:opacity-40 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors cursor-pointer"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
