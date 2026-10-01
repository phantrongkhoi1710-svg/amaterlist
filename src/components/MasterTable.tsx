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
  RotateCcw,
  CheckCircle2,
  Palette,
  CheckSquare,
  Square,
  X,
  AlertTriangle,
} from 'lucide-react';
import { MasterRowData, RowRevisionStatus } from '../types';
import { applyRowRevisionStatus } from '../utils/revisionManager';

interface MasterTableProps {
  data: MasterRowData[];
  masterHeaders: string[];
  onDeleteRow: (rowId: string) => void;
  onDeleteRows?: (rowIds: string[]) => void;
  onEditRow: (row: MasterRowData) => void;
  onUpdateRow?: (updatedRow: MasterRowData) => void;
  onAddRow?: (newRow?: MasterRowData) => void;
  onLoadImageSample?: () => void;
  onResetToImageHeaders?: () => void;
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

const QUICK_COLORS = [
  { label: 'Yellow (New/Updated)', hex: '#FFFF00', bg: 'bg-yellow-400', status: 'new_updated' as RowRevisionStatus },
  { label: 'Red Strikethrough (Delete Line)', hex: '#EF4444', bg: 'bg-rose-500', status: 'deleted' as RowRevisionStatus },
  { label: 'Green (Checked)', hex: '#86EFAC', bg: 'bg-emerald-400', status: 'normal' as RowRevisionStatus },
  { label: 'Blue (Design)', hex: '#93C5FD', bg: 'bg-sky-400', status: 'normal' as RowRevisionStatus },
  { label: 'Orange (Notice/Warning)', hex: '#FDBA74', bg: 'bg-orange-400', status: 'normal' as RowRevisionStatus },
  { label: 'Purple (Spare)', hex: '#D8B4FE', bg: 'bg-purple-400', status: 'normal' as RowRevisionStatus },
];

export const MasterTable: React.FC<MasterTableProps> = ({
  data,
  masterHeaders,
  onDeleteRow,
  onDeleteRows,
  onEditRow,
  onUpdateRow,
  onAddRow,
  onLoadImageSample,
  onResetToImageHeaders,
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
  const [selectedRevisionFilter, setSelectedRevisionFilter] = useState<'all' | RowRevisionStatus>('all');
  const [sortColumn, setSortColumn] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);

  // Multi-row selection for bulk delete / bulk color loading
  const [selectedRowIds, setSelectedRowIds] = useState<Set<string>>(new Set());

  // In-app Delete Confirmation Modal (no window.confirm)
  const [deleteModal, setDeleteModal] = useState<{
    type: 'single' | 'bulk';
    row?: MasterRowData;
    ids?: string[];
  } | null>(null);

  // In-app Next Rev Confirmation Modal
  const [isNextRevModalOpen, setIsNextRevModalOpen] = useState(false);

  // Row Color Palette Menu
  const [activeColorRowId, setActiveColorRowId] = useState<string | null>(null);

  // Editing cell state (for direct inline cell changes)
  const [editingCell, setEditingCell] = useState<{ rowId: string; header: string } | null>(null);
  const [editingValue, setEditingValue] = useState<string>('');

  // Determine split indices for three groups matching the uploaded image:
  // Group 1: Cols 1-5 (Yellow) - TO BE COMPLETED BY FUNCTIONAL DESIGN
  // Group 2: Cols 6-21 (Metallic Grey) - ARMATURE INFO FROM PIPE SPECIFICATION
  // Group 3: Cols 22-26 (White) - REVISION & SIGNATURE
  const group1Count = useMemo(() => {
    const idx = masterHeaders.findIndex(
      (h) =>
        h.toUpperCase().replace(/\s+/g, '') === 'SAPCODE' ||
        h.toUpperCase().includes('STD DRW') ||
        h.toUpperCase().includes('NORMALE') ||
        h.toUpperCase() === 'EXECUTION'
    );
    return idx !== -1 ? idx : Math.min(5, masterHeaders.length);
  }, [masterHeaders]);

  const group2Count = useMemo(() => {
    const idx = masterHeaders.findIndex(
      (h) =>
        h.toUpperCase().includes('INPUT') ||
        h.toUpperCase().includes('REV.') ||
        h.toUpperCase().includes('REV HIS') ||
        h.toUpperCase() === 'SIGNATURE'
    );
    if (idx !== -1 && idx > group1Count) {
      return idx - group1Count;
    }
    return Math.min(16, Math.max(0, masterHeaders.length - group1Count));
  }, [masterHeaders, group1Count]);

  const group3Count = useMemo(() => {
    return Math.max(0, masterHeaders.length - group1Count - group2Count);
  }, [masterHeaders, group1Count, group2Count]);

  // Statistics for revisions
  const revisionStats = useMemo(() => {
    let newUpdated = 0;
    let deleted = 0;
    let nextRev = 0;
    data.forEach((r) => {
      if (r._revisionStatus === 'new_updated') newUpdated++;
      else if (r._revisionStatus === 'deleted') deleted++;
      else if (r._revisionStatus === 'next_rev_after_deleted') nextRev++;
    });
    return { newUpdated, deleted, nextRev };
  }, [data]);

  // Get unique source files
  const sourceFiles = useMemo(() => {
    const files = new Set<string>();
    data.forEach((r) => {
      if (r._sourceFile) files.add(r._sourceFile);
    });
    return Array.from(files);
  }, [data]);

  // Filter data
  const filteredData = useMemo(() => {
    return data.filter((row) => {
      if (selectedFileFilter !== 'all' && row._sourceFile !== selectedFileFilter) {
        return false;
      }

      if (selectedRevisionFilter !== 'all') {
        const rowStatus = row._revisionStatus || 'normal';
        if (rowStatus !== selectedRevisionFilter) {
          return false;
        }
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
  }, [data, selectedFileFilter, selectedRevisionFilter, searchTerm]);

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

  // Selection helpers
  const isAllPaginatedSelected = useMemo(() => {
    if (paginatedData.length === 0) return false;
    return paginatedData.every((r) => selectedRowIds.has(r._id));
  }, [paginatedData, selectedRowIds]);

  const handleToggleSelectAll = () => {
    if (isAllPaginatedSelected) {
      // Unselect all on current page
      setSelectedRowIds((prev) => {
        const next = new Set(prev);
        paginatedData.forEach((r) => next.delete(r._id));
        return next;
      });
    } else {
      // Select all on current page
      setSelectedRowIds((prev) => {
        const next = new Set(prev);
        paginatedData.forEach((r) => next.add(r._id));
        return next;
      });
    }
  };

  const handleToggleSelectRow = (rowId: string) => {
    setSelectedRowIds((prev) => {
      const next = new Set(prev);
      if (next.has(rowId)) {
        next.delete(rowId);
      } else {
        next.add(rowId);
      }
      return next;
    });
  };

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

  // Toggle row revision status (Yellow / Red Strikethrough / Next Rev)
  const handleToggleRowStatus = (row: MasterRowData, status: RowRevisionStatus) => {
    if (!onUpdateRow) return;
    const targetStatus = row._revisionStatus === status ? 'normal' : status;
    const updated = applyRowRevisionStatus(row, targetStatus);
    onUpdateRow(updated);
  };

  // Toggle strikethrough for a single row independently
  const handleToggleRowStrikethrough = (row: MasterRowData) => {
    if (!onUpdateRow) return;
    const curStrike = row._revisionStatus === 'deleted' || !!row._hasStrikethrough;
    const updated: MasterRowData = {
      ...row,
      _hasStrikethrough: !curStrike,
    };
    onUpdateRow(updated);
  };

  // Set custom color, revision status, and strikethrough on a row
  const handleSetRowColor = (
    row: MasterRowData,
    colorHex?: string,
    status?: RowRevisionStatus,
    overrideStrike?: boolean
  ) => {
    if (!onUpdateRow) return;
    let updated = { ...row };
    if (status) {
      updated = applyRowRevisionStatus(updated, status);
    }
    updated._detectedColor = colorHex;
    if (overrideStrike !== undefined) {
      updated._hasStrikethrough = overrideStrike;
    } else if (colorHex === '#EF4444' || status === 'deleted') {
      updated._hasStrikethrough = true;
      updated._revisionStatus = 'deleted';
    } else if (colorHex === '#FFFF00' || status === 'new_updated') {
      updated._revisionStatus = 'new_updated';
    } else if (!colorHex) {
      updated._hasStrikethrough = false;
      updated._revisionStatus = 'normal';
    }
    onUpdateRow(updated);
    setActiveColorRowId(null);
  };

  // Bulk status update for selected rows
  const handleBulkApplyStatus = (status: RowRevisionStatus, customColor?: string) => {
    if (!onUpdateRow || selectedRowIds.size === 0) return;
    data.forEach((r) => {
      if (selectedRowIds.has(r._id)) {
        let updated = applyRowRevisionStatus(r, status);
        if (customColor) {
          updated._detectedColor = customColor;
        }
        onUpdateRow(updated);
      }
    });
  };

  // Bulk strikethrough toggle for selected rows
  const handleBulkToggleStrikethrough = () => {
    if (!onUpdateRow || selectedRowIds.size === 0) return;
    // Check if any is not struck through
    const anyNotStruck = data.some((r) => selectedRowIds.has(r._id) && !r._hasStrikethrough && r._revisionStatus !== 'deleted');
    data.forEach((r) => {
      if (selectedRowIds.has(r._id)) {
        onUpdateRow({
          ...r,
          _hasStrikethrough: anyNotStruck,
        });
      }
    });
  };

  // Execute confirmed deletion (Single row or Bulk rows)
  const handleExecuteDelete = () => {
    if (!deleteModal) return;
    if (deleteModal.type === 'single' && deleteModal.row) {
      onDeleteRow(deleteModal.row._id);
      setSelectedRowIds((prev) => {
        const next = new Set(prev);
        next.delete(deleteModal.row!._id);
        return next;
      });
    } else if (deleteModal.type === 'bulk' && deleteModal.ids) {
      if (onDeleteRows) {
        onDeleteRows(deleteModal.ids);
      } else {
        deleteModal.ids.forEach((id) => onDeleteRow(id));
      }
      setSelectedRowIds(new Set());
    }
    setDeleteModal(null);
  };

  // Transition all deleted rows to Next Rev
  const handleExecuteNextRevForDeleted = () => {
    if (!onUpdateRow) return;
    const deletedRows = data.filter((r) => r._revisionStatus === 'deleted');
    deletedRows.forEach((r) => {
      onUpdateRow(applyRowRevisionStatus(r, 'next_rev_after_deleted'));
    });
    setIsNextRevModalOpen(false);
  };

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
      {/* FLOATING BULK SELECTION ACTION BAR */}
      {selectedRowIds.size > 0 && (
        <div className="bg-slate-900 text-white rounded-xl p-3 shadow-lg border border-slate-700 flex flex-wrap items-center justify-between gap-3 animate-in fade-in duration-150">
          <div className="flex items-center gap-2">
            <span className="px-2.5 py-1 rounded-md bg-blue-600 text-white text-xs font-bold shadow-xs">
              Selected {selectedRowIds.size} valve rows
            </span>
            <span className="text-xs text-slate-300 hidden md:inline">
              Bulk Actions (Colors &amp; Line Deletions):
            </span>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {/* 🟡 Apply Yellow */}
            <button
              type="button"
              onClick={() => handleBulkApplyStatus('new_updated')}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-yellow-400 text-slate-950 hover:bg-yellow-300 transition-colors cursor-pointer"
              title="Apply yellow color to selected rows (New / Updated valve)"
            >
              <span className="w-2 h-2 rounded-full bg-yellow-600 inline-block"></span>
              <span>Apply Yellow</span>
            </button>

            {/* 🔴 Delete Line */}
            <button
              type="button"
              onClick={() => handleBulkApplyStatus('deleted')}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-rose-600 text-white hover:bg-rose-500 transition-colors cursor-pointer line-through"
              title="Delete line (Red colour with strikethrough)"
            >
              <span className="w-2 h-2 rounded-full bg-rose-300 inline-block"></span>
              <span>Delete Line</span>
            </button>

            {/* <s> Strikethrough */}
            <button
              type="button"
              onClick={handleBulkToggleStrikethrough}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-800 text-rose-300 border border-rose-500/50 hover:bg-slate-700 transition-colors cursor-pointer"
              title="Toggle strikethrough for selected rows"
            >
              <span className="font-serif text-sm font-black"><s>S</s></span>
              <span>Strikethrough ({selectedRowIds.size})</span>
            </button>

            {/* ⚪ Next Revision */}
            <button
              type="button"
              onClick={() => handleBulkApplyStatus('next_rev_after_deleted')}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-semibold bg-slate-800 text-slate-200 hover:bg-slate-700 border border-slate-700 transition-colors cursor-pointer"
              title="Apply Next Rev (Keep SFI & TAG, empty rest)"
            >
              <span>Next Rev</span>
            </button>

            {/* 🧹 Clear Color */}
            <button
              type="button"
              onClick={() => handleBulkApplyStatus('normal')}
              className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-slate-800 text-slate-300 hover:bg-slate-700 transition-colors cursor-pointer"
              title="Reset status to normal (No color)"
            >
              <span>Clear Color</span>
            </button>

            <span className="w-px h-5 bg-slate-700 mx-1 hidden sm:block"></span>

            {/* 🗑️ Delete Selected Rows */}
            <button
              type="button"
              onClick={() => setDeleteModal({ type: 'bulk', ids: Array.from(selectedRowIds) })}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-rose-700 hover:bg-rose-600 text-white transition-colors cursor-pointer shadow-sm"
              title="Permanently delete selected rows from Master table"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>Delete {selectedRowIds.size} rows</span>
            </button>

            {/* ✕ Deselect */}
            <button
              type="button"
              onClick={() => setSelectedRowIds(new Set())}
              className="px-2 py-1.5 rounded-lg text-xs text-slate-400 hover:text-white transition-colors cursor-pointer"
            >
              ✕ Deselect
            </button>
          </div>
        </div>
      )}

      {/* SECTION 1 TOOLBAR */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 p-3.5 shadow-sm">
        <div className="flex flex-col xl:flex-row xl:items-center xl:justify-between gap-3">
          {/* Left: Search & Filter & Sample */}
          <div className="flex flex-wrap items-center gap-2.5 flex-1">
            {/* Search */}
            <div className="relative flex-1 min-w-[200px] max-w-xs">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                id="table-search-input"
                type="text"
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                placeholder="Search TAG, SAP Code, Specs..."
                className="w-full pl-9 pr-3 py-2 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg text-xs text-slate-900 dark:text-slate-100 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all"
              />
            </div>

            {/* Revision Status Filter Pills */}
            <div className="flex items-center bg-slate-100 dark:bg-slate-800 p-0.5 rounded-lg border border-slate-200 dark:border-slate-700 text-xs">
              <button
                type="button"
                onClick={() => {
                  setSelectedRevisionFilter('all');
                  setCurrentPage(1);
                }}
                className={`px-2 py-1 rounded-md transition-colors font-medium cursor-pointer ${
                  selectedRevisionFilter === 'all'
                    ? 'bg-white dark:bg-slate-700 text-slate-900 dark:text-slate-100 shadow-xs'
                    : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                }`}
              >
                All ({data.length})
              </button>
              <button
                type="button"
                onClick={() => {
                  setSelectedRevisionFilter('new_updated');
                  setCurrentPage(1);
                }}
                className={`px-2 py-1 rounded-md transition-colors flex items-center gap-1 font-bold cursor-pointer ${
                  selectedRevisionFilter === 'new_updated'
                    ? 'bg-yellow-400 text-slate-950 shadow-xs'
                    : 'text-yellow-700 dark:text-yellow-300 hover:bg-yellow-100/50'
                }`}
                title="Filter New or Updated valve rows (Yellow)"
              >
                <span className="w-2 h-2 rounded-full bg-yellow-400 inline-block border border-yellow-600"></span>
                <span>New ({revisionStats.newUpdated})</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setSelectedRevisionFilter('deleted');
                  setCurrentPage(1);
                }}
                className={`px-2 py-1 rounded-md transition-colors flex items-center gap-1 font-bold cursor-pointer ${
                  selectedRevisionFilter === 'deleted'
                    ? 'bg-rose-600 text-white shadow-xs line-through'
                    : 'text-rose-600 dark:text-rose-400 hover:bg-rose-100/50'
                }`}
                title="Filter Deleted valve rows (Red strikethrough)"
              >
                <span className="w-2 h-2 rounded-full bg-rose-500 inline-block border border-rose-700"></span>
                <span>Deleted ({revisionStats.deleted})</span>
              </button>

              {revisionStats.deleted > 0 && (
                <button
                  type="button"
                  onClick={() => setIsNextRevModalOpen(true)}
                  className="px-2 py-1 rounded-md text-[11px] font-bold bg-amber-400 hover:bg-amber-300 text-slate-950 transition-colors shadow-xs cursor-pointer ml-1"
                  title="Clean up deleted valves: keep SFI & TAG, empty remaining cells for next revision"
                >
                  Move {revisionStats.deleted} deleted to Next Rev
                </button>
              )}
              <button
                type="button"
                onClick={() => {
                  setSelectedRevisionFilter('next_rev_after_deleted');
                  setCurrentPage(1);
                }}
                className={`px-2 py-1 rounded-md transition-colors flex items-center gap-1 font-medium cursor-pointer ${
                  selectedRevisionFilter === 'next_rev_after_deleted'
                    ? 'bg-slate-300 dark:bg-slate-600 text-slate-900 dark:text-slate-100 shadow-xs'
                    : 'text-slate-500 dark:text-slate-400 hover:bg-slate-200/50'
                }`}
                title="Filter Next Rev valves after deletion (Keeps SFI & TAG)"
              >
                <span>Next Rev ({revisionStats.nextRev})</span>
              </button>
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
                  className="bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg px-2.5 py-1.5 text-xs text-slate-800 dark:text-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  <option value="all">All files ({data.length})</option>
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
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold rounded-lg bg-yellow-400 hover:bg-yellow-500 text-slate-950 shadow-xs transition-colors cursor-pointer"
                title="Load sample rows matching standard 26 columns and revision colors"
              >
                <Sparkles className="w-3.5 h-3.5 text-slate-950" />
                <span>Load 4 Color Samples</span>
              </button>
            )}

            {/* Restore Standard 26 Columns Button */}
            {onResetToImageHeaders && (
              <button
                id="btn-reset-to-image-headers"
                onClick={onResetToImageHeaders}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-800 dark:text-slate-200 border border-slate-300 dark:border-slate-700 transition-colors cursor-pointer"
                title="Restore standard 26 column headers: TAG -> DESTINATION (Yellow), SAP CODE -> SIGN TEXT (Gray), INPUT -> SIGNATURE (White)"
              >
                <SlidersHorizontal className="w-3.5 h-3.5 text-blue-600" />
                <span>Standard 26 Cols ({masterHeaders.length}/26)</span>
              </button>
            )}

            {/* Add new valve row */}
            {onAddRow && (
              <button
                id="btn-add-valve-row"
                onClick={() => onAddRow()}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-slate-200 dark:bg-slate-800 dark:hover:bg-slate-700 text-slate-800 dark:text-slate-200 border border-slate-300 dark:border-slate-700 transition-colors cursor-pointer"
                title="Add a new valve row to Armature table"
              >
                <PlusCircle className="w-3.5 h-3.5 text-blue-600" />
                <span>+ Add New Valve</span>
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
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold border transition-all cursor-pointer ${
                  isFileUploaderOpen
                    ? 'bg-blue-50 text-blue-700 border-blue-300 dark:bg-blue-950/60 dark:text-blue-300 dark:border-blue-800'
                    : 'bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 border-slate-200 dark:border-slate-700 hover:bg-slate-50'
                }`}
                title="Open drop zone and source Excel file manager"
              >
                <FileUp className="w-3.5 h-3.5 text-blue-600" />
                <span>Import Excel File</span>
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
                className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-slate-700 hover:bg-slate-50 text-xs font-medium transition-colors cursor-pointer"
                title="View import history and Import_Log"
              >
                <Clock className="w-3.5 h-3.5 text-slate-500" />
                <span>Logs</span>
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
                className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 border border-slate-200 dark:border-slate-700 hover:bg-slate-50 text-xs font-medium transition-colors cursor-pointer"
                title="Configure column header aliases"
              >
                <SlidersHorizontal className="w-3.5 h-3.5 text-slate-500" />
                <span className="hidden sm:inline">Column Aliases</span>
              </button>
            )}

            {/* Outlook Mail */}
            <button
              id="btn-table-batch-mail"
              onClick={onOpenBatchMail}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-sky-50 dark:bg-sky-950/40 text-sky-700 dark:text-sky-300 hover:bg-sky-100 dark:hover:bg-sky-900/50 text-xs font-semibold border border-sky-200 dark:border-sky-800 transition-colors cursor-pointer"
              title="Compose report email via Outlook"
            >
              <Mail className="w-3.5 h-3.5 text-sky-600" />
              <span>Email Report ({data.length})</span>
            </button>

            {/* Export Excel (.xlsx) with loaded colors */}
            <button
              id="btn-table-export-excel"
              onClick={onExportExcel}
              disabled={data.length === 0}
              className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-xs transition-colors ${
                data.length > 0
                  ? 'bg-emerald-600 hover:bg-emerald-700 text-white cursor-pointer'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-400 border border-slate-200 dark:border-slate-700 cursor-not-allowed opacity-60'
              }`}
              title="Export complete Excel file preserving revision colors: Yellow (New/Updated), Red Strikethrough (Deleted), and Next Rev"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Excel (.xlsx)</span>
            </button>

            {/* Export CSV */}
            <button
              id="btn-table-export-csv"
              onClick={onExportCSV}
              disabled={data.length === 0}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-750 text-slate-700 dark:text-slate-300 text-xs font-medium transition-colors cursor-pointer"
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
            No data in Armature Master table
          </h3>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 max-w-md mx-auto">
            Click <strong>&quot;Import Excel File&quot;</strong> to load source files, or click <strong>&quot;Load 4 Color Samples&quot;</strong> to test revision colors (yellow, red strikethrough, and next revision).
          </p>
          <div className="mt-4 flex flex-wrap items-center justify-center gap-3">
            {onLoadImageSample && (
              <button
                onClick={onLoadImageSample}
                className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-yellow-400 hover:bg-yellow-500 text-slate-950 font-bold text-xs shadow-sm transition-all cursor-pointer"
              >
                <Sparkles className="w-4 h-4" />
                <span>Load 4 Color Samples</span>
              </button>
            )}
            {onToggleFileUploader && (
              <button
                onClick={onToggleFileUploader}
                className="inline-flex items-center gap-2 px-4 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-semibold text-xs shadow-sm transition-all cursor-pointer"
              >
                <FileUp className="w-4 h-4" />
                <span>Import Source Excel File</span>
              </button>
            )}
          </div>
        </div>
      ) : (
        /* Table Container - 2-Tier Visual Header with Checkboxes & Direct Color / Delete Controls */
        <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-300 dark:border-slate-800 shadow-sm overflow-hidden flex flex-col flex-1">
          <div className="overflow-x-auto max-h-[calc(100vh-250px)] min-h-[500px]">
            <table className="w-full text-left border-collapse text-xs">
              {/* STICKY 2-TIER HEADER */}
              <thead className="sticky top-0 z-20 shadow-xs">
                {/* LEVEL 1: SUPER HEADERS (Yellow, Gray, White Groups) */}
                <tr className="border-b border-slate-300 dark:border-slate-700">
                  {/* Empty space for checkbox, index and source file */}
                  <th
                    colSpan={3}
                    className="bg-slate-200 dark:bg-slate-850 px-3 py-2 text-center text-[10px] font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400 border-r border-slate-300 dark:border-slate-700"
                  >
                    DATA SOURCE
                  </th>

                  {/* GROUP 1: TO BE COMPLETED BY FUNCTIONAL DESIGN (Màu Vàng #FFFF00) */}
                  {group1Count > 0 && (
                    <th
                      colSpan={group1Count}
                      style={{ backgroundColor: '#FFFF00', color: '#000000' }}
                      className="px-4 py-2.5 text-center text-xs font-black uppercase tracking-wider border-r border-yellow-400 shadow-inner"
                    >
                      TO BE COMPLETED BY FUNCTIONAL DESIGN
                    </th>
                  )}

                  {/* GROUP 2: ARMATURE INFO FROM PIPE SPECIFICATION (Màu Xám Kim Loại #BFBFBF) */}
                  {group2Count > 0 && (
                    <th
                      colSpan={group2Count}
                      style={{ backgroundColor: '#BFBFBF', color: '#000000' }}
                      className="px-4 py-2.5 text-center text-xs font-black uppercase tracking-wider border-r border-slate-400 shadow-inner"
                    >
                      ARMATURE INFO FROM PIPE SPECIFICATION
                    </th>
                  )}

                  {/* GROUP 3: REVISION & SIGNATURE (Màu Trắng/Sáng #FFFFFF) */}
                  {group3Count > 0 && (
                    <th
                      colSpan={group3Count}
                      style={{ backgroundColor: '#FFFFFF', color: '#000000' }}
                      className="px-4 py-2.5 text-center text-xs font-black uppercase tracking-wider border-r border-slate-300 shadow-inner"
                    >
                      REVISION &amp; SIGNATURE
                    </th>
                  )}

                  {/* Actions Column */}
                  <th className="bg-slate-200 dark:bg-slate-850 px-3 py-2 text-center text-[11px] font-bold uppercase tracking-wider text-slate-600 dark:text-slate-400 min-w-[200px]">
                    COLOR &amp; LINE DELETION
                  </th>
                </tr>

                {/* LEVEL 2: INDIVIDUAL COLUMN HEADERS WITH SORT INDICATOR */}
                <tr className="border-b border-slate-300 dark:border-slate-700 font-semibold text-[11px]">
                  {/* Select All Checkbox */}
                  <th className="py-2.5 px-2.5 w-8 text-center bg-slate-100 dark:bg-slate-800 border-r border-slate-200 dark:border-slate-700">
                    <input
                      type="checkbox"
                      checked={isAllPaginatedSelected}
                      onChange={handleToggleSelectAll}
                      className="w-3.5 h-3.5 text-blue-600 rounded border-slate-300 dark:border-slate-600 focus:ring-blue-500 cursor-pointer"
                      title={isAllPaginatedSelected ? 'Deselect current page' : 'Select all on this page'}
                    />
                  </th>
                  <th className="py-2.5 px-2.5 w-10 text-center text-slate-400 font-mono bg-slate-100 dark:bg-slate-800 border-r border-slate-200 dark:border-slate-700">
                    #
                  </th>
                  <th className="py-2.5 px-3 text-slate-600 dark:text-slate-400 font-mono text-[11px] whitespace-nowrap bg-slate-100 dark:bg-slate-800 border-r border-slate-200 dark:border-slate-700">
                    Source File
                  </th>

                  {masterHeaders.map((header, colIdx) => {
                    const isSorted = sortColumn === header;
                    const isGroup1 = colIdx < group1Count;
                    const isGroup2 = colIdx >= group1Count && colIdx < group1Count + group2Count;

                    return (
                      <th
                        key={header}
                        onClick={() => handleSort(header)}
                        className={`py-2 px-3 cursor-pointer transition-colors whitespace-nowrap select-none border-r ${
                          isGroup1
                            ? 'bg-[#FFFF00] hover:bg-yellow-400 text-slate-950 font-bold border-yellow-500'
                            : isGroup2
                            ? 'bg-[#BFBFBF] hover:bg-slate-350 text-slate-950 font-bold border-slate-400'
                            : 'bg-white hover:bg-slate-100 text-slate-950 font-semibold border-slate-300'
                        }`}
                        title={`Click to sort by ${header}`}
                      >
                        <div className="flex items-center justify-between gap-2">
                          <span className="truncate">{header}</span>
                          {isSorted ? (
                            sortDirection === 'asc' ? (
                              <ArrowUp className="w-3.5 h-3.5 text-blue-700 shrink-0 font-bold" />
                            ) : (
                              <ArrowDown className="w-3.5 h-3.5 text-blue-700 shrink-0 font-bold" />
                            )
                          ) : (
                            <ArrowUpDown className="w-3 h-3 text-slate-500 opacity-40 hover:opacity-100 shrink-0" />
                          )}
                        </div>
                      </th>
                    );
                  })}

                  <th className="py-2 px-3 text-center min-w-[200px] whitespace-nowrap bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 font-semibold">
                    Row Actions
                  </th>
                </tr>
              </thead>

              {/* TABLE BODY WITH CONDITIONAL REVISION STYLES */}
              <tbody className="divide-y divide-slate-200 dark:divide-slate-800 font-normal">
                {paginatedData.map((row, idx) => {
                  const rowIndex = (currentPage - 1) * pageSize + idx + 1;
                  const isNewUpdated = row._revisionStatus === 'new_updated';
                  const isDeleted = row._revisionStatus === 'deleted';
                  const isNextRev = row._revisionStatus === 'next_rev_after_deleted';
                  const isStrikethrough = isDeleted || !!row._hasStrikethrough;
                  const isSelected = selectedRowIds.has(row._id);

                  // Row background and typography matching PRINCIPLE FOR REVISION DESCRIPTIONS
                  let rowClasses = 'hover:bg-amber-50/40 dark:hover:bg-slate-850/60 transition-colors';
                  let customStyle: React.CSSProperties = {};

                  if (isDeleted) {
                    rowClasses = 'bg-rose-100 hover:bg-rose-200/90 dark:bg-rose-950/70 text-rose-950 dark:text-rose-100 line-through';
                    customStyle = { borderLeft: '4px solid #EF4444' };
                  } else if (isNewUpdated) {
                    rowClasses = `bg-yellow-100 hover:bg-yellow-200/90 dark:bg-yellow-950/70 text-slate-950 dark:text-yellow-100 font-medium ${isStrikethrough ? 'line-through opacity-85' : ''}`;
                    customStyle = { borderLeft: '4px solid #F59E0B' };
                  } else if (isNextRev) {
                    rowClasses = 'bg-slate-100/90 hover:bg-slate-200/80 dark:bg-slate-850 text-slate-500 dark:text-slate-400 italic';
                    customStyle = { borderLeft: '4px solid #64748B' };
                  } else if (row._detectedColor && row._detectedColor !== '#FFFFFF') {
                    // Custom loaded color
                    rowClasses = `hover:bg-amber-50/40 dark:hover:bg-slate-850/60 transition-colors ${isStrikethrough ? 'line-through text-slate-900 dark:text-slate-100' : ''}`;
                    customStyle = {
                      backgroundColor: `${row._detectedColor}18`,
                      borderLeft: `4px solid ${row._detectedColor}`,
                    };
                  } else if (isStrikethrough) {
                    rowClasses = 'bg-rose-50/60 hover:bg-rose-100/80 dark:bg-rose-950/30 text-rose-950 dark:text-rose-200 line-through transition-colors';
                    customStyle = { borderLeft: '4px solid #F43F5E' };
                  }

                  if (isSelected) {
                    rowClasses += ' ring-1 ring-blue-500 bg-blue-50/30 dark:bg-blue-950/20';
                  }

                  return (
                    <tr
                      key={row._id}
                      className={rowClasses}
                      style={customStyle}
                    >
                      {/* Checkbox */}
                      <td className="py-2.5 px-2.5 text-center border-r border-slate-200 dark:border-slate-800">
                        <input
                          type="checkbox"
                          checked={isSelected}
                          onChange={() => handleToggleSelectRow(row._id)}
                          className="w-3.5 h-3.5 text-blue-600 rounded border-slate-300 dark:border-slate-600 focus:ring-blue-500 cursor-pointer"
                        />
                      </td>

                      <td className="py-2.5 px-2.5 text-center text-slate-400 font-mono text-[11px] border-r border-slate-200 dark:border-slate-800">
                        {rowIndex}
                      </td>

                      <td className="py-2.5 px-3 text-slate-500 text-[11px] font-mono whitespace-nowrap border-r border-slate-200 dark:border-slate-800">
                        <span className="truncate max-w-[120px] block" title={row._sourceFile}>
                          {row._sourceFile}
                        </span>
                      </td>

                      {masterHeaders.map((header, colIdx) => {
                        const val = row[header];
                        const isEmpty = val === undefined || val === null || String(val).trim() === '';
                        const isFunctionalGroup = colIdx < group1Count;
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
                            className={`py-2.5 px-3 whitespace-nowrap text-xs border-r border-slate-200 dark:border-slate-800 cursor-pointer ${
                              isStrikethrough
                                ? 'line-through text-rose-950 dark:text-rose-200'
                                : isNewUpdated
                                ? 'text-slate-950 dark:text-yellow-100'
                                : isNextRev
                                ? 'text-slate-500 italic'
                                : isFunctionalGroup
                                ? 'bg-amber-50/20 dark:bg-yellow-950/5'
                                : ''
                            } ${
                              isEmpty
                                ? 'text-slate-300 dark:text-slate-600 italic'
                                : isTag
                                ? 'font-bold text-blue-700 dark:text-blue-300 hover:underline'
                                : ''
                            }`}
                            title={
                              isTag
                                ? 'Click to edit this valve'
                                : 'Click to edit cell value'
                            }
                          >
                            {isEmpty ? '—' : String(val)}
                          </td>
                        );
                      })}

                      {/* Action buttons per row */}
                      <td className="py-2 px-2.5 text-center whitespace-nowrap border-l border-slate-200 dark:border-slate-800">
                        <div className="flex items-center justify-center gap-1.5">
                          {/* 🟡 Quick Toggle Yellow (New / Updated) */}
                          <button
                            type="button"
                            onClick={() => handleToggleRowStatus(row, 'new_updated')}
                            className={`px-2 py-1 rounded-md flex items-center gap-1 text-[11px] font-bold transition-all cursor-pointer ${
                              isNewUpdated
                                ? 'bg-yellow-400 text-slate-950 ring-2 ring-yellow-500 shadow-xs'
                                : 'bg-yellow-100 hover:bg-yellow-200 text-yellow-900 border border-yellow-300'
                            }`}
                            title="Yellow: New or Updated valve"
                          >
                            <span className="w-2 h-2 rounded-full bg-yellow-500 inline-block"></span>
                            <span>Yellow</span>
                          </button>

                          {/* 🔴 Quick Toggle Red Strikethrough (Delete Line) */}
                          <button
                            type="button"
                            onClick={() => handleToggleRowStatus(row, 'deleted')}
                            className={`px-2 py-1 rounded-md flex items-center gap-1 text-[11px] font-bold transition-all cursor-pointer ${
                              isDeleted
                                ? 'bg-rose-600 text-white ring-2 ring-rose-700 shadow-xs line-through'
                                : 'bg-rose-100 hover:bg-rose-200 text-rose-900 border border-rose-300'
                            }`}
                            title="Delete line (Red with strikethrough)"
                          >
                            <span className="w-2 h-2 rounded-full bg-rose-500 inline-block"></span>
                            <span>Delete Line</span>
                          </button>

                          {/* <s> Toggle Strikethrough Independently */}
                          <button
                            type="button"
                            onClick={() => handleToggleRowStrikethrough(row)}
                            className={`w-6 h-6 rounded-md flex items-center justify-center font-serif text-xs font-bold transition-all cursor-pointer ${
                              isStrikethrough
                                ? 'bg-rose-600 text-white ring-2 ring-rose-500 shadow-xs'
                                : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300'
                            }`}
                            title={isStrikethrough ? 'Strikethrough active: click to disable' : 'Toggle strikethrough'}
                          >
                            <s>S</s>
                          </button>

                          {/* ⚪ Quick Toggle Next Rev after deleted */}
                          <button
                            type="button"
                            onClick={() => handleToggleRowStatus(row, 'next_rev_after_deleted')}
                            className={`w-6 h-6 rounded-md flex items-center justify-center text-[9px] font-bold transition-all cursor-pointer ${
                              isNextRev
                                ? 'bg-slate-700 text-white ring-2 ring-slate-800 shadow-xs'
                                : 'bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300'
                            }`}
                            title="Next Rev: Keep SFI & TAG, empty remaining cells"
                          >
                            N
                          </button>

                          {/* 🎨 Quick Color Palette Popover */}
                          <div className="relative">
                            <button
                              type="button"
                              onClick={() => setActiveColorRowId(activeColorRowId === row._id ? null : row._id)}
                              className="w-6 h-6 rounded-md flex items-center justify-center text-slate-500 hover:text-slate-800 dark:hover:text-slate-200 bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 border border-slate-300 dark:border-slate-700 transition-colors cursor-pointer"
                              title="Set or change row color"
                            >
                              <Palette className="w-3.5 h-3.5" />
                            </button>

                            {activeColorRowId === row._id && (
                              <div className="absolute right-0 top-full mt-1 z-30 p-2.5 bg-white dark:bg-slate-800 rounded-xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col gap-2 min-w-[170px] text-left">
                                <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">
                                  Color &amp; Strikethrough
                                </span>
                                <div className="grid grid-cols-3 gap-1.5">
                                  {QUICK_COLORS.map((c) => (
                                    <button
                                      key={c.hex}
                                      type="button"
                                      onClick={() => handleSetRowColor(row, c.hex, c.status)}
                                      className={`w-6 h-6 rounded-md ${c.bg} border border-slate-400/40 hover:scale-110 transition-transform cursor-pointer shadow-xs`}
                                      title={c.label}
                                    />
                                  ))}
                                </div>

                                {/* Custom HTML Color Input */}
                                <div className="flex items-center justify-between gap-1 pt-1 border-t border-slate-100 dark:border-slate-700">
                                  <label className="flex items-center gap-1 text-[10px] text-slate-600 dark:text-slate-300 cursor-pointer">
                                    <input
                                      type="color"
                                      value={row._detectedColor || '#FFFF00'}
                                      onChange={(e) => handleSetRowColor(row, e.target.value)}
                                      className="w-4 h-4 rounded cursor-pointer border-0 p-0"
                                    />
                                    <span>Custom color</span>
                                  </label>
                                  <button
                                    type="button"
                                    onClick={() => handleSetRowColor(row, undefined, 'normal', false)}
                                    className="text-[10px] text-slate-500 hover:text-slate-800 dark:hover:text-slate-200 cursor-pointer"
                                  >
                                    Clear color
                                  </button>
                                </div>

                                {/* Strikethrough checkbox */}
                                <label className="flex items-center gap-1.5 text-[11px] pt-1 border-t border-slate-100 dark:border-slate-700 cursor-pointer select-none">
                                  <input
                                    type="checkbox"
                                    checked={isStrikethrough}
                                    onChange={() => handleToggleRowStrikethrough(row)}
                                    className="w-3.5 h-3.5 text-rose-600 rounded cursor-pointer"
                                  />
                                  <span className={isStrikethrough ? 'line-through font-bold text-rose-600' : 'text-slate-700 dark:text-slate-300'}>
                                    Strikethrough Line
                                  </span>
                                </label>
                              </div>
                            )}
                          </div>

                          <span className="w-px h-3.5 bg-slate-300 dark:bg-slate-700 mx-0.5"></span>

                          {/* Edit Row modal */}
                          <button
                            type="button"
                            onClick={() => onEditRow(row)}
                            className="text-slate-500 hover:text-blue-600 dark:text-slate-400 dark:hover:text-blue-400 transition-colors p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer"
                            title="Edit valve parameters"
                          >
                            <Edit3 className="w-3.5 h-3.5" />
                          </button>

                          {/* Outlook Mail */}
                          <button
                            type="button"
                            onClick={() => onMailRow(row)}
                            className="text-slate-500 hover:text-sky-600 dark:text-slate-400 dark:hover:text-sky-400 transition-colors p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-800 cursor-pointer"
                            title="Compose Outlook mail"
                          >
                            <Mail className="w-3.5 h-3.5" />
                          </button>

                          {/* 🗑️ Delete Line */}
                          <button
                            type="button"
                            onClick={() => setDeleteModal({ type: 'single', row })}
                            className="text-slate-400 hover:text-rose-600 transition-colors p-1 rounded hover:bg-rose-50 dark:hover:bg-rose-950/40 cursor-pointer"
                            title="Permanently delete this valve row from Master table"
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
                Showing{' '}
                <strong>
                  {Math.min((currentPage - 1) * pageSize + 1, sortedData.length)} -{' '}
                  {Math.min(currentPage * pageSize, sortedData.length)}
                </strong>{' '}
                of <strong>{sortedData.length}</strong> rows
                {filteredData.length !== data.length && ` (Filtered from ${data.length})`}
              </span>

              <span className="mx-2 text-slate-300 dark:text-slate-700">|</span>

              <div className="flex items-center gap-1.5">
                <span>Rows per page:</span>
                <select
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(Number(e.target.value));
                    setCurrentPage(1);
                  }}
                  className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded px-2 py-0.5 text-xs text-slate-700 dark:text-slate-300 font-medium cursor-pointer"
                >
                  <option value={25}>25</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                  <option value={99999}>All ({sortedData.length})</option>
                </select>
              </div>
            </div>

            {/* Pagination Controls */}
            <div className="flex items-center gap-1">
              <button
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="p-1 rounded border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-50 cursor-pointer"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <span className="px-2 font-mono text-xs">
                {currentPage} / {totalPages}
              </span>
              <button
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="p-1 rounded border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300 disabled:opacity-40 disabled:cursor-not-allowed hover:bg-slate-50 cursor-pointer"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </div>
      )}

      {/* IN-APP CONFIRMATION MODAL: DELETE LINE(S) */}
      {deleteModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
          <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-md w-full border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden p-6 space-y-4">
            <div className="flex items-start gap-3">
              <div className="w-10 h-10 rounded-xl bg-rose-100 dark:bg-rose-950/60 text-rose-600 flex items-center justify-center shrink-0">
                <Trash2 className="w-5 h-5" />
              </div>
              <div className="flex-1">
                <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                  {deleteModal.type === 'single'
                    ? `Confirm row deletion: ${deleteModal.row?.TAG || 'item'}`
                    : `Confirm deletion of ${deleteModal.ids?.length} valve rows`}
                </h3>
                <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                  {deleteModal.type === 'single'
                    ? 'This valve row will be removed from the Master table.'
                    : `All ${deleteModal.ids?.length} selected valve rows will be removed from the Master table.`}
                </p>
              </div>
            </div>
            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setDeleteModal(null)}
                className="px-3.5 py-2 rounded-lg text-xs font-semibold text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleExecuteDelete}
                className="inline-flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-bold bg-rose-600 hover:bg-rose-700 text-white shadow-sm transition-colors cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Confirm Delete</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* IN-APP CONFIRMATION MODAL: NEXT REV FOR DELETED */}
      {isNextRevModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
          <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-md w-full border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden p-6 space-y-4">
            <div className="flex items-start gap-3">
              <div className="w-10 h-10 rounded-xl bg-amber-100 dark:bg-amber-950/60 text-amber-600 flex items-center justify-center shrink-0">
                <AlertTriangle className="w-5 h-5" />
              </div>
              <div className="flex-1">
                <h3 className="text-sm font-bold text-slate-900 dark:text-slate-100">
                  Move {revisionStats.deleted} deleted valves to Next Rev?
                </h3>
                <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                  According to <strong>Revision Principle Rule 3</strong>: Keeps SFI &amp; TAG, clearing all other technical specification fields.
                </p>
              </div>
            </div>
            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setIsNextRevModalOpen(false)}
                className="px-3.5 py-2 rounded-lg text-xs font-semibold text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors cursor-pointer"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleExecuteNextRevForDeleted}
                className="inline-flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-bold bg-amber-500 hover:bg-amber-600 text-slate-950 shadow-sm transition-colors cursor-pointer"
              >
                <span>Confirm Move</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
