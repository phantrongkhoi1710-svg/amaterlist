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
  Eye,
  ChevronLeft,
  ChevronRight,
  Sparkles,
} from 'lucide-react';
import { MasterRowData } from '../types';

interface MasterTableProps {
  data: MasterRowData[];
  masterHeaders: string[];
  onDeleteRow: (rowId: string) => void;
  onExportExcel: () => void;
  onExportCSV: () => void;
}

export const MasterTable: React.FC<MasterTableProps> = ({
  data,
  masterHeaders,
  onDeleteRow,
  onExportExcel,
  onExportCSV,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedFileFilter, setSelectedFileFilter] = useState<string>('all');
  const [sortColumn, setSortColumn] = useState<string | null>(null);
  const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(25);

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
      // Filter by file
      if (selectedFileFilter !== 'all' && row._sourceFile !== selectedFileFilter) {
        return false;
      }

      // Filter by search term
      if (!searchTerm) return true;
      const term = searchTerm.toLowerCase();

      for (const h of masterHeaders) {
        const val = row[h];
        if (val && String(val).toLowerCase().includes(term)) {
          return true;
        }
      }
      return false;
    });
  }, [data, selectedFileFilter, searchTerm, masterHeaders]);

  // Sort data
  const sortedData = useMemo(() => {
    if (!sortColumn) return filteredData;

    return [...filteredData].sort((a, b) => {
      const valA = a[sortColumn] !== undefined && a[sortColumn] !== null ? String(a[sortColumn]) : '';
      const valB = b[sortColumn] !== undefined && b[sortColumn] !== null ? String(b[sortColumn]) : '';

      const comparison = valA.localeCompare(valB, undefined, { numeric: true, sensitivity: 'base' });
      return sortDirection === 'asc' ? comparison : -comparison;
    });
  }, [filteredData, sortColumn, sortDirection]);

  // Pagination
  const totalPages = Math.max(1, Math.ceil(sortedData.length / pageSize));
  const paginatedData = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return sortedData.slice(start, start + pageSize);
  }, [sortedData, currentPage, pageSize]);

  const handleSort = (colName: string) => {
    if (sortColumn === colName) {
      if (sortDirection === 'asc') {
        setSortDirection('desc');
      } else {
        setSortColumn(null);
      }
    } else {
      setSortColumn(colName);
      setSortDirection('asc');
    }
  };

  if (data.length === 0) {
    return (
      <div className="bg-white dark:bg-slate-900 rounded-2xl border border-slate-200 dark:border-slate-800 p-12 text-center shadow-sm">
        <div className="w-16 h-16 rounded-2xl bg-blue-50 dark:bg-blue-950/40 text-blue-500 mx-auto flex items-center justify-center mb-4">
          <FileSpreadsheet className="w-8 h-8" />
        </div>
        <h3 className="text-base font-semibold text-slate-900 dark:text-slate-100">
          Chưa có dữ liệu trong bảng Master (Total)
        </h3>
        <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 max-w-md mx-auto">
          Hãy kéo thả các file Excel vào tab &quot;Quản lý File & Import&quot; hoặc nhấn &quot;Nạp File Mẫu Test&quot; ở góc trên để trải nghiệm ngay.
        </p>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* Controls Bar */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 p-4 shadow-sm">
        <div className="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-3">
          {/* Search & File Filter */}
          <div className="flex flex-wrap items-center gap-3 flex-1">
            {/* Search Input */}
            <div className="relative flex-1 min-w-[220px] max-w-md">
              <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
              <input
                id="table-search-input"
                type="text"
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                placeholder="Tìm kiếm theo TAG, SFI, Nhà cung cấp..."
                className="w-full pl-9 pr-4 py-2 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-xs focus:ring-2 focus:ring-blue-500 focus:outline-none dark:text-slate-100"
              />
            </div>

            {/* Filter by file */}
            {sourceFiles.length > 1 && (
              <div className="flex items-center gap-1.5 text-xs text-slate-600 dark:text-slate-300">
                <Filter className="w-3.5 h-3.5 text-slate-400" />
                <select
                  id="select-file-filter"
                  value={selectedFileFilter}
                  onChange={(e) => {
                    setSelectedFileFilter(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="px-2.5 py-2 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-xs focus:ring-2 focus:ring-blue-500 focus:outline-none"
                >
                  <option value="all">Tất cả file nguồn ({sourceFiles.length})</option>
                  {sourceFiles.map((fn) => (
                    <option key={fn} value={fn}>
                      {fn}
                    </option>
                  ))}
                </select>
              </div>
            )}
          </div>

          {/* Export & Page Size */}
          <div className="flex items-center gap-2 justify-end">
            <button
              id="btn-table-export-excel"
              onClick={onExportExcel}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-emerald-600 hover:bg-emerald-500 text-white shadow-sm transition-colors cursor-pointer"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Tải Excel</span>
            </button>
            <button
              id="btn-table-export-csv"
              onClick={onExportCSV}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-100 dark:bg-slate-800 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 border border-slate-200 dark:border-slate-700 transition-colors"
            >
              <span>CSV</span>
            </button>
          </div>
        </div>
      </div>

      {/* Color Legend Bar */}
      <div className="flex flex-wrap items-center gap-3 px-2 text-[11px] text-slate-500 dark:text-slate-400">
        <span className="font-semibold text-slate-700 dark:text-slate-300">Phân nhóm cột:</span>
        <span className="inline-flex items-center gap-1">
          <span className="w-2.5 h-2.5 rounded-full bg-blue-600"></span>
          <span>Định danh cốt lõi (TAG, SUPPLIER, SFI, DESC)</span>
        </span>
        <span className="inline-flex items-center gap-1">
          <span className="w-2.5 h-2.5 rounded-full bg-teal-600"></span>
          <span>Truy xuất đơn hàng / Actuator</span>
        </span>
        <span className="inline-flex items-center gap-1">
          <span className="w-2.5 h-2.5 rounded-full bg-slate-600"></span>
          <span>Thông số kỹ thuật van</span>
        </span>
      </div>

      {/* Table Container */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
        <div className="overflow-x-auto max-h-[600px]">
          <table className="w-full text-left text-xs text-slate-700 dark:text-slate-300 border-collapse">
            <thead className="bg-slate-100 dark:bg-slate-800/95 text-slate-800 dark:text-slate-200 sticky top-0 z-20 border-b border-slate-200 dark:border-slate-700 font-semibold shadow-xs">
              <tr>
                <th className="py-3 px-3 w-10 text-center text-slate-400 font-mono">#</th>
                <th className="py-3 px-3 text-slate-500 font-mono text-[11px] whitespace-nowrap">File Nguồn</th>
                {masterHeaders.map((header) => {
                  const isSorted = sortColumn === header;
                  const isKeyField = ['TAG', 'SUPPLIER', 'SFI', 'DESCRIPTION'].includes(header);
                  const isActuator = ['ACTUATOR TAG', 'PO NUMBER', 'DESTINATION YARD'].includes(header);

                  return (
                    <th
                      key={header}
                      onClick={() => handleSort(header)}
                      className="py-3 px-3.5 cursor-pointer hover:bg-slate-200/60 dark:hover:bg-slate-700/60 transition-colors whitespace-nowrap select-none"
                    >
                      <div className="flex items-center gap-1.5">
                        <span
                          className={`w-1.5 h-3.5 rounded-full mr-0.5 ${
                            isKeyField
                              ? 'bg-blue-600'
                              : isActuator
                              ? 'bg-teal-600'
                              : 'bg-slate-500'
                          }`}
                        ></span>
                        <span className={isKeyField ? 'font-bold text-blue-900 dark:text-blue-300' : ''}>
                          {header}
                        </span>
                        {isSorted ? (
                          sortDirection === 'asc' ? (
                            <ArrowUp className="w-3 h-3 text-blue-500" />
                          ) : (
                            <ArrowDown className="w-3 h-3 text-blue-500" />
                          )
                        ) : (
                          <ArrowUpDown className="w-3 h-3 text-slate-400 opacity-40 hover:opacity-100" />
                        )}
                      </div>
                    </th>
                  );
                })}
                <th className="py-3 px-3 text-center w-12">Xoá</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-800 font-normal">
              {paginatedData.map((row, idx) => {
                const rowIndex = (currentPage - 1) * pageSize + idx + 1;
                return (
                  <tr
                    key={row._id}
                    className="hover:bg-blue-50/40 dark:hover:bg-blue-950/20 transition-colors"
                  >
                    <td className="py-2.5 px-3 text-center text-slate-400 font-mono text-[11px]">
                      {rowIndex}
                    </td>
                    <td className="py-2.5 px-3 text-slate-500 text-[11px] font-mono whitespace-nowrap">
                      <span className="truncate max-w-[140px] block" title={row._sourceFile}>
                        {row._sourceFile}
                      </span>
                    </td>
                    {masterHeaders.map((header) => {
                      const val = row[header];
                      const isEmpty = val === undefined || val === null || String(val).trim() === '';

                      // Styling for important keys like TAG or SUPPLIER
                      const isTag = header === 'TAG';
                      const isSupplier = header === 'SUPPLIER';

                      return (
                        <td
                          key={header}
                          className={`py-2.5 px-3.5 whitespace-nowrap text-xs ${
                            isEmpty
                              ? 'text-slate-300 dark:text-slate-600 italic'
                              : isTag
                              ? 'font-bold text-blue-600 dark:text-blue-400'
                              : isSupplier
                              ? 'font-medium text-emerald-700 dark:text-emerald-400'
                              : 'text-slate-800 dark:text-slate-200'
                          }`}
                        >
                          {isEmpty ? '—' : String(val)}
                        </td>
                      );
                    })}
                    <td className="py-2.5 px-3 text-center">
                      <button
                        onClick={() => onDeleteRow(row._id)}
                        className="text-slate-400 hover:text-rose-500 transition-colors p-1"
                        title="Xoá hàng này"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
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
              className="p-1.5 rounded-md border border-slate-200 dark:border-slate-700 disabled:opacity-40 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <span className="px-2 font-medium">
              Trang {currentPage} / {totalPages}
            </span>
            <button
              onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
              disabled={currentPage === totalPages}
              className="p-1.5 rounded-md border border-slate-200 dark:border-slate-700 disabled:opacity-40 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
