import React, { useState, useMemo } from 'react';
import {
  Clock,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  FileSpreadsheet,
  Download,
  Trash2,
  Search,
} from 'lucide-react';
import { ImportLogRow } from '../types';

interface ImportLogViewProps {
  logs: ImportLogRow[];
  onClearLogs: () => void;
  onExportExcel: () => void;
}

export const ImportLogView: React.FC<ImportLogViewProps> = ({
  logs,
  onClearLogs,
  onExportExcel,
}) => {
  const [filterStatus, setFilterStatus] = useState<'all' | 'ok' | 'warning' | 'error'>('all');
  const [searchTerm, setSearchTerm] = useState('');

  const filteredLogs = useMemo(() => {
    return logs.filter((log) => {
      // Status filter
      if (filterStatus === 'ok' && !log.status.startsWith('OK') && log.status !== 'OK') {
        return false;
      }
      if (filterStatus === 'warning' && !log.status.includes('WARNING') && !log.status.includes('MISSING')) {
        return false;
      }
      if (filterStatus === 'error' && !log.status.includes('ERROR')) {
        return false;
      }

      // Search term
      if (!searchTerm) return true;
      const term = searchTerm.toLowerCase();
      return (
        log.sourceFile.toLowerCase().includes(term) ||
        log.sourceSheet.toLowerCase().includes(term) ||
        log.status.toLowerCase().includes(term)
      );
    });
  }, [logs, filterStatus, searchTerm]);

  return (
    <div className="space-y-4">
      {/* Top action bar */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 p-4 shadow-sm">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <Clock className="w-5 h-5 text-blue-500" />
            <div>
              <h3 className="text-sm font-semibold text-slate-800 dark:text-slate-100">
                Nhật ký Quá trình Import (Sheet: Import_Log)
              </h3>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Tự động ghi lại kết quả xử lý từng file nguồn, số dòng nhập và cảnh báo lỗi
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={onExportExcel}
              disabled={logs.length === 0}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-emerald-600 hover:bg-emerald-500 text-white disabled:opacity-50 transition-colors"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Xuất Excel</span>
            </button>
            {logs.length > 0 && (
              <button
                onClick={onClearLogs}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium text-rose-500 hover:text-rose-600 hover:bg-rose-50 dark:hover:bg-rose-950/40 transition-colors"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>Xoá Log</span>
              </button>
            )}
          </div>
        </div>

        {/* Filter bar */}
        <div className="mt-3 pt-3 border-t border-slate-100 dark:border-slate-800 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-1.5">
            <button
              onClick={() => setFilterStatus('all')}
              className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors ${
                filterStatus === 'all'
                  ? 'bg-slate-800 text-white dark:bg-slate-200 dark:text-slate-900'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400'
              }`}
            >
              Tất cả ({logs.length})
            </button>
            <button
              onClick={() => setFilterStatus('ok')}
              className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors ${
                filterStatus === 'ok'
                  ? 'bg-emerald-600 text-white'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400'
              }`}
            >
              Thành công
            </button>
            <button
              onClick={() => setFilterStatus('warning')}
              className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors ${
                filterStatus === 'warning'
                  ? 'bg-amber-600 text-white'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400'
              }`}
            >
              Cảnh báo / Thiếu cột
            </button>
            <button
              onClick={() => setFilterStatus('error')}
              className={`px-2.5 py-1 rounded-md text-xs font-medium transition-colors ${
                filterStatus === 'error'
                  ? 'bg-rose-600 text-white'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400'
              }`}
            >
              Lỗi
            </button>
          </div>

          <div className="relative min-w-[200px]">
            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder="Lọc nhật ký..."
              className="w-full pl-8 pr-3 py-1 rounded-md border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-xs focus:ring-1 focus:ring-blue-500 focus:outline-none"
            />
          </div>
        </div>
      </div>

      {/* Log Table */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden">
        {logs.length === 0 ? (
          <div className="p-10 text-center text-slate-400 text-xs">
            Chưa có nhật ký nào. Nhật ký sẽ tự động sinh khi bạn tải file nguồn lên.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs text-slate-700 dark:text-slate-300 border-collapse">
              <thead className="bg-slate-100 dark:bg-slate-800 text-slate-800 dark:text-slate-200 font-semibold border-b border-slate-200 dark:border-slate-700">
                <tr>
                  <th className="py-2.5 px-3 whitespace-nowrap">TIME</th>
                  <th className="py-2.5 px-3 whitespace-nowrap">SOURCE FILE</th>
                  <th className="py-2.5 px-3 whitespace-nowrap">SOURCE SHEET</th>
                  <th className="py-2.5 px-3 text-center whitespace-nowrap">HEADER ROW</th>
                  <th className="py-2.5 px-3 text-center whitespace-nowrap">SOURCE ROWS</th>
                  <th className="py-2.5 px-3 text-center whitespace-nowrap">IMPORTED ROWS</th>
                  <th className="py-2.5 px-3 whitespace-nowrap">STATUS</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
                {filteredLogs.map((item) => {
                  const isSuccess = item.status === 'OK';
                  const isWarning = item.status.includes('WARNING') || item.status.includes('MISSING');
                  const isError = item.status.includes('ERROR');

                  return (
                    <tr key={item.id} className="hover:bg-slate-50/50 dark:hover:bg-slate-800/40">
                      <td className="py-2 px-3 font-mono text-[11px] text-slate-500 whitespace-nowrap">
                        {item.time}
                      </td>
                      <td className="py-2 px-3 font-medium text-slate-900 dark:text-slate-100 whitespace-nowrap">
                        {item.sourceFile}
                      </td>
                      <td className="py-2 px-3 text-slate-600 dark:text-slate-300 whitespace-nowrap">
                        {item.sourceSheet || '—'}
                      </td>
                      <td className="py-2 px-3 text-center font-mono">
                        {item.headerRow > 0 ? item.headerRow : '—'}
                      </td>
                      <td className="py-2 px-3 text-center font-mono">
                        {item.sourceRows}
                      </td>
                      <td className="py-2 px-3 text-center font-bold text-emerald-600 dark:text-emerald-400 font-mono">
                        {item.importedRows}
                      </td>
                      <td className="py-2 px-3">
                        <div className="flex items-center gap-1.5">
                          {isSuccess && <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500 flex-shrink-0" />}
                          {isWarning && <AlertTriangle className="w-3.5 h-3.5 text-amber-500 flex-shrink-0" />}
                          {isError && <XCircle className="w-3.5 h-3.5 text-rose-500 flex-shrink-0" />}
                          <span
                            className={`text-xs ${
                              isSuccess
                                ? 'text-emerald-700 dark:text-emerald-300 font-semibold'
                                : isWarning
                                ? 'text-amber-700 dark:text-amber-300 font-medium'
                                : isError
                                ? 'text-rose-700 dark:text-rose-300 font-medium'
                                : 'text-slate-600'
                            }`}
                          >
                            {item.status}
                          </span>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
};
