import React, { useState } from 'react';
import { X, Plus, Trash2, RotateCcw, SlidersHorizontal, Check } from 'lucide-react';
import { DEFAULT_HEADER_ALIASES, normalizeHeader } from '../utils/headerNormalizer';

interface AliasConfigModalProps {
  isOpen: boolean;
  onClose: () => void;
  aliases: Record<string, string[]>;
  setAliases: (aliases: Record<string, string[]>) => void;
  masterHeaders: string[];
}

export const AliasConfigModal: React.FC<AliasConfigModalProps> = ({
  isOpen,
  onClose,
  aliases,
  setAliases,
  masterHeaders,
}) => {
  const [selectedHeader, setSelectedHeader] = useState<string>(masterHeaders[0] || 'TAG');
  const [newAliasInput, setNewAliasInput] = useState('');

  if (!isOpen) return null;

  const normSelected = normalizeHeader(selectedHeader);
  const currentAliases = aliases[normSelected] || [];

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    const clean = newAliasInput.trim().toUpperCase();
    if (!clean) return;

    if (currentAliases.includes(clean)) {
      alert('Bí danh này đã có trong danh sách!');
      return;
    }

    setAliases({
      ...aliases,
      [normSelected]: [...currentAliases, clean],
    });
    setNewAliasInput('');
  };

  const handleRemove = (aliasToRemove: string) => {
    setAliases({
      ...aliases,
      [normSelected]: currentAliases.filter((a) => a !== aliasToRemove),
    });
  };

  const handleReset = () => {
    if (window.confirm('Khôi phục danh sách bí danh mặc định từ macro VBA?')) {
      setAliases({ ...DEFAULT_HEADER_ALIASES });
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-xl w-full border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden flex flex-col max-h-[85vh]">
        {/* Header */}
        <div className="px-6 py-4 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <SlidersHorizontal className="w-5 h-5 text-blue-500" />
            <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
              Quản lý Bí danh Cột (Column Aliases)
            </h2>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Body */}
        <div className="p-6 overflow-y-auto space-y-4">
          <p className="text-xs text-slate-500 dark:text-slate-400">
            Khi tên cột trong file nhà cung cấp khác với tên chuẩn trong Master, hệ thống sẽ căn cứ vào danh sách bí danh này để tự động ghép cột.
          </p>

          {/* Select master header */}
          <div>
            <label className="block text-xs font-semibold text-slate-700 dark:text-slate-300 mb-1.5">
              Chọn cột Master cần cấu hình:
            </label>
            <select
              value={selectedHeader}
              onChange={(e) => setSelectedHeader(e.target.value)}
              className="w-full px-3 py-2 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-xs font-medium focus:ring-2 focus:ring-blue-500 focus:outline-none"
            >
              {masterHeaders.map((h) => (
                <option key={h} value={h}>
                  {h} ({(aliases[normalizeHeader(h)] || []).length} alias)
                </option>
              ))}
            </select>
          </div>

          {/* Existing Aliases */}
          <div className="space-y-2">
            <span className="text-xs font-medium text-slate-700 dark:text-slate-300">
              Các bí danh được công nhận cho <strong>{selectedHeader}</strong>:
            </span>
            <div className="min-h-[100px] max-h-[160px] overflow-y-auto p-3 rounded-lg border border-slate-200 dark:border-slate-800 bg-slate-50/50 dark:bg-slate-950/50 flex flex-wrap gap-1.5 content-start">
              {currentAliases.length > 0 ? (
                currentAliases.map((al) => (
                  <span
                    key={al}
                    className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-md text-xs font-medium bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-200 border border-slate-200 dark:border-slate-700 shadow-xs"
                  >
                    <span>{al}</span>
                    <button
                      type="button"
                      onClick={() => handleRemove(al)}
                      className="text-slate-400 hover:text-rose-500 text-sm leading-none"
                    >
                      &times;
                    </button>
                  </span>
                ))
              ) : (
                <span className="text-xs text-slate-400 italic">
                  Chưa có bí danh nào cho cột này.
                </span>
              )}
            </div>
          </div>

          {/* Add new alias input */}
          <form onSubmit={handleAdd} className="flex items-center gap-2 pt-2">
            <input
              type="text"
              value={newAliasInput}
              onChange={(e) => setNewAliasInput(e.target.value)}
              placeholder="Thêm bí danh mới (vd: PO NO, ACTUATOR...)"
              className="flex-1 px-3 py-2 text-xs rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-slate-900 dark:text-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
            <button
              type="submit"
              className="inline-flex items-center gap-1 px-4 py-2 text-xs font-semibold rounded-lg bg-blue-600 hover:bg-blue-500 text-white shadow-sm transition-colors"
            >
              <Plus className="w-3.5 h-3.5" />
              Thêm
            </button>
          </form>
        </div>

        {/* Footer */}
        <div className="px-6 py-3.5 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 flex items-center justify-between">
          <button
            type="button"
            onClick={handleReset}
            className="inline-flex items-center gap-1.5 text-xs text-slate-500 hover:text-slate-700 dark:hover:text-slate-300"
          >
            <RotateCcw className="w-3.5 h-3.5" />
            Khôi phục mặc định VBA
          </button>
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-white transition-colors"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
};
