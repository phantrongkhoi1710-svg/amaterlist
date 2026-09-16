import React, { useState } from 'react';
import {
  Layers,
  Plus,
  Trash2,
  CheckCircle2,
  AlertCircle,
  HelpCircle,
  RotateCcw,
} from 'lucide-react';
import { DEFAULT_MASTER_COLUMNS, DEFAULT_HEADER_ALIASES, normalizeHeader } from '../utils/headerNormalizer';

interface MappingViewerProps {
  masterHeaders: string[];
  setMasterHeaders: (headers: string[]) => void;
  aliases: Record<string, string[]>;
  setAliases: (aliases: Record<string, string[]>) => void;
}

export const MappingViewer: React.FC<MappingViewerProps> = ({
  masterHeaders,
  setMasterHeaders,
  aliases,
  setAliases,
}) => {
  const [newColumnName, setNewColumnName] = useState('');
  const [selectedHeaderForAlias, setSelectedHeaderForAlias] = useState<string | null>(null);
  const [newAliasInput, setNewAliasInput] = useState('');

  const handleAddMasterColumn = (e: React.FormEvent) => {
    e.preventDefault();
    const clean = newColumnName.trim().toUpperCase();
    if (!clean) return;
    if (masterHeaders.includes(clean)) {
      alert('Cột này đã tồn tại trong danh sách Master!');
      return;
    }
    setMasterHeaders([...masterHeaders, clean]);
    setNewColumnName('');
  };

  const handleRemoveMasterColumn = (headerName: string) => {
    if (masterHeaders.length <= 1) {
      alert('Bảng Master cần có ít nhất 1 cột!');
      return;
    }
    setMasterHeaders(masterHeaders.filter((h) => h !== headerName));
  };

  const handleResetToDefaults = () => {
    if (window.confirm('Bạn có chắc muốn khôi phục về danh sách cột và Alias mặc định của macro VBA?')) {
      setMasterHeaders([...DEFAULT_MASTER_COLUMNS]);
      setAliases({ ...DEFAULT_HEADER_ALIASES });
    }
  };

  const handleAddAlias = (masterHeader: string) => {
    const clean = newAliasInput.trim().toUpperCase();
    if (!clean) return;

    const norm = normalizeHeader(masterHeader);
    const existing = aliases[norm] || [];
    if (existing.includes(clean)) {
      alert('Alias này đã tồn tại!');
      return;
    }

    setAliases({
      ...aliases,
      [norm]: [...existing, clean],
    });
    setNewAliasInput('');
  };

  const handleRemoveAlias = (masterHeader: string, aliasToRemove: string) => {
    const norm = normalizeHeader(masterHeader);
    const existing = aliases[norm] || [];
    setAliases({
      ...aliases,
      [norm]: existing.filter((a) => a !== aliasToRemove),
    });
  };

  return (
    <div className="space-y-6">
      {/* Introduction Card */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 p-5 shadow-sm">
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div>
            <h3 className="text-sm font-semibold text-slate-900 dark:text-slate-100 flex items-center gap-2">
              <Layers className="w-4 h-4 text-blue-500" />
              Cấu trúc Cột Bảng Master & Quy tắc Bí danh (Aliases)
            </h3>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
              Hệ thống tự động chuẩn hoá tên cột (viết hoa, xoá dấu gạch/chấm/khoảng trắng thừa) và đối chiếu với danh sách bí danh được khai báo từ macro VBA.
            </p>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={handleResetToDefaults}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 transition-colors"
            >
              <RotateCcw className="w-3.5 h-3.5" />
              Khôi phục mặc định VBA
            </button>
          </div>
        </div>

        {/* Add Master Column Form */}
        <form onSubmit={handleAddMasterColumn} className="mt-4 pt-4 border-t border-slate-100 dark:border-slate-800 flex items-center gap-2">
          <input
            type="text"
            value={newColumnName}
            onChange={(e) => setNewColumnName(e.target.value)}
            placeholder="Thêm tên cột Master mới (vd: VALVE TYPE, REMARKS, WEIGHT...)"
            className="flex-1 max-w-md px-3 py-1.5 text-xs rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-slate-900 dark:text-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          <button
            type="submit"
            className="inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-medium bg-blue-600 hover:bg-blue-500 text-white transition-colors"
          >
            <Plus className="w-3.5 h-3.5" />
            Thêm Cột
          </button>
        </form>
      </div>

      {/* Columns & Aliases Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {masterHeaders.map((header, index) => {
          const norm = normalizeHeader(header);
          const headerAliases = aliases[norm] || [];
          const isCoreRequired = ['TAG', 'SUPPLIER', 'SFI', 'DESCRIPTION'].includes(norm);

          return (
            <div
              key={header}
              className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 p-4 shadow-sm flex flex-col justify-between"
            >
              <div>
                <div className="flex items-start justify-between gap-2 mb-2">
                  <div className="flex items-center gap-2">
                    <span className="w-5 h-5 rounded-full bg-slate-100 dark:bg-slate-800 text-[11px] font-mono text-slate-500 flex items-center justify-center">
                      {index + 1}
                    </span>
                    <h4 className="text-xs font-bold text-slate-900 dark:text-slate-100 tracking-wide">
                      {header}
                    </h4>
                  </div>

                  <div className="flex items-center gap-1">
                    {isCoreRequired && (
                      <span className="text-[10px] px-1.5 py-0.5 rounded font-medium bg-blue-50 text-blue-700 dark:bg-blue-950/60 dark:text-blue-300 border border-blue-200 dark:border-blue-800">
                        Bắt buộc
                      </span>
                    )}
                    <button
                      onClick={() => handleRemoveMasterColumn(header)}
                      className="p-1 text-slate-300 hover:text-rose-500 transition-colors"
                      title="Xoá cột này khỏi Master"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                </div>

                {/* Aliases List */}
                <div className="mt-2 text-xs">
                  <span className="text-[11px] text-slate-400 block mb-1">
                    Các tên gọi khác (Aliases):
                  </span>
                  {headerAliases.length > 0 ? (
                    <div className="flex flex-wrap gap-1 mb-2">
                      {headerAliases.map((al) => (
                        <span
                          key={al}
                          className="inline-flex items-center gap-1 text-[11px] px-2 py-0.5 rounded bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border border-slate-200 dark:border-slate-700"
                        >
                          <span>{al}</span>
                          <button
                            type="button"
                            onClick={() => handleRemoveAlias(header, al)}
                            className="text-slate-400 hover:text-rose-500 ml-0.5"
                          >
                            &times;
                          </button>
                        </span>
                      ))}
                    </div>
                  ) : (
                    <p className="text-[11px] text-slate-400 italic mb-2">Chưa có alias nào (chỉ khớp chính xác)</p>
                  )}
                </div>
              </div>

              {/* Add alias input */}
              <div className="mt-3 pt-2.5 border-t border-slate-100 dark:border-slate-800/80">
                {selectedHeaderForAlias === header ? (
                  <div className="flex items-center gap-1">
                    <input
                      type="text"
                      value={newAliasInput}
                      onChange={(e) => setNewAliasInput(e.target.value)}
                      placeholder="Nhập tên alias..."
                      className="flex-1 px-2 py-1 text-xs rounded border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-slate-900 dark:text-slate-100 focus:outline-none focus:ring-1 focus:ring-blue-500"
                      onKeyDown={(e) => {
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          handleAddAlias(header);
                        }
                      }}
                      autoFocus
                    />
                    <button
                      type="button"
                      onClick={() => handleAddAlias(header)}
                      className="px-2 py-1 text-xs rounded bg-blue-600 text-white font-medium hover:bg-blue-500"
                    >
                      Lưu
                    </button>
                    <button
                      type="button"
                      onClick={() => setSelectedHeaderForAlias(null)}
                      className="px-1.5 py-1 text-xs text-slate-400 hover:text-slate-600"
                    >
                      Huỷ
                    </button>
                  </div>
                ) : (
                  <button
                    type="button"
                    onClick={() => {
                      setSelectedHeaderForAlias(header);
                      setNewAliasInput('');
                    }}
                    className="text-[11px] text-blue-600 dark:text-blue-400 hover:underline flex items-center gap-1"
                  >
                    <Plus className="w-3 h-3" /> Thêm alias cho cột này
                  </button>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};
