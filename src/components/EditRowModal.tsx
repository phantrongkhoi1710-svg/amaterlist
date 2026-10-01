import React, { useState, useEffect } from 'react';
import {
  X,
  Save,
  Mail,
  Edit3,
  AlertCircle,
  Tag,
  Settings,
  Layers,
  FileCheck2,
  Trash2,
  Palette,
} from 'lucide-react';
import { MasterRowData, RowRevisionStatus } from '../types';
import { RowEditChange } from '../utils/outlookMailer';
import { applyRowRevisionStatus } from '../utils/revisionManager';

interface EditRowModalProps {
  isOpen: boolean;
  onClose: () => void;
  row: MasterRowData | null;
  masterHeaders: string[];
  onSave: (updatedRow: MasterRowData, changes: RowEditChange[], autoComposeMail: boolean) => void;
  onDeleteRow?: (rowId: string) => void;
}

export const EditRowModal: React.FC<EditRowModalProps> = ({
  isOpen,
  onClose,
  row,
  masterHeaders,
  onSave,
  onDeleteRow,
}) => {
  const [formData, setFormData] = useState<Record<string, any>>({});
  const [revisionStatus, setRevisionStatus] = useState<RowRevisionStatus>('normal');
  const [customColor, setCustomColor] = useState<string | undefined>(undefined);
  const [hasStrikethrough, setHasStrikethrough] = useState<boolean>(false);
  const [autoEmail, setAutoEmail] = useState(true);
  const [showDeleteConfirm, setShowDeleteConfirm] = useState(false);

  useEffect(() => {
    if (row) {
      // Clone row data
      const data: Record<string, any> = {};
      masterHeaders.forEach((h) => {
        data[h] = row[h] ?? '';
      });
      setFormData(data);
      setRevisionStatus(row._revisionStatus || 'normal');
      setCustomColor(row._detectedColor);
      setHasStrikethrough(row._hasStrikethrough ?? (row._revisionStatus === 'deleted'));
      setAutoEmail(true);
      setShowDeleteConfirm(false);
    }
  }, [row, masterHeaders]);

  if (!isOpen || !row) return null;

  const handleChange = (header: string, value: string) => {
    setFormData((prev) => ({
      ...prev,
      [header]: value,
    }));
  };

  // Helper to clear other cells for Next Rev (Keep SFI & TAG)
  const handleApplyNextRevCleanup = () => {
    setFormData((prev) => {
      const updated = { ...prev };
      const preserved = new Set(['TAG', 'SFI', 'REV. HIS.', 'DATE', 'REV. DESCRIPTION', 'SIGNATURE']);
      Object.keys(updated).forEach((k) => {
        if (!preserved.has(k)) {
          updated[k] = '';
        }
      });
      if (!updated['REV. DESCRIPTION']) {
        updated['REV. DESCRIPTION'] = 'Deleted in previous revision';
      }
      return updated;
    });
    setRevisionStatus('next_rev_after_deleted');
    setHasStrikethrough(false);
  };

  // Calculate detected changes
  const changes: RowEditChange[] = [];
  masterHeaders.forEach((h) => {
    const oldVal = row[h] ?? '';
    const newVal = formData[h] ?? '';
    if (String(oldVal).trim() !== String(newVal).trim()) {
      changes.push({
        field: h,
        oldValue: oldVal,
        newValue: newVal,
      });
    }
  });

  const handleSubmit = (withMail: boolean) => {
    let updatedRow: MasterRowData = {
      ...row,
      ...formData,
      _revisionStatus: revisionStatus,
      _detectedColor:
        customColor ||
        (revisionStatus === 'new_updated'
          ? '#FFFF00'
          : revisionStatus === 'deleted'
          ? '#EF4444'
          : undefined),
      _hasStrikethrough: hasStrikethrough,
    };
    onSave(updatedRow, changes, withMail);
  };

  // Group 1 (Yellow): TO BE COMPLETED BY FUNCTIONAL DESIGN
  const functionalDesignCols = [
    'TAG',
    'ACTUATOR TAG',
    'P.O. NUMBER',
    'SUPPLIER',
    'DESTINATION (YARD)',
  ];

  // Group 2 (Gray): ARMATURE INFO FROM PIPE SPECIFICATION
  const pipeSpecCols = [
    'SAP CODE',
    'STD DRW NORMALE N°',
    'EXECUTION',
    'NRF N°',
    'SFI',
    'DESCRIPTION',
    'SIZE',
    'CONNECTION',
    'PRESSURE RATING',
    'HOUSING /BODY',
    'TYPE',
    'PIPE CLASS',
    'CLASS CERTIFICATE',
    'REMARKS',
    'SIGN TYPE',
    'SIGN TEXT',
  ];

  // Group 3 (White/Clear): REVISION & SIGNATURE
  const revisionCols = [
    'INPUT - SIGN TEXT',
    'REV. HIS.',
    'DATE',
    'REV. DESCRIPTION',
    'SIGNATURE',
  ];

  // Find remaining custom headers if any
  const remainingHeaders = masterHeaders.filter(
    (h) => !functionalDesignCols.includes(h) && !pipeSpecCols.includes(h) && !revisionCols.includes(h)
  );

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-4xl w-full border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden flex flex-col max-h-[92vh]">
        {/* Header */}
        <div className="px-6 py-4 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between bg-slate-50 dark:bg-slate-850">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 rounded-lg bg-blue-600 text-white flex items-center justify-center">
              <Edit3 className="w-4 h-4" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
                  Hiệu Chỉnh Thông Số Van: {formData.TAG || row.TAG || 'Chưa đặt TAG'}
                </h2>
                <span className="text-[11px] px-2 py-0.5 rounded bg-blue-100 dark:bg-blue-900/40 text-blue-700 dark:text-blue-300 font-medium">
                  {row._sourceFile}
                </span>
              </div>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Change Banner if any modified fields */}
        {changes.length > 0 && (
          <div className="px-6 py-2 bg-amber-50 dark:bg-amber-950/40 border-b border-amber-200 dark:border-amber-900/50 flex items-center justify-between text-xs text-amber-800 dark:text-amber-300">
            <div className="flex items-center gap-2">
              <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
              <span>
                Phát hiện <strong>{changes.length}</strong> trường đã được thay đổi giá trị
              </span>
            </div>
            <span className="text-[11px] italic text-amber-700 dark:text-amber-400">
              Sẽ được tự động tổng hợp vào bản thảo Email Outlook & file Excel đính kèm
            </span>
          </div>
        )}

        {/* Form Body */}
        <div className="p-6 overflow-y-auto space-y-6 text-xs">
          {/* PRINCIPLE FOR REVISION DESCRIPTIONS SELECTOR */}
          <div className="p-4 rounded-xl border border-slate-300 dark:border-slate-700 bg-slate-100 dark:bg-slate-800/60 space-y-2">
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-slate-850 dark:text-slate-100 uppercase tracking-wide flex items-center gap-1.5">
                <Tag className="w-3.5 h-3.5 text-blue-600" />
                Quy Chuẩn Revision Description (Nguyên Tắc Tô Màu &amp; Xóa Hàng)
              </span>
              {revisionStatus === 'next_rev_after_deleted' && (
                <button
                  type="button"
                  onClick={handleApplyNextRevCleanup}
                  className="px-2.5 py-1 rounded bg-slate-200 hover:bg-slate-300 dark:bg-slate-700 dark:hover:bg-slate-600 text-[11px] font-semibold text-slate-800 dark:text-slate-200 transition-colors"
                >
                  🧹 Làm trống các ô (chỉ giữ SFI &amp; TAG)
                </button>
              )}
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-4 gap-2 pt-1">
              <button
                type="button"
                onClick={() => setRevisionStatus('normal')}
                className={`px-3 py-2 rounded-lg border text-left text-xs transition-all cursor-pointer ${
                  revisionStatus === 'normal'
                    ? 'border-blue-500 bg-white dark:bg-slate-700 ring-2 ring-blue-400 font-bold shadow-xs'
                    : 'border-slate-200 dark:border-slate-700 bg-white/70 dark:bg-slate-800 text-slate-600 dark:text-slate-400'
                }`}
              >
                <div className="font-semibold text-slate-900 dark:text-slate-100">Bình thường</div>
                <div className="text-[10px] text-slate-400">Không tô màu</div>
              </button>

              <button
                type="button"
                onClick={() => setRevisionStatus('new_updated')}
                className={`px-3 py-2 rounded-lg border text-left text-xs transition-all cursor-pointer ${
                  revisionStatus === 'new_updated'
                    ? 'border-yellow-500 bg-yellow-300 text-slate-950 ring-2 ring-yellow-400 font-bold shadow-xs'
                    : 'border-yellow-300 bg-yellow-50 dark:bg-yellow-950/30 text-yellow-900 dark:text-yellow-200'
                }`}
              >
                <div className="font-bold flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-yellow-400 inline-block border border-yellow-600"></span>
                  New / Updated
                </div>
                <div className="text-[10px] opacity-80">Yellow coloured row</div>
              </button>

              <button
                type="button"
                onClick={() => setRevisionStatus('deleted')}
                className={`px-3 py-2 rounded-lg border text-left text-xs transition-all cursor-pointer ${
                  revisionStatus === 'deleted'
                    ? 'border-rose-600 bg-rose-500 text-white ring-2 ring-rose-400 font-bold shadow-xs line-through'
                    : 'border-rose-300 bg-rose-50 dark:bg-rose-950/30 text-rose-900 dark:text-rose-200 line-through'
                }`}
              >
                <div className="font-bold flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-rose-500 inline-block border border-rose-700"></span>
                  Deleted valve
                </div>
                <div className="text-[10px] opacity-80 no-underline">Red with strikethrough</div>
              </button>

              <button
                type="button"
                onClick={() => {
                  setRevisionStatus('next_rev_after_deleted');
                  handleApplyNextRevCleanup();
                }}
                className={`px-3 py-2 rounded-lg border text-left text-xs transition-all cursor-pointer ${
                  revisionStatus === 'next_rev_after_deleted'
                    ? 'border-slate-600 bg-slate-700 text-white ring-2 ring-slate-500 font-bold shadow-xs'
                    : 'border-slate-300 bg-slate-50 dark:bg-slate-800 text-slate-700 dark:text-slate-300'
                }`}
              >
                <div className="font-bold flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-slate-400 inline-block"></span>
                  Next revision
                </div>
                <div className="text-[10px] opacity-80">Keep SFI &amp; TAG only</div>
              </button>
            </div>

            {/* Strikethrough & Color Customizer Controls */}
            <div className="pt-3 border-t border-slate-200 dark:border-slate-700/60 space-y-3">
              {/* Row Strikethrough Toggle */}
              <div className="flex flex-wrap items-center justify-between gap-2 bg-white dark:bg-slate-800 p-2.5 rounded-lg border border-slate-200 dark:border-slate-700">
                <div className="flex items-center gap-2">
                  <span className="w-6 h-6 rounded flex items-center justify-center bg-rose-100 text-rose-700 dark:bg-rose-950 dark:text-rose-300 font-serif font-black text-xs">
                    <s>S</s>
                  </span>
                  <div>
                    <div className="font-bold text-slate-900 dark:text-slate-100 text-xs">
                      Gạch Ngang Dòng (Strikethrough)
                    </div>
                    <div className="text-[10px] text-slate-500">
                      Áp dụng hiệu ứng gạch ngang lên thông số van (dùng cho van bị hủy/xóa)
                    </div>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={() => setHasStrikethrough(!hasStrikethrough)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all cursor-pointer flex items-center gap-1.5 ${
                    hasStrikethrough
                      ? 'bg-rose-600 text-white ring-2 ring-rose-400 shadow-xs'
                      : 'bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 hover:bg-slate-200'
                  }`}
                >
                  <span className="text-sm font-serif"><s>S</s></span>
                  <span>{hasStrikethrough ? 'Đang BẬT gạch ngang' : 'BẬT gạch ngang'}</span>
                </button>
              </div>

              {/* Custom Color Palette bar with Native Color Picker */}
              <div className="flex flex-wrap items-center justify-between gap-2 bg-white dark:bg-slate-800 p-2.5 rounded-lg border border-slate-200 dark:border-slate-700">
                <div className="flex items-center gap-2">
                  <Palette className="w-4 h-4 text-blue-500" />
                  <div>
                    <div className="font-bold text-slate-900 dark:text-slate-100 text-xs">
                      Màu Sắc Dòng (Line Color)
                    </div>
                    <div className="text-[10px] text-slate-500">
                      Chọn màu tô cho toàn bộ dòng van trong bảng Master và file Excel
                    </div>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-1.5">
                  {[
                    { label: 'Vàng', hex: '#FFFF00', bg: 'bg-yellow-300' },
                    { label: 'Đỏ', hex: '#EF4444', bg: 'bg-rose-500' },
                    { label: 'Xanh lá', hex: '#86EFAC', bg: 'bg-emerald-300' },
                    { label: 'Xanh lam', hex: '#93C5FD', bg: 'bg-sky-300' },
                    { label: 'Cam', hex: '#FDBA74', bg: 'bg-orange-300' },
                    { label: 'Tím', hex: '#D8B4FE', bg: 'bg-purple-300' },
                    { label: 'Xanh ngọc', hex: '#67E8F9', bg: 'bg-cyan-300' },
                    { label: 'Hồng', hex: '#F472B6', bg: 'bg-pink-300' },
                  ].map((c) => (
                    <button
                      key={c.hex}
                      type="button"
                      onClick={() => {
                        setCustomColor(c.hex);
                        if (c.hex === '#FFFF00') setRevisionStatus('new_updated');
                        else if (c.hex === '#EF4444') setRevisionStatus('deleted');
                      }}
                      className={`w-5 h-5 rounded-full ${c.bg} transition-transform hover:scale-110 border ${
                        customColor === c.hex || (!customColor && revisionStatus === 'new_updated' && c.hex === '#FFFF00') || (!customColor && revisionStatus === 'deleted' && c.hex === '#EF4444')
                          ? 'ring-2 ring-blue-600 ring-offset-1 scale-110'
                          : 'border-slate-400/40'
                      }`}
                      title={`Tô màu ${c.label} (${c.hex})`}
                    />
                  ))}

                  {/* Native HTML Color Picker */}
                  <label
                    className="flex items-center gap-1 px-2 py-0.5 rounded border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-700 hover:bg-slate-100 text-[10px] text-slate-700 dark:text-slate-200 cursor-pointer"
                    title="Tự chọn màu tùy ý"
                  >
                    <input
                      type="color"
                      value={customColor || '#FFFF00'}
                      onChange={(e) => setCustomColor(e.target.value)}
                      className="w-4 h-4 rounded cursor-pointer border-0 p-0"
                    />
                    <span>Tự chọn</span>
                  </label>

                  <button
                    type="button"
                    onClick={() => {
                      setCustomColor(undefined);
                      setRevisionStatus('normal');
                    }}
                    className="px-2 py-0.5 rounded text-[10px] text-slate-500 hover:text-slate-800 dark:hover:text-slate-200 border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 cursor-pointer"
                    title="Xóa màu (mặc định không màu)"
                  >
                    Mặc định
                  </button>
                </div>
              </div>

              {/* Live Preview Card */}
              <div
                className="p-3 rounded-xl border flex items-center justify-between gap-3 text-xs transition-all shadow-xs"
                style={{
                  backgroundColor: customColor
                    ? `${customColor}22`
                    : revisionStatus === 'new_updated'
                    ? '#FEF08A44'
                    : revisionStatus === 'deleted'
                    ? '#FECDD344'
                    : revisionStatus === 'next_rev_after_deleted'
                    ? '#F1F5F9'
                    : '#FFFFFF',
                  borderLeft: `6px solid ${
                    customColor ||
                    (revisionStatus === 'new_updated'
                      ? '#EAB308'
                      : revisionStatus === 'deleted'
                      ? '#EF4444'
                      : revisionStatus === 'next_rev_after_deleted'
                      ? '#64748B'
                      : '#CBD5E1')
                  }`,
                }}
              >
                <div className="flex items-center gap-3">
                  <span className="text-[10px] font-bold uppercase tracking-wider text-slate-400">
                    Xem trước dòng van:
                  </span>
                  <span
                    className={`font-black text-sm ${
                      hasStrikethrough
                        ? 'line-through text-rose-700 dark:text-rose-300'
                        : 'text-slate-900 dark:text-slate-100'
                    }`}
                  >
                    {formData.TAG || row.TAG || 'V-TAG'}
                  </span>
                  <span
                    className={`text-xs ${
                      hasStrikethrough
                        ? 'line-through text-slate-500'
                        : 'text-slate-700 dark:text-slate-300'
                    }`}
                  >
                    {formData.DESCRIPTION || row.DESCRIPTION || formData.SFI || 'Thông số van Armature'}
                  </span>
                </div>

                <div className="flex items-center gap-2">
                  {hasStrikethrough && (
                    <span className="px-2 py-0.5 rounded text-[10px] bg-rose-600 text-white font-bold line-through">
                      Gạch ngang
                    </span>
                  )}
                  {customColor && (
                    <span
                      className="px-2 py-0.5 rounded text-[10px] font-mono font-bold border border-black/20"
                      style={{ backgroundColor: customColor, color: '#000000' }}
                    >
                      {customColor}
                    </span>
                  )}
                </div>
              </div>
            </div>
          </div>

          {/* Group 1: TO BE COMPLETED BY FUNCTIONAL DESIGN (YELLOW) */}
          <div className="p-4 rounded-xl border border-yellow-300 dark:border-yellow-800/60 bg-yellow-50/40 dark:bg-yellow-950/20">
            <div className="flex items-center gap-2 text-xs font-black text-amber-950 dark:text-yellow-200 uppercase tracking-wider mb-3">
              <span className="w-2.5 h-2.5 rounded-full bg-yellow-400 inline-block"></span>
              <span>1. TO BE COMPLETED BY FUNCTIONAL DESIGN (Màu vàng)</span>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
              {functionalDesignCols.map((field) => {
                const isChanged = changes.some((c) => c.field === field);
                return (
                  <div key={field} className={field === 'DESTINATION (YARD)' ? 'sm:col-span-2' : ''}>
                    <label className="block font-semibold text-slate-800 dark:text-slate-200 mb-1 flex items-center justify-between">
                      <span>{field}</span>
                      {isChanged && (
                        <span className="text-[10px] text-amber-600 dark:text-amber-400 font-bold">
                          Đã sửa
                        </span>
                      )}
                    </label>
                    <input
                      type="text"
                      value={formData[field] ?? ''}
                      onChange={(e) => handleChange(field, e.target.value)}
                      className={`w-full px-3 py-2 rounded-lg border text-xs focus:ring-2 focus:outline-none transition-colors ${
                        isChanged
                          ? 'border-amber-400 bg-amber-50/50 dark:bg-amber-950/20 text-slate-900 dark:text-slate-100 focus:ring-amber-500'
                          : 'border-yellow-300/80 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 focus:ring-yellow-400'
                      }`}
                    />
                  </div>
                );
              })}
            </div>
          </div>

          {/* Group 2: ARMATURE INFO FROM PIPE SPECIFICATION (GRAY) */}
          <div className="p-4 rounded-xl border border-slate-300 dark:border-slate-700 bg-slate-50/60 dark:bg-slate-850/40">
            <div className="flex items-center gap-2 text-xs font-black text-slate-900 dark:text-slate-100 uppercase tracking-wider mb-3">
              <span className="w-2.5 h-2.5 rounded-full bg-slate-400 inline-block"></span>
              <span>2. ARMATURE INFO FROM PIPE SPECIFICATION (Màu xám)</span>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
              {pipeSpecCols.map((field) => {
                const isChanged = changes.some((c) => c.field === field);
                const isSap = field === 'SAP CODE';
                return (
                  <div key={field} className={field === 'DESCRIPTION' ? 'sm:col-span-2' : ''}>
                    <label className="block font-semibold text-slate-700 dark:text-slate-300 mb-1 flex items-center justify-between truncate" title={field}>
                      <span className="truncate">{field}</span>
                      {isChanged && (
                        <span className="text-[10px] text-amber-600 dark:text-amber-400 font-bold ml-1">
                          Đã sửa
                        </span>
                      )}
                    </label>
                    <input
                      type="text"
                      value={formData[field] ?? ''}
                      onChange={(e) => handleChange(field, e.target.value)}
                      className={`w-full px-2.5 py-1.5 rounded-lg border text-xs focus:ring-2 focus:outline-none transition-colors ${
                        isChanged
                          ? 'border-amber-400 bg-amber-50/50 dark:bg-amber-950/20 text-slate-900 dark:text-slate-100 focus:ring-amber-500'
                          : isSap
                          ? 'border-blue-400 font-mono font-bold bg-white dark:bg-slate-800 text-blue-900 dark:text-blue-300'
                          : 'border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 focus:ring-blue-500'
                      }`}
                    />
                  </div>
                );
              })}
            </div>
          </div>

          {/* Group 3: REVISION & SIGNATURE (WHITE / LIGHT) */}
          <div className="p-4 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-900">
            <div className="flex items-center gap-2 text-xs font-black text-slate-900 dark:text-slate-100 uppercase tracking-wider mb-3">
              <span className="w-2.5 h-2.5 rounded-full bg-slate-300 inline-block"></span>
              <span>3. REVISION & SIGNATURE (Ký duyệt & Phiên bản)</span>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-3">
              {revisionCols.map((field) => {
                const isChanged = changes.some((c) => c.field === field);
                return (
                  <div key={field}>
                    <label className="block font-semibold text-slate-700 dark:text-slate-300 mb-1 flex items-center justify-between truncate" title={field}>
                      <span className="truncate">{field}</span>
                      {isChanged && (
                        <span className="text-[10px] text-amber-600 dark:text-amber-400 font-bold ml-1">
                          Đã sửa
                        </span>
                      )}
                    </label>
                    <input
                      type="text"
                      value={formData[field] ?? ''}
                      onChange={(e) => handleChange(field, e.target.value)}
                      className={`w-full px-2.5 py-1.5 rounded-lg border text-xs focus:ring-2 focus:outline-none transition-colors ${
                        isChanged
                          ? 'border-amber-400 bg-amber-50/50 dark:bg-amber-950/20 text-slate-900 dark:text-slate-100 focus:ring-amber-500'
                          : 'border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 focus:ring-blue-500'
                      }`}
                    />
                  </div>
                );
              })}
            </div>
          </div>

          {/* Group 4: Other custom fields if any exist */}
          {remainingHeaders.length > 0 && (
            <div>
              <div className="flex items-center gap-2 text-xs font-bold text-slate-700 dark:text-slate-300 uppercase tracking-wider mb-3">
                <Settings className="w-3.5 h-3.5" />
                <span>4. Các Trường Tùy Chỉnh Khác</span>
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
                {remainingHeaders.map((field) => {
                  const isChanged = changes.some((c) => c.field === field);
                  return (
                    <div key={field}>
                      <label className="block font-semibold text-slate-600 dark:text-slate-400 mb-1 flex items-center justify-between truncate" title={field}>
                        <span className="truncate">{field}</span>
                        {isChanged && (
                          <span className="text-[10px] text-amber-600 dark:text-amber-400 font-bold ml-1">
                            Đã sửa
                          </span>
                        )}
                      </label>
                      <input
                        type="text"
                        value={formData[field] ?? ''}
                        onChange={(e) => handleChange(field, e.target.value)}
                        className={`w-full px-2.5 py-1.5 rounded-lg border text-xs focus:ring-2 focus:outline-none transition-colors ${
                          isChanged
                            ? 'border-amber-400 bg-amber-50/50 dark:bg-amber-950/20 text-slate-900 dark:text-slate-100 focus:ring-amber-500'
                            : 'border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 focus:ring-blue-500'
                        }`}
                      />
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 bg-slate-50 dark:bg-slate-850 border-t border-slate-200 dark:border-slate-800 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-4">
            {onDeleteRow && (
              <>
                {!showDeleteConfirm ? (
                  <button
                    type="button"
                    onClick={() => setShowDeleteConfirm(true)}
                    className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold text-rose-600 hover:text-white hover:bg-rose-600 border border-rose-300 dark:border-rose-800 transition-colors cursor-pointer"
                    title="Xóa hẳn dòng van này khỏi bảng Master"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Xóa dòng này</span>
                  </button>
                ) : (
                  <div className="flex items-center gap-2 px-2.5 py-1 rounded-lg bg-rose-50 dark:bg-rose-950/50 border border-rose-300 dark:border-rose-800 text-xs">
                    <span className="text-rose-700 dark:text-rose-300 font-medium">Xác nhận xóa dòng này?</span>
                    <button
                      type="button"
                      onClick={() => {
                        onDeleteRow(row._id);
                        onClose();
                      }}
                      className="px-2 py-0.5 rounded bg-rose-600 text-white font-bold hover:bg-rose-700 transition-colors text-[11px]"
                    >
                      Xóa luôn
                    </button>
                    <button
                      type="button"
                      onClick={() => setShowDeleteConfirm(false)}
                      className="px-2 py-0.5 rounded bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 hover:bg-slate-300 text-[11px]"
                    >
                      Huỷ
                    </button>
                  </div>
                )}
              </>
            )}

            <label className="flex items-center gap-2 text-xs text-slate-600 dark:text-slate-400 cursor-pointer select-none">
              <input
                type="checkbox"
                checked={autoEmail}
                onChange={(e) => setAutoEmail(e.target.checked)}
                className="w-4 h-4 text-blue-600 rounded border-slate-300 dark:border-slate-700 focus:ring-blue-500"
              />
              <span className="hidden sm:inline">Tự động mở cửa sổ Email Outlook sau khi lưu</span>
            </label>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 rounded-lg text-xs font-medium text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-800 transition-colors"
            >
              Huỷ bỏ
            </button>

            <button
              type="button"
              onClick={() => handleSubmit(false)}
              className="inline-flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-semibold bg-slate-700 hover:bg-slate-800 text-white transition-colors"
            >
              <Save className="w-3.5 h-3.5" />
              <span>Chỉ Lưu Dữ Liệu</span>
            </button>

            <button
              type="button"
              onClick={() => handleSubmit(true)}
              className="inline-flex items-center gap-1.5 px-4 py-2 rounded-lg text-xs font-semibold bg-blue-600 hover:bg-blue-700 text-white shadow-sm transition-colors"
            >
              <Mail className="w-3.5 h-3.5" />
              <span>Lưu & Soạn Mail Outlook</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
