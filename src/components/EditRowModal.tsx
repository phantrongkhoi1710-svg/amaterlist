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
} from 'lucide-react';
import { MasterRowData } from '../types';
import { RowEditChange } from '../utils/outlookMailer';

interface EditRowModalProps {
  isOpen: boolean;
  onClose: () => void;
  row: MasterRowData | null;
  masterHeaders: string[];
  onSave: (updatedRow: MasterRowData, changes: RowEditChange[], autoComposeMail: boolean) => void;
}

export const EditRowModal: React.FC<EditRowModalProps> = ({
  isOpen,
  onClose,
  row,
  masterHeaders,
  onSave,
}) => {
  const [formData, setFormData] = useState<Record<string, any>>({});
  const [autoEmail, setAutoEmail] = useState(true);

  useEffect(() => {
    if (row) {
      // Clone row data
      const data: Record<string, any> = {};
      masterHeaders.forEach((h) => {
        data[h] = row[h] ?? '';
      });
      setFormData(data);
      setAutoEmail(true);
    }
  }, [row, masterHeaders]);

  if (!isOpen || !row) return null;

  const handleChange = (header: string, value: string) => {
    setFormData((prev) => ({
      ...prev,
      [header]: value,
    }));
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
    const updatedRow: MasterRowData = {
      ...row,
      ...formData,
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
  ];

  // Find remaining headers if any
  const remainingHeaders = masterHeaders.filter(
    (h) => !functionalDesignCols.includes(h) && !pipeSpecCols.includes(h)
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
          {/* Group 1: TO BE COMPLETED BY FUNCTIONAL DESIGN (YELLOW) */}
          <div className="p-4 rounded-xl border border-yellow-300 dark:border-yellow-800/60 bg-yellow-50/40 dark:bg-yellow-950/20">
            <div className="flex items-center gap-2 text-xs font-black text-amber-950 dark:text-yellow-200 uppercase tracking-wider mb-3">
              <span className="w-2.5 h-2.5 rounded-full bg-yellow-400 inline-block"></span>
              <span>1. TO BE COMPLETED BY FUNCTIONAL DESIGN</span>
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
              <span>2. ARMATURE INFO FROM PIPE SPECIFICATION (Phân tích từ Catalog)</span>
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

          {/* Group 3: Other custom fields if any exist */}
          {remainingHeaders.length > 0 && (
            <div>
              <div className="flex items-center gap-2 text-xs font-bold text-slate-700 dark:text-slate-300 uppercase tracking-wider mb-3">
                <Settings className="w-3.5 h-3.5" />
                <span>3. Các Trường Tùy Chỉnh Khác</span>
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
        <div className="px-6 py-4 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 flex flex-wrap items-center justify-between gap-3">
          {/* Checkbox option */}
          <label className="inline-flex items-center gap-2 cursor-pointer select-none text-xs text-slate-700 dark:text-slate-300">
            <input
              type="checkbox"
              checked={autoEmail}
              onChange={(e) => setAutoEmail(e.target.checked)}
              className="w-4 h-4 rounded text-blue-600 focus:ring-blue-500 border-slate-300"
            />
            <span className="flex items-center gap-1 font-medium">
              <Mail className="w-3.5 h-3.5 text-blue-600" />
              Tự động chuẩn bị Outlook và file Excel đính kèm sau khi lưu
            </span>
          </label>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-3.5 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-200 transition-colors cursor-pointer"
            >
              Hủy bỏ
            </button>

            {/* Save Only */}
            <button
              type="button"
              onClick={() => handleSubmit(false)}
              className="px-4 py-2 rounded-lg text-xs font-semibold bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 transition-colors flex items-center gap-1.5 cursor-pointer"
            >
              <Save className="w-3.5 h-3.5" />
              <span>Chỉ lưu dữ liệu</span>
            </button>

            {/* Save & Compose Outlook */}
            <button
              type="button"
              id="btn-save-and-mail"
              onClick={() => handleSubmit(true)}
              className="px-4 py-2 rounded-lg text-xs font-bold bg-blue-600 hover:bg-blue-500 text-white shadow-sm transition-all flex items-center gap-2 cursor-pointer"
            >
              <Mail className="w-3.5 h-3.5" />
              <span>Lưu & Soạn Mail (Kèm Excel)</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
