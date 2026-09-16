import React, { useState, useEffect } from 'react';
import {
  X,
  Mail,
  Send,
  ExternalLink,
  Copy,
  Check,
  FileEdit,
  Building2,
  BarChart3,
  CheckCircle2,
  FileSpreadsheet,
  Download,
  Paperclip,
  Table,
  HelpCircle,
} from 'lucide-react';
import {
  EmailDraft,
  OutlookEmailContext,
  generateEmailDraft,
  openInOutlookDesktop,
  openInOutlookWeb,
} from '../utils/outlookMailer';
import {
  exportValveChangeToExcel,
  exportToExcel,
  copyHtmlTableToClipboard,
} from '../utils/excelExporter';

interface OutlookEmailModalProps {
  isOpen: boolean;
  onClose: () => void;
  context: OutlookEmailContext | null;
}

export const OutlookEmailModal: React.FC<OutlookEmailModalProps> = ({
  isOpen,
  onClose,
  context,
}) => {
  const [draft, setDraft] = useState<EmailDraft>({
    to: '',
    cc: '',
    subject: '',
    body: '',
  });

  const [activeTemplate, setActiveTemplate] = useState<
    'row_edited' | 'vendor_inquiry' | 'batch_report'
  >('row_edited');

  const [copied, setCopied] = useState(false);
  const [copiedTable, setCopiedTable] = useState(false);
  const [downloadSuccess, setDownloadSuccess] = useState<string | null>(null);
  const [isExporting, setIsExporting] = useState(false);

  // Initialize draft whenever context changes
  useEffect(() => {
    if (context) {
      setActiveTemplate(context.type);
      const generated = generateEmailDraft(context);
      setDraft(generated);
    }
  }, [context]);

  // When switching template manually
  const handleSwitchTemplate = (type: 'row_edited' | 'vendor_inquiry' | 'batch_report') => {
    if (!context) return;
    setActiveTemplate(type);
    const updated = generateEmailDraft({ ...context, type });
    setDraft(updated);
  };

  if (!isOpen || !context) return null;

  // Determine Excel filename based on context
  const tag = context.rowTag || context.rowData?.TAG || 'Armature';
  const supplier = context.supplier || context.rowData?.SUPPLIER || 'ChuaRo';
  const excelFileName =
    activeTemplate === 'row_edited'
      ? `Phieu_CapNhat_Van_${String(tag).replace(/[^a-zA-Z0-9_-]/g, '_')}.xlsx`
      : activeTemplate === 'vendor_inquiry'
      ? `BangKe_Armature_CanXacNhan_${String(supplier).replace(/[^a-zA-Z0-9_-]/g, '_')}.xlsx`
      : 'Armature_Master_Export.xlsx';

  // Function to execute Excel file generation & download
  const handleDownloadExcel = async (): Promise<string> => {
    setIsExporting(true);
    try {
      if (activeTemplate === 'row_edited' && context.rowData && context.masterHeaders) {
        const generatedName = await exportValveChangeToExcel(
          context.masterHeaders,
          context.rowData,
          context.changes || [],
          excelFileName
        );
        setDownloadSuccess(`Đã tải file Excel: ${generatedName}`);
        setTimeout(() => setDownloadSuccess(null), 5000);
        return generatedName;
      } else if (context.allRows && context.masterHeaders) {
        await exportToExcel(
          context.masterHeaders,
          context.allRows,
          context.logs || [],
          excelFileName
        );
        setDownloadSuccess(`Đã tải file Excel: ${excelFileName}`);
        setTimeout(() => setDownloadSuccess(null), 5000);
        return excelFileName;
      } else if (context.rowData && context.masterHeaders) {
        const generatedName = await exportValveChangeToExcel(
          context.masterHeaders,
          context.rowData,
          context.changes || [],
          excelFileName
        );
        setDownloadSuccess(`Đã tải file Excel: ${generatedName}`);
        setTimeout(() => setDownloadSuccess(null), 5000);
        return generatedName;
      }
      return excelFileName;
    } catch (err) {
      console.error('Export excel error:', err);
      return excelFileName;
    } finally {
      setIsExporting(false);
    }
  };

  // Copy raw text body
  const handleCopy = () => {
    const fullText = `Tiêu đề: ${draft.subject}\n\n${draft.body}`;
    navigator.clipboard.writeText(fullText);
    setCopied(true);
    setTimeout(() => setCopied(false), 2200);
  };

  // Copy HTML formatted Excel table to clipboard
  const handleCopyFormattedTable = async () => {
    const { html, text } = generateHtmlTableForClipboard(context);
    await copyHtmlTableToClipboard(html, text);
    setCopiedTable(true);
    setTimeout(() => setCopiedTable(false), 2500);
  };

  // 1-Click Action: Download Excel AND launch Outlook Desktop
  const handleDownloadAndLaunchDesktop = async () => {
    await handleDownloadExcel();
    openInOutlookDesktop(draft);
    setDownloadSuccess(`Đã tải file ${excelFileName} và mở Outlook Desktop! Kéo file vừa tải vào cửa sổ Outlook để đính kèm.`);
  };

  // 1-Click Action: Download Excel AND launch Outlook Web
  const handleDownloadAndLaunchWeb = async () => {
    await handleDownloadExcel();
    openInOutlookWeb(draft, false);
    setDownloadSuccess(`Đã tải file ${excelFileName} và mở Outlook Web! Kéo file vừa tải vào khung thư để đính kèm.`);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-3xl w-full border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden flex flex-col max-h-[94vh]">
        {/* Header */}
        <div className="px-6 py-3.5 bg-gradient-to-r from-blue-700 via-indigo-700 to-blue-800 text-white flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-white/15 flex items-center justify-center backdrop-blur-xs shadow-inner">
              <Mail className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-base font-bold tracking-tight">Soạn & Gửi Email Qua Outlook</h2>
                <span className="text-[10px] px-2 py-0.5 rounded-full font-semibold bg-emerald-500/30 text-emerald-200 border border-emerald-400/30 flex items-center gap-1">
                  <Paperclip className="w-3 h-3" />
                  Kèm file Excel (.xlsx)
                </span>
              </div>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-white/80 hover:text-white hover:bg-white/10 transition-colors cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Template Selector Bar */}
        <div className="px-6 py-2 bg-slate-50 dark:bg-slate-850 border-b border-slate-200 dark:border-slate-800 flex flex-wrap items-center justify-between gap-2 text-xs">
          <div className="flex items-center gap-1.5">
            <span className="text-slate-500 font-medium mr-1">Mẫu thư:</span>
            <button
              onClick={() => handleSwitchTemplate('row_edited')}
              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium transition-colors cursor-pointer ${
                activeTemplate === 'row_edited'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'bg-slate-200/70 dark:bg-slate-800 text-slate-700 dark:text-slate-300 hover:bg-slate-300/70'
              }`}
            >
              <FileEdit className="w-3.5 h-3.5" />
              <span>Thay đổi van vừa sửa</span>
            </button>
            <button
              onClick={() => handleSwitchTemplate('vendor_inquiry')}
              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium transition-colors cursor-pointer ${
                activeTemplate === 'vendor_inquiry'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'bg-slate-200/70 dark:bg-slate-800 text-slate-700 dark:text-slate-300 hover:bg-slate-300/70'
              }`}
            >
              <Building2 className="w-3.5 h-3.5" />
              <span>Yêu cầu NCC bổ sung</span>
            </button>
            <button
              onClick={() => handleSwitchTemplate('batch_report')}
              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium transition-colors cursor-pointer ${
                activeTemplate === 'batch_report'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'bg-slate-200/70 dark:bg-slate-800 text-slate-700 dark:text-slate-300 hover:bg-slate-300/70'
              }`}
            >
              <BarChart3 className="w-3.5 h-3.5" />
              <span>Báo cáo tiến độ Master</span>
            </button>
          </div>

          {context.changes && context.changes.length > 0 && (
            <span className="text-[11px] font-semibold text-amber-700 dark:text-amber-300 bg-amber-50 dark:bg-amber-950/40 px-2 py-0.5 rounded border border-amber-200 dark:border-amber-900">
              Đã ghi nhận {context.changes.length} mục thay đổi
            </span>
          )}
        </div>

        {/* Modal Body: Email Form & Attachment Section */}
        <div className="p-5 sm:p-6 overflow-y-auto space-y-4 text-xs">
          {/* Notification banner if file downloaded or message */}
          {downloadSuccess && (
            <div className="p-2.5 rounded-lg bg-emerald-50 dark:bg-emerald-950/40 text-emerald-800 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 flex items-center gap-2 shadow-xs">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
              <span className="font-semibold text-xs">{downloadSuccess}</span>
            </div>
          )}

          {/* ATTACHMENT CARD: EXCEL FILE TO BE ATTACHED */}
          <div className="rounded-xl border border-emerald-200 dark:border-emerald-900/60 bg-emerald-50/50 dark:bg-emerald-950/20 p-3.5">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-xs">
                  <FileSpreadsheet className="w-5 h-5" />
                </div>
                <div>
                  <div className="flex items-center gap-2">
                    <span className="font-bold text-slate-900 dark:text-slate-100 text-xs">
                      {excelFileName}
                    </span>
                    <span className="text-[10px] font-semibold px-2 py-0.2 rounded bg-emerald-200 dark:bg-emerald-900 text-emerald-800 dark:text-emerald-200">
                      Excel .xlsx
                    </span>
                  </div>
                  <p className="text-[11px] text-slate-600 dark:text-slate-400">
                    {activeTemplate === 'row_edited'
                      ? 'Phiếu chi tiết thông số van & bảng highlight các trường vừa hiệu chỉnh'
                      : 'Bảng kê Master Data tổng hợp toàn bộ vật tư và nhật ký xử lý'}
                  </p>
                </div>
              </div>

              {/* Download & Copy Table Buttons */}
              <div className="flex items-center gap-2 shrink-0">
                <button
                  type="button"
                  id="btn-copy-excel-table"
                  onClick={handleCopyFormattedTable}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 text-slate-700 dark:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors shadow-2xs cursor-pointer"
                  title="Sao chép bảng định dạng màu sắc & kẻ khung để dán trực tiếp (Ctrl+V) vào nội dung thư Outlook"
                >
                  {copiedTable ? (
                    <>
                      <Check className="w-3.5 h-3.5 text-emerald-600" />
                      <span className="text-emerald-600 font-bold">Đã copy bảng!</span>
                    </>
                  ) : (
                    <>
                      <Table className="w-3.5 h-3.5 text-blue-600" />
                      <span>Copy bảng vào thư (Ctrl+V)</span>
                    </>
                  )}
                </button>

                <button
                  type="button"
                  id="btn-download-excel-attachment"
                  onClick={handleDownloadExcel}
                  disabled={isExporting}
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold bg-emerald-700 hover:bg-emerald-600 text-white transition-colors shadow-xs cursor-pointer disabled:opacity-50"
                  title="Tải file Excel này về máy tính"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>{isExporting ? 'Đang tạo...' : 'Tải file Excel'}</span>
                </button>
              </div>
            </div>
          </div>

          {/* Recipients Row */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div>
              <label className="block font-semibold text-slate-700 dark:text-slate-300 mb-1">
                Người nhận (To):
              </label>
              <input
                type="text"
                value={draft.to}
                onChange={(e) => setDraft({ ...draft, to: e.target.value })}
                placeholder="ví dụ: vendor@valvesupplier.com, manager@project.com"
                className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 focus:ring-2 focus:ring-blue-500 focus:outline-none"
              />
            </div>
            <div>
              <label className="block font-semibold text-slate-700 dark:text-slate-300 mb-1">
                Đồng gửi (CC):
              </label>
              <input
                type="text"
                value={draft.cc}
                onChange={(e) => setDraft({ ...draft, cc: e.target.value })}
                placeholder="cc_lead@project.com, piping_team@company.com"
                className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-slate-100 focus:ring-2 focus:ring-blue-500 focus:outline-none"
              />
            </div>
          </div>

          {/* Subject Row */}
          <div>
            <label className="block font-semibold text-slate-700 dark:text-slate-300 mb-1">
              Tiêu đề Email (Subject):
            </label>
            <input
              type="text"
              value={draft.subject}
              onChange={(e) => setDraft({ ...draft, subject: e.target.value })}
              className="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 font-semibold text-slate-900 dark:text-slate-100 focus:ring-2 focus:ring-blue-500 focus:outline-none"
            />
          </div>

          {/* Body Content */}
          <div>
            <div className="flex items-center justify-between mb-1">
              <label className="font-semibold text-slate-700 dark:text-slate-300 flex items-center gap-1.5">
                <span>Nội dung Email:</span>
                <span className="text-[11px] text-slate-400 font-normal">
                  (Đã tự động đính kèm thông báo file Excel)
                </span>
              </label>
              <button
                type="button"
                onClick={handleCopy}
                className="inline-flex items-center gap-1 text-[11px] font-medium text-slate-500 hover:text-slate-800 dark:hover:text-slate-200 transition-colors cursor-pointer"
              >
                {copied ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                    <span className="text-emerald-600 font-semibold">Đã sao chép!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5" />
                    <span>Sao chép text</span>
                  </>
                )}
              </button>
            </div>
            <textarea
              rows={9}
              value={draft.body}
              onChange={(e) => setDraft({ ...draft, body: e.target.value })}
              className="w-full px-3.5 py-2.5 rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 font-sans text-xs text-slate-800 dark:text-slate-200 focus:ring-2 focus:ring-blue-500 focus:outline-none leading-relaxed"
            />
          </div>
        </div>

        {/* Footer: Direct Outlook Actions */}
        <div className="px-6 py-3.5 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              className="px-3.5 py-2 rounded-lg text-xs font-semibold text-slate-600 hover:text-slate-800 dark:text-slate-400 dark:hover:text-slate-200 transition-colors cursor-pointer"
            >
              Đóng
            </button>
          </div>

          {/* 1-Click Launch & Attach Outlook Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            {/* Outlook Web / Office 365 */}
            <button
              id="btn-open-outlook-web"
              onClick={handleDownloadAndLaunchWeb}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-xs font-semibold bg-sky-600 hover:bg-sky-500 text-white shadow-xs transition-colors cursor-pointer"
              title="Tải file Excel và mở ngay trong Outlook Web (Office 365)"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              <span>1-Click: Tải Excel & Mở Outlook Web</span>
            </button>

            {/* Outlook Desktop App */}
            <button
              id="btn-open-outlook-desktop"
              onClick={handleDownloadAndLaunchDesktop}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold bg-blue-700 hover:bg-blue-600 text-white shadow-sm transition-all cursor-pointer"
              title="Tải file Excel và mở ngay ứng dụng Microsoft Outlook Desktop đã cài trên máy"
            >
              <Send className="w-3.5 h-3.5" />
              <span>1-Click: Tải Excel & Mở Outlook App</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

/**
 * Generate formatted HTML and PlainText tables to copy to clipboard for Outlook pasting
 */
function generateHtmlTableForClipboard(context: OutlookEmailContext): { html: string; text: string } {
  const tag = context.rowTag || context.rowData?.TAG || 'Armature';
  const supplier = context.supplier || context.rowData?.SUPPLIER || '';
  const changes = context.changes || [];
  const masterHeaders = context.masterHeaders || [];
  const rowData = context.rowData || {};

  let html = `<div style="font-family: Segoe UI, Arial, sans-serif; font-size: 13px; color: #1e293b; max-width: 650px;">`;
  html += `<h3 style="color: #1e3a8a; margin-bottom: 6px; border-bottom: 2px solid #1e3a8a; padding-bottom: 4px;">THÔNG TIN VẬT TƯ VAN: ${tag} (${supplier})</h3>`;

  if (changes.length > 0) {
    html += `<h4 style="color: #b45309; margin-top: 12px; margin-bottom: 6px;">1. CÁC THÔNG SỐ VỪA THAY ĐỔI:</h4>`;
    html += `<table style="border-collapse: collapse; width: 100%; border: 1px solid #cbd5e1; font-size: 12px; margin-bottom: 14px;">`;
    html += `<tr style="background-color: #d97706; color: white;">
      <th style="padding: 6px 10px; border: 1px solid #b45309; text-align: left;">Thông số (Field)</th>
      <th style="padding: 6px 10px; border: 1px solid #b45309; text-align: left;">Giá trị trước</th>
      <th style="padding: 6px 10px; border: 1px solid #b45309; text-align: left;">Giá trị cập nhật mới</th>
    </tr>`;
    changes.forEach((c) => {
      html += `<tr>
        <td style="padding: 6px 10px; border: 1px solid #e2e8f0; font-weight: bold; background-color: #fffbeb;">${c.field}</td>
        <td style="padding: 6px 10px; border: 1px solid #e2e8f0; text-decoration: line-through; color: #64748b;">${c.oldValue ?? '—'}</td>
        <td style="padding: 6px 10px; border: 1px solid #e2e8f0; font-weight: bold; color: #15803d; background-color: #ecfdf5;">${c.newValue ?? '—'}</td>
      </tr>`;
    });
    html += `</table>`;
  }

  if (masterHeaders.length > 0 && rowData) {
    html += `<h4 style="color: #1e3a8a; margin-top: 12px; margin-bottom: 6px;">2. BẢNG TOÀN BỘ THÔNG SỐ KỸ THUẬT HIỆN TẠI:</h4>`;
    html += `<table style="border-collapse: collapse; width: 100%; border: 1px solid #cbd5e1; font-size: 11.5px;">`;
    html += `<tr style="background-color: #1e3a8a; color: white;">
      <th style="padding: 6px 8px; border: 1px solid #0f172a; text-align: left; width: 40%;">Tên trường</th>
      <th style="padding: 6px 8px; border: 1px solid #0f172a; text-align: left;">Giá trị thiết kế</th>
    </tr>`;
    masterHeaders.forEach((h, idx) => {
      const isChanged = changes.some((c) => c.field === h);
      const val = rowData[h];
      html += `<tr style="background-color: ${isChanged ? '#eff6ff' : idx % 2 === 0 ? '#ffffff' : '#f8fafc'};">
        <td style="padding: 5px 8px; border: 1px solid #e2e8f0; font-weight: ${isChanged ? 'bold' : 'normal'}; color: ${isChanged ? '#1e40af' : '#334155'};">${h}</td>
        <td style="padding: 5px 8px; border: 1px solid #e2e8f0; font-weight: ${isChanged ? 'bold' : 'normal'}; color: ${isChanged ? '#1e40af' : '#0f172a'};">${val ?? '—'}</td>
      </tr>`;
    });
    html += `</table>`;
  }
  html += `</div>`;

  let text = `THÔNG TIN VẬT TƯ VAN: ${tag} (${supplier})\n`;
  if (changes.length > 0) {
    text += `\nCÁC THÔNG SỐ VỪA THAY ĐỔI:\n`;
    changes.forEach((c) => {
      text += `• ${c.field}: ${c.oldValue ?? ''} -> ${c.newValue ?? ''}\n`;
    });
  }
  if (masterHeaders.length > 0) {
    text += `\nTHÔNG SỐ KỸ THUẬT:\n`;
    masterHeaders.forEach((h) => {
      text += `• ${h}: ${rowData[h] ?? '—'}\n`;
    });
  }

  return { html, text };
}
