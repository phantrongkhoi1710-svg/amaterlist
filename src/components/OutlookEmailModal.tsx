import React, { useState, useEffect } from 'react';
import {
  X,
  Mail,
  Send,
  ExternalLink,
  Copy,
  Check,
  Sparkles,
  FileEdit,
  Building2,
  BarChart3,
  CheckCircle2,
} from 'lucide-react';
import {
  EmailDraft,
  OutlookEmailContext,
  generateEmailDraft,
  openInOutlookDesktop,
  openInOutlookWeb,
} from '../utils/outlookMailer';

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
  const [sendFeedback, setSendFeedback] = useState<string | null>(null);

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

  const handleCopy = () => {
    const fullText = `Tiêu đề: ${draft.subject}\n\n${draft.body}`;
    navigator.clipboard.writeText(fullText);
    setCopied(true);
    setTimeout(() => setCopied(false), 2200);
  };

  const handleLaunchDesktop = () => {
    openInOutlookDesktop(draft);
    setSendFeedback('Đã kích hoạt mở ứng dụng Microsoft Outlook trên máy!');
    setTimeout(() => setSendFeedback(null), 4000);
  };

  const handleLaunchWeb = (isPersonal = false) => {
    openInOutlookWeb(draft, isPersonal);
    setSendFeedback('Đã mở tab soạn thảo Outlook Web!');
    setTimeout(() => setSendFeedback(null), 4000);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-3xl w-full border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden flex flex-col max-h-[92vh]">
        {/* Header */}
        <div className="px-6 py-4 bg-gradient-to-r from-blue-700 to-indigo-800 text-white flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-white/15 flex items-center justify-center backdrop-blur-xs">
              <Mail className="w-5 h-5 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-base font-bold tracking-tight">Soạn & Gửi Email Qua Outlook</h2>
                <span className="text-[10px] px-2 py-0.5 rounded-full font-semibold bg-blue-500/40 text-blue-100 border border-blue-400/30">
                  Tự động sau khi Edit
                </span>
              </div>
              <p className="text-xs text-blue-100/80">
                Nội dung đã được chuẩn bị sẵn, mở trực tiếp bằng Microsoft Outlook Desktop hoặc Office 365
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-white/80 hover:text-white hover:bg-white/10 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Template Selector Bar */}
        <div className="px-6 py-2.5 bg-slate-50 dark:bg-slate-850 border-b border-slate-200 dark:border-slate-800 flex flex-wrap items-center justify-between gap-2 text-xs">
          <div className="flex items-center gap-1.5">
            <span className="text-slate-500 font-medium mr-1">Mẫu Email:</span>
            <button
              onClick={() => handleSwitchTemplate('row_edited')}
              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium transition-colors ${
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
              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium transition-colors ${
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
              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-medium transition-colors ${
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
            <span className="text-[11px] font-semibold text-amber-600 dark:text-amber-400 bg-amber-50 dark:bg-amber-950/40 px-2 py-0.5 rounded border border-amber-200 dark:border-amber-900">
              Đã ghi nhận {context.changes.length} mục thay đổi
            </span>
          )}
        </div>

        {/* Modal Body: Email Form */}
        <div className="p-6 overflow-y-auto space-y-4 text-xs">
          {/* Feedback alert if any */}
          {sendFeedback && (
            <div className="p-3 rounded-lg bg-emerald-50 dark:bg-emerald-950/40 text-emerald-800 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
              <span>{sendFeedback}</span>
            </div>
          )}

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
              <label className="font-semibold text-slate-700 dark:text-slate-300">
                Nội dung Email:
              </label>
              <button
                type="button"
                onClick={handleCopy}
                className="inline-flex items-center gap-1 text-[11px] font-medium text-slate-500 hover:text-slate-800 dark:hover:text-slate-200 transition-colors"
              >
                {copied ? (
                  <>
                    <Check className="w-3.5 h-3.5 text-emerald-600" />
                    <span className="text-emerald-600 font-semibold">Đã sao chép!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3.5 h-3.5" />
                    <span>Sao chép nội dung</span>
                  </>
                )}
              </button>
            </div>
            <textarea
              rows={11}
              value={draft.body}
              onChange={(e) => setDraft({ ...draft, body: e.target.value })}
              className="w-full px-3.5 py-2.5 rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 font-sans text-xs text-slate-800 dark:text-slate-200 focus:ring-2 focus:ring-blue-500 focus:outline-none leading-relaxed"
            />
          </div>
        </div>

        {/* Footer: Direct Outlook Actions */}
        <div className="px-6 py-4 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-2">
            <button
              onClick={handleCopy}
              className="inline-flex items-center gap-1.5 px-3 py-2 rounded-lg text-xs font-semibold bg-white dark:bg-slate-800 border border-slate-300 dark:border-slate-700 hover:bg-slate-100 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-200 transition-colors"
            >
              <Copy className="w-3.5 h-3.5" />
              <span>{copied ? 'Đã sao chép' : 'Sao chép mail'}</span>
            </button>
            <button
              onClick={onClose}
              className="px-3 py-2 rounded-lg text-xs text-slate-500 hover:text-slate-700 dark:hover:text-slate-300 transition-colors"
            >
              Đóng
            </button>
          </div>

          {/* Direct Outlook Launch Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            {/* Outlook Web / Office 365 */}
            <button
              id="btn-open-outlook-web"
              onClick={() => handleLaunchWeb(false)}
              className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-lg text-xs font-semibold bg-sky-600 hover:bg-sky-500 text-white shadow-xs transition-colors"
              title="Mở ngay trong Outlook Web (Office 365)"
            >
              <ExternalLink className="w-3.5 h-3.5" />
              <span>Mở Outlook Web (365)</span>
            </button>

            {/* Outlook Desktop App */}
            <button
              id="btn-open-outlook-desktop"
              onClick={handleLaunchDesktop}
              className="inline-flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold bg-blue-700 hover:bg-blue-600 text-white shadow-sm transition-all"
              title="Kích hoạt ứng dụng Microsoft Outlook Desktop đã cài trên máy tính"
            >
              <Send className="w-3.5 h-3.5" />
              <span>Mở Outlook App (Desktop)</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
