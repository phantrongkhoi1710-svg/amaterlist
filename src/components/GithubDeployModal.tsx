import React, { useState } from 'react';
import { X, Github, Check, Copy, ExternalLink, Terminal, Rocket, Sparkles } from 'lucide-react';

interface GithubDeployModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const GithubDeployModal: React.FC<GithubDeployModalProps> = ({ isOpen, onClose }) => {
  const [copiedIndex, setCopiedIndex] = useState<number | null>(null);

  if (!isOpen) return null;

  const copyToClipboard = (text: string, index: number) => {
    navigator.clipboard.writeText(text);
    setCopiedIndex(index);
    setTimeout(() => setCopiedIndex(null), 2000);
  };

  const gitCommands = `git init
git add .
git commit -m "Deploy Armature Import Tool to GitHub Pages"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo-name>.git
git push -u origin main`;

  const npmDeployCommand = `npm run deploy`;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-2xl w-full border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* Modal Header */}
        <div className="px-6 py-4 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-slate-900 dark:bg-slate-800 text-white flex items-center justify-center">
              <Github className="w-4 h-4" />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-900 dark:text-slate-100 flex items-center gap-2">
                Hướng Dẫn Deploy Lên GitHub Pages
                <span className="text-[10px] px-2 py-0.5 rounded-full font-medium bg-emerald-100 text-emerald-700 dark:bg-emerald-950/60 dark:text-emerald-300">
                  Sẵn sàng 100%
                </span>
              </h2>
              <p className="text-xs text-slate-500 dark:text-slate-400">
                Đã cấu hình sẵn file <code>.github/workflows/deploy.yml</code> & <code>base: &apos;./&apos;</code>
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-6 overflow-y-auto space-y-5 text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
          {/* Method 1: GitHub Actions (Recommended) */}
          <div className="border border-blue-200 dark:border-blue-900/40 bg-blue-50/60 dark:bg-blue-950/30 rounded-xl p-4">
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-sm font-bold text-blue-900 dark:text-blue-200 flex items-center gap-2">
                <Rocket className="w-4 h-4 text-blue-600" />
                Cách 1: Tự động hoá bằng GitHub Actions (Khuyên dùng)
              </h3>
              <span className="text-[10px] font-semibold bg-blue-600 text-white px-2 py-0.5 rounded">
                Tự động build khi Push
              </span>
            </div>
            <p className="text-slate-600 dark:text-slate-300 text-xs mb-3">
              Dự án đã có sẵn file cấu hình <code>.github/workflows/deploy.yml</code>. Bạn chỉ cần đẩy code lên repository:
            </p>

            {/* Code Box */}
            <div className="relative bg-slate-950 text-slate-200 rounded-lg p-3 font-mono text-[11px] overflow-x-auto">
              <button
                onClick={() => copyToClipboard(gitCommands, 1)}
                className="absolute top-2.5 right-2.5 inline-flex items-center gap-1 px-2 py-1 rounded bg-slate-800 hover:bg-slate-700 text-[10px] text-slate-300 transition-colors"
              >
                {copiedIndex === 1 ? (
                  <>
                    <Check className="w-3 h-3 text-emerald-400" />
                    <span className="text-emerald-400">Đã chép</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3 h-3" />
                    <span>Sao chép</span>
                  </>
                )}
              </button>
              <pre className="pr-16 leading-relaxed">{gitCommands}</pre>
            </div>

            {/* Step on GitHub UI */}
            <div className="mt-3 text-xs bg-white dark:bg-slate-900 p-3 rounded-lg border border-slate-200 dark:border-slate-800 space-y-1.5">
              <p className="font-semibold text-slate-800 dark:text-slate-200">
                Bước kích hoạt trên GitHub:
              </p>
              <ol className="list-decimal list-inside space-y-1 text-slate-600 dark:text-slate-300">
                <li>Vào Repo của bạn trên GitHub &rarr; chọn tab <strong>Settings</strong></li>
                <li>Chọn mục <strong>Pages</strong> ở thanh menu bên trái</li>
                <li>Tại mục <strong>Build and deployment &gt; Source</strong>: Chọn <strong>GitHub Actions</strong></li>
                <li>Hệ thống sẽ tự động chạy quy trình build và cung cấp link trang web công khai!</li>
              </ol>
            </div>
          </div>

          {/* Method 2: Manual gh-pages command */}
          <div className="border border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900/50 rounded-xl p-4">
            <div className="flex items-center justify-between mb-2">
              <h3 className="text-sm font-bold text-slate-800 dark:text-slate-200 flex items-center gap-2">
                <Terminal className="w-4 h-4 text-slate-600" />
                Cách 2: Deploy thủ công từ máy bằng lệnh <code>npm run deploy</code>
              </h3>
            </div>
            <p className="text-xs text-slate-500 dark:text-slate-400 mb-2">
              Gói <code>gh-pages</code> đã được cài đặt sẵn vào <code>package.json</code>. Bạn có thể mở terminal và chạy:
            </p>
            <div className="relative bg-slate-950 text-slate-200 rounded-lg p-2.5 font-mono text-[11px]">
              <button
                onClick={() => copyToClipboard(npmDeployCommand, 2)}
                className="absolute top-2 right-2 inline-flex items-center gap-1 px-2 py-0.5 rounded bg-slate-800 hover:bg-slate-700 text-[10px] text-slate-300 transition-colors"
              >
                {copiedIndex === 2 ? (
                  <>
                    <Check className="w-3 h-3 text-emerald-400" />
                    <span className="text-emerald-400">Đã chép</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-3 h-3" />
                    <span>Sao chép</span>
                  </>
                )}
              </button>
              <code>npm run deploy</code>
            </div>
          </div>
        </div>

        {/* Modal Footer */}
        <div className="px-6 py-3.5 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 flex items-center justify-between">
          <span className="text-[11px] text-slate-500">
            Xem thêm file <code>README.md</code> trong thư mục gốc
          </span>
          <button
            onClick={onClose}
            className="px-4 py-2 rounded-lg text-xs font-semibold bg-slate-900 dark:bg-slate-800 hover:bg-slate-800 dark:hover:bg-slate-700 text-white transition-colors"
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  );
};
