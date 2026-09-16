import React from 'react';
import {
  FileSpreadsheet,
  Download,
  Trash2,
  Sparkles,
  HelpCircle,
  FileText,
  SlidersHorizontal,
  Github,
  Mail,
  Table,
  BookOpen,
} from 'lucide-react';
import { downloadMasterTemplate } from '../utils/sampleData';
import { downloadCatalogTemplate } from '../utils/catalogManager';

interface HeaderProps {
  totalRows: number;
  totalFiles: number;
  catalogCount?: number;
  activeShipName?: string;
  onExportExcel: () => void;
  onExportCSV: () => void;
  onLoadSampleData: () => void;
  onClearData: () => void;
  onOpenAliases: () => void;
  onOpenHelp: () => void;
  onOpenGithubDeploy: () => void;
  onOpenOutlookMail: () => void;
  onOpenCatalogManager: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  totalRows,
  totalFiles,
  catalogCount = 0,
  activeShipName = 'Tàu Mẫu',
  onExportExcel,
  onExportCSV,
  onLoadSampleData,
  onClearData,
  onOpenAliases,
  onOpenHelp,
  onOpenGithubDeploy,
  onOpenOutlookMail,
  onOpenCatalogManager,
}) => {
  return (
    <header className="bg-slate-900 border-b border-slate-800 text-white sticky top-0 z-30 shadow-md">
      <div className="w-full px-3 sm:px-5 lg:px-6 py-2.5">
        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
          {/* Logo & Title */}
          <div className="flex items-center space-x-3">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-tr from-blue-600 to-indigo-500 flex items-center justify-center shadow-inner shadow-white/20 shrink-0">
              <FileSpreadsheet className="w-5 h-5 text-white" />
            </div>
            <div className="flex items-center gap-2">
              <h1 className="text-base font-bold tracking-tight text-white flex items-center gap-2">
                ARMATURE IMPORT TOOL
                <span className="text-[11px] px-2 py-0.5 rounded-full font-medium bg-blue-500/20 text-blue-300 border border-blue-500/30">
                  VBA Web Edition
                </span>
              </h1>
            </div>
          </div>

          {/* Quick Metrics & Actions */}
          <div className="flex flex-wrap items-center gap-2">
            <div className="flex items-center bg-slate-800/90 rounded-lg px-3 py-1.5 border border-slate-700/60 text-xs">
              <span className="text-slate-400 mr-1.5">Armature List:</span>
              <span className="font-bold text-emerald-400">{totalRows} hàng</span>
              <span className="mx-2 text-slate-600">|</span>
              <span className="text-slate-400 mr-1.5">Catalog ngầm:</span>
              <span className="font-semibold text-amber-300">{catalogCount} mã</span>
            </div>

            {/* Test Sample Data */}
            <button
              id="btn-load-sample"
              onClick={onLoadSampleData}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-indigo-600/30 hover:bg-indigo-600/50 text-indigo-200 border border-indigo-500/40 transition-colors shadow-sm"
              title="Tải 3 file Excel mẫu để thử nghiệm nhanh chức năng import"
            >
              <Sparkles className="w-3.5 h-3.5 text-indigo-300" />
              <span>Nạp File Mẫu Test</span>
            </button>

            {/* Download Master Template */}
            <button
              id="btn-download-template"
              onClick={downloadMasterTemplate}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors"
              title="Tải về file Excel mẫu Total + Import_Log trống"
            >
              <FileText className="w-3.5 h-3.5 text-slate-300" />
              <span className="hidden sm:inline">Mẫu Master</span>
            </button>

            {/* Hidden Catalog Manager Button */}
            <button
              id="btn-open-catalog-manager-header"
              onClick={onOpenCatalogManager}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-amber-300 border border-slate-700 transition-colors"
              title="Quản lý / Nạp file Catalog Excel nền theo từng tàu"
            >
              <BookOpen className="w-3.5 h-3.5 text-amber-400" />
              <span>Catalog Nền ({catalogCount})</span>
            </button>

            {/* Config Aliases */}
            <button
              id="btn-open-aliases"
              onClick={onOpenAliases}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors"
              title="Cấu hình quy tắc Alias tên cột"
            >
              <SlidersHorizontal className="w-3.5 h-3.5 text-slate-300" />
              <span className="hidden sm:inline">Alias Cột</span>
            </button>

            {/* Export Dropdown / Buttons */}
            <button
              id="btn-export-excel"
              onClick={onExportExcel}
              disabled={totalRows === 0}
              className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition-all ${
                totalRows > 0
                  ? 'bg-emerald-600 hover:bg-emerald-500 text-white cursor-pointer active:scale-95'
                  : 'bg-slate-800 text-slate-500 border border-slate-700 cursor-not-allowed opacity-60'
              }`}
              title="Xuất file Excel gồm 2 Sheet: Total & Import_Log theo chuẩn macro"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Xuất Excel (.xlsx)</span>
            </button>

            {/* Outlook Email Button */}
            <button
              id="btn-header-outlook-mail"
              onClick={onOpenOutlookMail}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-blue-700 hover:bg-blue-600 text-white shadow-sm transition-all"
              title="Soạn thảo và gửi email báo cáo qua Outlook"
            >
              <Mail className="w-3.5 h-3.5" />
              <span>Mail Outlook</span>
            </button>

            {/* Help / Guide */}
            <button
              id="btn-open-help"
              onClick={onOpenHelp}
              className="p-1.5 rounded-lg text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors"
              title="Xem hướng dẫn nguyên lý macro VBA"
            >
              <HelpCircle className="w-4 h-4" />
            </button>

            {/* GitHub Deploy Guide Button */}
            <button
              id="btn-open-github-deploy"
              onClick={onOpenGithubDeploy}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors"
              title="Xem hướng dẫn và lệnh Deploy lên GitHub Pages"
            >
              <Github className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Deploy GitHub</span>
            </button>

            {/* Clear button */}
            {totalRows > 0 && (
              <button
                id="btn-clear-master"
                onClick={onClearData}
                className="p-1.5 rounded-lg text-rose-400 hover:text-rose-300 hover:bg-rose-950/40 transition-colors"
                title="Xoá toàn bộ dữ liệu bảng Armature"
              >
                <Trash2 className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>
      </div>
    </header>
  );
};
