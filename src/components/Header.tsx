import React from 'react';
import { Link } from 'react-router-dom';
import {
  FileSpreadsheet,
  Download,
  Trash2,
  Sparkles,
  HelpCircle,
  FileText,
  SlidersHorizontal,
  Mail,
  BookOpen,
  ArrowLeft,
  LayoutGrid,
} from 'lucide-react';
import { downloadMasterTemplate } from '../utils/sampleData';

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
  onOpenGithubDeploy?: () => void;
  onOpenOutlookMail: () => void;
  onOpenCatalogManager: () => void;
}

export const Header: React.FC<HeaderProps> = ({
  totalRows,
  totalFiles,
  catalogCount = 0,
  activeShipName = 'Sample Ship',
  onExportExcel,
  onExportCSV,
  onLoadSampleData,
  onClearData,
  onOpenAliases,
  onOpenHelp,
  onOpenOutlookMail,
  onOpenCatalogManager,
}) => {
  return (
    <header className="bg-slate-900 border-b border-slate-800 text-white sticky top-0 z-30 shadow-md">
      <div className="w-full px-3 sm:px-5 lg:px-6 py-2.5">
        <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
          {/* Logo, Title & Back to Hub */}
          <div className="flex items-center space-x-2.5">
            <Link
              to="/"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-750 text-slate-300 hover:text-white border border-slate-700 text-xs font-bold transition-all shadow-sm active:scale-95 cursor-pointer group"
              title="Quay lại Trang chủ Hub Công cụ / Back to WebToolRush Hub"
            >
              <ArrowLeft className="w-3.5 h-3.5 text-slate-400 group-hover:-translate-x-0.5 transition-transform" />
              <LayoutGrid className="w-3.5 h-3.5 text-cyan-400" />
              <span className="hidden md:inline font-semibold text-slate-200">Tool Hub</span>
            </Link>

            <div className="w-8 h-8 rounded-lg bg-gradient-to-tr from-blue-600 to-indigo-500 flex items-center justify-center shadow-inner shadow-white/20 shrink-0">
              <FileSpreadsheet className="w-5 h-5 text-white" />
            </div>
            <div className="flex items-center gap-2">
              <h1 className="text-base font-bold tracking-tight text-white flex items-center gap-2">
                ARMATURE IMPORT TOOL
              </h1>
            </div>
          </div>

          {/* Quick Metrics & Actions */}
          <div className="flex flex-wrap items-center gap-2">
            <div className="flex items-center bg-slate-800/90 rounded-lg px-3 py-1.5 border border-slate-700/60 text-xs">
              <span className="text-slate-400 mr-1.5">Armature List:</span>
              <span className="font-bold text-emerald-400">{totalRows} rows</span>
              <span className="mx-2 text-slate-600">|</span>
              <span className="text-slate-400 mr-1.5">Catalog:</span>
              <span className="font-semibold text-amber-300">{catalogCount} items</span>
            </div>

            {/* Test Sample Data */}
            <button
              id="btn-load-sample"
              type="button"
              onClick={onLoadSampleData}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-indigo-600/30 hover:bg-indigo-600/50 text-indigo-200 border border-indigo-500/40 transition-colors shadow-sm cursor-pointer"
              title="Load 3 sample Excel files for quick testing"
            >
              <Sparkles className="w-3.5 h-3.5 text-indigo-300" />
              <span>Load Sample Files</span>
            </button>

            {/* Download Master Template */}
            <button
              id="btn-download-template"
              type="button"
              onClick={downloadMasterTemplate}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors cursor-pointer"
              title="Download Master template (26 standard columns)"
            >
              <FileText className="w-3.5 h-3.5 text-slate-300" />
              <span className="hidden sm:inline">Master Template (26 Cols)</span>
            </button>

            {/* Catalog Manager Button */}
            <button
              id="btn-open-catalog-manager-header"
              type="button"
              onClick={onOpenCatalogManager}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-amber-300 border border-slate-700 transition-colors cursor-pointer"
              title="Manage ship equipment catalog"
            >
              <BookOpen className="w-3.5 h-3.5 text-amber-400" />
              <span>Catalog ({catalogCount})</span>
            </button>

            {/* Config Aliases */}
            <button
              id="btn-open-aliases"
              type="button"
              onClick={onOpenAliases}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-medium bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700 transition-colors cursor-pointer"
              title="Configure column header mapping rules"
            >
              <SlidersHorizontal className="w-3.5 h-3.5 text-slate-300" />
              <span className="hidden sm:inline">Column Aliases</span>
            </button>

            {/* Export Button */}
            <button
              id="btn-export-excel"
              type="button"
              onClick={onExportExcel}
              disabled={totalRows === 0}
              className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition-all ${
                totalRows > 0
                  ? 'bg-emerald-600 hover:bg-emerald-500 text-white cursor-pointer active:scale-95'
                  : 'bg-slate-800 text-slate-500 border border-slate-700 cursor-not-allowed opacity-60'
              }`}
              title="Export complete Armature Master Table to Excel (.xlsx)"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Excel (.xlsx)</span>
            </button>

            {/* Outlook Email Button */}
            <button
              id="btn-header-outlook-mail"
              type="button"
              onClick={onOpenOutlookMail}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold bg-blue-700 hover:bg-blue-600 text-white shadow-sm transition-all cursor-pointer"
              title="Compose and send report email via Outlook"
            >
              <Mail className="w-3.5 h-3.5" />
              <span>Outlook Mail</span>
            </button>

            {/* Help / Guide */}
            <button
              id="btn-open-help"
              type="button"
              onClick={onOpenHelp}
              className="p-1.5 rounded-lg text-slate-400 hover:text-slate-200 hover:bg-slate-800 transition-colors cursor-pointer"
              title="View specification rules and documentation"
            >
              <HelpCircle className="w-4 h-4" />
            </button>

            {/* Clear button */}
            {totalRows > 0 && (
              <button
                id="btn-clear-master"
                type="button"
                onClick={onClearData}
                className="p-1.5 rounded-lg text-rose-400 hover:text-rose-300 hover:bg-rose-950/40 transition-colors cursor-pointer"
                title="Clear all Armature Table data"
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
