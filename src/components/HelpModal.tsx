import React from 'react';
import { X, CheckCircle2, HelpCircle, FileSpreadsheet, Code, ShieldCheck } from 'lucide-react';

interface HelpModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const HelpModal: React.FC<HelpModalProps> = ({ isOpen, onClose }) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-2xl w-full border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* Modal Header */}
        <div className="px-6 py-4 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-100 dark:bg-blue-950/60 text-blue-600 dark:text-blue-400 flex items-center justify-center">
              <Code className="w-4 h-4" />
            </div>
            <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
              Operating Principles &amp; VBA Macro Comparison
            </h2>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Content */}
        <div className="p-6 overflow-y-auto space-y-5 text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
          <div className="bg-blue-50 dark:bg-blue-950/30 border border-blue-200 dark:border-blue-900/40 rounded-xl p-4 text-blue-900 dark:text-blue-200">
            <p className="font-semibold text-sm mb-1">
              Web Application faithfully converted 100% from ARMATURE IMPORT TOOL Macro
            </p>
            <p>
              Import dozens of supplier Excel armature valve sheets simultaneously into the Master table without opening Excel, avoiding version conflicts, and running directly in your browser!
            </p>
          </div>

          <div className="space-y-3">
            <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-500" />
              1. Header Detection Rules
            </h3>
            <p>
              The system automatically scans up to the first 60 rows of each sheet in source files. A row is recognized as a header when it contains simultaneously:
            </p>
            <ul className="list-disc list-inside ml-2 space-y-1 font-mono text-slate-800 dark:text-slate-200">
              <li>Column containing <strong>TAG</strong></li>
              <li>Column containing <strong>SUPPLIER</strong></li>
              <li>Column containing <strong>SFI</strong> or <strong>DESCRIPTION</strong></li>
            </ul>
          </div>

          <div className="space-y-3">
            <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-500" />
              2. Title Normalization &amp; Alias Lookup
            </h3>
            <p>
              Before matching, headers are thoroughly cleaned: removes line breaks (\r, \n), tabs, non-breaking space (Chr 160), dots, hyphens, slashes, degree symbols (°), converts to UPPERCASE, and strips duplicate spaces.
            </p>
            <p>
              If column names do not match directly, the system automatically checks the Alias table (e.g. <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">PO NO</code> &rarr; <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">PO NUMBER</code>, <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">PRESSURE CLASS</code> &rarr; <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">PRESSURE RATING</code>, etc.).
            </p>
          </div>

          <div className="space-y-3">
            <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-500" />
              3. Data Row Validation (IsDataRow) &amp; Occurrence Matching
            </h3>
            <p>
              Only rows with at least one non-empty value in mapped columns are imported into Master (automatically filtering out blank or comment rows).
            </p>
            <p>
              If a header appears multiple times (e.g. <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">DESCRIPTION #1</code>, <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">DESCRIPTION #2</code>), the algorithm matches them sequentially by occurrence.
            </p>
          </div>

          <div className="space-y-3">
            <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-500" />
              4. Standardized Excel Export Structure
            </h3>
            <p>
              Clicking <strong>&quot;Export Excel (.xlsx)&quot;</strong> downloads a file with:
            </p>
            <ul className="list-disc list-inside ml-2 space-y-1">
              <li><strong>Sheet &quot;Total&quot;</strong>: Contains all merged Armature data, with row 2 as standard Header.</li>
              <li><strong>Sheet &quot;Import_Log&quot;</strong>: Contains complete timestamped logs, source file names, source sheet names, header rows, imported row counts, and error details.</li>
            </ul>
          </div>
        </div>

        {/* Modal Footer */}
        <div className="px-6 py-3.5 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-2 rounded-lg text-xs font-semibold bg-blue-600 hover:bg-blue-500 text-white shadow-sm transition-colors"
          >
            Got it, continue
          </button>
        </div>
      </div>
    </div>
  );
};
