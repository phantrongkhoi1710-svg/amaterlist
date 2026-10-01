import React, { useState, useMemo, useRef } from 'react';
import {
  Upload,
  Download,
  Search,
  Ship,
  FileSpreadsheet,
  Trash2,
  Plus,
  RefreshCw,
  CheckCircle2,
  SlidersHorizontal,
  ChevronLeft,
  ChevronRight,
  Database,
  ArrowRight,
  Sparkles,
} from 'lucide-react';
import { CatalogItem, CatalogProfile } from '../types';
import {
  CATALOG_COLUMNS,
  downloadCatalogTemplate,
  exportCatalogToExcel,
  parseCatalogFile,
} from '../utils/catalogManager';

interface CatalogViewerProps {
  catalogProfiles: CatalogProfile[];
  activeProfileId: string;
  setActiveProfileId: (id: string) => void;
  onSaveProfile: (profile: CatalogProfile) => void;
  onDeleteProfile: (id: string) => void;
  onAutoFillMaster?: (catalogItems: CatalogItem[]) => void;
  onExportToMaster?: (catalogItems: CatalogItem[], mode: 'replace' | 'append') => void;
}

export const CatalogViewer: React.FC<CatalogViewerProps> = ({
  catalogProfiles,
  activeProfileId,
  setActiveProfileId,
  onSaveProfile,
  onDeleteProfile,
  onAutoFillMaster,
  onExportToMaster,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [uploadSuccess, setUploadSuccess] = useState<string | null>(null);

  // New catalog form modal / inputs
  const [isCreatingNew, setIsCreatingNew] = useState(false);
  const [newShipName, setNewShipName] = useState('');
  const [newVersionName, setNewVersionName] = useState('Month ' + (new Date().getMonth() + 1) + '/' + new Date().getFullYear());

  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // Active catalog profile
  const activeProfile = useMemo(() => {
    return (
      catalogProfiles.find((p) => p.id === activeProfileId) ||
      catalogProfiles[0] ||
      null
    );
  }, [catalogProfiles, activeProfileId]);

  const items = activeProfile?.items || [];

  // Filter items
  const filteredItems = useMemo(() => {
    if (!searchTerm.trim()) return items;
    const term = searchTerm.toLowerCase().trim();
    return items.filter((item) => {
      return (
        (item['SAP Name'] && String(item['SAP Name']).toLowerCase().includes(term)) ||
        (item.Name && String(item.Name).toLowerCase().includes(term)) ||
        (item.Description && String(item.Description).toLowerCase().includes(term)) ||
        (item['Normale Nr.'] && String(item['Normale Nr.']).toLowerCase().includes(term)) ||
        (item['Model Number'] && String(item['Model Number']).toLowerCase().includes(term)) ||
        (item.Manufacturer && String(item.Manufacturer).toLowerCase().includes(term)) ||
        (item['Body Material'] && String(item['Body Material']).toLowerCase().includes(term)) ||
        (item['Pipe Size'] && String(item['Pipe Size']).toLowerCase().includes(term)) ||
        (item['Connection'] && String(item['Connection']).toLowerCase().includes(term)) ||
        (item['Pressure Nominal'] && String(item['Pressure Nominal']).toLowerCase().includes(term))
      );
    });
  }, [items, searchTerm]);

  // Pagination
  const totalPages = Math.max(1, Math.ceil(filteredItems.length / pageSize));
  const paginatedItems = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredItems.slice(start, start + pageSize);
  }, [filteredItems, currentPage, pageSize]);

  // File Upload Handler
  const handleFileUpload = async (file: File) => {
    setUploadError(null);
    setUploadSuccess(null);
    setIsUploading(true);

    try {
      const { items: parsedItems, sheetName, totalRows } = await parseCatalogFile(file);

      if (parsedItems.length === 0) {
        throw new Error('No valid Catalog items found in file.');
      }

      // Auto derive ship name from file name if creating new
      const derivedShipName =
        newShipName.trim() ||
        activeProfile?.shipName ||
        file.name.replace(/\.[^/.]+$/, '').replace(/_/g, ' ');

      const derivedVersion =
        newVersionName.trim() ||
        activeProfile?.versionName ||
        'Rev ' + new Date().toLocaleDateString('en-US');

      const updatedProfile: CatalogProfile = {
        id: activeProfile?.id && !isCreatingNew ? activeProfile.id : `cat-prof-${Date.now()}`,
        shipName: derivedShipName,
        versionName: derivedVersion,
        uploadedAt: new Date().toLocaleString('en-US'),
        sourceFileName: file.name,
        items: parsedItems,
      };

      onSaveProfile(updatedProfile);
      setActiveProfileId(updatedProfile.id);
      setIsCreatingNew(false);
      setUploadSuccess(
        `Successfully uploaded ${totalRows} catalog items for ${derivedShipName} (Sheet: ${sheetName})!`
      );
    } catch (err: any) {
      console.error(err);
      setUploadError(err.message || 'Error reading Catalog file.');
    } finally {
      setIsUploading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileUpload(e.dataTransfer.files[0]);
    }
  };

  return (
    <div className="space-y-4">
      {/* Top Bar: Profile Selector & Quick Action Buttons */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 p-3.5 shadow-sm">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
          {/* Left: Ship Profile Selector */}
          <div className="flex flex-wrap items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-600/10 dark:bg-blue-500/20 text-blue-600 dark:text-blue-400 flex items-center justify-center shrink-0">
              <Ship className="w-4 h-4" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-semibold text-slate-700 dark:text-slate-300">
                  Ship / Project Catalog:
                </span>
                <select
                  id="catalog-profile-select"
                  value={activeProfile?.id || ''}
                  onChange={(e) => {
                    setActiveProfileId(e.target.value);
                    setCurrentPage(1);
                  }}
                  className="px-2.5 py-1 text-xs font-medium rounded-lg border border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-800 text-slate-800 dark:text-slate-200 focus:outline-none focus:ring-1 focus:ring-blue-500"
                >
                  {catalogProfiles.map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.shipName} ({p.versionName}) - {p.items.length} items
                    </option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={() => {
                    setIsCreatingNew(true);
                    setNewShipName('');
                  }}
                  className="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium text-blue-600 dark:text-blue-400 hover:bg-blue-50 dark:hover:bg-blue-950/40 rounded-md transition-colors cursor-pointer"
                  title="Add catalog for a new ship or project"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>New Ship</span>
                </button>
              </div>
              <div className="text-[11px] text-slate-400 flex items-center gap-2 mt-0.5">
                <span>Updated: {activeProfile?.uploadedAt || 'Default'}</span>
                <span>&bull;</span>
                <span>File: {activeProfile?.sourceFileName || 'Sample Data'}</span>
              </div>
            </div>
          </div>

          {/* Right: Actions */}
          <div className="flex flex-wrap items-center gap-2">
            {/* Export all to Armature Master Table */}
            {onExportToMaster && (
              <button
                id="btn-catalog-export-to-master"
                type="button"
                onClick={() => onExportToMaster(items, 'append')}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-blue-600 hover:bg-blue-500 text-white shadow-sm transition-colors cursor-pointer"
                title="Convert all catalog items into standard 26-column Armature List rows"
              >
                <ArrowRight className="w-3.5 h-3.5" />
                <span>Export Catalog to Armature List</span>
              </button>
            )}

            {/* Auto-fill Master */}
            <button
              id="btn-catalog-autofill-master"
              type="button"
              onClick={() => onAutoFillMaster && onAutoFillMaster(items)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-semibold rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white shadow-sm transition-colors cursor-pointer"
              title="Match TAG / Model / Normale Nr to auto-fill technical specs in Master Table"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>Auto-Fill Specs into Master</span>
            </button>

            {/* Download Blank Template */}
            <button
              id="btn-download-catalog-blank"
              type="button"
              onClick={() => downloadCatalogTemplate(true, activeProfile?.shipName)}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors cursor-pointer"
              title="Download blank 31-column Excel catalog template"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
              <span>Blank Catalog Template</span>
            </button>

            {/* Download Sample Template */}
            <button
              id="btn-download-catalog-sample"
              type="button"
              onClick={() => downloadCatalogTemplate(false, activeProfile?.shipName)}
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-700 transition-colors cursor-pointer"
              title="Download sample Excel catalog with example items"
            >
              <Download className="w-3.5 h-3.5 text-blue-500" />
              <span>Sample Catalog Template</span>
            </button>

            {/* Export Current Catalog */}
            <button
              id="btn-export-catalog"
              type="button"
              onClick={() =>
                exportCatalogToExcel(
                  items,
                  activeProfile?.shipName || 'Catalog',
                  activeProfile?.versionName || 'Rev1'
                )
              }
              className="inline-flex items-center gap-1.5 px-2.5 py-1.5 text-xs font-medium rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 transition-colors cursor-pointer"
              title="Export current catalog to Excel (.xlsx)"
            >
              <Download className="w-3.5 h-3.5" />
              <span>Export Excel</span>
            </button>

            {/* Delete profile */}
            {catalogProfiles.length > 1 && (
              <button
                type="button"
                onClick={() => {
                  if (
                    confirm(
                      `Are you sure you want to delete Catalog "${activeProfile?.shipName}"?`
                    )
                  ) {
                    onDeleteProfile(activeProfile.id);
                  }
                }}
                className="p-1.5 text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-950/40 rounded-lg transition-colors cursor-pointer"
                title="Delete this ship catalog"
              >
                <Trash2 className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>

        {/* Modal / Inline form for creating new ship catalog */}
        {isCreatingNew && (
          <div className="mt-3 pt-3 border-t border-slate-200 dark:border-slate-800 flex flex-wrap items-center gap-3 bg-slate-50 dark:bg-slate-800/60 p-2.5 rounded-lg">
            <span className="text-xs font-semibold text-slate-700 dark:text-slate-300">
              Create new ship catalog:
            </span>
            <input
              type="text"
              placeholder="Ship / Project Name (e.g. Vessel H305, Bulk Carrier 64K...)"
              value={newShipName}
              onChange={(e) => setNewShipName(e.target.value)}
              className="px-2.5 py-1 text-xs rounded-md border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-200 w-64 focus:outline-none focus:ring-1 focus:ring-blue-500"
            />
            <input
              type="text"
              placeholder="Version / Date (e.g. Rev 2, Mar 2026...)"
              value={newVersionName}
              onChange={(e) => setNewVersionName(e.target.value)}
              className="px-2.5 py-1 text-xs rounded-md border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-200 w-44 focus:outline-none focus:ring-1 focus:ring-blue-500"
            />
            <span className="text-xs text-slate-500">
              (Then drag &amp; drop or select an Excel file below to load)
            </span>
            <button
              type="button"
              onClick={() => setIsCreatingNew(false)}
              className="px-2 py-1 text-xs text-slate-500 hover:text-slate-700 cursor-pointer ml-auto"
            >
              Close
            </button>
          </div>
        )}
      </div>

      {/* Upload Zone */}
      <div
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
        className={`bg-white dark:bg-slate-900 rounded-xl border-2 border-dashed p-4 text-center cursor-pointer transition-all ${
          isDragging
            ? 'border-blue-500 bg-blue-50/50 dark:bg-blue-950/20 shadow-md'
            : 'border-slate-300 dark:border-slate-700 hover:border-blue-400 dark:hover:border-blue-600'
        }`}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls,.xlsm,.xlsb,.csv"
          className="hidden"
          onChange={(e) => {
            if (e.target.files && e.target.files.length > 0) {
              handleFileUpload(e.target.files[0]);
            }
          }}
        />
        <div className="flex items-center justify-center gap-3">
          <div className="w-8 h-8 rounded-lg bg-blue-50 dark:bg-blue-950/50 text-blue-600 dark:text-blue-400 flex items-center justify-center shrink-0">
            {isUploading ? (
              <RefreshCw className="w-4 h-4 animate-spin text-blue-500" />
            ) : (
              <Upload className="w-4 h-4" />
            )}
          </div>
          <div className="text-left">
            <p className="text-xs font-semibold text-slate-800 dark:text-slate-100">
              {isUploading
                ? 'Analyzing and loading Catalog data...'
                : `Upload Catalog file for ${activeProfile?.shipName || 'this ship'}`}
            </p>
            <p className="text-[11px] text-slate-400">
              Drag &amp; drop <strong>.xlsx, .xls, .xlsm, .csv</strong> files here or click to browse from computer
            </p>
          </div>
        </div>
      </div>

      {/* Toast / Status messages */}
      {uploadSuccess && (
        <div className="p-2.5 rounded-lg bg-emerald-50 dark:bg-emerald-950/40 text-emerald-800 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800 flex items-center justify-between text-xs shadow-xs">
          <div className="flex items-center gap-2">
            <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
            <span className="font-semibold">{uploadSuccess}</span>
          </div>
          <button
            type="button"
            onClick={() => setUploadSuccess(null)}
            className="text-slate-400 hover:text-slate-600"
          >
            &times;
          </button>
        </div>
      )}

      {uploadError && (
        <div className="p-2.5 rounded-lg bg-rose-50 dark:bg-rose-950/40 text-rose-800 dark:text-rose-300 border border-rose-200 dark:border-rose-800 flex items-center justify-between text-xs shadow-xs">
          <span className="font-semibold">{uploadError}</span>
          <button
            type="button"
            onClick={() => setUploadError(null)}
            className="text-slate-400 hover:text-slate-600"
          >
            &times;
          </button>
        </div>
      )}

      {/* Catalog Data Table */}
      <div className="bg-white dark:bg-slate-900 rounded-xl border border-slate-200 dark:border-slate-800 shadow-sm overflow-hidden flex flex-col flex-1">
        {/* Table Controls */}
        <div className="p-3 border-b border-slate-200 dark:border-slate-800 flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-slate-50/70 dark:bg-slate-800/40">
          <div className="flex items-center gap-2.5 flex-1 max-w-md">
            <div className="relative w-full">
              <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input
                id="search-catalog-input"
                type="text"
                placeholder="Search by SAP Name, Normale Nr, Description, Material, Model..."
                value={searchTerm}
                onChange={(e) => {
                  setSearchTerm(e.target.value);
                  setCurrentPage(1);
                }}
                className="w-full pl-8 pr-3 py-1.5 text-xs rounded-lg border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-200 placeholder-slate-400 focus:outline-none focus:ring-1 focus:ring-blue-500"
              />
            </div>
            {searchTerm && (
              <button
                type="button"
                onClick={() => setSearchTerm('')}
                className="text-xs text-slate-400 hover:text-slate-600 whitespace-nowrap cursor-pointer"
              >
                Clear search
              </button>
            )}
          </div>

          <div className="flex items-center gap-3 text-xs text-slate-600 dark:text-slate-400">
            <span>
              Showing <strong>{filteredItems.length}</strong> / <strong>{items.length}</strong> items
            </span>
            <select
              value={pageSize}
              onChange={(e) => {
                setPageSize(Number(e.target.value));
                setCurrentPage(1);
              }}
              className="px-2 py-1 text-xs rounded border border-slate-300 dark:border-slate-700 bg-white dark:bg-slate-900 text-slate-800 dark:text-slate-200"
            >
              <option value={25}>25 rows/page</option>
              <option value={50}>50 rows/page</option>
              <option value={100}>100 rows/page</option>
              <option value={500}>500 rows/page</option>
            </select>
          </div>
        </div>

        {/* Maximized Scrollable Table */}
        <div className="overflow-x-auto max-h-[calc(100vh-270px)] min-h-[480px]">
          <table className="w-full text-left text-xs text-slate-700 dark:text-slate-300 border-collapse">
            <thead className="bg-slate-100 dark:bg-slate-800/95 text-slate-800 dark:text-slate-200 sticky top-0 z-20 border-b border-slate-200 dark:border-slate-700 font-semibold shadow-xs">
              <tr>
                <th className="py-2.5 px-3 w-10 text-center text-slate-400 font-mono">#</th>
                {CATALOG_COLUMNS.map((col) => {
                  const isCore = ['SAP Name', 'Name', 'Description', 'Normale Nr.', 'Model Number', 'Pipe Size'].includes(col);
                  return (
                    <th
                      key={col}
                      className={`py-2.5 px-3 whitespace-nowrap border-r border-slate-200/60 dark:border-slate-700/60 ${
                        isCore ? 'font-bold text-blue-900 dark:text-blue-300' : 'text-slate-700 dark:text-slate-300'
                      }`}
                    >
                      {col}
                    </th>
                  );
                })}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-800 font-normal">
              {paginatedItems.length === 0 ? (
                <tr>
                  <td
                    colSpan={CATALOG_COLUMNS.length + 1}
                    className="py-16 text-center text-slate-400"
                  >
                    <div className="flex flex-col items-center justify-center gap-2">
                      <Database className="w-8 h-8 text-slate-300 dark:text-slate-600" />
                      <p className="text-sm font-medium">
                        {searchTerm ? 'No matching valve items found' : 'No Catalog data for this ship yet'}
                      </p>
                      <p className="text-xs text-slate-400">
                        {searchTerm
                          ? 'Try searching with a different keyword'
                          : 'Upload an Excel Catalog file above to get started'}
                      </p>
                    </div>
                  </td>
                </tr>
              ) : (
                paginatedItems.map((item, idx) => {
                  const rowIndex = (currentPage - 1) * pageSize + idx + 1;
                  return (
                    <tr
                      key={item.id || idx}
                      className="hover:bg-blue-50/40 dark:hover:bg-blue-950/20 transition-colors"
                    >
                      <td className="py-2 px-3 text-center text-slate-400 font-mono text-[11px] bg-slate-50/50 dark:bg-slate-900/50">
                        {rowIndex}
                      </td>
                      {CATALOG_COLUMNS.map((col) => {
                        const val = item[col];
                        const isSapName = col === 'SAP Name' || col === 'Name';
                        const displayVal =
                          val !== undefined && val !== null && String(val).trim() !== ''
                            ? String(val)
                            : (col === 'SAP Name' ? (item.Name || '-') : '-');
                        return (
                          <td
                            key={col}
                            className={`py-2 px-3 whitespace-nowrap border-r border-slate-100 dark:border-slate-800/60 ${
                              isSapName ? 'font-mono font-bold text-blue-700 dark:text-blue-400 bg-blue-50/20 dark:bg-blue-950/20' : ''
                            }`}
                          >
                            {displayVal}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {/* Pagination Bar */}
        {totalPages > 1 && (
          <div className="p-2.5 border-t border-slate-200 dark:border-slate-800 flex items-center justify-between bg-slate-50/50 dark:bg-slate-900/50 text-xs">
            <span className="text-slate-500">
              Page {currentPage} / {totalPages}
            </span>
            <div className="flex items-center gap-1.5">
              <button
                type="button"
                disabled={currentPage === 1}
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                className="p-1 rounded border border-slate-300 dark:border-slate-700 disabled:opacity-40 hover:bg-slate-100 dark:hover:bg-slate-800"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>
              <button
                type="button"
                disabled={currentPage === totalPages}
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                className="p-1 rounded border border-slate-300 dark:border-slate-700 disabled:opacity-40 hover:bg-slate-100 dark:hover:bg-slate-800"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};
