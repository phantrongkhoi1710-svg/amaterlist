export interface MasterColumn {
  id: string;
  name: string;
  normalized: string;
  isCustom?: boolean;
}

export interface AliasMapping {
  masterHeader: string;
  aliases: string[];
}

export interface ImportLogRow {
  id: string;
  time: string;
  sourceFile: string;
  sourceSheet: string;
  headerRow: number;
  sourceRows: number;
  importedRows: number;
  skippedRows: number;
  status: string;
  warnings?: string[];
}

export interface MasterRowData {
  _id: string;
  _sourceFile: string;
  _sourceSheet: string;
  _importedAt: string;
  [key: string]: any;
}

export interface SourceFileInfo {
  id: string;
  file: File;
  name: string;
  size: number;
  lastModified: number;
  detectedSheet?: string;
  availableSheets?: string[];
  headerRow?: number;
  sourceRows?: number;
  importedRows?: number;
  skippedRows?: number;
  status?: 'pending' | 'processing' | 'success' | 'warning' | 'error' | 'skipped';
  statusMessage?: string;
  missingFields?: string[];
  matchedColumns?: { master: string; source: string; colIndex: number }[];
}

export type ImportMode = 'replace' | 'append';
