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

export interface CatalogItem {
  id: string;
  'SAP Name'?: string;
  Name?: string;
  Description?: string;
  'NRF Nr.'?: string;
  'Normale Nr.'?: string;
  Execution?: string;
  Manufacturer?: string;
  'Model Number'?: string;
  'Used In HVAC Generator'?: string;
  'Pipe Size'?: string;
  Connection?: string;
  'Pressure Nominal'?: string;
  'Pipe Class'?: string;
  'Body Material'?: string;
  'Disc Material'?: string;
  'Seat Material'?: string;
  'Stem Material'?: string;
  'Stem Height [mm]'?: string | number;
  'Stem Size'?: string;
  'Stem Shape'?: string;
  'ISO 5211'?: string;
  'Means Of Operation'?: string;
  'Operating Temp Min. [degC]'?: string | number;
  'Operating Temp Max. [degC]'?: string | number;
  'Setpoint Min'?: string | number;
  'Setpoint Max'?: string | number;
  'Dry Weight [kg]'?: string | number;
  'Typical Usage'?: string;
  'Testing Certificate'?: string;
  'MRP Type'?: string;
  'Sign Type'?: string;
  Comment?: string;
  [key: string]: any;
}

export interface CatalogProfile {
  id: string;
  shipName: string;
  versionName: string;
  uploadedAt: string;
  sourceFileName: string;
  items: CatalogItem[];
}
