import * as XLSX from 'xlsx';
import {
  normalizeHeader,
  getSourceColumnIndex,
  getHeaderOccurrence,
  DEFAULT_HEADER_ALIASES,
} from './headerNormalizer';
import { MasterRowData, SourceFileInfo, ImportLogRow } from '../types';

export const MAX_HEADER_SEARCH_ROW = 60;

export interface SheetDetectionResult {
  sheetName: string;
  headerRowIndex: number; // 0-based
  headers: string[];
}

/**
 * Checks a row array to see if it meets the criteria of an Armature source header:
 * Has TAG and SUPPLIER and (SFI or DESCRIPTION)
 */
export function isSourceHeaderRow(row: any[]): boolean {
  if (!row || !Array.isArray(row)) return false;

  let hasTag = false;
  let hasSupplier = false;
  let hasSFI = false;
  let hasDescription = false;

  const maxCols = Math.min(row.length, 150);

  for (let c = 0; c < maxCols; c++) {
    const val = normalizeHeader(row[c]);
    if (val === 'TAG') hasTag = true;
    else if (val === 'SUPPLIER') hasSupplier = true;
    else if (val === 'SFI') hasSFI = true;
    else if (val === 'DESCRIPTION') hasDescription = true;
  }

  return hasTag && hasSupplier && (hasSFI || hasDescription);
}

/**
 * Scans a 2D sheet data to find the header row index (0-based)
 */
export function findHeaderRowIndex(sheetData: any[][]): number {
  const scanLimit = Math.min(sheetData.length, MAX_HEADER_SEARCH_ROW);
  for (let r = 0; r < scanLimit; r++) {
    if (isSourceHeaderRow(sheetData[r])) {
      return r;
    }
  }
  return -1;
}

/**
 * Finds the suitable sheet and header row in a workbook
 */
export function findSourceSheetAndHeader(workbook: XLSX.WorkBook): {
  sheetName: string;
  headerRowIndex: number;
  sheetData: any[][];
} | null {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    const sheetData: any[][] = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      blankrows: false,
    });

    const headerIdx = findHeaderRowIndex(sheetData);
    if (headerIdx !== -1) {
      return { sheetName, headerRowIndex: headerIdx, sheetData };
    }
  }

  return null;
}

/**
 * Builds a header map: normalizedHeader -> list of 0-based column indices
 */
export function buildHeaderMap(headers: any[]): Map<string, number[]> {
  const map = new Map<string, number[]>();
  if (!headers) return map;

  for (let c = 0; c < headers.length; c++) {
    const norm = normalizeHeader(headers[c]);
    if (!norm) continue;

    if (!map.has(norm)) {
      map.set(norm, []);
    }
    map.get(norm)!.push(c);
  }

  return map;
}

/**
 * Check whether a row has at least one non-empty value in mapped source columns
 */
export function isDataRow(row: any[], mappedColIndices: number[]): boolean {
  if (!row || !Array.isArray(row)) return false;

  for (const colIdx of mappedColIndices) {
    if (colIdx < row.length) {
      const val = row[colIdx];
      if (val !== null && val !== undefined && String(val).trim().length > 0) {
        return true;
      }
    }
  }

  return false;
}

export interface ParseFileResult {
  fileInfo: SourceFileInfo;
  importedData: MasterRowData[];
  logRow: ImportLogRow;
}

/**
 * Parses a single Excel / CSV file against master columns & aliases
 */
export async function parseSourceFile(
  file: File,
  masterHeaders: string[],
  customAliases?: Record<string, string[]>,
  overrideSheetName?: string,
  overrideHeaderRow?: number
): Promise<ParseFileResult> {
  const nowStr = new Date().toLocaleString('vi-VN', {
    hour12: false,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  });

  const fileInfo: SourceFileInfo = {
    id: `${file.name}-${file.size}-${Date.now()}-${Math.random()}`,
    file,
    name: file.name,
    size: file.size,
    lastModified: file.lastModified,
    status: 'pending',
  };

  let workbook: XLSX.WorkBook;
  try {
    const buffer = await file.arrayBuffer();
    workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  } catch (err: any) {
    const logRow: ImportLogRow = {
      id: `log-${Date.now()}-${Math.random()}`,
      time: nowStr,
      sourceFile: file.name,
      sourceSheet: '',
      headerRow: 0,
      sourceRows: 0,
      importedRows: 0,
      skippedRows: 0,
      status: `ERROR - CANNOT OPEN FILE: ${err.message || 'Corrupted or unsupported format'}`,
    };
    fileInfo.status = 'error';
    fileInfo.statusMessage = logRow.status;
    return { fileInfo, importedData: [], logRow };
  }

  fileInfo.availableSheets = workbook.SheetNames;

  let chosenSheetName = overrideSheetName;
  let headerRowIndex = overrideHeaderRow !== undefined ? overrideHeaderRow : -1;
  let sheetData: any[][] = [];

  if (chosenSheetName && workbook.Sheets[chosenSheetName]) {
    sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[chosenSheetName], {
      header: 1,
      defval: '',
      blankrows: false,
    });
    if (headerRowIndex < 0) {
      headerRowIndex = findHeaderRowIndex(sheetData);
    }
  } else {
    const detected = findSourceSheetAndHeader(workbook);
    if (detected) {
      chosenSheetName = detected.sheetName;
      headerRowIndex = detected.headerRowIndex;
      sheetData = detected.sheetData;
    } else if (workbook.SheetNames.length > 0) {
      chosenSheetName = workbook.SheetNames[0];
      sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[chosenSheetName], {
        header: 1,
        defval: '',
        blankrows: false,
      });
      headerRowIndex = findHeaderRowIndex(sheetData);
    }
  }

  if (!chosenSheetName || !workbook.Sheets[chosenSheetName]) {
    const logRow: ImportLogRow = {
      id: `log-${Date.now()}-${Math.random()}`,
      time: nowStr,
      sourceFile: file.name,
      sourceSheet: '',
      headerRow: 0,
      sourceRows: 0,
      importedRows: 0,
      skippedRows: 0,
      status: 'ERROR - SOURCE SHEET NOT FOUND',
    };
    fileInfo.status = 'error';
    fileInfo.statusMessage = logRow.status;
    return { fileInfo, importedData: [], logRow };
  }

  fileInfo.detectedSheet = chosenSheetName;

  if (headerRowIndex < 0) {
    const logRow: ImportLogRow = {
      id: `log-${Date.now()}-${Math.random()}`,
      time: nowStr,
      sourceFile: file.name,
      sourceSheet: chosenSheetName,
      headerRow: 0,
      sourceRows: sheetData.length,
      importedRows: 0,
      skippedRows: sheetData.length,
      status: 'ERROR - HEADER ROW NOT FOUND (Requires TAG, SUPPLIER, and SFI/DESCRIPTION)',
    };
    fileInfo.status = 'error';
    fileInfo.statusMessage = logRow.status;
    return { fileInfo, importedData: [], logRow };
  }

  // 1-based display header row (like Excel row 2 or 3)
  fileInfo.headerRow = headerRowIndex + 1;

  const rawHeaders = sheetData[headerRowIndex] || [];
  const sourceMap = buildHeaderMap(rawHeaders);

  // Check mapping against master headers
  const missingFields: string[] = [];
  const matchedColumns: { master: string; source: string; colIndex: number }[] = [];
  const mappedSourceColIndices: number[] = [];

  for (let mIdx = 0; mIdx < masterHeaders.length; mIdx++) {
    const masterHeader = masterHeaders[mIdx];
    const occurrence = getHeaderOccurrence(masterHeaders, mIdx);
    const sourceCol = getSourceColumnIndex(
      masterHeader,
      occurrence,
      sourceMap,
      customAliases
    );

    if (sourceCol !== null) {
      const sourceHeaderName = String(rawHeaders[sourceCol] || '');
      matchedColumns.push({
        master: masterHeader,
        source: sourceHeaderName,
        colIndex: sourceCol,
      });
      mappedSourceColIndices.push(sourceCol);
    } else {
      missingFields.push(masterHeader);
    }
  }

  fileInfo.matchedColumns = matchedColumns;
  fileInfo.missingFields = missingFields;

  const dataRows = sheetData.slice(headerRowIndex + 1);
  const totalSourceRows = dataRows.length;
  fileInfo.sourceRows = totalSourceRows;

  if (totalSourceRows === 0) {
    const logRow: ImportLogRow = {
      id: `log-${Date.now()}-${Math.random()}`,
      time: nowStr,
      sourceFile: file.name,
      sourceSheet: chosenSheetName,
      headerRow: headerRowIndex + 1,
      sourceRows: 0,
      importedRows: 0,
      skippedRows: 0,
      status: 'WARNING - NO DATA',
    };
    fileInfo.status = 'warning';
    fileInfo.statusMessage = logRow.status;
    return { fileInfo, importedData: [], logRow };
  }

  // Extract valid rows
  const importedData: MasterRowData[] = [];
  let skippedRows = 0;

  for (let r = 0; r < dataRows.length; r++) {
    const row = dataRows[r];

    if (isDataRow(row, mappedSourceColIndices)) {
      const masterRow: MasterRowData = {
        _id: `${file.name}-${r}-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
        _sourceFile: file.name,
        _sourceSheet: chosenSheetName,
        _importedAt: nowStr,
      };

      // Map each master column
      for (let mIdx = 0; mIdx < masterHeaders.length; mIdx++) {
        const masterHeader = masterHeaders[mIdx];
        const occurrence = getHeaderOccurrence(masterHeaders, mIdx);
        const sourceCol = getSourceColumnIndex(
          masterHeader,
          occurrence,
          sourceMap,
          customAliases
        );

        if (sourceCol !== null && sourceCol < row.length) {
          const rawVal = row[sourceCol];
          masterRow[masterHeader] = rawVal !== undefined && rawVal !== null ? rawVal : '';
        } else {
          masterRow[masterHeader] = '';
        }
      }

      importedData.push(masterRow);
    } else {
      skippedRows++;
    }
  }

  fileInfo.importedRows = importedData.length;
  fileInfo.skippedRows = skippedRows;

  let statusStr = 'OK';
  if (missingFields.length > 0) {
    statusStr = `OK WITH MISSING FIELDS: ${missingFields.join(', ')}`;
    fileInfo.status = 'warning';
  } else {
    fileInfo.status = 'success';
  }
  fileInfo.statusMessage = statusStr;

  const logRow: ImportLogRow = {
    id: `log-${Date.now()}-${Math.random()}`,
    time: nowStr,
    sourceFile: file.name,
    sourceSheet: chosenSheetName,
    headerRow: headerRowIndex + 1,
    sourceRows: totalSourceRows,
    importedRows: importedData.length,
    skippedRows,
    status: statusStr,
    warnings: missingFields,
  };

  return { fileInfo, importedData, logRow };
}
