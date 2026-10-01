import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import {
  normalizeHeader,
  getSourceColumnIndex,
  getHeaderOccurrence,
  DEFAULT_HEADER_ALIASES,
} from './headerNormalizer';
import { MasterRowData, SourceFileInfo, ImportLogRow } from '../types';
import { isYellowColor, isRedColor, extractHexColor } from './revisionManager';

export const MAX_HEADER_SEARCH_ROW = 60;

export interface SheetDetectionResult {
  sheetName: string;
  headerRowIndex: number; // 0-based
  headers: string[];
}

/**
 * Checks a row array to see if it meets the criteria of an Armature source header or SAP Name Catalog:
 * 1. Standard Armature header: Has TAG and (SUPPLIER or SFI or DESCRIPTION or SAP CODE)
 * 2. SAP Name / Catalog header: Has SAP Name/Code and at least one spec column (Description, Normale Nr, Pipe Size, Manufacturer, Body Material, etc.)
 * 3. General header: Has 3 or more recognized columns from Armature master or Catalog specs
 */
export function isSourceHeaderRow(row: any[]): boolean {
  if (!row || !Array.isArray(row)) return false;

  let hasTag = false;
  let hasSupplier = false;
  let hasSFI = false;
  let hasDescription = false;
  let hasSapCodeOrName = false;
  let hasNormaleNr = false;
  let hasPipeSize = false;
  let hasPressure = false;
  let hasBodyMaterial = false;
  let hasModelNumber = false;
  let matchCount = 0;

  const maxCols = Math.min(row.length, 150);

  const sapNameAliases = [
    'SAP', 'SAP CODE', 'SAP NAME', 'SAP_NAME', 'SAPNAME', 'SAPCODE', 'SAP-NAME',
    'MATERIAL CODE', 'MATERIAL NO', 'MATERIAL NUMBER', 'MATERIAL', 'PART NUMBER', 'PART NO', 'VALVE CODE', 'VALVE NAME'
  ];

  const normaleAliases = [
    'NORMALE NR', 'NORMALE NO', 'NORMALE N', 'NORMALE N°', 'NORMALE', 'STD DRAWING', 'STANDARD DRAWING', 'STD DRW NORMALE'
  ];

  const sizeAliases = ['PIPE SIZE', 'DN', 'DIAMETER', 'VALVE SIZE', 'SIZE', 'NOMINAL SIZE'];
  const pressureAliases = ['PRESSURE', 'PRESSURE NOMINAL', 'PRESSURE CLASS', 'PRESSURE RATING', 'PN', 'RATING'];
  const bodyAliases = ['BODY MATERIAL', 'MATERIAL BODY', 'HOUSING /BODY', 'HOUSING/BODY', 'HOUSING BODY', 'BODY', 'BODY MAT'];

  for (let c = 0; c < maxCols; c++) {
    const val = normalizeHeader(row[c]);
    if (!val) continue;

    if (val === 'TAG' || val === 'TAG NO' || val === 'VALVE TAG' || val === 'TAG NUMBER') hasTag = true;
    if (val === 'SUPPLIER' || val === 'MANUFACTURER' || val === 'VENDOR' || val === 'MAKER') hasSupplier = true;
    if (val === 'SFI' || val === 'SFI CODE' || val === 'SYSTEM') hasSFI = true;
    if (val === 'DESCRIPTION' || val === 'DESC' || val === 'VALVE DESCRIPTION' || val === 'ITEM DESCRIPTION') hasDescription = true;

    if (sapNameAliases.includes(val)) {
      hasSapCodeOrName = true;
      matchCount++;
    } else if (normaleAliases.includes(val)) {
      hasNormaleNr = true;
      matchCount++;
    } else if (sizeAliases.includes(val)) {
      hasPipeSize = true;
      matchCount++;
    } else if (pressureAliases.includes(val)) {
      hasPressure = true;
      matchCount++;
    } else if (bodyAliases.includes(val)) {
      hasBodyMaterial = true;
      matchCount++;
    } else if (val === 'MODEL NUMBER' || val === 'MODEL' || val === 'MODEL NO' || val === 'MODEL NR') {
      hasModelNumber = true;
      matchCount++;
    } else if (val === 'EXECUTION' || val === 'CONNECTION' || val === 'NRF N' || val === 'NRF NR' || val === 'PIPE CLASS' || val === 'REMARKS') {
      matchCount++;
    }
  }

  // Case 1: Standard Armature List with TAG
  if (hasTag && (hasSupplier || hasSFI || hasDescription || hasSapCodeOrName)) {
    return true;
  }

  // Case 2: SAP Name / Catalog Export (doesn't have TAG column yet)
  if (
    hasSapCodeOrName &&
    (hasDescription || hasNormaleNr || hasPipeSize || hasPressure || hasBodyMaterial || hasSupplier || hasModelNumber)
  ) {
    return true;
  }

  // Case 3: Recognized at least 3 distinct pipe specification/catalog columns
  if (matchCount >= 3) {
    return true;
  }

  return false;
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

  if (mappedColIndices.length > 0) {
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

  // Fallback if mappedColIndices is empty: check if any cell has meaningful text
  return row.some((val) => val !== null && val !== undefined && String(val).trim().length > 0);
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
      status: 'ERROR - HEADER ROW NOT FOUND (Requires TAG, SUPPLIER, and SFI/DESCRIPTION or SAP Name/Catalog headers)',
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

  // Extract cell styling information (colors and strikethroughs) using ExcelJS for .xlsx/.xlsm files
  const rowStylesMap = new Map<number, { isYellow: boolean; isRed: boolean; hasStrike: boolean; customColor?: string }>();
  const rowStylesByTagMap = new Map<string, { isYellow: boolean; isRed: boolean; hasStrike: boolean; customColor?: string }>();

  if (file.name.toLowerCase().endsWith('.xlsx') || file.name.toLowerCase().endsWith('.xlsm')) {
    try {
      const buffer = await file.arrayBuffer();
      const excelWorkbook = new ExcelJS.Workbook();
      await excelWorkbook.xlsx.load(buffer);
      const ws =
        excelWorkbook.getWorksheet(chosenSheetName) ||
        excelWorkbook.worksheets.find(
          (s) => s.name.trim().toLowerCase() === chosenSheetName.trim().toLowerCase()
        ) ||
        excelWorkbook.worksheets[0];

      if (ws) {
        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          let hasYellow = false;
          let hasRed = false;
          let hasStrike = false;
          let foundCustomColor: string | undefined = undefined;
          let rowTagValue = '';

          row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
            // Check strikethrough in font
            if (cell.font?.strike) {
              hasStrike = true;
            }

            // Check font color for red
            if (cell.font?.color) {
              const fontArgb = (cell.font.color as any).argb || '';
              if (isRedColor(fontArgb)) {
                hasRed = true;
              }
            }

            // Check cell fill colors
            if (cell.fill && cell.fill.type === 'pattern') {
              const fg = cell.fill.fgColor;
              const bg = (cell.fill as any).bgColor;
              const fgArgb = (fg as any)?.argb || '';
              const bgArgb = (bg as any)?.argb || '';

              if (isYellowColor(fgArgb) || isYellowColor(bgArgb)) {
                hasYellow = true;
              }
              if (isRedColor(fgArgb) || isRedColor(bgArgb)) {
                hasRed = true;
              }

              const extracted = extractHexColor(fgArgb) || extractHexColor(bgArgb);
              if (extracted && !['#FFFFFF', '#000000', '#00000000'].includes(extracted)) {
                foundCustomColor = extracted;
              }
            }

            // Check cell value for TAG matching
            const cellText = String(cell.value || '').trim();
            if (cellText && (/^V-?\d+/i.test(cellText) || /^[A-Z0-9]+-[A-Z0-9]+/i.test(cellText))) {
              rowTagValue = cellText.toUpperCase();
            }
          });

          // Also check row-level fill if applied
          if ((row as any).fill && (row as any).fill.type === 'pattern') {
            const rFg = (row as any).fill.fgColor?.argb || '';
            if (isYellowColor(rFg)) hasYellow = true;
            if (isRedColor(rFg)) hasRed = true;
          }

          if (hasYellow || hasRed || hasStrike || foundCustomColor) {
            const styleObj = { isYellow: hasYellow, isRed: hasRed, hasStrike, customColor: foundCustomColor };
            rowStylesMap.set(rowNumber, styleObj);
            if (rowTagValue) {
              rowStylesByTagMap.set(rowTagValue, styleObj);
            }
          }
        });
      }
    } catch (e) {
      console.warn('Could not extract styles via ExcelJS:', e);
    }
  }

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

      // If TAG is missing (e.g. from SAP Name Catalog file), auto-generate standard TAG
      if (!masterRow.TAG || String(masterRow.TAG).trim() === '') {
        masterRow.TAG = `V-${String(r + 1).padStart(3, '0')}`;
      }

      // If DESTINATION (YARD) is missing, default to Yard A
      if (!masterRow['DESTINATION (YARD)'] || String(masterRow['DESTINATION (YARD)']).trim() === '') {
        masterRow['DESTINATION (YARD)'] = 'Yard A';
      }

      // Detect Revision Status based on Excel styling & content principles
      const excelRowNum = headerRowIndex + 1 + (r + 1);
      const tagKey = String(masterRow.TAG || '').trim().toUpperCase();
      const styleInfo = rowStylesMap.get(excelRowNum) || (tagKey ? rowStylesByTagMap.get(tagKey) : undefined);

      const revDesc = String(masterRow['REV. DESCRIPTION'] || '').toUpperCase();
      const remarks = String(masterRow['REMARKS'] || '').toUpperCase();
      const isExplicitlyDeleted = revDesc.includes('DELETE') || remarks.includes('DELETE') || remarks.includes('DELETED');

      if (styleInfo?.hasStrike || styleInfo?.isRed || isExplicitlyDeleted) {
        // Rule 2: Deleted valve/armature -> Red colour with strikethrough (delete line)
        masterRow._revisionStatus = 'deleted';
        masterRow._detectedColor = '#EF4444';
        masterRow._hasStrikethrough = true;
      } else if (styleInfo?.isYellow || revDesc.includes('NEW') || revDesc.includes('UPDATE')) {
        // Rule 1: New or updated valve/armature -> Yellow coloured row
        masterRow._revisionStatus = 'new_updated';
        masterRow._detectedColor = '#FFFF00';
        masterRow._hasStrikethrough = false;
      } else if (styleInfo?.customColor) {
        masterRow._revisionStatus = 'normal';
        masterRow._detectedColor = styleInfo.customColor;
        masterRow._hasStrikethrough = false;
      } else {
        // Rule 3: Next revision after deleted valve/armature -> Keep SFI and TAG (rest of cells are empty)
        const hasTagOrSfi = !!(masterRow.TAG || masterRow.SFI);
        const specKeys = ['DESCRIPTION', 'SIZE', 'CONNECTION', 'PRESSURE RATING', 'HOUSING /BODY', 'TYPE', 'PIPE CLASS'];
        const isRestEmpty = specKeys.every((k) => !masterRow[k] || String(masterRow[k]).trim() === '' || String(masterRow[k]).trim() === '-');

        if (hasTagOrSfi && isRestEmpty && (revDesc.includes('REV') || revDesc.includes('DELETE') || !masterRow['DESCRIPTION'])) {
          masterRow._revisionStatus = 'next_rev_after_deleted';
        } else {
          masterRow._revisionStatus = 'normal';
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
