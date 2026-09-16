import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { CatalogItem, MasterRowData } from '../types';

/**
 * All 31 Catalog columns matching the ship armature specification
 */
export const CATALOG_COLUMNS: string[] = [
  'SAP Name',
  'Description',
  'NRF Nr.',
  'Normale Nr.',
  'Execution',
  'Manufacturer',
  'Model Number',
  'Used In HVAC Generator',
  'Pipe Size',
  'Connection',
  'Pressure Nominal',
  'Pipe Class',
  'Body Material',
  'Disc Material',
  'Seat Material',
  'Stem Material',
  'Stem Height [mm]',
  'Stem Size',
  'Stem Shape',
  'ISO 5211',
  'Means Of Operation',
  'Operating Temp Min. [degC]',
  'Operating Temp Max. [degC]',
  'Setpoint Min',
  'Setpoint Max',
  'Dry Weight [kg]',
  'Typical Usage',
  'Testing Certificate',
  'MRP Type',
  'Sign Type',
  'Comment',
];

/**
 * 6 sample rows extracted directly from the user's ship catalog image
 */
export const DEFAULT_SAMPLE_CATALOG_ITEMS: CatalogItem[] = [
  {
    id: 'cat-sample-1',
    'SAP Name': 'A4010320608F',
    Name: 'A4010320608F',
    Description: 'Globe valve',
    'NRF Nr.': '',
    'Normale Nr.': 'NE401130',
    Execution: '',
    Manufacturer: 'M.Medana&ViscaS.r.l.',
    'Model Number': 'TAB 126.1',
    'Used In HVAC Generator': 'False',
    'Pipe Size': 'DN15',
    Connection: 'Flanged',
    'Pressure Nominal': 'PN16',
    'Pipe Class': '',
    'Body Material': 'Cast iron',
    'Disc Material': 'AISI 316',
    'Seat Material': 'AISI 316',
    'Stem Material': 'AISI 316',
    'Stem Height [mm]': '',
    'Stem Size': '',
    'Stem Shape': '',
    'ISO 5211': '',
    'Means Of Operation': 'Handwheel',
    'Operating Temp Min. [degC]': -10.0,
    'Operating Temp Max. [degC]': 120.0,
    'Setpoint Min': '',
    'Setpoint Max': '',
    'Dry Weight [kg]': 4.0,
    'Typical Usage': '',
    'Testing Certificate': 'FT',
    'MRP Type': '03',
    'Sign Type': '',
    Comment: '',
  },
  {
    id: 'cat-sample-2',
    'SAP Name': 'A401032060PF',
    Name: 'A401032060PF',
    Description: 'Globe valve',
    'NRF Nr.': '',
    'Normale Nr.': 'NE401130',
    Execution: '',
    Manufacturer: 'M.Medana&ViscaS.r.l.',
    'Model Number': 'TAB 126.1',
    'Used In HVAC Generator': 'False',
    'Pipe Size': 'DN15',
    Connection: 'Flanged',
    'Pressure Nominal': 'PN16',
    'Pipe Class': '',
    'Body Material': 'Cast iron',
    'Disc Material': 'AISI 316',
    'Seat Material': 'AISI 316',
    'Stem Material': 'AISI 316',
    'Stem Height [mm]': '',
    'Stem Size': '',
    'Stem Shape': '',
    'ISO 5211': '',
    'Means Of Operation': 'Handwheel',
    'Operating Temp Min. [degC]': -10.0,
    'Operating Temp Max. [degC]': 120.0,
    'Setpoint Min': '',
    'Setpoint Max': '',
    'Dry Weight [kg]': 4.0,
    'Typical Usage': '',
    'Testing Certificate': 'Project',
    'MRP Type': '01',
    'Sign Type': '',
    Comment: '',
  },
  {
    id: 'cat-sample-3',
    'SAP Name': 'A4010320708T',
    Name: 'A4010320708T',
    Description: 'Globe valve',
    'NRF Nr.': '',
    'Normale Nr.': 'NE401130',
    Execution: '',
    Manufacturer: 'M.Medana&ViscaS.r.l.',
    'Model Number': 'TAB 126.1',
    'Used In HVAC Generator': 'False',
    'Pipe Size': 'DN20',
    Connection: 'Flanged',
    'Pressure Nominal': 'PN16',
    'Pipe Class': '',
    'Body Material': 'Cast iron',
    'Disc Material': 'AISI 316',
    'Seat Material': 'AISI 316',
    'Stem Material': 'AISI 316',
    'Stem Height [mm]': '',
    'Stem Size': '',
    'Stem Shape': '',
    'ISO 5211': '',
    'Means Of Operation': 'Handwheel',
    'Operating Temp Min. [degC]': -10.0,
    'Operating Temp Max. [degC]': 120.0,
    'Setpoint Min': '',
    'Setpoint Max': '',
    'Dry Weight [kg]': 6.0,
    'Typical Usage': '',
    'Testing Certificate': 'FT',
    'MRP Type': '03',
    'Sign Type': '',
    Comment: '',
  },
  {
    id: 'cat-sample-4',
    'SAP Name': 'A401032070PT',
    Name: 'A401032070PT',
    Description: 'Globe valve',
    'NRF Nr.': '',
    'Normale Nr.': 'NE401130',
    Execution: '',
    Manufacturer: 'M.Medana&ViscaS.r.l.',
    'Model Number': 'TAB 126.1',
    'Used In HVAC Generator': 'False',
    'Pipe Size': 'DN20',
    Connection: 'Flanged',
    'Pressure Nominal': 'PN16',
    'Pipe Class': '',
    'Body Material': 'Cast iron',
    'Disc Material': 'AISI 316',
    'Seat Material': 'AISI 316',
    'Stem Material': 'AISI 316',
    'Stem Height [mm]': '',
    'Stem Size': '',
    'Stem Shape': '',
    'ISO 5211': '',
    'Means Of Operation': 'Handwheel',
    'Operating Temp Min. [degC]': -10.0,
    'Operating Temp Max. [degC]': 120.0,
    'Setpoint Min': '',
    'Setpoint Max': '',
    'Dry Weight [kg]': 6.0,
    'Typical Usage': '',
    'Testing Certificate': 'Project',
    'MRP Type': '01',
    'Sign Type': '',
    Comment: '',
  },
  {
    id: 'cat-sample-5',
    'SAP Name': 'A5051220602A',
    Name: 'A5051220602A',
    Description: 'SPECTACLE FLANGE',
    'NRF Nr.': '',
    'Normale Nr.': 'NI505122',
    Execution: '',
    Manufacturer: 'Standard Flange Corp',
    'Model Number': '',
    'Used In HVAC Generator': 'False',
    'Pipe Size': 'DN15',
    Connection: 'Flanged',
    'Pressure Nominal': 'PN16',
    'Pipe Class': 'LR',
    'Body Material': 'Steel',
    'Disc Material': '',
    'Seat Material': '',
    'Stem Material': '',
    'Stem Height [mm]': '',
    'Stem Size': '',
    'Stem Shape': '',
    'ISO 5211': '',
    'Means Of Operation': '',
    'Operating Temp Min. [degC]': '',
    'Operating Temp Max. [degC]': '',
    'Setpoint Min': '',
    'Setpoint Max': '',
    'Dry Weight [kg]': 1.8,
    'Typical Usage': 'Spectacle Flange Isolation',
    'Testing Certificate': 'YES',
    'MRP Type': '',
    'Sign Type': '',
    Comment: '',
  },
  {
    id: 'cat-sample-6',
    'SAP Name': 'V5518319100B',
    Name: 'V5518319100B',
    Description: 'HOSE COUPLING',
    'NRF Nr.': '',
    'Normale Nr.': 'VNE551831',
    Execution: 'D',
    Manufacturer: 'Camlock Systems',
    'Model Number': 'Camlock Type D',
    'Used In HVAC Generator': 'False',
    'Pipe Size': 'DN50',
    Connection: 'Threaded BSP',
    'Pressure Nominal': 'PN10',
    'Pipe Class': 'None',
    'Body Material': 'AISI 316',
    'Disc Material': '',
    'Seat Material': 'NBR',
    'Stem Material': '',
    'Stem Height [mm]': '',
    'Stem Size': '',
    'Stem Shape': '',
    'ISO 5211': '',
    'Means Of Operation': '',
    'Operating Temp Min. [degC]': -20.0,
    'Operating Temp Max. [degC]': 100.0,
    'Setpoint Min': '',
    'Setpoint Max': '',
    'Dry Weight [kg]': 0.85,
    'Typical Usage': 'Bunkering & Cargo Hose Connection',
    'Testing Certificate': 'NO',
    'MRP Type': '',
    'Sign Type': '',
    Comment: '',
  },
];

export interface ArmatureSpecLookupResult {
  found: boolean;
  matchedItem?: CatalogItem;
  armatureFields: Record<string, string>;
}

/**
 * Looks up catalog by SAP Code (Name) or Model / Normale Nr and returns
 * the exact columns for the Armature Master Table.
 */
export function lookupCatalogBySapCode(
  sapCode: string,
  catalogItems: CatalogItem[]
): ArmatureSpecLookupResult {
  if (!sapCode || !catalogItems || catalogItems.length === 0) {
    return { found: false, armatureFields: {} };
  }

  const cleanCode = sapCode.trim();
  const upperCode = cleanCode.toUpperCase();

  // 1. Search by exact SAP Name or Name
  let matched = catalogItems.find(
    (item) =>
      (item['SAP Name'] && item['SAP Name'].trim().toUpperCase() === upperCode) ||
      (item.Name && item.Name.trim().toUpperCase() === upperCode)
  );

  // 2. Search by Model Number
  if (!matched) {
    matched = catalogItems.find(
      (item) => item['Model Number'] && String(item['Model Number']).trim().toUpperCase() === upperCode
    );
  }

  // 3. Search by Normale Nr.
  if (!matched) {
    matched = catalogItems.find(
      (item) => item['Normale Nr.'] && String(item['Normale Nr.']).trim().toUpperCase() === upperCode
    );
  }

  // 4. Prefix or substring match if query length >= 4
  if (!matched && upperCode.length >= 4) {
    matched = catalogItems.find(
      (item) =>
        (item['SAP Name'] && item['SAP Name'].trim().toUpperCase().includes(upperCode)) ||
        (item.Name && item.Name.trim().toUpperCase().includes(upperCode))
    );
  }

  if (!matched) {
    return {
      found: false,
      armatureFields: {},
    };
  }

  // Map to armature columns matching the user's pipe specification
  const fields: Record<string, string> = {
    'STD DRW NORMALE N°': matched['Normale Nr.'] ? String(matched['Normale Nr.']) : '',
    'EXECUTION': matched.Execution ? String(matched.Execution) : '',
    'NRF N°': matched['NRF Nr.'] ? String(matched['NRF Nr.']) : '',
    'DESCRIPTION': matched.Description ? String(matched.Description) : '',
    'SIZE': matched['Pipe Size'] ? String(matched['Pipe Size']) : '',
    'CONNECTION': matched.Connection ? String(matched.Connection) : '',
    'PRESSURE RATING': matched['Pressure Nominal'] ? String(matched['Pressure Nominal']) : '',
    'HOUSING /BODY': matched['Body Material'] ? String(matched['Body Material']) : '',
    'TYPE': matched['Model Number']
      ? String(matched['Model Number'])
      : (matched['Sign Type'] ? String(matched['Sign Type']) : (matched.Execution ? String(matched.Execution) : '')),
    'PIPE CLASS': matched['Pipe Class'] ? String(matched['Pipe Class']) : '',
    'CLASS CERTIFICATE': matched['Testing Certificate'] ? String(matched['Testing Certificate']) : '',
    'REMARKS': matched.Comment ? String(matched.Comment) : '',
  };

  if (matched.Manufacturer) {
    fields['SUPPLIER'] = String(matched.Manufacturer);
  }

  return {
    found: true,
    matchedItem: matched,
    armatureFields: fields,
  };
}

/**
 * Downloads a blob buffer in the browser
 */
function downloadBlob(buffer: ArrayBuffer, fileName: string) {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/**
 * Generates an Excel catalog template (.xlsx) with all 31 columns
 * Can generate either a blank template or one populated with sample data
 */
export async function downloadCatalogTemplate(
  blankOnly: boolean = true,
  shipName?: string
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Armature Import Tool - Catalog Manager';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('Valve_Catalog', {
    views: [{ showGridLines: true, state: 'frozen', ySplit: 1 }],
  });

  // Setup column definitions
  worksheet.columns = CATALOG_COLUMNS.map((header) => {
    let width = Math.max(header.length + 5, 14);
    if (['Description', 'Manufacturer'].includes(header)) width = 24;
    if (['Name', 'Model Number', 'Normale Nr.'].includes(header)) width = 18;
    return {
      header,
      key: header,
      width,
    };
  });

  // Style header row (Row 1)
  const headerRow = worksheet.getRow(1);
  headerRow.height = 28;

  const headerBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF0F172A' } },
    left: { style: 'thin', color: { argb: 'FF334155' } },
    bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
    right: { style: 'thin', color: { argb: 'FF334155' } },
  };

  CATALOG_COLUMNS.forEach((_, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } }; // Dark Navy Blue
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: false };
    cell.border = headerBorder;
  });

  const dataBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  };

  if (blankOnly) {
    // 25 pre-formatted blank rows
    for (let r = 0; r < 25; r++) {
      const row = worksheet.addRow(new Array(CATALOG_COLUMNS.length).fill(''));
      row.height = 22;
      const isAlt = r % 2 === 1;
      for (let c = 1; c <= CATALOG_COLUMNS.length; c++) {
        const cell = row.getCell(c);
        cell.border = dataBorder;
        cell.font = { name: 'Segoe UI', size: 9.5 };
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        if (isAlt) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        }
      }
    }
  } else {
    // Populate with 4 sample rows
    DEFAULT_SAMPLE_CATALOG_ITEMS.forEach((item, rIdx) => {
      const rowData = CATALOG_COLUMNS.map((col) => item[col] ?? '');
      const row = worksheet.addRow(rowData);
      row.height = 22;
      const isAlt = rIdx % 2 === 1;
      for (let c = 1; c <= CATALOG_COLUMNS.length; c++) {
        const cell = row.getCell(c);
        cell.border = dataBorder;
        cell.font = { name: 'Segoe UI', size: 9.5 };
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        if (isAlt) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
        }
      }
    });
  }

  // Auto-filter
  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: CATALOG_COLUMNS.length },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  const safeShip = shipName ? `_${shipName.replace(/[^a-zA-Z0-9_-]/g, '_')}` : '';
  const fileName = `Catalog_Armature${safeShip}_${blankOnly ? 'Blank' : 'Sample'}.xlsx`;
  downloadBlob(buffer, fileName);
}

/**
 * Normalizes header string for fuzzy matching
 */
function cleanColName(str: any): string {
  if (str == null) return '';
  return String(str)
    .trim()
    .toUpperCase()
    .replace(/[\r\n\t]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/[._\-–—\[\]\(\)]/g, '');
}

/**
 * Parses an uploaded Catalog Excel or CSV file
 */
export async function parseCatalogFile(file: File): Promise<{
  items: CatalogItem[];
  sheetName: string;
  totalRows: number;
}> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('File không chứa bất kỳ sheet nào.');
  }

  // Find sheet containing catalog columns
  let bestSheetName = workbook.SheetNames[0];
  let bestHeaderRowIndex = -1;
  let bestMatchedCount = 0;

  const normalizedCatalogCols = CATALOG_COLUMNS.map((col) => ({
    original: col,
    clean: cleanColName(col),
  }));

  for (const sheetName of workbook.SheetNames) {
    const ws = workbook.Sheets[sheetName];
    if (!ws) continue;
    const data: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    const scanLimit = Math.min(data.length, 30);

    for (let r = 0; r < scanLimit; r++) {
      const row = data[r];
      if (!Array.isArray(row)) continue;

      let matchCount = 0;
      row.forEach((cellVal) => {
        const cleanVal = cleanColName(cellVal);
        if (cleanVal && normalizedCatalogCols.some((nc) => nc.clean === cleanVal)) {
          matchCount++;
        }
      });

      if (matchCount > bestMatchedCount) {
        bestMatchedCount = matchCount;
        bestHeaderRowIndex = r;
        bestSheetName = sheetName;
      }
    }
  }

  if (bestHeaderRowIndex === -1 || bestMatchedCount < 2) {
    // Fallback: use first row of first sheet
    bestSheetName = workbook.SheetNames[0];
    bestHeaderRowIndex = 0;
  }

  const ws = workbook.Sheets[bestSheetName];
  const sheetData: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const rawHeaders = (sheetData[bestHeaderRowIndex] || []).map((h: any) => String(h || '').trim());

  // Map header index to catalog column
  const colMapping: Record<number, string> = {};
  rawHeaders.forEach((rawH, idx) => {
    const cleanH = cleanColName(rawH);
    if (['SAPNAME', 'SAPCODE', 'SAP', 'NAME', 'VALVENAME', 'VALVECODE', 'PARTNUMBER'].includes(cleanH)) {
      colMapping[idx] = 'SAP Name';
      return;
    }
    const matched = normalizedCatalogCols.find((nc) => nc.clean === cleanH);
    if (matched) {
      colMapping[idx] = matched.original;
    } else if (rawH) {
      // Fallback direct name
      colMapping[idx] = rawH;
    }
  });

  const items: CatalogItem[] = [];
  for (let r = bestHeaderRowIndex + 1; r < sheetData.length; r++) {
    const row = sheetData[r];
    if (!Array.isArray(row)) continue;

    // Check if row is not completely empty
    const hasData = row.some((val) => val !== null && val !== undefined && String(val).trim() !== '');
    if (!hasData) continue;

    const item: CatalogItem = {
      id: `cat-${Date.now()}-${r}`,
      'SAP Name': '',
      Name: '',
    };

    row.forEach((cellVal, colIdx) => {
      const targetCol = colMapping[colIdx];
      if (targetCol) {
        let val = cellVal;
        if (typeof val === 'number') {
          // preserve number
        } else if (val !== null && val !== undefined) {
          val = String(val).trim();
        } else {
          val = '';
        }
        item[targetCol] = val;
      }
    });

    const primarySapName = item['SAP Name'] || item.Name || '';
    item['SAP Name'] = primarySapName;
    item.Name = primarySapName;

    if (item['SAP Name'] || item.Name || item['Model Number'] || item.Description || item['Normale Nr.']) {
      items.push(item);
    }
  }

  return {
    items,
    sheetName: bestSheetName,
    totalRows: items.length,
  };
}

/**
 * Exports catalog items to an Excel (.xlsx) file
 */
export async function exportCatalogToExcel(
  items: CatalogItem[],
  shipName: string = 'Ship_Catalog',
  versionName: string = 'Rev1'
): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Armature Import Tool';
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet('Valve_Catalog', {
    views: [{ showGridLines: true, state: 'frozen', ySplit: 1 }],
  });

  worksheet.columns = CATALOG_COLUMNS.map((header) => ({
    header,
    key: header,
    width: Math.max(header.length + 5, 14),
  }));

  const headerRow = worksheet.getRow(1);
  headerRow.height = 28;
  CATALOG_COLUMNS.forEach((_, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  const dataBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  };

  items.forEach((item, rIdx) => {
    const rowValues = CATALOG_COLUMNS.map((col) => item[col] ?? '');
    const row = worksheet.addRow(rowValues);
    row.height = 22;
    const isAlt = rIdx % 2 === 1;
    for (let c = 1; c <= CATALOG_COLUMNS.length; c++) {
      const cell = row.getCell(c);
      cell.border = dataBorder;
      cell.font = { name: 'Segoe UI', size: 9.5 };
      if (isAlt) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
      }
    }
  });

  worksheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: CATALOG_COLUMNS.length },
  };

  const buffer = await workbook.xlsx.writeBuffer();
  const safeTitle = `${shipName}_${versionName}`.replace(/[^a-zA-Z0-9_-]/g, '_');
  downloadBlob(buffer, `Catalog_${safeTitle}.xlsx`);
}

/**
 * Enriches Master rows with specifications looked up from the active Catalog
 */
export function autoFillMasterFromCatalog(
  masterRows: MasterRowData[],
  catalogItems: CatalogItem[]
): {
  updatedRows: MasterRowData[];
  matchedCount: number;
  fieldsUpdatedCount: number;
} {
  if (!catalogItems || catalogItems.length === 0 || !masterRows || masterRows.length === 0) {
    return { updatedRows: masterRows, matchedCount: 0, fieldsUpdatedCount: 0 };
  }

  // Build lookup maps for fast matching
  const catalogByName = new Map<string, CatalogItem>();
  const catalogByNormale = new Map<string, CatalogItem>();
  const catalogByModel = new Map<string, CatalogItem>();

  catalogItems.forEach((item) => {
    const sapName = (item['SAP Name'] || item.Name || '').trim().toUpperCase();
    if (sapName) {
      catalogByName.set(sapName, item);
    }
    if (item.Name) {
      catalogByName.set(item.Name.trim().toUpperCase(), item);
    }
    if (item['Normale Nr.']) {
      catalogByNormale.set(String(item['Normale Nr.']).trim().toUpperCase(), item);
    }
    if (item['Model Number']) {
      catalogByModel.set(String(item['Model Number']).trim().toUpperCase(), item);
    }
  });

  let matchedCount = 0;
  let fieldsUpdatedCount = 0;

  const updatedRows = masterRows.map((row) => {
    const tag = String(row['TAG'] || '').trim().toUpperCase();
    const desc = String(row['DESCRIPTION'] || '').trim().toUpperCase();
    const sfi = String(row['SFI'] || '').trim().toUpperCase();
    const stdDwg = String(row['STD DRW NORMALE N°'] || row['STD DRAWING'] || '').trim().toUpperCase();

    // Match priority: TAG === SAP Name -> TAG === Model -> SFI/STD DRAWING === Normale Nr. -> TAG in SAP Name
    let matchedItem: CatalogItem | undefined =
      catalogByName.get(tag) ||
      catalogByModel.get(tag) ||
      catalogByNormale.get(sfi) ||
      catalogByNormale.get(stdDwg);

    if (!matchedItem && tag) {
      // Fuzzy partial match
      for (const [cName, cItem] of catalogByName.entries()) {
        if (cName.includes(tag) || tag.includes(cName)) {
          matchedItem = cItem;
          break;
        }
      }
    }

    if (!matchedItem && desc) {
      for (const [cName, cItem] of catalogByName.entries()) {
        if (desc.includes(cName)) {
          matchedItem = cItem;
          break;
        }
      }
    }

    if (!matchedItem) {
      return row;
    }

    matchedCount++;
    const newRow = { ...row };

    const updateFieldIfEmpty = (masterKey: string, val: any) => {
      if (val !== undefined && val !== null && String(val).trim() !== '') {
        if (!newRow[masterKey] || String(newRow[masterKey]).trim() === '') {
          newRow[masterKey] = val;
          fieldsUpdatedCount++;
        }
      }
    };

    // Enrich specifications according to pipe specification
    updateFieldIfEmpty('STD DRW NORMALE N°', matchedItem['Normale Nr.']);
    updateFieldIfEmpty('EXECUTION', matchedItem.Execution);
    updateFieldIfEmpty('NRF N°', matchedItem['NRF Nr.']);
    updateFieldIfEmpty('DESCRIPTION', matchedItem.Description);
    updateFieldIfEmpty('SIZE', matchedItem['Pipe Size']);
    updateFieldIfEmpty('CONNECTION', matchedItem.Connection);
    updateFieldIfEmpty('PRESSURE RATING', matchedItem['Pressure Nominal']);
    updateFieldIfEmpty('HOUSING /BODY', matchedItem['Body Material']);
    updateFieldIfEmpty('TYPE', matchedItem['Model Number'] || matchedItem['Sign Type'] || matchedItem.Execution);
    updateFieldIfEmpty('PIPE CLASS', matchedItem['Pipe Class']);
    updateFieldIfEmpty('CLASS CERTIFICATE', matchedItem['Testing Certificate']);
    updateFieldIfEmpty('REMARKS', matchedItem.Comment);

    // Also support legacy column names if present in table
    updateFieldIfEmpty('HOUSING BODY', matchedItem['Body Material']);
    updateFieldIfEmpty('SUPPLIER', matchedItem.Manufacturer);
    updateFieldIfEmpty('STD DRAWING', matchedItem['Normale Nr.']);

    return newRow;
  });

  return {
    updatedRows,
    matchedCount,
    fieldsUpdatedCount,
  };
}
