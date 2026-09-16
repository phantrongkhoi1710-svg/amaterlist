import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { DEFAULT_MASTER_COLUMNS } from './headerNormalizer';

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
 * Creates a sample Excel file (as a Blob / File object) that simulates a typical supplier or sub-contractor Armature sheet.
 */
export function generateSampleExcelFile(
  fileName: string,
  variant: 'vendor_a' | 'vendor_b' | 'subsea_c'
): File {
  const wb = XLSX.utils.book_new();

  let headerRow = 1;
  let headers: string[] = [];
  let rows: any[][] = [];

  if (variant === 'vendor_a') {
    // Uses aliases: "PO NO", "ACTUATOR TAG NO", "PRESSURE CLASS", "CLASS CERT"
    // Header placed at row 4 (0-indexed 3) to test automatic row detection
    const introRows = [
      ['SUPPLIER QUOTATION & SPECIFICATION SHEET'],
      ['PROJECT: FPSO VIETSOVPETRO 04'],
      [''], // empty row
    ];
    headers = [
      'TAG',
      'SUPPLIER',
      'SFI',
      'DESCRIPTION',
      'ACTUATOR TAG NO', // Alias for ACTUATOR TAG
      'PO NO', // Alias for PO NUMBER
      'DESTINATION', // Alias for DESTINATION YARD
      'PRESSURE CLASS', // Alias for PRESSURE RATING
      'BODY', // Alias for HOUSING BODY
      'PIPE CLASS',
      'CLASS CERT', // Alias for CLASS CERTIFICATE
      'REV', // Alias for REV HIS
    ];
    rows = [
      ...introRows,
      headers,
      ['V-101-BV', 'EMERSON', '512.01', 'BALL VALVE 4" FLANGED 150#', 'ACT-101-A', 'PO-98210', 'VUNG TAU', '150#', 'A105 CS', '150-CS-01', 'DNV', '0'],
      ['V-102-BV', 'EMERSON', '512.01', 'BALL VALVE 6" FLANGED 300#', 'ACT-102-A', 'PO-98210', 'VUNG TAU', '300#', 'A105 CS', '300-CS-01', 'DNV', 'A'],
      ['V-103-GV', 'EMERSON', '512.02', 'GLOBE VALVE 2" NPT 800#', '', 'PO-98210', 'VUNG TAU', '800#', 'F316 SS', '800-SS-02', 'DNV', '0'],
      ['', '', '', '', '', '', '', '', '', '', '', ''], // Blank row to test IsDataRow filtering
      ['V-104-CK', 'EMERSON', '512.03', 'CHECK VALVE DUAL PLATE 8"', '', 'PO-98211', 'QUANG NGAI', '150#', 'WCB CS', '150-CS-01', 'ABS', 'B'],
    ];
  } else if (variant === 'vendor_b') {
    // Uses standard headers and aliases like "STD DRAWING", "SIGNTYPE", "SIGNTEXT"
    const introRows = [
      ['VALVE SCHEDULE - MODULE 12'],
      ['DATE: 2026-03-10'],
    ];
    headers = [
      'TAG',
      'SUPPLIER',
      'DESCRIPTION',
      'SFI',
      'ACTUATOR', // Alias
      'PURCHASE ORDER NUMBER', // Alias
      'DESTINATION YARD ACTUATOR', // Alias
      'STD DRAWING', // Alias
      'PRESSURE', // Alias
      'HOUSING', // Alias
      'SIGN TYPE',
      'SIGN TEXT',
      'INPUT SIGN TEXT',
      'REVISION HISTORY', // Alias
    ];
    rows = [
      ...introRows,
      headers,
      ['CV-201-P', 'KITZ CORP', 'CONTROL VALVE 3" PNEUMATIC', '514.10', 'ACT-KITZ-01', 'PO-55440', 'HAI PHONG', 'STD-DWG-514', '600#', 'CF8M', 'TYPE-1', 'DANGER HIGH PRESSURE', 'WARN-01', 'REV-1'],
      ['BV-202-M', 'KITZ CORP', 'BUTTERFLY VALVE 10" WAFER', '514.12', '', 'PO-55440', 'HAI PHONG', 'STD-DWG-514', '150#', 'DUCTILE IRON', 'TYPE-2', 'WATER INLET', 'INFO-02', 'REV-0'],
      ['', '', '', '', '', '', '', '', '', '', '', '', '', ''], // Blank row
      ['SV-203-S', 'KITZ CORP', 'SAFETY RELIEF VALVE 2"x3"', '514.15', '', 'PO-55442', 'HAI PHONG', 'STD-DWG-518', '900#', 'INCONEL 625', 'TYPE-1', 'CRITICAL RELIEF', 'ALERT-03', 'REV-2'],
    ];
  } else {
    // Subsea C
    headers = [
      'TAG',
      'SUPPLIER',
      'SFI',
      'DESCRIPTION',
      'ACTUATOR NO',
      'PO NO',
      'DESTINATION',
      'PRESSURE RATING',
      'HOUSING BODY',
      'PIPE CLASSIFICATION',
      'CLASS CERTIFICATE',
      'REV HIS',
    ];
    rows = [
      ['SUBSEA PIPING ARMATURE PACKAGE 2026'],
      headers,
      ['SB-301-BV', 'CAMERON', '611.01', 'SUBSEA BALL VALVE 12" 5000 PSI', 'HYD-ACT-301', 'PO-77190', 'DUNG QUAT', '5000 PSI', 'DUPLEX 2205', 'API-5000', 'DNV-GL', 'C'],
      ['SB-302-CK', 'CAMERON', '611.02', 'SUBSEA NON-RETURN VALVE 8"', '', 'PO-77190', 'DUNG QUAT', '5000 PSI', 'SUPER DUPLEX', 'API-5000', 'DNV-GL', 'B'],
    ];
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'ArmatureData');

  const u8arr = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([u8arr], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  return new File([blob], `${fileName}.xlsx`, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/**
 * Generates an empty / formatted Master template file for download with beautiful colors and row spacing
 */
export async function downloadMasterTemplate() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Armature Import Tool';
  workbook.created = new Date();

  // 1. Total Sheet
  const wsTotal = workbook.addWorksheet('Total', {
    views: [{ showGridLines: true, state: 'frozen', ySplit: 3 }],
  });

  const colsCount = DEFAULT_MASTER_COLUMNS.length;

  // Title Banner
  const titleRow = wsTotal.addRow(['ARMATURE MASTER TEMPLATE (FORM CHUẨN DỰ ÁN)']);
  titleRow.height = 36;
  wsTotal.mergeCells(1, 1, 1, colsCount);
  const titleCell = wsTotal.getCell(1, 1);
  titleCell.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

  // Subtitle / Guide Row
  const subRow = wsTotal.addRow([
    'Dòng 2: Hướng dẫn - Vui lòng nhập dữ liệu từ dòng 4 trở đi | Dòng 3 là tiêu đề chuẩn để macro và web tool nhận diện',
  ]);
  subRow.height = 22;
  wsTotal.mergeCells(2, 1, 2, colsCount);
  const subCell = wsTotal.getCell(2, 1);
  subCell.font = { name: 'Segoe UI', size: 9.5, italic: true, color: { argb: 'FF475569' } };
  subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
  subCell.alignment = { vertical: 'middle', horizontal: 'center' };
  subCell.border = { bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } } };

  // Headers
  const headerRow = wsTotal.addRow([...DEFAULT_MASTER_COLUMNS]);
  headerRow.height = 30;

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'medium', color: { argb: 'FF1E3A8A' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  DEFAULT_MASTER_COLUMNS.forEach((header, idx) => {
    const cell = headerRow.getCell(idx + 1);
    const isKeyField = ['TAG', 'SUPPLIER', 'SFI', 'DESCRIPTION'].includes(header);
    const isActuator = ['ACTUATOR TAG', 'PO NUMBER', 'DESTINATION YARD'].includes(header);

    let bgColor = 'FF2563EB';
    if (isKeyField) bgColor = 'FF1E3A8A';
    else if (isActuator) bgColor = 'FF0D9488';
    else bgColor = 'FF334155';

    cell.font = { name: 'Segoe UI', size: 10.5, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    cell.alignment = { vertical: 'middle', horizontal: isKeyField ? 'left' : 'center' };
    cell.border = thinBorder;
  });

  // Sample formatted empty rows with borders and spacing
  const dataBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  };

  // Add 10 blank spaced rows with borders
  for (let r = 0; r < 10; r++) {
    const emptyRow = wsTotal.addRow(new Array(colsCount).fill(''));
    emptyRow.height = 24;
    const rowBg = r % 2 === 0 ? 'FFFFFFFF' : 'FFF8FAFC';
    for (let c = 1; c <= colsCount; c++) {
      const cell = emptyRow.getCell(c);
      cell.border = dataBorder;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
    }
  }

  // Set widths
  DEFAULT_MASTER_COLUMNS.forEach((header, idx) => {
    const column = wsTotal.getColumn(idx + 1);
    column.width = Math.max(header.length + 6, 18);
  });

  // 2. Import_Log Sheet
  const wsLog = workbook.addWorksheet('Import_Log', {
    views: [{ showGridLines: true, state: 'frozen', ySplit: 2 }],
  });

  const logTitle = wsLog.addRow(['NHẬT KÝ QUÁ TRÌNH IMPORT (IMPORT LOG TEMPLATE)']);
  logTitle.height = 32;
  wsLog.mergeCells(1, 1, 1, 7);
  const logTitleCell = wsLog.getCell(1, 1);
  logTitleCell.font = { name: 'Segoe UI', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  logTitleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF334155' } };
  logTitleCell.alignment = { vertical: 'middle', horizontal: 'center' };

  const logHeaders = ['TIME', 'SOURCE FILE', 'SOURCE SHEET', 'HEADER ROW', 'SOURCE ROWS', 'IMPORTED ROWS', 'STATUS'];
  const logHeaderRow = wsLog.addRow(logHeaders);
  logHeaderRow.height = 26;

  logHeaders.forEach((h, idx) => {
    const cell = logHeaderRow.getCell(idx + 1);
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF475569' } };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = thinBorder;
  });

  wsLog.columns = [
    { width: 22 },
    { width: 34 },
    { width: 20 },
    { width: 14 },
    { width: 15 },
    { width: 16 },
    { width: 45 },
  ];

  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(buffer, 'Armature_Master_Template.xlsx');
}

/**
 * Creates the sample rows shown in the reference image (without SAP CODE in master):
 * Row 1: Spectacle flange NI505122
 * Row 2: Empty/dashed row
 * Row 3: Hose coupling VNE551831 (Camlock Type D)
 */
export function generateImageArmatureSampleRows(): any[] {
  return [
    {
      _id: 'sample-armature-img-1',
      _sourceFile: 'Functional_Design_PipeSpec.xlsx',
      TAG: '',
      'ACTUATOR TAG': '',
      'P.O. NUMBER': '',
      'SUPPLIER': '',
      'DESTINATION (YARD)': '',
      'STD DRW NORMALE N°': 'NI505122',
      'EXECUTION': '',
      'NRF N°': '',
      'SFI': '',
      'DESCRIPTION': 'SPECTACLE FLANGE',
      'SIZE': 'DN15',
      'CONNECTION': 'Flanged',
      'PRESSURE RATING': 'PN16',
      'HOUSING /BODY': 'Steel',
      'TYPE': '',
      'PIPE CLASS': 'LR',
      'CLASS CERTIFICATE': 'YES',
      'REMARKS': '',
    },
    {
      _id: 'sample-armature-img-2',
      _sourceFile: 'Functional_Design_PipeSpec.xlsx',
      TAG: '',
      'ACTUATOR TAG': '',
      'P.O. NUMBER': '',
      'SUPPLIER': '',
      'DESTINATION (YARD)': '',
      'STD DRW NORMALE N°': '-',
      'EXECUTION': '-',
      'NRF N°': '-',
      'SFI': '',
      'DESCRIPTION': '-',
      'SIZE': '-',
      'CONNECTION': '-',
      'PRESSURE RATING': '-',
      'HOUSING /BODY': '-',
      'TYPE': '-',
      'PIPE CLASS': '-',
      'CLASS CERTIFICATE': 'NO',
      'REMARKS': '',
    },
    {
      _id: 'sample-armature-img-3',
      _sourceFile: 'Functional_Design_PipeSpec.xlsx',
      TAG: '',
      'ACTUATOR TAG': '',
      'P.O. NUMBER': '',
      'SUPPLIER': '',
      'DESTINATION (YARD)': '',
      'STD DRW NORMALE N°': 'VNE551831',
      'EXECUTION': 'D',
      'NRF N°': '',
      'SFI': '',
      'DESCRIPTION': 'HOSE COUPLING',
      'SIZE': 'DN50',
      'CONNECTION': 'Threaded BSP',
      'PRESSURE RATING': 'PN10',
      'HOUSING /BODY': 'AISI 316',
      'TYPE': 'Camlock Type D',
      'PIPE CLASS': 'None',
      'CLASS CERTIFICATE': 'NO',
      'REMARKS': '',
    },
  ];
}
