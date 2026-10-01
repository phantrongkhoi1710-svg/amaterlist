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
      'SAP CODE',
      'STD DRW NORMALE N°',
      'PRESSURE CLASS', // Alias for PRESSURE RATING
      'BODY', // Alias for HOUSING BODY
      'PIPE CLASS',
      'CLASS CERT', // Alias for CLASS CERTIFICATE
      'SIGN TYPE',
      'SIGN TEXT',
      'REV', // Alias for REV HIS
    ];
    rows = [
      ...introRows,
      headers,
      ['V-101-BV', 'EMERSON', '512.01', 'BALL VALVE 4" FLANGED 150#', 'ACT-101-A', 'PO-98210', 'VUNG TAU', 'A4010320608F', 'NE401130', '150#', 'A105 CS', '150-CS-01', 'DNV', 'TAG-A', 'BALL VALVE 4"', '0'],
      ['V-102-BV', 'EMERSON', '512.01', 'BALL VALVE 6" FLANGED 300#', 'ACT-102-A', 'PO-98210', 'VUNG TAU', 'A401032060PF', 'NE401130', '300#', 'A105 CS', '300-CS-01', 'DNV', 'TAG-A', 'BALL VALVE 6"', 'A'],
      ['V-103-GV', 'EMERSON', '512.02', 'GLOBE VALVE 2" NPT 800#', '', 'PO-98210', 'VUNG TAU', 'A4010320708T', 'NE401130', '800#', 'F316 SS', '800-SS-02', 'DNV', 'TAG-B', 'GLOBE VALVE 2"', '0'],
      ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''], // Blank row to test IsDataRow filtering
      ['V-104-CK', 'EMERSON', '512.03', 'CHECK VALVE DUAL PLATE 8"', '', 'PO-98211', 'QUANG NGAI', 'A401032070PT', 'NE401130', '150#', 'WCB CS', '150-CS-01', 'ABS', 'TAG-C', 'CHECK VALVE 8"', 'B'],
    ];
  } else if (variant === 'vendor_b') {
    // Uses standard headers and aliases like "STD DRAWING", "SIGNTYPE", "SIGNTEXT"
    const introRows = [
      ['VALVE SCHEDULE - MODULE 12'],
      ['DATE: 2026-03-10'],
    ];
    headers = [
      'TAG',
      'ACTUATOR',
      'PURCHASE ORDER NUMBER',
      'SUPPLIER',
      'DESTINATION YARD ACTUATOR',
      'SAP CODE',
      'STD DRAWING',
      'EXECUTION',
      'NRF NUMBER',
      'SFI',
      'DESCRIPTION',
      'SIZE',
      'CONNECTION',
      'PRESSURE',
      'HOUSING',
      'TYPE',
      'PIPE CLASS',
      'CLASS CERTIFICATE',
      'REMARKS',
      'SIGN TYPE',
      'SIGN TEXT',
      'INPUT - SIGN TEXT',
      'REV. HIS.',
      'DATE',
      'REV. DESCRIPTION',
      'SIGNATURE',
    ];
    rows = [
      ...introRows,
      headers,
      ['CV-201-P', 'ACT-KITZ-01', 'PO-55440', 'KITZ CORP', 'HAI PHONG', 'A5051220602A', 'NI505122', '', '', '514.10', 'SPECTACLE FLANGE', 'DN15', 'Flanged', 'PN16', 'Steel', '', 'LR', 'YES', 'Spectacle isolation', 'TYPE-1', 'DANGER HIGH PRESSURE', 'WARN-01', 'REV-1', '2026-03-10', 'Approved for prod', 'KHOI'],
      ['BV-202-M', '', 'PO-55440', 'KITZ CORP', 'HAI PHONG', 'V5518319100B', 'VNE551831', 'D', '', '514.12', 'HOSE COUPLING', 'DN50', 'Threaded BSP', 'PN10', 'AISI 316', 'Camlock Type D', 'None', 'NO', 'Cargo connection', 'TYPE-2', 'WATER INLET', 'INFO-02', 'REV-0', '2026-03-10', 'First issue', 'KHOI'],
      ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''], // Blank row
      ['SV-203-S', '', 'PO-55442', 'KITZ CORP', 'HAI PHONG', 'A4010320608F', 'NE401130', '', '', '514.15', 'SAFETY RELIEF VALVE', 'DN25', 'Flanged', 'PN25', 'AISI 316', 'TAB 126.1', 'LR', 'YES', 'High pressure relief', 'TYPE-1', 'CRITICAL RELIEF', 'ALERT-03', 'REV-2', '2026-03-11', 'Size revised', 'KHOI'],
    ];
  } else {
    // Subsea C
    headers = [
      'TAG',
      'ACTUATOR NO',
      'PO NO',
      'SUPPLIER',
      'DESTINATION',
      'SAP CODE',
      'STD DRW NORMALE N°',
      'DESCRIPTION',
      'SIZE',
      'CONNECTION',
      'PRESSURE RATING',
      'HOUSING BODY',
      'PIPE CLASSIFICATION',
      'CLASS CERTIFICATE',
      'REV HIS',
    ];
    rows = [
      ['SUBSEA PIPING ARMATURE PACKAGE 2026'],
      headers,
      ['SB-301-BV', 'HYD-ACT-301', 'PO-77190', 'CAMERON', 'DUNG QUAT', 'A4010320708T', 'NE401130', 'SUBSEA BALL VALVE 12" 5000 PSI', 'DN300', 'Flanged', '5000 PSI', 'DUPLEX 2205', 'API-5000', 'DNV-GL', 'C'],
      ['SB-302-CK', '', 'PO-77190', 'CAMERON', 'DUNG QUAT', 'A401032070PT', 'NE401130', 'SUBSEA NON-RETURN VALVE 8"', 'DN200', 'Flanged', '5000 PSI', 'SUPER DUPLEX', 'API-5000', 'DNV-GL', 'B'],
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
 * Generates an empty / formatted Master template file for download with beautiful colors and row spacing matching the 26 columns in the image
 */
export async function downloadMasterTemplate() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Armature Import Tool';
  workbook.created = new Date();

  // 1. Total Sheet
  const wsTotal = workbook.addWorksheet('Total', {
    views: [{ showGridLines: true, state: 'frozen', ySplit: 4 }],
  });

  const colsCount = DEFAULT_MASTER_COLUMNS.length; // 26

  // Title Banner
  const titleRow = wsTotal.addRow(['ARMATURE MASTER TEMPLATE (CHUẨN 26 CỘT THEO PIPE SPECIFICATION)']);
  titleRow.height = 36;
  wsTotal.mergeCells(1, 1, 1, colsCount);
  const titleCell = wsTotal.getCell(1, 1);
  titleCell.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

  // Subtitle / Guide Row
  const subRow = wsTotal.addRow([
    'Vàng (Cột 1-5): Functional Design  |  Xám (Cột 6-21): Armature Pipe Spec & Catalog  |  Trắng (Cột 22-26): Revision & Sign-off History',
  ]);
  subRow.height = 22;
  wsTotal.mergeCells(2, 1, 2, colsCount);
  const subCell = wsTotal.getCell(2, 1);
  subCell.font = { name: 'Segoe UI', size: 9.5, italic: true, color: { argb: 'FF475569' } };
  subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
  subCell.alignment = { vertical: 'middle', horizontal: 'center' };
  subCell.border = { bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } } };

  // Row 3: Super Headers (3 Groups matching the uploaded image)
  const superHeaderRow = wsTotal.addRow(new Array(colsCount).fill(''));
  superHeaderRow.height = 26;

  // Group 1: Cols 1-5 (Yellow)
  wsTotal.mergeCells(3, 1, 3, 5);
  const g1 = wsTotal.getCell(3, 1);
  g1.value = 'TO BE COMPLETED BY FUNCTIONAL DESIGN';
  g1.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF000000' } };
  g1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Yellow #FFFF00
  g1.alignment = { vertical: 'middle', horizontal: 'center' };

  // Group 2: Cols 6-21 (Grey)
  wsTotal.mergeCells(3, 6, 3, 21);
  const g2 = wsTotal.getCell(3, 6);
  g2.value = 'ARMATURE INFO FROM PIPE SPECIFICATION';
  g2.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF000000' } };
  g2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } }; // Metallic Grey #BFBFBF
  g2.alignment = { vertical: 'middle', horizontal: 'center' };

  // Group 3: Cols 22-26 (White)
  wsTotal.mergeCells(3, 22, 3, 26);
  const g3 = wsTotal.getCell(3, 22);
  g3.value = 'REVISION & SIGNATURE';
  g3.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF000000' } };
  g3.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8FAFC' } };
  g3.alignment = { vertical: 'middle', horizontal: 'center' };

  // Row 4: Individual Column Headers (Row 4)
  const headerRow = wsTotal.addRow([...DEFAULT_MASTER_COLUMNS]);
  headerRow.height = 30;

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF94A3B8' } },
    left: { style: 'thin', color: { argb: 'FF94A3B8' } },
    bottom: { style: 'medium', color: { argb: 'FF0F172A' } },
    right: { style: 'thin', color: { argb: 'FF94A3B8' } },
  };

  DEFAULT_MASTER_COLUMNS.forEach((header, idx) => {
    const colNum = idx + 1;
    const cell = headerRow.getCell(colNum);

    let bgColor = 'FFBFBFBF'; // Grey for cols 6-21
    if (colNum <= 5) {
      bgColor = 'FFFFFF00'; // Yellow for cols 1-5
    } else if (colNum >= 22) {
      bgColor = 'FFFFFFFF'; // White for cols 22-26
    }

    cell.font = { name: 'Segoe UI', size: 9.5, bold: true, color: { argb: 'FF000000' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bgColor } };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
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
    column.width = Math.max(header.length + 5, 16);
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
  downloadBlob(buffer, 'Armature_Master_Template_26_Columns.xlsx');
}

/**
 * Creates the sample rows shown in the reference image with all 26 columns:
 * Row 1: Spectacle flange NI505122
 * Row 2: Empty/dashed row
 * Row 3: Hose coupling VNE551831 (Camlock Type D)
 */
export function generateImageArmatureSampleRows(): any[] {
  return [
    {
      _id: 'sample-armature-img-1',
      _sourceFile: 'Functional_Design_PipeSpec.xlsx',
      _revisionStatus: 'new_updated',
      _detectedColor: '#FFFF00',
      _hasStrikethrough: false,
      TAG: 'V-101-FL',
      'ACTUATOR TAG': '',
      'P.O. NUMBER': 'PO-2026-001',
      'SUPPLIER': 'Standard Flange Corp',
      'DESTINATION (YARD)': 'YARD-A',
      'SAP CODE': 'A5051220602A',
      'STD DRW NORMALE N°': 'NI505122',
      'EXECUTION': '',
      'NRF N°': '',
      'SFI': '512.01',
      'DESCRIPTION': 'SPECTACLE FLANGE',
      'SIZE': 'DN15',
      'CONNECTION': 'Flanged',
      'PRESSURE RATING': 'PN16',
      'HOUSING /BODY': 'Steel',
      'TYPE': '',
      'PIPE CLASS': 'LR',
      'CLASS CERTIFICATE': 'YES',
      'REMARKS': 'Isolation flange',
      'SIGN TYPE': 'TAG-STD',
      'SIGN TEXT': 'FLANGE DN15 PN16',
      'INPUT - SIGN TEXT': 'SP-01',
      'REV. HIS.': '0',
      'DATE': '2026-03-15',
      'REV. DESCRIPTION': 'New or updated valve (Yellow)',
      'SIGNATURE': 'PTK',
    },
    {
      _id: 'sample-armature-img-2',
      _sourceFile: 'Functional_Design_PipeSpec.xlsx',
      _revisionStatus: 'deleted',
      _detectedColor: '#EF4444',
      _hasStrikethrough: true,
      TAG: 'V-102-OLD',
      'ACTUATOR TAG': '',
      'P.O. NUMBER': 'PO-2026-001',
      'SUPPLIER': 'Standard Flange Corp',
      'DESTINATION (YARD)': 'YARD-A',
      'SAP CODE': 'A5051220602A',
      'STD DRW NORMALE N°': 'NI505122',
      'EXECUTION': '',
      'NRF N°': '',
      'SFI': '512.01',
      'DESCRIPTION': 'GATE VALVE 2" (DELETED LINE)',
      'SIZE': 'DN50',
      'CONNECTION': 'Flanged',
      'PRESSURE RATING': 'PN16',
      'HOUSING /BODY': 'Steel',
      'TYPE': '',
      'PIPE CLASS': 'LR',
      'CLASS CERTIFICATE': 'NO',
      'REMARKS': 'Deleted valve/armature: Red colour with strikethrough',
      'SIGN TYPE': 'TAG-DEL',
      'SIGN TEXT': 'DELETED VALVE',
      'INPUT - SIGN TEXT': '-',
      'REV. HIS.': '1',
      'DATE': '2026-03-18',
      'REV. DESCRIPTION': 'Deleted valve/armature',
      'SIGNATURE': 'PTK',
    },
    {
      _id: 'sample-armature-img-3',
      _sourceFile: 'Functional_Design_PipeSpec.xlsx',
      _revisionStatus: 'normal',
      TAG: 'V-202-CK',
      'ACTUATOR TAG': '',
      'P.O. NUMBER': 'PO-2026-002',
      'SUPPLIER': 'Camlock Systems',
      'DESTINATION (YARD)': 'YARD-B',
      'SAP CODE': 'V5518319100B',
      'STD DRW NORMALE N°': 'VNE551831',
      'EXECUTION': 'D',
      'NRF N°': '',
      'SFI': '514.02',
      'DESCRIPTION': 'HOSE COUPLING',
      'SIZE': 'DN50',
      'CONNECTION': 'Threaded BSP',
      'PRESSURE RATING': 'PN10',
      'HOUSING /BODY': 'AISI 316',
      'TYPE': 'Camlock Type D',
      'PIPE CLASS': 'None',
      'CLASS CERTIFICATE': 'NO',
      'REMARKS': 'Cargo hose coupling',
      'SIGN TYPE': 'LABEL-B',
      'SIGN TEXT': 'CARGO COUPLING DN50',
      'INPUT - SIGN TEXT': 'CK-02',
      'REV. HIS.': 'A',
      'DATE': '2026-03-20',
      'REV. DESCRIPTION': 'Material verified',
      'SIGNATURE': 'PTK',
    },
    {
      _id: 'sample-armature-img-4',
      _sourceFile: 'Functional_Design_PipeSpec.xlsx',
      _revisionStatus: 'next_rev_after_deleted',
      TAG: 'V-303-DEL',
      'ACTUATOR TAG': '',
      'P.O. NUMBER': '',
      'SUPPLIER': '',
      'DESTINATION (YARD)': '',
      'SAP CODE': '',
      'STD DRW NORMALE N°': '',
      'EXECUTION': '',
      'NRF N°': '',
      'SFI': '514.99',
      'DESCRIPTION': '',
      'SIZE': '',
      'CONNECTION': '',
      'PRESSURE RATING': '',
      'HOUSING /BODY': '',
      'TYPE': '',
      'PIPE CLASS': '',
      'CLASS CERTIFICATE': '',
      'REMARKS': '',
      'SIGN TYPE': '',
      'SIGN TEXT': '',
      'INPUT - SIGN TEXT': '',
      'REV. HIS.': '2',
      'DATE': '2026-03-25',
      'REV. DESCRIPTION': 'Next rev after deleted (Keep SFI & TAG)',
      'SIGNATURE': 'PTK',
    },
  ];
}
