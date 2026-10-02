/**
 * Export "Issued History" .xlsx using ExcelJS with exact styling,
 * merged header banners, logo image, column colors, and freeze panes.
 */

import ExcelJS from 'exceljs';
import { IsoDrawingRow, getExcelWeekNum } from './isoDrawingLogic';

export interface ExcelExportOptions {
  projectName: string;
  updatedBy: string;
  sourceFolderName: string;
  logoBase64?: string; // PNG or JPEG base64
  logoExtension?: 'png' | 'jpeg';
}

export async function generateIssuedHistoryExcel(
  rows: IsoDrawingRow[],
  options: ExcelExportOptions,
): Promise<Blob> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'WebToolRush ISO Drawing Manager';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Issued History', {
    views: [{ state: 'frozen', xSplit: 0, ySplit: 11 }],
  });

  // Default font Calibri across sheet
  sheet.properties.defaultRowHeight = 20;

  // Column widths: 5,40,40,7,10,10,5,13,20,8,8,8,8,8,12,15,15,15,20
  const widths = [5, 40, 40, 7, 10, 10, 5, 13, 20, 8, 8, 8, 8, 8, 12, 15, 15, 15, 20];
  widths.forEach((w, idx) => {
    sheet.getColumn(idx + 1).width = w;
  });

  // 1. Insert Logo at B3 or "Logo not found"
  if (options.logoBase64) {
    try {
      const imageId = workbook.addImage({
        base64: options.logoBase64,
        extension: options.logoExtension || 'png',
      });
      sheet.addImage(imageId, {
        tl: { col: 1, row: 2 }, // B3 (0-indexed col 1, row 2)
        ext: { width: 200, height: 60 },
      });
    } catch {
      const cellB3 = sheet.getCell('B3');
      cellB3.value = 'Logo not found';
      cellB3.font = { name: 'Calibri', bold: true, color: { argb: 'FFFF0000' } };
    }
  } else {
    const cellB3 = sheet.getCell('B3');
    cellB3.value = 'Logo not found';
    cellB3.font = { name: 'Calibri', bold: true, size: 11, color: { argb: 'FF888888' } };
  }

  // 2. Row 9:
  // B9 = "Week: {n}"
  const currentWeek = getExcelWeekNum(new Date());
  const cellB9 = sheet.getCell('B9');
  cellB9.value = `Week: ${currentWeek}`;
  cellB9.font = { name: 'Calibri', bold: true, size: 11 };

  // B10 = "Updated by:{name}"
  const cellB10 = sheet.getCell('B10');
  cellB10.value = `Updated by:${options.updatedBy || ''}`;
  cellB10.font = { name: 'Calibri', size: 11, color: { argb: 'FF555555' } };

  // C9:F9 merged: "ISO DRAWING ISSUED HISTORY" (bold, 16, color #003366, centered)
  sheet.mergeCells('C9:F9');
  const titleCell = sheet.getCell('C9');
  titleCell.value = 'ISO DRAWING ISSUED HISTORY';
  titleCell.font = {
    name: 'Calibri',
    bold: true,
    size: 16,
    color: { argb: 'FF003366' },
  };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // C10:F10 merged: Project Name (bold, 12, blue, centered)
  sheet.mergeCells('C10:F10');
  const projectCell = sheet.getCell('C10');
  projectCell.value = options.projectName || 'VARD 9 80 1005';
  projectCell.font = {
    name: 'Calibri',
    bold: true,
    size: 12,
    color: { argb: 'FF0066CC' },
  };
  projectCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // 3. Header row 11 (Columns A..S)
  const headers = [
    'STT',
    'FILE NAME',
    'PIPE NUMBER',
    'ZONE',
    'SYSTEM',
    'PIPE SPOOL',
    'REV',
    'ISSUED DATE\n(dd/mm/yy)',
    'REMARK',
    'Status',
    'BLOCK',
    'UNIT',
    'TANK',
    'COMP',
    'CLASS',
    'OVERBOARD',
    'PHASE',
    'TEMPLATE',
    'CONN. JOINT',
  ];

  const headerRow = sheet.getRow(11);
  headerRow.height = 40;

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } },
  };

  headers.forEach((h, idx) => {
    const colNumber = idx + 1;
    const cell = headerRow.getCell(colNumber);
    cell.value = h;
    cell.font = { name: 'Calibri', bold: true, size: 10, color: { argb: 'FF000000' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = thinBorder;

    // A:J (cols 1..10) fill #C6EFCE, K:S (cols 11..19) fill #CD853F
    if (colNumber <= 10) {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFC6EFCE' },
      };
    } else {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFCD853F' },
      };
    }
  });

  // Enable AutoFilter on row 11
  sheet.autoFilter = 'A11:S11';

  // 4. Data starts at row 12
  let currentRowIdx = 12;

  rows.forEach((row, rIdx) => {
    const r = sheet.getRow(currentRowIdx);
    r.height = 20;

    // Col A: STT
    r.getCell(1).value = row.stt || rIdx + 1;
    // Col B: FILE NAME
    r.getCell(2).value = row.fileName;
    // Col C: PIPE NUMBER
    r.getCell(3).value = row.pipeNumber;
    // Col D: ZONE (as TEXT with leading zeros preserved)
    const cellZone = r.getCell(4);
    cellZone.value = String(row.zone || '');
    cellZone.numFmt = '@';
    // Col E: SYSTEM
    r.getCell(5).value = row.system;
    // Col F: PIPE SPOOL (as TEXT with leading zeros preserved)
    const cellSpool = r.getCell(6);
    cellSpool.value = String(row.pipeSpool || '');
    cellSpool.numFmt = '@';
    // Col G: REV
    r.getCell(7).value = String(row.rev || '0');
    // Col H: ISSUED DATE
    r.getCell(8).value = row.issuedDate;
    // Col I: REMARK
    r.getCell(9).value = row.remark || '';
    // Col J: Status (Value "Official", fill #90EE90, black font)
    const statusCell = r.getCell(10);
    statusCell.value = row.status || 'Official';
    statusCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF90EE90' },
    };
    statusCell.font = { name: 'Calibri', bold: true, color: { argb: 'FF000000' } };
    statusCell.alignment = { horizontal: 'center', vertical: 'middle' };

    // Format all cells in row from Col A to S
    for (let c = 1; c <= 19; c++) {
      const cell = r.getCell(c);
      cell.border = thinBorder;
      if (c !== 10) {
        cell.font = { name: 'Calibri', size: 10 };
      }
      if ([1, 4, 5, 6, 7, 8, 10].includes(c)) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    }

    currentRowIdx++;
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/**
 * Generate standard file name: "{last 4 chars of project name} {source folder name}.xlsx"
 */
export function getIssuedHistoryFileName(projectName: string, sourceFolderName: string): string {
  const pClean = (projectName || '1005').trim();
  const last4 = pClean.length > 4 ? pClean.slice(-4) : pClean;
  const sClean = (sourceFolderName || 'CAD_FILES').trim();
  return `${last4} ${sClean}.xlsx`;
}
