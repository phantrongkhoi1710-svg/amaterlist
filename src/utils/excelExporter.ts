import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';
import { MasterRowData, ImportLogRow } from '../types';

/**
 * Downloads a blob buffer as a file in the browser
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
 * Generates and downloads a beautifully styled Excel workbook with:
 * - "Total" sheet (Master sheet with spacious row heights, colored headers, borders, zebra striping)
 * - "Import_Log" sheet (Formatted log with status tags)
 */
export async function exportToExcel(
  masterHeaders: string[],
  rows: MasterRowData[],
  logs: ImportLogRow[],
  fileName: string = 'Armature_Master_Export.xlsx'
) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Armature Import Tool';
  workbook.lastModifiedBy = 'Armature Import Tool';
  workbook.created = new Date();
  workbook.modified = new Date();

  // ==========================================
  // 1. SHEET: "Total" (MASTER DATA)
  // ==========================================
  const wsTotal = workbook.addWorksheet('Total', {
    views: [
      {
        showGridLines: true,
        state: 'frozen',
        ySplit: 3, // Freeze rows 1-3 so header stays visible
      },
    ],
  });

  const totalCols = Math.max(masterHeaders.length, 6);

  // --- ROW 1: TITLE BANNER ---
  const titleRow = wsTotal.addRow(['BẢNG TỔNG HỢP VẬT TƯ ARMATURE (MASTER DATA)']);
  titleRow.height = 36;
  wsTotal.mergeCells(1, 1, 1, totalCols);
  const titleCell = wsTotal.getCell(1, 1);
  titleCell.font = {
    name: 'Segoe UI',
    size: 14,
    bold: true,
    color: { argb: 'FFFFFFFF' },
  };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E3A8A' }, // Deep Royal Navy
  };
  titleCell.alignment = {
    vertical: 'middle',
    horizontal: 'center',
  };

  // --- ROW 2: SUBTITLE / METADATA INFO ---
  const dateStr = new Date().toLocaleString('vi-VN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
  });
  const infoRow = wsTotal.addRow([
    `Thời gian xuất: ${dateStr}  |  Tổng số thiết bị: ${rows.length} hàng  |  Trạng thái: Đã chuẩn hoá & khớp cột`,
  ]);
  infoRow.height = 22;
  wsTotal.mergeCells(2, 1, 2, totalCols);
  const infoCell = wsTotal.getCell(2, 1);
  infoCell.font = {
    name: 'Segoe UI',
    size: 9.5,
    italic: true,
    color: { argb: 'FF475569' }, // Slate 600
  };
  infoCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF1F5F9' }, // Light Slate
  };
  infoCell.alignment = {
    vertical: 'middle',
    horizontal: 'center',
  };
  infoCell.border = {
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  // --- ROW 3: TABLE HEADERS ---
  const headerRow = wsTotal.addRow([...masterHeaders]);
  headerRow.height = 30; // Generous height for header

  // Define border styles
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'medium', color: { argb: 'FF1E3A8A' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  // Style header cells with distinctive color groups
  masterHeaders.forEach((header, idx) => {
    const colNumber = idx + 1;
    const cell = headerRow.getCell(colNumber);

    // Color grouping for clarity
    const isKeyField = ['TAG', 'SUPPLIER', 'SFI', 'DESCRIPTION'].includes(header);
    const isActuatorOrOrder = ['ACTUATOR TAG', 'PO NUMBER', 'DESTINATION YARD'].includes(header);

    let bgColor = 'FF2563EB'; // Vibrant Blue
    if (isKeyField) {
      bgColor = 'FF1E3A8A'; // Deep Navy for primary ID keys
    } else if (isActuatorOrOrder) {
      bgColor = 'FF0D9488'; // Teal for logistics / order
    } else {
      bgColor = 'FF334155'; // Slate for technical specs
    }

    cell.font = {
      name: 'Segoe UI',
      size: 10.5,
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: bgColor },
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: isKeyField ? 'left' : 'center',
      wrapText: false,
    };
    cell.border = thinBorder;
  });

  // --- ROW 4+: DATA ROWS ---
  const dataBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  };

  rows.forEach((row, rIdx) => {
    const rowValues = masterHeaders.map((header) => {
      const val = row[header];
      return val !== undefined && val !== null ? val : '';
    });

    const dataRow = wsTotal.addRow(rowValues);
    // Generous row height so items never look crowded or stuck together
    dataRow.height = 24;

    // Alternating zebra color
    const isEven = rIdx % 2 === 0;
    const rowBg = isEven ? 'FFFFFFFF' : 'FFF8FAFC'; // Clean white vs very soft ice-slate

    masterHeaders.forEach((header, cIdx) => {
      const cell = dataRow.getCell(cIdx + 1);
      cell.border = dataBorder;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: rowBg },
      };

      const isTag = header === 'TAG';
      const isSupplier = header === 'SUPPLIER';

      cell.font = {
        name: 'Segoe UI',
        size: 10,
        bold: isTag,
        color: isTag
          ? { argb: 'FF1D4ED8' } // Bold Blue for Tag
          : isSupplier
          ? { argb: 'FF047857' } // Emerald for Supplier
          : { argb: 'FF1E293B' }, // Dark slate
      };

      cell.alignment = {
        vertical: 'middle',
        horizontal: isTag || isSupplier || header === 'DESCRIPTION' ? 'left' : 'center',
      };
    });
  });

  // Calculate and apply generous column widths (preventing cramped text)
  masterHeaders.forEach((header, idx) => {
    let maxLen = header.length;
    for (const r of rows.slice(0, 150)) {
      const val = String(r[header] || '');
      if (val.length > maxLen) maxLen = val.length;
    }
    const column = wsTotal.getColumn(idx + 1);
    // Minimum 15 width, + 4 padding for breathing room
    column.width = Math.min(Math.max(maxLen + 5, 16), 48);
  });

  // Add auto-filter on header row
  wsTotal.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: 3, column: masterHeaders.length },
  };

  // ==========================================
  // 2. SHEET: "Import_Log"
  // ==========================================
  const wsLog = workbook.addWorksheet('Import_Log', {
    views: [{ showGridLines: true, state: 'frozen', ySplit: 2 }],
  });

  // --- LOG TITLE BANNER ---
  const logTitleRow = wsLog.addRow(['NHẬT KÝ QUÁ TRÌNH IMPORT DỮ LIỆU (ARMATURE IMPORT LOG)']);
  logTitleRow.height = 32;
  wsLog.mergeCells(1, 1, 1, 7);
  const logTitleCell = wsLog.getCell(1, 1);
  logTitleCell.font = {
    name: 'Segoe UI',
    size: 12,
    bold: true,
    color: { argb: 'FFFFFFFF' },
  };
  logTitleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF334155' }, // Dark Slate
  };
  logTitleCell.alignment = {
    vertical: 'middle',
    horizontal: 'center',
  };

  // --- LOG HEADERS ---
  const logHeaders = [
    'TIME',
    'SOURCE FILE',
    'SOURCE SHEET',
    'HEADER ROW',
    'SOURCE ROWS',
    'IMPORTED ROWS',
    'STATUS',
  ];

  const logHeaderRow = wsLog.addRow(logHeaders);
  logHeaderRow.height = 26;

  logHeaders.forEach((h, idx) => {
    const cell = logHeaderRow.getCell(idx + 1);
    cell.font = {
      name: 'Segoe UI',
      size: 10,
      bold: true,
      color: { argb: 'FFFFFFFF' },
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF475569' },
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
    };
    cell.border = thinBorder;
  });

  // --- LOG DATA ROWS ---
  logs.forEach((log, lIdx) => {
    const isOk = log.status === 'OK';
    const isWarning = log.status.includes('WARNING') || log.status.includes('MISSING');
    const isError = log.status.includes('ERROR');

    const logRow = wsLog.addRow([
      log.time,
      log.sourceFile,
      log.sourceSheet,
      log.headerRow > 0 ? log.headerRow : '—',
      log.sourceRows,
      log.importedRows,
      log.status,
    ]);
    logRow.height = 22;

    const rowBg = lIdx % 2 === 0 ? 'FFFFFFFF' : 'FFF8FAFC';

    for (let c = 1; c <= 7; c++) {
      const cell = logRow.getCell(c);
      cell.border = dataBorder;
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: rowBg },
      };
      cell.alignment = {
        vertical: 'middle',
        horizontal: c === 1 || c === 2 || c === 7 ? 'left' : 'center',
      };

      if (c === 7) {
        // Status styling
        cell.font = {
          name: 'Segoe UI',
          size: 9.5,
          bold: true,
          color: isOk
            ? { argb: 'FF15803D' } // Green
            : isWarning
            ? { argb: 'FFB45309' } // Amber
            : isError
            ? { argb: 'FFB91C1C' } // Red
            : { argb: 'FF334155' },
        };
      } else {
        cell.font = { name: 'Segoe UI', size: 9.5, color: { argb: 'FF334155' } };
      }
    }
  });

  // Column widths for Log
  wsLog.columns = [
    { width: 22 }, // TIME
    { width: 34 }, // SOURCE FILE
    { width: 20 }, // SOURCE SHEET
    { width: 14 }, // HEADER ROW
    { width: 15 }, // SOURCE ROWS
    { width: 16 }, // IMPORTED ROWS
    { width: 48 }, // STATUS
  ];

  // Write and trigger download
  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(buffer, fileName);
}

/**
 * Export only the master table as CSV
 */
export function exportToCSV(
  masterHeaders: string[],
  rows: MasterRowData[],
  fileName: string = 'Armature_Total.csv'
) {
  const data = rows.map((r) => {
    const obj: Record<string, any> = {};
    for (const h of masterHeaders) {
      obj[h] = r[h] !== undefined && r[h] !== null ? r[h] : '';
    }
    return obj;
  });

  const ws = XLSX.utils.json_to_sheet(data, { header: masterHeaders });
  const csvOutput = XLSX.utils.sheet_to_csv(ws);

  const blob = new Blob(['\uFEFF' + csvOutput], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
