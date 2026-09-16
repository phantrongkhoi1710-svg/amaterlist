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
        ySplit: 4, // Freeze rows 1-4 so super-header and column headers stay visible
      },
    ],
  });

  const totalCols = Math.max(masterHeaders.length, 6);

  // --- ROW 1: TITLE BANNER ---
  const titleRow = wsTotal.addRow(['BẢNG TỔNG HỢP VẬT TƯ ARMATURE (MASTER DATA)']);
  titleRow.height = 34;
  wsTotal.mergeCells(1, 1, 1, totalCols);
  const titleCell = wsTotal.getCell(1, 1);
  titleCell.font = {
    name: 'Segoe UI',
    size: 13,
    bold: true,
    color: { argb: 'FFFFFFFF' },
  };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E293B' }, // Slate 800
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
    `Thời gian xuất: ${dateStr}  |  Tổng số thiết bị: ${rows.length} hàng  |  Phân tích từ Catalog & Đặc tính kỹ thuật ống`,
  ]);
  infoRow.height = 20;
  wsTotal.mergeCells(2, 1, 2, totalCols);
  const infoCell = wsTotal.getCell(2, 1);
  infoCell.font = {
    name: 'Segoe UI',
    size: 9,
    italic: true,
    color: { argb: 'FF475569' }, // Slate 600
  };
  infoCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF8FAFC' }, // Light Slate
  };
  infoCell.alignment = {
    vertical: 'middle',
    horizontal: 'center',
  };
  infoCell.border = {
    bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  // --- ROW 3: TWO GROUPED SUPER-HEADERS (YELLOW & GRAY AS IN PIPE SPECIFICATION) ---
  // Determine split point (where Pipe Specification begins, e.g. STD DRW NORMALE N° or column 5)
  const specColIdx = masterHeaders.findIndex(
    (h) =>
      h.toUpperCase().includes('STD DRW') ||
      h.toUpperCase().includes('NORMALE') ||
      h.toUpperCase() === 'EXECUTION'
  );
  const splitCol = specColIdx !== -1 ? specColIdx : 5; // Default 5 columns for functional design

  const superHeaderRow = wsTotal.addRow(new Array(totalCols).fill(''));
  superHeaderRow.height = 26;

  // Group 1: Functional Design (Yellow)
  const functionalDesignCols = Math.min(splitCol, totalCols);
  if (functionalDesignCols > 0) {
    wsTotal.mergeCells(3, 1, 3, functionalDesignCols);
    const g1Cell = wsTotal.getCell(3, 1);
    g1Cell.value = 'TO BE COMPLETED BY FUNCTIONAL DESIGN';
    g1Cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF000000' } };
    g1Cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } }; // Pure Yellow #FFFF00
    g1Cell.alignment = { vertical: 'middle', horizontal: 'center' };
    g1Cell.border = {
      top: { style: 'medium', color: { argb: 'FF000000' } },
      left: { style: 'medium', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'medium', color: { argb: 'FF000000' } },
    };
  }

  // Group 2: Armature info from pipe specification (Gray)
  if (totalCols > functionalDesignCols) {
    wsTotal.mergeCells(3, functionalDesignCols + 1, 3, totalCols);
    const g2Cell = wsTotal.getCell(3, functionalDesignCols + 1);
    g2Cell.value = 'ARMATURE INFO FROM PIPE SPECIFICATION';
    g2Cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FF000000' } };
    g2Cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBFBFBF' } }; // Metallic Gray #BFBFBF
    g2Cell.alignment = { vertical: 'middle', horizontal: 'center' };
    g2Cell.border = {
      top: { style: 'medium', color: { argb: 'FF000000' } },
      left: { style: 'medium', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
      right: { style: 'medium', color: { argb: 'FF000000' } },
    };
  }

  // --- ROW 4: INDIVIDUAL COLUMN HEADERS ---
  const headerRow = wsTotal.addRow([...masterHeaders]);
  headerRow.height = 28;

  const headerBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
    bottom: { style: 'medium', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
  };

  masterHeaders.forEach((header, idx) => {
    const colNumber = idx + 1;
    const cell = headerRow.getCell(colNumber);
    const isFunctionalCol = colNumber <= functionalDesignCols;

    cell.font = {
      name: 'Segoe UI',
      size: 9.5,
      bold: true,
      color: { argb: 'FF000000' },
    };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: isFunctionalCol ? 'FFFFFF55' : 'FFD9D9D9' }, // Light yellow vs soft gray
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: false,
    };
    cell.border = headerBorder;
  });

  // --- ROW 4+: DATA ROWS ---
  const dataBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
  };
  const thinBorder = dataBorder;

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

/**
 * Export a single updated valve details & changes into a dedicated Excel file
 */
export async function exportValveChangeToExcel(
  masterHeaders: string[],
  row: MasterRowData,
  changes: { field: string; oldValue: any; newValue: any }[],
  customFileName?: string
): Promise<string> {
  const tag = row.TAG || 'Armature';
  const outName =
    customFileName ||
    `Phieu_CapNhat_Van_${String(tag).replace(/[^a-zA-Z0-9_-]/g, '_')}.xlsx`;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Armature Import Tool';
  workbook.created = new Date();

  const ws = workbook.addWorksheet('ChiTiet_CapNhat', {
    views: [{ showGridLines: true }],
  });

  // Row 1: Title
  const titleRow = ws.addRow(['PHIẾU XÁC NHẬN CẬP NHẬT THÔNG SỐ VẬT TƯ VAN']);
  titleRow.height = 36;
  ws.mergeCells(1, 1, 1, 4);
  const titleCell = ws.getCell(1, 1);
  titleCell.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1E3A8A' }, // Deep Blue
  };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  // Row 2: Subtitle
  const subRow = ws.addRow([
    `Mã TAG: ${tag}  |  Nhà cung cấp: ${row.SUPPLIER || 'Chưa rõ'}  |  Thời gian: ${new Date().toLocaleString('vi-VN')}  |  Nguồn: ${row._sourceFile || 'Master Data'}`,
  ]);
  subRow.height = 22;
  ws.mergeCells(2, 1, 2, 4);
  const subCell = ws.getCell(2, 1);
  subCell.font = { name: 'Segoe UI', size: 9.5, italic: true, color: { argb: 'FF1E293B' } };
  subCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
  subCell.alignment = { horizontal: 'left', vertical: 'middle' };

  ws.addRow([]); // Blank line

  // Section 1: Changes table (if any)
  if (changes && changes.length > 0) {
    const sec1 = ws.addRow(['1. DANH SÁCH THÔNG SỐ VỪA THAY ĐỔI (RECENT CHANGES)']);
    sec1.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFB45309' } };
    ws.addRow([]);

    const chgHeader = ws.addRow([
      'STT',
      'TÊN THÔNG SỐ (FIELD)',
      'GIÁ TRỊ TRƯỚC ĐÓ (OLD)',
      'GIÁ TRỊ CẬP NHẬT MỚI (NEW)',
    ]);
    chgHeader.height = 26;
    chgHeader.eachCell((cell) => {
      cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD97706' } }; // Amber 600
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFB45309' } },
        bottom: { style: 'thin', color: { argb: 'FFB45309' } },
        left: { style: 'thin', color: { argb: 'FFB45309' } },
        right: { style: 'thin', color: { argb: 'FFB45309' } },
      };
    });

    changes.forEach((c, idx) => {
      const r = ws.addRow([
        idx + 1,
        c.field,
        String(c.oldValue ?? '—'),
        String(c.newValue ?? '—'),
      ]);
      r.height = 22;
      r.eachCell((cell, colNum) => {
        cell.font = {
          name: 'Segoe UI',
          size: 10,
          bold: colNum === 4,
          color: colNum === 4 ? { argb: 'FF15803D' } : { argb: 'FF334155' },
        };
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: colNum === 4 ? 'FFECFDF5' : 'FFFFFBEB' },
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        };
        if (colNum === 1) cell.alignment = { horizontal: 'center' };
      });
    });

    ws.addRow([]); // Blank line
  }

  // Section 2: All Specifications of this valve
  const sec2 = ws.addRow([
    changes && changes.length > 0
      ? '2. BẢNG TOÀN BỘ THÔNG SỐ KỸ THUẬT HIỆN TẠI (SPECIFICATIONS)'
      : '1. BẢNG TOÀN BỘ THÔNG SỐ KỸ THUẬT HIỆN TẠI (SPECIFICATIONS)',
  ]);
  sec2.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF1E3A8A' } };
  ws.addRow([]);

  const specHeader = ws.addRow([
    'STT',
    'TÊN TRƯỜNG DỮ LIỆU (HEADER)',
    'GIÁ TRỊ THIẾT KẾ (SPEC VALUE)',
    'GHI CHÚ',
  ]);
  specHeader.height = 26;
  specHeader.eachCell((cell) => {
    cell.font = { name: 'Segoe UI', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF0F172A' } },
      bottom: { style: 'thin', color: { argb: 'FF0F172A' } },
      left: { style: 'thin', color: { argb: 'FF0F172A' } },
      right: { style: 'thin', color: { argb: 'FF0F172A' } },
    };
  });

  masterHeaders.forEach((h, idx) => {
    const val = row[h];
    const isModified = changes?.some((c) => c.field === h);
    const r = ws.addRow([
      idx + 1,
      h,
      val !== undefined && val !== null ? String(val) : '—',
      isModified ? 'Vừa điều chỉnh' : '',
    ]);
    r.height = 20;
    r.eachCell((cell, colNum) => {
      cell.font = {
        name: 'Segoe UI',
        size: 9.5,
        bold: isModified || colNum === 2,
        color: isModified
          ? { argb: 'FF1E40AF' }
          : colNum === 4
          ? { argb: 'FFD97706' }
          : { argb: 'FF334155' },
      };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {
          argb: isModified
            ? 'FFEFF6FF'
            : idx % 2 === 0
            ? 'FFFFFFFF'
            : 'FFF8FAFC',
        },
      };
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      };
      if (colNum === 1 || colNum === 4) cell.alignment = { horizontal: 'center' };
    });
  });

  ws.columns = [
    { width: 8 },
    { width: 34 },
    { width: 36 },
    { width: 22 },
  ];

  const buffer = await workbook.xlsx.writeBuffer();
  downloadBlob(buffer, outName);
  return outName;
}

/**
 * Copies rich HTML table to clipboard so it can be pasted with colors & borders directly into Outlook
 */
export async function copyHtmlTableToClipboard(htmlString: string, plainText: string): Promise<boolean> {
  try {
    if (typeof ClipboardItem !== 'undefined' && navigator.clipboard && navigator.clipboard.write) {
      const blobHtml = new Blob([htmlString], { type: 'text/html' });
      const blobText = new Blob([plainText], { type: 'text/plain' });
      await navigator.clipboard.write([
        new ClipboardItem({
          'text/html': blobHtml,
          'text/plain': blobText,
        }),
      ]);
      return true;
    } else {
      await navigator.clipboard.writeText(plainText);
      return true;
    }
  } catch (err) {
    console.warn('ClipboardItem write failed, fallback to plain text:', err);
    await navigator.clipboard.writeText(plainText);
    return true;
  }
}
