/**
 * Pure business logic for ISO Drawing Manager
 * Ported from original VBA workbook with web-native enhancements.
 */

export interface IsoDrawingRow {
  id: string;
  stt: number;
  fileName: string; // Base name without extension (e.g. N1005-212-P-CS-002-59.1-821-001)
  pipeNumber: string; // Main part before underscore
  zone: string; // e.g. "212"
  system: string; // e.g. "821-002"
  pipeSpool: string; // e.g. "59.1-001"
  rev: string; // e.g. "0", "A", "1"
  issuedDate: string; // dd/mm/yyyy
  remark: string;
  status: string; // e.g. "Official"
  systemShort: string; // e.g. "8212"
  hasDWG: boolean;
  hasDXF: boolean;
  checkStatus: string; // "Valid", "Invalid name format", "DWG not found", etc.
  isValid: boolean;
  targetFolder?: string; // e.g. "S212/8212 - SEAWATER SYSTEM"
}

/**
 * Format a Date object to dd/mm/yyyy
 */
export function formatToDDMMYYYY(date: Date = new Date()): string {
  const d = String(date.getDate()).padStart(2, '0');
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const y = date.getFullYear();
  return `${d}/${m}/${y}`;
}

/**
 * Compute Excel-compatible WEEKNUM(date, 2) where Monday is first day of week.
 */
export function getExcelWeekNum(d: Date = new Date()): number {
  const target = new Date(d.valueOf());
  // ISO-8601 week number calculation (Monday start)
  const dayNr = (d.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = target.valueOf();
  target.setMonth(0, 1);
  if (target.getDay() !== 4) {
    target.setMonth(0, 1 + ((4 - target.getDay() + 7) % 7));
  }
  return 1 + Math.ceil((firstThursday - target.valueOf()) / 604800000);
}

/**
 * Parse an ISO drawing base name into technical attributes.
 * Example base name: "N1005-212-P-CS-002-59.1-821-001_0" or "N1005-212-P-CS-002-59.1-821-001"
 */
export function parseIsoFileName(baseName: string, defaultDate?: string): {
  pipeNumber: string;
  zone: string;
  system: string;
  pipeSpool: string;
  rev: string;
  issuedDate: string;
  status: string;
  systemShort: string;
  isValid: boolean;
  errorMessage?: string;
} {
  const todayStr = defaultDate || formatToDDMMYYYY();
  const trimmed = baseName.trim();

  // Split by "_" -> mainPart = [0], REV = [1] if exists else "0"
  const underParts = trimmed.split('_');
  const mainPart = underParts[0] || '';
  const rev = underParts.length > 1 && underParts[1] ? underParts[1] : '0';

  // Split main part by "-"
  const arr = mainPart.split('-');

  if (arr.length < 8) {
    return {
      pipeNumber: mainPart,
      zone: arr[1] || '',
      system: '',
      pipeSpool: '',
      rev,
      issuedDate: todayStr,
      status: 'Official',
      systemShort: '',
      isValid: false,
      errorMessage: `Invalid name format (${arr.length}/8 required segments)`,
    };
  }

  // Example: N1005-212-P-CS-002-59.1-821-001
  // arr[0] = N1005
  // arr[1] = 212       (ZONE)
  // arr[2] = P
  // arr[3] = CS
  // arr[4] = 002
  // arr[5] = 59.1
  // arr[6] = 821
  // arr[7] = 001
  const zone = String(arr[1]);
  const system = `${arr[6]}-${arr[4]}`;
  const pipeSpool = `${arr[5]}-${arr[7]}`;
  const systemShort = computeSystemShort(system);

  return {
    pipeNumber: mainPart,
    zone,
    system,
    pipeSpool,
    rev,
    issuedDate: todayStr,
    status: 'Official',
    systemShort,
    isValid: true,
  };
}

/**
 * Compute System Short Code (Column X in original VBA).
 * Logic:
 * raw = SYSTEM (e.g. "821-002")
 * part1 = first 3 chars ("821")
 * part2 = last 3 chars ("002") or part after hyphen
 * Remove every "0" character from part2 -> "2"
 * systemShort = part1 + remaining part2 -> "8212"
 */
export function computeSystemShort(system: string): string {
  if (!system) return '';
  const trimmed = system.trim();
  const dashIdx = trimmed.indexOf('-');

  let part1 = '';
  let part2 = '';

  if (dashIdx !== -1) {
    part1 = trimmed.slice(0, dashIdx).slice(0, 3);
    part2 = trimmed.slice(dashIdx + 1);
    if (part2.length > 3) {
      part2 = part2.slice(-3);
    }
  } else {
    part1 = trimmed.slice(0, 3);
    part2 = trimmed.slice(3);
  }

  // Remove all '0' characters from part2
  const cleanedPart2 = part2.replace(/0/g, '');
  return `${part1}${cleanedPart2}`;
}

/**
 * Match template folder name by systemShort code.
 * If template subfolder name CONTAINS systemShort (case-insensitive),
 * use that template subfolder name; otherwise use systemShort.
 */
export function matchTemplateFolder(systemShort: string, templateFolders: string[]): string {
  if (!systemShort) return 'UNKNOWN';
  if (!templateFolders || templateFolders.length === 0) return systemShort;

  const needle = systemShort.toLowerCase();
  const matched = templateFolders.find((folder) => folder.toLowerCase().includes(needle));
  return matched || systemShort;
}

/**
 * Compare two revision strings.
 * Compares numerically if both can be parsed as numbers, otherwise alphabetically.
 */
export function compareRevisions(revA: string, revB: string): number {
  const cleanA = String(revA || '').trim();
  const cleanB = String(revB || '').trim();

  const numA = Number(cleanA);
  const numB = Number(cleanB);

  const isNumA = !Number.isNaN(numA) && cleanA !== '';
  const isNumB = !Number.isNaN(numB) && cleanB !== '';

  if (isNumA && isNumB) {
    return numA - numB;
  }

  return cleanA.localeCompare(cleanB, undefined, { numeric: true, sensitivity: 'base' });
}

/**
 * Mark old revisions logic (port of Tool 7: Update Old Rev).
 * For each group key, find the row with the max revision.
 * Mark every other row in that group with "old rev" in the mark column.
 */
export interface OldRevRow {
  [colIndex: number]: unknown;
}

export function markOldRevisionsInMatrix(
  matrix: unknown[][],
  startRowIndex: number, // 0-based row index where data starts (e.g. 1 for row 2)
  keyColIndex: number, // 0-based col index
  revColIndex: number, // 0-based col index
  markColIndex: number, // 0-based col index
): { updatedMatrix: unknown[][]; markedCount: number; totalProcessed: number } {
  const updatedMatrix = matrix.map((r) => [...r]);
  const groupMap = new Map<string, { rowIndex: number; rev: string }[]>();

  let totalProcessed = 0;

  for (let i = startRowIndex; i < updatedMatrix.length; i++) {
    const row = updatedMatrix[i];
    if (!row) continue;

    const rawKey = row[keyColIndex];
    const rawRev = row[revColIndex];

    if (rawKey === undefined || rawKey === null || String(rawKey).trim() === '') {
      continue;
    }
    if (rawRev === undefined || rawRev === null || String(rawRev).trim() === '') {
      continue;
    }

    const key = String(rawKey).trim();
    const rev = String(rawRev).trim();

    if (!groupMap.has(key)) {
      groupMap.set(key, []);
    }
    groupMap.get(key)!.push({ rowIndex: i, rev });
    totalProcessed++;
  }

  let markedCount = 0;

  groupMap.forEach((items) => {
    if (items.length <= 1) return;

    // Find highest revision
    let maxItem = items[0];
    for (let j = 1; j < items.length; j++) {
      if (compareRevisions(items[j].rev, maxItem.rev) > 0) {
        maxItem = items[j];
      }
    }

    // Mark every other row with "old rev"
    items.forEach((item) => {
      if (item.rowIndex !== maxItem.rowIndex) {
        // Expand row if needed
        while (updatedMatrix[item.rowIndex].length <= markColIndex) {
          updatedMatrix[item.rowIndex].push('');
        }
        updatedMatrix[item.rowIndex][markColIndex] = 'old rev';
        markedCount++;
      }
    });
  });

  return { updatedMatrix, markedCount, totalProcessed };
}
