import { MasterRowData, RowRevisionStatus } from '../types';

/**
 * PRINCIPLE FOR REVISION DESCRIPTIONS:
 * 1. New or updated valve/armature:
 *    -> Yellow coloured row
 * 2. Deleted valve/armature:
 *    -> Red colour with strikethrough (delete line)
 * 3. Next revision after deleted valve/armature:
 *    -> Keep SFI and TAG (rest of cells are empty)
 */

export const REVISION_PRINCIPLES = {
  NEW_UPDATED: {
    label: 'Mới / Cập nhật (Yellow)',
    colorClass: 'bg-yellow-200 dark:bg-yellow-950/70 border-yellow-400 text-slate-950 dark:text-yellow-100',
    badgeClass: 'bg-yellow-300 text-slate-950 border-yellow-500 font-bold',
    hexBg: '#FFFF00',
    description: 'Yellow coloured row (Van mới hoặc đã cập nhật thông số)',
  },
  DELETED: {
    label: 'Đã xóa (Red + Strikethrough)',
    colorClass: 'bg-rose-200 dark:bg-rose-950/70 border-rose-500 text-rose-950 dark:text-rose-100 line-through',
    badgeClass: 'bg-rose-500 text-white font-bold line-through',
    hexBg: '#EF4444',
    description: 'Red colour with strikethrough (Van bị xóa bỏ khỏi sơ đồ)',
  },
  NEXT_REV_AFTER_DELETED: {
    label: 'Next Rev sau khi xóa (Keep SFI & TAG)',
    colorClass: 'bg-slate-100 dark:bg-slate-800/80 border-slate-300 dark:border-slate-700 text-slate-500 italic',
    badgeClass: 'bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 font-semibold',
    hexBg: '#F1F5F9',
    description: 'Keep SFI and TAG (rest of cells are empty)',
  },
};

/**
 * Extracts a clean #RRGGBB from an ARGB or hex string
 */
export function extractHexColor(argbOrHex?: string): string | null {
  if (!argbOrHex) return null;
  const clean = String(argbOrHex).replace('#', '').trim().toUpperCase();
  if (clean.length === 8) {
    // AARRGGBB -> #RRGGBB
    return `#${clean.substring(2)}`;
  }
  if (clean.length === 6) {
    return `#${clean}`;
  }
  if (clean.length === 3) {
    return `#${clean[0]}${clean[0]}${clean[1]}${clean[1]}${clean[2]}${clean[2]}`;
  }
  return null;
}

/**
 * Checks if a hex color or argb is yellow-ish
 */
export function isYellowColor(argbOrHex?: string): boolean {
  if (!argbOrHex) return false;
  const clean = String(argbOrHex).replace('#', '').trim().toUpperCase();
  const knownYellows = [
    'FFFF00', 'FFFFFF00', 'FEF08A', 'FFFFFEF08A', 'FDE047', 'FFFFFDE047',
    'FACC15', 'FFFFFACC15', 'FFF200', 'FFFFFFF200', 'FFF59D', 'FFFFF59D',
    'FFFFE599', 'FFE599', 'FFFFD966', 'FFD966', 'FFFFF2CC', 'FFF2CC',
    'FFFFE0', 'FFFFFFE0', 'FFFFCC', 'FFFFFFCC', 'FFF9C4', 'FFFFF9C4',
    'FFF176', 'FFFFF176', 'FFE082', 'FFFFE082', 'FFD54F', 'FFFFD54F',
    'FFEE58', 'FFFFEE58', 'FFEB3B', 'FFFFEB3B', 'FDD835', 'FFFFD835',
    'FBC02D', 'FFFFC02D', 'EAB308', 'FFEAB308', 'CA8A04', 'FFCA8A04'
  ];
  if (knownYellows.includes(clean)) {
    return true;
  }
  // If 8-char ARGB, extract RGB
  if (clean.length === 8) {
    const r = parseInt(clean.substring(2, 4), 16);
    const g = parseInt(clean.substring(4, 6), 16);
    const b = parseInt(clean.substring(6, 8), 16);
    // Yellow has strong Red and Green, with significantly lower Blue
    return r >= 170 && g >= 150 && b < 160 && (r + g) > (b * 2.2);
  }
  if (clean.length === 6) {
    const r = parseInt(clean.substring(0, 2), 16);
    const g = parseInt(clean.substring(2, 4), 16);
    const b = parseInt(clean.substring(4, 6), 16);
    return r >= 170 && g >= 150 && b < 160 && (r + g) > (b * 2.2);
  }
  return false;
}

/**
 * Checks if a hex color or argb is red-ish (e.g. deleted line)
 */
export function isRedColor(argbOrHex?: string): boolean {
  if (!argbOrHex) return false;
  const clean = String(argbOrHex).replace('#', '').trim().toUpperCase();
  const knownReds = [
    'FF0000', 'FFFF0000', 'EF4444', 'FFEF4444', 'F87171', 'FFF87171',
    'FCA5A5', 'FFFFCA5A5', 'DC2626', 'FFDC2626', 'E11D48', 'FFE11D48',
    'FF4D4D', 'FFFFFF4D4D', 'FFC7CE', 'FFFFC7CE', 'FFB8B8', 'FFFFB8B8',
    'FFFF9999', 'FF9999', 'FFE53E3E', 'E53E3E', 'B91C1C', 'FFB91C1C',
    '991B1B', 'FF991B1B', '7F1D1D', 'FF7F1D1D', 'F43F5E', 'FFF43F5E',
    'E11D48', 'FFE11D48', 'BE123C', 'FFBE123C'
  ];
  if (knownReds.includes(clean)) {
    return true;
  }
  if (clean.length === 8) {
    const r = parseInt(clean.substring(2, 4), 16);
    const g = parseInt(clean.substring(4, 6), 16);
    const b = parseInt(clean.substring(6, 8), 16);
    // Red has high Red, with low Green and Blue
    return r >= 170 && g < 150 && b < 150 && r > (g * 1.3) && r > (b * 1.3);
  }
  if (clean.length === 6) {
    const r = parseInt(clean.substring(0, 2), 16);
    const g = parseInt(clean.substring(2, 4), 16);
    const b = parseInt(clean.substring(4, 6), 16);
    return r >= 170 && g < 150 && b < 150 && r > (g * 1.3) && r > (b * 1.3);
  }
  return false;
}

/**
 * Applies a specific revision status to a row following the principle
 */
export function applyRowRevisionStatus(
  row: MasterRowData,
  status: RowRevisionStatus
): MasterRowData {
  const updated: MasterRowData = { ...row };

  switch (status) {
    case 'new_updated':
      updated._revisionStatus = 'new_updated';
      updated._detectedColor = '#FFFF00';
      updated._hasStrikethrough = false;
      if (!updated['REV. DESCRIPTION']) {
        updated['REV. DESCRIPTION'] = 'New/updated valve';
      }
      break;

    case 'deleted':
      updated._revisionStatus = 'deleted';
      updated._detectedColor = '#EF4444';
      updated._hasStrikethrough = true;
      if (!updated['REV. DESCRIPTION']) {
        updated['REV. DESCRIPTION'] = 'Deleted valve';
      }
      break;

    case 'next_rev_after_deleted': {
      // RULE 3: "Keep SFI and TAG (rest of cells are empty)"
      const tag = row.TAG || '';
      const sfi = row.SFI || '';
      const revHis = row['REV. HIS.'] || '';
      const revDate = row.DATE || new Date().toISOString().slice(0, 10);
      const revDesc = 'Deleted in previous revision';
      const sig = row.SIGNATURE || '';

      // Clear all properties on row except system props and SFI + TAG + Rev fields
      const preservedKeys = new Set([
        '_id',
        '_sourceFile',
        '_sourceSheet',
        '_importedAt',
        'TAG',
        'SFI',
        'REV. HIS.',
        'DATE',
        'REV. DESCRIPTION',
        'SIGNATURE',
      ]);

      for (const key of Object.keys(updated)) {
        if (!key.startsWith('_') && !preservedKeys.has(key)) {
          updated[key] = '';
        }
      }

      updated.TAG = tag;
      updated.SFI = sfi;
      updated['REV. HIS.'] = revHis;
      updated.DATE = revDate;
      updated['REV. DESCRIPTION'] = revDesc;
      updated.SIGNATURE = sig;
      updated._revisionStatus = 'next_rev_after_deleted';
      updated._detectedColor = undefined;
      updated._hasStrikethrough = false;
      break;
    }

    case 'normal':
    default:
      updated._revisionStatus = 'normal';
      updated._detectedColor = undefined;
      updated._hasStrikethrough = false;
      break;
  }

  return updated;
}
