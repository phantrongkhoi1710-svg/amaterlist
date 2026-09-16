/**
 * Replicates NormalizeHeader from the VBA script:
 * - Replaces non-breaking space (Chr 160) with regular space
 * - Replaces \r, \n, \t with space
 * - Converts to UPPERCASE
 * - Removes '.' and '°'
 * - Replaces '_', '-', '/' with space
 * - Collapses multiple spaces to a single space
 * - Trims leading/trailing whitespace
 */
export function normalizeHeader(val: any): string {
  if (val === null || val === undefined) return '';
  let s = String(val);

  // Replace non-breaking space
  s = s.replace(/\u00a0/g, ' ');

  // Replace line breaks and tabs
  s = s.replace(/[\r\n\t]+/g, ' ');

  // Uppercase
  s = s.toUpperCase();

  // Remove specific punctuation like . and °
  s = s.replace(/\./g, '');
  s = s.replace(/°/g, '');

  // Replace separators with space
  s = s.replace(/[_\-\/]/g, ' ');

  // Collapse multiple spaces
  s = s.replace(/\s+/g, ' ');

  return s.trim();
}

/**
 * Default standard aliases from VBA Armature Import macro
 */
export const DEFAULT_HEADER_ALIASES: Record<string, string[]> = {
  'ACTUATOR TAG': [
    'ACTUATOR',
    'ACTUATOR TAG NO',
    'ACTUATOR TAG NUMBER',
    'ACTUATOR NO',
  ],
  'PO NUMBER': [
    'PO NO',
    'PURCHASE ORDER',
    'PURCHASE ORDER NUMBER',
    'PURCHASE ORDER NO',
  ],
  'DESTINATION YARD': ['DESTINATION', 'DESTINATION YARD ACTUATOR'],
  'STD DRW NORMALE N': ['STD DRAWING', 'STANDARD DRAWING', 'STANDARD DRW'],
  'PRESSURE RATING': ['PRESSURE', 'PRESSURE CLASS', 'RATING'],
  'HOUSING BODY': ['HOUSING', 'BODY'],
  'PIPE CLASS': ['PIPE CLASSIFICATION', 'CLASS'],
  'CLASS CERTIFICATE': ['CLASS CERT', 'CERTIFICATE', 'CLASS CERTIFICATION'],
  'SIGN TYPE': ['SIGNTYPE'],
  'SIGN TEXT': ['SIGNTEXT'],
  'INPUT SIGN TEXT': ['INPUT SIGN', 'SIGN INPUT TEXT', 'INPUT SIGNTEXT'],
  'REV HIS': ['REVISION', 'REV', 'REV HISTORY', 'REVISION HISTORY'],
};

/**
 * Standard default master columns in Armature projects
 */
export const DEFAULT_MASTER_COLUMNS: string[] = [
  'TAG',
  'SUPPLIER',
  'SFI',
  'DESCRIPTION',
  'ACTUATOR TAG',
  'PO NUMBER',
  'DESTINATION YARD',
  'STD DRW NORMALE N',
  'PRESSURE RATING',
  'HOUSING BODY',
  'PIPE CLASS',
  'CLASS CERTIFICATE',
  'SIGN TYPE',
  'SIGN TEXT',
  'INPUT SIGN TEXT',
  'REV HIS',
];

/**
 * Get aliases for a normalized header name
 */
export function getAliases(
  normalizedHeader: string,
  customAliases?: Record<string, string[]>
): string[] {
  const norm = normalizeHeader(normalizedHeader);
  const aliasMap = customAliases || DEFAULT_HEADER_ALIASES;
  return aliasMap[norm] || [];
}

/**
 * Calculates the occurrence (1-based index) of a header in a list of master headers
 * e.g., if DESCRIPTION appears 2 times, the first is occurrence 1, second is occurrence 2
 */
export function getHeaderOccurrence(headers: string[], targetIndex: number): number {
  const targetHeader = normalizeHeader(headers[targetIndex]);
  if (!targetHeader) return 1;

  let count = 0;
  for (let i = 0; i <= targetIndex; i++) {
    if (normalizeHeader(headers[i]) === targetHeader) {
      count++;
    }
  }
  return count;
}

/**
 * Matches a master header & occurrence to a column in sourceMap
 * sourceMap: key = normalizedHeader -> array of 0-based column indices
 */
export function getSourceColumnIndex(
  masterHeader: string,
  occurrence: number,
  sourceMap: Map<string, number[]>,
  customAliases?: Record<string, string[]>
): number | null {
  const key = normalizeHeader(masterHeader);

  // 1. Exact match
  if (sourceMap.has(key)) {
    const cols = sourceMap.get(key)!;
    if (occurrence <= cols.length) {
      return cols[occurrence - 1];
    }
  }

  // 2. Alias match
  const aliases = getAliases(key, customAliases);
  for (const alias of aliases) {
    const normAlias = normalizeHeader(alias);
    if (sourceMap.has(normAlias)) {
      const cols = sourceMap.get(normAlias)!;
      if (occurrence <= cols.length) {
        return cols[occurrence - 1];
      }
    }
  }

  return null;
}
