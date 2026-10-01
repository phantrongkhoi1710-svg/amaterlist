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

  // Replace separators and brackets with space
  s = s.replace(/[_\-\/\(\)]/g, ' ');

  // Collapse multiple spaces
  s = s.replace(/\s+/g, ' ');

  return s.trim();
}

/**
 * Standard default master columns in Armature projects matching pipe specification and uploaded image:
 * Total: 26 columns
 * 1-5: TO BE COMPLETED BY FUNCTIONAL DESIGN (Yellow)
 * 6-21: ARMATURE INFO FROM PIPE SPECIFICATION (Grey)
 * 22-26: REVISION / SIGN-OFF HISTORY (White)
 */
export const DEFAULT_MASTER_COLUMNS: string[] = [
  'TAG',
  'ACTUATOR TAG',
  'P.O. NUMBER',
  'SUPPLIER',
  'DESTINATION (YARD)',
  'SAP CODE',
  'STD DRW NORMALE N°',
  'EXECUTION',
  'NRF N°',
  'SFI',
  'DESCRIPTION',
  'SIZE',
  'CONNECTION',
  'PRESSURE RATING',
  'HOUSING /BODY',
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

/**
 * Super Header Groups definition matching the user's specification image
 */
export const SUPER_HEADER_GROUPS = [
  {
    name: 'TO BE COMPLETED BY FUNCTIONAL DESIGN',
    color: '#FFFF00', // Yellow
    textColor: '#000000',
    columns: [
      'TAG',
      'ACTUATOR TAG',
      'P.O. NUMBER',
      'SUPPLIER',
      'DESTINATION (YARD)',
    ],
  },
  {
    name: 'ARMATURE INFO FROM PIPE SPECIFICATION',
    color: '#BFBFBF', // Metallic Grey
    textColor: '#000000',
    columns: [
      'SAP CODE',
      'STD DRW NORMALE N°',
      'EXECUTION',
      'NRF N°',
      'SFI',
      'DESCRIPTION',
      'SIZE',
      'CONNECTION',
      'PRESSURE RATING',
      'HOUSING /BODY',
      'TYPE',
      'PIPE CLASS',
      'CLASS CERTIFICATE',
      'REMARKS',
      'SIGN TYPE',
      'SIGN TEXT',
    ],
  },
  {
    name: 'REVISION & SIGNATURE',
    color: '#FFFFFF', // White
    textColor: '#000000',
    columns: [
      'INPUT - SIGN TEXT',
      'REV. HIS.',
      'DATE',
      'REV. DESCRIPTION',
      'SIGNATURE',
    ],
  },
];

/**
 * Default standard aliases from VBA Armature Import macro and pipe specs
 */
export const DEFAULT_HEADER_ALIASES: Record<string, string[]> = {
  'TAG': ['TAG NO', 'TAG NUMBER', 'VALVE TAG', 'EQUIPMENT TAG', 'TAGNAME', 'TAG_NO', 'ITEM TAG'],
  'ACTUATOR TAG': [
    'ACTUATOR',
    'ACTUATOR TAG NO',
    'ACTUATOR TAG NUMBER',
    'ACTUATOR NO',
    'ACTUATOR_TAG',
  ],
  'PO NUMBER': [
    'PO NO',
    'PO',
    'P.O. NUMBER',
    'P.O. NO',
    'PURCHASE ORDER',
    'PURCHASE ORDER NUMBER',
    'PURCHASE ORDER NO',
    'PO NUMBER',
    'P O NUMBER',
    'P.O.',
  ],
  'SUPPLIER': [
    'MANUFACTURER',
    'VENDOR',
    'MAKER',
    'BRAND',
    'HERSTELLER',
    'MANUFACTURER NAME',
    'MAKER NAME',
    'SUPPLIER',
  ],
  'DESTINATION YARD': ['DESTINATION', 'YARD', 'DESTINATION YARD ACTUATOR', 'DESTINATION (YARD)', 'SHIPYARD', 'YARD DESTINATION'],
  'SAP CODE': [
    'SAP',
    'SAP_CODE',
    'SAP CODE',
    'SAP NAME',
    'SAP_NAME',
    'SAPNAME',
    'SAPCODE',
    'SAP-NAME',
    'MATERIAL CODE',
    'MATERIAL NO',
    'MATERIAL NUMBER',
    'MATERIAL-NO',
    'MATERIAL',
    'PART NUMBER',
    'PART NO',
    'VALVE CODE',
    'VALVE NAME',
    'ITEM CODE',
    'ITEM NO',
    'NAME',
    'SAP NO',
    'SAP NUMBER',
  ],
  'STD DRW NORMALE N': [
    'STD DRAWING',
    'STANDARD DRAWING',
    'STANDARD DRW',
    'NORMALE NR',
    'NORMALE NO',
    'NORMALE N',
    'NORMALE N°',
    'NORMALE NR.',
    'NORMALE NO.',
    'NORMALE',
    'NORMALE NUMBER',
    'STD DRW NORMALE',
    'STD DRW NORMALE N°',
    'DRAWING NO',
    'STD DRW',
    'DRAWING NUMBER',
  ],
  'EXECUTION': ['EXEC', 'EXE', 'EXECUTION TYPE', 'EXECUTION'],
  'NRF N': ['NRF', 'NRF NO', 'NRF NR', 'NRF NUMBER', 'NRF N°', 'NRF NR.', 'NRF NO.'],
  'SFI': ['SFI CODE', 'SFI GROUP', 'SYSTEM', 'SFI NO', 'SFI-NR'],
  'DESCRIPTION': ['DESC', 'VALVE DESCRIPTION', 'ITEM DESCRIPTION', 'SHORT TEXT', 'MATERIAL DESCRIPTION', 'ITEM NAME'],
  'SIZE': ['PIPE SIZE', 'DN', 'DIAMETER', 'VALVE SIZE', 'NOMINAL SIZE', 'SIZE', 'DIMENSION'],
  'CONNECTION': ['CONN', 'END CONNECTION', 'CONNECTION TYPE', 'CONNECTION'],
  'PRESSURE RATING': [
    'PRESSURE',
    'PRESSURE CLASS',
    'RATING',
    'PRESSURE NOMINAL',
    'PN',
    'CLASS RATING',
    'NOMINAL PRESSURE',
    'PRESSURE RATING',
  ],
  'HOUSING BODY': [
    'HOUSING',
    'BODY',
    'BODY MATERIAL',
    'MATERIAL BODY',
    'HOUSING /BODY',
    'HOUSING/BODY',
    'BODY MAT',
    'BODY MAT.',
    'HOUSING BODY',
    'MATERIAL',
  ],
  'TYPE': ['MODEL NUMBER', 'MODEL', 'VALVE TYPE', 'CAMLOCK TYPE', 'MODEL NO', 'MODEL NR', 'SIGN TYPE'],
  'PIPE CLASS': ['PIPE CLASSIFICATION', 'CLASS', 'PIPE CLASS', 'PIPING CLASS'],
  'CLASS CERTIFICATE': [
    'CLASS CERT',
    'CERTIFICATE',
    'CLASS CERTIFICATION',
    'TESTING CERTIFICATE',
    'CERT',
    'TEST CERTIFICATE',
    'CERTIFICATE TYPE',
  ],
  'REMARKS': ['COMMENT', 'COMMENTS', 'NOTE', 'NOTES', 'REMARK', 'REMARKS'],
  'SIGN TYPE': ['SIGN_TYPE', 'SIGNTYPE', 'SIGN', 'TAG TYPE', 'LABEL TYPE'],
  'SIGN TEXT': ['SIGN_TEXT', 'SIGNTEXT', 'SIGN CONTENT', 'LABEL TEXT', 'ENGRAVING TEXT', 'PLATE TEXT'],
  'INPUT SIGN TEXT': ['INPUT - SIGN TEXT', 'INPUT SIGN TEXT', 'INPUT_SIGN_TEXT', 'SIGN INPUT', 'SIGN TEXT INPUT'],
  'REV HIS': ['REV. HIS.', 'REV HIS', 'REV. HIST.', 'REV HIST', 'REVISION HISTORY', 'REV', 'REVISION', 'REV NO', 'REV HIST.'],
  'DATE': ['REV DATE', 'REVISION DATE', 'MODIFIED DATE', 'UPDATE DATE', 'ISSUE DATE'],
  'REV DESCRIPTION': ['REV. DESCRIPTION', 'REV DESCRIPTION', 'REV DESC', 'REVISION DESCRIPTION', 'CHANGE DESCRIPTION', 'REV. DESC.'],
  'SIGNATURE': ['SIGN', 'SIGNED BY', 'BY', 'APPROVED BY', 'CHECKED BY', 'SIG', 'SIGN BY', 'SIGN OFF'],
};

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
