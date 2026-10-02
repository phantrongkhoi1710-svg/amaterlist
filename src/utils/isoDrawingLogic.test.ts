/**
 * Self-contained verification tests for isoDrawingLogic.ts
 */

import {
  parseIsoFileName,
  computeSystemShort,
  matchTemplateFolder,
  compareRevisions,
  markOldRevisionsInMatrix,
} from './isoDrawingLogic';

export function runIsoDrawingLogicSelfTest(): { pass: boolean; logs: string[] } {
  const logs: string[] = [];
  let allPass = true;

  function assert(condition: boolean, testName: string) {
    if (condition) {
      logs.push(`PASS: ${testName}`);
    } else {
      logs.push(`FAIL: ${testName}`);
      allPass = false;
    }
  }

  // Test 1: Given example parsing
  const ex1 = parseIsoFileName('N1005-212-P-CS-002-59.1-821-001', '01/10/2026');
  assert(ex1.isValid === true, 'Parse valid filename validity');
  assert(ex1.zone === '212', `Zone should be '212', got '${ex1.zone}'`);
  assert(ex1.system === '821-002', `System should be '821-002', got '${ex1.system}'`);
  assert(ex1.pipeSpool === '59.1-001', `Pipe spool should be '59.1-001', got '${ex1.pipeSpool}'`);
  assert(ex1.rev === '0', `Rev should default to '0', got '${ex1.rev}'`);
  assert(ex1.systemShort === '8212', `SystemShort should be '8212', got '${ex1.systemShort}'`);

  // Test 2: File with revision suffix
  const ex2 = parseIsoFileName('N1005-212-P-CS-002-59.1-821-001_B');
  assert(ex2.rev === 'B', `Rev should be 'B', got '${ex2.rev}'`);

  // Test 3: System short code edge cases
  assert(computeSystemShort('821-002') === '8212', 'SystemShort 821-002 -> 8212');
  assert(computeSystemShort('501-010') === '5011', 'SystemShort 501-010 -> 5011');
  assert(computeSystemShort('100-001') === '1001', 'SystemShort 100-001 -> 1001');

  // Test 4: Invalid filename
  const invalid = parseIsoFileName('Short-Name-01');
  assert(invalid.isValid === false, 'Detect invalid short filename');
  assert(Boolean(invalid.errorMessage), 'Provide helpful error message on invalid filename');

  // Test 5: Template folder matching
  const templates = ['S8212 - SEAWATER SYSTEM', 'S5011 - BILGE & BALLAST'];
  assert(
    matchTemplateFolder('8212', templates) === 'S8212 - SEAWATER SYSTEM',
    'Match template folder containing systemShort',
  );
  assert(
    matchTemplateFolder('9999', templates) === '9999',
    'Fallback to systemShort when no template matches',
  );

  // Test 6: Revision comparison
  assert(compareRevisions('0', '1') < 0, 'Rev 0 < 1');
  assert(compareRevisions('2', '10') < 0, 'Rev 2 < 10 (numeric)');
  assert(compareRevisions('B', 'A') > 0, 'Rev B > A');

  // Test 7: Mark old revisions in matrix
  const sampleMatrix: unknown[][] = [
    ['Header1', 'HeaderKey', 'HeaderRev', 'HeaderMark'],
    ['Row1', 'PIPE-01', '0', ''],
    ['Row2', 'PIPE-01', '1', ''],
    ['Row3', 'PIPE-02', '0', ''],
  ];
  const markRes = markOldRevisionsInMatrix(sampleMatrix, 1, 1, 2, 3);
  assert(markRes.markedCount === 1, `Expected 1 marked row, got ${markRes.markedCount}`);
  assert(markRes.updatedMatrix[1][3] === 'old rev', 'Row 1 (rev 0) marked as old rev');
  assert(markRes.updatedMatrix[2][3] === '', 'Row 2 (rev 1) kept as current rev');

  return { pass: allPass, logs };
}
