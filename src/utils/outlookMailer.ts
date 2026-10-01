import { MasterRowData, ImportLogRow } from '../types';

export interface RowEditChange {
  field: string;
  oldValue: any;
  newValue: any;
}

export interface OutlookEmailContext {
  type: 'row_edited' | 'batch_report' | 'vendor_inquiry';
  rowTag?: string;
  supplier?: string;
  sourceFile?: string;
  changes?: RowEditChange[];
  rowData?: MasterRowData;
  missingFields?: string[];
  summaryStats?: {
    totalRows: number;
    totalFiles: number;
    suppliers: string[];
  };
  masterHeaders?: string[];
  allRows?: MasterRowData[];
  logs?: ImportLogRow[];
}

export interface EmailDraft {
  to: string;
  cc: string;
  subject: string;
  body: string;
}

/**
 * Builds an automated email template based on the action context
 */
export function generateEmailDraft(context: OutlookEmailContext): EmailDraft {
  const timestamp = new Date().toLocaleString('en-US', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  if (context.type === 'row_edited') {
    const tag = context.rowTag || context.rowData?.TAG || 'N/A';
    const supplier = context.supplier || context.rowData?.SUPPLIER || 'Unknown';
    const sourceFile = context.sourceFile || context.rowData?._sourceFile || 'Master Total';
    const excelAttachmentName = `Valve_Update_Sheet_${String(tag).replace(/[^a-zA-Z0-9_-]/g, '_')}.xlsx`;

    const changeLines = (context.changes || []).map(
      (c) => `  • [${c.field}]: from "${c.oldValue ?? ''}" ➔ "${c.newValue ?? ''}"`
    );

    const subject = `[ARMATURE EXCEL ATTACHMENT] Valve Specification Update ${tag} (${supplier}) - Armature Project`;

    const body = `Dear Technical & Project Management Team,

Please be advised that specifications for valve ${tag} have been updated in the Armature Master system:

GENERAL INFORMATION:
- Valve Tag: ${tag}
- Supplier: ${supplier}
- Source File / Drawing: ${sourceFile}
- Update Timestamp: ${timestamp}

MODIFIED FIELDS SUMMARY:
${changeLines.length > 0 ? changeLines.join('\n') : '  • Synchronized and adjusted detailed specification fields.'}

CURRENT TECHNICAL SPECIFICATIONS:
- SFI Code: ${context.rowData?.SFI || 'N/A'}
- Description: ${context.rowData?.DESCRIPTION || 'N/A'}
- Valve Type: ${context.rowData?.['TYPE OF VALVE'] || 'N/A'}
- Size: ${context.rowData?.['SIZE INCH'] || 'N/A'}
- Pressure Rating: ${context.rowData?.['PRESSURE RATING'] || 'N/A'}
- Pipe Class: ${context.rowData?.['PIPE CLASS'] || 'N/A'}
- Actuator Tag: ${context.rowData?.['ACTUATOR TAG'] || 'N/A'}
- PO Number: ${context.rowData?.['PO NUMBER'] || 'N/A'}

📎 ATTACHMENT:
- File Name: ${excelAttachmentName}
  (Includes valve specification sheet, highlighted changes, and current full technical parameters)

Please review and reply if any discrepancies are noted.

Best regards,
Piping & Valve Engineering Team
Armature Consolidator Project`;

    return {
      to: '',
      cc: '',
      subject,
      body,
    };
  }

  if (context.type === 'vendor_inquiry') {
    const supplier = context.supplier || 'Valued Supplier';
    const tag = context.rowTag || context.rowData?.TAG || '';
    const missing = (context.missingFields || []).join(', ') || 'PO NUMBER, ACTUATOR TAG, CLASS CERTIFICATE';
    const safeSupplier = supplier.replace(/[^a-zA-Z0-9_-]/g, '_');

    const subject = `[DATA REQUEST] Clarification of Armature Valve Parameters - Vendor: ${supplier}`;

    const body = `Dear Technical & Sales Team - ${supplier},

During the import and standardization of Armature material specifications for the project, we identified technical fields requiring your confirmation:

- Vendor / Supplier: ${supplier}
${tag ? `- Equipment Tag: ${tag}\n` : ''}- Required / Missing Parameters: ${missing}
- Reference Schedule: ${context.sourceFile || 'Submission Schedule'}

📎 ATTACHMENT:
- File Name: Armature_Verification_List_${safeSupplier}.xlsx

Please confirm and update the above parameters so we can complete material technical datasheets for project execution.

Thank you for your cooperation!

Best regards,
Piping & Valve Materials Engineer
Armature Master Project`;

    return {
      to: '',
      cc: '',
      subject,
      body,
    };
  }

  // batch_report default
  const totalRows = context.summaryStats?.totalRows ?? 0;
  const totalFiles = context.summaryStats?.totalFiles ?? 0;
  const suppliers = context.summaryStats?.suppliers?.join(', ') || 'Multiple Suppliers/Vendors';

  const subject = `[IMPORT REPORT EXCEL] Armature Master Data Summary Report (${totalRows} valves)`;

  const body = `Dear Project Management & Technical Leads,

The Armature Master system presents the summary report for material import and data standardization progress:

OVERVIEW:
- Extraction Timestamp: ${timestamp}
- Total Source Files Imported: ${totalFiles} file(s)
- Total Valve Items in Master: ${totalRows} row(s)
- Recorded Suppliers: ${suppliers}

📎 ATTACHMENT:
- File Name: Armature_Master_Export.xlsx (includes 'Total' Master Sheet and 'Import_Log' Sheet)

The database is ready for review and project handover.

Best regards,
Armature Data Management Team`;

  return {
    to: '',
    cc: '',
    subject,
    body,
  };
}

/**
 * Opens Outlook Desktop application using standard mailto: URI
 */
export function openInOutlookDesktop(draft: EmailDraft) {
  const { to, cc, subject, body } = draft;
  const params: string[] = [];

  if (cc) params.push(`cc=${encodeURIComponent(cc)}`);
  if (subject) params.push(`subject=${encodeURIComponent(subject)}`);
  if (body) params.push(`body=${encodeURIComponent(body)}`);

  const query = params.length > 0 ? `?${params.join('&')}` : '';
  const mailtoUrl = `mailto:${encodeURIComponent(to || '')}${query}`;

  // Safe invocation that doesn't replace the current tab
  const a = document.createElement('a');
  a.href = mailtoUrl;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

/**
 * Opens Outlook Web (Office 365 or Outlook.live) in a new browser tab
 */
export function openInOutlookWeb(draft: EmailDraft, isPersonal = false) {
  const { to, cc, subject, body } = draft;
  const baseUrl = isPersonal
    ? 'https://outlook.live.com/mail/0/deeplink/compose'
    : 'https://outlook.office.com/mail/deeplink/compose';

  const params = new URLSearchParams();
  if (to) params.set('to', to);
  if (cc) params.set('cc', cc);
  if (subject) params.set('subject', subject);
  if (body) params.set('body', body);

  const finalUrl = `${baseUrl}?${params.toString()}`;
  window.open(finalUrl, '_blank', 'noopener,noreferrer');
}
