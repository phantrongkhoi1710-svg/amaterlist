import { MasterRowData } from '../types';

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
  const timestamp = new Date().toLocaleString('vi-VN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });

  if (context.type === 'row_edited') {
    const tag = context.rowTag || context.rowData?.TAG || 'N/A';
    const supplier = context.supplier || context.rowData?.SUPPLIER || 'Chưa rõ';
    const sourceFile = context.sourceFile || context.rowData?._sourceFile || 'Master Total';

    const changeLines = (context.changes || []).map(
      (c) => `  • [${c.field}]: từ "${c.oldValue ?? ''}" ➔ "${c.newValue ?? ''}"`
    );

    const subject = `[ARMATURE THAY ĐỔI] Cập nhật thông số van ${tag} (${supplier}) - Dự án Armature`;

    const body = `Kính gửi Quý Đội ngũ Kỹ thuật & Quản lý Dự án,

Chúng tôi xin thông báo vừa thực hiện hiệu chỉnh thông số cho van ${tag} trong hệ thống Armature Master:

THÔNG TIN CHUNG:
- Mã van (TAG): ${tag}
- Nhà cung cấp (SUPPLIER): ${supplier}
- File nguồn / Bản vẽ: ${sourceFile}
- Thời gian cập nhật: ${timestamp}

CHI TIẾT CÁC MỤC THAY ĐỔI:
${changeLines.length > 0 ? changeLines.join('\n') : '  • Đã hiệu chỉnh và đồng bộ các trường dữ liệu chi tiết.'}

THÔNG SỐ KỸ THUẬT HIỆN TẠI:
- SFI Code: ${context.rowData?.SFI || 'N/A'}
- Mô tả: ${context.rowData?.DESCRIPTION || 'N/A'}
- Loại van: ${context.rowData?.['TYPE OF VALVE'] || 'N/A'}
- Kích thước: ${context.rowData?.['SIZE INCH'] || 'N/A'}
- Áp suất thiết kế (Rating): ${context.rowData?.['PRESSURE RATING'] || 'N/A'}
- Pipe Class: ${context.rowData?.['PIPE CLASS'] || 'N/A'}
- Actuator Tag: ${context.rowData?.['ACTUATOR TAG'] || 'N/A'}
- PO Number: ${context.rowData?.['PO NUMBER'] || 'N/A'}

Vui lòng rà soát và phản hồi nếu có bất kỳ sai lệch nào.

Trân trọng,
Bộ phận Kỹ thuật Cơ điện / Piping
Dự án Armature Consolidator`;

    return {
      to: '',
      cc: '',
      subject,
      body,
    };
  }

  if (context.type === 'vendor_inquiry') {
    const supplier = context.supplier || 'Quý Nhà cung cấp';
    const tag = context.rowTag || context.rowData?.TAG || '';
    const missing = (context.missingFields || []).join(', ') || 'PO NUMBER, ACTUATOR TAG, CLASS CERTIFICATE';

    const subject = `[YÊU CẦU BỔ SUNG DỮ LIỆU] Làm rõ thông số van Armature - NCC: ${supplier}`;

    const body = `Kính gửi Bộ phận Kỹ thuật & Bán hàng - ${supplier},

Trong quá trình import và chuẩn hoá cơ sở dữ liệu vật tư Armature cho dự án, chúng tôi ghi nhận một số thông tin kỹ thuật cần Quý công ty xác nhận và bổ sung:

- Nhà cung cấp: ${supplier}
${tag ? `- Mã thiết bị (TAG): ${tag}\n` : ''}- Các trường thông tin cần làm rõ/còn thiếu: ${missing}
- File bảng kê đối chiếu: ${context.sourceFile || 'Bảng kê đệ trình'}

Kính đề nghị Quý công ty phản hồi cập nhật các thông số trên để chúng tôi hoàn thiện hồ sơ vật tư và tiến hành các bước tiếp theo của dự án.

Xin chân thành cảm ơn sự phối hợp của Quý công ty!

Trân trọng,
Kỹ sư Phụ trách Vật tư Piping & Valve
Dự án Armature Master`;

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
  const suppliers = context.summaryStats?.suppliers?.join(', ') || 'Nhiều nhà thầu/NCC';

  const subject = `[BÁO CÁO NHẬP LIỆU] Tổng hợp dữ liệu Armature Master (${totalRows} van)`;

  const body = `Kính gửi Ban Quản lý Dự án & Trưởng bộ phận,

Hệ thống Armature Master xin gửi báo cáo tổng hợp tiến độ import và chuẩn hoá dữ liệu vật tư van:

TỔNG QUAN HỒ SƠ:
- Thời gian trích xuất: ${timestamp}
- Tổng số file nguồn đã import: ${totalFiles} file
- Tổng số lượng thiết bị van trong Master: ${totalRows} dòng
- Danh sách nhà cung cấp đã ghi nhận: ${suppliers}

Tình trạng cơ sở dữ liệu đã sẵn sàng để xuất file Excel chuẩn (Total + Import_Log) phục vụ công tác kiểm tra và bàn giao.

Trân trọng,
Bộ phận Quản lý Dữ liệu Armature`;

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
