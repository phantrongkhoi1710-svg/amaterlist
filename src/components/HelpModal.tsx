import React from 'react';
import { X, CheckCircle2, HelpCircle, FileSpreadsheet, Code, ShieldCheck } from 'lucide-react';

interface HelpModalProps {
  isOpen: boolean;
  onClose: () => void;
}

export const HelpModal: React.FC<HelpModalProps> = ({ isOpen, onClose }) => {
  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs">
      <div className="bg-white dark:bg-slate-900 rounded-2xl max-w-2xl w-full border border-slate-200 dark:border-slate-800 shadow-2xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* Modal Header */}
        <div className="px-6 py-4 border-b border-slate-200 dark:border-slate-800 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-100 dark:bg-blue-950/60 text-blue-600 dark:text-blue-400 flex items-center justify-center">
              <Code className="w-4 h-4" />
            </div>
            <h2 className="text-base font-bold text-slate-900 dark:text-slate-100">
              Nguyên lý Hoạt động & Đối chiếu Macro VBA
            </h2>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Content */}
        <div className="p-6 overflow-y-auto space-y-5 text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
          <div className="bg-blue-50 dark:bg-blue-950/30 border border-blue-200 dark:border-blue-900/40 rounded-xl p-4 text-blue-900 dark:text-blue-200">
            <p className="font-semibold text-sm mb-1">
              Ứng dụng Web được chuyển đổi trung thực 100% từ Macro ARMATURE IMPORT TOOL
            </p>
            <p>
              Giúp bạn nhập cùng lúc hàng chục file Excel vật tư/van Armature của nhiều nhà cung cấp khác nhau vào bảng tổng hợp Master mà không cần mở Excel, không lo xung đột phiên bản và hoạt động trực tiếp ngay trên trình duyệt!
            </p>
          </div>

          <div className="space-y-3">
            <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-500" />
              1. Quy tắc nhận diện Dòng Tiêu đề (Header Detection)
            </h3>
            <p>
              Hệ thống tự động quét tối đa 60 dòng đầu tiên của từng sheet trong file nguồn. Một dòng được xác định là dòng tiêu đề khi chứa đồng thời:
            </p>
            <ul className="list-disc list-inside ml-2 space-y-1 font-mono text-slate-800 dark:text-slate-200">
              <li>Cột có tên chứa <strong>TAG</strong></li>
              <li>Cột có tên chứa <strong>SUPPLIER</strong></li>
              <li>Cột có tên chứa <strong>SFI</strong> hoặc <strong>DESCRIPTION</strong></li>
            </ul>
          </div>

          <div className="space-y-3">
            <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-500" />
              2. Chuẩn hoá Tiêu đề & Tra cứu Bí danh (Normalize & Aliases)
            </h3>
            <p>
              Trước khi so khớp, tiêu đề được làm sạch triệt để: loại bỏ ngắt dòng (\r, \n), dấu tab, ký tự non-breaking space (Chr 160), dấu chấm, dấu gạch ngang, gạch chéo, dấu độ (°), chuyển toàn bộ sang chữ IN HOA và gom cụm khoảng trắng thừa.
            </p>
            <p>
              Nếu tên cột không trùng khớp chính xác, hệ thống tự động đối chiếu qua bảng Alias (ví dụ: <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">PO NO</code> &rarr; <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">PO NUMBER</code>, <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">PRESSURE CLASS</code> &rarr; <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">PRESSURE RATING</code>, v.v.).
            </p>
          </div>

          <div className="space-y-3">
            <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-500" />
              3. Lọc dòng dữ liệu hợp lệ (IsDataRow) & Xử lý trùng lặp cột (Occurrence)
            </h3>
            <p>
              Chỉ các dòng có ít nhất một giá trị tại các cột đã ánh xạ mới được nhập vào Master (tự động loại bỏ hoàn toàn dòng trắng hoặc dòng ghi chú).
            </p>
            <p>
              Trường hợp một tiêu đề xuất hiện nhiều lần (ví dụ: <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">DESCRIPTION #1</code>, <code className="bg-slate-100 dark:bg-slate-800 px-1 py-0.5 rounded">DESCRIPTION #2</code>), thuật toán xếp lần lượt theo thứ tự xuất hiện (occurrence matching).
            </p>
          </div>

          <div className="space-y-3">
            <h3 className="font-bold text-slate-900 dark:text-slate-100 text-sm flex items-center gap-2">
              <CheckCircle2 className="w-4 h-4 text-emerald-500" />
              4. Xuất file kết quả chuẩn định dạng
            </h3>
            <p>
              Khi nhấn <strong>&quot;Xuất Excel (.xlsx)&quot;</strong>, file tải về sẽ có cấu trúc gồm:
            </p>
            <ul className="list-disc list-inside ml-2 space-y-1">
              <li><strong>Sheet &quot;Total&quot;</strong>: Chứa toàn bộ dữ liệu Armature đã gộp, dòng 2 là Header chuẩn.</li>
              <li><strong>Sheet &quot;Import_Log&quot;</strong>: Chứa đầy đủ nhật ký ngày giờ, tên file nguồn, sheet nguồn, dòng header, số dòng đã nhập và chi tiết cảnh báo lỗi.</li>
            </ul>
          </div>
        </div>

        {/* Modal Footer */}
        <div className="px-6 py-3.5 border-t border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-950 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-2 rounded-lg text-xs font-semibold bg-blue-600 hover:bg-blue-500 text-white shadow-sm transition-colors"
          >
            Đã hiểu, tiếp tục
          </button>
        </div>
      </div>
    </div>
  );
};
