import React, { useState } from 'react';
import { 
  DownloadCloud, 
  X, 
  CheckCircle2, 
  AlertCircle, 
  RefreshCw, 
  FileSpreadsheet, 
  Link as LinkIcon, 
  Clipboard, 
  Layers, 
  Database, 
  FileText, 
  Download, 
  Upload, 
  Check, 
  HelpCircle,
  Clock,
  Sparkles,
  ExternalLink,
  ShieldCheck
} from 'lucide-react';
import { GoogleSheetsConfig, ViolationRecord, BonusRecord, ClassInfo } from '../types';
import { 
  extractSheetIdFromUrl, 
  scanGoogleSheetConnection, 
  exportViolationsToCSV, 
  exportSummaryToCSV, 
  downloadFile,
  parseCSVText,
  mapRawToViolations,
  mapRawToClasses,
  mapRawToBonuses
} from '../utils/googleSheetsHelper';

interface GoogleSheetsImportModalProps {
  isOpen: boolean;
  onClose: () => void;
  config: GoogleSheetsConfig;
  onUpdateConfig: (config: GoogleSheetsConfig) => void;
  onPullAll: (mode: 'MERGE' | 'REPLACE') => Promise<void>;
  onPullTable: (table: 'violations' | 'classes' | 'bonuses' | 'rules', mode: 'MERGE' | 'REPLACE') => Promise<void>;
  onDirectImportData?: (data: { violations?: ViolationRecord[]; bonuses?: BonusRecord[]; classes?: ClassInfo[] }, mode: 'MERGE' | 'REPLACE') => void;
  isPulling: boolean;
  totalViolations: number;
  totalBonuses: number;
  totalClasses: number;
  weekNumber: number;
}

export const GoogleSheetsImportModal: React.FC<GoogleSheetsImportModalProps> = ({
  isOpen,
  onClose,
  config,
  onUpdateConfig,
  onPullAll,
  onPullTable,
  onDirectImportData,
  isPulling,
  totalViolations,
  totalBonuses,
  totalClasses,
  weekNumber
}) => {
  if (!isOpen) return null;

  const [activeTab, setActiveTab] = useState<'online' | 'guide' | 'csv'>('online');
  const [sheetInput, setSheetInput] = useState<string>(config.sheetId || config.webAppUrl || '');
  const [syncMode, setSyncMode] = useState<'MERGE' | 'REPLACE'>('MERGE');
  const [isScanning, setIsScanning] = useState<boolean>(false);
  const [scanResult, setScanResult] = useState<{
    status: 'idle' | 'success' | 'error';
    message: string;
    counts?: { violations: number; bonuses: number; classes: number };
  }>({
    status: config.isConnected ? 'success' : 'idle',
    message: config.isConnected 
      ? 'Hệ thống đã liên kết với Google Sheets. Sẵn sàng nhận dữ liệu bất cứ lúc nào.' 
      : ''
  });

  const [csvUploadResult, setCsvUploadResult] = useState<string | null>(null);

  // Dán nhanh link từ Clipboard
  const handlePasteClipboard = async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (text && text.trim()) {
        setSheetInput(text.trim());
      }
    } catch {
      // Trình duyệt có thể chặn quyền clipboard, người dùng có thể dán thủ công
    }
  };

  // Kiểm tra kết nối & Quét dữ liệu trực tuyến
  const handleScanConnection = async () => {
    const trimmed = sheetInput.trim();
    if (!trimmed) {
      setScanResult({
        status: 'error',
        message: 'Vui lòng nhập đường link Google Sheet (URL) hoặc Web App URL!'
      });
      return;
    }

    setIsScanning(true);
    setScanResult({
      status: 'idle',
      message: 'Đang kiểm tra kết nối tới Google Sheets...'
    });

    const isWebApp = trimmed.startsWith('https://script.google.com');
    const updatedConfig: GoogleSheetsConfig = {
      ...config,
      sheetId: isWebApp ? config.sheetId : extractSheetIdFromUrl(trimmed),
      webAppUrl: isWebApp ? trimmed : config.webAppUrl,
      isConnected: true,
    };

    try {
      const scan = await scanGoogleSheetConnection(updatedConfig);
      if (scan.connected) {
        onUpdateConfig({
          ...updatedConfig,
          isConnected: true,
          lastPulledAt: new Date().toLocaleTimeString('vi-VN'),
        });
        setScanResult({
          status: 'success',
          message: scan.message,
          counts: scan.counts
        });
      } else {
        setScanResult({
          status: 'error',
          message: scan.message
        });
      }
    } catch (err: any) {
      setScanResult({
        status: 'error',
        message: `Lỗi kết nối: ${err?.message || 'Không thể truy cập Google Sheet. Đảm bảo đã mở quyền xem công khai.'}`
      });
    } finally {
      setIsScanning(false);
    }
  };

  // Thực hiện nhận toàn bộ dữ liệu
  const handleExecutePullAll = async () => {
    // Lưu cấu hình link trước nếu người dùng có thay đổi
    const trimmed = sheetInput.trim();
    if (trimmed) {
      const isWebApp = trimmed.startsWith('https://script.google.com');
      const updatedConfig: GoogleSheetsConfig = {
        ...config,
        sheetId: isWebApp ? config.sheetId : extractSheetIdFromUrl(trimmed),
        webAppUrl: isWebApp ? trimmed : config.webAppUrl,
        isConnected: true
      };
      onUpdateConfig(updatedConfig);
    }

    await onPullAll(syncMode);
  };

  // Xử lý nạp từ file CSV tải về
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const content = event.target?.result as string;
      if (!content) return;

      try {
        const rawRows = parseCSVText(content);
        if (rawRows.length === 0) {
          setCsvUploadResult('File CSV rỗng hoặc định dạng không đúng tiêu chuẩn.');
          return;
        }

        // Tự động nhận diện loại bảng
        const violations = mapRawToViolations(rawRows);
        const classes = mapRawToClasses(rawRows);
        const bonuses = mapRawToBonuses(rawRows);

        if (violations.length > 0) {
          if (onDirectImportData) {
            onDirectImportData({ violations }, syncMode);
          }
          setCsvUploadResult(`✓ Đã nạp thành công ${violations.length} bản ghi vi phạm từ file ${file.name}!`);
        } else if (bonuses.length > 0) {
          if (onDirectImportData) {
            onDirectImportData({ bonuses }, syncMode);
          }
          setCsvUploadResult(`✓ Đã nạp thành công ${bonuses.length} điểm thưởng từ file ${file.name}!`);
        } else if (classes.length > 0) {
          if (onDirectImportData) {
            onDirectImportData({ classes }, syncMode);
          }
          setCsvUploadResult(`✓ Đã nạp thành công ${classes.length} lớp học từ file ${file.name}!`);
        } else {
          setCsvUploadResult('Không nhận diện được cột dữ liệu phù hợp. Vui lòng tải file mẫu bên tab Hướng Dẫn.');
        }
      } catch (err: any) {
        setCsvUploadResult(`Lỗi đọc file CSV: ${err?.message}`);
      }
    };
    reader.readAsText(file, 'UTF-8');
  };

  return (
    <div className="fixed inset-0 z-50 overflow-y-auto bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-3 sm:p-4 animate-in fade-in duration-200">
      <div 
        className="relative bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-3xl shadow-2xl max-w-3xl w-full overflow-hidden flex flex-col max-h-[92vh]"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header Bar */}
        <div className="bg-gradient-to-r from-emerald-700 via-teal-700 to-slate-900 text-white p-5 sm:p-6 flex items-start justify-between relative overflow-hidden">
          <div className="absolute right-0 top-0 translate-x-8 -translate-y-8 w-44 h-44 bg-emerald-400/10 rounded-full blur-2xl pointer-events-none" />
          
          <div className="flex items-center space-x-3.5 z-10">
            <div className="w-12 h-12 rounded-2xl bg-white/10 border border-white/20 backdrop-blur flex items-center justify-center text-white shadow-lg shadow-emerald-950/40">
              <DownloadCloud className="w-6 h-6 text-emerald-300" />
            </div>
            <div>
              <div className="flex items-center space-x-2">
                <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-200 border border-emerald-400/30">
                  Google Sheets Database
                </span>
                <span className="text-xs text-emerald-200/80">
                  Liên Đội TH&THCS Phước Hiệp
                </span>
              </div>
              <h2 className="text-lg sm:text-xl font-bold tracking-tight text-white mt-0.5">
                Trung Tâm Nhận Dữ Liệu Online Từ Google Sheets
              </h2>
              <p className="text-xs text-emerald-100/80 mt-0.5">
                Kéo toàn bộ vi phạm, điểm thưởng, sĩ số và danh sách 31 lớp về máy 1-Click
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            className="text-slate-300 hover:text-white p-1.5 rounded-xl hover:bg-white/10 transition-colors z-10"
            aria-label="Đóng cửa sổ"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Tab Navigation */}
        <div className="flex items-center space-x-1 border-b border-slate-200 dark:border-slate-800 bg-slate-50 dark:bg-slate-900/60 px-5 pt-2">
          <button
            onClick={() => setActiveTab('online')}
            className={`px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all flex items-center space-x-2 border-b-2 ${
              activeTab === 'online'
                ? 'border-emerald-500 text-emerald-600 dark:text-emerald-400 bg-white dark:bg-slate-900 shadow-sm'
                : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
            }`}
          >
            <FileSpreadsheet className="w-4 h-4" />
            <span>Nhận Online Trực Tiếp</span>
          </button>

          <button
            onClick={() => setActiveTab('guide')}
            className={`px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all flex items-center space-x-2 border-b-2 ${
              activeTab === 'guide'
                ? 'border-emerald-500 text-emerald-600 dark:text-emerald-400 bg-white dark:bg-slate-900 shadow-sm'
                : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
            }`}
          >
            <HelpCircle className="w-4 h-4" />
            <span>Hướng Dẫn 3 Bước & Bảng Mẫu</span>
          </button>

          <button
            onClick={() => setActiveTab('csv')}
            className={`px-4 py-2.5 text-xs font-bold rounded-t-xl transition-all flex items-center space-x-2 border-b-2 ${
              activeTab === 'csv'
                ? 'border-emerald-500 text-emerald-600 dark:text-emerald-400 bg-white dark:bg-slate-900 shadow-sm'
                : 'border-transparent text-slate-500 hover:text-slate-800 dark:hover:text-slate-200'
            }`}
          >
            <Upload className="w-4 h-4" />
            <span>Nhập Từ Tệp CSV</span>
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-5 sm:p-6 overflow-y-auto flex-1 space-y-5">
          {/* TAB 1: NHẬN ONLINE TRỰC TIẾP */}
          {activeTab === 'online' && (
            <div className="space-y-5">
              {/* Input Link Frame */}
              <div className="bg-slate-50 dark:bg-slate-800/60 p-4 sm:p-5 rounded-2xl border border-slate-200 dark:border-slate-700/80 space-y-3">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-1">
                  <label className="text-xs font-bold text-slate-800 dark:text-slate-200 flex items-center gap-1.5">
                    <LinkIcon className="w-3.5 h-3.5 text-emerald-500" />
                    <span>Đường Link Google Sheet Hoặc Web App URL:</span>
                  </label>
                  <span className="text-[11px] text-emerald-600 dark:text-emerald-400 font-semibold">
                    Hỗ trợ dán trực tiếp link chia sẻ bất kỳ
                  </span>
                </div>

                <div className="flex flex-col sm:flex-row items-center gap-2">
                  <div className="relative flex-1 w-full">
                    <input
                      type="text"
                      placeholder="Dán link bảng tính: https://docs.google.com/spreadsheets/d/.../edit"
                      value={sheetInput}
                      onChange={(e) => setSheetInput(e.target.value)}
                      className="w-full text-xs font-mono px-3.5 py-2.5 rounded-xl bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 text-slate-900 dark:text-white focus:outline-none focus:ring-2 focus:ring-emerald-500 pr-10 shadow-sm"
                    />
                    <button
                      type="button"
                      onClick={handlePasteClipboard}
                      title="Dán từ Clipboard"
                      className="absolute right-2 top-2 p-1 text-slate-400 hover:text-emerald-600 transition-colors"
                    >
                      <Clipboard className="w-4 h-4" />
                    </button>
                  </div>

                  <button
                    type="button"
                    onClick={handleScanConnection}
                    disabled={isScanning}
                    className="w-full sm:w-auto px-4 py-2.5 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-100 font-bold text-xs flex items-center justify-center space-x-1.5 transition-all shrink-0 border border-slate-700 shadow-sm"
                  >
                    <RefreshCw className={`w-3.5 h-3.5 ${isScanning ? 'animate-spin text-emerald-400' : ''}`} />
                    <span>{isScanning ? 'Đang quét...' : 'Kiểm Tra Kết Nối'}</span>
                  </button>
                </div>

                {/* Scan Status Box */}
                {scanResult.message && (
                  <div
                    className={`p-3 rounded-xl text-xs font-medium flex items-start space-x-2 transition-all ${
                      scanResult.status === 'success'
                        ? 'bg-emerald-50 dark:bg-emerald-950/40 text-emerald-800 dark:text-emerald-300 border border-emerald-200 dark:border-emerald-800'
                        : scanResult.status === 'error'
                        ? 'bg-rose-50 dark:bg-rose-950/40 text-rose-800 dark:text-rose-300 border border-rose-200 dark:border-rose-800'
                        : 'bg-blue-50 dark:bg-blue-950/40 text-blue-800 dark:text-blue-300 border border-blue-200 dark:border-blue-800'
                    }`}
                  >
                    {scanResult.status === 'success' ? (
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0 mt-0.5" />
                    ) : scanResult.status === 'error' ? (
                      <AlertCircle className="w-4 h-4 text-rose-600 dark:text-rose-400 shrink-0 mt-0.5" />
                    ) : (
                      <RefreshCw className="w-4 h-4 text-blue-600 dark:text-blue-400 shrink-0 mt-0.5 animate-spin" />
                    )}
                    <div className="flex-1">
                      <p>{scanResult.message}</p>
                      {scanResult.counts && (
                        <div className="flex flex-wrap gap-2 mt-2 pt-2 border-t border-emerald-200/50 dark:border-emerald-800/50">
                          <span className="bg-white/80 dark:bg-slate-900 px-2 py-0.5 rounded text-[11px] font-bold text-rose-600">
                            {scanResult.counts.violations} Vi phạm
                          </span>
                          <span className="bg-white/80 dark:bg-slate-900 px-2 py-0.5 rounded text-[11px] font-bold text-amber-600">
                            {scanResult.counts.bonuses} Điểm thưởng
                          </span>
                          <span className="bg-white/80 dark:bg-slate-900 px-2 py-0.5 rounded text-[11px] font-bold text-blue-600">
                            {scanResult.counts.classes} Lớp học
                          </span>
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>

              {/* Mode Selector */}
              <div className="flex flex-col sm:flex-row sm:items-center justify-between p-3.5 rounded-2xl bg-white dark:bg-slate-800/40 border border-slate-200 dark:border-slate-800 gap-3">
                <div>
                  <span className="text-xs font-bold text-slate-800 dark:text-slate-200 block">
                    Chế độ nạp dữ liệu:
                  </span>
                  <span className="text-[11px] text-slate-500">
                    {syncMode === 'MERGE' 
                      ? '✓ Khuyên dùng: Giữ lại dữ liệu hiện có trong ứng dụng và bổ sung thêm bản ghi từ Sheets' 
                      : '⚠ Xóa toàn bộ dữ liệu hiện có và thay thế 100% bằng dữ liệu mới từ Sheets'}
                  </span>
                </div>

                <div className="flex items-center space-x-1 bg-slate-100 dark:bg-slate-800 p-1 rounded-xl shrink-0">
                  <button
                    type="button"
                    onClick={() => setSyncMode('MERGE')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                      syncMode === 'MERGE'
                        ? 'bg-emerald-600 text-white shadow-sm'
                        : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                    }`}
                  >
                    Hợp nhất (Merge)
                  </button>
                  <button
                    type="button"
                    onClick={() => setSyncMode('REPLACE')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                      syncMode === 'REPLACE'
                        ? 'bg-rose-600 text-white shadow-sm'
                        : 'text-slate-600 dark:text-slate-400 hover:text-slate-900'
                    }`}
                  >
                    Ghi đè (Replace)
                  </button>
                </div>
              </div>

              {/* MAIN 1-CLICK ACTION CARD */}
              <div className="p-5 rounded-2xl bg-gradient-to-r from-emerald-600 via-teal-600 to-emerald-700 text-white shadow-lg shadow-emerald-700/20 relative overflow-hidden">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                  <div className="space-y-1">
                    <span className="text-[10px] font-bold uppercase tracking-wider px-2 py-0.5 rounded-full bg-white/20 text-emerald-100">
                      Thao tác khuyên dùng
                    </span>
                    <h3 className="text-base sm:text-lg font-black tracking-tight">
                      Nhận Toàn Bộ Cơ Sở Dữ Liệu
                    </h3>
                    <p className="text-xs text-emerald-100/90 max-w-md leading-relaxed">
                      Tự động tải về và nạp cả <strong>Bảng Vi Phạm</strong>, <strong>Bảng Điểm Thưởng</strong>, <strong>Danh Sách 31 Lớp</strong> và <strong>Quy Chế Thi Đua</strong>.
                    </p>
                  </div>

                  <button
                    type="button"
                    onClick={handleExecutePullAll}
                    disabled={isPulling}
                    className="px-5 py-3.5 rounded-xl bg-white text-emerald-800 hover:bg-emerald-50 font-black text-xs sm:text-sm shadow-xl flex items-center justify-center space-x-2 shrink-0 transition-transform active:scale-95 disabled:opacity-75"
                  >
                    <DownloadCloud className={`w-5 h-5 text-emerald-600 ${isPulling ? 'animate-bounce' : ''}`} />
                    <span>{isPulling ? 'Đang Nhận Dữ Liệu...' : 'NHẬN DATABASE NGAY'}</span>
                  </button>
                </div>
              </div>

              {/* Individual Table Options */}
              <div className="space-y-2">
                <span className="text-xs font-bold text-slate-700 dark:text-slate-300">
                  Hoặc chọn nhận riêng từng bảng dữ liệu:
                </span>
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
                  <button
                    type="button"
                    onClick={() => onPullTable('violations', syncMode)}
                    disabled={isPulling}
                    className="p-3 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800/80 hover:border-rose-400 dark:hover:border-rose-500 text-left transition-all group flex flex-col justify-between"
                  >
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-slate-800 dark:text-slate-200 group-hover:text-rose-500">
                        1. Sổ Vi Phạm
                      </span>
                      <Download className="w-3.5 h-3.5 text-rose-500 opacity-60 group-hover:opacity-100" />
                    </div>
                    <span className="text-[10px] text-slate-400 mt-1 block">
                      Sheet: Bang_Vi_Pham
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => onPullTable('bonuses', syncMode)}
                    disabled={isPulling}
                    className="p-3 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800/80 hover:border-amber-400 dark:hover:border-amber-500 text-left transition-all group flex flex-col justify-between"
                  >
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-slate-800 dark:text-slate-200 group-hover:text-amber-500">
                        2. Điểm Thưởng
                      </span>
                      <Download className="w-3.5 h-3.5 text-amber-500 opacity-60 group-hover:opacity-100" />
                    </div>
                    <span className="text-[10px] text-slate-400 mt-1 block">
                      Sheet: Bang_Diem_Thuong
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => onPullTable('classes', syncMode)}
                    disabled={isPulling}
                    className="p-3 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800/80 hover:border-blue-400 dark:hover:border-blue-500 text-left transition-all group flex flex-col justify-between"
                  >
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-slate-800 dark:text-slate-200 group-hover:text-blue-500">
                        3. Danh Sách Lớp
                      </span>
                      <Download className="w-3.5 h-3.5 text-blue-500 opacity-60 group-hover:opacity-100" />
                    </div>
                    <span className="text-[10px] text-slate-400 mt-1 block">
                      Sheet: Danh_Sach_Lop (31 Lớp)
                    </span>
                  </button>

                  <button
                    type="button"
                    onClick={() => onPullTable('rules', syncMode)}
                    disabled={isPulling}
                    className="p-3 rounded-xl border border-slate-200 dark:border-slate-800 bg-white dark:bg-slate-800/80 hover:border-indigo-400 dark:hover:border-indigo-500 text-left transition-all group flex flex-col justify-between"
                  >
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-bold text-slate-800 dark:text-slate-200 group-hover:text-indigo-500">
                        4. Quy Chế Điểm
                      </span>
                      <Download className="w-3.5 h-3.5 text-indigo-500 opacity-60 group-hover:opacity-100" />
                    </div>
                    <span className="text-[10px] text-slate-400 mt-1 block">
                      Sheet: Quy_Dinh_Vi_Pham
                    </span>
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: HƯỚNG DẪN 3 BƯỚC & TẢI MẪU */}
          {activeTab === 'guide' && (
            <div className="space-y-4">
              <div className="bg-slate-50 dark:bg-slate-800/50 p-4 rounded-2xl border border-slate-200 dark:border-slate-700/60 space-y-3">
                <h4 className="text-xs font-black uppercase text-slate-800 dark:text-slate-200 tracking-wider">
                  3 Bước Đơn Giản Nhất Để Nhận Dữ Liệu Từ Google Sheets:
                </h4>
                
                <div className="space-y-3 text-xs text-slate-600 dark:text-slate-300">
                  <div className="flex items-start space-x-3">
                    <span className="w-6 h-6 rounded-full bg-emerald-100 dark:bg-emerald-950 text-emerald-700 dark:text-emerald-300 font-bold flex items-center justify-center shrink-0">
                      1
                    </span>
                    <div>
                      <strong>Mở bảng tính Google Sheet của trường</strong> trên trình duyệt Google Drive.
                    </div>
                  </div>

                  <div className="flex items-start space-x-3">
                    <span className="w-6 h-6 rounded-full bg-emerald-100 dark:bg-emerald-950 text-emerald-700 dark:text-emerald-300 font-bold flex items-center justify-center shrink-0">
                      2
                    </span>
                    <div>
                      Bấm nút <strong>Chia sẻ (Share)</strong> ở góc trên bên phải &rarr; tại mục <em>"Quyền truy cập chung"</em>, chọn <strong>"Bất kỳ ai có đường liên kết đều có thể xem" (Viewer)</strong> &rarr; bấm <strong>Sao chép đường liên kết</strong>.
                    </div>
                  </div>

                  <div className="flex items-start space-x-3">
                    <span className="w-6 h-6 rounded-full bg-emerald-100 dark:bg-emerald-950 text-emerald-700 dark:text-emerald-300 font-bold flex items-center justify-center shrink-0">
                      3
                    </span>
                    <div>
                      Dán link vào ô ở tab <strong>"Nhận Online Trực Tiếp"</strong> và bấm nút <strong>"NHẬN DATABASE NGAY"</strong>. Ứng dụng sẽ tự động phân tích và nạp điểm!
                    </div>
                  </div>
                </div>
              </div>

              {/* Download CSV Templates */}
              <div className="p-4 rounded-2xl border border-slate-200 dark:border-slate-800 space-y-3">
                <div className="flex items-center justify-between">
                  <h4 className="text-xs font-bold text-slate-800 dark:text-slate-200">
                    Tải File Mẫu Chuẩn CSV Của Liên Đội TH&THCS Phước Hiệp:
                  </h4>
                  <span className="text-[10px] text-slate-400">Đã định dạng sẵn cột</span>
                </div>

                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      const csv = exportViolationsToCSV([], weekNumber);
                      downloadFile(csv, 'Mau_Bang_Vi_Pham_PhuocHiep.csv');
                    }}
                    className="p-2.5 rounded-xl border border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-800 text-xs font-semibold text-slate-700 dark:text-slate-300 flex items-center justify-center space-x-1.5 transition-colors"
                  >
                    <Download className="w-3.5 h-3.5 text-rose-500" />
                    <span>Mẫu Bang_Vi_Pham.csv</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      const headers = 'id,weekNumber,classId,date,ruleCode,title,pointsAdded,reason,approvedBy\n';
                      downloadFile(headers, 'Mau_Bang_Diem_Thuong_PhuocHiep.csv');
                    }}
                    className="p-2.5 rounded-xl border border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-800 text-xs font-semibold text-slate-700 dark:text-slate-300 flex items-center justify-center space-x-1.5 transition-colors"
                  >
                    <Download className="w-3.5 h-3.5 text-amber-500" />
                    <span>Mẫu Bang_Diem_Thuong.csv</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      const headers = 'id,name,gradeLevel,gradeNumber,teacherName,teacherEmail,studentCount,redStarStudent,roomNumber\n';
                      downloadFile(headers, 'Mau_Danh_Sach_Lop_PhuocHiep.csv');
                    }}
                    className="p-2.5 rounded-xl border border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-800 text-xs font-semibold text-slate-700 dark:text-slate-300 flex items-center justify-center space-x-1.5 transition-colors"
                  >
                    <Download className="w-3.5 h-3.5 text-blue-500" />
                    <span>Mẫu Danh_Sach_Lop.csv</span>
                  </button>
                </div>
              </div>
            </div>
          )}

          {/* TAB 3: NHẬP TỪ TỆP CSV */}
          {activeTab === 'csv' && (
            <div className="space-y-4">
              <div className="border-2 border-dashed border-slate-300 dark:border-slate-700 hover:border-emerald-500 p-8 rounded-2xl text-center space-y-3 transition-colors bg-slate-50/50 dark:bg-slate-800/30">
                <div className="w-12 h-12 rounded-2xl bg-emerald-100 dark:bg-emerald-950/60 text-emerald-600 dark:text-emerald-400 mx-auto flex items-center justify-center">
                  <Upload className="w-6 h-6" />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-slate-800 dark:text-slate-200">
                    Chọn Hoặc Kéo Thả File CSV Từ Máy Tính
                  </h4>
                  <p className="text-xs text-slate-500 mt-1 max-w-sm mx-auto">
                    Áp dụng khi bạn tải file bảng tính từ Google Sheets về máy dưới dạng .csv
                  </p>
                </div>

                <div>
                  <label className="inline-flex items-center space-x-2 px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs cursor-pointer shadow transition-colors">
                    <FileText className="w-4 h-4" />
                    <span>Chọn File CSV Để Nạp</span>
                    <input
                      type="file"
                      accept=".csv,text/csv"
                      onChange={handleFileUpload}
                      className="hidden"
                    />
                  </label>
                </div>
              </div>

              {csvUploadResult && (
                <div className="p-3.5 rounded-xl bg-slate-100 dark:bg-slate-800 text-xs font-medium text-slate-700 dark:text-slate-200 border border-slate-300 dark:border-slate-700 flex items-center space-x-2">
                  <ShieldCheck className="w-4 h-4 text-emerald-500 shrink-0" />
                  <span>{csvUploadResult}</span>
                </div>
              )}
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="bg-slate-50 dark:bg-slate-900 border-t border-slate-200 dark:border-slate-800 px-5 py-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs">
          <div className="flex items-center space-x-2 text-slate-500">
            <Clock className="w-3.5 h-3.5 text-amber-500" />
            <span>
              {config.lastPulledAt 
                ? `Nhận lần gần nhất: ${config.lastPulledAt}` 
                : 'Chưa thực hiện nhận dữ liệu'}
            </span>
          </div>

          <div className="flex items-center space-x-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 rounded-xl border border-slate-300 dark:border-slate-700 text-slate-700 dark:text-slate-200 font-semibold hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
            >
              Đóng
            </button>
            <button
              type="button"
              onClick={handleExecutePullAll}
              disabled={isPulling}
              className="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold flex items-center space-x-1.5 shadow transition-colors"
            >
              <DownloadCloud className="w-3.5 h-3.5" />
              <span>{isPulling ? 'Đang Nhận...' : 'Nhận Dữ Liệu'}</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
