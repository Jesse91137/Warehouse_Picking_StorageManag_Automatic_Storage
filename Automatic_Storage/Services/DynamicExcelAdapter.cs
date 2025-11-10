using Automatic_Storage.Utilities;
using System;
using System.Collections.Generic;
using System.Data;

namespace Automatic_Storage.Services
{
    /// <summary>
    /// 動態 Excel 服務的介接器，包裝現有的 dynamic excel 服務，
    /// 主要用於表單，會優先呼叫 dynamic 方法，若無則回退至 ExcelInteropHelper 的安全預設實作。
    /// </summary>
    public class DynamicExcelAdapter : IExcelService
    {
        /// <summary>
        /// dynamic 型別的 Excel 服務實例。
        /// </summary>
        private readonly dynamic _dyn;

        /// <summary>
        /// 建立 DynamicExcelAdapter 實例。
        /// </summary>
        /// <param name="dyn">dynamic 型別的 Excel 服務物件，不能為 null。</param>
        /// <exception cref="ArgumentNullException">dyn 為 null 時拋出。</exception>
        public DynamicExcelAdapter(dynamic dyn)
        {
            _dyn = dyn ?? throw new ArgumentNullException(nameof(dyn));
        }

        /// <summary>
        /// 載入 Excel 檔案的第一個工作表為 DataTable。
        /// </summary>
        /// <param name="path">Excel 檔案路徑。</param>
        /// <returns>第一個工作表的資料內容。</returns>
        public DataTable LoadFirstWorksheetToDataTable(string path)
        {
            try
            {
                var res = _dyn.LoadFirstWorksheetToDataTable(path);
                return res as DataTable ?? ExcelInteropHelper.LoadFirstWorksheetToDataTable(path);
            }
            catch { return ExcelInteropHelper.LoadFirstWorksheetToDataTable(path); }
        }

        /// <summary>
        /// 更新已出貨數量並新增一筆記錄到 Excel。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="materialCode">料號。</param>
        /// <param name="qty">數量。</param>
        /// <param name="operatorName">操作人員名稱。</param>
        public void UpdateShippedAndAppendRecord(string excelPath, string materialCode, int qty, string operatorName)
        {
            try { _dyn.UpdateShippedAndAppendRecord(excelPath, materialCode, qty, operatorName); }
            catch { ExcelInteropHelper.UpdateShippedAndAppendRecord(excelPath, materialCode, qty, operatorName); }
        }

        /// <summary>
        /// 保護 Excel 工作表，並可選擇性保護已出貨欄位（僅 xlsm）。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="password">保護密碼。</param>
        /// <param name="protectShippedColumnForXlsm">是否保護已出貨欄位（僅 xlsm 有效）。</param>
        public void ProtectWorksheet(string excelPath, string password, bool protectShippedColumnForXlsm)
        {
            try { _dyn.ProtectWorksheet(excelPath, password, protectShippedColumnForXlsm); }
            catch { ExcelInteropHelper.ProtectWorksheet(excelPath, password, protectShippedColumnForXlsm); }
        }

        /// <summary>
        /// 解除 Excel 工作表保護。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="password">保護密碼。</param>
        public void UnprotectWorksheet(string excelPath, string password)
        {
            try { _dyn.UnprotectWorksheet(excelPath, password); }
            catch { ExcelInteropHelper.UnprotectWorksheet(excelPath, password); }
        }

        /// <summary>
        /// 檢查 Excel 工作表是否已被保護。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <returns>若已保護則回傳 true，否則回傳 false。</returns>
        public bool IsWorksheetProtected(string excelPath)
        {
            try
            {
                var res = _dyn.IsWorksheetProtected(excelPath);
                if (res is bool) return (bool)res;
            }
            catch { }
            // fallback
            try { return ExcelInteropHelper.IsWorksheetProtected(excelPath); } catch { return false; }
        }

        /// <summary>
        /// 將記錄批次追加寫入 Excel 的「記錄」工作表。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="records">待寫入的記錄清單。</param>
        /// <param name="password">保護密碼。</param>
        /// <returns>成功寫入的記錄清單。</returns>
        public List<Dto.記錄Dto> AppendRecordsToLogSheet(string excelPath, List<Dto.記錄Dto> records, string password)
        {
            // TODO: 可直接搬移原 Interop 實作，或呼叫 Utilities/Helper
            return ExcelInteropHelper.AppendRecordsToLogSheet(excelPath, records, password);
        }
    }
}
