using Automatic_Storage.Utilities;
using System.Collections.Generic;
using System.Data;

namespace Automatic_Storage.Services
{
    /// <summary>
    /// Excel 服務，負責封裝所有 Excel 相關操作，並委派至 ExcelInteropHelper。
    /// 方便進行依賴注入（DI）及單元測試（Mock）。
    /// </summary>
    public class ExcelService : IExcelService
    {
        /// <summary>
        /// 載入指定 Excel 檔案的第一個工作表，並轉換為 <see cref="DataTable"/>。
        /// </summary>
        /// <param name="path">Excel 檔案路徑。</param>
        /// <returns>第一個工作表的資料表。</returns>
        public DataTable LoadFirstWorksheetToDataTable(string path)
        {
            return ExcelInteropHelper.LoadFirstWorksheetToDataTable(path);
        }

        /// <summary>
        /// 更新 Excel 檔案中指定料號的出貨數量，並於記錄區塊新增一筆操作紀錄。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="materialCode">料號。</param>
        /// <param name="qty">出貨數量。</param>
        /// <param name="operatorName">操作人員名稱。</param>
        public void UpdateShippedAndAppendRecord(string excelPath, string materialCode, int qty, string operatorName)
        {
            ExcelInteropHelper.UpdateShippedAndAppendRecord(excelPath, materialCode, qty, operatorName);
        }

        /// <summary>
        /// 保護 Excel 工作表，並可選擇性針對 xlsm 格式保護出貨欄位。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="password">保護密碼。</param>
        /// <param name="protectShippedColumnForXlsm">是否針對 xlsm 格式保護出貨欄位。</param>
        public void ProtectWorksheet(string excelPath, string password, bool protectShippedColumnForXlsm)
        {
            ExcelInteropHelper.ProtectWorksheet(excelPath, password, protectShippedColumnForXlsm);
        }

        /// <summary>
        /// 解除 Excel 工作表的保護。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="password">保護密碼。</param>
        public void UnprotectWorksheet(string excelPath, string password)
        {
            ExcelInteropHelper.UnprotectWorksheet(excelPath, password);
        }

        /// <summary>
        /// 檢查指定 Excel 工作表是否已被保護。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <returns>若已保護則回傳 true，否則回傳 false。</returns>
        public bool IsWorksheetProtected(string excelPath)
        {
            return ExcelInteropHelper.IsWorksheetProtected(excelPath);
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

