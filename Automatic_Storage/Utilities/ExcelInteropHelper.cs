using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Configuration;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
// NPOI for fast .xls/.xlsx reading (avoids COM)
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using static Automatic_Storage.Utilities.ComInterop;

namespace Automatic_Storage.Utilities
{
    /// <summary>
    /// Excel Interop helper: 把常見的讀取/寫入封裝成可呼叫的同步方法，
    /// 以便外層使用 Task.Run 執行，並確保 COM 物件釋放與例外處理一致性。
    /// 注意：呼叫端應在背景執行緒使用，以避免 UI 卡頓。
    /// </summary>
    public static class ExcelInteropHelper
    {

        /// <summary>
        /// Global excel password used by helper methods. Default to '1234'.
        /// 呼叫端（例如 Form）可在啟動時覆寫此值。
        /// </summary>
        public static string excelPassword = "1234";

        /// <summary>
        /// 批次將記錄寫入 Excel 的「記錄」工作表。
        /// 若工作表不存在則自動建立，並設定標題、格式、欄位對齊與保護。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="records">待寫入的記錄清單。</param>
        /// <param name="password">保護密碼。</param>
        /// <returns>成功寫入的記錄清單。</returns>
        public static System.Collections.Generic.List<Dto.記錄Dto> AppendRecordsToLogSheet(string excelPath, System.Collections.Generic.List<Dto.記錄Dto> records, string password)
        {
            var result = new System.Collections.Generic.List<Dto.記錄Dto>();
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath) || records == null || records.Count == 0) return result;
            Excel.Application xlApp = null; Excel.Workbook wb = null;
            try
            {
                xlApp = new Excel.Application(); xlApp.DisplayAlerts = false;
                wb = xlApp.Workbooks.Open(excelPath, ReadOnly: false, Password: (object)(password ?? excelPassword ?? string.Empty));
                Excel.Worksheet recordSheet = null;
                try { recordSheet = wb.Sheets["記錄"] as Excel.Worksheet; } catch { recordSheet = null; }
                if (recordSheet == null)
                {
                    recordSheet = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]) as Excel.Worksheet;
                    recordSheet.Name = "記錄";
                    // 批次寫入標題
                    var headerArr = new object[1, 4];
                    headerArr[0, 0] = "刷入時間"; headerArr[0, 1] = "料號"; headerArr[0, 2] = "數量"; headerArr[0, 3] = "操作者";
                    var hStart = recordSheet.Cells[1, 1] as Excel.Range;
                    var hEnd = recordSheet.Cells[1, 4] as Excel.Range;
                    var headerRange = recordSheet.Range[hStart, hEnd];
                    try { headerRange.Value2 = headerArr; } catch { }
                    try { headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; } catch { }
                    try { headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; } catch { }
                    try { headerRange.Font.Bold = true; } catch { }
                    try { recordSheet.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft; } catch { }
                    try { recordSheet.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft; } catch { }
                    try { recordSheet.Columns[3].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; } catch { }
                    try { recordSheet.Columns[4].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; } catch { }
                    try { recordSheet.AutoFilterMode = false; } catch { }
                    try { recordSheet.Columns[1].AutoFit(); recordSheet.Columns[2].AutoFit(); recordSheet.Columns[3].AutoFit(); recordSheet.Columns[4].AutoFit(); } catch { }
                }
                // 若工作表受到保護，先嘗試解除保護以便設定格式與寫入。解除保護應先於 NumberFormat 設定。
                    try { if (!string.IsNullOrEmpty(password)) recordSheet.Unprotect(password); } catch { try { recordSheet.Unprotect(excelPassword); } catch { } }
                // 強制將操作者欄整欄設定為文字格式，並在後續每次寫入時再次確保該儲存格為文字
                try { (recordSheet.Columns[4] as Excel.Range).NumberFormat = "@"; } catch { }
                var recUsed = recordSheet.UsedRange;
                int recLastRow = recUsed.Row + recUsed.Rows.Count - 1;
                int writeRow = recLastRow + 1;
                if (recUsed == null || (recUsed.Rows.Count == 1 && string.IsNullOrWhiteSpace(GetRangeString(recordSheet.Cells[1, 1] as Excel.Range)))) writeRow = 1;

                // Optimized: prepare a batch 2D array and write all rows in one Value2 assignment.
                // Set the operator column NumberFormat once before writing to reduce COM calls.
                try { (recordSheet.Columns[4] as Excel.Range).NumberFormat = "@"; } catch { }
                int count = records.Count;
                try
                {
                    var batch = new object[count, 4];
                    for (int i = 0; i < count; i++)
                    {
                        var dto = records[i];
                        if (dto != null && dto.刷入時間 is DateTime dt)
                            batch[i, 0] = dt.ToString("yyyy-MM-dd HH:mm:ss");
                        else
                            batch[i, 0] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        batch[i, 1] = dto?.料號;
                        batch[i, 2] = dto?.數量;
                        // write operator as plain string (rely on NumberFormat = "@"), avoid visible apostrophe
                        batch[i, 3] = dto?.操作者?.ToString() ?? string.Empty;
                    }

                    Excel.Range rStart = recordSheet.Cells[writeRow, 1] as Excel.Range;
                    Excel.Range rEnd = recordSheet.Cells[writeRow + count - 1, 4] as Excel.Range;
                    Excel.Range writeRange = null;
                    try { if (rStart != null && rEnd != null) writeRange = recordSheet.Range[rStart, rEnd]; } catch { writeRange = null; }

                    if (writeRange != null)
                    {
                        try
                        {
                            writeRange.Value2 = batch;
                            // ensure operator cells keep text format, and re-write per-cell if necessary
                            try
                            {
                                for (int i = 0; i < count; i++)
                                {
                                    var opCell = recordSheet.Cells[writeRow + i, 4] as Excel.Range;
                                    if (opCell != null)
                                    {
                                        try { opCell.NumberFormat = "@"; } catch { }
                                        try { opCell.Value2 = batch[i, 3]; } catch { }
                                    }
                                }
                            }
                            catch { }
                            // add to result list
                            result.AddRange(records);
                        }
                        catch
                        {
                            // fallback: per-row write with robust per-cell handling (preserve previous behavior)
                            for (int i = 0; i < count; i++)
                            {
                                int row = writeRow + i;
                                try { recordSheet.Cells[row, 1] = batch[i, 0]; } catch { }
                                try { recordSheet.Cells[row, 2] = batch[i, 1]; } catch { }
                                try { recordSheet.Cells[row, 3] = batch[i, 2]; } catch { }
                                try
                                {
                                    var opCell = recordSheet.Cells[row, 4] as Excel.Range;
                                    if (opCell != null)
                                    {
                                        try { opCell.NumberFormat = "@"; } catch { }
                                        try { opCell.Value2 = "'" + (batch[i, 3]?.ToString() ?? string.Empty); } catch { }
                                    }
                                    else
                                    {
                                        try { recordSheet.Cells[row, 4] = "'" + (batch[i, 3]?.ToString() ?? string.Empty); } catch { }
                                    }
                                }
                                catch { }
                            }
                            result.AddRange(records);
                        }
                    }
                    else
                    {
                        // writeRange not available, fallback to per-row write
                        for (int i = 0; i < count; i++)
                        {
                            int row = writeRow + i;
                            try { recordSheet.Cells[row, 1] = batch[i, 0]; } catch { }
                            try { recordSheet.Cells[row, 2] = batch[i, 1]; } catch { }
                            try { recordSheet.Cells[row, 3] = batch[i, 2]; } catch { }
                            try
                            {
                                var opCell = recordSheet.Cells[row, 4] as Excel.Range;
                                if (opCell != null)
                                {
                                    try { opCell.NumberFormat = "@"; } catch { }
                                    try { opCell.Value2 = "'" + (batch[i, 3]?.ToString() ?? string.Empty); } catch { }
                                }
                                else
                                {
                                    try { recordSheet.Cells[row, 4] = "'" + (batch[i, 3]?.ToString() ?? string.Empty); } catch { }
                                }
                            }
                            catch { }
                        }
                        result.AddRange(records);
                    }
                }
                catch
                {
                    // On unexpected errors during preparation, fallback to original per-row write
                    try
                    {
                        foreach (var dto in records)
                        {
                            try
                            {
                                var rowArr = new object[1, 4];
                                if (dto != null && dto.刷入時間 is DateTime dt) rowArr[0, 0] = dt.ToString("yyyy-MM-dd HH:mm:ss"); else rowArr[0, 0] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                                rowArr[0, 1] = dto?.料號;
                                rowArr[0, 2] = dto?.數量;
                                rowArr[0, 3] = dto?.操作者?.ToString() ?? string.Empty;
                                try { recordSheet.Cells[writeRow, 1] = rowArr[0, 0]; } catch { }
                                try { recordSheet.Cells[writeRow, 2] = rowArr[0, 1]; } catch { }
                                try { recordSheet.Cells[writeRow, 3] = rowArr[0, 2]; } catch { }
                                try
                                {
                                    var opCell = recordSheet.Cells[writeRow, 4] as Excel.Range;
                                    if (opCell != null)
                                    {
                                        try { opCell.NumberFormat = "@"; } catch { }
                                        try { opCell.Value2 = "'" + (rowArr[0, 3]?.ToString() ?? string.Empty); } catch { }
                                    }
                                    else
                                    {
                                        try { recordSheet.Cells[writeRow, 4] = "'" + (rowArr[0, 3]?.ToString() ?? string.Empty); } catch { }
                                    }
                                }
                                catch { }
                            }
                            catch { }
                            writeRow++;
                        }
                    }
                    catch { }
                }
                try { recordSheet.Columns[1].AutoFit(); recordSheet.Columns[2].AutoFit(); recordSheet.Columns[3].AutoFit(); recordSheet.Columns[4].AutoFit(); } catch { }
                // 保護工作表（解除保護後已寫入並設定格式）
                try { recordSheet.Cells.Locked = true; } catch { }
                try { if (!string.IsNullOrEmpty(password)) recordSheet.Protect(password, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing);
                    else recordSheet.Protect(excelPassword, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing);
                } catch { }
                wb.Save();
            }
            finally
            {
                TryReleaseWorkbook(ref wb);
                TryReleaseApplication(ref xlApp);
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
            return result;
        }

        /// <summary>
        /// 建立並準備「記錄」工作表（若不存在則新增），並設定標題、格式、欄位對齊與保護。
        /// 此工作表用於記錄每次刷入的時間、料號、數量與操作者。
        /// - 標題：刷入時間、料號、數量、操作者
        /// - 標題列置中、粗體
        /// - 時間與料號欄靠左，數量與操作者欄靠右
        /// - 自動調整欄寬
        /// - 關閉篩選功能
        /// - 鎖定所有儲存格並以 excelPassword 保護，禁止篩選
        /// </summary>
        /// <remarks>
        /// 此段程式碼於 <see cref="LoadFirstWorksheetToDataTable(string)"/> 及 <see cref="UpdateShippedAndAppendRecord(string, string, int, string)"/> 方法中出現。
        /// 若「記錄」工作表已存在則略過新增，否則自動建立並套用格式與保護。
        /// </remarks>
        public static DataTable LoadFirstWorksheetToDataTable(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return null;

            Excel.Application xlApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet sheet = null;
            Excel.Range usedRange = null;
            DataTable table = new DataTable();

            try
            {
                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;

                wb = xlApp.Workbooks.Open(path, ReadOnly: true, Password: (object)(excelPassword ?? string.Empty));
                sheet = wb.Sheets[1] as Excel.Worksheet;
                if (sheet == null) return null;

                usedRange = sheet.UsedRange;
                int firstRow = usedRange.Row;
                int firstCol = usedRange.Column;
                int rowsCount = usedRange.Rows.Count;
                int colsCount = usedRange.Columns.Count;
                int lastRow = firstRow + rowsCount - 1;

                object[,] values = null;
                try { var v = usedRange.Value2; if (v is object[,]) values = (object[,])v; }
                catch { values = null; }

                // header detection
                // 優先使用第三列 (第三列 = firstRow + 2) 作為標題，若第三列有任何非空欄位則採用；否則回退至原本的 heuristic（在前 10 列中尋找最多非空欄位）
                int headerRow = -1;
                int preferredOffset = 2; // 0-based offset -> third row
                try
                {
                    if (firstRow + preferredOffset <= lastRow)
                    {
                        int nonEmptyPref = 0;
                        if (values != null)
                        {
                            for (int c = 0; c < colsCount; c++) if (!string.IsNullOrWhiteSpace(values[1 + preferredOffset, 1 + c]?.ToString() ?? string.Empty)) nonEmptyPref++;
                        }
                        else
                        {
                            for (int c = firstCol; c < firstCol + colsCount; c++) if (!string.IsNullOrWhiteSpace(GetRangeString(sheet.Cells[firstRow + preferredOffset, c] as Excel.Range))) nonEmptyPref++;
                        }

                        if (nonEmptyPref > 0) headerRow = firstRow + preferredOffset;
                    }
                }
                catch { headerRow = -1; }

                if (headerRow < 0)
                {
                    int scanLimit = Math.Min(10, rowsCount); int bestNonEmpty = -1;
                    for (int rOffset = 0; rOffset < scanLimit; rOffset++)
                    {
                        int nonEmpty = 0;
                        if (values != null)
                        {
                            for (int c = 0; c < colsCount; c++) if (!string.IsNullOrWhiteSpace(values[1 + rOffset, 1 + c]?.ToString() ?? string.Empty)) nonEmpty++;
                        }
                        else
                        {
                            for (int c = firstCol; c < firstCol + colsCount; c++) if (!string.IsNullOrWhiteSpace(GetRangeString(sheet.Cells[firstRow + rOffset, c] as Excel.Range))) nonEmpty++;
                        }

                        if (nonEmpty > bestNonEmpty) { bestNonEmpty = nonEmpty; headerRow = firstRow + rOffset; }
                    }

                    if (headerRow < 0) return null;
                }

                // Build list of included columns: only visible columns (skip hidden) with non-empty headers
                var includedCols = new System.Collections.Generic.List<int>();
                for (int c = firstCol; c < firstCol + colsCount; c++)
                {
                    // Check if column is hidden
                    bool isHidden = false;
                    try
                    {
                        var colRange = sheet.Columns[c] as Excel.Range;
                        if (colRange != null)
                        {
                            try { isHidden = (bool)colRange.Hidden; }
                            catch { isHidden = false; }
                        }
                    }
                    catch { isHidden = false; }

                    // Skip hidden columns
                    if (isHidden) continue;

                    string rawName = null;
                    if (values != null) rawName = values[headerRow - firstRow + 1, c - firstCol + 1]?.ToString();
                    else rawName = GetRangeString(sheet.Cells[headerRow, c] as Excel.Range);
                    if (string.IsNullOrWhiteSpace(rawName)) continue; // skip columns without header
                    string colName = rawName;
                    string unique = colName; int idx = 1; while (table.Columns.Contains(unique)) { unique = colName + "_" + idx; idx++; }
                    table.Columns.Add(unique);
                    includedCols.Add(c);
                }

                // If no columns have headers, nothing to import
                if (includedCols.Count == 0) return table;

                // Add rows but only for included columns; skip rows that are empty across included columns
                for (int r = headerRow + 1; r <= lastRow; r++)
                {
                    var dr = table.NewRow();
                    bool anyNonEmpty = false;
                    for (int i = 0; i < includedCols.Count; i++)
                    {
                        int c = includedCols[i];
                        string txt = string.Empty;
                        if (values != null)
                        {
                            var obj = values[r - firstRow + 1, c - firstCol + 1]; txt = obj?.ToString() ?? string.Empty;
                        }
                        else txt = GetRangeString(sheet.Cells[r, c] as Excel.Range);
                        dr[i] = txt;
                        if (!string.IsNullOrWhiteSpace(txt)) anyNonEmpty = true;
                    }

                    if (anyNonEmpty) table.Rows.Add(dr);
                }

                // Create and prepare the '記錄' sheet once (do not do this per-row - it is expensive)
                try
                {
                    Excel.Worksheet recordSheet = null;
                    try { recordSheet = wb.Sheets["記錄"] as Excel.Worksheet; } catch { recordSheet = null; }
                    if (recordSheet == null)
                    {
                        recordSheet = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]) as Excel.Worksheet;
                        recordSheet.Name = "記錄";
                        // 批次寫入標題，減少逐格 COM 呼叫
                        Excel.Range headerRange = null;
                        try
                        {
                            var headerArr = new object[1, 4];
                            headerArr[0, 0] = "刷入時間"; headerArr[0, 1] = "料號"; headerArr[0, 2] = "數量"; headerArr[0, 3] = "操作者";
                            var hStart = recordSheet.Cells[1, 1] as Excel.Range;
                            var hEnd = recordSheet.Cells[1, 4] as Excel.Range;
                            headerRange = recordSheet.Range[hStart, hEnd];
                            try { headerRange.Value2 = headerArr; }
                            catch
                            {
                                // fallback to per-cell if batch fails
                                try { recordSheet.Cells[1, 1] = "刷入時間"; } catch { }
                                try { recordSheet.Cells[1, 2] = "料號"; } catch { }
                                try { recordSheet.Cells[1, 3] = "數量"; } catch { }
                                try { recordSheet.Cells[1, 4] = "操作者"; } catch { }
                            }
                        }
                        catch
                        {
                            try { recordSheet.Cells[1, 1] = "刷入時間"; recordSheet.Cells[1, 2] = "料號"; recordSheet.Cells[1, 3] = "數量"; recordSheet.Cells[1, 4] = "操作者"; } catch { }
                        }
                        // 標題：置中、垂直置中、粗體
                        try { if (headerRange != null) headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; } catch { }
                        try { if (headerRange != null) headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; } catch { }
                        try { if (headerRange != null) headerRange.Font.Bold = true; } catch { }

                        // 設定欄位內容對齊：時間靠左、料號靠左、數量靠右、操作者靠右
                        try { recordSheet.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft; } catch { }
                        try { recordSheet.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft; } catch { }
                        try { recordSheet.Columns[3].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; } catch { }
                        try { recordSheet.Columns[4].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; } catch { }
                        // 確保操作者欄為文字格式（防止前導 0 被移除）
                        try { (recordSheet.Columns[4] as Excel.Range).NumberFormat = "@"; } catch { }

                        // 記錄工作表不需要篩選：確保未啟用 AutoFilter
                        try { var sh = recordSheet as Excel.Worksheet; if (sh != null) { try { if (sh.FilterMode) { try { sh.ShowAllData(); } catch { } } } catch { } try { sh.AutoFilterMode = false; } catch { } } } catch { }
                        try { recordSheet.Columns[1].AutoFit(); recordSheet.Columns[2].AutoFit(); recordSheet.Columns[3].AutoFit(); recordSheet.Columns[4].AutoFit(); } catch { }
                    }

                // 無論是否為新建立的 sheet，強制將操作者欄設定為文字格式
                try { (recordSheet.Columns[4] as Excel.Range).NumberFormat = "@"; } catch { }

                    try { if (!string.IsNullOrEmpty(excelPassword)) recordSheet.Unprotect(excelPassword); } catch { }
                    try { recordSheet.Cells.Locked = true; } catch { }
                    // 記錄工作表保護：不允許篩選
                    try { recordSheet.Protect(excelPassword, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing); } catch { }
                }
                catch (Exception ex)
                {
                    // Log and continue; protection failure should not break the read flow
                    try { Logger.LogErrorAsync("Protecting record sheet failed: " + ex.Message).Wait(); } catch { }
                }

                // All relevant rows have been added above (only included headered columns and rows
                // that contain any value in those columns). Return the resulting DataTable.
                return table;
            }
            finally
            {
                TryReleaseRange(ref usedRange);
                TryReleaseWorksheet(ref sheet);
                if (wb != null) { try { wb.Close(false); } catch { } TryReleaseWorkbook(ref wb); }
                if (xlApp != null) { try { xlApp.Quit(); } catch { } TryReleaseApplication(ref xlApp); }
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// 以串流方式讀取 Excel 第一個工作表資料，並以批次方式回傳 DataTable。
        /// <para>
        /// 本方法會以 COM Interop 方式載入 Excel，確保隱藏欄位正確處理。
        /// 由於 NPOI 無法正確偵測所有 Excel 隱藏欄，故本方法不採用 NPOI 快速路徑。
        /// </para>
        /// <remarks>
        /// - 目前實作會一次性載入所有資料並呼叫 batchCallback，未真正分批。
        /// - 若未來需處理大型檔案，可依 batchSize 參數分批回傳。
        /// - 若載入失敗，將不會擲出例外，僅靜默結束。
        /// </remarks>
        /// <param name="path">Excel 檔案路徑。</param>
        /// <param name="batchCallback">每批資料回呼委派，參數為 DataTable。</param>
        /// <param name="batchSize">每批最大筆數（目前未實作分批，僅保留參數）。</param>
        public static void LoadFirstWorksheetToDataTableStreaming(string path, Action<DataTable> batchCallback, int batchSize = 1000)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return;
            try
            {
                // Always use COM-based loading to ensure hidden columns are properly skipped
                // NPOI fast path is deliberately disabled for this scenario
                var full = LoadFirstWorksheetToDataTable(path);
                if (full != null) batchCallback?.Invoke(full);
                return;
            }
            catch
            {
                // If COM loading also fails, silently return - error already handled in LoadFirstWorksheetToDataTable
                try { var full = LoadFirstWorksheetToDataTable(path); if (full != null) batchCallback?.Invoke(full); } catch { }
            }
        }

        /// <summary>
        /// 檢查指定 Excel 檔案中被隱藏的欄位，並從 uiTable 中移除對應的欄位（根據標頭名稱匹配）。
        /// 使用 NPOI 快速路徑 (.xlsx/.xls/.xlsm) 或 COM fallback，並回傳偵測到的隱藏標頭清單。
        /// 若發生任何錯誤，方法會靜默失敗並回傳空集合。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑</param>
        /// <param name="uiTable">將被修改的 DataTable（in-place）</param>
        /// <returns>偵測到的隱藏欄位標頭（原始字串）</returns>
        public static List<string> RemoveHiddenColumnsFromDataTable(string excelPath, DataTable uiTable)
        {
            var resultHeaders = new List<string>();
            try
            {
                if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath)) return resultHeaders;
                if (uiTable == null || uiTable.Columns.Count == 0) return resultHeaders;

                var ext = Path.GetExtension(excelPath)?.ToLowerInvariant() ?? string.Empty;

                // NPOI fast path for .xlsx/.xls and attempt for .xlsm
                if (ext == ".xlsx" || ext == ".xls")
                {
                    var hiddenHeaders = new List<string>();
                    try
                    {
                        using (var fs = File.OpenRead(excelPath))
                        {
                            IWorkbook wbN = ext == ".xlsx" ? (IWorkbook)new XSSFWorkbook(fs) : new HSSFWorkbook(fs);
                            if (wbN == null || wbN.NumberOfSheets <= 0) return resultHeaders;
                            ISheet? sheet = null;
                            try
                            {
                                for (int i = 0; i < wbN.NumberOfSheets; i++)
                                {
                                    var nm = wbN.GetSheetName(i) ?? string.Empty;
                                    if (string.Equals(nm.Trim(), "總表", StringComparison.OrdinalIgnoreCase)) { sheet = wbN.GetSheetAt(i); break; }
                                }
                            }
                            catch { sheet = null; }
                            if (sheet == null) sheet = wbN.GetSheetAt(0);
                            if (sheet == null) return resultHeaders;

                            // detect header row (prefer 3rd row)
                            int headerRowIdx = -1; int bestCount = -1; int scanLimit = Math.Min(10, sheet.LastRowNum + 1);
                            try
                            {
                                if (sheet.LastRowNum >= 2)
                                {
                                    var pref = sheet.GetRow(2);
                                    if (pref != null)
                                    {
                                        int cnt = 0;
                                        int first = pref.FirstCellNum >= 0 ? pref.FirstCellNum : 0;
                                        int last = pref.LastCellNum >= 0 ? pref.LastCellNum : first;
                                        for (int c = first; c <= last; c++) { var cell = pref.GetCell(c); if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString())) cnt++; }
                                        if (cnt > 0) headerRowIdx = 2;
                                    }
                                }
                            }
                            catch { headerRowIdx = -1; }
                            if (headerRowIdx < 0)
                            {
                                for (int r = 0; r < scanLimit; r++)
                                {
                                    var row = sheet.GetRow(r);
                                    if (row == null) continue;
                                    int cnt = 0;
                                    int first = row.FirstCellNum >= 0 ? row.FirstCellNum : 0;
                                    int last = row.LastCellNum >= 0 ? row.LastCellNum : first;
                                    for (int c = first; c <= last; c++) { var cell = row.GetCell(c); if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString())) cnt++; }
                                    if (cnt > bestCount) { bestCount = cnt; headerRowIdx = r; }
                                }
                            }
                            if (headerRowIdx < 0) headerRowIdx = 0;

                            var headerRow = sheet.GetRow(headerRowIdx);
                            if (headerRow == null) return resultHeaders;
                            int maxCol = headerRow.LastCellNum >= 0 ? headerRow.LastCellNum : 0;

                            for (int c = 0; c < maxCol; c++)
                            {
                                bool isHidden = false;
                                try { isHidden = sheet.IsColumnHidden(c); } catch { isHidden = false; }
                                if (!isHidden)
                                {
                                    try { isHidden = sheet.GetColumnWidth(c) <= 1; } catch { }
                                }
                                if (!isHidden) continue;
                                try
                                {
                                    var cell = headerRow.GetCell(c);
                                    var hn = cell?.ToString()?.Trim();
                                    if (!string.IsNullOrWhiteSpace(hn)) hiddenHeaders.Add(hn!);
                                }
                                catch { }
                            }
                        }
                    }
                    catch { hiddenHeaders.Clear(); }

                    var hiddenSanSet = new HashSet<string>(hiddenHeaders.Select(Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch));
                    var protectedHeaders = new[] { "備註", "remark", "note", "備考", "註記" };
                    var protectedNorm = new HashSet<string>(protectedHeaders.Select(Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch));

                    if (hiddenSanSet.Count > 0)
                    {
                        var toRemove = new List<DataColumn>();
                        foreach (DataColumn dc in uiTable.Columns)
                        {
                            try
                            {
                                var norm = Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch(dc.ColumnName);
                                if (!string.IsNullOrEmpty(norm) && hiddenSanSet.Contains(norm) && !protectedNorm.Contains(norm))
                                    toRemove.Add(dc);
                            }
                            catch { }
                        }
                        foreach (var dc in toRemove.Distinct().ToList()) { try { uiTable.Columns.Remove(dc); } catch { } }
                        resultHeaders = hiddenHeaders.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    }
                    else
                    {
                        resultHeaders = new List<string>();
                    }
                    return resultHeaders;
                }

                // xlsm: try NPOI then fallback to COM
                if (ext == ".xlsm")
                {
                    var hiddenHeaders = new List<string>();
                    bool npoiSucceeded = false;
                    try
                    {
                        using (var fs = File.OpenRead(excelPath))
                        {
                            IWorkbook wbN = new XSSFWorkbook(fs);
                            if (wbN != null && wbN.NumberOfSheets > 0)
                            {
                                ISheet? sheet = null;
                                try { for (int i = 0; i < wbN.NumberOfSheets; i++) { var nm = wbN.GetSheetName(i) ?? string.Empty; if (string.Equals(nm.Trim(), "總表", StringComparison.OrdinalIgnoreCase)) { sheet = wbN.GetSheetAt(i); break; } } } catch { sheet = null; }
                                if (sheet == null) sheet = wbN.GetSheetAt(0);
                                if (sheet != null)
                                {
                                    int headerRowIdx = -1; int bestCount = -1; int scanLimit = Math.Min(10, sheet.LastRowNum + 1);
                                    try
                                    {
                                        if (sheet.LastRowNum >= 2)
                                        {
                                            var pref = sheet.GetRow(2);
                                            if (pref != null)
                                            {
                                                int cnt = 0; int first = pref.FirstCellNum >= 0 ? pref.FirstCellNum : 0; int last = pref.LastCellNum >= 0 ? pref.LastCellNum : first;
                                                for (int c = first; c <= last; c++) { var cell = pref.GetCell(c); if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString())) cnt++; }
                                                if (cnt > 0) headerRowIdx = 2;
                                            }
                                        }
                                    }
                                    catch { headerRowIdx = -1; }
                                    if (headerRowIdx < 0)
                                    {
                                        for (int r = 0; r < scanLimit; r++) { var row = sheet.GetRow(r); if (row == null) continue; int cnt = 0; int first = row.FirstCellNum >= 0 ? row.FirstCellNum : 0; int last = row.LastCellNum >= 0 ? row.LastCellNum : first; for (int c = first; c <= last; c++) { var cell = row.GetCell(c); if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString())) cnt++; } if (cnt > bestCount) { bestCount = cnt; headerRowIdx = r; } }
                                    }
                                    if (headerRowIdx < 0) headerRowIdx = 0;

                                    var headerRow = sheet.GetRow(headerRowIdx);
                                    if (headerRow != null)
                                    {
                                        int maxCol = headerRow.LastCellNum >= 0 ? headerRow.LastCellNum : 0;
                                        for (int c = 0; c < maxCol; c++)
                                        {
                                            bool isHidden = false;
                                            try { isHidden = sheet.IsColumnHidden(c); } catch { isHidden = false; }
                                            if (!isHidden) { try { isHidden = sheet.GetColumnWidth(c) <= 1; } catch { } }
                                            if (!isHidden) continue;
                                            try { var cell = headerRow.GetCell(c); var hn = cell?.ToString()?.Trim(); if (!string.IsNullOrWhiteSpace(hn)) hiddenHeaders.Add(hn!); } catch { }
                                        }
                                    }
                                }
                            }
                        }
                        npoiSucceeded = true;
                    }
                    catch { npoiSucceeded = false; }

                    if (npoiSucceeded)
                    {
                        var hiddenSanSet = new HashSet<string>(hiddenHeaders.Select(Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch));
                        var protectedHeaders = new[] { "備註", "remark", "note", "備記", "註記" };
                        var protectedNorm = new HashSet<string>(protectedHeaders.Select(Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch));

                        if (hiddenSanSet.Count > 0)
                        {
                            var toRemove = new List<DataColumn>();
                            foreach (DataColumn dc in uiTable.Columns)
                            {
                                try
                                {
                                    var norm = Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch(dc.ColumnName);
                                    if (!string.IsNullOrEmpty(norm) && hiddenSanSet.Contains(norm) && !protectedNorm.Contains(norm)) toRemove.Add(dc);
                                }
                                catch { }
                            }
                            foreach (var dc in toRemove.Distinct().ToList()) { try { uiTable.Columns.Remove(dc); } catch { } }
                            resultHeaders = hiddenHeaders.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                        }
                        else resultHeaders = new List<string>();
                        return resultHeaders;
                    }
                    // fallthrough to COM
                }
            }
            catch { /* ignore and continue to COM path below */ }

            // COM fallback
            try
            {
                Excel.Application xlApp = null; Excel.Workbook wb = null; Excel.Worksheet ws = null;
                try
                {
                    xlApp = new Excel.Application { DisplayAlerts = false, Visible = false };
                    wb = xlApp.Workbooks.Open(excelPath, ReadOnly: true);
                    try { ws = wb.Worksheets["總表"] as Excel.Worksheet; } catch { ws = null; }
                    if (ws == null) ws = wb.Worksheets[1] as Excel.Worksheet;
                    if (ws == null) return resultHeaders;

                    var hiddenHeaders = new List<string>();
                    object[,]? cachedHeaderValues = null;
                    int cachedHeaderColsToScan = 0;
                    var used = ws.UsedRange; if (used == null) return resultHeaders;
                    int rowCount = used.Rows.Count; int colCount = used.Columns.Count;

                    bool skipXlsmScan = false; int maxColsToScan = colCount;
                    try
                    {
                        var app = ConfigurationManager.AppSettings;
                        if (app != null)
                        {
                            bool.TryParse(app["SkipHiddenScanForXlsm"], out skipXlsmScan);
                            if (int.TryParse(app["XlsmHiddenScanMaxColumns"], out int cfgMax) && cfgMax > 0)
                                maxColsToScan = Math.Min(colCount, cfgMax);
                            else maxColsToScan = Math.Min(colCount, 150);
                        }
                        else maxColsToScan = Math.Min(colCount, 150);
                    }
                    catch { maxColsToScan = Math.Min(colCount, 150); }

                    if (skipXlsmScan) return resultHeaders;

                    // detect header row
                    int headerSearchMax = Math.Min(3, Math.Max(1, rowCount));
                    int detectedHeaderRow = 1;
                    string GetCellString(object[,]? arr, int r, int c)
                    {
                        if (arr == null) return string.Empty;
                        int rb = arr.GetLowerBound(0);
                        int cb = arr.GetLowerBound(1);
                        object? v = null;
                        try { v = arr[rb + (r - 1), cb + (c - 1)]; } catch { return string.Empty; }
                        return v?.ToString()?.Trim() ?? string.Empty;
                    }

                    try
                    {
                        Func<string, string> SanitizeForMatch = Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch;
                        var uiSan = new HashSet<string>(uiTable.Columns.Cast<DataColumn>().Select(dc => SanitizeForMatch(dc.ColumnName)));

                        int colsToScan = Math.Min(colCount, maxColsToScan);
                        Excel.Range? hdrRange = null;
                        object[,]? hdrVals = null;
                        try
                        {
                            hdrRange = used.Range[used.Cells[1, 1], used.Cells[headerSearchMax, colsToScan]] as Excel.Range;
                            var raw = hdrRange?.Value2;
                            if (raw is object[,]) hdrVals = (object[,])raw;
                            else if (raw is object single)
                            {
                                hdrVals = new object[1, 1]; hdrVals[0, 0] = single;
                            }
                        }
                        catch { hdrVals = null; }
                        finally { try { if (hdrRange != null) ReleaseComObjectSafe(hdrRange); } catch { } }

                        int bestMatches = 0;
                        for (int hr = 1; hr <= headerSearchMax; hr++)
                        {
                            int matches = 0;
                            for (int c = 1; c <= colsToScan; c++)
                            {
                                try { string hn = GetCellString(hdrVals, hr, c); if (string.IsNullOrEmpty(hn)) continue; var hns = SanitizeForMatch(hn); if (uiSan.Contains(hns)) matches++; } catch { }
                            }
                            if (matches > bestMatches) { bestMatches = matches; detectedHeaderRow = hr; }
                        }

                        cachedHeaderValues = hdrVals; cachedHeaderColsToScan = colsToScan;
                    }
                    catch { detectedHeaderRow = 1; }

                    for (int c = 1; c <= maxColsToScan; c++)
                    {
                        bool isHidden = false;
                        try
                        {
                            var colRange = ws.Columns[c] as Excel.Range;
                            if (colRange != null)
                            {
                                try { isHidden = (bool)colRange.Hidden; } catch { isHidden = false; }
                                if (!isHidden) { try { var entire = colRange.EntireColumn as Excel.Range; if (entire != null) isHidden = (bool)entire.Hidden; } catch { } }
                                if (!isHidden) { try { double w = 0; try { w = (double)colRange.ColumnWidth; } catch { w = 0; } if (w <= 0.1) isHidden = true; } catch { } }
                            }
                        }
                        catch { isHidden = false; }

                        if (!isHidden) continue;

                        try
                        {
                            string hn = string.Empty;
                            if (cachedHeaderValues != null && detectedHeaderRow <= 3 && c <= cachedHeaderColsToScan)
                            {
                                hn = GetCellString(cachedHeaderValues, detectedHeaderRow, c);
                            }
                            else
                            {
                                var hv = (used.Cells[detectedHeaderRow, c] as Excel.Range)?.Value2;
                                hn = hv?.ToString()?.Trim() ?? string.Empty;
                            }
                            if (!string.IsNullOrEmpty(hn)) hiddenHeaders.Add(hn!);
                        }
                        catch { }
                    }
                    if (hiddenHeaders.Count == 0) return resultHeaders;

                    var hiddenSan = hiddenHeaders.Select(h => new { Orig = h, Norm = Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch(h) }).ToList();
                    var protectedHeaders2 = new[] { "備註", "remark", "note", "備考", "註記" };
                    var protectedNorm = new HashSet<string>(protectedHeaders2.Select(h => Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch(h)));

                    var toRemove = new List<DataColumn>();

                    // exact name
                    foreach (DataColumn dc in uiTable.Columns)
                    {
                        try { if (hiddenHeaders.Any(h => string.Equals(h, dc.ColumnName, StringComparison.OrdinalIgnoreCase))) { if (!protectedNorm.Contains(Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch(dc.ColumnName))) toRemove.Add(dc); } } catch { }
                    }

                    // sanitized exact
                    foreach (DataColumn dc in uiTable.Columns)
                    {
                        try { if (toRemove.Contains(dc)) continue; var dcn = Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch(dc.ColumnName); if (protectedNorm.Contains(dcn)) continue; if (hiddenSan.Any(h => !string.IsNullOrEmpty(h.Norm) && h.Norm == dcn)) toRemove.Add(dc); } catch { }
                    }

                    // substring sanitized
                    foreach (DataColumn dc in uiTable.Columns)
                    {
                        try { if (toRemove.Contains(dc)) continue; var dcn = Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch(dc.ColumnName); if (string.IsNullOrEmpty(dcn)) continue; if (protectedNorm.Contains(dcn)) continue; if (hiddenSan.Any(h => !string.IsNullOrEmpty(h.Norm) && (dcn.Contains(h.Norm) || h.Norm.Contains(dcn)))) toRemove.Add(dc); } catch { }
                    }

                    // position-based fallback
                    for (int c = 1; c <= (ws.UsedRange.Columns.Count); c++)
                    {
                        try
                        {
                            var colRange = ws.Columns[c] as Excel.Range; bool isHidden = false; if (colRange != null) { try { isHidden = (bool)colRange.Hidden; } catch { isHidden = false; } }
                            if (!isHidden) continue;
                            int uiIndex = c - 1;
                            if (uiIndex >= 0 && uiIndex < uiTable.Columns.Count) { var dc = uiTable.Columns[uiIndex]; if (!protectedNorm.Contains(Automatic_Storage.Utilities.TextParsing.SanitizeHeaderForMatch(dc.ColumnName))) if (!toRemove.Contains(dc)) toRemove.Add(dc); }
                        }
                        catch { }
                    }

                    foreach (var dc in toRemove.Distinct().ToList()) { try { uiTable.Columns.Remove(dc); } catch { } }
                    resultHeaders = hiddenHeaders.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    return resultHeaders;
                }
                catch { return resultHeaders; }
                finally { try { if (wb != null) { wb.Close(false); ReleaseComObjectSafe(wb); } } catch { } try { if (xlApp != null) { xlApp.Quit(); ReleaseComObjectSafe(xlApp); } } catch { } }
            }
            catch { return resultHeaders; }
        }

        /// <summary>
        /// 更新 Excel 主表的「實發數量」或「發料數量」欄位，並於「記錄」工作表新增一筆刷入記錄。
        /// <para>
        /// 操作流程：<br/>
        /// 1. 依據檔案自動偵測標題列與「料號」及「實發數量/發料數量」欄位。<br/>
        /// 2. 尋找指定料號的資料列，將「實發數量/發料數量」加上指定數量。<br/>
        /// 3. 若「記錄」工作表不存在則自動建立，並設定標題、格式、欄位對齊與保護。<br/>
        /// 4. 於「記錄」工作表新增一筆包含刷入時間、料號、數量、操作者的記錄。<br/>
        /// 5. 回寫 Excel 並針對主表進行欄位鎖定與保護（僅允許「實發數量/發料數量」可編輯，其餘鎖定）。
        /// </para>
        /// <remarks>
        /// - 若找不到料號或數量欄位，將擲出例外。
        /// - 若「記錄」工作表已存在則直接新增記錄，否則自動建立並套用格式與保護。
        /// - 本方法會自動釋放所有 COM 物件並強制 GC，避免 Excel 殘留於記憶體。
        /// </remarks>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="materialCode">要更新的料號。</param>
        /// <param name="qty">要加總的數量。</param>
        /// <param name="operatorName">操作者名稱，將記錄於「記錄」工作表。</param>
        /// <exception cref="FileNotFoundException">找不到指定 Excel 檔案時擲出。</exception>
        /// <exception cref="Exception">找不到標題列、料號欄或數量欄時擲出。</exception>
        public static void UpdateShippedAndAppendRecord(string excelPath, string materialCode, int qty, string operatorName)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath)) throw new FileNotFoundException("Excel not found", excelPath);

            Excel.Application xlApp = null; Excel.Workbook wb = null;
            try
            {
                xlApp = new Excel.Application(); xlApp.DisplayAlerts = false;
                wb = xlApp.Workbooks.Open(excelPath, ReadOnly: false, Password: (object)(excelPassword ?? string.Empty));
                Excel.Worksheet mainSheet = wb.Sheets[1] as Excel.Worksheet;
                var used = mainSheet.UsedRange;

                int firstRow = used.Row; int firstCol = used.Column; int rows = used.Rows.Count; int cols = used.Columns.Count; int lastRow = firstRow + rows - 1;
                object[,] values = null; try { var v = used.Value2; if (v is object[,]) values = (object[,])v; } catch { values = null; }

                int headerRow = -1; int scanLimit = Math.Min(10, rows); int bestNonEmpty = -1;
                for (int rOffset = 0; rOffset < scanLimit; rOffset++)
                {
                    int nonEmpty = 0;
                    if (values != null)
                    {
                        for (int c = 0; c < cols; c++) if (values[1 + rOffset, 1 + c] != null && !string.IsNullOrWhiteSpace(values[1 + rOffset, 1 + c].ToString())) nonEmpty++;
                    }
                    else
                    {
                        for (int c = firstCol; c < firstCol + cols; c++) if (!string.IsNullOrWhiteSpace(GetRangeString(mainSheet.Cells[firstRow + rOffset, c] as Excel.Range))) nonEmpty++;
                    }
                    if (nonEmpty > bestNonEmpty) { bestNonEmpty = nonEmpty; headerRow = firstRow + rOffset; }
                }

                if (headerRow < 0) throw new Exception("Cannot detect header row");

                int materialCol = -1; int shippedCol = -1;
                for (int c = 0; c < cols; c++)
                {
                    string h = null;
                    if (values != null) h = values[headerRow - firstRow + 1, c + 1]?.ToString(); else h = GetRangeString(mainSheet.Cells[headerRow, firstCol + c] as Excel.Range);
                    if (string.Equals(h, "料號", StringComparison.OrdinalIgnoreCase) || string.Equals(h, "品號", StringComparison.OrdinalIgnoreCase)) materialCol = firstCol + c;
                    if (string.Equals(h, "實發數量", StringComparison.OrdinalIgnoreCase) || string.Equals(h, "發料數量", StringComparison.OrdinalIgnoreCase)) shippedCol = firstCol + c;
                }

                if (materialCol <= 0) throw new Exception("Material column not found");
                if (shippedCol <= 0) throw new Exception("Shipped column not found");

                int targetRow = -1;
                if (values != null)
                {
                    for (int r = headerRow + 1; r <= lastRow; r++)
                    {
                        var cellObj = values[r - firstRow + 1, materialCol - firstCol + 1];
                        if (cellObj != null && string.Equals(cellObj.ToString().Trim(), materialCode.Trim(), StringComparison.OrdinalIgnoreCase)) { targetRow = r; break; }
                    }
                }
                else
                {
                    for (int r = headerRow + 1; r <= lastRow; r++)
                    {
                        var v = GetRangeString(mainSheet.Cells[r, materialCol] as Excel.Range);
                        if (!string.IsNullOrWhiteSpace(v) && string.Equals(v.Trim(), materialCode.Trim(), StringComparison.OrdinalIgnoreCase)) { targetRow = r; break; }
                    }
                }

                if (targetRow <= 0) throw new Exception("Material row not found in Excel");

                var shippedCell = mainSheet.Cells[targetRow, shippedCol] as Excel.Range;
                double currentVal = 0; try { var tmp = shippedCell.Value2; if (tmp != null) currentVal = Convert.ToDouble(tmp); } catch { currentVal = 0; }
                double newVal = currentVal + qty; shippedCell.Value2 = newVal;

                Excel.Worksheet recordSheet = null;
                try { recordSheet = wb.Sheets["記錄"] as Excel.Worksheet; } catch { recordSheet = null; }
                if (recordSheet == null)
                {
                    recordSheet = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]) as Excel.Worksheet;
                    recordSheet.Name = "記錄";
                    // 批次寫入標題避免逐格呼叫
                    try
                    {
                        var headerArr = new object[1, 4];
                        headerArr[0, 0] = "刷入時間"; headerArr[0, 1] = "料號"; headerArr[0, 2] = "數量"; headerArr[0, 3] = "操作者";
                        var hStart = recordSheet.Cells[1, 1] as Excel.Range;
                        var hEnd = recordSheet.Cells[1, 4] as Excel.Range;
                        var headerRange = recordSheet.Range[hStart, hEnd];
                        try { headerRange.Value2 = headerArr; }
                        catch
                        {
                            try { recordSheet.Cells[1, 1] = "刷入時間"; recordSheet.Cells[1, 2] = "料號"; recordSheet.Cells[1, 3] = "數量"; recordSheet.Cells[1, 4] = "操作者"; } catch { }
                        }
                        // 標題：置中、垂直置中、粗體
                        try { headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; } catch { }
                        try { headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; } catch { }
                        try { headerRange.Font.Bold = true; } catch { }

                        // 設定欄位內容對齊：時間靠左、料號靠左、數量靠右、操作者靠右
                        try { recordSheet.Columns[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft; } catch { }
                        try { recordSheet.Columns[2].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft; } catch { }
                        try { recordSheet.Columns[3].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; } catch { }
                        try { recordSheet.Columns[4].HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; } catch { }

                        // 記錄工作表不需要篩選：確保未啟用 AutoFilter
                        try { var sh = recordSheet as Excel.Worksheet; if (sh != null) { try { if (sh.FilterMode) { try { sh.ShowAllData(); } catch { } } } catch { } try { sh.AutoFilterMode = false; } catch { } } } catch { }
                        try { recordSheet.Columns[1].AutoFit(); recordSheet.Columns[2].AutoFit(); recordSheet.Columns[3].AutoFit(); recordSheet.Columns[4].AutoFit(); } catch { }
                    }
                    catch { }
                }

                var recUsed = recordSheet.UsedRange;
                int recLastRow = recUsed.Row + recUsed.Rows.Count - 1;
                int writeRow = recLastRow + 1;
                if (recUsed == null || (recUsed.Rows.Count == 1 && string.IsNullOrWhiteSpace(GetRangeString(recordSheet.Cells[1, 1] as Excel.Range)))) writeRow = 1;

                // 批次寫入一列資料以避免多次逐格 COM 呼叫
                    try
                    {
                        var rowArr = new object[1, 4];
                        rowArr[0, 0] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        rowArr[0, 1] = materialCode; rowArr[0, 2] = qty; rowArr[0, 3] = (operatorName ?? string.Empty);
                        var rStart = recordSheet.Cells[writeRow, 1] as Excel.Range;
                        var rEnd = recordSheet.Cells[writeRow, 4] as Excel.Range;
                        Excel.Range writeRange = null;
                        try { if (rStart != null && rEnd != null) writeRange = recordSheet.Range[rStart, rEnd]; } catch { }
                        // 解除保護 -> 設定第4欄為文字格式 -> 批次寫入；若失敗則逐格寫入並確保操作者為文字
                        try { try { recordSheet.Unprotect(excelPassword); } catch { } } catch { }
                        try { (recordSheet.Columns[4] as Excel.Range).NumberFormat = "@"; } catch { }
                        try { (recordSheet.Cells[writeRow, 4] as Excel.Range).NumberFormat = "@"; } catch { }
                        try
                        {
                            if (writeRange != null) writeRange.Value2 = rowArr;
                            // ensure the operator cell is explicitly set to Text and re-written per-cell
                            try
                            {
                                var opCell = recordSheet.Cells[writeRow, 4] as Excel.Range;
                                if (opCell != null)
                                {
                                    try { opCell.NumberFormat = "@"; } catch { }
                                    try { opCell.Value2 = operatorName; } catch { }
                                }
                            }
                            catch { }
                        }
                    catch
                    {
                        try { recordSheet.Cells[writeRow, 1] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); } catch { }
                        try { recordSheet.Cells[writeRow, 2] = materialCode; } catch { }
                        try { recordSheet.Cells[writeRow, 3] = qty; } catch { }
                        try
                        {
                            var opCell = recordSheet.Cells[writeRow, 4] as Excel.Range;
                            if (opCell != null)
                            {
                                try { opCell.NumberFormat = "@"; } catch { }
                                try { opCell.Value2 = "'" + (operatorName ?? string.Empty); } catch { }
                            }
                            else
                            {
                                try { recordSheet.Cells[writeRow, 4] = "'" + (operatorName ?? string.Empty); } catch { }
                            }
                        }
                        catch { }
                    }
                }
                catch
                {
                    try { recordSheet.Cells[writeRow, 1] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); } catch { }
                    try { recordSheet.Cells[writeRow, 2] = materialCode; } catch { }
                    try { recordSheet.Cells[writeRow, 3] = qty; } catch { }
                    try { var opCell2 = recordSheet.Cells[writeRow, 4] as Excel.Range; if (opCell2 != null) { try { opCell2.NumberFormat = "@"; } catch { } try { opCell2.Value2 = "'" + (operatorName ?? string.Empty); } catch { } } else { recordSheet.Cells[writeRow, 4] = "'" + (operatorName ?? string.Empty); } } catch { }
                }
                try { recordSheet.Columns[1].AutoFit(); recordSheet.Columns[2].AutoFit(); recordSheet.Columns[3].AutoFit(); recordSheet.Columns[4].AutoFit(); } catch { }

                // 在回寫後，針對主表設定欄位鎖定並保護（同 ProtectWorksheet 行為）
                try
                {
                    // reuse detected values from above (firstRow/firstCol/rows/cols/lastRow/values)
                    // 找出應鎖定的欄
                    int lockColIndex = -1;
                    var ext = Path.GetExtension(excelPath)?.ToLowerInvariant() ?? "";
                    bool isXlsm = ext == ".xlsm";
                    for (int c = 0; c < cols; c++)
                    {
                        string h = null;
                        if (values != null) h = values[headerRow - firstRow + 1, c + 1]?.ToString(); else h = GetRangeString(mainSheet.Cells[headerRow, firstCol + c] as Excel.Range);
                        if (isXlsm)
                        {
                            if (string.Equals(h, "實發數量", StringComparison.OrdinalIgnoreCase)) { lockColIndex = firstCol + c; break; }
                        }
                        else
                        {
                            if (string.Equals(h, "發料數量", StringComparison.OrdinalIgnoreCase)) { lockColIndex = firstCol + c; break; }
                        }
                    }

                    if (lockColIndex <= 0 && shippedCol > 0) lockColIndex = shippedCol;

                    try { if (!string.IsNullOrEmpty(excelPassword)) mainSheet.Unprotect(excelPassword); } catch { }
                    try { mainSheet.Cells.Locked = false; } catch { }
                    if (lockColIndex > 0)
                    {
                        try
                        {
                            var lockRange = mainSheet.Columns[lockColIndex] as Excel.Range;
                            lockRange.Locked = true;
                        }
                        catch { }
                    }
                    else
                    {
                        try { mainSheet.Cells.Locked = true; } catch { }
                    }

                    // 為第一列套用 AutoFilter（若尚未套用），再保護主表，允許篩選
                    try
                    {
                        int colCount = cols;
                        if (colCount <= 0)
                        {
                            try { colCount = (mainSheet.UsedRange.Columns.Count); } catch { }
                        }
                        if (colCount > 0)
                        {
                            var headerRange = mainSheet.Range[mainSheet.Cells[headerRow, firstCol], mainSheet.Cells[headerRow, firstCol + colCount - 1]];
                            try { headerRange.AutoFilter(Type.Missing); } catch { }
                        }
                    }
                    catch { }

                    // 保護主表，允許篩選
                    try { mainSheet.Protect(excelPassword, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing); } catch { }
                }
                catch { }

                wb.Save();
            }
            finally
            {
                TryReleaseWorkbook(ref wb);
                TryReleaseApplication(ref xlApp);
                GC.Collect(); GC.WaitForPendingFinalizers();
            }

        }

        /// <summary>
        /// 保護 Excel 工作表，僅允許特定欄位（如「發料數量」或「實發數量」）可編輯，其餘儲存格皆鎖定。
        /// 若為 xlsm 檔案且 <paramref name="protectShippedColumnForXlsm"/> 為 true，則鎖定「實發數量」欄，否則鎖定「發料數量」欄。
        /// 若找不到目標欄，則整張表皆鎖定。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="password">保護密碼，若為空則使用全域 <see cref="excelPassword"/>。</param>
        /// <param name="protectShippedColumnForXlsm">xlsm 檔案時是否鎖定「實發數量」欄。</param>
        public static void ProtectWorksheet(string excelPath, string password, bool protectShippedColumnForXlsm)
        {
            // If caller passed empty password, fallback to global configured password
            if (string.IsNullOrEmpty(password)) password = excelPassword;
            Excel.Application xlApp = null; Excel.Workbook wb = null;
            try
            {
                xlApp = new Excel.Application(); xlApp.DisplayAlerts = false;
                wb = xlApp.Workbooks.Open(excelPath, ReadOnly: false, Password: (object)(excelPassword ?? string.Empty));
                Excel.Worksheet mainSheet = wb.Sheets[1] as Excel.Worksheet;

                var used = mainSheet.UsedRange;
                int firstRow = used.Row;
                int firstCol = used.Column;
                int rows = used.Rows.Count;
                int cols = used.Columns.Count;
                int lastRow = firstRow + rows - 1;
                int lastCol = firstCol + cols - 1;

                object[,] values = null;
                try { var v = used.Value2; if (v is object[,]) values = (object[,])v; } catch { values = null; }

                // 偵測 header 同前邏輯
                int headerRow = -1; int scanLimit = Math.Min(10, rows); int bestNonEmpty = -1;
                for (int rOffset = 0; rOffset < scanLimit; rOffset++)
                {
                    int nonEmpty = 0;
                    if (values != null)
                    {
                        for (int c = 0; c < cols; c++) if (values[1 + rOffset, 1 + c] != null && !string.IsNullOrWhiteSpace(values[1 + rOffset, 1 + c].ToString())) nonEmpty++;
                    }
                    else
                    {
                        for (int c = firstCol; c < firstCol + cols; c++) if (!string.IsNullOrWhiteSpace(GetRangeString(mainSheet.Cells[firstRow + rOffset, c] as Excel.Range))) nonEmpty++;
                    }
                    if (nonEmpty > bestNonEmpty) { bestNonEmpty = nonEmpty; headerRow = firstRow + rOffset; }
                }

                if (headerRow < 0) headerRow = firstRow;

                // 找出要鎖定的欄位索引
                int shippedCol = -1;
                for (int c = 0; c < cols; c++)
                {
                    string h = null;
                    if (values != null) h = values[headerRow - firstRow + 1, c + 1]?.ToString(); else h = GetRangeString(mainSheet.Cells[headerRow, firstCol + c] as Excel.Range);
                    if (protectShippedColumnForXlsm)
                    {
                        if (string.Equals(h, "實發數量", StringComparison.OrdinalIgnoreCase) || string.Equals(h, "實發", StringComparison.OrdinalIgnoreCase)) shippedCol = firstCol + c;
                    }
                    else
                    {
                        if (string.Equals(h, "發料數量", StringComparison.OrdinalIgnoreCase) || string.Equals(h, "發料", StringComparison.OrdinalIgnoreCase)) shippedCol = firstCol + c;
                    }
                }

                // 預設先解保護，解鎖所有儲存格，然後再鎖定目標欄
                // 如果提供的 password 為空，已在上方改為使用 global excelPassword
                try { if (!string.IsNullOrEmpty(password)) mainSheet.Unprotect(password); } catch { }
                try { mainSheet.Cells.Locked = false; } catch { }

                if (shippedCol > 0)
                {
                    var lockRange = mainSheet.Range[mainSheet.Cells[headerRow, shippedCol], mainSheet.Cells[lastRow, shippedCol]];
                    lockRange.Locked = true;
                }
                else
                {
                    // 若沒找到目標欄，則鎖定整張表以保守處理
                    mainSheet.Cells.Locked = true;
                }

                // 最後保護工作表並允許篩選
                mainSheet.Protect(password, AllowFiltering: true);
                wb.Save();
            }
            finally
            {
                TryReleaseWorkbook(ref wb);
                TryReleaseApplication(ref xlApp);
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// 解除 Excel 工作表的保護（Unprotect）。
        /// 只針對第一個工作表，密碼為 <paramref name="password"/>，若為空則使用全域 <see cref="excelPassword"/>。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="password">解除保護密碼，若為空則使用全域 <see cref="excelPassword"/>。</param>
        public static void UnprotectWorksheet(string excelPath, string password)
        {
            // If caller passed empty password, fallback to global configured password
            if (string.IsNullOrEmpty(password)) password = excelPassword;
            Excel.Application xlApp = null; Excel.Workbook wb = null;
            try
            {
                xlApp = new Excel.Application(); xlApp.DisplayAlerts = false;
                wb = xlApp.Workbooks.Open(excelPath, ReadOnly: false, Password: (object)(excelPassword ?? string.Empty));
                Excel.Worksheet mainSheet = wb.Sheets[1] as Excel.Worksheet;
                try { if (!string.IsNullOrEmpty(password)) mainSheet.Unprotect(password); } catch { }
                wb.Save();
            }
            finally
            {
                TryReleaseWorkbook(ref wb);
                TryReleaseApplication(ref xlApp);
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// 檢查 Excel 工作表是否已被保護（Protect）。
        /// 只檢查第一個工作表，並以 <see cref="excelPassword"/> 嘗試開啟。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <returns>若已保護則回傳 true，否則回傳 false。</returns>
        public static bool IsWorksheetProtected(string excelPath)
        {
            Excel.Application xlApp = null; Excel.Workbook wb = null;
            try
            {
                xlApp = new Excel.Application(); xlApp.DisplayAlerts = false;
                wb = xlApp.Workbooks.Open(excelPath, ReadOnly: true, Password: (object)(excelPassword ?? string.Empty));
                Excel.Worksheet mainSheet = wb.Sheets[1] as Excel.Worksheet;
                bool prot = false;
                try { prot = mainSheet.ProtectContents; } catch { prot = false; }
                return prot;
            }
            finally
            {
                TryReleaseWorkbook(ref wb);
                TryReleaseApplication(ref xlApp);
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// 取得 Excel Range 的字串內容（Value2），若為 null 則回傳空字串。
        /// </summary>
        /// <param name="rng">Excel Range 物件。</param>
        /// <returns>儲存格內容的字串表示，若無內容則為空字串。</returns>
        private static string GetRangeString(Excel.Range rng)
        {
            if (rng == null) return string.Empty;
            try { var v = rng.Value2; if (v == null) return string.Empty; return v.ToString(); } catch { return string.Empty; }
        }

        /// <summary>
        /// 嘗試釋放 Excel Range COM 物件。
        /// 若 Range 尚未釋放，則釋放資源並將參考設為 null。
        /// </summary>
        /// <param name="rng">要釋放的 Range 參考，釋放後設為 null。</param>
        private static void TryReleaseRange(ref Excel.Range rng)
        {
            try { if (rng != null) { Marshal.ReleaseComObject(rng); rng = null; } } catch { rng = null; }
        }

        /// <summary>
        /// 嘗試釋放 Excel Worksheet COM 物件。
        /// 若 Worksheet 尚未釋放，則釋放資源並將參考設為 null。
        /// </summary>
        /// <param name="ws">要釋放的 Worksheet 參考，釋放後設為 null。</param>
        private static void TryReleaseWorksheet(ref Excel.Worksheet ws)
        {
            try { if (ws != null) { Marshal.ReleaseComObject(ws); ws = null; } } catch { ws = null; }
        }

        /// <summary>
        /// 嘗試釋放 Excel Workbook COM 物件。
        /// 若 Workbook 尚未關閉，則先關閉（不儲存變更），再釋放資源。
        /// </summary>
        /// <param name="wb">要釋放的 Workbook 參考，釋放後設為 null。</param>
        private static void TryReleaseWorkbook(ref Excel.Workbook wb)
        {
            try { if (wb != null) { wb.Close(false); Marshal.ReleaseComObject(wb); wb = null; } } catch { wb = null; }
        }

        /// <summary>
        /// 嘗試釋放 Excel Application COM 物件。
        /// 若 Application 尚未結束，則先結束，再釋放資源。
        /// </summary>
        /// <param name="app">要釋放的 Application 參考，釋放後設為 null。</param>
        private static void TryReleaseApplication(ref Excel.Application app)
        {
            try { if (app != null) { app.Quit(); Marshal.ReleaseComObject(app); app = null; } } catch { app = null; }
        }
    }
}
