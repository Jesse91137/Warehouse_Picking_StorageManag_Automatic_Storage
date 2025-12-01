using System;

namespace Automatic_Storage
{
    /// <summary>
    /// 提供系統操作記錄功能的類別。
    /// </summary>
    sealed class Log
    {
        /// <summary>
        /// 定義資料庫操作的類型。
        /// </summary>
        public enum CRUD
        {
            /// <summary>
            /// 新增資料。
            /// </summary>
            Insert = 0,
            /// <summary>
            /// 查詢資料。
            /// </summary>
            Select = 1,
            /// <summary>
            /// 更新資料。
            /// </summary>
            Update = 2,
            /// <summary>
            /// 刪除資料。
            /// </summary>
            Delete = 3
        }

        /// <summary>
        /// 包含所有操作按鈕的名稱。
        /// </summary>
        class Action_Button
        {
            /// <summary>入庫按鈕。</summary>
            public string btn_Input = string.Empty;
            /// <summary>出庫按鈕。</summary>
            public string btn_Out = string.Empty;
            /// <summary>批次入庫按鈕。</summary>
            public string btn_BatIn = string.Empty;
            /// <summary>批次出庫按鈕。</summary>
            public string btn_BatOut = string.Empty;
            /// <summary>查詢全部按鈕。</summary>
            public string button1 = string.Empty;
            /// <summary>料號+儲位合併查詢按鈕。</summary>
            public string btn_combi = string.Empty;
            /// <summary>儲位刪除按鈕。</summary>
            public string btn_delPosition = string.Empty;
            /// <summary>歷史_全部查詢按鈕。</summary>
            public string btn_findAll = string.Empty;
            /// <summary>歷史_料號+儲位合併查詢按鈕。</summary>
            public string btn_itemSite = string.Empty;
            /// <summary>歷史_返回按鈕。</summary>
            public string btn_reP2 = string.Empty;
            /// <summary>檔案選擇按鈕。</summary>
            public string selectButton = string.Empty;
            /// <summary>檔案上傳按鈕。</summary>
            public string commitButton = string.Empty;
            /// <summary>檔案_介面返回按鈕。</summary>
            public string btn_return = string.Empty;
            /// <summary>維護記錄按鈕。</summary>
            public string maintainRecord = string.Empty;
            /// <summary>歷史記錄按鈕。</summary>
            public string historyRecord = string.Empty;
        }

        /// <summary>
        /// 包含所有操作文字框的名稱。
        /// </summary>
        class Action_TextBox
        {
            /// <summary>料號記錄文字框。</summary>
            public string itemRecord = string.Empty;
            /// <summary>儲位記錄文字框。</summary>
            public string positionRecord = string.Empty;
        }

        /// <summary>
        /// 寫入操作記錄。
        /// </summary>
        /// <param name="EventName">事件名稱。</param>
        public void WriteLog(string EventName)
        {
            try
            {
                string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {EventName}";
                // 啟動非同步記錄但不等待（fire-and-forget），且不讓例外冒泡回呼叫端
                _ = Automatic_Storage.Utilities.Logger.LogInfoAsync(line);
            }
            catch
            {
                // logging must not throw
            }
        }
    }
}
