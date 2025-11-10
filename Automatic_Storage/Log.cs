using System;

namespace Automatic_Storage
{
    /// <summary>
    /// 提供系統操作記錄功能的類別。
    /// </summary>
    class Log
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
        struct Action_Button
        {
            /// <summary>入庫按鈕。</summary>
            public string btn_Input;
            /// <summary>出庫按鈕。</summary>
            public string btn_Out;
            /// <summary>批次入庫按鈕。</summary>
            public string btn_BatIn;
            /// <summary>批次出庫按鈕。</summary>
            public string btn_BatOut;
            /// <summary>查詢全部按鈕。</summary>
            public string button1;
            /// <summary>料號+儲位合併查詢按鈕。</summary>
            public string btn_combi;
            /// <summary>儲位刪除按鈕。</summary>
            public string btn_delPosition;
            /// <summary>歷史_全部查詢按鈕。</summary>
            public string btn_findAll;
            /// <summary>歷史_料號+儲位合併查詢按鈕。</summary>
            public string btn_itemSite;
            /// <summary>歷史_返回按鈕。</summary>
            public string btn_reP2;
            /// <summary>檔案選擇按鈕。</summary>
            public string selectButton;
            /// <summary>檔案上傳按鈕。</summary>
            public string commitButton;
            /// <summary>檔案_介面返回按鈕。</summary>
            public string btn_return;
            /// <summary>維護記錄按鈕。</summary>
            public string maintainRecord;
            /// <summary>歷史記錄按鈕。</summary>
            public string historyRecord;
        }

        /// <summary>
        /// 包含所有操作文字框的名稱。
        /// </summary>
        struct Action_TextBox
        {
            /// <summary>料號記錄文字框。</summary>
            public string itemRecord;
            /// <summary>儲位記錄文字框。</summary>
            public string positionRecord;
        }

        /// <summary>
        /// 寫入操作記錄。
        /// </summary>
        /// <param name="EventName">事件名稱。</param>
        public void WriteLog(string EventName)
        {
            try
            {
                string folder = AppDomain.CurrentDomain.BaseDirectory;
                string path = System.IO.Path.Combine(folder, "Automatic_Storage.log");
                string line = string.Format("[{0:yyyy-MM-dd HH:mm:ss}] {1}", DateTime.Now, EventName);
                // 以集中式 logger 處理日誌，保留原先的同步 API 行為且不會丟例外
                try
                {
                    // 使用集中式非同步 logger，採同步等待結果以維持 WriteLog 的同步契約
                    Automatic_Storage.Utilities.Logger.LogInfoAsync(line).GetAwaiter().GetResult();
                }
                catch
                {
                    // logging must not throw - 最終保留沉默處理以免影響業務流程
                }
            }
            catch
            {
                // logging must not throw
            }
        }
    }
}
