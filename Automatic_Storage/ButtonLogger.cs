using System;
using System.Data;
using System.Data.SqlClient;

namespace Automatic_Storage
{
    /// <summary>
    /// 提供記錄按鈕事件至資料庫的功能。
    /// </summary>
    public class _Logger
    {
        /// <summary>
        /// 將按鈕事件記錄至 Automatic_Storage_ButtonLogs 資料表。
        /// </summary>
        /// <param name="enevt">事件名稱。</param>
        /// <param name="date">事件發生時間。</param>
        /// <param name="sqls">事件相關的 SQL 指令。</param>
        public void LogEvent(string enevt, DateTime date, string sqls)
        {
            // 定義插入按鈕事件記錄的 SQL 指令
            string strsql = @"INSERT INTO Automatic_Storage_ButtonLogs(EventName, LogTime,EventSQL) 
                                                                VALUES(@EventN, @LogTime, @Eventsql)";

            // 建立 SQL 參數陣列，包含事件名稱、時間及相關 SQL 指令
            SqlParameter[] sqlParameter = new SqlParameter[]
            {
                new SqlParameter("EventN", enevt), // 事件名稱參數
                new SqlParameter("LogTime", date), // 事件時間參數
                new SqlParameter("Eventsql", sqls) // 事件 SQL 指令參數
            };
            // 執行 SQL 指令，將事件記錄寫入資料庫
            db.ExecueNonQuery(strsql, CommandType.Text, sqlParameter);
        }
    }
}
