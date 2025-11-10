using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Windows.Forms;

namespace Automatic_Storage
{
    class db
    {
        #region paramater method
        //本機測試
        //private static readonly String connStr = "server=192.168.6.57;database=Automatic_Storage_M;uid=sa;pwd=A12345678;Connect Timeout = 10";
        //正式環境
        private static readonly String connStr = "server=192.168.4.120;database=Automatic_Storage;uid=Auto_sa;pwd=A12345678;Connect Timeout = 10";
        //ConfigurationManager.ConnectionStrings["conString"].ConnectionString;
        //1. 執行insert/update/delete，回傳影響的資料列數


        /// <summary>
        /// 執行 SQL 指令 (Insert/Update/Delete)，回傳受影響的資料列數。
        /// </summary>
        /// <param name="sql">要執行的 SQL 指令或儲存過程名稱。</param>
        /// <param name="cmdType">指令型態 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>受影響的資料列數。</returns>
        public static int ExecueNonQuery(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            using (SqlConnection con = new SqlConnection(connStr))
            {
                con.Open();
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    //設置目前執行的是「存儲過程? 還是帶參數的sql 語句?」
                    cmd.CommandType = cmdType;
                    if (pms != null)
                    {
                        cmd.Parameters.AddRange(pms);
                    }

                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 執行 SQL 指令 (Insert/Update/Delete)，並記錄執行的 T-SQL、參數及按鈕名稱至日誌資料表。
        /// </summary>
        /// <param name="sql">要執行的 SQL 指令或儲存過程名稱。</param>
        /// <param name="cmdType">指令型態 (Text 或 StoredProcedure)。</param>
        /// <param name="buttonName">執行此操作的按鈕名稱。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>受影響的資料列數。</returns>
        public static int ExecueNonQuery(string sql, CommandType cmdType, string buttonName, params SqlParameter[] pms)
        {
            using (SqlConnection con = new SqlConnection(connStr))
            {
                con.Open();

                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    cmd.CommandType = cmdType;

                    if (pms != null)
                    {
                        cmd.Parameters.AddRange(pms);
                    }

                    // 在這裡加入日誌記錄
                    LogSqlCommand(cmd, buttonName);

                    return cmd.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 將執行的 SQL 指令、參數及按鈕名稱記錄到 Automatic_Storage_CommandLog 資料表。
        /// </summary>
        /// <param name="cmd">SqlCommand 物件，包含執行的 T-SQL 及參數。</param>
        /// <param name="buttonName">執行此操作的按鈕名稱。</param>
        private static void LogSqlCommand(SqlCommand cmd, string buttonName)
        {
            // 將執行的 T-SQL 與參數信息記錄下來，你可以自行擴展日誌記錄的方式
            string logMessage = $"Executing T-SQL: {cmd.CommandText}";
            string Parameters = $"Parameters: {string.Join(", ", cmd.Parameters.Cast<SqlParameter>().Select(p => $"{p.ParameterName}={p.Value}"))}\n";
            string ButtonName = $"Button Name: {buttonName}";

            // 在這裡加入你的日誌輸出邏輯，可以使用檔案輸出、資料庫紀錄、日誌庫等方式
            Console.WriteLine(logMessage);
            try
            {
                string sqlcomm = "insert into Automatic_Storage_CommandLog" +
                "(TsqlCommandText,Parameters,ButtonName,Datetimes) " +
                "values(@logMessage,@Parameters,@ButtonName,@date) ";
                SqlParameter[] sqlParameter = new SqlParameter[]
                {
                new SqlParameter("logMessage",logMessage.Trim()),
                new SqlParameter("Parameters",Parameters),
                new SqlParameter("ButtonName",ButtonName),
                new SqlParameter("date",DateTime.Now)
                };
                ExecueNonQuery(sqlcomm, CommandType.Text, sqlParameter);
            }
            catch (Exception qrr)
            {
                string rrr = qrr.Message;
            }
        }

        /// <summary>
        /// 執行 SQL 指令 (查詢)，回傳 SqlDataReader 物件。
        /// </summary>
        /// <param name="sql">要執行的 SQL 指令或儲存過程名稱。</param>
        /// <param name="cmdType">指令型態 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>SqlDataReader，查詢結果資料集。</returns>
        public static SqlDataReader ExecuteReader(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            // 建立 SQL 連線物件
            SqlConnection con = new SqlConnection(connStr);
            // 使用 SqlCommand 執行 SQL 指令
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                // 設定指令型態 (Text 或 StoredProcedure)
                cmd.CommandType = cmdType;
                // 如果有參數，加入參數至 SqlCommand
                if (pms != null)
                {
                    cmd.Parameters.AddRange(pms);
                }
                try
                {
                    // 開啟資料庫連線
                    con.Open();
                    // 執行查詢，並回傳 SqlDataReader，查詢結束自動關閉連線
                    return cmd.ExecuteReader(CommandBehavior.CloseConnection);
                }
                catch
                {
                    // 發生例外時，關閉並釋放連線資源
                    con.Close();
                    con.Dispose();
                    // 丟出例外
                    throw;
                }
            }
        }

        /// <summary>
        /// 查詢資料庫並回傳 DataTable。
        /// </summary>
        /// <param name="sql">要執行的 SQL 指令或儲存過程名稱。</param>
        /// <param name="cmdType">指令型態 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>查詢結果的 DataTable。</returns>
        /// <remarks>
        /// 使用 SqlDataAdapter 會自動建立 SQL 連線，無需手動建立連線物件。
        /// </remarks>
        public static DataTable ExecuteDataTable(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            /// <summary>
            /// 建立一個新的 DataTable 物件，用來儲存查詢結果。
            /// </summary>
            DataTable dt = new DataTable();

            /// <summary>
            /// 使用 SqlDataAdapter 物件執行 SQL 查詢，並自動建立連線。
            /// </summary>
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connStr))
            {
                /// <summary>
                /// 設定查詢的指令型態 (Text 或 StoredProcedure)。
                /// </summary>
                adapter.SelectCommand.CommandType = cmdType;

                /// <summary>
                /// 如果有 SQL 參數，則加入參數至 SqlCommand。
                /// </summary>
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms);
                }

                /// <summary>
                /// 執行查詢並將結果填入 DataTable。
                /// </summary>
                adapter.Fill(dt);

                /// <summary>
                /// 回傳查詢結果的 DataTable。
                /// </summary>
                return dt;
            }
        }

        /// <summary>
        /// 使用 SqlDataAdapter 執行 SQL 查詢，並回傳查詢結果的 DataSet。
        /// </summary>
        /// <param name="sql">要執行的 SQL 指令或儲存過程名稱。</param>
        /// <param name="cmdType">指令型態 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數陣列。</param>
        /// <returns>查詢結果的 DataSet。</returns>
        /// <remarks>
        /// 使用 SqlDataAdapter 會自動建立 SQL 連線，無需手動建立連線物件。
        /// </remarks>
        public static DataSet ExecuteDataSet(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            DataSet ds = new DataSet();
            // 使用 SqlDataAdapter，自動建立 SQL 連線，無需手動建立連線物件。
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connStr))
            {
                // 設定查詢的指令型態 (Text 或 StoredProcedure)。
                adapter.SelectCommand.CommandType = cmdType;
                // 如果有 SQL 參數，則加入參數至 SqlCommand。
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms);
                }
                // 執行查詢並將結果填入 DataSet。
                adapter.Fill(ds);
                // 回傳查詢結果的 DataSet。
                return ds;
            }
        }

        /// <summary>
        /// 查詢資料庫並回傳 DataSet。
        /// </summary>
        /// <param name="sql">要執行的 SQL 指令或儲存過程名稱。</param>
        /// <param name="cmdType">指令型態 (Text 或 StoredProcedure)。</param>
        /// <param name="pms">SQL 參數的 List 集合。</param>
        /// <returns>查詢結果的 DataSet。</returns>
        /// <remarks>
        /// 使用 SqlDataAdapter 會自動建立 SQL 連線，無需手動建立連線物件。
        /// </remarks>
        public static DataSet ExecuteDataSetPmsList(string sql, CommandType cmdType, List<SqlParameter> pms)
        {
            // 建立一個新的 DataSet 物件，用來儲存查詢結果。
            DataSet ds = new DataSet();
            // 使用 SqlDataAdapter，自動建立 SQL 連線，無需手動建立連線物件。
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connStr))
            {
                // 設定查詢的指令型態 (Text 或 StoredProcedure)。
                adapter.SelectCommand.CommandType = cmdType;
                // 如果有 SQL 參數，則加入參數至 SqlCommand。
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms.ToArray<SqlParameter>());//paralist.ToArray<SqlParameter>()
                }
                // 執行查詢並將結果填入 DataSet。
                adapter.Fill(ds);
                // 回傳查詢結果的 DataSet。
                return ds;
            }
        }

        #endregion

        #region Beer method
        /// <summary>
        /// 取得資料庫連線
        /// </summary>
        /// <returns></returns>
        public static SqlConnection GetCon()
        {
            //系統更新使用↓
            string cnstr = connStr;

            SqlConnection icn = new SqlConnection();
            icn.ConnectionString = cnstr;
            if (icn.State == ConnectionState.Open) icn.Close();
            icn.Open();
            return icn;
        }

        /// <summary>
        /// 執行 SQL 指令 (Insert/Update/Delete)，並回傳是否執行成功。
        /// </summary>
        /// <param name="cmdtxt">要執行的 SQL 指令。</param>
        /// <returns>執行成功回傳 true，失敗回傳 false。</returns>
        public static bool Exsql(string cmdtxt)
        {
            // 連接資料庫
            SqlConnection con = db.GetCon();
            // 建立 SQL 指令物件
            SqlCommand cmd = new SqlCommand(cmdtxt, con);
            try
            {
                // 執行 SQL 語句並返回受影響的行數
                cmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception e)
            {
                // 顯示錯誤訊息
                MessageBox.Show(e.ToString());
                return false;
            }
            finally
            {
                // 釋放連接物件資源
                con.Dispose();
                // 關閉資料庫連線
                con.Close();
            }
        }

        /// <summary>
        /// 取得資料集。
        /// </summary>
        /// <param name="cmdtxt">要執行的 SQL 指令。</param>
        /// <returns>查詢結果的 DataSet。</returns>
        /// <remarks>
        /// 此方法會建立 SQL 連線，使用 SqlDataAdapter 執行查詢，並將結果填入 DataSet。
        /// </remarks>
        public static DataSet reDs(string cmdtxt)
        {
            // 取得資料庫連線
            SqlConnection con = db.GetCon();
            // 建立 SqlDataAdapter 物件，執行 SQL 查詢
            SqlDataAdapter da = new SqlDataAdapter(cmdtxt, con);
            // 建立資料集 DataSet
            DataSet ds = new DataSet();
            // 將查詢結果填入 DataSet
            da.Fill(ds);
            // 回傳查詢結果的 DataSet
            return ds;
        }

        #endregion
    }
}
