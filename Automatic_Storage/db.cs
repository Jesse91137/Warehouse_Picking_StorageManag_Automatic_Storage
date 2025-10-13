using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Automatic_Storage
{
    class db
    {
        #region paramater method
        //本機測試
        private static readonly String connStr = "server=192.168.6.57;database=Automatic_Storage_M;uid=sa;pwd=A12345678;Connect Timeout = 10";
        //正式環境
        //private static readonly String connStr = "server=192.168.4.120;database=Automatic_Storage;uid=Auto_sa;pwd=A12345678;Connect Timeout = 10";
        //ConfigurationManager.ConnectionStrings["conString"].ConnectionString;    
        //1. 執行insert/update/delete，回傳影響的資料列數
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

        public static SqlDataReader ExecuteReader(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            SqlConnection con = new SqlConnection(connStr);
            using (SqlCommand cmd = new SqlCommand(sql, con))
            {
                cmd.CommandType = cmdType;
                if (pms != null)
                {
                    cmd.Parameters.AddRange(pms);
                }
                try
                {
                    con.Open();
                    return cmd.ExecuteReader(CommandBehavior.CloseConnection);
                }
                catch
                {
                    con.Close();
                    con.Dispose();
                    throw;
                }
            }
        }

        public static DataTable ExecuteDataTable(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            DataTable dt = new DataTable();
            //use SqlDataAdapter ,it will establish Sql connection.So ,it no need to create Connection by yourself.
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connStr))
            {
                adapter.SelectCommand.CommandType = cmdType;
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms);

                }
                adapter.Fill(dt);
                return dt;
            }
        }
        public static DataSet ExecuteDataSet(string sql, CommandType cmdType, params SqlParameter[] pms)
        {
            DataSet ds = new DataSet();
            //use SqlDataAdapter ,it will establish Sql connection.So ,it no need to create Connection by yourself.
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connStr))
            {
                adapter.SelectCommand.CommandType = cmdType;
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms);

                }
                adapter.Fill(ds);
                return ds;
            }
        }
        public static DataSet ExecuteDataSetPmsList(string sql, CommandType cmdType, List<SqlParameter> pms)
        {
            DataSet ds = new DataSet();
            //use SqlDataAdapter ,it will establish Sql connection.So ,it no need to create Connection by yourself.
            using (SqlDataAdapter adapter = new SqlDataAdapter(sql, connStr))
            {
                adapter.SelectCommand.CommandType = cmdType;
                if (pms != null)
                {
                    adapter.SelectCommand.Parameters.AddRange(pms.ToArray<SqlParameter>());//paralist.ToArray<SqlParameter>()

                }
                adapter.Fill(ds);
                return ds;
            }
        }
        #endregion

        #region Beer method
        public static SqlConnection GetCon()
        {
            //系統更新使用↓
            string cnstr = "server=192.168.4.120;database=Automatic_Storage;uid=Auto_sa;pwd=A12345678;Connect Timeout = 10";

            SqlConnection icn = new SqlConnection();
            icn.ConnectionString = cnstr;
            if (icn.State == ConnectionState.Open) icn.Close();
            icn.Open();
            return icn;
        }
        public static bool Exsql(string cmdtxt)
        {
            SqlConnection con = db.GetCon();//連接資料庫
            //con.Open();
            SqlCommand cmd = new SqlCommand(cmdtxt, con);
            try
            {
                cmd.ExecuteNonQuery();//執行SQL 語句並返回受影響的行數
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
            finally
            {
                con.Dispose();//釋放連接物件資源
                con.Close();
            }
        }
        public static DataSet reDs(string cmdtxt)
        {
            SqlConnection con = db.GetCon();
            SqlDataAdapter da = new SqlDataAdapter(cmdtxt, con);
            //建立資料集ds
            DataSet ds = new DataSet();
            da.Fill(ds);
            return ds;
        }
        #endregion
    }
}
