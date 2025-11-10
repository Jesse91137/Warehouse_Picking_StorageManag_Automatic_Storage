using Automatic_Storage.Utilities;
using System;
using System.Configuration;
using System.Data.SqlClient;

namespace Automatic_Storage.Services
{
    /// <summary>
    /// Excel 密碼提供者，負責從設定或資料庫取得密碼並做快取以降低查詢頻率。
    /// 這個實作會在無法取得時回退到 ExcelInteropHelper.excelPassword 預設值。
    /// </summary>
    public class ExcelPasswordProvider : IExcelPasswordProvider
    {
        /// <summary>
        /// 密碼快取鎖定物件，確保多執行緒安全。
        /// </summary>
        private readonly object _pwdLock = new object();

        /// <summary>
        /// 快取的密碼字串。
        /// </summary>
        private string _cachedPassword = null;

        /// <summary>
        /// 快取密碼的時間戳記。
        /// </summary>
        private DateTime _cachedAt = DateTime.MinValue;

        /// <summary>
        /// 密碼快取的有效時間，預設為 10 分鐘。
        /// </summary>
        private readonly TimeSpan _defaultCache = TimeSpan.FromMinutes(10);

        /// <summary>
        /// 密碼查詢失敗時的預設回退密碼。
        /// </summary>
        private readonly string _fallback = ExcelInteropHelper.excelPassword ?? "1234";

        /// <summary>
        /// 取得 Excel 密碼，優先從快取取得，否則從資料庫查詢，失敗則回退預設值。
        /// </summary>
        /// <returns>Excel 密碼字串。</returns>
        public string GetPassword()
        {
            lock (_pwdLock)
            {
                try
                {
                    // 檢查快取是否有效
                    if (!string.IsNullOrEmpty(_cachedPassword) && (DateTime.UtcNow - _cachedAt) < _defaultCache)
                        return _cachedPassword;

                    string connName = "DefaultConnection";
                    var cs = ConfigurationManager.ConnectionStrings[connName]?.ConnectionString;
                    if (string.IsNullOrWhiteSpace(cs))
                    {
                        _cachedPassword = _fallback;
                        _cachedAt = DateTime.UtcNow;
                        return _cachedPassword;
                    }

                    try
                    {
                        using (var conn = new SqlConnection(cs))
                        {
                            conn.Open();
                            string sql = "SELECT TOP 1 Ftp_Password FROM [dbo].[FtpServer_Table] WHERE Ftp_Server_Ip=N'檢料系統' AND Ftp_Server_OA_Ip=N'倉庫' AND Ftp_Username=@username AND Ftp_Server_name=N'儲位管理系統'";
                            using (var cmd = new SqlCommand(sql, conn))
                            {
                                cmd.Parameters.AddWithValue("@username", "Automatic_Storage");
                                var result = cmd.ExecuteScalar();
                                if (result != null && result != DBNull.Value)
                                    _cachedPassword = result.ToString();
                                else
                                    _cachedPassword = _fallback;
                                _cachedAt = DateTime.UtcNow;
                                return _cachedPassword;
                            }
                        }
                    }
                    catch
                    {
                        _cachedPassword = _fallback;
                        _cachedAt = DateTime.UtcNow;
                        return _cachedPassword;
                    }
                }
                catch
                {
                    _cachedPassword = _fallback;
                    _cachedAt = DateTime.UtcNow;
                    return _cachedPassword;
                }
            }
        }
    }
}
