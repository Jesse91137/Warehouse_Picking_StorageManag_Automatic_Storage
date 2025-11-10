using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Automatic_Storage.Utilities
{
    /// <summary>
    /// 簡單非同步檔案日誌工具，使用 <see cref="SemaphoreSlim"/> 保護多執行緒寫入。
    /// 用於記錄匯入、回寫、錯誤等事件，提升可維護性與可觀察性。
    /// </summary>
    public static class Logger
    {
        /// <summary>
        /// 控制同時存取日誌檔案的同步機制。
        /// </summary>
        private static readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);

        /// <summary>
        /// 日誌檔案儲存的資料夾路徑。
        /// </summary>
        private static readonly string _logFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");

        /// <summary>
        /// 類別靜態建構式，確保日誌資料夾存在。
        /// </summary>
        static Logger()
        {
            try
            {
                if (!Directory.Exists(_logFolder)) Directory.CreateDirectory(_logFolder);
            }
            catch { }
        }

        /// <summary>
        /// 取得指定分類的日誌檔案完整路徑（依日期區分）。
        /// </summary>
        /// <param name="name">日誌分類名稱。</param>
        /// <returns>日誌檔案的完整路徑。</returns>
        private static string LogFilePath(string name)
        {
            var file = name + "_" + DateTime.Now.ToString("yyyyMMdd") + ".log";
            return Path.Combine(_logFolder, file);
        }

        /// <summary>
        /// 非同步寫入資訊等級日誌。
        /// </summary>
        /// <param name="message">要記錄的訊息內容。</param>
        /// <returns>非同步作業。</returns>
        public static async Task LogInfoAsync(string message)
        {
            await WriteAsync("INFO", message, "application");
        }

        /// <summary>
        /// 非同步寫入警告等級日誌。
        /// </summary>
        /// <param name="message">要記錄的訊息內容。</param>
        /// <returns>非同步作業。</returns>
        public static async Task LogWarningAsync(string message)
        {
            await WriteAsync("WARN", message, "application");
        }

        /// <summary>
        /// 非同步寫入錯誤等級日誌。
        /// </summary>
        /// <param name="message">要記錄的訊息內容。</param>
        /// <returns>非同步作業。</returns>
        public static async Task LogErrorAsync(string message)
        {
            await WriteAsync("ERROR", message, "application");
        }

        /// <summary>
        /// 非同步寫入日誌檔案，根據等級與分類自動分檔。
        /// </summary>
        /// <param name="level">日誌等級（INFO/WARN/ERROR）。</param>
        /// <param name="message">要記錄的訊息內容。</param>
        /// <param name="category">日誌分類。</param>
        /// <returns>非同步作業。</returns>
        private static async Task WriteAsync(string level, string message, string category)
        {
            string path = LogFilePath(category);
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {level} {message}" + Environment.NewLine;
            await _semaphore.WaitAsync().ConfigureAwait(false);
            try
            {
                // 嘗試以 Task.Run 方式非同步寫入，支援舊版 .NET Framework
                try
                {
                    await Task.Run(() => File.AppendAllText(path, line, Encoding.UTF8)).ConfigureAwait(false);
                    return;
                }
                catch
                {
                    // 若失敗則以 FileStream 同步寫入，確保檔案共用控制
                    try
                    {
                        var dir = Path.GetDirectoryName(path);
                        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
                        using (var fs = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read))
                        using (var sw = new StreamWriter(fs, Encoding.UTF8))
                        {
                            sw.Write(line);
                            sw.Flush();
                        }
                        return;
                    }
                    catch
                    {
                        // 最後備援：寫入 Trace，避免例外拋出並保留診斷資訊
                        try { System.Diagnostics.Trace.WriteLine(line); } catch { }
                    }
                }
            }
            finally
            {
                _semaphore.Release();
            }
        }
    }
}
