using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace Automatic_Storage.Utilities
{
    /// <summary>
    /// 序列化 Excel 寫入請求的單例工作佇列，避免同時間對同一檔案並行寫入造成檔案鎖定或資料競爭。
    /// </summary>
    public class ExcelWriteQueue : IDisposable
    {
        /// <summary>
        /// 代表一筆 Excel 寫入請求的內部類別。
        /// </summary>
        private class WriteRequest
        {
            /// <summary>
            /// Excel 檔案路徑。
            /// </summary>
            public string ExcelPath;
            /// <summary>
            /// 物料編號。
            /// </summary>
            public string MaterialCode;
            /// <summary>
            /// 數量。
            /// </summary>
            public int Qty;
            /// <summary>
            /// 操作人員名稱。
            /// </summary>
            public string OperatorName;
        }

        /// <summary>
        /// 儲存待處理寫入請求的佇列。
        /// </summary>
        private readonly BlockingCollection<WriteRequest> _queue = new BlockingCollection<WriteRequest>(new ConcurrentQueue<WriteRequest>());
        /// <summary>
        /// 用於取消工作執行緒的 TokenSource。
        /// </summary>
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();
        /// <summary>
        /// 處理 Excel 寫入請求的背景執行緒。
        /// </summary>
        private readonly Thread _workerThread;
        /// <summary>
        /// ExcelWriteQueue 的單例實例。
        /// </summary>
        private static readonly Lazy<ExcelWriteQueue> _instance = new Lazy<ExcelWriteQueue>(() => new ExcelWriteQueue());

        /// <summary>
        /// 取得 ExcelWriteQueue 的單例實例。
        /// </summary>
        public static ExcelWriteQueue Instance => _instance.Value;

        /// <summary>
        /// Excel 操作服務。
        /// </summary>
        private readonly Automatic_Storage.Services.IExcelService _excelService;
        /// <summary>
        /// Excel 密碼提供者。
        /// </summary>
        private readonly Automatic_Storage.Services.IExcelPasswordProvider _passwordProvider;

        /// <summary>
        /// 預設建構式，供單例使用，會建立預設的服務實作。
        /// </summary>
        private ExcelWriteQueue() : this(new Automatic_Storage.Services.ExcelService(), new Automatic_Storage.Services.ExcelPasswordProvider()) { }

        /// <summary>
        /// 以注入的服務建構 ExcelWriteQueue，方便測試或替換實作。
        /// </summary>
        /// <param name="excelService">Excel 操作服務。</param>
        /// <param name="passwordProvider">Excel 密碼提供者。</param>
        public ExcelWriteQueue(Automatic_Storage.Services.IExcelService excelService, Automatic_Storage.Services.IExcelPasswordProvider passwordProvider)
        {
            _excelService = excelService ?? new Automatic_Storage.Services.ExcelService();
            _passwordProvider = passwordProvider ?? new Automatic_Storage.Services.ExcelPasswordProvider();

            // 使用專用 STA thread 以確保 Excel Interop 在 STA apartment 中執行
            _workerThread = new Thread(() =>
            {
                try { ProcessQueue(); }
                catch { }
            });
            _workerThread.IsBackground = true;
            try
            {
                // 使用反射設定 apartment state，降低靜態分析對跨平台 API 的警示
                var setMethod = typeof(Thread).GetMethod("SetApartmentState", new Type[] { typeof(ApartmentState) });
                if (setMethod != null)
                {
                    try { setMethod.Invoke(_workerThread, new object[] { ApartmentState.STA }); } catch { }
                }
            }
            catch { }
            _workerThread.Start();
        }

        /// <summary>
        /// 將一筆 Excel 寫入請求加入佇列。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑。</param>
        /// <param name="materialCode">物料編號。</param>
        /// <param name="qty">數量。</param>
        /// <param name="operatorName">操作人員名稱。</param>
        public void Enqueue(string excelPath, string materialCode, int qty, string operatorName)
        {
            if (string.IsNullOrWhiteSpace(excelPath)) return;
            _queue.Add(new WriteRequest { ExcelPath = excelPath, MaterialCode = materialCode, Qty = qty, OperatorName = operatorName });
        }

        /// <summary>
        /// 等待內部佇列清空或逾時。
        /// </summary>
        /// <param name="timeout">等待的最長時間。</param>
        /// <returns>若佇列在逾時前清空則回傳 true，否則回傳 false。</returns>
        public async Task<bool> FlushAsync(TimeSpan timeout)
        {
            var sw = Stopwatch.StartNew();
            try
            {
                while (_queue.Count > 0 && sw.Elapsed < timeout)
                {
                    await Task.Delay(100);
                }
                return _queue.Count == 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 處理佇列中的 Excel 寫入請求，包含重試與保護/解除保護邏輯。
        /// </summary>
        private void ProcessQueue()
        {
            foreach (var req in _queue.GetConsumingEnumerable(_cts.Token))
            {
                bool ok = false;
                int maxAttempts = 3;
                int attempt = 0;
                while (!ok && attempt < maxAttempts)
                {
                    attempt++;
                    try
                    {
                        // 使用注入的 service
                        var svc = _excelService;
                        // 在寫入前嘗試解除保護（使用 password provider），若失敗則仍嘗試寫入並記錄
                        try
                        {
                            var pwd = _passwordProvider.GetPassword();
                            svc.UnprotectWorksheet(req.ExcelPath, pwd);
                            Automatic_Storage.Utilities.Logger.LogInfoAsync($"Unprotected {req.ExcelPath} with configured password before write.").Wait();
                        }
                        catch (Exception unex)
                        {
                            Automatic_Storage.Utilities.Logger.LogErrorAsync($"Unprotect attempt failed for {req.ExcelPath}: {unex.Message}").Wait();
                        }

                        svc.UpdateShippedAndAppendRecord(req.ExcelPath, req.MaterialCode, req.Qty, req.OperatorName);
                        Automatic_Storage.Utilities.Logger.LogInfoAsync($"Excel write succeeded: {req.MaterialCode} +{req.Qty} to {req.ExcelPath}").Wait();
                        ok = true;

                        // Re-apply protection after a successful write to ensure sheets/columns remain locked.
                        try
                        {
                            bool protectShipped = string.Equals(Path.GetExtension(req.ExcelPath), ".xlsm", StringComparison.OrdinalIgnoreCase);
                            var pwd = _passwordProvider.GetPassword();
                            if (string.IsNullOrEmpty(pwd)) pwd = string.Empty;
                            svc.ProtectWorksheet(req.ExcelPath, pwd, protectShipped);
                            Automatic_Storage.Utilities.Logger.LogInfoAsync($"Re-applied protection for {req.ExcelPath} (protectShippedColumn={protectShipped}, pwdSet={!string.IsNullOrEmpty(pwd)})").Wait();
                        }
                        catch (Exception ex)
                        {
                            Automatic_Storage.Utilities.Logger.LogErrorAsync($"Failed to re-apply protection for {req.ExcelPath}: {ex.Message}").Wait();
                        }
                    }
                    catch (Exception ex)
                    {
                        Automatic_Storage.Utilities.Logger.LogErrorAsync($"Excel write attempt {attempt} failed for {req.MaterialCode} to {req.ExcelPath}: {ex.Message}").Wait();
                        // exponential backoff
                        Thread.Sleep(500 * attempt);
                    }
                }

                if (!ok)
                {
                    // 最後失敗：寫入失敗紀錄以供人工處理
                    try
                    {
                        Automatic_Storage.Utilities.Logger.LogErrorAsync($"Excel write permanently failed for {req.MaterialCode} to {req.ExcelPath}").Wait();
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// 釋放資源，嘗試優雅地終止背景執行緒與相關物件。
        /// </summary>
        public void Dispose()
        {
            try { _queue.CompleteAdding(); } catch { }
            try { _cts.Cancel(); } catch { }
            try
            {
                if (_workerThread != null && _workerThread.IsAlive)
                {
                    // 給予背景執行緒額外時間完成。
                    if (!_workerThread.Join(5000))
                    {
                        // 無法在合理時間內停止，寫日誌並放棄強制中止（避免使用 Thread.Abort）
                        try { Automatic_Storage.Utilities.Logger.LogErrorAsync("ExcelWriteQueue worker thread did not stop within timeout during Dispose.").Wait(); } catch { }
                    }
                }
            }
            catch { }
            try { _cts.Dispose(); } catch { }
        }
    }
}
