using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace Automatic_Storage.Utilities
{
    /// <summary>
    /// COM interop helpers (Release safely, STA helpers can be added here).
    /// </summary>
    public static class ComInterop
    {
        /// <summary>
        /// 安全釋放 COM 物件的工具方法。
        /// 會檢查物件是否為 COM 物件，若是則反覆呼叫 <see cref="Marshal.ReleaseComObject(object)"/> 直到參考計數為 0，
        /// 並於 Windows 平台下強制執行垃圾回收以確保資源釋放。
        /// 用於 Excel Interop 等 COM 物件釋放，避免記憶體洩漏。
        /// </summary>
        /// <param name="comObj">要釋放的 COM 物件，若為 null 或非 COM 物件則不執行任何動作。</param>
        public static void ReleaseComObjectSafe(object? comObj)
        {
            try
            {
                if (comObj == null) return;
                if (!Marshal.IsComObject(comObj)) return;
                // 在 Windows 上，儘量以 ReleaseComObject 反覆降低 RCW 的參考計數。
                // 若 ReleaseComObject 遇到例外或未能完全釋放，嘗試使用 FinalReleaseComObject 作為 fallback。
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    try
                    {
                        // 先嘗試以 ReleaseComObject 逐一釋放
                        try
                        {
                            while (Marshal.ReleaseComObject(comObj) > 0) { }
                        }
                        catch
                        {
                            // 若上面失敗，嘗試使用 FinalReleaseComObject
                            try { Marshal.FinalReleaseComObject(comObj); } catch { }
                        }
                    }
                    catch { }
                    finally
                    {
                        // 進行多輪 GC 與等待，增加 CLR 與 COM 釋放機會
                        try { comObj = null; } catch { }
                        try
                        {
                            for (int i = 0; i < 3; i++)
                            {
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                Thread.Sleep(50);
                            }
                        }
                        catch { }
                    }
                }
                else
                {
                    // 非 Windows 平台直接嘗試釋放並做 GC 保險措施
                    try { Marshal.FinalReleaseComObject(comObj); } catch { }
                    try { comObj = null; } catch { }
                    try { GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); } catch { }
                }
            }
            catch { }
        }
    }
}
