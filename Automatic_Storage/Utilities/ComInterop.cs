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
                try
                {
                    if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                    {
                        try { while (Marshal.ReleaseComObject(comObj) > 0) { } } catch { }
                    }
                }
                catch { }
                finally
                {
                    comObj = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }
            catch { }
        }
    }
}
