using System.Windows.Forms;

namespace Automatic_Storage
{
    /// <summary>
    /// UI 輔助工具類別（公開） - 提供全域的游標還原方法。
    /// </summary>
    public static class UiHelpers
    {
        /// <summary>
        /// 強制還原所有游標狀態（包括等待游標及所有 Form/Control 的游標）。
        /// 可安全地在任何執行緒或情境下呼叫此方法。
        /// </summary>
        /// <remarks>
        /// 此方法會嘗試將 <see cref="Application.UseWaitCursor"/> 設為 false，
        /// 並將 <see cref="Cursor.Current"/> 及所有開啟中的 <see cref="Form"/> 和其 <see cref="Control"/> 的游標設為 <see cref="Cursors.Default"/>。
        /// 最後會呼叫 <see cref="Cursor.Hide"/> 及 <see cref="Cursor.Show"/> 以確保游標狀態正確。
        /// 所有操作皆包裹於 try-catch 以避免例外中斷流程。
        /// </remarks>
        public static void EnsureCursorRestored()
        {
            try { Application.UseWaitCursor = false; } catch { }
            try { Cursor.Current = Cursors.Default; } catch { }
            try
            {
                foreach (Form f in Application.OpenForms)
                {
                    try { if (f != null) f.Cursor = Cursors.Default; } catch { }
                    try
                    {
                        foreach (Control c in f.Controls)
                        {
                            try { if (c != null) c.Cursor = Cursors.Default; } catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            try { Cursor.Hide(); Cursor.Show(); } catch { }
        }
    }
}
