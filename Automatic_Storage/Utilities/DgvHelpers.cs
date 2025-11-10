using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Automatic_Storage.Utilities
{
    /// <summary>
    /// 共用的 DataGridView helper
    /// 提供延遲且安全的方式在 UI 執行緒中隱藏選取視覺效果
    /// </summary>
    public static class DgvHelpers
    {
        /// <summary>
        /// 延遲在 UI 執行緒中把 DataGridView 的選取視覺隱藏，並嘗試把 CurrentCell 設為 null。
        /// - 若提供 owner 且 handle 已建立，會使用 owner.BeginInvoke
        /// - 否則嘗試使用 dgv.BeginInvoke
        /// - 若都不可用，會以非同步背景任務忽略（無安全 UI 動作）
        /// </summary>
        public static void HideSelectionInGrid(Control owner, DataGridView dgv)
        {
            try
            {
                if (dgv == null) return;

                Action act = () =>
                {
                    try
                    {
                        var defaultBack = dgv.DefaultCellStyle?.BackColor ?? SystemColors.Window;
                        var defaultFore = dgv.DefaultCellStyle?.ForeColor ?? SystemColors.ControlText;

                        try
                        {
                            dgv.DefaultCellStyle.SelectionBackColor = defaultBack;
                            dgv.DefaultCellStyle.SelectionForeColor = defaultFore;
                        }
                        catch { }

                        try
                        {
                            foreach (DataGridViewColumn col in dgv.Columns.Cast<DataGridViewColumn>())
                            {
                                try
                                {
                                    var s = col.DefaultCellStyle;
                                    s.SelectionBackColor = defaultBack;
                                    s.SelectionForeColor = defaultFore;
                                    col.DefaultCellStyle = s;
                                }
                                catch { }
                            }
                        }
                        catch { }

                        try { if (dgv.CurrentCell != null) dgv.CurrentCell = null; } catch { }
                        try { dgv.ClearSelection(); } catch { }
                    }
                    catch { }
                };

                if (owner != null && owner.IsHandleCreated)
                {
                    try { owner.BeginInvoke(act); return; } catch { }
                }

                if (dgv.IsHandleCreated)
                {
                    try { dgv.BeginInvoke(act); return; } catch { }
                }

                // fallback: run asynchronously but cannot touch UI safely
                Task.Run(() => { try { /* no-op fallback */ } catch { } });
            }
            catch { }
        }
    }
}
