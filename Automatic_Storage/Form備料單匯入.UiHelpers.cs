using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Automatic_Storage
{
    public partial class Form備料單匯入
    {
        /// <summary>
        /// UI 輔助工具類別，提供批次停用/還原按鈕、顯示/隱藏遮罩、游標還原等功能。
        /// 僅供 <see cref="Form備料單匯入"/> 內部使用。
        /// </summary>
        private static class UiHelpers
        {
            /// <summary>
            /// 遞迴停用所有 ButtonBase 及 ToolStripItem，並記錄原本的 Enabled 狀態。
            /// 用於長時間作業時暫時鎖定 UI，避免重複操作。
            /// </summary>
            /// <param name="owner">目標表單實例。</param>
            public static void SaveAndDisableAllButtons(Form備料單匯入 owner)
            {
                owner._prevControlStates = new Dictionary<Control, bool>();
                owner._prevToolStripItemStates = new Dictionary<ToolStripItem, bool>();

                void DisableRecursiveLocal(Control parent)
                {
                    if (parent == null) return;
                    foreach (Control ctl in parent.Controls)
                    {
                        try
                        {
                            if (ctl is ButtonBase btn)
                            {
                                if (!owner._prevControlStates.ContainsKey(btn)) owner._prevControlStates[btn] = btn.Enabled;
                                btn.Enabled = false;
                            }
                            if (ctl.HasChildren) DisableRecursiveLocal(ctl);
                        }
                        catch { }
                    }
                }

                try { owner.SafeAction(() => DisableRecursiveLocal(owner)); } catch { }

                try
                {
                    void CollectToolStripsLocal(Control parent)
                    {
                        if (parent == null) return;
                        foreach (Control ctl in parent.Controls)
                        {
                            try
                            {
                                if (ctl is ToolStrip ts)
                                {
                                    foreach (ToolStripItem item in ts.Items)
                                    {
                                        try { if (!owner._prevToolStripItemStates.ContainsKey(item)) owner._prevToolStripItemStates[item] = item.Enabled; item.Enabled = false; } catch { }
                                    }
                                }
                                if (ctl.HasChildren) CollectToolStripsLocal(ctl);
                            }
                            catch { }
                        }
                    }
                    owner.SafeAction(() => CollectToolStripsLocal(owner));
                }
                catch { }

                try
                {
                    if (owner.Owner is Form ownerForm)
                    {
                        var matches = ownerForm.Controls.Find("btn備料單匯入", true);
                        if (matches != null && matches.Length > 0)
                        {
                            foreach (Control c in matches)
                            {
                                try { owner.SafeAction(() => { if (!owner._prevControlStates.ContainsKey(c)) owner._prevControlStates[c] = c.Enabled; c.Enabled = false; }); } catch { }
                            }
                        }
                    }
                }
                catch { }
            }

            /// <summary>
            /// 還原先前被 SaveAndDisableAllButtons 停用的所有按鈕與工具列項目。
            /// 若 _keepImportButtonDisabledUntilClose 為 true，則匯入按鈕維持停用。
            /// </summary>
            /// <param name="owner">目標表單實例。</param>
            public static void RestoreAllButtons(Form備料單匯入 owner)
            {
                try
                {
                    if (owner._prevControlStates != null)
                    {
                        foreach (var kv in owner._prevControlStates)
                        {
                            try
                            {
                                owner.SafeAction(() =>
                                {
                                    if (owner._keepImportButtonDisabledUntilClose && owner.btn備料單匯入檔案 != null && object.ReferenceEquals(kv.Key, owner.btn備料單匯入檔案))
                                    {
                                        try { kv.Key.Enabled = false; } catch { }
                                        return;
                                    }

                                    if (kv.Key != null) kv.Key.Enabled = kv.Value;
                                });
                            }
                            catch { }
                        }
                        owner._prevControlStates = null;
                    }

                    if (owner._prevToolStripItemStates != null)
                    {
                        foreach (var kv in owner._prevToolStripItemStates)
                        {
                            try { owner.SafeAction(() => { if (kv.Key != null) kv.Key.Enabled = kv.Value; }); } catch { }
                        }
                        owner._prevToolStripItemStates = null;
                    }
                }
                catch { }

                try
                {
                    if (!owner._keepImportButtonDisabledUntilClose)
                    {
                        foreach (Form f in Application.OpenForms)
                        {
                            try
                            {
                                var matches = f.Controls.Find("btn備料單匯入", true);
                                if (matches != null && matches.Length > 0)
                                {
                                    foreach (Control c in matches)
                                    {
                                        try { owner.SafeAction(() => owner.SetControlEnabledSafe(c, true)); } catch { }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }

            /// <summary>
            /// 在所有開啟的表單上顯示半透明遮罩，顯示處理中訊息並攔截輸入。
            /// 用於長時間作業時避免誤操作。
            /// </summary>
            /// <param name="owner">目標表單實例。</param>
            /// <param name="message">顯示的訊息內容。</param>
            public static void ShowOperationOverlay(Form備料單匯入 owner, string message)
            {
                try
                {
                    if (owner._operationOverlays != null && owner._operationOverlays.Count > 0) return;
                    owner._operationOverlays = new List<Panel>();

                    owner.SafeAction(() =>
                    {
                        foreach (Form f in Application.OpenForms)
                        {
                            try
                            {
                                owner.SafeAction(() =>
                                {
                                    if (f == null || !f.Visible) return;
                                    try { if (f.WindowState == FormWindowState.Minimized) return; } catch { }

                                    var panel = new Panel();
                                    panel.BackColor = Color.FromArgb(230, Color.White);
                                    panel.Size = new Size(Math.Min(400, Math.Max(240, f.ClientSize.Width - 80)), 80);
                                    panel.Visible = false;
                                    panel.Enabled = false;
                                    panel.TabStop = false;
                                    panel.Cursor = Cursors.WaitCursor;
                                    panel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

                                    var lbl = new Label();
                                    lbl.AutoSize = false;
                                    lbl.TextAlign = ContentAlignment.MiddleCenter;
                                    lbl.Dock = DockStyle.Fill;
                                    lbl.Font = new Font(SystemFonts.MessageBoxFont.FontFamily, 12f, FontStyle.Regular);
                                    lbl.Text = message;
                                    panel.Controls.Add(lbl);

                                    owner.SafeBeginInvoke(f, () => owner.SafeAction(() =>
                                    {
                                        f.Controls.Add(panel);
                                        panel.BringToFront();
                                        panel.Left = Math.Max(10, (f.ClientSize.Width - panel.Width) / 2);
                                        panel.Top = Math.Max(10, (f.ClientSize.Height - panel.Height) / 2);
                                        panel.Visible = true;
                                    }));

                                    owner._operationOverlays.Add(panel);
                                });
                            }
                            catch { }
                        }
                    });
                }
                catch { }
            }

            /// <summary>
            /// 隱藏並移除所有由 ShowOperationOverlay 建立的遮罩面板，釋放資源。
            /// </summary>
            /// <param name="owner">目標表單實例。</param>
            public static void HideOperationOverlay(Form備料單匯入 owner)
            {
                try
                {
                    if (owner._operationOverlays == null || owner._operationOverlays.Count == 0) return;
                    owner.SafeAction(() =>
                    {
                        foreach (var panel in owner._operationOverlays.ToList())
                        {
                            try
                            {
                                owner.SafeAction(() =>
                                {
                                    var parent = panel.Parent as Control;
                                    if (parent != null)
                                    {
                                        owner.SafeAction(() =>
                                        {
                                            panel.Cursor = Cursors.Default;
                                            foreach (Control cc in panel.Controls)
                                            {
                                                try { owner.SafeAction(() => { if (cc != null) cc.Cursor = Cursors.Default; }); } catch { }
                                            }
                                        });

                                        owner.SafeBeginInvoke(parent, () => owner.SafeAction(() => { parent.Controls.Remove(panel); panel.Dispose(); }));
                                    }
                                });
                            }
                            catch { }
                        }
                    });
                }
                catch { }
            }

            /// <summary>
            /// 強制還原所有游標狀態為 Default，包括 Application、所有 Form 及其子控制項。
            /// 用於確保長時間作業結束後游標不殘留等待狀態。
            /// </summary>
            /// <param name="owner">目標表單實例。</param>
            public static void ResetAllCursors(Form備料單匯入 owner)
            {
                try
                {
                    owner.SafeAction(() => Application.UseWaitCursor = false);
                    owner.SafeAction(() => Cursor.Current = Cursors.Default);
                    owner.SafeAction(() => owner.Cursor = Cursors.Default);
                    owner.SafeAction(() =>
                    {
                        foreach (Form f in Application.OpenForms)
                        {
                            try
                            {
                                owner.SafeAction(() => f.Cursor = Cursors.Default);
                                foreach (Control c in f.Controls)
                                {
                                    try { owner.SafeAction(() => { if (c != null) c.Cursor = Cursors.Default; }); } catch { }
                                }
                            }
                            catch { }
                        }
                    });
                    owner.SafeAction(() => Application.DoEvents());
                }
                catch { }
            }
        }
    }
}
