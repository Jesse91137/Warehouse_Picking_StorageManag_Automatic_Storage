using System.Windows.Forms;

namespace Automatic_Storage.Services
{
    /// <summary>
    /// 實作: WinForms 使用者通知器。
    /// 使用 MessageBox 顯示通知。可在建構子中注入 parent form 以改善 UI 階層。
    /// </summary>
    public class WinFormsUserNotifier : IUserNotifier
    {
        /// <summary>
        /// 父表單參考，用於作為 MessageBox 的擁有者。
        /// </summary>
        private readonly Form? _parentForm;

        /// <summary>
        /// 建立 WinFormsUserNotifier 實例。
        /// </summary>
        /// <param name="parentForm">可選的父表單，作為 MessageBox 的擁有者。</param>
        public WinFormsUserNotifier(Form? parentForm = null)
        {
            _parentForm = parentForm;
        }

        /// <summary>
        /// 顯示資訊訊息（一般通知）。
        /// </summary>
        /// <param name="message">要顯示的訊息內容。</param>
        public void ShowInfo(string message)
        {
            try
            {
                if (_parentForm is not null && !_parentForm.Disposing)
                {
                    try { UiHelpers.EnsureCursorRestored(); } catch { }
                    try { MessageBox.Show(_parentForm, message, "資訊", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                    catch { try { UiHelpers.EnsureCursorRestored(); } catch { } MessageBox.Show(message, "資訊", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                }
                else
                {
                    MessageBox.Show(message, "資訊", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch { }
        }

        /// <summary>
        /// 顯示警告訊息。
        /// </summary>
        /// <param name="message">要顯示的訊息內容。</param>
        public void ShowWarning(string message)
        {
            try
            {
                if (_parentForm is not null && !_parentForm.Disposing)
                {
                    try { UiHelpers.EnsureCursorRestored(); } catch { }
                    try { MessageBox.Show(_parentForm, message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                    catch { try { UiHelpers.EnsureCursorRestored(); } catch { } MessageBox.Show(message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
                }
                else
                {
                    MessageBox.Show(message, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch { }
        }

        /// <summary>
        /// 顯示錯誤訊息。
        /// </summary>
        /// <param name="message">要顯示的訊息內容。</param>
        public void ShowError(string message)
        {
            try
            {
                if (_parentForm is not null && !_parentForm.Disposing)
                {
                    try { UiHelpers.EnsureCursorRestored(); } catch { }
                    try { MessageBox.Show(_parentForm, message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                    catch { try { UiHelpers.EnsureCursorRestored(); } catch { } MessageBox.Show(message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                }
                else
                {
                    MessageBox.Show(message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch { }
        }
    }
}
