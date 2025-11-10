namespace Automatic_Storage.Services
{
    /// <summary>
    /// 介面: 使用者通知
    /// 抽象化使用者訊息顯示，允許在 UI 表單中使用 MessageBox，在單元測試中使用 Mock。
    /// </summary>
    public interface IUserNotifier
    {
        /// <summary>
        /// 顯示資訊訊息（一般通知）。
        /// </summary>
        /// <param name="message">訊息內容</param>
        void ShowInfo(string message);

        /// <summary>
        /// 顯示警告訊息。
        /// </summary>
        /// <param name="message">訊息內容</param>
        void ShowWarning(string message);

        /// <summary>
        /// 顯示錯誤訊息。
        /// </summary>
        /// <param name="message">訊息內容</param>
        void ShowError(string message);
    }
}
