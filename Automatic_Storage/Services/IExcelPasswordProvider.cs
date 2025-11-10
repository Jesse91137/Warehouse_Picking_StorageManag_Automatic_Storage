namespace Automatic_Storage.Services
{
    /// <summary>
    /// 提供 Excel 保護或取消保護操作所需的密碼。
    /// 實作類別可從安全儲存區、組態或全域輔助工具讀取密碼。
    /// </summary>
    public interface IExcelPasswordProvider
    {
        /// <summary>
        /// 取得用於 Excel 保護或取消保護的密碼。
        /// </summary>
        /// <returns>Excel 密碼字串。</returns>
        string GetPassword();
    }
}
