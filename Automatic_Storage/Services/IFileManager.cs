using System.IO;
using System.Threading.Tasks;

namespace Automatic_Storage.Services
{
    /// <summary>
    /// 提供檔案管理相關職責的介面，例如備份、還原及開啟獨佔檔案。
    /// </summary>
    public interface IFileManager
    {
        /// <summary>
        /// 建立指定檔案的備份。
        /// </summary>
        /// <param name="path">要備份的原始檔案路徑。</param>
        /// <returns>回傳備份檔案的完整路徑。</returns>
        Task<string> CreateBackupAsync(string path);

        /// <summary>
        /// 從備份檔案還原至原始檔案位置。
        /// </summary>
        /// <param name="originalPath">原始檔案的路徑。</param>
        /// <param name="backupPath">備份檔案的路徑。</param>
        /// <returns>非同步作業。</returns>
        Task RestoreBackupAsync(string originalPath, string backupPath);

        /// <summary>
        /// 嘗試以獨佔模式開啟指定檔案的 FileStream。此同步方法仍保留以相容舊有實作，
        /// 建議使用非同步版本 <see cref="TryOpenExclusiveFileStreamAsync"/>。
        /// </summary>
        /// <param name="path">要開啟的檔案路徑。</param>
        /// <returns>成功時回傳 FileStream，否則回傳 null。</returns>
        [System.Obsolete("Use TryOpenExclusiveFileStreamAsync instead", false)]
        FileStream TryOpenExclusiveFileStream(string path);

        /// <summary>
        /// 非同步嘗試以獨佔模式開啟指定檔案的 FileStream。
        /// 新增此方法以便逐步移轉至 async-only API。
        /// </summary>
        /// <param name="path">要開啟的檔案路徑。</param>
        /// <returns>成功時回傳 FileStream，否則回傳 null。</returns>
        Task<FileStream> TryOpenExclusiveFileStreamAsync(string path);
    }
}
