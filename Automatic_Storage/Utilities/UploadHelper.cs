using System;
using System.IO;

namespace Automatic_Storage.Utilities
{
    /// <summary>
    /// 提供 Upload 資料夾的建立與路徑管理功能。
    /// </summary>
    public static class UploadHelper
    {
        /// <summary>
        /// 取得 Upload 資料夾完整路徑，若不存在則自動建立。
        /// </summary>
        /// <returns>Upload 資料夾完整路徑</returns>
        public static string GetUploadFolder()
        {
            string uploadDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Upload");
            if (!Directory.Exists(uploadDir))
            {
                Directory.CreateDirectory(uploadDir);
            }
            return uploadDir;
        }

        /// <summary>
        /// 取得指定檔案在 Upload 資料夾下的完整路徑，並確保資料夾存在。
        /// </summary>
        /// <param name="fileName">檔案名稱</param>
        /// <returns>完整路徑</returns>
        public static string GetUploadFilePath(string fileName)
        {
            string folder = GetUploadFolder();
            return Path.Combine(folder, fileName);
        }
    }
}