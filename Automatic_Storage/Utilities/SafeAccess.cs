using System;

namespace Automatic_Storage.Utilities
{

    /// <summary>
    /// 提供安全存取物件並轉換為常用型別的靜態方法。
    /// </summary>
    public static class SafeAccess
    {
        /// <summary>
        /// 安全地將物件轉換為去除前後空白的字串，若轉換失敗則回傳空字串。
        /// </summary>
        /// <param name="obj">要轉換的物件。</param>
        /// <returns>轉換後的字串，若失敗則回傳空字串。</returns>
        public static string SafeCellString(object? obj)
        {
            try
            {
                return obj?.ToString()?.Trim() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
