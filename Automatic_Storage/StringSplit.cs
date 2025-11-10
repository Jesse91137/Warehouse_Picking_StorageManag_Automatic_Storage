namespace Automatic_Storage
{
    /// <summary>
    /// 提供字串分割與擷取相關的靜態方法。
    /// </summary>
    class StringSplit
    {
        #region 字串處理

        /// <summary>
        /// 取得字串左側指定長度的子字串。
        /// </summary>
        /// <param name="param">原始字串。</param>
        /// <param name="length">要擷取的長度。</param>
        /// <returns>左側子字串。</returns>
        public static string StrLeft(string param, int length)
        {
            // 從字串最左側開始擷取指定長度的子字串
            string result = param.Substring(0, length);
            return result;
        }

        /// <summary>
        /// 取得字串右側指定長度的子字串。
        /// </summary>
        /// <param name="param">原始字串。</param>
        /// <param name="length">要擷取的長度。</param>
        /// <returns>右側子字串。</returns>
        public static string StrRight(string param, int length)
        {
            // 從字串最右側開始擷取指定長度的子字串
            string result = param.Substring(param.Length - length, length);
            return result;
        }

        /// <summary>
        /// 取得字串中間指定位置與長度的子字串。
        /// </summary>
        /// <param name="param">原始字串。</param>
        /// <param name="startIndex">起始索引位置。</param>
        /// <param name="length">要擷取的長度。</param>
        /// <returns>中間子字串。</returns>
        public static string StrMid(string param, int startIndex, int length)
        {
            // 從指定起始位置擷取指定長度的子字串
            string result = param.Substring(startIndex, length);
            return result;
        }

        #endregion
    }
}
