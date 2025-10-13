using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Automatic_Storage
{
    class StringSplit
    {
        #region 字串處理
        public static string StrLeft(string param, int length)
        {
            string result = param.Substring(0, length);
            return result;
        }
        public static string StrRight(string param, int length)
        {
            string result = param.Substring(param.Length - length, length);
            return result;
        }

        public static string StrMid(string param, int startIndex, int length)
        {
            string result = param.Substring(startIndex, length);
            return result;
        }

        #endregion
    }
}
