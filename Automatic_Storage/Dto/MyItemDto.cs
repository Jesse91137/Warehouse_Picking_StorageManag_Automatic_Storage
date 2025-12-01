using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automatic_Storage.Dto
{
    /// <summary>
    /// 表示下拉選單中的包裝種類項目。
    /// </summary>
    internal class MyItemDto
    {
        /// <summary>
        /// 建立 MyItem 物件並指定顯示文字與值。
        /// </summary>
        /// <param name="text">顯示於下拉選單的文字。</param>
        /// <param name="value">對應值。</param>
        public MyItemDto(string text, string value)
        {
            this.text = text; // 設定顯示文字
            this.value = value; // 設定對應值
        }

        /// <summary>
        /// 顯示於下拉選單的文字。
        /// </summary>
        public string text; // 顯示文字

        /// <summary>
        /// 對應的包裝代碼值。
        /// </summary>
        public string value; // 對應值

        /// <summary>
        /// 傳回顯示文字，供下拉選單顯示用。
        /// </summary>
        /// <returns>顯示文字。</returns>
        public override string ToString()
        {
            return text; // 回傳顯示文字
        }
    }
}
