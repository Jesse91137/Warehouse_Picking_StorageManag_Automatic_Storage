using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automatic_Storage.Dto
{

    /// <summary>
    /// 存檔結果結構。
    /// 用於表示存檔作業是否成功及相關錯誤訊息。
    /// </summary>
    internal class SaveResultDto
    {
        /// <summary>
        /// 取得或設定存檔是否成功。
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 取得或設定錯誤訊息。
        /// 若存檔失敗，將包含詳細錯誤說明；成功時為 null 或空字串。
        /// </summary>
        public string? ErrorMessage { get; set; }
    }
}
