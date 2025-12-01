using System;

namespace Automatic_Storage.Dto
{
    /// <summary>
    /// 表示一筆備料單的記錄資料，包含刷入時間、料號、數量及操作者資訊。
    /// </summary>
    public class 記錄Dto
    {
        /// <summary>
        /// Gets or sets the 刷入時間.
        /// </summary>
        /// <value>
        /// The 刷入時間.
        /// </value>
        public DateTime 刷入時間 { get; set; }

        /// <summary>
        /// Gets or sets the 料號.（可為昶亨料號或客戶料號）。
        /// </summary>
        /// <value>
        /// The 料號.
        /// </value>
        public string 料號 { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the 實際刷入的數量。
        /// </summary>
        /// <value>
        /// The 數量.
        /// </value>
        public int 數量 { get; set; }

        /// <summary>
        /// Gets or sets the 操作者名稱（登入帳號或系統使用者）。
        /// </summary>
        /// <value>
        /// The 操作者.
        /// </value>
        public string 操作者 { get; set; } = string.Empty;
    }
}
