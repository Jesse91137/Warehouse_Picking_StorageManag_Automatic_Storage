using System.Data;

namespace Automatic_Storage
{
    /// <summary>
    /// 提供入庫資料刪除相關功能的類別。
    /// </summary>
    class DelEngsr
    {
        /// <summary>
        /// 清除指定 <see cref="DataTable"/> 中的入庫資料。
        /// </summary>
        /// <param name="dt">要清除的資料表。</param>
        public void DelNow(DataTable dt)
        {
            //入庫清除
        }

        /// <summary>
        /// 清除程式啟動日前兩天的所有資料。
        /// </summary>
        public void Del2Day()
        {
            //程式啟動日前兩天的資料都清除
        }
    }
}
