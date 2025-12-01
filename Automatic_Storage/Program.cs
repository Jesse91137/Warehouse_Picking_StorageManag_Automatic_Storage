using System;
using System.IO;
using System.Windows.Forms;
using Automatic_Storage.Services;

namespace Automatic_Storage
{

    /// <summary>
    /// 應用程式的主要類別，負責啟動 Windows Forms 應用程式。
    /// </summary>
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            /// <summary>
            /// 啟用視覺化樣式，使應用程式的控制項符合目前 Windows 主題。
            /// </summary>
            Application.EnableVisualStyles();

            /// <summary>
            /// 設定應用程式使用預設的文字轉譯方式，提升文字顯示品質。
            /// </summary>
            Application.SetCompatibleTextRenderingDefault(false);

            // 確保 Upload 資料夾存在
            try
            {
                string uploadDir = Path.Combine(Application.StartupPath, "Upload");
                if (!Directory.Exists(uploadDir)) Directory.CreateDirectory(uploadDir);
            }
            catch
            {
                // CreateDirectory 可能因權限限制失敗，忽略該例外讓應用程式繼續執行
            }

            /// <summary>
            /// 啟動主視窗 Form1，進入應用程式主循環。
            /// </summary>
            Application.Run(new Form1());

            // 應用程式主循環結束後，嘗試優雅釋放 ExcelWriteQueue singleton
            try
            {
                try { Automatic_Storage.Utilities.ExcelWriteQueue.Instance.Dispose(); } catch { }
            }
            catch { }
        }
    }
}
