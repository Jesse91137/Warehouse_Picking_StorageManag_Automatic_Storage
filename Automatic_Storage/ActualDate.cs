using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Automatic_Storage
{
    /// <summary>
    /// 實際入庫日期異動視窗，提供入庫日期修改功能。
    /// </summary>
    public partial class ActualDate : Form
    {
        /// <summary>
        /// 建構函式，初始化 ActualDate 視窗元件。
        /// </summary>
        public ActualDate()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 主鍵，對應資料表的 sno 欄位。
        /// </summary>
        public string Sno { get; set; }

        /// <summary>
        /// 原始入庫日期，顯示於視窗中。
        /// </summary>
        public string AcD_O { get; set; }

        /// <summary>
        /// 異動後的入庫日期。
        /// </summary>
        public string AcD_M { get; set; }

        /// <summary>
        /// 異動入庫日期按鈕事件，檢查日期合法性並更新資料庫。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void btn_actualDate_Click(object sender, EventArgs e)
        {
            try
            {
                AcD_M = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
                TimeSpan Diff_dates = dateTimePicker1.Value.Date.Subtract(DateTime.Today);
                if (Diff_dates.TotalDays >= 1)
                {
                    MessageBox.Show("日期錯誤!! 請再次確認。");
                    return;
                }
                // index: 更新明細表
                string sql = @"update Automatic_Storage_Detail set Actual_InDate=@InDate where sno = @Sno ";
                SqlParameter[] paramers = new SqlParameter[]
                {
                        new SqlParameter("InDate",AcD_M),
                        new SqlParameter("Sno",Sno),
                };
                db.ExecueNonQuery(sql, CommandType.Text, paramers);

                // history: 更新入庫記錄表
                string sql_i = @"update Automatic_Storage_Input set Actual_InDate=@InDate where sno = @Sno ";
                SqlParameter[] paramers_i = new SqlParameter[]
                {
                    new SqlParameter("InDate",AcD_M),
                    new SqlParameter("Sno",Sno),
                };
                db.ExecueNonQuery(sql_i, CommandType.Text, paramers_i);

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// 視窗載入事件，顯示原始入庫日期。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void ActualDate_Load(object sender, EventArgs e)
        {
            // 載入視窗時，顯示原始入庫日期
            textBox1.Text = AcD_O;
        }

        /// <summary>
        /// 視窗關閉事件，通知主視窗重新整理資料。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void ActualDate_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 關閉視窗時，通知主視窗更新資料
            Form1 ower = (Form1)this.Owner;
            // 呼叫主視窗的 refreshData 方法
            ower.refreshData();
        }
    }
}
