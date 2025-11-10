using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Automatic_Storage
{
    /// <summary>
    /// PCB DC 修改視窗表單
    /// </summary>
    public partial class PcbDC : Form
    {
        /// <summary>
        /// 建構函式，初始化元件
        /// </summary>
        public PcbDC()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 主鍵編號
        /// </summary>
        public string Sno { get; set; }

        /// <summary>
        /// PCB DC 原始值
        /// </summary>
        public string PcbDC_O { get; set; }

        /// <summary>
        /// PCB DC 新值
        /// </summary>
        public string PcbDC_N { get; set; }

        /// <summary>
        /// 按下修改按鈕事件，執行 PCB DC 更新
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void btn_pcbDC_Click(object sender, EventArgs e)
        {
            try
            {
                // 取得新輸入的 PCB DC 值
                PcbDC_N = txtPcbNew.Text;

                // 檢查新值是否為合法日期格式
                if (!string.IsNullOrWhiteSpace(PcbDC_N) && !TryParseDate(PcbDC_N, out DateTime parsedDate))
                {
                    MessageBox.Show("日期錯誤!! 請再次確認。");
                    return;
                }

                // 更新明細資料表的 PCB_DC 欄位
                string sql = @"UPDATE Automatic_Storage_Detail SET PCB_DC=@PCB_DC WHERE Sno = @Sno ";
                SqlParameter[] paramers = new SqlParameter[]
                {
                        new SqlParameter("PCB_DC",PcbDC_N),
                        new SqlParameter("Sno",Sno),
                };
                db.ExecueNonQuery(sql, CommandType.Text, paramers);

                // 更新歷史資料表的 PCB_DC 欄位
                string sql_i = @"update Automatic_Storage_Input set PCB_DC=@PCB_DC where sno = @Sno ";
                SqlParameter[] paramers_i = new SqlParameter[]
                {
                        new SqlParameter("PCB_DC",PcbDC_N),
                        new SqlParameter("Sno",Sno),
                };
                db.ExecueNonQuery(sql_i, CommandType.Text, paramers_i);

                // 關閉視窗
                this.Close();
            }
            catch (Exception ex)
            {
                // 顯示錯誤訊息
                MessageBox.Show(ex.Message);
            }
        }

        #region 日期格式驗証
        /// <summary>
        /// 嘗試解析多種日期格式，支援常見格式及年週格式
        /// </summary>
        /// <param name="input">輸入字串</param>
        /// <param name="date">解析後的日期</param>
        /// <returns>是否解析成功</returns>
        public bool TryParseDate(string input, out DateTime date)
        {
            string[] formats = new[]
            {
                    "yyyy/M/d", "M/d/yyyy", "yyyy.MM.dd",
                    "M.d.yyyy", "yyyyMMdd", "yyyy/MM/dd",
                    "MM.dd.yyyy"
                };

            // 嘗試解析常見日期格式
            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(input, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    return true;
                }
            }

            // 處理年週格式（YYYYWW）
            if (input.Length == 6 && input.Substring(4, 2).All(char.IsDigit) && input.Substring(0, 4).All(char.IsDigit))
            {
                string yearPart = input.Substring(0, 4);
                string weekPart = input.Substring(4, 2);
                if (int.TryParse(yearPart, out int year) && int.TryParse(weekPart, out int week))
                {
                    date = FirstDateOfWeekISO8601(year, week);
                    return true;
                }
            }

            // 處理年週格式（YYWW），如 2427 對應 2024 年第 27 週
            if (input.Length == 4 && input.All(char.IsDigit))
            {
                string yearPart = "20" + input.Substring(0, 2); // 假設年分為 2000 年以後
                string weekPart = input.Substring(2, 2);
                if (int.TryParse(yearPart, out int year) && int.TryParse(weekPart, out int week))
                {
                    date = FirstDateOfWeekISO8601(year, week);
                    return true;
                }
            }

            // 解析失敗，回傳最小日期
            date = DateTime.MinValue;
            return false;
        }

        /// <summary>
        /// 根據 ISO8601 標準，計算指定年份與週數的第一天（週一）
        /// </summary>
        /// <param name="year">年份</param>
        /// <param name="weekOfYear">週數</param>
        /// <returns>該週的第一天日期</returns>
        private DateTime FirstDateOfWeekISO8601(int year, int weekOfYear)
        {
            // 取得指定年份的 1 月 1 日
            DateTime jan1 = new DateTime(year, 1, 1);
            // 計算到第一個星期四的天數差
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            // 取得第一個星期四的日期
            DateTime firstThursday = jan1.AddDays(daysOffset);
            // 取得第一個星期四所在的週數
            var cal = CultureInfo.InvariantCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            // 計算指定週數的日期
            var weekNum = weekOfYear;
            if (firstWeek <= 1)
            {
                // 若第一週為 1，週數減 1
                weekNum -= 1;
            }

            // 計算該週的第一個星期四
            var result = firstThursday.AddDays(weekNum * 7);
            // 回傳該週的週一
            return result.AddDays(-3);
        }
        #endregion

        /// <summary>
        /// 載入事件，顯示原始 PCB DC 值
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void PcbDC_Load(object sender, EventArgs e)
        {
            // 顯示原始 PCB DC 值
            txtPcbOld.Text = PcbDC_O;
        }

        /// <summary>
        /// 視窗關閉事件，通知主畫面重新整理資料
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void PcbDC_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 呼叫主畫面重新整理資料
            Form1 ower = (Form1)this.Owner;
            ower.refreshData();
        }
    }
}
