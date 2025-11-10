using System;
using System.Globalization;
using System.Linq;

namespace Automatic_Storage
{
    /// <summary>
    /// 提供日期解析與保存期限判斷的功能。
    /// </summary>
    class CMC_DC
    {
        /// <summary>
        /// 檢查輸入的日期是否超過兩年（即是否到達保存期限的前一個月）。
        /// </summary>
        /// <param name="inputDate">要檢查的日期字串，支援多種格式。</param>
        /// <returns>如果已超過保存期限則回傳 true，否則回傳 false。</returns>
        /// <exception cref="ArgumentException">當無法解析日期格式時拋出。</exception>
        public bool IsOverTwoYears(string inputDate)
        {
            bool f = false; // 宣告布林變數 f，初始值為 false
            if (TryParseDate(inputDate, out DateTime parsedDate)) // 嘗試解析輸入的日期字串
            {
                DateTime currentDate = DateTime.Now; // 取得目前的日期時間

                // 計算 parsedDate 加上 1 年 11 個月的日期
                DateTime expiryWarningDate = parsedDate.AddYears(1).AddMonths(11);

                // 判斷是否到達保存期限的前一個月
                return currentDate >= expiryWarningDate;
            }
            else
            {
                throw new ArgumentException("無法解析日期格式。"); // 無法解析時丟出例外
            }
        }

        /// <summary>
        /// 嘗試解析多種日期格式，包含常見日期格式與年周格式。
        /// </summary>
        /// <param name="input">要解析的日期字串。</param>
        /// <param name="date">解析成功時回傳的日期物件。</param>
        /// <returns>解析成功則回傳 true，否則回傳 false。</returns>
        private bool TryParseDate(string input, out DateTime date)
        {
            string[] formats = new[]
            {
                    "yyyy/M/d", "M/d/yyyy", "yyyy.MM.dd",
                    "M.d.yyyy", "yyyyMMdd", "yyyy/MM/dd",
                    "MM.dd.yyyy"
                }; // 支援的日期格式陣列

            // 常見格式日期
            foreach (var format in formats) // 逐一嘗試每種格式
            {
                if (DateTime.TryParseExact(input, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) // 嘗試解析
                {
                    return true; // 解析成功則回傳 true
                }
            }

            // 處理年周格式（YYYYWW）
            if (input.Length == 6 && input.Substring(4, 2).All(char.IsDigit) && input.Substring(0, 4).All(char.IsDigit)) // 判斷是否為 6 位數年周格式
            {
                string yearPart = input.Substring(0, 4); // 取出年份部分
                string weekPart = input.Substring(4, 2); // 取出週數部分
                if (int.TryParse(yearPart, out int year) && int.TryParse(weekPart, out int week)) // 嘗試轉換為整數
                {
                    date = FirstDateOfWeekISO8601(year, week); // 取得該週的第一天
                    return true; // 解析成功
                }
            }

            // 處理年周格式（YYYYWW），如 2427 對應 2024 年的第 27 周
            if (input.Length == 4 && input.All(char.IsDigit)) // 判斷是否為 4 位數年周格式
            {
                string yearPart = "20" + input.Substring(0, 2); // 假設年分為 2000 年以後的年份
                string weekPart = input.Substring(2, 2); // 取出週數部分
                if (int.TryParse(yearPart, out int year) && int.TryParse(weekPart, out int week)) // 嘗試轉換為整數
                {
                    date = FirstDateOfWeekISO8601(year, week); // 取得該週的第一天
                    return true; // 解析成功
                }
            }

            date = DateTime.MinValue; // 解析失敗時回傳最小日期
            return false; // 回傳 false
        }

        /// <summary>
        /// 根據 ISO 8601 標準，計算指定年份與週數的第一天（週一）。
        /// </summary>
        /// <param name="year">年份。</param>
        /// <param name="weekOfYear">週數。</param>
        /// <returns>該週的第一天（週一）。</returns>
        private DateTime FirstDateOfWeekISO8601(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1); // 取得該年一月一日
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek; // 計算與週四的天數差

            DateTime firstThursday = jan1.AddDays(daysOffset); // 取得該年第一個週四
            var cal = CultureInfo.InvariantCulture.Calendar; // 取得不變文化的行事曆
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday); // 取得第一週的週數

            var weekNum = weekOfYear; // 設定週數
            if (firstWeek <= 1) // 如果第一週小於等於 1
            {
                weekNum -= 1; // 週數減 1
            }

            var result = firstThursday.AddDays(weekNum * 7); // 計算該週的日期
            return result.AddDays(-3); // 回傳該週的週一
        }
    }
}
