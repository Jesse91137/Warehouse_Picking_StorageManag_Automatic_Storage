using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Automatic_Storage
{
    class CMC_DC
    {
        // 檢查輸入的日期是否超過兩年
        public bool IsOverTwoYears(string inputDate)
        {
            bool f = false;
            if (TryParseDate(inputDate, out DateTime parsedDate))
            {
                DateTime currentDate = DateTime.Now;

                // 計算 parsedDate 加上 1 年 11 個月的日期
                DateTime expiryWarningDate = parsedDate.AddYears(1).AddMonths(11);

                // 判斷是否到達保存期限的前一個月
                return currentDate >= expiryWarningDate;
            }
            else
            {
                throw new ArgumentException("無法解析日期格式。");
            }
        }

        // 解析多種日期格式
        private bool TryParseDate(string input, out DateTime date)
        {
            string[] formats = new[]
            {
                "yyyy/M/d", "M/d/yyyy", "yyyy.MM.dd",
                "M.d.yyyy", "yyyyMMdd", "yyyy/MM/dd",
                "MM.dd.yyyy"
            };

            // 常見格式日期
            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(input, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    return true;
                }
            }

            // 處理年周格式（YYYYWW）
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

            // 處理年周格式（YYYYWW），如 2427 對應 2024 年的第 27 周
            if (input.Length == 4 && input.All(char.IsDigit))
            {
                string yearPart = "20" + input.Substring(0, 2); // 假設年分為 2000 年以後的年份
                string weekPart = input.Substring(2, 2);
                if (int.TryParse(yearPart, out int year) && int.TryParse(weekPart, out int week))
                {
                    date = FirstDateOfWeekISO8601(year, week);
                    return true;
                }
            }

            date = DateTime.MinValue;
            return false;
        }

        // 計算標準下給定年分和週數的第一天
        private DateTime FirstDateOfWeekISO8601(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            DateTime firstThursday = jan1.AddDays(daysOffset);
            var cal = CultureInfo.InvariantCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            var weekNum = weekOfYear;
            if (firstWeek <= 1)
            {
                weekNum -= 1;
            }

            var result = firstThursday.AddDays(weekNum * 7);
            return result.AddDays(-3);
        }
    }
}
