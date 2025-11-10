using System.Text;
using System.Linq;
using System.Text.RegularExpressions;

namespace Automatic_Storage.Utilities
{
    /// <summary>
    /// Common text parsing helpers extracted from Form備料單匯入 for reuse and testing.
    /// </summary>
    public static class TextParsing
    {
        /// <summary>
        /// 將料號字串正規化以便比對：去除前後空白、轉大寫，移除不可見或特殊字元
        /// </summary>
        public static string NormalizeMaterialKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            var t = s.Trim().ToUpperInvariant();
            // 移除全形空白與特殊不可見字元
            t = t.Replace('\u3000', ' ').Trim();
            // 移除常見非英數符號
            var sb = new StringBuilder();
            foreach (var ch in t)
            {
                if (char.IsLetterOrDigit(ch)) sb.Append(ch);
            }
            return sb.ToString();
        }

        /// <summary>
        /// 將用於回填 textbox 的料號做輕量正規化：
        /// - 去除前後空白
        /// - 把全形空白與不間斷空白轉成半形空格
        /// - 移除常見的零寬或格式控制字元 (ZWSP, ZWNJ, ZWJ, BOM)
        /// - 收斂任意連續空白為單一半形空格
        /// </summary>
        public static string NormalizeForTextboxFill(string? s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            string t = s;
            try
            {
                // 替換全形空白與 no-break space 為一般空格
                t = t.Replace('\u3000', ' ');
                t = t.Replace('\u00A0', ' ');

                // 移除零寬與格式控制字元
                t = t.Replace("\u200B", "").Replace("\u200C", "").Replace("\u200D", "").Replace("\uFEFF", "");

                // 移除控制字元 (Cc) 與格式字元 (Cf)
                t = Regex.Replace(t, "[\\p{Cc}\\p{Cf}]+", "");

                // 移除 Private Use Area, Presentation Forms 等可能包含奇怪字形的區段
                // U+E000 - U+F8FF (Private Use), U+FB00 - U+FB4F (Presentation Forms-A), U+FE70 - U+FEFF (Presentation Forms-B)
                t = Regex.Replace(t, "[\uE000-\uF8FF\uFB00-\uFB4F\uFE70-\uFEFF]+", "");

                // 收斂任何空白序列為單一半形空格，並 Trim
                t = Regex.Replace(t.Trim(), "\\s+", " ");

                // 最後以白名單過濾：保留所有字母、數字、連字號、底線、點、斜線與空格
                // 這一步可確保像 Private Use 或 Presentation Forms 的奇怪字元被移除
                t = Regex.Replace(t, "[^\\p{L}\\p{Nd}\\-_.\\/ ]+", "");
            }
            catch
            {
                try { t = (s ?? string.Empty).Trim(); } catch { t = string.Empty; }
            }
            return t;
        }

        /// <summary>
        /// 彈性解析 decimal（處理逗號、全形字元、空白、Culture）
        /// </summary>
        public static bool TryParseDecimalFlexible(string s, out decimal result)
        {
            result = 0m;
            if (string.IsNullOrWhiteSpace(s)) return false;

            string t = s.Trim();

            // 移除全形空白與常見不可見字元
            t = t.Replace('\u3000', ' ').Trim();

            // 移除千分符常見格式（逗號），但保留小數點
            t = t.Replace(",", "");

            // 嘗試以目前文化解析
            if (decimal.TryParse(t, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.CurrentCulture, out result))
                return true;

            // 再試 InvariantCulture
            if (decimal.TryParse(t, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out result))
                return true;

            // 嘗試去掉所有非數字與小數點字元，再解析（最後保底）
            var filtered = new StringBuilder();
            foreach (var ch in t)
            {
                if (char.IsDigit(ch) || ch == '.' || ch == '-') filtered.Append(ch);
            }
            if (decimal.TryParse(filtered.ToString(), System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out result))
                return true;

            return false;
        }

        /// <summary>
        /// 比對用：將標頭名稱標準化（移除非英數字元並轉小寫）
        /// </summary>
        public static string SanitizeHeaderForMatch(string s)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            var arr = s.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray();
            return new string(arr);
        }

        /// <summary>
        /// 與舊有介面相容的解析函式，封裝至 TryParseDecimalFlexible。
        /// 保留舊方法名稱以減少對現有呼叫端的影響。
        /// </summary>
        public static bool TryParseDecimalValue(string s, out decimal result)
        {
            return TryParseDecimalFlexible(s ?? string.Empty, out result);
        }
    }
}
