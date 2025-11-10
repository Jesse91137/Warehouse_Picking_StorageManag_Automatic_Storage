using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Automatic_Storage.Utilities;

namespace Automatic_Storage.Services
{
    /// <summary>
    /// 提供建立料號快速索引與模糊比對的純邏輯服務。
    /// - 不直接修改 UI 樣式（例如標色），呼叫端負責視覺呈現。
    /// - 可由表單建立並重複使用（透過傳入 DataGridView 物件）。
    /// </summary>
    public class MaterialIndexService
    {
        /// <summary>
        /// 來源 DataGridView 物件，作為資料來源。
        /// </summary>
        private readonly DataGridView _dgv;

        /// <summary>
        /// 以料號為鍵，儲存對應的 DataGridViewRow 列集合。
        /// </summary>
        private readonly Dictionary<string, List<DataGridViewRow>> _materialIndex = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 以料號為鍵，儲存已發數量的加總。
        /// </summary>
        private readonly Dictionary<string, decimal> _materialShippedSums = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 取得料號索引資料（唯讀）。
        /// </summary>
        public IReadOnlyDictionary<string, List<DataGridViewRow>> MaterialIndex => _materialIndex;

        /// <summary>
        /// 取得已發數量加總資料（唯讀）。
        /// </summary>
        public IReadOnlyDictionary<string, decimal> MaterialShippedSums => _materialShippedSums;

        /// <summary>
        /// 建構 MaterialIndexService，需傳入 DataGridView 物件。
        /// </summary>
        /// <param name="dgv">資料來源的 DataGridView。</param>
        /// <exception cref="ArgumentNullException">若傳入的 dgv 為 null 則拋出例外。</exception>
        public MaterialIndexService(DataGridView dgv)
        {
            _dgv = dgv ?? throw new ArgumentNullException(nameof(dgv));
        }

        /// <summary>
        /// 建立內部索引與已發數量快取（從目前 DataGridView 讀取）。
        /// 呼叫端應在需要時先停止 DataGridView 的重繪以提升效能。
        /// </summary>
        public void BuildIndex()
        {
            _materialIndex.Clear();
            _materialShippedSums.Clear();
            if (_dgv?.Rows == null) return;

            int materialCol = FindColumnIndexByNames(new[] { "昶亨料號", "客戶料號" });
            if (materialCol < 0) return;

            foreach (DataGridViewRow row in _dgv.Rows)
            {
                if (row == null || row.IsNewRow) continue;
                try
                {
                    var raw = row.Cells[materialCol].Value;
                    var val = raw?.ToString()?.Trim();
                    if (string.IsNullOrEmpty(val)) continue;
                    var key = val;
                    if (!_materialIndex.TryGetValue(key, out var list))
                    {
                        list = new List<DataGridViewRow>(4);
                        _materialIndex[key] = list;
                    }
                    if (list.Count == 0 || !object.ReferenceEquals(list[list.Count - 1], row)) list.Add(row);
                }
                catch { }
            }

            // 建立已發數量總和快取
            try
            {
                int shippedCol = FindColumnIndexByNames(new[] { "實發數量", "發料數量" });
                if (shippedCol >= 0)
                {
                    foreach (DataGridViewRow row in _dgv.Rows)
                    {
                        try
                        {
                            if (row == null || row.IsNewRow) continue;
                            var rawMat = row.Cells[materialCol].Value?.ToString()?.Trim();
                            if (string.IsNullOrEmpty(rawMat)) continue;
                            var key = TextParsing.NormalizeMaterialKey(rawMat);
                            if (string.IsNullOrEmpty(key)) continue;
                            decimal v = 0m;
                            var sv = row.Cells.Count > shippedCol ? row.Cells[shippedCol].Value?.ToString() ?? string.Empty : string.Empty;
                            if (TextParsing.TryParseDecimalValue(sv, out decimal parsed)) v = parsed;
                            if (_materialShippedSums.TryGetValue(key, out decimal exist)) _materialShippedSums[key] = exist + v; else _materialShippedSums[key] = v;
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 清除所有索引與已發數量快取資料。
        /// </summary>
        public void Clear()
        {
            _materialIndex.Clear();
            _materialShippedSums.Clear();
        }

        /// <summary>
        /// 針對輸入字串做模糊匹配，回傳匹配到的列與欄位索引（來源欄位，昶亨或客戶料號）。
        /// 呼叫端可依結果做 UI 樣式變更或顯示備註。
        /// </summary>
        /// <param name="input">欲比對的輸入字串。</param>
        /// <returns>回傳所有符合條件的 DataGridViewRow 及欄位索引。</returns>
        public List<(DataGridViewRow row, int colIdx)> Match(string input)
        {
            var result = new List<(DataGridViewRow, int)>();
            if (string.IsNullOrEmpty(input)) return result;

            int chCol = -1, custCol = -1, remarkCol = -1;
            foreach (DataGridViewColumn col in _dgv.Columns)
            {
                var name = (col.HeaderText ?? string.Empty).Trim();
                if (name.Contains("昶亨料號")) chCol = col.Index;
                if (name.Contains("客戶料號")) custCol = col.Index;
                if (name.Contains("備註")) remarkCol = col.Index;
            }
            if (chCol < 0 && custCol < 0) return result;

            var inputTrimmed = input.Trim();
            var inputNorm = TextParsing.NormalizeMaterialKey(inputTrimmed);
            var inputUpper = inputTrimmed.ToUpperInvariant();

            // 1) exact normalized match
            foreach (DataGridViewRow row in _dgv.Rows)
            {
                if (row.IsNewRow) continue;
                if (chCol >= 0 && row.Cells.Count > chCol && row.Cells[chCol].Value != null)
                {
                    var val = row.Cells[chCol].Value?.ToString() ?? string.Empty;
                    var valNorm = TextParsing.NormalizeMaterialKey(val);
                    if (!string.IsNullOrEmpty(valNorm) && !string.IsNullOrEmpty(inputNorm) && valNorm == inputNorm)
                    {
                        result.Add((row, chCol));
                        continue;
                    }
                }
                if (custCol >= 0 && row.Cells.Count > custCol && row.Cells[custCol].Value != null)
                {
                    var val = row.Cells[custCol].Value?.ToString() ?? string.Empty;
                    var valNorm = TextParsing.NormalizeMaterialKey(val);
                    if (!string.IsNullOrEmpty(valNorm) && !string.IsNullOrEmpty(inputNorm) && valNorm == inputNorm)
                    {
                        result.Add((row, custCol));
                        continue;
                    }
                }
            }

            // 2) prefix/contains or uppercase contains
            if (result.Count == 0)
            {
                foreach (DataGridViewRow row in _dgv.Rows)
                {
                    if (row.IsNewRow) continue;
                    if (chCol >= 0 && row.Cells.Count > chCol && row.Cells[chCol].Value != null)
                    {
                        var val = row.Cells[chCol].Value?.ToString() ?? string.Empty;
                        var valNorm = TextParsing.NormalizeMaterialKey(val);
                        if (!string.IsNullOrEmpty(valNorm) && !string.IsNullOrEmpty(inputNorm) && (valNorm.StartsWith(inputNorm) || valNorm.Contains(inputNorm))) { result.Add((row, chCol)); continue; }
                        else if (!string.IsNullOrEmpty(val) && val.ToUpperInvariant().Contains(inputUpper)) { result.Add((row, chCol)); continue; }
                    }
                    if (custCol >= 0 && row.Cells.Count > custCol && row.Cells[custCol].Value != null)
                    {
                        var val = row.Cells[custCol].Value?.ToString() ?? string.Empty;
                        var valNorm = TextParsing.NormalizeMaterialKey(val);
                        if (!string.IsNullOrEmpty(valNorm) && !string.IsNullOrEmpty(inputNorm) && (valNorm.StartsWith(inputNorm) || valNorm.Contains(inputNorm))) { result.Add((row, custCol)); continue; }
                        else if (!string.IsNullOrEmpty(val) && val.ToUpperInvariant().Contains(inputUpper)) { result.Add((row, custCol)); continue; }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 依候選名稱尋找欄位索引（內部複用 Form 的行為，但不倚賴 Form 狀態）。
        /// </summary>
        /// <param name="names">候選欄位名稱集合。</param>
        /// <returns>找到的欄位索引，若找不到則回傳 -1。</returns>
        private int FindColumnIndexByNames(IEnumerable<string> names)
        {
            if (_dgv?.Columns == null || _dgv.Columns.Count == 0) return -1;
            /// <summary>
            /// 將欄位名稱標準化以利比對。
            /// </summary>
            /// <param name="s">原始欄位名稱。</param>
            /// <returns>標準化後的欄位名稱。</returns>
            string San(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                try { var global = TextParsing.SanitizeHeaderForMatch(s); if (!string.IsNullOrEmpty(global)) return global; } catch { }
                s = s.Replace('\u0020', ' ').Replace('\u3000', ' ').Trim();
                var sb = new System.Text.StringBuilder(s.Length);
                foreach (var ch in s)
                {
                    if (char.IsLetterOrDigit(ch) || (ch >= 0x4E00 && ch <= 0x9FFF)) sb.Append(char.ToLowerInvariant(ch));
                }
                return sb.ToString();
            }

            var exactTargets = names.Where(n => !string.IsNullOrWhiteSpace(n)).Select(n => n.Trim()).ToHashSet(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < _dgv.Columns.Count; i++)
            {
                var col = _dgv.Columns[i];
                var candidates = new[] { col.HeaderText, col.Name, col.DataPropertyName };
                if (candidates.Any(c => !string.IsNullOrWhiteSpace(c) && exactTargets.Contains(c.Trim()))) return i;
            }

            var normTargets = names.Select(San).Where(x => !string.IsNullOrEmpty(x)).ToList();
            for (int i = 0; i < _dgv.Columns.Count; i++)
            {
                var col = _dgv.Columns[i];
                var candidates = new[] { col.HeaderText, col.Name, col.DataPropertyName };
                var normColNames = candidates.Where(x => !string.IsNullOrWhiteSpace(x)).Select(San).ToArray();
                foreach (var nt in normTargets)
                {
                    if (normColNames.Any(nc => !string.IsNullOrEmpty(nc) && (nc.Contains(nt) || nt.Contains(nc)))) return i;
                }
            }
            return -1;
        }
    }
}
