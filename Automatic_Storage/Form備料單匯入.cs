using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
// NPOI for fast .xls/.xlsx reading (avoids COM)
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using static Automatic_Storage.Utilities.TextParsing;
using static Automatic_Storage.Utilities.ComInterop;
using Automatic_Storage.Services;
using Automatic_Storage.Utilities;
using Automatic_Storage.Dto;

namespace Automatic_Storage
{
    /// <summary>
    /// 表單: 備料單匯入
    /// - 保持 Designer 版面
    /// - 對大型 Excel 使用 background streaming 讀取，並在 UI 執行緒合併批次
    /// - 支援在關閉表單時取消讀取
    /// </summary>
    public partial class Form備料單匯入 : Form
    {
        // 當此旗標為 true 時，WriteBackShippedQuantitiesToExcelBatch 會在同一個 Workbook 會話中
        // 將 _records 一併寫入「記錄」工作表，避免重複開關檔案造成的效能負擔。
        // 此旗標僅在 SaveAsyncWithResult 的短暫期間內被設定。
        private bool _mergeAppendIntoWriteBack = false;


        #region 宣告變數
        /// <summary>
        /// The default password
        /// </summary>
        private string defaultPwd = "1234";
        /// <summary>
        /// 記錄Dto容器，儲存所有備料單的刷入記錄。
        /// 每當使用者於畫面上進行數量刷入時，會將該筆資料（包含刷入時間、料號、數量、操作者）
        /// 以 <see cref="Dto.記錄Dto"/> 物件形式加入此清單，待存檔時一併寫入 Excel「記錄」工作表。
        /// </summary>
        private List<Dto.記錄Dto> _records = new List<Dto.記錄Dto>();
        /// <summary>
        /// 最近一次成功寫入 Excel「記錄」工作表的記錄清單。
        /// 用於後續從 <see cref="_records"/> 中移除已寫入的記錄，避免重複寫入。
        /// 存檔成功後，會將本次實際寫入的 <see cref="Dto.記錄Dto"/> 物件存於此集合。
        /// </summary>
        private List<Dto.記錄Dto> _lastAppendedRecords = new List<Dto.記錄Dto>();
        /// <summary>
        /// 保存 DataGridView 中紅色高亮（短缺）狀態的料號 key 集合。
        /// 用於在 UI 重建或寫回 Excel 後，保留短缺標示。
        /// 內容為經過標準化的料號字串（不分大小寫），以利快速查詢與還原紅色提示。
        /// </summary>
        private HashSet<string> _preservedRedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// 目前載入的 Excel 檔案路徑（若有）。
        /// 儲存目前已經載入的 Excel 檔案完整路徑，供後續存檔、匯出等功能使用。
        /// 若尚未載入任何檔案則為 null。
        /// </summary>
        private string? currentExcelPath = null;

        /// <summary>
        /// 使用者通知器，用於顯示訊息（資訊、警告、錯誤）。
        /// 透過建構子注入，允許 UI 或測試環境使用不同的實作。
        /// 預設使用 WinFormsUserNotifier（MessageBox）。
        /// </summary>
        private IUserNotifier _userNotifier = null!;

        /// <summary>
        /// 匯入/背景作業的取消控制元件。
        /// 用於控制匯入 Excel 或其他長時間背景作業的取消，
        /// 讓使用者可在表單關閉或操作中斷時安全終止相關作業。
        /// </summary>
        private CancellationTokenSource? _cts = null;
        /// <summary>
        /// 儲存被停用的控制項及其先前的 Enabled 狀態（含各種按鈕，遞迴收集）。
        /// 於長時間作業期間暫時停用 UI，作業結束後可還原原本狀態。
        /// </summary>
        private Dictionary<Control, bool>? _prevControlStates = null;
        /// <summary>
        /// 儲存 ToolStripItem（例如 MenuStrip / ToolStrip 的按鈕）先前的 Enabled 狀態。
        /// 於長時間作業期間暫時停用工具列按鈕，作業結束後可還原原本狀態。
        /// </summary>
        private Dictionary<ToolStripItem, bool>? _prevToolStripItemStates = null;
        /// <summary>
        /// 標示目前 UI 是否正在匯入資料，避免重複操作。
        /// 當進行匯入作業時設為 true，作業結束後還原為 false。
        /// </summary>
        private volatile bool _isImporting = false;
        /// <summary>
        /// 標示目前 UI 是否正在匯出資料，避免重複操作。
        /// 當進行匯出作業時設為 true，作業結束後還原為 false。
        /// </summary>
        private volatile bool _isExporting = false;
        /// <summary>
        /// 標示目前 UI 是否正在存檔，避免重複操作。
        /// 當進行存檔作業時設為 true，作業結束後還原為 false。
        /// </summary>
        private volatile bool _isSaving = false;
        /// <summary>
        /// 按下匯入檔案後，強制保留匯入按鈕為不可按直到表單關閉。
        /// 用於防止重複匯入造成資料異常。
        /// </summary>
        private bool _keepImportButtonDisabledUntilClose = false;
        /// <summary>
        /// 在成功匯出檔案後，強制保留 Unlock 按鈕為不可按直到表單關閉。
        /// 用於避免匯出後立即被使用者重新解鎖導致資料不一致的操作。
        /// </summary>
        private bool _keepUnlockButtonDisabledUntilClose = false;
        /// <summary>
        /// 標示表單先前是否處於最小化狀態，用於偵測從 Minimized 還原。
        /// 以便於還原時正確處理 UI 狀態。
        /// </summary>
        private volatile bool _wasMinimized = false;
        /// <summary>
        /// 快取 Application.UseWaitCursor 的先前狀態，便於還原。
        /// 用於長時間作業時切換等待游標，作業結束後還原原本狀態。
        /// </summary>
        private bool _prevUseWaitCursor = false;
        /// <summary>
        /// 長時間操作時用於遮罩輸入的 overlay 面板集合。
        /// 每個 overlay 對應一個表單，顯示處理中訊息並攔截輸入。
        /// </summary>
        private List<Panel>? _operationOverlays = null;
        /// <summary>
        /// 操作遮罩上顯示訊息的 Label。
        /// 用於顯示目前處理中的提示文字。
        /// </summary>
        // _overlayLabel removed (was unused) — field deleted to reduce unused-member warnings.
        /// <summary>
        /// Excel 欄位名稱對應 DataTable 欄位索引的對應表。
        /// 主要用於存檔時將 DataTable 欄位正確對應回 Excel 欄位。
        /// </summary>
        private Dictionary<string, int>? _columnMapping = null;
        /// <summary>
        /// Excel 服務物件，允許外部注入以便測試或替換實作（可為 null）。
        /// 若有自訂 Excel 操作服務，可透過此欄位注入。
        /// </summary>
        // allow dynamic excel service to be nullable; we guard assignments and uses
        private dynamic? _excelService = null;
        /// <summary>
        /// 可選的 Typed Excel service（優先使用）。若為 null，會回退到舊有 dynamic 物件的 adapter。
        /// 注入此欄位可提供強型別、可測試的 Excel 實作而不破壞現有邏輯。
        /// </summary>
        private IExcelService? _typedExcelService = null;
        /// <summary>
        /// 快速索引：料號對應的 DataGridViewRow 清單。
        /// 以料號為 key，對應所有出現該料號的資料列，便於快速查找與操作。
        /// </summary>
        private Dictionary<string, List<DataGridViewRow>> _materialIndex = new Dictionary<string, List<DataGridViewRow>>(StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// The material index service
        /// </summary>
        private Automatic_Storage.Services.MaterialIndexService? _materialIndexService;
        /// <summary>
        /// 快取每個料號的已發/實發數量總和，避免在 CellFormatting 中對整張表做 O(N) 掃描。
        /// Key 為 NormalizeMaterialKey 後的料號字串。
        /// </summary>
        private Dictionary<string, decimal> _materialShippedSums = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// 用於標示料號儲存格的顏色（統一管理，便於未來調整）。
        /// 目前預設為黃色，標示比對成功的料號。
        /// </summary>
        private readonly Color _materialHighlightColor = Color.Yellow;
        /// <summary>
        /// 快取目前顯示的 DataTable，避免 Hide/Show 後資料來源遺失。
        /// 用於表單隱藏再顯示時還原資料。
        /// </summary>
        private DataTable? _currentUiTable = null;
        /// <summary>
        /// 標示 UI 是否有未儲存的變更（供匯出前檢查）。
        /// 若有資料異動但尚未存檔，則設為 true。
        /// </summary>
        private bool _isDirty = false;
        /// <summary>
        /// 當程式性地改變 UI 或欄位屬性時，暫時抑制標記為已修改，避免誤觸發未存檔提示。
        /// 例如自動鎖定/解鎖或批次更新時使用。
        /// </summary>
        private volatile bool _suspendDirtyMarking = false;

        /// <summary>
        /// Win32 API：用於 BeginUpdate/EndUpdate 控制畫面重繪（WM_SETREDRAW）。
        /// 透過呼叫此 API 可暫停或恢復控制項的重繪，提升大量更新時的效能。
        /// </summary>
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
        /// <summary>
        /// WM_SETREDRAW 訊息常數，用於控制控制項重繪。
        /// 傳送此訊息可暫停或恢復控制項的畫面更新。
        /// </summary>
        private const int WM_SETREDRAW = 0x000B;

        /// <summary>
        /// 操作者
        /// </summary>
        private string operatorName = "";

        /// <summary>
        /// Excel 回寫相關佇列/同步控制
        /// </summary>
        private CancellationTokenSource? _excelQueueCts = null;

        /// <summary>
        /// 讀取 Excel 時擷取的每個儲存格對齊資訊 (row -> list of alignments per column)
        /// </summary>
        private List<List<DataGridViewContentAlignment>>? _excelAlignments = null;

        /// <summary>
        /// Resize 去抖計時器，用於視窗調整大小時避免過度自動調整欄寬。
        /// </summary>
        /// <remarks>
        /// 當視窗尺寸變動時，會啟動此計時器，僅在使用者停止調整後才執行自動欄寬調整，
        /// 以提升效能並避免頻繁重繪造成的延遲。
        /// </remarks>
        private System.Windows.Forms.Timer? _resizeTimer = null;

        /// <summary>
        /// 顯示被裁切儲存格內容的 ToolTip。
        /// </summary>
        /// <remarks>
        /// 當 DataGridView 儲存格內容超出顯示範圍時，使用此 ToolTip 顯示完整內容。
        /// 於 CellMouseMove、CellMouseLeave、Scroll 等事件中動態顯示與隱藏。
        /// </remarks>
        private System.Windows.Forms.ToolTip? _cellToolTip = null;

        /// <summary>
        /// 密碼快取與鎖
        /// </summary>
        /// <returns></returns>
        private readonly object _pwdLock = new object();

        /// <summary>
        /// The cached excel password
        /// <summary>
        /// 快取的 Excel 密碼。
        /// </summary>
        /// <remarks>
        /// 由 <see cref="GetExcelPassword(int)"/> 取得並快取，減少重複查詢資料庫的次數。
        /// </remarks>
        private string _cachedExcelPassword = string.Empty;

        /// <summary>
        /// 快取的 Excel 密碼取得時間。
        /// </summary>
        /// <remarks>
        /// 用於判斷密碼快取是否過期，單位為 UTC。
        /// </remarks>
        private DateTime _cachedExcelPasswordAt = DateTime.MinValue;

        /// <summary>
        /// 取得目前 Excel 密碼的鎖物件。
        /// </summary>
        /// <remarks>
        /// 用於多執行緒環境下同步存取 Excel 密碼快取，確保密碼取得與更新的執行緒安全。
        /// </remarks>
        public object PwdLock => _pwdLock;

        /// <summary>
        /// 取得目前快取的 Excel 密碼。
        /// </summary>
        /// <remarks>
        /// 主要用於單元測試或診斷，避免直接存取 private 欄位。
        /// </remarks>
        public string CachedExcelPassword => _cachedExcelPassword;

        /// <summary>
        /// 取得目前快取的 Excel 密碼取得時間（UTC）。
        /// </summary>
        /// <remarks>
        /// 用於判斷密碼快取是否過期。
        /// </remarks>
        public DateTime CachedExcelPasswordAt => _cachedExcelPasswordAt;
        /// <summary>
        /// Excel 密碼提供者 (預設實作為 ExcelPasswordProvider)，用於將密碼取得責任從表單移出。
        /// </summary>
        private readonly IExcelPasswordProvider _excelPasswordProvider = new ExcelPasswordProvider();

        /// <summary>
        /// 儲存最近一次從 Excel 偵測到的「隱藏欄位」標頭名稱清單（保留原始標頭字串）。
        /// /// </summary>
        /// <remarks>
        /// 用於在資料綁定之後做後備隱藏，避免 streaming 路徑或比對誤差導致隱藏欄仍顯示在 UI。
        /// 內容為 Excel 標記為 Hidden 的欄位標頭，供 <see cref="HideColumnsByHeaders(DataGridView)"/> 及相關 UI 邏輯參考。
        /// </remarks>
        private List<string>? _lastHiddenHeaders = new List<string>();

        // 畫面層的編輯鎖定：控制哪些欄位在 UI 上可編輯（不影響 Excel 檔案的保護狀態）
        // 注意：_isEditingLocked 的語義是「是否鎖定」
        // 初始值：true = 鎖定狀態（按鈕顯「解鎖」，匯出、存檔不可按）
        private bool _isEditingLocked = true;

        /// <summary>
        /// 儲存編輯前的儲存格值以便在驗證失敗時還原
        /// </summary>
        private Dictionary<string, object?> _dgvCellPrevValues = new Dictionary<string, object?>();
        #endregion


        /// <summary>
        /// Initializes a new instance of the <see cref="Form備料單匯入"/> class.
        /// 建構子：確保 <see cref="InitializeComponent"/> 被呼叫，並啟動必要的背景工作。
        /// </summary>
        /// <param name="excelService">可選的 Excel 服務，允許外部注入以便測試或替換實作，可為 <see langword="null"/>。</param>
        /// <param name="userNotifier">可選的使用者通知器，用於顯示訊息。若為 null 則建立預設 WinForms 實作。</param>
        #region Constructor & Initialization
        /// <summary>
        /// 建構子：初始化表單元件並設定必要的服務。
        /// - 參數 <c>excelService</c> 為選用的 Excel 服務（允許 null，用於測試或替換實作），目前以 dynamic 接受以維持相容性。
        /// - 參數 <c>userNotifier</c> 為顯示訊息的使用者通知器，若為 null 則建立預設的 WinForms 實作。
        /// 此建構子僅進行元件初始化與欄位預設值設定，避免執行長時間工作的同步作業。
        /// </summary>
        /// <param name="excelService">可選的 Excel 服務實作（允許 null）。</param>
        /// <param name="userNotifier">用於顯示訊息的使用者通知器，若為 null 會建立預設實作。</param>
        public Form備料單匯入(dynamic? excelService = null, IUserNotifier? userNotifier = null)
        {
            InitCommon(excelService, userNotifier);
        }

        /// <summary>
        /// 建構子（接受強型別 IExcelService），會將 typed 實作同時注入到 dynamic 欄位以保留相容行為。
        /// </summary>
        /// <param name="excelService">強型別的 Excel service 實作</param>
        /// <param name="userNotifier">使用者通知器</param>
        public Form備料單匯入(IExcelService? excelService, IUserNotifier? userNotifier = null)
        {
            // 呼叫共用初始化（使用 dynamic 參數以維持與舊有 dynamic 路徑相容）
            InitCommon((dynamic?)excelService, userNotifier);
            try { _typedExcelService = excelService; } catch { }
        }

        /// <summary>
        /// 共用的建構子初始化邏輯，從兩個建構子抽出以避免在建構函式初始設定式使用 dynamic 導致 CS1975。
        /// </summary>
        /// <param name="excelService">可為 dynamic 或強型別實作（在呼叫端轉型）。</param>
        /// <param name="userNotifier">使用者通知器</param>
        private void InitCommon(dynamic? excelService = null, IUserNotifier? userNotifier = null)
        {
            InitializeComponent();
            // Only assign dynamic excel service if non-null to avoid nullable assignment warnings
            try { if (excelService != null) _excelService = excelService; } catch { }
            _userNotifier = userNotifier ?? new WinFormsUserNotifier(this);
            _cachedExcelPassword = GetExcelPassword();
            // 優先使用 Login.User_No（若專案提供此 static），否則 fallback 到 Environment.UserName
            operatorName = !string.IsNullOrWhiteSpace(Login.User_No) ? Login.User_No :
                               !string.IsNullOrWhiteSpace(Login.User_name) ? Login.User_name :
                               Environment.UserName;
            // 註冊 DataGridView 編輯相關事件，加入防呆與驗證
            try
            {
                if (this.dgv備料單 != null)
                {
                    this.dgv備料單.CellBeginEdit += Dgv備料單_CellBeginEdit;
                    this.dgv備料單.CellValidating += Dgv備料單_CellValidating;
                    this.dgv備料單.CellEndEdit += Dgv備料單_CellEndEdit;
                }
            }
            catch { }
        }

        /// <summary>
        /// 允許在建構後注入或切換 IExcelService 實作（同時保留舊有 dynamic 欄位以維持相容）。
        /// </summary>
        /// <param name="svc">IExcelService 實作，若為 null 則會清除 typed 實作。</param>
        public void SetExcelService(IExcelService? svc)
        {
            try { _typedExcelService = svc; } catch { }
            try
            {
                if (svc != null) _excelService = svc as dynamic;
                else _excelService = null;
            }
            catch { }
        }
        #endregion

        #region Fields
        /// <summary>
        /// 記錄可見的 Excel 欄位索引（1-based）。
        /// 用於追蹤目前顯示於 DataGridView 的 Excel 欄位位置，便於資料對應與欄位操作。
        /// </summary>
        // visibleColumnIndexes removed (was unused / shadowed by local variables) to clean up unused fields.
        /// <summary>
        /// 記錄上次比對到的 DataGridViewRow 集合（用於數量更新）。
        /// 當使用者輸入料號並比對成功時，將所有符合條件的資料列存於此集合，
        /// 以便後續進行數量累加、檢查或其他操作。
        /// </summary>
        private List<DataGridViewRow> _lastMatchedRows = new List<DataGridViewRow>();
        #endregion

        #region Misc Helpers
        /// <summary>
        /// 判斷 DataTable 是否有實際異動（有任何資料列狀態為 Added/Modified/Deleted 即視為有異動）
        /// </summary>
        /// <param name="dt">要檢查的 DataTable</param>
        /// <returns>有異動則回傳 true，否則 false</returns>
        private bool DataTableHasRealChanges(DataTable dt)
        {
            if (dt == null) return false;
            foreach (DataRow row in dt.Rows)
            {
                if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified || row.RowState == DataRowState.Deleted)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// DataGridView 大小變更時的處理器，供外部或內部調整欄寬/版面用。
        /// 此處採用去抖或直接呼叫欄寬調整的 helper，避免頻繁重繪造成效能問題。
        /// </summary>
        /// <param name="sender">事件來源（通常為 DataGridView）</param>
        /// <param name="e">事件參數</param>
        private void Dgv_SizeChanged(object sender, EventArgs e)
        {
            // 可根據需求調整 DataGridView 欄寬或其他行為
            try { AutoSizeColumnsFillNoHorizontalScroll(this.dgv備料單); } catch { }
        }

        /// <summary>
        /// 檢查指定的 Office Excel 檔案是否為加密封裝（僅針對 .xlsx/.xlsm 格式有效）。
        /// 會以快速 ASCII 掃描方式檢查檔案內容是否包含 OOXML 加密特徵字串。
        /// </summary>
        /// <param name="path">要檢查的 Excel 檔案完整路徑。</param>
        /// <returns>若檔案為加密封裝則回傳 true，否則回傳 false。</returns>
        private bool IsOfficeFileEncrypted(string path)
        {
            // 檢查 Office 檔案是否為加密封裝
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return false;
            var ext = Path.GetExtension(path)?.ToLowerInvariant() ?? string.Empty;
            try
            {
                if (ext == ".xlsx" || ext == ".xlsm")
                {
                    // 避免相依 ZipFile：改用快速 ASCII 掃描檔案內是否出現 OOXML 加密特徵字串
                    try
                    {
                        const int MaxScanBytes = 4 * 1024 * 1024; // 最多掃描前 4MB
                        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            var length = (int)Math.Min(fs.Length, MaxScanBytes);
                            var buffer = new byte[length];
                            int read = fs.Read(buffer, 0, length);
                            if (read > 0)
                            {
                                var text = System.Text.Encoding.ASCII.GetString(buffer, 0, read);
                                if (text.IndexOf("EncryptedPackage", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                                if (text.IndexOf("EncryptionInfo", StringComparison.OrdinalIgnoreCase) >= 0) return true;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 檢查目前是否有有效的 Excel 檔案可供操作，並在不合法時顯示提示。
        /// </summary>
        /// <returns>若有可用 Excel 檔案則回傳 true，否則 false。</returns>
        private bool CheckExcelAvailable()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(this.currentExcelPath) || !File.Exists(this.currentExcelPath))
                {
                    try { MessageBox.Show("請先匯入或選擇有效的 Excel 檔案。", "無檔案", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                    return false;
                }
            }
            catch { return false; }
            return true;
        }

        /// <summary>
        /// 暫停指定控制項的重繪（畫面更新）。
        /// </summary>
        /// <remarks>
        /// 透過傳送 Win32 訊息 <c>WM_SETREDRAW</c> 並將 <c>wParam</c> 設為 <c>0</c>
        /// 來暫停指定控制項的重繪，通常用於需要大量更新 UI 而暫時避免重複重繪以提升效能的情境。
        /// 本方法會靜默吃掉任何例外以維持與原有程式行為一致。
        /// 呼叫對應的復原方法請使用 <see cref="EndUpdate(System.Windows.Forms.Control)"/>。
        /// </remarks>
        /// <param name="ctl">要暫停重繪的控制項。若為 <c>null</c> 或控制項尚未建立 handle，則會被忽略。</param>
        private void BeginUpdate(Control ctl)
        {
            try
            {
                SendMessage(ctl.Handle, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
            }
            catch { }
        }

        /// <summary>
        /// 恢復指定控制項的重繪並強制重新整理畫面。
        /// </summary>
        /// <remarks>
        /// 透過傳送 Win32 訊息 <c>WM_SETREDRAW</c> 並將 <c>wParam</c> 設為 <c>1</c>
        /// 來恢復控制項的重繪，之後會呼叫 <see cref="System.Windows.Forms.Control.Refresh()"/> 強制重繪。
        /// 本方法會靜默吃掉任何例外以維持與原有程式行為一致。
        /// 對應的暫停方法為 <see cref="BeginUpdate(System.Windows.Forms.Control)"/>。
        /// </remarks>
        /// <param name="ctl">要恢復重繪的控制項。若為 <c>null</c> 或尚未建立 handle，則會被忽略。</param>
        private void EndUpdate(Control ctl)
        {
            try
            {
                SendMessage(ctl.Handle, WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
                ctl.Refresh();
            }
            catch { }
        }

        /// <summary>
        /// 使用反射設定控制項的 DoubleBuffered 屬性以改善大量繪製時的閃爍行為。
        /// </summary>
        /// <remarks>
        /// Windows Forms 的 <c>Control.DoubleBuffered</c> 為受保護的成員，無法直接於外部存取。
        /// 此方法使用 reflection 以非公開綁定旗標取得該屬性並設定值。
        /// 因為使用 reflection 具有相容性風險與例外可能性，本方法會靜默吃掉任何例外以維持原有程式行為。
        /// </remarks>
        /// <param name="ctl">目標控制項，若為 <c>null</c> 則忽略。</param>
        /// <param name="enabled">是否啟用 DoubleBuffered。</param>
        private void SetDoubleBuffered(Control ctl, bool enabled)
        {
            try
            {
                var prop = typeof(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                if (prop != null) prop.SetValue(ctl, enabled, null);
            }
            catch { }
        }

        /// <summary>
        /// 安全地執行指定的操作，將常見的 try/catch 包裝成單一 helper 以減少重複程式。
        /// </summary>
        /// <remarks>
        /// 此方法會在內部捕捉並忽略任何例外；用於不希望例外影響主流程的非關鍵性作業。
        /// 若呼叫端需要例外資訊以進行處理，請勿使用此 helper。
        /// </remarks>
        /// <param name="action">要執行的操作，可為 <c>null</c>（null 會被忽略）。</param>
        private void SafeAction(Action? action)
        {
            if (action == null) return;
            try { action(); } catch { }
        }

        #endregion

        #region DataGridView Event Handlers
        /// <summary>
        /// 在儲存格進入編輯模式前，記錄該儲存格的目前值。
        /// </summary>
        /// <param name="sender">事件觸發來源，通常是 `dgv備料單`。</param>
        /// <param name="e">包含儲存格位置（列與欄）的 <see cref="DataGridViewCellCancelEventArgs"/>。</param>
        /// <remarks>
        /// - 方法會把目前儲存格的值儲存到私有字典 `_dgvCellPrevValues`，鍵格式為 "{row}_{column}"。
        /// - 目的是在需要時能還原編輯前的值（例如編輯失敗或需回退時）。
        /// - 實作以防禦式編碼為主：任何索引或存取錯誤都會被捕捉並安全處理，方法不會向外拋出例外。
        /// </remarks>
        private void Dgv備料單_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            try
            {
                var dgv = this.dgv備料單;
                if (dgv == null) return;
                var key = e.RowIndex + "_" + e.ColumnIndex;
                try { _dgvCellPrevValues[key] = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value; } catch { _dgvCellPrevValues[key] = null; }
            }
            catch { }
        }

        /// <summary>
        /// 在儲存格驗證階段執行輸入有效性檢查，但不會變更儲存格的顏色或視覺樣式。
        /// </summary>
        /// <param name="sender">事件觸發來源，通常是 `dgv備料單`。</param>
        /// <param name="e">包含儲存格位置與可取消狀態的 <see cref="DataGridViewCellValidatingEventArgs"/>。</param>
        /// <remarks>
        /// - 此事件處理程序僅用於阻止無效資料被接受（例如格式錯誤或不允許的值），而不負責 UI 樣式變更。
        /// - 針對顏色或錯誤標示的更新，請參考 <c>CellEndEdit</c> 與 <c>CellFormatting</c> 的實作。
        /// - 方法採防禦式編碼：任何可能的索引或存取錯誤都會被捕捉處理，避免拋出未處理的例外。
        /// </remarks>
        private void Dgv備料單_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            try
            {
                // CellValidating 事件只用於防止無效資料進入編輯狀態
                // 不在此處改變顏色，避免因為點選就移除紅色標記
                var dgv = this.dgv備料單;
                if (dgv == null) return;
                int shippedCol = FindColumnIndexByNames(new[] { "實發數量", "發料數量" });
                if (shippedCol < 0) return;

                if (e.ColumnIndex == shippedCol)
                {
                    // 驗證邏輯在 CellEndEdit 中統一處理
                }
            }
            catch { }
        }


        #endregion


        /// <summary>
        /// 處理 `Dgv備料單` 的 <c>CellEndEdit</c> 事件。
        /// 在使用者結束儲存格編輯後執行欄位驗證、顏色標示（如短缺標紅）、還原先前值（如超出需求時）以及更新已修改狀態。
        /// </summary>
        /// <param name="sender">事件來源，通常為觸發事件的 <see cref="System.Windows.Forms.DataGridView"/>。</param>
        /// <param name="e">包含被編輯儲存格之列與欄索引的 <see cref="System.Windows.Forms.DataGridViewCellEventArgs"/>。</param>
        /// <remarks>
        /// - 主要檢查編輯欄位是否為發料／實發數量欄，若不是則跳過。
        /// - 會嘗試解析數值並與相同料號的累計發出數量及需求數量比較。
        /// - 若發出數量超出需求，會顯示警告並嘗試將值還原為編輯前的值。
        /// - 若發出數量未達需求，會標示儲存格為紅色，並在內部集合中保留該料號以便後續格式化時保留紅色標記。
        /// - 若編輯成功且為有效數值，會將儲存格標為白色並將表單標記為已修改（_isDirty）。
        /// - 方法實作會使用 `_dgvCellPrevValues` 來記錄與還原編輯前的值，並以安全的 try/catch 保護不應拋出例外。
        /// </remarks>
        private void Dgv備料單_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                var dgv = this.dgv備料單;
                if (dgv == null) return;

                // 清除列錯誤提示
                int shippedCol = FindColumnIndexByNames(new[] { "實發數量", "發料數量" });
                if (shippedCol < 0)
                {
                    return;
                }

                if (e.ColumnIndex != shippedCol)
                {
                    return;
                }

                // 解析目前編輯後的值（使用共用 helper）
                    decimal curVal = 0m;
                    string cellValue = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? string.Empty;
                try { var v = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString() ?? string.Empty; TextParsing.TryParseDecimalValue(v, out curVal); } catch { curVal = 0m; }

                // 取得料號 key（嘗試昶亨料號, 客戶料號）
                string matKey = string.Empty;
                int chCol = FindColumnIndexByNames(new[] { "昶亨料號" });
                int custCol = FindColumnIndexByNames(new[] { "客戶料號" });
                try
                {
                    if (chCol >= 0 && chCol < dgv.Columns.Count) matKey = dgv.Rows[e.RowIndex].Cells[chCol].Value?.ToString() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(matKey) && custCol >= 0 && custCol < dgv.Columns.Count) matKey = dgv.Rows[e.RowIndex].Cells[custCol].Value?.ToString() ?? string.Empty;
                }
                catch { }
                matKey = TextParsing.NormalizeMaterialKey(matKey);

                // 計算相同料號的累計發出數量
                decimal sum = 0m;
                try
                {
                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        if (row == null || row.IsNewRow) continue;
                        string rowKey = string.Empty;
                        try
                        {
                            if (chCol >= 0 && chCol < row.Cells.Count) rowKey = row.Cells[chCol].Value?.ToString() ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(rowKey) && custCol >= 0 && custCol < row.Cells.Count) rowKey = row.Cells[custCol].Value?.ToString() ?? string.Empty;
                        }
                        catch { }
                        if (TextParsing.NormalizeMaterialKey(rowKey) != matKey) continue;
                        try { if (row.Cells.Count > shippedCol) { var sv = row.Cells[shippedCol].Value?.ToString() ?? string.Empty; if (TextParsing.TryParseDecimalValue(sv, out decimal v)) sum += v; } } catch { }
                    }
                }
                catch { }

                // 【修正】取得目前編輯的儲存格及其狀態 - 先做這個，與是否有需求欄無關
                var currentCell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                bool isEmptyValue = string.IsNullOrEmpty(currentCell.Value?.ToString());
                bool isValidNumber = decimal.TryParse(currentCell.Value?.ToString(), out _);

                // ...existing code...

                // 【修正】第一步：先檢查編輯後的值是否為空 - 空值設為白色
                if (isEmptyValue)
                {
                    currentCell.Style.BackColor = Color.White;

                    // 移除 _preservedRedKeys 中的該料號，避免後續 CellFormatting 再次標紅
                    try
                    {
                        if (!string.IsNullOrEmpty(matKey) && _preservedRedKeys != null)
                        {
                            _preservedRedKeys.Remove(matKey);
                        }
                    }
                    catch { }
                }
                else
                {
                    // 第二步：如果不是空值，才檢查是否超出限制
                    // 取得需求欄位（依檔案類型呈現不同欄位名稱）
                    int demandCol = -1;
                    try
                    {
                        var ext = string.Empty;
                        try { ext = Path.GetExtension(this.currentExcelPath ?? string.Empty).ToLowerInvariant(); } catch { }
                        if (ext == ".xlsm")
                        {
                            demandCol = FindColumnIndexByNames(new[] { "需求數量" });
                        }
                        else
                        {
                            // 其他 Excel 檔案，優先找應領欄位
                            demandCol = FindColumnIndexByNames(new[] { "應領數量" });
                        }
                    }
                    catch { }

                    // 若找到需求欄且可解析，進行比較
                    if (demandCol >= 0 && demandCol < dgv.Columns.Count)
                    {
                        decimal demandVal = 0m;
                        bool demandParsed = false;
                        try { var s = dgv.Rows[e.RowIndex].Cells[demandCol].Value?.ToString() ?? string.Empty; demandParsed = TextParsing.TryParseDecimalValue(s, out demandVal); } catch { demandParsed = false; }

                        if (isValidNumber && demandParsed)
                        {
                            // 【修正邏輯】檢查是否超過需求（只有 > 時才警告，= 時不警告）
                            if (sum > demandVal)
                            {
                                // 【超出】發出數量超過需求 → 警告 + 還原舊值
                                string msg = "發出的數量已超出應領數量";
                                try
                                {
                                    var ext = Path.GetExtension(this.currentExcelPath ?? string.Empty).ToLowerInvariant();
                                    if (ext == ".xlsm") msg = "發出的數量已超出需求數量";
                                }
                                catch { }

                                // 顯示警告訊息
                                SafeShowMessage(msg, "數量超出", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                // 還原到編輯前的值（若有先前值）
                                var key = e.RowIndex + "_" + e.ColumnIndex;
                                if (_dgvCellPrevValues.TryGetValue(key, out var prev))
                                {
                                    var valueToRestore = prev;
                                    try
                                    {
                                        // 嘗試在必要時暫時解除對應 DataColumn 的 ReadOnly 屬性
                                        var col = dgv.Columns[e.ColumnIndex];
                                        var propName = col != null ? (!string.IsNullOrEmpty(col.DataPropertyName) ? col.DataPropertyName : col.Name) : null;
                                        if (dgv.DataSource is DataTable dt && !string.IsNullOrEmpty(propName) && dt.Columns.Contains(propName))
                                        {
                                            var dc = dt.Columns[propName];
                                            bool origReadOnly = dc.ReadOnly;
                                            try
                                            {
                                                if (origReadOnly) dc.ReadOnly = false;
                                                dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = valueToRestore;
                                            }
                                            finally
                                            {
                                                try { dc.ReadOnly = origReadOnly; } catch { }
                                            }
                                        }
                                        else
                                        {
                                            dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = valueToRestore;
                                        }
                                    }
                                    catch
                                    {
                                        // fallback: 直接寫到 cell
                                        try { dgv.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = valueToRestore; } catch { }
                                    }
                                }
                            }
                            else if (sum < demandVal)
                            {
                                // 【短缺】發出數量未達需求 → 標紅色，並加入 _preservedRedKeys
                                currentCell.Style.BackColor = Color.Red;

                                // 將料號加入 _preservedRedKeys，以便後續保留紅色標記
                                try
                                {
                                    if (!string.IsNullOrEmpty(matKey) && _preservedRedKeys != null)
                                    {
                                        _preservedRedKeys.Add(matKey);
                                    }
                                }
                                catch { }
                            }
                            else
                            {
                                // 【相等】發出數量相等需求 → 設為白色，並從 _preservedRedKeys 中移除
                                currentCell.Style.BackColor = Color.White;

                                // 移除 _preservedRedKeys 中的該料號，避免後續 CellFormatting 再次標紅
                                try
                                {
                                    if (!string.IsNullOrEmpty(matKey) && _preservedRedKeys != null)
                                    {
                                        _preservedRedKeys.Remove(matKey);
                                    }
                                }
                                catch { }
                            }
                        }
                        else if (isValidNumber)
                        {
                            // 編輯後為有效數字但無法解析需求值 → 移除紅色標記
                            currentCell.Style.BackColor = Color.White;
                        }
                    }
                    else if (isValidNumber)
                    {
                        // 找不到需求欄位，但編輯值是有效數字 → 移除紅色標記
                        currentCell.Style.BackColor = Color.White;
                    }
                }

                // 標示為已修改
                try { if (!_suspendDirtyMarking) _isDirty = true; } catch { }
            }
            catch { }
        }


        /// <summary>
        /// 保存目前表單上按鈕／ToolStripItem 的 Enabled 狀態，並將它們暫時停用。
        /// 呼叫此方法可避免長時間背景工作期間使用者誤操作。
        /// 還原工作請使用 <see cref="RestoreAllButtons"/>。
        /// </summary>
        #region UI Helpers
        /// <summary>
        /// 儲存目前資料並停用所有按鈕。
        /// </summary>
        /// <remarks>
        /// 此方法會先執行資料儲存動作，然後將所有互動按鈕設為不可用，避免重複操作。
        /// </remarks>
        private void SaveAndDisableAllButtons()
        {
            UiHelpers.SaveAndDisableAllButtons(this);
        }

        #region Excel Operations

        /// <summary>
        /// 在匯入後更新 Excel：建立或更新工作表「記錄」，設定標題樣式並根據檔案類型鎖定欄位或工作表。
        /// 使用 COM 操作，並盡量保留原有格式/公式/巨集。密碼使用欄位 `_cachedExcelPassword`。
        /// </summary>
        /// <param name="excelPath"></param>
        /// <param name="token"></param>
        private async Task UpdateExcelAfterImportAsync(string excelPath, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath)) return;

            await Task.Run(() =>
            {
                Excel.Application? xlApp = null; Excel.Workbook? wb = null; Excel.Worksheet? wsMain = null; Excel.Worksheet? wsLog = null; Excel.Range? rng = null;
                try
                {
                    // 使用 ExcelService 取代直接 new Excel.Application
                    System.Data.DataTable? dataTable = null;
                    if (!string.IsNullOrEmpty(currentExcelPath))
                    {
                        if (_typedExcelService is not null)
                        {
                            dataTable = _typedExcelService.LoadFirstWorksheetToDataTable(currentExcelPath);
                        }
                        else if (_excelService is not null)
                        {
                            dataTable = ((IExcelService)_excelService).LoadFirstWorksheetToDataTable(currentExcelPath);
                        }
                        else
                        {
                            xlApp = new Excel.Application { DisplayAlerts = false, Visible = false };
                        }
                    }
                    else
                    {
                        // 檔案路徑為 null 或空字串，無法載入 Excel
                        // TODO: 可加入錯誤提示或例外處理
                    }
                    // 使用 ReadOnly = false 以便回寫，保留巨集時不要改變檔案格式
                    wb = xlApp.Workbooks.Open(excelPath, ReadOnly: false, Password: string.IsNullOrEmpty(_cachedExcelPassword) ? Type.Missing : (object)_cachedExcelPassword);

                    // 取得總表
                    try { wsMain = wb.Worksheets["總表"] as Excel.Worksheet; } catch { wsMain = null; }
                    if (wsMain == null)
                    {
                        // 若找不到總表則使用第一個工作表
                        try { wsMain = wb.Worksheets[1] as Excel.Worksheet; } catch { wsMain = null; }
                        if (wsMain == null) return;
                    }

                    // 取得或建立「記錄」工作表，放在總表旁邊（確保放在後面）
                    // 宣告 creationError 在外層以便後續診斷區塊可存取
                    string? creationError = null;
                    try
                    {
                        // 1) 嘗試以不區分大小寫與前後空白的方式在現有工作表中尋找
                        foreach (Excel.Worksheet sh in wb.Worksheets)
                        {
                            try
                            {
                                var name = (sh.Name ?? string.Empty).ToString().Trim();
                                if (string.Equals(name, "記錄", StringComparison.OrdinalIgnoreCase))
                                {
                                    wsLog = sh;
                                    break;
                                }
                            }
                            catch { }
                        }

                        // 2) 若未找到，則嘗試新增一張工作表放在總表後（加強容錯）
                        if (wsLog == null)
                        {
                            try
                            {
                                // 若 workbook 被保護（結構保護）可能無法新增工作表，嘗試先解除保護
                                try
                                {
                                    if (!string.IsNullOrEmpty(_cachedExcelPassword))
                                    {
                                        try { wb.Unprotect(_cachedExcelPassword); } catch (Exception ex) { creationError = ex.ToString(); }
                                    }
                                    else
                                    {
                                        try { wb.Unprotect(Type.Missing); } catch (Exception ex) { creationError = ex.ToString(); }
                                    }
                                }
                                catch (Exception ex) { creationError = ex.ToString(); }

                                // 直接使用命名參數加強可讀性與相容性：在 wsMain 後新增一張工作表
                                try
                                {
                                    try { wsLog = wb.Worksheets.Add(After: wsMain) as Excel.Worksheet; } catch (Exception ex) { creationError = ex.ToString(); wsLog = null; }

                                    // 如果上面沒成功，嘗試用 xlApp.Worksheets.Add
                                    if (wsLog == null)
                                    {
                                        try { wsLog = xlApp.Worksheets.Add(After: wsMain) as Excel.Worksheet; } catch (Exception ex) { creationError = creationError ?? ex.ToString(); wsLog = null; }
                                    }

                                    // 再不行，嘗試 wb.Sheets.Add() 然後 Move 到 wsMain 之後
                                    if (wsLog == null)
                                    {
                                        try
                                        {
                                            var sh = wb.Sheets.Add() as Excel.Worksheet;
                                            if (sh != null)
                                            {
                                                try { sh.Move(After: wsMain); } catch (Exception ex) { creationError = creationError ?? ex.ToString(); }
                                                wsLog = sh;
                                            }
                                        }
                                        catch (Exception ex) { creationError = creationError ?? ex.ToString(); wsLog = null; }
                                    }

                                    // 若仍失敗，嘗試複製最後一張工作表作為 fallback，並放在總表之後
                                    if (wsLog == null)
                                    {
                                        try
                                        {
                                            var last = wb.Worksheets[wb.Worksheets.Count] as Excel.Worksheet;
                                            if (last != null)
                                            {
                                                last.Copy(After: wsMain);
                                                // 複製後，新的工作表位於總表之後，嘗試取得它
                                                try { wsLog = wb.Worksheets[wsMain.Index + 1] as Excel.Worksheet; } catch (Exception ex) { creationError = creationError ?? ex.ToString(); wsLog = null; }
                                            }
                                        }
                                        catch (Exception ex) { creationError = creationError ?? ex.ToString(); wsLog = null; }
                                    }

                                    // 確保如果成功建立，將其顯示並移動到總表後方
                                    if (wsLog != null)
                                    {
                                        try { wsLog.Visible = Excel.XlSheetVisibility.xlSheetVisible; } catch (Exception ex) { creationError = creationError ?? ex.ToString(); }
                                        try { wsLog.Move(After: wsMain); } catch (Exception ex) { creationError = creationError ?? ex.ToString(); }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    creationError = creationError ?? ex.ToString();
                                }

                                // 嘗試命名為 "記錄"，若有衝突則加數字後綴
                                if (wsLog != null)
                                {
                                    string baseName = "記錄";
                                    string tryName = baseName;
                                    int suffix = 1;
                                    bool named = false;
                                    while (!named)
                                    {
                                        try
                                        {
                                            wsLog.Name = tryName;
                                            named = true;
                                        }
                                        catch
                                        {
                                            tryName = baseName + "_" + suffix.ToString();
                                            suffix++;
                                            if (suffix > 50) { named = true; } // give up eventually
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // 若仍失敗，嘗試用最後建立的 worksheet 作為 fallback
                                try { wsLog = wb.Worksheets[wb.Worksheets.Count] as Excel.Worksheet; } catch { wsLog = null; }
                            }
                        }
                        else
                        {
                            // 如果已存在，確保為可見並移動到總表後方以確保位置
                            try { wsLog.Visible = Excel.XlSheetVisibility.xlSheetVisible; } catch { }
                            try { wsLog.Move(Type.Missing, wsMain); } catch { }
                        }
                    }
                    catch { wsLog = null; }

                    // 設定標題列（第一列）並應用樣式 - 確保 wsLog 不為 null
                    if (wsLog != null)
                    {
                        try
                        {
                            var headers = new[] { "刷入時間", "料號", "數量", "操作者" };

                            // 使用一個 headerRange 來設定整列樣式（明確建立值再設定樣式）
                            Excel.Range? headerRange = null;
                            try
                            {
                                // 先寫入標題文字（批次寫入，避免逐格 COM 呼叫）
                                try
                                {
                                    var headerArr = new object[1, headers.Length];
                                    for (int i = 0; i < headers.Length; i++) headerArr[0, i] = headers[i];
                                    var start = wsLog.Cells[1, 1] as Excel.Range;
                                    var end = wsLog.Cells[1, headers.Length] as Excel.Range;
                                    headerRange = wsLog.Range[start, end] as Excel.Range;
                                    if (headerRange != null)
                                    {
                                        try { headerRange.Value2 = headerArr; }
                                        catch
                                        {
                                            // fallback: per-cell write if batch fails
                                            for (int i = 0; i < headers.Length; i++)
                                            {
                                                try { var cell = (wsLog.Cells[1, i + 1] as Excel.Range); if (cell != null) cell.Value2 = headers[i]; } catch { }
                                            }
                                        }
                                    }
                                }
                                catch { headerRange = wsLog.Range[wsLog.Cells[1, 1], wsLog.Cells[1, headers.Length]] as Excel.Range; }
                            }
                            catch { headerRange = null; }

                            // 樣式：標題置中、垂直置中、粗體
                            if (headerRange != null)
                            {
                                try { headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; } catch { }
                                try { headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; } catch { }
                                try { headerRange.Font.Bold = true; } catch { }
                            }

                            // 設定欄位內容對齊：刷入時間(左), 料號(左), 數量(右), 操作者(右)
                            try
                            {
                                var col1 = wsLog.Columns[1] as Excel.Range; if (col1 != null) col1.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            }
                            catch { }
                            try
                            {
                                var col2 = wsLog.Columns[2] as Excel.Range; if (col2 != null) col2.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                            }
                            catch { }
                            try
                            {
                                var col3 = wsLog.Columns[3] as Excel.Range; if (col3 != null) col3.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                            }
                            catch { }
                            try
                            {
                                var col4 = wsLog.Columns[4] as Excel.Range; if (col4 != null) col4.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                                try { if (col4 != null) col4.NumberFormat = "@"; } catch { }
                            }
                            catch { }

                            // 針對第一列（標題）設定特殊對齊：標題本身置中、但資料欄內容屬性上面已設定
                            try { if (headerRange != null) headerRange.WrapText = false; } catch { }

                            // 自動調整欄寬（整欄 AutoFit）- 先 AutoFit，若需要可在之後調整最小欄寬
                            for (int i = 1; i <= headers.Length; i++)
                            {
                                try { var col = wsLog.Columns[i] as Excel.Range; if (col != null) col.AutoFit(); } catch { }
                            }

                            try { if (headerRange != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(headerRange); } catch { }
                        }
                        catch { }
                    }

                    // 如果仍然無法建立 wsLog，紀錄原因並通知使用者（但不擲出例外以不中斷流程）
                    try
                    {
                        if (wsLog == null)
                        {
                            try
                            {
                                string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Automatic_Storage", "logs");
                                try { Directory.CreateDirectory(logPath); } catch { }
                                string fn = Path.Combine(logPath, $"CreateRecordSheetFail_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                                try
                                {
                                    var sb = new System.Text.StringBuilder();
                                    sb.AppendLine($"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                                    try { sb.AppendLine($"ExcelPath: {excelPath}"); } catch { }
                                    try { sb.AppendLine($"Workbook Worksheets Count: {wb?.Worksheets?.Count}"); } catch { }
                                    try { sb.AppendLine($"Workbook ReadOnly: {wb?.ReadOnly}"); } catch { }
                                    if (!string.IsNullOrEmpty(creationError))
                                    {
                                        sb.AppendLine("CreationError:");
                                        sb.AppendLine(creationError);
                                    }
                                    else
                                    {
                                        sb.AppendLine("CreationError: (none captured)");
                                    }

                                    // debug file write removed
                                }
                                catch { }

                                SafeBeginInvoke(this, new Action(() =>
                                {
                                    try { MessageBox.Show("警告：無法建立工作表 '記錄'。請確認工作簿是否受結構保護或無法新增工作表。已在使用者資料資料夾建立偵錯檔案。", "建立工作表失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                                }));
                            }
                            catch { }
                        }
                    }
                    catch { }

                    // 「記錄」工作表不需要篩選：若已有篩選則清除，並關閉 AutoFilterMode（確保無篩選功能）
                    if (wsLog != null)
                    {
                        try { if (wsLog.FilterMode) { try { wsLog.ShowAllData(); } catch { } } } catch { }
                        try { wsLog.AutoFilterMode = false; } catch { }
                    }

                    // 依副檔名決定要鎖定的欄位（支援同義詞），且「記錄」一律保護（不允許篩選）
                    string ext = Path.GetExtension(excelPath).ToLowerInvariant();
                    bool isXlsm = ext == ".xlsm";
                    string[] targetNames = isXlsm
                        ? new[] { "實發數量" }
                        : new[] { "發料數量" };

                    // 鎖定總表對應欄位（將該欄設為 Locked = true，其它欄設為 unlocked），並啟用「偵測到的標題列」AutoFilter
                    try
                    {
                        // 先嘗試解除保護，避免在保護狀態下無法變更鎖定屬性
                        try { wsMain.Unprotect(string.IsNullOrEmpty(_cachedExcelPassword) ? Type.Missing : (object)_cachedExcelPassword); } catch { }

                        // 先把整張總表所有儲存格解鎖，之後只鎖指定欄
                        try { (wsMain.Cells as Excel.Range).Locked = false; } catch { }

                        // 動態偵測標題列（偏好第 3 列，否則掃描前 10 列取最多非空的列）
                        Excel.Range? used = null;
                        int firstRow = 1, firstCol = 1, rowsCount = 0, colsCount = 0, lastRow = 0;
                        try
                        {
                            used = wsMain.UsedRange as Excel.Range;
                            if (used != null)
                            {
                                try { firstRow = used.Row; } catch { firstRow = 1; }
                                try { firstCol = used.Column; } catch { firstCol = 1; }
                                try { rowsCount = used.Rows.Count; } catch { rowsCount = 0; }
                                try { colsCount = used.Columns.Count; } catch { colsCount = 0; }
                                lastRow = firstRow + Math.Max(0, rowsCount - 1);
                            }
                        }
                        catch { }

                        int headerRow = firstRow;
                        try
                        {
                            int pref = firstRow + 2; // 偏好第 3 列
                            bool prefHasAny = false;
                            if (rowsCount > 0 && colsCount > 0 && pref <= lastRow)
                            {
                                int nonEmpty = 0;
                                for (int c = 0; c < colsCount; c++)
                                {
                                    try
                                    {
                                        var v = (wsMain.Cells[pref, firstCol + c] as Excel.Range)?.Value2;
                                        {
                                            var vs = v?.ToString();
                                            if (!string.IsNullOrWhiteSpace(vs)) nonEmpty++;
                                        }
                                    }
                                    catch { }
                                }
                                prefHasAny = nonEmpty > 0;
                            }
                            if (!prefHasAny)
                            {
                                int scanLimit = Math.Min(10, Math.Max(0, rowsCount));
                                int best = -1; int bestRow = firstRow;
                                for (int rOff = 0; rOff < scanLimit; rOff++)
                                {
                                    int r = firstRow + rOff;
                                    int cnt = 0;
                                    for (int c = 0; c < colsCount; c++)
                                    {
                                        try
                                        {
                                            var v = (wsMain.Cells[r, firstCol + c] as Excel.Range)?.Value2;
                                            {
                                                var vs = v?.ToString();
                                                if (!string.IsNullOrWhiteSpace(vs)) cnt++;
                                            }
                                        }
                                        catch { }
                                    }
                                    if (cnt > best) { best = cnt; bestRow = r; }
                                }
                                headerRow = bestRow;
                            }
                            else headerRow = pref;
                        }
                        catch { headerRow = firstRow; }

                        // 在標題列找出要鎖的欄位（同義詞）
                        int lockColAbs = -1;
                        if (colsCount > 0)
                        {
                            for (int i = 0; i < colsCount; i++)
                            {
                                try
                                {
                                    var hv = (wsMain.Cells[headerRow, firstCol + i] as Excel.Range)?.Value2;
                                    var hn = hv?.ToString()?.Trim() ?? string.Empty;
                                    if (string.IsNullOrEmpty(hn)) continue;

                                    // Flexible matching for shipped-like header:
                                    // 1) exact case-insensitive match
                                    // 2) sanitized exact match (remove non-alnum, lower)
                                    // 3) substring match of target name
                                    // 4) short-token heuristic: header contains short token (e.g. '發料'/'實發') AND a quantity indicator ('數' or '量' or 'qty')
                                    bool matched = false;
                                    try
                                    {
                                        if (targetNames.Any(t => string.Equals(t, hn, StringComparison.OrdinalIgnoreCase))) matched = true;
                                    }
                                    catch { }

                                    if (!matched)
                                    {
                                        try
                                        {
                                            var hnSan = SanitizeHeaderForMatch(hn);
                                            if (!string.IsNullOrEmpty(hnSan) && targetNames.Any(t => SanitizeHeaderForMatch(t) == hnSan)) matched = true;
                                        }
                                        catch { }
                                    }

                                    if (!matched)
                                    {
                                        try
                                        {
                                            if (!string.IsNullOrEmpty(hn))
                                            {
                                                var hnLower = (hn ?? string.Empty).ToLowerInvariant();
                                                if (targetNames.Any(t => !string.IsNullOrEmpty(t) && hnLower.Contains((t ?? string.Empty).ToLowerInvariant()))) matched = true;
                                            }
                                        }
                                        catch { }
                                    }

                                    if (!matched)
                                    {
                                        try
                                        {
                                            var shortTokens = new[] { "發料數量", "實發數量" };
                                            var quantityWords = new[] { "數量" };
                                            if (!string.IsNullOrEmpty(hn))
                                            {
                                                var hl = (hn ?? string.Empty).ToLowerInvariant();
                                                foreach (var s in shortTokens)
                                                {
                                                    if (hl.Contains(s))
                                                    {
                                                        if (quantityWords.Any(q => hl.Contains(q))) { matched = true; break; }
                                                    }
                                                }
                                            }
                                        }
                                        catch { }
                                    }

                                    if (matched)
                                    {
                                        lockColAbs = firstCol + i;
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }

                        // 鎖定指定欄（找不到時保守鎖整表）
                        if (lockColAbs > 0)
                        {
                            try
                            {
                                var endRowForLock = Math.Max(headerRow, lastRow);
                                // 確保至少鎖到 headerRow + 100 行，以避免在 lastRow 計算異常時只鎖到標頭導致鎖定無效
                                endRowForLock = Math.Max(endRowForLock, headerRow + 100);

                                var lockRange = wsMain.Range[wsMain.Cells[headerRow, lockColAbs], wsMain.Cells[endRowForLock, lockColAbs]] as Excel.Range;
                                if (lockRange != null)
                                {
                                    try { lockRange.Locked = true; } catch { }
                                    try { if (lockRange != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(lockRange); } catch { }
                                }

                                // 驗證若上面方法在某些 Excel 版本或檔案情況下未生效，嘗試以整欄方式強制設定
                                try
                                {
                                    var sampleCell = wsMain.Cells[headerRow, lockColAbs] as Excel.Range;
                                    bool lockedOk = false;
                                    try { lockedOk = sampleCell != null && (bool)sampleCell.Locked; } catch { lockedOk = false; }
                                    try { if (sampleCell != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(sampleCell); } catch { }
                                    if (!lockedOk)
                                    {
                                        try
                                        {
                                            var colRange = wsMain.Columns[lockColAbs] as Excel.Range;
                                            if (colRange != null) { try { colRange.Locked = true; } catch { } try { Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(colRange); } catch { } }
                                        }
                                        catch { }
                                    }
                                }
                                catch { }
                            }
                            catch { }
                        }
                        else
                        {
                            try { (wsMain.Cells as Excel.Range).Locked = true; } catch { }
                        }

                        // 將 AutoFilter 套用至偵測到的標題列
                        try
                        {
                            if (colsCount > 0)
                            {
                                var headerRangeMain = wsMain.Range[wsMain.Cells[headerRow, firstCol], wsMain.Cells[headerRow, firstCol + colsCount - 1]] as Excel.Range;
                                if (headerRangeMain != null)
                                {
                                    try { headerRangeMain.AutoFilter(1); } catch { }
                                    try { if (headerRangeMain != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(headerRangeMain); } catch { }
                                }
                            }
                        }
                        catch { }
                        finally
                        {
                            try { if (used != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(used); } catch { }
                        }

                        // 允許在保護模式下使用 AutoFilter（僅授權篩選）
                        try
                        {
                            var allowFiltering = true;
                            wsMain.Protect(string.IsNullOrEmpty(_cachedExcelPassword) ? Type.Missing : (object)_cachedExcelPassword,
                                AllowFiltering: allowFiltering);
                        }
                        catch { }
                    }
                    catch { }

                    // 保護「記錄」工作表（所有格式），不允許使用篩選
                    if (wsLog != null)
                    {
                        try
                        {
                            wsLog.Protect(string.IsNullOrEmpty(_cachedExcelPassword) ? Type.Missing : (object)_cachedExcelPassword, AllowFiltering: false);
                        }
                        catch { }
                        // Diagnostic: record whether wsLog was found and protection state for post-mortem
                        try
                        {
                            // diagnostic logging removed
                        }
                        catch { }
                    }

                    // 儲存工作簿，保留原有格式/巨集/公式
                    try
                    {
                        if (wb != null) wb.Save();
                    }
                    catch { /* 忽略儲存錯誤，但不中斷流程 */ }
                }
                catch { }
                finally
                {
                    try { if (rng != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(rng); } catch { }
                    try { if (wsLog != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(wsLog); } catch { }
                    try { if (wsMain != null) Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(wsMain); } catch { }
                    try { if (wb != null) { wb.Close(true); Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(wb); } } catch { }
                    try { if (xlApp != null) { xlApp.Quit(); Automatic_Storage.Utilities.ComInterop.ReleaseComObjectSafe(xlApp); } } catch { }
                }
            }, token);
        }

        /// <summary>
        /// 已移至中央 helper：呼叫 `ExcelInteropHelper.RemoveHiddenColumnsFromDataTable` 以維持一致性。
        /// 保留此方法為向後相容的 wrapper；任何錯誤會被靜默處理以不阻斷 UI 流程。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑</param>
        /// <param name="uiTable">將被修改的 DataTable（in-place）</param>
        private void RemoveHiddenColumnsFromDataTable(string excelPath, DataTable uiTable)
        {
            try
            {
                ExcelInteropHelper.RemoveHiddenColumnsFromDataTable(excelPath, uiTable);
            }
            catch { }
        }

        /// <summary>
        /// 還原先前被 <see cref="SaveAndDisableAllButtons"/> 停用的所有按鈕與工具列項目。
        /// </summary>
        /// <remarks>
        /// 呼叫此方法會使用內部的 <c>UiHelpers.RestoreAllButtons(Form備料單匯入)</c>
        /// 還原按鈕狀態，並考量成員欄位如 <c>_keepImportButtonDisabledUntilClose</c>
        /// 以決定是否保留匯入按鈕的停用狀態。此方法為 private，僅用於
        /// 表單內部在背景作業結束或關閉時恢復 UI。
        /// </remarks>
        private void RestoreAllButtons()
        {
            UiHelpers.RestoreAllButtons(this);
        }

        /// <summary>
        /// 顯示一個半透明遮罩，攔截所有使用者輸入並顯示等待文字。
        /// 適用於長時間的匯入/存檔作業，避免錯誤點擊造成表單假死。
        /// </summary>
        private void ShowOperationOverlay(string message = "處理中，請稍候...")
        {
            UiHelpers.ShowOperationOverlay(this, message);
        }

        /// <summary>
        /// 隱藏先前建立的等待提示遮罩，並嘗試回復 UI 可互動狀態。
        /// 此方法會安全地移除 overlay 並釋放資源。
        /// </summary>
        private void HideOperationOverlay()
        {
            UiHelpers.HideOperationOverlay(this);
            _operationOverlays = null;
        }

        /// <summary>
        /// 強制還原所有游標狀態到 Default，包含 Application.UseWaitCursor 與每個 Form/Control 的 Cursor 屬性。
        /// 用於確保在任何操作完成或被取消後，游標不會殘留在等待狀態。
        /// </summary>
        private void ResetAllCursors()
        {
            UiHelpers.ResetAllCursors(this);
        }

        /// <summary>
        /// 安全地設定 Control.Enabled。若 Control 需要跨執行緒操作，會使用 <see cref="SafeBeginInvoke"/> 確保在 UI 執行緒執行。
        /// 注意：此方法已不檢查全域鎖定旗標，改由 UpdateButtonStates() 統一管理按鈕啟用狀態。
        /// </summary>
        /// <param name="ctl">目標 Control</param>
        /// <param name="enabled">是否啟用</param>
        private void SetControlEnabledSafe(Control ctl, bool enabled)
        {
            try
            {
                if (ctl == null) return;
                if (ctl.InvokeRequired) SafeBeginInvoke(ctl, () => { try { ctl.Enabled = enabled; } catch { } }); else ctl.Enabled = enabled;
            }
            catch { }
        }

        /// <summary>
        /// 更新所有按鈕的啟用狀態，根據目前的操作旗標（_isImporting, _isExporting, _isSaving）。
        /// 此方法應在每個操作開始/結束時呼叫，以確保按鈕狀態維持一致。
        /// 設計目標：按鈕級鎖定 + 分離旗標 + 局部遮罩 = 最佳效能與 UX
        /// </summary>
        private void UpdateButtonStates()
        {
            try
            {
                // 【關鍵】委派給 UpdateMainButtonsEnabled()
                // 該方法已實現完整的按鈕邏輯，包括 _isEditingLocked 狀態、資料存在檢查等
                // 按鈕啟用邏輯：
                // - 匯入：隨時可按（只要沒有操作進行中）
                // - 匯出/存檔：按鈕顯「鎖定」（_isEditingLocked=false）時且有資料時可按
                // - 解鎖：有資料且沒有操作進行中時可按
                UpdateMainButtonsEnabled();

                // 返回按鈕：不在任何操作中時可按
                bool canReturn = !_isImporting && !_isExporting && !_isSaving;
                SetControlEnabledSafe(this.btn備料單返回, canReturn);
            }
            catch { }
        }

        /// <summary>
        /// 在 Control 尚未建立 Handle 時安全地執行委派（BeginInvoke/Invoke）的 helper。
        /// 若 target control 不可用會嘗試使用目前表單（this）作為 fallback，最後退回 Background 執行以避免例外。
        /// </summary>
        /// <param name="ctl">欲執行委派之 Control（可為 null）</param>
        /// <param name="action">要在 UI 執行緒執行的動作</param>
        private void SafeBeginInvoke(Control ctl, Action action)
        {
            try
            {
                if (action == null) return;
                Control target = ctl ?? this;
                // 若 target 已 disposal 或不可用，嘗試使用 this
                if (target == null || target.IsDisposed)
                {
                    target = this;
                }
                if (target != null && !target.IsDisposed)
                {
                    if (target.InvokeRequired)
                    {
                        // 若 handle 尚未建立，先嘗試建立 Handle
                        try { var h = target.Handle; } catch { }
                        if (!target.IsHandleCreated)
                        {
                            // 若仍未建立 handle，改用 BeginInvoke on threadpool via BeginInvoke on this if possible
                            try
                            {
                                if (this != null && !this.IsDisposed && this.IsHandleCreated)
                                {
                                    this.BeginInvoke(action);
                                    return;
                                }
                            }
                            catch { }
                            // 退而求其次：在 background thread 執行，不安全但避免拋例外
                            Task.Run(action);
                            return;
                        }

                        try { target.BeginInvoke(action); return; } catch { }
                    }
                    else
                    {
                        try { action(); return; } catch { }
                    }
                }

                // fallback: run on threadpool to avoid throwing
                Task.Run(action);
            }
            catch { }
        }

        #endregion

        /// <summary>
        /// 隱藏先前由 Excel 偵測為隱藏的標頭所對應的 DataGridView 欄位。
        /// </summary>
        /// <remarks>
        /// - 使用成員欄位 <c>_lastHiddenHeaders</c> 的標頭清單，透過 <c>SanitizeHeaderForMatch</c> 進行標頭正規化後比對。
        /// - 若欄位看起來像是「實發數量」或「發料數量」等關鍵欄位，則會保留（不會自動隱藏）。
        /// - 方法為防禦式實作：當傳入的 <paramref name="dgv"/> 為 <see langword="null"/> 或無欄位時，會安全回退且不拋出例外。
        /// </remarks>
        /// <param name="dgv">要檢查並隱藏欄位的 <see cref="System.Windows.Forms.DataGridView"/> 控制項。</param>
        private void HideColumnsByHeaders(DataGridView dgv)
        {
            try
            {
                if (dgv == null || dgv.Columns == null || dgv.Columns.Count == 0) return;
                HashSet<string>? sanSet = null;
                if (_lastHiddenHeaders != null && _lastHiddenHeaders.Count > 0)
                {
                    sanSet = new HashSet<string>(_lastHiddenHeaders.Select(SanitizeHeaderForMatch));
                }

                // 保護重要的 shipped-like 欄位，即使標頭為空或被偵測為 Hidden 也不要自動隱藏
                // Reuse the centralized sanitizer for column name normalization
                Func<string, string> NormalizeColName = SanitizeHeaderForMatch;
                var shippedSynonyms = new[] { "實發數量", "發料數量" };
                var shippedNorms = new HashSet<string>(shippedSynonyms.Select(NormalizeColName));

                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    bool hide = false;

                    // 若標頭為空，預設視為要隱藏（通常為 streaming 路徑補上但無內容的欄位）
                    var header = col.HeaderText ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(header)) { hide = true; }

                    // 如果有從 Excel 偵測到被隱藏的標頭，依標準化後比對
                    if (!hide && sanSet != null)
                    {
                        var san = SanitizeHeaderForMatch(header);
                        if (sanSet.Contains(san)) { hide = true; }
                    }

                    // 如果欄位看起來像是發料/實發等重要欄位，保留它（不要隱藏）
                    try
                    {
                        var candidate = !string.IsNullOrEmpty(col.DataPropertyName) ? col.DataPropertyName : col.Name;
                        var candNorm = NormalizeColName(candidate ?? string.Empty);
                        var headerNorm = NormalizeColName(header ?? string.Empty);
                        if (!string.IsNullOrEmpty(candNorm) && shippedNorms.Contains(candNorm))
                        {
                            hide = false;
                        }
                        else if (!string.IsNullOrEmpty(headerNorm) && shippedNorms.Contains(headerNorm))
                        {
                            hide = false;
                        }
                    }
                    catch { }

                    // 設定可見性
                    try { col.Visible = !hide; } catch { }
                }
            }
            catch { }
        }

        /// <summary>
        /// 重置 DataGridView 於重新綁定前的狀態。
        /// 重置會清除或初始化欄位屬性、事件註冊、工具提示、去抖定時器等，避免欄位殘留導致 UI 與 Excel 不一致。
        /// </summary>
        /// <remarks>
        /// - 此方法會安全地處理 `dgv備料單` 的欄位（例如依據欄位值判斷是否隱藏），並嘗試保留重要欄位（例如發料/實發數量類欄位）即使目前為空值。
        /// - 會初始化或註冊必要的事件處理器（CellFormatting、SizeChanged、Scroll、CellMouseMove 等）與輔助物件（ToolTip、Resize timer）。
        /// - 呼叫者應在 UI 執行緒上執行此方法；方法內部以 try/catch 包裝個別步驟以避免初始化小錯誤冒泡。
        /// </remarks>
        private void ResetGridBeforeBind()
        {
            try
            {
                if (this.dgv備料單 == null) return;
                try
                {
                    if (this.dgv備料單 == null || this.dgv備料單.Columns == null || this.dgv備料單.Columns.Count == 0) return;

                    // 取得 DataTable（若有）以便判斷整欄是否全空
                    DataTable? dt = null;
                    try { dt = this.dgv備料單.DataSource as DataTable; } catch { dt = null; }

                    HashSet<string>? sanSet = null;
                    if (_lastHiddenHeaders != null && _lastHiddenHeaders.Count > 0)
                    {
                        sanSet = new HashSet<string>(_lastHiddenHeaders.Select(SanitizeHeaderForMatch));
                    }

                    // Debug logging removed: production build should not write per-column debug traces

                    // Prepare a set of protected column name norms (shipped-like columns) to avoid hiding them
                    // Reuse the centralized sanitizer for column name normalization
                    Func<string, string> NormalizeColName = SanitizeHeaderForMatch;
                    var shippedSynonyms = new[] { "實發數量", "發料數量" };
                    var shippedNorms = new HashSet<string>(shippedSynonyms.Select(NormalizeColName));

                    foreach (DataGridViewColumn col in this.dgv備料單.Columns)
                    {
                        bool hide = false;
                        string reason = "";

                        // 若標頭為空，視為要隱藏（通常為 streaming 路徑補上但無內容的欄位）
                        var header = col.HeaderText ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(header)) { hide = true; reason = "Header is empty"; }

                        // 如果有從 Excel 偵測到被隱藏的標頭，依標準化後比對
                        if (!hide && sanSet != null)
                        {
                            var san = SanitizeHeaderForMatch(header);
                            if (sanSet.Contains(san)) { hide = true; reason = $"Header matches hiddenHeaders [{header}]"; }
                        }

                        // 若尚未決定，且綁定來源為 DataTable，可檢查該欄是否全部為空值／DBNull
                        if (!hide && dt != null)
                        {
                            var colName = !string.IsNullOrEmpty(col.DataPropertyName) ? col.DataPropertyName : col.Name;
                            if (!string.IsNullOrEmpty(colName) && dt.Columns.Contains(colName))
                            {
                                bool anyNonEmpty = false;
                                foreach (DataRow r in dt.Rows)
                                {
                                    try
                                    {
                                        if (r == null) continue;
                                        var v = r[colName];
                                        if (v != null && v != DBNull.Value)
                                        {
                                            var s = v?.ToString() ?? string.Empty;
                                            if (!string.IsNullOrWhiteSpace(s)) { anyNonEmpty = true; break; }
                                        }
                                    }
                                    catch { }
                                }
                                // 如果該欄位看起來像是 "發料/實發 數量" 一類的重要欄位，即使目前全為空也不要自動隱藏
                                try
                                {
                                    var normCol = NormalizeColName(colName);
                                    if (!anyNonEmpty && !string.IsNullOrEmpty(normCol) && shippedNorms.Contains(normCol))
                                    {
                                        // 不隱藏重要的 shipped 欄位
                                        hide = false;
                                        reason = "Preserved important shipped-like column despite being empty";
                                    }
                                    else
                                    {
                                        if (!anyNonEmpty) { hide = true; reason = "All values empty/DBNull"; }
                                    }
                                }
                                catch { if (!anyNonEmpty) { hide = true; reason = "All values empty/DBNull"; } }
                            }
                        }

                        // 設定可見性
                        try { col.Visible = !hide; } catch { }
                        // per-column debug entry removed
                    }
                    // 輸出 debug log 到檔案
                    // file-based debug dump removed
                }
                catch { }
                // 依需求：移除背景回寫佇列與背景工作者，僅允許「匯入」與「存檔」進行 Excel 回寫
                // 確保 DataGridView 隨窗體放大/縮小（若 Anchor 無效，嘗試 Dock.Fill）
                try { this.dgv備料單.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right; } catch { }
                try { if (this.dgv備料單.Dock != DockStyle.Fill) this.dgv備料單.Dock = DockStyle.Fill; } catch { }
                // 強制設為 FullRowSelect，確保選取 cell 會選整列
                try { this.dgv備料單.SelectionMode = DataGridViewSelectionMode.FullRowSelect; } catch { }
                // Ensure selection visuals are hidden and no cell is focused
                try { this.BeginInvoke(new Action(() => { try { HideSelectionInGrid(this.dgv備料單); } catch { } })); } catch { }
                // 初次顯示先不進行自動欄寬，避免在後續重新綁定時重複計算造成延遲；
                // 綁定完成後會於 DataBindingComplete 統一調整
                //try { AutoSizeColumnsFillNoHorizontalScroll(this.dgv備料單); } catch { }
                // 初始化 Resize 去抖 timer（使用具名 handler 以便解除註冊）
                try
                {
                    _resizeTimer = new System.Windows.Forms.Timer();
                    _resizeTimer.Interval = 200; // ms
                    _resizeTimer.Tick += ResizeTimer_Tick;

                    // 在窗體大小改變時啟動去抖
                    this.Resize += Form_Resize;
                    // 若 DataGridView 大小改變也觸發
                    try { this.dgv備料單.SizeChanged += Dgv_SizeChanged; } catch { }

                    // 初始化 cell tooltip 並註冊事件
                    try
                    {
                        _cellToolTip = new System.Windows.Forms.ToolTip { ShowAlways = true, AutoPopDelay = 8000, InitialDelay = 300, ReshowDelay = 100 };
                        this.dgv備料單.CellMouseMove += Dgv備料單_CellMouseMove;
                        this.dgv備料單.CellMouseLeave += Dgv備料單_CellMouseLeave;
                        this.dgv備料單.Scroll += Dgv備料單_Scroll;
                        // 當游標回到料號輸入框時，清除搜尋標示
                        try { this.txt備料單料號.Enter += Txt備料單料號_Enter; } catch { }
                    }
                    catch { }
                }
                catch { }
            }
            catch
            {
                // 忽略初始化期間的小錯誤，避免在 UI 啟動時拋出未處理例外
            }
            // 註冊 DataGridView 的 CellFormatting 以確保顯示樣式在每次繪製時一致
            try { this.dgv備料單.CellFormatting += Dgv備料單_CellFormatting; } catch { }
        }

        /// <summary>
        /// 隱藏 DataGridView 的選取視覺效果並清除任何選取/目前儲存格
        /// - 設定選取顏色為與一般背景相同
        /// - 清除選取並嘗試移除 CurrentCell
        /// </summary>
        /// <param name="dgv"></param>
        private void HideSelectionInGrid(DataGridView dgv)
        {
            try
            {
                // Delegate to the centralized helper to ensure consistent, UI-safe behavior
                // Use the fully-qualified name to avoid needing additional using directives
                Automatic_Storage.Utilities.DgvHelpers.HideSelectionInGrid(this, dgv);
            }
            catch { }
        }

        /// <summary>
        /// 事件處理：使用者按下「匯入檔案」按鈕。
        /// 此方法會以非同步方式開啟檔案選擇對話方塊，並啟動 Excel 讀取與解析流程。
        /// 注意：此處僅負責啟動與協調，實際解析會在背景作業或分離方法中執行以避免阻塞 UI。
        /// </summary>
        /// <param name="sender">事件來源（按鈕）。</param>
        /// <param name="e">事件參數。</param>
        private async void btn備料單匯入檔案_Click(object sender, EventArgs e)
        {
            // 防呆：避免重入
            if (_isImporting)
            {
                try { MessageBox.Show("目前已有匯入作業在進行中，請稍後再試。", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                return;
            }

            // 【關鍵】在最外層先保存遊標狀態，便於在 finally 中還原
            Cursor prevCursor = Cursor.Current;
            bool prevUseWaitCursor = Application.UseWaitCursor;

            _isImporting = true;
            UpdateButtonStates();
            try
            {
                // Restore original import logic (from backup) with UI lock around the whole operation.
                // 使用者按下匯入檔案：不再強制在表單關閉前保留匯入按鈕為不可按，
                // 讓 DataBindingComplete/還原流程能在匯入完成後恢復按鈕狀態
                // (原先此 flag 會導致匯入完成後按鈕無法被還原，造成不可按的 bug)
                _keepImportButtonDisabledUntilClose = false;

                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Filter = "Excel 檔案|*.xls;*.xlsx;*.xlsm|All files|*.*";
                    ofd.Title = "選擇備料單 Excel 檔案";
                    ofd.Multiselect = false;

                    if (ofd.ShowDialog() != DialogResult.OK) return;
                    string path = ofd.FileName;

                    // 【修正】每次使用匯入之前，先清空 UI 與暫存，避免多次匯入資料累加到 dgv 上
                    try
                    {
                        // 解除綁定並清除舊資料
                        if (this.dgv備料單 != null)
                        {
                            try { this.dgv備料單.DataSource = null; } catch { }
                            try { this.dgv備料單.Rows.Clear(); } catch { }
                            try { this.dgv備料單.Refresh(); } catch { }
                        }

                        // 清除內部暫存與狀態
                        try { _currentUiTable = null; } catch { }
                        try { _records?.Clear(); } catch { }
                        try { _lastAppendedRecords?.Clear(); } catch { }
                        try { _lastMatchedRows?.Clear(); } catch { }
                        try { _materialIndex?.Clear(); } catch { }
                        try { _preservedRedKeys?.Clear(); } catch { }
                        try { _isDirty = false; } catch { }
                    }
                    catch { }

                    // 先鎖定 UI 與顯示等待遮罩，確保在任何 COM/IO 操作前就讓使用者看到等待狀態
                    try { try { _cts?.Cancel(); } catch { } _cts = new CancellationTokenSource(); } catch { }
                    try { SaveAndDisableAllButtons(); } catch { }
                    try { SetControlEnabledSafe(this.btn備料單返回, false); } catch { }
                    try { SetControlEnabledSafe(this.btn備料單Unlock, false); } catch { }
                    try { SetControlEnabledSafe(this.btn備料單匯出, false); } catch { }
                    try { SetControlEnabledSafe(this.btn備料單存檔, false); } catch { }
                    try { ShowOperationOverlay(); } catch { }
                    try { _prevUseWaitCursor = Application.UseWaitCursor; Application.UseWaitCursor = true; } catch { }
                    // prevCursor 和 prevUseWaitCursor 已在函式外層定義
                    try { Cursor.Current = Cursors.WaitCursor; } catch { }

                    // 讓步給 UI 執行緒，確保遮罩與游標能立即呈現
                    try { this.Refresh(); Application.DoEvents(); } catch { }
                    try { await Task.Yield(); } catch { }

                    // 初始化記錄容器與讀檔流程
                    _records = new List<Dto.記錄Dto>();
                    _records = new List<Dto.記錄Dto>();
                    try { try { _cts?.Cancel(); } catch { } _cts = new CancellationTokenSource(); } catch { }

                    DataTable? uiTable = null;
                    Exception? loadEx = null;
                    try
                    {
                        // 統一用注入型 IExcelService 讀取 Excel
                        uiTable = GetExcelService().LoadFirstWorksheetToDataTable(path);
                    }
                    catch (OperationCanceledException)
                    {
                        loadEx = new OperationCanceledException("匯入被取消。");
                    }
                    catch (Exception ex)
                    {
                        loadEx = ex;
                    }

                    if (loadEx != null)
                    {
                        SafeBeginInvoke(this, new Action(() =>
                        {
                            MessageBox.Show($"Excel 讀取失敗：{loadEx.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            this.currentExcelPath = null;
                            try { Application.UseWaitCursor = _prevUseWaitCursor; } catch { }
                            try { Cursor.Current = Cursors.Default; } catch { }
                            try { HideOperationOverlay(); } catch { }
                            try { RestoreAllButtons(); } catch { }
                        }));
                        return;
                    }

                    if (uiTable == null || uiTable.Rows.Count == 0)
                    {
                        SafeBeginInvoke(this, new Action(() =>
                        {
                            MessageBox.Show("讀取檔案後未取得資料。", "匯入失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            this.currentExcelPath = null;
                            try { Application.UseWaitCursor = _prevUseWaitCursor; } catch { }
                            try { Cursor.Current = Cursors.Default; } catch { }
                            try { HideOperationOverlay(); } catch { }
                            try { RestoreAllButtons(); } catch { }
                        }));
                        return;
                    }

                    // 公式處理
                    ExcelInteropHelper.RemoveHiddenColumnsFromDataTable(path, uiTable);
                    try { await Task.Run(() => ReplaceFormulasWithValuesFromExcel(path, uiTable, _cts?.Token ?? CancellationToken.None)); }
                    catch (OperationCanceledException)
                    {
                        SafeBeginInvoke(this, new Action(() =>
                        {
                            MessageBox.Show("匯入已被使用者取消。", "取消", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            try { Application.UseWaitCursor = _prevUseWaitCursor; } catch { }
                            try { Cursor.Current = Cursors.Default; } catch { }
                            try { HideOperationOverlay(); } catch { }
                            try { RestoreAllButtons(); } catch { }
                        }));
                        return;
                    }
                    catch (Exception ex)
                    {
                        SafeBeginInvoke(this, new Action(() =>
                        {
                            MessageBox.Show($"公式處理失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }

                    // 修正：取得 mapping 並處理錯誤
                    Dictionary<string, int> mapping;
                    string errMsg;
                    if (!ValidateAndMapColumns(uiTable, out mapping, out errMsg))
                    {
                        SafeBeginInvoke(this, new Action(() =>
                        {
                            MessageBox.Show($"Excel 欄位對應失敗：{errMsg}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                        return;
                    }
                    this._columnMapping = mapping;

                    // DataGridView 綁定 (簡化：採用原先流程)
                    try
                    {
                        try { if (this.dgv備料單 != null) { BeginUpdate(this.dgv備料單); SetDoubleBuffered(this.dgv備料單, true); } } catch { }
                        try
                        {
                            if (this.dgv備料單 is not null)
                            {
                                this.dgv備料單.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                            }
                        }
                        catch { }
                        ResetGridBeforeBind();
                        if (this.dgv備料單 != null && uiTable != null) RebuildGridColumnsFromDataTable(this.dgv備料單, uiTable);
                        try { if (this.dgv備料單 != null) this.dgv備料單.DataSource = uiTable; } catch { }
                        _currentUiTable = uiTable;
                        this.currentExcelPath = path;
                        // 綁定完成後將程式性修改標記為已接受（AcceptChanges），並清除暫存回寫記錄，
                        // 避免 ReplaceFormulasWithValuesFromExcel / NPOI 寫入造成 DataTable.GetChanges() 回傳變更而誤觸發未存檔提醒
                        try { _suspendDirtyMarking = true; } catch { }
                        try
                        {
                            try { uiTable?.AcceptChanges(); } catch { }
                            try { _isDirty = false; } catch { }
                            try { _records?.Clear(); } catch { }
                            try { _lastAppendedRecords?.Clear(); } catch { }
                        }
                        finally
                        {
                            try { _suspendDirtyMarking = false; } catch { }
                        }
                        int shippedCol = FindColumnIndexByNames(new[] { "實發數量", "發料數量" });
                        if (this.dgv備料單 is not null && shippedCol >= 0 && this.dgv備料單.Columns != null && shippedCol < this.dgv備料單.Columns.Count)
                        {
                            this.dgv備料單.Columns[shippedCol].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        }
                        var postToken = _cts?.Token ?? CancellationToken.None;
                        _ = UpdateExcelAfterImportAsync(this.currentExcelPath, postToken);
                        try { if (this.dgv備料單 != null) EndUpdate(this.dgv備料單); } catch { }
                        try { if (this.dgv備料單 != null) HideSelectionInGrid(this.dgv備料單); } catch { }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"綁定資料時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        this.currentExcelPath = null;
                    }

                    // UI 還原
                    try { HideOperationOverlay(); } catch { }
                    try { RestoreAllButtons(); } catch { }
                }
            }
            finally
            {
                // 【關鍵】無論成功或失敗，都必須還原遊標和等待狀態
                try { Application.UseWaitCursor = _prevUseWaitCursor; } catch { }
                try { Cursor.Current = prevCursor; this.Cursor = Cursors.Default; } catch { }
                try { HideOperationOverlay(); } catch { }

                _isImporting = false;
                UpdateButtonStates();

                // 【關鍵】匯入完成後，自動將畫面切換回「鎖定」狀態
                // 這樣按鈕會顯示「解鎖」，並且匯出/存檔按鈕將再次可按（有資料時）
                try { SetEditingLocked(true); } catch { }

                // 在取消或完成匯入後，確保按鈕狀態符合實際資料情況
                try { UpdateMainButtonsEnabled(); } catch { }
            }
        }

        /// <summary>
        /// 事件處理：使用者按下「匯出」按鈕。
        /// 會以非同步方式匯出目前 DataGridView 的內容為外部檔案（例如 CSV/Excel），並在完成後回報狀態。
        /// </summary>
        /// <param name="sender">事件來源（按鈕）。</param>
        /// <param name="e">事件參數。</param>
        private async void btn備料單匯出_Click(object sender, EventArgs e)
        {
            // 防呆：避免重入
            if (_isExporting)
            {
                try { MessageBox.Show("目前已有匯出作業在進行中，請稍後再試。", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                return;
            }
            if (!CheckExcelAvailable()) return;

            _isExporting = true;
            UpdateButtonStates();
            // 顯示執行中提示與等待游標，並暫時停用相關操作按鈕
            var prevCursor = Cursor.Current;
            _prevUseWaitCursor = Application.UseWaitCursor;
            SaveAndDisableAllButtons();
            ShowOperationOverlay("匯出中，請稍候...");
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;
            bool exportSucceeded = false;
            try
            {
                // 檢查是否有未存檔變更：優先判斷 _isDirty 或暫存 _records
                bool hasUnsavedChanges = false;
                try
                {
                    if (_isDirty) hasUnsavedChanges = true;
                    if (_records != null && _records.Count > 0) hasUnsavedChanges = true;

                    // 如果目前沒有 _isDirty / _records，但 DataTable 報告有變更，進一步檢查是否為使用者可編輯的欄位
                    var dt = this.dgv備料單?.DataSource as DataTable;
                    if (!hasUnsavedChanges && dt != null && DataTableHasRealChanges(dt))
                    {
                        bool userEditableChangeFound = false;

                        // Alternate approach: inspect the changes.Columns names for any column that looks like "實發/發料".
                        // This avoids relying on grid column index mapping which can be fragile after DataTable schema adjustments.
                        var shippedSynonyms = new[] { "實發數量", "發料數量" };
                        // Reuse centralized sanitizer for various local normalizations
                        Func<string, string> San = SanitizeHeaderForMatch;
                        var shippedNorms = new HashSet<string>(shippedSynonyms.Select(San));

                        try
                        {
                            var changes = dt.GetChanges();
                            if (changes != null)
                            {
                                // Log changed column names for diagnosis
                                try
                                {
                                    var changedCols = changes.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray();

                                }
                                catch { }

                                foreach (DataColumn ch in changes.Columns)
                                {
                                    try
                                    {
                                        var colName = ch.ColumnName ?? string.Empty;
                                        var colNorm = San(colName);
                                        if (string.IsNullOrEmpty(colNorm)) continue;
                                        if (shippedNorms.Contains(colNorm) || shippedNorms.Any(s => colNorm.Contains(s) || s.Contains(colNorm)))
                                        {
                                            // Found a changed column that looks like shipped -> inspect rows to confirm value change
                                            foreach (DataRow r in changes.Rows)
                                            {
                                                try
                                                {
                                                    object? orig = null, cur = null;
                                                    try { orig = r[colName, DataRowVersion.Original]; } catch { orig = null; }
                                                    try { cur = r[colName, DataRowVersion.Current]; } catch { cur = null; }
                                                    bool origEmpty = orig == null || orig == DBNull.Value || (orig is string os && string.IsNullOrWhiteSpace(os));
                                                    bool curEmpty = cur == null || cur == DBNull.Value || (cur is string cs && string.IsNullOrWhiteSpace(cs));
                                                    if (origEmpty && curEmpty) continue;
                                                    if (orig is string os2 && cur is string cs2)
                                                    {
                                                        if (!string.Equals(os2.Trim(), cs2.Trim(), StringComparison.Ordinal)) { userEditableChangeFound = true; break; }
                                                    }
                                                    else
                                                    {
                                                        if (!object.Equals(orig, cur)) { userEditableChangeFound = true; break; }
                                                    }
                                                }
                                                catch { }
                                            }
                                            if (userEditableChangeFound) break;
                                        }
                                    }
                                    catch { }
                                    if (userEditableChangeFound) break;
                                }
                                // Also write a sample of the first few changed cells into the debug log to aid diagnosis
                                try
                                {
                                    int rowIdx = 0;
                                    foreach (DataRow r in changes.Rows)
                                    {
                                        try
                                        {
                                            var samples = new List<string>();
                                            foreach (DataColumn c in changes.Columns)
                                            {
                                                try
                                                {
                                                    object? o1 = null, o2 = null;
                                                    try { o1 = r[c.ColumnName, DataRowVersion.Original]; } catch { o1 = null; }
                                                    try { o2 = r[c.ColumnName, DataRowVersion.Current]; } catch { o2 = null; }
                                                    samples.Add($"{c.ColumnName}:[{o1}]->[{o2}]");
                                                }
                                                catch { }
                                            }
                                            // Debug log removed
                                        }
                                        catch { }
                                        rowIdx++; if (rowIdx >= 5) break;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }

                        // 如果找不到使用者可編輯欄位的實際變更，則視為程式性變更（例如 ReadOnly/Focus/格式套用），不視為未存檔
                        if (userEditableChangeFound) hasUnsavedChanges = true;
                    }
                }
                catch { }

                // 診斷資訊已移除（留存 log 與 AppendDebugLog），避免在正常匯出流程彈出測試用視窗

                if (hasUnsavedChanges)
                {
                    var dr = MessageBox.Show("畫面有未存檔的變更，是否先存檔？\n(選是將先執行存檔，選否則直接繼續匯出)", "未存檔提醒", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (dr == DialogResult.Cancel)
                    {
                        // 使用者取消時，確保 UI 狀態被完全還原，避免卡在等待游標
                        EnsureUiFullyRestored();
                        return;
                    }
                    else if (dr == DialogResult.Yes)
                    {
                        // 先存檔，並取得詳細錯誤
                        SaveResultDto? saveResult = null;
                        try
                        {
                            saveResult = await SaveAsyncWithResult();
                        }
                        catch (Exception ex)
                        {
                            saveResult = new SaveResultDto { Success = false, ErrorMessage = ex.Message };
                        }
                        if (saveResult == null || !saveResult.Success)
                        {
                            // 讀取本地測試日誌已移除；僅使用 saveResult.ErrorMessage 作為錯誤來源
                            string? lastError = null;
                            var msg = "存檔失敗，已取消匯出。";
                            if (!string.IsNullOrEmpty(saveResult?.ErrorMessage))
                                msg += "\n\n錯誤主因：" + (saveResult?.ErrorMessage ?? string.Empty);
                            if (!string.IsNullOrEmpty(lastError))
                                msg += "\n\n詳細錯誤：" + lastError;
                            MessageBox.Show(msg, "取消匯出", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            // 在取消匯出前，確保 UI 狀態被完全還原，避免卡在等待游標或按鈕停用
                            EnsureUiFullyRestored();
                            return;
                        }
                    }
                    // 若為 No，則繼續匯出（不儲存）
                }

                // 顯示存檔對話框讓使用者選擇輸出位置
                bool showDialogCancelled = false;
                using (var sfd = new SaveFileDialog())
                {
                    sfd.Title = "另存為...";
                    sfd.Filter = "Excel 檔案 (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|所有檔案 (*.*)|*.*";
                    sfd.FileName = Path.GetFileName(this.currentExcelPath) ?? "export.xlsx";
                    // 關閉 SaveFileDialog 的內建覆蓋提示，改由程式自行顯示自訂覆蓋確認，避免兩次提示
                    try { sfd.OverwritePrompt = false; } catch { }
                    try { EnsureCursorRestored(); } catch { }
                    if (sfd.ShowDialog(this) != DialogResult.OK)
                    {
                        showDialogCancelled = true;
                    }
                    else
                    {
                        var dest = sfd.FileName;
                        try
                        {
                            // 若目的檔案已存在，顯示更完整的檔案資訊並詢問是否覆蓋
                            if (File.Exists(dest))
                            {
                                try
                                {
                                    var fi = new FileInfo(dest);
                                    string sizeText = fi.Length >= 1024 ? (fi.Length >= 1024 * 1024 ? (fi.Length / (1024.0 * 1024.0)).ToString("F2") + " MB" : (fi.Length / 1024.0).ToString("F1") + " KB") : fi.Length + " bytes";
                                    string info = $"檔案已存在：{dest}\r\n大小：{sizeText}\r\n最後修改：{fi.LastWriteTime:yyyy/MM/dd HH:mm}\r\n是否要覆蓋？";
                                    var over = MessageBox.Show(info, "確認覆蓋", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    if (over != DialogResult.Yes)
                                    {
                                        // 使用者選「否」時，也要還原游標後才 return
                                        try { Application.UseWaitCursor = false; Cursor.Current = Cursors.Default; this.Cursor = Cursors.Default; } catch { }
                                        return;
                                    }
                                }
                                catch
                                {
                                    var over = MessageBox.Show($"檔案已存在：{dest}\n是否要覆蓋？", "確認覆蓋", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                                    if (over != DialogResult.Yes)
                                    {
                                        // 使用者選「否」時，也要還原游標後才 return
                                        try { Application.UseWaitCursor = false; Cursor.Current = Cursors.Default; this.Cursor = Cursors.Default; } catch { }
                                        return;
                                    }
                                }
                            }

                            File.Copy(this.currentExcelPath, dest, true);
                            MessageBox.Show("已匯出檔案: " + dest, "匯出完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            // 標記匯出成功；只有成功時才會在 finally 中清除畫面/內部資料
                            exportSucceeded = true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("匯出失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                if (showDialogCancelled)
                {
                    // SaveFileDialog 被取消時，標記為取消狀態
                    // 不再在這裡做任何 return，讓流程走完 finally 才退出
                    // 這樣可以確保 finally 的 Application.UseWaitCursor = false 能被執行
                }
                // 注意：即使 showDialogCancelled，也要讓程式走到 finally
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯出失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 【關鍵】export 完成後，ALWAYS 強制設定 UseWaitCursor = false，無論原狀態為何
                // 這樣可以確保即使 prevCursor 被保存為非 Default 值，遊標也會被完全還原
                try { Application.UseWaitCursor = false; } catch { }
                try { Cursor.Current = Cursors.Default; } catch { }
                try { this.Cursor = Cursors.Default; } catch { }

                // 對所有 open forms 遍歷，明確設定 Cursor = Default
                try
                {
                    foreach (Form f in Application.OpenForms)
                    {
                        try
                        {
                            if (f != null)
                            {
                                f.Cursor = Cursors.Default;
                                // 對 form 的所有 controls 也直接設定
                                foreach (Control c in f.Controls)
                                {
                                    try { if (c != null) c.Cursor = Cursors.Default; } catch { }
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                try { HideOperationOverlay(); } catch { }
                try { RestoreAllButtons(); } catch { }

                // 結束 DataGridView 編輯、清除焦點、鎖定編輯
                try { if (this.dgv備料單 != null) { if (this.dgv備料單.IsCurrentCellInEditMode) { this.dgv備料單.EndEdit(); } try { this.dgv備料單.CancelEdit(); } catch { } try { this.dgv備料單.CurrentCell = null; } catch { } try { this.dgv備料單.ClearSelection(); } catch { } } } catch { }
                try { SetEditingLocked(true); } catch { }

                // 清除快速索引與暫存記錄
                try { ClearMaterialIndex(); } catch { }

                // 只有在實際匯出成功時，才會清除或重置 form 內部狀態與畫面資料
                if (exportSucceeded)
                {
                    try { ResetFormState(); } catch { }
                }

                // 解除匯出旗標並更新按鈕狀態
                try { _isExporting = false; } catch { }
                UpdateButtonStates();
                try { UpdateMainButtonsEnabled(); } catch { }

                // 如果匯出成功，則在匯出完成後將「解鎖」按鈕設為不可按，避免使用者在匯出後立即修改或解鎖
                try
                {
                    if (exportSucceeded)
                    {
                        try { _keepUnlockButtonDisabledUntilClose = true; } catch { }
                        try { SetControlEnabledSafe(this.btn備料單Unlock, false); } catch { }
                    }
                }
                catch { }

                // 最後強制把焦點移到 form，並呼叫 ReleaseCapture + 遊標重繪
                try { this.Focus(); } catch { }
                try { this.Activate(); } catch { }
                try { NativeHelpers.ReleaseCapture(); } catch { }
                try { Cursor.Hide(); Cursor.Show(); } catch { }
                try { Application.DoEvents(); } catch { }

                // 【診斷】記錄最終 UI 狀態快照以診斷遊標卡住問題
                try
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine($"[EXPORT_FINALLY_SNAPSHOT]");
                    sb.AppendLine($"  Application.UseWaitCursor={Application.UseWaitCursor}");
                    sb.AppendLine($"  Cursor.Current={Cursor.Current}");
                    sb.AppendLine($"  this.Cursor={this.Cursor}");
                    sb.AppendLine($"  prevCursor={prevCursor}");
                    sb.AppendLine($"  _prevUseWaitCursor={_prevUseWaitCursor}");
                    sb.AppendLine($"  exportSucceeded={exportSucceeded}");
                    sb.AppendLine($"  _isExporting={_isExporting}");

                    // 掃描所有 form 和頂層 control 的 Cursor 狀態
                    try
                    {
                        int fIdx = 0;
                        foreach (Form f in Application.OpenForms)
                        {
                            if (f != null)
                            {
                                fIdx++;
                                sb.AppendLine($"  [Form {fIdx}] {f.GetType().Name} @ {f.Handle}: Cursor={f.Cursor}");

                                // 記錄 form 內任何 Cursor != Default 的 control
                                try
                                {
                                    int nonDefaultCount = 0;
                                    foreach (Control c in f.Controls)
                                    {
                                        if (c != null && c.Cursor != Cursors.Default)
                                        {
                                            nonDefaultCount++;
                                            sb.AppendLine($"    ⚠ Control {c.GetType().Name}: Cursor={c.Cursor}");
                                        }
                                    }
                                    if (nonDefaultCount == 0)
                                    {
                                        sb.AppendLine($"    ✓ All top-level controls have Cursor.Default");
                                    }
                                }
                                catch (Exception ex) { sb.AppendLine($"    [Error scanning controls: {ex.Message}]"); }
                            }
                        }
                    }
                    catch (Exception ex) { sb.AppendLine($"  [Error scanning forms: {ex.Message}]"); }


                }
                catch { }
            }
        }

        /// <summary>
        /// 事件處理：使用者按下「返回」按鈕。
        /// 此方法會關閉表單或回到上一層畫面，並在必要時清理臨時狀態。
        /// </summary>
        /// <param name="sender">事件來源（按鈕）。</param>
        /// <param name="e">事件參數。</param>
        private void btn備料單返回_Click(object sender, EventArgs e)
        {
            // 需求：返回只隱藏，不要重置目前表單資料；再次由主畫面開啟時，恢復同一個實例
            try { RestoreAllButtons(); } catch { }
            try { HideOperationOverlay(); } catch { }

            // 隱藏目前視窗並回到 owner（不需驗證 Excel，避免在返回時出現異常）
            try { this.Hide(); } catch { }
            if (this.Owner is Form owner)
            {
                try { owner.Show(); owner.BringToFront(); owner.Activate(); } catch { }
            }
        }

        /// <summary>
        /// [事件] 備料單解鎖/鎖定按鈕點擊事件。
        /// - 若目前為鎖定狀態，點擊後顯示密碼輸入對話框，驗證成功則解鎖（可編輯）；失敗或取消則維持鎖定。
        /// - 若目前為解鎖狀態，點擊後直接鎖定（不可編輯）。
        /// - 過程中會自動處理游標、UI 狀態與按鈕啟用狀態，避免重複操作或異常狀態殘留。
        /// </summary>
        /// <param name="sender">事件來源（按鈕控制項）。</param>
        /// <param name="e">事件參數。</param>
        private void btn備料單Unlock_Click(object sender, EventArgs e)
        {
            // 防呆：避免在其他操作進行時解鎖
            if (_isImporting || _isExporting || _isSaving)
            {
                try { MessageBox.Show("目前有作業在進行中，請稍後再試。", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                return;
            }

            // 暫時禁用解鎖按鈕，避免使用者重複點擊
            try { SetControlEnabledSafe(this.btn備料單Unlock, false); } catch { }
            // 不在顯示密碼輸入時使用等待游標或遮罩，避免阻塞使用者輸入
            // 若有較長時間的背景工作再顯示 overlay / 等待游標

            SafeBeginInvoke(this, new Action(() =>
            {
                try
                {
                    if (_isEditingLocked)
                    {
                        // 在顯示密碼輸入對話框前，確保游標與 UI 狀態已恢復（避免對話框出現時仍為等待游標）
                        try { EnsureCursorRestored(); } catch { }

                        // 需要密碼才能解鎖畫面編輯：使用可重試的對話框，錯誤時不自動關閉視窗
                        string expected = GetExcelPassword();
                        bool ok = PromptForPasswordWithRetry(expected);
                        if (ok)
                        {
                            // 解鎖成功：設定為解鎖狀態（false = 解鎖）
                            // SetEditingLocked() 內部會呼叫 UpdateMainButtonsEnabled() 更新所有按鈕狀態
                            SetEditingLocked(false);
                        }
                        else
                        {
                            // 使用者取消或驗證失敗，不做任何變更，按鈕狀態保持原樣（鎖定狀態不變）
                        }
                    }
                    else
                    {
                        // 已解鎖，點擊即鎖定（不需密碼）
                        // SetEditingLocked() 內部會呼叫 UpdateMainButtonsEnabled() 更新所有按鈕狀態
                        SetEditingLocked(true);
                    }
                }
                catch (Exception ex)
                {
                    // Fire-and-forget log to avoid changing method signature; swallow any errors from logging
                    try { _ = Task.Run(() => Utilities.Logger.LogErrorAsync("Unlock UI failed: " + ex.Message)); } catch { }
                }
                finally
                {
                    // 強制還原游標與 UI 狀態（使用強制還原而非回復 prev，避免 prev 在其他流程被改變導致殘留）
                    try { Application.UseWaitCursor = false; } catch { }
                    try { Cursor.Current = Cursors.Default; } catch { }
                    try { this.Cursor = Cursors.Default; } catch { }
                    try { EnsureCursorRestored(); } catch { }
                    try { HideOperationOverlay(); } catch { }
                    try { if (!_keepUnlockButtonDisabledUntilClose) SetControlEnabledSafe(this.btn備料單Unlock, true); } catch { }
                }
            }));
        }

        #region DataGridView helpers

        /// <summary>
        /// 依據 DataTable 欄位重建 DataGridView 欄位，確保 UI 欄位與資料來源一致。
        /// 會清除原有欄位並建立對應的 DataPropertyName / HeaderText。
        /// </summary>
        /// <param name="dgv">目標 DataGridView</param>
        /// <param name="dt">來源 DataTable</param>
        private void RebuildGridColumnsFromDataTable(DataGridView dgv, DataTable dt)
        {
            if (dgv == null || dt == null) return;
            dgv.Columns.Clear();
            foreach (DataColumn col in dt.Columns)
            {
                var gridCol = new DataGridViewTextBoxColumn
                {
                    Name = col.ColumnName,
                    HeaderText = col.ColumnName,
                    DataPropertyName = col.ColumnName,
                    ReadOnly = true,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                };
                // ...existing code...
            }
        }
        /// <summary>
        /// 清除 DataGridView 中所有儲存格的高亮顏色（例如黃色標示），但保留短缺（紅色）標示。
        /// 此方法會將非紅色的儲存格背景色還原為預設顏色，並還原前景色。
        /// 常用於料號搜尋後移除黃色標示，或在資料重設時恢復原始樣式。
        /// </summary>
        private void ClearRowHighlights()
        {
            var dgv = this.dgv備料單;
            if (dgv == null || dgv.Rows == null) return;
            var defaultBack = dgv.DefaultCellStyle?.BackColor ?? SystemColors.Window;
            var preservedColor = Color.Red; // keep shortage red highlights
            foreach (DataGridViewRow row in dgv.Rows)
            {
                if (row == null || row.IsNewRow) continue;
                foreach (DataGridViewCell cell in row.Cells)
                {
                    try
                    {
                        // 如果儲存格目前已經是紅色（短缺標示），保留之，不要還原
                        var cur = cell.Style?.BackColor ?? Color.Empty;
                        if (cur == preservedColor) continue;
                        if (cell is not null && cell.Style is not null)
                        {
                            cell.Style.BackColor = defaultBack;
                        }
                        // 還原前景色為預設
                        try { cell.Style.ForeColor = dgv.DefaultCellStyle?.ForeColor ?? SystemColors.ControlText; } catch { }
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// 確保 UI 狀態與游標被完整還原。
        /// 在所有會導致流程提前 return 的路徑呼叫，避免游標或按鈕停留在等待狀態。
        /// </summary>
        private void EnsureUiFullyRestored()
        {
            try { Application.UseWaitCursor = false; } catch { }
            try { Cursor.Current = Cursors.Default; } catch { }
            try { this.Cursor = Cursors.Default; } catch { }
            try
            {
                foreach (Form f in Application.OpenForms)
                {
                    try { if (f != null) f.Cursor = Cursors.Default; } catch { }
                }
            }
            catch { }
            try
            {
                // 取消任何控制項的 mouse capture（避免有控制項鎖住滑鼠）
                foreach (Form f in Application.OpenForms)
                {
                    try
                    {
                        if (f == null) continue;
                        if (f.Capture) f.Capture = false;
                        foreach (Control c in f.Controls)
                        {
                            try { if (c != null && c.Capture) c.Capture = false; } catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            try { Application.DoEvents(); } catch { }
            try { NativeHelpers.ReleaseCapture(); } catch { }
            try { Cursor.Hide(); Cursor.Show(); } catch { }
            try { _isExporting = false; } catch { }
            try { UpdateButtonStates(); } catch { }
            try { RestoreAllButtons(); } catch { }
            try { HideOperationOverlay(); } catch { }
            try
            {
                var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ui_restore_debug.log");
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("--- EnsureUiFullyRestored DIAG " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " ---");
                try
                {
                    foreach (Form f in Application.OpenForms)
                    {
                        try
                        {
                            sb.AppendLine($"Form: Name={f?.Name} Text={f?.Text} Visible={f?.Visible} WindowState={f?.WindowState} Handle={f?.Handle} Cursor={f?.Cursor} Capture={f?.Capture} ActiveControl={(f?.ActiveControl != null ? f.ActiveControl.Name : "<none>")}");
                            if (f is not null && f.Controls is not null)
                            {
                                foreach (Control c in f.Controls)
                                {
                                    if (c is null) continue;
                                    try
                                    {
                                        string cType = c.GetType().Name;
                                        string cName = c.Name;
                                        string cInfo = $"  Control: Type={cType} Name={cName} Visible={c.Visible} Enabled={c.Enabled} Cursor={c.Cursor} Capture={c.Capture} Focused={c.Focused}";
                                        sb.AppendLine(cInfo);

                                        // 若控制項為 DataGridView，嘗試強制 EndEdit/CancelEdit 並記錄狀態
                                        if (c is DataGridView dgv)
                                        {
                                            try
                                            {
                                                if (dgv.Rows is not null)
                                                    sb.AppendLine($"    DataGridView: Rows={dgv.Rows.Count} Columns={(dgv.Columns != null ? dgv.Columns.Count : -1)} CurrentCell={(dgv.CurrentCell != null ? dgv.CurrentCell.RowIndex + "," + dgv.CurrentCell.ColumnIndex : "<none>")} IsCurrentCellInEditMode={dgv.IsCurrentCellInEditMode} EditMode={dgv.EditMode}");
                                                // 強制嘗試結束編輯與取消編輯
                                                try { if (dgv.IsCurrentCellInEditMode) dgv.EndEdit(); } catch { }
                                                try { if (dgv.IsCurrentCellInEditMode) dgv.CancelEdit(); } catch { }
                                                try { dgv.ClearSelection(); } catch { }
                                                try { dgv.CurrentCell = null; } catch { }
                                                sb.AppendLine($"    AfterForce: IsCurrentCellInEditMode={dgv.IsCurrentCellInEditMode} CurrentCell={(dgv.CurrentCell != null ? dgv.CurrentCell.RowIndex + "," + dgv.CurrentCell.ColumnIndex : "<none>")}");
                                            }
                                            catch (Exception exDgv)
                                            {
                                                sb.AppendLine($"    DataGridView exception: {exDgv}");
                                            }
                                        }
                                    }
                                    catch { }
                                    try { c.Cursor = Cursors.Default; } catch { }
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch { }
                sb.AppendLine("--- End DIAG ---\r\n");
                // Debug write removed
                // 確保將焦點移回安全控制項
                try { if (Application.OpenForms.Count > 0) { var f0 = Application.OpenForms[0]; try { f0.Activate(); if (f0.ActiveControl != null) f0.ActiveControl.Focus(); else f0.Focus(); } catch { } } } catch { }
            }
            catch { }
            try { ResetAllCursors(); } catch { }
        }

        /// <summary>
        /// 強制還原所有游標狀態（等待游標 & 各 Form/Control 游標）
        /// 可在顯示 MessageBox / SaveFileDialog 之前或之後呼叫，確保 UI 不會殘留等待游標
        /// </summary>
        private void EnsureCursorRestored()
        {
            try { Application.UseWaitCursor = false; } catch { }
            try { Cursor.Current = Cursors.Default; } catch { }
            try { this.Cursor = Cursors.Default; } catch { }
            try
            {
                foreach (Form f in Application.OpenForms)
                {
                    try { if (f != null) f.Cursor = Cursors.Default; } catch { }
                    try
                    {
                        if (f is not null && f.Controls is not null)
                        {
                            foreach (Control c in f.Controls)
                            {
                                if (c is not null)
                                {
                                    // ...existing code...
                                }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            try { Cursor.Hide(); Cursor.Show(); } catch { }
        }

        /// <summary>
        /// 提供與 Win32 使用者介面相關的本機 interop 幫助方法，用於還原 UI 狀態（例如釋放滑鼠捕捉）。
        /// 僅供 <see cref="Form備料單匯入"/> 類別內部使用。
        /// </summary>
        private static class NativeHelpers
        {
            /// <summary>
            /// 釋放目前執行緒對滑鼠的捕捉（mouse capture），使滑鼠輸入恢復到預設行為。
            /// </summary>
            /// <remarks>
            /// 此方法封裝 Win32 API <c>ReleaseCapture</c> 的呼叫。
            /// 在某些拖曳或控制項呼叫 <c>SetCapture</c> 後，系統會將滑鼠事件限制於單一視窗，
            /// 呼叫本方法可確保系統不再將事件鎖定於該視窗。
            /// </remarks>
            /// <returns>
            /// <see langword="true" /> if the function succeeds; otherwise, <see langword="false" />.
            /// </returns>
            [System.Runtime.InteropServices.DllImport("user32.dll")]
            public static extern bool ReleaseCapture();
        }

        /// <summary>
        /// 在 DataGridView 綁定完成後的事件處理器。
        /// 負責完成 UI 的後處理，包括欄位隱藏、欄寬調整、樣式套用與按鈕還原等。
        /// </summary>
        /// <param name="sender">事件來源（DataGridView）</param>
        /// <param name="e">事件參數</param>
        private void Dgv備料單_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            try
            {
                // 綁定完成前先嘗試隱藏被 Excel 標記為隱藏的欄位，避免後續自動寬度把隱藏欄撐開
                try { HideColumnsByHeaders(this.dgv備料單); } catch { }
                // 自動調整欄寬：先以內容計算每欄所需寬度，再改成 Fill 模式填滿整個 DataGridView，並只顯示垂直滾動條
                try { AutoSizeColumnsFillNoHorizontalScroll(this.dgv備料單); } catch { }
                // 將短缺上色延後（避免阻塞綁定流程）；在訊息迴圈空檔以 BeginInvoke 執行
                try { SafeBeginInvoke(this, new Action(() => { try { MarkShortagesInGrid(); } catch { } })); } catch { }
                try { AdjustGridAppearance(); } catch { }

                // 建立快速索引改延後（避免阻塞綁定流程）
                try { SafeBeginInvoke(this, new Action(() => { try { BuildMaterialIndex(); } catch { } })); } catch { }
                // 表單尺寸調整與欄位對齊也延後，進一步縮短 DataBindingComplete 的同步時間
                try { SafeBeginInvoke(this, new Action(() => { try { AdjustFormSizeToDataGrid(); } catch { } })); } catch { }
                try { SafeBeginInvoke(this, new Action(() => { try { ApplyAlignmentsToGrid(); } catch { } })); } catch { }

                // 將 DataGridView 設為僅呈現模式：禁止排序、禁止重新排序、並隱藏選取效果
                try
                {
                    var dgv = this.dgv備料單;
                    if (dgv != null)
                    {
                        try { dgv.AllowUserToAddRows = false; } catch { }
                        try { dgv.AllowUserToDeleteRows = false; } catch { }
                        try { dgv.AllowUserToOrderColumns = false; } catch { }
                        try { dgv.AllowUserToResizeRows = false; } catch { }
                        try { dgv.AllowUserToResizeColumns = true; } catch { }
                        try { dgv.ReadOnly = true; } catch { }
                        try { dgv.SelectionMode = DataGridViewSelectionMode.CellSelect; } catch { }

                        // 隱藏選取的視覺效果：將選取色設為與背景相同，並在綁定後清除選取
                        try
                        {
                            var bg = dgv.DefaultCellStyle?.BackColor ?? SystemColors.Window;
                            var fg = dgv.DefaultCellStyle?.ForeColor ?? SystemColors.ControlText;
                            if (dgv is not null && dgv.DefaultCellStyle is not null)
                            {
                                dgv.DefaultCellStyle.SelectionBackColor = bg;
                            }
                            if (dgv is not null && dgv.DefaultCellStyle is not null)
                            {
                                dgv.DefaultCellStyle.SelectionForeColor = fg;
                            }
                        }
                        catch { }

                        // 逐欄設為不可排序與唯讀（以展示為主）
                        try
                        {
                            foreach (DataGridViewColumn col in dgv.Columns)
                            {
                                try { col.SortMode = DataGridViewColumnSortMode.NotSortable; } catch { }
                                try { col.ReadOnly = true; } catch { }
                                try { col.Resizable = DataGridViewTriState.False; } catch { }
                            }
                        }
                        catch { }

                        // 清除任何選取（統一呼叫 HideSelectionInGrid 以確保一致行為）
                        try { HideSelectionInGrid(dgv); } catch { }
                    }
                }
                catch { }

                // 匯入完成後清空輸入欄位，避免殘留空白或前置空格，並把焦點放回料號欄
                try
                {
                    try { if (this.txt備料單料號 != null) this.txt備料單料號.Text = string.Empty; } catch { }
                    try { if (this.txt備料單數量 != null) this.txt備料單數量.Text = string.Empty; } catch { }
                    try { this.txt備料單料號?.Focus(); } catch { }
                }
                catch { }

                // 背景檢查檔案是否受保護，並更新 Unlock 按鈕狀態
                try
                {
                    if (!string.IsNullOrWhiteSpace(this.currentExcelPath) && File.Exists(this.currentExcelPath))
                    {
                        Task.Run(() =>
                        {
                            try
                            {
                                // 只檢查保護狀態以便決定按鈕是否可用，但不要在匯入流程自動修改按鈕文字。
                                // 按鈕文字應僅在使用者互動（按下 Unlock）或明確流程需要時才改變。
                                bool nowProtected = _excel_service_safe(() => GetExcelService().IsWorksheetProtected(this.currentExcelPath ?? string.Empty));
                                SafeBeginInvoke(this, new Action(() =>
                                {
                                    try { if (!_keepUnlockButtonDisabledUntilClose) SetControlEnabledSafe(this.btn備料單Unlock, true); } catch { }
                                }));
                            }
                            catch { }
                        });
                    }
                }
                catch { }
            }
            finally
            {
                try { HideSelectionInGrid(this.dgv備料單); } catch { }
                try { if (this.dgv備料單.Columns.Count > 0) this.dgv備料單.FirstDisplayedScrollingColumnIndex = 0; } catch { }

                // 在 DataBindingComplete 完成後，確保還原按鈕狀態與游標
                try
                {
                    // 還原等待游標
                    try { Application.UseWaitCursor = _prevUseWaitCursor; } catch { }
                    try { Cursor.Current = Cursors.Default; } catch { }
                    // 匯入並完成 UI 綁定後，允許匯入按鈕恢復
                    try { _keepImportButtonDisabledUntilClose = false; } catch { }
                    // 在 DataBinding 完成後，也清除匯出時設置的 keep flag，
                    // 以免先前一次成功匯出後將 Unlock 永久鎖住，造成下次匯入後按鈕仍不可按的問題。
                    try { _keepUnlockButtonDisabledUntilClose = false; } catch { }
                    // 解鎖依需求被鎖的按鈕（返回按鈕始終可用）
                    try { SetControlEnabledSafe(this.btn備料單返回, true); } catch { }
                    // 其他按鈕狀態由 UpdateMainButtonsEnabled 決定，此處不做無條件啟用
                    // 還原其他按鈕狀態
                    try { RestoreAllButtons(); } catch { }
                    try { HideOperationOverlay(); } catch { }
                    try { RestorePreservedRedHighlights(); } catch { }
                }
                catch { }
                // 確保 DataBinding 完成後，套用目前的畫面編輯鎖定狀態
                try { SetEditingLocked(_isEditingLocked); } catch { }
                // 根據 DataGridView 是否有資料，更新主按鈕狀態（此必須在最後呼叫以確保邏輯正確）
                try
                {
                    // If the UI now has data after binding/import, allow Unlock again even if it was disabled by an earlier export
                    try
                    {
                        bool hasRowsForUi = false;
                        var dgvLocal = this.dgv備料單;
                        try
                        {
                            if (dgvLocal != null && dgvLocal.DataSource is DataTable dtLocal && dtLocal.Rows != null && dtLocal.Rows.Count > 0)
                                hasRowsForUi = true;
                        }
                        catch { }
                        if (!hasRowsForUi && dgvLocal != null && dgvLocal.Rows != null)
                        {
                            try
                            {
                                foreach (DataGridViewRow row in dgvLocal.Rows)
                                {
                                    if (row != null && !row.IsNewRow)
                                    {
                                        hasRowsForUi = true;
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                        try
                        {
                            if (!hasRowsForUi && _currentUiTable != null && _currentUiTable.Rows != null && _currentUiTable.Rows.Count > 0)
                                hasRowsForUi = true;
                        }
                        catch { }

                        if (hasRowsForUi)
                        {
                            // clear the export-disable flag so Unlock can be enabled for the newly loaded data
                            try { _keepUnlockButtonDisabledUntilClose = false; } catch { }
                        }
                    }
                    catch { }

                    try { UpdateMainButtonsEnabled(); } catch { }
                }
                catch { }
            }
        }

        /// <summary>
        /// 設定畫面編輯鎖定狀態：會嘗試找出實發/發料數量欄位並切換 ReadOnly
        /// 同時更新 Unlock 按鈕顯示文字。
        /// </summary>
        /// <param name="locked">true 表示鎖定（不可編輯），false 表示解鎖（可編輯）</param>
        private void SetEditingLocked(bool locked)
        {
            // Update internal flag first
            _isEditingLocked = locked;

            // Temporarily suspend dirty marking while we programmatically adjust ReadOnly/Focus/etc.
            // This prevents SetEditingLocked from causing false-positive "unsaved changes".
            bool dtHadChangesBefore = false;
            try
            {
                var dtBefore = this.dgv備料單?.DataSource as DataTable;
                dtHadChangesBefore = dtBefore != null && dtBefore.GetChanges() != null;
            }
            catch { dtHadChangesBefore = false; }
            try { _suspendDirtyMarking = true; } catch { }

            try
            {
                var dgv = this.dgv備料單;
                if (dgv != null && dgv.Columns != null)
                {
                    // Find the preferred editable column (shippedCol)
                    int shippedCol = FindColumnIndexByNames(new[] { "實發數量", "發料數量" });

                    if (locked)
                    {
                        // Lock: make whole grid readonly to be safe
                        try { dgv.ReadOnly = true; } catch { }
                        // Also ensure every column is readonly
                        for (int i = 0; i < dgv.Columns.Count; i++)
                        {
                            try { dgv.Columns[i].ReadOnly = true; } catch { }
                        }
                    }
                    else
                    {
                        // Unlock: make grid editable, but enforce per-column readonly except the shippedCol
                        try { dgv.ReadOnly = false; } catch { }

                        for (int i = 0; i < dgv.Columns.Count; i++)
                        {
                            try
                            {
                                if (i == shippedCol) dgv.Columns[i].ReadOnly = false; else dgv.Columns[i].ReadOnly = true;
                            }
                            catch { }
                        }
                    }

                    // If bound to a DataTable, also sync the DataColumn.ReadOnly property for the target column
                    try
                    {
                        if (dgv.DataSource is DataTable dt && shippedCol >= 0 && shippedCol < dgv.Columns.Count)
                        {
                            var col = dgv.Columns[shippedCol];
                            var propName = !string.IsNullOrEmpty(col.DataPropertyName) ? col.DataPropertyName : col.Name;
                            if (!string.IsNullOrEmpty(propName) && dt.Columns.Contains(propName))
                            {
                                try { dt.Columns[propName].ReadOnly = locked; } catch { }
                            }
                        }
                    }
                    catch { }

                    // Re-assert settings after pending events (DataBindingComplete or other handlers may run afterwards)
                    try
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            try
                            {
                                var dgv2 = this.dgv備料單;
                                if (dgv2 != null && dgv2.Columns != null)
                                {
                                    // reapply per-column readonly as final authority
                                    for (int i = 0; i < dgv2.Columns.Count; i++)
                                    {
                                        try
                                        {
                                            if (!locked && i == shippedCol) dgv2.Columns[i].ReadOnly = false; else dgv2.Columns[i].ReadOnly = true;
                                        }
                                        catch { }
                                    }

                                    // sync DataTable again
                                    try
                                    {
                                        if (dgv2.DataSource is DataTable dt2 && shippedCol >= 0 && shippedCol < dgv2.Columns.Count)
                                        {
                                            var col2 = dgv2.Columns[shippedCol];
                                            var propName2 = !string.IsNullOrEmpty(col2.DataPropertyName) ? col2.DataPropertyName : col2.Name;
                                            if (!string.IsNullOrEmpty(propName2) && dt2.Columns.Contains(propName2))
                                            {
                                                try { dt2.Columns[propName2].ReadOnly = locked; } catch { }
                                            }
                                        }
                                    }
                                    catch { }

                                    // If unlocked, try to focus and begin edit again
                                    if (!locked)
                                    {
                                        try
                                        {
                                            int targetCol = shippedCol;
                                            if (targetCol >= 0 && targetCol < dgv2.Columns.Count)
                                            {
                                                foreach (DataGridViewRow row in dgv2.Rows)
                                                {
                                                    if (row == null || row.IsNewRow) continue;
                                                    var cell = row.Cells[targetCol];
                                                    if (cell == null) continue;
                                                    if (!cell.ReadOnly)
                                                    {
                                                        try { HideSelectionInGrid(dgv2); } catch { }
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                        catch { }
                                    }
                                }
                            }
                            catch { }
                        }));
                    }
                    catch { }


                    // Apply appearance adjustments
                    try { AdjustGridAppearance(); } catch { }

                    // If unlocked, focus first editable cell and begin edit
                    if (!locked)
                    {
                        try
                        {
                            int targetCol = shippedCol;
                            // fallback: find header containing '數量'
                            if (targetCol < 0)
                            {
                                for (int ci = 0; ci < dgv.Columns.Count; ci++)
                                {
                                    try
                                    {
                                        var col = dgv.Columns[ci];
                                        var hdr = (col?.HeaderText ?? col?.Name ?? string.Empty).ToString();
                                        if (!string.IsNullOrWhiteSpace(hdr) && hdr.IndexOf("數量", StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            targetCol = ci; break;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            // fallback: first numeric-typed column
                            if (targetCol < 0)
                            {
                                for (int ci = 0; ci < dgv.Columns.Count; ci++)
                                {
                                    try
                                    {
                                        var col = dgv.Columns[ci];
                                        var vt = col?.ValueType;
                                        if (vt == typeof(int) || vt == typeof(long) || vt == typeof(decimal) || vt == typeof(double) || vt == typeof(float))
                                        {
                                            targetCol = ci; break;
                                        }
                                    }
                                    catch { }
                                }
                            }

                            if (targetCol >= 0 && targetCol < dgv.Columns.Count)
                            {
                                foreach (DataGridViewRow row in dgv.Rows)
                                {
                                    if (row == null || row.IsNewRow) continue;
                                    var cell = row.Cells[targetCol];
                                    if (cell == null) continue;
                                    // Only begin edit if cell is editable
                                    if (!cell.ReadOnly)
                                    {
                                        try
                                        {
                                            int rowIndex = row.Index;
                                            try
                                            {
                                                if (dgv.Rows.Count > 0 && dgv.FirstDisplayedScrollingRowIndex != rowIndex)
                                                    dgv.FirstDisplayedScrollingRowIndex = Math.Max(0, rowIndex - 2);
                                            }
                                            catch { }

                                            try { HideSelectionInGrid(dgv); } catch { }

                                        }
                                        catch { }
                                        break;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            // 若在程式性操作前 DataTable 無變更，但現在出現變更，則視為程式性造成，呼叫 AcceptChanges 清除修改標記
            try
            {
                var dtAfter = this.dgv備料單?.DataSource as DataTable;
                bool dtHasChangesAfter = dtAfter != null && dtAfter.GetChanges() != null;
                if (!dtHadChangesBefore && dtHasChangesAfter)
                {
                    try
                    {
                        if (dtAfter is not null)
                        {
                            dtAfter.AcceptChanges();
                        }
                    }
                    catch { }
                }
            }
            catch { }

            try { this.btn備料單Unlock.Text = locked ? "解鎖" : "鎖定"; } catch { }
            // 同步料號輸入框的可編輯狀態：只有在畫面處於鎖定(locked==true)時允許輸入料號
            try
            {
                if (this.txt備料單料號 != null)
                {
                    // locked == true -> allow input -> ReadOnly = false
                    this.txt備料單料號.ReadOnly = !locked;
                }
            }
            catch { }
            // 恢復 dirty 標記
            try { _suspendDirtyMarking = false; } catch { }

            // 【關鍵】：在編輯鎖定狀態改變後，立即更新按鈕狀態
            // 確保鎖定/解鎖後，匯出/存檔/匯入按鈕狀態與編輯鎖定狀態保持同步
            try { UpdateMainButtonsEnabled(); } catch { }
        }

        /// <summary>
        /// 根據從 Excel 擷取的對齊資訊或內容推斷，將對齊套用到 DataGridView 欄位
        /// </summary>
        private void ApplyAlignmentsToGrid()
        {
            var dgv = this.dgv備料單;
            if (dgv == null || dgv.Columns == null) return;

            int colCount = dgv.Columns.Count;
            var decided = new DataGridViewContentAlignment[colCount];
            for (int i = 0; i < colCount; i++) decided[i] = DataGridViewContentAlignment.NotSet;

            // 優先使用 _excelAlignments 中的第一個有效對齊值（跳過標題列）
            if (_excelAlignments != null && _excelAlignments.Count > 1)
            {
                for (int c = 0; c < colCount; c++)
                {
                    for (int r = 1; r < _excelAlignments.Count; r++)
                    {
                        if (_excelAlignments[r] != null && c < _excelAlignments[r].Count)
                        {
                            var a = _excelAlignments[r][c];
                            if (a != DataGridViewContentAlignment.NotSet)
                            {
                                decided[c] = a;
                                break;
                            }
                        }
                    }
                }
            }

            // 對於仍未決定的欄位，根據欄位類型或內容推斷
            for (int c = 0; c < colCount; c++)
            {
                if (decided[c] != DataGridViewContentAlignment.NotSet) continue;
                var col = dgv.Columns[c];
                bool isNumericType = false;
                try { var vt = col.ValueType; if (vt == typeof(int) || vt == typeof(long) || vt == typeof(decimal) || vt == typeof(double) || vt == typeof(float)) isNumericType = true; } catch { }

                if (isNumericType)
                {
                    decided[c] = DataGridViewContentAlignment.MiddleRight;
                }
                else
                {
                    // sample some cells to see if majority numeric
                    int samples = 0; int numericCount = 0; int maxSample = Math.Min(100, dgv.Rows.Count);
                    for (int r = 0; r < maxSample; r++)
                    {
                        try
                        {
                            var row = dgv.Rows[r];
                            if (row == null || row.IsNewRow) continue;
                            var v = row.Cells[c].Value?.ToString();
                            if (string.IsNullOrWhiteSpace(v)) continue;
                            samples++;
                            if (decimal.TryParse(v.Replace(",", ""), System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out _)) numericCount++;
                        }
                        catch { }
                    }
                    if (samples > 0 && numericCount * 2 >= samples) decided[c] = DataGridViewContentAlignment.MiddleRight; else decided[c] = DataGridViewContentAlignment.MiddleLeft;
                }
            }

            // 套用到欄位 DefaultCellStyle
            for (int c = 0; c < colCount; c++)
            {
                try
                {
                    dgv.Columns[c].DefaultCellStyle.Alignment = decided[c];
                    // 標題居中
                    dgv.Columns[c].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                }
                catch { }
            }
        }

        /// <summary>
        /// 調整資料表格的外觀樣式。
        /// </summary>
        /// <remarks>
        /// 此方法會根據目前的顯示需求，設定欄位寬度、顏色或其他視覺效果，提升使用者體驗。
        /// </remarks>
        private void AdjustGridAppearance()
        {
            var dgv = this.dgv備料單;
            if (dgv == null) return;
            dgv.SuspendLayout();
            try
            {
                dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    var col = dgv.Columns[i];
                    var t = col?.ValueType;
                    if (t == typeof(int) || t == typeof(long) || t == typeof(decimal) || t == typeof(double))
                    {
                        if (col is not null && col.DefaultCellStyle is not null)
                        {
                            col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                        }
                    }
                }
                // 控制整表是否為唯讀，改為依畫面層的編輯鎖定狀態判斷。
                // 如果畫面處於鎖定，則整表唯讀；若解鎖，讓 per-column ReadOnly 決定哪些欄位可編輯。
                try { dgv.ReadOnly = _isEditingLocked; } catch { dgv.ReadOnly = true; }
                dgv.RowHeadersVisible = false;
                dgv.EnableHeadersVisualStyles = false;
                dgv.GridColor = Color.LightGray;
            }
            finally { try { dgv.ResumeLayout(); } catch { } }
        }

        #endregion

        /// <summary>
        /// 顯示一個密碼輸入對話視窗，若密碼錯誤會在視窗內顯示錯誤訊息並允許重試；使用者也可自行按取消關閉視窗。
        /// 回傳 true 表示密碼驗證成功並由使用者關閉視窗；false 表示使用者取消或視窗關閉時未通過驗證。
        /// </summary>
        /// <param name="expected">預期的密碼</param>
        /// <returns></returns>
        private bool PromptForPasswordWithRetry(string expected)
        {
            if (string.IsNullOrEmpty(expected)) return false;
            try
            {
                using (var dlg = new System.Windows.Forms.Form())
                {
                    dlg.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                    dlg.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
                    dlg.MinimizeBox = false;
                    dlg.MaximizeBox = false;
                    dlg.ShowInTaskbar = false;
                    dlg.ClientSize = new System.Drawing.Size(380, 150);
                    dlg.Text = "Excel 解鎖";

                    var lbl = new System.Windows.Forms.Label() { Left = 12, Top = 12, Width = 352, Height = 20, Text = "請輸入密碼以解鎖：" };
                    var txt = new System.Windows.Forms.TextBox() { Left = 12, Top = 36, Width = 352 }; txt.UseSystemPasswordChar = true;
                    var err = new System.Windows.Forms.Label() { Left = 12, Top = 66, Width = 352, Height = 24, ForeColor = System.Drawing.Color.Red, Text = string.Empty };
                    var btnOk = new System.Windows.Forms.Button() { Text = "確定", Left = 200, Width = 80, Top = 100, DialogResult = System.Windows.Forms.DialogResult.None };
                    var btnCancel = new System.Windows.Forms.Button() { Text = "取消", Left = 292, Width = 80, Top = 100, DialogResult = System.Windows.Forms.DialogResult.Cancel };

                    dlg.Controls.Add(lbl);
                    dlg.Controls.Add(txt);
                    dlg.Controls.Add(err);
                    dlg.Controls.Add(btnOk);
                    dlg.Controls.Add(btnCancel);

                    dlg.AcceptButton = btnOk;
                    dlg.CancelButton = btnCancel;

                    bool success = false;

                    btnOk.Click += (s, e) =>
                    {
                        try
                        {
                            var v = txt.Text ?? string.Empty;
                            if (!string.IsNullOrEmpty(v) && v == expected)
                            {
                                success = true;
                                dlg.DialogResult = System.Windows.Forms.DialogResult.OK;
                                dlg.Close();
                            }
                            else
                            {
                                err.Text = "密碼錯誤，請再試一次或按取消離開。";
                                txt.SelectAll();
                                txt.Focus();
                            }
                        }
                        catch { }
                    };
                    var result = dlg.ShowDialog();
                    return success;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 先以內容計算每欄的最佳寬度，然後切換到 Fill 模式讓欄位平均/按權重填滿整個 DataGridView，並避免出現水平捲軸。
        /// 同時設定 MinimumWidth 以防某些欄位被壓縮過小。只顯示垂直捲軸。
        /// </summary>
        /// <param name="dgv">目標 DataGridView</param>
        private void AutoSizeColumnsFillNoHorizontalScroll(DataGridView dgv)
        {
            if (dgv == null) return;
            try
            {
                // 先確保 layout 完成
                dgv.SuspendLayout();

                // 計算每個欄位所需的最佳寬度（依內容與欄頭）
                int colCount = dgv.Columns.Count;
                if (colCount == 0) return;

                var preferredWidths = new int[colCount];
                int totalPreferred = 0;
                for (int i = 0; i < colCount; i++)
                {
                    try
                    {
                        // 改用 DisplayedCells 僅估算目前畫面可見列，避免掃描整個資料集造成延遲
                        // 後續仍會切換為 Fill 以平均分配寬度，視覺上差異極小但可大幅減少計算時間
                        int pref = dgv.Columns[i].GetPreferredWidth(DataGridViewAutoSizeColumnMode.DisplayedCells, true);
                        // 最小寬度保護
                        pref = Math.Max(pref, 24);
                        preferredWidths[i] = pref;
                        totalPreferred += pref;
                    }
                    catch
                    {
                        preferredWidths[i] = dgv.Columns[i].Width;
                        totalPreferred += preferredWidths[i];
                    }
                }

                // 計算可用寬度（扣除垂直捲軸可能佔的空間）
                int totalWidth = dgv.ClientSize.Width;
                int vscrollWidth = SystemInformation.VerticalScrollBarWidth;
                bool needVScroll = dgv.Rows.Count > dgv.DisplayedRowCount(true);
                if (needVScroll) totalWidth -= vscrollWidth;

                if (totalPreferred <= 0) totalPreferred = 1;

                if (totalPreferred <= totalWidth && totalWidth > 0)
                {
                    // 如果所有 preferred 寬度加起來可以放得下，使用 Fill 並依 preferred 比例分配 FillWeight
                    dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    // 設定 MinimumWidth 與 FillWeight
                    for (int i = 0; i < colCount; i++)
                    {
                        var col = dgv.Columns[i];
                        try { col.MinimumWidth = preferredWidths[i]; } catch { }
                        try { col.FillWeight = Math.Max(1, preferredWidths[i]); } catch { }
                    }
                    dgv.ScrollBars = ScrollBars.Vertical;
                }
                else if (totalWidth > 0)
                {
                    // 否則把每個欄位寬度設定為 preferred，允許水平捲軸
                    dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                    for (int i = 0; i < colCount; i++)
                    {
                        try { dgv.Columns[i].Width = preferredWidths[i]; } catch { }
                    }
                    dgv.ScrollBars = ScrollBars.Both;
                }
                // 無額外全域覆寫，ScrollBars 已於各分支中設定
            }
            catch { }
        }


        // 替換此方法：MarkShortagesInGrid
        /// <summary>
        /// 標記資料表格中缺料的項目。
        /// </summary>
        /// <remarks>
        /// 此方法會根據缺料條件，將相關欄位或列以特殊樣式標示，方便使用者辨識。
        /// </remarks>
        private void MarkShortagesInGrid()
        {
            var dgv = this.dgv備料單;
            if (dgv == null || dgv.Rows == null) return;

            int demandCol = FindColumnIndexByNames(new[] { "需求數量", "應領數量" });
            int shippedCol = FindColumnIndexByNames(new[] { "實發數量", "發料數量" });
            if (demandCol < 0 || shippedCol < 0) return;

            // 儲存預設顏色以便還原（使用 DataGridView 的預設樣式作為基底）
            var defaultBack = dgv.DefaultCellStyle?.BackColor ?? SystemColors.Window;
            var defaultFore = dgv.DefaultCellStyle?.ForeColor ?? SystemColors.ControlText;

            dgv.SuspendLayout();
            try
            {
                var red = Color.Red;
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row == null || row.IsNewRow) continue;

                    decimal demand = 0m, shipped = 0m;
                    bool hasDemand = false, hasShipped = false;

                    if (demandCol >= 0)
                        hasDemand = TryParseDecimalFlexible(row.Cells[demandCol].Value?.ToString() ?? string.Empty, out demand);
                    if (shippedCol >= 0)
                        hasShipped = TryParseDecimalFlexible(row.Cells[shippedCol].Value?.ToString() ?? string.Empty, out shipped);

                    bool shortage = false;
                    // 只有當需求大於 0 且實發/發料數量可解析且大於 0 時，才判斷是否短缺
                    // 這樣可以避免空值或 0 被誤標為短缺
                    if (hasDemand && demand > 0m && hasShipped && shipped > 0m)
                    {
                        if (shipped < demand) shortage = true;
                    }

                    // 只對 "實發/發料" 那一個 cell 進行上色，避免改變整列樣式
                    if (shippedCol >= 0 && shippedCol < row.Cells.Count)
                    {
                        var cell = row.Cells[shippedCol];
                        try
                        {
                            if (shortage)
                            {
                                cell.Style.BackColor = red;
                                cell.Style.ForeColor = Color.Black;
                                try
                                {
                                    // 儲存紅色 key
                                    string? mat = null;
                                    int chCol = FindColumnIndexByNames(new[] { "昶亨料號" });
                                    int custCol = FindColumnIndexByNames(new[] { "客戶料號" });
                                    if (chCol >= 0 && chCol < row.Cells.Count) mat = row.Cells[chCol].Value?.ToString()?.Trim();
                                    if (string.IsNullOrEmpty(mat) && custCol >= 0 && custCol < row.Cells.Count) mat = row.Cells[custCol].Value?.ToString()?.Trim();
                                    if (!string.IsNullOrEmpty(mat) && _preservedRedKeys != null) _preservedRedKeys.Add(TextParsing.NormalizeMaterialKey(mat!));
                                }
                                catch { }
                            }
                            else
                            {
                                // 只有在可以明確判定『不再短缺』時才移除先前的紅色標示
                                // 如果缺乏可解析的 demand 或 shipped 值，保留先前的紅色狀態以避免在儲存失敗或讀取不全時移除標示
                                if (cell.Style.BackColor == red)
                                {
                                    bool canConfirmNoShortage = false;
                                    try
                                    {
                                        // 僅當 demand 與 shipped 都可解析，且 shipped >= demand（或需求為 0）時，才可以確定不是短缺
                                        if (hasDemand && hasShipped)
                                        {
                                            if (!(shipped < demand)) canConfirmNoShortage = true;
                                        }
                                        else
                                        {
                                            // 無法確認數值，保留紅色
                                            canConfirmNoShortage = false;
                                        }
                                    }
                                    catch { canConfirmNoShortage = false; }

                                    if (canConfirmNoShortage)
                                    {
                                        try
                                        {
                                            string? mat = null;
                                            int chCol = FindColumnIndexByNames(new[] { "昶亨料號" });
                                            int custCol = FindColumnIndexByNames(new[] { "客戶料號" });
                                            if (chCol >= 0 && chCol < row.Cells.Count) mat = row.Cells[chCol].Value?.ToString()?.Trim();
                                            if (string.IsNullOrEmpty(mat) && custCol >= 0 && custCol < row.Cells.Count) mat = row.Cells[custCol].Value?.ToString()?.Trim();
                                            if (!string.IsNullOrEmpty(mat) && _preservedRedKeys != null) _preservedRedKeys.Remove(TextParsing.NormalizeMaterialKey(mat!));
                                        }
                                        catch { }

                                        // 顯式還原為 DataGridView 預設顏色，避免 Color.Empty 在某些佈景導致黑底
                                        var __defBack = dgv.DefaultCellStyle?.BackColor ?? SystemColors.Window;
                                        var __defFore = dgv.DefaultCellStyle?.ForeColor ?? SystemColors.ControlText;
                                        cell.Style.BackColor = __defBack;
                                        cell.Style.ForeColor = __defFore;
                                    }
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            finally { try { dgv.ResumeLayout(); } catch { } }
        }


        /// <summary>
        /// 讀取 app.config 的開關，判斷是否啟用某日誌（含快取以降低高頻呼叫成本）
        /// </summary>
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, bool> _logEnabledCache =
            new System.Collections.Concurrent.ConcurrentDictionary<string, bool>(System.StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 判斷指定的 key 是否啟用日誌功能。
        /// </summary>
        /// <param name="key">要檢查的日誌設定鍵值。</param>
        /// <returns>
        /// <see langword="true" /> 表示該 key 已啟用日誌；否則為 <see langword="false" />。
        /// </returns>
        private static bool IsLogEnabled(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return false;

            if (_logEnabledCache.TryGetValue(key, out bool cached))
            {
                return cached;
            }

            try
            {
                var raw = System.Configuration.ConfigurationManager.AppSettings[key];
                bool parsed = !string.IsNullOrWhiteSpace(raw) && raw.Trim().Equals("true", StringComparison.OrdinalIgnoreCase);
                _logEnabledCache[key] = parsed;
                return parsed;
            }
            catch
            {
                // don't cache failures as true; return false
                _logEnabledCache[key] = false;
                return false;
            }
        }

        /// <summary>
        /// 重置表單狀態：清空 DataGridView、欄位、狀態變數等，回到初始狀態。
        /// </summary>
        private void ResetFormState()
        {
            try
            {
                // 清空 DataGridView
                if (this.dgv備料單 != null)
                {
                    if (this.dgv備料單.DataSource is DataTable dt)
                    {
                        dt.Clear();
                        this.dgv備料單.DataSource = null;
                    }
                    this.dgv備料單.Rows.Clear();
                    this.dgv備料單.Refresh();
                }

                // 清空記錄與狀態
                _records?.Clear();
                _lastAppendedRecords?.Clear();
                _lastMatchedRows?.Clear();
                _materialIndex?.Clear();
                _preservedRedKeys?.Clear();
                _excelAlignments = null;
                _currentUiTable = null;
                _isDirty = false;
                currentExcelPath = null;
                _columnMapping = null;
                _lastHiddenHeaders = null;
                // 其他欄位如有需要可依實際情境補充

                // 重設 UI 狀態
                try { SetEditingLocked(true); } catch { }
                try { RestoreAllButtons(); } catch { }
                try { HideOperationOverlay(); } catch { }
                try { UpdateMainButtonsEnabled(); } catch { }

                // 停用主要按鈕以防誤操作（匯入按鈕不強制設為 false，交由 UpdateMainButtonsEnabled 控制）
                try { SetControlEnabledSafe(this.btn備料單Unlock, false); } catch { }
                try { SetControlEnabledSafe(this.btn備料單匯出, false); } catch { }
                try { SetControlEnabledSafe(this.btn備料單存檔, false); } catch { }
                // 立即依狀態更新匯入按鈕（解鎖且無操作時可按）
                try { UpdateMainButtonsEnabled(); } catch { }
            }
            catch { }
        }

        /// <summary>
        /// 根據 dgv備料單 是否有資料，啟用/停用存檔、匯出、鎖定按鈕（防呆）
        /// 按照按鈕文字判斷狀態：
        /// - 按鈕顯示「解鎖」：_isEditingLocked=true（鎖定）→ 匯出、存檔可按（有資料時）
        /// - 按鈕顯示「鎖定」：_isEditingLocked=false（解鎖）→ 匯出、存檔不可按
        /// - 匯入按鈕：隨時可按（只要沒有操作進行中）
        /// </summary>
        private void UpdateMainButtonsEnabled()
        {
            try
            {
                // 匯入按鈕：隨時可按，只要沒有操作進行中，不受鎖定狀態影響
                bool canImport = !_isImporting && !_isExporting && !_isSaving;
                SetControlEnabledSafe(this.btn備料單匯入檔案, canImport);

                // 先檢查 UI 是否有資料
                var dgv = this.dgv備料單;
                bool hasRows = false;
                try
                {
                    if (dgv != null && dgv.DataSource is DataTable dt && dt.Rows != null && dt.Rows.Count > 0)
                    {
                        hasRows = true;
                    }
                }
                catch { }
                if (!hasRows && dgv != null && dgv.Rows != null)
                {
                    foreach (DataGridViewRow row in dgv.Rows)
                    {
                        if (row != null && !row.IsNewRow)
                        {
                            hasRows = true;
                            break;
                        }
                    }
                }
                try
                {
                    if (!hasRows && _currentUiTable != null && _currentUiTable.Rows != null && _currentUiTable.Rows.Count > 0)
                        hasRows = true;
                }
                catch { }

                // 匯出/存檔：按鈕顯示「解鎖」時（_isEditingLocked=true，即鎖定狀態）
                // 且無操作進行中且有資料時才可按
                bool canOperateWithData = (_isEditingLocked) && !_isImporting && !_isExporting && !_isSaving && hasRows;
                SetControlEnabledSafe(this.btn備料單匯出, canOperateWithData);
                SetControlEnabledSafe(this.btn備料單存檔, canOperateWithData);

                // Unlock 按鈕：有資料且沒有其他操作進行中時可按
                // 若在匯出成功後設定了 _keepUnlockButtonDisabledUntilClose，則強制保持不可按
                SetControlEnabledSafe(this.btn備料單Unlock, hasRows && !_isImporting && !_isExporting && !_isSaving && !_keepUnlockButtonDisabledUntilClose);
            }
            catch { }
        }

        #region Excel I/O

        /// <summary>
        /// 使用 COM 開啟 Excel，針對 uiTable 中看起來為公式的字串（以 '=' 開頭）讀取該儲存格的 Value2，
        /// 並以讀到的值覆寫 DataTable。此方法可被取消。
        /// </summary>
        /// <param name="excelPath">Excel 檔案路徑</param>
        /// <param name="uiTable">欲處理的 DataTable（會被 in-place 修改）</param>
        /// <param name="token">取消 Token</param>
        private void ReplaceFormulasWithValuesFromExcel(string excelPath, DataTable uiTable, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath)) return;
            if (uiTable == null || uiTable.Rows.Count == 0) return;

            // Timing instrumentation removed: swHeader/swScan/swRead/swWrite
            // var swHeader = System.Diagnostics.Stopwatch.StartNew();
            // var swScan = new System.Diagnostics.Stopwatch();
            // var swRead = new System.Diagnostics.Stopwatch();
            // var swWrite = new System.Diagnostics.Stopwatch();
            int formulaCellCount = 0;

            // Fast path: for .xlsx/.xls use NPOI to read values quickly and avoid COM;
            // Optional: also prefer NPOI for .xlsm (controlled by appSettings: PreferNpoiForXlsm=true)
            var ext = Path.GetExtension(excelPath)?.ToLowerInvariant() ?? string.Empty;
            bool preferNpoiForXlsm = false;
            try { preferNpoiForXlsm = IsLogEnabled("PreferNpoiForXlsm"); } catch { preferNpoiForXlsm = false; }
            if (ext == ".xlsx" || ext == ".xls" || (ext == ".xlsm" && preferNpoiForXlsm))
            {
                try
                {
                    // NPOI timers (declare here so message box can reference them)
                    // NPOI timing removed (swNpoiRead/swNpoiWrite)
                    // var swNpoiRead = new System.Diagnostics.Stopwatch();
                    // var swNpoiWrite = new System.Diagnostics.Stopwatch();
                    // Read sheet with NPOI and write values into uiTable
                    using (var fs = File.OpenRead(excelPath))
                    {
                        IWorkbook nwb = null;
                        // .xls → HSSFWorkbook；.xlsx / .xlsm(當 PreferNpoiForXlsm=true) → XSSFWorkbook
                        if (ext == ".xls")
                        {
                            nwb = new HSSFWorkbook(fs);
                        }
                        else // .xlsx 或 .xlsm（在 preferNpoiForXlsm 啟用時會進此分支）
                        {
                            nwb = new XSSFWorkbook(fs);
                        }
                        if (nwb == null || nwb.NumberOfSheets <= 0) throw new InvalidOperationException("No sheets");
                        var sheet = nwb.GetSheetAt(0);
                        // Determine header row similar to other loaders
                        int headerRowIdx = -1; int bestCount = -1; int scanLimit = Math.Min(10, sheet.LastRowNum + 1);
                        try
                        {
                            if (sheet.LastRowNum >= 2)
                            {
                                var prefRow = sheet.GetRow(2);
                                if (prefRow != null)
                                {
                                    int cnt = 0;
                                    int first = prefRow.FirstCellNum >= 0 ? prefRow.FirstCellNum : 0;
                                    int last = prefRow.LastCellNum >= 0 ? prefRow.LastCellNum : first;
                                    for (int c = first; c <= last; c++) { var cell = prefRow.GetCell(c); if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString())) cnt++; }
                                    if (cnt > 0) headerRowIdx = 2;
                                }
                            }
                        }
                        catch { headerRowIdx = -1; }
                        if (headerRowIdx < 0)
                        {
                            for (int r = 0; r < scanLimit; r++)
                            {
                                var row = sheet.GetRow(r);
                                if (row == null) continue;
                                int count = 0;
                                int first = row.FirstCellNum >= 0 ? row.FirstCellNum : 0;
                                int last = row.LastCellNum >= 0 ? row.LastCellNum : first;
                                for (int c = first; c <= last; c++) { var cell = row.GetCell(c); if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString())) count++; }
                                if (count > bestCount) { bestCount = count; headerRowIdx = r; }
                            }
                        }
                        if (headerRowIdx < 0) throw new InvalidOperationException("No header");
                        // 以欄名對齊：建立 NPOI 欄位索引 -> DataTable 欄位索引 的映射
                        var headerRow = sheet.GetRow(headerRowIdx);
                        int maxCol = headerRow?.LastCellNum >= 0 ? headerRow.LastCellNum : 0;
                        // Use centralized sanitizer for header normalization to keep behaviour consistent
                        var uiNameToIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        for (int i = 0; i < uiTable.Columns.Count; i++)
                        {
                            var nm = uiTable.Columns[i].ColumnName ?? string.Empty;
                            var san = SanitizeHeaderForMatch(nm);
                            if (!uiNameToIndex.ContainsKey(san)) uiNameToIndex[san] = i;
                        }
                        var columnMap = new List<(int uiIdx, int npoiIdx, string name)>();
                        for (int c = 0; c < maxCol; c++)
                        {
                            var cell = headerRow?.GetCell(c);
                            var rawName = cell?.ToString();
                            if (string.IsNullOrWhiteSpace(rawName)) continue;
                            var san = SanitizeHeaderForMatch(rawName);
                            if (string.IsNullOrEmpty(san)) continue;
                            if (uiNameToIndex.TryGetValue(san, out int uiIdx))
                            {
                                columnMap.Add((uiIdx, c, rawName));
                            }
                        }
                        if (columnMap.Count == 0)
                        {
                            throw new InvalidOperationException("Header mapping failed between Excel and DataTable");
                        }

                        // Write values to DataTable rows (matching uiTable rows by index)
                        int writeRow = 0;
                        // swNpoiRead.Start(); // timing removed
                        try
                        {
                            // 加速大量寫入：暫停約束/索引/事件
                            uiTable.BeginLoadData();
                            for (int r = headerRowIdx + 1; r <= sheet.LastRowNum && writeRow < uiTable.Rows.Count; r++)
                            {
                                var row = sheet.GetRow(r);
                                if (row == null) continue;
                                var dr = uiTable.Rows[writeRow];
                                bool any = false;
                                for (int i = 0; i < columnMap.Count; i++)
                                {
                                    var (uiIdx, npoiIdx, _) = columnMap[i];
                                    var cell = row.GetCell(npoiIdx);
                                    object val = null;
                                    try
                                    {
                                        if (cell != null)
                                        {
                                            // Count formula cells and try to use cached formula result when available
                                            if (cell.CellType == NPOI.SS.UserModel.CellType.Formula)
                                            {
                                                formulaCellCount++;
                                                var cached = cell.CachedFormulaResultType;
                                                switch (cached)
                                                {
                                                    case NPOI.SS.UserModel.CellType.Numeric:
                                                        val = cell.NumericCellValue; break;
                                                    case NPOI.SS.UserModel.CellType.String:
                                                        val = cell.StringCellValue; break;
                                                    case NPOI.SS.UserModel.CellType.Boolean:
                                                        val = cell.BooleanCellValue; break;
                                                    case NPOI.SS.UserModel.CellType.Error:
                                                        val = cell.ErrorCellValue; break;
                                                    default:
                                                        val = cell?.ToString() ?? string.Empty; break;
                                                }
                                            }
                                            else
                                            {
                                                switch (cell.CellType)
                                                {
                                                    case NPOI.SS.UserModel.CellType.Numeric: val = cell.NumericCellValue; break;
                                                    case NPOI.SS.UserModel.CellType.String: val = cell.StringCellValue; break;
                                                    case NPOI.SS.UserModel.CellType.Boolean: val = cell.BooleanCellValue; break;
                                                    case NPOI.SS.UserModel.CellType.Error: val = cell.ErrorCellValue; break;
                                                    default: val = cell?.ToString() ?? string.Empty; break;
                                                }
                                            }
                                        }
                                    }
                                    catch { val = cell?.ToString() ?? string.Empty; }

                                    // write value (preserve types where reasonable) - 以 uiIdx 寫入對應 DataTable 欄位
                                    try
                                    {
                                        // swNpoiWrite.Start(); // timing removed
                                        if (val == null) { dr[uiIdx] = DBNull.Value; }
                                        else
                                        {
                                            // treat empty strings as DB null
                                            if (val is string s)
                                            {
                                                if (string.IsNullOrWhiteSpace(s)) dr[uiIdx] = DBNull.Value; else dr[uiIdx] = s;
                                                any = any || !string.IsNullOrWhiteSpace(s);
                                            }
                                            else
                                            {
                                                dr[uiIdx] = val;
                                                any = true;
                                            }
                                        }
                                    }
                                    catch { }
                                    finally { try { /* timing removed */ } catch { } }
                                }
                                if (any) writeRow++; // only advance when row had any data
                            }
                        }
                        finally
                        {
                            try { uiTable.EndLoadData(); } catch { }
                            try { /* timing removed */ } catch { }
                        }
                    }
                    // Timing removed for scan/read/write/header
                    try { /* timing removed: swScan/ swRead/ swWrite/ swHeader */ } catch { }
                    // 測試用效能分析提示已移除
                    return;
                }
                catch (Exception ex)
                {
                    // On any failure, fall back to Interop chunked path below
                    try
                    {
                        var sb = new System.Text.StringBuilder();
                        sb.AppendLine($"Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                        sb.AppendLine($"File: {excelPath}");
                        sb.AppendLine("Exception:");
                        sb.AppendLine(ex.ToString());
                        // Debug write removed
                    }
                    catch { }
                }
            }

            Excel.Application? xlApp = null; Excel.Workbook? wb = null; Excel.Worksheet? ws = null; Excel.Range? used = null;
            // 優化：在 Interop 路徑中關閉畫面更新/事件與計算，以縮短讀取時間（結束時還原）
            bool prevScreenUpdating = true; bool prevEnableEvents = true; Excel.XlCalculation prevCalc = Excel.XlCalculation.xlCalculationAutomatic;
            try
            {
                xlApp = new Excel.Application { DisplayAlerts = false, Visible = false }; // headless
                try { prevScreenUpdating = xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
                try { prevEnableEvents = xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
                try { prevCalc = xlApp.Calculation; xlApp.Calculation = Excel.XlCalculation.xlCalculationManual; } catch { }
                wb = xlApp.Workbooks.Open(excelPath, ReadOnly: true);
                try { ws = wb.Worksheets["總表"] as Excel.Worksheet; } catch { ws = null; }
                if (ws == null) ws = wb.Worksheets[1] as Excel.Worksheet;
                if (ws == null) return;

                used = ws.UsedRange; if (used == null) return;
                int rowCount = used.Rows.Count; int colCount = used.Columns.Count;

                // 偵測 Excel 的標題列位置：比對 uiTable 的欄名與 Excel 某一列的文字，優先找 1..5 列
                int excelHeaderRow = 1;
                try
                {
                    int maxHeaderCheck = Math.Min(5, rowCount);
                    int limitCols = Math.Min(colCount, uiTable.Columns.Count);
                    Excel.Range? headerRange = null;
                    object headerValsObj = null;
                    try { headerRange = used.Range[used.Cells[1, 1], used.Cells[maxHeaderCheck, limitCols]]; } catch { headerRange = null; }
                    try { headerValsObj = headerRange?.Value2; } catch { headerValsObj = null; }

                    int bestMatchCount = -1;
                    if (headerValsObj is object[,] headerVals)
                    {
                        for (int hr = 1; hr <= maxHeaderCheck; hr++)
                        {
                            int match = 0;
                            for (int c = 1; c <= limitCols; c++)
                            {
                                try
                                {
                                    var hv = headerVals[hr, c];
                                    if (hv == null) continue;
                                    var hs = hv?.ToString()?.Trim() ?? string.Empty;
                                    var colName = uiTable.Columns[c - 1].ColumnName?.Trim() ?? string.Empty;
                                    if (string.IsNullOrEmpty(colName)) continue;
                                    if (string.Equals(hs, colName, StringComparison.OrdinalIgnoreCase) || hs.IndexOf(colName, StringComparison.OrdinalIgnoreCase) >= 0 || colName.IndexOf(hs, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        match++;
                                    }
                                }
                                catch { }
                            }
                            if (match > bestMatchCount)
                            {
                                bestMatchCount = match;
                                excelHeaderRow = hr;
                            }
                        }
                    }
                    else if (headerValsObj != null)
                    {
                        // 單列或單欄時 Value2 可能不是二維陣列，保底處理第一列
                        int match = 0;
                        for (int c = 1; c <= limitCols; c++)
                        {
                            try
                            {
                                var hv = (used.Cells[1, c] as Excel.Range)?.Value2;
                                if (hv == null) continue;
                                var hs = hv?.ToString()?.Trim() ?? string.Empty;
                                var colName = uiTable.Columns[c - 1].ColumnName?.Trim() ?? string.Empty;
                                if (string.IsNullOrEmpty(colName)) continue;
                                if (string.Equals(hs, colName, StringComparison.OrdinalIgnoreCase) || hs.IndexOf(colName, StringComparison.OrdinalIgnoreCase) >= 0 || colName.IndexOf(hs, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    match++;
                                }
                            }
                            catch { }
                        }
                        excelHeaderRow = 1;
                    }
                    try { if (headerRange != null) ReleaseComObjectSafe(headerRange); } catch { }
                }
                catch { excelHeaderRow = 1; }
                // timing removed: swHeader.Stop();

                int excelStartRow = Math.Max(2, excelHeaderRow + 1);
                int maxRowsToCheck = Math.Min(uiTable.Rows.Count, Math.Max(0, rowCount - excelStartRow + 1));
                // 以「欄名正規化」建立 Excel 欄 -> DataTable 欄 的映射，避免移除隱藏欄位後出現位移
                var uiMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < uiTable.Columns.Count; i++)
                {
                    var nm = uiTable.Columns[i].ColumnName ?? string.Empty;
                    var key = SanitizeHeaderForMatch(nm);
                    if (!uiMap.ContainsKey(key)) uiMap[key] = i;
                }
                var excelColToUi = new List<(int excelCol, int uiIdx)>();
                for (int c = 1; c <= colCount; c++)
                {
                    try
                    {
                        var hv = (used.Cells[excelHeaderRow, c] as Excel.Range)?.Value2;
                        var hn = hv?.ToString();
                        var key = SanitizeHeaderForMatch(hn);
                        if (string.IsNullOrEmpty(key)) continue;
                        if (uiMap.TryGetValue(key, out int uiIdx))
                        {
                            excelColToUi.Add((c, uiIdx));
                        }
                    }
                    catch { }
                }
                // 讀取到映射中最大欄位位置即可
                int maxColsToCheck = excelColToUi.Count > 0 ? excelColToUi.Max(p => p.excelCol) : Math.Min(uiTable.Columns.Count, colCount);

                // swScan.Start(); // timing removed

                // 以一次性批次讀取整個資料區塊的值，減少每個 cell 的 COM 呼叫
                Excel.Range dataRange = null;
                try
                {
                    dataRange = used.Range[used.Cells[excelStartRow, 1], used.Cells[excelStartRow + maxRowsToCheck - 1, maxColsToCheck]];
                }
                catch { dataRange = null; }

                // 不在此處一次性讀取整張表的 Value2，改為以 chunk 分段讀取以降低單次 I/O 與記憶體峰值

                // 僅在小範圍時嘗試計算公式儲存格數量（SpecialCells 對大範圍會非常慢）
                try
                {
                    if (dataRange != null && maxRowsToCheck * maxColsToCheck <= 10000)
                    {
                        try
                        {
                            var formulaRange = dataRange.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                            if (formulaRange != null)
                            {
                                try { formulaCellCount = Convert.ToInt32(formulaRange.Count); } catch { formulaCellCount = 0; }
                                try { if (formulaRange != null) ReleaseComObjectSafe(formulaRange); } catch { }
                            }
                        }
                        catch { formulaCellCount = 0; }
                    }
                }
                catch { formulaCellCount = 0; }

                // 使用分段 (chunk) 讀取/寫入，並使用 DataTable.BeginLoadData/EndLoadData 加快大量寫入
                // swWrite.Start(); // timing removed
                try
                {
                    uiTable.BeginLoadData();
                    // 自適應 chunk size：若資料量大，增加 chunkSize 減少 Range 建立次數
                    int chunkSize = 500; // default
                    try
                    {
                        if (maxRowsToCheck > 10000) chunkSize = 2000;
                        else if (maxRowsToCheck > 5000) chunkSize = 1000;
                        else if (maxRowsToCheck > 2000) chunkSize = 800;
                        else chunkSize = 500;
                    }
                    catch { chunkSize = 500; }
                    for (int offset = 0; offset < maxRowsToCheck; offset += chunkSize)
                    {
                        if (token.IsCancellationRequested) throw new OperationCanceledException();
                        int thisChunkRows = Math.Min(chunkSize, maxRowsToCheck - offset);

                        Excel.Range chunkRange = null;
                        object chunkValues = null;
                        try
                        {
                            chunkRange = used.Range[used.Cells[excelStartRow + offset, 1], used.Cells[excelStartRow + offset + thisChunkRows - 1, maxColsToCheck]];
                            // 讀取該區塊的值（計時並累加到 swRead）
                            try
                            {
                                // swRead.Start(); // timing removed
                                chunkValues = chunkRange.Value2;
                            }
                            catch { chunkValues = null; }
                            //finally { try { swRead.Stop(); } catch { } }
                        }
                        catch { chunkRange = null; chunkValues = null; }

                        if (chunkValues is object[,] carr)
                        {
                            int crow = carr.GetLength(0);
                            int ccol = carr.GetLength(1);
                            for (int r = 0; r < Math.Min(thisChunkRows, crow); r++)
                            {
                                if (token.IsCancellationRequested) throw new OperationCanceledException();
                                var dr = uiTable.Rows[offset + r];
                                // 僅根據映射寫入目標欄位
                                foreach (var (excelCol, uiIdx) in excelColToUi)
                                {
                                    try
                                    {
                                        if (excelCol <= 0 || excelCol > ccol) continue;
                                        var val = carr[r + 1, excelCol];
                                        if (val == null) { dr[uiIdx] = DBNull.Value; continue; }
                                        if (val is double || val is float || val is int || val is decimal)
                                        {
                                            if (uiTable.Columns[uiIdx].DataType == typeof(string)) dr[uiIdx] = val?.ToString() ?? string.Empty; else dr[uiIdx] = val;
                                        }
                                        else
                                        {
                                            dr[uiIdx] = val?.ToString() ?? string.Empty;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                        else if (chunkValues != null && chunkValues.GetType().IsArray)
                        {
                            try
                            {
                                var flat = chunkValues as object[];
                                if (flat != null)
                                {
                                    int idx = 0;
                                    for (int r = 0; r < thisChunkRows; r++)
                                    {
                                        if (token.IsCancellationRequested) throw new OperationCanceledException();
                                        var dr = uiTable.Rows[offset + r];
                                        for (int c = 0; c < maxColsToCheck; c++)
                                        {
                                            if (idx >= flat.Length) break;
                                            try
                                            {
                                                var val = flat[idx++];
                                                // 這個分支發生於單欄選取，為安全起見沿用舊寫法
                                                if (val == null) { dr[c] = DBNull.Value; continue; }
                                                if (val is double || val is float || val is int || val is decimal)
                                                {
                                                    if (uiTable.Columns[c].DataType == typeof(string)) dr[c] = val?.ToString() ?? string.Empty; else dr[c] = val;
                                                }
                                                else
                                                {
                                                    dr[c] = val?.ToString() ?? string.Empty;
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                        else
                        {
                            // fallback to per-cell read for this chunk
                            for (int r = 0; r < thisChunkRows; r++)
                            {
                                if (token.IsCancellationRequested) throw new OperationCanceledException();
                                var dr = uiTable.Rows[offset + r];
                                foreach (var (excelCol, uiIdx) in excelColToUi)
                                {
                                    try
                                    {
                                        Excel.Range cell = null;
                                        try { cell = (used.Cells[excelStartRow + offset + r, excelCol] as Excel.Range); } catch { cell = null; }
                                        if (cell == null) continue;
                                        object val = null;
                                        try { val = cell.Value2; } catch { val = null; }
                                        if (val == null) { dr[uiIdx] = DBNull.Value; continue; }
                                        if (val is double || val is float || val is int || val is decimal)
                                        {
                                            if (uiTable.Columns[uiIdx].DataType == typeof(string)) dr[uiIdx] = val?.ToString() ?? string.Empty; else dr[uiIdx] = val;
                                        }
                                        else
                                        {
                                            dr[uiIdx] = val?.ToString() ?? string.Empty;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }

                        try { if (chunkRange != null) ReleaseComObjectSafe(chunkRange); } catch { }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch { }
                finally
                {
                    try { uiTable.EndLoadData(); } catch { }
                }
                //swWrite.Stop();

                //swScan.Stop();

                try { if (dataRange != null) ReleaseComObjectSafe(dataRange); } catch { }
            }
            finally
            {
                try { if (used != null) ReleaseComObjectSafe(used); } catch { }
                try { if (ws != null) ReleaseComObjectSafe(ws); } catch { }
                try { if (wb != null) { wb.Close(false); ReleaseComObjectSafe(wb); } } catch { }
                // 還原 Excel 應用程式狀態
                try { if (xlApp != null) { try { xlApp.Calculation = prevCalc; } catch { } try { xlApp.EnableEvents = prevEnableEvents; } catch { } try { xlApp.ScreenUpdating = prevScreenUpdating; } catch { } } } catch { }
                try { if (xlApp != null) { xlApp.Quit(); ReleaseComObjectSafe(xlApp); } } catch { }
            }

            // 測試用效能分析提示已移除
        }

        #endregion

        #region Input & update

        #region Input Handlers

        /// <summary>
        /// 料號輸入框 KeyDown 事件: 按下 Enter 執行不區分大小寫的模糊比對
        /// </summary>
        private void Txt備料單料號_KeyDown(object sender, KeyEventArgs e)
        {
            // 禁止貼上
            if (e.Control && e.KeyCode == Keys.V) { e.SuppressKeyPress = true; return; }

            // 若目前畫面處於「可編輯」狀態（未鎖定），則禁止輸入料號，要求先按鎖定
            try
            {
                if (!_isEditingLocked)
                {
                    if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.None || e.KeyCode == Keys.Back || e.KeyCode == Keys.Delete || char.IsLetterOrDigit((char)e.KeyValue))
                    {
                        e.SuppressKeyPress = true;
                        try { SafeShowMessage("目前畫面為可編輯狀態；請先按『鎖定』後再輸入料號。", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                        try { if (this.btn備料單Unlock != null) this.btn備料單Unlock.Focus(); } catch { }
                        return;
                    }
                }
            }
            catch { }

            // 按下 Enter 執行料號比對（僅在畫面為鎖定狀態時執行）
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                string input = this.txt備料單料號.Text?.Trim();
                if (!string.IsNullOrEmpty(input))
                {
                    ProcessMaterialInput(input);
                }
            }
        }

        /// <summary>
        /// 處理備料單 DataGridView 的儲存格格式化事件。
        /// CellFormatting: 保證每次繪製時數字欄靠右，且根據短缺條件標示紅色
        /// </summary>
        /// <param name="sender">觸發事件的來源物件。</param>
        /// <param name="e">包含儲存格格式化相關資料的事件參數。</param>
        /// <remarks>
        /// 此方法用於自訂備料單 DataGridView 的儲存格顯示格式，例如顏色、字型或數值格式。
        /// </remarks>
        /// <example>
        /// <code language="csharp">
        /// // 範例：自訂儲存格顯示
        /// Dgv備料單_CellFormatting(sender, e);
        /// </code>
        /// </example>
        private void Dgv備料單_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                var dgv = sender as DataGridView;
                if (dgv == null) return;

                var __defaultBack = dgv.DefaultCellStyle?.BackColor ?? SystemColors.Window;
                var __defaultFore = dgv.DefaultCellStyle?.ForeColor ?? SystemColors.ControlText;

                int shippedCol = FindColumnIndexByNames(new[] { "實發數量", "發料數量" });

                // 根據檔案類型決定需求欄位
                int demandCol = -1;
                try
                {
                    var ext = Path.GetExtension(this.currentExcelPath ?? string.Empty).ToLowerInvariant();
                    if (ext == ".xlsm")
                    {
                        demandCol = FindColumnIndexByNames(new[] { "需求數量" });
                    }
                    else
                    {
                        demandCol = FindColumnIndexByNames(new[] { "應領數量" });
                    }
                }
                catch { demandCol = -1; }

                if (e.ColumnIndex == shippedCol)
                {
                    e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                    // --- 全新重構的邏輯 ---
                    // 步驟 1: 收集所有需要的資訊，但不做任何樣式變更

                    // 基本資訊
                    string cellValueStr = e.Value?.ToString()?.Trim() ?? string.Empty;
                    bool is_empty = string.IsNullOrEmpty(cellValueStr);
                    // style debug log removed

                    // 料號資訊
                    int chCol = FindColumnIndexByNames(new[] { "昶亨料號" });
                    int custCol = FindColumnIndexByNames(new[] { "客戶料號" });
                    string? matKey = null;
                    if (e.RowIndex >= 0 && e.RowIndex < dgv.Rows.Count)
                    {
                        if (chCol >= 0) matKey = dgv.Rows[e.RowIndex].Cells[chCol].Value?.ToString();
                        if (string.IsNullOrWhiteSpace(matKey) && custCol >= 0) matKey = dgv.Rows[e.RowIndex].Cells[custCol].Value?.ToString();
                    }
                    matKey = NormalizeMaterialKey(matKey ?? string.Empty);

                    // 狀態計算
                    bool is_preserved_red = !string.IsNullOrEmpty(matKey) && _preservedRedKeys.Contains(matKey);
                    bool is_equal = false;
                    bool is_shortage = false;
                    bool is_exceed = false;

                    if (!is_empty)
                    {
                        // 以先前建置的快取取得累計數量（若無快取則視為 0）
                        decimal sum = 0m;
                        try
                        {
                            if (!string.IsNullOrEmpty(matKey) && _materialShippedSums != null && _materialShippedSums.TryGetValue(matKey, out decimal cached))
                            {
                                sum = cached;
                            }
                            else
                            {
                                sum = 0m;
                            }
                        }
                        catch { sum = 0m; }

                        // 取得需求數量
                        if (demandCol >= 0 && demandCol < dgv.Columns.Count &&
                            TryParseDecimalFlexible(dgv.Rows[e.RowIndex].Cells[demandCol].Value?.ToString() ?? string.Empty, out decimal demand) && demand >= 0)
                        {
                            is_equal = (sum == demand);
                            is_shortage = (sum < demand);
                            is_exceed = (sum > demand);
                        }
                    }

                    // 步驟 2: 根據優先級，只做一次樣式決定
                    if (is_empty)
                    {
                        // 優先級 1: 空值 → 白色
                        e.CellStyle.BackColor = __defaultBack;
                        e.CellStyle.ForeColor = __defaultFore;
                    }
                    else if (is_equal)
                    {
                        // 優先級 2: 相等 → 白色
                        e.CellStyle.BackColor = __defaultBack;
                        e.CellStyle.ForeColor = __defaultFore;
                    }
                    else if (is_shortage || is_exceed)
                    {
                        // 優先級 3: 短缺或超出 → 紅色
                        e.CellStyle.BackColor = Color.Red;
                        e.CellStyle.ForeColor = Color.Black;
                    }
                    else
                    {
                        // 預設情況 → 白色
                        e.CellStyle.BackColor = __defaultBack;
                        e.CellStyle.ForeColor = __defaultFore;
                    }
                }
            }
            catch (Exception ex)
            {
                // 記錄日誌，避免在格式化事件中拋出異常導致程式崩潰
                // Debug write removed
            }
        }

        /// <summary>
        /// 數量輸入框 KeyDown 事件: 按下 Enter 執行數量更新與累加邏輯
        /// </summary>
        private void Txt備料單數量_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                var qtyText = this.txt備料單數量.Text?.Trim();
                // 診斷：記錄按下 Enter 時的游標/UseWaitCursor 狀態，方便追蹤殘留情況
                // Enter pressed snapshot diagnostic removed

                ApplyQuantityToSelectedRow(qtyText);
            }
        }

        /// <summary>
        /// 數量輸入框 KeyPress 事件: 只允許輸入數字
        /// </summary>
        private void Txt備料單數量_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 允許控制字元 (例如 Backspace, Delete)
            if (char.IsControl(e.KeyChar)) return;

            // 只允許數字,不允許其他字元(包括小數點、負號等)
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        /// <summary>
        /// 當料號輸入框取得焦點時，清除任何搜尋標示
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        private void Txt備料單料號_Enter(object sender, EventArgs e)
        {
            try
            {
                // 若目前畫面為可編輯（未鎖定），禁止切換到料號輸入，並提示使用者先按「鎖定」
                try
                {
                    if (!_isEditingLocked)
                    {
                        try { if (this.txt備料單料號 != null) this.txt備料單料號.ReadOnly = true; } catch { }
                        try { SafeShowMessage("目前畫面為可編輯狀態；請先按『鎖定』後再輸入料號。", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                        try { if (this.btn備料單Unlock != null) this.btn備料單Unlock.Focus(); } catch { }
                        return;
                    }
                }
                catch { }

                ClearRowHighlights();
            }
            catch { }
        }

        #endregion

        #region Processing

        /// <summary>
        /// 處理料號輸入: 進行不區分大小寫的模糊比對，標黃並顯示備註
        /// 處理原料輸入字串並執行相關資料更新。
        /// </summary>
        /// <param name="input">要處理的原料輸入字串。</param>
        /// <remarks>
        /// 此方法會解析 <paramref name="input"/>，並根據內容更新原料資料或觸發相關流程。
        /// </remarks>
        /// <example>
        /// <code language="csharp">
        /// // 範例：處理原料輸入
        /// ProcessMaterialInput("A12345");
        /// </code>
        /// </example>
        private void ProcessMaterialInput(string? input)
        {
            if (string.IsNullOrEmpty(input)) return;

            // 找出料號欄位索引
            int chCol = -1, custCol = -1, remarkCol = -1;
            foreach (DataGridViewColumn col in dgv備料單.Columns)
            {
                var name = (col.HeaderText ?? string.Empty).Trim();
                if (name.Contains("昶亨料號")) chCol = col.Index;
                if (name.Contains("客戶料號")) custCol = col.Index;
                if (name.Contains("備註")) remarkCol = col.Index;
            }
            if (chCol < 0 && custCol < 0) return;

            // 清除舊標色（只還原兩個可能的料號欄位）
            foreach (DataGridViewRow row in dgv備料單.Rows)
            {
                if (chCol >= 0 && row.Cells.Count > chCol)
                    row.Cells[chCol].Style.BackColor = Color.White;
                if (custCol >= 0 && row.Cells.Count > custCol)
                    row.Cells[custCol].Style.BackColor = Color.White;
            }

            // 使用 NormalizeMaterialKey 以移除特殊字元並做不分大小寫比對
            var inputTrimmed = input?.Trim() ?? string.Empty;
            var inputNorm = NormalizeMaterialKey(inputTrimmed);
            var inputUpper = inputTrimmed.ToUpperInvariant();

            // 新：比對時記錄來源欄位（昶亨或客戶）
            var matchedCells = new List<(DataGridViewRow row, int colIdx)>();
            // 1. 先做 Normalize 完全相等比對
            foreach (DataGridViewRow row in dgv備料單.Rows)
            {
                if (row.IsNewRow) continue;
                // 昶亨料號
                if (chCol >= 0 && row.Cells.Count > chCol && row.Cells[chCol].Value != null)
                {
                    var val = row.Cells[chCol].Value?.ToString() ?? string.Empty;
                    var valNorm = NormalizeMaterialKey(val);
                    if (!string.IsNullOrEmpty(valNorm) && !string.IsNullOrEmpty(inputNorm) && valNorm == inputNorm)
                    {
                        matchedCells.Add((row, chCol));
                        continue;
                    }
                }
                // 客戶料號
                if (custCol >= 0 && row.Cells.Count > custCol && row.Cells[custCol].Value != null)
                {
                    var val = row.Cells[custCol].Value?.ToString() ?? string.Empty;
                    var valNorm = NormalizeMaterialKey(val);
                    if (!string.IsNullOrEmpty(valNorm) && !string.IsNullOrEmpty(inputNorm) && valNorm == inputNorm)
                    {
                        matchedCells.Add((row, custCol));
                        continue;
                    }
                }
            }
            // 2. 若完全相等沒找到，再做前綴/包含比對
            if (matchedCells.Count == 0)
            {
                foreach (DataGridViewRow row in dgv備料單.Rows)
                {
                    if (row.IsNewRow) continue;
                    // 昶亨料號
                    if (chCol >= 0 && row.Cells.Count > chCol && row.Cells[chCol].Value != null)
                    {
                        var val = row.Cells[chCol].Value?.ToString() ?? string.Empty;
                        var valNorm = NormalizeMaterialKey(val);
                        if (!string.IsNullOrEmpty(valNorm) && !string.IsNullOrEmpty(inputNorm) && (valNorm.StartsWith(inputNorm) || valNorm.Contains(inputNorm)))
                        {
                            matchedCells.Add((row, chCol));
                            continue;
                        }
                        else if (!string.IsNullOrEmpty(val) && val.ToUpperInvariant().Contains(inputUpper))
                        {
                            matchedCells.Add((row, chCol));
                            continue;
                        }
                    }
                    // 客戶料號
                    if (custCol >= 0 && row.Cells.Count > custCol && row.Cells[custCol].Value != null)
                    {
                        var val = row.Cells[custCol].Value?.ToString() ?? string.Empty;
                        var valNorm = NormalizeMaterialKey(val);
                        if (!string.IsNullOrEmpty(valNorm) && !string.IsNullOrEmpty(inputNorm) && (valNorm.StartsWith(inputNorm) || valNorm.Contains(inputNorm)))
                        {
                            matchedCells.Add((row, custCol));
                            continue;
                        }
                        else if (!string.IsNullOrEmpty(val) && val.ToUpperInvariant().Contains(inputUpper))
                        {
                            matchedCells.Add((row, custCol));
                            continue;
                        }
                    }
                }
            }

            // 沒有比對到
            if (matchedCells.Count == 0)
            {
                try { _lastMatchedRows.Clear(); } catch { }
                try { txt備料單料號.BackColor = Color.MistyRose; } catch { }
                try { txt備料單料號.SelectAll(); } catch { }
                try { EnsureCursorRestored(); } catch { }
                MessageBox.Show("此料號不在發料清單內，請再確認", "找不到料號", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 標色所有比對到的格（只標比對來源欄位）
            foreach (var (row, colIdx) in matchedCells)
            {
                if (colIdx >= 0 && row.Cells.Count > colIdx)
                    row.Cells[colIdx].Style.BackColor = _materialHighlightColor;
            }

            // 找到多筆
            if (matchedCells.Count > 1)
            {
                try { _lastMatchedRows = matchedCells.Select(x => x.row).ToList(); } catch { }
                if (remarkCol >= 0)
                {
                    var remarkList = matchedCells
                        .Select(x => x.row.Cells[remarkCol].Value?.ToString())
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Distinct()
                        .ToList();
                    if (remarkList.Count > 0)
                    {
                        try { EnsureCursorRestored(); } catch { }
                        MessageBox.Show(string.Join("; ", remarkList), "注意事項", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                int visibleCount = matchedCells.Count(x => x.row.Visible);
                // 回填第一筆可見的料號（依來源欄位）
                try
                {
                    var firstVisible = matchedCells.FirstOrDefault(x => x.row.Visible);
                    if (firstVisible.row == null)
                        firstVisible = matchedCells[0];
                    string fill = null;
                    if (firstVisible.colIdx >= 0 && firstVisible.row.Cells.Count > firstVisible.colIdx)
                        fill = firstVisible.row.Cells[firstVisible.colIdx].Value?.ToString();
                    if (!string.IsNullOrEmpty(fill)) txt備料單料號.Text = NormalizeForTextboxFill(fill);
                    txt備料單料號.BackColor = SystemColors.Window;
                }
                catch { }
                try { EnsureCursorRestored(); } catch { }
                MessageBox.Show($"找到 {matchedCells.Count} 筆模糊匹配結果，其中 {visibleCount} 筆為可見列，已標示為黃色，請確認。", "模糊比對結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 單筆比對：自動回填料號並記錄該列
            var matched = matchedCells[0];
            try { _lastMatchedRows = new List<DataGridViewRow> { matched.row }; } catch { }
            string fillValue = null;
            if (matched.colIdx >= 0 && matched.row.Cells.Count > matched.colIdx && matched.row.Cells[matched.colIdx].Value != null)
                fillValue = matched.row.Cells[matched.colIdx].Value?.ToString() ?? string.Empty;

            if (!string.IsNullOrEmpty(fillValue))
            {
                // 去除所有空白再回填（使用 Regex 去除任意空白字元）
                try { txt備料單料號.Text = NormalizeForTextboxFill(fillValue); } catch { txt備料單料號.Text = fillValue?.Trim() ?? string.Empty; }
                try { txt備料單料號.BackColor = SystemColors.Window; } catch { }
            }

            // 顯示備註（若有）
            if (remarkCol >= 0 && remarkCol < matched.row.Cells.Count)
            {
                var remark = matched.row.Cells[remarkCol].Value?.ToString();
                if (!string.IsNullOrWhiteSpace(remark))
                {
                    try { EnsureCursorRestored(); } catch { }
                    MessageBox.Show(remark, "注意事項", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            try { this.txt備料單數量.Focus(); this.txt備料單數量.SelectAll(); } catch { }
            return;
        }

        // 共用驗證 helper：解析為大於 0 的整數
        #region Parsing Helpers
        /// <summary>
        /// 嘗試將數量字串解析為非負整數。
        /// </summary>
        /// <param name="qtyText">要解析的數量字串。</param>
        /// <param name="value">當此方法傳回時，包含解析後的非負整數。此參數在方法呼叫前未初始化。</param>
        /// <param name="errMsg">當此方法傳回時，包含錯誤訊息。此參數在方法呼叫前未初始化。</param>
        /// <returns>
        /// <see langword="true" /> 若解析成功且為非負整數；否則 <see langword="false" />。
        /// </returns>
        /// <remarks>
        /// 此方法用於安全地將輸入字串轉換為非負整數，並於失敗時提供錯誤訊息。
        /// </remarks>
        /// <example>
        /// <code language="csharp">
        /// // 範例：解析非負整數
        /// if (TryParseNonNegativeInteger("5", out var result, out var err))
        /// {
        ///     // result 會是 5，err 為空字串
        /// }
        /// </code>
        /// </example>
        private bool TryParseNonNegativeInteger(string? qtyText, out int value, out string errMsg)
        {
            value = 0; errMsg = null;
            try
            {
                if (string.IsNullOrWhiteSpace(qtyText))
                {
                    errMsg = "請輸入數量。";
                    return false;
                }
                if (!TryParseDecimalFlexible(qtyText ?? string.Empty, out decimal dec))
                {
                    errMsg = "請輸入數字。";
                    return false;
                }
                if (dec <= 0)
                {
                    errMsg = "數量需大於 0。";
                    return false;
                }
                if (decimal.Truncate(dec) != dec)
                {
                    errMsg = "數量需為整數。";
                    return false;
                }
                value = (int)dec;
                return true;
            }
            catch { errMsg = "數字解析失敗。"; return false; }
        }

        // 共用驗證 helper：解析為 decimal（接受非負或負視情境而定）
        /// <summary>
        /// 嘗試將字串型態的數量值解析為 <see cref="decimal"/> 型態。
        /// </summary>
        /// <param name="qty">要解析的數量字串。</param>
        /// <param name="value">
        /// 當此方法傳回時，包含解析後的 <see cref="decimal"/> 數值。此參數在方法呼叫前未初始化。
        /// </param>
        /// <returns>
        /// <see langword="true" /> 若解析成功；否則 <see langword="false" />。
        /// </returns>
        /// <remarks>
        /// 此方法用於安全地將輸入字串轉換為 <see cref="decimal"/>，避免例外發生並回傳解析結果。
        /// </remarks>
        /// <example>
        /// <code language="csharp">
        /// // 範例：解析數量字串
        /// if (TryParseDecimalValue("123.45", out var result))
        /// {
        ///     // result 會是 123.45
        /// }
        /// </code>
        /// </example>
        private bool TryParseDecimalValue(string? qty, out decimal value)
        {
            value = 0m;
            try
            {
                if (string.IsNullOrWhiteSpace(qty)) return false;
                if (!TryParseDecimalFlexible(qty ?? string.Empty, out decimal dec)) return false;
                value = dec;
                return true;
            }
            catch { return false; }
        }

        #endregion

        /// <summary>
        /// 將指定的數量文字套用至目前選取的資料列。
        /// </summary>
        /// <param name="qtyText">要套用的數量字串。</param>
        /// <remarks>
        /// 此方法會將 <paramref name="qtyText"/> 解析後，更新目前選取的資料列數量欄位。
        /// 若解析失敗則不會更新。
        /// </remarks>
        /// <example>
        /// <code language="csharp">
        /// // 範例：將數量 "10" 套用至選取列
        /// ApplyQuantityToSelectedRow("10");
        /// </code>
        /// </example>
        private void ApplyQuantityToSelectedRow(string? qtyText)
        {
            // 驗證數量輸入（使用共用 helper）
            if (!TryParseNonNegativeInteger(qtyText, out int qty, out string parseErr))
            {
                try { EnsureCursorRestored(); } catch { }
                try { SafeShowMessage(parseErr ?? "數量需為大於 0 的整數。", "數量錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                try { this.txt備料單數量.Focus(); this.txt備料單數量.SelectAll(); } catch { }
                return;
            }

            // 必須有比對到的列
            if (_lastMatchedRows == null || _lastMatchedRows.Count == 0)
            {
                try { EnsureCursorRestored(); } catch { }
                try { SafeShowMessage("請先輸入料號並比對成功(黃色標示)後再輸入數量。", "操作錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                try { this.txt備料單料號.Focus(); } catch { }
                return;
            }

            // 判斷檔案類型 (.xlsm 或其他)
            bool isXlsm = !string.IsNullOrEmpty(this.currentExcelPath) &&
                          Path.GetExtension(this.currentExcelPath).Equals(".xlsm", StringComparison.OrdinalIgnoreCase);

            // 取得欄位索引
            int demandCol = -1;
            int shippedCol = -1;

            if (isXlsm)
            {
                // .xlsm: 「需求數量」與「實發數量」
                demandCol = FindColumnIndexByNames(new[] { "需求數量" });
                shippedCol = FindColumnIndexByNames(new[] { "實發數量" });
            }
            else
            {
                // 其他: 「應領數量」與「發料數量」
                demandCol = FindColumnIndexByNames(new[] { "應領數量" });
                shippedCol = FindColumnIndexByNames(new[] { "發料數量" });
            }

            if (demandCol < 0 || shippedCol < 0)
            {
                string expectedCols = isXlsm ? "「需求數量」與「實發數量」" : "「應領數量」與「發料數量」";
                try { SafeShowMessage($"找不到必要欄位 {expectedCols}。", "欄位錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning); } catch { }
                return;
            }

            // 先檢查所有列是否都不會超量
            foreach (var row in _lastMatchedRows)
            {
                if (row == null) continue;
                decimal demandQty = 0;
                if (demandCol < row.Cells.Count)
                {
                    string demandStr = row.Cells[demandCol].Value?.ToString()?.Trim() ?? "0";
                    TryParseDecimalValue(demandStr, out demandQty);
                }
                decimal currentShipped = 0;
                if (shippedCol < row.Cells.Count)
                {
                    string shippedStr = row.Cells[shippedCol].Value?.ToString()?.Trim() ?? "0";
                    TryParseDecimalValue(shippedStr, out currentShipped);
                }
                decimal newShipped = currentShipped + qty;
                if (newShipped > demandQty)
                {
                    // 根據檔案類型顯示不同的訊息字句
                    string headerMsg = isXlsm ? "發出的數量已超出需求數量" : "發出的數量已超出應領數量";
                    // Debug log removed
                    try
                    {
                        try { EnsureCursorRestored(); } catch { }
                        SafeShowMessage(
                            headerMsg + "\n\n" +
                            $"目前已發: {currentShipped}\n" +
                            $"本次輸入: {qty}\n" +
                            $"累加後: {newShipped}\n" +
                            $"{(isXlsm ? "需求數量" : "應領數量")}: {demandQty}",
                            "超量警告",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                    catch { }
                    try { this.txt備料單數量.Focus(); this.txt備料單數量.SelectAll(); } catch { }
                    return;
                }
            }

            // 全部通過，開始更新與記錄

            foreach (var row in _lastMatchedRows)
            {
                if (row == null) continue;
                // 取得料號：優先使用使用者輸入/比對後回填的 txt 值（代表實際被比對到的欄位），
                // 若不存在再依序嘗試 昶亨料號 -> 客戶料號
                string materialCode = null;
                try { var txtVal = this.txt備料單料號?.Text?.Trim(); if (!string.IsNullOrEmpty(txtVal)) materialCode = txtVal; } catch { }

                int chCol = FindColumnIndexByNames(new[] { "昶亨料號" });
                int custCol = FindColumnIndexByNames(new[] { "客戶料號" });
                if (string.IsNullOrEmpty(materialCode))
                {
                    try
                    {
                        if (chCol >= 0 && chCol < row.Cells.Count)
                            materialCode = row.Cells[chCol].Value?.ToString()?.Trim();
                    }
                    catch { }
                    if (string.IsNullOrEmpty(materialCode))
                    {
                        try { if (custCol >= 0 && custCol < row.Cells.Count) materialCode = row.Cells[custCol].Value?.ToString()?.Trim(); } catch { }
                    }
                    if (materialCode == null) materialCode = string.Empty;
                }
                // 更新數量
                decimal currentShipped = 0;
                if (shippedCol < row.Cells.Count)
                {
                    string shippedStr = row.Cells[shippedCol].Value?.ToString()?.Trim() ?? "0";
                    TryParseDecimalValue(shippedStr, out currentShipped);
                }
                decimal newShipped = currentShipped + qty;

                // 嘗試以綁定資料列 (DataRowView) 更新底層 DataRow，因為 DataColumn 可能為 ReadOnly
                // 若無法使用 DataRow 更新，則退回到直接設定 DataGridView cell（並嘗試暫時解除 ReadOnly）
                try
                {
                    var dgv = this.dgv備料單;
                    bool updated = false;
                    if (row.DataBoundItem is DataRowView drv)
                    {
                        var col = (shippedCol >= 0 && dgv != null && shippedCol < dgv.Columns.Count) ? dgv.Columns[shippedCol] : null;
                        string dataColName = col?.DataPropertyName;
                        if (string.IsNullOrEmpty(dataColName)) dataColName = col?.Name;
                        if (!string.IsNullOrEmpty(dataColName) && drv.Row.Table.Columns.Contains(dataColName))
                        {
                            var dataCol = drv.Row.Table.Columns[dataColName];
                            bool origReadOnly = dataCol.ReadOnly;
                            try
                            {
                                if (origReadOnly) dataCol.ReadOnly = false;
                                drv.Row[dataColName] = newShipped;
                                updated = true;
                            }
                            finally
                            {
                                try { dataCol.ReadOnly = origReadOnly; } catch { }
                            }
                        }
                    }

                    if (!updated)
                    {
                        // fallback: 直接寫入 cell，但先嘗試解除 DataGridViewColumn.ReadOnly
                        try
                        {
                            var col = (shippedCol >= 0 && this.dgv備料單 != null && shippedCol < this.dgv備料單.Columns.Count) ? this.dgv備料單.Columns[shippedCol] : null;
                            bool origDgvReadOnly = col?.ReadOnly ?? false;
                            if (col != null && origDgvReadOnly) col.ReadOnly = false;
                            row.Cells[shippedCol].Value = newShipped;
                            if (col != null) col.ReadOnly = origDgvReadOnly;
                        }
                        catch
                        {
                            // swallow - UI will remain unchanged but avoid crash
                        }
                    }
                }
                catch (System.Data.ReadOnlyException)
                {
                    // 最後的保險：若仍拋出唯讀例外，嘗試短暫解除 DataGridViewColumn.ReadOnly
                    try
                    {
                        var col = (shippedCol >= 0 && this.dgv備料單 != null && shippedCol < this.dgv備料單.Columns.Count) ? this.dgv備料單.Columns[shippedCol] : null;
                        bool origDgvReadOnly = col?.ReadOnly ?? false;
                        if (col != null && origDgvReadOnly) col.ReadOnly = false;
                        row.Cells[shippedCol].Value = newShipped;
                        if (col != null) col.ReadOnly = origDgvReadOnly;
                    }
                    catch { }
                }
                catch { }

                // 寫入記錄
                #region 寫入記錄
                try
                {
                    _records.Add(new Dto.記錄Dto
                    {
                        刷入時間 = DateTime.Now,
                        料號 = materialCode,
                        數量 = qty,
                        操作者 = operatorName
                    });
                }
                catch { }
                #endregion
            }

            // 標示表單為已修改（需儲存），但若目前暫停 dirty 標記則略過
            try { if (!_suspendDirtyMarking) _isDirty = true; } catch { }

            // 移除黃色標示
            ClearRowHighlights();

            // 清空輸入欄位
            this.txt備料單料號.Text = "";
            this.txt備料單數量.Text = "";

            // 將焦點移回料號欄位
            this.txt備料單料號.Focus();

            // 更新快速索引與快取，然後重新標示短缺(如果實發仍小於需求，保持紅色提示)
            try
            {
                try { BuildMaterialIndex(); } catch { }
                MarkShortagesInGrid();
            }
            catch { }

            // 在完成成功流程後，額外記錄 UI/游標快照並確保游標被還原
            try
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("[APPLY_QTY_COMPLETED_SNAPSHOT]");
                try { sb.AppendLine($"  Application.UseWaitCursor={Application.UseWaitCursor}"); } catch { }
                try { sb.AppendLine($"  Cursor.Current={Cursor.Current}"); } catch { }
                try { sb.AppendLine($"  this.Cursor={this.Cursor}"); } catch { }
                // Debug log removed
            }
            catch { }

            try { EnsureCursorRestored(); } catch { }
        }

        #endregion

        #region Logging & Safe UI Helpers
        /// <summary>
        /// 以執行緒安全方式顯示訊息框。
        /// </summary>
        /// <param name="text">要顯示的訊息內容。</param>
        /// <param name="caption">訊息框標題。預設為空字串。</param>
        /// <param name="buttons">訊息框按鈕類型。預設為 <see cref="MessageBoxButtons.OK"/>。</param>
        /// <param name="icon">訊息框圖示類型。預設為 <see cref="MessageBoxIcon.None"/>。</param>
        /// <remarks>
        /// 此方法確保在多執行緒環境下安全顯示訊息框，避免 UI 執行緒衝突。
        /// </remarks>
        /// <example>
        /// <code language="csharp">
        /// // 範例：顯示警告訊息
        /// SafeShowMessage("資料儲存失敗", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        /// </code>
        /// </example>
        private void SafeShowMessage(string text, string caption = "", MessageBoxButtons buttons = MessageBoxButtons.OK, MessageBoxIcon icon = MessageBoxIcon.None)
        {
            try
            {
                try { MessageBox.Show(this, text, caption, buttons, icon); }
                catch { MessageBox.Show(text, caption, buttons, icon); }
            }
            catch
            {
                // Debug log removed
            }
        }
        #endregion

        #region Interop & STA Helpers
        /// <summary>
        /// 在 STA 執行緒上執行指定的工作，並回傳結果。
        /// Excel COM 必須在 STA 執行緒上操作，此 helper 確保整個 func 在同一 STA 執行緒上執行。
        /// </summary>
        /// <summary>
        /// 在 STA 執行緒中執行指定的委派並回傳結果。
        /// 此 helper 主要用於需要 STA 執行緒的作業（例如 COM 物件存取），
        /// 會建立臨時執行緒並等待完成後回傳值。
        /// </summary>
        /// <typeparam name="T">委派回傳型別。</typeparam>
        /// <param name="func">要在 STA 執行緒中執行的委派。</param>
        /// <returns>委派的回傳值。</returns>
        private T RunInSta<T>(Func<T> func)
        {
            // 若當前執行緒已是 STA，就直接執行
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                return func();
            }

            T result = default(T);
            Exception captured = null;
            var thread = new Thread(() =>
            {
                try { result = func(); }
                catch (Exception ex) { captured = ex; }
            });
            try
            {
                if (System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows))
                {
                    thread.SetApartmentState(ApartmentState.STA);
                }
            }
            catch { }
            thread.IsBackground = true;
            thread.Start();
            thread.Join();
            if (captured != null) throw captured;
            return result;
        }

        /// <summary>
        /// 執行與 Excel 服務相關的委派並安全地攔截例外；若需要，於 STA 執行緒中執行。
        /// 用途：包裝 _excelService 的呼叫，避免 UI 執行緒被阻塞與非 STA 造成的異常。
        /// </summary>
        private T _excel_service_safe<T>(Func<T> func)
        {
            try
            {
                // 多數 Excel/Interop 呼叫需在 STA 執行緒上
                return RunInSta(func);
            }
            catch
            {
                return default(T);
            }
        }

        /// <summary>
        /// 取得 IExcelService 實例：
        /// 1) 若已注入強型別 _typedExcelService，則回傳之
        /// 2) 否則若原本有傳入 dynamic _excelService，則以 DynamicExcelAdapter 包裝後回傳
        /// 3) 否則回傳預設的 NpoiExcelService
        /// 此方法保證不會回傳 null，且為輕量的工廠/路由函式，目的是在不破壞現有 logic 下逐步導入介面。
        /// </summary>
        private IExcelService GetExcelService()
        {
            try
            {
                if (_typedExcelService != null) return _typedExcelService;
                if (_excelService != null)
                {
                    try { return new DynamicExcelAdapter(_excelService); } catch { }
                }
            }
            catch { }
            // fallback to conservative implementation
            return new NpoiExcelService();
        }

        /// <summary>
        /// 檢查檔案是否被其他程序鎖定（嘗試以排他方式開啟）。
        /// 若檔案不存在則回傳 false。
        /// </summary>
        /// <summary>
        /// 檢查檔案是否被鎖定（改進版：優先使用 COM 策略）
        /// 策略 A: 嘗試用 COM 以讀寫模式開啟（最可靠）
        /// 策略 B: 檢查 Excel 程序是否持有檔案
        /// 策略 C: 嘗試以獨佔模式開啟檔案
        /// </summary>
        private bool IsFileLocked(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return false;

            //var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? Directory.GetCurrentDirectory(), "excel_write_error.log");
            var logMsg = new System.Text.StringBuilder();

            // 【策略】嘗試以獨佔模式打開檔案（最可靠）
            const int maxRetries = 5;
            const int delayMs = 50;

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                FileStream? fs = null;
                try
                {
                    // 嘗試以獨佔讀寫模式打開
                    fs = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    fs.Close();
                    fs.Dispose();

                    // 成功 = 檔案自由
                    logMsg.Clear();
                    logMsg.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [IsFileLocked] SUCCESS on attempt {attempt + 1}: File is FREE");
                    // try { WriteDebugFile(logPath, logMsg.ToString()); } catch { }
                    return false;
                }
                catch (IOException ioEx)
                {
                    // IOException = 檔案被鎖定
                    logMsg.Clear();
                    logMsg.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [IsFileLocked] Attempt {attempt + 1}/{maxRetries}: IOException - {ioEx.Message}");
                    //try { WriteDebugFile(logPath, logMsg.ToString()); } catch { }

                    if (attempt < maxRetries - 1)
                    {
                        System.Threading.Thread.Sleep(delayMs);
                    }
                    else
                    {
                        logMsg.Clear();
                        logMsg.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [IsFileLocked] FINAL: All {maxRetries} attempts failed, file is LOCKED");
                        //try { WriteDebugFile(logPath, logMsg.ToString()); } catch { }
                        return true;
                    }
                }
                catch (UnauthorizedAccessException uaEx)
                {
                    // 權限不足
                    logMsg.Clear();
                    logMsg.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [IsFileLocked] UnauthorizedAccessException: {uaEx.Message}");
                    //try { WriteDebugFile(logPath, logMsg.ToString()); } catch { }
                    return true;
                }
                catch (Exception ex)
                {
                    // 其他異常
                    logMsg.Clear();
                    logMsg.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [IsFileLocked] Unexpected exception {ex.GetType().Name}: {ex.Message}");
                    // try { File.AppendAllText(logPath, logMsg.ToString()); } catch { }

                    if (attempt < maxRetries - 1)
                    {
                        System.Threading.Thread.Sleep(delayMs);
                    }
                    else
                    {
                        return true;
                    }
                }
                finally
                {
                    try { fs?.Dispose(); } catch { }
                }
            }

            // 預設：所有重試都失敗
            logMsg.Clear();
            logMsg.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [IsFileLocked] FINAL DEFAULT: Returning TRUE (locked)");
            //try { WriteDebugFile(logPath, logMsg.ToString()); } catch { }
            return true;
        }
        #endregion

        #region Column & Material Index Helpers
        /// <summary>
        /// 依欄位候選名稱尋找對應的 DataGridView 欄位索引。
        /// 先做精確比對（忽略前後空白與大小寫），再做正規化包含式比對。
        /// </summary>
        private int FindColumnIndexByNames(IEnumerable<string> names)
        {
            if (this.dgv備料單?.Columns == null || this.dgv備料單.Columns.Count == 0) return -1;

            string San(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return string.Empty;

                // 優先使用集中式的正規化函式以維持行為一致性
                try
                {
                    var global = SanitizeHeaderForMatch(s);
                    if (!string.IsNullOrEmpty(global)) return global;
                }
                catch { }

                // fallback: 保留原本 local 行為（包含 CJK 範圍），以避免因集中函式尚未覆蓋某些語系行為而破壞邏輯
                s = s.Replace('\u0020', ' ').Replace('\u3000', ' ').Trim();
                var sb = new System.Text.StringBuilder(s.Length);
                foreach (var ch in s)
                {
                    if (char.IsLetterOrDigit(ch) || (ch >= 0x4E00 && ch <= 0x9FFF)) sb.Append(char.ToLowerInvariant(ch));
                }
                return sb.ToString();
            }

            // 1) 精確比對（不分大小寫、忽略多餘空白）
            var exactTargets = names
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Select(n => n.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < this.dgv備料單.Columns.Count; i++)
            {
                var col = this.dgv備料單.Columns[i];
                var candidates = new[] { col.HeaderText, col.Name, col.DataPropertyName };
                if (candidates.Any(c => !string.IsNullOrWhiteSpace(c) && exactTargets.Contains(c.Trim())))
                {
                    return i;
                }
            }

            // 2) 正規化包含式比對（移除非英數與中日韓字元、轉小寫）
            var normTargets = names.Select(San).Where(t => !string.IsNullOrEmpty(t)).ToList();
            for (int i = 0; i < this.dgv備料單.Columns.Count; i++)
            {
                var col = this.dgv備料單.Columns[i];
                var candidates = new[] { col.HeaderText, col.Name, col.DataPropertyName };
                var normColNames = candidates.Where(x => !string.IsNullOrWhiteSpace(x)).Select(San).ToArray();
                foreach (var nt in normTargets)
                {
                    if (normColNames.Any(nc => !string.IsNullOrEmpty(nc) && (nc.Contains(nt) || nt.Contains(nc))))
                        return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// 建立備料單的索引，將原始資料整理成可快速查詢的結構。
        /// </summary>
        /// <remarks>
        /// 此方法會根據目前的備料單資料，重新整理索引以提升查詢效率。
        /// 若資料量龐大，建議在非 UI 執行緒執行以避免介面卡頓。
        /// </remarks>
        /// <example>
        /// <code language="csharp">
        /// // 範例：呼叫方法以重建索引
        /// BuildMaterialIndex();
        /// </code>
        /// </example>
        private void BuildMaterialIndex()
        {
            try
            {
                // lazy init service
                try
                {
                    if (_materialIndexService == null && this.dgv備料單 != null)
                        _materialIndexService = new Automatic_Storage.Services.MaterialIndexService(this.dgv備料單);
                }
                catch { }

                if (_materialIndexService != null)
                {
                    try { _materialIndexService.BuildIndex(); } catch { }

                    // sync service results into existing backing fields to preserve current callers
                    try { _materialIndex.Clear(); } catch { }
                    try { _materialShippedSums.Clear(); } catch { }
                    try
                    {
                        foreach (var kv in _materialIndexService.MaterialIndex)
                        {
                            try { _materialIndex[kv.Key] = kv.Value.ToList(); } catch { }
                        }
                        foreach (var kv in _materialIndexService.MaterialShippedSums)
                        {
                            try { _materialShippedSums[kv.Key] = kv.Value; } catch { }
                        }
                    }
                    catch { }
                    return;
                }

                // fallback: original in-place implementation (kept for safety)
                try
                {
                    _materialIndex.Clear();
                    _materialShippedSums.Clear();
                    if (this.dgv備料單?.Rows == null) return;
                    int materialCol = FindColumnIndexByNames(new[] { "昶亨料號", "客戶料號" });
                    if (materialCol < 0) return;

                    foreach (DataGridViewRow row in this.dgv備料單.Rows)
                    {
                        if (row.IsNewRow) continue;
                        try
                        {
                            var raw = row.Cells[materialCol].Value;
                            var val = raw?.ToString()?.Trim();
                            if (string.IsNullOrEmpty(val)) continue;

                            var key = val;
                            if (!_materialIndex.TryGetValue(key, out List<DataGridViewRow> list))
                            {
                                list = new List<DataGridViewRow>(4);
                                _materialIndex[key] = list;
                            }
                            if (list.Count == 0 || !object.ReferenceEquals(list[list.Count - 1], row)) list.Add(row);
                        }
                        catch { }
                    }

                    try
                    {
                        int shippedCol = FindColumnIndexByNames(new[] { "實發數量", "發料數量" });
                        if (shippedCol >= 0)
                        {
                            foreach (DataGridViewRow row in this.dgv備料單.Rows)
                            {
                                try
                                {
                                    if (row == null || row.IsNewRow) continue;
                                    var rawMat = row.Cells[materialCol].Value?.ToString()?.Trim();
                                    if (string.IsNullOrEmpty(rawMat)) continue;
                                    var key = NormalizeMaterialKey(rawMat);
                                    if (string.IsNullOrEmpty(key)) continue;
                                    decimal v = 0m;
                                    var sv = row.Cells.Count > shippedCol ? row.Cells[shippedCol].Value?.ToString() ?? string.Empty : string.Empty;
                                    if (TryParseDecimalValue(sv, out decimal parsed)) v = parsed;
                                    if (_materialShippedSums.TryGetValue(key, out decimal exist)) _materialShippedSums[key] = exist + v;
                                    else _materialShippedSums[key] = v;
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }
                catch { }
            }
            catch { }
        }

        /// <summary>
        /// 清除備料單索引，將目前的索引資料全部移除。
        /// </summary>
        /// <remarks>
        /// 此方法會將所有已建立的備料單索引資料清空，適用於重新整理或重建索引前的初始化步驟。
        /// 執行後索引資料將不可查詢，需重新建立。
        /// </remarks>
        /// <example>
        /// <code language="csharp">
        /// // 範例：呼叫方法以清空索引
        /// ClearMaterialIndex();
        /// </code>
        /// </example>
        private void ClearMaterialIndex()
        {
            try { _materialIndex.Clear(); } catch { }
            try { _materialShippedSums.Clear(); } catch { }
            try { _materialIndexService?.Clear(); } catch { }
        }

        #endregion

        #region Form Lifecycle & Visibility
        /// <summary>
        /// 覆寫 <see cref="Form.OnFormClosing"/>，在表單關閉時執行必要的清理與取消邏輯（例如取消背景工作或等待佇列 flush）。
        /// 此方法嘗試釋放與表單相關的資源（計時器、Tooltip、事件註冊、背景 cancellation token 等），並還原 UI 狀態以避免影響其他視窗。
        /// </summary>
        /// <param name="e">包含關閉表單事件的資料。</param>
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            try
            {
                // 關閉(X)時才重置目前快取資料
                try { _currentUiTable = null; } catch { }
                // 取消任何背景工作
                try { _cts?.Cancel(); } catch { }
                try { _excelQueueCts?.Cancel(); } catch { }

                // 清理 resize timer 與事件註冊
                try { if (_resizeTimer != null) { _resizeTimer.Stop(); _resizeTimer.Tick -= ResizeTimer_Tick; _resizeTimer.Dispose(); _resizeTimer = null; } } catch { }
                try { this.Resize -= Form_Resize; } catch { }
                try { if (this.dgv備料單 != null) { this.dgv備料單.SizeChanged -= Dgv_SizeChanged; } } catch { }
                try { if (this.dgv備料單 != null) { this.dgv備料單.CellMouseMove -= Dgv備料單_CellMouseMove; this.dgv備料單.CellMouseLeave -= Dgv備料單_CellMouseLeave; this.dgv備料單.Scroll -= Dgv備料單_Scroll; } } catch { }
                try { if (_cellToolTip != null) { _cellToolTip.Dispose(); _cellToolTip = null; } } catch { }
            }
            catch { }
            // 額外保證性清理：關閉時一定要移除任何 overlay 並還原按鈕狀態，避免主表單被鎖定
            try
            {
                try { HideOperationOverlay(); } catch { }
            }
            catch { }
            try
            {
                try { RestoreAllButtons(); } catch { }
            }
            catch { }
            // 確保所有操作旗標都已清除
            try { _isImporting = false; _isExporting = false; _isSaving = false; } catch { }
            try { _keepImportButtonDisabledUntilClose = false; } catch { }
            // 若有 Owner（例如主畫面），在關閉時確保將其顯示回來
            try
            {
                if (this.Owner is Form ownerForm)
                {
                    try { ownerForm.Show(); ownerForm.BringToFront(); ownerForm.Activate(); } catch { }
                }
            }
            catch { }
        }

        /// <summary>
        /// 覆寫 <see cref="Form.OnShown"/>，當表單首次顯示時執行一次性的初始化或 UI 調整邏輯，例如設定焦點或觸發延後的載入動作。
        /// </summary>
        /// <param name="e">事件引數。</param>
        protected override void OnShown(EventArgs e)
        {
            /* Lines 2976-2999 omitted */
            base.OnShown(e);
            try { this.WindowState = FormWindowState.Maximized; this.BringToFront(); this.Activate(); } catch { }
            try
            {
                if (this.dgv備料單 != null && this.dgv備料單.DataSource == null && _currentUiTable != null)
                {
                    this.dgv備料單.DataSource = _currentUiTable;
                    this.dgv備料單.Refresh();
                    try { HideColumnsByHeaders(this.dgv備料單); } catch { }
                    try { BuildMaterialIndex(); } catch { }
                }
            }
            catch { }

            // 顯示時把輸入焦點放到料號輸入框（若存在）
            try
            {
                SafeBeginInvoke(this, new Action(() =>
                {
                    try { if (this.txt備料單料號 != null && !this.txt備料單料號.IsDisposed) { this.txt備料單料號.Focus(); this.txt備料單料號.SelectAll(); } } catch { }
                }));
            }
            catch { }

            // 一次性測試：寫入 dgv_style_debug.log（僅在 EnableDgvStyleLog 開啟時）
            try
            {
                // OnShown: previously had optional test write for dgv_style_debug.log; removed.
            }
            catch { }
            try { UpdateMainButtonsEnabled(); } catch { }
        }

        /// <summary>
        /// 覆寫 <see cref="Control.OnVisibleChanged"/>，當表單可見性改變時觸發（例如從最小化還原），可用於恢復或暫停 UI 行為與背景工作。
        /// </summary>
        /// <param name="e">事件引數。</param>
        protected override void OnVisibleChanged(EventArgs e)
        {
            /* Lines 3003-3019 omitted */
            base.OnVisibleChanged(e);
            try
            {
                if (this.Visible)
                {
                    try { this.WindowState = FormWindowState.Maximized; } catch { }
                    if (this.dgv備料單 != null && this.dgv備料單.DataSource == null && _currentUiTable != null)
                    {
                        try { this.dgv備料單.DataSource = _currentUiTable; this.dgv備料單.Refresh(); } catch { }
                        try { HideColumnsByHeaders(this.dgv備料單); } catch { }
                        try { BuildMaterialIndex(); } catch { }
                    }
                    try { SafeBeginInvoke(this, () => { try { if (this.txt備料單料號 != null && !this.txt備料單料號.IsDisposed) { this.txt備料單料號.Focus(); this.txt備料單料號.SelectAll(); } } catch { } }); } catch { }
                }
            }
            catch { }
        }

        #endregion

        #region Layout & Resize Helpers
        /// <summary>
        /// 處理定時觸發的事件，用於執行 UI 重新調整大小的邏輯。
        /// </summary>
        /// <param name="sender">引發此事件的物件。</param>
        /// <param name="e">包含事件相關資訊的 <see cref="EventArgs"/>。</param>
        private void ResizeTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                // 若目前視窗是最大化狀態，避免任何自動縮放計算以免把視窗縮小
                try { if (this.WindowState == FormWindowState.Maximized) { _resizeTimer?.Stop(); return; } } catch { }
                // 若視窗已最小化或不可見，跳過調整邏輯，避免在 Minimized 時改變位置/大小導致還原問題
                try
                {
                    if (this == null || this.IsDisposed) return;
                }
                catch { }

                try
                {
                    if (!this.Visible || this.WindowState == FormWindowState.Minimized)
                    {
                        _resizeTimer?.Stop();
                        return;
                    }
                }
                catch { }

                _resizeTimer?.Stop();
                // 確保 DataGridView 以 Dock.Fill 呈現，避免未延展至右側的空白
                try { if (this.dgv備料單 != null && this.dgv備料單.Dock != DockStyle.Fill) this.dgv備料單.Dock = DockStyle.Fill; } catch { }
                AutoSizeColumnsFillNoHorizontalScroll(this.dgv備料單);
                // 觸發父容器重新排版並刷新，以處理 Designer 中可能的容器 padding 或停駐狀態
                try
                {
                    var parent = this.dgv備料單?.Parent;
                    parent?.PerformLayout();
                    this.PerformLayout();
                    this.dgv備料單?.Invalidate();
                    this.dgv備料單?.Refresh();
                }
                catch { }
                try
                {
                    if (this.dgv備料單 != null)
                    {
                        int totalCols = 0; foreach (DataGridViewColumn c in this.dgv備料單.Columns) totalCols += c.Width;
                        if (totalCols > this.dgv備料單.ClientSize.Width) this.dgv備料單.ScrollBars = ScrollBars.Both;
                    }
                }
                catch { }
                try { AdjustFormSizeToDataGrid(); } catch { }
            }
            catch { }
        }

        // Form Resize handler - 啟動去抖
        /// <summary>
        /// 處理表單大小調整事件，根據視窗尺寸變化調整 UI 佈局。
        /// </summary>
        /// <param name="sender">觸發事件的物件。</param>
        /// <param name="e">包含事件資料的 <see cref="EventArgs"/> 物件。</param>
        private void Form_Resize(object sender, EventArgs e)
        {
            try
            {
                // 若切換到最大化，停止去抖並跳出，不再做自動調整
                try
                {
                    if (this.WindowState == FormWindowState.Maximized)
                    {
                        if (_resizeTimer != null) _resizeTimer.Stop();
                        return;
                    }
                }
                catch { }
                // 若最小化則不要啟動去抖計時器，避免在最小化/還原期間造成不必要的尺寸計算
                try
                {
                    if (this.WindowState == FormWindowState.Minimized)
                    {
                        // 記錄目前為最小化
                        _wasMinimized = true;
                        if (_resizeTimer != null) _resizeTimer.Stop();
                        return;
                    }

                    // 若先前為最小化但現在已還原，做還原處理
                    if (_wasMinimized && this.WindowState != FormWindowState.Minimized)
                    {
                        _wasMinimized = false;
                        try { HideOperationOverlay(); } catch { }
                        try { RestoreAllButtons(); } catch { }
                        try { if (!this.Visible) this.Show(); } catch { }
                        try { this.BringToFront(); this.Activate(); } catch { }
                    }
                }
                catch { }

                if (_resizeTimer != null) { _resizeTimer.Stop(); _resizeTimer.Start(); }
            }
            catch { }
        }

        /// <summary>
        /// DataGridView 滑鼠移動事件處理器。
        /// 當滑鼠游標移動到儲存格上時，若儲存格內容被裁切，則顯示完整內容的 ToolTip。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 DataGridView。</param>
        /// <param name="e">包含儲存格座標的事件參數。</param>
        private void Dgv備料單_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                if (_cellToolTip == null) return;
                if (e.RowIndex < 0 || e.ColumnIndex < 0) { _cellToolTip.Hide(this.dgv備料單); return; }
                var dgv = this.dgv備料單;
                var cell = dgv.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (cell == null) { _cellToolTip.Hide(dgv); return; }
                var text = cell.Value?.ToString() ?? string.Empty;
                if (string.IsNullOrEmpty(text)) { _cellToolTip.Hide(dgv); return; }

                var cellRect = dgv.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, false);
                using (var g = dgv.CreateGraphics())
                {
                    var ft = cell.InheritedStyle.Font ?? dgv.Font;
                    var size = g.MeasureString(text, ft);
                    if (size.Width > cellRect.Width - 4)
                    {
                        // 顯示於儲存格下方
                        var localPt = new Point(cellRect.Left + 2, cellRect.Bottom + 2);
                        _cellToolTip.Show(text, dgv, localPt, _cellToolTip.AutoPopDelay);
                    }
                    else
                    {
                        _cellToolTip.Hide(dgv);
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// DataGridView 滑鼠離開儲存格事件處理器。
        /// 當滑鼠游標離開儲存格時，隱藏顯示中的 ToolTip。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 DataGridView。</param>
        /// <param name="e">包含儲存格座標的事件參數。</param>
        private void Dgv備料單_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            try { _cellToolTip?.Hide(this.dgv備料單); } catch { }
        }

        /// <summary>
        /// DataGridView 捲動事件處理器。
        /// 當 DataGridView 捲動時，自動隱藏目前顯示的 ToolTip，避免內容殘留。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 DataGridView。</param>
        /// <param name="e">包含捲動資訊的事件參數。</param>
        private void Dgv備料單_Scroll(object sender, ScrollEventArgs e)
        {
            try { _cellToolTip?.Hide(this.dgv備料單); } catch { }
        }

        /// <summary>
        /// 根據 DataGridView 的內容計算並調整表單大小。
        /// 行為：計算欄位總寬與列高度，並將表單寬高調整到能完整顯示內容（含少量 margin），
        /// 但不超過工作區 (Screen.WorkingArea) 的大小，且保有最小寬/高保護值。
        /// </summary>
        private void AdjustFormSizeToDataGrid()
        {
            try
            {
                // 最大化狀態不做任何自動調整，避免被縮小
                try { if (this.WindowState == FormWindowState.Maximized) return; } catch { }
                if (this.dgv備料單 == null || this.dgv備料單.Columns.Count == 0) return;

                // 計算所需寬度：欄位總寬 + 行號/邊框 + 估算垂直捲軸寬
                int colsTotal = 0; foreach (DataGridViewColumn c in this.dgv備料單.Columns) colsTotal += c.Width;
                int vscrollWidth = SystemInformation.VerticalScrollBarWidth;
                int desiredDgvWidth = colsTotal + this.dgv備料單.RowHeadersWidth + 8; // small padding

                // 若目前內容高度可以顯示所有列，則不需要垂直捲軸寬度，否則加入寬度
                bool needVScroll = this.dgv備料單.Rows.Count > this.dgv備料單.DisplayedRowCount(true);
                if (needVScroll) desiredDgvWidth += vscrollWidth;

                // 計算所需高度：欄頭高度 + 所有列高度 + 小 padding
                int rowsTotalHeight = this.dgv備料單.ColumnHeadersHeight;
                foreach (DataGridViewRow r in this.dgv備料單.Rows) rowsTotalHeight += r.Height;
                int desiredDgvHeight = rowsTotalHeight + 8;

                // 將 dgv 的需求轉換為 form 的目標大小（考慮到其他控制項與邊界）
                // 估算額外寬度/高度：Form.ClientSize - dgv.ClientSize
                int extraWidth = Math.Max(0, this.ClientSize.Width - this.dgv備料單.ClientSize.Width);
                int extraHeight = Math.Max(0, this.ClientSize.Height - this.dgv備料單.ClientSize.Height);

                int targetWidth = desiredDgvWidth + extraWidth;
                int targetHeight = desiredDgvHeight + extraHeight;

                // clamp to screen working area
                var wa = Screen.FromControl(this).WorkingArea;
                int minW = Math.Max(600, this.MinimumSize.Width > 0 ? this.MinimumSize.Width : 600);
                int minH = Math.Max(400, this.MinimumSize.Height > 0 ? this.MinimumSize.Height : 400);
                int finalW = Math.Min(Math.Max(targetWidth, minW), wa.Width);
                int finalH = Math.Min(Math.Max(targetHeight, minH), wa.Height);

                // 只在變更顯著時調整表單尺寸，避免頻繁搖動
                if (Math.Abs(this.Width - finalW) > 8 || Math.Abs(this.Height - finalH) > 8)
                {
                    // 將表單置中於原先左上角附近（避免跳到螢幕邊界外）
                    int newX = this.Left; int newY = this.Top;
                    // 若超出工作區右邊，調整 Left
                    if (newX + finalW > wa.Right) newX = Math.Max(wa.Left, wa.Right - finalW);
                    if (newY + finalH > wa.Bottom) newY = Math.Max(wa.Top, wa.Bottom - finalH);

                    this.SuspendLayout();
                    try
                    {
                        this.SetBounds(newX, newY, finalW, finalH);
                        this.PerformLayout();
                        this.Refresh();
                    }
                    finally { try { this.ResumeLayout(); } catch { } }
                }
            }
            catch { }
        }

        /// <summary>
        /// 將 _preservedRedKeys 中標記為短缺的料號，於 DataGridView 中對應的「實發數量」或「發料數量」欄位恢復紅色高亮。
        /// 此方法會根據「昶亨料號」與「客戶料號」欄位比對，僅針對已標記的料號進行顏色還原。
        /// 常用於存檔、匯出或重新綁定資料後，確保短缺提示不會遺失。
        /// </summary>
        private void RestorePreservedRedHighlights()
        {
            try
            {
                var dgv = this.dgv備料單;
                if (dgv == null || dgv.Rows == null || dgv.Rows.Count == 0) return;

                int chCol = FindColumnIndexByNames(new[] { "昶亨料號" });
                int custCol = FindColumnIndexByNames(new[] { "客戶料號" });
                int shippedCol = FindColumnIndexByNames(new[] { "實發數量", "發料數量" });
                if (shippedCol < 0) return;

                foreach (DataGridViewRow row in dgv.Rows)
                {
                    try
                    {
                        if (row == null || row.IsNewRow) continue;
                        string materialCode = "";
                        if (chCol >= 0 && chCol < row.Cells.Count)
                            materialCode = row.Cells[chCol].Value?.ToString()?.Trim() ?? "";
                        if (string.IsNullOrEmpty(materialCode) && custCol >= 0 && custCol < row.Cells.Count)
                            materialCode = row.Cells[custCol].Value?.ToString()?.Trim() ?? "";

                        if (string.IsNullOrEmpty(materialCode)) continue;
                        var key = NormalizeMaterialKey(materialCode);
                        if (!string.IsNullOrEmpty(key) && _preservedRedKeys != null && _preservedRedKeys.Contains(key))
                        {
                            try { row.Cells[shippedCol].Style.BackColor = Color.Red; } catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }

        }

        #endregion

        /// <summary>
        /// 驗證 DataTable 是否包含必要欄位，並建立欄位名稱對應的索引對應表。
        /// </summary>
        /// <param name="dt">要驗證的資料表。</param>
        /// <param name="mapping">
        /// [out] 回傳欄位對應表，key 為 internal name（如 "Material", "Demand", "Shipped", "Remark"），value 為 DataTable 欄位索引。
        /// </param>
        /// <param name="errMsg">[out] 若驗證失敗，回傳錯誤訊息。</param>
        /// <returns>若驗證成功且 mapping 完整，回傳 true；否則回傳 false 並於 errMsg 說明原因。</returns>
        private bool ValidateAndMapColumns(DataTable dt, out Dictionary<string, int> mapping, out string errMsg)
        {
            // 回傳 mapping 的鍵值為我們統一使用的 internal name
            mapping = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            errMsg = null;
            if (dt == null) { errMsg = "資料表為空。"; return false; }

            // 欄位候選群組（越前面的名稱優先）
            var candidates = new Dictionary<string, string[]>
            {
                // 料號常見別名
                { "Material", new[] { "昶亨料號", "客戶料號"  } },
                // 需求/應領常見別名
                { "Demand",   new[] { "需求數量", "應領數量" } },
                // 實發/發料常見別名（使用者回報最常缺這一欄）
                { "Shipped",  new[] { "實發數量", "發料數量" } },
                { "Remark",   new[] { "備註" } }
            };

            // 建立欄位快取：去除前後空白並且保留原始名稱及簡化判斷
            var colNameMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                var raw = dt.Columns[i].ColumnName ?? string.Empty;
                var key = raw.Trim();
                if (colNameMap.ContainsKey(key))
                {
                    // 若重複欄位名稱，加入索引後綴以保證唯一
                    int suffix = 1; var uniq = key + "_" + suffix;
                    while (colNameMap.ContainsKey(uniq)) { suffix++; uniq = key + "_" + suffix; }
                    colNameMap[uniq] = i;
                }
                else colNameMap[key] = i;
            }

            // helper: 找到第一個 match（加入負向關鍵字以避免把「需求/應領」誤認為實發/發料）
            int FindFirstMatch(string[] names)
            {
                // 1) exact match of candidate to column name
                foreach (var n in names)
                {
                    var key = colNameMap.Keys.FirstOrDefault(k => string.Equals(k, n, StringComparison.OrdinalIgnoreCase));
                    if (key != null) return colNameMap[key];
                }

                // 2) column name contains candidate token (e.g., header "料號(料號)" should match "料號")
                foreach (var n in names)
                {
                    var key = colNameMap.Keys.FirstOrDefault(k => k.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (key != null) return colNameMap[key];
                }

                // 3) candidate token contains column name (handles short column headers like "料" vs "料號")
                foreach (var n in names)
                {
                    var key = colNameMap.Keys.FirstOrDefault(k => n.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (key != null) return colNameMap[key];
                }

                // 4) fuzzy by removing non-alphanumeric characters and comparing
                // Reuse central normalizer to keep behavior consistent
                var normalizedCandidates = names.Select(x => SanitizeHeaderForMatch(x)).Where(x => !string.IsNullOrEmpty(x)).ToArray();
                foreach (var ck in colNameMap.Keys)
                {
                    var nk = SanitizeHeaderForMatch(ck);
                    // 修正：當中文欄名去除非英數字會變成空字串時，先略過，避免任何 candidate 都被誤判符合
                    if (string.IsNullOrEmpty(nk)) continue;
                    if (normalizedCandidates.Any(cn => !string.IsNullOrEmpty(cn) && (nk.Contains(cn) || cn.Contains(nk)))) return colNameMap[ck];
                }

                return -1;
            }

            // 檢查每個必要欄位（Material, Demand, Shipped）
            var requiredInternal = new[] { "Material", "Demand", "Shipped" };
            foreach (var internalName in requiredInternal)
            {
                var idx = FindFirstMatch(candidates[internalName]);
                if (idx < 0)
                {
                    // 若找不到，提供更友善的錯誤訊息，列出目前檔案的欄位名稱以利排查
                    var available = string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
                    errMsg = $"找不到必要欄位：{internalName}（嘗試候選名稱：{string.Join(",", candidates[internalName])}）。\n目前偵測到的欄位：{available}\n請確認檔案的標題列是否為第 3 列或欄位名稱是否正確。";
                    return false;
                }
                mapping[internalName] = idx;
            }

            // 針對 Shipped 再次防呆：若誤抓到「需求/應領」欄，改以更嚴格規則重新尋找
            if (mapping.TryGetValue("Shipped", out var shippedIdx))
            {
                try
                {
                    var shippedName = dt.Columns[shippedIdx].ColumnName ?? string.Empty;
                    if (shippedName.Contains("需求") || shippedName.Contains("應領"))
                    {
                        // 僅接受包含「實發」或「發料」等正向詞的欄位
                        var strictNames = new[] { "實發數量", "發料數量" };
                        int strict = FindFirstMatch(strictNames);
                        if (strict >= 0) mapping["Shipped"] = strict;
                    }
                }
                catch { }
            }

            // Remark 非必要但如果存在則紀錄 index
            var remarkIdx = FindFirstMatch(candidates["Remark"]);
            if (remarkIdx >= 0) mapping["Remark"] = remarkIdx;

            return true;
        }

        /// <summary>
        /// 覆寫 <see cref="Form.ProcessCmdKey(ref Message, Keys)"/> 以攔截特定系統鍵（例如貼上），避免在料號或數量輸入時意外貼上造成不當資料。
        /// </summary>
        /// <param name="msg">Windows 訊息封裝結構，傳入時使用 ref。</param>
        /// <param name="keyData">按鍵資料，包含修飾鍵（Ctrl/Shift 等）與實際按鍵。</param>
        /// <returns>如果已處理該按鍵（例如阻擋貼上）則回傳 <see langword="true"/>，否則回傳基底實作的結果。</returns>
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            // 阻擋貼上技巧：Ctrl+V, Shift+Insert
            try
            {
                // 只在料號或數量輸入框時阻擋貼上，避免全域攔截造成其他快速鍵失效
                var focused = this.ActiveControl;
                bool isTargetTextbox = false;
                try { isTargetTextbox = (focused == this.txt備料單料號) || (focused == this.txt備料單數量); } catch { isTargetTextbox = false; }

                if (isTargetTextbox)
                {
                    if ((keyData & Keys.Control) == Keys.Control && (keyData & Keys.KeyCode) == Keys.V) return true;
                    if ((keyData & Keys.Shift) == Keys.Shift && (keyData & Keys.KeyCode) == Keys.Insert) return true;
                }
            }
            catch { }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        /// <summary>
        /// 從資料庫取得 Excel 密碼，並快取一段時間以減少查詢次數。
        /// 若查詢失敗則回傳預設密碼。
        /// </summary>
        /// <param name="cacheMinutes">密碼快取有效分鐘數，預設為 10 分鐘。</param>
        /// <returns>Excel 密碼字串。</returns>
        private string GetExcelPassword(int cacheMinutes = 10)
        {
            try
            {
                return _excelPasswordProvider.GetPassword();
            }
            catch
            {
                // 若 provider 發生例外，回退到表單原本的快取或預設值以避免破壞使用者流程
                lock (_pwdLock)
                {
                    if (!string.IsNullOrEmpty(_cachedExcelPassword)) return _cachedExcelPassword;
                    _cachedExcelPassword = defaultPwd;
                    _cachedExcelPasswordAt = DateTime.UtcNow;
                    return _cachedExcelPassword;
                }
            }
        }

        /// <summary>
        /// 將 DataTable 中的「實發/發料」數量批次寫回 Excel「總表」工作表，並記錄異動到「記錄」工作表。
        /// - 會自動偵測標頭列與欄位位置，僅覆寫非公式儲存格。
        /// - 若目標儲存格原本為公式，則保留公式不覆寫。
        /// - 支援自動解除/還原保護，並於寫入後重新保護工作表。
        /// - 會將本次異動量（delta）記錄到「記錄」工作表，避免重複寫入。
        /// - 所有 Excel COM 操作皆於 STA 執行緒執行，確保相容性。
        /// </summary>
        /// <param name="excelPath">目標 Excel 檔案路徑。</param>
        /// <param name="dt">來源資料表，需包含「料號」與「實發/發料」欄位。</param>
        /// <param name="columnMapping">欄位對應表，key 為 "Material"、"Shipped"。</param>
        /// <param name="excelPassword">Excel 保護密碼，若無則傳入 null。</param>
        /// <param name="errMsg">[out] 若失敗則回傳錯誤訊息。</param>
        /// <returns>成功寫入則回傳 true，否則 false。</returns>
        private bool WriteBackShippedQuantitiesToExcelBatch(string excelPath, DataTable dt, Dictionary<string, int> columnMapping, string excelPassword, out string errMsg)
        {
            errMsg = null;

            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath)) { errMsg = "Excel 檔案不存在。"; return false; }

            try
            {
                // 現在整個 Excel COM 流程交給 RunInSta 執行，並回傳 (ok, err)
                // timing removed: overall WriteBack stopwatch
                // var sw = System.Diagnostics.Stopwatch.StartNew();
                // 記錄 MAIN_SHEET 階段的起始 timestamp，方便後續 cross-check
                // debug file path declarations removed
                var result = RunInSta(() =>
                {
                    string localErr = null;
                    Excel.Application xlApp = null;
                    Excel.Workbook wb = null;
                    Excel.Range? used = null;
                    // 用於暫存原本的 Application 設定，操作完畢會還原
                    bool? prevScreenUpdating = null;
                    bool? prevEnableEvents = null;
                    bool? prevDisplayAlerts = null;
                    Excel.XlCalculation? prevCalculation = null;
                    try
                    {
                        xlApp = new Excel.Application { Visible = false, DisplayAlerts = false };
                        wb = xlApp.Workbooks.Open(excelPath, ReadOnly: false);

                        // 儲存並關閉可能影響性能的 Application 設定（會在 finally 中還原）
                        try { prevScreenUpdating = xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
                        try { prevEnableEvents = xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
                        try { prevDisplayAlerts = xlApp.DisplayAlerts; xlApp.DisplayAlerts = false; } catch { }
                        try { prevCalculation = xlApp.Calculation; xlApp.Calculation = Excel.XlCalculation.xlCalculationManual; } catch { }

                        Excel.Worksheet sheet = null;
                        try { sheet = wb.Worksheets["總表"] as Excel.Worksheet; if (sheet == null) sheet = wb.Worksheets[1] as Excel.Worksheet; } catch { sheet = wb.Worksheets[1] as Excel.Worksheet; }
                        if (sheet == null) { localErr = "找不到總表工作表。"; return (false, localErr); }

                        // 嘗試解除保護
                        try
                        {
                            if (sheet.ProtectContents)
                            {
                                try { if (!string.IsNullOrEmpty(excelPassword)) sheet.Unprotect(excelPassword); else sheet.Unprotect(); } catch { }
                            }
                        }
                        catch { }

                        used = sheet.UsedRange;
                        // 取得 UsedRange 的絕對起訖座標
                        int firstRow = used.Row;
                        int firstCol = used.Column;
                        int lastRow = firstRow + used.Rows.Count - 1;
                        int lastCol = firstCol + used.Columns.Count - 1;

                        // 在前幾列（由 UsedRange 的起始列開始）掃描標頭，找出「昶亨料號」/「客戶料號」與「實發/發料」所在欄位（絕對座標）
                        int headerRowAbs = -1;
                        int materialColAbs = -1, shippedColAbs = -1;
                        string[] materialTokens = new[] { "昶亨料號", "客戶料號" };
                        // 僅接受與「實發/發料」相關的欄名；明確排除「需求/應領」等欄，避免誤判
                        string[] shippedTokens = new[] { "實發數量", "發料數量" };
                        string[] shippedExclude = new[] { "需求數量", "應領數量" };

                        int maxScanHeaderRows = Math.Min(10, used.Rows.Count);
                        // diagnostic log builder for header scanning
                        var __scanLog = new System.Text.StringBuilder();
                        __scanLog.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ExcelPath: {excelPath}");
                        __scanLog.AppendLine($"Scanning header rows: firstRow={firstRow}, lastRow={lastRow}, maxScanHeaderRows={maxScanHeaderRows}");

                        // Use central helper to normalize header text (移除非英數字並小寫)
                        // Reuse existing SanitizeHeaderForMatch to keep normalization consistent across the form

                        for (int r = firstRow; r <= Math.Min(firstRow + maxScanHeaderRows - 1, lastRow) && (materialColAbs < 0 || shippedColAbs < 0); r++)
                        {
                            int tmpMatAbs = -1, tmpShipAbs = -1;
                            for (int c = firstCol; c <= lastCol; c++)
                            {
                                var hv = (sheet.Cells[r, c] as Excel.Range)?.Value2;
                                if (hv == null) continue;
                                var hs = hv?.ToString()?.Trim() ?? string.Empty;
                                // log each header cell evaluated (original and normalized)
                                string nh = SanitizeHeaderForMatch(hs);
                                try { __scanLog.AppendLine($"Row {r} Col {c}: raw='{hs}' norm='{nh}'"); } catch { }

                                // material detection: prefer normalized contains match
                                if (tmpMatAbs < 0 && materialTokens.Any(t => nh.IndexOf(SanitizeHeaderForMatch(t), StringComparison.OrdinalIgnoreCase) >= 0)) tmpMatAbs = c;

                                if (tmpShipAbs < 0)
                                {
                                    // 改良判別：優先接受明確含「數/量/Qty」字眼的標頭，避免像「發料倉」被誤判為發料數量欄
                                    bool foundPos = false;
                                    foreach (var t in shippedTokens)
                                    {
                                        try
                                        {
                                            var nt = SanitizeHeaderForMatch(t);
                                            if (nh.IndexOf(nt, StringComparison.OrdinalIgnoreCase) >= 0)
                                            {
                                                // 對於過短或過通用的 token（例如「發料」「實發」），要求標頭同時包含數量字眼
                                                if (nt.Length <= 3)
                                                {
                                                    if (!(nh.IndexOf("數", StringComparison.OrdinalIgnoreCase) >= 0 || nh.IndexOf("量", StringComparison.OrdinalIgnoreCase) >= 0 || nh.IndexOf("qty", StringComparison.OrdinalIgnoreCase) >= 0))
                                                    {
                                                        // 跳過此 token 的匹配，因為標頭看起來不是數量欄
                                                        continue;
                                                    }
                                                }
                                                foundPos = true;
                                                break;
                                            }
                                        }
                                        catch { }
                                    }
                                    // record why we matched or not for shipped candidates
                                    try
                                    {
                                        if (foundPos) __scanLog.AppendLine($"  -> shipped token matched candidate at Col {c}: raw='{hs}' norm='{nh}'");
                                        else __scanLog.AppendLine($"  -> shipped token NOT matched at Col {c}: raw='{hs}' norm='{nh}'");
                                    }
                                    catch { }
                                    bool hasPos = foundPos;
                                    bool hasNeg = shippedExclude.Any(t => nh.IndexOf(SanitizeHeaderForMatch(t), StringComparison.OrdinalIgnoreCase) >= 0);
                                    if (hasPos && !hasNeg) tmpShipAbs = c;
                                }
                                if (tmpMatAbs >= 0 && tmpShipAbs >= 0) break;
                            }
                            if (tmpMatAbs >= 0 && tmpShipAbs >= 0)
                            {
                                headerRowAbs = r; materialColAbs = tmpMatAbs; shippedColAbs = tmpShipAbs; break;
                            }
                            if (headerRowAbs < 0)
                            {
                                if (tmpMatAbs >= 0) { headerRowAbs = r; materialColAbs = tmpMatAbs; }
                                if (tmpShipAbs >= 0) { headerRowAbs = r; shippedColAbs = tmpShipAbs; }
                            }
                        }

                        // after scanning, dump final decision into log
                        // header scan debug write removed

                        if (headerRowAbs < 1 || materialColAbs < 0 || shippedColAbs < 0)
                        {
                            localErr = "在 Excel 找不到料號或發料/實發欄位（已掃描前 10 列）。";
                            return (false, localErr);
                        }

                        // 建立：料號文字 -> 絕對資料列索引 的映射
                        var sheetMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        // 為了減少 COM 呼叫次數，一次性讀取料號欄與發料欄到本機陣列，然後在記憶體中處理
                        int firstDataRowAbs = headerRowAbs + 1;
                        int lastDataRowAbs = lastRow;
                        int dataCount = Math.Max(0, lastDataRowAbs - firstDataRowAbs + 1);

                        object[,] materialVals = null;
                        object[,] shippedVals = null;
                        Excel.Range matRange = null;
                        Excel.Range shipRange = null;
                        try
                        {
                            if (dataCount > 0)
                            {
                                try
                                {
                                    matRange = sheet.Range[sheet.Cells[firstDataRowAbs, materialColAbs], sheet.Cells[lastDataRowAbs, materialColAbs]] as Excel.Range;
                                    var matRaw = matRange?.Value2;
                                    if (matRaw is object[,]) materialVals = (object[,])matRaw;
                                    else if (matRaw != null)
                                    {
                                        materialVals = new object[1, 1];
                                        materialVals[0, 0] = matRaw;
                                    }
                                }
                                catch { materialVals = null; }

                                try
                                {
                                    shipRange = sheet.Range[sheet.Cells[firstDataRowAbs, shippedColAbs], sheet.Cells[lastDataRowAbs, shippedColAbs]] as Excel.Range;
                                    var shipRaw = shipRange?.Value2;
                                    if (shipRaw is object[,]) shippedVals = (object[,])shipRaw;
                                    else if (shipRaw != null)
                                    {
                                        shippedVals = new object[1, 1];
                                        shippedVals[0, 0] = shipRaw;
                                    }
                                }
                                catch { shippedVals = null; }
                            }

                            if (materialVals != null)
                            {
                                int rb = materialVals.GetLowerBound(0);
                                int cb = materialVals.GetLowerBound(1);
                                int rows = materialVals.GetUpperBound(0) - rb + 1;
                                for (int idx = 0; idx < rows; idx++)
                                {
                                    try
                                    {
                                        var mv = materialVals[rb + idx, cb];
                                        if (mv == null) continue;
                                        var key = mv?.ToString()?.Trim() ?? string.Empty; if (string.IsNullOrEmpty(key)) continue;
                                        int sheetRow = firstDataRowAbs + idx;
                                        if (!sheetMap.ContainsKey(key)) sheetMap[key] = sheetRow;
                                    }
                                    catch { }
                                }
                            }
                            else
                            {
                                // 若無法一次性讀取，再回退到逐列讀取（保守回退）
                                for (int r = headerRowAbs + 1; r <= lastRow; r++)
                                {
                                    var mv = (sheet.Cells[r, materialColAbs] as Excel.Range)?.Value2;
                                    if (mv == null) continue;
                                    var key = mv?.ToString()?.Trim() ?? string.Empty; if (string.IsNullOrEmpty(key)) continue;
                                    if (!sheetMap.ContainsKey(key)) sheetMap[key] = r;
                                }
                            }
                        }
                        finally
                        {
                            try { if (matRange != null) ReleaseComObjectSafe(matRange); } catch { }
                            try { if (shipRange != null) ReleaseComObjectSafe(shipRange); } catch { }
                        }

                        int dtMaterialIdx = columnMapping.ContainsKey("Material") ? columnMapping["Material"] : -1;
                        int dtShippedIdx = columnMapping.ContainsKey("Shipped") ? columnMapping["Shipped"] : -1;
                        if (dtMaterialIdx < 0 || dtShippedIdx < 0) { localErr = "內部欄位對應遺失 (Material/Shipped)。"; return (false, localErr); }

                        var writeUpdates = new Dictionary<int, object>(); // rowAbs -> new value
                        var previousValues = new Dictionary<int, double>(); // rowAbs -> old value (for 記錄 用)
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            var matObj = dt.Rows[i][dtMaterialIdx]; if (matObj == null) continue;
                            var mat = matObj?.ToString()?.Trim() ?? string.Empty; if (string.IsNullOrEmpty(mat)) continue;
                            if (!sheetMap.TryGetValue(mat, out int targetRowAbs)) continue;
                            var dtValObj = dt.Rows[i][dtShippedIdx];
                            // 若資料為 null 或空白，視為使用者清除此欄位，應回寫空字串以清空 Excel 儲存格（不可跳過）
                            if (dtValObj == null)
                            {
                                writeUpdates[targetRowAbs] = string.Empty;
                                continue;
                            }

                            // 盡量寬鬆解析數值；若為空白字串則視為清空儲存格
                            double newVal;
                            var sVal = dtValObj?.ToString() ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(sVal))
                            {
                                // 使用者把值清空，回寫空字串以清除 Excel 儲存格內容
                                writeUpdates[targetRowAbs] = string.Empty;
                                continue;
                            }

                            if (!double.TryParse(sVal, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out newVal))
                            {
                                if (!double.TryParse(sVal, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out newVal))
                                {
                                    if (TryParseDecimalFlexible(sVal, out var dec)) newVal = (double)dec; else continue;
                                }
                            }
                            // 先擷取舊值（若能取到） - 優先從一次性讀取的 shippedVals 中取得，避免逐格 COM 呼叫
                            try
                            {
                                if (shippedVals != null)
                                {
                                    int rb = shippedVals.GetLowerBound(0);
                                    int cb = shippedVals.GetLowerBound(1);
                                    int idx = targetRowAbs - firstDataRowAbs; // zero-based
                                    if (idx >= 0 && (rb + idx) <= shippedVals.GetUpperBound(0))
                                    {
                                        var oldObj = shippedVals[rb + idx, cb];
                                        double oldD = 0d;
                                        if (oldObj != null && double.TryParse(oldObj?.ToString() ?? string.Empty, out oldD)) previousValues[targetRowAbs] = oldD;
                                        else previousValues[targetRowAbs] = 0d;
                                    }
                                    else previousValues[targetRowAbs] = 0d;
                                }
                                else
                                {
                                    var oldObj = (sheet.Cells[targetRowAbs, shippedColAbs] as Excel.Range)?.Value2;
                                    double oldD = 0d;
                                    if (oldObj != null && double.TryParse(oldObj?.ToString() ?? string.Empty, out oldD)) previousValues[targetRowAbs] = oldD; else previousValues[targetRowAbs] = 0d;
                                }
                            }
                            catch { previousValues[targetRowAbs] = 0d; }

                            writeUpdates[targetRowAbs] = newVal;
                        }

                        int writeCount = Math.Max(0, lastDataRowAbs - firstDataRowAbs + 1);

                        // 若沒有任何要寫入，跳過
                        if (writeUpdates == null || writeUpdates.Count == 0)
                        {
                            // nothing to write
                        }
                        else
                        {
                            // 為了最小化 COM 傳輸，找出要更新列的最小與最大絕對列，並針對連續區段分段批次寫入。
                            var rows = writeUpdates.Keys.OrderBy(r => r).ToArray();
                            int segStart = rows[0];
                            int segPrev = rows[0];

                            // Helper: write a window [start..end] inclusive using a single Range.Value2 batch
                            Action<int, int> writeWindow = (startRowAbs, endRowAbs) =>
                            {
                                try
                                {
                                    int windowCount = endRowAbs - startRowAbs + 1;
                                    if (windowCount <= 0) return;

                                    // 讀取現有值的區間以便保留沒有變更的儲存格
                                    object[,] existingVals = null;
                                    try
                                    {
                                        var existingRange = sheet.Range[sheet.Cells[startRowAbs, shippedColAbs], sheet.Cells[endRowAbs, shippedColAbs]] as Excel.Range;
                                        var existRaw = existingRange?.Value2;
                                        if (existRaw is object[,]) existingVals = (object[,])existRaw;
                                        else if (existRaw != null)
                                        {
                                            existingVals = new object[1, 1];
                                            existingVals[0, 0] = existRaw;
                                        }
                                    }
                                    catch { existingVals = null; }

                                    object[,] buffer = new object[windowCount, 1];
                                    for (int i = 0; i < windowCount; i++)
                                    {
                                        int sheetRow = startRowAbs + i;
                                        if (writeUpdates.TryGetValue(sheetRow, out object newV))
                                        {
                                            buffer[i, 0] = newV;
                                        }
                                        else if (existingVals != null)
                                        {
                                            try
                                            {
                                                int rb = existingVals.GetLowerBound(0);
                                                int cb = existingVals.GetLowerBound(1);
                                                if ((rb + i) <= existingVals.GetUpperBound(0)) buffer[i, 0] = existingVals[rb + i, cb];
                                                else buffer[i, 0] = Type.Missing;
                                            }
                                            catch { buffer[i, 0] = Type.Missing; }
                                        }
                                        else
                                        {
                                            buffer[i, 0] = Type.Missing;
                                        }
                                    }

                                    try
                                    {
                                        var startCell = sheet.Cells[startRowAbs, shippedColAbs];
                                        var endCell = sheet.Cells[endRowAbs, shippedColAbs];
                                        var range = sheet.Range[startCell, endCell];
                                        range.Value2 = buffer;
                                    }
                                    catch
                                    {
                                        // fallback to per-row write if batch fails
                                        for (int r = startRowAbs; r <= endRowAbs; r++)
                                        {
                                            try
                                            {
                                                if (writeUpdates.TryGetValue(r, out object v)) (sheet.Cells[r, shippedColAbs] as Excel.Range).Value2 = v;
                                            }
                                            catch { }
                                        }
                                    }
                                }
                                catch { }
                            };

                            // 分段：當 row 不是連續時，先寫前段再開始新段
                            for (int i = 1; i < rows.Length; i++)
                            {
                                if (rows[i] == segPrev + 1)
                                {
                                    segPrev = rows[i];
                                    continue;
                                }
                                // non-contiguous gap -> flush previous segment
                                writeWindow(segStart, segPrev);
                                segStart = rows[i];
                                segPrev = rows[i];
                            }
                            // flush last segment
                            writeWindow(segStart, segPrev);
                        }

                        // 【優化】簡化保護邏輯，減少 COM 呼叫
                        try
                        {
                            if (!string.IsNullOrEmpty(excelPassword))
                                sheet.Protect(excelPassword, AllowFiltering: true);
                            else
                                sheet.Protect(Type.Missing, AllowFiltering: true);
                        }
                        catch { }

                        // 如果 SaveAsyncWithResult 有設定 _mergeAppendIntoWriteBack，則在此同一 Workbook
                        // 會話中，把待寫入的 _records 一併寫入「記錄」工作表，避免重複開啟/儲存同一檔案所花費的時間。
                        if (this._mergeAppendIntoWriteBack && this._records != null && this._records.Count > 0)
                        {
                            try
                            {
                                // 盡量使用與 AppendRecordsToLogSheet 相同的邏輯，但在此使用已開啟的 wb 與 sheet
                                Excel.Worksheet wsLog = null;
                                try
                                {
                                    try { wsLog = wb.Worksheets["記錄"] as Excel.Worksheet; } catch { wsLog = null; }
                                    if (wsLog == null)
                                    {
                                        try { wsLog = wb.Worksheets.Cast<Excel.Worksheet>().FirstOrDefault(w => string.Equals((w?.Name ?? string.Empty).ToString().Trim(), "記錄", StringComparison.OrdinalIgnoreCase)); } catch { wsLog = null; }
                                    }
                                }
                                catch { wsLog = null; }

                                if (wsLog == null)
                                {
                                    try { wb.Unprotect(string.IsNullOrEmpty(_cachedExcelPassword) ? Type.Missing : (object)_cachedExcelPassword); } catch { }
                                    try
                                    {
                                        wsLog = wb.Worksheets.Add() as Excel.Worksheet;
                                        if (wsLog != null)
                                        {
                                            try { wsLog.Name = "記錄"; } catch { }
                                            try { wsLog.Cells[1, 1].Value2 = "刷入時間"; } catch { }
                                            try { wsLog.Cells[1, 2].Value2 = "料號"; } catch { }
                                            try { wsLog.Cells[1, 3].Value2 = "數量"; } catch { }
                                            try { wsLog.Cells[1, 4].Value2 = "操作者"; } catch { }
                                            try { var hdr = wsLog.Range["A1:D1"]; hdr.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; hdr.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter; hdr.Font.Bold = true; } catch { }
                                        }
                                    }
                                    catch { }
                                }

                                if (wsLog != null)
                                {
                                    try { wsLog.Unprotect(string.IsNullOrEmpty(_cachedExcelPassword) ? Type.Missing : (object)_cachedExcelPassword); } catch { }

                                    // 建立 existingKeys
                                    var existingKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                    try
                                    {
                                        var usedLog = wsLog.UsedRange;
                                        if (usedLog != null)
                                        {
                                            object[,] arr = null;
                                            try
                                            {
                                                var v = usedLog.Value2;
                                                if (v is object[,]) arr = (object[,])v;
                                                else if (v != null)
                                                {
                                                    arr = new object[usedLog.Rows.Count, usedLog.Columns.Count];
                                                    for (int rr = 1; rr <= usedLog.Rows.Count; rr++)
                                                        for (int cc = 1; cc <= usedLog.Columns.Count; cc++)
                                                            arr[rr - 1, cc - 1] = (rr == 1 && cc == 1) ? v : null;
                                                }
                                            }
                                            catch { arr = null; }

                                            if (arr != null)
                                            {
                                                int rowCountArr = arr.GetLength(0);
                                                for (int r = 2; r <= rowCountArr; r++)
                                                {
                                                    try
                                                    {
                                                        object tObj = null, mObj = null, qvObj = null, uObj = null;
                                                        try { tObj = arr[r - 1, 0]; } catch { }
                                                        try { mObj = arr[r - 1, 1]; } catch { }
                                                        try { qvObj = arr[r - 1, 2]; } catch { }
                                                        try { uObj = arr[r - 1, 3]; } catch { }

                                                        var m = mObj?.ToString();
                                                        var qv = qvObj?.ToString();
                                                        var u = uObj?.ToString();
                                                        if (string.IsNullOrWhiteSpace(m) || string.IsNullOrWhiteSpace(qv)) continue;
                                                        if (!int.TryParse(qv, out int q)) continue;

                                                        DateTime dt = DateTime.MinValue;
                                                        if (tObj is double dval) { try { dt = DateTime.FromOADate(dval); } catch { dt = DateTime.MinValue; } }
                                                        else { var str = tObj?.ToString(); if (!string.IsNullOrWhiteSpace(str)) { if (double.TryParse(str, out double dv)) { try { dt = DateTime.FromOADate(dv); } catch { DateTime.TryParse(str, out dt); } } else DateTime.TryParse(str, out dt); } }

                                                        if (dt == DateTime.MinValue) continue;
                                                        var timeKey = dt.ToString("yyyyMMddHHmmss");
                                                        var key = $"{timeKey}|{NormalizeMaterialKey(m)}|{q}|{(u ?? string.Empty)}";
                                                        existingKeys.Add(key);
                                                    }
                                                    catch { }
                                                }
                                            }
                                        }
                                    }
                                    catch { }

                                    // 準備要追加的實際列
                                    var toAppend = new List<Dto.記錄Dto>();
                                    var batchKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                    foreach (var rec in this._records)
                                    {
                                        if (rec == null) continue;
                                        var recTimeKey = rec.刷入時間.ToString("yyyyMMddHHmmss");
                                        var recKey = $"{recTimeKey}|{NormalizeMaterialKey(rec.料號)}|{rec.數量}|{(rec.操作者 ?? string.Empty)}";
                                        if (existingKeys.Contains(recKey)) continue;
                                        if (batchKeys.Contains(recKey)) continue;
                                        batchKeys.Add(recKey);
                                        toAppend.Add(rec);
                                    }

                                    if (toAppend.Count > 0)
                                    {
                                        int lastRowLog = 1;
                                        try { lastRowLog = wsLog.UsedRange.Rows.Count; } catch { lastRowLog = 1; }
                                        int appendStartRow = Math.Max(2, lastRowLog + 1);
                                        int appendRowCount = toAppend.Count;
                                        object[,] data = new object[appendRowCount, 4];
                                        for (int i = 0; i < appendRowCount; i++)
                                        {
                                            var rec = toAppend[i];
                                            try { data[i, 0] = rec.刷入時間.ToString("yyyy/MM/dd HH:mm:ss"); } catch { data[i, 0] = string.Empty; }
                                            try { data[i, 1] = rec.料號; } catch { data[i, 1] = string.Empty; }
                                            try { data[i, 2] = rec.數量; } catch { data[i, 2] = 0; }
                                            try { data[i, 3] = rec.操作者; } catch { data[i, 3] = string.Empty; }
                                        }

                                        // Batch attempt
                                        try
                                        {
                                            var startCell = wsLog.Cells[appendStartRow, 1];
                                            var endCell = wsLog.Cells[appendStartRow + appendRowCount - 1, 4];
                                            var range = wsLog.Range[startCell, endCell];
                                            range.Value2 = data;
                                            // 將 toAppend 全部視為成功新增
                                            try { _lastAppendedRecords = toAppend.ToList(); } catch { }
                                            // material write debug path removed
                                        }
                                        catch (Exception)
                                        {
                                            // 已移除：舊式臨時診斷檔路徑（excel_write_error.log）

                                            // fallback per-row
                                            try
                                            {
                                                int rr = appendStartRow;
                                                int fallbackCount = 0;
                                                foreach (var rec in toAppend)
                                                {
                                                    try
                                                    {
                                                        wsLog.Cells[rr, 1].Value2 = rec.刷入時間.ToString("yyyy/MM/dd HH:mm:ss");
                                                        wsLog.Cells[rr, 2].Value2 = rec.料號;
                                                        wsLog.Cells[rr, 3].Value2 = rec.數量;
                                                        wsLog.Cells[rr, 4].Value2 = rec.操作者;
                                                        rr++;
                                                    }
                                                    catch { fallbackCount++; rr++; }
                                                }
                                                try { _lastAppendedRecords = toAppend.ToList(); } catch { }
                                                // 已移除：舊式 material_write_debug.log 宣告（改用集中式 Logger）
                                            }
                                            catch { }
                                        }
                                    }

                                    // 為確保「記錄」工作表的 A:D 欄位可以完整顯示新增資料，僅針對該工作表執行 AutoFit
                                    try { wsLog.Columns["A:D"].AutoFit(); } catch { }
                                    try { wsLog.Protect(string.IsNullOrEmpty(_cachedExcelPassword) ? Type.Missing : (object)_cachedExcelPassword, AllowFiltering: false); } catch { }
                                }
                            }
                            catch { }
                        }

                        // 【關鍵修復】確保 Save() 成功執行，並驗證結果
                        try
                        {
                            wb.Save();
                            // 診斷：記錄成功存檔
                            // save-success diagnostic removed
                        }
                        catch (Exception saveEx)
                        {
                            // 【重要】存檔失敗時必須拋出異常，不能返回 true！
                            throw new Exception($"Excel 檔案儲存失敗: {saveEx.Message}", saveEx);
                        }

                        // 在關閉前嘗試還原 Application 設定，避免影響使用者環境
                        try { if (prevCalculation.HasValue) xlApp.Calculation = prevCalculation.Value; } catch { }
                        try { if (prevDisplayAlerts.HasValue) xlApp.DisplayAlerts = prevDisplayAlerts.Value; } catch { }
                        try { if (prevEnableEvents.HasValue) xlApp.EnableEvents = prevEnableEvents.Value; } catch { }
                        try { if (prevScreenUpdating.HasValue) xlApp.ScreenUpdating = prevScreenUpdating.Value; } catch { }
                        try { wb.Close(false); } catch { }
                        try { xlApp.Quit(); } catch { }
                        return (true, (string)null);
                    }
                    catch (Exception ex)
                    {
                        try { if (wb != null) wb.Close(false); } catch { }
                        try { if (xlApp != null) xlApp.Quit(); } catch { }
                        // Diagnostic: record full exception to excel_write_error.log to help reproduce COM/Excel errors
                        // exception diagnostic write removed
                        return (false, ex.Message);
                    }
                    finally
                    {
                        try { if (used != null) ReleaseComObjectSafe(used); } catch { }
                        try { if (wb != null) ReleaseComObjectSafe(wb); } catch { }
                        try { if (xlApp != null) ReleaseComObjectSafe(xlApp); } catch { }
                        // 在 STA 執行緒完全結束前，進行強制垃圾回收
                        // 這確保所有 COM 物件引用都被釋放
                        try
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                        }
                        catch { }
                    }
                });

                // result is (bool ok, string err)
                // 【優化】簡化 GC 策略，減少等待時間 (省時 200+ms)
                try
                {
                    GC.Collect();
                    System.Threading.Thread.Sleep(100);  // short cleanup pause
                }
                catch { }

                // 停止計時並顯示耗時（方便本地測試觀察）。顯示訊息以避免破壞原本邏輯；若 UI 無法顯示或在無頭環境，則捕捉例外並忽略。
                try { /* timing removed for main write elapsed */ } catch { }

                errMsg = result.Item2;
                return result.Item1;
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
        }

        #endregion

        /// <summary>
        /// 將目前表單資料存回 Excel 的可等待方法（可被其他流程呼叫並回傳成功/失敗）。
        /// 回傳 true 表示存檔成功；false 表示失敗或無資料可存。
        /// </summary>
        // 新 SaveAsyncWithResult，回傳 SaveResultDto
        /// <summary>
        /// 非同步儲存目前匯入資料的執行方法，回傳包含成功狀態與錯誤訊息的 <see cref="SaveResultDto"/>。
        /// 此方法封裝資料驗證、資料庫寫入以及必要的事務處理，需在非 UI 執行緒上執行或以 await 呼叫以避免凍結介面。
        /// </summary>
        /// <returns>操作結果（成功或失敗）與錯誤描述。</returns>
        private async Task<SaveResultDto> SaveAsyncWithResult()
        {
            // 防呆：避免重入
            var result = new SaveResultDto();
            if (_isSaving)
            {
                try { MessageBox.Show("目前已有存檔作業在進行中，請稍後再試。", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                result.Success = false;
                result.ErrorMessage = "存檔作業進行中，無法重複存檔。";
                return result;
            }
            if (!CheckExcelAvailable()) { result.Success = false; result.ErrorMessage = "Excel 未就緒。"; return result; }

            _isSaving = true;
            UpdateButtonStates();
            SaveAndDisableAllButtons();

            // 保存原始游標與 UseWaitCursor 狀態（使用本地變數以避免覆寫外層呼叫者的 _prevUseWaitCursor）
            var prevCursor = Cursor.Current;
            var prevUseWaitCursorLocal = Application.UseWaitCursor;

            // 設定等待游標
            ShowOperationOverlay("存檔中，請稍候...");
            Application.UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;
            try { await Task.Delay(80); } catch { }

            bool _showFinalMsg = false;
            string? _finalMsg = null;
            string? _finalTitle = null;
            MessageBoxIcon _finalIcon = MessageBoxIcon.None;

            try
            {
                var dt = this.dgv備料單.DataSource as DataTable;
                if (dt == null || this._columnMapping == null)
                {
                    _showFinalMsg = true; _finalMsg = "無可回寫的資料或欄位對應遺失。"; _finalTitle = "存檔"; _finalIcon = MessageBoxIcon.Warning;
                    result.Success = false;
                    result.ErrorMessage = _finalMsg;
                    return result;
                }

                string? err = null;
                bool ok = false;
                string pwd = GetExcelPassword();

                try
                {

                    if (!string.IsNullOrWhiteSpace(this.currentExcelPath) && File.Exists(this.currentExcelPath))
                    {
                        int retryCount = 0;
                        // 檢查檔案是否被鎖定
                        bool isLocked = IsFileLocked(this.currentExcelPath ?? string.Empty);

                        // initial file-lock diagnostic removed

                        while (isLocked)
                        {
                            retryCount++;

                            // 檢查是否有 Excel 程序正在執行
                            var excelProcesses = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                            string processInfo = excelProcesses.Length > 0
                                ? $"\n\n偵測到 {excelProcesses.Length} 個 Excel 程序正在執行。\n建議：請到工作管理員結束所有 Excel.exe 程序後再重試。"
                                : "\n\n未偵測到 Excel 程序，但檔案仍被其他程序鎖定。";

                            var dlg = MessageBox.Show(
                                $"欲回寫的 Excel 檔案目前正被開啟或鎖定。請先關閉該檔案後按「重試」。按「取消」可放棄本次存檔。{processInfo}\n\n(已重試 {retryCount} 次)",
                                "檔案被鎖定",
                                MessageBoxButtons.RetryCancel,
                                MessageBoxIcon.Warning,
                                MessageBoxDefaultButton.Button1);

                            if (dlg == DialogResult.Cancel)
                            {
                                _showFinalMsg = true;
                                _finalMsg = "使用者已取消：Excel 檔案被鎖定，存檔已中止。";
                                _finalTitle = "存檔已取消";
                                _finalIcon = MessageBoxIcon.Information;
                                result.Success = false;
                                result.ErrorMessage = "使用者取消存檔：檔案被鎖定。";
                                return result;
                            }

                            // 使用者選擇重試：進行強制 GC + 等候再檢查（避免 busy-loop）
                            try
                            {
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                GC.Collect();
                                await Task.Delay(1500);  // 等待 1.5 秒再檢查
                            }
                            catch { }

                            // 重新檢查是否仍被鎖定
                            isLocked = IsFileLocked(this.currentExcelPath ?? string.Empty);

                            // retry check diagnostic removed
                        }
                    }
                    else
                    {
                        // 【診斷】記錄為何跳過檔案鎖定檢查
                        // skipped file-lock check diagnostic removed
                    }

                    // 檔案已確認可以寫入，開始嘗試寫入並進行智能重試
                    const int maxWriteRetries = 3;
                    for (int attempt = 0; attempt < maxWriteRetries; attempt++)
                    {
                        try
                        {
                            // 在每次寫入前清理系統狀態
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();

                            if (attempt > 0)
                            {
                                try { await Task.Delay(2000 * attempt); } catch { }  // 指數退避延遲
                            }

                            // 【最終安全檢查】在真正寫入前,再次確認檔案可寫入且沒有 Excel 程序鎖定
                            if (!string.IsNullOrWhiteSpace(this.currentExcelPath) && File.Exists(this.currentExcelPath))
                            {
                                // 檢查是否仍有 Excel 程序執行中
                                var excelProcs = System.Diagnostics.Process.GetProcessesByName("EXCEL");
                                if (excelProcs.Length > 0)
                                {
                                    // 嘗試檢查檔案是否真的可寫入
                                    bool isStillLocked = IsFileLocked(this.currentExcelPath ?? string.Empty);
                                    if (isStillLocked)
                                    {
                                        ok = false;
                                        err = "檔案仍被 Excel 程序鎖定,無法寫入。請確保 Excel 已完全關閉。";
                                        // final safety-check diagnostic removed
                                        break;  // 停止重試,直接失敗
                                    }
                                }
                            }

                            // Measure total time for the full writeback (main sheet write + record append)
                            // __writeSw timing removed (Start)
                            // 啟用合併 append 行為，避免後續重複開檔造成額外延遲
                            try { _mergeAppendIntoWriteBack = true; } catch { }
                            try
                            {
                                ok = await Task.Run(() => WriteBackShippedQuantitiesToExcelBatch(this.currentExcelPath ?? string.Empty, dt, this._columnMapping, pwd, out err));
                            }
                            finally
                            {
                                try { _mergeAppendIntoWriteBack = false; } catch { }
                            }

                            if (ok)
                            {
                                // 寫入成功，進行強制垃圾回收以清理 COM 物件參考
                                try
                                {
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    GC.Collect();
                                    // 以保守輪詢替代固定等待：嘗試以非獨佔方式打開檔案，若可開啟則立即繼續
                                    try
                                    {
                                        var exeDir = AppDomain.CurrentDomain.BaseDirectory ?? Directory.GetCurrentDirectory();
                                        // diagnostic material_write_debug path removed
                                        int maxWaitMs = 3000;
                                        int intervalMs = 150;
                                        int waited = 0;
                                        bool released = false;

                                        while (waited < maxWaitMs)
                                        {
                                            try
                                            {
                                                using (var fs = new FileStream(this.currentExcelPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                                                {
                                                    // able to open exclusively -> file released
                                                    released = true;
                                                    break;
                                                }
                                            }
                                            catch
                                            {

                                                try { await Task.Delay(intervalMs); } catch { }
                                                waited += intervalMs;
                                            }
                                        }
                                        try
                                        {
                                            try { } catch { }
                                        }
                                        catch { }
                                    }
                                    catch { }
                                }
                                catch { }
                                break;  // 成功，退出重試迴圈
                            }
                            else
                            {
                                // 寫入失敗，嘗試重試
                                if (attempt < maxWriteRetries - 1)
                                {
                                    // Debug log removed
                                }
                                else
                                {
                                    // 最後一次嘗試失敗
                                    // Debug log removed
                                }
                            }
                        }
                        catch (Exception exWrite)
                        {
                            err = exWrite.Message;
                            if (attempt < maxWriteRetries - 1)
                            {
                                // Debug log removed
                            }
                            else
                            {
                                ok = false;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    ok = false; err = ex.Message;
                }

                if (!ok)
                {
                    _showFinalMsg = true; _finalMsg = "存檔失敗：" + (err ?? "未知錯誤"); _finalTitle = "錯誤"; _finalIcon = MessageBoxIcon.Error;
                    result.Success = false;
                    result.ErrorMessage = err;
                    // failure diagnostic removed
                }
                else
                {
                    // 【修訂】若 WriteBackShippedQuantitiesToExcelBatch 已回傳 ok=true，
                    // 直接信任其結果，不再做嚴格的時間戳驗證（因系統可能有時間延遲）
                    if (ok)
                    {
                        // 驗證成功，記錄存檔完成
                        // 暫停 dirty marking，避免後續操作誤觸 _isDirty 更新
                        bool prevSuspend = _suspendDirtyMarking;
                        try { _suspendDirtyMarking = true; }
                        catch { }

                        // 當 _mergeAppendIntoWriteBack 為 true 時，AppendRecords 已由 WriteBackShippedQuantitiesToExcelBatch
                        // 在同一 Workbook session 中完成；不再於此處另行開啟工作簿以避免重複開關檔案。
                        finally
                        {
                            try { /* timing removed: __writeSw stop/elapsed */ } catch { }
                        }

                        // 不要清空 _lastAppendedRecords（這會移除剛記錄的已新增清單），僅清空原始待寫清單 _records
                        try { _records?.Clear(); } catch { }

                        // 清空 DataTable 的變更追蹤（呼叫 AcceptChanges），確保存檔後沒有未儲存的變更記錄
                        try
                        {
                            if (dt != null)
                            {
                                var changes = dt.GetChanges();
                                if (changes != null) dt.AcceptChanges();
                            }
                        }
                        catch { }

                        // 確保 _isDirty 在 dirty marking 暫停時被設為 false
                        try { _isDirty = false; } catch { }

                        _showFinalMsg = true;
                        _finalMsg = "已成功存檔並保護 Excel 檔案。";
                        _finalTitle = "存檔完成";
                        _finalIcon = MessageBoxIcon.Information;
                        try { RestorePreservedRedHighlights(); } catch { }
                        result.Success = true;
                        result.ErrorMessage = string.Empty;

                        // 恢復 dirty marking 狀態
                        try { _suspendDirtyMarking = prevSuspend; } catch { }
                    }
                }
            }
            finally
            {
                // 優先還原游標：確保即使失敗也立即還原
                try { Cursor.Current = prevCursor; } catch { }
                try { Application.UseWaitCursor = prevUseWaitCursorLocal; } catch { }
                try { HideOperationOverlay(); } catch { }

                // 在還原 UI 時暫停 dirty marking，防止無意間觸發
                bool prevSuspend2 = _suspendDirtyMarking;
                try { _suspendDirtyMarking = true; } catch { }

                try { RestoreAllButtons(); } catch { }
                _isSaving = false;
                UpdateButtonStates();
                // 確保按鈕狀態符合實際資料情況
                try { UpdateMainButtonsEnabled(); } catch { }

                // 恢復 dirty marking 狀態
                try { _suspendDirtyMarking = prevSuspend2; } catch { }

                try
                {
                    if (_showFinalMsg && !string.IsNullOrEmpty(_finalMsg))
                    {
                        try { MessageBox.Show(this, _finalMsg, _finalTitle ?? string.Empty, MessageBoxButtons.OK, _finalIcon); } catch { }
                    }
                }
                catch { }
                try { RestorePreservedRedHighlights(); } catch { }
                try { MarkShortagesInGrid(); } catch { }
            }

            return result;
        }

        /// <summary>
        /// [事件] 備料單存檔按鈕點擊事件。
        /// 觸發時會檢查是否已有存檔作業進行中，並呼叫 SaveAsyncWithResult 進行存檔。
        /// 若存檔失敗，會顯示錯誤訊息（包含日誌內容）。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private async void btn備料單存檔_Click(object sender, EventArgs e)
        {
            // 防呆：避免重入
            if (_isSaving)
            {
                try { MessageBox.Show("目前已有存檔作業在進行中，請稍後再試。", "提醒", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                return;
            }
            if (!CheckExcelAvailable()) return;

            // Delegate reentrancy guard and UI state management to SaveAsyncWithResult
            var saveResult = await SaveAsyncWithResult();
            if (saveResult == null || !saveResult.Success)
            {
                string? errorMsg = saveResult?.ErrorMessage;

                // 移除測試性 / 偵錯用的外部日誌檔讀取（excel_write_error.log）
                // 直接使用 SaveResult 提供的錯誤訊息，若無則顯示通用錯誤提示。
                if (string.IsNullOrWhiteSpace(errorMsg))
                    errorMsg = "存檔失敗，請聯絡管理員或重新嘗試。";

                // 同步嘗試將錯誤寫入集中式 Logger（非必要，若 Logger 發生例外則忽略）
                try
                {
                    try { Automatic_Storage.Utilities.Logger.LogErrorAsync($"Save failed: {errorMsg}").GetAwaiter().GetResult(); } catch { }
                }
                catch { }

                try { MessageBox.Show(this, errorMsg, "存檔失敗", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
            }
        }

        /// <summary>
        /// 將 _records 追加寫入 Excel 的「記錄」工作表（保留標題列，僅新增不修改既有資料）。
        /// 會檢查既有資料以避免重複額外寫入。預期 records 為最近的批次。
        /// </summary>
        /// <param name="excelPath">目標 Excel 檔案的完整路徑。</param>
        /// <param name="records">待附加的記錄清單（每筆為 <see cref="Dto.記錄Dto"/>）。</param>
        /// <returns>成功附加並寫入的記錄清單。</returns>
        private List<Dto.記錄Dto> AppendRecordsToLogSheet(string excelPath, List<Dto.記錄Dto> records)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath)) return new List<Dto.記錄Dto>();
            if (records == null || records.Count == 0) return new List<Dto.記錄Dto>();

            // 改用注入型 ExcelService 實作
            try
            {
                // 若有強型別 IExcelService，優先使用
                if (_typedExcelService != null)
                {
                    return _typedExcelService.AppendRecordsToLogSheet(excelPath, records, _cachedExcelPassword);
                }
                // 若有 dynamic ExcelService，則動態呼叫
                if (_excelService != null)
                {
                    return _excelService.AppendRecordsToLogSheet(excelPath, records, _cachedExcelPassword);
                }
            }
            catch (Exception ex)
            {
                // 保留原有錯誤處理
                try { Automatic_Storage.Utilities.Logger.LogErrorAsync($"Excel 寫入失敗: {ex.Message}").GetAwaiter().GetResult(); } catch { }
            }
            // 若都無法寫入，回傳空集合
            return new List<Dto.記錄Dto>();
        }

        #endregion

    }

}