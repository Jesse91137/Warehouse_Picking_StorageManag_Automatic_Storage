using System; // 引用 System 命名空間
using System.Collections.Generic; // 引用泛型集合命名空間
using System.ComponentModel; // 引用元件模型命名空間
using System.Data; // 引用資料處理命名空間
using System.Data.OleDb; // 引用 OleDb 資料庫命名空間
using System.Data.SqlClient; // 引用 SQL Server 資料庫命名空間
using System.Diagnostics; // 引用診斷命名空間
using System.Drawing; // 引用繪圖命名空間
using System.Globalization; // 引用全球化命名空間
using System.IO; // 引用檔案處理命名空間
using System.Reflection; // 引用反射命名空間
using System.Runtime.InteropServices; // 引用互操作命名空間
using System.Security; // 引用安全性命名空間
using System.Text; // 引用文字處理命名空間
using System.Threading; // 引用執行緒命名空間
using System.Windows.Forms; // 引用視窗表單命名空間
using System.Windows.Threading; // 引用 WPF 計時器命名空間

namespace Automatic_Storage // 命名空間：Automatic_Storage
{
    public partial class Form1 : Form // Form1 類別，繼承 Form
    {
        // 單一實例：備料單匯入視窗（返回只隱藏；按 X 關閉才釋放並重置）
        private Form備料單匯入 importForm;
        public Form1() // 建構函式
        {
            InitializeComponent(); // 初始化元件
            isLoaded = false; // 設定初始載入狀態
        }

        #region 視窗ReSize // 視窗大小調整相關變數
        int X = new int();  // 視窗寬度
        int Y = new int(); // 視窗高度
        float fgX = new float(); // 寬度縮放比例
        float fgY = new float(); // 高度縮放比例
        bool isLoaded;  // 是否已設定各控制的尺寸資料到 Tag 屬性
        #endregion // 結束視窗大小調整區塊

        #region 參數 // 參數區塊

        /// <summary>
        /// Excel 輸入計數、查詢全部按鈕、料號查詢按鈕
        /// </summary>
        int inputexcelcount = 0, btn_fAll = 0, btn_itemP = 0; // Excel 輸入次數、查詢全部按鈕、料號查詢按鈕

        /// <summary>
        /// 資料表物件
        /// </summary>
        DataTable dt = new DataTable(); // 資料表

        /// <summary>
        /// 資料集物件
        /// </summary>
        DataSet dsData = new DataSet(); // 資料集

        /// <summary>
        /// 靜態 Mutex 物件
        /// </summary>
        static Mutex m; // 靜態 Mutex

        /// <summary>
        /// WPF 計時器
        /// </summary>
        DispatcherTimer dispatcherTimer = new DispatcherTimer(); // WPF 計時器

        /// <summary>
        /// SOP 路徑、名稱、舊版、最新版
        /// </summary>
        public string sop_path = "", SopName = "", version_old = "", version_new = ""; // SOP 路徑、名稱、舊版、最新版

        /// <summary>
        /// 設定檔名稱
        /// </summary>
        public string filename = "Setup.ini"; // 設定檔名稱

        /// <summary>
        /// 設定檔操作物件
        /// </summary>
        SetupIniIP ini = new SetupIniIP(); // 設定檔操作物件

        /// <summary>
        /// 日誌物件
        /// </summary>
        Log GetLog = new Log(); // 日誌物件

        /// <summary>
        /// CMC_DC 物件
        /// </summary>
        CMC_DC cmc = new CMC_DC(); // CMC_DC 物件

        // 分頁相關變數 // 分頁功能相關

        /// <summary>
        /// 每頁顯示 25 筆資料
        /// </summary>
        int pageSize = 25; // 分頁大小

        /// <summary>
        /// 當前頁碼
        /// </summary>
        private int currentPage = 1; // 當前頁碼

        /// <summary>
        /// 查詢類型：0=預設查詢，1=儲位查詢，2=料號查詢，3=料號+儲位組合查詢
        /// </summary>
        int queryType = 0; // 查詢類型

        // 儲位查詢相關變數 // 儲位查詢功能
        /// <summary>
        /// 存儲目前的儲位查詢條件
        /// </summary>
        string current_Position = ""; // 儲位查詢條件

        /// <summary>
        /// 查詢是否為歷史查詢 => true 為歷史查詢，false 為當日查詢
        /// </summary>
        bool is_History_Query = false; // 是否歷史查詢

        // 料號查詢相關變數 // 料號查詢功能
        /// <summary>
        /// 存儲目前的料號查詢條件
        /// </summary>
        string current_ItemNo = ""; // 料號查詢條件

        /// <summary>
        /// 勾選客
        /// </summary>
        bool is_Rad_E = true; // 勾選客

        // 料號+儲位組合查詢相關變數 // 組合查詢功能
        private string current_CombiItem = ""; // 存儲目前的料號查詢條件

        /// <summary>
        /// 存儲目前的儲位查詢條件
        /// </summary>
        private string current_CombiPosition = ""; // 組合儲位查詢條件

        /// <summary>
        /// 存儲組合查詢是否為歷史查詢
        /// </summary>
        private bool is_Combi_History_Query = false; // 組合查詢是否歷史

        /// <summary>
        /// 存儲組合查詢是否為英文料號查詢條件
        /// </summary>
        private bool is_Combi_E_Query = true; // 組合查詢是否英文料號

        #endregion // 結束參數區塊

        /// <summary>
        /// 設定檔操作類別，提供 ini 檔案的讀寫功能
        /// </summary>
        public class SetupIniIP // 設定檔操作類別
        {
            //api ini
            /// <summary>
            /// ini 檔案路徑
            /// </summary>
            public string path; // ini 檔案路徑

            /// <summary>
            /// 寫入 ini 檔案內容
            /// </summary>
            /// <param name="section">區段名稱</param>
            /// <param name="key">鍵名</param>
            /// <param name="val">值</param>
            /// <param name="filePath">檔案路徑</param>
            [DllImport("kernel32", CharSet = CharSet.Unicode)] // 匯入 Windows API，寫入 ini 檔案
            private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath); // 寫入 ini 檔案

            /// <summary>
            /// 讀取 ini 檔案內容
            /// </summary>
            /// <param name="section">區段名稱</param>
            /// <param name="key">鍵名</param>
            /// <param name="def">預設值</param>
            /// <param name="retVal">回傳字串</param>
            /// <param name="size">字串長度</param>
            /// <param name="filePath">檔案路徑</param>
            [DllImport("kernel32", CharSet = CharSet.Unicode)] // 匯入 Windows API，讀取 ini 檔案
            private static extern int GetPrivateProfileString(string section,
            string key, string def, StringBuilder retVal,
            int size, string filePath); // 讀取 ini 檔案

            /// <summary>
            /// 寫入指定區段與鍵值到 ini 檔案
            /// </summary>
            /// <param name="Section">區段名稱</param>
            /// <param name="Key">鍵名</param>
            /// <param name="Value">值</param>
            /// <param name="inipath">ini 檔案名稱</param>
            public void IniWriteValue(string Section, string Key, string Value, string inipath)
            {
                // 寫入 ini 檔案內容
                WritePrivateProfileString(Section, Key, Value, Application.StartupPath + "\\" + inipath); // 呼叫 Windows API 寫入
            }

            /// <summary>
            /// 讀取指定區段與鍵值的內容
            /// </summary>
            /// <param name="Section">區段名稱</param>
            /// <param name="Key">鍵名</param>
            /// <param name="inipath">ini 檔案名稱</param>
            /// <returns>回傳讀取到的字串內容</returns>
            public string IniReadValue(string Section, string Key, string inipath)
            {
                // 建立暫存字串
                StringBuilder temp = new StringBuilder(255); // 用來存放讀取結果
                // 讀取 ini 檔案內容
                int i = GetPrivateProfileString(Section, Key, "", temp, 255, Application.StartupPath + "\\" + inipath); // 呼叫 Windows API 讀取
                // 回傳讀取到的字串
                return temp.ToString(); // 回傳結果
            }
        }

        /// <summary>
        /// Form1_Load
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            X = this.Width;//獲取窗體的寬度
            Y = this.Height;//獲取窗體的高度
            isLoaded = true;// 已設定各控制項的尺寸到Tag屬性中
            SetTag(this);//調用方法

            this.Text += "  " + FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion.ToString();
            Form.CheckForIllegalCrossThreadCalls = false;
            if (IsMyMutex("Automatic_Storage"))
            {
                MessageBox.Show("程式正在執行中!!");
                Dispose();//關閉
            }
            version_old = ini.IniReadValue("Version", "version", filename);
            version_new = selectVerSQL_new("Auto_Storage_M"); // 測試機
            //lbl_ver.Text = "VER:V" + version_old;
            //判斷自動更新程式是否啟動
            int v_old = Convert.ToInt32(version_old.Replace(".", ""));
            int v_new = Convert.ToInt32(version_new.Replace(".", ""));
            if (v_old < v_new)
            {
                MessageBox.Show("有新版本更新VER: V" + version_new);
                autoupdate();
            }
            else
            {
                panel1.Visible = false;
                panel5.Visible = false;
                // 使用者登入
                Login login = new Login();
                login.Owner = this;
                login.ShowDialog();

                Initi();
                dataBind();

                // 顯示分頁按鈕和標籤
                btnPreviousPage.Visible = true;
                btnNextPage.Visible = true;
                // 初始化顯示頁碼
                if (lblCurrentPage != null)
                {
                    lblCurrentPage.Visible = true;
                    lblCurrentPage.Text = $"第 {currentPage} 頁";
                }
            }
            this.Text += "       " + Login.User_name;
        }

        /// <summary>
        /// 判斷指定名稱的 Mutex 是否已存在，避免程式重複執行
        /// </summary>
        /// <param name="prgname">Mutex 名稱（通常為程式名稱）</param>
        /// <returns>如果已存在則回傳 true，否則回傳 false</returns>
        /// <remarks>
        /// 此方法會建立一個新的 Mutex，並判斷是否已經有相同名稱的 Mutex 存在於系統中。
        /// 若 Mutex 已存在，表示程式已執行中，回傳 true；否則回傳 false。
        /// </remarks>
        /// <example>
        /// <code>
        /// if (IsMyMutex("Automatic_Storage")) { MessageBox.Show("程式正在執行中!!"); }
        /// </code>
        /// </example>
        private bool IsMyMutex(string prgname)
        {
            bool IsExist; // 是否已存在 Mutex
            m = new Mutex(true, prgname, out IsExist); // 建立 Mutex，並取得是否已存在
            GC.Collect(); // 釋放記憶體資源
            if (IsExist) // 如果 Mutex 是新建立的（表示沒有重複執行）
            {
                return false; // 沒有重複執行
            }
            else // 如果 Mutex 已存在（表示程式已執行中）
            {
                return true; // 已有重複執行
            }
        }


        #region 查詢

        /// <summary>
        /// 料號查詢 Enter
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txt_item_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                bool h_Flag = false; // 歷史查詢

                string strShift = string.Empty;

                // 只查料號，但可能被輸入了要清空
                txt_position.Text = string.Empty;

                //料號-歷史查詢
                if (txt_itemP5.Text != "" && txt_itemP5.Text.Length >= 3)
                {
                    h_Flag = true;
                }

                if (txt_item.Text != "" && txt_item.Text.Length >= 3 || h_Flag)
                {
                    try
                    {
                        // 重置當前頁碼到第一頁
                        currentPage = 1;

                        // 判斷是一般查詢或歷史查詢-true 為歷史查詢
                        // 先存儲查詢條件，以便在換頁時使用
                        strShift = h_Flag ? txt_itemP5.Text.Trim().ToUpper() : txt_item.Text.Trim().ToUpper();

                        // 存儲目前的查詢條件
                        current_ItemNo = strShift;
                        is_History_Query = h_Flag;
                        is_Rad_E = rad_E.Checked;
                        queryType = 2; // 設置為料號查詢

                        // 獲取第一頁數據
                        GetItemPagedData(currentPage);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("請確認料號是否正確或沒有輸入");
                }
                //輸入搜尋條件後反白
                txt_item.SelectAll();
                txt_itemP5.SelectAll();
            }
        }

        /// <summary>
        /// 儲位查詢 Enter
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void txt_position_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                //變數宣告
                bool h_Flag = false;    //歷史查詢
                // 儲位
                string position = txt_position.Text;

                // 只要查料號，但可能被輸入了要清空
                txt_item.Text = string.Empty;

                //分析儲位
                //if (txt_position.Text.IndexOf("-") > 0)
                //{
                //    string[] postion_split = txt_position.Text.Split('-');
                //    position = postion_split[0];
                //}

                //歷史查詢
                if (txt_siteP5.Text != "")
                {
                    h_Flag = true;
                }
                if (txt_position.Text != "" || h_Flag)
                {
                    try
                    {
                        // 重置當前頁碼到第一頁
                        currentPage = 1;

                        // 先存儲查詢條件，以便在換頁時使用
                        if (h_Flag)
                        {
                            position = txt_siteP5.Text;
                        }

                        // 存儲目前的查詢條件
                        current_Position = position;
                        is_History_Query = h_Flag;
                        queryType = 1; // 設置為儲位查詢

                        // 獲取第一頁數據
                        GetPositionPagedData(currentPage);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("請確認儲位是否正確或沒有輸入");
                }
            }
        }

        /// <summary>
        /// 查詢全部
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click_1(object sender, EventArgs e)
        {
            clsAll();
            queryType = 0; // 重置為預設查詢類型
            dataBind();
            // Excel匯出資料
            dataBind_Old();
        }

        /// <summary>
        /// 查詢By頁數
        /// </summary>
        public void dataBind()
        {
            DataTable currentData = GetPagedData(currentPage);
            if (currentData != null)
            {
                dataGridView1.DataSource = currentData;
                if (lblCurrentPage != null)
                {
                    lblCurrentPage.Text = $"第 {currentPage} 頁";
                }
                sumC();
            }
            btn_fAll = 0;
        }

        /// <summary>
        /// 料號+儲位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_itemSite_Click(object sender, EventArgs e)
        {
            btn_itemP = 1;
            btn_combi_Click(sender, e);
        }

        /// <summary>
        /// 料號+儲位搜尋
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_combi_Click(object sender, EventArgs e)
        {
            queryType = 3; // 查詢類型- 料號+儲位

            bool h_flag = false;

            if (txt_itemP5.Text != "" && txt_siteP5.Text != "" && txt_itemP5.Text.Length > 3)
            {
                h_flag = true;
            }
            if (txt_item.Text != "" && txt_position.Text != "" && txt_item.Text.Length > 3 || h_flag)
            {
                try
                {
                    string stritem = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name
                                    ,a.Mark,PCB_DC,CMC_DC,a.sno  
                                    from Automatic_Storage_Detail a 
                                    left join Automatic_Storage_Package b on a.Package = b.code 
                                    left join Automatic_Storage_User c on a.Input_UserNo = c.User_No 
                                    left join Automatic_Storage_User d on a.Output_UserNo = d.User_No ";
                    if (rad_E.Checked) //昶
                    {
                        stritem += "where a.Unit_No=@unitNo and  Position=@position and  Item_No_Master = @master and amount >0 " +
                                    "order by Position";
                    }
                    else if (rad_C.Checked) // 客
                    {
                        stritem += "where a.Unit_No=@unitNo and  Position=@position and  Item_No_Slave = @master and amount >0 " +
                                   "order by Position";
                    }
                    //歷史資料
                    if (btn_itemP == 1)
                    {
                        stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate" +
                                  ",d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  " +
                                "from Automatic_Storage_Detail a " +
                                "left join Automatic_Storage_Package b on a.Package = b.code " +
                                "left join Automatic_Storage_User c on a.Input_UserNo = c.User_No " +
                                "left join Automatic_Storage_User d on a.Output_UserNo = d.User_No " +
                                "where a.Unit_No=@unitNo and  Position=@position and  Item_No_Master = @master " +
                                "order by Position";
                    }
                    SqlParameter[] parm = new SqlParameter[]
                    {
                        new SqlParameter("unitNo", Login.Unit_No),
                        new SqlParameter("position", h_flag?txt_siteP5.Text?.Trim():txt_position.Text),
                        new SqlParameter("master", h_flag?txt_itemP5.Text?.Trim():txt_item.Text?.Trim().ToUpper())
                    };

                    dt = db.ExecuteDataTable(stritem, CommandType.Text, parm);
                    if (dt.Rows.Count > 0)
                    {

                        //Excel
                        //GetDateForExcel(stritem, rad_E.Checked, parm);


                        // 重置分頁狀態
                        currentPage = 1;

                        // 使用分頁功能顯示數據
                        DataTable pagedData = GetItemPositionPagedData(dt, currentPage);
                        dataGridView1.DataSource = pagedData;

                        // 更新頁碼顯示
                        if (lblCurrentPage != null)
                        {
                            lblCurrentPage.Text = $"第 {currentPage} 頁";
                            lblCurrentPage.Visible = true;
                        }

                        // 顯示分頁按鈕
                        btnPreviousPage.Visible = true;
                        btnNextPage.Visible = true;

                        sumC();
                    }
                    else
                    {
                        MessageBox.Show("查無資料，請確認輸入後重新查詢");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("料號及儲位資料不正確，請確認輸入後重新查詢");
            }
            btn_itemP = 0;
        }

        #endregion

        /// <summary>
        /// 匯出Excel-資料
        /// </summary>
        /// <returns></returns>
        private DataTable dataBind_Old()
        {
            string sql = string.Empty;
            //int offset = (pageNumber - 1) * pageSize;

            sql = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno    
                        from Automatic_Storage_Detail a
                        left join Automatic_Storage_Package b on a.Package=b.code
                        left join Automatic_Storage_User c on a.Input_UserNo=c.User_No 
                        left join Automatic_Storage_User d on a.Output_UserNo=d.User_No 
                        where a.Unit_No=@unitNo and amount >0 order by Actual_InDate";
            if (btn_fAll == 1)
            {
                sql = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,Package,Up_InDate,Input_UserNo,Up_OutDate,Output_UserNo,Mark,PCB_DC,CMC_DC,sno from Automatic_Storage_Detail 
                        where Unit_No=@unitNo order by Actual_InDate";
            }

            SqlParameter[] parm = new SqlParameter[]
            {
                new SqlParameter("unitNo",Login.Unit_No),
            };

            dt = db.ExecuteDataTable(sql, CommandType.Text, parm);

            sumC();

            btn_fAll = 0;

            // 不再直接指定到UI控制項
            return dt;
        }

        /// <summary>
        /// 日期解析
        /// </summary>
        /// <param name="inputString"></param>
        /// <returns></returns>
        public bool Week_Then25(string inputString)
        {
            // 假設輸入字串為 inputString
            //string inputString = textBox1.Text;

            DateTime inputDate;
            // 嘗試解析 YY/WW 格式
            if (inputString.Length == 5)
            {
                string[] y_w = inputString.Split('/');

                // 取得年份和週數
                int year = int.Parse(y_w[0]) + 2000; // 將兩位數年份轉換為四位數年份
                int week = int.Parse(y_w[1]);

                // 轉換為 DateTime 型別
                inputDate = new DateTime(year, 1, 1).AddDays((week - 1) * 7);

                // 取得指定年份的 1 月 1 日
                DateTime jan1 = new DateTime(year, 1, 1);

                // 取得指定年份的 1 月 1 日是星期幾
                int dayOfWeek = (int)jan1.DayOfWeek;

                // 計算指定週數的第一天
                int daysToAdd = (week - 1) * 7;
                daysToAdd -= dayOfWeek;
                inputDate = jan1.AddDays(daysToAdd);
            }
            // 嘗試解析 YYWW 格式
            else if (inputString.Length == 4 && int.TryParse(inputString, out int yearWeek))
            {
                int year = 2000 + (yearWeek / 100);
                int week = yearWeek % 100;
                inputDate = new DateTime(year, 1, 1).AddDays((week - 1) * 7);
            }
            // 嘗試解析 YYYYMMDD 格式
            else if (DateTime.TryParseExact(inputString, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out inputDate))
            {
                // 不需要進行額外的處理
            }
            // 嘗試解析 YYYY/MM/DD 格式
            else if (DateTime.TryParseExact(inputString, "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out inputDate))
            {
                // 不需要進行額外的處理
            }
            //else
            //{
            //    label1.Text = ("無法解析輸入字串");
            //    return ;
            //}

            // 取得當天日期
            DateTime currentDate = DateTime.Today;

            // 計算當天日期與輸入日期的時間差
            TimeSpan diff = currentDate - inputDate;

            // 取得時間差的絕對值
            TimeSpan diffAbs = diff.Duration();

            // 計算時間差的週數
            int weeks = (int)Math.Floor(diffAbs.TotalDays / 7);

            // 判斷週數是否超過 25 週
            if (weeks > 25)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 總數量
        /// </summary>
        private void sumC()
        {
            //計算總數
            List<SqlParameter> parmC = new List<SqlParameter>();
            string sumSql = string.Empty;


            // 判斷當前是哪種查詢模式
            bool isItemQuery = !string.IsNullOrWhiteSpace(txt_item.Text); // 料號
            bool isPositionQuery = !string.IsNullOrWhiteSpace(txt_position.Text); // 儲位
            bool isDateQuery = !string.IsNullOrWhiteSpace(txt_DTs.Text) && !string.IsNullOrWhiteSpace(txt_DTe.Text);//日期

            // 判斷歷史查詢
            bool h_Flag = false;
            if (txt_itemP5.Text != "" && txt_itemP5.Text.Length >= 3) // 料號
            {
                h_Flag = true;
            }
            else if (txt_siteP5.Text != "") // 儲位
            {
                h_Flag = true;
            }


            // 料號+儲位查詢模式 - 歷史查詢
            if (h_Flag)
            {
                // 基本查詢條件
                string whereClause = " where Unit_No=@unitNo and cast(Input_Date as date) >'2023-06-30' ";
                string outputWhereClause = " where Unit_No=@unitNo and cast(Output_Date as date) >'2023-06-30' ";

                parmC.Add(new SqlParameter("unitNo", Login.Unit_No));

                // 添加料號條件 - 歷史查詢
                if (!string.IsNullOrEmpty(txt_itemP5.Text))
                {
                    if (rad_E.Checked) //昶
                    {
                        whereClause += " and item_no_master = @master ";
                        outputWhereClause += @" and item_no_master = @master ";
                    }
                    else if (rad_C.Checked)//客
                    {
                        whereClause += " and Item_No_Slave = @master ";
                        outputWhereClause += @" and  Item_No_Slave = @master ";

                    }
                    parmC.Add(new SqlParameter("master", txt_itemP5.Text.Trim()));
                }

                // 添加儲位條件 - 歷史查詢
                if (!string.IsNullOrEmpty(txt_siteP5.Text))
                {
                    whereClause += " and Position = @position ";
                    outputWhereClause += " and Position = @position ";
                    parmC.Add(new SqlParameter("position", txt_siteP5.Text.Trim()));
                }


                // 計算入庫總量減去出庫總量
                sumSql = $@"select (select ISNULL(sum(Amount), 0) from Automatic_Storage_Input {whereClause}) - 
                            (select ISNULL(sum(Amount), 0) from Automatic_Storage_Output {outputWhereClause}) as C";
            }
            else if (isItemQuery && isPositionQuery) // 料號+儲位查詢模式 - 一般
            {
                sumSql = @"select sum(Amount) as C from Automatic_Storage_Detail ";
                if (rad_E.Checked) //昶
                {
                    sumSql += @" where Unit_No=@unitNo and  Position=@position and  Item_No_Master = @master and amount >0 ";
                    parmC.Add(new SqlParameter("unitNo", Login.Unit_No));
                    parmC.Add(new SqlParameter("position", txt_position.Text));
                    parmC.Add(new SqlParameter("master", txt_item.Text.Trim()));
                }
                else if (rad_C.Checked)//客
                {
                    sumSql += @" where Unit_No=@unitNo and  Position=@position and  Item_No_Slave = @master and amount >0 ";
                    parmC.Add(new SqlParameter("unitNo", Login.Unit_No));
                    parmC.Add(new SqlParameter("position", txt_position.Text));
                    parmC.Add(new SqlParameter("master", txt_item.Text.Trim()));
                }
                if (btn_itemP == 1)// 歷史查詢全部
                {
                    sumSql = @" select sum(Amount) as C from Automatic_Storage_Detail a where a.Unit_No=@unitNo and  Position=@position and Item_No_Master = @master";

                    parmC.Add(new SqlParameter("unitNo", Login.Unit_No));
                    parmC.Add(new SqlParameter("master", txt_item.Text.Trim()));
                    parmC.Add(new SqlParameter("position", txt_position.Text));

                }
            }
            // 一般查詢模式
            else
            {
                // 基本查詢
                sumSql = "select sum(Amount) as C from Automatic_Storage_Detail where Unit_No=@unitNo ";
                parmC.Add(new SqlParameter("unitNo", Login.Unit_No));

                // 添加儲位條件
                if (isPositionQuery)
                {
                    //是否歷史資料
                    if (is_History_Query)
                    {
                        sumSql += @" AND Up_InDate > '2023-06-30'";
                    }
                    else
                    {
                        sumSql += @" AND amount > 0";
                    }

                    sumSql += @" and Position like @position+'%' ";
                    parmC.Add(new SqlParameter("position", txt_position.Text));
                }

                // 添加料號條件
                if (isItemQuery)
                {
                    //是否歷史資料
                    #region
                    if (is_History_Query)
                    {
                        if (is_Rad_E)
                        {
                            sumSql = @"SELECT SUM(Amount) as C
                                    FROM
                                    (
                                          SELECT a.sno FROM Automatic_Storage_Input a
                                          WHERE a.Unit_No = @unitNo AND Item_No_Master = @master AND cast(a.Input_Date as date) > '2023-06-30'
                                          UNION
                                          SELECT a.sno FROM Automatic_Storage_Output a
                                          WHERE a.Unit_No = @unitNo AND Item_No_Master = @master AND cast(a.Output_Date as date) > '2023-06-30'
                                    ) AS t";

                        }
                        else
                        {

                            sumSql = @"SELECT SUM(Amount) as C
                                   FROM 
                                    (
                                        SELECT a.sno FROM Automatic_Storage_Input a 
                                        WHERE a.Unit_No=@unitNo AND Item_No_Slave=@master AND cast(a.Input_Date as date) >'2023-06-30'
                                        UNION
                                        SELECT a.sno FROM Automatic_Storage_Output a 
                                        WHERE a.Unit_No=@unitNo AND Item_No_Slave=@master AND cast(a.Output_Date as date) >'2023-06-30'
                                    ) AS t";
                        }
                    }
                    #endregion
                    else
                    {
                        if (is_Rad_E)
                        {
                            sumSql = $@"SELECT SUM(Amount) as C 
                            FROM Automatic_Storage_Detail                           
                            WHERE Unit_No=@unitNo AND Item_No_Master LIKE '%'+@master+'%' AND amount > 0 ";
                        }
                        else
                        {
                            sumSql = $@"SELECT SUM(Amount) as C 
                            FROM Automatic_Storage_Detail                             
                            WHERE Unit_No=@unitNo AND Item_No_Slave LIKE '%'+@master+'%' AND amount > 0 ";
                        }
                    }


                    // parmC.Add(new SqlParameter("unitNo", Login.Unit_No));
                    parmC.Add(new SqlParameter("master", txt_item.Text.Trim()));
                }

                // 添加日期條件
                if (isDateQuery)
                {
                    DateTime sdate = DateTime.ParseExact(txt_DTs.Text, "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                    DateTime edate = DateTime.ParseExact(txt_DTe.Text, "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                    string s_date = sdate.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    string e_date = edate.ToString("yyyy-MM-dd HH:mm:59.999");

                    if (rad_in.Checked)
                    {
                        sumSql += " and amount > 0 and Up_InDate between @startDate and @endDate ";
                    }
                    else if (rad_out.Checked)
                    {
                        sumSql += " and amount > 0 and Up_OutDate between @startDate and @endDate ";
                    }
                    parmC.Add(new SqlParameter("startDate", s_date.Trim()));
                    parmC.Add(new SqlParameter("endDate", e_date.Trim()));
                }
            }

            // 執行查詢並顯示結果
            try
            {
                DataSet ds = db.ExecuteDataSetPmsList(sumSql, CommandType.Text, parmC);
                if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0 && ds.Tables[0].Rows[0]["C"] != DBNull.Value)
                {
                    txt_sumC.Text = ds.Tables[0].Rows[0]["C"].ToString();
                }
                else
                {
                    txt_sumC.Text = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("計算總數量時發生錯誤: " + ex.Message);
                txt_sumC.Text = "0";
            }
        }


        #region
        /// <summary>
        /// 查詢數據並匯出到Excel
        /// </summary>
        /// <param name="e_Flag">true:昶 / false:客</param>
        /// <param name="pms"></param>
        private void GetDateForExcel(string strSql1, bool e_Flag, params SqlParameter[] pms)
        {
            string strSql = strSql1;
            switch (queryType)
            {
                case 0:
                    if (e_Flag)
                    {
                        strSql = $@"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name as User_name1,a.Mark,PCB_DC,CMC_DC,a.sno  
                            FROM Automatic_Storage_Detail a 
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                            WHERE a.Unit_No=@unitNo AND Item_No_Master LIKE '%'+@master+'%' AND amount > 0 
                            ORDER BY a.position";
                    }
                    else
                    {
                        strSql = $@"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name as User_name1,a.Mark,PCB_DC,CMC_DC,a.sno  
                            FROM Automatic_Storage_Detail a 
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                            WHERE a.Unit_No=@unitNo AND Item_No_Slave LIKE '%'+@master+'%' AND amount > 0 
                            ORDER BY a.position";
                    }

                    break;
                case 1:
                    strSql = $@"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                        FROM Automatic_Storage_Detail a 
                        LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                        LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                        LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                        WHERE a.Unit_No=@unitNo AND Position LIKE @position+'%' AND amount > 0 
                        ORDER BY a.sno";

                    break;
                case 2:
                    if (e_Flag)
                    {
                        strSql = $@"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name as User_name1,a.Mark,PCB_DC,CMC_DC,a.sno  
                            FROM Automatic_Storage_Detail a 
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                            WHERE a.Unit_No=@unitNo AND Item_No_Master LIKE '%'+@master+'%' AND amount > 0 
                            ORDER BY a.position";
                    }
                    else
                    {
                        strSql = $@"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name as User_name1,a.Mark,PCB_DC,CMC_DC,a.sno  
                            FROM Automatic_Storage_Detail a 
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                            WHERE a.Unit_No=@unitNo AND Item_No_Slave LIKE '%'+@master+'%' AND amount > 0 
                            ORDER BY a.position";
                    }

                    break;
                case 3:
                    if (!string.IsNullOrWhiteSpace(strSql1)) strSql = strSql1;
                    break;

            }

            // 創建新的參數集合，避免重複使用相同的參數實例
            SqlParameter[] newParams = null;
            if (pms != null && pms.Length > 0)
            {
                newParams = new SqlParameter[pms.Length];
                for (int i = 0; i < pms.Length; i++)
                {
                    // 為每個參數創建一個新的實例
                    newParams[i] = new SqlParameter(pms[i].ParameterName, pms[i].Value);
                    if (pms[i].SqlDbType != System.Data.SqlDbType.NVarChar)
                    {
                        newParams[i].SqlDbType = pms[i].SqlDbType;
                    }
                }
            }

            dt = db.ExecuteDataTable(strSql, CommandType.Text, newParams);
        }
        #endregion

        /// <summary>
        /// 入庫登記
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Input input = new Input();
                input.Owner = this;
                input.Show();
                GetLog.WriteLog(((Button)(sender)).Name);
            }
            catch (Exception ee)
            {

                throw;
            }
        }

        /// <summary>
        /// DataGridView 表頭欄位資料設定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView1.Columns[0].HeaderText = "入庫日期";                //Actul_Date
            dataGridView1.Columns[0].DefaultCellStyle.Format = "yyyy-MM-dd";                //Actul_Date
            dataGridView1.Columns[1].HeaderText = "料號_昶";                //Item_No_Master
            dataGridView1.Columns[2].HeaderText = "料號_客";                //Item_No_Master
            dataGridView1.Columns[3].HeaderText = "規格";                //Spec
            dataGridView1.Columns[4].HeaderText = "儲位";                //Position
            dataGridView1.Columns[5].HeaderText = "總數";                //Amount
            dataGridView1.Columns[6].HeaderText = "包裝型態";       //Package
            dataGridView1.Columns[7].HeaderText = "操作日期";
            dataGridView1.Columns[7].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";                //Actul_Date
            dataGridView1.Columns[8].HeaderText = "操作人員";
            dataGridView1.Columns[9].HeaderText = "出庫日期";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";                //Actul_Date
            dataGridView1.Columns[10].HeaderText = "出庫人員";
            dataGridView1.Columns[11].HeaderText = "備註";
            dataGridView1.Columns[12].HeaderText = "PCB";
            dataGridView1.Columns[13].HeaderText = "CMC";
            dataGridView1.Columns[14].Visible = false;
        }

        /// <summary>
        /// 重新整理資料，呼叫 dataBind 方法以更新畫面顯示
        /// </summary>
        /// <remarks>
        /// 此方法可用於外部觸發資料重新整理，例如查詢條件變更或操作完成後。
        /// </remarks>
        /// <example>
        /// <code>
        /// form1.refreshData();
        /// </code>
        /// </example>
        public void refreshData()
        {
            // 呼叫 dataBind 方法，更新資料顯示
            dataBind();
        }

        /// <summary>
        /// 出庫登記
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            OutPut output = new OutPut();
            output.Owner = this;
            output.Show();
            dataBind();
        }

        /// <summary>
        /// 檔案選擇對話框，提供使用者選擇 Excel 檔案進行批次入庫
        /// </summary>
        private OpenFileDialog fileDialog1; // 檔案選擇對話框

        /// <summary>
        /// 批次入庫
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel1.BringToFront();
            dispatcherTimer.Stop();
            fileDialog1 = new OpenFileDialog()
            {
                FileName = "Select a text file",
                Filter = "Text files (*.xls)|*.xls",
                Title = "Open text file"
            };

            //selectButton = new Button()
            //{
            //    Size = new Size(87,37),
            //    Location = new Point(270, 17),
            //    Font = new Font("微軟正黑體", 12, FontStyle.Bold),
            //    Text = "選擇檔案"
            //};

            //selectButton.Click += new EventHandler(selectButton_Click);
            //commitButton.Click += new EventHandler(commitButton_Click);
            //panel1.Controls.Add(selectButton);

        }

        /// <summary>
        /// 選擇檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void selectButton_Click(object sender, EventArgs e)
        {
            if (fileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var fileName = fileDialog1.FileName;

                    FileInfo _info = new FileInfo(fileName);
                    string _new = Application.StartupPath + "\\Upload\\" + _info.Name;
                    if (File.Exists(fileName))
                    {
                        _info.CopyTo(_new, true);
                        txt_path.Text = _new;
                    }
                    dataBind();
                }
                catch (SecurityException ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }

        /// <summary>
        /// 上傳入庫
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void commitButton_Click(object sender, EventArgs e)
        {
            try
            {
                commitButton.Enabled = false;
                string fileName = txt_path.Text;
                LoadExcel(fileName);
                //WriteExcelData();
                list_result.Items.Clear();
                this.inputExcelBW.WorkerSupportsCancellation = true; //允許中斷
                this.inputExcelBW.RunWorkerAsync(); //呼叫背景程式
                dataBind();

            }
            catch (SecurityException ex)
            {
                MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }

        }



        /// <summary>
        /// 清空
        /// </summary>
        private void clsAll()
        {
            txt_item.Text = ""; //料號查詢
            txt_position.Text = "";// 儲位查詢
        }

        /// <summary>
        /// 計時器
        /// </summary>
        public void TimerStart()
        {
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 3);
            dispatcherTimer.Start();
        }

        /// <summary>
        /// 不顯示panel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            panel1.Visible = false;
        }

        /// <summary>
        /// 讀取 Excel 檔案
        /// </summary>
        /// <param name="filename"></param>
        private void LoadExcel(string filename)
        {
            if (filename != "")
            {
                #region office 97-2003
                string excelString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                    "Data Source=" + filename + ";" +
                                    "Extended Properties='Excel 8.0;HDR=Yes;IMEX=1\'";
                #endregion

                #region office 2007
                //string excelString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename +
                //                 ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1';";
                #endregion

                OleDbConnection cnn = new OleDbConnection(excelString);
                cnn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = cnn;
                OleDbDataAdapter adapter = new OleDbDataAdapter();
                cmd.CommandText = "SELECT * FROM [Sheet1$] where (入庫日期<>'' or 入庫日期 is not null) ";
                adapter.SelectCommand = cmd;
                //DataSet dsData = new DataSet();
                adapter.Fill(dsData);
                cnn = null;
                cmd = null;
                adapter = null;
            }
            else
            {
                MessageBox.Show("請選擇檔案");
            }
        }

        /// <summary>
        /// 背景工作執行緒，提供非同步處理功能
        /// </summary>
        /// <remarks>
        /// 可用於執行長時間運算或 IO 操作，避免 UI 阻塞
        /// </remarks>
        /// <example>
        /// <code>
        /// backgroundWorker.DoWork += BackgroundWorker_DoWork;
        /// backgroundWorker.RunWorkerAsync();
        /// </code>
        /// </example>
        BackgroundWorker backgroundWorker = new BackgroundWorker(); // 建立 BackgroundWorker 物件，用於非同步背景作業

        /// <summary>
        /// 批次寫入 Excel 資料到資料庫，包含資料驗證、異常處理與進度條更新
        /// </summary>
        public void WriteExcelData()
        {
            Form.CheckForIllegalCrossThreadCalls = false; // 關閉跨執行緒檢查，避免 UI 執行緒錯誤
            try
            {
                // 宣告所有欄位變數
                string ItemE = "", ItemC = "", SIDE = "", Package = "", Mark = "", Spec = "", ActualDate = "", PCB = "", CMC = "";
                string AMOUNT = "";
                string sqlinput = string.Empty; // 入庫 SQL
                string sqldetail = string.Empty; // 明細 SQL
                string sqlchkFirst = string.Empty; // 檢查是否已存在 SQL
                string strsql_spec = string.Empty; // 規格查詢 SQL
                string search_spec = string.Empty; // 規格查詢條件

                progressBar1.Minimum = 0; // 設定進度條最小值
                progressBar1.Maximum = dsData.Tables[0].Rows.Count; // 設定進度條最大值
                progressBar1.Step = 1; // 設定進度條每次步進值

                // 逐筆處理 Excel 資料
                for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
                {
                    string strErrMsg = string.Empty; // 錯誤訊息
                    ActualDate = Convert.ToDateTime(dsData.Tables[0].Rows[i][0].ToString().Trim()).ToString("yyyy-MM-dd"); // 入庫日期
                    ItemE = (dsData.Tables[0].Rows[i][1].ToString().Trim().ToUpper()); // 昶料號
                    ItemC = (dsData.Tables[0].Rows[i][2].ToString().Trim().ToUpper()); // 客料號
                    AMOUNT = dsData.Tables[0].Rows[i][3].ToString().Trim(); // 數量
                    Package = dsData.Tables[0].Rows[i][4].ToString().Trim().ToUpper(); // 包裝型態
                    SIDE = dsData.Tables[0].Rows[i][5].ToString().Trim().ToUpper(); // 儲位
                    PCB = dsData.Tables[0].Rows[i][6].ToString().Trim().ToUpper(); // PCB
                    Mark = dsData.Tables[0].Rows[i][7].ToString().Trim().ToUpper(); // 備註
                    CMC = dsData.Tables[0].Rows[i][8].ToString().Trim().ToUpper(); // CMC

                    // 檢查儲位是否已設定
                    string strsql_p = @"select * from Automatic_Storage_Position where Position=@position and Unit_no=@unitno";
                    SqlParameter[] parm_p = new SqlParameter[]
                    {
                new SqlParameter("position",SIDE),
                new SqlParameter("unitno",Login.Unit_No)
                    };
                    DataSet dsP = db.ExecuteDataSet(strsql_p, CommandType.Text, parm_p); // 查詢儲位
                    strErrMsg = (dsP.Tables[0].Rows.Count == 0) ? " ,x儲位 " : ""; // 若儲位不存在，記錄錯誤

                    // 檢查規格是否已設定
                    if (!string.IsNullOrEmpty(ItemE))
                    {
                        search_spec = ItemE; // 以昶料號查詢
                        strsql_spec = @"select Item_E,item_C,spec from Automatic_Storage_Spec where Unit_no=@unitno and (Item_E =@search_spec )";
                    }
                    else if (!string.IsNullOrEmpty(ItemC))
                    {
                        search_spec = ItemC; // 以客料號查詢
                        strsql_spec = @"select Item_E,item_C,spec from Automatic_Storage_Spec where Unit_no=@unitno and (item_C = @search_spec )";
                    }

                    SqlParameter[] parm_spec = new SqlParameter[]
                    {
                new SqlParameter("unitno",Login.Unit_No),
                new SqlParameter("search_spec",search_spec)
                    };
                    DataSet ds_Spec = db.ExecuteDataSet(strsql_spec, CommandType.Text, parm_spec); // 查詢規格
                    if (ds_Spec.Tables[0].Rows.Count > 0)
                    {
                        ItemE = ds_Spec.Tables[0].Rows[0]["Item_E"].ToString().Trim(); // 取得昶料號
                        ItemC = ds_Spec.Tables[0].Rows[0]["item_C"].ToString().Trim(); // 取得客料號
                        Spec = ds_Spec.Tables[0].Rows[0]["Spec"].ToString().Trim(); // 取得規格
                    }
                    else
                    {
                        ItemE = (dsData.Tables[0].Rows[i][1].ToString().Trim().ToUpper()); // 若查無規格，保留原始資料
                        ItemC = (dsData.Tables[0].Rows[i][2].ToString().Trim().ToUpper());
                        Spec = "";
                        strErrMsg += " ,x規格 "; // 規格錯誤
                    }

                    // 查詢包裝型態是否存在
                    string strsql_pack = @"select Code from Automatic_Storage_Package where substring(Package_View,1,1) =@pack ";
                    SqlParameter[] parm_pack = new SqlParameter[]
                    {
                new SqlParameter("pack",Package)
                    };
                    DataSet dsPack = db.ExecuteDataSet(strsql_pack, CommandType.Text, parm_pack); // 查詢包裝
                    if (dsPack.Tables[0].Rows.Count > 0)
                    {
                        Package = dsPack.Tables[0].Rows[0]["Code"].ToString().Trim(); // 取得包裝代碼
                    }
                    else
                    {
                        strErrMsg += " ,x包裝 "; // 包裝錯誤
                    }

                    // 檢查入庫日期是否為未來日期
                    DateTime ACdt = Convert.ToDateTime(ActualDate); // 轉換入庫日期
                    TimeSpan timeSpan = ACdt.Subtract(DateTime.Today); // 計算與今天的差距
                    if (!string.IsNullOrEmpty(ActualDate))
                    {
                        if (timeSpan.TotalDays >= 1)
                        {
                            strErrMsg += " ,x日期 "; // 日期錯誤
                        }
                    }

                    // 20220323 檢查儲位、包裝、規格是否正確
                    if (dsP.Tables[0].Rows.Count == 0 || string.IsNullOrEmpty(Package) || string.IsNullOrEmpty(Spec))
                    {
                        txt_result.Text += ("第" + (i + 2) + "筆失敗 " + strErrMsg).ToString() + Environment.NewLine; // 顯示失敗訊息
                    }
                    else
                    {
                        try
                        {
                            // 寫入入庫資料
                            sqlinput = @"begin try
                                begin tran
                                INSERT INTO Automatic_Storage_Input
                                (Position,Amount,Item_No_Master,Item_No_Slave,Spec,
                                Unit_No,Package,Input_Date,Actual_InDate,Input_UserNo,Mark,PCB_DC,CMC_DC)   
                                values (@Position,@Amount,@Master,@Slave,@Spec,
                                @Unit_No,@package,@InDate,@ActualDate,@UserNo,@mark,@PCB,@CMC)
                                commit tran
                             end try
                             begin catch
                                rollback tran
                             end catch";
                            SqlParameter[] parm = new SqlParameter[]
                            {
                        new SqlParameter("Position",SIDE.Trim().ToUpper()), // 儲位
                        new SqlParameter("Amount",AMOUNT), // 數量
                        new SqlParameter("Master",ItemE), // 昶料號
                        new SqlParameter("Slave",ItemC), // 客料號
                        new SqlParameter("Spec",Spec), // 規格
                        new SqlParameter("Package",Package), // 包裝
                        new SqlParameter("Unit_No",Login.Unit_No), // 單位編號
                        new SqlParameter("InDate",DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")), // 入庫時間
                        new SqlParameter("ActualDate",ActualDate), // 實際日期
                        new SqlParameter("UserNo",Login.User_No), // 使用者編號
                        new SqlParameter("mark",Mark), // 備註
                        new SqlParameter("PCB",PCB), // PCB
                        new SqlParameter("CMC",CMC) // CMC
                            };
                            db.ExecueNonQuery(sqlinput, CommandType.Text, "Form_btn_BatIn", parm); // 執行入庫 SQL

                            //key : 入庫日期+料號+位置+包裝種類, if (key=新資料)  insert , else update
                            // 檢查明細資料是否已存在
                            sqlchkFirst = @"select * from Automatic_Storage_Detail 
                            where Item_No_Master=@Master and  Item_No_Slave=@Slave and Position=@Position and 
                            Package=@Package and Actual_InDate=@actualDate and Mark = @mark and PCB_DC = @PCB and CMC_DC =@CMC and amount >0 ";
                            SqlParameter[] parameters = new SqlParameter[]
                            {
                    new SqlParameter("Master",ItemE), // 昶料號
                    new SqlParameter("Slave",ItemC), // 客料號
                    new SqlParameter("Position",SIDE.Trim()), // 儲位
                    new SqlParameter("Package",Package.Trim()), // 包裝
                    new SqlParameter("actualDate",ActualDate), // 實際日期
                    new SqlParameter("mark",Mark), // 備註
                    new SqlParameter("PCB",PCB), // PCB
                    new SqlParameter("CMC",CMC) // CMC
                            };
                            DataSet dataSet = db.ExecuteDataSet(sqlchkFirst, CommandType.Text, parameters); // 查詢明細

                            if (dataSet.Tables[0].Rows.Count > 0)
                            {
                                string detailSno = dataSet.Tables[0].Rows[0]["Sno"].ToString(); // 取得明細編號
                                                                                                // 若資料已存在，更新數量+1與相關欄位+1
                                sqldetail = @"
                            begin try
                                begin tran
                                update Automatic_Storage_Detail set Amount=Amount+@Amount, Mark=@Mark, PCB_DC=@PCB ,CMC_DC=@CMC 
                                where Sno = @sno
                                commit tran
                             end try
                             begin catch
                                rollback tran
                             end catch";
                                SqlParameter[] parm2 = new SqlParameter[]
                                {
                        new SqlParameter("Amount",AMOUNT), // 數量
                        new SqlParameter("Mark",Mark), // 備註
                        new SqlParameter("PCB",PCB), // PCB
                        new SqlParameter("CMC",CMC), // CMC
                        new SqlParameter("Sno",detailSno) // 明細編號
                                };
                                db.ExecueNonQuery(sqldetail, CommandType.Text, "Form_btn_BatIn", parm2); // 執行更新 SQL
                            }
                            else
                            {
                                // 若為新資料，寫入明細資料表
                                sqldetail = @"
                            begin try
                                begin tran
                                INSERT INTO Automatic_Storage_Detail
                                  (Sno,Item_No_Master,Item_No_Slave,Spec,Unit_No,Position,Up_InDate,Actual_InDate,Input_UserNo,
                                    Amount,Package,Mark,PCB_DC,CMC_DC)	 
                                      SELECT top(1)Sno,Item_No_Master,Item_No_Slave,Spec,Unit_No,Position,GETDATE(),Actual_InDate,
                                    Input_UserNo,Amount,Package,Mark,PCB_DC,CMC_DC
                                  FROM Automatic_Storage_Input
                                      where Item_No_Master= @Master and Item_No_Slave = @Slave and Unit_No=@Unit_No 
                                  and Position=@position and Package=@Package and Actual_InDate=@actualDate and Spec=@Spec order by Input_Date desc
                                commit tran
                             end try
                             begin catch
                                rollback tran
                             end catch";
                                SqlParameter[] parm2 = new SqlParameter[]
                                {
                        new SqlParameter("Master",ItemE), // 昶料號
                        new SqlParameter("Slave",ItemC), // 客料號
                        new SqlParameter("Unit_No",Login.Unit_No), // 單位編號
                        new SqlParameter("position",SIDE.Trim()), // 儲位
                        new SqlParameter("Package",Package.Trim()), // 包裝
                        new SqlParameter("actualDate",ActualDate), // 實際日期
                        new SqlParameter("Spec",Spec), // 規格
                                };
                                db.ExecueNonQuery(sqldetail, CommandType.Text, "Form_btn_BatIn", parm2); // 執行新增 SQL
                            }
                            progressBar1.PerformStep(); // 進度條前進
                        }
                        catch (Exception ex)
                        {
                            list_result.Items.Add(ex.Message); // 顯示例外訊息
                        }
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("請確認資料"); // 顯示資料錯誤訊息
            }
        }

        /// <summary>
        /// 批次出庫
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, EventArgs e)
        {
            BatOut bO = new BatOut();
            bO.Owner = this;
            bO.ShowDialog();
        }

        /// <summary>
        /// 管理設定
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Maintain_Click(object sender, EventArgs e)
        {
            Maintain m = new Maintain();
            m.Owner = this;
            m.ShowDialog();
            dataBind();
        }

        /// <summary>
        /// 備料單匯入按鈕事件，開啟匯入視窗
        /// </summary>
        private void btn備料單匯入_Click(object sender, EventArgs e)
        {
            try
            {
                // 建立或重用單一實例
                if (importForm is null || importForm.IsDisposed)
                {
                    importForm = new Form備料單匯入();
                    importForm.Owner = this;
                    importForm.FormClosed += (s, args) =>
                    {
                        // 使用者按 X 關閉才釋放，下次再開會重置內容
                        importForm = null;
                        try { this.Activate(); } catch { }
                    };
                }

                // 顯示並最大化
                if (!importForm.Visible)
                {
                    try { importForm.WindowState = FormWindowState.Maximized; } catch { }
                    importForm.Show();
                }
                else
                {
                    try
                    {
                        if (importForm.WindowState == FormWindowState.Minimized)
                            importForm.WindowState = FormWindowState.Maximized;
                        else
                            importForm.WindowState = FormWindowState.Maximized;
                    }
                    catch { }
                    importForm.BringToFront();
                    importForm.Activate();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("開啟備料單匯入視窗失敗: " + ex.Message);
            }
        }

        /// <summary>
        /// 整批刪除
        /// </summary>
        private void InvokeDeleteThread()
        {
            DialogResult res = MessageBox.Show(
                           "你確認刪除嗎??" + Environment.NewLine +
                           "( 料號 : " + txt_item.Text.Trim() + Environment.NewLine +
                           "庫位 : " + txt_position.Text.Trim() + " ) ",
                                           "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res == DialogResult.Yes)
            {
                progressBar2.Visible = Enabled;
                string item = string.Empty;
                string posi = string.Empty;
                string spec = string.Empty;
                string package = string.Empty;

                string time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                try
                {
                    if (string.IsNullOrEmpty(txt_position.Text))
                    {
                        MessageBox.Show("尚未指定庫位");
                        return;
                    }
                    string stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,package,Up_InDate,Input_UserNo,Up_OutDate,Output_UserNo " +
                    "from Automatic_Storage_Detail where Unit_No=@unitNo and  Position like @Position+'%' and amount >0 " +
                    "order by Item_No_Master";
                    SqlParameter[] parm = new SqlParameter[]
                    {
                        new SqlParameter("unitNo",Login.Unit_No),
                        new SqlParameter("Position",txt_position.Text.Trim())
                    };
                    DataSet ds_p = db.ExecuteDataSet(stritem, CommandType.Text, parm);
                    progressBar2.Minimum = 0;
                    progressBar2.Maximum = ds_p.Tables[0].Rows.Count;
                    progressBar2.Step = 1;
                    foreach (DataRow row in ds_p.Tables[0].Rows)
                    {
                        item = row["Item_No_Master"].ToString();
                        posi = row["Position"].ToString();
                        spec = row["Spec"].ToString();
                        package = row["Package"].ToString();
                        string sqlOut = @"insert into Automatic_Storage_Output (Sno,Item_No_Master,Item_No_Slave,spec,Position,Amount,package,Unit_No,Output_UserNo,Output_Date) 
                                                   select sno,Item_No_Master,Item_No_Slave,spec,Position,Amount,package,@unitno,@outuser,@outdate from Automatic_Storage_Detail 
                                                   where spec=@spec and Position=@position and package=@package and Item_No_Master=@master_d and amount > 0";
                        SqlParameter[] parm1 = new SqlParameter[]
                        {
                        new SqlParameter("unitno",Login.Unit_No),
                        new SqlParameter("outuser",Login.User_No),
                        new SqlParameter("outdate",time),
                        new SqlParameter("master",item.Trim()),
                        new SqlParameter("spec",spec.Trim()),
                        new SqlParameter("position",posi.Trim()),
                        new SqlParameter("package",package.Trim()),
                        new SqlParameter("master_d",item.Trim())
                        };
                        db.ExecueNonQuery(sqlOut, CommandType.Text, "btndelPosition", parm1);

                        //Update Storage_Detail
                        string sqlDetail = @"update Automatic_Storage_Detail 
                                                           set Amount = 0 ,Amount_Unit = 0,Up_OutDate = @outdate , Output_UserNo = @outuser 
                                                           where Item_No_Master= @master and Position=@position";
                        SqlParameter[] parm2 = new SqlParameter[]
                        {
                        new SqlParameter("outdate",time),
                        new SqlParameter("outuser",Login.User_No),
                        new SqlParameter("master",item.Trim()),
                        new SqlParameter("position",posi.Trim()),
                        };
                        db.ExecueNonQuery(sqlDetail, CommandType.Text, "btndelPosition", parm2);
                        progressBar2.PerformStep();
                    }
                    dataBind();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btndelPosition_Click(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(InvokeDeleteThread));
            }
            else
            {
                InvokeDeleteThread();
            }
            progressBar2.Visible = false;
        }

        /// <summary>
        /// 初始化權限設定，根據使用者角色啟用或顯示對應功能按鈕
        /// </summary>
        /// <remarks>
        /// 此方法會根據目前登入使用者的角色，設定各功能按鈕的啟用或顯示狀態。
        /// 角色對應如下：
        /// 0 - 管理者：顯示刪除按鈕
        /// 1 - 出庫作業：啟用出庫按鈕
        /// 2 - 入庫作業：啟用入庫及批次入庫按鈕
        /// 3 - 資料維護：顯示維護按鈕
        /// 4 - 批次出庫(儲位)：啟用批次出庫按鈕
        /// </remarks>
        /// <example>
        /// <code>
        /// Initi(); // 執行初始化權限設定
        /// </code>
        /// </example>
        private void Initi()
        {
            // 預設所有按鈕關閉或隱藏
            btn_Input.Enabled = false; // 入庫按鈕關閉
            btn_BatIn.Enabled = false; // 批次入庫按鈕關閉
            btn_Out.Enabled = false; // 出庫按鈕關閉
            btn_BatOut.Enabled = false; // 批次出庫按鈕關閉
            btn_Maintain.Visible = false; // 維護按鈕隱藏
            btn_delPosition.Visible = false; // 刪除按鈕隱藏

            // 查詢使用者角色
            string sql_initi = @"select * from Automatic_Storage_UserRole where USER_ID=@userid";
            SqlParameter[] parm_initi = new SqlParameter[]
            {
            new SqlParameter("userid",Login.User_No) // 使用者編號參數
            };
            DataSet ds = db.ExecuteDataSet(sql_initi, CommandType.Text, parm_initi); // 執行查詢

            // 根據角色設定按鈕狀態
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                switch (row["role_id"].ToString())
                {
                    case "0": // 管理者
                        btn_delPosition.Visible = true; // 顯示刪除按鈕
                        break;
                    case "1": // 出庫作業
                        btn_Out.Enabled = true; // 啟用出庫按鈕
                        break;
                    case "2": // 入庫作業
                        btn_Input.Enabled = true; // 啟用入庫按鈕
                        btn_BatIn.Enabled = true; // 啟用批次入庫按鈕
                        break;
                    case "3": // 資料維護
                        btn_Maintain.Visible = true; // 顯示維護按鈕
                        break;
                    case "4": // 批次出庫(儲位)
                        btn_BatOut.Enabled = true; // 啟用批次出庫按鈕
                        break;
                    default:
                        break; // 其他角色不做任何設定
                }
            }
        }

        /// <summary>
        /// <summary>
        /// 取得字串的班別資訊，根據特定格式判斷是否需要截斷字串
        /// </summary>
        /// <param name="txt">輸入的原始字串</param>
        /// <returns>處理後的班別字串</returns>
        /// <remarks>
        /// 此方法會判斷字串最後三碼是否為班別代碼（如 -01 ~ -05），
        /// 若不是則將最後三碼移除，否則回傳原字串。
        /// </remarks>
        /// <example>
        /// <code>
        /// string result = subShift("ABC-01"); // 回傳 "ABC-01"
        /// string result = subShift("ABC-99"); // 回傳 "ABC"
        /// </code>
        /// </example>
        private string subShift(string txt)
        {
            // 取得字串最後三碼的第一個字元
            string sub = StringSplit.StrLeft(StringSplit.StrRight(txt, 3), 1);
            // 判斷是否為 "-" 且不是標準班別代碼（-01 ~ -05）
            string strShift = (sub == "-" && (StringSplit.StrRight(txt, 3).CompareTo("-01") != 0
                            || StringSplit.StrRight(txt, 3).CompareTo("-02") != 0
                            || StringSplit.StrRight(txt, 3).CompareTo("-03") != 0
                            || StringSplit.StrRight(txt, 3).CompareTo("-04") != 0
                            || StringSplit.StrRight(txt, 3).CompareTo("-05") != 0))
                   // 若符合條件則移除最後三碼，否則回傳原字串
                   ? StringSplit.StrLeft(txt, txt.Length - 3) : txt;

            // 回傳處理後的班別字串
            return strShift;
        }

        /// <summary>
        /// 返回 Excel 批次入庫畫面，隱藏 panel1
        /// </summary>
        /// <param name="sender">事件呼叫來源物件</param>
        /// <param name="e">事件參數</param>
        /// <remarks>
        /// 此方法用於批次入庫作業完成或取消時，返回主畫面並隱藏 panel1。
        /// </remarks>
        /// <example>
        /// <code>
        /// btn_return_Click(sender, e); // 隱藏 panel1，返回主畫面
        /// </code>
        /// </example>
        private void btn_return_Click(object sender, EventArgs e)
        {
            // 隱藏批次入庫的 panel1，返回主畫面
            panel1.Visible = false;
        }

        /// <summary>
        /// 背景工作執行緒，負責批次寫入 Excel 資料到資料庫
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        /// <remarks>
        /// 此方法會在 inputExcelBW 背景執行緒啟動時呼叫，執行批次 Excel 資料寫入動作。
        /// </remarks>
        /// <example>
        /// <code>
        /// inputExcelBW.DoWork += inputExcelBW_DoWork;
        /// inputExcelBW.RunWorkerAsync();
        /// </code>
        /// </example>
        private void inputExcelBW_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                // 執行批次寫入 Excel 資料到資料庫
                WriteExcelData();
            }
            catch (Exception)
            {
                // 捕捉例外但不做任何處理（建議可補充錯誤處理邏輯）
            }
        }

        /// <summary>
        /// 處理 Excel 批次入庫背景工作完成事件
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        /// <remarks>
        /// 此方法會在 inputExcelBW 背景執行緒完成後呼叫，負責更新進度條、顯示完成訊息、重新整理資料、啟用按鈕、清空資料集與路徑欄位，並重新啟動計時器。
        /// </remarks>
        /// <example>
        /// <code>
        /// inputExcelBW.RunWorkerCompleted += inputExcelBW_RunWorkerCompleted;
        /// </code>
        /// </example>
        private void inputExcelBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = inputexcelcount; // 設定進度條值為已處理的 Excel 筆數
            //progressBar1.Maximum = 100; // 可選：設定進度條最大值
            MessageBox.Show("執行完成"); // 顯示完成訊息
            dataBind(); // 重新整理資料顯示
            commitButton.Enabled = true; // 啟用提交按鈕
            dsData.Clear(); // 清空資料集
            txt_path.Text = ""; // 清空檔案路徑欄位
            TimerStart(); // 重新啟動計時器
        }

        /// <summary>
        /// 處理 Excel 批次入庫背景工作進度更新事件
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">進度變更事件參數，包含目前進度百分比</param>
        /// <remarks>
        /// 此方法會在 inputExcelBW 背景執行緒進度更新時呼叫，負責更新進度條顯示目前進度百分比。
        /// </remarks>
        /// <example>
        /// <code>
        /// inputExcelBW.ProgressChanged += inputExcelBW_ProgressChanged;
        /// </code>
        /// </example>
        private void inputExcelBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // 設定進度條的值為目前進度百分比
            progressBar1.Value = e.ProgressPercentage;
        }

        /// <summary>
        /// 版本更新記錄
        /// </summary>
        /// <param name="tool"></param>
        /// <returns></returns>
        private string selectVerSQL_new(string tool)//Version Check new
        {
            string sqlCmd = "";
            try
            {
                sqlCmd = "select *  FROM Program_Table where [Program_Name] ='" + tool + "'";
                DataSet ds = db.reDs(sqlCmd);
                if (ds.Tables[0].Rows.Count != 0)
                {
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        if (ds.Tables[0].Columns[i].ToString() == "Version")
                        {
                            version_new = ds.Tables[0].Rows[0][i].ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新異常");
            }
            return version_new;
        }

        /// <summary>
        /// Form1 顯示事件處理函式，視窗顯示時自動最大化
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        /// <remarks>
        /// 此方法會在 Form1 顯示時觸發，將視窗狀態設為最大化，確保使用者一開始就能看到完整畫面。
        /// </remarks>
        /// <example>
        /// <code>
        /// Form1_Shown(sender, e);
        /// </code>
        /// </example>
        private void Form1_Shown(object sender, EventArgs e)
        {
            // 將視窗狀態設為最大化
            this.WindowState = FormWindowState.Maximized;
        }


        /// <summary>
        /// Form1 視窗大小調整事件處理函式，根據視窗縮放比例自動調整所有控制項的尺寸與位置
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        /// <remarks>
        /// 此方法會在 Form1 視窗大小變更時觸發，根據原始寬高計算縮放比例，並呼叫 SetControls 方法遞迴調整所有控制項。
        /// 若原始寬高尚未初始化（X 或 Y 為 0），則不執行任何操作。
        /// </remarks>
        /// <example>
        /// <code>
        /// Form1_Resize(sender, e); // 視窗大小調整時自動呼叫
        /// </code>
        /// </example>
        private void Form1_Resize(object sender, EventArgs e)
        {
            // 若原始寬度或高度尚未設定，則直接返回不做任何處理
            if (X == 0 || Y == 0) return;
            // 計算目前寬度與原始寬度的縮放比例
            fgX = (float)this.Width / (float)X;
            // 計算目前高度與原始高度的縮放比例
            fgY = (float)this.Height / (float)Y;
            // 呼叫 SetControls 方法，根據縮放比例調整所有控制項
            SetControls(fgX, fgY, this);
        }

        /// <summary>
        /// 歷史查詢-查詢全部
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_findAll_Click(object sender, EventArgs e)
        {
            btn_fAll = 1;
            dataBind();
        }

        /// <summary>
        /// 匯出檔案路徑，儲存目前匯出 Excel 或 CSV 檔案的完整路徑
        /// </summary>
        /// <remarks>
        /// 此欄位用於記錄匯出檔案的路徑，供匯出作業及後續開啟檔案使用。
        /// 會依照匯出時間自動產生唯一檔名，避免重複覆蓋。
        /// </remarks>
        /// <example>
        /// <code>
        /// filePath = Application.StartupPath + "\\Upload\\查詢結果_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
        /// </code>
        /// </example>
        private string filePath = string.Empty; // 匯出檔案路徑

        /// <summary>
        /// 匯出
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_qexport_Click(object sender, EventArgs e)
        {
            progressBar2.Visible = Enabled;
            filePath = "";
            //string strFilePath = txt_path.Text;
            string newName = "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

            //檔案路徑
            filePath = Application.StartupPath + "\\Upload\\查詢結果" + newName;

            // 確保在匯出前重新獲取完整數據
            RefreshExportData();

            // 檢查是否有數據可供匯出
            if (dt == null || dt.Rows.Count == 0)
            {
                MessageBox.Show("沒有可匯出的數據！請先進行查詢。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                progressBar2.Visible = false;
                return;
            }

            this.outputExcelBW.WorkerSupportsCancellation = true; //允許中斷
            this.outputExcelBW.RunWorkerAsync(); //呼叫背景程式
        }

        /// <summary>
        /// 產生Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void outputExcelBW_DoWork(object sender, DoWorkEventArgs e)
        {
            expExcelSheet();
            GC.Collect();
            //this.Close();
        }

        /// <summary>
        /// 顯示匯出完成視窗
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void outputExcelBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("匯出完成");
            progressBar2.Visible = false;
        }

        /// <summary>
        /// 在匯出前重新獲取完整數據
        /// </summary>
        private void RefreshExportData()
        {
            try
            {
                // 根據當前查詢類型獲取完整數據（不分頁）
                switch (queryType)
                {
                    case 1: // 儲位查詢
                        string positionSql;
                        if (is_History_Query)
                        {
                            positionSql = @"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                                FROM Automatic_Storage_Detail a 
                                LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                WHERE a.Unit_No=@unitNo AND Position LIKE @position+'%' AND Up_InDate > '2023-06-30'
                                ORDER BY a.sno";
                        }
                        else
                        {
                            positionSql = @"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                                FROM Automatic_Storage_Detail a 
                                LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                WHERE a.Unit_No=@unitNo AND Position LIKE @position+'%' AND amount > 0
                                ORDER BY a.sno";
                        }

                        SqlParameter[] positionParams = new SqlParameter[]
                        {
                            new SqlParameter("unitNo", Login.Unit_No),
                            new SqlParameter("position", current_Position.Trim())
                        };

                        dt = db.ExecuteDataTable(positionSql, CommandType.Text, positionParams);
                        break;

                    case 2: // 料號查詢
                        string itemSql;
                        if (is_Rad_E)
                        {
                            itemSql = @"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                                FROM Automatic_Storage_Detail a 
                                LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                WHERE a.Unit_No=@unitNo AND Item_No_Master LIKE '%'+@master+'%' AND amount > 0
                                ORDER BY a.position";
                        }
                        else
                        {
                            itemSql = @"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                                FROM Automatic_Storage_Detail a 
                                LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                WHERE a.Unit_No=@unitNo AND Item_No_Slave LIKE '%'+@master+'%' AND amount > 0
                                ORDER BY a.position";
                        }

                        SqlParameter[] itemParams = new SqlParameter[]
                        {
                            new SqlParameter("unitNo", Login.Unit_No),
                            new SqlParameter("master", current_ItemNo.Trim())
                        };

                        dt = db.ExecuteDataTable(itemSql, CommandType.Text, itemParams);
                        break;

                    case 3: // 料號+儲位組合查詢
                        string combiSql;
                        if (is_Combi_E_Query)
                        {
                            combiSql = @"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                                FROM Automatic_Storage_Detail a 
                                LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                WHERE a.Unit_No=@unitNo AND Item_No_Master LIKE '%'+@master+'%' AND Position LIKE @position+'%' AND amount > 0
                                ORDER BY a.position";
                        }
                        else
                        {
                            combiSql = @"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                                FROM Automatic_Storage_Detail a 
                                LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                WHERE a.Unit_No=@unitNo AND Item_No_Slave LIKE '%'+@master+'%' AND Position LIKE @position+'%' AND amount > 0
                                ORDER BY a.position";
                        }

                        SqlParameter[] combiParams = new SqlParameter[]
                        {
                            new SqlParameter("unitNo", Login.Unit_No),
                            new SqlParameter("master", current_CombiItem.Trim()),
                            new SqlParameter("position", current_CombiPosition.Trim())
                        };

                        dt = db.ExecuteDataTable(combiSql, CommandType.Text, combiParams);
                        break;

                    default: // 預設查詢
                        string defaultSql = @"SELECT Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                            FROM Automatic_Storage_Detail a 
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                            WHERE a.Unit_No=@unitNo AND amount > 0
                            ORDER BY a.position";

                        SqlParameter[] defaultParams = new SqlParameter[]
                        {
                            new SqlParameter("unitNo", Login.Unit_No)
                        };

                        dt = db.ExecuteDataTable(defaultSql, CommandType.Text, defaultParams);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"獲取匯出數據時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                dt = new DataTable(); // 確保 dt 不為 null
            }
        }

        /// <summary>
        /// 歷史查詢
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_history_Click(object sender, EventArgs e)
        {
            txt_item.Text = "";
            txt_position.Text = "";
            panel5.Visible = true; //歷史查詢畫面
            panel5.BringToFront();
            panel_rad.BringToFront();
            dt.Clear();
        }

        /// <summary>
        /// 處理規格查詢的 Enter 鍵事件，根據輸入內容執行模糊查詢或歷史查詢
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">鍵盤事件參數</param>
        /// <remarks>
        /// 此方法會根據 txt_spec 或 txt_specP5 的內容，決定查詢模式（一般或歷史），
        /// 並執行 SQL 查詢，將結果顯示於 dataGridView1，並更新總數量。
        /// </remarks>
        /// <example>
        /// <code>
        /// txt_spec_KeyPress(sender, e);
        /// </code>
        /// </example>
        private void txt_spec_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否按下 Enter 鍵
            if (e.KeyChar == 13)
            {
                bool h_Flag = false; // 是否歷史查詢
                string strShift = string.Empty; // 查詢用的規格字串

                // 判斷是否有歷史查詢條件
                if (txt_specP5.Text != "")
                {
                    h_Flag = true;
                }

                // 判斷是否有查詢條件（一般或歷史）
                if (txt_spec.Text != "" || h_Flag)
                {
                    try
                    {
                        // 預設使用 txt_spec 的內容
                        strShift = txt_spec.Text.Trim().ToUpper();

                        // 一般查詢：模糊查詢規格，長度需 >= 3
                        string stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  " +
                        "from Automatic_Storage_Detail a " +
                        "left join Automatic_Storage_Package b on a.Package = b.code " +
                        "left join Automatic_Storage_User c on a.Input_UserNo = c.User_No " +
                        "left join Automatic_Storage_User d on a.Output_UserNo = d.User_No " +
                        "where a.Unit_No=@unitNo and  Spec like '%'+@spec+'%' and amount >0 " +
                        "order by position asc";

                        // 歷史查詢：查詢 Input/Output 資料表
                        if (h_Flag)
                        {
                            stritem = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Input_Date,c.User_name,'','',Mark,PCB_DC,CMC_DC,a.sno  
                                from Automatic_Storage_Input a 
                                left join Automatic_Storage_Package b on a.Package = b.code 
                                left join Automatic_Storage_User c on a.Input_UserNo = c.User_No 
                                where a.Unit_No=@unitNo and Spec like '%'+@spec+'%'
                                union
                                select null,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,null,'',Output_Date,d.User_name,Mark,PCB_DC,CMC_DC,a.sno  
                                from Automatic_Storage_Output a
                                left join Automatic_Storage_Package b on a.Package = b.code  
                                left join Automatic_Storage_User d on a.Output_UserNo = d.User_No
                                where a.Unit_No=@unitNo and Spec like '%'+@spec+'%' ";
                            // 歷史查詢時使用 txt_specP5 的內容
                            strShift = txt_specP5.Text.Trim().ToUpper();
                        }

                        // 建立 SQL 參數
                        SqlParameter[] parm = new SqlParameter[]
                        {
                    new SqlParameter("unitNo",Login.Unit_No),
                    new SqlParameter("spec",strShift)
                        };

                        // 執行查詢
                        dt = db.ExecuteDataTable(stritem, CommandType.Text, parm);

                        // 若有查詢結果則顯示於 dataGridView1，並標記數量為 0 的列
                        if (dt.Rows.Count > 0)
                        {
                            dataGridView1.DataSource = dt;
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                if (dt.Rows[i]["Amount"].ToString() == "0")
                                {
                                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.MediumVioletRed;
                                }
                            }
                            // 更新總數量
                            sumC();
                        }
                        else
                        {
                            // 查無資料時顯示提示訊息
                            MessageBox.Show("查無資料，請確認輸入後重新查詢");
                        }
                    }
                    catch (Exception ex)
                    {
                        // 查詢過程發生例外時顯示錯誤訊息
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    // 未輸入查詢條件時顯示提示訊息
                    MessageBox.Show("請確認料號是否正確或沒有輸入");
                }
                // 輸入搜尋條件後反白 txt_item 與 txt_itemP5
                txt_item.SelectAll();
                txt_itemP5.SelectAll();
            }
        }

        /// <summary>
        /// 工單出庫查詢事件處理函式，按下 Enter 鍵時根據工單號查詢出庫資料
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">鍵盤事件參數</param>
        /// <remarks>
        /// 此方法會在 txt_wonoOut 控制項按下 Enter 鍵時觸發，根據輸入的工單號 (Wo_No) 查詢 Automatic_Storage_Output 資料表，
        /// 並將查詢結果顯示於 dataGridView1。
        /// </remarks>
        /// <example>
        /// <code>
        /// txt_wonoOut_KeyPress(sender, e);
        /// </code>
        /// </example>
        private void txt_wonoOut_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否按下 Enter 鍵
            if (e.KeyChar == 13)
            {
                // 定義查詢工單出庫的 SQL 語句
                string sql_wonoOut = @" select Actual_InDate,a.Item_No_Master,a.Item_No_Slave,a.Spec,a.Position,
                                    a.Amount,b.Package,d.Input_Date,c.User_name,Output_Date,
                                    c.User_name,a.Mark,d.PCB_DC,CMC_DC  from Automatic_Storage_Output a
                                    left join Automatic_Storage_Package b on a.Package = b.code 
                                    left join Automatic_Storage_User c on a.Output_UserNo = c.User_No 
                                    left join Automatic_Storage_Input d on a.sno=d.sno
                                    where Wo_No =@wono order by position asc ";
                // 建立 SQL 參數，將輸入的工單號傳入查詢
                SqlParameter[] parm = new SqlParameter[]
                {
                new SqlParameter("wono",txt_wonoOut.Text.Trim())
                };
                // 執行查詢，取得結果 DataTable
                dt = db.ExecuteDataTable(sql_wonoOut, CommandType.Text, parm);
                // 將查詢結果顯示於 dataGridView1
                dataGridView1.DataSource = dt;
            }
        }

        /// <summary>
        /// 查詢後-GridView-資料驗証
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // 判斷是否為目標欄位
            if (dataGridView1.Columns[e.ColumnIndex].Name == "PCB_DC")
            {
                // 判斷該單元格的值是否符合條件
                if (!string.IsNullOrEmpty(e.Value?.ToString() ?? string.Empty) && Week_Then25(e.Value?.ToString() ?? string.Empty))
                {
                    // 設置該單元格的背景顏色為紅色
                    e.CellStyle.BackColor = Color.Red;
                }
            }
            if (dataGridView1.Columns[e.ColumnIndex].Name == "CMC_DC")
            {
                if (!string.IsNullOrEmpty(e.Value?.ToString() ?? string.Empty) && cmc.IsOverTwoYears(e.Value?.ToString() ?? string.Empty))
                {
                    // 設置該單元格的背景顏色為黃色
                    e.CellStyle.BackColor = Color.Yellow;
                }
            }
        }

        /// <summary>
        /// 處理日期查詢的 Enter 鍵事件，根據輸入的起始與結束時間執行資料查詢
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">鍵盤事件參數</param>
        /// <remarks>
        /// 此方法會在 txt_DTs 控制項按下 Enter 鍵時觸發，根據 txt_DTs 與 txt_DTe 的內容，
        /// 執行時間區間查詢，將結果顯示於 dataGridView1，並更新總數量。
        /// </remarks>
        /// <example>
        /// <code>
        /// txt_DTs_KeyPress(sender, e);
        /// </code>
        /// </example>
        private void txt_DTs_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否按下 Enter 鍵
            if (e.KeyChar == 13)
            {
                // 檢查起始與結束時間是否有輸入
                if (!string.IsNullOrEmpty(txt_DTs.Text) && !string.IsNullOrEmpty(txt_DTe.Text))
                {
                    // 檢查時間格式是否正確（長度為 12）
                    if (txt_DTs.Text.Length == 12 && txt_DTe.Text.Length == 12)
                    {
                        #region 時間查詢
                        // 解析起始時間
                        DateTime sdate = DateTime.ParseExact(txt_DTs.Text, "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                        // 解析結束時間
                        DateTime edate = DateTime.ParseExact(txt_DTe.Text, "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                        // 格式化起始時間字串
                        string s_date = sdate.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        // 格式化結束時間字串
                        string e_date = edate.ToString("yyyy-MM-dd HH:mm:59.999");

                        string strsql = "";
                        // 判斷查詢類型（入庫或出庫）
                        if (rad_in.Checked)
                        {
                            // 入庫查詢 SQL
                            strsql = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,
                             Amount,b.Package,Input_Date,c.User_name,'','',Mark,PCB_DC,CMC_DC,a.sno  
                             from Automatic_Storage_Input a 
                             left join Automatic_Storage_Package b on a.Package = b.code 
                             left join Automatic_Storage_User c on a.Input_UserNo = c.User_No 
                             where Input_Date between @startDate and @endDate";
                        }
                        else if (rad_out.Checked)
                        {
                            // 出庫查詢 SQL
                            strsql = @"select null,Item_No_Master,Item_No_Slave,Spec,Position,
                            Amount,b.Package,'','',Output_Date,d.User_name,Mark,PCB_DC,CMC_DC,a.sno   
                            from Automatic_Storage_Output a
                            left join Automatic_Storage_Package b on a.Package = b.code  
                            left join Automatic_Storage_User d on a.Output_UserNo = d.User_No
                            where Output_Date between @startDate and @endDate";
                        }
                        // 建立 SQL 參數
                        SqlParameter[] parm = new SqlParameter[]
                        {
                    new SqlParameter("startDate",s_date),
                    new SqlParameter("endDate",e_date)
                        };
                        // 執行查詢，取得結果 DataTable
                        dt = db.ExecuteDataTable(strsql, CommandType.Text, parm);
                        // 將查詢結果顯示於 dataGridView1
                        dataGridView1.DataSource = dt;
                        #endregion
                    }
                }
                else
                {
                    // 未輸入時間時顯示提示訊息
                    MessageBox.Show("請確認時間格式是否正確或沒有輸入");
                }
                // 更新總數量
                sumC();
            }
        }

        /// <summary>
        /// GridView 取選欄位
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            // 檢查是否有選取的列
            if (dataGridView1.CurrentRow == null)
            {
                return;
            }

            /*
             0=actual_indate
             1=item_no_master
            2=item_no_slave
            3=spec
            4=position
            5=amount
            6=package
            7=up_indate
            8=user_name
            9=up_outdate
            10=user_name1
            11=mark
            12=pcb_dc
            13=cmc_dc
            14=sno
             */
            //for (int i = 0; i < dataGridView1.Columns.Count; i++)
            //{
            //    string dda = dataGridView1.Columns[i].HeaderText.ToString();
            //}

            // 取得該列的 sno
            string sno = dataGridView1.CurrentRow.Cells["sno"].Value?.ToString();

            // 取得目前雙擊的欄位索引
            int columnIndex = dataGridView1.CurrentCell.ColumnIndex;

            // 判斷雙擊的欄位索引
            if (columnIndex == 0) // 針對 actual_indate 欄位 (索引 0)
            {
                // 取得目前的值
                string actualDate = dataGridView1.CurrentRow.Cells[0].Value?.ToString();
                if (DateTime.TryParse(actualDate, out DateTime parsedDate))
                {
                    actualDate = parsedDate.ToString("yyyy-MM-dd");
                }
                // 顯示 ActualDate 表單並傳遞相關資訊
                ActualDate acD = new ActualDate();
                acD.Owner = this;
                acD.Sno = sno;
                acD.AcD_O = actualDate;
                acD.Show();
            }
            else if (columnIndex == 12) // PCB_DC 欄位 (索引 12)
            {
                string strPcbDC = dataGridView1.CurrentRow.Cells[12].Value?.ToString();
                PcbDC pcbDC = new PcbDC();
                pcbDC.Owner = this;
                pcbDC.Sno = sno;
                pcbDC.PcbDC_O = strPcbDC;
                pcbDC.Show();
            }
        }

        /// <summary>
        /// 處理 txt_position 控制項的滑鼠雙擊事件，通常用於快速清空或選擇儲位查詢條件
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">滑鼠事件參數</param>
        /// <remarks>
        /// 此事件可用於提供使用者快速操作，例如清空儲位查詢欄位或彈出選擇視窗。
        /// </remarks>
        /// <example>
        /// <code>
        /// txt_position.MouseDoubleClick += txt_position_MouseDoubleClick;
        /// </code>
        /// </example>
        private void txt_position_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // 當使用者在 txt_position 控制項上雙擊滑鼠時觸發
            // 這裡可以加入清空 txt_position 或彈出儲位選擇視窗的邏輯
            // 目前尚未實作任何功能
        }

        /// <summary>
        /// 返回 - 原查詢頁面
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_reP2_Click(object sender, EventArgs e)
        {
            txt_itemP5.Text = "";
            txt_siteP5.Text = "";
            panel2.BringToFront();
            dataBind();
        }

        /// <summary>
        /// 將控制項的寬，高，左邊距，頂邊距和字體大小暫存到tag屬性中
        /// </summary>
        /// <param name="cons">遞歸控制項中的控制項</param>
        private void SetTag(Control cons)
        {
            foreach (Control con in cons.Controls)
            {
                con.Tag = con.Width + ":" + con.Height + ":" + con.Left + ":" + con.Top + ":" + con.Font.Size;
                if (con.Controls.Count > 0)
                    SetTag(con);
            }
        }

        /// <summary>
        /// 設置控制項的值
        /// </summary>
        /// <param name="newx"></param>
        /// <param name="newy"></param>
        /// <param name="cons"></param>
        private void SetControls(float newx, float newy, Control cons)
        {
            if (isLoaded)
            {
                //遍歷窗體中的控制項，重新設置控制項的值
                foreach (Control con in cons.Controls)
                {
                    string[] mytag = con.Tag.ToString().Split(new char[] { ':' });//獲取控制項的Tag屬性值，並分割後存儲字元串數組
                    float a = System.Convert.ToSingle(mytag[0]) * newx;//根據窗體縮放比例確定控制項的值，寬度
                    con.Width = (int)a;//寬度
                    a = System.Convert.ToSingle(mytag[1]) * newy;//高度
                    con.Height = (int)(a);
                    a = System.Convert.ToSingle(mytag[2]) * newx;//左邊距離
                    con.Left = (int)(a);
                    a = System.Convert.ToSingle(mytag[3]) * newy;//上邊緣距離
                    con.Top = (int)(a);
                    Single currentSize = System.Convert.ToSingle(mytag[4]) * newy;//字體大小
                    con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
                    if (con.Controls.Count > 0)
                    {
                        SetControls(newx, newy, con);
                    }
                }
            }
        }

        /// <summary>
        /// Excel匯出作業 - 提供高效能的 CSV 匯出和 Excel 匯出選項
        /// </summary>
        private void expExcelSheet()
        {
            try
            {
                // 詢問用戶選擇匯出格式
                DialogResult result = MessageBox.Show("請選擇匯出格式:\n\n是 - 匯出為 Excel 格式 (.xlsx)\n否 - 匯出為 CSV 格式 (.csv)\n\n注意: CSV 格式匯出速度更快，適合大量數據",
                    "選擇匯出格式", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // 用戶選擇了 Excel 格式
                    ExportToExcel();
                }
                else
                {
                    // 用戶選擇了 CSV 格式
                    ExportToCSV();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// 使用高效能方式匯出為 CSV 格式
        /// </summary>
        private void ExportToCSV()
        {
            try
            {
                // 設置進度條
                SetupProgressBar();

                // 確保目錄存在
                string csvFilePath = filePath;
                if (!csvFilePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    csvFilePath = Path.ChangeExtension(csvFilePath, ".csv");
                }

                string directoryPath = Path.GetDirectoryName(csvFilePath);
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                // 使用 StreamWriter 直接寫入文件，比 StringBuilder 更節省內存
                using (StreamWriter sw = new StreamWriter(csvFilePath, false, new UTF8Encoding(true)))
                {
                    // 寫入標題行
                    string[] headers = new string[14];
                    for (int i = 0; i <= 13; i++)
                    {
                        headers[i] = EscapeCsvField(dataGridView1.Columns[i].HeaderText);
                    }
                    sw.WriteLine(string.Join(",", headers));

                    // 寫入數據行
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string[] rowData = new string[14];
                        for (int j = 0; j <= 13; j++)
                        {
                            rowData[j] = EscapeCsvField(dt.Rows[i][j].ToString());
                        }
                        sw.WriteLine(string.Join(",", rowData));

                        // 更新進度條
                        UpdateProgressBar();
                    }
                }

                // 自動打開文件
                if (MessageBox.Show($"CSV 檔案已成功匯出至: {csvFilePath}\n\n是否要立即開啟此檔案?", "匯出成功",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(csvFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出 CSV 時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// 使用 Microsoft.Office.Interop.Excel 匯出為 Excel 格式
        /// </summary>
        private void ExportToExcel()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            try
            {
                // 設置進度條
                SetupProgressBar();

                // 創建 Excel 應用程序
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                // 創建新的工作簿
                workbook = excelApp.Workbooks.Add(Type.Missing);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];

                // 一次性設置整個範圍的值，而不是單獨設置每個單元格
                object[,] data = new object[dt.Rows.Count + 1, 14];

                // 填充標題行
                for (int i = 0; i <= 13; i++)
                {
                    data[0, i] = dataGridView1.Columns[i].HeaderText;
                }

                // 填充數據
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j <= 13; j++)
                    {
                        data[i + 1, j] = dt.Rows[i][j].ToString();
                    }
                    UpdateProgressBar();
                }

                // 一次性設置整個範圍的值
                Microsoft.Office.Interop.Excel.Range range = worksheet.Range[
                    worksheet.Cells[1, 1],
                    worksheet.Cells[dt.Rows.Count + 1, 14]];
                range.Value2 = data;

                // 設置標題行為粗體
                Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[
                    worksheet.Cells[1, 1],
                    worksheet.Cells[1, 14]];
                headerRange.Font.Bold = true;

                // 自動調整列寬
                worksheet.Columns.AutoFit();

                // 保存為 Excel 文件
                string excelFilePath = filePath;
                if (!excelFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    excelFilePath = Path.ChangeExtension(excelFilePath, ".xlsx");
                }

                // 確保目錄存在
                string directoryPath = Path.GetDirectoryName(excelFilePath);
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                // 保存工作簿
                workbook.SaveAs(excelFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);

                // 詢問是否打開文件
                if (MessageBox.Show($"Excel 檔案已成功匯出至: {excelFilePath}\n\n是否要立即開啟此檔案?", "匯出成功",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(excelFilePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯出 Excel 時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // 釋放資源
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// 設置進度條
        /// </summary>
        private void SetupProgressBar()
        {
            if (progressBar2.InvokeRequired)
            {
                progressBar2.Invoke(new Action(() =>
                {
                    progressBar2.Minimum = 0;
                    progressBar2.Maximum = dt.Rows.Count;
                    progressBar2.Step = 1;
                    progressBar2.Value = 0;
                }));
            }
            else
            {
                progressBar2.Minimum = 0;
                progressBar2.Maximum = dt.Rows.Count;
                progressBar2.Step = 1;
                progressBar2.Value = 0;
            }
        }

        /// <summary>
        /// 更新進度條
        /// </summary>
        private void UpdateProgressBar()
        {
            if (progressBar2.InvokeRequired)
            {
                progressBar2.Invoke(new Action(() => progressBar2.PerformStep()));
            }
            else
            {
                progressBar2.PerformStep();
            }
        }

        /// <summary>
        /// 處理 CSV 欄位，確保特殊字符被正確轉義
        /// </summary>
        private string EscapeCsvField(string field)
        {
            // 如果字段包含逗號、引號或換行符，則需要用引號括起來
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                // 將字段中的引號替換為兩個引號（CSV 格式的轉義方式）
                field = field.Replace("\"", "\"\"");
                // 用引號括起整個字段
                return $"\"{field}\"";
            }
            return field;
        }

        /// <summary>
        /// 自動更新
        /// </summary>
        public void autoupdate()
        {
            //寫入目前版本與程式名後執行更新

            Process p = new Process();
            p.StartInfo.FileName = System.Windows.Forms.Application.StartupPath + "\\AutoUpdate.exe";
            p.StartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath; //檔案所在的目錄
            p.Start();
            this.Close();
        }

        #region 換頁功能

        /// <summary>
        /// 查詢全部獲取指定頁的數據
        /// </summary>
        /// <param name="page">頁碼</param>
        /// <param name="checkOnly">是否只檢查是否有數據而不更新UI</param>
        /// <returns>指定頁的數據</returns>
        private DataTable GetPagedData(int page, bool checkOnly = false)
        {
            try
            {
                // 計算偏移量
                int offset = (page - 1) * pageSize;

                // 根據當前的查詢條件來決定使用什麼 SQL
                string sql = string.Empty;

                sql = $@"SELECT TOP {pageSize} Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno    
                    FROM Automatic_Storage_Detail a
                    LEFT JOIN Automatic_Storage_Package b ON a.Package=b.code
                    LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo=c.User_No 
                    LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo=d.User_No 
                    WHERE amount > 0 AND a.Unit_No=@unitNo AND a.sno NOT IN 
                    (SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a WHERE a.Unit_No=@unitNo AND amount > 0 ORDER BY a.Actual_InDate)
                    ORDER BY a.Actual_InDate";

                if (btn_fAll == 1)
                {
                    sql = $@"SELECT TOP {pageSize} Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,Package,Up_InDate,Input_UserNo,Up_OutDate,Output_UserNo,Mark,PCB_DC,CMC_DC,sno 
                    FROM Automatic_Storage_Detail 
                    WHERE Unit_No=@unitNo AND sno NOT IN 
                    (SELECT TOP {offset} sno FROM Automatic_Storage_Detail WHERE Unit_No=@unitNo ORDER BY Actual_InDate)
                    ORDER BY Actual_InDate";
                }

                // 料號查詢
                if (!string.IsNullOrEmpty(txt_item.Text))
                {
                    string stritem = $@"SELECT TOP {pageSize} a.Actual_InDate,a.Item_No_Master,a.Item_No_Slave,a.Spec,a.Position,a.Amount,b.Package
                                            ,a.Up_InDate,c.User_name,a.Up_OutDate,d.User_name,a.Mark,a.PCB_DC,a.CMC_DC,a.sno  
                                            FROM Automatic_Storage_Detail a 
                                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                            WHERE a.amount > 0 AND a.Unit_No=@unitNo AND a.Item_No_Master LIKE '%'+@master+'%' 
                                            AND a.sno NOT IN 
                                            (
                                                SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a 
                                                WHERE a.Unit_No=@unitNo AND a.Item_No_Master LIKE '%'+@master+'%' AND a.amount > 0 ORDER BY a.Actual_InDate
                                            )
                                            ORDER BY a.Actual_InDate";

                    if (rad_E.Checked)
                    {
                        stritem = $@"SELECT TOP {pageSize} a.Actual_InDate,a.Item_No_Master,a.Item_No_Slave,a.Spec,a.Position,a.Amount,b.Package
                                         ,a.Up_InDate,c.User_name,a.Up_OutDate,d.User_name,a.Mark,a.PCB_DC,a.CMC_DC,a.sno                              
                                        FROM Automatic_Storage_Detail a 
                                        LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                        LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                        LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                        WHERE a.amount > 0 AND a.Unit_No=@unitNo AND a.Item_No_Master LIKE '%'+@master+'%' AND a.sno NOT IN 
                                        (
                                            SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a WHERE a.Unit_No=@unitNo AND a.Item_No_Master LIKE '%'+@master+'%' 
                                            AND a.amount > 0 ORDER BY a.Actual_InDate
                                        )
                                        ORDER BY a.Actual_InDate";
                    }
                    else if (rad_C.Checked)
                    {
                        stritem = $@"SELECT TOP {pageSize} a.Actual_InDate,a.Item_No_Master,a.Item_No_Slave,a.Spec,a.Position,a.Amount,b.Package,a.Up_InDate,c.User_name
                                        ,a.Up_OutDate,d.User_name,a.Mark,a.PCB_DC,a.CMC_DC,a.sno  
                                        FROM Automatic_Storage_Detail a 
                                        LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                        LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                        LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                        WHERE a.amount > 0 AND a.Unit_No=@unitNo AND a.Item_No_Slave LIKE '%'+@master+'%' AND a.sno NOT IN 
                                        (
                                            SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a WHERE a.Unit_No=@unitNo AND a.Item_No_Slave LIKE '%'+@master+'%' 
                                            AND a.amount > 0 ORDER BY a.Actual_InDate
                                        )
                                        ORDER BY a.Actual_InDate";

                    }

                    SqlParameter[] parm = new SqlParameter[]
                    {
                        new SqlParameter("unitNo", Login.Unit_No),
                        new SqlParameter("master", txt_item.Text.Trim().ToUpper())
                    };

                    return db.ExecuteDataTable(stritem, CommandType.Text, parm);
                }
                // 儲位查詢
                else if (!string.IsNullOrEmpty(txt_position.Text))
                {
                    string position = txt_position.Text;
                    string stritem = $@"SELECT TOP {pageSize} a.Actual_InDate,a.Item_No_Master,a.Item_No_Slave,a.Spec,a.Position,a.Amount,b.Package,a.Up_InDate,c.User_name
                                ,a.Up_OutDate,d.User_name,a.Mark,a.PCB_DC,a.CMC_DC,a.sno  
                                FROM Automatic_Storage_Detail a 
                                LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                                LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                                LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                                WHERE a.amount > 0 AND a.Unit_No=@unitNo AND a.Position LIKE @position+'%' 
                                AND a.sno NOT IN 
                                (
                                    SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a WHERE a.Unit_No=@unitNo AND a.Position LIKE @position+'%' 
                                    AND a.amount > 0 ORDER BY a.Actual_InDate
                                )
                                ORDER BY a.Actual_InDate";

                    SqlParameter[] parm = new SqlParameter[]
                    {
                        new SqlParameter("unitNo", Login.Unit_No),
                        new SqlParameter("position", position.Trim())
                    };

                    return db.ExecuteDataTable(stritem, CommandType.Text, parm);
                }
                // 規格查詢
                else if (!string.IsNullOrEmpty(txt_spec.Text))
                {
                    string stritem = $@"SELECT TOP {pageSize} a.Actual_InDate,a.Item_No_Master,a.Item_No_Slave,a.Spec,a.Position,a.Amount,b.Package,a.Up_InDate,c.User_name,a.Up_OutDate,d.User_name,a.Mark,a.PCB_DC,a.CMC_DC,a.sno  
                        FROM Automatic_Storage_Detail a 
                        LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                        LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                        LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                        WHERE a.amount > 0 AND a.Unit_No=@unitNo AND a.Spec LIKE '%'+@spec+'%' AND a.sno NOT IN 
                        (SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a WHERE a.Unit_No=@unitNo AND a.Spec LIKE '%'+@spec+'%' AND a.amount > 0 ORDER BY a.Actual_InDate)
                        ORDER BY a.Actual_InDate";

                    SqlParameter[] parm = new SqlParameter[]
                    {
                        new SqlParameter("unitNo", Login.Unit_No),
                        new SqlParameter("spec", txt_spec.Text.Trim().ToUpper())
                    };

                    return db.ExecuteDataTable(stritem, CommandType.Text, parm);
                }
                // 預設查詢（全部資料）
                else
                {
                    // 預先查詢總筆數以計算總頁數
                    int totalRows = 0;
                    string countSql;

                    countSql = @"SELECT COUNT(*) as cnt
                            FROM Automatic_Storage_Detail a 
                            WHERE a.Unit_No=@unitNo 
                            AND amount > 0";


                    SqlParameter[] countParm = new SqlParameter[]
                    {
                        new SqlParameter("unitNo", Login.Unit_No)
                    };

                    using (DataTable dtCount = db.ExecuteDataTable(countSql, CommandType.Text, countParm))
                    {
                        if (dtCount.Rows.Count > 0)
                        {
                            totalRows = Convert.ToInt32(dtCount.Rows[0]["cnt"]);
                        }
                    }
                    int totalPages = (int)Math.Ceiling(totalRows / (double)pageSize);

                    SqlParameter[] parm = new SqlParameter[]
                    {
                        new SqlParameter("unitNo", Login.Unit_No)
                    };

                    DataTable currentData = db.ExecuteDataTable(sql, CommandType.Text, parm);

                    // 如果只是檢查是否有數據，直接返回結果
                    if (checkOnly || currentData == null || currentData.Rows.Count == 0)
                    {
                        return currentData;
                    }

                    // 更新UI
                    dataGridView1.DataSource = currentData;

                    // 設置行背景色
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        if ((dataGridView1.Rows[i].Cells["Amount"].Value?.ToString() ?? string.Empty) == "0")
                        {
                            dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.MediumVioletRed;
                        }
                    }

                    // 更新頁碼顯示
                    if (lblCurrentPage != null)
                    {
                        lblCurrentPage.Text = $"第 {page} 頁";
                        lblCurrentPage.Visible = true;
                    }

                    // 顯示分頁按鈕
                    btnPreviousPage.Visible = true;
                    btnNextPage.Visible = true;

                    // 顯示目前頁數/總頁數
                    if (lblPageInfo != null)
                    {
                        lblPageInfo.Text = $"第 {page} 頁 / 共 {totalPages} 頁";
                    }

                    sumC();

                    return currentData;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲取分頁數據時發生錯誤: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 獲取料號查詢的分頁數據
        /// </summary>
        /// <param name="page">頁碼</param>
        /// <param name="checkOnly">是否只檢查是否有數據而不更新UI</param>
        /// <returns>查詢結果的DataTable</returns>
        private DataTable GetItemPagedData(int page, bool checkOnly = false)
        {
            try
            {
                // 計算偏移量
                int offset = (page - 1) * pageSize;
                string stritem;
                string itemNo = current_ItemNo;
                bool h_Flag = is_History_Query;
                bool e_Flag = is_Rad_E;

                // 預先查詢總筆數以計算總頁數
                int totalRows = 0;
                string countSql;

                if (h_Flag)
                {
                    if (e_Flag)
                    {
                        countSql = @"SELECT COUNT(*) as cnt
                            FROM (SELECT a.sno FROM Automatic_Storage_Input a 
                                  WHERE a.Unit_No=@unitNo AND Item_No_Master=@master AND cast(a.Input_Date as date) >'2023-06-30'
                                  UNION
                                  SELECT a.sno FROM Automatic_Storage_Output a 
                                  WHERE a.Unit_No=@unitNo AND Item_No_Master=@master AND cast(a.Output_Date as date) >'2023-06-30') AS t";
                    }
                    else
                    {
                        countSql = @"SELECT COUNT(*) as cnt
                            FROM (SELECT a.sno FROM Automatic_Storage_Input a 
                                  WHERE a.Unit_No=@unitNo AND Item_No_Slave=@master AND cast(a.Input_Date as date) >'2023-06-30'
                                  UNION
                                  SELECT a.sno FROM Automatic_Storage_Output a 
                                  WHERE a.Unit_No=@unitNo AND Item_No_Slave=@master AND cast(a.Output_Date as date) >'2023-06-30') AS t";
                    }
                }
                else
                {
                    if (e_Flag)
                    {
                        countSql = @"SELECT COUNT(*) as cnt
                            FROM Automatic_Storage_Detail  
                            WHERE Unit_No=@unitNo 
                            AND Item_No_Master LIKE '%'+@master+'%' 
                            AND amount > 0";
                    }
                    else
                    {
                        countSql = @"SELECT COUNT(*) as cnt
                            FROM Automatic_Storage_Detail  
                            WHERE Unit_No=@unitNo 
                            AND Item_No_Slave LIKE '%'+@master+'%' 
                            AND amount > 0";
                    }
                }

                SqlParameter[] countParm = new SqlParameter[]
                {
                    new SqlParameter("unitNo", Login.Unit_No),
                    new SqlParameter("master", itemNo)
                };

                using (DataTable dtCount = db.ExecuteDataTable(countSql, CommandType.Text, countParm))
                {
                    if (dtCount.Rows.Count > 0)
                    {
                        totalRows = Convert.ToInt32(dtCount.Rows[0]["cnt"]);
                    }
                }
                int totalPages = (int)Math.Ceiling(totalRows / (double)pageSize);

                // 根據查詢條件構建 SQL
                if (h_Flag)
                {
                    if (e_Flag)
                    {
                        stritem = $@"SELECT TOP {pageSize} Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Input_Date as Up_InDate,c.User_name,null as Up_OutDate,'' as User_name1,Mark,PCB_DC,CMC_DC,a.sno   
                            FROM Automatic_Storage_Input a 
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                            WHERE a.Unit_No=@unitNo AND Item_No_Master=@master AND cast(a.Input_Date as date) >'2023-06-30' AND a.sno NOT IN 
                            (SELECT TOP {offset} a.sno FROM Automatic_Storage_Input a 
                             WHERE a.Unit_No=@unitNo AND Item_No_Master=@master AND cast(a.Input_Date as date) >'2023-06-30')
                            UNION
                            SELECT null as Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,null as Up_InDate,'' as User_name,Output_Date as Up_OutDate,d.User_name as User_name1,Mark,PCB_DC,CMC_DC,a.sno 
                            FROM Automatic_Storage_Output a
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code  
                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No
                            WHERE a.Unit_No=@unitNo AND Item_No_Master=@master AND cast(a.Output_Date as date) >'2023-06-30' AND a.sno NOT IN 
                            (SELECT TOP {offset} a.sno FROM Automatic_Storage_Output a 
                             WHERE a.Unit_No=@unitNo AND Item_No_Master=@master AND cast(a.Output_Date as date) >'2023-06-30')";
                    }
                    else
                    {
                        stritem = $@"SELECT TOP {pageSize} Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Input_Date as Up_InDate,c.User_name,null as Up_OutDate,'' as User_name1,Mark,PCB_DC,CMC_DC,a.sno  
                            FROM Automatic_Storage_Input a 
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                            WHERE a.Unit_No=@unitNo AND Item_No_Slave=@master AND cast(a.Input_Date as date) >'2023-06-30' AND a.sno NOT IN 
                            (SELECT TOP {offset} a.sno FROM Automatic_Storage_Input a 
                             WHERE a.Unit_No=@unitNo AND Item_No_Slave=@master AND cast(a.Input_Date as date) >'2023-06-30')
                            UNION
                            SELECT null as Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,null as Up_InDate,'' as User_name,Output_Date as Up_OutDate,d.User_name as User_name1,Mark,PCB_DC,CMC_DC,a.sno  
                            FROM Automatic_Storage_Output a
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code  
                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No
                            WHERE a.Unit_No=@unitNo AND Item_No_Slave=@master AND cast(a.Output_Date as date) >'2023-06-30' AND a.sno NOT IN 
                            (SELECT TOP {offset} a.sno FROM Automatic_Storage_Output a 
                             WHERE a.Unit_No=@unitNo AND Item_No_Slave=@master AND cast(a.Output_Date as date) >'2023-06-30')";
                    }
                }
                else
                {
                    if (e_Flag)
                    {
                        stritem = $@"SELECT TOP {pageSize} Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name as User_name1,a.Mark,PCB_DC,CMC_DC,a.sno  
                            FROM Automatic_Storage_Detail a 
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                            WHERE a.Unit_No=@unitNo AND Item_No_Master LIKE '%'+@master+'%' AND amount > 0 AND a.sno NOT IN 
                            (SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a 
                             WHERE a.Unit_No=@unitNo AND Item_No_Master LIKE '%'+@master+'%' AND amount > 0 ORDER BY a.Actual_InDate)
                            ORDER BY Actual_InDate";
                    }
                    else
                    {
                        stritem = $@"SELECT TOP {pageSize} Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name as User_name1,a.Mark,PCB_DC,CMC_DC,a.sno  
                            FROM Automatic_Storage_Detail a 
                            LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                            LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                            LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                            WHERE a.Unit_No=@unitNo AND Item_No_Slave LIKE '%'+@master+'%' AND amount > 0 AND a.sno NOT IN 
                            (SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a 
                             WHERE a.Unit_No=@unitNo AND Item_No_Slave LIKE '%'+@master+'%' AND amount > 0 ORDER BY a.Actual_InDate)
                            ORDER BY Actual_InDate";
                    }
                }

                SqlParameter[] parm = new SqlParameter[]
                {
                    new SqlParameter("unitNo", Login.Unit_No),
                    new SqlParameter("master", itemNo)
                };

                DataTable currentData = db.ExecuteDataTable(stritem, CommandType.Text, parm);

                // Excel
                //GetDateForExcel("", e_Flag, parm);


                // 如果只是檢查是否有數據，直接返回結果
                if (checkOnly || currentData == null || currentData.Rows.Count == 0)
                {
                    return currentData;
                }

                // 更新UI
                dataGridView1.DataSource = currentData;

                // 設置行背景色
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if ((dataGridView1.Rows[i].Cells["Amount"].Value?.ToString() ?? string.Empty) == "0")
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.MediumVioletRed;
                    }
                }

                // 更新頁碼顯示
                if (lblCurrentPage != null)
                {
                    lblCurrentPage.Text = $"第 {currentPage} 頁";
                    lblCurrentPage.Visible = true;
                }

                // 顯示分頁按鈕
                btnPreviousPage.Visible = true;
                btnNextPage.Visible = true;

                // 顯示目前頁數/總頁數
                if (lblPageInfo != null)
                {
                    lblPageInfo.Text = $"第 {currentPage} 頁 / 共 {totalPages} 頁";
                }
                sumC();

                return currentData;
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲取料號查詢分頁數據時發生錯誤: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 獲取儲位查詢的分頁數據
        /// </summary>
        /// <param name="page">頁碼</param>
        /// <param name="checkOnly">是否只檢查是否有數據而不更新UI</param>
        /// <returns>查詢結果的DataTable</returns>
        private DataTable GetPositionPagedData(int page, bool checkOnly = false)
        {
            try
            {
                // 計算偏移量
                int offset = (page - 1) * pageSize;
                string stritem;
                string position = current_Position;
                bool h_Flag = is_History_Query;

                // 預先查詢總筆數以計算總頁數
                int totalRows = 0;
                string countSql;
                if (h_Flag)
                {
                    countSql = @"SELECT COUNT(*) as cnt
                        FROM Automatic_Storage_Detail a 
                        WHERE a.Unit_No=@unitNo 
                        AND Position LIKE @position+'%' 
                        AND Up_InDate > '2023-06-30'";
                }
                else
                {
                    countSql = @"SELECT COUNT(*) as cnt
                        FROM Automatic_Storage_Detail a 
                        WHERE a.Unit_No=@unitNo 
                        AND Position LIKE @position+'%' 
                        AND amount > 0";
                }

                SqlParameter[] countParm = new SqlParameter[]
                {
                    new SqlParameter("unitNo", Login.Unit_No),
                    new SqlParameter("position", position.Trim())
                };

                using (DataTable dtCount = db.ExecuteDataTable(countSql, CommandType.Text, countParm))
                {
                    if (dtCount.Rows.Count > 0)
                    {
                        totalRows = Convert.ToInt32(dtCount.Rows[0]["cnt"]);
                    }
                }
                int totalPages = (int)Math.Ceiling(totalRows / (double)pageSize);

                // 根據查詢條件構建 SQL
                if (h_Flag)
                {
                    stritem = $@"SELECT TOP {pageSize} Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                        FROM Automatic_Storage_Detail a 
                        LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                        LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                        LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                        WHERE a.Unit_No=@unitNo AND Position LIKE @position+'%' AND Up_InDate > '2023-06-30' AND a.sno NOT IN 
                        (SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a 
                         WHERE a.Unit_No=@unitNo AND Position LIKE @position+'%' AND Up_InDate > '2023-06-30' ORDER BY a.Actual_InDate)
                        ORDER BY a.Actual_InDate";
                }
                else
                {
                    stritem = $@"SELECT TOP {pageSize} Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  
                        FROM Automatic_Storage_Detail a 
                        LEFT JOIN Automatic_Storage_Package b ON a.Package = b.code 
                        LEFT JOIN Automatic_Storage_User c ON a.Input_UserNo = c.User_No 
                        LEFT JOIN Automatic_Storage_User d ON a.Output_UserNo = d.User_No 
                        WHERE a.Unit_No=@unitNo AND Position LIKE @position+'%' AND amount > 0 AND a.sno NOT IN 
                        (SELECT TOP {offset} a.sno FROM Automatic_Storage_Detail a 
                         WHERE a.Unit_No=@unitNo AND Position LIKE @position+'%' AND amount > 0 ORDER BY a.Actual_InDate)
                        ORDER BY a.Actual_InDate";
                }

                SqlParameter[] parm = new SqlParameter[]
                {
                    new SqlParameter("unitNo", Login.Unit_No),
                    new SqlParameter("position", position.Trim())
                };

                DataTable currentData = db.ExecuteDataTable(stritem, CommandType.Text, parm);

                //Excel
                //GetDateForExcel("", false, parm);


                // 如果只是檢查是否有數據，直接返回結果
                if (checkOnly || currentData == null || currentData.Rows.Count == 0)
                {
                    return currentData;
                }

                // 更新UI
                dataGridView1.DataSource = currentData;

                // 設置行背景色
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if ((dataGridView1.Rows[i].Cells["Amount"].Value?.ToString() ?? string.Empty) == "0")
                    {
                        dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.MediumVioletRed;
                    }
                }

                // 更新頁碼顯示
                if (lblCurrentPage != null)
                {
                    lblCurrentPage.Text = $"第 {currentPage} 頁";
                    lblCurrentPage.Visible = true;
                }

                // 顯示分頁按鈕
                btnPreviousPage.Visible = true;
                btnNextPage.Visible = true;

                // 顯示目前頁數/總頁數
                if (lblPageInfo != null)
                {
                    lblPageInfo.Text = $"第 {currentPage} 頁 / 共 {totalPages} 頁";
                }
                sumC();

                return currentData;
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲取儲位查詢分頁數據時發生錯誤: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 上一頁
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPreviousPage_Click(object sender, EventArgs e)
        {
            if (currentPage <= 1)
            {
                MessageBox.Show("已經是第一頁");
                return;
            }

            currentPage--;

            // 根據查詢類型選擇不同的分頁方法
            if (queryType == 1) // 儲位查詢
            {
                GetPositionPagedData(currentPage);
            }
            else if (queryType == 2) // 料號查詢
            {
                GetItemPagedData(currentPage);
            }
            else // 預設查詢
            {
                GetPagedData(currentPage);
            }
        }

        /// <summary>
        /// 獲取料號+儲位查詢的分頁數據
        /// </summary>
        /// <param name="sourceData">源數據表</param>
        /// <param name="page">頁碼</param>
        /// <returns>分頁後的數據表</returns>
        private DataTable GetItemPositionPagedData(DataTable sourceData, int page)
        {
            try
            {
                if (sourceData == null || sourceData.Rows.Count == 0)
                {
                    return sourceData;
                }

                // 計算總頁數
                int totalRows = sourceData.Rows.Count;
                int totalPages = (int)Math.Ceiling(totalRows / (double)pageSize);

                // 計算當前頁的起始和結束索引
                int startIndex = (page - 1) * pageSize;
                int endIndex = Math.Min(startIndex + pageSize - 1, totalRows - 1);

                // 創建新的數據表來存儲分頁數據
                DataTable pagedData = sourceData.Clone();

                // 複製指定範圍的行到新數據表
                for (int i = startIndex; i <= endIndex; i++)
                {
                    pagedData.ImportRow(sourceData.Rows[i]);
                }

                // 更新頁碼信息
                if (lblPageInfo != null)
                {
                    lblPageInfo.Text = $"第 {page} 頁 / 共 {totalPages} 頁";
                }

                return pagedData;
            }
            catch (Exception ex)
            {
                MessageBox.Show("獲取料號+儲位查詢分頁數據時發生錯誤: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// 下一頁
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnNextPage_Click(object sender, EventArgs e)
        {
            // 根據查詢類型選擇不同的分頁方法
            if (queryType == 1) // 儲位查詢
            {
                // 檢查是否有更多數據，以決定是否允許繼續翻頁
                DataTable nextPageData = GetPositionPagedData(currentPage + 1, true);
                if (nextPageData != null && nextPageData.Rows.Count > 0)
                {
                    currentPage++;
                    GetPositionPagedData(currentPage);
                }
                else
                {
                    MessageBox.Show("已經是最後一頁");
                }
            }
            else if (queryType == 2) // 料號查詢
            {
                // 檢查是否有更多數據，以決定是否允許繼續翻頁
                DataTable nextPageData = GetItemPagedData(currentPage + 1, true);
                if (nextPageData != null && nextPageData.Rows.Count > 0)
                {
                    currentPage++;
                    GetItemPagedData(currentPage);
                }
                else
                {
                    MessageBox.Show("已經是最後一頁");
                }
            }
            else // 預設查詢
            {
                // 檢查是否有更多數據，以決定是否允許繼續翻頁
                DataTable nextPageData = GetPagedData(currentPage + 1, true);
                if (nextPageData != null && nextPageData.Rows.Count > 0)
                {
                    currentPage++;
                    GetPagedData(currentPage);
                }
                else
                {
                    MessageBox.Show("已經是最後一頁");
                }
            }
        }
        #endregion
    }
}
