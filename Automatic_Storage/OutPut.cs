using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace Automatic_Storage
{
    /// <summary>
    /// 出庫登記
    /// </summary>
    public partial class OutPut : Form
    {

        /// <summary>
        /// 出庫登記視窗的建構函式。
        /// </summary>
        /// <remarks>
        /// 初始化元件並設定 isLoaded 為 false，表示尚未載入控制項尺寸資料。
        /// </remarks>
        public OutPut()
        {
            InitializeComponent(); // 初始化視窗元件
            isLoaded = false;      // 尚未設定控制項尺寸資料到 Tag 屬性
        }

        /// <summary>
        /// 表示下拉選單中的包裝種類項目。
        /// </summary>
        /// <remarks>
        /// 此類別用於儲存包裝顯示文字與對應的值，並在 ComboBox 控制項中顯示。
        /// </remarks>
        /// <example>
        /// <code>
        /// MyItem item = new MyItem("紙箱", "A01");
        /// cbx_package.Items.Add(item);
        /// </code>
        /// </example>
        public class MyItem
        {
            /// <summary>
            /// 顯示於下拉選單的文字。
            /// </summary>
            public string text; // 顯示文字

            /// <summary>
            /// 對應的包裝代碼值。
            /// </summary>
            public string value; // 對應值

            /// <summary>
            /// 建立 MyItem 物件並指定顯示文字與值。
            /// </summary>
            /// <param name="text">顯示於下拉選單的文字。</param>
            /// <param name="value">對應的包裝代碼值。</param>
            public MyItem(string text, string value)
            {
                this.text = text; // 設定顯示文字
                this.value = value; // 設定對應值
            }

            /// <summary>
            /// 傳回顯示文字，供下拉選單顯示用。
            /// </summary>
            /// <returns>顯示文字。</returns>
            public override string ToString()
            {
                return text; // 回傳顯示文字
            }
        }

        #region 批次出庫來的參數
        /// <summary>
        /// 料號
        /// </summary>
        private string strItem;     //料號
        /// <summary>
        /// 規格
        /// </summary>
        private string strSpec;     //規格
        /// <summary>
        /// 包裝
        /// </summary>
        private string strPackage;  //包裝
        /// <summary>
        /// 單號
        /// </summary>
        private string strWono = "";  //單號
        /// <summary>
        /// 是否批次來
        /// </summary>
        private bool strBat = false;  //是否批次來

        /// <summary>
        /// 料號
        /// </summary>
        public string Item
        {
            set { strItem = value; }
        }
        /// <summary>
        /// 規格
        /// </summary>
        public string Spec
        {
            set { strSpec = value; }
        }
        /// <summary>
        /// 包裝
        /// </summary>
        public string Package
        {
            set { strPackage = value; }
        }
        /// <summary>
        /// 單號
        /// </summary>
        public string Wono
        {
            set { strWono = value; }
        }
        /// <summary>
        /// 是否批次來
        /// </summary>
        public bool Bat
        {
            set { strBat = value; }
        }
        /// <summary>
        /// 設定視窗欄位的值，將批次出庫的參數帶入對應控制項。
        /// </summary>
        /// <remarks>
        /// 此方法會將 strItem、strSpec、strWono、strPackage、strBat 等欄位值設定到對應的控制項或變數。
        /// </remarks>
        public void setValue()
        {
            textBox1.Text = strItem; // 設定料號到 textBox1
            //textBox2.Text = strItem; // (註解) 可選擇設定料號到 textBox2
            string sSpec = strSpec; // 將規格存到區域變數 sSpec
            string sWono = strWono; // 將單號存到區域變數 sWono
            cbx_package.SelectedItem = strPackage; // 設定包裝種類到下拉選單
            bool sBat = strBat; // 將是否批次來存到區域變數 sBat
        }
        #endregion
        #region 視窗ReSize
        /// <summary>
        /// 窗口寬度
        /// </summary>
        int X = new int();  //窗口寬度
        /// <summary>
        /// 窗口高度
        /// </summary>
        int Y = new int(); //窗口高度
        /// <summary>
        /// 寬度縮放比例
        /// </summary>
        float fgX = new float(); //寬度縮放比例
        /// <summary>
        /// 高度縮放比例
        /// </summary>
        float fgY = new float(); //高度縮放比例
        /// <summary>
        /// 是否已設定各控制的尺寸資料到Tag屬性
        /// </summary>
        // 是否已設定各控制的尺寸資料到Tag屬性
        bool isLoaded;  // 是否已設定各控制的尺寸資料到Tag屬性
        #endregion

        /// <summary>
        /// 用於處理 DateTime? 型別的 Nullable 轉換器。
        /// </summary>
        /// <remarks>
        /// 主要用於將控制項資料轉換為可為 null 的日期時間型別。
        /// </remarks>
        System.ComponentModel.NullableConverter nullableDateTime =
            new System.ComponentModel.NullableConverter(typeof(DateTime?)); // DateTime? 型別的 Nullable 轉換器

        /// <summary>
        /// 實際入庫日期字串。
        /// </summary>
        /// <remarks>
        /// 用於記錄目前選取資料的入庫日期，格式為 yyyy-MM-dd。
        /// </remarks>
        string actualDate = string.Empty; // 實際入庫日期

        /// <summary>
        /// 資料的序號欄位。
        /// </summary>
        /// <remarks>
        /// 用於記錄目前選取資料的序號。
        /// </remarks>
        string sno = string.Empty; // 資料序號

        /// <summary>
        /// 庫存數量。
        /// </summary>
        /// <remarks>
        /// 用於記錄目前選取資料的庫存數量。
        /// </remarks>
        double inventory = 0; // 庫存數量

        /// <summary>
        /// 出庫數量。
        /// </summary>
        /// <remarks>
        /// 用於記錄本次出庫的數量。
        /// </remarks>
        double Amount = 0; // 出庫數量

        /// <summary>
        /// 料號重複確認用字串。
        /// </summary>
        /// <remarks>
        /// 用於扣帳時確認料號是否一致。
        /// </remarks>
        string txt_dulCheck = string.Empty; // 料號重複確認用

        /// <summary>
        /// 搜尋用料號字串。
        /// </summary>
        /// <remarks>
        /// 用於查詢資料時的料號關鍵字。
        /// </remarks>
        string txt_search = string.Empty; // 搜尋用料號

        /// <summary>
        /// 料號拆解後的字串。
        /// </summary>
        /// <remarks>
        /// 用於儲存拆解後的料號資訊。
        /// </remarks>
        string strShift = string.Empty; // 拆解後料號

        /// <summary>
        /// 資料表物件，用於儲存查詢結果。
        /// </summary>
        /// <remarks>
        /// 出庫查詢、資料綁定等都會用到此 DataTable。
        /// </remarks>
        DataTable dt = new DataTable(); // 查詢結果資料表

        /// <summary>
        /// 出庫日期時間字串。
        /// </summary>
        /// <remarks>
        /// 格式為 yyyy-MM-dd HH:mm:ss，記錄本次出庫的時間。
        /// </remarks>
        string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // 出庫日期時間

        /// <summary>
        /// 昶亨料號 onClick
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                label10.Text = "";
                try
                {
                    if (!string.IsNullOrEmpty(textBox1.Text))
                    {
                        txt_search = textBox1.Text.Trim().ToUpper();
                    }
                    else if (!string.IsNullOrEmpty(txt_itemC.Text))
                    {
                        txt_search = txt_itemC.Text.Trim().ToUpper();
                    }

                    //subShift(txt_search);
                    //strShift = (sub == "-") ? StringSplit.StrLeft(txt_search, txt_search.Length - 3) : txt_search;

                    dt = dataBind();
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message);
                }
                if (dt.Rows.Count > 0)
                {
                    dataGridView1.DataSource = dt;
                    dataGridView1.Visible = true;
                    label5.Visible = false;
                    textBox2.Focus();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if ((dt.Rows[i]["Amount"]?.ToString() ?? string.Empty) == "0")
                        {
                            dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.MediumVioletRed;
                        }
                    }

                }
                else
                {
                    label5.Font = new Font("微軟正黑體", 18, FontStyle.Bold);
                    label5.ForeColor = Color.Red;
                    label5.Text = "找不到資料";
                }
            }
        }

        /// <summary>
        /// 視窗載入事件，初始化控制項尺寸與包裝種類下拉選單。// {
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        /// <remarks>
        /// 此方法會記錄視窗初始寬高，並將所有控制項的尺寸資訊暫存到 Tag 屬性。
        /// 同時會從資料庫載入所有包裝種類，並加入至 cbx_package 下拉選單。
        /// 若為批次出庫，則會依據批次參數查詢資料並綁定至 dataGridView1。
        /// </remarks>
        private void OutPut_Load(object sender, EventArgs e)
        {
            #region 獲取窗體info
            X = this.Width; // 取得窗體寬度
            Y = this.Height; // 取得窗體高度
            isLoaded = true; // 標記已設定各控制項尺寸到 Tag 屬性
            SetTag(this); // 遞迴暫存所有控制項的尺寸資訊到 Tag 屬性
            #endregion

            // Form_Load 設定包裝種類 cbox
            // textBox3.Visible = false; // 預設隱藏 textBox3
            string strsql = "select Package_View,code from Automatic_Storage_Package"; // 查詢所有包裝種類
            DataSet cbx_ds = db.ExecuteDataSet(strsql, CommandType.Text, null); // 執行 SQL 並取得資料集
            foreach (DataRow dr in cbx_ds.Tables[0].Rows)
            {
                // 將每一筆包裝資料加入 cbx_package 下拉選單
                cbx_package.Items.Add(new MyItem(dr["Package_View"]?.ToString() ?? string.Empty, dr["code"]?.ToString() ?? string.Empty));
            }
            // 判斷是否來自批次視窗
            if (strBat)
            {
                // 若為批次出庫，依據批次參數查詢出庫明細
                string sqlstr = "select Actual_InDate,Item_No_Master ,Item_No_Slave ,Amount ,Position ,Package,Spec  ,Mark ,Sno ,Reel_ID,PCB_DC,CMC_DC  " +
                            "from Automatic_Storage_Detail " +
                            "where Unit_No = @unitNo and Item_No_Master = @Item and Spec = @Spec " +
                            "and Amount > 0  order by position asc";
                SqlParameter[] parm = new SqlParameter[]
                {
                new SqlParameter("unitNo",Login.Unit_No), // 單位編號參數
                new SqlParameter("Item",strItem),         // 料號參數
                new SqlParameter("Spec",strSpec)          // 規格參數
                };
                dataGridView1.Visible = true; // 顯示資料表
                                              // textBox3.Visible = true; // (註解) 可選擇顯示 textBox3
                                              // bat_confirm.Visible = true; // (註解) 可選擇顯示批次確認按鈕
                textBox4.Text = ""; // 清空出庫數量欄位
                                    // txt_Amount_U.Text = ""; // (註解) 清空單位數欄位
                dt = db.ExecuteDataTable(sqlstr, CommandType.Text, parm); // 執行查詢並取得資料表
                dataGridView1.DataSource = dt; // 綁定查詢結果至資料表
            }
        }

        /// <summary>
        /// 料號確認狀態碼，預設為 -1 表示尚未確認。
        /// </summary>
        /// <remarks>
        /// cp 用於判斷料號是否正確，0 表示正確，-1 表示尚未確認或錯誤。
        /// </remarks>
        int cp = -1; // 料號確認狀態碼，預設為 -1

        //↓料號確認
        /// <summary>
        /// 處理 textBox2 的按鍵事件，判斷是否按下 Enter 鍵並執行出庫流程。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">按鍵事件參數。</param>
        /// <remarks>
        /// 當非批次出庫時，會依據輸入的料號或客戶料號執行扣帳，並重設相關欄位與重新查詢資料。
        /// 若為批次出庫則直接呼叫批次確認程序。
        /// </remarks>
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否按下 Enter 鍵且不是批次出庫
            if (e.KeyChar == 13 && !strBat)
            {
                // 若料號欄位有值則取料號
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    txt_dulCheck = textBox1.Text.Trim().ToUpper();
                }
                // 若客戶料號欄位有值則取客戶料號
                if (!string.IsNullOrEmpty(txt_itemC.Text))
                {
                    txt_dulCheck = txt_itemC.Text.Trim().ToUpper();
                }
                // 執行扣帳程序，若失敗則直接返回
                if (!Debit()) return;

                // 重設所有欄位並重新查詢資料
                textBox1.Focus();           // 料號欄位取得焦點
                textBox1.Text = "";         // 清空料號
                textBox2.Text = "";         // 清空客戶料號
                textBox3.Text = "";         // 清空儲位
                txt_Mark.Text = "";         // 清空備註
                                            //txt_Amount_U.Text = "";   // (註解) 清空單位數
                textBox4.Text = "";         // 清空出庫數量
                label10.Text = "";          // 清空提示訊息
                cbx_package.SelectedIndex = -1; // 重設包裝種類選擇
                dataGridView1.DataSource = dataBind(); // 重新查詢並綁定資料
            }
            // 判斷是否按下 Enter 鍵且為批次出庫
            if (e.KeyChar == 13 && strBat)
            {
                bat_confirm_Click(sender, e); // 執行批次確認程序
            }

        }

        /// <summary>
        /// 依據料號確認狀態碼 cp，更新儲位確認相關控制項狀態與提示文字。
        /// </summary>
        /// <remarks>
        /// 當 cp 為 0 時，表示料號確認正確，則將焦點移至出庫數量欄位；否則顯示錯誤訊息並隱藏儲位欄位。
        /// </remarks>
        public void lb4Change()
        {
            switch (cp)
            {
                case 0:
                    // cp 為 0，表示料號確認正確
                    if (dt.Rows.Count > 0)
                    {
                        // 若查詢結果有資料，將焦點移至出庫數量欄位
                        //label4.Font = new Font("微軟正黑體", 18, FontStyle.Bold);//微軟正黑體, 12pt, style=Bold
                        //label4.ForeColor = Color.Black;
                        //label4.Text = "儲位確認";
                        //textBox3.Visible = true;
                        textBox4.Focus(); // 將焦點移至 textBox4
                    }
                    else
                    {
                        // 若查詢結果無資料，清空料號與客戶料號欄位並將焦點移回料號欄位
                        textBox1.Text = ""; // 清空料號欄位
                        textBox2.Text = ""; // 清空客戶料號欄位
                        textBox1.Focus();   // 將焦點移回料號欄位
                    }
                    break;
                default:
                    // cp 非 0，表示料號確認錯誤
                    textBox3.Visible = false; // 隱藏儲位欄位
                    label4.Font = new Font("微軟正黑體", 18, FontStyle.Bold); // 設定提示文字字型
                    label4.ForeColor = Color.Red; // 設定提示文字顏色為紅色
                    label4.Text = "料卷有誤,請再次確認!!"; // 顯示錯誤訊息
                    break;
            }
        }

        /// <summary>
        /// 處理 textBox3 的按鍵事件，判斷是否按下 Enter 鍵並執行料號比對。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">按鍵事件參數。</param>
        /// <remarks>
        /// 當按下 Enter 鍵時，會將 textBox2 的文字轉成大寫後與 textBox1 或 txt_itemC 的文字進行比對，
        /// 並將比對結果存到 cp 變數，供後續料號確認流程使用。
        /// </remarks>
        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否按下 Enter 鍵 (KeyChar 為 13)
            if (e.KeyChar == 13)
            {
                // 取得 textBox2 的文字並轉成大寫，作為比對用字串
                string auditText = textBox2.Text.Trim().ToUpper();
                // 若 textBox1 有值，則與 auditText 進行比對
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    // 比對 textBox1 的文字與 auditText，結果存到 cp
                    cp = textBox1.Text.Trim().ToUpper().CompareTo(auditText);
                }
                // 若 txt_itemC 有值，則與 auditText 進行比對
                else if (!string.IsNullOrEmpty(txt_itemC.Text))
                {
                    // 比對 txt_itemC 的文字與 auditText，結果存到 cp
                    cp = txt_itemC.Text.Trim().ToUpper().CompareTo(auditText);
                }

                // 比對結果 cp 可供後續料號確認流程使用
                //if (cp == 0)
                //{
                //    lb4Change();
                //    //this.SelectNextControl(this.ActiveControl, true, true, true, false); //跳下一個元件
                //}
                //lb4Change();
            }
        }

        /// <summary>
        /// 執行出庫扣帳流程，包含資料驗證、資料庫新增出庫紀錄與更新庫存明細。
        /// </summary>
        /// <remarks>
        /// 此方法會依據目前輸入的料號、數量、包裝等資訊，進行出庫資料的新增與庫存明細的更新。
        /// 若資料驗證失敗或資料庫操作失敗則回傳 false。
        /// </remarks>
        /// <returns>
        /// 若出庫成功則回傳 true，否則回傳 false。
        /// </returns>
        public bool Debit()
        {
            // 檢查 textBox1 是否有值，若有則將料號存到 txt_dulCheck
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                txt_dulCheck = textBox1.Text.Trim().ToUpper();
            }
            // 檢查 txt_itemC 是否有值，若有則將客戶料號存到 txt_dulCheck
            if (!string.IsNullOrEmpty(txt_itemC.Text))
            {
                txt_dulCheck = txt_itemC.Text.Trim().ToUpper();
            }
            // 檢查出庫數量是否大於庫存數量，若超過則顯示錯誤訊息並返回 false
            if (Amount > inventory)
            {
                label10.Text = "總數量有誤,請重新確認!";
                return false;
            }
            // 檢查 textBox4 輸入數量是否大於庫存數量，若超過則顯示錯誤訊息並返回 false
            if (Convert.ToDouble(string.IsNullOrWhiteSpace(textBox4.Text) ? "0" : textBox4.Text) > inventory)
            {
                label10.Text = "總數量有誤,請重新確認!";
                return false;
            }
            // 檢查料號是否與 textBox2 輸入一致，若不一致則顯示錯誤訊息並返回 false
            if (txt_dulCheck != textBox2.Text.Trim().ToUpper())
            {
                label10.Text = "料號確認有誤,請重新確認!";
                return false;
            }
            try
            {
                // 取得目前選取的包裝種類 MyItem 物件
                MyItem myItem = (MyItem)this.cbx_package.SelectedItem;
                // 建立出庫資料新增 SQL 指令
                string sqlOut = @"insert into Automatic_Storage_Output (Sno,Item_No_Master,Item_No_Slave,Spec,Position,
                                   Amount,Package,Unit_No,Output_UserNo,Output_Date,Wo_No,Mark,Reel_ID,PCB_DC,CMC_DC ) 
                                   select Sno,Item_No_Master,Item_No_Slave,Spec,Position,@amount,Package,
                                   @unitno,@outuser,@outdate,@wono,@mark,@reelid,PCB_DC,CMC_DC from Automatic_Storage_Detail ";

                // 加上條件只處理 amount > 0 的資料
                sqlOut += @"where Sno=@sno and amount > 0";

                // 建立出庫資料新增所需的參數陣列
                SqlParameter[] parm1 = new SqlParameter[]
                {
                new SqlParameter("amount",textBox4.Text),
                new SqlParameter("unitno",Login.Unit_No),
                new SqlParameter("outuser",Login.User_No),
                new SqlParameter("outdate",dateTime),
                new SqlParameter("wono",strWono),
                new SqlParameter("mark",txt_Mark.Text.Trim()),
                new SqlParameter("reelid",txt_reelid.Text.Trim().ToUpper()),
                new SqlParameter("sno",sno.Trim()),
                };

                // 執行出庫資料新增，若失敗則重設 cp 狀態並呼叫 lb4Change，最後返回 false
                if (db.ExecueNonQuery(sqlOut, CommandType.Text, "單筆出庫text確認", parm1) == 0)
                {
                    cp = -1;
                    lb4Change();
                    return false;
                }

                // 建立庫存明細更新 SQL 指令
                string sqlDetail = @"update Automatic_Storage_Detail 
                                   set Amount = Amount-@amount ,
                                   Up_OutDate = @outdate ,Output_UserNo = @outuser ,Mark = @mark ,Reel_ID = @reelid ";

                // 若 textBox1 有值則以 Item_No_Master 為條件
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    sqlDetail += @"where Item_No_Master = @master and Position=@position and Package=@package 
                              and Actual_InDate=@actualDate and sno=@sno ";
                }

                // 若 txt_itemC 有值則以 Item_No_Slave 為條件
                if (!string.IsNullOrEmpty(txt_itemC.Text))
                {
                    sqlDetail += @"where Item_No_Slave = @master and Position=@position and Package=@package 
                            and Actual_InDate=@actualDate and sno=@sno";
                }

                // 建立庫存明細更新所需的參數陣列
                SqlParameter[] parm2 = new SqlParameter[]
                {
                new SqlParameter("amount",textBox4.Text.Trim()),
                new SqlParameter("outdate",dateTime),
                new SqlParameter("outuser",Login.User_No),
                new SqlParameter("mark",txt_Mark.Text.Trim().ToUpper()),
                new SqlParameter("reelid",txt_reelid.Text.Trim().ToUpper()),
                new SqlParameter("master",textBox2.Text.Trim()),
                new SqlParameter("position",textBox3.Text.Trim()),
                new SqlParameter("package",myItem.value),
                new SqlParameter("actualDate",actualDate),
                new SqlParameter("sno",sno)
                };

                // 執行庫存明細更新
                db.ExecueNonQuery(sqlDetail, CommandType.Text, "單筆出庫Detail更新", parm2);
            }
            catch (Exception ex)
            {
                // 若發生例外則顯示錯誤訊息
                MessageBox.Show(ex.Message);
            }
            // 出庫流程成功，回傳 true
            return true;
        }


        /// <summary>
        /// 查詢出庫明細資料，依據輸入的料號或客戶料號回傳對應的 DataTable。
        /// </summary>
        /// <remarks>
        /// 此方法會根據 textBox1 或 txt_itemC 的內容，組合 SQL 查詢語句，
        /// 並以參數方式帶入單位編號與料號，回傳查詢結果的 DataTable。
        /// </remarks>
        /// <returns>
        /// 回傳查詢結果的 DataTable，包含出庫明細資料。
        /// </returns>
        /// <example>
        /// <code>
        /// DataTable dt = dataBind();
        /// </code>
        /// </example>
        public DataTable dataBind()
        {
            // 建立查詢出庫明細的 SQL 語句
            string sqlstr = "select Actual_InDate,Item_No_Master ,Item_No_Slave ,Amount ,Position ,Package ,Spec ,Mark ,Sno ,Reel_ID,PCB_DC,CMC_DC  " +
                    "from Automatic_Storage_Detail ";
            // 如果 textBox1 有輸入料號，則以 Item_No_Master 為條件查詢
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                sqlstr += "where Unit_No = @unitNo and Item_No_Master = @Item " +
                      "and Amount > 0  order by Actual_InDate asc";
            }
            // 如果 txt_itemC 有輸入客戶料號，則以 Item_No_Slave 為條件查詢
            else if (!string.IsNullOrEmpty(txt_itemC.Text))
            {
                sqlstr += "where Unit_No = @unitNo and Item_No_Slave = @Item " +
                      "and Amount > 0  order by Actual_InDate asc";
            }

            // 建立 SQL 查詢參數，包含單位編號與料號
            SqlParameter[] parm = new SqlParameter[]
            {
            new SqlParameter("unitNo",Login.Unit_No), // 單位編號參數
            new SqlParameter("Item",txt_search)       // 料號參數
            };
            // 執行 SQL 查詢並回傳結果 DataTable
            return db.ExecuteDataTable(sqlstr, CommandType.Text, parm);
        }

        /// <summary>
        /// 視窗關閉事件，當視窗即將關閉時執行。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">視窗關閉事件參數。</param>
        /// <remarks>
        /// 若視窗標題為 "Form1"，則呼叫主視窗的 refreshData 方法以重新整理資料。
        /// </remarks>
        private void OutPut_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 判斷視窗標題是否為 "Form1"
            if (this.Text == "Form1")
            {
                // 取得主視窗物件並呼叫 refreshData 方法重新整理資料
                Form1 ower = (Form1)this.Owner;
                ower.refreshData();
            }
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
        /// 根據視窗縮放比例，重新設定所有控制項的寬度、高度、位置與字體大小。
        /// </summary>
        /// <param name="newx">寬度縮放比例。</param>
        /// <param name="newy">高度縮放比例。</param>
        /// <param name="cons">要調整的父控制項。</param>
        /// <remarks>
        /// 此方法會遞迴遍歷所有子控制項，根據 Tag 屬性暫存的原始尺寸資訊，依照目前縮放比例調整控制項的大小與位置。
        /// </remarks>
        /// <example>
        /// <code>
        /// SetControls(fgX, fgY, this);
        /// </code>
        /// </example>
        private void SetControls(float newx, float newy, Control cons)
        {
            // 判斷是否已載入控制項尺寸資料
            if (isLoaded)
            {
                // 遍歷父控制項中的所有子控制項
                foreach (Control con in cons.Controls)
                {
                    // 取得控制項的 Tag 屬性值並以 ':' 分割成字串陣列
                    string[] mytag = con.Tag.ToString().Split(new char[] { ':' });
                    // 根據縮放比例計算新的寬度
                    float a = System.Convert.ToSingle(mytag[0]) * newx;
                    con.Width = (int)a; // 設定控制項寬度
                                        // 根據縮放比例計算新的高度
                    a = System.Convert.ToSingle(mytag[1]) * newy;
                    con.Height = (int)(a); // 設定控制項高度
                                           // 根據縮放比例計算新的左邊距
                    a = System.Convert.ToSingle(mytag[2]) * newx;
                    con.Left = (int)(a); // 設定控制項左邊距
                                         // 根據縮放比例計算新的上邊緣距
                    a = System.Convert.ToSingle(mytag[3]) * newy;
                    con.Top = (int)(a); // 設定控制項頂邊距
                                        // 根據縮放比例計算新的字體大小
                    Single currentSize = System.Convert.ToSingle(mytag[4]) * newy;
                    con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit); // 設定字體大小
                                                                                                    // 若該控制項還有子控制項則遞迴呼叫本方法
                    if (con.Controls.Count > 0)
                    {
                        SetControls(newx, newy, con);
                    }
                }
            }
        }
        /// <summary>
        /// 視窗縮放事件，根據目前視窗大小重新調整所有控制項的尺寸與位置。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        /// <remarks>
        /// 此方法會根據原始視窗寬高，計算目前縮放比例，並呼叫 SetControls 方法遞迴調整所有控制項的寬度、高度、位置與字體大小。
        // <summary>
        /// <example>
        /// <code>
        /// // 在 Form 的 Resize 事件中呼叫
        /// this.Resize += OutPut_Resize;
        /// </code>
        /// </example>
        private void OutPut_Resize(object sender, EventArgs e)
        {
            // 若原始寬高為 0，則不執行縮放
            if (X == 0 || Y == 0) return;
            // 計算目前寬度縮放比例
            fgX = (float)this.Width / (float)X;
            // 計算目前高度縮放比例
            fgY = (float)this.Height / (float)Y;

            // 根據縮放比例調整所有控制項
            SetControls(fgX, fgY, this);
        }
        ///
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>/ </remarks>
        private void OutPut_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        /// <summary>
        /// 資料綁定完成事件，設定 DataGridView 欄位標題與顯示狀態。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">資料綁定完成事件參數。</param>
        /// <remarks>
        /// 此方法會依據資料表的欄位順序，設定每個欄位的標題文字，並將 Sno 欄位隱藏。
        /// </remarks>
        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView1.Columns[0].HeaderText = "入庫日期";                //Actual_InDate
            dataGridView1.Columns[1].HeaderText = "昶亨料號";                //Item_E
            dataGridView1.Columns[2].HeaderText = "客戶料號";                //Item_C
            dataGridView1.Columns[3].HeaderText = "總數";                         //Amount
            dataGridView1.Columns[4].HeaderText = "儲位";                         //Position
            dataGridView1.Columns[5].HeaderText = "包裝種類";                 //Package
            dataGridView1.Columns[6].HeaderText = "規格";                         //Spec
            dataGridView1.Columns[7].HeaderText = "備註";                         //Mark
            dataGridView1.Columns[8].HeaderText = "Sno";                          //Package
            dataGridView1.Columns[8].Visible = false;                                     //sno
            dataGridView1.Columns[9].HeaderText = "Reel_ID";                          //Package
            dataGridView1.Columns[10].HeaderText = "PCB";                          //Package
            dataGridView1.Columns[11].HeaderText = "CMC";
        }

        /// <summary>
        /// 處理 textBox4 的按鍵事件，僅允許數字與退格鍵輸入，並在按下 Enter 鍵時執行相關邏輯。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 textBox4。</param>
        /// <param name="e">按鍵事件參數。</param>
        /// <remarks>
        /// 此方法會檢查輸入是否為數字或退格鍵，若不是則阻止輸入。
        /// 當按下 Enter 鍵時，若 textBox4 有值則將焦點移至包裝種類下拉選單。
        /// 並檢查出庫數量是否超過庫存，若超過則顯示錯誤訊息。
        /// </remarks>
        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 僅允許數字(0-9)與退格鍵(8)輸入
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                // 阻止非數字與非退格鍵的輸入
                e.Handled = true;
                // 按下 Enter 鍵(13)時執行
                if (e.KeyChar == 13)
                {
                    // 若 textBox4 有值則將焦點移至包裝種類下拉選單
                    TextBox txb = (TextBox)sender;
                    // 根據觸發事件的控制項名稱執行對應邏輯
                    switch (txb.Name)
                    {
                        case "textBox4":
                            // 若 textBox4 有值則將焦點移至包裝種類下拉選單
                            if (!string.IsNullOrEmpty(textBox4.Text))
                            {
                                //包裝數x單位數.焦點
                                cbx_package.Focus();
                            }
                            break;
                        //case "txt_Amount_U":
                        //    if (!string.IsNullOrEmpty(txt_Amount_U.Text) &&
                        //        Convert.ToDouble(txt_Amount_U.Text) <= Convert.ToDouble(dt.Rows[0]["Amount_Unit"]))
                        //    {
                        //        //包裝數x單位數
                        //        double xNum = Convert.ToDouble(textBox4.Text);
                        //        Amount = xNum * Convert.ToDouble(txt_Amount_U.Text);
                        //        cbx_package.Focus();
                        //    }
                        //    else
                        //    {
                        //        label10.Text = "單位數量有誤,請重新確認!";
                        //    }
                        //    break;
                        default:
                            break;
                    }
                    // 檢查出庫數量是否超過庫存，若超過則顯示錯誤訊息
                    if (Amount > Convert.ToDouble(dt.Rows[0]["Amount"]?.ToString() ?? "0"))
                    {
                        label10.Text = "總數量有誤,請重新確認!";
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// 處理 textBox4 離開事件，檢查出庫數量並設定焦點至包裝種類下拉選單。
        /// </summary>
        /// <param name="sender">觸發事件的控制項物件。</param>
        /// <param name="e">事件參數。</param>
        /// <remarks>
        /// 當 textBox4 離開時，若有輸入數量則將焦點移至包裝種類下拉選單。
        /// 並檢查出庫數量是否超過庫存，若超過則顯示錯誤訊息。
        /// </remarks>
        private void textBox4_Leave(object sender, EventArgs e)
        {
            // 若 textBox4 有值則將焦點移至包裝種類下拉選單
            TextBox txb = (TextBox)sender;
            // 根據觸發事件的控制項名稱執行對應邏輯
            switch (txb.Name)
            {
                case "textBox4":
                    if (!string.IsNullOrEmpty(textBox4.Text))
                    {
                        //包裝數x單位數.焦點
                        cbx_package.Focus();
                    }
                    break;
                //case "txt_Amount_U":
                //    if (!string.IsNullOrEmpty(txt_Amount_U.Text) &&
                //                Convert.ToDouble(txt_Amount_U.Text) <= Convert.ToDouble(dt.Rows[0]["Amount_Unit"]))
                //    {
                //        //包裝數x單位數
                //        double xNum = Convert.ToDouble(textBox4.Text);
                //        Amount = xNum * Convert.ToDouble(txt_Amount_U.Text);
                //        //label10.Text = Amount.ToString();
                //        cbx_package.Focus();
                //    }
                //    else
                //    {
                //        label10.Text = "單位數量有誤,請重新確認!";
                //    }
                //    break;
                default:
                    break;
            }
            // 檢查出庫數量是否超過庫存，若超過則顯示錯誤訊息
            if (Amount > Convert.ToDouble(dt.Rows[0]["Amount"]?.ToString() ?? "0"))
            {
                label10.Text = "總數量有誤,請重新確認!";
                return;
            }
        }

        /// <summary>
        /// 批次出庫確認按鈕事件。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        /// <remarks>
        /// 此方法會執行出庫扣帳流程，若成功則將出庫數量與料號回傳給父視窗 BatOut，並關閉本視窗。
        /// </remarks>
        private void bat_confirm_Click(object sender, EventArgs e)
        {
            //[0]→ItemE, [1]→ItemC, [2]→單位數, [3]→總數, [4]→儲位, [5]→包裝, [6]→規格,
            BatOut father = (BatOut)this.Owner;
            //先確認msg.dialog=yes; 判斷是否父視窗啟動的程序
            //DialogResult res = MessageBox.Show(
            //    "確認出庫( "+ Environment.NewLine +
            //    "料號_昶 : " + dataGridView1.CurrentRow.Cells[1].Value.ToString() + Environment.NewLine +
            //    "出庫數量/Pcs : " + textBox4.Text + Environment.NewLine +
            //    //"總數 : " + txt_Amount_U.Text + Environment.NewLine +
            //    "儲位 : " + dataGridView1.CurrentRow.Cells[4].Value.ToString() + Environment.NewLine +
            //    "包裝 : " + dataGridView1.CurrentRow.Cells[5].Value.ToString() + " )",
            //                    "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            //if (res == DialogResult.Yes && strBat)
            //{
            //    if (!Debit()) return;
            //    father.MsgFromChildChk = textBox4.Text;
            //    father.MsgFromChildItemE = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            //    this.Close();
            //}
            //else
            //{
            //    return;
            //}

            //直接出庫
            if (!Debit()) return;
            //回傳出庫數量與料號
            father.MsgFromChildChk = textBox4.Text;
            father.MsgFromChildItemE = dataGridView1.CurrentRow.Cells[1].Value?.ToString() ?? string.Empty;
            father.MsgFromChildButton = true;
            this.Close();
        }

        /// <summary>
        /// DataGridView 雙擊事件，選取出庫明細資料並將相關欄位值帶入控制項。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 dataGridView1。</param>
        /// <param name="e">事件參數。</param>
        /// <remarks>
        /// 此事件會根據目前選取的資料列，將庫存數量、儲位、備註、入庫日期、序號等資訊帶入對應控制項。
        /// 並根據包裝代碼查詢包裝種類，設定下拉選單選項。
        /// 最後隱藏 visible_Panel 並將焦點移至出庫數量欄位。
        /// </remarks>
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            #region 先進先出管控
            //if (dataGridView1.CurrentRow.Index != 0)
            //{
            //    MessageBox.Show("Please，First In First Out！！");
            //    return;
            //}
            //for (int d = 1; d < dataGridView1.RowCount-1; d++)
            //{
            //    dataGridView1.Rows[d].DefaultCellStyle.BackColor = Color.Gray;
            //}
            #endregion
            //3→庫存數;4→儲位; 5→包裝
            /*
             0=actual_indate
             1=item_no_master
            2=item_no_slave
            3=amount
            4=position
            5=package
            6=spec
            7=mark
            8=sno
            9=reel_id
            10=pcb_dc
            11=cmc_dc
             */
            // 庫存數（保守 null 安全處理）
            var inventoryObj = dataGridView1.CurrentRow?.Cells[3].Value;
            inventory = inventoryObj != null ? Convert.ToDouble(inventoryObj) : 0.0;
            // 儲位
            textBox3.Text = dataGridView1.CurrentRow.Cells[4].Value?.ToString() ?? string.Empty;
            // 備註
            txt_Mark.Text = dataGridView1.CurrentRow.Cells[7].Value?.ToString() ?? string.Empty;
            // 入庫日期
            actualDate = Convert.ToDateTime(dataGridView1.CurrentRow?.Cells[0].Value?.ToString() ?? string.Empty).ToString("yyyy-MM-dd");

            sno = dataGridView1.CurrentRow?.Cells[8].Value?.ToString() ?? string.Empty;

            string strpackage = "select id from Automatic_Storage_Package " +
                "where code = '" + (dataGridView1.CurrentRow?.Cells[5].Value?.ToString() ?? string.Empty) + "'";
            DataSet dscpk = db.ExecuteDataSet(strpackage, CommandType.Text, null);
            cbx_package.SelectedIndex = Convert.ToInt32(dscpk.Tables[0].Rows[0]["id"]) - 1;

            visible_Panel.Visible = false;
            textBox4.Focus();
        }

        /// <summary>
        /// textBox1 文字變更事件，根據輸入狀態切換 txt_itemC 或 textBox1 的啟用狀態。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 textBox1。</param>
        /// <param name="e">事件參數。</param>
        /// <remarks>
        /// 當 textBox1 有輸入時，停用 txt_itemC；
        /// 當 txt_itemC 有輸入時，停用 textBox1。
        /// </remarks>
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // 根據 textBox1 與 txt_itemC 的輸入狀態切換啟用狀態
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                txt_itemC.Enabled = false;
            }
            else if (!string.IsNullOrEmpty(txt_itemC.Text))
            {
                textBox1.Enabled = false;
            }
        }

        /// <summary>
        /// 處理 OutPut 視窗關閉事件。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">視窗關閉事件參數。</param>
        /// <remarks>
        /// 若非批次出庫，則呼叫主視窗 Form1 的 refreshData 方法以重新整理資料。
        /// </remarks>
        private void OutPut_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!strBat)
            {
                Form1 ower = (Form1)this.Owner;
                ower.refreshData();
                //this.Close();
            }
        }
    }
}
