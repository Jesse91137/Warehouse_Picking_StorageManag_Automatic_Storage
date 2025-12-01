using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace Automatic_Storage
{
    /// <summary>
    /// Input 表單，負責入庫資料的操作與顯示。
    /// </summary>
    public partial class Input : Form
    {
        /// <summary>
        /// 建構子，初始化 Input 表單元件與狀態。
        /// </summary>
        public Input()
        {
            InitializeComponent(); // 初始化元件
            isLoaded = false; // 設定載入狀態為未載入
        }

        /// <summary>
        /// 自訂下拉選單項目類別，包含顯示文字與對應值。
        /// </summary>
        public class MyItem
        {
            /// <summary>
            /// 顯示於下拉選單的文字。
            /// </summary>
            public string text;
            /// <summary>
            /// 對應的值（通常為代碼或主鍵）。
            /// </summary>
            public string value;

            /// <summary>
            /// 建構子，初始化文字與值。
            /// </summary>
            /// <param name="text">顯示文字。</param>
            /// <param name="value">對應值。</param>
            public MyItem(string text, string value)
            {
                this.text = text; // 設定顯示文字
                this.value = value; // 設定對應值
            }
            /// <summary>
            /// 覆寫 ToString 方法，回傳顯示文字。
            /// </summary>
            /// <returns>顯示文字。</returns>
            public override string ToString()
            {
                return text; // 回傳顯示文字
            }
        }

        #region 窗體Size
        /// <summary>
        /// 窗口寬度。
        /// </summary>
        int X = new int();
        /// <summary>
        /// 窗口高度。
        /// </summary>
        int Y = new int();
        /// <summary>
        /// 寬度縮放比例。
        /// </summary>
        float fgX = new float();
        /// <summary>
        /// 高度縮放比例。
        /// </summary>
        float fgY = new float();
        /// <summary>
        /// 是否已設定各控制的尺寸資料到 Tag 屬性。
        /// </summary>
        bool isLoaded;
        #endregion

        /// <summary>
        /// 關閉按鈕事件，關閉表單並刷新主畫面資料。
        /// </summary>
        private void button2_Click(object sender, EventArgs e)
        {
            Form1 ower = (Form1)this.Owner;
            ower.refreshData();
            this.Close();
        }

        /// <summary>
        /// 處理料號輸入框的按鍵事件，判斷是否顯示 CMC 標籤並跳至數量欄位。
        /// </summary>
        private void txt_Item1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && !string.IsNullOrEmpty(ActiveControl?.Text))
            {
                try
                {
                    var text = txt_Item1.Text?.Trim();
                    if (!string.IsNullOrEmpty(text) && (text?.Length ?? 0) >= 4)
                    {
                        string auditCMC = text![(text!.Length - 4)].ToString().ToUpper();
                        if (auditCMC == "A")
                        {
                            labCMC.Visible = true;
                        }
                        else
                        {
                            labCMC.Visible = false;
                        }
                    }
                    else
                    {
                        labCMC.Visible = false;
                    }
                }
                catch (Exception ex)
                {
                    LogException(ex, "txt_Item1_KeyPress");
                    try { MessageBox.Show("料號處理發生錯誤，請聯絡系統管理員。"); } catch { }
                }

                // 將焦點移至儲位輸入框
                txt_Amount.Focus();
            }
        }

        /// <summary>
        /// 表單關閉事件，刷新主畫面資料。
        /// </summary>
        private void Input_FormClosing(object sender, FormClosingEventArgs e)
        {
            Form1 ower = (Form1)this.Owner;
            ower.refreshData();
        }

        /// <summary>
        /// 表單載入事件，初始化窗體尺寸與資料。
        /// </summary>
        private void Input_Load(object sender, EventArgs e)
        {
            #region 獲取窗體info
            X = this.Width;//獲取窗體的寬度
            Y = this.Height;//獲取窗體的高度
            isLoaded = true;// 已設定各控制項的尺寸到Tag屬性中
            SetTag(this);//調用方法
            #endregion

            init_Data();

            this.FormClosing += new FormClosingEventHandler(Input_FormClosing);
        }

        /// <summary>
        /// 初始化下拉選單資料來源。
        /// </summary>
        private async void init_Data()
        {
            /* 包裝種類 */
            string strsql = "select Package_View,code from Automatic_Storage_Package";
            try
            {
                // 顯示等待游標，提示使用者正在載入
                try { this.Cursor = Cursors.WaitCursor; } catch { }
                // 在背景執行查詢，避免在 UI 執行緒造成阻塞
                DataSet cbx_ds = await Task.Run(() => db.ExecuteDataSet(strsql, CommandType.Text, System.Array.Empty<SqlParameter>()));
                if (cbx_ds != null && cbx_ds.Tables.Count > 0)
                {
                    // 在 UI 執行緒新增選項
                    foreach (DataRow dr in cbx_ds.Tables[0].Rows)
                    {
                        cbx_package.Items.Add(new MyItem(dr["Package_View"].ToString(), dr["code"].ToString()));
                    }
                }
                try { this.Cursor = Cursors.Default; } catch { }
            }
            catch (Exception ex)
            {
                // 若發生例外，記錄或顯示錯誤，但不阻塞 UI
                try { MessageBox.Show("載入包裝資料失敗: " + ex.Message); } catch { }
            }
        }

        /// <summary>
        /// 清空所有輸入欄位與標籤，重設狀態。
        /// </summary>
        private void txt_clearn()
        {
            Label[] lbl = new Label[18] { lab_p1, lab_p2, lab_p3, lab_p4, lab_p5, lab_p6, lab_p7, lab_p8, lab_p9, lab_p10,
                                            lab_p11,lab_p12,lab_p13,lab_p14,lab_p15,lab_p16,lab_p17,lab_p18};
            foreach (var item in lbl)
            {
                item.Text = "";
            }
            foreach (Control ctrl in Controls)
            {
                if (ctrl is TextBox)
                {
                    ctrl.Text = "";
                }
            }
            cbx_package.SelectedIndex = -1;
            txt_Item1.Enabled = true;
            txt_item2.Enabled = true;
            txt_Item1.Focus();

        }

        /// <summary>
        /// 儲位明細-單筆入庫按鈕事件，執行入庫資料新增或更新。
        /// </summary>
        private void btn_only_commit_Click(object sender, EventArgs e)
        {
            string inputDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string sqlinput = string.Empty;
            string sqldetail = string.Empty;
            string sqlchkFirst = string.Empty;

            try
            {
                TextBox[] b = new TextBox[1] { txt_Storage1 }; //儲位
                /* 儲位表頭 */
                string strsql_p = @"select * from Automatic_Storage_Position where Position=@position and Unit_no=@unitno";
                SqlParameter[] parm_p = new SqlParameter[]
                {
                        new SqlParameter("position",b[0].Text.Trim().ToUpper()),
                        new SqlParameter("unitno",Login.Unit_No)
                };
                DataSet dsP = db.ExecuteDataSet(strsql_p, CommandType.Text, parm_p);

                if (dsP.Tables[0].Rows.Count > 0 && !string.IsNullOrEmpty(txt_Spec.Text))
                {
                    for (int i = 0; i < 1; i++)
                    {
                        TextBox[] c = new TextBox[1] { txt_Amount }; //總數量
                        TextBox[] d = new TextBox[1] { txt_Item1 }; //Item_Eversun
                        TextBox[] f = new TextBox[1] { txt_item2 }; //Item_Cutsome

                        if (b[0].Text.Trim() != "") //設定儲位才可以新增
                        {
                            if (d[i].Text.Trim().ToString() != "" || f[i].Text.Trim().ToString() != "") //有料號才insert
                            {
                                MyItem myItem = (MyItem)this.cbx_package.SelectedItem;
                                //重新設定數量
                                string amount = c[i].Text.Trim();//總數量
                                //string amount = "1"; ← 盤點前總數量為1的累加
                                //string amount_U = (txt_Amount_U.Text.Trim()); //單位
                                string master = (d[i].Text.Trim().ToUpper());//昶亨料號
                                string slave = (f[i].Text.Trim().ToUpper());//客戶料號
                                string item_MorS = (txt_Item1.Enabled) ? master : slave;

                                string package = myItem.value;
                                string spec = txt_Spec.Text.Trim();
                                string mark = txt_mark.Text.Trim();
                                string actualDate = dt_Picker.Value.ToString("yyyy/MM/dd");
                                string pcbdc = txt_pcbdc.Text.Trim();
                                string cmcdc = txt_cmcdc.Text.Trim();
                                //寫入input table
                                sqlinput = @"INSERT INTO Automatic_Storage_Input
                                        (Position,Amount,Item_No_Master,Item_No_Slave,Spec,Unit_No,Package,Input_Date,Actual_InDate,Input_UserNo,Mark,PCB_DC,CMC_DC)   
                                        values (@Position,@Amount,@Master,@Slave,@Spec,@Unit_No,@package,@InDate,@actualDate,@UserNo,@mark,@pcbdc,@cmcdc)";
                                SqlParameter[] parm = new SqlParameter[]
                                {
                                    new SqlParameter("Position",b[0].Text.Trim().ToUpper()),
                                    new SqlParameter("Amount",amount),
                                    //new SqlParameter("Amount_Unit",amount_U),
                                    new SqlParameter("Master",master),
                                    new SqlParameter("Slave",slave),
                                    new SqlParameter("Spec",spec),
                                    new SqlParameter("Package",package),
                                    new SqlParameter("Unit_No",Login.Unit_No),
                                    new SqlParameter("InDate",inputDate),
                                    new SqlParameter("actualDate",actualDate),
                                    new SqlParameter("UserNo",Login.User_No),
                                    new SqlParameter("mark",mark),
                                    new SqlParameter("pcbdc",pcbdc),
                                    new SqlParameter("cmcdc",cmcdc),
                                };
                                db.ExecueNonQuery(sqlinput, CommandType.Text, "btn_only_commit", parm);
                                #region 判斷該料號資料是否合併
                                //key : 入庫日期+料號+位置+包裝種類, if (key=新資料)  insert , else update
                                sqlchkFirst = @"select * from Automatic_Storage_Detail ";
                                if (txt_Item1.Enabled)
                                {
                                    sqlchkFirst += @"where Item_No_Master=@item_ms and Position=@Position 
                                                                and Package = @Package and Actual_InDate = @actualDate 
                                                                and Mark = @mark and PCB_DC = @pcbdc and CMC_DC = @cmcdc and Amount >0";
                                }
                                if (txt_item2.Enabled)
                                {
                                    sqlchkFirst += @"where Item_No_Slave=@item_ms and Position=@Position 
                                                                and Package = @Package and Actual_InDate = @actualDate 
                                                                and Mark = @mark and PCB_DC = @pcbdc and CMC_DC = @cmcdc and Amount >0";
                                }
                                //if (!string.IsNullOrEmpty(master))
                                //{
                                //    sqlchkFirst+= @"where Item_No_Master=@item_ms and Position=@Position 
                                //                                and Package = @Package and Actual_InDate=@actualDate";
                                //}
                                //else if (!string.IsNullOrEmpty(slave))
                                //{
                                //    sqlchkFirst += @"where Item_No_Slave=@item_ms and Position=@Position 
                                //                                and Package = @Package and Actual_InDate=@actualDate";
                                //}

                                SqlParameter[] parameters = new SqlParameter[]
                                {
                                    new SqlParameter("item_ms",item_MorS),
                                    new SqlParameter("Position",b[0].Text.Trim()),
                                    new SqlParameter("Package",package.Trim()),
                                    new SqlParameter("actualDate",actualDate),//目前為手動輸入,將來改自動帶入系統時間
                                    new SqlParameter("mark",mark),
                                    new SqlParameter("pcbdc",pcbdc),
                                    new SqlParameter("cmcdc",cmcdc)
                                };
                                DataSet dataSet = db.ExecuteDataSet(sqlchkFirst, CommandType.Text, parameters);
                                #endregion
                                if (dataSet.Tables[0].Rows.Count > 0)
                                {
                                    string detailSno = dataSet.Tables[0].Rows[0]["Sno"].ToString();
                                    //資料已存在update Amount+1 ,Amount_Unit+1
                                    sqldetail = @"update Automatic_Storage_Detail set Amount=Amount+@Amount, Mark=@Mark ,PCB_DC=@pcbdc , CMC_DC = @cmcdc ";
                                    sqldetail += @" where Sno = @Sno";

                                    SqlParameter[] parm2 = new SqlParameter[]
                                    {
                                        new SqlParameter("Amount",amount),
                                        //new SqlParameter("Amount_Unit",amount_U),
                                        new SqlParameter("Mark",mark),
                                        new SqlParameter("pcbdc",pcbdc),
                                        new SqlParameter("cmcdc",cmcdc),
                                        new SqlParameter("Sno",detailSno)
                                        //new SqlParameter("item_ms",item_MorS),
                                        //new SqlParameter("Position",b[0].Text.Trim()),
                                        //new SqlParameter("Package",package.Trim()),
                                        //new SqlParameter("actualDate",actualDate)//目前為手動輸入,將來改自動帶入系統時間
                                    };
                                    db.ExecueNonQuery(sqldetail, CommandType.Text, "btn_only_commit", parm2);
                                }
                                else
                                {
                                    //新資料
                                    //寫入detail table
                                    sqldetail = @"INSERT INTO Automatic_Storage_Detail
                                          (Sno,Item_No_Master,Item_No_Slave,Spec,Unit_No,Position,Up_InDate,Actual_InDate,Input_UserNo,Amount,Package,Mark,PCB_DC,CMC_DC)	 
	                                      SELECT top(1)Sno,Item_No_Master,Item_No_Slave,Spec,Unit_No,Position,GETDATE(),Actual_InDate,Input_UserNo,Amount,Package,Mark,PCB_DC,CMC_DC 
                                          FROM Automatic_Storage_Input
	                                      where (Item_No_Master= @Master or Item_No_Slave = @Slave) and Unit_No=@Unit_No and Position=@position and Package=@Package order by Input_Date desc";
                                    SqlParameter[] parm2 = new SqlParameter[]
                                    {
                                    //new SqlParameter("Reel_ID",c[i].Text.Trim()),
                                        new SqlParameter("Master",master),
                                        new SqlParameter("Slave",slave),
                                        new SqlParameter("Unit_No",Login.Unit_No),
                                        new SqlParameter("position",b[0].Text.Trim()),
                                        new SqlParameter("Package",package.Trim())
                                    };
                                    db.ExecueNonQuery(sqldetail, CommandType.Text, "btn_only_commit", parm2);
                                }
                            }
                        }
                        else
                        {
                            lbl_result.Text = "失敗, 請輸入儲位";
                        }
                    }
                    lbl_result.Text = "單筆入庫完成!!";
                }
                else
                {
                    if (dsP.Tables[0].Rows.Count == 0)
                    {
                        lbl_result.Text = "該儲位尚未設定";
                    }
                    if (string.IsNullOrEmpty(txt_Spec.Text))
                    {
                        lbl_result.Text = "該料號規格尚未設定";
                    }
                    return;
                }

                txt_clearn();
                labCMC.Visible = false;

            }
            catch (Exception ex)
            {
                lbl_result.Text = ex.Message.ToString();
            }
        }

        /// <summary>
        /// 取得移除尾端班別字串的結果。
        /// 若字串倒數第3~1碼為 "-01" ~ "-05" 以外的 "-xx" 格式，則移除尾端3碼，否則回傳原字串。
        /// </summary>
        /// <param name="txt">欲處理的原始字串。</param>
        /// <returns>移除班別後的字串。</returns>
        private string subShift(string txt)
        {
            string sub = StringSplit.StrLeft(StringSplit.StrRight(txt, 3), 1);
            string strShift = (sub == "-" && (StringSplit.StrRight(txt, 3).CompareTo("-01") != 0
                                            || StringSplit.StrRight(txt, 3).CompareTo("-02") != 0
                                            || StringSplit.StrRight(txt, 3).CompareTo("-03") != 0
                                            || StringSplit.StrRight(txt, 3).CompareTo("-04") != 0
                                            || StringSplit.StrRight(txt, 3).CompareTo("-05") != 0))
                       ? StringSplit.StrLeft(txt, txt.Length - 3) : txt;

            return strShift;
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
        /// 根據指定的縮放比例，遞迴調整所有控制項的寬度、高度、位置及字體大小。
        /// </summary>
        /// <param name="newx">寬度縮放比例。</param>
        /// <param name="newy">高度縮放比例。</param>
        /// <param name="cons">要調整的父控制項。</param>
        private void SetControls(float newx, float newy, Control cons)
        {
            if (isLoaded)
            {
                // 遍歷窗體中的所有控制項，根據縮放比例重新設定控制項的尺寸與位置
                foreach (Control con in cons.Controls)
                {
                    // 取得控制項的 Tag 屬性，並分割為寬度、高度、左邊距、頂邊距、字體大小
                    string[] mytag = con.Tag.ToString().Split(new char[] { ':' });
                    // 計算並設定寬度
                    float a = System.Convert.ToSingle(mytag[0]) * newx;
                    con.Width = (int)a;
                    // 計算並設定高度
                    a = System.Convert.ToSingle(mytag[1]) * newy;
                    con.Height = (int)(a);
                    // 計算並設定左邊距
                    a = System.Convert.ToSingle(mytag[2]) * newx;
                    con.Left = (int)(a);
                    // 計算並設定頂邊距
                    a = System.Convert.ToSingle(mytag[3]) * newy;
                    con.Top = (int)(a);
                    // 計算並設定字體大小
                    Single currentSize = System.Convert.ToSingle(mytag[4]) * newy;
                    con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
                    // 若控制項還有子控制項，則遞迴調整
                    if (con.Controls.Count > 0)
                    {
                        SetControls(newx, newy, con);
                    }
                }
            }
        }

        /// <summary>
        /// 將發生的例外寫入錯誤日誌，包含來源與檔案/行號（若可取得）。
        /// </summary>
        /// <param name="ex">例外物件。</param>
        /// <param name="context">額外的上下文字串，用來說明發生位置。</param>
        private void LogException(Exception ex, string context)
        {
            try
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context} - {ex.GetType()}: {ex.Message}");
                // 嘗試取得具檔案與行號的 StackTrace 資訊
                try
                {
                    var st = new StackTrace(ex, true);
                    for (int i = 0; i < st.FrameCount; i++)
                    {
                        var frame = st.GetFrame(i);
                        if (frame == null) continue;
                        var file = frame.GetFileName();
                        var line = frame.GetFileLineNumber();
                        sb.AppendLine($"   at {frame.GetMethod()} in {file}:{line}");
                    }
                }
                catch { }
                sb.AppendLine(ex.StackTrace ?? "");

                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "error.log");
                File.AppendAllText(logPath, sb.ToString());
            }
            catch { }
        }

        /// <summary>
        /// Input 表單顯示事件。
        /// 當 Input 表單顯示時，將視窗狀態設為最大化。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 Input 表單。</param>
        /// <param name="e">事件參數。</param>
        private void Input_Shown(object sender, EventArgs e)
        {
                // 延遲在下一個 message loop 週期設定最大化，以避免與其他 Show/WindowState 呼叫產生 race condition
                try
                {
                    this.BeginInvoke((Action)(() =>
                    {
                        try { this.WindowState = FormWindowState.Maximized; } catch { }
                    }));
                }
                catch { }
        }

        /// <summary>
        /// 視窗大小調整事件。
        /// 當使用者調整 Input 表單大小時，根據原始寬高比例縮放所有控制項。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 Input 表單。</param>
        /// <param name="e">事件參數。</param>
        private void Input_Resize(object sender, EventArgs e)
        {
            // 若尚未記錄原始寬高則不執行縮放
            if (X == 0 || Y == 0) return;
            // 計算寬度與高度縮放比例
            fgX = (float)this.Width / (float)X;
            fgY = (float)this.Height / (float)Y;

            // 依比例調整所有控制項
            SetControls(fgX, fgY, this);
        }

        /// <summary>
        /// 處理 txtbox1 的滑鼠點擊事件。
        /// 當使用者點擊 txtbox1 時，會將文字框內容清空，方便重新輸入。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txtbox1。</param>
        /// <param name="e">滑鼠點擊事件參數。</param>
        private void txtbox1_MouseClick(object sender, MouseEventArgs e)
        {
            TextBox text = (TextBox)sender;
            text.Text = "";
        }

        /// <summary>
        /// 處理料號輸入框離開事件，根據輸入料號查詢規格、儲位及訊息，並更新相關欄位與標籤。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Item1。</param>
        /// <param name="e">事件參數。</param>
        private void txt_Item1_Leave(object sender, EventArgs e)
        {
            // 宣告 SQL 查詢字串
            string sqlitem = "";
            string sqlpositionList = "";
            string ItemChoice = "";
            string msgcheck = "";
            // 建立儲位標籤陣列
            Label[] lbl = new Label[18] { lab_p1, lab_p2, lab_p3, lab_p4, lab_p5, lab_p6, lab_p7, lab_p8, lab_p9, lab_p10,
                                            lab_p11,lab_p12,lab_p13,lab_p14,lab_p15,lab_p16,lab_p17,lab_p18};

            // 檢查昶亨料號是否有輸入並防呆（避免字串長度不足導致 IndexOutOfRange）
            try
            {
                var text = txt_Item1.Text?.Trim();
                if (!string.IsNullOrEmpty(text) && (text?.Length ?? 0) >= 4)
                {
                    // 取得料號倒數第4碼判斷是否為 CMC
                    string auditCMC = text![(text!.Length - 4)].ToString().ToUpper();
                    // 若倒數第4碼為 A，顯示 CMC 標籤
                    labCMC.Visible = (auditCMC == "A");
                }
                else
                {
                    labCMC.Visible = false;
                }
            }
            catch (Exception ex)
            {
                LogException(ex, "txt_Item1_Leave - auditCMC");
                try { MessageBox.Show("料號檢查發生錯誤，請聯絡系統管理員。"); } catch { }
            }

            // 若昶亨料號與客戶料號皆未輸入，顯示提示訊息
            if (string.IsNullOrEmpty(txt_Item1.Text) && string.IsNullOrEmpty(txt_item2.Text))
            {
                lbl_result.Text = "料號尚未輸入!!";
            }
            else
            {
                // 若昶亨料號有輸入
                if (!string.IsNullOrEmpty(txt_Item1.Text))
                {
                    sqlitem = "select item_E,item_C,Spec from Automatic_Storage_Spec where item_E = @item";
                    sqlpositionList = "select Position from Automatic_Storage_Detail where Item_No_Master = @item and amount>0  group by Position";
                    ItemChoice = txt_Item1.Text ?? "";
                    msgcheck = "select * from Automatic_Storage_Msg where item_E =@item ";
                }
                // 若客戶料號有輸入
                if (!string.IsNullOrEmpty(txt_item2.Text))
                {
                    sqlitem = "select item_E,item_C,Spec from Automatic_Storage_Spec where item_C =@item";
                    sqlpositionList = "select Position from Automatic_Storage_Detail where Item_No_Slave = @item and amount>0  group by Position";
                    ItemChoice = txt_item2.Text ?? "";
                    msgcheck = "select * from Automatic_Storage_Msg where item_C =@item ";
                }
                SqlParameter[] parm = new SqlParameter[]
                {
                    new SqlParameter("item",ItemChoice)
                };
                SqlParameter[] parm2 = new SqlParameter[]
                {
                    new SqlParameter("item",ItemChoice)
                };
                SqlParameter[] parm3 = new SqlParameter[]
                {
                    new SqlParameter("item",ItemChoice)
                };
                // 查詢規格資料
                DataSet ds = db.ExecuteDataSet(sqlitem, CommandType.Text, parm);
                // 查詢儲位資料
                DataSet ds2 = db.ExecuteDataSet(sqlpositionList, CommandType.Text, parm2);
                // 查詢訊息資料
                DataSet ds3 = db.ExecuteDataSet(msgcheck, CommandType.Text, parm3);

                // 若有查到規格資料
                if (ds.Tables[0].Rows.Count > 0)
                {
                    // 顯示規格
                    txt_Spec.Text = ds.Tables[0].Rows[0]["Spec"].ToString();
                    // 若昶亨料號有輸入，帶出客戶料號
                    if (!string.IsNullOrEmpty(txt_Item1.Text))
                    {
                        txt_item2.Text = ds.Tables[0].Rows[0]["item_C"].ToString();
                    }
                    // 若客戶料號有輸入，帶出昶亨料號
                    if (!string.IsNullOrEmpty(txt_item2.Text))
                    {
                        txt_Item1.Text = ds.Tables[0].Rows[0]["item_E"].ToString();
                    }
                    //檢查msg資料表是否有訊息需要顯示
                    if (ds3.Tables[0].Rows.Count > 0)
                    {
                        MessageBox.Show(ds3.Tables[0].Rows[0]["msg"].ToString(), "注意!");
                    }
                }
                else
                {
                    lbl_result.Text = "找不到規格資料，請重新確認!!";
                    return;
                }

                // 若有查到儲位資料，顯示於標籤
                if (ds2.Tables[0].Rows.Count > 0)
                {
                    for (int lb = 0; lb < ds2.Tables[0].Rows.Count; lb++)
                    {
                        lbl[lb].Text = ds2.Tables[0].Rows[lb][0].ToString();
                    }

                }
            }
        }

        /// <summary>
        /// 處理數量輸入框的按鍵事件。
        /// 當使用者在 txt_Amount 輸入框按下 Enter 鍵時，會自動跳至下一個元件。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Amount。</param>
        /// <param name="e">按鍵事件參數。</param>
        private void txt_Amount_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否按下 Enter 鍵
            if (e.KeyChar == 13)
            {
                // 跳至下一個元件
                this.SelectNextControl(this.ActiveControl, true, true, true, false); //跳下一個元件
            }
        }

        /// <summary>
        /// 處理儲位輸入框的按鍵事件，當按下 Enter 鍵時跳至下一個元件。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Storage1。</param>
        /// <param name="e">按鍵事件參數。</param>
        private void txt_Storage1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否按下 Enter 鍵
            if (e.KeyChar == 13)
            {
                // 跳至下一個元件
                this.SelectNextControl(this.ActiveControl, true, true, true, false);
            }
        }

        /// <summary>
        /// 料號輸入框文字變更事件。
        /// 當 txt_Item1 或 txt_item2 內容變更時，根據焦點狀態調整欄位啟用狀態與結果標籤顯示。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Item1。</param>
        /// <param name="e">事件參數。</param>
        private void txt_Item1_TextChanged(object sender, EventArgs e)
        {
            // 若客戶料號欄位有值，則清空結果標籤
            if (!string.IsNullOrEmpty(txt_item2.Text))
            {
                lbl_result.Text = "";
            }
            // 若目前焦點在昶亨料號欄位，則停用客戶料號欄位
            if (txt_Item1.Focused)
            {
                txt_item2.Enabled = false;
            }
            // 若目前焦點在客戶料號欄位，則停用昶亨料號欄位
            else if (txt_item2.Focused)
            {
                txt_Item1.Enabled = false;
            }
        }

        /// <summary>
        /// 當使用者選擇不同的包裝種類時觸發此事件。
        /// 會將焦點移至儲位輸入框，方便使用者繼續輸入儲位資料。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為下拉選單。</param>
        /// <param name="e">事件參數。</param>
        private void cbx_package_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 將焦點移至儲位輸入框
            txt_Storage1.Focus();
        }

        /// <summary>
        /// 入庫日期選擇器值變更事件。
        /// 檢查選擇的日期是否大於今天，若是則顯示錯誤訊息並自動帶入今天日期。
        /// </summary>
        /// <param name="sender">事件來源物件。</param>
        /// <param name="e">事件參數。</param>
        private void dt_Picker_ValueChanged(object sender, EventArgs e)
        {
            // 計算選擇日期與今天的差距
            TimeSpan timeSpan = dt_Picker.Value.Date.Subtract(DateTime.Today);
            // 若選擇日期大於今天，顯示錯誤訊息
            if ((timeSpan.TotalDays >= 1))
            {
                MessageBox.Show("入庫日期錯誤!! 按下確認後系統將帶入今天日期。");
            }
            // 若選擇日期大於今天，則自動帶入今天日期，否則維持原值
            dt_Picker.Value = (timeSpan.TotalDays >= 1) ? DateTime.Now : dt_Picker.Value;
        }
    }
}
