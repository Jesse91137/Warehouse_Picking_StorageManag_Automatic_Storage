using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace Automatic_Storage
{
    /// <summary>
    /// 批次出庫明細視窗
    /// </summary>
    public partial class BatOutView : Form
    {
        /// <summary>
        /// 料號
        /// </summary>
        private string strItem;
        /// <summary>
        /// 機種
        /// </summary>
        private string strEngsr;

        /// <summary>
        /// 設定料號
        /// </summary>
        public string Item
        {
            set { strItem = value; } // 設定料號
        }
        //public string Engsr
        //{
        //    set { strEngsr = value; } // 設定機種
        //}

        /// <summary>
        /// 設定視窗值（目前僅料號）
        /// </summary>
        public void setValue()
        {
            string sItem = strItem; // 取得料號
                                    //string sEng = strEngsr; // 取得機種
        }

        /// <summary>
        /// 建構函式
        /// </summary>
        public BatOutView()
        {
            InitializeComponent(); // 初始化元件
            isLoaded = false; // 設定初始載入狀態
        }

        #region 視窗ReSize
        /// <summary>
        /// 視窗寬度（初始值，用於縮放計算）
        /// </summary>
        int X = new int();  //窗口寬度
        /// <summary>
        /// 視窗高度（初始值，用於縮放計算）
        /// </summary>
        int Y = new int(); //窗口高度
        /// <summary>
        /// 視窗寬度縮放比例
        /// </summary>
        float fgX = new float(); //寬度縮放比例
        /// <summary>
        /// 視窗高度縮放比例
        /// </summary>
        float fgY = new float(); //高度縮放比例
        /// <summary>
        /// 是否已設定各控制項的尺寸資料到 Tag 屬性
        /// </summary>
        bool isLoaded;  // 是否已設定各控制的尺寸資料到Tag屬性
        #endregion

        /// <summary>
        /// 搜尋文字（用於查詢功能）
        /// </summary>
        string txt_search = string.Empty; // 搜尋文字
        /// <summary>
        /// 班別（用於記錄班次資訊）
        /// </summary>
        string strShift = string.Empty; // 班別
        /// <summary>
        /// 出庫明細資料表
        /// </summary>
        DataTable dt = new DataTable(); // 出庫明細資料表
        /// <summary>
        /// 當前日期時間（格式：yyyy-MM-dd HH:mm:ss）
        /// </summary>
        string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // 當前日期時間

        /// <summary>
        /// 視窗載入事件
        /// </summary>
        private void OutPut_Load(object sender, EventArgs e)
        {
            #region 獲取窗體info

            X = this.Width;//獲取窗體的寬度
            Y = this.Height;//獲取窗體的高度
            isLoaded = true;// 已設定各控制項的尺寸到Tag屬性中
            SetTag(this);//調用方法                     +

            #endregion

            #region LoadBind

            dataGridView1.DataSource = dataBind(strItem); // 綁定資料到DataGridView
            #endregion

        }

        int cp = -1; // 當前索引

        /// <summary>
        /// 取得出庫明細資料
        /// </summary>
        /// <param name="item">料號</param>
        /// <returns>明細資料表</returns>
        public DataTable dataBind(string item)
        {
            // SQL查詢語句
            string sqlstr = "select Item_No_Master ,Item_No_Slave ,Spec " +
                                ",Amount_Unit ,Amount ,Position ,Package  " +
                                "from Automatic_Storage_Detail " +
                                "where Unit_No = @unitNo and Item_No_Master = @Item " +
                                "and Amount > 0  order by position asc";
            // SQL參數
            SqlParameter[] parm = new SqlParameter[]
            {
                    new SqlParameter("unitNo",Login.Unit_No), // 單位編號
                    new SqlParameter("Item",strItem) // 料號
            };
            return db.ExecuteDataTable(sqlstr, CommandType.Text, parm); // 執行查詢並回傳資料表
        }

        /// <summary>
        /// 視窗關閉事件
        /// </summary>
        private void OutPut_FormClosing(object sender, FormClosingEventArgs e)
        {

        }
        #region 視窗設定
        /// <summary>
        /// 將控制項的寬、高、左邊距、頂邊距和字體大小暫存到Tag屬性中
        /// </summary>
        /// <param name="cons">遞歸控制項中的控制項</param>
        private void SetTag(Control cons)
        {
            // 遍歷所有控制項
            foreach (Control con in cons.Controls)
            {
                // 將控制項的寬度、高度、左邊距、頂邊距和字體大小存到 Tag 屬性
                con.Tag = con.Width + ":" + con.Height + ":" + con.Left + ":" + con.Top + ":" + con.Font.Size;
                // 如果控制項還有子控制項則遞迴呼叫 SetTag
                if (con.Controls.Count > 0)
                    SetTag(con);
            }
        }

        /// <summary>
        /// 根據縮放比例調整控制項尺寸
        /// </summary>
        /// <param name="newx">寬度比例</param>
        /// <param name="newy">高度比例</param>
        /// <param name="cons">控制項</param>
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
        /// 視窗大小調整事件
        /// </summary>
        private void OutPut_Resize(object sender, EventArgs e)
        {
            if (X == 0 || Y == 0) return; // 若初始寬高為0則不處理
            fgX = (float)this.Width / (float)X; // 計算寬度縮放比例
            fgY = (float)this.Height / (float)Y; // 計算高度縮放比例

            SetControls(fgX, fgY, this); // 調整所有控制項尺寸
        }

        /// <summary>
        /// 視窗顯示事件，最大化視窗
        /// </summary>
        private void OutPut_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized; // 最大化視窗
        }
        #endregion

        /// <summary>
        /// DataGridView 滑鼠雙擊事件
        /// </summary>
        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //this.Close(); // 關閉視窗
        }

        /// <summary>
        /// DataGridView 雙擊事件
        /// </summary>
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
        }

        /// <summary>
        /// DataGridView 儲存格滑鼠雙擊事件
        /// </summary>
        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // 0→料號; 1→位置; 2→機種; 3→數量;

            BatOut father = (BatOut)this.Owner; // 取得父視窗

            int c_index = dataGridView1.CurrentCell.ColumnIndex; // 取得目前欄位索引

            switch (c_index)
            {
                case 1:
                    if (!string.IsNullOrEmpty(dataGridView1.CurrentRow.Cells[3].Value?.ToString() ?? string.Empty)) // 檢查數量欄位是否有值
                    {
                        if (Int32.Parse(dataGridView1.CurrentRow.Cells[3].Value?.ToString() ?? "0") > 0) // 檢查數量是否大於0
                        {
                            DialogResult res = MessageBox.Show("是否拿取( " + dataGridView1.CurrentRow.Cells[1].Value?.ToString() ?? string.Empty + " )位置１捲",
                                "請點兩下取出庫", MessageBoxButtons.YesNo, MessageBoxIcon.Question); // 顯示確認訊息
                            if (res == DialogResult.Yes) // 若選擇是
                            {
                                father.MsgFromChildPosition = dataGridView1.CurrentCell?.Value?.ToString() ?? string.Empty; // 設定父視窗位置
                                this.Close(); // 關閉視窗
                            }
                            else
                            {
                                return; // 不處理
                            }
                        }
                        else
                        {
                            MessageBox.Show("數量小於０不可取出庫!!"); // 顯示錯誤訊息
                            return;
                        }
                    }
                    else
                    {
                        return; // 不處理
                    }
                    break;
                case 2:
                    if (!string.IsNullOrEmpty(dataGridView1.CurrentRow?.Cells[2].Value?.ToString() ?? string.Empty)) // 檢查機種欄位是否有值
                    {
                        DialogResult res2 = MessageBox.Show("是否線上拿取( " + (dataGridView1.CurrentRow?.Cells[2].Value?.ToString() ?? string.Empty) + " )位置１捲",
                                "請點兩下取出庫", MessageBoxButtons.YesNo, MessageBoxIcon.Question); // 顯示確認訊息
                        if (res2 == DialogResult.Yes) // 若選擇是
                        {
                            father.MsgFromChildPosition = dataGridView1.CurrentRow?.Cells[1].Value?.ToString() ?? string.Empty; // 設定父視窗位置
                            father.MsgFromChildEngSr = dataGridView1.CurrentCell?.Value?.ToString() ?? string.Empty; // 設定父視窗機種
                            this.Close(); // 關閉視窗
                        }
                    }
                    else
                    {
                        MessageBox.Show("機種名稱空白，未在線上!!"); // 顯示錯誤訊息
                        return;
                    }
                    break;
                default:
                    father.MsgFromChildPosition = string.Empty; // 清空父視窗位置
                    father.MsgFromChildEngSr = string.Empty; // 清空父視窗機種
                    break;
            }
        }
    }
}
