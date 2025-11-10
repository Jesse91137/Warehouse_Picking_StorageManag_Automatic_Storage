using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Security;
using System.Windows.Forms;
using System.Windows.Threading;

namespace Automatic_Storage
{
    /// <summary>
    /// 管理儲位、規格、人員、訊息等功能的主視窗。
    /// </summary>
    public partial class Maintain : Form
    {
        /// <summary>
        /// 下拉選單用的項目類別，包含顯示文字與值。
        /// </summary>
        public class MyItem
        {
            /// <summary>
            /// 顯示文字
            /// </summary>
            public string text;
            /// <summary>
            /// 對應值
            /// </summary>
            public string value;

            /// <summary>
            /// 建構子，初始化顯示文字與值。
            /// </summary>
            /// <param name="text">顯示文字</param>
            /// <param name="value">對應值</param>
            public MyItem(string text, string value)
            {
                this.text = text;
                this.value = value;
            }
            /// <summary>
            /// 傳回顯示文字
            /// </summary>
            /// <returns>顯示文字</returns>
            public override string ToString()
            {
                return text;
            }
        }

        /// <summary>
        /// 建構子，初始化元件與資料表。
        /// </summary>
        public Maintain()
        {
            InitializeComponent(); // 初始化視窗元件
            isLoaded = false; // 標記尚未載入控制項尺寸
            dd = dd ?? columnsData(); // 初始化儲位資料表
        }

        /// <summary>
        /// Excel匯入筆數
        /// </summary>
        int inputexcelcount = 0, btn_fAll = 0, btn_itemP = 0;
        /// <summary>
        /// 初始化總數
        /// </summary>
        int initSum = 0;
        /// <summary>
        /// Excel資料集
        /// </summary>
        DataSet dsData = new DataSet();
        /// <summary>
        /// 計時器
        /// </summary>
        DispatcherTimer dispatcherTimer = new DispatcherTimer();

        #region 視窗ReSize
        /// <summary>
        /// 視窗寬度
        /// </summary>
        int X = new int();
        /// <summary>
        /// 視窗高度
        /// </summary>
        int Y = new int();
        /// <summary>
        /// 寬度縮放比例
        /// </summary>
        float fgX = new float();
        /// <summary>
        /// 高度縮放比例
        /// </summary>
        float fgY = new float();
        /// <summary>
        /// 是否已載入控制項尺寸
        /// </summary>
        bool isLoaded;
        #endregion

        /// <summary>
        /// 建立儲位資料表欄位
        /// </summary>
        /// <returns>儲位資料表</returns>
        private DataTable columnsData()
        {
            using (DataTable table = new DataTable())
            {
                table.Columns.Add("sno", typeof(string)); //序號
                table.Columns.Add("UNIT_NAME", typeof(string)); //單位名稱
                table.Columns.Add("User_name", typeof(string)); //使用者名稱
                table.Columns.Add("Position", typeof(string)); //儲位
                table.Columns.Add("Create_Date", typeof(string)); //建立日期
                return table;
            }
        }

        /// <summary>
        /// 建立明細資料表欄位
        /// </summary>
        /// <returns>明細資料表</returns>
        private DataTable columnsDataTable()
        {
            using (DataTable data = new DataTable())
            {
                data.Columns.Add("sno", typeof(string)); //序號
                data.Columns.Add("Actual_InDate", typeof(DateTime)); //實際入庫日
                data.Columns.Add("Item_No_Master", typeof(string)); //主料號
                data.Columns.Add("Spec", typeof(string)); //規格
                data.Columns.Add("Position", typeof(string)); //儲位
                data.Columns.Add("Package", typeof(string)); //包裝
                data.Columns.Add("Up_InDate", typeof(string)); //上架日期
                data.Columns.Add("User_name", typeof(string)); //使用者名稱
                data.Columns.Add("Mark", typeof(string)); //備註
                return data;
            }
        }

        /// <summary>
        /// 目前時間字串
        /// </summary>
        static string time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        /// <summary>
        /// 儲位資料表
        /// </summary>
        public DataTable dd;
        /// <summary>
        /// 明細資料表
        /// </summary>
        public DataTable dataT = new DataTable();
        /// <summary>
        /// 狀態旗標
        /// </summary>
        bool flag = false;
        /// <summary>
        /// 儲位序號
        /// </summary>
        static string position_sno;

        /// <summary>
        /// TabControl切換事件
        /// </summary>
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch ((sender as TabControl).SelectedIndex)
            {
                case 1:
                    try
                    {
                        checkedListBox1.Items.Clear();
                        dataBind(); //資料繫結
                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.Message);
                    }
                    break;
                case 2:
                    try
                    {
                        gv_Spec_Data(); //規格資料繫結
                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.Message);
                    }
                    break;
                case 3:
                    dataT = columnsDataTable(); //初始化明細資料表
                    break;
                case 4:
                    string sql = @"select UNIT_NAME as '單位名稱',ROLE_NAME as '權限',user_name as '使用者名稱' 
                                        from Automatic_Storage_UserRole a
                                        left join Automatic_Storage_Role b on a.ROLE_ID = b.ROLE_ID ,
                                        Automatic_Storage_User c, Automatic_Storage_UnitNo d 
                                        where c.User_No = a.USER_ID and c.Unit_No = d.UNIT_NO and c.User_No=@userNo";
                    SqlParameter[] parm = new SqlParameter[]
                    {
                            new SqlParameter("userNo",Login.User_No)
                    };
                    DataTable dt = db.ExecuteDataTable(sql, CommandType.Text, parm);

                    dataGridView1.DataSource = dt;
                    chk_dataBind(); //權限資料繫結
                    userDataBind(); //人員資料繫結
                    break;
                case 5:
                    DataMsg(); //訊息資料繫結
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 人員資料繫結
        /// </summary>
        private void userDataBind()
        {
            string sql_user = @"SELECT User_No as '工號',User_name as '姓名' from Automatic_Storage_User where User_No<>'02437' order by User_No";
            DataTable dt_user = db.ExecuteDataTable(sql_user, CommandType.Text, null);
            dataGridView3.DataSource = dt_user;
        }

        /// <summary>
        /// 儲位資料繫結
        /// </summary>
        private void dataBind()
        {
            dd.Clear(); // 清空儲位資料表
            string strsql = @"select a.Sno as sno,UNIT_NAME,User_name,Position,Create_Date 
                                    from Automatic_Storage_Position a
                                    left join Automatic_Storage_UnitNo b on a.Unit_No=b.UNIT_NO ,Automatic_Storage_User c 
                                    where a.Create_User=c.User_No and a.Unit_No=@unitno ";

            SqlParameter[] parameters = { new SqlParameter("unitno", Login.Unit_No) }; // 建立 SQL 參數，指定單位編號
            DataSet dt_all = db.ExecuteDataSet(strsql, CommandType.Text, parameters); // 執行 SQL 查詢，取得儲位資料集

            foreach (DataRow row in dt_all.Tables[0].Rows) // 逐筆處理查詢結果
            {
                DataRow dr = dd.NewRow(); // 新增一筆儲位資料
                dr["sno"] = row["sno"].ToString(); // 設定序號
                dr["UNIT_NAME"] = row["UNIT_NAME"].ToString(); // 設定單位名稱
                dr["User_name"] = row["User_name"].ToString(); // 設定使用者名稱
                dr["Position"] = row["Position"].ToString(); // 設定儲位
                dr["Create_Date"] = row["Create_Date"].ToString(); // 設定建立日期
                dd.Rows.Add(dr); // 加入儲位資料表
            }
            dataGridView2.DataSource = dd; // 將儲位資料表繫結到 DataGridView 顯示
        }
        /// <summary>
        /// 按下查詢人員權限按鈕時，根據工號或姓名查詢人員權限資料，並顯示於 dataGridView1。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void button2_Click(object sender, EventArgs e)
        {
            DataTable ddt = null;
            //dataGridView1.Rows.Clear();
            string strsql = @"select UNIT_NAME as '單位名稱',ROLE_NAME as '權限',user_name as '使用者名稱' 
                                from Automatic_Storage_UserRole a
                                left join Automatic_Storage_Role b on a.ROLE_ID = b.ROLE_ID ,
                                Automatic_Storage_User c, Automatic_Storage_UnitNo d 
                                where c.User_No = a.USER_ID and c.Unit_No = d.UNIT_NO and c.Unit_No=@unitno and (c.User_No=@userNo or c.User_name=@name)";
            SqlParameter[] parm = new SqlParameter[]
            {
                new SqlParameter("unitno",Login.Unit_No), // 單位編號參數
                new SqlParameter("userNo",txt_userid.Text.Trim()), // 工號參數
                new SqlParameter("name",txt_name.Text.Trim()) // 姓名參數
            };
            ddt = db.ExecuteDataTable(strsql, CommandType.Text, parm); // 執行 SQL 查詢，取得人員權限資料表
            dataGridView1.DataSource = ddt; // 將查詢結果繫結到 dataGridView1 顯示
        }

        /// <summary>
        /// 按下「儲位新增」按鈕時，依序檢查 12 個儲位輸入框，若有輸入且未重複則新增儲位，否則顯示重複警告。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void btn_submit_Click(object sender, EventArgs e)
        {
            try
            {
                TimerStart(); // 啟動計時器，2 秒後清空所有儲位輸入框
                // 建立 12 個儲位輸入框的陣列
                TextBox[] txt = new TextBox[12] { txt_Position1 , txt_Position2 , txt_Position3 , txt_Position4 , txt_Position5 , txt_Position6 ,
                            txt_Position7 , txt_Position8 ,txt_Position9,txt_Position10,txt_Position11,txt_Position12};
                for (int i = 0; i < txt.Length; i++)
                {
                    // 檢查儲位輸入框是否有輸入內容
                    if (txt[i].Text != "")
                    {
                        // 查詢該儲位是否已存在
                        string str_s = @"select * from Automatic_Storage_Position where Unit_No=@unitno and Position=@position";
                        SqlParameter[] parm_s = new SqlParameter[]
                        {
                                new SqlParameter("unitno",Login.Unit_No), // 單位編號參數
                                new SqlParameter("position",txt[i].Text.Trim()), // 儲位名稱參數
                        };
                        DataSet dataSet = db.ExecuteDataSet(str_s, CommandType.Text, parm_s); // 執行查詢
                        // 若查詢結果為 0 筆，表示儲位未重複
                        if (dataSet.Tables[0].Rows.Count == 0)
                        {
                            // 新增儲位資料
                            string strsql = @"insert into Automatic_Storage_Position (Unit_No,Position,Create_User,Create_Date) values
                                                        (@unitno, @position, @cr_user, @cr_date)";
                            SqlParameter[] parameters = new SqlParameter[]
                            {
                                    new SqlParameter("unitno",Login.Unit_No), // 單位編號
                                    new SqlParameter("position",txt[i].Text.Trim().ToUpper()), // 儲位名稱(轉大寫)
                                    new SqlParameter("cr_user",Login.User_No), // 建立人員編號
                                    new SqlParameter("cr_date",time) // 建立日期
                            };
                            db.ExecueNonQuery(strsql, CommandType.Text, parameters); // 執行新增
                            lab_finish.Text = "儲位新增完成"; // 顯示新增完成訊息
                        }
                        else
                        {
                            // 若儲位重複，顯示警告訊息
                            MessageBox.Show("第 " + (i + 1) + " 儲位重複,請再次確認");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 發生例外時顯示錯誤訊息
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// DataGridView2 儲位資料表格的儲存格點擊事件。
        /// 根據點擊的欄位執行不同動作：
        /// - 欄位0：進入儲位編輯模式。
        /// - 欄位1：執行儲位刪除。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 DataGridView2。</param>
        /// <param name="e">事件參數，包含點擊的儲存格位置。</param>
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    // 取得目前選取的列索引
                    int s = dataGridView2.CurrentRow.Index;
                    // 設定焦點到 DataGridView2
                    dataGridView2.Focus();
                    // 將目前儲存格設為第 5 欄（Position 欄）
                    dataGridView2.CurrentCell = dataGridView2[5, s];
                    // 進入編輯模式
                    dataGridView2.BeginEdit(true);
                    break;
                case 1:
                    // 執行儲位刪除
                    delete_Click();
                    break;
                default:
                    // 其他欄位不執行任何動作
                    break;
            }
        }

        /// <summary>
        /// DataGridView2 編輯控制項顯示事件。
        /// 判斷是否為 Position 欄位，若是則設定編輯狀態並將目前儲位資訊帶入編輯框。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 DataGridView2。</param>
        /// <param name="e">事件參數，包含編輯控制項。</param>
        private void dataGridView2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            // 判斷觸發事件的欄位是 Position 才進行以下動作
            if (((DataGridView)sender).Columns[((DataGridView)sender).CurrentCell.ColumnIndex].Name == "Position")
            {
                flag = true; // 設定狀態旗標為編輯中
                txt_new_p.Focus(); // 將焦點移至新儲位輸入框
                txt_new_p.Text = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[5].Value?.ToString() ?? string.Empty; // 帶入目前儲位名稱
                position_sno = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value?.ToString() ?? string.Empty; // 帶入目前儲位序號
                //TextBox txt = e.Control as TextBox;
                //if (txt != null)
                //{
                //    // 增加事件
                //    //txt.KeyPress += new KeyPressEventHandler(txt_KeyPress);
                //}
            }
        }

        /// <summary>
        /// txt_position_s_KeyPress 事件處理函式，當使用者在 txt_position_s 輸入框按下 Enter 鍵時觸發，
        /// 依據輸入的儲位名稱查詢資料庫，並將結果顯示於 dataGridView2。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_position_s 輸入框。</param>
        /// <param name="e">事件參數，包含按下的鍵。</param>
        private void txt_position_s_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 檢查是否按下 Enter 鍵
            if (e.KeyChar == 13)
            {
                try
                {
                    dd.Clear(); // 清空儲位資料表
                    // 建立查詢儲位的 SQL 語句
                    string strsql = @"select a.Sno as sno,UNIT_NAME,User_name,Position,Create_Date from Automatic_Storage_Position a
                                                 left join Automatic_Storage_UnitNo b on a.Unit_No=b.UNIT_NO ,Automatic_Storage_User c 
                                                 where a.Unit_No='M' and a.Create_User=c.User_No and a.Unit_No=@unitno and Position=@position ";
                    // 建立 SQL 參數，指定單位編號與儲位名稱
                    SqlParameter[] parameters = new SqlParameter[]
                    {
                        new SqlParameter("unitno",Login.Unit_No),
                        new SqlParameter("position",txt_position_s.Text.Trim())
                    };
                    // 執行 SQL 查詢，取得儲位資料集
                    DataSet dt = db.ExecuteDataSet(strsql, CommandType.Text, parameters);
                    // 逐筆處理查詢結果
                    foreach (DataRow row in dt.Tables[0].Rows)
                    {
                        DataRow dr = dd.NewRow(); // 新增一筆儲位資料
                        dr["sno"] = row["sno"].ToString(); // 設定序號
                        dr["UNIT_NAME"] = row["UNIT_NAME"].ToString(); // 設定單位名稱
                        dr["User_name"] = row["User_name"].ToString(); // 設定使用者名稱
                        dr["Position"] = row["Position"].ToString(); // 設定儲位
                        dr["Create_Date"] = row["Create_Date"].ToString(); // 設定建立日期
                        dd.Rows.Add(dr); // 加入儲位資料表
                    }
                    dataGridView2.DataSource = dd; // 將儲位資料表繫結到 DataGridView 顯示
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message); // 發生例外時顯示錯誤訊息
                }
            }
        }

        /// <summary>
        /// 按下「儲位更新」按鈕時，執行儲位名稱更新或刪除動作。
        /// 若新儲位名稱未重複則更新，否則刪除原儲位。
        /// 並同步更新明細表的儲位名稱。
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void button3_Click(object sender, EventArgs e)
        {
            // 檢查是否已選擇可編輯欄位
            if (flag)
            {
                // 查詢新儲位名稱是否已存在
                string strchk = @"select * from Automatic_Storage_Position where Position =@position";
                SqlParameter[] param = new SqlParameter[]
                {
                    new SqlParameter("position",txt_new_p.Text.Trim())

                };
                DataSet data = db.ExecuteDataSet(strchk, CommandType.Text, param);

                // 若新儲位名稱未重複，則執行更新
                if (data.Tables[0].Rows.Count == 0)
                {
                    string strsql = @"update Automatic_Storage_Position set Position=@position where sno=@sno";
                    SqlParameter[] sqlParameters = new SqlParameter[]
                    {
                    new SqlParameter("position",txt_new_p.Text.Trim()),
                    new SqlParameter("sno",position_sno)
                    };
                    db.ExecueNonQuery(strsql, CommandType.Text, sqlParameters);
                }
                // 若新儲位名稱重複，則刪除原儲位
                else
                {
                    string delsql = @"delete from Automatic_Storage_Position where sno=@sno";
                    SqlParameter[] delparam = new SqlParameter[]
                    {
                        new SqlParameter("sno",position_sno)
                    };
                    db.ExecueNonQuery(delsql, CommandType.Text, delparam);
                }

                // 重新繫結儲位資料表
                dataBind();

                #region Detail表同步更新
                // 同步更新明細表的儲位名稱
                string strsql2 = @"update Automatic_Storage_Detail set Position=@position1 where Position=@position2";
                SqlParameter[] sqlParameters2 = new SqlParameter[]
                {
                    new SqlParameter("position1",txt_new_p.Text.Trim()),
                    new SqlParameter("position2",txt_position_s.Text.Trim())
                };
                db.ExecueNonQuery(strsql2, CommandType.Text, sqlParameters2);
                #endregion
            }
            else
            {
                // 未選擇可編輯欄位時顯示提示訊息
                MessageBox.Show("未選擇更新欄位");
            }
        }
        /// <summary>
        /// 刪除儲位資料。
        /// - 目前為「不驗證」版本，直接依據選取列的序號刪除儲位。
        /// - 保留「有驗證」版本註解，未啟用：可檢查明細表是否尚有資料，若有則不允許刪除。
        /// </summary>
        private void delete_Click()
        {
            #region 刪除不驗證 for 小鄭 20210522
            // 取得目前選取列的儲位序號
            string id = dataGridView2.CurrentRow?.Cells[2].Value?.ToString() ?? string.Empty;
            // 建立刪除儲位的 SQL 語句
            string strsql = @"delete Automatic_Storage_Position where sno=@sno";
            SqlParameter[] sqlParameters = new SqlParameter[]
            {
                new SqlParameter("sno",id)
            };
            // 執行刪除儲位
            db.ExecueNonQuery(strsql, CommandType.Text, sqlParameters);
            // 重新繫結儲位資料表
            dataBind();
            #endregion

            #region 刪除有驗證 保留版
            // 確認明細表中該儲位是否尚有資料，若有則不允許刪除
            //string position = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[5].Value.ToString();
            //string sql_detail = @"select * from Automatic_Storage_Detail where Position=@position";
            //SqlParameter[] parm_detail = new SqlParameter[]
            //{
            //        new SqlParameter("position",position)
            //};
            //DataSet dataSet = db.ExecuteDataSet(sql_detail, CommandType.Text, parm_detail);
            //if (dataSet.Tables[0].Rows.Count == 0)
            //{
            //    string id = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value.ToString();
            //    string strsql = @"delete Automatic_Storage_Position where sno=@sno";
            //    SqlParameter[] sqlParameters = new SqlParameter[]
            //    {
            //        new SqlParameter("sno",id)
            //    };
            //    db.ExecueNonQuery(strsql, CommandType.Text, sqlParameters);
            //    dataBind();
            //}
            //else
            //{
            //    MessageBox.Show("儲位尚有資料,請確認後再刪除");
            //}
            #endregion
        }

        /// <summary>
        /// 處理 txt_Position1 的 KeyPress 事件。
        /// 當使用者在 txt_Position1 輸入框按下 Enter 鍵時，
        /// 自動跳至下一個控制項。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Position1。</param>
        /// <param name="e">事件參數，包含按下的鍵。</param>
        private void txt_Position1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 檢查是否按下 Enter 鍵
            if (e.KeyChar == 13)
            {
                this.SelectNextControl(this.ActiveControl, true, true, true, false); //跳下一個元件
            }
        }

        /// <summary>
        /// 權限資料繫結。
        /// 讀取 Automatic_Storage_Role 資料表，將所有角色名稱與角色代碼加入 checkedListBox1 控制項。
        /// </summary>
        private void chk_dataBind()
        {
            // 清空 checkedListBox1 的項目
            checkedListBox1.Items.Clear();
            string sqlcb = @"select ROLE_ID,ROLE_NAME from Automatic_Storage_Role ";
            DataSet ds = db.ExecuteDataSet(sqlcb, CommandType.Text, null);

            // 將查詢結果逐筆加入 checkedListBox1
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                checkedListBox1.Items.Add(new MyItem(row["ROLE_NAME"].ToString(), row["ROLE_ID"].ToString()));
            }
        }

        /// <summary>
        /// 按下「人員權限新增」按鈕時，將選取的權限賦予指定人員。
        /// 若該人員已存在權限，則先刪除原有權限再重新賦予。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void btn_role_add_Click(object sender, EventArgs e)
        {
            // 檢查是否有選取權限
            if (checkedListBox1.CheckedItems.Count > 0)
            {
                // 檢查人員規則是否已存在
                string sql_search = @"select * from Automatic_Storage_UserRole where USER_ID=@userid ";
                SqlParameter[] parm_se = new SqlParameter[]
                {
                    new SqlParameter("userid",txt_userid.Text.Trim())
                };
                DataSet data = db.ExecuteDataSet(sql_search, CommandType.Text, parm_se);

                // 若人員尚未有權限，則逐一新增選取的權限
                if (data.Tables[0].Rows.Count == 0)
                {
                    // 若人員尚未有權限，則逐一新增選取的權限
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        MyItem item = (MyItem)this.checkedListBox1.Items[i];

                        // 檢查該項目是否被選取
                        if (checkedListBox1.GetItemChecked(i))
                        {
                            string sql_ins = @"insert into Automatic_Storage_UserRole values (@userid,@roleid,@date)";
                            SqlParameter[] parm_role = new SqlParameter[]
                            {
                            new SqlParameter("userid",txt_userid_role.Text.Trim()),
                            new SqlParameter("roleid",item.value),
                            new SqlParameter("date",time)
                            };
                            db.ExecueNonQuery(sql_ins, CommandType.Text, parm_role);
                        }
                    }
                    MessageBox.Show("人員權限賦予完成");
                }
                else
                {
                    // 若人員已存在權限，則先刪除原有權限再重新賦予
                    string sql_del = @"delete from Automatic_Storage_UserRole where USER_ID=@userid";
                    SqlParameter[] parm_role2 = new SqlParameter[]
                    {
                        new SqlParameter("userid",txt_userid_role.Text.Trim())
                     };
                    db.ExecueNonQuery(sql_del, CommandType.Text, parm_role2);

                    // 重新賦予選取的權限
                    btn_role_add_Click(sender, e);
                    MessageBox.Show("人員權限更新完成");
                }
            }
            else
            {
                MessageBox.Show("至少選擇一項權限");
            }
        }

        /// <summary>
        /// 按下「新增人員」按鈕時，根據輸入的工號檢查人員是否已存在。
        /// 若不存在則新增人員資料，並將工號設為預設密碼。
        /// 新增成功後顯示人員姓名與工號，若重複則顯示警告訊息。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void btn_user_add_Click(object sender, EventArgs e)
        {
            try
            {
                // 查詢工號是否已存在
                string sql_search = @"select * from Automatic_Storage_User where User_No=@userid ";
                SqlParameter[] parm_se = new SqlParameter[]
                {
                    new SqlParameter("userid",txt_userid.Text.Trim())
                };
                DataSet data = db.ExecuteDataSet(sql_search, CommandType.Text, parm_se);
                if (data.Tables[0].Rows.Count == 0)
                {
                    // 若工號不存在，則新增人員資料
                    string sqladd = @"insert into Automatic_Storage_User values (@userid,@name,@unitno,@psw)";
                    SqlParameter[] parm_add = new SqlParameter[]
                    {
                        new SqlParameter("userid",txt_userid.Text.Trim()),
                        new SqlParameter("name",txt_name.Text.Trim()),
                        new SqlParameter("unitno",Login.Unit_No),
                        new SqlParameter("psw",txt_userid.Text.Trim())
                    };
                    db.ExecueNonQuery(sqladd, CommandType.Text, parm_add);
                    txt_userid_role.Text = txt_userid.Text;
                    MessageBox.Show("新增人員  :  " + txt_name.Text.Trim() + Environment.NewLine + "員工編號  :  " + txt_userid.Text.Trim());
                }
                else
                {
                    // 若工號已存在，顯示警告訊息
                    MessageBox.Show("不可重複新增人員");
                }
            }
            catch (Exception ee)
            {
                // 發生例外時顯示錯誤訊息
                MessageBox.Show(ee.Message);
            }
        }

        /// <summary>
        /// DataGridView3 的儲存格點擊事件。
        /// 當使用者點擊第 0 欄時，會刪除選取的人員資料及其權限資料。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 dataGridView3。</param>
        /// <param name="e">事件參數，包含點擊的儲存格位置。</param>
        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    // 取得目前選取列的人員編號
                    string id = dataGridView3.CurrentRow?.Cells[1].Value?.ToString() ?? string.Empty;
                    // 刪除人員資料
                    string strsql_u = @"delete Automatic_Storage_User where User_No=@userno";
                    SqlParameter[] parm_u = new SqlParameter[]
                    {
                            new SqlParameter("userno",id)
                    };
                    db.ExecueNonQuery(strsql_u, CommandType.Text, parm_u);
                    // 刪除人員權限資料
                    string strsql_r = @"delete from Automatic_Storage_UserRole where USER_ID=@userno";
                    SqlParameter[] parm_r = new SqlParameter[]
                    {
                            new SqlParameter("userno",id)
                    };
                    db.ExecueNonQuery(strsql_r, CommandType.Text, parm_r);

                    // 重新繫結人員資料
                    userDataBind();
                    break;
                default:
                    break;
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
        /// 依據指定的縮放比例，遞迴調整所有控制項的寬度、高度、位置與字體大小。
        /// </summary>
        /// <param name="newx">寬度縮放比例</param>
        /// <param name="newy">高度縮放比例</param>
        /// <param name="cons">要調整的父控制項</param>
        private void SetControls(float newx, float newy, Control cons)
        {
            if (isLoaded)
            {
                //遍歷窗體中的控制項，重新設置控制項的值
                foreach (Control con in cons.Controls)
                {
                    // 取得控制項的 Tag 屬性，格式為 "寬:高:左:上:字體大小"
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
                    // 計算並設定上邊距
                    a = System.Convert.ToSingle(mytag[3]) * newy;
                    con.Top = (int)(a);
                    // 計算並設定字體大小
                    Single currentSize = System.Convert.ToSingle(mytag[4]) * newy;
                    con.Font = new Font(con.Font.Name, currentSize, con.Font.Style, con.Font.Unit);
                    // 若有子控制項則遞迴調整
                    if (con.Controls.Count > 0)
                    {
                        SetControls(newx, newy, con);
                    }
                }
            }
        }

        /// <summary>
        /// 當主視窗顯示時，將視窗狀態設為最大化。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為主視窗。</param>
        /// <param name="e">事件參數。</param>
        private void Maintain_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        /// <summary>
        /// 視窗大小調整事件處理函式。<br/>
        /// 依據目前視窗的寬度與高度，計算縮放比例，並呼叫 <see cref="SetControls(float, float, Control)"/> 方法調整所有控制項的尺寸與位置。<br/>
        /// 若尚未取得初始寬高 (X, Y 為 0) 則不執行任何動作。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為主視窗。</param>
        /// <param name="e">事件參數。</param>
        private void Maintain_Resize(object sender, EventArgs e)
        {
            if (X == 0 || Y == 0) return;
            fgX = (float)this.Width / (float)X;
            fgY = (float)this.Height / (float)Y;

            SetControls(fgX, fgY, this);
        }

        /// <summary>
        /// 主視窗載入事件。
        /// 初始化窗體尺寸資訊，並設定所有控制項的 Tag 屬性以便後續縮放。
        /// 同時初始化檔案選擇對話框，僅允許選擇 .xls 格式的檔案。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為主視窗。</param>
        /// <param name="e">事件參數。</param>
        private void Maintain_Load(object sender, EventArgs e)
        {
            #region 獲取窗體info
            X = this.Width;//獲取窗體的寬度
            Y = this.Height;//獲取窗體的高度
            isLoaded = true;// 已設定各控制項的尺寸到Tag屬性中
            SetTag(this);//調用方法
            #endregion

            fileDialog1 = new OpenFileDialog()
            {
                FileName = "Select a text file",
                Filter = "Text files (*.xls)|*.xls",
                Title = "Open text file"
            };
        }

        /// <summary>
        /// 讀取 Automatic_Storage_Msg 資料表，將所有訊息資料繫結至 dataGridView5 顯示。
        /// </summary>
        private void DataMsg()
        {
            string sqlmsgs = @" select item_E as '昶亨料號',item_C as '客戶料號',msg as '訊息內容' 
                                from Automatic_Storage_Msg ";

            DataTable dtmg = db.ExecuteDataTable(sqlmsgs, CommandType.Text, null);
            dataGridView5.DataSource = dtmg;
        }

        /// <summary>
        /// 從 Automatic_Storage_Spec 資料表讀取指定單位的所有規格資料，
        /// 並將結果繫結至 gv_Spec DataGridView 顯示。
        /// </summary>
        private void gv_Spec_Data()
        {
            string sqlSpec = @"select Item_E as '昶亨料號' ,Item_C as '客戶料號' ,Spec as '規格' 
                               from Automatic_Storage_Spec where Unit_No = @unitno";
            SqlParameter[] parmSpec = new SqlParameter[]
            {
                new SqlParameter("unitno",Login.Unit_No)
            };
            DataTable dt_spec = db.ExecuteDataTable(sqlSpec, CommandType.Text, parmSpec);

            gv_Spec.DataSource = dt_spec;
        }

        /// <summary>
        /// btn_Spec_f_Click 事件處理函式。<br/>
        /// 依據使用者輸入的料號、客戶料號、規格等條件，查詢 Automatic_Storage_Spec 資料表，
        /// 並將查詢結果繫結至 gv_Spec DataGridView 顯示。<br/>
        /// - 若 txt_f_item_E 有值，則以「昶亨料號」模糊查詢。<br/>
        /// - 若 txt_f_item_C 有值，則以「客戶料號」模糊查詢。<br/>
        /// - 若 txt_f_item_Spec 有值，則以「規格」模糊查詢。<br/>
        /// </summary>
        /// <param name="sender">事件來源物件，通常為查詢按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void btn_Spec_f_Click(object sender, EventArgs e)
        {
            List<SqlParameter> parmF = new List<SqlParameter>();
            string sqlF = "select Item_E as '昶亨料號' ,Item_C as '客戶料號' ,Spec as '規格' from Automatic_Storage_Spec " +
                          "where Unit_no=@unitno ";

            // 基本條件：單位編號
            parmF.Add(new SqlParameter("unitno", Login.Unit_No));

            if (!string.IsNullOrEmpty(txt_f_item_E.Text))
            {
                sqlF += " and Item_E like @itemE +'%' ";
                parmF.Add(new SqlParameter("itemE", txt_f_item_E.Text.Trim()));
            }

            // 客戶料號
            if (!string.IsNullOrEmpty(txt_f_item_C.Text))
            {
                sqlF += " and Item_C like @itemC +'%' ";
                parmF.Add(new SqlParameter("itemC", txt_f_item_C.Text.Trim()));
            }

            // 規格
            if (!string.IsNullOrEmpty(txt_f_item_Spec.Text))
            {
                sqlF += " and Spec like @spec +'%' ";
                parmF.Add(new SqlParameter("spec", txt_f_item_Spec.Text.Trim()));
            }
            DataSet ds = db.ExecuteDataSetPmsList(sqlF, CommandType.Text, parmF);
            DataTable dt = ds.Tables[0];  //每次能讀取一張表
            gv_Spec.DataSource = dt;
        }

        /// <summary>
        /// 檔案選擇對話框，僅允許選擇 .xls 格式的檔案。
        /// </summary>
        private OpenFileDialog fileDialog1;

        /// <summary>
        /// 檔案選擇按鈕事件處理函式。<br/>
        /// 開啟檔案選擇對話框，選擇 Excel 檔案後，將檔案複製到指定目錄，並更新路徑顯示於 txt_path。<br/>
        /// 同時重新繫結儲位資料表。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為選擇檔案按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void selectButton_Click(object sender, EventArgs e)
        {
            // 顯示檔案選擇對話框
            if (fileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    // 取得選取的檔案名稱
                    var fileName = fileDialog1.FileName;

                    // 取得選取的檔案目錄
                    FileInfo _info = new FileInfo(fileName);

                    // 複製檔案指定目錄
                    string _new = Application.StartupPath + "\\Upload\\" + _info.Name;

                    // 若檔案已存在則複製檔案
                    if (File.Exists(fileName))
                    {
                        // 複製檔案到指定目錄
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
        /// commitButton 的 Click 事件處理函式。
        /// 1. 禁用 commitButton，避免重複執行。
        /// 2. 取得 txt_path 的檔案路徑，並載入 Excel 資料。
        /// 3. 清空 list_result 顯示區。
        /// 4. 啟用 inputExcelBW 的中斷功能，並以背景執行匯入作業。
        /// 5. 重新繫結規格資料至 gv_Spec。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 commitButton。</param>
        /// <param name="e">事件參數。</param>
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
                gv_Spec_Data();
            }
            catch (SecurityException ex)
            {
                MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// 讀取指定 Excel 檔案，將 Sheet1 的所有資料載入至 dsData。
        /// 支援 Office 97-2003 格式，連線字串使用 Microsoft.Jet.OLEDB.4.0。
        /// 若未選擇檔案則顯示提示訊息。
        /// </summary>
        /// <param name="filename">Excel 檔案完整路徑</param>
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
                cmd.CommandText = "SELECT * FROM [Sheet1$]";
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
        /// 將 Excel 資料寫入 Automatic_Storage_Spec 資料表。
        /// 1. 逐筆讀取 dsData 的每一列，取得料號、客戶料號、規格。
        /// 2. 若三者皆為空則記錄失敗訊息。
        /// 3. 若資料存在，先查詢資料表是否已有該料號。
        ///    - 若已存在則更新規格與客戶料號。
        ///    - 若不存在則新增一筆資料。
        /// 4. 每處理一筆資料即更新進度條。
        /// 5. 若發生例外則記錄錯誤訊息於 list_result。
        /// </summary>
        public void WriteExcelData()
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            try
            {
                string ItemE = "", ItemC = "", Spec = "";
                int[] count;
                inputexcelcount = dsData.Tables[0].Rows.Count;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = dsData.Tables[0].Rows.Count;
                progressBar1.Step = 1;
                for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
                {
                    count = new int[dsData.Tables[0].Rows.Count];
                    ItemE = dsData.Tables[0].Rows[i][0].ToString().Trim().ToUpper();
                    ItemC = dsData.Tables[0].Rows[i][1].ToString().Trim().ToUpper();
                    Spec = dsData.Tables[0].Rows[i][2].ToString().Trim().ToUpper();
                    // 檢查三個欄位是否皆為空
                    if ((string.IsNullOrEmpty(ItemE) && string.IsNullOrEmpty(ItemC) && string.IsNullOrEmpty(Spec)))
                    {
                        _ = (count[i] == 0) ? list_result.Items.Add("第" + (i + 1) + "筆上傳失敗").ToString() : "";
                    }
                    else
                    {
                        // 查詢資料表是否已有該料號
                        string sql_s = "select * from Automatic_Storage_Spec " +
                                                "where Unit_no=@unitno and Item_E=@itemE ";
                        SqlParameter[] parmS = new SqlParameter[]
                        {
                            new SqlParameter("unitno",Login.Unit_No),
                            new SqlParameter("itemE",ItemE.Trim())
                        };
                        DataSet dss = db.ExecuteDataSet(sql_s, CommandType.Text, parmS);
                        if (dss.Tables[0].Rows.Count > 0)
                        {
                            // 若已存在則更新規格與客戶料號
                            string sql_u = "update Automatic_Storage_Spec set Spec = @spec ,Item_C = @itemC where Unit_no=@unitno and Item_E=@itemE ";
                            SqlParameter[] parmU = new SqlParameter[]
                            {
                                new SqlParameter("spec",Spec),
                                new SqlParameter("itemC",ItemC),
                                new SqlParameter("unitno",Login.Unit_No),
                                new SqlParameter("itemE",ItemE.Trim())
                            };
                            count[i] = db.ExecueNonQuery(sql_u, CommandType.Text, parmU);
                        }
                        else
                        {
                            // 若不存在則新增一筆資料
                            string sql_i = "INSERT INTO Automatic_Storage_Spec ([Item_E],[Item_C],[Spec],[Unit_No],[Cr_User],[Cr_Date]) " +
                                            "VALUES(@itemE,@itemC,@spec,@unitno,@cruser,@crdate)";
                            SqlParameter[] parmI = new SqlParameter[]
                            {
                                new SqlParameter("itemE",ItemE.Trim()),
                                new SqlParameter("itemC",ItemC.Trim()),
                                new SqlParameter("spec",Spec.Trim()),
                                new SqlParameter("unitno",Login.Unit_No),
                                new SqlParameter("cruser",Login.User_No),
                                new SqlParameter("crdate",DateTime.Now)
                            };
                            count[i] = db.ExecueNonQuery(sql_i, CommandType.Text, parmI);
                        }

                    }
                    // 更新進度條
                    progressBar1.PerformStep();
                }
            }
            catch (Exception ex)
            {
                // 發生例外時記錄錯誤訊息
                list_result.Items.Add(ex.Message);
            }

            ////gv_Spec_Data();
        }

        /// <summary>
        /// inputExcelBW 的 DoWork 事件處理函式。
        /// 執行 Excel 資料匯入作業，並支援中斷功能。
        /// 1. 初始化進度計數器。
        /// 2. 檢查是否收到中斷請求，若是則取消作業。
        /// 3. 呼叫 WriteExcelData 方法進行資料匯入。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 inputExcelBW。</param>
        /// <param name="e">事件參數，包含作業狀態。</param>
        private void inputExcelBW_DoWork(object sender, DoWorkEventArgs e)
        {
            initSum = 0;
            if (inputExcelBW.CancellationPending) //如果被中斷...
                e.Cancel = true;
            try
            {
                WriteExcelData();
            }
            catch (Exception ex)
            {
                // 例外處理：此處未記錄錯誤訊息
            }
        }

        /// <summary>
        /// gv_Spec 的儲存格點擊事件。
        /// 當使用者點擊第 0 欄時，會詢問是否刪除該筆規格資料，若確認則刪除對應的 Automatic_Storage_Spec 資料表資料，
        /// 並重新繫結規格資料至 gv_Spec 顯示。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 gv_Spec DataGridView。</param>
        /// <param name="e">事件參數，包含點擊的儲存格位置。</param>
        private void gv_Spec_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    DialogResult res = MessageBox.Show("是否刪除( " + (gv_Spec.CurrentRow?.Cells[1].Value?.ToString() ?? string.Empty) + " )",
                                "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (res == DialogResult.Yes)
                    {
                        string itemE = gv_Spec.CurrentRow?.Cells[1].Value?.ToString() ?? string.Empty;
                        string itemC = gv_Spec.CurrentRow?.Cells[2].Value?.ToString() ?? string.Empty;
                        string spec = gv_Spec.CurrentRow?.Cells[3].Value?.ToString() ?? string.Empty;
                        string sql_d = "delete Automatic_Storage_Spec " +
                            "where Unit_no = @unitno and Item_E = @itemE and Item_C=@itemC and Spec=@spec ";
                        SqlParameter[] parmD = new SqlParameter[]
                        {
                                new SqlParameter("unitno",Login.Unit_No),
                                new SqlParameter("itemE",itemE),
                                new SqlParameter("itemC",itemC),
                                new SqlParameter("spec",spec),
                        };
                        db.ExecueNonQuery(sql_d, CommandType.Text, parmD);
                        gv_Spec_Data();
                    }
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// txt_Position1 的 TextChanged 事件處理函式。
        /// 當使用者在 txt_Position1 輸入框內容變更時，清空 lab_finish 的文字顯示。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Position1。</param>
        /// <param name="e">事件參數。</param>
        private void txt_Position1_TextChanged(object sender, EventArgs e)
        {
            lab_finish.Text = "";
        }

        /// <summary>
        /// 啟動 DispatcherTimer，於指定間隔後觸發 Tick 事件，執行清空儲位輸入框的動作。
        /// </summary>
        public void TimerStart()
        {
            // 建立 DispatcherTimer 實例
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            // 設定時間間隔為 2 秒
            dispatcherTimer.Interval = new TimeSpan(0, 0, 2);
            // 啟動計時器
            dispatcherTimer.Start();
        }

        /// <summary>
        /// DispatcherTimer 的 Tick 事件處理函式。
        /// 每次觸發時會清空 12 個儲位輸入框的內容。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 DispatcherTimer。</param>
        /// <param name="e">事件參數。</param>
        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            TextBox[] txt = new TextBox[12] { txt_Position1 , txt_Position2 , txt_Position3 , txt_Position4 , txt_Position5 , txt_Position6 ,
                            txt_Position7 , txt_Position8 ,txt_Position9,txt_Position10,txt_Position11,txt_Position12};

            for (int i = 0; i < txt.Length; i++)
            {
                txt[i].Text = "";
            }
        }

        /// <summary>
        /// txt_Position1 的 MouseClick 事件處理函式。<br/>
        /// 當使用者在 txt_Position1 輸入框點擊時，停止 DispatcherTimer，避免自動清空儲位輸入框。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_Position1。</param>
        /// <param name="e">事件參數，包含滑鼠點擊資訊。</param>
        private void txt_Position1_MouseClick(object sender, MouseEventArgs e)
        {
            // 停止計時器，避免自動清空儲位輸入框
            dispatcherTimer.Stop();
        }

        /// <summary>
        /// txt_itemE_mark 的 KeyPress 事件處理函式。<br/>
        /// 當使用者在 txt_itemE_mark 輸入框按下 Enter 鍵時，
        /// 會呼叫 <see cref="mark_KeyPress"/> 方法，查詢並顯示該料號的明細資料。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_itemE_mark。</param>
        /// <param name="e">事件參數，包含按下的鍵。</param>
        private void txt_itemE_mark_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否按下 Enter 鍵 (KeyChar 13)
            if (e.KeyChar == 13)
            {
                // 呼叫 mark_KeyPress 方法查詢並顯示明細資料
                mark_KeyPress();
            }
        }

        /// <summary>
        /// dataGridView4 的儲存格點擊事件。
        /// 當使用者點擊第 0 欄時，會將焦點移至第 9 欄（Mark 欄位），並進入編輯模式。
        /// 其他欄位則不執行任何動作。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 dataGridView4。</param>
        /// <param name="e">事件參數，包含點擊的儲存格位置。</param>
        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    // 取得目前選取的列索引
                    int s = dataGridView4.CurrentRow.Index;
                    // 設定焦點到 dataGridView4
                    dataGridView4.Focus();
                    // 將目前儲存格設為第 9 欄（Mark 欄）
                    dataGridView4.CurrentCell = dataGridView4[9, s];
                    // 進入編輯模式
                    dataGridView4.BeginEdit(true);
                    break;
                default:
                    // 其他欄位不執行任何動作
                    break;
            }
        }

        /// <summary>
        /// dataGridView4 的編輯控制項顯示事件處理函式。
        /// 當使用者在 dataGridView4 的儲存格進入編輯狀態時觸發。
        /// 僅當編輯的欄位名稱為 "m_Mark" 時，
        /// 會將 flag 設為 true，並將焦點移至 txt_mark，
        /// 並將目前選取列的 Mark 欄位值帶入 txt_mark 輸入框。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 dataGridView4。</param>
        /// <param name="e">事件參數，包含編輯控制項。</param>
        private void dataGridView4_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            // 判斷觸發事件的欄位是 Mark 才進行以下動作
            if (((DataGridView)sender).Columns[((DataGridView)sender).CurrentCell.ColumnIndex].Name == "m_Mark")
            {
                // 設定 flag 為 true，表示可以更新 Mark
                flag = true;
                // 將焦點移至 txt_mark 輸入框
                txt_mark.Focus();
                // 將目前選取列的 Mark 欄位值帶入 txt_mark 輸入框
                txt_mark.Text = dataGridView4.CurrentRow?.Cells[9].Value?.ToString() ?? string.Empty;
            }
        }

        /// <summary>
        /// btn_update_mark_Click 事件處理函式。<br/>
        /// 當使用者按下「更新備註」按鈕時，
        /// 會將目前選取的明細資料 (sno) 的 Mark 欄位更新為 txt_mark 的內容，
        /// 並同步更新 Automatic_Storage_Detail 表的 Mark 欄位。
        /// 更新完成後會重新查詢並顯示明細資料。
        /// 若未選擇可編輯欄位則顯示提示訊息。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為「更新備註」按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void btn_update_mark_Click(object sender, EventArgs e)
        {
            // 檢查是否有選擇可編輯的欄位
            if (flag)
            {
                // 取得目前選取列的 sno 值
                string id = dataGridView4.CurrentRow?.Cells[1].Value?.ToString() ?? string.Empty;
                string strsql = @"update Automatic_Storage_Input set Mark=@mark where sno=@sno";
                SqlParameter[] sqlParameters = new SqlParameter[]
                {
                    new SqlParameter("mark",txt_mark.Text.Trim()),
                    new SqlParameter("sno",id)
                };
                db.ExecueNonQuery(strsql, CommandType.Text, sqlParameters);

                #region Detail表同步更新
                //string id = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value.ToString();
                string strsql2 = @"update Automatic_Storage_Detail set Mark=@mark where sno=@sno";
                SqlParameter[] sqlParameters2 = new SqlParameter[]
                {
                    new SqlParameter("mark",txt_mark.Text.Trim()),
                    new SqlParameter("sno",id)
                };
                db.ExecueNonQuery(strsql2, CommandType.Text, sqlParameters2);
                #endregion
                mark_KeyPress();
            }
            else
            {
                MessageBox.Show("未選擇更新欄位");
            }
        }

        /// <summary>
        /// inputExcelBW 的 ProgressChanged 事件處理函式。<br/>
        /// 當背景工作進度改變時，更新 progressBar1 的值以反映目前進度。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 inputExcelBW。</param>
        /// <param name="e">事件參數，包含目前進度百分比。</param>
        private void inputExcelBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // 更新進度條的值
            progressBar1.Value = e.ProgressPercentage;
        }

        /// <summary>
        /// 按下「新增訊息」按鈕時，將使用者輸入的料號、客戶料號、訊息內容新增至 Automatic_Storage_Msg 資料表。
        /// 若料號欄位為空則不執行新增。
        /// 新增完成後會重新繫結訊息資料至 dataGridView5 顯示。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為「新增訊息」按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void btn_insert_Click(object sender, EventArgs e)
        {
            // 檢查料號欄位是否為空
            bool flag = false;
            flag = (!string.IsNullOrEmpty(txt_msgitemE.Text)) ? true : false;

            if (flag)
            {
                string sqlmsg = "insert into Automatic_Storage_Msg (item_E,item_C,msg) values (@itemE,@itemC,@Msg) ";
                SqlParameter[] sqls = new SqlParameter[]
                {
                    new SqlParameter("itemE",txt_msgitemE.Text.Trim()),
                    new SqlParameter("itemC",txt_msgitemC.Text.Trim()),
                    new SqlParameter("Msg",txt_msg.Text.Trim()),
                };
                db.ExecueNonQuery(sqlmsg, CommandType.Text, sqls);
            }
            DataMsg();
        }

        /// <summary>
        /// txt_msgitemE 的 Leave 事件處理函式。
        /// 當使用者離開 txt_msgitemE 輸入框時，
        /// 會自動查詢 Automatic_Storage_Spec 資料表，
        /// 若找到對應的料號，則自動填入客戶料號至 txt_msgitemC。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_msgitemE。</param>
        /// <param name="e">事件參數。</param>
        private void txt_msgitemE_Leave(object sender, EventArgs e)
        {
            #region 自動填入
            string sqlitem = "select item_E,item_C,Spec from Automatic_Storage_Spec where item_E = @item";
            SqlParameter[] parm = new SqlParameter[]
            {
                new SqlParameter("item",txt_msgitemE.Text.Trim())
            };
            DataSet ds = db.ExecuteDataSet(sqlitem, CommandType.Text, parm);
            if (ds.Tables[0].Rows.Count > 0)
            {
                txt_msgitemC.Text = ds.Tables[0].Rows[0]["item_C"].ToString();
            }
            #endregion
        }

        /// <summary>
        /// btn_search_Click 事件處理函式。<br/>
        /// 根據使用者輸入的料號、客戶料號、訊息內容查詢 Automatic_Storage_Msg 資料表，
        /// 並將查詢結果繫結至 dataGridView5 顯示，同時將第一筆訊息內容顯示於 txt_msg。
        /// 查詢條件：
        /// - 若 txt_msgitemE 有值，則以「昶亨料號」查詢。
        /// - 若 txt_msgitemC 有值，則以「客戶料號」查詢。
        /// - 若 txt_msg 有值，則以「訊息內容」查詢。
        /// 查詢完成後顯示「更新」按鈕。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為查詢按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void btn_search_Click(object sender, EventArgs e)
        {
            List<SqlParameter> parmM = new List<SqlParameter>();
            string sqlmsg = @"select item_E as '昶亨料號',item_C as '客戶料號',msg as '訊息內容' from Automatic_Storage_Msg where 1=1";
            if (!string.IsNullOrEmpty(txt_msgitemE.Text))
            {
                sqlmsg += @" and item_E=@itemE ";
                parmM.Add(new SqlParameter("itemE", txt_msgitemE.Text));
            }
            if (!string.IsNullOrEmpty(txt_msgitemC.Text))
            {
                sqlmsg += @" and item_C=@itemC ";
                parmM.Add(new SqlParameter("itemC", txt_msgitemC.Text));
            }
            if (!string.IsNullOrEmpty(txt_msg.Text))
            {
                sqlmsg += @" and msg=@Msg ";
                parmM.Add(new SqlParameter("Msg", txt_msg.Text));
            }

            DataSet dsmsg = db.ExecuteDataSetPmsList(sqlmsg, CommandType.Text, parmM);
            txt_msg.Text = dsmsg.Tables[0].Rows[0]["訊息內容"].ToString();
            DataTable dt = dsmsg.Tables[0];
            dataGridView5.DataSource = dt;
            btn_update.Visible = true;
        }

        /// <summary>
        /// dataGridView5 的儲存格點擊事件。
        /// 當使用者點擊第 0 欄時，會詢問是否刪除該筆訊息資料，
        /// 若確認則刪除對應的 Automatic_Storage_Msg 資料表資料，
        /// 並重新繫結訊息資料至 dataGridView5 顯示。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 dataGridView5。</param>
        /// <param name="e">事件參數，包含點擊的儲存格位置。</param>
        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 判斷點擊的是否為第 0 欄
            switch (e.ColumnIndex)
            {
                case 0:
                    // 顯示刪除確認對話框
                    DialogResult res = MessageBox.Show("是否刪除( " + (dataGridView5.CurrentRow?.Cells[1].Value?.ToString() ?? string.Empty) + " )",
                                "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    // 若使用者確認刪除
                    if (res == DialogResult.Yes)
                    {
                        string itemE = dataGridView5.CurrentRow?.Cells[1].Value?.ToString() ?? string.Empty;
                        string itemC = dataGridView5.CurrentRow?.Cells[2].Value?.ToString() ?? string.Empty;
                        string msg = dataGridView5.CurrentRow?.Cells[3].Value?.ToString() ?? string.Empty;
                        string sql_d = "delete Automatic_Storage_Msg " +
                            "where msg = @Msg and Item_E = @itemE and Item_C=@itemC ";
                        SqlParameter[] parmD = new SqlParameter[]
                        {
                            new SqlParameter("itemE",itemE),
                            new SqlParameter("itemC",itemC),
                            new SqlParameter("Msg",msg),
                        };
                        db.ExecueNonQuery(sql_d, CommandType.Text, parmD);
                        DataMsg();
                    }
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// btn_update_Click 事件處理函式。<br/>
        /// 當使用者按下「更新訊息」按鈕時，
        /// 會將目前輸入的訊息內容 (msg) 更新至 Automatic_Storage_Msg 資料表中，
        /// 條件為 item_E 與 item_C 皆符合。<br/>
        /// 更新完成後會隱藏「更新」按鈕，並清空相關輸入欄位，
        /// 並重新查詢訊息資料顯示於 dataGridView5。
        /// 若發生例外則顯示錯誤訊息。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為「更新訊息」按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void btn_update_Click(object sender, EventArgs e)
        {
            try
            {
                string sql_update = @"update Automatic_Storage_Msg set msg=@Msg 
                                        where item_E=@itemE and item_C=@itemC";
                SqlParameter[] sqls = new SqlParameter[]
                {
                    new SqlParameter("Msg",txt_msg.Text.Trim()),
                    new SqlParameter("itemE",txt_msgitemE.Text.Trim()),
                    new SqlParameter("itemC",txt_msgitemC.Text.Trim())
                };
                db.ExecueNonQuery(sql_update, CommandType.Text, sqls);
                btn_update.Visible = false;
                txt_msg.Text = "";
                txt_msgitemE.Text = "";
                txt_msgitemC.Text = "";
                DataMsg();
            }
            catch (Exception xx)
            {
                MessageBox.Show(xx.Message);
            }
        }

        /// <summary>
        /// inputExcelBW_RunWorkerCompleted 事件處理函式。<br/>
        /// 當 Excel 匯入背景作業完成時觸發，執行以下動作：
        /// 1. 設定進度條值為匯入筆數。
        /// 2. 顯示「執行完成」訊息。
        /// 3. 啟用 commitButton，允許再次匯入。
        /// 4. 清空 dsData 資料集。
        /// 5. 清空 txt_path 路徑欄位。
        /// 6. 重新繫結規格資料至 gv_Spec。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 inputExcelBW。</param>
        /// <param name="e">事件參數，包含作業完成狀態。</param>
        private void inputExcelBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = inputexcelcount;
            //progressBar1.Maximum = 100;
            MessageBox.Show("執行完成");
            commitButton.Enabled = true;
            dsData.Clear();// 清空資料集
            txt_path.Text = "";
            gv_Spec_Data();
        }

        /// <summary>
        /// 根據輸入的料號，查詢 Automatic_Storage_Detail 資料表，
        /// 並將結果顯示於 dataGridView4。
        /// 1. 清空明細資料表 dataT。
        /// 2. 以料號 (Item_No_Master) 為條件，查詢入庫明細資料，僅顯示庫存量大於 0 的資料。
        /// 3. 將查詢結果逐筆加入 dataT。
        /// 4. 將 dataT 資料繫結至 dataGridView4 顯示。
        /// 5. 若查詢過程發生例外，則顯示錯誤訊息。
        /// </summary>
        private void mark_KeyPress()
        {
            try
            {
                dataT.Clear();
                string strsql = @"select a.sno,cast(Actual_InDate as date)InDate,Item_No_Master,Spec,Position,b.Package,Up_InDate,c.User_name,a.Mark  
                                    from Automatic_Storage_Detail a 
                                    left join Automatic_Storage_Package b on a.Package = b.code 
                                    left join Automatic_Storage_User c on a.Input_UserNo = c.User_No 
                                    left join Automatic_Storage_User d on a.Output_UserNo = d.User_No
                                    where a.Unit_No='M' and  Item_No_Master = @master and amount >0 
                                    order by position asc";

                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("unitno",Login.Unit_No),
                    new SqlParameter("master",txt_itemE_mark.Text.Trim())
                };
                DataSet dt = db.ExecuteDataSet(strsql, CommandType.Text, parameters);
                foreach (DataRow row in dt.Tables[0].Rows)
                {
                    DataRow dr = dataT.NewRow();
                    dr["sno"] = row["sno"].ToString();
                    dr["Actual_InDate"] = row["InDate"].ToString();
                    dr["Item_No_Master"] = row["Item_No_Master"].ToString();
                    dr["Spec"] = row["Spec"].ToString();
                    dr["Position"] = row["Position"].ToString();
                    dr["Package"] = row["Package"].ToString();
                    dr["Up_InDate"] = row["Up_InDate"].ToString();
                    dr["User_name"] = row["User_name"].ToString();
                    dr["Mark"] = row["Mark"].ToString();
                    dataT.Rows.Add(dr);
                }
                dataGridView4.DataSource = dataT;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #region 未使用
        /// <summary>
        /// btn_spec_Click 事件處理函式。<br/>
        /// 當使用者按下「新增規格」按鈕時，
        /// 會將輸入的昶亨料號、客戶料號、規格等資料新增至 Automatic_Storage_Spec 資料表。<br/>
        /// 新增完成後可呼叫 gv_Spec_Data() 重新繫結規格資料顯示於 gv_Spec。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為「新增規格」按鈕。</param>
        /// <param name="e">事件參數。</param>
        private void btn_spec_Click(object sender, EventArgs e)
        {
            //string sql_i = "INSERT INTO Automatic_Storage_Spec ([Item_E],[Item_C],[Spec],[Unit_No],[Cr_User],[Cr_Date]) " +
            //                        "VALUES(@itemE,@itemC,@spec,@unitno,@cruser,@crdate)";
            //SqlParameter[] parmI = new SqlParameter[]
            //{
            //    new SqlParameter("itemE",txt_itemE_i.Text.Trim()),
            //    new SqlParameter("itemC",txt_itemC_i.Text.Trim()),
            //    new SqlParameter("spec",txt_spec_i.Text.Trim()),
            //    new SqlParameter("unitno",Login.Unit_No),
            //    new SqlParameter("cruser",Login.User_No),
            //    new SqlParameter("crdate",DateTime.Now)
            //};
            //db.ExecueNonQuery(sql_i, CommandType.Text, parmI);
            ////gv_Spec_Data();
        }

        /// <summary>
        /// Txt_itemE_mark_KeyPress 事件處理函式。<br/>
        /// 當使用者在 txt_itemE_mark 輸入框按下按鍵時觸發。<br/>
        /// 目前尚未實作，請依需求補充功能。
        /// </summary>
        /// <param name="sender">事件來源物件，通常為 txt_itemE_mark。</param>
        /// <param name="e">事件參數，包含按下的鍵。</param>
        private void Txt_itemE_mark_KeyPress(object sender, KeyPressEventArgs e)
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}
