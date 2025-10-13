using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Windows.Forms;
using System.Windows.Threading;

namespace Automatic_Storage
{
    public partial class Maintain : Form
    {
        public class MyItem
        {
            public string text;
            public string value;

            public MyItem(string text, string value)
            {
                this.text = text;
                this.value = value;
            }
            public override string ToString()
            {
                return text;
            }
        }
        public Maintain()
        {
            InitializeComponent();
            isLoaded = false;
            dd = dd ?? columnsData();
        }
        int inputexcelcount = 0, btn_fAll = 0, btn_itemP = 0;
        int initSum = 0;
        DataSet dsData = new DataSet();
        DispatcherTimer dispatcherTimer = new DispatcherTimer();
        #region 視窗ReSize
        int X = new int();  //窗口寬度
        int Y = new int(); //窗口高度
        float fgX = new float(); //寬度縮放比例
        float fgY = new float(); //高度縮放比例
        bool isLoaded;  // 是否已設定各控制的尺寸資料到Tag屬性
        #endregion

        private DataTable columnsData()
        {
            using (DataTable table = new DataTable())
            {
                // Add columns.
                table.Columns.Add("sno", typeof(string));
                table.Columns.Add("UNIT_NAME", typeof(string));
                table.Columns.Add("User_name", typeof(string));
                table.Columns.Add("Position", typeof(string));
                table.Columns.Add("Create_Date", typeof(string));
                return table;
            }
        }
        private DataTable columnsDataTable()
        {
            using (DataTable data = new DataTable())
            {
                // Add columns.
                data.Columns.Add("sno", typeof(string));
                data.Columns.Add("Actual_InDate", typeof(DateTime));
                data.Columns.Add("Item_No_Master", typeof(string));
                data.Columns.Add("Spec", typeof(string));                
                data.Columns.Add("Position", typeof(string));
                data.Columns.Add("Package", typeof(string));
                data.Columns.Add("Up_InDate", typeof(string));
                data.Columns.Add("User_name", typeof(string));
                data.Columns.Add("Mark", typeof(string));
                return data;
            }
        }

        static string time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        public DataTable dd;
        public DataTable dataT=new DataTable();
        bool flag = false;
        static string position_sno;
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch ((sender as TabControl).SelectedIndex)
            {
                case 1:
                    try
                    {
                        checkedListBox1.Items.Clear();
                        dataBind();                        
                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.Message);
                    }
                    break;
                case 2:
                    try
                    {
                        gv_Spec_Data();
                    }
                    catch (Exception ee)
                    {
                        MessageBox.Show(ee.Message);
                    }
                    break;
                case 3:
                    dataT = columnsDataTable();
                    
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
                    chk_dataBind();
                    userDataBind();
                    break;
                case 5:
                    DataMsg();
                    break;

                default:
                    break;
            }
        }
        private void userDataBind()
        {
            string sql_user = @"SELECT User_No as '工號',User_name as '姓名' from Automatic_Storage_User where User_No<>'02437' order by User_No";
            DataTable dt_user = db.ExecuteDataTable(sql_user, CommandType.Text, null);
            dataGridView3.DataSource = dt_user;
        }
        private void dataBind()
        {
            dd.Clear();
            string strsql = @"select a.Sno as sno,UNIT_NAME,User_name,Position,Create_Date from Automatic_Storage_Position a
                                                         left join Automatic_Storage_UnitNo b on a.Unit_No=b.UNIT_NO ,Automatic_Storage_User c 
                                                         where a.Create_User=c.User_No and a.Unit_No=@unitno ";

            SqlParameter[] parameters = { new SqlParameter("unitno", Login.Unit_No) };
            DataSet dt_all = db.ExecuteDataSet(strsql, CommandType.Text, parameters);

            foreach (DataRow row in dt_all.Tables[0].Rows)
            {
                DataRow dr = dd.NewRow();
                dr["sno"] = row["sno"].ToString();
                dr["UNIT_NAME"] = row["UNIT_NAME"].ToString();
                dr["User_name"] = row["User_name"].ToString();
                dr["Position"] = row["Position"].ToString();
                dr["Create_Date"] = row["Create_Date"].ToString();
                dd.Rows.Add(dr);
            }
            dataGridView2.DataSource = dd;
        }        
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
                 new SqlParameter("unitno",Login.Unit_No),
                new SqlParameter("userNo",txt_userid.Text.Trim()),
                new SqlParameter("name",txt_name.Text.Trim())
             };
            ddt = db.ExecuteDataTable(strsql, CommandType.Text, parm);
            dataGridView1.DataSource = ddt;
        }

        private void btn_submit_Click(object sender, EventArgs e)
        {
            try
            {
                TimerStart();
                TextBox[] txt = new TextBox[12] { txt_Position1 , txt_Position2 , txt_Position3 , txt_Position4 , txt_Position5 , txt_Position6 ,
                            txt_Position7 , txt_Position8 ,txt_Position9,txt_Position10,txt_Position11,txt_Position12};
                for (int i = 0; i < txt.Length; i++)
                {
                    if (txt[i].Text != "")
                    {
                        string str_s = @"select * from Automatic_Storage_Position where Unit_No=@unitno and Position=@position";
                        SqlParameter[] parm_s = new SqlParameter[]
                        {
                                new SqlParameter("unitno",Login.Unit_No),
                                new SqlParameter("position",txt[i].Text.Trim()),
                        };
                        DataSet dataSet= db.ExecuteDataSet(str_s, CommandType.Text, parm_s);
                        if (dataSet.Tables[0].Rows.Count == 0)
                        {
                            string strsql = @"insert into Automatic_Storage_Position (Unit_No,Position,Create_User,Create_Date) values
                                                        (@unitno, @position, @cr_user, @cr_date)";
                            SqlParameter[] parameters = new SqlParameter[]
                            {
                                    new SqlParameter("unitno",Login.Unit_No),
                                    new SqlParameter("position",txt[i].Text.Trim().ToUpper()),
                                    new SqlParameter("cr_user",Login.User_No),
                                    new SqlParameter("cr_date",time)
                            };
                            db.ExecueNonQuery(strsql, CommandType.Text, parameters);
                            lab_finish.Text = "儲位新增完成";
                        }
                        else
                        {
                            MessageBox.Show("第 " + (i + 1) + " 儲位重複,請再次確認");
                        }                        
                    }                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }           

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    //(sender as DataGridView).Rows.RemoveAt(e.RowIndex);
                    int s = dataGridView2.CurrentRow.Index;
                    //MessageBox.Show(dataGridView2.Rows[s].Cells[2].Value.ToString());
                    dataGridView2.Focus();
                    dataGridView2.CurrentCell = dataGridView2[5, s];
                    dataGridView2.BeginEdit(true);
                    break;
                case 1:
                    delete_Click();
                    break;
                default:
                    break;
            }
        }

        private void dataGridView2_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            // 判斷觸發事件的欄位是 Position 才進行以下動作
            if (((DataGridView)sender).Columns[((DataGridView)sender).CurrentCell.ColumnIndex].Name == "Position")
            {
                flag = true;
                txt_new_p.Focus();
                txt_new_p.Text= dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[5].Value.ToString();
                position_sno = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value.ToString();
                //TextBox txt = e.Control as TextBox;
                //if (txt != null)
                //{
                //    // 增加事件
                //    //txt.KeyPress += new KeyPressEventHandler(txt_KeyPress);
                //}
            }
        }

        private void txt_position_s_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar==13)
            {
                try
                {
                    dd.Clear();
                    string strsql = @"select a.Sno as sno,UNIT_NAME,User_name,Position,Create_Date from Automatic_Storage_Position a
                                                 left join Automatic_Storage_UnitNo b on a.Unit_No=b.UNIT_NO ,Automatic_Storage_User c 
                                                 where a.Unit_No='M' and a.Create_User=c.User_No and a.Unit_No=@unitno and Position=@position ";
                    SqlParameter[] parameters = new SqlParameter[]
                    {
                    new SqlParameter("unitno",Login.Unit_No),
                    new SqlParameter("position",txt_position_s.Text.Trim())
                    };
                    DataSet dt = db.ExecuteDataSet(strsql, CommandType.Text, parameters);
                    foreach (DataRow row in dt.Tables[0].Rows)
                    {
                        DataRow dr = dd.NewRow();
                        dr["sno"] = row["sno"].ToString();
                        dr["UNIT_NAME"] = row["UNIT_NAME"].ToString();
                        dr["User_name"] = row["User_name"].ToString();
                        dr["Position"] = row["Position"].ToString();
                        dr["Create_Date"] = row["Create_Date"].ToString();
                        dd.Rows.Add(dr);
                    }
                    dataGridView2.DataSource = dd;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (flag)
            {
                //string id=dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value.ToString();
                string strchk = @"select * from Automatic_Storage_Position where Position =@position";
                SqlParameter[] param = new SqlParameter[]
                {
                    new SqlParameter("position",txt_new_p.Text.Trim())
                    
                };
                DataSet data= db.ExecuteDataSet(strchk, CommandType.Text, param);
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
                else
                {
                    string delsql = @"delete from Automatic_Storage_Position where sno=@sno";
                    SqlParameter[] delparam = new SqlParameter[]
                    {                    
                        new SqlParameter("sno",position_sno)
                    };
                    db.ExecueNonQuery(delsql, CommandType.Text, delparam);
                }
                
                dataBind();
                #region Detail表同步更新
                //string id = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value.ToString();
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
                MessageBox.Show("未選擇更新欄位");
            }
        }
        private void delete_Click()
        {

            #region 刪除不驗證 for 小鄭 20210522
            string id = dataGridView2.Rows[dataGridView2.CurrentRow.Index].Cells[2].Value.ToString();
            string strsql = @"delete Automatic_Storage_Position where sno=@sno";
            SqlParameter[] sqlParameters = new SqlParameter[]
            {
                    new SqlParameter("sno",id)
            };
            db.ExecueNonQuery(strsql, CommandType.Text, sqlParameters);
            dataBind();
            #endregion

            #region 刪除有驗證 保留版
            //確認Detail_Table儲位.count是否>0
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

        private void txt_Position1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                this.SelectNextControl(this.ActiveControl, true, true, true, false); //跳下一個元件
            }
        }
        private void chk_dataBind() 
        {
            checkedListBox1.Items.Clear();
            string sqlcb = @"select ROLE_ID,ROLE_NAME from Automatic_Storage_Role ";
            DataSet ds= db.ExecuteDataSet(sqlcb, CommandType.Text, null);
            
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                checkedListBox1.Items.Add(new MyItem(row["ROLE_NAME"].ToString(), row["ROLE_ID"].ToString()));
            }
        }

        private void btn_role_add_Click(object sender, EventArgs e)
        {
            if (checkedListBox1.CheckedItems.Count>0)
            {
                //檢查人員規則是否已存在
                string sql_search = @"select * from Automatic_Storage_UserRole where USER_ID=@userid ";
                SqlParameter[] parm_se = new SqlParameter[]
                {
                    new SqlParameter("userid",txt_userid.Text.Trim())
                };
                DataSet data = db.ExecuteDataSet(sql_search, CommandType.Text, parm_se);
                if (data.Tables[0].Rows.Count == 0)
                {
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        MyItem item = (MyItem)this.checkedListBox1.Items[i];
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
                    string sql_del = @"delete from Automatic_Storage_UserRole where USER_ID=@userid";
                    SqlParameter[] parm_role2 = new SqlParameter[]
                    {
                        new SqlParameter("userid",txt_userid_role.Text.Trim())
                     };
                    db.ExecueNonQuery(sql_del, CommandType.Text, parm_role2);
                    btn_role_add_Click(sender, e);
                    MessageBox.Show("人員權限更新完成");
                }
            }
            else
            {
                MessageBox.Show("至少選擇一項權限");
            }
        }

        private void btn_user_add_Click(object sender, EventArgs e)
        {
            try
            {
                string sql_search = @"select * from Automatic_Storage_User where User_No=@userid ";
                SqlParameter[] parm_se = new SqlParameter[]
                {
                    new SqlParameter("userid",txt_userid.Text.Trim())
                };
                DataSet data = db.ExecuteDataSet(sql_search, CommandType.Text, parm_se);
                if (data.Tables[0].Rows.Count == 0)
                {
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
                    MessageBox.Show("不可重複新增人員");
                }
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.Message);                
            }
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    string id = dataGridView3.Rows[dataGridView3.CurrentRow.Index].Cells[1].Value.ToString();
                    string strsql_u = @"delete Automatic_Storage_User where User_No=@userno";
                    SqlParameter[] parm_u = new SqlParameter[]
                    {
                        new SqlParameter("userno",id)
                    };
                    db.ExecueNonQuery(strsql_u, CommandType.Text, parm_u);
                    string strsql_r = @"delete from Automatic_Storage_UserRole where USER_ID=@userno";
                    SqlParameter[] parm_r = new SqlParameter[]
                    {
                        new SqlParameter("userno",id)
                    };
                    db.ExecueNonQuery(strsql_r, CommandType.Text, parm_r);

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
        private void Maintain_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        private void Maintain_Resize(object sender, EventArgs e)
        {
            if (X == 0 || Y == 0) return;
            fgX = (float)this.Width / (float)X;
            fgY = (float)this.Height / (float)Y;

            SetControls(fgX, fgY, this);
        }

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
        private void DataMsg()
        {
            string sqlmsgs = @" select item_E as '昶亨料號',item_C as '客戶料號',msg as '訊息內容' 
                                                            from Automatic_Storage_Msg ";

            DataTable dtmg = db.ExecuteDataTable(sqlmsgs, CommandType.Text, null);
            dataGridView5.DataSource = dtmg;
        }
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
        private void btn_Spec_f_Click(object sender, EventArgs e)
        {
            List<SqlParameter> parmF = new List<SqlParameter>();
            string sqlF = "select Item_E as '昶亨料號' ,Item_C as '客戶料號' ,Spec as '規格' from Automatic_Storage_Spec " +
                "where Unit_no=@unitno ";
            parmF.Add(new SqlParameter("unitno", Login.Unit_No));
            if (!string.IsNullOrEmpty(txt_f_item_E.Text))
            {
                sqlF += " and Item_E like @itemE +'%' ";
                parmF.Add(new SqlParameter("itemE", txt_f_item_E.Text.Trim()));
            }
            if (!string.IsNullOrEmpty(txt_f_item_C.Text))
            {
                sqlF += " and Item_C like @itemC +'%' ";
                parmF.Add(new SqlParameter("itemC", txt_f_item_C.Text.Trim()));
            }
            if (!string.IsNullOrEmpty(txt_f_item_Spec.Text))
            {
                sqlF += " and Spec like @spec +'%' ";
                parmF.Add(new SqlParameter("spec", txt_f_item_Spec.Text.Trim()));
            }
            DataSet ds = db.ExecuteDataSetPmsList(sqlF, CommandType.Text, parmF);
            DataTable dt = ds.Tables[0];  //每次能讀取一張表
            gv_Spec.DataSource = dt;                        
        }

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
        private OpenFileDialog fileDialog1;
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
                    if ((string.IsNullOrEmpty(ItemE) && string.IsNullOrEmpty(ItemC) && string.IsNullOrEmpty(Spec)))
                    {
                        _ = (count[i] == 0) ? list_result.Items.Add("第" + (i + 1) + "筆上傳失敗").ToString() : "";
                    }
                    else
                    {
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
                    progressBar1.PerformStep();
                }
            }
            catch (Exception ex)
            {
                list_result.Items.Add(ex.Message);
            }
            
            ////gv_Spec_Data();
        }
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

            }
        }

        private void gv_Spec_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    DialogResult res = MessageBox.Show("是否刪除( " + gv_Spec.Rows[gv_Spec.CurrentRow.Index].Cells[1].Value.ToString() + " )",
                                "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (res == DialogResult.Yes)
                    {
                        string itemE = gv_Spec.Rows[gv_Spec.CurrentRow.Index].Cells[1].Value.ToString();
                        string itemC = gv_Spec.Rows[gv_Spec.CurrentRow.Index].Cells[2].Value.ToString();
                        string spec = gv_Spec.Rows[gv_Spec.CurrentRow.Index].Cells[3].Value.ToString();
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

        private void txt_Position1_TextChanged(object sender, EventArgs e)
        {
            lab_finish.Text = "";
        }

        public void TimerStart()
        {
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 2);
            dispatcherTimer.Start();
        }
        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            TextBox[] txt = new TextBox[12] { txt_Position1 , txt_Position2 , txt_Position3 , txt_Position4 , txt_Position5 , txt_Position6 ,
                            txt_Position7 , txt_Position8 ,txt_Position9,txt_Position10,txt_Position11,txt_Position12};
            for (int i = 0; i < txt.Length; i++)
            {
                txt[i].Text = "";
            }
        }

        private void txt_Position1_MouseClick(object sender, MouseEventArgs e)
        {
            dispatcherTimer.Stop();
        }

        private void txt_itemE_mark_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                mark_KeyPress();
            }
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    //(sender as DataGridView).Rows.RemoveAt(e.RowIndex);
                    int s = dataGridView4.CurrentRow.Index;
                    //MessageBox.Show(dataGridView2.Rows[s].Cells[2].Value.ToString());
                    dataGridView4.Focus();
                    dataGridView4.CurrentCell = dataGridView4[9, s];
                    dataGridView4.BeginEdit(true);
                    break;                
                default:
                    break;
            }
        }

        private void dataGridView4_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            // 判斷觸發事件的欄位是 Mark 才進行以下動作
            if (((DataGridView)sender).Columns[((DataGridView)sender).CurrentCell.ColumnIndex].Name == "m_Mark")
            {
                flag = true;
                txt_mark.Focus();
                txt_mark.Text = dataGridView4.Rows[dataGridView4.CurrentRow.Index].Cells[9].Value.ToString();
            }
        }

        private void btn_update_mark_Click(object sender, EventArgs e)
        {
            if (flag)
            {
                string id = dataGridView4.Rows[dataGridView4.CurrentRow.Index].Cells[1].Value.ToString();
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

        private void Txt_itemE_mark_KeyPress(object sender, KeyPressEventArgs e)
        {
            throw new NotImplementedException();
        }

        private void inputExcelBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }

        private void btn_insert_Click(object sender, EventArgs e)
        {
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

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            switch (e.ColumnIndex)
            {
                case 0:
                    DialogResult res = MessageBox.Show("是否刪除( " + dataGridView5.Rows[dataGridView5.CurrentRow.Index].Cells[1].Value.ToString() + " )",
                                "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (res == DialogResult.Yes)
                    {
                        string itemE = dataGridView5.Rows[dataGridView5.CurrentRow.Index].Cells[1].Value.ToString();
                        string itemC = dataGridView5.Rows[dataGridView5.CurrentRow.Index].Cells[2].Value.ToString();
                        string msg = dataGridView5.Rows[dataGridView5.CurrentRow.Index].Cells[3].Value.ToString();
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

        private void inputExcelBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = inputexcelcount;
            //progressBar1.Maximum = 100;
            MessageBox.Show("執行完成");
            commitButton.Enabled =true;
            dsData.Clear();
            txt_path.Text = "";
            gv_Spec_Data();
        }
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
    }
}
