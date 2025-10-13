using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Automatic_Storage
{
    public partial class OutPut : Form
    {
        public OutPut()
        {
            InitializeComponent();
            isLoaded = false;
        }
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

        #region 批次出庫來的參數
        private string strItem;         //料號
        private string strSpec;         //規格
        private string strPackage;  //包裝
        private string strWono="";  //單號
        private bool strBat = false;             //是否批次來

        public string Item
        {
            set { strItem = value; }
        }
        public string Spec
        {
            set { strSpec = value; }
        }
        public string Package
        {
            set { strPackage = value; }
        }
        public string Wono
        {
            set { strWono = value; }
        }
        public bool Bat
        {
            set { strBat = value; }
        }
        public void setValue()
        {
            textBox1.Text = strItem;
            //textBox2.Text = strItem;
            string sSpec = strSpec;
            string sWono = strWono;
            cbx_package.SelectedItem = strPackage;
            bool sBat = strBat;
        }
        #endregion
        #region 視窗ReSize
        int X = new int();  //窗口寬度
        int Y = new int(); //窗口高度
        float fgX = new float(); //寬度縮放比例
        float fgY = new float(); //高度縮放比例
        bool isLoaded;  // 是否已設定各控制的尺寸資料到Tag屬性
        #endregion

        System.ComponentModel.NullableConverter nullableDateTime =
                new System.ComponentModel.NullableConverter(typeof(DateTime?));        
        string actualDate=string.Empty;
        string sno = string.Empty;

        double inventory = 0;
        double Amount = 0;
        string txt_dulCheck = string.Empty;
        string txt_search =string.Empty;
        string strShift = string.Empty;
        DataTable dt = new DataTable();
        string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar==13)
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
                if (dt.Rows.Count>0)
                {
                    dataGridView1.DataSource = dt;
                    dataGridView1.Visible = true;
                    label5.Visible = false;
                    textBox2.Focus();
                    
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]["Amount"].ToString()=="0")
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

        private void OutPut_Load(object sender, EventArgs e)
        {
            #region 獲取窗體info
            X = this.Width;//獲取窗體的寬度
            Y = this.Height;//獲取窗體的高度
            isLoaded = true;// 已設定各控制項的尺寸到Tag屬性中
            SetTag(this);//調用方法
            #endregion
            
            //Form_Load 設定包裝種類cbox
            //textBox3.Visible = false;
            string strsql = "select Package_View,code from Automatic_Storage_Package";
            DataSet cbx_ds = db.ExecuteDataSet(strsql, CommandType.Text, null);
            foreach (DataRow dr in cbx_ds.Tables[0].Rows)
            {
                cbx_package.Items.Add(new MyItem(dr["Package_View"].ToString(), dr["code"].ToString()));
            }
            //判斷是否來自批次視窗
            if (strBat)
            {
                string sqlstr = "select Actual_InDate,Item_No_Master ,Item_No_Slave ,Amount ,Position ,Package,Spec  ,Mark ,Sno ,Reel_ID,PCB_DC,CMC_DC  " +                                        
                                        "from Automatic_Storage_Detail " +
                                        "where Unit_No = @unitNo and Item_No_Master = @Item and Spec = @Spec " +
                                        "and Amount > 0  order by position asc";
                SqlParameter[] parm = new SqlParameter[]
                {
                    new SqlParameter("unitNo",Login.Unit_No),
                    new SqlParameter("Item",strItem),
                    new SqlParameter("Spec",strSpec)
                };
                dataGridView1.Visible = true;
                //textBox3.Visible = true;
                //bat_confirm.Visible = true;
                textBox4.Text = "";
                //txt_Amount_U.Text = "";
                dt = db.ExecuteDataTable(sqlstr, CommandType.Text, parm);
                dataGridView1.DataSource = dt;

            }            
        }
        int cp = -1;
        //↓料號確認
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13 && !strBat)
            {
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    txt_dulCheck = textBox1.Text.Trim().ToUpper();
                }
                if (!string.IsNullOrEmpty(txt_itemC.Text))
                {
                    txt_dulCheck = txt_itemC.Text.Trim().ToUpper();
                }
                if (!Debit()) return;

                textBox1.Focus();
                textBox1.Text = "";
                textBox2.Text = "";
                textBox3.Text = "";
                txt_Mark.Text = "";
                //txt_Amount_U.Text = "";
                textBox4.Text = "";
                label10.Text = "";
                cbx_package.SelectedIndex = -1;
                dataGridView1.DataSource = dataBind();
            }
            if (e.KeyChar==13 && strBat)
            {
                bat_confirm_Click(sender, e);
            }
            
        }
        public void lb4Change()
        {
            switch (cp)
            {
                case 0:
                    if (dt.Rows.Count>0)
                    {
                        //微軟正黑體, 12pt, style=Bold
                        //label4.Font = new Font("微軟正黑體", 18, FontStyle.Bold);//微軟正黑體, 12pt, style=Bold
                        //label4.ForeColor = Color.Black;
                        //label4.Text = "儲位確認";
                        //textBox3.Visible = true;
                        textBox4.Focus();
                    }
                    else
                    {
                        textBox1.Text = "";
                        textBox2.Text = "";
                        textBox1.Focus();
                    }
                    break;                
                default:
                    textBox3.Visible = false;
                    label4.Font = new Font("微軟正黑體", 18, FontStyle.Bold);
                    label4.ForeColor = Color.Red;
                    label4.Text = "料卷有誤,請再次確認!!";
                    break;
            }
            
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                string auditText = textBox2.Text.Trim().ToUpper();
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    cp = textBox1.Text.Trim().ToUpper().CompareTo(auditText);
                }
                else if (!string.IsNullOrEmpty(txt_itemC.Text))
                {
                    cp = txt_itemC.Text.Trim().ToUpper().CompareTo(auditText);
                }

                //if (cp == 0)
                //{
                //    lb4Change();

                //    //this.SelectNextControl(this.ActiveControl, true, true, true, false); //跳下一個元件
                //}                
                //lb4Change();
            }
        }
        //text3box.keypress動作的扣帳程序
        public bool Debit()
        {
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                txt_dulCheck = textBox1.Text.Trim().ToUpper();
            }
            if (!string.IsNullOrEmpty(txt_itemC.Text))
            {
                txt_dulCheck = txt_itemC.Text.Trim().ToUpper();
            }
            //Convert.ToDouble(dt.Rows[0]["Amount"])
            if (Amount > inventory)
            {
                label10.Text = "總數量有誤,請重新確認!";
                return false;
            }
            if (Convert.ToDouble(textBox4.Text) > inventory)
            {
                label10.Text = "總數量有誤,請重新確認!";
                return false;
            }
            if (txt_dulCheck != textBox2.Text.Trim().ToUpper())
            {
                label10.Text = "料號確認有誤,請重新確認!";
                return false;
            }
            try
            {
                MyItem myItem = (MyItem)this.cbx_package.SelectedItem;
                //Insert Out_Storage
                string sqlOut = @"insert into Automatic_Storage_Output (Sno,Item_No_Master,Item_No_Slave,Spec,Position,
                                                       Amount,Package,Unit_No,Output_UserNo,Output_Date,Wo_No,Mark,Reel_ID,PCB_DC,CMC_DC ) 
                                                       select Sno,Item_No_Master,Item_No_Slave,Spec,Position,@amount,Package,
                                                       @unitno,@outuser,@outdate,@wono,@mark,@reelid,PCB_DC,CMC_DC from Automatic_Storage_Detail ";
                sqlOut += @"where Sno=@sno and amount > 0";
                //if (!string.IsNullOrEmpty(textBox1.Text))
                //{
                //    sqlOut += @" where Item_No_Master=@master and Position=@position and package=@package
                //                                        and Actual_InDate=@actualDate and Mark = @mark and amount > 0";
                //}
                //if (!string.IsNullOrEmpty(txt_itemC.Text))
                //{
                //    sqlOut += @" where Item_No_Slave=@master and Position=@position and package=@package
                //                                        and Actual_InDate=@actualDate and Mark = @mark and amount > 0";
                //}
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
                        //new SqlParameter("master",textBox2.Text.Trim()),
                        //new SqlParameter("position",textBox3.Text.Trim()),
                        //new SqlParameter("package",myItem.value),
                        //new SqlParameter("actualDate",actualDate),
                        //new SqlParameter("mark",actualDate)
                };
                //if (db.ExecueNonQuery(sqlOut, CommandType.Text, parm1) == 0)
                if (db.ExecueNonQuery(sqlOut, CommandType.Text,"單筆出庫text確認", parm1) == 0)
                {
                    cp = -1;
                    lb4Change();
                    return false;
                }

                //Update Storage_Detail                
                string sqlDetail = @"update Automatic_Storage_Detail 
                                                           set Amount = Amount-@amount ,
                                                           Up_OutDate = @outdate ,Output_UserNo = @outuser ,Mark = @mark ,Reel_ID = @reelid ";
                if (!string.IsNullOrEmpty(textBox1.Text))
                {
                    sqlDetail += @"where Item_No_Master = @master and Position=@position and Package=@package 
                                                  and Actual_InDate=@actualDate and sno=@sno ";
                }
                if (!string.IsNullOrEmpty(txt_itemC.Text))
                {
                    sqlDetail += @"where Item_No_Slave = @master and Position=@position and Package=@package 
                                                and Actual_InDate=@actualDate and sno=@sno";
                }
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
                db.ExecueNonQuery(sqlDetail, CommandType.Text, "單筆出庫Detail更新", parm2);                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return true;
        }
        public DataTable dataBind()
        {
            string sqlstr = "select Actual_InDate,Item_No_Master ,Item_No_Slave ,Amount ,Position ,Package ,Spec ,Mark ,Sno ,Reel_ID,PCB_DC,CMC_DC  " +
                                        "from Automatic_Storage_Detail ";
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                sqlstr += "where Unit_No = @unitNo and Item_No_Master = @Item " +
                                        "and Amount > 0  order by Actual_InDate asc";
            }
            else if (!string.IsNullOrEmpty(txt_itemC.Text)) 
            {
                sqlstr += "where Unit_No = @unitNo and Item_No_Slave = @Item " +
                                        "and Amount > 0  order by Actual_InDate asc";
            }

            SqlParameter[] parm = new SqlParameter[]
            {
                    new SqlParameter("unitNo",Login.Unit_No),
                    new SqlParameter("Item",txt_search)
            };
            return db.ExecuteDataTable(sqlstr, CommandType.Text, parm);
        }

        private void OutPut_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.Text=="Form1")
            {
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
        private void OutPut_Resize(object sender, EventArgs e)
        {
            if (X == 0 || Y == 0) return;
            fgX = (float)this.Width / (float)X;
            fgY = (float)this.Height / (float)Y;

            SetControls(fgX, fgY, this);
        }
        private void OutPut_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        private void txt_Engsr_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

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

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
                if (e.KeyChar == 13)
                {
                    TextBox txb = (TextBox)sender;
                    switch (txb.Name)
                    {
                        case "textBox4":
                            if (!string.IsNullOrEmpty(textBox4.Text))
                            {
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
                    if (Amount > Convert.ToDouble(dt.Rows[0]["Amount"]))
                    {
                        label10.Text = "總數量有誤,請重新確認!";
                        return;
                    }
                }
            }            
        }

        private void textBox4_Leave(object sender, EventArgs e)
        {
            TextBox txb = (TextBox)sender;
            switch (txb.Name)
            {
                case "textBox4":
                    if (!string.IsNullOrEmpty(textBox4.Text))
                    {
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
            if (Amount > Convert.ToDouble(dt.Rows[0]["Amount"]))
            {
                label10.Text = "總數量有誤,請重新確認!";
                return;
            }
        }

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
            if (!Debit()) return;
            father.MsgFromChildChk = textBox4.Text;
            father.MsgFromChildItemE = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            father.MsgFromChildButton = true;
            this.Close();
        }

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
            inventory = Convert.ToDouble(dataGridView1.CurrentRow.Cells[3].Value);
            textBox3.Text=dataGridView1.CurrentRow.Cells[4].Value.ToString();
            txt_Mark.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            actualDate = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[0].Value.ToString()).ToString("yyyy-MM-dd");
            sno= dataGridView1.CurrentRow.Cells[8].Value.ToString();
            
            string strpackage = "select id from Automatic_Storage_Package " +
                "where code = '" + dataGridView1.CurrentRow.Cells[5].Value.ToString() + "'";
            DataSet dscpk = db.ExecuteDataSet(strpackage, CommandType.Text, null);
            cbx_package.SelectedIndex = Convert.ToInt32(dscpk.Tables[0].Rows[0]["id"])-1;

            visible_Panel.Visible = false;
            textBox4.Focus();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {            
            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                txt_itemC.Enabled = false;
            }
            else if (!string.IsNullOrEmpty(txt_itemC.Text))
            {
                textBox1.Enabled = false;
            }
        }

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
