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
    public partial class Input : Form
    {
        public Input()
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
        #region 視窗ReSize
        int X = new int();  //窗口寬度
        int Y = new int(); //窗口高度
        float fgX = new float(); //寬度縮放比例
        float fgY = new float(); //高度縮放比例
        bool isLoaded;  // 是否已設定各控制的尺寸資料到Tag屬性
        #endregion
        _Logger logger = new _Logger();
        private void button2_Click(object sender, EventArgs e)
        {
            //logger.LogEvent(((Button)(sender)).Text, DateTime.Now,"");
            Form1 ower = (Form1)this.Owner;
            ower.refreshData();
            this.Close();
        }

        
        private void txt_Item1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar==13 && !string.IsNullOrEmpty(ActiveControl.Text))
            {                
                string auditCMC = txt_Item1.Text[txt_Item1.Text.Length - 4].ToString().ToUpper();

                if (auditCMC == "A")
                {
                    labCMC.Visible = true;
                }
                txt_Amount.Focus();
                //cbx_package.Focus();
            }
        }

        private void Input_FormClosing(object sender, FormClosingEventArgs e)
        {
            Form1 ower = (Form1)this.Owner;
            ower.refreshData();            
        }

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
        private void init_Data()
        {
            string strsql = "select Package_View,code from Automatic_Storage_Package";
            DataSet cbx_ds = db.ExecuteDataSet(strsql, CommandType.Text, null);
            foreach (DataRow dr in cbx_ds.Tables[0].Rows)
            {
                cbx_package.Items.Add(new MyItem(dr["Package_View"].ToString(), dr["code"].ToString()));
            }
        }
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
            //txtbox1.Text = "1";
            //txt_Amount_U.Text = "1";
            txt_Item1.Enabled = true;
            txt_item2.Enabled = true;
            txt_Item1.Focus();

        }

        
        //private void styleTXT(string str)
        //{
        //    foreach (Control ctrl in Controls)
        //    {
        //        if (ctrl is TextBox && str=="only")
        //        {
        //            ctrl.Enabled = false;
        //        }
        //        else
        //        {
        //            ctrl.Enabled = true;
        //        }
        //    }            
        //    btn_only_commit.Visible = true;
        //    txt_Item1.Enabled = true;
        //    txt_Amount1.Enabled = true;
        //    txt_Storage1.Enabled = true;
        //}
        
        
        private void btn_only_commit_Click(object sender, EventArgs e)
        {
            string inputDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string sqlinput = string.Empty;
            string sqldetail = string.Empty;
            string sqlchkFirst = string.Empty;
            
            try
            {
                TextBox[] b = new TextBox[1] { txt_Storage1 }; //儲位
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
                                string actualDate= dt_Picker.Value.ToString("yyyy/MM/dd");                                
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
                                //                                and Package = @Package and Actual_InDate = @actualDate";
                                //}
                                //else if (!string.IsNullOrEmpty(slave))
                                //{
                                //    sqlchkFirst += @"where Item_No_Slave=@item_ms and Position=@Position 
                                //                                and Package = @Package and Actual_InDate = @actualDate";
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
                                    //if (txt_Item1.Enabled)
                                    //{
                                    //    sqldetail += @"where Item_No_Master=@item_ms and Position=@Position 
                                    //                                and Package=@Package and Actual_InDate=@actualDate and Mark = @mark";
                                    //}
                                    //if (txt_item2.Enabled)
                                    //{
                                    //    sqldetail += @"where Item_No_Slave=@item_ms and Position=@Position 
                                    //                                    and Package=@Package and Actual_InDate=@actualDate";
                                    //}
                                    //if (!string.IsNullOrEmpty(master))
                                    //{ 
                                    //    sqldetail+= @"where Item_No_Master=@item_ms and Position=@Position 
                                    //                                and Package=@Package and Actual_InDate=@actualDate"; 
                                    //}
                                    //else if (!string.IsNullOrEmpty(slave))
                                    //{
                                    //    sqldetail += @"where Item_No_Slave=@item_ms and Position=@Position 
                                    //                                and Package=@Package and Actual_InDate=@actualDate";
                                    //}
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
                    if (dsP.Tables[0].Rows.Count ==0)
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

                //Response.Write(" <script>window.opener = null;window.open('', '_parent', '');window.self.close();</script>");
            }
            catch (Exception ex)
            {
                lbl_result.Text = ex.Message.ToString();
            }
        }
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

        private void Input_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        private void Input_Resize(object sender, EventArgs e)
        {
            if (X == 0 || Y == 0) return;
            fgX = (float)this.Width / (float)X;
            fgY = (float)this.Height / (float)Y;

            SetControls(fgX, fgY, this);
        }

        private void txt_Amount_U_TextChanged(object sender, EventArgs e)
        {
            //string nS1 = (string.IsNullOrEmpty(txtbox1.Text)) ? "0" : txtbox1.Text;
            //string nS2 = (string.IsNullOrEmpty(txt_Amount_U.Text)) ? "0" : txt_Amount_U.Text;

            //Double nI1 = Convert.ToDouble(nS1);
            //Double nI2 = Convert.ToDouble(nS2);

            //Double result = nI1 * nI2;
            //txt_Amount.Text = result.ToString();
        }

        private void txtbox1_MouseClick(object sender, MouseEventArgs e)
        {
            TextBox text = (TextBox)sender;
            text.Text = "";
        }

        private void txt_Item1_Leave(object sender, EventArgs e)
        {
            string sqlitem = "";
            string sqlpositionList = "";
            string ItemChoice = "";
            string msgcheck = "";
            Label[] lbl = new Label[18] { lab_p1, lab_p2, lab_p3, lab_p4, lab_p5, lab_p6, lab_p7, lab_p8, lab_p9, lab_p10,
                                            lab_p11,lab_p12,lab_p13,lab_p14,lab_p15,lab_p16,lab_p17,lab_p18};

            if (!string.IsNullOrEmpty(txt_Item1.Text.Trim()))
            {
                string auditCMC = txt_Item1.Text[txt_Item1.Text.Length - 4].ToString().ToUpper();

                if (auditCMC == "A")
                {
                    labCMC.Visible = true;
                }
            }
            
            if (string.IsNullOrEmpty(txt_Item1.Text) && string.IsNullOrEmpty(txt_item2.Text))
            {
                lbl_result.Text = "料號尚未輸入!!";
            }                        
            else
            {
                if (!string.IsNullOrEmpty(txt_Item1.Text))
                {
                    sqlitem = "select item_E,item_C,Spec from Automatic_Storage_Spec where item_E = @item";
                    sqlpositionList = "select Position from Automatic_Storage_Detail where Item_No_Master = @item and amount>0  group by Position";
                    ItemChoice = txt_Item1.Text;
                    msgcheck = "select * from Automatic_Storage_Msg where item_E =@item ";
                }
                if (!string.IsNullOrEmpty(txt_item2.Text))
                {
                    sqlitem = "select item_E,item_C,Spec from Automatic_Storage_Spec where item_C =@item";
                    sqlpositionList = "select Position from Automatic_Storage_Detail where Item_No_Slave = @item and amount>0  group by Position";
                    ItemChoice = txt_item2.Text;
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
                DataSet ds = db.ExecuteDataSet(sqlitem, CommandType.Text, parm);
                DataSet ds2 = db.ExecuteDataSet(sqlpositionList, CommandType.Text, parm2);
                DataSet ds3 = db.ExecuteDataSet(msgcheck, CommandType.Text, parm3);

                if (ds.Tables[0].Rows.Count > 0)
                {
                    txt_Spec.Text = ds.Tables[0].Rows[0]["Spec"].ToString();
                    if (!string.IsNullOrEmpty(txt_Item1.Text))
                    {
                        txt_item2.Text= ds.Tables[0].Rows[0]["item_C"].ToString();
                    }
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

                if (ds2.Tables[0].Rows.Count > 0)
                {
                    for (int lb = 0; lb < ds2.Tables[0].Rows.Count; lb++)
                    {
                        lbl[lb].Text = ds2.Tables[0].Rows[lb][0].ToString();
                    }
                    
                }
            }
            //return;
        }

        private void txtbox1_Leave(object sender, EventArgs e)
        {
            //if (string.IsNullOrEmpty(txtbox1.Text))
            //{
            //    MessageBox.Show("包裝數量尚未輸入!");
            //    txtbox1.Focus();
            //}
            //if (string.IsNullOrEmpty(txt_Amount_U.Text))
            //{
            //    MessageBox.Show("單位數量尚未輸入!");
            //    txt_Amount_U.Focus();
            //}

        }

        private void txtbox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
        }
        private void _KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                this.SelectNextControl(this.ActiveControl, true, true, true, false); //跳下一個元件
            }
        }

        private void txt_Amount_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                this.SelectNextControl(this.ActiveControl, true, true, true, false); //跳下一個元件
            }
        }

        private void txt_Storage1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                this.SelectNextControl(this.ActiveControl, true, true, true, false); //跳下一個元件
            }
        }

        private void txt_Item1_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txt_item2.Text))
            {
                lbl_result.Text = "";
            }
            if (txt_Item1.Focused)
            {
                txt_item2.Enabled = false;
            }
            else if (txt_item2.Focused)
            {
                txt_Item1.Enabled = false;
            }
        }

        private void cbx_package_SelectedIndexChanged(object sender, EventArgs e)
        {
            txt_Storage1.Focus();
        }

        private void dt_Picker_ValueChanged(object sender, EventArgs e)
        {
            TimeSpan timeSpan = dt_Picker.Value.Date.Subtract(DateTime.Today);
            if ((timeSpan.TotalDays >= 1))
            {
                MessageBox.Show("入庫日期錯誤!! 按下確認後系統將帶入今天日期。");
            }
            dt_Picker.Value = (timeSpan.TotalDays >= 1) ? DateTime.Now : dt_Picker.Value;
        }
    }
}
