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
    public partial class BatOutView : Form
    {
        private string strItem;
        private string strEngsr;

        public string Item
        {
            set { strItem = value; }
        }
        //public string Engsr
        //{
        //    set { strEngsr = value; }
        //}

        public void setValue()
        {
            string sItem = strItem;
            //string sEng = strEngsr;
        }
        public BatOutView()
        {
            InitializeComponent();
            isLoaded = false;
        }

        #region 視窗ReSize
        int X = new int();  //窗口寬度
        int Y = new int(); //窗口高度
        float fgX = new float(); //寬度縮放比例
        float fgY = new float(); //高度縮放比例
        bool isLoaded;  // 是否已設定各控制的尺寸資料到Tag屬性
        #endregion

        string txt_search =string.Empty;
        string strShift = string.Empty;
        DataTable dt = new DataTable();
        string dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        
        private void OutPut_Load(object sender, EventArgs e)
        {
            #region 獲取窗體info
            X = this.Width;//獲取窗體的寬度
            Y = this.Height;//獲取窗體的高度
            isLoaded = true;// 已設定各控制項的尺寸到Tag屬性中
            SetTag(this);//調用方法
            #endregion

            #region LoadBind
            dataGridView1.DataSource=dataBind(strItem);
            #endregion

        }
        int cp = -1;
                        
        public DataTable dataBind(string item)
        {
            string sqlstr = "select Item_No_Master ,Item_No_Slave ,Spec " +
                                        ",Amount_Unit ,Amount ,Position ,Package  " +
                                        "from Automatic_Storage_Detail " +
                                        "where Unit_No = @unitNo and Item_No_Master = @Item " +
                                        "and Amount > 0  order by position asc";
            SqlParameter[] parm = new SqlParameter[]
            {
                    new SqlParameter("unitNo",Login.Unit_No),
                    new SqlParameter("Item",strItem)
            };
            return db.ExecuteDataTable(sqlstr, CommandType.Text, parm);
        }
                     
        private void OutPut_FormClosing(object sender, FormClosingEventArgs e)
        {
                                    
        }
        #region 視窗設定
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
        #endregion

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {            
            //this.Close();
        }
        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            


        }
        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // 0→料號; 1→位置; 2→機種; 3→數量;   

            BatOut father = (BatOut)this.Owner;
            
            int c_index = dataGridView1.CurrentCell.ColumnIndex;
            
            switch (c_index)
            {
                case 1:
                    if (!string.IsNullOrEmpty(dataGridView1.CurrentRow.Cells[3].Value.ToString()))
                    {
                        if (Int32.Parse(dataGridView1.CurrentRow.Cells[3].Value.ToString()) > 0)
                        {
                            DialogResult res = MessageBox.Show("是否拿取( "+ dataGridView1.CurrentRow.Cells[1].Value.ToString()+" )位置１捲",
                                "請點兩下取出庫", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (res==DialogResult.Yes)
                            {
                                father.MsgFromChildPosition = dataGridView1.CurrentCell.Value.ToString();
                                this.Close();
                            }
                            else
                            {
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show("數量小於０不可取出庫!!");
                            return;
                        }
                    }
                    else
                    {
                        return;
                    }
                    break;
                case 2:
                    if (!string.IsNullOrEmpty(dataGridView1.CurrentRow.Cells[2].Value.ToString()))
                    {
                        DialogResult res2 = MessageBox.Show("是否線上拿取( " + dataGridView1.CurrentRow.Cells[2].Value.ToString() + " )位置１捲",
                                "請點兩下取出庫", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (res2 == DialogResult.Yes)
                        {
                            father.MsgFromChildPosition = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                            father.MsgFromChildEngSr = dataGridView1.CurrentCell.Value.ToString();
                            this.Close();
                        }
                    }
                    else
                    {
                        MessageBox.Show("機種名稱空白，未在線上!!");
                        return;
                    }                    
                    break;
                default:
                    father.MsgFromChildPosition = string.Empty;
                    father.MsgFromChildEngSr = string.Empty;
                    break;
            }            
        }
    }
}
