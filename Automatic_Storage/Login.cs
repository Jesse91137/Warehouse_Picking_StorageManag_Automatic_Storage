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
    public partial class Login : Form
    {
        public Login()
        {
            InitializeComponent();
        }
        public static string Role_ID = string.Empty;
        public static string User_No = string.Empty;
        public static string User_name = string.Empty;
        public static string Unit_No = string.Empty;

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter) //判斷是否輸入 enter
            {
                txt_psw.Focus();
            }
        }

        private void txt_psw_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter) //判斷是否輸入 enter
            {
                string sqlstr = "select * from Automatic_Storage_User WHERE User_No = @User_No and Password= @psw";

                SqlParameter[] parm = new SqlParameter[]
                {
                    new SqlParameter("User_No",txt_id.Text),
                    new SqlParameter("psw",txt_psw.Text)
                };
                DataSet ds = db.ExecuteDataSet(sqlstr, CommandType.Text, parm);
                if (ds.Tables[0].Rows.Count > 0)
                {
                    string sqlRole = @"select * from Automatic_Storage_UserRole where USER_ID=@User_No and ROLE_ID='3' ";
                    SqlParameter[] parm2 = new SqlParameter[]
                    {
                        new SqlParameter("User_No",txt_id.Text)
                    };
                    DataSet dsRole = db.ExecuteDataSet(sqlRole, CommandType.Text, parm2);
                    if (dsRole.Tables[0].Rows.Count > 0)
                    {
                        Role_ID = dsRole.Tables[0].Rows[0]["ROLE_ID"].ToString();
                    }

                    User_No = ds.Tables[0].Rows[0]["User_No"].ToString();
                    User_name = ds.Tables[0].Rows[0]["User_Name"].ToString();
                    Unit_No = ds.Tables[0].Rows[0]["Unit_No"].ToString();
                    
                    this.Hide();
                }
                else
                {
                    loginresult.Text="帳號密碼錯誤 或 無權限登錄！";
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
