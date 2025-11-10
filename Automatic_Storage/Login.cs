using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace Automatic_Storage
{
    public partial class Login : Form
    {
        /// <summary>
        /// 建構函式，初始化登入視窗元件
        /// </summary>
        public Login()
        {
            InitializeComponent(); // 初始化元件
        }

        /// <summary>
        /// 使用者角色編號，登入後儲存於此
        /// </summary>
        public static string Role_ID = string.Empty; // 使用者角色編號

        /// <summary>
        /// 使用者編號，登入後儲存於此
        /// </summary>
        public static string User_No = string.Empty; // 使用者編號

        /// <summary>
        /// 使用者姓名，登入後儲存於此
        /// </summary>
        public static string User_name = string.Empty; // 使用者姓名

        /// <summary>
        /// 單位編號，登入後儲存於此
        /// </summary>
        public static string Unit_No = string.Empty; // 單位編號

        /// <summary>
        /// 處理使用者在帳號輸入框按下鍵盤按鍵的事件
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">鍵盤按鍵事件參數</param>
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否輸入 enter
            if (e.KeyChar == (char)Keys.Enter)
            {
                // 將焦點移至密碼輸入框
                txt_psw.Focus();
            }
        }

        /// <summary>
        /// 使用者登入KeyPress事件
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">鍵盤按鍵事件參數</param>
        private void txt_psw_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 判斷是否輸入 enter
            if (e.KeyChar == (char)Keys.Enter)
            {
                // SQL查詢語句，根據帳號及密碼查詢使用者
                string sqlstr = "select * from Automatic_Storage_User WHERE User_No = @User_No and Password= @psw";

                // 建立SQL參數陣列
                SqlParameter[] parm = new SqlParameter[]
                {
                    new SqlParameter("User_No",txt_id.Text), // 帳號參數
                    new SqlParameter("psw",txt_psw.Text) // 密碼參數
                };

                // 執行SQL查詢，取得結果
                DataSet ds = db.ExecuteDataSet(sqlstr, CommandType.Text, parm);

                // 判斷是否有查詢到資料
                if (ds.Tables[0].Rows.Count > 0)
                {
                    // 查詢使用者角色是否為3
                    string sqlRole = @"select * from Automatic_Storage_UserRole where USER_ID=@User_No and ROLE_ID='3' ";
                    SqlParameter[] parm2 = new SqlParameter[]
                    {
                        new SqlParameter("User_No",txt_id.Text) // 帳號參數
                    };

                    // 執行角色查詢
                    DataSet dsRole = db.ExecuteDataSet(sqlRole, CommandType.Text, parm2);

                    // 若有角色資料則儲存角色編號
                    if (dsRole.Tables[0].Rows.Count > 0)
                    {
                        Role_ID = dsRole.Tables[0].Rows[0]["ROLE_ID"].ToString(); // 角色編號
                    }

                    // 儲存使用者編號
                    User_No = ds.Tables[0].Rows[0]["User_No"].ToString();
                    // 儲存使用者姓名
                    User_name = ds.Tables[0].Rows[0]["User_Name"].ToString();
                    // 儲存單位編號
                    Unit_No = ds.Tables[0].Rows[0]["Unit_No"].ToString();

                    // 隱藏登入視窗
                    this.Hide();
                }
                else
                {
                    // 顯示錯誤訊息
                    loginresult.Text = "帳號密碼錯誤 或 無權限登錄！";
                }
            }
        }

        /// <summary>
        /// 處理離開按鈕點擊事件，結束應用程式
        /// </summary>
        /// <param name="sender">事件來源物件</param>
        /// <param name="e">事件參數</param>
        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit(); // 結束應用程式
        }
    }
}
