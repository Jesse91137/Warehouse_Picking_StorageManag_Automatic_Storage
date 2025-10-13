using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Automatic_Storage
{
    public partial class ActualDate : Form
    {
        public ActualDate()
        {
            InitializeComponent();
        }
        public string Sno { get; set; }
        public string AcD_O { get; set; }
        public string AcD_M { get; set; }

        private void btn_actualDate_Click(object sender, EventArgs e)
        {
            try
            {
                AcD_M = dateTimePicker1.Value.Date.ToString("yyyy-MM-dd");
                TimeSpan Diff_dates = dateTimePicker1.Value.Date.Subtract(DateTime.Today);
                if (Diff_dates.TotalDays>=1)
                {
                    MessageBox.Show("日期錯誤!! 請再次確認。");
                    return;
                }
                //index
                string sql = @"update Automatic_Storage_Detail set Actual_InDate=@InDate where sno = @Sno ";
                SqlParameter[] paramers = new SqlParameter[]
                {
                new SqlParameter("InDate",AcD_M),
                new SqlParameter("Sno",Sno),
                };
                db.ExecueNonQuery(sql, CommandType.Text, paramers);

                //history
                string sql_i = @"update Automatic_Storage_Input set Actual_InDate=@InDate where sno = @Sno ";
                SqlParameter[] paramers_i = new SqlParameter[]
                {
                new SqlParameter("InDate",AcD_M),
                new SqlParameter("Sno",Sno),
                };
                db.ExecueNonQuery(sql_i, CommandType.Text, paramers_i);

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ActualDate_Load(object sender, EventArgs e)
        {
            textBox1.Text = AcD_O;
        }

        private void ActualDate_FormClosing(object sender, FormClosingEventArgs e)
        {
            Form1 ower = (Form1)this.Owner;
            ower.refreshData();
        }
    }
}
