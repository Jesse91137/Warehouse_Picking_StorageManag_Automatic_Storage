using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace Automatic_Storage
{
    public partial class Form1 : Form
    {
        public Form1()
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

        int inputexcelcount = 0, btn_fAll = 0, btn_itemP = 0;
        int initSum = 0;
        bool rad_choice = false;
        DataTable dt = new DataTable();
        DataSet dsData = new DataSet();
        static Mutex m;
        DispatcherTimer dispatcherTimer = new DispatcherTimer();
        public string sop_path = "", SopName = "", version_old = "", version_new = "";
        public string filename = "Setup.ini";
        SetupIniIP ini = new SetupIniIP();
        Log GetLog = new Log();
        int pageSize = 100;
        int currentPage = 1; // 默认第一页
        CMC_DC cmc = new CMC_DC();
        public class SetupIniIP
        { //api ini
            public string path;
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern long WritePrivateProfileString(string section,
            string key, string val, string filePath);
            [DllImport("kernel32", CharSet = CharSet.Unicode)]
            private static extern int GetPrivateProfileString(string section,
            string key, string def, StringBuilder retVal,
            int size, string filePath);
            public void IniWriteValue(string Section, string Key, string Value, string inipath)
            {
                WritePrivateProfileString(Section, Key, Value, Application.StartupPath + "\\" + inipath);
            }
            public string IniReadValue(string Section, string Key, string inipath)
            {
                StringBuilder temp = new StringBuilder(255);
                int i = GetPrivateProfileString(Section, Key, "", temp, 255, Application.StartupPath + "\\" + inipath);
                return temp.ToString();
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            X = this.Width;//獲取窗體的寬度
            Y = this.Height;//獲取窗體的高度
            isLoaded = true;// 已設定各控制項的尺寸到Tag屬性中
            SetTag(this);//調用方法

            this.Text += "  " + FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location).FileVersion.ToString();
            Form.CheckForIllegalCrossThreadCalls = false;
            if (IsMyMutex("Automatic_Storage"))
            {
                MessageBox.Show("程式正在執行中!!");
                Dispose();//關閉
            }
            version_old = ini.IniReadValue("Version", "version", filename);
            version_new = selectVerSQL_new("Auto_Storage_M");
            //lbl_ver.Text = "VER:V" + version_old;
            //判斷自動更新程式是否啟動
            int v_old = Convert.ToInt32(version_old.Replace(".", ""));
            int v_new = Convert.ToInt32(version_new.Replace(".", ""));
            if (v_old < v_new)
            {
                MessageBox.Show("有新版本更新VER: V" + version_new);
                autoupdate();
            }
            else
            {
                panel1.Visible = false;
                panel5.Visible = false;
                Login login = new Login();
                login.Owner = this;
                login.ShowDialog();

                Initi();
                dataBind();


            }
            this.Text += "       " + Login.User_name;
        }
        //private void LoadPage()
        //{
        //    DataTable dt = GetPagedData(currentPage);
        //    dataGridView1.DataSource = dt;
        //    //lblPageNumber.Text = $"Page {currentPage}";
        //}
        private bool IsMyMutex(string prgname)
        {
            bool IsExist;
            m = new Mutex(true, prgname, out IsExist);
            GC.Collect();
            if (IsExist)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        public void dataBind()
        {
            string sql = string.Empty;
            //int offset = (pageNumber - 1) * pageSize;

            sql = (Login.User_No == "02437")
                ? @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno    
                        from Automatic_Storage_Detail a
                        left join Automatic_Storage_Package b on a.Package=b.code
                        left join Automatic_Storage_User c on a.Input_UserNo=c.User_No 
                        left join Automatic_Storage_User d on a.Output_UserNo=d.User_No 
                        where amount >0 order by Actual_InDate"
                : @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno    
                        from Automatic_Storage_Detail a
                        left join Automatic_Storage_Package b on a.Package=b.code
                        left join Automatic_Storage_User c on a.Input_UserNo=c.User_No 
                        left join Automatic_Storage_User d on a.Output_UserNo=d.User_No 
                        where a.Unit_No=@unitNo and amount >0 order by Actual_InDate";
            if (btn_fAll == 1)
            {
                sql = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,Package,Up_InDate,Input_UserNo,Up_OutDate,Output_UserNo,Mark,PCB_DC,CMC_DC,sno from Automatic_Storage_Detail 
                        where Unit_No=@unitNo order by Actual_InDate";
            }
            //sql += " OFFSET @Offset ROWS " +
            //             " FETCH NEXT @PageSize ROWS ONLY ";
            SqlParameter[] parm = new SqlParameter[]
            {
                new SqlParameter("unitNo",Login.Unit_No),
                //new SqlParameter("Offset",offset),
                //new SqlParameter("PageSize",pageSize ),
            };
            dt = db.ExecuteDataTable(sql, CommandType.Text, parm);

            dataGridView1.DataSource = dt;

            sumC();
            //dataGridView1.AutoGenerateColumns = true;                        
            btn_fAll = 0;
        }
        public bool Week_Then25(string inputString)
        {
            // 假設輸入字串為 inputString
            //string inputString = textBox1.Text;

            DateTime inputDate;
            // 嘗試解析 YY/WW 格式
            if (inputString.Length == 5)
            {
                string[] y_w = inputString.Split('/');

                // 取得年份和週數
                int year = int.Parse(y_w[0]) + 2000; // 將兩位數年份轉換為四位數年份
                int week = int.Parse(y_w[1]);

                // 轉換為 DateTime 型別
                inputDate = new DateTime(year, 1, 1).AddDays((week - 1) * 7);

                // 取得指定年份的 1 月 1 日
                DateTime jan1 = new DateTime(year, 1, 1);

                // 取得指定年份的 1 月 1 日是星期幾
                int dayOfWeek = (int)jan1.DayOfWeek;

                // 計算指定週數的第一天
                int daysToAdd = (week - 1) * 7;
                daysToAdd -= dayOfWeek;
                inputDate = jan1.AddDays(daysToAdd);
            }
            // 嘗試解析 YYWW 格式
            else if (inputString.Length == 4 && int.TryParse(inputString, out int yearWeek))
            {
                int year = 2000 + (yearWeek / 100);
                int week = yearWeek % 100;
                inputDate = new DateTime(year, 1, 1).AddDays((week - 1) * 7);
            }
            // 嘗試解析 YYYYMMDD 格式
            else if (DateTime.TryParseExact(inputString, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out inputDate))
            {
                // 不需要進行額外的處理
            }
            // 嘗試解析 YYYY/MM/DD 格式
            else if (DateTime.TryParseExact(inputString, "yyyy/MM/dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out inputDate))
            {
                // 不需要進行額外的處理
            }
            //else
            //{
            //    label1.Text = ("無法解析輸入字串");
            //    return ;
            //}

            // 取得當天日期
            DateTime currentDate = DateTime.Today;

            // 計算當天日期與輸入日期的時間差
            TimeSpan diff = currentDate - inputDate;

            // 取得時間差的絕對值
            TimeSpan diffAbs = diff.Duration();

            // 計算時間差的週數
            int weeks = (int)Math.Floor(diffAbs.TotalDays / 7);

            // 判斷週數是否超過 25 週
            if (weeks > 25)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private void sumC()
        {
            //計算總數
            List<SqlParameter> parmC = new List<SqlParameter>();

            bool h_Flag = false;
            if (txt_itemP5.Text != "" && txt_itemP5.Text.Length >= 3)
            {
                h_Flag = true;
            }
            else if (txt_siteP5.Text != "")
            {
                h_Flag = true;
            }

            string sumC = "select sum(Amount) as C from Automatic_Storage_Detail where Unit_No=@unitNo and cast(Up_InDate as date) >'2023-06-30' ";

            string sumC_in = "select sum(Amount)   ";
            string sumC_Out = "select ISNULL(SUM(Amount), 0) as C from Automatic_Storage_Output where Unit_No=@unitNo and cast(Output_Date as date) >'2023-06-30' ";
            parmC.Add(new SqlParameter("unitNo", Login.Unit_No));

            if (h_Flag)
            {
                sumC = "";
                if (!string.IsNullOrEmpty(txt_itemP5.Text))
                {
                    if (rad_E.Checked)
                    {
                        sumC += " and Item_No_Master = @master ";
                    }
                    else
                    {
                        sumC += " and Item_No_Slave = @master ";
                    }
                    parmC.Add(new SqlParameter("master", txt_itemP5.Text.Trim()));
                }
                if (!string.IsNullOrEmpty(txt_siteP5.Text))
                {
                    sumC += " and Position = @position ";
                    parmC.Add(new SqlParameter("position", txt_siteP5.Text.Trim()));
                }
                string _in = " as C from Automatic_Storage_Input where Unit_No=@unitNo and cast(Input_Date as date) >'2023-06-30' ";
                sumC = sumC_in + " - ( " + sumC_Out + sumC + " ) " + _in + sumC;

            }
            else
            {
                if (!string.IsNullOrEmpty(txt_item.Text))
                {
                    sumC += " and item_no_master like @master+'%' ";
                    parmC.Add(new SqlParameter("master", txt_item.Text.Trim()));
                }
                if (!string.IsNullOrEmpty(txt_position.Text))
                {
                    sumC += " and Position like @position+'%' ";
                    parmC.Add(new SqlParameter("position", txt_position.Text));
                }
                if (!string.IsNullOrEmpty(txt_DTs.Text) && !string.IsNullOrEmpty(txt_DTe.Text))
                {
                    DateTime sdate = DateTime.ParseExact(txt_DTs.Text, "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                    DateTime edate = DateTime.ParseExact(txt_DTe.Text, "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                    string s_date = sdate.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    string e_date = edate.ToString("yyyy-MM-dd HH:mm:59.999");

                    if (rad_in.Checked)
                    {
                        sumC += " and amount > 0 and Up_InDate between @startDate and @endDate ";
                    }
                    else if (rad_out.Checked)
                    {
                        sumC += " and amount > 0 and Up_OutDate between @startDate and @endDate ";
                    }
                    parmC.Add(new SqlParameter("startDate", s_date.Trim()));
                    parmC.Add(new SqlParameter("endDate", e_date.Trim()));
                }
            }

            DataSet ds = db.ExecuteDataSetPmsList(sumC, CommandType.Text, parmC);
            txt_sumC.Text = ds.Tables[0].Rows[0]["C"].ToString();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Input input = new Input();
                input.Owner = this;
                input.Show();
                GetLog.WriteLog(((Button)(sender)).Name);
            }
            catch (Exception ee)
            {

                throw;
            }
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView1.Columns[0].HeaderText = "入庫日期";                //Actul_Date
            dataGridView1.Columns[0].DefaultCellStyle.Format = "yyyy-MM-dd";                //Actul_Date
            dataGridView1.Columns[1].HeaderText = "料號_昶";                //Item_No_Master
            dataGridView1.Columns[2].HeaderText = "料號_客";                //Item_No_Master
            dataGridView1.Columns[3].HeaderText = "規格";                //Spec
            dataGridView1.Columns[4].HeaderText = "儲位";                //Position            
            dataGridView1.Columns[5].HeaderText = "總數";                //Amount            
            dataGridView1.Columns[6].HeaderText = "包裝型態";       //Package            
            dataGridView1.Columns[7].HeaderText = "操作日期";
            dataGridView1.Columns[7].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";                //Actul_Date
            dataGridView1.Columns[8].HeaderText = "操作人員";
            dataGridView1.Columns[9].HeaderText = "出庫日期";
            dataGridView1.Columns[9].DefaultCellStyle.Format = "yyyy-MM-dd HH:mm:ss";                //Actul_Date
            dataGridView1.Columns[10].HeaderText = "出庫人員";
            dataGridView1.Columns[11].HeaderText = "備註";
            dataGridView1.Columns[12].HeaderText = "PCB";
            dataGridView1.Columns[13].HeaderText = "CMC";
            dataGridView1.Columns[14].Visible = false;
        }
        public void refreshData()
        {
            dataBind();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OutPut output = new OutPut();
            output.Owner = this;
            output.Show();
            dataBind();
        }
        //private Button selectButton;
        private OpenFileDialog fileDialog1;
        private void button3_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
            panel1.BringToFront();
            dispatcherTimer.Stop();
            fileDialog1 = new OpenFileDialog()
            {
                FileName = "Select a text file",
                Filter = "Text files (*.xls)|*.xls",
                Title = "Open text file"
            };

            //selectButton = new Button()
            //{
            //    Size = new Size(87,37),
            //    Location = new Point(270, 17),
            //    Font = new Font("微軟正黑體", 12, FontStyle.Bold),
            //    Text = "選擇檔案"
            //};

            //selectButton.Click += new EventHandler(selectButton_Click);
            //commitButton.Click += new EventHandler(commitButton_Click);
            //panel1.Controls.Add(selectButton);

        }
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
                dataBind();

            }
            catch (SecurityException ex)
            {
                MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                $"Details:\n\n{ex.StackTrace}");
            }

        }

        private void txt_item_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {

                bool h_Flag = false;
                string strShift = string.Empty;
                if (txt_itemP5.Text != "" && txt_itemP5.Text.Length >= 3)
                {
                    h_Flag = true;
                }

                if (txt_item.Text != "" && txt_item.Text.Length >= 3 || h_Flag)
                {
                    try
                    {
                        strShift = txt_item.Text.Trim().ToUpper();

                        //string stritem = "select Item_No_Master,Position,Amount,Up_InDate,Input_UserNo,Up_OutDate,Output_UserNo,Eng_SR " +
                        //"from Automatic_Storage_Detail where Unit_No=@unitNo and  Item_No_Master = @master and ( amount >0 " +
                        //" or eng_sr is not null )" +
                        //"order by eng_sr desc ,position asc";
                        //改模糊查詢(20210616 有一規則length需>3)此處改>=
                        string stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  " +
                        "from Automatic_Storage_Detail a " +
                        "left join Automatic_Storage_Package b on a.Package = b.code " +
                        "left join Automatic_Storage_User c on a.Input_UserNo = c.User_No " +
                        "left join Automatic_Storage_User d on a.Output_UserNo = d.User_No ";
                        if (rad_E.Checked)
                        {
                            stritem += "where a.Unit_No=@unitNo and  Item_No_Master like '%'+@master +'%' and amount >0 " +
                        "order by position asc";
                        }
                        else if (rad_C.Checked)
                        {
                            stritem += "where a.Unit_No=@unitNo and  Item_No_Slave like '%'+ @master +'%' and amount >0 " +
                        "order by position asc";
                        }

                        if (h_Flag)
                        {
                            if (rad_E.Checked)
                            {
                                stritem = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Input_Date,c.User_name,null,'',Mark,PCB_DC,CMC_DC,a.sno   
                                                        from Automatic_Storage_Input a 
                                                        left join Automatic_Storage_Package b on a.Package = b.code 
                                                        left join Automatic_Storage_User c on a.Input_UserNo = c.User_No 
                                                        where a.Unit_No=@unitNo and Item_No_Master=@master and cast(a.Input_Date as date) >'2023-06-30' 
                                                        union
                                                        select null,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,null,'',Output_Date,d.User_name,Mark,PCB_DC,CMC_DC,a.sno 
                                                        from Automatic_Storage_Output a
                                                        left join Automatic_Storage_Package b on a.Package = b.code  
                                                        left join Automatic_Storage_User d on a.Output_UserNo = d.User_No
                                                        where a.Unit_No=@unitNo and Item_No_Master=@master and cast(a.Output_Date as date) >'2023-06-30' ";
                            }
                            else if (rad_C.Checked)
                            {
                                stritem = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Input_Date,c.User_name,null,'',Mark,PCB_DC,CMC_DC,a.sno  
                                                        from Automatic_Storage_Input a 
                                                        left join Automatic_Storage_Package b on a.Package = b.code 
                                                        left join Automatic_Storage_User c on a.Input_UserNo = c.User_No 
                                                        where a.Unit_No=@unitNo and Item_No_Slave=@master and cast(a.Input_Date as date) >'2023-06-30'
                                                        union
                                                        select null,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,null,'',Output_Date,d.User_name,Mark,PCB_DC,CMC_DC,a.sno  
                                                        from Automatic_Storage_Output a
                                                        left join Automatic_Storage_Package b on a.Package = b.code  
                                                        left join Automatic_Storage_User d on a.Output_UserNo = d.User_No
                                                        where a.Unit_No=@unitNo and Item_No_Slave=@master and cast(a.Output_Date as date) >'2023-06-30' ";
                                //stritem += "where a.Unit_No=@unitNo and  Item_No_Slave = @master order by position asc";
                            }

                            strShift = txt_itemP5.Text.Trim().ToUpper();
                        }


                        SqlParameter[] parm = new SqlParameter[]
                        {
                        new SqlParameter("unitNo",Login.Unit_No),
                        new SqlParameter("master",strShift)
                        };
                        dt = db.ExecuteDataTable(stritem, CommandType.Text, parm);
                        if (dt.Rows.Count > 0)
                        {
                            dataGridView1.DataSource = dt;
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                if (dt.Rows[i]["Amount"].ToString() == "0")
                                {
                                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.MediumVioletRed;
                                }
                            }
                            sumC();
                        }
                        else
                        {
                            MessageBox.Show("查無資料，請確認輸入後重新查詢");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("請確認料號是否正確或沒有輸入");
                }
                //輸入搜尋條件後反白
                txt_item.SelectAll();
                txt_itemP5.SelectAll();
            }
        }

        private void txt_position_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                //變數宣告
                bool h_Flag = false;    //歷史查詢
                string position = txt_position.Text;

                //分析儲位
                //if (txt_position.Text.IndexOf("-") > 0)
                //{
                //    string[] postion_split = txt_position.Text.Split('-');
                //    position = postion_split[0];
                //}
                //歷史查詢
                if (txt_siteP5.Text != "")
                {
                    h_Flag = true;
                }
                if (txt_position.Text != "" || h_Flag)
                {
                    try
                    {
                        string stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  " +
                        "from Automatic_Storage_Detail a " +
                        "left join Automatic_Storage_Package b on a.Package = b.code " +
                        "left join Automatic_Storage_User c on a.Input_UserNo = c.User_No " +
                        "left join Automatic_Storage_User d on a.Output_UserNo = d.User_No " +
                        "where a.Unit_No=@unitNo and  Position like @position +'%' and amount >0 " +
                        "order by Item_No_Master";

                        if (h_Flag)
                        {
                            stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  " +
                        "from Automatic_Storage_Detail a " +
                        "left join Automatic_Storage_Package b on a.Package = b.code " +
                        "left join Automatic_Storage_User c on a.Input_UserNo = c.User_No " +
                        "left join Automatic_Storage_User d on a.Output_UserNo = d.User_No " +
                        "where a.Unit_No=@unitNo and  Position like @position +'%' and Up_InDate > '2023-06-30' " +
                        "order by Item_No_Master";
                            position = txt_siteP5.Text;
                        }

                        SqlParameter[] parm = new SqlParameter[]
                        {
                        new SqlParameter("unitNo",Login.Unit_No),
                        new SqlParameter("position",position.Trim())
                        };
                        dt = db.ExecuteDataTable(stritem, CommandType.Text, parm);
                        if (dt.Rows.Count > 0)
                        {
                            dataGridView1.DataSource = dt;
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                if (dt.Rows[i]["Amount"].ToString() == "0")
                                {
                                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.MediumVioletRed;
                                }
                            }
                            sumC();
                        }
                        else
                        {
                            MessageBox.Show("查無資料，請確認輸入後重新查詢");
                        }
                        sumC();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("請確認儲位是否正確或沒有輸入");
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            clsAll();
            dataBind();
        }
        private void clsAll()
        {
            txt_item.Text = "";
            txt_position.Text = "";
        }
        public void TimerStart()
        {
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 3);
            dispatcherTimer.Start();
        }
        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            panel1.Visible = false;
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
                cmd.CommandText = "SELECT * FROM [Sheet1$] where (入庫日期<>'' or 入庫日期 is not null) ";
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
        BackgroundWorker backgroundWorker = new BackgroundWorker();
        public void WriteExcelData()
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            try
            {
                string ItemE = "", ItemC = "", SIDE = "", Package = "", Mark = "", Spec = "", ActualDate = "", PCB = "", CMC = "";
                string xNum = "", AMOUNT_U = "";
                string AMOUNT = "";
                string USER_NO = "";
                string sqlinput = string.Empty;
                string sqldetail = string.Empty;
                string sqlchkFirst = string.Empty;
                string strsql_spec = string.Empty;
                string search_spec = string.Empty;

                int okQty = 0;
                int[] count;
                string inputDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                inputexcelcount = dsData.Tables[0].Rows.Count;
                progressBar1.Minimum = 0;
                progressBar1.Maximum = dsData.Tables[0].Rows.Count;
                progressBar1.Step = 1;
                for (int i = 0; i < dsData.Tables[0].Rows.Count; i++)
                {
                    string strErrMsg = string.Empty;
                    ActualDate = Convert.ToDateTime(dsData.Tables[0].Rows[i][0].ToString().Trim()).ToString("yyyy-MM-dd");
                    ItemE = (dsData.Tables[0].Rows[i][1].ToString().Trim().ToUpper());
                    ItemC = (dsData.Tables[0].Rows[i][2].ToString().Trim().ToUpper());
                    //乘數,最終不存DB
                    //xNum = (dsData.Tables[0].Rows[i][2].ToString().Trim() == "") ? "0" : dsData.Tables[0].Rows[i][2].ToString().Trim();                    
                    //被乘數,單位數量
                    //AMOUNT_U = (dsData.Tables[0].Rows[i][3].ToString().Trim() == "") ? "0" : dsData.Tables[0].Rows[i][3].ToString().Trim();
                    //總數,計算結果
                    //AMOUNT = "1"; ← 盤點前總數量為1的累加
                    AMOUNT = dsData.Tables[0].Rows[i][3].ToString().Trim();
                    Package = dsData.Tables[0].Rows[i][4].ToString().Trim().ToUpper();
                    SIDE = dsData.Tables[0].Rows[i][5].ToString().Trim().ToUpper();
                    PCB = dsData.Tables[0].Rows[i][6].ToString().Trim().ToUpper();
                    Mark = dsData.Tables[0].Rows[i][7].ToString().Trim().ToUpper();
                    CMC = dsData.Tables[0].Rows[i][8].ToString().Trim().ToUpper();
                    count = new int[dsData.Tables[0].Rows.Count];

                    //未設定的儲位不新增
                    string strsql_p = @"select * from Automatic_Storage_Position where Position=@position and Unit_no=@unitno";
                    SqlParameter[] parm_p = new SqlParameter[]
                    {
                        new SqlParameter("position",SIDE),
                        new SqlParameter("unitno",Login.Unit_No)
                    };
                    DataSet dsP = db.ExecuteDataSet(strsql_p, CommandType.Text, parm_p);
                    strErrMsg = (dsP.Tables[0].Rows.Count == 0) ? " ,x儲位 " : "";
                    //尚未設定規格不新增
                    if (!string.IsNullOrEmpty(ItemE))
                    {
                        search_spec = ItemE;
                        strsql_spec = @"select Item_E,item_C,spec from Automatic_Storage_Spec where Unit_no=@unitno and (Item_E =@search_spec )";
                    }
                    else if (!string.IsNullOrEmpty(ItemC))
                    {
                        search_spec = ItemC;
                        strsql_spec = @"select Item_E,item_C,spec from Automatic_Storage_Spec where Unit_no=@unitno and (item_C = @search_spec )";
                    }

                    SqlParameter[] parm_spec = new SqlParameter[]
                    {
                        new SqlParameter("unitno",Login.Unit_No),
                        new SqlParameter("search_spec",search_spec)
                        //new SqlParameter("itemC",ItemC)
                    };
                    DataSet ds_Spec = db.ExecuteDataSet(strsql_spec, CommandType.Text, parm_spec);
                    if (ds_Spec.Tables[0].Rows.Count > 0)
                    {
                        ItemE = ds_Spec.Tables[0].Rows[0]["Item_E"].ToString().Trim();
                        ItemC = ds_Spec.Tables[0].Rows[0]["item_C"].ToString().Trim();
                        Spec = ds_Spec.Tables[0].Rows[0]["Spec"].ToString().Trim();
                    }
                    else
                    {
                        ItemE = (dsData.Tables[0].Rows[i][1].ToString().Trim().ToUpper());
                        ItemC = (dsData.Tables[0].Rows[i][2].ToString().Trim().ToUpper());
                        Spec = "";
                        strErrMsg += " ,x規格 ";
                    }

                    //查詢Package
                    string strsql_pack = @"select Code from Automatic_Storage_Package where substring(Package_View,1,1) =@pack ";
                    SqlParameter[] parm_pack = new SqlParameter[]
                    {
                        new SqlParameter("pack",Package)
                    };
                    DataSet dsPack = db.ExecuteDataSet(strsql_pack, CommandType.Text, parm_pack);
                    if (dsPack.Tables[0].Rows.Count > 0)
                    {
                        Package = dsPack.Tables[0].Rows[0]["Code"].ToString().Trim();
                    }
                    else
                    {
                        strErrMsg += " ,x包裝 ";
                    }
                    //入庫日期錯誤(未來日期)                    
                    DateTime ACdt = Convert.ToDateTime(ActualDate);
                    TimeSpan timeSpan = ACdt.Subtract(DateTime.Today);
                    if (!string.IsNullOrEmpty(ActualDate))
                    {
                        if (timeSpan.TotalDays >= 1)
                        {
                            strErrMsg += " ,x日期 ";
                        }
                    }

                    //20220323判斷儲位,規格                    
                    if (dsP.Tables[0].Rows.Count == 0 || string.IsNullOrEmpty(Package) || string.IsNullOrEmpty(Spec))
                    {
                        //_ = (count[i] == 0) ? list_result.Items.Add("第" + (i + 1) + "筆失敗 "+ strErrMsg).ToString() : "";
                        _ = (count[i] == 0) ? txt_result.Text += ("第" + (i + 2) + "筆失敗 " + strErrMsg).ToString() + Environment.NewLine : "";
                    }
                    else
                    {
                        try
                        {
                            sqlinput = @"begin try
                                                    begin tran
                                                        INSERT INTO Automatic_Storage_Input
                                                        (Position,Amount,Item_No_Master,Item_No_Slave,Spec,
                                                        Unit_No,Package,Input_Date,Actual_InDate,Input_UserNo,Mark,PCB_DC,CMC_DC)   
                                                        values (@Position,@Amount,@Master,@Slave,@Spec,
                                                        @Unit_No,@package,@InDate,@ActualDate,@UserNo,@mark,@PCB,@CMC)
                                                    commit tran
                                                 end try
                                                 begin catch
                                                    rollback tran
                                                 end catch";
                            SqlParameter[] parm = new SqlParameter[]
                            {
                                    new SqlParameter("Position",SIDE.Trim().ToUpper()),
                                    new SqlParameter("Amount",AMOUNT),
                                    new SqlParameter("Master",ItemE),
                                    new SqlParameter("Slave",ItemC),
                                    new SqlParameter("Spec",Spec),
                                    new SqlParameter("Package",Package),
                                    new SqlParameter("Unit_No",Login.Unit_No),
                                    new SqlParameter("InDate",inputDate),
                                    new SqlParameter("ActualDate",ActualDate),
                                    new SqlParameter("UserNo",Login.User_No),
                                    new SqlParameter("mark",Mark),
                                    new SqlParameter("PCB",PCB),
                                    new SqlParameter("CMC",CMC)
                            };
                            db.ExecueNonQuery(sqlinput, CommandType.Text, "Form_btn_BatIn", parm);

                            //key : 入庫日期+料號+位置+包裝種類, if (key=新資料)  insert , else update
                            sqlchkFirst = @"select * from Automatic_Storage_Detail 
                                                           where Item_No_Master=@Master and  Item_No_Slave=@Slave and Position=@Position and 
                                                            Package=@Package and Actual_InDate=@actualDate and Mark = @mark and PCB_DC = @PCB and CMC_DC =@CMC and amount >0 ";
                            SqlParameter[] parameters = new SqlParameter[]
                            {
                                new SqlParameter("Master",ItemE),
                                new SqlParameter("Slave",ItemC),
                                new SqlParameter("Position",SIDE.Trim()),
                                new SqlParameter("Package",Package.Trim()),
                                new SqlParameter("actualDate",ActualDate),//目前為手動輸入,將來改自動帶入系統時間
                                new SqlParameter("mark",Mark),
                                new SqlParameter("PCB",PCB),
                                new SqlParameter("CMC",CMC)
                            };
                            DataSet dataSet = db.ExecuteDataSet(sqlchkFirst, CommandType.Text, parameters);

                            if (dataSet.Tables[0].Rows.Count > 0)
                            {
                                string detailSno = dataSet.Tables[0].Rows[0]["Sno"].ToString();
                                //資料已存在update Amount+1 ,Amount_Unit+1
                                sqldetail = @"
                                                begin try
                                                    begin tran
                                                        update Automatic_Storage_Detail set Amount=Amount+@Amount, Mark=@Mark, PCB_DC=@PCB ,CMC_DC=@CMC 
                                                        where Sno = @sno
                                                    commit tran
                                                 end try
                                                 begin catch
                                                    rollback tran
                                                 end catch";
                                /*where (Item_No_Master=@Master or Item_No_Slave = @Slave) and Position=@Position 
                                                        and Package=@Package and Actual_InDate=@actualDate*/
                                SqlParameter[] parm2 = new SqlParameter[]
                                {
                                        new SqlParameter("Amount",AMOUNT),
                                        //new SqlParameter("Amount_Unit",AMOUNT_U),
                                        new SqlParameter("Mark",Mark),
                                        new SqlParameter("PCB",PCB),
                                        new SqlParameter("CMC",CMC),
                                        new SqlParameter("Sno",detailSno)
                                        //new SqlParameter("Master",ItemE),
                                        //new SqlParameter("Slave",ItemC),
                                        //new SqlParameter("Position",SIDE.Trim()),
                                        //new SqlParameter("Package",Package.Trim()),
                                        //new SqlParameter("actualDate",ActualDate)//目前為手動輸入,將來改自動帶入系統時間
                                };
                                count[i] = db.ExecueNonQuery(sqldetail, CommandType.Text, "Form_btn_BatIn", parm2);
                            }
                            else
                            {
                                //新資料
                                //寫入detail table
                                sqldetail = @"
                                                begin try
                                                    begin tran
                                                        INSERT INTO Automatic_Storage_Detail
                                                          (Sno,Item_No_Master,Item_No_Slave,Spec,Unit_No,Position,Up_InDate,Actual_InDate,Input_UserNo,
                                                            Amount,Package,Mark,PCB_DC,CMC_DC)	 
	                                                      SELECT top(1)Sno,Item_No_Master,Item_No_Slave,Spec,Unit_No,Position,GETDATE(),Actual_InDate,
                                                            Input_UserNo,Amount,Package,Mark,PCB_DC,CMC_DC
                                                          FROM Automatic_Storage_Input
	                                                      where Item_No_Master= @Master and Item_No_Slave = @Slave and Unit_No=@Unit_No 
                                                          and Position=@position and Package=@Package and Actual_InDate=@actualDate and Spec=@Spec order by Input_Date desc
                                                    commit tran
                                                 end try
                                                 begin catch
                                                    rollback tran
                                                 end catch";
                                SqlParameter[] parm2 = new SqlParameter[]
                                {
                                        new SqlParameter("Master",ItemE),
                                        new SqlParameter("Slave",ItemC),
                                        new SqlParameter("Unit_No",Login.Unit_No),
                                        new SqlParameter("position",SIDE.Trim()),
                                        new SqlParameter("Package",Package.Trim()),
                                        new SqlParameter("actualDate",ActualDate),
                                        new SqlParameter("Spec",Spec),
                                };
                                count[i] = db.ExecueNonQuery(sqldetail, CommandType.Text, "Form_btn_BatIn", parm2);
                            }
                            okQty = okQty + 1;
                            //_ = (count[i] == 0) ? list_result.Items.Add("第" + (i + 1) + "筆失敗 " + strErrMsg).ToString() : "";
                            _ = (count[i] == 0) ? txt_result.Text += ("第" + (i + 2) + "筆失敗 " + strErrMsg).ToString() + Environment.NewLine : "";
                        }
                        catch (Exception ex)
                        {
                            list_result.Items.Add(ex.Message);
                        }
                    }
                    progressBar1.PerformStep();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("請確認資料");
            }

        }
        private void button4_Click(object sender, EventArgs e)
        {
            BatOut bO = new BatOut();
            bO.Owner = this;
            bO.ShowDialog();
        }
        private void Maintain_Click(object sender, EventArgs e)
        {
            Maintain m = new Maintain();
            m.Owner = this;
            m.ShowDialog();
            dataBind();
        }

        //料號+儲位搜尋
        private void btn_combi_Click(object sender, EventArgs e)
        {
            string strShift = string.Empty, position = string.Empty;
            bool h_flag = false;

            if (txt_itemP5.Text != "" && txt_siteP5.Text != "" && txt_itemP5.Text.Length > 3)
            {
                h_flag = true;
            }
            if (txt_item.Text != "" && txt_position.Text != "" && txt_item.Text.Length > 3 || h_flag)
            {
                try
                {
                    strShift = txt_item.Text.Trim().ToUpper();
                    position = txt_position.Text;
                    string stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC  " +
                    "from Automatic_Storage_Detail a " +
                    "left join Automatic_Storage_Package b on a.Package = b.code " +
                    "left join Automatic_Storage_User c on a.Input_UserNo = c.User_No " +
                    "left join Automatic_Storage_User d on a.Output_UserNo = d.User_No ";
                    if (rad_E.Checked)
                    {
                        stritem += "where a.Unit_No=@unitNo and  Position=@position and  Item_No_Master = @master and amount >0 " +
                    "order by Item_No_Master";
                    }
                    else if (rad_C.Checked)
                    {
                        stritem += "where a.Unit_No=@unitNo and  Position=@position and  Item_No_Slave = @master and amount >0 " +
                    "order by Item_No_Slave";
                    }

                    if (btn_itemP == 1)
                    {
                        stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC  " +
                    "from Automatic_Storage_Detail a " +
                    "left join Automatic_Storage_Package b on a.Package = b.code " +
                    "left join Automatic_Storage_User c on a.Input_UserNo = c.User_No " +
                    "left join Automatic_Storage_User d on a.Output_UserNo = d.User_No " +
                    "where a.Unit_No=@unitNo and  Position=@position and  Item_No_Master = @master " +
                    "order by Item_No_Master";
                        position = txt_siteP5.Text;
                        strShift = txt_itemP5.Text;
                    }
                    SqlParameter[] parm = new SqlParameter[]
                     {
                        new SqlParameter("unitNo", Login.Unit_No) ,
                        new SqlParameter("position", position.Trim()),
                        new SqlParameter("master", strShift)
                    };

                    dt = db.ExecuteDataTable(stritem, CommandType.Text, parm);
                    if (dt.Rows.Count > 0)
                    {
                        dataGridView1.DataSource = dt;
                        sumC();
                    }
                    else
                    {
                        MessageBox.Show("查無資料，請確認輸入後重新查詢");
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("料號及儲位資料不正確，請確認輸入後重新查詢");
            }
            btn_itemP = 0;
        }

        //整批刪除
        private void InvokeDeleteThread()
        {
            DialogResult res = MessageBox.Show(
                           "你確認刪除嗎??" + Environment.NewLine +
                           "( 料號 : " + txt_item.Text.Trim() + Environment.NewLine +
                           "庫位 : " + txt_position.Text.Trim() + " ) ",
                                           "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res == DialogResult.Yes)
            {
                progressBar2.Visible = Enabled;
                string item = string.Empty;
                string posi = string.Empty;
                string spec = string.Empty;
                string package = string.Empty;

                string time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                try
                {
                    if (string.IsNullOrEmpty(txt_position.Text))
                    {
                        MessageBox.Show("尚未指定庫位");
                        return;
                    }
                    string stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,spec,Position,Amount,package,Up_InDate,Input_UserNo,Up_OutDate,Output_UserNo " +
                    "from Automatic_Storage_Detail where Unit_No=@unitNo and  Position like @Position+'%' and amount >0 " +
                    "order by Item_No_Master";
                    SqlParameter[] parm = new SqlParameter[]
                    {
                        new SqlParameter("unitNo",Login.Unit_No),
                        new SqlParameter("Position",txt_position.Text.Trim())
                    };
                    DataSet ds_p = db.ExecuteDataSet(stritem, CommandType.Text, parm);
                    progressBar2.Minimum = 0;
                    progressBar2.Maximum = ds_p.Tables[0].Rows.Count;
                    progressBar2.Step = 1;
                    foreach (DataRow row in ds_p.Tables[0].Rows)
                    {
                        item = row["Item_No_Master"].ToString();
                        posi = row["Position"].ToString();
                        spec = row["Spec"].ToString();
                        package = row["Package"].ToString();
                        string sqlOut = @"insert into Automatic_Storage_Output (Sno,Item_No_Master,Item_No_Slave,spec,Position,Amount,package,Unit_No,Output_UserNo,Output_Date) 
                                                   select sno,Item_No_Master,Item_No_Slave,spec,Position,Amount,package,@unitno,@outuser,@outdate from Automatic_Storage_Detail 
                                                   where spec=@spec and Position=@position and package=@package and Item_No_Master=@master_d and amount > 0";
                        SqlParameter[] parm1 = new SqlParameter[]
                        {
                        new SqlParameter("unitno",Login.Unit_No),
                        new SqlParameter("outuser",Login.User_No),
                        new SqlParameter("outdate",time),
                        new SqlParameter("master",item.Trim()),
                        new SqlParameter("spec",spec.Trim()),
                        new SqlParameter("position",posi.Trim()),
                        new SqlParameter("package",package.Trim()),
                        new SqlParameter("master_d",item.Trim())
                        };
                        db.ExecueNonQuery(sqlOut, CommandType.Text, "btndelPosition", parm1);

                        //Update Storage_Detail
                        string sqlDetail = @"update Automatic_Storage_Detail 
                                                           set Amount = 0 ,Amount_Unit = 0,Up_OutDate = @outdate , Output_UserNo = @outuser 
                                                           where Item_No_Master= @master and Position=@position";
                        SqlParameter[] parm2 = new SqlParameter[]
                        {
                        new SqlParameter("outdate",time),
                        new SqlParameter("outuser",Login.User_No),
                        new SqlParameter("master",item.Trim()),
                        new SqlParameter("position",posi.Trim()),
                        };
                        db.ExecueNonQuery(sqlDetail, CommandType.Text, "btndelPosition", parm2);
                        progressBar2.PerformStep();
                    }
                    dataBind();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
        private void btndelPosition_Click(object sender, EventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(InvokeDeleteThread));
            }
            else
            {
                InvokeDeleteThread();
            }
            progressBar2.Visible = false;
        }
        private void Initi()
        {
            //預設全關
            btn_Input.Enabled = false; ;
            btn_BatIn.Enabled = false;
            btn_Out.Enabled = false;
            btn_BatOut.Enabled = false;
            btn_Maintain.Visible = false;
            btn_delPosition.Visible = false;
            //
            string sql_initi = @"select * from Automatic_Storage_UserRole where USER_ID=@userid";
            SqlParameter[] parm_initi = new SqlParameter[]
            {
                new SqlParameter("userid",Login.User_No)
            };
            DataSet ds = db.ExecuteDataSet(sql_initi, CommandType.Text, parm_initi);

            foreach (DataRow row in ds.Tables[0].Rows)
            {
                switch (row["role_id"].ToString())
                {
                    case "0"://Administrator
                        btn_delPosition.Visible = true;
                        break;
                    case "1"://Output Workbtn_BatOut.Enabled = true;                        
                        btn_Out.Enabled = true;
                        break;
                    case "2"://Input Work
                        btn_Input.Enabled = true;
                        btn_BatIn.Enabled = true;
                        break;
                    case "3"://Data Maintain
                        btn_Maintain.Visible = true;
                        break;
                    case "4"://BatOut(Position) Work 
                        btn_BatOut.Enabled = true;
                        break;
                    default:
                        break;
                }
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

        private void btn_return_Click(object sender, EventArgs e)
        {
            panel1.Visible = false;
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

        private void inputExcelBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Value = inputexcelcount;
            //progressBar1.Maximum = 100;
            MessageBox.Show("執行完成");
            dataBind();
            commitButton.Enabled = true;
            dsData.Clear();
            txt_path.Text = "";
            TimerStart();
        }

        private void updateBW_DoWork(object sender, DoWorkEventArgs e)
        {

        }

        private void updateBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

        }

        private void inputExcelBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
        }
        private string selectVerSQL_new(string tool)//Version Check new
        {
            string sqlCmd = "";
            bool result = false;
            try
            {
                sqlCmd = "select *  FROM Program_Table where [Program_Name] ='" + tool + "'";
                DataSet ds = db.reDs(sqlCmd);
                if (ds.Tables[0].Rows.Count != 0)
                {
                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        if (ds.Tables[0].Columns[i].ToString() == "Version")
                        {
                            version_new = ds.Tables[0].Rows[0][i].ToString();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("更新異常");
            }
            return version_new;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            if (X == 0 || Y == 0) return;
            fgX = (float)this.Width / (float)X;
            fgY = (float)this.Height / (float)Y;

            SetControls(fgX, fgY, this);
        }

        private void btn_findAll_Click(object sender, EventArgs e)
        {
            btn_fAll = 1;
            dataBind();
        }

        private void btn_itemSite_Click(object sender, EventArgs e)
        {
            btn_itemP = 1;
            btn_itemSite.Click += new EventHandler(btn_combi_Click);
        }
        string filePath = string.Empty;
        private void btn_qexport_Click(object sender, EventArgs e)
        {
            progressBar2.Visible = Enabled;
            filePath = "";
            //string strFilePath = txt_path.Text;
            string newName = "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

            //檔案路徑
            filePath = Application.StartupPath + "\\Upload\\查詢結果" + newName;
            //System.IO.FileInfo batchItemAttribute = new FileInfo(filePath)
            //{
            //    //設定檔案屬性為非唯讀
            //    Attributes = FileAttributes.Normal
            //};
            this.outputExcelBW.WorkerSupportsCancellation = true; //允許中斷
            this.outputExcelBW.RunWorkerAsync(); //呼叫背景程式     
        }

        private void outputExcelBW_DoWork(object sender, DoWorkEventArgs e)
        {
            if (outputExcelBW.CancellationPending) //如果被中斷...
                e.Cancel = true;
            expExcelSheet();
            GC.Collect();
            //this.Close();
        }

        private void outputExcelBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("匯出完成");
            progressBar2.Visible = false;
        }

        private void btn_history_Click(object sender, EventArgs e)
        {
            txt_item.Text = "";
            txt_position.Text = "";
            panel5.Visible = true;
            panel5.BringToFront();
            panel_rad.BringToFront();
            dt.Clear();
        }

        private void txt_spec_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                bool h_Flag = false;
                string strShift = string.Empty;
                if (txt_specP5.Text != "")
                {
                    h_Flag = true;
                }

                if (txt_spec.Text != "" || h_Flag)
                {
                    try
                    {
                        strShift = txt_spec.Text.Trim().ToUpper();

                        //改模糊查詢(20210616 有一規則length需>3)此處改>=
                        string stritem = "select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Up_InDate,c.User_name,Up_OutDate,d.User_name,a.Mark,PCB_DC,CMC_DC,a.sno  " +
                        "from Automatic_Storage_Detail a " +
                        "left join Automatic_Storage_Package b on a.Package = b.code " +
                        "left join Automatic_Storage_User c on a.Input_UserNo = c.User_No " +
                        "left join Automatic_Storage_User d on a.Output_UserNo = d.User_No " +
                        "where a.Unit_No=@unitNo and  Spec like '%'+@spec+'%' and amount >0 " +
                        "order by position asc";

                        if (h_Flag)
                        {
                            stritem = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Input_Date,c.User_name,'','',Mark,PCB_DC,CMC_DC,a.sno  
                                                        from Automatic_Storage_Input a 
                                                        left join Automatic_Storage_Package b on a.Package = b.code 
                                                        left join Automatic_Storage_User c on a.Input_UserNo = c.User_No 
                                                        where a.Unit_No=@unitNo and Spec like '%'+@spec+'%'
                                                        union
                                                        select null,Item_No_Master,Item_No_Slave,Spec,Position,Amount,b.Package,Output_Date,'','',d.User_name,Mark,PCB_DC,CMC_DC,a.sno  
                                                        from Automatic_Storage_Output a
                                                        left join Automatic_Storage_Package b on a.Package = b.code  
                                                        left join Automatic_Storage_User d on a.Output_UserNo = d.User_No
                                                        where a.Unit_No=@unitNo and Spec like '%'+@spec+'%' ";
                            strShift = txt_specP5.Text.Trim().ToUpper();
                        }


                        SqlParameter[] parm = new SqlParameter[]
                        {
                        new SqlParameter("unitNo",Login.Unit_No),
                        new SqlParameter("spec",strShift)
                        };
                        dt = db.ExecuteDataTable(stritem, CommandType.Text, parm);
                        if (dt.Rows.Count > 0)
                        {
                            dataGridView1.DataSource = dt;
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                if (dt.Rows[i]["Amount"].ToString() == "0")
                                {
                                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.MediumVioletRed;
                                }
                            }
                            sumC();
                        }
                        else
                        {
                            MessageBox.Show("查無資料，請確認輸入後重新查詢");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("請確認料號是否正確或沒有輸入");
                }
                //輸入搜尋條件後反白
                txt_item.SelectAll();
                txt_itemP5.SelectAll();
            }
        }

        private void txt_wonoOut_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                string sql_wonoOut = @" select Actual_InDate,a.Item_No_Master,a.Item_No_Slave,a.Spec,a.Position,
                                                                a.Amount,b.Package,d.Input_Date,c.User_name,Output_Date,
                                                                c.User_name,a.Mark,d.PCB_DC,CMC_DC  from Automatic_Storage_Output a
                                                                left join Automatic_Storage_Package b on a.Package = b.code 
                                                                left join Automatic_Storage_User c on a.Output_UserNo = c.User_No 
                                                                left join Automatic_Storage_Input d on a.sno=d.sno
                                                                where Wo_No =@wono order by position asc ";
                SqlParameter[] parm = new SqlParameter[]
                {
                    new SqlParameter("wono",txt_wonoOut.Text.Trim())
                };
                dt = db.ExecuteDataTable(sql_wonoOut, CommandType.Text, parm);
                dataGridView1.DataSource = dt;
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // 判斷是否為目標欄位
            if (dataGridView1.Columns[e.ColumnIndex].Name == "PCB_DC")
            {
                // 判斷該單元格的值是否符合條件
                if (!string.IsNullOrEmpty(e.Value.ToString()) && Week_Then25(e.Value.ToString()))
                {
                    // 設置該單元格的背景顏色為紅色
                    e.CellStyle.BackColor = Color.Red;
                }
            }
            if (dataGridView1.Columns[e.ColumnIndex].Name == "CMC_DC")
            {
                if (!string.IsNullOrEmpty(e.Value.ToString()) && cmc.IsOverTwoYears(e.Value.ToString()))
                {
                    // 設置該單元格的背景顏色為黃色
                    e.CellStyle.BackColor = Color.Yellow;
                }
            }
        }

        private void txt_DTs_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
            {
                if (!string.IsNullOrEmpty(txt_DTs.Text) && !string.IsNullOrEmpty(txt_DTe.Text))
                {
                    if (txt_DTs.Text.Length == 12 && txt_DTe.Text.Length == 12)
                    {
                        #region 時間查詢
                        DateTime sdate = DateTime.ParseExact(txt_DTs.Text, "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                        DateTime edate = DateTime.ParseExact(txt_DTe.Text, "yyyyMMddHHmm", CultureInfo.InvariantCulture);
                        string s_date = sdate.ToString("yyyy-MM-dd HH:mm:ss.fff");
                        string e_date = edate.ToString("yyyy-MM-dd HH:mm:59.999");
                        string strsql = "";
                        if (rad_in.Checked)
                        {
                            //strsql += "where amount > 0 and Up_InDate between @startDate and @endDate ";
                            strsql = @"select Actual_InDate,Item_No_Master,Item_No_Slave,Spec,Position,
                                                 Amount,b.Package,Input_Date,c.User_name,'','',Mark,PCB_DC,CMC_DC,a.sno  
                                                 from Automatic_Storage_Input a 
                                                 left join Automatic_Storage_Package b on a.Package = b.code 
                                                 left join Automatic_Storage_User c on a.Input_UserNo = c.User_No 
                                                 where Input_Date between @startDate and @endDate";
                        }
                        else if (rad_out.Checked)
                        {
                            //strsql += "where amount > 0 and Up_OutDate between @startDate and @endDate ";
                            strsql = @"select null,Item_No_Master,Item_No_Slave,Spec,Position,
                                                Amount,b.Package,'','',Output_Date,d.User_name,Mark,PCB_DC,CMC_DC,a.sno   
                                                from Automatic_Storage_Output a
                                                left join Automatic_Storage_Package b on a.Package = b.code  
                                                left join Automatic_Storage_User d on a.Output_UserNo = d.User_No
                                                where Output_Date between @startDate and @endDate";
                        }
                        SqlParameter[] parm = new SqlParameter[]
                        {
                            new SqlParameter("startDate",s_date),
                            new SqlParameter("endDate",e_date)
                        };
                        dt = db.ExecuteDataTable(strsql, CommandType.Text, parm);
                        dataGridView1.DataSource = dt;
                        #endregion
                    }
                }
                else
                {
                    MessageBox.Show("請確認時間格式是否正確或沒有輸入");
                }
                sumC();
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            /*
             0=actual_indate
             1=item_no_master
            2=item_no_slave
            3=spec
            4=position
            5=amount            
            6=package
            7=up_indate
            8=user_name
            9=up_outdate
            10=user_name1
            11=mark            
            12=pcb_dc
            13=cmc_dc
            14=sno
             */
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                string dda = dataGridView1.Columns[i].HeaderText.ToString();
            }
            string sno = dataGridView1.CurrentRow.Cells["sno"].Value.ToString();
            string actualDate = Convert.ToDateTime(dataGridView1.CurrentRow.Cells[0].Value.ToString()).ToString("yyyy-MM-dd");

            ActualDate acD = new ActualDate();
            acD.Owner = this;
            acD.Sno = sno;
            acD.AcD_O = actualDate;
            acD.Show();
        }

        private void btnPreviousPage_Click(object sender, EventArgs e)
        {
            if (currentPage > 1)
            {
                currentPage--;
                //dataBind(currentPage);
            }
        }

        private void btnNextPage_Click(object sender, EventArgs e)
        {
            // 在此处检查是否有更多数据，以决定是否允许继续翻页
            //DataTable dt = GetPagedData(currentPage + 1);
            if (dt.Rows.Count > 0)
            {
                currentPage++;
                //dataBind(currentPage);
            }
        }

        private void txt_position_MouseDoubleClick(object sender, MouseEventArgs e)
        {

        }

        private void btn_reP2_Click(object sender, EventArgs e)
        {
            txt_itemP5.Text = "";
            txt_siteP5.Text = "";
            panel2.BringToFront();
            dataBind();
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
        //Excel匯出作業↓
        private void expExcelSheet()
        {
            Form.CheckForIllegalCrossThreadCalls = false;
            try
            {
                progressBar2.Minimum = 0;
                progressBar2.Maximum = dt.Rows.Count;
                progressBar2.Step = 1;

                //應用程序
                Excel.Application excel_app = new Excel.Application();
                //開啟存在的檔案
                //Excel.Workbook excel_wb = excel_app.Workbooks.Open(filePath);
                //檔案
                Excel.Workbook excel_wb = excel_app.Workbooks.Add();
                //工作表
                Excel.Worksheet excel_ws = new Excel.Worksheet();
                excel_ws = excel_wb.Worksheets[1];
                excel_ws.Name = "Sheet1";

                for (int i = 0; i <= 13; i++)
                {
                    excel_app.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                }
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    excel_app.Cells[i + 2, 1] = dt.Rows[i][0].ToString();//日庫日期
                    excel_app.Cells[i + 2, 2] = dt.Rows[i][1].ToString();//料號昶
                    excel_app.Cells[i + 2, 3] = dt.Rows[i][2].ToString();//料號客
                    excel_app.Cells[i + 2, 4] = dt.Rows[i][3].ToString();//規格
                    excel_app.Cells[i + 2, 5] = dt.Rows[i][4].ToString();//儲位
                    excel_app.Cells[i + 2, 6] = dt.Rows[i][5].ToString();//數量
                    excel_app.Cells[i + 2, 7] = dt.Rows[i][6].ToString();//包裝型態
                    excel_app.Cells[i + 2, 8] = dt.Rows[i][7].ToString();//操作日期
                    excel_app.Cells[i + 2, 9] = dt.Rows[i][8].ToString();//操作人員
                    excel_app.Cells[i + 2, 10] = dt.Rows[i][9].ToString();//出庫日期
                    excel_app.Cells[i + 2, 11] = dt.Rows[i][10].ToString();//出庫人員
                    excel_app.Cells[i + 2, 12] = dt.Rows[i][11].ToString();//備註
                    excel_app.Cells[i + 2, 13] = dt.Rows[i][12].ToString();//PCB
                    excel_app.Cells[i + 2, 14] = dt.Rows[i][13].ToString();//CMC
                    progressBar2.PerformStep();
                }

                //存檔
                excel_wb.SaveAs(filePath);

                //關閉book
                excel_wb.Close(false, Type.Missing, Type.Missing);
                //關閉excel
                excel_app.Quit();
                //關閉&釋放                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel_app);
                excel_ws = null;
                excel_wb = null;
                excel_app = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void autoupdate()//自動更新
        {
            //寫入目前版本與程式名後執行更新

            Process p = new Process();
            p.StartInfo.FileName = System.Windows.Forms.Application.StartupPath + "\\AutoUpdate.exe";
            p.StartInfo.WorkingDirectory = System.Windows.Forms.Application.StartupPath; //檔案所在的目錄
            p.Start();
            this.Close();
        }
    }
}
